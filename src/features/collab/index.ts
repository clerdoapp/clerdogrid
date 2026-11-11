// Live collaboration (device to device over WebRTC, no server) and share links.
import { createElement } from 'react';
import { openDialog, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerOverlay, registerPanelSection, registerStartup, registerTopBarItem } from '../../ui/panelRegistry';
import { openShareHash, shareFile } from './share';
import { activeSession, endSession, ui } from './state';
import { InviteDialog, JoinDialog, RemoteCursors, SessionChip, ShareDialog, SharePanel, StartDialog } from './ui';
import './collab.css';
import { N } from '../../ui/i18n';

const clearHash = () => { try { history.replaceState(null, '', `${location.pathname}${location.search}`); } catch { /* sandboxed */ } };

/** Handles "#wb=" share links and "#join=" invite links. Returns true when a workbook was opened. */
const handleHash = (store: GridStore): boolean => {
  const h = location.hash;
  if (h.startsWith('#wb=')) {
    // A share link replaces the open workbook, so a session in progress is left first.
    if (activeSession()) endSession(store);
    const ok = openShareHash(store, h);
    clearHash();
    return ok;
  }
  const m = /^#join=([A-Za-z0-9._-]{8,})/.exec(h);
  if (m) {
    ui.joinCode = decodeURIComponent(m[1]!);
    clearHash();
    setTimeout(() => openDialog(store, 'collab.join'), 0);
  }
  return false;
};

export const register = (): void => {
  registerCommand({ id: 'collab.start', label: N('Start live session'), group: N('Share'), enabled: () => !activeSession(), run: (s) => openDialog(s, 'collab.start') });
  registerCommand({ id: 'collab.join', label: N('Join live session'), group: N('Share'), enabled: () => !activeSession() || activeSession()!.status === 'lost', run: (s) => openDialog(s, 'collab.join') });
  registerCommand({ id: 'collab.invite', label: N('Invite someone to the live session'), group: N('Share'), enabled: () => activeSession()?.role === 'host', run: (s) => openDialog(s, 'collab.invite') });
  registerCommand({ id: 'collab.people', label: N('People in the live session'), group: N('Share'), enabled: () => !!activeSession(), run: (s) => { ui.peopleOpen = !ui.peopleOpen; s.emit(); } });
  registerCommand({ id: 'collab.leave', label: N('Leave or end the live session'), group: N('Share'), enabled: () => !!activeSession(), run: (s) => endSession(s) });
  registerCommand({ id: 'share.link', label: N('Copy link'), group: N('Share'), run: (s) => openDialog(s, 'share.link', { readOnly: false }) });
  registerCommand({ id: 'share.linkReadOnly', label: N('Copy read-only link'), group: N('Share'), run: (s) => openDialog(s, 'share.link', { readOnly: true }) });
  registerCommand({ id: 'share.file', label: N('Share file'), group: N('Share'), run: (s) => { void shareFile(s); } });

  registerDialog('collab.start', StartDialog);
  registerDialog('collab.invite', InviteDialog);
  registerDialog('collab.join', JoinDialog);
  registerDialog('share.link', ShareDialog);

  registerPanelSection({ tab: 'file', title: N('Share'), order: 5, render: (s) => createElement(SharePanel, { store: s }) });
  registerOverlay({ id: 'collab.cursors', order: 25, render: (s, vp) => createElement(RemoteCursors, { store: s, vp }) });
  registerTopBarItem({ id: 'collab.people', order: 10, render: (s) => createElement(SessionChip, { store: s }) });

  registerStartup((store) => {
    if (typeof window === 'undefined') return false;
    window.addEventListener('hashchange', () => handleHash(store));
    window.addEventListener('beforeunload', (e) => {
      const s = activeSession();
      if (s?.role === 'host' && s.liveGuests) { e.preventDefault(); e.returnValue = ''; }
    });
    return handleHash(store);
  });
};
