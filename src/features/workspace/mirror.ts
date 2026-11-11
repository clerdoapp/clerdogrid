// New window: a second window on the same workbook where both windows edit. The windows are joined
// as a live session (src/features/collab): the first window is the host, the second a guest, and every
// change travels as the session's ops over a link between the windows (BroadcastChannel in the browser,
// relayed through the main process on the desktop).
import type { GridStore } from '../../app/store';
import { desktop } from '../../app/files';
import { Session, newPeerId, type Link } from '../collab/session';
import { activeSession, savedName, startSession, ui as collabUi } from '../collab/state';
import type { Msg } from '../collab/validate';
import { desktopWindow } from './desktop';
import { mirrorMode } from './history';
import { N, T, Tr } from '../../ui/i18n';

/** A frame between two windows. `to` is the window it's for; `kind` opens, carries or closes a link. */
type Frame = { t: 'wl'; from: string; to: string; kind: 'open' | 'msg' | 'close' | 'refuse'; m?: unknown; text?: string };

interface Bridge { mirrorPost?: (m: unknown) => void; onMirror?: (f: (m: unknown) => void) => void }
const bridge = (globalThis as { clerdo?: Bridge }).clerdo;

/** This window's id on the window channel (each tab or window has its own). */
export const windowId = Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const listeners = new Set<(f: Frame) => void>();
let channel: BroadcastChannel | null = null;
let wired = false;

const wire = (): void => {
  if (wired) return;
  wired = true;
  const deliver = (m: unknown) => {
    const f = m as Frame;
    if (!f || typeof f !== 'object' || f.t !== 'wl' || f.to !== windowId || typeof f.from !== 'string') return;
    for (const l of listeners) l(f);
  };
  if (bridge?.onMirror && bridge.mirrorPost) bridge.onMirror(deliver);
  else if (typeof BroadcastChannel !== 'undefined') { channel = new BroadcastChannel('clerdogrid-windows'); channel.onmessage = (e) => deliver(e.data); }
};
const post = (f: Frame): void => { wire(); if (bridge?.mirrorPost) bridge.mirrorPost(f); else channel?.postMessage(f); };
const listen = (fn: (f: Frame) => void): (() => void) => { wire(); listeners.add(fn); return () => listeners.delete(fn); };

/** A session link to another window. */
const windowLink = (peer: string): Link & { receive: (f: Frame) => void } => {
  let closed = false;
  const link: Link & { receive: (f: Frame) => void } = {
    onMessage: null, onClose: null,
    // Sent as JSON text, exactly as over a network link (structured cloning would keep undefined fields the session's checks reject).
    send: (m: Msg) => { if (!closed) post({ t: 'wl', from: windowId, to: peer, kind: 'msg', text: JSON.stringify(m) }); },
    close: () => { if (closed) return; closed = true; post({ t: 'wl', from: windowId, to: peer, kind: 'close' }); },
    receive: (f) => {
      if (closed) return;
      if (f.kind === 'msg') {
        let m: unknown;
        try { m = JSON.parse(f.text ?? 'null'); } catch { return; }
        link.onMessage?.(m);
      }
      else if (f.kind === 'close') { closed = true; link.onClose?.(); }
    },
  };
  return link;
};

export const mirrorState = { following: null as string | null, windows: 0 };
const links = new Map<string, ReturnType<typeof windowLink>>();

/** The first window: accepts second windows as guests of its live session. */
export const startMirrorSource = (store: GridStore): (() => void) => {
  const off = listen((f) => {
    if (f.kind === 'open') {
      let s = activeSession();
      if (s && s.role !== 'host') { post({ t: 'wl', from: windowId, to: f.from, kind: 'refuse', text: N('This window is in a live session someone else hosts, so it can\'t share its workbook with a new window.') }); return; }
      s ??= startSession(store, 'host', savedName() || T('Window 1'));
      const link = windowLink(f.from);
      links.set(f.from, link);
      mirrorState.windows = links.size;
      s.addGuestLink(link);
      const inner = link.onClose;
      link.onClose = () => { links.delete(f.from); mirrorState.windows = links.size; inner?.(); };
      return;
    }
    links.get(f.from)?.receive(f);
  });
  const bye = () => { for (const l of links.values()) l.close(); };
  if (typeof window !== 'undefined') window.addEventListener('beforeunload', bye);
  return () => { off(); if (typeof window !== 'undefined') window.removeEventListener('beforeunload', bye); };
};

const HANDOFF = 'clerdogrid.mirrorOf';

/** Opens a second window on this workbook. */
export const openMirrorWindow = (store: GridStore): void => {
  const s = activeSession();
  if (s && s.role !== 'host') { store.notify(T('This window is in a live session someone else hosts, so a second window can\'t join it.')); return; }
  try { localStorage.setItem(HANDOFF, JSON.stringify({ id: windowId, at: Date.now() })); } catch { /* private mode */ }
  if (desktop) { (desktopWindow?.newWindow ?? desktop.newWindow)(); return; }
  const w = window.open(`${location.pathname}?window=${windowId}`, '_blank');
  if (!w) store.notify(T('Your browser blocked the new window. Allow pop-ups for ClerdoGrid and try again.'));
};

/** Start-up: if this window was opened as a second window, join the first one's session. */
export const startMirrorIfAsked = (store: GridStore): boolean => {
  type Req = { id: string; at: number } | null;
  let req: Req = null;
  try { req = JSON.parse(localStorage.getItem(HANDOFF) ?? 'null') as Req; } catch { req = null; }
  const fromUrl = typeof location !== 'undefined' ? new URLSearchParams(location.search).get('window') : null;
  const id = fromUrl ?? (desktop && req && Date.now() - req.at < 30_000 ? req.id : null);
  if (!id || !/^[a-z0-9]{4,32}$/.test(id)) return false;
  try { localStorage.removeItem(HANDOFF); } catch { /* ignore */ }
  if (fromUrl) history.replaceState(null, '', location.pathname);
  mirrorMode.on = true;
  mirrorState.following = id;
  // Its own peer id: a tab opened from another can share that tab's session storage.
  collabUi.session?.leave();
  const session = new Session(store, 'guest', { id: newPeerId(), name: T('{name} (window 2)', { name: savedName() || T('You') }), color: '#3e63dd' });
  session.onUpdate = () => store.emit();
  session.onNotice = (t) => {
    if (t === T('Lost the connection to the host. Ask them for a new invite code to reconnect; your changes are kept.') || t === T('The host ended the session. You can keep working on your copy.')) { detachMirror(store, N('The first window closed. This window keeps the workbook as an unsaved copy.')); return; }
    store.notify(t);
  };
  collabUi.session = session;
  const link = windowLink(id);
  const off = listen((f) => {
    if (f.from !== id) return;
    if (f.kind === 'refuse') { off(); detachMirror(store, Tr(f.text) || T('The first window couldn\'t share its workbook.')); return; }
    link.receive(f);
  });
  stopFollowing = () => { off(); };
  // The first window makes its end of the link on "open", so that goes before the session's hello.
  post({ t: 'wl', from: windowId, to: id, kind: 'open' });
  session.attachHost(link);
  if (typeof window !== 'undefined') window.addEventListener('beforeunload', () => session.leave());
  return true;
};

let stopFollowing: (() => void) | null = null;

/** Leaves the shared session: this window keeps an independent, unsaved copy. */
export const detachMirror = (store: GridStore, note = N('This window is now a separate copy.')): void => {
  const s = collabUi.session;
  if (s && s.role === 'guest' && mirrorState.following) { s.leave(); collabUi.session = null; }
  stopFollowing?.();
  stopFollowing = null;
  mirrorState.following = null;
  mirrorMode.on = false;
  store.readOnly = false;
  store.file = { ...store.file, name: store.file.name.replace(/(:2| \(window 2\))$/, '') + ' copy', path: null, handle: undefined, dirty: true };
  store.notify(note);
  store.touch();
};
