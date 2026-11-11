// The live session of this window (at most one) and small UI state around it.
import type { GridStore } from '../../app/store';
import { Session, newPeerId } from './session';
import type { PeerInfo } from './validate';
import { T } from '../../ui/i18n';

export const ui = {
  session: null as Session | null,
  peopleOpen: false,
  /** Invite code from a #join= link, waiting for the Join dialog. */
  joinCode: '',
};

const NAME_KEY = 'clerdogrid.user';
const ID_KEY = 'clerdogrid.peer';

export const savedName = (): string => { try { return localStorage.getItem(NAME_KEY)?.trim() ?? ''; } catch { return ''; } };
export const saveName = (n: string): void => { try { localStorage.setItem(NAME_KEY, n.trim().slice(0, 60)); } catch { /* storage blocked */ } };

/** One id per browser tab, so a reload or reconnect is recognized as the same person. */
export const myPeerId = (): string => {
  try {
    const v = sessionStorage.getItem(ID_KEY);
    if (v && /^[A-Za-z0-9_-]{1,64}$/.test(v)) return v;
    const id = newPeerId();
    sessionStorage.setItem(ID_KEY, id);
    return id;
  } catch { return newPeerId(); }
};

export const startSession = (store: GridStore, role: 'host' | 'guest', name: string): Session => {
  ui.session?.leave();
  const me: PeerInfo = { id: myPeerId(), name: name.trim().slice(0, 60) || (role === 'host' ? T('Host') : T('Guest')), color: '#e5484d' };
  const s = new Session(store, role, me);
  s.onUpdate = () => store.emit();
  s.onNotice = (t) => store.notify(t);
  ui.session = s;
  store.emit();
  return s;
};

export const endSession = (store: GridStore): void => {
  ui.session?.leave();
  ui.session = null;
  ui.peopleOpen = false;
  store.emit();
};

export const activeSession = (): Session | null => (ui.session && ui.session.status !== 'ended' ? ui.session : null);

export const initials = (name: string): string => name.split(/\s+/).filter(Boolean).slice(0, 2).map((w) => [...w][0]!.toUpperCase()).join('') || '?';
