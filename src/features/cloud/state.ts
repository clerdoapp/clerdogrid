import { useSyncExternalStore } from 'react';
import type { GridStore } from '../../app/store';
import type { Workbook } from '../../model/workbook';
import { workbookBytes, workbookFromBytes } from '../../app/files';
import { openDialog } from '../../app/commands';
import { T } from '../../ui/i18n';
import { cloud, CloudError, fromBase64, me, signInUrl, toBase64, type CloudMeta, type Conflict, type Me } from './api';

/** The cloud is only part of the build served at grid.clerdo.app; the desktop and offline web builds never call it. */
export const cloudEnabled = (): boolean => import.meta.env.VITE_CLERDO_CLOUD === '1';

export type SessionState = { status: 'unknown' } | { status: 'signedOut' } | { status: 'signedIn'; me: Me } | { status: 'error' };
export type SaveStatus = 'saved' | 'saving' | 'dirty' | 'error' | 'conflict';

interface Binding { wb: Workbook; id: string; revision: number; name: string; status: SaveStatus; storage: CloudMeta['storage']; conflict?: Conflict }

const AUTOSAVE_MS = 1500;
const RETRY_MS = 10_000;

let session: SessionState = { status: 'unknown' };
let binding: Binding | null = null;
let version = 0;
const listeners = new Set<() => void>();
const watched = new WeakSet<GridStore>();
let timer: ReturnType<typeof setTimeout> | null = null;
let saving = false;

const changed = (): void => { version++; for (const l of listeners) l(); };
const subscribe = (l: () => void): (() => void) => { listeners.add(l); return () => listeners.delete(l); };

/** Re-renders a component when the session or the cloud workbook changes. */
export const useCloud = (): { session: SessionState; binding: Readonly<Binding> | null } => {
  useSyncExternalStore(subscribe, () => version);
  return { session, binding };
};

export const currentBinding = (store: GridStore): Readonly<Binding> | null => (binding && binding.wb === store.wb ? binding : null);

export const loadSession = async (force = false): Promise<SessionState> => {
  try {
    session = { status: 'signedIn', me: await me(force) };
  } catch (e) {
    session = e instanceof CloudError && e.status === 401 ? { status: 'signedOut' } : { status: 'error' };
  }
  changed();
  return session;
};

export const signIn = (): void => { location.assign(signInUrl(location.pathname + location.search)); };

export const signOut = async (): Promise<void> => {
  try { await cloud.logout(); } catch { /* the session ends on the server either way */ }
  session = { status: 'signedOut' };
  binding = null;
  setUrl(null);
  changed();
};

const setUrl = (id: string | null): void => {
  const url = new URL(location.href);
  url.searchParams.delete('s');
  if (id) url.searchParams.set('p', id); else url.searchParams.delete('p');
  history.replaceState(history.state, '', url.pathname + url.search + url.hash);
};

const bind = (store: GridStore, meta: CloudMeta): void => {
  binding = { wb: store.wb, id: meta.id, revision: meta.revision, name: meta.name, status: 'saved', storage: meta.storage };
  store.file = { ...store.file, dirty: false, lastSaved: Date.now(), path: null, handle: undefined };
  setUrl(meta.id);
  watch(store);
  changed();
};

/** Autosave: a cloud workbook is written 1.5 s after the last change. Opening another workbook ends the link. */
const watch = (store: GridStore): void => {
  if (watched.has(store)) return;
  watched.add(store);
  store.subscribe(() => {
    if (!binding) return;
    if (binding.wb !== store.wb) {
      binding = null;
      setUrl(null);
      changed();
      return;
    }
    if (store.file.dirty && binding.status !== 'conflict') {
      if (binding.status === 'saved') { binding.status = 'dirty'; changed(); }
      schedule(store, AUTOSAVE_MS);
    }
  });
  addEventListener('beforeunload', (e) => {
    if (binding && binding.wb === store.wb && (store.file.dirty || saving)) { e.preventDefault(); void saveNow(store); }
  });
};

const schedule = (store: GridStore, ms: number): void => {
  if (timer) clearTimeout(timer);
  timer = setTimeout(() => { timer = null; void saveNow(store); }, ms);
};

const encode = (store: GridStore): { b64: string; sheets: number } => ({ b64: toBase64(workbookBytes(store)), sheets: Math.max(1, store.wb.sheets.length) });

/** Writes the cloud workbook now. Changes made while a save is in flight are saved right after it. */
export const saveNow = async (store: GridStore): Promise<void> => {
  const b = currentBinding(store);
  if (!b || binding!.status === 'conflict') return;
  if (saving) { schedule(store, AUTOSAVE_MS); return; }
  if (store.editing) { schedule(store, AUTOSAVE_MS); return; }
  if (store.file.password) { binding!.status = 'dirty'; changed(); return; }
  let payload: { b64: string; sheets: number };
  try { payload = encode(store); } catch (e) {
    binding!.status = 'error';
    changed();
    store.notify(T('The workbook could not be saved: {p0}', { p0: e instanceof Error ? e.message : String(e) }));
    return;
  }
  const target = binding!;
  saving = true;
  store.file.dirty = false;
  target.status = 'saving';
  changed();
  try {
    const meta = await cloud.save(target.id, target.revision, payload.b64, payload.sheets, store.file.name || target.name);
    target.revision = meta.revision;
    target.name = meta.name;
    target.storage = meta.storage;
    target.status = store.file.dirty ? 'dirty' : 'saved';
    store.file.lastSaved = Date.now();
    store.emit();
    postToKelvo({ type: 'toast', text: meta.storage === 'locker' ? T('Saved in Locker') : T('Saved to the cloud') });
  } catch (e) {
    store.file.dirty = true;
    if (e instanceof CloudError && e.code === 'conflict') {
      const c = e.payload as { project?: CloudMeta; workbook?: string } | undefined;
      if (c?.project && typeof c.workbook === 'string') {
        target.status = 'conflict';
        target.conflict = { project: c.project, workbook: c.workbook };
        openDialog(store, 'cloud.conflict');
      } else target.status = 'error';
    } else if (e instanceof CloudError && e.status === 401) {
      target.status = 'error';
      session = { status: 'signedOut' };
      store.notify(T('Please sign in again to save.'));
    } else {
      target.status = 'error';
      if (e instanceof CloudError && (e.code === 'too_large' || e.code === 'validation' || e.code === 'limit')) store.notify(T('The workbook could not be saved: {p0}', { p0: e.message }));
      else schedule(store, RETRY_MS);
    }
  } finally {
    saving = false;
    changed();
  }
  if (binding === target && target.status === 'dirty') schedule(store, AUTOSAVE_MS);
};

/** Puts the workbook in the person's cloud and keeps saving it there. */
export const saveToCloud = async (store: GridStore): Promise<void> => {
  store.commitEdit();
  if (currentBinding(store)) return saveNow(store);
  let payload: { b64: string; sheets: number };
  try { payload = encode(store); } catch (e) {
    store.notify(T('The workbook could not be saved: {p0}', { p0: e instanceof Error ? e.message : String(e) }));
    return;
  }
  try {
    const meta = await cloud.create(store.file.name || 'Book1', payload.b64, payload.sheets);
    bind(store, meta);
    store.emit();
    store.notify(T('Saved to the cloud'));
  } catch (e) {
    if (e instanceof CloudError && e.status === 401) { signIn(); return; }
    store.notify(T('The workbook could not be saved: {p0}', { p0: e instanceof Error ? e.message : String(e) }));
  }
};

const loadInto = (store: GridStore, meta: CloudMeta, b64: string): void => {
  const wb = workbookFromBytes(`${meta.name}.xlsx`, fromBase64(b64));
  store.load(wb, meta.name);
  bind(store, meta);
};

export const openCloud = async (store: GridStore, id: string, ask = true): Promise<boolean> => {
  if (ask && store.file.dirty && !currentBinding(store) && !confirm(T('Open this workbook? Unsaved changes in the current one will be lost.'))) return false;
  try {
    const r = await cloud.get(id);
    loadInto(store, r.project, r.workbook);
    return true;
  } catch (e) {
    if (e instanceof CloudError && e.status === 401) { signIn(); return false; }
    store.notify(T('The cloud workbook could not be opened.'));
    return false;
  }
};

export const removeCloud = async (store: GridStore, meta: CloudMeta): Promise<boolean> => {
  if (!confirm(T('Delete “{name}” from the cloud? This can’t be undone.', { name: meta.name }))) return false;
  try {
    await cloud.remove(meta.id);
    if (binding?.id === meta.id) { binding = null; setUrl(null); store.file.dirty = true; }
    changed();
    return true;
  } catch (e) {
    store.notify(e instanceof Error ? e.message : String(e));
    return false;
  }
};

export const restoreVersion = async (store: GridStore, revision: number): Promise<void> => {
  const b = currentBinding(store);
  if (!b || !confirm(T('Restore this version? The current one stays in the history.'))) return;
  if (store.file.dirty) await saveNow(store);
  try {
    const r = await cloud.restore(b.id, revision);
    loadInto(store, r.project, r.workbook);
  } catch (e) {
    store.notify(e instanceof Error ? e.message : String(e));
  }
};

/** Settles a conflict: keep this window's workbook (saved over the other version) or load the other version. */
export const resolveConflict = async (store: GridStore, keep: 'mine' | 'theirs'): Promise<void> => {
  const b = binding;
  if (!b || b.wb !== store.wb || !b.conflict) return;
  const other = b.conflict;
  b.conflict = undefined;
  if (keep === 'theirs') {
    loadInto(store, other.project, other.workbook);
    return;
  }
  b.revision = other.project.revision;
  b.status = 'dirty';
  store.file.dirty = true;
  changed();
  await saveNow(store);
};

export const openShared = async (store: GridStore, token: string): Promise<boolean> => {
  try {
    const r = await cloud.shared(token);
    store.load(workbookFromBytes(`${r.name}.xlsx`, fromBase64(r.workbook)), r.name);
    setUrl(null);
    store.notify(T('This is a copy of a shared workbook. Save it to keep your changes.'));
    return true;
  } catch {
    store.notify(T('The shared workbook could not be opened.'));
    return false;
  }
};

/** Mod+S in a cloud workbook saves to the cloud; a signed-in person's new workbook goes to the cloud too. */
export const cloudSave = async (store: GridStore): Promise<boolean> => {
  if (currentBinding(store)) { store.commitEdit(); await saveNow(store); return true; }
  if (session.status === 'signedIn') { await saveToCloud(store); return true; }
  return false;
};

/** Startup: ?p=<id> opens a cloud workbook (signing in first), ?s=<token> opens a copy of a shared one. */
export const startCloud = async (store: GridStore): Promise<boolean> => {
  watch(store);
  const params = new URLSearchParams(location.search);
  const id = params.get('p');
  const token = params.get('s');
  const s = await loadSession();
  if (token && /^[A-Za-z0-9_-]{43}$/.test(token)) return openShared(store, token);
  if (id && /^[0-9a-f-]{36}$/.test(id)) {
    if (s.status === 'signedOut') { signIn(); return true; }
    return openCloud(store, id, false);
  }
  return false;
};

// Inside Kelvo (kelvo.clerdo.app embeds apps in a frame) the app says when it is ready and what it saved.
const KELVO_ORIGIN = 'https://kelvo.clerdo.app';
const inKelvo = (): boolean => { try { return window.parent !== window; } catch { return false; } };
export const postToKelvo = (msg: Record<string, unknown>): void => {
  if (!inKelvo()) return;
  try { window.parent.postMessage({ kelvo: 1, app: 'grid', ...msg }, KELVO_ORIGIN); } catch { /* not embedded by Kelvo */ }
};
