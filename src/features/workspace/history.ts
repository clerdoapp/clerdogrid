// Takes version snapshots while a workbook is open and handles previewing and restoring them.
import { Engine } from '../../engine/engine';
import { readXlsx } from '../../formats/xlsx';
import type { GridStore, Selection } from '../../app/store';
import type { Workbook } from '../../model/workbook';
import { desktop } from '../../app/files';
import { settings } from './settings';
import { addVersion, fileKey, getVersion, listVersions, loadVersion, moveVersions, snapshotBytes, type VersionMeta } from './versions';
import { resetViews } from './views';
import { N, T } from '../../ui/i18n';

/** Version list of the open file, refreshed after each change to the store of versions. */
export const versionList: { file: string; items: VersionMeta[]; loading: boolean; error: string | null } = { file: '', items: [], loading: false, error: null };

export const refreshVersions = async (store: GridStore): Promise<void> => {
  const file = fileKey(store);
  versionList.loading = true;
  try {
    versionList.items = await listVersions(file);
    versionList.file = file;
    versionList.error = null;
  } catch {
    versionList.items = [];
    versionList.error = T('Version history isn\'t available in this browser window.');
  }
  versionList.loading = false;
  store.emit();
};

const canSnapshot = (store: GridStore): boolean => !preview.current && !store.file.password && !mirrorMode.on && typeof indexedDB !== 'undefined';

let busy = false;
/** Stores a snapshot of the open workbook now (skipped when nothing changed since the last one). */
export const snapshotNow = async (store: GridStore, reason: VersionMeta['reason'], name?: string): Promise<VersionMeta | null> => {
  if (!canSnapshot(store) || busy) return null;
  busy = true;
  try {
    store.sheet.view.selection = store.range;
    const meta = await addVersion(fileKey(store), snapshotBytes(store.wb), reason, name);
    await refreshVersions(store);
    return meta;
  } catch {
    return null;
  } finally { busy = false; }
};

/** Read-only window mirroring another (set by mirror.ts). */
export const mirrorMode = { on: false };

/** Saves on save, and every few minutes while the workbook keeps changing. */
export const startVersionTimer = (store: GridStore): (() => void) => {
  let lastSaved = store.file.lastSaved;
  let lastKey = fileKey(store);
  let changedSince = false;
  let lastVersion = store.getVersion();
  let lastAuto = Date.now();
  const unsub = store.subscribe(() => {
    if (store.getVersion() !== lastVersion && store.file.dirty) changedSince = true;
    lastVersion = store.getVersion();
    const key = fileKey(store);
    if (key !== lastKey) {
      // Save as gave the file a new identity; versions follow it. A different workbook just switches lists.
      const prev = lastKey;
      lastKey = key;
      if (store.file.lastSaved && store.file.lastSaved !== lastSaved && prev.startsWith('name:')) void moveVersions(prev, key).then(() => refreshVersions(store));
      else void refreshVersions(store);
    }
    if (store.file.lastSaved && store.file.lastSaved !== lastSaved) {
      lastSaved = store.file.lastSaved;
      changedSince = false;
      if (settings.versionsOn) void snapshotNow(store, 'save');
    }
  });
  const t = setInterval(() => {
    if (!settings.versionsOn || !changedSince || store.editing) return;
    if (Date.now() - lastAuto < Math.max(1, settings.versionsEvery) * 60_000) return;
    lastAuto = Date.now();
    changedSince = false;
    void snapshotNow(store, 'auto');
  }, 15_000);
  return () => { unsub(); clearInterval(t); };
};

// ---------- Preview ----------

interface Saved { wb: Workbook; engine: Engine; sel: Selection; readOnly: boolean; dirty: boolean; panel: string | null }
export const preview: { current: { meta: VersionMeta; saved: Saved } | null } = { current: null };

export const openPreview = async (store: GridStore, id: string): Promise<void> => {
  store.commitEdit();
  const v = await loadVersion(id);
  if (!v) { store.notify(T('That version is no longer available.')); return; }
  if (preview.current) closePreview(store, false);
  preview.current = { meta: v.meta, saved: { wb: store.wb, engine: store.engine, sel: store.sel, readOnly: store.readOnly, dirty: store.file.dirty, panel: store.view.panel } };
  store.wb = v.wb;
  store.engine = new Engine(v.wb);
  if (v.wb.calc?.iterate) store.engine.iterate = true;
  store.engine.recalc(true);
  store.readOnly = true;
  const s = store.sheet.view.selection;
  const at = s ? { row: s.r1, col: s.c1 } : { row: 0, col: 0 };
  store.sel = { ranges: [s ?? { r1: 0, c1: 0, r2: 0, c2: 0 }], active: at, anchor: at };
  store.touch();
};

export const closePreview = (store: GridStore, touch = true): void => {
  const p = preview.current;
  if (!p) return;
  preview.current = null;
  store.wb = p.saved.wb;
  store.engine = p.saved.engine;
  store.sel = p.saved.sel;
  store.readOnly = p.saved.readOnly;
  store.file.dirty = p.saved.dirty;
  if (touch) store.touch();
};

/** Replaces the workbook with the previewed version as one undoable step, after keeping the current state as a version. */
export const restorePreview = async (store: GridStore): Promise<void> => {
  const p = preview.current;
  if (!p) return;
  const version = store.wb;
  closePreview(store, false);
  if (store.readOnly) { store.notify(T('This workbook is read-only.')); store.touch(); return; }
  await snapshotNow(store, 'restore', undefined);
  resetViews(store);
  store.change(N('Restore version'), 'book', () => {
    const wb = store.wb;
    wb.sheets = version.sheets; wb.names = version.names; wb.styles = version.styles; wb.activeSheet = Math.min(version.activeSheet, version.sheets.length - 1);
    wb.date1904 = version.date1904; wb.calc = version.calc; wb.protection = version.protection; wb.customViews = version.customViews;
    wb.scripts = version.scripts; wb.props = { ...version.props, ...wb.props, modified: wb.props?.modified };
    if (version.extraParts) wb.extraParts = version.extraParts;
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
  const s = store.sheet.view.selection;
  if (s) store.selectRange(s);
  store.notify(T('Restored the version from {p0}. Undo brings back what you had.', { p0: new Date(p.meta.at).toLocaleString() }));
};

const HANDOFF_KEY = 'clerdogrid.openVersion';

/** Opens the previewed version as a separate, unsaved workbook (a new tab or window). */
export const copyPreviewToNew = async (store: GridStore): Promise<void> => {
  const p = preview.current;
  if (!p) return;
  const name = `${store.file.name || 'Book1'} (${new Date(p.meta.at).toLocaleDateString()})`;
  try { localStorage.setItem(HANDOFF_KEY, JSON.stringify({ id: p.meta.id, name, at: Date.now() })); } catch { /* private mode */ }
  if (desktop) { desktop.newWindow(); store.notify(T('Opening the version in a new window.')); return; }
  const w = window.open(`${location.pathname}?version=${encodeURIComponent(p.meta.id)}`, '_blank');
  if (!w) {
    // Pop-ups are blocked: open it here instead, leaving the current workbook's versions in place.
    closePreview(store, false);
    if (store.file.dirty && !confirm(T('Open this version here? Unsaved changes in the current workbook will be lost.'))) { store.touch(); return; }
    await openHandoff(store, p.meta.id, name);
  }
};

const openHandoff = async (store: GridStore, id: string, name: string): Promise<boolean> => {
  const v = await getVersion(id);
  if (!v) return false;
  store.load(readXlsx(v.bytes), name);
  store.file.dirty = true;
  return true;
};

/** Start-up: a new tab or window asked to open a version as a new workbook. */
export const consumeHandoff = async (store: GridStore): Promise<boolean> => {
  type Req = { id: string; name: string; at: number } | null;
  let req: Req = null;
  try { req = JSON.parse(localStorage.getItem(HANDOFF_KEY) ?? 'null') as Req; } catch { req = null; }
  const fromUrl = typeof location !== 'undefined' ? new URLSearchParams(location.search).get('version') : null;
  if (!req || Date.now() - req.at > 60_000 || (fromUrl && fromUrl !== req.id) || (!fromUrl && !desktop)) return false;
  try { localStorage.removeItem(HANDOFF_KEY); } catch { /* ignore */ }
  if (fromUrl) history.replaceState(null, '', location.pathname);
  return openHandoff(store, req.id, req.name);
};
