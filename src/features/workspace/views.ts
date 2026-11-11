// Sheet views (named, plus one temporary view per sheet), custom views, split windows and full screen.
import { applyFilter, type SortKey } from '../../app/ops/data';
import { compareValues } from '../../engine/coerce';
import { isError, type Scalar } from '../../engine/types';
import { RowOrder } from '../../app/rowOrder';
import type { CellRange } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import type { AutoFilter, CustomView, NamedSheetView, Sheet } from '../../model/workbook';
import { newGuid } from '../../formats/xlsx-workspace';
import { activeGrid } from '../../ui/GridView';
import { desktopWindow } from './desktop';
import { setUi, ui } from './ui';
import { N, T } from '../../ui/i18n';

// ---------- Sheet display state ----------

interface ViewState { filter?: AutoFilter; hiddenRows: number[]; hiddenCols: number[] }

/** What a sheet view controls: the filter (with its sort) and rows or columns hidden by hand. */
export const captureView = (s: Sheet): ViewState => ({
  filter: s.autoFilter ? structuredClone(s.autoFilter) : undefined,
  hiddenRows: [...s.rows].filter(([, r]) => r.hidden && !r.filtered).map(([i]) => i).sort((a, b) => a - b),
  hiddenCols: [...s.cols].filter(([, c]) => c.hidden).map(([i]) => i).sort((a, b) => a - b),
});

/** Puts a view's filter and hidden rows and columns on the sheet (no undo step: switching views isn't an edit). */
export const applyView = (store: GridStore, sheetIndex: number, v: ViewState): void => {
  const s = store.wb.sheets[sheetIndex]!;
  for (const r of s.rows.values()) { if (r.hidden) r.hidden = false; delete r.filtered; }
  for (const c of s.cols.values()) if (c.hidden) c.hidden = false;
  for (const i of v.hiddenRows) { const r = s.rows.get(i) ?? {}; r.hidden = true; s.rows.set(i, r); }
  for (const i of v.hiddenCols) { const c = s.cols.get(i) ?? {}; c.hidden = true; s.cols.set(i, c); }
  if (v.filter) s.autoFilter = structuredClone(v.filter); else delete s.autoFilter;
  updateRowOrder(store, s);
  if (sheetIndex === store.sheetIndex) applyFilter(store);
  else {
    const was = store.wb.activeSheet;
    store.wb.activeSheet = sheetIndex;
    applyFilter(store);
    store.wb.activeSheet = was;
  }
  store.engine.recalc(true);
  store.file.dirty = true;
  store.touch();
};

/** Which view each sheet shows: null is the default view, '' is the temporary view. Keyed by sheet object. */
const active = new WeakMap<Sheet, string | null>();
/** The default view's state while another view is shown. */
const defaults = new WeakMap<Sheet, ViewState>();
const temporary = new WeakMap<Sheet, ViewState>();

export const activeViewOf = (s: Sheet): string | null => active.get(s) ?? null;

const saveCurrentInto = (s: Sheet): void => {
  const cur = active.get(s) ?? null;
  const state = captureView(s);
  if (cur === null) defaults.set(s, state);
  else if (cur === '') temporary.set(s, state);
  else {
    const v = s.sheetViews?.find((x) => x.id === cur);
    if (v) { v.filter = state.filter; v.hiddenRows = state.hiddenRows; v.hiddenCols = state.hiddenCols; }
  }
};

/** Switches the active sheet to a view: a named view's id, '' for the temporary view, or null for the default view. */
export const switchView = (store: GridStore, id: string | null): void => {
  const s = store.sheet;
  const cur = active.get(s) ?? null;
  if (cur === id) return;
  store.commitEdit();
  saveCurrentInto(s);
  let target: ViewState | undefined;
  if (id === null) target = defaults.get(s) ?? captureView(s);
  else if (id === '') target = temporary.get(s) ?? captureView(s);
  else target = s.sheetViews?.find((v) => v.id === id);
  if (!target) return;
  if (id === null) defaults.delete(s);
  active.set(s, id);
  applyView(store, store.sheetIndex, target);
};

const ensureIds = (s: Sheet): void => { for (const v of s.sheetViews ?? []) v.id ??= newGuid(); };

export const uniqueViewName = (s: Sheet, base: string): string => {
  const taken = new Set((s.sheetViews ?? []).map((v) => v.name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
};

/** New named view starting from what the sheet shows now; it becomes the active view. */
export const createView = (store: GridStore, name?: string): NamedSheetView => {
  const s = store.sheet;
  ensureIds(s);
  saveCurrentInto(s);
  const state = captureView(s);
  const v: NamedSheetView = { name: uniqueViewName(s, name?.trim() || `View${(s.sheetViews?.length ?? 0) + 1}`), id: newGuid(), ...state };
  (s.sheetViews ??= []).push(v);
  if ((active.get(s) ?? null) === null) defaults.set(s, captureView(s));
  active.set(s, v.id!);
  store.file.dirty = true;
  store.emit();
  return v;
};

/** Turns the temporary view into a named view. */
export const keepTemporary = (store: GridStore, name: string): void => {
  const s = store.sheet;
  if (active.get(s) !== '') return;
  const state = captureView(s);
  const v: NamedSheetView = { name: uniqueViewName(s, name.trim() || T('View')), id: newGuid(), ...state };
  (s.sheetViews ??= []).push(v);
  temporary.delete(s);
  active.set(s, v.id!);
  store.file.dirty = true;
  store.emit();
};

export const renameView = (store: GridStore, id: string, name: string): boolean => {
  const s = store.sheet;
  const v = s.sheetViews?.find((x) => x.id === id);
  const n = name.trim();
  if (!v || !n) return false;
  if (s.sheetViews!.some((x) => x !== v && x.name.toLowerCase() === n.toLowerCase())) { store.notify(T('Another view on this sheet has that name.')); return false; }
  v.name = n.slice(0, 255);
  store.file.dirty = true;
  store.emit();
  return true;
};

export const duplicateView = (store: GridStore, id: string): void => {
  const s = store.sheet;
  if ((active.get(s) ?? null) === id) saveCurrentInto(s);
  const v = s.sheetViews?.find((x) => x.id === id);
  if (!v) return;
  s.sheetViews!.push({ ...structuredClone(v), name: uniqueViewName(s, `${v.name} copy`), id: newGuid() });
  store.file.dirty = true;
  store.emit();
};

export const deleteView = (store: GridStore, id: string): void => {
  const s = store.sheet;
  if ((active.get(s) ?? null) === id) switchView(store, null);
  s.sheetViews = s.sheetViews?.filter((v) => v.id !== id);
  if (!s.sheetViews?.length) delete s.sheetViews;
  store.file.dirty = true;
  store.emit();
};

/** Before saving, writes the shown view's state into it and puts the default view back on the sheet data. */
export const settleViewsForSave = (store: GridStore): (() => void) => {
  const restore: (() => void)[] = [];
  store.wb.sheets.forEach((s, i) => {
    const cur = active.get(s) ?? null;
    if (cur === null) return;
    saveCurrentInto(s);
    const shown = captureView(s);
    const def = defaults.get(s);
    if (!def) return;
    applyViewSilently(s, def);
    restore.push(() => applyViewSilently(s, shown));
    void i;
  });
  return () => { for (const f of restore) f(); };
};

const applyViewSilently = (s: Sheet, v: ViewState): void => {
  for (const r of s.rows.values()) { if (r.hidden) r.hidden = false; }
  for (const c of s.cols.values()) if (c.hidden) c.hidden = false;
  for (const i of v.hiddenRows) { const r = s.rows.get(i) ?? {}; r.hidden = true; s.rows.set(i, r); }
  for (const i of v.hiddenCols) { const c = s.cols.get(i) ?? {}; c.hidden = true; s.cols.set(i, c); }
  for (const r of s.rows.values()) if (r.filtered) r.hidden = true;
  if (v.filter) s.autoFilter = structuredClone(v.filter); else delete s.autoFilter;
};

/** Forgets view bookkeeping (a new workbook was opened). */
export const resetViews = (store: GridStore): void => { for (const s of store.wb.sheets) { active.delete(s); defaults.delete(s); temporary.delete(s); store.rowOrders.delete(s); } };

/** Undo and redo replace sheet objects with copies: carry each sheet's view state over to its copy. */
export const carryViews = (store: GridStore, before: Sheet[]): void => {
  for (const old of before) {
    if (!active.has(old)) continue;
    const now = store.wb.sheets.find((x) => x.name === old.name);
    if (!now || now === old) continue;
    active.set(now, active.get(old)!);
    const d = defaults.get(old); if (d) defaults.set(now, d);
    const t = temporary.get(old); if (t) temporary.set(now, t);
    updateRowOrder(store, now);
  }
};

// ---------- Row order of a view's sort ----------

/** The data rows of a range in sorted order (same rules as sorting cells: blanks last, errors before them). */
export const sortedRows = (store: GridStore, sheetIndex: number, r1: number, r2: number, keys: { col: number; descending?: boolean; by?: SortKey['by']; color?: string; customList?: string[] }[]): number[] => {
  const val = (row: number, k: (typeof keys)[number]): Scalar => {
    if (k.by === 'cellColor') return store.style(row, k.col, sheetIndex).fill?.color === k.color ? 0 : 1;
    if (k.by === 'fontColor') return (store.style(row, k.col, sheetIndex).font?.color ?? '#000000') === k.color ? 0 : 1;
    const v = store.value(row, k.col, sheetIndex);
    if (k.customList && typeof v === 'string') { const i = k.customList.findIndex((x) => x.toLowerCase() === v.toLowerCase()); return i < 0 ? k.customList.length + 1 : i; }
    return v;
  };
  const rows: number[] = [];
  for (let r = r1; r <= r2; r++) rows.push(r);
  return rows.sort((a, b) => {
    for (const k of keys) {
      const x = val(a, k), y = val(b, k);
      if (x === null || y === null) { if (x === y) continue; return x === null ? 1 : -1; }
      const c = isError(x) || isError(y) ? (isError(x) && isError(y) ? 0 : isError(x) ? 1 : -1) : compareValues(x, y);
      if (c) return k.descending && !(isError(x) || isError(y)) ? -c : c;
    }
    return a - b;
  });
};

/** Shows the sheet's rows in its view's sort order (only while a named or temporary view is shown). */
export const updateRowOrder = (store: GridStore, s: Sheet): void => {
  const af = s.autoFilter;
  const si = store.wb.sheets.indexOf(s);
  if ((active.get(s) ?? null) === null || !af?.sort?.length || si < 0 || af.range.r2 <= af.range.r1) { store.rowOrders.delete(s); return; }
  const r1 = af.range.r1 + 1, r2 = Math.min(af.range.r2, r1 + 1_000_000);
  if (s.merges.some((m) => m.r2 >= r1 && m.r1 <= r2 && m.r2 > m.r1)) { store.rowOrders.delete(s); return; }
  const rows = sortedRows(store, si, r1, r2, af.sort.map((k) => ({ ...k, col: af.range.c1 + k.col })));
  store.rowOrders.set(s, new RowOrder(r1, Int32Array.from(rows)));
};

/** Sorting while a sheet view is shown reorders only that view's rows; the data and other views stay as they are. */
export const sortInView = (store: GridStore, r: CellRange, keys: SortKey[], header: boolean): boolean => {
  const s = store.sheet;
  if ((active.get(s) ?? null) === null) return false;
  let af = s.autoFilter;
  const body = { r1: header ? r.r1 + 1 : r.r1, r2: r.r2 };
  if (!af || af.range.c1 > r.c1 || af.range.c2 < r.c2 || af.range.r1 + 1 > body.r1 || af.range.r2 < body.r2) {
    if (af) { store.notify(T('Sort inside the filtered range of this view, or clear its filter first.')); return true; }
    if (body.r1 === 0) { store.notify(T('Sorting in a view needs a header row above the data.')); return true; }
    af = s.autoFilter = { range: { r1: body.r1 - 1, c1: r.c1, r2: body.r2, c2: r.c2 }, columns: [] };
  }
  af.sort = keys.slice(0, 64).map((k) => ({ col: k.col - af!.range.c1, ...(k.descending ? { descending: true } : {}) }));
  updateRowOrder(store, s);
  store.file.dirty = true;
  store.touch();
  store.notify(T('Sorted in this view only. The default view keeps the original order.'));
  return true;
};

// ---------- Custom views ----------

export const captureCustomView = (store: GridStore, name: string, includePrint: boolean, includeHidden: boolean, guid = newGuid()): CustomView => {
  store.sheet.view.selection = store.range;
  return {
    name: name.trim().slice(0, 255), guid, includePrint, includeHidden, activeSheet: store.sheetIndex,
    sheets: store.wb.sheets.map((s, i) => {
      const st = captureView(s);
      return {
        sheet: i, state: s.state, zoom: s.view.zoom, frozenRows: s.view.frozenRows, frozenCols: s.view.frozenCols, showGrid: s.view.showGrid, showHeaders: s.view.showHeaders,
        selection: s.view.selection ? { ...s.view.selection } : undefined,
        ...(includeHidden ? { hiddenRows: st.hiddenRows, hiddenCols: st.hiddenCols, filter: st.filter } : {}),
        ...(includePrint && s.pageSetup ? { pageSetup: structuredClone(s.pageSetup) } : {}),
      };
    }),
  };
};

export const addCustomView = (store: GridStore, name: string, includePrint: boolean, includeHidden: boolean): void => {
  const n = name.trim();
  if (!n) { store.notify(T('Give the view a name.')); return; }
  const existing = store.wb.customViews?.find((v) => v.name.toLowerCase() === n.toLowerCase());
  const view = captureCustomView(store, n, includePrint, includeHidden, existing?.guid);
  store.change(N('Add custom view'), 'book', () => {
    store.wb.customViews = [...(store.wb.customViews ?? []).filter((v) => v !== existing), view];
  });
};

export const deleteCustomView = (store: GridStore, guid: string): void => {
  store.change(N('Delete custom view'), 'book', () => {
    store.wb.customViews = store.wb.customViews?.filter((v) => v.guid !== guid);
    if (!store.wb.customViews?.length) delete store.wb.customViews;
  });
};

/** Shows a custom view: sheet visibility, zoom, panes, selection, and optionally hidden rows, filters and print settings. */
export const showCustomView = (store: GridStore, guid: string): void => {
  const v = store.wb.customViews?.find((x) => x.guid === guid);
  if (!v) return;
  if (v.sheets.some((sv) => sv.state && sv.state !== store.wb.sheets[sv.sheet]?.state) && store.wb.protection?.structure) {
    store.notify(T('This view hides or shows sheets, which the workbook protection doesn\'t allow.'));
    return;
  }
  store.change(N('Show custom view'), 'book', () => {
    for (const sv of v.sheets) {
      const s = store.wb.sheets[sv.sheet];
      if (!s) continue;
      if (sv.state) s.state = sv.state;
      if (sv.zoom) s.view.zoom = sv.zoom;
      s.view.frozenRows = sv.frozenRows ?? 0; s.view.frozenCols = sv.frozenCols ?? 0;
      if (sv.showGrid !== undefined) s.view.showGrid = sv.showGrid;
      if (sv.showHeaders !== undefined) s.view.showHeaders = sv.showHeaders;
      if (sv.selection) s.view.selection = { ...sv.selection };
      if (v.includeHidden) applyViewSilently(s, { hiddenRows: sv.hiddenRows ?? [], hiddenCols: sv.hiddenCols ?? [], filter: sv.filter });
      if (v.includePrint && sv.pageSetup) s.pageSetup = structuredClone(sv.pageSetup);
    }
    if (!store.wb.sheets.some((s) => s.state === 'visible')) store.wb.sheets[0]!.state = 'visible';
    let a = Math.min(v.activeSheet, store.wb.sheets.length - 1);
    if (store.wb.sheets[a]?.state !== 'visible') a = store.wb.sheets.findIndex((s) => s.state === 'visible');
    store.wb.activeSheet = a;
    // Filtered rows follow the restored filters.
    if (v.includeHidden) for (const sv of v.sheets) if (sv.filter && store.wb.sheets[sv.sheet]) {
      const was = store.wb.activeSheet;
      store.wb.activeSheet = sv.sheet;
      applyFilter(store);
      store.wb.activeSheet = was;
    }
  });
  const sel = store.sheet.view.selection;
  if (sel) store.selectRange(sel);
};

// ---------- Split window ----------

/** Splits the window at the top-left corner of the active cell, or removes the split. */
export const toggleSplit = (store: GridStore): void => {
  const v = store.sheet.view;
  if (v.split) {
    const back = v.split;
    delete v.split;
    store.touch();
    activeGrid.current?.scrollToPx?.(back.scrollX, back.scrollY);
    return;
  }
  if (v.frozenRows || v.frozenCols) { v.frozenRows = 0; v.frozenCols = 0; }
  const api = activeGrid.current;
  const vp = api?.viewport?.();
  const { row, col } = store.sel.active;
  if (!vp) return;
  const colX = vp.layout.cols.start(col), rowY = vp.layout.rows.start(row);
  let x = colX - vp.scrollX, y = rowY - vp.scrollY;
  const viewW = (vp.width - vp.headerW) / vp.zoom, viewH = (vp.height - vp.headerH) / vp.zoom;
  // At the top-left corner (or outside the view), split the middle of the window instead.
  if ((x <= 0 && y <= 0) || x >= viewW || y >= viewH) { x = Math.round(viewW / 2); y = Math.round(viewH / 2); }
  x = Math.max(0, x); y = Math.max(0, y);
  v.split = { x: x > 8 ? x : 0, y: y > 8 ? y : 0, scrollX: vp.scrollX, scrollY: vp.scrollY };
  store.touch();
  api?.scrollToPx?.(vp.scrollX + (v.split.x ? x : 0), vp.scrollY + (v.split.y ? y : 0));
};

// ---------- Full screen ----------

export const isFullScreen = (): boolean => ui.fullScreen;

export const toggleFullScreen = async (store: GridStore): Promise<void> => {
  const on = !ui.fullScreen;
  if (desktopWindow?.setFullScreen) {
    await desktopWindow.setFullScreen(on);
  } else if (typeof document !== 'undefined') {
    try {
      if (on && !document.fullscreenElement) await document.documentElement.requestFullscreen?.({ navigationUI: 'hide' });
      else if (!on && document.fullscreenElement) await document.exitFullscreen?.();
    } catch { /* not allowed here: fall back to hiding the chrome only */ }
  }
  setUi(store, { fullScreen: on });
  document.documentElement.toggleAttribute('data-fullscreen', on);
};
