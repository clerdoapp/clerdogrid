// Screen state for PivotTables (not saved, not undoable) and helpers to find the table in focus.
import type { GridStore } from '../../app/store';
import type { PivotLayout, PCell } from './compute';
import { cacheOf, layoutOf, pivotAt, pivotCellAt, type PivotRef } from './apply';
import type { PivotCache } from '../../model/pivot';

const load = (k: string, d: boolean): boolean => { try { const v = localStorage.getItem(k); return v === null ? d : v === '1'; } catch { return d; } };
export const save = (k: string, v: boolean): void => { try { localStorage.setItem(k, v ? '1' : '0'); } catch { /* private mode */ } };

export const pui = {
  /** Field list pane is open. */
  pane: false,
  /** Last table the pane showed (kept while the cell moves out of it). */
  paneTarget: null as { sheet: number; name: string } | null,
  /** Clicking a table value while typing a formula makes GETPIVOTDATA. */
  gpd: load('clerdogrid.pivot.gpd', true),
  /** Filter and sort popup opened from a field button. */
  popup: null as null | { sheet: number; name: string; field: number | 'rows' | 'cols'; x: number; y: number; page?: boolean },
  /** Selected slicer or timeline on the active sheet. */
  slicer: null as string | null,
  /** Field list search text. */
  search: '',
  /** Open the field list when a table cell is selected (off once the user closes it). */
  autoPane: true,
  /** Table seen under the active cell last time, to notice moving into a table. */
  lastSeen: '',
};

export const setPui = (store: GridStore, patch: Partial<typeof pui>): void => { Object.assign(pui, patch); store.emit(); };

/** The table under the active cell. */
export const activePivot = (store: GridStore): PivotRef | null => pivotAt(store.wb, store.sheetIndex, store.sel.active.row, store.sel.active.col);

/** The table the tools act on: the one under the active cell, else the one the pane last showed. */
export const currentPivot = (store: GridStore): PivotRef | null => {
  const a = activePivot(store);
  if (a) return a;
  const t = pui.paneTarget;
  if (!t || t.sheet !== store.sheetIndex) return null;
  const index = store.wb.sheets[t.sheet]?.pivots?.findIndex((p) => p.name === t.name) ?? -1;
  return index >= 0 ? { sheet: t.sheet, index, pt: store.wb.sheets[t.sheet]!.pivots![index]! } : null;
};

export interface Focus { ref: PivotRef; cache: PivotCache; layout: PivotLayout; cell: PCell | null }

/** Table, cache, layout and the layout cell under the active cell. */
export const focusOf = (store: GridStore): Focus | null => {
  const { row, col } = store.sel.active;
  const hit = pivotCellAt(store.wb, store.sheetIndex, row, col);
  const ref = hit?.ref ?? currentPivot(store);
  if (!ref) return null;
  const cache = cacheOf(store.wb, ref.pt.cacheId);
  const layout = hit?.layout ?? layoutOf(store.wb, ref.pt);
  if (!cache || !layout) return null;
  return { ref, cache, layout, cell: hit?.cell ?? null };
};

/** The axis field the active cell belongs to (item, caption or page field). */
export const focusField = (f: Focus | null): number | null => {
  if (!f?.cell) return null;
  if (f.cell.field !== undefined && f.cell.field >= 0) return f.cell.field;
  if (typeof f.cell.filter === 'number') return f.cell.filter;
  if (f.cell.filter === 'rows') return f.ref.pt.rows.find((x) => x >= 0) ?? null;
  if (f.cell.filter === 'cols') return f.ref.pt.cols.find((x) => x >= 0) ?? null;
  return null;
};
