// Changes to a sheet's print settings, each an undoable step.
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { colName, rangeAddr, type CellRange, type PageSetup } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { parseAreas } from './paginate';
import { N, T } from '../../ui/i18n';

export const setupOf = (store: GridStore, sheet = store.sheetIndex): PageSetup => store.wb.sheets[sheet]?.pageSetup ?? {};

/** Drops fields that equal the defaults so files stay small. */
const clean = (s: PageSetup): PageSetup | undefined => {
  const o: PageSetup = { ...s };
  for (const k of Object.keys(o) as (keyof PageSetup)[]) {
    const v = o[k];
    if (v === undefined || v === false || (Array.isArray(v) && !v.length) || v === '') delete o[k];
  }
  if (o.scale === 100) delete o.scale;
  if (o.scaleWithDoc === true) delete o.scaleWithDoc;
  if (o.alignWithMargins === true) delete o.alignWithMargins;
  if (o.cellComments === 'none') delete o.cellComments;
  if (o.errors === 'displayed') delete o.errors;
  if (o.headerPictures && !Object.keys(o.headerPictures).length) delete o.headerPictures;
  return Object.keys(o).length ? o : undefined;
};

export const updateSetup = (store: GridStore, label: string, patch: Partial<PageSetup> | ((s: PageSetup) => PageSetup), sheet = store.sheetIndex): void => {
  const sh = store.wb.sheets[sheet];
  if (!sh) return;
  const run = () => {
    const cur = sh.pageSetup ?? {};
    const next = typeof patch === 'function' ? patch({ ...cur }) : { ...cur, ...patch };
    const c = clean(next);
    if (c) sh.pageSetup = c; else delete sh.pageSetup;
  };
  if (sheet === store.sheetIndex) store.change(label, 'sheet', run);
  else store.change(label, 'book', run);
};

const areaText = (r: CellRange): string => {
  if (r.r1 === 0 && r.r2 >= MAX_ROWS - 1) return `${colName(r.c1)}:${colName(r.c2)}`;
  if (r.c1 === 0 && r.c2 >= MAX_COLS - 1) return `${r.r1 + 1}:${r.r2 + 1}`;
  return rangeAddr(r);
};

export const setPrintArea = (store: GridStore, ranges: CellRange[] = store.sel.ranges): void =>
  updateSetup(store, N('Set print area'), { printArea: ranges.map(areaText).join(',') });

export const addToPrintArea = (store: GridStore, ranges: CellRange[] = store.sel.ranges): void => {
  const cur = parseAreas(setupOf(store).printArea);
  updateSetup(store, N('Add to print area'), { printArea: [...cur, ...ranges].map(areaText).join(',') });
};

export const clearPrintArea = (store: GridStore): void => updateSetup(store, N('Clear print area'), { printArea: undefined });

export const setPrintTitles = (store: GridStore, rows: string | undefined, cols: string | undefined): void =>
  updateSetup(store, N('Print titles'), { printTitleRows: rows || undefined, printTitleCols: cols || undefined });

const isWholeRow = (r: CellRange) => r.c1 === 0 && r.c2 >= MAX_COLS - 1;
const isWholeCol = (r: CellRange) => r.r1 === 0 && r.r2 >= MAX_ROWS - 1;

/** Inserts a break above the active row and/or left of the active column, as the selection says. */
export const insertBreak = (store: GridStore): void => {
  const r = store.range;
  const { row, col } = store.sel.active;
  const addRow = !isWholeCol(r) && (isWholeRow(r) ? r.r1 : row) > 0 ? (isWholeRow(r) ? r.r1 : row) : null;
  const addCol = !isWholeRow(r) && (isWholeCol(r) ? r.c1 : col) > 0 ? (isWholeCol(r) ? r.c1 : col) : null;
  if (addRow === null && addCol === null) { store.notify(T('Select a cell below or to the right of where the page should break.')); return; }
  updateSetup(store, N('Insert page break'), (s) => ({
    ...s,
    rowBreaks: addRow !== null ? [...new Set([...(s.rowBreaks ?? []), addRow])].sort((a, b) => a - b) : s.rowBreaks,
    colBreaks: addCol !== null ? [...new Set([...(s.colBreaks ?? []), addCol])].sort((a, b) => a - b) : s.colBreaks,
  }));
};

export const removeBreak = (store: GridStore): void => {
  const { row, col } = store.sel.active;
  const s = setupOf(store);
  if (!s.rowBreaks?.includes(row) && !s.colBreaks?.includes(col)) { store.notify(T('There is no manual page break at the selected cell.')); return; }
  updateSetup(store, N('Remove page break'), (x) => ({ ...x, rowBreaks: x.rowBreaks?.filter((b) => b !== row), colBreaks: x.colBreaks?.filter((b) => b !== col) }));
};

export const resetBreaks = (store: GridStore): void => updateSetup(store, N('Reset page breaks'), { rowBreaks: undefined, colBreaks: undefined });

/** Moves a break (manual or automatic) from one row/column to another; a target outside the area removes it. */
export const moveBreak = (store: GridStore, axis: 'row' | 'col', from: number, to: number, area: CellRange): void => {
  const key = axis === 'row' ? 'rowBreaks' : 'colBreaks';
  const lo = axis === 'row' ? area.r1 : area.c1, hi = axis === 'row' ? area.r2 : area.c2;
  updateSetup(store, N('Move page break'), (s) => {
    const list = (s[key] ?? []).filter((b) => b !== from);
    if (to > lo && to <= hi) list.push(to);
    // A fit-to-page setting would ignore manual breaks, so moving one switches to scaling.
    const next: PageSetup = { ...s, [key]: [...new Set(list)].sort((a, b) => a - b) };
    if (next.fitWidth !== undefined || next.fitHeight !== undefined) { delete next.fitWidth; delete next.fitHeight; next.scale = 100; }
    return next;
  });
};

export const setPageView = (store: GridStore, view: 'normal' | 'pageBreakPreview' | 'pageLayout'): void => {
  if (view === 'normal') delete store.sheet.view.pageView; else store.sheet.view.pageView = view;
  store.emit();
};
