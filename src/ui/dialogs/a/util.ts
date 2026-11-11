import type { Scalar } from '../../../engine/types';
import { MAX_COLS, MAX_ROWS } from '../../../engine/types';
import { cellKey, internStyle, mergeStyle, newUserSheet as newSheet, rangeAddr, type CellRange, type CellStyle } from '../../../model/workbook';
import type { GridStore } from '../../../app/store';
import { parseRef } from '../../charts/chartData';

export interface SheetRange { sheet: number; name: string; range: CellRange }

/** Reads "A1:D9" or "Sheet2!A1:D9" (with or without $ and a leading =). */
export const resolveRange = (store: GridStore, text: string): SheetRange | null => {
  const p = parseRef(text);
  if (!p) return null;
  const sheet = p.sheet === undefined ? store.sheetIndex : store.wb.sheets.findIndex((s) => s.name.toLowerCase() === p.sheet!.toLowerCase());
  if (sheet < 0) return null;
  const r = p.range;
  if (r.r2 >= MAX_ROWS || r.c2 >= MAX_COLS) return null;
  return { sheet, name: store.wb.sheets[sheet]!.name, range: r };
};

export const rangeText = (store: GridStore, r: CellRange, sheet = store.sheetIndex): string =>
  sheet === store.sheetIndex ? rangeAddr(r) : `${store.wb.sheets[sheet]!.name}!${rangeAddr(r)}`;

/** Values of a block, clipped to the used area and to `maxRows` rows. */
export const readMatrix = (store: GridStore, sheet: number, r: CellRange, maxRows = 5000): Scalar[][] => {
  const b = store.engine.usedBounds(sheet);
  const r2 = Math.min(r.r2, Math.max(r.r1, b.rows), r.r1 + maxRows - 1);
  const c2 = Math.min(r.c2, Math.max(r.c1, b.cols), r.c1 + 255);
  const out: Scalar[][] = [];
  for (let row = r.r1; row <= r2; row++) {
    const line: Scalar[] = [];
    for (let col = r.c1; col <= c2; col++) line.push(store.value(row, col, sheet));
    out.push(line);
  }
  return out;
};

/** Clips a range to the sheet's used area, so whole-column selections stay small. */
export const clipToUsed = (store: GridStore, sheet: number, r: CellRange): CellRange => {
  const b = store.engine.usedBounds(sheet);
  return { ...r, r2: Math.min(r.r2, Math.max(r.r1, b.rows)), c2: Math.min(r.c2, Math.max(r.c1, b.cols)) };
};

export const uniqueSheetName = (store: GridStore, base: string): string => {
  const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
};

/** Adds a sheet at the end and makes it active. Call inside store.change(..., 'book', ...). */
export const addSheetAtEnd = (store: GridStore, base: string): number => {
  store.wb.sheets.push(newSheet(uniqueSheetName(store, base)));
  store.engine.sheetsChanged();
  store.wb.activeSheet = store.wb.sheets.length - 1;
  return store.wb.activeSheet;
};

/** Sets a value or formula ("=..." in `f`) and merges a style into the cell. */
export const put = (store: GridStore, sheet: number, row: number, col: number, v: Scalar, style?: CellStyle, f?: string): void => {
  if (f) store.engine.setFormula(sheet, row, col, f);
  else store.engine.setValue(sheet, row, col, v);
  if (!style) return;
  const s = store.wb.sheets[sheet]!;
  const k = cellKey(row, col);
  const id = internStyle(store.wb, mergeStyle(store.style(row, col, sheet), style));
  const c = s.cells.get(k);
  if (c) c.s = id; else if (id) s.cells.set(k, { v: null, s: id });
};

export const BOLD: CellStyle = { font: { bold: true } };

export const setColWidth = (store: GridStore, sheet: number, col: number, chars: number): void => {
  const s = store.wb.sheets[sheet]!;
  s.cols.set(col, { ...(s.cols.get(col) ?? {}), width: chars });
};

export const headerNames = (store: GridStore, sheet: number, r: CellRange, header: boolean): string[] => {
  const out: string[] = [];
  for (let c = r.c1; c <= r.c2; c++) {
    const v = header ? store.value(r.r1, c, sheet) : null;
    out.push(v === null || v === '' ? `Column ${colLetter(c)}` : String(typeof v === 'object' ? v.code : v));
  }
  return out;
};

const colLetter = (c: number): string => rangeAddr({ r1: 0, c1: c, r2: 0, c2: c }).replace(/\d+/g, '');

/** Whether any cell in the range has content. */
export const hasContent = (store: GridStore, sheet: number, r: CellRange): boolean => {
  const s = store.wb.sheets[sheet]!;
  for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
    const c = s.cells.get(cellKey(row, col));
    if (c && (c.v !== null || c.f)) return true;
  }
  return false;
};
