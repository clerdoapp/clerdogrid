import { adjustForStructure, deleteSheetInFormula, renameSheetInFormula, type Structural } from '../../engine/refactor';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { cellKey, keyCol, keyRow, newSheet, type CellRange, type Sheet } from '../../model/workbook';
import { cloneSheet } from '../clone';
import type { GridStore } from '../store';

const shiftRange = (r: CellRange, op: Structural): CellRange | null => {
  const del = op.count < 0 ? -op.count : 0;
  const end = op.at + del - 1;
  const span = (a: number, b: number): [number, number] | null => {
    if (op.count > 0) return [a >= op.at ? a + op.count : a, b >= op.at ? b + op.count : b];
    if (a >= op.at && b <= end) return null;
    const na = a > end ? a - del : a >= op.at ? op.at : a;
    const nb = b > end ? b - del : b >= op.at ? op.at - 1 : b;
    return nb < na ? null : [na, nb];
  };
  const max = op.axis === 'row' ? MAX_ROWS - 1 : MAX_COLS - 1;
  if (op.axis === 'row') {
    if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return r;
    const s = span(r.r1, r.r2);
    return s ? { ...r, r1: Math.min(s[0], max), r2: Math.min(s[1], max) } : null;
  }
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return r;
  const s = span(r.c1, r.c2);
  return s ? { ...r, c1: Math.min(s[0], max), c2: Math.min(s[1], max) } : null;
};

const shiftIndexMap = <T>(m: Map<number, T>, op: Structural, keyed: 'row' | 'col' | 'cell'): Map<number, T> => {
  const out = new Map<number, T>();
  const del = op.count < 0 ? -op.count : 0;
  const move = (i: number): number | null => {
    if (op.count > 0) return i >= op.at ? i + op.count : i;
    if (i >= op.at && i < op.at + del) return null;
    return i >= op.at + del ? i - del : i;
  };
  for (const [k, v] of m) {
    if (keyed === 'cell') {
      let r = keyRow(k), c = keyCol(k);
      const ni = move(op.axis === 'row' ? r : c);
      if (ni === null) continue;
      if (op.axis === 'row') r = ni; else c = ni;
      if (r >= MAX_ROWS || c >= MAX_COLS) continue;
      out.set(cellKey(r, c), v);
    } else if (keyed === op.axis) {
      const ni = move(k);
      if (ni !== null && ni < (keyed === 'row' ? MAX_ROWS : MAX_COLS)) out.set(ni, v);
    } else out.set(k, v);
  }
  return out;
};

/** Insert (count > 0) or delete (count < 0) whole rows or columns on the active sheet. */
export const changeStructure = (store: GridStore, axis: 'row' | 'col', at: number, count: number): void => {
  const sheetIdx = store.sheetIndex;
  const sheet = store.sheet;
  const op: Structural = { sheet: sheet.name, axis, at, count };
  if (count > 0) {
    // Refuse to push data off the end of the sheet, as spreadsheets do.
    const last = axis === 'row' ? MAX_ROWS : MAX_COLS;
    for (const k of sheet.cells.keys()) if ((axis === 'row' ? keyRow(k) : keyCol(k)) + count >= last) { store.notify("There isn't room to insert: data would be pushed off the end of the sheet."); return; }
  }
  if (count < 0 && sheet.tables.some((t) => (axis === 'row' ? t.range.r1 >= at && t.range.r1 < at - count && t.headerRow : false))) {
    store.notify('That would delete a table header row. Convert the table to a range first.');
    return;
  }
  store.change(count > 0 ? `Insert ${axis === 'row' ? 'rows' : 'columns'}` : `Delete ${axis === 'row' ? 'rows' : 'columns'}`, 'book', () => {
    store.wb.sheets.forEach((s, si) => {
      for (const c of s.cells.values()) {
        if (c.f) c.f = adjustForStructure(c.f, s.name, op);
        if (si === sheetIdx && c.arrayRange) c.arrayRange = shiftRange(c.arrayRange, op) ?? undefined;
      }
      for (const cf of s.conditionalFormats) for (const r of cf.rules) {
        if ('formula' in r) r.formula = adjustForStructure(r.formula, s.name, op);
        if ('formulas' in r) r.formulas = r.formulas.map((f) => adjustForStructure(f, s.name, op));
      }
      for (const dv of s.validations) {
        if (dv.formula1) dv.formula1 = adjustForStructure(dv.formula1.replace(/^=/, ''), s.name, op);
        if (dv.formula2) dv.formula2 = adjustForStructure(dv.formula2.replace(/^=/, ''), s.name, op);
      }
      for (const ch of s.charts) for (const ser of ch.series) {
        ser.values = adjustForStructure(ser.values, s.name, op);
        if (ser.categories) ser.categories = adjustForStructure(ser.categories, s.name, op);
        if (ser.name) ser.name = adjustForStructure(ser.name, s.name, op);
      }
    });
    for (const n of store.wb.names) n.formula = adjustForStructure(n.formula, n.sheet !== undefined ? store.wb.sheets[n.sheet]!.name : sheet.name, { ...op });
    sheet.cells = shiftIndexMap(sheet.cells, op, 'cell');
    sheet.hyperlinks = shiftIndexMap(sheet.hyperlinks, op, 'cell');
    sheet.notes = shiftIndexMap(sheet.notes, op, 'cell');
    if (axis === 'row') sheet.rows = shiftIndexMap(sheet.rows, op, 'row'); else sheet.cols = shiftIndexMap(sheet.cols, op, 'col');
    sheet.merges = sheet.merges.map((m) => shiftRange(m, op)).filter((m): m is CellRange => !!m && (m.r1 !== m.r2 || m.c1 !== m.c2));
    for (const cf of sheet.conditionalFormats) cf.ranges = cf.ranges.map((r) => shiftRange(r, op)).filter((r): r is CellRange => !!r);
    sheet.conditionalFormats = sheet.conditionalFormats.filter((cf) => cf.ranges.length);
    for (const dv of sheet.validations) dv.ranges = dv.ranges.map((r) => shiftRange(r, op)).filter((r): r is CellRange => !!r);
    sheet.validations = sheet.validations.filter((dv) => dv.ranges.length);
    for (const t of sheet.tables) { const r = shiftRange(t.range, op); if (r) t.range = r; }
    if (sheet.autoFilter) { const r = shiftRange(sheet.autoFilter.range, op); if (r) sheet.autoFilter.range = r; else delete sheet.autoFilter; }
    for (const ch of [...sheet.charts, ...sheet.pictures]) {
      const p = axis === 'row' ? ch.anchor.row : ch.anchor.col;
      if (p >= at) { const np = count > 0 ? p + count : Math.max(at, p + count); if (axis === 'row') ch.anchor.row = np; else ch.anchor.col = np; }
    }
    store.engine.sheetsChanged();
    store.engine.recalc(true);
    // Keep the selection on the same place.
    const r = store.range;
    store.sel = { ...store.sel, ranges: [r] };
  });
};

/** Inserts or deletes cells and shifts the neighbours (Insert > Shift cells down / right). */
export const shiftCells = (store: GridStore, r: CellRange, dir: 'down' | 'right' | 'up' | 'left'): void => {
  const sheet = store.sheet;
  const h = r.r2 - r.r1 + 1, w = r.c2 - r.c1 + 1;
  store.change('Shift cells', 'book', () => {
    const moved = new Map<number, ReturnType<typeof sheet.cells.get>>();
    const keep = new Map(sheet.cells);
    const inBand = (row: number, col: number) => (dir === 'down' || dir === 'up' ? col >= r.c1 && col <= r.c2 : row >= r.r1 && row <= r.r2);
    for (const [k, c] of keep) {
      const row = keyRow(k), col = keyCol(k);
      if (!inBand(row, col)) { moved.set(k, c); continue; }
      let nr = row, nc = col;
      if (dir === 'down' && row >= r.r1) nr = row + h;
      else if (dir === 'right' && col >= r.c1) nc = col + w;
      else if (dir === 'up') { if (row >= r.r1 && row <= r.r2) continue; if (row > r.r2) nr = row - h; }
      else if (dir === 'left') { if (col >= r.c1 && col <= r.c2) continue; if (col > r.c2) nc = col - w; }
      if (nr < MAX_ROWS && nc < MAX_COLS) moved.set(cellKey(nr, nc), c);
    }
    sheet.cells = moved as Sheet['cells'];
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
};

// ---------- Sheets ----------

const uniqueName = (store: GridStore, base: string): string => {
  const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  const stem = base.replace(/\s*\(\d+\)$/, '');
  for (let i = 2; ; i++) { const n = `${stem} (${i})`; if (!taken.has(n.toLowerCase())) return n; }
};

export const nextSheetName = (store: GridStore): string => {
  const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
  for (let i = store.wb.sheets.length + 1; ; i++) if (!taken.has(`sheet${i}`)) return `Sheet${i}`;
};

export const validSheetName = (store: GridStore, name: string, except?: number): string | null => {
  const n = name.trim();
  if (!n) return 'A sheet name can\'t be empty.';
  if (n.length > 31) return 'Sheet names can be at most 31 characters.';
  if (/[\\/?*[\]:]/.test(n) || n.startsWith("'") || n.endsWith("'")) return 'Sheet names can\'t contain \\ / ? * [ ] : or start or end with an apostrophe.';
  if (n.toLowerCase() === 'history') return '"History" is a reserved name.';
  if (store.wb.sheets.some((s, i) => i !== except && s.name.toLowerCase() === n.toLowerCase())) return 'Another sheet already has that name.';
  return null;
};

export const addSheet = (store: GridStore, at = store.sheetIndex + 1, name = nextSheetName(store)): void => {
  store.change('Insert sheet', 'book', () => {
    store.wb.sheets.splice(at, 0, newSheet(name));
    fixSheetIndexes(store, (i) => (i >= at ? i + 1 : i));
    store.wb.activeSheet = at;
    store.engine.sheetsChanged();
  });
  store.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
};

export const duplicateSheet = (store: GridStore, i = store.sheetIndex): void => {
  const src = store.wb.sheets[i]!;
  store.change('Duplicate sheet', 'book', () => {
    const copy = cloneSheet(src);
    copy.name = uniqueName(store, src.name);
    copy.tables = copy.tables.map((t) => ({ ...t, name: `${t.name}_${Math.random().toString(36).slice(2, 6)}` }));
    store.wb.sheets.splice(i + 1, 0, copy);
    fixSheetIndexes(store, (x) => (x > i ? x + 1 : x));
    store.wb.activeSheet = i + 1;
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
};

export const renameSheet = (store: GridStore, i: number, name: string): boolean => {
  const err = validSheetName(store, name, i);
  if (err) { store.notify(err); return false; }
  const old = store.wb.sheets[i]!.name;
  if (old === name.trim()) return true;
  store.change('Rename sheet', 'book', () => {
    const nn = name.trim();
    for (const s of store.wb.sheets) {
      for (const c of s.cells.values()) if (c.f) c.f = renameSheetInFormula(c.f, old, nn);
      for (const ch of s.charts) for (const ser of ch.series) {
        ser.values = renameSheetInFormula(ser.values, old, nn);
        if (ser.categories) ser.categories = renameSheetInFormula(ser.categories, old, nn);
        if (ser.name) ser.name = renameSheetInFormula(ser.name, old, nn);
      }
    }
    for (const n of store.wb.names) n.formula = renameSheetInFormula(n.formula, old, nn);
    store.wb.sheets[i]!.name = nn;
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
  return true;
};

export const deleteSheet = (store: GridStore, i = store.sheetIndex): void => {
  if (store.wb.sheets.filter((s) => s.state === 'visible').length <= 1 && store.wb.sheets[i]!.state === 'visible') { store.notify('A workbook needs at least one visible sheet.'); return; }
  const name = store.wb.sheets[i]!.name;
  store.change('Delete sheet', 'book', () => {
    store.wb.sheets.splice(i, 1);
    for (const s of store.wb.sheets) for (const c of s.cells.values()) if (c.f) c.f = deleteSheetInFormula(c.f, name);
    store.wb.names = store.wb.names.filter((n) => n.sheet !== i).map((n) => ({ ...n, formula: deleteSheetInFormula(n.formula, name) }));
    fixSheetIndexes(store, (x) => (x > i ? x - 1 : x));
    let a = Math.min(i, store.wb.sheets.length - 1);
    while (store.wb.sheets[a]?.state !== 'visible' && a > 0) a--;
    store.wb.activeSheet = a;
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
};

export const moveSheet = (store: GridStore, from: number, to: number): void => {
  if (from === to) return;
  store.change('Move sheet', 'book', () => {
    const [s] = store.wb.sheets.splice(from, 1);
    store.wb.sheets.splice(to, 0, s!);
    const map = (x: number) => (x === from ? to : from < to && x > from && x <= to ? x - 1 : from > to && x >= to && x < from ? x + 1 : x);
    fixSheetIndexes(store, map);
    store.wb.activeSheet = to;
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
};

const fixSheetIndexes = (store: GridStore, map: (i: number) => number): void => {
  for (const n of store.wb.names) if (n.sheet !== undefined) n.sheet = map(n.sheet);
};

export const setSheetState = (store: GridStore, i: number, state: Sheet['state']): void => {
  if (state !== 'visible' && store.wb.sheets.filter((s) => s.state === 'visible').length <= 1) { store.notify('A workbook needs at least one visible sheet.'); return; }
  store.change(state === 'visible' ? 'Unhide sheet' : 'Hide sheet', 'book', () => {
    store.wb.sheets[i]!.state = state;
    if (state !== 'visible' && store.wb.activeSheet === i) store.wb.activeSheet = store.wb.sheets.findIndex((s) => s.state === 'visible');
    if (state === 'visible') store.wb.activeSheet = i;
  });
};
