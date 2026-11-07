import { adjustForMove, shiftFormula } from '../../engine/refactor';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { cellKey, type Cell, type CellRange } from '../../model/workbook';
import type { GridStore, Pos } from '../store';

/** Drag-and-drop of the selection: moves (or copies with Ctrl) cells, keeping references pointed at moved data. */
export const moveRange = (store: GridStore, from: CellRange, to: Pos, copy = false): void => {
  const h = from.r2 - from.r1, w = from.c2 - from.c1;
  const target: CellRange = { r1: to.row, c1: to.col, r2: to.row + h, c2: to.col + w };
  if (target.r2 >= MAX_ROWS || target.c2 >= MAX_COLS) return;
  const occupied = !copy && [...store.sheet.cells.keys()].some((k) => {
    const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
    const inT = r >= target.r1 && r <= target.r2 && c >= target.c1 && c <= target.c2;
    const inF = r >= from.r1 && r <= from.r2 && c >= from.c1 && c <= from.c2;
    return inT && !inF && store.value(r, c) !== null;
  });
  if (occupied && !confirm("There's already data here. Do you want to replace it?")) return;
  const sheet = store.sheetIndex;
  store.change(copy ? 'Copy' : 'Move', 'book', () => {
    const cells: { r: number; c: number; cell: Cell }[] = [];
    for (let r = from.r1; r <= from.r2; r++) for (let c = from.c1; c <= from.c2; c++) {
      const cell = store.sheet.cells.get(cellKey(r, c));
      if (cell) cells.push({ r, c, cell: structuredClone(cell) });
    }
    if (!copy) for (const { r, c } of cells) store.engine.clearCell(sheet, r, c);
    for (let r = target.r1; r <= target.r2; r++) for (let c = target.c1; c <= target.c2; c++) store.engine.clearCell(sheet, r, c);
    for (const { r, c, cell } of cells) {
      const nr = r - from.r1 + to.row, nc = c - from.c1 + to.col;
      if (cell.f) {
        store.engine.setFormula(sheet, nr, nc, copy ? shiftFormula(cell.f, nr - r, nc - c) : cell.f);
        const placed = store.sheet.cells.get(cellKey(nr, nc))!;
        placed.s = cell.s;
      } else {
        store.sheet.cells.set(cellKey(nr, nc), cell);
        store.engine.markChanged(sheet, nr, nc);
      }
    }
    if (!copy) {
      const f = { sheet: store.sheet.name, ...from };
      const t = { sheet: store.sheet.name, row: to.row, col: to.col };
      for (const s of store.wb.sheets) for (const c of s.cells.values()) if (c.f) c.f = adjustForMove(c.f, s.name, f, t);
      store.sheet.merges = store.sheet.merges.map((m) => (m.r1 >= from.r1 && m.r2 <= from.r2 && m.c1 >= from.c1 && m.c2 <= from.c2 ? { r1: m.r1 - from.r1 + to.row, c1: m.c1 - from.c1 + to.col, r2: m.r2 - from.r1 + to.row, c2: m.c2 - from.c1 + to.col } : m));
    }
    store.engine.sheetsChanged();
    store.engine.recalc(true);
  });
  store.selectRange(target);
};
