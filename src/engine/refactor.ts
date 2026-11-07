// Rewrites formula references when cells are copied, filled, moved, or rows/columns/sheets change.
import { mapNode, nodeText, parseFormula, type Node, type RefNode } from './parser';
import { MAX_COLS, MAX_ROWS } from './types';

const REF_ERR: Node = { type: 'err', code: '#REF!' };

const rewrite = (formula: string, f: (n: Node) => Node): string => {
  try { return nodeText(mapNode(parseFormula(formula), f)); } catch { return formula; }
};

/** Copy/fill: relative parts move by (dRow, dCol); references pushed off the sheet become #REF!. */
export const shiftFormula = (formula: string, dRow: number, dCol: number): string =>
  rewrite(formula, (n) => {
    if (n.type !== 'ref') return n;
    const r = { ...n.ref };
    if (r.kind !== 'cols') { if (!r.r1Abs) r.r1 += dRow; if (!r.r2Abs) r.r2 += dRow; }
    if (r.kind !== 'rows') { if (!r.c1Abs) r.c1 += dCol; if (!r.c2Abs) r.c2 += dCol; }
    if (r.r1 < 0 || r.c1 < 0 || r.r2 >= MAX_ROWS || r.c2 >= MAX_COLS) return REF_ERR;
    if (r.r1 > r.r2) [r.r1, r.r2] = [r.r2, r.r1];
    if (r.c1 > r.c2) [r.c1, r.c2] = [r.c2, r.c1];
    return { type: 'ref', ref: r };
  });

export interface Structural {
  /** Sheet whose rows/columns change. */
  sheet: string;
  axis: 'row' | 'col';
  /** First index affected. */
  at: number;
  /** Positive inserts, negative deletes. */
  count: number;
}

const onSheet = (r: RefNode, formulaSheet: string, target: string): boolean => {
  const lc = target.toLowerCase();
  if (r.sheet === undefined) return formulaSheet.toLowerCase() === lc;
  if (r.sheet2 === undefined) return r.sheet.toLowerCase() === lc;
  return r.sheet.toLowerCase() === lc || r.sheet2.toLowerCase() === lc;
};

/** Adjusts one 0-based span [a, b] for an insert/delete at `at` of `count`. Null when fully deleted. */
const adjustSpan = (a: number, b: number, at: number, count: number): [number, number] | null => {
  if (count > 0) {
    if (a >= at) a += count;
    if (b >= at) b += count;
    return [a, b];
  }
  const del = -count, end = at + del - 1;
  if (a >= at && b <= end) return null;
  const shift = (x: number) => (x > end ? x - del : x >= at ? at : x);
  const na = a > end ? a - del : a >= at ? at : a;
  const nb = b > end ? b - del : b >= at ? at - 1 : b;
  void shift;
  return nb < na ? null : [na, nb];
};

/** Insert or delete rows/columns on a sheet, for a formula living on `formulaSheet`. */
export const adjustForStructure = (formula: string, formulaSheet: string, op: Structural): string =>
  rewrite(formula, (n) => {
    if (n.type !== 'ref' || n.ref.external || !onSheet(n.ref, formulaSheet, op.sheet)) return n;
    const r = { ...n.ref };
    if (op.axis === 'row') {
      if (r.kind === 'cols') return n;
      const s = adjustSpan(r.r1, r.r2, op.at, op.count);
      if (!s) return REF_ERR;
      [r.r1, r.r2] = s;
      if (r.r2 >= MAX_ROWS) r.r2 = MAX_ROWS - 1;
      if (r.r1 >= MAX_ROWS) return REF_ERR;
    } else {
      if (r.kind === 'rows') return n;
      const s = adjustSpan(r.c1, r.c2, op.at, op.count);
      if (!s) return REF_ERR;
      [r.c1, r.c2] = s;
      if (r.c2 >= MAX_COLS) r.c2 = MAX_COLS - 1;
      if (r.c1 >= MAX_COLS) return REF_ERR;
    }
    return { type: 'ref', ref: r };
  });

/** Cut and paste: references into the moved block follow it. */
export const adjustForMove = (
  formula: string, formulaSheet: string,
  from: { sheet: string; r1: number; c1: number; r2: number; c2: number }, to: { sheet: string; row: number; col: number },
): string => {
  const dR = to.row - from.r1, dC = to.col - from.c1;
  return rewrite(formula, (n) => {
    if (n.type !== 'ref' || n.ref.external || n.ref.sheet2 || !onSheet(n.ref, formulaSheet, from.sheet)) return n;
    const r = n.ref;
    if (r.r1 < from.r1 || r.r2 > from.r2 || r.c1 < from.c1 || r.c2 > from.c2) return n;
    const moved: RefNode = { ...r, r1: r.r1 + dR, r2: r.r2 + dR, c1: r.c1 + dC, c2: r.c2 + dC };
    if (to.sheet.toLowerCase() !== from.sheet.toLowerCase()) moved.sheet = to.sheet.toLowerCase() === formulaSheet.toLowerCase() ? undefined : to.sheet;
    return { type: 'ref', ref: moved };
  });
};

export const renameSheetInFormula = (formula: string, oldName: string, newName: string): string => {
  const lc = oldName.toLowerCase();
  return rewrite(formula, (n) => {
    if (n.type === 'ref' && !n.ref.external) {
      const r = { ...n.ref };
      if (r.sheet?.toLowerCase() === lc) r.sheet = newName;
      if (r.sheet2?.toLowerCase() === lc) r.sheet2 = newName;
      return { type: 'ref', ref: r };
    }
    if (n.type === 'name' && n.sheet?.toLowerCase() === lc) return { ...n, sheet: newName };
    return n;
  });
};

export const deleteSheetInFormula = (formula: string, name: string): string => {
  const lc = name.toLowerCase();
  return rewrite(formula, (n) => {
    if (n.type === 'ref' && n.ref.sheet?.toLowerCase() === lc && !n.ref.sheet2) return REF_ERR;
    return n;
  });
};

/** Does the formula mention this sheet at all (for warnings before deleting). */
export const mentionsSheet = (formula: string, name: string): boolean => {
  let hit = false;
  const lc = name.toLowerCase();
  try {
    mapNode(parseFormula(formula), (n) => {
      if (n.type === 'ref' && (n.ref.sheet?.toLowerCase() === lc || n.ref.sheet2?.toLowerCase() === lc)) hit = true;
      return n;
    });
  } catch { /* unreadable */ }
  return hit;
};
