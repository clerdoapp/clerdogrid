import { Err, ref, type CellError, type FnContext, type RangeRef } from './types';

/**
 * Drops blank rows and columns from the edges of a range, as TRIMRANGE and the trim-reference operators
 * do: `rows`/`cols` are 0 none, 1 leading, 2 trailing, 3 both. An all-blank range is #CALC!.
 */
export const trimRange = (r: RangeRef, ctx: FnContext, rows: number, cols: number): RangeRef | CellError => {
  const m = ctx.values(r);
  const blankRow = (i: number): boolean => (m[i] ?? []).every((c) => c === null);
  const blankCol = (j: number): boolean => m.every((row) => (row[j] ?? null) === null);
  let i1 = 0, i2 = r.r2 - r.r1, j1 = 0, j2 = r.c2 - r.c1;
  // Values come back trimmed to the used area, so trailing trims start from what came back.
  if (rows & 2) { i2 = Math.min(i2, m.length - 1); while (i2 >= 0 && blankRow(i2)) i2--; }
  if (rows & 1) while (i1 <= i2 && blankRow(i1)) i1++;
  if (cols & 2) { j2 = Math.min(j2, (m[0]?.length ?? 0) - 1); while (j2 >= 0 && blankCol(j2)) j2--; }
  if (cols & 1) while (j1 <= j2 && blankCol(j1)) j1++;
  if (i1 > i2 || j1 > j2) return Err.calc;
  return ref(r.sheet, r.r1 + i1, r.c1 + j1, r.r1 + i2, r.c1 + j2);
};
