// A small in-memory workbook for testing worksheet functions on their own, without the evaluator.
import type { FnContext, LambdaValue, Matrix, RangeRef, Scalar, Value } from '../src/engine/types';

export const makeCtx = (sheets: Scalar[][][] = [[]], names = ['Sheet1', 'Sheet2', 'Sheet3']): FnContext => ({
  sheet: 0, row: 0, col: 0, date1904: false,
  values(r: RangeRef): Matrix {
    const g = sheets[r.sheet] ?? [];
    const lastRow = Math.min(r.r2, Math.max(r.r1, g.length - 1));
    const width = Math.max(0, ...g.map((x) => x.length));
    const lastCol = Math.min(r.c2, Math.max(r.c1, width - 1));
    const out: Matrix = [];
    for (let i = r.r1; i <= lastRow; i++) {
      const row: Scalar[] = [];
      for (let j = r.c1; j <= lastCol; j++) row.push(g[i]?.[j] ?? null);
      out.push(row);
    }
    return out;
  },
  cell: (s, r, c) => sheets[s]?.[r]?.[c] ?? null,
  usedBounds: (s) => ({ rows: (sheets[s]?.length ?? 0) - 1, cols: Math.max(-1, ...(sheets[s] ?? []).map((x) => x.length - 1)) }),
  formula: () => null,
  sheetIndex: (n) => names.findIndex((x) => x.toLowerCase() === n.toLowerCase()),
  sheetName: (i) => names[i] ?? '',
  sheetCount: () => sheets.length,
  call: (fn: LambdaValue, args: Value[]) => fn.invoke(args),
  rowHidden: () => false,
  isSubtotalCell: () => false,
  numberFormat: () => 'General',
});

export const lambda = (f: (...a: Value[]) => Value, params = ['x']): LambdaValue => ({ kind: 'lambda', params, invoke: (a) => f(...a) });
