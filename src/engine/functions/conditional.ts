import { CellError, Err, isError, isLambda, isMatrix, isRef, isRefs, ref, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { first, num, toNumber } from '../coerce';
import { makeCriteria, toGrid, type Grid } from './criteria';

type R<T> = T | CellError;
type Pred = (v: Scalar) => boolean;

const crit = (v: Value, ctx: FnContext): Pred => makeCriteria(first(v, ctx), ctx.date1904);

/** Positions 1, 3, 5... (or 2, 4, 6...) for the criteria arguments, which spill when given arrays. */
const every2 = (from: number): number[] => Array.from({ length: Math.ceil((255 - from) / 2) }, (_, i) => from + i * 2);

/** Sum range resized to the criteria range's shape from its top-left cell, as spreadsheets do. */
const resized = (target: Value, like: Grid): Value =>
  isRef(target) ? ref(target.sheet, target.r1, target.c1, target.r1 + like.fullRows - 1, target.c1 + like.fullCols - 1) : target;

interface Scan { grids: Grid[]; preds: Pred[] }

const prepare = (pairs: [Value, Value][], ctx: FnContext): R<Scan> => {
  const grids: Grid[] = [], preds: Pred[] = [];
  for (const [rg, c] of pairs) {
    if (isRefs(rg) || isLambda(rg)) return Err.value;
    const g = toGrid(rg, ctx);
    if (isError(g)) return g;
    if (grids.length && (g.fullRows !== grids[0]!.fullRows || g.fullCols !== grids[0]!.fullCols)) return Err.value;
    grids.push(g);
    preds.push(crit(c, ctx));
  }
  return { grids, preds };
};

/**
 * Visits each cell position where every criterion holds. Only the stored part of each range is
 * walked; the count of matching blank cells past it is returned for COUNTIF(S).
 */
const scan = ({ grids, preds }: Scan, visit: (i: number, j: number) => void, extra: Grid[] = []): number => {
  const all = [...grids, ...extra];
  const g0 = grids[0]!;
  const rows = Math.min(g0.fullRows, Math.max(...all.map((g) => g.rows)));
  const cols = Math.min(g0.fullCols, Math.max(...all.map((g) => g.cols)));
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++)
      if (preds.every((p, k) => p(grids[k]!.at(i, j)))) visit(i, j);
  const rest = g0.fullRows * g0.fullCols - rows * cols;
  return rest > 0 && preds.every((p) => p(null)) ? rest : 0;
};

type Agg = 'sum' | 'avg' | 'max' | 'min' | 'count';

const finish = (agg: Agg, xs: number[]): Scalar => {
  if (agg === 'sum') return xs.reduce((a, b) => a + b, 0);
  if (agg === 'avg') return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : Err.div0;
  if (agg === 'max') return xs.length ? Math.max(...xs) : 0;
  return xs.length ? Math.min(...xs) : 0;
};

/** Shared body of the *IFS functions: `target` is the range aggregated (null for COUNTIFS). */
const ifs = (agg: Agg, target: Value | null, pairs: [Value, Value][], ctx: FnContext): Scalar => {
  const s = prepare(pairs, ctx);
  if (isError(s)) return s;
  if (agg === 'count') { let n = 0; const rest = scan(s, () => n++); return n + rest; }
  if (target === null || isRefs(target) || isLambda(target)) return Err.value;
  const t = toGrid(target, ctx);
  if (isError(t)) return t;
  const g0 = s.grids[0]!;
  if (t.fullRows !== g0.fullRows || t.fullCols !== g0.fullCols) return Err.value;
  const xs: number[] = [];
  let err: CellError | null = null;
  scan(s, (i, j) => {
    const v = t.at(i, j);
    if (typeof v === 'number') xs.push(v);
    else if (isError(v) && !err) err = v;
  }, [t]);
  return err ?? finish(agg, xs);
};

const pairsFrom = (a: Value[], start: number): R<[Value, Value][]> => {
  if ((a.length - start) % 2 !== 0) return Err.value;
  const out: [Value, Value][] = [];
  for (let i = start; i < a.length; i += 2) out.push([a[i]!, a[i + 1]!]);
  return out;
};

const ifsDef = (name: string, agg: Agg): FunctionDef => ({
  name, min: agg === 'count' ? 2 : 3, max: 255, scalar: every2(agg === 'count' ? 1 : 2),
  fn: (a, ctx) => {
    const p = pairsFrom(a, agg === 'count' ? 0 : 1);
    return isError(p) ? p : ifs(agg, agg === 'count' ? null : a[0]!, p, ctx);
  },
});

const ifDef = (name: string, agg: 'sum' | 'avg'): FunctionDef => ({
  name, min: 2, max: 3, scalar: [1],
  fn: (a, ctx) => {
    const rg = a[0]!;
    if (a[2] === undefined || a[2] === null) return ifs(agg, rg, [[rg, a[1]!]], ctx);
    const g = toGrid(rg, ctx);
    if (isError(g)) return g;
    const t = resized(a[2], g);
    // Arrays (not references) cannot be resized; spreadsheets require references here anyway.
    if (!isRef(t)) {
      const tg = toGrid(t, ctx);
      if (isError(tg)) return tg;
      if (tg.fullRows !== g.fullRows || tg.fullCols !== g.fullCols) return Err.value;
    }
    return ifs(agg, t, [[rg, a[1]!]], ctx);
  },
});

const COUNTIF: FunctionDef = {
  name: 'COUNTIF', min: 2, max: 2, scalar: [1],
  fn: (a, ctx) => ifs('count', null, [[a[0]!, a[1]!]], ctx),
};

const SUMPRODUCT: FunctionDef = {
  name: 'SUMPRODUCT', min: 1, max: 255,
  fn: (a, ctx) => {
    const gs: Grid[] = [];
    for (const v of a) {
      if (isRefs(v) || isLambda(v)) return Err.value;
      const g = toGrid(v, ctx);
      if (isError(g)) return g;
      if (gs.length && (g.fullRows !== gs[0]!.fullRows || g.fullCols !== gs[0]!.fullCols)) return Err.value;
      gs.push(g);
    }
    const rows = Math.min(gs[0]!.fullRows, Math.max(...gs.map((g) => g.rows)));
    const cols = Math.min(gs[0]!.fullCols, Math.max(...gs.map((g) => g.cols)));
    let sum = 0;
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        let p = 1;
        for (const g of gs) {
          const v = g.at(i, j);
          if (isError(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        sum += p;
      }
    return sum;
  },
};

// ---------------------------------------------------------------- statistics for SUBTOTAL / AGGREGATE

const sum = (xs: number[]): number => xs.reduce((a, b) => a + b, 0);
const mean = (xs: number[]): number => sum(xs) / xs.length;
const variance = (xs: number[], sample: boolean): R<number> => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return Err.div0;
  const m = mean(xs);
  return sum(xs.map((x) => (x - m) ** 2)) / (sample ? n - 1 : n);
};
const sorted = (xs: number[]): number[] => [...xs].sort((x, y) => x - y);

const percentileInc = (xs: number[], k: number): R<number> => {
  if (!xs.length || k < 0 || k > 1) return Err.num;
  const s = sorted(xs), h = (s.length - 1) * k, lo = Math.floor(h);
  return lo + 1 < s.length ? s[lo]! + (h - lo) * (s[lo + 1]! - s[lo]!) : s[lo]!;
};
const percentileExc = (xs: number[], k: number): R<number> => {
  const n = xs.length, h = (n + 1) * k;
  if (!n || h < 1 || h > n) return Err.num;
  const s = sorted(xs), lo = Math.floor(h);
  return lo < n ? s[lo - 1]! + (h - lo) * (s[lo]! - s[lo - 1]!) : s[n - 1]!;
};
const quartile = (xs: number[], q: number, inc: boolean): R<number> => {
  const k = Math.trunc(q);
  if (inc ? k < 0 || k > 4 : k < 1 || k > 3) return Err.num;
  return inc ? percentileInc(xs, k / 4) : percentileExc(xs, k / 4);
};
const mode = (xs: number[]): R<number> => {
  const counts = new Map<number, number>();
  let best: number | null = null, bestN = 1;
  for (const x of xs) {
    const n = (counts.get(x) ?? 0) + 1;
    counts.set(x, n);
    if (n > bestN) { bestN = n; best = x; }
  }
  // Ties go to the value seen first, so re-scan in order.
  if (best === null) return Err.na;
  return xs.find((x) => counts.get(x) === bestN)!;
};

/** Functions 1-13 on numbers; COUNTA (3) gets its own count. */
const basic = (fn: number, xs: number[], countA: number): R<number> => {
  switch (fn) {
    case 1: return xs.length ? mean(xs) : Err.div0;
    case 2: return xs.length;
    case 3: return countA;
    case 4: return xs.length ? Math.max(...xs) : 0;
    case 5: return xs.length ? Math.min(...xs) : 0;
    case 6: return xs.length ? xs.reduce((a, b) => a * b, 1) : 0;
    case 7: { const v = variance(xs, true); return isError(v) ? v : Math.sqrt(v); }
    case 8: { const v = variance(xs, false); return isError(v) ? v : Math.sqrt(v); }
    case 9: return sum(xs);
    case 10: return variance(xs, true);
    case 11: return variance(xs, false);
    case 12: return xs.length ? percentileInc(xs, 0.5) : Err.num;
    case 13: return mode(xs);
    default: return Err.value;
  }
};

interface Gather { skipHidden: boolean; skipErrors: boolean; skipNested: boolean; skipFiltered?: boolean }

/** Numbers and the COUNTA count from reference/array arguments, honouring the skip options. */
const gather = (args: Value[], ctx: FnContext, o: Gather): R<{ xs: number[]; countA: number }> => {
  const xs: number[] = [];
  let countA = 0;
  const take = (v: Scalar): CellError | null => {
    if (v === null) return null;
    if (isError(v)) return o.skipErrors ? null : v;
    countA++;
    if (typeof v === 'number') xs.push(v);
    return null;
  };
  let firstErr: CellError | null = null;
  const areas = (v: Value): Value[] => (isRefs(v) ? [...v.areas] : [v]);
  for (const arg of args) for (const v of areas(arg)) {
    if (isRef(v)) {
      const m = ctx.values(v);
      m.forEach((row, i) => {
        const r = v.r1 + i;
        if ((o.skipHidden && ctx.rowHidden(v.sheet, r)) || (o.skipFiltered && ctx.rowFiltered?.(v.sheet, r))) return;
        row.forEach((c, j) => {
          if (o.skipNested && ctx.isSubtotalCell(v.sheet, r, v.c1 + j)) return;
          const e = take(c);
          if (e && !firstErr) firstErr = e;
        });
      });
    } else if (isMatrix(v)) {
      for (const row of v) for (const c of row) { const e = take(c); if (e && !firstErr) firstErr = e; }
    } else if (isLambda(v)) return Err.value;
    else {
      const s = v as Scalar;
      const n = toNumber(s, ctx.date1904);
      if (isError(n)) { if (!o.skipErrors) return n; continue; }
      if (s !== null) { countA++; xs.push(n); }
    }
  }
  return firstErr ?? { xs, countA };
};

/** COUNT and COUNTA ignore errors in their data; every other function returns the first one. */
const run = (fn: number, args: Value[], ctx: FnContext, o: Gather): Scalar => {
  if (fn === 3 && !o.skipErrors) return countAll(args, ctx, o);
  const g = gather(args, ctx, { ...o, skipErrors: o.skipErrors || fn === 2 });
  return isError(g) ? g : basic(fn, g.xs, g.countA);
};

/** COUNTA including error cells. */
const countAll = (args: Value[], ctx: FnContext, o: Gather): number => {
  let n = 0;
  for (const arg of args) for (const v of isRefs(arg) ? arg.areas : [arg]) {
    if (isRef(v)) {
      ctx.values(v).forEach((row, i) => {
        const r = v.r1 + i;
        if ((o.skipHidden && ctx.rowHidden(v.sheet, r)) || (o.skipFiltered && ctx.rowFiltered?.(v.sheet, r))) return;
        row.forEach((c, j) => { if (c !== null && !(o.skipNested && ctx.isSubtotalCell(v.sheet, r, v.c1 + j))) n++; });
      });
    } else if (isMatrix(v)) n += v.flat().filter((c) => c !== null).length;
    else if (v !== null) n++;
  }
  return n;
};

const SUBTOTAL: FunctionDef = {
  name: 'SUBTOTAL', min: 2, max: 255, scalar: [0],
  fn: (a, ctx) => {
    const f = num(a[0]!, ctx);
    if (isError(f)) return f;
    const k = Math.trunc(f);
    const fn = k > 100 ? k - 100 : k;
    if (fn < 1 || fn > 11) return Err.value;
    for (const v of a.slice(1)) if (!isRef(v) && !isRefs(v)) return Err.value;
    // 1-11 skip rows a filter hides; 101-111 also skip rows hidden by hand.
    return run(fn, a.slice(1), ctx, { skipHidden: k > 100, skipFiltered: true, skipErrors: false, skipNested: true });
  },
};

const AGGREGATE: FunctionDef = {
  name: 'AGGREGATE', min: 3, max: 255, scalar: [0, 1],
  fn: (a, ctx) => {
    const f = num(a[0]!, ctx), op = num(a[1]!, ctx);
    if (isError(f)) return f; if (isError(op)) return op;
    const fn = Math.trunc(f), opt = Math.trunc(op);
    if (fn < 1 || fn > 19 || opt < 0 || opt > 7) return Err.value;
    const o: Gather = {
      skipNested: opt <= 3,
      skipHidden: opt === 1 || opt === 3 || opt === 5 || opt === 7,
      skipErrors: opt === 2 || opt === 3 || opt === 6 || opt === 7,
    };
    if (fn <= 13) {
      // The reference form takes references only; computed arrays are #VALUE!.
      for (const v of a.slice(2)) if (!isRef(v) && !isRefs(v)) return Err.value;
      return run(fn, a.slice(2), ctx, o);
    }
    if (a.length !== 4) return Err.value;
    const k = num(a[3]!, ctx);
    if (isError(k)) return k;
    const g = gather([a[2]!], ctx, o);
    if (isError(g)) return g;
    const xs = g.xs;
    switch (fn) {
      case 14: case 15: {
        const n = Math.ceil(k);
        if (n < 1 || n > xs.length) return Err.num;
        const s = sorted(xs);
        return fn === 14 ? s[s.length - n]! : s[n - 1]!;
      }
      case 16: return percentileInc(xs, k);
      case 17: return quartile(xs, k, true);
      case 18: return percentileExc(xs, k);
      default: return quartile(xs, k, false);
    }
  },
};

export const defs: FunctionDef[] = [
  ifDef('AVERAGEIF', 'avg'), ifsDef('AVERAGEIFS', 'avg'), COUNTIF, ifsDef('COUNTIFS', 'count'),
  ifsDef('MAXIFS', 'max'), ifsDef('MINIFS', 'min'), ifDef('SUMIF', 'sum'), ifsDef('SUMIFS', 'sum'),
  SUMPRODUCT, SUBTOTAL, AGGREGATE,
];
