import { CellError, Err, isError, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { cellsOf, collectNumbers, fromRange, num, optBool, optNum, toMatrix, type CollectOptions } from '../coerce';
import {
  betaIncInv, betaIncPair, betaPdf, chisqSf, fCdfPair, fPdf, gamma, gammaInv, gammaInvUpper, gammaln, gammaP, gammaPdf, gammaQ,
  lnChoose, normCdf, normHalf, normInv, normPdf, solvePositive, tCdf, tInv, tPdf, xlogy,
} from './special';
import { confidenceHalfWidth, detectSeason, etsStatistic, fitEts, forecastSteps, MAX_SEASON, prepareSeries, stepsAhead, type EtsModel, type EtsSeries } from '../ets';

type Num = number | CellError;
type Fn = FunctionDef['fn'];

const ok = (x: number): Num => (Number.isFinite(x) ? x : Err.num);
const fail = (v: unknown): v is CellError => v instanceof CellError;
const A_OPTS: CollectOptions = { boolsInRanges: true, textInRangesAsZero: true };

const sum = (xs: readonly number[]): number => {
  // Neumaier summation keeps results like DEVSQ exact for typical worksheet data.
  let s = 0, c = 0;
  for (const x of xs) {
    const t = s + x;
    c += Math.abs(s) >= Math.abs(x) ? s - t + x : x - t + s;
    s = t;
  }
  return s + c;
};
const mean = (xs: readonly number[]): number => sum(xs) / xs.length;
const devsq = (xs: readonly number[]): number => {
  const m = mean(xs);
  return sum(xs.map((x) => (x - m) * (x - m)));
};
const sorted = (xs: readonly number[]): number[] => [...xs].sort((a, b) => a - b);

/** Element-wise numeric function: every argument is coerced to a number; `dflt` fills omitted optional ones. */
const scalarFn = (name: string, min: number, max: number, f: (...x: number[]) => Num, dflt: (number | undefined)[] = []): FunctionDef => ({
  name, min, max, scalar: true,
  fn: (args, ctx) => {
    const xs: number[] = [];
    for (let i = 0; i < max; i++) {
      const a = args[i];
      const d = dflt[i];
      if (a === undefined || (a === null && d !== undefined)) { xs.push(d ?? NaN); continue; }
      const n = num(a, ctx);
      if (fail(n)) return n;
      xs.push(n);
    }
    const r = f(...xs);
    return fail(r) ? r : ok(r);
  },
});

/** Aggregate over every argument's numbers. */
const aggFn = (name: string, f: (xs: number[]) => Num, opt: CollectOptions = {}, min = 1): FunctionDef => ({
  name, min, max: 255,
  fn: (args, ctx) => {
    const xs = collectNumbers(args, ctx, opt);
    if (fail(xs)) return xs;
    const r = f(xs);
    return fail(r) ? r : ok(r);
  },
});

const data = (v: Value, ctx: FnContext): number[] | CellError => collectNumbers([v], ctx);

/** Paired numeric data from two equally sized sets; pairs where either side is not a number are skipped. */
const pairs = (a: Value, b: Value, ctx: FnContext): [number[], number[]] | CellError => {
  const xa = [...cellsOf(a, ctx)], xb = [...cellsOf(b, ctx)];
  if (xa.length !== xb.length) return Err.na;
  const p: number[] = [], q: number[] = [];
  for (let i = 0; i < xa.length; i++) {
    const u = xa[i]!, v = xb[i]!;
    if (isError(u)) return u;
    if (isError(v)) return v;
    if (typeof u === 'number' && typeof v === 'number') { p.push(u); q.push(v); }
  }
  return [p, q];
};

const pairFn = (name: string, f: (x: number[], y: number[]) => Num, swap = false): FunctionDef => ({
  name, min: 2, max: 2,
  fn: (args, ctx) => {
    const pr = pairs(args[0]!, args[1]!, ctx);
    if (fail(pr)) return pr;
    const r = swap ? f(pr[1], pr[0]) : f(pr[0], pr[1]);
    return fail(r) ? r : ok(r);
  },
});

// ---------- descriptive ----------

const variance = (xs: number[], sample: boolean): Num => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return Err.div0;
  return devsq(xs) / (sample ? n - 1 : n);
};
const stdev = (xs: number[], sample: boolean): Num => {
  const v = variance(xs, sample);
  return fail(v) ? v : Math.sqrt(v);
};

const median = (xs: number[]): Num => {
  if (!xs.length) return Err.num;
  const s = sorted(xs), h = s.length >> 1;
  return s.length % 2 ? s[h]! : (s[h - 1]! + s[h]!) / 2;
};

const modes = (xs: number[]): number[] => {
  const counts = new Map<number, number>();
  for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
  let best = 1;
  for (const c of counts.values()) best = Math.max(best, c);
  if (best < 2) return [];
  return [...counts].filter(([, c]) => c === best).map(([v]) => v);
};

const skew = (xs: number[]): Num => {
  const n = xs.length;
  if (n < 3) return Err.div0;
  const m = mean(xs), s = Math.sqrt(devsq(xs) / (n - 1));
  if (s === 0) return Err.div0;
  return (n / ((n - 1) * (n - 2))) * sum(xs.map((x) => ((x - m) / s) ** 3));
};
const skewP = (xs: number[]): Num => {
  const n = xs.length;
  if (n < 1) return Err.div0;
  const m = mean(xs), s = Math.sqrt(devsq(xs) / n);
  if (s === 0) return Err.div0;
  return sum(xs.map((x) => ((x - m) / s) ** 3)) / n;
};
const kurt = (xs: number[]): Num => {
  const n = xs.length;
  if (n < 4) return Err.div0;
  const m = mean(xs), s = Math.sqrt(devsq(xs) / (n - 1));
  if (s === 0) return Err.div0;
  const s4 = sum(xs.map((x) => ((x - m) / s) ** 4));
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * s4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
};

const percentileInc = (xs: number[], k: number): Num => {
  if (!xs.length || k < 0 || k > 1) return Err.num;
  const s = sorted(xs), h = k * (s.length - 1), lo = Math.floor(h);
  return lo + 1 < s.length ? s[lo]! + (h - lo) * (s[lo + 1]! - s[lo]!) : s[lo]!;
};
const percentileExc = (xs: number[], k: number): Num => {
  const n = xs.length;
  if (!n || k <= 0 || k >= 1) return Err.num;
  const h = k * (n + 1);
  if (h < 1 || h > n) return Err.num;
  const s = sorted(xs), lo = Math.floor(h);
  return lo < n ? s[lo - 1]! + (h - lo) * (s[lo]! - s[lo - 1]!) : s[n - 1]!;
};

/** PERCENTRANK: interpolates between the ranks of neighbouring values and truncates to `sig` digits. */
const percentRank = (xs: number[], x: number, sig: number, exc: boolean): Num => {
  sig = Math.trunc(sig);
  if (sig < 1) return Err.num;
  const n = xs.length;
  if (!n) return Err.num;
  const s = sorted(xs);
  if (x < s[0]! || x > s[n - 1]!) return Err.na;
  const rankOf = (v: number): number => {
    const below = s.filter((y) => y < v).length;
    return exc ? (below + 1) / (n + 1) : n === 1 ? 1 : below / (n - 1);
  };
  let r: number;
  if (s.includes(x)) r = rankOf(x);
  else {
    let i = 0;
    while (s[i + 1]! < x) i++;
    const lo = s[i]!, hi = s[i + 1]!;
    const rl = rankOf(lo), rh = rankOf(hi);
    r = rl + ((x - lo) / (hi - lo)) * (rh - rl);
  }
  const f = 10 ** sig;
  return Math.floor(r * f * (1 + 1e-12)) / f;
};

const rankFn = (name: string, avg: boolean): FunctionDef => ({
  name, min: 2, max: 3, scalar: [0, 2],
  fn: (args, ctx) => {
    const x = num(args[0]!, ctx);
    if (fail(x)) return x;
    const order = optNum(args[2], ctx, 0);
    if (fail(order)) return order;
    const xs = collectNumbers([args[1]!], ctx, { ignoreErrors: false });
    if (fail(xs)) return xs;
    const asc = order !== 0;
    let better = 0, same = 0;
    for (const v of xs) {
      if (v === x) same++;
      else if (asc ? v < x : v > x) better++;
    }
    if (!same) return Err.na;
    return avg ? better + (same + 1) / 2 : better + 1;
  },
});

const kthFn = (name: string, largest: boolean): FunctionDef => ({
  name, min: 2, max: 2, scalar: [1],
  fn: (args, ctx) => {
    const xs = data(args[0]!, ctx);
    if (fail(xs)) return xs;
    const kn = num(args[1]!, ctx);
    if (fail(kn)) return kn;
    const k = Math.ceil(kn);
    if (k < 1 || k > xs.length) return Err.num;
    const s = sorted(xs);
    return largest ? s[s.length - k]! : s[k - 1]!;
  },
});

const dataScalarFn = (name: string, min: number, max: number, f: (xs: number[], ...p: number[]) => Num, dflt: (number | undefined)[] = []): FunctionDef => ({
  name, min, max, scalar: Array.from({ length: max - 1 }, (_, i) => i + 1),
  fn: (args, ctx) => {
    const xs = data(args[0]!, ctx);
    if (fail(xs)) return xs;
    const ps: number[] = [];
    for (let i = 1; i < max; i++) {
      const d = dflt[i - 1];
      const p = d === undefined ? (args[i] === undefined ? NaN : num(args[i]!, ctx)) : optNum(args[i], ctx, d);
      if (fail(p)) return p;
      ps.push(p);
    }
    const r = f(xs, ...ps);
    return fail(r) ? r : ok(r);
  },
});

const countFn: Fn = (args, ctx) => {
  let n = 0;
  for (const a of args) {
    if (fromRange(a)) { for (const c of cellsOf(a, ctx)) if (typeof c === 'number') n++; }
    else if (a === null || typeof a === 'number' || typeof a === 'boolean') n++;
    else if (typeof a === 'string' && !fail(num(a, ctx))) n++;
  }
  return n;
};
const countaFn: Fn = (args, ctx) => {
  let n = 0;
  for (const a of args) {
    if (fromRange(a)) { for (const c of cellsOf(a, ctx)) if (c !== null) n++; }
    else n++;
  }
  return n;
};
const countblankFn: Fn = (args, ctx) => {
  let n = 0;
  for (const c of cellsOf(args[0]!, ctx)) if (c === null || c === '') n++;
  return n;
};

const frequencyFn: Fn = (args, ctx) => {
  const xs = collectNumbers([args[0]!], ctx);
  if (fail(xs)) return xs;
  const bins = collectNumbers([args[1]!], ctx);
  if (fail(bins)) return bins;
  if (!bins.length) return [[xs.length]];
  const order = bins.map((b, i) => [b, i] as const).sort((p, q) => p[0] - q[0] || p[1] - q[1]);
  const out = new Array<number>(bins.length + 1).fill(0);
  for (const x of xs) {
    const hit = order.find(([b]) => x <= b);
    out[hit ? hit[1] : bins.length]!++;
  }
  return out.map((c) => [c]);
};

const probFn: Fn = (args, ctx) => {
  const xs = [...cellsOf(args[0]!, ctx)], ps = [...cellsOf(args[1]!, ctx)];
  if (xs.length !== ps.length) return Err.na;
  const lo = num(args[2]!, ctx);
  if (fail(lo)) return lo;
  const hi = args[3] === undefined || args[3] === null ? lo : num(args[3], ctx);
  if (fail(hi)) return hi;
  let total = 0, acc = 0;
  for (let i = 0; i < xs.length; i++) {
    const x = xs[i]!, p = ps[i]!;
    if (isError(x)) return x;
    if (isError(p)) return p;
    if (typeof p !== 'number') continue;
    if (p <= 0 || p > 1) return Err.num;
    total += p;
    if (typeof x === 'number' && x >= lo && x <= hi) acc += p;
  }
  if (Math.abs(total - 1) > 1e-7) return Err.num;
  return acc;
};

// ---------- regression (LINEST / LOGEST / TREND / GROWTH) ----------

interface Fit {
  m: number[]; b: number; dropped: boolean[]; se: number[]; seb: number;
  r2: number; sey: number; F: number; df: number; ssreg: number; ssres: number;
}

/**
 * Least squares through modified Gram-Schmidt on centred columns. Collinear columns are dropped
 * (coefficient 0, standard error 0, degrees of freedom raised), which is what current spreadsheets does.
 */
const regress = (Y: number[], X: number[][], cst: boolean): Fit => {
  const n = Y.length, k = X[0]?.length ?? 0;
  const ym = cst ? mean(Y) : 0;
  const xm = Array.from({ length: k }, (_, j) => (cst ? mean(X.map((r) => r[j]!)) : 0));
  const cols = Array.from({ length: k }, (_, j) => X.map((r) => r[j]! - xm[j]!));
  const yc = Y.map((y) => y - ym);
  const Q: number[][] = [], kept: number[] = [], R: number[][] = [];
  const dropped = new Array<boolean>(k).fill(false);
  const dot = (a: number[], b: number[]): number => a.reduce((s, v, i) => s + v * b[i]!, 0);
  for (let j = 0; j < k; j++) {
    const v = [...cols[j]!];
    const n0 = Math.sqrt(dot(v, v));
    const rcol = new Array<number>(kept.length).fill(0);
    for (let pass = 0; pass < 2; pass++) {
      Q.forEach((q, i) => {
        const r = dot(q, v);
        rcol[i]! += r;
        for (let t = 0; t < n; t++) v[t]! -= r * q[t]!;
      });
    }
    const nrm = Math.sqrt(dot(v, v));
    if (n0 === 0 || nrm <= 1e-10 * n0) { dropped[j] = true; continue; }
    Q.push(v.map((x) => x / nrm));
    R.forEach((row, i) => row.push(rcol[i]!));
    R.push([...new Array<number>(kept.length).fill(0), nrm]);
    kept.push(j);
  }
  const p = kept.length;
  const c = Q.map((q) => dot(q, yc));
  const coef = new Array<number>(p).fill(0);
  for (let i = p - 1; i >= 0; i--) {
    let s = c[i]!;
    for (let t = i + 1; t < p; t++) s -= R[i]![t]! * coef[t]!;
    coef[i] = s / R[i]![i]!;
  }
  const m = new Array<number>(k).fill(0);
  kept.forEach((j, i) => (m[j] = coef[i]!));
  const b = cst ? ym - sum(m.map((v, j) => v * xm[j]!)) : 0;
  const res = Y.map((y, i) => y - b - sum(m.map((v, j) => v * X[i]![j]!)));
  const ssres = sum(res.map((e) => e * e));
  const sstot = sum(yc.map((y) => y * y));
  const ssreg = sstot - ssres;
  const df = n - p - (cst ? 1 : 0);
  const sey = Math.sqrt(ssres / df);
  // (XᵀX)⁻¹ = R⁻¹R⁻ᵀ over the kept columns.
  const Ri: number[][] = Array.from({ length: p }, () => new Array<number>(p).fill(0));
  for (let i = p - 1; i >= 0; i--) {
    Ri[i]![i] = 1 / R[i]![i]!;
    for (let j = i + 1; j < p; j++) {
      let s = 0;
      for (let t = i + 1; t <= j; t++) s += R[i]![t]! * Ri[t]![j]!;
      Ri[i]![j] = -s / R[i]![i]!;
    }
  }
  const inv = (a: number, bb: number): number => {
    let s = 0;
    for (let t = Math.max(a, bb); t < p; t++) s += Ri[a]![t]! * Ri[bb]![t]!;
    return s;
  };
  const se = new Array<number>(k).fill(0);
  kept.forEach((j, i) => (se[j] = sey * Math.sqrt(inv(i, i))));
  let q = 1 / n;
  for (let a = 0; a < p; a++) for (let bb = 0; bb < p; bb++) q += xm[kept[a]!]! * xm[kept[bb]!]! * inv(a, bb);
  return { m, b, dropped, se, seb: sey * Math.sqrt(q), r2: ssreg / sstot, sey, F: ssreg / p / (ssres / df), df, ssreg, ssres };
};

interface Design { Y: number[]; X: number[][]; k: number; byRow: boolean; single: boolean; xShape: number[][] | null; yShape: number[][] }

const numMatrix = (v: Value, ctx: FnContext): number[][] | CellError => {
  const m = toMatrix(v, ctx);
  if (fail(m)) return m;
  const out: number[][] = [];
  for (const row of m) {
    const r: number[] = [];
    for (const c of row) {
      if (isError(c)) return c;
      if (typeof c !== 'number') return Err.value;
      r.push(c);
    }
    out.push(r);
  }
  return out;
};

const design = (yv: Value, xv: Value | undefined, ctx: FnContext, logY: boolean): Design | CellError => {
  const ym = numMatrix(yv, ctx);
  if (fail(ym)) return ym;
  const ry = ym.length, cy = ym[0]!.length;
  let Y = ym.flat();
  if (logY) {
    if (Y.some((y) => y <= 0)) return Err.num;
    Y = Y.map(Math.log);
  }
  const yShape = ym;
  if (xv === undefined || xv === null) return { Y, X: Y.map((_, i) => [i + 1]), k: 1, byRow: false, single: true, xShape: null, yShape };
  const xm = numMatrix(xv, ctx);
  if (fail(xm)) return xm;
  const rx = xm.length, cx = xm[0]!.length;
  if (rx === ry && cx === cy) return { Y, X: xm.flat().map((x) => [x]), k: 1, byRow: false, single: true, xShape: xm, yShape };
  if (cy === 1 && rx === ry) return { Y, X: xm, k: cx, byRow: false, single: cx === 1, xShape: xm, yShape };
  if (ry === 1 && cx === cy) return { Y, X: Array.from({ length: cx }, (_, j) => xm.map((r) => r[j]!)), k: rx, byRow: true, single: rx === 1, xShape: xm, yShape };
  return Err.ref;
};

const linestFn = (logY: boolean): Fn => (args, ctx) => {
  const d = design(args[0]!, args[1], ctx, logY);
  if (fail(d)) return d;
  const cst = optBool(args[2], ctx, true);
  if (fail(cst)) return cst;
  const stats = optBool(args[3], ctx, false);
  if (fail(stats)) return stats;
  if (d.Y.length < 1) return Err.value;
  const f = regress(d.Y, d.X, cst);
  const v = (x: number): Scalar => (Number.isFinite(x) ? x : Err.num);
  const t = (x: number): Scalar => v(logY ? Math.exp(x) : x);
  const row1: Scalar[] = [...f.m.map(t).reverse(), t(f.b)];
  if (!stats) return [row1];
  const w = d.k + 1;
  const pad = (xs: Scalar[]): Scalar[] => [...xs, ...new Array<Scalar>(w - xs.length).fill(Err.na)];
  return [
    row1,
    [...f.se.map(v).reverse(), cst ? v(f.seb) : Err.na],
    pad([v(f.r2), v(f.sey)]),
    pad([v(f.F), f.df]),
    pad([v(f.ssreg), v(f.ssres)]),
  ];
};

const trendFn = (logY: boolean): Fn => (args, ctx) => {
  const d = design(args[0]!, args[1], ctx, logY);
  if (fail(d)) return d;
  const cst = optBool(args[3], ctx, true);
  if (fail(cst)) return cst;
  const f = regress(d.Y, d.X, cst);
  const out = (x: number): Scalar => {
    const y = logY ? Math.exp(x) : x;
    return Number.isFinite(y) ? y : Err.num;
  };
  const predict = (xs: number[]): number => f.b + sum(xs.map((x, j) => x * f.m[j]!));
  let nm: number[][];
  if (args[2] === undefined || args[2] === null) {
    if (!d.xShape) return d.yShape.map((row, i) => row.map((_, j) => out(predict([i * row.length + j + 1]))));
    nm = d.xShape;
  } else {
    const m = numMatrix(args[2], ctx);
    if (fail(m)) return m;
    nm = m;
  }
  if (d.single) return nm.map((row) => row.map((x) => out(predict([x]))));
  if (!d.byRow) return nm[0]!.length === d.k ? nm.map((row) => [out(predict(row))]) : Err.ref;
  return nm.length === d.k ? [nm[0]!.map((_, j) => out(predict(nm.map((r) => r[j]!))))] : Err.ref;
};

const simpleReg = (x: number[], y: number[]): { slope: number; icpt: number; sxx: number; syy: number; sxy: number; n: number } | CellError => {
  const n = x.length;
  if (!n) return Err.div0;
  const mx = mean(x), my = mean(y);
  const sxx = sum(x.map((v) => (v - mx) ** 2)), syy = sum(y.map((v) => (v - my) ** 2));
  const sxy = sum(x.map((v, i) => (v - mx) * (y[i]! - my)));
  if (sxx === 0) return Err.div0;
  return { slope: sxy / sxx, icpt: my - (sxy / sxx) * mx, sxx, syy, sxy, n };
};

const correl = (x: number[], y: number[]): Num => {
  const r = simpleReg(x, y);
  if (fail(r)) return r;
  if (r.syy === 0) return Err.div0;
  return r.sxy / Math.sqrt(r.sxx * r.syy);
};

const forecastFn: FunctionDef['fn'] = (args, ctx) => {
  const x = num(args[0]!, ctx);
  if (fail(x)) return x;
  const pr = pairs(args[1]!, args[2]!, ctx);
  if (fail(pr)) return pr;
  const r = simpleReg(pr[1], pr[0]);
  return fail(r) ? r : ok(r.icpt + r.slope * x);
};

// ---------- hypothesis tests ----------

const ttestFn: Fn = (args, ctx) => {
  const tails = num(args[2]!, ctx);
  if (fail(tails)) return tails;
  const type = num(args[3]!, ctx);
  if (fail(type)) return type;
  const tl = Math.trunc(tails), ty = Math.trunc(type);
  if (tl !== 1 && tl !== 2) return Err.num;
  if (ty < 1 || ty > 3) return Err.num;
  let t: number, df: number;
  if (ty === 1) {
    const pr = pairs(args[0]!, args[1]!, ctx);
    if (fail(pr)) return pr;
    const d = pr[0].map((v, i) => v - pr[1][i]!);
    const n = d.length;
    if (n < 2) return Err.div0;
    const s = Math.sqrt(devsq(d) / (n - 1));
    if (s === 0) return Err.div0;
    t = mean(d) / (s / Math.sqrt(n));
    df = n - 1;
  } else {
    const a = data(args[0]!, ctx);
    if (fail(a)) return a;
    const b = data(args[1]!, ctx);
    if (fail(b)) return b;
    const n1 = a.length, n2 = b.length;
    if (n1 < 2 || n2 < 2) return Err.div0;
    const v1 = devsq(a) / (n1 - 1), v2 = devsq(b) / (n2 - 1);
    if (ty === 2) {
      df = n1 + n2 - 2;
      const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
      t = (mean(a) - mean(b)) / Math.sqrt(sp * (1 / n1 + 1 / n2));
    } else {
      const q1 = v1 / n1, q2 = v2 / n2;
      t = (mean(a) - mean(b)) / Math.sqrt(q1 + q2);
      df = (q1 + q2) ** 2 / (q1 * q1 / (n1 - 1) + q2 * q2 / (n2 - 1));
    }
    if (!Number.isFinite(t)) return Err.div0;
  }
  return ok(tl * tCdf(-Math.abs(t), df));
};

const ftestFn: Fn = (args, ctx) => {
  const a = data(args[0]!, ctx);
  if (fail(a)) return a;
  const b = data(args[1]!, ctx);
  if (fail(b)) return b;
  if (a.length < 2 || b.length < 2) return Err.div0;
  const v1 = devsq(a) / (a.length - 1), v2 = devsq(b) / (b.length - 1);
  if (v1 === 0 || v2 === 0) return Err.div0;
  const [cdf, sf] = fCdfPair(v1 / v2, a.length - 1, b.length - 1);
  return ok(Math.min(1, 2 * Math.min(cdf, sf)));
};

const chitestFn: Fn = (args, ctx) => {
  const am = toMatrix(args[0]!, ctx);
  if (fail(am)) return am;
  const em = toMatrix(args[1]!, ctx);
  if (fail(em)) return em;
  const r = am.length, c = am[0]!.length;
  if (em.length !== r || em[0]!.length !== c) return Err.na;
  if (r === 1 && c === 1) return Err.na;
  let x2 = 0;
  for (let i = 0; i < r; i++) for (let j = 0; j < c; j++) {
    const a = am[i]![j]!, e = em[i]![j]!;
    if (isError(a)) return a;
    if (isError(e)) return e;
    if (typeof a !== 'number' || typeof e !== 'number') continue;
    if (e === 0) return Err.div0;
    x2 += (a - e) ** 2 / e;
  }
  const df = r > 1 && c > 1 ? (r - 1) * (c - 1) : r === 1 ? c - 1 : r - 1;
  return ok(chisqSf(x2, df));
};

const ztestFn: Fn = (args, ctx) => {
  const xs = data(args[0]!, ctx);
  if (fail(xs)) return xs;
  if (!xs.length) return Err.na;
  const x = num(args[1]!, ctx);
  if (fail(x)) return x;
  let s: Num;
  if (args[2] === undefined || args[2] === null) s = stdev(xs, true);
  else s = num(args[2], ctx);
  if (fail(s)) return s;
  if (s <= 0) return s === 0 ? Err.div0 : Err.num;
  const z = (mean(xs) - x) / (s / Math.sqrt(xs.length));
  return ok(normCdf(-z));
};

// ---------- distributions ----------

const isCum = (c: number): boolean => c !== 0;

const normDist = (x: number, m: number, s: number, c: number): Num => {
  if (s <= 0) return Err.num;
  const z = (x - m) / s;
  return isCum(c) ? normCdf(z) : normPdf(z) / s;
};
const normInvP = (p: number, m: number, s: number): Num => (p <= 0 || p >= 1 || s <= 0 ? Err.num : m + s * normInv(p));
const lognDist = (x: number, m: number, s: number, c: number): Num => {
  if (x <= 0 || s <= 0) return Err.num;
  const z = (Math.log(x) - m) / s;
  return isCum(c) ? normCdf(z) : normPdf(z) / (x * s);
};
const lognInv = (p: number, m: number, s: number): Num => (p <= 0 || p >= 1 || s <= 0 ? Err.num : Math.exp(m + s * normInv(p)));
const exponDist = (x: number, l: number, c: number): Num => {
  if (x < 0 || l <= 0) return Err.num;
  return isCum(c) ? -Math.expm1(-l * x) : l * Math.exp(-l * x);
};
const weibull = (x: number, a: number, b: number, c: number): Num => {
  if (x < 0 || a <= 0 || b <= 0) return Err.num;
  const u = (x / b) ** a;
  return isCum(c) ? -Math.expm1(-u) : (a / b) * (x / b) ** (a - 1) * Math.exp(-u);
};
const poisson = (xn: number, m: number, c: number): Num => {
  const x = Math.trunc(xn);
  if (x < 0 || m < 0) return Err.num;
  if (isCum(c)) return m === 0 ? 1 : gammaQ(x + 1, m);
  return m === 0 ? (x === 0 ? 1 : 0) : Math.exp(x * Math.log(m) - m - gammaln(x + 1));
};

const binomPmf = (k: number, n: number, p: number): number => Math.exp(lnChoose(n, k) + xlogy(k, p) + xlogy(n - k, 1 - p));
const binomCdf = (k: number, n: number, p: number): number => {
  if (k >= n) return 1;
  if (p === 0) return 1;
  if (p === 1) return 0;
  return betaIncPair(1 - p, p, n - k, k + 1)[0];
};
const binomDist = (sn: number, nn: number, p: number, c: number): Num => {
  const s = Math.trunc(sn), n = Math.trunc(nn);
  if (s < 0 || s > n || p < 0 || p > 1) return Err.num;
  return isCum(c) ? binomCdf(s, n, p) : binomPmf(s, n, p);
};
const binomRange = (nn: number, p: number, s1n: number, s2n: number): Num => {
  const n = Math.trunc(nn), s1 = Math.trunc(s1n), s2 = Number.isNaN(s2n) ? s1 : Math.trunc(s2n);
  if (n < 0 || p < 0 || p > 1 || s1 < 0 || s1 > n || s2 < s1 || s2 > n) return Err.num;
  let t = 0;
  for (let k = s1; k <= s2; k++) t += binomPmf(k, n, p);
  return t;
};
const binomInv = (nn: number, p: number, a: number): Num => {
  const n = Math.trunc(nn);
  if (n < 0 || p < 0 || p > 1 || a < 0 || a > 1) return Err.num;
  let lo = 0, hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (binomCdf(mid, n, p) >= a * (1 - 1e-14)) hi = mid;
    else lo = mid + 1;
  }
  return lo;
};
const negbinom = (fn: number, sn: number, p: number, c: number): Num => {
  const f = Math.trunc(fn), s = Math.trunc(sn);
  if (p < 0 || p > 1 || f < 0 || s < 1) return Err.num;
  if (isCum(c)) return p === 0 ? 0 : p === 1 ? 1 : betaIncPair(p, 1 - p, s, f + 1)[0];
  return Math.exp(lnChoose(f + s - 1, s - 1) + xlogy(s, p) + xlogy(f, 1 - p));
};
const hypgeom = (sn: number, nn: number, Mn: number, Nn: number, c: number): Num => {
  const s = Math.trunc(sn), n = Math.trunc(nn), M = Math.trunc(Mn), N = Math.trunc(Nn);
  if (s < 0 || s > Math.min(n, M) || s < Math.max(0, n - N + M)) return Err.num;
  if (n <= 0 || n > N || M <= 0 || M > N || N <= 0) return Err.num;
  const pmf = (k: number): number => Math.exp(lnChoose(M, k) + lnChoose(N - M, n - k) - lnChoose(N, n));
  if (!isCum(c)) return pmf(s);
  let t = 0;
  for (let k = Math.max(0, n - N + M); k <= s; k++) t += pmf(k);
  return Math.min(1, t);
};

const gammaDist = (x: number, a: number, b: number, c: number): Num => {
  if (x < 0 || a <= 0 || b <= 0) return Err.num;
  return isCum(c) ? gammaP(a, x / b) : gammaPdf(x, a, b);
};
const gammaInvP = (p: number, a: number, b: number): Num => {
  if (p < 0 || p >= 1 || a <= 0 || b <= 0) return Err.num;
  return p === 0 ? 0 : b * gammaInv(p, a);
};

const betaDist = (x: number, a: number, b: number, c: number, lo: number, hi: number): Num => {
  if (a <= 0 || b <= 0 || x < lo || x > hi || lo === hi) return Err.num;
  const z = (x - lo) / (hi - lo);
  return isCum(c) ? betaIncPair(z, (hi - x) / (hi - lo), a, b)[0] : betaPdf(z, a, b) / (hi - lo);
};
const betaInvP = (p: number, a: number, b: number, lo: number, hi: number): Num => {
  if (p <= 0 || p > 1 || a <= 0 || b <= 0 || lo >= hi) return Err.num;
  return lo + (hi - lo) * betaIncInv(p, a, b);
};

const dfOk = (k: number): boolean => k >= 1 && k <= 1e10;
const chisqDist = (x: number, kn: number, c: number): Num => {
  const k = Math.trunc(kn);
  if (x < 0 || !dfOk(k)) return Err.num;
  return isCum(c) ? gammaP(k / 2, x / 2) : gammaPdf(x, k / 2, 2);
};
const chisqRt = (x: number, kn: number): Num => {
  const k = Math.trunc(kn);
  if (x < 0 || !dfOk(k)) return Err.num;
  return chisqSf(x, k);
};
const chisqInv = (p: number, kn: number): Num => {
  const k = Math.trunc(kn);
  if (p < 0 || p >= 1 || !dfOk(k)) return Err.num;
  return p === 0 ? 0 : 2 * gammaInv(p, k / 2);
};
const chisqInvRt = (p: number, kn: number): Num => {
  const k = Math.trunc(kn);
  if (p <= 0 || p > 1 || !dfOk(k)) return Err.num;
  return p === 1 ? 0 : 2 * gammaInvUpper(p, k / 2);
};

const fArgs = (a: number, b: number): [number, number] | CellError => {
  const d1 = Math.trunc(a), d2 = Math.trunc(b);
  return d1 < 1 || d2 < 1 || d1 > 1e10 || d2 > 1e10 ? Err.num : [d1, d2];
};
const fDist = (x: number, a: number, b: number, c: number): Num => {
  const d = fArgs(a, b);
  if (fail(d) || x < 0) return Err.num;
  return isCum(c) ? fCdfPair(x, d[0], d[1])[0] : fPdf(x, d[0], d[1]);
};
const fRt = (x: number, a: number, b: number): Num => {
  const d = fArgs(a, b);
  if (fail(d) || x < 0) return Err.num;
  return fCdfPair(x, d[0], d[1])[1];
};
const fInv = (p: number, a: number, b: number, upper: boolean): Num => {
  const d = fArgs(a, b);
  if (fail(d) || p < 0 || p > 1) return Err.num;
  if (p === (upper ? 1 : 0)) return 0;
  if (p === (upper ? 0 : 1)) return Err.num;
  return solvePositive((x) => (upper ? p - fCdfPair(x, d[0], d[1])[1] : fCdfPair(x, d[0], d[1])[0] - p), 1);
};

const tDf = (v: number): number | CellError => {
  const k = Math.trunc(v);
  return k < 1 || k > 1e10 ? Err.num : k;
};
const tDist = (x: number, v: number, c: number): Num => {
  const k = tDf(v);
  if (fail(k)) return k;
  return isCum(c) ? tCdf(x, k) : tPdf(x, k);
};
const tDist2 = (x: number, v: number): Num => {
  const k = tDf(v);
  if (fail(k) || x < 0) return Err.num;
  return 2 * tCdf(-x, k);
};
const tDistRt = (x: number, v: number): Num => {
  const k = tDf(v);
  return fail(k) ? k : tCdf(-x, k);
};
const tInvP = (p: number, v: number): Num => {
  const k = tDf(v);
  if (fail(k) || p <= 0 || p >= 1) return Err.num;
  return p > 0.5 ? -tInv(1 - p, k) : tInv(p, k);
};
const tInv2 = (p: number, v: number): Num => {
  const k = tDf(v);
  if (fail(k) || p <= 0 || p > 1) return Err.num;
  return p === 1 ? 0 : -tInv(p / 2, k);
};

const confNorm = (a: number, s: number, nn: number): Num => {
  const n = Math.trunc(nn);
  if (a <= 0 || a >= 1 || s <= 0 || n < 1) return Err.num;
  return (-normInv(a / 2) * s) / Math.sqrt(n);
};
const confT = (a: number, s: number, nn: number): Num => {
  const n = Math.trunc(nn);
  if (a <= 0 || a >= 1 || s <= 0 || n < 1) return Err.num;
  if (n === 1) return Err.div0;
  return (-tInv(a / 2, n - 1) * s) / Math.sqrt(n);
};

// ---------- FORECAST.ETS family (AAA exponential smoothing, see ../ets.ts) ----------

const etsSeries = (vArg: Value, tArg: Value, compArg: Value | undefined, aggArg: Value | undefined, ctx: FnContext): EtsSeries | CellError => {
  const comp = optNum(compArg, ctx, 1);
  if (fail(comp)) return comp;
  const agg = optNum(aggArg, ctx, 0);
  if (fail(agg)) return agg;
  return prepareSeries([...cellsOf(vArg, ctx)], [...cellsOf(tArg, ctx)], comp, agg, ctx.date1904);
};

const etsPrepare = (vArg: Value, tArg: Value, seasonArg: Value | undefined, compArg: Value | undefined, aggArg: Value | undefined, ctx: FnContext):
  { series: EtsSeries; model: EtsModel } | CellError => {
  const season = optNum(seasonArg, ctx, 1);
  if (fail(season)) return season;
  const series = etsSeries(vArg, tArg, compArg, aggArg, ctx);
  if (fail(series)) return series;
  const sn = Math.round(season);
  if (sn < 0 || sn > MAX_SEASON) return Err.num;
  const m = sn === 1 ? detectSeason(series.y) : sn;
  if (m > 1 && series.y.length < 2 * m) return Err.num;
  return { series, model: fitEts(series.y, m) };
};

const etsFn: Fn = (args, ctx) => {
  const target = num(args[0]!, ctx);
  if (fail(target)) return target;
  const p = etsPrepare(args[1]!, args[2]!, args[3], args[4], args[5], ctx);
  if (fail(p)) return p;
  const h = stepsAhead(p.series, target);
  if (h < -1e-9) return Err.num;
  return ok(forecastSteps(p.model, p.series.y, Math.max(0, h)));
};

const etsConfFn: Fn = (args, ctx) => {
  const target = num(args[0]!, ctx);
  if (fail(target)) return target;
  const conf = optNum(args[3], ctx, 0.95);
  if (fail(conf)) return conf;
  if (conf <= 0 || conf >= 1) return Err.num;
  const p = etsPrepare(args[1]!, args[2]!, args[4], args[5], args[6], ctx);
  if (fail(p)) return p;
  const h = stepsAhead(p.series, target);
  if (h < -1e-9) return Err.num;
  return ok(confidenceHalfWidth(p.model, p.series.y.length, h, normInv(0.5 + conf / 2)));
};

const etsSeasonFn: Fn = (args, ctx) => {
  const s = etsSeries(args[0]!, args[1]!, args[2], args[3], ctx);
  return fail(s) ? s : detectSeason(s.y);
};

const etsStatFn: Fn = (args, ctx) => {
  const type = num(args[2]!, ctx);
  if (fail(type)) return type;
  const ty = Math.trunc(type);
  if (ty < 1 || ty > 8) return Err.num;
  const p = etsPrepare(args[0]!, args[1]!, args[3], args[4], args[5], ctx);
  if (fail(p)) return p;
  return ok(etsStatistic(p.model, p.series, ty));
};

// ---------- definitions ----------

const avg = (xs: number[]): Num => (xs.length ? mean(xs) : Err.div0);
const maxOf = (xs: number[]): number => xs.reduce((a, b) => Math.max(a, b), xs.length ? -Infinity : 0);
const minOf = (xs: number[]): number => xs.reduce((a, b) => Math.min(a, b), xs.length ? Infinity : 0);
const modeSngl = (xs: number[]): Num => modes(xs)[0] ?? Err.na;
const quartile = (exc: boolean) => (xs: number[], q: number): Num => {
  const k = Math.trunc(q);
  if (exc ? k < 1 || k > 3 : k < 0 || k > 4) return Err.num;
  return exc ? percentileExc(xs, k / 4) : percentileInc(xs, k / 4);
};
const avedev = (xs: number[]): Num => {
  if (!xs.length) return Err.num;
  const m = mean(xs);
  return mean(xs.map((x) => Math.abs(x - m)));
};
const geomean = (xs: number[]): Num => (!xs.length || xs.some((x) => x <= 0) ? Err.num : Math.exp(mean(xs.map(Math.log))));
const harmean = (xs: number[]): Num => (!xs.length || xs.some((x) => x <= 0) ? Err.num : xs.length / sum(xs.map((x) => 1 / x)));
const trimmean = (xs: number[], p: number): Num => {
  if (!xs.length || p < 0 || p >= 1) return Err.num;
  const k = Math.floor((xs.length * p) / 2);
  return mean(sorted(xs).slice(k, xs.length - k));
};
const covar = (sample: boolean) => (x: number[], y: number[]): Num => {
  const n = x.length;
  if (n < (sample ? 2 : 1)) return Err.div0;
  const mx = mean(x), my = mean(y);
  return sum(x.map((v, i) => (v - mx) * (y[i]! - my))) / (sample ? n - 1 : n);
};
const rsq = (x: number[], y: number[]): Num => {
  const r = correl(x, y);
  return fail(r) ? r : r * r;
};
const steyx = (x: number[], y: number[]): Num => {
  const n = x.length;
  if (n < 3) return Err.div0;
  const r = simpleReg(x, y);
  if (fail(r)) return r;
  return Math.sqrt(Math.max(0, r.syy - (r.sxy * r.sxy) / r.sxx) / (n - 2));
};
const slope = (x: number[], y: number[]): Num => {
  const r = simpleReg(x, y);
  return fail(r) ? r : r.slope;
};
const intercept = (x: number[], y: number[]): Num => {
  const r = simpleReg(x, y);
  return fail(r) ? r : r.icpt;
};
const permut = (nn: number, kn: number): Num => {
  const n = Math.trunc(nn), k = Math.trunc(kn);
  if (n < 0 || k < 0 || n < k) return Err.num;
  if (n <= 170) { let p = 1; for (let i = n - k + 1; i <= n; i++) p *= i; return p; }
  return Math.exp(gammaln(n + 1) - gammaln(n - k + 1));
};
const permutA = (nn: number, kn: number): Num => {
  const n = Math.trunc(nn), k = Math.trunc(kn);
  return n < 0 || k < 0 ? Err.num : n ** k;
};
const gammaFn = (x: number): Num => (x <= 0 && Number.isInteger(x) ? Err.num : gamma(x));
const gammaLn = (x: number): Num => (x <= 0 ? Err.num : gammaln(x));

const multi = (name: string, fn: Fn, min: number, max: number, scalar?: readonly number[]): FunctionDef =>
  scalar ? { name, min, max, fn, scalar } : { name, min, max, fn };

export const defs: FunctionDef[] = [
  // central tendency and spread
  aggFn('AVERAGE', avg),
  aggFn('AVERAGEA', avg, A_OPTS),
  aggFn('AVEDEV', avedev),
  aggFn('DEVSQ', (xs) => (xs.length ? devsq(xs) : Err.num)),
  aggFn('GEOMEAN', geomean),
  aggFn('HARMEAN', harmean),
  aggFn('MEDIAN', median),
  aggFn('MAX', maxOf),
  aggFn('MAXA', maxOf, A_OPTS),
  aggFn('MIN', minOf),
  aggFn('MINA', minOf, A_OPTS),
  aggFn('MODE', modeSngl),
  aggFn('MODE.SNGL', modeSngl),
  multi('MODE.MULT', (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    if (fail(xs)) return xs;
    const m = modes(xs);
    return m.length ? m.map((v) => [v]) : Err.na;
  }, 1, 255),
  aggFn('KURT', kurt),
  aggFn('SKEW', skew),
  aggFn('SKEW.P', skewP),
  ...(['STDEV', 'STDEV.S'] as const).map((n) => aggFn(n, (xs) => stdev(xs, true))),
  ...(['STDEVP', 'STDEV.P'] as const).map((n) => aggFn(n, (xs) => stdev(xs, false))),
  aggFn('STDEVA', (xs) => stdev(xs, true), A_OPTS),
  aggFn('STDEVPA', (xs) => stdev(xs, false), A_OPTS),
  ...(['VAR', 'VAR.S'] as const).map((n) => aggFn(n, (xs) => variance(xs, true))),
  ...(['VARP', 'VAR.P'] as const).map((n) => aggFn(n, (xs) => variance(xs, false))),
  aggFn('VARA', (xs) => variance(xs, true), A_OPTS),
  aggFn('VARPA', (xs) => variance(xs, false), A_OPTS),
  multi('COUNT', countFn, 1, 255),
  multi('COUNTA', countaFn, 1, 255),
  multi('COUNTBLANK', countblankFn, 1, 1),

  // order statistics
  kthFn('LARGE', true),
  kthFn('SMALL', false),
  ...(['PERCENTILE', 'PERCENTILE.INC'] as const).map((n) => dataScalarFn(n, 2, 2, percentileInc)),
  dataScalarFn('PERCENTILE.EXC', 2, 2, percentileExc),
  ...(['QUARTILE', 'QUARTILE.INC'] as const).map((n) => dataScalarFn(n, 2, 2, quartile(false))),
  dataScalarFn('QUARTILE.EXC', 2, 2, quartile(true)),
  ...(['PERCENTRANK', 'PERCENTRANK.INC'] as const).map((n) => dataScalarFn(n, 2, 3, (xs, x, s) => percentRank(xs, x!, s!, false), [undefined, 3])),
  dataScalarFn('PERCENTRANK.EXC', 2, 3, (xs, x, s) => percentRank(xs, x!, s!, true), [undefined, 3]),
  rankFn('RANK', false),
  rankFn('RANK.EQ', false),
  rankFn('RANK.AVG', true),
  dataScalarFn('TRIMMEAN', 2, 2, trimmean),
  multi('FREQUENCY', frequencyFn, 2, 2),
  multi('PROB', probFn, 3, 4, [2, 3]),

  // correlation and regression
  pairFn('CORREL', correl),
  pairFn('PEARSON', correl),
  pairFn('RSQ', rsq, true),
  pairFn('COVAR', covar(false)),
  pairFn('COVARIANCE.P', covar(false)),
  pairFn('COVARIANCE.S', covar(true)),
  pairFn('SLOPE', slope, true),
  pairFn('INTERCEPT', intercept, true),
  pairFn('STEYX', steyx, true),
  multi('FORECAST', forecastFn, 3, 3, [0]),
  multi('FORECAST.LINEAR', forecastFn, 3, 3, [0]),
  multi('LINEST', linestFn(false), 1, 4),
  multi('LOGEST', linestFn(true), 1, 4),
  multi('TREND', trendFn(false), 1, 4),
  multi('GROWTH', trendFn(true), 1, 4),
  multi('FORECAST.ETS', etsFn, 3, 6, [0]),
  multi('FORECAST.ETS.CONFINT', etsConfFn, 3, 7, [0]),
  multi('FORECAST.ETS.SEASONALITY', etsSeasonFn, 2, 4),
  multi('FORECAST.ETS.STAT', etsStatFn, 3, 6, [2]),

  // tests
  multi('T.TEST', ttestFn, 4, 4, [2, 3]),
  multi('TTEST', ttestFn, 4, 4, [2, 3]),
  multi('F.TEST', ftestFn, 2, 2),
  multi('FTEST', ftestFn, 2, 2),
  multi('CHISQ.TEST', chitestFn, 2, 2),
  multi('CHITEST', chitestFn, 2, 2),
  multi('Z.TEST', ztestFn, 2, 3, [1, 2]),
  multi('ZTEST', ztestFn, 2, 3, [1, 2]),

  // element-wise
  scalarFn('STANDARDIZE', 3, 3, (x, m, s) => (s! <= 0 ? Err.num : (x! - m!) / s!)),
  scalarFn('FISHER', 1, 1, (x) => (x! <= -1 || x! >= 1 ? Err.num : Math.atanh(x!))),
  scalarFn('FISHERINV', 1, 1, (y) => Math.tanh(y!)),
  scalarFn('GAMMA', 1, 1, (x) => gammaFn(x!)),
  scalarFn('GAMMALN', 1, 1, (x) => gammaLn(x!)),
  scalarFn('GAMMALN.PRECISE', 1, 1, (x) => gammaLn(x!)),
  scalarFn('GAUSS', 1, 1, (z) => normHalf(z!)),
  scalarFn('PHI', 1, 1, (x) => normPdf(x!)),
  scalarFn('PERMUT', 2, 2, (n, k) => permut(n!, k!)),
  scalarFn('PERMUTATIONA', 2, 2, (n, k) => permutA(n!, k!)),
  scalarFn('CONFIDENCE', 3, 3, (a, s, n) => confNorm(a!, s!, n!)),
  scalarFn('CONFIDENCE.NORM', 3, 3, (a, s, n) => confNorm(a!, s!, n!)),
  scalarFn('CONFIDENCE.T', 3, 3, (a, s, n) => confT(a!, s!, n!)),

  scalarFn('NORM.DIST', 4, 4, (x, m, s, c) => normDist(x!, m!, s!, c!)),
  scalarFn('NORMDIST', 4, 4, (x, m, s, c) => normDist(x!, m!, s!, c!)),
  scalarFn('NORM.S.DIST', 2, 2, (z, c) => normDist(z!, 0, 1, c!)),
  scalarFn('NORMSDIST', 1, 1, (z) => normCdf(z!)),
  scalarFn('NORM.INV', 3, 3, (p, m, s) => normInvP(p!, m!, s!)),
  scalarFn('NORMINV', 3, 3, (p, m, s) => normInvP(p!, m!, s!)),
  scalarFn('NORM.S.INV', 1, 1, (p) => normInvP(p!, 0, 1)),
  scalarFn('NORMSINV', 1, 1, (p) => normInvP(p!, 0, 1)),
  scalarFn('LOGNORM.DIST', 4, 4, (x, m, s, c) => lognDist(x!, m!, s!, c!)),
  scalarFn('LOGNORMDIST', 3, 3, (x, m, s) => lognDist(x!, m!, s!, 1)),
  scalarFn('LOGNORM.INV', 3, 3, (p, m, s) => lognInv(p!, m!, s!)),
  scalarFn('LOGINV', 3, 3, (p, m, s) => lognInv(p!, m!, s!)),
  scalarFn('EXPON.DIST', 3, 3, (x, l, c) => exponDist(x!, l!, c!)),
  scalarFn('EXPONDIST', 3, 3, (x, l, c) => exponDist(x!, l!, c!)),
  scalarFn('WEIBULL', 4, 4, (x, a, b, c) => weibull(x!, a!, b!, c!)),
  scalarFn('WEIBULL.DIST', 4, 4, (x, a, b, c) => weibull(x!, a!, b!, c!)),
  scalarFn('POISSON', 3, 3, (x, m, c) => poisson(x!, m!, c!)),
  scalarFn('POISSON.DIST', 3, 3, (x, m, c) => poisson(x!, m!, c!)),
  scalarFn('BINOM.DIST', 4, 4, (s, n, p, c) => binomDist(s!, n!, p!, c!)),
  scalarFn('BINOMDIST', 4, 4, (s, n, p, c) => binomDist(s!, n!, p!, c!)),
  scalarFn('BINOM.DIST.RANGE', 3, 4, (n, p, s1, s2) => binomRange(n!, p!, s1!, s2!), [undefined, undefined, undefined, NaN]),
  scalarFn('BINOM.INV', 3, 3, (n, p, a) => binomInv(n!, p!, a!)),
  scalarFn('CRITBINOM', 3, 3, (n, p, a) => binomInv(n!, p!, a!)),
  scalarFn('NEGBINOM.DIST', 4, 4, (f, s, p, c) => negbinom(f!, s!, p!, c!)),
  scalarFn('NEGBINOMDIST', 3, 3, (f, s, p) => negbinom(f!, s!, p!, 0)),
  scalarFn('HYPGEOM.DIST', 5, 5, (s, n, M, N, c) => hypgeom(s!, n!, M!, N!, c!)),
  scalarFn('HYPGEOMDIST', 4, 4, (s, n, M, N) => hypgeom(s!, n!, M!, N!, 0)),
  scalarFn('GAMMA.DIST', 4, 4, (x, a, b, c) => gammaDist(x!, a!, b!, c!)),
  scalarFn('GAMMADIST', 4, 4, (x, a, b, c) => gammaDist(x!, a!, b!, c!)),
  scalarFn('GAMMA.INV', 3, 3, (p, a, b) => gammaInvP(p!, a!, b!)),
  scalarFn('GAMMAINV', 3, 3, (p, a, b) => gammaInvP(p!, a!, b!)),
  scalarFn('BETA.DIST', 4, 6, (x, a, b, c, lo, hi) => betaDist(x!, a!, b!, c!, lo!, hi!), [undefined, undefined, undefined, undefined, 0, 1]),
  scalarFn('BETADIST', 3, 5, (x, a, b, lo, hi) => betaDist(x!, a!, b!, 1, lo!, hi!), [undefined, undefined, undefined, 0, 1]),
  scalarFn('BETA.INV', 3, 5, (p, a, b, lo, hi) => betaInvP(p!, a!, b!, lo!, hi!), [undefined, undefined, undefined, 0, 1]),
  scalarFn('BETAINV', 3, 5, (p, a, b, lo, hi) => betaInvP(p!, a!, b!, lo!, hi!), [undefined, undefined, undefined, 0, 1]),
  scalarFn('CHISQ.DIST', 3, 3, (x, k, c) => chisqDist(x!, k!, c!)),
  scalarFn('CHISQ.DIST.RT', 2, 2, (x, k) => chisqRt(x!, k!)),
  scalarFn('CHIDIST', 2, 2, (x, k) => chisqRt(x!, k!)),
  scalarFn('CHISQ.INV', 2, 2, (p, k) => chisqInv(p!, k!)),
  scalarFn('CHISQ.INV.RT', 2, 2, (p, k) => chisqInvRt(p!, k!)),
  scalarFn('CHIINV', 2, 2, (p, k) => chisqInvRt(p!, k!)),
  scalarFn('F.DIST', 4, 4, (x, a, b, c) => fDist(x!, a!, b!, c!)),
  scalarFn('F.DIST.RT', 3, 3, (x, a, b) => fRt(x!, a!, b!)),
  scalarFn('FDIST', 3, 3, (x, a, b) => fRt(x!, a!, b!)),
  scalarFn('F.INV', 3, 3, (p, a, b) => fInv(p!, a!, b!, false)),
  scalarFn('F.INV.RT', 3, 3, (p, a, b) => fInv(p!, a!, b!, true)),
  scalarFn('FINV', 3, 3, (p, a, b) => fInv(p!, a!, b!, true)),
  scalarFn('T.DIST', 3, 3, (x, v, c) => tDist(x!, v!, c!)),
  scalarFn('T.DIST.2T', 2, 2, (x, v) => tDist2(x!, v!)),
  scalarFn('T.DIST.RT', 2, 2, (x, v) => tDistRt(x!, v!)),
  scalarFn('TDIST', 3, 3, (x, v, t) => {
    const tl = Math.trunc(t!);
    if (tl !== 1 && tl !== 2) return Err.num;
    return tl === 1 ? (x! < 0 ? Err.num : tDistRt(x!, v!)) : tDist2(x!, v!);
  }),
  scalarFn('T.INV', 2, 2, (p, v) => tInvP(p!, v!)),
  scalarFn('T.INV.2T', 2, 2, (p, v) => tInv2(p!, v!)),
  scalarFn('TINV', 2, 2, (p, v) => tInv2(p!, v!)),
];

