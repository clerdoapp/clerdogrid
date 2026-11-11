// Exponential triple smoothing for FORECAST.ETS, FORECAST.ETS.CONFINT, FORECAST.ETS.SEASONALITY and
// FORECAST.ETS.STAT: the AAA model (additive error, additive trend, additive seasonality).
//
// Sources for the method (public literature, not any product's code):
// - Model and forecast variance: Hyndman, Koehler, Ord & Snyder, "Forecasting with Exponential
//   Smoothing: The State Space Approach" (2008), model ETS(A,A,A), chapter 6 (class 1 variance
//   expression  v_h = sigma^2 [1 + sum_{j=1}^{h-1} c_j^2],  c_j = alpha + beta*j + gamma*[j mod m = 0]).
// - Smoothing weights reported by FORECAST.ETS.STAT are the classic Holt-Winters weights (Winters 1960):
//   alpha (level), beta* (trend) and gamma* (season), each between 0 and 1. In state-space form
//   beta = alpha*beta* and gamma = (1-alpha)*gamma*.
// - Parameters are chosen by minimising the sum of squared one-step errors. For fixed weights the errors
//   are linear in the initial state, so the initial level, trend and seasonal indices are solved exactly by
//   least squares (seasonal indices constrained to sum to zero) and the weights are searched with
//   Nelder-Mead from a grid of starts. Long seasons fall back to the classical-decomposition start of
//   Hyndman et al. section 2.6.
// - Season detection: candidates are the significant peaks of the autocorrelation of the linearly
//   detrended series; the final choice (including "no season") minimises AICc of the fitted models.
// - Data preparation follows the documented arguments of these functions: a constant timeline step
//   (days, or whole months for monthly/yearly date timelines), duplicate timestamps aggregated with one of
//   seven methods, up to 30% missing points completed by interpolation or zeros.

import { CellError, Err, isError, type Scalar } from './types';
import { daysInMonth, fromSerial } from './date';

export const MAX_SEASON = 8760;

export interface EtsSeries {
  y: number[];
  /** Last timeline value. */
  tEnd: number;
  /** Step in timeline units (days), or 0 for a month-based timeline. */
  step: number;
  /** Step in whole months for monthly/yearly date timelines, else 0. */
  months: number;
  date1904: boolean;
  /** What FORECAST.ETS.STAT 8 reports. */
  stepStat: number;
}

const median = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  const k = s.length >> 1;
  return s.length % 2 ? s[k]! : (s[k - 1]! + s[k]!) / 2;
};
const mean = (xs: number[]): number => xs.reduce((a, b) => a + b, 0) / xs.length;

/** Aggregates the values that share one timestamp. 0 and 1 AVERAGE, 2 COUNT, 3 COUNTA, 4 MAX, 5 MEDIAN, 6 MIN, 7 SUM. */
const aggregate = (g: Scalar[], agg: number): number | null => {
  const nums = g.filter((x): x is number => typeof x === 'number');
  if (agg === 2) return nums.length;
  if (agg === 3) return g.filter((x) => x !== null && x !== '').length;
  if (!nums.length) return null;
  switch (agg) {
    case 4: return Math.max(...nums);
    case 5: return median(nums);
    case 6: return Math.min(...nums);
    case 7: return nums.reduce((a, b) => a + b, 0);
    default: return mean(nums);
  }
};

const isInt = (x: number): boolean => Math.abs(x - Math.round(x)) < 1e-9;

/** Turns values and timeline into an evenly spaced series. */
export const prepareSeries = (values: Scalar[], timeline: Scalar[], completion: number, aggregation: number, date1904 = false): EtsSeries | CellError => {
  if (values.length !== timeline.length) return Err.na;
  const comp = Math.trunc(completion), agg = Math.trunc(aggregation);
  if ((comp !== 0 && comp !== 1) || agg < 0 || agg > 7) return Err.num;
  const groups = new Map<number, Scalar[]>();
  for (let i = 0; i < values.length; i++) {
    const t = timeline[i]!, v = values[i]!;
    if (isError(t)) return t;
    if (isError(v)) return v;
    if (typeof t !== 'number') return Err.value;
    const g = groups.get(t);
    if (g) g.push(v); else groups.set(t, [v]);
  }
  const keys = [...groups.keys()].sort((a, b) => a - b);
  if (keys.length < 3) return Err.num;
  const diffs = keys.slice(1).map((k, i) => k - keys[i]!);
  const minDiff = Math.min(...diffs);
  let pos: number[] | null = null;
  let step = 0, months = 0, stepStat = minDiff;
  if (diffs.every((d) => isInt(d / minDiff))) {
    step = minDiff;
    pos = keys.map((k) => Math.round((k - keys[0]!) / step));
  } else if (keys.every((k) => isInt(k))) {
    // Monthly or yearly dates: the same day each month, or every one a month end.
    const parts = keys.map((k) => fromSerial(k, date1904));
    const sameDay = parts.every((p) => p.day === parts[0]!.day);
    const monthEnd = parts.every((p) => p.day === daysInMonth(p.year, p.month));
    if (sameDay || monthEnd) {
      const mi = parts.map((p) => p.year * 12 + p.month - 1);
      const md = mi.slice(1).map((x, i) => x - mi[i]!);
      const gcd = (a: number, b: number): number => (b ? gcd(b, a % b) : a);
      months = md.reduce((a, b) => gcd(a, b));
      if (months > 0) {
        pos = mi.map((x) => (x - mi[0]!) / months);
        stepStat = median(diffs);
      }
    }
  }
  if (!pos || pos.some((p) => !isInt(p))) return Err.num;
  const slots = pos[pos.length - 1]! + 1;
  if (slots > 10_000_000) return Err.num;
  const y: (number | null)[] = new Array<number | null>(slots).fill(null);
  keys.forEach((k, i) => { y[pos![i]!] = aggregate(groups.get(k)!, agg); });
  const missing = y.reduce<number>((a, v) => a + (v === null ? 1 : 0), 0);
  if (missing > 0.3 * slots) return Err.num;
  if (y.every((v) => v === null)) return Err.num;
  for (let i = 0; i < slots; i++) {
    if (y[i] !== null) continue;
    if (comp === 0) { y[i] = 0; continue; }
    let a = i - 1;
    while (a >= 0 && y[a] === null) a--;
    let b = i + 1;
    while (b < slots && y[b] === null) b++;
    const ya = a >= 0 ? y[a]! : null, yb = b < slots ? y[b]! : null;
    y[i] = ya === null ? yb : yb === null ? ya : ya + ((yb - ya) * (i - a)) / (b - a);
  }
  return { y: y as number[], tEnd: keys[keys.length - 1]!, step, months, date1904, stepStat };
};

/** Steps from the end of the timeline to `target` (fractional between points); negative when before it. */
export const stepsAhead = (s: EtsSeries, target: number): number => {
  if (!s.months) return (target - s.tEnd) / s.step;
  const a = fromSerial(s.tEnd, s.date1904), b = fromSerial(Math.floor(target), s.date1904);
  const monthsBetween = (b.year - a.year) * 12 + (b.month - a.month) + (b.day - a.day) / daysInMonth(b.year, b.month) + (target - Math.floor(target)) / daysInMonth(b.year, b.month);
  return monthsBetween / s.months;
};

// ---------- model ----------

export interface EtsModel {
  /** Holt-Winters smoothing weights, each in [0, 1]. */
  alpha: number;
  beta: number;
  gamma: number;
  m: number;
  /** Final level and trend, and seasonal indices for the next m steps (index 0 is one step ahead). */
  level: number;
  trend: number;
  season: number[];
  fitted: number[];
  sse: number;
}

interface Init { l: number; b: number; s: number[] }

/** One pass of the filter. Writes one-step errors into `err` (if given) and returns the final state and SSE. */
const filter = (y: ArrayLike<number>, m: number, a: number, bss: number, gss: number, x0: Init, err?: Float64Array, fitted?: number[]):
  { l: number; b: number; s: Float64Array; sse: number } => {
  let l = x0.l, b = x0.b;
  const seasonal = m > 1;
  const s = new Float64Array(seasonal ? m : 1);
  if (seasonal) for (let j = 0; j < m; j++) s[j] = x0.s[j]!;
  let sse = 0;
  const n = y.length;
  for (let t = 0; t < n; t++) {
    const j = seasonal ? t % m : 0;
    const f = l + b + (seasonal ? s[j]! : 0);
    const e = y[t]! - f;
    if (err) err[t] = e;
    if (fitted) fitted.push(f);
    sse += e * e;
    l = l + b + a * e;
    b = b + bss * e;
    if (seasonal) s[j] = s[j]! + gss * e;
  }
  return { l, b, s, sse };
};

/** Solves the symmetric system A x = r (Gaussian elimination with partial pivoting); null if singular. */
const solve = (A: Float64Array[], r: Float64Array): Float64Array | null => {
  const p = r.length;
  const M = A.map((row, i) => { const x = new Float64Array(p + 1); x.set(row); x[p] = r[i]!; return x; });
  for (let c = 0; c < p; c++) {
    let piv = c;
    for (let i = c + 1; i < p; i++) if (Math.abs(M[i]![c]!) > Math.abs(M[piv]![c]!)) piv = i;
    if (Math.abs(M[piv]![c]!) < 1e-300) return null;
    [M[c], M[piv]] = [M[piv]!, M[c]!];
    for (let i = c + 1; i < p; i++) {
      const f = M[i]![c]! / M[c]![c]!;
      if (f) for (let k = c; k <= p; k++) M[i]![k]! -= f * M[c]![k]!;
    }
  }
  const x = new Float64Array(p);
  for (let i = p - 1; i >= 0; i--) {
    let v = M[i]![p]!;
    for (let k = i + 1; k < p; k++) v -= M[i]![k]! * x[k]!;
    x[i] = v / M[i]![i]!;
  }
  return x;
};

/** Classical decomposition start (Hyndman et al. 2008, section 2.6.1). */
const heuristicInit = (y: number[], m: number): Init => {
  const n = y.length;
  if (m <= 1) {
    const k = Math.min(10, n);
    const xs = Array.from({ length: k }, (_, i) => i + 1);
    const { slope, icpt } = linreg(xs, y.slice(0, k));
    return { l: icpt, b: slope, s: [] };
  }
  const L = Math.min(n, Math.max(2 * m, Math.min(4 * m, n - (n % m))));
  // Centred moving average of order m (2 x m when m is even).
  const trend: (number | null)[] = new Array<number | null>(L).fill(null);
  const half = Math.floor(m / 2);
  for (let t = half; t < L - half; t++) {
    let acc = 0;
    if (m % 2) { for (let k = -half; k <= half; k++) acc += y[t + k]!; trend[t] = acc / m; }
    else {
      if (t + half >= L) continue;
      for (let k = -half; k <= half; k++) acc += (k === -half || k === half ? 0.5 : 1) * y[t + k]!;
      trend[t] = acc / m;
    }
  }
  const sums = new Array<number>(m).fill(0), cnt = new Array<number>(m).fill(0);
  for (let t = 0; t < L; t++) if (trend[t] !== null) { sums[t % m]! += y[t]! - trend[t]!; cnt[t % m]!++; }
  let s = sums.map((v, j) => (cnt[j] ? v / cnt[j]! : 0));
  const avg = mean(s);
  s = s.map((v) => v - avg);
  const k = Math.min(Math.max(10, m), n);
  const xs = Array.from({ length: k }, (_, i) => i + 1);
  const adj = y.slice(0, k).map((v, t) => v - s[t % m]!);
  const { slope, icpt } = linreg(xs, adj);
  return { l: icpt, b: slope, s };
};

const linreg = (x: number[], y: number[]): { slope: number; icpt: number } => {
  const n = x.length, mx = mean(x), my = mean(y);
  let sxy = 0, sxx = 0;
  for (let i = 0; i < n; i++) { sxy += (x[i]! - mx) * (y[i]! - my); sxx += (x[i]! - mx) ** 2; }
  const slope = sxx ? sxy / sxx : 0;
  return { slope, icpt: my - slope * mx };
};

const PROFILE_LIMIT = 4_000_000;

/**
 * For fixed weights, the least-squares initial state. The one-step errors are affine in the initial state,
 * so one filter run per state component gives the design matrix.
 */
const bestInit = (y: number[], m: number, a: number, bss: number, gss: number): Init | null => {
  const n = y.length;
  const seasonal = m > 1;
  const p = 2 + (seasonal ? m - 1 : 0);
  if (n * p * p > PROFILE_LIMIT || n <= p) return null;
  const zero: Init = { l: 0, b: 0, s: new Array<number>(m > 1 ? m : 0).fill(0) };
  const e0 = new Float64Array(n);
  filter(y, m, a, bss, gss, zero, e0);
  const zeros = new Float64Array(n);
  const cols: Float64Array[] = [];
  const unit = (k: number): Init => {
    const x: Init = { l: 0, b: 0, s: new Array<number>(m > 1 ? m : 0).fill(0) };
    if (k === 0) x.l = 1;
    else if (k === 1) x.b = 1;
    else { x.s[k - 2] = 1; x.s[m - 1] = -1; }
    return x;
  };
  for (let k = 0; k < p; k++) {
    const e = new Float64Array(n);
    filter(zeros, m, a, bss, gss, unit(k), e);
    cols.push(e);
  }
  // Minimise |e0 + E x|^2  ->  (E'E) x = -E'e0
  const A = cols.map(() => new Float64Array(p));
  const r = new Float64Array(p);
  for (let i = 0; i < p; i++) {
    const ci = cols[i]!;
    let acc = 0;
    for (let t = 0; t < n; t++) acc += ci[t]! * e0[t]!;
    r[i] = -acc;
    for (let j = i; j < p; j++) {
      const cj = cols[j]!;
      let v = 0;
      for (let t = 0; t < n; t++) v += ci[t]! * cj[t]!;
      A[i]![j] = v; A[j]![i] = v;
    }
  }
  // A tiny ridge keeps nearly collinear cases stable without moving well-posed solutions.
  for (let i = 0; i < p; i++) A[i]![i]! += 1e-12 * (A[i]![i]! || 1);
  const x = solve(A, r);
  if (!x || x.some((v) => !Number.isFinite(v))) return null;
  const s = new Array<number>(m > 1 ? m : 0).fill(0);
  if (seasonal) { let last = 0; for (let j = 0; j < m - 1; j++) { s[j] = x[2 + j]!; last -= x[2 + j]!; } s[m - 1] = last; }
  return { l: x[0]!, b: x[1]!, s };
};

const sigmoid = (u: number): number => 1 / (1 + Math.exp(-u));
const logit = (p: number): number => Math.log(p / (1 - p));
const EPS = 1e-6;
const U_MAX = 14;
const clampU = (u: number): number => Math.max(-U_MAX, Math.min(U_MAX, u));
const toWeights = (u: number[], seasonal: boolean): [number, number, number] => {
  const a = EPS + (1 - 2 * EPS) * sigmoid(clampU(u[0]!));
  const b = EPS + (1 - 2 * EPS) * sigmoid(clampU(u[1]!));
  const g = seasonal ? EPS + (1 - 2 * EPS) * sigmoid(clampU(u[2]!)) : 0;
  return [a, b, g];
};

/** Nelder-Mead minimiser (Lagarias et al. 1998 coefficients). */
const nelderMead = (f: (x: number[]) => number, x0: number[], scale: number, maxIter: number): { x: number[]; v: number } => {
  const d = x0.length;
  let simplex: number[][] = [x0];
  for (let i = 0; i < d; i++) { const p = [...x0]; p[i]! += scale; simplex.push(p); }
  let vals = simplex.map(f);
  for (let it = 0; it < maxIter; it++) {
    const order = vals.map((_, i) => i).sort((a, b) => vals[a]! - vals[b]!);
    simplex = order.map((i) => simplex[i]!);
    vals = order.map((i) => vals[i]!);
    const spread = Math.abs(vals[d]! - vals[0]!);
    const size = Math.max(...simplex.slice(1).map((p) => Math.max(...p.map((v, j) => Math.abs(v - simplex[0]![j]!)))));
    if (spread <= 1e-10 * (Math.abs(vals[0]!) + 1e-30) || size < 1e-6) break;
    const c = new Array<number>(d).fill(0);
    for (let i = 0; i < d; i++) for (let j = 0; j < d; j++) c[j]! += simplex[i]![j]! / d;
    const at = (t: number): number[] => c.map((v, j) => v + t * (simplex[d]![j]! - v));
    const xr = at(-1), fr = f(xr);
    if (fr < vals[0]!) {
      const xe = at(-2), fe = f(xe);
      if (fe < fr) { simplex[d] = xe; vals[d] = fe; } else { simplex[d] = xr; vals[d] = fr; }
    } else if (fr < vals[d - 1]!) { simplex[d] = xr; vals[d] = fr; }
    else {
      const outside = fr < vals[d]!;
      const xc = at(outside ? -0.5 : 0.5), fc = f(xc);
      if (fc < (outside ? fr : vals[d]!)) { simplex[d] = xc; vals[d] = fc; }
      else for (let i = 1; i <= d; i++) { simplex[i] = simplex[i]!.map((v, j) => simplex[0]![j]! + 0.5 * (v - simplex[0]![j]!)); vals[i] = f(simplex[i]!); }
    }
  }
  let best = 0;
  for (let i = 1; i <= d; i++) if (vals[i]! < vals[best]!) best = i;
  return { x: simplex[best]!, v: vals[best]! };
};

const fitCache = new Map<string, EtsModel>();
const seasonCache = new Map<string, number>();
const remember = <T>(cache: Map<string, T>, key: string, value: T): T => {
  if (cache.size >= 32) cache.delete(cache.keys().next().value!);
  cache.set(key, value);
  return value;
};

/** Fits the AAA model with season length m (m <= 1 means no season). Results are cached by series. */
export const fitEts = (y: number[], mIn: number): EtsModel => {
  const key = `${mIn > 1 ? mIn : 1}|${y.join(',')}`;
  return fitCache.get(key) ?? remember(fitCache, key, fitUncached(y, mIn));
};

const fitUncached = (y: number[], mIn: number): EtsModel => {
  const m = mIn > 1 ? mIn : 1;
  const seasonal = m > 1;
  const fixedInit = heuristicInit(y, m);
  const evaluate = (a: number, bh: number, gh: number): { sse: number; init: Init } => {
    const bss = a * bh, gss = (1 - a) * gh;
    const init = bestInit(y, m, a, bss, gss) ?? fixedInit;
    return { sse: filter(y, m, a, bss, gss, init).sse, init };
  };
  const cost = (u: number[]): number => { const [a, b, g] = toWeights(u, seasonal); const r = evaluate(a, b, g).sse; return Number.isFinite(r) ? r : Number.MAX_VALUE; };
  const grid: number[][] = [];
  for (const a of [0.05, 0.2, 0.4, 0.7, 0.95]) for (const b of [0.01, 0.1, 0.4]) for (const g of seasonal ? [0.01, 0.1, 0.3, 0.7] : [0]) grid.push([a, b, g]);
  const big = y.length * (seasonal ? m : 1) > 200_000;
  const starts = grid
    .map((w) => ({ u: [logit(w[0]!), logit(w[1]!), ...(seasonal ? [logit(w[2]!)] : [])], v: 0 }))
    .map((s) => ({ ...s, v: cost(s.u) }))
    .sort((p, q) => p.v - q.v)
    .slice(0, big ? 1 : 3);
  let best = { x: starts[0]!.u, v: starts[0]!.v };
  for (const s of starts) {
    let r = nelderMead(cost, s.u.map(clampU), 0.8, big ? 120 : 300);
    // Restart from the result until it stops improving, which guards against early collapse of the simplex.
    for (let k = 0; k < 3; k++) {
      const r2 = nelderMead(cost, r.x.map(clampU), 0.3, big ? 60 : 150);
      if (r2.v >= r.v - 1e-12 * Math.abs(r.v)) break;
      r = r2;
    }
    if (r.v < best.v) best = r;
  }
  const [a, bh, gh] = toWeights(best.x, seasonal);
  const { init } = evaluate(a, bh, gh);
  const fitted: number[] = [];
  const fin = filter(y, m, a, a * bh, (1 - a) * gh, init, undefined, fitted);
  const n = y.length;
  const season: number[] = [];
  if (seasonal) for (let h = 0; h < m; h++) season.push(fin.s[(n + h) % m]!);
  return { alpha: a, beta: bh, gamma: seasonal ? gh : 0, m, level: fin.l, trend: fin.b, season, fitted, sse: fin.sse };
};

/** Sum of squared one-step errors for given Holt-Winters weights, with the best initial state (for checks). */
export const etsSse = (y: number[], m: number, alpha: number, beta: number, gamma: number): number => {
  const bss = alpha * beta, gss = (1 - alpha) * gamma;
  const init = bestInit(y, m, alpha, bss, gss) ?? heuristicInit(y, m);
  return filter(y, m, alpha, bss, gss, init).sse;
};

/** Point forecast h whole steps after the last observation (h >= 1). */
export const forecastAt = (md: EtsModel, h: number): number =>
  md.level + h * md.trend + (md.m > 1 ? md.season[(h - 1) % md.m]! : 0);

/** Forecast for a possibly fractional number of steps ahead; 0 gives the last observed value. */
export const forecastSteps = (md: EtsModel, y: number[], h: number): number => {
  const at = (k: number): number => (k === 0 ? y[y.length - 1]! : forecastAt(md, k));
  const lo = Math.floor(h), frac = h - lo;
  return frac < 1e-12 ? at(lo) : at(lo) + frac * (at(lo + 1) - at(lo));
};

/** Forecast variance h steps ahead, in units of sigma^2. */
const varianceFactor = (md: EtsModel, h: number): number => {
  const bss = md.alpha * md.beta, gss = (1 - md.alpha) * md.gamma;
  let v = 1;
  for (let j = 1; j < h; j++) {
    const c = md.alpha + bss * j + (md.m > 1 && j % md.m === 0 ? gss : 0);
    v += c * c;
  }
  return v;
};

/** Half-width of the prediction interval h steps ahead for the given z value. */
export const confidenceHalfWidth = (md: EtsModel, n: number, h: number, z: number): number => {
  const sigma2 = md.sse / n;
  const hh = Math.max(1, h);
  const lo = Math.floor(hh), frac = hh - lo;
  const v = frac < 1e-12 ? varianceFactor(md, lo) : varianceFactor(md, lo) + frac * (varianceFactor(md, lo + 1) - varianceFactor(md, lo));
  return z * Math.sqrt(sigma2 * v);
};

const aicc = (md: EtsModel, n: number, scaleFloor: number): number => {
  const k = 4 + (md.m > 1 ? md.m : 0);
  const s2 = Math.max(md.sse / n, scaleFloor);
  const denom = n - k - 1;
  return n * Math.log(s2) + 2 * k + (denom > 0 ? (2 * k * (k + 1)) / denom : 1e9);
};

/** Detrended autocorrelation peaks, strongest first. */
const seasonCandidates = (y: number[]): number[] => {
  const n = y.length;
  const xs = y.map((_, i) => i);
  const { slope, icpt } = linreg(xs, y);
  const e = y.map((v, i) => v - icpt - slope * i);
  let denom = 0;
  for (const v of e) denom += v * v;
  if (denom <= 1e-24 * (1 + Math.abs(mean(y)))) return [];
  const maxLag = Math.min(Math.floor(n / 2), MAX_SEASON);
  const ac = new Float64Array(maxLag + 2);
  for (let l = 1; l <= Math.min(maxLag + 1, n - 1); l++) {
    let s = 0;
    for (let i = l; i < n; i++) s += e[i]! * e[i - l]!;
    ac[l] = s / denom;
  }
  const band = 1.96 / Math.sqrt(n);
  const peaks: { l: number; v: number }[] = [];
  for (let l = 2; l <= maxLag; l++) {
    if (ac[l]! > band && ac[l]! > ac[l - 1]! && ac[l]! >= ac[l + 1]!) peaks.push({ l, v: ac[l]! });
  }
  peaks.sort((a, b) => b.v - a.v || a.l - b.l);
  return peaks.slice(0, 4).map((p) => p.l);
};

/** Season length chosen for the series: 1 when there is no repeating pattern. */
export const detectSeason = (y: number[]): number => {
  const key = y.join(',');
  return seasonCache.get(key) ?? remember(seasonCache, key, detectUncached(y));
};

const detectUncached = (y: number[]): number => {
  const n = y.length;
  const cands = seasonCandidates(y).filter((m) => n >= 2 * m);
  if (!cands.length) return 1;
  const varY = y.reduce((a, v) => a + (v - mean(y)) ** 2, 0) / n;
  const floor = Math.max(1e-10 * varY, 1e-300);
  let best = 1, bestScore = aicc(fitEts(y, 1), n, floor);
  for (const m of [...cands].sort((a, b) => a - b)) {
    const sc = aicc(fitEts(y, m), n, floor);
    if (sc < bestScore - 1e-9) { best = m; bestScore = sc; }
  }
  return best;
};

/** The eight FORECAST.ETS.STAT statistics. */
export const etsStatistic = (md: EtsModel, s: EtsSeries, type: number): number => {
  const y = s.y, f = md.fitted, n = y.length;
  const err = y.map((v, i) => v - f[i]!);
  const mae = mean(err.map(Math.abs));
  switch (type) {
    case 1: return md.alpha;
    case 2: return md.beta;
    case 3: return md.gamma;
    case 4: {
      const naive = mean(y.slice(1).map((v, i) => Math.abs(v - y[i]!)));
      return naive === 0 ? (mae === 0 ? 0 : Infinity) : mae / naive;
    }
    case 5: return mean(y.map((v, i) => { const d = Math.abs(v) + Math.abs(f[i]!); return d === 0 ? 0 : (2 * Math.abs(err[i]!)) / d; }));
    case 6: return mae;
    case 7: return Math.sqrt(md.sse / n);
    default: return s.stepStat;
  }
};
