// Least-squares trendlines and error-bar amounts for charts.
import type { ErrorBars, Trendline } from '../../model/workbook';
import { N, T, Tr } from '../i18n';

export interface Fit {
  kind: Trendline['kind'];
  /** y for a given x (NaN outside the curve's domain). */
  at: (x: number) => number;
  /** Polynomial / linear coefficients, constant first. For exp and power: [c, b]. For log: [b, c] meaning c·ln(x) + b. */
  coef: number[];
  r2: number;
  /** Equation as runs of text; `sup` runs are exponents. */
  equation: { t: string; sup?: boolean }[];
  /** Moving average points, already computed. */
  points?: [number, number][];
}

type XY = [number, number];

/** Solves A·x = b in place by Gaussian elimination with partial pivoting. Returns null when singular. */
export const solve = (A: number[][], b: number[]): number[] | null => {
  const n = b.length;
  const M = A.map((row, i) => [...row, b[i]!]);
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(M[r]![col]!) > Math.abs(M[piv]![col]!)) piv = r;
    if (Math.abs(M[piv]![col]!) < 1e-300) return null;
    [M[col], M[piv]] = [M[piv]!, M[col]!];
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const f = M[r]![col]! / M[col]![col]!;
      if (!f) continue;
      for (let k = col; k <= n; k++) M[r]![k]! -= f * M[col]![k]!;
    }
  }
  return M.map((row, i) => row[n]! / row[i]!);
};

/**
 * Least squares polynomial of `order` through the points. With `intercept` the constant term is fixed.
 * x is centred and scaled first so high orders stay well conditioned.
 */
export const polyFit = (pts: XY[], order: number, intercept?: number): number[] | null => {
  const n = pts.length;
  if (!n) return null;
  if (intercept !== undefined) {
    // y - b0 = Σ a_k x^k (k ≥ 1): no centring, since the constant must stay at x = 0.
    const ys = pts.map(([, y]) => y - intercept);
    const A: number[][] = [], b: number[] = [];
    for (let i = 1; i <= order; i++) {
      A.push([]);
      for (let j = 1; j <= order; j++) A[i - 1]!.push(pts.reduce((s, [x]) => s + x ** (i + j), 0));
      b.push(pts.reduce((s, [x], k) => s + x ** i * ys[k]!, 0));
    }
    const sol = solve(A, b);
    return sol ? [intercept, ...sol] : null;
  }
  const mx = pts.reduce((s, [x]) => s + x, 0) / n;
  const sx = Math.max(1e-12, Math.max(...pts.map(([x]) => Math.abs(x - mx))));
  const u = pts.map(([x]) => (x - mx) / sx);
  const m = order + 1;
  const A = Array.from({ length: m }, (_, i) => Array.from({ length: m }, (_, j) => u.reduce((s, x) => s + x ** (i + j), 0)));
  const b = Array.from({ length: m }, (_, i) => u.reduce((s, x, k) => s + x ** i * pts[k]![1], 0));
  const c = solve(A, b);
  if (!c) return null;
  // Expand Σ c_i ((x - mx)/sx)^i back into powers of x.
  const out = new Array<number>(m).fill(0);
  for (let i = 0; i < m; i++) {
    const ci = c[i]! / sx ** i;
    for (let k = 0; k <= i; k++) out[k]! += ci * binom(i, k) * (-mx) ** (i - k);
  }
  return out;
};

const binom = (n: number, k: number): number => {
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return r;
};

const evalPoly = (c: number[], x: number): number => c.reduceRight((acc, a) => acc * x + a, 0);

const rSquared = (pts: XY[], f: (x: number) => number, centred: boolean): number => {
  const n = pts.length;
  if (n < 2) return 1;
  const mean = centred ? pts.reduce((s, [, y]) => s + y, 0) / n : 0;
  let res = 0, tot = 0;
  for (const [x, y] of pts) { res += (y - f(x)) ** 2; tot += (y - mean) ** 2; }
  return tot ? 1 - res / tot : 1;
};

/** Number text for an equation: about 4 significant figures, no trailing zeros. */
export const coefText = (v: number): string => {
  if (!Number.isFinite(v)) return '0';
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e6 || a < 1e-4) return v.toExponential(3).replace(/\.?0+e/, 'E').replace('E+', 'E');
  return String(+v.toPrecision(5));
};

const signed = (v: number, first: boolean): string => (first ? coefText(v) : v < 0 ? ` - ${coefText(-v)}` : ` + ${coefText(v)}`);

export const fitTrend = (raw: XY[], t: Trendline): Fit | null => {
  const pts = raw.filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
  if (t.kind === 'movingAvg') {
    const p = Math.max(2, Math.min(255, Math.round(t.period ?? 2)));
    if (pts.length < p) return null;
    const out: XY[] = [];
    for (let i = p - 1; i < pts.length; i++) {
      let s = 0;
      for (let k = i - p + 1; k <= i; k++) s += pts[k]![1];
      out.push([pts[i]![0], s / p]);
    }
    return { kind: 'movingAvg', at: () => NaN, coef: [], r2: NaN, equation: [], points: out };
  }
  if (pts.length < 2) return null;
  switch (t.kind) {
    case 'linear': case 'poly': {
      const order = t.kind === 'linear' ? 1 : Math.max(2, Math.min(6, Math.round(t.order ?? 2)));
      if (pts.length <= order && t.intercept === undefined) return null;
      const c = polyFit(pts, order, t.intercept);
      if (!c) return null;
      const f = (x: number) => evalPoly(c, x);
      const eq: Fit['equation'] = [{ t: 'y = ' }];
      let first = true;
      for (let k = order; k >= 0; k--) {
        const a = c[k]!;
        if (k > 0 && Math.abs(a) < 1e-15) continue;
        eq.push({ t: `${signed(a, first)}${k ? 'x' : ''}` });
        if (k > 1) eq.push({ t: String(k), sup: true });
        first = false;
      }
      return { kind: t.kind, at: f, coef: c, r2: rSquared(pts, f, t.intercept === undefined), equation: eq };
    }
    case 'exp': {
      const p = pts.filter(([, y]) => y > 0);
      if (p.length < 2) return null;
      const lin = p.map(([x, y]): XY => [x, Math.log(y)]);
      const fixed = t.intercept !== undefined && t.intercept > 0 ? Math.log(t.intercept) : undefined;
      const c = polyFit(lin, 1, fixed);
      if (!c) return null;
      const a = Math.exp(c[0]!), b = c[1]!;
      return {
        kind: 'exp', at: (x) => a * Math.exp(b * x), coef: [a, b], r2: rSquared(lin, (x) => c[0]! + b * x, fixed === undefined),
        equation: [{ t: `y = ${coefText(a)}e` }, { t: `${coefText(b)}x`, sup: true }],
      };
    }
    case 'log': {
      const p = pts.filter(([x]) => x > 0);
      if (p.length < 2) return null;
      const lin = p.map(([x, y]): XY => [Math.log(x), y]);
      const c = polyFit(lin, 1);
      if (!c) return null;
      return {
        kind: 'log', at: (x) => (x > 0 ? c[1]! * Math.log(x) + c[0]! : NaN), coef: [c[0]!, c[1]!], r2: rSquared(lin, (u) => c[0]! + c[1]! * u, true),
        equation: [{ t: `y = ${coefText(c[1]!)}ln(x)${signed(c[0]!, false)}` }],
      };
    }
    case 'power': {
      const p = pts.filter(([x, y]) => x > 0 && y > 0);
      if (p.length < 2) return null;
      const lin = p.map(([x, y]): XY => [Math.log(x), Math.log(y)]);
      const c = polyFit(lin, 1);
      if (!c) return null;
      const a = Math.exp(c[0]!), b = c[1]!;
      return {
        kind: 'power', at: (x) => (x > 0 ? a * x ** b : NaN), coef: [a, b], r2: rSquared(lin, (u) => c[0]! + b * u, true),
        equation: [{ t: `y = ${coefText(a)}x` }, { t: coefText(b), sup: true }],
      };
    }
    default: return null;
  }
};

/** Points along a fitted curve between x0 and x1, for drawing. */
export const sampleFit = (f: Fit, x0: number, x1: number, n = 80): XY[] => {
  if (f.points) return f.points;
  const out: XY[] = [];
  if (f.kind === 'linear') {
    for (const x of [x0, x1]) { const y = f.at(x); if (Number.isFinite(y)) out.push([x, y]); }
    return out;
  }
  for (let i = 0; i <= n; i++) {
    const x = x0 + ((x1 - x0) * i) / n;
    const y = f.at(x);
    if (Number.isFinite(y)) out.push([x, y]);
  }
  return out;
};

export const trendName = (t: Trendline, series: string): string => {
  if (t.name) return t.name;
  const k = t.kind === 'movingAvg' ? T('{n} per. moving avg.', { n: t.period ?? 2 }) : Tr({ linear: N('Linear'), exp: N('Exponential'), log: N('Logarithmic'), poly: N('Polynomial'), power: N('Power') }[t.kind]);
  return `${k} (${series})`;
};

// ---------- Error bars ----------

export const mean = (v: number[]): number => (v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0);
/** Sample standard deviation. */
export const stdev = (v: number[]): number => {
  if (v.length < 2) return 0;
  const m = mean(v);
  return Math.sqrt(v.reduce((s, x) => s + (x - m) ** 2, 0) / (v.length - 1));
};

/**
 * Low and high ends of each point's error bar. `values` are the plotted values (null for blanks);
 * `custom` holds resolved custom amounts. Standard deviation bars are centred on the series mean.
 */
export const errorRanges = (values: (number | null)[], e: ErrorBars, custom?: { plus?: (number | null)[]; minus?: (number | null)[] }): ([number, number] | null)[] => {
  const nums = values.filter((v): v is number => v !== null && Number.isFinite(v));
  const amount = e.value ?? (e.kind === 'percent' ? 5 : 1);
  const type = e.type ?? 'both';
  const sd = stdev(nums), m = mean(nums);
  return values.map((v, i) => {
    if (v === null || !Number.isFinite(v)) return null;
    let center = v, plus = 0, minus = 0;
    switch (e.kind) {
      case 'fixed': plus = minus = Math.abs(amount); break;
      case 'percent': plus = minus = Math.abs(v * amount) / 100; break;
      case 'stdDev': center = m; plus = minus = Math.abs(amount) * sd; break;
      case 'stdErr': plus = minus = nums.length ? sd / Math.sqrt(nums.length) : 0; break;
      case 'custom': {
        // One amount applies to every point; otherwise each point has its own (blank means none).
        const pick = (a?: (number | null)[]) => Math.abs((a?.length === 1 ? a[0] : a?.[i]) ?? 0);
        plus = pick(custom?.plus); minus = pick(custom?.minus);
        break;
      }
    }
    return [type === 'plus' ? center : center - minus, type === 'minus' ? center : center + plus];
  });
};

/** Parses an array constant like {1,2;3} into numbers; returns null when the text isn't one. */
export const parseNumberList = (s: string): number[] | null => {
  const m = /^=?\s*\{([^{}]*)\}\s*$/.exec(s);
  if (!m) return null;
  const parts = m[1]!.split(/[,;]/).map((x) => Number(x.trim()));
  return parts.every((x) => Number.isFinite(x)) ? parts : null;
};
