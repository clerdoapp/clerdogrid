import { CellError, Err, isError, type FnContext, type FunctionDef, type Value } from '../types';
import { cellsOf, collectNumbers, fromRange, num, optBool, toNumber } from '../coerce';
import { addMonths, dateArg, days360EU, days360US, isMonthEnd, yearFrac } from './datetime';

type Num = number | CellError;

/** Finite results only: NaN and ±Infinity become #NUM!. */
const fin = (x: number): Num => (Number.isFinite(x) ? x : Err.num);

/**
 * Reads numeric arguments. `spec[i]` is undefined for a required argument, or the default for an
 * optional one ('date' marks a required date, which is truncated and range-checked).
 */
const nums = (args: Value[], ctx: FnContext, spec: Array<number | 'date' | undefined>): number[] | CellError => {
  const out: number[] = [];
  for (let i = 0; i < spec.length; i++) {
    const s = spec[i], a = args[i];
    let n: Num;
    if (s === 'date') n = dateArg(a!, ctx);
    else if (s !== undefined && (a === undefined || a === null)) n = s;
    else n = num(a!, ctx);
    if (isError(n)) return n;
    out.push(n);
  }
  return out;
};

type Impl = (n: number[], ctx: FnContext) => Num;
const def = (name: string, spec: Array<number | 'date' | undefined>, impl: Impl): FunctionDef => {
  const min = spec.filter((s) => s === undefined || s === 'date').length;
  return {
    name, min, max: spec.length, scalar: true,
    fn: (args, ctx) => { const n = nums(args, ctx, spec); if (isError(n)) return n; const r = impl(n, ctx); return isError(r) ? r : fin(r); },
  };
};

// ---- Root finding: Newton from the guess, then a bracketing search with bisection/secant fallback.

const TOL = 1e-10;
export const solve = (f: (x: number) => number, guess: number, lo = -1 + 1e-10, hi = 1e10): Num => {
  const d = (x: number): number => { const h = Math.max(1e-7, Math.abs(x) * 1e-7); return (f(x + h) - f(x - h)) / (2 * h); };
  let x = guess;
  const scale = Math.max(1, Math.abs(f(guess)) || 1);
  for (let i = 0; i < 100; i++) {
    const fx = f(x);
    if (!Number.isFinite(fx)) break;
    if (Math.abs(fx) < 1e-12) return x;
    const dx = d(x);
    if (!dx || !Number.isFinite(dx)) break;
    let nx = x - fx / dx;
    if (nx <= lo) nx = (x + lo) / 2;
    if (Math.abs(nx - x) < TOL * Math.max(1, Math.abs(nx))) {
      if (Math.abs(f(nx)) < 1e-6 * scale) return nx;
      break; // stalled against the domain edge, not at a root
    }
    x = nx;
  }
  // Bracket a sign change outward from the guess.
  const g = Math.min(Math.max(guess, lo), hi);
  const points: number[] = [g];
  for (let step = 0.01; step < 1e6; step *= 1.6) points.push(g + step, g - step);
  let a = NaN, b = NaN;
  let prev = g, fprev = f(g);
  for (const p of points.slice(1).filter((p) => p > lo && p < hi).sort((u, v) => Math.abs(u - g) - Math.abs(v - g))) {
    const fp = f(p);
    if (Number.isFinite(fp) && Number.isFinite(fprev) && Math.sign(fp) !== Math.sign(fprev)) { a = Math.min(p, prev); b = Math.max(p, prev); break; }
    if (Math.sign(p - g) === Math.sign(prev - g) || prev === g) { prev = p; fprev = fp; }
  }
  if (Number.isNaN(a)) {
    // Fall back to a scan of the whole domain.
    let px = lo + 1e-9, pf = f(px);
    for (let t = -0.99; t < 1e3; t = t < 1 ? t + 0.01 : t * 1.1) {
      const ft = f(t);
      if (Number.isFinite(ft) && Number.isFinite(pf) && Math.sign(ft) !== Math.sign(pf)) { a = px; b = t; break; }
      px = t; pf = ft;
    }
    if (Number.isNaN(a)) return Err.num;
  }
  let fa = f(a);
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2, fm = f(m);
    if (Math.abs(fm) < 1e-12 || (b - a) / 2 < TOL * Math.max(1, Math.abs(m))) return m;
    if (Math.sign(fm) === Math.sign(fa)) { a = m; fa = fm; } else b = m;
  }
  return Err.num;
};

// ---- Time value of money

const fvOf = (r: number, n: number, pmt: number, pv: number, t: number): number =>
  r === 0 ? -(pv + pmt * n) : -(pv * (1 + r) ** n + (pmt * (1 + r * t) * ((1 + r) ** n - 1)) / r);
const pvOf = (r: number, n: number, pmt: number, fv: number, t: number): number =>
  r === 0 ? -(fv + pmt * n) : -(fv + (pmt * (1 + r * t) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n;
const pmtOf = (r: number, n: number, pv: number, fv: number, t: number): number =>
  r === 0 ? -(pv + fv) / n : -(r * (fv + pv * (1 + r) ** n)) / ((1 + r * t) * ((1 + r) ** n - 1));
const ipmtOf = (r: number, per: number, n: number, pv: number, fv: number, t: number): number => {
  const pmt = pmtOf(r, n, pv, fv, t);
  const ip = per === 1 ? (t ? 0 : -pv) : t ? fvOf(r, per - 2, pmt, pv, 1) - pmt : fvOf(r, per - 1, pmt, pv, 0);
  return ip * r;
};
const typeOf = (t: number): number => (t !== 0 ? 1 : 0);

// ---- Day counts and coupon schedules (bases 0–4)

const basisOk = (b: number): boolean => b >= 0 && b <= 4;
const freqOk = (f: number): boolean => f === 1 || f === 2 || f === 4;
const dayCount = (a: number, b: number, basis: number, d1904: boolean): number =>
  basis === 0 ? days360US(a, b, d1904) : basis === 4 ? days360EU(a, b, d1904) : b - a;
const periodDays = (p: number, n: number, basis: number, freq: number): number =>
  basis === 1 ? n - p : basis === 3 ? 365 / freq : 360 / freq;

interface Coupon { pcd: number; ncd: number; num: number }
/** Coupon dates around settlement, stepping back from maturity (month ends stay month ends). */
const coupons = (settle: number, mat: number, freq: number, d1904: boolean): Coupon => {
  const months = 12 / freq, eom = isMonthEnd(mat, d1904);
  let k = 1, pcd = addMonths(mat, -months, d1904, eom), ncd = mat;
  while (pcd > settle) { ncd = pcd; k++; pcd = addMonths(mat, -k * months, d1904, eom); }
  return { pcd, ncd, num: k };
};

interface Bond { A: number; E: number; DSC: number; N: number }
const bond = (settle: number, mat: number, freq: number, basis: number, d1904: boolean): Bond => {
  const c = coupons(settle, mat, freq, d1904);
  const A = dayCount(c.pcd, settle, basis, d1904);
  const E = periodDays(c.pcd, c.ncd, basis, freq);
  const DSC = basis === 0 || basis === 4 ? E - A : c.ncd - settle;
  return { A, E, DSC, N: c.num };
};

/** Checks shared by the security functions: settlement < maturity, frequency 1/2/4, basis 0–4. Dates truncate. */
const secCheck = (settle: number, mat: number, freq: number, basis: number): CellError | null =>
  settle >= mat || !freqOk(freq) || !basisOk(basis) ? Err.num : null;

const priceOf = (b: Bond, rate: number, yld: number, red: number, freq: number): number => {
  const c = (100 * rate) / freq, base = 1 + yld / freq, x = b.DSC / b.E;
  if (b.N === 1) return (red + c) / (1 + x * (yld / freq)) - (c * b.A) / b.E;
  let p = red / base ** (b.N - 1 + x);
  for (let k = 1; k <= b.N; k++) p += c / base ** (k - 1 + x);
  return p - (c * b.A) / b.E;
};

const durationOf = (settle: number, mat: number, cpn: number, yld: number, freq: number, basis: number, d1904: boolean): number => {
  const b = bond(settle, mat, freq, basis, d1904), x = b.DSC / b.E;
  const c = (cpn * 100) / freq, y = 1 + yld / freq;
  let dur = 0, p = 0;
  for (let k = 1; k <= b.N; k++) {
    const t = k - 1 + x, v = (c + (k === b.N ? 100 : 0)) / y ** t;
    dur += t * v; p += v;
  }
  return dur / p / freq;
};

// ---- Odd-period securities: quasi-coupon periods counted from an anchor coupon date.

interface Quasi { start: number; end: number; len: number }
const quasiPeriods = (anchor: number, from: number, to: number, freq: number, basis: number, d1904: boolean, backward: boolean): Quasi[] => {
  const months = 12 / freq, eom = isMonthEnd(anchor, d1904), out: Quasi[] = [];
  for (let k = 0; ; k++) {
    const a = addMonths(anchor, (backward ? -(k + 1) : k) * months, d1904, eom);
    const b = addMonths(anchor, (backward ? -k : k + 1) * months, d1904, eom);
    out.push({ start: a, end: b, len: periodDays(a, b, basis, freq) });
    if (backward ? a <= from : b >= to) break;
  }
  return backward ? out.reverse() : out;
};
/** Sum over periods of (days of [lo, hi] falling in the period) / period length. */
const fraction = (ps: Quasi[], lo: number, hi: number, basis: number, d1904: boolean): number =>
  ps.reduce((s, p) => {
    const a = Math.max(p.start, lo), b = Math.min(p.end, hi);
    return b > a ? s + dayCount(a, b, basis, d1904) / p.len : s;
  }, 0);

const oddFirstPrice = (s: number, m: number, iss: number, fc: number, rate: number, yld: number, red: number, freq: number, basis: number, d1904: boolean): number => {
  const ps = quasiPeriods(fc, iss, fc, freq, basis, d1904, true);
  const c = (100 * rate) / freq, base = 1 + yld / freq;
  const firstCoupon = c * fraction(ps, iss, fc, basis, d1904);
  const accrued = c * fraction(ps, iss, s, basis, d1904);
  const t1 = fraction(ps, s, fc, basis, d1904); // periods from settlement to the first coupon
  const n = coupons(fc, m, freq, d1904).num;
  let p = firstCoupon / base ** t1 + red / base ** (t1 + n);
  for (let k = 1; k <= n; k++) p += c / base ** (t1 + k);
  return p - accrued;
};

interface OddLast { dc: number; a: number; dsc: number }
const oddLast = (s: number, m: number, last: number, freq: number, basis: number, d1904: boolean): OddLast => {
  const ps = quasiPeriods(last, last, m, freq, basis, d1904, false);
  return { dc: fraction(ps, last, m, basis, d1904), a: fraction(ps, last, s, basis, d1904), dsc: fraction(ps, s, m, basis, d1904) };
};

// ---- Depreciation helpers (spreadsheet's DDB and VDB definitions)

const ddbOf = (cost: number, salvage: number, life: number, period: number, factor: number): number => {
  let rate = factor / life, old: number;
  if (rate >= 1) { rate = 1; old = period === 1 ? cost : 0; } else old = cost * (1 - rate) ** (period - 1);
  const nw = cost * (1 - rate) ** period;
  return Math.max(0, nw < salvage ? old - salvage : old - nw);
};
const interVdb = (cost: number, salvage: number, life: number, life1: number, period: number, factor: number): number => {
  let vdb = 0, sln = 0, rest = cost - salvage, nowSln = false;
  const end = Math.ceil(period);
  for (let i = 1; i <= end; i++) {
    let term: number;
    if (!nowSln) {
      const ddb = ddbOf(cost, salvage, life, i, factor);
      sln = rest / (life1 - (i - 1));
      if (sln > ddb) { term = sln; nowSln = true; } else { term = ddb; rest -= ddb; }
    } else term = sln;
    if (i === end) term *= period + 1 - end;
    vdb += term;
  }
  return vdb;
};
const vdbOf = (cost: number, salvage: number, life: number, start: number, end: number, factor: number, noSwitch: boolean): number => {
  const is = Math.floor(start), ie = Math.ceil(end);
  if (noSwitch) {
    let vdb = 0;
    for (let i = is + 1; i <= ie; i++) {
      let term = ddbOf(cost, salvage, life, i, factor);
      if (i === is + 1) term *= Math.min(end, is + 1) - start;
      else if (i === ie) term *= end + 1 - ie;
      vdb += term;
    }
    return vdb;
  }
  let part = 0;
  if (start !== is) {
    const v = cost - interVdb(cost, salvage, life, life, is, factor);
    part += (start - is) * interVdb(v, salvage, life, life - is, 1, factor);
  }
  if (end !== ie) {
    const v = cost - interVdb(cost, salvage, life, life, ie - 1, factor);
    part += (ie - end) * interVdb(v, salvage, life, life - (ie - 1), 1, factor);
  }
  const c = cost - interVdb(cost, salvage, life, life, is, factor);
  return interVdb(c, salvage, life, life - is, ie - is, factor) - part;
};

// ---- Cash-flow series (NPV, IRR, MIRR, XNPV, XIRR)

/** Numbers in a values argument: only real numbers count inside ranges/arrays; errors propagate. */
const series = (v: Value, ctx: FnContext): number[] | CellError => collectNumbers([v], ctx);

const npvOf = (r: number, vs: number[], from = 1): number => vs.reduce((s, v, i) => s + v / (1 + r) ** (i + from), 0);

/** XNPV/XIRR inputs: values and dates paired one-to-one, all numeric, no date before the first. */
const xflows = (vArg: Value, dArg: Value, ctx: FnContext): { v: number[]; d: number[] } | CellError => {
  const v: number[] = [], d: number[] = [];
  for (const c of cellsOf(vArg, ctx)) {
    if (isError(c)) return c;
    if (typeof c !== 'number') return Err.value;
    v.push(c);
  }
  for (const c of cellsOf(dArg, ctx)) {
    if (isError(c)) return c;
    const n = typeof c === 'number' ? c : typeof c === 'string' && !fromRange(dArg) ? toNumber(c, ctx.date1904) : Err.value;
    if (isError(n)) return n;
    if (n < 0) return Err.num;
    d.push(Math.floor(n));
  }
  if (v.length !== d.length || !v.length) return Err.num;
  if (d.some((x) => x < d[0]!)) return Err.num;
  return { v, d };
};
const xnpvOf = (r: number, v: number[], d: number[]): number => v.reduce((s, x, i) => s + x / (1 + r) ** ((d[i]! - d[0]!) / 365), 0);

const hasBothSigns = (vs: number[]): boolean => vs.some((x) => x > 0) && vs.some((x) => x < 0);

export const defs: FunctionDef[] = [
  // Time value of money
  def('FV', [undefined, undefined, undefined, 0, 0], ([r, n, pmt, pv, t]) => fvOf(r!, n!, pmt!, pv!, typeOf(t!))),
  def('PV', [undefined, undefined, undefined, 0, 0], ([r, n, pmt, fv, t]) => pvOf(r!, n!, pmt!, fv!, typeOf(t!))),
  def('PMT', [undefined, undefined, undefined, 0, 0], ([r, n, pv, fv, t]) => (n === 0 ? Err.num : pmtOf(r!, n!, pv!, fv!, typeOf(t!)))),
  def('IPMT', [undefined, undefined, undefined, undefined, 0, 0], ([r, per, n, pv, fv, t]) =>
    per! < 1 || per! > n! ? Err.num : ipmtOf(r!, per!, n!, pv!, fv!, typeOf(t!))),
  def('PPMT', [undefined, undefined, undefined, undefined, 0, 0], ([r, per, n, pv, fv, t]) =>
    per! < 1 || per! > n! ? Err.num : pmtOf(r!, n!, pv!, fv!, typeOf(t!)) - ipmtOf(r!, per!, n!, pv!, fv!, typeOf(t!))),
  def('NPER', [undefined, undefined, undefined, 0, 0], ([r, pmt, pv, fv, tt]) => {
    const t = typeOf(tt!);
    if (r === 0) return pmt === 0 ? Err.num : -(pv! + fv!) / pmt!;
    const x = (pmt! * (1 + r! * t) - fv! * r!) / (pmt! * (1 + r! * t) + pv! * r!);
    return x <= 0 || r! <= -1 ? Err.num : Math.log(x) / Math.log(1 + r!);
  }),
  def('RATE', [undefined, undefined, undefined, 0, 0, 0.1], ([n, pmt, pv, fv, tt, guess]) => {
    const t = typeOf(tt!);
    if (n! <= 0) return Err.num;
    const f = (r: number): number => (Math.abs(r) < 1e-12 ? pv! + pmt! * n! + fv! : pv! * (1 + r) ** n! + (pmt! * (1 + r * t) * ((1 + r) ** n! - 1)) / r + fv!);
    return solve(f, guess!);
  }),
  def('ISPMT', [undefined, undefined, undefined, undefined], ([r, per, n, pv]) => (n === 0 ? Err.div0 : pv! * r! * (per! / n! - 1))),
  def('CUMIPMT', [undefined, undefined, undefined, undefined, undefined, undefined], ([r, nn, pv, s, e, tt]) => {
    const n = nn!, start = Math.trunc(s!), end = Math.trunc(e!), t = tt!;
    if (r! <= 0 || n <= 0 || pv! <= 0 || start < 1 || end < start || end > n || (t !== 0 && t !== 1)) return Err.num;
    const pmt = pmtOf(r!, n, pv!, 0, t);
    let ip = 0, i = start;
    if (i === 1) { if (t === 0) ip = -pv!; i++; }
    for (; i <= end; i++) ip += t === 1 ? fvOf(r!, i - 2, pmt, pv!, 1) - pmt : fvOf(r!, i - 1, pmt, pv!, 0);
    return ip * r!;
  }),
  def('CUMPRINC', [undefined, undefined, undefined, undefined, undefined, undefined], ([r, nn, pv, s, e, tt]) => {
    const n = nn!, start = Math.trunc(s!), end = Math.trunc(e!), t = tt!;
    if (r! <= 0 || n <= 0 || pv! <= 0 || start < 1 || end < start || end > n || (t !== 0 && t !== 1)) return Err.num;
    const pmt = pmtOf(r!, n, pv!, 0, t);
    let pp = 0, i = start;
    if (i === 1) { pp = t === 0 ? pmt + pv! * r! : pmt; i++; }
    for (; i <= end; i++) pp += t === 1 ? pmt - (fvOf(r!, i - 2, pmt, pv!, 1) - pmt) * r! : pmt - fvOf(r!, i - 1, pmt, pv!, 0) * r!;
    return pp;
  }),
  def('EFFECT', [undefined, undefined], ([nom, p]) => {
    const n = Math.trunc(p!);
    return nom! <= 0 || n < 1 ? Err.num : (1 + nom! / n) ** n - 1;
  }),
  def('NOMINAL', [undefined, undefined], ([eff, p]) => {
    const n = Math.trunc(p!);
    return eff! <= 0 || n < 1 ? Err.num : n * ((1 + eff!) ** (1 / n) - 1);
  }),
  def('PDURATION', [undefined, undefined, undefined], ([r, pv, fv]) =>
    r! <= 0 || pv! <= 0 || fv! <= 0 ? Err.num : (Math.log(fv!) - Math.log(pv!)) / Math.log(1 + r!)),
  def('RRI', [undefined, undefined, undefined], ([n, pv, fv]) => (n! <= 0 || pv === 0 ? Err.num : (fv! / pv!) ** (1 / n!) - 1)),
  def('DOLLARDE', [undefined, undefined], ([x, f]) => {
    const fr = Math.trunc(f!);
    if (fr < 0) return Err.num;
    if (fr === 0) return Err.div0;
    const i = Math.trunc(x!);
    return i + ((x! - i) * 10 ** Math.ceil(Math.log10(fr))) / fr;
  }),
  def('DOLLARFR', [undefined, undefined], ([x, f]) => {
    const fr = Math.trunc(f!);
    if (fr < 0) return Err.num;
    if (fr === 0) return Err.div0;
    const i = Math.trunc(x!);
    return i + ((x! - i) * fr) / 10 ** Math.ceil(Math.log10(fr));
  }),

  // Cash-flow series
  {
    name: 'NPV', min: 2, max: 255,
    fn: (args, ctx) => {
      const r = num(args[0]!, ctx); if (isError(r)) return r;
      const vs = collectNumbers(args.slice(1), ctx); if (isError(vs)) return vs;
      return r === -1 ? Err.div0 : fin(npvOf(r, vs));
    },
  },
  {
    name: 'IRR', min: 1, max: 2,
    fn: ([v, g], ctx) => {
      const vs = series(v!, ctx); if (isError(vs)) return vs;
      const guess = g === undefined || g === null ? 0.1 : num(g, ctx); if (isError(guess)) return guess;
      if (!hasBothSigns(vs)) return Err.num;
      return solve((r) => npvOf(r, vs, 0), guess);
    },
  },
  {
    name: 'MIRR', min: 3, max: 3,
    fn: ([v, f, rr], ctx) => {
      const vs = series(v!, ctx); if (isError(vs)) return vs;
      const fr = num(f!, ctx); if (isError(fr)) return fr;
      const rer = num(rr!, ctx); if (isError(rer)) return rer;
      if (!hasBothSigns(vs)) return Err.div0;
      const n = vs.length;
      const pos = npvOf(rer, vs.map((x) => (x > 0 ? x : 0)), 0) * (1 + rer) ** (n - 1);
      const neg = npvOf(fr, vs.map((x) => (x < 0 ? x : 0)), 0);
      return fin((-pos / neg) ** (1 / (n - 1)) - 1);
    },
  },
  {
    name: 'XNPV', min: 3, max: 3,
    fn: ([r, v, d], ctx) => {
      const rate = num(r!, ctx); if (isError(rate)) return rate;
      const x = xflows(v!, d!, ctx); if (isError(x)) return x;
      return rate <= -1 ? Err.num : fin(xnpvOf(rate, x.v, x.d));
    },
  },
  {
    name: 'XIRR', min: 2, max: 3,
    fn: ([v, d, g], ctx) => {
      const x = xflows(v!, d!, ctx); if (isError(x)) return x;
      const guess = g === undefined || g === null ? 0.1 : num(g, ctx); if (isError(guess)) return guess;
      if (guess <= -1 || !hasBothSigns(x.v)) return Err.num;
      return solve((r) => xnpvOf(r, x.v, x.d), guess);
    },
  },
  {
    name: 'FVSCHEDULE', min: 2, max: 2, scalar: [0],
    fn: ([p, s], ctx) => {
      const pr = num(p!, ctx); if (isError(pr)) return pr;
      let v = pr;
      for (const c of cellsOf(s!, ctx)) {
        const r = toNumber(c, ctx.date1904);
        if (isError(r)) return r;
        v *= 1 + r;
      }
      return fin(v);
    },
  },

  // Depreciation
  def('SLN', [undefined, undefined, undefined], ([c, s, l]) => (l === 0 ? Err.div0 : (c! - s!) / l!)),
  def('SYD', [undefined, undefined, undefined, undefined], ([c, s, l, p]) =>
    l! <= 0 || p! <= 0 || p! > l! || s! < 0 ? Err.num : ((c! - s!) * (l! - p! + 1) * 2) / (l! * (l! + 1))),
  def('DB', [undefined, undefined, undefined, undefined, 12], ([cost, salvage, life, period, month]) => {
    const c = cost!, s = salvage!, l = life!, p = Math.trunc(period!), m = Math.trunc(month!);
    if (c < 0 || s < 0 || l <= 0 || p <= 0 || m < 1 || m > 12 || p > l + (m < 12 ? 1 : 0)) return Err.num;
    if (c === 0) return 0;
    const rate = Math.round((1 - (s / c) ** (1 / l)) * 1000) / 1000;
    let total = (c * rate * m) / 12;
    if (p === 1) return total;
    for (let i = 2; i < p; i++) total += (c - total) * rate;
    return p === Math.floor(l) + 1 ? ((c - total) * rate * (12 - m)) / 12 : (c - total) * rate;
  }),
  def('DDB', [undefined, undefined, undefined, undefined, 2], ([c, s, l, p, f]) =>
    c! < 0 || s! < 0 || l! <= 0 || p! <= 0 || f! <= 0 || p! > l! ? Err.num : ddbOf(c!, s!, l!, p!, f!)),
  {
    name: 'VDB', min: 5, max: 7, scalar: true,
    fn: (args, ctx) => {
      const n = nums(args, ctx, [undefined, undefined, undefined, undefined, undefined, 2]); if (isError(n)) return n;
      const ns = optBool(args[6], ctx, false); if (isError(ns)) return ns;
      const [c, s, l, st, en, f] = n as [number, number, number, number, number, number];
      if (c < 0 || s < 0 || l <= 0 || st < 0 || en < st || en > l || f <= 0) return Err.num;
      return fin(vdbOf(c, s, l, st, en, f, ns));
    },
  },
  def('AMORLINC', [undefined, 'date', 'date', undefined, undefined, undefined, 0], ([cost, pd, fp, salvage, period, rate, basis], ctx) => {
    const b = Math.trunc(basis!), per = Math.trunc(period!);
    if (!basisOk(b) || b === 2 || cost! < 0 || salvage! < 0 || salvage! > cost! || rate! <= 0 || per < 0 || pd! > fp!) return Err.num;
    const one = cost! * rate!, delta = cost! - salvage!;
    const first = yearFrac(pd!, fp!, b, ctx.date1904) * rate! * cost!;
    const full = Math.trunc((delta - first) / one);
    if (per === 0) return Math.min(first, delta);
    if (per <= full) return one;
    if (per === full + 1) return Math.max(0, delta - one * full - first);
    return 0;
  }),
  def('AMORDEGRC', [undefined, 'date', 'date', undefined, undefined, undefined, 0], ([cost, pd, fp, salvage, period, rate, basis], ctx) => {
    const b = Math.trunc(basis!), per = Math.trunc(period!);
    if (!basisOk(b) || b === 2 || cost! < 0 || salvage! < 0 || salvage! > cost! || rate! <= 0 || per < 0 || pd! > fp!) return Err.num;
    const life = 1 / rate!;
    if ((life > 0 && life < 1) || (life > 1 && life < 2) || (life > 2 && life < 3) || (life > 4 && life < 5)) return Err.num;
    const r = rate! * (life < 3 ? 1 : life < 5 ? 1.5 : life <= 6 ? 2 : 2.5);
    let dep = Math.round(yearFrac(pd!, fp!, b, ctx.date1904) * r * cost!);
    let c = cost! - dep, rest = c - salvage!;
    for (let n = 0; n < per; n++) {
      dep = Math.round(r * c);
      rest -= dep;
      if (rest < 0) return per - n <= 1 ? Math.round(c * 0.5) : 0;
      c -= dep;
    }
    return dep;
  }),

  // Coupons
  ...(['COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD'] as const).map((name) =>
    def(name, ['date', 'date', undefined, 0], ([s, m, f, bb], ctx) => {
      const freq = Math.trunc(f!), basis = Math.trunc(bb!), e = secCheck(s!, m!, freq, basis);
      if (e) return e;
      const c = coupons(s!, m!, freq, ctx.date1904), b = bond(s!, m!, freq, basis, ctx.date1904);
      switch (name) {
        case 'COUPDAYBS': return b.A;
        case 'COUPDAYS': return b.E;
        case 'COUPDAYSNC': return b.DSC;
        case 'COUPNCD': return c.ncd;
        case 'COUPNUM': return c.num;
        case 'COUPPCD': return c.pcd;
      }
    })),

  // Securities with periodic interest
  def('ACCRINT', ['date', 'date', 'date', undefined, 1000, undefined, 0, 1], ([iss, fi, s, rate, par, f, bb, cm], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    if (rate! <= 0 || par! <= 0 || !freqOk(freq) || !basisOk(basis) || iss! >= s!) return Err.num;
    const months = 12 / freq, eom = isMonthEnd(fi!, ctx.date1904);
    const at = (k: number): number => addMonths(fi!, k * months, ctx.date1904, eom);
    let start = iss!;
    if (!cm && s! > fi!) { let k = 0; while (at(k + 1) <= s!) k++; start = at(k); }
    let k = 0;
    while (at(k) > start) k--;
    while (at(k + 1) <= start) k++;
    let sum = 0;
    for (; at(k) < s!; k++) {
      const p = at(k), n = at(k + 1), a = Math.max(p, start), b = Math.min(n, s!);
      if (b > a) sum += dayCount(a, b, basis, ctx.date1904) / periodDays(p, n, basis, freq);
    }
    return (par! * rate! * sum) / freq;
  }),
  def('ACCRINTM', ['date', 'date', undefined, 1000, 0], ([iss, s, rate, par, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (rate! <= 0 || par! <= 0 || !basisOk(basis) || iss! >= s!) return Err.num;
    return par! * rate! * yearFrac(iss!, s!, basis, ctx.date1904);
  }),
  def('PRICE', ['date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, rate, yld, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    const e = secCheck(s!, m!, freq, basis); if (e) return e;
    if (rate! < 0 || yld! < 0 || red! <= 0) return Err.num;
    return priceOf(bond(s!, m!, freq, basis, ctx.date1904), rate!, yld!, red!, freq);
  }),
  def('YIELD', ['date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, rate, pr, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    const e = secCheck(s!, m!, freq, basis); if (e) return e;
    if (rate! < 0 || pr! <= 0 || red! <= 0) return Err.num;
    const b = bond(s!, m!, freq, basis, ctx.date1904);
    if (b.N === 1) {
      const c = rate! / freq, dirty = pr! / 100 + (b.A / b.E) * c;
      return ((red! / 100 + c - dirty) / dirty) * ((freq * b.E) / b.DSC);
    }
    return solve((y) => priceOf(b, rate!, y, red!, freq) - pr!, 0.05, -freq + 1e-10);
  }),
  def('DURATION', ['date', 'date', undefined, undefined, undefined, 0], ([s, m, c, y, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    const e = secCheck(s!, m!, freq, basis); if (e) return e;
    if (c! < 0 || y! < 0) return Err.num;
    return durationOf(s!, m!, c!, y!, freq, basis, ctx.date1904);
  }),
  def('MDURATION', ['date', 'date', undefined, undefined, undefined, 0], ([s, m, c, y, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    const e = secCheck(s!, m!, freq, basis); if (e) return e;
    if (c! < 0 || y! < 0) return Err.num;
    return durationOf(s!, m!, c!, y!, freq, basis, ctx.date1904) / (1 + y! / freq);
  }),
  def('ODDFPRICE', ['date', 'date', 'date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, iss, fc, rate, yld, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    if (!freqOk(freq) || !basisOk(basis) || !(m! > fc! && fc! > s! && s! > iss!) || rate! < 0 || yld! < 0 || red! <= 0) return Err.num;
    return oddFirstPrice(s!, m!, iss!, fc!, rate!, yld!, red!, freq, basis, ctx.date1904);
  }),
  def('ODDFYIELD', ['date', 'date', 'date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, iss, fc, rate, pr, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    if (!freqOk(freq) || !basisOk(basis) || !(m! > fc! && fc! > s! && s! > iss!) || rate! < 0 || pr! <= 0 || red! <= 0) return Err.num;
    return solve((y) => oddFirstPrice(s!, m!, iss!, fc!, rate!, y, red!, freq, basis, ctx.date1904) - pr!, rate! || 0.05, -freq + 1e-10);
  }),
  def('ODDLPRICE', ['date', 'date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, last, rate, yld, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    if (!freqOk(freq) || !basisOk(basis) || !(m! > s! && s! > last!) || rate! < 0 || yld! < 0 || red! <= 0) return Err.num;
    const o = oddLast(s!, m!, last!, freq, basis, ctx.date1904), c = (100 * rate!) / freq;
    return (red! + o.dc * c) / (1 + (o.dsc * yld!) / freq) - o.a * c;
  }),
  def('ODDLYIELD', ['date', 'date', 'date', undefined, undefined, undefined, undefined, 0], ([s, m, last, rate, pr, red, f, bb], ctx) => {
    const freq = Math.trunc(f!), basis = Math.trunc(bb!);
    if (!freqOk(freq) || !basisOk(basis) || !(m! > s! && s! > last!) || rate! < 0 || pr! <= 0 || red! <= 0) return Err.num;
    const o = oddLast(s!, m!, last!, freq, basis, ctx.date1904), c = (100 * rate!) / freq;
    const dirty = pr! + o.a * c;
    return ((red! + o.dc * c - dirty) / dirty) * (freq / o.dsc);
  }),

  // Discounted and at-maturity securities
  def('DISC', ['date', 'date', undefined, undefined, 0], ([s, m, pr, red, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || pr! <= 0 || red! <= 0 || !basisOk(basis)) return Err.num;
    return (red! - pr!) / red! / yearFrac(s!, m!, basis, ctx.date1904);
  }),
  def('INTRATE', ['date', 'date', undefined, undefined, 0], ([s, m, inv, red, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || inv! <= 0 || red! <= 0 || !basisOk(basis)) return Err.num;
    return (red! - inv!) / inv! / yearFrac(s!, m!, basis, ctx.date1904);
  }),
  def('RECEIVED', ['date', 'date', undefined, undefined, 0], ([s, m, inv, disc, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || inv! <= 0 || disc! <= 0 || !basisOk(basis)) return Err.num;
    const d = 1 - disc! * yearFrac(s!, m!, basis, ctx.date1904);
    return d <= 0 ? Err.num : inv! / d;
  }),
  def('PRICEDISC', ['date', 'date', undefined, undefined, 0], ([s, m, disc, red, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || disc! <= 0 || red! <= 0 || !basisOk(basis)) return Err.num;
    return red! * (1 - disc! * yearFrac(s!, m!, basis, ctx.date1904));
  }),
  def('YIELDDISC', ['date', 'date', undefined, undefined, 0], ([s, m, pr, red, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || pr! <= 0 || red! <= 0 || !basisOk(basis)) return Err.num;
    return (red! - pr!) / pr! / yearFrac(s!, m!, basis, ctx.date1904);
  }),
  def('PRICEMAT', ['date', 'date', 'date', undefined, undefined, 0], ([s, m, iss, rate, yld, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || rate! < 0 || yld! < 0 || !basisOk(basis)) return Err.num;
    const yf = (a: number, b: number): number => yearFrac(a, b, basis, ctx.date1904);
    return (100 + yf(iss!, m!) * rate! * 100) / (1 + yf(s!, m!) * yld!) - yf(iss!, s!) * rate! * 100;
  }),
  def('YIELDMAT', ['date', 'date', 'date', undefined, undefined, 0], ([s, m, iss, rate, pr, bb], ctx) => {
    const basis = Math.trunc(bb!);
    if (s! >= m! || rate! < 0 || pr! <= 0 || !basisOk(basis)) return Err.num;
    const yf = (a: number, b: number): number => yearFrac(a, b, basis, ctx.date1904);
    return ((1 + yf(iss!, m!) * rate!) / (pr! / 100 + yf(iss!, s!) * rate!) - 1) / yf(s!, m!);
  }),
  def('TBILLPRICE', ['date', 'date', undefined], ([s, m, disc]) => {
    const dsm = m! - s!;
    if (dsm <= 0 || dsm > 366 || disc! <= 0) return Err.num;
    const p = 100 * (1 - (disc! * dsm) / 360);
    return p <= 0 ? Err.num : p;
  }),
  def('TBILLYIELD', ['date', 'date', undefined], ([s, m, pr]) => {
    const dsm = m! - s!;
    if (dsm <= 0 || dsm > 366 || pr! <= 0) return Err.num;
    return ((100 - pr!) / pr!) * (360 / dsm);
  }),
  def('TBILLEQ', ['date', 'date', undefined], ([s, m, disc]) => {
    const dsm = m! - s!;
    if (dsm <= 0 || dsm > 366 || disc! <= 0) return Err.num;
    if (dsm <= 182) return (365 * disc!) / (360 - disc! * dsm);
    // Over half a year the standard method solves the bond-equivalent quadratic.
    const price = 100 * (1 - (disc! * dsm) / 360);
    if (price <= 0) return Err.num;
    const t = dsm / 365;
    return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (2 * t - 1);
  }),
];
