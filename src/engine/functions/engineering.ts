import { CellError, Err, isError, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { cellsOf, first, numberToText, toNumber } from '../coerce';

// Engineering functions refuse logical values where they expect numbers or text (=DEC2BIN(TRUE) is #VALUE!).
const eNum = (v: Value | undefined, ctx: FnContext): number | CellError => {
  const s = v === undefined ? null : first(v, ctx);
  return typeof s === 'boolean' ? Err.value : toNumber(s, ctx.date1904);
};
const eOpt = (v: Value | undefined, ctx: FnContext, dflt: number): number | CellError =>
  v === undefined || v === null ? dflt : eNum(v, ctx);
const eText = (v: Value | undefined, ctx: FnContext): string | CellError => {
  const s = v === undefined ? null : first(v, ctx);
  if (s === null) return '';
  if (typeof s === 'boolean') return Err.value;
  if (isError(s)) return s;
  return typeof s === 'number' ? numberToText(s) : s;
};
const fin = (x: number | CellError): number | CellError => (isError(x) ? x : Number.isFinite(x) ? (x === 0 ? 0 : x) : Err.num);
const snap = (x: number): number => (Number.isFinite(x) && x !== 0 ? Number(x.toPrecision(15)) : x);

// ---------------------------------------------------------------- Bessel functions

/** Composite 20-point Gauss–Legendre nodes and weights on [-1, 1], computed once by Newton's method. */
const GL = (() => {
  const n = 20, xs: number[] = [], ws: number[] = [];
  for (let i = 1; i <= n; i++) {
    let x = Math.cos((Math.PI * (i - 0.25)) / (n + 0.5)), dp = 0;
    for (let it = 0; it < 100; it++) {
      let p0 = 1, p1 = x;
      for (let k = 2; k <= n; k++) [p0, p1] = [p1, ((2 * k - 1) * x * p1 - (k - 1) * p0) / k];
      dp = (n * (x * p1 - p0)) / (x * x - 1);
      const dx = p1 / dp;
      x -= dx;
      if (Math.abs(dx) < 1e-16) break;
    }
    xs.push(x); ws.push(2 / ((1 - x * x) * dp * dp));
  }
  return { xs, ws };
})();
const integrate = (f: (t: number) => number, a: number, b: number, panels: number): number => {
  const h = (b - a) / panels;
  let s = 0;
  for (let p = 0; p < panels; p++) {
    const mid = a + (p + 0.5) * h;
    for (let i = 0; i < GL.xs.length; i++) s += GL.ws[i] * f(mid + (GL.xs[i] * h) / 2);
  }
  return (s * h) / 2;
};

const logFact = (n: number): number => {
  let s = 0;
  for (let i = 2; i <= n; i++) s += Math.log(i);
  return s;
};

/** Power series Σ s^k (x/2)^(2k+n) / (k!(n+k)!) with s = -1 for J and +1 for I. */
const besselSeries = (x: number, n: number, sign: number): number => {
  const h = Math.abs(x) / 2;
  if (h === 0) return n === 0 ? 1 : 0;
  let term = Math.exp(n * Math.log(h) - logFact(n)), sum = term;
  for (let k = 1; k < 2000; k++) {
    term *= (sign * h * h) / (k * (n + k));
    sum += term;
    if (Math.abs(term) <= Math.abs(sum) * 1e-17) break;
  }
  return x < 0 && n % 2 ? -sum : sum;
};

/** J_n(x) = (1/2π)∫ cos(nτ − x sin τ) dτ over a period; the trapezoidal rule converges exponentially. */
const besselJ = (x: number, n: number): number => {
  if ((x * x) / 4 < 4 * (n + 1)) return besselSeries(x, n, -1); // series terms only grow past this, losing digits
  const N = Math.ceil(n + Math.abs(x) + 64);
  let s = 0;
  for (let k = 0; k < N; k++) {
    const t = (2 * Math.PI * k) / N;
    s += Math.cos(n * t - x * Math.sin(t));
  }
  return s / N;
};

/** Y_0 and Y_1 from their integral representations, then the stable forward recurrence. */
const besselY01 = (x: number, n: 0 | 1): number => {
  const osc = integrate((t) => Math.sin(x * Math.sin(t) - n * t), 0, Math.PI, Math.ceil(x) + 4);
  const T = Math.asinh(60 / x) + 1;
  const tail = integrate((t) => (Math.exp(n * t) + (n ? -1 : 1) * Math.exp(-n * t)) * Math.exp(-x * Math.sinh(t)), 0, T, Math.ceil(T * 16) + 8);
  return (osc - tail) / Math.PI;
};
const besselY = (x: number, n: number): number => {
  let y0 = besselY01(x, 0);
  if (n === 0) return y0;
  let y1 = besselY01(x, 1);
  for (let k = 1; k < n; k++) [y0, y1] = [y1, ((2 * k) / x) * y1 - y0];
  return y1;
};

/** K_n(x) = ∫₀^∞ e^(−x cosh t) cosh(nt) dt; the integrand is even, so the trapezoidal rule is spectrally accurate. */
const besselK01 = (x: number, n: 0 | 1): number => {
  const h = 0.05;
  let s = 0.5;
  for (let k = 1; k < 20000; k++) {
    const t = k * h, term = Math.exp(-x * (Math.cosh(t) - 1)) * (n ? Math.cosh(t) : 1);
    s += term;
    if (term < s * 1e-18) break;
  }
  return s * h * Math.exp(-x);
};
const besselK = (x: number, n: number): number => {
  let k0 = besselK01(x, 0);
  if (n === 0) return k0;
  let k1 = besselK01(x, 1);
  for (let k = 1; k < n; k++) [k0, k1] = [k1, ((2 * k) / x) * k1 + k0];
  return k1;
};

const bessel = (name: string, f: (x: number, n: number) => number, positiveX: boolean): FunctionDef => ({
  name, min: 2, max: 2, scalar: true,
  fn: ([xv, nv], ctx) => {
    const x = eNum(xv, ctx); if (isError(x)) return x;
    const nr = eNum(nv, ctx); if (isError(nr)) return nr;
    const n = Math.trunc(nr);
    if (n < 0 || (positiveX && x <= 0)) return Err.num;
    return fin(f(x, n));
  },
});

// ---------------------------------------------------------------- base conversion

interface Base { radix: number; digits: RegExp; /** Largest positive value; negatives go down to -(max+1). */ max: number }
const BIN: Base = { radix: 2, digits: /^[01]{0,10}$/, max: 511 };
const OCT: Base = { radix: 8, digits: /^[0-7]{0,10}$/, max: 2 ** 29 - 1 };
const HEX: Base = { radix: 16, digits: /^[0-9A-F]{0,10}$/i, max: 2 ** 39 - 1 };

/** Reads a 10-digit two's-complement string: a full-width value with its top bit set is negative. */
const parseBase = (v: Value | undefined, ctx: FnContext, b: Base): number | CellError => {
  const t = eText(v, ctx);
  if (isError(t)) return t;
  if (!b.digits.test(t)) return Err.num;
  if (!t) return 0;
  const n = parseInt(t, b.radix), full = b.radix ** 10;
  return t.length === 10 && n > (full / 2) - 1 ? n - full : n;
};

/** Writes n in base b; negatives become ten digits and ignore `places`. */
const formatBase = (n: number, b: Base, placesArg: Value | undefined, ctx: FnContext): string | CellError => {
  if (n < -(b.max + 1) || n > b.max) return Err.num;
  if (n < 0) return (b.radix ** 10 + n).toString(b.radix).toUpperCase();
  const s = n.toString(b.radix).toUpperCase();
  if (placesArg === undefined || placesArg === null) return s;
  const p = eNum(placesArg, ctx);
  if (isError(p)) return p;
  const places = Math.trunc(p);
  if (places <= 0 || places > 10 || s.length > places) return Err.num;
  return s.padStart(places, '0');
};

const fromBase = (name: string, src: Base, to: Base | null): FunctionDef => ({
  name, min: 1, max: to ? 2 : 1, scalar: true,
  fn: ([v, places], ctx) => {
    const n = parseBase(v, ctx, src);
    if (isError(n)) return n;
    return to ? formatBase(n, to, places, ctx) : n;
  },
});

const fromDec = (name: string, to: Base): FunctionDef => ({
  name, min: 1, max: 2, scalar: true,
  fn: ([v, places], ctx) => {
    const n = eNum(v, ctx);
    if (isError(n)) return n;
    return formatBase(Math.trunc(n), to, places, ctx);
  },
});

// ---------------------------------------------------------------- bitwise

const BIT_LIMIT = 2 ** 48;
const bitArg = (v: Value | undefined, ctx: FnContext): number | CellError => {
  const n = eNum(v, ctx);
  if (isError(n)) return n;
  return n < 0 || n >= BIT_LIMIT || !Number.isInteger(n) ? Err.num : n;
};
const bitOp = (name: string, f: (a: bigint, b: bigint) => bigint): FunctionDef => ({
  name, min: 2, max: 2, scalar: true,
  fn: ([a, b], ctx) => {
    const x = bitArg(a, ctx); if (isError(x)) return x;
    const y = bitArg(b, ctx); if (isError(y)) return y;
    return Number(f(BigInt(x), BigInt(y)));
  },
});
const bitShift = (name: string, dir: 1 | -1): FunctionDef => ({
  name, min: 2, max: 2, scalar: true,
  fn: ([a, s], ctx) => {
    const x = bitArg(a, ctx); if (isError(x)) return x;
    const sr = eNum(s, ctx); if (isError(sr)) return sr;
    const k = Math.trunc(sr) * dir;
    if (Math.abs(k) > 53) return Err.num;
    const r = k >= 0 ? BigInt(x) << BigInt(k) : BigInt(x) >> BigInt(-k);
    return r >= BigInt(BIT_LIMIT) ? Err.num : Number(r);
  },
});

// ---------------------------------------------------------------- complex numbers

interface Complex { re: number; im: number; unit: 'i' | 'j' | '' }

const REAL_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
const IMAG_RE = /^[+-]?(?:(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?$/;

/** Parses spreadsheet's complex text: "3+4i", "-2.5j", "i", "1e3-2E-1i", or a plain number. */
const parseComplex = (s: Scalar): Complex | CellError => {
  if (isError(s)) return s;
  if (s === null || s === '') return { re: 0, im: 0, unit: '' };
  if (typeof s === 'boolean') return Err.value;
  if (typeof s === 'number') return { re: s, im: 0, unit: '' };
  const last = s[s.length - 1];
  if (last !== 'i' && last !== 'j') return REAL_RE.test(s) ? { re: Number(s), im: 0, unit: '' } : Err.num;
  const body = s.slice(0, -1);
  // The imaginary part starts at the last sign that is not leading and not an exponent sign.
  let split = 0;
  for (let i = body.length - 1; i > 0; i--) {
    if ((body[i] === '+' || body[i] === '-') && !/[eE]/.test(body[i - 1])) { split = i; break; }
  }
  const re = body.slice(0, split), im = body.slice(split);
  if ((re && !REAL_RE.test(re)) || !IMAG_RE.test(im)) return Err.num;
  const imNum = im === '' || im === '+' ? 1 : im === '-' ? -1 : Number(im);
  return { re: re ? Number(re) : 0, im: imNum, unit: last };
};

/** Formats as spreadsheets do: 15 significant digits, "i"/"-i" for unit imaginary parts, and no zero parts. */
const formatComplex = (re: number, im: number, unit: 'i' | 'j' | ''): string | CellError => {
  if (!Number.isFinite(re) || !Number.isFinite(im)) return Err.num;
  const u = unit || 'i';
  const r = snap(re), i = snap(im);
  const imText = i === 1 ? '' : i === -1 ? '-' : numberToText(i);
  if (i === 0) return numberToText(r);
  if (r === 0) return imText + u;
  return numberToText(r) + (i > 0 ? '+' : '') + imText + u;
};

/** Mixing "i" and "j" suffixes is #VALUE!. */
const unitOf = (...cs: Complex[]): 'i' | 'j' | '' | CellError => {
  let u: 'i' | 'j' | '' = '';
  for (const c of cs) {
    if (!c.unit) continue;
    if (u && c.unit !== u) return Err.value;
    u = c.unit;
  }
  return u;
};

const cArg = (v: Value | undefined, ctx: FnContext): Complex | CellError => parseComplex(v === undefined ? null : first(v, ctx));

type C = [number, number];
const cmul = ([a, b]: C, [c, d]: C): C => [a * c - b * d, a * d + b * c];
const cdiv = ([a, b]: C, [c, d]: C): C => {
  const den = c * c + d * d;
  return [(a * c + b * d) / den, (b * c - a * d) / den];
};
const cexp = ([a, b]: C): C => [Math.exp(a) * Math.cos(b), Math.exp(a) * Math.sin(b)];
const cln = ([a, b]: C): C => [Math.log(Math.hypot(a, b)), Math.atan2(b, a)];
const csin = ([a, b]: C): C => [Math.sin(a) * Math.cosh(b), Math.cos(a) * Math.sinh(b)];
const ccos = ([a, b]: C): C => [Math.cos(a) * Math.cosh(b), -Math.sin(a) * Math.sinh(b)];
const csinh = ([a, b]: C): C => [Math.sinh(a) * Math.cos(b), Math.cosh(a) * Math.sin(b)];
const ccosh = ([a, b]: C): C => [Math.cosh(a) * Math.cos(b), Math.sinh(a) * Math.sin(b)];
const recip = (c: C): C => cdiv([1, 0], c);

/** Unary complex → complex function; `null` from the body means #NUM!. */
const imFn = (name: string, f: (c: C) => C | null): FunctionDef => ({
  name, min: 1, max: 1, scalar: true,
  fn: ([v], ctx) => {
    const c = cArg(v, ctx);
    if (isError(c)) return c;
    const r = f([c.re, c.im]);
    return r ? formatComplex(r[0], r[1], c.unit) : Err.num;
  },
});
const isZero = ([a, b]: C): boolean => a === 0 && b === 0;

const imFold = (name: string, f: (a: C, b: C) => C, init: C): FunctionDef => ({
  name, min: 1, max: 255,
  fn: (args, ctx) => {
    let acc = init;
    let unit: 'i' | 'j' | '' = '';
    for (const a of args) {
      for (const cell of cellsOf(a, ctx)) {
        if (cell === null) continue;
        const c = parseComplex(cell);
        if (isError(c)) return c;
        const u = unitOf({ re: 0, im: 0, unit }, c);
        if (isError(u)) return u;
        unit = u;
        acc = f(acc, [c.re, c.im]);
      }
    }
    return formatComplex(acc[0], acc[1], unit);
  },
});

const imBinary = (name: string, f: (a: C, b: C) => C | null): FunctionDef => ({
  name, min: 2, max: 2, scalar: true,
  fn: ([x, y], ctx) => {
    const a = cArg(x, ctx); if (isError(a)) return a;
    const b = cArg(y, ctx); if (isError(b)) return b;
    const u = unitOf(a, b); if (isError(u)) return u;
    const r = f([a.re, a.im], [b.re, b.im]);
    return r ? formatComplex(r[0], r[1], u) : Err.num;
  },
});

// ---------------------------------------------------------------- CONVERT

type Category = 'mass' | 'length' | 'time' | 'pressure' | 'force' | 'energy' | 'power' | 'magnetism' | 'temperature' | 'volume' | 'area' | 'information' | 'speed';
interface Unit { cat: Category; factor: number; prefix?: boolean; binary?: boolean; /** Power the prefix is raised to (m2, m3). */ dim?: number }

const IN = 0.0254, FT = 0.3048, YD = 0.9144, MI = 1609.344, NMI = 1852, PICA = IN / 72, LY = 9.4607304725808e15;
const UNITS: Record<string, Unit> = {};
const add = (cat: Category, factor: number, names: string[], opt: Partial<Unit> = {}): void => {
  for (const n of names) UNITS[n] = { cat, factor, ...opt };
};
// Mass (grams)
add('mass', 1, ['g'], { prefix: true });
add('mass', 14593.902937206364, ['sg']);
add('mass', 453.59237, ['lbm']);
add('mass', 1.660538782e-24, ['u'], { prefix: true });
add('mass', 28.349523125, ['ozm']);
add('mass', 0.06479891, ['grain']);
add('mass', 45359.237, ['cwt', 'shweight']);
add('mass', 50802.34544, ['uk_cwt', 'lcwt', 'hweight']);
add('mass', 6350.29318, ['stone']);
add('mass', 907184.74, ['ton']);
add('mass', 1016046.9088, ['uk_ton', 'LTON', 'brton']);
// Length (metres)
add('length', 1, ['m'], { prefix: true });
add('length', MI, ['mi']);
add('length', NMI, ['Nmi']);
add('length', IN, ['in']);
add('length', FT, ['ft']);
add('length', YD, ['yd']);
add('length', 1e-10, ['ang'], { prefix: true });
add('length', 1.143, ['ell']);
add('length', LY, ['ly'], { prefix: true });
add('length', 3.08567758128155e16, ['parsec', 'pc'], { prefix: true });
add('length', PICA, ['Picapt', 'Pica']);
add('length', IN / 6, ['pica']);
add('length', (5280 * 1200) / 3937, ['survey_mi']);
// Time (seconds)
add('time', 31557600, ['yr']);
add('time', 86400, ['day', 'd']);
add('time', 3600, ['hr']);
add('time', 60, ['mn', 'min']);
add('time', 1, ['sec', 's'], { prefix: true });
// Pressure (pascals)
add('pressure', 1, ['Pa', 'p'], { prefix: true });
add('pressure', 101325, ['atm', 'at'], { prefix: true });
add('pressure', 133.322387415, ['mmHg'], { prefix: true });
add('pressure', 6894.75729316836, ['psi']);
add('pressure', 101325 / 760, ['Torr']);
// Force (newtons)
add('force', 1, ['N'], { prefix: true });
add('force', 1e-5, ['dyn', 'dy'], { prefix: true });
add('force', 4.4482216152605, ['lbf']);
add('force', 0.00980665, ['pond'], { prefix: true });
// Energy (joules)
add('energy', 1, ['J'], { prefix: true });
add('energy', 1e-7, ['e'], { prefix: true });
add('energy', 4.184, ['c'], { prefix: true });
add('energy', 4.1868, ['cal'], { prefix: true });
add('energy', 1.602176487e-19, ['eV', 'ev'], { prefix: true });
add('energy', 2684519.537696173, ['HPh', 'hh']);
add('energy', 3600, ['Wh', 'wh'], { prefix: true });
add('energy', 1 / 23.7304222192651, ['flb']);
add('energy', 1055.05585262, ['BTU', 'btu']);
// Power (watts)
add('power', 745.69987158227, ['HP', 'h']);
add('power', 735.49875, ['PS']);
add('power', 1, ['W', 'w'], { prefix: true });
// Magnetism (tesla)
add('magnetism', 1, ['T'], { prefix: true });
add('magnetism', 1e-4, ['ga'], { prefix: true });
// Temperature: factor unused, handled by toKelvin/fromKelvin
add('temperature', 1, ['C', 'cel', 'F', 'fah', 'Rank', 'Reau']);
add('temperature', 1, ['K', 'kel'], { prefix: true });
// Volume (cubic metres)
add('volume', 4.92892159375e-6, ['tsp']);
add('volume', 5e-6, ['tspm']);
add('volume', 1.478676478125e-5, ['tbs']);
add('volume', 2.95735295625e-5, ['oz']);
add('volume', 2.365882365e-4, ['cup']);
add('volume', 4.73176473e-4, ['pt', 'us_pt']);
add('volume', 5.6826125e-4, ['uk_pt']);
add('volume', 9.46352946e-4, ['qt']);
add('volume', 1.1365225e-3, ['uk_qt']);
add('volume', 3.785411784e-3, ['gal']);
add('volume', 4.54609e-3, ['uk_gal']);
add('volume', 1e-3, ['l', 'L', 'lt'], { prefix: true });
add('volume', 1e-30, ['ang3', 'ang^3'], { prefix: true, dim: 3 });
add('volume', 0.158987294928, ['barrel']);
add('volume', 0.03523907016688, ['bushel']);
add('volume', FT ** 3, ['ft3', 'ft^3']);
add('volume', IN ** 3, ['in3', 'in^3']);
add('volume', LY ** 3, ['ly3', 'ly^3']);
add('volume', 1, ['m3', 'm^3'], { prefix: true, dim: 3 });
add('volume', MI ** 3, ['mi3', 'mi^3']);
add('volume', YD ** 3, ['yd3', 'yd^3']);
add('volume', NMI ** 3, ['Nmi3', 'Nmi^3']);
add('volume', PICA ** 3, ['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3']);
add('volume', 2.8316846592, ['GRT', 'regton']);
add('volume', 1.13267386368, ['MTON']);
// Area (square metres)
add('area', 4046.8564224, ['uk_acre']);
add('area', 4046.872609874252, ['us_acre']);
add('area', 1e-20, ['ang2', 'ang^2'], { prefix: true, dim: 2 });
add('area', 100, ['ar'], { prefix: true });
add('area', FT ** 2, ['ft2', 'ft^2']);
add('area', 10000, ['ha']);
add('area', IN ** 2, ['in2', 'in^2']);
add('area', LY ** 2, ['ly2', 'ly^2']);
add('area', 1, ['m2', 'm^2'], { prefix: true, dim: 2 });
add('area', 2500, ['Morgen']);
add('area', MI ** 2, ['mi2', 'mi^2']);
add('area', NMI ** 2, ['Nmi2', 'Nmi^2']);
add('area', PICA ** 2, ['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2']);
add('area', YD ** 2, ['yd2', 'yd^2']);
// Information (bits)
add('information', 1, ['bit'], { prefix: true, binary: true });
add('information', 8, ['byte'], { prefix: true, binary: true });
// Speed (metres per second)
add('speed', 6080 * FT / 3600, ['admkn']);
add('speed', NMI / 3600, ['kn']);
add('speed', 1 / 3600, ['m/h', 'm/hr'], { prefix: true });
add('speed', 1, ['m/s', 'm/sec'], { prefix: true });
add('speed', MI / 3600, ['mph']);

const PREFIXES: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1,
  d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
const BINARY_PREFIXES: Record<string, number> = { ki: 2 ** 10, Mi: 2 ** 20, Gi: 2 ** 30, Ti: 2 ** 40, Pi: 2 ** 50, Ei: 2 ** 60, Zi: 2 ** 70, Yi: 2 ** 80 };

interface Resolved { unit: string; base: Unit; scale: number }
const resolveUnit = (name: string): Resolved | null => {
  const exact = UNITS[name];
  if (exact) return { unit: name, base: exact, scale: 1 };
  for (const len of [2, 1]) {
    const p = name.slice(0, len), rest = name.slice(len);
    const u = UNITS[rest];
    if (!u || !u.prefix) continue;
    const dec = PREFIXES[p], bin = u.binary ? BINARY_PREFIXES[p] : undefined;
    const f = bin ?? dec;
    if (f !== undefined) return { unit: rest, base: u, scale: f ** (u.dim ?? 1) };
  }
  return null;
};

const toKelvin = (x: number, u: string): number => {
  switch (u) {
    case 'C': case 'cel': return x + 273.15;
    case 'F': case 'fah': return ((x - 32) * 5) / 9 + 273.15;
    case 'Rank': return (x * 5) / 9;
    case 'Reau': return x * 1.25 + 273.15;
    default: return x;
  }
};
const fromKelvin = (k: number, u: string): number => {
  switch (u) {
    case 'C': case 'cel': return k - 273.15;
    case 'F': case 'fah': return ((k - 273.15) * 9) / 5 + 32;
    case 'Rank': return (k * 9) / 5;
    case 'Reau': return (k - 273.15) * 0.8;
    default: return k;
  }
};

// ---------------------------------------------------------------- error function

const SQRT_PI = Math.sqrt(Math.PI);
/** erf(x) = 2/√π · e^(−x²) · Σ 2ⁿx^(2n+1)/(1·3·…·(2n+1)); every term is positive, so no cancellation. */
const erfSeries = (x: number): number => {
  let term = x, sum = x;
  for (let n = 1; n < 500; n++) {
    term *= (2 * x * x) / (2 * n + 1);
    sum += term;
    if (term < sum * 1e-17) break;
  }
  return (2 / SQRT_PI) * Math.exp(-x * x) * sum;
};
/** erfc(x) for x ≥ 2 by the Laplace continued fraction, evaluated with the modified Lentz method. */
const erfcFraction = (x: number): number => {
  const tiny = 1e-300;
  let f = x, c = x, d = 0;
  for (let k = 1; k < 5000; k++) {
    const a = k / 2;
    d = x + a * d; d = d === 0 ? tiny : 1 / d;
    c = x + a / c; if (c === 0) c = tiny;
    const delta = c * d;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return Math.exp(-x * x) / (SQRT_PI * f);
};
export const erf = (x: number): number => {
  const a = Math.abs(x);
  const r = a < 3 ? erfSeries(a) : 1 - erfcFraction(a);
  return x < 0 ? -r : r;
};
export const erfc = (x: number): number => {
  if (x < 0) return 2 - erfc(-x);
  return x < 2 ? 1 - erfSeries(x) : erfcFraction(x);
};

const realFn = (name: string, f: (x: number) => number | CellError): FunctionDef => ({
  name, min: 1, max: 1, scalar: true,
  fn: ([v], ctx) => { const x = eNum(v, ctx); return isError(x) ? x : fin(f(x)); },
});

export const defs: FunctionDef[] = [
  bessel('BESSELJ', besselJ, false),
  bessel('BESSELI', (x, n) => besselSeries(x, n, 1), false),
  bessel('BESSELY', besselY, true),
  bessel('BESSELK', besselK, true),

  fromBase('BIN2DEC', BIN, null), fromBase('BIN2HEX', BIN, HEX), fromBase('BIN2OCT', BIN, OCT),
  fromBase('OCT2DEC', OCT, null), fromBase('OCT2BIN', OCT, BIN), fromBase('OCT2HEX', OCT, HEX),
  fromBase('HEX2DEC', HEX, null), fromBase('HEX2BIN', HEX, BIN), fromBase('HEX2OCT', HEX, OCT),
  fromDec('DEC2BIN', BIN), fromDec('DEC2OCT', OCT), fromDec('DEC2HEX', HEX),

  bitOp('BITAND', (a, b) => a & b),
  bitOp('BITOR', (a, b) => a | b),
  bitOp('BITXOR', (a, b) => a ^ b),
  bitShift('BITLSHIFT', 1),
  bitShift('BITRSHIFT', -1),

  {
    name: 'COMPLEX', min: 2, max: 3, scalar: true,
    fn: ([r, i, s], ctx) => {
      const re = eNum(r, ctx); if (isError(re)) return re;
      const im = eNum(i, ctx); if (isError(im)) return im;
      const sv = s === undefined || s === null ? '' : first(s, ctx);
      if (isError(sv)) return sv;
      if (sv !== '' && sv !== 'i' && sv !== 'j') return Err.value;
      return formatComplex(re, im, sv);
    },
  },
  {
    name: 'IMREAL', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const c = cArg(v, ctx); return isError(c) ? c : c.re; },
  },
  {
    name: 'IMAGINARY', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const c = cArg(v, ctx); return isError(c) ? c : c.im; },
  },
  {
    name: 'IMABS', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const c = cArg(v, ctx); return isError(c) ? c : fin(Math.hypot(c.re, c.im)); },
  },
  {
    name: 'IMARGUMENT', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      const c = cArg(v, ctx);
      if (isError(c)) return c;
      return c.re === 0 && c.im === 0 ? Err.div0 : Math.atan2(c.im, c.re);
    },
  },
  imFn('IMCONJUGATE', ([a, b]) => [a, -b]),
  imFn('IMEXP', cexp),
  imFn('IMLN', (c) => (isZero(c) ? null : cln(c))),
  imFn('IMLOG10', (c) => { if (isZero(c)) return null; const [a, b] = cln(c); return [a / Math.LN10, b / Math.LN10]; }),
  imFn('IMLOG2', (c) => { if (isZero(c)) return null; const [a, b] = cln(c); return [a / Math.LN2, b / Math.LN2]; }),
  imFn('IMSQRT', ([a, b]) => {
    const r = Math.sqrt(Math.hypot(a, b)), t = Math.atan2(b, a) / 2;
    return [r * Math.cos(t), r * Math.sin(t)];
  }),
  imFn('IMSIN', csin),
  imFn('IMCOS', ccos),
  imFn('IMTAN', (c) => cdiv(csin(c), ccos(c))),
  imFn('IMCOT', (c) => (isZero(c) ? null : cdiv(ccos(c), csin(c)))),
  imFn('IMSEC', (c) => recip(ccos(c))),
  imFn('IMCSC', (c) => (isZero(c) ? null : recip(csin(c)))),
  imFn('IMSINH', csinh),
  imFn('IMCOSH', ccosh),
  imFn('IMSECH', (c) => recip(ccosh(c))),
  imFn('IMCSCH', (c) => (isZero(c) ? null : recip(csinh(c)))),
  imBinary('IMSUB', ([a, b], [c, d]) => [a - c, b - d]),
  imBinary('IMDIV', (x, y) => (isZero(y) ? null : cdiv(x, y))),
  {
    name: 'IMPOWER', min: 2, max: 2, scalar: true,
    fn: ([v, p], ctx) => {
      const c = cArg(v, ctx); if (isError(c)) return c;
      const n = eNum(p, ctx); if (isError(n)) return n;
      const r = Math.hypot(c.re, c.im);
      if (r === 0) return n > 0 ? '0' : Err.num;
      if (c.im === 0 && Number.isInteger(n)) return formatComplex(c.re ** n, 0, c.unit);
      const t = Math.atan2(c.im, c.re) * n, m = r ** n;
      return formatComplex(m * Math.cos(t), m * Math.sin(t), c.unit);
    },
  },
  imFold('IMSUM', ([a, b], [c, d]) => [a + c, b + d], [0, 0]),
  imFold('IMPRODUCT', cmul, [1, 0]),

  {
    name: 'CONVERT', min: 3, max: 3, scalar: true,
    fn: ([v, f, t], ctx) => {
      const x = eNum(v, ctx); if (isError(x)) return x;
      const fs = first(f, ctx), ts = first(t, ctx);
      if (isError(fs)) return fs;
      if (isError(ts)) return ts;
      if (typeof fs !== 'string' || typeof ts !== 'string') return Err.na;
      const from = resolveUnit(fs), to = resolveUnit(ts);
      if (!from || !to || from.base.cat !== to.base.cat) return Err.na;
      if (from.base.cat === 'temperature') {
        const k = toKelvin(x * (from.unit === 'K' || from.unit === 'kel' ? from.scale : 1), from.unit);
        const r = fromKelvin(k, to.unit) / (to.unit === 'K' || to.unit === 'kel' ? to.scale : 1);
        return fin(snap(r));
      }
      return fin(snap((x * from.base.factor * from.scale) / (to.base.factor * to.scale)));
    },
  },

  {
    name: 'DELTA', min: 1, max: 2, scalar: true,
    fn: ([a, b], ctx) => {
      const x = eNum(a, ctx); if (isError(x)) return x;
      const y = eOpt(b, ctx, 0); if (isError(y)) return y;
      return x === y ? 1 : 0;
    },
  },
  {
    name: 'GESTEP', min: 1, max: 2, scalar: true,
    fn: ([a, b], ctx) => {
      const x = eNum(a, ctx); if (isError(x)) return x;
      const y = eOpt(b, ctx, 0); if (isError(y)) return y;
      return x >= y ? 1 : 0;
    },
  },
  {
    name: 'ERF', min: 1, max: 2, scalar: true,
    fn: ([a, b], ctx) => {
      const lo = eNum(a, ctx); if (isError(lo)) return lo;
      if (b === undefined || b === null) return fin(erf(lo));
      const hi = eNum(b, ctx); if (isError(hi)) return hi;
      return fin(erf(hi) - erf(lo));
    },
  },
  realFn('ERF.PRECISE', erf),
  realFn('ERFC', erfc),
  realFn('ERFC.PRECISE', erfc),
];
