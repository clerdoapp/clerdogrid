import { CellError, Err, isError, type FnContext, type FunctionDef, type Matrix, type Scalar, type Value } from '../types';
import { cellsOf, collectNumbers, fromRange, num, text, toMatrix, toNumber } from '../coerce';

/** Final numeric result: NaN/Infinity is #NUM!, and -0 is shown as 0. */
const fin = (x: number | CellError): number | CellError => (isError(x) ? x : Number.isFinite(x) ? (x === 0 ? 0 : x) : Err.num);

/** Rounds to spreadsheet's 15 significant digits, removing binary noise such as 0.1+0.2. */
const snap = (x: number): number => (Number.isFinite(x) ? Number(x.toPrecision(15)) : x);

/** Exact decimal shift n × 10^k through the decimal exponent, without multiplication error. */
const shift = (n: number, k: number): number => {
  const [m, e] = n.toExponential().split('e');
  return Number(`${m}e${Number(e) + k}`);
};

type RoundMode = 'near' | 'down' | 'up';
/** spreadsheets round the 15-digit decimal value, half away from zero: ROUND(2.675,2)=2.68, ROUND(-2.5,0)=-3. */
export const roundTo = (x: number, digits: number, mode: RoundMode): number => {
  if (x === 0 || !Number.isFinite(x)) return x;
  const d = Math.trunc(digits);
  const s = shift(Math.abs(snap(x)), d);
  if (!Number.isFinite(s)) return x;
  const r = mode === 'near' ? Math.round(s) : mode === 'down' ? Math.floor(s) : Math.ceil(s);
  return Math.sign(x) * shift(r, -d);
};

type NumBody = (...xs: number[]) => number | CellError;
/**
 * Element-wise numeric function. Optional arguments (index ≥ min) that are omitted take `defaults`;
 * NaN as a default lets the body tell "omitted" apart from a value.
 */
const numFn = (name: string, min: number, max: number, body: NumBody, defaults: number[] = []): FunctionDef => ({
  name, min, max, scalar: true,
  fn: (args, ctx) => {
    const xs: number[] = [];
    for (let i = 0; i < max; i++) {
      const a = args[i];
      const x = i >= min && (a === undefined || a === null) ? (defaults[i - min] ?? 0) : num(a ?? null, ctx);
      if (isError(x)) return x;
      xs.push(x);
    }
    return fin(body(...xs));
  },
});
const unary = (name: string, body: (x: number) => number | CellError): FunctionDef => numFn(name, 1, 1, body);

// spreadsheet's trigonometric functions refuse arguments of 2^27 or more.
const TRIG_LIMIT = 134217728;
const trig = (name: string, body: (x: number) => number | CellError): FunctionDef =>
  unary(name, (x) => (Math.abs(x) >= TRIG_LIMIT ? Err.num : body(x)));

const fact = (n: number): number => {
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return r;
};

const combin = (n: number, k: number): number | CellError => {
  n = Math.trunc(n); k = Math.trunc(k);
  if (n < 0 || k < 0 || n < k) return Err.num;
  k = Math.min(k, n - k);
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return r < 2 ** 53 ? Math.round(r) : r;
};

/** Multiple-of rounding used by CEILING/FLOOR families; the quotient is snapped so 0.234/0.01 is 23.4. */
const toMultiple = (x: number, sig: number, f: (q: number) => number): number => snap(f(snap(x / sig)) * sig);

/**
 * Integer arguments for GCD/LCM/MULTINOMIAL: numbers are truncated; text inside ranges and logical
 * values anywhere are #VALUE!; blanks are skipped.
 */
const integers = (args: readonly Value[], ctx: FnContext): number[] | CellError => {
  const out: number[] = [];
  for (const a of args) {
    const inRange = fromRange(a);
    for (const c of cellsOf(a, ctx)) {
      if (c === null) continue;
      if (isError(c)) return c;
      if (typeof c === 'boolean' || (inRange && typeof c === 'string')) return Err.value;
      const n = toNumber(c, ctx.date1904);
      if (isError(n)) return n;
      out.push(Math.trunc(n));
    }
  }
  return out;
};

const gcd2 = (a: number, b: number): number => {
  while (b) [a, b] = [b, a % b];
  return a;
};

/** A matrix argument made only of numbers (MDETERM, MINVERSE, MMULT); anything else is #VALUE!. */
const numericMatrix = (v: Value, ctx: FnContext): number[][] | CellError => {
  const m = toMatrix(v, ctx);
  if (isError(m)) return m;
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

const squareMatrix = (v: Value, ctx: FnContext): number[][] | CellError => {
  const m = numericMatrix(v, ctx);
  if (isError(m)) return m;
  return m.length && m.every((r) => r.length === m.length) ? m : Err.value;
};

const determinant = (src: number[][]): number => {
  const a = src.map((r) => [...r]);
  const n = a.length;
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return 0;
    if (p !== c) { [a[p], a[c]] = [a[c], a[p]]; det = -det; }
    det *= a[c][c];
    for (let r = c + 1; r < n; r++) {
      const f = a[r][c] / a[c][c];
      for (let k = c; k < n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return det;
};

const inverse = (src: number[][]): number[][] | CellError => {
  const n = src.length;
  const a = src.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return Err.num;
    [a[p], a[c]] = [a[c], a[p]];
    const pv = a[c][c];
    for (let k = 0; k < 2 * n; k++) a[c][k] /= pv;
    for (let r = 0; r < n; r++) {
      if (r === c || a[r][c] === 0) continue;
      const f = a[r][c];
      for (let k = 0; k < 2 * n; k++) a[r][k] -= f * a[c][k];
    }
  }
  return a.map((r) => r.slice(n).map((x) => (x === 0 ? 0 : x)));
};

/** Pairs of numbers from two equally sized arrays (SUMX2MY2 and friends); non-numeric pairs are skipped. */
const pairSum = (name: string, f: (x: number, y: number) => number): FunctionDef => ({
  name, min: 2, max: 2,
  fn: ([a, b], ctx) => {
    const xs = [...cellsOf(a, ctx)], ys = [...cellsOf(b, ctx)];
    for (const c of [...xs, ...ys]) if (isError(c)) return c;
    if (xs.length !== ys.length) return Err.na;
    let s = 0;
    xs.forEach((x, i) => { const y = ys[i]; if (typeof x === 'number' && typeof y === 'number') s += f(x, y); });
    return fin(s);
  },
});

const aggregate = (name: string, f: (xs: number[]) => number): FunctionDef => ({
  name, min: 1, max: 255,
  fn: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    return isError(xs) ? xs : fin(f(xs));
  },
});

// ROMAN's concise forms (1–4) follow the classic algorithm used by spreadsheet apps.
const ROMAN_CHARS = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];
const roman = (value: number, mode: number): string => {
  let out = '', v = value;
  const maxIdx = ROMAN_VALUES.length - 1;
  for (let i = 0; i <= maxIdx / 2; i++) {
    let idx = 2 * i;
    const digit = Math.floor(v / ROMAN_VALUES[idx]);
    if (digit % 5 === 4) {
      const idx2 = digit === 4 ? idx - 1 : idx - 2;
      let steps = 0;
      while (steps < mode && idx < maxIdx) {
        steps++;
        if (ROMAN_VALUES[idx2] - ROMAN_VALUES[idx + 1] <= v) idx++;
        else steps = mode;
      }
      out += ROMAN_CHARS[idx] + ROMAN_CHARS[idx2];
      v = v + ROMAN_VALUES[idx] - ROMAN_VALUES[idx2];
    } else {
      if (digit > 4) out += ROMAN_CHARS[idx - 1];
      out += ROMAN_CHARS[idx].repeat(digit % 5);
      v %= ROMAN_VALUES[idx];
    }
  }
  return out;
};

const ROMAN_DIGITS: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };

const scalarOf = (v: Value | undefined): Scalar => (v === undefined ? null : (v as Scalar));

export const defs: FunctionDef[] = [
  unary('ABS', Math.abs),
  unary('ACOS', Math.acos),
  unary('ACOSH', Math.acosh),
  unary('ACOT', (x) => Math.PI / 2 - Math.atan(x)),
  unary('ACOTH', (x) => (Math.abs(x) <= 1 ? Err.num : 0.5 * Math.log((x + 1) / (x - 1)))),
  unary('ASIN', Math.asin),
  unary('ASINH', Math.asinh),
  unary('ATAN', Math.atan),
  numFn('ATAN2', 2, 2, (x, y) => (x === 0 && y === 0 ? Err.div0 : Math.atan2(y, x))),
  unary('ATANH', (x) => (Math.abs(x) >= 1 ? Err.num : Math.atanh(x))),
  trig('COS', Math.cos),
  unary('COSH', Math.cosh),
  trig('COT', (x) => (x === 0 ? Err.div0 : 1 / Math.tan(x))),
  trig('COTH', (x) => (x === 0 ? Err.div0 : 1 / Math.tanh(x))),
  trig('CSC', (x) => (x === 0 ? Err.div0 : 1 / Math.sin(x))),
  trig('CSCH', (x) => (x === 0 ? Err.div0 : 1 / Math.sinh(x))),
  trig('SEC', (x) => 1 / Math.cos(x)),
  trig('SECH', (x) => 1 / Math.cosh(x)),
  trig('SIN', Math.sin),
  unary('SINH', Math.sinh),
  trig('TAN', Math.tan),
  unary('TANH', Math.tanh),
  unary('DEGREES', (x) => (x * 180) / Math.PI),
  unary('RADIANS', (x) => (x * Math.PI) / 180),
  unary('EXP', Math.exp),
  unary('LN', (x) => (x <= 0 ? Err.num : Math.log(x))),
  unary('LOG10', (x) => (x <= 0 ? Err.num : Math.log10(x))),
  numFn('LOG', 1, 2, (x, b) => {
    if (x <= 0 || b <= 0) return Err.num;
    if (b === 1) return Err.div0;
    return b === 10 ? Math.log10(x) : b === 2 ? Math.log2(x) : Math.log(x) / Math.log(b);
  }, [10]),
  unary('SQRT', (x) => (x < 0 ? Err.num : Math.sqrt(x))),
  unary('SQRTPI', (x) => (x < 0 ? Err.num : Math.sqrt(x * Math.PI))),
  unary('SIGN', Math.sign),
  unary('INT', Math.floor),
  numFn('TRUNC', 1, 2, (x, d) => roundTo(x, d, 'down'), [0]),
  numFn('ROUND', 2, 2, (x, d) => roundTo(x, d, 'near')),
  numFn('ROUNDDOWN', 2, 2, (x, d) => roundTo(x, d, 'down')),
  numFn('ROUNDUP', 2, 2, (x, d) => roundTo(x, d, 'up')),
  numFn('MROUND', 2, 2, (x, m) => {
    if (m === 0 || x === 0) return 0;
    if (Math.sign(x) !== Math.sign(m)) return Err.num;
    return snap(roundTo(snap(x / m), 0, 'near') * m);
  }),
  unary('EVEN', (x) => {
    const v = Math.ceil(Math.abs(snap(x)));
    return (x < 0 ? -1 : 1) * (v % 2 ? v + 1 : v);
  }),
  unary('ODD', (x) => {
    const v = Math.ceil(Math.abs(snap(x)));
    return (x < 0 ? -1 : 1) * (v % 2 ? v : v + 1);
  }),
  // CEILING/FLOOR (current rules): a negative number with a positive significance rounds toward zero
  // for CEILING and away from zero for FLOOR; a positive number with a negative significance is #NUM!.
  numFn('CEILING', 2, 2, (x, s) => {
    if (x === 0 || s === 0) return 0;
    if (x > 0 && s < 0) return Err.num;
    return toMultiple(x, s, Math.ceil);
  }),
  numFn('FLOOR', 2, 2, (x, s) => {
    if (x === 0) return 0;
    if (s === 0) return Err.div0;
    if (x > 0 && s < 0) return Err.num;
    return toMultiple(x, s, Math.floor);
  }),
  numFn('CEILING.MATH', 1, 3, (x, s, mode) => {
    s = Math.abs(s);
    if (s === 0 || x === 0) return 0;
    return toMultiple(x, s, x < 0 && mode !== 0 ? Math.floor : Math.ceil);
  }, [1, 0]),
  numFn('FLOOR.MATH', 1, 3, (x, s, mode) => {
    s = Math.abs(s);
    if (s === 0 || x === 0) return 0;
    return toMultiple(x, s, x < 0 && mode !== 0 ? Math.ceil : Math.floor);
  }, [1, 0]),
  numFn('CEILING.PRECISE', 1, 2, (x, s) => (s === 0 || x === 0 ? 0 : toMultiple(x, Math.abs(s), Math.ceil)), [1]),
  numFn('ISO.CEILING', 1, 2, (x, s) => (s === 0 || x === 0 ? 0 : toMultiple(x, Math.abs(s), Math.ceil)), [1]),
  numFn('FLOOR.PRECISE', 1, 2, (x, s) => (s === 0 || x === 0 ? 0 : toMultiple(x, Math.abs(s), Math.floor)), [1]),
  numFn('MOD', 2, 2, (n, d) => {
    if (d === 0) return Err.div0;
    const q = snap(n / d); // MOD(0.3, 0.1) is 0, not 0.1
    return Number.isInteger(q) ? 0 : snap(n - d * Math.floor(q));
  }),
  numFn('QUOTIENT', 2, 2, (n, d) => (d === 0 ? Err.div0 : Math.trunc(n / d))),
  numFn('POWER', 2, 2, (b, e) => {
    if (b === 0 && e === 0) return Err.num;
    if (b === 0 && e < 0) return Err.div0;
    if (b < 0 && !Number.isInteger(e)) return Err.num;
    return Math.pow(b, e);
  }),
  unary('FACT', (n) => (n < 0 ? Err.num : fact(Math.trunc(n)))),
  unary('FACTDOUBLE', (n) => {
    n = Math.trunc(n);
    if (n < -1) return Err.num;
    let r = 1;
    for (let i = n; i > 1; i -= 2) r *= i;
    return r;
  }),
  numFn('COMBIN', 2, 2, combin),
  numFn('COMBINA', 2, 2, (n, k) => {
    n = Math.trunc(n); k = Math.trunc(k);
    if (n < 0 || k < 0) return Err.num;
    return n === 0 && k === 0 ? 1 : combin(n + k - 1, k);
  }),
  { name: 'PI', min: 0, max: 0, fn: () => Math.PI },
  { name: 'RAND', min: 0, max: 0, volatile: true, fn: () => Math.random() },
  {
    name: 'RANDBETWEEN', min: 2, max: 2, volatile: true, scalar: true,
    fn: ([a, b], ctx) => {
      const lo = num(a, ctx);
      if (isError(lo)) return lo;
      const hi = num(b, ctx);
      if (isError(hi)) return hi;
      const l = Math.ceil(lo), h = Math.floor(hi);
      return l > h ? Err.num : l + Math.floor(Math.random() * (h - l + 1));
    },
  },
  {
    name: 'GCD', min: 1, max: 255,
    fn: (args, ctx) => {
      const xs = integers(args, ctx);
      if (isError(xs)) return xs;
      if (xs.some((x) => x < 0 || x >= 2 ** 53)) return Err.num;
      return xs.reduce(gcd2, 0);
    },
  },
  {
    name: 'LCM', min: 1, max: 255,
    fn: (args, ctx) => {
      const xs = integers(args, ctx);
      if (isError(xs)) return xs;
      if (xs.some((x) => x < 0 || x >= 2 ** 53)) return Err.num;
      if (xs.some((x) => x === 0)) return 0;
      return fin(xs.reduce((a, b) => (a / gcd2(a, b)) * b, 1));
    },
  },
  {
    name: 'MULTINOMIAL', min: 1, max: 255,
    fn: (args, ctx) => {
      const xs = integers(args, ctx);
      if (isError(xs)) return xs;
      if (xs.some((x) => x < 0)) return Err.num;
      // Product of binomials C(a1+…+ai, ai) avoids the overflow of (Σa)!/Πa!.
      let total = 0, r = 1;
      for (const x of xs) {
        total += x;
        const c = combin(total, x);
        if (isError(c)) return c;
        r *= c;
      }
      return fin(r);
    },
  },
  aggregate('SUM', (xs) => xs.reduce((a, b) => a + b, 0)),
  aggregate('SUMSQ', (xs) => xs.reduce((a, b) => a + b * b, 0)),
  aggregate('PRODUCT', (xs) => (xs.length ? xs.reduce((a, b) => a * b, 1) : 0)),
  pairSum('SUMX2MY2', (x, y) => x * x - y * y),
  pairSum('SUMX2PY2', (x, y) => x * x + y * y),
  pairSum('SUMXMY2', (x, y) => (x - y) ** 2),
  {
    name: 'SERIESSUM', min: 4, max: 4,
    fn: ([xv, nv, mv, cv], ctx) => {
      const x = num(xv, ctx); if (isError(x)) return x;
      const n = num(nv, ctx); if (isError(n)) return n;
      const m = num(mv, ctx); if (isError(m)) return m;
      let s = 0, i = 0;
      for (const c of cellsOf(cv, ctx)) {
        if (isError(c)) return c;
        if (typeof c !== 'number') return Err.value;
        s += c * Math.pow(x, n + i++ * m);
      }
      return fin(s);
    },
  },
  {
    name: 'MDETERM', min: 1, max: 1,
    fn: ([a], ctx) => { const m = squareMatrix(a, ctx); return isError(m) ? m : fin(snap(determinant(m))); },
  },
  {
    name: 'MINVERSE', min: 1, max: 1,
    fn: ([a], ctx) => { const m = squareMatrix(a, ctx); return isError(m) ? m : inverse(m); },
  },
  {
    name: 'MMULT', min: 2, max: 2,
    fn: ([av, bv], ctx) => {
      const a = numericMatrix(av, ctx); if (isError(a)) return a;
      const b = numericMatrix(bv, ctx); if (isError(b)) return b;
      if (!a.length || !b.length || a[0].length !== b.length) return Err.value;
      const out: Matrix = a.map((row) => b[0].map((_, j) => {
        let s = 0;
        for (let k = 0; k < row.length; k++) s += row[k] * b[k][j];
        return fin(s);
      }));
      return out;
    },
  },
  {
    name: 'MUNIT', min: 1, max: 1,
    fn: ([d], ctx) => {
      const n = num(d, ctx);
      if (isError(n)) return n;
      const k = Math.trunc(n);
      if (k < 1) return Err.value;
      return Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => (i === j ? 1 : 0)));
    },
  },
  {
    name: 'ARABIC', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      const t = text(v, ctx);
      if (isError(t)) return t;
      let s = t.trim().toUpperCase();
      if (s.length > 255) return Err.value;
      const neg = s.startsWith('-');
      if (neg) s = s.slice(1);
      let total = 0;
      for (let i = 0; i < s.length; i++) {
        const cur = ROMAN_DIGITS[s[i]], next = ROMAN_DIGITS[s[i + 1] ?? ''] ?? 0;
        if (cur === undefined) return Err.value;
        total += cur < next ? -cur : cur;
      }
      return neg ? -total : total;
    },
  },
  {
    name: 'ROMAN', min: 1, max: 2, scalar: true,
    fn: ([v, f], ctx) => {
      const n = num(v, ctx);
      if (isError(n)) return n;
      const form = scalarOf(f);
      let mode = 0;
      if (typeof form === 'boolean') mode = form ? 0 : 4;
      else if (form !== null) {
        const m = num(form, ctx);
        if (isError(m)) return m;
        mode = Math.trunc(m);
      }
      const k = Math.trunc(n);
      if (k < 0 || k > 3999 || mode < 0 || mode > 4) return Err.value;
      return roman(k, mode);
    },
  },
  {
    name: 'BASE', min: 2, max: 3, scalar: true,
    fn: ([v, r, l], ctx) => {
      const n = num(v, ctx); if (isError(n)) return n;
      const radix = num(r, ctx); if (isError(radix)) return radix;
      const len = l === undefined || l === null ? 0 : num(l, ctx);
      if (isError(len)) return len;
      const x = Math.trunc(n), rd = Math.trunc(radix), ml = Math.trunc(len);
      if (x < 0 || x >= 2 ** 53 || rd < 2 || rd > 36 || ml < 0 || ml > 255) return Err.num;
      return x.toString(rd).toUpperCase().padStart(ml, '0');
    },
  },
  {
    name: 'DECIMAL', min: 2, max: 2, scalar: true,
    fn: ([v, r], ctx) => {
      const t = text(v, ctx); if (isError(t)) return t;
      const radix = num(r, ctx); if (isError(radix)) return radix;
      const rd = Math.trunc(radix);
      if (rd < 2 || rd > 36 || t.length > 255) return Err.num;
      let out = 0;
      for (const ch of t.toLowerCase()) {
        const d = parseInt(ch, 36);
        if (Number.isNaN(d) || d >= rd) return Err.num;
        out = out * rd + d;
      }
      return out >= 2 ** 53 ? Err.num : out;
    },
  },
];
