import { CellError, Err, isError, isLambda, isMatrix, isRef, isRefs, type FnContext, type Matrix, type RangeRef, type Scalar, type Value } from './types';
import { parseDateTime } from './date';

/** Parses text the way spreadsheet turns typed text into a number: 1,234.5  -12%  $3  1e3  (5)  and dates. */
export const parseNumber = (text: string, date1904 = false): number | null => {
  let s = text.trim();
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  let pct = 0;
  while (s.endsWith('%')) { pct++; s = s.slice(0, -1).trim(); }
  s = s.replace(/^([+-]?)\s*[$€£¥]\s*/, '$1').replace(/\s*[$€£¥]$/, '');
  if (/^[+-]?(\d{1,3}(,\d{3})+|\d+)?(\.\d*)?([eE][+-]?\d+)?$/.test(s) && /\d/.test(s)) {
    let n = Number(s.replace(/,/g, ''));
    if (!Number.isFinite(n)) return null;
    for (let i = 0; i < pct; i++) n /= 100;
    return neg ? -n : n;
  }
  if (pct || neg) return null;
  return parseDateTime(s, date1904);
};

/** Single value from anything: top-left of an array or range (implicit intersection is not applied). */
export const first = (v: Value, ctx: FnContext): Scalar => {
  if (isMatrix(v)) return v[0]?.[0] ?? null;
  if (isRef(v)) return ctx.cell(v.sheet, v.r1, v.c1);
  if (isRefs(v)) return v.areas[0] ? ctx.cell(v.areas[0].sheet, v.areas[0].r1, v.areas[0].c1) : Err.ref;
  if (isLambda(v)) return Err.calc;
  return v;
};

/** spreadsheet's number coercion for one value. Blank is 0, TRUE is 1, numeric text is read, other text is #VALUE!. */
export const toNumber = (v: Scalar, date1904 = false): number | CellError => {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const n = parseNumber(v, date1904);
  return n === null ? Err.value : n;
};

export const toBool = (v: Scalar): boolean | CellError => {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return v;
  const u = v.trim().toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return Err.value;
};

/** Shortest decimal spreadsheets show for a number in "General" text conversion (15 significant digits). */
export const numberToText = (n: number): string => {
  if (Number.isInteger(n) && Math.abs(n) < 1e21) return String(n);
  if (n === 0) return '0';
  const p = Number(n.toPrecision(15));
  const abs = Math.abs(p);
  if (abs >= 1e21 || abs < 1e-9) return p.toExponential().replace(/\.?0+e/, 'E').replace('E+', 'E+').toUpperCase();
  return String(p);
};

export const toText = (v: Scalar): string | CellError => {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v;
  return numberToText(v);
};

/** Any value as a rectangular matrix. Scalars become 1×1. */
export const toMatrix = (v: Value, ctx: FnContext): Matrix | CellError => {
  if (isMatrix(v)) return v;
  if (isRef(v)) return ctx.values(v);
  if (isRefs(v)) return v.areas.length === 1 ? ctx.values(v.areas[0]!) : Err.value;
  if (isLambda(v)) return Err.calc;
  return [[v]];
};

/** Each cell of every argument, in order, flattened. Union references contribute every area. */
export function* cellsOf(v: Value, ctx: FnContext): Generator<Scalar> {
  if (isRefs(v)) { for (const a of v.areas) yield* cellsOf(a, ctx); return; }
  if (isRef(v) || isMatrix(v)) {
    const m = isRef(v) ? ctx.values(v) : v;
    for (const row of m) for (const c of row) yield c;
    return;
  }
  if (isLambda(v)) { yield Err.calc; return; }
  yield v;
}

export const fromRange = (v: Value): boolean => isRef(v) || isRefs(v) || isMatrix(v);

export interface CollectOptions {
  /** Count TRUE/FALSE found inside ranges and arrays (AVERAGEA-style). */
  boolsInRanges?: boolean;
  /** Count text inside ranges and arrays as 0 (AVERAGEA-style). */
  textInRangesAsZero?: boolean;
  /** Skip error values found in ranges instead of returning them. */
  ignoreErrors?: boolean;
}

/**
 * Numbers for SUM-like functions, following spreadsheet's rules: values typed directly as arguments are
 * coerced (TRUE counts 1, "3" counts 3, "abc" is #VALUE!), while inside ranges and arrays only real
 * numbers count, unless the options say otherwise. The first error found is returned.
 */
export const collectNumbers = (args: readonly Value[], ctx: FnContext, opt: CollectOptions = {}): number[] | CellError => {
  const out: number[] = [];
  for (const a of args) {
    if (fromRange(a)) {
      for (const c of cellsOf(a, ctx)) {
        if (typeof c === 'number') out.push(c);
        else if (isError(c)) { if (!opt.ignoreErrors) return c; }
        else if (typeof c === 'boolean') { if (opt.boolsInRanges) out.push(c ? 1 : 0); }
        else if (typeof c === 'string') { if (opt.textInRangesAsZero) out.push(0); }
      }
    } else {
      const s = isLambda(a) ? Err.calc : (a as Scalar);
      if (s === null) continue;
      const n = toNumber(s, ctx.date1904);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
};

export const num = (v: Value, ctx: FnContext): number | CellError => toNumber(first(v, ctx), ctx.date1904);
export const int = (v: Value, ctx: FnContext): number | CellError => {
  const n = num(v, ctx);
  return isError(n) ? n : Math.trunc(n);
};
export const text = (v: Value, ctx: FnContext): string | CellError => toText(first(v, ctx));
export const bool = (v: Value, ctx: FnContext): boolean | CellError => toBool(first(v, ctx));

/** Optional argument: missing (undefined) or an omitted argument (null from `=F(1,)`) gives the default. */
export const optNum = (v: Value | undefined, ctx: FnContext, dflt: number): number | CellError =>
  v === undefined || v === null ? dflt : num(v, ctx);
export const optBool = (v: Value | undefined, ctx: FnContext, dflt: boolean): boolean | CellError =>
  v === undefined || v === null ? dflt : bool(v, ctx);

/** Builds a result array; an empty result is #CALC!, as spreadsheets do for FILTER and friends. */
export const matrix = (rows: Scalar[][]): Matrix | CellError => (rows.length && rows[0]!.length ? rows : Err.calc);

export const dims = (m: Matrix): [number, number] => [m.length, m[0]?.length ?? 0];

/** Numbers equal to 15 significant digits count as equal, which is how spreadsheet compares. */
export const approxEqual = (a: number, b: number): boolean => a === b || Math.abs(a - b) <= Math.abs(a) * 1e-15;

/** spreadsheet ordering for sorting and comparison: numbers < text < FALSE < TRUE; blanks sort last. */
export const compareValues = (a: Scalar, b: Scalar): number => {
  const rank = (v: Scalar): number => (v === null ? 5 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 4);
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, undefined, { sensitivity: 'accent' }) || 0;
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return 0;
};

export const refArea = (r: RangeRef): [number, number] => [r.r2 - r.r1 + 1, r.c2 - r.c1 + 1];
