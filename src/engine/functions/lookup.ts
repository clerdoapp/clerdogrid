import {
  CellError, Err, MAX_COLS, MAX_ROWS, isError, isLambda, isMatrix, isRef, isRefs, isScalar, ref,
  type FnContext, type FunctionDef, type LambdaValue, type Matrix, type RangeRef, type Scalar, type Value,
} from '../types';
import { approxEqual, cellsOf, compareValues, first, int, matrix, num, optBool, optNum, refArea, text, toMatrix } from '../coerce';
import { textCompare, textMatcher } from './criteria';

type R<T> = T | CellError;

// ---------------------------------------------------------------- shared helpers

/**
 * Values of an array argument at its real shape. References are padded with blanks to their full
 * size, except whole-column/whole-row references, which stay trimmed to the used area.
 */
export const arr = (v: Value, ctx: FnContext): R<Matrix> => {
  if (!isRef(v)) return toMatrix(v, ctx);
  const m = ctx.values(v);
  const [R0, C0] = refArea(v);
  const rows = v.r1 === 0 && v.r2 === MAX_ROWS - 1 ? m.length : R0;
  const cols = v.c1 === 0 && v.c2 === MAX_COLS - 1 ? (m[0]?.length ?? 0) : C0;
  if (m.length === rows && (m[0]?.length ?? 0) === cols) return m;
  return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => m[i]?.[j] ?? null));
};

const shape = (v: Value, ctx: FnContext): R<[number, number]> => {
  if (isRef(v)) return refArea(v);
  if (isRefs(v)) return Err.ref;
  if (isError(v)) return v;
  const m = toMatrix(v, ctx);
  return isError(m) ? m : [m.length, m[0]?.length ?? 0];
};

const transpose = (m: Matrix): Matrix => (m[0] ?? []).map((_, j) => m.map((row) => row[j] ?? null));
const flat = (m: Matrix): Scalar[] => m.flat();
const sub = (r: RangeRef, i1: number, j1: number, i2: number, j2: number): RangeRef =>
  ref(r.sheet, r.r1 + i1, r.c1 + j1, r.r1 + i2, r.c1 + j2);
const scalarOf = (v: Value, ctx: FnContext): Scalar => (isScalar(v) ? v : first(v, ctx));
const nums = (v: Value, ctx: FnContext): R<number[]> => {
  const out: number[] = [];
  for (const c of cellsOf(v, ctx)) {
    const n = num(c, ctx);
    if (isError(n)) return n;
    out.push(Math.trunc(n));
  }
  return out;
};

/** One scalar from a LAMBDA result; arrays are #CALC! (a nested array). */
const one = (v: Value, ctx: FnContext): Scalar => {
  if (isRef(v)) return v.r1 === v.r2 && v.c1 === v.c2 ? ctx.cell(v.sheet, v.r1, v.c1) : Err.calc;
  if (isMatrix(v)) return v.length === 1 && v[0]!.length === 1 ? v[0]![0]! : Err.calc;
  return isScalar(v) ? v : Err.calc;
};

const lam = (v: Value | undefined, n: number): R<LambdaValue> =>
  v !== undefined && isLambda(v) ? (v.params.length && v.params.length !== n ? Err.value : v) : Err.value;

// ---------------------------------------------------------------- lookup comparisons

const typeTag = (v: Scalar): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null ? 3 : 4);
const sameType = (a: Scalar, b: Scalar): boolean => typeTag(a) === typeTag(b) && typeTag(a) < 3;

/** Compares two same-typed values (numbers to 15 digits, text case-insensitively). */
const cmp = (a: Scalar, b: Scalar): number => {
  if (typeof a === 'number' && typeof b === 'number') return approxEqual(a, b) ? 0 : a - b;
  if (typeof a === 'string' && typeof b === 'string') return textCompare(a, b);
  return Number(a) - Number(b);
};

const exactMatcher = (v: Scalar, wildcards: boolean): ((c: Scalar) => boolean) => {
  if (typeof v === 'string') {
    const t = textMatcher(v, wildcards);
    return (c) => typeof c === 'string' && t(c);
  }
  if (v === null) return (c) => c === null;
  return (c) => sameType(c, v) && cmp(c, v) === 0;
};

/**
 * spreadsheet's binary search for approximate MATCH/VLOOKUP/LOOKUP: the last position whose value is
 * <= v (or >= v when `desc`). Cells of another type are skipped by stepping back to the nearest
 * same-typed cell, so unsorted data gives the same "wrong" answers Spreadsheets give.
 */
const approxSearch = (cells: readonly Scalar[], v: Scalar, desc: boolean): number => {
  let lo = 0, hi = cells.length - 1, found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let j = mid;
    while (j >= lo && !sameType(cells[j]!, v)) j--;
    if (j < lo) { lo = mid + 1; continue; }
    const c = cmp(cells[j]!, v);
    if (desc ? c >= 0 : c <= 0) { found = j; lo = mid + 1; } else hi = j - 1;
  }
  return found;
};

const lookupKey = (v: Scalar): Scalar => (v === null ? 0 : v);

/** 1-D view of a lookup vector: its cells plus the orientation. */
const vector = (v: Value, ctx: FnContext): R<{ cells: Scalar[]; vertical: boolean; len: number }> => {
  const m = arr(v, ctx);
  if (isError(m)) return m;
  const [R0, C0] = isRef(v) ? refArea(v) : [m.length, m[0]?.length ?? 0];
  if (R0 !== 1 && C0 !== 1) return Err.na;
  const vertical = !(R0 === 1 && C0 !== 1);
  return { cells: vertical ? m.map((r) => r[0] ?? null) : (m[0] ?? []), vertical, len: vertical ? R0 : C0 };
};

const matchIndex = (cells: Scalar[], v: Scalar, type: number): number => {
  if (type === 0) return cells.findIndex(exactMatcher(v, true));
  return approxSearch(cells, v, type < 0);
};

// ---------------------------------------------------------------- XLOOKUP / XMATCH core

const xsearch = (cells: Scalar[], v: Scalar, mode: number, search: number): R<number> => {
  if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search)) return Err.value;
  if (Math.abs(search) === 2) {
    if (mode === 2) return Err.value;
    const desc = search === -2;
    const i = approxSearch(cells, v, desc);
    if (i >= 0 && cmp(cells[i]!, v) === 0) return i;
    const nextAfter = (k: number): number => { for (let j = k + 1; j < cells.length; j++) if (sameType(cells[j]!, v)) return j; return -1; };
    if (mode === 0) return -1;
    // ascending: i is the next smaller; descending: i is the next larger.
    return (mode === -1) !== desc ? i : nextAfter(i);
  }
  const order = search === 1 ? cells.map((_, i) => i) : cells.map((_, i) => cells.length - 1 - i);
  const eq = exactMatcher(v, mode === 2);
  for (const i of order) if (eq(cells[i]!)) return i;
  if (mode === 0 || mode === 2) return -1;
  let best = -1;
  for (const i of order) {
    const c = cells[i]!;
    if (!sameType(c, v)) continue;
    const d = cmp(c, v);
    if (mode === -1 ? d < 0 && (best < 0 || cmp(c, cells[best]!) > 0) : d > 0 && (best < 0 || cmp(c, cells[best]!) < 0)) best = i;
  }
  return best;
};

// ---------------------------------------------------------------- simple reference functions

const colName = (c: number): string => {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

const quoteSheet = (s: string): string =>
  /^[A-Za-z_À-￿][\w.À-￿]*$/.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) && !/^R\d*C\d*$/i.test(s) ? s : `'${s.replace(/'/g, "''")}'`;

const ADDRESS: FunctionDef = {
  name: 'ADDRESS', min: 2, max: 5, scalar: true,
  fn: (a, ctx) => {
    const r = int(a[0]!, ctx), c = int(a[1]!, ctx), abs = optNum(a[2], ctx, 1), a1 = optBool(a[3], ctx, true);
    if (isError(r)) return r; if (isError(c)) return c; if (isError(abs)) return abs; if (isError(a1)) return a1;
    const t = Math.trunc(abs);
    if (t < 1 || t > 4) return Err.value;
    const absR = t === 1 || t === 2, absC = t === 1 || t === 3;
    let addr: string;
    if (a1) {
      if (r < 1 || r > MAX_ROWS || c < 1 || c > MAX_COLS) return Err.value;
      addr = `${absC ? '$' : ''}${colName(c)}${absR ? '$' : ''}${r}`;
    } else {
      if ((absR && r < 1) || (absC && c < 1) || Math.abs(r) > MAX_ROWS || Math.abs(c) > MAX_COLS) return Err.value;
      const part = (abs: boolean, n: number): string => (abs ? String(n) : n ? `[${n}]` : '');
      addr = `R${part(absR, r)}C${part(absC, c)}`;
    }
    if (a[4] === undefined || a[4] === null) return addr;
    const s = text(a[4], ctx);
    if (isError(s)) return s;
    return `${s ? quoteSheet(s) : ''}!${addr}`;
  },
};

const AREAS: FunctionDef = {
  name: 'AREAS', min: 1, max: 1,
  fn: ([v]) => (isRef(v!) ? 1 : isRefs(v!) ? v.areas.length : Err.value),
};

const rowOrCol = (isRow: boolean): FunctionDef => ({
  name: isRow ? 'ROW' : 'COLUMN', min: 0, max: 1,
  fn: (a, ctx) => {
    const v = a[0];
    if (v === undefined || v === null) return (isRow ? ctx.row : ctx.col) + 1;
    if (isRefs(v)) return Err.ref;
    if (!isRef(v)) return Err.value;
    const lo = isRow ? v.r1 : v.c1, hi = isRow ? v.r2 : v.c2;
    if (lo === hi) return lo + 1;
    const seq = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i + 1);
    return isRow ? seq.map((n) => [n]) : [seq];
  },
});

const ROWS: FunctionDef = { name: 'ROWS', min: 1, max: 1, fn: ([v], ctx) => { const s = shape(v!, ctx); return isError(s) ? s : s[0]; } };
const COLUMNS: FunctionDef = { name: 'COLUMNS', min: 1, max: 1, fn: ([v], ctx) => { const s = shape(v!, ctx); return isError(s) ? s : s[1]; } };

const HYPERLINK: FunctionDef = {
  name: 'HYPERLINK', min: 1, max: 2, scalar: true,
  fn: (a, ctx) => (a[1] === undefined || a[1] === null ? first(a[0]!, ctx) : first(a[1], ctx)),
};

// ---------------------------------------------------------------- INDEX / OFFSET

const INDEX: FunctionDef = {
  name: 'INDEX', min: 1, max: 4, scalar: [1, 2, 3],
  fn: (a, ctx) => {
    let src = a[0]!;
    const areaN = optNum(a[3], ctx, 1);
    if (isError(areaN)) return areaN;
    const an = Math.trunc(areaN);
    if (isRefs(src)) {
      if (an < 1 || an > src.areas.length) return Err.ref;
      src = src.areas[an - 1]!;
    } else if (an !== 1) return Err.ref;
    if (isLambda(src)) return Err.value;
    const m = isRef(src) ? null : toMatrix(src, ctx);
    if (isError(m)) return m;
    const [R0, C0] = isRef(src) ? refArea(src) : [m!.length, m![0]!.length];
    let row = optNum(a[1], ctx, 0), col = optNum(a[2], ctx, 0);
    if (isError(row)) return row; if (isError(col)) return col;
    row = Math.trunc(row); col = Math.trunc(col);
    // With only row_num given, a one-row array is indexed along its columns.
    if (a[2] === undefined && R0 === 1 && C0 > 1) { col = row; row = 0; }
    if (row < 0 || col < 0) return Err.value;
    if (row > R0 || col > C0) return Err.ref;
    if (R0 === 1 && row === 0 && col > 0) row = 1;
    if (C0 === 1 && col === 0 && row > 0) col = 1;
    const i1 = row ? row - 1 : 0, i2 = row ? row - 1 : R0 - 1, j1 = col ? col - 1 : 0, j2 = col ? col - 1 : C0 - 1;
    if (isRef(src)) return sub(src, i1, j1, i2, j2);
    const out = m!.slice(i1, i2 + 1).map((r) => r.slice(j1, j2 + 1));
    return out.length === 1 && out[0]!.length === 1 ? out[0]![0]! : out;
  },
};

const OFFSET: FunctionDef = {
  name: 'OFFSET', min: 3, max: 5, volatile: true,
  fn: (a, ctx) => {
    const base = a[0]!;
    if (!isRef(base)) return Err.value;
    const vals = [a[1], a[2], a[3], a[4]].map((v, k) => (v === undefined || v === null ? (k === 2 ? base.r2 - base.r1 + 1 : k === 3 ? base.c2 - base.c1 + 1 : 0) : num(v, ctx)));
    for (const v of vals) if (isError(v)) return v;
    const [dr, dc, h, w] = (vals as number[]).map(Math.trunc) as [number, number, number, number];
    if (h === 0 || w === 0) return Err.ref;
    const r1 = base.r1 + dr, c1 = base.c1 + dc;
    // A negative height/width extends up/left from the offset cell.
    const r2 = h > 0 ? r1 + h - 1 : r1 + h + 1, c2 = w > 0 ? c1 + w - 1 : c1 + w + 1;
    if (Math.min(r1, r2) < 0 || Math.min(c1, c2) < 0 || Math.max(r1, r2) >= MAX_ROWS || Math.max(c1, c2) >= MAX_COLS) return Err.ref;
    return ref(base.sheet, r1, c1, r2, c2);
  },
};

// ---------------------------------------------------------------- MATCH family

const MATCH: FunctionDef = {
  name: 'MATCH', min: 2, max: 3, scalar: [0],
  fn: (a, ctx) => {
    const v = lookupKey(scalarOf(a[0]!, ctx));
    if (isError(v)) return v;
    const t = optNum(a[2], ctx, 1);
    if (isError(t)) return t;
    const vec = vector(a[1]!, ctx);
    if (isError(vec)) return vec;
    const i = matchIndex(vec.cells, v, Math.sign(t));
    return i < 0 ? Err.na : i + 1;
  },
};

const vhlookup = (vertical: boolean): FunctionDef => ({
  name: vertical ? 'VLOOKUP' : 'HLOOKUP', min: 3, max: 4, scalar: [0],
  fn: (a, ctx) => {
    const v = lookupKey(scalarOf(a[0]!, ctx));
    if (isError(v)) return v;
    const idx = int(a[2]!, ctx), approx = optBool(a[3], ctx, true);
    if (isError(idx)) return idx; if (isError(approx)) return approx;
    const t = a[1]!;
    if (isRefs(t) || isLambda(t)) return Err.value;
    const m0 = isRef(t) ? ctx.values(t) : toMatrix(t, ctx);
    if (isError(m0)) return m0;
    const m = vertical ? m0 : transpose(m0);
    const width = isRef(t) ? refArea(t)[vertical ? 1 : 0] : m[0]?.length ?? 0;
    if (idx < 1) return Err.value;
    if (idx > width) return Err.ref;
    const keys = m.map((r) => r[0] ?? null);
    const i = approx ? approxSearch(keys, v, false) : keys.findIndex(exactMatcher(v, true));
    return i < 0 ? Err.na : (m[i]![idx - 1] ?? null);
  },
});

const LOOKUP: FunctionDef = {
  name: 'LOOKUP', min: 2, max: 3, scalar: [0],
  fn: (a, ctx) => {
    const v = lookupKey(scalarOf(a[0]!, ctx));
    if (isError(v)) return v;
    const lv = a[1]!;
    if (a[2] === undefined) {
      const m = arr(lv, ctx);
      if (isError(m)) return m;
      const R0 = m.length, C0 = m[0]!.length;
      const byRow = C0 > R0;
      const keys = byRow ? m[0]! : m.map((r) => r[0]!);
      const i = approxSearch(keys, v, false);
      if (i < 0) return Err.na;
      return byRow ? m[R0 - 1]![i]! : m[i]![C0 - 1]!;
    }
    const vec = vector(lv, ctx);
    if (isError(vec)) return vec;
    const i = approxSearch(vec.cells, v, false);
    if (i < 0) return Err.na;
    const rv = a[2];
    if (isRef(rv)) {
      // A result reference extends past its end in its own direction (or the lookup's, for one cell).
      const [R0, C0] = refArea(rv);
      const vert = R0 > 1 ? true : C0 > 1 ? false : vec.vertical;
      const r = vert ? rv.r1 + i : rv.r1, c = vert ? rv.c1 : rv.c1 + i;
      return r >= MAX_ROWS || c >= MAX_COLS ? Err.ref : ctx.cell(rv.sheet, r, c);
    }
    const rm = arr(rv, ctx);
    if (isError(rm)) return rm;
    if (rm.length !== 1 && rm[0]!.length !== 1) return Err.na;
    const cells = flat(rm);
    return i < cells.length ? cells[i]! : Err.na;
  },
};

const XLOOKUP: FunctionDef = {
  name: 'XLOOKUP', min: 3, max: 6, scalar: [0],
  fn: (a, ctx) => {
    const v = scalarOf(a[0]!, ctx);
    if (isError(v)) return v;
    const mode = optNum(a[4], ctx, 0), search = optNum(a[5], ctx, 1);
    if (isError(mode)) return mode; if (isError(search)) return search;
    const vec = vector(a[1]!, ctx);
    if (isError(vec)) return vec.code === '#N/A' ? Err.value : vec;
    const ret = a[2]!;
    const rs = shape(ret, ctx);
    if (isError(rs)) return rs;
    if ((vec.vertical ? rs[0] : rs[1]) !== vec.len) return Err.value;
    const i = xsearch(vec.cells, v, Math.trunc(mode), Math.trunc(search));
    if (isError(i)) return i;
    if (i < 0) return a[3] === undefined || a[3] === null ? Err.na : a[3];
    if (isRef(ret)) return vec.vertical ? sub(ret, i, 0, i, rs[1] - 1) : sub(ret, 0, i, rs[0] - 1, i);
    const m = arr(ret, ctx);
    if (isError(m)) return m;
    const out = vec.vertical ? [m[i]!] : m.map((r) => [r[i]!]);
    return out.length === 1 && out[0]!.length === 1 ? out[0]![0]! : out;
  },
};

const XMATCH: FunctionDef = {
  name: 'XMATCH', min: 2, max: 4, scalar: [0],
  fn: (a, ctx) => {
    const v = scalarOf(a[0]!, ctx);
    if (isError(v)) return v;
    const mode = optNum(a[2], ctx, 0), search = optNum(a[3], ctx, 1);
    if (isError(mode)) return mode; if (isError(search)) return search;
    const vec = vector(a[1]!, ctx);
    if (isError(vec)) return vec.code === '#N/A' ? Err.value : vec;
    const i = xsearch(vec.cells, v, Math.trunc(mode), Math.trunc(search));
    return isError(i) ? i : i < 0 ? Err.na : i + 1;
  },
};

// ---------------------------------------------------------------- array shaping

const pick = (byCol: boolean): FunctionDef => ({
  name: byCol ? 'CHOOSECOLS' : 'CHOOSEROWS', min: 2, max: 255,
  fn: (a, ctx) => {
    const m0 = arr(a[0]!, ctx);
    if (isError(m0)) return m0;
    const m = byCol ? transpose(m0) : m0;
    const out: Matrix = [];
    for (const x of a.slice(1)) {
      const ns = nums(x, ctx);
      if (isError(ns)) return ns;
      for (const n of ns) {
        if (n === 0 || Math.abs(n) > m.length) return Err.value;
        out.push(m[n > 0 ? n - 1 : m.length + n]!);
      }
    }
    return byCol ? transpose(out) : out;
  },
});

const takeDrop = (take: boolean): FunctionDef => ({
  name: take ? 'TAKE' : 'DROP', min: 2, max: 3,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    const r = a[1] === null || a[1] === undefined ? null : int(a[1], ctx);
    const c = a[2] === null || a[2] === undefined ? null : int(a[2], ctx);
    if (isError(r)) return r; if (isError(c)) return c;
    const span = (n: number | null, len: number): [number, number] => {
      if (n === null) return [0, len];
      const k = Math.min(Math.abs(n), len);
      if (take) return n >= 0 ? [0, k] : [len - k, len];
      return n >= 0 ? [k, len] : [0, len - k];
    };
    if (take && (r === 0 || c === 0)) return Err.calc;
    const [r1, r2] = span(r, m.length), [c1, c2] = span(c, m[0]!.length);
    return matrix(m.slice(r1, r2).map((row) => row.slice(c1, c2)));
  },
});

const EXPAND: FunctionDef = {
  name: 'EXPAND', min: 2, max: 4,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    const r = a[1] === null ? m.length : int(a[1]!, ctx);
    const c = a[2] === undefined || a[2] === null ? m[0]!.length : int(a[2], ctx);
    if (isError(r)) return r; if (isError(c)) return c;
    if (r < m.length || c < m[0]!.length) return Err.value;
    if (r > MAX_ROWS || c > MAX_COLS) return Err.num;
    const pad = a[3] === undefined ? Err.na : scalarOf(a[3], ctx);
    return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => (i < m.length && j < m[0]!.length ? m[i]![j]! : pad)));
  },
};

const FILTER: FunctionDef = {
  name: 'FILTER', min: 2, max: 3,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    const inc = arr(a[1]!, ctx);
    if (isError(inc)) return inc;
    const R0 = m.length, C0 = m[0]!.length;
    let byRow: boolean;
    if (inc[0]!.length === 1 && inc.length === R0) byRow = true;
    else if (inc.length === 1 && inc[0]!.length === C0) byRow = false;
    else return Err.value;
    const keep: boolean[] = [];
    for (const v of flat(inc)) {
      if (isError(v)) return v;
      if (typeof v === 'string') return Err.value;
      keep.push(typeof v === 'number' ? v !== 0 : v === true);
    }
    const out = byRow ? m.filter((_, i) => keep[i]) : m.map((r) => r.filter((_, j) => keep[j]));
    if (!out.length || !out[0]!.length) return a[2] === undefined ? Err.calc : a[2];
    return out;
  },
};

const stack = (vertical: boolean): FunctionDef => ({
  name: vertical ? 'VSTACK' : 'HSTACK', min: 1, max: 255,
  fn: (a, ctx) => {
    // HSTACK(SUM, AVERAGE): a list of functions for GROUPBY and PIVOTBY.
    if (a.length && a.every((v) => isLambda(v))) {
      const list = (a as LambdaValue[]).flatMap((l) => l.list ?? [l]);
      return { kind: 'lambda', params: ['x'], invoke: () => Err.calc, list, vertical } satisfies LambdaValue;
    }
    const parts: Matrix[] = [];
    for (const v of a) {
      const m = arr(v, ctx);
      if (isError(m)) return m;
      parts.push(vertical ? m : transpose(m));
    }
    const w = Math.max(...parts.map((p) => p[0]!.length));
    const out = parts.flatMap((p) => p.map((r) => (r.length < w ? [...r, ...Array<Scalar>(w - r.length).fill(Err.na)] : r)));
    return vertical ? out : transpose(out);
  },
});

const toVec = (asCol: boolean): FunctionDef => ({
  name: asCol ? 'TOCOL' : 'TOROW', min: 1, max: 3,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    const ig = optNum(a[1], ctx, 0), byCol = optBool(a[2], ctx, false);
    if (isError(ig)) return ig; if (isError(byCol)) return byCol;
    const mode = Math.trunc(ig);
    if (mode < 0 || mode > 3) return Err.value;
    const cells = flat(byCol ? transpose(m) : m).filter((c) => !((mode & 1 && c === null) || (mode & 2 && isError(c))));
    if (!cells.length) return Err.calc;
    return asCol ? cells.map((c) => [c]) : [cells];
  },
});

const TRANSPOSE: FunctionDef = {
  name: 'TRANSPOSE', min: 1, max: 1,
  fn: ([v], ctx) => { const m = arr(v!, ctx); return isError(m) ? m : transpose(m); },
};

const wrap = (rows: boolean): FunctionDef => ({
  name: rows ? 'WRAPROWS' : 'WRAPCOLS', min: 2, max: 3,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    if (m.length !== 1 && m[0]!.length !== 1) return Err.value;
    const n = int(a[1]!, ctx);
    if (isError(n)) return n;
    if (n < 1) return Err.num;
    const pad = a[2] === undefined ? Err.na : scalarOf(a[2], ctx);
    const cells = flat(m);
    const out: Matrix = [];
    for (let i = 0; i < cells.length; i += n) out.push(Array.from({ length: n }, (_, k) => (i + k < cells.length ? cells[i + k]! : pad)));
    return rows ? out : transpose(out);
  },
});

// ---------------------------------------------------------------- SORT / SORTBY / UNIQUE

/** Sort comparator: blanks stay last in both directions, as in other spreadsheets. */
const sortCmp = (x: Scalar, y: Scalar, desc: boolean): number => {
  if (x === null || y === null) return x === y ? 0 : x === null ? 1 : -1;
  const c = compareValues(x, y);
  return desc ? -c : c;
};

const SORT: FunctionDef = {
  name: 'SORT', min: 1, max: 4,
  fn: (a, ctx) => {
    const m0 = arr(a[0]!, ctx);
    if (isError(m0)) return m0;
    const byCol = optBool(a[3], ctx, false);
    if (isError(byCol)) return byCol;
    const m = byCol ? transpose(m0) : m0;
    const idx = a[1] === undefined || a[1] === null ? [1] : nums(a[1], ctx);
    const ord = a[2] === undefined || a[2] === null ? [1] : nums(a[2], ctx);
    if (isError(idx)) return idx; if (isError(ord)) return ord;
    const w = m[0]!.length;
    if (idx.some((k) => k < 1 || k > w) || ord.some((o) => o !== 1 && o !== -1)) return Err.value;
    if (ord.length > 1 && ord.length !== idx.length) return Err.value;
    const keys = idx.map((k, n) => ({ k: k - 1, desc: (ord.length === 1 ? ord[0] : ord[n]) === -1 }));
    const out = [...m].sort((x, y) => { for (const { k, desc } of keys) { const c = sortCmp(x[k]!, y[k]!, desc); if (c) return c; } return 0; });
    return byCol ? transpose(out) : out;
  },
};

const SORTBY: FunctionDef = {
  name: 'SORTBY', min: 1, max: 255,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    if (a.length === 1) return SORT.fn([m], ctx);
    const keys: { cells: Scalar[]; desc: boolean }[] = [];
    let byCol: boolean | null = null;
    for (let i = 1; i < a.length; i += 2) {
      const b = arr(a[i]!, ctx);
      if (isError(b)) return b;
      const o = a[i + 1] === undefined || a[i + 1] === null ? 1 : int(a[i + 1]!, ctx);
      if (isError(o)) return o;
      if (o !== 1 && o !== -1) return Err.value;
      const fitsRows = b[0]!.length === 1 && b.length === m.length;
      const fitsCols = b.length === 1 && b[0]!.length === m[0]!.length;
      const dir: boolean | null = fitsRows && byCol !== true ? false : fitsCols && byCol !== false ? true : null;
      if (dir === null) return Err.value;
      byCol = dir;
      keys.push({ cells: flat(b), desc: o === -1 });
    }
    const n = byCol ? m[0]!.length : m.length;
    const order = Array.from({ length: n }, (_, i) => i).sort((x, y) => {
      for (const k of keys) { const c = sortCmp(k.cells[x]!, k.cells[y]!, k.desc); if (c) return c; }
      return x - y;
    });
    return byCol ? m.map((r) => order.map((j) => r[j]!)) : order.map((i) => m[i]!);
  },
};

const keyOf = (v: Scalar): string =>
  v === null ? 'b' : typeof v === 'number' ? `n${v}` : typeof v === 'string' ? `s${v.toLowerCase()}` : typeof v === 'boolean' ? `l${v}` : `e${v.code}`;

const UNIQUE: FunctionDef = {
  name: 'UNIQUE', min: 1, max: 3,
  fn: (a, ctx) => {
    const m0 = arr(a[0]!, ctx);
    if (isError(m0)) return m0;
    const byCol = optBool(a[1], ctx, false), once = optBool(a[2], ctx, false);
    if (isError(byCol)) return byCol; if (isError(once)) return once;
    const m = byCol ? transpose(m0) : m0;
    const seen = new Map<string, { row: Scalar[]; n: number }>();
    for (const r of m) {
      const k = r.map(keyOf).join('\u0000');
      const e = seen.get(k);
      if (e) e.n++; else seen.set(k, { row: r, n: 1 });
    }
    const out = [...seen.values()].filter((e) => !once || e.n === 1).map((e) => e.row);
    const res = matrix(out);
    return isError(res) ? res : byCol ? transpose(res) : res;
  },
};

// ---------------------------------------------------------------- generators

const SEQUENCE: FunctionDef = {
  name: 'SEQUENCE', min: 1, max: 4,
  fn: (a, ctx) => {
    const v = [optNum(a[0], ctx, 1), optNum(a[1], ctx, 1), optNum(a[2], ctx, 1), optNum(a[3], ctx, 1)];
    for (const x of v) if (isError(x)) return x;
    const [r0, c0, start, step] = v as number[];
    const r = Math.trunc(r0!), c = Math.trunc(c0!);
    if (r < 0 || c < 0) return Err.value;
    if (r === 0 || c === 0) return Err.calc;
    if (r * c > MAX_ROWS * 16) return Err.num;
    return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => start! + (i * c + j) * step!));
  },
};

const RANDARRAY: FunctionDef = {
  name: 'RANDARRAY', min: 0, max: 5, volatile: true,
  fn: (a, ctx) => {
    const v = [optNum(a[0], ctx, 1), optNum(a[1], ctx, 1), optNum(a[2], ctx, 0), optNum(a[3], ctx, 1)];
    for (const x of v) if (isError(x)) return x;
    const integer = optBool(a[4], ctx, false);
    if (isError(integer)) return integer;
    const [r0, c0, lo, hi] = v as number[];
    const r = Math.trunc(r0!), c = Math.trunc(c0!);
    if (r < 0 || c < 0 || lo! > hi!) return Err.value;
    if (r === 0 || c === 0) return Err.calc;
    if (integer && (!Number.isInteger(lo) || !Number.isInteger(hi))) return Err.value;
    const gen = (): number => (integer ? lo! + Math.floor(Math.random() * (hi! - lo! + 1)) : lo! + Math.random() * (hi! - lo!));
    return Array.from({ length: r }, () => Array.from({ length: c }, gen));
  },
};

const TRIMRANGE: FunctionDef = {
  name: 'TRIMRANGE', min: 1, max: 3,
  fn: (a, ctx) => {
    const v = a[0]!;
    const tr = optNum(a[1], ctx, 3), tc = optNum(a[2], ctx, 3);
    if (isError(tr)) return tr; if (isError(tc)) return tc;
    if (![0, 1, 2, 3].includes(Math.trunc(tr)) || ![0, 1, 2, 3].includes(Math.trunc(tc))) return Err.value;
    if (isRefs(v)) return Err.value;
    const m = isRef(v) ? ctx.values(v) : toMatrix(v, ctx);
    if (isError(m)) return m;
    const [R0, C0] = isRef(v) ? refArea(v) : [m.length, m[0]!.length];
    const blankRow = (i: number): boolean => (m[i] ?? []).every((c) => c === null);
    const blankCol = (j: number): boolean => m.every((r) => (r[j] ?? null) === null);
    let i1 = 0, i2 = R0 - 1, j1 = 0, j2 = C0 - 1;
    // Cells past the trimmed values are blank, so trailing trims start from the used area.
    if (tr & 2) { i2 = Math.min(i2, m.length - 1); while (i2 >= 0 && blankRow(i2)) i2--; }
    if (tr & 1) while (i1 <= i2 && blankRow(i1)) i1++;
    if (tc & 2) { j2 = Math.min(j2, (m[0]?.length ?? 0) - 1); while (j2 >= 0 && blankCol(j2)) j2--; }
    if (tc & 1) while (j1 <= j2 && blankCol(j1)) j1++;
    if (i1 > i2 || j1 > j2) return Err.calc;
    if (isRef(v)) return sub(v, i1, j1, i2, j2);
    return m.slice(i1, i2 + 1).map((r) => r.slice(j1, j2 + 1));
  },
};

// ---------------------------------------------------------------- LAMBDA helpers

const byRowCol = (rows: boolean): FunctionDef => ({
  name: rows ? 'BYROW' : 'BYCOL', min: 2, max: 2,
  fn: (a, ctx) => {
    const m = arr(a[0]!, ctx);
    if (isError(m)) return m;
    const f = lam(a[1], 1);
    if (isError(f)) return f;
    const slices = rows ? m.map((r) => [r]) : transpose(m).map((c) => c.map((x) => [x]));
    const res = slices.map((s) => one(ctx.call(f, [s]), ctx));
    return rows ? res.map((x) => [x]) : [res];
  },
});

const MAP: FunctionDef = {
  name: 'MAP', min: 2, max: 255,
  fn: (a, ctx) => {
    const f = lam(a[a.length - 1], a.length - 1);
    if (isError(f)) return f;
    const ms: Matrix[] = [];
    for (const v of a.slice(0, -1)) { const m = arr(v, ctx); if (isError(m)) return m; ms.push(m); }
    const R0 = Math.max(...ms.map((m) => m.length)), C0 = Math.max(...ms.map((m) => m[0]!.length));
    // A 1×1 argument is broadcast; any other size mismatch fills with #N/A.
    const at = (m: Matrix, i: number, j: number): Scalar | undefined =>
      m.length === 1 && m[0]!.length === 1 ? m[0]![0]! : m[i]?.[j];
    return Array.from({ length: R0 }, (_, i) => Array.from({ length: C0 }, (_, j) => {
      const xs = ms.map((m) => at(m, i, j));
      return xs.some((x) => x === undefined) ? Err.na : one(ctx.call(f, xs as Scalar[]), ctx);
    }));
  },
};

const reduceScan = (scan: boolean): FunctionDef => ({
  name: scan ? 'SCAN' : 'REDUCE', min: 2, max: 3,
  fn: (a, ctx) => {
    const [init, src, fv] = a.length === 2 ? [null, a[0]!, a[1]] : [a[0]!, a[1]!, a[2]];
    const f = lam(fv, 2);
    if (isError(f)) return f;
    const m = arr(src, ctx);
    if (isError(m)) return m;
    let acc: Value = isRef(init) ? one(init, ctx) : init;
    const out = m.map((r) => r.map((x) => {
      acc = ctx.call(f, [acc, x]);
      return scan ? one(acc, ctx) : null;
    }));
    return scan ? out : acc;
  },
});

const MAKEARRAY: FunctionDef = {
  name: 'MAKEARRAY', min: 3, max: 3,
  fn: (a, ctx) => {
    const r = int(a[0]!, ctx), c = int(a[1]!, ctx);
    if (isError(r)) return r; if (isError(c)) return c;
    if (r < 1 || c < 1) return Err.value;
    const f = lam(a[2], 2);
    if (isError(f)) return f;
    return Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => one(ctx.call(f, [i + 1, j + 1]), ctx)));
  },
};

export const defs: FunctionDef[] = [
  ADDRESS, AREAS, pick(true), pick(false), rowOrCol(false), COLUMNS, takeDrop(false), EXPAND, FILTER,
  vhlookup(false), stack(false), HYPERLINK, INDEX, LOOKUP, MATCH, OFFSET, rowOrCol(true), ROWS, SORT, SORTBY,
  takeDrop(true), toVec(true), toVec(false), TRANSPOSE, UNIQUE, vhlookup(true), stack(true), wrap(false), wrap(true),
  XLOOKUP, XMATCH, SEQUENCE, RANDARRAY, TRIMRANGE, byRowCol(true), byRowCol(false), MAP, reduceScan(false),
  reduceScan(true), MAKEARRAY,
];
