// Operational transformation: rewrite an operation so it applies after a concurrent one.
// xform(a, b, aLater) returns a' such that applying b then a' gives the same result as applying a then xform(b, a, !aLater).
// "Later" is the host's order: when two ops set the same thing, the later one wins.
import { adjustForStructure, deleteSheetInFormula, renameSheetInFormula, type Structural } from '../../engine/refactor';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import type { CellRange } from '../../model/workbook';
import type { Pos, Selection } from '../../app/store';
import type { CellEntry, Op } from './ops';
import type { Json } from './wire';

export type Struct = Extract<Op, { t: 'struct' }>;

// ---------- Positions ----------

/** New index of `i` after an insert (n > 0) or delete (n < 0) at `at`; null when deleted. */
export const shiftIndex = (i: number, at: number, n: number): number | null => {
  if (n > 0) return i >= at ? i + n : i;
  const del = -n;
  if (i >= at && i < at + del) return null;
  return i >= at + del ? i - del : i;
};

/** Moves a range through an insert or delete; ranges that are entirely deleted give null; whole rows/columns stay. */
export const shiftRange = (r: CellRange, axis: 'row' | 'col', at: number, n: number): CellRange | null => {
  const max = axis === 'row' ? MAX_ROWS - 1 : MAX_COLS - 1;
  const [a, b] = axis === 'row' ? [r.r1, r.r2] : [r.c1, r.c2];
  if (a === 0 && b === max) return { ...r };
  let na: number, nb: number;
  if (n > 0) { na = a >= at ? a + n : a; nb = b >= at ? b + n : b; }
  else {
    const del = -n, end = at + del - 1;
    if (a >= at && b <= end) return null;
    na = a > end ? a - del : a >= at ? at : a;
    nb = b > end ? b - del : b >= at ? at - 1 : b;
    if (nb < na) return null;
  }
  na = Math.min(na, max); nb = Math.min(nb, max);
  return axis === 'row' ? { ...r, r1: na, r2: nb } : { ...r, c1: na, c2: nb };
};

/** A cursor moves with its cell; a deleted cell's cursor lands on the first row/column after the deletion. */
export const shiftPos = (p: Pos, axis: 'row' | 'col', at: number, n: number): Pos => {
  const i = axis === 'row' ? p.row : p.col;
  const ni = shiftIndex(i, at, n) ?? at;
  const max = axis === 'row' ? MAX_ROWS - 1 : MAX_COLS - 1;
  const v = Math.max(0, Math.min(max, ni));
  return axis === 'row' ? { ...p, row: v } : { ...p, col: v };
};

export const shiftSelection = (sel: Selection, axis: 'row' | 'col', at: number, n: number): Selection => {
  const active = shiftPos(sel.active, axis, at, n);
  const ranges = sel.ranges.map((r) => shiftRange(r, axis, at, n)).filter((r): r is CellRange => !!r);
  return { ranges: ranges.length ? ranges : [{ r1: active.row, c1: active.col, r2: active.row, c2: active.col }], active, anchor: shiftPos(sel.anchor, axis, at, n) };
};

// ---------- Formulas inside ops ----------

type FormulaFn = (formula: string, hostSheet: string) => string;

const mapJsonFormulas = (v: Json, keys: string[], fn: (f: string) => string, depth = 0): Json => {
  if (depth > 12 || v === null || typeof v !== 'object') return v;
  if (Array.isArray(v)) return v.map((x) => mapJsonFormulas(x, keys, fn, depth + 1));
  const out: Record<string, Json> = {};
  for (const [k, x] of Object.entries(v)) {
    if (keys.includes(k) && typeof x === 'string') out[k] = fn(x);
    else if (keys.includes(k) && Array.isArray(x) && x.every((y) => typeof y === 'string')) out[k] = x.map((y) => fn(y as string));
    else out[k] = mapJsonFormulas(x, keys, fn, depth + 1);
  }
  return out;
};

const safe = (fn: FormulaFn): FormulaFn => (f, host) => { try { return fn(f, host); } catch { return f; } };

/** Rewrites every formula an op carries (cells, chart series, rules, validations, names). */
const mapFormulas = (op: Op, raw: FormulaFn): Op => {
  const fn = safe(raw);
  switch (op.t) {
    case 'cells': return { ...op, c: op.c.map(([r, c, w]) => [r, c, w?.f ? { ...w, f: fn(w.f, op.s) } : w] as CellEntry) };
    case 'item': return op.k === 'charts' ? { ...op, v: mapJsonFormulas(op.v, ['values', 'categories', 'name', 'sizes'], (f) => fn(f, op.s)) } : op;
    case 'prop':
      if (op.k === 'conditionalFormats') return { ...op, v: mapJsonFormulas(op.v, ['formula', 'formulas'], (f) => fn(f, op.s)) };
      if (op.k === 'validations') return { ...op, v: mapJsonFormulas(op.v, ['formula1', 'formula2'], (f) => fn(f.replace(/^=/, ''), op.s)) };
      if (op.k === 'charts') return { ...op, v: mapJsonFormulas(op.v, ['values', 'categories', 'name', 'sizes'], (f) => fn(f, op.s)) };
      return op;
    case 'book':
      return op.k === 'names' ? { ...op, v: mapJsonFormulas(op.v, ['formula'], (f) => fn(f, '')) } : op;
    default: return op;
  }
};

// ---------- Ranges inside ops ----------

type Rng = { r1: number; c1: number; r2: number; c2: number };
const isRng = (v: unknown): v is Rng => !!v && typeof v === 'object' && ['r1', 'c1', 'r2', 'c2'].every((k) => typeof (v as Record<string, unknown>)[k] === 'number');

/** Moves every {r1,c1,r2,c2} inside a JSON value; list entries whose range is deleted are removed (and objects left with no ranges). */
const shiftJsonRanges = (v: Json, axis: 'row' | 'col', at: number, n: number, depth = 0): Json | undefined => {
  if (depth > 12 || v === null || typeof v !== 'object') return v;
  if (isRng(v)) return (shiftRange(v as CellRange, axis, at, n) as unknown as Json) ?? undefined;
  if (Array.isArray(v)) {
    const out: Json[] = [];
    for (const x of v) { const y = shiftJsonRanges(x, axis, at, n, depth + 1); if (y !== undefined) out.push(y); }
    return out;
  }
  const out: Record<string, Json> = {};
  for (const [k, x] of Object.entries(v)) {
    const y = shiftJsonRanges(x, axis, at, n, depth + 1);
    if (y === undefined) { if (k === 'range') return undefined; continue; }
    if (k === 'ranges' && Array.isArray(y) && !y.length && Array.isArray(x) && x.length) return undefined;
    out[k] = y;
  }
  return out;
};

const shiftAnchor = (v: Json, axis: 'row' | 'col', at: number, n: number): Json => {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return v;
  const a = v.anchor;
  if (!a || typeof a !== 'object' || Array.isArray(a)) return v;
  const p = axis === 'row' ? a.row : a.col;
  if (typeof p !== 'number' || p < at) return v;
  const np = n > 0 ? p + n : Math.max(at, p + n);
  return { ...v, anchor: { ...a, [axis]: np } };
};

// ---------- Transform ----------

const sheetOf = (op: Op): string | null => ('s' in op && op.t !== 'sheetAdd' ? op.s : null);

const xformStruct = (a: Op, b: Struct, aLater: boolean): Op[] => {
  const st: Structural = { sheet: b.nm, axis: b.a, at: b.at, count: b.n };
  // Ops name sheets by id; unqualified references in a formula are on its own sheet.
  const fixed = mapFormulas(a, (f, host) => adjustForStructure(f, host === b.s ? b.nm : `\u0001${host}`, st));
  if (sheetOf(fixed) !== b.s) return [fixed];
  const op = fixed;
  switch (op.t) {
    case 'cells': {
      const c: CellEntry[] = [];
      for (const [r, col, w] of op.c) {
        const nr = b.a === 'row' ? shiftIndex(r, b.at, b.n) : r, nc = b.a === 'col' ? shiftIndex(col, b.at, b.n) : col;
        if (nr === null || nc === null || nr >= MAX_ROWS || nc >= MAX_COLS) continue;
        let cell = w;
        if (w?.ar) { const r2 = shiftRange({ r1: w.ar[0], c1: w.ar[1], r2: w.ar[2], c2: w.ar[3] }, b.a, b.at, b.n); cell = { ...w, ar: r2 ? [r2.r1, r2.c1, r2.r2, r2.c2] : undefined }; if (!r2) delete cell.ar; }
        c.push([nr, nc, cell]);
      }
      return c.length ? [{ ...op, c }] : [];
    }
    case 'keyed': {
      const c: [number, number, Json][] = [];
      for (const [r, col, v] of op.c) {
        const nr = b.a === 'row' ? shiftIndex(r, b.at, b.n) : r, nc = b.a === 'col' ? shiftIndex(col, b.at, b.n) : col;
        if (nr !== null && nc !== null && nr < MAX_ROWS && nc < MAX_COLS) c.push([nr, nc, v]);
      }
      return c.length ? [{ ...op, c }] : [];
    }
    case 'dim': {
      if (op.a !== b.a) return [op];
      const c = op.c.map(([i, v]) => [shiftIndex(i, b.at, b.n), v] as const).filter((e): e is [number, typeof e[1]] => e[0] !== null && e[0] < (b.a === 'row' ? MAX_ROWS : MAX_COLS));
      return c.length ? [{ ...op, c }] : [];
    }
    case 'item': return [{ ...op, v: shiftAnchor(op.v, b.a, b.at, b.n) }];
    case 'prop': {
      if (op.k === 'charts' || op.k === 'pictures' || op.k === 'shapes' || op.k === 'slicers') {
        return [{ ...op, v: Array.isArray(op.v) ? op.v.map((x) => shiftAnchor(x, b.a, b.at, b.n)) : op.v }];
      }
      if (op.k === 'view' || op.k === 'pageSetup' || op.k === 'protection') return [op];
      const v = shiftJsonRanges(op.v, b.a, b.at, b.n);
      return [{ ...op, v: v === undefined ? null : v }];
    }
    case 'struct': {
      if (op.a !== b.a) return [op];
      return structVsStruct(op, b, aLater);
    }
    default: return [op];
  }
};

const structVsStruct = (a: Struct, b: Struct, aLater: boolean): Struct[] => {
  const p = a.at, q = b.at;
  if (a.n > 0 && b.n > 0) return [p < q || (p === q && !aLater) ? a : { ...a, at: p + b.n }];
  if (a.n > 0) { // insert vs delete
    const m = -b.n;
    if (p <= q) return [a];
    if (p >= q + m) return [{ ...a, at: p - m }];
    return [{ ...a, at: q }];
  }
  const m = -a.n;
  if (b.n > 0) { // delete vs insert
    const k = b.n;
    if (q <= p) return [{ ...a, at: p + k }];
    if (q >= p + m) return [a];
    // The insert landed inside the deleted block: delete around it, keeping the inserted lines.
    return [{ ...a, at: q + k, n: -(p + m - q) }, { ...a, at: p, n: -(q - p) }];
  }
  const m2 = -b.n;
  const overlap = Math.max(0, Math.min(p + m, q + m2) - Math.max(p, q));
  const left = m - overlap;
  if (left <= 0) return [];
  const start = p <= q ? p : p >= q + m2 ? p - m2 : q;
  return [{ ...a, at: start, n: -left }];
};

const keysOf = (op: Op): Set<string> | null => {
  switch (op.t) {
    case 'cells': return new Set(op.c.map(([r, c]) => `${r},${c}`));
    case 'keyed': return new Set(op.c.map(([r, c]) => `${r},${c}`));
    case 'dim': return new Set(op.c.map(([i]) => String(i)));
    default: return null;
  }
};

/** Two ops that set the same things: the later one wins, so the earlier one loses those parts. */
const sameTarget = (a: Op, b: Op): boolean => {
  if (a.t !== b.t) return false;
  switch (a.t) {
    case 'cells': return a.s === (b as typeof a).s;
    case 'keyed': return a.s === (b as typeof a).s && a.k === (b as typeof a).k;
    case 'dim': return a.s === (b as typeof a).s && a.a === (b as typeof a).a;
    case 'item': return a.s === (b as typeof a).s && a.k === (b as typeof a).k && a.id === (b as typeof a).id;
    case 'prop': return a.s === (b as typeof a).s && a.k === (b as typeof a).k;
    case 'book': return a.k === (b as typeof a).k;
    case 'sheetPos': return a.s === (b as typeof a).s;
    default: return false;
  }
};

/** The name an op gives a sheet: a new sheet's name or a rename's new name. */
const claimOf = (o: Op): { id: string; name: string } | null => (o.t === 'sheetAdd' ? { id: o.s, name: o.name } : o.t === 'sheetRen' ? { id: o.s, name: o.to } : null);

/** A name for a sheet that lost a clash: the name with part of the sheet's id, so it can't clash again. */
export const clashName = (name: string, id: string): string => `${name.slice(0, 24)} (${id.slice(1, 5)})`;

/** Two sheets given the same name at once: the later claim gets a clash name, on every copy. */
const nameClash = (a: Op, bId: string, bName: string, aLater: boolean): Op[] => {
  const c = claimOf(a);
  if (!c || c.id === bId || c.name.toLowerCase() !== bName.toLowerCase()) return [a];
  if (aLater) {
    const to = clashName(c.name, c.id);
    return [a.t === 'sheetAdd' ? { ...a, name: to } : { ...(a as Extract<Op, { t: 'sheetRen' }>), to }];
  }
  return [{ t: 'sheetRen', s: bId, from: bName, to: clashName(bName, bId) }, a];
};

export const xform = (a: Op, b: Op, aLater: boolean): Op[] => {
  switch (b.t) {
    case 'struct': return xformStruct(a, b, aLater);
    case 'sheetDel': {
      // Bringing a deleted sheet back (undo) while someone deletes it: whichever the host ordered later wins.
      if (a.t === 'sheetAdd' && a.s === b.s) return aLater ? [a] : [];
      if (a.t === 'sheetDel' && a.s === b.s) return [];
      // Other edits to the deleted sheet still land on its kept copy (apply.ts), so every copy keeps the same one.
      return [mapFormulas(a, (f) => deleteSheetInFormula(f, b.name))];
    }
    case 'sheetRen': {
      const fixed = mapFormulas(a, (f) => renameSheetInFormula(f, b.from, b.to));
      if (fixed.t === 'sheetRen' && fixed.s === b.s) return aLater ? [{ ...fixed, from: b.to }] : [];
      // A sheet's name, like its place, is one value: a later re-add (undo of a delete) sets it again.
      if (fixed.t === 'sheetAdd' && fixed.s === b.s) return [aLater ? fixed : { ...fixed, name: b.to }];
      if (fixed.t === 'struct' && fixed.s === b.s) return [{ ...fixed, nm: b.to }];
      if (fixed.t === 'sheetDel' && fixed.s === b.s) return [{ ...fixed, name: b.to }];
      return nameClash(fixed, b.s, b.to, aLater);
    }
    case 'sheetAdd':
      if (a.t === 'sheetAdd' && a.s === b.s) return aLater ? [] : [a];
      if (a.t === 'sheetDel' && a.s === b.s) return aLater ? [a] : [];
      if (a.t === 'sheetRen' && a.s === b.s) return aLater ? [{ ...a, from: b.name }] : [];
      if (a.t === 'sheetPos' && a.s === b.s) return aLater ? [a] : [];
      return nameClash(a, b.s, b.name, aLater);
    case 'sheetPos':
      if (a.t === 'sheetAdd' && a.s === b.s) return [aLater ? a : { ...a, pos: b.pos }];
      break;
    default: break;
  }
  if (aLater || !sameTarget(a, b)) return [a];
  const bk = keysOf(b);
  if (!bk) return [];
  if (a.t === 'cells') { const c = a.c.filter(([r, col]) => !bk.has(`${r},${col}`)); return c.length ? [{ ...a, c }] : []; }
  if (a.t === 'keyed') { const c = a.c.filter(([r, col]) => !bk.has(`${r},${col}`)); return c.length ? [{ ...a, c }] : []; }
  if (a.t === 'dim') { const c = a.c.filter(([i]) => !bk.has(String(i))); return c.length ? [{ ...a, c }] : []; }
  return [a];
};

/** Transforms two lists of ops against each other: returns [A after B, B after A]. */
export const xformLists = (A: Op[], B: Op[], aLater: boolean): [Op[], Op[]] => {
  if (!A.length || !B.length) return [A, B];
  if (A.length === 1 && B.length === 1) return [xform(A[0]!, B[0]!, aLater), xform(B[0]!, A[0]!, !aLater)];
  if (A.length > 1) {
    const [a1, B1] = xformLists([A[0]!], B, aLater);
    const [a2, B2] = xformLists(A.slice(1), B1, aLater);
    return [[...a1, ...a2], B2];
  }
  const [A1, b1] = xformLists(A, [B[0]!], aLater);
  const [A2, b2] = xformLists(A1, B.slice(1), aLater);
  return [A2, [...b1, ...b2]];
};
