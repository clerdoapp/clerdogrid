// Plain-JSON forms of model values for messages: tagged values, cells, rows/columns, sheets and whole workbooks.
import { CellError, ERROR_CODES, MAX_COLS, MAX_ROWS, type ErrorCode, type Scalar } from '../../engine/types';
import {
  cellKey, internStyle, keyCol, keyRow, newSheet,
  type Cell, type CellRange, type CellStyle, type RichRun, type Sheet, type Workbook,
} from '../../model/workbook';
import { b64urlDecode, b64urlEncode } from './b64';

export type Json = null | boolean | number | string | Json[] | { [k: string]: Json };

export const MAX_DEPTH = 48;
const BAD_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
export const safeKey = (k: string): boolean => !BAD_KEYS.has(k) && k.length <= 64;

// ---------- Tagged values ----------
// Map -> {"$m": [[k, v], ...]}, Uint8Array -> {"$b": base64url}, CellError -> {"$e": code}. Undefined fields are dropped.

export const toWire = (x: unknown, depth = 0): Json => {
  if (depth > MAX_DEPTH) throw new Error('Too deep');
  if (x === null || x === undefined) return null;
  if (typeof x === 'number') return Number.isFinite(x) ? x : null;
  if (typeof x === 'string' || typeof x === 'boolean') return x;
  if (x instanceof CellError) return { $e: x.code };
  if (x instanceof Uint8Array) return { $b: b64urlEncode(x) };
  if (x instanceof Map) return { $m: [...x].map(([k, v]) => [toWire(k, depth + 1), toWire(v, depth + 1)]) };
  if (Array.isArray(x)) return x.map((v) => toWire(v, depth + 1));
  if (typeof x === 'object') {
    const out: Record<string, Json> = {};
    for (const [k, v] of Object.entries(x as Record<string, unknown>)) if (v !== undefined && typeof v !== 'function' && safeKey(k)) out[k] = toWire(v, depth + 1);
    return out;
  }
  return null;
};

export const fromWire = (x: unknown, depth = 0): unknown => {
  if (depth > MAX_DEPTH) throw new Error('Too deep');
  if (x === null || typeof x === 'boolean' || typeof x === 'string') return x;
  if (typeof x === 'number') { if (!Number.isFinite(x)) throw new Error('Bad number'); return x; }
  if (Array.isArray(x)) return x.map((v) => fromWire(v, depth + 1));
  if (typeof x === 'object') {
    const o = x as Record<string, unknown>;
    const keys = Object.keys(o);
    if (keys.length === 1) {
      if (keys[0] === '$e') return errorOf(o.$e);
      if (keys[0] === '$b') { if (typeof o.$b !== 'string') throw new Error('Bad bytes'); return b64urlDecode(o.$b); }
      if (keys[0] === '$m') {
        if (!Array.isArray(o.$m)) throw new Error('Bad map');
        const m = new Map<unknown, unknown>();
        for (const e of o.$m) {
          if (!Array.isArray(e) || e.length !== 2) throw new Error('Bad map');
          m.set(fromWire(e[0], depth + 1), fromWire(e[1], depth + 1));
        }
        return m;
      }
    }
    const out: Record<string, unknown> = {};
    for (const k of keys) if (safeKey(k)) out[k] = fromWire(o[k], depth + 1);
    return out;
  }
  throw new Error('Bad value');
};

const errorOf = (code: unknown): CellError => {
  if (typeof code !== 'string' || !ERROR_CODES.includes(code as ErrorCode)) throw new Error('Bad error value');
  return new CellError(code as ErrorCode);
};

/** Deep equality for model values (Maps, bytes, errors; undefined fields ignored). */
export const deepEq = (a: unknown, b: unknown, depth = 0): boolean => {
  if (a === b) return true;
  if (depth > MAX_DEPTH) return false;
  if (a === undefined || b === undefined || a === null || b === null || typeof a !== 'object' || typeof b !== 'object') {
    return (a ?? null) === (b ?? null) || (typeof a === 'number' && typeof b === 'number' && Number.isNaN(a) && Number.isNaN(b));
  }
  if (a instanceof CellError || b instanceof CellError) return a instanceof CellError && b instanceof CellError && a.code === b.code;
  if (a instanceof Uint8Array || b instanceof Uint8Array) {
    if (!(a instanceof Uint8Array && b instanceof Uint8Array) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }
  if (a instanceof Map || b instanceof Map) {
    if (!(a instanceof Map && b instanceof Map) || a.size !== b.size) return false;
    for (const [k, v] of a) if (!b.has(k) || !deepEq(v, b.get(k), depth + 1)) return false;
    return true;
  }
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) {
    const bb = b as unknown[];
    if (a.length !== bb.length) return false;
    for (let i = 0; i < a.length; i++) if (!deepEq(a[i], bb[i], depth + 1)) return false;
    return true;
  }
  const ao = a as Record<string, unknown>, bo = b as Record<string, unknown>;
  const ka = Object.keys(ao).filter((k) => ao[k] !== undefined), kb = Object.keys(bo).filter((k) => bo[k] !== undefined);
  if (ka.length !== kb.length) return false;
  for (const k of ka) if (!deepEq(ao[k], bo[k], depth + 1)) return false;
  return true;
};

/** JSON with sorted keys, for hashing. */
export const stable = (x: unknown): string => {
  if (x === null || x === undefined) return 'null';
  if (typeof x !== 'object') return JSON.stringify(x);
  if (Array.isArray(x)) return `[${x.map(stable).join(',')}]`;
  const o = x as Record<string, unknown>;
  return `{${Object.keys(o).filter((k) => o[k] !== undefined).sort().map((k) => `${JSON.stringify(k)}:${stable(o[k])}`).join(',')}}`;
};

// ---------- Cells ----------

export type WScalar = number | string | boolean | { $e: string };
/** A cell in a message. `st` is the style itself (style ids differ between copies); `si` indexes a snapshot's style list. */
export interface WCell { v?: WScalar; f?: string; dyn?: true; ar?: [number, number, number, number]; st?: CellStyle; si?: number; rich?: RichRun[] }

export const scalarToWire = (v: Scalar): WScalar | undefined => {
  if (v === null || v === undefined) return undefined;
  if (v instanceof CellError) return { $e: v.code };
  if (typeof v === 'number') return Number.isFinite(v) ? v : undefined;
  return v;
};
export const scalarFromWire = (v: WScalar | undefined): Scalar => {
  if (v === undefined) return null;
  if (typeof v === 'object') return errorOf(v.$e);
  return v;
};

const rngOut = (r: CellRange): [number, number, number, number] => [r.r1, r.c1, r.r2, r.c2];
export const rngIn = (a: [number, number, number, number]): CellRange => ({ r1: a[0], c1: a[1], r2: a[2], c2: a[3] });

const nonEmptyStyle = (s: CellStyle | undefined): CellStyle | undefined => (s && Object.keys(s).some((k) => (s as Record<string, unknown>)[k] !== undefined) ? s : undefined);

/** `styleIds` true writes `si` (the index into `styles`) instead of the style object. */
export const encodeCell = (c: Cell | undefined, styles: CellStyle[], styleIds = false): WCell | null => {
  if (!c) return null;
  const out: WCell = {};
  if (c.f) {
    out.f = c.f;
    if (c.dyn) out.dyn = true;
    if (c.arrayRange) out.ar = rngOut(c.arrayRange);
  } else {
    const v = scalarToWire(c.v);
    if (v !== undefined) out.v = v;
  }
  if (c.s) {
    if (styleIds) out.si = c.s;
    else { const st = nonEmptyStyle(styles[c.s]); if (st) out.st = st; }
  }
  if (c.rich?.length && !c.f) out.rich = c.rich;
  return out.f === undefined && out.v === undefined && out.st === undefined && out.si === undefined && out.rich === undefined ? null : out;
};

export const sameCell = (a: WCell | null, b: WCell | null): boolean => (a === null || b === null ? a === b : deepEq(a, b));

/** Builds a model cell; the style goes into `wb.styles` (or comes from `table` for snapshot cells). */
export const decodeCell = (w: WCell, wb: Workbook, table?: number[]): Cell => {
  const c: Cell = { v: w.f ? null : scalarFromWire(w.v) };
  if (w.f) { c.f = w.f; if (w.dyn) c.dyn = true; if (w.ar) c.arrayRange = rngIn(w.ar); }
  if (w.si !== undefined && table) { const id = table[w.si]; if (id) c.s = id; }
  else if (w.st) { const id = internStyle(wb, w.st); if (id) c.s = id; }
  if (w.rich && !w.f) c.rich = w.rich;
  return c;
};

// ---------- Rows and columns ----------

export type WDim = Record<string, Json> & { st?: Json };
export const encodeDim = (info: { s?: number } | undefined, styles: CellStyle[]): WDim | null => {
  if (!info) return null;
  const { s, ...rest } = info as { s?: number } & Record<string, unknown>;
  const out = toWire(rest) as WDim;
  const st = s ? nonEmptyStyle(styles[s]) : undefined;
  if (st) out.st = toWire(st);
  return Object.keys(out).length ? out : null;
};
export const decodeDim = (w: WDim, wb: Workbook): Record<string, unknown> => {
  const { st, ...rest } = w;
  const info = fromWire(rest) as Record<string, unknown>;
  if (st) { const id = internStyle(wb, fromWire(st) as CellStyle); if (id) info.s = id; }
  return info;
};

// ---------- Sheet parts ----------

/** Cell-keyed maps synced key by key. */
export const KEYED = ['notes', 'hyperlinks', 'comments'] as const;
export type KeyedKey = typeof KEYED[number];
/** Lists of objects with ids, synced item by item. */
export const ITEMS = ['charts', 'pictures', 'shapes', 'ink'] as const;
export type ItemKey = typeof ITEMS[number];
/** Everything else on a sheet is a "prop" synced as a whole value. These are handled specially or are per-person. */
const SPECIAL = new Set<string>(['name', 'cid', 'cpos', 'cells', 'rows', 'cols', ...KEYED, ...ITEMS]);
/** Parts of the sheet view that belong to each person. */
const PERSONAL_VIEW = ['selection', 'zoom'];

export const sheetPropKeys = (...sheets: (Sheet | undefined)[]): string[] => {
  const keys = new Set<string>();
  for (const s of sheets) if (s) for (const k of Object.keys(s)) if (!SPECIAL.has(k) && safeKey(k)) keys.add(k);
  return [...keys];
};

/** A sheet prop as it is shared (the view without selection and zoom). */
export const sharedProp = (s: Sheet, k: string): unknown => {
  const v = (s as unknown as Record<string, unknown>)[k];
  if (k === 'view' && v && typeof v === 'object') {
    const copy = { ...(v as Record<string, unknown>) };
    for (const p of PERSONAL_VIEW) delete copy[p];
    return copy;
  }
  return v;
};

export const setSheetProp = (s: Sheet, k: string, v: unknown): void => {
  const o = s as unknown as Record<string, unknown>;
  if (k === 'view') {
    const cur = s.view;
    const next = { ...(v && typeof v === 'object' ? v as Record<string, unknown> : {}) } as Record<string, unknown>;
    for (const p of PERSONAL_VIEW) next[p] = (cur as unknown as Record<string, unknown>)[p];
    s.view = { ...newSheet('x').view, ...next } as Sheet['view'];
    return;
  }
  if (v === null || v === undefined) delete o[k]; else o[k] = v;
};

const SPECIAL_TOP = new Set(['name', 'cid', 'cpos', 'cells', 'rows', 'cols']);

export interface WSheet {
  name: string;
  /** Live-session id and sort key (sheets.ts). */
  cid?: string;
  cpos?: string;
  cells: [number, number, WCell][];
  rows: [number, WDim][];
  cols: [number, WDim][];
  /** Every other part as tagged JSON (keyed maps, item lists and props). */
  parts: Record<string, Json>;
}

export const encodeSheet = (s: Sheet, styles: CellStyle[], styleIds = false): WSheet => {
  const cells: [number, number, WCell][] = [];
  for (const [k, c] of s.cells) { const w = encodeCell(c, styles, styleIds); if (w) cells.push([keyRow(k), keyCol(k), w]); }
  const dims = (m: Map<number, { s?: number }>) => {
    const out: [number, WDim][] = [];
    for (const [i, info] of m) {
      if (styleIds) { const { s, ...rest } = info as { s?: number }; const w = toWire(rest) as WDim; if (s) w.si = s; out.push([i, w]); }
      else { const w = encodeDim(info, styles); if (w) out.push([i, w]); }
    }
    return out;
  };
  const parts: Record<string, Json> = {};
  for (const k of [...KEYED, ...ITEMS]) { const v = (s as unknown as Record<string, unknown>)[k]; if (v !== undefined) parts[k] = toWire(v); }
  for (const k of sheetPropKeys(s)) { const v = sharedProp(s, k); if (v !== undefined) parts[k] = toWire(v); }
  return { name: s.name, ...(s.cid ? { cid: s.cid } : {}), ...(s.cpos ? { cpos: s.cpos } : {}), cells, rows: dims(s.rows), cols: dims(s.cols), parts };
};

export const decodeSheet = (w: WSheet, wb: Workbook, table?: number[]): Sheet => {
  const s = newSheet(w.name);
  if (w.cid) s.cid = w.cid;
  if (w.cpos) s.cpos = w.cpos;
  for (const [r, c, wc] of w.cells) s.cells.set(cellKey(r, c), decodeCell(wc, wb, table));
  const dim = (list: [number, WDim][], into: Map<number, unknown>) => {
    for (const [i, wd] of list) {
      if (table && typeof wd.si === 'number') { const { si, ...rest } = wd; const info = fromWire(rest) as Record<string, unknown>; const id = table[si]; if (id) info.s = id; into.set(i, info); }
      else into.set(i, decodeDim(wd, wb));
    }
  };
  dim(w.rows, s.rows as Map<number, unknown>);
  dim(w.cols, s.cols as Map<number, unknown>);
  for (const [k, v] of Object.entries(w.parts)) {
    if (!safeKey(k) || SPECIAL_TOP.has(k)) continue;
    setSheetProp(s, k, fromWire(v));
  }
  if (!(s.notes instanceof Map)) s.notes = new Map();
  if (!(s.hyperlinks instanceof Map)) s.hyperlinks = new Map();
  if (s.comments !== undefined && !(s.comments instanceof Map)) delete s.comments;
  for (const k of ['charts', 'pictures', 'merges', 'conditionalFormats', 'validations', 'tables'] as const) if (!Array.isArray(s[k])) (s as unknown as Record<string, unknown>)[k] = [];
  return s;
};

// ---------- Workbook ----------

/** Workbook fields that aren't shared as book props. */
const BOOK_SKIP = new Set(['sheets', 'styles', 'activeSheet']);
export const bookPropKeys = (...books: (Workbook | undefined)[]): string[] => {
  const keys = new Set<string>();
  for (const b of books) if (b) for (const k of Object.keys(b)) if (!BOOK_SKIP.has(k) && safeKey(k)) keys.add(k);
  return [...keys];
};

export interface WBook { styles: Json; sheets: WSheet[]; book: Record<string, Json>; active: number }

export const encodeBook = (wb: Workbook): WBook => ({
  styles: toWire(wb.styles),
  sheets: wb.sheets.map((s) => encodeSheet(s, wb.styles, true)),
  book: Object.fromEntries(bookPropKeys(wb).filter((k) => (wb as unknown as Record<string, unknown>)[k] !== undefined).map((k) => [k, toWire((wb as unknown as Record<string, unknown>)[k])])),
  active: wb.activeSheet,
});

export const decodeBook = (w: WBook): Workbook => {
  const wb: Workbook = { sheets: [], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const styles = fromWire(w.styles);
  const table: number[] = [];
  if (Array.isArray(styles)) styles.forEach((st, i) => { table[i] = i === 0 ? 0 : internStyle(wb, (st && typeof st === 'object' ? st : {}) as CellStyle); });
  for (const [k, v] of Object.entries(w.book)) if (safeKey(k) && !BOOK_SKIP.has(k)) (wb as unknown as Record<string, unknown>)[k] = fromWire(v);
  if (!Array.isArray(wb.names)) wb.names = [];
  wb.date1904 = wb.date1904 === true;
  wb.sheets = w.sheets.map((s) => decodeSheet(s, wb, table));
  wb.activeSheet = Math.max(0, Math.min(wb.sheets.length - 1, w.active | 0));
  return wb;
};

// ---------- Hash ----------

/** FNV-1a over the shared content, to check that two copies agree. Formula results and personal view state are left out. */
export const bookHash = (wb: Workbook): string => {
  let h1 = 0x811c9dc5, h2 = 0x01000193 ^ 0x5bd1e995;
  const feed = (s: string) => {
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      h1 = Math.imul(h1 ^ c, 0x01000193);
      h2 = Math.imul(h2 ^ c, 0x5bd1e995) ^ (h2 >>> 15);
    }
    h1 ^= 0x1f; h2 ^= 0x2f;
  };
  const book: Record<string, unknown> = {};
  for (const k of bookPropKeys(wb)) if (k !== 'extraParts' && k !== 'props') book[k] = summarize((wb as unknown as Record<string, unknown>)[k]);
  feed(stable(book));
  for (const s of wb.sheets) {
    feed(s.name);
    const keys = [...s.cells.keys()].sort((a, b) => a - b);
    for (const k of keys) {
      const w = encodeCell(s.cells.get(k), wb.styles);
      if (w) { feed(String(k)); feed(stable(w)); }
    }
    for (const m of [s.rows, s.cols] as Map<number, { s?: number }>[]) {
      for (const i of [...m.keys()].sort((a, b) => a - b)) { const w = encodeDim(m.get(i), wb.styles); if (w) feed(`${i}:${stable(w)}`); }
      feed('|');
    }
    for (const k of KEYED) {
      const m = s[k] as Map<number, unknown> | undefined;
      if (m?.size) for (const key of [...m.keys()].sort((a, b) => a - b)) feed(`${k}${key}:${stable(summarize(m.get(key)))}`);
    }
    for (const k of ITEMS) { const v = s[k] as unknown[] | undefined; if (v?.length) feed(`${k}:${stable(summarize(v))}`); }
    for (const k of sheetPropKeys(s).sort()) { const v = sharedProp(s, k); if (v !== undefined && !(Array.isArray(v) && !v.length)) feed(`${k}:${stable(summarize(v))}`); }
  }
  return (h1 >>> 0).toString(36) + (h2 >>> 0).toString(36);
};

/** Bytes become their length and a short checksum; Maps become sorted entry lists. */
const summarize = (x: unknown, depth = 0): unknown => {
  if (depth > MAX_DEPTH || x === null || x === undefined || typeof x !== 'object') return x;
  if (x instanceof CellError) return { $e: x.code };
  if (x instanceof Uint8Array) { let h = 0; for (let i = 0; i < x.length; i += 97) h = (h * 31 + x[i]!) | 0; return { $b: x.length, h }; }
  if (x instanceof Map) return [...x].map(([k, v]) => [k, summarize(v, depth + 1)]).sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  if (Array.isArray(x)) return x.map((v) => summarize(v, depth + 1));
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(x)) if (v !== undefined) out[k] = summarize(v, depth + 1);
  return out;
};

export const inBounds = (row: number, col: number): boolean => Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0 && row < MAX_ROWS && col < MAX_COLS;
