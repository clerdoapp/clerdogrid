// Every message from another device is checked here before anything uses it. Unknown op types are ignored;
// anything malformed throws and the whole message is dropped.
import { ERROR_CODES, MAX_COLS, MAX_ROWS, type ErrorCode } from '../../engine/types';
import type { Op } from './ops';
import { ITEMS, KEYED, MAX_DEPTH, safeKey, type WBook, type WCell, type WSheet } from './wire';

export const LIMITS = {
  /** Largest reassembled message (a full workbook snapshot). */
  message: 160 * 1024 * 1024,
  /** Largest single frame on the data channel. */
  frame: 256 * 1024,
  cellsPerOp: 2_000_000,
  opsPerMessage: 10_000,
  pending: 2_000,
  sheets: 1_000,
  text: 32_767,
  formula: 16_384,
  name: 80,
  picture: 40 * 1024 * 1024,
  ranges: 256,
  peers: 64,
};

function fail(what: string): never { throw new Error(`Invalid message: ${what}`); }
const isObj = (x: unknown): x is Record<string, unknown> => !!x && typeof x === 'object' && !Array.isArray(x);
const int = (x: unknown, min: number, max: number, what: string): number => (Number.isInteger(x) && (x as number) >= min && (x as number) <= max ? x as number : fail(what));
const str = (x: unknown, max: number, what: string): string => (typeof x === 'string' && x.length <= max ? x : fail(what));
const row = (x: unknown) => int(x, 0, MAX_ROWS - 1, 'row');
const col = (x: unknown) => int(x, 0, MAX_COLS - 1, 'column');
const onlyKeys = (o: Record<string, unknown>, keys: string[], what: string) => { for (const k of Object.keys(o)) if (!keys.includes(k)) fail(`${what} field ${k.slice(0, 20)}`); };

/** Sheet ids (sheets.ts) and sort keys. */
export const sheetId = (x: unknown): string => { const s = str(x, 40, 'sheet id'); if (!/^[A-Za-z0-9_-]{1,40}$/.test(s)) fail('sheet id'); return s; };
const sortKey = (x: unknown): string => { const s = str(x, 400, 'sheet position'); if (!/^[0-9a-z]*[1-9a-z]$/.test(s)) fail('sheet position'); return s; };

export const sheetName = (x: unknown): string => {
  const s = str(x, 31, 'sheet name');
  if (!s.trim() || /[\\/?*[\]:]/.test(s)) fail('sheet name');
  return s;
};

/** Tagged JSON of bounded depth; bytes only as base64url strings. */
export const json = (x: unknown, depth = 0): void => {
  if (depth > MAX_DEPTH) fail('nesting');
  if (x === null || typeof x === 'boolean') return;
  if (typeof x === 'number') { if (!Number.isFinite(x)) fail('number'); return; }
  if (typeof x === 'string') { if (x.length > LIMITS.message) fail('string'); return; }
  if (Array.isArray(x)) { for (const v of x) json(v, depth + 1); return; }
  if (isObj(x)) {
    for (const [k, v] of Object.entries(x)) { if (!safeKey(k)) fail('key'); json(v, depth + 1); }
    if ('$b' in x && (typeof x.$b !== 'string' || !/^[A-Za-z0-9_-]*$/.test(x.$b))) fail('bytes');
    if ('$e' in x && !ERROR_CODES.includes(x.$e as ErrorCode)) fail('error value');
    if ('$m' in x && !Array.isArray(x.$m)) fail('map');
    return;
  }
  fail('value');
};

const rng4 = (x: unknown): void => {
  if (!Array.isArray(x) || x.length !== 4) fail('range');
  const [r1, c1, r2, c2] = x as unknown[];
  if (row(r1) > row(r2) || col(c1) > col(c2)) fail('range');
};

export const wcell = (x: unknown, allowSi = false): WCell => {
  if (!isObj(x)) return fail('cell');
  onlyKeys(x, ['v', 'f', 'dyn', 'ar', 'st', 'si', 'rich'], 'cell');
  if ('v' in x) {
    const v = x.v;
    if (typeof v === 'string') str(v, LIMITS.text, 'text');
    else if (typeof v === 'number') { if (!Number.isFinite(v)) fail('number'); }
    else if (typeof v !== 'boolean') { if (!isObj(v) || !ERROR_CODES.includes(v.$e as ErrorCode) || Object.keys(v).length !== 1) fail('cell value'); }
  }
  if ('f' in x) str(x.f, LIMITS.formula, 'formula');
  if ('dyn' in x && x.dyn !== true) fail('dyn');
  if ('ar' in x) rng4(x.ar);
  if ('st' in x) { if (!isObj(x.st)) fail('style'); json(x.st); }
  if ('si' in x) { if (!allowSi) fail('style index'); int(x.si, 0, 10_000_000, 'style index'); }
  if ('rich' in x) {
    if (!Array.isArray(x.rich) || x.rich.length > 10_000) fail('rich text');
    for (const r of x.rich as unknown[]) { if (!isObj(r)) fail('rich text'); str(r.text, LIMITS.text, 'rich text'); json(r); }
  }
  return x as WCell;
};

const linkOk = (v: unknown): boolean => typeof v === 'string' && v.length <= 2083 && (/^(https:|mailto:)/i.test(v) || v.startsWith('#') || !/^[a-z][a-z0-9+.-]*:/i.test(v));

const dimInfo = (x: unknown, allowSi: boolean): void => {
  if (!isObj(x)) fail('row or column');
  json(x);
  for (const k of ['width', 'height']) if (k in x && (typeof x[k] !== 'number' || (x[k] as number) < 0 || (x[k] as number) > 100_000)) fail('size');
  if ('si' in x && (!allowSi || !Number.isInteger(x.si))) fail('style index');
  if ('st' in x && !isObj(x.st)) fail('style');
};

const SHEET_FIXED = new Set(['name', 'cells', 'rows', 'cols', ...KEYED]);

const pictureOk = (v: unknown): void => {
  if (!isObj(v)) return;
  if (typeof v.mime === 'string' && !/^image\/[a-z0-9.+-]+$/i.test(v.mime)) fail('picture type');
  const d = v.data;
  if (isObj(d) && typeof d.$b === 'string' && d.$b.length * 0.75 > LIMITS.picture) fail('picture size');
};

export const wsheet = (x: unknown, allowSi = false): WSheet => {
  if (!isObj(x)) return fail('sheet');
  sheetName(x.name);
  if (x.cid !== undefined) sheetId(x.cid);
  if (x.cpos !== undefined) sortKey(x.cpos);
  if (!Array.isArray(x.cells) || x.cells.length > LIMITS.cellsPerOp) fail('cells');
  for (const e of x.cells as unknown[]) { if (!Array.isArray(e) || e.length !== 3) fail('cell entry'); row(e[0]); col(e[1]); wcell(e[2], allowSi); }
  for (const [k, max] of [['rows', MAX_ROWS - 1], ['cols', MAX_COLS - 1]] as const) {
    const list = x[k];
    if (!Array.isArray(list) || list.length > 2_000_000) fail(k);
    for (const e of list as unknown[]) { if (!Array.isArray(e) || e.length !== 2) fail(k); int(e[0], 0, max, k); dimInfo(e[1], allowSi); }
  }
  if (!isObj(x.parts)) fail('sheet parts');
  const parts = x.parts as Record<string, unknown>;
  for (const [k, v] of Object.entries(parts)) {
    if (!safeKey(k) || ['name', 'cid', 'cpos', 'cells', 'rows', 'cols'].includes(k)) fail('sheet part');
    json(v);
  }
  keyedParts(parts);
  return x as unknown as WSheet;
};

const keyedParts = (parts: Record<string, unknown>) => {
  const hl = parts.hyperlinks;
  if (hl !== undefined) {
    if (!isObj(hl) || !Array.isArray(hl.$m)) fail('links');
    for (const e of (hl as { $m: unknown[] }).$m) if (!Array.isArray(e) || !linkOk(e[1])) fail('link');
  }
  const pics = parts.pictures;
  if (Array.isArray(pics)) pics.forEach(pictureOk);
};

export const wbook = (x: unknown): WBook => {
  if (!isObj(x)) return fail('workbook');
  json(x.styles);
  if (!Array.isArray(x.styles) || !Array.isArray(x.sheets) || !x.sheets.length || x.sheets.length > LIMITS.sheets) fail('workbook');
  for (const s of x.sheets as unknown[]) wsheet(s, true);
  if (!isObj(x.book)) fail('workbook');
  for (const [k, v] of Object.entries(x.book)) { if (!safeKey(k)) fail('workbook key'); json(v); }
  int(x.active, 0, LIMITS.sheets, 'active sheet');
  return x as unknown as WBook;
};

/** Returns the op, or null for an op type this version doesn't know (ignored). Throws when a known op is malformed. */
export const op = (x: unknown): Op | null => {
  if (!isObj(x) || typeof x.t !== 'string') return fail('op');
  const s = () => sheetId(x.s);
  switch (x.t) {
    case 'cells':
      s();
      if (!Array.isArray(x.c) || x.c.length > LIMITS.cellsPerOp) fail('cells');
      for (const e of x.c as unknown[]) { if (!Array.isArray(e) || e.length !== 3) fail('cell entry'); row(e[0]); col(e[1]); if (e[2] !== null) wcell(e[2]); }
      break;
    case 'keyed':
      s();
      if (!(KEYED as readonly unknown[]).includes(x.k) || !Array.isArray(x.c) || x.c.length > LIMITS.cellsPerOp) fail('keyed');
      for (const e of x.c as unknown[]) {
        if (!Array.isArray(e) || e.length !== 3) fail('keyed entry');
        row(e[0]); col(e[1]); json(e[2]);
        if (e[2] !== null && x.k === 'hyperlinks' && !linkOk(e[2])) fail('link');
        if (e[2] !== null && x.k !== 'hyperlinks' && !isObj(e[2])) fail('keyed value');
      }
      break;
    case 'dim':
      s();
      if ((x.a !== 'row' && x.a !== 'col') || !Array.isArray(x.c) || x.c.length > 2_000_000) fail('dim');
      for (const e of x.c as unknown[]) { if (!Array.isArray(e) || e.length !== 2) fail('dim entry'); int(e[0], 0, x.a === 'row' ? MAX_ROWS - 1 : MAX_COLS - 1, 'index'); if (e[1] !== null) dimInfo(e[1], false); }
      break;
    case 'item':
      s();
      if (!(ITEMS as readonly unknown[]).includes(x.k)) fail('item kind');
      str(x.id, 128, 'item id'); int(x.i, 0, 1_000_000, 'item index');
      if (x.v !== null) { if (!isObj(x.v)) fail('item'); json(x.v); if (x.k === 'pictures') pictureOk(x.v); }
      break;
    case 'prop':
      s();
      if (typeof x.k !== 'string' || !/^[A-Za-z][A-Za-z0-9]{0,39}$/.test(x.k) || SHEET_FIXED.has(x.k)) fail('prop');
      json(x.v);
      if (x.k === 'pictures' && Array.isArray(x.v)) x.v.forEach(pictureOk);
      break;
    case 'struct': {
      s();
      if (x.a !== 'row' && x.a !== 'col') fail('axis');
      const max = x.a === 'row' ? MAX_ROWS : MAX_COLS;
      sheetName(x.nm);
      int(x.at, 0, max - 1, 'position');
      const n = int(x.n, -max, max, 'count');
      if (n === 0) fail('count');
      break;
    }
    case 'sheetAdd': s(); sheetName(x.name); sortKey(x.pos); wsheet(x.d); break;
    case 'sheetDel': s(); sheetName(x.name); break;
    case 'sheetRen': s(); sheetName(x.from); sheetName(x.to); break;
    case 'sheetPos': s(); sortKey(x.pos); break;
    case 'book':
      if (typeof x.k !== 'string' || !/^[A-Za-z][A-Za-z0-9]{0,39}$/.test(x.k) || x.k === 'sheets' || x.k === 'styles' || x.k === 'activeSheet' || x.k === 'extraParts') fail('book prop');
      json(x.v);
      break;
    default: return null;
  }
  return x as unknown as Op;
};

export const ops = (x: unknown): Op[] => {
  if (!Array.isArray(x) || x.length > LIMITS.opsPerMessage) return fail('ops');
  return x.map(op).filter((o): o is Op => o !== null);
};

// ---------- Messages ----------

export interface PeerInfo { id: string; name: string; color: string; ro?: boolean; away?: boolean; host?: boolean }
export interface Presence { id: string; sheet: string; ranges: [number, number, number, number][]; active: [number, number]; editing: [number, number] | null }

export type Msg =
  | { t: 'hello'; proto: number; id: string; name: string; resume?: { acked: number; pending: { id: number; base: number; ops: Op[] }[] } }
  | { t: 'welcome'; you: string; color: string; seq: number; book: WBook; name: string; ro: boolean; roster: PeerInfo[] }
  | { t: 'op'; id: number; base?: number; seq?: number; origin?: string; ops: Op[] }
  | { t: 'presence'; p: Presence }
  | { t: 'roster'; peers: PeerInfo[] }
  | { t: 'snap'; seq: number; book: WBook; name: string }
  | { t: 'hash'; seq: number; h: string }
  | { t: 'resync' } | { t: 'bye' } | { t: 'end' } | { t: 'ping' } | { t: 'pong' };

export const peerId = (x: unknown): string => { const s = str(x, 64, 'peer id'); if (!/^[A-Za-z0-9_-]{1,64}$/.test(s)) fail('peer id'); return s; };
const color = (x: unknown): string => { const s = str(x, 9, 'color'); if (!/^#[0-9a-fA-F]{6}$/.test(s)) fail('color'); return s; };
const personName = (x: unknown): string => str(x, LIMITS.name, 'name');

const peerInfo = (x: unknown): PeerInfo => {
  if (!isObj(x)) return fail('peer');
  peerId(x.id); personName(x.name); color(x.color);
  for (const k of ['ro', 'away', 'host']) if (k in x && typeof x[k] !== 'boolean') fail('peer');
  return { id: x.id as string, name: x.name as string, color: x.color as string, ro: x.ro === true || undefined, away: x.away === true || undefined, host: x.host === true || undefined };
};

const presence = (x: unknown): Presence => {
  if (!isObj(x)) return fail('presence');
  peerId(x.id); sheetId(x.sheet);
  if (!Array.isArray(x.ranges) || !x.ranges.length || x.ranges.length > LIMITS.ranges) fail('presence ranges');
  (x.ranges as unknown[]).forEach(rng4);
  const pos = (p: unknown) => { if (!Array.isArray(p) || p.length !== 2) fail('position'); row(p[0]); col(p[1]); };
  pos(x.active);
  if (x.editing !== null) pos(x.editing);
  return x as unknown as Presence;
};

/** Checks a parsed message. Throws when it is malformed. */
export const message = (x: unknown): Msg => {
  if (!isObj(x) || typeof x.t !== 'string') return fail('message');
  switch (x.t) {
    case 'hello': {
      int(x.proto, 1, 1000, 'protocol'); peerId(x.id); personName(x.name);
      if (x.resume !== undefined) {
        const r = x.resume;
        if (!isObj(r)) fail('resume');
        int((r as Record<string, unknown>).acked, 0, Number.MAX_SAFE_INTEGER, 'resume');
        const p = (r as Record<string, unknown>).pending;
        if (!Array.isArray(p) || p.length > LIMITS.pending) fail('resume');
        for (const e of p as unknown[]) {
          if (!isObj(e)) fail('resume');
          int(e.id, 1, Number.MAX_SAFE_INTEGER, 'op id'); int(e.base, 0, Number.MAX_SAFE_INTEGER, 'op base');
          e.ops = ops(e.ops);
        }
      }
      break;
    }
    case 'welcome':
      peerId(x.you); color(x.color); int(x.seq, 0, Number.MAX_SAFE_INTEGER, 'seq'); wbook(x.book); str(x.name, 255, 'file name');
      if (typeof x.ro !== 'boolean' || !Array.isArray(x.roster) || x.roster.length > LIMITS.peers) fail('welcome');
      x.roster = (x.roster as unknown[]).map(peerInfo);
      break;
    case 'op':
      int(x.id, 1, Number.MAX_SAFE_INTEGER, 'op id');
      if (x.base !== undefined) int(x.base, 0, Number.MAX_SAFE_INTEGER, 'op base');
      if (x.seq !== undefined) int(x.seq, 1, Number.MAX_SAFE_INTEGER, 'seq');
      if (x.origin !== undefined) peerId(x.origin);
      x.ops = ops(x.ops);
      break;
    case 'presence': presence(x.p); break;
    case 'roster':
      if (!Array.isArray(x.peers) || x.peers.length > LIMITS.peers) fail('roster');
      x.peers = (x.peers as unknown[]).map(peerInfo);
      break;
    case 'snap': int(x.seq, 0, Number.MAX_SAFE_INTEGER, 'seq'); wbook(x.book); str(x.name, 255, 'file name'); break;
    case 'hash': int(x.seq, 0, Number.MAX_SAFE_INTEGER, 'seq'); str(x.h, 40, 'hash'); break;
    case 'resync': case 'bye': case 'end': case 'ping': case 'pong': break;
    default: return fail('message type');
  }
  return x as unknown as Msg;
};
