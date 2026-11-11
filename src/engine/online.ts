// Internet access for worksheet functions (WEBSERVICE, IMAGE, STOCKHISTORY). Off unless the person turns
// on "Allow functions to use the internet"; while off these functions return #CONNECT! and nothing is
// fetched. Requests are https only, without cookies or referrer, with size and time limits. Results are
// fetched in the background: the formula shows #BUSY! until the answer arrives, then the workbook
// recalculates (the functions are volatile, so they read the cached answer on the next pass).

import { CellError, type Matrix, type Scalar } from './types';

export const ONLINE_KEY = 'clerdogrid.online';

export const busy = (): CellError => new CellError('#BUSY!', 'Waiting for an answer from the internet.');
export const offline = (): CellError => new CellError('#CONNECT!', 'Turn on "Allow functions to use the internet" to use this function.');

let allowed = false;
const settleListeners = new Set<() => void>();
const changeListeners = new Set<() => void>();

export const onlineAllowed = (): boolean => allowed;

// The desktop app relays requests through its main process (so cross-site rules don't block them); the
// web build fetches directly.
interface Relay {
  netSetAllowed(on: boolean): Promise<boolean>;
  netFetch(req: { url: string; kind: 'text' | 'binary'; allowed: boolean; maxBytes?: number; timeoutMs?: number }):
    Promise<{ ok: true; status: number; mime: string; text?: string; data?: Uint8Array } | { ok: false; error: string }>;
}
const relay = (): Relay | null => {
  const c = (globalThis as { clerdo?: Partial<Relay> }).clerdo;
  return c && typeof c.netFetch === 'function' && typeof c.netSetAllowed === 'function' ? (c as Relay) : null;
};
const tellRelay = (): void => { void relay()?.netSetAllowed(allowed).catch(() => undefined); };

/** Turns internet access for functions on or off (and remembers it on this device when `persist`). */
export const setOnlineAllowed = (on: boolean, persist = false): void => {
  allowed = on;
  tellRelay();
  if (!on) cache.clear();
  if (persist) { try { localStorage.setItem(ONLINE_KEY, on ? '1' : '0'); } catch { /* private mode */ } }
  for (const l of changeListeners) l();
};

export const initOnline = (): void => {
  try { allowed = localStorage.getItem(ONLINE_KEY) === '1'; } catch { allowed = false; }
  tellRelay();
};

/** Called (once per batch) when background answers arrive, so the app can recalculate and redraw. */
export const onOnlineSettled = (l: () => void): (() => void) => { settleListeners.add(l); return () => settleListeners.delete(l); };
export const onOnlineChanged = (l: () => void): (() => void) => { changeListeners.add(l); return () => changeListeners.delete(l); };

type Answer = Scalar | Matrix;
interface Entry { done: boolean; value: Answer; at: number; ttl: number; promise?: Promise<void> }
const cache = new Map<string, Entry>();
let settleTimer: ReturnType<typeof setTimeout> | null = null;
const settle = (): void => {
  if (settleTimer) return;
  settleTimer = setTimeout(() => { settleTimer = null; for (const l of settleListeners) l(); }, 0);
};

/**
 * The cached answer for `key`, or #BUSY! while `load` runs. A stale answer is shown while it refreshes.
 * `load` must resolve to a value (errors as CellError); a rejection becomes `onError`.
 */
export const cachedFetch = (key: string, ttlMs: number, load: () => Promise<Answer>, onError: () => Answer = () => new CellError('#VALUE!')): Answer => {
  const e = cache.get(key);
  const now = Date.now();
  if (e && (!e.done || now - e.at < e.ttl)) return e.done ? e.value : (e.value ?? busy());
  const entry: Entry = { done: false, value: e?.done ? e.value : busy(), at: now, ttl: ttlMs };
  cache.set(key, entry);
  if (cache.size > 2000) { const first = cache.keys().next().value; if (first !== undefined && first !== key) cache.delete(first); }
  entry.promise = load().then((v) => v, () => onError()).then((v) => {
    entry.done = true; entry.value = v; entry.at = Date.now();
    settle();
  });
  return entry.value;
};

/** Waits for every request in flight (for tests and for "Refresh"). */
export const settled = async (): Promise<void> => { await Promise.all([...cache.values()].map((e) => e.promise)); };
export const clearOnlineCache = (): void => cache.clear();

export interface FetchLimits { maxBytes: number; timeoutMs: number }

/** Validates a URL for fetching: absolute https with a host, no credentials. */
export const safeHttpsUrl = (raw: string): URL | null => {
  if (raw.length > 2048) return null;
  try {
    const u = new URL(raw);
    if (u.protocol !== 'https:' || !u.hostname || u.username || u.password) return null;
    return u;
  } catch { return null; }
};

/** Fetches text over https with a size and time limit. Rejects on any failure. */
export const fetchText = async (url: URL, limits: FetchLimits): Promise<string> => {
  const r = relay();
  if (r) {
    if (!allowed) throw new Error('offline');
    const res = await r.netFetch({ url: url.href, kind: 'text', allowed, maxBytes: limits.maxBytes, timeoutMs: limits.timeoutMs });
    if (!res.ok || typeof res.text !== 'string') throw new Error(res.ok ? 'no text' : res.error);
    return res.text;
  }
  const f = (globalThis as { fetch?: typeof fetch }).fetch;
  if (!f) throw new Error('no fetch');
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), limits.timeoutMs);
  try {
    const res = await f(url.href, { credentials: 'omit', referrerPolicy: 'no-referrer', redirect: 'follow', signal: ctrl.signal, cache: 'no-store' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    if (res.url && !res.url.startsWith('https:')) throw new Error('redirected away from https');
    const len = Number(res.headers.get('content-length') ?? '0');
    if (len > limits.maxBytes) throw new Error('too large');
    if (!res.body) { const t = await res.text(); if (t.length > limits.maxBytes) throw new Error('too large'); return t; }
    const reader = res.body.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > limits.maxBytes) { void reader.cancel(); throw new Error('too large'); }
      chunks.push(value);
    }
    const all = new Uint8Array(total);
    let o = 0;
    for (const c of chunks) { all.set(c, o); o += c.byteLength; }
    return new TextDecoder('utf-8').decode(all);
  } finally { clearTimeout(timer); }
};

/**
 * Image bytes through the desktop relay, or null when there is no relay (the page then loads the
 * picture itself). Rejects when the relay refuses.
 */
export const fetchImageViaRelay = async (url: string): Promise<{ data: Uint8Array; mime: string } | null> => {
  const r = relay();
  if (!r) return null;
  if (!allowed) throw new Error('offline');
  const u = safeHttpsUrl(url);
  if (!u) throw new Error('bad address');
  const res = await r.netFetch({ url: u.href, kind: 'binary', allowed, maxBytes: 10_000_000, timeoutMs: 20_000 });
  if (!res.ok || !res.data) throw new Error(res.ok ? 'no data' : res.error);
  return { data: res.data, mime: res.mime };
};

// ---------- IMAGE cells ----------

export interface CellImage {
  url: string;
  alt: string;
  /** 0 fit keeping proportions, 1 fill the cell, 2 original size, 3 custom height and width. */
  sizing: number;
  height?: number;
  width?: number;
  /** The formula cell that made it (a spilled IMAGE array registers every cell it covers). */
  anchor: { sheet: number; row: number; col: number };
}

const images = new Map<string, CellImage>();
const imgKey = (sheet: number, row: number, col: number): string => `${sheet}:${row}:${col}`;

export const registerCellImage = (sheet: number, row: number, col: number, img: CellImage): void => {
  if (images.size > 50_000) images.clear();
  images.set(imgKey(sheet, row, col), img);
};
export const cellImage = (sheet: number, row: number, col: number): CellImage | undefined => images.get(imgKey(sheet, row, col));
export const clearCellImages = (): void => images.clear();
