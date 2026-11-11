// Fetches for worksheet functions (WEBSERVICE, STOCKHISTORY, IMPORTCSV, IMPORTTEXT, IMAGE), done in the
// main process so the page's cross-site rules don't block them. Every request is checked here: the
// person must have allowed internet access (this process keeps its own copy of that setting, and the page
// must also say so), https only, no credentials or cookies, no local or private-network hosts, redirects
// only to addresses that pass the same checks, and caps on size, time and requests in flight.
// Host names are checked again at connection time: the default fetch resolves names itself and refuses
// to connect when any resolved address is private, so a name can't be pointed at the local network.
// No Electron imports, so it can be tested on its own.

import { request } from 'node:https';
import { lookup as dnsLookup, type LookupAddress } from 'node:dns';
import { isIP } from 'node:net';
import { Readable } from 'node:stream';

export interface RelayRequest { url: string; kind: 'text' | 'binary'; allowed: boolean; maxBytes?: number; timeoutMs?: number }
export type RelayResponse =
  | { ok: true; status: number; mime: string; text?: string; data?: Uint8Array; url: string }
  | { ok: false; error: string };

export const LIMITS = { maxBytes: 20_000_000, timeoutMs: 30_000, redirects: 5, inFlight: 8 };

let allowed = false;
export const setRelayAllowed = (on: unknown): boolean => { allowed = on === true; return allowed; };
export const relayAllowed = (): boolean => allowed;

const PRIVATE_V4 = [/^0\./, /^10\./, /^127\./, /^169\.254\./, /^172\.(1[6-9]|2\d|3[01])\./, /^192\.168\./, /^100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\./, /^22[4-9]\.|^2[3-5]\d\./];

/** True for an IP address on the public internet. */
export const isPublicAddress = (raw: string): boolean => {
  const a = raw.toLowerCase().replace(/^\[|\]$/g, '').replace(/%.*$/, '');
  const v = isIP(a);
  if (v === 4) return !PRIVATE_V4.some((re) => re.test(a)) && a !== '255.255.255.255';
  if (v !== 6) return false;
  // An IPv4 address carried in the last 32 bits, written dotted or as two hex groups.
  const tail = (prefix: string): string | null => {
    if (!a.startsWith(prefix)) return null;
    const rest = a.slice(prefix.length);
    if (isIP(rest) === 4) return rest;
    const m = /^([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(rest);
    if (!m) return null;
    const hi = parseInt(m[1]!, 16), lo = parseInt(m[2]!, 16);
    return `${hi >> 8}.${hi & 255}.${lo >> 8}.${lo & 255}`;
  };
  const mapped = tail('::ffff:') ?? tail('::ffff:0:');
  if (mapped) return isPublicAddress(mapped);
  if (a === '::1' || a === '::' || /^::ffff:/.test(a)) return false;
  // Link-local fe80::/10, unique-local fc00::/7, multicast ff00::/8, documentation 2001:db8::/32, NAT64 64:ff9b::/96 to private v4.
  if (/^(fe[89ab]|f[cd]|ff)/.test(a) || a.startsWith('2001:db8:')) return false;
  const nat64 = tail('64:ff9b::');
  if (nat64) return isPublicAddress(nat64);
  return true;
};

type LookupCb = (err: NodeJS.ErrnoException | null, address?: string | LookupAddress[], family?: number) => void;

/** DNS lookup that fails when a name resolves to any non-public address. */
export const safeLookup = (host: string, opts: { all?: boolean; family?: number } | number, cb: LookupCb, resolve = dnsLookup): void => {
  const o = typeof opts === 'number' ? { family: opts } : opts ?? {};
  resolve(host, { family: o.family ?? 0, all: true }, (err, addrs) => {
    if (err) { cb(err); return; }
    const list = addrs as unknown as LookupAddress[];
    if (!list.length || list.some((x) => !isPublicAddress(x.address))) {
      cb(Object.assign(new Error('address not allowed'), { code: 'EADDRBLOCKED' }));
      return;
    }
    if (o.all) cb(null, list);
    else cb(null, list[0]!.address, list[0]!.family);
  });
};

/** A fetch over node:https that connects only through safeLookup and never follows redirects or sends cookies. */
export const pinnedFetch = (href: string, init: RequestInit = {}): Promise<Response> => new Promise((resolve, reject) => {
  const headers: Record<string, string> = {};
  new Headers(init.headers).forEach((v, k) => { headers[k] = v; });
  const req = request(href, { method: 'GET', headers, lookup: safeLookup as never, agent: false, signal: init.signal ?? undefined }, (res) => {
    const h = new Headers();
    for (const [k, v] of Object.entries(res.headers)) if (v !== undefined) h.set(k, Array.isArray(v) ? v.join(', ') : String(v));
    const status = res.statusCode ?? 502;
    if (status < 200 || status > 599) { res.destroy(); reject(new Error('bad status')); return; }
    const body = status === 204 || status === 205 || status === 304 ? null : (Readable.toWeb(res) as unknown as ReadableStream<Uint8Array>);
    if (!body) res.resume();
    resolve(new Response(body, { status, headers: h }));
  });
  req.on('error', reject);
  req.end();
});

/** The URL if it is safe to request: absolute https, a public host name or address, no user or password. */
export const checkUrl = (raw: unknown): URL | null => {
  if (typeof raw !== 'string' || raw.length > 2048) return null;
  let u: URL;
  try { u = new URL(raw); } catch { return null; }
  if (u.protocol !== 'https:' || u.username || u.password) return null;
  if (u.port && u.port !== '443') return null;
  const host = u.hostname.toLowerCase().replace(/^\[|\]$/g, '');
  if (!host || host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local') || host.endsWith('.internal')) return null;
  if (isIP(host) && !isPublicAddress(host)) return null;
  if (/^\d+$/.test(host) || /^0x/i.test(host)) return null;
  return u;
};

/** Validates the shape of a request coming from the page. */
export const checkRequest = (req: unknown): { url: URL; kind: 'text' | 'binary'; maxBytes: number; timeoutMs: number } | string => {
  if (!req || typeof req !== 'object') return 'bad request';
  const r = req as Partial<RelayRequest>;
  if (r.allowed !== true || !allowed) return 'internet access is off';
  if (r.kind !== 'text' && r.kind !== 'binary') return 'bad request';
  const url = checkUrl(r.url);
  if (!url) return 'address not allowed';
  const num = (v: unknown, cap: number, dflt: number): number => (typeof v === 'number' && Number.isFinite(v) && v > 0 ? Math.min(cap, Math.floor(v)) : dflt);
  return { url, kind: r.kind, maxBytes: num(r.maxBytes, LIMITS.maxBytes, 2_000_000), timeoutMs: num(r.timeoutMs, LIMITS.timeoutMs, 15_000) };
};

let inFlight = 0;

/** Runs one checked request with the given fetch (the platform's by default). */
export const relayFetch = async (req: unknown, f: typeof fetch = pinnedFetch as typeof fetch): Promise<RelayResponse> => {
  const c = checkRequest(req);
  if (typeof c === 'string') return { ok: false, error: c };
  if (inFlight >= LIMITS.inFlight) return { ok: false, error: 'too many requests' };
  inFlight++;
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), c.timeoutMs);
  try {
    let url = c.url;
    let res: Response | null = null;
    for (let hop = 0; hop <= LIMITS.redirects; hop++) {
      res = await f(url.href, {
        method: 'GET', redirect: 'manual', signal: ctrl.signal, credentials: 'omit', referrerPolicy: 'no-referrer',
        headers: { accept: c.kind === 'binary' ? 'image/*' : '*/*' },
      });
      if (res.status >= 300 && res.status < 400) {
        const loc = res.headers.get('location');
        let next: URL | null = null;
        try { next = loc ? checkUrl(new URL(loc, url).href) : null; } catch { next = null; }
        if (!next) return { ok: false, error: 'redirect not allowed' };
        url = next;
        res = null;
        continue;
      }
      break;
    }
    if (!res) return { ok: false, error: 'too many redirects' };
    if (!res.ok) return { ok: false, error: `HTTP ${res.status}` };
    const mime = (res.headers.get('content-type') ?? '').split(';')[0]!.trim().toLowerCase();
    if (c.kind === 'binary' && !mime.startsWith('image/')) return { ok: false, error: 'not an image' };
    const len = Number(res.headers.get('content-length') ?? '0');
    if (len > c.maxBytes) return { ok: false, error: 'too large' };
    const chunks: Uint8Array[] = [];
    let total = 0;
    if (res.body) {
      const reader = res.body.getReader();
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        total += value.byteLength;
        if (total > c.maxBytes) { void reader.cancel(); return { ok: false, error: 'too large' }; }
        chunks.push(value);
      }
    }
    const data = new Uint8Array(total);
    let o = 0;
    for (const ch of chunks) { data.set(ch, o); o += ch.byteLength; }
    if (c.kind === 'binary') return { ok: true, status: res.status, mime, data, url: url.href };
    return { ok: true, status: res.status, mime, text: new TextDecoder('utf-8').decode(data), url: url.href };
  } catch (e) {
    const err = e as NodeJS.ErrnoException & { cause?: NodeJS.ErrnoException };
    if (err.name === 'AbortError') return { ok: false, error: 'timed out' };
    if (err.code === 'EADDRBLOCKED' || err.cause?.code === 'EADDRBLOCKED') return { ok: false, error: 'address not allowed' };
    return { ok: false, error: 'network error' };
  } finally {
    clearTimeout(timer);
    inFlight--;
  }
};
