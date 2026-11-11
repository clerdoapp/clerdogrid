// base64url (RFC 4648 §5) and deflate helpers for codes, links and binary data in messages.
import { deflateSync, Inflate } from 'fflate';

const ALPHA = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
const LOOKUP = (() => {
  const t = new Int16Array(128).fill(-1);
  for (let i = 0; i < ALPHA.length; i++) t[ALPHA.charCodeAt(i)] = i;
  t['+'.charCodeAt(0)] = 62; t['/'.charCodeAt(0)] = 63;
  return t;
})();

export const b64urlEncode = (bytes: Uint8Array): string => {
  let out = '';
  const n = bytes.length;
  for (let i = 0; i < n; i += 3) {
    const a = bytes[i]!, b = i + 1 < n ? bytes[i + 1]! : 0, c = i + 2 < n ? bytes[i + 2]! : 0;
    const x = (a << 16) | (b << 8) | c;
    out += ALPHA[(x >> 18) & 63]! + ALPHA[(x >> 12) & 63]!;
    if (i + 1 < n) out += ALPHA[(x >> 6) & 63]!;
    if (i + 2 < n) out += ALPHA[x & 63]!;
  }
  return out;
};

/** Accepts base64url or standard base64, with or without padding. Throws on other characters. */
export const b64urlDecode = (text: string, maxBytes = Infinity): Uint8Array => {
  const s = text.replace(/[=\s]+$/g, '').replace(/\s+/g, '');
  if (s.length % 4 === 1) throw new Error('Bad code');
  const len = Math.floor((s.length * 3) / 4);
  if (len > maxBytes) throw new Error('Too large');
  const out = new Uint8Array(len);
  let o = 0, acc = 0, bits = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i);
    const v = ch < 128 ? LOOKUP[ch]! : -1;
    if (v < 0) throw new Error('Bad code');
    acc = (acc << 6) | v; bits += 6;
    if (bits >= 8) { bits -= 8; out[o++] = (acc >> bits) & 0xff; }
  }
  return out.subarray(0, o);
};

export const packText = (text: string): string => b64urlEncode(deflateSync(new TextEncoder().encode(text), { level: 9 }));

/** Inflates with a size cap so a tiny code can't expand into a huge string. */
export const unpackText = (code: string, maxBytes: number): string => {
  const raw = b64urlDecode(code, maxBytes);
  return new TextDecoder().decode(inflateCapped(raw, maxBytes));
};

export const packBytes = (bytes: Uint8Array): string => b64urlEncode(deflateSync(bytes, { level: 9 }));
export const unpackBytes = (code: string, maxBytes: number): Uint8Array => inflateCapped(b64urlDecode(code, maxBytes), maxBytes);

const inflateCapped = (raw: Uint8Array, maxBytes: number): Uint8Array => {
  // fflate's sync inflate has no size limit; the stream API lets us stop early.
  const parts: Uint8Array[] = [];
  let total = 0, over = false;
  const inf = new Inflate((chunk) => {
    total += chunk.length;
    if (total > maxBytes) { over = true; return; }
    parts.push(chunk);
  });
  const step = 64 * 1024;
  for (let i = 0; i < raw.length && !over; i += step) inf.push(raw.subarray(i, i + step), i + step >= raw.length);
  if (over) throw new Error('Too large');
  if (!raw.length) return new Uint8Array(0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};
