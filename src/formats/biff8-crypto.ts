// Password-based RC4 encryption of binary workbooks (MS-OFFCRYPTO 2.3.6 RC4 and 2.3.5 RC4 CryptoAPI,
// MS-XLS 2.2.10), plus the MD5, SHA-1 and RC4 primitives they need. Synchronous so readers stay simple.

export const md5 = (data: Uint8Array): Uint8Array => {
  const K = new Uint32Array(64);
  for (let i = 0; i < 64; i++) K[i] = Math.floor(Math.abs(Math.sin(i + 1)) * 2 ** 32) >>> 0;
  const S = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
  const len = data.length;
  const padded = new Uint8Array(((len + 8) >> 6) * 64 + 64);
  padded.set(data);
  padded[len] = 0x80;
  const dv = new DataView(padded.buffer);
  dv.setUint32(padded.length - 8, (len * 8) >>> 0, true);
  dv.setUint32(padded.length - 4, Math.floor(len / 0x20000000), true);
  let a0 = 0x67452301, b0 = 0xefcdab89, c0 = 0x98badcfe, d0 = 0x10325476;
  const M = new Uint32Array(16);
  for (let off = 0; off < padded.length; off += 64) {
    for (let i = 0; i < 16; i++) M[i] = dv.getUint32(off + i * 4, true);
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F: number, g: number;
      if (i < 16) { F = (B & C) | (~B & D); g = i; }
      else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) % 16; }
      else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) % 16; }
      else { F = C ^ (B | ~D); g = (7 * i) % 16; }
      F = (F + A + K[i]! + M[g]!) >>> 0;
      A = D; D = C; C = B;
      const s = S[(i >> 4) * 4 + (i % 4)]!;
      B = (B + ((F << s) | (F >>> (32 - s)))) >>> 0;
    }
    a0 = (a0 + A) >>> 0; b0 = (b0 + B) >>> 0; c0 = (c0 + C) >>> 0; d0 = (d0 + D) >>> 0;
  }
  const out = new Uint8Array(16);
  const o = new DataView(out.buffer);
  o.setUint32(0, a0, true); o.setUint32(4, b0, true); o.setUint32(8, c0, true); o.setUint32(12, d0, true);
  return out;
};

export const sha1 = (data: Uint8Array): Uint8Array => {
  const len = data.length;
  const padded = new Uint8Array(((len + 8) >> 6) * 64 + 64);
  padded.set(data);
  padded[len] = 0x80;
  const dv = new DataView(padded.buffer);
  dv.setUint32(padded.length - 4, (len * 8) >>> 0, false);
  dv.setUint32(padded.length - 8, Math.floor(len / 0x20000000), false);
  let h0 = 0x67452301, h1 = 0xefcdab89, h2 = 0x98badcfe, h3 = 0x10325476, h4 = 0xc3d2e1f0;
  const w = new Uint32Array(80);
  for (let off = 0; off < padded.length; off += 64) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(off + i * 4, false);
    for (let i = 16; i < 80; i++) { const x = w[i - 3]! ^ w[i - 8]! ^ w[i - 14]! ^ w[i - 16]!; w[i] = (x << 1) | (x >>> 31); }
    let a = h0, b = h1, c = h2, d = h3, e = h4;
    for (let i = 0; i < 80; i++) {
      let f: number, k: number;
      if (i < 20) { f = (b & c) | (~b & d); k = 0x5a827999; }
      else if (i < 40) { f = b ^ c ^ d; k = 0x6ed9eba1; }
      else if (i < 60) { f = (b & c) | (b & d) | (c & d); k = 0x8f1bbcdc; }
      else { f = b ^ c ^ d; k = 0xca62c1d6; }
      const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]!) >>> 0;
      e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = t;
    }
    h0 = (h0 + a) >>> 0; h1 = (h1 + b) >>> 0; h2 = (h2 + c) >>> 0; h3 = (h3 + d) >>> 0; h4 = (h4 + e) >>> 0;
  }
  const out = new Uint8Array(20);
  const o = new DataView(out.buffer);
  [h0, h1, h2, h3, h4].forEach((h, i) => o.setUint32(i * 4, h, false));
  return out;
};

export class Rc4 {
  private s = new Uint8Array(256);
  private i = 0;
  private j = 0;
  constructor(key: Uint8Array) {
    for (let k = 0; k < 256; k++) this.s[k] = k;
    let j = 0;
    for (let k = 0; k < 256; k++) {
      j = (j + this.s[k]! + key[k % key.length]!) & 0xff;
      const t = this.s[k]!; this.s[k] = this.s[j]!; this.s[j] = t;
    }
  }
  /** XORs the key stream into `buf` in place. */
  apply(buf: Uint8Array, from = 0, to = buf.length): void {
    const s = this.s;
    let { i, j } = this;
    for (let p = from; p < to; p++) {
      i = (i + 1) & 0xff;
      j = (j + s[i]!) & 0xff;
      const t = s[i]!; s[i] = s[j]!; s[j] = t;
      buf[p]! ^= s[(s[i]! + s[j]!) & 0xff]!;
    }
    this.i = i; this.j = j;
  }
  skip(n: number): void { this.apply(new Uint8Array(n)); }
}

const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};
const utf16 = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) { out[i * 2] = s.charCodeAt(i) & 0xff; out[i * 2 + 1] = s.charCodeAt(i) >> 8; }
  return out;
};
const le32 = (n: number): Uint8Array => new Uint8Array([n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff]);
const same = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((x, i) => x === b[i]);

/** Makes the RC4 key for one 1024-byte block of the stream. */
export type BlockKey = (block: number) => Uint8Array;

export interface Rc4Params { kind: 'rc4' | 'cryptoapi'; salt: Uint8Array; verifier: Uint8Array; verifierHash: Uint8Array; keyBits: number }

export const DEFAULT_PASSWORD = 'VelvetSweatshop';

const keyFor = (p: Rc4Params, password: string): BlockKey => {
  const pw = utf16(password.slice(0, 255));
  if (p.kind === 'rc4') {
    const h0 = md5(pw).subarray(0, 5);
    const buf = new Uint8Array(21 * 16);
    for (let i = 0; i < 16; i++) { buf.set(h0, i * 21); buf.set(p.salt, i * 21 + 5); }
    const h1 = md5(buf).subarray(0, 5);
    return (block) => md5(concat(h1, le32(block)));
  }
  const h0 = sha1(concat(p.salt, pw));
  const bytes = Math.max(5, Math.min(16, (p.keyBits || 40) / 8));
  return (block) => {
    const h = sha1(concat(h0, le32(block))).subarray(0, bytes);
    if (bytes === 5) { const k = new Uint8Array(16); k.set(h); return k; }
    return h;
  };
};

/** Returns the block-key function when the password is right, otherwise null. */
export const checkPassword = (p: Rc4Params, password: string): BlockKey | null => {
  const key = keyFor(p, password);
  const r = new Rc4(key(0));
  const v = p.verifier.slice();
  const h = p.verifierHash.slice();
  r.apply(v);
  r.apply(h);
  const want = p.kind === 'rc4' ? md5(v) : sha1(v);
  return same(want, h.subarray(0, want.length)) ? key : null;
};

/** Makes verifier fields for a password (used by the writer's tests and for re-encryption). */
export const makeRc4Params = (password: string, salt: Uint8Array, kind: Rc4Params['kind'] = 'cryptoapi'): Rc4Params => {
  const verifier = new Uint8Array(16);
  for (let i = 0; i < 16; i++) verifier[i] = (salt[i]! * 7 + 13) & 0xff;
  const p: Rc4Params = { kind, salt, verifier, verifierHash: new Uint8Array(kind === 'rc4' ? 16 : 20), keyBits: 128 };
  const key = keyFor(p, password);
  const v = verifier.slice();
  const h = kind === 'rc4' ? md5(verifier) : sha1(verifier);
  const r = new Rc4(key(0));
  r.apply(v);
  r.apply(h);
  return { ...p, verifier: v, verifierHash: h };
};

/** Position-based RC4 stream over a byte stream: rekeys every 1024 bytes. */
export class StreamCipher {
  private rc4: Rc4 | null = null;
  private block = -1;
  private pos = 0;
  constructor(private key: BlockKey) {}
  /** XORs the key stream for absolute positions [at, at + len) into buf[from..]. */
  apply(buf: Uint8Array, from: number, len: number, at: number): void {
    let done = 0;
    while (done < len) {
      const p = at + done;
      const block = Math.floor(p / 1024);
      if (block !== this.block || p < this.pos) {
        this.rc4 = new Rc4(this.key(block));
        this.block = block;
        this.pos = block * 1024;
      }
      if (p > this.pos) { this.rc4!.skip(p - this.pos); this.pos = p; }
      const n = Math.min(len - done, (block + 1) * 1024 - p);
      this.rc4!.apply(buf, from + done, from + done + n);
      this.pos += n;
      done += n;
    }
  }
}

// ---------- XOR obfuscation (MS-OFFCRYPTO 2.3.7, method 1), used by older workbooks ----------

const XOR_PAD = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00];
const XOR_INITIAL = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];
const XOR_MATRIX = [
  0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09, 0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf, 0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68,
  0x5ad0, 0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40, 0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496,
  0x390d, 0x721a, 0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9, 0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0, 0xb861, 0x60e3, 0xc1c6, 0x93ad,
  0x377b, 0x6ef6, 0xddec, 0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10, 0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168, 0x76b4, 0xed68, 0xcaf1,
  0x85c3, 0x1ba7, 0x374e, 0x6e9c, 0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc, 0x1021, 0x2042,
  0x4084, 0x8108, 0x1231, 0x2462, 0x48c4,
];
const xorBytes = (p: string): number[] => [...p.slice(0, 15)].map((c) => { const x = c.charCodeAt(0); return (x & 0xff) || (x >> 8); });

export const xorKey = (password: string): number => {
  const b = xorBytes(password);
  if (!b.length) return 0;
  let k = XOR_INITIAL[b.length - 1]!;
  let m = 0x68;
  for (let i = b.length - 1; i >= 0; i--) {
    let ch = b[i]!;
    for (let j = 0; j < 7; j++) { if (ch & 0x40) k ^= XOR_MATRIX[m]!; ch <<= 1; m--; }
  }
  return k;
};

export const xorVerifier = (password: string): number => {
  const b = xorBytes(password);
  let v = 0;
  for (let i = b.length - 1; i >= 0; i--) v = (((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ b[i]!;
  return (((v >> 14) & 1) | ((v << 1) & 0x7fff)) ^ b.length ^ 0xce4b;
};

/** The 16-byte array for a password, or null when it doesn't match the file's key and verifier. */
export const xorArray = (password: string, key: number, verifier: number): Uint8Array | null => {
  if (xorKey(password) !== key || xorVerifier(password) !== verifier) return null;
  const a = [...xorBytes(password), ...XOR_PAD].slice(0, 16);
  const ror1 = (x: number) => ((x >> 1) | (x << 7)) & 0xff;
  return Uint8Array.from(a.map((x, i) => ror1(x ^ (i % 2 ? key >> 8 : key & 0xff))));
};

/** Undoes XOR obfuscation of one record's data at stream position `pos`. */
export const xorDecode = (buf: Uint8Array, from: number, to: number, recordSize: number, arr: Uint8Array): void => {
  for (let p = from; p < to; p++) {
    const v = buf[p]! ^ arr[(p + recordSize) % 16]!;
    buf[p] = ((v >> 5) | (v << 3)) & 0xff;
  }
};

/** Applies XOR obfuscation (the inverse of xorDecode). */
export const xorEncode = (buf: Uint8Array, from: number, to: number, recordSize: number, arr: Uint8Array): void => {
  for (let p = from; p < to; p++) {
    const v = ((buf[p]! << 5) | (buf[p]! >> 3)) & 0xff;
    buf[p] = v ^ arr[(p + recordSize) % 16]!;
  }
};
