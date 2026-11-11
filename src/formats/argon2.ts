// BLAKE2b (RFC 7693) and Argon2id (RFC 9106), for the newer OpenDocument key derivation.
// 64-bit words are kept as (low, high) pairs of 32-bit numbers.

const IV = new Uint32Array([
  0xf3bcc908, 0x6a09e667, 0x84caa73b, 0xbb67ae85, 0xfe94f82b, 0x3c6ef372, 0x5f1d36f1, 0xa54ff53a,
  0xade682d1, 0x510e527f, 0x2b3e6c1f, 0x9b05688c, 0xfb41bd6b, 0x1f83d9ab, 0x137e2179, 0x5be0cd19,
]);
const SIGMA = [
  [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15], [14, 10, 4, 8, 9, 15, 13, 6, 1, 12, 0, 2, 11, 7, 5, 3],
  [11, 8, 12, 0, 5, 2, 15, 13, 10, 14, 3, 6, 7, 1, 9, 4], [7, 9, 3, 1, 13, 12, 11, 14, 2, 6, 5, 10, 4, 0, 15, 8],
  [9, 0, 5, 7, 2, 4, 10, 15, 14, 1, 11, 12, 6, 8, 3, 13], [2, 12, 6, 10, 0, 11, 8, 3, 4, 13, 7, 5, 15, 14, 1, 9],
  [12, 5, 1, 15, 14, 13, 4, 10, 0, 7, 6, 3, 9, 2, 8, 11], [13, 11, 7, 14, 12, 1, 3, 9, 5, 0, 15, 4, 8, 6, 2, 10],
  [6, 15, 14, 9, 11, 3, 0, 8, 12, 2, 13, 7, 1, 4, 10, 5], [10, 2, 8, 4, 7, 6, 1, 5, 15, 11, 9, 14, 3, 12, 13, 0],
];

// v[2i] = low, v[2i+1] = high
const add64 = (v: Uint32Array, a: number, b: number): void => {
  const lo = v[a]! + v[b]!;
  v[a + 1] = (v[a + 1]! + v[b + 1]! + (lo > 0xffffffff ? 1 : 0)) >>> 0;
  v[a] = lo >>> 0;
};
const addM = (v: Uint32Array, a: number, lo: number, hi: number): void => {
  const l = v[a]! + lo;
  v[a + 1] = (v[a + 1]! + hi + (l > 0xffffffff ? 1 : 0)) >>> 0;
  v[a] = l >>> 0;
};
const xorRot = (v: Uint32Array, a: number, b: number, n: number): void => {
  const lo = v[a]! ^ v[b]!, hi = v[a + 1]! ^ v[b + 1]!;
  if (n === 32) { v[a] = hi; v[a + 1] = lo; return; }
  if (n < 32) { v[a] = ((lo >>> n) | (hi << (32 - n))) >>> 0; v[a + 1] = ((hi >>> n) | (lo << (32 - n))) >>> 0; return; }
  const m = n - 32;
  v[a] = ((hi >>> m) | (lo << (32 - m))) >>> 0; v[a + 1] = ((lo >>> m) | (hi << (32 - m))) >>> 0;
};

const blake2bCompress = (h: Uint32Array, block: Uint32Array, t: number, last: boolean): void => {
  const v = new Uint32Array(32);
  v.set(h); v.set(IV, 16);
  v[24]! ^= t >>> 0; v[25]! ^= Math.floor(t / 0x100000000) >>> 0;
  if (last) { v[28] = ~v[28]! >>> 0; v[29] = ~v[29]! >>> 0; }
  const g = (a: number, b: number, c: number, d: number, x: number, y: number) => {
    add64(v, a, b); addM(v, a, block[x * 2]!, block[x * 2 + 1]!); xorRot(v, d, a, 32);
    add64(v, c, d); xorRot(v, b, c, 24);
    add64(v, a, b); addM(v, a, block[y * 2]!, block[y * 2 + 1]!); xorRot(v, d, a, 16);
    add64(v, c, d); xorRot(v, b, c, 63);
  };
  for (let r = 0; r < 12; r++) {
    const s = SIGMA[r % 10]!;
    g(0, 8, 16, 24, s[0]!, s[1]!); g(2, 10, 18, 26, s[2]!, s[3]!); g(4, 12, 20, 28, s[4]!, s[5]!); g(6, 14, 22, 30, s[6]!, s[7]!);
    g(0, 10, 20, 30, s[8]!, s[9]!); g(2, 12, 22, 24, s[10]!, s[11]!); g(4, 14, 16, 26, s[12]!, s[13]!); g(6, 8, 18, 28, s[14]!, s[15]!);
  }
  for (let i = 0; i < 16; i++) h[i] = (h[i]! ^ v[i]! ^ v[i + 16]!) >>> 0;
};

/** BLAKE2b with an output of 1-64 bytes (no key). */
export const blake2b = (data: Uint8Array, outLen = 64): Uint8Array => {
  const h = IV.slice();
  h[0]! ^= 0x01010000 ^ outLen;
  const block = new Uint32Array(32);
  const buf = new Uint8Array(128);
  let t = 0;
  let off = 0;
  const load = () => { for (let i = 0; i < 32; i++) block[i] = (buf[i * 4]! | (buf[i * 4 + 1]! << 8) | (buf[i * 4 + 2]! << 16) | (buf[i * 4 + 3]! << 24)) >>> 0; };
  while (data.length - off > 128) {
    buf.set(data.subarray(off, off + 128)); off += 128; t += 128;
    load(); blake2bCompress(h, block, t, false);
  }
  buf.fill(0); buf.set(data.subarray(off)); t += data.length - off;
  load(); blake2bCompress(h, block, t, true);
  const out = new Uint8Array(outLen);
  for (let i = 0; i < outLen; i++) out[i] = (h[i >> 2]! >>> (8 * (i & 3))) & 0xff;
  return out;
};

const le32 = (n: number): Uint8Array => new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
const cat = (...p: Uint8Array[]): Uint8Array => { const o = new Uint8Array(p.reduce((n, x) => n + x.length, 0)); let i = 0; for (const x of p) { o.set(x, i); i += x.length; } return o; };

/** Variable-length hash H' from RFC 9106. */
const hPrime = (len: number, x: Uint8Array): Uint8Array => {
  const input = cat(le32(len), x);
  if (len <= 64) return blake2b(input, len);
  const out = new Uint8Array(len);
  let v = blake2b(input, 64);
  let off = 0;
  const r = Math.ceil(len / 32) - 2;
  for (let i = 0; i < r; i++) { out.set(v.subarray(0, 32), off); off += 32; if (i < r - 1) v = blake2b(v, 64); }
  out.set(blake2b(v, len - 32 * r), off);
  return out;
};

// ---------- Argon2 compression ----------

/** a += b + 2 * low32(a) * low32(b), on pairs in `w`. */
const fBlaMka = (w: Uint32Array, a: number, b: number): void => {
  const x = w[a]!, y = w[b]!;
  const xl = x & 0xffff, xh = x >>> 16, yl = y & 0xffff, yh = y >>> 16;
  const ll = xl * yl;
  const t = xl * yh + xh * yl + Math.floor(ll / 65536);
  let hi = xh * yh + Math.floor(t / 65536);
  let lo = Math.imul(x, y) >>> 0;
  hi = ((hi << 1) | (lo >>> 31)) >>> 0;
  lo = (lo << 1) >>> 0;
  add64(w, a, b);
  addM(w, a, lo, hi);
};
const gb = (w: Uint32Array, a: number, b: number, c: number, d: number): void => {
  fBlaMka(w, a, b); xorRot(w, d, a, 32);
  fBlaMka(w, c, d); xorRot(w, b, c, 24);
  fBlaMka(w, a, b); xorRot(w, d, a, 16);
  fBlaMka(w, c, d); xorRot(w, b, c, 63);
};
/** The permutation P over 16 words given by their pair offsets. */
const perm = (w: Uint32Array, o: number[]): void => {
  gb(w, o[0]!, o[4]!, o[8]!, o[12]!); gb(w, o[1]!, o[5]!, o[9]!, o[13]!); gb(w, o[2]!, o[6]!, o[10]!, o[14]!); gb(w, o[3]!, o[7]!, o[11]!, o[15]!);
  gb(w, o[0]!, o[5]!, o[10]!, o[15]!); gb(w, o[1]!, o[6]!, o[11]!, o[12]!); gb(w, o[2]!, o[7]!, o[8]!, o[13]!); gb(w, o[3]!, o[4]!, o[9]!, o[14]!);
};
const ROWS: number[][] = Array.from({ length: 8 }, (_, i) => Array.from({ length: 16 }, (_, j) => (i * 16 + j) * 2));
const COLS: number[][] = Array.from({ length: 8 }, (_, i) => Array.from({ length: 16 }, (_, j) => (((j >> 1) * 16) + i * 2 + (j & 1)) * 2));

const R = new Uint32Array(256), Q = new Uint32Array(256);
/** out = G(x, y) (xor out's old value when `xor`). Blocks are 256 32-bit words. */
const compressG = (mem: Uint32Array, x: number, y: number, out: number, xor: boolean): void => {
  for (let i = 0; i < 256; i++) R[i] = mem[x + i]! ^ mem[y + i]!;
  Q.set(R);
  for (const row of ROWS) perm(Q, row);
  for (const col of COLS) perm(Q, col);
  if (xor) for (let i = 0; i < 256; i++) mem[out + i] = (mem[out + i]! ^ Q[i]! ^ R[i]!) >>> 0;
  else for (let i = 0; i < 256; i++) mem[out + i] = (Q[i]! ^ R[i]!) >>> 0;
};

export interface Argon2Params { iterations: number; memoryKiB: number; lanes: number; length: number; secret?: Uint8Array; ad?: Uint8Array }

/** Argon2id, version 0x13. */
export const argon2id = (password: Uint8Array, salt: Uint8Array, p: Argon2Params): Uint8Array => {
  const lanes = Math.max(1, Math.min(255, p.lanes));
  const tCost = Math.max(1, Math.min(64, p.iterations));
  if (p.memoryKiB > 4 * 1024 * 1024) throw new Error('The key derivation asks for too much memory.');
  const mPrime = Math.max(8 * lanes, 4 * lanes * Math.floor(p.memoryKiB / (4 * lanes)));
  const q = mPrime / lanes;
  const seg = q / 4;
  const secret = p.secret ?? new Uint8Array(0), ad = p.ad ?? new Uint8Array(0);
  const h0 = blake2b(cat(le32(lanes), le32(p.length), le32(p.memoryKiB), le32(tCost), le32(0x13), le32(2),
    le32(password.length), password, le32(salt.length), salt, le32(secret.length), secret, le32(ad.length), ad), 64);
  const mem = new Uint32Array(mPrime * 256);
  const at = (lane: number, col: number) => (lane * q + col) * 256;
  const putBytes = (off: number, b: Uint8Array) => { for (let i = 0; i < 256; i++) mem[off + i] = (b[i * 4]! | (b[i * 4 + 1]! << 8) | (b[i * 4 + 2]! << 16) | (b[i * 4 + 3]! << 24)) >>> 0; };
  for (let l = 0; l < lanes; l++) {
    putBytes(at(l, 0), hPrime(1024, cat(h0, le32(0), le32(l))));
    putBytes(at(l, 1), hPrime(1024, cat(h0, le32(1), le32(l))));
  }
  // scratch blocks for address generation live after the memory
  const scratch = new Uint32Array(256 * 4);
  const zero = 0, input = 256, tmp = 512, addr = 768;
  const sg = (x: number, y: number, out: number) => {
    for (let i = 0; i < 256; i++) R[i] = scratch[x + i]! ^ scratch[y + i]!;
    Q.set(R);
    for (const row of ROWS) perm(Q, row);
    for (const col of COLS) perm(Q, col);
    for (let i = 0; i < 256; i++) scratch[out + i] = (Q[i]! ^ R[i]!) >>> 0;
  };
  for (let pass = 0; pass < tCost; pass++) {
    for (let slice = 0; slice < 4; slice++) {
      for (let lane = 0; lane < lanes; lane++) {
        const indep = pass === 0 && slice < 2;
        if (indep) {
          scratch.fill(0);
          scratch[input] = pass; scratch[input + 2] = lane; scratch[input + 4] = slice; scratch[input + 6] = mPrime; scratch[input + 8] = tCost; scratch[input + 10] = 2;
        }
        const start = pass === 0 && slice === 0 ? 2 : 0;
        if (indep && start) { scratch[input + 12]!++; sg(zero, input, tmp); sg(zero, tmp, addr); }
        for (let j = start; j < seg; j++) {
          const col = slice * seg + j;
          const prev = col === 0 ? q - 1 : col - 1;
          let j1: number, j2: number;
          if (indep) {
            if (j % 128 === 0 && j) { scratch[input + 12]!++; sg(zero, input, tmp); sg(zero, tmp, addr); }
            if (j % 128 === 0 && !j && !start) { scratch[input + 12]!++; sg(zero, input, tmp); sg(zero, tmp, addr); }
            j1 = scratch[addr + (j % 128) * 2]!; j2 = scratch[addr + (j % 128) * 2 + 1]!;
          } else {
            const pb = at(lane, prev);
            j1 = mem[pb]!; j2 = mem[pb + 1]!;
          }
          const refLane = pass === 0 && slice === 0 ? lane : j2 % lanes;
          let area: number;
          if (pass === 0) area = refLane === lane ? slice * seg + j - 1 : slice * seg - (j === 0 ? 1 : 0);
          else area = refLane === lane ? q - seg + j - 1 : q - seg - (j === 0 ? 1 : 0);
          // relative position: area - 1 - (area * (j1^2 >> 32) >> 32)
          const x = mulHi(j1, j1);
          const y = mulHi(area >>> 0, x);
          const rel = area - 1 - y;
          const startPos = pass === 0 ? 0 : ((slice + 1) * seg) % q;
          const refCol = (startPos + rel) % q;
          compressG(mem, at(lane, prev), at(refLane, refCol), at(lane, col), pass > 0);
        }
      }
    }
  }
  const c = new Uint32Array(256);
  for (let l = 0; l < lanes; l++) { const o = at(l, q - 1); for (let i = 0; i < 256; i++) c[i]! ^= mem[o + i]!; }
  const cb = new Uint8Array(1024);
  for (let i = 0; i < 256; i++) { cb[i * 4] = c[i]! & 255; cb[i * 4 + 1] = (c[i]! >>> 8) & 255; cb[i * 4 + 2] = (c[i]! >>> 16) & 255; cb[i * 4 + 3] = c[i]! >>> 24; }
  return hPrime(p.length, cb);
};

/** High 32 bits of a 32x32-bit product. */
const mulHi = (a: number, b: number): number => {
  const al = a & 0xffff, ah = a >>> 16, bl = b & 0xffff, bh = b >>> 16;
  const t = al * bh + ah * bl + Math.floor((al * bl) / 65536);
  return (ah * bh + Math.floor(t / 65536)) >>> 0;
};
