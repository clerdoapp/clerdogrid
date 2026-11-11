// Synchronous hash and block-cipher primitives for workbook passwords and encrypted files.
// SHA-1, SHA-256 and SHA-512 follow FIPS 180-4, AES follows FIPS 197. They run synchronously
// because password hashing repeats a hash up to 100,000 times, which is slow through async APIs.

export type HashName = 'SHA-1' | 'SHA-256' | 'SHA-384' | 'SHA-512';

const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};
export { concat };

// ---------- SHA-512 / SHA-384 ----------

const K512 = (() => {
  const hex = [
    '428a2f98d728ae22', '7137449123ef65cd', 'b5c0fbcfec4d3b2f', 'e9b5dba58189dbbc', '3956c25bf348b538', '59f111f1b605d019', '923f82a4af194f9b', 'ab1c5ed5da6d8118',
    'd807aa98a3030242', '12835b0145706fbe', '243185be4ee4b28c', '550c7dc3d5ffb4e2', '72be5d74f27b896f', '80deb1fe3b1696b1', '9bdc06a725c71235', 'c19bf174cf692694',
    'e49b69c19ef14ad2', 'efbe4786384f25e3', '0fc19dc68b8cd5b5', '240ca1cc77ac9c65', '2de92c6f592b0275', '4a7484aa6ea6e483', '5cb0a9dcbd41fbd4', '76f988da831153b5',
    '983e5152ee66dfab', 'a831c66d2db43210', 'b00327c898fb213f', 'bf597fc7beef0ee4', 'c6e00bf33da88fc2', 'd5a79147930aa725', '06ca6351e003826f', '142929670a0e6e70',
    '27b70a8546d22ffc', '2e1b21385c26c926', '4d2c6dfc5ac42aed', '53380d139d95b3df', '650a73548baf63de', '766a0abb3c77b2a8', '81c2c92e47edaee6', '92722c851482353b',
    'a2bfe8a14cf10364', 'a81a664bbc423001', 'c24b8b70d0f89791', 'c76c51a30654be30', 'd192e819d6ef5218', 'd69906245565a910', 'f40e35855771202a', '106aa07032bbd1b8',
    '19a4c116b8d2d0c8', '1e376c085141ab53', '2748774cdf8eeb99', '34b0bcb5e19b48a8', '391c0cb3c5c95a63', '4ed8aa4ae3418acb', '5b9cca4f7763e373', '682e6ff3d6b2b8a3',
    '748f82ee5defb2fc', '78a5636f43172f60', '84c87814a1f0ab72', '8cc702081a6439ec', '90befffa23631e28', 'a4506cebde82bde9', 'bef9a3f7b2c67915', 'c67178f2e372532b',
    'ca273eceea26619c', 'd186b8c721c0c207', 'eada7dd6cde0eb1e', 'f57d4f7fee6ed178', '06f067aa72176fba', '0a637dc5a2c898a6', '113f9804bef90dae', '1b710b35131c471b',
    '28db77f523047d84', '32caab7b40c72493', '3c9ebe0a15c9bebc', '431d67c49c100d4c', '4cc5d4becb3e42b6', '597f299cfc657e2a', '5fcb6fab3ad6faec', '6c44198c4a475817',
  ];
  const out = new Uint32Array(160);
  hex.forEach((h, i) => { out[i * 2] = parseInt(h.slice(0, 8), 16); out[i * 2 + 1] = parseInt(h.slice(8), 16); });
  return out;
})();

const IV512 = [0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1, 0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179];
const IV384 = [0xcbbb9d5d, 0xc1059ed8, 0x629a292a, 0x367cd507, 0x9159015a, 0x3070dd17, 0x152fecd8, 0xf70e5939, 0x67332667, 0xffc00b31, 0x8eb44a87, 0x68581511, 0xdb0c2e0d, 0x64f98fa7, 0x47b5481d, 0xbefa4fa4];

const W = new Uint32Array(160);

const sha512Core = (msg: Uint8Array, iv: number[], outWords: number): Uint8Array => {
  const len = msg.length;
  const blocks = Math.ceil((len + 17) / 128);
  const buf = new Uint8Array(blocks * 128);
  buf.set(msg);
  buf[len] = 0x80;
  const bits = len * 8;
  const dv = new DataView(buf.buffer);
  dv.setUint32(buf.length - 8, Math.floor(bits / 0x100000000));
  dv.setUint32(buf.length - 4, bits >>> 0);
  const h = Uint32Array.from(iv);
  for (let b = 0; b < blocks; b++) {
    const off = b * 128;
    for (let i = 0; i < 32; i++) W[i] = dv.getUint32(off + i * 4);
    for (let i = 16; i < 80; i++) {
      const xh = W[(i - 15) * 2]!, xl = W[(i - 15) * 2 + 1]!;
      const s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
      const s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
      const yh = W[(i - 2) * 2]!, yl = W[(i - 2) * 2 + 1]!;
      const s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
      const s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
      let lo = (s0l >>> 0) + (s1l >>> 0) + W[(i - 7) * 2 + 1]! + W[(i - 16) * 2 + 1]!;
      const hi = s0h + s1h + W[(i - 7) * 2]! + W[(i - 16) * 2]! + Math.floor(lo / 0x100000000);
      lo >>>= 0;
      W[i * 2] = hi >>> 0; W[i * 2 + 1] = lo;
    }
    let ah = h[0]!, al = h[1]!, bh = h[2]!, bl = h[3]!, ch = h[4]!, cl = h[5]!, dh = h[6]!, dl = h[7]!;
    let eh = h[8]!, el = h[9]!, fh = h[10]!, fl = h[11]!, gh = h[12]!, gl = h[13]!, hh = h[14]!, hl = h[15]!;
    for (let i = 0; i < 80; i++) {
      const S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
      const S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
      const chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
      let t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + K512[i * 2 + 1]! + W[i * 2 + 1]!;
      const t1h = (hh + S1h + chh + K512[i * 2]! + W[i * 2]! + Math.floor(t1l / 0x100000000)) >>> 0;
      t1l >>>= 0;
      const S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
      const S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
      const mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
      let t2l = (S0l >>> 0) + (mjl >>> 0);
      const t2h = (S0h + mjh + Math.floor(t2l / 0x100000000)) >>> 0;
      t2l >>>= 0;
      hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
      let nl = (dl >>> 0) + t1l;
      eh = (dh + t1h + Math.floor(nl / 0x100000000)) >>> 0; el = nl >>> 0;
      dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
      nl = t1l + t2l;
      ah = (t1h + t2h + Math.floor(nl / 0x100000000)) >>> 0; al = nl >>> 0;
    }
    const add = (i: number, xh: number, xl: number) => {
      const lo = h[i + 1]! + (xl >>> 0);
      h[i] = (h[i]! + xh + Math.floor(lo / 0x100000000)) >>> 0;
      h[i + 1] = lo >>> 0;
    };
    add(0, ah, al); add(2, bh, bl); add(4, ch, cl); add(6, dh, dl); add(8, eh, el); add(10, fh, fl); add(12, gh, gl); add(14, hh, hl);
  }
  const out = new Uint8Array(outWords * 4);
  const odv = new DataView(out.buffer);
  for (let i = 0; i < outWords; i++) odv.setUint32(i * 4, h[i]!);
  return out;
};

export const sha512 = (msg: Uint8Array): Uint8Array => sha512Core(msg, IV512, 16);
export const sha384 = (msg: Uint8Array): Uint8Array => sha512Core(msg, IV384, 12);

// ---------- SHA-256 ----------

const K256 = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

export const sha256 = (msg: Uint8Array): Uint8Array => {
  const len = msg.length;
  const blocks = Math.ceil((len + 9) / 64);
  const buf = new Uint8Array(blocks * 64);
  buf.set(msg);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(buf.length - 8, Math.floor((len * 8) / 0x100000000));
  dv.setUint32(buf.length - 4, (len * 8) >>> 0);
  const h = new Uint32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const w = new Uint32Array(64);
  for (let b = 0; b < blocks; b++) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(b * 64 + i * 4);
    for (let i = 16; i < 64; i++) {
      const x = w[i - 15]!, y = w[i - 2]!;
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
      w[i] = (w[i - 16]! + s0 + w[i - 7]! + s1) >>> 0;
    }
    let a = h[0]!, bb = h[1]!, c = h[2]!, d = h[3]!, e = h[4]!, f = h[5]!, g = h[6]!, hh = h[7]!;
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const t1 = (hh + S1 + ((e & f) ^ (~e & g)) + K256[i]! + w[i]!) >>> 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const t2 = (S0 + ((a & bb) ^ (a & c) ^ (bb & c))) >>> 0;
      hh = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = bb; bb = a; a = (t1 + t2) >>> 0;
    }
    h[0] = (h[0]! + a) >>> 0; h[1] = (h[1]! + bb) >>> 0; h[2] = (h[2]! + c) >>> 0; h[3] = (h[3]! + d) >>> 0;
    h[4] = (h[4]! + e) >>> 0; h[5] = (h[5]! + f) >>> 0; h[6] = (h[6]! + g) >>> 0; h[7] = (h[7]! + hh) >>> 0;
  }
  const out = new Uint8Array(32);
  const odv = new DataView(out.buffer);
  for (let i = 0; i < 8; i++) odv.setUint32(i * 4, h[i]!);
  return out;
};

// ---------- SHA-1 ----------

export const sha1 = (msg: Uint8Array): Uint8Array => {
  const len = msg.length;
  const blocks = Math.ceil((len + 9) / 64);
  const buf = new Uint8Array(blocks * 64);
  buf.set(msg);
  buf[len] = 0x80;
  const dv = new DataView(buf.buffer);
  dv.setUint32(buf.length - 8, Math.floor((len * 8) / 0x100000000));
  dv.setUint32(buf.length - 4, (len * 8) >>> 0);
  let h0 = 0x67452301, h1 = 0xefcdab89, h2 = 0x98badcfe, h3 = 0x10325476, h4 = 0xc3d2e1f0;
  const w = new Uint32Array(80);
  for (let b = 0; b < blocks; b++) {
    for (let i = 0; i < 16; i++) w[i] = dv.getUint32(b * 64 + i * 4);
    for (let i = 16; i < 80; i++) { const x = w[i - 3]! ^ w[i - 8]! ^ w[i - 14]! ^ w[i - 16]!; w[i] = (x << 1) | (x >>> 31); }
    let a = h0, bb = h1, c = h2, d = h3, e = h4;
    for (let i = 0; i < 80; i++) {
      const f = i < 20 ? (bb & c) | (~bb & d) : i < 40 ? bb ^ c ^ d : i < 60 ? (bb & c) | (bb & d) | (c & d) : bb ^ c ^ d;
      const k = i < 20 ? 0x5a827999 : i < 40 ? 0x6ed9eba1 : i < 60 ? 0x8f1bbcdc : 0xca62c1d6;
      const t = (((a << 5) | (a >>> 27)) + f + e + k + w[i]!) >>> 0;
      e = d; d = c; c = ((bb << 30) | (bb >>> 2)) >>> 0; bb = a; a = t;
    }
    h0 = (h0 + a) >>> 0; h1 = (h1 + bb) >>> 0; h2 = (h2 + c) >>> 0; h3 = (h3 + d) >>> 0; h4 = (h4 + e) >>> 0;
  }
  const out = new Uint8Array(20);
  const odv = new DataView(out.buffer);
  [h0, h1, h2, h3, h4].forEach((x, i) => odv.setUint32(i * 4, x));
  return out;
};

export const hashFn = (name: string): ((m: Uint8Array) => Uint8Array) | null => {
  switch (name.toUpperCase().replace(/-/g, '')) {
    case 'SHA1': return sha1;
    case 'SHA256': return sha256;
    case 'SHA384': return sha384;
    case 'SHA512': return sha512;
    default: return null;
  }
};
export const hashBlock = (name: string): number => (/384|512/.test(name) ? 128 : 64);

export const hmac = (hashName: string, key: Uint8Array, data: Uint8Array): Uint8Array => {
  const H = hashFn(hashName);
  if (!H) throw new Error(`Unsupported hash ${hashName}`);
  const bs = hashBlock(hashName);
  let k = key.length > bs ? H(key) : key;
  const pad = new Uint8Array(bs);
  pad.set(k);
  k = pad;
  const ipad = k.map((b) => b ^ 0x36), opad = k.map((b) => b ^ 0x5c);
  return H(concat(opad, H(concat(ipad, data))));
};

// ---------- AES (FIPS 197) ----------

const SBOX = new Uint8Array(256), INV = new Uint8Array(256);
const TE: Uint32Array[] = [0, 1, 2, 3].map(() => new Uint32Array(256));
const TD: Uint32Array[] = [0, 1, 2, 3].map(() => new Uint32Array(256));
const gmul = (a: number, b: number): number => {
  let p = 0;
  for (let i = 0; i < 8; i++) {
    if (b & 1) p ^= a;
    const hi = a & 0x80;
    a = (a << 1) & 0xff;
    if (hi) a ^= 0x1b;
    b >>= 1;
  }
  return p;
};
(() => {
  // Multiplicative inverses in GF(2^8), then the affine transform.
  const inv = new Uint8Array(256);
  for (let x = 1; x < 256; x++) for (let y = 1; y < 256; y++) if (gmul(x, y) === 1) { inv[x] = y; break; }
  for (let x = 0; x < 256; x++) {
    const b = inv[x]!;
    let s = b;
    for (let i = 1; i <= 4; i++) s ^= ((b << i) | (b >>> (8 - i))) & 0xff;
    s ^= 0x63;
    SBOX[x] = s; INV[s] = x;
  }
  const rotr = (w: number, n: number) => ((w >>> n) | (w << (32 - n))) >>> 0;
  for (let x = 0; x < 256; x++) {
    const s = SBOX[x]!;
    const e = ((gmul(s, 2) << 24) | (s << 16) | (s << 8) | gmul(s, 3)) >>> 0;
    const si = INV[x]!;
    const d = ((gmul(si, 14) << 24) | (gmul(si, 9) << 16) | (gmul(si, 13) << 8) | gmul(si, 11)) >>> 0;
    for (let k = 0; k < 4; k++) { TE[k]![x] = rotr(e, 8 * k); TD[k]![x] = rotr(d, 8 * k); }
  }
})();

/** Expanded AES key with encrypt and decrypt round keys. */
export class Aes {
  private ek: Uint32Array;
  private dk: Uint32Array;
  private rounds: number;
  constructor(key: Uint8Array) {
    if (![16, 24, 32].includes(key.length)) throw new Error('AES keys are 16, 24 or 32 bytes.');
    const nk = key.length / 4;
    const nr = nk + 6;
    this.rounds = nr;
    const total = (nr + 1) * 4;
    const ek = new Uint32Array(total);
    for (let i = 0; i < nk; i++) ek[i] = ((key[i * 4]! << 24) | (key[i * 4 + 1]! << 16) | (key[i * 4 + 2]! << 8) | key[i * 4 + 3]!) >>> 0;
    const sub = (t: number) => ((SBOX[t >>> 24]! << 24) | (SBOX[(t >>> 16) & 255]! << 16) | (SBOX[(t >>> 8) & 255]! << 8) | SBOX[t & 255]!) >>> 0;
    let rcon = 1;
    for (let i = nk; i < total; i++) {
      let t = ek[i - 1]!;
      if (i % nk === 0) { t = sub(((t << 8) | (t >>> 24)) >>> 0) ^ (rcon << 24); rcon = gmul(rcon, 2); }
      else if (nk > 6 && i % nk === 4) t = sub(t);
      ek[i] = (ek[i - nk]! ^ t) >>> 0;
    }
    const dk = new Uint32Array(total);
    for (let r = 0; r <= nr; r++) for (let c = 0; c < 4; c++) {
      const w = ek[(nr - r) * 4 + c]!;
      dk[r * 4 + c] = r === 0 || r === nr ? w
        : (TD[0]![SBOX[w >>> 24]!]! ^ TD[1]![SBOX[(w >>> 16) & 255]!]! ^ TD[2]![SBOX[(w >>> 8) & 255]!]! ^ TD[3]![SBOX[w & 255]!]!) >>> 0;
    }
    this.ek = ek; this.dk = dk;
  }

  private s = new Uint32Array(4);
  private t = new Uint32Array(4);

  private block(src: Uint8Array, off: number, out: Uint8Array, outOff: number, decrypt: boolean): void {
    const k = decrypt ? this.dk : this.ek;
    const T = decrypt ? TD : TE;
    const box = decrypt ? INV : SBOX;
    // Column shift: +1 per row when encrypting (ShiftRows), -1 when decrypting.
    const sh = decrypt ? 3 : 1;
    let s = this.s, t = this.t;
    for (let c = 0; c < 4; c++) s[c] = (((src[off + c * 4]! << 24) | (src[off + c * 4 + 1]! << 16) | (src[off + c * 4 + 2]! << 8) | src[off + c * 4 + 3]!) ^ k[c]!) >>> 0;
    let ki = 4;
    for (let r = 1; r < this.rounds; r++) {
      for (let c = 0; c < 4; c++) {
        t[c] = (T[0]![s[c]! >>> 24]! ^ T[1]![(s[(c + sh) & 3]! >>> 16) & 255]! ^ T[2]![(s[(c + 2 * sh) & 3]! >>> 8) & 255]! ^ T[3]![s[(c + 3 * sh) & 3]! & 255]! ^ k[ki + c]!) >>> 0;
      }
      ki += 4;
      const x = s; s = t; t = x;
    }
    for (let c = 0; c < 4; c++) {
      const w = ((box[s[c]! >>> 24]! << 24) | (box[(s[(c + sh) & 3]! >>> 16) & 255]! << 16) | (box[(s[(c + 2 * sh) & 3]! >>> 8) & 255]! << 8) | box[s[(c + 3 * sh) & 3]! & 255]!) ^ k[ki + c]!;
      out[outOff + c * 4] = w >>> 24; out[outOff + c * 4 + 1] = (w >>> 16) & 255; out[outOff + c * 4 + 2] = (w >>> 8) & 255; out[outOff + c * 4 + 3] = w & 255;
    }
  }

  /** CBC without padding; the input length must be a multiple of 16. */
  encryptCbc(data: Uint8Array, iv: Uint8Array): Uint8Array {
    if (data.length % 16) throw new Error('Data must be a multiple of 16 bytes.');
    const out = new Uint8Array(data.length);
    const x = new Uint8Array(16);
    let prev = iv.subarray(0, 16);
    for (let off = 0; off < data.length; off += 16) {
      for (let j = 0; j < 16; j++) x[j] = data[off + j]! ^ prev[j]!;
      this.block(x, 0, out, off, false);
      prev = out.subarray(off, off + 16);
    }
    return out;
  }

  decryptCbc(data: Uint8Array, iv: Uint8Array): Uint8Array {
    if (data.length % 16) throw new Error('Data must be a multiple of 16 bytes.');
    const out = new Uint8Array(data.length);
    let prev = iv.subarray(0, 16);
    for (let off = 0; off < data.length; off += 16) {
      this.block(data, off, out, off, true);
      for (let j = 0; j < 16; j++) out[off + j]! ^= prev[j]!;
      prev = data.subarray(off, off + 16);
    }
    return out;
  }

  encryptEcb(data: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.length - (data.length % 16));
    for (let off = 0; off < out.length; off += 16) this.block(data, off, out, off, false);
    return out;
  }

  decryptEcb(data: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.length - (data.length % 16));
    for (let off = 0; off < out.length; off += 16) this.block(data, off, out, off, true);
    return out;
  }
}

export const randomBytes = (n: number): Uint8Array => {
  const out = new Uint8Array(n);
  globalThis.crypto.getRandomValues(out);
  return out;
};

export const toBase64 = (b: Uint8Array): string => {
  let s = '';
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(s);
};
export const fromBase64 = (s: string): Uint8Array => {
  const bin = atob(s.replace(/\s+/g, ''));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
};

export const utf16le = (s: string): Uint8Array => {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out[i * 2] = c & 255; out[i * 2 + 1] = c >> 8; }
  return out;
};

export const u32le = (n: number): Uint8Array => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n >>> 0, true); return b; };
