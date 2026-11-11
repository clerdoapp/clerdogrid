// Blowfish (Schneier, 1993), used by early OpenDocument encryption (Blowfish CFB).
// The initial P-array and S-boxes are the hexadecimal digits of pi; they are computed here
// once, when first needed, rather than written out.

let digits: Uint32Array | null = null;

/** The first 1042 32-bit words of pi's fractional hexadecimal expansion. */
const piWords = (): Uint32Array => {
  if (digits) return digits;
  const hexDigits = 1042 * 8;
  const scale = 16n ** BigInt(hexDigits + 16);
  const atanInv = (x: bigint): bigint => {
    let sum = 0n, term = scale / x, n = 1n, sign = 1n;
    const x2 = x * x;
    while (term) { sum += sign * (term / n); term /= x2; n += 2n; sign = -sign; }
    return sum;
  };
  const pi = 16n * atanInv(5n) - 4n * atanInv(239n);
  const hex = pi.toString(16).slice(1, 1 + hexDigits);
  const out = new Uint32Array(1042);
  for (let i = 0; i < 1042; i++) out[i] = parseInt(hex.slice(i * 8, i * 8 + 8), 16) >>> 0;
  digits = out;
  return out;
};

export class Blowfish {
  private p: Uint32Array;
  private s: Uint32Array[];
  constructor(key: Uint8Array) {
    const pi = piWords();
    this.p = pi.slice(0, 18);
    this.s = [0, 1, 2, 3].map((i) => pi.slice(18 + i * 256, 18 + (i + 1) * 256));
    let k = 0;
    for (let i = 0; i < 18; i++) {
      let w = 0;
      for (let j = 0; j < 4; j++) { w = ((w << 8) | key[k % key.length]!) >>> 0; k++; }
      this.p[i]! ^= w;
    }
    const lr = [0, 0];
    for (let i = 0; i < 18; i += 2) { this.block(lr); this.p[i] = lr[0]!; this.p[i + 1] = lr[1]!; }
    for (const box of this.s) for (let i = 0; i < 256; i += 2) { this.block(lr); box[i] = lr[0]!; box[i + 1] = lr[1]!; }
  }
  private f(x: number): number {
    const [a, b, c, d] = this.s as [Uint32Array, Uint32Array, Uint32Array, Uint32Array];
    return ((((a[x >>> 24]! + b[(x >>> 16) & 255]!) >>> 0) ^ c[(x >>> 8) & 255]!) + d[x & 255]!) >>> 0;
  }
  /** Encrypts [left, right] in place. */
  block(lr: number[]): void {
    let l = lr[0]! >>> 0, r = lr[1]! >>> 0;
    for (let i = 0; i < 16; i += 2) {
      l = (l ^ this.p[i]!) >>> 0; r = (r ^ this.f(l)) >>> 0;
      r = (r ^ this.p[i + 1]!) >>> 0; l = (l ^ this.f(r)) >>> 0;
    }
    lr[0] = (r ^ this.p[17]!) >>> 0;
    lr[1] = (l ^ this.p[16]!) >>> 0;
  }
  /** CFB mode with a 64-bit feedback, decrypting. */
  decryptCfb(data: Uint8Array, iv: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.length);
    const reg = iv.slice(0, 8);
    const lr = [0, 0];
    for (let off = 0; off < data.length; off += 8) {
      lr[0] = ((reg[0]! << 24) | (reg[1]! << 16) | (reg[2]! << 8) | reg[3]!) >>> 0;
      lr[1] = ((reg[4]! << 24) | (reg[5]! << 16) | (reg[6]! << 8) | reg[7]!) >>> 0;
      this.block(lr);
      const ks = [lr[0]! >>> 24, (lr[0]! >>> 16) & 255, (lr[0]! >>> 8) & 255, lr[0]! & 255, lr[1]! >>> 24, (lr[1]! >>> 16) & 255, (lr[1]! >>> 8) & 255, lr[1]! & 255];
      const n = Math.min(8, data.length - off);
      for (let i = 0; i < n; i++) out[off + i] = data[off + i]! ^ ks[i]!;
      for (let i = 0; i < n; i++) reg[i] = data[off + i]!;
    }
    return out;
  }
}
