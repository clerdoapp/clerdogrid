import { T } from '../../ui/i18n';
// QR code encoder written from ISO/IEC 18004: byte mode, error correction level M, versions 1 to 40.

/** Error correction codewords per block, level M, versions 1-40 (ISO/IEC 18004 table 9). */
const EC_PER_BLOCK = [10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28];
/** Number of error correction blocks, level M, versions 1-40. */
const BLOCKS = [1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49];

/** Modules that carry data or error correction (everything but function patterns and format/version info). */
const rawModules = (v: number): number => {
  let n = (16 * v + 128) * v + 64;
  if (v >= 2) {
    const a = Math.floor(v / 7) + 2;
    n -= (25 * a - 10) * a - 55;
    if (v >= 7) n -= 36;
  }
  return n;
};
const totalCodewords = (v: number): number => Math.floor(rawModules(v) / 8);
const dataCodewords = (v: number): number => totalCodewords(v) - EC_PER_BLOCK[v - 1]! * BLOCKS[v - 1]!;

/** Byte-mode capacity in bytes for a version at level M. */
export const capacity = (v: number): number => Math.floor((dataCodewords(v) * 8 - 4 - (v < 10 ? 8 : 16)) / 8);

// ---------- Reed-Solomon over GF(256), polynomial x^8 + x^4 + x^3 + x^2 + 1 ----------

const EXP = new Uint8Array(512), LOG = new Uint8Array(256);
(() => { let x = 1; for (let i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 0x100) x ^= 0x11d; } for (let i = 255; i < 512; i++) EXP[i] = EXP[i - 255]!; })();
const mul = (a: number, b: number): number => (a && b ? EXP[LOG[a]! + LOG[b]!]! : 0);

const generator = (degree: number): Uint8Array => {
  // Coefficients of (x - a^0)(x - a^1)...(x - a^(degree-1)), highest power first, leading 1 dropped.
  let g = [1];
  for (let i = 0; i < degree; i++) {
    const next = new Array<number>(g.length + 1).fill(0);
    for (let j = 0; j < g.length; j++) { next[j] ^= g[j]!; next[j + 1] ^= mul(g[j]!, EXP[i]!); }
    g = next;
  }
  return Uint8Array.from(g.slice(1));
};

const remainder = (data: Uint8Array, gen: Uint8Array): Uint8Array => {
  const r = new Uint8Array(gen.length);
  for (const b of data) {
    const f = b ^ r[0]!;
    r.copyWithin(0, 1); r[r.length - 1] = 0;
    for (let i = 0; i < gen.length; i++) r[i] ^= mul(gen[i]!, f);
  }
  return r;
};

// ---------- Codewords ----------

const codewords = (bytes: Uint8Array, v: number): Uint8Array => {
  const bits: number[] = [];
  const put = (val: number, len: number) => { for (let i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1); };
  put(0b0100, 4);
  put(bytes.length, v < 10 ? 8 : 16);
  for (const b of bytes) put(b, 8);
  const cap = dataCodewords(v) * 8;
  put(0, Math.min(4, cap - bits.length));
  while (bits.length % 8) bits.push(0);
  const data = new Uint8Array(dataCodewords(v));
  for (let i = 0; i < bits.length / 8; i++) { let x = 0; for (let j = 0; j < 8; j++) x = (x << 1) | bits[i * 8 + j]!; data[i] = x; }
  for (let i = bits.length / 8, pad = 0xec; i < data.length; i++, pad ^= 0xec ^ 0x11) data[i] = pad;

  const nb = BLOCKS[v - 1]!, ec = EC_PER_BLOCK[v - 1]!, total = totalCodewords(v);
  const shortLen = Math.floor(total / nb), longCount = total % nb;
  const gen = generator(ec);
  const blocks: Uint8Array[] = [], ecs: Uint8Array[] = [];
  for (let i = 0, k = 0; i < nb; i++) {
    const len = shortLen - ec + (i >= nb - longCount ? 1 : 0);
    const d = data.subarray(k, k + len);
    k += len;
    blocks.push(d);
    ecs.push(remainder(d, gen));
  }
  const out: number[] = [];
  const maxLen = shortLen - ec + (longCount ? 1 : 0);
  for (let i = 0; i < maxLen; i++) for (const b of blocks) if (i < b.length) out.push(b[i]!);
  for (let i = 0; i < ec; i++) for (const e of ecs) out.push(e[i]!);
  return Uint8Array.from(out);
};

// ---------- Matrix ----------

const alignmentPositions = (v: number): number[] => {
  if (v === 1) return [];
  const n = Math.floor(v / 7) + 2, size = v * 4 + 17;
  const step = v === 32 ? 26 : Math.ceil((v * 4 + 4) / (n * 2 - 2)) * 2;
  const out = [6];
  for (let p = size - 7; out.length < n; p -= step) out.splice(1, 0, p);
  return out;
};

class Matrix {
  size: number;
  dark: Uint8Array;
  fixed: Uint8Array;
  constructor(public v: number) {
    this.size = v * 4 + 17;
    this.dark = new Uint8Array(this.size * this.size);
    this.fixed = new Uint8Array(this.size * this.size);
  }
  set(x: number, y: number, on: boolean, fixed = true) { const i = y * this.size + x; this.dark[i] = on ? 1 : 0; if (fixed) this.fixed[i] = 1; }
  get(x: number, y: number) { return this.dark[y * this.size + x] === 1; }
  isFixed(x: number, y: number) { return this.fixed[y * this.size + x] === 1; }
}

const drawFunctionPatterns = (m: Matrix): void => {
  const s = m.size;
  for (let i = 0; i < s; i++) { m.set(6, i, i % 2 === 0); m.set(i, 6, i % 2 === 0); }
  const finder = (cx: number, cy: number) => {
    for (let dy = -4; dy <= 4; dy++) for (let dx = -4; dx <= 4; dx++) {
      const x = cx + dx, y = cy + dy;
      if (x < 0 || y < 0 || x >= s || y >= s) continue;
      const d = Math.max(Math.abs(dx), Math.abs(dy));
      m.set(x, y, d !== 2 && d !== 4);
    }
  };
  finder(3, 3); finder(s - 4, 3); finder(3, s - 4);
  const pos = alignmentPositions(m.v);
  for (let i = 0; i < pos.length; i++) for (let j = 0; j < pos.length; j++) {
    if ((i === 0 && j === 0) || (i === 0 && j === pos.length - 1) || (i === pos.length - 1 && j === 0)) continue;
    for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) m.set(pos[i]! + dx, pos[j]! + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
  }
  drawFormat(m, 0); // reserve
  drawVersion(m);
};

/** Level M is 00; 15 bits = 5 data bits + 10 BCH bits, masked with 101010000010010. */
export const formatBits = (mask: number): number => {
  const data = (0b00 << 3) | mask;
  let rem = data;
  for (let i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
  return ((data << 10) | rem) ^ 0x5412;
};

const drawFormat = (m: Matrix, mask: number): void => {
  const bits = formatBits(mask), s = m.size;
  const bit = (i: number) => ((bits >>> i) & 1) === 1;
  for (let i = 0; i <= 5; i++) m.set(8, i, bit(i));
  m.set(8, 7, bit(6)); m.set(8, 8, bit(7)); m.set(7, 8, bit(8));
  for (let i = 9; i < 15; i++) m.set(14 - i, 8, bit(i));
  for (let i = 0; i < 8; i++) m.set(s - 1 - i, 8, bit(i));
  for (let i = 8; i < 15; i++) m.set(8, s - 15 + i, bit(i));
  m.set(8, s - 8, true); // the dark module
};

const drawVersion = (m: Matrix): void => {
  if (m.v < 7) return;
  let rem = m.v;
  for (let i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1f25);
  const bits = (m.v << 12) | rem;
  for (let i = 0; i < 18; i++) {
    const on = ((bits >>> i) & 1) === 1, a = m.size - 11 + (i % 3), b = Math.floor(i / 3);
    m.set(a, b, on); m.set(b, a, on);
  }
};

const placeData = (m: Matrix, data: Uint8Array): void => {
  let i = 0;
  const s = m.size;
  for (let right = s - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (let vert = 0; vert < s; vert++) for (let j = 0; j < 2; j++) {
      const x = right - j, upward = ((right + 1) & 2) === 0, y = upward ? s - 1 - vert : vert;
      if (m.isFixed(x, y)) continue;
      if (i < data.length * 8) m.set(x, y, ((data[i >>> 3]! >>> (7 - (i & 7))) & 1) === 1, false);
      i++;
    }
  }
};

const MASKS: ((x: number, y: number) => boolean)[] = [
  (x, y) => (x + y) % 2 === 0,
  (_x, y) => y % 2 === 0,
  (x) => x % 3 === 0,
  (x, y) => (x + y) % 3 === 0,
  (x, y) => (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0,
  (x, y) => ((x * y) % 2) + ((x * y) % 3) === 0,
  (x, y) => (((x * y) % 2) + ((x * y) % 3)) % 2 === 0,
  (x, y) => (((x + y) % 2) + ((x * y) % 3)) % 2 === 0,
];

const applyMask = (m: Matrix, k: number): void => {
  const f = MASKS[k]!;
  for (let y = 0; y < m.size; y++) for (let x = 0; x < m.size; x++) if (!m.isFixed(x, y) && f(x, y)) m.dark[y * m.size + x]! ^= 1;
};

/** Penalty score from the standard's four rules; the mask with the lowest score is used. */
const penalty = (m: Matrix): number => {
  const s = m.size;
  let score = 0, darkCount = 0;
  const line = (get: (i: number) => boolean) => {
    let runColor = get(0), run = 1;
    for (let i = 1; i < s; i++) {
      const c = get(i);
      if (c === runColor) { run++; if (run === 5) score += 3; else if (run > 5) score++; }
      else { runColor = c; run = 1; }
    }
    // Finder-like 1:1:3:1:1 with four light modules on one side.
    const p = (i: number) => (i < 0 || i >= s ? false : get(i));
    for (let i = -4; i < s; i++) {
      if (p(i) && !p(i + 1) && p(i + 2) && p(i + 3) && p(i + 4) && !p(i + 5) && p(i + 6)) {
        if ((!p(i - 1) && !p(i - 2) && !p(i - 3) && !p(i - 4)) || (!p(i + 7) && !p(i + 8) && !p(i + 9) && !p(i + 10))) score += 40;
      }
    }
  };
  for (let y = 0; y < s; y++) line((i) => m.get(i, y));
  for (let x = 0; x < s; x++) line((i) => m.get(x, i));
  for (let y = 0; y < s - 1; y++) for (let x = 0; x < s - 1; x++) {
    const c = m.get(x, y);
    if (c === m.get(x + 1, y) && c === m.get(x, y + 1) && c === m.get(x + 1, y + 1)) score += 3;
  }
  for (let i = 0; i < m.dark.length; i++) darkCount += m.dark[i]!;
  const total = s * s;
  score += (Math.ceil(Math.abs(darkCount * 20 - total * 10) / total) - 1) * 10;
  return score;
};

export interface QrCode { size: number; version: number; mask: number; dark: (x: number, y: number) => boolean }

/** Encodes bytes (or UTF-8 text). `mask` forces a mask pattern (for tests). Throws when the data won't fit in version 40. */
export const encodeQr = (input: string | Uint8Array, mask?: number): QrCode => {
  const bytes = typeof input === 'string' ? new TextEncoder().encode(input) : input;
  let v = 1;
  while (v <= 40 && capacity(v) < bytes.length) v++;
  if (v > 40) throw new Error(T('Too much data for a QR code.'));
  const data = codewords(bytes, v);
  const m = new Matrix(v);
  drawFunctionPatterns(m);
  placeData(m, data);
  let best = mask ?? 0;
  if (mask === undefined) {
    let bestScore = Infinity;
    for (let k = 0; k < 8; k++) {
      applyMask(m, k); drawFormat(m, k);
      const sc = penalty(m);
      if (sc < bestScore) { bestScore = sc; best = k; }
      applyMask(m, k);
    }
  }
  applyMask(m, best);
  drawFormat(m, best);
  return { size: m.size, version: v, mask: best, dark: (x, y) => m.get(x, y) };
};

/** An SVG path (one unit per module, with a 4-module quiet zone) for drawing the code. */
export const qrPath = (q: QrCode): string => {
  let d = '';
  for (let y = 0; y < q.size; y++) {
    let x = 0;
    while (x < q.size) {
      if (!q.dark(x, y)) { x++; continue; }
      let w = 1;
      while (x + w < q.size && q.dark(x + w, y)) w++;
      d += `M${x + 4} ${y + 4}h${w}v1h${-w}z`;
      x += w;
    }
  }
  return d;
};
