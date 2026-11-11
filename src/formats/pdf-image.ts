// PNG decoding (RFC 2083 / ISO 15948) and JPEG header reading (ITU T.81) for PDF export.
import { unzlibSync } from 'fflate';

export interface RasterImage { width: number; height: number; channels: 1 | 3; data: Uint8Array; alpha?: Uint8Array }
export interface JpegInfo { width: number; height: number; components: number; adobeInverted: boolean }

const MAX_PIXELS = 100_000_000;

export const isPng = (b: Uint8Array): boolean => b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
export const isJpeg = (b: Uint8Array): boolean => b.length > 3 && b[0] === 0xff && b[1] === 0xd8;

export const jpegInfo = (b: Uint8Array): JpegInfo | null => {
  if (!isJpeg(b)) return null;
  let p = 2;
  let adobe = false;
  while (p + 4 <= b.length) {
    if (b[p] !== 0xff) { p++; continue; }
    const m = b[p + 1]!;
    if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { p += 2; continue; }
    const len = (b[p + 2]! << 8) | b[p + 3]!;
    if (m === 0xee && len >= 12 && String.fromCharCode(...b.subarray(p + 4, p + 9)) === 'Adobe') adobe = true;
    if ((m >= 0xc0 && m <= 0xcf) && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
      const height = (b[p + 5]! << 8) | b[p + 6]!, width = (b[p + 7]! << 8) | b[p + 8]!, components = b[p + 9]!;
      if (!width || !height || ![1, 3, 4].includes(components)) return null;
      return { width, height, components, adobeInverted: adobe && components === 4 };
    }
    if (m === 0xd9 || m === 0xda) break;
    p += 2 + len;
  }
  return null;
};

const paeth = (a: number, b: number, c: number): number => {
  const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
};

export const decodePng = (b: Uint8Array): RasterImage => {
  if (!isPng(b)) throw new Error('Not a PNG image.');
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let p = 8;
  let w = 0, h = 0, depth = 8, ctype = 0, interlace = 0;
  let palette: Uint8Array | undefined, trns: Uint8Array | undefined;
  const idat: Uint8Array[] = [];
  while (p + 8 <= b.length) {
    const len = v.getUint32(p);
    const type = String.fromCharCode(b[p + 4]!, b[p + 5]!, b[p + 6]!, b[p + 7]!);
    const data = b.subarray(p + 8, Math.min(b.length, p + 8 + len));
    if (type === 'IHDR') { w = v.getUint32(p + 8); h = v.getUint32(p + 12); depth = data[8]!; ctype = data[9]!; interlace = data[12]!; }
    else if (type === 'PLTE') palette = data;
    else if (type === 'tRNS') trns = data;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (!w || !h || w * h > MAX_PIXELS) throw new Error('Unsupported PNG image.');
  const total = idat.reduce((a, d) => a + d.length, 0);
  const z = new Uint8Array(total);
  let q = 0;
  for (const d of idat) { z.set(d, q); q += d.length; }
  const raw = unzlibSync(z);
  const chans = ctype === 0 ? 1 : ctype === 2 ? 3 : ctype === 3 ? 1 : ctype === 4 ? 2 : 4;
  const bitsPP = chans * depth;
  const bpp = Math.max(1, bitsPP >> 3);
  const outCh: 1 | 3 = ctype === 0 || ctype === 4 ? 1 : 3;
  const hasAlpha = ctype === 4 || ctype === 6 || !!trns;
  const data = new Uint8Array(w * h * outCh);
  const alpha = hasAlpha ? new Uint8Array(w * h).fill(255) : undefined;
  const max = (1 << depth) - 1;

  const sample = (row: Uint8Array, i: number): number => {
    if (depth === 8) return row[i]!;
    if (depth === 16) return row[i * 2]!;
    const bit = i * depth;
    return (row[bit >> 3]! >> (8 - depth - (bit & 7))) & max;
  };
  const sample16 = (row: Uint8Array, i: number): number => (depth === 16 ? (row[i * 2]! << 8) | row[i * 2 + 1]! : sample(row, i));
  const put = (row: Uint8Array, x: number, ox: number, oy: number) => {
    const o = oy * w + ox;
    const s = x * chans;
    const to8 = (n: number) => (depth === 16 ? n : depth === 8 ? n : Math.round((n * 255) / max));
    if (ctype === 3) {
      const idx = sample(row, x);
      data[o * 3] = palette?.[idx * 3] ?? 0; data[o * 3 + 1] = palette?.[idx * 3 + 1] ?? 0; data[o * 3 + 2] = palette?.[idx * 3 + 2] ?? 0;
      if (alpha && trns) alpha[o] = idx < trns.length ? trns[idx]! : 255;
      return;
    }
    if (ctype === 0) {
      const g = sample(row, s);
      data[o] = to8(g);
      if (alpha && trns && trns.length >= 2 && sample16(row, s) === ((trns[0]! << 8) | trns[1]!) >> (depth === 16 ? 0 : 0)) alpha[o] = 0;
      return;
    }
    if (ctype === 4) { data[o] = to8(sample(row, s)); alpha![o] = to8(sample(row, s + 1)); return; }
    data[o * 3] = to8(sample(row, s)); data[o * 3 + 1] = to8(sample(row, s + 1)); data[o * 3 + 2] = to8(sample(row, s + 2));
    if (ctype === 6) alpha![o] = to8(sample(row, s + 3));
    else if (alpha && trns && trns.length >= 6) {
      const r = sample16(row, s), g = sample16(row, s + 1), bl = sample16(row, s + 2);
      if (r === ((trns[0]! << 8) | trns[1]!) && g === ((trns[2]! << 8) | trns[3]!) && bl === ((trns[4]! << 8) | trns[5]!)) alpha[o] = 0;
    }
  };

  let pos = 0;
  const pass = (pw: number, ph: number, map: (x: number, y: number) => [number, number]) => {
    if (!pw || !ph) return;
    const stride = Math.ceil((pw * bitsPP) / 8);
    let prev = new Uint8Array(stride);
    for (let y = 0; y < ph; y++) {
      const f = raw[pos++] ?? 0;
      const cur = raw.slice(pos, pos + stride);
      pos += stride;
      for (let i = 0; i < stride; i++) {
        const a = i >= bpp ? cur[i - bpp]! : 0, up = prev[i]!, c = i >= bpp ? prev[i - bpp]! : 0;
        cur[i] = (cur[i]! + (f === 1 ? a : f === 2 ? up : f === 3 ? (a + up) >> 1 : f === 4 ? paeth(a, up, c) : 0)) & 255;
      }
      for (let x = 0; x < pw; x++) { const [ox, oy] = map(x, y); put(cur, x, ox, oy); }
      prev = cur;
    }
  };
  if (!interlace) pass(w, h, (x, y) => [x, y]);
  else {
    const P = [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]] as const;
    for (const [x0, y0, dx, dy] of P) pass(Math.ceil((w - x0) / dx), Math.ceil((h - y0) / dy), (x, y) => [x0 + x * dx, y0 + y * dy]);
  }
  const img: RasterImage = { width: w, height: h, channels: outCh, data };
  if (alpha && alpha.some((a) => a !== 255)) img.alpha = alpha;
  return img;
};
