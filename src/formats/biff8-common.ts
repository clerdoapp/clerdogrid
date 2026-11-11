// Pieces shared by the binary workbook readers and writer: byte reading, style enumerations,
// OfficeArt drawing records (MS-ODRAW) for pictures, and grid geometry for anchors.
import type { BorderStyle, Chart, HAlign, Sheet, VAlign } from '../model/workbook';
import { inflateSync, unzlibSync } from 'fflate';
import { fileColWidth } from './xlsx-chart';

export const BORDER_STYLES: (BorderStyle | undefined)[] = [
  undefined, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
];
export const PATTERNS = [
  'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
  'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625',
];
export const H_ALIGNS: HAlign[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
export const V_ALIGNS: VAlign[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
export const UNDERLINES: Record<number, 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'> = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };

/** Set while reading an older (BIFF5) workbook: strings are single bytes in the file's code page. */
export const byteStrings: { decoder: TextDecoder | null } = { decoder: null };

export class Rd {
  o: number;
  readonly dv: DataView;
  constructor(readonly b: Uint8Array, start = 0) { this.o = start; this.dv = new DataView(b.buffer, b.byteOffset, b.byteLength); }
  get left(): number { return this.b.length - this.o; }
  private need(n: number): void { if (this.o + n > this.b.length) throw new RangeError('Record ends early.'); }
  u8(): number { this.need(1); return this.b[this.o++]!; }
  u16(): number { this.need(2); const v = this.dv.getUint16(this.o, true); this.o += 2; return v; }
  i16(): number { this.need(2); const v = this.dv.getInt16(this.o, true); this.o += 2; return v; }
  u32(): number { this.need(4); const v = this.dv.getUint32(this.o, true); this.o += 4; return v; }
  i32(): number { this.need(4); const v = this.dv.getInt32(this.o, true); this.o += 4; return v; }
  f64(): number { this.need(8); const v = this.dv.getFloat64(this.o, true); this.o += 8; return v; }
  bytes(n: number): Uint8Array { this.need(n); const v = this.b.subarray(this.o, this.o + n); this.o += n; return v; }
  skip(n: number): void { this.o = Math.min(this.b.length, this.o + n); }
  utf16(n: number): string {
    const m = Math.min(n, Math.floor(this.left / 2));
    let s = '';
    for (let i = 0; i < m; i++) s += String.fromCharCode(this.dv.getUint16(this.o + i * 2, true));
    this.o += m * 2;
    return s;
  }
  latin1(n: number): string {
    const m = Math.min(n, this.left);
    let s = '';
    for (let i = 0; i < m; i++) s += String.fromCharCode(this.b[this.o + i]!);
    this.o += m;
    return s;
  }
  /** Characters after a flags byte (fHighByte); code-page bytes in older files. */
  chars(n: number): string {
    if (byteStrings.decoder) return this.codepage(n);
    const hi = this.u8() & 1; return hi ? this.utf16(n) : this.latin1(n);
  }
  codepage(n: number): string { const m = Math.min(n, this.left); const s = byteStrings.decoder ? byteStrings.decoder.decode(this.b.subarray(this.o, this.o + m)) : this.latin1(m); if (byteStrings.decoder) this.o += m; return s; }
  /** XLUnicodeString: 16-bit count. Rich runs and phonetic data are skipped. */
  xlString(): string {
    const n = this.u16();
    if (byteStrings.decoder) return this.codepage(n);
    const flags = this.u8();
    const runs = flags & 8 ? this.u16() : 0;
    const ext = flags & 4 ? this.u32() : 0;
    const s = flags & 1 ? this.utf16(n) : this.latin1(n);
    this.skip(runs * 4 + ext);
    return s;
  }
  /** ShortXLUnicodeString: 8-bit count. */
  shortString(): string { const n = this.u8(); return this.chars(n); }
  /** XLWideString (xlsb): 32-bit count of UTF-16 code units. */
  wide(): string { const n = this.u32(); if (n === 0xffffffff) return ''; return this.utf16(n); }
}

export const hex = (r: number, g: number, b: number): string => '#' + [r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();

// ---------- Geometry ----------

export const colPx = (s: Sheet, c: number): number => (s.cols.get(c)?.hidden ? 0 : Math.round(fileColWidth(s.cols.get(c)?.width ?? s.defaultColWidth) * 7));
export const rowPx = (s: Sheet, r: number): number => (s.rows.get(r)?.hidden ? 0 : ((s.rows.get(r)?.height ?? s.defaultRowHeight) * 96) / 72);

/** Anchor from two cell corners with fractional offsets (columns in 1/1024, rows in 1/256). */
export const anchorFromCells = (s: Sheet, c1: number, dx1: number, r1: number, dy1: number, c2: number, dx2: number, r2: number, dy2: number): Chart['anchor'] => {
  const fx = (c: number, d: number) => (colPx(s, c) * Math.min(d, 1024)) / 1024;
  const fy = (r: number, d: number) => (rowPx(s, r) * Math.min(d, 256)) / 256;
  let w = 0, h = 0;
  for (let c = c1; c < c2 && c < c1 + 16384; c++) w += colPx(s, c);
  for (let r = r1; r < r2 && r < r1 + 1_048_576; r++) h += rowPx(s, r);
  return {
    col: c1, row: r1, dx: Math.round(fx(c1, dx1)), dy: Math.round(fy(r1, dy1)),
    width: Math.max(1, Math.round(w - fx(c1, dx1) + fx(c2, dx2))), height: Math.max(1, Math.round(h - fy(r1, dy1) + fy(r2, dy2))),
  };
};

/** The inverse: bottom-right cell and fractions for an anchor. */
export const cellsFromAnchor = (s: Sheet, a: Chart['anchor']): { c1: number; dx1: number; r1: number; dy1: number; c2: number; dx2: number; r2: number; dy2: number } => {
  const cw = (c: number) => Math.max(1, colPx(s, c));
  const rh = (r: number) => Math.max(1, rowPx(s, r));
  let x = a.dx + a.width, c = a.col;
  while (x >= cw(c) && c < 255) x -= cw(c++);
  let y = a.dy + a.height, r = a.row;
  while (y >= rh(r) && r < 65535) y -= rh(r++);
  return {
    c1: a.col, dx1: Math.min(1023, Math.round((a.dx / cw(a.col)) * 1024)), r1: a.row, dy1: Math.min(255, Math.round((a.dy / rh(a.row)) * 256)),
    c2: c, dx2: Math.min(1023, Math.round((x / cw(c)) * 1024)), r2: r, dy2: Math.min(255, Math.round((y / rh(r)) * 256)),
  };
};

// ---------- OfficeArt (MS-ODRAW) ----------

export interface ArtRecord { ver: number; inst: number; type: number; data: Uint8Array; children?: ArtRecord[] }

/** Splits OfficeArt records; containers (version 0xF) are parsed recursively. */
export const artRecords = (b: Uint8Array, depth = 0): ArtRecord[] => {
  const out: ArtRecord[] = [];
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let o = 0;
  while (o + 8 <= b.length) {
    const vi = dv.getUint16(o, true);
    const type = dv.getUint16(o + 2, true);
    const len = dv.getUint32(o + 4, true);
    const data = b.subarray(o + 8, Math.min(b.length, o + 8 + len));
    const rec: ArtRecord = { ver: vi & 0xf, inst: vi >> 4, type, data };
    if (rec.ver === 0xf && depth < 16) rec.children = artRecords(data, depth + 1);
    out.push(rec);
    o += 8 + len;
  }
  return out;
};

export interface Blip { data: Uint8Array; mime: string }

const BLIP_MIME: Record<number, string> = { 0xf01d: 'image/jpeg', 0xf02a: 'image/jpeg', 0xf01e: 'image/png', 0xf01f: 'image/bmp', 0xf029: 'image/tiff' };

/** Pictures in the drawing group's blip store, by 1-based index (null where a picture can't be shown). */
export const readBlipStore = (dgg: Uint8Array): (Blip | null)[] => {
  const out: (Blip | null)[] = [];
  const top = artRecords(dgg);
  const store = top.flatMap((r) => r.children ?? []).find((r) => r.type === 0xf001);
  for (const fbse of store?.children ?? []) {
    if (fbse.type !== 0xf007 || fbse.data.length < 36) { out.push(null); continue; }
    const cbName = fbse.data[33]!;
    const blipRecs = artRecords(fbse.data.subarray(36 + cbName));
    const blip = blipRecs[0];
    if (blip && (blip.type === 0xf01a || blip.type === 0xf01b)) { out.push(metafile(blip)); continue; }
    const mime = blip ? BLIP_MIME[blip.type] : undefined;
    if (!blip || !mime) { out.push(null); continue; }
    // one or two 16-byte ids, then a tag byte
    const skip = (blip.inst & 1 ? 32 : 16) + 1;
    let data = blip.data.subarray(skip);
    if (blip.type === 0xf01f) data = dibToBmp(data);
    out.push({ data: data.slice(), mime });
  }
  return out;
};

/** EMF or WMF picture: a metafile header, then (usually compressed) data. */
const metafile = (blip: ArtRecord): Blip | null => {
  const at = blip.inst & 1 ? 32 : 16;
  if (blip.data.length < at + 34) return null;
  const dv = new DataView(blip.data.buffer, blip.data.byteOffset, blip.data.byteLength);
  const size = dv.getUint32(at, true);
  const saved = dv.getUint32(at + 28, true);
  const compressed = blip.data[at + 32] === 0;
  const body = blip.data.subarray(at + 34, at + 34 + saved);
  let data: Uint8Array;
  try { data = compressed ? (body[0] === 0x78 ? unzlibSync(body) : inflateSync(body)) : body.slice(); } catch { return null; }
  if (size && data.length > size) data = data.subarray(0, size);
  return { data: data.slice(), mime: blip.type === 0xf01a ? 'image/x-emf' : 'image/x-wmf' };
};

/** Adds a file header to a device-independent bitmap so it can be shown as a .bmp. */
export const dibToBmp = (dib: Uint8Array): Uint8Array => {
  if (dib.length < 40) return dib;
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const headerSize = dv.getUint32(0, true);
  const bpp = dv.getUint16(14, true);
  const colors = dv.getUint32(32, true) || (bpp <= 8 ? 1 << bpp : 0);
  const compression = dv.getUint32(16, true);
  const masks = compression === 3 && headerSize === 40 ? 12 : 0;
  const out = new Uint8Array(14 + dib.length);
  const o = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  o.setUint32(2, out.length, true);
  o.setUint32(10, 14 + headerSize + masks + colors * 4, true);
  out.set(dib, 14);
  return out;
};

export interface ArtShape { spt?: number; fill?: string; line?: string; noFill?: boolean; pib?: number; anchor?: { c1: number; dx1: number; r1: number; dy1: number; c2: number; dx2: number; r2: number; dy2: number }; alt?: string; name?: string }

/** Shapes of a sheet drawing with their picture index and cell anchor. */
export const readShapes = (dg: Uint8Array): ArtShape[] => {
  const out: ArtShape[] = [];
  const visit = (recs: ArtRecord[], depth: number): void => {
    for (const r of recs) {
      if (r.type === 0xf004 && r.children) {
        const sh: ArtShape = {};
        for (const c of r.children) {
          if (c.type === 0xf00b || c.type === 0xf122) {
            const dv = new DataView(c.data.buffer, c.data.byteOffset, c.data.byteLength);
            const n = c.inst;
            let complexAt = n * 6;
            for (let i = 0; i < n && i * 6 + 6 <= c.data.length; i++) {
              const id = dv.getUint16(i * 6, true);
              const val = dv.getUint32(i * 6 + 2, true);
              const pid = id & 0x3fff;
              const complex = id & 0x8000;
              if (pid === 0x0104 && !complex) sh.pib = val;
              const bgr = (v: number) => '#' + [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
              if (pid === 0x0181 && !complex && !(val & 0xff000000)) sh.fill = bgr(val);
              if (pid === 0x01c0 && !complex && !(val & 0xff000000)) sh.line = bgr(val);
              if (pid === 0x01bf && !complex && (val & 0x00100000) && !(val & 0x10)) sh.noFill = true;
              if (complex) {
                const bytes = c.data.subarray(complexAt, complexAt + val);
                if (pid === 0x0381 || pid === 0x0380) {
                  let s = '';
                  for (let k = 0; k + 1 < bytes.length; k += 2) { const ch = bytes[k]! | (bytes[k + 1]! << 8); if (!ch) break; s += String.fromCharCode(ch); }
                  if (pid === 0x0381) sh.alt = s; else sh.name = s;
                }
                complexAt += val;
              }
            }
          } else if (c.type === 0xf00a) sh.spt = c.inst;
          else if (c.type === 0xf010 && c.data.length >= 18) {
            const dv = new DataView(c.data.buffer, c.data.byteOffset, c.data.byteLength);
            const g = (i: number) => dv.getUint16(2 + i * 2, true);
            sh.anchor = { c1: g(0), dx1: g(1), r1: g(2), dy1: g(3), c2: g(4), dx2: g(5), r2: g(6), dy2: g(7) };
          }
        }
        out.push(sh);
      } else if (r.children && depth < 16) visit(r.children, depth + 1);
    }
  };
  visit(artRecords(dg), 0);
  return out;
};
