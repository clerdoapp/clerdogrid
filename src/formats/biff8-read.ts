// Binary workbook reader (.xls, BIFF8 per the MS-XLS open specification) inside a compound file.
import {
  cellKey, internStyle, newSheet, type AutoFilter, type BorderSide, type Cell, type CellRange, type CellStyle, type CfOperator, type CfRule,
  type DataValidation, type FilterColumn, type RichRun, type Sheet, type Workbook, type Chart, type Shape, type ShapeKind,
} from '../model/workbook';
import { CellError, type Scalar } from '../engine/types';
import { quoteSheet } from '../engine/parser';
import { cfbGet, readCfb } from './cfb';
import { DEFAULT_PASSWORD, StreamCipher, checkPassword, xorArray, xorDecode, type Rc4Params } from './biff8-crypto';
import { ERR_BY_CODE, Exp, decodeRgce, type DecodeCtx } from './biff8-formula';
import {
  BORDER_STYLES, H_ALIGNS, PATTERNS, Rd, UNDERLINES, V_ALIGNS, anchorFromCells, byteStrings, hex, readBlipStore, readShapes,
} from './biff8-common';
import { INDEXED_COLORS, builtinFormat } from './numfmt';
import { DEFAULT_THEME, applyTint } from './xlsx-styles';
import { shownColWidth } from './xlsx-chart';
import { queriesFromXml } from './queries-xml';
import { readXlsChart, type ChartRec } from './biff8-chart';

export class PasswordRequired extends Error {
  constructor(readonly wrong: boolean) { super(wrong ? 'That password is not right.' : 'This workbook is protected with a password.'); }
}

const SAFE_LINK = /^(https:|mailto:|#)/i;

interface Rec { type: number; data: Uint8Array; pos: number; cont: Uint8Array[] }

// Record types
const BOF = 0x0809, EOF = 0x000a, CONTINUE = 0x003c, FILEPASS = 0x002f;
const NO_ENCRYPT = new Set([BOF, FILEPASS, 0x0194 /* USREXCL */, 0x0195 /* FILELOCK */, 0x00e1 /* INTERFACEHDR */, 0x0196 /* RRDINFO */, 0x0138 /* RRDHEAD */]);

const parseFilepass = (d: Uint8Array, biff5: boolean): Rc4Params | { key: number; verifier: number } | null => {
  const r = new Rd(d);
  if (biff5) return { key: r.u16(), verifier: r.u16() };
  const type = r.u16();
  if (type === 0) return { key: r.u16(), verifier: r.u16() };
  const major = r.u16(), minor = r.u16();
  if (major === 1 && minor === 1) {
    return { kind: 'rc4', salt: r.bytes(16).slice(), verifier: r.bytes(16).slice(), verifierHash: r.bytes(16).slice(), keyBits: 40 };
  }
  if ((major === 2 || major === 3 || major === 4) && minor === 2) {
    r.u32(); // flags
    const hsize = r.u32();
    const hstart = r.o;
    r.u32(); r.u32(); r.u32(); r.u32();
    const keyBits = r.u32();
    r.o = hstart + hsize;
    const saltSize = r.u32();
    const salt = r.bytes(saltSize).slice();
    const verifier = r.bytes(16).slice();
    const hashSize = r.u32();
    return { kind: 'cryptoapi', salt, verifier, verifierHash: r.bytes(hashSize).slice(), keyBits: keyBits || 40 };
  }
  return null;
};

/** Splits the stream into records, decrypting after FILEPASS. CONTINUE records attach to the record before. */
let lastKey: ReturnType<typeof checkPassword> = null;
const splitRecords = (stream: Uint8Array, password?: string): Rec[] => {
  lastKey = null;
  const s = stream.slice();
  const dv = new DataView(s.buffer);
  const out: Rec[] = [];
  let cipher: StreamCipher | null = null;
  let xor: Uint8Array | null = null;
  const biff5 = s.length >= 8 && dv.getUint16(0, true) === BOF && dv.getUint16(4, true) !== 0x0600;
  let o = 0;
  while (o + 4 <= s.length) {
    const type = dv.getUint16(o, true);
    const size = dv.getUint16(o + 2, true);
    const end = Math.min(s.length, o + 4 + size);
    if ((cipher || xor) && !NO_ENCRYPT.has(type)) {
      const from = type === 0x0085 ? o + 8 : o + 4;
      if (end > from) { if (cipher) cipher.apply(s, from, end - from, from); else xorDecode(s, from, end, size, xor!); }
    }
    const data = s.subarray(o + 4, end);
    if (type === FILEPASS && !cipher && !xor) {
      const p = parseFilepass(data, biff5);
      if (p === null) throw new Error('This workbook uses an encryption ClerdoGrid cannot open.');
      if ('kind' in p) {
        const key = checkPassword(p, DEFAULT_PASSWORD) ?? (password !== undefined ? checkPassword(p, password) : null);
        if (!key) throw new PasswordRequired(password !== undefined);
        cipher = new StreamCipher(key);
        lastKey = key;
      } else {
        xor = xorArray(DEFAULT_PASSWORD, p.key, p.verifier) ?? (password !== undefined ? xorArray(password, p.key, p.verifier) : null);
        if (!xor) throw new PasswordRequired(password !== undefined);
      }
    }
    if (type === CONTINUE && out.length) out[out.length - 1]!.cont.push(data);
    else out.push({ type, data, pos: o, cont: [] });
    if (type === 0 && size === 0 && out.length > 2 && s.length - o > 16 && s.subarray(o, o + 16).every((x) => x === 0)) break;
    o = end;
  }
  return out;
};

/** Reads across a record and its CONTINUE records; character runs re-read the flags byte at each boundary. */
class ContReader {
  private seg = 0;
  private o = 0;
  constructor(private parts: Uint8Array[]) {}
  private cur(): Uint8Array | undefined {
    while (this.seg < this.parts.length && this.o >= this.parts[this.seg]!.length) { this.seg++; this.o = 0; }
    return this.parts[this.seg];
  }
  get done(): boolean { return !this.cur(); }
  u8(): number { const p = this.cur(); if (!p) throw new RangeError('end'); return p[this.o++]!; }
  u16(): number { return this.u8() | (this.u8() << 8); }
  u32(): number { return (this.u16() | (this.u16() << 16)) >>> 0; }
  skip(n: number): void { for (let i = 0; i < n; i++) { const p = this.cur(); if (!p) return; const k = Math.min(n - i, p.length - this.o); this.o += k; i += k - 1; } }
  chars(n: number, hi: boolean): string {
    let s = '';
    let high = hi;
    while (s.length < n) {
      let p = this.cur();
      if (!p) break;
      if (this.o === 0 && s.length > 0) { high = !!(p[this.o++]! & 1); p = this.cur(); if (!p) break; }
      const need = n - s.length;
      if (high) {
        const k = Math.min(need, Math.floor((p.length - this.o) / 2));
        for (let i = 0; i < k; i++) s += String.fromCharCode(p[this.o + i * 2]! | (p[this.o + i * 2 + 1]! << 8));
        this.o += k * 2;
        if (k === 0) this.o = p.length;
      } else {
        const k = Math.min(need, p.length - this.o);
        for (let i = 0; i < k; i++) s += String.fromCharCode(p[this.o + i]!);
        this.o += k;
      }
      if (s.length < n) { this.seg++; this.o = 0; }
    }
    return s;
  }
  /** XLUnicodeRichExtendedString. */
  richString(): { text: string; runs: [number, number][] } {
    const n = this.u16();
    const flags = this.u8();
    const cRun = flags & 8 ? this.u16() : 0;
    const ext = flags & 4 ? this.u32() : 0;
    const text = this.chars(n, !!(flags & 1));
    const runs: [number, number][] = [];
    for (let i = 0; i < cRun; i++) runs.push([this.u16(), this.u16()]);
    this.skip(ext);
    return { text, runs };
  }
}

interface Font { name: string; size: number; bold: boolean; italic: boolean; strike: boolean; underline?: RichRun['font'] extends infer F ? F extends { underline?: infer U } ? U : never : never; color?: string; vert?: 'superscript' | 'subscript' }
interface Xf { font: number; fmt: number; locked: boolean; hidden: boolean; h: number; wrap: boolean; v: number; rot: number; indent: number; shrink: boolean; order: number;
  bl: number; br: number; bt: number; bb: number; cl: number; cr: number; ct: number; cb: number; diag: number; dgDiag: number; cDiag: number; fls: number; fore: number; back: number;
  ext?: Partial<Record<'fore' | 'back' | 'text' | 'top' | 'bottom' | 'left' | 'right', string>> }

interface Supbook { self: boolean; addin: boolean; names: string[] }

export interface XlsReadOptions { password?: string }

/** Reads an .xls workbook. Throws PasswordRequired when it is encrypted with a password other than the default one. */
export const readXls = (bytes: Uint8Array, opts: XlsReadOptions = {}): Workbook => {
  const streams = readCfb(bytes);
  if (cfbGet(streams, 'EncryptedPackage')) throw new Error('ENCRYPTED_PACKAGE');
  const stream = cfbGet(streams, 'Workbook') ?? cfbGet(streams, 'Book');
  if (!stream) throw new Error("This file isn't a workbook ClerdoGrid can read.");
  const recs = splitRecords(stream, opts.password);
  const first = recs[0];
  const version = first && first.type === BOF ? new Rd(first.data).u16() : 0;
  if (version !== 0x0600 && version !== 0x0500) throw new Error("This workbook's format is too old for ClerdoGrid to read.");
  const b5 = version === 0x0500;
  // older files store text as single bytes in the workbook's code page
  const cpRec = b5 ? recs.find((r) => r.type === 0x0042) : undefined;
  const cp = cpRec ? new Rd(cpRec.data).u16() : 1252;
  byteStrings.decoder = b5 ? codepageDecoder(cp) : null;
  try { return readBook(); } finally { byteStrings.decoder = null; }

  function readBook(): Workbook {

  const wb: Workbook = { sheets: [], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const palette = INDEXED_COLORS.slice();
  const fonts: Font[] = [];
  const formats = new Map<number, string>();
  const xfs: Xf[] = [];
  const sst: { text: string; runs: [number, number][] }[] = [];
  const bound: { pos: number; state: number; dt: number; name: string }[] = [];
  const supbooks: Supbook[] = [];
  const xti: { book: number; first: number; last: number }[] = [];
  const rawNames: { flags: number; name: string; builtin: number; itab: number; rgce: Uint8Array; rgcb: Uint8Array }[] = [];
  let dgg: Uint8Array | null = null;
  let activeBound = 0;
  let wbProtect = false, wbPassword = 0;
  const calc: NonNullable<Workbook['calc']> = {};

  let i = 1;
  for (; i < recs.length; i++) {
    const rec = recs[i]!;
    const r = new Rd(rec.data);
    try {
      switch (rec.type) {
        case EOF: break;
        case 0x0022: wb.date1904 = r.u16() === 1; break;
        case 0x0031: {
          const size = r.u16() / 20, grbit = r.u16(), icv = r.u16(), bls = r.u16(), sss = r.u16(), uls = r.u8();
          r.skip(3);
          fonts.push({ name: r.shortString(), size, bold: bls >= 700, italic: !!(grbit & 2), strike: !!(grbit & 8), underline: UNDERLINES[uls], color: icv === 0x7fff ? undefined : String(icv), vert: sss === 1 ? 'superscript' : sss === 2 ? 'subscript' : undefined });
          break;
        }
        case 0x041e: { const id = r.u16(); formats.set(id, b5 ? r.shortString() : r.xlString()); break; }
        case 0x00e0: {
          if (b5) {
            const font = r.u16(), fmt = r.u16(), f1 = r.u16(), al = r.u8(), ori = r.u8() & 3;
            const b1 = r.u32(), b2 = r.u32();
            xfs.push({
              font, fmt, locked: !!(f1 & 1), hidden: !!(f1 & 2), h: al & 7, wrap: !!(al & 8), v: (al >> 4) & 7, rot: ori === 1 ? 255 : ori === 2 ? 90 : ori === 3 ? 180 : 0,
              indent: 0, shrink: false, order: 0,
              fore: b1 & 0x7f, back: (b1 >> 7) & 0x7f, fls: (b1 >> 16) & 0x3f, bb: (b1 >> 22) & 7, cb: (b1 >>> 25) & 0x7f,
              bt: b2 & 7, bl: (b2 >> 3) & 7, br: (b2 >> 6) & 7, ct: (b2 >> 9) & 0x7f, cl: (b2 >> 16) & 0x7f, cr: (b2 >> 23) & 0x7f, diag: 0, dgDiag: 0, cDiag: 0,
            });
            break;
          }
          const font = r.u16(), fmt = r.u16(), f1 = r.u16(), al = r.u8(), rot = r.u8(), f2 = r.u8();
          r.u8();
          const b1 = r.u32(), b2 = r.u32(), b3 = r.u16();
          xfs.push({
            font, fmt, locked: !!(f1 & 1), hidden: !!(f1 & 2), h: al & 7, wrap: !!(al & 8), v: (al >> 4) & 7, rot, indent: f2 & 0xf, shrink: !!(f2 & 0x10), order: (f2 >> 6) & 3,
            bl: b1 & 0xf, br: (b1 >> 4) & 0xf, bt: (b1 >> 8) & 0xf, bb: (b1 >> 12) & 0xf, cl: (b1 >> 16) & 0x7f, cr: (b1 >> 23) & 0x7f, diag: (b1 >>> 30) & 3,
            ct: b2 & 0x7f, cb: (b2 >> 7) & 0x7f, cDiag: (b2 >> 14) & 0x7f, dgDiag: (b2 >> 21) & 0xf, fls: (b2 >>> 26) & 0x3f, fore: b3 & 0x7f, back: (b3 >> 7) & 0x7f,
          });
          break;
        }
        case 0x087d: readXfExt(r, xfs); break;
        case 0x0092: { const n = r.u16(); for (let k = 0; k < n && k < 56; k++) { const b = r.bytes(4); palette[8 + k] = hex(b[0]!, b[1]!, b[2]!); } break; }
        case 0x0085: { const pos = r.u32(), state = r.u8() & 3, dt = r.u8(); bound.push({ pos, state, dt, name: r.shortString() }); break; }
        case 0x00fc: {
          const cr = new ContReader([rec.data, ...rec.cont]);
          cr.u32();
          const n = cr.u32();
          for (let k = 0; k < n && !cr.done; k++) sst.push(cr.richString());
          break;
        }
        case 0x01ae: {
          const ctab = r.u16(), cch = r.u16();
          supbooks.push({ self: cch === 0x0401, addin: cch === 0x3a01 || (ctab === 1 && cch === 0x3a01), names: [] });
          break;
        }
        case 0x0023: {
          const sb = supbooks.at(-1);
          if (sb) { r.u16(); r.u32(); sb.names.push(r.shortString()); }
          break;
        }
        case 0x0017: { const n = r.u16(); for (let k = 0; k < n; k++) xti.push({ book: r.u16(), first: r.i16(), last: r.i16() }); break; }
        case 0x0018: {
          const flags = r.u16(); r.u8(); const cch = r.u8(); const cce = r.u16(); r.u16(); const itab = r.u16(); r.skip(4);
          const builtin = flags & 0x20;
          let name: string;
          if (b5) name = r.codepage(cch);
          else { const hi = r.u8() & 1; name = hi ? r.utf16(cch) : r.latin1(cch); }
          const code = builtin ? name.charCodeAt(0) : -1;
          if (builtin) name = BUILTIN_NAMES[code] ?? name;
          const rest = rec.cont.length ? concat(rec.data.subarray(r.o), ...rec.cont) : rec.data.subarray(r.o);
          rawNames.push({ flags, name, builtin: code, itab, rgce: rest.subarray(0, cce), rgcb: rest.subarray(cce) });
          break;
        }
        case 0x003d: r.skip(10); activeBound = r.u16(); break;
        case 0x0012: wbProtect = r.u16() === 1; break;
        case 0x0013: wbPassword = r.u16(); break;
        case 0x00eb: dgg = concat(rec.data, ...rec.cont); break;
        default: break;
      }
    } catch { /* a damaged record is skipped */ }
    if (rec.type === EOF) break;
  }

  const color = (icv: number): string | undefined => (icv >= 64 || icv === 0x7fff ? undefined : palette[icv]);
  const fontAt = (i: number): Font | undefined => fonts[i >= 4 ? i - 1 : i];
  const fontStyle = (f: Font | undefined): CellStyle['font'] => {
    if (!f) return undefined;
    const out: NonNullable<CellStyle['font']> = {};
    if (f.name && f.name !== 'Calibri') out.name = f.name;
    if (f.size && f.size !== 11) out.size = f.size;
    if (f.bold) out.bold = true;
    if (f.italic) out.italic = true;
    if (f.strike) out.strike = true;
    if (f.underline) out.underline = f.underline;
    if (f.vert) out.vertAlign = f.vert;
    const c = f.color !== undefined ? color(+f.color) : undefined;
    if (c && c !== '#000000') out.color = c;
    return Object.keys(out).length ? out : undefined;
  };
  const styleCache = new Map<number, number>();
  const styleOf = (ixfe: number): number => {
    const hit = styleCache.get(ixfe);
    if (hit !== undefined) return hit;
    const x = xfs[ixfe];
    if (!x) return 0;
    const st: CellStyle = {};
    const f = fontStyle(fontAt(x.font));
    if (f) st.font = f;
    if (x.ext?.text && st.font?.color !== undefined) st.font.color = x.ext.text;
    else if (x.ext?.text && x.ext.text !== '#000000') st.font = { ...st.font, color: x.ext.text };
    const code = formats.get(x.fmt) ?? builtinFormat(x.fmt);
    if (code && code !== 'General') st.numFmt = code;
    if (x.fls === 1) {
      const c = x.ext?.fore ?? color(x.fore) ?? (x.fore === 64 ? '#000000' : undefined);
      if (c) st.fill = { color: c };
    } else if (x.fls > 1) {
      const fill: NonNullable<CellStyle['fill']> = { pattern: PATTERNS[x.fls] ?? 'gray125' };
      const bg = x.ext?.back ?? color(x.back), fg = x.ext?.fore ?? color(x.fore);
      if (bg) fill.color = bg;
      if (fg) fill.patternColor = fg;
      st.fill = fill;
    }
    const side = (dg: number, icv: number, ext?: string): BorderSide | undefined => {
      const style = BORDER_STYLES[dg];
      if (!style) return undefined;
      const c = ext ?? color(icv);
      return c && c !== '#000000' ? { style, color: c } : { style };
    };
    const border: NonNullable<CellStyle['border']> = {};
    const L = side(x.bl, x.cl, x.ext?.left), R = side(x.br, x.cr, x.ext?.right), T = side(x.bt, x.ct, x.ext?.top), B = side(x.bb, x.cb, x.ext?.bottom);
    if (L) border.left = L;
    if (R) border.right = R;
    if (T) border.top = T;
    if (B) border.bottom = B;
    const D = side(x.dgDiag, x.cDiag);
    if (D && x.diag & 1) border.diagonalDown = D;
    if (D && x.diag & 2) border.diagonalUp = D;
    if (Object.keys(border).length) st.border = border;
    const align: NonNullable<CellStyle['align']> = {};
    if (x.h) align.h = H_ALIGNS[x.h];
    if (x.v !== 2 && V_ALIGNS[x.v]) align.v = V_ALIGNS[x.v];
    if (x.wrap) align.wrap = true;
    if (x.indent) align.indent = x.indent;
    if (x.rot) align.rotation = x.rot > 90 && x.rot <= 180 ? 90 - x.rot : x.rot;
    if (x.shrink) align.shrink = true;
    if (x.order === 2) align.rtl = true;
    if (Object.keys(align).length) st.align = align;
    if (!x.locked || x.hidden) st.protection = { ...(x.locked ? {} : { locked: false }), ...(x.hidden ? { hidden: true } : {}) };
    const id = internStyle(wb, st);
    styleCache.set(ixfe, id);
    return id;
  };

  const richOf = (text: string, runs: [number, number][]): RichRun[] | undefined => {
    if (!runs.length) return undefined;
    const out: RichRun[] = [];
    const sorted = runs.filter(([ich]) => ich < text.length).sort((a, b) => a[0] - b[0]);
    if (!sorted.length) return undefined;
    if (sorted[0]![0] > 0) out.push({ text: text.slice(0, sorted[0]![0]) });
    sorted.forEach(([ich, ifnt], k) => {
      const end = sorted[k + 1]?.[0] ?? text.length;
      const f = fontStyle(fontAt(ifnt));
      out.push(f ? { text: text.slice(ich, end), font: f } : { text: text.slice(ich, end) });
    });
    return out.filter((r) => r.text);
  };

  // worksheets and chart sheets; macro sheets and modules are left out
  const sheetOfBound = new Map<number, number>();
  bound.forEach((b, k) => { if (b.dt === 0 || (b.dt === 2 && !b5)) sheetOfBound.set(k, sheetOfBound.size); });
  const sheetName = (k: number): string | undefined => bound[k]?.name;
  const ctxFor = (row: number, col: number): DecodeCtx => ({
    xlsb: false, row, col,
    ...(b5 ? { biff5: {
      text: (b: Uint8Array) => (byteStrings.decoder ?? new TextDecoder('windows-1252')).decode(b),
      sheetsByTab: (x: number, y: number) => { const a2 = sheetName(x), b2 = sheetName(y); return a2 ? { first: a2, last: b2 } : null; },
    } } : {}),
    sheets: (ixti) => {
      const x = xti[ixti];
      const sb = x ? supbooks[x.book] : undefined;
      if (!x || !sb?.self || x.first < 0) return null;
      const a = sheetName(x.first), b = sheetName(x.last);
      return a ? { first: a, last: b } : null;
    },
    name: (idx) => {
      const n = rawNames[idx - 1];
      if (!n) return undefined;
      if (n.builtin >= 0) return { name: `_xlnm.${n.name}`, sheet: n.itab ? sheetName(n.itab - 1) : undefined };
      return { name: n.name, sheet: n.itab ? sheetName(n.itab - 1) : undefined };
    },
    externName: (ixti, idx) => {
      const x = xti[ixti];
      const sb = x ? supbooks[x.book] : undefined;
      return sb?.names[idx - 1];
    },
  });
  const decode = (rgce: Uint8Array, rgcb: Uint8Array, row: number, col: number): string | Exp | undefined => {
    try { return decodeRgce(rgce, rgcb, ctxFor(row, col)); } catch { return undefined; }
  };

  const blips = dgg ? readBlipStore(dgg) : [];
  const recAt = new Map<number, number>();
  recs.forEach((r, k) => recAt.set(r.pos, k));
  const filterRanges = new Map<number, CellRange>();
  const printTitles = new Map<number, string>();

  bound.forEach((b) => {
    if (b.dt === 2 && !b5) { readChartSheet(b); return; }
    if (b.dt !== 0) return;
    const start = recAt.get(b.pos);
    const sheet = newSheet(b.name.slice(0, 31) || `Sheet${wb.sheets.length + 1}`);
    sheet.state = b.state === 1 ? 'hidden' : b.state === 2 ? 'veryHidden' : 'visible';
    wb.sheets.push(sheet);
    if (start === undefined) return;
    readSheet(start, sheet, wb.sheets.length - 1);
  });

  function readChartSheet(b: (typeof bound)[number]): void {
    const sheet = newSheet(b.name.slice(0, 31) || `Chart${wb.sheets.length + 1}`);
    sheet.state = b.state === 1 ? 'hidden' : b.state === 2 ? 'veryHidden' : 'visible';
    sheet.kind = 'chart';
    sheet.view.showGrid = false;
    sheet.view.showHeaders = false;
    wb.sheets.push(sheet);
    const start = recAt.get(b.pos);
    if (start === undefined) return;
    const sub: ChartRec[] = [];
    let depth = 1;
    for (let j = start + 1; j < recs.length && depth; j++) {
      const x = recs[j]!;
      if (x.type === BOF) depth++; else if (x.type === EOF) depth--;
      sub.push({ type: x.type, data: x.cont.length ? concat(x.data, ...x.cont) : x.data });
    }
    try {
      const c = readXlsChart(sub, (rgce) => { const f = decode(rgce, new Uint8Array(0), 0, 0); return typeof f === 'string' ? f : undefined; });
      if (c) sheet.charts = [{ id: 'chart1', anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 910, height: 660 }, ...c }];
    } catch { /* an unreadable chart leaves an empty chart sheet */ }
  }

  function readSheet(start: number, sheet: Sheet, sheetIdx: number): void {
    const shared = new Map<string, { range: CellRange; rgce: Uint8Array; rgcb: Uint8Array; array: boolean }>();
    const expCells: { row: number; col: number; at: Exp }[] = [];
    let pendingString: Cell | null = null;
    const drawing: Uint8Array[] = [];
    const objText = new Map<number, string>();
    const notes: { row: number; col: number; id: number; author: string }[] = [];
    let lastObj = -1;
    const objs: { ot: number; id: number; chart?: Omit<Chart, 'id' | 'anchor'> }[] = [];
    let protect = false, password = 0, protOptions: Record<string, boolean> | undefined;
    let frozen = false, paneX = 0, paneY = 0;
    let cfHeader: { ranges: CellRange[]; base: { row: number; col: number }; rules: CfRule[] } | null = null;
    const filterCols: FilterColumn[] = [];
    const ps: NonNullable<Sheet['pageSetup']> = {};
    const margins: Partial<NonNullable<NonNullable<Sheet['pageSetup']>['margins']>> = {};
    let fitToPage = false;
    let setup: { paper: number; scale: number; fitW: number; fitH: number; portrait: boolean; valid: boolean; header: number; footer: number } | null = null;

    const put = (row: number, col: number, cell: Cell) => { if (row < 65536 && col < 256) sheet.cells.set(cellKey(row, col), cell); };
    const withStyle = (cell: Cell, ixfe: number): Cell => { const s = styleOf(ixfe); if (s) cell.s = s; return cell; };

    for (let j = start + 1; j < recs.length; j++) {
      const rec = recs[j]!;
      if (rec.type === EOF) break;
      if (rec.type === BOF) {
        // embedded chart substream: read it for the drawing object just before it
        const sub: ChartRec[] = [];
        let depth = 1;
        while (++j < recs.length && depth) {
          const x = recs[j]!;
          if (x.type === BOF) depth++; else if (x.type === EOF) depth--;
          sub.push({ type: x.type, data: x.cont.length ? concat(x.data, ...x.cont) : x.data });
        }
        j--;
        const obj = objs.at(-1);
        if (obj && obj.ot === 5 && !b5) {
          try { obj.chart = readXlsChart(sub, (rgce) => { const f = decode(rgce, new Uint8Array(0), 0, 0); return typeof f === 'string' ? f : undefined; }) ?? undefined; } catch { /* skip a damaged chart */ }
        }
        continue;
      }
      const r = new Rd(rec.data);
      try {
        switch (rec.type) {
          case 0x0203: { const row = r.u16(), col = r.u16(), x = r.u16(); put(row, col, withStyle({ v: r.f64() }, x)); break; }
          case 0x027e: { const row = r.u16(), col = r.u16(), x = r.u16(); put(row, col, withStyle({ v: rk(r.u32()) }, x)); break; }
          case 0x00bd: {
            const row = r.u16(), c1 = r.u16();
            const n = (rec.data.length - 6) / 6;
            for (let k = 0; k < n; k++) { const x = r.u16(); put(row, c1 + k, withStyle({ v: rk(r.u32()) }, x)); }
            break;
          }
          case 0x00fd: {
            const row = r.u16(), col = r.u16(), x = r.u16(), s = sst[r.u32()];
            const cell: Cell = withStyle({ v: s?.text ?? '' }, x);
            const rich = s ? richOf(s.text, s.runs) : undefined;
            if (rich) cell.rich = rich;
            put(row, col, cell);
            break;
          }
          case 0x0204: case 0x00d6: { const row = r.u16(), col = r.u16(), x = r.u16(); put(row, col, withStyle({ v: r.xlString() }, x)); break; }
          case 0x0205: {
            const row = r.u16(), col = r.u16(), x = r.u16(), v = r.u8(), isErr = r.u8();
            put(row, col, withStyle({ v: isErr ? new CellError(ERR_BY_CODE[v] ?? '#N/A') : !!v }, x));
            break;
          }
          case 0x0201: { const row = r.u16(), col = r.u16(), x = r.u16(); const s = styleOf(x); if (s) put(row, col, { v: null, s }); break; }
          case 0x00be: {
            const row = r.u16(), c1 = r.u16();
            const n = (rec.data.length - 6) / 2;
            for (let k = 0; k < n; k++) { const s = styleOf(r.u16()); if (s) put(row, c1 + k, { v: null, s }); }
            break;
          }
          case 0x0006: {
            const row = r.u16(), col = r.u16(), x = r.u16();
            const val = rec.data.subarray(6, 14);
            r.skip(8);
            r.u16(); r.u32();
            const cce = r.u16();
            const rgce = rec.data.subarray(r.o, r.o + cce);
            const rgcb = rec.data.subarray(r.o + cce);
            let v: Scalar = null;
            let wantString = false;
            if (val[6] === 0xff && val[7] === 0xff) {
              if (val[0] === 0) wantString = true;
              else if (val[0] === 1) v = !!val[2];
              else if (val[0] === 2) v = new CellError(ERR_BY_CODE[val[2]!] ?? '#N/A');
              else if (val[0] === 3) v = '';
            } else v = new DataView(val.buffer, val.byteOffset, 8).getFloat64(0, true);
            const cell = withStyle({ v }, x);
            const f = decode(rgce, rgcb, row, col);
            if (f instanceof Exp) expCells.push({ row, col, at: f });
            else if (f) cell.f = f;
            put(row, col, cell);
            pendingString = wantString ? cell : null;
            break;
          }
          case 0x0207: if (pendingString) { pendingString.v = r.xlString(); pendingString = null; } break;
          case 0x04bc: case 0x0221: {
            const r1 = r.u16(), r2 = r.u16(), c1 = r.u8(), c2 = r.u8();
            if (rec.type === 0x04bc) { r.u8(); r.u8(); } else { r.u16(); r.u32(); }
            const cce = r.u16();
            const all = rec.cont.length ? concat(rec.data, ...rec.cont) : rec.data;
            shared.set(`${r1},${c1}`, { range: { r1, r2, c1, c2 }, rgce: all.subarray(r.o, r.o + cce), rgcb: all.subarray(r.o + cce), array: rec.type === 0x0221 });
            break;
          }
          case 0x0208: {
            const row = r.u16(); r.skip(4);
            const h = r.u16(); r.skip(4);
            const fl = r.u16(), fl2 = r.u16();
            const info: NonNullable<ReturnType<Sheet['rows']['get']>> = {};
            if (fl & 0x40 || (h & 0x7fff) !== 255) { if (fl & 0x40) info.height = (h & 0x7fff) / 20; }
            if (fl & 0x20) info.hidden = true;
            if (fl & 7) info.outline = fl & 7;
            if (fl & 0x80) { const s = styleOf(fl2 & 0xfff); if (s) info.s = s; }
            if (Object.keys(info).length) sheet.rows.set(row, info);
            break;
          }
          case 0x007d: {
            const c1 = r.u16(), c2 = r.u16(), w = r.u16(), x = r.u16(), fl = r.u16();
            const info: NonNullable<ReturnType<Sheet['cols']['get']>> = { width: shownColWidth(w / 256) };
            if (fl & 1) info.hidden = true;
            const lvl = (fl >> 8) & 7;
            if (lvl) info.outline = lvl;
            const s = styleOf(x);
            if (s) info.s = s;
            for (let c = c1; c <= Math.min(c2, 255); c++) sheet.cols.set(c, { ...info });
            break;
          }
          case 0x0055: sheet.defaultColWidth = r.u16() + 0.43; break;
          case 0x0099: sheet.defaultColWidth = shownColWidth(r.u16() / 256); break;
          case 0x0225: { r.u16(); const h = r.u16(); if (h) sheet.defaultRowHeight = h / 20; break; }
          case 0x00e5: {
            const n = r.u16();
            for (let k = 0; k < n; k++) { const r1 = r.u16(), r2 = r.u16(), c1 = r.u16(), c2 = r.u16(); sheet.merges.push({ r1, r2, c1, c2 }); }
            break;
          }
          case 0x01b8: readHlink(rec.data, sheet); break;
          case 0x001c: {
            if (b5) {
              // older notes carry their text; a row of 0xFFFF continues the previous note
              const row = r.u16(), col = r.u16(), cch = r.u16();
              const text = r.codepage(Math.min(cch, rec.data.length - 6)).replace(/\r\n?/g, '\n');
              if (row === 0xffff && notes.length) { const last = notes.at(-1)!; objText.set(last.id, (objText.get(last.id) ?? '') + text); }
              else { const id = -(notes.length + 1); notes.push({ row, col, id, author: '' }); objText.set(id, text); }
              break;
            }
            const row = r.u16(), col = r.u16(); r.u16(); const id = r.u16(); notes.push({ row, col, id, author: r.xlString() }); break;
          }
          case 0x005d: { const ft = r.u16(); r.u16(); if (ft === 0x15) { const ot = r.u16(); lastObj = r.u16(); objs.push({ ot, id: lastObj }); } break; }
          case 0x01b6: {
            r.skip(10);
            const cch = r.u16();
            if (cch && rec.cont.length) {
              const cr = new ContReader(rec.cont);
              const hi = !!(cr.u8() & 1);
              objText.set(lastObj, cr.chars(cch, hi).replace(/\r\n?/g, '\n'));
            }
            break;
          }
          case 0x00ec: drawing.push(rec.data, ...rec.cont); break;
          case 0x023e: {
            const g = r.u16();
            sheet.view.showGrid = !!(g & 2);
            sheet.view.showHeaders = !!(g & 4);
            sheet.view.rtl = !!(g & 0x40);
            frozen = !!(g & 8);
            break;
          }
          case 0x00a0: { const num = r.i16(), den = r.i16(); if (num > 0 && den > 0) sheet.view.zoom = Math.max(10, Math.min(400, Math.round((num / den) * 100))); break; }
          case 0x0041: paneX = r.u16(); paneY = r.u16(); break;
          case 0x001d: {
            r.u8(); const row = r.u16(), col = r.u16(); r.u16(); const n = r.u16();
            if (n) { const r1 = r.u16(), r2 = r.u16(), c1 = r.u8(), c2 = r.u8(); sheet.view.selection = { r1, r2, c1, c2 }; }
            else sheet.view.selection = { r1: row, r2: row, c1: col, c2: col };
            break;
          }
          case 0x0012: protect = r.u16() === 1; break;
          case 0x0013: password = r.u16(); break;
          case 0x0867: {
            r.skip(12);
            if (r.u16() !== 2) break;
            r.skip(5);
            const f = r.u16();
            const allowed = (bit: number) => !!(f & (1 << bit));
            protOptions = {
              objects: !allowed(0), scenarios: !allowed(1), formatCells: !allowed(2), formatColumns: !allowed(3), formatRows: !allowed(4),
              insertColumns: !allowed(5), insertRows: !allowed(6), insertHyperlinks: !allowed(7), deleteColumns: !allowed(8), deleteRows: !allowed(9),
              selectLockedCells: !allowed(10), sort: !allowed(11), autoFilter: !allowed(12), pivotTables: !allowed(13), selectUnlockedCells: !allowed(14),
            };
            break;
          }
          case 0x01b0: {
            r.u16(); r.u16();
            const b = { r1: r.u16(), r2: r.u16(), c1: r.u16(), c2: r.u16() };
            const n = r.u16();
            const ranges: CellRange[] = [];
            for (let k = 0; k < n; k++) { const r1 = r.u16(), r2 = r.u16(), c1 = r.u16(), c2 = r.u16(); ranges.push({ r1, r2, c1, c2 }); }
            cfHeader = { ranges, base: { row: b.r1, col: b.c1 }, rules: [] };
            sheet.conditionalFormats.push({ ranges, rules: cfHeader.rules });
            break;
          }
          case 0x01b1: {
            if (!cfHeader) break;
            const ct = r.u8(), cp = r.u8(), n1 = r.u16(), n2 = r.u16();
            const style = readDxfn(r, color, fontAt, formats);
            const f1 = n1 ? decode(rec.data.subarray(r.o, r.o + n1), new Uint8Array(0), cfHeader.base.row, cfHeader.base.col) : undefined;
            const f2 = n2 ? decode(rec.data.subarray(r.o + n1, r.o + n1 + n2), new Uint8Array(0), cfHeader.base.row, cfHeader.base.col) : undefined;
            if (ct === 1 && CF_OPS[cp] && typeof f1 === 'string') {
              const formulas = [f1];
              if (typeof f2 === 'string' && (cp === 1 || cp === 2)) formulas.push(f2);
              cfHeader.rules.push({ type: 'cellIs', operator: CF_OPS[cp]!, formulas, style });
            } else if (ct === 2 && typeof f1 === 'string') cfHeader.rules.push({ type: 'expression', formula: f1, style });
            break;
          }
          case 0x01be: {
            const fl = r.u32();
            const strs = [r.xlString(), r.xlString(), r.xlString(), r.xlString()].map((s) => (s === '\0' ? '' : s));
            const n1 = r.u16(); r.u16();
            const g1 = rec.data.subarray(r.o, r.o + n1); r.skip(n1);
            const n2 = r.u16(); r.u16();
            const g2 = rec.data.subarray(r.o, r.o + n2); r.skip(n2);
            const cnt = r.u16();
            const ranges: CellRange[] = [];
            for (let k = 0; k < cnt; k++) { const r1 = r.u16(), r2 = r.u16(), c1 = r.u16(), c2 = r.u16(); ranges.push({ r1, r2, c1, c2 }); }
            if (!ranges.length) break;
            const base = ranges[0]!;
            const type = DV_TYPES[fl & 0xf] ?? 'any';
            const dv: DataValidation = { ranges, type };
            const op = DV_OPS[(fl >> 20) & 0xf];
            if (op && type !== 'any' && type !== 'list' && type !== 'custom') dv.operator = op;
            if (fl & 0x80 && g1.length > 2 && g1[0] === 0x17) {
              const s = decode(g1, new Uint8Array(0), base.r1, base.c1);
              if (typeof s === 'string') dv.formula1 = s.replace(/\u0000/g, ',');
            } else if (n1) { const s = decode(g1, new Uint8Array(0), base.r1, base.c1); if (typeof s === 'string') dv.formula1 = s; }
            if (n2) { const s = decode(g2, new Uint8Array(0), base.r1, base.c1); if (typeof s === 'string') dv.formula2 = s; }
            if (fl & 0x100) dv.allowBlank = true;
            if (type === 'list') dv.showDropdown = !(fl & 0x200);
            if (fl & 0x40000 && strs[2]) dv.prompt = { title: strs[0] || undefined, text: strs[2] };
            if (fl & 0x80000) dv.error = { style: (['stop', 'warning', 'information'] as const)[(fl >> 4) & 7] ?? 'stop', title: strs[1] || undefined, text: strs[3] ?? '' };
            sheet.validations.push(dv);
            break;
          }
          case 0x009e: {
            const col = r.u16(), g = r.u16();
            const d1 = rec.data.subarray(4, 14), d2 = rec.data.subarray(14, 24);
            r.o = 24;
            const fc: FilterColumn = { col };
            if (g & 0x10) fc.top10 = { top: !!(g & 0x20), percent: !!(g & 0x40), value: g >> 7 };
            else {
              const rules: { operator: CfOperator; value: string }[] = [];
              for (const d of [d1, d2]) {
                const vt = d[0]!, sign = d[1]!;
                if (vt === 0) continue;
                const dv = new DataView(d.buffer, d.byteOffset, 10);
                let value = '';
                if (vt === 2) value = String(rk(dv.getUint32(2, true)));
                else if (vt === 4) value = String(dv.getFloat64(2, true));
                else if (vt === 6) { const cch = d[6]!; value = r.chars(cch); }
                else if (vt === 8) value = d[3] ? String(ERR_BY_CODE[d[2]!] ?? '#N/A') : d[2] ? 'TRUE' : 'FALSE';
                else if (vt === 0x0c) { fc.blanks = true; continue; }
                else if (vt === 0x0e) { rules.push({ operator: 'notEqual', value: '' }); continue; }
                rules.push({ operator: FILTER_OPS[sign] ?? 'equal', value });
              }
              if (rules.length) fc.custom = { and: (g & 3) === 0, rules };
            }
            filterCols.push(fc);
            break;
          }
          case 0x00a1: {
            const paper = r.u16(), scale = r.u16(); r.u16(); const fitW = r.u16(), fitH = r.u16(), g = r.u16();
            r.u16(); r.u16();
            setup = { paper, scale, fitW, fitH, portrait: !!(g & 2), valid: !(g & 4), header: r.f64(), footer: r.f64() };
            break;
          }
          case 0x0026: margins.left = r.f64(); break;
          case 0x0027: margins.right = r.f64(); break;
          case 0x0028: margins.top = r.f64(); break;
          case 0x0029: margins.bottom = r.f64(); break;
          case 0x0014: if (rec.data.length) { const s = b5 ? r.shortString() : r.xlString(); if (s) ps.header = s; } break;
          case 0x0015: if (rec.data.length) { const s = b5 ? r.shortString() : r.xlString(); if (s) ps.footer = s; } break;
          case 0x0083: if (r.u16()) ps.centerH = true; break;
          case 0x0084: if (r.u16()) ps.centerV = true; break;
          case 0x002b: if (r.u16()) ps.gridlines = true; break;
          case 0x002a: if (r.u16()) ps.headings = true; break;
          case 0x0081: fitToPage = !!(r.u16() & 0x100); break;
          case 0x001b: case 0x001a: {
            const n = r.u16();
            const list: number[] = [];
            for (let k = 0; k < n; k++) { list.push(r.u16()); r.skip(4); }
            if (list.length) { if (rec.type === 0x001b) ps.rowBreaks = list; else ps.colBreaks = list; }
            break;
          }
          case 0x0862: {
            r.skip(12);
            const cb = r.u32();
            const icv = r.u32() & 0x7f;
            let c = color(icv);
            if (cb >= 0x28 && r.left >= 20) {
              r.u32(); r.u32();
              const xtype = r.u32(), val = r.u32(), tint = r.f64();
              if (xtype === 2) c = hex(val & 0xff, (val >> 8) & 0xff, (val >> 16) & 0xff);
              else if (xtype === 3 && DEFAULT_THEME[val]) c = applyTint(DEFAULT_THEME[val]!, tint);
              else if (xtype === 1) c = color(val) ?? c;
            }
            if (c && icv !== 0x7f) sheet.tabColor = c;
            break;
          }
          case 0x000d: if (r.u16() === 0) calc.manual = true; break;
          case 0x0011: if (r.u16() === 1) calc.iterate = true; break;
          case 0x000c: calc.maxIterations = r.u16(); break;
          case 0x0010: calc.maxChange = r.f64(); break;
          default: break;
        }
      } catch { /* a damaged record is skipped */ }
    }

    for (const e of expCells) {
      const def = shared.get(`${e.at.row},${e.at.col}`);
      if (!def) continue;
      const cell = sheet.cells.get(cellKey(e.row, e.col));
      if (!cell) continue;
      if (def.array) {
        if (e.row === def.range.r1 && e.col === def.range.c1) {
          const f = decode(def.rgce, def.rgcb, e.row, e.col);
          if (typeof f === 'string') { cell.f = f; cell.arrayRange = def.range; }
        }
      } else {
        const f = decode(def.rgce, def.rgcb, e.row, e.col);
        if (typeof f === 'string') cell.f = f;
      }
    }
    for (const n of notes) {
      const text = objText.get(n.id);
      if (text !== undefined) sheet.notes.set(cellKey(n.row, n.col), n.author ? { author: n.author, text } : { text });
    }
    if (drawing.length) {
      // drawing objects pair with OBJ records in order: charts, text boxes and shapes
      const anchored = readShapes(concat(...drawing)).filter((x) => x.anchor);
      anchored.forEach((sh, i) => {
        const obj = objs[i];
        if (!obj || !sh.anchor) return;
        const a = sh.anchor;
        const anchor = anchorFromCells(sheet, a.c1, a.dx1, a.r1, a.dy1, a.c2, a.dx2, a.r2, a.dy2);
        if (obj.chart) { sheet.charts.push({ id: `chart${sheet.charts.length + 1}`, anchor, ...obj.chart, ...(sh.alt ? { alt: sh.alt } : {}) }); return; }
        const kind = obj.ot === 6 ? 'textBox' : obj.ot === 3 ? 'ellipse' : obj.ot === 1 ? 'line' : obj.ot === 2 ? 'rect' : obj.ot === 0x1e ? SHAPE_KINDS[sh.spt ?? 1] : undefined;
        if (!kind) return;
        const shape: Shape = { id: `shape${(sheet.shapes?.length ?? 0) + 1}`, kind, anchor };
        const text = objText.get(obj.id);
        if (text) shape.text = text;
        if (sh.fill && !sh.noFill) shape.fill = sh.fill;
        if (sh.noFill) shape.fill = 'none';
        if (sh.line) shape.line = { color: sh.line, width: 0.75 };
        if (sh.alt) shape.alt = sh.alt;
        (sheet.shapes ??= []).push(shape);
      });
    }
    if (drawing.length && blips.length) {
      let k = 0;
      for (const sh of readShapes(concat(...drawing))) {
        if (!sh.pib || !sh.anchor) continue;
        const blip = blips[sh.pib - 1];
        if (!blip) continue;
        const a = sh.anchor;
        const pic: Sheet['pictures'][number] = { id: `xls-${sheetIdx}-${++k}`, data: blip.data, mime: blip.mime, anchor: anchorFromCells(sheet, a.c1, a.dx1, a.r1, a.dy1, a.c2, a.dx2, a.r2, a.dy2) };
        if (sh.alt) pic.alt = sh.alt;
        sheet.pictures.push(pic);
      }
    }
    if (frozen && (paneX || paneY)) { sheet.view.frozenCols = paneX; sheet.view.frozenRows = paneY; }
    if (protect) {
      const p: NonNullable<Sheet['protection']> = { sheet: true };
      if (password) p.password = password.toString(16).toUpperCase().padStart(4, '0');
      if (protOptions) p.options = protOptions;
      sheet.protection = p;
    }
    if (setup) {
      if (setup.valid) {
        ps.orientation = setup.portrait ? 'portrait' : 'landscape';
        if (setup.paper) ps.paper = setup.paper;
        if (setup.scale && setup.scale !== 100) ps.scale = setup.scale;
      }
      if (fitToPage) { ps.fitWidth = setup.fitW; ps.fitHeight = setup.fitH; }
      ps.margins = { left: margins.left ?? 0.75, right: margins.right ?? 0.75, top: margins.top ?? 1, bottom: margins.bottom ?? 1, header: setup.header, footer: setup.footer };
    }
    if (Object.keys(ps).length) sheet.pageSetup = { ...sheet.pageSetup, ...ps };
    // drop column entries that only repeat the most common width
    if (sheet.cols.size > 64) {
      const freq = new Map<number, number>();
      for (const c of sheet.cols.values()) if (!c.hidden && !c.s && !c.outline && c.width !== undefined) freq.set(c.width, (freq.get(c.width) ?? 0) + 1);
      const top = [...freq].sort((a, b) => b[1] - a[1])[0];
      if (top && top[1] > 64) {
        sheet.defaultColWidth = top[0];
        for (const [k, c] of [...sheet.cols]) if (!c.hidden && !c.s && !c.outline && c.width === top[0]) sheet.cols.delete(k);
      }
    }
    if (filterCols.length) (sheet as Sheet & { _filterCols?: FilterColumn[] })._filterCols = filterCols;
  }

  // defined names
  rawNames.forEach((n) => {
    let local = n.itab ? sheetOfBound.get(n.itab - 1) : undefined;
    if (n.itab && local === undefined) return;
    const f = decode(n.rgce, n.rgcb, 0, 0);
    if (typeof f !== 'string') return;
    const named = /^_xlnm\.(\w+)$/i.exec(n.name);
    if (named && n.builtin < 0) {
      const code = Object.entries(BUILTIN_NAMES).find(([, v]) => v.toLowerCase() === named[1]!.toLowerCase());
      if (code) { n.builtin = +code[0]; n.name = code[1]; }
    }
    if (n.builtin >= 0 && local === undefined) {
      const sh = /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec(f);
      const nm = sh ? (sh[1]?.replace(/''/g, "'") ?? sh[2]) : undefined;
      const k = nm !== undefined ? wb.sheets.findIndex((x) => x.name === nm) : -1;
      if (k >= 0) local = k;
    }
    if (n.builtin === 13 && local !== undefined) {
      const m = /!?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(f);
      const rg = m ? parseRef(m[1]!) : null;
      if (rg) filterRanges.set(local, rg);
      return;
    }
    if (n.builtin === 6 && local !== undefined) {
      const sh = wb.sheets[local]!;
      sh.pageSetup = { ...sh.pageSetup, printArea: f.split(',').map((p) => p.slice(p.lastIndexOf('!') + 1)).join(',') };
      return;
    }
    if (n.builtin === 7 && local !== undefined) {
      printTitles.set(local, f);
      return;
    }
    if (n.flags & 0x0e) return; // function, macro or VBA procedure names
    const d: Workbook['names'][number] = { name: n.builtin >= 0 ? `_xlnm.${n.name}` : n.name, formula: f };
    if (local !== undefined) d.sheet = local;
    if (n.flags & 1) d.hidden = true;
    wb.names.push(d);
  });
  for (const [k, f] of printTitles) {
    const sh = wb.sheets[k]!;
    for (const part of f.split(',')) {
      const p = part.slice(part.lastIndexOf('!') + 1);
      if (/^\$?\d+:\$?\d+$/.test(p)) sh.pageSetup = { ...sh.pageSetup, printTitleRows: p };
      else if (/^\$?[A-Z]+:\$?[A-Z]+$/.test(p)) sh.pageSetup = { ...sh.pageSetup, printTitleCols: p };
    }
  }
  wb.sheets.forEach((sh, k) => {
    const ext = sh as Sheet & { _filterCols?: FilterColumn[] };
    const cols = ext._filterCols ?? [];
    delete ext._filterCols;
    const range = filterRanges.get(k);
    if (!range) return;
    const af: AutoFilter = { range, columns: cols.filter((c) => c.custom || c.top10 || c.blanks).map((c) => ({ ...c, col: c.col })) };
    sh.autoFilter = af;
  });

  if (!wb.sheets.length) throw new Error('This workbook has no worksheets.');
  wb.activeSheet = Math.min(sheetOfBound.get(activeBound) ?? 0, wb.sheets.length - 1);
  if (wbProtect) wb.protection = { structure: true, ...(wbPassword ? { password: wbPassword.toString(16).toUpperCase().padStart(4, '0') } : {}) };
  if (Object.keys(calc).length) wb.calc = calc;
  const qs = cfbGet(streams, 'ClerdoGridQueries');
  if (qs) {
    const q = qs.slice();
    if (lastKey) new StreamCipher(lastKey).apply(q, 0, q.length, 0);
    const list = queriesFromXml(new TextDecoder().decode(q));
    if (list.length) wb.queries = list;
  }
  void quoteSheet;
  return wb;
  }
};

/** Decoder for a Windows code page number from the CODEPAGE record. */
const codepageDecoder = (cp: number): TextDecoder => {
  const name = cp === 1200 ? 'utf-16le' : cp === 0x8000 || cp === 10000 ? 'macintosh' : cp === 367 ? 'windows-1252' : cp === 932 ? 'shift_jis' : cp === 936 ? 'gbk' : cp === 949 ? 'euc-kr' : cp === 950 ? 'big5' : `windows-${cp}`;
  try { return new TextDecoder(name); } catch { return new TextDecoder('windows-1252'); }
};

const parseRef = (s: string): CellRange | null => {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(s);
  if (!m) return null;
  const col = (t: string) => [...t].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const c1 = col(m[1]!), r1 = +m[2]! - 1;
  const c2 = m[3] ? col(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
};

const SHAPE_KINDS: Record<number, ShapeKind> = {
  1: 'rect', 2: 'roundRect', 3: 'ellipse', 4: 'diamond', 5: 'triangle', 9: 'hexagon', 12: 'star5', 13: 'rightArrow', 20: 'line', 56: 'pentagon', 61: 'callout', 66: 'leftArrow', 67: 'downArrow', 68: 'upArrow', 202: 'textBox',
};
const CF_OPS: Record<number, CfOperator> = { 1: 'between', 2: 'notBetween', 3: 'equal', 4: 'notEqual', 5: 'greaterThan', 6: 'lessThan', 7: 'greaterThanOrEqual', 8: 'lessThanOrEqual' };
const DV_OPS: Record<number, CfOperator> = { 0: 'between', 1: 'notBetween', 2: 'equal', 3: 'notEqual', 4: 'greaterThan', 5: 'lessThan', 6: 'greaterThanOrEqual', 7: 'lessThanOrEqual' };
const DV_TYPES: Record<number, DataValidation['type']> = { 0: 'any', 1: 'whole', 2: 'decimal', 3: 'list', 4: 'date', 5: 'time', 6: 'textLength', 7: 'custom' };
const FILTER_OPS: Record<number, CfOperator> = { 1: 'lessThan', 2: 'equal', 3: 'lessThanOrEqual', 4: 'greaterThan', 5: 'notEqual', 6: 'greaterThanOrEqual' };

/** Differential format of a conditional-format rule (DXFN). */
const readDxfn = (r: Rd, color: (icv: number) => string | undefined, _fontAt: unknown, formats: Map<number, string>): CellStyle => {
  const f = r.u32();
  const f2 = r.u16();
  const st: CellStyle = {};
  if (f & (1 << 25)) {
    if (f2 & 1) { const cb = r.u16(); const start = r.o; const s = r.xlString(); if (s) st.numFmt = s; r.o = start + cb - 2; }
    else { r.u8(); const id = r.u8(); const code = formats.get(id) ?? builtinFormat(id); if (code) st.numFmt = code; }
  }
  if (f & (1 << 26)) {
    const start = r.o;
    const cch = r.u8();
    const name = cch ? r.chars(Math.min(cch, 31)) : '';
    r.o = start + 64;
    const height = r.i32(), ts = r.u32(), bls = r.u16(), sss = r.u16(), uls = r.u8();
    r.o = start + 80;
    const icv = r.i32();
    r.u32();
    const tsNinch = r.u32(), sssNinch = r.u32(), ulsNinch = r.u32(), blsNinch = r.u32();
    r.o = start + 118;
    const font: NonNullable<CellStyle['font']> = {};
    if (name) font.name = name;
    if (height > 0) font.size = height / 20;
    if (!(tsNinch & 2) && ts & 2) font.italic = true;
    if (!(tsNinch & 0x80) && ts & 0x80) font.strike = true;
    if ((!blsNinch || !(tsNinch & 2)) && bls >= 700) font.bold = true;
    if (!sssNinch && sss) font.vertAlign = sss === 1 ? 'superscript' : 'subscript';
    if (!ulsNinch && UNDERLINES[uls]) font.underline = UNDERLINES[uls];
    const c = icv >= 0 ? color(icv) : undefined;
    if (c) font.color = c;
    if (Object.keys(font).length) st.font = font;
  }
  if (f & (1 << 27)) r.skip(8);
  if (f & (1 << 28)) {
    const b1 = r.u32(), b2 = r.u32();
    const side = (dg: number, icv: number, ninch: boolean): BorderSide | undefined => {
      const s = BORDER_STYLES[dg];
      if (ninch || !s) return undefined;
      const c = color(icv);
      return c ? { style: s, color: c } : { style: s };
    };
    const border: NonNullable<CellStyle['border']> = {};
    const L = side(b1 & 0xf, (b1 >> 16) & 0x7f, !!(f & (1 << 10))), R = side((b1 >> 4) & 0xf, (b1 >> 23) & 0x7f, !!(f & (1 << 11)));
    const T = side((b1 >> 8) & 0xf, b2 & 0x7f, !!(f & (1 << 12))), B = side((b1 >> 12) & 0xf, (b2 >> 7) & 0x7f, !!(f & (1 << 13)));
    if (L) border.left = L;
    if (R) border.right = R;
    if (T) border.top = T;
    if (B) border.bottom = B;
    if (Object.keys(border).length) st.border = border;
  }
  if (f & (1 << 29)) {
    const p = r.u16(), c = r.u16();
    const fls = (p >> 10) & 0x3f;
    const fore = c & 0x7f, back = (c >> 7) & 0x7f;
    const flsNinch = !!(f & (1 << 16));
    const fill = flsNinch || fls === 1 ? color(flsNinch ? back : fore) ?? color(back) : color(back);
    if (fill && (flsNinch || fls >= 1)) st.fill = { color: fill };
  }
  if (f & (1 << 30)) r.skip(2);
  return st;
};

/** XFEXT: full colours for an XF. */
const readXfExt = (r: Rd, xfs: Xf[]): void => {
  r.skip(12); r.u16();
  const ixfe = r.u16();
  r.u16();
  const n = r.u16();
  const x = xfs[ixfe];
  if (!x) return;
  const ext: NonNullable<Xf['ext']> = {};
  for (let k = 0; k < n && r.left >= 4; k++) {
    const type = r.u16(), cb = r.u16();
    const start = r.o;
    if (cb >= 12 && [4, 5, 7, 8, 9, 10, 13].includes(type)) {
      const xt = r.u16(), tint = r.i16(), val = r.u32();
      let c: string | undefined;
      if (xt === 2) c = hex(val & 0xff, (val >> 8) & 0xff, (val >> 16) & 0xff);
      else if (xt === 3 && DEFAULT_THEME[val]) c = applyTint(DEFAULT_THEME[val]!, tint / 32767);
      const key = ({ 4: 'fore', 5: 'back', 7: 'top', 8: 'bottom', 9: 'left', 10: 'right', 13: 'text' } as const)[type as 4];
      if (c && key) ext[key] = c;
    }
    r.o = start + cb - 4;
  }
  if (Object.keys(ext).length) x.ext = ext;
};

const URL_MONIKER = 'E0C9EA79F9BACE118C8200AA004BA90B';
const guid = (b: Uint8Array): string => [...b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();

/** HLINK record: keeps web, mail and in-workbook links only. */
const readHlink = (d: Uint8Array, sheet: Sheet): void => {
  const r = new Rd(d);
  const r1 = r.u16(), r2 = r.u16(), c1 = r.u16(), c2 = r.u16();
  r.skip(16);
  r.u32();
  const fl = r.u32();
  const hstr = () => { const n = r.u32(); const s = r.utf16(n); return s.replace(/\0+$/, ''); };
  if (fl & 0x14) hstr();
  if (fl & 0x80) hstr();
  let url = '';
  if (fl & 1) {
    if (fl & 0x100) url = hstr();
    else {
      const id = guid(r.bytes(16));
      if (id === URL_MONIKER) { const len = r.u32(); url = r.utf16(Math.floor(len / 2)).replace(/\0[\s\S]*$/, ''); }
      else return;
    }
  }
  let loc = '';
  if (fl & 8) loc = hstr();
  const target = url ? (loc ? `${url}#${loc}` : url) : loc ? `#${loc}` : '';
  if (!target || !SAFE_LINK.test(target)) return;
  let n = 0;
  for (let i = r1; i <= r2 && n < 10000; i++) for (let j = c1; j <= c2 && n < 10000; j++, n++) sheet.hyperlinks.set(cellKey(i, j), target);
};

const BUILTIN_NAMES: Record<number, string> = {
  0: 'Consolidate_Area', 1: 'Auto_Open', 2: 'Auto_Close', 3: 'Extract', 4: 'Database', 5: 'Criteria', 6: 'Print_Area', 7: 'Print_Titles',
  8: 'Recorder', 9: 'Data_Form', 10: 'Auto_Activate', 11: 'Auto_Deactivate', 12: 'Sheet_Title', 13: '_FilterDatabase',
};

const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};

const rk = (v: number): number => {
  let n: number;
  if (v & 2) n = v >> 2;
  else {
    const b = new DataView(new ArrayBuffer(8));
    b.setUint32(4, v & 0xfffffffc, true);
    n = b.getFloat64(0, true);
  }
  return v & 1 ? n / 100 : n;
};
