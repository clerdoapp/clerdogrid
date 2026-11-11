// Binary workbook writer (.xlsb per the MS-XLSB open specification). Cells, styles, names and sheet
// settings are binary records; drawings and charts are the same XML parts an .xlsx package uses.
import { zipSync, type Zippable } from 'fflate';
import {
  addr, keyCol, keyRow,
  type BorderSide, type CellRange, type CellStyle, type CfRule, type ConditionalFormat, type Note, type RichRun, type Sheet, type Workbook,
} from '../model/workbook';
import { CellError, isError, type Scalar } from '../engine/types';
import type { Node } from '../engine/parser';
import { CODE_BY_ERR, encodeFormula, type EncodeCtx } from './biff8-formula';
import { BORDER_STYLES, H_ALIGNS, PATTERNS, V_ALIGNS } from './biff8-common';
import { builtinId } from './numfmt';
import { XML_HEAD, el, esc } from './xml';
import { DRAWING_HEAD, anchorXml, fileColWidth, pictureXml } from './xlsx-chart';
import { cellLookup, chartAnchorXml, type PartWriter } from './xlsx-chartex';
import { OBJECTS_CT, OBJECTS_PATH, OBJECTS_REL, objectsPart, shapeXml, withDescr } from './xlsx-objects';
import { addFuturePrefixes, quoteSheet } from './formula-xml';
import { ftabByName } from './biff8-ftab';

const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
/** Our own part for settings the binary records here don't carry (color scales, data bars, icon sets, tables, filter criteria). */
export const XLSB_EXTRA_PATH = 'xl/clerdogrid/binaryExtras.json';
export const XLSB_EXTRA_REL = 'https://schemas.clerdo.app/grid/2026/relationships/binaryExtras';
const XLSB_EXTRA_CT = 'application/vnd.clerdo.grid.binary-extras+json';
const MAX_ROWS = 1048576, MAX_COLS = 16384;
const SAFE_LINK = /^(https:|mailto:|#)/i;
const IMG_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/tiff': 'tiff', 'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' };

const enc = new TextEncoder();

class B {
  b: number[] = [];
  u8(v: number): this { this.b.push(v & 0xff); return this; }
  u16(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff); return this; }
  u32(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); return this; }
  f64(v: number): this { this.b.push(...new Uint8Array(new Float64Array([v]).buffer)); return this; }
  bytes(a: ArrayLike<number>): this { for (let i = 0; i < a.length; i++) this.b.push(a[i]!); return this; }
  /** XLWideString: 32-bit count of UTF-16 code units. */
  wide(s: string): this { this.u32(s.length); for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i)); return this; }
  /** XLNullableWideString: absent strings are stored as the count 0xFFFFFFFF. */
  nwide(s: string | undefined): this { return s === undefined ? this.u32(0xffffffff) : this.wide(s); }
  rfx(r: CellRange): this { return this.u32(r.r1).u32(r.r2).u32(r.c1).u32(r.c2); }
  formula(t: { rgce: Uint8Array; rgcb: Uint8Array } | null): this { return t ? this.u32(t.rgce.length).bytes(t.rgce).u32(t.rgcb.length).bytes(t.rgcb) : this.u32(0).u32(0); }
}

/** A part of records: variable-length type and size before each body. */
class Part {
  private out: number[] = [];
  rec(type: number, body: B | number[] = []): this {
    const data = body instanceof B ? body.b : body;
    let t = type;
    do { let x = t & 0x7f; t >>= 7; if (t) x |= 0x80; this.out.push(x); } while (t);
    let n = data.length;
    do { let x = n & 0x7f; n >>>= 7; if (n) x |= 0x80; this.out.push(x); } while (n);
    for (let i = 0; i < data.length; i++) this.out.push(data[i]!);
    return this;
  }
  bytes(): Uint8Array { return new Uint8Array(this.out); }
}

class Rels {
  private list: string[] = [];
  add(type: string, target: string, external = false): string {
    const id = `rId${this.list.length + 1}`;
    this.list.push(el('Relationship', { Id: id, Type: type.startsWith('http') ? type : RT + type, Target: target, TargetMode: external ? 'External' : undefined }));
    return id;
  }
  get size(): number { return this.list.length; }
  xml(): string { return `${XML_HEAD}<Relationships xmlns="${NS_PKG_REL}">${this.list.join('')}</Relationships>`; }
}

/** BrtColor: "#RRGGBB" as an RGB color, or automatic. */
const color = (b: B, c: string | undefined, auto: 'auto' | 'text' | 'bg' = 'auto'): B => {
  const m = c && /^#([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(c);
  if (m) {
    const v = parseInt(m[1]!, 16);
    return b.u8(1 | (2 << 1)).u8(0).u16(0).u8(v >> 16).u8((v >> 8) & 0xff).u8(v & 0xff).u8(0xff);
  }
  if (auto === 'text') return b.u8(1 | (3 << 1)).u8(1).u16(0).u32(0xff000000);
  if (auto === 'bg') return b.u8(1 | (1 << 1)).u8(64).u16(0).u32(0xff000000);
  return b.u8(0).u8(0).u16(0).u32(0);
};

const UNDERLINE_CODES: Record<string, number> = { single: 1, double: 2, singleAccounting: 0x21, doubleAccounting: 0x22 };
type Font = NonNullable<CellStyle['font']>;
const DEFAULT_FONT: Font = { name: 'Calibri', size: 11 };

const fontRec = (f: Font): B => {
  const b = new B().u16(Math.round((f.size ?? 11) * 20)).u16((f.italic ? 2 : 0) | (f.strike ? 8 : 0)).u16(f.bold ? 700 : 400)
    .u16(f.vertAlign === 'superscript' ? 1 : f.vertAlign === 'subscript' ? 2 : 0).u8(f.underline ? UNDERLINE_CODES[f.underline] ?? 1 : 0).u8(2).u8(0).u8(0);
  color(b, f.color, 'text');
  return b.u8(0).wide(f.name ?? 'Calibri');
};

const fillRec = (f: CellStyle['fill']): B => {
  const b = new B();
  const solid = f?.color ?? f?.gradient?.stops[0]?.color;
  if (f?.pattern && f.pattern !== 'solid' && f.pattern !== 'none') {
    b.u32(Math.max(0, PATTERNS.indexOf(f.pattern)));
    color(b, f.patternColor, 'text'); color(b, f.color, 'bg');
  } else if (solid) { b.u32(1); color(b, solid); color(b, undefined, 'bg'); }
  else { b.u32(0); color(b, undefined, 'text'); color(b, undefined, 'bg'); }
  return b.u32(0).f64(0).f64(0).f64(0).f64(0).f64(0).u32(0);
};

const side = (b: B, s: BorderSide | undefined): B => {
  b.u8(s ? Math.max(0, BORDER_STYLES.indexOf(s.style)) : 0).u8(0);
  return color(b, s?.color, s ? 'text' : 'auto');
};
const borderRec = (x: CellStyle['border']): B => {
  const b = new B().u8((x?.diagonalDown ? 1 : 0) | (x?.diagonalUp ? 2 : 0));
  side(b, x?.top); side(b, x?.bottom); side(b, x?.left); side(b, x?.right);
  return side(b, x?.diagonalUp ?? x?.diagonalDown);
};

/** Fonts, fills, borders, number formats and cell formats, shared between equal entries. */
class Styles {
  fonts: string[] = [];
  fills: string[] = [JSON.stringify(null), JSON.stringify({ pattern: 'gray125' })];
  borders: string[] = [JSON.stringify(null)];
  numFmts = new Map<string, number>();
  xfs: B[] = [];
  dxfs: B[] = [];
  private dxfIndex = new Map<string, number>();

  constructor(styles: readonly CellStyle[]) {
    this.fontIndex(styles[0]?.font);
    for (const st of styles.length ? styles : [{}]) this.xfs.push(this.xf(st));
  }
  private intern(list: string[], v: unknown): number {
    const k = JSON.stringify(v ?? null);
    const i = list.indexOf(k);
    if (i >= 0) return i;
    list.push(k);
    return list.length - 1;
  }
  fontIndex(f: Font | undefined): number { return this.intern(this.fonts, { ...DEFAULT_FONT, ...f }); }
  numFmtId(code: string | undefined): number {
    if (!code || code === 'General') return 0;
    const b = builtinId(code);
    if (b !== undefined) return b;
    let id = this.numFmts.get(code);
    if (id === undefined) { id = 164 + this.numFmts.size; this.numFmts.set(code, id); }
    return id;
  }
  private xf(st: CellStyle): B {
    const font = st.font ? this.fontIndex(st.font) : 0;
    const f = st.fill;
    const fill = f && (f.color || (f.pattern && f.pattern !== 'none') || f.gradient) ? this.intern(this.fills, f) : 0;
    const border = st.border && Object.keys(st.border).length ? this.intern(this.borders, st.border) : 0;
    const fmt = this.numFmtId(st.numFmt);
    const a = st.align ?? {};
    const rot = a.rotation === undefined ? 0 : a.rotation < 0 ? 90 - a.rotation : a.rotation;
    const h = a.h ? Math.max(0, H_ALIGNS.indexOf(a.h)) : 0;
    const v = a.v ? Math.max(0, V_ALIGNS.indexOf(a.v)) : 2;
    const locked = st.protection?.locked !== false, hidden = !!st.protection?.hidden;
    const flags = h | (v << 3) | (a.wrap ? 0x40 : 0) | (a.shrink ? 0x100 : 0) | (a.rtl ? 2 << 10 : 0) | (locked ? 0x1000 : 0) | (hidden ? 0x2000 : 0);
    const used = (fmt ? 1 : 0) | (font ? 2 : 0) | (st.align ? 4 : 0) | (border ? 8 : 0) | (fill ? 16 : 0) | (st.protection ? 32 : 0);
    return new B().u16(0).u16(fmt).u16(font).u16(fill).u16(border).u8(Math.min(255, rot)).u8(Math.min(250, a.indent ?? 0)).u16(flags).u8(used).u8(0);
  }
  /** BrtDXF: a differential format as a list of XFProp entries. */
  dxf(st: CellStyle): number {
    const props: B[] = [];
    const prop = (type: number, data: B) => props.push(new B().u16(type).u16(data.b.length + 4).bytes(data.b));
    const col = (c: string | undefined) => {
      const v = parseInt((c ?? '#000000').slice(1, 7), 16) || 0;
      return new B().u8(1 | (2 << 1)).u8(0).u16(0).u8(v >> 16).u8((v >> 8) & 0xff).u8(v & 0xff).u8(0xff);
    };
    const f = st.fill;
    if (f?.pattern && f.pattern !== 'solid' && f.pattern !== 'none') {
      prop(0, new B().u8(Math.max(0, PATTERNS.indexOf(f.pattern))));
      if (f.patternColor) prop(1, col(f.patternColor));
      if (f.color) prop(2, col(f.color));
    } else if (f?.color || f?.gradient) {
      prop(0, new B().u8(1));
      const c = f.color ?? f.gradient!.stops[0]!.color;
      prop(1, col(c)); prop(2, col(c));
    }
    const fo = st.font;
    if (fo?.color) prop(5, col(fo.color));
    const bd = st.border;
    const bside = (type: number, s: BorderSide | undefined) => { if (s) prop(type, col(s.color).u16(Math.max(0, BORDER_STYLES.indexOf(s.style)))); };
    bside(6, bd?.top); bside(7, bd?.bottom); bside(8, bd?.left); bside(9, bd?.right);
    const a = st.align;
    if (a?.h) prop(15, new B().u8(Math.max(0, H_ALIGNS.indexOf(a.h))));
    if (a?.v) prop(16, new B().u8(Math.max(0, V_ALIGNS.indexOf(a.v))));
    if (a?.wrap) prop(20, new B().u8(1));
    if (fo?.name) { const n = fo.name.slice(0, 32); const x = new B().u16(n.length); for (let i = 0; i < n.length; i++) x.u16(n.charCodeAt(i)); prop(24, x); }
    if (fo?.bold !== undefined) prop(25, new B().u16(fo.bold ? 700 : 400));
    if (fo?.underline !== undefined) prop(26, new B().u16(UNDERLINE_CODES[fo.underline] ?? 1));
    if (fo?.vertAlign) prop(27, new B().u16(fo.vertAlign === 'superscript' ? 1 : 2));
    if (fo?.italic !== undefined) prop(28, new B().u8(fo.italic ? 1 : 0));
    if (fo?.strike !== undefined) prop(29, new B().u8(fo.strike ? 1 : 0));
    if (fo?.size) prop(36, new B().u32(Math.round(fo.size * 20)));
    if (st.numFmt) prop(41, new B().u16(this.numFmtId(st.numFmt)));
    if (st.protection?.locked !== undefined) prop(43, new B().u8(st.protection.locked ? 1 : 0));
    if (st.protection?.hidden !== undefined) prop(44, new B().u8(st.protection.hidden ? 1 : 0));
    const b = new B().u16(0).u16(props.length);
    for (const p of props) b.bytes(p.b);
    const key = b.b.join(',');
    let i = this.dxfIndex.get(key);
    if (i === undefined) { i = this.dxfs.length; this.dxfs.push(b); this.dxfIndex.set(key, i); }
    return i;
  }
  part(): Uint8Array {
    const p = new Part().rec(278);
    if (this.numFmts.size) {
      p.rec(615, new B().u32(this.numFmts.size));
      for (const [code, id] of this.numFmts) p.rec(44, new B().u16(id).wide(code));
      p.rec(616);
    }
    p.rec(611, new B().u32(this.fonts.length));
    for (const f of this.fonts) p.rec(43, fontRec(JSON.parse(f) as Font));
    p.rec(612).rec(603, new B().u32(this.fills.length));
    this.fills.forEach((f, i) => {
      if (i === 1) { const b = new B().u32(17); color(b, undefined, 'text'); color(b, undefined, 'bg'); p.rec(45, b.u32(0).f64(0).f64(0).f64(0).f64(0).f64(0).u32(0)); }
      else p.rec(45, fillRec(JSON.parse(f) as CellStyle['fill']));
    });
    p.rec(604).rec(613, new B().u32(this.borders.length));
    for (const b of this.borders) p.rec(46, borderRec(JSON.parse(b) as CellStyle['border']));
    p.rec(614).rec(626, new B().u32(1)).rec(47, new B().u16(0xffff).u16(0).u16(0).u16(0).u16(0).u8(0).u8(0).u16(0x1010).u8(0).u8(0)).rec(627);
    p.rec(617, new B().u32(this.xfs.length));
    for (const x of this.xfs) p.rec(47, x);
    p.rec(618).rec(619, new B().u32(1)).rec(48, new B().u32(0).u16(1).u8(0).u8(0xff).wide('Normal')).rec(620);
    p.rec(505, new B().u32(this.dxfs.length));
    for (const d of this.dxfs) p.rec(507, d);
    p.rec(506).rec(508, new B().u32(0).wide('TableStyleMedium2').wide('PivotStyleLight16')).rec(509);
    return p.rec(279).bytes();
  }
}

/** What the .xlsb file can't hold; ClerdoGrid keeps some of it in its own part. */
export const xlsbLimits = (wb: Workbook): string[] => {
  const out: string[] = [];
  if (wb.sheets.some((s) => s.kind === 'chart')) out.push('Chart sheets are saved as worksheets with the chart on them.');
  if (wb.sheets.some((s) => s.pivots?.length || s.slicers?.length)) out.push('PivotTables and slicers are left out.');
  if (wb.sheets.some((s) => s.sparklines?.length)) out.push('Sparklines are left out.');
  if (wb.sheets.some((s) => s.tables.length || s.conditionalFormats.some((c) => c.rules.some((r) => r.type === 'colorScale' || r.type === 'dataBar' || r.type === 'iconSet')))) {
    out.push('Tables, color scales, data bars and icon sets are kept for ClerdoGrid only; other apps show the cells without them.');
  }
  if (wb.sheets.some((s) => s.pageSetup?.headerPictures && Object.keys(s.pageSetup.headerPictures).length)) out.push('Pictures in headers and footers are left out.');
  if (wb.scripts?.length) out.push('Scripts are left out.');
  return out;
};

const CF_OPS: Record<string, number> = { between: 1, notBetween: 2, equal: 3, notEqual: 4, greaterThan: 5, lessThan: 6, greaterThanOrEqual: 7, lessThanOrEqual: 8 };
const DV_OPS: Record<string, number> = { between: 0, notBetween: 1, equal: 2, notEqual: 3, greaterThan: 4, lessThan: 5, greaterThanOrEqual: 6, lessThanOrEqual: 7 };
const DV_TYPES: Record<string, number> = { any: 0, whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7 };
const TEXT_OPS: Record<string, number> = { containsText: 0, notContainsText: 1, beginsWith: 2, endsWith: 3 };
const PERIODS: Record<string, number> = { today: 15, tomorrow: 16, yesterday: 17, last7Days: 18, lastMonth: 19, nextMonth: 20, thisWeek: 21, nextWeek: 22, lastWeek: 23, thisMonth: 24 };
const PERIOD_OPS: Record<string, number> = { today: 0, yesterday: 1, last7Days: 2, thisWeek: 3, lastWeek: 4, lastMonth: 5, tomorrow: 6, nextWeek: 7, nextMonth: 8, thisMonth: 9 };

const textRuleFormula = (r: Extract<CfRule, { text: string }>, first: string): string => {
  const t = `"${r.text.replace(/"/g, '""')}"`;
  switch (r.type) {
    case 'containsText': return `NOT(ISERROR(SEARCH(${t},${first})))`;
    case 'notContainsText': return `ISERROR(SEARCH(${t},${first}))`;
    case 'beginsWith': return `LEFT(${first},LEN(${t}))=${t}`;
    default: return `RIGHT(${first},LEN(${t}))=${t}`;
  }
};
const PERIOD_FORMULA: Record<string, (a: string) => string> = {
  today: (a) => `FLOOR(${a},1)=TODAY()`,
  yesterday: (a) => `FLOOR(${a},1)=TODAY()-1`,
  tomorrow: (a) => `FLOOR(${a},1)=TODAY()+1`,
  last7Days: (a) => `AND(TODAY()-FLOOR(${a},1)<=6,FLOOR(${a},1)<=TODAY())`,
  thisWeek: (a) => `AND(TODAY()-ROUNDDOWN(${a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${a},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
  lastWeek: (a) => `AND(TODAY()-ROUNDDOWN(${a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${a},0)<(WEEKDAY(TODAY())+7))`,
  nextWeek: (a) => `AND(ROUNDDOWN(${a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${a},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
  thisMonth: (a) => `AND(MONTH(${a})=MONTH(TODAY()),YEAR(${a})=YEAR(TODAY()))`,
  lastMonth: (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0-1)),YEAR(${a})=YEAR(EDATE(TODAY(),0-1)))`,
  nextMonth: (a) => `AND(MONTH(${a})=MONTH(EDATE(TODAY(),0+1)),YEAR(${a})=YEAR(EDATE(TODAY(),0+1)))`,
};

interface Extras { cf?: Record<string, ConditionalFormat[]>; tables?: Record<string, Sheet['tables']>; filters?: Record<string, NonNullable<Sheet['autoFilter']>> }

/** Writes an .xlsb workbook. */
export const writeXlsb = (wb: Workbook): Uint8Array => {
  if (!wb.sheets.length) throw new Error('A workbook needs at least one sheet.');
  const files: Zippable = {};
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml'], ['bin', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main']]);
  const putText = (path: string, text: string, ct?: string) => { files[path] = enc.encode(text); if (ct) overrides.set('/' + path, ct); };
  const putBin = (path: string, data: Uint8Array, ct: string) => { files[path] = data; overrides.set('/' + path, ct); };
  const counters = { drawing: 0, chart: 0, image: 0, comments: 0 };
  const styles = new Styles(wb.styles);
  const sheets = wb.sheets;
  const sheetIdx = new Map(sheets.map((s, i) => [s.name.toUpperCase(), i]));
  const extras: Extras = {};

  // ---------- names and formula context ----------
  interface NameRec { name: string; formula?: string; sheet?: number; hidden?: boolean; func?: boolean; builtin?: boolean; comment?: string }
  const names: NameRec[] = [];
  for (const n of wb.names) names.push({ name: n.name, formula: n.formula, sheet: n.sheet, hidden: n.hidden, comment: n.comment, builtin: /^_xlnm\./i.test(n.name) });
  const absRef = (p: string) => p.replace(/\$?([A-Z]+)\$?(\d+)/g, '$$$1$$$2');
  sheets.forEach((s, i) => {
    const q = quoteSheet(s.name);
    if (s.pageSetup?.printArea && !names.some((n) => n.sheet === i && /^_xlnm\.Print_Area$/i.test(n.name))) {
      names.push({ name: '_xlnm.Print_Area', builtin: true, sheet: i, formula: s.pageSetup.printArea.split(',').map((p) => `${q}!${absRef(p.trim())}`).join(',') });
    }
    if (s.autoFilter) names.push({ name: '_xlnm._FilterDatabase', builtin: true, hidden: true, sheet: i, formula: `${q}!${absRef(addr(s.autoFilter.range.r1, s.autoFilter.range.c1))}:${absRef(addr(s.autoFilter.range.r2, s.autoFilter.range.c2))}` });
  });
  const nameIdx = (name: string, sheet?: number): number | undefined => {
    const u = name.toUpperCase();
    let k = names.findIndex((n) => !n.func && n.name.toUpperCase() === u && n.sheet === sheet);
    if (k < 0 && sheet === undefined) k = names.findIndex((n) => !n.func && n.name.toUpperCase() === u);
    return k < 0 ? undefined : k + 1;
  };
  const funcName = (fn: string): number => {
    const stored = ftabByName(fn) ? fn : addFuturePrefixes(`${fn}()`).slice(0, -2).replace(/^(?!_xlfn\.)/, '_xlfn.');
    const k = names.findIndex((n) => n.func && n.name === stored);
    if (k >= 0) return k + 1;
    names.push({ name: stored, func: true, hidden: true });
    return names.length;
  };
  const xti: [number, number][] = [];
  const ixti = (a: number, b: number): number => {
    const k = xti.findIndex(([x, y]) => x === a && y === b);
    if (k >= 0) return k;
    xti.push([a, b]);
    return xti.length - 1;
  };
  const tableRef = (s: Sheet | undefined) => (n: Extract<Node, { type: 'table' }>) => {
    for (const sh of sheets) {
      const t = sh.tables.find((x) => x.name.toUpperCase() === n.table.toUpperCase());
      if (!t) continue;
      const items = n.spec.items;
      let r1 = t.range.r1 + (t.headerRow ? 1 : 0), r2 = t.range.r2 - (t.totalsRow ? 1 : 0);
      if (items.includes('#All')) { r1 = t.range.r1; r2 = t.range.r2; }
      else if (items.includes('#Headers') && !items.includes('#Data')) { r1 = r2 = t.range.r1; }
      else if (items.includes('#Totals') && !items.includes('#Data')) { r1 = r2 = t.range.r2; }
      let c1 = t.range.c1, c2 = t.range.c2;
      if (n.spec.col1) {
        const a = t.columns.findIndex((c) => c.name.toUpperCase() === n.spec.col1!.toUpperCase());
        const b = n.spec.col2 ? t.columns.findIndex((c) => c.name.toUpperCase() === n.spec.col2!.toUpperCase()) : a;
        if (a < 0 || b < 0) return undefined;
        c1 = t.range.c1 + Math.min(a, b); c2 = t.range.c1 + Math.max(a, b);
      }
      return { r1, r2, c1, c2, r1Abs: true, r2Abs: true, c1Abs: true, c2Abs: true, kind: r1 === r2 && c1 === c2 ? 'cell' as const : 'area' as const, sheet: sh === s ? undefined : sh.name };
    }
    return undefined;
  };
  const ctxFor = (mode: EncodeCtx['mode'], sheet?: Sheet, row = 0, col = 0): EncodeCtx => ({
    xlsb: true, mode, row, col,
    sheetIndex: (n) => sheetIdx.get(n.toUpperCase()),
    ixti: (a, b) => ixti(a < 0 ? 0 : a, b < 0 ? 0 : b),
    nameIndex: nameIdx,
    addin: funcName,
    table: tableRef(sheet),
  });
  const compile = (f: string, ctx: EncodeCtx, array = false): { rgce: Uint8Array; rgcb: Uint8Array } | null => {
    try { return encodeFormula(f, ctx, array); } catch { return null; }
  };

  // ---------- shared strings ----------
  const sst: B[] = [];
  const sstIndex = new Map<string, number>();
  let sstTotal = 0;
  const stringIndex = (text: string, rich?: RichRun[]): number => {
    sstTotal++;
    const b = new B();
    if (rich?.length) {
      const runs: [number, number][] = [];
      let p = 0;
      for (const r of rich) { if (r.text) runs.push([p, styles.fontIndex(r.font)]); p += r.text.length; }
      b.u8(1).wide(text).u32(runs.length);
      for (const [ich, ifnt] of runs) b.u16(ich).u16(ifnt);
    } else b.u8(0).wide(text);
    const key = b.b.join(',');
    let i = sstIndex.get(key);
    if (i === undefined) { i = sst.length; sst.push(b); sstIndex.set(key, i); }
    return i;
  };

  // ---------- sheets ----------
  const wbRels = new Rels();
  const bundles: { rid: string; name: string; state: number }[] = [];
  sheets.forEach((s, si) => {
    const rels = new Rels();
    const p = new Part().rec(129);
    const ps = s.pageSetup ?? {};
    const fit = ps.fitWidth !== undefined || ps.fitHeight !== undefined;
    const wsFlags = 0x0409 | (s.outlinePr?.summaryBelow === false ? 0 : 0x40) | (s.outlinePr?.summaryRight === false ? 0 : 0x80) | (fit ? 0x100 : 0);
    const tab = new B().u16(wsFlags).u8(0);
    color(tab, s.tabColor);
    p.rec(147, tab.u32(0xffffffff).u32(0xffffffff).wide(''));

    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const k of s.cells.keys()) {
      const r = keyRow(k), c = keyCol(k);
      if (r < r1) r1 = r;
      if (r > r2) r2 = r;
      if (c < c1) c1 = c;
      if (c > c2) c2 = c;
    }
    p.rec(148, new B().rfx(r2 < 0 ? { r1: 0, r2: 0, c1: 0, c2: 0 } : { r1, r2, c1, c2 }));

    const v = s.view;
    const viewFlags = 0x10 | (v.showGrid ? 4 : 0) | (v.showHeaders ? 8 : 0) | (v.rtl ? 0x20 : 0) | (si === wb.activeSheet ? 0x40 : 0) | 0x100 | 0x200;
    const xlView = v.pageView === 'pageLayout' ? 2 : v.pageView === 'pageBreakPreview' ? 1 : 0;
    p.rec(133).rec(137, new B().u16(viewFlags).u32(xlView).u32(0).u32(0).u32(64).u16(v.zoom).u16(xlView === 0 ? v.zoom : 0).u16(xlView === 1 ? v.zoom : 0).u16(xlView === 2 ? v.zoom : 0).u32(0));
    let pnn = 3;
    if (v.frozenRows || v.frozenCols) {
      pnn = v.frozenRows && v.frozenCols ? 0 : v.frozenRows ? 2 : 1;
      p.rec(151, new B().f64(v.frozenCols).f64(v.frozenRows).u32(v.frozenRows).u32(v.frozenCols).u32(pnn).u8(1 | 2));
    }
    const sel = v.selection ?? { r1: 0, r2: 0, c1: 0, c2: 0 };
    p.rec(152, new B().u32(pnn).u32(sel.r1).u32(sel.c1).u32(0).u32(1).rfx(sel));
    p.rec(138).rec(134);

    const outRow = Math.max(0, ...[...s.rows.values()].map((r) => r.outline ?? 0));
    const outCol = Math.max(0, ...[...s.cols.values()].map((c) => c.outline ?? 0));
    const dx = Math.abs(s.defaultColWidth - 8.43) > 0.001 ? Math.round(fileColWidth(s.defaultColWidth) * 256) : 0xffffffff;
    p.rec(485, new B().u32(dx).u16(8).u16(Math.round(s.defaultRowHeight * 20)).u16(s.defaultRowHeight !== 15 ? 1 : 0).u8(outRow).u8(outCol));

    if (s.cols.size) {
      p.rec(390);
      const cols = [...s.cols].filter(([c]) => c < MAX_COLS).sort((a, b) => a[0] - b[0]);
      let i = 0;
      while (i < cols.length) {
        const [start, info] = cols[i]!;
        const key = JSON.stringify(info);
        let end = start;
        while (i + 1 < cols.length && cols[i + 1]![0] === end + 1 && JSON.stringify(cols[i + 1]![1]) === key) { i++; end++; }
        const w = Math.round(fileColWidth(info.width ?? s.defaultColWidth) * 256);
        const fl = (info.hidden ? 1 : 0) | (info.width !== undefined ? 2 : 0) | ((info.outline ?? 0) & 7) << 8 | (info.collapsed ? 0x1000 : 0);
        p.rec(60, new B().u32(start).u32(end).u32(w).u32(info.s && info.s < wb.styles.length ? info.s : 0).u16(fl));
        i++;
      }
      p.rec(391);
    }

    // cells
    p.rec(145);
    const byRow = new Map<number, number[]>();
    for (const k of s.cells.keys()) {
      const r = keyRow(k);
      if (r >= MAX_ROWS) continue;
      let arr = byRow.get(r);
      if (!arr) byRow.set(r, (arr = []));
      arr.push(k);
    }
    for (const [r, info] of s.rows) if (r < MAX_ROWS && !byRow.has(r) && (info.height !== undefined || info.hidden || info.s || info.outline)) byRow.set(r, []);
    for (const r of [...byRow.keys()].sort((a, b) => a - b)) {
      const info = s.rows.get(r);
      const keys = byRow.get(r)!.sort((a, b) => a - b);
      const fl = ((info?.outline ?? 0) & 7) << 8 | (info?.collapsed ? 0x800 : 0) | (info?.hidden ? 0x1000 : 0) | (info?.height !== undefined ? 0x2000 : 0) | (info?.s ? 0x4000 : 0);
      const hdr = new B().u32(r).u32(info?.s && info.s < wb.styles.length ? info.s : 0).u16(Math.round((info?.height ?? s.defaultRowHeight) * 20)).u16(fl).u8(0);
      if (keys.length) hdr.u32(1).u32(keyCol(keys[0]!)).u32(keyCol(keys[keys.length - 1]!)); else hdr.u32(0);
      p.rec(0, hdr);
      for (const k of keys) {
        const c = keyCol(k);
        const cell = s.cells.get(k)!;
        const st = cell.s && cell.s < wb.styles.length ? cell.s : 0;
        const head = () => new B().u32(c).u32(st);
        const val: Scalar = cell.v;
        if (cell.f) {
          const array = !!(cell.arrayRange || cell.dyn);
          const range = cell.arrayRange ?? { r1: r, r2: r, c1: c, c2: c };
          const t = compile(cell.f, ctxFor('cell', s, r, c), array);
          if (t) {
            const fbody = (b: B) => (array ? b.u16(0).u32(5).u8(0x01).u32(r).u32(0) : b.u16(0).formula(t));
            if (typeof val === 'number') p.rec(9, fbody(head().f64(Number.isFinite(val) ? val : 0)));
            else if (typeof val === 'boolean') p.rec(10, fbody(head().u8(val ? 1 : 0)));
            else if (isError(val)) p.rec(11, fbody(head().u8(CODE_BY_ERR[val.code] ?? 0x2a)));
            else p.rec(8, fbody(head().wide(typeof val === 'string' ? val : '')));
            if (array) p.rec(426, new B().rfx(range).u8(0).formula(t));
            continue;
          }
        }
        if (val === null || val === undefined) { if (st) p.rec(1, head()); }
        else if (typeof val === 'number') p.rec(5, head().f64(Number.isFinite(val) ? val : 0));
        else if (typeof val === 'boolean') p.rec(4, head().u8(val ? 1 : 0));
        else if (val instanceof CellError || isError(val)) p.rec(3, head().u8(CODE_BY_ERR[(val as CellError).code] ?? 0x2a));
        else p.rec(7, head().u32(stringIndex(String(val), cell.rich)));
      }
    }
    p.rec(146);

    if (s.protection?.sheet) {
      const keys = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
      const o = s.protection.options ?? {};
      const b = new B().u16(s.protection.password ? parseInt(s.protection.password, 16) || 0 : 0).u32(1);
      // the binary flags say what is allowed, the opposite of the option names
      const offByDefault = new Set(['objects', 'scenarios', 'selectLockedCells', 'selectUnlockedCells']);
      for (const k of keys) b.u32((o[k] ?? !offByDefault.has(k)) ? 0 : 1);
      p.rec(535, b);
    }
    if (s.autoFilter) {
      p.rec(161, new B().rfx(s.autoFilter.range)).rec(162);
      if (s.autoFilter.columns.length || s.autoFilter.sort?.length) (extras.filters ??= {})[si] = s.autoFilter;
    }
    if (s.merges.length) {
      p.rec(177, new B().u32(s.merges.length));
      for (const m of s.merges) p.rec(176, new B().rfx(m));
      p.rec(178);
    }

    // conditional formats: rules with formulas and a format are records; scales, bars and icons go to our part
    let prio = 1;
    const kept: ConditionalFormat[] = [];
    for (const cf of s.conditionalFormats) {
      const ranges = cf.ranges.filter((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS);
      if (!ranges.length) continue;
      const base = ranges[0]!;
      const first = addr(base.r1, base.c1);
      const recs: B[] = [];
      const own: CfRule[] = [];
      for (const rule of cf.rules) {
        if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') { own.push(rule); prio++; continue; }
        const ctx = ctxFor('rel', s, base.r1, base.c1);
        let type = 2, sub = 1, param = 0, flags = rule.stop ? 2 : 0;
        let f1: string | undefined, f2: string | undefined, text: string | undefined;
        switch (rule.type) {
          case 'cellIs': type = 1; sub = 0; param = CF_OPS[rule.operator] ?? 3; f1 = rule.formulas[0]; f2 = rule.formulas[1]; break;
          case 'expression': f1 = rule.formula; break;
          case 'top10': type = 5; sub = 5; param = rule.rank; flags |= (rule.bottom ? 8 : 0) | (rule.percent ? 16 : 0); break;
          case 'aboveAverage': sub = rule.below ? (rule.equal ? 30 : 26) : (rule.equal ? 29 : 25); param = rule.equal ? 0 : Math.min(3, rule.stdDev ?? 0); flags |= rule.below ? 0 : 4; break;
          case 'duplicateValues': sub = 27; break;
          case 'uniqueValues': sub = 7; break;
          case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith':
            sub = 8; param = TEXT_OPS[rule.type]!; text = rule.text; f1 = textRuleFormula(rule, first); break;
          case 'containsBlanks': sub = 9; f1 = `LEN(TRIM(${first}))=0`; break;
          case 'notContainsBlanks': sub = 10; f1 = `LEN(TRIM(${first}))>0`; break;
          case 'containsErrors': sub = 11; f1 = `ISERROR(${first})`; break;
          case 'notContainsErrors': sub = 12; f1 = `NOT(ISERROR(${first}))`; break;
          case 'timePeriod': sub = PERIODS[rule.period] ?? 15; param = PERIOD_OPS[rule.period] ?? 0; f1 = PERIOD_FORMULA[rule.period]!(first); break;
        }
        const t1 = f1 !== undefined ? compile(f1, ctx) : null;
        const t2 = f2 !== undefined ? compile(f2, ctx) : null;
        if (f1 !== undefined && !t1) continue;
        const size = (t: typeof t1) => (t ? 8 + t.rgce.length + t.rgcb.length : 0);
        const b = new B().u32(type).u32(sub).u32('style' in rule ? styles.dxf(rule.style) : 0xffffffff).u32(prio++).u32(param).u32(0).u32(0)
          .u16(flags).u32(size(t1)).u32(size(t2)).u32(0).nwide(text);
        if (t1) b.formula(t1);
        if (t2) b.formula(t2);
        recs.push(b);
      }
      if (own.length) kept.push({ ranges, rules: own });
      if (!recs.length) continue;
      const head = new B().u32(recs.length).u32(0).u32(ranges.length);
      for (const g of ranges) head.rfx({ ...g, r2: Math.min(g.r2, MAX_ROWS - 1), c2: Math.min(g.c2, MAX_COLS - 1) });
      p.rec(461, head);
      for (const r of recs) p.rec(463, r).rec(464);
      p.rec(462);
    }
    if (kept.length) (extras.cf ??= {})[si] = kept;

    const dvs = s.validations.filter((d) => d.ranges.some((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS));
    if (dvs.length) {
      p.rec(573, new B().u16(0).u32(0).u32(0).u32(0xffffffff).u32(dvs.length));
      for (const d of dvs) {
        const ranges = d.ranges.filter((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS);
        const base = ranges[0]!;
        // plain references (relative flags set) measured from the first cell, as other apps read them
        const ctx = ctxFor('cell', s, base.r1, base.c1);
        let t1: { rgce: Uint8Array; rgcb: Uint8Array } | null = null;
        let lookup = false;
        if (d.type === 'list' && d.formula1 && /^".*"$/s.test(d.formula1)) {
          const items = d.formula1.slice(1, -1).replace(/""/g, '"').split(',').join('\0').slice(0, 255);
          const x = new B().u8(0x17).u16(items.length);
          for (let i = 0; i < items.length; i++) x.u16(items.charCodeAt(i));
          t1 = { rgce: new Uint8Array(x.b), rgcb: new Uint8Array(0) };
          lookup = true;
        } else if (d.formula1) t1 = compile(d.formula1, ctx);
        const t2 = d.formula2 ? compile(d.formula2, ctx) : null;
        const err = d.error ? Math.max(0, ['stop', 'warning', 'information'].indexOf(d.error.style)) : 0;
        const fl = (DV_TYPES[d.type] ?? 0) | (err << 4) | (lookup ? 0x80 : 0) | (d.allowBlank ? 0x100 : 0) | (d.type === 'list' && d.showDropdown === false ? 0x200 : 0)
          | (d.prompt ? 0x40000 : 0) | (d.error ? 0x80000 : 0) | ((DV_OPS[d.operator ?? 'between'] ?? 0) << 20);
        const b = new B().u32(fl >>> 0).u32(ranges.length);
        for (const g of ranges) b.rfx({ ...g, r2: Math.min(g.r2, MAX_ROWS - 1), c2: Math.min(g.c2, MAX_COLS - 1) });
        b.wide(d.error?.title ?? '').wide(d.error?.text ?? '').wide(d.prompt?.title ?? '').wide(d.prompt?.text ?? '');
        b.formula(t1).formula(t2);
        p.rec(64, b);
      }
      p.rec(574);
    }

    for (const [k, target] of [...s.hyperlinks].sort((a, b) => a[0] - b[0])) {
      if (!SAFE_LINK.test(target)) continue;
      const r = keyRow(k), c = keyCol(k);
      const b = new B().rfx({ r1: r, r2: r, c1: c, c2: c });
      if (target.startsWith('#')) b.nwide(undefined).wide(target.slice(1));
      else b.wide(rels.add('hyperlink', target, true)).wide('');
      p.rec(494, b.wide('').wide(''));
    }

    p.rec(477, new B().u16((ps.centerH ? 1 : 0) | (ps.centerV ? 2 : 0) | (ps.headings ? 4 : 0) | (ps.gridlines ? 8 : 0)));
    const m = ps.margins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
    p.rec(476, new B().f64(m.left).f64(m.right).f64(m.top).f64(m.bottom).f64(m.header).f64(m.footer));
    p.rec(478, new B().u32(ps.paper ?? 9).u32(ps.scale ?? 100).u32(ps.horizontalDpi ?? 600).u32(ps.verticalDpi ?? 600).u32(ps.copies ?? 1)
      .u32(ps.firstPageNumber ?? 1).u32(fit ? ps.fitWidth ?? 1 : 1).u32(fit ? ps.fitHeight ?? 1 : 1)
      .u16((ps.order === 'overThenDown' ? 1 : 0) | (ps.orientation === 'landscape' ? 2 : 0) | (ps.orientation ? 0 : 0x40) | (ps.firstPageNumber !== undefined ? 0x80 : 0)).nwide(undefined));
    if (ps.header || ps.footer || ps.evenHeader || ps.firstHeader) {
      p.rec(479, new B().u16((ps.differentOddEven ? 1 : 0) | (ps.differentFirst ? 2 : 0) | (ps.scaleWithDoc === false ? 0 : 4) | (ps.alignWithMargins === false ? 0 : 8))
        .nwide(ps.header).nwide(ps.footer).nwide(ps.evenHeader).nwide(ps.evenFooter).nwide(ps.firstHeader).nwide(ps.firstFooter)).rec(480);
    }

    // drawings: charts, pictures and shapes in one XML drawing part
    if (s.charts.length || s.pictures.length || s.shapes?.length) {
      const dn = ++counters.drawing;
      const dRels = new Rels();
      const w: PartWriter = {
        put: (path, text, ct) => putText(path, text, ct),
        rel: (type, target) => dRels.add(type, target),
        nextChart: () => ++counters.chart,
        look: cellLookup(wb),
      };
      let body = '';
      let shapeId = 1;
      for (const c of s.charts) body += chartAnchorXml(c, s, ++shapeId, w, withDescr);
      for (const pic of s.pictures) {
        const ext = IMG_EXT[pic.mime] ?? 'png';
        const n = ++counters.image;
        files[`xl/media/image${n}.${ext}`] = pic.data;
        defaults.set(ext, pic.mime);
        const rid = dRels.add('image', `../media/image${n}.${ext}`);
        body += anchorXml(pic.anchor, s, withDescr(pictureXml(++shapeId, `Picture ${n}`, rid, pic.anchor), pic.alt));
      }
      for (const sh of s.shapes ?? []) body += anchorXml(sh.anchor, s, shapeXml(++shapeId, sh));
      putText(`xl/drawings/drawing${dn}.xml`, `${XML_HEAD}${DRAWING_HEAD}${body}</xdr:wsDr>`, CT + 'drawing+xml');
      putText(`xl/drawings/_rels/drawing${dn}.xml.rels`, dRels.xml());
      p.rec(550, new B().wide(rels.add('drawing', `../drawings/drawing${dn}.xml`)));
    }
    if (s.notes.size) {
      const n = ++counters.comments;
      putBin(`xl/comments${n}.bin`, commentsPart(s.notes), 'application/vnd.ms-excel.comments');
      putText(`xl/drawings/vmlDrawing${n}.vml`, notesVml(s.notes, n));
      defaults.set('vml', CT + 'vmlDrawing');
      rels.add('comments', `../comments${n}.bin`);
      p.rec(551, new B().wide(rels.add('vmlDrawing', `../drawings/vmlDrawing${n}.vml`)));
    }
    if (s.tables.length) (extras.tables ??= {})[si] = s.tables;
    p.rec(130);
    putBin(`xl/worksheets/sheet${si + 1}.bin`, p.bytes(), 'application/vnd.ms-excel.worksheet');
    if (rels.size) putText(`xl/worksheets/_rels/sheet${si + 1}.bin.rels`, rels.xml());
    bundles.push({ rid: wbRels.add('worksheet', `worksheets/sheet${si + 1}.bin`), name: s.name, state: s.state === 'hidden' ? 1 : s.state === 'veryHidden' ? 2 : 0 });
  });

  // ---------- workbook ----------
  // name formulas are compiled after the sheets so names added for functions come first in the list
  const nameRecs: B[] = [];
  for (let i = 0; i < names.length; i++) {
    const n = names[i]!;
    const t = n.formula !== undefined ? compile(n.formula, ctxFor('name')) : null;
    const flags = (n.hidden ? 1 : 0) | (n.func ? 0x0a : 0) | (n.builtin ? 0x20 : 0) | (n.func && n.name.startsWith('_xlfn.') ? 1 << 17 : 0);
    const body = n.func ? { rgce: new Uint8Array([0x1c, 0x1d]), rgcb: new Uint8Array(0) } : t;
    const rec = new B().u32(flags).u8(0).u32(n.sheet ?? 0xffffffff).wide(n.name).formula(body).nwide(n.comment);
    if (n.func) rec.nwide(undefined).nwide(undefined).nwide(undefined).nwide(undefined);
    nameRecs.push(rec);
  }
  const p = new Part().rec(131);
  p.rec(153, new B().u32(wb.date1904 ? 1 : 0).u32(166925).wide(''));
  if (wb.protection?.structure || wb.protection?.windows) {
    p.rec(534, new B().u16(wb.protection.password ? parseInt(wb.protection.password, 16) || 0 : 0).u16(0).u16((wb.protection.structure ? 1 : 0) | (wb.protection.windows ? 2 : 0)));
  }
  const active = Math.max(0, Math.min(wb.activeSheet, sheets.length - 1));
  p.rec(135).rec(158, new B().u32(0).u32(0).u32(28800).u32(12300).u32(600).u32(0).u32(active).u8(0x78)).rec(136);
  p.rec(143);
  bundles.forEach((b, i) => p.rec(156, new B().u32(b.state).u32(i + 1).wide(b.rid).wide(b.name)));
  p.rec(144);
  if (xti.length) {
    p.rec(353).rec(357).rec(362, (() => { const b = new B().u32(xti.length); for (const [a, c] of xti) b.u32(0).u32(a).u32(c); return b; })()).rec(354);
  }
  for (const n of nameRecs) p.rec(39, n);
  const calc = wb.calc ?? {};
  p.rec(157, new B().u32(191029).u32(calc.manual ? 0 : calc.autoNoTable ? 2 : 1).u32(calc.maxIterations ?? 100).f64(calc.maxChange ?? 0.001).u32(0xffffffff)
    .u16(1 | (calc.r1c1 ? 0 : 2) | (calc.iterate ? 4 : 0) | (calc.precisionAsDisplayed ? 0 : 8)));
  p.rec(132);
  putBin('xl/workbook.bin', p.bytes(), 'application/vnd.ms-excel.sheet.binary.macroEnabled.main');

  putBin('xl/styles.bin', styles.part(), 'application/vnd.ms-excel.styles');
  wbRels.add('styles', 'styles.bin');
  if (sst.length) {
    const sp = new Part().rec(159, new B().u32(sstTotal).u32(sst.length));
    for (const b of sst) sp.rec(19, b);
    putBin('xl/sharedStrings.bin', sp.rec(160).bytes(), 'application/vnd.ms-excel.sharedStrings');
    wbRels.add('sharedStrings', 'sharedStrings.bin');
  }
  putText('xl/theme/theme1.xml', THEME, CT + 'theme+xml');
  wbRels.add('theme', 'theme/theme1.xml');
  const objects = objectsPart(wb);
  if (objects) { putText(OBJECTS_PATH, objects, OBJECTS_CT); wbRels.add(OBJECTS_REL, OBJECTS_PATH.replace(/^xl\//, '')); }
  if (extras.cf || extras.tables || extras.filters) {
    putText(XLSB_EXTRA_PATH, JSON.stringify({ app: 'ClerdoGrid', version: 1, ...extras }), XLSB_EXTRA_CT);
    wbRels.add(XLSB_EXTRA_REL, XLSB_EXTRA_PATH.replace(/^xl\//, ''));
  }
  // kept custom XML items (such as saved data queries)
  for (const [path, data] of wb.extraParts ?? []) {
    if (!/^customXml\/(item\d+\.xml|itemProps\d+\.xml|_rels\/item\d+\.xml\.rels)$/i.test(path) || files[path]) continue;
    files[path] = data;
    if (/itemProps/i.test(path)) overrides.set('/' + path, CT + 'customXmlProperties+xml');
    else if (/^customXml\/item\d+\.xml$/i.test(path)) wbRels.add('customXml', `../${path}`);
  }
  putText('xl/_rels/workbook.bin.rels', wbRels.xml());

  const rootRels = new Rels();
  rootRels.add('officeDocument', 'xl/workbook.bin');
  rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
  rootRels.add('extended-properties', 'docProps/app.xml');
  putText('_rels/.rels', rootRels.xml());
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const pr = wb.props ?? {};
  putText('docProps/core.xml', `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + (pr.title ? `<dc:title>${esc(pr.title)}</dc:title>` : '') + (pr.creator ? `<dc:creator>${esc(pr.creator)}</dc:creator>` : '')
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${esc(pr.created ?? now)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${esc(pr.modified ?? now)}</dcterms:modified></cp:coreProperties>`,
  'application/vnd.openxmlformats-package.core-properties+xml');
  putText('docProps/app.xml', `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ClerdoGrid</Application></Properties>`, CT + 'extended-properties+xml');

  const ct = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + [...defaults].map(([e, t]) => el('Default', { Extension: e, ContentType: t })).join('')
    + [...overrides].map(([pn, t]) => el('Override', { PartName: pn, ContentType: t })).join('')
    + '</Types>';
  return zipSync({ '[Content_Types].xml': enc.encode(ct), ...files }, { level: 6 });
};

const commentsPart = (notes: Map<number, Note>): Uint8Array => {
  const authors: string[] = [];
  const list = [...notes].sort((a, b) => a[0] - b[0]).map(([k, n]) => {
    const a = n.author ?? '';
    let i = authors.indexOf(a);
    if (i < 0) { i = authors.length; authors.push(a); }
    return { k, i, text: n.text };
  });
  const p = new Part().rec(628).rec(630);
  for (const a of authors) p.rec(632, new B().wide(a));
  p.rec(631).rec(633);
  for (const c of list) {
    const r = keyRow(c.k), col = keyCol(c.k);
    p.rec(635, new B().u32(c.i).rfx({ r1: r, r2: r, c1: col, c2: col }).bytes(new Uint8Array(16)));
    p.rec(637, new B().u8(0).wide(c.text.replace(/\n/g, '\r\n')));
    p.rec(636);
  }
  return p.rec(634).rec(629).bytes();
};

/** Note boxes for apps that draw notes from the legacy drawing part. */
const notesVml = (notes: Map<number, Note>, block: number): string => {
  const shapes = [...notes.keys()].sort((a, b) => a - b).map((k, i) => {
    const r = keyRow(k), c = keyCol(k);
    return `<v:shape id="_x0000_s${block * 1024 + 1 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`
      + '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
      + `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 15, ${r + 3}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  });
  return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
    + `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${block}"/></o:shapelayout>`
    + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
    + shapes.join('') + '</xml>';
};

const THEME_COLORS: [string, string][] = [['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'], ['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'],
  ['dk2', '<a:srgbClr val="44546A"/>'], ['lt2', '<a:srgbClr val="E7E6E6"/>'], ['accent1', '<a:srgbClr val="4472C4"/>'], ['accent2', '<a:srgbClr val="ED7D31"/>'],
  ['accent3', '<a:srgbClr val="A5A5A5"/>'], ['accent4', '<a:srgbClr val="FFC000"/>'], ['accent5', '<a:srgbClr val="5B9BD5"/>'], ['accent6', '<a:srgbClr val="70AD47"/>'],
  ['hlink', '<a:srgbClr val="0563C1"/>'], ['folHlink', '<a:srgbClr val="954F72"/>']];
const fontScheme = (latin: string): string => `<a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/>`;
const fillStyle = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
const THEME = `${XML_HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Default"><a:themeElements>`
  + `<a:clrScheme name="Default">${THEME_COLORS.map(([n, c]) => `<a:${n}>${c}</a:${n}>`).join('')}</a:clrScheme>`
  + `<a:fontScheme name="Default"><a:majorFont>${fontScheme('Calibri Light')}</a:majorFont><a:minorFont>${fontScheme('Calibri')}</a:minorFont></a:fontScheme>`
  + `<a:fmtScheme name="Default"><a:fillStyleLst>${fillStyle}${fillStyle}${fillStyle}</a:fillStyleLst>`
  + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
  + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
  + `<a:bgFillStyleLst>${fillStyle}${fillStyle}${fillStyle}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
