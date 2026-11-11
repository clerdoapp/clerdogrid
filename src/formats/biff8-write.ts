// Binary workbook writer (.xls, BIFF8 per the MS-XLS open specification) in a compound file.
// Older format limits apply: 65,536 rows, 256 columns, 56 custom colours, 4,000 cell styles.
import {
  DEFAULT_FONT, keyCol, keyRow, type Cell, type CellRange, type CellStyle, type CfRule, type FilterColumn, type Note, type Picture, type RichRun, type Sheet, type Workbook,
} from '../model/workbook';
import { CellError, type Scalar } from '../engine/types';
import type { Node } from '../engine/parser';
import { writeCfb } from './cfb';
import { CODE_BY_ERR, encodeFormula, type EncodeCtx } from './biff8-formula';
import { ftabByName } from './biff8-ftab';
import { addFuturePrefixes } from './formula-xml';
import { BORDER_STYLES, H_ALIGNS, PATTERNS, V_ALIGNS, cellsFromAnchor } from './biff8-common';
import { INDEXED_COLORS, builtinId } from './numfmt';
import { fileColWidth } from './xlsx-chart';
import { xlsChartRecords } from './biff8-chart';

const OLD_CHART_TYPES = /^(column|stackedColumn|percentColumn|bar|stackedBar|percentBar|line|stackedLine|percentLine|lineMarkers|stackedLineMarkers|percentLineMarkers|area|stackedArea|percentArea|pie|doughnut|scatter|scatterLinesMarkers|scatterLines|bubble|radar|radarMarkers|radarFilled|combo|comboSecondary|comboAreaColumn)$/;
import { StreamCipher, checkPassword, makeRc4Params, DEFAULT_PASSWORD } from './biff8-crypto';
import { md5 } from './biff8-crypto';
import { zlibSync } from 'fflate';
import { queriesToXml } from './queries-xml';

export const QUERIES_STREAM = 'ClerdoGridQueries';

const MAX_ROWS = 65536, MAX_COLS = 256, MAX_REC = 8224;

class W {
  b: number[] = [];
  u8(v: number): this { this.b.push(v & 0xff); return this; }
  u16(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff); return this; }
  u32(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); return this; }
  i16(v: number): this { return this.u16(v & 0xffff); }
  f64(v: number): this { this.b.push(...new Uint8Array(new Float64Array([v]).buffer)); return this; }
  bytes(a: ArrayLike<number>): this { for (let i = 0; i < a.length; i++) this.b.push(a[i]!); return this; }
  zeros(n: number): this { for (let i = 0; i < n; i++) this.b.push(0); return this; }
  /** Characters with a flags byte: Latin-1 when possible, else UTF-16. */
  chars(s: string): this {
    const hi = /[^\x00-\xff]/.test(s);
    this.u8(hi ? 1 : 0);
    for (let i = 0; i < s.length; i++) { if (hi) this.u16(s.charCodeAt(i)); else this.u8(s.charCodeAt(i)); }
    return this;
  }
  xlString(s: string): this { const t = s.slice(0, 32767); return this.u16(t.length).chars(t); }
  shortString(s: string): this { const t = s.slice(0, 255); return this.u8(t.length).chars(t); }
  get length(): number { return this.b.length; }
}

interface Rec { type: number; data: Uint8Array }
const rec = (type: number, w: W | number[] | Uint8Array = []): Rec => ({ type, data: w instanceof W ? new Uint8Array(w.b) : w instanceof Uint8Array ? w : new Uint8Array(w) });
/** Splits long data into a record plus CONTINUE records. */
const recLong = (type: number, data: Uint8Array): Rec[] => {
  const out: Rec[] = [{ type, data: data.subarray(0, MAX_REC) }];
  for (let o = MAX_REC; o < data.length; o += MAX_REC) out.push({ type: 0x003c, data: data.subarray(o, o + MAX_REC) });
  return out;
};

export interface XlsWriteOptions {
  /** Encrypts the file; without a password the workbook opens normally. */
  password?: string;
}

/** Things the older format can't hold, for a warning before saving. */
export const xlsLimits = (wb: Workbook): string[] => {
  const out = new Set<string>();
  for (const s of wb.sheets) {
    for (const k of s.cells.keys()) {
      if (keyRow(k) >= MAX_ROWS) out.add('Rows after 65,536 are left out.');
      if (keyCol(k) >= MAX_COLS) out.add('Columns after IV are left out.');
    }
    if (s.charts.some((c) => !OLD_CHART_TYPES.test(c.type))) out.add('Chart types the older format lacks are saved as the nearest simple type.');
    if (s.kind === 'chart') out.add('Chart sheets are saved as worksheets with the chart on them.');
    if (s.tables.length) out.add('Tables become plain ranges.');
    if (s.pictures.some((p) => !/png|jpe?g|bmp|x-emf|x-wmf/i.test(p.mime))) out.add('Pictures that are not PNG, JPEG, BMP, EMF or WMF are left out.');
    if (s.conditionalFormats.some((c) => c.rules.some((r) => r.type !== 'cellIs' && r.type !== 'expression'))) out.add('Some conditional formats are left out.');
  }
  if (wb.styles.length > 4000) out.add('Cell styles after the first 4,000 use the default style.');
  return [...out];
};

// ---------- Colours ----------

class Palette {
  colors = INDEXED_COLORS.slice(8, 64).map((c) => c.toUpperCase());
  private used = new Set<number>();
  private cache = new Map<string, number>();
  private reserved = new Set<number>();
  constructor(all: string[]) {
    const want = [...new Set(all.map((c) => c.slice(0, 7).toUpperCase()))];
    for (const c of want) { const i = this.colors.indexOf(c); if (i >= 0) this.reserved.add(i); }
    let slot = 55;
    for (const c of want) {
      if (this.colors.includes(c)) continue;
      while (slot >= 0 && this.reserved.has(slot)) slot--;
      if (slot < 0) break;
      this.colors[slot] = c;
      this.reserved.add(slot);
      slot--;
    }
  }
  /** Palette index (8-63) for a colour; the closest one when it isn't in the palette. */
  index(c: string | undefined): number | undefined {
    if (!c) return undefined;
    const key = c.slice(0, 7).toUpperCase();
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    let best = 0, bestD = Infinity;
    const n = parseInt(key.slice(1), 16);
    this.colors.forEach((p, i) => {
      const m = parseInt(p.slice(1), 16);
      const d = (((n >> 16) & 255) - ((m >> 16) & 255)) ** 2 + (((n >> 8) & 255) - ((m >> 8) & 255)) ** 2 + ((n & 255) - (m & 255)) ** 2;
      if (d < bestD) { bestD = d; best = i; }
    });
    this.used.add(best);
    this.cache.set(key, best + 8);
    return best + 8;
  }
}

const collectColors = (wb: Workbook): string[] => {
  const out: string[] = [];
  const font = (f?: CellStyle['font']) => { if (f?.color) out.push(f.color); };
  const style = (s: CellStyle) => {
    font(s.font);
    if (s.fill?.color) out.push(s.fill.color);
    if (s.fill?.patternColor) out.push(s.fill.patternColor);
    for (const b of Object.values(s.border ?? {})) if (b?.color) out.push(b.color);
  };
  wb.styles.slice(0, 4000).forEach(style);
  for (const s of wb.sheets) {
    if (s.tabColor) out.push(s.tabColor);
    for (const cf of s.conditionalFormats) for (const r of cf.rules) if ('style' in r) style(r.style);
    for (const c of s.cells.values()) for (const r of c.rich ?? []) font(r.font);
  }
  return out;
};

// ---------- Main ----------

export const writeXls = (wb: Workbook, opts: XlsWriteOptions = {}): Uint8Array => {
  const palette = new Palette(collectColors(wb));
  const sheets = wb.sheets;
  const sheetIdx = new Map(sheets.map((s, i) => [s.name.toUpperCase(), i]));

  // fonts: four copies of the default, index 4 is never used
  type F = NonNullable<CellStyle['font']>;
  const fontKeys = new Map<string, number>();
  const fontRecs: F[] = [];
  const fontIndex = (f: F | undefined): number => {
    const full: F = { name: f?.name ?? DEFAULT_FONT.name, size: f?.size ?? DEFAULT_FONT.size, ...f };
    const key = JSON.stringify([full.name, full.size, !!full.bold, !!full.italic, full.underline ?? '', !!full.strike, full.color ?? '', full.vertAlign ?? '']);
    const hit = fontKeys.get(key);
    if (hit !== undefined) return hit;
    fontRecs.push(full);
    const idx = fontRecs.length - 1;
    const ifnt = idx >= 4 ? idx + 1 : idx;
    fontKeys.set(key, ifnt);
    return ifnt;
  };
  for (let i = 0; i < 4; i++) { fontRecs.push({ name: DEFAULT_FONT.name, size: DEFAULT_FONT.size }); }
  fontKeys.set(JSON.stringify([DEFAULT_FONT.name, DEFAULT_FONT.size, false, false, '', false, '', '']), 0);

  const formatIds = new Map<string, number>();
  const formatRecs: [number, string][] = [];
  const formatId = (code: string | undefined): number => {
    if (!code || code === 'General') return 0;
    const b = builtinId(code);
    if (b !== undefined) return b;
    const hit = formatIds.get(code);
    if (hit !== undefined) return hit;
    const id = 164 + formatRecs.length;
    formatIds.set(code, id);
    formatRecs.push([id, code]);
    return id;
  };

  // XFs: 0-14 style records, 15 the default cell format, then one per workbook style
  const xfOf = (styleId: number | undefined): number => (!styleId || styleId >= 4000 || !wb.styles[styleId] ? 15 : 15 + styleId);
  const xfRecs: Rec[] = [];
  const xfData = (st: CellStyle, isStyle: boolean): W => {
    const w = new W();
    w.u16(fontIndex(st.font));
    w.u16(formatId(st.numFmt));
    const locked = st.protection?.locked !== false;
    w.u16((locked ? 1 : 0) | (st.protection?.hidden ? 2 : 0) | (isStyle ? 0xfff4 : 0));
    const a = st.align ?? {};
    const h = Math.max(0, H_ALIGNS.indexOf(a.h ?? 'general'));
    const v = a.v ? Math.max(0, V_ALIGNS.indexOf(a.v)) : 2;
    w.u8(h | (a.wrap ? 8 : 0) | (v << 4));
    const rot = a.rotation === undefined ? 0 : a.rotation === 255 ? 255 : a.rotation < 0 ? Math.min(180, 90 - a.rotation) : Math.min(90, a.rotation);
    w.u8(rot);
    w.u8(Math.min(15, a.indent ?? 0) | (a.shrink ? 0x10 : 0) | (a.rtl ? 0x80 : 0));
    w.u8(isStyle ? 0 : 0xfc);
    const b = st.border ?? {};
    const dg = (s?: { style: string }) => (s ? Math.max(0, BORDER_STYLES.indexOf(s.style as never)) : 0);
    const ic = (s?: { color?: string }) => (s ? palette.index(s.color) ?? 8 : 0);
    const diag = b.diagonalDown ?? b.diagonalUp;
    w.u32((dg(b.left) | (dg(b.right) << 4) | (dg(b.top) << 8) | (dg(b.bottom) << 12) | (ic(b.left) << 16) | (ic(b.right) << 23) | ((b.diagonalDown ? 1 : 0) << 30) | ((b.diagonalUp ? 1 : 0) << 31)) >>> 0);
    const f = st.fill;
    let fls = 0, fore = 64, back = 65;
    if (f?.pattern && f.pattern !== 'solid' && f.pattern !== 'none') {
      fls = Math.max(1, PATTERNS.indexOf(f.pattern));
      fore = palette.index(f.patternColor) ?? 64;
      back = palette.index(f.color) ?? 65;
    } else if (f?.color) { fls = 1; fore = palette.index(f.color)!; back = 65; }
    else if (f?.gradient?.stops[0]) { fls = 1; fore = palette.index(f.gradient.stops[0].color)!; }
    w.u32((ic(b.top) | (ic(b.bottom) << 7) | (ic(diag) << 14) | (dg(diag) << 21) | (fls << 26)) >>> 0);
    w.u16(fore | (back << 7));
    return w;
  };
  for (let i = 0; i < 15; i++) xfRecs.push(rec(0x00e0, xfData({}, true)));
  xfRecs.push(rec(0x00e0, xfData(wb.styles[0] ?? {}, false)));
  for (let i = 1; i < Math.min(wb.styles.length, 4000); i++) xfRecs.push(rec(0x00e0, xfData(wb.styles[i] ?? {}, false)));

  // names: user names first, then built-ins, then names for functions the table lacks
  interface NameRec { name: string; builtin?: number; sheet?: number; hidden?: boolean; func?: boolean; formula?: string }
  const names: NameRec[] = [];
  for (const n of wb.names) {
    const m = /^_xlnm\.(\w+)$/i.exec(n.name);
    const code = m ? BUILTIN_CODES[m[1]!.toLowerCase()] : undefined;
    names.push({ name: code !== undefined ? '' : n.name, builtin: code, sheet: n.sheet, hidden: n.hidden, formula: n.formula });
  }
  sheets.forEach((s, i) => {
    const q = quote(s.name);
    if (s.pageSetup?.printArea) names.push({ name: '', builtin: 6, sheet: i, formula: s.pageSetup.printArea.split(',').map((p) => `${q}!${absRef(p.trim())}`).join(',') });
    const titles = [s.pageSetup?.printTitleCols, s.pageSetup?.printTitleRows].filter(Boolean).map((p) => `${q}!${absRef(p!)}`);
    if (titles.length) names.push({ name: '', builtin: 7, sheet: i, formula: titles.join(',') });
    if (s.autoFilter) names.push({ name: '', builtin: 13, sheet: i, hidden: true, formula: `${q}!${absRange(s.autoFilter.range)}` });
  });
  const nameIdx = (name: string, sheet?: number): number | undefined => {
    const u = name.toUpperCase();
    let k = names.findIndex((n) => !n.builtin && !n.func && n.name.toUpperCase() === u && n.sheet === sheet);
    if (k < 0 && sheet === undefined) k = names.findIndex((n) => !n.func && n.builtin === undefined && n.name.toUpperCase() === u);
    if (k < 0) {
      const m = /^_xlnm\.(\w+)$/i.exec(name);
      const code = m ? BUILTIN_CODES[m[1]!.toLowerCase()] : undefined;
      if (code !== undefined) k = names.findIndex((n) => n.builtin === code && (sheet === undefined || n.sheet === sheet));
    }
    return k < 0 ? undefined : k + 1;
  };
  const addinNames: string[] = [];
  const funcName = (fn: string): number | { ixti: number; index: number } => {
    const f = ftabByName(fn);
    if (f && f.index >= 380 && f.index <= 479) {
      let k = addinNames.indexOf(fn);
      if (k < 0) { addinNames.push(fn); k = addinNames.length - 1; }
      return { ixti: ixti(-2, -2, 1), index: k + 1 };
    }
    const stored = addFuturePrefixes(`${fn}()`).slice(0, -2).replace(/^(?!_xlfn\.)/, '_xlfn.');
    const k = names.findIndex((n) => n.func && n.name === stored);
    if (k >= 0) return k + 1;
    names.push({ name: stored, func: true, hidden: true });
    return names.length;
  };
  const xti: [number, number, number][] = [];
  const ixti = (a: number, b: number, book = 0): number => {
    const k = xti.findIndex(([x, y, z]) => x === a && y === b && z === book);
    if (k >= 0) return k;
    xti.push([a, b, book]);
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
    mode, row, col,
    sheetIndex: (n) => sheetIdx.get(n.toUpperCase()),
    ixti: (a, b) => ixti(a, b),
    nameIndex: nameIdx,
    addin: funcName,
    table: tableRef(sheet),
  });
  const compile = (f: string, ctx: EncodeCtx, array = false): { rgce: Uint8Array; rgcb: Uint8Array } | null => {
    try { return encodeFormula(f, ctx, array); } catch { return null; }
  };

  // ---------- Shared strings ----------
  const sstList: { text: string; runs: [number, number][] }[] = [];
  const sstMap = new Map<string, number>();
  let sstTotal = 0;
  const sstIndex = (text: string, rich?: RichRun[]): number => {
    sstTotal++;
    const runs: [number, number][] = [];
    if (rich?.length) {
      let p = 0;
      for (const r of rich) { runs.push([p, fontIndex(r.font)]); p += r.text.length; }
    }
    const key = runs.length ? `${text}\u0000${JSON.stringify(runs)}` : text;
    const hit = sstMap.get(key);
    if (hit !== undefined) return hit;
    sstList.push({ text: text.slice(0, 32767), runs: runs.filter(([i]) => i < Math.min(text.length, 32767)) });
    sstMap.set(key, sstList.length - 1);
    return sstList.length - 1;
  };

  // ---------- Sheets (built first so formulas register names, sheets and strings) ----------
  const drawingGroup = new DrawingGroup();
  const sheetRecs = sheets.map((s, i) => sheetRecords(s, i));

  function sheetRecords(s: Sheet, index: number): Rec[] {
    const out: Rec[] = [];
    out.push(rec(0x0809, new W().u16(0x0600).u16(0x0010).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000006)));
    const calc = wb.calc ?? {};
    out.push(rec(0x000d, new W().u16(calc.manual ? 0 : 1)));
    out.push(rec(0x000c, new W().u16(calc.maxIterations ?? 100)));
    out.push(rec(0x000f, new W().u16(1)));
    out.push(rec(0x0011, new W().u16(calc.iterate ? 1 : 0)));
    out.push(rec(0x0010, new W().f64(calc.maxChange ?? 0.001)));
    out.push(rec(0x005f, new W().u16(1)));
    const ps = s.pageSetup ?? {};
    out.push(rec(0x002a, new W().u16(ps.headings ? 1 : 0)));
    out.push(rec(0x002b, new W().u16(ps.gridlines ? 1 : 0)));
    out.push(rec(0x0082, new W().u16(1)));
    const outline = Math.max(0, ...[...s.rows.values()].map((r) => r.outline ?? 0)), outlineC = Math.max(0, ...[...s.cols.values()].map((c) => c.outline ?? 0));
    out.push(rec(0x0080, new W().u16(outline ? 12 * outline + 17 : 0).u16(outlineC ? 12 * outlineC + 17 : 0).u16(outline ? outline + 1 : 0).u16(outlineC ? outlineC + 1 : 0)));
    out.push(rec(0x0225, new W().u16(0).u16(Math.round(s.defaultRowHeight * 20))));
    const fit = ps.fitWidth !== undefined || ps.fitHeight !== undefined;
    out.push(rec(0x0081, new W().u16(0x04c1 | (fit ? 0x100 : 0))));
    if (ps.rowBreaks?.length) { const w = new W().u16(Math.min(1026, ps.rowBreaks.length)); for (const r of ps.rowBreaks.slice(0, 1026)) w.u16(r).u16(0).u16(255); out.push(rec(0x001b, w)); }
    if (ps.colBreaks?.length) { const w = new W().u16(Math.min(1026, ps.colBreaks.length)); for (const c of ps.colBreaks.slice(0, 1026)) w.u16(c).u16(0).u16(65535); out.push(rec(0x001a, w)); }
    out.push(rec(0x0014, ps.header ? new W().xlString(ps.header.slice(0, 255)) : []));
    out.push(rec(0x0015, ps.footer ? new W().xlString(ps.footer.slice(0, 255)) : []));
    out.push(rec(0x0083, new W().u16(ps.centerH ? 1 : 0)));
    out.push(rec(0x0084, new W().u16(ps.centerV ? 1 : 0)));
    const m = ps.margins;
    if (m) {
      out.push(rec(0x0026, new W().f64(m.left)), rec(0x0027, new W().f64(m.right)), rec(0x0028, new W().f64(m.top)), rec(0x0029, new W().f64(m.bottom)));
    }
    out.push(rec(0x00a1, new W().u16(ps.paper ?? 9).u16(ps.scale ?? 100).u16(ps.firstPageNumber ?? 1).u16(ps.fitWidth ?? 1).u16(ps.fitHeight ?? 1)
      .u16((ps.orientation === 'landscape' ? 0 : 2) | (ps.firstPageNumber ? 0x80 : 0) | (ps.order === 'overThenDown' ? 1 : 0)).u16(600).u16(600)
      .f64(m?.header ?? 0.3).f64(m?.footer ?? 0.3).u16(1)));
    if (s.protection?.sheet) {
      out.push(rec(0x0012, new W().u16(1)));
      const o = s.protection.options ?? {};
      out.push(rec(0x0063, new W().u16(o.objects === false ? 0 : 1)));
      out.push(rec(0x00dd, new W().u16(o.scenarios === false ? 0 : 1)));
      const pw = parseInt(s.protection.password ?? '', 16);
      if (pw && /^[0-9a-f]{1,4}$/i.test(s.protection.password ?? '')) out.push(rec(0x0013, new W().u16(pw)));
    }
    if (s.autoFilter && s.autoFilter.range.c1 < MAX_COLS) {
      const af = s.autoFilter;
      const width = Math.min(af.range.c2, MAX_COLS - 1) - af.range.c1 + 1;
      out.push(rec(0x009d, new W().u16(width)));
      for (const fc of af.columns) {
        const w = autoFilterRecord(fc.col - af.range.c1, fc);
        if (w) out.push(rec(0x009e, w));
      }
    }
    out.push(rec(0x0055, new W().u16(Math.max(1, Math.round(s.defaultColWidth - 0.43)))));
    out.push(rec(0x0099, new W().u16(Math.min(65535, Math.round(fileColWidth(s.defaultColWidth) * 256)))));
    const colEntries = [...s.cols].filter(([c]) => c < MAX_COLS).sort((a, b) => a[0] - b[0]);
    for (const [c, info] of colEntries) {
      const w = Math.round(fileColWidth(info.width ?? s.defaultColWidth) * 256);
      out.push(rec(0x007d, new W().u16(c).u16(c).u16(Math.min(65535, w)).u16(xfOf(info.s)).u16((info.hidden ? 1 : 0) | (info.width !== undefined ? 2 : 0) | ((info.outline ?? 0) << 8)).u16(0)));
    }
    // cells by row
    const byRow = new Map<number, [number, Cell][]>();
    let maxR = -1, maxC = -1, minR = MAX_ROWS, minC = MAX_COLS;
    for (const [k, c] of s.cells) {
      const r = keyRow(k), col = keyCol(k);
      if (r >= MAX_ROWS || col >= MAX_COLS) continue;
      if (c.v === null && !c.f && !c.s) continue;
      let list = byRow.get(r);
      if (!list) byRow.set(r, (list = []));
      list.push([col, c]);
      maxR = Math.max(maxR, r); maxC = Math.max(maxC, col); minR = Math.min(minR, r); minC = Math.min(minC, col);
    }
    for (const r of s.rows.keys()) if (r < MAX_ROWS && !byRow.has(r)) byRow.set(r, []);
    out.push(rec(0x0200, new W().u32(maxR < 0 ? 0 : minR).u32(maxR + 1).u16(maxR < 0 ? 0 : minC).u16(maxC + 1).u16(0)));
    // array formula members point at their anchor
    const arrayAt = new Map<number, { r: number; c: number }>();
    for (const [k, c] of s.cells) {
      const a = c.arrayRange;
      if (!a || !c.f) continue;
      for (let r = a.r1; r <= Math.min(a.r2, MAX_ROWS - 1); r++) for (let cc = a.c1; cc <= Math.min(a.c2, MAX_COLS - 1); cc++) {
        if (r === keyRow(k) && cc === keyCol(k)) continue;
        arrayAt.set(r * MAX_COLS + cc, { r: keyRow(k), c: keyCol(k) });
        if (!byRow.has(r)) byRow.set(r, []);
        if (!byRow.get(r)!.some(([x]) => x === cc)) byRow.get(r)!.push([cc, s.cells.get(r * 16384 + cc) ?? { v: null }]);
      }
    }
    const rows = [...byRow.keys()].sort((a, b) => a - b);
    for (let i = 0; i < rows.length; i += 32) {
      const block = rows.slice(i, i + 32);
      for (const r of block) {
        const cells = byRow.get(r)!.sort((a, b) => a[0] - b[0]);
        const info = s.rows.get(r) ?? {};
        const first = cells[0]?.[0] ?? 0, last = cells.length ? cells.at(-1)![0] + 1 : 0;
        const h = info.height !== undefined ? Math.round(info.height * 20) : Math.round(s.defaultRowHeight * 20);
        const fl = ((info.outline ?? 0) & 7) | (info.hidden ? 0x20 : 0) | (info.height !== undefined ? 0x40 : 0) | (info.s ? 0x80 : 0) | 0x100;
        out.push(rec(0x0208, new W().u16(r).u16(first).u16(last).u16(h & 0x7fff).u16(0).u16(0).u16(fl).u16(info.s ? xfOf(info.s) : 15)));
      }
      for (const r of block) {
        for (const [c, cell] of byRow.get(r)!) out.push(...cellRecords(s, r, c, cell, arrayAt.get(r * MAX_COLS + c)));
      }
    }
    // drawing objects: pictures and notes
    out.push(...drawingRecords(s, index));
    // view
    const v = s.view;
    const frozen = v.frozenRows > 0 || v.frozenCols > 0;
    const active = index === wb.activeSheet;
    out.push(rec(0x023e, new W().u16((v.showGrid ? 2 : 0) | (v.showHeaders ? 4 : 0) | (frozen ? 8 | 0x100 : 0) | 0x10 | 0x20 | (v.rtl ? 0x40 : 0) | 0x80 | (active ? 0x600 : 0))
      .u16(0).u16(0).u16(0x40).u16(0).u16(0).u16(0).u16(0).u16(0)));
    if (v.zoom && v.zoom !== 100) { const g = gcd(v.zoom, 100); out.push(rec(0x00a0, new W().u16(v.zoom / g).u16(100 / g))); }
    let pnn = 3;
    if (frozen) {
      pnn = v.frozenRows && v.frozenCols ? 0 : v.frozenRows ? 2 : 1;
      out.push(rec(0x0041, new W().u16(v.frozenCols).u16(v.frozenRows).u16(v.frozenRows).u16(v.frozenCols).u8(pnn).u8(0)));
    }
    const sel = v.selection ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    const cl = (x: number, m: number) => Math.min(Math.max(0, x), m - 1);
    out.push(rec(0x001d, new W().u8(pnn).u16(cl(sel.r1, MAX_ROWS)).u16(cl(sel.c1, MAX_COLS)).u16(0).u16(1).u16(cl(sel.r1, MAX_ROWS)).u16(cl(sel.r2, MAX_ROWS)).u8(cl(sel.c1, MAX_COLS)).u8(cl(sel.c2, MAX_COLS))));
    const merges = s.merges.filter((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS);
    for (let i = 0; i < merges.length; i += 1026) {
      const part = merges.slice(i, i + 1026);
      const w = new W().u16(part.length);
      for (const g of part) w.u16(g.r1).u16(Math.min(g.r2, MAX_ROWS - 1)).u16(g.c1).u16(Math.min(g.c2, MAX_COLS - 1));
      out.push(rec(0x00e5, w));
    }
    out.push(...cfRecords(s));
    for (const [k, url] of s.hyperlinks) {
      const r = keyRow(k), c = keyCol(k);
      if (r >= MAX_ROWS || c >= MAX_COLS) continue;
      const h = hlink(r, c, url);
      if (h) out.push(h);
    }
    out.push(...dvRecords(s));
    if (s.tabColor) {
      const icv = palette.index(s.tabColor)!;
      out.push(rec(0x0862, new W().u16(0x0862).u16(0).zeros(8).u32(0x14).u32(icv)));
    }
    out.push(rec(0x000a));
    return out;
  }

  function cellRecords(s: Sheet, r: number, c: number, cell: Cell, arrayAnchor?: { r: number; c: number }): Rec[] {
    const xf = xfOf(cell.s);
    const head = () => new W().u16(r).u16(c).u16(xf);
    if (arrayAnchor) {
      const w = head();
      formulaValue(w, cell.v);
      w.u16(1).u32(0).u16(5).u8(0x01).u16(arrayAnchor.r).u16(arrayAnchor.c);
      return [rec(0x0006, w), ...(typeof cell.v === 'string' ? [rec(0x0207, new W().xlString(cell.v))] : [])];
    }
    if (cell.f) {
      const array = !!cell.arrayRange;
      const tokens = compile(cell.f, ctxFor('cell', s, r, c), array);
      if (tokens) {
        const w = head();
        formulaValue(w, cell.v);
        if (array) {
          w.u16(1).u32(0).u16(5).u8(0x01).u16(r).u16(c);
          const a = cell.arrayRange!;
          const arr = new W().u16(a.r1).u16(Math.min(a.r2, MAX_ROWS - 1)).u8(a.c1).u8(Math.min(a.c2, MAX_COLS - 1)).u16(1).u32(0).u16(tokens.rgce.length).bytes(tokens.rgce).bytes(tokens.rgcb);
          return [rec(0x0006, w), rec(0x0221, arr), ...(typeof cell.v === 'string' ? [rec(0x0207, new W().xlString(cell.v))] : [])];
        }
        w.u16(1).u32(0).u16(tokens.rgce.length).bytes(tokens.rgce).bytes(tokens.rgcb);
        return [rec(0x0006, w), ...(typeof cell.v === 'string' ? [rec(0x0207, new W().xlString(cell.v))] : [])];
      }
    }
    const v = cell.v;
    if (v === null) return [rec(0x0201, head())];
    if (typeof v === 'number') {
      const rk = toRk(v);
      return rk !== null ? [rec(0x027e, head().u32(rk))] : [rec(0x0203, head().f64(v))];
    }
    if (typeof v === 'boolean') return [rec(0x0205, head().u8(v ? 1 : 0).u8(0))];
    if (v instanceof CellError) return [rec(0x0205, head().u8(CODE_BY_ERR[v.code] ?? 0x2a).u8(1))];
    return [rec(0x00fd, head().u32(sstIndex(v, cell.rich)))];
  }

  function formulaValue(w: W, v: Scalar): void {
    if (typeof v === 'number') w.f64(v);
    else if (typeof v === 'string') w.bytes([v === '' ? 3 : 0, 0, 0, 0, 0, 0, 0xff, 0xff]);
    else if (typeof v === 'boolean') w.bytes([1, 0, v ? 1 : 0, 0, 0, 0, 0xff, 0xff]);
    else if (v instanceof CellError) w.bytes([2, 0, CODE_BY_ERR[v.code] ?? 0x2a, 0, 0, 0, 0xff, 0xff]);
    else w.f64(0);
  }

  function cfRecords(s: Sheet): Rec[] {
    const out: Rec[] = [];
    let id = 1;
    for (const cf of s.conditionalFormats) {
      const ranges = cf.ranges.filter((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS).map((g) => ({ ...g, r2: Math.min(g.r2, MAX_ROWS - 1), c2: Math.min(g.c2, MAX_COLS - 1) }));
      if (!ranges.length) continue;
      const base = ranges[0]!;
      const rules: Rec[] = [];
      for (const rule of cf.rules.slice(0, 3)) {
        const r = cfRule(rule, s, base);
        if (r) rules.push(r);
      }
      if (!rules.length) continue;
      const bound = ranges.reduce((a, g) => ({ r1: Math.min(a.r1, g.r1), c1: Math.min(a.c1, g.c1), r2: Math.max(a.r2, g.r2), c2: Math.max(a.c2, g.c2) }));
      const w = new W().u16(rules.length).u16((id++ << 1) | 1).u16(bound.r1).u16(bound.r2).u16(bound.c1).u16(bound.c2).u16(ranges.length);
      for (const g of ranges) w.u16(g.r1).u16(g.r2).u16(g.c1).u16(g.c2);
      out.push(rec(0x01b0, w), ...rules);
    }
    return out;
  }

  function cfRule(rule: CfRule, s: Sheet, base: CellRange): Rec | null {
    let ct: number, cp = 0;
    let f1: string | undefined, f2: string | undefined;
    if (rule.type === 'cellIs') { ct = 1; cp = CF_OP_CODES[rule.operator] ?? 3; f1 = rule.formulas[0]; f2 = rule.formulas[1]; }
    else if (rule.type === 'expression') { ct = 2; f1 = rule.formula; }
    else return null;
    const ctx = ctxFor('rel', s, base.r1, base.c1);
    const t1 = f1 !== undefined ? compile(f1, ctx) : null;
    const t2 = f2 !== undefined ? compile(f2, ctx) : null;
    if (f1 !== undefined && !t1) return null;
    const w = new W().u8(ct).u8(cp).u16(t1?.rgce.length ?? 0).u16(t2?.rgce.length ?? 0);
    dxfn(w, rule.style);
    if (t1) w.bytes(t1.rgce);
    if (t2) w.bytes(t2.rgce);
    return rec(0x01b1, w);
  }

  function dxfn(w: W, st: CellStyle): void {
    const font = st.font, fill = st.fill?.color, b = st.border;
    let f = 0x3ff | (b?.left ? 0 : 1 << 10) | (b?.right ? 0 : 1 << 11) | (b?.top ? 0 : 1 << 12) | (b?.bottom ? 0 : 1 << 13) | (3 << 14);
    if (!fill) f |= 7 << 16;
    f |= 1 << 19; // number format not set
    f |= 1 << 20; // font index not set
    f |= 7 << 22;
    if (font) f |= 1 << 26;
    if (b && (b.left || b.right || b.top || b.bottom)) f |= 1 << 28;
    if (fill) f |= 1 << 29;
    w.u32(f >>> 0).u16(0);
    if (font) {
      w.u8(0).zeros(63);
      w.u32(font.size ? Math.round(font.size * 20) : 0xffffffff);
      w.u32((font.italic ? 2 : 0) | (font.strike ? 0x80 : 0));
      w.u16(font.bold ? 700 : 400).u16(font.vertAlign === 'superscript' ? 1 : font.vertAlign === 'subscript' ? 2 : 0).u8(font.underline ? 1 : 0).u8(0).u16(0);
      w.u32(font.color ? palette.index(font.color)! : 0xffffffff).u32(0);
      w.u32((font.italic === undefined && font.bold === undefined ? 2 : 0) | (font.strike === undefined ? 0x80 : 0));
      w.u32(font.vertAlign ? 0 : 1).u32(font.underline ? 0 : 1).u32(font.bold === undefined ? 1 : 0).u32(0);
      w.u32(0xffffffff).u32(0).u16(1);
    }
    if (b && (b.left || b.right || b.top || b.bottom)) {
      const dg = (x?: { style: string }) => (x ? Math.max(0, BORDER_STYLES.indexOf(x.style as never)) : 0);
      const ic = (x?: { color?: string }) => (x ? palette.index(x.color) ?? 8 : 0);
      w.u32((dg(b.left) | (dg(b.right) << 4) | (dg(b.top) << 8) | (dg(b.bottom) << 12) | (ic(b.left) << 16) | (ic(b.right) << 23)) >>> 0);
      w.u32((ic(b.top) | (ic(b.bottom) << 7)) >>> 0);
    }
    if (fill) { const i = palette.index(fill)!; w.u16(1 << 10).u16(i | (i << 7)); }
  }

  function dvRecords(s: Sheet): Rec[] {
    const list = s.validations.filter((d) => d.ranges.some((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS));
    if (!list.length) return [];
    const out: Rec[] = [rec(0x01b2, new W().u16(0).u32(0).u32(0).u32(0xffffffff).u32(list.length))];
    for (const d of list) {
      const ranges = d.ranges.filter((g) => g.r1 < MAX_ROWS && g.c1 < MAX_COLS);
      const base = ranges[0]!;
      const typ = DV_TYPE_CODES[d.type] ?? 0;
      const ctx = ctxFor('rel', s, base.r1, base.c1);
      let t1: { rgce: Uint8Array } | null = null;
      let lookup = false;
      if (d.type === 'list' && d.formula1 && /^".*"$/s.test(d.formula1)) {
        const items = d.formula1.slice(1, -1).replace(/""/g, '"').split(',').join('\0');
        const w = new W().u8(0x17).u8(Math.min(255, items.length)).chars(items.slice(0, 255));
        t1 = { rgce: new Uint8Array(w.b) };
        lookup = true;
      } else if (d.formula1) t1 = compile(d.formula1, ctx);
      const t2 = d.formula2 ? compile(d.formula2, ctx) : null;
      const err = d.error ? ['stop', 'warning', 'information'].indexOf(d.error.style) : 0;
      const fl = typ | (Math.max(0, err) << 4) | (lookup ? 0x80 : 0) | (d.allowBlank ? 0x100 : 0) | (d.type === 'list' && d.showDropdown === false ? 0x200 : 0)
        | (d.prompt ? 0x40000 : 0) | (d.error ? 0x80000 : 0) | ((DV_OP_CODES[d.operator ?? 'between'] ?? 0) << 20);
      const str = (x?: string) => (x ? x.slice(0, 255) : '\0');
      const w = new W().u32(fl >>> 0).xlString(str(d.prompt?.title)).xlString(str(d.error?.title)).xlString(str(d.prompt?.text)).xlString(str(d.error?.text));
      w.u16(t1?.rgce.length ?? 0).u16(0).bytes(t1?.rgce ?? []);
      w.u16(t2?.rgce.length ?? 0).u16(0).bytes(t2?.rgce ?? []);
      w.u16(ranges.length);
      for (const g of ranges) w.u16(g.r1).u16(Math.min(g.r2, MAX_ROWS - 1)).u16(g.c1).u16(Math.min(g.c2, MAX_COLS - 1));
      out.push(rec(0x01be, w));
    }
    return out;
  }

  function drawingRecords(s: Sheet, index: number): Rec[] {
    const pics = s.pictures.map((p) => ({ p, blip: drawingGroup.add(p) })).filter((x) => x.blip > 0);
    const notes = [...s.notes].filter(([k]) => keyRow(k) < MAX_ROWS && keyCol(k) < MAX_COLS);
    const charts = s.charts;
    const shapes = s.shapes ?? [];
    const count = pics.length + notes.length + charts.length + shapes.length;
    if (!count) return [];
    const dgId = drawingGroup.newDrawing(count + 1);
    const base = dgId * 1024;
    let spid = base;
    // records after the first MSODRAWING; their drawing bytes count toward the container lengths
    const tail: Rec[] = [];
    const drawingBytes: Uint8Array[] = [];
    const shapeRecs: { draw: Uint8Array; obj: Rec; extra?: Rec[]; noText?: boolean }[] = [];
    let objId = 0;
    for (const { p, blip } of pics) {
      const id = ++spid;
      const a = cellsFromAnchor(s, p.anchor);
      const sp = spContainer(75, id, 0xa00, [[0x4104, blip], [0x01bf, 0x00100000]], a);
      const obj = new W().u16(0x15).u16(0x12).u16(8).u16(++objId).u16(0x6011).zeros(12)
        .u16(0x07).u16(2).u16(0xffff).u16(0x08).u16(2).u16(1).u16(0).u16(0);
      shapeRecs.push({ draw: sp, obj: rec(0x005d, obj) });
    }
    const noteRecs: Rec[] = [];
    for (const [k, note] of notes) {
      const r = keyRow(k), c = keyCol(k);
      const id = ++spid;
      const a = { c1: Math.min(255, c + 1), dx1: 15, r1: Math.max(0, r - 1), dy1: 10, c2: Math.min(255, c + 3), dx2: 15, r2: Math.min(65535, r + 3), dy2: 4 };
      const sp = spContainer(202, id, 0xa00, [[0x0080, 0], [0x00bf, 0x00080008], [0x0158, 0], [0x0181, 0x08000050], [0x0183, 0x08000050], [0x01bf, 0x00100010], [0x01c0, 0x08000051], [0x023f, 0x00030003], [0x03bf, 0x00020002]], a, true);
      const oid = ++objId;
      const obj = new W().u16(0x15).u16(0x12).u16(0x19).u16(oid).u16(0x4011).zeros(12).u16(0x0d).u16(0x16).bytes(md5(new TextEncoder().encode(`${index}:${k}`))).u16(0).u32(0).u16(0).u16(0);
      const text = noteText(note).slice(0, 32767);
      const txo = new W().u16(0x0212).u16(0).zeros(6).u16(text.length).u16(16).u16(0).u16(0);
      const textRec = new W().chars(text.replace(/\n/g, '\n'));
      const runs = new W().u16(0).u16(0).u32(0).u16(text.length).u16(0).u32(0);
      const clientText = new Uint8Array([0, 0, 0x0d, 0xf0, 0, 0, 0, 0]);
      shapeRecs.push({ draw: sp, obj: rec(0x005d, obj), extra: [rec(0x00ec, clientText), rec(0x01b6, txo), ...recLong(0x003c, new Uint8Array(textRec.b)), rec(0x003c, runs)] });
      noteRecs.push(rec(0x001c, new W().u16(r).u16(c).u16(0).u16(oid).xlString(note.author ?? '').u8(0)));
    }
    for (const c of charts) {
      const id = ++spid;
      const sp = spContainer(201, id, 0xa00, [[0x007f, 0x01040104], [0x00bf, 0x00080008], [0x0181, 0x0800004e], [0x0183, 0x0800004d], [0x01bf, 0x00110010], [0x01c0, 0x0800004d], [0x01ff, 0x00080008], [0x023f, 0x00020000], [0x03bf, 0x00080000]], cellsFromAnchor(s, c.anchor));
      const obj = new W().u16(0x15).u16(0x12).u16(5).u16(++objId).u16(0x6011).zeros(12).u16(0).u16(0);
      const stream = xlsChartRecords(c, (f) => compile(f, ctxFor('name', s))?.rgce ?? null).map((r) => rec(r.type, r.data));
      shapeRecs.push({ draw: sp, obj: rec(0x005d, obj), extra: stream, noText: true });
    }
    for (const sh of shapes) {
      const id = ++spid;
      const text = (sh.text ?? '').slice(0, 32767);
      const bgr = (c: string | undefined, dflt: number) => { const v = c ? parseInt(c.slice(1, 7), 16) : NaN; return Number.isFinite(v) ? ((v & 0xff) << 16) | (v & 0xff00) | (v >> 16) : dflt; };
      const props: [number, number][] = [[0x0080, 0], [0x00bf, 0x00080008], [0x0181, bgr(sh.fill, 0x00ffffff)], [0x01bf, sh.fill === 'none' ? 0x00100000 : 0x00100010],
        [0x01c0, bgr(sh.line?.color, 0x00000000)], [0x01cb, Math.round((sh.line?.width ?? 0.75) * 12700)], [0x01ff, 0x00080008]];
      if (sh.rotation) props.push([0x0004, Math.round(sh.rotation * 65536)]);
      const spt = SHAPE_TYPES[sh.kind] ?? 1;
      const sp = spContainer(spt, id, 0xa00, props.sort((a, b) => a[0] - b[0]), cellsFromAnchor(s, sh.anchor), !!text);
      const ot = sh.kind === 'textBox' ? 6 : sh.kind === 'rect' ? 2 : sh.kind === 'ellipse' ? 3 : sh.kind === 'line' || sh.kind === 'arrowLine' ? 1 : 0x1e;
      const obj = new W().u16(0x15).u16(0x12).u16(ot).u16(++objId).u16(0x6011).zeros(12).u16(0).u16(0);
      if (!text) { shapeRecs.push({ draw: sp, obj: rec(0x005d, obj), noText: true }); continue; }
      const h = sh.align === 'left' ? 1 : sh.align === 'right' ? 3 : 2;
      const txo = new W().u16((h << 1) | (2 << 4) | 0x200).u16(0).zeros(6).u16(text.length).u16(16).u16(0).u16(0);
      const runs = new W().u16(0).u16(0).u32(0).u16(text.length).u16(0).u32(0);
      const clientText = new Uint8Array([0, 0, 0x0d, 0xf0, 0, 0, 0, 0]);
      shapeRecs.push({ draw: sp, obj: rec(0x005d, obj), extra: [rec(0x00ec, clientText), rec(0x01b6, txo), ...recLong(0x003c, new Uint8Array(new W().chars(text).b)), rec(0x003c, runs)] });
    }
    // container headers: OfficeArtDgContainer > FDG, OfficeArtSpgrContainer > group shape, shapes
    const shapesLen = shapeRecs.reduce((n, x) => n + x.draw.length + (x.extra && !x.noText ? 8 : 0), 0);
    const groupShape = concat(art(0xf, 0, 0xf004, concat(art(1, 0, 0xf009, new Uint8Array(16)), art(2, 0, 0xf00a, le32s(base, 5))), 0));
    const spgrLen = groupShape.length + shapesLen;
    const fdg = art(0, dgId, 0xf008, le32s(count + 1, spid));
    const dgLen = fdg.length + 8 + spgrLen;
    const head = concat(hdr(0xf, 0, 0xf002, dgLen), fdg, hdr(0xf, 0, 0xf003, spgrLen), groupShape);
    shapeRecs.forEach((x, i) => {
      const data = i === 0 ? concat(head, x.draw) : x.draw;
      drawingBytes.push(data);
      tail.push(rec(0x00ec, data), x.obj, ...(x.extra ?? []));
    });
    return [...tail, ...noteRecs];
  }

  // ---------- Workbook globals ----------
  const namesRecs: Rec[] = [];
  const nameFormulas = names.map((n) => (n.formula ? compile(n.formula, ctxFor('name', n.sheet !== undefined ? sheets[n.sheet] : undefined)) : null));
  names.forEach((n, i) => {
    const body = n.func ? { rgce: new Uint8Array([0x1c, 0x1d]), rgcb: new Uint8Array(0) } : nameFormulas[i];
    const nm = n.builtin !== undefined ? String.fromCharCode(n.builtin) : n.name.slice(0, 255);
    const flags = (n.hidden ? 1 : 0) | (n.func ? 0x0a : 0) | (n.builtin !== undefined ? 0x20 : 0);
    const w = new W().u16(flags).u8(0).u8(nm.length).u16(body?.rgce.length ?? 0).u16(0).u16(n.sheet !== undefined ? n.sheet + 1 : 0).u32(0).chars(nm);
    if (body) w.bytes(body.rgce).bytes(body.rgcb);
    namesRecs.push(rec(0x0018, w));
  });

  const globals: Rec[] = [];
  globals.push(rec(0x0809, new W().u16(0x0600).u16(0x0005).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000006)));
  const filepassAt = globals.length;
  globals.push(rec(0x00e1, new W().u16(1200)));
  globals.push(rec(0x00c1, new W().u16(0)));
  globals.push(rec(0x00e2));
  globals.push(rec(0x005c, new W().xlString('ClerdoGrid').zeros(112 - 3 - 10)));
  globals.push(rec(0x0042, new W().u16(1200)));
  globals.push(rec(0x0161, new W().u16(0)));
  globals.push(rec(0x013d, (() => { const w = new W(); sheets.forEach((_, i) => w.u16(i + 1)); return w; })()));
  globals.push(rec(0x0019, new W().u16(0)));
  globals.push(rec(0x0012, new W().u16(wb.protection?.structure ? 1 : 0)));
  const wpw = parseInt(wb.protection?.password ?? '', 16);
  globals.push(rec(0x0013, new W().u16(wb.protection?.structure && /^[0-9a-f]{1,4}$/i.test(wb.protection.password ?? '') ? wpw : 0)));
  globals.push(rec(0x01af, new W().u16(0)));
  globals.push(rec(0x01bc, new W().u16(0)));
  const firstVisible = Math.max(0, sheets.findIndex((s) => s.state === 'visible'));
  globals.push(rec(0x003d, new W().u16(0x0168).u16(0x001e).u16(0x3a5c).u16(0x2328).u16(0x38).u16(wb.activeSheet).u16(firstVisible).u16(1).u16(600)));
  globals.push(rec(0x0040, new W().u16(0)));
  globals.push(rec(0x008d, new W().u16(0)));
  globals.push(rec(0x0022, new W().u16(wb.date1904 ? 1 : 0)));
  globals.push(rec(0x000e, new W().u16(1)));
  globals.push(rec(0x01b7, new W().u16(0)));
  globals.push(rec(0x00da, new W().u16(0)));
  for (const f of fontRecs) {
    const w = new W().u16(Math.round((f.size ?? 11) * 20)).u16((f.italic ? 2 : 0) | (f.strike ? 8 : 0)).u16(f.color ? palette.index(f.color)! : 0x7fff)
      .u16(f.bold ? 700 : 400).u16(f.vertAlign === 'superscript' ? 1 : f.vertAlign === 'subscript' ? 2 : 0)
      .u8(({ single: 1, double: 2, singleAccounting: 0x21, doubleAccounting: 0x22 } as Record<string, number>)[f.underline ?? ''] ?? 0).u8(0).u8(1).u8(0).shortString(f.name ?? DEFAULT_FONT.name);
    globals.push(rec(0x0031, w));
  }
  for (const [id, code] of formatRecs) globals.push(rec(0x041e, new W().u16(id).xlString(code)));
  globals.push(...xfRecs);
  globals.push(rec(0x0293, new W().u16(0x8000).u8(0).u8(0xff)));
  { const w = new W().u16(56); for (const c of palette.colors) { const n = parseInt(c.slice(1), 16); w.u8(n >> 16).u8(n >> 8).u8(n).u8(0); } globals.push(rec(0x0092, w)); }
  globals.push(rec(0x0160, new W().u16(1)));
  const boundAt = globals.length;
  sheets.forEach((s) => globals.push(rec(0x0085, new W().u32(0).u8(s.state === 'hidden' ? 1 : s.state === 'veryHidden' ? 2 : 0).u8(0).shortString(s.name.slice(0, 31)))));
  globals.push(rec(0x008c, new W().u16(1).u16(1)));
  if (xti.length || names.length) {
    globals.push(rec(0x01ae, new W().u16(sheets.length).u16(0x0401)));
    if (addinNames.length) {
      globals.push(rec(0x01ae, new W().u16(1).u16(0x3a01)));
      for (const n of addinNames) globals.push(rec(0x0023, new W().u16(0).u32(0).shortString(n).u16(2).u8(0x1c).u8(0x17)));
    }
    const w = new W().u16(xti.length);
    for (const [a, b, book] of xti) w.u16(book).u16(a & 0xffff).u16(b & 0xffff);
    globals.push(rec(0x0017, w));
  }
  globals.push(...namesRecs);
  const dgg = drawingGroup.bytes();
  if (dgg) globals.push(...recLong(0x00eb, dgg));
  const { recs: sstRecs, buckets } = buildSst(sstList, sstTotal);
  const sstAt = globals.length;
  globals.push(...sstRecs);
  const extsstAt = globals.length;
  globals.push(rec(0x00ff, new W().u16(8).zeros(buckets.length * 8)));
  globals.push(rec(0x000a));

  // encryption record (sizes known now, so positions can be fixed up)
  let cipherKey: ReturnType<typeof checkPassword> = null;
  if (opts.password !== undefined) {
    const salt = new Uint8Array(16);
    globalThis.crypto.getRandomValues(salt);
    const pw = opts.password || DEFAULT_PASSWORD;
    const p = makeRc4Params(pw, salt, 'cryptoapi');
    cipherKey = checkPassword(p, pw);
    const csp = 'Microsoft Enhanced Cryptographic Provider v1.0\0';
    const header = new W().u32(0x0c).u32(0).u32(0x6801).u32(0x8004).u32(128).u32(1).u32(0).u32(0);
    for (const ch of csp) header.u16(ch.charCodeAt(0));
    const w = new W().u16(1).u16(4).u16(2).u32(0x0c).u32(header.length).bytes(header.b).u32(16).bytes(p.salt).bytes(p.verifier).u32(20).bytes(p.verifierHash);
    globals.splice(filepassAt, 0, rec(0x002f, w));
  }

  // layout: positions of every record
  const all = [...globals, ...sheetRecs.flat()];
  const size = all.reduce((n, r) => n + 4 + r.data.length, 0);
  const out = new Uint8Array(size);
  const dv = new DataView(out.buffer);
  const pos: number[] = [];
  let o = 0;
  for (const r of all) { pos.push(o); o += 4 + r.data.length; }
  const offset = opts.password !== undefined ? 1 : 0;
  let sheetStart = globals.length;
  sheetRecs.forEach((list, i) => {
    const b = all[boundAt + offset + i]!;
    new DataView(b.data.buffer, b.data.byteOffset, 4).setUint32(0, pos[sheetStart]!, true);
    sheetStart += list.length;
  });
  // EXTSST: absolute position of every eighth string
  const ext = all[extsstAt + offset]!;
  const edv = new DataView(ext.data.buffer, ext.data.byteOffset, ext.data.byteLength);
  buckets.forEach((bk, i) => {
    edv.setUint32(2 + i * 8, pos[sstAt + offset + bk.rec]! + 4 + bk.off, true);
    edv.setUint16(6 + i * 8, bk.off + 4, true);
  });
  o = 0;
  for (const r of all) {
    dv.setUint16(o, r.type, true);
    dv.setUint16(o + 2, r.data.length, true);
    out.set(r.data, o + 4);
    o += 4 + r.data.length;
  }
  if (cipherKey) {
    const cipher = new StreamCipher(cipherKey);
    let seenFilepass = false;
    all.forEach((r, i) => {
      if (r.type === 0x002f) { seenFilepass = true; return; }
      if (!seenFilepass || NO_ENCRYPT.has(r.type) || !r.data.length) return;
      const from = pos[i]! + 4 + (r.type === 0x0085 ? 4 : 0);
      const end = pos[i]! + 4 + r.data.length;
      if (end > from) cipher.apply(out, from, end - from, from);
    });
  }
  const streams = new Map([['Workbook', out], ['\u0005SummaryInformation', summaryInfo(wb)]]);
  if (wb.queries?.length) {
    // ClerdoGrid's own stream; other apps ignore streams they don't know
    const q = new TextEncoder().encode(queriesToXml(wb.queries));
    if (cipherKey) new StreamCipher(cipherKey).apply(q, 0, q.length, 0);
    streams.set(QUERIES_STREAM, q);
  }
  return writeCfb(streams);
};

const NO_ENCRYPT = new Set([0x0809, 0x002f, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138]);
const BUILTIN_CODES: Record<string, number> = {
  consolidate_area: 0, auto_open: 1, auto_close: 2, extract: 3, database: 4, criteria: 5, print_area: 6, print_titles: 7, recorder: 8, data_form: 9,
  auto_activate: 10, auto_deactivate: 11, sheet_title: 12, _filterdatabase: 13,
};
const CF_OP_CODES: Record<string, number> = { between: 1, notBetween: 2, equal: 3, notEqual: 4, greaterThan: 5, lessThan: 6, greaterThanOrEqual: 7, lessThanOrEqual: 8 };
const DV_OP_CODES: Record<string, number> = { between: 0, notBetween: 1, equal: 2, notEqual: 3, greaterThan: 4, lessThan: 5, greaterThanOrEqual: 6, lessThanOrEqual: 7 };
const DV_TYPE_CODES: Record<string, number> = { any: 0, whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7 };

const FILTER_SIGN: Record<string, number> = { lessThan: 1, equal: 2, lessThanOrEqual: 3, greaterThan: 4, notEqual: 5, greaterThanOrEqual: 6 };

/** AUTOFILTER: up to two conditions, or a top/bottom rule, for one column. */
const autoFilterRecord = (entry: number, fc: FilterColumn): W | null => {
  if (entry < 0 || entry > 255) return null;
  let grbit = 0;
  let rules: { operator: string; value: string }[] = [];
  if (fc.top10) {
    grbit = 0x10 | (fc.top10.top ? 0x20 : 0) | (fc.top10.percent ? 0x40 : 0) | (Math.min(500, Math.max(1, Math.round(fc.top10.value))) << 7);
  } else if (fc.custom) { rules = fc.custom.rules.slice(0, 2); if (!fc.custom.and) grbit |= 1; }
  else if (fc.values && fc.values.length <= 2) { rules = fc.values.map((v) => ({ operator: 'equal', value: v })); grbit |= 1; }
  else if (fc.blanks) rules = [];
  else return null;
  const w = new W().u16(entry).u16(grbit);
  const strings: string[] = [];
  for (let i = 0; i < 2; i++) {
    const r = rules[i];
    if (!r) { w.zeros(10); continue; }
    const sign = FILTER_SIGN[r.operator] ?? 2;
    const n = Number(r.value);
    if (r.value !== '' && Number.isFinite(n)) w.u8(4).u8(sign).f64(n);
    else { const t = r.value.slice(0, 255); w.u8(6).u8(sign).u32(0).u8(t.length).u8(1).u8(0).u8(0); strings.push(t); }
  }
  for (const t of strings) w.chars(t);
  return w;
};

const gcd = (a: number, b: number): number => (b ? gcd(b, a % b) : a);
const colLetters = (c: number): string => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const absRange = (r: CellRange): string => `$${colLetters(r.c1)}$${r.r1 + 1}:$${colLetters(r.c2)}$${r.r2 + 1}`;
const absRef = (p: string): string => p.replace(/\$?([A-Za-z]{1,3})?\$?(\d+)?/g, (all, c: string | undefined, r: string | undefined) => (all ? `${c ? '$' + c.toUpperCase() : ''}${r ? '$' + r : ''}` : ''));
const quote = (name: string): string => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const noteText = (n: Note): string => n.text;

const toRk = (v: number): number | null => {
  if (Number.isInteger(v) && v >= -(2 ** 29) && v < 2 ** 29) return ((v << 2) | 2) >>> 0;
  const x100 = v * 100;
  if (Number.isInteger(x100) && x100 >= -(2 ** 29) && x100 < 2 ** 29 && x100 / 100 === v) return ((x100 << 2) | 3) >>> 0;
  const b = new DataView(new ArrayBuffer(8));
  b.setFloat64(0, v, true);
  if (b.getUint32(0, true) === 0 && (b.getUint32(4, true) & 3) === 0) return b.getUint32(4, true) >>> 0;
  return null;
};

const concat = (...parts: Uint8Array[]): Uint8Array => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
};
const le32s = (...v: number[]): Uint8Array => { const w = new W(); for (const x of v) w.u32(x); return new Uint8Array(w.b); };
const hdr = (ver: number, inst: number, type: number, len: number): Uint8Array => new Uint8Array(new W().u16(ver | (inst << 4)).u16(type).u32(len).b);
const art = (ver: number, inst: number, type: number, data: Uint8Array, len = data.length): Uint8Array => concat(hdr(ver, inst, type, len), data);

/** A shape container: FSP, properties, cell anchor and client data (plus room for a text box record). */
const SHAPE_TYPES: Record<string, number> = {
  rect: 1, roundRect: 2, ellipse: 3, diamond: 4, triangle: 5, hexagon: 9, star5: 12, rightArrow: 13, line: 20, arrowLine: 20, pentagon: 56,
  leftArrow: 66, downArrow: 67, upArrow: 68, callout: 61, textBox: 202,
};

const spContainer = (shapeType: number, spid: number, flags: number, props: [number, number][], a: ReturnType<typeof cellsFromAnchor>, textbox = false): Uint8Array => {
  const fsp = art(2, shapeType, 0xf00a, le32s(spid, flags));
  const ow = new W();
  for (const [id, v] of props) ow.u16(id).u32(v);
  const fopt = art(3, props.length, 0xf00b, new Uint8Array(ow.b));
  const anchor = art(0, 0, 0xf010, new Uint8Array(new W().u16(2).u16(a.c1).u16(a.dx1).u16(a.r1).u16(a.dy1).u16(a.c2).u16(a.dx2).u16(a.r2).u16(a.dy2).b));
  const client = hdr(0, 0, 0xf011, 0);
  const body = concat(fsp, fopt, anchor, client);
  return concat(hdr(0xf, 0, 0xf004, body.length + (textbox ? 8 : 0)), body);
};

/** Pictures for the drawing group's blip store, and the drawing id counters. */
class DrawingGroup {
  private blips: Uint8Array[] = [];
  private keys = new Map<string, number>();
  private drawings: number[] = [];
  add(p: Picture): number {
    let data = p.data;
    let kind: 'png' | 'jpeg' | 'dib';
    if (/png/i.test(p.mime)) kind = 'png';
    else if (/jpe?g/i.test(p.mime)) kind = 'jpeg';
    else if (/bmp/i.test(p.mime) && data.length > 14 && data[0] === 0x42 && data[1] === 0x4d) { kind = 'dib'; data = data.subarray(14); }
    else if (/x-emf|x-wmf/i.test(p.mime)) return this.addMetafile(p);
    else return 0;
    const uid = md5(data);
    const key = [...uid].join(',');
    const hit = this.keys.get(key);
    if (hit) return hit;
    const recType = kind === 'png' ? 0xf01e : kind === 'jpeg' ? 0xf01d : 0xf01f;
    const inst = kind === 'png' ? 0x6e0 : kind === 'jpeg' ? 0x46a : 0x7a8;
    const blip = art(0, inst, recType, concat(uid, new Uint8Array([0xff]), data));
    const bt = kind === 'png' ? 6 : kind === 'jpeg' ? 5 : 7;
    const fbse = new W().u8(bt).u8(bt).bytes(uid).u16(0xff).u32(blip.length).u32(1).u32(0).u8(0).u8(0).u8(0).u8(0);
    this.blips.push(art(2, bt, 0xf007, concat(new Uint8Array(fbse.b), blip)));
    this.keys.set(key, this.blips.length);
    return this.blips.length;
  }
  private addMetafile(p: Picture): number {
    const emf = /emf/i.test(p.mime);
    const uid = md5(p.data);
    const key = [...uid].join(',');
    const hit = this.keys.get(key);
    if (hit) return hit;
    const packed = zlibSync(p.data);
    const emu = (px: number) => Math.round(px * 9525);
    const head = new W().bytes(uid).u32(p.data.length).u32(0).u32(0).u32(Math.round(p.anchor.width * 15)).u32(Math.round(p.anchor.height * 15))
      .u32(emu(p.anchor.width)).u32(emu(p.anchor.height)).u32(packed.length).u8(0).u8(0xfe);
    const blip = art(0, emf ? 0x3d4 : 0x216, emf ? 0xf01a : 0xf01b, concat(new Uint8Array(head.b), packed));
    const bt = emf ? 2 : 3;
    const fbse = new W().u8(bt).u8(bt).bytes(uid).u16(0xff).u32(blip.length).u32(1).u32(0).u8(0).u8(0).u8(0).u8(0);
    this.blips.push(art(2, bt, 0xf007, concat(new Uint8Array(fbse.b), blip)));
    this.keys.set(key, this.blips.length);
    return this.blips.length;
  }
  newDrawing(shapes: number): number { this.drawings.push(shapes); return this.drawings.length; }
  bytes(): Uint8Array | null {
    if (!this.drawings.length) return null;
    const n = this.drawings.length;
    const fdgg = new W().u32((n + 1) * 1024).u32(n + 1).u32(this.drawings.reduce((a, b) => a + b, 0)).u32(n);
    this.drawings.forEach((count, i) => fdgg.u32(i + 1).u32(count + 1));
    const parts = [art(0, 0, 0xf006, new Uint8Array(fdgg.b))];
    if (this.blips.length) parts.push(art(0xf, this.blips.length, 0xf001, concat(...this.blips)));
    const opt = new W().u16(0x00bf).u32(0x00080008).u16(0x0181).u32(0x08000041).u16(0x01c0).u32(0x08000040);
    parts.push(art(3, 3, 0xf00b, new Uint8Array(opt.b)));
    const body = concat(...parts);
    return art(0xf, 0, 0xf000, body);
  }
}

/** SST plus CONTINUE records; strings never split inside their header, character runs re-state their flags. */
const buildSst = (list: { text: string; runs: [number, number][] }[], total: number): { recs: Rec[]; buckets: { rec: number; off: number }[] } => {
  const recs: Rec[] = [];
  const buckets: { rec: number; off: number }[] = [];
  let cur = new W().u32(total).u32(list.length);
  const flush = () => { recs.push(rec(recs.length ? 0x003c : 0x00fc, cur)); cur = new W(); };
  list.forEach((s, i) => {
    const hi = /[^\x00-\xff]/.test(s.text);
    const headLen = 3 + (s.runs.length ? 2 : 0);
    if (cur.length + headLen + (hi ? 2 : 1) > MAX_REC) flush();
    if (i % 8 === 0) buckets.push({ rec: recs.length, off: cur.length });
    cur.u16(s.text.length).u8((hi ? 1 : 0) | (s.runs.length ? 8 : 0));
    if (s.runs.length) cur.u16(s.runs.length);
    let p = 0;
    while (p < s.text.length) {
      const room = Math.floor((MAX_REC - cur.length) / (hi ? 2 : 1));
      if (room <= 0) { flush(); cur.u8(hi ? 1 : 0); continue; }
      const part = s.text.slice(p, p + room);
      for (let k = 0; k < part.length; k++) { if (hi) cur.u16(part.charCodeAt(k)); else cur.u8(part.charCodeAt(k)); }
      p += part.length;
      if (p < s.text.length) { flush(); cur.u8(hi ? 1 : 0); }
    }
    for (const [ich, ifnt] of s.runs) {
      if (cur.length + 4 > MAX_REC) flush();
      cur.u16(ich).u16(ifnt);
    }
  });
  flush();
  return { recs, buckets };
};

/** HLINK for a web, mail or in-workbook link. */
const hlink = (r: number, c: number, target: string): Rec | null => {
  if (!/^(https?:|mailto:|#)/i.test(target)) return null;
  const w = new W().u16(r).u16(r).u16(c).u16(c);
  w.bytes([0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b]).u32(2);
  const hash = target.indexOf('#');
  const url = target.startsWith('#') ? '' : hash > 0 && !/^mailto:/i.test(target) ? target.slice(0, hash) : target;
  const loc = target.startsWith('#') ? target.slice(1) : hash > 0 && !/^mailto:/i.test(target) ? target.slice(hash + 1) : '';
  w.u32((url ? 3 : 0) | (loc ? 8 : 0));
  if (url) {
    w.bytes([0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b]);
    w.u32((url.length + 1) * 2);
    for (const ch of url) w.u16(ch.charCodeAt(0));
    w.u16(0);
  }
  if (loc) { w.u32(loc.length + 1); for (const ch of loc) w.u16(ch.charCodeAt(0)); w.u16(0); }
  return rec(0x01b8, w);
};

/** A minimal summary-information property set naming the application. */
const summaryInfo = (wb: Workbook): Uint8Array => {
  const props: [number, string][] = [[0x12, 'ClerdoGrid']];
  if (wb.props?.title) props.push([0x02, wb.props.title]);
  if (wb.props?.creator) props.push([0x04, wb.props.creator]);
  const body = new W();
  const entries: [number, number][] = [];
  const valuesStart = 8 + props.length * 8 + 8;
  const values = new W();
  values.u32(2).u16(1).u16(0).u16(1252).u16(0);
  entries.push([1, valuesStart]);
  for (const [id, s] of props) {
    const off = valuesStart + values.length;
    const bytes = [...s.slice(0, 250)].map((ch) => (ch.charCodeAt(0) < 256 ? ch.charCodeAt(0) : 63));
    values.u32(30).u32(bytes.length + 1).bytes(bytes).u8(0);
    while (values.length % 4) values.u8(0);
    entries.push([id, off]);
  }
  body.u32(8 + entries.length * 8 + values.length).u32(entries.length);
  for (const [id, off] of entries) body.u32(id).u32(off);
  body.bytes(values.b);
  const head = new W().u16(0xfffe).u16(0).u32(0x00020006).zeros(16).u32(1)
    .bytes([0xe0, 0x85, 0x9f, 0xf2, 0xf9, 0x4f, 0x68, 0x10, 0xab, 0x91, 0x08, 0x00, 0x2b, 0x27, 0xb3, 0xd9]).u32(48);
  return new Uint8Array([...head.b, ...body.b]);
};
