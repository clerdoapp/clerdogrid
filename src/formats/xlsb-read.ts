// Binary workbook reader (.xlsb per the MS-XLSB open specification): an OPC package of binary record parts.
import { unzipSync } from 'fflate';
import {
  cellKey, internStyle, newSheet, type BorderSide, type Cell, type CellRange, type CellStyle, type CfOperator, type CfRule, type ConditionalFormat, type DataValidation, type Picture, type RichRun, type Sheet, type TableDef, type Workbook,
} from '../model/workbook';
import { CellError, type Scalar } from '../engine/types';
import { attr, find, kid, kids, parseRoot, type XNode } from './xml';
import { ERR_BY_CODE, Exp, decodeRgce, type DecodeCtx } from './biff8-formula';
import { BORDER_STYLES, H_ALIGNS, PATTERNS, Rd, UNDERLINES, V_ALIGNS, hex } from './biff8-common';
import { INDEXED_COLORS, builtinFormat } from './numfmt';
import { applyTint, readTheme } from './xlsx-styles';
import { readAnchor, shownColWidth } from './xlsx-chart';
import { readAnchorChart } from './xlsx-chartex';
import { OBJECTS_REL_SHORT, applyObjectsPart, readShape } from './xlsx-objects';
import { readExtras } from './xlsb-extras';

const MAX_PART = 1 << 30;
const SAFE_LINK = /^(https:|mailto:|#)/i;

interface Rec { type: number; data: Uint8Array }

/** Splits a binary part into records (variable-length type and size). */
export const xlsbRecords = (b: Uint8Array): Rec[] => {
  const out: Rec[] = [];
  let o = 0;
  while (o < b.length) {
    let type = 0;
    for (let i = 0; i < 2 && o < b.length; i++) { const x = b[o++]!; type |= (x & 0x7f) << (7 * i); if (!(x & 0x80)) break; }
    let size = 0;
    for (let i = 0; i < 4 && o < b.length; i++) { const x = b[o++]!; size |= (x & 0x7f) << (7 * i); if (!(x & 0x80)) break; }
    if (o + size > b.length) break;
    out.push({ type, data: b.subarray(o, o + size) });
    o += size;
  }
  return out;
};

const dirOf = (p: string): string => p.slice(0, p.lastIndexOf('/') + 1);
const resolve = (base: string, target: string): string => {
  if (target.startsWith('/')) return target.slice(1);
  const out: string[] = [];
  for (const p of (dirOf(base) + target).split('/')) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); }
  return out.join('/');
};

const rk = (v: number): number => {
  let n: number;
  if (v & 2) n = v >> 2;
  else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xfffffffc, true); n = b.getFloat64(0, true); }
  return v & 1 ? n / 100 : n;
};

/** Reads an .xlsb workbook. */
export const readXlsb = (data: Uint8Array): Workbook => {
  let total = 0;
  const files = unzipSync(data, {
    filter: (f) => { total += f.originalSize; if (f.originalSize > MAX_PART || total > 2 ** 31) throw new Error('The file is too large to open.'); return true; },
  });
  const lower = new Map(Object.keys(files).map((k) => [k.toLowerCase(), k]));
  const get = (p: string): Uint8Array | undefined => files[p] ?? files[lower.get(p.toLowerCase()) ?? ''];
  const text = (p: string): string | undefined => { const b = get(p); return b ? new TextDecoder().decode(b) : undefined; };
  const rels = (part: string): { id: string; type: string; target: string; external: boolean }[] => {
    const t = text(`${dirOf(part)}_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`);
    if (!t) return [];
    return kids(parseRoot(t), 'Relationship').map((r) => {
      const external = attr(r, 'TargetMode') === 'External';
      const target = attr(r, 'Target') ?? '';
      const type = attr(r, 'Type') ?? '';
      return { id: attr(r, 'Id') ?? '', type: type.slice(type.lastIndexOf('/') + 1), target: external ? target : resolve(part, target), external };
    });
  };
  const root = rels('');
  const wbPath = root.find((r) => r.type === 'officeDocument')?.target ?? 'xl/workbook.bin';
  const wbBin = get(wbPath);
  if (!wbBin) throw new Error("This file isn't a workbook ClerdoGrid can read.");
  const wbRels = rels(wbPath);
  const relTarget = (type: string) => wbRels.find((r) => r.type === type)?.target;
  const theme = readTheme(relTarget('theme') ? text(relTarget('theme')!) : undefined);

  const wb: Workbook = { sheets: [], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const bundles: { state: number; rel: string; name: string }[] = [];
  const rawNames: { flags: number; itab: number; name: string; rgce: Uint8Array; rgcb: Uint8Array }[] = [];
  const supbooks: { self: boolean; names: string[] }[] = [];
  const xti: { book: number; first: number; last: number }[] = [];
  let activeTab = 0;
  const calc: NonNullable<Workbook['calc']> = {};

  for (const r of xlsbRecords(wbBin)) {
    const d = new Rd(r.data);
    try {
      switch (r.type) {
        case 153: if (d.u32() & 1) wb.date1904 = true; break;
        case 156: { const state = d.u32(); d.u32(); const rel = d.wide(); bundles.push({ state, rel, name: d.wide() }); break; }
        case 39: {
          const flags = d.u32(); d.u8(); const itab = d.u32(); const name = d.wide();
          const cce = d.u32(); const rgce = d.bytes(cce); const cb = d.u32(); const rgcb = d.bytes(cb);
          rawNames.push({ flags, itab, name, rgce, rgcb });
          break;
        }
        case 357: case 358: supbooks.push({ self: true, names: [] }); break;
        case 355: case 667: supbooks.push({ self: false, names: [] }); break;
        case 362: { const n = d.u32(); for (let i = 0; i < n; i++) xti.push({ book: d.u32(), first: d.i32(), last: d.i32() }); break; }
        case 158: d.skip(24); activeTab = d.u32(); break;
        case 157: {
          d.u32(); const auto = d.u32(); const count = d.u32(); const delta = d.f64(); d.i32(); const fl = d.u16();
          if (auto === 0) calc.manual = true;
          if (fl & 4) { calc.iterate = true; calc.maxIterations = count; calc.maxChange = delta; }
          break;
        }
        case 534: { d.u16(); d.u16(); const fl = d.u16(); if (fl & 1) wb.protection = { structure: true }; break; }
        default: break;
      }
    } catch { /* skip damaged records */ }
  }

  // ---------- Styles ----------
  type Color = string | undefined;
  const color = (d: Rd): Color => {
    const fl = d.u8(), idx = d.u8(), tint = d.i16(), r = d.u8(), g = d.u8(), b = d.u8();
    d.u8();
    const kind = fl >> 1;
    let c: string | undefined;
    if (kind === 2) c = hex(r, g, b);
    else if (kind === 1) c = idx < 64 ? INDEXED_COLORS[idx] : undefined;
    else if (kind === 3) c = theme[idx];
    return c && tint ? applyTint(c, tint / 32767) : c;
  };
  /** BrtDXF: a list of XFProp entries (type, size, data). */
  const readDxf = (d: Rd): CellStyle => {
    d.u16();
    const n = d.u16();
    const st: CellStyle = {};
    const font = (): NonNullable<CellStyle['font']> => (st.font ??= {});
    const border = (): NonNullable<CellStyle['border']> => (st.border ??= {});
    const align = (): NonNullable<CellStyle['align']> => (st.align ??= {});
    let pattern: number | undefined, fg: Color, bg: Color;
    for (let i = 0; i < n && d.left >= 4; i++) {
      const type = d.u16(), cb = d.u16();
      const end = d.o + cb - 4;
      const p = new Rd(d.bytes(Math.max(0, Math.min(cb - 4, d.left))));
      try {
        switch (type) {
          case 0: pattern = p.u8(); break;
          case 1: fg = color(p); break;
          case 2: bg = color(p); break;
          case 5: { const c = color(p); if (c) font().color = c; break; }
          case 6: case 7: case 8: case 9: {
            const c = color(p); const s = BORDER_STYLES[p.u16()];
            if (s) border()[(['top', 'bottom', 'left', 'right'] as const)[type - 6]] = c && c !== '#000000' ? { style: s, color: c } : { style: s };
            break;
          }
          case 15: align().h = H_ALIGNS[p.u8()]; break;
          case 16: align().v = V_ALIGNS[p.u8()]; break;
          case 20: if (p.u8()) align().wrap = true; break;
          case 24: { const n = p.u16(); font().name = p.utf16(Math.min(n, 255)); break; }
          case 25: font().bold = p.u16() >= 700; break;
          case 26: { const u = UNDERLINES[p.u16()]; if (u) font().underline = u; break; }
          case 27: { const v = p.u16(); if (v) font().vertAlign = v === 1 ? 'superscript' : 'subscript'; break; }
          case 28: font().italic = !!p.u8(); break;
          case 29: font().strike = !!p.u8(); break;
          case 36: font().size = p.u32() / 20; break;
          case 38: { const n = p.u16(); const hi = p.u8() & 1; const code = hi ? p.utf16(n) : p.latin1(n); if (code && code !== 'General') st.numFmt = code; break; }
          case 41: { const id = p.u16(); const code = formats.get(id) ?? builtinFormat(id); if (code && code !== 'General') st.numFmt = code; break; }
          case 43: (st.protection ??= {}).locked = !!p.u8(); break;
          case 44: (st.protection ??= {}).hidden = !!p.u8(); break;
          default: break;
        }
      } catch { /* skip a damaged entry */ }
      d.o = Math.max(d.o, end);
    }
    if (pattern === 1 || (pattern === undefined && (fg || bg))) { const c = fg ?? bg; if (c) st.fill = { color: c }; }
    else if (pattern && PATTERNS[pattern]) st.fill = { pattern: PATTERNS[pattern], ...(bg ? { color: bg } : {}), ...(fg ? { patternColor: fg } : {}) };
    return st;
  };
  const fonts: CellStyle['font'][] = [];
  const fills: CellStyle['fill'][] = [];
  const borders: CellStyle['border'][] = [];
  const formats = new Map<number, string>();
  const xfs: CellStyle[] = [];
  const dxfs: CellStyle[] = [];
  const stylesPath = relTarget('styles');
  const sb = stylesPath ? get(stylesPath) : undefined;
  let inCellXfs = false;
  for (const r of sb ? xlsbRecords(sb) : []) {
    const d = new Rd(r.data);
    try {
      switch (r.type) {
        case 44: { const id = d.u16(); formats.set(id, d.wide()); break; }
        case 43: {
          const size = d.u16() / 20, grbit = d.u16(), bls = d.u16(), sss = d.u16(), uls = d.u8();
          d.skip(3);
          const c = color(d);
          d.u8();
          const name = d.wide();
          const f: NonNullable<CellStyle['font']> = {};
          if (name && name !== 'Calibri') f.name = name;
          if (size && size !== 11) f.size = size;
          if (bls >= 700) f.bold = true;
          if (grbit & 2) f.italic = true;
          if (grbit & 8) f.strike = true;
          if (UNDERLINES[uls]) f.underline = UNDERLINES[uls];
          if (sss) f.vertAlign = sss === 1 ? 'superscript' : 'subscript';
          if (c && c !== '#000000') f.color = c;
          fonts.push(Object.keys(f).length ? f : undefined);
          break;
        }
        case 45: {
          const fls = d.u32(), fg = color(d), bg = color(d);
          if (fls === 0) fills.push(undefined);
          else if (fls === 1) fills.push(fg ? { color: fg } : undefined);
          else fills.push({ pattern: PATTERNS[fls] ?? 'gray125', ...(bg ? { color: bg } : {}), ...(fg ? { patternColor: fg } : {}) });
          break;
        }
        case 46: {
          const fl = d.u8();
          const side = (): BorderSide | undefined => { const dg = d.u8(); d.u8(); const c = color(d); const s = BORDER_STYLES[dg]; return s ? (c && c !== '#000000' ? { style: s, color: c } : { style: s }) : undefined; };
          const top = side(), bottom = side(), left = side(), right = side(), diag = side();
          const b: NonNullable<CellStyle['border']> = {};
          if (left) b.left = left;
          if (right) b.right = right;
          if (top) b.top = top;
          if (bottom) b.bottom = bottom;
          if (diag && fl & 1) b.diagonalDown = diag;
          if (diag && fl & 2) b.diagonalUp = diag;
          borders.push(Object.keys(b).length ? b : undefined);
          break;
        }
        case 507: dxfs.push(readDxf(d)); break;
        case 617: inCellXfs = true; break;
        case 618: inCellXfs = false; break;
        case 47: {
          if (!inCellXfs) break;
          d.u16();
          const ifmt = d.u16(), ifnt = d.u16(), ifill = d.u16(), ibdr = d.u16(), trot = d.u8(), indent = d.u8(), fl = d.u16();
          const st: CellStyle = {};
          const font = fonts[ifnt];
          if (font) st.font = { ...font };
          const code = formats.get(ifmt) ?? builtinFormat(ifmt);
          if (code && code !== 'General') st.numFmt = code;
          const fill = fills[ifill];
          if (fill) st.fill = { ...fill };
          const b = borders[ibdr];
          if (b) st.border = { ...b };
          const a: NonNullable<CellStyle['align']> = {};
          const h = fl & 7, v = (fl >> 3) & 7;
          if (h) a.h = H_ALIGNS[h];
          if (v !== 2 && V_ALIGNS[v]) a.v = V_ALIGNS[v];
          if (fl & 0x40) a.wrap = true;
          if (indent) a.indent = indent;
          if (trot) a.rotation = trot > 90 && trot <= 180 ? 90 - trot : trot;
          if (fl & 0x100) a.shrink = true;
          if (((fl >> 10) & 3) === 2) a.rtl = true;
          if (Object.keys(a).length) st.align = a;
          if (!(fl & 0x1000) || fl & 0x2000) st.protection = { ...(fl & 0x1000 ? {} : { locked: false }), ...(fl & 0x2000 ? { hidden: true } : {}) };
          xfs.push(st);
          break;
        }
        default: break;
      }
    } catch { /* skip */ }
  }
  const styleIds = new Map<number, number>();
  const styleOf = (ixfe: number): number => {
    let id = styleIds.get(ixfe);
    if (id === undefined) { id = xfs[ixfe] ? internStyle(wb, xfs[ixfe]!) : 0; styleIds.set(ixfe, id); }
    return id;
  };

  // ---------- Shared strings ----------
  const sst: { text: string; rich?: RichRun[] }[] = [];
  const sstPath = relTarget('sharedStrings');
  const sstBin = sstPath ? get(sstPath) : undefined;
  const richStr = (d: Rd): { text: string; rich?: RichRun[] } => {
    const fl = d.u8();
    const t = d.wide();
    if (!(fl & 1)) return { text: t };
    const n = d.u32();
    const runs: [number, number][] = [];
    for (let i = 0; i < n && d.left >= 4; i++) runs.push([d.u16(), d.u16()]);
    if (!runs.length) return { text: t };
    const rich: RichRun[] = [];
    if (runs[0]![0] > 0) rich.push({ text: t.slice(0, runs[0]![0]) });
    runs.forEach(([ich, ifnt], k) => {
      const part = t.slice(ich, runs[k + 1]?.[0] ?? t.length);
      if (!part) return;
      const f = fonts[ifnt];
      rich.push(f ? { text: part, font: { ...f } } : { text: part });
    });
    return { text: t, rich };
  };
  for (const r of sstBin ? xlsbRecords(sstBin) : []) if (r.type === 19) { try { sst.push(richStr(new Rd(r.data))); } catch { sst.push({ text: '' }); } }

  // ---------- Sheets ----------
  const sheetNames = bundles.map((b) => b.name);
  const tables = new Map<number, { name: string; columns: string[] }>();
  const nameCtx = (row: number, col: number): DecodeCtx => ({
    xlsb: true, row, col,
    sheets: (ixti) => {
      const x = xti[ixti];
      if (!x || !supbooks[x.book]?.self || x.first < 0) return null;
      const a = sheetNames[x.first], b = sheetNames[x.last];
      return a ? { first: a, last: b } : null;
    },
    name: (idx) => {
      const n = rawNames[idx - 1];
      return n ? { name: n.name, sheet: n.itab !== 0xffffffff ? sheetNames[n.itab] : undefined } : undefined;
    },
    externName: (ixti, idx) => supbooks[xti[ixti]?.book ?? -1]?.names[idx - 1],
    table: (id) => tables.get(id),
  });
  const decode = (rgce: Uint8Array, rgcb: Uint8Array, row: number, col: number): string | Exp | undefined => {
    try { return decodeRgce(rgce, rgcb, nameCtx(row, col)); } catch { return undefined; }
  };

  // tables first, so structured references can be named
  const sheetParts = bundles.map((b) => wbRels.find((r) => r.id === b.rel)?.target);
  const sheetTables: TableDef[][] = sheetParts.map((p) => {
    if (!p) return [];
    const out: TableDef[] = [];
    for (const rel of rels(p)) {
      if (rel.type !== 'table') continue;
      const bin = get(rel.target);
      if (!bin) continue;
      let t: (TableDef & { id: number }) | null = null;
      for (const r of xlsbRecords(bin)) {
        const d = new Rd(r.data);
        try {
          if (r.type === 343) {
            const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32();
            d.u32(); const id = d.u32(); const hdr = d.u32(), tot = d.u32(); d.u32();
            d.skip(28);
            d.wide(); const display = d.wide();
            t = { id, name: display, range: { r1, r2, c1, c2 }, headerRow: hdr > 0, totalsRow: tot > 0, columns: [] };
          } else if (r.type === 347 && t) {
            d.skip(24);
            t.columns.push({ name: d.wide() });
          }
        } catch { /* skip */ }
      }
      if (t) { tables.set(t.id, { name: t.name, columns: t.columns.map((c) => c.name) }); const { id: _id, ...def } = t; void _id; out.push(def); }
    }
    return out;
  });

  const formulaAt = (d: Rd, base: CellRange): string | undefined => {
    const cce = d.u32(); const rgce = d.bytes(cce); const cb = d.u32(); const rgcb = d.bytes(cb);
    if (!cce) return undefined;
    const f = decode(rgce, rgcb, base.r1, base.c1);
    return typeof f === 'string' ? f : undefined;
  };
  const readCfRule = (d: Rd, base: CellRange): CfRule | null => {
    const type = d.u32(), sub = d.u32(), dxf = d.u32(); d.u32(); const param = d.u32(); d.u32(); d.u32();
    const fl = d.u16(); const n1 = d.u32(), n2 = d.u32(); d.u32();
    const text = d.wide();
    const f1 = n1 ? formulaAt(d, base) : undefined;
    const f2 = n2 ? formulaAt(d, base) : undefined;
    const style: CellStyle = dxfs[dxf] ? structuredClone(dxfs[dxf]!) : {};
    const stop = fl & 2 ? { stop: true } : {};
    if (type === 1) {
      const op = CF_OPS[param];
      if (!op || f1 === undefined) return null;
      return { type: 'cellIs', operator: op, formulas: f2 !== undefined ? [f1, f2] : [f1], style, ...stop };
    }
    if (type === 5) return { type: 'top10', rank: param, ...(fl & 8 ? { bottom: true } : {}), ...(fl & 16 ? { percent: true } : {}), style, ...stop };
    if (type !== 2) return null;
    switch (sub) {
      case 1: return f1 !== undefined ? { type: 'expression', formula: f1, style, ...stop } : null;
      case 7: return { type: 'uniqueValues', style, ...stop };
      case 27: return { type: 'duplicateValues', style, ...stop };
      case 8: return { type: (['containsText', 'notContainsText', 'beginsWith', 'endsWith'] as const)[param] ?? 'containsText', text, style, ...stop };
      case 9: return { type: 'containsBlanks', style, ...stop };
      case 10: return { type: 'notContainsBlanks', style, ...stop };
      case 11: return { type: 'containsErrors', style, ...stop };
      case 12: return { type: 'notContainsErrors', style, ...stop };
      case 25: case 26: case 29: case 30:
        return { type: 'aboveAverage', ...(sub === 26 || sub === 30 ? { below: true } : {}), ...(sub >= 29 ? { equal: true } : {}), ...(param ? { stdDev: param } : {}), style, ...stop };
      default: {
        const period = PERIODS[sub];
        return period ? { type: 'timePeriod', period, style, ...stop } : null;
      }
    }
  };
  const readDv = (d: Rd): DataValidation | null => {
    const fl = d.u32();
    const n = d.u32();
    const ranges: CellRange[] = [];
    for (let k = 0; k < n && d.left >= 16; k++) { const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32(); ranges.push({ r1, r2, c1, c2 }); }
    if (!ranges.length) return null;
    const errTitle = d.wide(), errText = d.wide(), promptTitle = d.wide(), promptText = d.wide();
    const base = ranges[0]!;
    const type = DV_TYPES[fl & 0xf] ?? 'any';
    const dv: DataValidation = { ranges, type };
    const f1 = formulaAt(d, base), f2 = formulaAt(d, base);
    if (f1 !== undefined) dv.formula1 = type === 'list' && /^".*"$/s.test(f1) ? f1.replace(/\u0000/g, ',') : f1;
    if (f2 !== undefined) dv.formula2 = f2;
    const op = DV_OPS[(fl >> 20) & 0xf];
    if (op && type !== 'any' && type !== 'list' && type !== 'custom') dv.operator = op;
    if (fl & 0x100) dv.allowBlank = true;
    if (type === 'list') dv.showDropdown = !(fl & 0x200);
    if (fl & 0x40000 && promptText) dv.prompt = { ...(promptTitle ? { title: promptTitle } : {}), text: promptText };
    if (fl & 0x80000) dv.error = { style: (['stop', 'warning', 'information'] as const)[(fl >> 4) & 7] ?? 'stop', ...(errTitle ? { title: errTitle } : {}), text: errText };
    return dv;
  };
  /** Drawing parts are the same XML as in .xlsx: charts, pictures and shapes. */
  const readDrawing = (path: string, sheet: Sheet): void => {
    const t = text(path);
    if (!t) return;
    let root: XNode;
    try { root = parseRoot(t); } catch { return; }
    const drels = new Map(rels(path).map((r) => [r.id, r]));
    const xml = (rid: string): XNode | undefined => {
      const r = drels.get(rid);
      const x = r && !r.external ? text(r.target) : undefined;
      try { return x ? parseRoot(x) : undefined; } catch { return undefined; }
    };
    for (const a of root.children) {
      if (!/Anchor$/.test(a.name)) continue;
      try {
        const anchor = readAnchor(a, sheet);
        const got = readAnchorChart(a, xml, `chart${sheet.charts.length + 1}`, sheet);
        if (got) { if (got.descr) got.chart.alt = got.descr; sheet.charts.push(got.chart); continue; }
        const pic = find(a, 'pic');
        const blip = pic && find(pic, 'blip');
        if (blip) {
          const rel = drels.get(attr(blip, 'embed') ?? '');
          const data = rel && !rel.external ? get(rel.target) : undefined;
          if (rel && data) {
            const ext = rel.target.slice(rel.target.lastIndexOf('.') + 1).toLowerCase();
            const p: Picture = { id: `picture${sheet.pictures.length + 1}`, data, mime: IMAGE_TYPES[ext] ?? 'application/octet-stream', anchor };
            const descr = attr(find(pic, 'cNvPr'), 'descr');
            if (descr) p.alt = descr;
            sheet.pictures.push(p);
          }
          continue;
        }
        const sp = kid(a, 'sp');
        if (sp) (sheet.shapes ??= []).push(readShape(sp, anchor, `shape${(sheet.shapes?.length ?? 0) + 1}`));
      } catch { /* skip a damaged anchor */ }
    }
  };

  bundles.forEach((b, i) => {
    const sheet = newSheet(b.name.slice(0, 31) || `Sheet${i + 1}`);
    sheet.state = b.state === 1 ? 'hidden' : b.state === 2 ? 'veryHidden' : 'visible';
    sheet.tables = sheetTables[i] ?? [];
    wb.sheets.push(sheet);
    const part = sheetParts[i];
    const bin = part ? get(part) : undefined;
    if (!bin || !part) return;
    const srels = rels(part);
    const shared: { range: CellRange; rgce: Uint8Array; rgcb: Uint8Array; array: boolean }[] = [];
    const exp: { row: number; col: number; at: Exp }[] = [];
    let row = 0;
    let lastCell: { row: number; col: number } | null = null;
    const put = (col: number, cell: Cell) => { sheet.cells.set(cellKey(row, col), cell); lastCell = { row, col }; };
    const ps: NonNullable<Sheet['pageSetup']> = {};
    let fit = false;
    let cf: ConditionalFormat | null = null;
    for (const r of xlsbRecords(bin)) {
      const d = new Rd(r.data);
      try {
        if (r.type >= 1 && r.type <= 11 || r.type === 62) {
          const col = d.u32();
          const ixfe = d.u32() & 0xffffff;
          const s = styleOf(ixfe);
          const cell: Cell = { v: null };
          if (s) cell.s = s;
          let fml = false;
          switch (r.type) {
            case 1: if (!s) continue; break;
            case 2: cell.v = rk(d.u32()); break;
            case 3: cell.v = new CellError(ERR_BY_CODE[d.u8()] ?? '#N/A'); break;
            case 4: cell.v = !!d.u8(); break;
            case 5: cell.v = d.f64(); break;
            case 6: cell.v = d.wide(); break;
            case 7: { const e = sst[d.u32()]; cell.v = e?.text ?? ''; if (e?.rich) cell.rich = e.rich; break; }
            case 62: { const e = richStr(d); cell.v = e.text; if (e.rich) cell.rich = e.rich; break; }
            case 8: cell.v = d.wide(); fml = true; break;
            case 9: cell.v = d.f64(); fml = true; break;
            case 10: cell.v = !!d.u8(); fml = true; break;
            case 11: cell.v = new CellError(ERR_BY_CODE[d.u8()] ?? '#N/A'); fml = true; break;
          }
          if (fml) {
            d.u16();
            const cce = d.u32(); const rgce = d.bytes(cce); const cb = d.u32(); const rgcb = d.bytes(cb);
            const f = decode(rgce, rgcb, row, col);
            if (f instanceof Exp) exp.push({ row, col, at: f });
            else if (f) cell.f = f;
          }
          put(col, cell);
          continue;
        }
        switch (r.type) {
          case 0: {
            row = d.u32();
            const ixfe = d.u32(); const h = d.u16(); const fl = d.u16();
            const info: NonNullable<ReturnType<Sheet['rows']['get']>> = {};
            if (fl & 0x2000) info.height = h / 20;
            if (fl & 0x1000) info.hidden = true;
            const lvl = (fl >> 8) & 7;
            if (lvl) info.outline = lvl;
            if (fl & 0x4000) { const s = styleOf(ixfe); if (s) info.s = s; }
            if (Object.keys(info).length) sheet.rows.set(row, info);
            break;
          }
          case 425: case 426: {
            const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32();
            if (r.type === 426) d.u8();
            const cce = d.u32(); const rgce = d.bytes(cce); const cb = d.u32(); const rgcb = d.bytes(cb);
            shared.push({ range: { r1, r2, c1, c2 }, rgce, rgcb, array: r.type === 426 });
            void lastCell;
            break;
          }
          case 60: {
            const c1 = d.u32(), c2 = d.u32(), w = d.u32(), ixfe = d.u32(), fl = d.u16();
            const info: NonNullable<ReturnType<Sheet['cols']['get']>> = { width: shownColWidth(w / 256) };
            if (fl & 1) info.hidden = true;
            const lvl = (fl >> 8) & 7;
            if (lvl) info.outline = lvl;
            const s = styleOf(ixfe);
            if (s) info.s = s;
            for (let c = c1; c <= Math.min(c2, 16383) && c - c1 < 16384; c++) sheet.cols.set(c, { ...info });
            break;
          }
          case 485: {
            const dx = d.u32(); const cch = d.u16(); const h = d.u16();
            if (dx !== 0xffffffff && dx) sheet.defaultColWidth = shownColWidth(dx / 256); else if (cch) sheet.defaultColWidth = cch + 0.43;
            if (h) sheet.defaultRowHeight = h / 20;
            break;
          }
          case 147: {
            const fl = d.u16(); d.u8();
            if (fl & 0x100) fit = true;
            const c = color(d);
            const raw = r.data[3]! >> 1;
            if (c && raw !== 0) sheet.tabColor = c;
            break;
          }
          case 137: {
            const fl = d.u16(); const xlView = d.u32(); d.u32(); d.u32(); d.u32(); const scale = d.u16();
            if (xlView === 2) sheet.view.pageView = 'pageLayout'; else if (xlView === 1) sheet.view.pageView = 'pageBreakPreview';
            sheet.view.showGrid = !!(fl & 4);
            sheet.view.showHeaders = !!(fl & 8);
            sheet.view.rtl = !!(fl & 0x20);
            if (scale) sheet.view.zoom = Math.max(10, Math.min(400, scale));
            break;
          }
          case 151: {
            const x = d.f64(), y = d.f64(); d.u32(); d.u32(); d.u32(); const fl = d.u8();
            if (fl & 1) { sheet.view.frozenCols = Math.round(x); sheet.view.frozenRows = Math.round(y); }
            break;
          }
          case 152: {
            d.u32(); const ra = d.u32(), ca = d.u32(); d.u32(); const n = d.u32();
            if (n) { const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32(); sheet.view.selection = { r1, r2, c1, c2 }; }
            else sheet.view.selection = { r1: ra, r2: ra, c1: ca, c2: ca };
            break;
          }
          case 176: { const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32(); sheet.merges.push({ r1, r2, c1, c2 }); break; }
          case 494: {
            const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32();
            const relId = d.wide(), loc = d.wide();
            const rel = relId ? srels.find((x) => x.id === relId) : undefined;
            let target = rel ? rel.target : '';
            if (loc) target = rel ? `${target}#${loc}` : `#${loc}`;
            if (!target || !SAFE_LINK.test(target)) break;
            let n = 0;
            for (let a = r1; a <= r2 && n < 10000; a++) for (let b2 = c1; b2 <= c2 && n < 10000; b2++, n++) sheet.hyperlinks.set(cellKey(a, b2), target);
            break;
          }
          case 535: {
            const pw = d.u16(); const locked = d.u32();
            if (!locked) break;
            const keys = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
            const options: Record<string, boolean> = {};
            for (const k of keys) if (d.left >= 4) options[k] = !d.u32();
            sheet.protection = { sheet: true, ...(pw ? { password: pw.toString(16).toUpperCase().padStart(4, '0') } : {}), options };
            break;
          }
          case 476: ps.margins = { left: d.f64(), right: d.f64(), top: d.f64(), bottom: d.f64(), header: d.f64(), footer: d.f64() }; break;
          case 478: {
            const paper = d.u32(), scale = d.u32(); d.u32(); d.u32(); d.u32(); const start = d.u32(); const fw = d.u32(), fh = d.u32(); const fl = d.u16();
            if (paper) ps.paper = paper;
            if (scale && scale !== 100) ps.scale = scale;
            ps.orientation = fl & 2 ? 'landscape' : 'portrait';
            if (fl & 0x80 && start !== 1) ps.firstPageNumber = start;
            if (fit) { ps.fitWidth = fw; ps.fitHeight = fh; }
            break;
          }
          case 477: { const fl = d.u16(); if (fl & 1) ps.centerH = true; if (fl & 2) ps.centerV = true; if (fl & 4) ps.headings = true; if (fl & 8) ps.gridlines = true; break; }
          case 479: { d.u16(); const h = d.wide(), f = d.wide(); if (h) ps.header = h; if (f) ps.footer = f; break; }
          case 161: { const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32(); sheet.autoFilter = { range: { r1, r2, c1, c2 }, columns: [] }; break; }
          case 461: {
            d.u32(); d.u32();
            const n = d.u32();
            const ranges: CellRange[] = [];
            for (let k = 0; k < n && d.left >= 16; k++) { const r1 = d.u32(), r2 = d.u32(), c1 = d.u32(), c2 = d.u32(); ranges.push({ r1, r2, c1, c2 }); }
            cf = ranges.length ? { ranges, rules: [] } : null;
            if (cf) sheet.conditionalFormats.push(cf);
            break;
          }
          case 462: cf = null; break;
          case 463: {
            if (!cf) break;
            const rule = readCfRule(d, cf.ranges[0]!);
            if (rule) cf.rules.push(rule);
            break;
          }
          case 64: { const dv = readDv(d); if (dv) sheet.validations.push(dv); break; }
          case 550: { const id = d.wide(); const rel = srels.find((x) => x.id === id); if (rel && !rel.external) readDrawing(rel.target, sheet); break; }
          default: break;
        }
      } catch { /* skip damaged records */ }
    }
    for (const e of exp) {
      const def = shared.find((s) => s.range.r1 === e.at.row && e.row >= s.range.r1 && e.row <= s.range.r2 && e.col >= s.range.c1 && e.col <= s.range.c2);
      const cell = sheet.cells.get(cellKey(e.row, e.col));
      if (!def || !cell) continue;
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
    if (Object.keys(ps).length) sheet.pageSetup = ps;
    // notes
    for (const rel of srels) {
      if (rel.type !== 'comments') continue;
      const cb = get(rel.target);
      if (!cb) continue;
      const authors: string[] = [];
      let at: { row: number; col: number; author?: string } | null = null;
      for (const r of xlsbRecords(cb)) {
        const d = new Rd(r.data);
        try {
          if (r.type === 632) authors.push(d.wide());
          else if (r.type === 635) { const iau = d.u32(); const r1 = d.u32(); d.u32(); const c1 = d.u32(); at = { row: r1, col: c1, author: authors[iau] }; }
          else if (r.type === 637 && at) {
            const t = richStr(d).text.replace(/\r\n?/g, '\n');
            sheet.notes.set(cellKey(at.row, at.col), at.author ? { author: at.author, text: t } : { text: t });
          }
        } catch { /* skip */ }
      }
    }
  });

  // names
  rawNames.forEach((n) => {
    if (n.flags & 0x0e) return;
    const local = n.itab !== 0xffffffff ? n.itab : undefined;
    const f = decode(n.rgce, n.rgcb, 0, 0);
    if (typeof f !== 'string') return;
    if (/^_xlnm\._FilterDatabase$/i.test(n.name)) return;
    if (/^_xlnm\.Print_Area$/i.test(n.name) && local !== undefined && wb.sheets[local]) {
      const sh = wb.sheets[local]!;
      sh.pageSetup = { ...sh.pageSetup, printArea: f.split(',').map((p) => p.slice(p.lastIndexOf('!') + 1).replace(/\$/g, '')).join(',') };
      return;
    }
    const d: Workbook['names'][number] = { name: n.name, formula: f };
    if (local !== undefined) d.sheet = local;
    if (n.flags & 1) d.hidden = true;
    wb.names.push(d);
  });

  if (!wb.sheets.length) throw new Error('This workbook has no worksheets.');
  const own = (type: string) => { const t = wbRels.find((r) => r.type === type && !r.external)?.target; return t ? text(t) : undefined; };
  applyObjectsPart(wb, own(OBJECTS_REL_SHORT));
  // custom XML items (such as saved data queries) are kept for the next save
  for (const p of Object.keys(files)) {
    if (/^customXml\/(item\d+\.xml|itemProps\d+\.xml|_rels\/item\d+\.xml\.rels)$/i.test(p)) (wb.extraParts ??= new Map()).set(p, files[p]!);
  }
  readExtras(wb, own('binaryExtras'));
  wb.activeSheet = Math.min(activeTab, wb.sheets.length - 1);
  if (Object.keys(calc).length) wb.calc = calc;
  return wb;
};

const CF_OPS: Record<number, CfOperator> = { 1: 'between', 2: 'notBetween', 3: 'equal', 4: 'notEqual', 5: 'greaterThan', 6: 'lessThan', 7: 'greaterThanOrEqual', 8: 'lessThanOrEqual' };
const DV_OPS: Record<number, CfOperator> = { 0: 'between', 1: 'notBetween', 2: 'equal', 3: 'notEqual', 4: 'greaterThan', 5: 'lessThan', 6: 'greaterThanOrEqual', 7: 'lessThanOrEqual' };
const DV_TYPES: DataValidation['type'][] = ['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'];
const PERIODS: Record<number, Extract<CfRule, { type: 'timePeriod' }>['period']> = {
  15: 'today', 16: 'tomorrow', 17: 'yesterday', 18: 'last7Days', 19: 'lastMonth', 20: 'nextMonth', 21: 'thisWeek', 22: 'nextWeek', 23: 'lastWeek', 24: 'thisMonth',
};
const IMAGE_TYPES: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp' };

export type { Scalar };
