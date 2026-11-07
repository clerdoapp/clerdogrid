// .xlsx reader (ECMA-376 Part 1 SpreadsheetML, Part 2 OPC packaging).
import { unzipSync } from 'fflate';
import {
  cellKey, newSheet, parseRange,
  type Cell, type CellRange, type CellStyle, type CfOperator, type CfRule, type CfValue, type ConditionalFormat,
  type DataValidation, type DefinedName, type FilterColumn, type Note, type Picture, type RichRun, type Sheet, type TableDef, type Workbook,
} from '../model/workbook';
import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../engine/types';
import { toSerial } from '../engine/date';
import { attr, boolAttr, decodeXstr, deepText, find, findAll, kid, kids, numAttr, parseRoot, type XNode } from './xml';
import { readColor, readFont, readStyles, readTheme } from './xlsx-styles';
import { INDEXED_COLORS } from './numfmt';
import { readAnchor, readChart, shownColWidth } from './xlsx-chart';
import { shiftFormula, stripFuturePrefixes } from './formula-xml';
import { OBJECTS_REL_SHORT, applyObjectsPart, readProtectedRanges, readShape } from './xlsx-objects';

const MAX_PART = 1 << 30;
const MAX_TOTAL = 2 ** 31;

interface Rel { id: string; type: string; target: string; external: boolean }

const relType = (t: string): string => t.slice(t.lastIndexOf('/') + 1);

const dirOf = (p: string): string => p.slice(0, p.lastIndexOf('/') + 1);
const resolve = (base: string, target: string): string => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
};
const relsPath = (part: string): string => `${dirOf(part)}_rels/${part.slice(part.lastIndexOf('/') + 1)}.rels`;

const SAFE_LINK = /^(https:|mailto:|#)/i;

/** Formula text from the file: strips future-function prefixes and turns ANCHORARRAY(A1) into A1#. */
const fx = (f: string): string =>
  stripFuturePrefixes(f).replace(/\bANCHORARRAY\(((?:'[^']+'|[A-Za-z0-9_.]+!)?\$?[A-Za-z]{1,3}\$?\d+)\)/g, '$1#');

const ranges = (sqref: string | undefined): CellRange[] =>
  (sqref ?? '').split(/\s+/).map((r) => parseRange(r)).filter((r): r is CellRange => !!r);

class Package {
  files: Record<string, Uint8Array>;
  private dec = new TextDecoder('utf-8');
  private dec16 = new TextDecoder('utf-16le');
  private lower = new Map<string, string>();

  constructor(data: Uint8Array) {
    let total = 0;
    this.files = unzipSync(data, {
      filter: (f) => {
        total += f.originalSize;
        if (f.originalSize > MAX_PART || total > MAX_TOTAL) throw new Error('The file is too large to open.');
        return true;
      },
    });
    for (const k of Object.keys(this.files)) this.lower.set(k.toLowerCase(), k);
  }

  name(p: string): string | undefined {
    return this.files[p] ? p : this.lower.get(p.toLowerCase());
  }
  bytes(p: string): Uint8Array | undefined {
    const n = this.name(p);
    return n ? this.files[n] : undefined;
  }
  text(p: string): string | undefined {
    const b = this.bytes(p);
    if (!b) return undefined;
    if (b[0] === 0xff && b[1] === 0xfe) return this.dec16.decode(b.subarray(2));
    return this.dec.decode(b);
  }
  xml(p: string): XNode | undefined {
    const t = this.text(p);
    return t === undefined ? undefined : parseRoot(t);
  }
  rels(part: string): Rel[] {
    const root = this.xml(relsPath(part));
    return kids(root, 'Relationship').map((r) => {
      const external = attr(r, 'TargetMode') === 'External';
      const target = attr(r, 'Target') ?? '';
      return { id: attr(r, 'Id') ?? '', type: relType(attr(r, 'Type') ?? ''), target: external ? target : resolve(part, target), external };
    });
  }
}

const readRich = (si: XNode, theme: readonly string[]): { text: string; rich?: RichRun[] } => {
  const runs = kids(si, 'r');
  if (!runs.length) return { text: decodeXstr(kid(si, 't')?.text ?? '') };
  const rich: RichRun[] = runs.map((r) => {
    const run: RichRun = { text: decodeXstr(kid(r, 't')?.text ?? '') };
    const rp = kid(r, 'rPr');
    const font = rp && readFont(rp, theme, INDEXED_COLORS, false);
    if (font) run.font = font;
    return run;
  });
  return { text: rich.map((r) => r.text).join(''), rich };
};

const errorOf = (s: string): CellError => new CellError((ERROR_CODES.includes(s as ErrorCode) ? s : '#VALUE!') as ErrorCode);

const isoToSerial = (s: string, date1904: boolean): number | string => {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s);
  if (!m) {
    const t = /^T?(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(s);
    return t ? (+t[1]! * 3600 + +t[2]! * 60 + +(t[3] ?? 0)) / 86400 : s;
  }
  const day = toSerial(+m[1]!, +m[2]!, +m[3]!, date1904);
  return day + ((+(m[4] ?? 0)) * 3600 + (+(m[5] ?? 0)) * 60 + +(m[6] ?? 0)) / 86400;
};

const parseRef = (r: string): { row: number; col: number } | null => {
  const m = /^([A-Z]{1,3})(\d+)$/i.exec(r);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: +m[2]! - 1, col: col - 1 };
};

interface Ctx {
  pkg: Package;
  theme: string[];
  xfs: CellStyle[];
  dxfs: CellStyle[];
  strings: { text: string; rich?: RichRun[] }[];
  date1904: boolean;
  dynamicCm: Set<string>;
}

const cfValue = (n: XNode): CfValue => {
  const v: CfValue = { type: (attr(n, 'type') ?? 'num') as CfValue['type'] };
  const val = attr(n, 'val') ?? kid(n, 'f')?.text;
  if (val !== undefined) v.value = fx(val);
  return v;
};

const readCfRule = (r: XNode, ctx: Ctx, dxfOverride?: CellStyle): CfRule | undefined => {
  const type = attr(r, 'type');
  const style: CellStyle = dxfOverride ?? ctx.dxfs[numAttr(r, 'dxfId') ?? -1] ?? {};
  const stop = boolAttr(r, 'stopIfTrue') || undefined;
  const withStop = <T extends object>(o: T): T => (stop ? { ...o, stop } : o);
  const formulas = kids(r, 'formula').concat(kids(r, 'f')).map((f) => fx(f.text));
  switch (type) {
    case 'cellIs':
      return withStop({ type, operator: (attr(r, 'operator') ?? 'equal') as CfOperator, formulas, style });
    case 'expression':
      return withStop({ type, formula: formulas[0] ?? 'FALSE', style });
    case 'colorScale': {
      const cs = kid(r, 'colorScale');
      const vals = kids(cs, 'cfvo').map(cfValue);
      const cols = kids(cs, 'color').map((c) => readColor(c, ctx.theme) ?? '#000000');
      return { type, stops: vals.map((value, i) => ({ value, color: cols[i] ?? '#000000' })) };
    }
    case 'dataBar': {
      const db = kid(r, 'dataBar');
      const vals = kids(db, 'cfvo').map(cfValue);
      const rule: CfRule = { type, min: vals[0] ?? { type: 'min' }, max: vals[1] ?? { type: 'max' }, color: readColor(kid(db, 'color') ?? kid(db, 'fillColor'), ctx.theme) ?? '#638EC6' };
      if (attr(db, 'showValue') === '0') rule.showValue = false;
      if (attr(db, 'gradient') === '0') rule.gradient = false;
      return rule;
    }
    case 'iconSet': {
      const is = kid(r, 'iconSet');
      const rule: CfRule = { type, set: attr(is, 'iconSet') ?? '3TrafficLights1', values: kids(is, 'cfvo').map(cfValue) };
      if (boolAttr(is, 'reverse')) rule.reverse = true;
      if (attr(is, 'showValue') === '0') rule.showValue = false;
      return rule;
    }
    case 'top10': {
      const rule: CfRule = { type, rank: numAttr(r, 'rank') ?? 10, style };
      if (boolAttr(r, 'bottom')) rule.bottom = true;
      if (boolAttr(r, 'percent')) rule.percent = true;
      return withStop(rule);
    }
    case 'aboveAverage': {
      const rule: CfRule = { type, style };
      if (attr(r, 'aboveAverage') === '0') rule.below = true;
      if (boolAttr(r, 'equalAverage')) rule.equal = true;
      const sd = numAttr(r, 'stdDev');
      if (sd) rule.stdDev = sd;
      return withStop(rule);
    }
    case 'duplicateValues': case 'uniqueValues':
    case 'containsBlanks': case 'notContainsBlanks': case 'containsErrors': case 'notContainsErrors':
      return withStop({ type, style });
    case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith':
      return withStop({ type, text: attr(r, 'text') ?? '', style });
    case 'timePeriod':
      return withStop({ type, period: (attr(r, 'timePeriod') ?? 'today') as Extract<CfRule, { type: 'timePeriod' }>['period'], style });
    default:
      return undefined;
  }
};

const readConditionalFormats = (ws: XNode, ctx: Ctx): ConditionalFormat[] => {
  const out: ConditionalFormat[] = [];
  const linked = new Set<string>();
  for (const cf of kids(ws, 'conditionalFormatting')) {
    const rules: CfRule[] = [];
    let priority = Infinity;
    for (const r of kids(cf, 'cfRule')) {
      const rule = readCfRule(r, ctx);
      const id = find(r, 'id')?.text;
      if (id) linked.add(id.trim());
      if (rule) rules.push(rule);
      priority = Math.min(priority, numAttr(r, 'priority') ?? Infinity);
    }
    const rs = ranges(attr(cf, 'sqref'));
    if (rules.length && rs.length) out.push(Number.isFinite(priority) ? { ranges: rs, rules, priority } : { ranges: rs, rules });
  }
  // x14 rules that aren't extensions of a rule above (e.g. data bars or icon sets only stored there).
  for (const cf of findAll(kid(ws, 'extLst'), 'conditionalFormatting')) {
    const rules: CfRule[] = [];
    let priority = Infinity;
    for (const r of kids(cf, 'cfRule')) {
      if (linked.has(attr(r, 'id') ?? '')) continue;
      const dxfNode = kid(r, 'dxf');
      const style = dxfNode ? dxfFromNode(dxfNode, ctx) : undefined;
      const rule = readCfRule(r, ctx, style);
      if (rule) rules.push(rule);
      priority = Math.min(priority, numAttr(r, 'priority') ?? Infinity);
    }
    const rs = ranges(kid(cf, 'sqref')?.text);
    if (rules.length && rs.length) out.push(Number.isFinite(priority) ? { ranges: rs, rules, priority } : { ranges: rs, rules });
  }
  return out.sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
};

const dxfFromNode = (n: XNode, ctx: Ctx): CellStyle => {
  const font = kid(n, 'font');
  const fill = kid(n, 'fill');
  const s: CellStyle = {};
  if (font) { const f = readFont(font, ctx.theme, INDEXED_COLORS, false); if (f) s.font = f; }
  if (fill) {
    const pf = kid(fill, 'patternFill');
    const c = readColor(kid(pf, 'bgColor') ?? kid(pf, 'fgColor'), ctx.theme);
    if (c) s.fill = { color: c };
  }
  return s;
};

const readValidations = (ws: XNode): DataValidation[] => {
  const nodes = kids(kid(ws, 'dataValidations'), 'dataValidation').concat(findAll(kid(ws, 'extLst'), 'dataValidation'));
  const out: DataValidation[] = [];
  for (const d of nodes) {
    const rs = ranges(attr(d, 'sqref') ?? kid(d, 'sqref')?.text);
    if (!rs.length) continue;
    const t = attr(d, 'type') ?? 'none';
    const v: DataValidation = { ranges: rs, type: (t === 'none' ? 'any' : t) as DataValidation['type'] };
    const op = attr(d, 'operator') ?? (/^(whole|decimal|date|time|textLength)$/.test(v.type) ? 'between' : undefined);
    if (op) v.operator = op as CfOperator;
    const f1 = kid(d, 'formula1'), f2 = kid(d, 'formula2');
    const ftext = (n: XNode): string => fx(kid(n, 'f')?.text ?? n.text);
    if (f1) v.formula1 = ftext(f1);
    if (f2) v.formula2 = ftext(f2);
    if (boolAttr(d, 'allowBlank')) v.allowBlank = true;
    if (v.type === 'list') v.showDropdown = !boolAttr(d, 'showDropDown');
    const prompt = attr(d, 'prompt');
    if (boolAttr(d, 'showInputMessage') && (prompt || attr(d, 'promptTitle'))) {
      v.prompt = { text: prompt ?? '' };
      const pt = attr(d, 'promptTitle');
      if (pt) v.prompt.title = pt;
    }
    if (boolAttr(d, 'showErrorMessage')) {
      v.error = { style: (attr(d, 'errorStyle') ?? 'stop') as 'stop' | 'warning' | 'information', text: attr(d, 'error') ?? '' };
      const et = attr(d, 'errorTitle');
      if (et) v.error.title = et;
    }
    out.push(v);
  }
  return out;
};

const readAutoFilter = (af: XNode, ctx: Ctx, sortNode?: XNode): Sheet['autoFilter'] => {
  const range = parseRange(attr(af, 'ref') ?? '');
  if (!range) return undefined;
  const columns: FilterColumn[] = [];
  for (const fc of kids(af, 'filterColumn')) {
    const c: FilterColumn = { col: numAttr(fc, 'colId') ?? 0 };
    const filters = kid(fc, 'filters');
    if (filters) {
      c.values = kids(filters, 'filter').map((f) => attr(f, 'val') ?? '');
      if (boolAttr(filters, 'blank')) c.blanks = true;
    }
    const custom = kid(fc, 'customFilters');
    if (custom) {
      c.custom = { rules: kids(custom, 'customFilter').map((f) => ({ operator: (attr(f, 'operator') ?? 'equal') as CfOperator, value: attr(f, 'val') ?? '' })) };
      if (boolAttr(custom, 'and')) c.custom.and = true;
    }
    const top = kid(fc, 'top10');
    if (top) c.top10 = { top: attr(top, 'top') !== '0', percent: boolAttr(top, 'percent'), value: numAttr(top, 'val') ?? 10 };
    const color = kid(fc, 'colorFilter');
    if (color) {
      const dxf = ctx.dxfs[numAttr(color, 'dxfId') ?? -1];
      const col = dxf?.fill?.color ?? dxf?.font?.color;
      if (col) c.color = col;
    }
    columns.push(c);
  }
  const out: NonNullable<Sheet['autoFilter']> = { range, columns };
  const ss = kid(af, 'sortState') ?? sortNode;
  if (ss) {
    const sort = kids(ss, 'sortCondition').map((s) => {
      const r = parseRange(attr(s, 'ref') ?? '');
      const o: { col: number; descending?: boolean } = { col: r ? r.c1 - range.c1 : 0 };
      if (boolAttr(s, 'descending')) o.descending = true;
      return o;
    });
    if (sort.length) out.sort = sort;
  }
  return out;
};

const readTable = (t: XNode): TableDef | undefined => {
  const range = parseRange(attr(t, 'ref') ?? '');
  if (!range) return undefined;
  const def: TableDef = {
    name: attr(t, 'displayName') ?? attr(t, 'name') ?? 'Table',
    range,
    headerRow: attr(t, 'headerRowCount') !== '0',
    totalsRow: (numAttr(t, 'totalsRowCount') ?? 0) > 0,
    columns: kids(kid(t, 'tableColumns'), 'tableColumn').map((c) => {
      const col: TableDef['columns'][number] = { name: decodeXstr(attr(c, 'name') ?? '') };
      const fn = attr(c, 'totalsRowFunction');
      if (fn && fn !== 'none') col.totalsFunction = fn === 'custom' ? fx(kid(c, 'totalsRowFormula')?.text ?? '') : fn;
      const calc = kid(c, 'calculatedColumnFormula');
      if (calc) col.calculatedFormula = fx(calc.text);
      return col;
    }),
  };
  const si = kid(t, 'tableStyleInfo');
  if (si) {
    const name = attr(si, 'name');
    if (name) def.style = name;
    def.bandedRows = boolAttr(si, 'showRowStripes');
    def.bandedCols = boolAttr(si, 'showColumnStripes');
  }
  return def;
};

const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', webp: 'image/webp' };

const readDrawing = (path: string, sheet: Sheet, ctx: Ctx): void => {
  const root = ctx.pkg.xml(path);
  if (!root) return;
  const rels = new Map(ctx.pkg.rels(path).map((r) => [r.id, r]));
  for (const a of root.children) {
    if (!/Anchor$/.test(a.name)) continue;
    const anchor = readAnchor(a, sheet);
    const frame = find(a, 'graphicFrame');
    const chartRef = frame && find(frame, 'chart');
    if (chartRef) {
      const rel = rels.get(attr(chartRef, 'id') ?? '');
      const cx = rel && !rel.external ? ctx.pkg.xml(rel.target) : undefined;
      const chart = cx && readChart(cx, `chart${sheet.charts.length + 1}`, anchor);
      const descr = attr(find(frame, 'cNvPr'), 'descr');
      if (chart && descr) chart.alt = descr;
      if (chart) sheet.charts.push(chart);
      continue;
    }
    const pic = find(a, 'pic');
    const blip = pic && find(pic, 'blip');
    if (blip) {
      const rel = rels.get(attr(blip, 'embed') ?? '');
      const data = rel && !rel.external ? ctx.pkg.bytes(rel.target) : undefined;
      if (rel && data) {
        const ext = rel.target.slice(rel.target.lastIndexOf('.') + 1).toLowerCase();
        const p: Picture = { id: `picture${sheet.pictures.length + 1}`, data, mime: MIME[ext] ?? 'application/octet-stream', anchor };
        const descr = attr(find(pic, 'cNvPr'), 'descr');
        if (descr) p.alt = descr;
        sheet.pictures.push(p);
      }
    }
    const sp = kid(a, 'sp');
    if (sp) (sheet.shapes ??= []).push(readShape(sp, anchor, `shape${(sheet.shapes?.length ?? 0) + 1}`));
  }
};

const readComments = (path: string, sheet: Sheet, ctx: Ctx): void => {
  const root = ctx.pkg.xml(path);
  if (!root) return;
  const authors = kids(kid(root, 'authors'), 'author').map((a) => a.text);
  for (const c of kids(kid(root, 'commentList'), 'comment')) {
    const p = parseRef(attr(c, 'ref') ?? '');
    if (!p) continue;
    const textNode = kid(c, 'text');
    const text = textNode ? decodeXstr(kids(textNode, 'r').length ? kids(textNode, 'r').map((r) => kid(r, 't')?.text ?? '').join('') : kid(textNode, 't')?.text ?? '') : '';
    const note: Note = { text };
    const author = authors[numAttr(c, 'authorId') ?? -1];
    if (author) note.author = author;
    sheet.notes.set(cellKey(p.row, p.col), note);
  }
};

const readSheet = (path: string, name: string, ctx: Ctx): Sheet => {
  const sheet = newSheet(name);
  const ws = ctx.pkg.xml(path);
  if (!ws) return sheet;
  const rels = new Map(ctx.pkg.rels(path).map((r) => [r.id, r]));

  const pr = kid(ws, 'sheetPr');
  const tab = readColor(kid(pr, 'tabColor'), ctx.theme);
  if (tab) sheet.tabColor = tab;

  const fp = kid(ws, 'sheetFormatPr');
  const dch = numAttr(fp, 'defaultColWidth');
  const base = numAttr(fp, 'baseColWidth');
  if (dch) sheet.defaultColWidth = shownColWidth(dch);
  else if (base !== undefined && base !== 8) sheet.defaultColWidth = Math.round((base + 0.43) * 100) / 100;
  const drh = numAttr(fp, 'defaultRowHeight');
  if (drh) sheet.defaultRowHeight = drh;

  const view = kid(kid(ws, 'sheetViews'), 'sheetView');
  if (view) {
    sheet.view.showGrid = attr(view, 'showGridLines') !== '0';
    sheet.view.showHeaders = attr(view, 'showRowColHeaders') !== '0';
    sheet.view.zoom = numAttr(view, 'zoomScale') || 100;
    sheet.view.rtl = boolAttr(view, 'rightToLeft');
    const pane = kid(view, 'pane');
    if (pane && (attr(pane, 'state') === 'frozen' || attr(pane, 'state') === 'frozenSplit')) {
      sheet.view.frozenCols = Math.round(numAttr(pane, 'xSplit') ?? 0);
      sheet.view.frozenRows = Math.round(numAttr(pane, 'ySplit') ?? 0);
    }
    const sels = kids(view, 'selection');
    const sel = sels[sels.length - 1];
    const r = sel && (parseRange((attr(sel, 'sqref') ?? '').split(' ')[0] ?? '') ?? parseRange(attr(sel, 'activeCell') ?? ''));
    if (r) sheet.view.selection = r;
  }

  for (const c of kids(kid(ws, 'cols'), 'col')) {
    const min = (numAttr(c, 'min') ?? 1) - 1, max = Math.min((numAttr(c, 'max') ?? 1) - 1, 16383);
    const info: { width?: number; hidden?: boolean; s?: number; outline?: number } = {};
    const w = numAttr(c, 'width');
    if (w !== undefined && (boolAttr(c, 'customWidth') || Math.abs(shownColWidth(w) - sheet.defaultColWidth) > 0.01)) info.width = shownColWidth(w);
    if (boolAttr(c, 'hidden')) info.hidden = true;
    const s = numAttr(c, 'style');
    if (s) info.s = s;
    const o = numAttr(c, 'outlineLevel');
    if (o) info.outline = o;
    if (!Object.keys(info).length) continue;
    for (let i = min; i <= max; i++) sheet.cols.set(i, { ...info });
  }

  const shared = new Map<string, { f: string; row: number; col: number }>();
  const spills: CellRange[] = [];
  let rowIdx = -1;
  for (const row of kids(kid(ws, 'sheetData'), 'row')) {
    rowIdx = (numAttr(row, 'r') ?? rowIdx + 2) - 1;
    const info: { height?: number; hidden?: boolean; s?: number; outline?: number } = {};
    const ht = numAttr(row, 'ht');
    if (ht !== undefined && boolAttr(row, 'customHeight')) info.height = ht;
    if (boolAttr(row, 'hidden')) info.hidden = true;
    if (boolAttr(row, 'customFormat')) { const s = numAttr(row, 's'); if (s) info.s = s; }
    const o = numAttr(row, 'outlineLevel');
    if (o) info.outline = o;
    if (Object.keys(info).length) sheet.rows.set(rowIdx, info);
    let colIdx = -1;
    for (const c of kids(row, 'c')) {
      const p = parseRef(attr(c, 'r') ?? '');
      if (p) { colIdx = p.col; } else colIdx++;
      const r = p ? p.row : rowIdx;
      const t = attr(c, 't') ?? 'n';
      const vNode = kid(c, 'v');
      const raw = vNode?.text;
      let v: Scalar = null;
      let rich: RichRun[] | undefined;
      if (t === 'inlineStr') { const is = kid(c, 'is'); if (is) { const x = readRich(is, ctx.theme); v = x.text; rich = x.rich; } }
      else if (raw !== undefined) {
        switch (t) {
          case 's': { const x = ctx.strings[+raw]; v = x?.text ?? ''; rich = x?.rich; break; }
          case 'str': v = decodeXstr(raw); break;
          case 'b': v = raw === '1' || raw === 'true'; break;
          case 'e': v = errorOf(raw.trim()); break;
          case 'd': v = isoToSerial(raw, ctx.date1904); break;
          default: { const n = Number(raw); v = raw.trim() !== '' && Number.isFinite(n) ? n : raw; }
        }
      }
      const cell: Cell = { v };
      if (rich) cell.rich = rich;
      const s = numAttr(c, 's');
      if (s && s < ctx.xfs.length) cell.s = s;
      const f = kid(c, 'f');
      if (f) {
        const ft = attr(f, 't');
        const text = f.text;
        if (ft === 'shared') {
          const si = attr(f, 'si') ?? '';
          if (text.trim()) { shared.set(si, { f: fx(text), row: r, col: colIdx }); cell.f = fx(text); }
          else {
            const m = shared.get(si);
            if (m) cell.f = shiftFormula(m.f, r - m.row, colIdx - m.col);
          }
        } else if (text.trim()) cell.f = fx(text);
        if (ft === 'array' && cell.f) {
          const ref = parseRange(attr(f, 'ref') ?? '') ?? { r1: r, c1: colIdx, r2: r, c2: colIdx };
          const cm = attr(c, 'cm');
          if (cm !== undefined && ctx.dynamicCm.has(cm)) { cell.dyn = true; spills.push(ref); }
          else cell.arrayRange = ref;
        }
      }
      if (v === null && !cell.f && !cell.s) continue;
      sheet.cells.set(cellKey(r, colIdx), cell);
    }
  }
  // A dynamic array's spilled cells are its results, recalculated on load, not cell contents.
  for (const sp of spills) {
    for (let r = sp.r1; r <= sp.r2; r++) {
      for (let c = sp.c1; c <= sp.c2; c++) {
        if (r === sp.r1 && c === sp.c1) continue;
        const k = cellKey(r, c);
        const cell = sheet.cells.get(k);
        if (!cell || cell.f) continue;
        if (cell.s) sheet.cells.set(k, { v: null, s: cell.s });
        else sheet.cells.delete(k);
      }
    }
  }

  for (const m of kids(kid(ws, 'mergeCells'), 'mergeCell')) {
    const r = parseRange(attr(m, 'ref') ?? '');
    if (r && (r.r1 !== r.r2 || r.c1 !== r.c2)) sheet.merges.push(r);
  }

  sheet.conditionalFormats = readConditionalFormats(ws, ctx);
  sheet.validations = readValidations(ws);
  const af = kid(ws, 'autoFilter');
  if (af) sheet.autoFilter = readAutoFilter(af, ctx, kid(ws, 'sortState'));

  for (const h of kids(kid(ws, 'hyperlinks'), 'hyperlink')) {
    const r = parseRange(attr(h, 'ref') ?? '');
    if (!r) continue;
    const rel = rels.get(attr(h, 'id') ?? '');
    let target = rel ? rel.target : '';
    const loc = attr(h, 'location');
    if (loc) target = rel ? `${target}#${loc}` : `#${loc}`;
    if (!target || !SAFE_LINK.test(target)) continue;
    let n = 0;
    for (let i = r.r1; i <= r.r2 && n < 10000; i++) for (let j = r.c1; j <= r.c2 && n < 10000; j++, n++) sheet.hyperlinks.set(cellKey(i, j), target);
  }

  readProtectedRanges(ws, sheet);
  const prot = kid(ws, 'sheetProtection');
  if (prot && boolAttr(prot, 'sheet')) {
    const p: NonNullable<Sheet['protection']> = { sheet: true };
    const pw = attr(prot, 'password');
    if (pw) p.password = pw;
    const options: Record<string, boolean> = {};
    for (const k in prot.attrs) {
      if (k === 'sheet' || k === 'password' || /^(algorithmName|hashValue|saltValue|spinCount)$/.test(k)) continue;
      options[k] = boolAttr(prot, k);
    }
    if (Object.keys(options).length) p.options = options;
    sheet.protection = p;
  }

  const ps = kid(ws, 'pageSetup');
  const pm = kid(ws, 'pageMargins');
  if (ps || pm) {
    const setup: NonNullable<Sheet['pageSetup']> = {};
    const o = attr(ps, 'orientation');
    if (o === 'portrait' || o === 'landscape') setup.orientation = o;
    const paper = numAttr(ps, 'paperSize');
    if (paper) setup.paper = paper;
    if (boolAttr(kid(pr, 'pageSetUpPr'), 'fitToPage')) {
      setup.fitWidth = numAttr(ps, 'fitToWidth') ?? 1;
      setup.fitHeight = numAttr(ps, 'fitToHeight') ?? 1;
    }
    if (pm) {
      setup.margins = {
        left: numAttr(pm, 'left') ?? 0.7, right: numAttr(pm, 'right') ?? 0.7, top: numAttr(pm, 'top') ?? 0.75,
        bottom: numAttr(pm, 'bottom') ?? 0.75, header: numAttr(pm, 'header') ?? 0.3, footer: numAttr(pm, 'footer') ?? 0.3,
      };
    }
    if (Object.keys(setup).length) sheet.pageSetup = setup;
  }

  for (const tp of kids(kid(ws, 'tableParts'), 'tablePart')) {
    const rel = rels.get(attr(tp, 'id') ?? '');
    const t = rel && !rel.external ? ctx.pkg.xml(rel.target) : undefined;
    const def = t && readTable(t);
    if (def) sheet.tables.push(def);
  }
  for (const rel of rels.values()) {
    if (rel.external) continue;
    if (rel.type === 'comments') readComments(rel.target, sheet, ctx);
    else if (rel.type === 'drawing') readDrawing(rel.target, sheet, ctx);
  }
  return sheet;
};

/** Package parts kept for writing back: macros, custom XML data and custom document properties. */
const isKeepable = (p: string): boolean =>
  /^xl\/vbaProject\.bin$/i.test(p) || /^customXml\/(item\d+\.xml|itemProps\d+\.xml|_rels\/item\d+\.xml\.rels)$/i.test(p) || /^docProps\/custom\.xml$/i.test(p);

export const readXlsx = (data: Uint8Array): Workbook => {
  const pkg = new Package(data);
  const rootRels = pkg.rels('');
  const wbPath = rootRels.find((r) => r.type === 'officeDocument')?.target ?? 'xl/workbook.xml';
  const wbx = pkg.xml(wbPath);
  if (!wbx || wbx.name !== 'workbook') throw new Error('This file is not a workbook.');
  const wbRels = pkg.rels(wbPath);
  const relById = new Map(wbRels.map((r) => [r.id, r]));
  const relOf = (type: string): Rel | undefined => wbRels.find((r) => r.type === type);
  const date1904 = boolAttr(kid(wbx, 'workbookPr'), 'date1904');

  const themeRel = relOf('theme');
  const theme = readTheme(themeRel ? pkg.text(themeRel.target) : undefined);
  const stylesRel = relOf('styles');
  const { xfs, dxfs } = readStyles(stylesRel ? pkg.text(stylesRel.target) : undefined, theme);
  const ssRel = relOf('sharedStrings');
  const ssx = ssRel ? pkg.xml(ssRel.target) : undefined;
  const strings = kids(ssx, 'si').map((si) => readRich(si, theme));

  const dynamicCm = new Set<string>();
  const mdRel = relOf('sheetMetadata');
  const md = mdRel ? pkg.xml(mdRel.target) : undefined;
  if (md) {
    const types = kids(kid(md, 'metadataTypes'), 'metadataType').map((t) => attr(t, 'name'));
    const future = kids(md, 'futureMetadata').find((f) => attr(f, 'name') === 'XLDAPR');
    const dynBk = kids(future, 'bk').map((b) => attr(find(b, 'dynamicArrayProperties'), 'fDynamic') === '1');
    kids(kid(md, 'cellMetadata'), 'bk').forEach((bk, i) => {
      const rc = kid(bk, 'rc');
      if (types[(numAttr(rc, 't') ?? 1) - 1] === 'XLDAPR' && dynBk[numAttr(rc, 'v') ?? 0]) dynamicCm.add(String(i + 1));
    });
  }

  const ctx: Ctx = { pkg, theme, xfs, dxfs, strings, date1904, dynamicCm };
  const wb: Workbook = { sheets: [], names: [], styles: xfs, date1904, activeSheet: 0 };

  for (const s of kids(kid(wbx, 'sheets'), 'sheet')) {
    const rel = relById.get(attr(s, 'id') ?? '');
    if (!rel || rel.type !== 'worksheet') continue;
    const sheet = readSheet(rel.target, attr(s, 'name') ?? `Sheet${wb.sheets.length + 1}`, ctx);
    const state = attr(s, 'state');
    if (state === 'hidden' || state === 'veryHidden') sheet.state = state;
    wb.sheets.push(sheet);
  }
  if (!wb.sheets.length) throw new Error('This workbook has no worksheets.');
  wb.activeSheet = Math.min(numAttr(kid(kid(wbx, 'bookViews'), 'workbookView'), 'activeTab') ?? 0, wb.sheets.length - 1);

  for (const dn of kids(kid(wbx, 'definedNames'), 'definedName')) {
    const name = attr(dn, 'name') ?? '';
    const local = numAttr(dn, 'localSheetId');
    const formula = fx(dn.text);
    if (/^_xlnm\._FilterDatabase$/i.test(name)) continue;
    if (/^_xlnm\.Print_Area$/i.test(name) && local !== undefined && wb.sheets[local]) {
      const sh = wb.sheets[local]!;
      sh.pageSetup = { ...sh.pageSetup, printArea: formula.split(',').map((p) => p.slice(p.lastIndexOf('!') + 1)).join(',') };
      continue;
    }
    const d: DefinedName = { name, formula };
    if (local !== undefined) d.sheet = local;
    if (boolAttr(dn, 'hidden')) d.hidden = true;
    const comment = attr(dn, 'comment');
    if (comment) d.comment = comment;
    wb.names.push(d);
  }

  const calc = kid(wbx, 'calcPr');
  if (calc) {
    const c: NonNullable<Workbook['calc']> = {};
    if (boolAttr(calc, 'iterate')) c.iterate = true;
    const it = numAttr(calc, 'iterateCount');
    if (it !== undefined) c.maxIterations = it;
    const ch = numAttr(calc, 'iterateDelta');
    if (ch !== undefined) c.maxChange = ch;
    if (attr(calc, 'calcMode') === 'manual') c.manual = true;
    if (Object.keys(c).length) wb.calc = c;
  }

  const coreRel = rootRels.find((r) => r.type === 'core-properties');
  const core = coreRel ? pkg.xml(coreRel.target) : undefined;
  if (core) {
    const props: NonNullable<Workbook['props']> = {};
    const get = (n: string): string | undefined => { const x = kid(core, n); return x ? deepText(x).trim() || undefined : undefined; };
    const title = get('title'), creator = get('creator'), created = get('created'), modified = get('modified');
    if (title) props.title = title;
    if (creator) props.creator = creator;
    if (created) props.created = created;
    if (modified) props.modified = modified;
    if (Object.keys(props).length) wb.props = props;
  }

  const objRel = relOf(OBJECTS_REL_SHORT);
  if (objRel && !objRel.external) applyObjectsPart(wb, pkg.text(objRel.target));

  // Saved scripts are only stored as text here; they never run when a file opens.
  const scriptsRel = relOf('scripts');
  const scriptsX = scriptsRel ? pkg.xml(scriptsRel.target) : undefined;
  if (scriptsX?.name === 'scripts') {
    const list = kids(scriptsX, 'script').slice(0, 500).map((n, i) => ({
      id: (attr(n, 'id') ?? '').replace(/[^\w-]/g, '').slice(0, 64) || `s${i + 1}`,
      name: (attr(n, 'name') ?? '').slice(0, 200) || `Script ${i + 1}`,
      code: n.text.slice(0, 1_000_000),
    }));
    if (list.length) wb.scripts = list;
  }

  const extra = new Map<string, Uint8Array>();
  for (const p of Object.keys(pkg.files)) if (isKeepable(p)) extra.set(p, pkg.files[p]!);
  if (extra.size) wb.extraParts = extra;
  return wb;
};
