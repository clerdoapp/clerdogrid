// .xlsx writer: a minimal, valid SpreadsheetML package for everything the Workbook model holds.
import { zipSync, type Zippable } from 'fflate';
import {
  addr, keyCol, keyRow, rangeAddr,
  type CellRange, type CfRule, type RichRun, type CfValue, type ConditionalFormat, type Note, type TableDef, type Workbook,
} from '../model/workbook';
import { isError, type Scalar } from '../engine/types';
import { XML_HEAD, el, esc, escXstr } from './xml';
import { StyleWriter, argb, fontXml } from './xlsx-styles';
import { DRAWING_HEAD, anchorXml, chartFrameXml, chartXml, fileColWidth, pictureXml } from './xlsx-chart';
import { addFuturePrefixes, quoteSheet } from './formula-xml';
import { OBJECTS_CT, OBJECTS_PATH, OBJECTS_REL, legacyNotes, objectsPart, protectedRangesXml, shapeXml, withDescr } from './xlsx-objects';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.';

const enc = new TextEncoder();

/** Formula as stored: future-function prefixes, and A1# spill refs as ANCHORARRAY(A1). */
const fout = (f: string): string =>
  esc(addFuturePrefixes(f.replace(/((?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?(\$?[A-Za-z]{1,3}\$?\d+)#/g, (_, sh: string | undefined, a: string) => `_xlfn.ANCHORARRAY(${sh ?? ''}${a})`)));

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

class SharedStrings {
  private index = new Map<string, number>();
  items: string[] = [];
  count = 0;
  add(text: string, rich?: RichRun[]): number {
    this.count++;
    const xml = rich && rich.length
      ? rich.map((r) => `<r>${r.font ? fontXml(r.font, 'rPr', false) : ''}<t xml:space="preserve">${escXstr(r.text)}</t></r>`).join('')
      : `<t${/^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : ''}>${escXstr(text)}</t>`;
    let i = this.index.get(xml);
    if (i === undefined) { i = this.items.length; this.items.push(xml); this.index.set(xml, i); }
    return i;
  }
  xml(): string {
    return `${XML_HEAD}<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.items.length}">${this.items.map((x) => `<si>${x}</si>`).join('')}</sst>`;
  }
}

const numText = (n: number): string => (Number.isFinite(n) ? String(n) : '0');

const cellValueXml = (v: Scalar, sst: SharedStrings, rich: RichRun[] | undefined, formula: boolean): { t?: string; v?: string } => {
  if (v === null) return {};
  if (typeof v === 'number') return Number.isFinite(v) ? { v: numText(v) } : { t: 'e', v: '#NUM!' };
  if (typeof v === 'boolean') return { t: 'b', v: v ? '1' : '0' };
  if (isError(v)) return { t: 'e', v: v.code };
  if (formula) return { t: 'str', v: escXstr(v) };
  return { t: 's', v: String(sst.add(v, rich)) };
};

const cfvoXml = (c: CfValue): string => el('cfvo', { type: c.type === 'autoMin' ? 'min' : c.type === 'autoMax' ? 'max' : c.type, val: c.value !== undefined ? c.value : undefined });

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

const cfXml = (cf: ConditionalFormat, startPriority: number, styles: StyleWriter): { xml: string; used: number } => {
  const first = cf.ranges[0] ? addr(cf.ranges[0].r1, cf.ranges[0].c1) : 'A1';
  let p = startPriority;
  const rules = cf.rules.map((r) => {
    const base = { priority: p++, stopIfTrue: 'stop' in r && r.stop ? true : undefined };
    const dxfId = 'style' in r ? styles.dxf(r.style) : undefined;
    const f = (x: string): string => `<formula>${fout(x)}</formula>`;
    switch (r.type) {
      case 'cellIs': return el('cfRule', { type: r.type, dxfId, ...base, operator: r.operator }, r.formulas.map(f).join(''));
      case 'expression': return el('cfRule', { type: r.type, dxfId, ...base }, f(r.formula));
      case 'colorScale':
        return el('cfRule', { type: r.type, ...base }, `<colorScale>${r.stops.map((s) => cfvoXml(s.value)).join('')}${r.stops.map((s) => el('color', { rgb: argb(s.color) })).join('')}</colorScale>`);
      case 'dataBar':
        return el('cfRule', { type: r.type, ...base }, `<dataBar${r.showValue === false ? ' showValue="0"' : ''}>${cfvoXml(r.min)}${cfvoXml(r.max)}${el('color', { rgb: argb(r.color) })}</dataBar>`);
      case 'iconSet':
        return el('cfRule', { type: r.type, ...base }, `<iconSet${el('x', { iconSet: r.set, reverse: r.reverse || undefined, showValue: r.showValue === false ? false : undefined }).slice(2, -2)}>${r.values.map(cfvoXml).join('')}</iconSet>`);
      case 'top10': return el('cfRule', { type: r.type, dxfId, ...base, percent: r.percent || undefined, bottom: r.bottom || undefined, rank: r.rank });
      case 'aboveAverage': return el('cfRule', { type: r.type, dxfId, ...base, aboveAverage: r.below ? false : undefined, equalAverage: r.equal || undefined, stdDev: r.stdDev || undefined });
      case 'duplicateValues': case 'uniqueValues': return el('cfRule', { type: r.type, dxfId, ...base });
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith':
        return el('cfRule', { type: r.type, dxfId, ...base, operator: r.type === 'containsText' ? 'containsText' : r.type === 'notContainsText' ? 'notContains' : r.type, text: r.text }, f(textRuleFormula(r, first)));
      case 'containsBlanks': return el('cfRule', { type: r.type, dxfId, ...base }, f(`LEN(TRIM(${first}))=0`));
      case 'notContainsBlanks': return el('cfRule', { type: r.type, dxfId, ...base }, f(`LEN(TRIM(${first}))>0`));
      case 'containsErrors': return el('cfRule', { type: r.type, dxfId, ...base }, f(`ISERROR(${first})`));
      case 'notContainsErrors': return el('cfRule', { type: r.type, dxfId, ...base }, f(`NOT(ISERROR(${first}))`));
      case 'timePeriod': return el('cfRule', { type: r.type, dxfId, ...base, timePeriod: r.period }, f(PERIOD_FORMULA[r.period]!(first)));
    }
    return '';
  });
  return { xml: el('conditionalFormatting', { sqref: cf.ranges.map(rangeAddr).join(' ') }, rules.join('')), used: p - startPriority };
};

const tableXml = (t: TableDef, id: number): string => {
  const cols = t.columns.map((c, i) => {
    const std = c.totalsFunction && /^(sum|min|max|average|count|countNums|stdDev|var)$/.test(c.totalsFunction);
    const inner = (c.calculatedFormula ? `<calculatedColumnFormula>${fout(c.calculatedFormula)}</calculatedColumnFormula>` : '')
      + (c.totalsFunction && !std ? `<totalsRowFormula>${fout(c.totalsFunction)}</totalsRowFormula>` : '');
    return el('tableColumn', { id: i + 1, name: escXstrAttr(c.name), totalsRowFunction: c.totalsFunction ? (std ? c.totalsFunction : 'custom') : undefined }, inner);
  }).join('');
  const dataRange: CellRange = { ...t.range, r2: t.totalsRow ? t.range.r2 - 1 : t.range.r2 };
  return `${XML_HEAD}<table xmlns="${NS_MAIN}"${el('x', {
    id, name: t.name, displayName: t.name, ref: rangeAddr(t.range),
    headerRowCount: t.headerRow ? undefined : 0, totalsRowCount: t.totalsRow ? 1 : undefined, totalsRowShown: t.totalsRow ? undefined : false,
  }).slice(2, -2)}>${t.headerRow ? el('autoFilter', { ref: rangeAddr(dataRange) }) : ''}<tableColumns count="${t.columns.length}">${cols}</tableColumns>`
    + el('tableStyleInfo', { name: t.style ?? 'TableStyleMedium2', showFirstColumn: false, showLastColumn: false, showRowStripes: t.bandedRows ?? true, showColumnStripes: t.bandedCols ?? false })
    + '</table>';
};
// `el` escapes attributes itself; table column names only need control characters encoded.
const escXstrAttr = (s: string): string => s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);

const commentsXml = (notes: Map<number, Note>): string => {
  const authors: string[] = [];
  const list = [...notes].sort((a, b) => a[0] - b[0]).map(([k, n]) => {
    const author = n.author ?? '';
    let ai = authors.indexOf(author);
    if (ai < 0) { ai = authors.length; authors.push(author); }
    return `<comment ref="${addr(keyRow(k), keyCol(k))}" authorId="${ai}"><text><r><t xml:space="preserve">${escXstr(n.text)}</t></r></text></comment>`;
  });
  return `${XML_HEAD}<comments xmlns="${NS_MAIN}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${list.join('')}</commentList></comments>`;
};

/** Legacy VML shapes so notes show in apps that draw them from VML; each part gets its own 1024-id block. */
const vmlXml = (notes: Map<number, Note>, block: number): string => {
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

const SAFE_LINK = /^(https:|mailto:|#)/i;
const IMG_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/tiff': 'tiff', 'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' };

interface Out {
  files: Zippable;
  overrides: Map<string, string>;
  defaults: Map<string, string>;
  counters: { table: number; drawing: number; chart: number; image: number; comments: number };
  styles: StyleWriter;
  sst: SharedStrings;
  usesDynamic: boolean;
}

const put = (o: Out, path: string, text: string, ct?: string): void => {
  o.files[path] = enc.encode(text);
  if (ct) o.overrides.set('/' + path, ct);
};

const sheetXml = (wb: Workbook, si: number, o: Out): string => {
  const s = wb.sheets[si]!;
  const rels = new Rels();
  let x = `${XML_HEAD}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`;

  const fit = s.pageSetup && (s.pageSetup.fitWidth !== undefined || s.pageSetup.fitHeight !== undefined);
  const prInner = (s.tabColor ? el('tabColor', { rgb: argb(s.tabColor) }) : '') + (fit ? '<pageSetUpPr fitToPage="1"/>' : '');
  if (prInner) x += `<sheetPr>${prInner}</sheetPr>`;

  // used range
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const k of s.cells.keys()) {
    const r = keyRow(k), c = keyCol(k);
    if (r < r1) r1 = r;
    if (r > r2) r2 = r;
    if (c < c1) c1 = c;
    if (c > c2) c2 = c;
  }
  x += el('dimension', { ref: r2 < 0 ? 'A1' : rangeAddr({ r1, c1, r2, c2 }) });

  const v = s.view;
  let pane = '';
  let selPane: string | undefined;
  if (v.frozenRows || v.frozenCols) {
    selPane = v.frozenRows && v.frozenCols ? 'bottomRight' : v.frozenRows ? 'bottomLeft' : 'topRight';
    pane = el('pane', { xSplit: v.frozenCols || undefined, ySplit: v.frozenRows || undefined, topLeftCell: addr(v.frozenRows, v.frozenCols), activePane: selPane, state: 'frozen' });
  }
  const sel = v.selection ? el('selection', { pane: selPane, activeCell: addr(v.selection.r1, v.selection.c1), sqref: rangeAddr(v.selection) }) : selPane ? el('selection', { pane: selPane }) : '';
  x += `<sheetViews>${el('sheetView', {
    showGridLines: v.showGrid ? undefined : false, showRowColHeaders: v.showHeaders ? undefined : false, rightToLeft: v.rtl || undefined,
    tabSelected: si === wb.activeSheet || undefined, zoomScale: v.zoom !== 100 ? v.zoom : undefined, zoomScaleNormal: v.zoom !== 100 ? v.zoom : undefined, workbookViewId: 0,
  }, pane + sel)}</sheetViews>`;

  x += el('sheetFormatPr', {
    baseColWidth: 8,
    defaultColWidth: Math.abs(s.defaultColWidth - 8.43) > 0.001 ? fileColWidth(s.defaultColWidth) : undefined,
    defaultRowHeight: s.defaultRowHeight,
    customHeight: s.defaultRowHeight !== 15 || undefined,
  });

  if (s.cols.size) {
    const cols = [...s.cols].sort((a, b) => a[0] - b[0]);
    const groups: { min: number; max: number; key: string; info: (typeof cols)[number][1] }[] = [];
    for (const [c, info] of cols) {
      const key = JSON.stringify(info);
      const last = groups[groups.length - 1];
      if (last && last.max === c - 1 && last.key === key) last.max = c;
      else groups.push({ min: c, max: c, key, info });
    }
    x += `<cols>${groups.map((g) => el('col', {
      min: g.min + 1, max: g.max + 1, width: fileColWidth(g.info.width ?? s.defaultColWidth), style: g.info.s || undefined,
      hidden: g.info.hidden || undefined, customWidth: g.info.width !== undefined || undefined, outlineLevel: g.info.outline || undefined,
    })).join('')}</cols>`;
  }

  // rows
  const byRow = new Map<number, number[]>();
  for (const k of s.cells.keys()) {
    const r = keyRow(k);
    let arr = byRow.get(r);
    if (!arr) byRow.set(r, (arr = []));
    arr.push(k);
  }
  for (const r of s.rows.keys()) if (!byRow.has(r)) byRow.set(r, []);
  const rowsSorted = [...byRow.keys()].sort((a, b) => a - b);
  let data = '';
  for (const r of rowsSorted) {
    const info = s.rows.get(r);
    let cells = '';
    for (const k of byRow.get(r)!.sort((a, b) => a - b)) {
      const cell = s.cells.get(k)!;
      const ref = addr(r, keyCol(k));
      const val = cellValueXml(cell.v, o.sst, cell.rich, !!cell.f);
      let fx = '';
      let cm: number | undefined;
      if (cell.f) {
        if (cell.dyn) {
          fx = `<f t="array" ref="${ref}">${fout(cell.f)}</f>`;
          cm = 1;
          o.usesDynamic = true;
        } else if (cell.arrayRange) fx = `<f t="array" ref="${rangeAddr(cell.arrayRange)}">${fout(cell.f)}</f>`;
        else fx = `<f>${fout(cell.f)}</f>`;
      }
      const s0 = cell.s && cell.s < wb.styles.length ? cell.s : undefined;
      cells += `<c r="${ref}"${s0 ? ` s="${s0}"` : ''}${val.t ? ` t="${val.t}"` : ''}${cm ? ` cm="${cm}"` : ''}>${fx}${val.v !== undefined ? `<v>${val.v}</v>` : ''}</c>`;
    }
    data += `<row r="${r + 1}"${info?.s ? ` s="${info.s}" customFormat="1"` : ''}${info?.height !== undefined ? ` ht="${info.height}" customHeight="1"` : ''}${info?.hidden ? ' hidden="1"' : ''}${info?.outline ? ` outlineLevel="${info.outline}"` : ''}>${cells}</row>`;
  }
  x += data ? `<sheetData>${data}</sheetData>` : '<sheetData/>';

  if (s.protection?.sheet) {
    const opts: Record<string, boolean> = { ...s.protection.options };
    x += el('sheetProtection', { password: s.protection.password, sheet: true, ...opts });
  }
  x += protectedRangesXml(s);
  if (s.autoFilter) {
    const af = s.autoFilter;
    const cols = af.columns.map((c) => {
      let inner = '';
      if (c.values || c.blanks) inner = el('filters', { blank: c.blanks || undefined }, (c.values ?? []).map((val) => el('filter', { val })).join(''));
      else if (c.custom) inner = el('customFilters', { and: c.custom.and || undefined }, c.custom.rules.map((r) => el('customFilter', { operator: r.operator === 'equal' ? undefined : r.operator, val: r.value })).join(''));
      else if (c.top10) inner = el('top10', { top: c.top10.top ? undefined : false, percent: c.top10.percent || undefined, val: c.top10.value });
      else if (c.color) inner = el('colorFilter', { dxfId: o.styles.dxf({ fill: { color: c.color } }) });
      return el('filterColumn', { colId: c.col }, inner);
    }).join('');
    const sort = af.sort?.length
      ? el('sortState', { ref: rangeAddr({ ...af.range, r1: af.range.r1 + 1 }) }, af.sort.map((srt) => el('sortCondition', {
        descending: srt.descending || undefined, ref: rangeAddr({ r1: af.range.r1 + 1, r2: af.range.r2, c1: af.range.c1 + srt.col, c2: af.range.c1 + srt.col }),
      })).join(''))
      : '';
    x += el('autoFilter', { ref: rangeAddr(af.range) }, cols + sort);
  }
  if (s.merges.length) x += `<mergeCells count="${s.merges.length}">${s.merges.map((m) => el('mergeCell', { ref: rangeAddr(m) })).join('')}</mergeCells>`;
  let prio = 1;
  for (const cf of s.conditionalFormats) {
    const { xml, used } = cfXml(cf, prio, o.styles);
    prio += used;
    x += xml;
  }
  if (s.validations.length) {
    x += `<dataValidations count="${s.validations.length}">${s.validations.map((d) => el('dataValidation', {
      type: d.type === 'any' ? undefined : d.type, errorStyle: d.error && d.error.style !== 'stop' ? d.error.style : undefined,
      operator: d.operator && d.operator !== 'between' ? d.operator : undefined, allowBlank: d.allowBlank || undefined,
      showDropDown: d.type === 'list' && d.showDropdown === false ? true : undefined,
      showInputMessage: d.prompt ? true : undefined, showErrorMessage: d.error ? true : undefined,
      errorTitle: d.error?.title, error: d.error?.text, promptTitle: d.prompt?.title, prompt: d.prompt?.text, sqref: d.ranges.map(rangeAddr).join(' '),
    }, (d.formula1 !== undefined ? `<formula1>${fout(d.formula1)}</formula1>` : '') + (d.formula2 !== undefined ? `<formula2>${fout(d.formula2)}</formula2>` : ''))).join('')}</dataValidations>`;
  }
  if (s.hyperlinks.size) {
    let hl = '';
    for (const [k, target] of [...s.hyperlinks].sort((a, b) => a[0] - b[0])) {
      if (!SAFE_LINK.test(target)) continue;
      const ref = addr(keyRow(k), keyCol(k));
      if (target.startsWith('#')) hl += el('hyperlink', { ref, location: target.slice(1), display: target.slice(1) });
      else hl += el('hyperlink', { ref, 'r:id': rels.add('hyperlink', target, true) });
    }
    if (hl) x += `<hyperlinks>${hl}</hyperlinks>`;
  }
  const m = s.pageSetup?.margins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
  x += el('pageMargins', m);
  const ps = s.pageSetup;
  if (ps && (ps.orientation || ps.paper || fit)) {
    x += el('pageSetup', { paperSize: ps.paper, fitToWidth: fit ? ps.fitWidth ?? 1 : undefined, fitToHeight: fit ? ps.fitHeight ?? 1 : undefined, orientation: ps.orientation });
  }

  // drawings (charts, pictures and shapes share one drawing part)
  if (s.charts.length || s.pictures.length || s.shapes?.length) {
    const dn = ++o.counters.drawing;
    const dPath = `xl/drawings/drawing${dn}.xml`;
    const dRels = new Rels();
    let body = '';
    let shapeId = 1;
    for (const c of s.charts) {
      const cn = ++o.counters.chart;
      put(o, `xl/charts/chart${cn}.xml`, chartXml(c), CT + 'drawingml.chart+xml');
      const rid = dRels.add('chart', `../charts/chart${cn}.xml`);
      body += anchorXml(c.anchor, s, withDescr(chartFrameXml(++shapeId, c.title ?? `Chart ${shapeId - 1}`, rid), c.alt));
    }
    for (const p of s.pictures) {
      const ext = IMG_EXT[p.mime] ?? 'png';
      const n = ++o.counters.image;
      o.files[`xl/media/image${n}.${ext}`] = p.data;
      o.defaults.set(ext, p.mime);
      const rid = dRels.add('image', `../media/image${n}.${ext}`);
      body += anchorXml(p.anchor, s, withDescr(pictureXml(++shapeId, `Picture ${n}`, rid, p.anchor), p.alt));
    }
    for (const sh of s.shapes ?? []) body += anchorXml(sh.anchor, s, shapeXml(++shapeId, sh));
    put(o, dPath, `${XML_HEAD}${DRAWING_HEAD}${body}</xdr:wsDr>`, CT + 'drawing+xml');
    put(o, `xl/drawings/_rels/drawing${dn}.xml.rels`, dRels.xml());
    x += `<drawing r:id="${rels.add('drawing', `../drawings/drawing${dn}.xml`)}"/>`;
  }
  const notes = legacyNotes(s);
  if (notes.size) {
    const n = ++o.counters.comments;
    put(o, `xl/comments${n}.xml`, commentsXml(notes), CT + 'spreadsheetml.comments+xml');
    o.files[`xl/drawings/vmlDrawing${n}.vml`] = enc.encode(vmlXml(notes, n));
    o.defaults.set('vml', CT + 'vmlDrawing');
    rels.add('comments', `../comments${n}.xml`);
    x += `<legacyDrawing r:id="${rels.add('vmlDrawing', `../drawings/vmlDrawing${n}.vml`)}"/>`;
  }
  if (s.tables.length) {
    const ids = s.tables.map((t) => {
      const n = ++o.counters.table;
      put(o, `xl/tables/table${n}.xml`, tableXml(t, n), CT + 'spreadsheetml.table+xml');
      return rels.add('table', `../tables/table${n}.xml`);
    });
    x += `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`;
  }
  x += '</worksheet>';
  if (rels.size) put(o, `xl/worksheets/_rels/sheet${si + 1}.xml.rels`, rels.xml());
  return x;
};

const METADATA = `${XML_HEAD}<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">`
  + '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
  + '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
  + '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

const NS_SCRIPTS = 'http://schemas.clerdo.app/grid/2026/scripts';
const RT_SCRIPTS = 'http://schemas.clerdo.app/grid/2026/relationships/scripts';

export const writeXlsx = (wb: Workbook): Uint8Array => {
  const o: Out = {
    files: {}, overrides: new Map(), defaults: new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]),
    counters: { table: 0, drawing: 0, chart: 0, image: 0, comments: 0 },
    styles: new StyleWriter(wb.styles), sst: new SharedStrings(), usesDynamic: false,
  };
  const sheets = wb.sheets.length ? wb.sheets : [];
  if (!sheets.length) throw new Error('A workbook needs at least one sheet.');
  const wbRels = new Rels();
  const sheetEls: string[] = [];
  sheets.forEach((s, i) => {
    put(o, `xl/worksheets/sheet${i + 1}.xml`, sheetXml(wb, i, o), CT + 'spreadsheetml.worksheet+xml');
    const rid = wbRels.add('worksheet', `worksheets/sheet${i + 1}.xml`);
    sheetEls.push(el('sheet', { name: s.name, sheetId: i + 1, state: s.state === 'visible' ? undefined : s.state, 'r:id': rid }));
  });

  // defined names, plus each sheet's print area and autofilter database
  const names: string[] = [];
  const nameEl = (name: string, formula: string, sheet?: number, hidden?: boolean, comment?: string): string =>
    el('definedName', { name, comment, localSheetId: sheet, hidden: hidden || undefined }, fout(formula));
  for (const n of wb.names) names.push(nameEl(n.name, n.formula, n.sheet, n.hidden, n.comment));
  sheets.forEach((s, i) => {
    const q = quoteSheet(s.name);
    if (s.autoFilter) names.push(nameEl('_xlnm._FilterDatabase', `${q}!${rangeAddr(s.autoFilter.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`, i, true));
    if (s.pageSetup?.printArea) {
      const pa = s.pageSetup.printArea.split(',').map((p) => `${q}!${p.trim().replace(/\$?([A-Z]+)\$?(\d+)/g, '$$$1$$$2')}`).join(',');
      names.push(nameEl('_xlnm.Print_Area', pa, i));
    }
  });

  const activeTab = Math.max(0, Math.min(wb.activeSheet, sheets.length - 1));
  const calc = wb.calc ?? {};
  const wbXml = `${XML_HEAD}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`
    + el('workbookPr', { date1904: wb.date1904 || undefined, defaultThemeVersion: 166925 })
    + `<bookViews>${el('workbookView', { xWindow: 0, yWindow: 0, windowWidth: 28800, windowHeight: 12300, activeTab: activeTab || undefined })}</bookViews>`
    + `<sheets>${sheetEls.join('')}</sheets>`
    + (names.length ? `<definedNames>${names.join('')}</definedNames>` : '')
    + el('calcPr', { calcId: 191029, calcMode: calc.manual ? 'manual' : undefined, iterate: calc.iterate || undefined, iterateCount: calc.maxIterations, iterateDelta: calc.maxChange, fullCalcOnLoad: true })
    + '</workbook>';

  put(o, 'xl/styles.xml', o.styles.xml(), CT + 'spreadsheetml.styles+xml');
  wbRels.add('styles', 'styles.xml');
  put(o, 'xl/theme/theme1.xml', THEME, CT + 'theme+xml');
  wbRels.add('theme', 'theme/theme1.xml');
  const objects = objectsPart(wb);
  if (objects) {
    put(o, OBJECTS_PATH, objects, OBJECTS_CT);
    wbRels.add(OBJECTS_REL, OBJECTS_PATH.replace(/^xl\//, ''));
  }
  if (o.sst.items.length) {
    put(o, 'xl/sharedStrings.xml', o.sst.xml(), CT + 'spreadsheetml.sharedStrings+xml');
    wbRels.add('sharedStrings', 'sharedStrings.xml');
  }
  if (o.usesDynamic) {
    put(o, 'xl/metadata.xml', METADATA, CT + 'spreadsheetml.sheetMetadata+xml');
    wbRels.add('sheetMetadata', 'metadata.xml');
  }

  if (wb.scripts?.length) {
    const items = wb.scripts.map((sc) => el('script', { id: sc.id, name: sc.name }, esc(sc.code))).join('');
    put(o, 'xl/clerdo/scripts.xml', `${XML_HEAD}<scripts xmlns="${NS_SCRIPTS}">${items}</scripts>`, 'application/vnd.clerdo.grid.scripts+xml');
    wbRels.add(RT_SCRIPTS, 'clerdo/scripts.xml');
  }

  // parts kept from the file this workbook was read from
  let macro = false;
  const rootRels = new Rels();
  rootRels.add('officeDocument', 'xl/workbook.xml');
  rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
  rootRels.add('extended-properties', 'docProps/app.xml');
  for (const [p, data] of wb.extraParts ?? []) {
    if (o.files[p] || p.includes('..') || p.startsWith('/')) continue;
    o.files[p] = data;
    if (/^xl\/vbaProject\.bin$/i.test(p)) {
      macro = true;
      o.defaults.set('bin', 'application/vnd.ms-office.vbaProject');
      wbRels.add('http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin');
    } else if (/^customXml\/itemProps\d+\.xml$/i.test(p)) o.overrides.set('/' + p, CT + 'customXmlProperties+xml');
    else if (/^docProps\/custom\.xml$/i.test(p)) {
      o.overrides.set('/' + p, CT + 'custom-properties+xml');
      rootRels.add('custom-properties', 'docProps/custom.xml');
    }
  }
  const customItems = Object.keys(o.files).filter((p) => /^customXml\/item\d+\.xml$/i.test(p));
  for (const p of customItems) wbRels.add('customXml', `../${p}`);

  put(o, 'xl/workbook.xml', wbXml, macro ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : CT + 'spreadsheetml.sheet.main+xml');
  put(o, 'xl/_rels/workbook.xml.rels', wbRels.xml());
  put(o, '_rels/.rels', rootRels.xml());

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const p = wb.props ?? {};
  put(o, 'docProps/core.xml', `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + (p.title ? `<dc:title>${esc(p.title)}</dc:title>` : '') + (p.creator ? `<dc:creator>${esc(p.creator)}</dc:creator>` : '')
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${esc(p.created ?? now)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${esc(p.modified ?? now)}</dcterms:modified></cp:coreProperties>`,
  'application/vnd.openxmlformats-package.core-properties+xml');
  put(o, 'docProps/app.xml', `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>ClerdoGrid</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>`
    + `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>`
    + `<TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">${sheets.map((s) => `<vt:lpstr>${esc(s.name)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>`,
  CT + 'extended-properties+xml');

  const ct = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + [...o.defaults].map(([e, t]) => el('Default', { Extension: e, ContentType: t })).join('')
    + [...o.overrides].map(([pn, t]) => el('Override', { PartName: pn, ContentType: t })).join('')
    + '</Types>';
  const files: Zippable = { '[Content_Types].xml': enc.encode(ct), ...o.files };
  return zipSync(files, { level: 6 });
};

const THEME_COLORS: [string, string][] = [['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'], ['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'],
  ['dk2', '<a:srgbClr val="44546A"/>'], ['lt2', '<a:srgbClr val="E7E6E6"/>'], ['accent1', '<a:srgbClr val="4472C4"/>'], ['accent2', '<a:srgbClr val="ED7D31"/>'],
  ['accent3', '<a:srgbClr val="A5A5A5"/>'], ['accent4', '<a:srgbClr val="FFC000"/>'], ['accent5', '<a:srgbClr val="5B9BD5"/>'], ['accent6', '<a:srgbClr val="70AD47"/>'],
  ['hlink', '<a:srgbClr val="0563C1"/>'], ['folHlink', '<a:srgbClr val="954F72"/>']];
const fontScheme = (latin: string): string => `<a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/>`;
const fillStyle = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
const THEME = `${XML_HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>`
  + `<a:clrScheme name="Office">${THEME_COLORS.map(([n, c]) => `<a:${n}>${c}</a:${n}>`).join('')}</a:clrScheme>`
  + `<a:fontScheme name="Office"><a:majorFont>${fontScheme('Calibri Light')}</a:majorFont><a:minorFont>${fontScheme('Calibri')}</a:minorFont></a:fontScheme>`
  + `<a:fmtScheme name="Office"><a:fillStyleLst>${fillStyle}${fillStyle}${fillStyle}</a:fillStyleLst>`
  + '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
  + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
  + `<a:bgFillStyleLst>${fillStyle}${fillStyle}${fillStyle}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

