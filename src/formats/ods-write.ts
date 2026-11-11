// OpenDocument spreadsheet writer (.ods / .ots, ODF 1.3).
import { zipSync, strToU8, type Zippable } from 'fflate';
import {
  DEFAULT_FONT, keyCol, keyRow, type BorderSide, type Cell, type CellRange, type CellStyle, type CfRule, type CfValue, type DataValidation, type RichRun,
  type Sheet, type Workbook,
} from '../model/workbook';
import { CellError } from '../engine/types';
import { fromSerial, timeOf } from '../engine/date';
import { parseFormula, type Node, type RefNode } from '../engine/parser';
import { NS_DECL, tag, textXml } from './ods-xml';
import { odfExpr, odfRef, toOdfFormula, type OdfFormulaCtx } from './ods-formula';
import { currencyCode, dataStyleXml, valueTypeOf } from './ods-numfmt';
import { chartExtra, chartObjectXml } from './ods-chart';
import { formatValue } from './numfmt';
import { fileColWidth } from './xlsx-chart';
import { esc } from './xml';
import { queriesToXml } from './queries-xml';
import { encryptStream, odfKey } from './ods-crypto';
import { sha1, toBase64 } from './crypto';

const MIME = 'application/vnd.oasis.opendocument.spreadsheet';
const MIME_TEMPLATE = 'application/vnd.oasis.opendocument.spreadsheet-template';
const MAX_ROW = 1_048_576, MAX_COL = 16_384;

export interface OdsWriteOptions { template?: boolean; /** Encrypts every stream with this password. */ password?: string }

const inch = (pt: number): string => `${(pt / 72).toFixed(4)}in`;
const colLetters = (c: number): string => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const odfSheet = (s: string): string => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
/** Plain ODF range address: "Sheet1.A1:Sheet1.B5". */
const rangeAddress = (sheet: string, r: CellRange): string => {
  const q = odfSheet(sheet);
  const one = r.r1 === r.r2 && r.c1 === r.c2;
  return one ? `${q}.${colLetters(r.c1)}${r.r1 + 1}` : `${q}.${colLetters(r.c1)}${r.r1 + 1}:${q}.${colLetters(r.c2)}${r.r2 + 1}`;
};
const absAddress = (sheet: string, r: CellRange): string => {
  const q = odfSheet(sheet);
  const one = r.r1 === r.r2 && r.c1 === r.c2;
  return one ? `$${q}.$${colLetters(r.c1)}$${r.r1 + 1}` : `$${q}.$${colLetters(r.c1)}$${r.r1 + 1}:.$${colLetters(r.c2)}$${r.r2 + 1}`;
};

const BORDER_OUT: Record<string, string> = {
  thin: '0.74pt solid', medium: '1.76pt solid', thick: '2.5pt solid', hair: '0.1pt solid', dashed: '0.74pt dashed', mediumDashed: '1.76pt dashed', dotted: '0.74pt dotted',
  double: '2.6pt double', dashDot: '0.74pt dash-dot', mediumDashDot: '1.76pt dash-dot', dashDotDot: '0.74pt dash-dot-dot', mediumDashDotDot: '1.76pt dash-dot-dot', slantDashDot: '1.76pt dash-dot',
};
// inner line, gap, outer line: the proportions other apps read as a double border
const DOUBLE_WIDTHS = '0.0071in 0.0213in 0.0071in';
const borderOut = (b: BorderSide | undefined): string | undefined => (b ? `${BORDER_OUT[b.style] ?? '0.74pt solid'} ${(b.color ?? '#000000').slice(0, 7).toLowerCase()}` : undefined);

const textProps = (f: CellStyle['font'] | undefined, fonts: Set<string>): string => {
  if (!f || !Object.keys(f).length) return '';
  if (f.name) fonts.add(f.name);
  return tag('style:text-properties', {
    'style:font-name': f.name, 'fo:font-size': f.size ? `${f.size}pt` : undefined, 'fo:font-weight': f.bold ? 'bold' : f.bold === false ? 'normal' : undefined,
    'style:font-weight-asian': f.bold ? 'bold' : undefined, 'style:font-weight-complex': f.bold ? 'bold' : undefined,
    'fo:font-style': f.italic ? 'italic' : f.italic === false ? 'normal' : undefined,
    'style:text-underline-style': f.underline ? 'solid' : undefined, 'style:text-underline-width': f.underline ? 'auto' : undefined, 'style:text-underline-color': f.underline ? 'font-color' : undefined,
    'style:text-underline-type': f.underline === 'double' || f.underline === 'doubleAccounting' ? 'double' : undefined,
    'style:text-line-through-style': f.strike ? 'solid' : undefined, 'fo:color': f.color?.slice(0, 7).toLowerCase(),
    'style:text-position': f.vertAlign === 'superscript' ? 'super 58%' : f.vertAlign === 'subscript' ? 'sub 58%' : undefined,
  });
};

const cellStyleInner = (st: CellStyle, fonts: Set<string>, dataStyle?: string): { attrs: Record<string, string | undefined>; inner: string } => {
  const a = st.align ?? {};
  const b = st.border ?? {};
  const prot = st.protection ? (st.protection.locked === false ? (st.protection.hidden ? 'formula-hidden' : 'none') : st.protection.hidden ? 'hidden-and-protected' : 'protected') : undefined;
  const rot = a.rotation === undefined || a.rotation === 255 ? undefined : a.rotation < 0 ? 360 + a.rotation : a.rotation;
  const tc = tag('style:table-cell-properties', {
    'fo:background-color': st.fill?.color?.slice(0, 7).toLowerCase() ?? st.fill?.gradient?.stops[0]?.color.slice(0, 7).toLowerCase(),
    'fo:border-left': borderOut(b.left), 'fo:border-right': borderOut(b.right), 'fo:border-top': borderOut(b.top), 'fo:border-bottom': borderOut(b.bottom),
    'style:diagonal-tl-br': borderOut(b.diagonalDown), 'style:diagonal-bl-tr': borderOut(b.diagonalUp),
    'style:border-line-width-left': b.left?.style === 'double' ? DOUBLE_WIDTHS : undefined, 'style:border-line-width-right': b.right?.style === 'double' ? DOUBLE_WIDTHS : undefined,
    'style:border-line-width-top': b.top?.style === 'double' ? DOUBLE_WIDTHS : undefined, 'style:border-line-width-bottom': b.bottom?.style === 'double' ? DOUBLE_WIDTHS : undefined,
    'style:vertical-align': a.v === 'top' ? 'top' : a.v === 'center' ? 'middle' : a.v === 'bottom' ? 'bottom' : a.v ? 'middle' : undefined,
    'fo:wrap-option': a.wrap ? 'wrap' : undefined, 'style:rotation-angle': rot !== undefined ? String(rot) : undefined, 'style:rotation-align': rot !== undefined ? 'none' : undefined,
    'style:direction': a.rotation === 255 ? 'ttb' : undefined, 'style:shrink-to-fit': a.shrink ? 'true' : undefined, 'style:cell-protect': prot,
    'style:text-align-source': a.h && a.h !== 'general' ? 'fix' : undefined,
  });
  const h = a.h === 'left' ? 'start' : a.h === 'right' ? 'end' : a.h === 'center' || a.h === 'centerContinuous' ? 'center' : a.h === 'justify' || a.h === 'distributed' ? 'justify' : a.h === 'fill' ? 'start' : undefined;
  const pp = h || a.indent || a.rtl ? tag('style:paragraph-properties', { 'fo:text-align': h, 'fo:margin-left': a.indent ? `${a.indent * 10}pt` : undefined, 'style:writing-mode': a.rtl ? 'rl-tb' : undefined }) : '';
  return { attrs: { 'style:data-style-name': dataStyle }, inner: tc + pp + textProps(st.font, fonts) };
};

/** Writes the workbook as an OpenDocument spreadsheet (or template). */
export const writeOds = (wb: Workbook, opts: OdsWriteOptions = {}): Uint8Array => {
  const fonts = new Set<string>([DEFAULT_FONT.name]);
  const autoStyles: string[] = [];
  const commonStyles: string[] = [];
  const files: Zippable = {};
  const manifest: [string, string][] = [];

  // ---------- data and cell styles ----------
  const dataNames = new Map<string, string>();
  const dataStyle = (code: string | undefined): string | undefined => {
    if (!code || code === 'General') return undefined;
    let n = dataNames.get(code);
    if (!n) { n = `N${dataNames.size + 100}`; dataNames.set(code, n); autoStyles.push(dataStyleXml(n, code)); }
    return n;
  };
  const cellStyleName = (i: number | undefined): string | undefined => (i && wb.styles[i] && Object.keys(wb.styles[i]!).length ? `ce${i}` : undefined);
  wb.styles.forEach((st, i) => {
    if (!i || !Object.keys(st).length) return;
    const { attrs, inner } = cellStyleInner(st, fonts, dataStyle(st.numFmt));
    autoStyles.push(tag('style:style', { 'style:name': `ce${i}`, 'style:family': 'table-cell', 'style:parent-style-name': 'Default', ...attrs }, inner));
  });
  const textNames = new Map<string, string>();
  const textStyle = (f: RichRun['font']): string | undefined => {
    if (!f || !Object.keys(f).length) return undefined;
    const key = JSON.stringify(f);
    let n = textNames.get(key);
    if (!n) { n = `T${textNames.size + 1}`; textNames.set(key, n); autoStyles.push(tag('style:style', { 'style:name': n, 'style:family': 'text' }, textProps(f, fonts))); }
    return n;
  };
  const colNames = new Map<string, string>();
  const colStyle = (pt: number): string => {
    const key = pt.toFixed(2);
    let n = colNames.get(key);
    if (!n) { n = `co${colNames.size + 1}`; colNames.set(key, n); autoStyles.push(tag('style:style', { 'style:name': n, 'style:family': 'table-column' }, tag('style:table-column-properties', { 'fo:break-before': 'auto', 'style:column-width': inch(pt) }))); }
    return n;
  };
  const rowNames = new Map<string, string>();
  const rowStyle = (pt: number, custom: boolean, breakBefore = false): string => {
    const key = `${pt.toFixed(2)}|${custom}|${breakBefore}`;
    let n = rowNames.get(key);
    if (!n) { n = `ro${rowNames.size + 1}`; rowNames.set(key, n); autoStyles.push(tag('style:style', { 'style:name': n, 'style:family': 'table-row' }, tag('style:table-row-properties', { 'style:row-height': inch(pt), 'fo:break-before': breakBefore ? 'page' : 'auto', 'style:use-optimal-row-height': !custom }))); }
    return n;
  };
  const colPt = (chars: number): number => (Math.round(fileColWidth(chars) * 7) * 72) / 96;

  // ---------- formulas ----------
  const tableRef = (n: Extract<Node, { type: 'table' }>, from: Sheet): RefNode | undefined => {
    for (const sh of wb.sheets) {
      const t = sh.tables.find((x) => x.name.toUpperCase() === n.table.toUpperCase());
      if (!t) continue;
      const items = n.spec.items;
      let r1 = t.range.r1 + (t.headerRow ? 1 : 0), r2 = t.range.r2 - (t.totalsRow ? 1 : 0);
      if (items.includes('#All')) { r1 = t.range.r1; r2 = t.range.r2; }
      else if (items.includes('#Headers') && !items.includes('#Data')) r1 = r2 = t.range.r1;
      else if (items.includes('#Totals') && !items.includes('#Data')) r1 = r2 = t.range.r2;
      let c1 = t.range.c1, c2 = t.range.c2;
      if (n.spec.col1) {
        const x = t.columns.findIndex((c) => c.name.toUpperCase() === n.spec.col1!.toUpperCase());
        const y = n.spec.col2 ? t.columns.findIndex((c) => c.name.toUpperCase() === n.spec.col2!.toUpperCase()) : x;
        if (x < 0 || y < 0) return undefined;
        c1 = t.range.c1 + Math.min(x, y); c2 = t.range.c1 + Math.max(x, y);
      }
      return { r1, r2, c1, c2, r1Abs: true, r2Abs: true, c1Abs: true, c2Abs: true, kind: r1 === r2 && c1 === c2 ? 'cell' : 'area', sheet: sh === from ? undefined : sh.name };
    }
    return undefined;
  };
  const fctx = (s: Sheet): OdfFormulaCtx => ({ table: (n) => tableRef(n, s) });
  const odfFormula = (f: string, s: Sheet): string | undefined => { try { return toOdfFormula(f, fctx(s)); } catch { return undefined; } };
  const odfBare = (f: string, s: Sheet): string | undefined => { try { return odfExpr(parseFormula(f), fctx(s)); } catch { return undefined; } };

  // ---------- validations ----------
  const validationXml: string[] = [];
  const validationName = new Map<DataValidation, string>();
  wb.sheets.forEach((s) => {
    for (const dv of s.validations) {
      const base = dv.ranges[0];
      if (!base) continue;
      const cond = validationCondition(dv, s);
      if (!cond && dv.type !== 'any') continue;
      const name = `val${validationXml.length + 1}`;
      validationName.set(dv, name);
      const help = dv.prompt ? tag('table:help-message', { 'table:title': dv.prompt.title, 'table:display': true }, dv.prompt.text.split('\n').map((l) => `<text:p>${textXml(l)}</text:p>`).join('')) : '';
      const err = dv.error ? tag('table:error-message', { 'table:message-type': dv.error.style, 'table:title': dv.error.title, 'table:display': true }, dv.error.text.split('\n').map((l) => `<text:p>${textXml(l)}</text:p>`).join(''))
        : tag('table:error-message', { 'table:message-type': 'stop', 'table:display': false });
      validationXml.push(tag('table:content-validation', {
        'table:name': name, 'table:condition': cond ? `of:${cond}` : undefined, 'table:allow-empty-cell': !!dv.allowBlank,
        'table:display-list': dv.type === 'list' ? (dv.showDropdown === false ? 'none' : 'unsorted') : undefined,
        'table:base-cell-address': rangeAddress(s.name, { r1: base.r1, c1: base.c1, r2: base.r1, c2: base.c1 }),
      }, help + err));
    }
  });

  function validationCondition(dv: DataValidation, s: Sheet): string | undefined {
    const f1 = dv.formula1 !== undefined ? odfBare(dv.formula1, s) : undefined;
    const f2 = dv.formula2 !== undefined ? odfBare(dv.formula2, s) : undefined;
    const op = (lhs: string): string | undefined => {
      if (f1 === undefined) return undefined;
      switch (dv.operator ?? 'between') {
        case 'between': return f2 !== undefined ? `${lhs === 'cell-content()' ? 'cell-content-is-between' : 'cell-content-text-length-is-between'}(${f1},${f2})` : undefined;
        case 'notBetween': return f2 !== undefined ? `${lhs === 'cell-content()' ? 'cell-content-is-not-between' : 'cell-content-text-length-is-not-between'}(${f1},${f2})` : undefined;
        case 'equal': return `${lhs}=${f1}`;
        case 'notEqual': return `${lhs}!=${f1}`;
        case 'greaterThan': return `${lhs}>${f1}`;
        case 'lessThan': return `${lhs}<${f1}`;
        case 'greaterThanOrEqual': return `${lhs}>=${f1}`;
        case 'lessThanOrEqual': return `${lhs}<=${f1}`;
      }
    };
    switch (dv.type) {
      case 'any': return undefined;
      case 'list': {
        if (!dv.formula1) return undefined;
        if (/^".*"$/s.test(dv.formula1)) return `cell-content-is-in-list(${dv.formula1.slice(1, -1).replace(/""/g, '"').split(',').map((x) => `"${x.replace(/"/g, '""')}"`).join(';')})`;
        return f1 ? `cell-content-is-in-list(${f1})` : undefined;
      }
      case 'custom': return f1 ? `is-true-formula(${f1})` : undefined;
      case 'textLength': return op('cell-content-text-length()');
      default: {
        const kind = { whole: 'cell-content-is-whole-number()', decimal: 'cell-content-is-decimal-number()', date: 'cell-content-is-date()', time: 'cell-content-is-time()' }[dv.type];
        const c = op('cell-content()');
        return c ? `${kind} and ${c}` : undefined;
      }
    }
  }

  // ---------- conditional format styles (common styles in styles.xml) ----------
  const cfStyleNames = new Map<string, string>();
  const cfStyle = (st: CellStyle): string => {
    const key = JSON.stringify(st);
    let n = cfStyleNames.get(key);
    if (!n) {
      n = `ClerdoCF${cfStyleNames.size + 1}`;
      cfStyleNames.set(key, n);
      const { attrs, inner } = cellStyleInner(st, fonts);
      void attrs;
      commonStyles.push(tag('style:style', { 'style:name': n, 'style:family': 'table-cell', 'style:parent-style-name': 'Default' }, inner));
    }
    return n;
  };

  // ---------- pictures and charts ----------
  let picN = 0, objN = 0;
  const frameXml = (_s: Sheet, anchor: Sheet['pictures'][number]['anchor'], inner: string, extra: Record<string, string | undefined> = {}): string =>
    tag('draw:frame', {
      'draw:z-index': 0, 'draw:name': extra.name, 'draw:style-name': 'gr1',
      'svg:width': inch(anchor.width * 0.75), 'svg:height': inch(anchor.height * 0.75), 'svg:x': inch(anchor.dx * 0.75), 'svg:y': inch(anchor.dy * 0.75),
      ...Object.fromEntries(Object.entries(extra).filter(([k]) => k !== 'name')),
    }, inner);
  const anchoredAt = new Map<Sheet, Map<number, string[]>>();
  wb.sheets.forEach((s) => {
    const at = new Map<number, string[]>();
    anchoredAt.set(s, at);
    const add = (row: number, col: number, x: string) => { const k = row * MAX_COL + col; const l = at.get(k) ?? []; l.push(x); at.set(k, l); };
    for (const p of s.pictures) {
      const ext = (/png/.test(p.mime) ? 'png' : /jpe?g/.test(p.mime) ? 'jpg' : /gif/.test(p.mime) ? 'gif' : /svg/.test(p.mime) ? 'svg' : /bmp/.test(p.mime) ? 'bmp' : /webp/.test(p.mime) ? 'webp' : /tiff/.test(p.mime) ? 'tif' : 'bin');
      const path = `Pictures/image${++picN}.${ext}`;
      files[path] = [p.data, { level: 0 }];
      manifest.push([path, p.mime]);
      const alt = p.alt ? `<svg:desc>${esc(p.alt)}</svg:desc>` : '';
      add(p.anchor.row, p.anchor.col, frameXml(s, p.anchor, tag('draw:image', { 'xlink:href': path, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad', 'draw:mime-type': p.mime }, '<text:p/>') + alt, { name: `Picture ${picN}` }));
    }
    for (const c of s.charts) {
      const dir = `Object ${++objN}`;
      files[`${dir}/content.xml`] = strToU8(chartObjectXml(c, s.name));
      files[`${dir}/styles.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles${NS_DECL('office', 'style')} office:version="1.3"><office:styles/></office:document-styles>`);
      manifest.push([`${dir}/`, 'application/vnd.oasis.opendocument.chart'], [`${dir}/content.xml`, 'text/xml'], [`${dir}/styles.xml`, 'text/xml']);
      const ranges = c.series.flatMap((x) => [x.values, x.categories, x.name]).filter((x): x is string => !!x && !x.startsWith('=')).map((x) => {
        try { const n = parseFormula(x); return n.type === 'ref' ? odfRef(n.ref, s.name, true).replace(/\$/g, '') : ''; } catch { return ''; }
      }).filter(Boolean);
      const alt = c.alt ? `<svg:desc>${esc(c.alt)}</svg:desc>` : '';
      add(c.anchor.row, c.anchor.col, frameXml(s, c.anchor, tag('draw:object', { 'draw:notify-on-update-of-ranges': ranges.join(' ') || undefined, 'xlink:href': `./${dir}`, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad' }, '<loext:p/>'.replace('loext:p', 'text:p')) + alt,
        { name: `Chart ${objN}`, 'clerdo:chart': chartExtra(c) }));
    }
  });

  // ---------- tables ----------
  const tableStyles: string[] = [];
  const pageLayouts: string[] = [];
  const masterPages: string[] = [];
  const tables = wb.sheets.map((s, si) => {
    const taName = `ta${si + 1}`;
    const master = `PageStyle_${si + 1}`;
    tableStyles.push(tag('style:style', { 'style:name': taName, 'style:family': 'table', 'style:master-page-name': master },
      tag('style:table-properties', { 'table:display': s.state === 'visible', 'style:writing-mode': s.view.rtl ? 'rl-tb' : 'lr-tb', 'tableooo:tab-color': s.tabColor?.slice(0, 7).toLowerCase() })));
    pageStyle(s, si, master);
    return tableXml(s, si, taName);
  });

  function pageStyle(s: Sheet, si: number, master: string): void {
    const ps = s.pageSetup ?? {};
    const m = ps.margins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
    const land = ps.orientation === 'landscape';
    const paper = ps.paper === 1 ? [8.5, 11] : ps.paper === 5 ? [8.5, 14] : ps.paper === 9 || ps.paper === undefined ? [8.2677, 11.6929] : [8.2677, 11.6929];
    const [w, h] = land ? [paper[1]!, paper[0]!] : [paper[0]!, paper[1]!];
    const print = [ps.gridlines && 'grid', ps.headings && 'headers', 'charts', 'drawings', 'objects', 'zero-values'].filter(Boolean).join(' ');
    const center = ps.centerH && ps.centerV ? 'both' : ps.centerH ? 'horizontal' : ps.centerV ? 'vertical' : 'none';
    const hasHeader = !!ps.header, hasFooter = !!ps.footer;
    const hTop = hasHeader ? m.header : m.top, fBottom = hasFooter ? m.footer : m.bottom;
    const area = (gap: number) => Math.max(0.1, gap);
    pageLayouts.push(tag('style:page-layout', { 'style:name': `pm${si + 1}` },
      tag('style:page-layout-properties', {
        'fo:page-width': `${w}in`, 'fo:page-height': `${h}in`, 'style:print-orientation': land ? 'landscape' : 'portrait',
        'fo:margin-left': `${m.left}in`, 'fo:margin-right': `${m.right}in`, 'fo:margin-top': `${hTop}in`, 'fo:margin-bottom': `${fBottom}in`,
        'style:print': print, 'style:table-centering': center, 'style:first-page-number': ps.firstPageNumber ?? 'continue',
        'style:scale-to': ps.fitWidth === undefined && ps.fitHeight === undefined && ps.scale ? `${ps.scale}%` : undefined,
        'style:scale-to-X': ps.fitWidth !== undefined || ps.fitHeight !== undefined ? ps.fitWidth ?? 0 : undefined,
        'style:scale-to-Y': ps.fitWidth !== undefined || ps.fitHeight !== undefined ? ps.fitHeight ?? 0 : undefined,
        'style:print-page-order': ps.order === 'overThenDown' ? 'ltr' : 'ttb',
      })
      + tag('style:header-style', {}, hasHeader ? tag('style:header-footer-properties', { 'fo:min-height': `${area(m.top - m.header - 0.1).toFixed(3)}in`, 'fo:margin-bottom': '0.1in' }) : '')
      + tag('style:footer-style', {}, hasFooter ? tag('style:header-footer-properties', { 'fo:min-height': `${area(m.bottom - m.footer - 0.1).toFixed(3)}in`, 'fo:margin-top': '0.1in' }) : '')));
    const region = (code: string | undefined, tagName: string): string => {
      if (!code) return tag(tagName, { 'style:display': false });
      const parts: Record<'L' | 'C' | 'R', string> = { L: '', C: '', R: '' };
      let cur: 'L' | 'C' | 'R' = 'C';
      const re = /&([LCR])|&"[^"]*"|&\d+|&[BIUSXYKEPN&DTFAZG]|([^&]+)/g;
      let mm: RegExpExecArray | null;
      while ((mm = re.exec(code))) {
        if (mm[1]) cur = mm[1] as 'L' | 'C' | 'R';
        else if (mm[0] === '&P') parts[cur] += '\u0001';
        else if (mm[0] === '&N') parts[cur] += '\u0002';
        else if (mm[0] === '&A') parts[cur] += '\u0003';
        else if (mm[0] === '&D') parts[cur] += '\u0004';
        else if (mm[0] === '&&') parts[cur] += '&';
        else if (mm[2]) parts[cur] += mm[2];
      }
      const para = (t: string) => `<text:p>${esc(t).replace(/\u0001/g, '<text:page-number>1</text:page-number>').replace(/\u0002/g, '<text:page-count>1</text:page-count>').replace(/\u0003/g, '<text:sheet-name>???</text:sheet-name>').replace(/\u0004/g, '<text:date/>')}</text:p>`;
      return tag(tagName, {}, (['L', 'C', 'R'] as const).filter((k) => parts[k]).map((k) => tag(k === 'L' ? 'style:region-left' : k === 'C' ? 'style:region-center' : 'style:region-right', {}, para(parts[k]))).join(''));
    };
    masterPages.push(tag('style:master-page', { 'style:name': master, 'style:page-layout-name': `pm${si + 1}` }, region(ps.header, 'style:header') + region(ps.footer, 'style:footer')));
  }

  function tableXml(s: Sheet, si: number, taName: string): string {
    let maxR = -1, maxC = -1;
    for (const k of s.cells.keys()) { maxR = Math.max(maxR, keyRow(k)); maxC = Math.max(maxC, keyCol(k)); }
    for (const m of s.merges) { maxR = Math.max(maxR, m.r2); maxC = Math.max(maxC, m.c2); }
    for (const r of s.rows.keys()) maxR = Math.max(maxR, r);
    for (const c of s.cols.keys()) maxC = Math.max(maxC, c);
    for (const k of s.notes.keys()) { maxR = Math.max(maxR, keyRow(k)); maxC = Math.max(maxC, keyCol(k)); }
    for (const k of s.hyperlinks.keys()) { maxR = Math.max(maxR, keyRow(k)); maxC = Math.max(maxC, keyCol(k)); }
    for (const p of [...s.pictures, ...s.charts]) { maxR = Math.max(maxR, p.anchor.row); maxC = Math.max(maxC, p.anchor.col); }
    // validation ranges stay within the used area plus a margin
    const dvCells = new Map<number, string>();
    const limR = Math.min(MAX_ROW - 1, Math.max(maxR, 0) + 1000), limC = Math.min(MAX_COL - 1, Math.max(maxC, 0) + 50);
    for (const dv of s.validations) {
      const name = validationName.get(dv);
      if (!name) continue;
      for (const g of dv.ranges) {
        const r2 = Math.min(g.r2, limR), c2 = Math.min(g.c2, limC);
        for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { dvCells.set(r * MAX_COL + c, name); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c); }
      }
    }
    const covered = new Set<number>();
    const spans = new Map<number, CellRange>();
    for (const m of s.merges) {
      spans.set(m.r1 * MAX_COL + m.c1, m);
      for (let r = m.r1; r <= Math.min(m.r2, maxR); r++) for (let c = m.c1; c <= Math.min(m.c2, maxC); c++) if (r !== m.r1 || c !== m.c1) covered.add(r * MAX_COL + c);
    }
    const anchored = anchoredAt.get(s)!;

    // columns
    const cols: string[] = [];
    const defPt = colPt(s.defaultColWidth);
    const colAttrs = (c: number) => {
      const info = s.cols.get(c);
      return { style: colStyle(info?.width !== undefined ? colPt(info.width) : defPt), hidden: !!info?.hidden, cs: cellStyleName(info?.s) ?? 'Default' };
    };
    let c = 0;
    const lastCol = Math.max(maxC, 0);
    while (c <= lastCol) {
      const x = colAttrs(c);
      let n = 1;
      while (c + n <= lastCol) { const y = colAttrs(c + n); if (y.style !== x.style || y.hidden !== x.hidden || y.cs !== x.cs) break; n++; }
      cols.push(tag('table:table-column', { 'table:style-name': x.style, 'table:number-columns-repeated': n > 1 ? n : undefined, 'table:visibility': x.hidden ? 'collapse' : undefined, 'table:default-cell-style-name': x.cs }));
      c += n;
    }
    if (lastCol < MAX_COL - 1) cols.push(tag('table:table-column', { 'table:style-name': colStyle(defPt), 'table:number-columns-repeated': MAX_COL - 1 - lastCol, 'table:default-cell-style-name': 'Default' }));

    // rows
    const titles = /^\$?(\d+):\$?(\d+)$/.exec(s.pageSetup?.printTitleRows ?? '');
    const tr1 = titles ? +titles[1]! - 1 : -1, tr2 = titles ? +titles[2]! - 1 : -1;
    const breaks = new Set(s.pageSetup?.rowBreaks ?? []);
    const rowsOut: string[] = [];
    const defRowPt = s.defaultRowHeight;
    // columns with something to write, by row (cells, merges, validation, frames, notes, links)
    const byRow = new Map<number, Set<number>>();
    const mark = (k: number) => { const r = Math.floor(k / MAX_COL); let set = byRow.get(r); if (!set) byRow.set(r, (set = new Set())); set.add(k % MAX_COL); };
    for (const k of s.cells.keys()) mark(k);
    for (const k of covered) mark(k);
    for (const k of dvCells.keys()) mark(k);
    for (const k of spans.keys()) mark(k);
    for (const k of anchored.keys()) mark(k);
    for (const k of s.notes.keys()) mark(k);
    for (const k of s.hyperlinks.keys()) mark(k);
    const rowXml = (r: number): string => {
      const info = s.rows.get(r);
      const st = rowStyle(info?.height ?? defRowPt, info?.height !== undefined, breaks.has(r));
      const cellsXml: string[] = [];
      let col = 0;
      const present = [...(byRow.get(r) ?? [])].sort((a, b) => a - b);
      let prev = '';
      let repeat = 0;
      const flush = () => { if (!repeat) return; cellsXml.push(repeat > 1 ? prev.replace(/^<table:(table-cell|covered-table-cell)/, `<table:$1 table:number-columns-repeated="${repeat}"`) : prev); repeat = 0; prev = ''; };
      const push = (x: string, n = 1) => { if (x === prev && !x.includes('</')) { repeat += n; return; } flush(); prev = x; repeat = n; };
      for (const cc of present) {
        if (cc > col) push('<table:table-cell/>', cc - col);
        push(cellXml(s, r, cc, dvCells.get(r * MAX_COL + cc), covered.has(r * MAX_COL + cc), spans.get(r * MAX_COL + cc), anchored.get(r * MAX_COL + cc)));
        col = cc + 1;
      }
      flush();
      if (col < MAX_COL) cellsXml.push(tag('table:table-cell', { 'table:number-columns-repeated': MAX_COL - col }));
      return tag('table:table-row', { 'table:style-name': st, 'table:visibility': info?.hidden ? (info.filtered ? 'filter' : 'collapse') : undefined, 'table:default-cell-style-name': cellStyleName(info?.s) }, cellsXml.join(''));
    };
    const lastRow = Math.max(maxR, 0);
    let pendingRow = '', pendingCount = 0;
    const flushRows = () => { if (!pendingCount) return; rowsOut.push(pendingCount > 1 ? pendingRow.replace('<table:table-row', `<table:table-row table:number-rows-repeated="${pendingCount}"`) : pendingRow); pendingCount = 0; };
    const emit = (x: string, n: number) => { if (x === pendingRow) pendingCount += n; else { flushRows(); pendingRow = x; pendingCount = n; } };
    // rows that differ from an empty default row; runs between them are written once, repeated
    const special = [...new Set([...byRow.keys(), ...s.rows.keys(), ...breaks])].filter((x) => x <= lastRow).sort((a, b) => a - b);
    let r = 0;
    const emptyRow = (row: number) => rowXml(row);
    for (const sr of [...special, lastRow + 1]) {
      if (sr > r) {
        // empty rows r..sr-1 (the print-title block is written on its own)
        if (tr1 >= r && tr1 < sr) {
          if (tr1 > r) emit(emptyRow(r), tr1 - r);
          flushRows();
          const head: string[] = [];
          for (let h = tr1; h <= Math.min(tr2, MAX_ROW - 1); h++) head.push(rowXml(h));
          rowsOut.push(`<table:table-header-rows>${head.join('')}</table:table-header-rows>`);
          r = Math.min(tr2, MAX_ROW - 1) + 1;
          if (sr > r) emit(emptyRow(r), sr - r);
        } else emit(emptyRow(r), sr - r);
        r = Math.max(r, sr);
      }
      if (sr > lastRow || sr < r) continue;
      if (sr === tr1) {
        flushRows();
        const head: string[] = [];
        for (let h = tr1; h <= Math.min(tr2, MAX_ROW - 1); h++) head.push(rowXml(h));
        rowsOut.push(`<table:table-header-rows>${head.join('')}</table:table-header-rows>`);
        r = Math.min(tr2, MAX_ROW - 1) + 1;
        continue;
      }
      emit(rowXml(sr), 1);
      r = sr + 1;
    }
    flushRows();
    if (lastRow < MAX_ROW - 1) rowsOut.push(tag('table:table-row', { 'table:style-name': rowStyle(defRowPt, false), 'table:number-rows-repeated': MAX_ROW - 1 - lastRow }, tag('table:table-cell', { 'table:number-columns-repeated': MAX_COL })));

    const localNames = wb.names.filter((n) => n.sheet === si).map((n) => nameXml(n, s)).filter(Boolean).join('');
    const pr = s.pageSetup?.printArea?.split(',').map((p) => p.trim().replace(/\$/g, '')).map((p) => {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(p);
      if (!m) return '';
      const c1 = colNum(m[1]!), r1 = +m[2]! - 1, c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
      return rangeAddress(s.name, { r1, c1, r2, c2 });
    }).filter(Boolean).join(' ');
    return tag('table:table', {
      'table:name': s.name, 'table:style-name': taName, 'table:protected': s.protection?.sheet ? true : undefined,
      ...protectionKey(s.protection?.sheet ? s.protection.password : undefined),
      'table:print-ranges': pr || undefined,
    }, cols.join('') + rowsOut.join('') + cfXml(s) + (localNames ? `<table:named-expressions>${localNames}</table:named-expressions>` : ''));
  }

  function cellXml(s: Sheet, r: number, c: number, dv: string | undefined, isCovered: boolean, span: CellRange | undefined, frames: string[] | undefined): string {
    const cell: Cell | undefined = s.cells.get(r * 16384 + c);
    const style = cellStyleName(cell?.s);
    const at: Record<string, string | number | boolean | undefined> = { 'table:style-name': style, 'table:content-validation-name': dv };
    if (span) { at['table:number-columns-spanned'] = span.c2 - span.c1 + 1; at['table:number-rows-spanned'] = span.r2 - span.r1 + 1; }
    let inner = '';
    const note = s.notes.get(r * 16384 + c);
    if (note) inner += tag('office:annotation', {}, (note.author ? `<dc:creator>${esc(note.author)}</dc:creator>` : '') + note.text.split('\n').map((l) => `<text:p>${textXml(l)}</text:p>`).join(''));
    if (frames) inner += frames.join('');
    if (cell) {
      if (cell.f) {
        const f = odfFormula(cell.f, s);
        if (f) {
          at['table:formula'] = f;
          if (cell.arrayRange) { at['table:number-matrix-columns-spanned'] = cell.arrayRange.c2 - cell.arrayRange.c1 + 1; at['table:number-matrix-rows-spanned'] = cell.arrayRange.r2 - cell.arrayRange.r1 + 1; }
        }
      }
      const v = cell.v;
      const code = wb.styles[cell.s ?? 0]?.numFmt;
      let shown = '';
      if (typeof v === 'number') {
        const t = valueTypeOf(code);
        at['office:value-type'] = t;
        at['calcext:value-type'] = t;
        if (t === 'date') at['office:date-value'] = isoDate(v, wb.date1904);
        else if (t === 'time') at['office:time-value'] = duration(v);
        else at['office:value'] = String(v);
        if (t === 'currency') at['office:currency'] = currencyCode(code ?? '');
        try { shown = formatValue(v, code ?? 'General', { date1904: wb.date1904 }).text; } catch { shown = String(v); }
      } else if (typeof v === 'boolean') {
        at['office:value-type'] = 'boolean'; at['calcext:value-type'] = 'boolean'; at['office:boolean-value'] = v;
        shown = v ? 'TRUE' : 'FALSE';
      } else if (v instanceof CellError) {
        at['office:value-type'] = 'string'; at['calcext:value-type'] = 'error'; at['office:string-value'] = '';
        shown = v.code;
      } else if (typeof v === 'string') {
        at['office:value-type'] = 'string'; at['calcext:value-type'] = 'string';
        if (cell.f) at['office:string-value'] = v;
        shown = v;
      }
      const link = s.hyperlinks.get(r * 16384 + c);
      const href = link ? (link.startsWith('#') ? `#${odfLocation(link.slice(1))}` : link) : undefined;
      let body: string;
      if (cell.rich?.length && typeof v === 'string') {
        // split runs into paragraphs at line breaks
        const paras: string[] = [''];
        for (const run of cell.rich) {
          run.text.split('\n').forEach((part, i) => {
            if (i) paras.push('');
            if (!part) return;
            const ts = textStyle(run.font);
            paras[paras.length - 1] += ts ? `<text:span text:style-name="${ts}">${textXml(part)}</text:span>` : textXml(part);
          });
        }
        body = paras.map((p) => `<text:p>${href ? `<text:a xlink:href="${esc(href)}" xlink:type="simple">${p}</text:a>` : p}</text:p>`).join('');
      } else if (shown !== '' || v !== null) {
        body = shown.split('\n').map((l) => `<text:p>${href ? `<text:a xlink:href="${esc(href)}" xlink:type="simple">${textXml(l)}</text:a>` : textXml(l)}</text:p>`).join('');
      } else body = '';
      inner += body;
    } else {
      const link = s.hyperlinks.get(r * 16384 + c);
      if (link) {
        const href = link.startsWith('#') ? `#${odfLocation(link.slice(1))}` : link;
        at['office:value-type'] = 'string';
        inner += `<text:p><text:a xlink:href="${esc(href)}" xlink:type="simple">${textXml(link)}</text:a></text:p>`;
      }
    }
    return tag(isCovered ? 'table:covered-table-cell' : 'table:table-cell', at, inner);
  }

  function odfLocation(loc: string): string {
    try {
      const n = parseFormula(loc);
      if (n.type === 'ref') return odfRef(n.ref, undefined, true).replace(/\$/g, '').replace(/^\./, '');
    } catch { /* keep as is */ }
    return loc;
  }

  function cfXml(s: Sheet): string {
    const out: string[] = [];
    for (const cf of s.conditionalFormats) {
      if (!cf.ranges.length) continue;
      const base = cf.ranges[0]!;
      const baseAddr = rangeAddress(s.name, { r1: base.r1, c1: base.c1, r2: base.r1, c2: base.c1 });
      const items = cf.rules.map((rule) => cfRuleXml(rule, s, baseAddr)).filter(Boolean).join('');
      if (!items) continue;
      out.push(tag('calcext:conditional-format', { 'calcext:target-range-address': cf.ranges.map((g) => rangeAddress(s.name, g)).join(' ') }, items));
    }
    return out.length ? `<calcext:conditional-formats>${out.join('')}</calcext:conditional-formats>` : '';
  }

  function cfValueAttrs(v: CfValue | undefined): Record<string, string | undefined> {
    if (!v) return { 'calcext:type': 'minimum' };
    switch (v.type) {
      case 'min': return { 'calcext:type': 'minimum' };
      case 'max': return { 'calcext:type': 'maximum' };
      case 'autoMin': return { 'calcext:type': 'auto-minimum' };
      case 'autoMax': return { 'calcext:type': 'auto-maximum' };
      case 'percent': return { 'calcext:type': 'percent', 'calcext:value': v.value };
      case 'percentile': return { 'calcext:type': 'percentile', 'calcext:value': v.value };
      case 'formula': return { 'calcext:type': 'formula', 'calcext:value': v.value };
      default: return { 'calcext:type': 'number', 'calcext:value': v.value };
    }
  }

  function cfRuleXml(rule: CfRule, s: Sheet, base: string): string {
    const cond = (value: string, st: CellStyle) => tag('calcext:condition', { 'calcext:apply-style-name': cfStyle(st), 'calcext:value': value, 'calcext:base-cell-address': base });
    const ex = (f: string) => odfBare(f, s);
    switch (rule.type) {
      case 'cellIs': {
        const f = rule.formulas.map(ex);
        if (f.some((x) => x === undefined)) return '';
        const ops: Record<string, string> = { equal: '=', notEqual: '!=', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=' };
        if (rule.operator === 'between' || rule.operator === 'notBetween') return cond(`${rule.operator === 'between' ? 'between' : 'not-between'}(${f[0]},${f[1] ?? f[0]})`, rule.style);
        return cond(`${ops[rule.operator]}${f[0]}`, rule.style);
      }
      case 'expression': { const f = ex(rule.formula); return f ? cond(`formula-is(${f})`, rule.style) : ''; }
      case 'containsText': return cond(`contains-text("${rule.text.replace(/"/g, '""')}")`, rule.style);
      case 'notContainsText': return cond(`not-contains-text("${rule.text.replace(/"/g, '""')}")`, rule.style);
      case 'beginsWith': return cond(`begins-with("${rule.text.replace(/"/g, '""')}")`, rule.style);
      case 'endsWith': return cond(`ends-with("${rule.text.replace(/"/g, '""')}")`, rule.style);
      case 'duplicateValues': return cond('duplicate', rule.style);
      case 'uniqueValues': return cond('unique', rule.style);
      case 'containsErrors': return cond('is-error', rule.style);
      case 'notContainsErrors': return cond('is-no-error', rule.style);
      case 'containsBlanks': return cond('is-blank', rule.style);
      case 'notContainsBlanks': return cond('is-no-blank', rule.style);
      case 'top10': return cond(`${rule.bottom ? 'bottom' : 'top'}-${rule.percent ? 'percent' : 'elements'}(${rule.rank})`, rule.style);
      case 'aboveAverage': return cond(`${rule.below ? 'below' : 'above'}${rule.equal ? '-equal' : ''}-average`, rule.style);
      case 'timePeriod': {
        const map: Record<string, string> = { today: 'today', yesterday: 'yesterday', tomorrow: 'tomorrow', last7Days: 'last-7-days', thisWeek: 'this-week', lastWeek: 'last-week', nextWeek: 'next-week', thisMonth: 'this-month', lastMonth: 'last-month', nextMonth: 'next-month' };
        return tag('calcext:date-is', { 'calcext:date': map[rule.period], 'calcext:style': cfStyle(rule.style) });
      }
      case 'colorScale': return tag('calcext:color-scale', {}, rule.stops.map((st) => tag('calcext:color-scale-entry', { ...cfValueAttrs(st.value), 'calcext:color': st.color.slice(0, 7).toLowerCase() })).join(''));
      case 'dataBar': return tag('calcext:data-bar', { 'calcext:positive-color': rule.color.slice(0, 7).toLowerCase(), 'calcext:negative-color': '#ff0000', 'calcext:gradient': rule.gradient !== false, 'calcext:show-value': rule.showValue !== false },
        tag('calcext:formatting-entry', cfValueAttrs(rule.min)) + tag('calcext:formatting-entry', cfValueAttrs(rule.max)));
      case 'iconSet': return tag('calcext:icon-set', { 'calcext:icon-set-type': rule.set, 'calcext:show-value': rule.showValue !== false }, rule.values.map((v) => tag('calcext:formatting-entry', cfValueAttrs(v))).join(''));
      default: return '';
    }
  }

  function nameXml(n: Workbook['names'][number], s?: Sheet): string {
    if (/^_xlnm\./i.test(n.name)) return '';
    const owner = s ?? wb.sheets[0]!;
    try {
      const node = parseFormula(n.formula);
      const refs = node.type === 'ref' ? [node.ref] : node.type === 'union' && node.items.every((x) => x.type === 'ref') ? node.items.map((x) => (x as Extract<Node, { type: 'ref' }>).ref) : null;
      if (refs && refs.every((x) => x.sheet !== undefined)) {
        return tag('table:named-range', { 'table:name': n.name, 'table:base-cell-address': absAddress(owner.name, { r1: 0, c1: 0, r2: 0, c2: 0 }), 'table:cell-range-address': refs.map((x) => odfRef(x, undefined, true)).join(' '), 'clerdo:hidden': n.hidden ? true : undefined });
      }
    } catch { return ''; }
    const f = odfFormula(n.formula, owner);
    return f ? tag('table:named-expression', { 'table:name': n.name, 'table:base-cell-address': absAddress(owner.name, { r1: 0, c1: 0, r2: 0, c2: 0 }), 'table:expression': f, 'clerdo:hidden': n.hidden ? true : undefined }) : '';
  }

  const dbRanges = wb.sheets.map((s, i) => {
    const af = s.autoFilter;
    if (!af) return '';
    const conds: string[] = [];
    for (const fc of af.columns) {
      const field = fc.col - af.range.c1;
      if (field < 0) continue;
      if (fc.values?.length) {
        conds.push(tag('table:filter-condition', { 'table:field-number': field, 'table:value': fc.values[0], 'table:operator': '=' }, fc.values.map((v) => tag('table:filter-set-item', { 'table:value': v })).join('')));
      } else if (fc.top10) {
        conds.push(tag('table:filter-condition', { 'table:field-number': field, 'table:value': fc.top10.value, 'table:operator': `${fc.top10.top ? 'top' : 'bottom'} ${fc.top10.percent ? 'percent' : 'values'}`, 'table:data-type': 'number' }));
      } else if (fc.blanks) conds.push(tag('table:filter-condition', { 'table:field-number': field, 'table:value': '', 'table:operator': 'empty' }));
      else if (fc.custom) {
        const ops: Record<string, string> = { equal: '=', notEqual: '!=', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=' };
        const items = fc.custom.rules.map((r) => tag('table:filter-condition', { 'table:field-number': field, 'table:value': r.value, 'table:operator': ops[r.operator] ?? '=', 'table:data-type': r.value !== '' && Number.isFinite(Number(r.value)) ? 'number' : 'text' }));
        conds.push(items.length > 1 ? tag(fc.custom.and ? 'table:filter-and' : 'table:filter-or', {}, items.join('')) : items.join(''));
      }
    }
    return tag('table:database-range', { 'table:name': `__Anonymous_Sheet_DB__${i}`, 'table:target-range-address': rangeAddress(s.name, af.range), 'table:display-filter-buttons': true },
      conds.length ? `<table:filter>${conds.length > 1 ? `<table:filter-and>${conds.join('')}</table:filter-and>` : conds[0]}</table:filter>` : '');
  }).join('');

  const globalNames = wb.names.filter((n) => n.sheet === undefined).map((n) => nameXml(n)).filter(Boolean).join('');
  const calcSettings = tag('table:calculation-settings', { 'table:automatic-find-labels': false, 'table:use-regular-expressions': false, 'table:use-wildcards': true },
    (wb.date1904 ? tag('table:null-date', { 'table:date-value': '1904-01-01' }) : '')
    + (wb.calc?.iterate ? tag('table:iteration', { 'table:status': 'enable', 'table:steps': wb.calc.maxIterations ?? 100, 'table:minimum-difference': wb.calc.maxChange ?? 0.001 }) : ''));

  autoStyles.push(tag('style:style', { 'style:name': 'gr1', 'style:family': 'graphic' }, tag('style:graphic-properties', { 'draw:stroke': 'none', 'draw:fill': 'none', 'style:protect': 'size' })));
  const fontDecls = () => `<office:font-face-decls>${[...fonts].map((f) => tag('style:font-face', { 'style:name': f, 'svg:font-family': /\s/.test(f) ? `'${f}'` : f })).join('')}</office:font-face-decls>`;
  const body = calcSettings + (validationXml.length ? `<table:content-validations>${validationXml.join('')}</table:content-validations>` : '') + tables.join('')
    + (globalNames ? `<table:named-expressions>${globalNames}</table:named-expressions>` : '') + (dbRanges ? `<table:database-ranges>${dbRanges}</table:database-ranges>` : '');
  const wbProtection = () => {
    if (!wb.protection?.structure) return '';
    const k = protectionKey(wb.protection.password);
    return ' table:structure-protected="true"' + Object.entries(k).map(([a2, v]) => ` ${a2}="${esc(v)}"`).join('');
  };
  const docNs = NS_DECL('office', 'style', 'text', 'table', 'draw', 'fo', 'xlink', 'dc', 'meta', 'number', 'svg', 'chart', 'of', 'calcext', 'loext', 'tableooo', 'clerdo');
  const contentXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content${docNs} office:version="1.3">${fontDecls()}`
    + `<office:automatic-styles>${tableStyles.join('')}${autoStyles.join('')}</office:automatic-styles><office:body><office:spreadsheet${wbProtection()}>${body}</office:spreadsheet></office:body></office:document-content>`;

  const def = wb.styles[0] ?? {};
  const defaultStyle = tag('style:default-style', { 'style:family': 'table-cell' }, tag('style:paragraph-properties', { 'style:tab-stop-distance': '0.5in' }) + textProps({ name: DEFAULT_FONT.name, size: DEFAULT_FONT.size, ...def.font }, fonts));
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles${docNs} office:version="1.3">${fontDecls()}<office:styles>${defaultStyle}`
    + tag('style:style', { 'style:name': 'Default', 'style:family': 'table-cell' }) + commonStyles.join('')
    + `</office:styles><office:automatic-styles>${pageLayouts.join('')}</office:automatic-styles><office:master-styles>${masterPages.join('')}</office:master-styles></office:document-styles>`;

  const item = (name: string, type: string, value: string | number | boolean) => tag('config:config-item', { 'config:name': name, 'config:type': type }, esc(String(value)));
  const tablesCfg = wb.sheets.map((s) => {
    const v = s.view;
    const sel = v.selection ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    return tag('config:config-item-map-entry', { 'config:name': s.name },
      item('CursorPositionX', 'int', sel.c1) + item('CursorPositionY', 'int', sel.r1)
      + item('HorizontalSplitMode', 'short', v.frozenCols ? 2 : 0) + item('VerticalSplitMode', 'short', v.frozenRows ? 2 : 0)
      + item('HorizontalSplitPosition', 'int', v.frozenCols) + item('VerticalSplitPosition', 'int', v.frozenRows)
      + item('ActiveSplitRange', 'short', v.frozenRows || v.frozenCols ? 2 : 2)
      + item('PositionLeft', 'int', 0) + item('PositionRight', 'int', v.frozenCols) + item('PositionTop', 'int', 0) + item('PositionBottom', 'int', v.frozenRows)
      + item('ZoomType', 'short', 0) + item('ZoomValue', 'int', v.zoom || 100) + item('PageViewZoomValue', 'int', 60)
      + item('ShowGrid', 'boolean', v.showGrid) + item('HasColumnRowHeaders', 'boolean', v.showHeaders));
  }).join('');
  const active = wb.sheets[wb.activeSheet] ?? wb.sheets[0];
  const settingsXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings${NS_DECL('office', 'config', 'xlink')} office:version="1.3"><office:settings>`
    + `<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>`
    + item('ViewId', 'string', 'view1') + `<config:config-item-map-named config:name="Tables">${tablesCfg}</config:config-item-map-named>`
    + item('ActiveTable', 'string', active?.name ?? '') + item('ShowGrid', 'boolean', active?.view.showGrid ?? true) + item('HasColumnRowHeaders', 'boolean', active?.view.showHeaders ?? true)
    + `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`;

  const now = new Date().toISOString().replace(/\.\d+Z$/, '');
  const p = wb.props ?? {};
  const metaXml = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta${NS_DECL('office', 'meta', 'dc', 'xlink')} office:version="1.3"><office:meta>`
    + `<meta:generator>ClerdoGrid</meta:generator>${p.title ? `<dc:title>${esc(p.title)}</dc:title>` : ''}${p.creator ? `<meta:initial-creator>${esc(p.creator)}</meta:initial-creator><dc:creator>${esc(p.creator)}</dc:creator>` : ''}`
    + `<meta:creation-date>${esc((p.created ?? now).replace(/Z$/, ''))}</meta:creation-date><dc:date>${esc((p.modified ?? now).replace(/Z$/, ''))}</dc:date>`
    + `<meta:document-statistic meta:table-count="${wb.sheets.length}"/></office:meta></office:document-meta>`;

  if (wb.queries?.length) { files['clerdo/queries.xml'] = strToU8(queriesToXml(wb.queries)); manifest.push(['clerdo/queries.xml', 'text/xml']); }
  const mime = opts.template ? MIME_TEMPLATE : MIME;
  const plain: [string, Uint8Array, string][] = [
    ['content.xml', strToU8(contentXml), 'text/xml'], ['styles.xml', strToU8(stylesXml), 'text/xml'], ['settings.xml', strToU8(settingsXml), 'text/xml'], ['meta.xml', strToU8(metaXml), 'text/xml'],
    ...manifest.filter(([f]) => !f.endsWith('/')).map(([f, t]): [string, Uint8Array, string] => { const v = files[f]!; return [f, Array.isArray(v) ? (v[0] as Uint8Array) : (v as Uint8Array), t]; }),
  ];
  const dirs = manifest.filter(([f]) => f.endsWith('/'));
  const key = opts.password ? odfKey(opts.password) : null;
  const out: Zippable = { mimetype: [strToU8(mime), { level: 0 }] };
  const entries: string[] = [];
  const stored: [string, Zippable[string]][] = [];
  for (const [f, data, t] of plain) {
    if (key) {
      const e = encryptStream(data, key);
      stored.push([f, [e.data, { level: 0 }]]);
      entries.push(`<manifest:file-entry manifest:full-path="${esc(f)}" manifest:media-type="${esc(t)}" manifest:size="${data.length}">${e.encryption}</manifest:file-entry>`);
    } else {
      stored.push([f, /^Pictures\//.test(f) ? [data, { level: 0 }] : data]);
      entries.push(tag('manifest:file-entry', { 'manifest:full-path': f, 'manifest:media-type': t }));
    }
  }
  for (const [d, t] of dirs) entries.push(tag('manifest:file-entry', { 'manifest:full-path': d, 'manifest:media-type': t }));
  const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">`
    + tag('manifest:file-entry', { 'manifest:full-path': '/', 'manifest:version': '1.3', 'manifest:media-type': mime }) + entries.join('') + '</manifest:manifest>';
  out['META-INF/manifest.xml'] = strToU8(manifestXml);
  for (const [f, v] of stored) out[f] = v;
  return zipSync(out, { level: 6 });
};

/** The legacy 16-bit sheet password hash as an ODF protection key other apps accept. */
export const protectionKey = (hex: string | undefined): Record<string, string> => {
  if (!hex || !/^[0-9A-F]{1,4}$/i.test(hex)) return {};
  const n = parseInt(hex, 16);
  return {
    'table:protection-key': toBase64(sha1(new Uint8Array([n >> 8, n & 0xff]))),
    'table:protection-key-digest-algorithm': 'http://docs.oasis-open.org/office/ns/table/legacy-hash-excel',
    'loext:protection-key-digest-algorithm-2': 'http://www.w3.org/2000/09/xmldsig#sha1',
  };
};

const colNum = (s: string): number => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

const isoDate = (serial: number, date1904: boolean): string => {
  const d = fromSerial(Math.floor(serial), date1904);
  const t = timeOf(serial - Math.floor(serial));
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  const base = `${p(d.year, 4)}-${p(d.month)}-${p(d.day)}`;
  if (serial === Math.floor(serial)) return base;
  return `${base}T${p(t.hour)}:${p(t.minute)}:${p(t.second)}${t.ms ? '.' + p(t.ms, 3) : ''}`;
};

const duration = (days: number): string => {
  const neg = days < 0;
  let secs = Math.abs(days) * 86400;
  const h = Math.floor(secs / 3600); secs -= h * 3600;
  const m = Math.floor(secs / 60); secs -= m * 60;
  const s = Math.round(secs * 1000) / 1000;
  return `${neg ? '-' : ''}PT${h}H${String(m).padStart(2, '0')}M${String(s).padStart(2, '0')}S`;
};
