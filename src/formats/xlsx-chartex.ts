// Extended chart parts (cx: namespace http://schemas.microsoft.com/office/drawing/2014/chartex) for the
// newer chart kinds, the drawing markup that offers them next to a classic fallback chart, and chart sheets.
import { cellKey, parseRange, type Chart, type ChartOptions, type ChartSeries, type ChartType, type Sheet, type Workbook } from '../model/workbook';
import type { Scalar } from '../engine/types';
import { attr, esc, find, findAll, kid, kids, numAttr, type XNode } from './xml';
import { CLERDO_EXT_URI, EMU_PER_PX, EXTENDED, anchorXml, chartFrameXml, chartXml, readAnchor, readChart } from './xlsx-chart';

export const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const CHARTEX_CT = 'application/vnd.ms-office.chartex+xml';
export const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const CHARTSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** Namespace a reader must know to draw each kind (newer kinds came in later revisions). */
const REQUIRES: Partial<Record<ChartType, [string, string]>> = {
  funnel: ['cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex'],
  regionMap: ['cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex'],
};
const CX1: [string, string] = ['cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'];

const LAYOUT: Partial<Record<ChartType, string>> = {
  histogram: 'clusteredColumn', pareto: 'clusteredColumn', boxWhisker: 'boxWhisker', waterfall: 'waterfall', funnel: 'funnel',
  treemap: 'treemap', sunburst: 'sunburst', regionMap: 'regionMap',
};

export const isExtended = (t: ChartType): boolean => EXTENDED.includes(t);

// ---------------------------------------------------------------- cached values

export type CellLookup = (ref: string) => { values: Scalar[]; cols: number } | null;

const REF = /^(?:'((?:[^']|'')+)'|([^!'\s:]+))!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)$/;

/** Reads the stored values of a sheet range (for the value caches other apps show before recalculating). */
export const cellLookup = (wb: Workbook): CellLookup => (ref) => {
  const m = REF.exec(ref.trim().replace(/^=/, ''));
  if (!m) return null;
  const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]!;
  const sheet = wb.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
  const r = parseRange(m[3]!.replace(/\$/g, ''));
  if (!sheet || !r) return null;
  const values: Scalar[] = [];
  const rows = Math.min(r.r2, r.r1 + 20000), cols = Math.min(r.c2, r.c1 + 64);
  for (let row = r.r1; row <= rows; row++) for (let col = r.c1; col <= cols; col++) values.push(sheet.cells.get(cellKey(row, col))?.v ?? null);
  return { values, cols: cols - r.c1 + 1 };
};

// ---------------------------------------------------------------- writing

const txData = (text: string | undefined, look: CellLookup): string => {
  if (!text) return '';
  const t = text.trim();
  if (/^=/.test(t) || REF.test(t)) {
    const ref = t.replace(/^=/, '');
    const cached = look(ref)?.values.map((x) => (x === null ? '' : String(typeof x === 'object' ? '' : x))).join(' ') ?? '';
    return `<cx:tx><cx:txData><cx:f>${esc(ref)}</cx:f><cx:v>${esc(cached)}</cx:v></cx:txData></cx:tx>`;
  }
  return `<cx:tx><cx:txData><cx:v>${esc(text)}</cx:v></cx:txData></cx:tx>`;
};

const numDim = (type: string, ref: string, look: CellLookup): string => {
  const c = look(ref);
  const pts = (c?.values ?? []).map((x, i) => (typeof x === 'number' && Number.isFinite(x) ? `<cx:pt idx="${i}">${x}</cx:pt>` : '')).join('');
  return `<cx:numDim type="${type}"><cx:f>${esc(ref.replace(/^=/, ''))}</cx:f><cx:lvl ptCount="${c?.values.length ?? 0}" formatCode="General">${pts}</cx:lvl></cx:numDim>`;
};

/** Category text; a range several columns wide is a hierarchy, written innermost level first. */
const strDim = (ref: string, look: CellLookup): string => {
  const c = look(ref);
  const cols = c?.cols ?? 1;
  const n = c ? Math.ceil(c.values.length / cols) : 0;
  const text = (x: Scalar) => (x === null ? '' : typeof x === 'object' ? '' : String(x));
  const lvls: string[] = [];
  for (let k = cols - 1; k >= 0; k--) {
    let pts = '';
    for (let i = 0; i < n; i++) { const t = text(c!.values[i * cols + k] ?? null); if (t) pts += `<cx:pt idx="${i}">${esc(t)}</cx:pt>`; }
    lvls.push(`<cx:lvl ptCount="${n}">${pts}</cx:lvl>`);
  }
  return `<cx:strDim type="cat"><cx:f>${esc(ref.replace(/^=/, ''))}</cx:f>${lvls.join('')}</cx:strDim>`;
};

const hexOf = (s: string | undefined): string | undefined => (s && /^#?[0-9a-fA-F]{6}$/.test(s) ? s.replace('#', '').toUpperCase() : undefined);
const uid = (k: number): string => `{${(0x10000000 + k).toString(16).toUpperCase()}-C1E0-4D2B-8A3C-0000000000${String(k).padStart(2, '0').slice(-2)}}`;

export const chartExXml = (c: Chart, look: CellLookup): string => {
  const layout = LAYOUT[c.type] ?? 'clusteredColumn';
  const o = c.options ?? {};
  const hierarchy = c.type === 'treemap' || c.type === 'sunburst';
  const map = c.type === 'regionMap';
  const valueType = hierarchy ? 'size' : map ? 'colorVal' : 'val';
  const list = c.type === 'histogram' || c.type === 'boxWhisker' ? c.series : c.series.slice(0, 1);
  const data = list.map((s, i) => `<cx:data id="${i}">${s.categories ? strDim(s.categories, look) : ''}${numDim(valueType, s.values, look)}</cx:data>`).join('');
  const labels = c.dataLabels || (c.dataLabels === undefined && (hierarchy || c.type === 'funnel'))
    ? `<cx:dataLabels pos="${hierarchy ? 'inEnd' : map ? 'ctr' : 'outEnd'}"><cx:visibility seriesName="0" categoryName="${hierarchy ? 1 : 0}" value="1"/></cx:dataLabels>` : '';
  const layoutPr = (s: ChartSeries): string => {
    void s;
    switch (c.type) {
      case 'histogram': {
        const bin = o.binWidth ? `<cx:binSize val="${o.binWidth}"/>` : o.binCount ? `<cx:binCount val="${o.binCount}"/>` : '';
        return `<cx:layoutPr><cx:binning intervalClosed="r">${bin}</cx:binning></cx:layoutPr>`;
      }
      case 'pareto': return '<cx:layoutPr><cx:aggregation/></cx:layoutPr>';
      case 'boxWhisker': return `<cx:layoutPr><cx:visibility meanLine="0" meanMarker="${o.hideMean ? 0 : 1}" nonoutliers="0" outliers="${o.hideOutliers ? 0 : 1}"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr>`;
      case 'waterfall': return `<cx:layoutPr><cx:visibility connectorLines="1"/>${o.totals?.length ? `<cx:subtotals>${o.totals.map((i) => `<cx:idx val="${i}"/>`).join('')}</cx:subtotals>` : ''}</cx:layoutPr>`;
      case 'treemap': return '<cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>';
      case 'regionMap': {
        const lbl = o.map?.labels === 'all' ? 'showAll' : o.map?.labels === 'best' ? 'bestFitOnly' : 'none';
        return `<cx:layoutPr><cx:regionLabelLayout val="${lbl}"/><cx:geography cultureLanguage="en-US" cultureRegion="US" attribution=""/></cx:layoutPr>`;
      }
      default: return '';
    }
  };
  const colors = map ? (() => {
    const m = o.map ?? {};
    const lo = hexOf(m.low), mid = hexOf(m.mid), hi = hexOf(m.high);
    if (!lo && !hi && !m.diverging) return '';
    const cc = (h: string | undefined, d: string) => `<a:srgbClr val="${h ?? d}"/>`;
    return `<cx:valueColors><cx:minColor>${cc(lo, m.diverging ? 'C94C3C' : 'D9F0E3')}</cx:minColor>${m.diverging ? `<cx:midColor>${cc(mid, 'F5F1E4')}</cx:midColor>` : ''}<cx:maxColor>${cc(hi, '0C7A4B')}</cx:maxColor></cx:valueColors><cx:valueColorPositions count="${m.diverging ? 3 : 2}"/>`;
  })() : '';
  const color = (s: ChartSeries) => (hexOf(s.color) ? `<cx:spPr><a:solidFill><a:srgbClr val="${hexOf(s.color)}"/></a:solidFill></cx:spPr>` : '');
  let series = list.map((s, i) => `<cx:series layoutId="${layout}" uniqueId="${uid(i)}">${txData(s.name, look)}${color(s)}${colors}${labels}<cx:dataId val="${i}"/>${layoutPr(s)}${c.type === 'pareto' ? '<cx:axisId val="1"/>' : ''}</cx:series>`).join('');
  if (c.type === 'pareto') series += `<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="${uid(90)}"><cx:axisId val="2"/></cx:series>`;
  const axes = hierarchy || map ? ''
    : c.type === 'funnel' ? '<cx:axis id="0"><cx:catScaling gapWidth="0.06"/><cx:tickLabels/></cx:axis>'
      : `<cx:axis id="0"><cx:catScaling gapWidth="${c.type === 'histogram' ? 0 : 0.5}"/>${c.axisTitles?.x ? `<cx:title>${txData(c.axisTitles.x, look)}</cx:title>` : ''}<cx:tickLabels/></cx:axis>`
      + `<cx:axis id="1"><cx:valScaling${o.max !== undefined ? ` max="${o.max}"` : ''}${o.min !== undefined ? ` min="${o.min}"` : ''}/>${c.axisTitles?.y ? `<cx:title>${txData(c.axisTitles.y, look)}</cx:title>` : ''}${o.gridlines === false ? '' : '<cx:majorGridlines/>'}<cx:tickLabels/></cx:axis>`
      + (c.type === 'pareto' ? '<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/><cx:tickLabels/></cx:axis>' : '');
  const legend = c.legend && c.legend !== 'none' ? `<cx:legend pos="${{ right: 'r', bottom: 'b', top: 't', left: 'l' }[c.legend]}" align="ctr" overlay="0"/>` : '';
  const title = c.title ? `<cx:title pos="t" align="ctr" overlay="0">${txData(c.title, look)}</cx:title>` : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${NS_CX}">`
    + `<cx:chartData>${data}</cx:chartData><cx:chart>${title}<cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${axes}</cx:plotArea>${legend}</cx:chart></cx:chartSpace>`;
};

// ---------------------------------------------------------------- reading

const LAYOUT_IN: Record<string, ChartType> = { boxWhisker: 'boxWhisker', waterfall: 'waterfall', funnel: 'funnel', treemap: 'treemap', sunburst: 'sunburst', regionMap: 'regionMap' };
const LEGEND_IN: Record<string, Chart['legend']> = { r: 'right', b: 'bottom', t: 'top', l: 'left' };

const textOf = (tx: XNode | undefined): string | undefined => {
  if (!tx) return undefined;
  const f = find(tx, 'f');
  if (f?.text.trim()) return '=' + f.text.trim();
  const v = find(tx, 'v');
  if (v?.text) return v.text;
  const t = findAll(tx, 't').map((x) => x.text).join('');
  return t || undefined;
};

/** A chart from an extended chart part written by another app. */
export const readChartEx = (root: XNode, id: string, anchor: Chart['anchor']): Chart | undefined => {
  if (root.name !== 'chartSpace') return undefined;
  const chart = kid(root, 'chart');
  const region = kid(kid(chart, 'plotArea'), 'plotAreaRegion');
  const all = kids(region, 'series');
  const main = all.filter((s) => attr(s, 'layoutId') !== 'paretoLine');
  if (!main.length) return undefined;
  const layout = attr(main[0], 'layoutId') ?? '';
  const pareto = all.some((s) => attr(s, 'layoutId') === 'paretoLine');
  const type: ChartType | undefined = pareto ? 'pareto' : layout === 'clusteredColumn' ? 'histogram' : LAYOUT_IN[layout];
  if (!type) return undefined;
  const dataById = new Map(kids(kid(root, 'chartData'), 'data').map((d) => [attr(d, 'id') ?? '', d]));
  const o: ChartOptions = {};
  const series: ChartSeries[] = [];
  for (const s of main) {
    const d = dataById.get(attr(kid(s, 'dataId'), 'val') ?? '');
    const nd = kids(d, 'numDim')[0], sd = kids(d, 'strDim')[0];
    const values = find(nd, 'f')?.text.trim();
    if (!values) continue;
    const ser: ChartSeries = { values };
    const cats = find(sd, 'f')?.text.trim();
    if (cats) ser.categories = cats;
    const name = textOf(kid(s, 'tx'));
    if (name) ser.name = name.replace(/^=/, '');
    const col = attr(find(kid(s, 'spPr'), 'srgbClr'), 'val');
    if (col && /^[0-9A-Fa-f]{6}$/.test(col)) ser.color = '#' + col.toUpperCase();
    series.push(ser);
    const lp = kid(s, 'layoutPr');
    const bin = kid(lp, 'binning');
    const bc = numAttr(kid(bin, 'binCount'), 'val'), bs = numAttr(kid(bin, 'binSize'), 'val');
    if (bc && bc >= 1) o.binCount = Math.min(1000, Math.round(bc));
    if (bs && bs > 0) o.binWidth = bs;
    const vis = kid(lp, 'visibility');
    if (attr(vis, 'meanMarker') === '0') o.hideMean = true;
    if (attr(vis, 'outliers') === '0') o.hideOutliers = true;
    const subs = kids(kid(lp, 'subtotals'), 'idx').map((x) => numAttr(x, 'val')).filter((x): x is number => x !== undefined && Number.isInteger(x) && x >= 0);
    if (subs.length) o.totals = subs.slice(0, 4096);
    const rl = attr(kid(lp, 'regionLabelLayout'), 'val');
    const vc = kid(s, 'valueColors');
    if (type === 'regionMap') {
      const m: NonNullable<ChartOptions['map']> = {};
      if (rl === 'showAll') m.labels = 'all'; else if (rl === 'bestFitOnly') m.labels = 'best';
      const cc = (n: string) => { const v = attr(find(kid(vc, n), 'srgbClr'), 'val'); return v && /^[0-9A-Fa-f]{6}$/.test(v) ? '#' + v.toUpperCase() : undefined; };
      const lo = cc('minColor'), mid = cc('midColor'), hi = cc('maxColor');
      if (lo) m.low = lo; if (hi) m.high = hi;
      if (mid) { m.mid = mid; m.diverging = true; }
      if (Object.keys(m).length) o.map = m;
    }
  }
  if (!series.length) return undefined;
  const c: Chart = { id, type, series, anchor };
  const title = textOf(kid(kid(chart, 'title'), 'tx'));
  if (title) c.title = title;
  const lg = kid(chart, 'legend');
  c.legend = lg ? LEGEND_IN[attr(lg, 'pos') ?? 'r'] ?? 'right' : 'none';
  const dl = find(kid(main[0], 'dataLabels'), 'visibility');
  if (kid(main[0], 'dataLabels') && attr(dl, 'value') !== '0') c.dataLabels = true;
  const axes = kids(kid(chart, 'plotArea'), 'axis');
  const ax0 = axes.find((a) => attr(a, 'id') === '0'), ax1 = axes.find((a) => attr(a, 'id') === '1');
  const x = textOf(kid(kid(ax0, 'title'), 'tx')), y = textOf(kid(kid(ax1, 'title'), 'tx'));
  if (x || y) c.axisTitles = { ...(x ? { x } : {}), ...(y ? { y } : {}) };
  const vs = kid(ax1, 'valScaling');
  const mn = Number(attr(vs, 'min')), mx = Number(attr(vs, 'max'));
  if (attr(vs, 'min') && Number.isFinite(mn)) o.min = mn;
  if (attr(vs, 'max') && Number.isFinite(mx)) o.max = mx;
  if (ax1 && !kid(ax1, 'majorGridlines')) o.gridlines = false;
  if (Object.keys(o).length) c.options = o;
  return c;
};

/** The chart behind one drawing anchor: a classic chart, or an extended chart with its fallback. */
export const readAnchorChart = (a: XNode, load: (rid: string) => XNode | undefined, id: string, sheet: Sheet): { chart: Chart; descr?: string } | undefined => {
  const frames = findAll(a, 'graphicFrame');
  if (!frames.length) return undefined;
  const anchor = readAnchor(a, sheet);
  let classic: Chart | undefined, extended: Chart | undefined, descr: string | undefined, ownExt = false;
  for (const f of frames) {
    const gd = find(f, 'graphicData');
    const ref = find(f, 'chart');
    if (!ref) continue;
    descr ??= attr(find(f, 'cNvPr'), 'descr');
    const root = load(attr(ref, 'id') ?? '');
    if (!root) continue;
    if (attr(gd, 'uri') === NS_CX || root.name === 'chartSpace' && kid(root, 'chartData')) extended ??= readChartEx(root, id, anchor);
    else {
      classic ??= readChart(root, id, anchor);
      ownExt ||= kids(kid(root, 'extLst'), 'ext').some((e) => attr(e, 'uri') === CLERDO_EXT_URI);
    }
  }
  // Our own fallback carries everything; another app's extended part wins over a plain fallback.
  const chart = classic && ownExt && classic.type !== 'column' ? classic : extended ?? classic;
  return chart ? { chart, ...(descr ? { descr } : {}) } : undefined;
};

// ---------------------------------------------------------------- drawing parts

export interface PartWriter {
  put: (path: string, text: string, contentType?: string) => void;
  /** Adds a relationship from the drawing part; returns its id. */
  rel: (type: string, target: string) => string;
  nextChart: () => number;
  look: CellLookup;
}

/** Writes a chart's parts and returns the drawing markup for it (without the anchor). */
export const chartFrame = (c: Chart, shapeId: number, w: PartWriter, withDescr: (xml: string, alt?: string) => string): string => {
  const n = w.nextChart();
  w.put(`xl/charts/chart${n}.xml`, chartXml(c), CT_CHART);
  const rid = w.rel('chart', `../charts/chart${n}.xml`);
  const name = c.title && !c.title.startsWith('=') ? c.title : `Chart ${shapeId - 1}`;
  const classic = withDescr(chartFrameXml(shapeId, name, rid), c.alt);
  if (!isExtended(c.type)) return classic;
  w.put(`xl/charts/chartEx${n}.xml`, chartExXml(c, w.look), CHARTEX_CT);
  const xid = w.rel(CHARTEX_REL, `../charts/chartEx${n}.xml`);
  const [pfx, ns] = REQUIRES[c.type] ?? CX1;
  const frame = withDescr(`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${NS_CX}"><cx:chart xmlns:cx="${NS_CX}" r:id="${xid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`, c.alt);
  return `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:${pfx}="${ns}" Requires="${pfx}">${frame}</mc:Choice><mc:Fallback>${classic}</mc:Fallback></mc:AlternateContent>`;
};

export const chartAnchorXml = (c: Chart, sheet: Sheet, shapeId: number, w: PartWriter, withDescr: (xml: string, alt?: string) => string): string =>
  anchorXml(c.anchor, sheet, chartFrame(c, shapeId, w, withDescr));

// ---------------------------------------------------------------- chart sheets

const CHART_SHEET_SIZE = { w: 910, h: 660 };

/** The chartsheet part and its drawing (one chart filling the sheet). */
export const chartsheetXml = (s: Sheet, w: PartWriter & { drawing: (body: string) => string }, withDescr: (xml: string, alt?: string) => string, tabSelected: boolean): string => {
  const c = s.charts[0];
  let drawing = '';
  if (c) {
    const frame = chartFrame(c, 2, w, withDescr);
    const ext = `<xdr:ext cx="${Math.round(CHART_SHEET_SIZE.w * EMU_PER_PX)}" cy="${Math.round(CHART_SHEET_SIZE.h * EMU_PER_PX)}"/>`;
    drawing = `<drawing r:id="${w.drawing(`<xdr:absoluteAnchor><xdr:pos x="0" y="0"/>${ext}${frame}<xdr:clientData/></xdr:absoluteAnchor>`)}"/>`;
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `${s.tabColor && /^#[0-9a-f]{6}$/i.test(s.tabColor) ? `<sheetPr><tabColor rgb="FF${s.tabColor.slice(1).toUpperCase()}"/></sheetPr>` : ''}`
    + `<sheetViews><sheetView${tabSelected ? ' tabSelected="1"' : ''} zoomToFit="1" workbookViewId="0"/></sheetViews>`
    + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    + `${drawing}</chartsheet>`;
};

/** Reads a chartsheet part: its drawing's first chart becomes the sheet's chart. */
export const readChartsheet = (cs: XNode, sheet: Sheet, rels: Map<string, { target: string; external: boolean }>, xml: (path: string) => XNode | undefined, relsOf: (part: string) => { id: string; target: string; external: boolean }[]): void => {
  sheet.kind = 'chart';
  sheet.view.showGrid = false;
  sheet.view.showHeaders = false;
  const rid = attr(kid(cs, 'drawing'), 'id');
  const rel = rid ? rels.get(rid) : undefined;
  const d = rel && !rel.external ? xml(rel.target) : undefined;
  if (!d || !rel) return;
  const drels = new Map(relsOf(rel.target).map((r) => [r.id, r]));
  for (const a of d.children) {
    if (!/Anchor$/.test(a.name)) continue;
    const got = readAnchorChart(a, (id) => { const r = drels.get(id); return r && !r.external ? xml(r.target) : undefined; }, 'chart1', sheet);
    if (got) {
      got.chart.anchor = { row: 0, col: 0, dx: 0, dy: 0, width: CHART_SHEET_SIZE.w, height: CHART_SHEET_SIZE.h };
      if (got.descr) got.chart.alt = got.descr;
      sheet.charts = [got.chart];
      break;
    }
  }
  const tab = attr(find(kid(cs, 'sheetPr'), 'tabColor'), 'rgb');
  if (tab && /^[0-9A-Fa-f]{8}$/.test(tab)) sheet.tabColor = '#' + tab.slice(2).toUpperCase();
};

