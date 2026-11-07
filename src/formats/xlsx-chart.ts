// DrawingML charts (ECMA-376 Part 1, 21.2) and spreadsheet drawing anchors (20.5).
import type { Chart, ChartOptions, ChartSeries, ChartType, Sheet } from '../model/workbook';
import { attr, esc, find, findAll, kid, kids, numAttr, type XNode } from './xml';

export const EMU_PER_PX = 9525;

/** Column width in characters (as shown) → the width stored in the file, which includes cell padding. */
export const fileColWidth = (chars: number): number => Math.trunc(((chars * 7 + 5) / 7) * 256) / 256;
export const shownColWidth = (file: number): number => Math.round(((file * 7 - 5) / 7) * 100) / 100;

const colPx = (s: Sheet, c: number): number => Math.round(fileColWidth(s.cols.get(c)?.width ?? s.defaultColWidth) * 7);
const rowPx = (s: Sheet, r: number): number => ((s.rows.get(r)?.height ?? s.defaultRowHeight) * 96) / 72;

interface Marker { col: number; colOff: number; row: number; rowOff: number }
const readMarker = (n: XNode | undefined): Marker => ({
  col: +(kid(n, 'col')?.text ?? 0), colOff: +(kid(n, 'colOff')?.text ?? 0),
  row: +(kid(n, 'row')?.text ?? 0), rowOff: +(kid(n, 'rowOff')?.text ?? 0),
});

/** Anchor of a drawing object (twoCellAnchor, oneCellAnchor or absoluteAnchor) in grid terms. */
export const readAnchor = (a: XNode, sheet: Sheet): Chart['anchor'] => {
  const from = readMarker(kid(a, 'from'));
  const out = { row: from.row, col: from.col, dx: Math.round(from.colOff / EMU_PER_PX), dy: Math.round(from.rowOff / EMU_PER_PX), width: 0, height: 0 };
  if (a.name === 'twoCellAnchor') {
    const to = readMarker(kid(a, 'to'));
    let w = 0, h = 0;
    for (let c = from.col; c < to.col && c < from.col + 16384; c++) w += colPx(sheet, c);
    for (let r = from.row; r < to.row && r < from.row + 1_048_576; r++) h += rowPx(sheet, r);
    out.width = Math.round(w + (to.colOff - from.colOff) / EMU_PER_PX);
    out.height = Math.round(h + (to.rowOff - from.rowOff) / EMU_PER_PX);
  } else {
    const ext = kid(a, 'ext');
    out.width = Math.round((numAttr(ext, 'cx') ?? 0) / EMU_PER_PX);
    out.height = Math.round((numAttr(ext, 'cy') ?? 0) / EMU_PER_PX);
    if (a.name === 'absoluteAnchor') {
      // place by walking the grid from the pixel position
      let x = (numAttr(kid(a, 'pos'), 'x') ?? 0) / EMU_PER_PX, y = (numAttr(kid(a, 'pos'), 'y') ?? 0) / EMU_PER_PX;
      let c = 0, r = 0;
      while (c < 16383 && x >= colPx(sheet, c)) x -= colPx(sheet, c++);
      while (r < 1_048_575 && y >= rowPx(sheet, r)) y -= rowPx(sheet, r++);
      Object.assign(out, { col: c, row: r, dx: Math.round(x), dy: Math.round(y) });
    }
  }
  out.width = Math.max(0, out.width);
  out.height = Math.max(0, out.height);
  return out;
};

const markerXml = (tag: string, m: Marker): string =>
  `<xdr:${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>${m.colOff}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${m.rowOff}</xdr:rowOff></xdr:${tag}>`;

export const anchorXml = (a: Chart['anchor'], sheet: Sheet, body: string): string => {
  let x = a.dx + a.width, c = a.col;
  while (c < 16383 && x >= colPx(sheet, c)) x -= colPx(sheet, c++);
  let y = a.dy + a.height, r = a.row;
  while (r < 1_048_575 && y >= rowPx(sheet, r)) y -= rowPx(sheet, r++);
  return '<xdr:twoCellAnchor editAs="oneCell">'
    + markerXml('from', { col: a.col, colOff: Math.round(a.dx * EMU_PER_PX), row: a.row, rowOff: Math.round(a.dy * EMU_PER_PX) })
    + markerXml('to', { col: c, colOff: Math.round(x * EMU_PER_PX), row: r, rowOff: Math.round(y * EMU_PER_PX) })
    + body + '<xdr:clientData/></xdr:twoCellAnchor>';
};

export const chartFrameXml = (id: number, name: string, rid: string): string =>
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
  + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
  + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;

export const pictureXml = (id: number, name: string, rid: string, a: Chart['anchor']): string =>
  `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
  + `<xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
  + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(a.width * EMU_PER_PX)}" cy="${Math.round(a.height * EMU_PER_PX)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;

export const DRAWING_HEAD = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

// ---------------------------------------------------------------- ClerdoGrid extension
//
// Kinds with no DrawingML equivalent (histogram, Pareto, box and whisker, treemap, sunburst,
// waterfall, funnel) are saved as a clustered column chart, the closest shape other apps can show.
// The real type and the chart's options travel in an extension list inside the same chart part
// (ECMA-376 Part 1, 21.2.2.64 and Part 3 extensibility): apps that don't know it skip it.
// On read, the extension only applies while the chart still has the fallback shape it was saved
// with, so a chart that another app has changed keeps that app's version.

export const CLERDO_EXT_URI = '{C6A0D6B2-6F43-4E57-9B1D-2F5E1C3A7D10}';
const CLERDO_NS = 'https://clerdo.app/ns/grid/chart/1';

const EXTENDED: ChartType[] = ['histogram', 'pareto', 'boxWhisker', 'treemap', 'sunburst', 'waterfall', 'funnel'];

const ALL_TYPES: ChartType[] = [
  'column', 'stackedColumn', 'percentColumn', 'bar', 'stackedBar', 'percentBar',
  'line', 'stackedLine', 'percentLine', 'lineMarkers', 'stackedLineMarkers', 'percentLineMarkers',
  'area', 'stackedArea', 'percentArea', 'pie', 'doughnut',
  'scatter', 'scatterSmoothMarkers', 'scatterSmooth', 'scatterLinesMarkers', 'scatterLines', 'bubble',
  'histogram', 'pareto', 'boxWhisker', 'combo', 'comboSecondary', 'comboAreaColumn',
  'treemap', 'sunburst', 'waterfall', 'radar', 'radarMarkers', 'radarFilled', 'funnel',
];
const isType = (t: unknown): t is ChartType => typeof t === 'string' && (ALL_TYPES as string[]).includes(t);

const finiteNum = (v: unknown): number | undefined => (typeof v === 'number' && Number.isFinite(v) ? v : undefined);
const intList = (v: unknown): number[] | undefined =>
  Array.isArray(v) ? v.filter((x): x is number => Number.isInteger(x) && x >= 0 && x < 100000).slice(0, 4096) : undefined;

/** Options from untrusted JSON: only known fields with the right types are kept. */
export const sanitizeOptions = (raw: unknown): ChartOptions | undefined => {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return undefined;
  const r = raw as Record<string, unknown>;
  const o: ChartOptions = {};
  const sec = intList(r.secondary); if (sec?.length) o.secondary = sec;
  const tot = intList(r.totals); if (tot) o.totals = tot;
  const bc = finiteNum(r.binCount); if (bc !== undefined && bc >= 1) o.binCount = Math.min(1000, Math.round(bc));
  const bw = finiteNum(r.binWidth); if (bw !== undefined && bw > 0) o.binWidth = bw;
  if (r.labelMode === 'value' || r.labelMode === 'percent' || r.labelMode === 'both') o.labelMode = r.labelMode;
  if (r.gridlines === false) o.gridlines = false;
  const mn = finiteNum(r.min); if (mn !== undefined) o.min = mn;
  const mx = finiteNum(r.max); if (mx !== undefined) o.max = mx;
  if (typeof r.numFmt === 'string' && r.numFmt.length <= 255) o.numFmt = r.numFmt;
  const st = finiteNum(r.style); if (st !== undefined && Number.isInteger(st) && st > 0 && st < 32) o.style = st;
  if (r.hideMean === true) o.hideMean = true;
  if (r.hideOutliers === true) o.hideOutliers = true;
  return Object.keys(o).length ? o : undefined;
};

// ---------------------------------------------------------------- chart reading

const richText = (n: XNode | undefined): string | undefined => {
  if (!n) return undefined;
  const tx = kid(n, 'tx');
  const f = find(tx, 'f');
  if (f) return f.text;
  const paras = findAll(tx, 'p');
  if (paras.length) return paras.map((p) => findAll(p, 't').map((t) => t.text).join('')).join('\n');
  const t = findAll(n, 't').map((x) => x.text).join('');
  return t || undefined;
};

const seriesRef = (n: XNode | undefined): string | undefined => {
  const f = find(n, 'f');
  return f ? f.text.trim() : undefined;
};

const seriesColor = (ser: XNode): string | undefined => {
  const sp = kid(ser, 'spPr');
  const fill = kid(sp, 'solidFill') ?? kid(kid(sp, 'ln'), 'solidFill') ?? kid(kid(kid(ser, 'marker'), 'spPr'), 'solidFill');
  const c = attr(kid(fill, 'srgbClr'), 'val');
  return c && /^[0-9A-Fa-f]{6}$/.test(c) ? '#' + c.toUpperCase() : undefined;
};

const readSeries = (ser: XNode, sub?: ChartSeries['type']): ChartSeries => {
  const s: ChartSeries = { values: seriesRef(kid(ser, 'val') ?? kid(ser, 'yVal')) ?? '' };
  const tx = kid(ser, 'tx');
  if (tx) { const f = find(tx, 'f'); const v = find(tx, 'v'); const name = f?.text ?? v?.text; if (name) s.name = name; }
  const cat = seriesRef(kid(ser, 'cat') ?? kid(ser, 'xVal'));
  if (cat) s.categories = cat;
  const color = seriesColor(ser);
  if (color) s.color = color;
  if (sub) s.type = sub;
  const sizes = seriesRef(kid(ser, 'bubbleSize'));
  if (sizes) s.sizes = sizes;
  return s;
};

const LEGEND: Record<string, Chart['legend']> = { r: 'right', b: 'bottom', t: 'top', l: 'left', tr: 'right' };

/** Whether any series draws point markers. */
const hasMarkers = (g: XNode): boolean => kids(g, 'ser').some((s) => {
  const m = kid(s, 'marker');
  return m ? (attr(kid(m, 'symbol'), 'val') ?? 'auto') !== 'none' : attr(kid(g, 'marker'), 'val') !== '0';
});
const serFlag = (g: XNode, tag: string): boolean => kids(g, 'ser').some((s) => attr(kid(s, tag), 'val') === '1' || attr(kid(s, tag), 'val') === 'true');
const lineHidden = (g: XNode): boolean => kids(g, 'ser').length > 0 && kids(g, 'ser').every((s) => !!kid(kid(kid(s, 'spPr'), 'ln'), 'noFill'));

const groupType = (g: XNode): ChartType | undefined => {
  const grouping = attr(kid(g, 'grouping'), 'val') ?? 'standard';
  const stack = grouping === 'stacked' ? 'stacked' : grouping === 'percentStacked' ? 'percent' : '';
  switch (g.name) {
    case 'barChart': case 'bar3DChart': {
      const bar = attr(kid(g, 'barDir'), 'val') === 'bar';
      if (stack === 'stacked') return bar ? 'stackedBar' : 'stackedColumn';
      if (stack === 'percent') return bar ? 'percentBar' : 'percentColumn';
      return bar ? 'bar' : 'column';
    }
    case 'lineChart': case 'line3DChart': case 'stockChart': {
      const m = g.name === 'lineChart' && hasMarkers(g);
      if (stack === 'stacked') return m ? 'stackedLineMarkers' : 'stackedLine';
      if (stack === 'percent') return m ? 'percentLineMarkers' : 'percentLine';
      return m ? 'lineMarkers' : 'line';
    }
    case 'areaChart': case 'area3DChart': return stack === 'stacked' ? 'stackedArea' : stack === 'percent' ? 'percentArea' : 'area';
    case 'pieChart': case 'pie3DChart': case 'ofPieChart': return 'pie';
    case 'doughnutChart': return 'doughnut';
    case 'scatterChart': {
      const style = attr(kid(g, 'scatterStyle'), 'val') ?? 'marker';
      const lines = !lineHidden(g) && style !== 'marker' && style !== 'none';
      const markers = hasMarkers(g);
      const smooth = serFlag(g, 'smooth') || ((style === 'smooth' || style === 'smoothMarker') && !kids(g, 'ser').some((s) => kid(s, 'smooth')));
      if (!lines) return 'scatter';
      if (smooth) return markers ? 'scatterSmoothMarkers' : 'scatterSmooth';
      return markers ? 'scatterLinesMarkers' : 'scatterLines';
    }
    case 'bubbleChart': return 'bubble';
    case 'radarChart': {
      const style = attr(kid(g, 'radarStyle'), 'val') ?? 'standard';
      if (style === 'filled') return 'radarFilled';
      return hasMarkers(g) ? 'radarMarkers' : 'radar';
    }
    default: return undefined;
  }
};

const subKind = (t: ChartType): NonNullable<ChartSeries['type']> =>
  /line/i.test(t) ? 'line' : /area/i.test(t) ? 'area' : 'column';

const readExt = (root: XNode): { type: ChartType; base: ChartType; options?: ChartOptions } | undefined => {
  const ext = kids(kid(root, 'extLst'), 'ext').find((e) => attr(e, 'uri') === CLERDO_EXT_URI);
  const el = kid(ext, 'chart');
  const type = attr(el, 'type'), base = attr(el, 'base');
  if (!isType(type) || !isType(base)) return undefined;
  let options: ChartOptions | undefined;
  const json = attr(el, 'options');
  if (json && json.length < 200000) {
    try { options = sanitizeOptions(JSON.parse(json)); } catch { options = undefined; }
  }
  return { type, base, ...(options ? { options } : {}) };
};

export const readChart = (root: XNode, id: string, anchor: Chart['anchor']): Chart | undefined => {
  const chart = kid(root, 'chart');
  const plot = kid(chart, 'plotArea');
  if (!plot) return undefined;
  const groups = plot.children.filter((c) => /Chart$/.test(c.name));
  const known = groups.map((g) => ({ g, t: groupType(g) })).filter((x): x is { g: XNode; t: ChartType } => !!x.t && kids(x.g, 'ser').length > 0);
  if (!known.length) return undefined;

  // Axis pairs: the first group's axes are primary; any other value axis is secondary.
  const axIdsOf = (g: XNode) => kids(g, 'axId').map((a) => attr(a, 'val') ?? '');
  const primaryAx = new Set(axIdsOf(known[0]!.g));
  const onSecondary = (g: XNode) => { const ids = axIdsOf(g); return ids.length > 0 && !ids.some((x) => primaryAx.has(x)); };

  // Groups that differ only in axis (one per axis) are one type, not a combo.
  const combo = new Set(known.map((k) => k.t)).size > 1;
  const series: ChartSeries[] = [];
  const secondary: number[] = [];
  for (const { g, t } of known) {
    const sec = onSecondary(g);
    for (const ser of kids(g, 'ser')) {
      if (sec) secondary.push(series.length);
      series.push(readSeries(ser, combo ? subKind(t) : undefined));
    }
  }
  let type: ChartType = known[0]!.t;
  const options: ChartOptions = {};
  if (combo) {
    const kinds = known.map((k) => k.t);
    const barCol = kinds.some((t) => t === 'column');
    const area = kinds.some((t) => t === 'stackedArea');
    const lineSecondary = known.filter((k) => subKind(k.t) === 'line').every((k) => onSecondary(k.g)) && known.filter((k) => subKind(k.t) !== 'line').every((k) => !onSecondary(k.g));
    if (barCol && area && known.length === 2) type = 'comboAreaColumn';
    else if (barCol && kinds.some((t) => subKind(t) === 'line') && lineSecondary && secondary.length) type = 'comboSecondary';
    else type = 'combo';
    if (type === 'comboAreaColumn') series.forEach((s) => { if (s.type === 'line') s.type = 'column'; });
  }
  const defaultSecondary = type === 'comboSecondary' ? series.flatMap((s, i) => (s.type === 'line' ? [i] : [])) : [];
  if (secondary.length && secondary.join() !== defaultSecondary.join()) options.secondary = secondary;
  if (type === 'comboSecondary' && !secondary.length) type = 'combo';

  const c: Chart = { id, type, series, anchor };
  const title = richText(kid(chart, 'title'));
  if (title) c.title = title;
  const legend = kid(chart, 'legend');
  c.legend = legend ? LEGEND[attr(kid(legend, 'legendPos'), 'val') ?? 'r'] ?? 'right' : 'none';

  const axes = plot.children.filter((n) => /Ax$/.test(n.name));
  const axById = (idv: string) => axes.find((a) => attr(kid(a, 'axId'), 'val') === idv);
  const pIds = axIdsOf(known[0]!.g);
  const pAxes = pIds.map(axById).filter((a): a is XNode => !!a);
  const scatterLike = known[0]!.g.name === 'scatterChart' || known[0]!.g.name === 'bubbleChart';
  const catAx = scatterLike
    ? pAxes.find((a) => /^[bt]$/.test(attr(kid(a, 'axPos'), 'val') ?? '')) ?? pAxes[0]
    : pAxes.find((a) => a.name === 'catAx' || a.name === 'dateAx' || a.name === 'serAx');
  const valAx = pAxes.find((a) => a.name === 'valAx' && a !== catAx);
  const secAx = known.filter((k) => onSecondary(k.g)).flatMap((k) => axIdsOf(k.g)).map(axById).find((a) => a?.name === 'valAx');
  const x = richText(kid(catAx, 'title')), y = richText(kid(valAx, 'title')), y2 = richText(kid(secAx, 'title'));
  if (x || y || y2) c.axisTitles = { ...(x ? { x } : {}), ...(y ? { y } : {}), ...(y2 ? { y2 } : {}) };
  if (valAx) {
    const sc = kid(valAx, 'scaling');
    const mn = numAttr(kid(sc, 'min'), 'val'), mx = numAttr(kid(sc, 'max'), 'val');
    if (mn !== undefined && Number.isFinite(mn)) options.min = mn;
    if (mx !== undefined && Number.isFinite(mx)) options.max = mx;
    if (!kid(valAx, 'majorGridlines')) options.gridlines = false;
    const nf = kid(valAx, 'numFmt');
    const code = attr(nf, 'formatCode');
    if (code && attr(nf, 'sourceLinked') === '0' && code !== 'General') options.numFmt = code;
  }

  const flag = (n: XNode | undefined, tag: string) => attr(kid(n, tag), 'val') === '1' || attr(kid(n, tag), 'val') === 'true';
  const lbls = known.flatMap(({ g }) => [kid(g, 'dLbls'), ...kids(g, 'ser').map((s) => kid(s, 'dLbls'))]).filter((n): n is XNode => !!n);
  const showVal = lbls.some((n) => flag(n, 'showVal')), showPct = lbls.some((n) => flag(n, 'showPercent'));
  if (showVal || showPct) c.dataLabels = true;
  if (type === 'pie' || type === 'doughnut') {
    if (showVal && showPct) options.labelMode = 'both';
    else if (showVal) options.labelMode = 'value';
  }

  const ext = readExt(root);
  if (ext && ext.base === type) {
    c.type = ext.type;
    if (ext.options) c.options = ext.options; else delete c.options;
    if (!['combo', 'comboSecondary', 'comboAreaColumn'].includes(c.type)) c.series = c.series.map(({ type: _t, ...s }) => s);
  } else if (Object.keys(options).length) c.options = options;
  return c;
};

// ---------------------------------------------------------------- chart writing

const REF_RE = /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

const txRich = (text: string): string =>
  `<c:tx><c:rich><a:bodyPr/><a:lstStyle/>${text.split('\n').map((l) => `<a:p><a:r><a:t>${esc(l)}</a:t></a:r></a:p>`).join('')}</c:rich></c:tx>`;
const titleXml = (text: string | undefined): string => (text ? `<c:title>${txRich(text)}<c:overlay val="0"/></c:title>` : '');
const v = (tag: string, val: string | number): string => `<c:${tag} val="${val}"/>`;

const dLblsXml = (show: boolean, pct = false): string =>
  `<c:dLbls>${v('showLegendKey', 0)}${v('showVal', show ? 1 : 0)}${v('showCatName', 0)}${v('showSerName', 0)}${v('showPercent', pct ? 1 : 0)}${v('showBubbleSize', 0)}</c:dLbls>`;

const refXml = (ref: string | undefined, num: boolean): string => {
  if (!ref) return '';
  return num ? `<c:numRef><c:f>${esc(ref)}</c:f></c:numRef>` : `<c:strRef><c:f>${esc(ref)}</c:f></c:strRef>`;
};

type SerKind = 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar' | 'bubble';
interface SerStyle { markers?: boolean; smooth?: boolean; lines?: boolean; filled?: boolean }

const solid = (color: string | undefined): string => (color ? `<a:solidFill><a:srgbClr val="${color.replace('#', '').slice(0, 6).toUpperCase()}"/></a:solidFill>` : '');

const serXml = (s: ChartSeries, i: number, kind: SerKind, st: SerStyle = {}): string => {
  let x = v('idx', i) + v('order', i);
  if (s.name) x += REF_RE.test(s.name) ? `<c:tx>${refXml(s.name, false)}</c:tx>` : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
  // Without a colour of its own, a series takes the theme's automatic colour.
  const fill = solid(s.color);
  const marker = (on: boolean) => (on ? `<c:marker>${v('symbol', 'circle')}${v('size', 5)}${fill ? `<c:spPr>${fill}<a:ln>${fill}</a:ln></c:spPr>` : ''}</c:marker>` : `<c:marker>${v('symbol', 'none')}</c:marker>`);
  if (kind === 'line' || (kind === 'radar' && !st.filled)) x += `${fill ? `<c:spPr><a:ln w="28575" cap="rnd">${fill}<a:round/></a:ln></c:spPr>` : ''}${marker(!!st.markers)}`;
  else if (kind === 'radar') x += `${fill ? `<c:spPr>${fill}</c:spPr>` : ''}${marker(false)}`;
  else if (kind === 'scatter') x += `<c:spPr><a:ln w="${st.lines ? 28575 : 19050}" cap="rnd">${st.lines ? fill : '<a:noFill/>'}${st.lines ? '<a:round/>' : ''}</a:ln></c:spPr>${marker(!!st.markers)}`;
  else if (kind !== 'pie' && fill) x += `<c:spPr>${fill}</c:spPr>`;
  if (kind === 'bar' || kind === 'bubble') x += v('invertIfNegative', 0);
  if (kind === 'scatter' || kind === 'bubble') {
    x += `${s.categories ? `<c:xVal>${refXml(s.categories, true)}</c:xVal>` : ''}<c:yVal>${refXml(s.values, true)}</c:yVal>`;
    if (kind === 'bubble') x += `<c:bubbleSize>${s.sizes ? refXml(s.sizes, true) : refXml(s.values, true)}</c:bubbleSize>${v('bubble3D', 0)}`;
    else x += v('smooth', st.smooth ? 1 : 0);
  } else {
    x += s.categories ? `<c:cat>${refXml(s.categories, false)}</c:cat>` : '';
    x += `<c:val>${refXml(s.values, true)}</c:val>`;
    if (kind === 'line') x += v('smooth', 0);
  }
  return `<c:ser>${x}</c:ser>`;
};

const AX_CAT = 500000001, AX_VAL = 500000002, AX_CAT2 = 500000003, AX_VAL2 = 500000004;

interface AxisPlan { horizontal: boolean; scatter: boolean; secondary: boolean }

const axesXml = (c: Chart, p: AxisPlan): string => {
  const catPos = p.horizontal ? 'l' : 'b', valPos = p.horizontal ? 'b' : 'l';
  const o = c.options ?? {};
  const scaling = (withBounds: boolean) => `<c:scaling>${v('orientation', 'minMax')}${withBounds && o.max !== undefined ? v('max', o.max) : ''}${withBounds && o.min !== undefined ? v('min', o.min) : ''}</c:scaling>`;
  const common = (id: number, cross: number, pos: string, title: string | undefined, opts: { grid: boolean; bounds: boolean; del?: boolean; fmt?: string; crosses?: string }): string =>
    v('axId', id) + scaling(opts.bounds) + v('delete', opts.del ? 1 : 0) + v('axPos', pos)
    + (opts.grid ? '<c:majorGridlines/>' : '') + titleXml(title)
    + (opts.fmt ? `<c:numFmt formatCode="${esc(opts.fmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>')
    + v('majorTickMark', 'out') + v('minorTickMark', 'none') + v('tickLblPos', 'nextTo') + v('crossAx', cross) + v('crosses', opts.crosses ?? 'autoZero');
  const grid = o.gridlines !== false;
  const catTail = `${v('auto', 1)}${v('lblAlgn', 'ctr')}${v('lblOffset', 100)}${v('noMultiLvlLbl', 0)}`;
  const cat = p.scatter
    ? `<c:valAx>${common(AX_CAT, AX_VAL, catPos, c.axisTitles?.x, { grid: false, bounds: false })}${v('crossBetween', 'midCat')}</c:valAx>`
    : `<c:catAx>${common(AX_CAT, AX_VAL, catPos, c.axisTitles?.x, { grid: false, bounds: false })}${catTail}</c:catAx>`;
  const val = `<c:valAx>${common(AX_VAL, AX_CAT, valPos, c.axisTitles?.y, { grid, bounds: true, fmt: o.numFmt })}${v('crossBetween', p.scatter ? 'midCat' : 'between')}</c:valAx>`;
  if (!p.secondary) return cat + val;
  const cat2 = `<c:catAx>${common(AX_CAT2, AX_VAL2, 'b', undefined, { grid: false, bounds: false, del: true })}${catTail}</c:catAx>`;
  const val2 = `<c:valAx>${common(AX_VAL2, AX_CAT2, 'r', c.axisTitles?.y2, { grid: false, bounds: false, fmt: o.numFmt, crosses: 'max' })}${v('crossBetween', 'between')}</c:valAx>`;
  return cat + val + cat2 + val2;
};

/** The DrawingML type a chart is saved as. */
export const savedAs = (t: ChartType): ChartType => (EXTENDED.includes(t) ? 'column' : t);

const comboKindOf = (c: Chart, s: ChartSeries, i: number): NonNullable<ChartSeries['type']> => {
  if (s.type) return s.type;
  if (c.type === 'comboAreaColumn') return i < Math.max(1, Math.ceil(c.series.length / 2)) && !(c.series.length > 1 && i === c.series.length - 1) ? 'area' : 'column';
  return i === 0 ? 'column' : 'line';
};

export const chartXml = (c: Chart): string => {
  const o = c.options ?? {};
  const type = savedAs(c.type);
  const dl = dLblsXml(!!c.dataLabels);
  const pieDl = () => {
    const mode = o.labelMode ?? 'percent';
    return c.dataLabels ? dLblsXml(mode !== 'percent', mode !== 'value') : dLblsXml(false);
  };
  const ax = (sec: boolean) => (sec ? v('axId', AX_CAT2) + v('axId', AX_VAL2) : v('axId', AX_CAT) + v('axId', AX_VAL));
  const sec = new Set(o.secondary ?? (c.type === 'comboSecondary' ? c.series.flatMap((s, i) => (comboKindOf(c, s, i) === 'line' ? [i] : [])) : []));
  const indexed = c.series.map((s, i) => ({ s, i }));
  const grp = (g: string) => (g === 'stacked' ? 'stacked' : g === 'percent' ? 'percentStacked' : g);
  const bar = (dir: 'col' | 'bar', grouping: string, list: { s: ChartSeries; i: number }[], secondary = false): string =>
    `<c:barChart>${v('barDir', dir)}${v('grouping', grp(grouping))}${v('varyColors', 0)}${list.map(({ s, i }) => serXml(s, i, 'bar')).join('')}${dl}`
    + `${v('gapWidth', 150)}${grouping !== 'clustered' ? v('overlap', 100) : ''}${ax(secondary)}</c:barChart>`;
  const line = (grouping: string, list: { s: ChartSeries; i: number }[], markers: boolean, secondary = false): string =>
    `<c:lineChart>${v('grouping', grp(grouping))}${v('varyColors', 0)}${list.map(({ s, i }) => serXml(s, i, 'line', { markers })).join('')}${dl}${v('marker', 1)}${ax(secondary)}</c:lineChart>`;
  const area = (grouping: string, list: { s: ChartSeries; i: number }[], secondary = false): string =>
    `<c:areaChart>${v('grouping', grp(grouping))}${v('varyColors', 0)}${list.map(({ s, i }) => serXml(s, i, 'area')).join('')}${dl}${ax(secondary)}</c:areaChart>`;
  /** One group per axis for types that may put some series on the right axis. */
  const split = (make: (list: { s: ChartSeries; i: number }[], secondary: boolean) => string): string => {
    const a = indexed.filter((x) => !sec.has(x.i)), b = indexed.filter((x) => sec.has(x.i));
    return (a.length ? make(a, false) : '') + (b.length ? make(b, true) : '');
  };

  let body = '';
  const plan: AxisPlan = { horizontal: false, scatter: false, secondary: false };
  let axes = true;
  const stackOf = (t: ChartType) => (/^stacked/.test(t) ? 'stacked' : /^percent/.test(t) ? 'percent' : '');
  switch (type) {
    case 'column': body = split((l, s) => bar('col', 'clustered', l, s)); break;
    case 'stackedColumn': case 'percentColumn': body = bar('col', stackOf(type), indexed); break;
    case 'bar': case 'stackedBar': case 'percentBar': body = bar('bar', stackOf(type) || 'clustered', indexed); plan.horizontal = true; break;
    case 'line': case 'lineMarkers': body = split((l, s) => line('standard', l, type === 'lineMarkers', s)); break;
    case 'stackedLine': case 'percentLine': case 'stackedLineMarkers': case 'percentLineMarkers':
      body = line(stackOf(type), indexed, /Markers$/.test(type)); break;
    case 'area': body = split((l, s) => area('standard', l, s)); break;
    case 'stackedArea': case 'percentArea': body = area(stackOf(type), indexed); break;
    case 'pie': case 'doughnut': {
      axes = false;
      const list = type === 'pie' ? c.series.slice(0, 1) : c.series;
      const sers = list.map((s, i) => serXml(s, i, 'pie')).join('');
      body = type === 'pie'
        ? `<c:pieChart>${v('varyColors', 1)}${sers}${pieDl()}${v('firstSliceAng', 0)}</c:pieChart>`
        : `<c:doughnutChart>${v('varyColors', 1)}${sers}${pieDl()}${v('firstSliceAng', 0)}${v('holeSize', 50)}</c:doughnutChart>`;
      break;
    }
    case 'scatter': case 'scatterSmoothMarkers': case 'scatterSmooth': case 'scatterLinesMarkers': case 'scatterLines': {
      plan.scatter = true;
      const st: SerStyle = {
        markers: type === 'scatter' || type === 'scatterSmoothMarkers' || type === 'scatterLinesMarkers',
        smooth: type === 'scatterSmooth' || type === 'scatterSmoothMarkers',
        lines: type !== 'scatter',
      };
      const style = !st.lines ? 'lineMarker' : st.smooth ? (st.markers ? 'smoothMarker' : 'smooth') : st.markers ? 'lineMarker' : 'line';
      body = `<c:scatterChart>${v('scatterStyle', style)}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'scatter', st)).join('')}${dl}${ax(false)}</c:scatterChart>`;
      break;
    }
    case 'bubble':
      plan.scatter = true;
      body = `<c:bubbleChart>${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'bubble')).join('')}${dl}${v('bubble3D', 0)}${v('bubbleScale', 100)}${v('showNegBubbles', 0)}${ax(false)}</c:bubbleChart>`;
      break;
    case 'radar': case 'radarMarkers': case 'radarFilled': {
      const style = type === 'radarFilled' ? 'filled' : type === 'radarMarkers' ? 'marker' : 'standard';
      body = `<c:radarChart>${v('radarStyle', style)}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'radar', { markers: type === 'radarMarkers', filled: type === 'radarFilled' })).join('')}${dl}${ax(false)}</c:radarChart>`;
      break;
    }
    case 'combo': case 'comboSecondary': case 'comboAreaColumn': {
      // Areas first so columns and lines draw over them.
      for (const secondary of [false, true]) {
        const of = (k: string) => indexed.filter(({ s, i }) => comboKindOf(c, s, i) === k && sec.has(i) === secondary);
        const a = of('area'), b = of('column'), l = of('line');
        if (a.length) body += area(c.type === 'comboAreaColumn' ? 'stacked' : 'standard', a, secondary);
        if (b.length) body += bar('col', 'clustered', b, secondary);
        if (l.length) body += line('standard', l, true, secondary);
      }
      break;
    }
    default: body = bar('col', 'clustered', indexed);
  }
  plan.secondary = sec.size > 0 && body.includes(`val="${AX_VAL2}"`);
  const legend = c.legend === 'none' ? '' : `<c:legend>${v('legendPos', { right: 'r', bottom: 'b', top: 't', left: 'l' }[c.legend ?? 'right'])}<c:overlay val="0"/></c:legend>`;
  // The extension is written when the file alone can't bring the chart back as it is.
  const needsExt = type !== c.type || o.style !== undefined || o.labelMode !== undefined && type !== 'pie' && type !== 'doughnut'
    || o.totals !== undefined || o.binCount !== undefined || o.binWidth !== undefined || o.hideMean || o.hideOutliers;
  const ext = needsExt
    ? `<c:extLst><c:ext uri="${CLERDO_EXT_URI}"><cg:chart xmlns:cg="${CLERDO_NS}" type="${c.type}" base="${type}"${c.options ? ` options="${esc(JSON.stringify(c.options))}"` : ''}/></c:ext></c:extLst>`
    : '';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + v('roundedCorners', 0) + '<c:chart>' + (c.title ? titleXml(c.title) + v('autoTitleDeleted', 0) : v('autoTitleDeleted', 1))
    + `<c:plotArea><c:layout/>${body}${axes ? axesXml(c, plan) : ''}</c:plotArea>`
    + legend + v('plotVisOnly', 1) + v('dispBlanksAs', 'gap') + '</c:chart>' + ext + '</c:chartSpace>';
};
