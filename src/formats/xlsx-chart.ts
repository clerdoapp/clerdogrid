// DrawingML charts (ECMA-376 Part 1, 21.2) and spreadsheet drawing anchors (20.5).
import type {
  AxisOptions, Chart, ChartOptions, ChartSeries, ChartType, DataLabelOptions, ErrorBars, FillFormat, LineFormat, MarkerFormat, MarkerSymbol, PointFormat, Sheet, ShapeFormat, Trendline,
} from '../model/workbook';
import { attr, esc, find, findAll, kid, kids, numAttr, type XNode } from './xml';
import { sanitizeOptions } from './chart-options';

export { sanitizeOptions };
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
// waterfall, funnel, filled map) are saved as a clustered column chart, the closest shape other apps
// can show; they also get an extended chart part (xlsx-chartex.ts) that newer apps draw natively.
// The real type and the chart's options travel in an extension list inside the same chart part
// (ECMA-376 Part 1, 21.2.2.64 and Part 3 extensibility): apps that don't know it skip it.
// On read, the extension only applies while the chart still has the fallback shape it was saved
// with, so a chart that another app has changed keeps that app's version.

export const CLERDO_EXT_URI = '{C6A0D6B2-6F43-4E57-9B1D-2F5E1C3A7D10}';
const CLERDO_NS = 'https://clerdo.app/ns/grid/chart/1';
const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
const EXT_LABEL_RANGE = '{02D57815-91ED-43cb-92C2-25804820EDAC}';
const EXT_SHOW_LABEL_RANGE = '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}';

export const EXTENDED: ChartType[] = ['histogram', 'pareto', 'boxWhisker', 'treemap', 'sunburst', 'waterfall', 'funnel', 'regionMap'];

const ALL_TYPES: ChartType[] = [
  'column', 'stackedColumn', 'percentColumn', 'bar', 'stackedBar', 'percentBar',
  'line', 'stackedLine', 'percentLine', 'lineMarkers', 'stackedLineMarkers', 'percentLineMarkers',
  'area', 'stackedArea', 'percentArea', 'pie', 'doughnut',
  'scatter', 'scatterSmoothMarkers', 'scatterSmooth', 'scatterLinesMarkers', 'scatterLines', 'bubble',
  'histogram', 'pareto', 'boxWhisker', 'combo', 'comboSecondary', 'comboAreaColumn',
  'treemap', 'sunburst', 'waterfall', 'radar', 'radarMarkers', 'radarFilled', 'funnel',
  'stockHLC', 'stockOHLC', 'stockVHLC', 'stockVOHLC', 'surface3D', 'surfaceWireframe3D', 'contour', 'contourWireframe',
  'column3DClustered', 'column3DStacked', 'column3DPercent', 'column3D', 'bar3DClustered', 'bar3DStacked', 'bar3DPercent',
  'cylinderColumn', 'coneColumn', 'pyramidColumn', 'line3D', 'area3D', 'stackedArea3D', 'percentArea3D',
  'pie3D', 'pieExploded', 'pie3DExploded', 'doughnutExploded', 'barOfPie', 'pieOfPie', 'regionMap', 'bubble3D',
];
export const isType = (t: unknown): t is ChartType => typeof t === 'string' && (ALL_TYPES as string[]).includes(t);

// ---------------------------------------------------------------- shape properties

const hex = (c: string | undefined): string | undefined => (c && /^#?[0-9a-fA-F]{6}$/.test(c) ? c.replace('#', '').toUpperCase() : undefined);
const clr = (c: string, alpha?: number): string => `<a:srgbClr val="${hex(c)}">${alpha !== undefined && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''}</a:srgbClr>`;
const solid = (color: string | undefined, alpha?: number): string => (hex(color) ? `<a:solidFill>${clr(color!, alpha)}</a:solidFill>` : '');

const DASH_OUT: Record<string, string> = { solid: 'solid', dot: 'sysDot', dash: 'dash', dashDot: 'dashDot', longDash: 'lgDash', longDashDot: 'lgDashDot' };
const DASH_IN: Record<string, LineFormat['dash']> = { solid: 'solid', dot: 'dot', sysDot: 'dot', sysDash: 'dash', dash: 'dash', dashDot: 'dashDot', sysDashDot: 'dashDot', lgDash: 'longDash', lgDashDot: 'longDashDot', lgDashDotDot: 'longDashDot', sysDashDotDot: 'dashDot' };

const fillXml = (f: FillFormat | undefined, fallback?: string): string => {
  if (f?.none) return '<a:noFill/>';
  const alpha = f?.transparency !== undefined ? 1 - f.transparency / 100 : undefined;
  if (f?.gradient && hex(f.gradient.from) && hex(f.gradient.to)) {
    const ang = Math.round((((f.gradient.angle ?? 90) % 360) + 360) % 360 * 60000);
    return `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${clr(f.gradient.from, alpha)}</a:gs><a:gs pos="100000">${clr(f.gradient.to, alpha)}</a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
  }
  return solid(f?.color ?? fallback, alpha);
};

const lnXml = (l: LineFormat | undefined, fallback?: string, defaultW?: number, round = false): string => {
  if (!l && !fallback && !defaultW) return '';
  if (l?.none) return '<a:ln><a:noFill/></a:ln>';
  const w = l?.width !== undefined ? Math.round(l.width * 12700) : defaultW;
  const color = solid(l?.color ?? fallback);
  const dash = l?.dash && l.dash !== 'solid' ? `<a:prstDash val="${DASH_OUT[l.dash]}"/>` : '';
  if (!color && !dash && w === undefined) return '';
  return `<a:ln${w !== undefined ? ` w="${w}"` : ''}${round ? ' cap="rnd"' : ''}>${color}${dash}${round ? '<a:round/>' : ''}</a:ln>`;
};

const spPr = (fill: string, ln: string): string => (fill || ln ? `<c:spPr>${fill}${ln}</c:spPr>` : '');
const shapeXml = (s: ShapeFormat | undefined): string => (s ? spPr(fillXml(s.fill), lnXml(s.line)) : '');

const readColor = (n: XNode | undefined): { color?: string; alpha?: number } => {
  const c = kid(n, 'srgbClr');
  const v = attr(c, 'val');
  if (!v || !/^[0-9A-Fa-f]{6}$/.test(v)) return {};
  const a = numAttr(kid(c, 'alpha'), 'val');
  return { color: '#' + v.toUpperCase(), ...(a !== undefined && a < 100000 ? { alpha: a / 100000 } : {}) };
};

const readFill = (sp: XNode | undefined): FillFormat | undefined => {
  if (!sp) return undefined;
  if (kid(sp, 'noFill')) return { none: true };
  const g = kid(sp, 'gradFill');
  if (g) {
    const stops = kids(kid(g, 'gsLst'), 'gs').map((s) => ({ pos: numAttr(s, 'pos') ?? 0, ...readColor(s) })).filter((s) => s.color).sort((a, b) => a.pos - b.pos);
    if (stops.length >= 2) {
      const ang = numAttr(kid(g, 'lin'), 'ang');
      const f: FillFormat = { gradient: { from: stops[0]!.color!, to: stops.at(-1)!.color!, ...(ang !== undefined && ang !== 5400000 ? { angle: Math.round(ang / 60000) } : {}) } };
      if (stops[0]!.alpha !== undefined) f.transparency = Math.round((1 - stops[0]!.alpha) * 100);
      return f;
    }
  }
  const s = readColor(kid(sp, 'solidFill'));
  if (s.color) return { color: s.color, ...(s.alpha !== undefined ? { transparency: Math.round((1 - s.alpha) * 100) } : {}) };
  return undefined;
};

const readLine = (ln: XNode | undefined): LineFormat | undefined => {
  if (!ln) return undefined;
  if (kid(ln, 'noFill')) return { none: true };
  const out: LineFormat = {};
  const c = readColor(kid(ln, 'solidFill'));
  if (c.color) out.color = c.color;
  const w = numAttr(ln, 'w');
  if (w !== undefined) out.width = Math.round((w / 12700) * 100) / 100;
  const d = attr(kid(ln, 'prstDash'), 'val');
  if (d && d !== 'solid' && DASH_IN[d]) out.dash = DASH_IN[d];
  return Object.keys(out).length ? out : undefined;
};

const readShape = (n: XNode | undefined): ShapeFormat | undefined => {
  const sp = kid(n, 'spPr');
  if (!sp) return undefined;
  const fill = readFill(sp), line = readLine(kid(sp, 'ln'));
  return fill || line ? { ...(fill ? { fill } : {}), ...(line ? { line } : {}) } : undefined;
};

// ---------------------------------------------------------------- chart reading

const val = (n: XNode | undefined, tag: string): string | undefined => attr(kid(n, tag), 'val');
const num = (n: XNode | undefined, tag: string): number | undefined => { const v = Number(val(n, tag)); return val(n, tag) !== undefined && Number.isFinite(v) ? v : undefined; };
const flag = (n: XNode | undefined, tag: string, dflt = false): boolean => { const v = val(n, tag); return v === undefined ? (kid(n, tag) ? true : dflt) : v === '1' || v === 'true'; };

const richText = (n: XNode | undefined): string | undefined => {
  if (!n) return undefined;
  const tx = kid(n, 'tx');
  const f = find(tx, 'f');
  if (f) return '=' + f.text.trim();
  const paras = findAll(tx, 'p');
  if (paras.length) return paras.map((p) => findAll(p, 't').map((t) => t.text).join('')).join('\n');
  const t = findAll(n, 't').map((x) => x.text).join('');
  return t || undefined;
};

const seriesRef = (n: XNode | undefined): string | undefined => {
  const f = find(n, 'f');
  if (f) return f.text.trim();
  // Literal values (numLit/strLit) become an array constant.
  const lit = kid(n, 'numLit') ?? kid(n, 'strLit');
  if (!lit) return undefined;
  const pts = kids(lit, 'pt').map((p) => kid(p, 'v')?.text ?? '');
  if (!pts.length || pts.length > 4096) return undefined;
  return `={${pts.map((v) => (lit.name === 'numLit' && Number.isFinite(Number(v)) ? v : `"${v.replace(/"/g, '""')}"`)).join(',')}}`;
};

const seriesColor = (ser: XNode, filled: boolean): string | undefined => {
  const sp = kid(ser, 'spPr');
  // Filled shapes take their fill colour; an outline colour of its own is a line format.
  const fill = kid(sp, 'solidFill') ?? (filled && (kid(sp, 'gradFill') || kid(sp, 'noFill')) ? undefined : kid(kid(sp, 'ln'), 'solidFill') ?? (kid(sp, 'ln') ? undefined : kid(kid(kid(ser, 'marker'), 'spPr'), 'solidFill')));
  const c = attr(kid(fill, 'srgbClr'), 'val');
  return c && /^[0-9A-Fa-f]{6}$/.test(c) ? '#' + c.toUpperCase() : undefined;
};

const SYMBOLS: MarkerSymbol[] = ['auto', 'none', 'circle', 'square', 'diamond', 'triangle', 'x', 'star', 'plus', 'dash', 'dot'];

const readMarkerFmt = (m: XNode | undefined, seriesCol: string | undefined): MarkerFormat | undefined => {
  if (!m) return undefined;
  const out: MarkerFormat = {};
  const sym = val(m, 'symbol');
  if (sym && sym !== 'circle' && sym !== 'none' && sym !== 'auto' && (SYMBOLS as string[]).includes(sym)) out.symbol = sym as MarkerSymbol;
  const size = num(m, 'size');
  if (size !== undefined && size !== 5) out.size = Math.max(2, Math.min(72, size));
  const sp = kid(m, 'spPr');
  const f = readColor(kid(sp, 'solidFill')).color, l = readColor(kid(kid(sp, 'ln'), 'solidFill')).color;
  if (f && f !== seriesCol) out.fill = f;
  if (l && l !== seriesCol) out.line = l;
  return Object.keys(out).length ? out : undefined;
};

const POS_IN: Record<string, DataLabelOptions['position']> = { bestFit: 'bestFit', b: 'below', ctr: 'center', inBase: 'insideBase', inEnd: 'insideEnd', l: 'left', outEnd: 'outsideEnd', r: 'right', t: 'above' };
const POS_OUT: Record<string, string> = { bestFit: 'bestFit', below: 'b', center: 'ctr', insideBase: 'inBase', insideEnd: 'inEnd', left: 'l', outsideEnd: 'outEnd', right: 'r', above: 't' };

/** Data label settings from a dLbls element; `range` is the series' label range (c15). */
const readLabels = (n: XNode | undefined, range?: string): DataLabelOptions | undefined => {
  if (!n || kid(n, 'delete') && flag(n, 'delete')) return n ? {} : undefined;
  const o: DataLabelOptions = {};
  if (flag(n, 'showVal')) o.value = true;
  if (flag(n, 'showSerName')) o.seriesName = true;
  if (flag(n, 'showCatName')) o.categoryName = true;
  if (flag(n, 'showPercent')) o.percent = true;
  if (flag(n, 'showBubbleSize')) o.bubbleSize = true;
  if (flag(n, 'showLegendKey')) o.legendKey = true;
  const pos = val(n, 'dLblPos');
  if (pos && POS_IN[pos]) o.position = POS_IN[pos];
  const sep = kid(n, 'separator')?.text;
  if (sep !== undefined && sep !== '' && sep !== ', ') o.separator = sep;
  if (kid(n, 'showLeaderLines') && !flag(n, 'showLeaderLines')) o.leaderLines = false;
  const nf = kid(n, 'numFmt');
  const code = attr(nf, 'formatCode');
  if (code && attr(nf, 'sourceLinked') !== '1' && code !== 'General') o.numFmt = code;
  const showRange = findAll(kid(n, 'extLst'), 'showDataLabelsRange').some((x) => attr(x, 'val') !== '0');
  if (showRange && range) o.fromCells = range;
  return o;
};

const TREND_IN: Record<string, Trendline['kind']> = { linear: 'linear', exp: 'exp', log: 'log', poly: 'poly', power: 'power', movingAvg: 'movingAvg' };
const readTrend = (t: XNode): Trendline | undefined => {
  const kind = TREND_IN[val(t, 'trendlineType') ?? ''];
  if (!kind) return undefined;
  const o: Trendline = { kind };
  const name = kid(t, 'name')?.text;
  if (name) o.name = name;
  const ln = readLine(kid(kid(t, 'spPr'), 'ln'));
  if (ln) o.line = ln;
  const order = num(t, 'order'); if (kind === 'poly' && order !== undefined) o.order = Math.max(2, Math.min(6, Math.round(order)));
  const period = num(t, 'period'); if (kind === 'movingAvg' && period !== undefined) o.period = Math.max(2, Math.min(255, Math.round(period)));
  const fw = num(t, 'forward'); if (fw) o.forward = fw;
  const bw = num(t, 'backward'); if (bw) o.backward = bw;
  const ic = num(t, 'intercept'); if (ic !== undefined) o.intercept = ic;
  if (flag(t, 'dispEq')) o.showEquation = true;
  if (flag(t, 'dispRSqr')) o.showR2 = true;
  return o;
};

const ERR_IN: Record<string, ErrorBars['kind']> = { fixedVal: 'fixed', percentage: 'percent', stdDev: 'stdDev', stdErr: 'stdErr', cust: 'custom' };
const readErr = (e: XNode): ErrorBars | undefined => {
  const kind = ERR_IN[val(e, 'errValType') ?? ''];
  if (!kind) return undefined;
  const o: ErrorBars = { kind };
  const dir = val(e, 'errDir'); if (dir === 'x') o.dir = 'x';
  const type = val(e, 'errBarType'); if (type === 'plus' || type === 'minus') o.type = type;
  if (flag(e, 'noEndCap')) o.noCap = true;
  const v = num(e, 'val'); if (v !== undefined && kind !== 'custom' && kind !== 'stdErr') o.value = v;
  if (kind === 'custom') {
    const p = seriesRef(kid(e, 'plus')), m = seriesRef(kid(e, 'minus'));
    if (p) o.plus = p;
    if (m) o.minus = m;
  }
  const ln = readLine(kid(kid(e, 'spPr'), 'ln'));
  if (ln) o.line = ln;
  return o;
};

/** Series whose lines are hidden by the chart type itself (marker-only scatter, stock prices). */
interface ReadCtx { lineImplied: boolean; filled: boolean; group: string }

const readSeries = (ser: XNode, sub: ChartSeries['type'] | undefined, rc: ReadCtx): ChartSeries => {
  const s: ChartSeries = { values: seriesRef(kid(ser, 'val') ?? kid(ser, 'yVal')) ?? '' };
  const tx = kid(ser, 'tx');
  if (tx) { const f = find(tx, 'f'); const v = find(tx, 'v'); const name = f?.text ?? v?.text; if (name) s.name = name; }
  const cat = seriesRef(kid(ser, 'cat') ?? kid(ser, 'xVal'));
  if (cat) s.categories = cat;
  const ml = kid(kid(ser, 'cat'), 'multiLvlStrRef');
  if (ml) {
    const n = numAttr(kid(kid(ml, 'multiLvlStrCache'), 'ptCount'), 'val') ?? 0;
    const levels = kids(kid(ml, 'multiLvlStrCache'), 'lvl').map((l) => {
      const out = Array<string>(Math.min(n, 4096)).fill('');
      for (const p of kids(l, 'pt')) { const i = numAttr(p, 'idx') ?? -1; if (i >= 0 && i < out.length) out[i] = kid(p, 'v')?.text ?? ''; }
      return out;
    });
    if (levels.length) s.multiLevel = { ref: find(ml, 'f')?.text.trim() ?? cat ?? '', levels };
  }
  const color = seriesColor(ser, rc.filled);
  if (color) s.color = color;
  if (sub) s.type = sub;
  const sizes = seriesRef(kid(ser, 'bubbleSize'));
  if (sizes) s.sizes = sizes;
  // Formatting beyond the plain colour.
  const sp = kid(ser, 'spPr');
  if (rc.filled) {
    const fill = readFill(sp);
    if (fill && (fill.gradient || fill.none || fill.transparency !== undefined)) s.fill = fill;
    const ln = readLine(kid(sp, 'ln'));
    if (ln) s.line = ln;
  } else {
    const ln = readLine(kid(sp, 'ln'));
    if (ln) {
      const { color: _c, ...rest } = ln;
      if (rest.width === 2.25) delete rest.width;
      // A hidden line is part of marker-only and stock types.
      if (!(ln.none && rc.lineImplied) && Object.keys(rest).length) s.line = rest;
    }
    const fill = readFill(sp);
    if (fill && (fill.gradient || fill.transparency !== undefined)) s.fill = fill;
  }
  const mk = readMarkerFmt(kid(ser, 'marker'), color);
  if (mk) s.marker = mk;
  if (flag(ser, 'invertIfNegative')) s.invertIfNegative = true;
  const ex = num(ser, 'explosion'); if (ex) s.explosion = ex;
  if (rc.group === 'lineChart' && flag(ser, 'smooth')) s.smooth = true;
  const points: PointFormat[] = kids(ser, 'dPt').map((d) => {
    const p: PointFormat = { idx: num(d, 'idx') ?? 0 };
    const sh = readShape(d);
    if (sh?.fill) p.fill = sh.fill;
    if (sh?.line) p.line = sh.line;
    const pm = readMarkerFmt(kid(d, 'marker'), color);
    if (pm) p.marker = pm;
    const pe = num(d, 'explosion'); if (pe) p.explosion = pe;
    return p;
  }).filter((p) => Object.keys(p).length > 1);
  if (points.length) s.points = points;
  const range = findAll(kid(ser, 'extLst'), 'datalabelsRange').map((r) => find(r, 'f')?.text.trim()).find(Boolean);
  const lbls = kid(ser, 'dLbls');
  if (lbls) { const l = readLabels(lbls, range); if (l) s.labels = l; }
  const trends = kids(ser, 'trendline').map(readTrend).filter((t): t is Trendline => !!t);
  if (trends.length) s.trendlines = trends;
  const errs = kids(ser, 'errBars').map(readErr).filter((e): e is ErrorBars => !!e);
  if (errs.length) s.errorBars = errs;
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
    case 'barChart': {
      const bar = attr(kid(g, 'barDir'), 'val') === 'bar';
      if (stack === 'stacked') return bar ? 'stackedBar' : 'stackedColumn';
      if (stack === 'percent') return bar ? 'percentBar' : 'percentColumn';
      return bar ? 'bar' : 'column';
    }
    case 'bar3DChart': {
      const bar = attr(kid(g, 'barDir'), 'val') === 'bar';
      const shape = val(g, 'shape') ?? kids(g, 'ser').map((s) => val(s, 'shape')).find(Boolean) ?? 'box';
      if (stack === 'stacked') return bar ? 'bar3DStacked' : 'column3DStacked';
      if (stack === 'percent') return bar ? 'bar3DPercent' : 'column3DPercent';
      if (bar) return 'bar3DClustered';
      if (grouping === 'standard') return 'column3D';
      return shape.startsWith('cylinder') ? 'cylinderColumn' : shape.startsWith('cone') ? 'coneColumn' : shape.startsWith('pyramid') ? 'pyramidColumn' : 'column3DClustered';
    }
    case 'lineChart': {
      const m = hasMarkers(g);
      if (stack === 'stacked') return m ? 'stackedLineMarkers' : 'stackedLine';
      if (stack === 'percent') return m ? 'percentLineMarkers' : 'percentLine';
      return m ? 'lineMarkers' : 'line';
    }
    case 'line3DChart': return 'line3D';
    case 'stockChart': return kids(g, 'ser').length >= 4 ? 'stockOHLC' : 'stockHLC';
    case 'areaChart': return stack === 'stacked' ? 'stackedArea' : stack === 'percent' ? 'percentArea' : 'area';
    case 'area3DChart': return stack === 'stacked' ? 'stackedArea3D' : stack === 'percent' ? 'percentArea3D' : 'area3D';
    case 'pieChart': return kids(g, 'ser').some((s) => (num(s, 'explosion') ?? 0) > 0) ? 'pieExploded' : 'pie';
    case 'pie3DChart': return kids(g, 'ser').some((s) => (num(s, 'explosion') ?? 0) > 0) ? 'pie3DExploded' : 'pie3D';
    case 'ofPieChart': return val(g, 'ofPieType') === 'bar' ? 'barOfPie' : 'pieOfPie';
    case 'doughnutChart': return kids(g, 'ser').some((s) => (num(s, 'explosion') ?? 0) > 0) ? 'doughnutExploded' : 'doughnut';
    case 'surfaceChart': return flag(g, 'wireframe') ? 'contourWireframe' : 'contour';
    case 'surface3DChart': return flag(g, 'wireframe') ? 'surfaceWireframe3D' : 'surface3D';
    case 'scatterChart': {
      const style = attr(kid(g, 'scatterStyle'), 'val') ?? 'marker';
      const lines = !lineHidden(g) && style !== 'marker' && style !== 'none';
      const markers = hasMarkers(g);
      const smooth = serFlag(g, 'smooth') || ((style === 'smooth' || style === 'smoothMarker') && !kids(g, 'ser').some((s) => kid(s, 'smooth')));
      if (!lines) return 'scatter';
      if (smooth) return markers ? 'scatterSmoothMarkers' : 'scatterSmooth';
      return markers ? 'scatterLinesMarkers' : 'scatterLines';
    }
    case 'bubbleChart': return flag(g, 'bubble3D') || serFlag(g, 'bubble3D') ? 'bubble3D' : 'bubble';
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

const UNITS = ['hundreds', 'thousands', 'tenThousands', 'hundredThousands', 'millions', 'tenMillions', 'hundredMillions', 'billions', 'trillions'];
const TICKS = ['none', 'out', 'in', 'cross'];

/** Axis settings from a catAx/valAx/dateAx/serAx element. `legacy` puts the value axis basics in the old chart-wide fields. */
const readAxis = (ax: XNode | undefined): AxisOptions => {
  const o: AxisOptions = {};
  if (!ax) return o;
  const sc = kid(ax, 'scaling');
  const lb = num(sc, 'logBase'); if (lb !== undefined && lb >= 2) o.logBase = lb;
  if (val(sc, 'orientation') === 'maxMin') o.reverse = true;
  const mn = num(sc, 'min'), mx = num(sc, 'max');
  if (mn !== undefined) o.min = mn;
  if (mx !== undefined) o.max = mx;
  if (flag(ax, 'delete')) o.hidden = true;
  if (kid(ax, 'majorGridlines')) o.majorGrid = true;
  if (kid(ax, 'minorGridlines')) o.minorGrid = true;
  const nf = kid(ax, 'numFmt');
  const code = attr(nf, 'formatCode');
  if (code && attr(nf, 'sourceLinked') === '0' && code !== 'General') o.numFmt = code;
  const mt = val(ax, 'majorTickMark'); if (mt && mt !== 'out' && TICKS.includes(mt)) o.majorTick = mt as AxisOptions['majorTick'];
  const nt = val(ax, 'minorTickMark'); if (nt && nt !== 'none' && TICKS.includes(nt)) o.minorTick = nt as AxisOptions['minorTick'];
  const lp = val(ax, 'tickLblPos'); if (lp === 'low' || lp === 'high' || lp === 'none') o.labelPos = lp;
  const ln = readLine(kid(kid(ax, 'spPr'), 'ln')); if (ln) o.line = ln;
  const mu = num(ax, 'majorUnit'); if (mu !== undefined && mu > 0 && ax.name === 'valAx') o.majorUnit = mu;
  const nu = num(ax, 'minorUnit'); if (nu !== undefined && nu > 0 && ax.name === 'valAx') o.minorUnit = nu;
  const du = kid(ax, 'dispUnits');
  const bu = val(du, 'builtInUnit');
  if (bu && UNITS.includes(bu)) { o.dispUnits = bu as AxisOptions['dispUnits']; if (!kid(du, 'dispUnitsLbl')) o.dispUnitsLabel = false; }
  if (ax.name === 'dateAx') {
    const bt = val(ax, 'baseTimeUnit');
    if (bt === 'days' || bt === 'months' || bt === 'years') o.baseUnit = bt; else o.dateAxis = true;
  } else if (ax.name === 'catAx' && val(ax, 'auto') === '0') o.dateAxis = false;
  return o;
};

/** Where an axis meets the axis it crosses (crosses/crossesAt on the element). */
const readCross = (ax: XNode | undefined): Pick<AxisOptions, 'crosses' | 'crossesAt'> => {
  const at = num(ax, 'crossesAt');
  if (at !== undefined) return { crossesAt: at };
  const c = val(ax, 'crosses');
  return c === 'min' || c === 'max' ? { crosses: c } : {};
};

const nonEmpty = <T extends object>(o: T): T | undefined => (Object.keys(o).length ? o : undefined);

export const readChart = (root: XNode, id: string, anchor: Chart['anchor']): Chart | undefined => {
  const chart = kid(root, 'chart');
  const plot = kid(chart, 'plotArea');
  if (!plot) return undefined;
  const groups = plot.children.filter((c) => /Chart$/.test(c.name));
  let known = groups.map((g) => ({ g, t: groupType(g) })).filter((x): x is { g: XNode; t: ChartType } => !!x.t && kids(x.g, 'ser').length > 0);
  if (!known.length) return undefined;

  // Volume stock charts: a column group of volumes plus the stock group.
  const stock = known.find((k) => k.g.name === 'stockChart');
  let stockType: ChartType | undefined;
  if (stock) {
    const vol = known.find((k) => k.g.name === 'barChart');
    const open = kids(stock.g, 'ser').length >= 4;
    stockType = vol ? (open ? 'stockVOHLC' : 'stockVHLC') : stock.t;
    known = vol ? [vol, stock] : [stock];
  }

  // Axis pairs: the first group's axes are primary; any other value axis is secondary.
  const axIdsOf = (g: XNode) => kids(g, 'axId').map((a) => attr(a, 'val') ?? '');
  const primaryAx = new Set(axIdsOf(known[0]!.g));
  const onSecondary = (g: XNode) => { const ids = axIdsOf(g); return ids.length > 0 && !ids.some((x) => primaryAx.has(x)); };

  // Groups that differ only in axis (one per axis) are one type, not a combo.
  const combo = !stockType && new Set(known.map((k) => k.t)).size > 1;
  const series: ChartSeries[] = [];
  const secondary: number[] = [];
  for (const { g, t } of known) {
    const sec = onSecondary(g);
    const scatterMarkersOnly = g.name === 'scatterChart' && groupType(g) === 'scatter';
    const rc: ReadCtx = { group: g.name, lineImplied: scatterMarkersOnly || g.name === 'stockChart', filled: !/^(line|line3D|scatter|radar|stock)Chart$/.test(g.name) || (g.name === 'radarChart' && val(g, 'radarStyle') === 'filled') };
    for (const ser of kids(g, 'ser')) {
      if (sec && !stockType) secondary.push(series.length);
      const s = readSeries(ser, combo ? subKind(t) : undefined, rc);
      // Markers are part of the type for line, scatter and radar charts.
      if (s.marker && Object.keys(s.marker).length === 0) delete s.marker;
      series.push(s);
    }
  }
  let type: ChartType = stockType ?? known[0]!.t;
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
  // Exploded types carry the default pull-out; only other amounts are kept per series.
  if (type === 'pieExploded' || type === 'pie3DExploded' || type === 'doughnutExploded') series.forEach((s) => { if (s.explosion === 18) delete s.explosion; });
  // The close series of a high-low-close chart is marked with dashes by default.
  if (type === 'stockHLC' || type === 'stockVHLC') { const last = series.at(-1); if (last?.marker && last.marker.symbol === 'dash' && Object.keys(last.marker).length === 1) delete last.marker; }

  const c: Chart = { id, type, series, anchor };
  const titleNode = kid(chart, 'title');
  const title = richText(titleNode);
  if (title) c.title = title;
  const legend = kid(chart, 'legend');
  c.legend = legend ? LEGEND[attr(kid(legend, 'legendPos'), 'val') ?? 'r'] ?? 'right' : 'none';
  if (legend && flag(legend, 'overlay')) options.legendOverlay = true;

  const axes = plot.children.filter((n) => /Ax$/.test(n.name));
  const axById = (idv: string) => axes.find((a) => attr(kid(a, 'axId'), 'val') === idv);
  const pIds = axIdsOf(known[0]!.g);
  const pAxes = pIds.map(axById).filter((a): a is XNode => !!a);
  const scatterLike = known[0]!.g.name === 'scatterChart' || known[0]!.g.name === 'bubbleChart';
  const catAx = scatterLike
    ? pAxes.find((a) => /^[bt]$/.test(attr(kid(a, 'axPos'), 'val') ?? '')) ?? pAxes[0]
    : pAxes.find((a) => a.name === 'catAx' || a.name === 'dateAx');
  const valAx = pAxes.find((a) => a.name === 'valAx' && a !== catAx);
  const secGroups = stockType && known.length > 1 ? [known[1]!] : known.filter((k) => onSecondary(k.g));
  const secAx = secGroups.flatMap((k) => axIdsOf(k.g)).map(axById).find((a) => a?.name === 'valAx' && (!scatterLike || /^[lr]$/.test(attr(kid(a, 'axPos'), 'val') ?? '')));
  const x = richText(kid(catAx, 'title')), y = richText(kid(valAx, 'title')), y2 = richText(kid(secAx, 'title'));
  if (x || y || y2) c.axisTitles = { ...(x ? { x } : {}), ...(y ? { y } : {}), ...(y2 ? { y2 } : {}) };
  const axOpts: NonNullable<ChartOptions['axes']> = {};
  if (valAx) {
    const a = readAxis(valAx);
    if (a.min !== undefined) { options.min = a.min; delete a.min; }
    if (a.max !== undefined) { options.max = a.max; delete a.max; }
    if (a.numFmt) { options.numFmt = a.numFmt; delete a.numFmt; }
    if (!a.majorGrid) options.gridlines = false;
    delete a.majorGrid;
    // Where the category axis crosses this one is stored on the category axis.
    Object.assign(a, readCross(catAx));
    if (nonEmpty(a)) axOpts.y = a;
  }
  if (catAx) {
    const a = readAxis(catAx);
    if (!scatterLike) { delete a.min; delete a.max; } else if (!kid(catAx, 'majorGridlines')) a.majorGrid = false;
    if (scatterLike && a.majorGrid) delete a.majorGrid;
    Object.assign(a, readCross(valAx));
    if (nonEmpty(a)) axOpts.x = a;
  }
  if (secAx) {
    const a = readAxis(secAx);
    delete a.majorGrid;
    if (nonEmpty(a)) axOpts.y2 = a;
  }
  if (nonEmpty(axOpts)) options.axes = axOpts;

  // Chart-wide labels: group dLbls.
  const groupLbls = known.map(({ g }) => kid(g, 'dLbls')).filter((n): n is XNode => !!n);
  const gl = groupLbls.map((n) => readLabels(n)).find((l) => l && Object.keys(l).length);
  const pieLike = /^(pie|doughnut|pie3D|pieExploded|pie3DExploded|doughnutExploded|pieOfPie|barOfPie)$/.test(type);
  const showFromSeries = series.some((s) => s.labels && (s.labels.value || s.labels.percent)) && !gl;
  if (gl && (gl.value || gl.percent || gl.seriesName || gl.categoryName || gl.bubbleSize)) {
    c.dataLabels = true;
    const { value, percent, ...rest } = gl;
    if (pieLike && !Object.keys(rest).length) {
      if (value && percent) options.labelMode = 'both';
      else if (value) options.labelMode = 'value';
    } else if (!(value && !percent && !Object.keys(rest).length)) options.labels = gl;
  } else if (showFromSeries && pieLike) {
    // Pies saved with per-series labels only.
    const l = series[0]!.labels!;
    c.dataLabels = true;
    if (l.value && l.percent) options.labelMode = 'both'; else if (l.value) options.labelMode = 'value';
    delete series[0]!.labels;
  }

  // Group settings.
  const g0 = known.find((k) => !stockType || k.g.name === 'stockChart')!.g;
  const gw = num(g0, 'gapWidth'); if (gw !== undefined && gw !== 150 && !/ofPie/.test(g0.name)) options.gapWidth = gw;
  const ov = num(g0, 'overlap'); if (ov !== undefined && ov !== (val(g0, 'grouping') === 'clustered' || !val(g0, 'grouping') ? 0 : 100)) options.overlap = ov;
  const vc = val(g0, 'varyColors');
  if (vc !== undefined) { const on = vc === '1' || vc === 'true'; if (on !== pieLike) options.varyColors = on; }
  const fsa = num(g0, 'firstSliceAng'); if (fsa) options.firstSliceAngle = fsa;
  const hs = num(g0, 'holeSize'); if (hs !== undefined && hs !== 50) options.holeSize = hs;
  const sh = val(g0, 'shape');
  if (g0.name === 'bar3DChart' && sh) {
    const shape = sh.startsWith('cylinder') ? 'cylinder' : sh.startsWith('cone') ? 'cone' : sh.startsWith('pyramid') ? 'pyramid' : 'box';
    const implied = type === 'cylinderColumn' ? 'cylinder' : type === 'coneColumn' ? 'cone' : type === 'pyramidColumn' ? 'pyramid' : 'box';
    if (shape !== implied) options.shape = shape;
  }
  if (g0.name === 'ofPieChart') {
    const op: NonNullable<ChartOptions['ofPie']> = {};
    const st = val(g0, 'splitType');
    if (st === 'val' || st === 'percent') op.split = st;
    const sp = num(g0, 'splitPos'); if (sp !== undefined) op.splitValue = sp;
    const ss = num(g0, 'secondPieSize'); if (ss !== undefined && ss !== 75) op.secondSize = ss;
    const og = num(g0, 'gapWidth'); if (og !== undefined && og !== 100) op.gapWidth = og;
    if (nonEmpty(op)) options.ofPie = op;
  }
  const gd = num(g0, 'gapDepth');
  const bs = num(g0, 'bubbleScale'); if (bs !== undefined && bs !== 100) options.bubbleScale = bs;
  const lineGroups = known.map((k) => k.g).filter((g) => /^(line|stock|area|line3D|area3D)Chart$/.test(g.name));
  const isStock = !!stockType;
  if (lineGroups.some((g) => kid(g, 'dropLines'))) options.dropLines = true;
  const hl = lineGroups.some((g) => kid(g, 'hiLowLines'));
  if (hl !== isStock) options.hiLowLines = hl;
  const ud = lineGroups.map((g) => kid(g, 'upDownBars')).find(Boolean);
  const udDefault = type === 'stockOHLC' || type === 'stockVOHLC';
  if (!!ud !== udDefault) options.upDownBars = !!ud;
  if (ud) {
    const up = readFill(kid(kid(ud, 'upBars'), 'spPr')), down = readFill(kid(kid(ud, 'downBars'), 'spPr'));
    if (up?.color) options.upFill = up.color;
    if (down?.color) options.downFill = down.color;
  }
  if (known.some(({ g }) => g.name === 'barChart' && kids(g, 'serLines').length)) options.seriesLines = true;

  // 3D view, walls and floor.
  const v3 = kid(chart, 'view3D');
  if (v3) {
    const v: NonNullable<ChartOptions['view3D']> = {};
    const d3 = /3D|^(cylinder|cone|pyramid)Column$|^column3D/.test(type) || /^surface/.test(type);
    const rx = num(v3, 'rotX'), ry = num(v3, 'rotY'), p = num(v3, 'perspective'), dp = num(v3, 'depthPercent'), hp = num(v3, 'hPercent');
    const pie3 = type === 'pie3D' || type === 'pie3DExploded';
    if (rx !== undefined && rx !== (pie3 ? 30 : 15) && d3) v.rotX = rx;
    if (ry !== undefined && ry !== (pie3 ? 0 : 20) && d3) v.rotY = ry;
    if (p !== undefined && p !== 30 && d3) v.perspective = p;
    if (flag(v3, 'rAngAx') && d3 && !pie3) v.rightAngle = true;
    if (dp !== undefined && dp !== 100 && d3) v.depthPercent = dp;
    if (hp !== undefined && hp !== 100 && d3) v.heightPercent = hp;
    if (gd !== undefined && gd !== 150) v.gapDepth = gd;
    if (nonEmpty(v)) options.view3D = v;
  } else if (gd !== undefined && gd !== 150) options.view3D = { gapDepth: gd };
  const walls = readShape(kid(chart, 'backWall')), floor = readShape(kid(chart, 'floor'));
  if (walls) options.walls = walls;
  if (floor) options.floor = floor;

  // Data table, plot and chart area.
  const dt = kid(plot, 'dTable');
  if (dt) {
    const t: NonNullable<ChartOptions['dataTable']> = {};
    if (!flag(dt, 'showHorzBorder', true)) t.horizontal = false;
    if (!flag(dt, 'showVertBorder', true)) t.vertical = false;
    if (!flag(dt, 'showOutline', true)) t.outline = false;
    if (!flag(dt, 'showKeys', true)) t.keys = false;
    options.dataTable = t;
  }
  const pa = readShape(plot); if (pa) options.plotArea = pa;
  const ca = readShape(root); if (ca) options.chartArea = ca;
  const blanks = val(chart, 'dispBlanksAs'); if (blanks === 'zero' || blanks === 'span') options.blanks = blanks;
  if (kid(chart, 'plotVisOnly') && !flag(chart, 'plotVisOnly')) options.showHidden = true;

  const ext = readExt(root);
  if (ext && ext.base === type) {
    c.type = ext.type;
    if (ext.options) c.options = ext.options; else delete c.options;
    if (!['combo', 'comboSecondary', 'comboAreaColumn'].includes(c.type)) c.series = c.series.map(({ type: _t, ...s }) => s);
  } else if (Object.keys(options).length) c.options = options;
  // PivotChart: the PivotTable it follows ("[Book]Sheet!Name" in the file).
  const pivotName = kid(kid(root, 'pivotSource'), 'name')?.text.trim();
  if (pivotName) c.pivot = pivotName.replace(/^\[[^\]]*\]/, '').replace(/^'(.*)'!/, (_m, n: string) => `${n.replace(/''/g, "'")}!`);
  return c;
};

// ---------------------------------------------------------------- chart writing

const REF_RE = /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

const txRich = (text: string): string =>
  `<c:tx><c:rich><a:bodyPr/><a:lstStyle/>${text.split('\n').map((l) => `<a:p><a:r><a:t>${esc(l)}</a:t></a:r></a:p>`).join('')}</c:rich></c:tx>`;
/** A title: plain text, or a link to a cell written as =Sheet1!A1. */
const titleXml = (text: string | undefined): string => {
  if (!text) return '';
  const ref = text.trim().replace(/^=/, '');
  if (text.trim().startsWith('=') && REF_RE.test(ref)) return `<c:title><c:tx><c:strRef><c:f>${esc(ref)}</c:f></c:strRef></c:tx><c:overlay val="0"/></c:title>`;
  return `<c:title>${txRich(text)}<c:overlay val="0"/></c:title>`;
};
const v = (tag: string, value: string | number): string => `<c:${tag} val="${value}"/>`;
const b = (x: boolean | undefined): number => (x ? 1 : 0);

type SerKind = 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar' | 'bubble' | 'surface';
type PosKind = 'bar' | 'barStacked' | 'line' | 'pie' | 'none';
const POS_OK: Record<PosKind, string[]> = {
  bar: ['ctr', 'inEnd', 'inBase', 'outEnd'], barStacked: ['ctr', 'inEnd', 'inBase'], line: ['ctr', 'l', 'r', 't', 'b'], pie: ['ctr', 'inEnd', 'outEnd', 'bestFit'], none: [],
};

const dLblsXml = (o: DataLabelOptions | null, pos: PosKind, withRange = false): string => {
  const show = o ?? {};
  const p = show.position && show.position !== 'auto' ? POS_OUT[show.position] : undefined;
  return '<c:dLbls>'
    + (show.numFmt ? `<c:numFmt formatCode="${esc(show.numFmt)}" sourceLinked="0"/>` : '')
    + (p && POS_OK[pos].includes(p) ? v('dLblPos', p) : '')
    + v('showLegendKey', b(show.legendKey)) + v('showVal', b(show.value)) + v('showCatName', b(show.categoryName)) + v('showSerName', b(show.seriesName))
    + v('showPercent', b(show.percent)) + v('showBubbleSize', b(show.bubbleSize))
    + (show.separator !== undefined ? `<c:separator>${esc(show.separator)}</c:separator>` : '')
    + (pos === 'pie' ? v('showLeaderLines', show.leaderLines === false ? 0 : 1) : '')
    + (withRange ? `<c:extLst><c:ext uri="${EXT_SHOW_LABEL_RANGE}" xmlns:c15="${C15}"><c15:showDataLabelsRange val="1"/></c:ext></c:extLst>` : '')
    + '</c:dLbls>';
};

/** Categories on several levels (§21.2.2.102 multiLvlStrRef), innermost level first. */
const multiLevelXml = (m: NonNullable<ChartSeries['multiLevel']>): string => {
  const n = Math.max(0, ...m.levels.map((l) => l.length));
  const lvls = m.levels.map((l) => `<c:lvl>${l.map((v, i) => (v === '' ? '' : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`)).join('')}</c:lvl>`).join('');
  return `<c:multiLvlStrRef><c:f>${esc(m.ref.replace(/^=/, ''))}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls}</c:multiLvlStrCache></c:multiLvlStrRef>`;
};

const refXml = (ref: string | undefined, isNum: boolean): string => {
  if (!ref) return '';
  const lit = /^=?\s*\{([^{}]*)\}\s*$/.exec(ref);
  if (lit) {
    const parts = lit[1]!.split(/[,;]/).map((x) => x.trim());
    if (isNum) return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${parts.length}"/>${parts.map((p, i) => `<c:pt idx="${i}"><c:v>${esc(String(Number(p) || 0))}</c:v></c:pt>`).join('')}</c:numLit>`;
    return `<c:strLit><c:ptCount val="${parts.length}"/>${parts.map((p, i) => `<c:pt idx="${i}"><c:v>${esc(p.replace(/^"|"$/g, '').replace(/""/g, '"'))}</c:v></c:pt>`).join('')}</c:strLit>`;
  }
  const f = ref.replace(/^=/, '');
  return isNum ? `<c:numRef><c:f>${esc(f)}</c:f></c:numRef>` : `<c:strRef><c:f>${esc(f)}</c:f></c:strRef>`;
};

interface SerStyle { marker?: MarkerSymbol; markers?: boolean; smooth?: boolean; lines?: boolean; filled?: boolean; stock?: boolean; pos: PosKind; noErrDir?: boolean; explosion?: number; shape?: string; trend?: boolean }

const markerXmlOf = (m: MarkerFormat | undefined, on: boolean, color: string | undefined, dflt?: MarkerSymbol): string => {
  const sym = m?.symbol && m.symbol !== 'auto' ? m.symbol : on ? dflt ?? 'circle' : 'none';
  if (sym === 'none') return `<c:marker>${v('symbol', 'none')}</c:marker>`;
  const fill = m?.fill ?? color, line = m?.line ?? color;
  const sp = spPr(solid(fill), hex(line) ? `<a:ln>${solid(line)}</a:ln>` : '');
  return `<c:marker>${v('symbol', sym)}${v('size', Math.round(Math.max(2, Math.min(72, m?.size ?? 5))))}${sp}</c:marker>`;
};

const trendXml = (t: Trendline): string => {
  const shown = t.showEquation || t.showR2;
  return '<c:trendline>' + (t.name ? `<c:name>${esc(t.name)}</c:name>` : '')
    + (t.line ? spPr('', lnXml(t.line)) : '')
    + v('trendlineType', t.kind)
    + (t.kind === 'poly' ? v('order', Math.max(2, Math.min(6, Math.round(t.order ?? 2)))) : '')
    + (t.kind === 'movingAvg' ? v('period', Math.max(2, Math.min(255, Math.round(t.period ?? 2)))) : '')
    + (t.forward && t.kind !== 'movingAvg' ? v('forward', t.forward) : '') + (t.backward && t.kind !== 'movingAvg' ? v('backward', t.backward) : '')
    + (t.intercept !== undefined && ['linear', 'exp', 'poly'].includes(t.kind) ? v('intercept', t.intercept) : '')
    + v('dispRSqr', b(t.showR2)) + v('dispEq', b(t.showEquation))
    + (shown ? '<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/></c:trendlineLbl>' : '')
    + '</c:trendline>';
};

const ERR_OUT: Record<ErrorBars['kind'], string> = { fixed: 'fixedVal', percent: 'percentage', stdDev: 'stdDev', stdErr: 'stdErr', custom: 'cust' };
const errXml = (e: ErrorBars, withDir: boolean): string =>
  '<c:errBars>' + (withDir ? v('errDir', e.dir ?? 'y') : '') + v('errBarType', e.type ?? 'both') + v('errValType', ERR_OUT[e.kind]) + v('noEndCap', b(e.noCap))
  + (e.kind === 'custom' ? `${e.plus ? `<c:plus>${refXml(e.plus, true)}</c:plus>` : ''}${e.minus ? `<c:minus>${refXml(e.minus, true)}</c:minus>` : ''}` : '')
  + (e.kind !== 'custom' && e.kind !== 'stdErr' ? v('val', e.value ?? (e.kind === 'percent' ? 5 : 1)) : '')
  + (e.line ? spPr('', lnXml(e.line)) : '') + '</c:errBars>';

const serXml = (s: ChartSeries, i: number, kind: SerKind, st: SerStyle): string => {
  let x = v('idx', i) + v('order', i);
  const nm = s.name?.replace(/^=/, '');
  if (s.name) x += REF_RE.test(nm!) ? `<c:tx>${refXml(nm, false)}</c:tx>` : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
  // Without a colour of its own, a series takes the theme's automatic colour.
  const color = s.color ?? s.fill?.color;
  if (kind === 'line' || (kind === 'radar' && !st.filled)) {
    const ln = st.stock && !s.line ? '<a:ln w="19050"><a:noFill/></a:ln>' : lnXml(s.line, color, 28575, true);
    x += spPr(fillXml(s.fill?.gradient || s.fill?.transparency !== undefined ? s.fill : undefined), ln) + markerXmlOf(s.marker, !!st.markers, color, st.marker);
  } else if (kind === 'scatter') {
    const ln = st.lines ? lnXml(s.line, color, 28575, true) : s.line && !s.line.none ? lnXml(s.line, color, 19050, true) : '<a:ln w="19050" cap="rnd"><a:noFill/></a:ln>';
    x += spPr('', ln) + markerXmlOf(s.marker, !!st.markers, color);
  } else if (kind === 'radar') x += spPr(fillXml(s.fill, color), lnXml(s.line)) + markerXmlOf(undefined, false, color);
  else if (kind === 'pie') x += spPr(s.fill ? fillXml(s.fill) : '', lnXml(s.line));
  else x += spPr(fillXml(s.fill, color), lnXml(s.line));
  if (kind === 'bar' || kind === 'bubble') x += v('invertIfNegative', b(s.invertIfNegative));
  if (kind === 'pie') { const ex = s.explosion ?? st.explosion; if (ex) x += v('explosion', Math.round(ex)); }
  const pts = kind === 'surface' ? [] : (s.points ?? []).filter((p) => p.idx >= 0);
  for (const p of pts.sort((a, b2) => a.idx - b2.idx)) {
    let d = v('idx', p.idx);
    if (kind === 'bar' || kind === 'bubble') d += v('invertIfNegative', b(s.invertIfNegative));
    if ((kind === 'line' || kind === 'scatter' || kind === 'radar') && p.marker) d += markerXmlOf(p.marker, true, color);
    if (kind === 'bubble') d += v('bubble3D', 0);
    if (kind === 'pie' && p.explosion !== undefined) d += v('explosion', Math.round(p.explosion));
    d += spPr(fillXml(p.fill), lnXml(p.line));
    x += `<c:dPt>${d}</c:dPt>`;
  }
  if (s.labels && kind !== 'surface') x += dLblsXml(s.labels, st.pos, !!s.labels.fromCells);
  if (st.trend && s.trendlines?.length && (kind === 'bar' || kind === 'line' || kind === 'area' || kind === 'scatter' || kind === 'bubble')) x += s.trendlines.slice(0, 8).map(trendXml).join('');
  if (s.errorBars?.length && (kind === 'bar' || kind === 'line' || kind === 'area' || kind === 'scatter' || kind === 'bubble')) {
    const xy = kind === 'scatter' || kind === 'bubble';
    const list = xy ? [s.errorBars.find((e) => (e.dir ?? 'y') === 'x'), s.errorBars.find((e) => (e.dir ?? 'y') === 'y')].filter((e): e is ErrorBars => !!e) : s.errorBars.filter((e) => (e.dir ?? 'y') === 'y').slice(0, 1);
    x += list.map((e) => errXml(e, xy)).join('');
  }
  if (kind === 'scatter' || kind === 'bubble') {
    x += `${s.categories ? `<c:xVal>${refXml(s.categories, true)}</c:xVal>` : ''}<c:yVal>${refXml(s.values, true)}</c:yVal>`;
    if (kind === 'bubble') x += `<c:bubbleSize>${s.sizes ? refXml(s.sizes, true) : refXml(s.values, true)}</c:bubbleSize>${v('bubble3D', b(st.filled))}`;
    else x += v('smooth', b(s.smooth ?? st.smooth));
  } else {
    x += s.multiLevel ? `<c:cat>${multiLevelXml(s.multiLevel)}</c:cat>` : s.categories ? `<c:cat>${refXml(s.categories, false)}</c:cat>` : '';
    x += `<c:val>${refXml(s.values, true)}</c:val>`;
    if (kind === 'bar' && st.shape) x += v('shape', st.shape);
    if (kind === 'line') x += v('smooth', b(s.smooth));
  }
  if (s.labels?.fromCells) x += `<c:extLst><c:ext uri="${EXT_LABEL_RANGE}" xmlns:c15="${C15}"><c15:datalabelsRange><c15:f>${esc(s.labels.fromCells.replace(/^=/, ''))}</c15:f></c15:datalabelsRange></c:ext></c:extLst>`;
  return `<c:ser>${x}</c:ser>`;
};

const AX_CAT = 500000001, AX_VAL = 500000002, AX_CAT2 = 500000003, AX_VAL2 = 500000004, AX_SER = 500000005;

interface AxisPlan { horizontal: boolean; scatter: boolean; secondary: boolean; serAx: boolean; dateDefault?: boolean }

const scalingXml = (a: AxisOptions, bounds: boolean): string =>
  `<c:scaling>${a.logBase && a.logBase >= 2 ? v('logBase', a.logBase) : ''}${v('orientation', a.reverse ? 'maxMin' : 'minMax')}${bounds && a.max !== undefined ? v('max', a.max) : ''}${bounds && a.min !== undefined ? v('min', a.min) : ''}</c:scaling>`;

const crossXml = (a: Pick<AxisOptions, 'crosses' | 'crossesAt'>, dflt = 'autoZero'): string =>
  a.crossesAt !== undefined ? v('crossesAt', a.crossesAt) : v('crosses', a.crosses === 'min' || a.crosses === 'max' ? a.crosses : dflt);

const axesXml = (c: Chart, p: AxisPlan): string => {
  const catPos = p.horizontal ? 'l' : 'b', valPos = p.horizontal ? 'b' : 'l';
  const o = c.options ?? {};
  const ax = o.axes ?? {};
  const yA: AxisOptions = { ...(ax.y ?? {}), ...(o.min !== undefined ? { min: o.min } : {}), ...(o.max !== undefined ? { max: o.max } : {}), ...(o.numFmt ? { numFmt: o.numFmt } : {}) };
  const xA: AxisOptions = ax.x ?? {};
  const y2A: AxisOptions = ax.y2 ?? {};
  const common = (id: number, cross: number, pos: string, a: AxisOptions, title: string | undefined, grid: { major: boolean; minor: boolean }, bounds: boolean, del: boolean, crossWith: Pick<AxisOptions, 'crosses' | 'crossesAt'>, crossDefault = 'autoZero'): string =>
    v('axId', id) + scalingXml(a, bounds) + v('delete', b(del || a.hidden)) + v('axPos', pos)
    + (grid.major ? '<c:majorGridlines/>' : '') + (grid.minor ? '<c:minorGridlines/>' : '') + titleXml(title)
    + (a.numFmt ? `<c:numFmt formatCode="${esc(a.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>')
    + v('majorTickMark', a.majorTick ?? 'out') + v('minorTickMark', a.minorTick ?? 'none') + v('tickLblPos', a.labelPos ?? 'nextTo')
    + (a.line ? spPr('', lnXml(a.line)) : '')
    + v('crossAx', cross) + crossXml(crossWith, crossDefault);
  const valTail = (a: AxisOptions, between: string) => v('crossBetween', between)
    + (a.majorUnit ? v('majorUnit', a.majorUnit) : '') + (a.minorUnit ? v('minorUnit', a.minorUnit) : '')
    + (a.dispUnits ? `<c:dispUnits>${v('builtInUnit', a.dispUnits)}${a.dispUnitsLabel === false ? '' : '<c:dispUnitsLbl/>'}</c:dispUnits>` : '');
  const yGrid = { major: o.gridlines !== false && yA.majorGrid !== false, minor: !!yA.minorGrid };
  const xGrid = { major: p.scatter ? xA.majorGrid !== false : !!xA.majorGrid, minor: !!xA.minorGrid };
  const ycross = { crosses: yA.crosses, crossesAt: yA.crossesAt }, xcross = { crosses: xA.crosses, crossesAt: xA.crossesAt };
  const date = !p.scatter && (xA.dateAxis === true || !!xA.baseUnit);
  const catTail = `${v('auto', xA.dateAxis === false ? 0 : 1)}${v('lblAlgn', 'ctr')}${v('lblOffset', 100)}${v('noMultiLvlLbl', 0)}`;
  const cat = p.scatter
    ? `<c:valAx>${common(AX_CAT, AX_VAL, catPos, xA, c.axisTitles?.x, xGrid, true, false, ycross)}${valTail(xA, 'midCat')}</c:valAx>`
    : date
      ? `<c:dateAx>${common(AX_CAT, AX_VAL, catPos, xA, c.axisTitles?.x, xGrid, false, false, ycross)}${v('auto', 1)}${v('lblOffset', 100)}${xA.baseUnit ? v('baseTimeUnit', xA.baseUnit) : ''}</c:dateAx>`
      : `<c:catAx>${common(AX_CAT, AX_VAL, catPos, xA, c.axisTitles?.x, xGrid, false, false, ycross)}${catTail}</c:catAx>`;
  const val = `<c:valAx>${common(AX_VAL, AX_CAT, valPos, yA, c.axisTitles?.y, yGrid, true, false, xcross)}${valTail(yA, p.scatter ? 'midCat' : 'between')}</c:valAx>`;
  const ser = p.serAx ? `<c:serAx>${v('axId', AX_SER)}<c:scaling>${v('orientation', 'minMax')}</c:scaling>${v('delete', 0)}${v('axPos', 'b')}<c:numFmt formatCode="General" sourceLinked="1"/>${v('majorTickMark', 'out')}${v('minorTickMark', 'none')}${v('tickLblPos', 'nextTo')}${v('crossAx', AX_VAL)}${v('crosses', 'autoZero')}</c:serAx>` : '';
  if (!p.secondary) return cat + val + ser;
  const cat2 = p.scatter
    ? `<c:valAx>${common(AX_CAT2, AX_VAL2, 'b', {}, undefined, { major: false, minor: false }, false, true, {})}${v('crossBetween', 'midCat')}</c:valAx>`
    : `<c:catAx>${common(AX_CAT2, AX_VAL2, 'b', {}, undefined, { major: false, minor: false }, false, true, {})}${v('auto', 1)}${v('lblAlgn', 'ctr')}${v('lblOffset', 100)}${v('noMultiLvlLbl', 0)}</c:catAx>`;
  const val2 = `<c:valAx>${common(AX_VAL2, AX_CAT2, 'r', y2A, c.axisTitles?.y2, { major: false, minor: !!y2A.minorGrid }, true, false, {}, 'max')}${valTail(y2A, 'between')}</c:valAx>`;
  return cat + val + ser + cat2 + val2;
};

/** The DrawingML type a chart is saved as. */
export const savedAs = (t: ChartType): ChartType => (EXTENDED.includes(t) ? 'column' : t);

const comboKindOf = (c: Chart, s: ChartSeries, i: number): NonNullable<ChartSeries['type']> => {
  if (s.type) return s.type;
  if (c.type === 'comboAreaColumn') return i < Math.max(1, Math.ceil(c.series.length / 2)) && !(c.series.length > 1 && i === c.series.length - 1) ? 'area' : 'column';
  return i === 0 ? 'column' : 'line';
};

/** Option fields with a native DrawingML form; any other field needs the extension. */
const NATIVE = new Set<keyof ChartOptions>([
  'secondary', 'gridlines', 'min', 'max', 'numFmt', 'axes', 'view3D', 'shape', 'gapWidth', 'overlap', 'varyColors', 'firstSliceAngle', 'holeSize', 'ofPie',
  'upDownBars', 'upFill', 'downFill', 'hiLowLines', 'dropLines', 'seriesLines', 'dataTable', 'labels', 'legendOverlay', 'chartArea', 'plotArea', 'walls', 'floor',
  'blanks', 'showHidden', 'bubbleScale', 'labelMode',
]);

const SHAPE_OUT: Record<string, string> = { box: 'box', cylinder: 'cylinder', cone: 'cone', pyramid: 'pyramid' };

/** "Sheet 2!PivotTable1" as a file reference, quoting the sheet name when needed. */
const pivotSourceName = (key: string): string => {
  const i = key.lastIndexOf('!');
  const sheet = key.slice(0, i), name = key.slice(i + 1);
  return `${/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}!${name}`;
};

export const chartXml = (c: Chart): string => {
  const o = c.options ?? {};
  const type = savedAs(c.type);
  const chartLbls = c.dataLabels ? (o.labels ?? { value: true }) : null;
  const dl = (pos: PosKind) => dLblsXml(chartLbls, pos);
  const pieDl = () => {
    if (!c.dataLabels) return dLblsXml(null, 'pie');
    if (o.labels) return dLblsXml(o.labels, 'pie');
    const mode = o.labelMode ?? 'percent';
    return dLblsXml({ value: mode !== 'percent', percent: mode !== 'value' }, 'pie');
  };
  const ax = (sec: boolean, ser = false) => (sec ? v('axId', AX_CAT2) + v('axId', AX_VAL2) : v('axId', AX_CAT) + v('axId', AX_VAL) + (ser ? v('axId', AX_SER) : ''));
  const sec = new Set(o.secondary ?? (c.type === 'comboSecondary' ? c.series.flatMap((s, i) => (comboKindOf(c, s, i) === 'line' ? [i] : [])) : []));
  const indexed = c.series.map((s, i) => ({ s, i }));
  const grp = (g: string) => (g === 'stacked' ? 'stacked' : g === 'percent' ? 'percentStacked' : g);
  const lineEls = () => (o.dropLines ? '<c:dropLines/>' : '') + (o.hiLowLines ? '<c:hiLowLines/>' : '') + (o.upDownBars ? upDownXml() : '');
  const upDownXml = () => `<c:upDownBars>${v('gapWidth', 150)}<c:upBars>${spPr(solid(o.upFill), '')}</c:upBars><c:downBars>${spPr(solid(o.downFill), '')}</c:downBars></c:upDownBars>`;
  const bar = (dir: 'col' | 'bar', grouping: string, list: { s: ChartSeries; i: number }[], secondary = false): string => {
    const stackedG = grouping !== 'clustered';
    return `<c:barChart>${v('barDir', dir)}${v('grouping', grp(grouping))}${v('varyColors', b(o.varyColors))}${list.map(({ s, i }) => serXml(s, i, 'bar', { pos: stackedG ? 'barStacked' : 'bar', trend: !stackedG })).join('')}${dl(stackedG ? 'barStacked' : 'bar')}`
      + `${v('gapWidth', Math.round(o.gapWidth ?? 150))}${stackedG ? v('overlap', Math.round(o.overlap ?? 100)) : o.overlap !== undefined ? v('overlap', Math.round(o.overlap)) : ''}${stackedG && o.seriesLines ? '<c:serLines/>' : ''}${ax(secondary)}</c:barChart>`;
  };
  const line = (grouping: string, list: { s: ChartSeries; i: number }[], markers: boolean, secondary = false): string =>
    `<c:lineChart>${v('grouping', grp(grouping))}${v('varyColors', b(o.varyColors))}${list.map(({ s, i }) => serXml(s, i, 'line', { markers, pos: 'line', trend: grouping === 'standard' })).join('')}${dl('line')}${lineEls()}${v('marker', 1)}${ax(secondary)}</c:lineChart>`;
  const area = (grouping: string, list: { s: ChartSeries; i: number }[], secondary = false): string =>
    `<c:areaChart>${v('grouping', grp(grouping))}${v('varyColors', b(o.varyColors))}${list.map(({ s, i }) => serXml(s, i, 'area', { pos: 'none', trend: grouping === 'standard' })).join('')}${dl('none')}${o.dropLines ? '<c:dropLines/>' : ''}${ax(secondary)}</c:areaChart>`;
  /** One group per axis for types that may put some series on the right axis. */
  const split = (make: (list: { s: ChartSeries; i: number }[], secondary: boolean) => string): string => {
    const a = indexed.filter((x) => !sec.has(x.i)), bb = indexed.filter((x) => sec.has(x.i));
    if (!a.length) return make(bb, false);
    return make(a, false) + (bb.length ? make(bb, true) : '');
  };

  let body = '';
  const plan: AxisPlan = { horizontal: false, scatter: false, secondary: false, serAx: false };
  let axes = true;
  let view3D = '';
  const stackOf = (t: ChartType) => (/^stacked|Stacked$|^stackedArea3D/.test(t) ? 'stacked' : /^percent|Percent$|^percentArea3D/.test(t) ? 'percent' : '');
  const v3 = o.view3D ?? {};
  const view = (d: { rotX: number; rotY: number; persp: number; rAng?: boolean }) =>
    `<c:view3D>${v('rotX', Math.round(v3.rotX ?? d.rotX))}${v3.heightPercent !== undefined ? v('hPercent', Math.round(v3.heightPercent)) : ''}${v('rotY', Math.round(v3.rotY ?? d.rotY))}${v3.depthPercent !== undefined ? v('depthPercent', Math.round(v3.depthPercent)) : ''}${v('rAngAx', b(v3.rightAngle ?? d.rAng))}${v('perspective', Math.round(v3.perspective ?? d.persp))}</c:view3D>`;
  switch (type) {
    case 'column': body = split((l, s) => bar('col', 'clustered', l, s)); break;
    case 'stackedColumn': case 'percentColumn': body = split((l, s) => bar('col', stackOf(type), l, s)); break;
    case 'bar': case 'stackedBar': case 'percentBar': body = bar('bar', stackOf(type) || 'clustered', indexed); plan.horizontal = true; break;
    case 'line': case 'lineMarkers': body = split((l, s) => line('standard', l, type === 'lineMarkers', s)); break;
    case 'stackedLine': case 'percentLine': case 'stackedLineMarkers': case 'percentLineMarkers':
      body = split((l, s) => line(stackOf(type), l, /Markers$/.test(type), s)); break;
    case 'area': body = split((l, s) => area('standard', l, s)); break;
    case 'stackedArea': case 'percentArea': body = split((l, s) => area(stackOf(type), l, s)); break;
    case 'pie': case 'pieExploded': case 'doughnut': case 'doughnutExploded': case 'pie3D': case 'pie3DExploded': {
      axes = false;
      const dough = type.startsWith('doughnut');
      const exploded = /Exploded$/.test(type) ? 18 : undefined;
      const list = dough ? c.series : c.series.slice(0, 1);
      const vary = v('varyColors', o.varyColors === false ? 0 : 1);
      const sers = list.map((s, i) => serXml(s, i, 'pie', { pos: 'pie', explosion: exploded })).join('');
      const fsa = v('firstSliceAng', Math.round(o.firstSliceAngle ?? 0));
      if (type.startsWith('pie3D')) { view3D = view({ rotX: 30, rotY: 0, persp: 30 }); body = `<c:pie3DChart>${vary}${sers}${pieDl()}</c:pie3DChart>`; }
      else body = dough ? `<c:doughnutChart>${vary}${sers}${pieDl()}${fsa}${v('holeSize', Math.round(o.holeSize ?? 50))}</c:doughnutChart>` : `<c:pieChart>${vary}${sers}${pieDl()}${fsa}</c:pieChart>`;
      break;
    }
    case 'pieOfPie': case 'barOfPie': {
      axes = false;
      const op = o.ofPie ?? {};
      const splitType = op.split ?? 'pos';
      const sers = c.series.slice(0, 1).map((s, i) => serXml(s, i, 'pie', { pos: 'pie' })).join('');
      body = `<c:ofPieChart>${v('ofPieType', type === 'barOfPie' ? 'bar' : 'pie')}${v('varyColors', o.varyColors === false ? 0 : 1)}${sers}${pieDl()}${v('gapWidth', Math.round(op.gapWidth ?? 100))}`
        + `${v('splitType', splitType)}${op.splitValue !== undefined ? v('splitPos', op.splitValue) : ''}${v('secondPieSize', Math.round(op.secondSize ?? 75))}<c:serLines/></c:ofPieChart>`;
      break;
    }
    case 'scatter': case 'scatterSmoothMarkers': case 'scatterSmooth': case 'scatterLinesMarkers': case 'scatterLines': {
      plan.scatter = true;
      const st: SerStyle = {
        markers: type === 'scatter' || type === 'scatterSmoothMarkers' || type === 'scatterLinesMarkers',
        smooth: type === 'scatterSmooth' || type === 'scatterSmoothMarkers',
        lines: type !== 'scatter', pos: 'line', trend: true,
      };
      const style = !st.lines ? 'lineMarker' : st.smooth ? (st.markers ? 'smoothMarker' : 'smooth') : st.markers ? 'lineMarker' : 'line';
      body = split((l, s) => `<c:scatterChart>${v('scatterStyle', style)}${v('varyColors', 0)}${l.map(({ s: ss, i }) => serXml(ss, i, 'scatter', st)).join('')}${dl('line')}${ax(s)}</c:scatterChart>`);
      break;
    }
    case 'bubble': case 'bubble3D': {
      plan.scatter = true;
      const d3 = type === 'bubble3D';
      body = `<c:bubbleChart>${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'bubble', { pos: 'line', filled: d3, trend: true })).join('')}${dl('line')}${v('bubble3D', b(d3))}${v('bubbleScale', Math.round(o.bubbleScale ?? 100))}${v('showNegBubbles', 0)}${ax(false)}</c:bubbleChart>`;
      break;
    }
    case 'radar': case 'radarMarkers': case 'radarFilled': {
      const style = type === 'radarFilled' ? 'filled' : type === 'radarMarkers' ? 'marker' : 'standard';
      body = `<c:radarChart>${v('radarStyle', style)}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'radar', { markers: type === 'radarMarkers', filled: type === 'radarFilled', pos: 'none' })).join('')}${dl('none')}${ax(false)}</c:radarChart>`;
      break;
    }
    case 'stockHLC': case 'stockOHLC': case 'stockVHLC': case 'stockVOHLC': {
      const vol = type === 'stockVHLC' || type === 'stockVOHLC';
      const open = type === 'stockOHLC' || type === 'stockVOHLC';
      const prices = indexed.slice(vol ? 1 : 0, (vol ? 1 : 0) + (open ? 4 : 3));
      const hl = o.hiLowLines ?? true, ud = o.upDownBars ?? open;
      const last = prices.at(-1)?.i;
      const stockBody = `<c:stockChart>${prices.map(({ s, i }) => serXml(s, i, 'line', { markers: !open && i === last, marker: 'dash', stock: true, pos: 'line', trend: true })).join('')}${dl('line')}${o.dropLines ? '<c:dropLines/>' : ''}${hl ? '<c:hiLowLines/>' : ''}${ud ? upDownXml() : ''}${ax(vol)}</c:stockChart>`;
      if (vol) {
        body = `<c:barChart>${v('barDir', 'col')}${v('grouping', 'clustered')}${v('varyColors', 0)}${serXml(c.series[0]!, 0, 'bar', { pos: 'bar' })}${dl('bar')}${v('gapWidth', Math.round(o.gapWidth ?? 150))}${ax(false)}</c:barChart>${stockBody}`;
        plan.secondary = true;
      } else body = stockBody;
      break;
    }
    case 'surface3D': case 'surfaceWireframe3D': case 'contour': case 'contourWireframe': {
      const contour = type.startsWith('contour');
      const wire = /Wireframe/.test(type);
      plan.serAx = true;
      view3D = contour ? `<c:view3D>${v('rotX', 90)}${v('rotY', 0)}${v('rAngAx', 0)}${v('perspective', 0)}</c:view3D>` : view({ rotX: 15, rotY: 20, persp: 30 });
      const sers = c.series.map((s, i) => serXml(s, i, 'surface', { pos: 'none' })).join('');
      body = contour ? `<c:surfaceChart>${v('wireframe', b(wire))}${sers}${ax(false, true)}</c:surfaceChart>` : `<c:surface3DChart>${v('wireframe', b(wire))}${sers}${ax(false, true)}</c:surface3DChart>`;
      break;
    }
    case 'column3DClustered': case 'column3DStacked': case 'column3DPercent': case 'column3D': case 'cylinderColumn': case 'coneColumn': case 'pyramidColumn':
    case 'bar3DClustered': case 'bar3DStacked': case 'bar3DPercent': {
      const horizontal = type.startsWith('bar3D');
      const rows = type === 'column3D';
      const st = stackOf(type);
      const implied = type === 'cylinderColumn' ? 'cylinder' : type === 'coneColumn' ? 'cone' : type === 'pyramidColumn' ? 'pyramid' : 'box';
      const shape = SHAPE_OUT[o.shape ?? implied] ?? 'box';
      const fileShape = st && (shape === 'cone' || shape === 'pyramid') ? `${shape}ToMax` : shape;
      plan.horizontal = horizontal;
      plan.serAx = rows;
      view3D = view({ rotX: 15, rotY: 20, persp: 30 });
      const grouping = rows ? 'standard' : st ? grp(st) : 'clustered';
      body = `<c:bar3DChart>${v('barDir', horizontal ? 'bar' : 'col')}${v('grouping', grouping)}${v('varyColors', b(o.varyColors))}${indexed.map(({ s, i }) => serXml(s, i, 'bar', { pos: 'none', trend: false })).join('')}${dl('none')}`
        + `${v('gapWidth', Math.round(o.gapWidth ?? 150))}${v3.gapDepth !== undefined ? v('gapDepth', Math.round(v3.gapDepth)) : ''}${v('shape', fileShape)}${ax(false, rows)}</c:bar3DChart>`;
      break;
    }
    case 'line3D': {
      plan.serAx = true;
      view3D = view({ rotX: 15, rotY: 20, persp: 30 });
      body = `<c:line3DChart>${v('grouping', 'standard')}${v('varyColors', b(o.varyColors))}${indexed.map(({ s, i }) => serXml(s, i, 'line', { pos: 'none' })).join('')}${dl('none')}${o.dropLines ? '<c:dropLines/>' : ''}${v3.gapDepth !== undefined ? v('gapDepth', Math.round(v3.gapDepth)) : ''}${ax(false, true)}</c:line3DChart>`;
      break;
    }
    case 'area3D': case 'stackedArea3D': case 'percentArea3D': {
      const st = stackOf(type);
      plan.serAx = !st;
      view3D = view({ rotX: 15, rotY: 20, persp: 30 });
      body = `<c:area3DChart>${v('grouping', st ? grp(st) : 'standard')}${v('varyColors', b(o.varyColors))}${indexed.map(({ s, i }) => serXml(s, i, 'area', { pos: 'none' })).join('')}${dl('none')}${o.dropLines ? '<c:dropLines/>' : ''}${v3.gapDepth !== undefined ? v('gapDepth', Math.round(v3.gapDepth)) : ''}${ax(false, !st)}</c:area3DChart>`;
      break;
    }
    case 'combo': case 'comboSecondary': case 'comboAreaColumn': {
      // Areas first so columns and lines draw over them.
      for (const secondary of [false, true]) {
        const of = (k: string) => indexed.filter(({ s, i }) => comboKindOf(c, s, i) === k && sec.has(i) === secondary);
        const a = of('area'), bb = of('column'), l = of('line');
        if (a.length) body += area(c.type === 'comboAreaColumn' ? 'stacked' : 'standard', a, secondary);
        if (bb.length) body += bar('col', 'clustered', bb, secondary);
        if (l.length) body += line('standard', l, true, secondary);
      }
      break;
    }
    default: body = bar('col', 'clustered', indexed);
  }
  plan.secondary = plan.secondary || (sec.size > 0 && body.includes(`val="${AX_VAL2}"`));
  const legend = c.legend === 'none' ? '' : `<c:legend>${v('legendPos', { right: 'r', bottom: 'b', top: 't', left: 'l' }[c.legend ?? 'right'])}<c:overlay val="${b(o.legendOverlay)}"/></c:legend>`;
  const walls = view3D && (o.walls || o.floor)
    ? `<c:floor>${v('thickness', 0)}${shapeXml(o.floor)}</c:floor><c:sideWall>${v('thickness', 0)}${shapeXml(o.walls)}</c:sideWall><c:backWall>${v('thickness', 0)}${shapeXml(o.walls)}</c:backWall>` : '';
  const dt = o.dataTable && axes && !plan.scatter
    ? `<c:dTable>${v('showHorzBorder', b(o.dataTable.horizontal !== false))}${v('showVertBorder', b(o.dataTable.vertical !== false))}${v('showOutline', b(o.dataTable.outline !== false))}${v('showKeys', b(o.dataTable.keys !== false))}</c:dTable>` : '';
  // The extension is written when the file alone can't bring the chart back as it is.
  const needsExt = type !== c.type || Object.keys(o).some((k) => !NATIVE.has(k as keyof ChartOptions)) || (o.labelMode !== undefined && !/pie|doughnut|Pie/.test(type));
  const ext = needsExt
    ? `<c:extLst><c:ext uri="${CLERDO_EXT_URI}"><cg:chart xmlns:cg="${CLERDO_NS}" type="${c.type}" base="${type}"${c.options ? ` options="${esc(JSON.stringify(c.options))}"` : ''}/></c:ext></c:extLst>`
    : '';
  const blanks = o.blanks ?? 'gap';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + v('roundedCorners', 0)
    + (c.pivot ? `<c:pivotSource><c:name>${esc(`[Book1]${pivotSourceName(c.pivot)}`)}</c:name><c:fmtId val="0"/></c:pivotSource>` : '')
    + '<c:chart>' + (c.title ? titleXml(c.title) + v('autoTitleDeleted', 0) : v('autoTitleDeleted', 1))
    + view3D + walls
    + `<c:plotArea><c:layout/>${body}${axes ? axesXml(c, plan) : ''}${dt}${shapeXml(o.plotArea)}</c:plotArea>`
    + legend + v('plotVisOnly', o.showHidden ? 0 : 1) + v('dispBlanksAs', blanks) + '</c:chart>' + shapeXml(o.chartArea) + ext + '</c:chartSpace>';
};
