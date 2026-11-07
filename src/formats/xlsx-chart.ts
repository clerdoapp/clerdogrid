// DrawingML charts (ECMA-376 Part 1, 21.2) and spreadsheet drawing anchors (20.5).
import type { Chart, ChartSeries, ChartType, Sheet } from '../model/workbook';
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
  const fill = kid(sp, 'solidFill') ?? kid(kid(sp, 'ln'), 'solidFill');
  const c = attr(kid(fill, 'srgbClr'), 'val');
  return c && /^[0-9A-Fa-f]{6}$/.test(c) ? '#' + c.toUpperCase() : undefined;
};

const readSeries = (ser: XNode, sub?: 'column' | 'line'): ChartSeries => {
  const s: ChartSeries = { values: seriesRef(kid(ser, 'val') ?? kid(ser, 'yVal')) ?? '' };
  const tx = kid(ser, 'tx');
  if (tx) { const f = find(tx, 'f'); const v = find(tx, 'v'); const name = f?.text ?? v?.text; if (name) s.name = name; }
  const cat = seriesRef(kid(ser, 'cat') ?? kid(ser, 'xVal'));
  if (cat) s.categories = cat;
  const color = seriesColor(ser);
  if (color) s.color = color;
  if (sub) s.type = sub;
  return s;
};

const LEGEND: Record<string, Chart['legend']> = { r: 'right', b: 'bottom', t: 'top', l: 'left', tr: 'right' };

export const readChart = (root: XNode, id: string, anchor: Chart['anchor']): Chart | undefined => {
  const chart = kid(root, 'chart');
  const plot = kid(chart, 'plotArea');
  if (!plot) return undefined;
  const groups = plot.children.filter((c) => /Chart$/.test(c.name));
  if (!groups.length) return undefined;
  const typeOf = (g: XNode): ChartType | undefined => {
    const grouping = attr(kid(g, 'grouping'), 'val') ?? 'standard';
    switch (g.name) {
      case 'barChart': case 'bar3DChart': {
        const bar = attr(kid(g, 'barDir'), 'val') === 'bar';
        if (grouping === 'stacked') return bar ? 'stackedBar' : 'stackedColumn';
        if (grouping === 'percentStacked') return bar ? 'stackedBar' : 'percentColumn';
        return bar ? 'bar' : 'column';
      }
      case 'lineChart': case 'line3DChart': case 'stockChart': return 'line';
      case 'areaChart': case 'area3DChart': return grouping === 'stacked' || grouping === 'percentStacked' ? 'stackedArea' : 'area';
      case 'pieChart': case 'pie3DChart': case 'ofPieChart': return 'pie';
      case 'doughnutChart': return 'doughnut';
      case 'scatterChart': case 'bubbleChart': return 'scatter';
      case 'radarChart': return 'radar';
      default: return undefined;
    }
  };
  const types = groups.map(typeOf);
  const known = groups.filter((_, i) => types[i]);
  if (!known.length) return undefined;
  const combo = known.length > 1;
  const series: ChartSeries[] = [];
  for (const g of known) {
    const t = typeOf(g)!;
    const sub = combo ? (t === 'line' ? 'line' : 'column') : undefined;
    for (const ser of kids(g, 'ser')) series.push(readSeries(ser, sub));
  }
  const c: Chart = { id, type: combo ? 'combo' : types.find(Boolean)!, series, anchor };
  const title = richText(kid(chart, 'title'));
  if (title) c.title = title;
  const legend = kid(chart, 'legend');
  c.legend = legend ? LEGEND[attr(kid(legend, 'legendPos'), 'val') ?? 'r'] ?? 'right' : 'none';
  const axTitle = (ax: XNode | undefined): string | undefined => richText(kid(ax, 'title'));
  const catAx = kid(plot, 'catAx') ?? kid(plot, 'dateAx') ?? kids(plot, 'valAx').find((a) => /^[bt]$/.test(attr(kid(a, 'axPos'), 'val') ?? ''));
  const valAx = kids(plot, 'valAx').find((a) => a !== catAx);
  const x = axTitle(catAx), y = axTitle(valAx);
  if (x || y) c.axisTitles = { ...(x ? { x } : {}), ...(y ? { y } : {}) };
  const dl = known.some((g) => attr(kid(kid(g, 'dLbls'), 'showVal'), 'val') === '1' || kids(g, 'ser').some((s) => attr(kid(kid(s, 'dLbls'), 'showVal'), 'val') === '1'));
  if (dl) c.dataLabels = true;
  return c;
};

// ---------------------------------------------------------------- chart writing

const REF_RE = /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-￿]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

const txRich = (text: string): string =>
  `<c:tx><c:rich><a:bodyPr/><a:lstStyle/>${text.split('\n').map((l) => `<a:p><a:r><a:t>${esc(l)}</a:t></a:r></a:p>`).join('')}</c:rich></c:tx>`;
const titleXml = (text: string | undefined): string => (text ? `<c:title>${txRich(text)}<c:overlay val="0"/></c:title>` : '');
const v = (tag: string, val: string | number): string => `<c:${tag} val="${val}"/>`;

const dLblsXml = (show: boolean): string =>
  `<c:dLbls>${v('showLegendKey', 0)}${v('showVal', show ? 1 : 0)}${v('showCatName', 0)}${v('showSerName', 0)}${v('showPercent', 0)}${v('showBubbleSize', 0)}</c:dLbls>`;

const refXml = (ref: string | undefined, num: boolean): string => {
  if (!ref) return '';
  return num ? `<c:numRef><c:f>${esc(ref)}</c:f></c:numRef>` : `<c:strRef><c:f>${esc(ref)}</c:f></c:strRef>`;
};

const serXml = (s: ChartSeries, i: number, kind: 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'radar'): string => {
  let x = v('idx', i) + v('order', i);
  if (s.name) x += REF_RE.test(s.name) ? `<c:tx>${refXml(s.name, false)}</c:tx>` : `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
  // Without a colour of its own, a series takes the theme's automatic colour.
  const fill = s.color ? `<a:solidFill><a:srgbClr val="${s.color.replace('#', '').slice(0, 6).toUpperCase()}"/></a:solidFill>` : '';
  if (kind === 'line' || kind === 'radar') x += `${fill ? `<c:spPr><a:ln w="28575" cap="rnd">${fill}<a:round/></a:ln></c:spPr>` : ''}<c:marker>${v('symbol', 'none')}</c:marker>`;
  else if (kind === 'scatter') x += `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker>${v('symbol', 'circle')}${v('size', 5)}${fill ? `<c:spPr>${fill}</c:spPr>` : ''}</c:marker>`;
  else if (kind !== 'pie' && fill) x += `<c:spPr>${fill}</c:spPr>`;
  if (kind === 'bar') x += v('invertIfNegative', 0);
  if (kind === 'scatter') x += `${s.categories ? `<c:xVal>${refXml(s.categories, true)}</c:xVal>` : ''}<c:yVal>${refXml(s.values, true)}</c:yVal>${v('smooth', 0)}`;
  else {
    x += s.categories ? `<c:cat>${refXml(s.categories, false)}</c:cat>` : '';
    x += `<c:val>${refXml(s.values, true)}</c:val>`;
    if (kind === 'line') x += v('smooth', 0);
  }
  return `<c:ser>${x}</c:ser>`;
};

const AX_CAT = 500000001, AX_VAL = 500000002;

const axesXml = (c: Chart, horizontal: boolean, scatter: boolean): string => {
  const catPos = horizontal ? 'l' : 'b', valPos = horizontal ? 'b' : 'l';
  const common = (id: number, cross: number, pos: string, title: string | undefined, grid: boolean): string =>
    v('axId', id) + `<c:scaling>${v('orientation', 'minMax')}</c:scaling>` + v('delete', 0) + v('axPos', pos)
    + (grid ? '<c:majorGridlines/>' : '') + titleXml(title) + '<c:numFmt formatCode="General" sourceLinked="1"/>'
    + v('majorTickMark', 'out') + v('minorTickMark', 'none') + v('tickLblPos', 'nextTo') + v('crossAx', cross) + v('crosses', 'autoZero');
  const cat = scatter
    ? `<c:valAx>${common(AX_CAT, AX_VAL, catPos, c.axisTitles?.x, false)}${v('crossBetween', 'midCat')}</c:valAx>`
    : `<c:catAx>${common(AX_CAT, AX_VAL, catPos, c.axisTitles?.x, false)}${v('auto', 1)}${v('lblAlgn', 'ctr')}${v('lblOffset', 100)}${v('noMultiLvlLbl', 0)}</c:catAx>`;
  const val = `<c:valAx>${common(AX_VAL, AX_CAT, valPos, c.axisTitles?.y, true)}${v('crossBetween', scatter ? 'midCat' : 'between')}</c:valAx>`;
  return cat + val;
};

export const chartXml = (c: Chart): string => {
  const dl = dLblsXml(!!c.dataLabels);
  const axIds = v('axId', AX_CAT) + v('axId', AX_VAL);
  const bar = (dir: 'col' | 'bar', grouping: string, list: ChartSeries[], offset = 0): string =>
    `<c:barChart>${v('barDir', dir)}${v('grouping', grouping)}${v('varyColors', 0)}${list.map((s, i) => serXml(s, i + offset, 'bar')).join('')}${dl}`
    + `${v('gapWidth', 150)}${grouping !== 'clustered' ? v('overlap', 100) : ''}${axIds}</c:barChart>`;
  const line = (list: ChartSeries[], offset = 0): string =>
    `<c:lineChart>${v('grouping', 'standard')}${v('varyColors', 0)}${list.map((s, i) => serXml(s, i + offset, 'line')).join('')}${dl}${v('marker', 1)}${axIds}</c:lineChart>`;
  let body = '';
  let axes = true, horizontal = false, scatter = false;
  switch (c.type) {
    case 'column': body = bar('col', 'clustered', c.series); break;
    case 'bar': body = bar('bar', 'clustered', c.series); horizontal = true; break;
    case 'stackedColumn': body = bar('col', 'stacked', c.series); break;
    case 'stackedBar': body = bar('bar', 'stacked', c.series); horizontal = true; break;
    case 'percentColumn': body = bar('col', 'percentStacked', c.series); break;
    case 'line': body = line(c.series); break;
    case 'area': case 'stackedArea':
      body = `<c:areaChart>${v('grouping', c.type === 'area' ? 'standard' : 'stacked')}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'area')).join('')}${dl}${axIds}</c:areaChart>`;
      break;
    case 'pie': case 'doughnut': {
      axes = false;
      const sers = c.series.map((s, i) => serXml(s, i, 'pie')).join('');
      body = c.type === 'pie'
        ? `<c:pieChart>${v('varyColors', 1)}${sers}${dl}${v('firstSliceAng', 0)}</c:pieChart>`
        : `<c:doughnutChart>${v('varyColors', 1)}${sers}${dl}${v('firstSliceAng', 0)}${v('holeSize', 50)}</c:doughnutChart>`;
      break;
    }
    case 'scatter':
      scatter = true;
      body = `<c:scatterChart>${v('scatterStyle', 'lineMarker')}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'scatter')).join('')}${dl}${axIds}</c:scatterChart>`;
      break;
    case 'radar':
      body = `<c:radarChart>${v('radarStyle', 'marker')}${v('varyColors', 0)}${c.series.map((s, i) => serXml(s, i, 'radar')).join('')}${dl}${axIds}</c:radarChart>`;
      break;
    case 'combo': {
      const cols = c.series.filter((s) => s.type !== 'line');
      const lines = c.series.filter((s) => s.type === 'line');
      body = (cols.length ? bar('col', 'clustered', cols) : '') + (lines.length ? line(lines, cols.length) : '');
      break;
    }
  }
  const legend = c.legend === 'none' ? '' : `<c:legend>${v('legendPos', { right: 'r', bottom: 'b', top: 't', left: 'l' }[c.legend ?? 'right'])}<c:overlay val="0"/></c:legend>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + v('roundedCorners', 0) + '<c:chart>' + (c.title ? titleXml(c.title) + v('autoTitleDeleted', 0) : v('autoTitleDeleted', 1))
    + `<c:plotArea><c:layout/>${body}${axes ? axesXml(c, horizontal, scatter) : ''}</c:plotArea>`
    + legend + v('plotVisOnly', 1) + v('dispBlanksAs', 'gap') + '</c:chart></c:chartSpace>';
};

