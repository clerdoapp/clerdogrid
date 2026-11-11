// Charts as OpenDocument chart objects (ODF 1.3 Part 3 §11): "Object N/content.xml" inside the package.
import type { Chart, ChartSeries, ChartType } from '../model/workbook';
import { parseFormula, type RefNode } from '../engine/parser';
import { a, child, children, elems, findDeep, odfText, tag, textXml, NS_DECL, type ONode } from './ods-xml';
import { gridRef, odfRef, splitAddresses } from './ods-formula';
import { isType, sanitizeOptions } from './xlsx-chart';

/** "Sheet1.A1:Sheet1.B5" (ODF range address) → "Sheet1!$A$1:$B$5". */
export const gridRange = (addr: string | undefined): string | undefined => {
  if (!addr) return undefined;
  const first = splitAddresses(addr)[0] ?? '';
  const r = gridRef(first.includes('.') ? first : `.${first}`);
  if (r === '#REF!') return undefined;
  return r.replace(/(^|!|:)([A-Z]+)(\d+)/g, '$1$$$2$$$3');
};

const refOf = (formula: string | undefined): RefNode | undefined => {
  if (!formula) return undefined;
  try {
    const n = parseFormula(formula.replace(/^=/, ''));
    return n.type === 'ref' ? n.ref : undefined;
  } catch { return undefined; }
};

/** ClerdoGrid range text → ODF range address ("Sheet1.A1:Sheet1.B5"). */
export const odfRange = (formula: string | undefined, sheet: string): string | undefined => {
  const r = refOf(formula);
  return r ? odfRef(r, sheet, true).replace(/\$(?=[A-Z]|\d)/g, '').replace(/^\$/, '').replace(/:\$/, ':') : undefined;
};

const CLASS_OUT: Partial<Record<ChartType, { cls: string; vertical?: boolean; stacked?: boolean; percent?: boolean; symbols?: boolean; lines?: boolean }>> = {
  column: { cls: 'bar' }, stackedColumn: { cls: 'bar', stacked: true }, percentColumn: { cls: 'bar', percent: true },
  bar: { cls: 'bar', vertical: true }, stackedBar: { cls: 'bar', vertical: true, stacked: true }, percentBar: { cls: 'bar', vertical: true, percent: true },
  line: { cls: 'line' }, stackedLine: { cls: 'line', stacked: true }, percentLine: { cls: 'line', percent: true },
  lineMarkers: { cls: 'line', symbols: true }, stackedLineMarkers: { cls: 'line', stacked: true, symbols: true }, percentLineMarkers: { cls: 'line', percent: true, symbols: true },
  area: { cls: 'area' }, stackedArea: { cls: 'area', stacked: true }, percentArea: { cls: 'area', percent: true },
  pie: { cls: 'circle' }, doughnut: { cls: 'ring' },
  scatter: { cls: 'scatter', symbols: true, lines: false }, scatterLines: { cls: 'scatter', lines: true }, scatterLinesMarkers: { cls: 'scatter', lines: true, symbols: true },
  scatterSmooth: { cls: 'scatter', lines: true }, scatterSmoothMarkers: { cls: 'scatter', lines: true, symbols: true },
  bubble: { cls: 'bubble' }, radar: { cls: 'radar' }, radarMarkers: { cls: 'radar', symbols: true }, radarFilled: { cls: 'filled-radar' },
};

/** The chart object document for a chart on `sheet`. */
export const chartObjectXml = (c: Chart, sheet: string): string => {
  const spec = CLASS_OUT[c.type] ?? { cls: 'bar' };
  const styles: string[] = [];
  styles.push(tag('style:style', { 'style:name': 'ch1', 'style:family': 'chart' }, tag('style:graphic-properties', { 'draw:stroke': 'none', 'draw:fill-color': '#ffffff' })));
  styles.push(tag('style:style', { 'style:name': 'pa1', 'style:family': 'chart' }, tag('style:chart-properties', {
    'chart:vertical': spec.vertical ?? false, 'chart:stacked': spec.stacked ?? false, 'chart:percentage': spec.percent ?? false,
    'chart:three-dimensional': false, 'chart:lines': spec.lines, 'chart:symbol-type': spec.symbols ? 'automatic' : spec.cls === 'line' || spec.cls === 'radar' ? 'none' : undefined,
    'chart:interpolation': c.type.startsWith('scatterSmooth') ? 'cubic-spline' : undefined,
  })));
  const series = c.series.map((s, i) => {
    const name = `se${i + 1}`;
    styles.push(tag('style:style', { 'style:name': name, 'style:family': 'chart' }, tag('style:chart-properties', { 'chart:data-label-number': c.dataLabels ? 'value' : undefined, 'chart:symbol-type': spec.symbols ? 'automatic' : undefined })
      + (s.color ? tag('style:graphic-properties', { 'svg:stroke-color': s.color, 'draw:fill-color': s.color }) : '')));
    const values = odfRange(s.values, sheet);
    const label = odfRange(s.name, sheet);
    const cls = c.type.startsWith('combo') && s.type ? `chart:${s.type === 'column' ? 'bar' : s.type}` : undefined;
    let inner = '';
    if ((spec.cls === 'scatter' || spec.cls === 'bubble') && s.categories) inner += tag('chart:domain', { 'table:cell-range-address': odfRange(s.categories, sheet) });
    if (spec.cls === 'bubble') {
      const sizes = odfRange(s.sizes, sheet);
      return tag('chart:series', { 'chart:style-name': name, 'chart:values-cell-range-address': sizes ?? values, 'chart:label-cell-address': label, 'clerdo:name': !label ? s.name : undefined },
        tag('chart:domain', { 'table:cell-range-address': values }) + (s.categories ? tag('chart:domain', { 'table:cell-range-address': odfRange(s.categories, sheet) }) : ''));
    }
    return tag('chart:series', { 'chart:style-name': name, 'chart:values-cell-range-address': values, 'chart:label-cell-address': label, 'chart:class': cls, 'clerdo:name': !label && s.name ? s.name : undefined }, inner);
  }).join('');
  const cats = spec.cls !== 'scatter' && spec.cls !== 'bubble' ? odfRange(c.series[0]?.categories, sheet) : undefined;
  const axes = spec.cls === 'circle' || spec.cls === 'ring' ? ''
    : tag('chart:axis', { 'chart:dimension': 'x', 'chart:name': 'primary-x' }, (c.axisTitles?.x ? tag('chart:title', {}, `<text:p>${textXml(c.axisTitles.x)}</text:p>`) : '') + (cats ? tag('chart:categories', { 'table:cell-range-address': cats }) : ''))
      + tag('chart:axis', { 'chart:dimension': 'y', 'chart:name': 'primary-y' }, (c.axisTitles?.y ? tag('chart:title', {}, `<text:p>${textXml(c.axisTitles.y)}</text:p>`) : '') + (c.options?.gridlines === false ? '' : '<chart:grid chart:class="major"/>'));
  const legendPos = ({ right: 'end', left: 'start', top: 'top', bottom: 'bottom' } as Record<string, string>)[c.legend ?? 'right'];
  const w = `${(c.anchor.width / 96 * 2.54).toFixed(3)}cm`, h = `${(c.anchor.height / 96 * 2.54).toFixed(3)}cm`;
  const body = tag('chart:chart', { 'svg:width': w, 'svg:height': h, 'chart:class': `chart:${spec.cls}`, 'chart:style-name': 'ch1' },
    (c.title ? tag('chart:title', {}, `<text:p>${textXml(c.title)}</text:p>`) : '')
    + (c.legend === 'none' ? '' : tag('chart:legend', { 'chart:legend-position': legendPos }))
    + tag('chart:plot-area', { 'chart:style-name': 'pa1', 'chart:data-source-has-labels': 'none' }, axes + series + '<chart:wall/><chart:floor/>'));
  return `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content${NS_DECL('office', 'style', 'text', 'table', 'draw', 'fo', 'xlink', 'svg', 'chart', 'number', 'clerdo')} office:version="1.3">`
    + `<office:automatic-styles>${styles.join('')}</office:automatic-styles><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
};

/** The settings ClerdoGrid keeps on the frame so its own charts come back exactly. */
export const chartExtra = (c: Chart): string => JSON.stringify({ type: c.type, options: c.options, series: c.series, legend: c.legend, dataLabels: c.dataLabels });

const LEGEND_IN: Record<string, Chart['legend']> = { end: 'right', start: 'left', top: 'top', bottom: 'bottom', 'top-end': 'right', 'bottom-end': 'right', 'top-start': 'left', 'bottom-start': 'left' };

/** Reads a chart object; returns undefined when it has no cell ranges to plot. */
export const readChartObject = (doc: ONode, id: string, anchor: Chart['anchor'], extra?: string): Chart | undefined => {
  const chart = findDeep(doc, 'chart:chart');
  if (!chart) return undefined;
  const styles = new Map<string, ONode>();
  for (const s of children(findDeep(doc, 'office:automatic-styles'), 'style:style')) styles.set(a(s, 'style:name') ?? '', s);
  const props = (name: string | undefined) => child(styles.get(name ?? ''), 'style:chart-properties');
  const cls = (a(chart, 'chart:class') ?? '').replace(/^chart:/, '');
  const plot = child(chart, 'chart:plot-area');
  const pp = props(a(plot, 'chart:style-name'));
  const stacked = a(pp, 'chart:stacked') === 'true', percent = a(pp, 'chart:percentage') === 'true', vertical = a(pp, 'chart:vertical') === 'true';
  const sym = (s?: ONode) => { const t = a(s, 'chart:symbol-type'); return t !== undefined && t !== 'none'; };
  const seriesNodes = children(plot, 'chart:series');
  const markers = sym(pp) || seriesNodes.some((s) => sym(props(a(s, 'chart:style-name'))));
  let type: ChartType;
  const pre = stacked ? 'stacked' : percent ? 'percent' : '';
  switch (cls) {
    case 'bar': type = (vertical ? (pre ? `${pre}Bar` : 'bar') : pre ? `${pre}Column` : 'column') as ChartType; break;
    case 'line': type = (pre ? `${pre}${markers ? 'LineMarkers' : 'Line'}` : markers ? 'lineMarkers' : 'line') as ChartType; break;
    case 'area': type = (pre ? `${pre}Area` : 'area') as ChartType; break;
    case 'circle': type = 'pie'; break;
    case 'ring': type = 'doughnut'; break;
    case 'scatter': type = a(pp, 'chart:lines') === 'true' ? (markers ? 'scatterLinesMarkers' : 'scatterLines') : 'scatter'; break;
    case 'bubble': type = 'bubble'; break;
    case 'radar': type = markers ? 'radarMarkers' : 'radar'; break;
    case 'filled-radar': type = 'radarFilled'; break;
    default: type = 'column';
  }
  const catAxis = children(plot, 'chart:axis').find((x) => a(x, 'chart:dimension') === 'x');
  const cats = gridRange(a(child(catAxis, 'chart:categories'), 'table:cell-range-address'));
  const series: ChartSeries[] = [];
  for (const s of seriesNodes) {
    const domains = children(s, 'chart:domain').map((d) => gridRange(a(d, 'table:cell-range-address')));
    let values = gridRange(a(s, 'chart:values-cell-range-address'));
    const out: ChartSeries = { values: values ?? '' };
    if (type === 'bubble') { out.sizes = values; values = domains[0]; out.values = values ?? ''; if (domains[1]) out.categories = domains[1]; }
    else if (type.startsWith('scatter')) { if (domains[0]) out.categories = domains[0]; }
    else if (cats) out.categories = cats;
    const name = gridRange(a(s, 'chart:label-cell-address')) ?? a(s, 'clerdo:name');
    if (name) out.name = name;
    const sc = a(s, 'chart:class')?.replace(/^chart:/, '');
    if (sc && sc !== cls && (sc === 'line' || sc === 'area' || sc === 'bar')) out.type = sc === 'bar' ? 'column' : sc;
    const g = child(styles.get(a(s, 'chart:style-name') ?? ''), 'style:graphic-properties');
    const color = a(g, 'draw:fill-color') ?? a(g, 'svg:stroke-color');
    if (color && /^#[0-9a-f]{6}$/i.test(color)) out.color = color.toUpperCase();
    if (out.values) series.push(out);
  }
  if (!series.length) return undefined;
  if (series.some((s) => s.type)) type = 'combo';
  const c: Chart = { id, type, series, anchor };
  const title = child(chart, 'chart:title');
  if (title) c.title = children(title, 'text:p').map(odfText).join('\n');
  const legend = child(chart, 'chart:legend');
  c.legend = legend ? LEGEND_IN[a(legend, 'chart:legend-position') ?? 'end'] ?? 'right' : 'none';
  const yAxis = children(plot, 'chart:axis').find((x) => a(x, 'chart:dimension') === 'y');
  const xt = child(catAxis, 'chart:title'), yt = child(yAxis, 'chart:title');
  if (xt || yt) c.axisTitles = { ...(xt ? { x: children(xt, 'text:p').map(odfText).join(' ') } : {}), ...(yt ? { y: children(yt, 'text:p').map(odfText).join(' ') } : {}) };
  if (seriesNodes.some((s) => { const v = a(props(a(s, 'chart:style-name')), 'chart:data-label-number'); return v !== undefined && v !== 'none'; })) c.dataLabels = true;
  if (yAxis && !child(yAxis, 'chart:grid')) c.options = { ...c.options, gridlines: false };
  // ClerdoGrid's own settings, used only while the chart still plots the same ranges
  if (extra) {
    try {
      const x = JSON.parse(extra) as Record<string, unknown>;
      const sameRanges = Array.isArray(x.series) && (x.series as ChartSeries[]).length === series.length && (x.series as ChartSeries[]).every((s, i) => typeof s?.values === 'string' && s.values.replace(/\$/g, '') === series[i]!.values.replace(/\$/g, ''));
      if (sameRanges) {
        if (isType(x.type)) c.type = x.type;
        const o = sanitizeOptions(x.options);
        if (o) c.options = o;
        (x.series as ChartSeries[]).forEach((s, i) => {
          const t = series[i]!;
          if (typeof s.name === 'string') t.name = s.name;
          if (typeof s.color === 'string' && /^#[0-9a-f]{6}$/i.test(s.color)) t.color = s.color;
          if (s.type === 'column' || s.type === 'line' || s.type === 'area') t.type = s.type;
          if (typeof s.sizes === 'string') t.sizes = s.sizes;
          if (typeof s.categories === 'string') t.categories = s.categories;
          if (typeof s.values === 'string') t.values = s.values;
        });
        if (x.legend === 'none' || x.legend === 'right' || x.legend === 'left' || x.legend === 'top' || x.legend === 'bottom') c.legend = x.legend;
        if (typeof x.dataLabels === 'boolean') c.dataLabels = x.dataLabels;
      }
    } catch { /* ignore */ }
  }
  void elems;
  return c;
};
