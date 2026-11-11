import type { ChartType } from '../../model/workbook';
import { N, tr as translate } from '../i18n';

export type ChartGroup = 'Column' | 'Bar' | 'Line' | 'Area' | 'Pie' | 'Scatter' | 'Stock' | 'Surface' | '3D' | 'Map' | 'Statistical' | 'Combo' | 'Hierarchy' | 'Other';
export const CHART_GROUPS: ChartGroup[] = ['Column', 'Bar', 'Line', 'Area', 'Pie', 'Scatter', 'Stock', 'Surface', '3D', 'Map', 'Statistical', 'Combo', 'Hierarchy', 'Other'];
/** The group names for the interface (translated where shown). */
export const CHART_GROUP_LABELS: string[] = [N('Column'), N('Bar'), N('Line'), N('Area'), N('Pie'), N('Scatter'), N('Stock'), N('Surface'), N('3D'), N('Map'), N('Statistical'), N('Combo'), N('Hierarchy'), N('Other')];

export const CHART_TYPES: { type: ChartType; label: string; group: ChartGroup }[] = [
  { type: 'column', label: N('Column'), group: 'Column' },
  { type: 'stackedColumn', label: N('Stacked column'), group: 'Column' },
  { type: 'percentColumn', label: N('100% column'), group: 'Column' },
  { type: 'bar', label: N('Bar'), group: 'Bar' },
  { type: 'stackedBar', label: N('Stacked bar'), group: 'Bar' },
  { type: 'percentBar', label: N('100% bar'), group: 'Bar' },
  { type: 'line', label: N('Line'), group: 'Line' },
  { type: 'stackedLine', label: N('Stacked line'), group: 'Line' },
  { type: 'percentLine', label: N('100% line'), group: 'Line' },
  { type: 'lineMarkers', label: N('Line with points'), group: 'Line' },
  { type: 'stackedLineMarkers', label: N('Stacked line with points'), group: 'Line' },
  { type: 'percentLineMarkers', label: N('100% line with points'), group: 'Line' },
  { type: 'area', label: N('Area'), group: 'Area' },
  { type: 'stackedArea', label: N('Stacked area'), group: 'Area' },
  { type: 'percentArea', label: N('100% area'), group: 'Area' },
  { type: 'pie', label: N('Pie'), group: 'Pie' },
  { type: 'pieExploded', label: N('Exploded pie'), group: 'Pie' },
  { type: 'doughnut', label: N('Doughnut'), group: 'Pie' },
  { type: 'doughnutExploded', label: N('Exploded doughnut'), group: 'Pie' },
  { type: 'pieOfPie', label: N('Pie of pie'), group: 'Pie' },
  { type: 'barOfPie', label: N('Bar of pie'), group: 'Pie' },
  { type: 'scatter', label: N('Scatter'), group: 'Scatter' },
  { type: 'scatterSmoothMarkers', label: N('Smooth lines with points'), group: 'Scatter' },
  { type: 'scatterSmooth', label: N('Smooth lines'), group: 'Scatter' },
  { type: 'scatterLinesMarkers', label: N('Straight lines with points'), group: 'Scatter' },
  { type: 'scatterLines', label: N('Straight lines'), group: 'Scatter' },
  { type: 'bubble', label: N('Bubble'), group: 'Scatter' },
  { type: 'bubble3D', label: N('Bubble with depth'), group: 'Scatter' },
  { type: 'stockHLC', label: N('High-low-close'), group: 'Stock' },
  { type: 'stockOHLC', label: N('Open-high-low-close'), group: 'Stock' },
  { type: 'stockVHLC', label: N('Volume-high-low-close'), group: 'Stock' },
  { type: 'stockVOHLC', label: N('Volume-open-high-low-close'), group: 'Stock' },
  { type: 'surface3D', label: N('3D surface'), group: 'Surface' },
  { type: 'surfaceWireframe3D', label: N('Wireframe 3D surface'), group: 'Surface' },
  { type: 'contour', label: N('Contour'), group: 'Surface' },
  { type: 'contourWireframe', label: N('Wireframe contour'), group: 'Surface' },
  { type: 'column3DClustered', label: N('3D clustered column'), group: '3D' },
  { type: 'column3DStacked', label: N('3D stacked column'), group: '3D' },
  { type: 'column3DPercent', label: N('3D 100% column'), group: '3D' },
  { type: 'column3D', label: N('3D column'), group: '3D' },
  { type: 'cylinderColumn', label: N('Cylinder column'), group: '3D' },
  { type: 'coneColumn', label: N('Cone column'), group: '3D' },
  { type: 'pyramidColumn', label: N('Pyramid column'), group: '3D' },
  { type: 'bar3DClustered', label: N('3D clustered bar'), group: '3D' },
  { type: 'bar3DStacked', label: N('3D stacked bar'), group: '3D' },
  { type: 'bar3DPercent', label: N('3D 100% bar'), group: '3D' },
  { type: 'line3D', label: N('3D line'), group: '3D' },
  { type: 'area3D', label: N('3D area'), group: '3D' },
  { type: 'stackedArea3D', label: N('3D stacked area'), group: '3D' },
  { type: 'percentArea3D', label: N('3D 100% area'), group: '3D' },
  { type: 'pie3D', label: N('3D pie'), group: '3D' },
  { type: 'pie3DExploded', label: N('Exploded 3D pie'), group: '3D' },
  { type: 'regionMap', label: N('Filled map'), group: 'Map' },
  { type: 'histogram', label: N('Histogram'), group: 'Statistical' },
  { type: 'pareto', label: N('Pareto'), group: 'Statistical' },
  { type: 'boxWhisker', label: N('Box and whisker'), group: 'Statistical' },
  { type: 'combo', label: N('Column and line'), group: 'Combo' },
  { type: 'comboSecondary', label: N('Column and line, two axes'), group: 'Combo' },
  { type: 'comboAreaColumn', label: N('Stacked area and column'), group: 'Combo' },
  { type: 'treemap', label: N('Treemap'), group: 'Hierarchy' },
  { type: 'sunburst', label: N('Sunburst'), group: 'Hierarchy' },
  { type: 'waterfall', label: N('Waterfall'), group: 'Other' },
  { type: 'funnel', label: N('Funnel'), group: 'Other' },
  { type: 'radar', label: N('Radar'), group: 'Other' },
  { type: 'radarMarkers', label: N('Radar with points'), group: 'Other' },
  { type: 'radarFilled', label: N('Filled radar'), group: 'Other' },
];

export const ALL_CHART_TYPES: ChartType[] = CHART_TYPES.map((t) => t.type);
export const isChartType = (t: unknown): t is ChartType => typeof t === 'string' && (ALL_CHART_TYPES as string[]).includes(t);
export const chartLabel = (t: ChartType): string => translate(CHART_TYPES.find((x) => x.type === t)?.label ?? t);
export const chartGroup = (t: ChartType): ChartGroup => CHART_TYPES.find((x) => x.type === t)?.group ?? 'Column';

export type Family = 'cartesian' | 'pie' | 'scatter' | 'bubble' | 'radar' | 'histogram' | 'pareto' | 'box' | 'treemap' | 'sunburst' | 'waterfall' | 'funnel'
  | 'stock' | 'surface' | 'bar3d' | 'line3d' | 'area3d' | 'ofPie' | 'map';
export type Stack = 'none' | 'stacked' | 'percent';
export type SeriesKind = 'column' | 'line' | 'area';

export interface Traits {
  family: Family;
  /** Kind of series for plain category charts. */
  base: SeriesKind;
  horizontal: boolean;
  stack: Stack;
  markers: boolean;
  smooth: boolean;
  /** Scatter: draw lines between points. */
  lines: boolean;
  combo: boolean;
  filled: boolean;
  /** Drawn with depth. */
  d3: boolean;
  /** Pie slices pulled apart. */
  exploded: boolean;
  /** 3D column with one row per series (not clustered). */
  rows3d: boolean;
  shape: 'box' | 'cylinder' | 'cone' | 'pyramid';
  /** Surface: wireframe, and seen from the top (contour). */
  wire: boolean;
  contour: boolean;
  /** Stock: first series is volume / has an open series. */
  volume: boolean;
  open: boolean;
  /** Pie of pie: the second plot is a bar. */
  ofBar: boolean;
}

const T = (family: Family, extra: Partial<Traits> = {}): Traits => ({
  family, base: 'column', horizontal: false, stack: 'none', markers: false, smooth: false, lines: false, combo: false, filled: false,
  d3: false, exploded: false, rows3d: false, shape: 'box', wire: false, contour: false, volume: false, open: false, ofBar: false, ...extra,
});

const TRAITS: Record<ChartType, Traits> = {
  column: T('cartesian'), stackedColumn: T('cartesian', { stack: 'stacked' }), percentColumn: T('cartesian', { stack: 'percent' }),
  bar: T('cartesian', { horizontal: true }), stackedBar: T('cartesian', { horizontal: true, stack: 'stacked' }), percentBar: T('cartesian', { horizontal: true, stack: 'percent' }),
  line: T('cartesian', { base: 'line' }), stackedLine: T('cartesian', { base: 'line', stack: 'stacked' }), percentLine: T('cartesian', { base: 'line', stack: 'percent' }),
  lineMarkers: T('cartesian', { base: 'line', markers: true }), stackedLineMarkers: T('cartesian', { base: 'line', stack: 'stacked', markers: true }),
  percentLineMarkers: T('cartesian', { base: 'line', stack: 'percent', markers: true }),
  area: T('cartesian', { base: 'area' }), stackedArea: T('cartesian', { base: 'area', stack: 'stacked' }), percentArea: T('cartesian', { base: 'area', stack: 'percent' }),
  pie: T('pie'), doughnut: T('pie', { filled: true }), pieExploded: T('pie', { exploded: true }), doughnutExploded: T('pie', { filled: true, exploded: true }),
  pie3D: T('pie', { d3: true }), pie3DExploded: T('pie', { d3: true, exploded: true }),
  pieOfPie: T('ofPie'), barOfPie: T('ofPie', { ofBar: true }),
  scatter: T('scatter', { markers: true }), scatterSmoothMarkers: T('scatter', { markers: true, smooth: true, lines: true }),
  scatterSmooth: T('scatter', { smooth: true, lines: true }), scatterLinesMarkers: T('scatter', { markers: true, lines: true }), scatterLines: T('scatter', { lines: true }),
  bubble: T('bubble'), bubble3D: T('bubble', { d3: true }),
  stockHLC: T('stock', { base: 'line' }), stockOHLC: T('stock', { base: 'line', open: true }),
  stockVHLC: T('stock', { base: 'line', volume: true }), stockVOHLC: T('stock', { base: 'line', volume: true, open: true }),
  surface3D: T('surface', { d3: true }), surfaceWireframe3D: T('surface', { d3: true, wire: true }),
  contour: T('surface', { contour: true }), contourWireframe: T('surface', { contour: true, wire: true }),
  column3DClustered: T('bar3d', { d3: true }), column3DStacked: T('bar3d', { d3: true, stack: 'stacked' }), column3DPercent: T('bar3d', { d3: true, stack: 'percent' }),
  column3D: T('bar3d', { d3: true, rows3d: true }),
  cylinderColumn: T('bar3d', { d3: true, shape: 'cylinder' }), coneColumn: T('bar3d', { d3: true, shape: 'cone' }), pyramidColumn: T('bar3d', { d3: true, shape: 'pyramid' }),
  bar3DClustered: T('bar3d', { d3: true, horizontal: true }), bar3DStacked: T('bar3d', { d3: true, horizontal: true, stack: 'stacked' }),
  bar3DPercent: T('bar3d', { d3: true, horizontal: true, stack: 'percent' }),
  line3D: T('line3d', { d3: true, base: 'line', rows3d: true }),
  area3D: T('area3d', { d3: true, base: 'area', rows3d: true }), stackedArea3D: T('area3d', { d3: true, base: 'area', stack: 'stacked' }),
  percentArea3D: T('area3d', { d3: true, base: 'area', stack: 'percent' }),
  regionMap: T('map'),
  histogram: T('histogram'), pareto: T('pareto'), boxWhisker: T('box'),
  combo: T('cartesian', { combo: true }), comboSecondary: T('cartesian', { combo: true }), comboAreaColumn: T('cartesian', { combo: true }),
  treemap: T('treemap'), sunburst: T('sunburst'),
  waterfall: T('waterfall'), funnel: T('funnel'),
  radar: T('radar', { base: 'line' }), radarMarkers: T('radar', { base: 'line', markers: true }), radarFilled: T('radar', { base: 'area', filled: true }),
};

export const traits = (t: ChartType): Traits => TRAITS[t] ?? TRAITS.column;

/** Default way each series of a combo chart is drawn. */
export const comboKind = (type: ChartType, i: number, n: number): SeriesKind => {
  if (type === 'comboAreaColumn') return i < Math.max(1, Math.ceil(n / 2)) && !(n > 1 && i === n - 1) ? 'area' : 'column';
  return i === 0 ? 'column' : 'line';
};

/** Types whose value axis can be split into two. */
export const supportsSecondary = (t: ChartType): boolean => {
  const tr = traits(t);
  return (tr.family === 'cartesian' && !tr.horizontal) || tr.family === 'scatter';
};

/** Types that show their categories as slices rather than as axis labels. */
export const usesCategoryColors = (t: ChartType): boolean => ['pie', 'ofPie', 'funnel', 'treemap', 'sunburst', 'map'].includes(traits(t).family);
export const hasAxes = (t: ChartType): boolean => !['pie', 'ofPie', 'treemap', 'sunburst', 'funnel', 'radar', 'map'].includes(traits(t).family) && !(traits(t).family === 'surface' && traits(t).contour);
/** Category charts drawn on a flat plot with a category axis (trendlines, error bars, labels, data table apply). */
export const isCategoryChart = (t: ChartType): boolean => ['cartesian', 'stock'].includes(traits(t).family);
/** Charts that can carry trendlines. */
export const supportsTrendlines = (t: ChartType): boolean => {
  const tr = traits(t);
  return (tr.family === 'cartesian' && tr.stack === 'none') || tr.family === 'scatter' || tr.family === 'bubble' || tr.family === 'stock';
};
export const supportsErrorBars = (t: ChartType): boolean => {
  const tr = traits(t);
  return tr.family === 'cartesian' || tr.family === 'scatter' || tr.family === 'bubble';
};
export const is3D = (t: ChartType): boolean => traits(t).d3 || (traits(t).family === 'surface');

/** Short list for the tool panel's one-click buttons. */
export const QUICK_CHARTS: [ChartType, string][] = [
  ['column', N('Column')], ['bar', N('Bar')], ['lineMarkers', N('Line')], ['area', N('Area')], ['pie', N('Pie')], ['doughnut', N('Doughnut')],
  ['scatter', N('Scatter')], ['stackedColumn', N('Stacked')], ['combo', N('Combo')], ['histogram', N('Histogram')], ['waterfall', N('Waterfall')], ['treemap', N('Treemap')],
  ['stockHLC', N('Stock')], ['column3DClustered', N('3D column')], ['regionMap', N('Map')],
];
