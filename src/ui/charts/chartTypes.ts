import type { ChartType } from '../../model/workbook';

export type ChartGroup = 'Column' | 'Bar' | 'Line' | 'Area' | 'Pie' | 'Scatter' | 'Statistical' | 'Combo' | 'Hierarchy' | 'Other';
export const CHART_GROUPS: ChartGroup[] = ['Column', 'Bar', 'Line', 'Area', 'Pie', 'Scatter', 'Statistical', 'Combo', 'Hierarchy', 'Other'];

export const CHART_TYPES: { type: ChartType; label: string; group: ChartGroup }[] = [
  { type: 'column', label: 'Column', group: 'Column' },
  { type: 'stackedColumn', label: 'Stacked column', group: 'Column' },
  { type: 'percentColumn', label: '100% column', group: 'Column' },
  { type: 'bar', label: 'Bar', group: 'Bar' },
  { type: 'stackedBar', label: 'Stacked bar', group: 'Bar' },
  { type: 'percentBar', label: '100% bar', group: 'Bar' },
  { type: 'line', label: 'Line', group: 'Line' },
  { type: 'stackedLine', label: 'Stacked line', group: 'Line' },
  { type: 'percentLine', label: '100% line', group: 'Line' },
  { type: 'lineMarkers', label: 'Line with points', group: 'Line' },
  { type: 'stackedLineMarkers', label: 'Stacked line with points', group: 'Line' },
  { type: 'percentLineMarkers', label: '100% line with points', group: 'Line' },
  { type: 'area', label: 'Area', group: 'Area' },
  { type: 'stackedArea', label: 'Stacked area', group: 'Area' },
  { type: 'percentArea', label: '100% area', group: 'Area' },
  { type: 'pie', label: 'Pie', group: 'Pie' },
  { type: 'doughnut', label: 'Doughnut', group: 'Pie' },
  { type: 'scatter', label: 'Scatter', group: 'Scatter' },
  { type: 'scatterSmoothMarkers', label: 'Smooth lines with points', group: 'Scatter' },
  { type: 'scatterSmooth', label: 'Smooth lines', group: 'Scatter' },
  { type: 'scatterLinesMarkers', label: 'Straight lines with points', group: 'Scatter' },
  { type: 'scatterLines', label: 'Straight lines', group: 'Scatter' },
  { type: 'bubble', label: 'Bubble', group: 'Scatter' },
  { type: 'histogram', label: 'Histogram', group: 'Statistical' },
  { type: 'pareto', label: 'Pareto', group: 'Statistical' },
  { type: 'boxWhisker', label: 'Box and whisker', group: 'Statistical' },
  { type: 'combo', label: 'Column and line', group: 'Combo' },
  { type: 'comboSecondary', label: 'Column and line, two axes', group: 'Combo' },
  { type: 'comboAreaColumn', label: 'Stacked area and column', group: 'Combo' },
  { type: 'treemap', label: 'Treemap', group: 'Hierarchy' },
  { type: 'sunburst', label: 'Sunburst', group: 'Hierarchy' },
  { type: 'waterfall', label: 'Waterfall', group: 'Other' },
  { type: 'funnel', label: 'Funnel', group: 'Other' },
  { type: 'radar', label: 'Radar', group: 'Other' },
  { type: 'radarMarkers', label: 'Radar with points', group: 'Other' },
  { type: 'radarFilled', label: 'Filled radar', group: 'Other' },
];

export const ALL_CHART_TYPES: ChartType[] = CHART_TYPES.map((t) => t.type);
export const isChartType = (t: unknown): t is ChartType => typeof t === 'string' && (ALL_CHART_TYPES as string[]).includes(t);
export const chartLabel = (t: ChartType): string => CHART_TYPES.find((x) => x.type === t)?.label ?? t;
export const chartGroup = (t: ChartType): ChartGroup => CHART_TYPES.find((x) => x.type === t)?.group ?? 'Column';

export type Family = 'cartesian' | 'pie' | 'scatter' | 'bubble' | 'radar' | 'histogram' | 'pareto' | 'box' | 'treemap' | 'sunburst' | 'waterfall' | 'funnel';
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
}

const T = (family: Family, extra: Partial<Traits> = {}): Traits => ({
  family, base: 'column', horizontal: false, stack: 'none', markers: false, smooth: false, lines: false, combo: false, filled: false, ...extra,
});

const TRAITS: Record<ChartType, Traits> = {
  column: T('cartesian'), stackedColumn: T('cartesian', { stack: 'stacked' }), percentColumn: T('cartesian', { stack: 'percent' }),
  bar: T('cartesian', { horizontal: true }), stackedBar: T('cartesian', { horizontal: true, stack: 'stacked' }), percentBar: T('cartesian', { horizontal: true, stack: 'percent' }),
  line: T('cartesian', { base: 'line' }), stackedLine: T('cartesian', { base: 'line', stack: 'stacked' }), percentLine: T('cartesian', { base: 'line', stack: 'percent' }),
  lineMarkers: T('cartesian', { base: 'line', markers: true }), stackedLineMarkers: T('cartesian', { base: 'line', stack: 'stacked', markers: true }),
  percentLineMarkers: T('cartesian', { base: 'line', stack: 'percent', markers: true }),
  area: T('cartesian', { base: 'area' }), stackedArea: T('cartesian', { base: 'area', stack: 'stacked' }), percentArea: T('cartesian', { base: 'area', stack: 'percent' }),
  pie: T('pie'), doughnut: T('pie', { filled: true }),
  scatter: T('scatter', { markers: true }), scatterSmoothMarkers: T('scatter', { markers: true, smooth: true, lines: true }),
  scatterSmooth: T('scatter', { smooth: true, lines: true }), scatterLinesMarkers: T('scatter', { markers: true, lines: true }), scatterLines: T('scatter', { lines: true }),
  bubble: T('bubble'),
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
  return tr.family === 'cartesian' && tr.stack === 'none' && !tr.horizontal;
};

/** Types that show their categories as slices rather than as axis labels. */
export const usesCategoryColors = (t: ChartType): boolean => ['pie', 'funnel', 'treemap', 'sunburst'].includes(traits(t).family);
export const hasAxes = (t: ChartType): boolean => !['pie', 'treemap', 'sunburst', 'funnel', 'radar'].includes(traits(t).family);

/** Short list for the tool panel's one-click buttons. */
export const QUICK_CHARTS: [ChartType, string][] = [
  ['column', 'Column'], ['bar', 'Bar'], ['lineMarkers', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'],
  ['scatter', 'Scatter'], ['stackedColumn', 'Stacked'], ['combo', 'Combo'], ['histogram', 'Histogram'], ['waterfall', 'Waterfall'], ['treemap', 'Treemap'],
];
