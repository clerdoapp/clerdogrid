import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { cellKey, newWorkbook, type Chart, type ChartType } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { chartXml, readChart, savedAs, sanitizeOptions } from '../src/formats/xlsx-chart';
import { parseRoot } from '../src/formats/xml';
import { buildSeries, detectHeaders, niceScale, resolveChart, secondaryOf } from '../src/ui/charts/chartData';
import { boxStats, buildTree, guessTotals, histogramBins, linePath, paretoRows, squarify, sunburstArcs, treeDepth, waterfallSteps } from '../src/ui/charts/chartStats';
import { ALL_CHART_TYPES, CHART_TYPES, traits } from '../src/ui/charts/chartTypes';

beforeAll(() => registerAllFunctions());

describe('axis ticks', () => {
  it('covers the range with round steps', () => {
    const s = niceScale(3, 97, 6);
    expect(s.min).toBe(0);
    expect(s.max).toBe(100);
    expect(s.step).toBe(20);
    expect(s.ticks).toEqual([0, 20, 40, 60, 80, 100]);
  });
  it('handles negatives, tiny ranges and a single value', () => {
    const neg = niceScale(-42, 17, 6);
    expect(neg.min).toBeLessThanOrEqual(-42);
    expect(neg.max).toBeGreaterThanOrEqual(17);
    expect(neg.ticks).toContain(0);
    const tiny = niceScale(0.001, 0.0042, 5);
    expect(tiny.ticks.every((t) => Number.isFinite(t))).toBe(true);
    expect(tiny.ticks[1]! - tiny.ticks[0]!).toBeCloseTo(tiny.step);
    const one = niceScale(5, 5);
    expect(one.min).toBeLessThan(5);
    expect(one.max).toBeGreaterThan(5);
  });
});

describe('series building', () => {
  const m = [[null, 'North', 'South'], ['Jan', 1, 2], ['Feb', 3, 4], ['Mar', 5, 6]];
  const r = { r1: 0, c1: 0, r2: 3, c2: 2 };
  it('one series per column with categories from the first column', () => {
    const s = buildSeries('Data', r, m);
    expect(s).toEqual([
      { name: 'Data!$B$1', values: 'Data!$B$2:$B$4', categories: 'Data!$A$2:$A$4' },
      { name: 'Data!$C$1', values: 'Data!$C$2:$C$4', categories: 'Data!$A$2:$A$4' },
    ]);
  });
  it('rows when asked', () => {
    const s = buildSeries('Data', r, m, { orientation: 'rows' });
    expect(s).toHaveLength(3);
    expect(s[0]).toEqual({ name: 'Data!$A$2', values: 'Data!$B$2:$C$2', categories: 'Data!$B$1:$C$1' });
  });
  it('combos mark their series kinds', () => {
    expect(buildSeries('Data', r, m, { type: 'combo' }).map((s) => s.type)).toEqual(['column', 'line']);
    expect(buildSeries('Data', r, m, { type: 'comboSecondary' }).map((s) => s.type)).toEqual(['column', 'line']);
    expect(buildSeries('Data', { r1: 0, c1: 0, r2: 3, c2: 3 }, [[null, 'A', 'B', 'C'], ['x', 1, 2, 3], ['y', 1, 2, 3], ['z', 1, 2, 3]], { type: 'comboAreaColumn' }).map((s) => s.type)).toEqual(['area', 'area', 'column']);
  });
  it('bubble takes y and size columns in pairs after the x column', () => {
    const b = [['X', 'Y', 'Size'], [1, 2, 3], [4, 5, 6]];
    const s = buildSeries('S', { r1: 0, c1: 0, r2: 2, c2: 2 }, b, { type: 'bubble' });
    expect(s).toEqual([{ name: 'S!$B$1', values: 'S!$B$2:$B$3', categories: 'S!$A$2:$A$3', sizes: 'S!$C$2:$C$3' }]);
  });
  it('hierarchy charts use every leading text column as a level', () => {
    const t = [['Region', 'City', 'Sales'], ['East', 'Boston', 5], ['East', 'Albany', 3], ['West', 'Reno', 4]];
    const s = buildSeries('S', { r1: 0, c1: 0, r2: 3, c2: 2 }, t, { type: 'treemap' });
    expect(s).toEqual([{ name: 'S!$C$1', values: 'S!$C$2:$C$4', categories: 'S!$A$2:$B$4' }]);
    expect(detectHeaders(t)).toEqual({ firstRow: true, firstCol: true });
  });
});

describe('resolving live data', () => {
  const store = () => {
    const s = new GridStore();
    const sh = s.wb.sheets[0]!;
    const rows: (string | number)[][] = [['Region', 'City', 'Sales'], ['East', 'Boston', 5], ['', 'Albany', 3], ['West', 'Reno', 4]];
    rows.forEach((r, ri) => r.forEach((v, ci) => { if (v !== '') s.engine.setValue(0, ri, ci, v); }));
    void sh;
    return s;
  };
  it('splits multi-column categories into paths and fills blank parents', () => {
    const s = store();
    const name = s.wb.sheets[0]!.name;
    const c: Chart = { id: 'c', type: 'treemap', series: [{ values: `${name}!$C$2:$C$4`, categories: `${name}!$A$2:$B$4` }], anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 1, height: 1 } };
    const d = resolveChart(s, c);
    expect(d.paths).toEqual([['East', 'Boston'], ['East', 'Albany'], ['West', 'Reno']]);
    expect(d.categories).toEqual(['East / Boston', 'East / Albany', 'West / Reno']);
    const tree = buildTree(d.paths, d.series[0]!.values);
    expect(tree.value).toBe(12);
    expect(tree.children.map((n) => [n.name, n.value])).toEqual([['East', 8], ['West', 4]]);
  });
  it('two-axis combo puts its line series on the right', () => {
    const c: Chart = { id: 'c', type: 'comboSecondary', series: [{ values: 'A1', type: 'column' }, { values: 'B1', type: 'line' }], anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 1, height: 1 } };
    expect([...secondaryOf(c)]).toEqual([1]);
    expect([...secondaryOf({ ...c, options: { secondary: [0] } })]).toEqual([0]);
  });
});

describe('histogram bins', () => {
  const data = [1, 2, 2, 3, 3, 3, 4, 4, 5, 9];
  it('counts every value exactly once', () => {
    for (const opts of [{}, { binCount: 4 }, { binWidth: 2 }, { binWidth: 0.7 }]) {
      const bins = histogramBins(data, opts);
      expect(bins.reduce((a, b) => a + b.count, 0)).toBe(data.length);
      expect(bins[0]!.lo).toBeLessThanOrEqual(1);
      expect(bins.at(-1)!.hi).toBeGreaterThanOrEqual(9);
    }
  });
  it('uses the number of bins asked for', () => {
    const bins = histogramBins(data, { binCount: 4 });
    expect(bins).toHaveLength(4);
    expect(bins.map((b) => b.count)).toEqual([6, 3, 0, 1]);
  });
  it('uses the bin width asked for, with upper edges inclusive', () => {
    const bins = histogramBins([0, 2, 2.5, 4, 6], { binWidth: 2 });
    expect(bins.map((b) => [b.lo, b.hi, b.count])).toEqual([[0, 2, 2], [2, 4, 2], [4, 6, 1]]);
    expect(bins[0]!.label).toBe('[0, 2]');
    expect(bins[1]!.label).toBe('(2, 4]');
  });
  it('picks tidy automatic bins', () => {
    const many = Array.from({ length: 100 }, (_, i) => i);
    const bins = histogramBins(many);
    expect(bins.length).toBeGreaterThanOrEqual(5);
    expect(bins.length).toBeLessThanOrEqual(30);
    const w = bins[0]!.hi - bins[0]!.lo;
    expect([1, 2, 2.5, 5, 10, 20, 25, 50]).toContain(w);
    expect(histogramBins([])).toEqual([]);
    expect(histogramBins([7, 7, 7]).reduce((a, b) => a + b.count, 0)).toBe(3);
  });
});

describe('Pareto', () => {
  it('sorts descending and accumulates to 100%', () => {
    const rows = paretoRows(['a', 'b', 'c', 'd'], [10, 40, null, 50]);
    expect(rows.map((r) => r.label)).toEqual(['d', 'b', 'a']);
    expect(rows.map((r) => r.cumulative)).toEqual([0.5, 0.9, 1]);
  });
});

describe('box and whisker', () => {
  it('quartiles, mean, whiskers and outliers', () => {
    const s = boxStats([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 40])!;
    expect(s.median).toBe(6);
    expect(s.q1).toBe(3);
    expect(s.q3).toBe(9);
    expect(s.mean).toBeCloseTo(95 / 11);
    expect(s.outliers).toEqual([40]);
    expect(s.highWhisker).toBe(10);
    expect(s.lowWhisker).toBe(1);
  });
  it('even counts interpolate', () => {
    const s = boxStats([1, 2, 3, 4])!;
    expect(s.median).toBe(2.5);
    expect(s.q1).toBe(1.25);
    expect(s.q3).toBe(3.75);
    expect(boxStats([])).toBeNull();
    expect(boxStats([5])!.median).toBe(5);
  });
});

describe('waterfall', () => {
  it('running totals, decreases and total bars', () => {
    const steps = waterfallSteps([100, -30, 50, null, -20, 0], [5]);
    expect(steps.map((s) => [s.start, s.end, s.kind])).toEqual([
      [0, 100, 'up'], [100, 70, 'down'], [70, 120, 'up'], [120, 120, 'up'], [120, 100, 'down'], [0, 100, 'total'],
    ]);
  });
  it('guesses totals from labels', () => {
    expect(guessTotals(['Start', 'Sales', 'Costs', 'Net', 'Total'])).toEqual([3, 4]);
    expect(guessTotals(['Total'])).toEqual([]);
  });
});

describe('treemap and sunburst', () => {
  it('squarified tiles fill the box and match the weights', () => {
    const weights = [6, 6, 4, 3, 2, 2, 1];
    const box = { x: 10, y: 20, w: 600, h: 400 };
    const tiles = squarify(weights.map((w, i) => ({ item: i, weight: w })), box);
    expect(tiles).toHaveLength(weights.length);
    const area = tiles.reduce((a, t) => a + t.w * t.h, 0);
    expect(area).toBeCloseTo(box.w * box.h, 6);
    const total = weights.reduce((a, b) => a + b, 0);
    for (const t of tiles) {
      expect(t.w * t.h).toBeCloseTo((weights[t.item]! / total) * box.w * box.h, 6);
      expect(t.x).toBeGreaterThanOrEqual(box.x - 1e-9);
      expect(t.y).toBeGreaterThanOrEqual(box.y - 1e-9);
      expect(t.x + t.w).toBeLessThanOrEqual(box.x + box.w + 1e-6);
      expect(t.y + t.h).toBeLessThanOrEqual(box.y + box.h + 1e-6);
    }
    // Squarified tiles stay reasonably close to square.
    const worst = Math.max(...tiles.map((t) => Math.max(t.w / t.h, t.h / t.w)));
    expect(worst).toBeLessThan(4);
  });
  it('sunburst rings split angles by value', () => {
    const tree = buildTree([['A', 'x'], ['A', 'y'], ['B', '']], [1, 3, 4]);
    expect(treeDepth(tree)).toBe(2);
    const arcs = sunburstArcs(tree, 0);
    const inner = arcs.filter((a) => a.depth === 1);
    expect(inner.map((a) => a.node.name)).toEqual(['A', 'B']);
    expect(inner[0]!.a1 - inner[0]!.a0).toBeCloseTo(Math.PI);
    const y = arcs.find((a) => a.node.name === 'y')!;
    expect(y.depth).toBe(2);
    expect(y.a1 - y.a0).toBeCloseTo((Math.PI * 3) / 4);
  });
});

describe('line paths', () => {
  it('straight and smooth', () => {
    expect(linePath([[0, 0], [10, 5]], false)).toBe('M0,0 L10,5');
    const smooth = linePath([[0, 0], [10, 10], [20, 0]], true);
    expect(smooth.startsWith('M0,0 C')).toBe(true);
    expect(smooth.endsWith('20,0')).toBe(true);
  });
});

describe('chart types', () => {
  it('every type has a gallery entry and traits', () => {
    expect(new Set(ALL_CHART_TYPES).size).toBe(CHART_TYPES.length);
    for (const t of ALL_CHART_TYPES) expect(traits(t).family).toBeTruthy();
  });
});

// ---------- .xlsx ----------

const anchor = { row: 2, col: 4, dx: 0, dy: 0, width: 480, height: 300 };
const ref = (r: string) => `Data!${r}`;
const sample = (type: ChartType, extra: Partial<Chart> = {}): Chart => ({
  id: 'x', type, anchor,
  series: [
    { name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') },
    { name: 'South', values: ref('$C$2:$C$5'), categories: ref('$A$2:$A$5') },
  ],
  ...extra,
});

const roundTrip = (charts: Chart[]): Chart[] => {
  const wb = newWorkbook();
  const s = wb.sheets[0]!;
  s.name = 'Data';
  [['Q', 'North', 'South', 'Size'], ['Q1', 1, 2, 3], ['Q2', 4, 5, 6], ['Q3', 7, 8, 9], ['Total', 10, 11, 12]].forEach((r, ri) => r.forEach((v, ci) => s.cells.set(cellKey(ri, ci), { v })));
  s.charts = charts.map((c, i) => ({ ...c, id: `chart${i + 1}` }));
  return readXlsx(writeXlsx(wb)).sheets[0]!.charts;
};

describe('xlsx round trip', () => {
  const direct: Chart[] = [
    sample('stackedColumn'), sample('percentColumn'), sample('percentBar'), sample('stackedBar'),
    sample('line'), sample('lineMarkers'), sample('stackedLine'), sample('percentLine'), sample('stackedLineMarkers'), sample('percentLineMarkers'),
    sample('area'), sample('stackedArea'), sample('percentArea'),
    sample('scatter'), sample('scatterSmoothMarkers'), sample('scatterSmooth'), sample('scatterLinesMarkers'), sample('scatterLines'),
    { id: 'x', type: 'bubble', anchor, series: [{ name: 'B', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5'), sizes: ref('$D$2:$D$5') }] },
    sample('radar'), sample('radarMarkers'), sample('radarFilled'),
    sample('combo', { series: [{ name: 'A', values: ref('$B$2:$B$5'), type: 'column' }, { name: 'B', values: ref('$C$2:$C$5'), type: 'line' }] }),
    sample('comboSecondary', { series: [{ name: 'A', values: ref('$B$2:$B$5'), type: 'column' }, { name: 'B', values: ref('$C$2:$C$5'), type: 'line' }], axisTitles: { y: 'Units', y2: 'Rate' } }),
    sample('comboAreaColumn', { series: [{ name: 'A', values: ref('$B$2:$B$5'), type: 'area' }, { name: 'B', values: ref('$C$2:$C$5'), type: 'column' }] }),
    sample('column', { options: { secondary: [1] } }),
    sample('doughnut', { dataLabels: true, options: { labelMode: 'both' } }),
    sample('pie', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }], dataLabels: true, options: { labelMode: 'value' } }),
    sample('column', { options: { min: -5, max: 50, gridlines: false, numFmt: '#,##0.00' }, axisTitles: { x: 'Quarter', y: 'Amount' }, legend: 'top', title: 'Sales' }),
  ];
  it('kinds with a direct DrawingML form come back unchanged', () => {
    const back = roundTrip(direct);
    back.forEach((c, i) => expect({ ...c, id: 'x' }, direct[i]!.type).toEqual({ ...direct[i]!, id: 'x', legend: direct[i]!.legend ?? 'right' }));
  });

  const extended: Chart[] = [
    sample('histogram', { series: [{ name: 'North', values: ref('$B$2:$B$5') }], options: { binWidth: 2.5 } }),
    sample('histogram', { series: [{ name: 'North', values: ref('$B$2:$B$5') }], options: { binCount: 7 } }),
    sample('pareto', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }] }),
    sample('boxWhisker', { options: { hideMean: true } }),
    sample('treemap', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }], options: { labelMode: 'percent', style: 2 } }),
    sample('sunburst', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }] }),
    sample('waterfall', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }], options: { totals: [3] }, dataLabels: true }),
    sample('funnel', { series: [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }] }),
    sample('stackedColumn', { options: { style: 1 } }),
  ];
  it('extended kinds come back with their real type and options', () => {
    const back = roundTrip(extended);
    back.forEach((c, i) => expect({ ...c, id: 'x' }, extended[i]!.type).toEqual({ ...extended[i]!, id: 'x', legend: 'right' }));
  });

  it('extended kinds are saved as a clustered column other apps can show', () => {
    const xml = chartXml(extended[6]!);
    expect(savedAs('waterfall')).toBe('column');
    expect(xml).toContain('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>');
    expect(xml).toContain('type="waterfall" base="column"');
    // The extension list is the last child of chartSpace, after c:chart.
    expect(xml.indexOf('<c:extLst>')).toBeGreaterThan(xml.indexOf('</c:chart>'));
  });

  it('ignores the extension when another app changed the chart', () => {
    const xml = chartXml(extended[6]!).replace('<c:barDir val="col"/><c:grouping val="clustered"/>', '<c:barDir val="bar"/><c:grouping val="clustered"/>');
    const c = readChart(parseRoot(xml), 'c', anchor)!;
    expect(c.type).toBe('bar');
  });

  it('writes the right DrawingML for each style', () => {
    expect(chartXml(sample('scatterSmooth'))).toContain('<c:scatterStyle val="smooth"/>');
    expect(chartXml(sample('scatterSmooth'))).toContain('<c:smooth val="1"/>');
    expect(chartXml(sample('radarFilled'))).toContain('<c:radarStyle val="filled"/>');
    expect(chartXml(sample('percentArea'))).toContain('<c:grouping val="percentStacked"/>');
    const two = chartXml(sample('comboSecondary', { series: [{ values: 'Data!$B$2:$B$5', type: 'column' }, { values: 'Data!$C$2:$C$5', type: 'line' }] }));
    expect(two).toContain('<c:axPos val="r"/>');
    expect(two.match(/<c:valAx>/g)).toHaveLength(2);
  });

  it('distrusts the options in a file', () => {
    expect(sanitizeOptions({ binCount: 'x', min: Infinity, numFmt: 4, secondary: [1, -1, 2.5, 3], style: 99, evil: true, labelMode: 'nope' })).toEqual({ secondary: [1, 3] });
    expect(sanitizeOptions('x')).toBeUndefined();
    const bad = chartXml(sample('funnel')).replace(/options="[^"]*"/, 'options="{not json"').replace('type="funnel"', 'type="funnel"');
    const withOpts = bad.includes('options=') ? bad : bad.replace('base="column"', 'base="column" options="{not json"');
    expect(readChart(parseRoot(withOpts), 'c', anchor)!.type).toBe('funnel');
    expect(readChart(parseRoot(chartXml(sample('funnel')).replace('type="funnel"', 'type="pizza"')), 'c', anchor)!.type).toBe('column');
  });
});
