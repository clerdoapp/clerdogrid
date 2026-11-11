import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8, zipSync } from 'fflate';
import { cellKey, newSheet, newWorkbook, type Chart, type ChartSeries, type ChartType } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { CLERDO_EXT_URI, chartXml, readChart } from '../src/formats/xlsx-chart';
import { cellLookup, chartExXml, readChartEx } from '../src/formats/xlsx-chartex';
import { sanitizeOptions, sanitizeSeriesFormat } from '../src/formats/chart-options';
import { parseRoot } from '../src/formats/xml';

const anchor = { row: 2, col: 4, dx: 0, dy: 0, width: 480, height: 300 };
const ref = (r: string) => `Data!${r}`;
const two: ChartSeries[] = [
  { name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') },
  { name: 'South', values: ref('$C$2:$C$5'), categories: ref('$A$2:$A$5') },
];
const sample = (type: ChartType, extra: Partial<Chart> = {}, series: ChartSeries[] = two): Chart => ({ id: 'x', type, anchor, series: structuredClone(series), ...extra });

const book = () => {
  const wb = newWorkbook();
  const s = wb.sheets[0]!;
  s.name = 'Data';
  [['Q', 'North', 'South', 'Size', 'Vol', 'Open', 'High', 'Low', 'Close'], ['Q1', 1, 2, 3, 100, 10, 12, 9, 11], ['Q2', 4, 5, 6, 120, 11, 13, 10, 12], ['Q3', 7, 8, 9, 90, 12, 14, 11, 13], ['Total', 10, 11, 12, 80, 13, 15, 12, 12]]
    .forEach((r, ri) => r.forEach((v, ci) => s.cells.set(cellKey(ri, ci), { v })));
  return wb;
};

const roundTrip = (charts: Chart[]): Chart[] => {
  const wb = book();
  wb.sheets[0]!.charts = charts.map((c, i) => ({ ...c, id: `chart${i + 1}` }));
  return readXlsx(writeXlsx(wb)).sheets[0]!.charts;
};
const same = (back: Chart, orig: Chart) => expect({ ...back, id: 'x' }, orig.type).toEqual({ ...orig, id: 'x', legend: orig.legend ?? 'right' });

/** Reads a chart's DrawingML without our extension, as another app would leave it. */
const native = (c: Chart): Chart => {
  const xml = chartXml(c).replace(new RegExp(`<c:extLst><c:ext uri="${CLERDO_EXT_URI.replace(/[{}]/g, '\\$&')}">.*?</c:ext></c:extLst>`), '');
  expect(xml).not.toContain(CLERDO_EXT_URI);
  return readChart(parseRoot(xml), 'x', anchor)!;
};

describe('new chart types in DrawingML', () => {
  const types: ChartType[] = [
    'column3DClustered', 'column3DStacked', 'column3DPercent', 'column3D', 'cylinderColumn', 'coneColumn', 'pyramidColumn',
    'bar3DClustered', 'bar3DStacked', 'bar3DPercent', 'line3D', 'area3D', 'stackedArea3D', 'percentArea3D',
    'pie3D', 'pieExploded', 'pie3DExploded', 'doughnutExploded', 'pieOfPie', 'barOfPie', 'surface3D', 'surfaceWireframe3D', 'contour', 'contourWireframe', 'bubble3D',
  ];
  for (const t of types) it(`${t} saves and reopens as itself`, () => {
    const series = /pie|Pie/.test(t) ? two.slice(0, 1) : t === 'bubble3D' ? [{ ...two[0]!, sizes: ref('$D$2:$D$5') }] : two;
    const c = sample(t, {}, series);
    const [back] = roundTrip([c]);
    same(back!, c);
    expect(chartXml(c)).not.toContain(CLERDO_EXT_URI);
  });

  it('stock charts, with and without volume', () => {
    const s = (cols: string[]): ChartSeries[] => cols.map((col) => ({ name: col, values: ref(`$${col}$2:$${col}$5`), categories: ref('$A$2:$A$5') }));
    const list = [sample('stockHLC', {}, s(['G', 'H', 'I'])), sample('stockOHLC', {}, s(['F', 'G', 'H', 'I'])), sample('stockVHLC', {}, s(['E', 'G', 'H', 'I'])), sample('stockVOHLC', {}, s(['E', 'F', 'G', 'H', 'I']))];
    const back = roundTrip(list);
    back.forEach((c, i) => same(c, list[i]!));
    const xml = chartXml(list[3]!);
    expect(xml).toContain('<c:stockChart>');
    expect(xml).toContain('<c:hiLowLines/>');
    expect(xml).toContain('<c:upDownBars>');
    expect(xml.indexOf('<c:barChart>')).toBeLessThan(xml.indexOf('<c:stockChart>'));
    expect(chartXml(list[0]!)).toContain('<c:symbol val="dash"/>');
  });

  it('writes the right 3D elements', () => {
    const x = chartXml(sample('coneColumn', { options: { view3D: { rotX: 25, rotY: 40, perspective: 60, depthPercent: 150, gapDepth: 80 } } }));
    expect(x).toMatch(/<c:view3D><c:rotX val="25"\/><c:rotY val="40"\/><c:depthPercent val="150"\/><c:rAngAx val="0"\/><c:perspective val="60"\/><\/c:view3D>/);
    expect(x).toContain('<c:shape val="cone"/>');
    expect(x).toContain('<c:gapDepth val="80"/>');
    expect(chartXml(sample('column3DStacked', { options: { shape: 'pyramid' } }))).toContain('<c:shape val="pyramidToMax"/>');
    expect(chartXml(sample('column3D'))).toContain('<c:serAx>');
    expect(chartXml(sample('contour'))).toContain('<c:surfaceChart><c:wireframe val="0"/>');
    expect(chartXml(sample('barOfPie', {}, two.slice(0, 1)))).toContain('<c:ofPieType val="bar"/>');
  });
});

describe('chart elements in DrawingML', () => {
  const rich: Chart[] = [
    sample('scatter', {}, [{ ...two[0]!, trendlines: [{ kind: 'poly', order: 3, forward: 2, backward: 1, showEquation: true, showR2: true, name: 'Fit', line: { color: '#FF0000', width: 1.5, dash: 'dash' } }, { kind: 'movingAvg', period: 3 }],
      errorBars: [{ dir: 'x', kind: 'fixed', value: 0.5 }, { kind: 'custom', plus: ref('$D$2:$D$5'), minus: '={1,2,3,4}', noCap: true }] }]),
    sample('lineMarkers', {}, [{ ...two[0]!, trendlines: [{ kind: 'linear', intercept: 2 }, { kind: 'exp' }, { kind: 'log' }, { kind: 'power' }], errorBars: [{ kind: 'percent', value: 10, type: 'plus' }] }, { ...two[1]!, errorBars: [{ kind: 'stdDev', value: 2 }] }]),
    sample('column', { options: { gapWidth: 60, overlap: -20, varyColors: true } }, [{ ...two[0]!, errorBars: [{ kind: 'stdErr', type: 'minus' }], invertIfNegative: true, fill: { gradient: { from: '#112233', to: '#445566', angle: 45 } }, line: { color: '#000000', width: 0.75 },
      points: [{ idx: 1, fill: { color: '#FF0000' } }, { idx: 2, line: { color: '#00FF00', dash: 'dot' } }] }]),
    sample('lineMarkers', {}, [{ ...two[0]!, line: { width: 3, dash: 'longDash' }, marker: { symbol: 'diamond', size: 8, fill: '#123456', line: '#654321' }, smooth: true, points: [{ idx: 0, marker: { symbol: 'star', size: 12 } }] }]),
    sample('pie', {}, [{ ...two[0]!, points: [{ idx: 2, explosion: 30 }] }]),
    sample('pieExploded', {}, [{ ...two[0]!, explosion: 30 }]),
    sample('column', { dataLabels: true, options: { labels: { value: true, seriesName: true, categoryName: true, separator: '; ', position: 'insideEnd', numFmt: '0.0' } } }),
    sample('column', {}, [{ ...two[0]!, labels: { value: true, fromCells: ref('$A$2:$A$5') } }, { ...two[1]!, labels: { categoryName: true, position: 'center' } }]),
    sample('pie', { dataLabels: true, options: { labels: { categoryName: true, percent: true, position: 'outsideEnd', leaderLines: false } } }, two.slice(0, 1)),
    sample('lineMarkers', { options: { upDownBars: true, upFill: '#00AA00', downFill: '#AA0000', hiLowLines: true, dropLines: true } }),
    sample('stackedColumn', { options: { seriesLines: true, dataTable: { keys: false, vertical: false } } }),
    sample('line', { options: { min: 1, max: 20, numFmt: '0.00', gridlines: false, axes: {
      y: { logBase: 10, reverse: true, majorUnit: 5, minorUnit: 1, dispUnits: 'thousands', majorTick: 'cross', minorTick: 'in', labelPos: 'low', minorGrid: true, crossesAt: 3, line: { color: '#333333', width: 1 } },
      x: { reverse: true, crosses: 'max', majorGrid: true, minorGrid: true, labelPos: 'high', hidden: true, baseUnit: 'months' },
    } } }),
    sample('column', { options: { secondary: [1], axes: { y: { dispUnits: 'millions', dispUnitsLabel: false }, y2: { min: 0, max: 50, numFmt: '0%', reverse: true }, x: { dateAxis: false } } }, axisTitles: { x: '=Data!$A$1', y: 'Amount', y2: '=Data!$C$1' } }),
    sample('scatterLines', { options: { secondary: [1], axes: { x: { min: 0, max: 12, majorGrid: false, logBase: 2 } } } }),
    sample('area', { title: '=Data!$B$1', legend: 'left', options: { legendOverlay: true, blanks: 'span', showHidden: true, chartArea: { fill: { color: '#F0F0F0' }, line: { color: '#000000', width: 1 } }, plotArea: { fill: { gradient: { from: '#FFFFFF', to: '#EEEEEE' } } } } }),
    sample('column3DStacked', { options: { view3D: { rotX: 10, rotY: 30, rightAngle: true, heightPercent: 80 }, walls: { fill: { color: '#EEEEEE' } }, floor: { fill: { color: '#DDDDDD' } }, shape: 'cylinder' } }),
    sample('pieOfPie', { options: { ofPie: { split: 'val', splitValue: 5, secondSize: 60, gapWidth: 150 } } }, two.slice(0, 1)),
    sample('doughnut', { options: { holeSize: 70, firstSliceAngle: 90, varyColors: false } }),
    sample('bubble', { options: { bubbleScale: 150 } }, [{ ...two[0]!, sizes: ref('$D$2:$D$5') }]),
    sample('stockOHLC', { options: { upDownBars: false, hiLowLines: false } }, ['F', 'G', 'H', 'I'].map((c) => ({ name: c, values: ref(`$${c}$2:$${c}$5`), categories: ref('$A$2:$A$5') }))),
  ];
  it('come back from DrawingML alone', () => {
    rich.forEach((c) => same(native(c), c));
  });
  it('come back through a saved file', () => {
    roundTrip(rich).forEach((c, i) => same(c, rich[i]!));
  });
  it('writes value-from-cells labels with the shared extension', () => {
    const x = chartXml(rich.find((c) => c.series.some((s) => s.labels?.fromCells))!);
    expect(x).toContain('<c15:datalabelsRange><c15:f>Data!$A$2:$A$5</c15:f></c15:datalabelsRange>');
    expect(x).toContain('<c15:showDataLabelsRange val="1"/>');
  });
  it('links titles to cells', () => {
    expect(chartXml(rich.find((c) => c.title?.startsWith('='))!)).toContain('<c:title><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>');
  });
  it('keeps non-native settings in the extension', () => {
    const c = sample('column', { options: { look: 2, style: 3 } });
    expect(chartXml(c)).toContain(CLERDO_EXT_URI);
    same(roundTrip([c])[0]!, c);
  });
  it('only writes label positions the chart kind allows', () => {
    const stacked = chartXml(sample('stackedColumn', { dataLabels: true, options: { labels: { value: true, position: 'outsideEnd' } } }));
    expect(stacked).not.toContain('dLblPos');
    const line = chartXml(sample('line', { dataLabels: true, options: { labels: { value: true, position: 'above' } } }));
    expect(line).toContain('<c:dLblPos val="t"/>');
    expect(chartXml(sample('area', { dataLabels: true, options: { labels: { value: true, position: 'center' } } }))).not.toContain('dLblPos');
  });
});

describe('extended chart parts', () => {
  const wb = book();
  const look = cellLookup(wb);
  const one = [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }];
  const kinds: Chart[] = [
    sample('histogram', { options: { binCount: 4 } }, [{ name: 'North', values: ref('$B$2:$B$5') }]),
    sample('pareto', {}, one),
    sample('boxWhisker', { options: { hideMean: true, hideOutliers: true } }, two),
    sample('waterfall', { options: { totals: [3] }, dataLabels: true }, one),
    sample('funnel', {}, one),
    sample('treemap', {}, [{ name: 'North', values: ref('$B$2:$B$5'), categories: ref('$A$2:$A$5') }]),
    sample('sunburst', {}, one),
    sample('regionMap', { options: { map: { low: '#FFFFFF', high: '#0C7A4B', labels: 'best' } } }, one),
  ];
  it('writes cx parts with cached values', () => {
    const x = chartExXml(kinds[1]!, look);
    expect(x).toContain('xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"');
    expect(x).toContain('layoutId="clusteredColumn"');
    expect(x).toContain('layoutId="paretoLine" ownerIdx="0"');
    expect(x).toContain('<cx:pt idx="1">4</cx:pt>');
    expect(x).toContain('<cx:pt idx="0">Q1</cx:pt>');
    expect(chartExXml(kinds[3]!, look)).toContain('<cx:subtotals><cx:idx val="3"/></cx:subtotals>');
    expect(chartExXml(kinds[7]!, look)).toContain('<cx:numDim type="colorVal">');
  });
  it('reads cx parts from other apps', () => {
    kinds.forEach((c) => {
      const back = readChartEx(parseRoot(chartExXml(c, look)), 'x', anchor)!;
      expect(back.type, c.type).toBe(c.type);
      expect(back.series[0]!.values).toBe(c.series[0]!.values.replace(/^=/, ''));
    });
    const h = readChartEx(parseRoot(chartExXml(kinds[0]!, look)), 'x', anchor)!;
    expect(h.options?.binCount).toBe(4);
    const w = readChartEx(parseRoot(chartExXml(kinds[3]!, look)), 'x', anchor)!;
    expect(w.options?.totals).toEqual([3]);
    const b = readChartEx(parseRoot(chartExXml(kinds[2]!, look)), 'x', anchor)!;
    expect(b.options).toMatchObject({ hideMean: true, hideOutliers: true });
    const m = readChartEx(parseRoot(chartExXml(kinds[7]!, look)), 'x', anchor)!;
    expect(m.options?.map).toMatchObject({ low: '#FFFFFF', high: '#0C7A4B', labels: 'best' });
  });
  it('saves the drawing with an alternate content choice and a classic fallback', () => {
    const w2 = book();
    w2.sheets[0]!.charts = kinds.map((c, i) => ({ ...c, id: `chart${i + 1}` }));
    const files = unzipSync(writeXlsx(w2));
    const drawing = strFromU8(files['xl/drawings/drawing1.xml']!);
    expect(drawing).toContain('<mc:AlternateContent');
    expect(drawing).toContain('Requires="cx1"');
    expect(drawing).toContain('Requires="cx2"');
    expect(drawing).toContain('Requires="cx4"');
    expect(drawing).toContain('<mc:Fallback>');
    const rels = strFromU8(files['xl/drawings/_rels/drawing1.xml.rels']!);
    expect(rels).toContain('http://schemas.microsoft.com/office/2014/relationships/chartEx');
    const ct = strFromU8(files['[Content_Types].xml']!);
    expect(ct).toContain('application/vnd.ms-office.chartex+xml');
    expect(files['xl/charts/chartEx1.xml']).toBeTruthy();
    // Reopened, our own fallback (with its settings) wins.
    const back = readXlsx(writeXlsx(w2)).sheets[0]!.charts;
    back.forEach((c, i) => same(c, kinds[i]!));
  });
  it('uses another app\'s extended part when the fallback is a plain chart', () => {
    const w2 = book();
    w2.sheets[0]!.charts = [{ ...kinds[3]!, id: 'chart1' }];
    const files = unzipSync(writeXlsx(w2));
    // Drop our extension from the fallback, as another app's file wouldn't have it.
    files['xl/charts/chart1.xml'] = new TextEncoder().encode(strFromU8(files['xl/charts/chart1.xml']!).replace(/<c:extLst>.*<\/c:extLst>/, ''));
    const back = readXlsx(zipSync(files)).sheets[0]!.charts[0]!;
    expect(back.type).toBe('waterfall');
    expect(back.options?.totals).toEqual([3]);
  });
});

describe('chart sheets', () => {
  it('save and reopen as chart sheets', () => {
    const wb = book();
    const cs = newSheet('Sales chart');
    cs.kind = 'chart';
    cs.tabColor = '#12A565';
    cs.charts = [{ ...sample('column', { title: 'Sales' }), id: 'chart1', anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 910, height: 660 } }];
    wb.sheets.push(cs);
    const cs2 = newSheet('Map');
    cs2.kind = 'chart';
    cs2.charts = [{ ...sample('waterfall', { options: { totals: [3] } }, two.slice(0, 1)), id: 'chart1', anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 910, height: 660 } }];
    wb.sheets.push(cs2);
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    expect(strFromU8(files['xl/workbook.xml']!)).toContain('name="Sales chart"');
    expect(strFromU8(files['xl/chartsheets/sheet2.xml']!)).toContain('<chartsheet');
    expect(strFromU8(files['xl/_rels/workbook.xml.rels']!)).toContain('relationships/chartsheet');
    const back = readXlsx(bytes);
    expect(back.sheets.map((s) => s.kind)).toEqual([undefined, 'chart', 'chart']);
    expect(back.sheets[1]!.charts[0]).toMatchObject({ type: 'column', title: 'Sales' });
    expect(back.sheets[1]!.tabColor).toBe('#12A565');
    expect(back.sheets[2]!.charts[0]).toMatchObject({ type: 'waterfall', options: { totals: [3] } });
  });
});

describe('untrusted settings', () => {
  it('keep only known, well-formed fields', () => {
    expect(sanitizeOptions({ axes: { y: { logBase: 1, min: 'x', reverse: true, evil: 1 } }, view3D: { rotX: 400, rotY: 30 }, map: { low: 'red', high: '#00FF00' }, look: 3 }))
      .toEqual({ axes: { y: { reverse: true } }, view3D: { rotY: 30 }, map: { high: '#00FF00' }, look: 3 });
    expect(sanitizeSeriesFormat({ trendlines: [{ kind: 'poly', order: 9 }, { kind: 'nope' }], values: 'A1', marker: { symbol: 'star', size: 500 } }))
      .toEqual({ trendlines: [{ kind: 'poly' }], marker: { symbol: 'star' } });
  });
});
