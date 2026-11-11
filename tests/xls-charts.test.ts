import { describe, expect, it } from 'vitest';
import { cellKey, newSheet, type Chart, type Workbook } from '../src/model/workbook';
import { writeXls, xlsLimits } from '../src/formats/biff8-write';
import { readXls } from '../src/formats/biff8-read';
import { readXlsChart, xlsChartRecords } from '../src/formats/biff8-chart';

const book = (): Workbook => {
  const wb: Workbook = { sheets: [newSheet('Data')], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const s = wb.sheets[0]!;
  ['Mon', 'Tue', 'Wed', 'Thu'].forEach((d, i) => {
    s.cells.set(cellKey(i + 1, 0), { v: d });
    s.cells.set(cellKey(i + 1, 1), { v: i * 3 + 1 });
    s.cells.set(cellKey(i + 1, 2), { v: 10 - i });
  });
  s.cells.set(cellKey(0, 1), { v: 'Sales' });
  return wb;
};
const series = () => [
  { name: 'Data!$B$1', values: 'Data!$B$2:$B$5', categories: 'Data!$A$2:$A$5', color: '#FF0000' },
  { name: 'Cost', values: 'Data!$C$2:$C$5', categories: 'Data!$A$2:$A$5', color: '#00B050' },
];
const anchor = (i: number): Chart['anchor'] => ({ row: 8 + i * 16, col: 4, dx: 0, dy: 0, width: 400, height: 240 });

describe('.xls charts', () => {
  const types: Chart['type'][] = ['column', 'percentColumn', 'stackedBar', 'lineMarkers', 'line', 'stackedArea', 'pie', 'doughnut', 'scatter', 'scatterLinesMarkers', 'bubble', 'radarMarkers', 'radarFilled', 'combo', 'comboSecondary'];
  const wb = book();
  types.forEach((type, i) => {
    const c: Chart = { id: `c${i}`, type, series: series(), anchor: anchor(i), legend: i % 2 ? 'bottom' : 'right' };
    if (type === 'pie' || type === 'doughnut') c.series = [c.series[0]!];
    if (type === 'bubble') c.series = [{ ...c.series[0]!, sizes: 'Data!$C$2:$C$5' }];
    if (type.startsWith('combo')) { c.series[0]!.type = 'column'; c.series[1]!.type = 'line'; }
    if (type === 'comboSecondary') c.options = { secondary: [1] };
    if (i === 0) { c.title = 'Weekly'; c.axisTitles = { x: 'Day', y: 'Amount' }; }
    if (i === 1) c.title = '=Data!$B$1';
    wb.sheets[0]!.charts.push(c);
  });
  wb.sheets[0]!.shapes = [
    { id: 's1', kind: 'textBox', text: 'Hello box', anchor: { row: 1, col: 4, dx: 0, dy: 0, width: 150, height: 40 }, fill: '#FFFFCC' },
    { id: 's2', kind: 'ellipse', anchor: { row: 1, col: 8, dx: 0, dy: 0, width: 80, height: 40 }, fill: '#00AA00' },
    { id: 's3', kind: 'rightArrow', text: 'Go', anchor: { row: 1, col: 10, dx: 0, dy: 0, width: 80, height: 40 }, fill: '#3366FF' },
  ];
  wb.sheets[0]!.notes.set(cellKey(1, 0), { author: 'A', text: 'note' });
  const back = readXls(writeXls(wb));
  const s = back.sheets[0]!;

  it('saves and reopens every chart kind the format has', () => {
    expect(s.charts.map((c) => c.type)).toEqual(types);
    s.charts.forEach((c, i) => {
      const want = wb.sheets[0]!.charts[i]!;
      expect(c.series.map((x) => [x.name, x.values, x.categories])).toEqual(want.series.map((x) => [x.name, x.values, x.categories]));
      expect(c.legend).toBe(want.legend);
      expect(c.anchor).toMatchObject({ row: want.anchor.row, col: want.anchor.col });
    });
    expect(s.charts[0]).toMatchObject({ title: 'Weekly', axisTitles: { x: 'Day', y: 'Amount' } });
    expect(s.charts[0]!.series.map((x) => x.color)).toEqual(['#FF0000', '#00B050']);
    expect(s.charts[1]!.title).toBe('=Data!$B$1');
    expect(s.charts[10]!.series[0]!.sizes).toBe('Data!$C$2:$C$5');
    expect(s.charts[13]!.series.map((x) => x.type)).toEqual(['column', 'line']);
    expect(s.charts[14]!.options?.secondary).toEqual([1]);
  });
  it('keeps shapes, text boxes and notes together', () => {
    expect(s.shapes?.map((x) => [x.kind, x.text, x.fill])).toEqual([['textBox', 'Hello box', '#FFFFCC'], ['ellipse', undefined, '#00AA00'], ['rightArrow', 'Go', '#3366FF']]);
    expect(s.notes.get(cellKey(1, 0))).toEqual({ author: 'A', text: 'note' });
  });
  it('no longer warns about charts and shapes', () => {
    expect(xlsLimits(wb)).toEqual([]);
    const other = book();
    other.sheets[0]!.charts.push({ id: 'h', type: 'histogram', series: series(), anchor: anchor(0) });
    expect(xlsLimits(other)).toEqual(['Chart types the older format lacks are saved as the nearest simple type.']);
  });
  it('reads a chart substream on its own', () => {
    const c: Chart = { id: 'x', type: 'stackedLineMarkers', series: series(), anchor: anchor(0), legend: 'top', dataLabels: true };
    const recs = xlsChartRecords(c, () => null);
    expect(recs[0]!.type).toBe(0x0809);
    expect(recs.at(-1)!.type).toBe(0x000a);
    // without references the series have no values, so nothing can be drawn
    expect(readXlsChart(recs, () => undefined)).toBeNull();
  });
});
