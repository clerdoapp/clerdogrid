import { beforeAll, describe, expect, test } from 'vitest';
import { registerAllFunctions } from '../src/engine/functions';
import { GridStore } from '../src/app/store';
import { analyze, findCorrelations, findOutliers, findTrend, linearFit, matchColumn, parseAsk, summarizeColumns } from '../src/features/automate/insights';
import { insertSummary, readBlock, runAsk, summaryFormula } from '../src/features/automate/insightActions';
import type { Scalar } from '../src/engine/types';

beforeAll(() => registerAllFunctions());

const headers = ['Month', 'Region', 'Sales', 'Units'];
const rows: Scalar[][] = [
  [1, 'East', 100, 10], [2, 'West', 120, 12], [3, 'East', 140, 15], [4, 'North', 160, 16],
  [5, 'West', 180, 19], [6, 'East', 200, 20], [7, 'East', 220, 21], [8, 'West', 240, 25],
];

describe('summaries', () => {
  test('numbers and text', () => {
    const cols = summarizeColumns(['N', 'T'], [[1, 'a'], [2, 'b'], [3, 'a'], [null, ''], [10, 'a']]);
    expect(cols[0]).toMatchObject({ type: 'number', count: 4, blanks: 1, distinct: 4, min: 1, max: 10, mean: 4, median: 2.5, sum: 16 });
    expect(cols[1]).toMatchObject({ type: 'text', count: 4, blanks: 1, distinct: 2 });
    expect(cols[1]!.top![0]).toEqual({ value: 'a', count: 3 });
  });

  test('dates come from the caller', () => {
    const cols = summarizeColumns(['D'], [[45000], [45001]], (c) => c === 0);
    expect(cols[0]!.type).toBe('date');
  });
});

describe('patterns', () => {
  test('outliers beyond 1.5 IQR', () => {
    const data: Scalar[][] = [[10], [11], [12], [11], [10], [12], [11], [95]];
    const cols = summarizeColumns(['X'], data);
    const out = findOutliers(cols, data);
    expect(out.map((o) => o.value)).toEqual([95]);
    expect(out[0]!.row).toBe(7);
  });

  test('linear fit', () => {
    const f = linearFit([1, 2, 3, 4], [3, 5, 7, 9]);
    expect(f.slope).toBeCloseTo(2);
    expect(f.intercept).toBeCloseTo(1);
    expect(f.r2).toBeCloseTo(1);
  });

  test('trend over a sequence column', () => {
    const cols = summarizeColumns(headers, rows);
    const t = findTrend(cols, rows)!;
    expect(t.xName).toBe('Month');
    expect(t.yName).toBe('Sales');
    expect(t.slope).toBeCloseTo(20);
    expect(t.r2).toBeCloseTo(1);
  });

  test('correlations above 0.6 only', () => {
    const cols = summarizeColumns(['A', 'B', 'C'], [[1, 2, 5], [2, 4, 1], [3, 6, 4], [4, 8, 2], [5, 10, 3]]);
    const c = findCorrelations(cols, [[1, 2, 5], [2, 4, 1], [3, 6, 4], [4, 8, 2], [5, 10, 3]]);
    expect(c.length).toBe(1);
    expect(c[0]).toMatchObject({ aName: 'A', bName: 'B' });
    expect(c[0]!.r).toBeCloseTo(1);
  });

  test('analyze gives top categories and visuals', () => {
    const ins = analyze(headers, rows);
    expect(ins.topCategories[0]).toMatchObject({ catName: 'Region', valName: 'Sales' });
    expect(ins.topCategories[0]!.items[0]).toEqual({ key: 'East', total: 660 });
    expect(ins.visuals.map((v) => v.kind)).toContain('line');
    expect(ins.correlations.some((c) => c.aName === 'Sales' && c.bName === 'Units')).toBe(true);
  });
});

describe('ask', () => {
  test('column matching', () => {
    expect(matchColumn(headers, 'sales')).toBe(2);
    expect(matchColumn(headers, 'regions')).toBe(1);
    expect(matchColumn(headers, 'each region')).toBe(1);
    expect(matchColumn(headers, 'unit')).toBe(3);
    expect(matchColumn(headers, 'colour')).toBe(-1);
  });

  test('patterns', () => {
    expect(parseAsk('total sales by region', headers)).toEqual({ kind: 'group', agg: 'sum', val: 2, by: 1 });
    expect(parseAsk('Average Units per Region?', headers)).toEqual({ kind: 'group', agg: 'average', val: 3, by: 1 });
    expect(parseAsk('sales by region', headers)).toEqual({ kind: 'group', agg: 'sum', val: 2, by: 1 });
    expect(parseAsk('top 5 region by sales', headers)).toEqual({ kind: 'top', n: 5, by: 1, val: 2, bottom: false });
    expect(parseAsk('bottom 2 regions by units', headers)).toEqual({ kind: 'top', n: 2, by: 1, val: 3, bottom: true });
    expect(parseAsk('count by region', headers)).toEqual({ kind: 'group', agg: 'count', by: 1 });
    expect(parseAsk('sales over time', headers)).toEqual({ kind: 'trend', val: 2, x: 0 });
    expect(parseAsk('units over month', headers)).toEqual({ kind: 'trend', val: 3, x: 0 });
    expect(parseAsk('show me max sales', headers)).toEqual({ kind: 'total', agg: 'max', val: 2 });
    expect(parseAsk('what is the weather', headers)).toBeNull();
    expect(parseAsk('total colour by region', headers)).toBeNull();
  });
});

describe('inserting', () => {
  const setup = () => {
    const s = new GridStore();
    s.change('fill', 'book', () => [headers, ...rows].forEach((r, i) => r.forEach((v, j) => s.enter(i, j, String(v)))));
    s.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    return s;
  };

  test('reads the region around the active cell with headers', () => {
    const b = readBlock(setup())!;
    expect(b.headers).toEqual(headers);
    expect(b.rows.length).toBe(8);
    expect(b.range).toEqual({ r1: 0, c1: 0, r2: 8, c2: 3 });
  });

  test('summary table with chart as one undo step', () => {
    const s = setup();
    const b = readBlock(s)!;
    expect(summaryFormula(b, { by: 1, val: 2, agg: 'sum', title: 'x' })).toBe("GROUPBY(Sheet1!$B$2:$B$9,Sheet1!$C$2:$C$9,LAMBDA(x,SUM(x)),0,1,-2)");
    insertSummary(s, b, { by: 1, val: 2, agg: 'sum', chart: 'column', title: 'Sales by Region' });
    expect(s.sheet.name).toBe('Sales by Region');
    expect([s.value(0, 0), s.value(0, 1)]).toEqual(['Region', 'Total of Sales']);
    expect([s.value(1, 0), s.value(1, 1)]).toEqual(['East', 660]);
    expect([s.value(2, 0), s.value(2, 1)]).toEqual(['West', 540]);
    expect([s.value(4, 0), s.value(4, 1)]).toEqual(['Total', 1360]);
    expect(s.sheet.charts.length).toBe(1);
    expect(s.sheet.charts[0]!.series[0]!.values).toBe("'Sales by Region'!$B$2:$B$4");
    s.undo();
    expect(s.wb.sheets.length).toBe(1);
  });

  test('ask: top N and totals', () => {
    const s = setup();
    const b = readBlock(s)!;
    expect(runAsk(s, b, parseAsk('top 2 region by sales', b.headers)!)).toMatch(/Top 2/);
    expect([s.value(0, 0), s.value(1, 0), s.value(2, 0), s.value(3, 0)]).toEqual(['Region', 'East', 'West', null]);
    s.setActiveSheet(0);
    expect(runAsk(s, b, parseAsk('average sales', b.headers)!)).toBe('Average of Sales: 170');
  });

  test('ask: count by and over time', () => {
    const s = setup();
    const b = readBlock(s)!;
    runAsk(s, b, parseAsk('count by region', b.headers)!);
    expect([s.value(1, 0), s.value(1, 1)]).toEqual(['East', 4]);
    s.setActiveSheet(0);
    runAsk(s, b, parseAsk('sales over time', b.headers)!);
    expect(s.sheet.charts.at(-1)!.type).toBe('line');
    expect(s.sheet.charts.at(-1)!.series[0]!.categories).toBe('Sheet1!$A$2:$A$9');
  });
});
