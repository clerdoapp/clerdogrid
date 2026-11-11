import { beforeAll, describe, expect, it } from 'vitest';
import { registerAllFunctions } from '../src/engine/functions';
import type { Scalar } from '../src/engine/types';
import { VALUES_FIELD, type PivotTable } from '../src/model/pivot';
import { buildCache, groupByRange, groupItems, addCalcField } from '../src/features/pivot/cache';
import { computePivot } from '../src/features/pivot/compute';
import { newPivotTable, placeField, setField, setLayout, makeDataField, remapTable } from '../src/features/pivot/model';

beforeAll(() => registerAllFunctions());

const HEADER = ['Region', 'Product', 'Year', 'Sales', 'Units'];
const ROWS: Scalar[][] = [
  ['East', 'Apple', 2023, 10, 1],
  ['East', 'Pear', 2023, 20, 2],
  ['East', 'Apple', 2024, 30, 3],
  ['West', 'Apple', 2023, 5, 1],
  ['West', 'Pear', 2024, 15, 5],
  ['West', 'Pear', 2024, 25, 5],
  ['North', 'Apple', 2024, 7, 7],
];
const cache = () => buildCache(1, { kind: 'range', sheet: 'Data', range: { r1: 0, c1: 0, r2: 7, c2: 4 } }, { header: HEADER, rows: ROWS, formats: [] });
const grid = (pt: PivotTable, c = cache()) => computePivot(pt, c).cells.map((r) => r.map((x) => (x ? x.v : null)));
const base = (c = cache()) => newPivotTable('PivotTable1', c, { row: 2, col: 0 });

describe('pivot compute', () => {
  it('rows and one value in compact form', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    expect(grid(pt, c)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['East', 60], ['North', 7], ['West', 45], ['Grand Total', 112],
    ]);
    const l = computePivot(pt, c);
    expect([l.firstHeaderRow, l.firstDataRow, l.firstDataCol]).toEqual([1, 1, 1]);
    expect(l.rowItems).toEqual([{ x: [0] }, { x: [1] }, { x: [2] }, { t: 'grand', x: [0] }]);
  });

  it('rows, columns, subtotals on top in compact form', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 1, 'rows');
    pt = placeField(pt, c, 2, 'cols');
    pt = placeField(pt, c, 3, 'data');
    expect(grid(pt, c)).toEqual([
      ['Sum of Sales', 'Column Labels', null, null],
      ['Row Labels', 2023, 2024, 'Grand Total'],
      ['East', 30, 30, 60],
      ['Apple', 10, 30, 40],
      ['Pear', 20, null, 20],
      ['North', null, 7, 7],
      ['Apple', null, 7, 7],
      ['West', 5, 40, 45],
      ['Apple', 5, null, 5],
      ['Pear', null, 40, 40],
      ['Grand Total', 35, 77, 112],
    ]);
    const l = computePivot(pt, c);
    expect(l.cells[3]![0]!.indent).toBe(1);
    expect(l.rowItems.slice(0, 3)).toEqual([{ x: [0] }, { r: 1, x: [0] }, { r: 1, x: [1] }]);
    expect(l.colItems).toEqual([{ x: [0] }, { x: [1] }, { t: 'grand', x: [0] }]);
  });

  it('tabular form puts subtotals at the bottom', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 1, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = setLayout(pt, 'tabular');
    expect(grid(pt, c)).toEqual([
      ['Region', 'Product', 'Sum of Sales'],
      ['East', 'Apple', 40], [null, 'Pear', 20], ['East Total', null, 60],
      ['North', 'Apple', 7], ['North Total', null, 7],
      ['West', 'Apple', 5], [null, 'Pear', 40], ['West Total', null, 45],
      ['Grand Total', null, 112],
    ]);
    const l = computePivot(pt, c);
    expect(l.rowItems.slice(0, 3)).toEqual([{ x: [0, 0] }, { r: 1, x: [1] }, { t: 'default', x: [0] }]);
  });

  it('several values with Σ in columns', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = placeField(pt, c, 4, 'data');
    expect(pt.cols).toEqual([VALUES_FIELD]);
    expect(grid(pt, c)).toEqual([
      ['Row Labels', 'Sum of Sales', 'Sum of Units'],
      ['East', 60, 6], ['North', 7, 7], ['West', 45, 11], ['Grand Total', 112, 24],
    ]);
    expect(computePivot(pt, c).firstHeaderRow).toBe(0);
  });

  it('Σ in rows', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = placeField(pt, c, 4, 'data');
    pt = placeField(pt, c, VALUES_FIELD, 'rows');
    expect(grid(pt, c).slice(0, 4)).toEqual([['Row Labels', null], ['East', null], ['Sum of Sales', 60], ['Sum of Units', 6]]);
    expect(grid(pt, c).slice(-2)).toEqual([['Total Sum of Sales', 112], ['Total Sum of Units', 24]]);
  });

  it('summary functions', () => {
    const c = cache();
    const run = (agg: PivotTable['data'][number]['agg']) => {
      let pt = placeField(base(c), c, 0, 'rows');
      pt = { ...pt, data: [{ ...makeDataField(pt, c, 3, agg) }] };
      return grid(pt, c).map((r) => r[1]);
    };
    expect(run('count')).toEqual(['Count of Sales', 3, 1, 3, 7]);
    expect(run('average')).toEqual(['Average of Sales', 20, 7, 15, 16]);
    expect(run('max')).toEqual(['Max of Sales', 30, 7, 25, 30]);
    expect(run('min')).toEqual(['Min of Sales', 10, 7, 5, 5]);
    expect(run('product')).toEqual(['Product of Sales', 6000, 7, 1875, 6000 * 7 * 1875]);
    const sd = run('stdDev');
    expect(sd[1]).toBeCloseTo(10, 10);
    expect(run('var')[3]).toBeCloseTo(100, 10);
    expect(run('varp')[3]).toBeCloseTo(200 / 3, 10);
    expect(run('stdDevp')[2]).toBe(0);
    expect(run('distinctCount')).toEqual(['Distinct Count of Sales', 3, 1, 3, 7]);
  });

  it('show values as', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 2, 'cols');
    pt = placeField(pt, c, 3, 'data');
    const shown = (patch: Partial<PivotTable['data'][number]>) => grid({ ...pt, data: [{ ...pt.data[0]!, ...patch }] }, c);
    const g = shown({ showAs: 'percentOfTotal' });
    expect(g[2]![1]).toBeCloseTo(30 / 112);
    expect(g[5]![3]).toBe(1);
    expect(shown({ showAs: 'percentOfCol' })[2]![1]).toBeCloseTo(30 / 35);
    expect(shown({ showAs: 'percentOfRow' })[2]![1]).toBeCloseTo(0.5);
    // Difference from 2023 along the Year field (index 0 in its items).
    const diff = shown({ showAs: 'difference', baseField: 2, baseItem: 0 });
    expect(diff[2]).toEqual(['East', null, 0, null]);
    expect(diff[4]).toEqual(['West', null, 35, null]);
    const pd = shown({ showAs: 'percentDiff', baseField: 2, baseItem: 'prev' });
    expect(pd[4]![2]).toBeCloseTo(7);
    const run = shown({ showAs: 'runTotal', baseField: 2 });
    expect(run[4]).toEqual(['West', 5, 45, null]);
    const rank = shown({ showAs: 'rankDescending', baseField: 0 });
    expect(rank.slice(2, 5).map((r) => r[3])).toEqual([1, 3, 2]);
    const idx = shown({ showAs: 'index' });
    expect(idx[2]![1]).toBeCloseTo((30 * 112) / (60 * 35));
  });

  it('filters: hidden items, label, value and top N', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    const hide = setField(pt, 0, { items: [{ x: 0 }, { x: 1, hidden: true }, { x: 2 }] });
    expect(grid(hide, c)).toEqual([['Row Labels', 'Sum of Sales'], ['East', 60], ['West', 45], ['Grand Total', 105]]);
    const lab = setField(pt, 0, { filters: [{ kind: 'label', op: 'beginsWith', value1: 'w' }] });
    expect(grid(lab, c)).toEqual([['Row Labels', 'Sum of Sales'], ['West', 45], ['Grand Total', 45]]);
    const val = setField(pt, 0, { filters: [{ kind: 'value', op: 'greaterThan', dataField: 0, value1: 40 }] });
    expect(grid(val, c).map((r) => r[0])).toEqual(['Row Labels', 'East', 'West', 'Grand Total']);
    const top = setField(pt, 0, { filters: [{ kind: 'top', top: true, by: 'count', value: 1, dataField: 0 }] });
    expect(grid(top, c)).toEqual([['Row Labels', 'Sum of Sales'], ['East', 60], ['Grand Total', 60]]);
  });

  it('report filter with one item', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 1, 'pages');
    pt = placeField(pt, c, 3, 'data');
    pt = { ...pt, pages: [{ field: 1, item: 1 }] };
    const l = computePivot(pt, c);
    expect(l.pageCells.map((p) => [p.dr, p.dc, p.cell.v])).toEqual([[-2, 0, 'Product'], [-2, 1, 'Pear']]);
    expect(l.cells.map((r) => r.map((x) => x?.v ?? null))).toEqual([['Row Labels', 'Sum of Pear'.replace('Pear', 'Sales')], ['East', 20], ['West', 40], ['Grand Total', 60]]);
  });

  it('sort by value descending and by label', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    expect(grid(setField(pt, 0, { sort: 'descending', sortByData: 0 }), c).map((r) => r[0])).toEqual(['Row Labels', 'East', 'West', 'North', 'Grand Total']);
    expect(grid(setField(pt, 0, { sort: 'descending' }), c).map((r) => r[0])).toEqual(['Row Labels', 'West', 'North', 'East', 'Grand Total']);
  });

  it('collapse hides details', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 1, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = setField(pt, 0, { items: [{ x: 0, collapsed: true }, { x: 1 }, { x: 2 }] });
    const l = computePivot(pt, c);
    expect(l.cells.map((r) => r[0]?.v)).toEqual(['Row Labels', 'East', 'North', 'Apple', 'West', 'Apple', 'Pear', 'Grand Total']);
    expect(l.cells[1]![0]!.button).toBe('collapsed');
    expect(l.cells[2]![0]!.button).toBe('expanded');
  });

  it('custom subtotal functions', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 1, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = setField(pt, 0, { subtotals: ['sum', 'max'] });
    const g = grid(pt, c);
    expect(g.slice(1, 6)).toEqual([['East', null], ['Apple', 40], ['Pear', 20], ['East Sum', 60], ['East Max', 30]]);
  });

  it('grouping dates and numbers, manual groups', () => {
    const header = ['Date', 'Amount'];
    const rows: Scalar[][] = [[45292, 10], [45383, 20], [45658, 30], [45700, 40]]; // 2024-01-01, 2024-04-01, 2025-01-01, 2025-02-12
    let c = buildCache(2, { kind: 'range', sheet: 'S', range: { r1: 0, c1: 0, r2: 4, c2: 1 } }, { header, rows, formats: ['m/d/yyyy', undefined] });
    const g = groupByRange(c, 0, [{ by: 'years' }, { by: 'quarters' }]);
    c = g.cache;
    let pt = newPivotTable('P', c, { row: 0, col: 0 });
    pt = placeField(pt, c, g.fields[1]!, 'rows');
    pt = placeField(pt, c, g.fields[0]!, 'rows');
    pt = placeField(pt, c, 1, 'data');
    expect(grid(pt, c).map((r) => r[0])).toEqual(['Row Labels', 2024, 'Qtr1', 'Qtr2', 2025, 'Qtr1', 'Grand Total']);
    expect(grid(pt, c).map((r) => r[1])).toEqual(['Sum of Amount', 30, 10, 20, 70, 70, 100]);

    let c2 = buildCache(3, { kind: 'range', sheet: 'S', range: { r1: 0, c1: 0, r2: 4, c2: 1 } }, { header: ['N', 'V'], rows: [[3, 1], [12, 1], [17, 1], [25, 1]], formats: [] });
    c2 = groupByRange(c2, 0, [{ by: 'range', start: 0, end: 29, interval: 10 }]).cache;
    let p2 = newPivotTable('P', c2, { row: 0, col: 0 });
    p2 = placeField(p2, c2, 0, 'rows');
    p2 = placeField(p2, c2, 1, 'data');
    expect(grid(p2, c2)).toEqual([['Row Labels', 'Sum of V'], ['0-9', 1], ['10-19', 2], ['20-29', 1], ['Grand Total', 4]]);

    const c3 = cache();
    const mg = groupItems(c3, 0, [{ name: 'Coast', items: [0, 2] }]);
    let p3 = newPivotTable('P', mg.cache, { row: 0, col: 0 });
    p3 = placeField(p3, mg.cache, mg.field, 'rows');
    p3 = placeField(p3, mg.cache, 3, 'data');
    expect(grid(p3, mg.cache)).toEqual([['Row Labels', 'Sum of Sales'], ['Coast', 105], ['North', 7], ['Grand Total', 112]]);
  });

  it('calculated fields and items', () => {
    const c = addCalcField(cache(), 'Price', 'Sales/Units');
    let pt = newPivotTable('P', c, { row: 0, col: 0 });
    pt = placeField(pt, c, 0, 'rows');
    pt = placeField(pt, c, 5, 'data');
    expect(grid(pt, c)).toEqual([['Row Labels', 'Sum of Price'], ['East', 10], ['North', 1], ['West', 45 / 11], ['Grand Total', 112 / 24]]);
    const c2 = cache();
    let p2 = newPivotTable('P', c2, { row: 0, col: 0 });
    p2 = placeField(p2, c2, 0, 'rows');
    p2 = placeField(p2, c2, 3, 'data');
    p2 = setField(p2, 0, { items: [...p2.fields[0]!.items, { x: -1, name: 'Both', formula: "East+'West'" }] });
    expect(grid(p2, c2)).toEqual([['Row Labels', 'Sum of Sales'], ['East', 60], ['North', 7], ['West', 45], ['Both', 105], ['Grand Total', 217]]);
  });

  it('remaps after a refresh', () => {
    const c = cache();
    let pt = placeField(base(c), c, 0, 'rows');
    pt = placeField(pt, c, 3, 'data');
    const r = remapTable(pt, [0, 1, 2, 3, 4], [[0, 1, 2]], c);
    expect(grid(r, c)).toEqual(grid(pt, c));
  });

  it('calculated items have inner items and count in totals for every function', () => {
    const c = cache();
    const run = (agg: PivotTable['data'][number]['agg']) => {
      let pt = placeField(base(c), c, 0, 'rows');
      pt = placeField(pt, c, 1, 'rows');
      pt = { ...pt, data: [makeDataField(pt, c, 3, agg)] };
      pt = setField(pt, 0, { items: [...pt.fields[0]!.items, { x: -1, name: 'Both', formula: 'East+West' }] });
      return grid(pt, c).map((r) => [r[0], r[1]]);
    };
    expect(run('sum')).toEqual([
      ['Row Labels', 'Sum of Sales'], ['East', 60], ['Apple', 40], ['Pear', 20], ['North', 7], ['Apple', 7], ['West', 45], ['Apple', 5], ['Pear', 40],
      ['Both', 105], ['Apple', 45], ['Pear', 60], ['Grand Total', 217],
    ]);
    const max = run('max');
    expect(max.slice(9)).toEqual([['Both', 55], ['Apple', 35], ['Pear', 45], ['Grand Total', 55]]);
    const avg = run('average');
    expect(avg.slice(9)).toEqual([['Both', 35], ['Apple', 25], ['Pear', 40], ['Grand Total', 147 / 8]]);
    expect(run('count').slice(9)).toEqual([['Both', 6], ['Apple', 3], ['Pear', 3], ['Grand Total', 13]]);
    expect(run('min').at(-1)).toEqual(['Grand Total', 5]);
    expect(run('product').at(-1)![1]).toBe(10 * 20 * 30 * 5 * 15 * 25 * 7 * (6000 + 1875));
    // Moved among the other items, and placed as a column item too.
    let pt = placeField(base(c), c, 0, 'cols');
    pt = placeField(pt, c, 2, 'rows');
    pt = placeField(pt, c, 3, 'data');
    pt = setField(pt, 0, { items: [{ x: 0 }, { x: -1, name: 'EW', formula: 'East+West' }, { x: 1 }, { x: 2 }] });
    expect(grid(pt, c).slice(1, 4)).toEqual([['Row Labels', 'East', 'EW', 'North', 'West', 'Grand Total'], [2023, 30, 35, null, 5, 70], [2024, 30, 70, 7, 40, 147]]);
  });
});
