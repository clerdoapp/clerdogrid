import { beforeAll, describe, expect, it } from 'vitest';
import { Engine } from '../src/engine/engine';
import { registerAllFunctions } from '../src/engine/functions';
import { CellError, type Scalar } from '../src/engine/types';
import { newWorkbook } from '../src/model/workbook';

beforeAll(() => registerAllFunctions());

// Region | Item | Year | Sales | Units
const DATA: Scalar[][] = [
  ['Region', 'Item', 'Year', 'Sales', 'Units'],
  ['East', 'a', 2023, 10, 1], ['West', 'a', 2023, 20, 2], ['East', 'b', 2024, 30, 3], ['West', 'b', 2024, 40, 4], ['East', 'a', 2024, 5, 5], ['North', 'b', 2023, 7, 6],
];

const run = (formula: string): Scalar[][] => {
  const wb = newWorkbook();
  const e = new Engine(wb);
  DATA.forEach((r, i) => r.forEach((v, j) => e.setValue(0, i, j, v)));
  e.setFormula(0, 0, 10, formula);
  e.recalc();
  const a = e.spillAnchorOf(0, 0, 10);
  const first = e.getValue(0, 0, 10);
  if (!a) return [[first instanceof CellError ? first.code : first]];
  const out: Scalar[][] = [];
  for (let r = 0; r < 40; r++) {
    const row: Scalar[] = [];
    for (let c = 10; c < 40; c++) {
      if (!e.isSpilled(0, r, c) && !(r === 0 && c === 10)) break;
      const v = e.getValue(0, r, c);
      row.push(v instanceof CellError ? v.code : v === '' ? null : v);
    }
    if (!row.length) break;
    out.push(row);
  }
  return out;
};

describe('GROUPBY', () => {
  it('sums by one field with totals', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,SUM)')).toEqual([['East', 45], ['North', 7], ['West', 60], ['Total', 112]]);
  });
  it('shows headers and takes several value columns', () => {
    expect(run('GROUPBY(A1:A7,D1:E7,SUM,3)')).toEqual([['Region', 'Sales', 'Units'], ['East', 45, 9], ['North', 7, 6], ['West', 60, 6], ['Total', 112, 21]]);
  });
  it('takes several functions with HSTACK', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,HSTACK(SUM,AVERAGE,COUNT),0,0)')).toEqual([[null, 'SUM', 'AVERAGE', 'COUNT'], ['East', 45, 15, 3], ['North', 7, 7, 1], ['West', 60, 30, 2]]);
    expect(run('GROUPBY(A1:A7,D1:E7,HSTACK(SUM,MAX),3,0)')).toEqual([
      [null, 'SUM', 'SUM', 'MAX', 'MAX'], ['Region', 'Sales', 'Units', 'Sales', 'Units'], ['East', 45, 9, 30, 5], ['North', 7, 6, 7, 6], ['West', 60, 6, 40, 4],
    ]);
  });
  it('takes functions stacked vertically', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,VSTACK(SUM,MIN),0,0)')).toEqual([['East', 'SUM', 45], [null, 'MIN', 5], ['North', 'SUM', 7], [null, 'MIN', 7], ['West', 'SUM', 60], [null, 'MIN', 20]]);
  });
  it('subtotals each level, top or bottom', () => {
    expect(run('GROUPBY(A2:B7,D2:D7,SUM,0,2)')).toEqual([
      ['East', 'a', 15], ['East', 'b', 30], ['East', null, 45], ['North', 'b', 7], ['North', null, 7], ['West', 'a', 20], ['West', 'b', 40], ['West', null, 60], ['Total', null, 112],
    ]);
    expect(run('GROUPBY(A2:B7,D2:D7,SUM,0,-2)').slice(0, 4)).toEqual([['Total', null, 112], ['East', null, 45], ['East', 'a', 15], ['East', 'b', 30]]);
    expect(run('GROUPBY(A2:B7,D2:D7,SUM,0,1)').at(-1)).toEqual(['Total', null, 112]);
  });
  it('sorts by key and by value columns', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,-1)').map((r) => r[0])).toEqual(['West', 'North', 'East']);
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,2)').map((r) => r[0])).toEqual(['North', 'East', 'West']);
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,-2)').map((r) => r[1])).toEqual([60, 45, 7]);
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,4)')).toEqual(['#VALUE!'].map((x) => [x]));
  });
  it('filters rows', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,,C2:C7=2024)')).toEqual([['East', 35], ['West', 40]]);
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,0,0,,C2:C7=1999)')).toEqual([['#CALC!']]);
  });
  it('groups whole rows in table mode', () => {
    expect(run('GROUPBY(A2:B7,D2:D7,SUM,0,1,,,1)')).toEqual([['East', 'a', 15], ['East', 'b', 30], ['North', 'b', 7], ['West', 'a', 20], ['West', 'b', 40], ['Total', null, 112]]);
    expect(run('GROUPBY(A2:B7,D2:D7,SUM,0,2,,,1)')).toEqual([['#VALUE!']]);
  });
  it('PERCENTOF and LAMBDA aggregations', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,PERCENTOF,0,0)').map((r) => r[1])).toEqual([45 / 112, 7 / 112, 60 / 112]);
    expect(run('GROUPBY(A2:A7,D2:D7,LAMBDA(x,MAX(x)-MIN(x)),0,0)').map((r) => r[1])).toEqual([25, 0, 20]);
  });
  it('generates headers with field_headers 2', () => {
    expect(run('GROUPBY(A2:A7,D2:D7,SUM,2,0)')[0]).toEqual(['Row Field 1', 'Value 1']);
  });
});

describe('PIVOTBY', () => {
  it('cross-tabulates with totals', () => {
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,SUM)')).toEqual([
      [null, 2023, 2024, 'Total'], ['East', 10, 35, 45], ['North', 7, null, 7], ['West', 20, 40, 60], ['Total', 37, 75, 112],
    ]);
  });
  it('controls row and column totals and sorting', () => {
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,SUM,0,0,-1,0)')).toEqual([[null, 2023, 2024], ['West', 20, 40], ['North', 7, null], ['East', 10, 35]]);
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,SUM,0,-1,,0,-1)')).toEqual([[null, 2024, 2023], ['Total', 75, 37], ['East', 35, 10], ['North', null, 7], ['West', 40, 20]]);
  });
  it('several values and functions', () => {
    expect(run('PIVOTBY(A1:A7,C1:C7,D1:E7,SUM,3,0,,0)')).toEqual([
      [null, 2023, 2023, 2024, 2024], ['Region', 'Sales', 'Units', 'Sales', 'Units'], ['East', 10, 1, 35, 8], ['North', 7, 6, null, null], ['West', 20, 2, 40, 4],
    ]);
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,HSTACK(SUM,COUNT),0,0,,0)')).toEqual([
      [null, 2023, 2023, 2024, 2024], [null, 'SUM', 'COUNT', 'SUM', 'COUNT'], ['East', 10, 1, 35, 2], ['North', 7, 1, null, null], ['West', 20, 1, 40, 1],
    ]);
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,VSTACK(SUM,COUNT),0,0,,0)')).toEqual([
      [null, null, 2023, 2024], ['East', 'SUM', 10, 35], [null, 'COUNT', 1, 2], ['North', 'SUM', 7, null], [null, 'COUNT', 1, null], ['West', 'SUM', 20, 40], [null, 'COUNT', 1, 1],
    ]);
  });
  it('PERCENTOF relative to column, row, grand and parent totals', () => {
    const f = (rel: number) => run(`PIVOTBY(A2:B7,C2:C7,D2:D7,PERCENTOF,0,2,,1,,,${rel})`);
    // Rows: East(a,b,sub) North(b,sub) West(a,b,sub) Total; columns 2023, 2024, Total.
    const col = f(0);
    expect(col[1]).toEqual(['East', 'a', 10 / 37, 5 / 75, 15 / 112]);
    const row = f(1);
    expect(row[1]).toEqual(['East', 'a', 10 / 15, 5 / 15, 1]);
    const grand = f(2);
    expect(grand[1]).toEqual(['East', 'a', 10 / 112, 5 / 112, 15 / 112]);
    // Parent row total: East a in 2023 over East in 2023.
    const parentRow = f(4);
    expect(parentRow[1]).toEqual(['East', 'a', 1, 5 / 35, 15 / 45]);
    expect(parentRow[3]).toEqual(['East', null, 10 / 37, 35 / 75, 45 / 112]);
    // Parent column total with one column level: the row's own total.
    expect(f(3)[1]).toEqual(['East', 'a', 10 / 15, 5 / 15, 1]);
  });
  it('filters and reports bad arguments', () => {
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,SUM,0,0,,0,,B2:B7="b")')).toEqual([[null, 2023, 2024], ['East', null, 30], ['North', 7, null], ['West', null, 40]]);
    expect(run('PIVOTBY(A2:A7,C2:C6,D2:D7,SUM)')).toEqual([['#VALUE!']]);
    expect(run('PIVOTBY(A2:A7,C2:C7,D2:D7,SUM,0,0,,0,,,9)')).toEqual([['#VALUE!']]);
  });
});
