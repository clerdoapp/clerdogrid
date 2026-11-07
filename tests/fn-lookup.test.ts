import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/lookup';
import { Err, MAX_ROWS, ref, type FnContext, type Scalar, type Value } from '../src/engine/types';
import { lambda, makeCtx } from './fnctx';

const call = (name: string, args: Value[], ctx: FnContext = makeCtx()): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const sumOf = (...a: Value[]): Value => (a[0] as Scalar[][]).flat().reduce<number>((s, x) => s + (typeof x === 'number' ? x : 0), 0);
const SUM = lambda((x) => sumOf(x));

describe('ADDRESS / ROW / COLUMN / AREAS / ROWS', () => {
  it('formats A1 and R1C1 addresses', () => {
    expect(call('ADDRESS', [2, 3])).toBe('$C$2');
    expect(call('ADDRESS', [2, 3, 2])).toBe('C$2');
    expect(call('ADDRESS', [2, 3, 2, false])).toBe('R2C[3]');
    expect(call('ADDRESS', [2, 3, 1, false, '[Book1]Sheet1'])).toBe("'[Book1]Sheet1'!R2C3");
    expect(call('ADDRESS', [2, 3, 1, false, 'DATA SHEET'])).toBe("'DATA SHEET'!R2C3");
    expect(call('ADDRESS', [1, 28, 4])).toBe('AB1');
    expect(call('ADDRESS', [1, 16384])).toBe('$XFD$1');
    expect(call('ADDRESS', [0, 1])).toEqual(Err.value);
    expect(call('ADDRESS', [1, 1, 1, true, 'Data'])).toBe('Data!$A$1');
  });
  it('ROW and COLUMN', () => {
    const ctx = { ...makeCtx(), row: 4, col: 2 };
    expect(call('ROW', [], ctx)).toBe(5);
    expect(call('COLUMN', [], ctx)).toBe(3);
    expect(call('ROW', [ref(0, 2, 0, 4, 0)])).toEqual([[3], [4], [5]]);
    expect(call('COLUMN', [ref(0, 0, 1, 0, 3)])).toEqual([[2, 3, 4]]);
    expect(call('ROWS', [[[1, 2], [3, 4], [5, 6]]])).toBe(3);
    expect(call('COLUMNS', [ref(0, 0, 0, MAX_ROWS - 1, 3)])).toBe(4);
    expect(call('AREAS', [{ kind: 'refs', areas: [ref(0, 0, 0), ref(0, 1, 1)] }])).toBe(2);
  });
});

describe('INDEX / OFFSET', () => {
  const grid: Scalar[][] = [['Apples', 'Lemons'], ['Bananas', 'Pears']];
  const ctx = makeCtx([grid]);
  it('array form', () => {
    expect(call('INDEX', [grid, 2, 2])).toBe('Pears');
    expect(call('INDEX', [grid, 2, 1])).toBe('Bananas');
    expect(call('INDEX', [[[1, 2, 3]], 2])).toBe(2);
    expect(call('INDEX', [grid, 0, 2])).toEqual([['Lemons'], ['Pears']]);
    expect(call('INDEX', [grid, 3, 1])).toEqual(Err.ref);
  });
  it('reference form returns refs', () => {
    expect(call('INDEX', [ref(0, 0, 0, 1, 1), 2, 2], ctx)).toEqual(ref(0, 1, 1));
    expect(call('INDEX', [ref(0, 0, 0, 9, 0), 5], ctx)).toEqual(ref(0, 4, 0));
    expect(call('INDEX', [ref(0, 0, 0, 1, 1), 0, 1], ctx)).toEqual(ref(0, 0, 0, 1, 0));
    const u = { kind: 'refs' as const, areas: [ref(0, 0, 0, 1, 1), ref(0, 5, 5, 7, 7)] };
    expect(call('INDEX', [u, 2, 3, 2], ctx)).toEqual(ref(0, 6, 7));
    expect(call('INDEX', [u, 1, 1, 3], ctx)).toEqual(Err.ref);
  });
  it('OFFSET', () => {
    expect(call('OFFSET', [ref(0, 2, 2), 3, -2, 3, 3], ctx)).toEqual(ref(0, 5, 0, 7, 2));
    expect(call('OFFSET', [ref(0, 2, 2), 3, -3], ctx)).toEqual(Err.ref);
    expect(call('OFFSET', [ref(0, 0, 0, 1, 1), 1, 1], ctx)).toEqual(ref(0, 1, 1, 2, 2));
    expect(call('OFFSET', [ref(0, 0, 0), 0, 0, 0, 1], ctx)).toEqual(Err.ref);
  });
});

describe('MATCH / VLOOKUP / HLOOKUP / LOOKUP', () => {
  const v = [[25], [38], [40], [41]];
  it('MATCH types', () => {
    expect(call('MATCH', [39, v, 1])).toBe(2);
    expect(call('MATCH', [41, v, 0])).toBe(4);
    expect(call('MATCH', [40, v, -1])).toEqual(Err.na);
    expect(call('MATCH', [40, [[41], [40], [38], [25]], -1])).toBe(2);
    expect(call('MATCH', [10, v])).toEqual(Err.na);
    expect(call('MATCH', ['b*', [['apple'], ['Banana'], ['cherry']], 0])).toBe(2);
    expect(call('MATCH', ['a~*', [['ab'], ['a*'], ['a']], 0])).toBe(2);
    expect(call('MATCH', ['?', [['ab'], ['c']], 0])).toBe(2);
  });
  it('MATCH approximate on unsorted data follows binary search', () => {
    // A linear scan would stop at 9 and return 1; spreadsheet's bisection skips it and returns 5.
    expect(call('MATCH', [5, [[1], [9], [2], [3], [4]], 1])).toBe(5);
    expect(call('VLOOKUP', [5, [[1, 'a'], [9, 'b'], [2, 'c'], [3, 'd'], [4, 'e']], 2])).toBe('e');
  });
  it('VLOOKUP exact and approximate', () => {
    const t: Scalar[][] = [[0.457, 3.55, 500], [0.525, 3.25, 400], [0.606, 2.93, 300], [0.675, 2.75, 250], [0.746, 2.57, 200]];
    expect(call('VLOOKUP', [1, t, 2])).toBe(2.57);
    expect(call('VLOOKUP', [0.7, t, 3, false])).toEqual(Err.na);
    expect(call('VLOOKUP', [0.1, t, 2, true])).toEqual(Err.na);
    expect(call('VLOOKUP', [2, t, 2, true])).toBe(2.57);
    expect(call('VLOOKUP', [0.525, t, 4])).toEqual(Err.ref);
    expect(call('VLOOKUP', [0.525, t, 0])).toEqual(Err.value);
    const names: Scalar[][] = [['Davolio', 'Nancy'], ['Fuller', 'Andrew']];
    expect(call('VLOOKUP', ['ful*', names, 2, false])).toBe('Andrew');
    expect(call('VLOOKUP', ['FULLER', names, 2, false])).toBe('Andrew');
  });
  it('HLOOKUP', () => {
    const t: Scalar[][] = [['Axles', 'Bearings', 'Bolts'], [4, 4, 9], [5, 7, 10], [6, 8, 11]];
    expect(call('HLOOKUP', ['Axles', t, 2, true])).toBe(4);
    expect(call('HLOOKUP', ['Bearings', t, 3, false])).toBe(7);
    expect(call('HLOOKUP', ['B', t, 3, true])).toBe(5);
    expect(call('HLOOKUP', ['Bolts', t, 5])).toEqual(Err.ref);
  });
  it('LOOKUP vector and array forms', () => {
    const lv = [[4.14], [4.19], [5.17], [5.77], [6.39]], rv = [['red'], ['orange'], ['yellow'], ['green'], ['blue']];
    expect(call('LOOKUP', [4.19, lv, rv])).toBe('orange');
    expect(call('LOOKUP', [5.75, lv, rv])).toBe('yellow');
    expect(call('LOOKUP', [7.66, lv, rv])).toBe('blue');
    expect(call('LOOKUP', [0, lv, rv])).toEqual(Err.na);
    expect(call('LOOKUP', ['C', [['a', 'b', 'c', 'd'], [1, 2, 3, 4]]])).toBe(3);
    expect(call('LOOKUP', ['bump', [['a', 1], ['b', 2], ['c', 3]]])).toBe(2);
  });
});

describe('XLOOKUP / XMATCH', () => {
  const countries = [['China'], ['India'], ['United States'], ['Brazil']];
  const codes = [['+86'], ['+91'], ['+1'], ['+55']];
  it('exact, not found, if_not_found', () => {
    expect(call('XLOOKUP', ['Brazil', countries, codes])).toBe('+55');
    expect(call('XLOOKUP', ['Peru', countries, codes])).toEqual(Err.na);
    expect(call('XLOOKUP', ['Peru', countries, codes, 'Not found'])).toBe('Not found');
    expect(call('XLOOKUP', ['x', countries, [[1], [2]]])).toEqual(Err.value);
  });
  it('match modes and search modes', () => {
    const v = [[10], [20], [30], [40]], r = [['a'], ['b'], ['c'], ['d']];
    expect(call('XLOOKUP', [25, v, r, null, -1])).toBe('b');
    expect(call('XLOOKUP', [25, v, r, null, 1])).toBe('c');
    expect(call('XLOOKUP', [45, v, r, 'none', 1])).toBe('none');
    expect(call('XLOOKUP', [25, v, r, null, 1, 2])).toBe('c');
    expect(call('XLOOKUP', [25, v, r, null, -1, 2])).toBe('b');
    expect(call('XLOOKUP', [25, [[40], [30], [20], [10]], r, null, 1, -2])).toBe('b');
    expect(call('XLOOKUP', [25, [[40], [30], [20], [10]], r, null, -1, -2])).toBe('c');
    expect(call('XLOOKUP', ['B', [['a'], ['b'], ['b']], [[1], [2], [3]], null, 0, -1])).toBe(3);
    expect(call('XLOOKUP', ['b?', [['a'], ['bc']], [[1], [2]], null, 2])).toBe(2);
    expect(call('XLOOKUP', ['b?', [['a'], ['bc']], [[1], [2]]])).toEqual(Err.na);
  });
  it('returns a whole row and a reference when given a range', () => {
    expect(call('XLOOKUP', [2, [[1], [2]], [[1, 'x', 'y'], [2, 'p', 'q']]])).toEqual([[2, 'p', 'q']]);
    const ctx = makeCtx([[[1, 'a'], [2, 'b']]]);
    expect(call('XLOOKUP', [2, ref(0, 0, 0, 1, 0), ref(0, 0, 1, 1, 1)], ctx)).toEqual(ref(0, 1, 1));
  });
  it('XMATCH', () => {
    expect(call('XMATCH', ['Gra?e', [['Apple'], ['Grape'], ['Pear']], 2])).toBe(2);
    expect(call('XMATCH', [4.5, [[1], [3], [5], [7]], 1])).toBe(3);
    expect(call('XMATCH', [4.5, [[1], [3], [5], [7]], -1])).toBe(2);
    expect(call('XMATCH', [9, [[1], [3]]])).toEqual(Err.na);
  });
});

describe('array shaping', () => {
  const m: Scalar[][] = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
  it('TAKE / DROP / CHOOSE*', () => {
    expect(call('TAKE', [m, 2])).toEqual([[1, 2, 3], [4, 5, 6]]);
    expect(call('TAKE', [m, -1, -2])).toEqual([[8, 9]]);
    expect(call('TAKE', [m, null, 1])).toEqual([[1], [4], [7]]);
    expect(call('DROP', [m, 1, 1])).toEqual([[5, 6], [8, 9]]);
    expect(call('DROP', [m, 3])).toEqual(Err.calc);
    expect(call('CHOOSECOLS', [m, 1, -1])).toEqual([[1, 3], [4, 6], [7, 9]]);
    expect(call('CHOOSEROWS', [m, [[3, 1]]])).toEqual([[7, 8, 9], [1, 2, 3]]);
    expect(call('CHOOSEROWS', [m, 4])).toEqual(Err.value);
  });
  it('EXPAND / STACKs / WRAP / TOCOL', () => {
    expect(call('EXPAND', [[[1, 2]], 2, 3])).toEqual([[1, 2, Err.na], [Err.na, Err.na, Err.na]]);
    expect(call('EXPAND', [[[1, 2]], 2, null, 0])).toEqual([[1, 2], [0, 0]]);
    expect(call('EXPAND', [m, 2])).toEqual(Err.value);
    expect(call('VSTACK', [[[1, 2]], [[3]]])).toEqual([[1, 2], [3, Err.na]]);
    expect(call('HSTACK', [[[1], [2]], [[3]]])).toEqual([[1, 3], [2, Err.na]]);
    expect(call('WRAPROWS', [[[1, 2, 3, 4, 5]], 2])).toEqual([[1, 2], [3, 4], [5, Err.na]]);
    expect(call('WRAPCOLS', [[[1, 2, 3]], 2, 0])).toEqual([[1, 3], [2, 0]]);
    expect(call('TOCOL', [[[1, null], [Err.div0, 4]], 3])).toEqual([[1], [4]]);
    expect(call('TOROW', [m, 0, true])).toEqual([[1, 4, 7, 2, 5, 8, 3, 6, 9]]);
    expect(call('TOCOL', [m, 4])).toEqual(Err.value);
    expect(call('TRANSPOSE', [[[1, 2]]])).toEqual([[1], [2]]);
  });
  it('FILTER', () => {
    const data: Scalar[][] = [['Apple', 'East', 10], ['Pear', 'West', 20], ['Fig', 'East', 30]];
    expect(call('FILTER', [data, [[true], [false], [true]]])).toEqual([['Apple', 'East', 10], ['Fig', 'East', 30]]);
    expect(call('FILTER', [data, [[0], [0], [0]]])).toEqual(Err.calc);
    expect(call('FILTER', [data, [[0], [0], [0]], 'none'])).toBe('none');
    expect(call('FILTER', [data, [[1, 0, 1]]])).toEqual([['Apple', 10], ['Pear', 20], ['Fig', 30]]);
    expect(call('FILTER', [data, [[1], [0]]])).toEqual(Err.value);
  });
  it('SORT / SORTBY / UNIQUE', () => {
    const d: Scalar[][] = [['b', 2], ['a', 3], ['c', 1]];
    expect(call('SORT', [d])).toEqual([['a', 3], ['b', 2], ['c', 1]]);
    expect(call('SORT', [d, 2, -1])).toEqual([['a', 3], ['b', 2], ['c', 1]]);
    expect(call('SORT', [[[3, 1, 2]], 1, 1, true])).toEqual([[1, 2, 3]]);
    expect(call('SORT', [[[null], [2], ['x'], [true], [1]]])).toEqual([[1], [2], ['x'], [true], [null]]);
    expect(call('SORT', [d, 3])).toEqual(Err.value);
    const names = [['Tom'], ['Fred'], ['Amy'], ['Sal'], ['Hector']];
    const region = [['West'], ['East'], ['West'], ['East'], ['West']];
    const age = [[39], [52], [22], [40], [22]];
    // Reference example: region ascending, then age descending.
    expect(call('SORTBY', [names, region, 1, age, -1])).toEqual([['Fred'], ['Sal'], ['Tom'], ['Amy'], ['Hector']]);
    expect(call('SORTBY', [names, age])).toEqual([['Amy'], ['Hector'], ['Tom'], ['Sal'], ['Fred']]);
    expect(call('SORTBY', [names, [[1], [2]]])).toEqual(Err.value);
    expect(call('UNIQUE', [[['a'], ['A'], ['b'], ['a']]])).toEqual([['a'], ['b']]);
    expect(call('UNIQUE', [[['a'], ['b'], ['a']], false, true])).toEqual([['b']]);
    expect(call('UNIQUE', [[[1, 2, 1]], true])).toEqual([[1, 2]]);
    expect(call('UNIQUE', [[['a'], ['a']], false, true])).toEqual(Err.calc);
  });
  it('SEQUENCE / RANDARRAY / TRIMRANGE', () => {
    expect(call('SEQUENCE', [2, 3, 3, 2])).toEqual([[3, 5, 7], [9, 11, 13]]);
    expect(call('SEQUENCE', [0])).toEqual(Err.calc);
    const r = call('RANDARRAY', [3, 2, 1, 6, true]) as number[][];
    expect(r).toHaveLength(3);
    expect(r.flat().every((x) => Number.isInteger(x) && x >= 1 && x <= 6)).toBe(true);
    expect(call('RANDARRAY', [1, 1, 5, 1])).toEqual(Err.value);
    const ctx = makeCtx([[[null, null], [null, 5], [null, 6]]]);
    expect(call('TRIMRANGE', [ref(0, 0, 0, 99, 9)], ctx)).toEqual(ref(0, 1, 1, 2, 1));
    expect(call('TRIMRANGE', [ref(0, 0, 0, 99, 9), 2, 0], ctx)).toEqual(ref(0, 0, 0, 2, 9));
  });
});

describe('LAMBDA helpers', () => {
  const m: Scalar[][] = [[1, 2], [3, 4]];
  it('BYROW / BYCOL / MAP / MAKEARRAY', () => {
    expect(call('BYROW', [m, SUM])).toEqual([[3], [7]]);
    expect(call('BYCOL', [m, SUM])).toEqual([[4, 6]]);
    expect(call('BYROW', [m, lambda((x) => x)])).toEqual([[Err.calc], [Err.calc]]);
    expect(call('MAP', [m, lambda((x) => (x as number) * 10)])).toEqual([[10, 20], [30, 40]]);
    expect(call('MAP', [m, [[1, 1], [0, 0]], lambda((x, y) => (y ? x : 0), ['a', 'b'])])).toEqual([[1, 2], [0, 0]]);
    expect(call('MAKEARRAY', [2, 3, lambda((r, c) => (r as number) * (c as number), ['r', 'c'])])).toEqual([[1, 2, 3], [2, 4, 6]]);
    expect(call('MAKEARRAY', [0, 3, lambda(() => 1, ['r', 'c'])])).toEqual(Err.value);
  });
  it('REDUCE / SCAN', () => {
    const add = lambda((a, b) => (a as number) + (b as number), ['a', 'b']);
    expect(call('REDUCE', [0, m, add])).toBe(10);
    expect(call('SCAN', [0, m, add])).toEqual([[1, 3], [6, 10]]);
    expect(call('REDUCE', [m, add])).toBe(10);
    expect(call('REDUCE', [0, m, lambda((a) => a)])).toEqual(Err.value);
  });
});

describe('GROUPBY / PIVOTBY / PERCENTOF / misc', () => {
  const region = [['East'], ['West'], ['East'], ['West'], ['North']];
  const sales = [[10], [20], [30], [40], [5]];
  it('groups and totals', () => {
    expect(call('GROUPBY', [region, sales, SUM])).toEqual([['East', 40], ['North', 5], ['West', 60], ['Total', 105]]);
    expect(call('GROUPBY', [region, sales, SUM, 0, 0, -2])).toEqual([['West', 60], ['East', 40], ['North', 5]]);
    expect(call('GROUPBY', [region, sales, SUM, 0, -1])).toEqual([['Total', 105], ['East', 40], ['North', 5], ['West', 60]]);
    expect(call('GROUPBY', [region, sales, SUM, 0, 0, null, [[true], [true], [false], [false], [false]]])).toEqual([['East', 10], ['West', 20]]);
  });
  it('headers and multi-level subtotals', () => {
    const rf = [['Region', 'Item'], ['East', 'a'], ['East', 'b'], ['West', 'a']];
    const v = [['Sales'], [1], [2], [4]];
    expect(call('GROUPBY', [rf, v, SUM, 3])).toEqual([
      ['Region', 'Item', 'Sales'], ['East', 'a', 1], ['East', 'b', 2], ['East', null, 3], ['West', 'a', 4], ['West', null, 4], ['Total', null, 7],
    ]);
  });
  it('PERCENTOF as a two-argument aggregator', () => {
    const PCT = lambda((s, t) => call('PERCENTOF', [s!, t!]), ['s', 't']);
    expect(call('GROUPBY', [[['a'], ['b'], ['a']], [[1], [2], [1]], PCT, 0, 0])).toEqual([['a', 0.5], ['b', 0.5]]);
    expect(call('PERCENTOF', [[[1], [3]], [[1], [3], [4]]])).toBe(0.5);
    expect(call('PERCENTOF', [1, 0])).toEqual(Err.div0);
  });
  it('PIVOTBY', () => {
    const rows = [['A'], ['A'], ['B']], cols = [['x'], ['y'], ['x']], v = [[1], [2], [3]];
    expect(call('PIVOTBY', [rows, cols, v, SUM, 0])).toEqual([
      [null, 'x', 'y', 'Total'], ['A', 1, 2, 3], ['B', 3, null, 3], ['Total', 4, 2, 6],
    ]);
  });
  it('HYPERLINK and GETPIVOTDATA', () => {
    expect(call('HYPERLINK', ['https://clerdo.app', 'Clerdo'])).toBe('Clerdo');
    expect(call('HYPERLINK', ['https://clerdo.app'])).toBe('https://clerdo.app');
    expect(call('GETPIVOTDATA', ['Sales', ref(0, 0, 0)])).toEqual(Err.ref);
  });
});
