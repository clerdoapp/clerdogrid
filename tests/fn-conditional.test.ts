import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/conditional';
import { makeCriteria } from '../src/engine/functions/criteria';
import { Err, MAX_ROWS, ref, type FnContext, type Scalar, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const call = (name: string, args: Value[], ctx: FnContext): Value => defs.find((d) => d.name === name)!.fn(args, ctx);

describe('makeCriteria', () => {
  const t = (c: Scalar, v: Scalar): boolean => makeCriteria(c)(v);
  it('numbers and operators', () => {
    expect(t(5, 5)).toBe(true);
    expect(t(5, '5')).toBe(true);
    expect(t('>5', 6)).toBe(true);
    expect(t('>5', '6')).toBe(false);
    expect(t('>=5', 5)).toBe(true);
    expect(t('<>5', 'x')).toBe(true);
    expect(t('<>5', 5)).toBe(false);
    expect(t('<5', null)).toBe(false);
    expect(t('=5', 5)).toBe(true);
  });
  it('text, wildcards and escapes', () => {
    expect(t('apple', 'APPLE')).toBe(true);
    expect(t('ap*', 'Apples')).toBe(true);
    expect(t('?pple', 'apple')).toBe(true);
    expect(t('a~*', 'a*')).toBe(true);
    expect(t('a~*', 'ab')).toBe(false);
    expect(t('a~?', 'a?')).toBe(true);
    expect(t('<>ap*', 'apple')).toBe(false);
    expect(t('<>ap*', 5)).toBe(true);
    expect(t('>b', 'c')).toBe(true);
    expect(t('>b', 5)).toBe(false);
    expect(t('*', 'x')).toBe(true);
    expect(t('*', 5)).toBe(false);
  });
  it('blanks', () => {
    expect(t('', null)).toBe(true);
    expect(t('', '')).toBe(true);
    expect(t('=', null)).toBe(true);
    expect(t('=', '')).toBe(false);
    expect(t('<>', null)).toBe(false);
    expect(t('<>', '')).toBe(true);
    expect(t(null, 0)).toBe(true);
    expect(t(null, null)).toBe(false);
  });
  it('dates, booleans, errors', () => {
    expect(t('>1/1/2024', 45293)).toBe(true);
    expect(t('>1/1/2024', 45292)).toBe(false);
    expect(t(true, true)).toBe(true);
    expect(t('TRUE', true)).toBe(true);
    expect(t(true, 1)).toBe(false);
    expect(t('#N/A', Err.na)).toBe(true);
  });
});

// The reference COUNTIF example: A2:B5
const fruit: Scalar[][] = [['apples', 32], ['oranges', 54], ['peaches', 75], ['apples', 86]];
const ctxF = makeCtx([fruit]);
const A = ref(0, 0, 0, 3, 0), B = ref(0, 0, 1, 3, 1);

describe('COUNTIF / COUNTIFS', () => {
  it('reference examples', () => {
    expect(call('COUNTIF', [A, 'apples'], ctxF)).toBe(2);
    expect(call('COUNTIF', [B, '>55'], ctxF)).toBe(2);
    expect(call('COUNTIF', [B, '<>75'], ctxF)).toBe(3);
    expect(call('COUNTIF', [A, '*es'], ctxF)).toBe(4);
    expect(call('COUNTIF', [A, '?????es'], ctxF)).toBe(2);
    expect(call('COUNTIF', [A, '*'], ctxF)).toBe(4);
  });
  it('blank cells, including beyond the used area', () => {
    const ctx = makeCtx([[[1], [null], [''], ['x']]]);
    expect(call('COUNTIF', [ref(0, 0, 0, 9, 0), ''], ctx)).toBe(8);
    expect(call('COUNTIF', [ref(0, 0, 0, 9, 0), '='], ctx)).toBe(7);
    expect(call('COUNTIF', [ref(0, 0, 0, 9, 0), '<>'], ctx)).toBe(3);
    expect(call('COUNTIF', [ref(0, 0, 0, MAX_ROWS - 1, 0), '<>'], ctx)).toBe(3);
  });
  it('COUNTIFS', () => {
    const g: Scalar[][] = [['Davidoski', 'Yes', 'No', 'No'], ['Burke', 'Yes', 'Yes', 'No'], ['Sundaram', 'Yes', 'Yes', 'Yes'], ['Levitan', 'No', 'Yes', 'Yes']];
    const ctx = makeCtx([g]);
    expect(call('COUNTIFS', [ref(0, 0, 1, 0, 3), '=Yes'], ctx)).toBe(1);
    expect(call('COUNTIFS', [ref(0, 0, 1, 3, 1), '=Yes', ref(0, 0, 2, 3, 2), '=Yes'], ctx)).toBe(2);
    expect(call('COUNTIFS', [ref(0, 0, 1, 3, 1), 'Yes', ref(0, 0, 2, 2, 2), 'Yes'], ctx)).toEqual(Err.value);
  });
});

describe('SUMIF / SUMIFS / AVERAGEIF(S) / MAXIFS / MINIFS', () => {
  const prop: Scalar[][] = [[100000, 7000, 250000], [200000, 14000], [300000, 21000], [400000, 28000]];
  const ctx = makeCtx([prop]);
  const P = ref(0, 0, 0, 3, 0), C = ref(0, 0, 1, 3, 1);
  it('SUMIF reference examples', () => {
    expect(call('SUMIF', [P, '>160000', C], ctx)).toBe(63000);
    expect(call('SUMIF', [P, '>160000'], ctx)).toBe(900000);
    expect(call('SUMIF', [P, 300000, C], ctx)).toBe(21000);
    expect(call('SUMIF', [P, '>' + 250000, C], ctx)).toBe(49000);
  });
  it('SUMIF resizes sum_range to the criteria shape', () => {
    expect(call('SUMIF', [P, '>160000', ref(0, 0, 1)], ctx)).toBe(63000);
  });
  it('SUMIFS and size checks', () => {
    const g: Scalar[][] = [[5, 'Apples', 'Tom'], [4, 'Apples', 'Sarah'], [15, 'Artichokes', 'Tom'], [3, 'Artichokes', 'Sarah'], [22, 'Bananas', 'Tom'], [12, 'Bananas', 'Sarah'], [10, 'Carrots', 'Tom'], [33, 'Carrots', 'Sarah']];
    const c2 = makeCtx([g]);
    const Q = ref(0, 0, 0, 7, 0), Pr = ref(0, 0, 1, 7, 1), S = ref(0, 0, 2, 7, 2);
    expect(call('SUMIFS', [Q, Pr, '=A*', S, 'Tom'], c2)).toBe(20);
    expect(call('SUMIFS', [Q, Pr, '<>Bananas', S, 'Tom'], c2)).toBe(30);
    expect(call('SUMIFS', [Q, ref(0, 0, 1, 6, 1), 'A*'], c2)).toEqual(Err.value);
    expect(call('SUMIFS', [Q, Pr], c2)).toEqual(Err.value);
  });
  it('AVERAGEIF / AVERAGEIFS / MAXIFS / MINIFS', () => {
    expect(call('AVERAGEIF', [C, '<23000'], ctx)).toBe(14000);
    expect(call('AVERAGEIF', [P, '<95000'], ctx)).toEqual(Err.div0);
    expect(call('AVERAGEIF', [P, '>250000', C], ctx)).toBe(24500);
    expect(call('AVERAGEIFS', [C, P, '>100000', P, '<400000'], ctx)).toBe(17500);
    expect(call('MAXIFS', [C, P, '<300000'], ctx)).toBe(14000);
    expect(call('MINIFS', [C, P, '>100000'], ctx)).toBe(14000);
    expect(call('MAXIFS', [C, P, '>1e9'], ctx)).toBe(0);
  });
  it('errors in the summed cells propagate only when matched', () => {
    const c3 = makeCtx([[['a', 1], ['b', Err.div0]]]);
    expect(call('SUMIF', [ref(0, 0, 0, 1, 0), 'a', ref(0, 0, 1, 1, 1)], c3)).toBe(1);
    expect(call('SUMIF', [ref(0, 0, 0, 1, 0), 'b', ref(0, 0, 1, 1, 1)], c3)).toEqual(Err.div0);
  });
});

describe('SUMPRODUCT', () => {
  const ctx = makeCtx();
  it('multiplies and sums, text as zero, size mismatch is #VALUE!', () => {
    expect(call('SUMPRODUCT', [[[3, 4], [8, 6], [1, 9]], [[2, 7], [6, 7], [5, 3]]], ctx)).toBe(156);
    expect(call('SUMPRODUCT', [[[1], ['x'], [3]], [[2], [2], [2]]], ctx)).toBe(8);
    expect(call('SUMPRODUCT', [[[1], [2]], [[1], [2], [3]]], ctx)).toEqual(Err.value);
    expect(call('SUMPRODUCT', [[[1], [Err.na]]], ctx)).toEqual(Err.na);
    expect(call('SUMPRODUCT', [[[true], [2]]], ctx)).toBe(2);
  });
});

describe('SUBTOTAL / AGGREGATE', () => {
  const g: Scalar[][] = [[120], [10], [150], [23], [Err.div0]];
  const hidden = new Set([1]);
  const ctx: FnContext = { ...makeCtx([g]), rowHidden: (_s, r) => hidden.has(r), isSubtotalCell: (_s, r) => r === 3 };
  const R4 = ref(0, 0, 0, 3, 0), R5 = ref(0, 0, 0, 4, 0);
  it('SUBTOTAL skips nested subtotals and, for 101+, hidden rows', () => {
    expect(call('SUBTOTAL', [9, R4], ctx)).toBe(280);
    expect(call('SUBTOTAL', [109, R4], ctx)).toBe(270);
    expect(call('SUBTOTAL', [1, R4], ctx)).toBe(280 / 3);
    expect(call('SUBTOTAL', [2, R5], ctx)).toBe(3);
    expect(call('SUBTOTAL', [3, R5], ctx)).toBe(4);
    expect(call('SUBTOTAL', [9, R5], ctx)).toEqual(Err.div0);
    expect(call('SUBTOTAL', [12, R4], ctx)).toEqual(Err.value);
    expect(call('SUBTOTAL', [9, [[1]]], ctx)).toEqual(Err.value);
  });
  it('AGGREGATE options and array form', () => {
    expect(call('AGGREGATE', [9, 6, R5], ctx)).toBe(303);
    expect(call('AGGREGATE', [9, 7, R5], ctx)).toBe(293);
    expect(call('AGGREGATE', [9, 3, R5], ctx)).toBe(270);
    expect(call('AGGREGATE', [4, 6, R5], ctx)).toBe(150);
    expect(call('AGGREGATE', [14, 6, R5, 2], ctx)).toBe(120);
    expect(call('AGGREGATE', [15, 6, R5, 1], ctx)).toBe(10);
    expect(call('AGGREGATE', [12, 6, R5], ctx)).toBe(71.5);
    expect(call('AGGREGATE', [14, 6, [[1, 5, Err.na, 3]], 1], ctx)).toBe(5);
    expect(call('AGGREGATE', [16, 6, [[1, 2, 3, 4]], 0.5], ctx)).toBe(2.5);
    expect(call('AGGREGATE', [17, 6, [[1, 2, 3, 4]], 1], ctx)).toBe(1.75);
    expect(call('AGGREGATE', [19, 6, [[1, 2, 3, 4]], 1], ctx)).toBe(1.25);
    expect(call('AGGREGATE', [14, 6, R5, 10], ctx)).toEqual(Err.num);
    expect(call('AGGREGATE', [9, 0, R5], ctx)).toEqual(Err.div0);
    expect(call('AGGREGATE', [20, 0, R5], ctx)).toEqual(Err.value);
  });
});
