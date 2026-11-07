import { describe, expect, it } from 'vitest';
import { defs as logical } from '../src/engine/functions/logical';
import { defs as info } from '../src/engine/functions/info';
import { CellError, Err, ref, type FnContext, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const grid = [
  [1, 'text', true, null],
  [0, '', new CellError('#N/A'), 3.5],
];
const base = makeCtx([grid, [], []]);
const ctx: FnContext = { ...base, formula: (s, r, c) => (s === 0 && r === 0 && c === 3 ? '=A1+1' : null), numberFormat: (_s, r, c) => (r === 1 && c === 3 ? '0.00%' : 'General') };
const all = [...logical, ...info];
const call = (name: string, ...args: Value[]): Value => all.find((d) => d.name === name)!.fn(args, ctx);

describe('logical', () => {
  it('AND/OR/XOR/NOT', () => {
    expect(call('AND', true, true)).toBe(true);
    expect(call('AND', true, false)).toBe(false);
    expect(call('AND', 2 > 1, 'TRUE')).toBe(true);
    expect(call('AND', 'abc')).toEqual(Err.value);
    expect(call('AND', ref(0, 0, 0, 0, 2))).toBe(true); // text in ranges is ignored
    expect(call('AND', ref(0, 1, 0, 1, 1))).toBe(false);
    expect(call('AND', ref(0, 0, 1, 0, 1))).toEqual(Err.value); // no logical values
    expect(call('AND', ref(0, 0, 0, 1, 3))).toEqual(Err.na);
    expect(call('OR', false, [[0, 1]])).toBe(true);
    expect(call('OR', false, null)).toBe(false);
    expect(call('XOR', 3 > 0, 2 < 9)).toBe(false);
    expect(call('XOR', 3 > 12, 4 > 6)).toBe(false);
    expect(call('XOR', true, true, true)).toBe(true);
    expect(call('NOT', false)).toBe(true);
    expect(call('NOT', 0)).toBe(true);
    expect(call('NOT', 'x')).toEqual(Err.value);
    expect(call('TRUE')).toBe(true);
    expect(call('FALSE')).toBe(false);
  });
});

describe('IS functions', () => {
  it('type tests', () => {
    expect(call('ISBLANK', null)).toBe(true);
    expect(call('ISBLANK', '')).toBe(false);
    expect(call('ISERR', Err.na)).toBe(false);
    expect(call('ISERR', Err.value)).toBe(true);
    expect(call('ISERROR', Err.na)).toBe(true);
    expect(call('ISNA', Err.na)).toBe(true);
    expect(call('ISLOGICAL', true)).toBe(true);
    expect(call('ISNONTEXT', null)).toBe(true);
    expect(call('ISNONTEXT', 'x')).toBe(false);
    expect(call('ISNUMBER', '5')).toBe(false);
    expect(call('ISNUMBER', 5)).toBe(true);
    expect(call('ISTEXT', 'x')).toBe(true);
    expect(call('ISEVEN', -1)).toBe(false);
    expect(call('ISEVEN', 2.5)).toBe(true);
    expect(call('ISEVEN', 5)).toBe(false);
    expect(call('ISODD', -1)).toBe(true);
    expect(call('ISODD', '3')).toBe(true);
    expect(call('ISODD', null)).toBe(false);
    expect(call('ISODD', true)).toEqual(Err.value);
    expect(call('ISEVEN', 'x')).toEqual(Err.value);
    expect(all.find((d) => d.name === 'ISBLANK')!.scalar).toBe(true);
  });
  it('ISREF/ISFORMULA/FORMULATEXT', () => {
    expect(call('ISREF', ref(0, 0, 0))).toBe(true);
    expect(call('ISREF', { kind: 'refs', areas: [ref(0, 0, 0)] })).toBe(true);
    expect(call('ISREF', 5)).toBe(false);
    expect(call('ISREF', [[1]])).toBe(false);
    expect(call('ISFORMULA', ref(0, 0, 3))).toBe(true);
    expect(call('ISFORMULA', ref(0, 0, 0))).toBe(false);
    expect(call('ISFORMULA', 5)).toEqual(Err.value);
    expect(call('FORMULATEXT', ref(0, 0, 3))).toBe('=A1+1');
    expect(call('FORMULATEXT', ref(0, 0, 0))).toEqual(Err.na);
  });
});

describe('info', () => {
  it('N/NA/TYPE/ERROR.TYPE', () => {
    expect(call('N', 7)).toBe(7);
    expect(call('N', 'Even')).toBe(0);
    expect(call('N', true)).toBe(1);
    expect(call('N', Err.div0)).toEqual(Err.div0);
    expect(call('NA')).toEqual(Err.na);
    expect(call('TYPE', 'Smith')).toBe(2);
    expect(call('TYPE', 2)).toBe(1);
    expect(call('TYPE', null)).toBe(1);
    expect(call('TYPE', true)).toBe(4);
    expect(call('TYPE', Err.na)).toBe(16);
    expect(call('TYPE', [[1, 2, 3]])).toBe(64);
    expect(call('TYPE', ref(0, 0, 1))).toBe(2);
    expect(call('TYPE', ref(0, 0, 0, 1, 1))).toBe(64);
    expect(call('TYPE', { kind: 'lambda', params: [], invoke: () => 1 })).toBe(128);
    expect(call('ERROR.TYPE', Err.null)).toBe(1);
    expect(call('ERROR.TYPE', Err.div0)).toBe(2);
    expect(call('ERROR.TYPE', Err.na)).toBe(7);
    expect(call('ERROR.TYPE', Err.calc)).toBe(14);
    expect(call('ERROR.TYPE', 5)).toEqual(Err.na);
  });
  it('CELL/INFO/SHEET/SHEETS', () => {
    expect(call('CELL', 'address', ref(0, 1, 3))).toBe('$D$2');
    expect(call('CELL', 'address', ref(1, 0, 27))).toBe('Sheet2!$AB$1');
    expect(call('CELL', 'row', ref(0, 4, 2))).toBe(5);
    expect(call('CELL', 'col', ref(0, 4, 2))).toBe(3);
    expect(call('CELL', 'contents', ref(0, 0, 1))).toBe('text');
    expect(call('CELL', 'type', ref(0, 0, 1))).toBe('l');
    expect(call('CELL', 'type', ref(0, 0, 3))).toBe('b');
    expect(call('CELL', 'type', ref(0, 0, 0))).toBe('v');
    expect(call('CELL', 'format', ref(0, 0, 0))).toBe('G');
    expect(call('CELL', 'format', ref(0, 1, 3))).toBe('P2');
    expect(call('CELL', 'prefix', ref(0, 0, 1))).toBe("'");
    expect(call('CELL', 'sheet', ref(1, 0, 0))).toBe('Sheet2');
    expect(call('CELL', 'filename', ref(0, 0, 0))).toBe('');
    expect(call('CELL', 'bogus', ref(0, 0, 0))).toEqual(Err.value);
    expect(call('CELL', 'row')).toBe(1);
    expect(call('INFO', 'recalc')).toBe('Automatic');
    expect(call('INFO', 'system')).toBe('pcdos');
    expect(call('INFO', 'numfile')).toBe(3);
    expect(call('INFO', 'nope')).toEqual(Err.value);
    expect(call('SHEET')).toBe(1);
    expect(call('SHEET', ref(2, 0, 0))).toBe(3);
    expect(call('SHEET', 'Sheet2')).toBe(2);
    expect(call('SHEET', 'Missing')).toEqual(Err.na);
    expect(call('SHEETS')).toBe(3);
    expect(call('SHEETS', ref(0, 0, 0))).toBe(1);
  });
});
