import { beforeAll, describe, expect, it } from 'vitest';
import { Engine } from '../src/engine/engine';
import { nodeText, parseFormula } from '../src/engine/parser';
import { registerFunctions } from '../src/engine/registry';
import { collectNumbers, num } from '../src/engine/coerce';
import { CellError, isError, type FunctionDef } from '../src/engine/types';
import { newSheet, newWorkbook, type Workbook } from '../src/model/workbook';

const SUM: FunctionDef = { name: 'SUM', min: 1, max: 255, fn: (a, ctx) => { const n = collectNumbers(a, ctx); return isError(n) ? n : n.reduce((x, y) => x + y, 0); } };
const SEQ: FunctionDef = { name: 'SEQTEST', min: 1, max: 1, fn: (a, ctx) => { const n = num(a[0]!, ctx) as number; return Array.from({ length: n }, (_, i) => [i + 1]); } };
const ABSX: FunctionDef = { name: 'ABSX', min: 1, max: 1, scalar: true, fn: (a) => Math.abs(a[0] as number) };

beforeAll(() => registerFunctions([SUM, SEQ, ABSX]));

const book = (): { wb: Workbook; e: Engine; set: (a: string, v: string | number) => void; get: (a: string) => unknown } => {
  const wb = newWorkbook();
  wb.sheets.push(newSheet('Data Sheet'));
  const e = new Engine(wb);
  const pos = (a: string) => { const m = /^([A-Z]+)(\d+)$/.exec(a)!; return [+m[2]! - 1, m[1]!.charCodeAt(0) - 65] as const; };
  return {
    wb, e,
    set: (a, v) => { const [r, c] = pos(a); if (typeof v === 'string' && v.startsWith('=')) e.setFormula(0, r, c, v.slice(1)); else e.setValue(0, r, c, v); e.recalc(); },
    get: (a) => { const [r, c] = pos(a); const v = e.getValue(0, r, c); return v instanceof CellError ? v.code : v; },
  };
};

describe('parser', () => {
  it.each([
    ['1+2*3', '1+2*3'], ['(1+2)*3', '(1+2)*3'], ['-2^2', '-2^2'], ['SUM(A1:B2,$C$3)', 'SUM(A1:B2,$C$3)'],
    ["'My Sheet'!A1", "'My Sheet'!A1"], ['Sheet1:Sheet3!A1', 'Sheet1:Sheet3!A1'], ['A:A', 'A:A'], ['1:3', '1:3'],
    ['{1,2;3,4}', '{1,2;3,4}'], ['"a""b"&"c"', '"a""b"&"c"'], ['LAMBDA(x,x+1)(2)', 'LAMBDA(x,x+1)(2)'],
    ['Table1[Amount]', 'Table1[Amount]'], ['A1#', 'A1#'], ['IF(A1>0,,1)', 'IF(A1>0,,1)'], ['50%', '50%'], ['_xlfn.XLOOKUP(1,A:A,B:B)', 'XLOOKUP(1,A:A,B:B)'],
  ])('round-trips %s', (src, out) => expect(nodeText(parseFormula(src))).toBe(out));

  it('rejects broken formulas', () => {
    expect(() => parseFormula('SUM(1,')).toThrow();
    expect(() => parseFormula('"abc')).toThrow();
  });
});

describe('engine', () => {
  it('calculates and updates dependents', () => {
    const b = book();
    b.set('A1', 2); b.set('A2', 3); b.set('A3', '=A1*A2'); b.set('A4', '=SUM(A1:A3)+1');
    expect(b.get('A3')).toBe(6); expect(b.get('A4')).toBe(12);
    b.set('A1', 10);
    expect(b.get('A3')).toBe(30); expect(b.get('A4')).toBe(44);
  });

  it('follows operator rules', () => {
    const b = book();
    b.set('A1', '=-2^2'); b.set('A2', '=1/0'); b.set('A3', '="a"&1&TRUE'); b.set('A4', '="abc"="ABC"'); b.set('A5', '=0.1+0.2-0.3'); b.set('A6', '=50%');
    b.set('A7', '="3"+4'); b.set('A8', '="x"+1'); b.set('A9', '=B9=0');
    expect([b.get('A1'), b.get('A2'), b.get('A3'), b.get('A4'), b.get('A5'), b.get('A6'), b.get('A7'), b.get('A8'), b.get('A9')])
      .toEqual([4, '#DIV/0!', 'a1TRUE', true, 0, 0.5, 7, '#VALUE!', true]);
  });

  it('spills arrays and reports blocked spills', () => {
    const b = book();
    b.set('A1', '=SEQTEST(3)');
    expect([b.get('A1'), b.get('A2'), b.get('A3')]).toEqual([1, 2, 3]);
    b.set('B1', '=SUM(A1#)');
    expect(b.get('B1')).toBe(6);
    b.set('A2', 'x');
    expect(b.get('A1')).toBe('#SPILL!');
    b.set('A2', null as unknown as number);
    b.e.recalc();
    expect(b.get('A1')).toBe(1);
    b.set('C1', '={1,2;3,4}*10');
    expect([b.get('C1'), b.get('D2')]).toEqual([10, 40]);
    b.set('E1', '=ABSX({-1,-2})');
    expect([b.get('E1'), b.get('F1')]).toEqual([1, 2]);
  });

  it('handles IF, IFERROR, LET, LAMBDA, CHOOSE, SWITCH lazily', () => {
    const b = book();
    b.set('A1', '=IF(1>2,1/0,"ok")'); b.set('A2', '=IFERROR(1/0,"bad")'); b.set('A3', '=LET(x,2,y,x*3,x+y)');
    b.set('A4', '=LAMBDA(a,b,a*b)(3,4)'); b.set('A5', '=CHOOSE(2,"a","b","c")'); b.set('A6', '=SWITCH(3,1,"one",3,"three","other")');
    b.set('A7', '=IF({TRUE,FALSE},1,2)'); b.set('A9', '=IFS(1>2,"a",TRUE,"b")');
    expect([b.get('A1'), b.get('A2'), b.get('A3'), b.get('A4'), b.get('A5'), b.get('A6'), b.get('A7'), b.get('B7'), b.get('A9')])
      .toEqual(['ok', 'bad', 8, 12, 'b', 'three', 1, 2, 'b']);
  });

  it('uses defined names, other sheets, INDIRECT and tables', () => {
    const b = book();
    b.wb.names.push({ name: 'Rate', formula: '0.5' }, { name: 'Twice', formula: 'LAMBDA(x,x*2)' });
    b.e.setValue(1, 0, 0, 40);
    b.set('A1', "='Data Sheet'!A1*Rate"); b.set('A2', '=Twice(5)'); b.set('A3', '=SUM(INDIRECT("A1:A2"))');
    expect([b.get('A1'), b.get('A2'), b.get('A3')]).toEqual([20, 10, 30]);
    b.e.setValue(1, 0, 0, 100); b.e.recalc();
    expect(b.get('A1')).toBe(50);
    b.wb.sheets[0]!.tables.push({ name: 'T', range: { r1: 9, c1: 0, r2: 11, c2: 1 }, headerRow: true, totalsRow: false, columns: [{ name: 'Item' }, { name: 'Qty' }] });
    b.set('A10', 'Item'); b.set('B10', 'Qty'); b.set('B11', 4); b.set('B12', 6);
    b.set('D1', '=SUM(T[Qty])');
    expect(b.get('D1')).toBe(10);
  });

  it('finds circular references without hanging', () => {
    const b = book();
    b.set('A1', '=A2+1'); b.set('A2', '=A1+1');
    expect(b.e.circular.length).toBeGreaterThan(0);
  });

  it('recalculates a long chain without overflowing the stack', () => {
    const b = book();
    for (let r = 1; r < 20000; r++) b.e.setFormula(0, r, 0, `A${r}+1`);
    b.e.setValue(0, 0, 0, 1);
    b.e.recalc();
    expect(b.e.getValue(0, 19999, 0)).toBe(20000);
    b.e.setValue(0, 0, 0, 5);
    b.e.recalc();
    expect(b.e.getValue(0, 19999, 0)).toBe(20004);
  });
});
