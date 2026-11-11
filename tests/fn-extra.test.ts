import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { Engine } from '../src/engine/engine';
import { registerAllFunctions } from '../src/engine/functions';
import { defs } from '../src/engine/functions/extra';
import { nodeText, parseFormula, trimRefsToFile } from '../src/engine/parser';
import { addFuturePrefixes, stripFuturePrefixes } from '../src/formats/formula-xml';
import { clearOnlineCache, setOnlineAllowed, settled } from '../src/engine/online';
import { newWorkbook } from '../src/model/workbook';
import { CellError, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

beforeAll(() => registerAllFunctions());
const g = globalThis as unknown as { fetch?: unknown };
const realFetch = g.fetch;
afterEach(() => { setOnlineAllowed(false); clearOnlineCache(); g.fetch = realFetch; });

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const code = (v: unknown): string | null => (v instanceof CellError ? v.code : null);

describe('trim-reference operators', () => {
  it('parse and print', () => {
    for (const f of ['SUM(A1.:.E10)', 'SUM(A1.:E10)', 'SUM(A1:.E10)', 'COUNTA(A:.A)', 'SUM(Sheet2!$B$2.:.C9)', 'ROWS(1:.5)']) {
      expect(nodeText(parseFormula(f))).toBe(f);
    }
    expect(parseFormula('A1:.B5')).toMatchObject({ type: 'ref', ref: { trim: 'trailing', r1: 0, r2: 4 } });
  });
  it('files store them as _TRO_ calls, and read back as operators', () => {
    expect(trimRefsToFile('SUM(A1.:.E10)+COUNTA(A:.A)')).toBe('SUM(_xlfn._TRO_ALL(A1:E10))+COUNTA(_xlfn._TRO_TRAILING(A:A))');
    expect(addFuturePrefixes('SUM(A1.:E10)')).toBe('SUM(_xlfn._TRO_LEADING(A1:E10))');
    expect(nodeText(parseFormula(stripFuturePrefixes('SUM(_xlfn._TRO_ALL(A1:E10))')))).toBe('SUM(A1.:.E10)');
    expect(addFuturePrefixes('SUM(A1:E10)')).toBe('SUM(A1:E10)');
  });
  it('evaluate like TRIMRANGE', () => {
    const e = new Engine(newWorkbook());
    for (const [r, c, v] of [[2, 0, 5], [3, 0, 6], [4, 1, 7]] as const) e.setValue(0, r, c, v);
    e.setFormula(0, 0, 4, 'ROWS(A1:.B1000)');
    e.setFormula(0, 1, 4, 'ROWS(A1.:B1000)');
    e.setFormula(0, 2, 4, 'ROWS(A1.:.B1000)');
    e.setFormula(0, 3, 4, 'SUM(A:.A)');
    e.setFormula(0, 4, 4, 'COLUMNS(A1.:.C1000)');
    e.setFormula(0, 5, 4, 'ROWS(G1.:.G9)');
    e.recalc();
    expect(e.getValue(0, 0, 4)).toBe(5);
    expect(e.getValue(0, 1, 4)).toBe(998);
    expect(e.getValue(0, 2, 4)).toBe(3);
    expect(e.getValue(0, 3, 4)).toBe(11);
    expect(e.getValue(0, 4, 4)).toBe(2);
    expect(code(e.getValue(0, 5, 4))).toBe('#CALC!');
    // Adding data below grows the trimmed range.
    e.setValue(0, 9, 0, 1);
    e.recalc();
    expect(e.getValue(0, 0, 4)).toBe(10);
  });
});

describe('EUROCONVERT', () => {
  it('converts with the fixed rates and rounding rules', () => {
    expect(call('EUROCONVERT', 1.2, 'DEM', 'EUR')).toBe(0.61);
    expect(call('EUROCONVERT', 1, 'FRF', 'EUR', false, 3)).toBe(0.15);
    expect(call('EUROCONVERT', 1, 'FRF', 'DEM', true, 3)).toBeCloseTo(0.29728616, 10);
    expect(call('EUROCONVERT', 1, 'FRF', 'DEM', false, 3)).toBe(0.3);
    expect(call('EUROCONVERT', 1, 'EUR', 'ITL')).toBe(1936);
    expect(call('EUROCONVERT', 1, 'EUR', 'PTE')).toBe(200.5);
    expect(call('EUROCONVERT', 100, 'eur', 'hrk')).toBe(753.45);
    expect(code(call('EUROCONVERT', 1, 'USD', 'EUR'))).toBe('#VALUE!');
    expect(code(call('EUROCONVERT', 1, 'FRF', 'DEM', false, 2))).toBe('#VALUE!');
  });
});

describe('other catalog functions', () => {
  it('aliases and blocked functions', () => {
    const e = new Engine(newWorkbook());
    e.setFormula(0, 0, 0, 'ECMA.CEILING(2.5,1)');
    e.setFormula(0, 0, 1, 'JIS("ABC")');
    e.setFormula(0, 0, 2, 'DBCS("ABC")');
    e.setFormula(0, 0, 3, 'CALL("kernel32","x","A")');
    e.setFormula(0, 0, 4, 'RTD("a","","b")');
    e.recalc();
    expect(e.getValue(0, 0, 0)).toBe(3);
    expect(e.getValue(0, 0, 1)).toBe(e.getValue(0, 0, 2));
    expect(code(e.getValue(0, 0, 3))).toBe('#BLOCKED!');
    expect(code(e.getValue(0, 0, 4))).toBe('#N/A');
  });
  it('IMPORTCSV and IMPORTTEXT read https files when allowed', async () => {
    expect(code(call('IMPORTCSV', 'https://example.com/a.csv'))).toBe('#CONNECT!');
    expect(code(call('IMPORTCSV', 'C:\\data\\a.csv'))).toBe('#BLOCKED!');
    setOnlineAllowed(true);
    g.fetch = async () => new Response('name,qty\nPens,3\nInk,12\nPaper,7\n');
    expect(code(call('IMPORTCSV', 'https://example.com/a.csv', 1))).toBe('#BUSY!');
    await settled();
    expect(call('IMPORTCSV', 'https://example.com/a.csv', 1)).toEqual([['Pens', 3], ['Ink', 12], ['Paper', 7]]);
    expect(call('IMPORTCSV', 'https://example.com/a.csv', 1, 2)).toEqual([['Pens', 3], ['Ink', 12]]);
    g.fetch = async () => new Response('a;b\n1,5;2\n');
    call('IMPORTTEXT', 'https://example.com/t.txt', ';', 1, null, 'utf-8', 'de-DE');
    await settled();
    expect(call('IMPORTTEXT', 'https://example.com/t.txt', ';', 1, null, 'utf-8', 'de-DE')).toEqual([[1.5, 2]]);
    expect(code(call('IMPORTTEXT', 'https://example.com/t.txt', ';;'))).toBe('#VALUE!');
    // The locale reads the raw text: "1.234" is a thousand-grouped number in German, a date uses day.month.
    g.fetch = async () => new Response('Wert;Datum;Anteil;Preis\n1.234;15.03.2024;12,5 %;3,50 €\n0,5;01.02.2023;-7;(2,25)\n');
    call('IMPORTCSV', 'https://example.com/de.csv', 1, null, 'de-DE');
    await settled();
    const { toSerial } = await import('../src/engine/date');
    expect(call('IMPORTCSV', 'https://example.com/de.csv', 1, null, 'de-DE')).toEqual([
      [1234, toSerial(2024, 3, 15), 0.125, 3.5],
      [0.5, toSerial(2023, 2, 1), -7, -2.25],
    ]);
    g.fetch = async () => new Response('a,b\n"1,234.5",3/15/2024\n');
    call('IMPORTCSV', 'https://example.com/us.csv', 1, null, 'en-US');
    await settled();
    expect(call('IMPORTCSV', 'https://example.com/us.csv', 1, null, 'en-US')).toEqual([[1234.5, toSerial(2024, 3, 15)]]);
    expect(code(call('IMPORTCSV', 'https://example.com/us.csv', 1, null, 'xx-nope-locale!'))).toBe('#VALUE!');
  });
});
