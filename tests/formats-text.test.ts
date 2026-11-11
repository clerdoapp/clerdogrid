import { describe, expect, it } from 'vitest';
import { cellKey, newWorkbook, type DataQuery } from '../src/model/workbook';
import {
  convertRows, decodeText, detectBreaks, guessDelimiter, guessOptions, importText, looksFixedWidth, parseDateOrder, parseLocaleNumber, splitDelimited, splitFixed, splitLines,
} from '../src/formats/text-import';
import { queriesFromXml, queriesToXml } from '../src/formats/queries-xml';
import { readXlsx } from '../src/formats/xlsx-read';
import { writeXlsx } from '../src/formats/xlsx-write';
import { readOds } from '../src/formats/ods-read';
import { writeOds } from '../src/formats/ods-write';
import { workbookFromBytes } from '../src/app/files';
import { writeXls } from '../src/formats/biff8-write';
import { loadQueries, stashQueries, workbookBlock } from '../src/features/formats/ops';
import { toSerial } from '../src/engine/date';

const enc = (s: string) => new TextEncoder().encode(s);

describe('text decoding', () => {
  it('detects byte-order marks, UTF-16 and falls back to Windows-1252', () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x41])).text).toBe('A');
    expect(decodeText(new Uint8Array([0xff, 0xfe, 0x41, 0, 0x42, 0]))).toEqual({ text: 'AB', encoding: 'utf-16le' });
    expect(decodeText(new Uint8Array([0x41, 0, 0x42, 0, 0x43, 0, 0x44, 0])).encoding).toBe('utf-16le');
    expect(decodeText(new Uint8Array([0x63, 0x61, 0x66, 0xe9]))).toEqual({ text: 'café', encoding: 'windows-1252' });
    expect(decodeText(enc('café ✓')).text).toBe('café ✓');
  });
  it('decodes chosen encodings', () => {
    expect(decodeText(new Uint8Array([0x82, 0xa0]), 'shift_jis').text).toBe('あ');
    expect(decodeText(new Uint8Array([0xc7, 0xe1]), 'windows-1256').text).toBe('ال');
    expect(decodeText(new Uint8Array([0xc4, 0xe3]), 'gb18030').text).toBe('你');
    expect(decodeText(new Uint8Array([0xe9]), 'iso-8859-1').text).toBe('é');
  });
});

describe('delimited text', () => {
  it('splits with qualifiers, embedded separators and line breaks', () => {
    expect(splitDelimited('a,"b,c",d\n"x ""y""","line\nbreak",z\n', { delimiters: [','], qualifier: '"' })).toEqual([['a', 'b,c', 'd'], ['x "y"', 'line\nbreak', 'z']]);
  });
  it('merges repeated separators and supports several at once and multi-character ones', () => {
    expect(splitDelimited('a  b\tc', { delimiters: [' ', '\t'], qualifier: '', mergeDelimiters: true })).toEqual([['a', 'b', 'c']]);
    expect(splitDelimited('a::b::c', { delimiters: ['::'], qualifier: '"' })).toEqual([['a', 'b', 'c']]);
    expect(splitDelimited("'a;b';c", { delimiters: [';'], qualifier: "'" })).toEqual([['a;b', 'c']]);
  });
  it('guesses the separator', () => {
    expect(guessDelimiter('a;b;c\n1;2;3\n')).toBe(';');
    expect(guessDelimiter('a\tb\n1\t2\n')).toBe('\t');
    expect(guessDelimiter('a,b\n1,2\n')).toBe(',');
  });
  it('guesses a comma decimal separator with semicolons', () => {
    const o = guessOptions(enc('name;amount\napple;1,5\npear;2,25\n'));
    expect(o.delimited.delimiters).toEqual([';']);
    expect(o.locale.decimal).toBe(',');
    const { values } = importText(enc('name;amount\napple;1,5\npear;1.234,25\n'), o);
    expect(values[1]).toEqual(['apple', 1.5]);
    expect(values[2]).toEqual(['pear', 1234.25]);
  });
  it('starts at a chosen row', () => {
    const o = { ...guessOptions(enc('title\nskip\na,b\n1,2\n')), startRow: 3 };
    o.delimited = { delimiters: [','], qualifier: '"' };
    expect(importText(enc('title\nskip\na,b\n1,2\n'), o).values).toEqual([['a', 'b'], [1, 2]]);
  });
});

describe('fixed-width text', () => {
  const text = 'Name      Qty  Price\nApple      12   1.50\nPear        3  12.00\n';
  it('finds column breaks', () => {
    const lines = splitLines(text);
    expect(looksFixedWidth(text)).toBe(true);
    expect(detectBreaks(lines)).toEqual([10, 15]);
    expect(splitFixed(lines, [10, 15])).toEqual([['Name', 'Qty', 'Price'], ['Apple', '12', '1.50'], ['Pear', '3', '12.00']]);
  });
  it('imports .prn files as fixed width', () => {
    const o = guessOptions(enc(text), 'list.prn');
    expect(o.kind).toBe('fixed');
    expect(importText(enc(text), o).values[2]).toEqual(['Pear', 3, 12]);
  });
});

describe('column conversion', () => {
  it('reads numbers in different locales', () => {
    expect(parseLocaleNumber('1,234.5', { decimal: '.', thousands: ',' })?.n).toBe(1234.5);
    expect(parseLocaleNumber('1.234,5', { decimal: ',', thousands: '.' })?.n).toBe(1234.5);
    expect(parseLocaleNumber('1 234,5', { decimal: ',', thousands: ' ' })?.n).toBe(1234.5);
    expect(parseLocaleNumber("1'234.5", { decimal: '.', thousands: "'" })?.n).toBe(1234.5);
    expect(parseLocaleNumber('(12.5)', { decimal: '.', thousands: ',' })?.n).toBe(-12.5);
    expect(parseLocaleNumber('12.5%', { decimal: '.', thousands: ',' })).toEqual({ n: 0.125, fmt: '0.0%' });
    expect(parseLocaleNumber('$1,200', { decimal: '.', thousands: ',' })?.n).toBe(1200);
    expect(parseLocaleNumber('12,34,5', { decimal: '.', thousands: ',' })).toBeNull();
    expect(parseLocaleNumber('abc', { decimal: '.', thousands: ',' })).toBeNull();
  });
  it('reads dates in a chosen order', () => {
    expect(parseDateOrder('03/04/2024', 'MDY')).toBe(toSerial(2024, 3, 4));
    expect(parseDateOrder('03/04/2024', 'DMY')).toBe(toSerial(2024, 4, 3));
    expect(parseDateOrder('2024-04-03 14:30', 'YMD')).toBeCloseTo(toSerial(2024, 4, 3) + 14.5 / 24, 8);
    expect(parseDateOrder('3 Apr 24', 'DMY')).toBe(toSerial(2024, 4, 3));
    expect(parseDateOrder('13/13/2024', 'MDY')).toBeNull();
  });
  it('applies column types: text, dates and skipped columns', () => {
    const r = convertRows([['007', '1/2/2024', 'x', 'TRUE', '#N/A']], { formats: ['text', 'DMY', 'skip'], locale: { decimal: '.', thousands: ',' } });
    expect(r.values[0]![0]).toBe('007');
    expect(r.formats[0]![0]).toBe('@');
    expect(r.values[0]![1]).toBe(toSerial(2024, 2, 1));
    expect(r.values[0]!.length).toBe(4);
    expect(r.values[0]![2]).toBe(true);
  });
});

describe('Get Data queries', () => {
  const q: DataQuery = { id: 'q1', name: 'Sales "2024"', kind: 'text', file: 'sales.csv', options: JSON.stringify({ kind: 'delimited' }), target: { sheet: 'Data', row: 2, col: 1, rows: 10, cols: 3 }, refreshed: '2026-11-08T10:00:00.000Z' };
  const q2: DataQuery = { id: 'q2', name: 'Totals', kind: 'workbook', file: 'book.xlsx', source: { sheet: 'Sheet 1', range: 'A1:B4' }, target: { sheet: 'T', row: 0, col: 0, rows: 4, cols: 2 } };
  it('writes and reads the query part', () => {
    expect(queriesFromXml(queriesToXml([q, q2]))).toEqual([q, q2]);
    expect(queriesFromXml('<queries xmlns="other"><query kind="text"/></queries>')).toEqual([]);
  });
  it('keeps queries through .xlsx save and reopen', () => {
    const wb = newWorkbook();
    wb.queries = [q, q2];
    const store = { wb } as unknown as Parameters<typeof stashQueries>[0];
    const undo = stashQueries(store);
    const bytes = writeXlsx(wb);
    undo?.();
    expect(wb.extraParts).toBeUndefined();
    const back = readXlsx(bytes);
    loadQueries(back);
    expect(back.queries).toEqual([q, q2]);
  });
  it('keeps queries through .ods save and reopen', () => {
    const wb = newWorkbook();
    wb.queries = [q];
    expect(readOds(writeOds(wb)).queries).toEqual([q]);
  });
  it('takes a range, table or whole sheet from another workbook', () => {
    const wb = newWorkbook();
    const s = wb.sheets[0]!;
    for (let r = 0; r < 4; r++) for (let c = 0; c < 3; c++) s.cells.set(cellKey(r, c), { v: r * 10 + c });
    s.tables.push({ name: 'T1', range: { r1: 1, c1: 1, r2: 2, c2: 2 }, headerRow: false, totalsRow: false, columns: [{ name: 'a' }, { name: 'b' }] });
    const val = (si: number, r: number, c: number) => wb.sheets[si]!.cells.get(cellKey(r, c))?.v ?? null;
    expect((workbookBlock(wb, 'Sheet1', undefined, val) as { values: unknown[][] }).values.length).toBe(4);
    expect((workbookBlock(wb, 'Sheet1', 'B2:C3', val) as { values: unknown[][] }).values).toEqual([[11, 12], [21, 22]]);
    expect((workbookBlock(wb, 'Sheet1', 'T1', val) as { values: unknown[][] }).values).toEqual([[11, 12], [21, 22]]);
    expect(workbookBlock(wb, 'Nope', undefined, val)).toBe('That sheet was not found in the file.');
  });
});

describe('opening files by type', () => {
  const wb = newWorkbook();
  wb.sheets[0]!.cells.set(0, { v: 'hello' });
  it('routes .ods, .ots, .xls, .xlt and template names to the right reader', () => {
    expect(workbookFromBytes('a.ods', writeOds(wb)).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('a.ots', writeOds(wb, { template: true })).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('a.xls', writeXls(wb)).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('a.xlt', writeXls(wb)).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('a.xltm', writeXlsx(wb)).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('renamed.bin', writeOds(wb)).sheets[0]!.cells.get(0)!.v).toBe('hello');
    expect(workbookFromBytes('a.txt', enc('x,1\ny,2\n')).sheets[0]!.cells.get(cellKey(1, 1))!.v).toBe(2);
  });
});

describe('dialogs render', () => {
  it('shows the text import, workbook import, export and password dialogs', async () => {
    const { createElement } = await import('react');
    const { renderToString } = await import('react-dom/server');
    const { GridStore } = await import('../src/app/store');
    const d = await import('../src/features/formats/dialogs');
    const store = new GridStore();
    const g = globalThis as { window?: unknown };
    const hadWindow = 'window' in g;
    if (!hadWindow) g.window = { innerWidth: 1200 };
    const close = () => undefined;
    const file = { name: 'list.csv', data: enc('a;b\n1,5;2\n') };
    const html = renderToString(createElement(d.TextImportDialog, { store, close, data: { file, mode: 'insert' } }));
    expect(html).toContain('Import list.csv');
    expect(html).toContain('1,5');
    const fixed = renderToString(createElement(d.TextImportDialog, { store, close, data: { file: { name: 'p.prn', data: enc('ab   cd\nef   gh\n') }, mode: 'open' } }));
    expect(fixed).toContain('fmt-break');
    const wb = newWorkbook();
    wb.sheets[0]!.cells.set(0, { v: 1 });
    expect(renderToString(createElement(d.WorkbookImportDialog, { store, close, data: { file: { name: 'b.xlsx', data: new Uint8Array() }, wb } }))).toContain('Sheet1');
    expect(renderToString(createElement(d.ExportDialog, { store, close, data: { kind: 'json' } }))).toContain('First row has field names');
    expect(renderToString(createElement(d.XlsPasswordDialog, { store, close, data: { name: 'x.xls', data: new Uint8Array(), path: null, wrong: true } }))).toContain('not right');
    if (!hadWindow) delete g.window;
  });
});

describe('Get Data queries in older workbooks and undo', () => {
  const q: DataQuery = { id: 'q9', name: 'Old', kind: 'workbook', file: 'a.xlsx', path: '/home/me/a.xlsx', source: { sheet: 'S' }, target: { sheet: 'T', row: 0, col: 0, rows: 1, cols: 1 } };
  it('keeps queries in .xls, also when encrypted', async () => {
    const { readXls } = await import('../src/formats/biff8-read');
    const wb = newWorkbook();
    wb.queries = [q];
    expect(readXls(writeXls(wb)).queries).toEqual([q]);
    expect(readXls(writeXls(wb, { password: 'x' }), { password: 'x' }).queries).toEqual([q]);
  });
  it('undoes removing a query and an import', async () => {
    const { GridStore } = await import('../src/app/store');
    const { deleteQuery, placeBlock } = await import('../src/features/formats/ops');
    const store = new GridStore();
    const nq: DataQuery = { ...q, id: 'q10', target: { sheet: '', row: 0, col: 0, rows: 0, cols: 0 } };
    placeBlock(store, nq, { values: [[1, 2]], formats: [[undefined, undefined]] }, 'Import');
    expect(store.wb.queries?.map((x) => x.id)).toEqual(['q10']);
    deleteQuery(store, store.wb.queries![0]!);
    expect(store.wb.queries).toBeUndefined();
    store.undo();
    expect(store.wb.queries?.map((x) => x.id)).toEqual(['q10']);
    store.undo();
    expect(store.wb.queries).toBeUndefined();
  });
});
