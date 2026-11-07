import { describe, expect, test } from 'vitest';
import { readCsv, writeCsv } from '../src/formats/csv';
import { CellError } from '../src/engine/types';

describe('readCsv', () => {
  test('RFC 4180 quoting, CRLF and BOM', () => {
    const { rows } = readCsv('﻿name,note\r\n"Smith, J","said ""hi""\r\nnext line"\r\nplain,\r\n');
    expect(rows).toEqual([['name', 'note'], ['Smith, J', 'said "hi"\r\nnext line'], ['plain', null]]);
  });

  test('LF and CR line ends, no trailing newline, ragged rows', () => {
    expect(readCsv('a,b\nc\rd,e,f').rows).toEqual([['a', 'b'], ['c'], ['d', 'e', 'f']]);
  });

  test('delimiter detection', () => {
    expect(readCsv('a;b;c\n1;2;3').rows).toEqual([['a', 'b', 'c'], [1, 2, 3]]);
    expect(readCsv('a\tb\n1\t2').rows).toEqual([['a', 'b'], [1, 2]]);
    expect(readCsv('a|b|c\n1|2|3').rows).toEqual([['a', 'b', 'c'], [1, 2, 3]]);
    expect(readCsv('x;"1,5";y\nq;"2,5";z').rows).toEqual([['x', '1,5', 'y'], ['q', '2,5', 'z']]);
    expect(readCsv('sep=;\na;b\n1;2').rows).toEqual([['a', 'b'], [1, 2]]);
    expect(readCsv('a;b\n1;2', { delimiter: ',' }).rows).toEqual([['a;b'], ['1;2']]);
  });

  test('values convert the way typed entries do, with formats', () => {
    const { rows, formats } = readCsv('1234,1.5,-7,"1,234.5",15%,$9.99,TRUE,false,#N/A,3/15/2024,14:30,1e3,abc,=\"007\", 42 ,007');
    expect(rows[0]).toEqual([1234, 1.5, -7, 1234.5, 0.15, 9.99, true, false, new CellError('#N/A'), 45366, 14.5 / 24, 1000, 'abc', '007', 42, 7]);
    expect(formats[0]).toEqual([undefined, undefined, undefined, '#,##0.00', '0%', '$#,##0.00', undefined, undefined, undefined, 'm/d/yyyy', 'h:mm', '0.00E+00', undefined, undefined, undefined, undefined]);
  });

  test('empty input', () => {
    expect(readCsv('')).toEqual({ rows: [], formats: [] });
    expect(readCsv('\n').rows).toEqual([[null]]);
  });

  test('unterminated quote reads to the end', () => {
    expect(readCsv('a,"b\nc').rows).toEqual([['a', 'b\nc']]);
  });
});

describe('writeCsv', () => {
  test('quotes when needed and uses CRLF', () => {
    expect(writeCsv([['a', 'b,c', 'say "x"'], ['line\nbreak', ' pad', '']])).toBe('a,"b,c","say ""x"""\r\n"line\nbreak"," pad",\r\n');
    expect(writeCsv([['a;b', 'c,d']], ';')).toBe('"a;b";c,d\r\n');
    expect(writeCsv([])).toBe('');
  });

  test('round trip', () => {
    const rows = [['x', 'y, z', 'q"uote'], ['multi\r\nline', '', 'end']];
    expect(readCsv(writeCsv(rows)).rows).toEqual([['x', 'y, z', 'q"uote'], ['multi\r\nline', null, 'end']]);
  });
});
