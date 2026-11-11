import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/text';
import { Err, ref, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx([[['a', 'b', null, 'c'], [1, true, Err.div0, 'd']]]);
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);

describe('characters and codes', () => {
  it('CHAR/CODE/UNICHAR/UNICODE', () => {
    expect(call('CHAR', 65)).toBe('A');
    expect(call('CHAR', 33)).toBe('!');
    expect(call('CHAR', 128)).toBe('€');
    expect(call('CHAR', 0)).toEqual(Err.value);
    expect(call('CHAR', 256)).toEqual(Err.value);
    expect(call('CODE', 'A')).toBe(65);
    expect(call('CODE', '!')).toBe(33);
    expect(call('CODE', '€')).toBe(128);
    expect(call('CODE', '')).toEqual(Err.value);
    expect(call('CODE', 12)).toBe(49);
    expect(call('UNICHAR', 66)).toBe('B');
    expect(call('UNICHAR', 32)).toBe(' ');
    expect(call('UNICHAR', 0)).toEqual(Err.value);
    expect(call('UNICHAR', 128512)).toBe('😀');
    expect(call('UNICODE', ' ')).toBe(32);
    expect(call('UNICODE', 'B')).toBe(66);
    expect(call('UNICODE', '😀')).toBe(128512);
  });
  it('CLEAN/TRIM/case', () => {
    expect(call('CLEAN', '\x09Monthly report\x0A')).toBe('Monthly report');
    expect(call('TRIM', ' First Quarter    Earnings ')).toBe('First Quarter Earnings');
    expect(call('TRIM', ' x')).toBe(' x');
    expect(call('LOWER', 'E. E. Cummings')).toBe('e. e. cummings');
    expect(call('UPPER', 'total')).toBe('TOTAL');
    expect(call('PROPER', 'this is a TITLE')).toBe('This Is A Title');
    expect(call('PROPER', '2-way street')).toBe('2-Way Street');
    expect(call('PROPER', '76BudGet')).toBe('76Budget');
    expect(call('PROPER', "don't")).toBe("Don'T");
    expect(call('UPPER', true)).toBe('TRUE');
  });
  it('ASC/DBCS', () => {
    expect(call('ASC', 'ＧＲＩＤ')).toBe('GRID');
    expect(call('ASC', 'エクセル')).toBe('ｴｸｾﾙ');
    expect(call('DBCS', 'GRID')).toBe('ＧＲＩＤ');
    expect(call('DBCS', 'ｴｸｾﾙ')).toBe('エクセル');
    expect(call('ASC', 'ガ')).toBe('ｶﾞ');
    expect(call('DBCS', 'ｶﾞ')).toBe('ガ');
  });
});

describe('slicing and searching', () => {
  it('LEFT/RIGHT/MID/LEN', () => {
    expect(call('LEFT', 'Sale Price', 4)).toBe('Sale');
    expect(call('LEFT', 'Sweden')).toBe('S');
    expect(call('LEFT', 'abc', -1)).toEqual(Err.value);
    expect(call('LEFT', 'abc', 10)).toBe('abc');
    expect(call('RIGHT', 'Sale Price', 5)).toBe('Price');
    expect(call('RIGHT', 'Stock Number')).toBe('r');
    expect(call('RIGHT', 'abc', 0)).toBe('');
    expect(call('MID', 'Fluid Flow', 1, 5)).toBe('Fluid');
    expect(call('MID', 'Fluid Flow', 7, 20)).toBe('Flow');
    expect(call('MID', 'Fluid Flow', 20, 5)).toBe('');
    expect(call('MID', 'Fluid Flow', 0, 5)).toEqual(Err.value);
    expect(call('LEN', 'Phoenix, AZ')).toBe(11);
    expect(call('LEN', null)).toBe(0);
    expect(call('LEN', '     One   ')).toBe(11);
    expect(call('LEN', 1.5)).toBe(3);
    expect(call('LEFT', Err.na, 1)).toEqual(Err.na);
  });
  it('FIND/SEARCH', () => {
    expect(call('FIND', 'M', 'Miriam McGovern')).toBe(1);
    expect(call('FIND', 'm', 'Miriam McGovern')).toBe(6);
    expect(call('FIND', 'M', 'Miriam McGovern', 3)).toBe(8);
    expect(call('FIND', 'x', 'abc')).toEqual(Err.value);
    expect(call('FIND', '', 'abc', 2)).toBe(2);
    expect(call('FIND', 'a', 'abc', 0)).toEqual(Err.value);
    expect(call('FIND', 'a', 'abc', 5)).toEqual(Err.value);
    expect(call('FIND', '*', 'a*c')).toBe(2);
    expect(call('SEARCH', 'e', 'Statements', 6)).toBe(7);
    expect(call('SEARCH', 'margin', 'Profit Margin')).toBe(8);
    expect(call('SEARCH', 'p?o', 'Profit')).toBe(1);
    expect(call('SEARCH', 'f*t', 'Profit')).toBe(4);
    expect(call('SEARCH', '~*', 'a*c')).toBe(2);
    expect(call('SEARCH', '~?', 'what?')).toBe(5);
    expect(call('SEARCH', '.', 'abc')).toEqual(Err.value);
  });
  it('REPLACE/SUBSTITUTE/REPT/EXACT', () => {
    expect(call('REPLACE', 'abcdefghijk', 6, 5, '*')).toBe('abcde*k');
    expect(call('REPLACE', '2009', 3, 2, '10')).toBe('2010');
    expect(call('REPLACE', '123456', 1, 3, '@')).toBe('@456');
    expect(call('REPLACE', 'abc', 0, 1, 'x')).toEqual(Err.value);
    expect(call('SUBSTITUTE', 'Sales Data', 'Sales', 'Cost')).toBe('Cost Data');
    expect(call('SUBSTITUTE', 'Quarter 1, 2008', '1', '2', 1)).toBe('Quarter 2, 2008');
    expect(call('SUBSTITUTE', 'Quarter 1, 2011', '1', '2', 3)).toBe('Quarter 1, 2012');
    expect(call('SUBSTITUTE', 'aaaa', 'aa', 'b', 2)).toBe('aab');
    expect(call('SUBSTITUTE', 'abc', 'x', 'y', 0)).toEqual(Err.value);
    expect(call('SUBSTITUTE', 'abc', '', 'y')).toBe('abc');
    expect(call('REPT', '*-', 3)).toBe('*-*-*-');
    expect(call('REPT', '-', 10)).toBe('----------');
    expect(call('REPT', 'ab', 20000)).toEqual(Err.value);
    expect(call('REPT', 'ab', -1)).toEqual(Err.value);
    expect(call('EXACT', 'word', 'word')).toBe(true);
    expect(call('EXACT', 'Word', 'word')).toBe(false);
    expect(call('EXACT', 1, '1')).toBe(true);
  });
});

describe('joining', () => {
  it('CONCAT/CONCATENATE/TEXTJOIN', () => {
    expect(call('CONCAT', ref(0, 0, 0, 0, 3))).toBe('abc');
    expect(call('CONCAT', 'The', ' ', 'sun')).toBe('The sun');
    expect(call('CONCAT', ref(0, 0, 0, 1, 3))).toEqual(Err.div0);
    expect(call('CONCATENATE', 'Stream population for ', 'brook trout', ' is ', 32, '/mile')).toBe('Stream population for brook trout is 32/mile');
    expect(call('CONCATENATE', true, 1.5)).toBe('TRUE1.5');
    expect(call('TEXTJOIN', ', ', true, ref(0, 0, 0, 0, 3))).toBe('a, b, c');
    expect(call('TEXTJOIN', ', ', false, ref(0, 0, 0, 0, 3))).toBe('a, b, , c');
    expect(call('TEXTJOIN', [[',', ';']], true, [['a', 'b', 'c', 'd']])).toBe('a,b;c,d');
    expect(call('TEXTJOIN', '', true, 'x', 'y')).toBe('xy');
    expect(call('CONCAT', 'a'.repeat(20000), 'b'.repeat(20000))).toEqual(Err.value);
  });
  it('T/VALUE/NUMBERVALUE/VALUETOTEXT/ARRAYTOTEXT', () => {
    expect(call('T', 'Rainfall')).toBe('Rainfall');
    expect(call('T', 19)).toBe('');
    expect(call('T', true)).toBe('');
    expect(call('VALUE', '$1,000')).toBe(1000);
    expect(call('VALUE', '16:48:00')).toBeCloseTo(0.7, 12);
    expect(call('VALUE', '50%')).toBe(0.5);
    expect(call('VALUE', '(12)')).toBe(-12);
    expect(call('VALUE', 'abc')).toEqual(Err.value);
    expect(call('VALUE', true)).toEqual(Err.value);
    expect(call('VALUE', null)).toBe(0);
    expect(call('NUMBERVALUE', '2.500,27', ',', '.')).toBe(2500.27);
    expect(call('NUMBERVALUE', '3.5%')).toBe(0.035);
    expect(call('NUMBERVALUE', '3 . 5%')).toBe(0.035);
    expect(call('NUMBERVALUE', '')).toBe(0);
    expect(call('NUMBERVALUE', '1,2.5', ',', ',')).toEqual(Err.value);
    expect(call('NUMBERVALUE', '1.2.3')).toEqual(Err.value);
    expect(call('VALUETOTEXT', 'Hello')).toBe('Hello');
    expect(call('VALUETOTEXT', 'Hello', 1)).toBe('"Hello"');
    expect(call('VALUETOTEXT', true, 1)).toBe('TRUE');
    expect(call('VALUETOTEXT', Err.value)).toBe('#VALUE!');
    expect(call('VALUETOTEXT', 1234.01234, 1)).toBe('1234.01234');
    expect(call('VALUETOTEXT', 'x', 2)).toEqual(Err.value);
    expect(call('ARRAYTOTEXT', [[true, Err.value], [1234.01234, 'Seattle']])).toBe('TRUE, #VALUE!, 1234.01234, Seattle');
    expect(call('ARRAYTOTEXT', [[true, Err.value], [1234.01234, 'Seattle']], 1)).toBe('{TRUE,#VALUE!;1234.01234,"Seattle"}');
  });
});

describe('number formatting', () => {
  it('DOLLAR/FIXED', () => {
    expect(call('DOLLAR', 1234.567, 2)).toBe('$1,234.57');
    expect(call('DOLLAR', 1234.567, -2)).toBe('$1,200');
    expect(call('DOLLAR', -1234.567, -2)).toBe('($1,200)');
    expect(call('DOLLAR', -0.123, 4)).toBe('($0.1230)');
    expect(call('DOLLAR', 99.888)).toBe('$99.89');
    expect(call('DOLLAR', -0.001)).toBe('$0.00');
    expect(call('FIXED', 1234.567, 1)).toBe('1,234.6');
    expect(call('FIXED', 1234.567, -1)).toBe('1,230');
    expect(call('FIXED', -1234.567, -1, true)).toBe('-1230');
    expect(call('FIXED', 44.332)).toBe('44.33');
    expect(call('FIXED', 2.5, 0)).toBe('3');
    expect(call('FIXED', 1.005, 2)).toBe('1.01');
    expect(call('FIXED', 1234567, -3)).toBe('1,235,000');
    expect(call('FIXED', 'abc')).toEqual(Err.value);
  });
  it('BAHTTEXT', () => {
    expect(call('BAHTTEXT', 1234)).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน');
    expect(call('BAHTTEXT', 0)).toBe('ศูนย์บาทถ้วน');
    expect(call('BAHTTEXT', 21)).toBe('ยี่สิบเอ็ดบาทถ้วน');
    expect(call('BAHTTEXT', 11.25)).toBe('สิบเอ็ดบาทยี่สิบห้าสตางค์');
    expect(call('BAHTTEXT', 1000001)).toBe('หนึ่งล้านเอ็ดบาทถ้วน');
    expect(call('BAHTTEXT', 1)).toBe('หนึ่งบาทถ้วน');
  });
});

describe('TEXTBEFORE/TEXTAFTER/TEXTSPLIT', () => {
  const s = 'Little red riding hood';
  it('TEXTBEFORE/TEXTAFTER', () => {
    expect(call('TEXTBEFORE', s, 'Red')).toEqual(Err.na);
    expect(call('TEXTBEFORE', s, 'red')).toBe('Little ');
    expect(call('TEXTBEFORE', s, 'Red', 1, 1)).toBe('Little ');
    expect(call('TEXTBEFORE', s, ' ', -1)).toBe('Little red riding');
    expect(call('TEXTBEFORE', s, ' ', 2)).toBe('Little red');
    expect(call('TEXTBEFORE', s, ' ', 4)).toEqual(Err.na);
    expect(call('TEXTBEFORE', s, ' ', 4, 0, 1)).toBe(s);
    expect(call('TEXTBEFORE', s, ' ', 4, 0, 0, 'none')).toBe('none');
    expect(call('TEXTBEFORE', s, ' ', 0)).toEqual(Err.value);
    expect(call('TEXTBEFORE', s, '')).toBe('');
    expect(call('TEXTAFTER', s, 'hood')).toBe('');
    expect(call('TEXTAFTER', s, 'red')).toBe(' riding hood');
    expect(call('TEXTAFTER', s, 'basket')).toEqual(Err.na);
    expect(call('TEXTAFTER', s, 'RED', 1, 1)).toBe(' riding hood');
    expect(call('TEXTAFTER', s, ' ', -1)).toBe('hood');
    expect(call('TEXTAFTER', s, ' ', -4, 0, 1)).toBe(s);
    expect(call('TEXTAFTER', s, [['x', 'red']])).toBe(' riding hood');
    expect(call('TEXTAFTER', s, '')).toBe(s);
  });
  it('TEXTSPLIT', () => {
    expect(call('TEXTSPLIT', 'Dakota Lennon Sanchez', ' ')).toEqual([['Dakota', 'Lennon', 'Sanchez']]);
    expect(call('TEXTSPLIT', '1,2,3;4,5,6', ',', ';')).toEqual([['1', '2', '3'], ['4', '5', '6']]);
    expect(call('TEXTSPLIT', 'Do. Or do not. There is no try. -Anonymous', '.')).toEqual([['Do', ' Or do not', ' There is no try', ' -Anonymous']]);
    expect(call('TEXTSPLIT', 'Do. Or do not. There is no try. -Anonymous', [['.', '-']])).toEqual([['Do', ' Or do not', ' There is no try', ' ', 'Anonymous']]);
    expect(call('TEXTSPLIT', 'Do. Or do not. There is no try. -Anonymous', null, '.')).toEqual([['Do'], [' Or do not'], [' There is no try'], [' -Anonymous']]);
    expect(call('TEXTSPLIT', 'Do. Or do not. There is no try. -Anonymous', [['.', '-']], null, true)).toEqual([['Do', ' Or do not', ' There is no try', ' ', 'Anonymous']]);
    expect(call('TEXTSPLIT', 'a,,b', ',', null, true)).toEqual([['a', 'b']]);
    expect(call('TEXTSPLIT', 'a,b;c', ',', ';')).toEqual([['a', 'b'], ['c', Err.na]]);
    expect(call('TEXTSPLIT', 'a,b;c', ',', ';', false, 0, '-')).toEqual([['a', 'b'], ['c', '-']]);
    expect(call('TEXTSPLIT', 'aXbxc', 'x', null, false, 1)).toEqual([['a', 'b', 'c']]);
    expect(call('TEXTSPLIT', 'abc', '')).toEqual(Err.value);
  });
});

describe('regular expressions', () => {
  it('REGEXTEST/REGEXEXTRACT/REGEXREPLACE', () => {
    expect(call('REGEXTEST', 'alpha', '[a-z]')).toBe(true);
    expect(call('REGEXTEST', 'ALPHA', '[a-z]')).toBe(false);
    expect(call('REGEXTEST', 'ALPHA', '[a-z]', 1)).toBe(true);
    expect(call('REGEXTEST', '(378) 555-4195', '^\\(\\d{3}\\) \\d{3}-\\d{4}$')).toBe(true);
    expect(call('REGEXTEST', 'x', '(')).toEqual(Err.value);
    expect(call('REGEXTEST', 'ABC', '(?i)abc')).toBe(true);
    expect(call('REGEXEXTRACT', 'DylanWilliams', '[A-Z][a-z]+')).toBe('Dylan');
    expect(call('REGEXEXTRACT', 'DylanWilliams', '[A-Z][a-z]+', 1)).toEqual([['Dylan'], ['Williams']]);
    expect(call('REGEXEXTRACT', 'Sonia Rees (378) 555-4195', '\\((\\d{3})\\) (\\d{3})-(\\d{4})', 2)).toEqual([['378', '555', '4195']]);
    expect(call('REGEXEXTRACT', 'abc', '\\d')).toEqual(Err.na);
    expect(call('REGEXEXTRACT', 'abc', '\\w', 3)).toEqual(Err.value);
    expect(call('REGEXREPLACE', 'Sonia Brown', '[A-Z]', '*')).toBe('*onia *rown');
    expect(call('REGEXREPLACE', 'Sonia Brown', '([A-Z])(\\w+) ([A-Z])(\\w+)', '$3$4, $1$2')).toBe('Brown, Sonia');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', 2)).toBe('a1b#c3');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', -1)).toBe('a1b2c#');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', 5)).toBe('a1b2c3');
    expect(call('REGEXREPLACE', 'ABC', 'b', 'x', 0, 1)).toBe('AxC');
    expect(call('REGEXREPLACE', 'ab', '(?P<x>a)', '${x}${x}')).toBe('aab');
  });
});

describe('misc', () => {
  it('PHONETIC/online functions/scalar flags', () => {
    expect(call('PHONETIC', ref(0, 0, 0, 0, 1))).toBe('ab');
    // TRANSLATE and DETECTLANGUAGE live with the other online functions (tests/fn-web.test.ts).
    expect(defs.some((d) => d.name === 'TRANSLATE' || d.name === 'DETECTLANGUAGE')).toBe(false);
    expect(defs.find((d) => d.name === 'LEFT')!.scalar).toBe(true);
    expect(defs.find((d) => d.name === 'CONCAT')!.scalar).toBeUndefined();
    expect(new Set(defs.map((d) => d.name)).size).toBe(defs.length);
  });
});
