import { describe, expect, test } from 'vitest';
import { builtinFormat, formatGeneral, formatValue, guessFormat, isDateFormat } from '../src/formats/numfmt';
import { CellError } from '../src/engine/types';
import { defs } from '../src/engine/functions/textfmt';
import { makeCtx } from './fnctx';

const t = (v: number | string, code: string): string => formatValue(v, code).text;

// [value, format, expected text] as shown in a wide-enough cell.
const CASES: [number | string, string, string][] = [
  // General
  [0, 'General', '0'], [1, 'General', '1'], [-1.5, 'General', '-1.5'], [0.1 + 0.2, 'General', '0.3'],
  [1234567890, 'General', '1234567890'], [12345678901, 'General', '12345678901'], [123456789012, 'General', '1.23457E+11'],
  [0.000123, 'General', '0.000123'], [0.0000123, 'General', '1.23E-05'], [1 / 3, 'General', '0.333333333'],
  [2 / 3, 'General', '0.666666667'], [-2 / 3, 'General', '-0.66666667'], [1e21, 'General', '1E+21'], [123.456789012, 'General', '123.456789'],
  ['text', 'General', 'text'],
  // plain digits
  [1234.567, '0', '1235'], [1234.5, '0', '1235'], [-1234.5, '0', '-1235'], [0.5, '0', '1'], [2.5, '0', '3'], [-0.4, '0', '-0'],
  [1234.567, '0.00', '1234.57'], [1.005, '0.00', '1.01'], [1.0049999, '0.00', '1.00'], [0.125, '0.00', '0.13'], [2.675, '0.00', '2.68'],
  [1234.567, '#,##0', '1,235'], [1234567.891, '#,##0.00', '1,234,567.89'], [-1234567.891, '#,##0.00', '-1,234,567.89'],
  [0, '#,##0.00', '0.00'], [0.5, '#,##0', '1'], [999.999, '#,##0.00', '1,000.00'], [123, '00000', '00123'], [5, '000.0', '005.0'],
  [0.5, '#.##', '.5'], [5, '#.##', '5.'], [0, '#', ''], [0, '#.#', '.'], [12.3, '???.???', ' 12.3  '], [1.5, '0.0?', '1.5 '],
  [12345.678, '#,###.#', '12,345.7'], [1234567, '#,##0,', '1,235'], [1234567, '0.0,,', '1.2'], [1234567, '#,##0,,"M"', '1M'],
  [100, '0.00,', '0.10'], [1234, '0,000', '1,234'], [7, '"Total: "0', 'Total: 7'], [7, '0" units"', '7 units'], [7, '\\#0', '#7'],
  [3, '0_)', '3 '], [3, '(0)', '(3)'], [12, '0 "kg"', '12 kg'], [1234.5, '[$€-x-euro2] #,##0.00', '€ 1,234.50'], [1234.5, '[$-409]#,##0.00', '1,234.50'],
  [1234.5, '#,##0.00 [$USD]', '1,234.50 USD'], [123456789, '#,##0', '123,456,789'], [1e15, '#,##0', '1,000,000,000,000,000'],
  [0.000001234, '0.000000000', '0.000001234'], [123456789012345, '0', '123456789012345'],
  // percent
  [0.5, '0%', '50%'], [0.1234, '0.00%', '12.34%'], [-0.05, '0%', '-5%'], [1.5, '0%', '150%'], [0.00005, '0.0%', '0.0%'], [0.12345, '0.0%', '12.3%'],
  // scientific
  [12345, '0.00E+00', '1.23E+04'], [0.000123, '0.00E+00', '1.23E-04'], [0, '0.00E+00', '0.00E+00'], [-12345, '0.00E+00', '-1.23E+04'],
  [1e100, '0.00E+00', '1.00E+100'], [9.999, '0.00E+00', '1.00E+01'], [12345, '0E+0', '1E+4'], [12345, '##0.0E+0', '12.3E+3'],
  [1234567, '##0.0E+0', '1.2E+6'], [0.00012, '##0.0E+0', '120.0E-6'], [12345, '0.0E-0', '1.2E4'], [0.012, '0.0E-0', '1.2E-2'],
  [123456, '00.00E+00', '12.35E+04'],
  // fractions
  [1.25, '# ?/?', '1 1/4'], [0.5, '# ?/?', ' 1/2'], [2, '# ?/?', '2    '], [3.14159, '# ??/??', '3 14/99'], [3.14159, '# ?/?', '3 1/7'], [3.14159, '# ???/???', '3  16/113'],
  [0.75, '# ??/16', ' 12/16'], [1.3, '# ?/8', '1 2/8'], [1.25, '?/?', '5/4'], [-1.25, '# ?/?', '-1 1/4'], [0.333, '# ?/?', ' 1/3'],
  [5.5, '0 ?/?', '5 1/2'], [0, '# ?/?', '0    '], [0.99, '# ?/?', '1    '], [2.5, '# ?/2', '2 1/2'], [0.1, '# ?/10', ' 1/10'],
  // sections
  [5, '0;(0)', '5'], [-5, '0;(0)', '(5)'], [0, '0;(0)', '0'], [0, '0;-0;"zero"', 'zero'], [-5, '0;-0;"zero"', '-5'],
  ['abc', '0;-0;0;"[" @ "]"', '[ abc ]'], ['abc', '0.00', 'abc'], ['abc', '"pre "@', 'pre abc'], [12, '"pre "@', '12'],
  [5, '0;;', '5'], [-5, '0;;', ''], [0, '0;;', ''], [-5, ';;', ''], [5, '"pos";"neg";"zero"', 'pos'], [-5, '"pos";"neg";"zero"', 'neg'],
  [-1234.567, '#,##0.00_);(#,##0.00)', '(1,234.57)'], [1234.567, '#,##0.00_);(#,##0.00)', '1,234.57 '], [-0.001, '0.00;-0.00', '-0.00'],
  [3, 'General;"neg"', '3'], [-3, 'General;"neg"', 'neg'], [-3, '0;General', '3'], [12, '"Num "General', 'Num 12'],
  // conditions
  [150, '[>100]"big";"small"', 'big'], [50, '[>100]"big";"small"', 'small'], [-50, '[>100]"big";"small"', '-small'],
  [150, '[>=100]0.0;[<0]"neg";0', '150.0'], [-5, '[>=100]0.0;[<0]"neg";0', 'neg'], [5, '[>=100]0.0;[<0]"neg";0', '5'],
  [1234567, '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', '1.2M'], [12345, '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', '12.3K'],
  [999, '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', '999'], [-5, '[<0]"minus "0;0', 'minus 5'],
  // dates
  [45366, 'm/d/yyyy', '3/15/2024'], [45366, 'mm/dd/yy', '03/15/24'], [45366, 'd-mmm-yy', '15-Mar-24'], [45366, 'dddd, mmmm d, yyyy', 'Friday, March 15, 2024'],
  [45366, 'ddd', 'Fri'], [45366, 'mmmmm', 'M'], [45366, 'yyyy-mm-dd', '2024-03-15'], [45366, 'd mmm', '15 mmm'.replace('mmm', 'Mar')],
  [1, 'm/d/yyyy', '1/1/1900'], [60, 'm/d/yyyy', '2/29/1900'], [61, 'm/d/yyyy', '3/1/1900'], [0, 'm/d/yyyy', '1/0/1900'],
  [2958465, 'yyyy-mm-dd', '9999-12-31'], [45366.75, 'm/d/yyyy h:mm', '3/15/2024 18:00'], [36526, 'yy', '00'], [45366, 'e', '2024'],
  [45366, '[$-409]mmmm d, yyyy;@', 'March 15, 2024'], [45366, 'dd"th" mmm', '15th Mar'], [45366, 'dd\\/mm\\/yyyy', '15/03/2024'],
  // times
  [0.5, 'h:mm', '12:00'], [0.75, 'h:mm AM/PM', '6:00 PM'], [0.25, 'h:mm AM/PM', '6:00 AM'], [0, 'h:mm AM/PM', '12:00 AM'], [0.5, 'h:mm AM/PM', '12:00 PM'],
  [0.6041666667, 'hh:mm:ss', '14:30:00'], [0.6041666667, 'h:mm:ss A/P', '2:30:00 P'], [0.6041666667, 'h:mm a/p', '2:30 p'],
  [0.000694444, 'mm:ss', '01:00'], [0.5000057, 'hh:mm:ss.000', '12:00:00.492'], [0.50000579, 'h:mm:ss.0', '12:00:00.5'], [0.50000579, 'ss.00', '00.50'],
  [0.99998, 'hh:mm:ss', '23:59:58'], [0.999999, 'hh:mm:ss', '00:00:00'], [45366.6041666667, 'm/d/yyyy h:mm AM/PM', '3/15/2024 2:30 PM'],
  [0.0104166667, 'h:mm', '0:15'], [0.5, 'h "hours" m "min"', '12 hours 0 min'], [0.52, 'h:m', '12:28'],
  // m disambiguation: month unless after h or before s
  [45366.6041666667, 'mm', '03'], [45366.6041666667, 'hh:mm', '14:30'], [45366.6041666667, 'mm:ss', '30:00'], [45366.6041666667, 'yyyy-mm-dd hh:mm', '2024-03-15 14:30'],
  // elapsed
  [1.5, '[h]:mm:ss', '36:00:00'], [2.25, '[h]:mm', '54:00'], [0.0625, '[mm]:ss', '90:00'], [0.01, '[ss]', '864'], [1.000011574, '[h]:mm:ss', '24:00:01'],
  [3.5, '[h]', '84'], [-1, '[h]:mm', '########'.slice(0, 11).padEnd(11, '#')],
  // colors and fill keep the text
  [-5, '0;[Red]0', '5'], [5, '[Blue]0', '5'], [5, '[Color10]0', '5'],
  // accounting
  [1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', ' $1,234.50 '], [-1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', ' $(1,234.50)'],
  [0, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', ' $-   '], ['abc', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', ' abc '],
  [1234.5, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', ' 1,235 '], [0, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', ' - '],
  [1234.5, '"$"#,##0_);[Red]\\("$"#,##0\\)', '$1,235 '], [-1234.5, '"$"#,##0_);[Red]\\("$"#,##0\\)', '($1,235)'],
];

describe('formatValue', () => {
  test.each(CASES)('%s with %s', (v, code, expected) => {
    expect(t(v, code)).toBe(expected);
  });

  test('colors', () => {
    expect(formatValue(-5, '0;[Red]0').color).toBe('#FF0000');
    expect(formatValue(5, '0;[Red]0').color).toBeUndefined();
    expect(formatValue(5, '[Blue]0').color).toBe('#0000FF');
    expect(formatValue(5, '[Green]0').color).toBe('#00FF00');
    expect(formatValue(5, '[Color10]0').color).toBe('#008000');
    expect(formatValue(5, '[COLOR3]0').color).toBe('#FF0000');
    expect(formatValue(150, '[Red][>100]0;[Blue]0').color).toBe('#FF0000');
    expect(formatValue('x', '0;0;0;[Magenta]@').color).toBe('#FF00FF');
  });

  test('fill character and alignment', () => {
    const r = formatValue(1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
    expect(r.fill).toBe(' ');
    expect(r.align).toBe('right');
    expect(formatValue(5, '0*-').fill).toBe('-');
    expect(formatValue(5, '**0').fill).toBe('*');
    expect(formatValue('a', '@').align).toBe('left');
    expect(formatValue(1, '0').fill).toBeUndefined();
  });

  test('non-numbers', () => {
    expect(formatValue(true, '0.00').text).toBe('TRUE');
    expect(formatValue(null, '0.00').text).toBe('');
    expect(formatValue(new CellError('#N/A'), '0.00').text).toBe('#N/A');
  });

  test('dates out of range show ####', () => {
    expect(t(-1, 'm/d/yyyy')).toMatch(/^#+$/);
    expect(t(2958466, 'm/d/yyyy')).toMatch(/^#+$/);
    expect(formatValue(-1, 'm/d/yyyy', { width: 5 }).text).toBe('#####');
  });

  test('1904 date system', () => {
    expect(formatValue(0, 'm/d/yyyy', { date1904: true }).text).toBe('1/1/1904');
    expect(formatValue(43904, 'yyyy-mm-dd', { date1904: true }).text).toBe('2024-03-15');
  });

  test('system date/time tags', () => {
    expect(t(45366, '[$-F800]dddd, mmmm dd, yyyy')).toBe('Friday, March 15, 2024');
    expect(t(0.6041666667, '[$-F400]h:mm:ss AM/PM')).toBe('2:30:00 PM');
  });
});

describe('formatGeneral widths', () => {
  const G: [number, number, string][] = [
    [123456789012, 11, '1.23457E+11'], [1234567890123, 11, '1.23457E+12'], [12345678901.5, 11, '12345678902'], [99999999999, 11, '99999999999'],
    [999999999999, 11, '1E+12'], [0.1 + 0.2, 11, '0.3'], [Math.PI, 11, '3.141592654'], [-Math.PI, 11, '-3.14159265'], [0.00001, 11, '1E-05'],
    [0.0001, 11, '0.0001'], [1e-10, 11, '1E-10'], [1.5e-7, 11, '1.5E-07'], [123.456, 5, '123.5'], [123456, 5, '1E+05'], [0.5, 3, '0.5'],
    [-1e-20, 11, '-1E-20'], [1.23456789e25, 11, '1.23457E+25'], [100, 11, '100'], [1e11, 11, '1E+11'],
  ];
  test.each(G)('%s at width %s', (n, w, s) => expect(formatGeneral(n, w)).toBe(s));
});

describe('builtinFormat / isDateFormat', () => {
  test('built-in ids', () => {
    expect(builtinFormat(0)).toBe('General');
    expect(builtinFormat(4)).toBe('#,##0.00');
    expect(builtinFormat(9)).toBe('0%');
    expect(builtinFormat(14)).toBe('m/d/yyyy');
    expect(builtinFormat(22)).toBe('m/d/yyyy h:mm');
    expect(builtinFormat(44)).toBe('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
    expect(builtinFormat(49)).toBe('@');
    expect(builtinFormat(100)).toBeUndefined();
    for (let i = 0; i <= 49; i++) if (builtinFormat(i)) expect(() => formatValue(1234.5, builtinFormat(i)!)).not.toThrow();
  });
  test('date detection', () => {
    for (const c of ['m/d/yyyy', 'h:mm', '[h]:mm:ss', 'yyyy', 'mmm-yy', '[$-409]d-mmm;@', 'ss.0', 'h AM/PM', '[Red]m/d/yy']) expect(isDateFormat(c), c).toBe(true);
    for (const c of ['General', '0.00', '#,##0', '"d"0', '\\m0', '@', '0.00E+00', '[Red]0', '"Date: "@']) expect(isDateFormat(c), c).toBe(false);
  });
});

describe('guessFormat', () => {
  const G: [string, string | undefined][] = [
    ['15%', '0%'], ['12.5%', '0.00%'], ['$1,234.50', '$#,##0.00'], ['$5', '$#,##0'], ['1,234', '#,##0'], ['1,234.5', '#,##0.00'],
    ['3/15/2024', 'm/d/yyyy'], ['2024-03-15', 'm/d/yyyy'], ['15-Mar', 'd-mmm'], ['Mar 15', 'd-mmm'], ['15-Mar-2024', 'd-mmm-yy'],
    ['March 15, 2024', 'd-mmm-yy'], ['Mar-24', 'd-mmm'], ['Mar-2024', 'mmm-yy'], ['14:30', 'h:mm'], ['14:30:15', 'h:mm:ss'], ['2:30 PM', 'h:mm AM/PM'],
    ['3/15/2024 14:30', 'm/d/yyyy h:mm'], ['1e5', '0.00E+00'], ['1.5E-3', '0.00E+00'], ['1/2', 'd-mmm'], ['1 1/2', '# ?/?'], ['2 3/16', '# ??/??'],
    ['123', undefined], ['-1.5', undefined], ['abc', undefined], ['', undefined], ['13/45', undefined], ['1,23', undefined],
  ];
  test.each(G)('%s', (s, f) => expect(guessFormat(s)).toBe(f));
});

describe('TEXT()', () => {
  const TEXT = defs.find((d) => d.name === 'TEXT')!;
  const ctx = makeCtx();
  const call = (v: unknown, f: unknown): unknown => TEXT.fn([v as never, f as never], ctx);
  test('formats like the cell format', () => {
    expect(call(1234.567, '#,##0.00')).toBe('1,234.57');
    expect(call(45366, 'dddd')).toBe('Friday');
    expect(call(0.5, '0%')).toBe('50%');
    expect(call('1234.5', '0.0')).toBe('1234.5');
    expect(call('abc', '0.00')).toBe('abc');
    expect(call('abc', '0;0;0;"<"@">"')).toBe('<abc>');
    expect(call(null, '0.00')).toBe('0.00');
    expect(call(true, '0')).toBe('TRUE');
    expect(call(-5, '0;(0)')).toBe('(5)');
    expect(call(5, '0*x')).toBe('5');
    expect(call(new CellError('#N/A'), '0')).toEqual(new CellError('#N/A'));
    expect(call(-1, 'yyyy')).toEqual(new CellError('#VALUE!'));
    expect(call(1234567890123, 'General')).toBe('1.23457E+12');
  });
});
