import { useMemo } from 'react';
import { formatValue, isDateFormat } from '../../formats/numfmt';
import type { Scalar } from '../../engine/types';
import { Field } from './Dialog';
import { Check } from './shared';

export type NumCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percent' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom';

export const CATEGORIES: [NumCategory, string][] = [
  ['general', 'General'], ['number', 'Number'], ['currency', 'Currency'], ['accounting', 'Accounting'], ['date', 'Date'], ['time', 'Time'],
  ['percent', 'Percentage'], ['fraction', 'Fraction'], ['scientific', 'Scientific'], ['text', 'Text'], ['special', 'Special'], ['custom', 'Custom'],
];

const SYMBOLS: [string, string][] = [['', 'None'], ['$', '$'], ['€', '€'], ['£', '£'], ['¥', '¥'], ['₹', '₹'], ['CHF', 'CHF'], ['SAR', 'SAR'], ['AED', 'AED']];
export const DATE_FORMATS = ['m/d/yyyy', 'd/m/yyyy', 'yyyy-mm-dd', 'm/d/yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'd mmmm yyyy', 'mmmm yyyy', 'm/d/yyyy h:mm', 'yyyy-mm-dd hh:mm:ss'];
export const TIME_FORMATS = ['h:mm', 'h:mm AM/PM', 'h:mm:ss', 'h:mm:ss AM/PM', 'hh:mm:ss', 'mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yyyy h:mm AM/PM'];
const FRACTIONS: [string, string][] = [['# ?/?', 'Up to one digit'], ['# ??/??', 'Up to two digits'], ['# ???/???', 'Up to three digits'], ['# ?/2', 'Halves'], ['# ?/4', 'Quarters'], ['# ?/8', 'Eighths'], ['# ??/16', 'Sixteenths'], ['# ?/10', 'Tenths'], ['# ??/100', 'Hundredths']];
const SPECIAL: [string, string][] = [['00000', 'Postal code (5 digits)'], ['00000-0000', 'Postal code + 4'], ['[<=9999999]###-####;(###) ###-####', 'Phone number'], ['000-00-0000', 'ID number (3-2-4)'], ['0000 0000 0000 0000', 'Groups of four']];
const CUSTOM_STARTERS = ['General', '0', '0.00', '#,##0', '#,##0.00', '#,##0;[Red]-#,##0', '0%', '0.00%', '0.00E+00', '# ?/?', 'm/d/yyyy', 'd-mmm-yy', 'h:mm AM/PM', '[h]:mm:ss', '@', '"Total: "0', '[Blue]#,##0;[Red]-#,##0;0', '0.0,,"M"', '#,##0.00 "kg"'];

export interface NumOpts { cat: NumCategory; dec: number; thousands: boolean; neg: number; symbol: string; pattern: string; custom: string }

const digits = (dec: number, thousands: boolean) => (thousands ? '#,##0' : '0') + (dec > 0 ? `.${'0'.repeat(dec)}` : '');
const quoteSym = (sym: string) => (sym ? (sym.length === 1 && sym === '$' ? '"$"' : `"${sym}"`) : '');
const negative = (base: string, neg: number) => [base, `${base};[Red]${base}`, `${base}_);(${base})`, `${base}_);[Red](${base})`][neg] ?? base;

/** Builds a format code from the dialog's options. */
export const buildCode = (o: NumOpts): string => {
  switch (o.cat) {
    case 'general': return 'General';
    case 'number': return negative(digits(o.dec, o.thousands), o.neg);
    case 'currency': return negative(`${quoteSym(o.symbol)}${digits(o.dec, true)}`, o.neg);
    case 'accounting': {
      const d = digits(o.dec, true), s = o.symbol ? `${quoteSym(o.symbol)}* ` : '* ';
      return `_(${s}${d}_);_(${s}\\(${d}\\);_(${s}"-"${o.dec ? '??' : ''}_);_(@_)`;
    }
    case 'percent': return `${digits(o.dec, false)}%`;
    case 'scientific': return `${digits(o.dec, false)}E+00`;
    case 'text': return '@';
    case 'date': case 'time': case 'fraction': case 'special': return o.pattern;
    case 'custom': return o.custom || 'General';
  }
};

/** Best guess at which category and options produced an existing code. */
export const parseCode = (code: string): NumOpts => {
  const base: NumOpts = { cat: 'custom', dec: 2, thousands: false, neg: 0, symbol: '$', pattern: '', custom: code };
  const decOf = (s: string) => (/\.(0+)/.exec(s)?.[1]?.length ?? 0);
  if (!code || code === 'General') return { ...base, cat: 'general' };
  if (code === '@') return { ...base, cat: 'text' };
  if (DATE_FORMATS.includes(code)) return { ...base, cat: 'date', pattern: code };
  if (TIME_FORMATS.includes(code)) return { ...base, cat: 'time', pattern: code };
  if (FRACTIONS.some(([c]) => c === code)) return { ...base, cat: 'fraction', pattern: code };
  if (SPECIAL.some(([c]) => c === code)) return { ...base, cat: 'special', pattern: code };
  if (/^0(\.0+)?%$/.test(code)) return { ...base, cat: 'percent', dec: decOf(code) };
  if (/^0(\.0+)?E\+00$/.test(code)) return { ...base, cat: 'scientific', dec: decOf(code) };
  if (code.startsWith('_(')) {
    const sym = /"([^"]+)"\*/.exec(code)?.[1] ?? '';
    return { ...base, cat: 'accounting', dec: decOf(code), symbol: sym };
  }
  const first = code.split(';')[0]!.replace(/_\)/, '');
  const neg = !code.includes(';') ? 0 : code.includes('[Red](') ? 3 : code.includes(';(') ? 2 : code.includes('[Red]') ? 1 : -1;
  const sym = /^"([^"]+)"/.exec(first)?.[1];
  const rest = sym !== undefined ? first.replace(/^"[^"]+"/, '') : first;
  if (neg >= 0 && /^(#,##)?0(\.0+)?$/.test(rest)) {
    if (sym !== undefined) return { ...base, cat: 'currency', dec: decOf(rest), symbol: sym, neg };
    return { ...base, cat: 'number', dec: decOf(rest), thousands: rest.startsWith('#,##'), neg };
  }
  if (isDateFormat(code)) return { ...base, cat: /[dy]/i.test(code) ? 'date' : 'time', pattern: code };
  return base;
};

export const NumberTab = ({ o, set, sample, date1904, formats }: { o: NumOpts; set: (p: Partial<NumOpts>) => void; sample: Scalar; date1904: boolean; formats: string[] }) => {
  const code = buildCode(o);
  const shown = useMemo(() => {
    try { return formatValue(sample ?? 1234.5, code, { date1904 }); } catch { return { text: 'This format code has a problem.', color: undefined }; }
  }, [sample, code, date1904]);
  const negSamples = ['-1,234.10', '1,234.10', '(1,234.10)', '(1,234.10)'];
  const pick = (list: [string, string][]) => (
    <div className="b-list" role="listbox">
      {list.map(([c, label]) => <button key={c} role="option" aria-selected={o.pattern === c} className={o.pattern === c ? 'on' : ''} onClick={() => set({ pattern: c })}>
        <span>{label}</span><span className="muted small">{formatValue(typeof sample === 'number' ? sample : 45678.5625, c, { date1904 }).text}</span></button>)}
    </div>
  );
  const decimals = <Field label="Decimal places"><input type="number" min={0} max={30} value={o.dec} onChange={(e) => set({ dec: Math.max(0, Math.min(30, Math.trunc(+e.target.value || 0))) })} /></Field>;
  const symbol = <Field label="Symbol"><select value={o.symbol} onChange={(e) => set({ symbol: e.target.value })}>{SYMBOLS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>;
  const negatives = (
    <Field label="Negative numbers">
      <div className="b-list" role="listbox">{negSamples.map((t, i) => <button key={i} className={o.neg === i ? 'on' : ''} aria-selected={o.neg === i} role="option" onClick={() => set({ neg: i })} style={{ color: i % 2 ? '#d92d20' : undefined }}>{t}</button>)}</div>
    </Field>
  );
  return (
    <div className="b-split">
      <div className="b-list b-cats" role="listbox" aria-label="Category">
        {CATEGORIES.map(([c, l]) => <button key={c} role="option" aria-selected={o.cat === c} className={o.cat === c ? 'on' : ''} onClick={() => {
          const p: Partial<NumOpts> = { cat: c };
          if (c === 'date' && !DATE_FORMATS.includes(o.pattern)) p.pattern = DATE_FORMATS[0];
          if (c === 'time' && !TIME_FORMATS.includes(o.pattern)) p.pattern = TIME_FORMATS[0];
          if (c === 'fraction' && !FRACTIONS.some(([x]) => x === o.pattern)) p.pattern = FRACTIONS[0]![0];
          if (c === 'special' && !SPECIAL.some(([x]) => x === o.pattern)) p.pattern = SPECIAL[0]![0];
          if (c === 'custom') p.custom = code;
          set(p);
        }}>{l}</button>)}
      </div>
      <div className="b-col">
        <Field label="Sample"><div className="preview b-sample" style={{ color: shown.color }}>{shown.text || ' '}</div></Field>
        {o.cat === 'general' && <p className="muted small">No special format. Numbers show as typed, up to 11 digits.</p>}
        {o.cat === 'number' && <>{decimals}<Check label="Use thousands separator" checked={o.thousands} onChange={(v) => set({ thousands: v })} />{negatives}</>}
        {o.cat === 'currency' && <><div className="grid2">{decimals}{symbol}</div>{negatives}</>}
        {o.cat === 'accounting' && <><div className="grid2">{decimals}{symbol}</div><p className="muted small">Lines up symbols and decimal points in a column.</p></>}
        {o.cat === 'percent' && <>{decimals}<p className="muted small">Multiplies by 100 and adds a percent sign.</p></>}
        {o.cat === 'scientific' && decimals}
        {o.cat === 'text' && <p className="muted small">Shows exactly what you type, even numbers.</p>}
        {o.cat === 'date' && pick(DATE_FORMATS.map((d) => [d, d]))}
        {o.cat === 'time' && pick(TIME_FORMATS.map((d) => [d, d]))}
        {o.cat === 'fraction' && pick(FRACTIONS)}
        {o.cat === 'special' && pick(SPECIAL)}
        {o.cat === 'custom' && <>
          <Field label="Format code"><input className="mono" value={o.custom} spellCheck={false} onChange={(e) => set({ custom: e.target.value })} /></Field>
          <div className="b-list b-short" role="listbox">{[...new Set([...formats, ...CUSTOM_STARTERS])].map((c) => <button key={c} className={`mono${o.custom === c ? ' on' : ''}`} onClick={() => set({ custom: c })}>{c}</button>)}</div>
          <p className="muted small">Up to four parts separated by semicolons: positive; negative; zero; text.</p>
        </>}
      </div>
    </div>
  );
};
