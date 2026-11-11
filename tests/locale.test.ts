import { afterEach, describe, expect, it } from 'vitest';
import {
  delocalizeFormula, delocalizeInput, displayFormat, editDateText, initLocale, localizeFormula, localizeNumberText,
  LOCALE_KEY, LOCALES, regionalFormat, separators, setLocale, setNativeDigits, setSeparators, activeLocale,
} from '../src/app/locale';
import { dateNames, tagForLcid } from '../src/formats/locales';
import { formatValue, guessFormat } from '../src/formats/numfmt';
import { parseNumber } from '../src/engine/coerce';
import { toSerial } from '../src/engine/date';
import { defs as textDefs } from '../src/engine/functions/textfmt';
import { makeCtx } from './fnctx';

afterEach(() => { setSeparators('', ''); setLocale('en-US', { nativeDigits: false }); });

const TEXT = (v: number, f: string): unknown => textDefs.find((d) => d.name === 'TEXT')!.fn([v, f], makeCtx());
const d2024 = toSerial(2024, 3, 15);

describe('locale table', () => {
  it('covers the app languages', () => {
    for (const lang of ['en', 'ar', 'de', 'es', 'fr', 'it', 'pt', 'nl', 'sv', 'pl', 'tr', 'ru', 'uk', 'zh', 'ja', 'ko', 'hi', 'id', 'he']) {
      const loc = LOCALES.find((l) => l.id.startsWith(lang));
      expect(loc, lang).toBeDefined();
      const n = dateNames(loc!.id);
      expect(n.months).toHaveLength(12);
      expect(n.days).toHaveLength(7);
      if (lang !== 'en') expect(n.months[0], lang).not.toBe('January');
    }
  });
  it('maps Windows locale ids', () => {
    expect(tagForLcid(0x0407)).toBe('de-DE');
    expect(tagForLcid(0x040c)).toBe('fr-FR');
    expect(tagForLcid(0x0c07)).toBe('de-AT');
    expect(tagForLcid(0x0007)).toBe('de-DE');
    expect(tagForLcid(0x2010401 & 0xffff)).toBe('ar-SA');
  });
  it('falls back to language, then to English', () => {
    expect(setLocale('de-AT').id).toBe('de-DE');
    expect(setLocale('xx-YY').id).toBe('en-US');
  });
});

describe('German region', () => {
  it('reads typed numbers and dates', () => {
    setLocale('de-DE');
    expect(separators).toEqual({ decimal: ',', thousands: '.' });
    expect(delocalizeInput('1.234,5')).toBe('1234.5');
    expect(parseNumber(delocalizeInput('1.234,5'))).toBe(1234.5);
    expect(parseNumber(delocalizeInput('-0,25'))).toBe(-0.25);
    expect(parseNumber(delocalizeInput('12,5 %'))).toBeCloseTo(0.125, 12);
    expect(parseNumber(delocalizeInput('1.234'))).toBe(1234);
    expect(parseNumber(delocalizeInput('15.03.2024'))).toBe(d2024);
    expect(parseNumber(delocalizeInput('15. März 2024'))).toBe(d2024);
    expect(parseNumber(delocalizeInput('3,50 €'))).toBe(3.5);
    expect(guessFormat(delocalizeInput('3,50 €'))).toBe('#,##0.00\\ [$€-407]');
    expect(guessFormat(delocalizeInput('15.03.2024'))).toBe('m/d/yyyy');
    expect(parseNumber(delocalizeInput('15.03.'))).toBe(toSerial(new Date().getFullYear(), 3, 15));
    expect(guessFormat(delocalizeInput('15.03'))).toBe('d-mmm');
  });
  it('shows numbers, dates and names the German way', () => {
    setLocale('de-DE');
    expect(localizeNumberText('1,234.5', '#,##0.0')).toBe('1.234,5');
    expect(displayFormat('m/d/yyyy')).toBe('dd.mm.yyyy');
    expect(formatValue(d2024, displayFormat('m/d/yyyy')!).text).toBe('15.03.2024');
    expect(TEXT(d2024, 'dddd, d. mmmm yyyy')).toBe('Freitag, 15. März 2024');
    expect(editDateText(2024, 3, 5)).toBe('05.03.2024');
    expect(activeLocale().firstDay).toBe(1);
  });
  it('formula bar uses ";" and "," while storing the canonical form', () => {
    setLocale('de-DE');
    const stored = 'SUM(A1,1.5,{1,2;3,4})+IF(B1="x,y.z",ROUND(C1,2),0.5)';
    const shown = localizeFormula(stored);
    expect(shown).toBe('SUM(A1;1,5;{1.2;3.4})+IF(B1="x,y.z";ROUND(C1;2);0,5)');
    expect(delocalizeFormula(shown)).toBe(stored);
    expect(localizeFormula("'Q1.2024'!A1+_xlfn.FORECAST.ETS(A1,B1:B9,C1:C9)")).toBe("'Q1.2024'!A1+_xlfn.FORECAST.ETS(A1;B1:B9;C1:C9)");
    expect(localizeFormula('Table1[[#This Row],[Price]]*1.1')).toBe('Table1[[#This Row],[Price]]*1,1');
    // Typing commas between references still works.
    expect(delocalizeFormula('SUM(A1,B1)')).toBe('SUM(A1,B1)');
  });
});

describe('other regions', () => {
  it('French, Swedish and Japanese dates', () => {
    setLocale('fr-FR');
    expect(parseNumber(delocalizeInput('15/03/2024'))).toBe(d2024);
    expect(parseNumber(delocalizeInput('1 234,5'))).toBe(1234.5);
    expect(TEXT(d2024, 'mmmm')).toBe('mars');
    setLocale('sv-SE');
    expect(formatValue(d2024, displayFormat('m/d/yyyy')!).text).toBe('2024-03-15');
    setLocale('ja-JP');
    expect(formatValue(d2024, displayFormat('m/d/yyyy')!).text).toBe('2024/03/15');
    expect(parseNumber('2024/3/15')).toBe(d2024);
  });
  it('US order stays month first', () => {
    expect(parseNumber('3/15/2024')).toBe(d2024);
    expect(parseNumber('15/3/2024')).toBeNull();
    setLocale('en-GB');
    expect(parseNumber('15/3/2024')).toBe(d2024);
    expect(parseNumber('3/15/2024')).toBeNull();
  });
  it('[$-xxxx] codes pick names and digits regardless of the region', () => {
    expect(formatValue(d2024, '[$-40C]dddd d mmmm').text).toBe('vendredi 15 mars');
    expect(formatValue(d2024, '[$-407]mmm').text).toMatch(/^Mär/);
    expect(formatValue(d2024, '[$-2000401]yyyy').text).toBe('٢٠٢٤');
    expect(formatValue(1234.5, '[$€-407]#,##0.00').text).toBe('€1,234.50');
    setLocale('de-DE');
    expect(formatValue(d2024, '[$-409]mmmm').text).toBe('March');
  });
  it('Arabic-Indic digits option', () => {
    setLocale('ar-SA', { nativeDigits: true });
    expect(localizeNumberText('1,234.5', '#,##0.0')).toBe('١٬٢٣٤٫٥');
    expect(delocalizeInput('١٢٫٥')).toBe('12.5');
    expect(parseNumber(delocalizeInput('١٢٫٥'))).toBe(12.5);
    setNativeDigits(false, false);
    expect(localizeNumberText('12.5')).toBe('12.5');
    setLocale('de-DE', { nativeDigits: true });
    expect(localizeNumberText('12.5')).toBe('12,5');
  });
  it('built-in currency formats follow the region', () => {
    expect(regionalFormat('"$"#,##0.00')).toBe('"$"#,##0.00');
    setLocale('de-DE');
    expect(regionalFormat('"$"#,##0.00')).toBe('#,##0.00\\ [$€-407]');
    expect(formatValue(1234.5, regionalFormat('"$"#,##0.00')).text).toBe('1,234.50 €');
    expect(regionalFormat('_("$"* #,##0.00_)')).toBe('_([$€-407]* #,##0.00_)');
    setLocale('ja-JP');
    expect(formatValue(1234, regionalFormat('"$"#,##0')).text).toBe('¥1,234');
  });
  it('Options separators override the region', () => {
    setLocale('en-US');
    setSeparators(',', '.');
    expect(localizeNumberText('1,234.5')).toBe('1.234,5');
    expect(delocalizeInput('1.234,5')).toBe('1234.5');
    expect(localizeFormula('ROUND(A1,2.5)')).toBe('ROUND(A1;2,5)');
    setSeparators('', '');
    expect(localizeNumberText('1,234.5')).toBe('1,234.5');
  });
  it('array constants use each region\'s separators', () => {
    const f = 'SUM({1.5,2;3,4.25})';
    const want: Record<string, string> = {
      'en-US': f, 'en-GB': f, 'ja-JP': f, 'de-CH': f,
      'de-DE': 'SUM({1,5.2;3.4,25})', 'fr-FR': 'SUM({1,5.2;3.4,25})',
      'es-ES': 'SUM({1,5\\2;3\\4,25})', 'it-IT': 'SUM({1,5\\2;3\\4,25})', 'pt-BR': 'SUM({1,5\\2;3\\4,25})',
      'nl-NL': 'SUM({1,5\\2;3\\4,25})', 'sv-SE': 'SUM({1,5\\2;3\\4,25})', 'ru-RU': 'SUM({1,5\\2;3\\4,25})',
    };
    for (const [id, shown] of Object.entries(want)) {
      setLocale(id);
      expect(localizeFormula(f), id).toBe(shown);
    }
  });
  it('every region round-trips formulas', () => {
    const formulas = [
      'SUM(A1:B2,1.5,{1,2;3,4})', 'IF(A1>0.25,"a,b;c",{1.5;2.5})', 'VLOOKUP(A1,{1,"x";2,"y"},2,FALSE)',
      "ROUND('Q1.5'!A1*1.05,2)+COUNTIF(A:A,\">1,5\")", 'TEXTJOIN(",",TRUE,{"a","b"})', 'LET(x,0.5,y,{0.1,0.2},x+SUM(y))',
      '_xlfn.FORECAST.ETS(A1,B1:B9,C1:C9,1,1)', 'Table1[[#This Row],[Price]]*1.1',
    ];
    for (const loc of LOCALES) {
      setLocale(loc.id);
      for (const f of formulas) expect(delocalizeFormula(localizeFormula(f)), `${loc.id}: ${f}`).toBe(f);
    }
    setLocale('en-US');
    setSeparators(',', '.');
    for (const f of formulas) expect(delocalizeFormula(localizeFormula(f)), `override: ${f}`).toBe(f);
    expect(localizeFormula('SUM({1.5,2})')).toBe('SUM({1,5\\2})');
  });
  it('remembers the region on this device', () => {
    const store = new Map<string, string>();
    const g = globalThis as unknown as { localStorage?: unknown };
    const old = g.localStorage;
    g.localStorage = { getItem: (k: string) => store.get(k) ?? null, setItem: (k: string, v: string) => { store.set(k, v); }, removeItem: (k: string) => { store.delete(k); } };
    try {
      setLocale('pl-PL', { persist: true });
      expect(JSON.parse(store.get(LOCALE_KEY)!)).toEqual({ id: 'pl-PL', nativeDigits: false });
      setLocale('en-US');
      initLocale();
      expect(activeLocale().id).toBe('pl-PL');
      store.set(LOCALE_KEY, 'tr-TR');
      initLocale();
      expect(activeLocale().id).toBe('tr-TR');
    } finally { g.localStorage = old; }
  });
});
