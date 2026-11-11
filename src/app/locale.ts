// Region settings: how numbers, dates and formulas are shown and typed. Workbooks always store numbers,
// dates and formulas in the file format's own form ("1234.5", serial dates, English functions with ","
// between arguments); this module converts to and from what the person sees for the active region.
//
// The region comes from localStorage "clerdogrid.locale" ({ id, nativeDigits }), else navigator.language.
// File > Options can still override the decimal and thousands separators through setSeparators().

import { activeLocale, fromNativeDigits, localeById, LOCALES, setActiveLocaleInfo, toNativeDigits, type LocaleInfo } from '../formats/locales';

export { LOCALES, activeLocale, type LocaleInfo };

export const LOCALE_KEY = 'clerdogrid.locale';

/** Effective separators for display and entry (the region's, unless overridden in Options). */
export const separators = { decimal: '.', thousands: ',' };
let override: { decimal: string; thousands: string } | null = null;
let nativeDigits = false;
const listeners = new Set<() => void>();

const read = (): { id?: string; nativeDigits?: boolean } => {
  try {
    const raw = localStorage.getItem(LOCALE_KEY);
    if (!raw) return {};
    if (raw.startsWith('{')) return JSON.parse(raw) as { id?: string; nativeDigits?: boolean };
    return { id: raw };
  } catch { return {}; }
};
const write = (): void => {
  try { localStorage.setItem(LOCALE_KEY, JSON.stringify({ id: activeLocale().id, nativeDigits })); } catch { /* private mode */ }
};

const syncSeparators = (): void => {
  const loc = activeLocale();
  separators.decimal = override?.decimal || loc.decimal;
  separators.thousands = override?.thousands || loc.thousands;
  if (separators.decimal === separators.thousands) separators.thousands = separators.decimal === ',' ? '.' : ',';
};

/** Options override for the separators; empty strings mean "use the region's". */
export const setSeparators = (decimal: string, thousands: string): void => {
  override = decimal || thousands ? { decimal, thousands } : null;
  syncSeparators();
  for (const l of listeners) l();
};

export const onLocaleChange = (l: () => void): (() => void) => { listeners.add(l); return () => listeners.delete(l); };

/** Makes a region active. `persist` saves it on this device. */
export const setLocale = (id: string, opts: { persist?: boolean; nativeDigits?: boolean } = {}): LocaleInfo => {
  const loc = localeById(id) ?? LOCALES[0]!;
  setActiveLocaleInfo(loc);
  if (opts.nativeDigits !== undefined) nativeDigits = opts.nativeDigits && !!loc.nativeDigits;
  else if (!loc.nativeDigits) nativeDigits = false;
  syncSeparators();
  if (opts.persist) write();
  for (const l of listeners) l();
  return loc;
};

export const useNativeDigits = (): boolean => nativeDigits;
export const setNativeDigits = (on: boolean, persist = true): void => {
  nativeDigits = on && !!activeLocale().nativeDigits;
  if (persist) write();
  for (const l of listeners) l();
};

/** Picks the saved region, or the browser's language, at start-up. */
export const initLocale = (): void => {
  const saved = read();
  const nav = typeof navigator !== 'undefined' ? navigator.language : undefined;
  setLocale(saved.id ?? nav ?? 'en-US', { nativeDigits: !!saved.nativeDigits });
};

// ---------- numbers and dates ----------

const DATE_CODE = /[dmyhs]/i;
const stripLit = (fmt: string): string => fmt.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');

/** Number format code to display with: the built-in system short date follows the region. */
export const displayFormat = (code: string | undefined): string | undefined => {
  if (!code) return code;
  const loc = activeLocale();
  if (loc.id === 'en-US') return code;
  if (code === 'm/d/yyyy' || code === 'mm-dd-yy') return loc.shortDate;
  if (code === 'm/d/yyyy h:mm') return `${loc.shortDate} ${loc.time.includes('AM/PM') ? 'h:mm AM/PM' : 'h:mm'}`;
  return code;
};

/** Swaps separators inside a formatted number and applies native digits when chosen. */
export const localizeNumberText = (text: string, fmt?: string): string => {
  if (!text) return text;
  let out = text;
  const isDate = !!fmt && fmt !== 'General' && DATE_CODE.test(stripLit(fmt));
  if (!isDate && (separators.decimal !== '.' || separators.thousands !== ',')) {
    out = out.replace(/(\d)([.,])(?=\d)|([.,])(?=\d)/g, (_m, d: string | undefined, sep: string | undefined, lone: string | undefined) => {
      const ch = sep ?? lone!;
      return (d ?? '') + (ch === '.' ? separators.decimal : separators.thousands);
    });
  }
  if (nativeDigits && activeLocale().nativeDigits) out = toNativeDigits(out, activeLocale().nativeDigits!);
  return out;
};

const esc = (c: string): string => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Turns a typed entry written the region's way into the stored form: native digits become ASCII and
 * numbers lose their regional separators ("1.234,5" -> "1234.5" in German). Dates are left for the date
 * reader, which knows the region's day/month order; other text is returned unchanged.
 */
export const delocalizeInput = (text: string): string => {
  const t0 = fromNativeDigits(text);
  const dec = separators.decimal, th = separators.thousands;
  const t = t0.trim();
  if (dec === '.' && th === ',') return t0;
  const loc = activeLocale();
  // Dates like 15.03.2024 or 1.5 (day and month) are left for the date reader.
  const ds = esc(loc.dateSep);
  if (new RegExp(`^\\d{1,4}${ds}\\d{1,2}(${ds}\\d{1,4})?${loc.dateSep === '.' ? '\\.?' : ''}$`).test(t)) {
    // "15.03" would read as a number, so dotted dates are handed on with slashes.
    return loc.dateSep === '.' && dec !== '.' ? t.replace(/\.$/, '').split('.').join('/') : t0;
  }
  const thx = th === ' ' || th === ' ' || th === ' ' ? '[\\u00a0\\u202f ]' : esc(th);
  const num = new RegExp(`^([-+(]?[^\\d\\s]{0,4}\\s?)(\\d{1,3}(?:${thx}\\d{3})+|\\d+)?(${esc(dec)}\\d*)?(\\s?[^\\d]{0,4})$`);
  const m = num.exec(t);
  if (!m || (!m[2] && !m[3]) || (m[3] === dec && !m[2])) return t0;
  const whole = (m[2] ?? '').replace(new RegExp(thx, 'g'), '');
  const frac = m[3] ? '.' + m[3].slice(dec.length) : '';
  return `${m[1]}${whole}${frac}${m[4]}`;
};

/** A date as typed in the formula bar for the active region (day/month order and separator). */
export const editDateText = (year: number, month: number, day: number): string => {
  const loc = activeLocale();
  const s = loc.dateSep;
  const pad = (n: number): string => (/(dd|mm)/.test(loc.shortDate) ? String(n).padStart(2, '0') : String(n));
  if (loc.dateOrder === 'DMY') return `${pad(day)}${s}${pad(month)}${s}${year}`;
  if (loc.dateOrder === 'YMD') return `${year}${s}${pad(month)}${s}${pad(day)}`;
  return `${month}/${day}/${year}`;
};

// ---------- formulas ----------

/**
 * Walks formula text outside string literals, quoted sheet names and [structured] parts, calling `f` for
 * each character with whether it sits inside an array constant.
 */
const mapFormula = (src: string, f: (ch: string, i: number, inArray: boolean) => string): string => {
  let out = '';
  let brace = 0, bracket = 0;
  for (let i = 0; i < src.length; i++) {
    const c = src[i]!;
    if (c === '"' || (c === "'" && bracket === 0)) {
      let j = i + 1;
      while (j < src.length) { if (src[j] === c) { if (src[j + 1] === c) { j += 2; continue; } break; } j++; }
      out += src.slice(i, j + 1);
      i = j;
      continue;
    }
    if (c === '[') bracket++;
    else if (c === ']') bracket = Math.max(0, bracket - 1);
    if (bracket > 0 || c === ']') { out += c; continue; }
    if (c === '{') brace++;
    else if (c === '}') brace = Math.max(0, brace - 1);
    out += f(c, i, brace > 0);
  }
  return out;
};

const isDigit = (c: string | undefined): boolean => c !== undefined && c >= '0' && c <= '9';
const isIdent = (c: string | undefined): boolean => c !== undefined && /[A-Za-z_.$\d]/.test(c);

/**
 * Separators the formula bar uses for the region: decimal mark, argument separator, and the column and
 * row separators inside array constants. An Options decimal-mark override wins and the others step aside.
 */
export const formulaSeparators = (): { dec: string; list: string; col: string; row: string } => {
  const loc = activeLocale();
  const dec = separators.decimal;
  const list = dec === ',' ? ';' : loc.listSep;
  let col = dec === loc.decimal && list === loc.listSep ? loc.arrayCol : list === ',' ? ',' : '\\';
  if (col === dec || (col === list && list !== ',')) col = '\\';
  return { dec, list, col, row: loc.arrayRow };
};

const canonicalSeps = (x: { dec: string; list: string; col: string; row: string }): boolean => x.dec === '.' && x.list === ',' && x.col === ',' && x.row === ';';

/** Stored formula -> what the formula bar shows for the region ("," decimals, ";" between arguments...). */
export const localizeFormula = (formula: string): string => {
  const sep = formulaSeparators();
  if (canonicalSeps(sep)) return formula;
  return mapFormula(formula, (c, i, inArray) => {
    if (inArray && c === ',') return sep.col;
    if (inArray && c === ';') return sep.row;
    if (c === ',') return sep.list;
    if (c === '.' && sep.dec !== '.' && isDigit(formula[i + 1]) && (isDigit(formula[i - 1]) || !isIdent(formula[i - 1]))) {
      // Only inside number literals: walk back over digits and make sure the run starts a number.
      let j = i - 1;
      while (j >= 0 && isDigit(formula[j])) j--;
      if (j >= 0 && /[A-Za-z_$]/.test(formula[j]!)) return c;
      return sep.dec;
    }
    return c;
  });
};

/** Formula typed for the region -> stored form. */
export const delocalizeFormula = (formula: string): string => {
  const sep = formulaSeparators();
  const f = fromNativeDigits(formula);
  if (canonicalSeps(sep)) return f;
  return mapFormula(f, (c, i, inArray) => {
    if (inArray && c === sep.col) return ',';
    if (inArray && c === sep.row) return ';';
    if (c === sep.dec && sep.dec !== '.' && isDigit(f[i + 1]) && (isDigit(f[i - 1]) || !isIdent(f[i - 1]))) {
      let j = i - 1;
      while (j >= 0 && isDigit(f[j])) j--;
      if (!(j >= 0 && /[A-Za-z_$]/.test(f[j]!))) return '.';
    }
    if (c === sep.list && !inArray) return ',';
    return c;
  });
};

/** Currency format code for the region, used when someone types an amount with the region's symbol. */
export const currencyFormat = (decimals: boolean): string => {
  const loc = activeLocale();
  const num = decimals ? '#,##0.00' : '#,##0';
  const tag = `[$${loc.currency}-${loc.lcid.toString(16).toUpperCase()}]`;
  return loc.currencyAfter ? `${num}${loc.currencySpace ? '\\ ' : ''}${tag}` : `${tag}${loc.currencySpace ? '\\ ' : ''}${num}`;
};

/** Built-in currency formats use the region's currency symbol (and position) instead of "$". */
export const regionalFormat = (code: string): string => {
  const loc = activeLocale();
  if ((loc.currency === '$' && !loc.currencyAfter) || !code.includes('"$"')) return code;
  if (code === '"$"#,##0.00' || code === '"$"#,##0') return currencyFormat(code.endsWith('.00'));
  return code.split('"$"').join(`[$${loc.currency}-${loc.lcid.toString(16).toUpperCase()}]`);
};

/** The argument separator people type in the formula bar for the active region. */
export const formulaListSep = (): string => formulaSeparators().list;
