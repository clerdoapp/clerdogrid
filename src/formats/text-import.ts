// Text import: decoding, delimited and fixed-width splitting, and column conversion with a chosen
// number locale and date order. Used by the Text Import dialog and refreshable text queries.
import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../engine/types';
import { toSerial } from '../engine/date';
import { guessFormat } from './numfmt';

export const ENCODINGS: readonly [string, string][] = [
  ['auto', 'Detect'], ['utf-8', 'Unicode (UTF-8)'], ['utf-16le', 'Unicode (UTF-16 LE)'], ['utf-16be', 'Unicode (UTF-16 BE)'],
  ['windows-1252', 'Western (Windows-1252)'], ['iso-8859-1', 'Western (ISO-8859-1)'], ['iso-8859-2', 'Central European (ISO-8859-2)'],
  ['iso-8859-5', 'Cyrillic (ISO-8859-5)'], ['windows-1251', 'Cyrillic (Windows-1251)'], ['iso-8859-6', 'Arabic (ISO-8859-6)'], ['windows-1256', 'Arabic (Windows-1256)'],
  ['iso-8859-7', 'Greek (ISO-8859-7)'], ['iso-8859-8', 'Hebrew (ISO-8859-8)'], ['iso-8859-9', 'Turkish (ISO-8859-9)'], ['iso-8859-15', 'Western (ISO-8859-15)'],
  ['shift_jis', 'Japanese (Shift_JIS)'], ['euc-jp', 'Japanese (EUC-JP)'], ['gb18030', 'Chinese Simplified (GB18030)'], ['big5', 'Chinese Traditional (Big5)'],
  ['euc-kr', 'Korean (EUC-KR)'], ['koi8-r', 'Cyrillic (KOI8-R)'],
];

/** Decodes bytes; "auto" uses a byte-order mark, then UTF-8 if valid, else Windows-1252. */
export const decodeText = (bytes: Uint8Array, encoding = 'auto'): { text: string; encoding: string } => {
  let enc = encoding;
  let body = bytes;
  if (enc === 'auto') {
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) { enc = 'utf-8'; body = bytes.subarray(3); }
    else if (bytes[0] === 0xff && bytes[1] === 0xfe) { enc = 'utf-16le'; body = bytes.subarray(2); }
    else if (bytes[0] === 0xfe && bytes[1] === 0xff) { enc = 'utf-16be'; body = bytes.subarray(2); }
    else {
      // UTF-16 without a mark: many zero bytes in alternate positions
      let evenZeros = 0, oddZeros = 0;
      const n = Math.min(bytes.length, 4096);
      for (let i = 0; i < n; i++) if (bytes[i] === 0) { if (i % 2) oddZeros++; else evenZeros++; }
      if (oddZeros > n / 4 && evenZeros < n / 40) enc = 'utf-16le';
      else if (evenZeros > n / 4 && oddZeros < n / 40) enc = 'utf-16be';
      else {
        try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes), encoding: 'utf-8' }; } catch { enc = 'windows-1252'; }
      }
    }
  } else if (enc === 'utf-8' && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) body = bytes.subarray(3);
  else if (enc === 'utf-16le' && bytes[0] === 0xff && bytes[1] === 0xfe) body = bytes.subarray(2);
  else if (enc === 'utf-16be' && bytes[0] === 0xfe && bytes[1] === 0xff) body = bytes.subarray(2);
  let text: string;
  try { text = new TextDecoder(enc).decode(body); } catch { text = new TextDecoder('windows-1252').decode(body); enc = 'windows-1252'; }
  return { text: text.charCodeAt(0) === 0xfeff ? text.slice(1) : text, encoding: enc };
};

export const splitLines = (text: string): string[] => {
  const lines = text.split(/\r\n|\r|\n/);
  if (lines.length && lines.at(-1) === '') lines.pop();
  return lines;
};

export interface DelimitedOptions {
  delimiters: string[];
  /** Runs of delimiters count as one. */
  mergeDelimiters?: boolean;
  /** '"', "'" or '' for none. */
  qualifier: string;
}

/** Splits delimited text; qualified fields may contain delimiters and line breaks. */
export const splitDelimited = (text: string, o: DelimitedOptions, maxRows = Infinity): string[][] => {
  const rows: string[][] = [];
  const delims = new Set(o.delimiters.filter((d) => d.length === 1));
  const multi = o.delimiters.filter((d) => d.length > 1);
  const q = o.qualifier;
  let row: string[] = [];
  let field = '';
  let i = 0;
  const n = text.length;
  let quoted = false;
  let fieldStart = true;
  const isDelim = (at: number): number => {
    const ch = text[at]!;
    if (delims.has(ch)) return 1;
    for (const m of multi) if (text.startsWith(m, at)) return m.length;
    return 0;
  };
  const endField = () => { row.push(field); field = ''; fieldStart = true; };
  const endRow = () => { endField(); rows.push(row); row = []; };
  while (i < n && rows.length < maxRows) {
    const ch = text[i]!;
    if (quoted) {
      if (ch === q) {
        if (text[i + 1] === q) { field += q; i += 2; continue; }
        quoted = false; i++; continue;
      }
      field += ch; i++; continue;
    }
    if (q && ch === q && fieldStart) { quoted = true; fieldStart = false; i++; continue; }
    const d = isDelim(i);
    if (d) {
      endField();
      i += d;
      if (o.mergeDelimiters) while (i < n && isDelim(i)) i += isDelim(i);
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      endRow();
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    field += ch;
    fieldStart = false;
    i++;
  }
  if ((field !== '' || row.length) && rows.length < maxRows) endRow();
  return rows;
};

/** Picks the delimiter that splits the sample into the most consistent rows. */
export const guessDelimiter = (text: string): string => {
  const sample = splitLines(text.slice(0, 32 * 1024)).slice(0, 30);
  let best = ',', bestScore = -1;
  for (const d of [',', ';', '\t', '|', ' ']) {
    const counts = splitDelimited(sample.join('\n'), { delimiters: [d], qualifier: '"', mergeDelimiters: d === ' ' }).map((r) => r.length);
    if (!counts.length) continue;
    const freq = new Map<number, number>();
    for (const c of counts) freq.set(c, (freq.get(c) ?? 0) + 1);
    const [mode, hits] = [...freq].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]!;
    if (mode < 2) continue;
    const score = (hits / counts.length) * 1000 + mode - (d === ' ' ? 200 : 0);
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
};

/** True when the text looks like fixed-width columns rather than delimited fields. */
export const looksFixedWidth = (text: string): boolean => {
  const lines = splitLines(text.slice(0, 32 * 1024)).slice(0, 30).filter((l) => l.trim());
  if (lines.length < 2) return false;
  if (/[,;\t|]/.test(lines.join(''))) return false;
  return detectBreaks(lines).length > 0;
};

/** Column breaks: positions where every sampled line has a space and text starts after it. */
export const detectBreaks = (lines: string[]): number[] => {
  const sample = lines.filter((l) => l.trim()).slice(0, 200);
  if (!sample.length) return [];
  const width = Math.max(...sample.map((l) => l.length));
  const blank = new Array<boolean>(width + 1).fill(true);
  for (const l of sample) for (let i = 0; i < width; i++) if (l[i] !== undefined && l[i] !== ' ') blank[i] = false;
  const breaks: number[] = [];
  for (let i = 1; i < width; i++) if (blank[i - 1] && !blank[i]) breaks.push(i);
  return breaks;
};

export const splitFixed = (lines: string[], breaks: number[]): string[][] => {
  const b = [...new Set(breaks)].filter((x) => x > 0).sort((x, y) => x - y);
  return lines.map((l) => {
    const out: string[] = [];
    let start = 0;
    for (const x of b) { out.push(l.slice(start, x).trim()); start = x; }
    out.push(l.slice(start).trim());
    return out;
  });
};

export type ColumnFormat = 'general' | 'text' | 'MDY' | 'DMY' | 'YMD' | 'MYD' | 'DYM' | 'YDM' | 'skip';
export const COLUMN_FORMATS: readonly [ColumnFormat, string][] = [
  ['general', 'General'], ['text', 'Text'], ['MDY', 'Date (MDY)'], ['DMY', 'Date (DMY)'], ['YMD', 'Date (YMD)'], ['MYD', 'Date (MYD)'], ['DYM', 'Date (DYM)'], ['YDM', 'Date (YDM)'], ['skip', 'Skip column'],
];

export interface NumberLocale { decimal: string; thousands: string }
export const LOCALES: readonly [string, string, NumberLocale][] = [
  ['en', '1,234.56', { decimal: '.', thousands: ',' }], ['de', '1.234,56', { decimal: ',', thousands: '.' }],
  ['fr', '1 234,56', { decimal: ',', thousands: ' ' }], ['ch', "1'234.56", { decimal: '.', thousands: "'" }], ['plain', '1234.56', { decimal: '.', thousands: '' }],
];

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/** Number in the given locale (with %, currency signs and parentheses for negatives), or null. */
export const parseLocaleNumber = (raw: string, loc: NumberLocale): { n: number; fmt?: string } | null => {
  let s = raw.trim();
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  if (/-$/.test(s) && !/^-/.test(s)) { neg = true; s = s.slice(0, -1); }
  let pct = false;
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  let cur = '';
  const cm = /^([$€£¥])\s*|\s*([$€£¥])$/.exec(s);
  if (cm) { cur = cm[1] ?? cm[2]!; s = s.replace(cm[0], ''); }
  if (s.startsWith('-')) { neg = !neg; s = s.slice(1); } else if (s.startsWith('+')) s = s.slice(1);
  if (loc.thousands) {
    const t = loc.thousands === ' ' ? /[   ]/g : new RegExp(`\\${loc.thousands}`, 'g');
    const parts = s.split(loc.decimal);
    if (parts.length > 2) return null;
    const groups = parts[0]!.split(t);
    if (groups.length > 1 && (groups.slice(1).some((g) => g.length !== 3) || !groups[0])) return null;
    s = parts[0]!.replace(t, '') + (parts[1] !== undefined ? '.' + parts[1] : '');
  } else if (loc.decimal !== '.') s = s.replace(loc.decimal, '.');
  if (!/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return null;
  let n = Number(s);
  if (!Number.isFinite(n)) return null;
  if (neg) n = -n;
  if (pct) n /= 100;
  const dec = (s.split('.')[1] ?? '').replace(/[eE].*/, '').length;
  const fmt = pct ? (dec ? `0.${'0'.repeat(dec)}%` : '0%') : cur ? `"${cur}"#,##0${dec ? '.' + '0'.repeat(Math.min(dec, 2)) : ''}` : loc.thousands && raw.includes(loc.thousands) ? `#,##0${dec ? '.' + '0'.repeat(dec) : ''}` : undefined;
  return { n, fmt };
};

/** A date in the given field order ("1/2/2024", "2024-01-02", "2 Jan 2024", with optional time). */
export const parseDateOrder = (raw: string, order: string): number | null => {
  const s = raw.trim();
  const m = /^(\w+)[./\-\s]+(\w+)[./\-\s]+(\w+)(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([ap]\.?m\.?)?)?$/i.exec(s);
  if (!m) return null;
  const parts: Record<string, string> = { [order[0]!]: m[1]!, [order[1]!]: m[2]!, [order[2]!]: m[3]! };
  const monthOf = (t: string) => { const i = MONTHS.indexOf(t.slice(0, 3).toLowerCase()); return i >= 0 ? i + 1 : +t; };
  let y = +parts.Y!, mo = monthOf(parts.M!), d = +parts.D!;
  if (!Number.isInteger(y) || !Number.isInteger(mo) || !Number.isInteger(d)) return null;
  if (parts.Y!.length <= 2) y += y < 30 ? 2000 : 1900;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  let serial = toSerial(y, mo, d);
  if (m[4]) {
    let h = +m[4];
    if (m[7]) { const pm = /^p/i.test(m[7]); if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
    serial += (h * 3600 + +m[5]! * 60 + +(m[6] ?? 0)) / 86400;
  }
  return serial;
};

export interface ConvertOptions { formats?: ColumnFormat[]; locale: NumberLocale; dateOrder?: string; trim?: boolean }

/** Turns text fields into values: numbers in the chosen locale, dates in the chosen order, TRUE/FALSE, errors. */
export const convertRows = (rows: string[][], o: ConvertOptions): { values: Scalar[][]; formats: (string | undefined)[][]; skipped: Set<number> } => {
  const skipped = new Set<number>();
  (o.formats ?? []).forEach((f, i) => { if (f === 'skip') skipped.add(i); });
  const values: Scalar[][] = [];
  const formats: (string | undefined)[][] = [];
  for (const r of rows) {
    const vr: Scalar[] = [];
    const fr: (string | undefined)[] = [];
    r.forEach((raw0, i) => {
      if (skipped.has(i)) return;
      const raw = o.trim ? raw0.trim() : raw0;
      const f = o.formats?.[i] ?? 'general';
      if (raw === '') { vr.push(null); fr.push(undefined); return; }
      if (f === 'text') { vr.push(raw); fr.push('@'); return; }
      if (f !== 'general') {
        const d = parseDateOrder(raw, f);
        if (d !== null) { vr.push(d); fr.push(d % 1 ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd'); }
        else { vr.push(raw); fr.push(undefined); }
        return;
      }
      const up = raw.trim().toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') { vr.push(up === 'TRUE'); fr.push(undefined); return; }
      if (ERROR_CODES.includes(up as ErrorCode)) { vr.push(new CellError(up as ErrorCode)); fr.push(undefined); return; }
      const num = parseLocaleNumber(raw, o.locale);
      if (num) { vr.push(num.n); fr.push(num.fmt); return; }
      const d = o.dateOrder ? parseDateOrder(raw, o.dateOrder) : /^\d{4}-\d{2}-\d{2}/.test(raw.trim()) ? parseDateOrder(raw.trim().replace('T', ' '), 'YMD') : null;
      if (d !== null) { vr.push(d); fr.push(guessFormat(raw.trim()) ?? (d % 1 ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd')); return; }
      vr.push(raw);
      fr.push(undefined);
    });
    values.push(vr);
    formats.push(fr);
  }
  return { values, formats, skipped };
};

export interface TextImportOptions {
  encoding: string;
  kind: 'delimited' | 'fixed';
  startRow: number;
  delimited: DelimitedOptions;
  breaks: number[];
  formats: ColumnFormat[];
  locale: NumberLocale;
  dateOrder?: string;
  trim?: boolean;
}

/** Sensible starting options for a file. */
export const guessOptions = (bytes: Uint8Array, fileName = ''): TextImportOptions => {
  const { text, encoding } = decodeText(bytes);
  const fixed = /\.prn$/i.test(fileName) || looksFixedWidth(text);
  const lines = splitLines(text.slice(0, 64 * 1024));
  const d = guessDelimiter(text);
  // a comma decimal separator is likely when ";" separates fields and numbers contain ","
  const commaDecimal = d === ';' && /\d,\d/.test(text.slice(0, 8192));
  return {
    encoding, kind: fixed ? 'fixed' : 'delimited', startRow: 1,
    delimited: { delimiters: [d], qualifier: '"', mergeDelimiters: d === ' ' },
    breaks: fixed ? detectBreaks(lines) : [], formats: [],
    locale: commaDecimal ? { decimal: ',', thousands: '.' } : { decimal: '.', thousands: ',' },
  };
};

/** Splits (and optionally converts) a text file with the given options. */
export const parseText = (bytes: Uint8Array, o: TextImportOptions, maxRows = Infinity): { fields: string[][]; text: string } => {
  const { text } = decodeText(bytes, o.encoding);
  const skip = Math.max(0, o.startRow - 1);
  if (o.kind === 'fixed') {
    const lines = splitLines(text).slice(skip, skip + maxRows);
    return { fields: splitFixed(lines, o.breaks), text };
  }
  const all = splitDelimited(text, o.delimited, skip + maxRows);
  return { fields: all.slice(skip), text };
};

export const importText = (bytes: Uint8Array, o: TextImportOptions): { values: Scalar[][]; formats: (string | undefined)[][] } => {
  const { fields } = parseText(bytes, o);
  const { values, formats } = convertRows(fields, { formats: o.formats, locale: o.locale, dateOrder: o.dateOrder, trim: o.trim ?? o.kind === 'fixed' });
  return { values, formats };
};
