import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../engine/types';
import { parseNumber } from '../engine/coerce';
import { guessFormat } from './numfmt';

const CANDIDATES = [',', ';', '\t', '|'];

/** Picks the delimiter that splits the first lines into the most consistent, widest rows. */
const detectDelimiter = (text: string): string => {
  const sample = text.slice(0, 64 * 1024);
  let best = ',';
  let bestScore = -1;
  for (const d of CANDIDATES) {
    const all = splitRows(sample, d);
    // the last sampled line may be cut short
    const counts = all.slice(0, all.length > 1 && sample.length < text.length ? Math.min(20, all.length - 1) : 20).map((r) => r.fields.length);
    if (!counts.length) continue;
    const freq = new Map<number, number>();
    for (const c of counts) freq.set(c, (freq.get(c) ?? 0) + 1);
    const mode = [...freq].sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]![0];
    if (mode < 2) continue;
    const consistent = counts.filter((c) => c === mode).length / counts.length;
    const score = consistent * 1000 + mode;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
};

interface RawRow { fields: string[] }

/** RFC 4180 with the usual leniency: CR, LF or CRLF line ends, stray quotes inside unquoted fields kept. */
const splitRows = (text: string, d: string): RawRow[] => {
  const rows: RawRow[] = [];
  let fields: string[] = [];
  let i = 0;
  const n = text.length;
  while (i <= n) {
    let val = '';
    if (text[i] === '"') {
      i++;
      for (;;) {
        const j = text.indexOf('"', i);
        if (j < 0) { val += text.slice(i); i = n; break; }
        val += text.slice(i, j);
        if (text[j + 1] === '"') { val += '"'; i = j + 2; }
        else { i = j + 1; break; }
      }
      // anything between the closing quote and the delimiter is kept, as spreadsheets do
      while (i < n && text[i] !== d && text[i] !== '\n' && text[i] !== '\r') val += text[i++];
    } else {
      let j = i;
      while (j < n && text[j] !== d && text[j] !== '\n' && text[j] !== '\r') j++;
      val = text.slice(i, j);
      i = j;
    }
    fields.push(val);
    if (i < n && text[i] === d) { i++; continue; }
    rows.push({ fields });
    fields = [];
    if (i >= n) break;
    if (text[i] === '\r' && text[i + 1] === '\n') i += 2;
    else i++;
    if (i >= n) break;
  }
  return rows;
};

const convert = (raw: string): [Scalar, string | undefined] => {
  if (raw === '') return [null, undefined];
  const s = raw.trim();
  const up = s.toUpperCase();
  if (up === 'TRUE') return [true, undefined];
  if (up === 'FALSE') return [false, undefined];
  if (ERROR_CODES.includes(up as ErrorCode)) return [new CellError(up as ErrorCode), undefined];
  // spreadsheet converts quoted fields as well ("007" becomes 7); only the ="..." idiom forces text.
  if (/^=".*"$/.test(s)) return [s.slice(2, -1), undefined];
  const n = parseNumber(s);
  if (n === null) return [raw, undefined];
  return [n, guessFormat(s)];
};

export const readCsv = (text: string, opts: { delimiter?: string } = {}): { rows: Scalar[][]; formats: (string | undefined)[][] } => {
  let t = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  let delim = opts.delimiter;
  const sep = /^sep=(.)\r?\n/i.exec(t);
  if (sep) { delim ??= sep[1]; t = t.slice(sep[0].length); }
  delim ??= detectDelimiter(t);
  const raw = t === '' ? [] : splitRows(t, delim);
  const rows: Scalar[][] = [];
  const formats: (string | undefined)[][] = [];
  for (const r of raw) {
    const vals: Scalar[] = [];
    const fmts: (string | undefined)[] = [];
    r.fields.forEach((f) => {
      const [v, fm] = convert(f);
      vals.push(v);
      fmts.push(fm);
    });
    rows.push(vals);
    formats.push(fmts);
  }
  return { rows, formats };
};

export const writeCsv = (rows: string[][], delimiter = ','): string =>
  rows
    .map((r) => r.map((f) => (f.includes(delimiter) || /["\r\n]/.test(f) || /^\s|\s$/.test(f) ? `"${f.replace(/"/g, '""')}"` : f)).join(delimiter))
    .join('\r\n') + (rows.length ? '\r\n' : '');
