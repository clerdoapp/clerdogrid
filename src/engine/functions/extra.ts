// Functions that complete the catalog: trim-reference forms, EUROCONVERT, aliases, file imports and
// the ones that would run outside code (CALL, REGISTER.ID, RTD), which are blocked.
import { readCsv, splitCsv } from '../../formats/csv';
import { localeById, type LocaleInfo } from '../../formats/locales';
import { daysInMonth, toSerial } from '../date';
import { first, optBool, text, toText } from '../coerce';
import { getFunction } from '../registry';
import { cachedFetch, fetchText, offline, onlineAllowed, safeHttpsUrl } from '../online';
import { trimRange } from '../trimref';
import { CellError, Err, ERROR_CODES, isError, isRef, type ErrorCode, type FnContext, type FunctionDef, type Matrix, type Scalar, type Value } from '../types';

// ---------- _TRO_ALL / _TRO_LEADING / _TRO_TRAILING (file form of A1.:.B9, A1.:B9, A1:.B9) ----------

const tro = (name: string, mode: number): FunctionDef => ({
  name, min: 1, max: 1,
  fn: ([v], ctx) => {
    if (isRef(v!)) return trimRange(v, ctx, mode, mode);
    const t = getFunction('TRIMRANGE');
    return t ? t.fn([v!, mode, mode], ctx) : Err.value;
  },
});

// ---------- EUROCONVERT ----------
// Fixed conversion rates set when each currency joined the euro, with the calculation precision used
// for rounding (digits after the point).
const EURO: Record<string, [number, number]> = {
  EUR: [1, 2], BEF: [40.3399, 0], LUF: [40.3399, 0], DEM: [1.95583, 2], ESP: [166.386, 0], FRF: [6.55957, 2], IEP: [0.787564, 2],
  ITL: [1936.27, 0], NLG: [2.20371, 2], ATS: [13.7603, 2], PTE: [200.482, 1], FIM: [5.94573, 2], GRD: [340.75, 0], SIT: [239.64, 2],
  CYP: [0.585274, 2], MTL: [0.4293, 2], SKK: [30.126, 2], EEK: [15.6466, 2], LVL: [0.702804, 2], LTL: [3.4528, 2], HRK: [7.5345, 2],
};
const roundTo = (x: number, d: number): number => { const f = 10 ** d; return Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f; };

const euroconvert = (args: Value[], ctx: FnContext): Value => {
  const n = first(args[0]!, ctx);
  if (isError(n)) return n;
  const x = typeof n === 'number' ? n : n === null ? 0 : Number(toText(n));
  if (!Number.isFinite(x)) return Err.value;
  const src = text(args[1]!, ctx); if (isError(src)) return src;
  const dst = text(args[2]!, ctx); if (isError(dst)) return dst;
  const a = EURO[src.trim().toUpperCase()], b = EURO[dst.trim().toUpperCase()];
  if (!a || !b) return Err.value;
  const full = optBool(args[3], ctx, false); if (isError(full)) return full;
  let tri: number | null = null;
  if (args[4] !== undefined && args[4] !== null) {
    const t = first(args[4], ctx);
    if (isError(t)) return t;
    const tn = typeof t === 'number' ? Math.trunc(t) : NaN;
    if (!Number.isFinite(tn) || tn < 3) return Err.value;
    tri = tn;
  }
  let eur = x / a[0];
  // Between two national currencies the euro amount may be rounded first (triangulation).
  if (tri !== null && a !== EURO.EUR && b !== EURO.EUR) eur = roundTo(eur, tri);
  const out = eur * b[0];
  return full ? out : roundTo(out, b[1]);
};

// ---------- IMPORTCSV / IMPORTTEXT ----------

const sliceRows = (rows: Scalar[][], skip: number, take: number | null): Scalar[][] => {
  let r = skip >= 0 ? rows.slice(skip) : rows.slice(0, Math.max(0, rows.length + skip));
  if (take !== null) r = take >= 0 ? r.slice(0, take) : r.slice(Math.max(0, r.length + take));
  return r;
};

/**
 * One field as the given region writes it: "1.234,5", "12,5 %", "3,50 €", "15.03.2024" (in its day/month
 * order), TRUE/FALSE, error codes, ="..." for forced text; anything else stays text.
 */
export const parseLocalField = (raw: string, loc: LocaleInfo, date1904: boolean): Scalar => {
  if (raw === '') return null;
  const s = raw.trim();
  const up = s.toUpperCase();
  if (up === 'TRUE') return true;
  if (up === 'FALSE') return false;
  if ((ERROR_CODES as readonly string[]).includes(up)) return new CellError(up as ErrorCode);
  if (/^=".*"$/.test(s)) return s.slice(2, -1);
  const e = (c: string): string => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const th = loc.thousands === '\u00a0' || loc.thousands === '\u202f' ? '[\\u00a0\\u202f ]' : e(loc.thousands);
  const cur = e(loc.currency);
  const num = new RegExp(`^([-+]?)\\(?(?:${cur}\\s?)?([-+]?)(\\d{1,3}(?:${th}\\d{3})+|\\d+)?(?:${e(loc.decimal)}(\\d+))?(?:\\s?${cur})?\\)?\\s?(%?)$`);
  const m = num.exec(s);
  if (m && (m[3] || m[4])) {
    let v = Number(`${(m[3] ?? '0').replace(new RegExp(th, 'g'), '')}.${m[4] ?? '0'}`);
    if (m[1] === '-' || m[2] === '-' || (s.startsWith('(') && s.endsWith(')'))) v = -v;
    return m[5] ? v / 100 : v;
  }
  const ds = e(loc.dateSep);
  const d = new RegExp(`^(\\d{1,4})${ds}(\\d{1,2})${ds}(\\d{1,4})$`).exec(s) ?? /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (d) {
    const [a, b, c] = [+d[1]!, +d[2]!, +d[3]!];
    let y: number, mo: number, dd: number;
    if (d[1]!.length === 4 || loc.dateOrder === 'YMD') { y = a; mo = b; dd = c; }
    else if (loc.dateOrder === 'DMY') { dd = a; mo = b; y = c; }
    else { mo = a; dd = b; y = c; }
    if (y < 100) y += y < 30 ? 2000 : 1900;
    if (mo >= 1 && mo <= 12 && dd >= 1 && dd <= daysInMonth(y, mo)) return toSerial(y, mo, dd, date1904);
  }
  return raw;
};

const importFile = (args: Value[], ctx: FnContext, csv: boolean): Value => {
  const path = text(args[0]!, ctx); if (isError(path)) return path;
  const numArg = (v: Value | undefined): number | null | CellError => {
    if (v === undefined || v === null) return null;
    const s = first(v, ctx);
    if (s === null) return null;
    if (isError(s)) return s;
    const n = typeof s === 'number' ? s : Number(toText(s));
    return Number.isFinite(n) ? Math.trunc(n) : Err.value;
  };
  let delim: string | undefined = csv ? ',' : undefined;
  let k = 1;
  if (!csv) {
    const d = args[1] === undefined || args[1] === null ? '' : text(args[1], ctx);
    if (isError(d)) return d;
    delim = d === '' ? undefined : d === '\\t' ? '\t' : d;
    if (delim && delim.length !== 1) return Err.value;
    k = 2;
  }
  const skip = numArg(args[k]); if (isError(skip)) return skip;
  const take = numArg(args[k + 1]); if (isError(take)) return take;
  if (!csv && args[4] !== undefined && args[4] !== null) {
    const enc = text(args[4], ctx); if (isError(enc)) return enc;
    if (enc && !/^(utf-?8|65001)$/i.test(enc.trim())) return Err.value;
  }
  const locArg = args[csv ? 3 : 5];
  let loc: ReturnType<typeof localeById> | undefined;
  if (locArg !== undefined && locArg !== null) {
    const l = text(locArg, ctx); if (isError(l)) return l;
    if (l.trim()) { loc = localeById(l.trim()); if (!loc) return Err.value; }
  }
  // A CSV written for a region that uses "," as its decimal mark separates fields with ";".
  if (csv && loc) delim = loc.listSep;
  const url = safeHttpsUrl(path.trim());
  if (!url) return /^(https?|file):|^[A-Za-z]:\\|^[\\/]/.test(path.trim()) ? new CellError('#BLOCKED!', 'Formulas can only import files from https addresses.') : Err.value;
  if (!onlineAllowed()) return offline();
  const all = cachedFetch(`imp:${delim ?? ''}:${loc?.id ?? ''}:${url.href}`, 5 * 60_000, async () => {
    const body = await fetchText(url, { maxBytes: 20_000_000, timeoutMs: 30_000 });
    // With a locale the raw field text is read the region's way; otherwise the usual CSV reading applies.
    const rows: Scalar[][] = loc
      ? splitCsv(body, delim ? { delimiter: delim } : {}).map((r) => r.map((f) => parseLocalField(f, loc!, ctx.date1904)))
      : readCsv(body, delim ? { delimiter: delim } : {}).rows;
    if (!rows.length) return Err.calc;
    const w = Math.max(...rows.map((x) => x.length));
    return rows.map((x) => [...x, ...new Array<Scalar>(w - x.length).fill(null)]) as Matrix;
  });
  if (!Array.isArray(all)) return all;
  const r = sliceRows(all, skip ?? 0, take);
  return r.length ? r : Err.calc;
};

const alias = (name: string, target: string): FunctionDef => ({
  name, min: 0, max: 255,
  fn: (args, ctx) => { const d = getFunction(target); return d ? d.fn(args, ctx) : Err.name; },
});

const blocked = (name: string, min: number, max: number): FunctionDef => ({
  name, min, max, fn: () => new CellError('#BLOCKED!', 'This function runs outside code, which Grid never allows.'),
});

export const defs: FunctionDef[] = [
  tro('_TRO_ALL', 3), tro('_TRO_LEADING', 1), tro('_TRO_TRAILING', 2),
  { name: 'EUROCONVERT', min: 3, max: 5, scalar: [0, 1, 2, 3, 4], fn: euroconvert },
  { name: 'IMPORTCSV', min: 1, max: 4, volatile: true, fn: (a, c) => importFile(a, c, true) },
  { name: 'IMPORTTEXT', min: 1, max: 6, volatile: true, fn: (a, c) => importFile(a, c, false) },
  { ...alias('ECMA.CEILING', 'CEILING'), min: 2, max: 2, scalar: true },
  { ...alias('JIS', 'DBCS'), min: 1, max: 1, scalar: true },
  blocked('CALL', 1, 255),
  blocked('REGISTER.ID', 2, 3),
  { name: 'RTD', min: 3, max: 255, fn: () => new CellError('#N/A', 'No real-time data server is available.') },
  { name: 'COPILOT', min: 1, max: 255, fn: () => new CellError('#CONNECT!', 'ClerdoGrid does not send your data to an AI service.') },
];
