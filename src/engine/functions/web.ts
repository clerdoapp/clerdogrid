import { first, text, toMatrix, toText } from '../coerce';
import { fromSerial, toSerial } from '../date';
import { cachedFetch, fetchText, offline, onlineAllowed, registerCellImage, safeHttpsUrl } from '../online';
import { CellError, Err, isError, type FnContext, type FunctionDef, type Matrix, type Scalar, type Value } from '../types';
import { xpath, XPathError } from '../xpath';

const MAX_TEXT = 32767;
const asNumber = (s: string): Scalar => (/^\s*-?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?\s*$/.test(s) ? Number(s) : s);

// ---------- WEBSERVICE ----------

const webservice = (args: Value[], ctx: FnContext): Value => {
  const u = text(args[0]!, ctx); if (isError(u)) return u;
  const url = safeHttpsUrl(u.trim());
  if (!url) return Err.value;
  if (!onlineAllowed()) return offline();
  return cachedFetch(`ws:${url.href}`, 60_000, async () => {
    const t = await fetchText(url, { maxBytes: 2_000_000, timeoutMs: 15_000 });
    return t.length > MAX_TEXT ? Err.value : t;
  });
};

// ---------- FILTERXML (offline: works on text already in the workbook) ----------

const filterxml = (args: Value[], ctx: FnContext): Value => {
  const xml = text(args[0]!, ctx); if (isError(xml)) return xml;
  const xp = text(args[1]!, ctx); if (isError(xp)) return xp;
  try {
    const r = xpath(xml, xp);
    if ('value' in r) return typeof r.value === 'string' ? asNumber(r.value) : typeof r.value === 'number' && !Number.isFinite(r.value) ? Err.num : r.value;
    if (!r.nodes.length) return Err.value;
    if (r.nodes.length === 1) return asNumber(r.nodes[0]!);
    return r.nodes.map((s) => [asNumber(s)]);
  } catch (e) {
    if (e instanceof XPathError) return Err.value;
    throw e;
  }
};

// ---------- IMAGE ----------

const image = (args: Value[], ctx: FnContext): Value => {
  const alt = args[1] === undefined || args[1] === null ? '' : text(args[1], ctx);
  if (isError(alt)) return alt;
  const opt = (v: Value | undefined): number | null | CellError => {
    if (v === undefined || v === null) return null;
    const s = first(v, ctx);
    if (s === null) return null;
    const n = typeof s === 'number' ? s : Number(toText(s));
    return isError(s) ? s : Number.isFinite(n) ? n : Err.value;
  };
  const sizing = opt(args[2]); if (isError(sizing)) return sizing;
  const height = opt(args[3]); if (isError(height)) return height;
  const width = opt(args[4]); if (isError(width)) return width;
  const sz = sizing === null ? 0 : Math.trunc(sizing);
  if (sz < 0 || sz > 3) return Err.value;
  if (sz === 3 ? !(height !== null && height > 0 && width !== null && width > 0) : height !== null || width !== null) return Err.value;
  const src = toMatrix(args[0]!, ctx); if (isError(src)) return src;
  const out: Matrix = src.map((row, i) => row.map((cell, j): Scalar => {
    if (isError(cell)) return cell;
    const url = typeof cell === 'string' ? safeHttpsUrl(cell.trim()) : null;
    if (!url) return Err.value;
    if (!onlineAllowed()) return offline();
    registerCellImage(ctx.sheet, ctx.row + i, ctx.col + j, {
      url: url.href, alt, sizing: sz, height: height ?? undefined, width: width ?? undefined, anchor: { sheet: ctx.sheet, row: ctx.row, col: ctx.col },
    });
    return alt;
  }));
  return out.length === 1 && out[0]!.length === 1 ? out[0]![0]! : out;
};

// ---------- STOCKHISTORY ----------
// Source: Stooq's free CSV download (https://stooq.com/q/d/l/), which needs no key. It is the only
// key-less daily price history we know of; when it can't be reached from this app (for example because
// the browser blocks the cross-site request), the function returns #CONNECT!.

const MIC_SUFFIX: Record<string, string> = {
  XNAS: 'us', XNYS: 'us', ARCX: 'us', BATS: 'us', XASE: 'us', XLON: 'uk', XETR: 'de', XFRA: 'de', XTKS: 'jp', XHKG: 'hk', XWAR: 'pl', XBUD: 'hu',
};
const stooqSymbol = (s: string): string | null => {
  const t = s.trim();
  if (!t || t.length > 32) return null;
  const m = /^([A-Z]{4}):([A-Za-z0-9.\-^]+)$/.exec(t);
  if (m) { const suf = MIC_SUFFIX[m[1]!]; return suf ? `${m[2]!.toLowerCase()}.${suf}` : null; }
  if (/^[A-Za-z0-9\-^]+\.[A-Za-z]{2}$/.test(t)) return t.toLowerCase();
  if (/^\^?[A-Za-z0-9\-]+$/.test(t)) return t.startsWith('^') ? t.toLowerCase() : `${t.toLowerCase()}.us`;
  return null;
};
const HEADERS = ['Date', 'Close', 'Open', 'High', 'Low', 'Volume'];

export const parseStooqCsv = (csv: string, date1904: boolean): { date: number; open: number; high: number; low: number; close: number; volume: number }[] | null => {
  const lines = csv.trim().split(/\r?\n/);
  if (!/^date,open,high,low,close/i.test(lines[0] ?? '')) return null;
  const out = [];
  for (const ln of lines.slice(1)) {
    const [d, o, h, l, c, v] = ln.split(',');
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d ?? '');
    if (!m) continue;
    out.push({ date: toSerial(+m[1]!, +m[2]!, +m[3]!, date1904), open: Number(o), high: Number(h), low: Number(l), close: Number(c), volume: Number(v ?? 0) || 0 });
  }
  return out;
};

const ymd = (serial: number, date1904: boolean): string => { const p = fromSerial(serial, date1904); return `${p.year}${String(p.month).padStart(2, '0')}${String(p.day).padStart(2, '0')}`; };

const stockhistory = (args: Value[], ctx: FnContext): Value => {
  const stock = text(args[0]!, ctx); if (isError(stock)) return stock;
  const num = (v: Value | undefined, d: number): number | CellError => {
    if (v === undefined || v === null) return d;
    const s = first(v, ctx);
    if (s === null) return d;
    if (isError(s)) return s;
    const n = typeof s === 'number' ? s : Number(toText(s));
    return Number.isFinite(n) ? n : Err.value;
  };
  const start = num(args[1], NaN); if (isError(start)) return start;
  if (!Number.isFinite(start)) return Err.value;
  const end = num(args[2], start); if (isError(end)) return end;
  const interval = num(args[3], 0); if (isError(interval)) return interval;
  const headers = num(args[4], 1); if (isError(headers)) return headers;
  const props: number[] = [];
  for (let i = 5; i < args.length; i++) { const p = num(args[i], -1); if (isError(p)) return p; if (p < 0 || p > 5 || !Number.isInteger(p)) return Err.value; props.push(p); }
  if (!props.length) props.push(0, 1);
  if (![0, 1, 2].includes(interval) || ![0, 1, 2].includes(headers) || end < start) return Err.value;
  const sym = stooqSymbol(stock);
  if (!sym) return Err.value;
  if (!onlineAllowed()) return offline();
  const s0 = Math.floor(start), s1 = Math.floor(end);
  const url = safeHttpsUrl(`https://stooq.com/q/d/l/?s=${encodeURIComponent(sym)}&d1=${ymd(s0, ctx.date1904)}&d2=${ymd(s1, ctx.date1904)}&i=${'dwm'[interval]}`)!;
  return cachedFetch(`stock:${url.href}`, 15 * 60_000, async () => {
    const csv = await fetchText(url, { maxBytes: 5_000_000, timeoutMs: 20_000 });
    const rows = parseStooqCsv(csv, ctx.date1904);
    if (!rows) return Err.na;
    const kept = rows.filter((r) => r.date >= s0 && r.date <= s1).sort((a, b) => a.date - b.date);
    if (!kept.length) return Err.na;
    const pick = (r: (typeof kept)[number], p: number): Scalar => [r.date, r.close, r.open, r.high, r.low, r.volume][p]!;
    const out: Matrix = [];
    if (headers === 2) out.push(props.map((_, i) => (i === 0 ? stock.toUpperCase() : '')));
    if (headers >= 1) out.push(props.map((p) => HEADERS[p]!));
    for (const r of kept) out.push(props.map((p) => pick(r, p)));
    return out;
  }, () => offline());
};

// ---------- TRANSLATE / DETECTLANGUAGE (the browser's on-device translator, when it has one) ----------

interface TranslatorLike { translate(text: string): Promise<string> }
interface TranslatorStatic { availability?(o: { sourceLanguage: string; targetLanguage: string }): Promise<string>; create(o: { sourceLanguage: string; targetLanguage: string }): Promise<TranslatorLike> }
interface DetectorLike { detect(text: string): Promise<{ detectedLanguage: string; confidence: number }[]> }
interface DetectorStatic { create(): Promise<DetectorLike> }
const api = (): { T?: TranslatorStatic; D?: DetectorStatic } => {
  const g = globalThis as { Translator?: TranslatorStatic; LanguageDetector?: DetectorStatic };
  return { T: g.Translator, D: g.LanguageDetector };
};
const LANG = /^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/;
const userLang = (): string => ((globalThis as { navigator?: { language?: string } }).navigator?.language ?? 'en').split('-')[0]!.toLowerCase();

let detector: Promise<DetectorLike> | null = null;
const detect = async (s: string): Promise<string | null> => {
  const { D } = api();
  if (!D) return null;
  detector ??= D.create();
  const r = await (await detector).detect(s);
  const top = r.find((x) => x.detectedLanguage !== 'und');
  return top ? top.detectedLanguage : null;
};
const translators = new Map<string, Promise<TranslatorLike>>();

const translate = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const src = args[1] === undefined || args[1] === null ? '' : text(args[1], ctx); if (isError(src)) return src;
  const dst = args[2] === undefined || args[2] === null ? userLang() : text(args[2], ctx); if (isError(dst)) return dst;
  const from = src.trim().toLowerCase(), to = dst.trim().toLowerCase();
  if ((from && from !== 'auto' && !LANG.test(from)) || !LANG.test(to)) return Err.value;
  if (s === '') return '';
  const { T } = api();
  if (!T) return new CellError('#CONNECT!', 'This browser has no built-in translator.');
  return cachedFetch(`tr:${from}>${to}:${s}`, 24 * 3600_000, async () => {
    const source = from && from !== 'auto' ? from : await detect(s);
    if (!source) return Err.value;
    if (source.split('-')[0] === to.split('-')[0]) return s;
    if (T.availability && (await T.availability({ sourceLanguage: source, targetLanguage: to })) === 'unavailable') return Err.value;
    const k = `${source}>${to}`;
    let tr = translators.get(k);
    if (!tr) { tr = T.create({ sourceLanguage: source, targetLanguage: to }); translators.set(k, tr); }
    const out = await (await tr).translate(s);
    return out.length > MAX_TEXT ? Err.value : out;
  });
};

const detectLanguage = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  if (!s.trim()) return Err.value;
  if (!api().D) return new CellError('#CONNECT!', 'This browser has no built-in language detector.');
  return cachedFetch(`dl:${s}`, 24 * 3600_000, async () => (await detect(s)) ?? Err.value);
};

export const defs: FunctionDef[] = [
  {
    name: 'ENCODEURL', min: 1, max: 1, scalar: true,
    fn: ([v]) => { const t = toText(v as Scalar); return isError(t) ? t : encodeURIComponent(t).replace(/[!'()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`); },
  },
  { name: 'WEBSERVICE', min: 1, max: 1, scalar: true, volatile: true, fn: webservice },
  { name: 'FILTERXML', min: 2, max: 2, scalar: true, fn: filterxml },
  { name: 'IMAGE', min: 1, max: 5, scalar: [1, 2, 3, 4], volatile: true, fn: image },
  { name: 'STOCKHISTORY', min: 2, max: 11, volatile: true, fn: stockhistory },
  { name: 'TRANSLATE', min: 1, max: 3, scalar: true, volatile: true, fn: translate },
  { name: 'DETECTLANGUAGE', min: 1, max: 1, scalar: true, volatile: true, fn: detectLanguage },
  { name: 'FIELDVALUE', min: 2, max: 2, fn: () => Err.field },
];
