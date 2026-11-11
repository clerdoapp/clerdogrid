import { CellError, Err, isError, isLambda, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { compileRegex, expandReplacement, RegexError, toCodePoints, type Regex } from '../regex';
import { cellsOf, first, fromRange, int, num, optBool, optNum, parseNumber, text, toMatrix, toText } from '../coerce';

export const MAX_TEXT = 32767;
const out = (s: string): string | CellError => (s.length > MAX_TEXT ? Err.value : s);

/** Windows-1252 code points 0x80–0x9F; the five undefined slots map to themselves, as spreadsheets do. */
const CP1252 = [
  0x20ac, 0x81, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8d, 0x017d, 0x8f,
  0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x9d, 0x017e, 0x0178,
];
const charOf = (n: number): string => String.fromCharCode(n >= 0x80 && n <= 0x9f ? CP1252[n - 0x80]! : n);
const codeOf = (ch: string): number => {
  const c = ch.charCodeAt(0);
  if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) return c;
  const i = CP1252.indexOf(c);
  return i >= 0 ? 0x80 + i : 63; // characters outside the code page read as "?"
};

/** Rounds half away from zero on the decimal value, the way spreadsheet's ROUND does. */
export const roundHalfAway = (n: number, d: number): number => {
  if (!Number.isFinite(n)) return n;
  const sign = n < 0 ? -1 : 1, a = Math.abs(n);
  if (d >= 0) {
    const r = Number(`${Math.round(Number(`${Number(a.toPrecision(15))}e${d}`))}e-${d}`);
    return sign * r;
  }
  const f = 10 ** -d;
  return sign * Math.round(Number(a.toPrecision(15)) / f) * f;
};

/** Fixed-point text with optional thousands separators; negative `decimals` round left of the point. */
export const formatFixed = (n: number, decimals: number, commas: boolean): string => {
  const r = roundHalfAway(n, decimals);
  let s = Math.abs(r).toFixed(Math.max(0, Math.min(decimals, 100)));
  if (commas) {
    const [i, f] = s.split('.') as [string, string | undefined];
    s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f !== undefined ? `.${f}` : '');
  }
  return (r < 0 && Number(s.replace(/,/g, '')) !== 0 ? '-' : '') + s;
};

const strArg = (v: Value | undefined, ctx: FnContext, dflt = ''): string | CellError => (v === undefined ? dflt : text(v, ctx));

/** spreadsheet wildcard pattern (* ? ~) as a RegExp source. */
export const wildcardSource = (p: string): string => {
  let s = '';
  for (let i = 0; i < p.length; i++) {
    const ch = p[i]!;
    if (ch === '~' && i + 1 < p.length && '*?~'.includes(p[i + 1]!)) s += `\\${p[++i]}`;
    else if (ch === '*') s += '[\\s\\S]*';
    else if (ch === '?') s += '[\\s\\S]';
    else s += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  return s;
};

const find = (args: Value[], ctx: FnContext, insensitive: boolean): Value => {
  const needle = text(args[0]!, ctx); if (isError(needle)) return needle;
  const hay = text(args[1]!, ctx); if (isError(hay)) return hay;
  const st = optNum(args[2], ctx, 1); if (isError(st)) return st;
  const start = Math.trunc(st);
  if (start < 1 || start > hay.length + 1) return Err.value;
  if (needle === '') return start;
  if (!insensitive) {
    const i = hay.indexOf(needle, start - 1);
    return i < 0 ? Err.value : i + 1;
  }
  const re = new RegExp(wildcardSource(needle), 'giu');
  re.lastIndex = start - 1;
  const m = re.exec(hay);
  return m ? m.index + 1 : Err.value;
};

const left = (args: Value[], ctx: FnContext, right: boolean): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const n = optNum(args[1], ctx, 1); if (isError(n)) return n;
  if (n < 0) return Err.value;
  const k = Math.trunc(n);
  return right ? (k === 0 ? '' : s.slice(-k)) : s.slice(0, k);
};

const mid = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const st = int(args[1]!, ctx); if (isError(st)) return st;
  const n = int(args[2]!, ctx); if (isError(n)) return n;
  if (st < 1 || n < 0) return Err.value;
  return s.substr(st - 1, n);
};

const replace = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const st = int(args[1]!, ctx); if (isError(st)) return st;
  const n = int(args[2]!, ctx); if (isError(n)) return n;
  const r = text(args[3]!, ctx); if (isError(r)) return r;
  if (st < 1 || n < 0) return Err.value;
  return out(s.slice(0, st - 1) + r + s.slice(st - 1 + n));
};

// Full-width ↔ half-width conversion for ASC/DBCS: ASCII, the ideographic space and katakana.
const HALF_KANA = '｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ';
const FULL_KANA = '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン';
const asc = (s: string): string =>
  [...s.normalize('NFD')].map((ch) => {
    const c = ch.charCodeAt(0);
    if (c >= 0xff01 && c <= 0xff5e) return String.fromCharCode(c - 0xfee0);
    if (c === 0x3000) return ' ';
    if (c === 0x3099) return 'ﾞ';
    if (c === 0x309a) return 'ﾟ';
    const k = FULL_KANA.indexOf(ch);
    return k >= 0 ? HALF_KANA[k]! : ch;
  }).join('').normalize('NFC');
const dbcs = (s: string): string =>
  [...s].map((ch) => {
    const c = ch.charCodeAt(0);
    if (c >= 0x21 && c <= 0x7e) return String.fromCharCode(c + 0xfee0);
    if (c === 0x20) return '　';
    if (ch === 'ﾞ') return '゙';
    if (ch === 'ﾟ') return '゚';
    const k = HALF_KANA.indexOf(ch);
    return k >= 0 ? FULL_KANA[k]! : ch;
  }).join('').normalize('NFC');

const THAI_DIGITS = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
const THAI_PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
/** Thai words for a whole number, in groups of six digits joined by ล้าน (million). */
const thaiNumber = (n: number): string => {
  const groups: string[] = [];
  let digits = n.toFixed(0);
  while (digits.length > 6) { groups.unshift(digits.slice(-6)); digits = digits.slice(0, -6); }
  groups.unshift(digits);
  let res = '';
  groups.forEach((g, gi) => {
    let w = '';
    for (let i = 0; i < g.length; i++) {
      const d = +g[i]!, place = g.length - 1 - i;
      if (d === 0) continue;
      if (place === 0 && d === 1 && (res !== '' || w !== '')) w += 'เอ็ด';
      else if (place === 1 && d === 1) w += 'สิบ';
      else if (place === 1 && d === 2) w += 'ยี่สิบ';
      else w += THAI_DIGITS[d]! + THAI_PLACES[place]!;
    }
    res += w + (gi < groups.length - 1 ? 'ล้าน' : '');
  });
  return res;
};
const bahtText = (n: number): string => {
  const neg = n < 0;
  const satangTotal = Math.round(roundHalfAway(Math.abs(n), 2) * 100);
  const baht = Math.floor(satangTotal / 100), satang = satangTotal % 100;
  let s: string;
  if (baht === 0 && satang === 0) s = 'ศูนย์บาทถ้วน';
  else s = (baht ? `${thaiNumber(baht)}บาท` : '') + (satang ? `${thaiNumber(satang)}สตางค์` : 'ถ้วน');
  return (neg ? 'ลบ' : '') + s;
};

/** VALUETOTEXT/ARRAYTOTEXT rendering of one value; strict mode quotes text. */
const valueText = (v: Scalar, strict: boolean): string => {
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  if (isError(v)) return v.code;
  return toText(v) as string;
};
const formatArg = (v: Value | undefined, ctx: FnContext): boolean | CellError => {
  const f = optNum(v, ctx, 0);
  if (isError(f)) return f;
  return f === 0 ? false : f === 1 ? true : Err.value;
};

// ---- Regular expressions: formulas use PCRE2; this maps the common syntax onto JavaScript's engine.
/** Compiles a PCRE2-style pattern; a bad pattern (or one that runs away) is #VALUE!. */
const toRegex = (pattern: string, insensitive: boolean): Regex | CellError => {
  try { return compileRegex(pattern, insensitive); } catch (e) { if (e instanceof RegexError) return Err.value; throw e; }
};
const guard = <T>(f: () => T): T | CellError => {
  try { return f(); } catch (e) { if (e instanceof RegexError) return Err.value; throw e; }
};
const caseArg = (v: Value | undefined, ctx: FnContext): boolean | CellError => {
  const c = optNum(v, ctx, 0);
  if (isError(c)) return c;
  const k = Math.trunc(c);
  return k === 0 ? false : k === 1 ? true : Err.value;
};

const regexTest = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const p = text(args[1]!, ctx); if (isError(p)) return p;
  const ci = caseArg(args[2], ctx); if (isError(ci)) return ci;
  const re = toRegex(p, ci); if (isError(re)) return re;
  return guard(() => re.test(s));
};

const regexExtract = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const p = text(args[1]!, ctx); if (isError(p)) return p;
  const mode = optNum(args[2], ctx, 0); if (isError(mode)) return mode;
  const ci = caseArg(args[3], ctx); if (isError(ci)) return ci;
  const md = Math.trunc(mode);
  if (md !== 0 && md !== 1 && md !== 2) return Err.value;
  const re = toRegex(p, ci); if (isError(re)) return re;
  return guard((): Value => {
    if (md === 1) {
      const all = re.matchAll(s).map((m) => [m.whole] as Scalar[]);
      return all.length ? all : Err.na;
    }
    const m = re.exec(s);
    if (!m) return Err.na;
    if (md === 0) return m.whole;
    return m.groups.length ? [m.groups.map((g) => g ?? '')] : Err.na;
  });
};

const regexReplace = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const p = text(args[1]!, ctx); if (isError(p)) return p;
  const r = text(args[2]!, ctx); if (isError(r)) return r;
  const occ = optNum(args[3], ctx, 0); if (isError(occ)) return occ;
  const ci = caseArg(args[4], ctx); if (isError(ci)) return ci;
  const re = toRegex(p, ci); if (isError(re)) return re;
  return guard((): Value => {
    const all = re.matchAll(s);
    const n = Math.trunc(occ);
    const target = n === 0 ? 0 : n > 0 ? n : all.length + n + 1;
    if (n !== 0 && (target < 1 || target > all.length)) return s;
    const cps = toCodePoints(s);
    const slice = (a: number, b: number): string => String.fromCodePoint(...cps.slice(a, b));
    let res = '', at = 0;
    all.forEach((m, idx) => {
      if (target !== 0 && idx + 1 !== target) return;
      res += slice(at, m.start) + expandReplacement(r, re, m);
      at = m.end;
    });
    res += slice(at, cps.length);
    return out(res);
  });
};

// ---- TEXTBEFORE / TEXTAFTER / TEXTSPLIT

const delimiters = (v: Value, ctx: FnContext): string[] | CellError => {
  const ds: string[] = [];
  for (const c of cellsOf(v, ctx)) {
    const t = toText(c);
    if (isError(t)) return t;
    ds.push(t);
  }
  return ds;
};

/** Every delimiter match in `s`, left to right, non-overlapping, longest delimiter first at a position. */
const delimMatches = (s: string, ds: string[], ci: boolean): Array<[number, number]> => {
  const hay = ci ? s.toLowerCase() : s;
  const needles = ds.filter((d) => d !== '').map((d) => (ci ? d.toLowerCase() : d)).sort((a, b) => b.length - a.length);
  const res: Array<[number, number]> = [];
  for (let i = 0; i < hay.length;) {
    const d = needles.find((x) => hay.startsWith(x, i));
    if (d) { res.push([i, d.length]); i += d.length; } else i++;
  }
  return res;
};

const textAround = (args: Value[], ctx: FnContext, after: boolean): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const ds = delimiters(args[1]!, ctx); if (isError(ds)) return ds;
  const inst = optNum(args[2], ctx, 1); if (isError(inst)) return inst;
  const mm = optNum(args[3], ctx, 0); if (isError(mm)) return mm;
  const me = optNum(args[4], ctx, 0); if (isError(me)) return me;
  const n = Math.trunc(inst);
  if (n === 0 || Math.abs(n) > s.length || (mm !== 0 && mm !== 1)) return Err.value;
  const notFound = (): Value => (args[5] === undefined ? Err.na : first(args[5], ctx));
  if (ds.some((d) => d === '')) return (n > 0) === after ? s : ''; // an empty delimiter matches at the start (or end, counting back)
  const ms = delimMatches(s, ds, mm === 1);
  if (me) { if (n > 0) ms.push([s.length, 0]); else ms.unshift([0, 0]); }
  const m = n > 0 ? ms[n - 1] : ms[ms.length + n];
  if (!m) return notFound();
  return after ? s.slice(m[0] + m[1]) : s.slice(0, m[0]);
};

const splitBy = (s: string, ds: string[], ci: boolean): string[] => {
  const parts: string[] = [];
  let pos = 0;
  for (const [i, len] of delimMatches(s, ds, ci)) { parts.push(s.slice(pos, i)); pos = i + len; }
  parts.push(s.slice(pos));
  return parts;
};

const textSplit = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const optDelims = (v: Value | undefined): string[] | CellError => (v === undefined || v === null ? [] : delimiters(v, ctx));
  const cd = optDelims(args[1]); if (isError(cd)) return cd;
  const rd = optDelims(args[2]); if (isError(rd)) return rd;
  const ignore = optBool(args[3], ctx, false); if (isError(ignore)) return ignore;
  const mm = optNum(args[4], ctx, 0); if (isError(mm)) return mm;
  if (!cd.some((d) => d !== '') && !rd.some((d) => d !== '')) return Err.value;
  const pad: Scalar = args[5] === undefined ? Err.na : first(args[5], ctx);
  const ci = mm === 1;
  let rows = rd.length ? splitBy(s, rd, ci) : [s];
  if (ignore) rows = rows.filter((r) => r !== '');
  let grid = rows.map((r) => (cd.length ? splitBy(r, cd, ci) : [r]));
  if (ignore) grid = grid.map((r) => r.filter((c) => c !== '')).filter((r) => r.length);
  if (!grid.length) return Err.calc;
  const width = Math.max(...grid.map((r) => r.length));
  return grid.map((r) => [...r, ...Array<Scalar>(width - r.length).fill(pad)]);
};

const textJoin = (args: Value[], ctx: FnContext): Value => {
  const dm = toMatrix(args[0]!, ctx); if (isError(dm)) return dm;
  const delims: string[] = [];
  for (const row of dm) for (const c of row) { const t = toText(c); if (isError(t)) return t; delims.push(t); }
  const ignore = optBool(args[1], ctx, false); if (isError(ignore)) return ignore;
  const parts: string[] = [];
  for (const a of args.slice(2)) {
    for (const c of cellsOf(a, ctx)) {
      const t = toText(c);
      if (isError(t)) return t;
      if (ignore && t === '') continue;
      parts.push(t);
    }
  }
  let res = parts[0] ?? '';
  for (let i = 1; i < parts.length; i++) res += (delims.length ? delims[(i - 1) % delims.length] : '') + parts[i];
  return out(res);
};

const concat = (args: Value[], ctx: FnContext): Value => {
  let s = '';
  for (const a of args) for (const c of cellsOf(a, ctx)) {
    const t = toText(c);
    if (isError(t)) return t;
    s += t;
    if (s.length > MAX_TEXT) return Err.value;
  }
  return s;
};

const numberValue = (args: Value[], ctx: FnContext): Value => {
  const s = text(args[0]!, ctx); if (isError(s)) return s;
  const dsep = strArg(args[1] ?? undefined, ctx, '.'); if (isError(dsep)) return dsep;
  const gsep = strArg(args[2] ?? undefined, ctx, ','); if (isError(gsep)) return gsep;
  const dec = (args[1] === null ? '.' : dsep)[0], grp = (args[2] === null ? ',' : gsep)[0];
  if (!dec || !grp || dec === grp) return Err.value;
  let t = s.replace(/\s+/g, '');
  if (t === '') return 0;
  let pct = 0;
  while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
  const parts = t.split(dec);
  if (parts.length > 2) return Err.value;
  const frac = parts[1] ?? '';
  if (frac.includes(grp)) return Err.value;
  const whole = parts[0]!.split(grp).join('');
  const lit = whole + (parts.length > 1 ? `.${frac}` : '');
  if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(lit)) return Err.value;
  return Number(lit) / 100 ** pct;
};

const scalarText = (name: string, f: (s: string) => string, min = 1): FunctionDef => ({
  name, min, max: 1, scalar: true,
  fn: ([v], ctx) => { const s = text(v!, ctx); return isError(s) ? s : out(f(s)); },
});

export const defs: FunctionDef[] = [
  {
    name: 'ARRAYTOTEXT', min: 1, max: 2,
    fn: ([v, f], ctx) => {
      const strict = formatArg(f, ctx); if (isError(strict)) return strict;
      const m = toMatrix(v!, ctx); if (isError(m)) return m;
      const res = strict
        ? `{${m.map((r) => r.map((c) => valueText(c, true)).join(',')).join(';')}}`
        : m.flat().map((c) => valueText(c, false)).join(', ');
      return out(res);
    },
  },
  scalarText('ASC', asc),
  scalarText('DBCS', dbcs),
  {
    name: 'BAHTTEXT', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const n = num(v!, ctx); return isError(n) ? n : bahtText(n); },
  },
  {
    name: 'CHAR', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const n = int(v!, ctx); return isError(n) ? n : n < 1 || n > 255 ? Err.value : charOf(n); },
  },
  scalarText('CLEAN', (s) => s.replace(/[\x00-\x1f]/g, '')),
  {
    name: 'CODE', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const s = text(v!, ctx); return isError(s) ? s : s === '' ? Err.value : codeOf(s); },
  },
  { name: 'CONCAT', min: 1, max: 255, fn: concat },
  { name: 'CONCATENATE', min: 1, max: 255, scalar: true, fn: concat },
  {
    name: 'DOLLAR', min: 1, max: 2, scalar: true,
    fn: ([v, d], ctx) => {
      const n = num(v!, ctx); if (isError(n)) return n;
      const dd = optNum(d, ctx, 2); if (isError(dd)) return dd;
      if (dd > 127) return Err.value;
      const s = formatFixed(Math.abs(n), Math.trunc(dd), true);
      return roundHalfAway(n, Math.trunc(dd)) < 0 ? `($${s})` : `$${s}`;
    },
  },
  {
    name: 'EXACT', min: 2, max: 2, scalar: true,
    fn: ([a, b], ctx) => { const x = text(a!, ctx); if (isError(x)) return x; const y = text(b!, ctx); return isError(y) ? y : x === y; },
  },
  { name: 'FIND', min: 2, max: 3, scalar: true, fn: (a, ctx) => find(a, ctx, false) },
  { name: 'FINDB', min: 2, max: 3, scalar: true, fn: (a, ctx) => find(a, ctx, false) },
  {
    name: 'FIXED', min: 1, max: 3, scalar: true,
    fn: ([v, d, nc], ctx) => {
      const n = num(v!, ctx); if (isError(n)) return n;
      const dd = optNum(d, ctx, 2); if (isError(dd)) return dd;
      const noCommas = optBool(nc, ctx, false); if (isError(noCommas)) return noCommas;
      if (dd > 127) return Err.value;
      return formatFixed(n, Math.trunc(dd), !noCommas);
    },
  },
  { name: 'LEFT', min: 1, max: 2, scalar: true, fn: (a, ctx) => left(a, ctx, false) },
  { name: 'LEFTB', min: 1, max: 2, scalar: true, fn: (a, ctx) => left(a, ctx, false) },
  { name: 'RIGHT', min: 1, max: 2, scalar: true, fn: (a, ctx) => left(a, ctx, true) },
  { name: 'RIGHTB', min: 1, max: 2, scalar: true, fn: (a, ctx) => left(a, ctx, true) },
  { name: 'LEN', min: 1, max: 1, scalar: true, fn: ([v], ctx) => { const s = text(v!, ctx); return isError(s) ? s : s.length; } },
  { name: 'LENB', min: 1, max: 1, scalar: true, fn: ([v], ctx) => { const s = text(v!, ctx); return isError(s) ? s : s.length; } },
  scalarText('LOWER', (s) => s.toLowerCase()),
  scalarText('UPPER', (s) => s.toUpperCase()),
  scalarText('PROPER', (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase())),
  { name: 'MID', min: 3, max: 3, scalar: true, fn: mid },
  { name: 'MIDB', min: 3, max: 3, scalar: true, fn: mid },
  { name: 'NUMBERVALUE', min: 1, max: 3, scalar: true, fn: numberValue },
  {
    name: 'PHONETIC', min: 1, max: 1,
    fn: ([v], ctx) => {
      if (!fromRange(v!) && !isError(v)) return Err.na;
      let s = '';
      for (const c of cellsOf(v!, ctx)) { if (isError(c)) return c; if (typeof c === 'string') s += c; }
      return out(s);
    },
  },
  { name: 'REGEXEXTRACT', min: 2, max: 4, scalar: true, fn: regexExtract },
  { name: 'REGEXREPLACE', min: 3, max: 5, scalar: true, fn: regexReplace },
  { name: 'REGEXTEST', min: 2, max: 3, scalar: true, fn: regexTest },
  { name: 'REPLACE', min: 4, max: 4, scalar: true, fn: replace },
  { name: 'REPLACEB', min: 4, max: 4, scalar: true, fn: replace },
  {
    name: 'REPT', min: 2, max: 2, scalar: true,
    fn: ([a, b], ctx) => {
      const s = text(a!, ctx); if (isError(s)) return s;
      const n = int(b!, ctx); if (isError(n)) return n;
      if (n < 0 || s.length * n > MAX_TEXT) return Err.value;
      return s.repeat(n);
    },
  },
  { name: 'SEARCH', min: 2, max: 3, scalar: true, fn: (a, ctx) => find(a, ctx, true) },
  { name: 'SEARCHB', min: 2, max: 3, scalar: true, fn: (a, ctx) => find(a, ctx, true) },
  {
    name: 'SUBSTITUTE', min: 3, max: 4, scalar: true,
    fn: ([a, b, c, d], ctx) => {
      const s = text(a!, ctx); if (isError(s)) return s;
      const o = text(b!, ctx); if (isError(o)) return o;
      const r = text(c!, ctx); if (isError(r)) return r;
      if (d === undefined) return o === '' ? s : out(s.split(o).join(r));
      const k = int(d, ctx); if (isError(k)) return k;
      if (k < 1) return Err.value;
      if (o === '') return s;
      let i = -o.length;
      for (let n = 0; n < k; n++) { i = s.indexOf(o, i + o.length); if (i < 0) return s; }
      return out(s.slice(0, i) + r + s.slice(i + o.length));
    },
  },
  { name: 'T', min: 1, max: 1, scalar: true, fn: ([v]) => (typeof v === 'string' || isError(v) ? v : '') },
  { name: 'TEXTAFTER', min: 2, max: 6, scalar: [0, 2, 3, 4, 5], fn: (a, ctx) => textAround(a, ctx, true) },
  { name: 'TEXTBEFORE', min: 2, max: 6, scalar: [0, 2, 3, 4, 5], fn: (a, ctx) => textAround(a, ctx, false) },
  { name: 'TEXTJOIN', min: 3, max: 255, fn: textJoin },
  { name: 'TEXTSPLIT', min: 2, max: 6, scalar: [0, 3, 4, 5], fn: textSplit },
  scalarText('TRIM', (s) => s.replace(/ +/g, ' ').replace(/^ | $/g, '')),
  {
    name: 'UNICHAR', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      const n = int(v!, ctx); if (isError(n)) return n;
      if (n < 1 || n > 0x10ffff) return Err.value;
      if (n >= 0xd800 && n <= 0xdfff) return Err.na;
      return String.fromCodePoint(n);
    },
  },
  {
    name: 'UNICODE', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => { const s = text(v!, ctx); return isError(s) ? s : s === '' ? Err.value : s.codePointAt(0)!; },
  },
  {
    name: 'VALUE', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      if (v === null || typeof v === 'number' || isError(v)) return v ?? 0;
      if (typeof v !== 'string') return Err.value;
      return parseNumber(v, ctx.date1904) ?? Err.value;
    },
  },
  {
    name: 'VALUETOTEXT', min: 1, max: 2, scalar: [0],
    fn: ([v, f], ctx) => {
      const strict = formatArg(f, ctx); if (isError(strict)) return strict;
      return out(valueText(isLambda(v!) ? Err.calc : (v as Scalar), strict));
    },
  },
];
