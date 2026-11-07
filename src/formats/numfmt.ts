// spreadsheet number formatting (ECMA-376 Part 1, 18.8.31 numFmts) with spreadsheet's display rules.
import { isError, type Scalar } from '../engine/types';
import { fromSerial, weekdayOf } from '../engine/date';

export interface Formatted {
  text: string;
  color?: string;
  align?: 'left' | 'right';
  /** Repeat character from `*x`; the UI repeats it at `fillAt` to fill the cell width. */
  fill?: string;
  fillAt?: number;
}

/** The legacy indexed palette (indexes 0-63, plus 64/65 system foreground/background). */
export const INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
].map((c) => '#' + c);

const NAMED_COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
  magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
};

const BUILTIN: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm',
  37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};
// East-Asian locale built-ins (27-36, 50-58) are dates or times; map them to the en-US equivalents.
for (const i of [27, 28, 29, 30, 31, 36, 50, 51, 52, 53, 54, 57, 58]) BUILTIN[i] = 'm/d/yyyy';
for (const i of [32, 33, 34, 35, 55, 56]) BUILTIN[i] = 'h:mm:ss';

export const builtinFormat = (id: number): string | undefined => BUILTIN[id];

/** The built-in id for a code, when it has one (used by the writer so ids 0-49 stay built-in). */
export const builtinId = (code: string): number | undefined => {
  for (const k in BUILTIN) if (BUILTIN[k] === code && (+k < 27 || +k > 36) && +k < 50) return +k;
  return undefined;
};

// ---------------------------------------------------------------- lexing

type Tok =
  | { k: 'lit'; v: string }
  | { k: 'dig'; v: string } // literal digit 1-9 (fraction denominators)
  | { k: 'ph'; v: '0' | '#' | '?' }
  | { k: 'dot' } | { k: 'comma' } | { k: 'pct' } | { k: 'slash' } | { k: 'at' } | { k: 'gen' }
  | { k: 'exp'; v: '+' | '-' }
  | { k: 'fill'; v: string }
  | { k: 'dt'; v: string; n?: number };

type Kind = 'num' | 'date' | 'text' | 'general';
interface Cond { op: string; v: number }
interface Section { toks: Tok[]; color?: string; cond?: Cond; kind: Kind; ampm: boolean; sub: number }

const splitSections = (code: string): string[] => {
  const out: string[] = [];
  let cur = '';
  for (let i = 0; i < code.length; i++) {
    const c = code[i]!;
    if (c === '"') {
      const e = code.indexOf('"', i + 1);
      const end = e < 0 ? code.length : e;
      cur += code.slice(i, end + 1);
      i = end;
    } else if (c === '\\' || c === '_' || c === '*') { cur += c + (code[i + 1] ?? ''); i++; }
    else if (c === '[') {
      const e = code.indexOf(']', i);
      const end = e < 0 ? code.length - 1 : e;
      cur += code.slice(i, end + 1);
      i = end;
    } else if (c === ';') { out.push(cur); cur = ''; }
    else cur += c;
  }
  out.push(cur);
  return out;
};

const lexSection = (src: string): Section => {
  const toks: Tok[] = [];
  let color: string | undefined;
  let cond: Cond | undefined;
  const lit = (v: string): void => {
    const last = toks[toks.length - 1];
    if (last && last.k === 'lit') last.v += v;
    else toks.push({ k: 'lit', v });
  };
  for (let i = 0; i < src.length; i++) {
    const c = src[i]!;
    const lc = c.toLowerCase();
    if (c === '"') {
      const e = src.indexOf('"', i + 1);
      const end = e < 0 ? src.length : e;
      lit(src.slice(i + 1, end));
      i = end;
    } else if (c === '\\') { lit(src[i + 1] ?? ''); i++; }
    else if (c === '_') { lit(' '); i++; }
    else if (c === '*') { toks.push({ k: 'fill', v: src[i + 1] ?? ' ' }); i++; }
    else if (c === '[') {
      const e = src.indexOf(']', i);
      const end = e < 0 ? src.length : e;
      const body = src.slice(i + 1, end);
      i = end;
      const lb = body.toLowerCase();
      let m: RegExpExecArray | null;
      if (NAMED_COLORS[lb]) color = NAMED_COLORS[lb];
      else if ((m = /^color\s*(\d+)$/.exec(lb))) color = INDEXED_COLORS[+m[1]! + 7] ?? color;
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:e[+-]?\d+)?)$/.exec(lb))) cond = { op: m[1]!, v: +m[2]! };
      else if (/^(h+|m+|s+)$/.test(lb)) toks.push({ k: 'dt', v: `[${lb[0]}]` });
      else if (body[0] === '$') {
        const dash = body.indexOf('-', 1);
        const cur = body.slice(1, dash < 0 ? undefined : dash);
        if (cur) lit(cur);
      }
      // other tags ([DBNum1], [NatNum], ...) are ignored
    } else if (lc === 'g' && /^general/i.test(src.slice(i, i + 7))) { toks.push({ k: 'gen' }); i += 6; }
    else if (c === '0' || c === '#' || c === '?') toks.push({ k: 'ph', v: c });
    else if (c >= '1' && c <= '9') toks.push({ k: 'dig', v: c });
    else if (c === '.') toks.push({ k: 'dot' });
    else if (c === ',') toks.push({ k: 'comma' });
    else if (c === '%') toks.push({ k: 'pct' });
    else if (c === '/') toks.push({ k: 'slash' });
    else if (c === '@') toks.push({ k: 'at' });
    else if ((c === 'E' || c === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) { toks.push({ k: 'exp', v: src[i + 1] as '+' | '-' }); i++; }
    else if (lc === 'a' && /^am\/pm/i.test(src.slice(i, i + 5))) { toks.push({ k: 'dt', v: 'ampm' }); i += 4; }
    else if (lc === 'a' && /^a\/p/i.test(src.slice(i, i + 3))) { toks.push({ k: 'dt', v: src[i]! === 'a' ? 'ap' : 'AP', n: src[i + 2] === 'p' ? 1 : 0 }); i += 2; }
    else if ('ymdhse'.includes(lc)) {
      let j = i;
      while (j + 1 < src.length && src[j + 1]!.toLowerCase() === lc) j++;
      toks.push({ k: 'dt', v: (lc === 'e' ? 'y' : lc).repeat(lc === 'e' ? 4 : j - i + 1) });
      i = j;
    } else if (lc === 'b' && (src[i + 1] === '1' || src[i + 1] === '2')) i++;
    else lit(c);
  }
  const hasDt = toks.some((t) => t.k === 'dt');
  let sub = 0;
  if (hasDt) {
    // In date sections ".0", ".00", ".000" after seconds are fractional seconds; other placeholders are literal.
    const out: Tok[] = [];
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i]!;
      if (t.k === 'dot' && toks[i + 1]?.k === 'ph' && (toks[i + 1] as { v: string }).v === '0') {
        let n = 0;
        while (toks[i + 1]?.k === 'ph' && (toks[i + 1] as { v: string }).v === '0') { n++; i++; }
        out.push({ k: 'dt', v: 'sub', n });
        sub = Math.max(sub, Math.min(n, 3));
      } else if (t.k === 'dt' || t.k === 'fill' || t.k === 'at') out.push(t);
      else out.push({ k: 'lit', v: t.k === 'lit' || t.k === 'dig' || t.k === 'ph' ? t.v : t.k === 'dot' ? '.' : t.k === 'comma' ? ',' : t.k === 'pct' ? '%' : t.k === 'slash' ? '/' : t.k === 'exp' ? 'E' + t.v : 'General' });
    }
    toks.length = 0;
    toks.push(...out);
    // m/mm right after an hour or right before seconds is minutes.
    const dts = toks.filter((t): t is Extract<Tok, { k: 'dt' }> => t.k === 'dt');
    dts.forEach((t, i) => {
      if (t.v !== 'm' && t.v !== 'mm') return;
      const prev = dts[i - 1]?.v ?? '';
      const next = dts[i + 1]?.v ?? '';
      if (prev[0] === 'h' || prev === '[h]' || next[0] === 's' || next === '[s]') t.v = t.v === 'm' ? 'M' : 'MM';
    });
  }
  const kind: Kind = hasDt ? 'date'
    : toks.some((t) => t.k === 'at') ? 'text'
      : toks.some((t) => t.k === 'gen') ? 'general' : 'num';
  return { toks, color, cond, kind, ampm: toks.some((t) => t.k === 'dt' && (t.v === 'ampm' || t.v === 'ap' || t.v === 'AP')), sub };
};

const cache = new Map<string, Section[]>();
const SYS_LONG_DATE = 'dddd, mmmm d, yyyy';
const SYS_TIME = 'h:mm:ss AM/PM';

const parseFormat = (code: string): Section[] => {
  let hit = cache.get(code);
  if (hit) return hit;
  let c = code;
  if (/\[\$-F800\]/i.test(c)) c = SYS_LONG_DATE;
  else if (/\[\$-F400\]/i.test(c)) c = SYS_TIME;
  hit = splitSections(c).slice(0, 4).map(lexSection);
  if (cache.size > 2000) cache.clear();
  cache.set(code, hit);
  return hit;
};

export const isDateFormat = (code: string): boolean => {
  const secs = parseFormat(code);
  return secs[0]?.kind === 'date';
};

// ---------------------------------------------------------------- decimal helpers

/** a >= 0 rounded half away from zero to d decimals at 15 significant digits: [intDigits ('' for 0), fracDigits]. */
const fixed = (a: number, d: number): [string, string] => {
  if (a === 0) return ['', '0'.repeat(d)];
  const [mant, ex] = a.toExponential(14).split('e') as [string, string];
  return roundDigits(mant.replace('.', ''), +ex + 1, d);
};

const roundDigits = (digits0: string, point0: number, d: number): [string, string] => {
  let digits = digits0;
  let point = point0;
  if (point < 0) { digits = '0'.repeat(-point) + digits; point = 0; }
  const keep = point + d;
  if (keep < digits.length) {
    const up = digits.charCodeAt(keep) >= 53;
    digits = digits.slice(0, keep);
    if (up) {
      const inc = increment(digits);
      if (inc.length > digits.length) point++;
      digits = inc;
    }
  } else digits = digits.padEnd(keep, '0');
  return [digits.slice(0, point).replace(/^0+/, ''), digits.slice(point)];
};

const increment = (s: string): string => {
  const a = s.split('');
  let i = a.length - 1;
  for (; i >= 0; i--) {
    if (a[i] === '9') a[i] = '0';
    else { a[i] = String.fromCharCode(a[i]!.charCodeAt(0) + 1); break; }
  }
  return i < 0 ? '1' + a.join('') : a.join('');
};

/** Significant digits and decimal exponent (a = 0.D × 10^point... expressed as digits with `point` integer digits). */
const sig = (a: number): { digits: string; e10: number } => {
  const [mant, ex] = a.toExponential(14).split('e') as [string, string];
  return { digits: mant.replace('.', '').replace(/0+$/, '') || '0', e10: +ex };
};

const plainDecimal = (a: number): string => {
  const { digits, e10 } = sig(a);
  if (e10 >= digits.length - 1) return digits + '0'.repeat(e10 - digits.length + 1);
  if (e10 >= 0) return digits.slice(0, e10 + 1) + '.' + digits.slice(e10 + 1);
  return '0.' + '0'.repeat(-e10 - 1) + digits;
};

const expStr = (e: number, minDigits = 2): string => 'E' + (e < 0 ? '-' : '+') + String(Math.abs(e)).padStart(minDigits, '0');

const sigCount = (s: string): number => s.replace(/E.*$/, '').replace(/[^0-9]/g, '').replace(/^0+/, '').length;

/** spreadsheet's General format for a number in a column `width` characters wide (11 in a standard column). */
export const formatGeneral = (n: number, cellWidth = 11): string => {
  let width = cellWidth;
  if (n === 0) return '0';
  if (!Number.isFinite(n)) return '#NUM!';
  const sign = n < 0 ? '-' : '';
  const a = Math.abs(n);
  width -= sign.length;
  const plain = plainDecimal(a);
  const { e10 } = sig(a);
  if (plain.length <= width && e10 >= -4) return sign + plain;
  let dec: string | undefined;
  const intDigits = e10 >= 0 ? e10 + 1 : 1;
  if (intDigits <= width && e10 >= -4) {
    const [i, f] = fixed(a, Math.max(0, width - intDigits - 1));
    const s = ((i || '0') + (f ? '.' + f : '')).replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
    if (s !== '0' && s.length <= width) dec = s;
  }
  let e = e10;
  const md = Math.max(0, width - 2 - expStr(e).length);
  const all = a.toExponential(14).split('e')[0]!.replace('.', '');
  let [mi, mf] = roundDigits(all, 1, md);
  if (mi.length > 1) { e++; [mi, mf] = roundDigits(all, 0, md); }
  mf = mf.replace(/0+$/, '');
  const sci = (mi || '0') + (mf ? '.' + mf : '') + expStr(e);
  if (dec === undefined) return sign + sci;
  return sign + (sigCount(dec) >= sigCount(sci) ? dec : sci);
};

// ---------------------------------------------------------------- rendering

class Out {
  s = '';
  fill?: string;
  fillAt?: number;
  add(t: string): void { this.s += t; }
  setFill(c: string): void { if (this.fill === undefined) { this.fill = c; this.fillAt = this.s.length; } }
}

const passThrough = (t: Tok, o: Out): boolean => {
  if (t.k === 'lit' || t.k === 'dig') o.add(t.v);
  else if (t.k === 'fill') o.setFill(t.v);
  else return false;
  return true;
};

/** Writes integer digits D into placeholders right-to-left; extra digits go into the leftmost one. */
const fillInt = (D: string, phs: { v: string }[], group: boolean): string[] => {
  const res: string[] = new Array<string>(phs.length).fill('');
  let p = D.length - 1;
  let count = 0;
  const put = (ch: string, isDigit: boolean): string => {
    let s = ch;
    if (group && isDigit && count > 0 && count % 3 === 0) s += ',';
    if (isDigit) count++;
    return s;
  };
  for (let i = phs.length - 1; i >= 0; i--) {
    const v = phs[i]!.v;
    let s = '';
    if (p >= 0) s = put(D[p--]!, true);
    else if (v === '0') s = put('0', true);
    else if (v === '?') s = ' ';
    if (i === 0) while (p >= 0) s = put(D[p--]!, true) + s;
    res[i] = s;
  }
  return res;
};

/** Decimal digits F into placeholders, trimming trailing zeros for # and ?. */
const fillDec = (F: string, phs: { v: string }[]): string[] => {
  const res = phs.map((_, i) => F[i] ?? '0');
  for (let i = phs.length - 1; i >= 0; i--) {
    if (res[i] !== '0' || phs[i]!.v === '0') break;
    res[i] = phs[i]!.v === '?' ? ' ' : '';
  }
  return res;
};

const gcd = (a: number, b: number): number => (b ? gcd(b, a % b) : a);

/** Closest fraction to x (0 <= x < 1) with denominator <= maxDen. */
const bestFraction = (x: number, maxDen: number): [number, number] => {
  let p0 = 0, q0 = 1, p1 = 1, q1 = 0;
  let r = x;
  for (let k = 0; k < 64; k++) {
    const a = Math.floor(r);
    const q2 = q0 + a * q1;
    if (q2 > maxDen) break;
    [p0, q0, p1, q1] = [p1, q1, p0 + a * p1, q2];
    if (r - a < 1e-12) break;
    r = 1 / (r - a);
  }
  // semiconvergent candidate
  const t = q1 === 0 ? 0 : Math.floor((maxDen - q0) / q1);
  const cands: [number, number][] = [[p1, q1], [p0 + t * p1, q0 + t * q1]];
  let best: [number, number] = [Math.round(x), 1];
  let err = Math.abs(x - best[0]);
  for (const [p, q] of cands) {
    if (q <= 0 || q > maxDen) continue;
    const e = Math.abs(x - p / q);
    if (e < err - 1e-15) { best = [p, q]; err = e; }
  }
  return best;
};

const renderNumber = (sec: Section, v: number, o: Out): void => {
  const toks = sec.toks;
  let a = Math.abs(v);
  const neg = v < 0;
  const phIdx: number[] = [];
  toks.forEach((t, i) => { if (t.k === 'ph') phIdx.push(i); });
  const pct = toks.filter((t) => t.k === 'pct').length;
  for (let i = 0; i < pct; i++) a *= 100;
  const slash = toks.findIndex((t, i) => t.k === 'slash' && phIdx.some((p) => p < i) && toks.slice(i + 1).some((x) => x.k === 'ph' || x.k === 'dig'));
  const expI = toks.findIndex((t) => t.k === 'exp');
  const firstPh = phIdx[0] ?? -1;
  const lastPh = phIdx[phIdx.length - 1] ?? -1;
  const dot = toks.findIndex((t, i) => t.k === 'dot' && (expI < 0 || i < expI) && (slash < 0));
  // Commas: between integer placeholders → grouping; right after the last placeholder of the number → scale by 1000.
  let group = false;
  let scale = 0;
  const intEnd = dot >= 0 ? dot : expI >= 0 ? expI : lastPh + 1;
  toks.forEach((t, i) => {
    if (t.k !== 'comma' || firstPh < 0 || i < firstPh) return;
    const nextPh = phIdx.find((p) => p > i);
    if (i < intEnd && nextPh !== undefined && nextPh < intEnd) group = true;
    else if (nextPh === undefined || (dot >= 0 && i < dot && !toks.slice(i + 1, dot).some((x) => x.k === 'ph'))) scale++;
  });
  for (let i = 0; i < scale; i++) a /= 1000;
  const out: string[] = toks.map(() => '');
  const sign = neg ? '-' : '';

  if (slash >= 0) {
    // numerator placeholders: the contiguous run just before '/'; whole part: placeholders before that run.
    let ns = slash - 1;
    while (ns >= 0 && toks[ns]!.k !== 'ph') ns--;
    let ne = ns;
    while (ne - 1 >= 0 && toks[ne - 1]!.k === 'ph') ne--;
    const numPh = phIdx.filter((p) => p >= ne && p <= ns);
    const wholePh = phIdx.filter((p) => p < ne);
    const denToks: number[] = [];
    for (let i = slash + 1; i < toks.length && (toks[i]!.k === 'ph' || toks[i]!.k === 'dig'); i++) denToks.push(i);
    const fixedDen = denToks.some((i) => toks[i]!.k === 'dig') ? +denToks.map((i) => (toks[i] as { v: string }).v).join('') : 0;
    const denPhCount = denToks.length;
    let whole = wholePh.length ? Math.floor(a) : 0;
    const frac = wholePh.length ? a - whole : a;
    let num: number, den: number;
    if (fixedDen) { den = fixedDen; num = Math.round(frac * den); }
    else {
      const ip = wholePh.length ? 0 : Math.floor(frac);
      [num, den] = bestFraction(frac - ip, Math.pow(10, Math.min(denPhCount, 7)) - 1);
      num += ip * den;
    }
    if (wholePh.length && num >= den) { whole += Math.floor(num / den); num %= den; }
    const g = fixedDen ? 1 : gcd(num, den) || 1;
    if (!fixedDen && num) { num /= g; den /= g; }
    const blankFrac = wholePh.length > 0 && num === 0;
    if (wholePh.length) {
      const D = whole ? String(whole) : blankFrac ? '0' : '';
      fillInt(D, wholePh.map((i) => toks[i] as { v: string }), false).forEach((s, j) => { out[wholePh[j]!] = s; });
    }
    const numStr = blankFrac ? '' : !wholePh.length && num === 0 ? '0' : String(num);
    fillInt(numStr, numPh.map((i) => toks[i] as { v: string }), false).forEach((s, j) => { out[numPh[j]!] = s; });
    if (fixedDen) denToks.forEach((i, j) => { out[i] = j === 0 ? String(den) : ''; });
    else {
      const ds = blankFrac ? '' : String(den);
      denToks.forEach((i, j) => {
        const v2 = (toks[i] as { v: string }).v;
        out[i] = j === 0 ? ds.slice(0, Math.max(1, ds.length - denToks.length + 1)) : ds[ds.length - denToks.length + j] ?? (v2 === '0' && !blankFrac ? '0' : v2 === '#' ? '' : ' ');
      });
      if (ds.length < denToks.length && !blankFrac) {
        // left-align the denominator: digits first, padding after
        const pads = denToks.map((i) => (toks[i] as { v: string }).v);
        const cells = ds.split('');
        for (let j = ds.length; j < denToks.length; j++) cells.push(pads[j] === '0' ? '0' : pads[j] === '#' ? '' : ' ');
        denToks.forEach((i, j) => { out[i] = cells[j] ?? ''; });
      }
    }
    if (blankFrac) {
      // the fraction (and the separator before it) becomes spaces of the same width
      let start = ne;
      while (start - 1 >= 0 && toks[start - 1]!.k === 'lit' && wholePh.length && start - 1 > wholePh[wholePh.length - 1]!) start--;
      for (let i = start; i <= (denToks[denToks.length - 1] ?? slash); i++) {
        const t = toks[i]!;
        out[i] = ' '.repeat(t.k === 'lit' ? t.v.length : 1);
      }
    }
    o.add(sign);
    toks.forEach((t, i) => {
      if (i >= ne && i <= (denToks[denToks.length - 1] ?? slash) && blankFrac) { o.add(out[i]!); return; }
      if (t.k === 'ph' || i === slash || denToks.includes(i)) o.add(i === slash ? '/' : out[i]!);
      else if (t.k === 'pct') o.add('%');
      else if (t.k === 'dot') o.add('.');
      else if (t.k === 'gen') o.add(formatGeneral(a));
      else passThrough(t, o);
    });
    return;
  }

  const intPh = phIdx.filter((p) => p < intEnd);
  const decPh = phIdx.filter((p) => p > intEnd && (expI < 0 || p < expI));
  const expPh = expI >= 0 ? phIdx.filter((p) => p > expI) : [];
  let I: string, F: string;
  let expText = '';
  if (expI >= 0) {
    const nInt = intPh.length;
    const eng = nInt > 1 && intPh.some((p) => (toks[p] as { v: string }).v === '#');
    let X = 0;
    let r: [string, string] = ['', '0'.repeat(decPh.length)];
    if (a !== 0) {
      const all = a.toExponential(14).split('e');
      const digits = all[0]!.replace('.', '');
      const e10 = +all[1]!;
      X = eng ? Math.floor(e10 / nInt) * nInt : e10 - (nInt > 0 ? nInt - 1 : -1);
      r = roundDigits(digits, e10 + 1 - X, decPh.length);
      const maxInt = eng ? nInt : Math.max(nInt, 0);
      if (r[0].length > maxInt && (eng || r[0].length > Math.max(nInt, 1) || nInt === 0)) {
        X += eng ? nInt : 1;
        r = roundDigits(digits, e10 + 1 - X, decPh.length);
      }
    }
    [I, F] = r;
    const es = expStr(X, expPh.length || 1);
    expText = 'E' + (X < 0 ? '-' : (toks[expI] as { v: string }).v === '+' ? '+' : '') + es.slice(2);
  } else {
    [I, F] = fixed(a, decPh.length);
  }
  const bare = !intPh.length && !!I && phIdx.length > 0;
  if (bare) {
    // no integer placeholders: digits go right before the decimal point / first decimal placeholder
    const at = dot >= 0 ? dot : decPh[0] ?? toks.length;
    out[at] = I;
  }
  fillInt(I, intPh.map((i) => toks[i] as { v: string }), group).forEach((s, j) => { out[intPh[j]!] = s; });
  fillDec(F, decPh.map((i) => toks[i] as { v: string })).forEach((s, j) => { out[decPh[j]!] = s; });
  o.add(sign);
  toks.forEach((t, i) => {
    if (bare && i === (dot >= 0 ? dot : decPh[0] ?? toks.length)) o.add(out[i]!);
    if (t.k === 'ph') { if (!expPh.includes(i)) o.add(out[i]!); }
    else if (t.k === 'dot') o.add('.');
    else if (t.k === 'pct') o.add('%');
    else if (t.k === 'exp') o.add(expText);
    else if (t.k === 'comma') { if (firstPh < 0 || i < firstPh) o.add(','); }
    else if (t.k === 'slash') o.add('/');
    else if (t.k === 'gen') o.add(formatGeneral(a));
    else passThrough(t, o);
  });
  if (bare && dot < 0 && !decPh.length) o.add(I);
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const p2 = (n: number): string => (n < 10 ? '0' : '') + n;

const renderDate = (sec: Section, v: number, date1904: boolean, o: Out): boolean => {
  if (v < 0 || v >= (date1904 ? 2957004 : 2958466)) return false;
  const f = Math.pow(10, sec.sub);
  const units = Math.round(v * 86400 * f);
  const unitsPerDay = 86400 * f;
  const day = Math.floor(units / unitsPerDay);
  const rem = units - day * unitsPerDay;
  const secs = Math.floor(rem / f);
  const H = Math.floor(secs / 3600), Mi = Math.floor((secs % 3600) / 60), S = secs % 60;
  const subUnits = rem - secs * f;
  const { year, month, day: D } = fromSerial(day, date1904);
  const wd = weekdayOf(day, date1904);
  const totalSecs = Math.floor(units / f);
  for (const t of sec.toks) {
    if (t.k !== 'dt') { passThrough(t, o); continue; }
    const len = t.v.length;
    if (t.v === 'sub') { o.add('.' + String(subUnits).padStart(sec.sub, '0').slice(0, t.n ?? 0).padEnd(t.n ?? 0, '0')); continue; }
    switch (t.v[0]) {
      case 'y': o.add(len <= 2 ? p2(year % 100) : String(year)); break;
      case 'm':
        o.add(len === 1 ? String(month) : len === 2 ? p2(month) : len === 3 ? MONTHS[month - 1]!.slice(0, 3) : len === 5 ? MONTHS[month - 1]![0]! : MONTHS[month - 1]!);
        break;
      case 'M': o.add(len === 1 ? String(Mi) : p2(Mi)); break;
      case 'd': o.add(len === 1 ? String(D) : len === 2 ? p2(D) : len === 3 ? DAYS[wd]!.slice(0, 3) : DAYS[wd]!); break;
      case 'h': {
        const h = sec.ampm ? ((H + 11) % 12) + 1 : H;
        o.add(len === 1 ? String(h) : p2(h));
        break;
      }
      case 's': o.add(len === 1 ? String(S) : p2(S)); break;
      case '[': {
        const which = t.v[1];
        const n = which === 'h' ? Math.floor(totalSecs / 3600) : which === 'm' ? Math.floor(totalSecs / 60) : totalSecs;
        o.add(which === 'h' ? String(n) : p2(n));
        break;
      }
      default:
        if (t.v === 'ampm') o.add(H >= 12 ? 'PM' : 'AM');
        else if (t.v === 'ap' || t.v === 'AP') o.add(H >= 12 ? (t.n ? 'p' : 'P') : t.v === 'ap' ? 'a' : 'A');
    }
  }
  return true;
};

const renderText = (sec: Section, s: string, o: Out): void => {
  for (const t of sec.toks) {
    if (t.k === 'at') o.add(s);
    else if (!passThrough(t, o)) {
      if (t.k === 'ph') o.add(t.v);
      else if (t.k === 'dt') o.add(t.v);
    }
  }
};

const testCond = (c: Cond, v: number): boolean => {
  switch (c.op) {
    case '<': return v < c.v;
    case '>': return v > c.v;
    case '<=': return v <= c.v;
    case '>=': return v >= c.v;
    case '<>': return v !== c.v;
    default: return v === c.v;
  }
};

const GENERAL: Section = { toks: [{ k: 'gen' }], kind: 'general', ampm: false, sub: 0 };

/** Section for a number and whether its sign is implied by the section (so it's dropped). */
const pickSection = (secs: Section[], v: number): [Section | undefined, boolean] => {
  let ns = secs;
  if (secs.length === 4) ns = secs.slice(0, 3);
  else if (secs.length > 1 && secs[secs.length - 1]!.kind === 'text') ns = secs.slice(0, -1);
  else if (secs.length === 1 && secs[0]!.kind === 'text') return [GENERAL, false];
  if (ns.length === 1) return [ns[0], false];
  if (ns[0]!.cond || ns[1]!.cond) {
    // With conditions, an unconditioned second section of a two-section format takes everything else.
    const c0 = ns[0]!.cond ?? { op: '>', v: 0 };
    const c1 = ns[1]!.cond ?? (ns.length === 2 ? undefined : { op: '<', v: 0 });
    // A section whose condition only negative numbers can meet shows no minus sign.
    const negOnly = (c: Cond | undefined): boolean => !!c && (c.op === '<' || c.op === '<=') && c.v <= 0;
    if (testCond(c0, v)) return [ns[0], negOnly(c0)];
    if (!c1 || testCond(c1, v)) return [ns[1], negOnly(c1)];
    if (ns[2]) return [ns[2], false];
    return [undefined, false];
  }
  if (v > 0 || (v === 0 && ns.length === 2)) return [ns[0], false];
  if (v < 0) return [ns[1], true];
  return [ns[2] ?? ns[0], false];
};

const hashes = (w?: number): string => '#'.repeat(Math.max(1, w ?? 11));

export const formatValue = (v: Scalar, code: string, opts: { date1904?: boolean; width?: number } = {}): Formatted => {
  if (v === null) return { text: '' };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
  if (isError(v)) return { text: v.code };
  const secs = parseFormat(code || 'General');
  const o = new Out();
  if (typeof v === 'string') {
    const tsec = secs.length === 4 ? secs[3] : secs.find((s) => s.kind === 'text');
    if (!tsec) return { text: v, align: 'left' };
    renderText(tsec, v, o);
    return done(o, tsec.color, 'left');
  }
  if (!Number.isFinite(v)) return { text: '#NUM!' };
  const [sec, dropSign] = pickSection(secs, v);
  if (!sec) return { text: hashes(opts.width), align: 'right' };
  const n = dropSign ? Math.abs(v) : v;
  switch (sec.kind) {
    case 'general': {
      const gen = sec.toks.length === 1;
      if (gen) return done(Object.assign(o, { s: formatGeneral(n, opts.width ?? 11) }), sec.color, 'right');
      for (const t of sec.toks) {
        if (t.k === 'gen') o.add(formatGeneral(n));
        else if (!passThrough(t, o)) o.add(t.k === 'pct' ? '%' : t.k === 'dot' ? '.' : t.k === 'comma' ? ',' : t.k === 'ph' ? t.v : '');
      }
      break;
    }
    case 'date':
      if (!renderDate(sec, n, !!opts.date1904, o)) return { text: hashes(opts.width), align: 'right' };
      break;
    case 'text':
      renderText(sec, formatGeneral(n), o);
      break;
    default:
      renderNumber(sec, n, o);
  }
  return done(o, sec.color, 'right');
};

const done = (o: Out, color: string | undefined, align: 'left' | 'right'): Formatted => {
  const r: Formatted = { text: o.s, align };
  if (color) r.color = color;
  if (o.fill !== undefined) { r.fill = o.fill; r.fillAt = o.fillAt; }
  return r;
};

// ---------------------------------------------------------------- typed-entry formats

/** The format Spreadsheets give a cell when this text is typed into it (undefined: stays General). */
export const guessFormat = (input: string): string | undefined => {
  const s = input.trim();
  if (!s) return undefined;
  let m: RegExpExecArray | null;
  if ((m = /^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(\.\d*)?\s*%$/.exec(s)) && /\d/.test(s)) return m[1] && m[1].length > 1 ? '0.00%' : '0%';
  if ((m = /^\(?[+-]?\$\s*(?:\d{1,3}(?:,\d{3})+|\d+)?(\.\d*)?\)?$/.exec(s)) && /\d/.test(s)) return m[1] && m[1].length > 1 ? '$#,##0.00' : '$#,##0';
  if ((m = /^\(?[+-]?\d{1,3}(?:,\d{3})+(\.\d*)?\)?$/.exec(s))) return m[1] && m[1].length > 1 ? '#,##0.00' : '#,##0';
  if (/^[+-]?(\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(s)) return '0.00E+00';
  if (/^[+-]?\d+(\.\d*)?$/.test(s)) return undefined;
  if ((m = /^[+-]?\d+\s+(\d+)\/(\d+)$/.exec(s))) return +m[2]! > 0 && +m[1]! < +m[2]! ? (m[2]!.length > 1 ? '# ??/??' : '# ?/?') : undefined;
  const time = /^(\d{1,2}):(\d{1,2})(:\d{1,2}(?:\.\d+)?)?\s*([ap]m?)?$/i;
  if ((m = time.exec(s))) {
    if (+m[2]! > 59) return undefined;
    return m[4] ? (m[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : m[3] ? 'h:mm:ss' : 'h:mm';
  }
  const dt = /^(.+?)\s+(\d{1,2}:\d{1,2}(?::\d{1,2})?\s*(?:[ap]m?)?)$/i.exec(s);
  if (dt && guessDate(dt[1]!) && time.test(dt[2]!)) return 'm/d/yyyy h:mm';
  return guessDate(s);
};

const MON = '(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\\.?';
const validDate = (y: number, mo: number, d: number): boolean => mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(Date.UTC(y, mo, 0)).getUTCDate();

const guessDate = (s: string): string | undefined => {
  let m: RegExpExecArray | null;
  const yr = new Date().getFullYear();
  if ((m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s))) return validDate(+m[1]!, +m[2]!, +m[3]!) ? 'm/d/yyyy' : undefined;
  if ((m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2}|\d{4})$/.exec(s))) return validDate(m[3]!.length === 2 ? 2000 + +m[3]! : +m[3]!, +m[1]!, +m[2]!) ? 'm/d/yyyy' : undefined;
  if ((m = /^(\d{1,2})[-/](\d{1,2})$/.exec(s))) {
    if (validDate(yr, +m[1]!, +m[2]!)) return 'd-mmm';
    // "3/2024" style month/year
    return undefined;
  }
  if ((m = /^(\d{1,2})[-/](\d{4})$/.exec(s))) return +m[1]! >= 1 && +m[1]! <= 12 ? 'mmm-yy' : undefined;
  if (new RegExp(`^(\\d{1,2})[-\\s/](${MON})$`, 'i').test(s)) return 'd-mmm';
  if (new RegExp(`^(\\d{1,2})[-\\s/](${MON})[-\\s/,]+(\\d{2}|\\d{4})$`, 'i').test(s)) return 'd-mmm-yy';
  if (new RegExp(`^(${MON})[-\\s/](\\d{1,2})$`, 'i').test(s)) return 'd-mmm';
  if (new RegExp(`^(${MON})\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})$`, 'i').test(s)) return 'd-mmm-yy';
  if (new RegExp(`^(${MON})[-\\s/](\\d{2}|\\d{4})$`, 'i').test(s)) return 'mmm-yy';
  return undefined;
};
