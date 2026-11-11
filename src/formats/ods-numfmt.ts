// OpenDocument data styles (number:*-style, ODF 1.3 Part 3 §16.29) ↔ number format codes.
import { a, child, elems, odfText, tag, type ONode } from './ods-xml';
import { isDateFormat } from './numfmt';

const COLOR_NAMES: Record<string, string> = { '#000000': 'Black', '#0000FF': 'Blue', '#00FFFF': 'Cyan', '#00FF00': 'Green', '#FF00FF': 'Magenta', '#FF0000': 'Red', '#FFFFFF': 'White', '#FFFF00': 'Yellow' };
const NAMED: Record<string, string> = Object.fromEntries(Object.entries(COLOR_NAMES).map(([k, v]) => [v.toLowerCase(), k]));

// ---------- Reading ----------

const quoteLit = (s: string): string => {
  if (!s) return '';
  if (/^[ $\-+/():!^&'~{}<>=,.%]+$/.test(s)) return s;
  return `"${s.replace(/"/g, '')}"`;
};

const numberPart = (n: ONode): string => {
  const dec = +(a(n, 'number:decimal-places') ?? 0) || 0;
  const minDec = a(n, 'number:min-decimal-places') !== undefined ? +(a(n, 'number:min-decimal-places') ?? 0) : dec;
  const minInt = +(a(n, 'number:min-integer-digits') ?? 0) || 0;
  const grouping = a(n, 'number:grouping') === 'true';
  let int = '0'.repeat(minInt) || '#';
  if (grouping) { int = int.padStart(4, '#'); int = `${int.slice(0, -3)},${int.slice(-3)}`; }
  let out = int;
  if (dec > 0) out += '.' + '0'.repeat(Math.min(minDec, dec)) + '#'.repeat(Math.max(0, dec - minDec));
  const factor = +(a(n, 'number:display-factor') ?? 1) || 1;
  if (factor >= 1000) out += ','.repeat(Math.round(Math.log10(factor) / 3));
  return out;
};

const sectionCode = (style: ONode): string => {
  const kind = style.name;
  const truncate = a(style, 'number:truncate-on-overflow') === 'false';
  let firstTime = true;
  let out = '';
  let color = '';
  for (const k of elems(style)) {
    const long = a(k, 'number:style') === 'long';
    switch (k.name) {
      case 'number:number': out += numberPart(k); break;
      case 'number:scientific-number': {
        const dec = +(a(k, 'number:decimal-places') ?? 0) || 0;
        out += '0'.repeat(Math.max(1, +(a(k, 'number:min-integer-digits') ?? 1) || 1)) + (dec ? '.' + '0'.repeat(dec) : '') + 'E+' + '0'.repeat(Math.max(1, +(a(k, 'number:min-exponent-digits') ?? 2) || 2));
        break;
      }
      case 'number:fraction': {
        const mi = +(a(k, 'number:min-integer-digits') ?? 0) || 0;
        const num = Math.max(1, +(a(k, 'number:min-numerator-digits') ?? 1) || 1);
        const fixed = a(k, 'number:denominator-value');
        const den = Math.max(1, +(a(k, 'number:min-denominator-digits') ?? 1) || 1);
        out += (mi ? '#'.repeat(Math.max(0, mi - 1)) + '0 ' : '# ') + '?'.repeat(num) + '/' + (fixed ?? '?'.repeat(den));
        break;
      }
      case 'number:text': {
        const t = odfText(k);
        out += t.split('%').map(quoteLit).join('%');
        break;
      }
      case 'number:currency-symbol': { const t = odfText(k); out += t === '$' ? '$' : `[$${t}]`; break; }
      case 'number:text-content': out += '@'; break;
      case 'number:day': out += long ? 'dd' : 'd'; break;
      case 'number:month': out += a(k, 'number:textual') === 'true' ? (long ? 'mmmm' : 'mmm') : long ? 'mm' : 'm'; break;
      case 'number:year': out += long ? 'yyyy' : 'yy'; break;
      case 'number:day-of-week': out += long ? 'dddd' : 'ddd'; break;
      case 'number:era': out += long ? 'gggg' : 'g'; break;
      case 'number:quarter': out += long ? '"Q"q' : 'q'; break;
      case 'number:week-of-year': out += 'ww'; break;
      case 'number:hours': {
        const h = long ? 'hh' : 'h';
        out += truncate && firstTime && kind === 'number:time-style' ? `[${h}]` : h;
        firstTime = false;
        break;
      }
      case 'number:minutes': out += long ? 'mm' : 'mm'.slice(0, long ? 2 : 1); break;
      case 'number:seconds': {
        const dec = +(a(k, 'number:decimal-places') ?? 0) || 0;
        out += (long ? 'ss' : 's') + (dec ? '.' + '0'.repeat(dec) : '');
        break;
      }
      case 'number:am-pm': out += 'AM/PM'; break;
      case 'number:boolean': out += 'TRUE'; break;
      case 'style:text-properties': {
        const c = (a(k, 'fo:color') ?? '').toUpperCase();
        if (COLOR_NAMES[c]) color = `[${COLOR_NAMES[c]}]`;
        break;
      }
      default: break;
    }
  }
  if (kind === 'number:boolean-style') return '"TRUE";"TRUE";"FALSE"';
  return color + out;
};

/** Format code for a data style; mapped sub-styles become extra sections. */
export const readDataStyle = (style: ONode, byName: Map<string, ONode>, depth = 0): string | undefined => {
  const own = a(style, 'clerdo:format-code');
  if (own) return own;
  const main = sectionCode(style);
  const maps = elems(style).filter((k) => k.name === 'style:map');
  if (!maps.length || depth > 3) return main || undefined;
  const parts: { cond: string; code: string }[] = [];
  for (const m of maps) {
    const sub = byName.get(a(m, 'style:apply-style-name') ?? '');
    const cond = (a(m, 'style:condition') ?? '').replace(/^value\(\)\s*/, '').replace(/\s+/g, '');
    if (!sub) continue;
    parts.push({ cond, code: readDataStyle(sub, byName, depth + 1) ?? 'General' });
  }
  const pos = parts.find((p) => p.cond === '>=0' || p.cond === '>0');
  const neg = parts.find((p) => p.cond === '<0' || p.cond === '<=0');
  const zero = parts.find((p) => p.cond === '=0');
  if (pos && !neg && parts.length === 1) return `${pos.code};${main}`;
  if (pos && neg && !zero) return `${pos.code};${neg.code};${main}`;
  if (pos && neg && zero) return `${pos.code};${neg.code};${zero.code};${main}`;
  return [...parts.map((p) => `[${p.cond}]${p.code}`), main].join(';');
};

// ---------- Writing ----------

const splitSections = (code: string): string[] => {
  const out: string[] = [];
  let buf = '', q = false, br = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    if (ch === '\\' && !q) { buf += ch + (code[i + 1] ?? ''); i++; continue; }
    if (ch === '"') q = !q;
    else if (!q && ch === '[') br = true;
    else if (!q && ch === ']') br = false;
    if (ch === ';' && !q && !br) { out.push(buf); buf = ''; continue; }
    buf += ch;
  }
  out.push(buf);
  return out;
};

interface Tok { k: 'lit' | 'num' | 'date' | 'text' | 'percent' | 'color' | 'cond' | 'elapsed' | 'currency' | 'ampm'; v: string }

const tokenize = (sec: string): Tok[] => {
  const toks: Tok[] = [];
  let i = 0;
  const lit = (s: string) => { const last = toks.at(-1); if (last?.k === 'lit') last.v += s; else toks.push({ k: 'lit', v: s }); };
  while (i < sec.length) {
    const ch = sec[i]!;
    const rest = sec.slice(i);
    let m: RegExpExecArray | null;
    if (ch === '"') { const j = sec.indexOf('"', i + 1); lit(sec.slice(i + 1, j < 0 ? sec.length : j)); i = j < 0 ? sec.length : j + 1; continue; }
    if (ch === '\\') { lit(sec[i + 1] ?? ''); i += 2; continue; }
    if (ch === '_') { lit(' '); i += 2; continue; }
    if (ch === '*') { i += 2; continue; }
    if (ch === '[') {
      const j = sec.indexOf(']', i);
      const body = sec.slice(i + 1, j < 0 ? sec.length : j);
      i = j < 0 ? sec.length : j + 1;
      if (/^(h+|m+|s+)$/i.test(body)) toks.push({ k: 'elapsed', v: body.toLowerCase() });
      else if (/^\$/.test(body)) { const sym = body.slice(1).split('-')[0]!; if (sym) toks.push({ k: 'currency', v: sym }); }
      else if (/^(<=|>=|<>|<|>|=)/.test(body)) toks.push({ k: 'cond', v: body });
      else toks.push({ k: 'color', v: body });
      continue;
    }
    if ((m = /^(AM\/PM|A\/P)/i.exec(rest))) { toks.push({ k: 'ampm', v: m[1]! }); i += m[1]!.length; continue; }
    if ((m = /^(y+|d+|h+|s+(?:\.0+)?|e+|g+|m+)/i.exec(rest))) { toks.push({ k: 'date', v: m[1]!.toLowerCase() }); i += m[1]!.length; continue; }
    if ((m = /^[0#?,.]*[0#?][0#?,.]*(?:[Ee][+-][0#]+)?(?:\s*[?#0]+\/[?#0-9]+)?/.exec(rest)) && m[0]) {
      toks.push({ k: 'num', v: m[0] });
      i += m[0].length;
      continue;
    }
    if (ch === '%') { toks.push({ k: 'percent', v: '%' }); i++; continue; }
    if (ch === '@') { toks.push({ k: 'text', v: '@' }); i++; continue; }
    if (/^General/i.test(rest)) { toks.push({ k: 'num', v: 'General' }); i += 7; continue; }
    if (/[$€£¥]/.test(ch)) { toks.push({ k: 'currency', v: ch }); i++; continue; }
    lit(ch);
    i++;
  }
  return toks;
};

const numberXml = (v: string): string => {
  if (v === 'General') return tag('number:number', { 'number:min-integer-digits': 1 });
  const frac = /^([#0,]*)\s*([?#0]+)\/([?#0-9]+)$/.exec(v);
  if (frac) {
    const intDigits = (frac[1]!.match(/0/g) ?? []).length;
    const den = frac[3]!;
    return tag('number:fraction', {
      'number:min-integer-digits': frac[1] ? Math.max(intDigits, 0) : undefined,
      'number:min-numerator-digits': frac[2]!.length,
      ...(/^\d+$/.test(den) ? { 'number:denominator-value': den } : { 'number:min-denominator-digits': den.length }),
    });
  }
  const sci = /^([#0,]*)(?:\.([0#]*))?[Ee][+-]([0#]+)$/.exec(v);
  if (sci) {
    return tag('number:scientific-number', { 'number:decimal-places': sci[2]?.length ?? 0, 'number:min-integer-digits': (sci[1]!.match(/0/g) ?? []).length || 1, 'number:min-exponent-digits': sci[3]!.length });
  }
  const trailing = /,+$/.exec(v)?.[0].length ?? 0;
  const body = trailing ? v.slice(0, -trailing) : v;
  const [int = '', dec = ''] = body.split('.');
  return tag('number:number', {
    'number:decimal-places': dec.length,
    'number:min-decimal-places': (dec.match(/0/g) ?? []).length,
    'number:min-integer-digits': (int.match(/0/g) ?? []).length,
    'number:grouping': int.includes(',') ? true : undefined,
    'number:display-factor': trailing ? 1000 ** trailing : undefined,
  });
};

const dateXml = (v: string, afterHour: boolean, isTime: boolean): string => {
  const ch = v[0];
  const long = v.length >= 2;
  if (ch === 'y') return tag('number:year', { 'number:style': v.length > 2 ? 'long' : undefined });
  if (ch === 'd') return v.length >= 3 ? tag('number:day-of-week', { 'number:style': v.length > 3 ? 'long' : undefined }) : tag('number:day', { 'number:style': long ? 'long' : undefined });
  if (ch === 'h') return tag('number:hours', { 'number:style': long ? 'long' : undefined });
  if (ch === 's') { const dec = v.split('.')[1]?.length ?? 0; return tag('number:seconds', { 'number:style': v.startsWith('ss') ? 'long' : undefined, 'number:decimal-places': dec || undefined }); }
  if (ch === 'm') {
    if (afterHour || (isTime && v.length <= 2)) return tag('number:minutes', { 'number:style': long ? 'long' : undefined });
    return tag('number:month', { 'number:textual': v.length >= 3 ? true : undefined, 'number:style': v.length === 2 || v.length >= 4 ? 'long' : undefined });
  }
  if (ch === 'e') return tag('number:year', { 'number:style': 'long' });
  if (ch === 'g') return tag('number:era', { 'number:style': v.length > 2 ? 'long' : undefined });
  return '';
};

const sectionXml = (name: string, sec: string, code: string, maps: string): string => {
  const toks = tokenize(sec);
  const has = (k: Tok['k']) => toks.some((t) => t.k === k);
  const isDate = toks.some((t) => t.k === 'date' && /^[yde]/.test(t.v)) || (toks.some((t) => t.k === 'date' && t.v.startsWith('m')) && !toks.some((t) => t.k === 'date' && /^[hs]/.test(t.v)));
  const isTime = !isDate && (has('elapsed') || toks.some((t) => t.k === 'date' && /^[hs]/.test(t.v)));
  const kind = has('text') && !has('num') ? 'number:text-style' : isDate ? 'number:date-style' : isTime ? 'number:time-style' : has('percent') ? 'number:percentage-style' : has('currency') ? 'number:currency-style' : 'number:number-style';
  let inner = '';
  let color = '';
  let lastWasHour = false;
  let elapsed = false;
  toks.forEach((t, idx) => {
    switch (t.k) {
      case 'lit': inner += tag('number:text', {}, escLit(t.v)); lastWasHour = lastWasHour && /^[:\s]*$/.test(t.v); break;
      case 'num': inner += numberXml(t.v); break;
      case 'percent': inner += tag('number:text', {}, '%'); break;
      case 'text': inner += '<number:text-content/>'; break;
      case 'currency': inner += tag('number:currency-symbol', {}, escLit(t.v)); break;
      case 'ampm': inner += '<number:am-pm/>'; break;
      case 'color': { const c = NAMED[t.v.toLowerCase()]; if (c) color = c; break; }
      case 'elapsed': {
        elapsed = true;
        const ch = t.v[0];
        inner += ch === 'h' ? tag('number:hours', { 'number:style': t.v.length > 1 ? 'long' : undefined }) : ch === 'm' ? tag('number:minutes', { 'number:style': t.v.length > 1 ? 'long' : undefined }) : tag('number:seconds', { 'number:style': t.v.length > 1 ? 'long' : undefined });
        lastWasHour = ch === 'h';
        break;
      }
      case 'date': {
        const nextIsSecond = toks.slice(idx + 1).find((x) => x.k === 'date')?.v.startsWith('s');
        inner += dateXml(t.v, t.v.startsWith('m') && (lastWasHour || !!nextIsSecond), isTime);
        lastWasHour = t.v.startsWith('h');
        break;
      }
      default: break;
    }
  });
  const props = color ? tag('style:text-properties', { 'fo:color': color }) : '';
  return tag(kind, {
    'style:name': name, 'clerdo:format-code': code || undefined,
    'number:truncate-on-overflow': elapsed && kind === 'number:time-style' ? false : undefined,
  }, props + inner + maps);
};

const escLit = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Data style elements for a format code: sub-styles for extra sections plus the main style named `name`. */
export const dataStyleXml = (name: string, code: string): string => {
  const secs = splitSections(code);
  if (secs.length === 1) return sectionXml(name, secs[0]!, code, '');
  const condOf = (s: string) => /^\[(<=|>=|<>|<|>|=)([^\]]*)\]/.exec(s);
  let subs = '';
  let maps = '';
  const explicit = secs.some((s) => condOf(s));
  const conds = explicit ? secs.slice(0, -1).map((s) => { const m = condOf(s); return m ? `${m[1]}${m[2]}` : '>=0'; })
    : secs.length === 2 ? ['>=0'] : secs.length === 3 ? ['>0', '<0'] : ['>0', '<0', '=0'];
  const mainSec = explicit ? secs.at(-1)! : secs.length === 2 ? secs[1]! : secs.length === 3 ? secs[2]! : secs[3]!;
  conds.forEach((c, i) => {
    const sub = `${name}P${i}`;
    subs += sectionXml(sub, secs[i]!.replace(/^\[(<=|>=|<>|<|>|=)[^\]]*\]/, ''), '', '');
    maps += tag('style:map', { 'style:condition': `value()${c}`, 'style:apply-style-name': sub });
  });
  return subs + sectionXml(name, mainSec.replace(/^\[(<=|>=|<>|<|>|=)[^\]]*\]/, ''), code, maps);
};

/** Value type ODF uses for a number shown with this format. */
export const valueTypeOf = (code: string | undefined): 'float' | 'percentage' | 'currency' | 'date' | 'time' => {
  if (!code || code === 'General') return 'float';
  const first = splitSections(code)[0]!;
  const stripped = first.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, (m) => (/^\[\$/.test(m) ? '$' : ''));
  if (isDateFormat(code)) return /[yd]/i.test(stripped) || /m{3,}/i.test(stripped) || (/m/i.test(stripped) && !/[hs]/i.test(stripped)) ? 'date' : 'time';
  if (stripped.includes('%')) return 'percentage';
  if (/[$€£¥]/.test(stripped) || /\[\$[^\]]+\]/.test(first)) return 'currency';
  return 'float';
};

export const currencyCode = (code: string): string => {
  const m = /\[\$([^\]-]+)/.exec(code)?.[1] ?? /[$€£¥]/.exec(code)?.[0] ?? '$';
  return ({ $: 'USD', '€': 'EUR', '£': 'GBP', '¥': 'JPY' } as Record<string, string>)[m] ?? (/^[A-Z]{3}$/.test(m) ? m : 'USD');
};

export { child };
