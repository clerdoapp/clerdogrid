// OpenFormula (OpenDocument 1.3 Part 4) ↔ ClerdoGrid formula text.
// ODF writes references as [.A1], [$Sheet1.A1:.B2], separates arguments with ";", uses "~" for union,
// "!" for intersection, "|" between array rows, and prefixes newer functions with COM.MICROSOFT.
import { parseFormula, quoteSheet, type Node, type RefNode } from '../engine/parser';
import { stripFuturePrefixes } from './formula-xml';

const LEGACY = new Set(['CHIDIST', 'CHIINV', 'CHITEST', 'FDIST', 'FINV', 'NORMSDIST', 'NORMSINV', 'TDIST']);
const NEWER = new Set(`CONCAT TEXTJOIN IFS SWITCH MAXIFS MINIFS XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA
TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND MAP REDUCE SCAN
MAKEARRAY BYROW BYCOL ISOMITTED ARRAYTOTEXT VALUETOTEXT AGGREGATE ENCODEURL FILTERXML WEBSERVICE IMAGE GROUPBY PIVOTBY PERCENTOF
REGEXTEST REGEXEXTRACT REGEXREPLACE TRIMRANGE STOCKHISTORY CEILING FLOOR FORMULATEXT ISFORMULA SHEET SHEETS DAYS IFNA`.split(/\s+/));
const STANDARD_DOTTED = new Set(['ERROR.TYPE']);
// Names whose OpenFormula form is standard (no prefix), even though newer than the first edition.
const PLAIN = new Set(['IFNA', 'DAYS', 'SHEET', 'SHEETS', 'FORMULATEXT', 'ISFORMULA']);

export const odfFunctionName = (name: string): string => {
  const u = name.toUpperCase();
  if (LEGACY.has(u)) return `LEGACY.${u}`;
  if (PLAIN.has(u) || STANDARD_DOTTED.has(u)) return u;
  if (u.includes('.') || NEWER.has(u)) return `COM.MICROSOFT.${u}`;
  return u;
};

export const gridFunctionName = (name: string): string => {
  let u = name.toUpperCase();
  for (let i = 0; i < 3; i++) u = u.replace(/^(COM\.MICROSOFT\.|ORG\.OPENOFFICE\.|ORG\.LIBREOFFICE\.|_XLFN\._XLWS\.|_XLFN\.|_XLWS\.|LEGACY\.)/, '');
  if (u === 'CEILING.XCL') return 'CEILING';
  if (u === 'FLOOR.XCL') return 'FLOOR';
  return u;
};

const colLetters = (c: number): string => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const odfSheet = (s: string): string => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);

/** "[$Sheet1.$A$1:.B2]" style reference text (without brackets when `bare`). */
export const odfRef = (r: RefNode, defaultSheet?: string, bare = false): string => {
  const d = (abs: boolean) => (abs ? '$' : '');
  const sheet = r.sheet ?? defaultSheet;
  const p1 = sheet !== undefined ? `$${odfSheet(sheet)}.` : '.';
  const p2 = r.sheet2 !== undefined && r.sheet2 !== r.sheet ? `$${odfSheet(r.sheet2)}.` : bare && sheet !== undefined ? `$${odfSheet(sheet)}.` : '.';
  let s: string;
  if (r.kind === 'cols') s = `${p1}${d(r.c1Abs)}${colLetters(r.c1)}:${p2}${d(r.c2Abs)}${colLetters(r.c2)}`;
  else if (r.kind === 'rows') s = `${p1}${d(r.r1Abs)}${r.r1 + 1}:${p2}${d(r.r2Abs)}${r.r2 + 1}`;
  else {
    const a1 = `${d(r.c1Abs)}${colLetters(r.c1)}${d(r.r1Abs)}${r.r1 + 1}`;
    s = r.kind === 'cell' && !r.sheet2 ? p1 + a1 : `${p1}${a1}:${p2}${d(r.c2Abs)}${colLetters(r.c2)}${d(r.r2Abs)}${r.r2 + 1}`;
  }
  return bare ? s : `[${s}]`;
};

const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5, ' ': 8, ':': 9 };
const numText = (v: number): string => (Number.isInteger(v) && Math.abs(v) < 1e21 ? String(v) : String(v).replace('e', 'E'));

export interface OdfFormulaCtx {
  /** Expands a structured table reference, which ODF can't express. */
  table?: (n: Extract<Node, { type: 'table' }>) => RefNode | undefined;
}

/** ClerdoGrid formula (no "=") → ODF attribute value "of:=...". Throws on text that doesn't parse. */
export const toOdfFormula = (formula: string, ctx: OdfFormulaCtx = {}): string => `of:=${odfExpr(parseFormula(formula), ctx)}`;

export const odfExpr = (n: Node, ctx: OdfFormulaCtx = {}, parent = 0): string => {
  const go = (x: Node, p = 0) => odfExpr(x, ctx, p);
  switch (n.type) {
    case 'num': return numText(n.value);
    case 'str': return `"${n.value.replace(/"/g, '""')}"`;
    case 'bool': return n.value ? 'TRUE()' : 'FALSE()';
    case 'err': return n.code;
    case 'missing': return '';
    case 'array': return `{${n.rows.map((r) => r.map((x) => go(x)).join(';')).join('|')}}`;
    case 'ref': return n.ref.external ? '#REF!' : odfRef(n.ref);
    case 'name': return n.name;
    case 'table': { const r = ctx.table?.(n); return r ? odfRef(r) : '#REF!'; }
    case 'call': return `${odfFunctionName(n.name)}(${n.args.map((x) => go(x)).join(';')})`;
    case 'invoke': return `${go(n.callee, 10)}(${n.args.map((x) => go(x)).join(';')})`;
    case 'unary': return n.op === '%' ? `${go(n.arg, 7)}%` : n.op === '@' ? go(n.arg, parent) : `${n.op}${go(n.arg, 6)}`;
    case 'union': return `(${n.items.map((x) => go(x, 10)).join('~')})`;
    case 'spill': return go(n.arg, parent);
    case 'binary': {
      const p = PREC[n.op]!;
      const s = `${go(n.left, p)}${n.op === ' ' ? '!' : n.op}${go(n.right, p + 1)}`;
      return p < parent ? `(${s})` : s;
    }
  }
};

/** Splits a space-separated list of ODF addresses, keeping quoted sheet names whole. */
export const splitAddresses = (s: string): string[] => {
  const out: string[] = [];
  let buf = '', q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === "'") { if (q && s[i + 1] === "'") { buf += "''"; i++; continue; } q = !q; }
    if (/\s/.test(ch) && !q) { if (buf) out.push(buf); buf = ''; continue; }
    buf += ch;
  }
  if (buf) out.push(buf);
  return out;
};

/** One reference inside [...] → ClerdoGrid reference text, or "#REF!". */
export const gridRef = (inner: string): string => {
  const s = inner.trim();
  if (s.startsWith("'") && s.includes("'#")) return '#REF!'; // another document
  const parts: string[] = [];
  let buf = '', q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === "'") { if (q && s[i + 1] === "'") { buf += "''"; i++; continue; } q = !q; buf += ch; continue; }
    if (ch === ':' && !q) { parts.push(buf); buf = ''; continue; }
    buf += ch;
  }
  parts.push(buf);
  if (parts.length > 2 || /#REF!/i.test(s)) return '#REF!';
  const one = (p: string) => {
    let rest = p.trim();
    let sheet: string | undefined;
    if (rest.startsWith('$')) rest = rest.slice(1);
    if (rest.startsWith("'")) {
      let j = 1, name = '';
      for (; j < rest.length; j++) { if (rest[j] === "'") { if (rest[j + 1] === "'") { name += "'"; j++; continue; } break; } name += rest[j]; }
      sheet = name;
      rest = rest.slice(j + 1);
    } else {
      const dot = rest.lastIndexOf('.');
      if (dot > 0) { sheet = rest.slice(0, dot); rest = rest.slice(dot); }
    }
    if (!rest.startsWith('.')) return null;
    const m = /^\.(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(rest);
    if (!m || (!m[2] && !m[4])) return null;
    return { sheet, cAbs: !!m[1], col: m[2] ? m[2].toUpperCase() : undefined, rAbs: !!m[3] || (!m[2] && !!m[1]), row: m[4] };
  };
  const a1 = one(parts[0]!);
  const b1 = parts[1] !== undefined ? one(parts[1]) : null;
  if (!a1 || (parts[1] !== undefined && !b1)) return '#REF!';
  const d = (x: boolean) => (x ? '$' : '');
  const piece = (p: NonNullable<ReturnType<typeof one>>) => `${p.col ? d(p.cAbs) + p.col : ''}${p.row ? d(p.rAbs) + p.row : ''}`;
  let prefix = '';
  if (a1.sheet !== undefined) {
    const s2 = b1?.sheet !== undefined && b1.sheet !== a1.sheet ? b1.sheet : undefined;
    const names = s2 ? [a1.sheet, s2] : [a1.sheet];
    prefix = (names.every((x) => quoteSheet(x) === x) ? names.join(':') : `'${names.join(':').replace(/'/g, "''")}'`) + '!';
  }
  return b1 && (piece(b1) !== piece(a1) || !a1.col || !a1.row) ? `${prefix}${piece(a1)}:${piece(b1)}` : prefix + piece(a1);
};

/** ODF formula attribute → ClerdoGrid formula text (no "="), or undefined when it can't be read. */
export const fromOdfFormula = (raw: string): string | undefined => {
  let s = raw.trim();
  const ns = /^([A-Za-z][\w.-]*):/.exec(s);
  if (ns && !s.startsWith('[')) {
    const p = ns[1]!.toLowerCase();
    s = s.slice(ns[0].length);
    if (p === 'msoxl') { s = s.replace(/^=/, ''); return valid(s); }
    if (p !== 'of' && p !== 'oooc' && p !== 'ooow') return undefined;
  }
  s = s.replace(/^=/, '');
  let out = '';
  let arr = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === '"') {
      let j = i + 1;
      for (; j < s.length; j++) { if (s[j] === '"') { if (s[j + 1] === '"') { j++; continue; } break; } }
      out += s.slice(i, j + 1);
      i = j;
      continue;
    }
    if (ch === '[') {
      let j = i + 1, q = false;
      for (; j < s.length; j++) { if (s[j] === "'") q = !q; else if (s[j] === ']' && !q) break; }
      out += gridRef(s.slice(i + 1, j));
      i = j;
      continue;
    }
    if (ch === '{') { arr++; out += ch; continue; }
    if (ch === '}') { arr = Math.max(0, arr - 1); out += ch; continue; }
    if (arr) { out += ch === ';' ? ',' : ch === '|' ? ';' : ch; continue; }
    if (ch === ';' || ch === '~') { out += ','; continue; }
    if (ch === '!') { out += ' '; continue; }
    const id = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(s.slice(i));
    if (id) {
      const after = s.slice(i + id[0].length);
      const call = /^\s*\(/.test(after);
      if (call) {
        const name = gridFunctionName(id[0]);
        if ((name === 'TRUE' || name === 'FALSE') && /^\s*\(\s*\)/.test(after)) {
          out += name;
          i += id[0].length + /^\s*\(\s*\)/.exec(after)![0].length - 1;
          continue;
        }
        out += name;
      } else out += id[0];
      i += id[0].length - 1;
      continue;
    }
    out += ch;
  }
  return valid(stripFuturePrefixes(out));
};

const valid = (f: string): string | undefined => {
  try { parseFormula(f); return f; } catch { return undefined; }
};
