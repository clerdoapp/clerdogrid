import { ERROR_CODES, MAX_COLS, MAX_ROWS, type ErrorCode } from './types';

export type Node =
  | { type: 'num'; value: number }
  | { type: 'str'; value: string }
  | { type: 'bool'; value: boolean }
  | { type: 'err'; code: ErrorCode }
  | { type: 'missing' }
  | { type: 'array'; rows: Node[][] }
  | { type: 'ref'; ref: RefNode }
  | { type: 'name'; name: string; sheet?: string }
  | { type: 'table'; table: string; spec: TableSpec }
  | { type: 'call'; name: string; args: Node[] }
  | { type: 'invoke'; callee: Node; args: Node[] }
  | { type: 'unary'; op: '-' | '+' | '%' | '@'; arg: Node }
  | { type: 'binary'; op: BinOp; left: Node; right: Node }
  | { type: 'union'; items: Node[] }
  | { type: 'spill'; arg: Node };

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=' | ':' | ' ';

/** A cell, area, whole-column or whole-row reference. Rows/cols are 0-based. */
export interface RefNode {
  sheet?: string;
  /** Last sheet of a 3D reference (Sheet1:Sheet3!A1). */
  sheet2?: string;
  r1: number; c1: number; r2: number; c2: number;
  r1Abs: boolean; c1Abs: boolean; r2Abs: boolean; c2Abs: boolean;
  kind: 'cell' | 'area' | 'cols' | 'rows';
  /** External workbook reference: always #REF! here. */
  external?: boolean;
}

export interface TableSpec {
  /** #All, #Data, #Headers, #Totals, #This Row (@). Data when absent. */
  items: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[];
  col1?: string;
  col2?: string;
}

export class FormulaError extends Error {}

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrorCode }
  | { t: 'ref'; v: RefNode }
  | { t: 'name'; v: string; sheet?: string }
  | { t: 'table'; table: string; spec: TableSpec }
  | { t: 'func'; v: string }
  | { t: 'op'; v: string }
  | { t: '(' | ')' | ',' | ';' | '{' | '}' | '#' | 'sp' };

const COL = '\\$?[A-Za-z]{1,3}';
const ROW = '\\$?[0-9]{1,7}';
const CELL_RE = new RegExp(`^(${COL})(${ROW})(?![A-Za-z0-9_.(\\[])`);
const AREA_RE = new RegExp(`^(${COL})(${ROW}):(${COL})(${ROW})(?![A-Za-z0-9_.(\\[])`);
const COLS_RE = new RegExp(`^(${COL}):(${COL})(?![A-Za-z0-9_.(\\[])`);
const ROWS_RE = new RegExp(`^(${ROW}):(${ROW})(?![A-Za-z0-9_.(\\[])`);
const NAME_RE = /^[A-Za-z_\\À-￿][A-Za-z0-9_.\\?À-￿]*/;
const SHEET_PLAIN = /^([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*(?::[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)?)!/;

const colOf = (s: string): number => {
  let n = 0;
  for (const ch of s.replace('$', '').toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
};

const cellPart = (c: string, r: string) => ({ col: colOf(c), row: +r.replace('$', '') - 1, cAbs: c.startsWith('$'), rAbs: r.startsWith('$') });

const validRef = (n: RefNode): boolean => n.r1 >= 0 && n.c1 >= 0 && n.r2 < MAX_ROWS && n.c2 < MAX_COLS && n.r1 <= MAX_ROWS && n.c1 < MAX_COLS;

/** Tries to read a reference (cell, area, columns, rows) at the start of `s`. */
const readRef = (s: string): { len: number; ref: RefNode } | null => {
  let m = AREA_RE.exec(s);
  if (m) {
    const a = cellPart(m[1]!, m[2]!), b = cellPart(m[3]!, m[4]!);
    const ref: RefNode = {
      r1: Math.min(a.row, b.row), r2: Math.max(a.row, b.row), c1: Math.min(a.col, b.col), c2: Math.max(a.col, b.col),
      r1Abs: a.row <= b.row ? a.rAbs : b.rAbs, r2Abs: a.row <= b.row ? b.rAbs : a.rAbs,
      c1Abs: a.col <= b.col ? a.cAbs : b.cAbs, c2Abs: a.col <= b.col ? b.cAbs : a.cAbs, kind: 'area',
    };
    return validRef(ref) ? { len: m[0].length, ref } : null;
  }
  m = CELL_RE.exec(s);
  if (m) {
    const a = cellPart(m[1]!, m[2]!);
    const ref: RefNode = { r1: a.row, r2: a.row, c1: a.col, c2: a.col, r1Abs: a.rAbs, r2Abs: a.rAbs, c1Abs: a.cAbs, c2Abs: a.cAbs, kind: 'cell' };
    return validRef(ref) ? { len: m[0].length, ref } : null;
  }
  m = COLS_RE.exec(s);
  if (m) {
    const a = colOf(m[1]!), b = colOf(m[2]!);
    const ref: RefNode = { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a, b), c2: Math.max(a, b), r1Abs: true, r2Abs: true, c1Abs: m[1]!.startsWith('$'), c2Abs: m[2]!.startsWith('$'), kind: 'cols' };
    return validRef(ref) ? { len: m[0].length, ref } : null;
  }
  m = ROWS_RE.exec(s);
  if (m) {
    const a = +m[1]!.replace('$', '') - 1, b = +m[2]!.replace('$', '') - 1;
    const ref: RefNode = { r1: Math.min(a, b), r2: Math.max(a, b), c1: 0, c2: MAX_COLS - 1, r1Abs: m[1]!.startsWith('$'), r2Abs: m[2]!.startsWith('$'), c1Abs: true, c2Abs: true, kind: 'rows' };
    return validRef(ref) ? { len: m[0].length, ref } : null;
  }
  return null;
};

const readTableSpec = (s: string, at: number): { end: number; spec: TableSpec } => {
  // s[at] === '['
  let depth = 0, i = at;
  for (; i < s.length; i++) {
    const ch = s[i];
    if (ch === "'") { i++; continue; }
    if (ch === '[') depth++;
    else if (ch === ']') { depth--; if (depth === 0) break; }
  }
  if (depth !== 0) throw new FormulaError('Missing ] in a table reference.');
  const inner = s.slice(at + 1, i);
  const spec: TableSpec = { items: [] };
  const unq = (x: string) => x.replace(/'(.)/g, '$1').trim();
  if (!inner.includes('[')) {
    if (inner.startsWith('@')) { spec.items.push('#This Row'); if (inner.length > 1) spec.col1 = unq(inner.slice(1)); }
    else if (inner.startsWith('#')) spec.items.push(normItem(inner));
    else if (inner) spec.col1 = unq(inner);
    return { end: i + 1, spec };
  }
  const parts: string[] = [];
  const re = /\[((?:[^\]']|'.)*)\]|(@)|(:)/g;
  let m: RegExpExecArray | null;
  let range = false;
  while ((m = re.exec(inner))) {
    if (m[2]) spec.items.push('#This Row');
    else if (m[3]) range = true;
    else {
      const p = unq(m[1]!);
      if (p.startsWith('#')) spec.items.push(normItem(p));
      else parts.push(p);
    }
  }
  if (inner.trim().startsWith('@') && !spec.items.includes('#This Row')) spec.items.push('#This Row');
  spec.col1 = parts[0];
  if (range || parts.length > 1) spec.col2 = parts[1];
  return { end: i + 1, spec };
};

const normItem = (x: string): TableSpec['items'][number] => {
  const k = x.toLowerCase().replace(/\s+/g, ' ');
  if (k === '#all') return '#All';
  if (k === '#headers') return '#Headers';
  if (k === '#totals') return '#Totals';
  if (k === '#this row') return '#This Row';
  return '#Data';
};

export const tokenize = (src: string): Tok[] => {
  const out: Tok[] = [];
  let i = 0;
  const s = src;
  while (i < s.length) {
    const ch = s[i]!;
    const rest = s.slice(i);
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      while (i < s.length && /\s/.test(s[i]!)) i++;
      out.push({ t: 'sp' });
      continue;
    }
    if (ch === '"') {
      let v = '';
      i++;
      for (;;) {
        if (i >= s.length) throw new FormulaError('A text value is missing its closing quote.');
        if (s[i] === '"') { if (s[i + 1] === '"') { v += '"'; i += 2; continue; } i++; break; }
        v += s[i++];
      }
      out.push({ t: 'str', v });
      continue;
    }
    if (ch === '#') {
      const code = ERROR_CODES.find((c) => rest.toUpperCase().startsWith(c));
      if (code) { out.push({ t: 'err', v: code }); i += code.length; continue; }
      out.push({ t: '#' }); i++; continue;
    }
    if (/[0-9.]/.test(ch) && !readRef(rest)) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(rest);
      if (m) {
        // A row range like 1:3 is a reference, not a number.
        const rr = ROWS_RE.exec(rest);
        if (!rr || rr[0].length <= m[0].length) { out.push({ t: 'num', v: Number(m[0]) }); i += m[0].length; continue; }
      }
    }
    // Sheet-qualified references: 'My Sheet'!A1, Sheet1!A1:B2, Sheet1:Sheet3!A1, [1]Sheet1!A1
    let sheet: string | undefined, sheet2: string | undefined, external = false, plen = 0;
    if (ch === "'") {
      let j = i + 1, name = '';
      for (;;) {
        if (j >= s.length) throw new FormulaError('A sheet name is missing its closing quote.');
        if (s[j] === "'") { if (s[j + 1] === "'") { name += "'"; j += 2; continue; } j++; break; }
        name += s[j++];
      }
      if (s[j] !== '!') throw new FormulaError('Expected ! after a sheet name.');
      const ext = /^\[([^\]]+)\](.*)$/.exec(name);
      if (ext) { external = true; name = ext[2]!; }
      [sheet, sheet2] = name.split(':') as [string, string | undefined];
      plen = j + 1 - i;
    } else {
      const ext = /^\[[^\]]+\]/.exec(rest);
      const r2 = ext ? rest.slice(ext[0].length) : rest;
      const m = SHEET_PLAIN.exec(r2);
      if (m && (readRef(r2.slice(m[0].length)) || /^#REF!/i.test(r2.slice(m[0].length)) || NAME_RE.test(r2.slice(m[0].length)))) {
        [sheet, sheet2] = m[1]!.split(':') as [string, string | undefined];
        plen = m[0].length + (ext ? ext[0].length : 0);
        external = !!ext;
      }
    }
    if (plen) {
      const after = s.slice(i + plen);
      const r = readRef(after);
      if (r) { out.push({ t: 'ref', v: { ...r.ref, sheet, sheet2, external } }); i += plen + r.len; continue; }
      if (/^#REF!/i.test(after)) { out.push({ t: 'err', v: '#REF!' }); i += plen + 5; continue; }
      const n = NAME_RE.exec(after);
      if (n) { out.push({ t: 'name', v: n[0], sheet }); i += plen + n[0].length; continue; }
    }
    const r = readRef(rest);
    if (r) { out.push({ t: 'ref', v: r.ref }); i += r.len; continue; }
    if (ch === '[') {
      const { end, spec } = readTableSpec(s, i);
      out.push({ t: 'table', table: '', spec }); i = end; continue;
    }
    const n = NAME_RE.exec(rest);
    if (n) {
      let name = n[0];
      i += name.length;
      if (s[i] === '[') {
        const { end, spec } = readTableSpec(s, i);
        out.push({ t: 'table', table: name, spec }); i = end; continue;
      }
      const up = name.toUpperCase();
      // Peek past spaces for a call.
      let j = i;
      while (s[j] === ' ') j++;
      if (s[j] === '(') { name = up.replace(/^(_XLFN\.|_XLWS\.)+/, ''); out.push({ t: 'func', v: name }); i = j; continue; }
      if (up === 'TRUE' || up === 'FALSE') { out.push({ t: 'bool', v: up === 'TRUE' }); continue; }
      out.push({ t: 'name', v: name.replace(/^_xlpm\./i, '') });
      continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { out.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%:@'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
    if ('(),;{}'.includes(ch)) { out.push({ t: ch as '(' }); i++; continue; }
    throw new FormulaError(`Unexpected character "${ch}".`);
  }
  return out;
};

const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

class Parser {
  private i = 0;
  constructor(private readonly toks: Tok[]) {}

  parse(): Node {
    this.skip();
    const n = this.expr(0);
    this.skip();
    if (this.i < this.toks.length) throw new FormulaError('There is something extra at the end of the formula.');
    return n;
  }

  private peek(): Tok | undefined { return this.toks[this.i]; }
  private skip(): void { while (this.peek()?.t === 'sp') this.i++; }

  private expr(min: number): Node {
    let left = this.unary();
    for (;;) {
      const save = this.i;
      this.skip();
      const t = this.peek();
      if (!t || t.t !== 'op') { this.i = save; break; }
      const p = PREC[t.v];
      if (p === undefined || p < min) { this.i = save; break; }
      this.i++;
      this.skip();
      const right = this.expr(t.v === '^' ? p + 1 : p + 1);
      left = { type: 'binary', op: t.v as BinOp, left, right };
    }
    return left;
  }

  private unary(): Node {
    this.skip();
    const t = this.peek();
    if (t?.t === 'op' && (t.v === '-' || t.v === '+' || t.v === '@')) {
      this.i++;
      const arg = this.unary();
      return { type: 'unary', op: t.v, arg };
    }
    return this.postfix();
  }

  private postfix(): Node {
    let n = this.rangeExpr();
    for (;;) {
      const t = this.peek();
      if (t?.t === 'op' && t.v === '%') { this.i++; n = { type: 'unary', op: '%', arg: n }; continue; }
      break;
    }
    return n;
  }

  // Reference operators bind tightest: ':' then intersection (space).
  private rangeExpr(): Node {
    let n = this.colonExpr();
    for (;;) {
      const save = this.i;
      if (this.peek()?.t !== 'sp') break;
      this.skip();
      const t = this.peek();
      if (t && (t.t === 'ref' || t.t === 'name' || t.t === '(' || t.t === 'func' || t.t === 'table') && isRefLike(n)) {
        const right = this.colonExpr();
        n = { type: 'binary', op: ' ', left: n, right };
        continue;
      }
      this.i = save;
      break;
    }
    return n;
  }

  private colonExpr(): Node {
    let n = this.primary();
    for (;;) {
      const t = this.peek();
      if (t?.t === 'op' && t.v === ':') { this.i++; n = { type: 'binary', op: ':', left: n, right: this.primary() }; continue; }
      if (t?.t === '#') { this.i++; n = { type: 'spill', arg: n }; continue; }
      if (t?.t === '(' && (n.type === 'call' || n.type === 'invoke' || n.type === 'name')) {
        // LAMBDA(x, x+1)(5), or a name holding a LAMBDA: MyFn(3)
        if (n.type === 'name') { this.i++; n = { type: 'call', name: n.name.toUpperCase(), args: this.args() }; continue; }
        this.i++;
        n = { type: 'invoke', callee: n, args: this.args() };
        continue;
      }
      break;
    }
    return n;
  }

  private args(): Node[] {
    const args: Node[] = [];
    this.skip();
    if (this.peek()?.t === ')') { this.i++; return args; }
    for (;;) {
      this.skip();
      const t = this.peek();
      if (t?.t === ',' || t?.t === ')') args.push({ type: 'missing' });
      else args.push(this.expr(0));
      this.skip();
      const sep = this.peek();
      if (sep?.t === ',') { this.i++; continue; }
      if (sep?.t === ')') { this.i++; if (args.length && args.at(-1)!.type === 'missing' && args.length === 1) return []; return args; }
      throw new FormulaError('Expected , or ) in a function call.');
    }
  }

  private primary(): Node {
    this.skip();
    const t = this.toks[this.i++];
    if (!t) throw new FormulaError('The formula ends too early.');
    switch (t.t) {
      case 'num': return { type: 'num', value: t.v };
      case 'str': return { type: 'str', value: t.v };
      case 'bool': return { type: 'bool', value: t.v };
      case 'err': return { type: 'err', code: t.v };
      case 'ref': return { type: 'ref', ref: t.v };
      case 'name': return t.sheet ? { type: 'name', name: t.v, sheet: t.sheet } : { type: 'name', name: t.v };
      case 'table': return { type: 'table', table: t.table, spec: t.spec };
      case 'func': {
        this.skip();
        if (this.toks[this.i]?.t !== '(') throw new FormulaError('Expected ( after a function name.');
        this.i++;
        return { type: 'call', name: t.v, args: this.args() };
      }
      case '(': {
        const items: Node[] = [this.expr(0)];
        this.skip();
        while (this.peek()?.t === ',') { this.i++; items.push(this.expr(0)); this.skip(); }
        if (this.toks[this.i++]?.t !== ')') throw new FormulaError('A ( is missing its closing ).');
        return items.length === 1 ? items[0]! : { type: 'union', items };
      }
      case '{': return this.arrayLit();
      case 'op':
        if (t.v === '-' || t.v === '+') { const arg = this.primary(); return { type: 'unary', op: t.v, arg }; }
        throw new FormulaError(`Unexpected "${t.v}".`);
      default:
        throw new FormulaError('Unexpected symbol in the formula.');
    }
  }

  private arrayLit(): Node {
    const rows: Node[][] = [[]];
    for (;;) {
      this.skip();
      const n = this.unary();
      if (!['num', 'str', 'bool', 'err', 'unary'].includes(n.type)) throw new FormulaError('Arrays can only hold numbers, text, TRUE/FALSE and errors.');
      rows.at(-1)!.push(n);
      this.skip();
      const t = this.toks[this.i++];
      if (t?.t === ',') continue;
      if (t?.t === ';') { rows.push([]); continue; }
      if (t?.t === '}') break;
      throw new FormulaError('An array is missing its closing }.');
    }
    if (rows.some((r) => r.length !== rows[0]!.length)) throw new FormulaError('Every row of an array needs the same number of values.');
    return { type: 'array', rows };
  }
}

const isRefLike = (n: Node): boolean =>
  n.type === 'ref' || n.type === 'name' || n.type === 'table' || n.type === 'union' || (n.type === 'binary' && (n.op === ':' || n.op === ' ')) ||
  (n.type === 'call' && ['INDEX', 'OFFSET', 'INDIRECT', 'CHOOSE', 'IF', 'XLOOKUP'].includes(n.name));

const cache = new Map<string, Node>();

/** Parses a formula (without the leading "="). Results are cached by text. */
export const parseFormula = (formula: string): Node => {
  const hit = cache.get(formula);
  if (hit) return hit;
  const node = new Parser(tokenize(formula)).parse();
  if (cache.size > 50_000) cache.clear();
  cache.set(formula, node);
  return node;
};

// ---------- Writing formulas back to text ----------

const colName = (c: number): string => {
  let s = '';
  let n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

export const quoteSheet = (name: string): string =>
  /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !readRef(name) && !/^(TRUE|FALSE)$/i.test(name) && !/^R\d*C\d*$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

export const refText = (r: RefNode): string => {
  let prefix = '';
  if (r.sheet !== undefined) {
    const both = r.sheet2 ? [r.sheet, r.sheet2] : [r.sheet];
    prefix = (both.every((x) => quoteSheet(x) === x) ? both.join(':') : `'${both.join(':').replace(/'/g, "''")}'`) + '!';
  }
  const d = (abs: boolean) => (abs ? '$' : '');
  if (r.kind === 'cols') return `${prefix}${d(r.c1Abs)}${colName(r.c1)}:${d(r.c2Abs)}${colName(r.c2)}`;
  if (r.kind === 'rows') return `${prefix}${d(r.r1Abs)}${r.r1 + 1}:${d(r.r2Abs)}${r.r2 + 1}`;
  const a = `${d(r.c1Abs)}${colName(r.c1)}${d(r.r1Abs)}${r.r1 + 1}`;
  if (r.kind === 'cell') return prefix + a;
  return `${prefix}${a}:${d(r.c2Abs)}${colName(r.c2)}${d(r.r2Abs)}${r.r2 + 1}`;
};

const tableText = (n: Extract<Node, { type: 'table' }>): string => {
  const { items, col1, col2 } = n.spec;
  const esc = (c: string) => c.replace(/([[\]#'@])/g, "'$1");
  const parts: string[] = [];
  const thisRow = items.includes('#This Row');
  for (const it of items) if (it !== '#This Row') parts.push(`[${it}]`);
  if (col1) parts.push(col2 ? `[${esc(col1)}]:[${esc(col2)}]` : `[${esc(col1)}]`);
  if (thisRow && parts.length <= 1 && !col2) return `${n.table}[@${col1 ? (/[^A-Za-z0-9_]/.test(col1) ? `[${esc(col1)}]` : esc(col1)) : ''}]`;
  if (thisRow) parts.unshift('@');
  if (parts.length === 1 && !thisRow && col1 && !col2) return `${n.table}[${esc(col1)}]`;
  return `${n.table}[${parts.join(',')}]`;
};

const PREC_OUT: Record<string, number> = { ...PREC, ':': 9, ' ': 8 };

export const nodeText = (n: Node, parentPrec = 0): string => {
  switch (n.type) {
    case 'num': return numText(n.value);
    case 'str': return `"${n.value.replace(/"/g, '""')}"`;
    case 'bool': return n.value ? 'TRUE' : 'FALSE';
    case 'err': return n.code;
    case 'missing': return '';
    case 'array': return `{${n.rows.map((r) => r.map((x) => nodeText(x)).join(',')).join(';')}}`;
    case 'ref': return n.ref.external ? '#REF!' : refText(n.ref);
    case 'name': return n.sheet ? `${quoteSheet(n.sheet)}!${n.name}` : n.name;
    case 'table': return tableText(n);
    case 'call': return `${n.name}(${n.args.map((a) => nodeText(a)).join(',')})`;
    case 'invoke': return `${nodeText(n.callee, 10)}(${n.args.map((a) => nodeText(a)).join(',')})`;
    case 'unary': return n.op === '%' ? `${nodeText(n.arg, 7)}%` : `${n.op}${nodeText(n.arg, 6)}`;
    case 'union': return `(${n.items.map((x) => nodeText(x)).join(',')})`;
    case 'spill': return `${nodeText(n.arg, 10)}#`;
    case 'binary': {
      const p = PREC_OUT[n.op]!;
      const s = `${nodeText(n.left, p)}${n.op}${nodeText(n.right, p + 1)}`;
      return p < parentPrec ? `(${s})` : s;
    }
  }
};

const numText = (v: number): string => {
  if (Number.isInteger(v) && Math.abs(v) < 1e21) return String(v);
  const s = String(v);
  return s.includes('e') ? s.replace('e', 'E') : s;
};

/** Visits every node, depth first. */
export const walk = (n: Node, f: (n: Node) => void): void => {
  f(n);
  switch (n.type) {
    case 'array': n.rows.forEach((r) => r.forEach((x) => walk(x, f))); break;
    case 'call': n.args.forEach((x) => walk(x, f)); break;
    case 'invoke': walk(n.callee, f); n.args.forEach((x) => walk(x, f)); break;
    case 'unary': case 'spill': walk(n.arg, f); break;
    case 'binary': walk(n.left, f); walk(n.right, f); break;
    case 'union': n.items.forEach((x) => walk(x, f)); break;
    default: break;
  }
};

/** Rebuilds a tree with `f` applied to each node, children first. */
export const mapNode = (n: Node, f: (n: Node) => Node): Node => {
  let c: Node;
  switch (n.type) {
    case 'array': c = { ...n, rows: n.rows.map((r) => r.map((x) => mapNode(x, f))) }; break;
    case 'call': c = { ...n, args: n.args.map((x) => mapNode(x, f)) }; break;
    case 'invoke': c = { ...n, callee: mapNode(n.callee, f), args: n.args.map((x) => mapNode(x, f)) }; break;
    case 'unary': c = { ...n, arg: mapNode(n.arg, f) }; break;
    case 'spill': c = { ...n, arg: mapNode(n.arg, f) }; break;
    case 'binary': c = { ...n, left: mapNode(n.left, f), right: mapNode(n.right, f) }; break;
    case 'union': c = { ...n, items: n.items.map((x) => mapNode(x, f)) }; break;
    default: c = n;
  }
  return f(c);
};
