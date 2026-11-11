// XPath 1.0 (W3C Recommendation, 1999) evaluator for FILTERXML, with its own small XML reader.
// Supports all thirteen axes except namespace, abbreviated syntax, predicates, unions, the operators,
// and the core function library except id() and lang(). Prefixes in the expression resolve to the
// namespaces the document declares for them; namespace-uri() and local-name() report a node's parts.

import { decodeEntities } from '../formats/xml';

export class XPathError extends Error {}

interface PNode {
  kind: 'root' | 'element' | 'attribute' | 'text';
  name: string;
  value: string;
  parent: PNode | null;
  children: PNode[];
  attrs: PNode[];
  order: number;
  /** Namespace URI ('' for none) and the name without its prefix. */
  ns: string;
  local: string;
}

/** Prefix -> namespace URI declarations found in the document (the first one for each prefix wins). */
let docPrefixes = new Map<string, string>();

/**
 * Parses XML into the XPath data model, keeping qualified names, mixed content order and document order.
 * Throws XPathError for malformed input (unclosed or mismatched tags, a missing root element).
 */
export const buildDocument = (xml: string): PNode => {
  let order = 0;
  const root: PNode = { kind: 'root', name: '', value: '', parent: null, children: [], attrs: [], order: order++, ns: '', local: '' };
  const stack: PNode[] = [root];
  // Namespace scopes, one per open element; "xml" is always bound.
  const scopes: Map<string, string>[] = [new Map([['xml', 'http://www.w3.org/XML/1998/namespace']])];
  const declared = new Map<string, string>();
  const split = (q: string): [string, string] => { const i = q.indexOf(':'); return i < 0 ? ['', q] : [q.slice(0, i), q.slice(i + 1)]; };
  let i = xml.charCodeAt(0) === 0xfeff ? 1 : 0;
  const n = xml.length;
  const text = (t: string): void => {
    if (!t) return;
    const top = stack[stack.length - 1]!;
    if (top === root) { if (t.trim()) throw new XPathError('text outside the root element'); return; }
    const last = top.children[top.children.length - 1];
    if (last?.kind === 'text') last.value += t;
    else top.children.push({ kind: 'text', name: '', value: t, parent: top, children: [], attrs: [], order: order++, ns: '', local: '' });
  };
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) { text(decodeEntities(xml.slice(i))); break; }
    if (lt > i) text(decodeEntities(xml.slice(i, lt)));
    if (xml.startsWith('<!--', lt)) { const e = xml.indexOf('-->', lt + 4); if (e < 0) throw new XPathError('unclosed comment'); i = e + 3; continue; }
    if (xml.startsWith('<![CDATA[', lt)) { const e = xml.indexOf(']]>', lt + 9); if (e < 0) throw new XPathError('unclosed CDATA'); text(xml.slice(lt + 9, e)); i = e + 3; continue; }
    if (xml.startsWith('<?', lt)) { const e = xml.indexOf('?>', lt + 2); if (e < 0) throw new XPathError('unclosed declaration'); i = e + 2; continue; }
    if (xml.startsWith('<!', lt)) {
      // DOCTYPE, possibly with an internal subset in brackets.
      let depth = 0, j = lt + 2;
      for (; j < n; j++) { const c = xml[j]; if (c === '[') depth++; else if (c === ']') depth--; else if (c === '>' && depth <= 0) break; }
      i = j + 1; continue;
    }
    const gt = (() => { let q: string | null = null; for (let j = lt + 1; j < n; j++) { const c = xml[j]!; if (q) { if (c === q) q = null; } else if (c === '"' || c === "'") q = c; else if (c === '>') return j; } return -1; })();
    if (gt < 0) throw new XPathError('unclosed tag');
    const body = xml.slice(lt + 1, gt);
    i = gt + 1;
    if (body.startsWith('/')) {
      const name = body.slice(1).trim();
      const top = stack.pop();
      if (!top || top === root || top.name !== name) throw new XPathError(`mismatched </${name}>`);
      scopes.pop();
      continue;
    }
    const selfClose = body.endsWith('/');
    const inner = selfClose ? body.slice(0, -1) : body;
    const m = /^([^\s/>]+)/.exec(inner);
    if (!m) throw new XPathError('bad tag');
    const parent = stack[stack.length - 1]!;
    if (parent === root && root.children.some((c) => c.kind === 'element')) throw new XPathError('more than one root element');
    const el: PNode = { kind: 'element', name: m[1]!, value: '', parent, children: [], attrs: [], order: order++, ns: '', local: '' };
    const re = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    const rest = inner.slice(m[0].length);
    const raw: [string, string][] = [];
    const scope = new Map(scopes[scopes.length - 1]!);
    for (let a = re.exec(rest); a; a = re.exec(rest)) {
      const k = a[1]!, v = decodeEntities(a[3] ?? a[4] ?? '');
      if (k === 'xmlns') scope.set('', v);
      else if (k.startsWith('xmlns:')) { scope.set(k.slice(6), v); if (!declared.has(k.slice(6))) declared.set(k.slice(6), v); }
      else raw.push([k, v]);
    }
    const [pre, local] = split(el.name);
    if (pre && !scope.has(pre)) throw new XPathError(`undeclared prefix ${pre}`);
    el.ns = scope.get(pre) ?? '';
    el.local = local;
    for (const [k, v] of raw) {
      const [ap, al] = split(k);
      if (ap && !scope.has(ap)) throw new XPathError(`undeclared prefix ${ap}`);
      // Attributes without a prefix are in no namespace.
      el.attrs.push({ kind: 'attribute', name: k, value: v, parent: el, children: [], attrs: [], order: order++, ns: ap ? scope.get(ap)! : '', local: al });
    }
    parent.children.push(el);
    if (!selfClose) { stack.push(el); scopes.push(scope); }
  }
  if (stack.length > 1) throw new XPathError(`unclosed <${stack[stack.length - 1]!.name}>`);
  if (!root.children.some((c) => c.kind === 'element')) throw new XPathError('no XML element found');
  docPrefixes = declared;
  return root;
};


const stringValue = (n: PNode): string => {
  if (n.kind === 'attribute' || n.kind === 'text') return n.value;
  let s = '';
  const walk = (x: PNode): void => { for (const c of x.children) { if (c.kind === 'text') s += c.value; else walk(c); } };
  walk(n);
  return s;
};

type XValue = PNode[] | string | number | boolean;

// ---------- tokenizer ----------

type Tok = { t: 'num'; v: number } | { t: 'str'; v: string } | { t: 'name'; v: string } | { t: 'op'; v: string } | { t: 'var'; v: string };

const tokenize = (src: string): Tok[] => {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i]!;
    if (/\s/.test(c)) { i++; continue; }
    if (c === '"' || c === "'") {
      const j = src.indexOf(c, i + 1);
      if (j < 0) throw new XPathError('unterminated string');
      out.push({ t: 'str', v: src.slice(i + 1, j) }); i = j + 1; continue;
    }
    const num = /^(\d+(\.\d*)?|\.\d+)/.exec(src.slice(i));
    if (num) { out.push({ t: 'num', v: Number(num[0]) }); i += num[0].length; continue; }
    const two = src.slice(i, i + 2);
    if (['//', '..', '::', '!=', '<=', '>='].includes(two)) { out.push({ t: 'op', v: two }); i += 2; continue; }
    if ('/()[]@,|+-=<>*.$'.includes(c)) {
      if (c === '$') {
        const m = /^\$([A-Za-z_][\w.-]*)/.exec(src.slice(i));
        if (!m) throw new XPathError('bad variable');
        out.push({ t: 'var', v: m[1]! }); i += m[0].length; continue;
      }
      out.push({ t: 'op', v: c }); i++; continue;
    }
    const nm = /^[A-Za-z_À-￿][\w.\-·À-￿]*(:(\*|[A-Za-z_À-￿][\w.\-·À-￿]*))?/.exec(src.slice(i));
    if (nm) { out.push({ t: 'name', v: nm[0] }); i += nm[0].length; continue; }
    throw new XPathError(`unexpected character ${c}`);
  }
  return out;
};

// ---------- parser ----------

type Expr =
  | { k: 'num'; v: number } | { k: 'str'; v: string }
  | { k: 'bin'; op: string; a: Expr; b: Expr }
  | { k: 'neg'; a: Expr }
  | { k: 'fn'; name: string; args: Expr[] }
  | { k: 'path'; abs: boolean; filter?: Expr; filterPreds?: Expr[]; steps: Step[] };

interface Step { axis: string; test: { k: 'name'; v: string } | { k: 'type'; v: string }; preds: Expr[] }

const AXES = new Set(['ancestor', 'ancestor-or-self', 'attribute', 'child', 'descendant', 'descendant-or-self', 'following', 'following-sibling', 'parent', 'preceding', 'preceding-sibling', 'self']);
const NODE_TYPES = new Set(['node', 'text', 'comment', 'processing-instruction']);

class Parser {
  private i = 0;
  constructor(private toks: Tok[]) {}
  private peek(k = 0): Tok | undefined { return this.toks[this.i + k]; }
  private isOp(v: string, k = 0): boolean { const t = this.peek(k); return !!t && t.t === 'op' && t.v === v; }
  private eatOp(v: string): boolean { if (this.isOp(v)) { this.i++; return true; } return false; }
  private expect(v: string): void { if (!this.eatOp(v)) throw new XPathError(`expected ${v}`); }
  /** Called only right after an operand, where "and", "or", "div" and "mod" can only be operators. */
  private isOperatorName(v: string): boolean {
    const t = this.peek();
    return !!t && t.t === 'name' && t.v === v;
  }

  parse(): Expr {
    const e = this.or();
    if (this.i < this.toks.length) throw new XPathError('unexpected token');
    return e;
  }
  private or(): Expr { let a = this.and(); while (this.isOperatorName('or')) { this.i++; a = { k: 'bin', op: 'or', a, b: this.and() }; } return a; }
  private and(): Expr { let a = this.eq(); while (this.isOperatorName('and')) { this.i++; a = { k: 'bin', op: 'and', a, b: this.eq() }; } return a; }
  private eq(): Expr {
    let a = this.rel();
    for (;;) {
      const op = this.isOp('=') ? '=' : this.isOp('!=') ? '!=' : null;
      if (!op) return a;
      this.i++; a = { k: 'bin', op, a, b: this.rel() };
    }
  }
  private rel(): Expr {
    let a = this.add();
    for (;;) {
      const op = ['<=', '>=', '<', '>'].find((o) => this.isOp(o));
      if (!op) return a;
      this.i++; a = { k: 'bin', op, a, b: this.add() };
    }
  }
  private add(): Expr {
    let a = this.mul();
    for (;;) {
      const op = this.isOp('+') ? '+' : this.isOp('-') ? '-' : null;
      if (!op) return a;
      this.i++; a = { k: 'bin', op, a, b: this.mul() };
    }
  }
  private mul(): Expr {
    let a = this.unary();
    for (;;) {
      let op: string | null = null;
      if (this.isOp('*')) op = '*';
      else if (this.isOperatorName('div')) op = 'div';
      else if (this.isOperatorName('mod')) op = 'mod';
      if (!op) return a;
      this.i++; a = { k: 'bin', op, a, b: this.unary() };
    }
  }
  private unary(): Expr { if (this.eatOp('-')) return { k: 'neg', a: this.unary() }; return this.union(); }
  private union(): Expr { let a = this.pathExpr(); while (this.eatOp('|')) a = { k: 'bin', op: '|', a, b: this.pathExpr() }; return a; }

  private pathExpr(): Expr {
    const t = this.peek();
    const startsFilter = t && (t.t === 'num' || t.t === 'str' || t.t === 'var' || (t.t === 'op' && t.v === '(')
      || (t.t === 'name' && this.isOp('(', 1) && !NODE_TYPES.has(t.v) && !AXES.has(t.v)));
    if (startsFilter) {
      const prim = this.primary();
      const preds: Expr[] = [];
      while (this.isOp('[')) preds.push(this.pred());
      if (this.isOp('/') || this.isOp('//')) {
        const steps: Step[] = [];
        this.relSteps(steps);
        return { k: 'path', abs: false, filter: prim, filterPreds: preds, steps };
      }
      return preds.length ? { k: 'path', abs: false, filter: prim, filterPreds: preds, steps: [] } : prim;
    }
    const steps: Step[] = [];
    if (this.eatOp('/')) {
      const n = this.peek();
      const more = n && (n.t === 'name' || (n.t === 'op' && ['*', '@', '.', '..'].includes(n.v)));
      if (more) this.steps(steps);
      return { k: 'path', abs: true, steps };
    }
    if (this.eatOp('//')) {
      steps.push({ axis: 'descendant-or-self', test: { k: 'type', v: 'node' }, preds: [] });
      this.steps(steps);
      return { k: 'path', abs: true, steps };
    }
    this.steps(steps);
    return { k: 'path', abs: false, steps };
  }
  private relSteps(steps: Step[]): void {
    for (;;) {
      if (this.eatOp('/')) this.step(steps);
      else if (this.eatOp('//')) { steps.push({ axis: 'descendant-or-self', test: { k: 'type', v: 'node' }, preds: [] }); this.step(steps); }
      else return;
    }
  }
  private steps(steps: Step[]): void { this.step(steps); this.relSteps(steps); }
  private step(steps: Step[]): void {
    if (this.eatOp('.')) { steps.push({ axis: 'self', test: { k: 'type', v: 'node' }, preds: [] }); return; }
    if (this.eatOp('..')) { steps.push({ axis: 'parent', test: { k: 'type', v: 'node' }, preds: [] }); return; }
    let axis = 'child';
    if (this.eatOp('@')) axis = 'attribute';
    else {
      const t = this.peek();
      if (t?.t === 'name' && this.isOp('::', 1)) {
        if (!AXES.has(t.v)) throw new XPathError(`unknown axis ${t.v}`);
        axis = t.v; this.i += 2;
      }
    }
    let test: Step['test'];
    const t = this.peek();
    if (t?.t === 'op' && t.v === '*') { this.i++; test = { k: 'name', v: '*' }; }
    else if (t?.t === 'name') {
      this.i++;
      if (NODE_TYPES.has(t.v) && this.isOp('(')) {
        this.i++;
        if (this.peek()?.t === 'str') this.i++;
        this.expect(')');
        test = { k: 'type', v: t.v };
      } else test = { k: 'name', v: t.v };
    } else throw new XPathError('step expected');
    const preds: Expr[] = [];
    while (this.isOp('[')) preds.push(this.pred());
    steps.push({ axis, test, preds });
  }
  private pred(): Expr { this.expect('['); const e = this.or(); this.expect(']'); return e; }
  private primary(): Expr {
    const t = this.peek()!;
    this.i++;
    if (t.t === 'num') return { k: 'num', v: t.v };
    if (t.t === 'str') return { k: 'str', v: t.v };
    if (t.t === 'var') throw new XPathError('variables are not supported');
    if (t.t === 'op' && t.v === '(') { const e = this.or(); this.expect(')'); return e; }
    if (t.t === 'name') {
      this.expect('(');
      const args: Expr[] = [];
      if (!this.isOp(')')) { do args.push(this.or()); while (this.eatOp(',')); }
      this.expect(')');
      return { k: 'fn', name: t.v, args };
    }
    throw new XPathError('unexpected token');
  }
}

// ---------- evaluation ----------

interface Ctx { node: PNode; pos: number; size: number }

const docOrder = (ns: PNode[]): PNode[] => [...new Set(ns)].sort((a, b) => a.order - b.order);

const descendants = (n: PNode, out: PNode[]): void => { for (const c of n.children) { out.push(c); descendants(c, out); } };

const axisNodes = (n: PNode, axis: string): PNode[] => {
  const out: PNode[] = [];
  switch (axis) {
    case 'child': return n.children;
    case 'attribute': return n.attrs;
    case 'self': return [n];
    case 'parent': return n.parent ? [n.parent] : [];
    case 'descendant': descendants(n, out); return out;
    case 'descendant-or-self': out.push(n); descendants(n, out); return out;
    case 'ancestor': for (let p = n.parent; p; p = p.parent) out.push(p); return out;
    case 'ancestor-or-self': for (let p: PNode | null = n; p; p = p.parent) out.push(p); return out;
    case 'following-sibling': case 'preceding-sibling': {
      if (!n.parent || n.kind === 'attribute') return [];
      const sib = n.parent.children, i = sib.indexOf(n);
      return axis === 'following-sibling' ? sib.slice(i + 1) : sib.slice(0, i).reverse();
    }
    case 'following': case 'preceding': {
      let root = n; while (root.parent) root = root.parent;
      const all: PNode[] = []; descendants(root, all);
      const anc = new Set<PNode>(); for (let p = n.parent; p; p = p.parent) anc.add(p);
      const base = n.kind === 'attribute' ? n.parent! : n;
      const endOrder = (x: PNode): number => { let last = x; while (last.children.length) last = last.children[last.children.length - 1]!; return last.order; };
      if (axis === 'following') return all.filter((x) => x.order > endOrder(base));
      return all.filter((x) => x.order < base.order && !anc.has(x)).reverse();
    }
    default: return [];
  }
};

const matches = (n: PNode, step: Step): boolean => {
  if (step.test.k === 'type') {
    if (step.test.v === 'node') return true;
    if (step.test.v === 'text') return n.kind === 'text';
    return false;
  }
  const principal = step.axis === 'attribute' ? 'attribute' : 'element';
  if (n.kind !== principal) return false;
  const t = step.test.v;
  if (t === '*') return true;
  // A prefix in the expression stands for the namespace the document declares for it; a name without one
  // means no namespace (XPath 1.0, section 2.3).
  const i = t.indexOf(':');
  const pre = i < 0 ? '' : t.slice(0, i), local = i < 0 ? t : t.slice(i + 1);
  const uri = pre ? docPrefixes.get(pre) : '';
  if (uri === undefined) throw new XPathError(`unknown prefix ${pre}`);
  return n.ns === uri && (local === '*' || n.local === local);
};

const toStr = (v: XValue): string => {
  if (Array.isArray(v)) return v.length ? stringValue(docOrder(v)[0]!) : '';
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  if (typeof v === 'number') return Number.isNaN(v) ? 'NaN' : v === Infinity ? 'Infinity' : v === -Infinity ? '-Infinity' : Number.isInteger(v) ? String(v) : String(v);
  return v;
};
const toNum = (v: XValue): number => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const s = toStr(v).trim();
  return /^-?(\d+(\.\d*)?|\.\d+)$/.test(s) ? Number(s) : NaN;
};
const toBool = (v: XValue): boolean => (Array.isArray(v) ? v.length > 0 : typeof v === 'number' ? v !== 0 && !Number.isNaN(v) : typeof v === 'string' ? v.length > 0 : v);

const compare = (op: string, a: XValue, b: XValue): boolean => {
  const cmp = (x: string | number | boolean, y: string | number | boolean): boolean => {
    if (op === '=' || op === '!=') {
      let r: boolean;
      if (typeof x === 'boolean' || typeof y === 'boolean') r = toBool(x) === toBool(y);
      else if (typeof x === 'number' || typeof y === 'number') r = toNum(x) === toNum(y);
      else r = x === y;
      return op === '=' ? r : !r;
    }
    const p = toNum(x), q = toNum(y);
    return op === '<' ? p < q : op === '<=' ? p <= q : op === '>' ? p > q : p >= q;
  };
  if (Array.isArray(a) && Array.isArray(b)) return a.some((x) => b.some((y) => cmp(stringValue(x), stringValue(y))));
  if (Array.isArray(a)) { const bb = b as string | number | boolean; return typeof bb === 'boolean' ? cmp(toBool(a), bb) : a.some((x) => cmp(typeof bb === 'number' ? toNum(stringValue(x)) : stringValue(x), bb)); }
  if (Array.isArray(b)) { const aa = a as string | number | boolean; return typeof aa === 'boolean' ? cmp(aa, toBool(b)) : b.some((y) => cmp(aa, typeof aa === 'number' ? toNum(stringValue(y)) : stringValue(y))); }
  return cmp(a as string | number | boolean, b as string | number | boolean);
};

const round = (x: number): number => (Number.isNaN(x) || !Number.isFinite(x) ? x : Math.floor(x + 0.5));

const callFn = (name: string, args: Expr[], c: Ctx): XValue => {
  const ev = (i: number): XValue => evaluate(args[i]!, c);
  const str = (i: number): string => (args.length > i ? toStr(ev(i)) : stringValue(c.node));
  const need = (min: number, max = min): void => { if (args.length < min || args.length > max) throw new XPathError(`wrong number of arguments for ${name}()`); };
  switch (name) {
    case 'last': need(0); return c.size;
    case 'position': need(0); return c.pos;
    case 'count': { need(1); const v = ev(0); if (!Array.isArray(v)) throw new XPathError('count() needs a node-set'); return v.length; }
    case 'name': case 'local-name': {
      need(0, 1);
      const ns = args.length ? ev(0) : [c.node];
      if (!Array.isArray(ns)) throw new XPathError(`${name}() needs a node-set`);
      const n = docOrder(ns)[0];
      if (!n || (n.kind !== 'element' && n.kind !== 'attribute')) return '';
      return name === 'name' ? n.name : n.local;
    }
    case 'namespace-uri': {
      need(0, 1);
      const ns = args.length ? ev(0) : [c.node];
      if (!Array.isArray(ns)) throw new XPathError('namespace-uri() needs a node-set');
      return docOrder(ns)[0]?.ns ?? '';
    }
    case 'string': need(0, 1); return str(0);
    case 'concat': if (args.length < 2) throw new XPathError('concat() needs two or more arguments'); return args.map((_, i) => toStr(ev(i))).join('');
    case 'starts-with': need(2); return str(0).startsWith(str(1));
    case 'contains': need(2); return str(0).includes(str(1));
    case 'substring-before': { need(2); const s = str(0), t = str(1), i = s.indexOf(t); return i < 0 ? '' : s.slice(0, i); }
    case 'substring-after': { need(2); const s = str(0), t = str(1), i = s.indexOf(t); return i < 0 ? '' : s.slice(i + t.length); }
    case 'substring': {
      need(2, 3);
      const s = [...str(0)];
      const start = round(toNum(ev(1)));
      const len = args.length > 2 ? round(toNum(ev(2))) : Infinity;
      return s.filter((_, i) => i + 1 >= start && i + 1 < start + len).join('');
    }
    case 'string-length': need(0, 1); return [...str(0)].length;
    case 'normalize-space': need(0, 1); return str(0).replace(/[ \t\r\n]+/g, ' ').trim();
    case 'translate': {
      need(3);
      const from = [...str(1)], to = [...str(2)];
      return [...str(0)].map((ch) => { const i = from.indexOf(ch); return i < 0 ? ch : (to[i] ?? ''); }).join('');
    }
    case 'boolean': need(1); return toBool(ev(0));
    case 'not': need(1); return !toBool(ev(0));
    case 'true': need(0); return true;
    case 'false': need(0); return false;
    case 'number': need(0, 1); return args.length ? toNum(ev(0)) : toNum(stringValue(c.node));
    case 'sum': { need(1); const v = ev(0); if (!Array.isArray(v)) throw new XPathError('sum() needs a node-set'); return v.reduce((a, n) => a + toNum(stringValue(n)), 0); }
    case 'floor': need(1); return Math.floor(toNum(ev(0)));
    case 'ceiling': need(1); return Math.ceil(toNum(ev(0)));
    case 'round': need(1); return round(toNum(ev(0)));
    default: throw new XPathError(`unknown function ${name}()`);
  }
};

const applyPreds = (nodes: PNode[], preds: Expr[], reverse: boolean): PNode[] => {
  let cur = nodes;
  for (const p of preds) {
    const size = cur.length;
    cur = cur.filter((node, i) => {
      const pos = i + 1;
      const v = evaluate(p, { node, pos, size });
      return typeof v === 'number' ? v === pos : toBool(v);
    });
    void reverse;
  }
  return cur;
};

const REVERSE = new Set(['ancestor', 'ancestor-or-self', 'preceding', 'preceding-sibling']);

const evaluate = (e: Expr, c: Ctx): XValue => {
  switch (e.k) {
    case 'num': return e.v;
    case 'str': return e.v;
    case 'neg': return -toNum(evaluate(e.a, c));
    case 'fn': return callFn(e.name, e.args, c);
    case 'bin': {
      if (e.op === 'or') return toBool(evaluate(e.a, c)) || toBool(evaluate(e.b, c));
      if (e.op === 'and') return toBool(evaluate(e.a, c)) && toBool(evaluate(e.b, c));
      const a = evaluate(e.a, c), b = evaluate(e.b, c);
      if (e.op === '|') {
        if (!Array.isArray(a) || !Array.isArray(b)) throw new XPathError('| needs node-sets');
        return docOrder([...a, ...b]);
      }
      if (['=', '!=', '<', '<=', '>', '>='].includes(e.op)) return compare(e.op, a, b);
      const x = toNum(a), y = toNum(b);
      switch (e.op) {
        case '+': return x + y;
        case '-': return x - y;
        case '*': return x * y;
        case 'div': return x / y;
        default: return x % y;
      }
    }
    case 'path': {
      let nodes: PNode[];
      if (e.filter) {
        const v = evaluate(e.filter, c);
        if (!Array.isArray(v)) { if (e.steps.length || e.filterPreds?.length) throw new XPathError('a path needs a node-set'); return v; }
        nodes = applyPreds(docOrder(v), e.filterPreds ?? [], false);
      } else if (e.abs) { let r = c.node; while (r.parent) r = r.parent; nodes = [r]; }
      else nodes = [c.node];
      for (const st of e.steps) {
        const next: PNode[] = [];
        for (const n of nodes) {
          const cand = axisNodes(n, st.axis).filter((x) => matches(x, st));
          next.push(...applyPreds(cand, st.preds, REVERSE.has(st.axis)));
        }
        nodes = docOrder(next);
      }
      return nodes;
    }
  }
};

export type XPathResult = { nodes: string[] } | { value: string | number | boolean };

/** Evaluates `expr` against `xml`. Throws XPathError for bad XML paths. */
export const xpath = (xml: string, expr: string): XPathResult => {
  const doc = buildDocument(xml);
  const e = new Parser(tokenize(expr)).parse();
  const v = evaluate(e, { node: doc, pos: 1, size: 1 });
  if (Array.isArray(v)) return { nodes: v.map(stringValue) };
  return { value: v };
};
