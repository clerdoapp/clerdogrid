// Formulas of calculated fields and calculated items: field or item names combined with operators
// and worksheet functions, e.g. Sales*0.1, 'Unit Price'*Units, East+West or Region[East]*2.
import { getFunction } from '../../engine/registry';
import { toNumber, toText } from '../../engine/coerce';
import { CellError, Err, isError, isMatrix, type FnContext, type Scalar, type Value } from '../../engine/types';
import { T } from '../../ui/i18n';

export type PNode =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; name: string; item?: string }
  | { t: 'call'; name: string; args: PNode[] }
  | { t: 'bin'; op: string; a: PNode; b: PNode }
  | { t: 'neg'; a: PNode }
  | { t: 'pct'; a: PNode };

type Tok = { k: 'num'; v: number } | { k: 'str'; v: string } | { k: 'name'; v: string } | { k: 'q'; v: string } | { k: 'op'; v: string };

const tokenize = (src: string): Tok[] => {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i]!;
    if (/\s/.test(c)) { i++; continue; }
    if (/[0-9.]/.test(c)) {
      const m = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(src.slice(i));
      if (!m) throw new Error(T('Bad number'));
      out.push({ k: 'num', v: +m[0] }); i += m[0].length; continue;
    }
    if (c === '"' || c === "'") {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= src.length) throw new Error(T('Missing quote'));
        if (src[j] === c) { if (src[j + 1] === c) { s += c; j += 2; continue; } break; }
        s += src[j++];
      }
      out.push({ k: c === '"' ? 'str' : 'q', v: s }); i = j + 1; continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { out.push({ k: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%(),[]'.includes(c)) { out.push({ k: 'op', v: c }); i++; continue; }
    const m = /^[\p{L}_\\][\p{L}\p{N}_.\\]*/u.exec(src.slice(i));
    if (!m) throw new Error(T('Unexpected "{c}"', { c }));
    out.push({ k: 'name', v: m[0] }); i += m[0].length;
  }
  return out;
};

export const parsePivotFormula = (src: string): PNode => {
  const toks = tokenize(src.replace(/^\s*=/, ''));
  let p = 0;
  const peek = () => toks[p];
  const isOp = (v: string) => { const t = toks[p]; return t?.k === 'op' && t.v === v; };
  const expect = (v: string) => { if (!isOp(v)) throw new Error(T('Expected "{v}"', { v })); p++; };
  const bin = (next: () => PNode, ops: string[]) => (): PNode => {
    let a = next();
    for (;;) {
      const t = peek();
      if (t?.k === 'op' && ops.includes(t.v)) { p++; a = { t: 'bin', op: t.v, a, b: next() }; } else return a;
    }
  };
  const primary = (): PNode => {
    const t = toks[p++];
    if (!t) throw new Error(T('Unexpected end'));
    if (t.k === 'num') return { t: 'num', v: t.v };
    if (t.k === 'str') return { t: 'str', v: t.v };
    if (t.k === 'op' && t.v === '(') { const e = cmp(); expect(')'); return e; }
    if (t.k === 'name' || t.k === 'q') {
      if (t.k === 'name' && isOp('(')) {
        p++;
        const args: PNode[] = [];
        if (!isOp(')')) for (;;) { args.push(cmp()); if (isOp(',')) { p++; continue; } break; }
        expect(')');
        return { t: 'call', name: t.v.toUpperCase(), args };
      }
      if (isOp('[')) {
        p++;
        const it = toks[p++];
        if (!it || it.k === 'op') throw new Error(T('Expected an item'));
        expect(']');
        return { t: 'ref', name: t.v, item: String(it.v) };
      }
      return { t: 'ref', name: t.v };
    }
    throw new Error(T('Unexpected "{v}"', { v: t.v }));
  };
  const postfix = (): PNode => { let a = primary(); while (isOp('%')) { p++; a = { t: 'pct', a }; } return a; };
  const unary = (): PNode => {
    if (isOp('-')) { p++; return { t: 'neg', a: unary() }; }
    if (isOp('+')) { p++; return unary(); }
    return postfix();
  };
  const pow = bin(unary, ['^']);
  const term = bin(pow, ['*', '/']);
  const add = bin(term, ['+', '-']);
  const cat = bin(add, ['&']);
  const cmp = bin(cat, ['=', '<>', '<', '>', '<=', '>=']);
  const e = cmp();
  if (p < toks.length) throw new Error(T('Unexpected text at the end'));
  return e;
};

/** Names referenced by a formula (with the item when written as Field[Item]). */
export const formulaRefs = (n: PNode, out: { name: string; item?: string }[] = []): { name: string; item?: string }[] => {
  switch (n.t) {
    case 'ref': out.push({ name: n.name, item: n.item }); break;
    case 'call': n.args.forEach((a) => formulaRefs(a, out)); break;
    case 'bin': formulaRefs(n.a, out); formulaRefs(n.b, out); break;
    case 'neg': case 'pct': formulaRefs(n.a, out); break;
  }
  return out;
};

const stubCtx: FnContext = {
  sheet: 0, row: 0, col: 0, date1904: false,
  values: () => [[null]], cell: () => null, usedBounds: () => ({ rows: -1, cols: -1 }), formula: () => null,
  sheetIndex: () => -1, sheetName: () => '', sheetCount: () => 0, call: (fn, args) => fn.invoke(args),
  rowHidden: () => false, isSubtotalCell: () => false, numberFormat: () => 'General',
};

const numOf = (v: Scalar): number | CellError => (v === null ? 0 : toNumber(v));

export const evalPivotFormula = (n: PNode, resolve: (name: string, item?: string) => Scalar): Scalar => {
  const ev = (x: PNode): Scalar => {
    switch (x.t) {
      case 'num': return x.v;
      case 'str': return x.v;
      case 'ref': return resolve(x.name, x.item);
      case 'neg': { const a = numOf(ev(x.a)); return isError(a) ? a : -a; }
      case 'pct': { const a = numOf(ev(x.a)); return isError(a) ? a : a / 100; }
      case 'call': {
        const def = getFunction(x.name);
        if (!def) return new CellError('#NAME?', x.name);
        const args: Value[] = x.args.map(ev);
        if (args.length < def.min || args.length > def.max) return Err.value;
        try {
          const r = def.fn(args, stubCtx);
          if (isMatrix(r)) return r[0]?.[0] ?? null;
          return r === null || typeof r !== 'object' || r instanceof CellError ? r as Scalar : Err.value;
        } catch { return Err.value; }
      }
      case 'bin': {
        const a = ev(x.a), b = ev(x.b);
        if (x.op === '&') { const s = toText(a ?? ''), t = toText(b ?? ''); return isError(s) ? s : isError(t) ? t : s + t; }
        if (['=', '<>', '<', '>', '<=', '>='].includes(x.op)) {
          if (isError(a)) return a; if (isError(b)) return b;
          const c = typeof a === 'string' && typeof b === 'string' ? a.toLowerCase().localeCompare(b.toLowerCase()) : (Number(a ?? 0) - Number(b ?? 0));
          return x.op === '=' ? c === 0 : x.op === '<>' ? c !== 0 : x.op === '<' ? c < 0 : x.op === '>' ? c > 0 : x.op === '<=' ? c <= 0 : c >= 0;
        }
        const p = numOf(a), q = numOf(b);
        if (isError(p)) return p; if (isError(q)) return q;
        switch (x.op) {
          case '+': return p + q;
          case '-': return p - q;
          case '*': return p * q;
          case '/': return q === 0 ? Err.div0 : p / q;
          case '^': { const r = Math.pow(p, q); return Number.isFinite(r) ? r : Err.num; }
        }
        return Err.value;
      }
    }
  };
  return ev(n);
};
