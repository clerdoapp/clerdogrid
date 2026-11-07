import { compareValues, first, numberToText, toBool, toNumber, toText } from './coerce';
import { parseFormula, FormulaError, type Node, type RefNode } from './parser';
import { getFunction } from './registry';
import {
  CellError, Err, isError, isLambda, isMatrix, isRef, isRefs, MAX_COLS, MAX_ROWS, ref,
  type FnContext, type FunctionDef, type LambdaValue, type Matrix, type RangeRef, type Scalar, type Value,
} from './types';

/** What the evaluator needs from the workbook. */
export interface Host {
  getValue(sheet: number, row: number, col: number): Scalar;
  usedBounds(sheet: number): { rows: number; cols: number };
  sheetIndex(name: string): number;
  sheetName(index: number): string;
  sheetCount(): number;
  formula(sheet: number, row: number, col: number): string | null;
  /** Defined name lookup: sheet-scoped first, then workbook. Returns the formula text. */
  definedName(name: string, sheet: number): string | null;
  table(name: string): { sheet: number; r1: number; c1: number; r2: number; c2: number; header: boolean; totals: boolean; columns: string[] } | null;
  tableAt(sheet: number, row: number, col: number): string | null;
  spillRange(sheet: number, row: number, col: number): RangeRef | null;
  rowHidden(sheet: number, row: number): boolean;
  rowFiltered(sheet: number, row: number): boolean;
  isSubtotalCell(sheet: number, row: number, col: number): boolean;
  numberFormat(sheet: number, row: number, col: number): string;
  readonly date1904: boolean;
}

interface Frame {
  sheet: number;
  row: number;
  col: number;
  vars: Map<string, Value> | null;
  /** Names currently being expanded, to stop a name that refers to itself. */
  names: Set<string>;
  depth: number;
}

const OMITTED = Symbol('omitted');
type Omitted = { kind: 'ref'; [OMITTED]: true };
const omittedMarker = new CellError('#VALUE!', 'omitted');

const lower = (s: string) => s.toLowerCase();

export class Evaluator {
  constructor(private readonly host: Host) {}

  evaluateFormula(formula: string, sheet: number, row: number, col: number): Value {
    let node: Node;
    try { node = parseFormula(formula); } catch (e) { if (e instanceof FormulaError) return Err.name; throw e; }
    const f: Frame = { sheet, row, col, vars: null, names: new Set(), depth: 0 };
    try {
      return this.eval(node, f);
    } catch (e) {
      if (e instanceof RangeError) return new CellError('#NUM!', 'Too deeply nested');
      throw e;
    }
  }

  /** Turns a result into what a cell holds: refs become values, blanks in results become 0. */
  finalize(v: Value): Scalar | Matrix {
    if (isRef(v)) {
      if (v.r1 === v.r2 && v.c1 === v.c2) return blankToZero(this.host.getValue(v.sheet, v.r1, v.c1));
      return this.values(v, false).map((r) => r.map(blankToZero));
    }
    if (isRefs(v)) return v.areas.length === 1 ? this.finalize(v.areas[0]!) : Err.value;
    if (isLambda(v)) return Err.calc;
    if (isMatrix(v)) {
      if (v.length === 1 && v[0]!.length === 1) return blankToZero(v[0]![0]!);
      return v.map((r) => r.map(blankToZero));
    }
    return v;
  }

  ctxFor(f: Frame): FnContext {
    const host = this.host;
    return {
      sheet: f.sheet, row: f.row, col: f.col, date1904: host.date1904,
      values: (r) => this.values(r, true),
      cell: (s, r, c) => host.getValue(s, r, c),
      usedBounds: (s) => host.usedBounds(s),
      formula: (s, r, c) => host.formula(s, r, c),
      sheetIndex: (n) => host.sheetIndex(n),
      sheetName: (i) => host.sheetName(i),
      sheetCount: () => host.sheetCount(),
      call: (fn, args) => fn.invoke(args),
      rowHidden: (s, r) => host.rowHidden(s, r),
      rowFiltered: (s, r) => host.rowFiltered(s, r),
      isSubtotalCell: (s, r, c) => host.isSubtotalCell(s, r, c),
      numberFormat: (s, r, c) => host.numberFormat(s, r, c),
    };
  }

  /** Values of a range. Trimmed to the used area when `trim` is set, keeping at least one cell. */
  values(r: RangeRef, trim: boolean): Matrix {
    let { r2, c2 } = r;
    if (trim) {
      const b = this.host.usedBounds(r.sheet);
      r2 = Math.max(r.r1, Math.min(r2, b.rows));
      c2 = Math.max(r.c1, Math.min(c2, b.cols));
    }
    const out: Matrix = [];
    for (let i = r.r1; i <= r2; i++) {
      const row: Scalar[] = new Array(c2 - r.c1 + 1);
      for (let j = r.c1; j <= c2; j++) row[j - r.c1] = this.host.getValue(r.sheet, i, j);
      out.push(row);
    }
    return out;
  }

  private resolveRef(n: RefNode, f: Frame): Value {
    if (n.external) return Err.ref;
    const s1 = n.sheet === undefined ? f.sheet : this.host.sheetIndex(n.sheet);
    if (s1 < 0) return Err.ref;
    if (n.sheet2 !== undefined) {
      const s2 = this.host.sheetIndex(n.sheet2);
      if (s2 < 0) return Err.ref;
      const areas: RangeRef[] = [];
      for (let s = Math.min(s1, s2); s <= Math.max(s1, s2); s++) areas.push(ref(s, n.r1, n.c1, n.r2, n.c2));
      return { kind: 'refs', areas };
    }
    return ref(s1, n.r1, n.c1, n.r2, n.c2);
  }

  eval(n: Node, f: Frame): Value {
    switch (n.type) {
      case 'num': return n.value;
      case 'str': return n.value;
      case 'bool': return n.value;
      case 'err': return new CellError(n.code);
      case 'missing': return null;
      case 'array': return n.rows.map((r) => r.map((x) => { const v = this.eval(x, f); return isMatrix(v) ? v[0]![0]! : (v as Scalar); }));
      case 'ref': return this.resolveRef(n.ref, f);
      case 'name': return this.name(n.name, n.sheet, f);
      case 'table': return this.table(n, f);
      case 'union': {
        const areas: RangeRef[] = [];
        for (const it of n.items) {
          const v = this.eval(it, f);
          if (isRef(v)) areas.push(v);
          else if (isRefs(v)) areas.push(...v.areas);
          else return isError(v) ? v : Err.value;
        }
        return { kind: 'refs', areas };
      }
      case 'spill': {
        const v = this.eval(n.arg, f);
        if (!isRef(v)) return Err.ref;
        return this.host.spillRange(v.sheet, v.r1, v.c1) ?? Err.ref;
      }
      case 'unary': return this.unary(n.op, this.eval(n.arg, f), f);
      case 'binary': return this.binary(n.op, n.left, n.right, f);
      case 'invoke': {
        const callee = this.eval(n.callee, f);
        if (!isLambda(callee)) return isError(callee) ? callee : Err.value;
        return callee.invoke(n.args.map((a) => this.evalArg(a, f)));
      }
      case 'call': return this.call(n.name, n.args, f);
    }
  }

  private evalArg(a: Node, f: Frame): Value {
    return a.type === 'missing' ? (omittedMarker as unknown as Value) : this.eval(a, f);
  }

  private name(name: string, sheetName: string | undefined, f: Frame): Value {
    const key = lower(name);
    if (!sheetName && f.vars?.has(key)) return f.vars.get(key)!;
    const sheet = sheetName ? this.host.sheetIndex(sheetName) : f.sheet;
    if (sheet < 0) return Err.ref;
    const formula = this.host.definedName(name, sheet);
    if (formula !== null) {
      if (f.names.has(key) || f.depth > 64) return Err.ref;
      let node: Node;
      try { node = parseFormula(formula); } catch { return Err.name; }
      f.names.add(key);
      try { return this.eval(node, { ...f, vars: null, depth: f.depth + 1 }); } finally { f.names.delete(key); }
    }
    const t = this.host.table(name);
    if (t) return ref(t.sheet, t.r1 + (t.header ? 1 : 0), t.c1, t.r2 - (t.totals ? 1 : 0), t.c2);
    return new CellError('#NAME?', name);
  }

  private table(n: Extract<Node, { type: 'table' }>, f: Frame): Value {
    const tname = n.table || this.host.tableAt(f.sheet, f.row, f.col);
    if (!tname) return Err.ref;
    const t = this.host.table(tname);
    if (!t) return new CellError('#NAME?', tname);
    const items = n.spec.items.length ? n.spec.items : ['#Data'];
    const dataTop = t.r1 + (t.header ? 1 : 0);
    const dataBottom = t.r2 - (t.totals ? 1 : 0);
    let r1 = Infinity, r2 = -Infinity;
    for (const it of items) {
      if (it === '#All') { r1 = Math.min(r1, t.r1); r2 = Math.max(r2, t.r2); }
      else if (it === '#Data') { r1 = Math.min(r1, dataTop); r2 = Math.max(r2, dataBottom); }
      else if (it === '#Headers') { if (!t.header) return Err.ref; r1 = Math.min(r1, t.r1); r2 = Math.max(r2, t.r1); }
      else if (it === '#Totals') { if (!t.totals) return Err.ref; r1 = Math.min(r1, t.r2); r2 = Math.max(r2, t.r2); }
      else {
        if (f.sheet !== t.sheet || f.row < dataTop || f.row > dataBottom) return Err.value;
        r1 = Math.min(r1, f.row); r2 = Math.max(r2, f.row);
      }
    }
    let c1 = t.c1, c2 = t.c2;
    if (n.spec.col1) {
      const i = t.columns.findIndex((c) => lower(c) === lower(n.spec.col1!));
      if (i < 0) return Err.ref;
      c1 = c2 = t.c1 + i;
      if (n.spec.col2) {
        const j = t.columns.findIndex((c) => lower(c) === lower(n.spec.col2!));
        if (j < 0) return Err.ref;
        c1 = t.c1 + Math.min(i, j); c2 = t.c1 + Math.max(i, j);
      }
    }
    if (r1 > r2) return Err.ref;
    return ref(t.sheet, r1, c1, r2, c2);
  }

  // ---------- Operators ----------

  private unary(op: '-' | '+' | '%' | '@', v: Value, f: Frame): Value {
    if (op === '@') {
      if (isRef(v)) return this.intersectCell(v, f);
      if (isMatrix(v)) return v[0]![0]!;
      return v;
    }
    if (op === '+') return v;
    return this.lift1(v, f, (x) => {
      const n = toNumber(x, this.host.date1904);
      if (isError(n)) return n;
      return op === '-' ? -n : n / 100;
    });
  }

  /** Implicit intersection: the cell of a range on the formula's own row or column. */
  private intersectCell(r: RangeRef, f: Frame): Scalar {
    if (r.r1 === r.r2 && r.c1 === r.c2) return this.host.getValue(r.sheet, r.r1, r.c1);
    if (r.c1 === r.c2 && f.row >= r.r1 && f.row <= r.r2) return this.host.getValue(r.sheet, f.row, r.c1);
    if (r.r1 === r.r2 && f.col >= r.c1 && f.col <= r.c2) return this.host.getValue(r.sheet, r.r1, f.col);
    return Err.value;
  }

  private binary(op: string, ln: Node, rn: Node, f: Frame): Value {
    if (op === ':' || op === ' ') {
      const l = this.eval(ln, f), r = this.eval(rn, f);
      if (isError(l)) return l;
      if (isError(r)) return r;
      const a = this.asRef(l), b = this.asRef(r);
      if (!a || !b) return Err.value;
      if (a.sheet !== b.sheet) return Err.value;
      if (op === ':') return ref(a.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
      const r1 = Math.max(a.r1, b.r1), c1 = Math.max(a.c1, b.c1), r2 = Math.min(a.r2, b.r2), c2 = Math.min(a.c2, b.c2);
      return r1 > r2 || c1 > c2 ? Err.null : ref(a.sheet, r1, c1, r2, c2);
    }
    const l = this.eval(ln, f), r = this.eval(rn, f);
    return this.lift2(l, r, f, (a, b) => scalarOp(op, a, b, this.host.date1904));
  }

  private asRef(v: Value): RangeRef | null {
    if (isRef(v)) return v;
    if (isRefs(v) && v.areas.length === 1) return v.areas[0]!;
    return null;
  }

  /** A value as a matrix for element-wise work; single cells come back as scalars. */
  private arrayOf(v: Value, f: Frame): Scalar | Matrix {
    if (isRef(v)) return v.r1 === v.r2 && v.c1 === v.c2 ? this.host.getValue(v.sheet, v.r1, v.c1) : this.values(v, false);
    if (isRefs(v)) return v.areas.length === 1 ? this.arrayOf(v.areas[0]!, f) : Err.value;
    if (isLambda(v)) return Err.calc;
    if (isMatrix(v) && v.length === 1 && v[0]!.length === 1) return v[0]![0]!;
    return v;
  }

  private lift1(v: Value, f: Frame, op: (x: Scalar) => Scalar): Value {
    const a = this.arrayOf(v, f);
    if (!isMatrix(a)) return op(a);
    return a.map((row) => row.map(op));
  }

  private lift2(l: Value, r: Value, f: Frame, op: (a: Scalar, b: Scalar) => Scalar): Value {
    const a = this.arrayOf(l, f), b = this.arrayOf(r, f);
    if (!isMatrix(a) && !isMatrix(b)) return op(a, b);
    return broadcast([a, b], (xs) => op(xs[0]!, xs[1]!));
  }

  // ---------- Function calls ----------

  private call(name: string, args: Node[], f: Frame): Value {
    switch (name) {
      case 'IF': return this.fnIf(args, f);
      case 'IFS': {
        if (args.length < 2 || args.length % 2) return Err.na;
        for (let i = 0; i < args.length; i += 2) {
          const c = this.condition(this.eval(args[i]!, f), f);
          if (isError(c)) return c;
          if (isMatrix(c)) return Err.value;
          if (c) return this.eval(args[i + 1]!, f);
        }
        return Err.na;
      }
      case 'IFERROR': case 'IFNA': {
        if (args.length !== 2) return Err.na;
        const v = this.eval(args[0]!, f);
        const test = (x: Scalar) => isError(x) && (name === 'IFERROR' || x.code === '#N/A');
        const a = this.arrayOf(v, f);
        if (isMatrix(a)) {
          if (!a.some((r) => r.some(test))) return v;
          const alt = this.eval(args[1]!, f);
          const b = this.arrayOf(alt, f);
          return broadcast([a, b], ([x, y]) => (test(x!) ? (y === null ? 0 : y!) : x!));
        }
        if (!test(a)) return v;
        const alt = this.eval(args[1]!, f);
        return alt === null ? 0 : alt;
      }
      case 'SWITCH': {
        if (args.length < 3) return Err.na;
        const v = first(this.eval(args[0]!, f), this.ctxFor(f));
        if (isError(v)) return v;
        let i = 1;
        for (; i + 1 < args.length; i += 2) {
          const c = first(this.eval(args[i]!, f), this.ctxFor(f));
          if (isError(c)) return c;
          if (scalarOp('=', v, c, this.host.date1904) === true) return this.eval(args[i + 1]!, f);
        }
        return i < args.length ? this.eval(args[i]!, f) : Err.na;
      }
      case 'CHOOSE': {
        if (args.length < 2) return Err.na;
        const idx = this.arrayOf(this.eval(args[0]!, f), f);
        const pick = (x: Scalar): Value => {
          const n = toNumber(x, this.host.date1904);
          if (isError(n)) return n;
          const k = Math.trunc(n);
          if (k < 1 || k >= args.length) return Err.value;
          return this.eval(args[k]!, f);
        };
        if (!isMatrix(idx)) return pick(idx);
        return idx.map((r) => r.map((x) => { const v = this.arrayOf(pick(x), f); return isMatrix(v) ? v[0]![0]! : v; }));
      }
      case 'LET': {
        if (args.length < 3 || args.length % 2 === 0) return Err.na;
        const vars = new Map(f.vars ?? []);
        const g: Frame = { ...f, vars };
        for (let i = 0; i + 1 < args.length; i += 2) {
          const nm = args[i]!;
          if (nm.type !== 'name') return Err.name;
          vars.set(lower(nm.name), this.eval(args[i + 1]!, g));
        }
        return this.eval(args.at(-1)!, g);
      }
      case 'LAMBDA': return this.lambda(args, f);
      case 'ISOMITTED': {
        const a = args[0];
        if (!a || a.type !== 'name') return false;
        const v = f.vars?.get(lower(a.name));
        return v === (omittedMarker as unknown as Value);
      }
      case 'INDIRECT': return this.indirect(args, f);
      default: break;
    }
    const lambdaVar = f.vars?.get(lower(name));
    if (lambdaVar !== undefined && isLambda(lambdaVar)) return lambdaVar.invoke(args.map((a) => this.evalArg(a, f)));
    const def = getFunction(name);
    if (!def) {
      const named = this.host.definedName(name, f.sheet);
      if (named !== null) {
        const v = this.name(name, undefined, f);
        if (isLambda(v)) return v.invoke(args.map((a) => this.evalArg(a, f)));
      }
      return new CellError('#NAME?', name);
    }
    if (args.length < def.min || args.length > def.max) return new CellError('#N/A', 'Wrong number of arguments');
    const vals = args.map((a) => {
      const v = this.evalArg(a, f);
      return v === (omittedMarker as unknown as Value) ? null : v;
    });
    // A function name passed as a value (eta-reduced lambda), e.g. =BYROW(A1:C3, SUM)
    return this.apply(def, vals, f);
  }

  apply(def: FunctionDef, vals: Value[], f: Frame): Value {
    const ctx = this.ctxFor(f);
    const lifted = def.scalar === true ? vals.map((_, i) => i) : def.scalar ? def.scalar.filter((i) => i < vals.length) : [];
    if (!lifted.length) return guard(def.fn(vals, ctx));
    const prepared = vals.slice();
    let arrays = false;
    for (const i of lifted) {
      const a = this.arrayOf(vals[i]!, f);
      prepared[i] = a;
      if (isMatrix(a)) arrays = true;
    }
    if (!arrays) return guard(def.fn(prepared, ctx));
    const parts = lifted.map((i) => prepared[i] as Scalar | Matrix);
    return broadcast(parts, (xs) => {
      const call = prepared.slice();
      lifted.forEach((i, k) => { call[i] = xs[k]!; });
      const r = guard(def.fn(call, ctx));
      if (isRef(r)) return this.host.getValue(r.sheet, r.r1, r.c1);
      if (isMatrix(r)) return r[0]?.[0] ?? null;
      return isLambda(r) || isRefs(r) ? Err.calc : r;
    });
  }

  private fnIf(args: Node[], f: Frame): Value {
    if (args.length < 1 || args.length > 3) return Err.na;
    const condV = this.eval(args[0]!, f);
    const c = this.arrayOf(condV, f);
    const branch = (i: number): Value => {
      const a = args[i];
      if (!a) return i === 1 ? true : false;
      if (a.type === 'missing') return 0;
      return this.eval(a, f);
    };
    if (!isMatrix(c)) {
      const b = toCond(c);
      if (isError(b)) return b;
      return branch(b ? 1 : 2);
    }
    const t = this.arrayOf(branch(1), f), e = this.arrayOf(branch(2), f);
    return broadcast([c, t, e], ([x, y, z]) => {
      const b = toCond(x!);
      if (isError(b)) return b;
      return b ? y! : z!;
    });
  }

  private condition(v: Value, f: Frame): boolean | CellError | Matrix {
    const c = this.arrayOf(v, f);
    if (isMatrix(c)) return c;
    return toCond(c);
  }

  private lambda(args: Node[], f: Frame): Value {
    if (!args.length) return Err.value;
    const params: string[] = [];
    for (const p of args.slice(0, -1)) {
      if (p.type !== 'name') return Err.value;
      params.push(lower(p.name.replace(/^_xlpm\./i, '')));
    }
    const body = args.at(-1)!;
    const captured = f.vars;
    const lam: LambdaValue = {
      kind: 'lambda',
      params,
      invoke: (vals: Value[]) => {
        if (vals.length > params.length) return Err.value;
        const vars = new Map(captured ?? []);
        params.forEach((p, i) => vars.set(p, i < vals.length ? vals[i]! : (omittedMarker as unknown as Value)));
        if (f.depth > 900) return new CellError('#NUM!', 'Too deeply nested');
        return this.eval(body, { ...f, vars, depth: f.depth + 1 });
      },
    };
    return lam;
  }

  private indirect(args: Node[], f: Frame): Value {
    if (args.length < 1 || args.length > 2) return Err.na;
    const ctx = this.ctxFor(f);
    const t = toText(first(this.eval(args[0]!, f), ctx));
    if (isError(t)) return t;
    const a1 = args[1] ? toBool(first(this.eval(args[1], f), ctx)) : true;
    if (isError(a1)) return a1;
    const text = t.trim();
    if (!text) return Err.ref;
    if (!a1) return this.r1c1(text, f);
    try {
      const node = parseFormula(text);
      if (node.type === 'ref' || node.type === 'name' || node.type === 'table' || (node.type === 'binary' && node.op === ':')) {
        const v = this.eval(node, f);
        return isRef(v) || isRefs(v) ? v : Err.ref;
      }
    } catch { /* not a reference */ }
    return Err.ref;
  }

  private r1c1(text: string, f: Frame): Value {
    const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(.*)$/.exec(text);
    if (!m) return Err.ref;
    const sheetName = m[1]?.replace(/''/g, "'") ?? m[2];
    const sheet = sheetName ? this.host.sheetIndex(sheetName) : f.sheet;
    if (sheet < 0) return Err.ref;
    const part = (p: string): { row: number; col: number } | null => {
      const q = /^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i.exec(p);
      if (!q) return null;
      const v = (s: string | undefined, base: number) => (s === undefined ? base : s.startsWith('[') ? base + +s.slice(1, -1) : +s - 1);
      return { row: v(q[1], f.row), col: v(q[2], f.col) };
    };
    const [a, b] = m[3]!.split(':');
    const p = part(a!), q = b ? part(b) : p;
    if (!p || !q) return Err.ref;
    if (p.row < 0 || q.row < 0 || p.col < 0 || q.col < 0 || p.row >= MAX_ROWS || q.row >= MAX_ROWS || p.col >= MAX_COLS || q.col >= MAX_COLS) return Err.ref;
    return ref(sheet, p.row, p.col, q.row, q.col);
  }
}

const blankToZero = (v: Scalar): Scalar => (v === null ? 0 : v);

const guard = (v: Value): Value => {
  if (typeof v === 'number' && !Number.isFinite(v)) return Err.num;
  if (v === (omittedMarker as unknown as Value)) return null;
  return v;
};

const toCond = (x: Scalar): boolean | CellError => {
  if (isError(x)) return x;
  if (typeof x === 'string') {
    const u = x.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return Err.value;
  }
  return toBool(x);
};

/** Element-wise over several scalars/arrays, broadcasting single rows/columns; mismatched cells are #N/A. */
export const broadcast = (parts: (Scalar | Matrix)[], op: (xs: Scalar[]) => Scalar): Matrix => {
  let rows = 1, cols = 1;
  for (const p of parts) if (isMatrix(p)) { rows = Math.max(rows, p.length); cols = Math.max(cols, p[0]!.length); }
  const out: Matrix = [];
  for (let i = 0; i < rows; i++) {
    const row: Scalar[] = new Array(cols);
    for (let j = 0; j < cols; j++) {
      let bad = false;
      const xs = parts.map((p) => {
        if (!isMatrix(p)) return p;
        const h = p.length, w = p[0]!.length;
        const ii = h === 1 ? 0 : i, jj = w === 1 ? 0 : j;
        if (ii >= h || jj >= w) { bad = true; return null; }
        return p[ii]![jj]!;
      });
      row[j] = bad ? Err.na : op(xs);
    }
    out.push(row);
  }
  return out;
};

/** Arithmetic, text joining and comparison on two values, as spreadsheets do it. */
export const scalarOp = (op: string, a: Scalar, b: Scalar, date1904: boolean): Scalar => {
  if (op === '&') {
    if (isError(a)) return a;
    if (isError(b)) return b;
    const s = (toText(a) as string) + (toText(b) as string);
    return s.length > 32767 ? Err.value : s;
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    if (isError(a)) return a;
    if (isError(b)) return b;
    const c = compareForOp(a, b);
    switch (op) {
      case '=': return c === 0;
      case '<>': return c !== 0;
      case '<': return c < 0;
      case '>': return c > 0;
      case '<=': return c <= 0;
      default: return c >= 0;
    }
  }
  const x = toNumber(a, date1904);
  if (isError(x)) return x;
  const y = toNumber(b, date1904);
  if (isError(y)) return y;
  let r: number;
  switch (op) {
    case '+': r = x + y; break;
    case '-': r = x - y; break;
    case '*': r = x * y; break;
    case '/': if (y === 0) return Err.div0; r = x / y; break;
    case '^':
      if (x === 0 && y === 0) return Err.num;
      if (x === 0 && y < 0) return Err.div0;
      r = Math.pow(x, y);
      if (Number.isNaN(r)) return Err.num;
      break;
    default: return Err.value;
  }
  if (!Number.isFinite(r)) return Err.num;
  // Sums and differences that land within rounding noise of zero show as 0, as in other spreadsheets.
  if ((op === '+' || op === '-') && r !== 0 && Math.abs(r) < Math.max(Math.abs(x), Math.abs(y)) * 2.2e-16 * 4) r = 0;
  return r;
};

const compareForOp = (a: Scalar, b: Scalar): number => {
  // A blank compares as 0 against numbers, "" against text, FALSE against booleans.
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === 'number' && typeof b === 'number') {
    if (a === b) return 0;
    const pa = Number(a.toPrecision(15)), pb = Number(b.toPrecision(15));
    return pa === pb ? 0 : pa < pb ? -1 : 1;
  }
  return compareValues(a, b);
};

export { numberToText };
export type { Omitted };
