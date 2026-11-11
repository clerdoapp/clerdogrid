// GROUPBY, PIVOTBY, PERCENTOF and GETPIVOTDATA.
import { collectNumbers, compareValues, first, int, toMatrix } from '../coerce';
import { getFunction } from '../registry';
import {
  CellError, Err, MAX_COLS, MAX_ROWS, isError, isLambda, isMatrix, isRef, isRefs, isScalar,
  type FnContext, type FunctionDef, type LambdaValue, type Matrix, type Scalar, type Value,
} from '../types';

type R<T> = T | CellError;

const arr = (v: Value, ctx: FnContext): R<Matrix> => {
  if (!isRef(v)) return toMatrix(v, ctx);
  const m = ctx.values(v);
  const rows = v.r1 === 0 && v.r2 === MAX_ROWS - 1 ? m.length : v.r2 - v.r1 + 1;
  const cols = v.c1 === 0 && v.c2 === MAX_COLS - 1 ? (m[0]?.length ?? 0) : v.c2 - v.c1 + 1;
  if (m.length === rows && (m[0]?.length ?? 0) === cols) return m;
  return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => m[i]?.[j] ?? null));
};

const keyOf = (v: Scalar): string =>
  v === null ? 'b' : typeof v === 'number' ? `n${v}` : typeof v === 'string' ? `s${v.toLowerCase()}` : typeof v === 'boolean' ? `l${v}` : `e${v.code}`;

const sortCmp = (x: Scalar, y: Scalar, desc: boolean): number => {
  if (x === null || y === null) return x === y ? 0 : x === null ? 1 : -1;
  if (isError(x) || isError(y)) return isError(x) === isError(y) ? 0 : isError(x) ? 1 : -1;
  const c = compareValues(x, y);
  return desc ? -c : c;
};

/** One scalar from a LAMBDA result; arrays are #CALC! (a nested array). */
const one = (v: Value, ctx: FnContext): Scalar => {
  if (isRef(v)) return v.r1 === v.r2 && v.c1 === v.c2 ? ctx.cell(v.sheet, v.r1, v.c1) : Err.calc;
  if (isMatrix(v)) return v.length === 1 && v[0]!.length === 1 ? v[0]![0]! : Err.calc;
  return isScalar(v) ? v : Err.calc;
};

// ---------------------------------------------------------------- PERCENTOF

const PERCENTOF: FunctionDef = {
  name: 'PERCENTOF', min: 2, max: 2,
  fn: (a, ctx) => {
    const s = collectNumbers([a[0]!], ctx), t = collectNumbers([a[1]!], ctx);
    if (isError(s)) return s; if (isError(t)) return t;
    const tot = t.reduce((x, y) => x + y, 0);
    return tot === 0 ? Err.div0 : s.reduce((x, y) => x + y, 0) / tot;
  },
};

// ---------------------------------------------------------------- shared grouping

interface GNode { key: Scalar[]; rows: number[]; kids: GNode[]; parent: GNode | null }
interface Line { keys: Scalar[]; rows: number[]; kind: 'leaf' | 'sub' | 'total'; node: GNode | null; parentRows: number[] }
interface SortSpec { idx: number; desc: boolean }

/** A function given to GROUPBY/PIVOTBY: one LAMBDA, a function name, or several stacked with HSTACK/VSTACK. */
interface Agg { f: LambdaValue; label: string; two: boolean }

const aggsOf = (v: Value | undefined): R<{ list: Agg[]; vertical: boolean }> => {
  if (v === undefined || !isLambda(v)) return Err.value;
  const list = (v.list ?? [v]).map((f): Agg => {
    const def = f.label ? getFunction(f.label) : undefined;
    const two = f.params.length === 2 || (!!def && def.min === 2);
    return { f, label: f.label ?? (f.params.length === 2 ? 'PERCENTOF' : 'LAMBDA'), two };
  });
  if (list.some((x) => x.f.params.length > 2)) return Err.value;
  return { list, vertical: !!v.list && !!v.vertical && v.list.length > 1 };
};

/**
 * Groups `rows` level by level. In hierarchy mode each level groups one field column; in table mode
 * one level groups the whole tuple. `cmp` orders siblings at a level.
 */
const buildTree = (keys: Matrix, rows: number[], level: number, nf: number, table: boolean, cmp: (level: number, x: GNode, y: GNode) => number, parent: GNode | null): GNode[] => {
  const map = new Map<string, GNode>();
  for (const r of rows) {
    const key = table ? keys[r]! : [keys[r]![level]!];
    const k = key.map(keyOf).join('\u0000');
    let n = map.get(k);
    if (!n) map.set(k, (n = { key, rows: [], kids: [], parent }));
    n.rows.push(r);
  }
  const nodes = [...map.values()];
  if (!table && level + 1 < nf) for (const n of nodes) n.kids = buildTree(keys, n.rows, level + 1, nf, table, cmp, n);
  return nodes.sort((x, y) => cmp(level, x, y));
};

/** Flattens a group tree into lines: subtotals for levels above |depth|-1, and a grand total when depth is not 0. */
const flatten = (tree: GNode[], all: number[], nf: number, depth: number): Line[] => {
  const top = depth < 0, d = Math.abs(depth);
  const out: Line[] = [];
  const walk = (nodes: GNode[], prefix: Scalar[], level: number): void => {
    for (const n of nodes) {
      const keys = [...prefix, ...n.key];
      const parentRows = n.parent?.rows ?? all;
      if (!n.kids.length) { out.push({ keys, rows: n.rows, kind: 'leaf', node: n, parentRows }); continue; }
      const sub: Line = { keys: [...keys, ...Array<Scalar>(nf - keys.length).fill(null)], rows: n.rows, kind: 'sub', node: n, parentRows };
      const withSub = level < d - 1;
      if (withSub && top) out.push(sub);
      walk(n.kids, keys, level + 1);
      if (withSub && !top) out.push(sub);
    }
  };
  walk(tree, [], 0);
  if (d >= 1) {
    const tot: Line = { keys: ['Total', ...Array<Scalar>(Math.max(0, nf - 1)).fill(null)], rows: all, kind: 'total', node: null, parentRows: all };
    if (top) out.unshift(tot); else out.push(tot);
  }
  return out;
};

interface Fields { names: Scalar[]; data: Matrix }

/** Splits headers off (field_headers 1/3, or auto-detected) and reports whether to show them. */
const headerMode = (fh: number | null, vals: Matrix, multi: boolean): R<{ has: boolean; show: boolean }> => {
  if (fh === null) {
    const has = typeof vals[0]?.[0] === 'string' && typeof vals[1]?.[0] === 'number';
    return { has, show: has && multi };
  }
  if (![0, 1, 2, 3].includes(fh)) return Err.value;
  return { has: fh === 1 || fh === 3, show: fh >= 2 };
};

const split = (m: Matrix, has: boolean, prefix: string): Fields => ({
  names: has ? m[0]!.map((v, j) => (v === null || v === '' ? `${prefix} ${j + 1}` : v)) : m[0]!.map((_, j) => `${prefix} ${j + 1}`),
  data: has ? m.slice(1) : m,
});

const nums = (v: Value, ctx: FnContext): R<number[]> => {
  const m = arr(v, ctx);
  if (isError(m)) return m;
  const out: number[] = [];
  for (const c of m.flat()) {
    if (c === null) continue;
    if (typeof c !== 'number') return isError(c) ? c : Err.value;
    out.push(Math.trunc(c));
  }
  return out;
};

const parseSort = (v: Value | undefined, ctx: FnContext, width: number): R<SortSpec[]> => {
  if (v === undefined || v === null) return [];
  const ns = nums(v, ctx);
  if (isError(ns)) return ns;
  if (ns.some((n) => n === 0 || Math.abs(n) > width)) return Err.value;
  return ns.map((n) => ({ idx: Math.abs(n) - 1, desc: n < 0 }));
};

/** Calls an aggregation over some rows of one value column; two-argument functions also get the base rows. */
const runAgg = (a: Agg, ctx: FnContext, vals: Matrix, rows: number[], col: number, base: number[]): Scalar => {
  if (!rows.length) return null;
  const pick = (rs: number[]): Matrix => (rs.length ? rs.map((r) => [vals[r]![col]!]) : [[null]]);
  return one(ctx.call(a.f, a.two ? [pick(rows), pick(base)] : [pick(rows)]), ctx);
};

const sorter = (specs: SortSpec[], nf: number, table: boolean, agg: (rows: number[], valueCol: number) => Scalar) =>
  (level: number, x: GNode, y: GNode): number => {
    for (const s of specs) {
      let c = 0;
      if (s.idx >= nf) c = sortCmp(agg(x.rows, s.idx - nf), agg(y.rows, s.idx - nf), s.desc);
      else if (table) c = sortCmp(x.key[s.idx]!, y.key[s.idx]!, s.desc);
      else if (s.idx === level) c = sortCmp(x.key[0]!, y.key[0]!, s.desc);
      else continue;
      if (c) return c;
    }
    for (let i = 0; i < x.key.length; i++) { const c = sortCmp(x.key[i]!, y.key[i]!, false); if (c) return c; }
    return 0;
  };

const filterRows = (v: Value | undefined, ctx: FnContext, n: number, has: boolean): R<number[]> => {
  const idx = Array.from({ length: n }, (_, i) => i);
  if (v === undefined || v === null) return idx;
  const m = arr(v, ctx);
  if (isError(m)) return m;
  if (m.length !== 1 && m[0]!.length !== 1) return Err.value;
  let cells = m.flat();
  if (has && cells.length === n + 1) cells = cells.slice(1);
  if (cells.length !== n) return Err.value;
  const out: number[] = [];
  for (let i = 0; i < n; i++) {
    const c = cells[i]!;
    if (isError(c)) return c;
    if (c === true || (typeof c === 'number' && c !== 0)) out.push(i);
  }
  return out;
};

const optInt = (v: Value | undefined, ctx: FnContext): R<number | null> => (v === undefined || v === null ? null : int(v, ctx));

/** Empty places in the result show as empty, not as 0. */
const blanks = (m: Matrix): Matrix => m.map((r) => r.map((v) => (v === null ? '' : v)));

const vectorsOk = (...ms: Matrix[]): boolean => ms.every((m) => m.length === ms[0]!.length);

// ---------------------------------------------------------------- GROUPBY

const GROUPBY: FunctionDef = {
  name: 'GROUPBY', min: 3, max: 8,
  fn: (a, ctx) => {
    const rf0 = arr(a[0]!, ctx), v0 = arr(a[1]!, ctx);
    if (isError(rf0)) return rf0; if (isError(v0)) return v0;
    if (!vectorsOk(rf0, v0)) return Err.value;
    const ag = aggsOf(a[2]);
    if (isError(ag)) return ag;
    const fh = optInt(a[3], ctx), td = optInt(a[4], ctx), rel = optInt(a[7], ctx);
    if (isError(fh)) return fh; if (isError(td)) return td; if (isError(rel)) return rel;
    if (rel !== null && rel !== 0 && rel !== 1) return Err.value;
    const nf = rf0[0]!.length, nv = v0[0]!.length, nfn = ag.list.length;
    const hm = headerMode(fh, v0, nf > 1 || nv > 1);
    if (isError(hm)) return hm;
    const rf = split(rf0, hm.has, 'Row Field'), vals = split(v0, hm.has, 'Value');
    const table = rel === 1;
    const depth = td ?? (table ? 1 : nf);
    if (table && Math.abs(depth) > 1) return Err.value;
    const horiz = ag.vertical ? 1 : nfn;
    const specs = parseSort(a[5], ctx, nf + nv * horiz);
    if (isError(specs)) return specs;
    const rows = filterRows(a[6], ctx, rf.data.length, hm.has);
    if (isError(rows)) return rows;
    if (!rows.length) return Err.calc;
    const sortAgg = (rs: number[], k: number) => runAgg(ag.list[ag.vertical ? 0 : Math.floor(k / nv)]!, ctx, vals.data, rs, k % nv, rows);
    const tree = buildTree(rf.data, rows, 0, nf, table, sorter(specs, nf, table, sortAgg), null);
    const lines = flatten(tree, rows, nf, depth);
    const out: Matrix = [];
    const blankKeys = Array<Scalar>(nf).fill(null);
    if (ag.vertical) {
      if (hm.show) out.push([...rf.names, null, ...vals.names]);
      for (const l of lines) ag.list.forEach((fn, fi) => out.push([...(fi === 0 ? l.keys : blankKeys), fn.label, ...vals.names.map((_, c) => runAgg(fn, ctx, vals.data, l.rows, c, rows))]));
      return blanks(out);
    }
    if (nfn > 1) out.push([...(hm.show && nv === 1 ? rf.names : blankKeys), ...ag.list.flatMap((fn) => vals.names.map(() => fn.label))]);
    if (hm.show && (nfn === 1 || nv > 1)) out.push([...rf.names, ...ag.list.flatMap(() => vals.names)]);
    for (const l of lines) out.push([...l.keys, ...ag.list.flatMap((fn) => vals.names.map((_, c) => runAgg(fn, ctx, vals.data, l.rows, c, rows)))]);
    return blanks(out);
  },
};

// ---------------------------------------------------------------- PIVOTBY

const PIVOTBY: FunctionDef = {
  name: 'PIVOTBY', min: 4, max: 11,
  fn: (a, ctx) => {
    const rf0 = arr(a[0]!, ctx), cf0 = arr(a[1]!, ctx), v0 = arr(a[2]!, ctx);
    if (isError(rf0)) return rf0; if (isError(cf0)) return cf0; if (isError(v0)) return v0;
    if (!vectorsOk(rf0, cf0, v0)) return Err.value;
    const ag = aggsOf(a[3]);
    if (isError(ag)) return ag;
    const ints = [a[4], a[5], a[7], a[10]].map((v) => optInt(v, ctx));
    for (const x of ints) if (isError(x)) return x;
    const [fh, rtd, ctd, relTo] = ints as (number | null)[];
    if (relTo !== null && (relTo! < 0 || relTo! > 4)) return Err.value;
    const nr = rf0[0]!.length, nc = cf0[0]!.length, nv = v0[0]!.length, nfn = ag.list.length;
    const hm = headerMode(fh, v0, true);
    if (isError(hm)) return hm;
    const rf = split(rf0, hm.has, 'Row Field'), cf = split(cf0, hm.has, 'Col Field'), vals = split(v0, hm.has, 'Value');
    const rows = filterRows(a[9], ctx, rf.data.length, hm.has);
    if (isError(rows)) return rows;
    if (!rows.length) return Err.calc;
    const horiz = ag.vertical ? 1 : nfn;
    const rspecs = parseSort(a[6], ctx, nr + nv * horiz), cspecs = parseSort(a[8], ctx, nc + nv * horiz);
    if (isError(rspecs)) return rspecs; if (isError(cspecs)) return cspecs;
    const sortAgg = (rs: number[], k: number) => runAgg(ag.list[ag.vertical ? 0 : Math.floor(k / nv) % nfn]!, ctx, vals.data, rs, k % nv, rows);
    const rLines = flatten(buildTree(rf.data, rows, 0, nr, false, sorter(rspecs, nr, false, sortAgg), null), rows, nr, rtd ?? nr);
    const cLines = flatten(buildTree(cf.data, rows, 0, nc, false, sorter(cspecs, nc, false, sortAgg), null), rows, nc, ctd ?? nc);
    const cSets = cLines.map((l) => new Set(l.rows));
    const and = (x: number[], s: Set<number>) => x.filter((r) => s.has(r));
    const perCol = horiz * nv;
    const out: Matrix = [];
    const lead = nr + (ag.vertical ? 1 : 0);
    const pad = () => Array<Scalar>(lead).fill(null);
    for (let l = 0; l < nc; l++) out.push([...pad(), ...cLines.flatMap((c) => Array<Scalar>(perCol).fill(c.keys[l]!))]);
    if (!ag.vertical && nfn > 1) out.push([...pad(), ...cLines.flatMap(() => ag.list.flatMap((fn) => vals.names.map(() => fn.label)))]);
    if (nv > 1) out.push([...pad(), ...cLines.flatMap(() => (ag.vertical ? vals.names : ag.list.flatMap(() => vals.names)))]);
    if (hm.show) out[out.length - 1]!.splice(0, nr, ...rf.names);
    for (const rl of rLines) {
      const fns = ag.vertical ? ag.list.map((f) => [f]) : [ag.list];
      fns.forEach((group, gi) => {
        const row: Scalar[] = gi === 0 ? [...rl.keys] : Array<Scalar>(nr).fill(null);
        if (ag.vertical) row.push(group[0]!.label);
        cLines.forEach((cl, ci) => {
          const rs = and(rl.rows, cSets[ci]!);
          // relative_to: 0 column totals, 1 row totals, 2 grand total, 3 parent column total, 4 parent row total.
          let base: number[];
          switch (relTo ?? 0) {
            case 1: base = rl.rows; break;
            case 2: base = rows; break;
            case 3: base = and(rl.rows, new Set(cl.kind === 'total' ? rows : cl.parentRows)); break;
            case 4: base = and(rl.kind === 'total' ? rows : rl.parentRows, cSets[ci]!); break;
            default: base = cl.rows;
          }
          for (const fn of group) for (let v = 0; v < nv; v++) row.push(runAgg(fn, ctx, vals.data, rs, v, base));
        });
        out.push(row);
      });
    }
    return blanks(out);
  },
};

// ---------------------------------------------------------------- GETPIVOTDATA

const GETPIVOTDATA: FunctionDef = {
  name: 'GETPIVOTDATA', min: 2, max: 254, volatile: true,
  fn: (a, ctx) => {
    const df = first(a[0]!, ctx);
    if (isError(df)) return df;
    const t = a[1]!;
    if (isRefs(t)) return Err.ref;
    if (!isRef(t)) return Err.ref;
    if ((a.length - 2) % 2) return Err.ref;
    const pairs: [string, Scalar][] = [];
    for (let i = 2; i + 1 < a.length; i += 2) {
      const f = first(a[i]!, ctx), it = first(a[i + 1]!, ctx);
      if (isError(f)) return f; if (isError(it)) return it;
      pairs.push([String(f ?? ''), it]);
    }
    if (!ctx.pivotData) return Err.ref;
    return ctx.pivotData(t.sheet, t.r1, t.c1, String(df ?? ''), pairs);
  },
};

export const defs: FunctionDef[] = [GROUPBY, PIVOTBY, PERCENTOF, GETPIVOTDATA];
