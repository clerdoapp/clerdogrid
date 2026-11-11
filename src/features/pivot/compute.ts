// PivotTable calculation: filters records, builds the row and column trees, aggregates values,
// applies "show values as", and lays everything out as a grid of cells. Pure: no store access.
import { formatValue } from '../../formats/numfmt';
import { CellError, Err, isError, type Scalar } from '../../engine/types';
import {
  VALUES_FIELD,
  type AxisItem, type DataField, type ItemType, type PivotAgg, type PivotCache, type PivotField, type PivotFilter, type PivotTable, type Subtotal,
} from '../../model/pivot';
import { compareItems, dbIndex, fieldItems, itemKey, recordItems } from './cache';
import { evalPivotFormula, parsePivotFormula, type PNode } from './formula';

// ---------------------------------------------------------------- accumulation

export class Acc {
  recs = 0; n = 0; nums = 0; sum = 0; prod = 1; min = Infinity; max = -Infinity; mean = 0; m2 = 0;
  err: CellError | null = null;
  distinct: Set<string> | null = null;
  add(v: Scalar, distinct = false): void {
    this.recs++;
    if (v === null || v === '') return;
    if (isError(v)) { this.err ??= v; this.n++; return; }
    this.n++;
    if (distinct) (this.distinct ??= new Set()).add(itemKey(v));
    if (typeof v !== 'number') return;
    this.nums++;
    this.sum += v; this.prod *= v;
    if (v < this.min) this.min = v;
    if (v > this.max) this.max = v;
    const d = v - this.mean;
    this.mean += d / this.nums;
    this.m2 += d * (v - this.mean);
  }
  /** Adds an already computed value (from a calculated item) into a total. */
  addValue(v: Scalar): void { this.add(v); }
  merge(o: Acc): void {
    if (!o.recs) return;
    if (o.err) this.err ??= o.err;
    const n = this.nums + o.nums;
    if (o.nums) {
      const d = o.mean - this.mean;
      this.m2 += o.m2 + (d * d * this.nums * o.nums) / n;
      this.mean += (d * o.nums) / n;
    }
    this.recs += o.recs; this.n += o.n; this.nums = n; this.sum += o.sum; this.prod *= o.prod;
    this.min = Math.min(this.min, o.min); this.max = Math.max(this.max, o.max);
    if (o.distinct) { this.distinct ??= new Set(); for (const k of o.distinct) this.distinct.add(k); }
  }
  result(agg: PivotAgg): Scalar {
    if (!this.recs) return null;
    if (this.err && agg !== 'count' && agg !== 'countNums' && agg !== 'distinctCount') return this.err;
    switch (agg) {
      case 'sum': return this.sum;
      case 'count': return this.n;
      case 'countNums': return this.nums;
      case 'distinctCount': return this.distinct?.size ?? 0;
      case 'average': return this.nums ? this.sum / this.nums : Err.div0;
      case 'max': return this.nums ? this.max : 0;
      case 'min': return this.nums ? this.min : 0;
      case 'product': return this.nums ? this.prod : 0;
      case 'stdDev': return this.nums > 1 ? Math.sqrt(this.m2 / (this.nums - 1)) : Err.div0;
      case 'stdDevp': return this.nums ? Math.sqrt(this.m2 / this.nums) : Err.div0;
      case 'var': return this.nums > 1 ? this.m2 / (this.nums - 1) : Err.div0;
      case 'varp': return this.nums ? this.m2 / this.nums : Err.div0;
    }
  }
}

const SUBTOTAL_AGG: Record<Exclude<Subtotal, 'default'>, PivotAgg> = {
  sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums',
  stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp',
};
const SUBTOTAL_TYPE: Record<Subtotal, ItemType> = {
  default: 'default', sum: 'sum', count: 'countA', average: 'avg', max: 'max', min: 'min', product: 'product', countNums: 'count',
  stdDev: 'stdDev', stdDevp: 'stdDevP', var: 'var', varp: 'varP',
};
export const TYPE_SUBTOTAL: Partial<Record<ItemType, Subtotal>> = Object.fromEntries(Object.entries(SUBTOTAL_TYPE).map(([k, v]) => [v, k])) as Partial<Record<ItemType, Subtotal>>;

export const AGG_LABEL: Record<PivotAgg, string> = {
  sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count Numbers',
  stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp', distinctCount: 'Distinct Count',
};

// ---------------------------------------------------------------- trees

export interface PNodeT {
  id: number;
  /** Cache field, or -1 for the root. */
  field: number;
  /** Index into the pivot field's items (display slot), -1 for the root. */
  slot: number;
  depth: number;
  recs: number[];
  kids: PNodeT[];
  parent: PNodeT | null;
  /** Calculated item formula (set on the item and on the inner items under it). */
  calc?: PNode;
  /** The calculated item node this node sits under (itself for the item). */
  calcRoot?: PNodeT;
  key: string;
}

export interface PathEl { field: number; slot: number }

export interface Line {
  kind: 'item' | 'sub' | 'grand' | 'blank';
  node: PNodeT;
  /** Data field the line belongs to when Σ Values is on this axis. */
  data: number;
  path: PathEl[];
  fn?: Subtotal;
  showValues: boolean;
  /** Item line with deeper levels hidden (a +/- button applies). */
  expandable?: 'expanded' | 'collapsed';
}

export type CellRole = 'pageLabel' | 'pageValue' | 'dataCaption' | 'colCaption' | 'rowCaption' | 'colItem' | 'rowItem' | 'data' | 'empty';

export interface PCell {
  v: Scalar;
  role: CellRole;
  numFmt?: string;
  /** Row or column item belongs to a subtotal or grand total line. */
  total?: 'sub' | 'grand';
  /** Indent level in compact form. */
  indent?: number;
  rowLine?: number;
  colLine?: number;
  data?: number;
  /** Axis field the label belongs to (row/col items, captions, page fields). */
  field?: number;
  slot?: number;
  button?: 'expanded' | 'collapsed';
  /** Captions with a filter button: the axis fields they cover. */
  filter?: 'rows' | 'cols' | number;
}

export interface PivotLayout {
  width: number;
  height: number;
  cells: (PCell | null)[][];
  /** Report filter cells above the table: offsets from the anchor (negative rows). */
  pageCells: { dr: number; dc: number; cell: PCell }[];
  pageRows: number;
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
  rowLines: Line[];
  colLines: Line[];
  rowItems: AxisItem[];
  colItems: AxisItem[];
  /** Record indexes behind a value cell (for drill-down). */
  recordsAt: (rowLine: number, colLine: number) => number[];
  /** Raw (before show-values-as) and shown value for a data field at two nodes. */
  valueAt: (d: number, rowNode: PNodeT, colNode: PNodeT, fn?: Subtotal) => Scalar;
  rowRoot: PNodeT;
  colRoot: PNodeT;
  /** Records left after filters. */
  records: number[];
  /** Label of an item in a field. */
  caption: (field: number, slot: number) => string;
  errors: string[];
}

const isTabular = (pt: PivotTable, f: number): boolean => (f === VALUES_FIELD ? pt.layout === 'tabular' : (pt.fields[f]?.outline ?? pt.layout !== 'tabular') === false);
const isCompact = (pt: PivotTable, f: number): boolean => {
  if (f === VALUES_FIELD) return pt.layout === 'compact';
  const pf = pt.fields[f];
  return !isTabular(pt, f) && (pf?.compact ?? pt.layout === 'compact');
};

export const fieldCaption = (pt: PivotTable, cache: PivotCache, f: number): string =>
  f === VALUES_FIELD ? (pt.dataCaption ?? 'Values') : (pt.fields[f]?.name ?? cache.fields[f]?.name ?? '');

export const dataFieldName = (cache: PivotCache, df: DataField): string => df.name || `${AGG_LABEL[df.agg]} of ${cache.fields[df.field]?.name ?? ''}`;

const labelMatch = (f: Extract<PivotFilter, { kind: 'label' }>, caption: string, raw: Scalar): boolean => {
  const t = caption.toLowerCase(), a = f.value1.toLowerCase(), b = (f.value2 ?? '').toLowerCase();
  const wild = (pat: string) => new RegExp(`^${pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?')}$`, 's');
  const num = typeof raw === 'number' ? raw : NaN;
  const na = Number(f.value1), nb = Number(f.value2);
  const cmp = (x: string, n: number) => (!Number.isNaN(num) && !Number.isNaN(n) && x.trim() !== '' ? num - n : t.localeCompare(x, undefined, { sensitivity: 'accent' }));
  switch (f.op) {
    case 'equal': return wild(a).test(t);
    case 'notEqual': return !wild(a).test(t);
    case 'beginsWith': return wild(`${a}*`).test(t);
    case 'notBeginsWith': return !wild(`${a}*`).test(t);
    case 'endsWith': return wild(`*${a}`).test(t);
    case 'notEndsWith': return !wild(`*${a}`).test(t);
    case 'contains': return wild(`*${a}*`).test(t);
    case 'notContains': return !wild(`*${a}*`).test(t);
    case 'greaterThan': return cmp(a, na) > 0;
    case 'greaterThanOrEqual': return cmp(a, na) >= 0;
    case 'lessThan': return cmp(a, na) < 0;
    case 'lessThanOrEqual': return cmp(a, na) <= 0;
    case 'between': return cmp(a, na) >= 0 && cmp(b, nb) <= 0;
    case 'notBetween': return cmp(a, na) < 0 || cmp(b, nb) > 0;
  }
};

const valueMatch = (f: Extract<PivotFilter, { kind: 'value' }>, v: Scalar): boolean => {
  if (typeof v !== 'number') return false;
  const a = f.value1, b = f.value2 ?? f.value1;
  switch (f.op) {
    case 'equal': return v === a;
    case 'notEqual': return v !== a;
    case 'greaterThan': return v > a;
    case 'greaterThanOrEqual': return v >= a;
    case 'lessThan': return v < a;
    case 'lessThanOrEqual': return v <= a;
    case 'between': return v >= Math.min(a, b) && v <= Math.max(a, b);
    case 'notBetween': return v < Math.min(a, b) || v > Math.max(a, b);
  }
};

/** The cache field items of every slot of a pivot field, normalized so every cache item has a slot. */
export const normalizeFields = (pt: PivotTable, cache: PivotCache): PivotField[] =>
  cache.fields.map((cf, f) => {
    const pf = pt.fields[f] ?? { items: [] };
    if (cf.formula) return pf;
    const n = fieldItems(cf).length;
    const seen = new Set<number>();
    const items = pf.items.filter((it) => (it.x < 0 ? !!it.formula : it.x < n && !seen.has(it.x) && (seen.add(it.x), true)));
    if (items.length - items.filter((it) => it.x < 0).length < n) {
      const missing: number[] = [];
      for (let x = 0; x < n; x++) if (!seen.has(x)) missing.push(x);
      const all = fieldItems(cf);
      const ascending = items.every((it, i) => i === 0 || it.x < 0 || items[i - 1]!.x < 0 || compareItems(all[items[i - 1]!.x]!, all[it.x]!) <= 0);
      const hideNew = pf.includeNew === false && items.some((it) => it.hidden);
      for (const x of missing.sort((a, b) => compareItems(all[a]!, all[b]!))) {
        const it = hideNew ? { x, hidden: true } : { x };
        if (ascending && !cf.group) {
          const at = items.findIndex((o) => o.x >= 0 && compareItems(all[o.x]!, all[x]!) > 0);
          if (at >= 0) { items.splice(at, 0, it); continue; }
        }
        items.push(it);
      }
    }
    return { ...pf, items };
  });

export const computePivot = (pt0: PivotTable, cache: PivotCache, date1904 = false): PivotLayout => {
  const errors: string[] = [];
  const fields = normalizeFields(pt0, cache);
  const pt: PivotTable = { ...pt0, fields };
  const items = recordItems(cache, date1904);
  const nRec = cache.records.length;
  const nData = pt.data.length;
  const rowAxis = pt.rows.filter((f) => f === VALUES_FIELD ? nData > 1 : !!cache.fields[f]);
  const colAxis = pt.cols.filter((f) => f === VALUES_FIELD ? nData > 1 : !!cache.fields[f]);
  // Σ Values goes into the columns when there are several data fields and it wasn't placed.
  if (nData > 1 && !rowAxis.includes(VALUES_FIELD) && !colAxis.includes(VALUES_FIELD)) colAxis.push(VALUES_FIELD);
  const rowFields = rowAxis.filter((f) => f !== VALUES_FIELD);
  const colFields = colAxis.filter((f) => f !== VALUES_FIELD);

  // slot of each cache item per field
  const slotOf = fields.map((pf) => {
    const m = new Map<number, number>();
    pf.items.forEach((it, i) => { if (it.x >= 0) m.set(it.x, i); });
    return m;
  });
  const recSlot = (f: number, r: number): number => slotOf[f]!.get(items[f]![r]!) ?? -1;

  const caption = (f: number, slot: number): string => {
    const it = fields[f]?.items[slot];
    if (!it) return '';
    if (it.name) return it.name;
    if (it.x < 0) return '';
    const v = fieldItems(cache.fields[f]!)[it.x] ?? null;
    return captionOf(v, cache.fields[f]!.group ? undefined : (fields[f]!.numFmt ?? cache.fields[f]!.numFmt), cache.fields[f]!.dates, date1904);
  };
  const labelValue = (f: number, slot: number): Scalar => {
    const it = fields[f]?.items[slot];
    if (!it) return null;
    if (it.name) return it.name;
    if (it.x < 0) return it.formula ? '' : null;
    const cf = cache.fields[f]!;
    const v = fieldItems(cf)[it.x] ?? null;
    // Year groups read as numbers, like the dates they come from.
    if (cf.group?.range?.by === 'years' && typeof v === 'string' && /^\d{1,4}$/.test(v)) return +v;
    return v === null || v === '' ? '(blank)' : v;
  };

  // ---------- record filters: hidden items, report filters, label filters
  const allowed: (Uint8Array | null)[] = fields.map((pf, f) => {
    if (cache.fields[f]!.formula) return null;
    const onAxis = rowFields.includes(f) || colFields.includes(f);
    const page = pt.pages.find((p) => p.field === f);
    const hasHidden = pf.items.some((it) => it.hidden);
    const labels = onAxis ? (pf.filters ?? []).filter((x): x is Extract<PivotFilter, { kind: 'label' }> => x.kind === 'label') : [];
    if (!hasHidden && !labels.length && !(page && page.item !== undefined)) return null;
    const ok = new Uint8Array(pf.items.length);
    pf.items.forEach((it, s) => {
      let pass = !it.hidden;
      if (page && page.item !== undefined) pass = s === page.item;
      if (pass && labels.length) {
        const raw = it.x >= 0 ? fieldItems(cache.fields[f]!)[it.x] ?? null : null;
        pass = labels.every((lf) => labelMatch(lf, caption(f, s), raw));
      }
      ok[s] = pass ? 1 : 0;
    });
    return ok;
  });
  // Timeline date ranges test the source dates themselves.
  const dateRanges: { di: number; start: number; end: number }[] = [];
  fields.forEach((pf, f) => {
    for (const x of pf.filters ?? []) {
      if (x.kind !== 'dates') continue;
      const base = cache.fields[f]!.group?.base ?? f;
      const di = dbIndex(cache, base);
      if (di >= 0) dateRanges.push({ di, start: Math.floor(x.start), end: Math.floor(x.end) + 1 });
    }
  });
  let recs: number[] = [];
  for (let r = 0; r < nRec; r++) {
    let pass = true;
    for (let f = 0; f < allowed.length && pass; f++) { const a = allowed[f]; if (a) { const s = recSlot(f, r); pass = s >= 0 && a[s] === 1; } }
    for (const d of dateRanges) { if (!pass) break; const v = cache.records[r]![d.di]; pass = typeof v === 'number' && v >= d.start && v < d.end; }
    if (pass) recs.push(r);
  }

  // ---------- record values
  const valueCache = new Map<number, (r: number) => Scalar>();
  const recValue = (f: number): ((r: number) => Scalar) => {
    let fn = valueCache.get(f);
    if (fn) return fn;
    const di = dbIndex(cache, f);
    fn = di >= 0 ? (r) => cache.records[r]![di] ?? null : () => null;
    valueCache.set(f, fn);
    return fn;
  };
  const calcFields = new Map<number, PNode | null>();
  const calcFieldNode = (f: number): PNode | null => {
    if (!calcFields.has(f)) {
      try { calcFields.set(f, parsePivotFormula(cache.fields[f]!.formula!)); } catch (e) { errors.push(`${cache.fields[f]!.name}: ${(e as Error).message}`); calcFields.set(f, null); }
    }
    return calcFields.get(f)!;
  };
  const fieldByName = (name: string): number => cache.fields.findIndex((cf) => cf.name.toLowerCase() === name.toLowerCase());
  // Source fields whose sums a data field needs.
  const sourcesOf = (f: number, seen = new Set<number>()): number[] => {
    const cf = cache.fields[f];
    if (!cf?.formula) return [f];
    if (seen.has(f)) return [];
    seen.add(f);
    const node = calcFieldNode(f);
    if (!node) return [];
    const out: number[] = [];
    const walk = (n: PNode) => {
      if (n.t === 'ref') { const g = fieldByName(n.name); if (g >= 0) out.push(...sourcesOf(g, seen)); }
      else if (n.t === 'call') n.args.forEach(walk);
      else if (n.t === 'bin') { walk(n.a); walk(n.b); }
      else if (n.t === 'neg' || n.t === 'pct') walk(n.a);
    };
    walk(node);
    return out;
  };
  const needed = [...new Set(pt.data.flatMap((d) => sourcesOf(d.field)))];
  const neededIdx = new Map(needed.map((f, i) => [f, i]));
  const distinctFor = new Set(pt.data.filter((d) => d.agg === 'distinctCount').map((d) => d.field));

  // ---------- trees
  let nextId = 0;
  const calcCache = new Map<string, PNode | null>();
  const calcNode = (f: number, s: number): PNode | null => {
    const k = `${f}:${s}`;
    if (!calcCache.has(k)) {
      const src = fields[f]!.items[s]!.formula!;
      try { calcCache.set(k, parsePivotFormula(src)); } catch (e) { errors.push(`${caption(f, s)}: ${(e as Error).message}`); calcCache.set(k, null); }
    }
    return calcCache.get(k)!;
  };

  /** Slot of an item named in a calculated item formula ("East" or Region[East]). */
  const slotByName = (f: number, name: string, item?: string): number => {
    if (item !== undefined && fieldCaption(pt, cache, f).toLowerCase() !== name.toLowerCase() && cache.fields[f]!.name.toLowerCase() !== name.toLowerCase()) return -1;
    const want = (item ?? name).toLowerCase();
    return fields[f]!.items.findIndex((_, i) => caption(f, i).toLowerCase() === want);
  };
  const formulaSlots = (f: number, node: PNode): number[] => {
    const out: number[] = [];
    const walk = (n: PNode) => {
      if (n.t === 'ref') { const s = slotByName(f, n.name, n.item); if (s >= 0) { const it = fields[f]!.items[s]!; if (it.x >= 0) out.push(s); else { const sub = calcNode(f, s); if (sub && sub !== node) out.push(...formulaSlots(f, sub)); } } }
      else if (n.t === 'call') n.args.forEach(walk);
      else if (n.t === 'bin') { walk(n.a); walk(n.b); }
      else if (n.t === 'neg' || n.t === 'pct') walk(n.a);
    };
    walk(node);
    return out;
  };

  const buildTree = (axisFields: number[], rs: number[]): PNodeT => {
    const root: PNodeT = { id: nextId++, field: -1, slot: -1, depth: -1, recs: rs, kids: [], parent: null, key: '' };
    // `shape` are the records that decide which items appear; under a calculated item they are the
    // records of the items its formula uses, and the nodes keep no records of their own.
    const grow = (node: PNodeT, level: number, shape: number[] = node.recs): void => {
      if (level >= axisFields.length) return;
      const f = axisFields[level]!;
      const pf = fields[f]!;
      const buckets = new Map<number, number[]>();
      for (const r of shape) { const s = recSlot(f, r); if (s < 0) continue; let b = buckets.get(s); if (!b) buckets.set(s, (b = [])); b.push(r); }
      const ok = allowed[f];
      const inCalc = node.calcRoot;
      pf.items.forEach((it, s) => {
        if (it.hidden || (ok && !ok[s] && it.x >= 0)) return;
        const b = buckets.get(s);
        if (it.x < 0) {
          if (!it.formula || inCalc) return;
          const cn = calcNode(f, s);
          if (!cn) return;
          const kid: PNodeT = { id: nextId++, field: f, slot: s, depth: level, recs: [], kids: [], parent: node, calc: cn, key: `${node.key}/${f}:${s}` };
          kid.calcRoot = kid;
          node.kids.push(kid);
          const used = new Set(formulaSlots(f, cn));
          grow(kid, level + 1, used.size ? shape.filter((r) => used.has(recSlot(f, r))) : shape);
          return;
        }
        if (!b && !pf.showAll) return;
        const kid: PNodeT = inCalc
          ? { id: nextId++, field: f, slot: s, depth: level, recs: [], kids: [], parent: node, calc: node.calc, calcRoot: inCalc, key: `${node.key}/${f}:${s}` }
          : { id: nextId++, field: f, slot: s, depth: level, recs: b ?? [], kids: [], parent: node, key: `${node.key}/${f}:${s}` };
        node.kids.push(kid);
        grow(kid, level + 1, b ?? []);
      });
    };
    grow(root, 0);
    return root;
  };

  // Value filters and top N narrow the records, one axis field at a time.
  const valueFilterPass = (axisFields: number[], otherFields: number[], isRow: boolean): void => {
    axisFields.forEach((f, level) => {
      const flt = (fields[f]!.filters ?? []).filter((x): x is Extract<PivotFilter, { kind: 'value' | 'top' }> => x.kind === 'value' || x.kind === 'top');
      if (!flt.length) return;
      for (const fl of flt) {
        if (fl.dataField < 0 || fl.dataField >= nData) continue;
        const tree = buildTree(axisFields.slice(0, level + 1), recs);
        const drop = new Set<number>();
        const groups: PNodeT[][] = [];
        const collect = (n: PNodeT) => { if (n.depth === level - 1 || (level === 0 && n.depth === -1)) { groups.push(n.kids.filter((k) => !k.calc)); return; } n.kids.forEach(collect); };
        collect(tree);
        for (const g of groups) {
          const vals = g.map((k) => simpleValue(fl.dataField, k.recs));
          if (fl.kind === 'value') g.forEach((k, i) => { if (!valueMatch(fl, vals[i]!)) for (const r of k.recs) drop.add(r); });
          else {
            const order = g.map((_, i) => i).filter((i) => typeof vals[i] === 'number').sort((a, b) => (fl.top ? (vals[b] as number) - (vals[a] as number) : (vals[a] as number) - (vals[b] as number)));
            const keep = new Set<number>();
            if (fl.by === 'count') order.slice(0, Math.max(0, Math.floor(fl.value))).forEach((i) => keep.add(i));
            else {
              const total = order.reduce((s, i) => s + (vals[i] as number), 0);
              const target = fl.by === 'percent' ? (total * fl.value) / 100 : fl.value;
              let run = 0;
              for (const i of order) { if (run >= target && keep.size) break; keep.add(i); run += vals[i] as number; }
            }
            g.forEach((k, i) => { if (!keep.has(i)) for (const r of k.recs) drop.add(r); });
          }
        }
        if (drop.size) recs = recs.filter((r) => !drop.has(r));
      }
    });
    void otherFields; void isRow;
  };
  const simpleValue = (d: number, rs: number[]): Scalar => {
    const acc = new Acc();
    const df = pt.data[d]!;
    if (cache.fields[df.field]?.formula) return calcFieldValue(df.field, (g) => { const a = new Acc(); const fv = recValue(g); for (const r of rs) a.add(fv(r)); return a; });
    const fv = recValue(df.field);
    for (const r of rs) acc.add(fv(r), df.agg === 'distinctCount');
    return acc.result(df.agg);
  };
  const calcFieldValue = (f: number, accOf: (g: number) => Acc, depth = 0): Scalar => {
    const node = calcFieldNode(f);
    if (!node || depth > 20) return Err.ref;
    return evalPivotFormula(node, (name) => {
      const g = fieldByName(name);
      if (g < 0) return new CellError('#NAME?', name);
      if (cache.fields[g]!.formula) return calcFieldValue(g, accOf, depth + 1);
      const a = accOf(g);
      return a.err ?? a.sum;
    });
  };
  valueFilterPass(rowFields, colFields, true);
  valueFilterPass(colFields, rowFields, false);

  const rowRoot = buildTree(rowFields, recs);
  const colRoot = buildTree(colFields, recs);

  // record -> deepest column node
  const colLeaf = new Int32Array(nRec).fill(-1);
  const colNodes: PNodeT[] = [];
  const indexCols = (n: PNodeT) => { colNodes[n.id] = n; if (!n.calc) for (const r of n.recs) colLeaf[r] = n.id; n.kids.forEach(indexCols); };
  indexCols(colRoot);

  // per row node: column node id -> accumulators for the needed fields
  const buckets = new Map<number, Map<number, Acc[]>>();
  const bucketOf = (rn: PNodeT): Map<number, Acc[]> => {
    let m = buckets.get(rn.id);
    if (m) return m;
    m = new Map();
    const fvs = needed.map((f) => recValue(f));
    const dist = needed.map((f) => distinctFor.has(f));
    for (const r of rn.recs) {
      for (let c: PNodeT | null = colNodes[colLeaf[r]!] ?? null; c; c = c.parent) {
        let accs = m.get(c.id);
        if (!accs) { accs = needed.map(() => new Acc()); m.set(c.id, accs); }
        for (let i = 0; i < fvs.length; i++) accs[i]!.add(fvs[i]!(r), dist[i]);
      }
    }
    buckets.set(rn.id, m);
    return m;
  };
  const accAt = (rn: PNodeT, cn: PNodeT, f: number): Acc => bucketOf(rn).get(cn.id)?.[neededIdx.get(f)!] ?? new Acc();

  const memo = new Map<string, Scalar>();
  const descCalcs = (n: PNodeT): PNodeT[] => {
    const out: PNodeT[] = [];
    const walk = (x: PNodeT) => { for (const k of x.kids) { if (k.calc) out.push(k); else walk(k); } };
    walk(n);
    return out;
  };
  /**
   * A calculated item's value in one summarised cell: its formula over the values the other items of
   * its field have in that same cell (same inner items, same position on the other axis).
   */
  const calcValue = (d: number, fn: Subtotal, isRow: boolean, formula: PNode, root: PNodeT, inner: PathEl[], other: PNodeT, depth: number): Scalar => {
    const F = root.field, base = root.parent!;
    return evalPivotFormula(formula, (name, item) => {
      const s = slotByName(F, name, item);
      if (s < 0) return new CellError('#NAME?', name);
      const it = fields[F]!.items[s]!;
      if (it.x < 0) {
        const sub = calcNode(F, s);
        if (!sub || depth > 20) return Err.ref;
        const v = calcValue(d, fn, isRow, sub, root, inner, other, depth + 1);
        return v === null ? 0 : v;
      }
      const recs = base.recs.filter((r) => recSlot(F, r) === s && inner.every((e) => recSlot(e.field, r) === e.slot));
      const t: PNodeT = { id: -1, field: F, slot: s, depth: root.depth, recs, kids: [], parent: base, key: `~${F}:${s}:${inner.map((e) => `${e.field}.${e.slot}`).join(',')}` };
      const v = isRow ? valueAt(d, t, other, fn, depth + 1) : valueAt(d, other, t, fn, depth + 1);
      return v === null ? 0 : v;
    });
  };
  const innerPath = (n: PNodeT): PathEl[] => {
    const out: PathEl[] = [];
    for (let x: PNodeT | null = n; x && x !== n.calcRoot; x = x.parent) out.unshift({ field: x.field, slot: x.slot });
    return out;
  };
  /** A total that takes in calculated item values, for each summary function. */
  const combine = (agg: PivotAgg, base: Acc, extras: Scalar[]): Scalar => {
    const err = extras.find((v): v is CellError => isError(v));
    if (err) return err;
    const nums = extras.filter((v): v is number => typeof v === 'number');
    if (!nums.length) return base.result(agg);
    const r = base.result(agg);
    const real = typeof r === 'number' ? r : null;
    if (isError(r)) return r;
    switch (agg) {
      case 'sum': case 'count': case 'countNums': case 'distinctCount': return (real ?? 0) + nums.reduce((a, b) => a + b, 0);
      case 'max': return Math.max(...(base.nums ? [real!] : []), ...nums);
      case 'min': return Math.min(...(base.nums ? [real!] : []), ...nums);
      case 'product': return (base.nums ? real! : 1) * nums.reduce((a, b) => a * b, 1);
      default: {
        // Average and spread: each calculated cell counts as one more value.
        const acc = new Acc();
        acc.merge(base);
        for (const v of nums) acc.add(v);
        return acc.result(agg);
      }
    }
  };

  /** Raw aggregate of data field d at a row node and a column node. */
  const valueAt = (d: number, rn: PNodeT, cn: PNodeT, fn: Subtotal = 'default', depth = 0): Scalar => {
    const df = pt.data[d];
    if (!df) return null;
    const k = `${d}|${rn.id}|${cn.id}|${fn}|${rn.key}|${cn.key}`;
    if (rn.id >= 0 && cn.id >= 0 && memo.has(k)) return memo.get(k)!;
    let out: Scalar;
    if (depth > 30) out = Err.ref;
    else if (rn.calc || cn.calc) {
      const calcN = rn.calc ? rn : cn;
      out = calcValue(d, fn, !!rn.calc, calcN.calc!, calcN.calcRoot ?? calcN, innerPath(calcN), rn.calc ? cn : rn, depth);
    } else {
      const agg: PivotAgg = fn === 'default' ? df.agg : SUBTOTAL_AGG[fn];
      const recsHere = rn.id < 0 || cn.id < 0 ? intersect(rn.recs, cn.recs) : null;
      if (cache.fields[df.field]?.formula) {
        const any = recsHere ? recsHere.length > 0 : bucketOf(rn).has(cn.id);
        out = any ? calcFieldValue(df.field, (g) => (recsHere ? accOver(recsHere, g) : accAt(rn, cn, g))) : null;
      } else {
        const base = recsHere ? accOver(recsHere, df.field, df.agg === 'distinctCount') : accAt(rn, cn, df.field);
        const extra = [...descCalcs(rn).map((c) => valueAt(d, c, cn, fn, depth + 1)), ...descCalcs(cn).map((c) => valueAt(d, rn, c, fn, depth + 1))];
        out = extra.length ? combine(agg, base, extra) : base.result(agg);
      }
    }
    if (rn.id >= 0 && cn.id >= 0) memo.set(k, out);
    return out;
  };
  const accOver = (rs: number[], f: number, distinct = false): Acc => { const a = new Acc(); const fv = recValue(f); for (const r of rs) a.add(fv(r), distinct); return a; };

  // ---------- sorting
  const sortKids = (n: PNodeT, isRow: boolean): void => {
    if (n.kids.length > 1) {
      const pf = fields[n.kids[0]!.field]!;
      const dir = pf.sort === 'descending' ? -1 : 1;
      if (pf.sortByData !== undefined && pf.sortByData < nData) {
        const d = pf.sortByData;
        const val = (k: PNodeT) => (isRow ? valueAt(d, k, colRoot) : valueAt(d, rowRoot, k));
        const vals = new Map(n.kids.map((k) => [k.id, val(k)]));
        const order = new Map(n.kids.map((k, i) => [k.id, i]));
        n.kids.sort((a, b) => {
          const x = vals.get(a.id) ?? null, y = vals.get(b.id) ?? null;
          const c = typeof x === 'number' && typeof y === 'number' ? x - y : typeof x === 'number' ? -1 : typeof y === 'number' ? 1 : 0;
          return c ? c * dir : order.get(a.id)! - order.get(b.id)!;
        });
      } else if (pf.sort) {
        const all = fieldItems(cache.fields[n.kids[0]!.field]!);
        const raw = (k: PNodeT): Scalar => { const it = pf.items[k.slot]!; return it.name ?? (it.x >= 0 ? all[it.x] ?? null : caption(k.field, k.slot)); };
        const grouped = !!cache.fields[n.kids[0]!.field]!.group;
        n.kids.sort((a, b) => dir * (grouped ? (pf.items[a.slot]!.x - pf.items[b.slot]!.x) : compareItems(raw(a), raw(b))));
      }
    }
    n.kids.forEach((k) => sortKids(k, isRow));
  };
  sortKids(rowRoot, true);
  sortKids(colRoot, false);

  // ---------- lines
  const subtotalsOf = (f: number): Subtotal[] => (f === VALUES_FIELD ? [] : fields[f]!.subtotals ?? ['default']);
  const genLines = (axis: number[], root: PNodeT, isRow: boolean, grand: boolean): Line[] => {
    const lines: Line[] = [];
    if (!axis.length) { if (nData || isRow) lines.push({ kind: 'item', node: root, data: -1, path: [], showValues: true }); return lines; }
    const dataLevel = axis.indexOf(VALUES_FIELD);
    const walk = (node: PNodeT, li: number, path: PathEl[], d: number): void => {
      const f = axis[li]!;
      const last = li === axis.length - 1;
      if (f === VALUES_FIELD) {
        for (let di = 0; di < nData; di++) {
          const p2 = [...path, { field: VALUES_FIELD, slot: di }];
          if (last) { lines.push({ kind: 'item', node, data: di, path: p2, showValues: true }); continue; }
          if (isRow && !isTabular(pt, f)) lines.push({ kind: 'item', node, data: di, path: p2, showValues: true });
          walk(node, li + 1, p2, di);
        }
        return;
      }
      const pf = fields[f]!;
      for (const kid of node.kids) {
        const p2 = [...path, { field: f, slot: kid.slot }];
        const collapsed = !!pf.items[kid.slot]?.collapsed && !last;
        if (last || collapsed) {
          lines.push({ kind: 'item', node: kid, data: d, path: p2, showValues: true, expandable: !last ? (collapsed ? 'collapsed' : undefined) : undefined });
          continue;
        }
        const subs = subtotalsOf(f);
        const dataBelow = dataLevel > li && d < 0;
        const top = isRow && !isTabular(pt, f) && pf.subtotalTop !== false && subs.length === 1 && subs[0] === 'default';
        if (isRow && !isTabular(pt, f)) lines.push({ kind: 'item', node: kid, data: d, path: p2, showValues: top && !dataBelow, expandable: 'expanded' });
        walk(kid, li + 1, p2, d);
        if (subs.length && (!top || dataBelow)) {
          for (const fn of subs) {
            if (dataBelow) for (let di = 0; di < nData; di++) lines.push({ kind: 'sub', node: kid, data: di, path: p2, fn, showValues: true });
            else if (!(top && fn === 'default')) lines.push({ kind: 'sub', node: kid, data: d, path: p2, fn, showValues: true });
          }
        }
        if (isRow && pf.blankRow) lines.push({ kind: 'blank', node: kid, data: d, path: p2, showValues: false });
      }
    };
    walk(root, 0, [], -1);
    if (grand && axis.some((f) => f !== VALUES_FIELD)) {
      if (dataLevel >= 0) for (let di = 0; di < nData; di++) lines.push({ kind: 'grand', node: root, data: di, path: [], showValues: true });
      else lines.push({ kind: 'grand', node: root, data: -1, path: [], showValues: true });
    }
    return lines;
  };
  const rowLines = genLines(rowAxis, rowRoot, true, pt.colGrand);
  const colLines = genLines(colAxis, colRoot, false, pt.rowGrand);

  // ---------- show values as
  const pathMaps = (root: PNodeT) => {
    const m = new Map<string, PNodeT>();
    const walk = (n: PNodeT) => { m.set(n.key, n); n.kids.forEach(walk); };
    walk(root);
    return m;
  };
  const rowMap = pathMaps(rowRoot), colMap = pathMaps(colRoot);
  const ancestorAt = (n: PNodeT, f: number): PNodeT | null => { for (let x: PNodeT | null = n; x && x.field >= 0; x = x.parent) if (x.field === f) return x; return null; };
  const withSlot = (n: PNodeT, anc: PNodeT, slot: number, map: Map<string, PNodeT>): PNodeT | null => {
    const prefix = anc.key;
    const repl = `${anc.parent?.key ?? ''}/${anc.field}:${slot}`;
    const key = repl + n.key.slice(prefix.length);
    return map.get(key) ?? null;
  };

  const showValue = (d: number, rl: Line, cl: Line, raw: Scalar): Scalar => {
    const df = pt.data[d]!;
    const mode = df.showAs ?? 'normal';
    if (mode === 'normal' || isError(raw)) return raw;
    const rn = rl.node, cn = cl.node;
    const fn = rl.fn ?? cl.fn ?? 'default';
    const v = (r: PNodeT, c: PNodeT) => valueAt(d, r, c, fn);
    const div = (a: Scalar, b: Scalar): Scalar => (a === null ? null : typeof a !== 'number' ? a : typeof b !== 'number' ? (isError(b) ? b : Err.div0) : b === 0 ? Err.div0 : a / b);
    const n0 = (x: Scalar) => (x === null ? 0 : x);
    switch (mode) {
      case 'percentOfTotal': return div(raw, v(rowRoot, colRoot));
      case 'percentOfCol': return div(raw, v(rowRoot, cn));
      case 'percentOfRow': return div(raw, v(rn, colRoot));
      case 'percentOfParentRow': return div(raw, v(rl.kind === 'grand' ? rowRoot : rn.parent ?? rn, cn));
      case 'percentOfParentCol': return div(raw, v(rn, cl.kind === 'grand' ? colRoot : cn.parent ?? cn));
      case 'index': {
        const g = v(rowRoot, colRoot), rt = v(rn, colRoot), ct = v(rowRoot, cn);
        return typeof raw === 'number' && typeof g === 'number' && typeof rt === 'number' && typeof ct === 'number' ? (rt * ct === 0 ? Err.div0 : (raw * g) / (rt * ct)) : raw;
      }
    }
    const B = df.baseField ?? -1;
    const onRows = rowFields.includes(B);
    const node = onRows ? rn : cn;
    const anc = B >= 0 ? ancestorAt(node, B) : null;
    if (mode === 'percentOfParent') {
      if (!anc) return null;
      return div(raw, onRows ? v(anc, cn) : v(rn, anc));
    }
    if (!anc) return null;
    const map = onRows ? rowMap : colMap;
    const sibs = anc.parent?.kids.filter((k) => !k.calc) ?? [anc];
    const at = (s: PNodeT): Scalar => { const n2 = withSlot(node, anc, s.slot, map); return n2 ? (onRows ? v(n2, cn) : v(rn, n2)) : null; };
    if (mode === 'percent' || mode === 'difference' || mode === 'percentDiff') {
      let baseNode: PNodeT | undefined;
      const idx = sibs.indexOf(anc);
      if (df.baseItem === 'prev') baseNode = sibs[idx - 1];
      else if (df.baseItem === 'next') baseNode = sibs[idx + 1];
      else baseNode = sibs.find((s) => s.slot === df.baseItem) ?? undefined;
      if (baseNode === anc) return mode === 'percent' ? (raw === null ? null : 1) : null;
      if (!baseNode) {
        if (df.baseItem === 'prev' || df.baseItem === 'next') return mode === 'percent' ? (raw === null ? null : 1) : null;
        // The base item has no data here.
        return mode === 'difference' ? raw : Err.na;
      }
      const b = at(baseNode);
      if (mode === 'percent') return div(raw, b);
      if (mode === 'difference') return typeof raw === 'number' || raw === null ? (raw === null && b === null ? null : (n0(raw) as number) - (typeof b === 'number' ? b : 0)) : raw;
      if (raw === null) return null;
      return typeof raw === 'number' && typeof b === 'number' ? (b === 0 ? Err.div0 : (raw - b) / b) : b === null ? Err.na : raw;
    }
    if (mode === 'runTotal' || mode === 'percentOfRunningTotal') {
      let run = 0;
      for (const s of sibs) { const x = at(s); if (typeof x === 'number') run += x; if (s === anc) break; }
      if (mode === 'runTotal') return run;
      let tot = 0;
      for (const s of sibs) { const x = at(s); if (typeof x === 'number') tot += x; }
      return tot === 0 ? Err.div0 : run / tot;
    }
    if (mode === 'rankAscending' || mode === 'rankDescending') {
      if (typeof raw !== 'number') return null;
      const vals = sibs.map(at).filter((x): x is number => typeof x === 'number');
      const better = vals.filter((x) => (mode === 'rankAscending' ? x < raw : x > raw)).length;
      return better + 1;
    }
    return raw;
  };

  // ---------- grid
  const W0 = rowAxis.length ? 0 : colAxis.length && nData ? 1 : 0;
  const labelCol: number[] = [];
  rowAxis.forEach((_, li) => { labelCol[li] = li === 0 ? 0 : isCompact(pt, rowAxis[li - 1]!) ? labelCol[li - 1]! : labelCol[li - 1]! + 1; });
  const W = rowAxis.length ? labelCol[rowAxis.length - 1]! + 1 : W0;
  const onlyData = colAxis.length === 1 && colAxis[0] === VALUES_FIELD;
  const nColLevels = colAxis.length;
  const colLevelRow = (li: number) => (onlyData ? li : 1 + li);
  const firstHeaderRow = onlyData ? 0 : 1;
  const firstDataRow = nColLevels === 0 ? 1 : colLevelRow(nColLevels - 1) + 1;
  const firstDataCol = W;
  const height = firstDataRow + rowLines.length;
  const width = W + colLines.length;
  const cells: (PCell | null)[][] = Array.from({ length: height }, () => Array<PCell | null>(Math.max(1, width)).fill(null));
  const set = (r: number, c: number, cell: PCell) => { if (r >= 0 && r < height && c >= 0 && c < width) cells[r]![c] = cell; };
  const indentOf = (li: number) => { let n = 0; for (let j = 0; j < li; j++) if (labelCol[j] === labelCol[li]) n++; return n; };
  const dataName = (d: number) => dataFieldName(cache, pt.data[d]!);
  const fieldFmt = (f: number): string | undefined => (cache.fields[f]?.group ? undefined : fields[f]?.numFmt ?? cache.fields[f]?.numFmt ?? (cache.fields[f]?.dates ? 'm/d/yyyy' : undefined));
  const pathLabel = (el: PathEl): Scalar => (el.field === VALUES_FIELD ? dataName(el.slot) : labelValue(el.field, el.slot));
  const pathCaption = (el: PathEl): string => (el.field === VALUES_FIELD ? dataName(el.slot) : caption(el.field, el.slot));
  const fnLabel = (fn: Subtotal) => (fn === 'default' ? 'Total' : ({ sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' } as const)[fn]);
  const grandCaption = pt.grandTotalCaption ?? 'Grand Total';

  // Corner captions.
  if (nColLevels > 0 && !onlyData) {
    if (W > 0) set(0, 0, { v: nData === 1 && pt.showHeaders !== false ? dataName(0) : null, role: 'dataCaption', data: nData === 1 ? 0 : undefined });
    if (pt.layout === 'compact' || colAxis.length === 1) {
      const cap = colAxis.length === 1 && pt.layout !== 'compact' ? fieldCaption(pt, cache, colAxis[0]!) : (pt.colHeaderCaption ?? 'Column Labels');
      set(0, W, { v: pt.showHeaders !== false ? cap : null, role: 'colCaption', filter: 'cols', field: colAxis.length === 1 ? colAxis[0] : undefined });
    } else colAxis.forEach((f, i) => set(0, W + i, { v: pt.showHeaders !== false ? fieldCaption(pt, cache, f) : null, role: 'colCaption', filter: f === VALUES_FIELD ? undefined : f, field: f }));
  }
  const hdr = firstDataRow - 1;
  for (let li = 0; li < rowAxis.length; li++) {
    const c = labelCol[li]!;
    const shared = rowAxis.filter((_, j) => labelCol[j] === c);
    if (cells[hdr]![c]?.role === 'rowCaption') continue;
    const cap = shared.length > 1 ? (pt.rowHeaderCaption ?? 'Row Labels') : pt.layout === 'compact' && rowAxis.length >= 1 && shared.length === 1 && isCompact(pt, shared[0]!) ? (pt.rowHeaderCaption ?? 'Row Labels') : fieldCaption(pt, cache, shared[0]!);
    set(hdr, c, { v: pt.showHeaders !== false ? cap : null, role: 'rowCaption', filter: shared.length > 1 ? 'rows' : (shared[0] === VALUES_FIELD ? undefined : shared[0]), field: shared.length === 1 ? shared[0] : undefined });
  }

  // Column headers.
  colLines.forEach((cl, j) => {
    const c = W + j;
    const prev = colLines[j - 1];
    if (cl.kind === 'grand') {
      set(firstHeaderRow, c, { v: cl.data >= 0 ? `Total ${dataName(cl.data)}` : grandCaption, role: 'colItem', total: 'grand', colLine: j });
      return;
    }
    if (!colAxis.length) {
      if (nData) set(0, c, { v: nData === 1 ? dataName(0) : null, role: 'dataCaption', colLine: j, data: 0 });
      return;
    }
    const k = cl.path.length;
    cl.path.forEach((el, li) => {
      const same = prev && prev.kind !== 'grand' && prev.path.length > li && prev.path.slice(0, li + 1).every((p, i) => p.field === cl.path[i]!.field && p.slot === cl.path[i]!.slot) && prev.kind === 'item';
      const isLast = li === k - 1;
      if (cl.kind === 'sub') {
        if (!isLast) return;
        const lbl = cl.data >= 0 && cl.path.every((p) => p.field !== VALUES_FIELD) ? `${pathCaption(el)} ${dataName(cl.data)}` : `${pathCaption(el)} ${fnLabel(cl.fn!)}`;
        set(colLevelRow(li), c, { v: lbl, role: 'colItem', total: 'sub', colLine: j, field: el.field, slot: el.slot });
        return;
      }
      if (same && !isLast) return;
      set(colLevelRow(li), c, {
        v: pathLabel(el), role: 'colItem', colLine: j, field: el.field, slot: el.slot, numFmt: el.field >= 0 ? fieldFmt(el.field) : undefined,
        button: pt.showDrill && isLast && cl.expandable ? cl.expandable : pt.showDrill && !isLast && el.field >= 0 && li < colFields.length - 1 && cl.path[li + 1]?.field !== undefined && !same ? 'expanded' : undefined,
      });
    });
  });

  // Row labels and values.
  rowLines.forEach((rl, i) => {
    const r = firstDataRow + i;
    const prev = rowLines[i - 1];
    if (rl.kind === 'grand') {
      set(r, 0, { v: rl.data >= 0 ? `Total ${dataName(rl.data)}` : grandCaption, role: 'rowItem', total: 'grand', rowLine: i });
      if (W === 0 && width > 0) cells[r]![0] = null;
    } else if (rl.kind !== 'blank') {
      if (!rowAxis.length) {
        if (W > 0) set(r, 0, { v: nData === 1 ? dataName(0) : null, role: 'rowItem', rowLine: i, data: nData === 1 ? 0 : undefined });
      }
      const k = rl.path.length;
      rl.path.forEach((el, li) => {
        const isLast = li === k - 1;
        const c = labelCol[li]!;
        if (rl.kind === 'sub') {
          if (isLast) {
            const lbl = rl.data >= 0 && rl.path.every((p) => p.field !== VALUES_FIELD) ? `${pathCaption(el)} ${dataName(rl.data)}` : `${pathCaption(el)} ${fnLabel(rl.fn!)}`;
            set(r, c, { v: lbl, role: 'rowItem', total: 'sub', rowLine: i, field: el.field, slot: el.slot, indent: isCompact(pt, el.field) || li > 0 ? indentOf(li) : 0 });
          } else if (el.field >= 0 && fields[el.field]?.repeatLabels && labelCol[li] !== labelCol[k - 1]) set(r, c, { v: pathLabel(el), role: 'rowItem', rowLine: i, field: el.field, slot: el.slot, numFmt: fieldFmt(el.field) });
          return;
        }
        if (!isLast) {
          if (labelCol[li] === labelCol[k - 1]) return;
          const firstOfGroup = !prev || prev.kind === 'grand' || prev.path.length <= li || !prev.path.slice(0, li + 1).every((p, j) => p.field === rl.path[j]!.field && p.slot === rl.path[j]!.slot);
          const repeat = el.field >= 0 && !!fields[el.field]?.repeatLabels;
          if (!(repeat || (isTabular(pt, rowAxis[li]!) && firstOfGroup))) return;
          set(r, c, { v: pathLabel(el), role: 'rowItem', rowLine: i, field: el.field, slot: el.slot, numFmt: el.field >= 0 ? fieldFmt(el.field) : undefined, button: pt.showDrill && firstOfGroup && el.field >= 0 ? 'expanded' : undefined });
          return;
        }
        set(r, c, {
          v: pathLabel(el), role: 'rowItem', rowLine: i, field: el.field, slot: el.slot, numFmt: el.field >= 0 ? fieldFmt(el.field) : undefined,
          indent: indentOf(li), button: pt.showDrill && rl.expandable ? rl.expandable : undefined,
        });
      });
    }
    colLines.forEach((cl, j) => {
      const c = W + j;
      if (rl.kind === 'blank') return;
      const d = rl.data >= 0 ? rl.data : cl.data >= 0 ? cl.data : nData === 1 ? 0 : -1;
      if (d < 0) { set(r, c, { v: null, role: 'data', rowLine: i, colLine: j }); return; }
      if (!rl.showValues || !cl.showValues) { set(r, c, { v: null, role: 'data', rowLine: i, colLine: j, data: d }); return; }
      const fn: Subtotal = rl.kind === 'sub' ? rl.fn! : cl.kind === 'sub' ? cl.fn! : 'default';
      const raw = valueAt(d, rl.node, cl.node, fn);
      let v = showValue(d, rl, cl, raw);
      if (v === null && pt.showEmpty && pt.emptyText !== undefined) v = pt.emptyText;
      if (isError(v) && pt.showError && pt.errorText !== undefined) v = pt.errorText;
      set(r, c, { v, role: 'data', rowLine: i, colLine: j, data: d, numFmt: valueFormat(pt.data[d]!, cache), total: rl.kind === 'grand' || cl.kind === 'grand' ? 'grand' : rl.kind === 'sub' || cl.kind === 'sub' || (rl.showValues && rl.expandable === 'expanded') ? 'sub' : undefined });
    });
  });

  // ---------- report filters
  const pageCells: PivotLayout['pageCells'] = [];
  const pages = pt.pages.filter((p) => cache.fields[p.field]);
  const wrap = pt.pageWrap && pt.pageWrap > 0 ? pt.pageWrap : 0;
  const overThenDown = !!pt.pageOverThenDown;
  let pageRows = 0;
  if (pages.length) {
    const nPerLine = wrap || pages.length;
    const nLines = Math.ceil(pages.length / nPerLine);
    const rowsUsed = overThenDown ? nLines : Math.min(pages.length, nPerLine);
    pageRows = rowsUsed + 1;
    pages.forEach((p, k) => {
      const line = Math.floor(k / nPerLine), pos = k % nPerLine;
      const dr = overThenDown ? line : pos;
      const dc = overThenDown ? pos * 3 : line * 3;
      const pf = fields[p.field]!;
      const hidden = pf.items.some((it) => it.hidden);
      const val = p.item !== undefined ? labelValue(p.field, p.item) : hidden ? '(Multiple Items)' : '(All)';
      pageCells.push({ dr: dr - pageRows, dc, cell: { v: fieldCaption(pt, cache, p.field), role: 'pageLabel', field: p.field, filter: p.field } });
      pageCells.push({ dr: dr - pageRows, dc: dc + 1, cell: { v: val, role: 'pageValue', field: p.field, slot: p.item, filter: p.field, numFmt: p.item !== undefined ? fieldFmt(p.field) : undefined } });
    });
  }

  // ---------- rowItems / colItems
  const axisItems = (lines: Line[]): AxisItem[] => {
    const out: AxisItem[] = [];
    let prev: PathEl[] = [];
    for (const l of lines) {
      const x = l.path.map((p) => p.slot);
      const it: AxisItem = { x };
      if (l.kind === 'grand') { it.t = 'grand'; it.x = [0]; prev = []; if (l.data > 0) it.i = l.data; out.push(it); continue; }
      if (l.kind === 'sub') it.t = SUBTOTAL_TYPE[l.fn!];
      if (l.kind === 'blank') it.t = 'blank';
      let common = 0;
      while (common < prev.length && common < l.path.length && prev[common]!.field === l.path[common]!.field && prev[common]!.slot === l.path[common]!.slot) common++;
      const r = Math.max(0, Math.min(common, l.path.length - 1));
      if (r) it.r = r;
      it.x = x.slice(r);
      if (l.data > 0) it.i = l.data;
      out.push(it);
      prev = l.path;
    }
    return out;
  };

  const recordsAt = (ri: number, ci: number): number[] => {
    const rl = rowLines[ri], cl = colLines[ci];
    if (!rl || !cl) return [];
    return intersect(rl.node.recs, cl.node.recs);
  };

  return {
    width, height, cells, pageCells, pageRows, firstHeaderRow, firstDataRow, firstDataCol, rowLines, colLines,
    rowItems: axisItems(rowLines), colItems: axisItems(colLines),
    recordsAt, valueAt: (d, rn, cn, fn) => valueAt(d, rn, cn, fn ?? 'default'), rowRoot, colRoot, records: recs, caption, errors,
  };
};

const intersect = (a: number[], b: number[]): number[] => {
  if (a.length > b.length) [a, b] = [b, a];
  const s = new Set(b);
  return a.filter((x) => s.has(x));
};

export const captionOf = (v: Scalar, fmt: string | undefined, dates: boolean | undefined, date1904: boolean): string => {
  if (v === null || v === '') return '(blank)';
  if (isError(v)) return v.code;
  if (typeof v === 'number') return formatValue(v, fmt ?? (dates ? 'm/d/yyyy' : 'General'), { date1904 }).text;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v;
};

const PERCENT_MODES = new Set(['percentOfTotal', 'percentOfCol', 'percentOfRow', 'percentOfParentRow', 'percentOfParentCol', 'percentOfParent', 'percent', 'percentDiff', 'percentOfRunningTotal']);
export const valueFormat = (df: DataField, cache: PivotCache): string | undefined => {
  if (df.numFmt) return df.numFmt;
  const mode = df.showAs ?? 'normal';
  if (PERCENT_MODES.has(mode)) return '0.00%';
  if (mode === 'rankAscending' || mode === 'rankDescending') return '0';
  if (mode === 'index') return '0.00';
  if (df.agg === 'count' || df.agg === 'countNums' || df.agg === 'distinctCount') return undefined;
  const f = cache.fields[df.field];
  if (f?.dates && (df.agg === 'max' || df.agg === 'min')) return f.numFmt ?? 'm/d/yyyy';
  return f?.dates ? undefined : f?.numFmt;
};
