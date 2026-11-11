import { Evaluator, type Host } from './evaluator';
import { parseFormula, walk, type Node } from './parser';
import { getFunction, VOLATILE_SPECIAL } from './registry';
import { CellError, Err, isError, isMatrix, MAX_COLS, MAX_ROWS, ref, type Matrix, type RangeRef, type Scalar } from './types';
import { cellKey, keyCol, keyRow, type Cell, type Workbook } from '../model/workbook';
import { pivotResolver } from './pivotHook';

/** "sheet:key" for one formula cell. */
type Id = string;
const idOf = (sheet: number, key: number): Id => `${sheet}:${key}`;
const splitId = (id: Id): [number, number] => {
  const i = id.indexOf(':');
  return [+id.slice(0, i), +id.slice(i + 1)];
};

interface Area { sheet: number; r1: number; c1: number; r2: number; c2: number }
interface Deps { cells: Area[]; volatile: boolean }

interface SpillState { anchors: Map<number, { rows: number; cols: number }>; owner: Map<number, number>; values: Map<number, Scalar> }

const BUCKET_LIMIT = 64;

/**
 * Keeps every formula in a workbook calculated. Changes go through setCell/setFormula/setValue
 * and then recalc() works out what is affected and recalculates only that, in dependency order.
 */
export class Engine implements Host {
  readonly evaluator: Evaluator;
  private deps = new Map<Id, Deps>();
  private cellDependents = new Map<string, Set<Id>>();
  private rangeDependents = new Map<number, { cols: Map<number, { r1: number; r2: number; id: Id }[]>; wide: (Area & { id: Id })[] }>();
  private volatile = new Set<Id>();
  private spills: SpillState[] = [];
  private pending = new Set<Id>();
  private changedCells: Area[] = [];
  private bounds = new Map<number, { rows: number; cols: number }>();
  /** Formulas showing #SPILL!, with the area they want, so they retry when that area frees up. */
  private blockedSpills = new Map<Id, Area>();
  /** Cells found in a circular reference during the last calculation. */
  circular: { sheet: number; row: number; col: number }[] = [];
  /** Iterative calculation settings (File > Options > Formulas in other apps). */
  iterate = false;
  maxIterations = 100;
  maxChange = 0.001;
  /** Manual calculation: only formulas that were just entered are calculated until a full recalculation. */
  manual = false;
  /** Some formulas may be out of date because calculation is manual. */
  stale = false;
  /** Precision as displayed: rounds each numeric formula result to what its number format shows. */
  roundResult: ((sheet: number, row: number, col: number, v: number) => number) | null = null;

  constructor(public wb: Workbook) {
    this.evaluator = new Evaluator(this);
    this.rebuild();
  }

  get date1904(): boolean { return this.wb.date1904; }

  // ---------- Host ----------

  getValue(sheet: number, row: number, col: number): Scalar {
    const s = this.wb.sheets[sheet];
    if (!s) return Err.ref;
    const key = cellKey(row, col);
    const c = s.cells.get(key);
    if (c && (c.v !== null || c.f)) return c.v;
    const sp = this.spills[sheet];
    if (sp) { const v = sp.values.get(key); if (v !== undefined) return v; }
    return c?.v ?? null;
  }

  usedBounds(sheet: number): { rows: number; cols: number } {
    let b = this.bounds.get(sheet);
    if (!b) {
      let rows = -1, cols = -1;
      const s = this.wb.sheets[sheet];
      if (s) {
        for (const [k, c] of s.cells) if (c.v !== null || c.f) { rows = Math.max(rows, keyRow(k)); cols = Math.max(cols, keyCol(k)); }
        const sp = this.spills[sheet];
        if (sp) for (const k of sp.values.keys()) { rows = Math.max(rows, keyRow(k)); cols = Math.max(cols, keyCol(k)); }
      }
      b = { rows, cols };
      this.bounds.set(sheet, b);
    }
    return b;
  }

  sheetIndex(name: string): number {
    const n = name.toLowerCase();
    return this.wb.sheets.findIndex((s) => s.name.toLowerCase() === n);
  }
  sheetName(i: number): string { return this.wb.sheets[i]?.name ?? ''; }
  sheetCount(): number { return this.wb.sheets.length; }
  formula(sheet: number, row: number, col: number): string | null {
    return this.wb.sheets[sheet]?.cells.get(cellKey(row, col))?.f ?? null;
  }

  definedName(name: string, sheet: number): string | null {
    const n = name.toLowerCase();
    const local = this.wb.names.find((d) => d.sheet === sheet && d.name.toLowerCase() === n);
    if (local) return local.formula;
    const global = this.wb.names.find((d) => d.sheet === undefined && d.name.toLowerCase() === n);
    return global ? global.formula : null;
  }

  table(name: string) {
    const n = name.toLowerCase();
    for (let i = 0; i < this.wb.sheets.length; i++) {
      const t = this.wb.sheets[i]!.tables.find((x) => x.name.toLowerCase() === n);
      if (t) return { sheet: i, ...t.range, header: t.headerRow, totals: t.totalsRow, columns: t.columns.map((c) => c.name) };
    }
    return null;
  }

  tableAt(sheet: number, row: number, col: number): string | null {
    const t = this.wb.sheets[sheet]?.tables.find((x) => row >= x.range.r1 && row <= x.range.r2 && col >= x.range.c1 && col <= x.range.c2);
    return t?.name ?? null;
  }

  spillRange(sheet: number, row: number, col: number): RangeRef | null {
    const a = this.spills[sheet]?.anchors.get(cellKey(row, col));
    if (!a) {
      const c = this.wb.sheets[sheet]?.cells.get(cellKey(row, col));
      return c?.f ? ref(sheet, row, col) : null;
    }
    return ref(sheet, row, col, row + a.rows - 1, col + a.cols - 1);
  }

  spillAnchorOf(sheet: number, row: number, col: number): { row: number; col: number } | null {
    const k = this.spills[sheet]?.owner.get(cellKey(row, col));
    return k === undefined ? null : { row: keyRow(k), col: keyCol(k) };
  }

  /** True for cells showing part of another cell's spilled result. */
  isSpilled(sheet: number, row: number, col: number): boolean {
    const k = cellKey(row, col);
    const o = this.spills[sheet]?.owner.get(k);
    return o !== undefined && o !== k;
  }

  rowHidden(sheet: number, row: number): boolean { return !!this.wb.sheets[sheet]?.rows.get(row)?.hidden; }
  rowFiltered(sheet: number, row: number): boolean { const r = this.wb.sheets[sheet]?.rows.get(row); return !!(r?.hidden && r.filtered); }

  isSubtotalCell(sheet: number, row: number, col: number): boolean {
    const f = this.wb.sheets[sheet]?.cells.get(cellKey(row, col))?.f;
    return !!f && /^\s*(SUBTOTAL|AGGREGATE)\s*\(/i.test(f);
  }

  numberFormat(sheet: number, row: number, col: number): string {
    const s = this.wb.sheets[sheet];
    const c = s?.cells.get(cellKey(row, col));
    return (c?.s ? this.wb.styles[c.s]?.numFmt : undefined) ?? 'General';
  }

  pivotData(sheet: number, row: number, col: number, dataField: string, pairs: [string, Scalar][]): Scalar {
    const r = pivotResolver();
    return r ? r(this.wb, sheet, row, col, dataField, pairs) : Err.ref;
  }

  // ---------- Dependencies ----------

  /** Re-reads every formula in the workbook. Call after loading or large structural changes. */
  rebuild(): void {
    this.deps.clear();
    this.cellDependents.clear();
    this.rangeDependents.clear();
    this.volatile.clear();
    this.bounds.clear();
    this.spills = this.wb.sheets.map(() => ({ anchors: new Map(), owner: new Map(), values: new Map() }));
    this.wb.sheets.forEach((s, si) => {
      for (const [k, c] of s.cells) if (c.f) { this.track(si, k, c.f); this.pending.add(idOf(si, k)); }
    });
  }

  private track(sheet: number, key: number, formula: string): void {
    const id = idOf(sheet, key);
    this.untrack(id);
    const d = this.analyse(formula, sheet, keyRow(key), keyCol(key));
    this.deps.set(id, d);
    if (d.volatile) this.volatile.add(id);
    for (const a of d.cells) {
      if (a.r1 === a.r2 && a.c1 === a.c2) {
        const k = `${a.sheet}:${cellKey(a.r1, a.c1)}`;
        let set = this.cellDependents.get(k);
        if (!set) this.cellDependents.set(k, (set = new Set()));
        set.add(id);
        continue;
      }
      let idx = this.rangeDependents.get(a.sheet);
      if (!idx) this.rangeDependents.set(a.sheet, (idx = { cols: new Map(), wide: [] }));
      if (a.c2 - a.c1 + 1 > BUCKET_LIMIT) { idx.wide.push({ ...a, id }); continue; }
      for (let c = a.c1; c <= a.c2; c++) {
        let list = idx.cols.get(c);
        if (!list) idx.cols.set(c, (list = []));
        list.push({ r1: a.r1, r2: a.r2, id });
      }
    }
  }

  private untrack(id: Id): void {
    const d = this.deps.get(id);
    if (!d) return;
    this.deps.delete(id);
    this.volatile.delete(id);
    for (const a of d.cells) {
      if (a.r1 === a.r2 && a.c1 === a.c2) { this.cellDependents.get(`${a.sheet}:${cellKey(a.r1, a.c1)}`)?.delete(id); continue; }
      const idx = this.rangeDependents.get(a.sheet);
      if (!idx) continue;
      if (a.c2 - a.c1 + 1 > BUCKET_LIMIT) { idx.wide = idx.wide.filter((w) => w.id !== id); continue; }
      for (let c = a.c1; c <= a.c2; c++) {
        const list = idx.cols.get(c);
        if (list) idx.cols.set(c, list.filter((x) => x.id !== id));
      }
    }
  }

  /** The cells and ranges a formula reads, found from its text (names and tables expanded). */
  private analyse(formula: string, sheet: number, row: number, col: number, seen = new Set<string>()): Deps {
    const out: Deps = { cells: [], volatile: false };
    let node: Node;
    try { node = parseFormula(formula); } catch { return out; }
    walk(node, (n) => {
      if (n.type === 'ref') {
        if (n.ref.external) return;
        const s1 = n.ref.sheet === undefined ? sheet : this.sheetIndex(n.ref.sheet);
        const s2 = n.ref.sheet2 === undefined ? s1 : this.sheetIndex(n.ref.sheet2);
        if (s1 < 0 || s2 < 0) return;
        for (let s = Math.min(s1, s2); s <= Math.max(s1, s2); s++) out.cells.push({ sheet: s, r1: n.ref.r1, c1: n.ref.c1, r2: n.ref.r2, c2: n.ref.c2 });
      } else if (n.type === 'name') {
        const key = n.name.toLowerCase();
        if (seen.has(key)) return;
        const s = n.sheet ? this.sheetIndex(n.sheet) : sheet;
        const f = s >= 0 ? this.definedName(n.name, s) : null;
        if (f !== null) {
          seen.add(key);
          const d = this.analyse(f, s, row, col, seen);
          out.cells.push(...d.cells);
          out.volatile ||= d.volatile;
        } else {
          const t = this.table(n.name);
          if (t) out.cells.push({ sheet: t.sheet, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
        }
      } else if (n.type === 'table') {
        const name = n.table || this.tableAt(sheet, row, col);
        const t = name ? this.table(name) : null;
        if (t) out.cells.push({ sheet: t.sheet, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
      } else if (n.type === 'call') {
        if (VOLATILE_SPECIAL.has(n.name) || getFunction(n.name)?.volatile) out.volatile = true;
        if (!getFunction(n.name) && !['IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'CHOOSE', 'LET', 'LAMBDA', 'ISOMITTED', 'INDIRECT'].includes(n.name)) {
          const f = this.definedName(n.name, sheet);
          if (f !== null && !seen.has(n.name.toLowerCase())) {
            seen.add(n.name.toLowerCase());
            const d = this.analyse(f, sheet, row, col, seen);
            out.cells.push(...d.cells);
            out.volatile ||= d.volatile;
          }
        }
      } else if (n.type === 'spill') {
        out.volatile = true;
      }
    });
    return out;
  }

  /** Formula cells that read a given cell directly. */
  private dependentsOf(sheet: number, row: number, col: number, into: Set<Id>): void {
    const set = this.cellDependents.get(`${sheet}:${cellKey(row, col)}`);
    if (set) for (const id of set) into.add(id);
    const idx = this.rangeDependents.get(sheet);
    if (!idx) return;
    const list = idx.cols.get(col);
    if (list) for (const x of list) if (row >= x.r1 && row <= x.r2) into.add(x.id);
    for (const w of idx.wide) if (row >= w.r1 && row <= w.r2 && col >= w.c1 && col <= w.c2) into.add(w.id);
  }

  // ---------- Editing ----------

  /** Stores a formula (without "="). Throws FormulaError when it can't be read. */
  setFormula(sheet: number, row: number, col: number, formula: string): void {
    parseFormula(formula);
    const s = this.wb.sheets[sheet]!;
    const key = cellKey(row, col);
    const c: Cell = s.cells.get(key) ?? { v: null };
    c.f = formula;
    delete c.arrayRange;
    delete c.rich;
    s.cells.set(key, c);
    this.track(sheet, key, formula);
    this.markChanged(sheet, row, col);
    this.pending.add(idOf(sheet, key));
  }

  setValue(sheet: number, row: number, col: number, v: Scalar): void {
    const s = this.wb.sheets[sheet]!;
    const key = cellKey(row, col);
    const c = s.cells.get(key);
    if (c?.f) this.untrack(idOf(sheet, key));
    if (c) {
      delete c.f; delete c.dyn; delete c.arrayRange; delete c.rich;
      c.v = v;
      if (v === null && !c.s) s.cells.delete(key);
    } else if (v !== null) s.cells.set(key, { v });
    this.clearSpill(sheet, key);
    this.markChanged(sheet, row, col);
  }

  /** Removes a cell's content and formatting. */
  clearCell(sheet: number, row: number, col: number): void {
    const key = cellKey(row, col);
    if (this.wb.sheets[sheet]?.cells.get(key)?.f) this.untrack(idOf(sheet, key));
    this.wb.sheets[sheet]?.cells.delete(key);
    this.clearSpill(sheet, key);
    this.markChanged(sheet, row, col);
  }

  /** Tells the engine a cell changed outside setFormula/setValue (for example after a paste). */
  markChanged(sheet: number, row: number, col: number): void {
    this.bounds.delete(sheet);
    const owner = this.spills[sheet]?.owner.get(cellKey(row, col));
    if (owner !== undefined && owner !== cellKey(row, col)) this.pending.add(idOf(sheet, owner));
    for (const [id, a] of this.blockedSpills) if (a.sheet === sheet && row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2) this.pending.add(id);
    this.changedCells.push({ sheet, r1: row, c1: col, r2: row, c2: col });
  }

  /** Forgets everything about a sheet's formulas and recalculates it all, after rows or columns move. */
  refreshSheet(sheet: number): void {
    for (const id of [...this.deps.keys()]) if (splitId(id)[0] === sheet) this.untrack(id);
    const sp = this.spills[sheet];
    if (sp) { sp.anchors.clear(); sp.owner.clear(); sp.values.clear(); }
    this.bounds.delete(sheet);
    const s = this.wb.sheets[sheet];
    if (!s) return;
    for (const [k, c] of s.cells) if (c.f) { this.track(sheet, k, c.f); this.pending.add(idOf(sheet, k)); }
    this.changedCells.push({ sheet, r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 });
  }

  // ---------- Calculation ----------

  /** Recalculates everything affected by changes since the last call. Returns the cells whose value changed. */
  recalc(full = false): { sheet: number; row: number; col: number }[] {
    if (this.manual && !full) return this.recalcEntered();
    if (full) { for (const id of this.deps.keys()) this.pending.add(id); this.stale = false; }
    const dirty = new Set<Id>(this.pending);
    for (const id of this.volatile) dirty.add(id);
    const queue: Area[] = this.changedCells;
    this.changedCells = [];
    this.pending.clear();
    for (const a of queue) this.expandArea(a, dirty);
    const changed: { sheet: number; row: number; col: number }[] = [];
    this.circular = [];
    for (let pass = 0; pass < 8 && dirty.size; pass++) {
      const spillChanges: Area[] = [];
      this.closure(dirty);
      this.calculate(dirty, changed, spillChanges);
      dirty.clear();
      for (const a of spillChanges) this.expandArea(a, dirty);
    }
    return changed;
  }

  /** Manual mode: calculates just the formulas that were set, and remembers that their dependents are out of date. */
  private recalcEntered(): { sheet: number; row: number; col: number }[] {
    const entered = new Set<Id>(this.pending);
    const affected = new Set<Id>(entered);
    for (const a of this.changedCells) this.expandArea(a, affected);
    this.changedCells = [];
    this.pending.clear();
    const changed: { sheet: number; row: number; col: number }[] = [];
    if (entered.size) { this.circular = []; this.calculate(entered, changed, []); }
    this.closure(affected);
    if ([...affected].some((id) => !entered.has(id))) this.stale = true;
    return changed;
  }

  private expandArea(a: Area, into: Set<Id>): void {
    if (a.r2 - a.r1 > 4096 || a.c2 - a.c1 > 256) {
      // Huge change (whole sheet): everything on that sheet and everything reading it.
      for (const id of this.deps.keys()) {
        const [s] = splitId(id);
        if (s === a.sheet || this.deps.get(id)!.cells.some((x) => x.sheet === a.sheet)) into.add(id);
      }
      return;
    }
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) this.dependentsOf(a.sheet, r, c, into);
  }

  /** Adds everything downstream of the dirty set. */
  private closure(dirty: Set<Id>): void {
    const stack = [...dirty];
    while (stack.length) {
      const id = stack.pop()!;
      const [s, k] = splitId(id);
      const before = dirty.size;
      const next = new Set<Id>();
      this.dependentsOf(s, keyRow(k), keyCol(k), next);
      const sp = this.spills[s]?.anchors.get(k);
      if (sp) {
        const r0 = keyRow(k), c0 = keyCol(k);
        for (let r = r0; r < r0 + sp.rows; r++) for (let c = c0; c < c0 + sp.cols; c++) this.dependentsOf(s, r, c, next);
      }
      for (const n of next) if (!dirty.has(n)) { dirty.add(n); stack.push(n); }
      void before;
    }
  }

  private calculate(dirty: Set<Id>, changed: { sheet: number; row: number; col: number }[], spillChanges: Area[]): void {
    // Index dirty formulas by sheet and column so range precedents can be found quickly.
    const index = new Map<number, Map<number, number[]>>();
    for (const id of dirty) {
      const [s, k] = splitId(id);
      let cols = index.get(s);
      if (!cols) index.set(s, (cols = new Map()));
      const c = keyCol(k);
      let rows = cols.get(c);
      if (!rows) cols.set(c, (rows = []));
      rows.push(keyRow(k));
    }
    for (const cols of index.values()) for (const rows of cols.values()) rows.sort((a, b) => a - b);

    const state = new Map<Id, 1 | 2>();
    const precedentsIn = (id: Id): Id[] => {
      const d = this.deps.get(id);
      if (!d) return [];
      const out: Id[] = [];
      for (const a of d.cells) {
        const cols = index.get(a.sheet);
        if (!cols) continue;
        const scan = (c: number, rows: number[]) => {
          let lo = 0, hi = rows.length;
          while (lo < hi) { const m = (lo + hi) >> 1; if (rows[m]! < a.r1) lo = m + 1; else hi = m; }
          for (let i = lo; i < rows.length && rows[i]! <= a.r2; i++) out.push(idOf(a.sheet, cellKey(rows[i]!, c)));
        };
        if (a.c2 - a.c1 > cols.size) { for (const [c, rows] of cols) if (c >= a.c1 && c <= a.c2) scan(c, rows); }
        else for (let c = a.c1; c <= a.c2; c++) { const rows = cols.get(c); if (rows) scan(c, rows); }
      }
      return out;
    };

    const order = [...dirty].sort((x, y) => { const [a, ka] = splitId(x), [b, kb] = splitId(y); return a - b || ka - kb; });
    const cyclic = new Set<Id>();
    for (const start of order) {
      if (state.get(start) === 2) continue;
      const stack: { id: Id; expanded: boolean }[] = [{ id: start, expanded: false }];
      while (stack.length) {
        const top = stack[stack.length - 1]!;
        const st = state.get(top.id);
        if (st === 2) { stack.pop(); continue; }
        if (top.expanded) {
          stack.pop();
          state.set(top.id, 2);
          this.compute(top.id, changed, spillChanges);
          continue;
        }
        top.expanded = true;
        state.set(top.id, 1);
        for (const p of precedentsIn(top.id)) {
          if (p === top.id) { cyclic.add(p); continue; }
          const ps = state.get(p);
          if (ps === 2) continue;
          if (ps === 1) { cyclic.add(p); cyclic.add(top.id); continue; }
          stack.push({ id: p, expanded: false });
        }
      }
    }
    if (cyclic.size) {
      for (const id of cyclic) { const [s, k] = splitId(id); this.circular.push({ sheet: s, row: keyRow(k), col: keyCol(k) }); }
      if (this.iterate) {
        const ids = [...cyclic];
        for (let i = 0; i < this.maxIterations; i++) {
          let delta = 0;
          for (const id of ids) {
            const [s, k] = splitId(id);
            const before = this.wb.sheets[s]?.cells.get(k)?.v;
            this.compute(id, changed, spillChanges);
            const after = this.wb.sheets[s]?.cells.get(k)?.v;
            if (typeof before === 'number' && typeof after === 'number') delta = Math.max(delta, Math.abs(after - before));
            else if (before !== after) delta = Infinity;
          }
          if (delta < this.maxChange) break;
        }
      }
    }
  }

  private compute(id: Id, changed: { sheet: number; row: number; col: number }[], spillChanges: Area[]): void {
    const [sheet, key] = splitId(id);
    const s = this.wb.sheets[sheet];
    const cell = s?.cells.get(key);
    if (!s || !cell?.f) return;
    const row = keyRow(key), col = keyCol(key);
    this.blockedSpills.delete(id);
    const raw = this.evaluator.evaluateFormula(cell.f, sheet, row, col);
    let result = this.evaluator.finalize(raw);
    const before = cell.v;
    if (cell.arrayRange) {
      // Legacy array formula: fills its fixed range; extra cells get #N/A.
      const m = isMatrix(result) ? result : [[result]];
      const ar = cell.arrayRange;
      this.setSpill(sheet, key, ar.r2 - ar.r1 + 1, ar.c2 - ar.c1 + 1, (i, j) => {
        const h = m.length, w = m[0]!.length;
        const ii = h === 1 ? 0 : i, jj = w === 1 ? 0 : j;
        return ii < h && jj < w ? m[ii]![jj]! : Err.na;
      }, spillChanges);
      result = m[0]![0]!;
    } else if (isMatrix(result)) {
      const rows = result.length, cols = result[0]!.length;
      const m = result;
      if (row + rows > MAX_ROWS || col + cols > MAX_COLS || this.blocked(sheet, key, rows, cols)) {
        this.clearSpill(sheet, key, spillChanges);
        this.blockedSpills.set(id, { sheet, r1: row, c1: col, r2: row + rows - 1, c2: col + cols - 1 });
        result = new CellError('#SPILL!');
      } else {
        this.setSpill(sheet, key, rows, cols, (i, j) => m[i]![j]!, spillChanges);
        cell.dyn = true;
        result = m[0]![0]!;
      }
    } else {
      this.clearSpill(sheet, key, spillChanges);
    }
    if (this.roundResult && typeof result === 'number') result = this.roundResult(sheet, row, col, result);
    cell.v = result;
    if (!sameValue(before, result)) changed.push({ sheet, row, col });
  }

  private blocked(sheet: number, anchor: number, rows: number, cols: number): boolean {
    const s = this.wb.sheets[sheet]!;
    const sp = this.spills[sheet]!;
    const r0 = keyRow(anchor), c0 = keyCol(anchor);
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        if (!i && !j) continue;
        const k = cellKey(r0 + i, c0 + j);
        const c = s.cells.get(k);
        if (c && (c.f || c.v !== null)) return true;
        const o = sp.owner.get(k);
        if (o !== undefined && o !== anchor) return true;
        if (s.merges.some((m) => r0 + i >= m.r1 && r0 + i <= m.r2 && c0 + j >= m.c1 && c0 + j <= m.c2)) return true;
      }
    }
    return false;
  }

  private setSpill(sheet: number, anchor: number, rows: number, cols: number, at: (i: number, j: number) => Scalar, spillChanges: Area[]): void {
    const sp = this.spills[sheet]!;
    const old = sp.anchors.get(anchor);
    const r0 = keyRow(anchor), c0 = keyCol(anchor);
    let differs = !old || old.rows !== rows || old.cols !== cols;
    if (old && differs) this.clearSpill(sheet, anchor, spillChanges);
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const k = cellKey(r0 + i, c0 + j);
        const v = at(i, j);
        if (!differs && (i || j) && !sameValue(sp.values.get(k) ?? null, v)) differs = true;
        sp.owner.set(k, anchor);
        if (i || j) sp.values.set(k, v);
      }
    }
    sp.anchors.set(anchor, { rows, cols });
    this.bounds.delete(sheet);
    if (differs) spillChanges.push({ sheet, r1: r0, c1: c0, r2: r0 + rows - 1, c2: c0 + cols - 1 });
  }

  private clearSpill(sheet: number, anchor: number, spillChanges?: Area[]): void {
    const sp = this.spills[sheet];
    const old = sp?.anchors.get(anchor);
    if (!sp || !old) return;
    const r0 = keyRow(anchor), c0 = keyCol(anchor);
    for (let i = 0; i < old.rows; i++) for (let j = 0; j < old.cols; j++) {
      const k = cellKey(r0 + i, c0 + j);
      if (sp.owner.get(k) === anchor) { sp.owner.delete(k); sp.values.delete(k); }
    }
    sp.anchors.delete(anchor);
    this.bounds.delete(sheet);
    const area = { sheet, r1: r0, c1: c0, r2: r0 + old.rows - 1, c2: c0 + old.cols - 1 };
    if (spillChanges) spillChanges.push(area); else this.changedCells.push(area);
  }

  /** Evaluates a formula without storing it (for conditional formatting, validation, Goal Seek...). */
  evaluate(formula: string, sheet: number, row: number, col: number): Scalar | Matrix {
    return this.evaluator.finalize(this.evaluator.evaluateFormula(formula, sheet, row, col));
  }

  /** Sheets added or removed: resize per-sheet state and recalculate. */
  sheetsChanged(): void {
    this.rebuild();
  }
}

const sameValue = (a: Scalar | undefined, b: Scalar): boolean => {
  if (a === b) return true;
  if (isError(a) && isError(b)) return a.code === b.code;
  return false;
};

export type { Matrix };
