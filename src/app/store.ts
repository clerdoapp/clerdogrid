import { Engine } from '../engine/engine';
import { FormulaError, nodeText, parseFormula } from '../engine/parser';
import { parseNumber } from '../engine/coerce';
import { CellError, ERROR_CODES, MAX_COLS, MAX_ROWS, type Scalar } from '../engine/types';
import { guessFormat } from '../formats/numfmt';
import {
  cellKey, internStyle, keyCol, keyRow, mergeStyle, newWorkbook,
  type Cell, type CellRange, type CellStyle, type Sheet, type Workbook,
} from '../model/workbook';
import { cloneSheet } from './clone';

export interface Pos { row: number; col: number }

export interface Selection {
  /** Every selected area; the last one holds the active cell. */
  ranges: CellRange[];
  active: Pos;
  /** Fixed corner while extending with Shift. */
  anchor: Pos;
}

export interface Editing {
  row: number;
  col: number;
  text: string;
  /** Started by typing over the cell (arrow keys then commit), rather than F2/double-click (arrows move the caret). */
  replace: boolean;
  /** The formula bar has focus instead of the cell. */
  inBar: boolean;
}

export type Snapshot =
  | { kind: 'cells'; sheet: number; area: CellRange; cells: [number, Cell][] }
  | { kind: 'sheet'; sheet: number; data: Sheet }
  | { kind: 'book'; data: Workbook };

interface UndoStep { label: string; snaps: Snapshot[]; selBefore: Selection; sheetBefore: number; meta?: unknown[] }

/** What a finished change, undo or redo did, for observers such as live collaboration. */
export interface ChangeEvent {
  kind: 'change' | 'undo' | 'redo';
  label: string;
  /** State before each part of the change, in the order taken (the first one covering a cell is the true "before"). */
  snaps: Snapshot[];
  /** The sheet objects at the start, so renamed, moved, added and removed sheets can be told apart. */
  sheetsBefore: Sheet[];
  /** Notes left by operations with store.markOp(), e.g. a row insert. */
  meta: unknown[];
  /** Free space for observers between begin() and end(). */
  extra: Record<string, unknown>;
}
export interface ChangeObserver { begin?: (e: ChangeEvent, scope: { area: CellRange; sheet?: number }[] | 'sheet' | 'book') => void; end: (e: ChangeEvent) => void }

export interface FileState { name: string; path: string | null; dirty: boolean; lastSaved: number | null; handle?: unknown; /** Password the file is encrypted with when saved (kept in memory only). */ password?: string; /** File type Save writes ("xlsx", "xls", "ods", "ots"); xlsx when missing. */ format?: string }

export interface ViewState {
  showFormulas: boolean;
  theme: 'system' | 'light' | 'dark';
  panel: string | null;
  zoom: number;
}

export type Listener = () => void;

const clampPos = (p: Pos): Pos => ({ row: Math.max(0, Math.min(MAX_ROWS - 1, p.row)), col: Math.max(0, Math.min(MAX_COLS - 1, p.col)) });
export const editRangeKey = (sheet: string, r: CellRange): string => `${sheet}|${r.r1},${r.c1}:${r.r2},${r.c2}`;
export const rangeOf = (a: Pos, b: Pos): CellRange => ({ r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) });
export const inRange = (r: CellRange, row: number, col: number): boolean => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;

export class GridStore {
  wb: Workbook;
  engine: Engine;
  sel: Selection = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } };
  editing: Editing | null = null;
  file: FileState = { name: 'Book1', path: null, dirty: false, lastSaved: null };
  view: ViewState = { showFormulas: false, theme: 'system', panel: 'home', zoom: 1 };
  /** Where the active cell goes after Enter (File > Options). */
  enterDirection: 'down' | 'right' | 'up' | 'left' | 'none' = 'down';
  /** Formulas are shown and typed with R1C1 references. */
  r1c1 = false;
  toast: string | null = null;
  /** Range being copied or cut, shown with the moving border. */
  marquee: { sheet: number; range: CellRange; cut: boolean } | null = null;
  dialog: { id: string; data?: unknown } | null = null;
  menu: { x: number; y: number; kind: 'cell' | 'row' | 'col' | 'sheet'; data?: unknown } | null = null;
  /** Increments whenever cells or layout change, so the grid knows to redraw. */
  layoutVersion = 0;
  private version = 0;
  private listeners = new Set<Listener>();
  private undoStack: UndoStep[] = [];
  private redoStack: UndoStep[] = [];
  private open: UndoStep | null = null;
  private toastTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(wb: Workbook = newWorkbook()) {
    this.wb = wb;
    this.engine = new Engine(wb);
    this.engine.recalc(true);
  }

  // ---------- Subscription ----------

  subscribe = (l: Listener): (() => void) => { this.listeners.add(l); return () => this.listeners.delete(l); };
  getVersion = (): number => this.version;
  emit(): void { this.version++; for (const l of this.listeners) l(); }
  touch(): void { this.layoutVersion++; this.emit(); }

  notify(text: string): void {
    this.toast = text;
    if (this.toastTimer) clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => { this.toast = null; this.emit(); }, 3200);
    this.emit();
  }

  // ---------- Workbook ----------

  get sheetIndex(): number { return this.wb.activeSheet; }
  get sheet(): Sheet { return this.wb.sheets[this.wb.activeSheet]!; }

  load(wb: Workbook, name: string, path: string | null = null, handle?: unknown): void {
    if (!wb.sheets.length) wb.sheets.push(newWorkbook().sheets[0]!);
    if (!wb.styles.length) wb.styles.push({});
    this.wb = wb;
    this.wb.activeSheet = Math.min(Math.max(0, wb.activeSheet), wb.sheets.length - 1);
    if (this.sheet.state !== 'visible') this.wb.activeSheet = Math.max(0, wb.sheets.findIndex((s) => s.state === 'visible'));
    this.engine = new Engine(wb);
    if (wb.calc?.iterate) { this.engine.iterate = true; this.engine.maxIterations = wb.calc.maxIterations ?? 100; this.engine.maxChange = wb.calc.maxChange ?? 0.001; }
    this.engine.recalc(true);
    this.undoStack = []; this.redoStack = [];
    const ext = /\.([a-z]+)$/i.exec(name)?.[1]?.toLowerCase();
    const format = ext === 'xls' || ext === 'xlt' ? 'xls' : ext === 'ods' || ext === 'ots' || ext === 'xlsb' ? ext : undefined;
    this.file = { name: name.replace(/\.(xlsx|xlsm|xltx|xltm|xlam|xlsb|xls|xlt|ods|ots|csv|tsv|txt|prn)$/i, ''), path, dirty: false, lastSaved: null, handle, ...(format ? { format } : {}) };
    const s = this.sheet.view.selection;
    const start = s ? { row: s.r1, col: s.c1 } : { row: 0, col: 0 };
    this.sel = { ranges: [s ?? { r1: 0, c1: 0, r2: 0, c2: 0 }], active: start, anchor: start };
    this.editing = null;
    this.marquee = null;
    this.touch();
  }

  setActiveSheet(i: number): void {
    if (i === this.wb.activeSheet || !this.wb.sheets[i]) return;
    this.commitEdit();
    this.sheet.view.selection = this.sel.ranges.at(-1);
    this.wb.activeSheet = i;
    const s = this.sheet.view.selection;
    const start = s ? { row: s.r1, col: s.c1 } : { row: 0, col: 0 };
    this.sel = { ranges: [s ?? { r1: 0, c1: 0, r2: 0, c2: 0 }], active: start, anchor: start };
    this.touch();
  }

  cell(row: number, col: number, sheet = this.sheetIndex): Cell | undefined {
    return this.wb.sheets[sheet]?.cells.get(cellKey(row, col));
  }

  value(row: number, col: number, sheet = this.sheetIndex): Scalar {
    return this.engine.getValue(sheet, row, col);
  }

  style(row: number, col: number, sheet = this.sheetIndex): CellStyle {
    const s = this.wb.sheets[sheet]!;
    const c = s.cells.get(cellKey(row, col));
    const id = c?.s ?? s.rows.get(row)?.s ?? s.cols.get(col)?.s ?? 0;
    return this.wb.styles[id] ?? {};
  }

  /** Text shown in the formula bar / in-cell editor for a cell. */
  inputText(row: number, col: number): string {
    const c = this.cell(row, col);
    if (c?.f) return `=${this.shownFormula(c.f, row, col)}`;
    if (this.engine.isSpilled(this.sheetIndex, row, col)) {
      const a = this.engine.spillAnchorOf(this.sheetIndex, row, col)!;
      const f = this.cell(a.row, a.col)?.f;
      return f ? `=${this.shownFormula(f, a.row, a.col)}` : '';
    }
    const v = c?.v ?? null;
    if (v === null) return '';
    if (v instanceof CellError) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return /^[=+\-]|^\d/.test(v) && parseNumber(v) !== null ? `'${v}` : v;
    const fmt = this.style(row, col).numFmt;
    return localizeNumberText(editableNumber(v, fmt), fmt);
  }

  /** Formula text as the person sees it: A1 as stored, or converted to R1C1. */
  shownFormula(f: string, row: number, col: number): string { return localizeFormula(this.r1c1 ? toR1C1(f, row, col) : f); }

  /** Moves the active cell after Enter (Shift reverses), following the Enter direction setting. */
  moveAfterEnter(shift: boolean): void {
    const d = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1], none: [0, 0] }[this.enterDirection] as [number, number];
    if (!d[0] && !d[1]) return;
    this.move(shift ? -d[0] : d[0], shift ? -d[1] : d[1], true);
  }

  // ---------- Undo ----------

  /**
   * Runs a change as one undo step. `scope` says what to snapshot first: an area, a whole sheet, or
   * the whole book (for sheet-level changes).
   */
  change(label: string, scope: { area: CellRange; sheet?: number }[] | 'sheet' | 'book', fn: () => void): void {
    // Changes from other people in a live session: no undo step, nothing reported back.
    if (this.remote) { fn(); return; }
    const outer = !this.open;
    if (outer && this.readOnly) { this.notify(T('This workbook is read-only.')); return; }
    let event: ChangeEvent | null = null;
    if (outer) {
      this.open = { label, snaps: [], selBefore: structuredClone(this.sel), sheetBefore: this.sheetIndex };
      if (this.observers.size) {
        event = { kind: 'change', label, snaps: this.open.snaps, sheetsBefore: [...this.wb.sheets], meta: [], extra: {} };
        for (const o of this.observers) o.begin?.(event, scope);
      }
    }
    const step = this.open!;
    if (scope === 'book') step.snaps.push({ kind: 'book', data: cloneBook(this.wb) });
    else if (scope === 'sheet') step.snaps.push({ kind: 'sheet', sheet: this.sheetIndex, data: cloneSheet(this.sheet) });
    else for (const s of scope) step.snaps.push(this.snapCells(s.sheet ?? this.sheetIndex, s.area));
    try {
      fn();
    } finally {
      if (outer) {
        this.open = null;
        this.undoStack.push(step);
        if (this.undoStack.length > 200) this.undoStack.shift();
        this.redoStack = [];
        this.file.dirty = true;
        this.engine.recalc();
        if (event) { event.meta = step.meta ?? []; for (const o of this.observers) o.end(event); }
        this.touch();
      }
    }
  }

  // ---------- Collaboration hooks ----------

  private observers = new Set<ChangeObserver>();
  /** True while applying someone else's change: change() skips undo and observers. */
  remote = false;
  /** Blocks every local change (a read-only shared copy or a view-only session). */
  readOnly = false;

  observe(o: ChangeObserver): () => void { this.observers.add(o); return () => this.observers.delete(o); }

  /** Lets an operation describe what it did (e.g. rows inserted) for observers; part of the open undo step. */
  markOp(meta: unknown): void { if (this.open && !this.remote) (this.open.meta ??= []).push(meta); }

  /** Applies a change made by someone else: no undo step, not reported to observers, then recalculates. */
  applyRemote(fn: () => void): void {
    const was = this.remote;
    this.remote = true;
    try { fn(); } finally {
      this.remote = was;
      if (!was) { this.file.dirty = true; this.engine.recalc(); this.touch(); }
    }
  }

  /**
   * Visits every step in the undo history (newest first), then the redo history (next redo first), so its snapshots
   * can be patched after someone else's change and undo doesn't revert their edits. A visitor can keep state between
   * steps (positions move through the steps in between). Returning false drops that step and every step behind it.
   */
  rebaseHistory(visit: (step: { snaps: Snapshot[]; meta: readonly unknown[]; sel: Selection }, stack: 'undo' | 'redo') => boolean): void {
    for (const [stack, kind] of [[this.undoStack, 'undo'], [this.redoStack, 'redo']] as const) {
      for (let i = stack.length - 1; i >= 0; i--) {
        const st = stack[i]!;
        const view = { snaps: st.snaps, meta: st.meta ?? [], sel: st.selBefore };
        const keep = visit(view, kind);
        st.selBefore = view.sel;
        if (!keep) { stack.splice(0, i + 1); break; }
      }
    }
  }

  private snapCells(sheet: number, area: CellRange): Snapshot {
    const s = this.wb.sheets[sheet]!;
    const cells: [number, Cell][] = [];
    const big = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1) > s.cells.size;
    if (big) { for (const [k, c] of s.cells) if (inRange(area, keyRow(k), keyCol(k))) cells.push([k, structuredClone(c)]); }
    else for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) {
      const cell = s.cells.get(cellKey(r, c));
      if (cell) cells.push([cellKey(r, c), structuredClone(cell)]);
    }
    return { kind: 'cells', sheet, area: { ...area }, cells };
  }

  private restore(snaps: Snapshot[]): Snapshot[] {
    const inverse: Snapshot[] = [];
    for (const snap of [...snaps].reverse()) {
      if (snap.kind === 'book') {
        inverse.push({ kind: 'book', data: cloneBook(this.wb) });
        const data = cloneBook(snap.data);
        this.wb.sheets = data.sheets; this.wb.names = data.names; this.wb.styles = data.styles; this.wb.activeSheet = data.activeSheet;
        this.wb.pivotCaches = data.pivotCaches; this.wb.slicerCaches = data.slicerCaches;
        this.wb.protection = data.protection; this.wb.customViews = data.customViews; this.wb.calc = data.calc; this.wb.date1904 = data.date1904;
        this.wb.queries = data.queries;
        this.engine.sheetsChanged();
      } else if (snap.kind === 'sheet') {
        inverse.push({ kind: 'sheet', sheet: snap.sheet, data: cloneSheet(this.wb.sheets[snap.sheet]!) });
        this.wb.sheets[snap.sheet] = cloneSheet(snap.data);
        this.engine.sheetsChanged();
      } else {
        inverse.push(this.snapCells(snap.sheet, snap.area));
        const s = this.wb.sheets[snap.sheet]!;
        const keys: number[] = [];
        for (const k of s.cells.keys()) if (inRange(snap.area, keyRow(k), keyCol(k))) keys.push(k);
        for (const k of keys) {
          const c = s.cells.get(k)!;
          if (c.f) this.engine.clearCell(snap.sheet, keyRow(k), keyCol(k)); else s.cells.delete(k);
          this.engine.markChanged(snap.sheet, keyRow(k), keyCol(k));
        }
        for (const [k, c] of snap.cells) {
          const copy = structuredClone(c);
          if (copy.f) {
            const v = copy.v;
            this.engine.setFormula(snap.sheet, keyRow(k), keyCol(k), copy.f);
            const placed = s.cells.get(k)!;
            Object.assign(placed, copy, { v });
          } else {
            s.cells.set(k, copy);
            this.engine.markChanged(snap.sheet, keyRow(k), keyCol(k));
          }
        }
      }
    }
    return inverse;
  }

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  get undoLabel(): string | undefined { return this.undoStack.at(-1)?.label; }

  undo(): void { this.swap(this.undoStack, this.redoStack); }
  redo(): void { this.swap(this.redoStack, this.undoStack); }

  /** Undo and redo are the same move: restore a step's snapshots and file the inverse on the other stack. */
  private swap(from: UndoStep[], to: UndoStep[]): void {
    this.commitEdit();
    const step = from.pop();
    if (!step) return;
    if (this.readOnly) { from.push(step); return; }
    const sel = structuredClone(this.sel), sheet = this.sheetIndex;
    const sheetsBefore = [...this.wb.sheets];
    const inverse = this.restore(step.snaps);
    to.push({ label: step.label, snaps: inverse, selBefore: sel, sheetBefore: sheet, meta: step.meta });
    this.wb.activeSheet = Math.min(step.sheetBefore, this.wb.sheets.length - 1);
    this.sel = step.selBefore;
    this.engine.recalc();
    if (this.observers.size) {
      const e: ChangeEvent = { kind: from === this.undoStack ? 'undo' : 'redo', label: step.label, snaps: inverse, sheetsBefore, meta: [], extra: {} };
      for (const o of this.observers) o.end(e);
    }
    this.file.dirty = true;
    this.touch();
  }

  // ---------- Selection ----------

  get range(): CellRange { return this.sel.ranges.at(-1)!; }

  /** Rows shown in another order by a sheet view (display row → data row), per sheet. */
  rowOrders = new WeakMap<Sheet, RowOrder>();
  rowOrderOf(s: Sheet = this.sheet): RowOrder | null { return this.rowOrders.get(s) ?? null; }
  /** Cells the person may select (sheet protection can forbid locked or unlocked cells). */
  selectable: ((row: number, col: number) => boolean) | null = null;
  /** Ranges before the area being extended with Shift (a sheet view's row order can split one area into several). */
  private extendBase = 0;

  select(pos: Pos, extend = false, add = false): void {
    pos = clampPos(pos);
    const order = this.rowOrderOf();
    if (!extend && this.selectable && !this.selectable(pos.row, pos.col)) { this.notify(T('That cell is protected and can\'t be selected.')); return; }
    if (extend && order) {
      const a = this.sel.anchor;
      const runs = order.dataRuns(order.toDisplay(a.row), order.toDisplay(pos.row), Math.min(a.col, pos.col), Math.max(a.col, pos.col));
      // The run holding the active cell goes last, as the active area.
      const i = runs.findIndex((r) => inRange(r, this.sel.active.row, this.sel.active.col));
      if (i >= 0) runs.push(...runs.splice(i, 1));
      this.sel = { ...this.sel, ranges: [...this.sel.ranges.slice(0, Math.min(this.extendBase, this.sel.ranges.length - 1)), ...runs], active: this.sel.active };
    } else if (extend) {
      const r = this.expandForMerges(rangeOf(this.sel.anchor, pos));
      this.sel = { ...this.sel, ranges: [...this.sel.ranges.slice(0, Math.min(this.extendBase, this.sel.ranges.length - 1)), r], active: this.sel.active };
    } else {
      this.extendBase = add ? this.sel.ranges.length : 0;
      const m = this.mergeAt(pos.row, pos.col);
      const r = m ?? { r1: pos.row, c1: pos.col, r2: pos.row, c2: pos.col };
      const start = m ? { row: m.r1, col: m.c1 } : pos;
      this.sel = { ranges: add ? [...this.sel.ranges, r] : [r], active: start, anchor: start };
    }
    this.emit();
  }

  selectRange(r: CellRange, active?: Pos): void {
    const a = active ?? { row: r.r1, col: r.c1 };
    if (this.selectable && !this.selectable(a.row, a.col)) { this.notify(T('That cell is protected and can\'t be selected.')); return; }
    this.extendBase = 0;
    this.sel = { ranges: [r], active: a, anchor: a };
    this.emit();
  }

  /** Move the active cell (arrow keys, Enter, Tab). Inside a multi-cell selection, Enter/Tab cycle within it. */
  move(dRow: number, dCol: number, withinSelection = false): void {
    const r = this.range;
    const multi = r.r1 !== r.r2 || r.c1 !== r.c2;
    if (withinSelection && multi) {
      let { row, col } = this.sel.active;
      if (dRow) { row += dRow; if (row > r.r2) { row = r.r1; col = col + 1 > r.c2 ? r.c1 : col + 1; } else if (row < r.r1) { row = r.r2; col = col - 1 < r.c1 ? r.c2 : col - 1; } }
      if (dCol) { col += dCol; if (col > r.c2) { col = r.c1; row = row + 1 > r.r2 ? r.r1 : row + 1; } else if (col < r.c1) { col = r.c2; row = row - 1 < r.r1 ? r.r2 : row - 1; } }
      this.sel = { ...this.sel, active: { row, col } };
      this.emit();
      return;
    }
    let { row, col } = this.sel.active;
    const m = this.mergeAt(row, col);
    if (m) { if (dRow > 0) row = m.r2; if (dCol > 0) col = m.c2; }
    // Rows move in display order (a sheet view may show them sorted).
    const order = this.rowOrderOf();
    const toD = (p: number) => (order ? order.toData(p) : p);
    let p = (order ? order.toDisplay(row) : row) + dRow;
    col += dCol;
    // Skip hidden rows and columns, and cells protection doesn't let anyone select.
    const s = this.sheet;
    const skipHidden = () => {
      while (dRow && p > 0 && p < MAX_ROWS - 1 && s.rows.get(toD(p))?.hidden) p += Math.sign(dRow);
      while (dCol && col > 0 && col < MAX_COLS - 1 && s.cols.get(col)?.hidden) col += Math.sign(dCol);
    };
    skipHidden();
    const ok = () => { const c = clampPos({ row: p, col }); return !this.selectable || this.selectable(toD(c.row), c.col); };
    if (this.selectable && (dRow || dCol)) {
      for (let tries = 0; !ok() && tries < 5000; tries++) {
        if (p + dRow < 0 || p + dRow >= MAX_ROWS || col + dCol < 0 || col + dCol >= MAX_COLS) return;
        p += dRow; col += dCol;
        skipHidden();
      }
      if (!ok()) return;
    }
    const c = clampPos({ row: p, col });
    this.select({ row: toD(c.row), col: c.col });
  }

  /** Ctrl+arrow: jump to the edge of the current block of data. */
  edge(pos: Pos, dRow: number, dCol: number): Pos {
    const order = this.rowOrderOf();
    if (order) {
      // Walk in display order, then report the data row.
      const shown = this.edgeIn({ row: order.toDisplay(pos.row), col: pos.col }, dRow, dCol, (r) => order.toData(r));
      return { row: order.toData(shown.row), col: shown.col };
    }
    return this.edgeIn(pos, dRow, dCol, (r) => r);
  }

  private edgeIn(pos: Pos, dRow: number, dCol: number, toData: (r: number) => number): Pos {
    const has = (r: number, c: number) => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS && this.value(toData(r), c) !== null;
    let { row, col } = pos;
    const b = this.engine.usedBounds(this.sheetIndex);
    const limitR = dRow > 0 ? MAX_ROWS - 1 : 0, limitC = dCol > 0 ? MAX_COLS - 1 : 0;
    const next = () => { row += dRow; col += dCol; };
    const atLimit = () => (dRow && row === limitR) || (dCol && col === limitC);
    if (atLimit()) return pos;
    const startFilled = has(row, col);
    next();
    if (startFilled && has(row, col)) {
      while (!atLimit() && has(row + dRow, col + dCol)) next();
      return { row, col };
    }
    while (!atLimit() && !has(row, col)) {
      if ((dRow > 0 && row > b.rows) || (dCol > 0 && col > b.cols)) return { row: dRow ? limitR : row, col: dCol ? limitC : col };
      next();
    }
    return { row, col };
  }

  mergeAt(row: number, col: number, sheet = this.sheetIndex): CellRange | undefined {
    return this.wb.sheets[sheet]!.merges.find((m) => inRange(m, row, col));
  }

  private expandForMerges(r: CellRange): CellRange {
    let out = { ...r };
    for (let changed = true; changed;) {
      changed = false;
      for (const m of this.sheet.merges) {
        const hit = m.r1 <= out.r2 && m.r2 >= out.r1 && m.c1 <= out.c2 && m.c2 >= out.c1;
        if (hit && (m.r1 < out.r1 || m.r2 > out.r2 || m.c1 < out.c1 || m.c2 > out.c2)) {
          out = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
          changed = true;
        }
      }
    }
    return out;
  }

  // ---------- Editing ----------

  beginEdit(text?: string, inBar = false): void {
    const { row, col } = this.sel.active;
    if (this.readOnly) { this.notify(T('This workbook is read-only.')); return; }
    if (this.isLocked(row, col)) { this.notify(T('This cell is protected. Unprotect the sheet to change it.')); return; }
    const anchor = this.engine.spillAnchorOf(this.sheetIndex, row, col);
    if (anchor && (anchor.row !== row || anchor.col !== col) && text === undefined) {
      this.select(anchor);
    }
    this.editing = { row: this.sel.active.row, col: this.sel.active.col, text: text ?? this.inputText(this.sel.active.row, this.sel.active.col), replace: text !== undefined, inBar };
    this.emit();
  }

  setEditText(text: string): void {
    if (!this.editing) return;
    this.editing = { ...this.editing, text };
    this.emit();
  }

  cancelEdit(): void {
    if (!this.editing) return;
    this.editing = null;
    this.emit();
  }

  /** Stores what was typed. `fill` puts the same entry into every selected cell (Ctrl+Enter). Returns false if the formula is broken. */
  commitEdit(fill = false, arrayFormula = false): boolean {
    const e = this.editing;
    if (!e) return true;
    let text = e.text;
    if (text.startsWith('=')) text = `=${delocalizeFormula(text.slice(1))}`;
    if (this.r1c1 && text.startsWith('=')) text = `=${fromR1C1(text.slice(1), e.row, e.col)}`;
    if (text.startsWith('=')) {
      const fixed = autoCloseParens(text.slice(1));
      try { parseFormula(fixed); text = `=${fixed}`; } catch (err) {
        if (err instanceof FormulaError) { this.notify(T('There\'s a problem with this formula. {message}', { message: err.message })); return false; }
        throw err;
      }
    }
    const err = this.validateEntry(e.row, e.col, text);
    if (err) { this.notify(err); return false; }
    this.editing = null;
    const r = this.range;
    if (arrayFormula && text.startsWith('=')) {
      this.change(N('Array formula'), [{ area: r }], () => this.setArrayFormula(r, text.slice(1)));
      return true;
    }
    const targets = fill ? this.sel.ranges : [{ r1: e.row, c1: e.col, r2: e.row, c2: e.col }];
    this.change(N('Typing'), targets.map((area) => ({ area })), () => {
      for (const t of targets) for (let row = t.r1; row <= t.r2; row++) for (let col = t.c1; col <= t.c2; col++) {
        const entry = text.startsWith('=') && (row !== e.row || col !== e.col) ? `=${shiftText(text.slice(1), row - e.row, col - e.col)}` : text;
        this.enter(row, col, entry);
      }
    });
    return true;
  }

  private setArrayFormula(r: CellRange, formula: string): void {
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) if (row !== r.r1 || col !== r.c1) this.engine.setValue(this.sheetIndex, row, col, null);
    this.engine.setFormula(this.sheetIndex, r.r1, r.c1, canonical(formula));
    this.cell(r.r1, r.c1)!.arrayRange = { ...r };
  }

  /** Interprets typed text: formula, number (with a format guessed from how it was typed), TRUE/FALSE, error, or text. */
  enter(row: number, col: number, text: string, sheet = this.sheetIndex): void {
    const s = this.wb.sheets[sheet]!;
    if (text === '') { this.engine.setValue(sheet, row, col, null); return; }
    if (text.startsWith('=') && text.length > 1) {
      this.engine.setFormula(sheet, row, col, canonical(text.slice(1)));
      const c = s.cells.get(cellKey(row, col))!;
      if (!c.s) {
        const fmt = formulaFormat(text.slice(1), (r, cc) => this.style(r, cc, sheet).numFmt);
        if (fmt) c.s = internStyle(this.wb, { numFmt: fmt });
      }
      return;
    }
    if (text.startsWith("'")) { this.engine.setValue(sheet, row, col, text.slice(1)); return; }
    const up = text.trim().toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') { this.engine.setValue(sheet, row, col, up === 'TRUE'); return; }
    const errCode = ERROR_CODES.find((c) => c === up);
    if (errCode) { this.engine.setValue(sheet, row, col, new CellError(errCode)); return; }
    const plain = delocalizeInput(text);
    const n = parseNumber(plain, this.wb.date1904);
    if (n !== null && !/^0\d+$/.test(plain.trim())) {
      this.engine.setValue(sheet, row, col, n);
      const cur = this.style(row, col, sheet);
      const fmt = guessFormat(plain);
      if (fmt && (!cur.numFmt || cur.numFmt === 'General')) {
        const c = s.cells.get(cellKey(row, col))!;
        c.s = internStyle(this.wb, mergeStyle(cur, { numFmt: fmt }));
      }
      return;
    }
    this.engine.setValue(sheet, row, col, text.length > 32767 ? text.slice(0, 32767) : text);
  }

  // ---------- Formatting ----------

  /** Merges a style change into every selected cell (and whole rows/columns when they're selected). */
  applyStyle(patch: CellStyle | ((s: CellStyle) => CellStyle), label = N('Format')): void {
    const f = typeof patch === 'function' ? patch : (s: CellStyle) => mergeStyle(s, patch);
    const whole = this.sel.ranges.some((r) => (r.r1 === 0 && r.r2 === MAX_ROWS - 1) || (r.c1 === 0 && r.c2 === MAX_COLS - 1));
    this.change(label, whole ? 'sheet' : this.sel.ranges.map((area) => ({ area })), () => {
      for (const r of this.sel.ranges) this.styleRange(r, f);
    });
  }

  styleRange(r: CellRange, f: (s: CellStyle) => CellStyle, sheet = this.sheetIndex): void {
    const s = this.wb.sheets[sheet]!;
    const wholeCols = r.r1 === 0 && r.r2 === MAX_ROWS - 1;
    const wholeRows = r.c1 === 0 && r.c2 === MAX_COLS - 1;
    if (wholeCols) for (let c = r.c1; c <= r.c2 && c - r.c1 < 2000; c++) {
      const info = s.cols.get(c) ?? {};
      info.s = internStyle(this.wb, f(this.wb.styles[info.s ?? 0] ?? {}));
      s.cols.set(c, info);
    }
    if (wholeRows && !wholeCols) for (let rr = r.r1; rr <= r.r2 && rr - r.r1 < 100000; rr++) {
      const info = s.rows.get(rr) ?? {};
      info.s = internStyle(this.wb, f(this.wb.styles[info.s ?? 0] ?? {}));
      s.rows.set(rr, info);
    }
    if (wholeCols || wholeRows) {
      for (const [k, c] of s.cells) if (inRange(r, keyRow(k), keyCol(k))) c.s = internStyle(this.wb, f(this.wb.styles[c.s ?? 0] ?? {}));
      return;
    }
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
      const k = cellKey(row, col);
      const c = s.cells.get(k);
      const base = this.style(row, col, sheet);
      const id = internStyle(this.wb, f(base));
      if (c) c.s = id;
      else if (id) s.cells.set(k, { v: null, s: id });
    }
  }

  /** Allow-edit ranges unlocked with their password in this session ("Sheet|A1:B2"). */
  unlockedRanges = new Set<string>();

  isLocked(row: number, col: number): boolean {
    if (!this.sheet.protection?.sheet) return false;
    if (this.sheet.editRanges?.some((e) => inRange(e.range, row, col) && (!e.password || this.unlockedRanges.has(editRangeKey(this.sheet.name, e.range))))) return false;
    return this.style(row, col).protection?.locked !== false;
  }

  private validateEntry(row: number, col: number, text: string): string | null {
    const dv = this.sheet.validations.find((v) => v.ranges.some((r) => inRange(r, row, col)));
    if (!dv || dv.type === 'any' || text.startsWith('=')) return null;
    if (text === '' && dv.allowBlank !== false) return null;
    const ok = checkValidation(this, dv, row, col, text);
    if (ok || dv.error?.style === 'information' || dv.error?.style === 'warning') return null;
    return dv.error?.text || T('This value doesn\'t match the data validation rules defined for this cell.');
  }
}

// ---------- Helpers ----------

import { checkValidation } from './validation';
import { shiftFormula } from '../engine/refactor';
import { fromR1C1, toR1C1 } from '../engine/r1c1';
import { delocalizeFormula, delocalizeInput, editDateText, localizeFormula, localizeNumberText } from './locale';
import type { RowOrder } from './rowOrder';

const shiftText = (f: string, dr: number, dc: number): string => shiftFormula(f, dr, dc);

/** Upper-cases function names and cell references the way spreadsheets show them, keeping text as typed. */
export const canonical = (formula: string): string => {
  try { return nodeText(parseFormula(formula)); } catch { return formula; }
};

const autoCloseParens = (f: string): string => {
  let depth = 0, inStr = false;
  for (const ch of f) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === '(') depth++;
    else if (!inStr && ch === ')') depth--;
  }
  return depth > 0 && !inStr ? f + ')'.repeat(depth) : f;
};

/** Spreadsheets give =A1 the format of A1 when A1 is a date/currency/percent, and =TODAY() a date format. */
const formulaFormat = (formula: string, fmtOf: (r: number, c: number) => string | undefined): string | undefined => {
  const f = formula.trim().toUpperCase();
  if (/^(TODAY)\(\)$/.test(f)) return 'm/d/yyyy';
  if (/^NOW\(\)$/.test(f)) return 'm/d/yyyy h:mm';
  if (/^(DATE|EDATE|EOMONTH|WORKDAY|WORKDAY\.INTL|DATEVALUE)\(/.test(f)) return 'm/d/yyyy';
  if (/^(TIME|TIMEVALUE)\(/.test(f)) return 'h:mm AM/PM';
  const m = /^(?:SUM|AVERAGE|MIN|MAX)?\(?\$?([A-Z]{1,3})\$?(\d+)/.exec(f);
  if (m) {
    let col = 0;
    for (const ch of m[1]!) col = col * 26 + ch.charCodeAt(0) - 64;
    const fmt = fmtOf(+m[2]! - 1, col - 1);
    if (fmt && fmt !== 'General' && fmt !== '@') return fmt;
  }
  return undefined;
};

const editableNumber = (v: number, fmt: string | undefined): string => {
  if (fmt && /%/.test(fmt.replace(/"[^"]*"/g, ''))) return `${+(v * 100).toPrecision(15)}%`;
  if (fmt && /[dmyhs]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/^[#0.,]+$/.test(fmt)) {
    return dateText(v, fmt);
  }
  return String(+v.toPrecision(15));
};

import { fromSerial, timeOf } from '../engine/date';
import { N, T } from '../ui/i18n';
const dateText = (v: number, fmt: string): string => {
  const hasDate = /[dy]|m{3,}/i.test(fmt.replace(/"[^"]*"/g, '')) || v >= 1;
  const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"/g, '')) || v % 1 !== 0;
  const parts: string[] = [];
  if (hasDate && v >= 1) { const d = fromSerial(v); parts.push(editDateText(d.year, d.month, d.day)); }
  if (hasTime) {
    const t = timeOf(v);
    const h12 = t.hour % 12 || 12;
    parts.push(`${h12}:${String(t.minute).padStart(2, '0')}:${String(t.second).padStart(2, '0')} ${t.hour < 12 ? 'AM' : 'PM'}`);
  }
  return parts.join(' ') || String(v);
};

export const cloneBook = (wb: Workbook): Workbook => ({
  ...wb, sheets: wb.sheets.map(cloneSheet), names: structuredClone(wb.names), styles: structuredClone(wb.styles),
  // Pivot caches are replaced, never changed in place, so copying the list is enough.
  pivotCaches: wb.pivotCaches && [...wb.pivotCaches], slicerCaches: wb.slicerCaches && structuredClone(wb.slicerCaches),
  queries: wb.queries && structuredClone(wb.queries),
});
