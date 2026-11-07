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

type Snapshot =
  | { kind: 'cells'; sheet: number; area: CellRange; cells: [number, Cell][] }
  | { kind: 'sheet'; sheet: number; data: Sheet }
  | { kind: 'book'; data: Workbook };

interface UndoStep { label: string; snaps: Snapshot[]; selBefore: Selection; sheetBefore: number }

export interface FileState { name: string; path: string | null; dirty: boolean; lastSaved: number | null; handle?: unknown }

export interface ViewState {
  showFormulas: boolean;
  theme: 'system' | 'light' | 'dark';
  panel: string | null;
  zoom: number;
}

export type Listener = () => void;

const clampPos = (p: Pos): Pos => ({ row: Math.max(0, Math.min(MAX_ROWS - 1, p.row)), col: Math.max(0, Math.min(MAX_COLS - 1, p.col)) });
export const rangeOf = (a: Pos, b: Pos): CellRange => ({ r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) });
export const inRange = (r: CellRange, row: number, col: number): boolean => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;

export class GridStore {
  wb: Workbook;
  engine: Engine;
  sel: Selection = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } };
  editing: Editing | null = null;
  file: FileState = { name: 'Book1', path: null, dirty: false, lastSaved: null };
  view: ViewState = { showFormulas: false, theme: 'system', panel: 'home', zoom: 1 };
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
    this.file = { name: name.replace(/\.(xlsx|xlsm|xltx|csv|tsv|txt)$/i, ''), path, dirty: false, lastSaved: null, handle };
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
    if (c?.f) return `=${c.f}`;
    if (this.engine.isSpilled(this.sheetIndex, row, col)) {
      const a = this.engine.spillAnchorOf(this.sheetIndex, row, col)!;
      const f = this.cell(a.row, a.col)?.f;
      return f ? `=${f}` : '';
    }
    const v = c?.v ?? null;
    if (v === null) return '';
    if (v instanceof CellError) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return /^[=+\-]|^\d/.test(v) && parseNumber(v) !== null ? `'${v}` : v;
    return editableNumber(v, this.style(row, col).numFmt);
  }

  // ---------- Undo ----------

  /**
   * Runs a change as one undo step. `scope` says what to snapshot first: an area, a whole sheet, or
   * the whole book (for sheet-level changes).
   */
  change(label: string, scope: { area: CellRange; sheet?: number }[] | 'sheet' | 'book', fn: () => void): void {
    const outer = !this.open;
    if (outer) this.open = { label, snaps: [], selBefore: structuredClone(this.sel), sheetBefore: this.sheetIndex };
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
        this.touch();
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
    const sel = structuredClone(this.sel), sheet = this.sheetIndex;
    const inverse = this.restore(step.snaps);
    to.push({ label: step.label, snaps: inverse, selBefore: sel, sheetBefore: sheet });
    this.wb.activeSheet = Math.min(step.sheetBefore, this.wb.sheets.length - 1);
    this.sel = step.selBefore;
    this.engine.recalc();
    this.file.dirty = true;
    this.touch();
  }

  // ---------- Selection ----------

  get range(): CellRange { return this.sel.ranges.at(-1)!; }

  select(pos: Pos, extend = false, add = false): void {
    pos = clampPos(pos);
    if (extend) {
      const r = this.expandForMerges(rangeOf(this.sel.anchor, pos));
      this.sel = { ...this.sel, ranges: [...this.sel.ranges.slice(0, -1), r], active: this.sel.active };
    } else {
      const m = this.mergeAt(pos.row, pos.col);
      const r = m ?? { r1: pos.row, c1: pos.col, r2: pos.row, c2: pos.col };
      const start = m ? { row: m.r1, col: m.c1 } : pos;
      this.sel = { ranges: add ? [...this.sel.ranges, r] : [r], active: start, anchor: start };
    }
    this.emit();
  }

  selectRange(r: CellRange, active?: Pos): void {
    const a = active ?? { row: r.r1, col: r.c1 };
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
    row += dRow; col += dCol;
    // Skip hidden rows and columns.
    const s = this.sheet;
    while (dRow && row > 0 && row < MAX_ROWS - 1 && s.rows.get(row)?.hidden) row += Math.sign(dRow);
    while (dCol && col > 0 && col < MAX_COLS - 1 && s.cols.get(col)?.hidden) col += Math.sign(dCol);
    this.select({ row, col });
  }

  /** Ctrl+arrow: jump to the edge of the current block of data. */
  edge(pos: Pos, dRow: number, dCol: number): Pos {
    const has = (r: number, c: number) => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS && this.value(r, c) !== null;
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
    if (this.isLocked(row, col)) { this.notify('This cell is protected. Unprotect the sheet to change it.'); return; }
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
    if (text.startsWith('=')) {
      const fixed = autoCloseParens(text.slice(1));
      try { parseFormula(fixed); text = `=${fixed}`; } catch (err) {
        if (err instanceof FormulaError) { this.notify(`There's a problem with this formula. ${err.message}`); return false; }
        throw err;
      }
    }
    const err = this.validateEntry(e.row, e.col, text);
    if (err) { this.notify(err); return false; }
    this.editing = null;
    const r = this.range;
    if (arrayFormula && text.startsWith('=')) {
      this.change('Array formula', [{ area: r }], () => this.setArrayFormula(r, text.slice(1)));
      return true;
    }
    const targets = fill ? this.sel.ranges : [{ r1: e.row, c1: e.col, r2: e.row, c2: e.col }];
    this.change('Typing', targets.map((area) => ({ area })), () => {
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
    const n = parseNumber(text, this.wb.date1904);
    if (n !== null && !/^0\d/.test(text.trim())) {
      this.engine.setValue(sheet, row, col, n);
      const cur = this.style(row, col, sheet);
      const fmt = guessFormat(text);
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
  applyStyle(patch: CellStyle | ((s: CellStyle) => CellStyle), label = 'Format'): void {
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

  isLocked(row: number, col: number): boolean {
    if (!this.sheet.protection?.sheet) return false;
    return this.style(row, col).protection?.locked !== false;
  }

  private validateEntry(row: number, col: number, text: string): string | null {
    const dv = this.sheet.validations.find((v) => v.ranges.some((r) => inRange(r, row, col)));
    if (!dv || dv.type === 'any' || text.startsWith('=')) return null;
    if (text === '' && dv.allowBlank !== false) return null;
    const ok = checkValidation(this, dv, row, col, text);
    if (ok || dv.error?.style === 'information' || dv.error?.style === 'warning') return null;
    return dv.error?.text || 'This value doesn\'t match the data validation rules defined for this cell.';
  }
}

// ---------- Helpers ----------

import { checkValidation } from './validation';
import { shiftFormula } from '../engine/refactor';

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
const dateText = (v: number, fmt: string): string => {
  const hasDate = /[dy]|m{3,}/i.test(fmt.replace(/"[^"]*"/g, '')) || v >= 1;
  const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"/g, '')) || v % 1 !== 0;
  const parts: string[] = [];
  if (hasDate && v >= 1) { const d = fromSerial(v); parts.push(`${d.month}/${d.day}/${d.year}`); }
  if (hasTime) {
    const t = timeOf(v);
    const h12 = t.hour % 12 || 12;
    parts.push(`${h12}:${String(t.minute).padStart(2, '0')}:${String(t.second).padStart(2, '0')} ${t.hour < 12 ? 'AM' : 'PM'}`);
  }
  return parts.join(' ') || String(v);
};

export const cloneBook = (wb: Workbook): Workbook => ({ ...wb, sheets: wb.sheets.map(cloneSheet), names: structuredClone(wb.names), styles: structuredClone(wb.styles) });
