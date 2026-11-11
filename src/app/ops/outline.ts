// Outlines (grouped rows and columns with collapse and expand), Auto Outline, Subtotal and Consolidate.
import { parseFormula, walk } from '../../engine/parser';
import { MAX_COLS, MAX_ROWS, isError, type Scalar } from '../../engine/types';
import { shiftFormula } from '../../engine/refactor';
import { cellKey, colName, internStyle, keyCol, keyRow, type Cell, type CellRange, type CellStyle, type ColInfo, type RowInfo, type Sheet } from '../../model/workbook';
import { changeStructure } from './structure';
import { currentRegion } from './data';
import type { GridStore } from '../store';
import { N, T } from '../../ui/i18n';

export const MAX_OUTLINE = 7;
export type Axis = 'row' | 'col';

const mapOf = (s: Sheet, axis: Axis): Map<number, RowInfo | ColInfo> => (axis === 'row' ? s.rows : s.cols);
const levelOf = (s: Sheet, axis: Axis, i: number): number => mapOf(s, axis).get(i)?.outline ?? 0;
const summaryAfter = (s: Sheet, axis: Axis): boolean => (axis === 'row' ? s.outlinePr?.summaryBelow !== false : s.outlinePr?.summaryRight !== false);

export const maxLevel = (s: Sheet, axis: Axis): number => {
  let m = 0;
  for (const v of mapOf(s, axis).values()) if ((v.outline ?? 0) > m) m = v.outline!;
  return m;
};

export interface OutlineGroup { start: number; end: number; level: number; summary: number; collapsed: boolean }

/** Every group (run of lines at or above a level) on an axis, outer groups first. */
export const outlineGroups = (s: Sheet, axis: Axis): OutlineGroup[] => {
  const map = mapOf(s, axis);
  const max = maxLevel(s, axis);
  if (!max) return [];
  const idx = [...map.entries()].filter(([, v]) => (v.outline ?? 0) > 0).map(([k]) => k).sort((a, b) => a - b);
  const out: OutlineGroup[] = [];
  const after = summaryAfter(s, axis);
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  for (let lvl = 1; lvl <= max; lvl++) {
    let start = -1, prev = -2;
    const flush = () => {
      if (start < 0) return;
      const summary = after ? prev + 1 : start - 1;
      const sumInfo = summary >= 0 && summary < limit ? map.get(summary) : undefined;
      out.push({ start, end: prev, level: lvl, summary, collapsed: !!sumInfo?.collapsed || allHidden(map, start, prev) });
      start = -1;
    };
    for (const i of idx) {
      if ((map.get(i)!.outline ?? 0) < lvl) continue;
      if (i !== prev + 1) { flush(); start = i; }
      if (start < 0) start = i;
      prev = i;
    }
    flush();
  }
  return out;
};

const allHidden = (map: Map<number, RowInfo | ColInfo>, a: number, b: number): boolean => {
  for (let i = a; i <= b; i++) if (!map.get(i)?.hidden) return false;
  return true;
};

const setInfo = (s: Sheet, axis: Axis, i: number, patch: Partial<RowInfo & ColInfo>): void => {
  const map = mapOf(s, axis);
  const next = { ...(map.get(i) ?? {}), ...patch } as RowInfo & ColInfo;
  for (const k of Object.keys(next) as (keyof typeof next)[]) if (next[k] === undefined || next[k] === false || (k === 'outline' && !next[k])) delete next[k];
  if (Object.keys(next).length) map.set(i, next); else map.delete(i);
};

/** Group (delta 1) or ungroup (delta -1) lines. Ungrouping shows lines that were hidden by a collapsed group. */
export const groupLines = (store: GridStore, axis: Axis, from: number, to: number, delta: 1 | -1): void => {
  const s = store.sheet;
  if (delta > 0) {
    for (let i = from; i <= to; i++) if (levelOf(s, axis, i) >= MAX_OUTLINE) { store.notify(T('Outlines can have at most 8 levels.')); return; }
  }
  store.change(delta > 0 ? N('Group') : N('Ungroup'), 'sheet', () => {
    for (let i = from; i <= to; i++) {
      const lvl = Math.max(0, Math.min(MAX_OUTLINE, levelOf(s, axis, i) + delta));
      setInfo(s, axis, i, { outline: lvl || undefined, ...(delta < 0 && mapOf(s, axis).get(i)?.hidden && !isFiltered(s, axis, i) ? { hidden: undefined } : {}) });
    }
    // A summary line left with no group loses its collapsed mark.
    const after = summaryAfter(s, axis);
    for (const i of [from - 1, to + 1]) if (i >= 0 && !groupEndingAt(s, axis, i, after)) setInfo(s, axis, i, { collapsed: undefined });
  });
};

const isFiltered = (s: Sheet, axis: Axis, i: number) => axis === 'row' && !!s.rows.get(i)?.filtered;
const groupEndingAt = (s: Sheet, axis: Axis, summary: number, after: boolean): boolean => levelOf(s, axis, after ? summary - 1 : summary + 1) > levelOf(s, axis, summary);

/** Lines a Group command should act on for the selection. */
export const groupTarget = (store: GridStore): { axis: Axis; from: number; to: number } => {
  const r = store.range;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return { axis: 'row', from: r.r1, to: r.r2 };
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return { axis: 'col', from: r.c1, to: r.c2 };
  if (r.c1 === r.c2 && r.r1 !== r.r2) return { axis: 'row', from: r.r1, to: r.r2 };
  if (r.r1 === r.r2 && r.c1 !== r.c2) return { axis: 'col', from: r.c1, to: r.c2 };
  return { axis: 'row', from: r.r1, to: r.r2 };
};

/** Collapses or expands one group. Expanding keeps nested collapsed groups closed. */
export const setGroupCollapsed = (store: GridStore, axis: Axis, g: OutlineGroup, collapsed: boolean): void => {
  store.change(collapsed ? N('Hide detail') : N('Show detail'), 'sheet', () => applyCollapsed(store.sheet, axis, g, collapsed));
};

const applyCollapsed = (s: Sheet, axis: Axis, g: OutlineGroup, collapsed: boolean): void => {
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (g.summary >= 0 && g.summary < limit) setInfo(s, axis, g.summary, { collapsed: collapsed || undefined });
  if (collapsed) { for (let i = g.start; i <= g.end; i++) setInfo(s, axis, i, { hidden: true }); return; }
  const inner = outlineGroups(s, axis).filter((x) => x.level > g.level && x.start >= g.start && x.end <= g.end);
  for (let i = g.start; i <= g.end; i++) {
    const hiddenByInner = inner.some((x) => i >= x.start && i <= x.end && !!mapOf(s, axis).get(x.summary)?.collapsed);
    if (!hiddenByInner && !isFiltered(s, axis, i)) setInfo(s, axis, i, { hidden: undefined });
  }
};

/** Level buttons: shows levels up to n (n = max+1 shows everything). */
export const showLevel = (store: GridStore, axis: Axis, n: number): void => {
  store.change(N('Show level'), 'sheet', () => {
    const s = store.sheet;
    for (const g of outlineGroups(s, axis)) {
      const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
      if (g.summary >= 0 && g.summary < limit) setInfo(s, axis, g.summary, { collapsed: g.level >= n || undefined });
    }
    for (const [i, info] of [...mapOf(s, axis)]) {
      const lvl = info.outline ?? 0;
      if (!lvl) continue;
      if (lvl >= n) setInfo(s, axis, i, { hidden: true });
      else if (!isFiltered(s, axis, i)) setInfo(s, axis, i, { hidden: undefined });
    }
  });
};

/** The group whose summary line or detail holds the active cell, innermost first. */
const groupAtCursor = (store: GridStore, axis: Axis, collapsedWanted: boolean): OutlineGroup | null => {
  const pos = axis === 'row' ? store.sel.active.row : store.sel.active.col;
  const gs = outlineGroups(store.sheet, axis).sort((a, b) => b.level - a.level);
  return gs.find((g) => g.summary === pos && g.collapsed === collapsedWanted) ?? gs.find((g) => pos >= g.start && pos <= g.end && g.collapsed === collapsedWanted) ?? null;
};

export const showDetail = (store: GridStore, show: boolean): boolean => {
  const t = groupTarget(store);
  for (const axis of [t.axis, t.axis === 'row' ? 'col' : 'row'] as Axis[]) {
    const g = groupAtCursor(store, axis, show);
    if (g) { setGroupCollapsed(store, axis, g, !show); return true; }
  }
  store.notify(show ? T('There is no hidden detail here.') : T('There is no detail to hide here.'));
  return false;
};

/** Removes outline levels in the selection (or the whole sheet for one cell) and shows the lines again. */
export const clearOutline = (store: GridStore): void => {
  const r = store.range;
  const whole = r.r1 === r.r2 && r.c1 === r.c2;
  store.change(N('Clear outline'), 'sheet', () => {
    const s = store.sheet;
    for (const axis of ['row', 'col'] as Axis[]) {
      const lo = whole ? 0 : axis === 'row' ? r.r1 : r.c1, hi = whole ? Infinity : axis === 'row' ? r.r2 : r.c2;
      for (const [i, info] of [...mapOf(s, axis)]) {
        if (i < lo - 1 || i > hi + 1) continue;
        if (info.outline && i >= lo && i <= hi) setInfo(s, axis, i, { outline: undefined, hidden: info.hidden && !isFiltered(s, axis, i) ? undefined : info.hidden });
        if (info.collapsed) setInfo(s, axis, i, { collapsed: undefined });
      }
    }
  });
};

// ---------------------------------------------------------------- Auto Outline

interface Span { start: number; end: number }

/**
 * Builds an outline from summary formulas: a formula that refers to a block directly above it
 * (or to its left) makes that block a group. Groups inside groups get deeper levels.
 */
export const autoOutline = (store: GridStore): number => {
  const s = store.sheet;
  const r = store.range.r1 === store.range.r2 && store.range.c1 === store.range.c2 ? currentRegion(store) : store.range;
  const rowSpans: Span[] = [], colSpans: Span[] = [];
  const below = s.outlinePr?.summaryBelow !== false, right = s.outlinePr?.summaryRight !== false;
  for (const [k, c] of s.cells) {
    if (!c.f) continue;
    const row = keyRow(k), col = keyCol(k);
    if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
    let node;
    try { node = parseFormula(c.f); } catch { continue; }
    const above: { r1: number; r2: number }[] = [], left: { c1: number; c2: number }[] = [];
    walk(node, (n) => {
      if (n.type !== 'ref' || n.ref.sheet || n.ref.kind === 'cols' || n.ref.kind === 'rows') return;
      const a = n.ref;
      if (a.c1 === col && a.c2 === col && (below ? a.r2 < row : a.r1 > row)) above.push({ r1: a.r1, r2: a.r2 });
      if (a.r1 === row && a.r2 === row && (right ? a.c2 < col : a.c1 > col)) left.push({ c1: a.c1, c2: a.c2 });
    });
    // A block of cells, or several cells, next to the formula: the block is its detail.
    if (above.length > 1 || (above.length === 1 && above[0]!.r1 !== above[0]!.r2)) {
      rowSpans.push(below ? { start: Math.min(...above.map((x) => x.r1)), end: row - 1 } : { start: row + 1, end: Math.max(...above.map((x) => x.r2)) });
    }
    if (left.length > 1 || (left.length === 1 && left[0]!.c1 !== left[0]!.c2)) {
      colSpans.push(right ? { start: Math.min(...left.map((x) => x.c1)), end: col - 1 } : { start: col + 1, end: Math.max(...left.map((x) => x.c2)) });
    }
  }
  // A span that starts at another span's summary line takes in that span's detail too.
  const widen = (spans: Span[], after: boolean): Span[] => {
    const out = spans.map((x) => ({ ...x }));
    for (let changed = true, guard = 0; changed && guard < 50; guard++) {
      changed = false;
      for (const sp of out) for (const o of out) {
        if (o === sp) continue;
        if (after && o.end === sp.start - 1 && o.start < sp.start && o.end < sp.end) { sp.start = o.start; changed = true; }
        if (!after && o.start === sp.end + 1 && o.end > sp.end && o.start > sp.start) { sp.end = o.end; changed = true; }
      }
    }
    return out;
  };
  rowSpans.splice(0, rowSpans.length, ...widen(rowSpans, below));
  colSpans.splice(0, colSpans.length, ...widen(colSpans, right));
  const levels = (spans: Span[]): Map<number, number> => {
    const uniq = [...new Map(spans.map((x) => [`${x.start}:${x.end}`, x])).values()];
    const m = new Map<number, number>();
    for (const sp of uniq) for (let i = sp.start; i <= sp.end; i++) m.set(i, Math.min(MAX_OUTLINE, (m.get(i) ?? 0) + 1));
    return m;
  };
  const rl = levels(rowSpans), cl = levels(colSpans);
  if (!rl.size && !cl.size) { store.notify(T('Couldn\'t find summary formulas to build an outline from.')); return 0; }
  store.change(N('Auto outline'), 'sheet', () => {
    for (const [i, info] of [...s.rows]) if (info.outline && i >= r.r1 && i <= r.r2) setInfo(s, 'row', i, { outline: undefined, collapsed: undefined });
    for (const [i, info] of [...s.cols]) if (info.outline && i >= r.c1 && i <= r.c2) setInfo(s, 'col', i, { outline: undefined, collapsed: undefined });
    for (const [i, l] of rl) setInfo(s, 'row', i, { outline: l });
    for (const [i, l] of cl) setInfo(s, 'col', i, { outline: l });
  });
  return rl.size + cl.size;
};

// ---------------------------------------------------------------- Subtotal

export const SUBTOTAL_FUNCS: [number, string][] = [
  [9, N('Sum')], [3, N('Count')], [1, N('Average')], [4, N('Max')], [5, N('Min')], [6, N('Product')], [2, N('Count Numbers')], [7, 'StdDev'], [8, 'StdDevp'], [10, N('Var')], [11, N('Varp')],
];
const FN_NAME = new Map(SUBTOTAL_FUNCS.map(([n, l]) => [n, l]));

export interface SubtotalOptions {
  /** Column (sheet index) whose changes start a new group. */
  groupBy: number;
  fn: number;
  columns: number[];
  replace: boolean;
  pageBreaks: boolean;
  summaryBelow: boolean;
}

const isSubtotalRow = (s: Sheet, row: number, c1: number, c2: number): boolean => {
  for (let c = c1; c <= c2; c++) if (/\bSUBTOTAL\s*\(/i.test(s.cells.get(cellKey(row, c))?.f ?? '')) return true;
  return false;
};

/** The list a Subtotal applies to: the selection, or the region around the active cell, with its header row. */
export const subtotalRange = (store: GridStore): CellRange => {
  const r = store.range;
  return r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(store) : r;
};

const textOf = (v: Scalar): string => (v === null ? '' : isError(v) ? v.code : String(v));

interface Row { src: number | null; kind: 'data' | 'sub' | 'grand'; level: number; label?: string; fn?: number; cols?: number[]; group?: [number, number] }

/**
 * Inserts a subtotal row after (or before) each group and a grand total, and outlines the list.
 * With `replace` off, new subtotals nest inside the ones already there.
 */
export const applySubtotals = (store: GridStore, range: CellRange, o: SubtotalOptions): void => {
  const s = store.sheet;
  const header = range.r1;
  store.change(N('Subtotal'), 'book', () => {
    let r = range;
    if (o.replace) r = removeSubtotalsIn(store, r);
    s.outlinePr = { ...s.outlinePr, summaryBelow: o.summaryBelow };
    // Rows of the list as they are now.
    const rows: Row[] = [];
    let layers = 0;
    for (let row = header + 1; row <= r.r2; row++) {
      if (isSubtotalRow(s, row, r.c1, r.c2)) {
        const isGrand = /grand/i.test(textOf(store.value(row, o.groupBy))) || (levelOf(s, 'row', row) === 0 && row === r.r2);
        const lvl = levelOf(s, 'row', row);
        if (!isGrand) layers = Math.max(layers, lvl);
        rows.push({ src: row, kind: isGrand ? 'grand' : 'sub', level: lvl });
      } else rows.push({ src: row, kind: 'data', level: 0 });
    }
    const newLayer = layers + 1;
    // Insert new subtotal rows inside each run of data rows.
    const next: Row[] = [];
    let run: Row[] = [];
    const key = (x: Row) => textOf(store.value(x.src!, o.groupBy)).toLowerCase();
    const flush = () => {
      let i = 0;
      while (i < run.length) {
        let j = i;
        while (j + 1 < run.length && key(run[j + 1]!) === key(run[i]!)) j++;
        const label = textOf(store.value(run[i]!.src!, o.groupBy));
        const sub: Row = { src: null, kind: 'sub', level: newLayer, label: `${label} ${FN_NAME.get(o.fn) === 'Sum' ? 'Total' : FN_NAME.get(o.fn)}`, fn: o.fn, cols: o.columns };
        if (!o.summaryBelow) next.push(sub);
        next.push(...run.slice(i, j + 1));
        if (o.summaryBelow) next.push(sub);
        i = j + 1;
      }
      run = [];
    };
    for (const x of rows) { if (x.kind === 'data') run.push(x); else { flush(); next.push(x); } }
    flush();
    if (!next.some((x) => x.kind === 'grand')) {
      const g: Row = { src: null, kind: 'grand', level: 0, label: `Grand ${FN_NAME.get(o.fn) === 'Sum' ? 'Total' : FN_NAME.get(o.fn)}`, fn: o.fn, cols: o.columns };
      if (o.summaryBelow) next.push(g); else next.unshift(g);
    }
    const added = next.length - rows.length;
    // Make room below the list, then rewrite the list rows in their new order.
    if (added > 0) changeStructure(store, 'row', r.r2 + 1, added);
    writeRows(store, header, r, rows.length, next, o, newLayer);
  });
};

/** Writes the rearranged list and its outline levels and subtotal formulas. */
const writeRows = (store: GridStore, header: number, r: CellRange, oldCount: number, next: Row[], o: SubtotalOptions, deepest: number): void => {
  const s = store.sheet;
  const saved = new Map<number, { cells: Map<number, Cell>; info: RowInfo | undefined }>();
  for (let row = header + 1; row <= header + oldCount; row++) {
    const cells = new Map<number, Cell>();
    for (let c = r.c1; c <= r.c2; c++) { const cell = s.cells.get(cellKey(row, c)); if (cell) cells.set(c, structuredClone(cell)); }
    saved.set(row, { cells, info: s.rows.get(row) ? { ...s.rows.get(row)! } : undefined });
  }
  for (let row = header + 1; row <= header + next.length; row++) {
    for (let c = r.c1; c <= r.c2; c++) store.engine.setValue(store.sheetIndex, row, c, null);
    s.rows.delete(row);
  }
  if (o.pageBreaks) s.pageSetup = { ...s.pageSetup, rowBreaks: (s.pageSetup?.rowBreaks ?? []).filter((b) => b <= header || b > header + next.length) };
  const si = store.sheetIndex;
  // Final positions, to build each subtotal's range.
  const pos = next.map((_, i) => header + 1 + i);
  const levelFor = (x: Row): number => (x.kind === 'grand' ? 0 : x.kind === 'sub' ? x.level : deepest + 1);
  next.forEach((x, i) => {
    const row = pos[i]!;
    if (x.src !== null) {
      const sv = saved.get(x.src)!;
      for (const [c, cell] of sv.cells) {
        const copy = structuredClone(cell);
        if (copy.f) {
          if (x.kind !== 'data') continue;
          store.engine.setFormula(si, row, c, shiftFormula(copy.f, row - x.src, 0));
          const placed = s.cells.get(cellKey(row, c));
          if (placed && copy.s) placed.s = copy.s;
        } else { s.cells.set(cellKey(row, c), copy); store.engine.markChanged(si, row, c); }
      }
      if (x.kind !== 'data') {
        // Existing subtotal rows: keep their labels and rebuild their formulas below.
        for (const [c, cell] of sv.cells) if (!cell.f) { s.cells.set(cellKey(row, c), structuredClone(cell)); }
        x.cols = [...sv.cells].filter(([, cell]) => /\bSUBTOTAL\s*\(/i.test(cell.f ?? '')).map(([c]) => c);
        const f0 = [...sv.cells.values()].find((cell) => /\bSUBTOTAL\s*\(/i.test(cell.f ?? ''))?.f ?? '';
        x.fn = Number(/SUBTOTAL\s*\(\s*(\d+)/i.exec(f0)?.[1] ?? o.fn);
      }
      const info = sv.info ? { ...sv.info } : {};
      delete info.outline; delete info.collapsed;
      if (!(info.hidden && info.filtered)) delete info.hidden;
      const lvl = levelFor(x);
      s.rows.set(row, { ...info, ...(lvl ? { outline: lvl } : {}) });
    } else {
      const lvl = levelFor(x);
      if (lvl) s.rows.set(row, { outline: lvl });
      store.engine.setValue(si, row, o.groupBy, x.label ?? null);
      const bold = { font: { bold: true } };
      const id = internBold(store, bold);
      const lc = s.cells.get(cellKey(row, o.groupBy));
      if (lc) lc.s = id;
    }
  });
  // Subtotal formulas: each covers the data and deeper subtotal rows of its group.
  next.forEach((x, i) => {
    if (x.kind === 'data' || !x.cols) return;
    const row = pos[i]!;
    let a: number, b: number;
    if (x.kind === 'grand') { a = pos[0]!; b = pos[pos.length - 1]!; if (o.summaryBelow) b = row - 1; else a = row + 1; }
    else {
      a = i; b = i;
      const lvl = x.level;
      if (o.summaryBelow) { while (a - 1 >= 0 && next[a - 1]!.kind !== 'grand' && !(next[a - 1]!.kind === 'sub' && next[a - 1]!.level <= lvl)) a--; b = i - 1; a = pos[a]!; b = pos[b]!; }
      else { while (b + 1 < next.length && next[b + 1]!.kind !== 'grand' && !(next[b + 1]!.kind === 'sub' && next[b + 1]!.level <= lvl)) b++; a = pos[i + 1]!; b = pos[b]!; }
    }
    for (const c of x.cols) {
      if (a > b) continue;
      store.engine.setFormula(si, row, c, `SUBTOTAL(${x.fn ?? o.fn},${colName(c)}${a + 1}:${colName(c)}${b + 1})`);
      const placed = s.cells.get(cellKey(row, c));
      if (placed) placed.s = internBold(store, { font: { bold: true }, ...(store.style(a, c).numFmt ? { numFmt: store.style(a, c).numFmt } : {}) });
    }
    if (o.pageBreaks && x.kind === 'sub' && x.level === deepest && o.summaryBelow) s.pageSetup = { ...s.pageSetup, rowBreaks: [...(s.pageSetup?.rowBreaks ?? []), row + 1].sort((p, q) => p - q) };
  });
};

const internBold = (store: GridStore, style: CellStyle): number => internStyle(store.wb, style);

/** Deletes subtotal rows from a list and its outline. Returns the list's new range. Runs inside a change. */
const removeSubtotalsIn = (store: GridStore, r: CellRange): CellRange => {
  const s = store.sheet;
  const keep: number[] = [];
  let removed = 0;
  for (let row = r.r1 + 1; row <= r.r2; row++) { if (isSubtotalRow(s, row, r.c1, r.c2)) removed++; else keep.push(row); }
  for (let row = r.r1 + 1; row <= r.r2; row++) { const info = s.rows.get(row); if (info) { const n = { ...info }; delete n.outline; delete n.collapsed; if (!n.filtered) delete n.hidden; if (Object.keys(n).length) s.rows.set(row, n); else s.rows.delete(row); } }
  if (s.pageSetup?.rowBreaks) s.pageSetup.rowBreaks = s.pageSetup.rowBreaks.filter((b) => b <= r.r1 || b > r.r2 + 1);
  if (!removed) return r;
  // Pull kept rows up, then delete the spare rows at the end of the list.
  const si = store.sheetIndex;
  keep.forEach((src, i) => {
    const dst = r.r1 + 1 + i;
    if (src === dst) return;
    for (let c = r.c1; c <= r.c2; c++) {
      const cell = s.cells.get(cellKey(src, c));
      store.engine.setValue(si, dst, c, null);
      if (!cell) continue;
      const copy = structuredClone(cell);
      if (copy.f) { store.engine.setFormula(si, dst, c, shiftFormula(copy.f, dst - src, 0)); const p = s.cells.get(cellKey(dst, c)); if (p && copy.s) p.s = copy.s; }
      else { s.cells.set(cellKey(dst, c), copy); store.engine.markChanged(si, dst, c); }
      store.engine.setValue(si, src, c, null);
      s.cells.delete(cellKey(src, c));
    }
    const info = s.rows.get(src);
    if (info) s.rows.set(dst, info); else s.rows.delete(dst);
  });
  changeStructure(store, 'row', r.r2 - removed + 1, -removed);
  return { ...r, r2: r.r2 - removed };
};

export const removeSubtotals = (store: GridStore, range: CellRange): void => {
  store.change(N('Remove subtotals'), 'book', () => { removeSubtotalsIn(store, range); });
};

export const hasSubtotals = (store: GridStore, r: CellRange): boolean => {
  for (let row = r.r1 + 1; row <= r.r2; row++) if (isSubtotalRow(store.sheet, row, r.c1, r.c2)) return true;
  return false;
};

// ---------------------------------------------------------------- select visible cells

/** Splits the selection into blocks of visible cells (Alt+;), so copy leaves hidden rows and columns out. */
export const selectVisible = (store: GridStore): void => {
  const s = store.sheet;
  const b = store.engine.usedBounds(store.sheetIndex);
  const out: CellRange[] = [];
  for (const r0 of store.sel.ranges) {
    const r = { ...r0, r2: r0.r2 === MAX_ROWS - 1 ? Math.max(r0.r1, b.rows) : r0.r2, c2: r0.c2 === MAX_COLS - 1 ? Math.max(r0.c1, b.cols) : r0.c2 };
    const runs = (lo: number, hi: number, hidden: (i: number) => boolean): [number, number][] => {
      const res: [number, number][] = [];
      let st = -1;
      for (let i = lo; i <= hi + 1; i++) {
        const vis = i <= hi && !hidden(i);
        if (vis && st < 0) st = i;
        if (!vis && st >= 0) { res.push([st, i - 1]); st = -1; }
      }
      return res;
    };
    const rr = runs(r.r1, r.r2, (i) => !!s.rows.get(i)?.hidden);
    const cc = runs(r.c1, r.c2, (i) => !!s.cols.get(i)?.hidden);
    for (const [r1, r2] of rr) for (const [c1, c2] of cc) out.push({ r1, r2, c1, c2 });
  }
  if (!out.length) { store.notify(T('No visible cells in the selection.')); return; }
  const a = store.sel.active;
  const active = out.some((x) => a.row >= x.r1 && a.row <= x.r2 && a.col >= x.c1 && a.col <= x.c2) ? a : { row: out[0]!.r1, col: out[0]!.c1 };
  const last = out.findIndex((x) => active.row >= x.r1 && active.row <= x.r2 && active.col >= x.c1 && active.col <= x.c2);
  const ranges = [...out.slice(0, last), ...out.slice(last + 1), out[last]!];
  store.sel = { ranges, active, anchor: active };
  store.emit();
};

// ---------------------------------------------------------------- Consolidate

export type ConsolidateFn = 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';
export const CONSOLIDATE_FUNCS: [ConsolidateFn, string, string][] = [
  ['sum', N('Sum'), 'SUM'], ['count', N('Count'), 'COUNTA'], ['average', N('Average'), 'AVERAGE'], ['max', N('Max'), 'MAX'], ['min', N('Min'), 'MIN'], ['product', N('Product'), 'PRODUCT'],
  ['countNums', N('Count Numbers'), 'COUNT'], ['stdDev', 'StdDev', 'STDEV.S'], ['stdDevp', 'StdDevp', 'STDEV.P'], ['var', N('Var'), 'VAR.S'], ['varp', N('Varp'), 'VAR.P'],
];

export interface ConsolidateOptions { fn: ConsolidateFn; refs: { sheet: number; range: CellRange }[]; topRow: boolean; leftCol: boolean; links: boolean }

const aggregate = (fn: ConsolidateFn, vals: Scalar[]): Scalar => {
  const nonBlank = vals.filter((v) => v !== null && v !== '');
  const nums = nonBlank.filter((v): v is number => typeof v === 'number');
  const n = nums.length, sum = nums.reduce((a, b) => a + b, 0);
  const mean = n ? sum / n : 0;
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  switch (fn) {
    case 'sum': return sum;
    case 'count': return nonBlank.length;
    case 'countNums': return n;
    case 'average': return n ? mean : null;
    case 'max': return n ? Math.max(...nums) : 0;
    case 'min': return n ? Math.min(...nums) : 0;
    case 'product': return n ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'stdDev': return n > 1 ? Math.sqrt(ss / (n - 1)) : null;
    case 'stdDevp': return n ? Math.sqrt(ss / n) : null;
    case 'var': return n > 1 ? ss / (n - 1) : null;
    case 'varp': return n ? ss / n : null;
  }
};

/**
 * Combines several ranges into one table at the active cell: by position, or matched by the labels
 * in their top row and/or left column. With links, each result is a formula over linked detail rows
 * that are grouped and collapsed under it.
 */
export const consolidate = (store: GridStore, o: ConsolidateOptions): string | null => {
  if (!o.refs.length) return T('Add at least one reference.');
  const dest = store.sel.active;
  const destSheet = store.sheetIndex;
  if (o.refs.some((r) => r.sheet === destSheet && dest.row >= r.range.r1 && dest.row <= r.range.r2 && dest.col >= r.range.c1 && dest.col <= r.range.c2)) return T('The destination overlaps a source range.');
  const val = (sheet: number, row: number, col: number) => store.value(row, col, sheet);
  const sheetName = (i: number) => store.wb.sheets[i]!.name;
  const q = (i: number) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName(i)) ? sheetName(i) : `'${sheetName(i).replace(/'/g, "''")}'`);
  type Src = { ref: { sheet: number; range: CellRange }; rows: Map<string, number>; cols: Map<string, number> };
  const rowLabels: string[] = [], colLabels: string[] = [];
  const addLabel = (list: string[], l: string) => { if (!list.some((x) => x.toLowerCase() === l.toLowerCase())) list.push(l); };
  const srcs: Src[] = o.refs.map((ref) => {
    const r = ref.range;
    const rows = new Map<string, number>(), cols = new Map<string, number>();
    const r0 = r.r1 + (o.topRow ? 1 : 0), c0 = r.c1 + (o.leftCol ? 1 : 0);
    for (let row = r0; row <= r.r2; row++) {
      const l = o.leftCol ? textOf(val(ref.sheet, row, r.c1)) : String(row - r0);
      if (o.leftCol && !l) continue;
      if (!rows.has(l.toLowerCase())) rows.set(l.toLowerCase(), row);
      addLabel(rowLabels, l);
    }
    for (let col = c0; col <= r.c2; col++) {
      const l = o.topRow ? textOf(val(ref.sheet, r.r1, col)) : String(col - c0);
      if (o.topRow && !l) continue;
      if (!cols.has(l.toLowerCase())) cols.set(l.toLowerCase(), col);
      addLabel(colLabels, l);
    }
    return { ref, rows, cols };
  });
  const fnName = CONSOLIDATE_FUNCS.find((f) => f[0] === o.fn)![2];
  const offR = o.topRow ? 1 : 0, offC = o.leftCol ? 1 : 0;
  store.change(N('Consolidate'), 'book', () => {
    const s = store.sheet;
    if (o.topRow) colLabels.forEach((l, j) => store.engine.setValue(destSheet, dest.row, dest.col + offC + j, l));
    let outRow = dest.row + offR;
    for (const rl of rowLabels) {
      const hits = colLabels.map((cl) => srcs.flatMap((sr) => {
        const row = sr.rows.get(rl.toLowerCase()), col = sr.cols.get(cl.toLowerCase());
        return row !== undefined && col !== undefined ? [{ sr, row, col }] : [];
      }));
      if (o.links) {
        // One linked detail row per source, then the summary row.
        const detailSources = srcs.filter((sr) => sr.rows.has(rl.toLowerCase()));
        const first = outRow;
        detailSources.forEach((sr, k) => {
          const row = first + k;
          if (o.leftCol) store.engine.setValue(destSheet, row, dest.col, sheetName(sr.ref.sheet));
          colLabels.forEach((cl, j) => {
            const srow = sr.rows.get(rl.toLowerCase())!, scol = sr.cols.get(cl.toLowerCase());
            if (scol === undefined) return;
            store.engine.setFormula(destSheet, row, dest.col + offC + j, `${q(sr.ref.sheet)}!$${colName(scol)}$${srow + 1}`);
          });
          s.rows.set(row, { ...(s.rows.get(row) ?? {}), outline: 1, hidden: true });
        });
        const sumRow = first + detailSources.length;
        if (o.leftCol) store.engine.setValue(destSheet, sumRow, dest.col, rl);
        colLabels.forEach((_, j) => {
          if (!detailSources.length) return;
          const c = dest.col + offC + j;
          store.engine.setFormula(destSheet, sumRow, c, `${fnName}(${colName(c)}${first + 1}:${colName(c)}${sumRow})`);
        });
        s.rows.set(sumRow, { ...(s.rows.get(sumRow) ?? {}), collapsed: detailSources.length > 0 || undefined });
        outRow = sumRow + 1;
      } else {
        if (o.leftCol) store.engine.setValue(destSheet, outRow, dest.col, rl);
        hits.forEach((list, j) => { if (list.length) store.engine.setValue(destSheet, outRow, dest.col + offC + j, aggregate(o.fn, list.map((h) => val(h.sr.ref.sheet, h.row, h.col)))); });
        outRow++;
      }
    }
    if (o.links) s.outlinePr = { ...s.outlinePr, summaryBelow: true };
  });
  return null;
};

export { colName };

/** Keeps PivotTable sources, slicer connections and PivotChart links on a renamed sheet. */
export const renamePivotSheet = (wb: import('../../model/workbook').Workbook, old: string, name: string): void => {
  const same = (a: string) => a.toLowerCase() === old.toLowerCase();
  if (wb.pivotCaches) wb.pivotCaches = wb.pivotCaches.map((c) => (c.source.kind === 'range' && same(c.source.sheet) ? { ...c, source: { ...c.source, sheet: name } } : c));
  for (const sc of wb.slicerCaches ?? []) if (sc.source.kind === 'pivot') sc.source.pivots = sc.source.pivots.map((p) => (same(p.sheet) ? { ...p, sheet: name } : p));
  for (const s of wb.sheets) for (const ch of s.charts) {
    if (!ch.pivot) continue;
    const i = ch.pivot.lastIndexOf('!');
    if (same(ch.pivot.slice(0, i))) ch.pivot = `${name}${ch.pivot.slice(i)}`;
  }
};

/**
 * Inserting or deleting rows or columns moves PivotTables, slicers and the ranges caches read from.
 * Runs inside changeStructure, before cell anchors move.
 */
export const shiftPivots = (wb: import('../../model/workbook').Workbook, sheet: number, axis: Axis, at: number, count: number): void => {
  const s = wb.sheets[sheet]!;
  const mv = (i: number): number => (i >= at ? Math.max(at, i + count) : i);
  const isRow = axis === 'row';
  if (s.pivots) s.pivots = s.pivots.map((pt) => {
    const anchor = isRow ? { ...pt.anchor, row: mv(pt.anchor.row) } : { ...pt.anchor, col: mv(pt.anchor.col) };
    const loc = pt.location;
    const location = loc ? { ...loc, ref: isRow ? { ...loc.ref, r1: mv(loc.ref.r1), r2: Math.max(mv(loc.ref.r1), mv(loc.ref.r2)) } : { ...loc.ref, c1: mv(loc.ref.c1), c2: Math.max(mv(loc.ref.c1), mv(loc.ref.c2)) } } : loc;
    return { ...pt, anchor, location };
  });
  for (const sl of s.slicers ?? []) { if (isRow) sl.anchor.row = mv(sl.anchor.row); else sl.anchor.col = mv(sl.anchor.col); }
  if (wb.pivotCaches) wb.pivotCaches = wb.pivotCaches.map((c) => {
    if (c.source.kind !== 'range' || c.source.sheet.toLowerCase() !== s.name.toLowerCase()) return c;
    const r = c.source.range;
    const del = count < 0 ? -count : 0;
    const end = (i: number) => (count > 0 ? (i >= at ? i + count : i) : i >= at + del ? i - del : i >= at ? at - 1 : i);
    const range = isRow ? { ...r, r1: mv(r.r1), r2: Math.max(mv(r.r1), end(r.r2)) } : { ...r, c1: mv(r.c1), c2: Math.max(mv(r.c1), end(r.c2)) };
    return { ...c, source: { ...c.source, range } };
  });
};
