import { compareValues, parseNumber, toText } from '../../engine/coerce';
import { shiftFormula } from '../../engine/refactor';
import { CellError, isError, MAX_COLS, MAX_ROWS, type Scalar } from '../../engine/types';
import { cellKey, type Cell, type CellRange, type FilterColumn } from '../../model/workbook';
import { displayText } from './clipboard';
import { compareOp } from '../validation';
import type { GridStore } from '../store';

/** The block of data around the active cell, like other spreadsheets's Ctrl+A / current region. */
export const currentRegion = (store: GridStore, row = store.sel.active.row, col = store.sel.active.col): CellRange => {
  const has = (r: number, c: number) => r >= 0 && c >= 0 && store.value(r, c) !== null;
  let r = { r1: row, c1: col, r2: row, c2: col };
  for (let grew = true; grew;) {
    grew = false;
    const tryEdge = (cond: () => boolean, f: () => void) => { if (cond()) { f(); grew = true; } };
    const anyIn = (r1: number, c1: number, r2: number, c2: number) => { for (let i = r1; i <= r2; i++) for (let j = c1; j <= c2; j++) if (has(i, j)) return true; return false; };
    tryEdge(() => r.r1 > 0 && anyIn(r.r1 - 1, Math.max(0, r.c1 - 1), r.r1 - 1, r.c2 + 1), () => { r = { ...r, r1: r.r1 - 1 }; });
    tryEdge(() => r.r2 < MAX_ROWS - 1 && anyIn(r.r2 + 1, Math.max(0, r.c1 - 1), r.r2 + 1, r.c2 + 1), () => { r = { ...r, r2: r.r2 + 1 }; });
    tryEdge(() => r.c1 > 0 && anyIn(Math.max(0, r.r1 - 1), r.c1 - 1, r.r2 + 1, r.c1 - 1), () => { r = { ...r, c1: r.c1 - 1 }; });
    tryEdge(() => r.c2 < MAX_COLS - 1 && anyIn(Math.max(0, r.r1 - 1), r.c2 + 1, r.r2 + 1, r.c2 + 1), () => { r = { ...r, c2: r.c2 + 1 }; });
  }
  return r;
};

/** Guesses whether the first row of a block is a header row (text over numbers). */
export const looksLikeHeader = (store: GridStore, r: CellRange): boolean => {
  if (r.r1 === r.r2) return false;
  let textTop = 0, numBelow = 0;
  for (let c = r.c1; c <= r.c2; c++) {
    if (typeof store.value(r.r1, c) === 'string') textTop++;
    const below = store.value(r.r1 + 1, c);
    if (typeof below !== 'string') numBelow++;
  }
  const bold = store.style(r.r1, r.c1).font?.bold && !store.style(r.r1 + 1, r.c1).font?.bold;
  return textTop === r.c2 - r.c1 + 1 && (numBelow > 0 || !!bold);
};

/** The range a sort or filter applies to: the selection, or the current region if only one cell is selected. */
export const dataRange = (store: GridStore): CellRange => {
  const r = store.range;
  if (r.r1 === r.r2 && r.c1 === r.c2) return currentRegion(store);
  const b = store.engine.usedBounds(store.sheetIndex);
  return { ...r, r2: Math.min(r.r2, Math.max(r.r1, b.rows)), c2: Math.min(r.c2, Math.max(r.c1, b.cols)) };
};

export interface SortKey { col: number; descending?: boolean; by?: 'value' | 'cellColor' | 'fontColor'; color?: string; customList?: string[] }

/** Sorts the rows (or columns with byCols) of a range, moving formats with values and keeping formulas pointing at the same relative rows. */
export const sortRange = (store: GridStore, r: CellRange, keys: SortKey[], header: boolean, byCols = false, caseSensitive = false): void => {
  const sheet = store.sheetIndex;
  const start = byCols ? r.c1 + (header ? 1 : 0) : r.r1 + (header ? 1 : 0);
  const end = byCols ? r.c2 : r.r2;
  if (end <= start - 1) return;
  if (store.sheet.merges.some((m) => m.r1 <= r.r2 && m.r2 >= r.r1 && m.c1 <= r.c2 && m.c2 >= r.c1)) { store.notify('To sort, all merged cells in the range need to be the same size. Unmerge them first.'); return; }
  const lines: { idx: number; cells: (Cell | undefined)[]; hidden?: boolean }[] = [];
  const span = byCols ? [r.r1, r.r2] : [r.c1, r.c2];
  for (let i = start; i <= end; i++) {
    const cells: (Cell | undefined)[] = [];
    for (let j = span[0]!; j <= span[1]!; j++) { const c = byCols ? store.cell(j, i) : store.cell(i, j); cells.push(c ? structuredClone(c) : undefined); }
    lines.push({ idx: i, cells });
  }
  const keyVal = (ln: (typeof lines)[number], k: SortKey): Scalar => {
    const pos = k.col - span[0]!;
    const c = ln.cells[pos];
    const row = byCols ? span[0]! + pos : ln.idx, col = byCols ? ln.idx : span[0]! + pos;
    if (k.by === 'cellColor') return store.style(row, col).fill?.color === k.color ? 0 : 1;
    if (k.by === 'fontColor') return (store.style(row, col).font?.color ?? '#000000') === k.color ? 0 : 1;
    const v = c?.f ? store.value(row, col) : c?.v ?? null;
    if (k.customList && typeof v === 'string') { const i = k.customList.findIndex((x) => x.toLowerCase() === v.toLowerCase()); return i < 0 ? k.customList.length + 1 : i; }
    return v;
  };
  const sorted = [...lines].sort((a, b) => {
    for (const k of keys) {
      const x = keyVal(a, k), y = keyVal(b, k);
      // Blanks always go last, errors just before them.
      if (x === null || y === null) { if (x === y) continue; return x === null ? 1 : -1; }
      let c: number;
      if (isError(x) || isError(y)) c = isError(x) && isError(y) ? 0 : isError(x) ? 1 : -1;
      else if (caseSensitive && typeof x === 'string' && typeof y === 'string') c = x < y ? -1 : x > y ? 1 : 0;
      else c = compareValues(x, y);
      if (c) return k.descending && !(isError(x) || isError(y)) ? -c : c;
    }
    return a.idx - b.idx;
  });
  const area: CellRange = byCols ? { ...r, c1: start } : { ...r, r1: start };
  store.change('Sort', [{ area }], () => {
    sorted.forEach((ln, k) => {
      const dest = start + k;
      ln.cells.forEach((c, pos) => {
        const row = byCols ? span[0]! + pos : dest, col = byCols ? dest : span[0]! + pos;
        const key = cellKey(row, col);
        if (!c) { store.engine.clearCell(sheet, row, col); return; }
        const copy = structuredClone(c);
        if (copy.f) {
          const f = byCols ? shiftFormula(copy.f, 0, dest - ln.idx) : shiftFormula(copy.f, dest - ln.idx, 0);
          store.engine.setFormula(sheet, row, col, f);
          const placed = store.sheet.cells.get(key)!;
          placed.s = copy.s;
        } else {
          store.engine.setValue(sheet, row, col, copy.v);
          const placed = store.sheet.cells.get(key);
          if (placed) placed.s = copy.s; else if (copy.s) store.sheet.cells.set(key, { v: null, s: copy.s });
        }
      });
    });
  });
};

// ---------- AutoFilter ----------

export const toggleFilter = (store: GridStore): void => {
  const s = store.sheet;
  if (s.autoFilter) {
    store.change('Remove filter', 'sheet', () => {
      for (const row of s.rows.values()) if (row.filtered) { row.hidden = false; delete row.filtered; }
      delete s.autoFilter;
    });
    return;
  }
  const r = dataRange(store);
  store.change('Filter', 'sheet', () => { s.autoFilter = { range: r, columns: [] }; });
};

export const filterValues = (store: GridStore, col: number): { text: string; count: number }[] => {
  const af = store.sheet.autoFilter;
  if (!af) return [];
  const map = new Map<string, number>();
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    const t = displayText(store, r, col);
    map.set(t, (map.get(t) ?? 0) + 1);
  }
  return [...map].map(([text, count]) => ({ text, count })).sort((a, b) => {
    const x = parseNumber(a.text), y = parseNumber(b.text);
    if (a.text === '') return 1;
    if (b.text === '') return -1;
    return x !== null && y !== null ? x - y : a.text.localeCompare(b.text);
  });
};

const passes = (store: GridStore, f: FilterColumn, row: number): boolean => {
  const text = displayText(store, row, f.col);
  const v = store.value(row, f.col);
  if (f.values) return text === '' ? !!f.blanks || f.values.includes('') : f.values.includes(text);
  if (f.color) return store.style(row, f.col).fill?.color === f.color;
  if (f.custom) {
    const test = (rule: { operator: string; value: string }): boolean => {
      const target = rule.value;
      const num = parseNumber(target);
      if (rule.operator === 'equal' || rule.operator === 'notEqual') {
        const wild = new RegExp(`^${target.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i');
        const eq = typeof v === 'number' && num !== null ? v === num : wild.test(text);
        return rule.operator === 'equal' ? eq : !eq;
      }
      if (typeof v === 'number' && num !== null) return compareOp(rule.operator as never, v, num, num);
      const c = text.localeCompare(target, undefined, { sensitivity: 'base' });
      return compareOp(rule.operator as never, c, 0, 0);
    };
    const results = f.custom.rules.map(test);
    return f.custom.and ? results.every(Boolean) : results.some(Boolean);
  }
  if (f.top10) {
    const af = store.sheet.autoFilter!;
    const nums: number[] = [];
    for (let r = af.range.r1 + 1; r <= af.range.r2; r++) { const x = store.value(r, f.col); if (typeof x === 'number') nums.push(x); }
    if (typeof v !== 'number') return false;
    nums.sort((a, b) => (f.top10!.top ? b - a : a - b));
    const n = f.top10.percent ? Math.max(1, Math.floor((nums.length * f.top10.value) / 100)) : f.top10.value;
    const cut = nums[Math.min(nums.length, n) - 1];
    return cut !== undefined && (f.top10.top ? v >= cut : v <= cut);
  }
  return true;
};

/** Applies the sheet's filter conditions, hiding rows that don't pass. */
export const applyFilter = (store: GridStore): void => {
  const af = store.sheet.autoFilter;
  if (!af) return;
  for (let r = af.range.r1 + 1; r <= af.range.r2; r++) {
    const show = af.columns.every((f) => passes(store, f, r));
    const info = store.sheet.rows.get(r) ?? {};
    info.hidden = !show;
    info.filtered = !show || undefined;
    store.sheet.rows.set(r, info);
  }
};

export const setFilterColumn = (store: GridStore, f: FilterColumn | null, col: number): void => {
  const af = store.sheet.autoFilter;
  if (!af) return;
  store.change('Filter', 'sheet', () => {
    af.columns = af.columns.filter((c) => c.col !== col);
    if (f) af.columns.push(f);
    applyFilter(store);
  });
};

export const clearFilters = (store: GridStore): void => {
  const af = store.sheet.autoFilter;
  if (!af) return;
  store.change('Clear filter', 'sheet', () => { af.columns = []; applyFilter(store); });
};

// ---------- Other data tools ----------

export const removeDuplicates = (store: GridStore, r: CellRange, cols: number[], header: boolean): number => {
  const sheet = store.sheetIndex;
  const seen = new Set<string>();
  const keep: number[] = [];
  const first = r.r1 + (header ? 1 : 0);
  for (let row = first; row <= r.r2; row++) {
    const key = cols.map((c) => { const v = store.value(row, c); return typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${String(v)}`; }).join('\u0001');
    if (!seen.has(key)) { seen.add(key); keep.push(row); }
  }
  const removed = r.r2 - first + 1 - keep.length;
  if (!removed) return 0;
  store.change('Remove duplicates', [{ area: r }], () => {
    const rows = keep.map((row) => { const out: (Cell | undefined)[] = []; for (let c = r.c1; c <= r.c2; c++) out.push(structuredClone(store.cell(row, c))); return out; });
    for (let row = first; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) store.engine.clearCell(sheet, row, c);
    rows.forEach((cells, i) => cells.forEach((cell, j) => {
      if (!cell) return;
      const row = first + i, col = r.c1 + j;
      if (cell.f) store.engine.setFormula(sheet, row, col, shiftFormula(cell.f, row - keep[i]!, 0));
      else store.engine.setValue(sheet, row, col, cell.v);
      const placed = store.sheet.cells.get(cellKey(row, col));
      if (placed) placed.s = cell.s; else if (cell.s) store.sheet.cells.set(cellKey(row, col), { v: null, s: cell.s });
    }));
  });
  return removed;
};

export const textToColumns = (store: GridStore, opts: { delimiters: string[]; other?: string; treatConsecutive?: boolean; qualifier?: string; fixedWidths?: number[] }): void => {
  const r = store.range;
  const col = r.c1;
  const lines: string[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const v = store.value(row, col);
    const s = v === null ? '' : String(toText(v) as string);
    if (opts.fixedWidths?.length) {
      const parts: string[] = [];
      let at = 0;
      for (const w of [...opts.fixedWidths, Infinity]) { parts.push(s.slice(at, at + w).trim()); at += w; if (at >= s.length) break; }
      lines.push(parts);
      continue;
    }
    const delims = [...opts.delimiters, ...(opts.other ? [opts.other] : [])].map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const re = new RegExp(`(${delims.join('|') || ','})${opts.treatConsecutive ? '+' : ''}`);
    const q = opts.qualifier ?? '"';
    const parts: string[] = [];
    let cur = '', inQ = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i]!;
      if (q && ch === q) { inQ = !inQ; continue; }
      if (!inQ && re.test(ch)) {
        parts.push(cur); cur = '';
        if (opts.treatConsecutive) while (i + 1 < s.length && re.test(s[i + 1]!)) i++;
        continue;
      }
      cur += ch;
    }
    parts.push(cur);
    lines.push(parts);
  }
  const width = Math.max(...lines.map((l) => l.length));
  store.change('Text to columns', [{ area: { r1: r.r1, c1: col, r2: r.r2, c2: col + width - 1 } }], () => {
    lines.forEach((parts, i) => parts.forEach((p, j) => store.enter(r.r1 + i, col + j, p)));
  });
};

export interface FindOptions { text: string; matchCase?: boolean; wholeCell?: boolean; lookIn?: 'formulas' | 'values'; allSheets?: boolean; byColumns?: boolean }
export interface Found { sheet: number; row: number; col: number; text: string }

export const findAll = (store: GridStore, o: FindOptions): Found[] => {
  if (!o.text) return [];
  const pat = o.text.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\\\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?');
  const re = new RegExp(o.wholeCell ? `^${pat}$` : pat, o.matchCase ? '' : 'i');
  const out: Found[] = [];
  const sheets = o.allSheets ? store.wb.sheets.map((_, i) => i) : [store.sheetIndex];
  for (const si of sheets) {
    const s = store.wb.sheets[si]!;
    const keys = [...s.cells.keys()].sort((a, b) => (o.byColumns ? (a % MAX_COLS) - (b % MAX_COLS) || a - b : a - b));
    for (const k of keys) {
      const c = s.cells.get(k)!;
      const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
      const text = o.lookIn !== 'values' && c.f ? `=${c.f}` : c.v === null ? '' : displayText(store, row, col, si);
      if (text && re.test(text)) out.push({ sheet: si, row, col, text });
    }
  }
  return out;
};

export const replaceAll = (store: GridStore, o: FindOptions, replacement: string, only?: Found): number => {
  const hits = only ? [only] : findAll(store, { ...o, lookIn: 'formulas' });
  if (!hits.length) return 0;
  const pat = o.text.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  const re = new RegExp(o.wholeCell ? `^${pat}$` : pat, o.matchCase ? 'g' : 'gi');
  const bySheet = new Map<number, Found[]>();
  for (const h of hits) bySheet.set(h.sheet, [...(bySheet.get(h.sheet) ?? []), h]);
  store.change('Replace', 'book', () => {
    for (const h of hits) {
      const c = store.wb.sheets[h.sheet]!.cells.get(cellKey(h.row, h.col));
      if (!c) continue;
      const src = c.f ? `=${c.f}` : String(c.v ?? '');
      const next = src.replace(re, replacement.replace(/\$/g, '$$$$'));
      try { store.enter(h.row, h.col, next, h.sheet); } catch { /* keep the original if the new formula is broken */ }
    }
  });
  return hits.length;
};

/** Goal Seek: changes one input cell until a formula cell reaches a target. */
export const goalSeek = (store: GridStore, setCell: { row: number; col: number }, target: number, byCell: { row: number; col: number }): { ok: boolean; value: number; result: Scalar } => {
  const sheet = store.sheetIndex;
  const orig = store.value(byCell.row, byCell.col);
  const f = (x: number): number => {
    store.engine.setValue(sheet, byCell.row, byCell.col, x);
    store.engine.recalc();
    const v = store.value(setCell.row, setCell.col);
    return typeof v === 'number' ? v - target : NaN;
  };
  let x0 = typeof orig === 'number' ? orig : 0;
  let x1 = x0 === 0 ? 1 : x0 * 1.01;
  let f0 = f(x0), f1 = f(x1);
  let ok = false;
  for (let i = 0; i < 200 && Number.isFinite(f1); i++) {
    if (Math.abs(f1) < 1e-7) { ok = true; break; }
    const d = f1 - f0;
    const x2 = d === 0 ? x1 + (Math.random() - 0.5) : x1 - (f1 * (x1 - x0)) / d;
    x0 = x1; f0 = f1; x1 = x2; f1 = f(x1);
  }
  if (!ok && Math.abs(f1) < 1e-7) ok = true;
  const result = store.value(setCell.row, setCell.col);
  // Undoable: put the original back, then apply the answer as one step.
  store.engine.setValue(sheet, byCell.row, byCell.col, orig);
  store.engine.recalc();
  if (ok) store.change('Goal Seek', [{ area: { r1: byCell.row, c1: byCell.col, r2: byCell.row, c2: byCell.col } }], () => store.engine.setValue(sheet, byCell.row, byCell.col, +x1.toPrecision(15)));
  return { ok, value: x1, result };
};

/** Data > Group / Ungroup: outline levels on rows or columns, with collapse by hiding. */
export const groupLines = (store: GridStore, axis: 'row' | 'col', from: number, to: number, delta: 1 | -1): void => {
  store.change(delta > 0 ? 'Group' : 'Ungroup', 'sheet', () => {
    const map = axis === 'row' ? store.sheet.rows : store.sheet.cols;
    for (let i = from; i <= to; i++) {
      const info = map.get(i) ?? {};
      info.outline = Math.max(0, Math.min(7, (info.outline ?? 0) + delta));
      if (!info.outline) { delete info.outline; if (delta < 0) info.hidden = false; }
      map.set(i, info);
    }
  });
};

export const setLinesHidden = (store: GridStore, axis: 'row' | 'col', from: number, to: number, hidden: boolean): void => {
  store.change(hidden ? 'Hide' : 'Unhide', 'sheet', () => {
    const map = axis === 'row' ? store.sheet.rows : store.sheet.cols;
    for (let i = from; i <= to; i++) { const info = map.get(i) ?? {}; info.hidden = hidden; map.set(i, info); }
  });
};

export const errorText = (v: Scalar): string => (v instanceof CellError ? v.code : '');
