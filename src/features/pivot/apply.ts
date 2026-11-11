// PivotTables in the workbook: reading sources, drawing tables into cells, refreshing caches,
// drill-down and linked charts. Every change goes through store.change() so it can be undone.
import type { GridStore } from '../../app/store';
import { formatValue } from '../../formats/numfmt';
import { MAX_COLS, MAX_ROWS, isError, type Scalar } from '../../engine/types';
import { quoteSheet } from '../../engine/parser';
import {
  cellKey, colName, internStyle, mergeStyle, newSheet, rangeAddr,
  type CellRange, type CellStyle, type Chart, type Sheet, type Workbook,
} from '../../model/workbook';
import { VALUES_FIELD, type PivotCache, type PivotSource, type PivotTable } from '../../model/pivot';
import { buildCache, refreshCache, type SourceData } from './cache';
import { computePivot, type Line, type PCell, type PivotLayout } from './compute';
import { remapTable } from './model';
import { cellStyleFor } from './styles';
import { N, T } from '../../ui/i18n';

// ---------------------------------------------------------------- lookups

export interface PivotRef { sheet: number; index: number; pt: PivotTable }

export const allPivots = (wb: Workbook): PivotRef[] =>
  wb.sheets.flatMap((s, sheet) => (s.pivots ?? []).map((pt, index) => ({ sheet, index, pt })));

export const cacheOf = (wb: Workbook, id: number): PivotCache | undefined => wb.pivotCaches?.find((c) => c.id === id);

const layoutMemo = new WeakMap<PivotCache, Map<string, PivotLayout>>();
const sig = (pt: PivotTable): string => JSON.stringify({ ...pt, rowItems: undefined, colItems: undefined, location: undefined, formats: undefined, anchor: undefined });

/** The table's computed layout, memoized per cache object and definition. */
export const layoutOf = (wb: Workbook, pt: PivotTable): PivotLayout | null => {
  const cache = cacheOf(wb, pt.cacheId);
  if (!cache) return null;
  let m = layoutMemo.get(cache);
  if (!m) { m = new Map(); layoutMemo.set(cache, m); }
  const k = sig(pt);
  let l = m.get(k);
  if (!l) {
    l = computePivot(pt, cache, wb.date1904);
    if (m.size > 64) m.clear();
    m.set(k, l);
  }
  return l;
};

/** Body area (location) of a table; with `pages` the report filter rows above are included. */
export const pivotArea = (pt: PivotTable, l: PivotLayout, pages = false): CellRange => {
  const top = pt.anchor.row - (pages ? l.pageRows : 0);
  const pageWidth = pages ? Math.max(0, ...l.pageCells.map((p) => p.dc + 1)) : 0;
  return { r1: Math.max(0, top), c1: pt.anchor.col, r2: pt.anchor.row + Math.max(1, l.height) - 1, c2: pt.anchor.col + Math.max(1, l.width, pageWidth) - 1 };
};

/** Area a table occupied when last drawn. */
export const drawnArea = (pt: PivotTable): CellRange | null => {
  const loc = pt.location;
  if (!loc) return null;
  return { ...loc.ref, r1: Math.max(0, loc.ref.r1 - (loc.pageRows ?? 0)) };
};

export const pivotAt = (wb: Workbook, sheet: number, row: number, col: number): PivotRef | null => {
  const list = wb.sheets[sheet]?.pivots ?? [];
  for (let index = 0; index < list.length; index++) {
    const pt = list[index]!;
    const a = drawnArea(pt);
    if (a && row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2) return { sheet, index, pt };
  }
  return null;
};

/** The layout cell under a sheet cell, if it belongs to a table. */
export const pivotCellAt = (wb: Workbook, sheet: number, row: number, col: number): { ref: PivotRef; layout: PivotLayout; cell: PCell | null } | null => {
  const ref = pivotAt(wb, sheet, row, col);
  if (!ref) return null;
  const layout = layoutOf(wb, ref.pt);
  if (!layout) return null;
  const dr = row - ref.pt.anchor.row, dc = col - ref.pt.anchor.col;
  if (dr < 0) {
    const p = layout.pageCells.find((x) => x.dr === dr && x.dc === dc);
    return { ref, layout, cell: p?.cell ?? null };
  }
  return { ref, layout, cell: layout.cells[dr]?.[dc] ?? null };
};

export const nextPivotName = (wb: Workbook): string => {
  const taken = new Set(allPivots(wb).map((p) => p.pt.name.toLowerCase()));
  for (let i = 1; ; i++) if (!taken.has(`pivottable${i}`)) return `PivotTable${i}`;
};

export const nextCacheId = (wb: Workbook): number => Math.max(0, ...(wb.pivotCaches ?? []).map((c) => c.id)) + 1;

// ---------------------------------------------------------------- sources

export const sourceText = (src: PivotSource): string =>
  src.kind === 'name' ? src.name : `${quoteSheet(src.sheet)}!${rangeAddr(src.range)}`;

/** Resolves a source to a sheet range (header row first). */
export const resolveSource = (store: GridStore, src: PivotSource): { sheet: number; range: CellRange } | string => {
  if (src.kind === 'range') {
    const sheet = store.wb.sheets.findIndex((s) => s.name.toLowerCase() === src.sheet.toLowerCase());
    if (sheet < 0) return T('The sheet "{name}" no longer exists.', { name: src.sheet });
    return { sheet, range: src.range };
  }
  for (let i = 0; i < store.wb.sheets.length; i++) {
    const t = store.wb.sheets[i]!.tables.find((x) => x.name.toLowerCase() === src.name.toLowerCase());
    if (t) return { sheet: i, range: { ...t.range, r2: t.range.r2 - (t.totalsRow ? 1 : 0) } };
  }
  const dn = store.wb.names.find((n) => n.name.toLowerCase() === src.name.toLowerCase());
  if (dn) {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(dn.formula.trim().replace(/^=/, ''));
    if (m) {
      const sheet = store.engine.sheetIndex((m[1] ?? m[2] ?? '').replace(/''/g, "'"));
      const a = colIdx(m[3]!), b = colIdx(m[5] ?? m[3]!);
      if (sheet >= 0) return { sheet, range: { r1: +m[4]! - 1, c1: Math.min(a, b), r2: +(m[6] ?? m[4]!) - 1, c2: Math.max(a, b) } };
    }
    return T('The name "{name}" doesn\'t refer to a range.', { name: src.name });
  }
  return T('There\'s no table or name called "{name}".', { name: src.name });
};

const colIdx = (s: string): number => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };

/** Reads a source block: header row, data rows (trailing blank rows dropped) and each column's format. */
export const readSource = (store: GridStore, src: PivotSource): SourceData | string => {
  const at = resolveSource(store, src);
  if (typeof at === 'string') return at;
  const b = store.engine.usedBounds(at.sheet);
  const r = { ...at.range, r2: Math.min(at.range.r2, Math.max(at.range.r1, b.rows)), c2: Math.min(at.range.c2, Math.max(at.range.c1, b.cols)) };
  if (r.c2 - r.c1 > 16383) return T('The source is too wide.');
  if (r.r2 <= r.r1) return T('The source needs a header row and at least one row of data.');
  const header: Scalar[] = [];
  for (let c = r.c1; c <= r.c2; c++) header.push(store.value(r.r1, c, at.sheet));
  if (header.every((h) => h === null || h === '')) return T('The first row of the source needs column names.');
  const rows: Scalar[][] = [];
  for (let row = r.r1 + 1; row <= r.r2; row++) {
    const line: Scalar[] = [];
    for (let c = r.c1; c <= r.c2; c++) line.push(store.value(row, c, at.sheet));
    rows.push(line);
  }
  while (rows.length > 1 && rows[rows.length - 1]!.every((v) => v === null || v === '')) rows.pop();
  const formats = header.map((_, i) => store.style(r.r1 + 1, r.c1 + i, at.sheet).numFmt);
  return { header, rows, formats };
};

// ---------------------------------------------------------------- drawing

const lineKey = (l: Line | undefined): string => (l ? `${l.kind}:${l.fn ?? ''}:${l.data}:${l.path.map((p) => `${p.field}.${p.slot}`).join(',')}` : '');
const cellFormatKey = (l: PivotLayout, c: PCell): string => {
  switch (c.role) {
    case 'data': return `d|${lineKey(l.rowLines[c.rowLine!])}|${lineKey(l.colLines[c.colLine!])}`;
    case 'rowItem': return `r|${lineKey(l.rowLines[c.rowLine!])}|${c.field}`;
    case 'colItem': return `c|${lineKey(l.colLines[c.colLine!])}|${c.field}`;
    default: return `${c.role}|${c.field ?? ''}|${c.data ?? ''}`;
  }
};

/** Style ids last written into each table's cells, to tell the user's own formatting apart. */
const written = new WeakMap<Sheet, Map<string, Map<number, { key: string; s: number }>>>();

export interface OverlapIssue { kind: 'pivot' | 'cells'; range: CellRange }

/** Would drawing this table overwrite another table or cells with content? */
export const findOverlap = (store: GridStore, sheet: number, pt: PivotTable, l: PivotLayout): OverlapIssue | null => {
  const area = pivotArea(pt, l, true);
  const s = store.wb.sheets[sheet]!;
  for (const other of s.pivots ?? []) {
    if (other.name === pt.name) continue;
    const a = drawnArea(other);
    if (a && a.r1 <= area.r2 && a.r2 >= area.r1 && a.c1 <= area.c2 && a.c2 >= area.c1) return { kind: 'pivot', range: a };
  }
  const own = s.pivots?.find((p) => p.name === pt.name);
  const old = own ? drawnArea(own) : null;
  const inOld = (r: number, c: number) => !!old && r >= old.r1 && r <= old.r2 && c >= old.c1 && c <= old.c2;
  const cells = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
  const check = (k: number, r: number, c: number) => {
    if (inOld(r, c)) return false;
    const cell = s.cells.get(k);
    return !!cell && (cell.v !== null || !!cell.f);
  };
  if (cells > s.cells.size) {
    for (const [k] of s.cells) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2 && check(k, r, c)) return { kind: 'cells', range: area }; }
  } else for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) if (check(cellKey(r, c), r, c)) return { kind: 'cells', range: area };
  return null;
};

const clearArea = (store: GridStore, sheet: number, a: CellRange): void => {
  const s = store.wb.sheets[sheet]!;
  for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
    const k = cellKey(r, c);
    const cell = s.cells.get(k);
    if (!cell) continue;
    store.engine.setValue(sheet, r, c, null);
    const after = s.cells.get(k);
    if (after) { delete after.s; if (after.v === null && !after.f) s.cells.delete(k); }
  }
};

/**
 * Draws a table into its sheet and returns the definition with its new location and axis items.
 * Call inside store.change(). Keeps formatting the user added when preserveFormatting is on.
 */
export const drawPivot = (store: GridStore, sheet: number, pt0: PivotTable): PivotTable => {
  const wb = store.wb;
  const s = wb.sheets[sheet]!;
  let pt = pt0;
  const l = layoutOf(wb, pt);
  if (!l) return pt;
  // Room for report filters above the table.
  if (pt.anchor.row < l.pageRows) pt = { ...pt, anchor: { ...pt.anchor, row: l.pageRows } };

  // Formatting the user changed since the last draw.
  const byTable = written.get(s) ?? new Map<string, Map<number, { key: string; s: number }>>();
  written.set(s, byTable);
  const last = byTable.get(pt.name);
  const formats = new Map((pt.formats ?? []).map((f) => [f.key, f.style]));
  if (last && pt.preserveFormatting) {
    for (const [k, w] of last) {
      const now = s.cells.get(k)?.s ?? 0;
      if (now !== w.s) formats.set(w.key, now);
    }
  }
  const old = drawnArea(pt0);
  if (old) clearArea(store, sheet, old);

  const put = (r: number, c: number, cell: PCell, style: CellStyle, key: string, map: Map<number, { key: string; s: number }>) => {
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
    const v = cell.v;
    store.engine.setValue(sheet, r, c, v);
    const k = cellKey(r, c);
    const own = pt.preserveFormatting ? formats.get(key) : undefined;
    // Formatting read from a file may be partial: lay it over the table's own style.
    const sid = own !== undefined ? internStyle(wb, mergeStyle(style, wb.styles[own] ?? {})) : internStyle(wb, style);
    const existing = s.cells.get(k);
    if (existing) existing.s = sid || undefined;
    else if (sid) s.cells.set(k, { v: null, s: sid });
    map.set(k, { key, s: sid });
  };

  const map = new Map<number, { key: string; s: number }>();
  const { row: R, col: C } = pt.anchor;
  for (const p of l.pageCells) put(R + p.dr, C + p.dc, p.cell, cellStyleFor(pt.style, p.cell, false, false), cellFormatKey(l, p.cell), map);
  for (let i = 0; i < l.height; i++) for (let j = 0; j < l.width; j++) {
    const cell = l.cells[i]![j];
    if (!cell) continue;
    const bandRow = (cell.rowLine ?? 0) % 2 === 1, bandCol = (cell.colLine ?? 0) % 2 === 1;
    put(R + i, C + j, cell, cellStyleFor(pt.style, cell, bandRow, bandCol, pt.showDrill), cellFormatKey(l, cell), map);
  }
  byTable.set(pt.name, map);

  if (pt.autoFit) autoFitColumns(store, sheet, pt, l);
  const ref: CellRange = { r1: R, c1: C, r2: R + Math.max(1, l.height) - 1, c2: C + Math.max(1, l.width) - 1 };
  const next: PivotTable = {
    ...pt,
    location: { ref, firstHeaderRow: l.firstHeaderRow, firstDataRow: l.firstDataRow, firstDataCol: l.firstDataCol, pageRows: l.pageRows || undefined },
    rowItems: l.rowItems, colItems: l.colItems,
    formats: formats.size ? [...formats].map(([key, style]) => ({ key, style })) : undefined,
  };
  updateLinkedCharts(store, sheet, next, l);
  return next;
};

const autoFitColumns = (store: GridStore, sheet: number, pt: PivotTable, l: PivotLayout): void => {
  const s = store.wb.sheets[sheet]!;
  const widths: number[] = [];
  const text = (c: PCell): string => (c.v === null ? '' : typeof c.v === 'number' ? formatValue(c.v, c.numFmt ?? 'General', { date1904: store.wb.date1904 }).text : isError(c.v) ? c.v.code : String(c.v));
  for (let j = 0; j < l.width; j++) {
    let w = 0;
    for (let i = 0; i < l.height; i++) { const c = l.cells[i]![j]; if (c && c.role !== 'colCaption') w = Math.max(w, text(c).length + (c.indent ?? 0) * 2 + (c.button ? 2 : 0) + (c.filter !== undefined ? 3 : 0)); }
    widths[j] = w;
  }
  for (const p of l.pageCells) widths[p.dc] = Math.max(widths[p.dc] ?? 0, text(p.cell).length + 3);
  widths.forEach((w, j) => {
    const col = pt.anchor.col + j;
    const chars = Math.max(8.43, Math.min(60, w * 1.05 + 1.5));
    const info = s.cols.get(col) ?? {};
    s.cols.set(col, { ...info, width: Math.round(chars * 100) / 100 });
  });
};

const replacePivot = (store: GridStore, sheet: number, name: string, pt: PivotTable): void => {
  const s = store.wb.sheets[sheet]!;
  const list = [...(s.pivots ?? [])];
  const i = list.findIndex((p) => p.name === name);
  if (i >= 0) list[i] = pt; else list.push(pt);
  s.pivots = list;
};

// ---------------------------------------------------------------- changes

const overlapOk = (store: GridStore, sheet: number, pt: PivotTable, ask: boolean): boolean => {
  const l = layoutOf(store.wb, pt);
  if (!l) return false;
  const issue = findOverlap(store, sheet, pt.anchor.row < l.pageRows ? { ...pt, anchor: { ...pt.anchor, row: l.pageRows } } : pt, l);
  if (!issue) return true;
  if (issue.kind === 'pivot') { store.notify(T('A PivotTable can\'t overlap another PivotTable.')); return false; }
  if (!ask) return false;
  return typeof confirm === 'function' ? confirm(T('There\'s already data in {p0}. Replace it?', { p0: rangeAddr(issue.range) })) : true;
};

/**
 * Changes a table's definition and redraws it as one undo step. Returns false when nothing changed
 * (the table is missing or the user cancelled an overlap).
 */
export const updatePivot = (store: GridStore, sheet: number, name: string, fn: (pt: PivotTable) => PivotTable, label = N('Change PivotTable'), ask = true): boolean => {
  const cur = store.wb.sheets[sheet]?.pivots?.find((p) => p.name === name);
  if (!cur) return false;
  const next = fn(cur);
  if (next === cur) return false;
  if (next.deferLayout && cur.deferLayout) {
    store.change(label, sheet === store.sheetIndex ? 'sheet' : 'book', () => replacePivot(store, sheet, name, next));
    return true;
  }
  if (!overlapOk(store, sheet, next, ask)) return false;
  store.change(label, sheet === store.sheetIndex ? 'sheet' : 'book', () => replacePivot(store, sheet, name, drawPivot(store, sheet, next)));
  return true;
};

/** Draws a table again (after deferred changes), optionally changing settings that don't need a check. */
export const redrawPivot = (store: GridStore, sheet: number, name: string, patch: Partial<PivotTable> = {}): boolean => {
  const cur = store.wb.sheets[sheet]?.pivots?.find((p) => p.name === name);
  if (!cur) return false;
  const next = { ...cur, ...patch, deferLayout: 'deferLayout' in patch ? patch.deferLayout : cur.deferLayout };
  if (!overlapOk(store, sheet, next, true)) return false;
  store.change(N('Update PivotTable'), sheet === store.sheetIndex ? 'sheet' : 'book', () => replacePivot(store, sheet, name, drawPivot(store, sheet, next)));
  return true;
};

/** Replaces a cache (after grouping or adding fields) and redraws every table on it. */
export const replaceCache = (store: GridStore, cache: PivotCache, remap: { fields: number[]; items: number[][] } | null, label: string, edit?: { sheet: number; name: string; fn: (pt: PivotTable) => PivotTable }): void => {
  store.change(label, 'book', () => {
    store.wb.pivotCaches = (store.wb.pivotCaches ?? []).map((c) => (c.id === cache.id ? cache : c));
    for (const ref of allPivots(store.wb)) {
      if (ref.pt.cacheId !== cache.id) continue;
      let pt = remap ? remapTable(ref.pt, remap.fields, remap.items, cache) : { ...ref.pt, fields: cache.fields.map((_, i) => ref.pt.fields[i] ?? { items: [] }) };
      if (edit && edit.sheet === ref.sheet && edit.name === ref.pt.name) pt = edit.fn(pt);
      replacePivot(store, ref.sheet, ref.pt.name, drawPivot(store, ref.sheet, pt));
    }
  });
};

/** Re-reads the sources of the given caches (all when omitted) and redraws their tables. */
export const refreshPivots = (store: GridStore, cacheIds?: number[], label = N('Refresh')): string | null => {
  const caches = (store.wb.pivotCaches ?? []).filter((c) => !cacheIds || cacheIds.includes(c.id));
  const fresh: { cache: PivotCache; remap: ReturnType<typeof refreshCache>['remap'] }[] = [];
  for (const c of caches) {
    const data = readSource(store, c.source);
    if (typeof data === 'string') return data;
    fresh.push(refreshCache(c, data, store.wb.date1904));
  }
  if (!freshFits(store, fresh)) return '';
  store.change(label, 'book', () => applyFresh(store, fresh));
  return null;
};

/** Checks every table a refresh would redraw for collisions; asks before overwriting cells. */
const freshFits = (store: GridStore, fresh: { cache: PivotCache; remap: { fields: number[]; items: number[][] } }[]): boolean => {
  let asked = false;
  for (const { cache, remap } of fresh) {
    for (const ref of allPivots(store.wb)) {
      if (ref.pt.cacheId !== cache.id) continue;
      const pt = remapTable(ref.pt, remap.fields, remap.items, cache);
      const l = computePivot(pt, cache, store.wb.date1904);
      const at = pt.anchor.row < l.pageRows ? { ...pt, anchor: { ...pt.anchor, row: l.pageRows } } : pt;
      const issue = findOverlap(store, ref.sheet, at, l);
      if (!issue) continue;
      if (issue.kind === 'pivot') { store.notify(T('Refreshing {name} would overlap another PivotTable. Move one of them first.', { name: pt.name })); return false; }
      if (!asked) {
        asked = true;
        if (typeof confirm === 'function' && !confirm(T('Refreshing {name} would replace data in {p1}. Replace it?', { name: pt.name, p1: rangeAddr(issue.range) }))) return false;
      }
    }
  }
  return true;
};

const applyFresh = (store: GridStore, fresh: { cache: PivotCache; remap: { fields: number[]; items: number[][] } }[]): void => {
  for (const { cache, remap } of fresh) {
    store.wb.pivotCaches = (store.wb.pivotCaches ?? []).map((c) => (c.id === cache.id ? cache : c));
    for (const ref of allPivots(store.wb)) {
      if (ref.pt.cacheId !== cache.id) continue;
      const pt = remapTable(ref.pt, remap.fields, remap.items, cache);
      replacePivot(store, ref.sheet, ref.pt.name, drawPivot(store, ref.sheet, pt));
    }
  }
};

/** Refresh for caches marked "refresh when opening", without an undo step. */
const openedBooks = new WeakSet<Workbook>();
export const refreshOnOpen = (store: GridStore): void => {
  const wb = store.wb;
  if (openedBooks.has(wb)) return;
  openedBooks.add(wb);
  const ids = (wb.pivotCaches ?? []).filter((c) => c.refreshOnLoad).map((c) => c.id);
  if (!ids.length) return;
  const fresh: { cache: PivotCache; remap: { fields: number[]; items: number[][] } }[] = [];
  for (const c of (wb.pivotCaches ?? []).filter((x) => ids.includes(x.id))) {
    const data = readSource(store, c.source);
    if (typeof data !== 'string') fresh.push(refreshCache(c, data, wb.date1904));
  }
  if (!fresh.length) return;
  const dirty = store.file.dirty;
  applyFresh(store, fresh);
  store.engine.recalc();
  store.file.dirty = dirty;
  store.touch();
};

/** Points a table's cache at a new source and refreshes it. Other tables sharing the cache follow. */
export const changeSource = (store: GridStore, cacheId: number, source: PivotSource): string | null => {
  const old = cacheOf(store.wb, cacheId);
  if (!old) return T('This PivotTable has no data.');
  const data = readSource(store, source);
  if (typeof data === 'string') return data;
  const fresh = refreshCache({ ...old, source }, data, store.wb.date1904);
  if (!freshFits(store, [fresh])) return '';
  store.change(N('Change data source'), 'book', () => applyFresh(store, [fresh]));
  return null;
};

export interface CreateOptions { source: PivotSource; dest: { sheet: number; row: number; col: number } | null; name?: string; useCache?: number; build?: (pt: PivotTable, cache: PivotCache) => PivotTable }

/** Creates a table on a new sheet (dest null) or at a cell. Returns the place it went, or an error message. */
export const createPivot = (store: GridStore, o: CreateOptions, newTable: (name: string, cache: PivotCache, anchor: { row: number; col: number }) => PivotTable): { sheet: number; name: string } | string => {
  let cache = o.useCache !== undefined ? cacheOf(store.wb, o.useCache) : undefined;
  if (!cache) {
    const data = readSource(store, o.source);
    if (typeof data === 'string') return data;
    cache = buildCache(nextCacheId(store.wb), o.source, data);
  }
  const name = o.name ?? nextPivotName(store.wb);
  let pt = newTable(name, cache, o.dest ? { row: o.dest.row, col: o.dest.col } : { row: 2, col: 0 });
  if (o.build) pt = o.build(pt, cache);
  if (o.dest) {
    const l = computePivot(pt, cache, store.wb.date1904);
    const issue = findOverlap(store, o.dest.sheet, pt, l);
    if (issue?.kind === 'pivot') return T("A PivotTable can't overlap another PivotTable.");
    if (issue && typeof confirm === 'function' && !confirm(T('There\'s already data in {p0}. Replace it?', { p0: rangeAddr(issue.range) }))) return '';
  }
  let at = o.dest?.sheet ?? -1;
  const finalCache = cache;
  store.change(N('Insert PivotTable'), 'book', () => {
    if (!store.wb.pivotCaches?.some((c) => c.id === finalCache.id)) store.wb.pivotCaches = [...(store.wb.pivotCaches ?? []), finalCache];
    if (at < 0) {
      const base = 'Sheet';
      const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
      let n = store.wb.sheets.length + 1;
      while (taken.has(`${base}${n}`.toLowerCase())) n++;
      store.wb.sheets.splice(store.sheetIndex, 0, newSheet(`${base}${n}`));
      store.engine.sheetsChanged();
      at = store.sheetIndex;
    }
    store.wb.activeSheet = at;
    replacePivot(store, at, name, drawPivot(store, at, pt));
  });
  const placed = store.wb.sheets[at]?.pivots?.find((p) => p.name === name);
  if (placed) store.selectRange({ r1: placed.anchor.row, c1: placed.anchor.col, r2: placed.anchor.row, c2: placed.anchor.col });
  return { sheet: at, name };
};

export const deletePivot = (store: GridStore, sheet: number, name: string): void => {
  store.change(N('Delete PivotTable'), 'book', () => {
    const s = store.wb.sheets[sheet]!;
    const pt = s.pivots?.find((p) => p.name === name);
    if (!pt) return;
    const a = drawnArea(pt);
    if (a) clearArea(store, sheet, a);
    s.pivots = (s.pivots ?? []).filter((p) => p.name !== name);
    if (!s.pivots.length) delete s.pivots;
    // Drop caches nothing uses any more.
    const used = new Set(allPivots(store.wb).map((p) => p.pt.cacheId));
    for (const sc of store.wb.slicerCaches ?? []) if (sc.source.kind === 'pivot') used.add(sc.source.cacheId);
    store.wb.pivotCaches = (store.wb.pivotCaches ?? []).filter((c) => used.has(c.id));
    if (!store.wb.pivotCaches.length) delete store.wb.pivotCaches;
  });
};

/** Moves a table to a new top-left cell. */
export const movePivot = (store: GridStore, sheet: number, name: string, to: { sheet: number; row: number; col: number }): boolean => {
  const s = store.wb.sheets[sheet]!;
  const pt = s.pivots?.find((p) => p.name === name);
  if (!pt) return false;
  const moved: PivotTable = { ...pt, anchor: { row: to.row, col: to.col }, location: undefined };
  const l = layoutOf(store.wb, moved)!;
  const issue = findOverlap(store, to.sheet, { ...moved, name: to.sheet === sheet ? '\u0000' : moved.name }, l);
  const ownArea = drawnArea(pt);
  if (issue && !(issue.kind === 'cells' && to.sheet === sheet && ownArea)) {
    if (issue.kind === 'pivot') { store.notify(T('A PivotTable can\'t overlap another PivotTable.')); return false; }
    if (typeof confirm === 'function' && !confirm(T('There\'s already data in {p0}. Replace it?', { p0: rangeAddr(issue.range) }))) return false;
  }
  store.change(N('Move PivotTable'), 'book', () => {
    if (ownArea) clearArea(store, sheet, ownArea);
    s.pivots = (s.pivots ?? []).filter((p) => p.name !== name);
    replacePivot(store, to.sheet, name, drawPivot(store, to.sheet, moved));
    store.wb.activeSheet = to.sheet;
  });
  return true;
};

// ---------------------------------------------------------------- drill-down

/** Puts the source records behind a value cell on a new sheet, as a table. */
export const drillDown = (store: GridStore, sheet: number, row: number, col: number): boolean => {
  const hit = pivotCellAt(store.wb, sheet, row, col);
  if (!hit?.cell || hit.cell.role !== 'data' || hit.cell.rowLine === undefined || hit.cell.colLine === undefined) return false;
  const cache = cacheOf(store.wb, hit.ref.pt.cacheId);
  if (!cache) return false;
  const recs = hit.layout.recordsAt(hit.cell.rowLine, hit.cell.colLine);
  const dbFields = cache.fields.filter((f) => f.database !== false);
  store.change(N('Show details'), 'book', () => {
    const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
    let n = store.wb.sheets.length + 1;
    while (taken.has(`sheet${n}`)) n++;
    const at = store.sheetIndex;
    store.wb.sheets.splice(at, 0, newSheet(`Sheet${n}`));
    store.engine.sheetsChanged();
    store.wb.activeSheet = at;
    dbFields.forEach((f, c) => store.engine.setValue(at, 0, c, f.name));
    recs.forEach((r, i) => dbFields.forEach((f, c) => {
      store.engine.setValue(at, i + 1, c, cache.records[r]![c] ?? null);
      const fmt = f.numFmt ?? (f.dates ? 'm/d/yyyy' : undefined);
      if (fmt) { const cell = store.wb.sheets[at]!.cells.get(cellKey(i + 1, c)); if (cell) cell.s = internStyle(store.wb, mergeStyle({}, { numFmt: fmt })); }
    }));
    const tables = new Set(store.wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())));
    let tn = 1;
    while (tables.has(`table${tn}`)) tn++;
    store.wb.sheets[at]!.tables.push({
      name: `Table${tn}`, range: { r1: 0, c1: 0, r2: Math.max(1, recs.length), c2: Math.max(0, dbFields.length - 1) }, headerRow: true, totalsRow: false,
      columns: dbFields.map((f) => ({ name: f.name })), style: 'TableStyleMedium7', bandedRows: true,
    });
    dbFields.forEach((f, c) => store.wb.sheets[at]!.cols.set(c, { width: Math.min(40, Math.max(10, f.name.length + 4)) }));
  });
  store.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
  return true;
};

// ---------------------------------------------------------------- PivotCharts

/**
 * Series for a chart linked to a table: one per column line, leaf rows only. Contiguous leaf rows
 * use cell ranges; otherwise the values travel as arrays, refreshed whenever the table changes.
 */
export const pivotChartSeries = (store: GridStore, sheet: number, pt: PivotTable, l: PivotLayout): Chart['series'] => {
  const sh = quoteSheet(store.wb.sheets[sheet]!.name);
  const { row: R, col: C } = pt.anchor;
  let lines = l.rowLines.map((rl, i) => ({ rl, i })).filter(({ rl }) => rl.kind === 'item' && (!rl.expandable || rl.expandable === 'collapsed'));
  if (!lines.length) lines = l.rowLines.map((rl, i) => ({ rl, i })).filter(({ rl }) => rl.kind !== 'blank');
  const rows = lines.map(({ i }) => R + l.firstDataRow + i);
  const deep = lines.some(({ rl }) => rl.path.filter((p) => p.field !== VALUES_FIELD).length > 1);
  const contiguous = !deep && rows.every((r, i) => i === 0 || r === rows[i - 1]! + 1);
  const literal = (vals: (string | number | null)[]) => `={${vals.map((v) => (typeof v === 'number' ? String(v) : `"${String(v ?? '').replace(/"/g, '""')}"`)).join(',')}}`;
  const colRange = (c: number) => `${sh}!$${colName(c)}$${rows[0]! + 1}:$${colName(c)}$${rows[rows.length - 1]! + 1}`;
  let categories: string | undefined;
  let multiLevel: Chart['series'][number]['multiLevel'];
  if (rows.length && l.firstDataCol > 0) {
    const labels = lines.map(({ rl }) => rl.path.filter((p) => p.field !== VALUES_FIELD).map((p) => l.caption(p.field, p.slot)));
    categories = contiguous ? colRange(C) : literal(labels.map((p, i) => p.join(' ') || (lines[i]!.rl.kind === 'grand' ? 'Total' : '')));
    if (deep) {
      // One level per row field, innermost first; outer labels only where they change.
      const depth = Math.max(...labels.map((p) => p.length));
      const levels: string[][] = [];
      for (let d = depth - 1; d >= 0; d--) {
        levels.push(labels.map((p, i) => {
          const v = p[d] ?? '';
          if (d === depth - 1) return v;
          const prev = labels[i - 1];
          return !prev || prev.slice(0, d + 1).join('\u0000') !== p.slice(0, d + 1).join('\u0000') ? v : '';
        }));
      }
      multiLevel = { ref: `${sh}!$${colName(C)}$${rows[0]! + 1}:$${colName(C + l.firstDataCol - 1)}$${rows[rows.length - 1]! + 1}`, levels };
    }
  }
  const series: Chart['series'] = [];
  l.colLines.forEach((cl, j) => {
    if (cl.kind === 'grand' && l.colLines.some((x) => x.kind !== 'grand')) return;
    if (cl.kind === 'sub' || cl.kind === 'blank') return;
    const c = C + l.firstDataCol + j;
    const headerRow = R + Math.max(0, l.firstDataRow - 1);
    const values = !rows.length ? '' : contiguous ? colRange(c) : literal(lines.map(({ i }) => { const v = l.cells[l.firstDataRow + i]?.[l.firstDataCol + j]?.v ?? null; return typeof v === 'number' ? v : null; }));
    series.push({ name: `=${sh}!$${colName(c)}$${headerRow + 1}`, values, categories, ...(multiLevel ? { multiLevel } : {}) });
  });
  return series;
};

const updateLinkedCharts = (store: GridStore, sheet: number, pt: PivotTable, l: PivotLayout): void => {
  const key = `${store.wb.sheets[sheet]!.name}!${pt.name}`.toLowerCase();
  for (const s of store.wb.sheets) for (const ch of s.charts) {
    if (ch.pivot?.toLowerCase() !== key) continue;
    ch.series = pivotChartSeries(store, sheet, pt, l).map((x, i) => ({ ...x, color: ch.series[i]?.color, type: ch.series[i]?.type }));
  }
};

export const pivotKey = (wb: Workbook, sheet: number, pt: PivotTable): string => `${wb.sheets[sheet]!.name}!${pt.name}`;

export const findByKey = (wb: Workbook, key: string): PivotRef | null => {
  const i = key.lastIndexOf('!');
  const sheetName = key.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"), name = key.slice(i + 1);
  const sheet = wb.sheets.findIndex((s) => s.name.toLowerCase() === sheetName.toLowerCase());
  const index = wb.sheets[sheet]?.pivots?.findIndex((p) => p.name.toLowerCase() === name.toLowerCase()) ?? -1;
  return index >= 0 ? { sheet, index, pt: wb.sheets[sheet]!.pivots![index]! } : null;
};

export const fieldsOnAxes = (pt: PivotTable): number[] => [...pt.rows, ...pt.cols].filter((f) => f !== VALUES_FIELD);
