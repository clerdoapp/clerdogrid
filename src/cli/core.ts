// Workbook operations shared by the command-line tool and the MCP server. Everything runs through
// GridStore so edits behave exactly as they do in the app (typed-input parsing, reference updates).
import { GridStore } from '../app/store';
import { addSheet as addSheetOp, changeStructure, deleteSheet as deleteSheetOp, renameSheet as renameSheetOp, validSheetName } from '../app/ops/structure';
import { registerAllFunctions } from '../engine/functions';
import { functionNames, SPECIAL_FORMS } from '../engine/registry';
import { CATALOG, catalogByName, searchCatalog, signature, type FnInfo } from '../engine/catalog';
import { FormulaError, parseFormula, quoteSheet, walk } from '../engine/parser';
import { CellError, isError, MAX_COLS, MAX_ROWS, type Matrix, type Scalar } from '../engine/types';
import {
  addr, cellKey, colIndex, keyCol, keyRow, mergeStyle, newWorkbook, rangeAddr,
  type CellRange, type CellStyle, type Chart, type ChartType, type HAlign, type Workbook,
} from '../model/workbook';
import { TEMPLATES } from '../ui/dialogs/a/templates';
import { EXIT, UserError } from './errors';
import { displayText, jsonValue, scalarText, usedRange } from './io';

let registered = false;
export const ensureFunctions = (): void => {
  if (registered) return;
  registerAllFunctions();
  registered = true;
};

/** A workbook opened headlessly. Store notices (which the app shows as toasts) become errors. */
export const openStore = (wb: Workbook): GridStore => {
  ensureFunctions();
  const store = new GridStore(newWorkbook());
  store.notify = (text: string) => { throw new UserError(text); };
  store.load(wb, 'Book1');
  return store;
};

// ---------- Addresses ----------

export const sheetIndex = (store: GridStore, name?: string | number): number => {
  if (name === undefined || name === '') return store.wb.activeSheet;
  if (typeof name === 'number') {
    if (!store.wb.sheets[name]) throw new UserError(`There is no sheet number ${name + 1}.`);
    return name;
  }
  const i = store.wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());
  if (i >= 0) return i;
  if (/^\d+$/.test(name) && store.wb.sheets[+name - 1]) return +name - 1;
  throw new UserError(`There is no sheet named "${name}". Sheets: ${store.wb.sheets.map((s) => s.name).join(', ')}.`);
};

export interface Target { sheet: number; range: CellRange }

/**
 * Reads "B3", "A1:C9", "Sheet 2!A1:B4", "'My sheet'!A:A", "3:5" or a defined name.
 * Whole columns or rows are cut to the used area.
 */
export const parseTarget = (store: GridStore, text: string, sheetName?: string): Target => {
  let t = text.trim().replace(/\$/g, '');
  let sheet = sheetIndex(store, sheetName);
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    let sn = t.slice(0, bang);
    if (sn.startsWith("'") && sn.endsWith("'")) sn = sn.slice(1, -1).replace(/''/g, "'");
    sheet = sheetIndex(store, sn);
    t = t.slice(bang + 1);
  }
  const used = usedRange(store.engine, sheet) ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
  const cell = /^([A-Za-z]{1,3})(\d{1,7})$/;
  const parts = t.split(':');
  if (parts.length <= 2 && parts.every((p) => cell.test(p))) {
    const [a, b] = parts.map((p) => { const m = cell.exec(p)!; return { row: +m[2]! - 1, col: colIndex(m[1]!) }; });
    const e = b ?? a!;
    const r = { r1: Math.min(a!.row, e.row), c1: Math.min(a!.col, e.col), r2: Math.max(a!.row, e.row), c2: Math.max(a!.col, e.col) };
    if (r.r1 < 0 || r.c2 >= MAX_COLS || r.r2 >= MAX_ROWS) throw new UserError(`"${text}" is outside the sheet.`);
    return { sheet, range: r };
  }
  if (parts.length === 2 && parts.every((p) => /^[A-Za-z]{1,3}$/.test(p))) {
    const c1 = colIndex(parts[0]!), c2 = colIndex(parts[1]!);
    return { sheet, range: { r1: 0, c1: Math.min(c1, c2), r2: used.r2, c2: Math.max(c1, c2) } };
  }
  if (parts.length === 2 && parts.every((p) => /^\d+$/.test(p))) {
    const r1 = +parts[0]! - 1, r2 = +parts[1]! - 1;
    return { sheet, range: { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: used.c2 } };
  }
  const dn = store.wb.names.find((n) => n.name.toLowerCase() === t.toLowerCase() && (n.sheet === undefined || n.sheet === sheet));
  if (dn && !text.includes('!')) return parseTarget(store, dn.formula.replace(/^=/, ''), undefined);
  const tbl = store.wb.sheets.flatMap((s, i) => s.tables.map((tb) => ({ i, tb }))).find((x) => x.tb.name.toLowerCase() === t.toLowerCase());
  if (tbl) return { sheet: tbl.i, range: tbl.tb.range };
  throw new UserError(`"${text}" isn't a cell, range, defined name or table.`);
};

export const cellCount = (r: CellRange): number => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);

// ---------- Reading ----------

export const valueOut = (v: Scalar | Matrix): unknown => (Array.isArray(v) ? v.map((row) => row.map(jsonValue)) : jsonValue(v));
export const valueText = (v: Scalar | Matrix): string => (Array.isArray(v) ? v.map((row) => row.map(scalarText).join('\t')).join('\n') : scalarText(v));

export interface RangeData {
  sheet: string;
  range: string;
  values: (string | number | boolean | null)[][];
  text: string[][];
  formulas: (string | null)[][];
}

export const readRange = (store: GridStore, t: Target, limit = Infinity): RangeData => {
  if (cellCount(t.range) > limit) throw new UserError(`That range has ${cellCount(t.range)} cells; the limit is ${limit}. Read a smaller range.`);
  const s = store.wb.sheets[t.sheet]!;
  const values: RangeData['values'] = [], text: string[][] = [], formulas: (string | null)[][] = [];
  for (let r = t.range.r1; r <= t.range.r2; r++) {
    const vr: RangeData['values'][number] = [], tr: string[] = [], fr: (string | null)[] = [];
    for (let c = t.range.c1; c <= t.range.c2; c++) {
      vr.push(jsonValue(store.engine.getValue(t.sheet, r, c)));
      tr.push(displayText(store.wb, store.engine, t.sheet, r, c));
      const f = s.cells.get(cellKey(r, c))?.f;
      fr.push(f ? `=${f}` : null);
    }
    values.push(vr); text.push(tr); formulas.push(fr);
  }
  return { sheet: s.name, range: rangeAddr(t.range), values, text, formulas };
};

export const summary = (store: GridStore) => {
  const wb = store.wb;
  let formulas = 0, cells = 0;
  const sheets = wb.sheets.map((s, i) => {
    let f = 0, n = 0;
    for (const c of s.cells.values()) { if (c.f) f++; if (c.v !== null || c.f) n++; }
    formulas += f; cells += n;
    const used = usedRange(store.engine, i);
    return {
      name: s.name,
      state: s.state,
      used: used ? rangeAddr(used) : null,
      rows: used ? used.r2 + 1 : 0,
      cols: used ? used.c2 + 1 : 0,
      cells: n,
      formulas: f,
      tables: s.tables.map((t) => ({ name: t.name, range: rangeAddr(t.range), columns: t.columns.map((c) => c.name) })),
      charts: s.charts.map((c) => ({ id: c.id, type: c.type, title: c.title ?? null, series: c.series.length, at: addr(c.anchor.row, c.anchor.col) })),
      pictures: s.pictures.length,
      merges: s.merges.length,
      conditionalFormats: s.conditionalFormats.length,
      validations: s.validations.length,
      frozen: s.view.frozenRows || s.view.frozenCols ? { rows: s.view.frozenRows, cols: s.view.frozenCols } : null,
      protected: !!s.protection?.sheet,
    };
  });
  return {
    sheets,
    activeSheet: wb.sheets[wb.activeSheet]?.name ?? null,
    names: wb.names.map((n) => ({ name: n.name, refersTo: `=${n.formula.replace(/^=/, '')}`, scope: n.sheet === undefined ? null : wb.sheets[n.sheet]?.name ?? null })),
    totals: { sheets: wb.sheets.length, cells, formulas, styles: wb.styles.length },
    date1904: wb.date1904,
    properties: wb.props ?? {},
    scripts: wb.scripts?.length ?? 0,
  };
};

// ---------- Writing ----------

export type Input = string | number | boolean | null;

const putValue = (store: GridStore, sheet: number, row: number, col: number, v: Input): void => {
  if (typeof v === 'string') { store.enter(row, col, v, sheet); return; }
  if (typeof v === 'number' && !Number.isFinite(v)) { store.engine.setValue(sheet, row, col, new CellError('#NUM!')); return; }
  store.engine.setValue(sheet, row, col, v);
};

/** Strings are read as typed input: "=SUM(A1:A3)" is a formula, "12%" a number, "'007" text. */
export const writeCells = (store: GridStore, sheet: number, row: number, col: number, rows: Input[][]): CellRange => {
  if (!rows.length) throw new UserError('Nothing to write.');
  const width = Math.max(...rows.map((r) => r.length));
  const area = { r1: row, c1: col, r2: row + rows.length - 1, c2: col + Math.max(1, width) - 1 };
  if (area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS) throw new UserError('Those values would run off the end of the sheet.');
  if (store.wb.sheets[sheet]!.protection?.sheet) throw new UserError(`Sheet "${store.wb.sheets[sheet]!.name}" is protected.`);
  try {
    store.change('Write cells', [{ area, sheet }], () => {
      rows.forEach((r, i) => r.forEach((v, j) => putValue(store, sheet, row + i, col + j, v)));
    });
  } catch (e) {
    if (e instanceof FormulaError) throw new UserError(`That formula can't be read: ${e.message}`);
    throw e;
  }
  return area;
};

export const recalculate = (store: GridStore) => {
  store.engine.recalc(true);
  return { circular: store.engine.circular.map((c) => `${quoteSheet(store.wb.sheets[c.sheet]!.name)}!${addr(c.row, c.col)}`) };
};

export const evaluate = (store: GridStore, formula: string, sheet: number): Scalar | Matrix => {
  const f = formula.trim().replace(/^=/, '');
  if (!f) throw new UserError('Give a formula to calculate.');
  try { parseFormula(f); } catch (e) { throw new UserError(`That formula can't be read: ${e instanceof Error ? e.message : String(e)}`); }
  return store.engine.evaluate(f, sheet, 0, 0);
};

// ---------- Sheets and structure ----------

const withSheet = <T>(store: GridStore, sheet: number, fn: () => T): T => {
  const before = store.wb.activeSheet;
  store.wb.activeSheet = sheet;
  try { return fn(); } finally { store.wb.activeSheet = Math.min(before, store.wb.sheets.length - 1); }
};

export const addSheet = (store: GridStore, name?: string, at?: number): string => {
  if (store.wb.protection?.structure) throw new UserError('The workbook structure is protected.');
  const n = name?.trim();
  if (n) { const err = validSheetName(store, n); if (err) throw new UserError(err); }
  const pos = at === undefined ? store.wb.sheets.length : Math.max(0, Math.min(at, store.wb.sheets.length));
  const before = store.wb.activeSheet;
  addSheetOp(store, pos, n || undefined);
  const made = store.wb.sheets[pos]!.name;
  store.wb.activeSheet = before >= pos ? before + 1 : before;
  return made;
};

export const renameSheet = (store: GridStore, sheet: number, name: string): void => {
  if (store.wb.protection?.structure) throw new UserError('The workbook structure is protected.');
  renameSheetOp(store, sheet, name);
};

export const deleteSheet = (store: GridStore, sheet: number): void => {
  if (store.wb.protection?.structure) throw new UserError('The workbook structure is protected.');
  const active = store.wb.activeSheet;
  deleteSheetOp(store, sheet);
  store.wb.activeSheet = Math.max(0, Math.min(active > sheet ? active - 1 : active, store.wb.sheets.length - 1));
};

/** count > 0 inserts before `at` (0-based); count < 0 deletes from `at`. */
export const changeRows = (store: GridStore, sheet: number, axis: 'row' | 'col', at: number, count: number): void => {
  if (!Number.isInteger(at) || at < 0 || at >= (axis === 'row' ? MAX_ROWS : MAX_COLS)) throw new UserError(`That ${axis === 'row' ? 'row' : 'column'} is outside the sheet.`);
  if (!Number.isInteger(count) || count === 0 || Math.abs(count) > 100000) throw new UserError('Count must be a whole number from 1 to 100000.');
  if (store.wb.sheets[sheet]!.protection?.sheet) throw new UserError(`Sheet "${store.wb.sheets[sheet]!.name}" is protected.`);
  withSheet(store, sheet, () => changeStructure(store, axis, at, count));
};

// ---------- Formatting and charts ----------

export interface FormatPatch {
  numberFormat?: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  fontName?: string; fontSize?: number; color?: string; fill?: string | null;
  align?: HAlign; valign?: 'top' | 'center' | 'bottom'; wrap?: boolean; border?: 'none' | 'thin' | 'medium' | 'thick' | 'double';
}

const HEX = /^#?[0-9a-f]{6}$/i;
const hex = (c: string, what: string): string => {
  if (!HEX.test(c)) throw new UserError(`${what} must be a colour like #1F7A4D.`);
  return (c.startsWith('#') ? c : `#${c}`).toUpperCase();
};

export const formatStyle = (p: FormatPatch): ((s: CellStyle) => CellStyle) => {
  const over: CellStyle = {};
  const font: NonNullable<CellStyle['font']> = {};
  if (p.bold !== undefined) font.bold = p.bold;
  if (p.italic !== undefined) font.italic = p.italic;
  if (p.underline !== undefined) font.underline = p.underline ? 'single' : undefined;
  if (p.strike !== undefined) font.strike = p.strike;
  if (p.fontName) font.name = p.fontName.slice(0, 64);
  if (p.fontSize !== undefined) { if (!(p.fontSize >= 1 && p.fontSize <= 409)) throw new UserError('Font size must be from 1 to 409.'); font.size = p.fontSize; }
  if (p.color) font.color = hex(p.color, 'Text colour');
  if (Object.keys(font).length) over.font = font;
  if (p.numberFormat !== undefined) over.numFmt = p.numberFormat.slice(0, 255) || 'General';
  if (p.fill) over.fill = { color: hex(p.fill, 'Fill colour'), pattern: 'solid' };
  const align: NonNullable<CellStyle['align']> = {};
  if (p.align) align.h = p.align;
  if (p.valign) align.v = p.valign;
  if (p.wrap !== undefined) align.wrap = p.wrap;
  if (Object.keys(align).length) over.align = align;
  return (s) => {
    let out = mergeStyle(s, over);
    if (p.fill === null) delete out.fill;
    if (p.border) {
      const side = p.border === 'none' ? undefined : { style: p.border };
      out = { ...out, border: side ? { left: side, right: side, top: side, bottom: side } : undefined };
    }
    return JSON.parse(JSON.stringify(out)) as CellStyle;
  };
};

export const setFormat = (store: GridStore, t: Target, p: FormatPatch, limit = Infinity): void => {
  if (cellCount(t.range) > limit) throw new UserError(`That range has ${cellCount(t.range)} cells; the limit is ${limit}.`);
  const f = formatStyle(p);
  store.change('Format', [{ area: t.range, sheet: t.sheet }], () => store.styleRange(t.range, f, t.sheet));
};

export const CHART_TYPES: readonly ChartType[] = [
  'column', 'stackedColumn', 'percentColumn', 'bar', 'stackedBar', 'percentBar', 'line', 'stackedLine', 'lineMarkers', 'area', 'stackedArea',
  'pie', 'doughnut', 'scatter', 'scatterLines', 'bubble', 'histogram', 'pareto', 'boxWhisker', 'combo', 'treemap', 'sunburst', 'waterfall', 'radar', 'funnel',
];

/** Makes a chart from a block whose first row holds series names and first column the categories. */
export const createChart = (store: GridStore, t: Target, o: { type?: string; title?: string; at?: string; seriesInColumns?: boolean }): Chart => {
  const type = (o.type ?? 'column') as ChartType;
  if (!CHART_TYPES.includes(type)) throw new UserError(`Unknown chart type "${o.type}". Types: ${CHART_TYPES.join(', ')}.`);
  const r = t.range;
  if (r.r2 <= r.r1 && r.c2 <= r.c1) throw new UserError('A chart needs a range with at least two cells.');
  const sn = quoteSheet(store.wb.sheets[t.sheet]!.name);
  const abs = (r1: number, c1: number, r2: number, c2: number) => `${sn}!$${addr(r1, c1).replace(/(\d+)$/, '$$$1')}:$${addr(r2, c2).replace(/(\d+)$/, '$$$1')}`;
  const one = (row: number, col: number) => `${sn}!$${addr(row, col).replace(/(\d+)$/, '$$$1')}`;
  const textAt = (row: number, col: number) => typeof store.engine.getValue(t.sheet, row, col) === 'string';
  const byCols = o.seriesInColumns ?? (r.r2 - r.r1 >= r.c2 - r.c1);
  const anyText = (cells: [number, number][]) => cells.some(([row, col]) => textAt(row, col));
  const span = (a: number, b: number) => Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);
  // Series names along the first row (or column), categories down the first column (or along the first row).
  const header = byCols ? anyText(span(r.c1 + 1, r.c2).map((c) => [r.r1, c])) : anyText(span(r.r1 + 1, r.r2).map((rr) => [rr, r.c1]));
  const cats = byCols ? anyText(span(r.r1 + (header ? 1 : 0), r.r2).map((rr) => [rr, r.c1])) && r.c1 < r.c2 : anyText(span(r.c1 + (header ? 1 : 0), r.c2).map((c) => [r.r1, c])) && r.r1 < r.r2;
  const series: Chart['series'] = [];
  if (byCols) {
    const d0 = r.r1 + (header ? 1 : 0);
    for (let c = r.c1 + (cats ? 1 : 0); c <= r.c2; c++) series.push({ name: header ? one(r.r1, c) : undefined, values: abs(d0, c, r.r2, c), categories: cats ? abs(d0, r.c1, r.r2, r.c1) : undefined });
  } else {
    const d0 = r.c1 + (header ? 1 : 0);
    for (let rr = r.r1 + (cats ? 1 : 0); rr <= r.r2; rr++) series.push({ name: header ? one(rr, r.c1) : undefined, values: abs(rr, d0, rr, r.c2), categories: cats ? abs(r.r1, d0, r.r1, r.c2) : undefined });
  }
  if (!series.length) throw new UserError('That range has no numbers to chart.');
  const at = o.at ? parseTarget(store, o.at, store.wb.sheets[t.sheet]!.name).range : { r1: r.r1, c1: r.c2 + 2 };
  const sheet = store.wb.sheets[t.sheet]!;
  const id = `chart${Date.now().toString(36)}${sheet.charts.length + 1}`;
  const chart: Chart = { id, type, title: o.title?.slice(0, 255) || undefined, series, anchor: { row: at.r1, col: at.c1, dx: 0, dy: 0, width: 480, height: 288 }, legend: series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none' };
  store.change('Insert chart', 'book', () => { sheet.charts.push(chart); });
  return chart;
};

// ---------- Search, checks and comparison ----------

export interface FindHit { sheet: string; cell: string; value: string | number | boolean | null; text: string; formula: string | null }

export const find = (store: GridStore, query: string, o: { sheet?: number; matchCase?: boolean; whole?: boolean; formulas?: boolean; limit?: number } = {}): FindHit[] => {
  if (!query) throw new UserError('Give some text to find.');
  const hits: FindHit[] = [];
  const q = o.matchCase ? query : query.toLowerCase();
  const limit = o.limit ?? 1000;
  const sheets = o.sheet !== undefined ? [o.sheet] : store.wb.sheets.map((_, i) => i);
  for (const i of sheets) {
    const s = store.wb.sheets[i]!;
    const keys = [...s.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
    for (const k of keys) {
      const c = s.cells.get(k)!;
      const row = keyRow(k), col = keyCol(k);
      const text = displayText(store.wb, store.engine, i, row, col);
      const cands = [text, scalarText(store.engine.getValue(i, row, col))];
      if (o.formulas !== false && c.f) cands.push(`=${c.f}`);
      const hit = cands.some((x) => { const y = o.matchCase ? x : x.toLowerCase(); return o.whole ? y === q : y.includes(q); });
      if (hit) {
        hits.push({ sheet: s.name, cell: addr(row, col), value: jsonValue(store.engine.getValue(i, row, col)), text, formula: c.f ? `=${c.f}` : null });
        if (hits.length >= limit) return hits;
      }
    }
  }
  return hits;
};

export interface Problem { sheet: string; cell: string; kind: 'error' | 'ref' | 'circular' | 'parse' | 'unknownFunction'; detail: string; formula?: string }

export const validate = (store: GridStore): Problem[] => {
  const problems: Problem[] = [];
  const known = new Set(functionNames());
  const special = new Set<string>(SPECIAL_FORMS);
  const { circular } = recalculate(store);
  const circ = new Set(circular);
  store.wb.sheets.forEach((s, i) => {
    const keys = [...s.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
    for (const k of keys) {
      const c = s.cells.get(k)!;
      const cell = addr(keyRow(k), keyCol(k));
      const full = `${quoteSheet(s.name)}!${cell}`;
      const formula = c.f ? `=${c.f}` : undefined;
      if (c.f) {
        try {
          const node = parseFormula(c.f);
          const unknown = new Set<string>();
          const local = new Set<string>();
          walk(node, (n) => {
            if (n.type !== 'call') return;
            const fn = n.name.toUpperCase();
            const names = fn === 'LET' ? n.args.filter((_, j) => j % 2 === 0 && j < n.args.length - 1) : fn === 'LAMBDA' ? n.args.slice(0, -1) : [];
            for (const a of names) if (a.type === 'name') local.add(a.name.toUpperCase());
          });
          let badRef = false;
          walk(node, (n) => {
            if (n.type === 'err' && n.code === '#REF!') badRef = true;
            if (n.type === 'call' && !known.has(n.name.toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, '')) && !special.has(n.name.toUpperCase()) && !local.has(n.name.toUpperCase()) && !store.wb.names.some((d) => d.name.toUpperCase() === n.name.toUpperCase())) unknown.add(n.name.toUpperCase());
          });
          if (badRef) problems.push({ sheet: s.name, cell, kind: 'ref', detail: 'Refers to a deleted cell (#REF!).', formula });
          for (const u of unknown) problems.push({ sheet: s.name, cell, kind: 'unknownFunction', detail: `Unknown function ${u}.`, formula });
        } catch (e) {
          problems.push({ sheet: s.name, cell, kind: 'parse', detail: `Can't be read: ${e instanceof Error ? e.message : String(e)}`, formula });
        }
        if (circ.has(full)) problems.push({ sheet: s.name, cell, kind: 'circular', detail: 'Part of a circular reference.', formula });
      }
      const v = store.engine.getValue(i, keyRow(k), keyCol(k));
      if (isError(v) && !problems.some((p) => p.sheet === s.name && p.cell === cell && p.kind !== 'circular')) problems.push({ sheet: s.name, cell, kind: 'error', detail: `Shows ${v.code}.`, formula });
    }
  });
  return problems;
};

export interface Difference { sheet: string; cell?: string; kind: 'sheetAdded' | 'sheetRemoved' | 'value' | 'formula' | 'added' | 'removed'; a?: string; b?: string }

const cellDesc = (store: GridStore, sheet: number, k: number): string => {
  const c = store.wb.sheets[sheet]!.cells.get(k);
  if (!c) return '';
  return c.f ? `=${c.f}` : scalarText(store.engine.getValue(sheet, keyRow(k), keyCol(k)));
};

export const diff = (a: GridStore, b: GridStore, o: { valuesOnly?: boolean; limit?: number } = {}): Difference[] => {
  const out: Difference[] = [];
  const limit = o.limit ?? 10000;
  const names = (s: GridStore) => new Map(s.wb.sheets.map((x, i) => [x.name.toLowerCase(), i]));
  const na = names(a), nb = names(b);
  for (const [n, i] of na) if (!nb.has(n)) out.push({ sheet: a.wb.sheets[i]!.name, kind: 'sheetRemoved' });
  for (const [n, i] of nb) if (!na.has(n)) out.push({ sheet: b.wb.sheets[i]!.name, kind: 'sheetAdded' });
  for (const [n, ia] of na) {
    const ib = nb.get(n);
    if (ib === undefined) continue;
    const sa = a.wb.sheets[ia]!, sb = b.wb.sheets[ib]!;
    const keys = [...new Set([...sa.cells.keys(), ...sb.cells.keys()])].sort((x, y) => keyRow(x) - keyRow(y) || keyCol(x) - keyCol(y));
    for (const k of keys) {
      const ca = sa.cells.get(k), cb = sb.cells.get(k);
      const va = a.engine.getValue(ia, keyRow(k), keyCol(k)), vb = b.engine.getValue(ib, keyRow(k), keyCol(k));
      const emptyA = !ca || (ca.v === null && !ca.f), emptyB = !cb || (cb.v === null && !cb.f);
      if (emptyA && emptyB) continue;
      const cell = addr(keyRow(k), keyCol(k));
      if (emptyA) out.push({ sheet: sa.name, cell, kind: 'added', b: cellDesc(b, ib, k) });
      else if (emptyB) out.push({ sheet: sa.name, cell, kind: 'removed', a: cellDesc(a, ia, k) });
      else if (!o.valuesOnly && (ca!.f ?? '') !== (cb!.f ?? '')) out.push({ sheet: sa.name, cell, kind: 'formula', a: cellDesc(a, ia, k), b: cellDesc(b, ib, k) });
      else if (!sameScalar(va, vb)) out.push({ sheet: sa.name, cell, kind: 'value', a: scalarText(va), b: scalarText(vb) });
      if (out.length >= limit) return out;
    }
  }
  return out;
};

const sameScalar = (x: Scalar, y: Scalar): boolean => {
  if (x === y) return true;
  if (x instanceof CellError && y instanceof CellError) return x.code === y.code;
  if (typeof x === 'number' && typeof y === 'number') return Math.abs(x - y) <= 1e-9 * Math.max(1, Math.abs(x), Math.abs(y));
  return false;
};

// ---------- Functions and templates ----------

export const describeFunction = (name: string) => {
  const f = catalogByName(name);
  if (!f) {
    const q = name.trim().toUpperCase();
    const near = [...new Set([...CATALOG.filter((x) => editDistance(x.name, q) <= Math.max(1, Math.floor(q.length / 4))).map((x) => x.name), ...searchCatalog(name, 3).map((x) => x.name)])].slice(0, 5);
    throw new UserError(`There is no function named ${name.toUpperCase()}.${near.length ? ` Did you mean ${near.join(', ')}?` : ''}`);
  }
  return fnOut(f);
};

const editDistance = (a: string, b: string): number => {
  if (Math.abs(a.length - b.length) > 3) return 99;
  let prev = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j]! + 1, cur[j - 1]! + 1, prev[j - 1]! + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[b.length]!;
};

export const fnOut = (f: FnInfo) => ({ name: f.name, category: f.category, signature: signature(f), about: f.about, args: f.args.map((a) => ({ name: a.name, about: a.about, optional: !!a.optional, repeats: !!a.repeat })), returns: f.returns ?? null });

export const listFunctions = (query?: string, category?: string, limit = 10000): FnInfo[] => {
  let list = query ? searchCatalog(query, limit) : [...CATALOG];
  if (category) list = list.filter((f) => f.category.toLowerCase() === category.toLowerCase());
  return list.slice(0, limit);
};

export const templateList = () => [{ id: 'blank', name: 'Blank workbook', about: 'One empty sheet.' }, ...TEMPLATES.map((t) => ({ id: t.id, name: t.name, about: t.about }))];

export const fromTemplate = (id: string): Workbook => {
  ensureFunctions();
  if (!id || id === 'blank') return newWorkbook();
  const t = TEMPLATES.find((x) => x.id === id.toLowerCase());
  if (!t) throw new UserError(`Unknown template "${id}". Templates: ${templateList().map((x) => x.id).join(', ')}.`, EXIT.usage);
  return t.make();
};
