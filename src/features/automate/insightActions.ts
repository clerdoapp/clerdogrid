import { isDateFormat } from '../../formats/numfmt';
import { currentRegion, looksLikeHeader } from '../../app/ops/data';
import type { GridStore } from '../../app/store';
import { CellError, type Scalar } from '../../engine/types';
import { colName, internStyle, newSheet, type CellRange, type Chart, type ChartSeries, type ChartType } from '../../model/workbook';
import { refString } from '../../ui/charts/chartData';
import { analyze, type Ask, type Insights } from './insights';

export interface DataBlock { sheet: number; sheetName: string; range: CellRange; headers: string[]; rows: Scalar[][]; hasHeader: boolean }

const MAX_ROWS_READ = 50_000;
const MAX_COLS_READ = 100;

/** The selection, or the block of data around the active cell, with its header row. */
export const readBlock = (store: GridStore): DataBlock | null => {
  const sel = store.range;
  const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
  let r = single ? currentRegion(store) : { ...sel };
  const b = store.engine.usedBounds(store.sheetIndex);
  if (b.rows < 0) return null;
  r = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, b.rows, r.r1 + MAX_ROWS_READ), c2: Math.min(r.c2, b.cols, r.c1 + MAX_COLS_READ - 1) };
  if (r.r2 < r.r1 || r.c2 < r.c1) return null;
  const hasHeader = looksLikeHeader(store, r) || (r.r2 > r.r1 && Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => store.value(r.r1, r.c1 + i)).every((v) => typeof v === 'string' && v !== ''));
  const headers: string[] = [];
  for (let c = r.c1; c <= r.c2; c++) {
    const v = hasHeader ? store.value(r.r1, c) : null;
    headers.push(v === null || v === '' ? `Column ${colName(c)}` : v instanceof CellError ? v.code : String(v));
  }
  const rows: Scalar[][] = [];
  for (let row = r.r1 + (hasHeader ? 1 : 0); row <= r.r2; row++) {
    if (store.sheet.rows.get(row)?.hidden) continue;
    const line: Scalar[] = [];
    for (let c = r.c1; c <= r.c2; c++) line.push(store.value(row, c));
    rows.push(line);
  }
  if (!rows.length) return null;
  return { sheet: store.sheetIndex, sheetName: store.sheet.name, range: r, headers, rows, hasHeader };
};

export const analyzeBlock = (store: GridStore, block: DataBlock): Insights => {
  const firstData = block.range.r1 + (block.hasHeader ? 1 : 0);
  const isDate = (col: number): boolean => {
    for (let row = firstData; row <= Math.min(block.range.r2, firstData + 20); row++) {
      if (typeof store.value(row, block.range.c1 + col, block.sheet) === 'number') return isDateFormat(store.style(row, block.range.c1 + col, block.sheet).numFmt ?? 'General');
    }
    return false;
  };
  return analyze(block.headers, block.rows, isDate);
};

// ---------- Inserting results ----------

const newId = () => `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

const uniqueName = (store: GridStore, base: string): string => {
  const clean = base.replace(/[\\/?*[\]:]/g, ' ').replace(/^'+|'+$/g, '').replace(/\s+/g, ' ').trim().slice(0, 28) || 'Summary';
  const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
  if (!taken.has(clean.toLowerCase())) return clean;
  for (let i = 2; ; i++) { const n = `${clean.slice(0, 27)} ${i}`; if (!taken.has(n.toLowerCase())) return n; }
};

/** Column `i` of the block's data rows (without the header). */
const dataCol = (block: DataBlock, i: number): string => {
  const top = block.range.r1 + (block.hasHeader ? 1 : 0);
  return refString(block.sheetName, { r1: top, c1: block.range.c1 + i, r2: block.range.r2, c2: block.range.c1 + i });
};
const headCell = (block: DataBlock, i: number): string | undefined =>
  block.hasHeader ? refString(block.sheetName, { r1: block.range.r1, c1: block.range.c1 + i, r2: block.range.r1, c2: block.range.c1 + i }) : undefined;

const AGG_LAMBDA: Record<string, string> = { sum: 'SUM', average: 'AVERAGE', count: 'COUNTA', min: 'MIN', max: 'MAX' };
const AGG_LABEL: Record<string, string> = { sum: 'Total', average: 'Average', count: 'Count', min: 'Lowest', max: 'Highest' };

export interface SummarySpec { by: number; val?: number; agg: 'sum' | 'average' | 'count' | 'min' | 'max'; top?: number; bottom?: boolean; chart?: ChartType; title: string }

/** The GROUPBY formula for a summary (sorted by value, largest first, optionally only the first N). */
export const summaryFormula = (block: DataBlock, s: SummarySpec): string => {
  const keys = dataCol(block, s.by);
  const vals = dataCol(block, s.val ?? s.by);
  const fn = AGG_LAMBDA[s.agg]!;
  const sort = s.agg === 'count' || s.val !== undefined ? (s.bottom ? 2 : -2) : 1;
  const g = `GROUPBY(${keys},${vals},LAMBDA(x,${fn}(x)),0,${s.top ? 0 : 1},${sort})`;
  return s.top ? `TAKE(${g},${s.top})` : g;
};

/** Adds a sheet with a summary table (a GROUPBY formula) and, optionally, a chart of it. One undo step. */
export const insertSummary = (store: GridStore, block: DataBlock, s: SummarySpec): void => {
  const valName = s.val !== undefined ? block.headers[s.val]! : 'Rows';
  store.change('Insert summary', 'book', () => {
    const name = uniqueName(store, s.title);
    store.wb.sheets.push(newSheet(name));
    store.engine.sheetsChanged();
    const si = store.wb.sheets.length - 1;
    const bold = internStyle(store.wb, { font: { bold: true }, fill: { color: '#E1F3E9' } });
    store.engine.setValue(si, 0, 0, block.headers[s.by]!);
    store.engine.setValue(si, 0, 1, `${AGG_LABEL[s.agg]} of ${valName}`.replace('Count of Rows', 'Count'));
    store.wb.sheets[si]!.cells.get(0)!.s = bold;
    store.wb.sheets[si]!.cells.get(1)!.s = bold;
    store.engine.setFormula(si, 1, 0, summaryFormula(block, s));
    store.wb.sheets[si]!.cols.set(0, { width: Math.min(40, Math.max(14, block.headers[s.by]!.length + 4)) });
    store.wb.sheets[si]!.cols.set(1, { width: 18 });
    store.engine.recalc();
    if (s.chart) {
      const spill = store.engine.spillRange(si, 1, 0);
      // The chart leaves out the grand total row that GROUPBY adds at the bottom.
      const last = spill ? spill.r2 : 1;
      const r2 = Math.max(1, s.top ? last : last - 1);
      const series: ChartSeries[] = [{ name: refString(name, { r1: 0, c1: 1, r2: 0, c2: 1 }), values: refString(name, { r1: 1, c1: 1, r2, c2: 1 }), categories: refString(name, { r1: 1, c1: 0, r2, c2: 0 }) }];
      store.wb.sheets[si]!.charts.push({ id: newId(), type: s.chart, title: s.title, series, anchor: { row: 0, col: 3, dx: 8, dy: 4, width: 480, height: 288 }, legend: s.chart === 'pie' || s.chart === 'doughnut' ? 'right' : 'none' });
    }
    store.wb.activeSheet = si;
  });
  store.selectRange({ r1: 1, c1: 0, r2: 1, c2: 0 });
};

/** Puts a chart of source columns beside the data on its own sheet. One undo step. */
export const insertSourceChart = (store: GridStore, block: DataBlock, type: ChartType, x: number, ys: number[], title: string): void => {
  const series: ChartSeries[] = ys.map((y) => ({ name: headCell(block, y) ?? block.headers[y], values: dataCol(block, y), categories: dataCol(block, x) }));
  const chart: Chart = {
    id: newId(), type, title, series, legend: ys.length > 1 ? 'bottom' : 'none',
    anchor: { row: block.range.r1, col: block.range.c2 + 2, dx: 0, dy: 0, width: 480, height: 288 },
    axisTitles: type === 'scatter' ? { x: block.headers[x], y: block.headers[ys[0]!] } : undefined,
  };
  if (store.sheetIndex !== block.sheet) store.setActiveSheet(block.sheet);
  store.change('Insert chart', 'sheet', () => { store.sheet.charts.push(chart); });
};

/** Result of an "ask" question: a single number to show, or something inserted. */
export const runAsk = (store: GridStore, block: DataBlock, ask: Ask): string => {
  const h = block.headers;
  switch (ask.kind) {
    case 'group': {
      const title = ask.agg === 'count' ? `Count by ${h[ask.by]}` : `${AGG_LABEL[ask.agg]} ${h[ask.val!]} by ${h[ask.by]}`;
      insertSummary(store, block, { by: ask.by, val: ask.agg === 'count' ? undefined : ask.val, agg: ask.agg, chart: 'column', title });
      return `Added a summary sheet: ${title}.`;
    }
    case 'top': {
      const title = `${ask.bottom ? 'Bottom' : 'Top'} ${ask.n} ${h[ask.by]} by ${h[ask.val]}`;
      insertSummary(store, block, { by: ask.by, val: ask.val, agg: 'sum', top: ask.n, bottom: ask.bottom, chart: 'bar', title });
      return `Added a summary sheet: ${title}.`;
    }
    case 'trend': {
      const title = `${h[ask.val]} over ${h[ask.x]}`;
      insertSourceChart(store, block, 'line', ask.x, [ask.val], title);
      return `Added a line chart: ${title}.`;
    }
    case 'total': {
      const nums = block.rows.map((r) => r[ask.val]).filter((v): v is number => typeof v === 'number');
      const filled = block.rows.filter((r) => r[ask.val] !== null && r[ask.val] !== '').length;
      const v = ask.agg === 'count' ? filled : !nums.length ? 0
        : ask.agg === 'sum' ? nums.reduce((a, b) => a + b, 0)
          : ask.agg === 'average' ? nums.reduce((a, b) => a + b, 0) / nums.length
            : ask.agg === 'min' ? Math.min(...nums) : Math.max(...nums);
      return `${AGG_LABEL[ask.agg]} of ${h[ask.val]}: ${+v.toPrecision(12)}`;
    }
  }
};
