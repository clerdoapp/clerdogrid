import { quoteSheet } from '../../engine/parser';
import { isMatrix, type Scalar } from '../../engine/types';
import { formatValue } from '../../formats/numfmt';
import { colIndex, colName, type CellRange, type Chart, type ChartSeries, type ChartType } from '../../model/workbook';
import { displayText } from '../../app/ops/clipboard';
import type { GridStore } from '../../app/store';
import { comboKind, traits, type SeriesKind } from './chartTypes';
import { parseNumberList } from './trend';
import { N } from '../i18n';

// ---------- Axis ticks ----------

export interface Scale { min: number; max: number; step: number; ticks: number[] }

const clean = (v: number): number => +v.toPrecision(12);

/** A "nice" number near x: 1, 2 or 5 times a power of ten. */
export const niceNum = (x: number, round: boolean): number => {
  if (x <= 0 || !Number.isFinite(x)) return 1;
  const exp = Math.floor(Math.log10(x));
  const f = x / 10 ** exp;
  const nf = round ? (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) : f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
  return nf * 10 ** exp;
};

/** Round axis bounds and evenly spaced ticks that cover [lo, hi]. */
export const niceScale = (lo: number, hi: number, maxTicks = 6): Scale => {
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) { lo = 0; hi = 1; }
  if (lo > hi) [lo, hi] = [hi, lo];
  if (lo === hi) {
    if (lo === 0) hi = 1;
    else { const d = Math.abs(lo) * 0.1; lo -= d; hi += d; }
  }
  const n = Math.max(2, maxTicks);
  const step = niceNum((hi - lo) / (n - 1), true);
  const min = clean(Math.floor(lo / step + 1e-9) * step);
  const max = clean(Math.ceil(hi / step - 1e-9) * step);
  const ticks: number[] = [];
  for (let i = 0; min + i * step <= max + step / 2 && i < 100; i++) ticks.push(clean(min + i * step));
  return { min, max: Math.max(max, ticks.at(-1) ?? max), step, ticks };
};

/** Value-axis range for a chart: bars and areas always start at zero; lines do when zero is close. */
export const valueExtent = (values: number[], includeZero: boolean): [number, number] => {
  if (!values.length) return [0, 1];
  let lo = Math.min(...values), hi = Math.max(...values);
  if (includeZero || (lo > 0 && lo < hi * 0.6) || (hi < 0 && hi > lo * 0.6)) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  return [lo, hi];
};

// ---------- References ----------

export interface ParsedRef { sheet?: string; range: CellRange }

const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^!'\s:]+))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/;

/** Parses "Sheet1!$B$2:$B$10", "'My data'!A1" or "B2:B9". */
export const parseRef = (text: string): ParsedRef | null => {
  const m = REF_RE.exec(text.trim().replace(/^=/, ''));
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const c1 = colIndex(m[3]!), r1 = +m[4]! - 1;
  const c2 = m[5] ? colIndex(m[5]) : c1, r2 = m[6] ? +m[6] - 1 : r1;
  if (r1 < 0 || r2 < 0) return null;
  return { sheet, range: { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) } };
};

const abs = (row: number, col: number) => `$${colName(col)}$${row + 1}`;
export const refString = (sheet: string, r: CellRange): string =>
  `${quoteSheet(sheet)}!${abs(r.r1, r.c1)}${r.r1 === r.r2 && r.c1 === r.c2 ? '' : `:${abs(r.r2, r.c2)}`}`;

// ---------- Building series from a block of cells ----------

export type Orientation = 'cols' | 'rows';
export interface Headers { firstRow: boolean; firstCol: boolean }

const isText = (v: Scalar): boolean => typeof v === 'string' && v !== '';

/** Whether the block has a header row (series names) and a first column of labels (categories). */
export const detectHeaders = (m: Scalar[][]): Headers => {
  const rows = m.length, cols = m[0]?.length ?? 0;
  if (!rows || !cols) return { firstRow: false, firstCol: false };
  const corner = m[0]![0] ?? null;
  const colRest = m.slice(1).map((r) => r[0] ?? null);
  const rowRest = m[0]!.slice(1);
  const firstCol = cols > 1 && (colRest.some(isText) || (corner === null && rows > 1 && colRest.every((v) => v !== null)));
  const firstRow = rows > 1 && (rowRest.length ? rowRest.every((v) => v === null || isText(v)) && rowRest.some(isText) : isText(corner))
    && (cols === 1 || !firstCol || corner === null || isText(corner));
  return { firstRow, firstCol };
};

export const defaultOrientation = (m: Scalar[][], h: Headers): Orientation => {
  const dataRows = m.length - (h.firstRow ? 1 : 0);
  const dataCols = (m[0]?.length ?? 0) - (h.firstCol ? 1 : 0);
  return dataRows >= dataCols ? 'cols' : 'rows';
};

/**
 * Series formulas for a block of cells. `m` holds the block's values (used only for header detection).
 * By columns, each data column is a series; by rows, each data row is.
 */
const labelColumns = (m: Scalar[][], h: Headers): number => {
  const body = m.slice(h.firstRow ? 1 : 0);
  const cols = m[0]?.length ?? 0;
  let n = 0;
  while (n < cols - 1 && body.every((r) => typeof r[n] !== 'number')) n++;
  return Math.max(n, h.firstCol ? 1 : 0);
};

export const buildSeries = (sheet: string, r: CellRange, m: Scalar[][], opts: { orientation?: Orientation; headers?: Headers; type?: ChartType } = {}): ChartSeries[] => {
  let h = opts.headers ?? detectHeaders(m);
  const type = opts.type ?? 'column';
  const tr0 = traits(type);
  const fam = tr0.family;
  const hierarchy = fam === 'treemap' || fam === 'sunburst';
  const xy = fam === 'scatter' || fam === 'bubble';
  // Stock data starts with a column of dates when there's one more column than prices.
  const stockCols = 3 + (tr0.open ? 1 : 0) + (tr0.volume ? 1 : 0);
  if (fam === 'stock' && !h.firstCol && (m[0]?.length ?? 0) > stockCols) h = { ...h, firstCol: true };
  // Hierarchy charts read every leading text column as a level; their data always runs down columns.
  const orient = hierarchy ? 'cols' : opts.orientation ?? (xy || fam === 'stock' ? 'cols' : defaultOrientation(m, h));
  const labelCols = hierarchy ? labelColumns(m, h) : h.firstCol ? 1 : 0;
  const top = r.r1 + (h.firstRow ? 1 : 0), left = r.c1 + (orient === 'cols' ? labelCols : h.firstCol ? 1 : 0);
  if (top > r.r2 || left > r.c2) return [];
  const out: ChartSeries[] = [];
  if (orient === 'cols') {
    let first = left;
    let categories = labelCols ? refString(sheet, { r1: top, c1: r.c1, r2: r.r2, c2: r.c1 + labelCols - 1 }) : undefined;
    // An x-y chart without a label column uses its first column as the x values.
    if (xy && !labelCols && r.c2 > left) { categories = refString(sheet, { r1: top, c1: left, r2: r.r2, c2: left }); first = left + 1; }
    const step = fam === 'bubble' ? 2 : 1;
    for (let c = first; c <= r.c2; c += step) {
      const s: ChartSeries = {
        name: h.firstRow ? refString(sheet, { r1: r.r1, c1: c, r2: r.r1, c2: c }) : `Series ${out.length + 1}`,
        values: refString(sheet, { r1: top, c1: c, r2: r.r2, c2: c }),
        ...(categories ? { categories } : {}),
      };
      if (fam === 'bubble' && c + 1 <= r.c2) s.sizes = refString(sheet, { r1: top, c1: c + 1, r2: r.r2, c2: c + 1 });
      out.push(s);
    }
  } else {
    let first = top;
    let categories = h.firstRow ? refString(sheet, { r1: r.r1, c1: left, r2: r.r1, c2: r.c2 }) : undefined;
    if (xy && !h.firstRow && r.r2 > top) { categories = refString(sheet, { r1: top, c1: left, r2: top, c2: r.c2 }); first = top + 1; }
    const step = fam === 'bubble' ? 2 : 1;
    for (let row = first; row <= r.r2; row += step) {
      const s: ChartSeries = {
        name: h.firstCol ? refString(sheet, { r1: row, c1: r.c1, r2: row, c2: r.c1 }) : `Series ${out.length + 1}`,
        values: refString(sheet, { r1: row, c1: left, r2: row, c2: r.c2 }),
        ...(categories ? { categories } : {}),
      };
      if (fam === 'bubble' && row + 1 <= r.r2) s.sizes = refString(sheet, { r1: row + 1, c1: left, r2: row + 1, c2: r.c2 });
      out.push(s);
    }
  }
  if (traits(type).combo) out.forEach((s, i) => { s.type = comboKind(type, i, out.length); });
  return out;
};

/** The block that a chart's series were built from (bounding box of its references on one sheet). */
export const chartSourceRange = (chart: Chart): { sheet?: string; range: CellRange } | null => {
  let box: CellRange | null = null;
  let sheet: string | undefined;
  for (const s of chart.series) for (const f of [s.values, s.categories, s.name]) {
    const p = f ? parseRef(f) : null;
    if (!p) continue;
    sheet ??= p.sheet;
    box = box ? { r1: Math.min(box.r1, p.range.r1), c1: Math.min(box.c1, p.range.c1), r2: Math.max(box.r2, p.range.r2), c2: Math.max(box.c2, p.range.c2) } : { ...p.range };
  }
  return box ? { sheet, range: box } : null;
};

/** Series run down columns unless each series' values lie along a row. */
export const chartOrientation = (chart: Chart): Orientation => {
  const p = chart.series[0] ? parseRef(chart.series[0].values) : null;
  return p && p.range.r1 === p.range.r2 && p.range.c2 > p.range.c1 ? 'rows' : 'cols';
};

// ---------- Recommendations ----------

const DATEISH = /^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|q[1-4]|mon|tue|wed|thu|fri|sat|sun)/i;

/** Chart types that suit a block, best first. */
export const recommendTypes = (m: Scalar[][], h: Headers = detectHeaders(m)): ChartType[] => {
  const o = defaultOrientation(m, h);
  const body = m.slice(h.firstRow ? 1 : 0).map((r) => r.slice(h.firstCol ? 1 : 0));
  const nums = body.flat().filter((v): v is number => typeof v === 'number');
  const seriesCount = o === 'cols' ? body[0]?.length ?? 0 : body.length;
  const points = o === 'cols' ? body.length : body[0]?.length ?? 0;
  const labels = h.firstCol ? m.slice(h.firstRow ? 1 : 0).map((r) => r[0]) : h.firstRow ? m[0]!.slice(h.firstCol ? 1 : 0) : [];
  const timeLike = labels.length > 2 && labels.every((v) => (typeof v === 'number' && v > 1900 && v < 2200) || (typeof v === 'string' && DATEISH.test(v)));
  const longLabels = labels.some((v) => typeof v === 'string' && v.length > 14);
  const out: ChartType[] = [];
  const push = (...t: ChartType[]) => { for (const x of t) if (!out.includes(x)) out.push(x); };
  // Column names that say what the data is: prices, or places.
  const header = h.firstRow ? m[0]!.map((v) => (typeof v === 'string' ? v.toLowerCase() : '')) : [];
  const named = (w: string) => header.some((x) => x.includes(w));
  if (named('high') && named('low') && named('close')) push(named('volume') ? (named('open') ? 'stockVOHLC' : 'stockVHLC') : named('open') ? 'stockOHLC' : 'stockHLC');
  if (h.firstCol && seriesCount === 1 && /countr|state|nation|region|province|territor/.test(header[0] ?? '')) push('regionMap');
  if (!h.firstCol && !h.firstRow && seriesCount >= 2 && o === 'cols' && nums.length === body.flat().filter((v) => v !== null).length) push('scatter');
  if (timeLike || points > 12) push('line', 'column', 'area');
  if (seriesCount === 1 && points <= 8 && nums.length && nums.every((n) => n >= 0)) push(longLabels ? 'bar' : 'column', 'pie', 'doughnut');
  if (seriesCount === 1 && points >= 12 && !h.firstCol) push('histogram', 'boxWhisker');
  if (seriesCount === 1 && points >= 3 && points <= 15 && h.firstCol && nums.length && nums.every((n) => n >= 0)) push('pareto', 'treemap');
  if (seriesCount > 1) push('column', 'stackedColumn', 'line', 'bar', 'combo');
  if (longLabels) push('bar');
  push('column', 'line', 'bar', 'pie');
  return out.slice(0, 5);
};

// ---------- Resolving live data ----------

export interface ResolvedSeries {
  name: string; values: (number | null)[]; color: string; type: SeriesKind;
  /** Bubble sizes. */
  sizes: (number | null)[] | null;
  /** Measured on the second value axis. */
  secondary: boolean;
  /** Position in chart.series. */
  index: number;
  src: ChartSeries;
  /** This series' own x values (scatter and bubble). */
  xs: (number | null)[] | null;
  /** Label text per point from the label range, when set. */
  cellLabels?: string[];
  /** Custom error amounts per error-bar set. */
  errors?: { plus?: (number | null)[]; minus?: (number | null)[] }[];
}
export interface ResolvedChart {
  categories: string[];
  /** Numeric x values for scatter and bubble charts. */
  xs: (number | null)[] | null;
  /** Category levels per point (treemap and sunburst). */
  paths: string[][];
  series: ResolvedSeries[];
  /** Number format for axis and label text: the chart's own, else that of the first value cell. */
  fmt: string;
  /** Category cell values (numbers or null) and their format, for date axes. */
  catValues: (number | null)[];
  catFmt: string;
  /** Title texts with cell links resolved. */
  title?: string;
  axisTitles: { x?: string; y?: string; y2?: string };
}

export const PALETTE_SIZE = 10;
export const STYLE_COUNT = 5;
export const STYLE_NAMES = [N('Mixed'), N('Greens'), N('Forest and sun'), N('Mint and slate'), N('Soft')];
const STYLE_SIZE = 8;
/** Colour of series (or slice) `i` in colour set `style`; a series' own colour wins. */
export const seriesColor = (i: number, custom?: string, style = 0): string => {
  if (custom) return custom;
  if (!style || style >= STYLE_COUNT) return `var(--chart-${(i % PALETTE_SIZE) + 1})`;
  return `var(--chart-s${style}-${(i % STYLE_SIZE) + 1})`;
};

/** Series indexes on the second axis, including the defaults for the two-axis combo and volume stock charts. */
export const secondaryOf = (chart: Chart): Set<number> => {
  if (chart.options?.secondary) return new Set(chart.options.secondary);
  if (chart.type === 'comboSecondary') return new Set(chart.series.flatMap((s, i) => ((s.type ?? comboKind(chart.type, i, chart.series.length)) === 'line' ? [i] : [])));
  return new Set();
};

const MAX_POINTS = 4000;

interface Cells { values: Scalar[]; texts: string[]; fmt: string; hidden: boolean[] }

const sheetOf = (store: GridStore, name: string | undefined, sheetIdx: number): number =>
  name === undefined ? sheetIdx : store.wb.sheets.findIndex((s) => s.name.toLowerCase() === name.toLowerCase());

const readFormula = (store: GridStore, formula: string, sheetIdx: number): Cells => {
  const p = parseRef(formula);
  if (p) {
    const si = sheetOf(store, p.sheet, sheetIdx);
    if (si < 0) return { values: [], texts: [], fmt: 'General', hidden: [] };
    const sh = store.wb.sheets[si]!;
    const r = p.range;
    const bounds = store.engine.usedBounds(si);
    const r2 = Math.min(r.r2, Math.max(r.r1, bounds.rows)), c2 = Math.min(r.c2, Math.max(r.c1, bounds.cols));
    const values: Scalar[] = [], texts: string[] = [], hidden: boolean[] = [];
    for (let row = r.r1; row <= r2; row++) for (let col = r.c1; col <= c2; col++) {
      if (values.length >= MAX_POINTS) break;
      values.push(store.value(row, col, si));
      texts.push(displayText(store, row, col, si));
      hidden.push(!!sh.rows.get(row)?.hidden || !!sh.cols.get(col)?.hidden);
    }
    return { values, texts, fmt: store.engine.numberFormat(si, r.r1, r.c1), hidden };
  }
  const list = parseNumberList(formula);
  if (list) return { values: list, texts: list.map(String), fmt: 'General', hidden: list.map(() => false) };
  try {
    const v = store.engine.evaluate(formula.replace(/^=/, ''), sheetIdx, 0, 0);
    const values = (isMatrix(v) ? v.flat() : [v]).slice(0, MAX_POINTS);
    return { values, texts: values.map((x) => (x === null ? '' : String(typeof x === 'object' ? x.code : x))), fmt: 'General', hidden: values.map(() => false) };
  } catch {
    return { values: [], texts: [], fmt: 'General', hidden: [] };
  }
};

/** Text of a title: a cell link (=Sheet1!A1) shows the cell's text. */
export const resolveText = (store: GridStore, text: string | undefined, sheetIdx = store.sheetIndex): string | undefined => {
  if (!text) return text;
  if (!/^=/.test(text.trim())) return text;
  const c = readFormula(store, text.trim(), sheetIdx);
  return c.texts.filter(Boolean).join(' ');
};

/** Reads a range or list of numbers (custom error amounts, sparkline data). */
export const readNumbers = (store: GridStore, formula: string, sheetIdx = store.sheetIndex): (number | null)[] => readFormula(store, formula, sheetIdx).values.map(num);
export const readTexts = (store: GridStore, formula: string, sheetIdx = store.sheetIndex): string[] => readFormula(store, formula, sheetIdx).texts;

const looksLikeRef = (s: string): boolean => /^=/.test(s) || parseRef(s) !== null;
const num = (x: Scalar): number | null => (typeof x === 'number' && Number.isFinite(x) ? x : null);

export const resolveChart = (store: GridStore, chart: Chart, sheetIdx = store.sheetIndex): ResolvedChart => {
  let catCells: Cells | null = null;
  let fmt = 'General';
  const tr = traits(chart.type);
  const o = chart.options ?? {};
  const style = o.style ?? 0;
  const second = secondaryOf(chart);
  const xy = tr.family === 'scatter' || tr.family === 'bubble';
  const raw = chart.series.map((s, i) => {
    const v = readFormula(store, s.values, sheetIdx);
    if (i === 0) fmt = v.fmt;
    const cats = s.categories ? readFormula(store, s.categories, sheetIdx) : null;
    if (cats && !catCells) catCells = cats;
    return { s, i, v, cats };
  });
  // Points in hidden rows or columns are left out unless the chart says to show them.
  const hiddenMask = !o.showHidden && raw[0] ? raw[0].v.hidden : [];
  const keep = (k: number) => !hiddenMask[k];
  const filt = <T,>(a: T[]): T[] => (hiddenMask.some(Boolean) ? a.filter((_, k) => keep(k)) : a);
  const blankZero = o.blanks === 'zero';
  const series = raw.map(({ s, i, v, cats }): ResolvedSeries => {
    const name = s.name ? (looksLikeRef(s.name) ? readFormula(store, s.name, sheetIdx).texts.join(' ') : s.name) : '';
    const values = filt(v.values.map(num)).map((x) => (x === null && blankZero ? 0 : x));
    const out: ResolvedSeries = {
      name: name || `Series ${i + 1}`,
      values,
      // Stock volume defaults to a quiet colour so the price marks stand out.
      color: tr.family === 'stock' && tr.volume && i === 0 && chart.series.length > 1 && !s.color && !s.fill?.color ? 'var(--chart-volume, var(--chart-9))' : seriesColor(i, s.color ?? s.fill?.color, style),
      type: tr.combo ? s.type ?? comboKind(chart.type, i, chart.series.length) : tr.base,
      sizes: s.sizes ? filt(readFormula(store, s.sizes, sheetIdx).values.map(num)) : null,
      secondary: second.has(i),
      index: i,
      src: s,
      xs: xy && cats ? filt(cats.values.map(num)) : null,
    };
    const lbl = s.labels?.fromCells ?? (s.labels ? undefined : o.labels?.fromCells);
    if (lbl) out.cellLabels = filt(readFormula(store, lbl, sheetIdx).texts);
    if (s.errorBars?.some((e) => e.kind === 'custom')) {
      out.errors = s.errorBars.map((e) => (e.kind !== 'custom' ? {} : {
        plus: e.plus ? filt(readFormula(store, e.plus, sheetIdx).values.map(num)) : undefined,
        minus: e.minus ? filt(readFormula(store, e.minus, sheetIdx).values.map(num)) : undefined,
      }));
    }
    return out;
  });
  const cc = catCells as Cells | null;
  const catTexts = cc ? filt(cc.texts) : [];
  const catValues = cc ? filt(cc.values) : [];
  const n = Math.max(0, ...series.map((s) => s.values.length));
  // A category range may span several columns (hierarchy levels); split it into one path per point.
  const k = n && catTexts.length > n ? Math.max(1, Math.round(catTexts.length / n)) : 1;
  const paths: string[][] = [];
  for (let i = 0; i < n; i++) paths.push(k === 1 ? [catTexts[i] ?? ''] : catTexts.slice(i * k, i * k + k));
  // Blank parent cells repeat the label above them, as in an outline.
  if (k > 1) for (let i = 1; i < paths.length; i++) {
    const p = paths[i]!, prev = paths[i - 1]!;
    for (let d = 0; d < k - 1 && p[d] === '' && p.slice(d + 1).some((x) => x !== ''); d++) p[d] = prev[d] ?? '';
  }
  let categories = k === 1 ? catTexts.slice(0, n) : paths.map((p) => p.filter(Boolean).join(' / '));
  if (categories.length < n) categories = [...categories, ...Array.from({ length: n - categories.length }, (_, i) => String(categories.length + i + 1))];
  const xs = xy ? (catValues.length ? catValues.slice(0, Math.max(n, 1)).map(num) : null) : null;
  const at = chart.axisTitles ?? {};
  return {
    categories: categories.slice(0, Math.max(n, 1)), xs, paths, series, fmt: o.numFmt || o.axes?.y?.numFmt || fmt,
    catValues: k === 1 ? catValues.slice(0, n).map(num) : [], catFmt: cc?.fmt ?? 'General',
    title: resolveText(store, chart.title, sheetIdx),
    axisTitles: { x: resolveText(store, at.x, sheetIdx), y: resolveText(store, at.y, sheetIdx), y2: resolveText(store, at.y2, sheetIdx) },
  };
};

/** Short text for a number on an axis or label, using the source cells' format. */
export const formatNumber = (v: number, fmt: string, compact = false): string => {
  if ((fmt === 'General' || !fmt) && compact) {
    const a = Math.abs(v);
    const unit = a >= 1e12 ? ['T', 1e12] as const : a >= 1e9 ? ['B', 1e9] as const : a >= 1e6 ? ['M', 1e6] as const : a >= 1e4 ? ['K', 1e3] as const : null;
    if (unit) return `${clean(v / unit[1])}${unit[0]}`;
  }
  return formatValue(clean(v), fmt || 'General').text;
};
