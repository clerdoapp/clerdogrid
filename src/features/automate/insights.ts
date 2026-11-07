import type { Scalar } from '../../engine/types';
import { CellError } from '../../engine/types';

export type ColType = 'number' | 'date' | 'text' | 'boolean' | 'empty';

export interface ColumnSummary {
  index: number;
  name: string;
  type: ColType;
  count: number;
  blanks: number;
  distinct: number;
  min?: number; max?: number; mean?: number; median?: number; sum?: number;
  q1?: number; q3?: number;
  top?: { value: string; count: number }[];
}

export interface Outlier { col: number; name: string; row: number; value: number; low: number; high: number }
export interface Trend { x: number; y: number; xName: string; yName: string; slope: number; intercept: number; r2: number; n: number }
export interface Correlation { a: number; b: number; aName: string; bName: string; r: number; n: number }
export interface TopCategory { cat: number; val: number; catName: string; valName: string; items: { key: string; total: number }[] }
export interface Visual { kind: 'column' | 'line' | 'pie' | 'scatter' | 'bar'; title: string; cat?: number; val: number; val2?: number; agg?: 'sum' | 'average' | 'count' }

export interface Insights {
  rows: number;
  columns: ColumnSummary[];
  outliers: Outlier[];
  trend?: Trend;
  correlations: Correlation[];
  topCategories: TopCategory[];
  visuals: Visual[];
}

const median = (sorted: number[]): number => {
  const n = sorted.length;
  if (!n) return NaN;
  const m = Math.floor(n / 2);
  return n % 2 ? sorted[m]! : (sorted[m - 1]! + sorted[m]!) / 2;
};

/** Quartile by linear interpolation between closest ranks (inclusive method). */
export const quantile = (sorted: number[], p: number): number => {
  if (!sorted.length) return NaN;
  const pos = (sorted.length - 1) * p;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo]! + (sorted[hi]! - sorted[lo]!) * (pos - lo);
};

const isBlank = (v: Scalar): boolean => v === null || v === '';
const textOf = (v: Scalar): string => (v instanceof CellError ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v ?? ''));

/** Summaries for each column. `isDate(col)` tells whether a numeric column is shown as dates. */
export const summarizeColumns = (headers: string[], rows: Scalar[][], isDate: (col: number) => boolean = () => false): ColumnSummary[] =>
  headers.map((name, index) => {
    const vals = rows.map((r) => r[index] ?? null);
    const filled = vals.filter((v) => !isBlank(v));
    const nums = filled.filter((v): v is number => typeof v === 'number');
    const texts = filled.filter((v) => typeof v === 'string');
    const bools = filled.filter((v) => typeof v === 'boolean');
    let type: ColType = 'empty';
    if (filled.length) {
      if (nums.length >= filled.length * 0.8) type = isDate(index) ? 'date' : 'number';
      else if (bools.length >= filled.length * 0.8) type = 'boolean';
      else type = 'text';
    }
    const distinct = new Set(filled.map((v) => textOf(v).toLowerCase())).size;
    const s: ColumnSummary = { index, name, type, count: filled.length, blanks: vals.length - filled.length, distinct };
    if ((type === 'number' || type === 'date') && nums.length) {
      const sorted = [...nums].sort((a, b) => a - b);
      s.min = sorted[0]; s.max = sorted.at(-1);
      s.sum = nums.reduce((a, b) => a + b, 0);
      s.mean = s.sum / nums.length;
      s.median = median(sorted);
      s.q1 = quantile(sorted, 0.25); s.q3 = quantile(sorted, 0.75);
    }
    if (type === 'text' || type === 'boolean') {
      const counts = new Map<string, { value: string; count: number }>();
      for (const v of [...texts, ...bools]) {
        const t = textOf(v), k = t.toLowerCase();
        const e = counts.get(k);
        if (e) e.count++; else counts.set(k, { value: t, count: 1 });
      }
      s.top = [...counts.values()].sort((a, b) => b.count - a.count || a.value.localeCompare(b.value)).slice(0, 5);
    }
    return s;
  });

/** Values beyond 1.5 × the interquartile range from the quartiles. Needs at least 4 numbers. */
export const findOutliers = (cols: ColumnSummary[], rows: Scalar[][]): Outlier[] => {
  const out: Outlier[] = [];
  for (const c of cols) {
    if (c.type !== 'number' || c.q1 === undefined || c.q3 === undefined || c.count < 4) continue;
    const iqr = c.q3 - c.q1;
    const low = c.q1 - 1.5 * iqr, high = c.q3 + 1.5 * iqr;
    rows.forEach((r, i) => {
      const v = r[c.index];
      if (typeof v === 'number' && (v < low || v > high)) out.push({ col: c.index, name: c.name, row: i, value: v, low, high });
    });
  }
  return out;
};

export const linearFit = (xs: number[], ys: number[]): { slope: number; intercept: number; r2: number } => {
  const n = xs.length;
  const mx = xs.reduce((a, b) => a + b, 0) / n, my = ys.reduce((a, b) => a + b, 0) / n;
  let sxy = 0, sxx = 0, syy = 0;
  for (let i = 0; i < n; i++) { const dx = xs[i]! - mx, dy = ys[i]! - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
  const slope = sxx ? sxy / sxx : 0;
  const r2 = sxx && syy ? (sxy * sxy) / (sxx * syy) : 0;
  return { slope, intercept: my - slope * mx, r2 };
};

export const pearson = (xs: number[], ys: number[]): number => {
  const { r2, slope } = linearFit(xs, ys);
  return Math.sign(slope) * Math.sqrt(r2);
};

const pairs = (rows: Scalar[][], a: number, b: number): [number[], number[]] => {
  const xs: number[] = [], ys: number[] = [];
  for (const r of rows) { const x = r[a], y = r[b]; if (typeof x === 'number' && typeof y === 'number') { xs.push(x); ys.push(y); } }
  return [xs, ys];
};

/** Strictly increasing numbers with a constant step, like 1, 2, 3 or 2019, 2020, 2021. */
const isSequence = (rows: Scalar[][], col: number): boolean => {
  const v = rows.map((r) => r[col]).filter((x): x is number => typeof x === 'number');
  if (v.length < 3 || v.length < rows.length * 0.8) return false;
  const step = v[1]! - v[0]!;
  return step > 0 && v.every((x, i) => i === 0 || Math.abs(x - v[i - 1]! - step) < 1e-9 * Math.max(1, Math.abs(step)));
};

const TIME_WORDS = /date|time|day|week|month|quarter|year|period|step|^#$|^no\.?$|^n$|index/i;

/** Columns that order the rows in time or sequence: dates, else a time-like name, else the first even sequence. */
export const sequenceColumns = (cols: ColumnSummary[], rows: Scalar[][]): number[] => {
  const dates = cols.filter((c) => c.type === 'date').map((c) => c.index);
  if (dates.length) return dates;
  const seq = cols.filter((c) => c.type === 'number' && isSequence(rows, c.index));
  const named = seq.filter((c) => TIME_WORDS.test(c.name.trim()));
  if (named.length) return named.map((c) => c.index);
  return seq.slice(0, 1).map((c) => c.index);
};

/** The numeric column that changes most consistently over a date or sequence column (highest R²). */
export const findTrend = (cols: ColumnSummary[], rows: Scalar[][]): Trend | undefined => {
  const xs = sequenceColumns(cols, rows);
  let best: Trend | undefined;
  for (const xi of xs) for (const y of cols) {
    if (xs.includes(y.index) || y.type !== 'number') continue;
    const x = cols[xi]!;
    const [px, py] = pairs(rows, x.index, y.index);
    if (px.length < 3) continue;
    const f = linearFit(px, py);
    if (!best || f.r2 > best.r2) best = { x: x.index, y: y.index, xName: x.name, yName: y.name, ...f, n: px.length };
  }
  return best;
};

export const findCorrelations = (cols: ColumnSummary[], rows: Scalar[][], threshold = 0.6, skip: number[] = []): Correlation[] => {
  const nums = cols.filter((c) => c.type === 'number' && !skip.includes(c.index));
  const out: Correlation[] = [];
  for (let i = 0; i < nums.length; i++) for (let j = i + 1; j < nums.length; j++) {
    const a = nums[i]!, b = nums[j]!;
    const [xs, ys] = pairs(rows, a.index, b.index);
    if (xs.length < 3) continue;
    const r = pearson(xs, ys);
    if (Math.abs(r) > threshold) out.push({ a: a.index, b: b.index, aName: a.name, bName: b.name, r, n: xs.length });
  }
  return out.sort((x, y) => Math.abs(y.r) - Math.abs(x.r));
};

export const groupTotals = (rows: Scalar[][], cat: number, val: number, agg: 'sum' | 'average' | 'count' = 'sum'): { key: string; total: number }[] => {
  const m = new Map<string, { key: string; sum: number; n: number }>();
  for (const r of rows) {
    const k = r[cat];
    if (isBlank(k ?? null)) continue;
    const t = textOf(k ?? null);
    const e = m.get(t.toLowerCase()) ?? { key: t, sum: 0, n: 0 };
    const v = r[val];
    if (agg === 'count') e.n++;
    else if (typeof v === 'number') { e.sum += v; e.n++; }
    m.set(t.toLowerCase(), e);
  }
  return [...m.values()].map((e) => ({ key: e.key, total: agg === 'count' ? e.n : agg === 'average' ? (e.n ? e.sum / e.n : 0) : e.sum }))
    .sort((a, b) => b.total - a.total);
};

/** Text columns with a manageable number of categories, each with the numeric column that varies most between them. */
export const findTopCategories = (cols: ColumnSummary[], rows: Scalar[][], skip: number[] = []): TopCategory[] => {
  const cats = cols.filter((c) => c.type === 'text' && c.distinct >= 2 && c.distinct <= Math.max(12, rows.length / 2) && c.distinct < rows.length);
  const vals = cols.filter((c) => c.type === 'number' && !skip.includes(c.index));
  const out: TopCategory[] = [];
  for (const c of cats) {
    const v = vals[0];
    if (!v) continue;
    out.push({ cat: c.index, val: v.index, catName: c.name, valName: v.name, items: groupTotals(rows, c.index, v.index).slice(0, 5) });
  }
  return out.slice(0, 3);
};

export const suggestVisuals = (cols: ColumnSummary[], trend: Trend | undefined, corr: Correlation[], tops: TopCategory[]): Visual[] => {
  const out: Visual[] = [];
  if (trend) out.push({ kind: 'line', title: `${trend.yName} over ${trend.xName}`, cat: trend.x, val: trend.y });
  for (const t of tops.slice(0, 2)) {
    const few = (cols[t.cat]?.distinct ?? 99) <= 6;
    out.push({ kind: few ? 'pie' : 'column', title: `${t.valName} by ${t.catName}`, cat: t.cat, val: t.val, agg: 'sum' });
  }
  if (corr[0]) out.push({ kind: 'scatter', title: `${corr[0].bName} against ${corr[0].aName}`, val: corr[0].a, val2: corr[0].b });
  return out;
};

export const analyze = (headers: string[], rows: Scalar[][], isDate?: (col: number) => boolean): Insights => {
  const columns = summarizeColumns(headers, rows, isDate);
  const outliers = findOutliers(columns, rows);
  const trend = findTrend(columns, rows);
  const seq = sequenceColumns(columns, rows);
  const correlations = findCorrelations(columns, rows, 0.6, seq);
  const topCategories = findTopCategories(columns, rows, seq);
  return { rows: rows.length, columns, outliers, trend, correlations, topCategories, visuals: suggestVisuals(columns, trend, correlations, topCategories) };
};

// ---------- "Ask about your data" ----------

export type Ask =
  | { kind: 'group'; agg: 'sum' | 'average' | 'count' | 'min' | 'max'; val?: number; by: number }
  | { kind: 'top'; n: number; by: number; val: number; bottom?: boolean }
  | { kind: 'trend'; val: number; x: number }
  | { kind: 'total'; agg: 'sum' | 'average' | 'count' | 'min' | 'max'; val: number };

const AGG_WORDS: [RegExp, 'sum' | 'average' | 'count' | 'min' | 'max'][] = [
  [/^(?:total|sum(?: of)?|sum)$/, 'sum'], [/^(?:average|avg|mean)$/, 'average'], [/^(?:count(?: of)?|number of|how many)$/, 'count'],
  [/^(?:min(?:imum)?|lowest|smallest)$/, 'min'], [/^(?:max(?:imum)?|highest|largest|biggest)$/, 'max'],
];

const norm = (s: string): string => s.toLowerCase().replace(/[_\-]+/g, ' ').replace(/\s+/g, ' ').trim();

/** Finds a column by name: exact, then singular/plural, then contains. */
export const matchColumn = (headers: string[], phrase: string): number => {
  const p = norm(phrase).replace(/^(?:the|each|every|per)\s+/, '');
  if (!p) return -1;
  const hs = headers.map(norm);
  const singular = (x: string) => x.replace(/(?:ies)$/, 'y').replace(/(?<=[^s])s$/, '');
  let i = hs.indexOf(p);
  if (i < 0) i = hs.findIndex((h) => singular(h) === singular(p));
  if (i < 0) i = hs.findIndex((h) => h.includes(p) || (h.length > 2 && p.includes(h)));
  return i;
};

/** Reads simple English questions. Returns null with nothing understood. Pure pattern matching, no network. */
export const parseAsk = (text: string, headers: string[], cols?: ColumnSummary[]): Ask | null => {
  const q = norm(text).replace(/[?.!]+$/, '').replace(/^(?:show(?: me)?|what is|what's|what are|give me|list|find)\s+/, '').trim();
  let m: RegExpExecArray | null;

  m = /^(top|bottom)\s+(\d{1,3})\s+(.+?)\s+by\s+(.+)$/.exec(q);
  if (m) {
    const by = matchColumn(headers, m[3]!), val = matchColumn(headers, m[4]!);
    if (by >= 0 && val >= 0) return { kind: 'top', n: Math.max(1, Math.min(100, +m[2]!)), by, val, bottom: m[1] === 'bottom' };
  }
  m = /^(?:count|number of rows|how many rows|rows)\s+(?:by|per|for each)\s+(.+)$/.exec(q);
  if (m) { const by = matchColumn(headers, m[1]!); if (by >= 0) return { kind: 'group', agg: 'count', by }; }

  m = /^(.+?)\s+(?:over time|by date|over (?:the )?(.+))$/.exec(q);
  if (m && /over|by date/.test(q)) {
    const val = matchColumn(headers, m[1]!.replace(/^(?:total|average|sum of)\s+/, ''));
    let x = m[2] ? matchColumn(headers, m[2]) : -1;
    if (x < 0 && cols) x = cols.find((c) => c.type === 'date')?.index ?? -1;
    if (x < 0) x = headers.findIndex((h) => /date|time|month|year|day|week|period/i.test(h));
    if (val >= 0 && x >= 0 && val !== x) return { kind: 'trend', val, x };
  }

  m = /^(?:(total|sum(?: of)?|average|avg|mean|count(?: of)?|number of|min(?:imum)?|max(?:imum)?|lowest|highest|largest|smallest)\s+)?(.+?)\s+(?:by|per|for each|grouped by|across)\s+(.+)$/.exec(q);
  if (m) {
    const aggWord = m[1] ?? 'total';
    const agg = AGG_WORDS.find(([re]) => re.test(aggWord))?.[1] ?? 'sum';
    const val = matchColumn(headers, m[2]!), by = matchColumn(headers, m[3]!);
    if (by >= 0 && agg === 'count') return { kind: 'group', agg, by, val: val >= 0 ? val : undefined };
    if (val >= 0 && by >= 0 && val !== by) return { kind: 'group', agg, val, by };
  }

  m = /^(total|sum(?: of)?|average|avg|mean|count(?: of)?|min(?:imum)?|max(?:imum)?|lowest|highest)\s+(.+)$/.exec(q);
  if (m && !/\s(?:by|per|over|for each)\s/.test(` ${m[2]} `)) {
    const agg = AGG_WORDS.find(([re]) => re.test(m![1]!))?.[1] ?? 'sum';
    const val = matchColumn(headers, m[2]!);
    if (val >= 0) return { kind: 'total', agg, val };
  }
  return null;
};
