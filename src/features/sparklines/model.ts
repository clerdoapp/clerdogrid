// Sparklines: small charts drawn inside cells. Pure helpers (no drawing) so they can be tested.
import type { GridStore } from '../../app/store';
import type { CellRange, Sheet, Sparkline, SparklineGroup, SparklineKind } from '../../model/workbook';
import { parseRef, refString } from '../../ui/charts/chartData';
import { N, T } from '../../ui/i18n';

/** Colour presets in the green palette: series, negative, markers, high, low, first, last. */
export const SPARK_STYLES: { name: string; color: string; negativeColor: string; markersColor: string; highColor: string; lowColor: string; firstColor: string; lastColor: string }[] = [
  { name: N('Green'), color: '#12A565', negativeColor: '#E03131', markersColor: '#0C7A4B', highColor: '#0A5C39', lowColor: '#E8A317', firstColor: '#2F6FDF', lastColor: '#7048E8' },
  { name: N('Deep green'), color: '#0C7A4B', negativeColor: '#C2410C', markersColor: '#0A5C39', highColor: '#06402A', lowColor: '#F08C00', firstColor: '#4CC38A', lastColor: '#4CC38A' },
  { name: N('Mint'), color: '#4CC38A', negativeColor: '#D6336C', markersColor: '#2E9C6E', highColor: '#0C7A4B', lowColor: '#D6336C', firstColor: '#8FD9B2', lastColor: '#0C7A4B' },
  { name: N('Forest and sun'), color: '#0A5C39', negativeColor: '#E8A317', markersColor: '#E8A317', highColor: '#12A565', lowColor: '#C27A0E', firstColor: '#12A565', lastColor: '#E8A317' },
  { name: N('Slate and green'), color: '#5C6F82', negativeColor: '#E03131', markersColor: '#12A565', highColor: '#12A565', lowColor: '#E03131', firstColor: '#34495E', lastColor: '#12A565' },
  { name: N('Soft'), color: '#8FD9B2', negativeColor: '#E8A0A8', markersColor: '#6CC49A', highColor: '#2BB380', lowColor: '#E7B97A', firstColor: '#9AB8D6', lastColor: '#C9A3D4' },
];

export const defaultGroup = (kind: SparklineKind, sparklines: Sparkline[]): SparklineGroup => {
  const st = SPARK_STYLES[0]!;
  return {
    kind, sparklines, color: st.color, negativeColor: st.negativeColor, markersColor: st.markersColor, highColor: st.highColor, lowColor: st.lowColor,
    firstColor: st.firstColor, lastColor: st.lastColor, axisColor: '#000000', lineWeight: 0.75, displayEmptyCellsAs: 'gap',
    ...(kind === 'winLoss' ? { negative: true } : {}),
  };
};

/** One sparkline per row or column of `data`, placed along `location`. Returns an error message when they don't line up. */
export const buildSparklines = (sheetName: string, data: { sheet?: string; range: CellRange }, location: CellRange): Sparkline[] | string => {
  const locRows = location.r2 - location.r1 + 1, locCols = location.c2 - location.c1 + 1;
  if (locRows > 1 && locCols > 1) return T('The location must be one row or one column of cells.');
  const n = Math.max(locRows, locCols);
  const r = data.range;
  const dRows = r.r2 - r.r1 + 1, dCols = r.c2 - r.c1 + 1;
  const src = data.sheet ?? sheetName;
  const at = (k: number) => (locRows > 1 ? { row: location.r1 + k, col: location.c1 } : { row: location.r1, col: location.c1 + k });
  if (n === 1) {
    if (dRows > 1 && dCols > 1) return T('For one sparkline, choose one row or one column of data.');
    return [{ ref: refString(src, r), ...at(0) }];
  }
  // Data rows go to a column of cells, data columns to a row of cells; a square block follows the location.
  const byRows = locRows > 1 ? dRows === n : dRows === n && dCols !== n;
  if (byRows) return Array.from({ length: n }, (_, k) => ({ ref: refString(src, { r1: r.r1 + k, c1: r.c1, r2: r.r1 + k, c2: r.c2 }), ...at(k) }));
  if (dCols === n) return Array.from({ length: n }, (_, k) => ({ ref: refString(src, { r1: r.r1, c1: r.c1 + k, r2: r.r2, c2: r.c1 + k }), ...at(k) }));
  return T("The data has {rows} rows and {cols} columns, so it doesn't match {count} location cells.", { rows: dRows, cols: dCols, count: n });
};

export const groupAt = (sheet: Sheet, row: number, col: number): { group: SparklineGroup; index: number } | null => {
  for (const g of sheet.sparklines ?? []) {
    const index = g.sparklines.findIndex((s) => s.row === row && s.col === col);
    if (index >= 0) return { group: g, index };
  }
  return null;
};

export const groupsIn = (sheet: Sheet, r: CellRange): SparklineGroup[] =>
  (sheet.sparklines ?? []).filter((g) => g.sparklines.some((s) => s.row >= r.r1 && s.row <= r.r2 && s.col >= r.c1 && s.col <= r.c2));

/** Values of one sparkline. Hidden rows and columns are left out unless the group shows them. */
export const sparkValues = (store: GridStore, sheetIdx: number, g: SparklineGroup, sp: Sparkline): (number | null)[] => {
  const p = parseRef(sp.ref);
  if (!p) return [];
  const si = p.sheet === undefined ? sheetIdx : store.wb.sheets.findIndex((s) => s.name.toLowerCase() === p.sheet!.toLowerCase());
  const sh = store.wb.sheets[si];
  if (!sh) return [];
  const r = p.range;
  const out: (number | null)[] = [];
  for (let row = r.r1; row <= r.r2 && out.length < 2000; row++) for (let col = r.c1; col <= r.c2 && out.length < 2000; col++) {
    if (!g.displayHidden && (sh.rows.get(row)?.hidden || sh.cols.get(col)?.hidden)) continue;
    const v = store.value(row, col, si);
    out.push(typeof v === 'number' && Number.isFinite(v) ? v : v === null && g.displayEmptyCellsAs === 'zero' ? 0 : null);
  }
  return out;
};

export interface SparkScale { min: number; max: number }

/** Axis range for each sparkline of a group, from its axis settings. */
export const groupScales = (g: SparklineGroup, all: (number | null)[][]): SparkScale[] => {
  const own = all.map((v) => {
    const n = v.filter((x): x is number => x !== null);
    return n.length ? { min: Math.min(...n), max: Math.max(...n) } : { min: 0, max: 0 };
  });
  const gmin = Math.min(...own.map((o) => o.min)), gmax = Math.max(...own.map((o) => o.max));
  return own.map((o) => ({
    min: g.minAxis === 'custom' && g.manualMin !== undefined ? g.manualMin : g.minAxis === 'group' ? gmin : o.min,
    max: g.maxAxis === 'custom' && g.manualMax !== undefined ? g.manualMax : g.maxAxis === 'group' ? gmax : o.max,
  }));
};

export interface SparkMark { x: number; y: number; color: string }
export interface SparkBar { x: number; w: number; y0: number; y1: number; color: string }
export interface SparkGeom { lines: [number, number][][]; bars: SparkBar[]; marks: SparkMark[]; axis: number | null }

/**
 * Shapes of one sparkline inside a box. `xs` (dates) space points by value; otherwise points are evenly spaced.
 * Line sparklines highlight points with markers; column and win/loss ones colour the bars.
 */
export const sparkGeometry = (g: SparklineGroup, vals: (number | null)[], box: { x: number; y: number; w: number; h: number }, sc: SparkScale, xs?: (number | null)[]): SparkGeom => {
  const out: SparkGeom = { lines: [], bars: [], marks: [], axis: null };
  const n = vals.length;
  if (!n) return out;
  const pad = Math.max(1.5, Math.min(box.h, box.w) * 0.12);
  const x0 = box.x + pad, w = Math.max(1, box.w - pad * 2), y0 = box.y + pad, h = Math.max(1, box.h - pad * 2);
  const winLoss = g.kind === 'winLoss';
  let lo = winLoss ? -1 : sc.min, hi = winLoss ? 1 : sc.max;
  if (g.kind === 'column' && lo > 0 && g.minAxis !== 'custom') lo = 0;
  if (g.kind === 'column' && hi < 0 && g.maxAxis !== 'custom') hi = 0;
  if (hi === lo) { hi += 1; lo -= g.kind === 'line' ? 1 : 0; }
  const Y = (v: number) => y0 + h - ((Math.min(Math.max(v, lo), hi) - lo) / (hi - lo)) * h;
  const dates = xs && xs.length === n && xs.every((d) => d !== null) ? (xs as number[]) : null;
  const dmin = dates ? Math.min(...dates) : 0, dmax = dates ? Math.max(...dates) : 0;
  const slot = (i: number) => (dates ? (dmax > dmin ? (dates[i]! - dmin) / (dmax - dmin) : 0.5) : n === 1 ? 0.5 : g.kind === 'line' ? i / (n - 1) : (i + 0.5) / n);
  const X = (i: number) => { const t = slot(i); return x0 + (g.rightToLeft ? 1 - t : t) * w; };
  const nums = vals.map((v, i) => ({ v, i })).filter((p): p is { v: number; i: number } => p.v !== null);
  if (!nums.length) return out;
  const maxV = Math.max(...nums.map((p) => p.v)), minV = Math.min(...nums.map((p) => p.v));
  const firstI = nums[0]!.i, lastI = nums.at(-1)!.i;
  const special = (p: { v: number; i: number }): string | undefined => {
    if (g.high && p.v === maxV) return g.highColor;
    if (g.low && p.v === minV) return g.lowColor;
    if (g.first && p.i === firstI) return g.firstColor;
    if (g.last && p.i === lastI) return g.lastColor;
    if (g.negative && p.v < 0) return g.negativeColor;
    return undefined;
  };
  const color = g.color ?? '#12A565';
  if (g.showAxis && lo < 0 && hi > 0) out.axis = Y(0);
  else if (winLoss) out.axis = g.showAxis ? Y(0) : null;
  if (g.kind === 'line') {
    let run: [number, number][] = [];
    vals.forEach((v, i) => {
      if (v === null) { if (g.displayEmptyCellsAs !== 'span' && run.length) { out.lines.push(run); run = []; } return; }
      run.push([X(i), Y(v)]);
    });
    if (run.length) out.lines.push(run);
    for (const p of nums) {
      const c = special(p) ?? (g.markers ? g.markersColor ?? color : undefined);
      if (c) out.marks.push({ x: X(p.i), y: Y(p.v), color: c });
    }
    return out;
  }
  const bw = Math.max(1, (dates ? w / Math.max(n, 2) : w / n) * 0.78);
  for (const p of nums) {
    if (winLoss && p.v === 0) continue;
    const v = winLoss ? Math.sign(p.v) : p.v;
    const base = Y(winLoss ? 0 : Math.min(Math.max(0, lo), hi));
    const top = Y(v);
    out.bars.push({ x: X(p.i) - bw / 2, w: bw, y0: Math.min(base, top), y1: Math.max(base, top), color: special(p) ?? (p.v < 0 && g.negative ? g.negativeColor ?? color : color) });
  }
  return out;
};
