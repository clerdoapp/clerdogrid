// Axis scales and the axis frame shared by flat category and x-y charts.
import type { ReactNode } from 'react';
import type { AxisOptions, Chart, DisplayUnits } from '../../model/workbook';
import { isDateFormat } from '../../formats/numfmt';
import { formatNumber, niceScale, type ResolvedChart } from './chartData';
import { AXIS, FONT, GRID, ink, lineCss, muted, textW, trunc, type DrawCtx } from './svgKit';
import { N, Tr } from '../i18n';

export interface ValScale {
  min: number; max: number; ticks: number[]; minor: number[]; log: number;
  /** Position of v along the axis, 0 at the start and 1 at the end (after reversing). */
  pos: (v: number) => number;
  reverse: boolean;
  /** Divisor for tick labels and the unit name shown by the axis. */
  unit: number; unitLabel?: string;
}

const UNITS: Record<DisplayUnits, [number, string]> = {
  hundreds: [1e2, N('Hundreds')], thousands: [1e3, N('Thousands')], tenThousands: [1e4, N('Ten thousands')], hundredThousands: [1e5, N('Hundred thousands')],
  millions: [1e6, N('Millions')], tenMillions: [1e7, N('Ten millions')], hundredMillions: [1e8, N('Hundred millions')], billions: [1e9, N('Billions')], trillions: [1e12, N('Trillions')],
};
export const DISPLAY_UNITS = Object.entries(UNITS).map(([k, [, label]]) => ({ value: k as DisplayUnits, label }));

const clean = (v: number): number => +v.toPrecision(12);

/** Axis settings with the older chart-wide fields folded into the value axis. */
export const axisOf = (chart: Chart, which: 'x' | 'y' | 'y2'): AxisOptions => {
  const o = chart.options ?? {};
  const a = o.axes?.[which] ?? {};
  if (which !== 'y') return a;
  return {
    ...a,
    ...(o.min !== undefined && a.min === undefined ? { min: o.min } : {}),
    ...(o.max !== undefined && a.max === undefined ? { max: o.max } : {}),
    ...(o.numFmt && !a.numFmt ? { numFmt: o.numFmt } : {}),
    ...(o.gridlines === false && a.majorGrid === undefined ? { majorGrid: false } : {}),
  };
};

export const makeValueScale = (lo: number, hi: number, nt: number, ax: AxisOptions = {}): ValScale => {
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) { lo = 0; hi = 1; }
  const reverse = !!ax.reverse;
  const [unit, unitLabel] = ax.dispUnits ? UNITS[ax.dispUnits] ?? [1, undefined] : [1, undefined];
  const done = (min: number, max: number, ticks: number[], minor: number[], log: number): ValScale => {
    const span = log ? Math.log(max / min) : max - min || 1;
    const raw = log ? (v: number) => (v > 0 ? Math.log(v / min) / span : -0.05) : (v: number) => (v - min) / span;
    return { min, max, ticks, minor, log, pos: reverse ? (v) => 1 - raw(v) : raw, reverse, unit, unitLabel };
  };
  if (ax.logBase && ax.logBase >= 2) {
    const b = Math.min(1000, ax.logBase);
    if (hi <= 0) hi = 1;
    if (lo <= 0) lo = Math.min(hi / b, 1);
    let e0 = Math.floor(Math.log(lo) / Math.log(b) + 1e-9), e1 = Math.ceil(Math.log(hi) / Math.log(b) - 1e-9);
    if (e1 <= e0) e1 = e0 + 1;
    let min = b ** e0, max = b ** e1;
    if (ax.min !== undefined && ax.min > 0) { min = ax.min; e0 = Math.log(min) / Math.log(b); }
    if (ax.max !== undefined && ax.max > min) { max = ax.max; e1 = Math.log(max) / Math.log(b); }
    const ticks: number[] = [], minor: number[] = [];
    const every = Math.max(1, Math.ceil((Math.ceil(e1) - Math.floor(e0)) / 10));
    for (let e = Math.ceil(e0 - 1e-9); e <= e1 + 1e-9; e += every) ticks.push(clean(b ** e));
    if (b <= 10) for (let e = Math.floor(e0); e < e1; e++) for (let k = 2; k < b; k++) { const v = k * b ** e; if (v > min && v < max) minor.push(v); }
    return done(min, max, ticks, minor, b);
  }
  let min: number, max: number, step: number, ticks: number[];
  if (ax.majorUnit && ax.majorUnit > 0) {
    step = ax.majorUnit;
    min = ax.min ?? Math.floor(Math.min(lo, hi) / step + 1e-9) * step;
    max = ax.max ?? Math.ceil(Math.max(lo, hi) / step - 1e-9) * step;
    if (max <= min) max = min + step;
    if ((max - min) / step > 500) step = (max - min) / 500;
    ticks = [];
    for (let v = min; v <= max + step * 1e-6; v += step) ticks.push(clean(v));
  } else {
    let l = ax.min ?? lo, h = ax.max ?? hi;
    if (l >= h) h = l + (ax.max === undefined ? Math.max(1, Math.abs(l) * 0.1) : 0) || l + 1;
    if (l >= h) l = h - 1;
    const s = niceScale(l, h, nt);
    step = s.step;
    min = ax.min ?? s.min; max = ax.max ?? s.max;
    ticks = s.ticks.filter((t) => t >= min - 1e-9 && t <= max + 1e-9);
    if (ax.min !== undefined && (!ticks.length || Math.abs(ticks[0]! - min) > step * 0.25)) ticks.unshift(min);
    if (ax.max !== undefined && Math.abs(ticks.at(-1)! - max) > step * 0.25) ticks.push(max);
  }
  const mstep = ax.minorUnit && ax.minorUnit > 0 ? ax.minorUnit : step / 5;
  const minor: number[] = [];
  if ((max - min) / mstep <= 1000) for (let v = min; v <= max + mstep * 1e-6; v += mstep) { const cv = clean(v); if (!ticks.some((t) => Math.abs(t - cv) < mstep * 1e-6)) minor.push(cv); }
  return done(min, max, ticks, minor, 0);
};

export const tickLabel = (v: number, s: ValScale, fmt: string): string => formatNumber(v / s.unit, fmt, s.unit === 1);

// ---------- Dates on the category axis ----------

const EPOCH = Date.UTC(1899, 11, 30);
export const serialDate = (n: number): Date => new Date(EPOCH + Math.round(n * 86400000));
const unitIndex = (serial: number, unit: 'days' | 'months' | 'years'): number => {
  if (unit === 'days') return Math.floor(serial);
  const d = serialDate(serial);
  return unit === 'months' ? d.getUTCFullYear() * 12 + d.getUTCMonth() : d.getUTCFullYear();
};
const indexSerial = (idx: number, unit: 'days' | 'months' | 'years'): number => {
  if (unit === 'days') return idx;
  const t = unit === 'months' ? Date.UTC(Math.floor(idx / 12), idx % 12, 1) : Date.UTC(idx, 0, 1);
  return (t - EPOCH) / 86400000;
};

export interface DateCats { unit: 'days' | 'months' | 'years'; first: number; count: number; bucket: number[]; labels: string[] }

/** Date axis layout when the categories are dates (or the axis is set to dates). */
export const dateCategories = (data: ResolvedChart, ax: AxisOptions): DateCats | null => {
  if (ax.dateAxis === false) return null;
  const vals = data.catValues;
  if (!vals.length || vals.some((v) => v === null)) return null;
  if (ax.dateAxis !== true && !isDateFormat(data.catFmt)) return null;
  const nums = vals as number[];
  let unit = ax.baseUnit;
  if (!unit) {
    const ds = nums.map(serialDate);
    unit = ds.every((d) => d.getUTCDate() === 1) ? (ds.every((d) => d.getUTCMonth() === 0) && nums.length > 1 ? 'years' : 'months') : 'days';
  }
  const bucket = nums.map((v) => unitIndex(v, unit!));
  const first = Math.min(...bucket), last = Math.max(...bucket);
  const count = last - first + 1;
  if (count > 4000 || count < 1) return null;
  const fmt = unit === 'years' ? 'yyyy' : unit === 'months' ? 'mmm yy' : isDateFormat(data.catFmt) ? data.catFmt : 'd mmm';
  const labels = Array.from({ length: count }, (_, i) => formatNumber(indexSerial(first + i, unit!), fmt));
  return { unit, first, count, bucket: bucket.map((b) => b - first), labels };
};

// ---------- The frame ----------

export interface ValSpec { lo: number; hi: number; fmt: string; ax: AxisOptions; title?: string; /** Automatic crossing at the minimum rather than at zero. */ crossMin?: boolean }
export interface TableRow { name: string; color: string; cells: string[] }

export interface Frame {
  px: number; py: number; pw: number; ph: number;
  /** Number of category slots (date buckets on a date axis). */
  n: number; band: number;
  /** Centre of point i's category slot. */
  catPos: (i: number) => number;
  /** Position of a continuous x value: category number (1-based) on category axes, the value itself on value axes. */
  xPos: (x: number) => number;
  valPos: (v: number) => number;
  valPos2: (v: number) => number;
  /** Where bars and areas start: the base axis line. */
  zero: number; zero2: number;
  scale: ValScale; scale2: ValScale | null; xScale: ValScale | null;
  slotOf: (i: number) => number;
}

interface FrameOpts {
  /** Category labels, or a numeric x spec for x-y charts. */
  cats?: string[];
  xVal?: ValSpec;
  date?: DateCats | null;
  catAx: AxisOptions;
  horizontal: boolean;
  y: ValSpec;
  y2?: ValSpec | null;
  /** Data table rows drawn under the plot, lined up with the categories. */
  table?: TableRow[] | null;
  tableOpts?: { keys?: boolean; horizontal?: boolean; vertical?: boolean; outline?: boolean };
  xTitle?: string;
}

const tickMark = (els: ReactNode[], key: string, kind: AxisOptions['majorTick'], x: number, y: number, vertical: boolean, outward: number, len = 5): void => {
  if (!kind || kind === 'none') return;
  const a = kind === 'in' ? -len : kind === 'cross' ? -len : 0, b = kind === 'in' ? 0 : len;
  if (vertical) els.push(<line key={key} x1={x + a * outward} x2={x + b * outward} y1={y} y2={y} style={AXIS} />);
  else els.push(<line key={key} x1={x} x2={x} y1={y + a * outward} y2={y + b * outward} style={AXIS} />);
};

export const buildFrame = (c: DrawCtx, o: FrameOpts): Frame => {
  const { plot, thumb, compact, els, chart } = c;
  const horizontal = o.horizontal;
  const showLabels = !thumb;
  const cats = o.date ? o.date.labels : o.cats ?? [];
  const n = o.xVal ? 1 : Math.max(1, cats.length);
  const yTitle = !thumb && !compact ? o.y.title ?? '' : '';
  const xTitle = !thumb && !compact ? o.xTitle ?? '' : '';
  const y2Title = !thumb && !compact && o.y2 ? o.y2.title ?? '' : '';
  const valueLen = horizontal ? plot.w : plot.h;
  const nt = thumb ? 3 : Math.max(3, Math.min(8, Math.floor(valueLen / (horizontal ? 70 : 40))));
  const yAx = o.y.ax, y2Ax = o.y2?.ax ?? {};
  const scale = makeValueScale(o.y.lo, o.y.hi, nt, yAx);
  const scale2 = o.y2 ? makeValueScale(o.y2.lo, o.y2.hi, scale.ticks.length, y2Ax) : null;
  const xScale = o.xVal ? makeValueScale(o.xVal.lo, o.xVal.hi, thumb ? 3 : Math.max(3, Math.floor(plot.w / 70)), o.xVal.ax) : null;
  const yLabelsOn = showLabels && !yAx.hidden && yAx.labelPos !== 'none';
  const y2LabelsOn = showLabels && !!scale2 && !y2Ax.hidden && y2Ax.labelPos !== 'none';
  const catLabelsOn = showLabels && !o.catAx.hidden && o.catAx.labelPos !== 'none' && !o.table;
  const tickText = scale.ticks.map((t) => tickLabel(t, scale, o.y.fmt));
  const tickText2 = scale2 ? scale2.ticks.map((t) => tickLabel(t, scale2, o.y2!.fmt)) : [];
  const xTickText = xScale ? xScale.ticks.map((t) => tickLabel(t, xScale, o.xVal!.fmt)) : [];
  const maxCat = xScale ? Math.max(0, ...xTickText.map((s) => textW(s))) : Math.max(0, ...cats.map((s) => textW(s)));

  let px = plot.x, pw = plot.w, ph = plot.h;
  const py = plot.y;
  let rotate = false, skip = 1;
  const tableRowH = 16;
  const tableRows = o.table ? o.table.length + 1 : 0;
  const keyW = o.table ? Math.min(plot.w * 0.28, Math.max(...o.table.map((r) => textW(r.name))) + (o.tableOpts?.keys !== false ? 22 : 10)) : 0;
  const yUnitLabel = scale.unitLabel && yAx.dispUnitsLabel !== false && !thumb ? Tr(scale.unitLabel) : '';
  if (!thumb) {
    if (yTitle) { px += 18; pw -= 18; }
    if (y2Title) pw -= 18;
    if (xTitle) ph -= 18;
    if (yUnitLabel) { if (horizontal) ph -= 14; else { px += 14; pw -= 14; } }
    if (o.table) ph -= tableRows * tableRowH + 4;
    if (horizontal) {
      const g = catLabelsOn ? Math.min(maxCat, plot.w * 0.3) + 8 : 4;
      px += g; pw -= g;
      if (yLabelsOn) ph -= 18;
      skip = Math.max(1, Math.ceil(13 / (ph / n)));
    } else {
      const g = yLabelsOn ? Math.max(...tickText.map((t) => textW(t))) + 8 : 4;
      // Value labels sit on the right when the axis crosses at the last category.
      const yRight = yAx.labelPos === 'high' || (yAx.labelPos !== 'low' && !xScale && !o.table && (o.catAx.crosses === 'max') !== !!o.catAx.reverse);
      const lead = Math.max(yRight ? 4 : g, keyW - (yTitle ? 18 : 0) - (yUnitLabel ? 14 : 0));
      px += lead;
      pw -= lead + (yRight ? g : 0);
      if (y2LabelsOn) pw -= Math.max(...tickText2.map((t) => textW(t))) + 8;
      const band = pw / n;
      if (xScale) { ph -= 18; pw -= 6; }
      else if (catLabelsOn && maxCat + 6 > band) {
        rotate = true;
        ph -= Math.min(maxCat * 0.72 + 10, plot.h * 0.3);
        skip = Math.max(1, Math.ceil(14 / band));
      } else if (catLabelsOn) ph -= 18;
    }
  }
  pw = Math.max(10, pw); ph = Math.max(10, ph);
  const band = (horizontal ? ph : pw) / n;
  const valLen = horizontal ? pw : ph;
  const valPos = (v: number) => (horizontal ? px + scale.pos(v) * pw : py + ph - scale.pos(v) * ph);
  const valPos2 = scale2 ? (v: number) => py + ph - scale2.pos(v) * ph : valPos;
  const catRev = !!o.catAx.reverse;
  const slotOf = (i: number) => (o.date ? o.date.bucket[i] ?? i : i);
  // Bar charts list categories top to bottom, like the rows they come from.
  const slotPos = (k: number) => {
    const kk = catRev ? n - 1 - k : k;
    return horizontal ? py + band * (kk + 0.5) : px + band * (kk + 0.5);
  };
  const catPos = (i: number) => slotPos(slotOf(i));
  const xPos = xScale
    ? (x: number) => px + xScale.pos(x) * pw
    : (x: number) => { const kk = catRev ? n - x : x - 1; return horizontal ? py + band * (kk + 0.5) : px + band * (kk + 0.5); };

  // Where the base axis crosses the value axis.
  const crossVal = (s: ValScale, ax: AxisOptions, atMin = false) => {
    if (ax.crosses === 'min') return s.reverse ? s.max : s.min;
    if (ax.crosses === 'max') return s.reverse ? s.min : s.max;
    if (ax.crossesAt !== undefined && Number.isFinite(ax.crossesAt)) return Math.min(Math.max(ax.crossesAt, Math.min(s.min, s.max)), Math.max(s.min, s.max));
    return s.log || atMin ? (s.reverse ? s.max : s.min) : Math.min(Math.max(0, s.min), s.max);
  };
  const zero = valPos(crossVal(scale, yAx, o.y.crossMin));
  const zero2 = scale2 ? valPos2(crossVal(scale2, y2Ax)) : zero;
  // Where the value axis sits along the base axis.
  const valAxisAt = (() => {
    const ca = o.catAx;
    if (xScale) {
      const v = ca.crosses === 'max' ? xScale.max : ca.crosses === 'min' ? xScale.min : ca.crossesAt !== undefined ? ca.crossesAt : xScale.log || o.xVal?.crossMin ? xScale.min : Math.min(Math.max(0, xScale.min), xScale.max);
      return px + xScale.pos(v) * pw;
    }
    const atMax = ca.crosses === 'max';
    if (ca.crossesAt !== undefined && ca.crossesAt >= 1) return horizontal ? xPos(ca.crossesAt) + band / 2 : xPos(ca.crossesAt) - band / 2;
    const start = horizontal ? (catRev !== atMax ? py : py + ph) : (catRev !== atMax ? px + pw : px);
    return start;
  })();

  const lineScale = c.lineScale;
  const yMajorGrid = yAx.majorGrid !== false && !(c.chart.options?.look !== undefined && LOOK_NO_GRID.has(c.chart.options.look));
  const axisStroke = (ax: AxisOptions) => ({ ...AXIS, ...lineCss(ax.line, 'var(--chart-axis)', 1 * lineScale) });

  // Minor then major gridlines for the value axis.
  const gridLine = (k: string, p: number, minor: boolean) => {
    const st = minor ? { ...GRID, strokeOpacity: 0.55, strokeDasharray: '2 3' } : GRID;
    els.push(horizontal ? <line key={k} x1={p} x2={p} y1={py} y2={py + ph} style={st} /> : <line key={k} x1={px} x2={px + pw} y1={p} y2={p} style={st} />);
  };
  if (yAx.minorGrid) scale.minor.forEach((t, k) => gridLine(`gm${k}`, valPos(t), true));
  if (yMajorGrid) scale.ticks.forEach((t, k) => gridLine(`g${k}`, valPos(t), false));
  // Gridlines along the base axis.
  const catAx = o.catAx;
  const baseGrid = (k: string, p: number, minor: boolean) => {
    const st = minor ? { ...GRID, strokeOpacity: 0.55, strokeDasharray: '2 3' } : GRID;
    els.push(horizontal ? <line key={k} x1={px} x2={px + pw} y1={p} y2={p} style={st} /> : <line key={k} x1={p} x2={p} y1={py} y2={py + ph} style={st} />);
  };
  if (xScale) {
    if (catAx.minorGrid) xScale.minor.forEach((t, k) => baseGrid(`xgm${k}`, px + xScale.pos(t) * pw, true));
    if (catAx.majorGrid ?? true) xScale.ticks.forEach((t, k) => baseGrid(`xg${k}`, px + xScale.pos(t) * pw, false));
  } else {
    if (catAx.minorGrid) for (let k = 0; k < n; k++) baseGrid(`cgm${k}`, slotPos(k), true);
    if (catAx.majorGrid) for (let k = 0; k <= n; k++) baseGrid(`cg${k}`, horizontal ? py + k * band : px + k * band, false);
  }

  // Value axis labels.
  const yLabelX = (() => {
    if (yAx.labelPos === 'high') return { x: px + pw + 6, anchor: 'start' as const };
    if (yAx.labelPos === 'low') return { x: px - 6, anchor: 'end' as const };
    return valAxisAt > px + pw / 2 ? { x: px + pw + 6, anchor: 'start' as const } : { x: px - 6, anchor: 'end' as const };
  })();
  if (yLabelsOn) scale.ticks.forEach((t, k) => {
    const p = valPos(t);
    if (horizontal) {
      const ly = yAx.labelPos === 'high' ? py - 5 : py + ph + 13;
      els.push(<text key={`t${k}`} x={Math.min(p, plot.x + plot.w - textW(tickText[k]!) / 2)} y={ly} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>);
    } else els.push(<text key={`t${k}`} x={yLabelX.x} y={p + 4} textAnchor={yLabelX.anchor} style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>);
  });
  if (!thumb) scale.ticks.forEach((t, k) => {
    const p = valPos(t);
    if (horizontal) tickMark(els, `tm${k}`, yAx.majorTick, p, zero === px ? py + ph : py + ph, false, 1);
    else tickMark(els, `tm${k}`, yAx.majorTick, valAxisAt, p, true, valAxisAt > px + pw / 2 ? 1 : -1);
  });
  if (!thumb && yAx.minorTick && yAx.minorTick !== 'none') scale.minor.forEach((t, k) => {
    const p = valPos(t);
    if (horizontal) tickMark(els, `tmm${k}`, yAx.minorTick, p, py + ph, false, 1, 3);
    else tickMark(els, `tmm${k}`, yAx.minorTick, valAxisAt, p, true, valAxisAt > px + pw / 2 ? 1 : -1, 3);
  });
  if (scale2) {
    if (y2LabelsOn) scale2.ticks.forEach((t, k) => els.push(<text key={`t2${k}`} x={px + pw + 6} y={valPos2(t) + 4} textAnchor="start" style={{ ...muted, fontSize: FONT }}>{tickText2[k]}</text>));
    if (!y2Ax.hidden) els.push(<line key="axis2" x1={px + pw} x2={px + pw} y1={py} y2={py + ph} style={axisStroke(y2Ax)} />);
    if (!thumb) scale2.ticks.forEach((t, k) => tickMark(els, `tm2${k}`, y2Ax.majorTick, px + pw, valPos2(t), true, 1));
  }
  // Axis lines: the base axis at the crossing value, the value axis where it crosses the base.
  if (!catAx.hidden) els.push(horizontal
    ? <line key="axis" x1={zero} x2={zero} y1={py} y2={py + ph} style={axisStroke(catAx)} />
    : <line key="axis" x1={px} x2={px + pw} y1={zero} y2={zero} style={axisStroke(catAx)} />);
  if (xScale && !yAx.hidden) els.push(<line key="vaxis" x1={valAxisAt} x2={valAxisAt} y1={py} y2={py + ph} style={axisStroke(yAx)} />);
  else if ((yAx.line && !yAx.hidden) || (yAx.majorTick && yAx.majorTick !== 'none')) els.push(horizontal
    ? <line key="vaxis" x1={px} x2={px + pw} y1={valAxisAt} y2={valAxisAt} style={axisStroke(yAx)} />
    : <line key="vaxis" x1={valAxisAt} x2={valAxisAt} y1={py} y2={py + ph} style={axisStroke(yAx)} />);

  // Base axis labels and tick marks.
  const catLabelY = catAx.labelPos === 'high' ? py - 6 : catAx.labelPos === 'low' ? py + ph + 14 : (horizontal ? 0 : Math.max(py + ph, zero) + 14);
  if (xScale) {
    if (catLabelsOn) xScale.ticks.forEach((t, k) => els.push(<text key={`tx${k}`} x={Math.min(px + xScale.pos(t) * pw, plot.x + plot.w - textW(xTickText[k]!) / 2)} y={catAx.labelPos === 'high' ? py - 6 : py + ph + 14} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{xTickText[k]}</text>));
    if (!thumb) xScale.ticks.forEach((t, k) => tickMark(els, `txm${k}`, catAx.majorTick, px + xScale.pos(t) * pw, zero, false, 1));
  } else {
    if (catLabelsOn) cats.forEach((label, k) => {
      if (k % skip) return;
      const p = slotPos(k);
      if (horizontal) {
        const lx = catAx.labelPos === 'high' ? px + pw + 6 : px - 6;
        els.push(<text key={`c${k}`} x={lx} y={p + 4} textAnchor={catAx.labelPos === 'high' ? 'start' : 'end'} style={{ ...muted, fontSize: FONT }}>{trunc(label, Math.min(maxCat, plot.w * 0.3))}</text>);
      } else if (rotate) {
        const y0 = catAx.labelPos === 'low' || catAx.labelPos === 'high' ? catLabelY - 4 : py + ph + 10;
        els.push(<text key={`c${k}`} x={p} y={y0} textAnchor="end" transform={`rotate(-45 ${p} ${y0})`} style={{ ...muted, fontSize: FONT }}>{trunc(label, (plot.h * 0.3 - 10) / 0.72)}</text>);
      } else els.push(<text key={`c${k}`} x={p} y={catAx.labelPos === 'nextTo' || !catAx.labelPos ? py + ph + 14 : catLabelY} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{label}</text>);
    });
    if (!thumb) for (let k = 0; k <= n; k++) {
      const p = horizontal ? py + k * band : px + k * band;
      if (horizontal) tickMark(els, `cm${k}`, catAx.majorTick, zero, p, true, -1);
      else tickMark(els, `cm${k}`, catAx.majorTick, p, zero, false, 1);
    }
  }
  if (yUnitLabel) els.push(horizontal
    ? <text key="yu" x={px + pw} y={py + ph + 28} textAnchor="end" style={{ ...muted, fontSize: 10, fontStyle: 'italic' }}>{yUnitLabel}</text>
    : <text key="yu" x={plot.x + (yTitle ? 26 : 8)} y={py + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + (yTitle ? 26 : 8)} ${py + ph / 2})`} style={{ ...muted, fontSize: 10, fontStyle: 'italic' }}>{yUnitLabel}</text>);
  if (yTitle) els.push(horizontal
    ? <text key="yt" x={px + pw / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(yTitle, pw)}</text>
    : <text key="yt" x={plot.x + 10} y={py + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(yTitle, ph)}</text>);
  if (y2Title) els.push(<text key="y2t" x={plot.x + plot.w - 8} y={py + ph / 2} textAnchor="middle" transform={`rotate(90 ${plot.x + plot.w - 8} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(y2Title, ph)}</text>);
  if (xTitle) els.push(horizontal
    ? <text key="xt" x={plot.x + 10} y={py + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(xTitle, ph)}</text>
    : <text key="xt" x={px + pw / 2} y={plot.y + plot.h - 4 - (o.table ? 0 : 0)} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(xTitle, pw)}</text>);

  // Data table: categories across, one row per series.
  if (o.table && !thumb) {
    const t = o.tableOpts ?? {};
    const top = py + ph + 4;
    const x0 = px - keyW;
    const lineSt = { stroke: 'var(--chart-grid)', strokeWidth: 1 };
    const rows = o.table;
    const allRows = tableRows;
    els.push(<g key="dtable">
      {(t.outline ?? true) && <rect x={x0} y={top} width={keyW + pw} height={allRows * tableRowH} style={{ fill: 'none', ...lineSt }} />}
      {(t.horizontal ?? true) && Array.from({ length: allRows - 1 }, (_, r) => <line key={`h${r}`} x1={x0} x2={px + pw} y1={top + (r + 1) * tableRowH} y2={top + (r + 1) * tableRowH} style={lineSt} />)}
      {(t.vertical ?? true) && Array.from({ length: n + 1 }, (_, k) => <line key={`v${k}`} x1={px + k * band} x2={px + k * band} y1={top} y2={top + allRows * tableRowH} style={lineSt} />)}
      {cats.map((label, k) => <text key={`hc${k}`} x={slotPos(k)} y={top + 12} textAnchor="middle" style={{ ...muted, fontSize: 10 }}>{trunc(label, band - 4, 10)}</text>)}
      {rows.map((r, ri) => (
        <g key={`r${ri}`}>
          {t.keys !== false && <rect x={x0 + 5} y={top + (ri + 1) * tableRowH + 4} width={9} height={9} rx={2} style={{ fill: r.color }} />}
          <text x={x0 + (t.keys !== false ? 18 : 5)} y={top + (ri + 1) * tableRowH + 12} style={{ ...ink, fontSize: 10 }}>{trunc(r.name, keyW - 22, 10)}</text>
          {r.cells.map((v, k) => <text key={k} x={slotPos(k)} y={top + (ri + 1) * tableRowH + 12} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{trunc(v, band - 4, 10)}</text>)}
        </g>
      ))}
    </g>);
  }

  if (c.hot !== null && !thumb && !xScale && c.hot < n) {
    const b = slotPos(c.hot) - band / 2;
    els.push(horizontal
      ? <rect key="hot" x={px} y={b} width={pw} height={band} style={{ fill: 'var(--chart-hover)' }} />
      : <rect key="hot" x={b} y={py} width={band} height={ph} style={{ fill: 'var(--chart-hover)' }} />);
  }
  void valLen; void chart;
  return { px, py, pw, ph, n, band, catPos, xPos, valPos, valPos2, zero, zero2, scale, scale2, xScale, slotOf };
};

/** Looks that hide gridlines (kept in sync with LOOKS in svgKit). */
const LOOK_NO_GRID = new Set([3]);

/** Category slot under the pointer, or null outside the plot. */
export const slotAt = (f: Frame, horizontal: boolean, reverse: boolean, x: number, y: number): number | null => {
  if (x < f.px || x > f.px + f.pw || y < f.py || y > f.py + f.ph) return null;
  let k = Math.floor(((horizontal ? y - f.py : x - f.px) / (horizontal ? f.ph : f.pw)) * f.n);
  k = Math.min(f.n - 1, Math.max(0, k));
  return reverse ? f.n - 1 - k : k;
};
