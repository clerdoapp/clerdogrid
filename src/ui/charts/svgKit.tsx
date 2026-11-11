// Shared pieces for the chart drawing modules.
import type { CSSProperties, ReactNode } from 'react';
import type { Chart, DataLabelOptions, FillFormat, LineFormat, MarkerSymbol, PointFormat } from '../../model/workbook';
import { formatNumber, type ResolvedChart, type ResolvedSeries } from './chartData';
import type { Traits } from './chartTypes';
import { N } from '../i18n';

export const FONT = 11;
export const textW = (s: string, size = FONT): number => s.length * size * 0.56;
export const trunc = (s: string, maxW: number, size = FONT): string => {
  if (textW(s, size) <= maxW) return s;
  const n = Math.max(0, Math.floor(maxW / (size * 0.56)) - 1);
  return n ? `${s.slice(0, n)}…` : '';
};

export const ink = { fill: 'var(--chart-ink)' } as const;
export const muted = { fill: 'var(--chart-muted)' } as const;
export const onFill = { fill: 'var(--chart-on)' } as const;
export const GRID = { stroke: 'var(--chart-grid)', strokeWidth: 1 } as const;
export const AXIS = { stroke: 'var(--chart-axis)', strokeWidth: 1 } as const;
export const WF = { up: 'var(--chart-up)', down: 'var(--chart-down)', total: 'var(--chart-total)' } as const;

export interface TipRow { color: string; label: string; value: string }
export interface Tip { x: number; y: number; title: string; rows: TipRow[] }
export interface Box { x: number; y: number; w: number; h: number }
export interface Hover { onMove: (x: number, y: number) => void; onLeave: () => void }
export interface LegendEntry { label: string; color: string; line: boolean; dash?: string; marker?: MarkerSymbol }

export interface DrawCtx {
  chart: Chart; data: ResolvedChart; tr: Traits; plot: Box; thumb: boolean; compact: boolean; hot: number | null; els: ReactNode[];
  /** Gradient and clip definitions. */
  defs: ReactNode[];
  setTip: (t: Tip | null) => void; setHot: (i: number | null) => void; hover: Hover; clip: string;
  /** Element-based charts set tips from the shapes themselves. */
  elementTips: boolean;
  style: number;
  /** Unique prefix for ids inside this chart. */
  uid: string;
  /** Room taken at the bottom by a data table, filled in by the axis frame. */
  lineScale: number;
  /** Repaint request after async work (map shapes). */
  refresh: () => void;
}

export const fmtv = (v: number, fmt: string) => formatNumber(v, fmt);
export const ptPx = (pt: number): number => (pt * 96) / 72;

const DASH: Record<string, (w: number) => string | undefined> = {
  solid: () => undefined, dot: (w) => `${w} ${w * 1.6}`, dash: (w) => `${w * 4} ${w * 2.5}`, dashDot: (w) => `${w * 4} ${w * 2} ${w} ${w * 2}`,
  longDash: (w) => `${w * 8} ${w * 3}`, longDashDot: (w) => `${w * 8} ${w * 3} ${w} ${w * 3}`,
};
export const dashArray = (dash: string | undefined, w: number): string | undefined => (dash ? DASH[dash]?.(Math.max(1, w)) : undefined);

/** Stroke style from a line format, falling back to a colour and pixel width. */
export const lineCss = (l: LineFormat | undefined, color: string, width: number): CSSProperties => {
  if (l?.none) return { stroke: 'none' };
  const w = l?.width !== undefined ? ptPx(l.width) : width;
  return { stroke: l?.color ?? color, strokeWidth: w, strokeDasharray: dashArray(l?.dash, w) };
};

let gradSeq = 0;
/** Fill style from a fill format; gradients are added to the chart's defs. */
export const fillCss = (c: Pick<DrawCtx, 'defs' | 'uid'>, f: FillFormat | undefined, color: string, opacity = 1): CSSProperties => {
  if (f?.none) return { fill: 'none' };
  const alpha = f?.transparency !== undefined ? Math.max(0, Math.min(1, 1 - f.transparency / 100)) : 1;
  if (f?.gradient) {
    const g = f.gradient;
    const id = `${c.uid}g${++gradSeq}`;
    const a = ((g.angle ?? 90) * Math.PI) / 180;
    const dx = Math.cos(a) / 2, dy = Math.sin(a) / 2;
    c.defs.push(
      <linearGradient key={id} id={id} x1={0.5 - dx} y1={0.5 - dy} x2={0.5 + dx} y2={0.5 + dy}>
        <stop offset="0" style={{ stopColor: g.from }} /><stop offset="1" style={{ stopColor: g.to }} />
      </linearGradient>,
    );
    return { fill: `url(#${id})`, fillOpacity: alpha * opacity };
  }
  return { fill: f?.color ?? color, fillOpacity: alpha * opacity };
};

export const pointFormat = (s: ResolvedSeries, i: number): PointFormat | undefined => s.src.points?.find((p) => p.idx === i);

/** A point marker centred at x, y. */
export const markerEl = (key: string, symbol: MarkerSymbol | undefined, x: number, y: number, size: number, fill: string, stroke: string, sw = 2): ReactNode => {
  const r = size / 2;
  const st = { fill, stroke, strokeWidth: sw, strokeLinejoin: 'round' as const };
  switch (symbol ?? 'circle') {
    case 'none': return null;
    case 'square': return <rect key={key} x={x - r} y={y - r} width={size} height={size} style={st} />;
    case 'diamond': return <path key={key} d={`M${x},${y - r * 1.25} L${x + r * 1.25},${y} L${x},${y + r * 1.25} L${x - r * 1.25},${y} Z`} style={st} />;
    case 'triangle': return <path key={key} d={`M${x},${y - r * 1.2} L${x + r * 1.1},${y + r * 0.8} L${x - r * 1.1},${y + r * 0.8} Z`} style={st} />;
    case 'x': return <path key={key} d={`M${x - r},${y - r} L${x + r},${y + r} M${x - r},${y + r} L${x + r},${y - r}`} style={{ ...st, fill: 'none', stroke: stroke }} />;
    case 'plus': return <path key={key} d={`M${x - r},${y} L${x + r},${y} M${x},${y - r} L${x},${y + r}`} style={{ ...st, fill: 'none' }} />;
    case 'star': {
      let d = '';
      for (let k = 0; k < 10; k++) {
        const a = -Math.PI / 2 + (k * Math.PI) / 5, rr = k % 2 ? r * 0.5 : r * 1.25;
        d += `${k ? 'L' : 'M'}${(x + Math.cos(a) * rr).toFixed(2)},${(y + Math.sin(a) * rr).toFixed(2)} `;
      }
      return <path key={key} d={`${d}Z`} style={st} />;
    }
    case 'dash': return <line key={key} x1={x - r * 1.3} x2={x + r * 1.3} y1={y} y2={y} style={{ stroke, strokeWidth: Math.max(2, sw + 0.5) }} />;
    case 'dot': return <circle key={key} cx={x} cy={y} r={Math.max(1.5, r * 0.5)} style={{ fill: stroke }} />;
    default: return <circle key={key} cx={x} cy={y} r={r} style={st} />;
  }
};

/** Label settings for a series: its own, else the chart's, else just the value. */
export const labelOpts = (chart: Chart, s: ResolvedSeries): DataLabelOptions | null => {
  if (s.src.labels) {
    const l = s.src.labels;
    return l.value || l.seriesName || l.categoryName || l.percent || l.bubbleSize || l.fromCells ? l : null;
  }
  if (!chart.dataLabels) return null;
  return chart.options?.labels ?? { value: true };
};

/** Text of one data label. */
export const labelText = (
  o: DataLabelOptions, p: { value: number | null; series: string; category: string; percent?: number; size?: number | null; cell?: string },
  fmt: string, defaultPercent = false,
): string => {
  const parts: string[] = [];
  if (o.fromCells && p.cell) parts.push(p.cell);
  if (o.seriesName) parts.push(p.series);
  if (o.categoryName) parts.push(p.category);
  if (o.value && p.value !== null) parts.push(formatNumber(p.value, o.numFmt || fmt));
  if ((o.percent || (defaultPercent && !o.value && !o.seriesName && !o.categoryName)) && p.percent !== undefined) parts.push(formatNumber(p.percent, o.numFmt && /%/.test(o.numFmt) ? o.numFmt : '0%'));
  if (o.bubbleSize && p.size !== null && p.size !== undefined) parts.push(formatNumber(p.size, 'General'));
  return parts.join(o.separator ?? ', ');
};

/** Draws a possibly multi-line equation text with superscripts. */
export const richTextEl = (key: string, runs: { t: string; sup?: boolean }[], x: number, y: number, anchor: 'start' | 'middle' | 'end', size = 10.5): ReactNode => (
  <text key={key} x={x} y={y} textAnchor={anchor} style={{ ...ink, fontSize: size }}>
    {runs.map((r, i) => (r.sup ? <tspan key={i} baselineShift="super" style={{ fontSize: size * 0.72 }}>{r.t}</tspan> : <tspan key={i}>{r.t}</tspan>))}
  </text>
);

/** Look presets: CSS variable overrides and frame styling. */
export interface Look { name: string; vars?: Record<string, string>; border?: boolean; plotBorder?: boolean; plotTint?: boolean; lineScale?: number; noGrid?: boolean; rounded?: boolean }
export const LOOKS: Look[] = [
  { name: N('Plain') },
  { name: N('Outlined'), border: true, plotBorder: true },
  { name: N('Soft panel'), plotTint: true, rounded: true },
  { name: N('Bold lines'), lineScale: 1.6, noGrid: true },
  { name: N('Dark'), border: false, vars: { '--chart-bg': '#17231d', '--chart-ink': '#eef4f0', '--chart-muted': '#a9b8af', '--chart-grid': '#2c3c33', '--chart-axis': '#5f7368', '--chart-on': '#ffffff', '--chart-hover': 'rgba(255,255,255,0.07)' } },
  { name: N('Light panel'), vars: { '--chart-bg': '#f3f8f5' }, border: true },
];
