import { useId, useRef, useState, type ReactNode } from 'react';
import type { Chart } from '../../model/workbook';
import { formatNumber, niceScale, seriesColor, valueExtent, type ResolvedChart, type Scale } from './chartData';
import { boxStats, buildTree, guessTotals, histogramBins, linePath, paretoRows, squarify, sunburstArcs, treeDepth, waterfallSteps, type BoxStats, type TreeNode } from './chartStats';
import { traits, type SeriesKind, type Traits } from './chartTypes';

const FONT = 11;
const textW = (s: string, size = FONT): number => s.length * size * 0.56;
const trunc = (s: string, maxW: number, size = FONT): string => {
  if (textW(s, size) <= maxW) return s;
  const n = Math.max(0, Math.floor(maxW / (size * 0.56)) - 1);
  return n ? `${s.slice(0, n)}…` : '';
};

const ink = { fill: 'var(--chart-ink)' } as const;
const muted = { fill: 'var(--chart-muted)' } as const;
const onFill = { fill: 'var(--chart-on)' } as const;
const GRID = { stroke: 'var(--chart-grid)', strokeWidth: 1 } as const;

interface TipRow { color: string; label: string; value: string }
interface Tip { x: number; y: number; title: string; rows: TipRow[] }
interface Box { x: number; y: number; w: number; h: number }
interface Hover { onMove: (x: number, y: number) => void; onLeave: () => void }
interface LegendEntry { label: string; color: string; line: boolean }

interface DrawCtx {
  chart: Chart; data: ResolvedChart; tr: Traits; plot: Box; thumb: boolean; compact: boolean; hot: number | null; els: ReactNode[];
  setTip: (t: Tip | null) => void; setHot: (i: number | null) => void; hover: Hover; clip: string;
  /** Element-based charts set tips from the shapes themselves. */
  elementTips: boolean;
  style: number;
}

const WF = { up: 'var(--chart-up)', down: 'var(--chart-down)', total: 'var(--chart-total)' } as const;

/** Legend entries for the chart's family. */
const legendFor = (data: ResolvedChart, tr: Traits, style: number): LegendEntry[] => {
  switch (tr.family) {
    case 'pie': return data.categories.map((c, i) => ({ label: c, color: seriesColor(i, undefined, style), line: false }));
    case 'treemap': case 'sunburst': {
      const tree = buildTree(data.paths, data.series[0]?.values ?? []);
      return tree.children.map((c, i) => ({ label: c.name, color: seriesColor(i, undefined, style), line: false }));
    }
    case 'funnel': case 'histogram': return [];
    case 'waterfall': return [{ label: 'Increase', color: WF.up, line: false }, { label: 'Decrease', color: WF.down, line: false }, { label: 'Total', color: WF.total, line: false }];
    case 'pareto': return data.series[0] ? [{ label: data.series[0].name, color: data.series[0].color, line: false }, { label: 'Cumulative %', color: seriesColor(1, undefined, style), line: true }] : [];
    default: return data.series.map((s) => ({ label: s.name, color: s.color, line: (s.type === 'line' || tr.family === 'scatter' && tr.lines) && tr.family !== 'radar' }));
  }
};

/** Draws a chart. `thumb` drops title, legend, labels and hover for small previews. */
export const ChartSvg = ({ chart, data, width, height, thumb = false }: { chart: Chart; data: ResolvedChart; width: number; height: number; thumb?: boolean }) => {
  const svg = useRef<SVGSVGElement>(null);
  const [tip, setTip] = useState<Tip | null>(null);
  const [hot, setHot] = useState<number | null>(null);
  const clip = `cc${useId().replace(/[^a-zA-Z0-9_-]/g, '')}`;
  const W = Math.max(40, width), H = Math.max(30, height);
  const compact = !thumb && (W < 220 || H < 140);
  const pad = thumb ? 4 : compact ? 6 : 10;
  const tr = traits(chart.type);
  const style = chart.options?.style ?? 0;
  const els: ReactNode[] = [];
  let top = pad, bottom = H - pad, left = pad, right = W - pad;

  const title = thumb ? '' : chart.title?.trim() ?? '';
  if (title) {
    const size = compact ? 12 : 14;
    els.push(<text key="title" x={W / 2} y={top + size - 1} textAnchor="middle" style={{ ...ink, fontSize: size, fontWeight: 600 }}>{trunc(title, W - 2 * pad, size)}</text>);
    top += size + 10;
  }

  const legendItems = legendFor(data, tr, style);
  const auto = legendItems.length > 1 || tr.family === 'pie' ? 'bottom' : 'none';
  const legend = thumb ? 'none' : chart.legend ?? (compact ? 'none' : auto);
  if (legend !== 'none' && legendItems.length) {
    const itemW = (l: string) => Math.min(160, textW(l)) + 26;
    if (legend === 'top' || legend === 'bottom') {
      const rows: LegendEntry[][] = [[]];
      for (let used = 0, k = 0; k < legendItems.length; k++) {
        const it = legendItems[k]!;
        if (used + itemW(it.label) > right - left && rows.at(-1)!.length) { if (rows.length === (compact ? 1 : 3)) break; rows.push([]); used = 0; }
        rows.at(-1)!.push(it); used += itemW(it.label);
      }
      const h = rows.length * 16 + 4;
      const y0 = legend === 'top' ? top : bottom - h + 4;
      rows.forEach((row, ri) => {
        const total = row.reduce((a, it) => a + itemW(it.label), 0);
        let x = Math.max(left, (W - total) / 2);
        row.forEach((it, i) => { els.push(<LegendItem key={`lg${ri}-${i}`} x={x} y={y0 + ri * 16 + 8} {...it} max={160} />); x += itemW(it.label); });
      });
      if (legend === 'top') top += h + 4; else bottom -= h + 4;
    } else {
      const colW = Math.min(W * 0.35, Math.max(...legendItems.map((it) => itemW(it.label))));
      const shown = legendItems.slice(0, Math.max(1, Math.floor((bottom - top) / 16)));
      const y0 = top + (bottom - top - shown.length * 16) / 2 + 8;
      const x0 = legend === 'left' ? left : right - colW;
      shown.forEach((it, i) => els.push(<LegendItem key={`lg${i}`} x={x0} y={y0 + i * 16} {...it} max={colW - 26} />));
      if (legend === 'left') left += colW + 6; else right -= colW + 6;
    }
  }

  const plot: Box = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  const hasData = data.series.some((s) => s.values.some((v) => v !== null));
  const toLocal = (cx: number, cy: number): [number, number] => {
    const r = svg.current?.getBoundingClientRect();
    return r && r.width && r.height ? [((cx - r.left) * W) / r.width, ((cy - r.top) * H) / r.height] : [0, 0];
  };
  const hover: Hover = { onMove: () => undefined, onLeave: () => { setTip(null); setHot(null); } };
  const ctx: DrawCtx = { chart, data, tr, plot, thumb, compact, hot, els, setTip, setHot, hover, clip, elementTips: false, style };

  if (!hasData) {
    els.push(<text key="empty" x={W / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: thumb ? 9 : 12 }}>{thumb || compact ? 'No data' : 'Add numbers to the data range to see a chart.'}</text>);
  } else {
    switch (tr.family) {
      case 'pie': drawPie(ctx); break;
      case 'radar': drawRadar(ctx); break;
      case 'scatter': case 'bubble': drawScatter(ctx); break;
      case 'box': drawBox(ctx); break;
      case 'treemap': drawTreemap(ctx); break;
      case 'sunburst': drawSunburst(ctx); break;
      case 'funnel': drawFunnel(ctx); break;
      default: drawCartesian(ctx, cartesianModel(ctx));
    }
  }

  return (
    <div className="chartsvg">
      <svg ref={svg} viewBox={`0 0 ${W} ${H}`} width="100%" height="100%" role="img" aria-label={title || 'Chart'}
        onPointerMove={thumb ? undefined : (e) => {
          if (ctx.elementTips) { if (!(e.target as Element).closest?.('[data-tip]')) hover.onLeave(); return; }
          const [x, y] = toLocal(e.clientX, e.clientY); ctx.hover.onMove(x, y);
        }}
        onPointerLeave={thumb ? undefined : () => hover.onLeave()}>
        <defs><clipPath id={clip}><rect x={plot.x - 2} y={plot.y - 2} width={plot.w + 4} height={plot.h + 4} /></clipPath></defs>
        {els}
      </svg>
      {tip && !thumb && (
        <div className="charttip" style={{ left: `${(tip.x / W) * 100}%`, top: `${(tip.y / H) * 100}%`, transform: `translate(${tip.x > W * 0.6 ? 'calc(-100% - 10px)' : '10px'}, ${tip.y > H * 0.6 ? 'calc(-100% - 6px)' : '6px'})` }}>
          {tip.title && <div className="tiptitle">{tip.title}</div>}
          {tip.rows.map((r, i) => <div key={i} className="tiprow"><i style={{ background: r.color }} /><span>{r.label}</span><b>{r.value}</b></div>)}
        </div>
      )}
    </div>
  );
};

const LegendItem = ({ x, y, label, color, line, max }: LegendEntry & { x: number; y: number; max: number }) => (
  <g>
    {line ? <line x1={x} x2={x + 14} y1={y} y2={y} style={{ stroke: color, strokeWidth: 2.5, strokeLinecap: 'round' }} /> : <rect x={x + 2} y={y - 5} width={10} height={10} rx={2} style={{ fill: color }} />}
    <text x={x + 18} y={y + 4} style={{ ...ink, fontSize: FONT }}>{trunc(label, max)}</text>
  </g>
);

const fmtv = (v: number, fmt: string) => formatNumber(v, fmt);

/** A nice scale, then pinned to the chart's fixed min and max when set. */
const valueScale = (lo: number, hi: number, ticks: number, fixed?: { min?: number; max?: number }): Scale => {
  const fmin = fixed?.min, fmax = fixed?.max;
  if (fmin !== undefined && Number.isFinite(fmin)) lo = fmin;
  if (fmax !== undefined && Number.isFinite(fmax)) hi = fmax;
  if (lo >= hi) hi = lo + 1;
  const s = niceScale(lo, hi, ticks);
  if (fmin === undefined && fmax === undefined) return s;
  const min = fmin ?? s.min, max = fmax ?? s.max;
  const list = s.ticks.filter((t) => t >= min - 1e-9 && t <= max + 1e-9);
  if (!list.length || Math.abs(list[0]! - min) > s.step * 0.25) list.unshift(min);
  if (Math.abs(list.at(-1)! - max) > s.step * 0.25) list.push(max);
  return { min, max, step: s.step, ticks: list };
};

// ---------- Axis frame shared by category charts ----------

interface Frame {
  px: number; py: number; pw: number; ph: number; n: number; band: number;
  valPos: (v: number) => number; valPos2: (v: number) => number; catPos: (i: number) => number; zero: number;
}

const axisFrame = (c: DrawCtx, o: { cats: string[]; horizontal: boolean; lo: number; hi: number; fmt: string; second?: { lo: number; hi: number; fmt: string; fixed?: boolean } | null; fixedPrimary?: boolean }): Frame => {
  const { chart, plot, thumb, compact, els } = c;
  const { cats, horizontal } = o;
  const n = Math.max(1, cats.length);
  const yTitle = !thumb && !compact ? chart.axisTitles?.y ?? '' : '';
  const xTitle = !thumb && !compact ? chart.axisTitles?.x ?? '' : '';
  const y2Title = !thumb && !compact && o.second ? chart.axisTitles?.y2 ?? '' : '';
  const valueLen = horizontal ? plot.w : plot.h;
  const nt = thumb ? 3 : Math.max(3, Math.min(8, Math.floor(valueLen / (horizontal ? 70 : 40))));
  const fixed = o.fixedPrimary ? undefined : { min: chart.options?.min, max: chart.options?.max };
  const scale = valueScale(o.lo, o.hi, nt, fixed);
  const scale2 = o.second ? (o.second.fixed ? { min: o.second.lo, max: o.second.hi, step: (o.second.hi - o.second.lo) / 4, ticks: [0, 1, 2, 3, 4].map((k) => o.second!.lo + (k * (o.second!.hi - o.second!.lo)) / 4) } : niceScale(o.second.lo, o.second.hi, scale.ticks.length)) : null;
  const tickText = scale.ticks.map((t) => formatNumber(t, o.fmt, true));
  const tickText2 = scale2 ? scale2.ticks.map((t) => formatNumber(t, o.second!.fmt, true)) : [];
  const maxCat = Math.max(0, ...cats.map((s) => textW(s)));

  let px = plot.x, pw = plot.w, ph = plot.h;
  const py = plot.y;
  let rotate = false, skip = 1;
  if (thumb) { /* no labels */ } else {
    if (yTitle) { px += 18; pw -= 18; }
    if (y2Title) pw -= 18;
    if (xTitle) ph -= 18;
    if (horizontal) {
      const g = Math.min(maxCat, plot.w * 0.3) + 8;
      px += g; pw -= g; ph -= 18;
      skip = Math.max(1, Math.ceil(13 / (ph / n)));
    } else {
      const g = Math.max(...tickText.map((t) => textW(t))) + 8;
      px += g; pw -= g;
      if (scale2) pw -= Math.max(...tickText2.map((t) => textW(t))) + 8;
      const band = pw / n;
      if (maxCat + 6 > band) {
        rotate = true;
        ph -= Math.min(maxCat * 0.72 + 10, plot.h * 0.3);
        skip = Math.max(1, Math.ceil(14 / band));
      } else ph -= 18;
    }
  }
  pw = Math.max(10, pw); ph = Math.max(10, ph);
  const span = scale.max - scale.min || 1;
  const band = (horizontal ? ph : pw) / n;
  const valPos = (v: number) => (horizontal ? px + ((v - scale.min) / span) * pw : py + ph - ((v - scale.min) / span) * ph);
  const valPos2 = scale2 ? (v: number) => py + ph - ((v - scale2.min) / (scale2.max - scale2.min || 1)) * ph : valPos;
  const catPos = (i: number) => (horizontal ? py + band * (i + 0.5) : px + band * (i + 0.5));
  const zero = valPos(Math.min(Math.max(0, scale.min), scale.max));
  const grid = chart.options?.gridlines !== false;

  scale.ticks.forEach((t, k) => {
    const p = valPos(t);
    if (grid) els.push(horizontal ? <line key={`g${k}`} x1={p} x2={p} y1={py} y2={py + ph} style={GRID} /> : <line key={`g${k}`} x1={px} x2={px + pw} y1={p} y2={p} style={GRID} />);
    if (!thumb) els.push(horizontal
      ? <text key={`t${k}`} x={Math.min(p, plot.x + plot.w - textW(tickText[k]!) / 2)} y={py + ph + 13} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>
      : <text key={`t${k}`} x={px - 6} y={p + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>);
  });
  if (scale2 && !thumb) scale2.ticks.forEach((t, k) => els.push(<text key={`t2${k}`} x={px + pw + 6} y={valPos2(t) + 4} textAnchor="start" style={{ ...muted, fontSize: FONT }}>{tickText2[k]}</text>));
  if (scale2) els.push(<line key="axis2" x1={px + pw} x2={px + pw} y1={py} y2={py + ph} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1 }} />);
  els.push(horizontal
    ? <line key="axis" x1={zero} x2={zero} y1={py} y2={py + ph} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1 }} />
    : <line key="axis" x1={px} x2={px + pw} y1={zero} y2={zero} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1 }} />);

  if (!thumb) cats.forEach((label, i) => {
    if (i % skip) return;
    const p = catPos(i);
    if (horizontal) els.push(<text key={`c${i}`} x={px - 6} y={p + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{trunc(label, Math.min(maxCat, plot.w * 0.3))}</text>);
    else if (rotate) els.push(<text key={`c${i}`} x={p} y={py + ph + 10} textAnchor="end" transform={`rotate(-45 ${p} ${py + ph + 10})`} style={{ ...muted, fontSize: FONT }}>{trunc(label, (plot.h * 0.3 - 10) / 0.72)}</text>);
    else els.push(<text key={`c${i}`} x={p} y={py + ph + 14} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{label}</text>);
  });
  if (yTitle) els.push(<text key="yt" x={plot.x + 10} y={py + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(yTitle, ph)}</text>);
  if (y2Title) els.push(<text key="y2t" x={plot.x + plot.w - 8} y={py + ph / 2} textAnchor="middle" transform={`rotate(90 ${plot.x + plot.w - 8} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(y2Title, ph)}</text>);
  if (xTitle) els.push(<text key="xt" x={px + pw / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(xTitle, pw)}</text>);

  if (c.hot !== null && !thumb && c.hot < n) {
    const b = c.hot;
    els.push(horizontal
      ? <rect key="hot" x={px} y={py + b * band} width={pw} height={band} style={{ fill: 'var(--chart-hover)' }} />
      : <rect key="hot" x={px + b * band} y={py} width={band} height={ph} style={{ fill: 'var(--chart-hover)' }} />);
  }
  return { px, py, pw, ph, n, band, valPos, valPos2, catPos, zero };
};

/** Category under the pointer, or null outside the plot. */
const bandAt = (f: Frame, horizontal: boolean, x: number, y: number): number | null => {
  if (x < f.px || x > f.px + f.pw || y < f.py || y > f.py + f.ph) return null;
  return Math.min(f.n - 1, Math.max(0, Math.floor(((horizontal ? y - f.py : x - f.px) / (horizontal ? f.ph : f.pw)) * f.n)));
};

// ---------- Column, bar, line, area and their stacks; combos, histogram, Pareto, waterfall ----------

interface Seg { i: number; v0: number; v1: number; raw: number | null }
interface CSeries { name: string; color: string; kind: SeriesKind; secondary: boolean; segs: Seg[]; colors?: string[]; markers: boolean; smooth: boolean; stacked: boolean; label?: (g: Seg) => string }
interface CModel {
  cats: string[]; series: CSeries[]; horizontal: boolean; fmt: string; fmt2: string;
  lo: number; hi: number; second: { lo: number; hi: number } | null; fixed2?: boolean; fixedPrimary?: boolean;
  gap: number; connectors?: boolean; tip: (i: number) => TipRow[];
}

const stackSegs = (vals: (number | null)[][], n: number, mode: 'none' | 'stacked' | 'percent'): Seg[][] => {
  const out: Seg[][] = vals.map(() => []);
  for (let i = 0; i < n; i++) {
    let pos = 0, neg = 0;
    const total = mode === 'percent' ? vals.reduce((a, v) => a + Math.abs(v[i] ?? 0), 0) : 1;
    vals.forEach((v, si) => {
      const raw = v[i] ?? null;
      let x = raw ?? 0;
      if (mode === 'percent') x = total ? x / total : 0;
      if (mode === 'none') out[si]!.push({ i, v0: 0, v1: x, raw });
      else if (x >= 0) { out[si]!.push({ i, v0: pos, v1: pos + x, raw }); pos += x; }
      else { out[si]!.push({ i, v0: neg, v1: neg + x, raw }); neg += x; }
    });
  }
  return out;
};

const cartesianModel = (c: DrawCtx): CModel => {
  const { chart, data, tr, style } = c;
  const fmt = data.fmt;
  if (tr.family === 'histogram') {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
    const bins = histogramBins(all, { binCount: chart.options?.binCount, binWidth: chart.options?.binWidth }, (v) => formatNumber(v, fmt, true));
    const cats = bins.map((b) => b.label);
    const inBin = (v: number, bi: number) => (v > bins[bi]!.lo || (bi === 0 && v >= bins[bi]!.lo)) && (v <= bins[bi]!.hi || bi === bins.length - 1);
    const series: CSeries[] = data.series.map((s) => {
      const own = bins.map((_, bi) => s.values.filter((v): v is number => v !== null && inBin(v, bi)).length);
      return { name: s.name, color: s.color, kind: 'column', secondary: false, segs: own.map((v, i) => ({ i, v0: 0, v1: v, raw: v })), markers: false, smooth: false, stacked: false };
    });
    if (data.series.length === 1) series[0]!.segs = bins.map((b, i) => ({ i, v0: 0, v1: b.count, raw: b.count }));
    const hi = Math.max(1, ...series.flatMap((s) => s.segs.map((g) => g.v1)));
    return { cats, series, horizontal: false, fmt: '0', fmt2: '0', lo: 0, hi, second: null, gap: series.length === 1 ? 0.97 : 0.86, tip: (i) => series.map((s) => ({ color: s.color, label: s.name, value: String(s.segs[i]?.raw ?? 0) })) };
  }
  if (tr.family === 'pareto') {
    const s = data.series[0]!;
    const rows = paretoRows(data.categories, s.values);
    const lineColor = seriesColor(1, undefined, style);
    const bars: CSeries = { name: s.name, color: s.color, kind: 'column', secondary: false, segs: rows.map((r, i) => ({ i, v0: 0, v1: r.value, raw: r.value })), markers: false, smooth: false, stacked: false };
    const line: CSeries = { name: 'Cumulative %', color: lineColor, kind: 'line', secondary: true, segs: rows.map((r, i) => ({ i, v0: 0, v1: r.cumulative, raw: r.cumulative })), markers: true, smooth: false, stacked: false, label: (g) => formatNumber(g.v1, '0%') };
    return {
      cats: rows.map((r) => r.label), series: [bars, line], horizontal: false, fmt, fmt2: '0%', lo: 0, hi: Math.max(0, ...rows.map((r) => r.value)), second: { lo: 0, hi: 1 }, fixed2: true, gap: 0.8,
      tip: (i) => [{ color: s.color, label: s.name, value: fmtv(rows[i]?.value ?? 0, fmt) }, { color: lineColor, label: 'Cumulative', value: formatNumber(rows[i]?.cumulative ?? 0, '0.0%') }],
    };
  }
  if (tr.family === 'waterfall') {
    const s = data.series[0]!;
    const totals = chart.options?.totals ?? guessTotals(data.categories);
    const steps = waterfallSteps(s.values, totals);
    const colors = steps.map((st) => WF[st.kind]);
    const ser: CSeries = { name: s.name, color: WF.up, kind: 'column', secondary: false, segs: steps.map((st) => ({ i: st.index, v0: st.start, v1: st.end, raw: st.value })), colors, markers: false, smooth: false, stacked: true, label: (g) => fmtv(g.raw ?? 0, fmt) };
    const vals = steps.flatMap((st) => [st.start, st.end]);
    return {
      cats: data.categories, series: [ser], horizontal: false, fmt, fmt2: fmt, lo: Math.min(0, ...vals), hi: Math.max(0, ...vals), second: null, gap: 0.7, connectors: true,
      tip: (i) => {
        const st = steps[i];
        if (!st) return [];
        return st.kind === 'total' ? [{ color: WF.total, label: 'Total', value: fmtv(st.end, fmt) }]
          : [{ color: WF[st.kind], label: st.kind === 'up' ? 'Increase' : 'Decrease', value: fmtv(st.value, fmt) }, { color: 'var(--chart-axis)', label: 'Running total', value: fmtv(st.end, fmt) }];
      },
    };
  }
  // Plain, stacked and combo category charts. Each kind of series stacks within its own group.
  const n = Math.max(1, data.categories.length);
  const kinds = data.series.map((s) => s.type);
  const modeFor = (k: SeriesKind): 'none' | 'stacked' | 'percent' => (tr.combo ? (chart.type === 'comboAreaColumn' && k === 'area' ? 'stacked' : 'none') : tr.stack);
  const segs: Seg[][] = data.series.map(() => []);
  for (const axis of [false, true]) for (const k of ['column', 'line', 'area'] as SeriesKind[]) {
    const idx = data.series.flatMap((s, i) => (kinds[i] === k && s.secondary === axis ? [i] : []));
    if (!idx.length) continue;
    const st = stackSegs(idx.map((i) => data.series[i]!.values), n, modeFor(k));
    idx.forEach((i, j) => { segs[i] = st[j]!; });
  }
  const percent = tr.stack === 'percent';
  const series: CSeries[] = data.series.map((s, i) => ({
    name: s.name, color: s.color, kind: s.type, secondary: s.secondary, segs: segs[i]!, stacked: modeFor(s.type) !== 'none',
    markers: tr.markers || (tr.combo && s.type === 'line' && n <= 40), smooth: false,
    label: percent ? (g) => formatNumber(g.v1 - g.v0, '0%') : undefined,
  }));
  const ext = (sec: boolean): [number, number] | null => {
    const list = series.filter((s) => s.secondary === sec);
    if (!list.length) return null;
    const vals = list.flatMap((s) => s.segs.filter((g) => g.raw !== null || s.stacked).flatMap((g) => (s.stacked ? [g.v0, g.v1] : [g.v1])));
    if (percent) return [Math.min(0, ...vals), Math.max(0, ...vals)];
    return valueExtent(vals, list.some((s) => s.kind !== 'line'));
  };
  const p = ext(false) ?? [0, 1], s2 = ext(true);
  return {
    cats: data.categories, series, horizontal: tr.horizontal, fmt: percent ? '0%' : fmt, fmt2: fmt, lo: p[0], hi: p[1], second: s2 ? { lo: s2[0], hi: s2[1] } : null,
    gap: tr.stack !== 'none' ? 0.62 : 0.74, fixedPrimary: false,
    tip: (i) => data.series.map((s, si) => {
      const g = segs[si]![i];
      const v = g?.raw ?? null;
      return { color: s.color, label: s.name, value: v === null ? 'Blank' : percent ? `${fmtv(v, data.fmt)} (${formatNumber(g!.v1 - g!.v0, '0.0%')})` : fmtv(v, data.fmt) };
    }),
  };
};

const drawCartesian = (c: DrawCtx, m: CModel): void => {
  const { chart, thumb, els } = c;
  const { horizontal } = m;
  const f = axisFrame(c, { cats: m.cats, horizontal, lo: m.lo, hi: m.hi, fmt: m.fmt, second: m.second ? { ...m.second, fmt: m.fmt2, fixed: m.fixed2 } : null, fixedPrimary: m.fixedPrimary });
  const { band, catPos } = f;
  const pos = (s: CSeries, v: number) => (s.secondary ? f.valPos2(v) : f.valPos(v));
  const labels = !thumb && !c.compact && chart.dataLabels && f.n * m.series.length <= 240;
  const parts: ReactNode[] = [];

  // Areas first, so bars and lines sit on top.
  m.series.forEach((s, si) => {
    if (s.kind !== 'area') return;
    const usable = s.stacked ? s.segs : s.segs.filter((g) => g.raw !== null);
    if (!usable.length) return;
    const xy = (g: Seg, v: number) => `${catPos(g.i)},${pos(s, v)}`;
    const upper = usable.map((g) => xy(g, g.v1));
    const base = s.secondary ? f.valPos2(0) : f.zero;
    const lower = [...usable].reverse().map((g) => (s.stacked ? xy(g, g.v0) : `${catPos(g.i)},${base}`));
    parts.push(<polygon key={`a${si}`} points={[...upper, ...lower].join(' ')} style={{ fill: s.color, fillOpacity: s.stacked ? 0.82 : 0.38 }} />);
    parts.push(<polyline key={`al${si}`} points={upper.join(' ')} style={{ fill: 'none', stroke: s.color, strokeWidth: 1.6, strokeLinejoin: 'round' }} />);
  });

  const bars = m.series.map((s, si) => ({ s, si })).filter((x) => x.s.kind === 'column');
  const stackedBars = bars.length > 0 && bars.every((b) => b.s.stacked);
  const group = band * m.gap;
  const slots = stackedBars ? 1 : bars.length;
  bars.forEach(({ s, si }, j) => {
    const w = group / Math.max(1, slots);
    s.segs.forEach((g) => {
      if (g.raw === null && !s.stacked) return;
      const off = -group / 2 + (stackedBars ? 0 : j) * w;
      const a = pos(s, g.v0), b = pos(s, g.v1);
      const cp = catPos(g.i) + off;
      const gapPx = w > 6 ? 1 : 0.4;
      const r = horizontal ? { x: Math.min(a, b), y: cp, width: Math.abs(b - a), height: Math.max(0.5, w - gapPx) } : { x: cp, y: Math.min(a, b), width: Math.max(0.5, w - gapPx), height: Math.abs(b - a) };
      parts.push(<rect key={`b${si}-${g.i}`} {...r} rx={Math.min(2, w / 6)} style={{ fill: s.colors?.[g.i] ?? s.color }} />);
      if (!labels || g.raw === null) return;
      const t = s.label ? s.label(g) : fmtv(g.raw, m.fmt);
      const inside = s.stacked && !m.connectors;
      if (inside) {
        if ((horizontal ? r.width : r.height) > 14 && textW(t, 10) < (horizontal ? r.width : w)) els.push(<text className="onfill" key={`l${si}-${g.i}`} x={r.x + r.width / 2} y={r.y + r.height / 2 + 3.5} textAnchor="middle" style={{ ...onFill, fontSize: 10 }}>{t}</text>);
      } else if (horizontal) els.push(<text key={`l${si}-${g.i}`} x={g.v1 >= g.v0 ? Math.max(a, b) + 4 : Math.min(a, b) - 4} y={cp + w / 2 + 3.5} textAnchor={g.v1 >= g.v0 ? 'start' : 'end'} style={{ ...ink, fontSize: 10 }}>{t}</text>);
      else els.push(<text key={`l${si}-${g.i}`} x={cp + w / 2} y={g.v1 >= g.v0 ? Math.min(a, b) - 4 : Math.max(a, b) + 12} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{t}</text>);
    });
    if (m.connectors) s.segs.forEach((g, k) => {
      const next = s.segs[k + 1];
      if (!next) return;
      const y = pos(s, g.v1);
      parts.push(<line key={`cn${k}`} x1={catPos(g.i) + group / 2} x2={catPos(next.i) - group / 2} y1={y} y2={y} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1, strokeDasharray: '3 2' }} />);
    });
  });

  m.series.forEach((s, si) => {
    if (s.kind !== 'line') return;
    const runs: [number, number][][] = [[]];
    s.segs.forEach((g) => { if (g.raw === null && !s.stacked) { if (runs.at(-1)!.length) runs.push([]); } else runs.at(-1)!.push(horizontal ? [pos(s, g.v1), catPos(g.i)] : [catPos(g.i), pos(s, g.v1)]); });
    runs.forEach((r, k) => r.length > 1 && parts.push(<path key={`ln${si}-${k}`} d={linePath(r, s.smooth)} style={{ fill: 'none', stroke: s.color, strokeWidth: thumb ? 1.6 : 2.2, strokeLinejoin: 'round', strokeLinecap: 'round' }} />));
    s.segs.forEach((g) => {
      if (g.raw === null && !s.stacked) return;
      const single = runs.some((r) => r.length === 1);
      const show = s.markers || single || c.hot === g.i;
      if (!show) return;
      const [x, y] = horizontal ? [pos(s, g.v1), catPos(g.i)] : [catPos(g.i), pos(s, g.v1)];
      parts.push(<circle key={`m${si}-${g.i}`} cx={x} cy={y} r={thumb ? 1.6 : c.hot === g.i ? 4.5 : c.compact ? 2 : 3} style={{ fill: 'var(--chart-bg)', stroke: s.color, strokeWidth: 2 }} />);
    });
    if (labels) s.segs.forEach((g) => g.raw !== null && els.push(<text key={`l${si}-${g.i}`} x={catPos(g.i)} y={pos(s, g.v1) - 7} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{s.label ? s.label(g) : fmtv(g.raw, s.secondary ? m.fmt2 : m.fmt)}</text>));
  });
  els.splice(els.length, 0, <g key="marks" clipPath={`url(#${c.clip})`}>{parts}</g>);

  if (thumb) return;
  c.hover.onMove = (x, y) => {
    const i = bandAt(f, horizontal, x, y);
    if (i === null) { c.setTip(null); c.setHot(null); return; }
    c.setHot(i);
    c.setTip({ x, y, title: m.cats[i] ?? '', rows: m.tip(i) });
  };
};

// ---------- Box and whisker ----------

const drawBox = (c: DrawCtx): void => {
  const { chart, data, thumb, els } = c;
  const n = Math.max(0, ...data.series.map((s) => s.values.length));
  const labels = data.categories.slice(0, n);
  const unique = [...new Set(labels)];
  // Repeated labels group the rows; otherwise each series is one box.
  const grouped = chart.series.some((s) => s.categories) && unique.length < labels.length && unique.length <= 50;
  const cats = grouped ? unique : data.series.map((s) => s.name);
  const stats: (BoxStats | null)[][] = grouped
    ? unique.map((u) => data.series.map((s) => boxStats(s.values.filter((v, i): v is number => v !== null && labels[i] === u))))
    : data.series.map((s) => [boxStats(s.values.filter((v): v is number => v !== null))]);
  const all = stats.flat().filter((s): s is BoxStats => !!s);
  if (!all.length) return;
  const lo = Math.min(...all.map((s) => s.min)), hi = Math.max(...all.map((s) => s.max));
  const f = axisFrame(c, { cats, horizontal: false, lo: lo >= 0 && lo < hi * 0.4 ? 0 : lo, hi, fmt: data.fmt });
  const per = grouped ? data.series.length : 1;
  const group = f.band * 0.7, w = group / per;
  const showMean = !chart.options?.hideMean, showOut = !chart.options?.hideOutliers;
  const parts: ReactNode[] = [];
  stats.forEach((row, ci) => row.forEach((st, si) => {
    if (!st) return;
    const color = grouped ? data.series[si]!.color : data.series[ci]!.color;
    const x0 = f.catPos(ci) - group / 2 + si * w + w * 0.12, bw = w * 0.76, xm = x0 + bw / 2;
    const Y = f.valPos;
    const k = `${ci}-${si}`;
    parts.push(<line key={`w${k}`} x1={xm} x2={xm} y1={Y(st.lowWhisker)} y2={Y(st.highWhisker)} style={{ stroke: color, strokeWidth: 1.4 }} />);
    parts.push(<line key={`wl${k}`} x1={xm - bw / 4} x2={xm + bw / 4} y1={Y(st.lowWhisker)} y2={Y(st.lowWhisker)} style={{ stroke: color, strokeWidth: 1.4 }} />);
    parts.push(<line key={`wh${k}`} x1={xm - bw / 4} x2={xm + bw / 4} y1={Y(st.highWhisker)} y2={Y(st.highWhisker)} style={{ stroke: color, strokeWidth: 1.4 }} />);
    parts.push(<rect key={`bx${k}`} x={x0} y={Y(st.q3)} width={Math.max(1, bw)} height={Math.max(1, Y(st.q1) - Y(st.q3))} rx={2} style={{ fill: color, fillOpacity: 0.32, stroke: color, strokeWidth: 1.4 }} />);
    parts.push(<line key={`md${k}`} x1={x0} x2={x0 + bw} y1={Y(st.median)} y2={Y(st.median)} style={{ stroke: color, strokeWidth: 2.2 }} />);
    if (showMean) {
      const s = thumb ? 2 : 4, y = Y(st.mean);
      parts.push(<path key={`mn${k}`} d={`M${xm - s},${y - s} L${xm + s},${y + s} M${xm - s},${y + s} L${xm + s},${y - s}`} style={{ stroke: 'var(--chart-ink)', strokeWidth: 1.6 }} />);
    }
    if (showOut) st.outliers.forEach((v, oi) => parts.push(<circle key={`o${k}-${oi}`} cx={xm} cy={Y(v)} r={thumb ? 1.4 : 3} style={{ fill: 'var(--chart-bg)', stroke: color, strokeWidth: 1.5 }} />));
    if (!thumb && !c.compact && chart.dataLabels) els.push(<text key={`bl${k}`} x={x0 + bw + 3} y={Y(st.median) + 3.5} style={{ ...ink, fontSize: 10 }}>{fmtv(st.median, data.fmt)}</text>);
  }));
  els.push(<g key="marks" clipPath={`url(#${c.clip})`}>{parts}</g>);
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    const ci = bandAt(f, false, x, y);
    if (ci === null) { c.setTip(null); c.setHot(null); return; }
    const si = Math.min(per - 1, Math.max(0, Math.floor((x - (f.catPos(ci) - group / 2)) / w)));
    const st = stats[ci]?.[si];
    c.setHot(ci);
    if (!st) { c.setTip(null); return; }
    const color = grouped ? data.series[si]!.color : data.series[ci]!.color;
    const v = (l: string, n: number) => ({ color, label: l, value: fmtv(n, data.fmt) });
    c.setTip({ x, y, title: grouped ? `${cats[ci]} · ${data.series[si]!.name}` : cats[ci] ?? '', rows: [v('Maximum', st.max), v('Upper quartile', st.q3), v('Median', st.median), v('Lower quartile', st.q1), v('Minimum', st.min), v('Mean', st.mean), ...(st.outliers.length ? [{ color, label: 'Outliers', value: String(st.outliers.length) }] : [])] });
  };
};

// ---------- Scatter and bubble ----------

const drawScatter = (c: DrawCtx): void => {
  const { chart, data, tr, plot, thumb, compact, els } = c;
  const bubble = tr.family === 'bubble';
  type P = { si: number; i: number; x: number; y: number; z: number };
  const pts: P[] = data.series.flatMap((s, si) => s.values.map((y, i) => ({ si, i, x: data.xs ? data.xs[i] ?? null : i + 1, y, z: bubble ? s.sizes?.[i] ?? 1 : 0 }))
    .filter((p): p is P => p.x !== null && p.y !== null && (!bubble || p.z !== null && p.z > 0)));
  if (!pts.length) return;
  const [xl, xh] = valueExtent(pts.map((p) => p.x), false);
  const [yl, yh] = valueExtent(pts.map((p) => p.y), false);
  let px = plot.x, pw = plot.w, ph = plot.h;
  const yTitle = !thumb && !compact ? chart.axisTitles?.y ?? '' : '', xTitle = !thumb && !compact ? chart.axisTitles?.x ?? '' : '';
  let ys = valueScale(yl, yh, thumb ? 3 : Math.max(2, Math.min(8, Math.floor(plot.h / 40))), { min: chart.options?.min, max: chart.options?.max });
  const yText = ys.ticks.map((t) => formatNumber(t, data.fmt, true));
  if (!thumb) {
    const g = Math.max(...yText.map((t) => textW(t))) + 8 + (yTitle ? 18 : 0);
    px += g; pw -= g; ph -= 18 + (xTitle ? 18 : 0);
  }
  const zMax = Math.max(...pts.map((p) => p.z), 1e-9);
  const rMax = bubble ? Math.max(3, Math.min(pw, ph) * (thumb ? 0.08 : 0.09)) : 0;
  const rOf = (z: number) => Math.max(thumb ? 1.2 : 2.5, Math.sqrt(z / zMax) * rMax);
  const nx = thumb ? 3 : Math.max(3, Math.floor(pw / 70));
  let xs = niceScale(xl, xh, nx);
  if (bubble) {
    // Leave room for the largest bubbles inside the plot.
    const padX = (rMax / Math.max(10, pw - 2 * rMax)) * (xh - xl || 1), padY = (rMax / Math.max(10, ph - 2 * rMax)) * (yh - yl || 1);
    ys = valueScale(yl - padY, yh + padY, thumb ? 3 : Math.max(2, Math.min(8, Math.floor(plot.h / 40))), { min: chart.options?.min, max: chart.options?.max });
    xs = niceScale(xl - padX, xh + padX, nx);
  }
  const yTicks = ys.ticks.map((t) => formatNumber(t, data.fmt, true));
  const X = (v: number) => px + ((v - xs.min) / (xs.max - xs.min || 1)) * pw;
  const Y = (v: number) => plot.y + ph - ((v - ys.min) / (ys.max - ys.min || 1)) * ph;
  const grid = chart.options?.gridlines !== false;
  ys.ticks.forEach((t, k) => {
    if (grid) els.push(<line key={`gy${k}`} x1={px} x2={px + pw} y1={Y(t)} y2={Y(t)} style={GRID} />);
    if (!thumb) els.push(<text key={`ty${k}`} x={px - 6} y={Y(t) + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{yTicks[k]}</text>);
  });
  xs.ticks.forEach((t, k) => {
    if (grid) els.push(<line key={`gx${k}`} x1={X(t)} x2={X(t)} y1={plot.y} y2={plot.y + ph} style={GRID} />);
    if (!thumb) els.push(<text key={`tx${k}`} x={Math.min(X(t), plot.x + plot.w - textW(formatNumber(t, 'General', true)) / 2)} y={plot.y + ph + 14} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{formatNumber(t, 'General', true)}</text>);
  });
  els.push(<line key="ax" x1={px} x2={px + pw} y1={plot.y + ph} y2={plot.y + ph} style={{ stroke: 'var(--chart-axis)' }} />);
  if (yTitle) els.push(<text key="yt" x={plot.x + 10} y={plot.y + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${plot.y + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(yTitle, ph)}</text>);
  if (xTitle) els.push(<text key="xt" x={px + pw / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(xTitle, pw)}</text>);
  const parts: ReactNode[] = [];
  if (tr.lines) data.series.forEach((s, si) => {
    const line = pts.filter((p) => p.si === si).sort((a, b) => a.x - b.x || a.i - b.i).map((p): [number, number] => [X(p.x), Y(p.y)]);
    if (line.length > 1) parts.push(<path key={`sl${si}`} d={linePath(line, tr.smooth)} style={{ fill: 'none', stroke: s.color, strokeWidth: thumb ? 1.4 : 2, strokeLinejoin: 'round', strokeLinecap: 'round' }} />);
  });
  const order = bubble ? pts.map((p, k) => ({ p, k })).sort((a, b) => b.p.z - a.p.z) : pts.map((p, k) => ({ p, k }));
  order.forEach(({ p, k }) => {
    const color = data.series[p.si]!.color;
    const on = c.hot === k;
    if (bubble) parts.push(<circle key={`p${k}`} cx={X(p.x)} cy={Y(p.y)} r={rOf(p.z)} style={{ fill: color, fillOpacity: on ? 0.9 : 0.6, stroke: color, strokeWidth: 1 }} />);
    else if (tr.markers || on) parts.push(<circle key={`p${k}`} cx={X(p.x)} cy={Y(p.y)} r={thumb ? 1.8 : on ? 5.5 : compact ? 2.4 : 4} style={{ fill: tr.lines ? 'var(--chart-bg)' : color, fillOpacity: tr.lines ? 1 : 0.85, stroke: tr.lines ? color : 'var(--chart-bg)', strokeWidth: tr.lines ? 2 : 1 }} />);
    if (!thumb && !compact && chart.dataLabels && pts.length <= 200) els.push(<text key={`pl${k}`} x={X(p.x)} y={Y(p.y) - (bubble ? 0 : 7)} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{formatNumber(bubble ? p.z : p.y, data.fmt)}</text>);
  });
  els.push(<g key="marks" clipPath={`url(#${c.clip})`}>{parts}</g>);
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    let best = -1, bd = Infinity;
    pts.forEach((p, k) => {
      const d = Math.sqrt((X(p.x) - x) ** 2 + (Y(p.y) - y) ** 2) - (bubble ? rOf(p.z) : 0);
      if (d < bd && d < 18) { bd = d; best = k; }
    });
    if (best < 0) { c.setTip(null); c.setHot(null); return; }
    const p = pts[best]!;
    const s = data.series[p.si]!;
    c.setHot(best);
    const rows: TipRow[] = [{ color: s.color, label: chart.axisTitles?.x || 'X', value: formatNumber(p.x, 'General') }, { color: s.color, label: chart.axisTitles?.y || 'Y', value: formatNumber(p.y, data.fmt) }];
    if (bubble) rows.push({ color: s.color, label: 'Size', value: formatNumber(p.z, 'General') });
    c.setTip({ x: X(p.x), y: Y(p.y), title: data.xs ? s.name : `${s.name} · ${data.categories[p.i] ?? ''}`, rows });
  };
};

// ---------- Pie and doughnut ----------

const labelText = (mode: 'value' | 'percent' | 'both', v: number, share: number, fmt: string): string =>
  mode === 'value' ? formatNumber(v, fmt, true) : mode === 'percent' ? formatNumber(share, '0%') : `${formatNumber(v, fmt, true)} (${formatNumber(share, '0%')})`;

const arcPath = (cx: number, cy: number, r: number, inner: number, a0: number, a1: number): string => {
  if (a1 - a0 >= Math.PI * 2 - 1e-6) a1 = a0 + Math.PI * 2 - 1e-4;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const p = (rad: number, a: number) => `${(cx + Math.cos(a) * rad).toFixed(2)},${(cy + Math.sin(a) * rad).toFixed(2)}`;
  if (!inner) return `M${cx},${cy} L${p(r, a0)} A${r},${r} 0 ${large} 1 ${p(r, a1)} Z`;
  return `M${p(r, a0)} A${r},${r} 0 ${large} 1 ${p(r, a1)} L${p(inner, a1)} A${inner},${inner} 0 ${large} 0 ${p(inner, a0)} Z`;
};

const drawPie = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, style } = c;
  c.elementTips = true;
  const doughnut = chart.type === 'doughnut';
  const rings = doughnut ? data.series : data.series.slice(0, 1);
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const R = Math.max(4, Math.min(plot.w, plot.h) / 2 - (thumb ? 0 : 6));
  const hole = doughnut ? R * 0.52 : 0;
  const ringW = (R - hole) / rings.length;
  const mode = chart.options?.labelMode ?? 'percent';
  let grand = 0;
  rings.forEach((s, ri) => {
    const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (ri === 0) grand = total;
    if (!total) return;
    const outer = R - ri * ringW, inner = doughnut ? outer - ringW : 0;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const start = a0;
      a0 = a1;
      if (!v) return;
      const mid = (start + a1) / 2;
      const hot = c.hot === ri * 10000 + i;
      const off = hot && rings.length === 1 ? 5 : 0;
      const ox = Math.cos(mid) * off, oy = Math.sin(mid) * off;
      const color = seriesColor(i, undefined, style);
      els.push(<path key={`s${ri}-${i}`} data-tip="" d={arcPath(cx + ox, cy + oy, outer, inner, start, a1)} style={{ fill: color, fillOpacity: hot && rings.length > 1 ? 0.8 : 1, stroke: 'var(--chart-bg)', strokeWidth: thumb ? 0.6 : 1.5 }}
        onPointerEnter={thumb ? undefined : () => { c.setHot(ri * 10000 + i); c.setTip({ x: cx + Math.cos(mid) * (outer - ringW / 2), y: cy + Math.sin(mid) * (outer - ringW / 2), title: data.categories[i] ?? '', rows: [{ color, label: s.name, value: `${fmtv(v, data.fmt)} (${formatNumber(v / total, '0.0%')})` }] }); }} />);
      const t = labelText(mode, v, v / total, data.fmt);
      const lr = inner ? (outer + inner) / 2 : outer * 0.64;
      const arcLen = (a1 - start) * lr;
      if (!thumb && !compact && chart.dataLabels && arcLen > textW(t, 11) * 0.7 && (inner ? outer - inner > 14 : true)) {
        els.push(<text className="onfill" key={`sl${ri}-${i}`} x={cx + ox + Math.cos(mid) * lr} y={cy + oy + Math.sin(mid) * lr + 4} textAnchor="middle" style={{ ...onFill, fontSize: 11, fontWeight: 600, pointerEvents: 'none' }}>{t}</text>);
      }
    });
  });
  if (!grand) { els.push(<text key="none" x={cx} y={cy} textAnchor="middle" style={{ ...muted, fontSize: 11 }}>No positive values</text>); return; }
  if (hole && !thumb && hole > 22) els.push(<text key="tot" x={cx} y={cy + 5} textAnchor="middle" style={{ ...ink, fontSize: Math.min(16, hole / 2.6), fontWeight: 600, pointerEvents: 'none' }}>{formatNumber(grand, data.fmt, true)}</text>);
};

// ---------- Radar ----------

const drawRadar = (c: DrawCtx): void => {
  const { data, tr, plot, thumb, compact, els, chart } = c;
  const n = Math.max(3, data.categories.length);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sc = valueScale(Math.min(0, ...vals), Math.max(...vals), thumb ? 3 : 5, { min: chart.options?.min, max: chart.options?.max });
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const labelRoom = thumb ? 2 : compact ? 10 : Math.min(60, Math.max(22, Math.max(0, ...data.categories.map((t) => textW(t))) * 0.6));
  const R = Math.max(6, Math.min(plot.w / 2 - labelRoom, plot.h / 2 - (thumb ? 2 : 16)));
  const ang = (i: number) => -Math.PI / 2 + (i / n) * Math.PI * 2;
  const rad = (v: number) => Math.max(0, ((v - sc.min) / (sc.max - sc.min || 1))) * R;
  const pt = (i: number, v: number): [number, number] => [cx + Math.cos(ang(i)) * rad(v), cy + Math.sin(ang(i)) * rad(v)];
  const grid = chart.options?.gridlines !== false;
  sc.ticks.forEach((t, k) => {
    if (t === sc.min) return;
    if (grid) els.push(<polygon key={`r${k}`} points={Array.from({ length: n }, (_, i) => pt(i, t).join(',')).join(' ')} style={{ fill: 'none', ...GRID }} />);
    if (!thumb && !compact) els.push(<text key={`rt${k}`} x={cx + 3} y={cy - rad(t) + 10} style={{ ...muted, fontSize: 9 }}>{formatNumber(t, data.fmt, true)}</text>);
  });
  for (let i = 0; i < n; i++) {
    els.push(<line key={`sp${i}`} x1={cx} y1={cy} x2={cx + Math.cos(ang(i)) * R} y2={cy + Math.sin(ang(i)) * R} style={{ stroke: c.hot === i ? 'var(--chart-axis)' : 'var(--chart-grid)' }} />);
    if (!thumb && !compact) {
      const a = ang(i), lx = cx + Math.cos(a) * (R + 8), ly = cy + Math.sin(a) * (R + 8) + 4 + (Math.sin(a) > 0.9 ? 6 : 0);
      els.push(<text key={`cl${i}`} x={lx} y={ly} textAnchor={Math.abs(Math.cos(a)) < 0.2 ? 'middle' : Math.cos(a) > 0 ? 'start' : 'end'} style={{ ...muted, fontSize: FONT }}>{trunc(data.categories[i] ?? '', labelRoom / 0.6)}</text>);
    }
  }
  data.series.forEach((s, si) => {
    const pts = Array.from({ length: n }, (_, i) => pt(i, s.values[i] ?? Math.max(0, sc.min)));
    els.push(<polygon key={`rs${si}`} points={pts.map((p) => p.join(',')).join(' ')} style={{ fill: s.color, fillOpacity: tr.filled ? 0.45 : 0.08, stroke: s.color, strokeWidth: thumb ? 1.2 : 2, strokeLinejoin: 'round' }} />);
    if (tr.markers || (c.hot !== null && !thumb)) pts.forEach((p, i) => (tr.markers || c.hot === i) && els.push(<circle key={`rm${si}-${i}`} cx={p[0]} cy={p[1]} r={thumb ? 1.4 : c.hot === i ? 4.5 : compact ? 2 : 3} style={{ fill: 'var(--chart-bg)', stroke: s.color, strokeWidth: 2 }} />));
    if (!thumb && !compact && chart.dataLabels) s.values.forEach((v, i) => v !== null && els.push(<text key={`rl${si}-${i}`} x={pt(i, v)[0]} y={pt(i, v)[1] - 6} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{fmtv(v, data.fmt)}</text>));
  });
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    if ((x - cx) ** 2 + (y - cy) ** 2 > (R + 16) ** 2) { c.setTip(null); c.setHot(null); return; }
    let a = Math.atan2(y - cy, x - cx) + Math.PI / 2;
    if (a < 0) a += Math.PI * 2;
    const i = Math.round((a / (Math.PI * 2)) * n) % n;
    c.setHot(i);
    c.setTip({ x, y, title: data.categories[i] ?? '', rows: data.series.map((s) => ({ color: s.color, label: s.name, value: s.values[i] === null || s.values[i] === undefined ? 'Blank' : fmtv(s.values[i]!, data.fmt) })) });
  };
};

// ---------- Funnel ----------

const drawFunnel = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, style } = c;
  c.elementTips = true;
  const s = data.series[0]!;
  const rows = s.values.map((v, i) => ({ v: v !== null && v > 0 ? v : 0, i, label: data.categories[i] ?? '' })).filter((r) => r.v > 0);
  if (!rows.length) return;
  const max = Math.max(...rows.map((r) => r.v));
  const first = rows[0]!.v;
  const labelW = thumb || compact ? 0 : Math.min(plot.w * 0.3, Math.max(...rows.map((r) => textW(r.label))) + 10);
  const x0 = plot.x + labelW, w = plot.w - labelW;
  const h = plot.h / rows.length;
  const mode = chart.options?.labelMode ?? 'value';
  rows.forEach((r, k) => {
    const bw = Math.max(2, (r.v / max) * w);
    const x = x0 + (w - bw) / 2, y = plot.y + k * h;
    const color = chart.series[0]?.color ?? seriesColor(style ? k : 0, undefined, style);
    const hot = c.hot === k;
    els.push(<rect key={`f${k}`} data-tip="" x={x} y={y + h * 0.06} width={bw} height={Math.max(1, h * 0.88)} rx={2} style={{ fill: color, fillOpacity: hot ? 0.82 : 1 }}
      onPointerEnter={thumb ? undefined : () => { c.setHot(k); c.setTip({ x: x0 + w / 2, y: y + h / 2, title: r.label, rows: [{ color, label: s.name, value: fmtv(r.v, data.fmt) }, { color, label: 'Of first', value: formatNumber(r.v / first, '0.0%') }] }); }} />);
    if (labelW) els.push(<text key={`fl${k}`} x={x0 - 8} y={y + h / 2 + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{trunc(r.label, labelW - 10)}</text>);
    if (!thumb && h >= 13 && (chart.dataLabels ?? true)) {
      const t = labelText(mode, r.v, r.v / first, data.fmt);
      const inside = textW(t, 11) + 8 < bw;
      els.push(<text key={`fv${k}`} x={inside ? x0 + w / 2 : x + bw + 4} y={y + h / 2 + 4} textAnchor={inside ? 'middle' : 'start'} className={inside ? 'onfill' : undefined} style={{ ...(inside ? onFill : ink), fontSize: 11, fontWeight: 600, pointerEvents: 'none' }}>{t}</text>);
    }
  });
};

// ---------- Treemap ----------

const drawTreemap = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, style } = c;
  c.elementTips = true;
  const tree = buildTree(data.paths, data.series[0]?.values ?? []);
  if (!tree.value) return;
  const deep = treeDepth(tree) > 1;
  const gap = thumb ? 0.6 : 1.5;
  const head = thumb || compact ? 0 : 17;
  const labels = !thumb && (chart.dataLabels ?? true);
  const mode = chart.options?.labelMode ?? 'value';
  let k = 0;
  const leaf = (n: TreeNode, x: number, y: number, w: number, h: number) => {
    const color = seriesColor(n.color ?? 0, undefined, style);
    const id = k++;
    const hot = c.hot === id;
    els.push(<rect key={`t${id}`} data-tip="" x={x + gap / 2} y={y + gap / 2} width={Math.max(0.5, w - gap)} height={Math.max(0.5, h - gap)} rx={thumb ? 0.5 : 2}
      style={{ fill: color, fillOpacity: hot ? 0.78 : 1 }}
      onPointerEnter={thumb ? undefined : () => { c.setHot(id); c.setTip({ x: x + w / 2, y: y + h / 2, title: n.path.join(' / '), rows: [{ color, label: data.series[0]!.name, value: fmtv(n.value, data.fmt) }, { color, label: 'Share', value: formatNumber(n.value / tree.value, '0.0%') }] }); }} />);
    if (labels && w > 34 && h > 18) {
      els.push(<text className="onfill" key={`tn${id}`} x={x + 6} y={y + 15} style={{ ...onFill, fontSize: FONT, fontWeight: 600, pointerEvents: 'none' }}>{trunc(n.name, w - 12)}</text>);
      if (h > 34) els.push(<text className="onfill" key={`tv${id}`} x={x + 6} y={y + 29} style={{ ...onFill, fontSize: 10, opacity: 0.9, pointerEvents: 'none' }}>{trunc(labelText(mode, n.value, n.value / tree.value, data.fmt), w - 12, 10)}</text>);
    }
  };
  const tiles = squarify(tree.children.map((n) => ({ item: n, weight: n.value })), { x: plot.x, y: plot.y, w: plot.w, h: plot.h });
  tiles.forEach((t) => {
    const n = t.item;
    if (!deep || !n.children.length) { leaf(n, t.x, t.y, t.w, t.h); return; }
    const hh = head && t.h > head * 2.2 && t.w > 40 ? head : 0;
    if (hh) els.push(<text key={`th${n.name}`} x={t.x + 5} y={t.y + 12} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(n.name, t.w - 10)}</text>);
    const fill = (node: TreeNode, box: { x: number; y: number; w: number; h: number }) => {
      squarify(node.children.map((ch) => ({ item: ch, weight: ch.value })), box).forEach((s) => {
        if (s.item.children.length) fill(s.item, s); else leaf(s.item, s.x, s.y, s.w, s.h);
      });
    };
    fill(n, { x: t.x, y: t.y + hh, w: t.w, h: t.h - hh });
  });
};

// ---------- Sunburst ----------

const drawSunburst = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, style } = c;
  c.elementTips = true;
  const tree = buildTree(data.paths, data.series[0]?.values ?? []);
  if (!tree.value) return;
  const depth = Math.max(1, treeDepth(tree));
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const R = Math.max(6, Math.min(plot.w, plot.h) / 2 - (thumb ? 0 : 4));
  const r0 = R * 0.18;
  const ring = (R - r0) / depth;
  const labels = !thumb && !compact && (chart.dataLabels ?? true);
  sunburstArcs(tree).forEach((a, k) => {
    if (a.a1 - a.a0 < 1e-4) return;
    const inner = r0 + (a.depth - 1) * ring, outer = inner + ring;
    const color = seriesColor(a.node.color ?? 0, undefined, style);
    const opacity = Math.max(0.42, 1 - (a.depth - 1) * 0.22);
    const mid = (a.a0 + a.a1) / 2, mr = (inner + outer) / 2;
    const hot = c.hot === k;
    els.push(<path key={`sb${k}`} data-tip="" d={arcPath(cx, cy, outer, inner, a.a0, a.a1)} style={{ fill: color, fillOpacity: hot ? opacity * 0.75 : opacity, stroke: 'var(--chart-bg)', strokeWidth: thumb ? 0.5 : 1.2 }}
      onPointerEnter={thumb ? undefined : () => { c.setHot(k); c.setTip({ x: cx + Math.cos(mid) * mr, y: cy + Math.sin(mid) * mr, title: a.node.path.join(' / '), rows: [{ color, label: data.series[0]!.name, value: fmtv(a.node.value, data.fmt) }, { color, label: 'Share', value: formatNumber(a.node.value / tree.value, '0.0%') }] }); }} />);
    if (labels && (a.a1 - a.a0) * mr > textW(a.node.name) + 6 && ring > 14) {
      els.push(<text key={`sbl${k}`} x={cx + Math.cos(mid) * mr} y={cy + Math.sin(mid) * mr + 4} textAnchor="middle" className={a.depth === 1 ? 'onfill' : undefined} style={{ ...(a.depth === 1 ? onFill : ink), fontSize: 10.5, fontWeight: 600, pointerEvents: 'none' }}>{a.node.name}</text>);
    }
  });
};
