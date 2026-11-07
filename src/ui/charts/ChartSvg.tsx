import { useRef, useState, type ReactNode } from 'react';
import type { Chart, ChartType } from '../../model/workbook';
import { formatNumber, niceScale, seriesColor, valueExtent, type ResolvedChart, type ResolvedSeries, type Scale } from './chartData';

const FONT = 11;
const textW = (s: string, size = FONT): number => s.length * size * 0.56;
const trunc = (s: string, maxW: number, size = FONT): string => {
  if (textW(s, size) <= maxW) return s;
  const n = Math.max(1, Math.floor(maxW / (size * 0.56)) - 1);
  return `${s.slice(0, n)}…`;
};

const ink = { fill: 'var(--chart-ink)' } as const;
const muted = { fill: 'var(--chart-muted)' } as const;

interface Tip { x: number; y: number; title: string; rows: { color: string; label: string; value: string }[] }
interface Box { x: number; y: number; w: number; h: number }
interface Hover { onMove: (x: number, y: number) => void; onLeave: () => void }

const HORIZONTAL: ChartType[] = ['bar', 'stackedBar'];
const STACKED: ChartType[] = ['stackedColumn', 'stackedBar', 'stackedArea', 'percentColumn'];

/** Draws a chart. `thumb` drops title, legend, labels and hover for small previews. */
export const ChartSvg = ({ chart, data, width, height, thumb = false }: { chart: Chart; data: ResolvedChart; width: number; height: number; thumb?: boolean }) => {
  const svg = useRef<SVGSVGElement>(null);
  const [tip, setTip] = useState<Tip | null>(null);
  const [hot, setHot] = useState<number | null>(null);
  const W = Math.max(40, width), H = Math.max(30, height);
  const pad = thumb ? 4 : 10;
  const els: ReactNode[] = [];
  let top = pad, bottom = H - pad, left = pad, right = W - pad;

  const title = thumb ? '' : chart.title?.trim() ?? '';
  if (title) {
    els.push(<text key="title" x={W / 2} y={top + 13} textAnchor="middle" style={{ ...ink, fontSize: 14, fontWeight: 600 }}>{trunc(title, W - 2 * pad, 14)}</text>);
    top += 24;
  }

  const pieLike = chart.type === 'pie' || chart.type === 'doughnut';
  const legendItems = pieLike ? data.categories.map((c, i) => ({ label: c, color: seriesColor(i), line: false }))
    : data.series.map((s) => ({ label: s.name, color: s.color, line: s.type === 'line' && chart.type !== 'radar' }));
  const legend = thumb ? 'none' : chart.legend ?? (legendItems.length > 1 || pieLike ? 'bottom' : 'none');
  if (legend !== 'none' && legendItems.length) {
    const itemW = (l: string) => Math.min(160, textW(l)) + 26;
    if (legend === 'top' || legend === 'bottom') {
      const rows: (typeof legendItems)[] = [[]];
      let used = 0;
      for (const it of legendItems) {
        if (used + itemW(it.label) > right - left && rows.at(-1)!.length) { if (rows.length === 3) break; rows.push([]); used = 0; }
        rows.at(-1)!.push(it); used += itemW(it.label);
      }
      const h = rows.length * 16 + 4;
      const y0 = legend === 'top' ? top : bottom - h + 4;
      rows.forEach((row, ri) => {
        const total = row.reduce((a, it) => a + itemW(it.label), 0);
        let x = (W - total) / 2;
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
    return r ? [((cx - r.left) * W) / r.width, ((cy - r.top) * H) / r.height] : [0, 0];
  };
  const hover: Hover = {
    onMove: (x, y) => { void x; void y; },
    onLeave: () => { setTip(null); setHot(null); },
  };
  const ctx: DrawCtx = { chart, data, plot, thumb, hot, els, setTip, setHot, hover };

  if (!hasData) {
    els.push(<text key="empty" x={W / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: thumb ? 9 : 12 }}>{thumb ? 'No data' : 'Add numbers to the data range to see a chart.'}</text>);
  } else if (pieLike) drawPie(ctx);
  else if (chart.type === 'radar') drawRadar(ctx);
  else if (chart.type === 'scatter') drawScatter(ctx);
  else drawCartesian(ctx);

  return (
    <div className="chartsvg">
      <svg ref={svg} viewBox={`0 0 ${W} ${H}`} width="100%" height="100%" role="img" aria-label={title || 'Chart'}
        onPointerMove={thumb ? undefined : (e) => { const [x, y] = toLocal(e.clientX, e.clientY); ctx.hover.onMove(x, y); }}
        onPointerLeave={thumb ? undefined : () => hover.onLeave()}>
        {els}
      </svg>
      {tip && !thumb && (
        <div className="charttip" style={{ left: `${(tip.x / W) * 100}%`, top: `${(tip.y / H) * 100}%`, transform: `translate(${tip.x > W * 0.6 ? 'calc(-100% - 10px)' : '10px'}, ${tip.y > H * 0.6 ? 'calc(-100% - 6px)' : '6px'})` }}>
          <div className="tiptitle">{tip.title}</div>
          {tip.rows.map((r, i) => <div key={i} className="tiprow"><i style={{ background: r.color }} /><span>{r.label}</span><b>{r.value}</b></div>)}
        </div>
      )}
    </div>
  );
};

const LegendItem = ({ x, y, label, color, line, max }: { x: number; y: number; label: string; color: string; line: boolean; max: number }) => (
  <g>
    {line ? <line x1={x} x2={x + 14} y1={y} y2={y} style={{ stroke: color, strokeWidth: 2.5 }} /> : <rect x={x + 2} y={y - 5} width={10} height={10} rx={2} style={{ fill: color }} />}
    <text x={x + 18} y={y + 4} style={{ ...ink, fontSize: FONT }}>{trunc(label, max)}</text>
  </g>
);

interface DrawCtx {
  chart: Chart; data: ResolvedChart; plot: Box; thumb: boolean; hot: number | null; els: ReactNode[];
  setTip: (t: Tip | null) => void; setHot: (i: number | null) => void; hover: Hover;
}

const valText = (v: number, fmt: string) => formatNumber(v, fmt);

// ---------- Column, bar, line, area, stacked, percent, combo ----------

interface Seg { s: number; i: number; v0: number; v1: number; raw: number | null }

const drawCartesian = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, els } = c;
  const type = chart.type;
  const horizontal = HORIZONTAL.includes(type);
  const stacked = STACKED.includes(type);
  const percent = type === 'percentColumn';
  const n = Math.max(1, data.categories.length);
  const series = data.series;
  const fmt = percent ? '0%' : data.fmt;

  const segs: Seg[][] = series.map(() => []);
  for (let i = 0; i < n; i++) {
    let pos = 0, neg = 0;
    const total = percent ? series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0) : 1;
    series.forEach((s, si) => {
      const raw = s.values[i] ?? null;
      let v = raw ?? 0;
      if (percent) v = total ? Math.abs(v) / total : 0;
      if (stacked) {
        if (v >= 0) { segs[si]!.push({ s: si, i, v0: pos, v1: pos + v, raw }); pos += v; }
        else { segs[si]!.push({ s: si, i, v0: neg, v1: neg + v, raw }); neg += v; }
      } else segs[si]!.push({ s: si, i, v0: 0, v1: v, raw });
    });
  }
  const all = segs.flat().filter((g) => g.raw !== null || stacked).flatMap((g) => [g.v0, g.v1]);
  const barish = type !== 'line';
  const [lo, hi] = percent ? [0, 1] : valueExtent(stacked ? all : segs.flat().filter((g) => g.raw !== null).map((g) => g.v1), barish);

  const yTitle = !thumb && chart.axisTitles?.y ? chart.axisTitles.y : '';
  const xTitle = !thumb && chart.axisTitles?.x ? chart.axisTitles.x : '';
  const valueLen = horizontal ? plot.w : plot.h;
  const scale = niceScale(lo, hi, thumb ? 3 : Math.max(2, Math.min(8, Math.floor(valueLen / (horizontal ? 70 : 40)))));
  const tickText = scale.ticks.map((t) => formatNumber(t, fmt, true));
  const cats = data.categories;
  const maxCat = Math.max(0, ...cats.map((s) => textW(s)));

  let px = plot.x, py = plot.y, pw = plot.w, ph = plot.h;
  let rotate = false, skip = 1;
  if (!thumb) {
    if (yTitle) { px += 18; pw -= 18; }
    if (xTitle) ph -= 18;
    if (horizontal) {
      const g = Math.min(maxCat, plot.w * 0.3) + 8;
      px += g; pw -= g; ph -= 18;
      const band = ph / n;
      skip = Math.max(1, Math.ceil(13 / band));
    } else {
      const g = Math.max(...tickText.map((t) => textW(t))) + 8;
      px += g; pw -= g;
      const band = pw / n;
      if (maxCat + 6 > band) {
        rotate = true;
        const h = Math.min(maxCat * 0.72 + 10, plot.h * 0.3);
        ph -= h;
        skip = Math.max(1, Math.ceil(14 / band));
      } else ph -= 18;
    }
  }
  pw = Math.max(10, pw); ph = Math.max(10, ph);
  const span = scale.max - scale.min || 1;
  const band = (horizontal ? ph : pw) / n;
  const valPos = (v: number) => (horizontal ? px + ((v - scale.min) / span) * pw : py + ph - ((v - scale.min) / span) * ph);
  const catPos = (i: number) => (horizontal ? py + band * (i + 0.5) : px + band * (i + 0.5));
  const zero = valPos(Math.min(Math.max(0, scale.min), scale.max));

  // Gridlines and value-axis labels.
  scale.ticks.forEach((t, k) => {
    const p = valPos(t);
    els.push(horizontal
      ? <line key={`g${k}`} x1={p} x2={p} y1={py} y2={py + ph} style={{ stroke: 'var(--chart-grid)', strokeWidth: 1 }} />
      : <line key={`g${k}`} x1={px} x2={px + pw} y1={p} y2={p} style={{ stroke: 'var(--chart-grid)', strokeWidth: 1 }} />);
    if (!thumb) els.push(horizontal
      ? <text key={`t${k}`} x={p} y={py + ph + 13} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>
      : <text key={`t${k}`} x={px - 6} y={p + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{tickText[k]}</text>);
  });
  els.push(horizontal
    ? <line key="axis" x1={zero} x2={zero} y1={py} y2={py + ph} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1 }} />
    : <line key="axis" x1={px} x2={px + pw} y1={zero} y2={zero} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1 }} />);

  // Category labels.
  if (!thumb) cats.forEach((label, i) => {
    if (i % skip) return;
    const p = catPos(i);
    if (horizontal) els.push(<text key={`c${i}`} x={px - 6} y={p + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{trunc(label, Math.min(maxCat, plot.w * 0.3))}</text>);
    else if (rotate) els.push(<text key={`c${i}`} x={p} y={py + ph + 10} textAnchor="end" transform={`rotate(-45 ${p} ${py + ph + 10})`} style={{ ...muted, fontSize: FONT }}>{trunc(label, (plot.h * 0.3 - 10) / 0.72)}</text>);
    else els.push(<text key={`c${i}`} x={p} y={py + ph + 14} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{label}</text>);
  });
  if (yTitle) els.push(<text key="yt" x={plot.x + 10} y={py + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${py + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(yTitle, ph)}</text>);
  if (xTitle) els.push(<text key="xt" x={px + pw / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{trunc(xTitle, pw)}</text>);

  if (c.hot !== null && !thumb) {
    const b = c.hot;
    els.push(horizontal
      ? <rect key="hot" x={px} y={py + b * band} width={pw} height={band} style={{ fill: 'var(--chart-hover)' }} />
      : <rect key="hot" x={px + b * band} y={py} width={band} height={ph} style={{ fill: 'var(--chart-hover)' }} />);
  }

  const isBar = (s: ResolvedSeries) => type === 'combo' ? s.type === 'column' : !['line', 'area', 'stackedArea'].includes(type);
  const barSeries = series.map((s, i) => ({ s, i })).filter((x) => isBar(x.s));
  const group = band * (stacked ? 0.62 : 0.74);
  const labels = !thumb && chart.dataLabels && n * series.length <= 240;

  // Bars.
  barSeries.forEach(({ s, i: si }, j) => {
    const w = stacked ? group : group / barSeries.length;
    segs[si]!.forEach((g) => {
      if (g.raw === null && !stacked) return;
      const off = stacked ? -group / 2 : -group / 2 + j * w;
      const a = valPos(g.v0), b = valPos(g.v1);
      const cp = catPos(g.i) + off;
      const r = horizontal ? { x: Math.min(a, b), y: cp, width: Math.abs(b - a), height: Math.max(0.5, w - 1) } : { x: cp, y: Math.min(a, b), width: Math.max(0.5, w - 1), height: Math.abs(b - a) };
      els.push(<rect key={`b${si}-${g.i}`} {...r} rx={Math.min(2, w / 6)} style={{ fill: s.color }} />);
      if (labels && g.raw !== null) {
        const t = percent ? formatNumber(g.v1 - g.v0, '0%') : valText(g.raw, fmt);
        if (stacked) {
          if ((horizontal ? r.width : r.height) > 14 && textW(t, 10) < (horizontal ? r.width : w)) els.push(<text key={`l${si}-${g.i}`} x={r.x + r.width / 2} y={r.y + r.height / 2 + 3.5} textAnchor="middle" style={{ fill: '#fff', fontSize: 10 }}>{t}</text>);
        } else if (horizontal) els.push(<text key={`l${si}-${g.i}`} x={g.v1 >= 0 ? b + 4 : b - 4} y={cp + w / 2 + 3.5} textAnchor={g.v1 >= 0 ? 'start' : 'end'} style={{ ...ink, fontSize: 10 }}>{t}</text>);
        else els.push(<text key={`l${si}-${g.i}`} x={cp + w / 2} y={g.v1 >= 0 ? b - 4 : b + 12} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{t}</text>);
      }
    });
  });

  // Areas and lines.
  series.forEach((s, si) => {
    if (isBar(s)) return;
    const pts = segs[si]!;
    const area = type === 'area' || type === 'stackedArea';
    if (area) {
      const xy = (g: Seg, v: number) => `${catPos(g.i)},${valPos(v)}`;
      const usable = stacked ? pts : pts.filter((g) => g.raw !== null);
      if (usable.length) {
        const upper = usable.map((g) => xy(g, g.v1));
        const lower = [...usable].reverse().map((g) => xy(g, stacked ? g.v0 : Math.max(0, scale.min)));
        els.push(<polygon key={`a${si}`} points={[...upper, ...lower].join(' ')} style={{ fill: s.color, fillOpacity: stacked ? 0.85 : 0.35 }} />);
        els.push(<polyline key={`al${si}`} points={upper.join(' ')} style={{ fill: 'none', stroke: s.color, strokeWidth: 1.6 }} />);
      }
    } else {
      // Break the line at blank cells.
      const runs: string[][] = [[]];
      pts.forEach((g) => { if (g.raw === null) { if (runs.at(-1)!.length) runs.push([]); } else runs.at(-1)!.push(`${catPos(g.i)},${valPos(g.v1)}`); });
      runs.forEach((r, k) => r.length > 1 && els.push(<polyline key={`ln${si}-${k}`} points={r.join(' ')} style={{ fill: 'none', stroke: s.color, strokeWidth: thumb ? 1.6 : 2.2, strokeLinejoin: 'round', strokeLinecap: 'round' }} />));
      if (n <= 40 || runs.some((r) => r.length === 1)) pts.forEach((g) => g.raw !== null && els.push(<circle key={`m${si}-${g.i}`} cx={catPos(g.i)} cy={valPos(g.v1)} r={thumb ? 1.6 : c.hot === g.i ? 4.5 : 3} style={{ fill: 'var(--chart-bg)', stroke: s.color, strokeWidth: 2 }} />));
    }
    if (labels) pts.forEach((g) => g.raw !== null && els.push(<text key={`l${si}-${g.i}`} x={catPos(g.i)} y={valPos(g.v1) - 7} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{valText(g.raw, fmt)}</text>));
  });

  if (thumb) return;
  c.hover.onMove = (x, y) => {
    const inside = x >= px && x <= px + pw && y >= py && y <= py + ph;
    if (!inside) { c.setTip(null); c.setHot(null); return; }
    const i = Math.min(n - 1, Math.max(0, Math.floor(((horizontal ? y - py : x - px) / (horizontal ? ph : pw)) * n)));
    c.setHot(i);
    c.setTip({
      x, y, title: cats[i] ?? '',
      rows: series.map((s, si) => {
        const g = segs[si]![i];
        const v = g?.raw ?? null;
        const value = v === null ? 'Blank' : percent ? `${valText(v, data.fmt)} (${formatNumber(g!.v1 - g!.v0, '0.0%')})` : valText(v, data.fmt);
        return { color: s.color, label: s.name, value };
      }),
    });
  };
};

// ---------- Scatter ----------

const drawScatter = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, els } = c;
  const pts = data.series.flatMap((s, si) => s.values.map((y, i) => ({ si, i, x: data.xs?.[i] ?? (data.xs ? null : i + 1), y })).filter((p): p is { si: number; i: number; x: number; y: number } => p.x !== null && p.y !== null));
  const [xl, xh] = valueExtent(pts.map((p) => p.x), false);
  const [yl, yh] = valueExtent(pts.map((p) => p.y), false);
  let px = plot.x, pw = plot.w, ph = plot.h;
  const yTitle = !thumb ? chart.axisTitles?.y ?? '' : '', xTitle = !thumb ? chart.axisTitles?.x ?? '' : '';
  const ys = niceScale(yl, yh, thumb ? 3 : Math.max(2, Math.floor(plot.h / 40)));
  const yText = ys.ticks.map((t) => formatNumber(t, data.fmt, true));
  if (!thumb) {
    const g = Math.max(...yText.map((t) => textW(t))) + 8 + (yTitle ? 18 : 0);
    px += g; pw -= g; ph -= 18 + (xTitle ? 18 : 0);
  }
  const xs = niceScale(xl, xh, thumb ? 3 : Math.max(2, Math.floor(pw / 70)));
  const X = (v: number) => px + ((v - xs.min) / (xs.max - xs.min || 1)) * pw;
  const Y = (v: number) => plot.y + ph - ((v - ys.min) / (ys.max - ys.min || 1)) * ph;
  gridY(els, ys, Y, px, pw, thumb ? null : yText);
  xs.ticks.forEach((t, k) => {
    els.push(<line key={`gx${k}`} x1={X(t)} x2={X(t)} y1={plot.y} y2={plot.y + ph} style={{ stroke: 'var(--chart-grid)' }} />);
    if (!thumb) els.push(<text key={`tx${k}`} x={X(t)} y={plot.y + ph + 14} textAnchor="middle" style={{ ...muted, fontSize: FONT }}>{formatNumber(t, 'General', true)}</text>);
  });
  if (yTitle) els.push(<text key="yt" x={plot.x + 10} y={plot.y + ph / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 10} ${plot.y + ph / 2})`} style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{yTitle}</text>);
  if (xTitle) els.push(<text key="xt" x={px + pw / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: FONT, fontWeight: 600 }}>{xTitle}</text>);
  const hotKey = c.hot;
  pts.forEach((p, k) => {
    els.push(<circle key={`p${k}`} cx={X(p.x)} cy={Y(p.y)} r={thumb ? 1.8 : hotKey === k ? 5.5 : 4} style={{ fill: data.series[p.si]!.color, fillOpacity: 0.85, stroke: 'var(--chart-bg)', strokeWidth: 1 }} />);
    if (!thumb && chart.dataLabels && pts.length <= 200) els.push(<text key={`pl${k}`} x={X(p.x)} y={Y(p.y) - 7} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{formatNumber(p.y, data.fmt)}</text>);
  });
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    let best = -1, bd = 24 * 24;
    pts.forEach((p, k) => { const d = (X(p.x) - x) ** 2 + (Y(p.y) - y) ** 2; if (d < bd) { bd = d; best = k; } });
    if (best < 0) { c.setTip(null); c.setHot(null); return; }
    const p = pts[best]!;
    const s = data.series[p.si]!;
    c.setHot(best);
    c.setTip({ x: X(p.x), y: Y(p.y), title: s.name, rows: [{ color: s.color, label: 'X', value: formatNumber(p.x, 'General') }, { color: s.color, label: 'Y', value: formatNumber(p.y, data.fmt) }] });
  };
};

const gridY = (els: ReactNode[], s: Scale, Y: (v: number) => number, px: number, pw: number, text: string[] | null) => {
  s.ticks.forEach((t, k) => {
    els.push(<line key={`gy${k}`} x1={px} x2={px + pw} y1={Y(t)} y2={Y(t)} style={{ stroke: 'var(--chart-grid)' }} />);
    if (text) els.push(<text key={`ty${k}`} x={px - 6} y={Y(t) + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{text[k]}</text>);
  });
};

// ---------- Pie and doughnut ----------

const drawPie = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, els } = c;
  const s = data.series[0]!;
  const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const r = Math.max(4, Math.min(plot.w, plot.h) / 2 - (thumb ? 0 : 6));
  const inner = chart.type === 'doughnut' ? r * 0.56 : 0;
  if (!total) { els.push(<text key="none" x={cx} y={cy} textAnchor="middle" style={{ ...muted, fontSize: 11 }}>No positive values</text>); return; }
  let a0 = -Math.PI / 2;
  const slices = vals.map((v, i) => { const a1 = a0 + (v / total) * Math.PI * 2; const sl = { i, v, a0, a1 }; a0 = a1; return sl; });
  slices.forEach((sl) => {
    if (!sl.v) return;
    const mid = (sl.a0 + sl.a1) / 2;
    const off = c.hot === sl.i ? 5 : 0;
    const ox = Math.cos(mid) * off, oy = Math.sin(mid) * off;
    els.push(<path key={`s${sl.i}`} d={arcPath(cx + ox, cy + oy, r, inner, sl.a0, sl.a1)} style={{ fill: seriesColor(sl.i), stroke: 'var(--chart-bg)', strokeWidth: thumb ? 0.6 : 1.5 }}
      onPointerEnter={thumb ? undefined : () => { c.setHot(sl.i); c.setTip({ x: cx + Math.cos(mid) * r * 0.7, y: cy + Math.sin(mid) * r * 0.7, title: data.categories[sl.i] ?? '', rows: [{ color: seriesColor(sl.i), label: s.name, value: `${formatNumber(sl.v, data.fmt)} (${formatNumber(sl.v / total, '0.0%')})` }] }); }} />);
    if (!thumb && chart.dataLabels && sl.a1 - sl.a0 > 0.28) {
      const lr = inner ? (r + inner) / 2 : r * 0.64;
      els.push(<text key={`sl${sl.i}`} x={cx + ox + Math.cos(mid) * lr} y={cy + oy + Math.sin(mid) * lr + 4} textAnchor="middle" style={{ fill: '#fff', fontSize: 11, fontWeight: 600, pointerEvents: 'none' }}>{formatNumber(sl.v / total, '0%')}</text>);
    }
  });
  if (inner && !thumb) els.push(<text key="tot" x={cx} y={cy + 5} textAnchor="middle" style={{ ...ink, fontSize: 14, fontWeight: 600, pointerEvents: 'none' }}>{formatNumber(total, data.fmt, true)}</text>);
  c.hover.onMove = (x, y) => { if ((x - cx) ** 2 + (y - cy) ** 2 > r * r || (x - cx) ** 2 + (y - cy) ** 2 < inner * inner) { c.setTip(null); c.setHot(null); } };
};

const arcPath = (cx: number, cy: number, r: number, inner: number, a0: number, a1: number): string => {
  // A full circle can't be one arc; split it.
  if (a1 - a0 >= Math.PI * 2 - 1e-6) a1 = a0 + Math.PI * 2 - 1e-4;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const p = (rad: number, a: number) => `${cx + Math.cos(a) * rad},${cy + Math.sin(a) * rad}`;
  if (!inner) return `M${cx},${cy} L${p(r, a0)} A${r},${r} 0 ${large} 1 ${p(r, a1)} Z`;
  return `M${p(r, a0)} A${r},${r} 0 ${large} 1 ${p(r, a1)} L${p(inner, a1)} A${inner},${inner} 0 ${large} 0 ${p(inner, a0)} Z`;
};

// ---------- Radar ----------

const drawRadar = (c: DrawCtx): void => {
  const { data, plot, thumb, els } = c;
  const n = Math.max(3, data.categories.length);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sc = niceScale(Math.min(0, ...vals), Math.max(...vals), thumb ? 3 : 5);
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const R = Math.max(6, Math.min(plot.w, plot.h) / 2 - (thumb ? 2 : 22));
  const ang = (i: number) => -Math.PI / 2 + (i / n) * Math.PI * 2;
  const rad = (v: number) => ((v - sc.min) / (sc.max - sc.min || 1)) * R;
  const pt = (i: number, v: number) => `${cx + Math.cos(ang(i)) * rad(v)},${cy + Math.sin(ang(i)) * rad(v)}`;
  sc.ticks.forEach((t, k) => {
    if (t === sc.min) return;
    els.push(<polygon key={`r${k}`} points={Array.from({ length: n }, (_, i) => pt(i, t)).join(' ')} style={{ fill: 'none', stroke: 'var(--chart-grid)' }} />);
    if (!thumb) els.push(<text key={`rt${k}`} x={cx + 3} y={cy - rad(t) + 10} style={{ ...muted, fontSize: 9 }}>{formatNumber(t, data.fmt, true)}</text>);
  });
  for (let i = 0; i < n; i++) {
    els.push(<line key={`sp${i}`} x1={cx} y1={cy} x2={cx + Math.cos(ang(i)) * R} y2={cy + Math.sin(ang(i)) * R} style={{ stroke: c.hot === i ? 'var(--chart-axis)' : 'var(--chart-grid)' }} />);
    if (!thumb) {
      const a = ang(i), lx = cx + Math.cos(a) * (R + 10), ly = cy + Math.sin(a) * (R + 10) + 4;
      els.push(<text key={`cl${i}`} x={lx} y={ly} textAnchor={Math.abs(Math.cos(a)) < 0.2 ? 'middle' : Math.cos(a) > 0 ? 'start' : 'end'} style={{ ...muted, fontSize: FONT }}>{trunc(data.categories[i] ?? '', 90)}</text>);
    }
  }
  data.series.forEach((s, si) => {
    const pts = Array.from({ length: n }, (_, i) => pt(i, s.values[i] ?? Math.max(0, sc.min)));
    els.push(<polygon key={`rs${si}`} points={pts.join(' ')} style={{ fill: s.color, fillOpacity: 0.14, stroke: s.color, strokeWidth: thumb ? 1.2 : 2, strokeLinejoin: 'round' }} />);
    if (!thumb && c.chart.dataLabels) s.values.forEach((v, i) => v !== null && els.push(<text key={`rl${si}-${i}`} x={cx + Math.cos(ang(i)) * rad(v)} y={cy + Math.sin(ang(i)) * rad(v) - 6} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{formatNumber(v, data.fmt)}</text>));
  });
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    if ((x - cx) ** 2 + (y - cy) ** 2 > (R + 16) ** 2) { c.setTip(null); c.setHot(null); return; }
    let a = Math.atan2(y - cy, x - cx) + Math.PI / 2;
    if (a < 0) a += Math.PI * 2;
    const i = Math.round((a / (Math.PI * 2)) * n) % n;
    c.setHot(i);
    c.setTip({ x, y, title: data.categories[i] ?? '', rows: data.series.map((s) => ({ color: s.color, label: s.name, value: s.values[i] === null || s.values[i] === undefined ? 'Blank' : formatNumber(s.values[i]!, data.fmt) })) });
  };
};
