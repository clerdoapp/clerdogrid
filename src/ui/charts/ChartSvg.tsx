import { useId, useReducer, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import type { Chart } from '../../model/workbook';
import { seriesColor, type ResolvedChart } from './chartData';
import { buildTree } from './chartStats';
import { traits, type Traits } from './chartTypes';
import { cartesianModel, drawBox, drawCartesian, drawScatter } from './drawCartesian';
import { draw3dBars, draw3dLineArea } from './draw3d';
import { drawMap } from './drawMap';
import { drawFunnel, drawRadar, drawSunburst, drawTreemap } from './drawOther';
import { drawOfPie, drawPie, ofPieSplit } from './drawPie';
import { bandColor, bandLabel, drawSurface, surfaceBands } from './drawSurface';
import { axisOf } from './axes';
import { trendName } from './trend';
import { FONT, fillCss, ink, lineCss, LOOKS, markerEl, muted, textW, trunc, WF, type Box, type DrawCtx, type Hover, type LegendEntry, type Tip } from './svgKit';
import { N, T } from '../i18n';

/** Legend entries for the chart's family. */
const legendFor = (chart: Chart, data: ResolvedChart, tr: Traits, style: number): LegendEntry[] => {
  switch (tr.family) {
    case 'pie': {
      const vary = chart.options?.varyColors ?? true;
      if (!vary) return data.series.slice(0, tr.filled ? undefined : 1).map((s) => ({ label: s.name, color: s.color, line: false }));
      return data.categories.map((c, i) => ({ label: c, color: data.series[0]?.src.points?.find((p) => p.idx === i)?.fill?.color ?? seriesColor(i, undefined, style), line: false }));
    }
    case 'ofPie': {
      const s = data.series[0];
      if (!s) return [];
      const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
      const second = ofPieSplit(vals, chart.options?.ofPie?.split, chart.options?.ofPie?.splitValue);
      return data.categories.map((c, i) => ({ label: c, color: seriesColor(i, undefined, style), line: false, second: second[i] }))
        .concat([{ label: N('Other'), color: seriesColor(vals.length, undefined, style), line: false, second: false }]).map(({ second: _s, ...e }) => e);
    }
    case 'treemap': case 'sunburst': {
      const tree = buildTree(data.paths, data.series[0]?.values ?? []);
      return tree.children.map((c, i) => ({ label: c.name, color: seriesColor(i, undefined, style), line: false }));
    }
    case 'funnel': case 'histogram': case 'map': return [];
    case 'surface': {
      const b = surfaceBands(chart, data);
      const fmt = axisOf(chart, 'y').numFmt || data.fmt;
      return b.slice(0, -1).map((lo, k) => ({ label: bandLabel(lo, b[k + 1]!, fmt), color: bandColor(k, style), line: false }));
    }
    case 'waterfall': return [{ label: N('Increase'), color: WF.up, line: false }, { label: N('Decrease'), color: WF.down, line: false }, { label: N('Total'), color: WF.total, line: false }];
    case 'pareto': return data.series[0] ? [{ label: data.series[0].name, color: data.series[0].color, line: false }, { label: N('Cumulative %'), color: seriesColor(1, undefined, style), line: true }] : [];
    default: {
      const out: LegendEntry[] = [];
      const vary = (chart.options?.varyColors ?? false) && data.series.length === 1 && (tr.family === 'cartesian' || tr.family === 'bar3d');
      if (vary) return data.categories.map((c, i) => ({ label: c, color: seriesColor(i, undefined, style), line: false }));
      const prices = tr.family === 'stock' ? data.series.slice(tr.volume && data.series.length > 1 ? 1 : 0) : [];
      for (const s of data.series) {
        if (tr.family === 'stock' && prices.includes(s)) {
          // Price series show only what's drawn: their own line or markers (close gets a dash).
          const ownLine = !!s.src.line && !s.src.line.none;
          const mk = s.src.marker?.symbol ?? (!tr.open && s === prices.at(-1) ? 'dash' : undefined);
          if (ownLine || (mk && mk !== 'none')) out.push({ label: s.name, color: s.color, line: ownLine, dash: s.src.line?.dash, marker: mk });
          continue;
        }
        const isLine = (s.type === 'line' || (tr.family === 'scatter' && tr.lines)) && tr.family !== 'radar' && tr.family !== 'line3d' && tr.family !== 'stock';
        out.push({ label: s.name, color: s.color, line: isLine, dash: s.src.line?.dash, marker: s.src.marker?.symbol });
        s.src.trendlines?.forEach((t) => out.push({ label: trendName(t, s.name), color: t.line?.color ?? s.color, line: true, dash: t.line?.dash ?? 'dash' }));
      }
      return out;
    }
  }
};

/** Draws a chart. `thumb` drops title, legend, labels and hover for small previews. */
export const ChartSvg = ({ chart, data, width, height, thumb = false }: { chart: Chart; data: ResolvedChart; width: number; height: number; thumb?: boolean }) => {
  const svg = useRef<SVGSVGElement>(null);
  const [tip, setTip] = useState<Tip | null>(null);
  const [hot, setHot] = useState<number | null>(null);
  const [, refresh] = useReducer((x: number) => x + 1, 0);
  const uid = `cc${useId().replace(/[^a-zA-Z0-9_-]/g, '')}`;
  const clip = `${uid}clip`;
  const W = Math.max(40, width), H = Math.max(30, height);
  const compact = !thumb && (W < 220 || H < 140);
  const pad = thumb ? 4 : compact ? 6 : 10;
  const tr = traits(chart.type);
  const o = chart.options ?? {};
  const style = o.style ?? 0;
  const look = LOOKS[o.look ?? 0] ?? LOOKS[0]!;
  const els: ReactNode[] = [];
  const defs: ReactNode[] = [];
  const back: ReactNode[] = [];
  let top = pad, bottom = H - pad, left = pad, right = W - pad;
  let overlayLegend: ReactNode[] | null = null;

  // Chart area background and border.
  const ca = o.chartArea;
  if (ca?.fill || ca?.line || look.border) {
    const st: CSSProperties = { ...(ca?.fill ? fillCss({ defs, uid }, ca.fill, 'var(--chart-bg)') : { fill: 'none' }), ...(ca?.line || look.border ? lineCss(ca?.line, 'var(--chart-grid)', 1) : {}) };
    back.push(<rect key="ca" x={0.5} y={0.5} width={W - 1} height={H - 1} rx={look.rounded ? 8 : 0} style={st} />);
  }

  const title = thumb ? '' : data.title?.trim() ?? chart.title?.trim() ?? '';
  if (title) {
    const size = compact ? 12 : 14;
    els.push(<text key="title" x={W / 2} y={top + size - 1} textAnchor="middle" style={{ ...ink, fontSize: size, fontWeight: 600 }}>{trunc(title, W - 2 * pad, size)}</text>);
    top += size + 10;
  }

  const legendItems = legendFor(chart, data, tr, style);
  const auto = legendItems.length > 1 || tr.family === 'pie' || tr.family === 'surface' ? (tr.family === 'surface' ? 'right' : 'bottom') : 'none';
  const legend = thumb ? 'none' : chart.legend ?? (compact ? 'none' : auto);
  const overlay = !!o.legendOverlay;
  if (legend !== 'none' && legendItems.length) {
    const itemW = (l: string) => Math.min(160, textW(l)) + 26;
    const legendEls: ReactNode[] = [];
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
        row.forEach((it, i) => { legendEls.push(<LegendItem key={`lg${ri}-${i}`} x={x} y={y0 + ri * 16 + 8} {...it} max={160} />); x += itemW(it.label); });
      });
      if (!overlay) { if (legend === 'top') top += h + 4; else bottom -= h + 4; }
    } else {
      const colW = Math.min(W * 0.35, Math.max(...legendItems.map((it) => itemW(it.label))));
      const shown = legendItems.slice(0, Math.max(1, Math.floor((bottom - top) / 16)));
      const y0 = top + (bottom - top - shown.length * 16) / 2 + 8;
      const x0 = legend === 'left' ? left : right - colW;
      if (overlay) legendEls.push(<rect key="lgbg" x={x0 - 4} y={y0 - 12} width={colW + 4} height={shown.length * 16 + 8} rx={4} style={{ fill: 'var(--chart-bg)', fillOpacity: 0.85 }} />);
      shown.forEach((it, i) => legendEls.push(<LegendItem key={`lg${i}`} x={x0} y={y0 + i * 16} {...it} max={colW - 26} />));
      if (!overlay) { if (legend === 'left') left += colW + 6; else right -= colW + 6; }
    }
    // Overlaid legends draw last, above the plot.
    if (overlay) overlayLegend = legendEls; else els.push(...legendEls);
  }

  const plot: Box = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  const pa = o.plotArea;
  if (pa?.fill || pa?.line || look.plotBorder || look.plotTint) {
    const st: CSSProperties = { ...(pa?.fill ? fillCss({ defs, uid }, pa.fill, 'var(--chart-bg)') : look.plotTint ? { fill: 'var(--chart-hover)' } : { fill: 'none' }), ...(pa?.line || look.plotBorder ? lineCss(pa?.line, 'var(--chart-grid)', 1) : {}) };
    back.push(<rect key="pa" x={plot.x} y={plot.y} width={plot.w} height={plot.h} rx={look.rounded ? 6 : 0} style={st} />);
  }
  const hasData = data.series.some((s) => s.values.some((v) => v !== null));
  const toLocal = (cx: number, cy: number): [number, number] => {
    const r = svg.current?.getBoundingClientRect();
    return r && r.width && r.height ? [((cx - r.left) * W) / r.width, ((cy - r.top) * H) / r.height] : [0, 0];
  };
  const hover: Hover = { onMove: () => undefined, onLeave: () => { setTip(null); setHot(null); } };
  const ctx: DrawCtx = { chart, data, tr, plot, thumb, compact, hot, els, defs, setTip, setHot, hover, clip, elementTips: false, style, uid, lineScale: look.lineScale ?? 1, refresh };

  if (!hasData) {
    els.push(<text key="empty" x={W / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: thumb ? 9 : 12 }}>{thumb || compact ? T('No data') : T('Add numbers to the data range to see a chart.')}</text>);
  } else {
    switch (tr.family) {
      case 'pie': drawPie(ctx); break;
      case 'ofPie': drawOfPie(ctx); break;
      case 'radar': drawRadar(ctx); break;
      case 'scatter': case 'bubble': drawScatter(ctx); break;
      case 'box': drawBox(ctx); break;
      case 'treemap': drawTreemap(ctx); break;
      case 'sunburst': drawSunburst(ctx); break;
      case 'funnel': drawFunnel(ctx); break;
      case 'bar3d': draw3dBars(ctx); break;
      case 'line3d': case 'area3d': draw3dLineArea(ctx); break;
      case 'surface': drawSurface(ctx); break;
      case 'map': drawMap(ctx); break;
      default: drawCartesian(ctx, cartesianModel(ctx));
    }
  }
  if (overlayLegend) els.push(...overlayLegend);

  return (
    <div className="chartsvg" style={look.vars as CSSProperties | undefined}>
      <svg ref={svg} direction="ltr" viewBox={`0 0 ${W} ${H}`} width="100%" height="100%" role="img" aria-label={title || T('Chart')}
        style={look.vars ? { background: 'var(--chart-bg)' } : undefined}
        onPointerMove={thumb ? undefined : (e) => {
          if (ctx.elementTips) { if (!(e.target as Element).closest?.('[data-tip]')) hover.onLeave(); return; }
          const [x, y] = toLocal(e.clientX, e.clientY); ctx.hover.onMove(x, y);
        }}
        onPointerLeave={thumb ? undefined : () => hover.onLeave()}>
        <defs><clipPath id={clip}><rect x={plot.x - 2} y={plot.y - 2} width={plot.w + 4} height={plot.h + 4} /></clipPath>{defs}</defs>
        {back}
        {els}
      </svg>
      {tip && !thumb && (
        <div className="charttip" style={{ left: `${(tip.x / W) * 100}%`, top: `${(tip.y / H) * 100}%`, transform: `translate(${tip.x > W * 0.6 ? 'calc(-100% - 10px)' : '10px'}, ${tip.y > H * 0.6 ? 'calc(-100% - 6px)' : '6px'})` }}>
          {tip.title && <div className="tiptitle">{tip.title}</div>}
          {tip.rows.map((r, i) => <div key={i} className="tiprow"><i style={{ background: r.color }} /><span>{r.label}</span>{r.value && <b>{r.value}</b>}</div>)}
        </div>
      )}
    </div>
  );
};

const LegendItem = ({ x, y, label, color, line, max, dash, marker }: LegendEntry & { x: number; y: number; max: number }) => (
  <g>
    {line
      ? <line x1={x} x2={x + 14} y1={y} y2={y} style={{ stroke: color, strokeWidth: 2.5, strokeLinecap: 'round', strokeDasharray: dash && dash !== 'solid' ? '4 3' : undefined }} />
      : marker && marker !== 'auto' && marker !== 'none'
        ? markerEl('m', marker, x + 7, y, 9, marker === 'dash' ? color : color, color, 1.5)
        : <rect x={x + 2} y={y - 5} width={10} height={10} rx={2} style={{ fill: color }} />}
    {line && marker && marker !== 'none' && marker !== 'auto' && markerEl('mm', marker, x + 7, y, 7, 'var(--chart-bg)', color, 1.5)}
    <text x={x + 18} y={y + 4} style={{ ...ink, fontSize: FONT }}>{trunc(label, max)}</text>
  </g>
);
