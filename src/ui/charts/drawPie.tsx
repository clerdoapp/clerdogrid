// Pie, doughnut, exploded and 3D pies, pie of pie and bar of pie.
import type { ReactNode } from 'react';
import type { DataLabelOptions } from '../../model/workbook';
import { formatNumber, seriesColor, type ResolvedSeries } from './chartData';
import { fillCss, fmtv, ink, labelText, lineCss, muted, onFill, pointFormat, textW, type DrawCtx } from './svgKit';
import { N, T } from '../i18n';

export const arcPath = (cx: number, cy: number, r: number, inner: number, a0: number, a1: number, ky = 1): string => {
  if (a1 - a0 >= Math.PI * 2 - 1e-6) a1 = a0 + Math.PI * 2 - 1e-4;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const p = (rad: number, a: number) => `${(cx + Math.cos(a) * rad).toFixed(2)},${(cy + Math.sin(a) * rad * ky).toFixed(2)}`;
  if (!inner) return `M${cx},${cy} L${p(r, a0)} A${r},${r * ky} 0 ${large} 1 ${p(r, a1)} Z`;
  return `M${p(r, a0)} A${r},${r * ky} 0 ${large} 1 ${p(r, a1)} L${p(inner, a1)} A${inner},${inner * ky} 0 ${large} 0 ${p(inner, a0)} Z`;
};

/** Label settings for pies: explicit options, else the older label mode (percent by default). */
const pieLabels = (c: DrawCtx, s: ResolvedSeries): DataLabelOptions | null => {
  const { chart } = c;
  if (s.src.labels) { const l = s.src.labels; return l.value || l.percent || l.seriesName || l.categoryName || l.fromCells ? l : null; }
  if (!chart.dataLabels) return null;
  if (chart.options?.labels) return chart.options.labels;
  const mode = chart.options?.labelMode ?? 'percent';
  return { value: mode !== 'percent', percent: mode !== 'value' };
};

interface Slice { i: number; v: number; a0: number; a1: number; color: string; off: number }

const sliceColor = (c: DrawCtx, s: ResolvedSeries, i: number): string => {
  const pf = pointFormat(s, i);
  if (pf?.fill?.color) return pf.fill.color;
  const vary = c.chart.options?.varyColors ?? true;
  return vary ? seriesColor(i, undefined, c.style) : s.color;
};

const explodeOf = (c: DrawCtx, s: ResolvedSeries, i: number): number => {
  const pf = pointFormat(s, i);
  return Math.max(0, Math.min(400, pf?.explosion ?? s.src.explosion ?? (c.tr.exploded ? 18 : 0)));
};

export const drawPie = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, tr } = c;
  c.elementTips = true;
  const o = chart.options ?? {};
  const doughnut = tr.filled;
  const rings = doughnut ? data.series : data.series.slice(0, 1);
  const tilt = tr.d3 ? Math.max(10, Math.min(90, o.view3D?.rotX ?? 30)) : 90;
  const ky = Math.sin((tilt * Math.PI) / 180);
  const depthPx = tr.d3 ? Math.max(4, Math.min(plot.w, plot.h) * 0.11 * ((o.view3D?.heightPercent ?? 100) / 100) * Math.cos((tilt * Math.PI) / 180)) : 0;
  const maxExpl = Math.max(0, ...rings.flatMap((s) => s.values.map((_, i) => explodeOf(c, s, i))));
  const labelSpace = !thumb && !compact && rings.some((s) => { const l = pieLabels(c, s); return l && (l.position === 'outsideEnd' || l.position === 'bestFit' || !l.position); }) ? 18 : 0;
  const cx = plot.x + plot.w / 2, cy = plot.y + (plot.h - depthPx) / 2;
  const fitR = Math.min(plot.w / 2 - labelSpace * 1.6, (plot.h - depthPx) / 2 / ky - labelSpace) - (thumb ? 0 : 4);
  const R = Math.max(4, (fitR * (tr.d3 ? 0.88 : 1)) / (1 + maxExpl / 100));
  const hole = doughnut ? R * Math.max(0.1, Math.min(0.9, (o.holeSize ?? 52) / 100)) : 0;
  const ringW = (R - hole) / rings.length;
  const start = -Math.PI / 2 + ((o.firstSliceAngle ?? 0) * Math.PI) / 180;
  let grand = 0;
  const tops: ReactNode[] = [], walls: { y: number; el: ReactNode }[] = [], labels: ReactNode[] = [];
  rings.forEach((s, ri) => {
    const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (ri === 0) grand = total;
    if (!total) return;
    const outer = R - ri * ringW, inner = doughnut ? outer - ringW : 0;
    let a = start;
    const slices: Slice[] = [];
    vals.forEach((v, i) => {
      const a1 = a + (v / total) * Math.PI * 2;
      if (v) slices.push({ i, v, a0: a, a1, color: sliceColor(c, s, i), off: (explodeOf(c, s, i) / 100) * R });
      a = a1;
    });
    const lo = pieLabels(c, s);
    slices.forEach((sl) => {
      const mid = (sl.a0 + sl.a1) / 2;
      const hot = c.hot === ri * 10000 + sl.i;
      const off = sl.off + (hot && rings.length === 1 && !tr.d3 ? 5 : 0);
      const ox = Math.cos(mid) * off, oy = Math.sin(mid) * off * ky;
      const pf = pointFormat(s, sl.i);
      const fill = fillCss(c, pf?.fill ?? s.src.fill, sl.color, hot && rings.length > 1 ? 0.8 : 1);
      const stroke = pf?.line ?? s.src.line ? lineCss(pf?.line ?? s.src.line, 'var(--chart-bg)', 1.5) : { stroke: 'var(--chart-bg)', strokeWidth: thumb ? 0.6 : 1.5 };
      const enter = thumb ? undefined : () => { c.setHot(ri * 10000 + sl.i); c.setTip({ x: cx + Math.cos(mid) * (outer - ringW / 2), y: cy + Math.sin(mid) * (outer - ringW / 2) * ky, title: data.categories[sl.i] ?? '', rows: [{ color: sl.color, label: s.name, value: `${fmtv(sl.v, data.fmt)} (${formatNumber(sl.v / total, '0.0%')})` }] }); };
      const scx = cx + ox, scy = cy + oy;
      if (tr.d3) {
        // Front-facing parts of the outer wall: angles with sin > 0 face the viewer.
        const segs: [number, number][] = [];
        for (let k = -1; k <= 2; k++) {
          const lo2 = Math.max(sl.a0, k * Math.PI * 2), hi2 = Math.min(sl.a1, k * Math.PI * 2 + Math.PI);
          if (hi2 > lo2) segs.push([lo2, hi2]);
        }
        segs.forEach(([b0, b1], k) => {
          const P = (ang: number, dy: number) => `${(scx + Math.cos(ang) * outer).toFixed(2)},${(scy + Math.sin(ang) * outer * ky + dy).toFixed(2)}`;
          const d = `M${P(b0, 0)} A${outer},${outer * ky} 0 0 1 ${P(b1, 0)} L${P(b1, depthPx)} A${outer},${outer * ky} 0 0 0 ${P(b0, depthPx)} Z`;
          walls.push({ y: scy + Math.sin((b0 + b1) / 2) * outer * ky + 1000, el: <g key={`w${ri}-${sl.i}-${k}`}><path d={d} style={{ ...fill, stroke: 'none' }} /><path d={d} style={{ fill: '#000', fillOpacity: 0.22, stroke: 'var(--chart-bg)', strokeWidth: 0.6 }} /></g> });
        });
        // Radial walls show when slices are pulled apart.
        if (sl.off > 0 || rings.length > 1) for (const ang of [sl.a0, sl.a1]) {
          const ex = scx + Math.cos(ang) * outer, ey = scy + Math.sin(ang) * outer * ky;
          const d = `M${scx},${scy} L${ex},${ey} L${ex},${ey + depthPx} L${scx},${scy + depthPx} Z`;
          walls.push({ y: Math.min(scy, ey), el: <g key={`r${ri}-${sl.i}-${ang}`}><path d={d} style={{ ...fill, stroke: 'none' }} /><path d={d} style={{ fill: '#000', fillOpacity: 0.32 }} /></g> });
        }
      }
      tops.push(<path key={`s${ri}-${sl.i}`} data-tip="" d={arcPath(scx, scy, outer, inner, sl.a0, sl.a1, ky)} style={{ ...fill, ...stroke }} onPointerEnter={enter} />);
      if (!lo || thumb || compact) return;
      const t = labelText(lo, { value: sl.v, series: s.name, category: data.categories[sl.i] ?? '', percent: sl.v / total, cell: s.cellLabels?.[sl.i] }, lo.numFmt || data.fmt, true);
      if (!t) return;
      const pos = lo.position ?? 'bestFit';
      const lr = inner ? (outer + inner) / 2 : outer * (pos === 'insideEnd' ? 0.8 : 0.64);
      const arcLen = (sl.a1 - sl.a0) * lr;
      const fitsInside = arcLen > textW(t, 11) * 0.7 && (inner ? outer - inner > 14 : true);
      if (pos === 'center' || pos === 'insideEnd' || (pos === 'bestFit' && fitsInside) || inner) {
        if (!fitsInside && pos === 'bestFit') return;
        labels.push(<text className="onfill" key={`sl${ri}-${sl.i}`} x={scx + Math.cos(mid) * lr} y={scy + Math.sin(mid) * lr * ky + 4} textAnchor="middle" style={{ ...onFill, fontSize: 11, fontWeight: 600, pointerEvents: 'none' }}>{t}</text>);
        return;
      }
      // Outside, with a leader line back to the slice.
      const ex = scx + Math.cos(mid) * outer, ey = scy + Math.sin(mid) * outer * ky;
      const lx = scx + Math.cos(mid) * (outer + 14), ly = scy + Math.sin(mid) * (outer + 14) * ky + (Math.sin(mid) > 0 ? depthPx : 0);
      const right = Math.cos(mid) >= 0;
      if (lo.leaderLines !== false) labels.push(<polyline key={`ll${ri}-${sl.i}`} points={`${ex},${ey} ${lx},${ly} ${lx + (right ? 6 : -6)},${ly}`} style={{ fill: 'none', stroke: 'var(--chart-muted)', strokeWidth: 1 }} />);
      labels.push(<text key={`sl${ri}-${sl.i}`} x={lx + (right ? 9 : -9)} y={ly + 4} textAnchor={right ? 'start' : 'end'} style={{ ...ink, fontSize: 11, pointerEvents: 'none' }}>{t}</text>);
    });
  });
  if (!grand) { els.push(<text key="none" x={cx} y={cy} textAnchor="middle" style={{ ...muted, fontSize: 11 }}>{T('No positive values')}</text>); return; }
  walls.sort((a, b) => a.y - b.y).forEach((w) => els.push(w.el));
  els.push(...tops, ...labels);
  if (hole && !thumb && hole > 22 && !tr.d3) els.push(<text key="tot" x={cx} y={cy + 5} textAnchor="middle" style={{ ...ink, fontSize: Math.min(16, hole / 2.6), fontWeight: 600, pointerEvents: 'none' }}>{formatNumber(grand, data.fmt, true)}</text>);
};

/** Which points go to the second plot of a pie of pie. */
export const ofPieSplit = (vals: number[], split: 'pos' | 'val' | 'percent' | undefined, value: number | undefined): boolean[] => {
  const n = vals.length;
  const total = vals.reduce((a, b) => a + b, 0);
  if (split === 'val') return vals.map((v) => v < (value ?? 0));
  if (split === 'percent') return vals.map((v) => (total ? (v / total) * 100 : 0) < (value ?? 10));
  const k = Math.max(1, Math.min(n - 1, Math.round(value ?? Math.max(1, Math.round(n / 3)))));
  return vals.map((_, i) => i >= n - k);
};

export const drawOfPie = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els, tr } = c;
  c.elementTips = true;
  const o = chart.options ?? {};
  const s = data.series[0]!;
  const vals = s.values.map((v) => (v !== null && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (!total) { els.push(<text key="none" x={plot.x + plot.w / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: 11 }}>{T('No positive values')}</text>); return; }
  const second = ofPieSplit(vals, o.ofPie?.split, o.ofPie?.splitValue);
  if (!second.some(Boolean) || second.every(Boolean)) second.fill(false).fill(true, Math.max(1, vals.length - 1));
  const otherSum = vals.reduce((a, v, i) => a + (second[i] ? v : 0), 0);
  const size2 = Math.max(0.05, Math.min(2, (o.ofPie?.secondSize ?? 75) / 100));
  const gapF = Math.max(0, Math.min(5, (o.ofPie?.gapWidth ?? 100) / 100));
  const lo = pieLabels(c, s);
  const labelRoom = lo && !thumb && !compact ? 26 : 0;
  // Radius so both plots and the gap fit across the width.
  const R = Math.max(4, Math.min((plot.h - labelRoom * 1.2) / 2, (plot.w - labelRoom * 2) / (2 + 2 * size2 * (tr.ofBar ? 0.5 : 1) + gapF * 0.6)));
  const R2 = R * size2;
  const cx = plot.x + labelRoom + R + (plot.w - labelRoom * 2 - (2 * R + gapF * R * 0.6 + 2 * R2 * (tr.ofBar ? 0.5 : 1))) / 2, cy = plot.y + plot.h / 2;
  const cx2 = cx + R + gapF * R * 0.6 + R2 * (tr.ofBar ? 0.5 : 1);
  const otherColor = seriesColor(vals.length, undefined, c.style);
  // Main pie: first-plot points, then one "Other" slice centred on the right.
  const otherAngle = (otherSum / total) * Math.PI * 2;
  let a = otherAngle / 2;
  const mainStroke = { stroke: 'var(--chart-bg)', strokeWidth: thumb ? 0.6 : 1.5 };
  const label = (key: string, x: number, y: number, v: number, cat: string, share: number, inside: boolean) => {
    if (!lo || thumb || compact) return;
    const t = labelText(lo, { value: v, series: s.name, category: cat, percent: share }, lo.numFmt || data.fmt, true);
    if (t) els.push(<text key={key} className={inside ? 'onfill' : undefined} x={x} y={y} textAnchor="middle" style={{ ...(inside ? onFill : ink), fontSize: 10.5, fontWeight: 600, pointerEvents: 'none' }}>{t}</text>);
  };
  vals.forEach((v, i) => {
    if (second[i] || !v) return;
    const a1 = a + (v / total) * Math.PI * 2;
    const color = sliceColor(c, s, i), mid = (a + a1) / 2;
    const off = (explodeOf(c, s, i) / 100) * R;
    const ox = Math.cos(mid) * off, oy = Math.sin(mid) * off;
    const enter = thumb ? undefined : () => { c.setHot(i); c.setTip({ x: cx + Math.cos(mid) * R * 0.6, y: cy + Math.sin(mid) * R * 0.6, title: data.categories[i] ?? '', rows: [{ color, label: s.name, value: `${fmtv(v, data.fmt)} (${formatNumber(v / total, '0.0%')})` }] }); };
    els.push(<path key={`m${i}`} data-tip="" d={arcPath(cx + ox, cy + oy, R, 0, a, a1)} style={{ ...fillCss(c, pointFormat(s, i)?.fill ?? s.src.fill, color, c.hot === i ? 0.82 : 1), ...mainStroke }} onPointerEnter={enter} />);
    if ((a1 - a) * R * 0.64 > 24) label(`ml${i}`, cx + ox + Math.cos(mid) * R * 0.64, cy + oy + Math.sin(mid) * R * 0.64 + 4, v, data.categories[i] ?? '', v / total, true);
    a = a1;
  });
  const oa0 = -otherAngle / 2, oa1 = otherAngle / 2;
  els.push(<path key="other" data-tip="" d={arcPath(cx, cy, R, 0, oa0, oa1)} style={{ fill: otherColor, ...mainStroke }}
    onPointerEnter={thumb ? undefined : () => { c.setHot(-1); c.setTip({ x: cx + R * 0.6, y: cy, title: N('Other'), rows: [{ color: otherColor, label: s.name, value: `${fmtv(otherSum, data.fmt)} (${formatNumber(otherSum / total, '0.0%')})` }] }); }} />);
  if (otherAngle * R * 0.64 > 24) label('ol', cx + R * 0.64, cy + 4, otherSum, T('Other'), otherSum / total, true);
  const e0: [number, number] = [cx + Math.cos(oa0) * R, cy + Math.sin(oa0) * R], e1: [number, number] = [cx + Math.cos(oa1) * R, cy + Math.sin(oa1) * R];
  const lineSt = { stroke: 'var(--chart-muted)', strokeWidth: 1 };
  const idx = vals.map((_, i) => i).filter((i) => second[i] && vals[i]! > 0);
  if (tr.ofBar) {
    const bw = R2 * 0.55, bh = R2 * 2, bx = cx2 - bw / 2, by = cy - R2;
    let y = by + bh;
    idx.forEach((i) => {
      const h = (vals[i]! / otherSum) * bh;
      const color = sliceColor(c, s, i);
      y -= h;
      const yy = y;
      els.push(<rect key={`b${i}`} data-tip="" x={bx} y={yy} width={bw} height={Math.max(0.5, h)} style={{ fill: color, stroke: 'var(--chart-bg)', strokeWidth: 1 }}
        onPointerEnter={thumb ? undefined : () => { c.setHot(i); c.setTip({ x: bx + bw, y: yy + h / 2, title: data.categories[i] ?? '', rows: [{ color, label: s.name, value: `${fmtv(vals[i]!, data.fmt)} (${formatNumber(vals[i]! / total, '0.0%')})` }] }); }} />);
      if (h > 13) label(`bl${i}`, bx + bw + 22, yy + h / 2 + 4, vals[i]!, data.categories[i] ?? '', vals[i]! / total, false);
    });
    if (!thumb) {
      els.push(<line key="c0" x1={e0[0]} y1={e0[1]} x2={bx} y2={by} style={lineSt} />);
      els.push(<line key="c1" x1={e1[0]} y1={e1[1]} x2={bx} y2={by + bh} style={lineSt} />);
    }
    return;
  }
  // Second pie, its slices centred on the left so they face the first pie.
  const sweep = Math.PI * 2;
  let b = 0;
  idx.forEach((i) => {
    const b1 = b + (vals[i]! / otherSum) * sweep;
    const color = sliceColor(c, s, i), mid = (b + b1) / 2;
    els.push(<path key={`p2${i}`} data-tip="" d={arcPath(cx2, cy, R2, 0, b, b1)} style={{ ...fillCss(c, pointFormat(s, i)?.fill ?? s.src.fill, color, c.hot === i ? 0.82 : 1), ...mainStroke }}
      onPointerEnter={thumb ? undefined : () => { c.setHot(i); c.setTip({ x: cx2 + Math.cos(mid) * R2 * 0.6, y: cy + Math.sin(mid) * R2 * 0.6, title: data.categories[i] ?? '', rows: [{ color, label: s.name, value: `${fmtv(vals[i]!, data.fmt)} (${formatNumber(vals[i]! / total, '0.0%')})` }] }); }} />);
    if ((b1 - b) * R2 * 0.62 > 22) label(`p2l${i}`, cx2 + Math.cos(mid) * R2 * 0.62, cy + Math.sin(mid) * R2 * 0.62 + 4, vals[i]!, data.categories[i] ?? '', vals[i]! / total, true);
    b = b1;
  });
  if (!thumb) {
    els.push(<line key="c0" x1={e0[0]} y1={e0[1]} x2={cx2} y2={cy - R2} style={lineSt} />);
    els.push(<line key="c1" x1={e1[0]} y1={e1[1]} x2={cx2} y2={cy + R2} style={lineSt} />);
  }
};
