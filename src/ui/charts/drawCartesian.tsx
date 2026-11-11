// Column, bar, line, area, combos, stock, histogram, Pareto, waterfall, box and whisker, scatter and bubble.
import type { ReactNode } from 'react';
import type { AxisOptions, ErrorBars, Trendline } from '../../model/workbook';
import { formatNumber, seriesColor, valueExtent, type ResolvedSeries } from './chartData';
import { boxStats, guessTotals, histogramBins, linePath, paretoRows, waterfallSteps, type BoxStats } from './chartStats';
import type { SeriesKind } from './chartTypes';
import { axisOf, buildFrame, dateCategories, slotAt, type Frame, type TableRow } from './axes';
import { errorRanges, fitTrend, sampleFit, trendName } from './trend';
import {
  AXIS, fillCss, fmtv, ink, labelOpts, labelText, lineCss, markerEl, onFill, pointFormat, ptPx, richTextEl, textW, WF,
  type DrawCtx, type TipRow,
} from './svgKit';
import { N, T } from '../i18n';

export interface Seg { i: number; v0: number; v1: number; raw: number | null }
export interface CSeries {
  name: string; color: string; kind: SeriesKind; secondary: boolean; segs: Seg[]; colors?: string[]; markers: boolean; smooth: boolean; stacked: boolean;
  label?: (g: Seg) => string;
  /** The resolved series this was built from (formatting, labels, trendlines). */
  src?: ResolvedSeries;
  lineHidden?: boolean;
  marker?: string;
}
export interface CModel {
  cats: string[]; series: CSeries[]; horizontal: boolean; fmt: string; fmt2: string;
  lo: number; hi: number; second: { lo: number; hi: number } | null; fixed2?: boolean; fixedPrimary?: boolean;
  gap: number; connectors?: boolean; tip: (i: number) => TipRow[];
  /** Percent stacking: labels show shares by default. */
  percent?: boolean;
  /** Elements forced on by the chart type (stock). */
  hiLow?: boolean; upDown?: boolean;
  /** Use the date axis when the categories are dates. */
  dates?: boolean;
}

export const stackSegs = (vals: (number | null)[][], n: number, mode: 'none' | 'stacked' | 'percent'): Seg[][] => {
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

export const cartesianModel = (c: DrawCtx): CModel => {
  const { chart, data, tr, style } = c;
  const fmt = data.fmt;
  if (tr.family === 'histogram') {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
    const bins = histogramBins(all, { binCount: chart.options?.binCount, binWidth: chart.options?.binWidth }, (v) => formatNumber(v, fmt, true));
    const cats = bins.map((b) => b.label);
    const inBin = (v: number, bi: number) => (v > bins[bi]!.lo || (bi === 0 && v >= bins[bi]!.lo)) && (v <= bins[bi]!.hi || bi === bins.length - 1);
    const series: CSeries[] = data.series.map((s) => {
      const own = bins.map((_, bi) => s.values.filter((v): v is number => v !== null && inBin(v, bi)).length);
      return { name: s.name, color: s.color, kind: 'column', secondary: false, segs: own.map((v, i) => ({ i, v0: 0, v1: v, raw: v })), markers: false, smooth: false, stacked: false, src: s };
    });
    if (data.series.length === 1) series[0]!.segs = bins.map((b, i) => ({ i, v0: 0, v1: b.count, raw: b.count }));
    const hi = Math.max(1, ...series.flatMap((s) => s.segs.map((g) => g.v1)));
    return { cats, series, horizontal: false, fmt: '0', fmt2: '0', lo: 0, hi, second: null, gap: series.length === 1 ? 0.97 : 0.86, tip: (i) => series.map((s) => ({ color: s.color, label: s.name, value: String(s.segs[i]?.raw ?? 0) })) };
  }
  if (tr.family === 'pareto') {
    const s = data.series[0]!;
    const rows = paretoRows(data.categories, s.values);
    const lineColor = seriesColor(1, undefined, style);
    const bars: CSeries = { name: s.name, color: s.color, kind: 'column', secondary: false, segs: rows.map((r, i) => ({ i, v0: 0, v1: r.value, raw: r.value })), markers: false, smooth: false, stacked: false, src: s };
    const line: CSeries = { name: T('Cumulative %'), color: lineColor, kind: 'line', secondary: true, segs: rows.map((r, i) => ({ i, v0: 0, v1: r.cumulative, raw: r.cumulative })), markers: true, smooth: false, stacked: false, label: (g) => formatNumber(g.v1, '0%') };
    return {
      cats: rows.map((r) => r.label), series: [bars, line], horizontal: false, fmt, fmt2: '0%', lo: 0, hi: Math.max(0, ...rows.map((r) => r.value)), second: { lo: 0, hi: 1 }, fixed2: true, gap: 0.8,
      tip: (i) => [{ color: s.color, label: s.name, value: fmtv(rows[i]?.value ?? 0, fmt) }, { color: lineColor, label: N('Cumulative'), value: formatNumber(rows[i]?.cumulative ?? 0, '0.0%') }],
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
        return st.kind === 'total' ? [{ color: WF.total, label: N('Total'), value: fmtv(st.end, fmt) }]
          : [{ color: WF[st.kind], label: st.kind === 'up' ? N('Increase') : N('Decrease'), value: fmtv(st.value, fmt) }, { color: 'var(--chart-axis)', label: N('Running total'), value: fmtv(st.end, fmt) }];
      },
    };
  }
  const n = Math.max(1, data.categories.length);
  if (tr.family === 'stock') {
    const vol = tr.volume && data.series.length > 1 ? data.series[0]! : null;
    const prices = vol ? data.series.slice(1) : data.series;
    const last = prices.length - 1;
    const series: CSeries[] = [];
    if (vol) series.push({ name: vol.name, color: vol.color, kind: 'column', secondary: false, segs: vol.values.map((v, i) => ({ i, v0: 0, v1: v ?? 0, raw: v })), markers: false, smooth: false, stacked: false, src: vol });
    prices.forEach((s, k) => {
      const ownLine = !!s.src.line && !s.src.line.none;
      series.push({
        name: s.name, color: s.color, kind: 'line', secondary: !!vol, segs: s.values.map((v, i) => ({ i, v0: 0, v1: v ?? 0, raw: v })),
        markers: !!s.src.marker || (!tr.open && k === last), marker: s.src.marker?.symbol ?? (!tr.open && k === last ? 'dash' : 'none'), smooth: false, stacked: false, src: s, lineHidden: !ownLine,
      });
    });
    const pv = prices.flatMap((s) => s.values.filter((v): v is number => v !== null));
    const vv = vol ? vol.values.filter((v): v is number => v !== null) : [];
    const [plo, phi] = valueExtent(pv, false);
    return {
      cats: data.categories, series, horizontal: false, fmt, fmt2: fmt,
      lo: vol ? 0 : plo, hi: vol ? Math.max(1, ...vv) : phi, second: vol ? { lo: plo, hi: phi } : null, gap: 0.6, hiLow: true, upDown: tr.open, dates: true,
      tip: (i) => data.series.map((s) => ({ color: s.color, label: s.name, value: s.values[i] === null || s.values[i] === undefined ? T('Blank') : fmtv(s.values[i]!, fmt) })),
    };
  }
  // Plain, stacked and combo category charts. Each kind of series stacks within its own group.
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
    markers: tr.markers || (tr.combo && s.type === 'line' && n <= 40) || (!!s.src.marker && s.src.marker.symbol !== 'none'), smooth: !!s.src.smooth, src: s,
    marker: s.src.marker?.symbol, lineHidden: !!s.src.line?.none,
  }));
  const ext = (sec: boolean): [number, number] | null => {
    const list = series.filter((s) => s.secondary === sec);
    if (!list.length) return null;
    const vals = list.flatMap((s) => s.segs.filter((g) => g.raw !== null || s.stacked).flatMap((g) => (s.stacked ? [g.v0, g.v1] : [g.v1])));
    if (percent) return [Math.min(0, ...vals), Math.max(0, ...vals)];
    return valueExtent(vals, list.some((s) => s.kind !== 'line'));
  };
  const p = ext(false) ?? ext(true) ?? [0, 1], s2 = series.some((s) => !s.secondary) ? ext(true) : null;
  if (!series.some((s) => !s.secondary)) series.forEach((s) => { s.secondary = false; });
  return {
    cats: data.categories, series, horizontal: tr.horizontal, fmt: percent ? '0%' : fmt, fmt2: fmt, lo: p[0], hi: p[1], second: s2 ? { lo: s2[0], hi: s2[1] } : null,
    gap: tr.stack !== 'none' ? 0.62 : 0.74, fixedPrimary: false, percent, dates: true,
    tip: (i) => data.series.map((s, si) => {
      const g = segs[si]![i];
      const v = g?.raw ?? null;
      return { color: s.color, label: s.name, value: v === null ? T('Blank') : percent ? `${fmtv(v, data.fmt)} (${formatNumber(g!.v1 - g!.v0, '0.0%')})` : fmtv(v, data.fmt) };
    }),
  };
};

/** Trendline drawing shared by category and x-y charts. */
const drawTrends = (c: DrawCtx, parts: ReactNode[], key: string, list: Trendline[] | undefined, pts: [number, number][], toScreen: (x: number, y: number) => [number, number], color: string, seriesName: string, xRange: [number, number]): void => {
  if (!list?.length || pts.length < 2) return;
  list.forEach((t, ti) => {
    const fit = fitTrend(pts, t);
    if (!fit) return;
    const x0 = xRange[0] - (t.backward ?? 0), x1 = xRange[1] + (t.forward ?? 0);
    const line = sampleFit(fit, x0, x1).map(([x, y]) => toScreen(x, y)).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
    if (line.length < 2) return;
    const st = lineCss(t.line, color, 1.6 * c.lineScale);
    parts.push(<path key={`${key}t${ti}`} d={linePath(line, false)} style={{ fill: 'none', ...st, strokeDasharray: t.line?.dash ? st.strokeDasharray : '5 3', strokeLinecap: 'round' }}><title>{trendName(t, seriesName)}</title></path>);
    if (c.thumb || c.compact || (!t.showEquation && !t.showR2)) return;
    const runs: { t: string; sup?: boolean }[] = [];
    if (t.showEquation) runs.push(...fit.equation);
    const [ex, ey] = line.at(-1)!;
    c.els.push(richTextEl(`${key}te${ti}`, runs, Math.min(ex, c.plot.x + c.plot.w - 4), ey - 8 - (t.showR2 && t.showEquation ? 13 : 0), 'end'));
    if (t.showR2 && Number.isFinite(fit.r2)) c.els.push(richTextEl(`${key}tr${ti}`, [{ t: 'R' }, { t: '2', sup: true }, { t: ` = ${(+fit.r2.toFixed(4)).toString()}` }], Math.min(ex, c.plot.x + c.plot.w - 4), ey - 8, 'end'));
  });
};

/** Error bars along the value direction (and x for x-y charts). */
const drawErrors = (parts: ReactNode[], key: string, s: ResolvedSeries, list: ErrorBars[] | undefined, values: (number | null)[], at: (i: number, v: number) => [number, number], vertical: boolean, dir: 'x' | 'y'): void => {
  list?.forEach((e, ei) => {
    if ((e.dir ?? 'y') !== dir) return;
    const ranges = errorRanges(values, e, s.errors?.[ei]);
    const st = lineCss(e.line, 'var(--chart-ink)', 1.1);
    const cap = e.noCap ? 0 : 4;
    ranges.forEach((r, i) => {
      if (!r) return;
      const [a0, a1] = at(i, r[0]), [b0, b1] = at(i, r[1]);
      parts.push(<line key={`${key}e${ei}-${i}`} x1={a0} y1={a1} x2={b0} y2={b1} style={st} />);
      if (!cap) return;
      for (const [x, y] of [[a0, a1], [b0, b1]] as const) parts.push(vertical
        ? <line key={`${key}ec${ei}-${i}-${x}-${y}`} x1={x - cap} x2={x + cap} y1={y} y2={y} style={st} />
        : <line key={`${key}ec${ei}-${i}-${x}-${y}`} x1={x} x2={x} y1={y - cap} y2={y + cap} style={st} />);
    });
  });
};

export const drawCartesian = (c: DrawCtx, m: CModel): void => {
  const { chart, data, thumb, els } = c;
  const { horizontal } = m;
  const o = chart.options ?? {};
  const catAx: AxisOptions = axisOf(chart, 'x');
  const yAx: AxisOptions = axisOf(chart, 'y');
  // Counts get whole-number ticks.
  if (m.fmt === '0' && yAx.majorUnit === undefined && m.hi <= 8) yAx.majorUnit = 1;
  const y2Ax: AxisOptions = m.fixed2 ? { min: m.second!.lo, max: m.second!.hi, majorUnit: (m.second!.hi - m.second!.lo) / 4, numFmt: m.fmt2 } : axisOf(chart, 'y2');
  const date = m.dates ? dateCategories(data, catAx) : null;
  const tableOn = !!o.dataTable && !thumb && !c.compact;
  const table: TableRow[] | null = tableOn ? m.series.map((s) => ({ name: s.name, color: s.color, cells: m.cats.map((_, i) => { const g = s.segs.find((x) => x.i === i); return g?.raw === null || g?.raw === undefined ? '' : fmtv(g.raw, s.secondary ? m.fmt2 : m.fmt); }) })) : null;
  const f = buildFrame(c, {
    cats: m.cats, date: tableOn ? null : date, catAx, horizontal,
    y: { lo: m.lo, hi: m.hi, fmt: yAx.numFmt || m.fmt, ax: yAx, title: data.axisTitles.y },
    y2: m.second ? { lo: m.second.lo, hi: m.second.hi, fmt: y2Ax.numFmt || m.fmt2, ax: y2Ax, title: data.axisTitles.y2 } : null,
    table, tableOpts: o.dataTable, xTitle: data.axisTitles.x,
  });
  const { catPos } = f;
  const band = f.band;
  const frameClip = `${c.uid}fr`;
  c.defs.push(<clipPath key={frameClip} id={frameClip}><rect x={f.px - 8} y={f.py - 8} width={f.pw + 16} height={f.ph + 16} /></clipPath>);
  const pos = (s: CSeries, v: number) => (s.secondary ? f.valPos2(v) : f.valPos(v));
  const base = (s: CSeries) => (s.secondary ? f.zero2 : f.zero);
  const labelsAllowed = !thumb && !c.compact && f.n * m.series.length <= 400;
  const parts: ReactNode[] = [];
  const ls = c.lineScale;
  const span = o.blanks === 'span';
  const xy = (s: CSeries, i: number, v: number): [number, number] => (horizontal ? [pos(s, v), catPos(i)] : [catPos(i), pos(s, v)]);
  const labelFor = (s: CSeries, g: Seg): string | null => {
    if (!s.src) return chart.dataLabels && s.label ? s.label(g) : null;
    const lo = labelOpts(chart, s.src);
    if (!lo) return null;
    if (s.label && !s.src.src.labels && !chart.options?.labels) return s.label(g);
    const share = m.percent ? g.v1 - g.v0 : undefined;
    return labelText(lo, { value: g.raw, series: s.name, category: m.cats[g.i] ?? '', percent: share, cell: s.src.cellLabels?.[g.i] }, s.secondary ? m.fmt2 : (lo.numFmt || data.fmt), !!m.percent);
  };

  // Areas first, so bars and lines sit on top.
  m.series.forEach((s, si) => {
    if (s.kind !== 'area') return;
    const usable = s.stacked || span ? s.segs.filter((g) => s.stacked || g.raw !== null) : s.segs.filter((g) => g.raw !== null);
    if (!usable.length) return;
    const pt = (g: Seg, v: number) => xy(s, g.i, v).join(',');
    const upper = usable.map((g) => pt(g, g.v1));
    const lower = [...usable].reverse().map((g) => (s.stacked ? pt(g, g.v0) : (horizontal ? `${base(s)},${catPos(g.i)}` : `${catPos(g.i)},${base(s)}`)));
    const fill = fillCss(c, s.src?.src.fill, s.color, s.stacked ? 0.82 : 0.38);
    parts.push(<polygon key={`a${si}`} points={[...upper, ...lower].join(' ')} style={fill} />);
    parts.push(<polyline key={`al${si}`} points={upper.join(' ')} style={{ fill: 'none', ...lineCss(s.src?.src.line, s.color, 1.6 * ls), strokeLinejoin: 'round' }} />);
    if (!labelsAllowed) return;
    usable.forEach((g) => { const t = labelFor(s, g); if (t) { const [x, y] = xy(s, g.i, s.stacked ? (g.v0 + g.v1) / 2 : g.v1); els.push(<text key={`l${si}-${g.i}`} x={x} y={y + (s.stacked ? 3.5 : -6)} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{t}</text>); } });
  });

  // Bars: width from the gap and overlap settings.
  const bars = m.series.map((s, si) => ({ s, si })).filter((x) => x.s.kind === 'column');
  const stackedBars = bars.length > 0 && bars.every((b) => b.s.stacked);
  const slots = stackedBars ? 1 : bars.length;
  let w: number, step: number;
  if (o.gapWidth !== undefined || o.overlap !== undefined) {
    const gw = (o.gapWidth ?? 150) / 100, ov = (o.overlap ?? (stackedBars ? 100 : 0)) / 100;
    const k = stackedBars ? 1 : slots - (slots - 1) * ov;
    w = band / (k + gw);
    step = stackedBars ? 0 : w * (1 - ov);
  } else {
    w = (band * m.gap) / Math.max(1, slots);
    step = stackedBars ? 0 : w;
  }
  const group = w + step * (Math.max(1, slots) - 1);
  const vary = (o.varyColors ?? false) && m.series.length === 1;
  bars.forEach(({ s, si }, j) => {
    const src = s.src;
    s.segs.forEach((g) => {
      if (g.raw === null && !s.stacked) return;
      const off = -group / 2 + (stackedBars ? 0 : j) * step;
      const b = pos(s, g.v1);
      const a0 = s.stacked ? pos(s, g.v0) : base(s);
      const cp = catPos(g.i) + off;
      const gapPx = w > 6 ? 1 : 0.4;
      const r = horizontal ? { x: Math.min(a0, b), y: cp, width: Math.abs(b - a0), height: Math.max(0.5, w - gapPx) } : { x: cp, y: Math.min(a0, b), width: Math.max(0.5, w - gapPx), height: Math.abs(b - a0) };
      const pf = src ? pointFormat(src, g.i) : undefined;
      const color = s.colors?.[g.i] ?? (vary ? seriesColor(g.i, undefined, c.style) : s.color);
      const neg = (g.raw ?? 0) < 0 && (src?.src.invertIfNegative ?? false);
      const fill = neg ? { fill: 'var(--chart-bg)' } : fillCss(c, pf?.fill ?? src?.src.fill, pf?.fill?.color ?? color);
      const stroke = neg ? { stroke: color, strokeWidth: 1.2 } : pf?.line ?? src?.src.line ? lineCss(pf?.line ?? src?.src.line, color, 1) : {};
      parts.push(<rect key={`b${si}-${g.i}`} {...r} rx={Math.min(2, w / 6)} style={{ ...fill, ...stroke }} />);
      if (!labelsAllowed || g.raw === null) return;
      const t = labelFor(s, g);
      if (!t) return;
      const lp = (src?.src.labels ?? o.labels)?.position ?? (s.stacked && !m.connectors ? 'center' : 'outsideEnd');
      const up = g.v1 >= g.v0;
      const fits = (horizontal ? r.width : r.height) > 14 && textW(t, 10) < (horizontal ? r.width : w) + 2;
      if ((lp === 'center' || lp === 'insideEnd' || lp === 'insideBase') && (fits || s.stacked)) {
        if (!fits) return;
        let lx = r.x + r.width / 2, ly = r.y + r.height / 2 + 3.5;
        if (lp !== 'center') {
          const atEnd = lp === 'insideEnd';
          if (horizontal) lx = (up === atEnd) ? r.x + r.width - textW(t, 10) / 2 - 4 : r.x + textW(t, 10) / 2 + 4;
          else ly = (up === atEnd) ? r.y + 13 : r.y + r.height - 5;
        }
        els.push(<text className="onfill" key={`l${si}-${g.i}`} x={lx} y={ly} textAnchor="middle" style={{ ...onFill, fontSize: 10 }}>{t}</text>);
      } else if (horizontal) els.push(<text key={`l${si}-${g.i}`} x={up ? Math.max(a0, b) + 4 : Math.min(a0, b) - 4} y={cp + w / 2 + 3.5} textAnchor={up ? 'start' : 'end'} style={{ ...ink, fontSize: 10 }}>{t}</text>);
      else els.push(<text key={`l${si}-${g.i}`} x={cp + w / 2} y={up ? Math.min(a0, b) - 4 : Math.max(a0, b) + 12} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{t}</text>);
    });
    if (m.connectors) s.segs.forEach((g, k) => {
      const next = s.segs[k + 1];
      if (!next) return;
      const y = pos(s, g.v1);
      parts.push(<line key={`cn${k}`} x1={catPos(g.i) + group / 2} x2={catPos(next.i) - group / 2} y1={y} y2={y} style={{ stroke: 'var(--chart-axis)', strokeWidth: 1, strokeDasharray: '3 2' }} />);
    });
    // Series lines join the ends of each stacked segment across categories.
    if (o.seriesLines && s.stacked && !m.connectors) s.segs.forEach((g, k) => {
      const next = s.segs[k + 1];
      if (!next) return;
      const a = pos(s, g.v1), b = pos(s, next.v1);
      parts.push(horizontal
        ? <line key={`sl${si}-${k}`} x1={a} x2={b} y1={catPos(g.i) + w / 2} y2={catPos(next.i) - w / 2} style={{ ...AXIS }} />
        : <line key={`sl${si}-${k}`} x1={catPos(g.i) + w / 2} x2={catPos(next.i) - w / 2} y1={a} y2={b} style={{ ...AXIS }} />);
    });
  });

  // High-low lines and up-down bars between line series.
  const lineSeries = m.series.filter((s) => s.kind === 'line');
  if ((o.hiLowLines ?? m.hiLow) && lineSeries.length > 1) for (let i = 0; i < m.cats.length; i++) {
    const vals = lineSeries.map((s) => s.segs[i]).filter((g): g is Seg => !!g && g.raw !== null);
    if (vals.length < 2) continue;
    const s0 = lineSeries[0]!;
    const hi = Math.max(...vals.map((g) => g.v1)), lo = Math.min(...vals.map((g) => g.v1));
    const [x1, y1] = xy(s0, i, hi), [x2, y2] = xy(s0, i, lo);
    parts.push(<line key={`hl${i}`} x1={x1} y1={y1} x2={x2} y2={y2} style={{ stroke: 'var(--chart-ink)', strokeWidth: 1.2 * ls }} />);
  }
  if ((o.upDownBars ?? m.upDown) && lineSeries.length > 1) {
    const first = lineSeries[0]!, last = lineSeries.at(-1)!;
    const bw = Math.max(2, band * 0.45);
    first.segs.forEach((g) => {
      const e = last.segs[g.i];
      if (g.raw === null || !e || e.raw === null) return;
      const up = e.v1 >= g.v1;
      const a = pos(first, g.v1), b = pos(first, e.v1);
      const fill = up ? o.upFill ?? 'var(--chart-up)' : o.downFill ?? 'var(--chart-down)';
      parts.push(horizontal
        ? <rect key={`ud${g.i}`} x={Math.min(a, b)} y={catPos(g.i) - bw / 2} width={Math.max(1, Math.abs(b - a))} height={bw} style={{ fill, stroke: 'var(--chart-ink)', strokeWidth: 0.8 }} />
        : <rect key={`ud${g.i}`} x={catPos(g.i) - bw / 2} y={Math.min(a, b)} width={bw} height={Math.max(1, Math.abs(b - a))} style={{ fill, stroke: 'var(--chart-ink)', strokeWidth: 0.8 }} />);
    });
  }

  m.series.forEach((s, si) => {
    if (s.kind !== 'line') return;
    const src = s.src;
    const runs: [number, number][][] = [[]];
    s.segs.forEach((g) => {
      if (g.raw === null && !s.stacked) { if (!span && runs.at(-1)!.length) runs.push([]); return; }
      runs.at(-1)!.push(xy(s, g.i, g.v1));
    });
    const stroke = lineCss(src?.src.line, s.color, (thumb ? 1.6 : 2.2) * ls);
    if (!s.lineHidden) runs.forEach((r, k) => r.length > 1 && parts.push(<path key={`ln${si}-${k}`} d={linePath(r, s.smooth)} style={{ fill: 'none', ...stroke, strokeLinejoin: 'round', strokeLinecap: 'round' }} />));
    // Drop lines from each point down to the base axis.
    if (o.dropLines) s.segs.forEach((g) => {
      if (g.raw === null) return;
      const [x, y] = xy(s, g.i, g.v1);
      parts.push(horizontal ? <line key={`dl${si}-${g.i}`} x1={x} x2={base(s)} y1={y} y2={y} style={{ ...AXIS, strokeDasharray: '2 2' }} /> : <line key={`dl${si}-${g.i}`} x1={x} x2={x} y1={y} y2={base(s)} style={{ ...AXIS, strokeDasharray: '2 2' }} />);
    });
    s.segs.forEach((g) => {
      if (g.raw === null && !s.stacked) return;
      const pf = src ? pointFormat(src, g.i) : undefined;
      const single = runs.some((r) => r.length === 1) && !s.lineHidden;
      const show = s.markers || single || c.hot === g.i || !!pf?.marker;
      if (!show) return;
      const mk = pf?.marker ?? src?.src.marker;
      const sym = (mk?.symbol && mk.symbol !== 'auto' ? mk.symbol : s.marker && s.marker !== 'auto' ? s.marker : 'circle') as Parameters<typeof markerEl>[1];
      if (sym === 'none' && c.hot !== g.i) return;
      const [x, y] = xy(s, g.i, g.v1);
      const size = thumb ? 3.2 : c.hot === g.i ? 9 : mk?.size ? Math.max(2, Math.min(72, mk.size)) * 1.2 : c.compact ? 4 : 6;
      parts.push(markerEl(`m${si}-${g.i}`, sym === 'none' ? 'circle' : sym, x, y, size, mk?.fill ?? 'var(--chart-bg)', mk?.line ?? s.color, 2));
    });
    if (labelsAllowed) s.segs.forEach((g) => {
      if (g.raw === null) return;
      const t = labelFor(s, g);
      if (!t) return;
      const [x, y] = xy(s, g.i, g.v1);
      const lp = (src?.src.labels ?? o.labels)?.position ?? 'above';
      const at = lp === 'below' ? [x, y + 15, 'middle'] : lp === 'left' ? [x - 7, y + 3.5, 'end'] : lp === 'right' ? [x + 7, y + 3.5, 'start'] : lp === 'center' ? [x, y + 3.5, 'middle'] : [x, y - 7, 'middle'];
      els.push(<text key={`l${si}-${g.i}`} x={at[0] as number} y={at[1] as number} textAnchor={at[2] as 'middle'} style={{ ...ink, fontSize: 10 }}>{t}</text>);
    });
  });

  // Error bars and trendlines, measured in category numbers along the base axis.
  m.series.forEach((s, si) => {
    const src = s.src;
    if (!src) return;
    const vals = s.segs.map((g) => (g.raw === null ? null : g.v1));
    drawErrors(parts, `s${si}`, src, src.src.errorBars, vals, (i, v) => {
      const off = s.kind === 'column' && !stackedBars ? -group / 2 + bars.findIndex((b) => b.s === s) * step + w / 2 : 0;
      return horizontal ? [pos(s, v), catPos(i) + off] : [catPos(i) + off, pos(s, v)];
    }, !horizontal, 'y');
    if (!src.src.trendlines?.length || s.stacked) return;
    const pts = s.segs.filter((g) => g.raw !== null).map((g): [number, number] => [date ? (date.bucket[g.i] ?? g.i) + 1 : g.i + 1, g.v1]);
    const xmin = Math.min(...pts.map((p) => p[0])), xmax = Math.max(...pts.map((p) => p[0]));
    drawTrends(c, parts, `s${si}`, src.src.trendlines, pts, (x, y) => (horizontal ? [pos(s, y), f.xPos(x)] : [f.xPos(x), pos(s, y)]), s.color, s.name, [xmin, xmax]);
  });
  els.push(<g key="marks" clipPath={`url(#${frameClip})`}>{parts}</g>);

  if (thumb) return;
  c.hover.onMove = (x, y) => {
    const k = slotAt(f, horizontal, !!catAx.reverse, x, y);
    if (k === null) { c.setTip(null); c.setHot(null); return; }
    const i = date ? date.bucket.indexOf(k) : k;
    c.setHot(k);
    if (i < 0) { c.setTip(null); return; }
    c.setTip({ x, y, title: m.cats[i] ?? '', rows: m.tip(i) });
  };
};

// ---------- Box and whisker ----------

export const drawBox = (c: DrawCtx): void => {
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
  const yAx = axisOf(chart, 'y');
  const f = buildFrame(c, { cats, catAx: axisOf(chart, 'x'), horizontal: false, y: { lo: lo >= 0 && lo < hi * 0.4 ? 0 : lo, hi, fmt: yAx.numFmt || data.fmt, ax: yAx, title: data.axisTitles.y }, xTitle: data.axisTitles.x });
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
    parts.push(<rect key={`bx${k}`} x={x0} y={Math.min(Y(st.q3), Y(st.q1))} width={Math.max(1, bw)} height={Math.max(1, Math.abs(Y(st.q1) - Y(st.q3)))} rx={2} style={{ fill: color, fillOpacity: 0.32, stroke: color, strokeWidth: 1.4 }} />);
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
    const ci = slotAt(f, false, !!axisOf(chart, 'x').reverse, x, y);
    if (ci === null) { c.setTip(null); c.setHot(null); return; }
    const si = Math.min(per - 1, Math.max(0, Math.floor((x - (f.catPos(ci) - group / 2)) / w)));
    const st = stats[ci]?.[si];
    c.setHot(ci);
    if (!st) { c.setTip(null); return; }
    const color = grouped ? data.series[si]!.color : data.series[ci]!.color;
    const v = (l: string, nn: number) => ({ color, label: l, value: fmtv(nn, data.fmt) });
    c.setTip({ x, y, title: grouped ? `${cats[ci]} · ${data.series[si]!.name}` : cats[ci] ?? '', rows: [v(T('Maximum'), st.max), v(T('Upper quartile'), st.q3), v(T('Median'), st.median), v(T('Lower quartile'), st.q1), v(T('Minimum'), st.min), v(T('Mean'), st.mean), ...(st.outliers.length ? [{ color, label: N('Outliers'), value: String(st.outliers.length) }] : [])] });
  };
};

// ---------- Scatter and bubble ----------

export const drawScatter = (c: DrawCtx): void => {
  const { chart, data, tr, thumb, compact, els } = c;
  const o = chart.options ?? {};
  const bubble = tr.family === 'bubble';
  type P = { si: number; i: number; x: number; y: number; z: number };
  const pts: P[] = data.series.flatMap((s, si) => s.values.map((y, i) => {
    const xs = s.xs ?? data.xs;
    return { si, i, x: xs ? xs[i] ?? null : i + 1, y, z: bubble ? s.sizes?.[i] ?? 1 : 0 };
  }).filter((p): p is P => p.x !== null && p.y !== null && (!bubble || p.z !== null && (p.z > 0 || (o as { showNegBubbles?: boolean }).showNegBubbles === true))));
  if (!pts.length) return;
  const xAx = axisOf(chart, 'x'), yAx = axisOf(chart, 'y'), y2Ax = axisOf(chart, 'y2');
  const sec = (si: number) => data.series[si]!.secondary;
  const prim = pts.filter((p) => !sec(p.si)), second = pts.filter((p) => sec(p.si));
  const [xl, xh] = valueExtent(pts.map((p) => p.x), false);
  const [yl, yh] = valueExtent((prim.length ? prim : pts).map((p) => p.y), false);
  const zMax = Math.max(...pts.map((p) => Math.abs(p.z)), 1e-9);
  const scale = (o.bubbleScale ?? 100) / 100;
  const rMaxGuess = bubble ? Math.max(3, Math.min(c.plot.w, c.plot.h) * (thumb ? 0.08 : 0.09) * scale) : 0;
  // Leave room for the largest bubbles inside the plot.
  const padX = bubble ? (rMaxGuess / Math.max(10, c.plot.w - 2 * rMaxGuess)) * (xh - xl || 1) : 0, padY = bubble ? (rMaxGuess / Math.max(10, c.plot.h - 2 * rMaxGuess)) * (yh - yl || 1) : 0;
  const s2 = second.length && prim.length ? valueExtent(second.map((p) => p.y), false) : null;
  const f: Frame = buildFrame(c, {
    xVal: { lo: xl - padX, hi: xh + padX, fmt: xAx.numFmt || 'General', ax: xAx, crossMin: !(xl < 0 && xh > 0) }, catAx: xAx, horizontal: false,
    y: { lo: yl - padY, hi: yh + padY, fmt: yAx.numFmt || data.fmt, ax: yAx, title: data.axisTitles.y, crossMin: !(yl < 0 && yh > 0) },
    y2: s2 ? { lo: s2[0], hi: s2[1], fmt: y2Ax.numFmt || data.fmt, ax: y2Ax, title: data.axisTitles.y2 } : null, xTitle: data.axisTitles.x,
  });
  const rMax = bubble ? Math.max(3, Math.min(f.pw, f.ph) * (thumb ? 0.08 : 0.09) * scale) : 0;
  const frameClip = `${c.uid}fr`;
  c.defs.push(<clipPath key={frameClip} id={frameClip}><rect x={f.px - 8} y={f.py - 8} width={f.pw + 16} height={f.ph + 16} /></clipPath>);
  const rOf = (z: number) => Math.max(thumb ? 1.2 : 2.5, Math.sqrt(Math.abs(z) / zMax) * rMax);
  const X = (v: number) => f.xPos(v);
  const Yp = (si: number) => (s2 && sec(si) ? f.valPos2 : f.valPos);
  const parts: ReactNode[] = [];
  const ls = c.lineScale;
  data.series.forEach((s, si) => {
    const own = pts.filter((p) => p.si === si);
    const Y = Yp(si);
    const lineOn = tr.lines ? !s.src.line?.none : !!s.src.line && !s.src.line.none && !!(s.src.line.color || s.src.line.width);
    if (lineOn) {
      const line = own.map((p): [number, number] => [X(p.x), Y(p.y)]);
      if (line.length > 1) parts.push(<path key={`sl${si}`} d={linePath(line, s.src.smooth ?? tr.smooth)} style={{ fill: 'none', ...lineCss(s.src.line, s.color, (thumb ? 1.4 : 2) * ls), strokeLinejoin: 'round', strokeLinecap: 'round' }} />);
    }
    drawErrors(parts, `x${si}`, s, s.src.errorBars, own.map((p) => p.x), (i, v) => [X(v), Y(own[i]!.y)], false, 'x');
    drawErrors(parts, `y${si}`, s, s.src.errorBars, own.map((p) => p.y), (i, v) => [X(own[i]!.x), Y(v)], true, 'y');
    if (s.src.trendlines?.length && own.length > 1) {
      const xsr = own.map((p) => p.x);
      drawTrends(c, parts, `t${si}`, s.src.trendlines, own.map((p) => [p.x, p.y]), (x, y) => [X(x), Y(y)], s.color, s.name, [Math.min(...xsr), Math.max(...xsr)]);
    }
  });
  if (bubble && tr.d3) data.series.forEach((s, si) => {
    const id = `${c.uid}bb${si}`;
    c.defs.push(<radialGradient key={id} id={id} cx="0.4" cy="0.35" r="0.65" fx="0.32" fy="0.28">
      <stop offset="0" style={{ stopColor: '#ffffff', stopOpacity: 0.95 }} /><stop offset="0.45" style={{ stopColor: s.color }} /><stop offset="1" style={{ stopColor: s.color }} />
    </radialGradient>);
  });
  const order = bubble ? pts.map((p, k) => ({ p, k })).sort((a, b) => Math.abs(b.p.z) - Math.abs(a.p.z)) : pts.map((p, k) => ({ p, k }));
  order.forEach(({ p, k }) => {
    const s = data.series[p.si]!;
    const color = s.color;
    const on = c.hot === k;
    const pf = pointFormat(s, p.i);
    const Y = Yp(p.si);
    if (bubble) {
      const fill = tr.d3 ? { fill: `url(#${c.uid}bb${p.si})`, fillOpacity: on ? 1 : 0.92 } : fillCss(c, pf?.fill ?? s.src.fill, color, on ? 0.9 : 0.6);
      parts.push(<circle key={`p${k}`} cx={X(p.x)} cy={Y(p.y)} r={rOf(p.z)} style={{ ...fill, stroke: p.z < 0 ? color : tr.d3 ? 'none' : color, strokeWidth: 1, strokeDasharray: p.z < 0 ? '3 2' : undefined }} />);
    } else {
      const mk = pf?.marker ?? s.src.marker;
      const markersOn = tr.markers || (!!mk && mk.symbol !== 'none' && mk.symbol !== undefined);
      if (markersOn || on) {
        const size = thumb ? 3.6 : on ? 11 : mk?.size ? Math.max(2, Math.min(72, mk.size)) * 1.2 : compact ? 4.8 : 8;
        const sym = mk?.symbol && mk.symbol !== 'auto' && mk.symbol !== 'none' ? mk.symbol : 'circle';
        parts.push(markerEl(`p${k}`, sym, X(p.x), Y(p.y), size, mk?.fill ?? (tr.lines ? 'var(--chart-bg)' : color), mk?.line ?? (tr.lines ? color : 'var(--chart-bg)'), tr.lines ? 2 : 1));
      }
    }
    if (!thumb && !compact && pts.length <= 300) {
      const lo = labelOpts(chart, s);
      if (lo) {
        const t = labelText(lo.value === undefined && !lo.seriesName && !lo.categoryName && !lo.bubbleSize ? { ...lo, value: true } : lo, { value: p.y, series: s.name, category: formatNumber(p.x, 'General'), size: bubble ? p.z : null, cell: s.cellLabels?.[p.i] }, lo.numFmt || data.fmt);
        const lp = (s.src.labels ?? o.labels)?.position ?? (bubble ? 'center' : 'above');
        const dy = lp === 'center' ? 3.5 : lp === 'below' ? 15 : -7;
        const dx = lp === 'left' ? -7 : lp === 'right' ? 7 : 0;
        els.push(<text key={`pl${k}`} x={X(p.x) + dx} y={Y(p.y) + dy} textAnchor={lp === 'left' ? 'end' : lp === 'right' ? 'start' : 'middle'} style={{ ...ink, fontSize: 10 }}>{t}</text>);
      }
    }
  });
  els.push(<g key="marks" clipPath={`url(#${bubble ? c.clip : frameClip})`}>{parts}</g>);
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    let best = -1, bd = Infinity;
    pts.forEach((p, k) => {
      const d = Math.sqrt((X(p.x) - x) ** 2 + (Yp(p.si)(p.y) - y) ** 2) - (bubble ? rOf(p.z) : 0);
      if (d < bd && d < 18) { bd = d; best = k; }
    });
    if (best < 0) { c.setTip(null); c.setHot(null); return; }
    const p = pts[best]!;
    const s = data.series[p.si]!;
    c.setHot(best);
    const rows: TipRow[] = [{ color: s.color, label: data.axisTitles.x || 'X', value: formatNumber(p.x, 'General') }, { color: s.color, label: data.axisTitles.y || 'Y', value: formatNumber(p.y, data.fmt) }];
    if (bubble) rows.push({ color: s.color, label: N('Size'), value: formatNumber(p.z, 'General') });
    c.setTip({ x: X(p.x), y: Yp(p.si)(p.y), title: (s.xs ?? data.xs) ? s.name : `${s.name} · ${data.categories[p.i] ?? ''}`, rows });
  };
  void ptPx;
};
