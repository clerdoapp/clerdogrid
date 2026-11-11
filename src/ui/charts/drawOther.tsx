// Radar, funnel, treemap and sunburst.
import { formatNumber, seriesColor } from './chartData';
import { buildTree, squarify, sunburstArcs, treeDepth, type TreeNode } from './chartStats';
import { axisOf, makeValueScale } from './axes';
import { arcPath } from './drawPie';
import { FONT, fillCss, fmtv, GRID, ink, lineCss, markerEl, muted, onFill, textW, trunc, type DrawCtx } from './svgKit';
import { N, T } from '../i18n';

const modeText = (mode: 'value' | 'percent' | 'both', v: number, share: number, fmt: string): string =>
  mode === 'value' ? formatNumber(v, fmt, true) : mode === 'percent' ? formatNumber(share, '0%') : `${formatNumber(v, fmt, true)} (${formatNumber(share, '0%')})`;

export const drawRadar = (c: DrawCtx): void => {
  const { data, tr, plot, thumb, compact, els, chart } = c;
  const n = Math.max(3, data.categories.length);
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sc = makeValueScale(Math.min(0, ...vals), Math.max(...vals), thumb ? 3 : 5, { ...axisOf(chart, 'y'), logBase: undefined, reverse: false });
  const cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2;
  const labelRoom = thumb ? 2 : compact ? 10 : Math.min(60, Math.max(22, Math.max(0, ...data.categories.map((t) => textW(t))) * 0.6));
  const R = Math.max(6, Math.min(plot.w / 2 - labelRoom, plot.h / 2 - (thumb ? 2 : 16)));
  const ang = (i: number) => -Math.PI / 2 + (i / n) * Math.PI * 2;
  const rad = (v: number) => Math.max(0, sc.pos(v)) * R;
  const pt = (i: number, v: number): [number, number] => [cx + Math.cos(ang(i)) * rad(v), cy + Math.sin(ang(i)) * rad(v)];
  const grid = axisOf(chart, 'y').majorGrid !== false;
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
    els.push(<polygon key={`rs${si}`} points={pts.map((p) => p.join(',')).join(' ')} style={{ ...fillCss(c, s.src.fill, s.color, tr.filled ? 0.45 : 0.08), ...lineCss(s.src.line, s.color, (thumb ? 1.2 : 2) * c.lineScale), strokeLinejoin: 'round' }} />);
    const mk = s.src.marker;
    const markersOn = (tr.markers && mk?.symbol !== 'none') || (!!mk?.symbol && mk.symbol !== 'none' && mk.symbol !== 'auto');
    if (markersOn || (c.hot !== null && !thumb)) pts.forEach((p, i) => (markersOn || c.hot === i) && els.push(markerEl(`rm${si}-${i}`, mk?.symbol && mk.symbol !== 'auto' && mk.symbol !== 'none' ? mk.symbol : 'circle', p[0], p[1], thumb ? 2.8 : c.hot === i ? 9 : mk?.size ? mk.size * 1.2 : compact ? 4 : 6, mk?.fill ?? 'var(--chart-bg)', mk?.line ?? s.color, 2)));
    if (!thumb && !compact && chart.dataLabels) s.values.forEach((v, i) => v !== null && els.push(<text key={`rl${si}-${i}`} x={pt(i, v)[0]} y={pt(i, v)[1] - 6} textAnchor="middle" style={{ ...ink, fontSize: 10 }}>{fmtv(v, data.fmt)}</text>));
  });
  if (thumb) return;
  c.hover.onMove = (x, y) => {
    if ((x - cx) ** 2 + (y - cy) ** 2 > (R + 16) ** 2) { c.setTip(null); c.setHot(null); return; }
    let a = Math.atan2(y - cy, x - cx) + Math.PI / 2;
    if (a < 0) a += Math.PI * 2;
    const i = Math.round((a / (Math.PI * 2)) * n) % n;
    c.setHot(i);
    c.setTip({ x, y, title: data.categories[i] ?? '', rows: data.series.map((s) => ({ color: s.color, label: s.name, value: s.values[i] === null || s.values[i] === undefined ? T('Blank') : fmtv(s.values[i]!, data.fmt) })) });
  };
};

// ---------- Funnel ----------

export const drawFunnel = (c: DrawCtx): void => {
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
      onPointerEnter={thumb ? undefined : () => { c.setHot(k); c.setTip({ x: x0 + w / 2, y: y + h / 2, title: r.label, rows: [{ color, label: s.name, value: fmtv(r.v, data.fmt) }, { color, label: N('Of first'), value: formatNumber(r.v / first, '0.0%') }] }); }} />);
    if (labelW) els.push(<text key={`fl${k}`} x={x0 - 8} y={y + h / 2 + 4} textAnchor="end" style={{ ...muted, fontSize: FONT }}>{trunc(r.label, labelW - 10)}</text>);
    if (!thumb && h >= 13 && (chart.dataLabels ?? true)) {
      const t = modeText(mode, r.v, r.v / first, data.fmt);
      const inside = textW(t, 11) + 8 < bw;
      els.push(<text key={`fv${k}`} x={inside ? x0 + w / 2 : x + bw + 4} y={y + h / 2 + 4} textAnchor={inside ? 'middle' : 'start'} className={inside ? 'onfill' : undefined} style={{ ...(inside ? onFill : ink), fontSize: 11, fontWeight: 600, pointerEvents: 'none' }}>{t}</text>);
    }
  });
};

// ---------- Treemap ----------

export const drawTreemap = (c: DrawCtx): void => {
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
      onPointerEnter={thumb ? undefined : () => { c.setHot(id); c.setTip({ x: x + w / 2, y: y + h / 2, title: n.path.join(' / '), rows: [{ color, label: data.series[0]!.name, value: fmtv(n.value, data.fmt) }, { color, label: N('Share'), value: formatNumber(n.value / tree.value, '0.0%') }] }); }} />);
    if (labels && w > 34 && h > 18) {
      els.push(<text className="onfill" key={`tn${id}`} x={x + 6} y={y + 15} style={{ ...onFill, fontSize: FONT, fontWeight: 600, pointerEvents: 'none' }}>{trunc(n.name, w - 12)}</text>);
      if (h > 34) els.push(<text className="onfill" key={`tv${id}`} x={x + 6} y={y + 29} style={{ ...onFill, fontSize: 10, opacity: 0.9, pointerEvents: 'none' }}>{trunc(modeText(mode, n.value, n.value / tree.value, data.fmt), w - 12, 10)}</text>);
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

export const drawSunburst = (c: DrawCtx): void => {
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
      onPointerEnter={thumb ? undefined : () => { c.setHot(k); c.setTip({ x: cx + Math.cos(mid) * mr, y: cy + Math.sin(mid) * mr, title: a.node.path.join(' / '), rows: [{ color, label: data.series[0]!.name, value: fmtv(a.node.value, data.fmt) }, { color, label: N('Share'), value: formatNumber(a.node.value / tree.value, '0.0%') }] }); }} />);
    if (labels && (a.a1 - a.a0) * mr > textW(a.node.name) + 6 && ring > 14) {
      els.push(<text key={`sbl${k}`} x={cx + Math.cos(mid) * mr} y={cy + Math.sin(mid) * mr + 4} textAnchor="middle" className={a.depth === 1 ? 'onfill' : undefined} style={{ ...(a.depth === 1 ? onFill : ink), fontSize: 10.5, fontWeight: 600, pointerEvents: 'none' }}>{a.node.name}</text>);
    }
  });
};
