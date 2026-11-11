// Filled map: each region in the data is coloured by its value.
import type { ReactNode } from 'react';
import { formatNumber } from './chartData';
import { equalEarth, geoData, loadGeo, matchRegions, usProject, type Region } from './geo';
import { fmtv, ink, muted, textW, type DrawCtx } from './svgKit';
import { T as tl } from '../i18n';

const hex = (s: string | undefined): [number, number, number] | null => {
  const m = s && /^#?([0-9a-f]{6})$/i.exec(s.trim());
  if (!m) return null;
  const n = parseInt(m[1]!, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};
const mix = (a: [number, number, number], b: [number, number, number], t: number): string =>
  `rgb(${a.map((v, i) => Math.round(v + (b[i]! - v) * t)).join(',')})`;

/** Fill for a share t (0..1) of the value range. */
export const mapColor = (t: number, o: { low?: string; mid?: string; high?: string; diverging?: boolean } = {}): { fill: string; fillOpacity: number } => {
  t = Math.max(0, Math.min(1, t));
  const lo = hex(o.low), hi = hex(o.high), mid = hex(o.mid);
  if (o.diverging) {
    const L = lo ?? [201, 76, 60], M = mid ?? [245, 241, 228], H = hi ?? [12, 122, 75];
    return { fill: t < 0.5 ? mix(L, M, t * 2) : mix(M, H, (t - 0.5) * 2), fillOpacity: 1 };
  }
  if (lo && hi) return { fill: mix(lo, hi, t), fillOpacity: 1 };
  return { fill: o.high ?? 'var(--chart-1)', fillOpacity: 0.14 + 0.86 * t };
};

const pathOf = (r: Region, P: (lon: number, lat: number) => [number, number]): string => {
  let d = '';
  for (const ring of r.rings) {
    for (let i = 0; i < ring.length; i += 2) {
      let lon = ring[i]!;
      if (r.a2 === 'AK' && lon > 0) lon -= 360;
      const [x, y] = P(lon, ring[i + 1]!);
      d += `${i ? 'L' : 'M'}${x.toFixed(1)},${y.toFixed(1)}`;
    }
    d += 'Z';
  }
  return d;
};

export const drawMap = (c: DrawCtx): void => {
  const { chart, data, plot, thumb, compact, els } = c;
  c.elementTips = true;
  const geo = geoData();
  if (!geo) {
    void loadGeo().then(() => c.refresh());
    els.push(<text key="load" x={plot.x + plot.w / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: thumb ? 9 : 12 }}>{tl('Loading map…')}</text>);
    return;
  }
  const mo = chart.options?.map ?? {};
  const s = data.series[0];
  if (!s) return;
  const names = data.categories;
  const match = matchRegions(geo, names, mo.area ?? 'auto');
  const value = new Map<number, { v: number; i: number }>();
  match.index.forEach((ri, i) => { const v = s.values[i]; if (ri >= 0 && v !== null && v !== undefined) value.set(ri, { v, i }); });
  const vals = [...value.values()].map((x) => x.v);
  const lo = vals.length ? Math.min(...vals) : 0, hi = vals.length ? Math.max(...vals) : 1;
  const us = match.area === 'us';
  const proj = (r: Region | null) => (lon: number, lat: number): [number, number] => (us ? usProject(r, lon, lat) : equalEarth(lon, lat));
  // Fit to the regions with data (or the whole map when showing the world).
  const showAll = mo.area === 'world' || mo.area === 'us' || !value.size;
  const fitSet = match.regions.filter((r, i) => (showAll ? !(r.a3 === 'ATA' && !value.has(i)) : value.has(i)));
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const r of fitSet) {
    const P = proj(r);
    for (const ring of r.rings) for (let i = 0; i < ring.length; i += 2) {
      let lon = ring[i]!;
      if (r.a2 === 'AK' && lon > 0) lon -= 360;
      const [x, y] = P(lon, ring[i + 1]!);
      if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
    }
  }
  if (!Number.isFinite(x0)) return;
  const legendH = !thumb && !compact && vals.length ? 30 : 0;
  const area = { x: plot.x + 2, y: plot.y + 2, w: plot.w - 4, h: plot.h - 4 - legendH };
  const padX = (x1 - x0) * 0.03, padY = (y1 - y0) * 0.03;
  x0 -= padX; x1 += padX; y0 -= padY; y1 += padY;
  const k = Math.min(area.w / (x1 - x0 || 1), area.h / (y1 - y0 || 1));
  const ox = area.x + (area.w - (x1 - x0) * k) / 2, oy = area.y + (area.h - (y1 - y0) * k) / 2;
  const T = (r: Region) => { const P = proj(r); return (lon: number, lat: number): [number, number] => { const [x, y] = P(lon, lat); return [ox + (x - x0) * k, oy + (y - y0) * k]; }; };
  const shapes: ReactNode[] = [];
  const labels: ReactNode[] = [];
  const clipId = `${c.uid}map`;
  c.defs.push(<clipPath key={clipId} id={clipId}><rect x={area.x} y={area.y} width={area.w} height={area.h} /></clipPath>);
  match.regions.forEach((r, ri) => {
    if (r.a3 === 'ATA' && !value.has(ri) && !showAll) return;
    const hit = value.get(ri);
    const t = hit && hi > lo ? (hit.v - lo) / (hi - lo) : hit ? 1 : 0;
    const st = hit ? mapColor(t, mo) : { fill: 'var(--chart-grid)', fillOpacity: 0.9 };
    const d = pathOf(r, T(r));
    const hot = c.hot === ri;
    shapes.push(<path key={`r${ri}`} data-tip={hit ? '' : undefined} d={d} style={{ ...st, fillRule: 'evenodd', stroke: hot ? 'var(--chart-ink)' : 'var(--chart-bg)', strokeWidth: hot ? 1.4 : thumb ? 0.3 : 0.6, strokeLinejoin: 'round' }}
      onPointerEnter={thumb || !hit ? undefined : () => {
        c.setHot(ri);
        const P = T(r);
        const [cx, cy] = P((r.bbox[0] + r.bbox[2]) / 2, (r.bbox[1] + r.bbox[3]) / 2);
        c.setTip({ x: cx, y: cy, title: names[hit.i] ?? r.name, rows: [{ color: st.fill, label: s.name, value: fmtv(hit.v, data.fmt) }] });
      }} />);
    const lm = mo.labels ?? 'none';
    if (hit && lm !== 'none' && !thumb && !compact) {
      const P = T(r);
      // Label at the centre of the region's largest ring.
      const big = r.rings.reduce((m, ring) => (ring.length > m.length ? ring : m), r.rings[0]!);
      let bx0 = Infinity, by0 = Infinity, bx1 = -Infinity, by1 = -Infinity;
      for (let i = 0; i < big.length; i += 2) { const [x, y] = P(big[i]!, big[i + 1]!); bx0 = Math.min(bx0, x); bx1 = Math.max(bx1, x); by0 = Math.min(by0, y); by1 = Math.max(by1, y); }
      const label = lm === 'all' ? `${names[hit.i] ?? r.name}` : (r.a2 && us ? r.a2 : names[hit.i] ?? r.name);
      if (lm === 'all' || textW(label, 9.5) < bx1 - bx0) labels.push(<text key={`l${ri}`} x={(bx0 + bx1) / 2} y={(by0 + by1) / 2 + 3} textAnchor="middle" style={{ ...ink, fontSize: 9.5, pointerEvents: 'none', paintOrder: 'stroke', stroke: 'var(--chart-bg)', strokeWidth: 2.5, strokeOpacity: 0.7 }}>{label}</text>);
    }
  });
  els.push(<g key="map" clipPath={`url(#${clipId})`}>{shapes}{labels}</g>);
  if (legendH) {
    const lw = Math.min(220, plot.w * 0.5), lx = plot.x + plot.w - lw - 6, ly = plot.y + plot.h - legendH + 6;
    const id = `${c.uid}mapg`;
    const stops = Array.from({ length: 11 }, (_, i) => i / 10);
    c.defs.push(<linearGradient key={id} id={id} x1="0" x2="1" y1="0" y2="0">{stops.map((t) => { const st = mapColor(t, mo); return <stop key={t} offset={t} style={{ stopColor: st.fill, stopOpacity: st.fillOpacity }} />; })}</linearGradient>);
    els.push(<rect key="lg" x={lx} y={ly} width={lw} height={9} rx={2} style={{ fill: `url(#${id})`, stroke: 'var(--chart-grid)' }} />);
    els.push(<text key="lgl" x={lx} y={ly + 21} style={{ ...muted, fontSize: 10 }}>{formatNumber(lo, data.fmt, true)}</text>);
    els.push(<text key="lgh" x={lx + lw} y={ly + 21} textAnchor="end" style={{ ...muted, fontSize: 10 }}>{formatNumber(hi, data.fmt, true)}</text>);
    const missing = names.filter((_, i) => match.index[i]! < 0 && names[i]).length;
    if (missing) els.push(<text key="miss" x={plot.x + 4} y={ly + 21} style={{ ...muted, fontSize: 10 }}>{tl('{count, plural, one {# name not found} other {# names not found}}', { count: missing })}</text>);
  }
};
