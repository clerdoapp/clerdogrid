// 3D column, bar, line and area charts: a camera projection with painter's ordering, in SVG.
import type { ReactNode } from 'react';
import type { View3D } from '../../model/workbook';
import { formatNumber, seriesColor } from './chartData';
import { axisOf, makeValueScale, tickLabel, type ValScale } from './axes';
import { stackSegs } from './drawCartesian';
import { fillCss, fmtv, ink, labelOpts, labelText, muted, textW, trunc, type DrawCtx } from './svgKit';
import { T } from '../i18n';

export type V3 = [number, number, number];
export interface Camera {
  /** World point → screen x, y and depth (larger is farther). */
  project: (p: V3) => [number, number, number];
  /** True when the camera looks from the right (so the left side wall is the far one). */
  fromRight: boolean;
  fromAbove: boolean;
}

/**
 * Camera for a box of size w × h × d (x right, y up, z away from the viewer), fitted into `area`.
 * rotX tilts (looking down), rotY turns around the vertical axis, perspective sets the field of view.
 */
export const makeCamera = (v: View3D, size: V3, area: { x: number; y: number; w: number; h: number }, extra: V3[] = []): Camera => {
  const rx = ((v.rotX ?? 15) * Math.PI) / 180, ry = ((v.rotY ?? 20) * Math.PI) / 180;
  const [w, h, d] = size;
  const c = [w / 2, h / 2, d / 2];
  const persp = v.rightAngle ? 0 : Math.max(0, Math.min(240, v.perspective ?? 30));
  const span = Math.hypot(w, h, d);
  const fov = (persp / 2) * (Math.PI / 180);
  const dist = fov > 0.001 ? span / 2 / Math.tan(fov / 2) + span / 2 : Infinity;
  const raw = (p: V3): [number, number, number] => {
    const x = p[0] - c[0]!, y = p[1] - c[1]!, z = p[2] - c[2]!;
    if (v.rightAngle) {
      // Oblique projection: the front face stays a rectangle.
      const kx = Math.sin(ry) * 0.75, ky = Math.sin(rx) * 0.75;
      return [x + z * kx, y + z * ky, z - y * 0.001];
    }
    const x1 = x * Math.cos(ry) - z * Math.sin(ry);
    const z1 = x * Math.sin(ry) + z * Math.cos(ry);
    const y2 = y * Math.cos(rx) + z1 * Math.sin(rx);
    const z2 = z1 * Math.cos(rx) - y * Math.sin(rx);
    const s = Number.isFinite(dist) ? dist / (dist + z2) : 1;
    return [x1 * s, y2 * s, z2];
  };
  const corners: V3[] = [];
  for (const x of [0, w]) for (const y of [0, h]) for (const z of [0, d]) corners.push([x, y, z]);
  const pts = [...corners, ...extra].map(raw);
  const minX = Math.min(...pts.map((p) => p[0])), maxX = Math.max(...pts.map((p) => p[0]));
  const minY = Math.min(...pts.map((p) => p[1])), maxY = Math.max(...pts.map((p) => p[1]));
  const k = Math.min(area.w / (maxX - minX || 1), area.h / (maxY - minY || 1));
  const ox = area.x + (area.w - (maxX - minX) * k) / 2, oy = area.y + (area.h - (maxY - minY) * k) / 2;
  return {
    project: (p) => { const r = raw(p); return [ox + (r[0] - minX) * k, oy + (maxY - r[1]) * k, r[2]]; },
    fromRight: v.rightAngle ? (v.rotY ?? 20) >= 0 : Math.sin(ry) >= 0,
    fromAbove: (v.rotX ?? 15) >= 0,
  };
};

export interface Face { pts: [number, number][]; depth: number; fill: string; shade: number; style?: Record<string, unknown>; key: string; tip?: () => void; stroke?: string; fillOpacity?: number }

/** Polygon area sign in screen space (y down): positive when the points run clockwise on screen. */
export const winding = (p: [number, number][]): number => {
  let a = 0;
  for (let i = 0; i < p.length; i++) { const [x1, y1] = p[i]!, [x2, y2] = p[(i + 1) % p.length]!; a += x1 * y2 - x2 * y1; }
  return a;
};

export const faceEls = (faces: Face[], thumb: boolean): ReactNode[] => faces.sort((a, b) => b.depth - a.depth).map((f) => {
  const d = `M${f.pts.map(([x, y]) => `${x.toFixed(2)},${y.toFixed(2)}`).join(' L')} Z`;
  return (
    <g key={f.key} data-tip={f.tip ? '' : undefined} onPointerEnter={f.tip}>
      <path d={d} style={{ fill: f.fill, fillOpacity: f.fillOpacity ?? 1, stroke: f.stroke ?? 'none', strokeWidth: 0.6, strokeLinejoin: 'round', ...(f.style ?? {}) }} />
      {f.shade !== 0 && <path d={d} style={{ fill: f.shade > 0 ? '#000' : '#fff', fillOpacity: Math.min(0.6, Math.abs(f.shade)), stroke: thumb ? 'none' : 'rgba(0,0,0,0.12)', strokeWidth: 0.5 }} />}
    </g>
  );
});

/** Shading for a face from its outward normal: tops light, sides darker. */
const SHADE = { top: -0.08, front: 0.04, side: 0.2, back: 0.12, bottom: 0.3 } as const;

/** Faces of a box (or cylinder, cone, pyramid) standing on [x0,x1]×[y0,y1]×[z0,z1]. */
export const solidFaces = (cam: Camera, key: string, b: { x0: number; x1: number; y0: number; y1: number; z0: number; z1: number }, shape: 'box' | 'cylinder' | 'cone' | 'pyramid', fill: string, horizontal: boolean, tip?: () => void, taper?: [number, number]): Face[] => {
  const out: Face[] = [];
  const P = (p: V3) => cam.project(p);
  const face = (k: string, pts: V3[], shade: number) => {
    const pr = pts.map(P);
    const scr = pr.map((p): [number, number] => [p[0], p[1]]);
    if (winding(scr) <= 0) return;
    out.push({ key: `${key}${k}`, pts: scr, depth: pr.reduce((a, p) => a + p[2], 0) / pr.length, fill, shade, tip });
  };
  const { x0, x1, y0, y1, z0, z1 } = b;
  if (shape === 'box') {
    face('f', [[x0, y0, z0], [x0, y1, z0], [x1, y1, z0], [x1, y0, z0]], SHADE.front);
    face('k', [[x1, y0, z1], [x1, y1, z1], [x0, y1, z1], [x0, y0, z1]], SHADE.back);
    face('l', [[x0, y0, z1], [x0, y1, z1], [x0, y1, z0], [x0, y0, z0]], SHADE.side);
    face('r', [[x1, y0, z0], [x1, y1, z0], [x1, y1, z1], [x1, y0, z1]], SHADE.side);
    face('t', [[x0, y1, z0], [x0, y1, z1], [x1, y1, z1], [x1, y1, z0]], horizontal ? SHADE.side : SHADE.top);
    face('b', [[x0, y0, z1], [x0, y0, z0], [x1, y0, z0], [x1, y0, z1]], SHADE.bottom);
    return out;
  }
  // Round and pointed shapes along the value direction: y for columns, x for bars.
  const N = shape === 'pyramid' ? 4 : 20;
  const cx = (x0 + x1) / 2, cz = (z0 + z1) / 2;
  const rx = (x1 - x0) / 2, rz = (z1 - z0) / 2;
  const [t0, t1] = taper ?? (shape === 'cylinder' ? [1, 1] : [1, 0]);
  const ring = (k: number, lvl: number, scale: number): V3 => {
    const a = (k / N) * Math.PI * 2 + (shape === 'pyramid' ? Math.PI / 4 : 0);
    const r = shape === 'pyramid' ? Math.SQRT2 : 1;
    if (horizontal) {
      const cy = (y0 + y1) / 2, ry = (y1 - y0) / 2;
      return [lvl, cy + Math.sin(a) * ry * r * scale, cz + Math.cos(a) * rz * r * scale];
    }
    return [cx + Math.cos(a) * rx * r * scale, lvl, cz + Math.sin(a) * rz * r * scale];
  };
  const lo = horizontal ? x0 : y0, hi = horizontal ? x1 : y1;
  for (let k = 0; k < N; k++) {
    const a0 = ring(k, lo, t0), a1 = ring(k + 1, lo, t0), b1 = ring(k + 1, hi, t1), b0 = ring(k, hi, t1);
    const mid = ((k + 0.5) / N) * Math.PI * 2;
    const light = 0.05 + 0.2 * (0.5 - 0.5 * Math.cos(mid - 0.6));
    const pts: V3[] = horizontal ? [a0, b0, b1, a1] : [a0, a1, b1, b0];
    face(`s${k}`, t1 === 0 && t0 !== 0 ? (horizontal ? [a0, b0, a1] : [a0, a1, b0]) : pts, light);
  }
  if (t1 > 0) face('t', Array.from({ length: N }, (_, k) => ring(horizontal ? k : N - k, hi, t1)), horizontal ? SHADE.side : SHADE.top);
  if (t0 > 0) face('b', Array.from({ length: N }, (_, k) => ring(horizontal ? N - k : k, lo, t0)), SHADE.bottom);
  return out;
};

const view = (c: DrawCtx): View3D => c.chart.options?.view3D ?? {};

interface Walls { size: [number, number, number]; scale: ValScale; cam: Camera; horizontal: boolean; valueToWorld: (v: number) => number; /** Tick and category labels, drawn over the shapes. */ labels: ReactNode[] }

/** Back wall, side wall and floor with gridlines and tick labels; returns the camera. */
export const drawWalls = (c: DrawCtx, o: { size: [number, number, number]; lo: number; hi: number; fmt: string; cats: string[]; rows?: string[]; horizontal: boolean; percent?: boolean }): Walls => {
  const { plot, thumb, els, chart } = c;
  const yAx = axisOf(chart, 'y');
  const nt = thumb ? 3 : Math.max(3, Math.min(7, Math.floor(plot.h / 45)));
  const scale = makeValueScale(o.lo, o.hi, nt, yAx);
  const [W, H, D] = o.size;
  const lab = thumb ? 0 : 1;
  const maxTick = Math.max(...scale.ticks.map((t) => textW(tickLabel(t, scale, o.fmt))));
  const maxCat = Math.max(0, ...o.cats.map((s) => textW(s)));
  const titleY = !thumb && !c.compact && c.data.axisTitles.y ? 16 : 0, titleX = !thumb && !c.compact && c.data.axisTitles.x ? 16 : 0;
  const area = {
    x: plot.x + lab * (o.horizontal ? Math.min(maxCat, plot.w * 0.25) + 16 : maxTick + 10) + titleY,
    y: plot.y + 4,
    w: plot.w - lab * ((o.horizontal ? Math.min(maxCat, plot.w * 0.25) + 16 : maxTick + 10) + (o.rows ? Math.min(90, Math.max(0, ...o.rows.map((r) => textW(r)))) + 8 : 6)) - titleY,
    h: plot.h - lab * (o.horizontal ? 22 : 24) - 8 - titleX,
  };
  const cam = makeCamera(view(c), o.size, area);
  const labels: ReactNode[] = [];
  const vw = (v: number) => scale.pos(v) * (o.horizontal ? W : H);
  const wallFill = chart.options?.walls?.fill?.color ?? 'var(--chart-wall, rgba(127,140,133,0.07))';
  const floorFill = chart.options?.floor?.fill?.color ?? 'var(--chart-wall, rgba(127,140,133,0.10))';
  const wallLine = { stroke: 'var(--chart-grid)', strokeWidth: 1 };
  const poly = (pts: V3[]) => pts.map((p) => cam.project(p)).map((p) => `${p[0].toFixed(2)},${p[1].toFixed(2)}`).join(' ');
  const sideX = cam.fromRight ? 0 : W;
  const floorY = cam.fromAbove ? 0 : H;
  els.push(<polygon key="wb" points={poly([[0, 0, D], [W, 0, D], [W, H, D], [0, H, D]])} style={{ fill: wallFill, ...wallLine }} />);
  els.push(<polygon key="ws" points={poly([[sideX, 0, 0], [sideX, 0, D], [sideX, H, D], [sideX, H, 0]])} style={{ fill: wallFill, ...wallLine }} />);
  els.push(<polygon key="wf" points={poly([[0, floorY, 0], [W, floorY, 0], [W, floorY, D], [0, floorY, D]])} style={{ fill: floorFill, ...wallLine }} />);
  const grid = yAx.majorGrid !== false;
  scale.ticks.forEach((t, k) => {
    const v = vw(t);
    if (o.horizontal) {
      if (grid) els.push(<polyline key={`g${k}`} points={poly([[v, 0, 0], [v, floorY, D], [v, H, D]].map((p, i) => (i === 0 ? [v, floorY, 0] : p)) as V3[])} style={{ fill: 'none', ...wallLine }} />);
      if (!thumb) { const [x, y] = cam.project([v, floorY, 0]); labels.push(<text key={`t${k}`} x={x} y={y + 14} textAnchor="middle" style={{ ...muted, fontSize: 10.5 }}>{tickLabel(t, scale, o.fmt)}</text>); }
    } else {
      if (grid) els.push(<polyline key={`g${k}`} points={poly([[sideX, v, 0], [sideX, v, D], [sideX === 0 ? W : 0, v, D]])} style={{ fill: 'none', ...wallLine }} />);
      if (!thumb) {
        const [x, y] = cam.project([sideX, v, sideX === 0 ? 0 : D]);
        const left = sideX === 0;
        labels.push(<text key={`t${k}`} x={left ? x - 6 : cam.project([0, v, 0])[0] - 6} y={(left ? y : cam.project([0, v, 0])[1]) + 4} textAnchor="end" style={{ ...muted, fontSize: 10.5 }}>{tickLabel(t, scale, o.fmt)}</text>);
      }
    }
  });
  if (!thumb) {
    const n = Math.max(1, o.cats.length);
    const skip = Math.max(1, Math.ceil((o.horizontal ? 14 : Math.max(16, maxCat * 0.9)) / ((o.horizontal ? area.h : area.w) / n)));
    o.cats.forEach((label, i) => {
      if (i % skip) return;
      const t = (i + 0.5) / n;
      if (o.horizontal) {
        const [x, y] = cam.project([0, H - t * H, 0]);
        labels.push(<text key={`c${i}`} x={x - 12} y={y + 4} textAnchor="end" style={{ ...muted, fontSize: 10.5 }}>{trunc(label, plot.w * 0.25)}</text>);
      } else {
        const [x, y] = cam.project([t * W, floorY, 0]);
        labels.push(<text key={`c${i}`} x={x} y={y + 14} textAnchor="middle" style={{ ...muted, fontSize: 10.5 }}>{trunc(label, Math.max(30, (area.w / n) * 1.2))}</text>);
      }
    });
    let lastY = Infinity;
    o.rows?.forEach((label, r) => {
      const z = ((r + 0.5) / o.rows!.length) * D;
      const [x, y] = cam.project([cam.fromRight ? W : 0, floorY, z]);
      if (Math.abs(lastY - y) < 11) return;
      lastY = y;
      labels.push(<text key={`r${r}`} x={x + (cam.fromRight ? 6 : -6)} y={y + 4} textAnchor={cam.fromRight ? 'start' : 'end'} style={{ ...muted, fontSize: 10.5 }}>{trunc(label, 90)}</text>);
    });
    if (titleY) labels.push(<text key="yt" x={plot.x + 9} y={area.y + area.h / 2} textAnchor="middle" transform={`rotate(-90 ${plot.x + 9} ${area.y + area.h / 2})`} style={{ ...ink, fontSize: 11, fontWeight: 600 }}>{trunc(c.data.axisTitles.y ?? '', area.h)}</text>);
    if (titleX) labels.push(<text key="xt" x={area.x + area.w / 2} y={plot.y + plot.h - 4} textAnchor="middle" style={{ ...ink, fontSize: 11, fontWeight: 600 }}>{trunc(c.data.axisTitles.x ?? '', area.w)}</text>);
  }
  return { size: o.size, scale, cam, horizontal: o.horizontal, valueToWorld: vw, labels };
};

export const draw3dBars = (c: DrawCtx): void => {
  const { chart, data, tr, thumb, els } = c;
  c.elementTips = true;
  const o = chart.options ?? {};
  const v = view(c);
  const n = Math.max(1, data.categories.length);
  const rows = tr.rows3d ? data.series.length : 1;
  const mode = tr.stack;
  const segs = stackSegs(data.series.map((s) => s.values), n, mode);
  const vals = segs.flatMap((ss) => ss.flatMap((g) => (mode === 'none' ? [g.v1] : [g.v0, g.v1])));
  const lo = Math.min(0, ...vals), hi = Math.max(0, ...vals, mode === 'percent' ? 1 : 0);
  const horizontal = tr.horizontal;
  const depthPct = Math.max(20, Math.min(2000, v.depthPercent ?? 100)) / 100;
  const W = 1, H = 0.62 * Math.max(0.05, (v.heightPercent ?? 100) / 100);
  const D = Math.min(2.5, (tr.rows3d ? 0.16 * rows : 0.2) * depthPct);
  const fmt = mode === 'percent' ? '0%' : axisOf(chart, 'y').numFmt || data.fmt;
  const walls = drawWalls(c, { size: horizontal ? [W, H, D] : [W, H, D], lo, hi, fmt, cats: data.categories, rows: tr.rows3d ? data.series.map((s) => s.name) : undefined, horizontal, percent: mode === 'percent' });
  const { cam, valueToWorld } = walls;
  const shape = o.shape ?? tr.shape;
  const gap = Math.max(0, Math.min(500, o.gapWidth ?? 150)) / 100;
  const gapDepth = Math.max(0, Math.min(500, v.gapDepth ?? 150)) / 100;
  const clustered = mode === 'none' && !tr.rows3d;
  const per = clustered ? data.series.length : 1;
  const band = (horizontal ? H : W) / n;
  const bw = band / (per + gap);
  const rowD = D / rows;
  const zw = rowD / (1 + gapDepth);
  const faces = [];
  const base = valueToWorld(Math.min(Math.max(0, walls.scale.min), walls.scale.max));
  for (let si = 0; si < data.series.length; si++) {
    const s = data.series[si]!;
    const r = tr.rows3d ? si : 0;
    const z0 = r * rowD + (rowD - zw) / 2, z1 = z0 + zw;
    for (const g of segs[si]!) {
      if (g.raw === null && mode === 'none') continue;
      const slot = clustered ? si : 0;
      const c0 = g.i * band + (band - bw * per) / 2 + slot * bw;
      const a = mode === 'none' ? base : valueToWorld(g.v0), b = valueToWorld(g.v1);
      const color = (o.varyColors && data.series.length === 1) ? seriesColor(g.i, undefined, c.style) : s.color;
      const fill = fillCss(c, s.src.points?.find((p) => p.idx === g.i)?.fill ?? s.src.fill, color).fill as string;
      const tip = thumb ? undefined : () => { c.setHot(si * 10000 + g.i); c.setTip({ x: cam.project([c0, b, z0])[0], y: cam.project([c0, b, z0])[1], title: data.categories[g.i] ?? '', rows: [{ color, label: s.name, value: g.raw === null ? T('Blank') : mode === 'percent' ? `${fmtv(g.raw, data.fmt)} (${formatNumber(g.v1 - g.v0, '0.0%')})` : fmtv(g.raw, data.fmt) }] }); };
      // Pointed shapes taper over the whole stack in stacked charts.
      let taper: [number, number] | undefined;
      if ((shape === 'cone' || shape === 'pyramid') && mode !== 'none') {
        const top = valueToWorld(Math.max(...segs.map((ss) => ss[g.i]?.v1 ?? 0)));
        const t = (y: number) => Math.max(0, 1 - (y - base) / (top - base || 1));
        taper = [t(Math.min(a, b)), t(Math.max(a, b))];
      }
      const box = horizontal
        ? { x0: Math.min(a, b), x1: Math.max(a, b), y0: H - (c0 + bw * 0.92), y1: H - c0, z0, z1 }
        : { x0: c0 + bw * 0.04, x1: c0 + bw * 0.96, y0: Math.min(a, b), y1: Math.max(a, b), z0, z1 };
      if ((horizontal ? box.x1 - box.x0 : box.y1 - box.y0) < 1e-6) continue;
      const negative = g.v1 < g.v0;
      faces.push(...solidFaces(cam, `b${si}-${g.i}`, box, shape, fill, horizontal, tip, negative && taper ? [taper[1], taper[0]] : taper));
    }
  }
  els.push(...faceEls(faces, thumb), ...walls.labels);
  // Labels on top of each bar.
  if (thumb || c.compact) return;
  data.series.forEach((s, si) => {
    const lo2 = labelOpts(chart, s);
    if (!lo2) return;
    const r = tr.rows3d ? si : 0;
    const zc = r * rowD + rowD / 2;
    segs[si]!.forEach((g) => {
      if (g.raw === null) return;
      const slot = clustered ? si : 0;
      const cc = g.i * band + (band - bw * per) / 2 + slot * bw + bw / 2;
      const end = valueToWorld(g.v1);
      const [x, y] = cam.project(horizontal ? [end, H - cc, zc - zw / 2] : [cc, end, zc - zw / 2]);
      const t = labelText(lo2, { value: g.raw, series: s.name, category: data.categories[g.i] ?? '', percent: mode === 'percent' ? g.v1 - g.v0 : undefined, cell: s.cellLabels?.[g.i] }, lo2.numFmt || data.fmt, mode === 'percent');
      els.push(<text key={`l${si}-${g.i}`} x={horizontal ? x + 5 : x} y={horizontal ? y + 4 : y - 5} textAnchor={horizontal ? 'start' : 'middle'} style={{ ...ink, fontSize: 10, pointerEvents: 'none' }}>{t}</text>);
    });
  });
};

export const draw3dLineArea = (c: DrawCtx): void => {
  const { chart, data, tr, thumb, els } = c;
  c.elementTips = true;
  const v = view(c);
  const n = Math.max(1, data.categories.length);
  const area = tr.family === 'area3d';
  const mode = area ? tr.stack : 'none';
  const segs = stackSegs(data.series.map((s) => s.values), n, mode);
  const vals = segs.flatMap((ss) => ss.flatMap((g) => (mode === 'none' ? (g.raw === null ? [] : [g.v1]) : [g.v0, g.v1])));
  const lo = Math.min(area ? 0 : Infinity, ...vals), hi = Math.max(0, ...vals);
  const rows = tr.rows3d ? data.series.length : 1;
  const depthPct = Math.max(20, Math.min(2000, v.depthPercent ?? 100)) / 100;
  const W = 1, H = 0.62 * Math.max(0.05, (v.heightPercent ?? 100) / 100), D = Math.min(2.5, 0.16 * rows * depthPct);
  const fmt = mode === 'percent' ? '0%' : axisOf(chart, 'y').numFmt || data.fmt;
  const walls = drawWalls(c, { size: [W, H, D], lo: Number.isFinite(lo) ? lo : 0, hi, fmt, cats: data.categories, rows: tr.rows3d ? data.series.map((s) => s.name) : undefined, horizontal: false });
  const { cam, valueToWorld } = walls;
  const rowD = D / rows;
  const gapDepth = Math.max(0, Math.min(500, v.gapDepth ?? 150)) / 100;
  const zw = area ? rowD / (1 + gapDepth * 0.3) : rowD / (1 + gapDepth);
  const X = (i: number) => ((i + 0.5) / n) * W;
  const base = valueToWorld(Math.min(Math.max(0, walls.scale.min), walls.scale.max));
  const faces: Parameters<typeof faceEls>[0] = [];
  const P = (p: V3): [number, number, number] => cam.project(p);
  const add = (key: string, pts: V3[], fill: string, shade: number, tip?: () => void, twoSided = true) => {
    const pr = pts.map(P);
    const scr = pr.map((p): [number, number] => [p[0], p[1]]);
    if (!twoSided && winding(scr) <= 0) return;
    faces.push({ key, pts: scr, depth: pr.reduce((a, p) => a + p[2], 0) / pr.length, fill, shade, tip });
  };
  data.series.forEach((s, si) => {
    const r = tr.rows3d ? si : 0;
    const z0 = r * rowD + (rowD - zw) / 2, z1 = z0 + zw;
    const color = fillCss(c, s.src.fill, s.color).fill as string;
    const tip = thumb ? undefined : () => { c.setHot(si); c.setTip({ x: P([X(0), valueToWorld(segs[si]![0]?.v1 ?? 0), z0])[0], y: P([X(0), valueToWorld(segs[si]![0]?.v1 ?? 0), z0])[1], title: s.name, rows: s.values.slice(0, 12).map((val, i) => ({ color, label: data.categories[i] ?? '', value: val === null ? T('Blank') : fmtv(val, data.fmt) })) }); };
    const pts = segs[si]!;
    if (area) {
      const top = pts.map((g) => valueToWorld(g.v1)), bot = pts.map((g) => (mode === 'none' ? base : valueToWorld(g.v0)));
      // Top strips between neighbouring points.
      for (let i = 0; i < pts.length - 1; i++) {
        add(`at${si}-${i}`, [[X(i), top[i]!, z0], [X(i), top[i]!, z1], [X(i + 1), top[i + 1]!, z1], [X(i + 1), top[i + 1]!, z0]], color, -0.06, tip);
        if (mode !== 'none') add(`ab${si}-${i}`, [[X(i), bot[i]!, z0], [X(i + 1), bot[i + 1]!, z0], [X(i + 1), bot[i + 1]!, z1], [X(i), bot[i]!, z1]], color, 0.3, tip, false);
      }
      const front: V3[] = [...top.map((y, i): V3 => [X(i), y, z0]), ...[...bot].reverse().map((y, k): V3 => [X(pts.length - 1 - k), y, z0])];
      // The front face is drawn last among this series' faces so strips behind it stay hidden.
      const fpr = front.map(P);
      faces.push({ key: `af${si}`, pts: fpr.map((p) => [p[0], p[1]]), depth: Math.min(...fpr.map((p) => p[2])) - 1e-6, fill: color, shade: 0.04, tip });
      const last = pts.length - 1;
      add(`ae${si}`, [[X(last), bot[last]!, z0], [X(last), top[last]!, z0], [X(last), top[last]!, z1], [X(last), bot[last]!, z1]], color, 0.2, tip);
      add(`as${si}`, [[X(0), bot[0]!, z1], [X(0), top[0]!, z1], [X(0), top[0]!, z0], [X(0), bot[0]!, z0]], color, 0.2, tip);
    } else {
      const th = H * 0.025;
      for (let i = 0; i < pts.length - 1; i++) {
        const a = pts[i]!, b = pts[i + 1]!;
        if (a.raw === null || b.raw === null) continue;
        const ya = valueToWorld(a.v1), yb = valueToWorld(b.v1);
        add(`lt${si}-${i}`, [[X(i), ya, z0], [X(i), ya, z1], [X(i + 1), yb, z1], [X(i + 1), yb, z0]], color, -0.05, tip);
        add(`lf${si}-${i}`, [[X(i), ya, z0], [X(i + 1), yb, z0], [X(i + 1), yb - th, z0], [X(i), ya - th, z0]], color, 0.18, tip);
        add(`lb${si}-${i}`, [[X(i), ya - th, z0], [X(i + 1), yb - th, z0], [X(i + 1), yb - th, z1], [X(i), ya - th, z1]], color, 0.32, tip, false);
      }
    }
  });
  els.push(...faceEls(faces, thumb), ...walls.labels);
  if (thumb || c.compact) return;
  data.series.forEach((s, si) => {
    const lo2 = labelOpts(chart, s);
    if (!lo2) return;
    const r = tr.rows3d ? si : 0;
    const z0 = r * rowD + (rowD - zw) / 2;
    segs[si]!.forEach((g) => {
      if (g.raw === null) return;
      const [x, y] = P([X(g.i), valueToWorld(g.v1), z0]);
      els.push(<text key={`l${si}-${g.i}`} x={x} y={y - 6} textAnchor="middle" style={{ ...ink, fontSize: 10, pointerEvents: 'none' }}>{labelText(lo2, { value: g.raw, series: s.name, category: data.categories[g.i] ?? '', percent: mode === 'percent' ? g.v1 - g.v0 : undefined }, lo2.numFmt || data.fmt, mode === 'percent')}</text>);
    });
  });
};
