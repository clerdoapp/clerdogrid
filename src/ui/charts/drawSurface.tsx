// Surface charts: 3D surface, wireframe, contour and wireframe contour. Values are split into bands.
import type { ReactNode } from 'react';
import type { Chart } from '../../model/workbook';
import { formatNumber, seriesColor, type ResolvedChart } from './chartData';
import { axisOf, makeValueScale } from './axes';
import { drawWalls, faceEls, type Face, type V3 } from './draw3d';
import { fmtv, muted, textW, trunc, type DrawCtx } from './svgKit';
import { N, T } from '../i18n';

/** Band edges for a surface chart: the value axis ticks over the data range. */
export const surfaceBands = (chart: Chart, data: ResolvedChart): number[] => {
  const vals = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  if (!vals.length) return [0, 1];
  const sc = makeValueScale(Math.min(...vals), Math.max(...vals), 7, { ...axisOf(chart, 'y'), logBase: undefined, reverse: false });
  return sc.ticks.length >= 2 ? sc.ticks : [sc.min, sc.max];
};

export const bandColor = (k: number, style: number): string => seriesColor(k, undefined, style);
export const bandLabel = (lo: number, hi: number, fmt: string): string => `${formatNumber(lo, fmt, true)}–${formatNumber(hi, fmt, true)}`;

type P3 = [number, number, number]; // x, z, value

/** Clips a polygon to lo ≤ value ≤ hi (Sutherland–Hodgman on the value coordinate). */
export const clipBand = (poly: P3[], lo: number, hi: number): P3[] => {
  const clip = (pts: P3[], keep: (v: number) => boolean, edge: number): P3[] => {
    const out: P3[] = [];
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i]!, b = pts[(i + 1) % pts.length]!;
      const ia = keep(a[2]), ib = keep(b[2]);
      if (ia) out.push(a);
      if (ia !== ib) {
        const t = (edge - a[2]) / (b[2] - a[2]);
        out.push([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, edge]);
      }
    }
    return out;
  };
  let p = clip(poly, (v) => v >= lo, lo);
  if (p.length >= 3) p = clip(p, (v) => v <= hi, hi);
  return p.length >= 3 ? p : [];
};

const bandOf = (bands: number[], v: number): number => {
  for (let k = 0; k < bands.length - 1; k++) if (v <= bands[k + 1]! + 1e-12) return k;
  return bands.length - 2;
};

export const drawSurface = (c: DrawCtx): void => {
  const { chart, data, tr, plot, thumb, els, style } = c;
  c.elementTips = true;
  const rows = data.series.length;
  const cols = Math.max(0, ...data.series.map((s) => s.values.length));
  if (rows < 2 || cols < 2) {
    els.push(<text key="few" x={plot.x + plot.w / 2} y={plot.y + plot.h / 2} textAnchor="middle" style={{ ...muted, fontSize: 11 }}>{T('A surface needs at least two rows and two columns of numbers.')}</text>);
    return;
  }
  const grid = data.series.map((s) => Array.from({ length: cols }, (_, i) => s.values[i] ?? null));
  const bands = surfaceBands(chart, data);
  const fmt = axisOf(chart, 'y').numFmt || data.fmt;
  const nb = bands.length - 1;
  // Cell triangles in (x, z, value) with x and z in 0..1.
  const X = (i: number) => i / (cols - 1), Z = (r: number) => r / (rows - 1);
  const tris: P3[][] = [];
  for (let r = 0; r < rows - 1; r++) for (let i = 0; i < cols - 1; i++) {
    const v00 = grid[r]![i], v10 = grid[r]![i + 1], v01 = grid[r + 1]![i], v11 = grid[r + 1]![i + 1];
    if (v00 === null || v10 === null || v01 === null || v11 === null) continue;
    const a: P3 = [X(i), Z(r), v00], b: P3 = [X(i + 1), Z(r), v10], cc: P3 = [X(i + 1), Z(r + 1), v11], d: P3 = [X(i), Z(r + 1), v01];
    tris.push([a, b, cc], [a, cc, d]);
  }
  // Grid edges for wireframes, split where they cross band edges.
  const edges: [P3, P3][] = [];
  for (let r = 0; r < rows; r++) for (let i = 0; i < cols; i++) {
    const v = grid[r]![i];
    if (v === null) continue;
    if (i + 1 < cols && grid[r]![i + 1] !== null) edges.push([[X(i), Z(r), v], [X(i + 1), Z(r), grid[r]![i + 1]!]]);
    if (r + 1 < rows && grid[r + 1]![i] !== null) edges.push([[X(i), Z(r), v], [X(i), Z(r + 1), grid[r + 1]![i]!]]);
  }
  const splitEdge = ([a, b]: [P3, P3]): { p: [P3, P3]; band: number }[] => {
    const cuts = [0, 1];
    for (const l of bands) { const t = (l - a[2]) / (b[2] - a[2]); if (t > 1e-9 && t < 1 - 1e-9) cuts.push(t); }
    cuts.sort((x, y) => x - y);
    const at = (t: number): P3 => [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t];
    return cuts.slice(1).map((t, k) => { const p0 = at(cuts[k]!), p1 = at(t); return { p: [p0, p1] as [P3, P3], band: bandOf(bands, (p0[2] + p1[2]) / 2) }; });
  };
  const tip = (k: number) => (thumb ? undefined : () => { c.setHot(k); c.setTip({ x: plot.x + plot.w / 2, y: plot.y + 20, title: N('Range'), rows: [{ color: bandColor(k, style), label: bandLabel(bands[k]!, bands[k + 1]!, fmt), value: '' }] }); });

  if (tr.contour) {
    // Seen from above: categories across, series down.
    const labW = thumb ? 0 : Math.min(plot.w * 0.25, Math.max(...data.series.map((s) => textW(s.name))) + 10);
    const labH = thumb ? 0 : 18;
    const side = { x: plot.x + labW, y: plot.y, w: Math.max(10, plot.w - labW - 4), h: Math.max(10, plot.h - labH) };
    const S = (x: number, z: number): [number, number] => [side.x + x * side.w, side.y + side.h - z * side.h];
    const out: ReactNode[] = [];
    if (!tr.wire) for (const t of tris) for (let k = 0; k < nb; k++) {
      const p = clipBand(t, bands[k]!, bands[k + 1]!);
      if (!p.length) continue;
      const color = bandColor(k, style);
      out.push(<path key={`c${out.length}`} data-tip="" onPointerEnter={tip(k)} d={`M${p.map((q) => S(q[0], q[1]).map((v) => v.toFixed(2)).join(',')).join(' L')} Z`} style={{ fill: color, stroke: color, strokeWidth: 0.6 }} />);
    }
    else for (const e of edges) for (const s of splitEdge(e)) {
      const [a, b] = [S(s.p[0][0], s.p[0][1]), S(s.p[1][0], s.p[1][1])];
      out.push(<line key={`w${out.length}`} x1={a[0]} y1={a[1]} x2={b[0]} y2={b[1]} style={{ stroke: bandColor(s.band, style), strokeWidth: thumb ? 1 : 1.6 }} />);
    }
    els.push(<rect key="frame" x={side.x} y={side.y} width={side.w} height={side.h} style={{ fill: 'none', stroke: 'var(--chart-grid)' }} />, ...out);
    if (!thumb) {
      const skip = Math.max(1, Math.ceil(Math.max(...data.categories.map((t) => textW(t))) / (side.w / cols) / 0.9));
      data.categories.forEach((t, i) => i % skip === 0 && els.push(<text key={`cx${i}`} x={S(X(i), 0)[0]} y={side.y + side.h + 14} textAnchor="middle" style={{ ...muted, fontSize: 10.5 }}>{trunc(t, 80)}</text>));
      const rskip = Math.max(1, Math.ceil(14 / (side.h / rows)));
      data.series.forEach((s, r) => r % rskip === 0 && els.push(<text key={`cz${r}`} x={side.x - 6} y={S(0, Z(r))[1] + 4} textAnchor="end" style={{ ...muted, fontSize: 10.5 }}>{trunc(s.name, labW - 10)}</text>));
    }
    return;
  }

  const v = chart.options?.view3D ?? {};
  const all = grid.flat().filter((x): x is number => x !== null);
  const walls = drawWalls(c, { size: [1, 0.55 * ((v.heightPercent ?? 100) / 100), 0.9 * Math.max(0.2, Math.min(20, (v.depthPercent ?? 100) / 100))], lo: Math.min(...all, bands[0]!), hi: Math.max(...all, bands.at(-1)!), fmt, cats: data.categories, rows: data.series.map((s) => s.name), horizontal: false });
  const D = walls.size[2];
  const W3 = (q: P3): V3 => [q[0], walls.valueToWorld(q[2]), q[1] * D];
  if (tr.wire) {
    const segsOut: { d: number; el: ReactNode }[] = [];
    for (const e of edges) for (const s of splitEdge(e)) {
      const a = walls.cam.project(W3(s.p[0])), b = walls.cam.project(W3(s.p[1]));
      segsOut.push({ d: (a[2] + b[2]) / 2, el: <line key={`w${segsOut.length}`} x1={a[0]} y1={a[1]} x2={b[0]} y2={b[1]} style={{ stroke: bandColor(s.band, style), strokeWidth: thumb ? 1 : 1.5, strokeLinecap: 'round' }} /> });
    }
    els.push(...segsOut.sort((x, y) => y.d - x.d).map((s) => s.el), ...walls.labels);
    return;
  }
  const faces: Face[] = [];
  tris.forEach((t, ti) => {
    // Light from the upper left: shade by how much the triangle faces up.
    const [a, b, cc] = t.map(W3) as [V3, V3, V3];
    const u = [b[0] - a[0], b[1] - a[1], b[2] - a[2]], w = [cc[0] - a[0], cc[1] - a[1], cc[2] - a[2]];
    const nrm = [u[1]! * w[2]! - u[2]! * w[1]!, u[2]! * w[0]! - u[0]! * w[2]!, u[0]! * w[1]! - u[1]! * w[0]!];
    const len = Math.hypot(...nrm) || 1;
    const up = Math.abs(nrm[1]! / len);
    const shade = (1 - up) * 0.35 - 0.04;
    for (let k = 0; k < nb; k++) {
      const p = clipBand(t, bands[k]!, bands[k + 1]!);
      if (!p.length) continue;
      const pr = p.map((q) => walls.cam.project(W3(q)));
      const color = bandColor(k, style);
      faces.push({ key: `s${ti}-${k}`, pts: pr.map((q) => [q[0], q[1]]), depth: pr.reduce((acc, q) => acc + q[2], 0) / pr.length, fill: color, shade, stroke: color, tip: tip(k) });
    }
  });
  els.push(...faceEls(faces, thumb), ...walls.labels);
  void fmtv;
};
