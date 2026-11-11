// Drawing surface for printed pages. Coordinates are points from the page's top-left corner.
// Two implementations: PDF (pdfPainter.ts) and SVG for preview and printing (svgPainter.ts).
import type { FontBook } from '../../formats/pdf-text';
import type { PathCmd, StrokeStyle } from '../../formats/pdf';
import { parseXml, type XNode } from '../../formats/xml';
import { chartLightVars } from '../../ui/charts/palette';

export type { PathCmd, StrokeStyle };

export interface TextStyle { family?: string; bold?: boolean; italic?: boolean; size: number; color: string; underline?: 'single' | 'double'; strike?: boolean }

export interface Painter {
  readonly fonts: FontBook;
  /** Colors become shades of gray. */
  mono: boolean;
  save(): void;
  restore(): void;
  clip(x: number, y: number, w: number, h: number): void;
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void;
  rect(x: number, y: number, w: number, h: number, fill: string, alpha?: number): void;
  gradient(x: number, y: number, w: number, h: number, angle: number, stops: { pos: number; color: string }[]): void;
  path(cmds: PathCmd[], fill?: string, stroke?: StrokeStyle, opts?: { evenOdd?: boolean; fillAlpha?: number; strokeAlpha?: number }): void;
  /** Draws one line of text with its left edge at x and baseline at y; returns its width. */
  text(text: string, x: number, y: number, style: TextStyle, rtl?: boolean): number;
  image(data: Uint8Array, mime: string, x: number, y: number, w: number, h: number): void;
  link(x: number, y: number, w: number, h: number, uri: string): void;
}

export const grayOf = (c: string): string => {
  const m = /^#([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(c);
  if (!m) return c;
  const n = parseInt(m[1]!, 16);
  const l = Math.round(0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255));
  const h = l.toString(16).padStart(2, '0');
  return `#${h}${h}${h}${m[2] ?? ''}`;
};

const NAMED: Record<string, string> = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', grey: '#808080',
  orange: '#ffa500', purple: '#800080', cyan: '#00ffff', magenta: '#ff00ff', silver: '#c0c0c0', navy: '#000080', teal: '#008080', maroon: '#800000', transparent: '#00000000',
};

/** Light-theme values of the chart color variables (the same palette the grid uses). */
const chartVars: Record<string, string> = { '--sheet': '#ffffff', ...chartLightVars(), '--chart-hover': 'transparent' };

/** CSS color text → "#rrggbb" or "#rrggbbaa"; undefined for none. */
export const cssColor = (v: string | undefined, vars: Record<string, string> = chartVars, depth = 0): string | undefined => {
  if (!v) return undefined;
  let s = v.trim();
  const vm = /^var\(\s*(--[\w-]+)\s*(?:,\s*([^)]+))?\)$/.exec(s);
  if (vm) return depth > 4 ? undefined : cssColor(vars[vm[1]!] ?? vm[2], vars, depth + 1);
  s = s.toLowerCase();
  if (s === 'none' || s === '') return undefined;
  if (s === 'currentcolor') return '#000000';
  if (NAMED[s]) return NAMED[s];
  if (/^#[0-9a-f]{3}$/.test(s)) return '#' + s.slice(1).split('').map((x) => x + x).join('');
  if (/^#[0-9a-f]{6}([0-9a-f]{2})?$/.test(s)) return s;
  const rgb = /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(s);
  if (rgb) {
    const h = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
    const a = rgb[4] !== undefined ? +rgb[4] : 1;
    return `#${h(+rgb[1]!)}${h(+rgb[2]!)}${h(+rgb[3]!)}${a < 1 ? h(a * 255) : ''}`;
  }
  return '#000000';
};

export const chartColorVars = (): Record<string, string> => chartVars;

// ---------- SVG path data ----------

const arcToBeziers = (x1: number, y1: number, rx: number, ry: number, phi: number, large: boolean, sweep: boolean, x2: number, y2: number): PathCmd[] => {
  if (!rx || !ry) return [['L', x2, y2]];
  rx = Math.abs(rx); ry = Math.abs(ry);
  const sin = Math.sin((phi * Math.PI) / 180), cos = Math.cos((phi * Math.PI) / 180);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
  const lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let co = Math.sqrt(Math.max(0, num / den));
  if (large === sweep) co = -co;
  const cxp = (co * rx * y1p) / ry, cyp = (-co * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  if (sweep && dt < 0) dt += 2 * Math.PI;
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
  const out: PathCmd[] = [];
  const step = dt / n;
  const k = (4 / 3) * Math.tan(step / 4);
  const pt = (t: number): [number, number] => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
  const dpt = (t: number): [number, number] => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
  for (let i = 0; i < n; i++) {
    const a = t1 + i * step, b = a + step;
    const [ax, ay] = pt(a), [bx, by] = pt(b), [dax, day] = dpt(a), [dbx, dby] = dpt(b);
    out.push(['C', ax + k * dax, ay + k * day, bx - k * dbx, by - k * dby, bx, by]);
  }
  return out;
};

/** SVG path data → absolute moves, lines and cubic curves. */
export const parsePathData = (d: string): PathCmd[] => {
  const toks = d.match(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g) ?? [];
  const out: PathCmd[] = [];
  let i = 0, cmd = '', x = 0, y = 0, sx = 0, sy = 0, lcx = 0, lcy = 0, lqx = 0, lqy = 0, prev = '';
  const n = () => +toks[i++]!;
  while (i < toks.length) {
    if (/^[a-zA-Z]$/.test(toks[i]!)) cmd = toks[i++]!;
    else if (!cmd) { i++; continue; }
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const ox = rel ? x : 0, oy = rel ? y : 0;
    switch (C) {
      case 'M': x = ox + n(); y = oy + n(); sx = x; sy = y; out.push(['M', x, y]); cmd = rel ? 'l' : 'L'; break;
      case 'L': x = ox + n(); y = oy + n(); out.push(['L', x, y]); break;
      case 'H': x = ox + n(); out.push(['L', x, y]); break;
      case 'V': y = oy + n(); out.push(['L', x, y]); break;
      case 'C': { const a = ox + n(), b = oy + n(), c = ox + n(), e = oy + n(); x = ox + n(); y = oy + n(); out.push(['C', a, b, c, e, x, y]); lcx = c; lcy = e; break; }
      case 'S': {
        const a = /[CS]/.test(prev) ? 2 * x - lcx : x, b = /[CS]/.test(prev) ? 2 * y - lcy : y;
        const c = ox + n(), e = oy + n(); x = ox + n(); y = oy + n(); out.push(['C', a, b, c, e, x, y]); lcx = c; lcy = e; break;
      }
      case 'Q': case 'T': {
        let qx: number, qy: number;
        if (C === 'Q') { qx = ox + n(); qy = oy + n(); } else { qx = /[QT]/.test(prev) ? 2 * x - lqx : x; qy = /[QT]/.test(prev) ? 2 * y - lqy : y; }
        const ex = ox + n(), ey = oy + n();
        out.push(['C', x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), ex + (2 / 3) * (qx - ex), ey + (2 / 3) * (qy - ey), ex, ey]);
        lqx = qx; lqy = qy; x = ex; y = ey; break;
      }
      case 'A': {
        const rx = n(), ry = n(), rot = n(), la = n(), sw = n(); const ex = ox + n(), ey = oy + n();
        out.push(...arcToBeziers(x, y, rx, ry, rot, !!la, !!sw, ex, ey)); x = ex; y = ey; break;
      }
      case 'Z': out.push(['Z']); x = sx; y = sy; cmd = ''; break;
      default: i++;
    }
    prev = C;
  }
  return out;
};

const ellipsePath = (cx: number, cy: number, rx: number, ry: number): PathCmd[] => {
  const k = 0.5522847498;
  return [['M', cx + rx, cy], ['C', cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry], ['C', cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy],
    ['C', cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry], ['C', cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy], ['Z']];
};
export const roundRectPath = (x: number, y: number, w: number, h: number, r: number): PathCmd[] => {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  if (!r) return [['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']];
  const k = 0.5522847498 * r;
  return [['M', x + r, y], ['L', x + w - r, y], ['C', x + w - r + k, y, x + w, y + r - k, x + w, y + r], ['L', x + w, y + h - r], ['C', x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h],
    ['L', x + r, y + h], ['C', x + r - k, y + h, x, y + h - r + k, x, y + h - r], ['L', x, y + r], ['C', x, y + r - k, x + r - k, y, x + r, y], ['Z']];
};

// ---------- SVG markup → painter calls (charts, shapes) ----------

interface SvgStyle { fill?: string; stroke?: string; strokeWidth: number; fillOpacity: number; strokeOpacity: number; opacity: number; fontSize: number; fontWeight: number; anchor: string; dash?: number[]; family?: string; cap?: string; join?: string }

const parseStyle = (n: XNode, parent: SvgStyle, vars: Record<string, string>): SvgStyle => {
  const props: Record<string, string> = {};
  for (const k of ['fill', 'stroke', 'stroke-width', 'fill-opacity', 'stroke-opacity', 'opacity', 'font-size', 'font-weight', 'text-anchor', 'stroke-dasharray', 'font-family', 'stroke-linecap', 'stroke-linejoin']) if (n.attrs[k] !== undefined) props[k] = n.attrs[k]!;
  for (const decl of (n.attrs.style ?? '').split(';')) { const i = decl.indexOf(':'); if (i > 0) props[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).trim(); }
  const s: SvgStyle = { ...parent, opacity: 1 };
  if (props.fill !== undefined) s.fill = props.fill;
  if (props.stroke !== undefined) s.stroke = props.stroke;
  if (props['stroke-width']) s.strokeWidth = parseFloat(props['stroke-width']) || 0;
  if (props['fill-opacity']) s.fillOpacity = parseFloat(props['fill-opacity']);
  if (props['stroke-opacity']) s.strokeOpacity = parseFloat(props['stroke-opacity']);
  if (props.opacity) s.opacity = parseFloat(props.opacity);
  if (props['font-size']) s.fontSize = parseFloat(props['font-size']) || s.fontSize;
  if (props['font-weight']) s.fontWeight = props['font-weight'] === 'bold' ? 700 : parseInt(props['font-weight'], 10) || 400;
  if (props['text-anchor']) s.anchor = props['text-anchor'];
  if (props['font-family']) s.family = props['font-family'].split(',')[0]!.replace(/["']/g, '').trim();
  if (props['stroke-dasharray']) s.dash = props['stroke-dasharray'] === 'none' ? undefined : props['stroke-dasharray'].split(/[\s,]+/).map(Number).filter((x) => x >= 0);
  if (props['stroke-linecap']) s.cap = props['stroke-linecap'];
  if (props['stroke-linejoin']) s.join = props['stroke-linejoin'];
  void vars;
  return s;
};

const applyTransform = (p: Painter, t: string | undefined): void => {
  if (!t) return;
  for (const m of t.matchAll(/(\w+)\s*\(([^)]*)\)/g)) {
    const a = m[2]!.split(/[\s,]+/).filter(Boolean).map(Number);
    switch (m[1]) {
      case 'translate': p.transform(1, 0, 0, 1, a[0] ?? 0, a[1] ?? 0); break;
      case 'scale': p.transform(a[0] ?? 1, 0, 0, a[1] ?? a[0] ?? 1, 0, 0); break;
      case 'rotate': {
        const r = ((a[0] ?? 0) * Math.PI) / 180, cx = a[1] ?? 0, cy = a[2] ?? 0;
        p.transform(1, 0, 0, 1, cx, cy); p.transform(Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0); p.transform(1, 0, 0, 1, -cx, -cy);
        break;
      }
      case 'matrix': if (a.length === 6) p.transform(a[0]!, a[1]!, a[2]!, a[3]!, a[4]!, a[5]!); break;
      case 'skewX': p.transform(1, 0, Math.tan(((a[0] ?? 0) * Math.PI) / 180), 1, 0, 0); break;
      case 'skewY': p.transform(1, Math.tan(((a[0] ?? 0) * Math.PI) / 180), 0, 1, 0, 0); break;
    }
  }
};

const numAttr = (n: XNode, k: string, d = 0): number => { const v = parseFloat(n.attrs[k] ?? ''); return Number.isFinite(v) ? v : d; };

/** Draws SVG markup (as made by the chart and shape renderers) into a box. */
export const drawSvg = (p: Painter, markup: string, x: number, y: number, w: number, h: number, vars: Record<string, string> = chartVars): void => {
  let doc: XNode;
  try { doc = parseXml(markup); } catch { return; }
  const find = (n: XNode): XNode | undefined => (n.name === 'svg' ? n : n.children.map(find).find(Boolean));
  const svg = find(doc);
  if (!svg) return;
  const vb = (svg.attrs.viewBox ?? '').split(/[\s,]+/).map(Number);
  const vw = vb.length === 4 && vb[2]! > 0 ? vb[2]! : numAttr(svg, 'width', w), vh = vb.length === 4 && vb[3]! > 0 ? vb[3]! : numAttr(svg, 'height', h);
  const clips = new Map<string, XNode>();
  const collect = (n: XNode) => { if (n.name === 'clipPath' && n.attrs.id) clips.set(n.attrs.id, n); n.children.forEach(collect); };
  collect(svg);
  p.save();
  p.clip(x, y, w, h);
  p.transform(w / vw, 0, 0, h / vh, x, y);
  if (vb.length === 4) p.transform(1, 0, 0, 1, -vb[0]!, -vb[1]!);
  const base: SvgStyle = { fill: '#000000', strokeWidth: 1, fillOpacity: 1, strokeOpacity: 1, opacity: 1, fontSize: 12, fontWeight: 400, anchor: 'start', family: 'Hanken Grotesk' };
  const paint = (cmds: PathCmd[], s: SvgStyle, alpha: number) => {
    const fill = cssColor(s.fill, vars), stroke = cssColor(s.stroke, vars);
    const fa = s.fillOpacity * alpha * (fill && fill.length === 9 ? parseInt(fill.slice(7), 16) / 255 : 1);
    const sa = s.strokeOpacity * alpha * (stroke && stroke.length === 9 ? parseInt(stroke.slice(7), 16) / 255 : 1);
    p.path(cmds, fill?.slice(0, 7), stroke && s.strokeWidth > 0 ? { color: stroke.slice(0, 7), width: s.strokeWidth, dash: s.dash, cap: s.cap as StrokeStyle['cap'], join: s.join as StrokeStyle['join'] } : undefined, { fillAlpha: fa, strokeAlpha: sa });
  };
  const walk = (n: XNode, parent: SvgStyle, alpha: number) => {
    if (n.name === 'defs' || n.name === 'clipPath' || n.name === 'title' || n.name === 'desc' || n.name === 'marker') return;
    const s = parseStyle(n, parent, vars);
    const a = alpha * s.opacity;
    const tf = n.attrs.transform;
    const clipRef = /url\(#([^)]+)\)/.exec(n.attrs['clip-path'] ?? n.attrs.clipPath ?? '')?.[1];
    const scoped = !!tf || !!clipRef;
    if (scoped) { p.save(); applyTransform(p, tf); }
    if (clipRef) {
      const r = clips.get(clipRef)?.children.find((c) => c.name === 'rect');
      if (r) p.clip(numAttr(r, 'x'), numAttr(r, 'y'), numAttr(r, 'width'), numAttr(r, 'height'));
    }
    switch (n.name) {
      case 'svg': case 'g': case 'a': for (const c of n.children) walk(c, s, a); break;
      case 'rect': {
        const rw = numAttr(n, 'width'), rh = numAttr(n, 'height');
        if (rw > 0 && rh > 0) paint(roundRectPath(numAttr(n, 'x'), numAttr(n, 'y'), rw, rh, numAttr(n, 'rx', numAttr(n, 'ry'))), s, a);
        break;
      }
      case 'circle': { const r = numAttr(n, 'r'); if (r > 0) paint(ellipsePath(numAttr(n, 'cx'), numAttr(n, 'cy'), r, r), s, a); break; }
      case 'ellipse': paint(ellipsePath(numAttr(n, 'cx'), numAttr(n, 'cy'), numAttr(n, 'rx'), numAttr(n, 'ry')), s, a); break;
      case 'line': paint([['M', numAttr(n, 'x1'), numAttr(n, 'y1')], ['L', numAttr(n, 'x2'), numAttr(n, 'y2')]], { ...s, fill: 'none' }, a); break;
      case 'polyline': case 'polygon': {
        const pts = (n.attrs.points ?? '').trim().split(/[\s,]+/).map(Number);
        const cmds: PathCmd[] = [];
        for (let i = 0; i + 1 < pts.length; i += 2) cmds.push([i ? 'L' : 'M', pts[i]!, pts[i + 1]!]);
        if (n.name === 'polygon') cmds.push(['Z']);
        paint(cmds, s, a);
        break;
      }
      case 'path': paint(parsePathData(n.attrs.d ?? ''), s, a); break;
      case 'text': {
        const txt = (n.text + n.children.map((c) => c.text).join('')).replace(/\s+/g, ' ');
        if (!txt.trim()) break;
        const fill = cssColor(s.fill, vars) ?? '#000000';
        const st: TextStyle = { family: s.family, bold: s.fontWeight >= 600, size: s.fontSize, color: fill.slice(0, 7) };
        const tw = p.fonts.measure(txt, st, s.fontSize);
        const tx = numAttr(n, 'x') - (s.anchor === 'middle' ? tw / 2 : s.anchor === 'end' ? tw : 0);
        const dy = n.attrs['dominant-baseline'] === 'middle' || n.attrs['dominant-baseline'] === 'central' ? s.fontSize * 0.35 : 0;
        p.text(txt, tx, numAttr(n, 'y') + dy + numAttr(n, 'dy'), st);
        break;
      }
    }
    if (scoped) p.restore();
  };
  for (const c of svg.children) walk(c, base, 1);
  p.restore();
};
