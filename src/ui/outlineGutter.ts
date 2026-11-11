// The outline gutter left of the row headers and above the column headers: level buttons 1..8 and
// +/- buttons for each group. Drawn on the grid canvas; clicks come from GridView. Everything here is in logical
// pixels: on a right-to-left sheet the canvas is mirrored, so the row gutter ends up on the right.
import type { GridStore } from '../app/store';
import type { Sheet } from '../model/workbook';
import { maxLevel, outlineGroups, setGroupCollapsed, showLevel, type Axis, type OutlineGroup } from '../app/ops/outline';
import { paneFor, panesL, rectIn, type Pane, type Theme, type Viewport } from './render';

const STEP = 14;

/** Width of the row gutter and height of the column gutter, in screen pixels. */
export const outlineSize = (sheet: Sheet, zoom: number): { w: number; h: number } => {
  if (!sheet.view.showHeaders) return { w: 0, h: 0 };
  const r = maxLevel(sheet, 'row'), c = maxLevel(sheet, 'col');
  return { w: r ? Math.round((r + 1) * STEP * zoom + 6) : 0, h: c ? Math.round((c + 1) * STEP * zoom + 6) : 0 };
};

interface Box { x: number; y: number; s: number }
const center = (level: number, zoom: number) => 3 + (level - 1) * STEP * zoom + (STEP * zoom) / 2;

/** Screen boxes of the level buttons for one axis. */
const levelButtons = (vp: Viewport, axis: Axis, levels: number): Box[] => {
  const z = vp.zoom, s = Math.round(11 * z);
  const ow = vp.outlineW ?? 0, oh = vp.outlineH ?? 0;
  return Array.from({ length: levels + 1 }, (_, i) => (axis === 'row'
    ? { x: center(i + 1, z) - s / 2, y: oh + (vp.headerH - oh - s) / 2, s }
    : { x: ow + (vp.headerW - ow - s) / 2, y: center(i + 1, z) - s / 2, s }));
};

interface GroupMark { g: OutlineGroup; axis: Axis; box: Box | null; line: [number, number, number, number] | null }

/** Lines and +/- boxes of the groups in view. */
const marks = (store: GridStore, vp: Viewport, axis: Axis): GroupMark[] => {
  const s = store.sheet;
  const z = vp.zoom, size = Math.round(9 * z);
  const panes = panesL(vp);
  const out: GroupMark[] = [];
  const pos = (i: number): { a: number; len: number } | null => {
    const len = axis === 'row' ? vp.layout.rows.size(i) : vp.layout.cols.size(i);
    if (!len) return null;
    const p: Pane = axis === 'row' ? paneFor(panes, vp, i, vp.frozenCols) : paneFor(panes, vp, vp.frozenRows, i);
    const r = rectIn(p, vp, { r1: i, c1: i, r2: i, c2: i });
    return axis === 'row' ? { a: r.y, len: r.h } : { a: r.x, len: r.w };
  };
  const span = axis === 'row' ? [vp.headerH, vp.height] : [vp.headerW, vp.width];
  for (const g of outlineGroups(s, axis)) {
    const c = center(g.level, z);
    const sp = pos(g.summary);
    const box = sp && sp.a + sp.len > span[0]! && sp.a < span[1]! ? (axis === 'row' ? { x: c - size / 2, y: sp.a + (sp.len - size) / 2, s: size } : { x: sp.a + (sp.len - size) / 2, y: c - size / 2, s: size }) : null;
    let line: GroupMark['line'] = null;
    if (!g.collapsed) {
      const first = firstShown(vp, axis, g.start, g.end), last = lastShown(vp, axis, g.start, g.end);
      const a = first !== null ? pos(first) : null, b = last !== null ? pos(last) : null;
      if (a && b) {
        const from = a.a + 2, to = b.a + b.len - 2;
        line = axis === 'row' ? [c, Math.max(from, span[0]!), c, Math.min(to, span[1]!)] : [Math.max(from, span[0]!), c, Math.min(to, span[1]!), c];
      }
    }
    if (box || line) out.push({ g, axis, box, line });
  }
  return out;
};

const firstShown = (vp: Viewport, axis: Axis, a: number, b: number): number | null => { for (let i = a; i <= b && i - a < 100000; i++) if ((axis === 'row' ? vp.layout.rows.size(i) : vp.layout.cols.size(i)) > 0) return i; return null; };
const lastShown = (vp: Viewport, axis: Axis, a: number, b: number): number | null => { for (let i = b; i >= a && b - i < 100000; i--) if ((axis === 'row' ? vp.layout.rows.size(i) : vp.layout.cols.size(i)) > 0) return i; return null; };

const drawBox = (ctx: CanvasRenderingContext2D, b: Box, theme: Theme, label: string | null, plus: boolean | null, mirrored = false): void => {
  ctx.fillStyle = theme.bg;
  ctx.fillRect(b.x, b.y, b.s, b.s);
  ctx.strokeStyle = theme.headerInk;
  ctx.lineWidth = 1;
  ctx.strokeRect(Math.round(b.x) + 0.5, Math.round(b.y) + 0.5, Math.round(b.s) - 1, Math.round(b.s) - 1);
  ctx.fillStyle = theme.headerInk;
  if (label) {
    if (!mirrored) { ctx.fillText(label, b.x + b.s / 2, b.y + b.s / 2 + 0.5); return; }
    ctx.save(); ctx.scale(-1, 1); ctx.fillText(label, -(b.x + b.s / 2), b.y + b.s / 2 + 0.5); ctx.restore();
    return;
  }
  const m = Math.round(b.y + b.s / 2) + 0.5, cx = Math.round(b.x + b.s / 2) + 0.5;
  ctx.beginPath(); ctx.moveTo(b.x + 2, m); ctx.lineTo(b.x + b.s - 2, m);
  if (plus) { ctx.moveTo(cx, b.y + 2); ctx.lineTo(cx, b.y + b.s - 2); }
  ctx.stroke();
};

export const drawOutline = (ctx: CanvasRenderingContext2D, store: GridStore, vp: Viewport, theme: Theme): void => {
  const ow = vp.outlineW ?? 0, oh = vp.outlineH ?? 0;
  if (!ow && !oh) return;
  const s = store.sheet;
  ctx.save();
  ctx.font = `600 ${(9 * vp.zoom).toFixed(1)}px "Hanken Grotesk", system-ui, sans-serif`;
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillStyle = theme.header;
  if (ow) ctx.fillRect(0, 0, ow, vp.height);
  if (oh) ctx.fillRect(0, 0, vp.width, oh);
  ctx.fillStyle = theme.headerLine;
  if (ow) ctx.fillRect(ow - 1, 0, 1, vp.height);
  if (oh) ctx.fillRect(0, oh - 1, vp.width, 1);
  for (const axis of ['row', 'col'] as Axis[]) {
    const levels = maxLevel(s, axis);
    if (!levels) continue;
    levelButtons(vp, axis, levels).forEach((b, i) => drawBox(ctx, b, theme, String(i + 1), null, !!vp.rtl));
    ctx.save();
    ctx.beginPath();
    if (axis === 'row') ctx.rect(0, vp.headerH, ow, vp.height - vp.headerH); else ctx.rect(vp.headerW, 0, vp.width - vp.headerW, oh);
    ctx.clip();
    for (const m of marks(store, vp, axis)) {
      if (m.line) {
        ctx.strokeStyle = theme.headerInk; ctx.lineWidth = 1;
        const [x1, y1, x2, y2] = m.line;
        ctx.beginPath();
        if (axis === 'row') { ctx.moveTo(Math.round(x1) + 0.5, y1); ctx.lineTo(Math.round(x2) + 0.5, y2); }
        else { ctx.moveTo(x1, Math.round(y1) + 0.5); ctx.lineTo(x2, Math.round(y2) + 0.5); }
        ctx.stroke();
      }
      if (m.box) drawBox(ctx, m.box, theme, null, m.g.collapsed);
    }
    ctx.restore();
  }
  ctx.restore();
};

const inBox = (b: Box, x: number, y: number, pad = 3) => x >= b.x - pad && x <= b.x + b.s + pad && y >= b.y - pad && y <= b.y + b.s + pad;

/** Handles a click in the gutters. Returns true when the point was in a gutter. */
export const outlineClick = (store: GridStore, vp: Viewport, screenX: number, y: number): boolean => {
  const x = vp.rtl ? vp.width - screenX : screenX;
  const ow = vp.outlineW ?? 0, oh = vp.outlineH ?? 0;
  const inRowGutter = ow > 0 && x < ow;
  const inColGutter = oh > 0 && y < oh;
  const inRowLevels = ow > 0 && x < ow && y < vp.headerH;
  const inColLevels = oh > 0 && y < oh && x < vp.headerW;
  if (!inRowGutter && !inColGutter) return false;
  const s = store.sheet;
  for (const axis of ['row', 'col'] as Axis[]) {
    const levels = maxLevel(s, axis);
    if (!levels || (axis === 'row' ? !inRowLevels : !inColLevels)) continue;
    const hit = levelButtons(vp, axis, levels).findIndex((b) => inBox(b, x, y));
    if (hit >= 0) { showLevel(store, axis, hit + 1); return true; }
  }
  const axis: Axis = inRowGutter && y >= vp.headerH ? 'row' : 'col';
  if ((axis === 'row' && !inRowGutter) || (axis === 'col' && (!inColGutter || x < vp.headerW))) return true;
  const m = marks(store, vp, axis).sort((a, b) => b.g.level - a.g.level).find((k) => k.box && inBox(k.box, x, y));
  if (m) setGroupCollapsed(store, axis, m.g, !m.g.collapsed);
  return true;
};
