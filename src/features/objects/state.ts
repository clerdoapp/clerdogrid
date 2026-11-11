// UI state for the objects feature (not saved, not undoable) plus screen geometry helpers.
import type { Chart } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { anchorBox, paneFor, panesOf, rectIn, screenToSheet, sheetToScreen, type Viewport } from '../../ui/render';

export type DrawTool = 'select' | 'pen' | 'marker' | 'highlighter' | 'eraser';

export const ui = {
  /** Selected shape on the active sheet. */
  shape: null as string | null,
  /** Shape whose text is being typed. */
  editingShape: null as string | null,
  tool: 'select' as DrawTool,
  penColor: '#1f2a24',
  penWidth: 2,
  hlColor: '#ffe066',
  /** Cell under the mouse with a note or comment. */
  hover: null as { sheet: number; row: number; col: number } | null,
  /** Popover opened by a tap or a command; stays until dismissed. */
  pinned: null as { sheet: number; row: number; col: number; compose?: boolean } | null,
  commentsPane: false,
  showResolved: true,
  showAllNotes: false,
  /** A pen has been used, so touches scroll instead of drawing (palm rejection). */
  penSeen: false,
};

export const setUi = (store: GridStore, patch: Partial<typeof ui>): void => { Object.assign(ui, patch); store.emit(); };

export type Box = { left: number; top: number; width: number; height: number };

/** Screen box (relative to the grid) of an anchored object. */
export const place = (vp: Viewport, a: Chart['anchor']): Box => anchorBox(vp, a);

export const cellBox = (vp: Viewport, row: number, col: number, store: GridStore): Box => {
  const m = store.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
  const r = rectIn(paneFor(panesOf(vp), vp, m.r1, m.c1), vp, m);
  return { left: r.x, top: r.y, width: r.w, height: r.h };
};

/** Absolute sheet pixel position of an anchor's top-left corner (100% zoom). */
export const anchorPx = (vp: Viewport, a: Chart['anchor']): { x: number; y: number } => ({ x: vp.layout.cols.start(a.col) + a.dx, y: vp.layout.rows.start(a.row) + a.dy });

/** Anchor for a box given in absolute sheet pixels. */
export const anchorAt = (vp: Viewport, x: number, y: number, width: number, height: number): Chart['anchor'] => {
  const cx = Math.max(0, x), cy = Math.max(0, y);
  const col = vp.layout.cols.at(cx), row = vp.layout.rows.at(cy);
  return { col, row, dx: Math.round(cx - vp.layout.cols.start(col)), dy: Math.round(cy - vp.layout.rows.start(row)), width: Math.max(4, Math.round(width)), height: Math.max(2, Math.round(height)) };
};

/**
 * Sheet pixels (100% zoom) to screen pixels relative to the grid, honouring frozen and split panes: the point is
 * placed in the pane that shows it (the main pane first), or in the main pane when none does.
 */
export const toScreen = (vp: Viewport, sx: number, sy: number): [number, number] => sheetToScreen(vp, sx, sy);

/** Screen pixels relative to the grid to sheet pixels, using the pane under the point. */
export const toSheet = (vp: Viewport, x: number, y: number): [number, number] => screenToSheet(vp, x, y);

/** Top-left of the visible cells in sheet pixels, for placing new objects. */
export const visibleOrigin = (vp: Viewport): { x: number; y: number } => {
  const lx = vp.headerW + vp.layout.cols.start(vp.frozenCols) * vp.zoom;
  const [x, y] = toSheet(vp, vp.rtl ? vp.width - lx : lx, vp.headerH + vp.layout.rows.start(vp.frozenRows) * vp.zoom);
  return { x, y };
};

/** Last viewport the overlays were drawn with, so commands can place things in view. */
export const lastVp = { current: null as Viewport | null };

export const fmtTime = (iso: string): string => {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const today = new Date();
  return d.toDateString() === today.toDateString() ? d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : d.toLocaleDateString([], { month: 'short', day: 'numeric', year: d.getFullYear() === today.getFullYear() ? undefined : 'numeric' });
};
