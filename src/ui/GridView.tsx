import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { cellKey, colName, parseRange, type CellRange } from '../model/workbook';
import { layoutOf, pxToChars, pxToPt } from '../app/layout';
import { applyPainter, painter, run } from '../app/commands';
import { autoFill } from '../app/ops/fill';
import { clearClipboardMarquee } from '../app/ops/clipboard';
import { inRange, rangeOf, type GridStore, type Pos } from '../app/store';
import { panesOf, readTheme, SheetRenderer, cssFont, rectIn, paneFor, type Viewport, type Overlay } from './render';
import { handleGridKey } from './keys';
import { FormulaAssist } from './FormulaAssist';
import { ChartLayer } from './ChartLayer';
import { listOptions } from '../app/validation';
import { filterPopupFor } from './FilterPopup';

const REF_COLORS = ['#2f6fdf', '#d6336c', '#0c8a50', '#9c36b5', '#e8590c', '#0c8599', '#b08800'];
const COL_HEADER = 22;

/** Scrollable area is capped (browsers limit element sizes), so large sheets map scroll position proportionally. */
const MAX_SCROLL = 8_000_000;

export interface GridApi { scrollTo: (row: number, col: number) => void }

export const refsInFormula = (text: string): string[] => {
  if (!text.startsWith('=')) return [];
  const out: string[] = [];
  const re = /(?:'[^']+'|[A-Za-z_][\w.]*)?!?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g;
  const stripped = text.replace(/"[^"]*"/g, (m) => ' '.repeat(m.length));
  let m: RegExpExecArray | null;
  while ((m = re.exec(stripped))) if (!/[A-Za-z_(]/.test(stripped[m.index + m[0].length] ?? '')) out.push(m[0]);
  return out;
};

export const GridView = ({ store, apiRef }: { store: GridStore; apiRef?: React.MutableRefObject<GridApi | null> }) => {
  const wrap = useRef<HTMLDivElement>(null);
  const scroller = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const editor = useRef<HTMLTextAreaElement>(null);
  const renderer = useRef<SheetRenderer | null>(null);
  const [size, setSize] = useState({ w: 800, h: 600 });
  const [scroll, setScroll] = useState({ x: 0, y: 0 });
  const [dragState, setDragState] = useState<{ fill?: CellRange | null; resize?: { axis: 'row' | 'col'; index: number; start: number; orig: number; px: number } | null; move?: CellRange | null }>({});
  const [marqueePhase, setPhase] = useState(0);
  const [list, setList] = useState<{ row: number; col: number; options: string[] } | null>(null);
  const [filterCol, setFilterCol] = useState<{ col: number; x: number; y: number } | null>(null);
  const pointer = useRef<{ mode: 'select' | 'fill' | 'resize' | 'pickRef' | 'move' | 'header'; start: Pos; headerAxis?: 'row' | 'col'; refStart?: number; refLen?: number } | null>(null);

  const sheet = store.sheet;
  const zoom = (sheet.view.zoom / 100) * (matchMedia('(pointer: coarse)').matches ? 1.15 : 1);
  const layout = layoutOf(sheet, store.layoutVersion, store.wb.styles);
  const digits = String(Math.min(MAX_ROWS, layout.rows.at(scroll.y + size.h / zoom) + 1)).length;
  const headerW = sheet.view.showHeaders ? Math.round(Math.max(3, digits) * 8 * zoom + 14) : 0;
  const headerH = sheet.view.showHeaders ? Math.round(COL_HEADER * zoom) : 0;
  const totalW = layout.cols.start(MAX_COLS), totalH = layout.rows.start(MAX_ROWS);
  // How far the content can scroll in sheet pixels, and the spacer size used to provide it.
  const usedB = store.engine.usedBounds(store.sheetIndex);
  const extentRows = Math.min(MAX_ROWS, Math.max(usedB.rows + 200, layout.rows.at(scroll.y + size.h / zoom) + 100, 1000));
  const extentCols = Math.min(MAX_COLS, Math.max(usedB.cols + 30, layout.cols.at(scroll.x + size.w / zoom) + 20, 52));
  const contentH = Math.min(totalH, layout.rows.start(extentRows)), contentW = Math.min(totalW, layout.cols.start(extentCols));
  const spacerH = Math.min(MAX_SCROLL, contentH * zoom + headerH), spacerW = Math.min(MAX_SCROLL, contentW * zoom + headerW);
  const scaleY = spacerH >= MAX_SCROLL ? contentH / ((MAX_SCROLL - headerH) / zoom) : 1;
  const scaleX = spacerW >= MAX_SCROLL ? contentW / ((MAX_SCROLL - headerW) / zoom) : 1;

  const vp: Viewport = {
    layout, zoom, headerW, headerH, width: size.w, height: size.h,
    scrollX: scroll.x, scrollY: scroll.y, frozenRows: sheet.view.frozenRows, frozenCols: sheet.view.frozenCols,
  };

  // ---------- Size and scroll ----------
  useLayoutEffect(() => {
    const el = wrap.current!;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  const onScroll = useCallback(() => {
    const el = scroller.current!;
    setScroll({ x: (el.scrollLeft / zoom) * scaleX, y: (el.scrollTop / zoom) * scaleY });
    if (list) setList(null);
  }, [zoom, scaleX, scaleY, list]);

  const scrollTo = useCallback((row: number, col: number) => {
    const el = scroller.current;
    if (!el) return;
    const fr = sheet.view.frozenRows, fc = sheet.view.frozenCols;
    const viewH = (size.h - headerH) / zoom - layout.rows.start(fr), viewW = (size.w - headerW) / zoom - layout.cols.start(fc);
    let { x, y } = scroll;
    if (row >= fr) {
      const top = layout.rows.start(row) - layout.rows.start(fr), bottom = layout.rows.start(row + 1) - layout.rows.start(fr);
      if (top < y) y = top; else if (bottom > y + viewH) y = bottom - viewH;
    }
    if (col >= fc) {
      const left = layout.cols.start(col) - layout.cols.start(fc), right = layout.cols.start(col + 1) - layout.cols.start(fc);
      if (left < x) x = left; else if (right > x + viewW) x = Math.min(left, right - viewW);
    }
    if (x !== scroll.x || y !== scroll.y) {
      el.scrollTop = (Math.max(0, y) * zoom) / scaleY;
      el.scrollLeft = (Math.max(0, x) * zoom) / scaleX;
      setScroll({ x: Math.max(0, x), y: Math.max(0, y) });
    }
  }, [scroll, size, zoom, layout, headerH, headerW, sheet.view.frozenRows, sheet.view.frozenCols, scaleX, scaleY]);
  if (apiRef) apiRef.current = { scrollTo };

  // Keep the active cell in view when it moves.
  const active = store.sel.active;
  const lastActive = useRef('');
  useEffect(() => {
    const k = `${store.sheetIndex}:${active.row}:${active.col}`;
    if (k !== lastActive.current) { lastActive.current = k; scrollTo(active.row, active.col); }
  });

  // Reset the scroll when switching sheets.
  useEffect(() => { const el = scroller.current; if (el) { el.scrollTop = 0; el.scrollLeft = 0; } setScroll({ x: 0, y: 0 }); }, [store.sheetIndex]);

  // ---------- Drawing ----------
  const editRefs = store.editing ? refsInFormula(store.editing.text) : [];
  const overlay: Overlay = {
    refs: editRefs.map((t, i) => {
      const local = t.replace(/^.*!/, '');
      const r = parseRange(local.replace(/\$/g, ''));
      const onOther = t.includes('!') && !t.toLowerCase().startsWith(sheet.name.toLowerCase()) && !t.toLowerCase().startsWith(`'${sheet.name.toLowerCase()}'`);
      return r && !onOther ? { range: r, color: REF_COLORS[i % REF_COLORS.length]! } : null;
    }).filter((x): x is { range: CellRange; color: string } => !!x),
    fillPreview: dragState.fill ?? null,
    dragRange: dragState.move ?? null,
    marqueePhase,
    editing: store.editing && !store.editing.inBar ? store.editing : store.editing ? { row: -1, col: -1 } : null,
    showFormulas: store.view.showFormulas,
  };

  useEffect(() => {
    const c = canvas.current;
    if (!c) return;
    const dpr = window.devicePixelRatio || 1;
    if (c.width !== Math.round(size.w * dpr) || c.height !== Math.round(size.h * dpr)) { c.width = Math.round(size.w * dpr); c.height = Math.round(size.h * dpr); }
    const ctx = c.getContext('2d')!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    renderer.current ??= new SheetRenderer(ctx);
    renderer.current.draw(store, vp, readTheme(wrap.current!), overlay);
  });

  // Animated copy marquee.
  useEffect(() => {
    if (!store.marquee) return;
    const t = setInterval(() => setPhase((p) => (p + 1) % 18), 90);
    return () => clearInterval(t);
  }, [store.marquee]);

  // Repaint when fonts finish loading.
  useEffect(() => {
    const f = () => store.emit();
    document.fonts?.addEventListener?.('loadingdone', f);
    return () => document.fonts?.removeEventListener?.('loadingdone', f);
  }, [store]);

  // ---------- Hit testing ----------
  const hit = (clientX: number, clientY: number): { area: 'cell' | 'colHeader' | 'rowHeader' | 'corner'; row: number; col: number; x: number; y: number } => {
    const rect = wrap.current!.getBoundingClientRect();
    const x = clientX - rect.left, y = clientY - rect.top;
    const fh = layout.rows.start(vp.frozenRows) * zoom, fw = layout.cols.start(vp.frozenCols) * zoom;
    const sx = x - headerW < fw ? (x - headerW) / zoom : (x - headerW - fw) / zoom + layout.cols.start(vp.frozenCols) + scroll.x;
    const sy = y - headerH < fh ? (y - headerH) / zoom : (y - headerH - fh) / zoom + layout.rows.start(vp.frozenRows) + scroll.y;
    const col = layout.cols.at(Math.max(0, sx)), row = layout.rows.at(Math.max(0, sy));
    const area = x < headerW && y < headerH ? 'corner' : y < headerH ? 'colHeader' : x < headerW ? 'rowHeader' : 'cell';
    return { area, row, col, x, y };
  };

  const resizeEdge = (h: ReturnType<typeof hit>): { axis: 'row' | 'col'; index: number } | null => {
    if (h.area === 'colHeader') {
      const panes = panesOf(vp);
      const p = paneFor(panes, vp, 0, h.col);
      const left = rectIn(p, vp, { r1: 0, c1: h.col, r2: 0, c2: h.col });
      if (Math.abs(h.x - (left.x + left.w)) <= 4) return { axis: 'col', index: h.col };
      if (Math.abs(h.x - left.x) <= 4 && h.col > 0) return { axis: 'col', index: h.col - 1 };
    }
    if (h.area === 'rowHeader') {
      const panes = panesOf(vp);
      const p = paneFor(panes, vp, h.row, 0);
      const top = rectIn(p, vp, { r1: h.row, c1: 0, r2: h.row, c2: 0 });
      if (Math.abs(h.y - (top.y + top.h)) <= 3) return { axis: 'row', index: h.row };
      if (Math.abs(h.y - top.y) <= 3 && h.row > 0) return { axis: 'row', index: h.row - 1 };
    }
    return null;
  };

  const onFillHandle = (h: ReturnType<typeof hit>): boolean => {
    if (store.sel.ranges.length !== 1 || store.editing) return false;
    const r = store.range;
    const p = paneFor(panesOf(vp), vp, r.r2, r.c2);
    const rr = rectIn(p, vp, r);
    return Math.abs(h.x - (rr.x + rr.w)) <= 6 && Math.abs(h.y - (rr.y + rr.h)) <= 6;
  };

  const onSelBorder = (h: ReturnType<typeof hit>): boolean => {
    if (store.sel.ranges.length !== 1 || store.editing) return false;
    const r = store.range;
    if (r.r2 - r.r1 > 5000 || r.c2 - r.c1 > 1000) return false;
    const p = paneFor(panesOf(vp), vp, r.r1, r.c1);
    const rr = rectIn(p, vp, r);
    const nearX = Math.abs(h.x - rr.x) <= 3 || Math.abs(h.x - (rr.x + rr.w)) <= 3;
    const nearY = Math.abs(h.y - rr.y) <= 3 || Math.abs(h.y - (rr.y + rr.h)) <= 3;
    const inX = h.x >= rr.x - 3 && h.x <= rr.x + rr.w + 3, inY = h.y >= rr.y - 3 && h.y <= rr.y + rr.h + 3;
    return (nearX && inY) || (nearY && inX);
  };

  // ---------- Pointer ----------
  const autoScrollTimer = useRef<ReturnType<typeof setInterval> | null>(null);
  const lastPointer = useRef({ x: 0, y: 0 });

  const insertRefWhileEditing = (pos: Pos, extendTo?: Pos) => {
    const e = store.editing!;
    const ta = editor.current;
    const caret = ta?.selectionStart ?? e.text.length;
    const ptr = pointer.current;
    const a = rangeOf(pos, extendTo ?? pos);
    const text = a.r1 === a.r2 && a.c1 === a.c2 ? `${colName(a.c1)}${a.r1 + 1}` : `${colName(a.c1)}${a.r1 + 1}:${colName(a.c2)}${a.r2 + 1}`;
    const start = ptr?.mode === 'pickRef' && ptr.refStart !== undefined ? ptr.refStart : caret;
    const len = ptr?.mode === 'pickRef' && ptr.refLen !== undefined ? ptr.refLen : 0;
    const sheetPrefix = store.sheetIndex !== editSheet.current ? `${quoteName(sheet.name)}!` : '';
    const ins = sheetPrefix + text;
    const next = e.text.slice(0, start) + ins + e.text.slice(start + len);
    store.setEditText(next);
    if (pointer.current) { pointer.current.refStart = start; pointer.current.refLen = ins.length; }
    requestAnimationFrame(() => { const t = editor.current; if (t) { t.focus(); t.setSelectionRange(start + ins.length, start + ins.length); } });
  };
  const editSheet = useRef(store.sheetIndex);

  const pointingFormula = (): boolean => {
    const e = store.editing;
    if (!e || !e.text.startsWith('=')) return false;
    const caret = editor.current?.selectionStart ?? e.text.length;
    const before = e.text.slice(0, caret).trimEnd();
    return /[=(,+\-*/^&<>:;{]$/.test(before) || (pointer.current?.mode === 'pickRef');
  };

  const onPointerDown = (ev: React.PointerEvent) => {
    if (ev.button === 2) return;
    const h = hit(ev.clientX, ev.clientY);
    wrap.current?.focus({ preventScroll: true });
    setList(null);
    // Touch: a tap selects, dragging scrolls (the browser handles it).
    if (ev.pointerType === 'touch') { touchStart.current = { x: ev.clientX, y: ev.clientY, t: Date.now(), h }; return; }
    (ev.target as Element).setPointerCapture?.(ev.pointerId);
    lastPointer.current = { x: ev.clientX, y: ev.clientY };
    const edge = resizeEdge(h);
    if (edge) {
      const orig = edge.axis === 'col' ? layout.cols.size(edge.index) : layout.rows.size(edge.index);
      pointer.current = { mode: 'resize', start: { row: h.row, col: h.col } };
      setDragState({ resize: { ...edge, start: edge.axis === 'col' ? h.x : h.y, orig, px: orig } });
      return;
    }
    if (h.area === 'corner') { store.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, store.sel.active); return; }
    if (h.area === 'cell' && af(h)) return;
    if (store.editing && pointingFormula() && h.area === 'cell') {
      pointer.current = { mode: 'pickRef', start: { row: h.row, col: h.col } };
      insertRefWhileEditing({ row: h.row, col: h.col });
      return;
    }
    if (store.editing) { if (!store.commitEdit()) return; }
    if (h.area === 'colHeader' || h.area === 'rowHeader') {
      const axis = h.area === 'colHeader' ? 'col' : 'row';
      const r = axis === 'col' ? { r1: 0, c1: h.col, r2: MAX_ROWS - 1, c2: h.col } : { r1: h.row, c1: 0, r2: h.row, c2: MAX_COLS - 1 };
      if (ev.shiftKey) {
        const a = store.sel.anchor;
        store.selectRange(axis === 'col' ? { r1: 0, c1: Math.min(a.col, h.col), r2: MAX_ROWS - 1, c2: Math.max(a.col, h.col) } : { r1: Math.min(a.row, h.row), c1: 0, r2: Math.max(a.row, h.row), c2: MAX_COLS - 1 }, store.sel.active);
      } else {
        const act = axis === 'col' ? { row: Math.max(0, layout.rows.at(scroll.y + layout.rows.start(sheet.view.frozenRows))), col: h.col } : { row: h.row, col: Math.max(0, layout.cols.at(scroll.x + layout.cols.start(sheet.view.frozenCols))) };
        if (ev.ctrlKey || ev.metaKey) store.sel = { ranges: [...store.sel.ranges, r], active: act, anchor: { row: h.row, col: h.col } };
        else store.sel = { ranges: [r], active: act, anchor: { row: h.row, col: h.col } };
        store.emit();
      }
      pointer.current = { mode: 'header', start: { row: h.row, col: h.col }, headerAxis: axis };
      return;
    }
    if (onFillHandle(h)) { pointer.current = { mode: 'fill', start: { row: h.row, col: h.col } }; setDragState({ fill: store.range }); return; }
    if (onSelBorder(h)) { pointer.current = { mode: 'move', start: { row: h.row, col: h.col } }; setDragState({ move: store.range }); return; }
    store.select({ row: h.row, col: h.col }, ev.shiftKey, ev.ctrlKey || ev.metaKey);
    pointer.current = { mode: 'select', start: { row: h.row, col: h.col } };
  };

  const touchStart = useRef<{ x: number; y: number; t: number; h: ReturnType<typeof hit> } | null>(null);
  const lastTap = useRef<{ t: number; row: number; col: number } | null>(null);

  /** Clicking a filter button or a validation list arrow. Returns true if handled. */
  const af = (h: ReturnType<typeof hit>): boolean => {
    const f = filterPopupFor(store, h.row, h.col);
    if (f) {
      const p = paneFor(panesOf(vp), vp, h.row, h.col);
      const rr = rectIn(p, vp, { r1: h.row, c1: h.col, r2: h.row, c2: h.col });
      if (h.x >= rr.x + rr.w - 18 * zoom && h.y >= rr.y + rr.h - 18 * zoom) {
        store.select({ row: h.row, col: h.col });
        setFilterCol({ col: h.col, x: rr.x + rr.w - 220 < 0 ? rr.x : rr.x + rr.w - 220, y: rr.y + rr.h });
        return true;
      }
    }
    return false;
  };

  const onPointerMove = (ev: React.PointerEvent) => {
    const h = hit(ev.clientX, ev.clientY);
    lastPointer.current = { x: ev.clientX, y: ev.clientY };
    const ptr = pointer.current;
    if (!ptr) {
      // Cursor feedback.
      const el = wrap.current!;
      const edge = resizeEdge(h);
      el.style.cursor = edge ? (edge.axis === 'col' ? 'col-resize' : 'row-resize') : h.area === 'colHeader' ? 'url("data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 width=%2716%27 height=%2716%27%3E%3Cpath d=%27M8 14V2M4 10l4 4 4-4%27 stroke=%27black%27 stroke-width=%272%27 fill=%27none%27/%3E%3C/svg%3E") 8 14, s-resize'
        : h.area === 'rowHeader' ? 'url("data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 width=%2716%27 height=%2716%27%3E%3Cpath d=%27M14 8H2M10 4l4 4-4 4%27 stroke=%27black%27 stroke-width=%272%27 fill=%27none%27/%3E%3C/svg%3E") 14 8, e-resize'
          : onFillHandle(h) ? 'crosshair' : onSelBorder(h) ? 'move' : painter.style ? 'copy' : 'cell';
      return;
    }
    edgeScroll(ev.clientX, ev.clientY);
    applyDrag(h);
  };

  const applyDrag = (h: ReturnType<typeof hit>) => {
    const ptr = pointer.current;
    if (!ptr) return;
    if (ptr.mode === 'resize' && dragState.resize) {
      const d = dragState.resize;
      const delta = ((d.axis === 'col' ? h.x : h.y) - d.start) / zoom;
      setDragState({ resize: { ...d, px: Math.max(0, d.orig + delta) } });
      return;
    }
    if (ptr.mode === 'select') store.select({ row: h.row, col: h.col }, true);
    else if (ptr.mode === 'header') {
      const s = store.sel;
      const r = ptr.headerAxis === 'col' ? { r1: 0, c1: Math.min(ptr.start.col, h.col), r2: MAX_ROWS - 1, c2: Math.max(ptr.start.col, h.col) } : { r1: Math.min(ptr.start.row, h.row), c1: 0, r2: Math.max(ptr.start.row, h.row), c2: MAX_COLS - 1 };
      store.sel = { ...s, ranges: [...s.ranges.slice(0, -1), r] };
      store.emit();
    } else if (ptr.mode === 'pickRef') insertRefWhileEditing(ptr.start, { row: h.row, col: h.col });
    else if (ptr.mode === 'fill') {
      const r = store.range;
      const dr = h.row > r.r2 ? h.row - r.r2 : h.row < r.r1 ? h.row - r.r1 : 0;
      const dc = h.col > r.c2 ? h.col - r.c2 : h.col < r.c1 ? h.col - r.c1 : 0;
      let t: CellRange = r;
      if (Math.abs(dr) >= Math.abs(dc) && dr) t = dr > 0 ? { ...r, r2: h.row } : { ...r, r1: h.row };
      else if (dc) t = dc > 0 ? { ...r, c2: h.col } : { ...r, c1: h.col };
      else if (h.row < r.r2 && h.row >= r.r1 && r.r2 > r.r1) t = { ...r, r2: h.row };
      setDragState({ fill: t });
    } else if (ptr.mode === 'move') {
      const r = store.range;
      const dr = h.row - ptr.start.row, dc = h.col - ptr.start.col;
      setDragState({ move: { r1: Math.max(0, r.r1 + dr), c1: Math.max(0, r.c1 + dc), r2: Math.max(0, r.r1 + dr) + r.r2 - r.r1, c2: Math.max(0, r.c1 + dc) + r.c2 - r.c1 } });
    }
  };

  const edgeScroll = (cx: number, cy: number) => {
    const rect = wrap.current!.getBoundingClientRect();
    const dx = cx < rect.left + headerW ? -1 : cx > rect.right - 10 ? 1 : 0;
    const dy = cy < rect.top + headerH ? -1 : cy > rect.bottom - 10 ? 1 : 0;
    if (!dx && !dy) { if (autoScrollTimer.current) { clearInterval(autoScrollTimer.current); autoScrollTimer.current = null; } return; }
    if (autoScrollTimer.current) return;
    autoScrollTimer.current = setInterval(() => {
      const el = scroller.current!;
      el.scrollTop += dy * 30; el.scrollLeft += dx * 30;
      applyDrag(hit(lastPointer.current.x, lastPointer.current.y));
    }, 40);
  };

  const onPointerUp = (ev: React.PointerEvent) => {
    if (autoScrollTimer.current) { clearInterval(autoScrollTimer.current); autoScrollTimer.current = null; }
    if (ev.pointerType === 'touch' && touchStart.current) {
      const t = touchStart.current;
      touchStart.current = null;
      if (Math.hypot(ev.clientX - t.x, ev.clientY - t.y) > 10 || Date.now() - t.t > 600) return;
      const h = t.h;
      if (h.area === 'cell' && af(h)) return;
      if (store.editing && pointingFormula() && h.area === 'cell') { insertRefWhileEditing({ row: h.row, col: h.col }); return; }
      if (store.editing && !store.commitEdit()) return;
      if (h.area === 'colHeader') { store.selectRange({ r1: 0, c1: h.col, r2: MAX_ROWS - 1, c2: h.col }); return; }
      if (h.area === 'rowHeader') { store.selectRange({ r1: h.row, c1: 0, r2: h.row, c2: MAX_COLS - 1 }); return; }
      const dbl = lastTap.current && Date.now() - lastTap.current.t < 400 && lastTap.current.row === h.row && lastTap.current.col === h.col;
      lastTap.current = { t: Date.now(), row: h.row, col: h.col };
      if (dbl) { store.beginEdit(); return; }
      if (inRange(store.range, h.row, h.col) && store.range.r1 !== store.range.r2 && false) return;
      store.select({ row: h.row, col: h.col });
      if (painter.style) applyPainter(store);
      return;
    }
    const ptr = pointer.current;
    pointer.current = null;
    if (!ptr) return;
    if (ptr.mode === 'resize' && dragState.resize) {
      const d = dragState.resize;
      setDragState({});
      const targets = (d.axis === 'col' ? store.sel.ranges.some((r) => r.r1 === 0 && r.r2 === MAX_ROWS - 1 && d.index >= r.c1 && d.index <= r.c2) : store.sel.ranges.some((r) => r.c1 === 0 && r.c2 === MAX_COLS - 1 && d.index >= r.r1 && d.index <= r.r2))
        ? store.sel.ranges.flatMap((r) => { const out: number[] = []; const a = d.axis === 'col' ? [r.c1, r.c2] : [r.r1, r.r2]; for (let i = a[0]!; i <= a[1]! && out.length < 5000; i++) out.push(i); return out; })
        : [d.index];
      store.change(d.axis === 'col' ? 'Column width' : 'Row height', 'sheet', () => {
        for (const i of targets) {
          if (d.axis === 'col') { const info = sheet.cols.get(i) ?? {}; info.width = pxToChars(d.px); info.hidden = d.px === 0 ? true : undefined; sheet.cols.set(i, info); }
          else { const info = sheet.rows.get(i) ?? {}; info.height = pxToPt(d.px); info.hidden = d.px === 0 ? true : undefined; sheet.rows.set(i, info); }
        }
      });
      return;
    }
    if (ptr.mode === 'fill' && dragState.fill) {
      const t = dragState.fill;
      setDragState({});
      const r = store.range;
      if (t.r1 !== r.r1 || t.r2 !== r.r2 || t.c1 !== r.c1 || t.c2 !== r.c2) autoFill(store, r, t, ev.ctrlKey || ev.altKey);
      return;
    }
    if (ptr.mode === 'move' && dragState.move) {
      const t = dragState.move;
      setDragState({});
      const r = store.range;
      if (t.r1 !== r.r1 || t.c1 !== r.c1) void import('../app/ops/move').then((m) => m.moveRange(store, r, { row: t.r1, col: t.c1 }, ev.ctrlKey || ev.altKey));
      return;
    }
    if (ptr.mode === 'select' && painter.style) applyPainter(store);
    if (ptr.mode === 'pickRef') { editor.current?.focus(); }
  };

  const onDoubleClick = (ev: React.MouseEvent) => {
    const h = hit(ev.clientX, ev.clientY);
    const edge = resizeEdge(h);
    if (edge) {
      store.selectRange(edge.axis === 'col' ? { r1: 0, c1: edge.index, r2: MAX_ROWS - 1, c2: edge.index } : { r1: edge.index, c1: 0, r2: edge.index, c2: MAX_COLS - 1 });
      run(store, edge.axis === 'col' ? 'cells.autofitCols' : 'cells.autofitRows');
      return;
    }
    if (h.area !== 'cell') return;
    if (onFillHandle(h)) {
      // Double-clicking the fill handle fills down as far as the neighbouring column has data.
      const r = store.range;
      let end = r.r2;
      const side = r.c1 > 0 && store.value(r.r2 + 1, r.c1 - 1) !== null ? r.c1 - 1 : r.c2 + 1;
      while (end + 1 < MAX_ROWS && store.value(end + 1, side) !== null) end++;
      if (end > r.r2) autoFill(store, r, { ...r, r2: end });
      return;
    }
    const link = sheet.hyperlinks.get(cellKey(h.row, h.col));
    if (link && (ev.ctrlKey || ev.metaKey)) { openLink(store, link); return; }
    store.beginEdit();
  };

  const onContextMenu = (ev: React.MouseEvent) => {
    ev.preventDefault();
    const h = hit(ev.clientX, ev.clientY);
    if (h.area === 'cell' && !store.sel.ranges.some((r) => inRange(r, h.row, h.col))) store.select({ row: h.row, col: h.col });
    if (h.area === 'colHeader' && !store.sel.ranges.some((r) => r.r1 === 0 && r.r2 === MAX_ROWS - 1 && h.col >= r.c1 && h.col <= r.c2)) store.selectRange({ r1: 0, c1: h.col, r2: MAX_ROWS - 1, c2: h.col });
    if (h.area === 'rowHeader' && !store.sel.ranges.some((r) => r.c1 === 0 && r.c2 === MAX_COLS - 1 && h.row >= r.r1 && h.row <= r.r2)) store.selectRange({ r1: h.row, c1: 0, r2: h.row, c2: MAX_COLS - 1 });
    store.menu = { x: ev.clientX, y: ev.clientY, kind: h.area === 'colHeader' ? 'col' : h.area === 'rowHeader' ? 'row' : 'cell' };
    store.emit();
  };

  // ---------- Keyboard ----------
  const onKeyDown = (ev: React.KeyboardEvent) => {
    // Keys from the cell editor bubble up here after it has handled them.
    if (ev.target === editor.current || (ev.target as Element).closest?.('.dvlist,.filterpop,.objlayer')) return;
    const k = ev.key;
    const printable = k.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !(k === ' ' && ev.shiftKey) && !ev.nativeEvent.isComposing;
    if (printable && !(ev.target instanceof HTMLInputElement) && !(ev.target instanceof HTMLTextAreaElement)) {
      ev.preventDefault();
      // Keys that arrive before the editor has focus still land in it.
      if (store.editing && !store.editing.inBar) store.setEditText(store.editing.text + k);
      else store.beginEdit(k);
      return;
    }
    if (store.editing && !store.editing.inBar) {
      // The editor is still opening: treat commit keys the way the editor would.
      const moves: Record<string, [number, number]> = { Enter: [ev.shiftKey ? -1 : 1, 0], Tab: [0, ev.shiftKey ? -1 : 1], ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (k === 'Escape') { ev.preventDefault(); store.cancelEdit(); return; }
      if (k === 'Backspace') { ev.preventDefault(); store.setEditText(store.editing.text.slice(0, -1)); return; }
      const m = moves[k];
      if (m) { ev.preventDefault(); if (store.commitEdit()) store.move(m[0], m[1], k === 'Enter' || k === 'Tab'); return; }
    }
    if (handleGridKey(store, ev.nativeEvent, { pageRows: Math.max(1, Math.floor((size.h - headerH) / (20 * zoom)) - 1) })) {
      ev.preventDefault();
      if (store.marquee && ev.key === 'Escape') clearClipboardMarquee(store);
    }
  };

  // Typing into the grid starts editing with that character.
  const onBeforeInput = (ev: React.FormEvent<HTMLDivElement>) => {
    const data = (ev.nativeEvent as InputEvent).data;
    if (!store.editing && data && !(ev.target instanceof HTMLTextAreaElement)) {
      ev.preventDefault();
      store.beginEdit(data);
    }
  };

  // Focus the in-cell editor when editing starts.
  useEffect(() => {
    const e = store.editing;
    if (e && !e.inBar) {
      editSheet.current = store.sheetIndex;
      const t = editor.current;
      if (t && document.activeElement !== t) { t.focus({ preventScroll: true }); t.setSelectionRange(t.value.length, t.value.length); }
      const dv = sheet.validations.find((v) => v.type === 'list' && v.showDropdown !== false && v.ranges.some((r) => inRange(r, e.row, e.col)));
      void dv;
    } else if (!e && document.activeElement === editor.current) wrap.current?.focus({ preventScroll: true });
  }, [store.editing, store.sheetIndex, sheet.validations]);

  // Validation list arrow next to the active cell.
  const dv = sheet.validations.find((v) => v.type === 'list' && v.showDropdown !== false && v.ranges.some((r) => inRange(r, active.row, active.col)));
  const activeRect = (() => {
    const panes = panesOf(vp);
    const m = store.mergeAt(active.row, active.col) ?? { r1: active.row, c1: active.col, r2: active.row, c2: active.col };
    return rectIn(paneFor(panes, vp, m.r1, m.c1), vp, m);
  })();

  // ---------- Editor ----------
  const e = store.editing;
  const editStyle = e ? store.style(e.row, e.col) : null;
  const editRect = e ? (() => {
    const panes = panesOf(vp);
    const m = store.mergeAt(e.row, e.col) ?? { r1: e.row, c1: e.col, r2: e.row, c2: e.col };
    return rectIn(paneFor(panes, vp, m.r1, m.c1), vp, m);
  })() : null;

  const onEditorKey = (ev: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ed = store.editing;
    if (!ed) return;
    if (assistKey.current?.(ev.nativeEvent)) { ev.preventDefault(); return; }
    const mod = ev.ctrlKey || ev.metaKey;
    if (ev.key === 'Enter' && (ev.altKey || (ev.metaKey && ev.ctrlKey))) {
      ev.preventDefault();
      const t = ev.currentTarget, pos = t.selectionStart;
      store.setEditText(ed.text.slice(0, pos) + '\n' + ed.text.slice(t.selectionEnd));
      requestAnimationFrame(() => t.setSelectionRange(pos + 1, pos + 1));
      return;
    }
    if (ev.key === 'Enter') {
      ev.preventDefault();
      if (mod && ev.shiftKey) { store.commitEdit(false, true); return; }
      if (!store.commitEdit(mod)) return;
      if (!mod) store.move(ev.shiftKey ? -1 : 1, 0, true);
      wrap.current?.focus({ preventScroll: true });
      return;
    }
    if (ev.key === 'Tab') {
      ev.preventDefault();
      if (!store.commitEdit()) return;
      store.move(0, ev.shiftKey ? -1 : 1, true);
      wrap.current?.focus({ preventScroll: true });
      return;
    }
    if (ev.key === 'Escape') { ev.preventDefault(); store.cancelEdit(); wrap.current?.focus({ preventScroll: true }); return; }
    if (ev.key === 'F4') { ev.preventDefault(); cycleAbsolute(ev.currentTarget, store); return; }
    if (ed.replace && !ed.text.startsWith('=') && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(ev.key)) {
      ev.preventDefault();
      if (!store.commitEdit()) return;
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[ev.key]!;
      store.move(d[0]!, d[1]!);
      wrap.current?.focus({ preventScroll: true });
      return;
    }
    if (ed.replace && ed.text.startsWith('=') && pointingFormula() && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(ev.key)) {
      // Arrow keys point at cells while building a formula.
      ev.preventDefault();
      const ptr = pointer.current ?? { mode: 'pickRef' as const, start: { ...store.sel.active } };
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[ev.key]!;
      const base = ptr.mode === 'pickRef' && pointer.current ? ptr.start : { row: ed.row, col: ed.col };
      const next = { row: Math.max(0, base.row + d[0]!), col: Math.max(0, base.col + d[1]!) };
      pointer.current = { ...ptr, mode: 'pickRef', start: next };
      insertRefWhileEditing(next);
      return;
    }
    if (pointer.current?.mode === 'pickRef' && !['Shift', 'Control', 'Meta', 'Alt'].includes(ev.key)) pointer.current = null;
  };
  const assistKey = useRef<((e: KeyboardEvent) => boolean) | null>(null);

  const resizeLine = dragState.resize ? (() => {
    const d = dragState.resize;
    const panes = panesOf(vp);
    if (d.axis === 'col') { const p = paneFor(panes, vp, 0, d.index); const r = rectIn(p, vp, { r1: 0, c1: d.index, r2: 0, c2: d.index }); return { left: r.x + d.px * zoom, top: 0, width: 1, height: size.h, label: `Width: ${pxToChars(d.px)} (${Math.round(d.px)} pixels)` }; }
    const p = paneFor(panes, vp, d.index, 0); const r = rectIn(p, vp, { r1: d.index, c1: 0, r2: d.index, c2: 0 });
    return { left: 0, top: r.y + d.px * zoom, width: size.w, height: 1, label: `Height: ${pxToPt(d.px)} (${Math.round(d.px)} pixels)` };
  })() : null;

  return (
    <div className="gridwrap" ref={wrap} tabIndex={0} role="grid" aria-label={`${sheet.name} worksheet`} aria-rowcount={MAX_ROWS} aria-colcount={MAX_COLS}
      onKeyDown={onKeyDown} onBeforeInput={onBeforeInput}>
      <div className="scroller" ref={scroller} onScroll={onScroll}
        onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={() => { pointer.current = null; touchStart.current = null; setDragState({}); }}
        onDoubleClick={onDoubleClick} onContextMenu={onContextMenu}
        onWheel={(ev) => { if (ev.ctrlKey || ev.metaKey) { ev.preventDefault(); run(store, ev.deltaY < 0 ? 'view.zoomIn' : 'view.zoomOut'); } }}>
        <div className="spacer" style={{ width: spacerW, height: spacerH }}>
          <canvas ref={canvas} className="sheetcanvas" style={{ width: size.w, height: size.h }} />
        </div>
      </div>
      <ChartLayer store={store} vp={vp} />
      {dv && !e && (
        <button className="dvarrow" style={{ left: activeRect.x + activeRect.w + 1, top: activeRect.y + activeRect.h - 18 * zoom, height: 18 * zoom, width: 18 * zoom }}
          aria-label="Choose from the list" onPointerDown={(ev) => ev.stopPropagation()} onClick={() => setList(list ? null : { row: active.row, col: active.col, options: listOptions(store, dv, active.row, active.col) })}>
          <svg viewBox="0 0 10 10" width="10" height="10"><path d="M2 3.5 5 6.5 8 3.5" stroke="currentColor" strokeWidth="1.6" fill="none" strokeLinecap="round" /></svg>
        </button>
      )}
      {list && (
        <div className="dvlist" role="listbox" style={{ left: activeRect.x, top: activeRect.y + activeRect.h + 2, minWidth: activeRect.w }}>
          {list.options.map((o) => (
            <button key={o} role="option" aria-selected={false} onClick={() => { store.change('Choose', [{ area: { r1: list.row, c1: list.col, r2: list.row, c2: list.col } }], () => store.enter(list.row, list.col, o)); setList(null); wrap.current?.focus(); }}>{o}</button>
          ))}
          {!list.options.length && <div className="muted">No choices</div>}
        </div>
      )}
      {e && !e.inBar && editRect && editStyle && (
        <>
          <textarea ref={editor} className="celleditor" spellCheck={false} value={e.text} aria-label="Cell editor"
            style={{
              left: editRect.x, top: editRect.y, minWidth: editRect.w, minHeight: editRect.h,
              font: cssFont(editStyle, zoom), color: editStyle.font?.color ?? undefined, background: editStyle.fill?.color ?? undefined,
              textAlign: editStyle.align?.h === 'center' ? 'center' : editStyle.align?.h === 'right' ? 'right' : 'left',
              width: Math.min(size.w - editRect.x - 4, Math.max(editRect.w, measureWidth(e.text, cssFont(editStyle, zoom)) + 14)),
              height: Math.max(editRect.h, (e.text.split('\n').length) * (((editStyle.font?.size ?? 11) * 4) / 3) * zoom * 1.25 + 6),
              whiteSpace: editStyle.align?.wrap ? 'pre-wrap' : 'pre',
            }}
            onChange={(ev) => store.setEditText(ev.target.value)} onKeyDown={onEditorKey}
            onBlur={(ev) => { if (ev.relatedTarget && (ev.relatedTarget as Element).closest?.('.formulabar,.assist,.dialog,.panel,.toolpanel,.topbar')) return; }} />
          <FormulaAssist store={store} anchor={{ x: editRect.x, y: editRect.y + Math.max(editRect.h, 22) + 2 }} input={editor} keyRef={assistKey} />
        </>
      )}
      {resizeLine && <><div className="resizeline" style={resizeLine} /><div className="resizetip" style={{ left: Math.min(resizeLine.left + 8, size.w - 200), top: Math.min(resizeLine.top + 8, size.h - 30) }}>{resizeLine.label}</div></>}
      {filterCol && (
        <FilterPopupHost store={store} col={filterCol.col} x={filterCol.x} y={filterCol.y} onClose={() => { setFilterCol(null); wrap.current?.focus(); }} />
      )}
    </div>
  );
};

import { FilterPopup } from './FilterPopup';
const FilterPopupHost = (p: { store: GridStore; col: number; x: number; y: number; onClose: () => void }) => <FilterPopup {...p} />;

let measureCtx: CanvasRenderingContext2D | null = null;
const measureWidth = (text: string, font: string): number => {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  if (!measureCtx) return text.length * 7;
  measureCtx.font = font;
  return Math.max(...text.split('\n').map((l) => measureCtx!.measureText(l).width));
};

const quoteName = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

/** F4 while editing: A1 → $A$1 → A$1 → $A1 → A1 for the reference at the caret. */
const cycleAbsolute = (t: HTMLTextAreaElement | HTMLInputElement, store: GridStore): void => {
  const text = t.value, pos = t.selectionStart ?? text.length;
  const re = /\$?[A-Za-z]{1,3}\$?\d+/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    if (pos >= m.index && pos <= m.index + m[0].length) {
      const raw = m[0];
      const p = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(raw)!;
      const state = (p[1] ? 2 : 0) + (p[3] ? 1 : 0);
      const next = { 0: [true, true], 3: [false, true], 1: [true, false], 2: [false, false] }[state as 0 | 1 | 2 | 3]!;
      const rep = `${next[0] ? '$' : ''}${p[2]}${next[1] ? '$' : ''}${p[4]}`;
      const out = text.slice(0, m.index) + rep + text.slice(m.index + raw.length);
      store.setEditText(out);
      requestAnimationFrame(() => t.setSelectionRange(m!.index + rep.length, m!.index + rep.length));
      return;
    }
  }
};
export { cycleAbsolute };

export const openLink = (store: GridStore, link: string): void => {
  if (link.startsWith('#')) {
    const target = link.slice(1);
    const [sh, ref] = target.includes('!') ? [target.slice(0, target.lastIndexOf('!')).replace(/^'|'$/g, ''), target.slice(target.lastIndexOf('!') + 1)] : [store.sheet.name, target];
    const si = store.engine.sheetIndex(sh);
    const r = parseRange(ref.replace(/\$/g, ''));
    if (si >= 0) store.setActiveSheet(si);
    if (r) store.selectRange(r);
    return;
  }
  if (/^(https?:|mailto:)/i.test(link)) window.open(link, '_blank', 'noopener');
};
