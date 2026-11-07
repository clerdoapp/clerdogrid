import { useRef, useState } from 'react';
import type { InkStroke } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import type { Viewport } from '../../ui/render';
import { lastVp, toScreen, toSheet, ui } from './state';
import { addStroke, eraseStrokes, newId, strokesAt } from './ops';

/** The default pen color shows as the sheet's ink color, so it stays visible in dark mode. */
export const DEFAULT_PEN = '#1f2a24';

const TOOL_STYLE: Record<InkStroke['tool'], { opacity: number; cap: 'round' | 'butt' }> = {
  pen: { opacity: 1, cap: 'round' }, marker: { opacity: 0.92, cap: 'round' }, highlighter: { opacity: 0.38, cap: 'butt' },
};

/** Smooth path through the points (quadratic curves via midpoints), in screen pixels. */
const pathOf = (pts: [number, number][]): string => {
  if (!pts.length) return '';
  if (pts.length === 1) { const [x, y] = pts[0]!; return `M${x},${y}l0.01,0`; }
  let d = `M${pts[0]![0]},${pts[0]![1]}`;
  for (let i = 1; i < pts.length - 1; i++) {
    const [x, y] = pts[i]!, [nx, ny] = pts[i + 1]!;
    d += `Q${x},${y} ${(x + nx) / 2},${(y + ny) / 2}`;
  }
  const last = pts[pts.length - 1]!;
  return `${d}L${last[0]},${last[1]}`;
};

export const InkLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  lastVp.current = vp;
  const ink = store.sheet.ink ?? [];
  const drawing = ui.tool !== 'select';
  const [live, setLive] = useState<InkStroke | null>(null);
  const [erasing, setErasing] = useState<Set<string>>(new Set());
  const active = useRef<{ id: number; type: string; pressures: number[]; pan?: { x: number; y: number } } | null>(null);
  const host = useRef<SVGSVGElement>(null);
  if (!ink.length && !drawing) return null;

  const local = (e: React.PointerEvent): [number, number] => {
    const r = host.current!.getBoundingClientRect();
    const [x, y] = toSheet(vp, e.clientX - r.left, e.clientY - r.top);
    return [Math.round(x * 10) / 10, Math.round(y * 10) / 10];
  };

  const down = (e: React.PointerEvent) => {
    if (!drawing || e.button > 0) return;
    if (e.pointerType === 'pen') ui.penSeen = true;
    // Palm rejection: once a pen has been used, a touch pans the sheet instead of drawing.
    if (e.pointerType === 'touch' && (ui.penSeen || active.current?.type === 'pen')) {
      active.current = { id: e.pointerId, type: 'touch', pressures: [], pan: { x: e.clientX, y: e.clientY } };
      (e.currentTarget as Element).setPointerCapture(e.pointerId);
      return;
    }
    if (active.current) return;
    e.preventDefault();
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
    active.current = { id: e.pointerId, type: e.pointerType, pressures: [] };
    const p = local(e);
    if (ui.tool === 'eraser') { setErasing(new Set(strokesAt(ink, p, 6 / vp.zoom).map((k) => k.id))); return; }
    const tool = ui.tool === 'highlighter' ? 'highlighter' : ui.tool === 'marker' ? 'marker' : 'pen';
    const width = tool === 'highlighter' ? Math.max(10, ui.penWidth * 5) : tool === 'marker' ? ui.penWidth * 2.2 : ui.penWidth;
    if (e.pointerType === 'pen' && e.pressure > 0) active.current.pressures.push(e.pressure);
    setLive({ id: newId('ink'), tool, color: tool === 'highlighter' ? ui.hlColor : ui.penColor, width, points: [p] });
  };

  const move = (e: React.PointerEvent) => {
    const a = active.current;
    if (!a || a.id !== e.pointerId) return;
    if (a.pan) {
      const sc = document.querySelector('.gridwrap .scroller') as HTMLElement | null;
      if (sc) { sc.scrollLeft -= e.clientX - a.pan.x; sc.scrollTop -= e.clientY - a.pan.y; }
      a.pan = { x: e.clientX, y: e.clientY };
      return;
    }
    const events = (e.nativeEvent as PointerEvent).getCoalescedEvents?.() ?? [];
    const list = events.length ? events : [e.nativeEvent];
    if (ui.tool === 'eraser') {
      const hit = new Set(erasing);
      const r = host.current!.getBoundingClientRect();
      for (const ev of list) { const [x, y] = toSheet(vp, ev.clientX - r.left, ev.clientY - r.top); for (const k of strokesAt(ink, [x, y], 6 / vp.zoom)) hit.add(k.id); }
      if (hit.size !== erasing.size) setErasing(hit);
      return;
    }
    if (!live) return;
    const r = host.current!.getBoundingClientRect();
    const pts = [...live.points];
    for (const ev of list) {
      const [x, y] = toSheet(vp, ev.clientX - r.left, ev.clientY - r.top);
      const last = pts[pts.length - 1]!;
      if (Math.hypot(x - last[0], y - last[1]) < 0.8 / vp.zoom) continue;
      pts.push([Math.round(x * 10) / 10, Math.round(y * 10) / 10]);
      if (ev.pointerType === 'pen' && ev.pressure > 0) a.pressures.push(ev.pressure);
    }
    setLive({ ...live, points: pts });
  };

  const up = (e: React.PointerEvent) => {
    const a = active.current;
    if (!a || a.id !== e.pointerId) return;
    active.current = null;
    if (a.pan) return;
    if (ui.tool === 'eraser') { const ids = erasing; setErasing(new Set()); eraseStrokes(store, ids); return; }
    if (live) {
      let width = live.width;
      // Pen pressure scales the stroke (half pressure keeps the chosen width).
      if (a.pressures.length && live.tool !== 'highlighter') width = Math.max(0.5, Math.round(width * (0.5 + a.pressures.reduce((x, y) => x + y, 0) / a.pressures.length) * 10) / 10);
      addStroke(store, { ...live, width });
      setLive(null);
    }
  };

  const strokes = live ? [...ink, live] : ink;
  return (
    <svg ref={host} className={`inklayer${drawing ? ' on' : ''}`} width={vp.width} height={vp.height}
      style={{ position: 'absolute', left: 0, top: 0, pointerEvents: drawing ? 'auto' : 'none', touchAction: 'none', cursor: drawing ? (ui.tool === 'eraser' ? 'cell' : 'crosshair') : undefined }}
      onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={() => { active.current = null; setLive(null); setErasing(new Set()); }}
      onWheel={(e) => { const sc = document.querySelector('.gridwrap .scroller') as HTMLElement | null; if (sc && !e.ctrlKey) { sc.scrollLeft += e.deltaX; sc.scrollTop += e.deltaY; } }}>
      <defs><clipPath id="inkclip"><rect x={vp.headerW} y={vp.headerH} width={Math.max(0, vp.width - vp.headerW)} height={Math.max(0, vp.height - vp.headerH)} /></clipPath></defs>
      <g clipPath="url(#inkclip)">
        {strokes.map((k) => {
          const st = TOOL_STYLE[k.tool];
          const pts = k.points.map(([x, y]) => toScreen(vp, x, y));
          return <path key={k.id} className={k.tool === 'highlighter' ? 'hl' : undefined} d={pathOf(pts)} fill="none" stroke={k.color.toLowerCase() === DEFAULT_PEN ? 'var(--cell-ink)' : k.color}
            strokeWidth={k.width * vp.zoom} strokeOpacity={erasing.has(k.id) ? 0.15 : st.opacity} strokeLinecap={st.cap} strokeLinejoin="round" />;
        })}
      </g>
    </svg>
  );
};
