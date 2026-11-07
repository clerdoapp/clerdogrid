import { useEffect, useRef, useState, type ReactElement } from 'react';
import { ArrowDownToLine, ArrowUpToLine, Pencil, Trash2, ScanText } from 'lucide-react';
import type { Shape, ShapeKind } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import type { Viewport } from '../../ui/render';
import { fontFamily } from '../../ui/render';
import { anchorAt, anchorPx, place, setUi, ui, type Box } from './state';
import { deleteShape, orderShape, updateShape } from './ops';
import { luminance } from './checks';

const S = 1.9;

/** Outline of each shape kind inside a w×h box (SVG path data). */
export const shapePath = (kind: ShapeKind, w: number, h: number): string => {
  const p = (pts: [number, number][]) => `M${pts.map(([x, y]) => `${x * w},${y * h}`).join('L')}Z`;
  switch (kind) {
    case 'ellipse': return `M0,${h / 2}A${w / 2},${h / 2} 0 1 0 ${w},${h / 2}A${w / 2},${h / 2} 0 1 0 0,${h / 2}Z`;
    case 'roundRect': { const r = Math.min(w, h) * 0.16; return `M${r},0H${w - r}Q${w},0 ${w},${r}V${h - r}Q${w},${h} ${w - r},${h}H${r}Q0,${h} 0,${h - r}V${r}Q0,0 ${r},0Z`; }
    case 'triangle': return p([[0.5, 0], [1, 1], [0, 1]]);
    case 'diamond': return p([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]);
    case 'pentagon': return p([[0.5, 0], [1, 0.38], [0.81, 1], [0.19, 1], [0, 0.38]]);
    case 'hexagon': return p([[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]]);
    case 'star5': {
      const pts: [number, number][] = [];
      for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + (i * Math.PI) / 5, r = i % 2 ? 0.19 : 0.5; pts.push([0.5 + r * Math.cos(a) * 1.0, 0.53 + r * Math.sin(a) * 1.05]); }
      return p(pts);
    }
    case 'rightArrow': return p([[0, 0.25], [0.62, 0.25], [0.62, 0], [1, 0.5], [0.62, 1], [0.62, 0.75], [0, 0.75]]);
    case 'leftArrow': return p([[1, 0.25], [0.38, 0.25], [0.38, 0], [0, 0.5], [0.38, 1], [0.38, 0.75], [1, 0.75]]);
    case 'upArrow': return p([[0.25, 1], [0.25, 0.38], [0, 0.38], [0.5, 0], [1, 0.38], [0.75, 0.38], [0.75, 1]]);
    case 'downArrow': return p([[0.25, 0], [0.25, 0.62], [0, 0.62], [0.5, 1], [1, 0.62], [0.75, 0.62], [0.75, 0]]);
    case 'callout': return p([[0, 0], [1, 0], [1, 0.75], [0.42, 0.75], [0.2, 1], [0.25, 0.75], [0, 0.75]]);
    case 'line': case 'arrowLine': return `M0,0L${w},${h}`;
    default: return `M0,0H${w}V${h}H0Z`;
  }
};

const isLine = (k: ShapeKind) => k === 'line' || k === 'arrowLine';
/** Automatic text color: readable on the shape's own fill, or the sheet's ink when it has none. */
const inkFor = (sh: Shape): string => (sh.fill ? (luminance(sh.fill) > 0.4 ? '#111111' : '#ffffff') : 'var(--cell-ink)');

export const ShapeSvg = ({ shape, w, h, zoom }: { shape: Shape; w: number; h: number; zoom: number }) => {
  const lw = (shape.line?.width ?? 0) * zoom;
  const dash = shape.line?.dash === 'dash' ? `${lw * 4} ${lw * 2}` : shape.line?.dash === 'dot' ? `${lw} ${lw * 1.6}` : undefined;
  const pad = Math.max(lw, isLine(shape.kind) ? 8 : 0);
  const id = `ah-${shape.id}`;
  return (
    <svg width={w + pad * 2} height={h + pad * 2} style={{ position: 'absolute', left: -pad, top: -pad, overflow: 'visible' }} aria-hidden="true">
      {shape.kind === 'arrowLine' && <defs><marker id={id} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0L10,5L0,10Z" fill={shape.line?.color ?? '#0C7A4B'} /></marker></defs>}
      <path transform={`translate(${pad},${pad})`} d={shapePath(shape.kind, w, h)} fill={isLine(shape.kind) ? 'none' : shape.fill ?? 'none'}
        stroke={lw > 0 ? shape.line!.color : 'none'} strokeWidth={lw} strokeDasharray={dash} strokeLinejoin="round" strokeLinecap="round"
        markerEnd={shape.kind === 'arrowLine' ? `url(#${id})` : undefined} />
    </svg>
  );
};

type Drag = { id: string; mode: 'move' | 'rotate' | string; sx: number; sy: number; orig: { x: number; y: number; w: number; h: number; rot: number }; cx: number; cy: number; moved: boolean };

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;

const focusGrid = () => (document.querySelector('.gridwrap') as HTMLElement | null)?.focus({ preventScroll: true });

export const ShapeLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const shapes = store.sheet.shapes ?? [];
  const [live, setLive] = useState<{ id: string; x: number; y: number; w: number; h: number; rot: number } | null>(null);
  const drag = useRef<Drag | null>(null);
  if (!shapes.length) return null;
  if (ui.shape && !shapes.some((s) => s.id === ui.shape)) ui.shape = null;

  const begin = (e: React.PointerEvent, sh: Shape, mode: Drag['mode']) => {
    if (e.button === 2) return;
    e.stopPropagation();
    if (ui.editingShape === sh.id && mode === 'move') return;
    e.preventDefault();
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
    if (ui.shape !== sh.id || ui.editingShape) setUi(store, { shape: sh.id, editingShape: null, pinned: null });
    focusGrid();
    const a = anchorPx(vp, sh.anchor);
    const b = place(vp, sh.anchor);
    drag.current = { id: sh.id, mode, sx: e.clientX, sy: e.clientY, orig: { x: a.x, y: a.y, w: sh.anchor.width, h: sh.anchor.height, rot: sh.rotation ?? 0 }, cx: 0, cy: 0, moved: false };
    const host = (e.currentTarget as Element).closest('.gridwrap')!.getBoundingClientRect();
    drag.current.cx = host.left + b.left + b.width / 2; drag.current.cy = host.top + b.top + b.height / 2;
  };

  const move = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d) return;
    const z = vp.zoom;
    let dx = (e.clientX - d.sx) / z, dy = (e.clientY - d.sy) / z;
    if (Math.abs(dx) + Math.abs(dy) > 1) d.moved = true;
    const o = d.orig;
    if (d.mode === 'move') { setLive({ id: d.id, x: o.x + dx, y: o.y + dy, w: o.w, h: o.h, rot: o.rot }); return; }
    if (d.mode === 'rotate') {
      let deg = (Math.atan2(e.clientY - d.cy, e.clientX - d.cx) * 180) / Math.PI + 90;
      if (e.shiftKey) deg = Math.round(deg / 15) * 15;
      setLive({ id: d.id, ...o, rot: Math.round(((deg % 360) + 360) % 360) });
      return;
    }
    // Resize in the shape's own (rotated) frame.
    const rad = (-o.rot * Math.PI) / 180;
    [dx, dy] = [dx * Math.cos(rad) - dy * Math.sin(rad), dx * Math.sin(rad) + dy * Math.cos(rad)];
    let { x, y, w, h } = o;
    const m = d.mode;
    if (m.includes('e')) w = o.w + dx;
    if (m.includes('s')) h = o.h + dy;
    if (m.includes('w')) { w = o.w - dx; x = o.x + dx; }
    if (m.startsWith('n')) { h = o.h - dy; y = o.y + dy; }
    if (e.shiftKey && o.w > 0 && o.h > 0 && m.length === 2) {
      const ratio = o.w / o.h;
      if (Math.abs(w / o.w) > Math.abs(h / o.h)) { const nh = w / ratio; if (m.startsWith('n')) y += h - nh; h = nh; }
      else { const nw = h * ratio; if (m.includes('w')) x += w - nw; w = nw; }
    }
    const minS = isLine(shapes.find((s) => s.id === d.id)!.kind) ? 0 : 8;
    if (w < minS) { if (m.includes('w')) x -= minS - w; w = minS; }
    if (h < minS) { if (m.startsWith('n')) y -= minS - h; h = minS; }
    setLive({ id: d.id, x, y, w, h, rot: o.rot });
  };

  const end = (e: React.PointerEvent) => {
    const d = drag.current;
    drag.current = null;
    const l = live;
    setLive(null);
    if (!d || !l || !d.moved) return;
    void e;
    const anchor = anchorAt(vp, l.x, l.y, l.w, l.h);
    updateShape(store, d.id, { anchor, rotation: l.rot || undefined }, d.mode === 'move' ? 'Move shape' : d.mode === 'rotate' ? 'Rotate shape' : 'Resize shape');
  };

  return (
    <>
      {shapes.map((sh) => {
        const lv = live?.id === sh.id ? live : null;
        let box: Box = place(vp, sh.anchor);
        if (lv) {
          const o = anchorPx(vp, sh.anchor);
          box = { left: box.left + (lv.x - o.x) * vp.zoom, top: box.top + (lv.y - o.y) * vp.zoom, width: lv.w * vp.zoom, height: lv.h * vp.zoom };
        }
        const rot = lv ? lv.rot : sh.rotation ?? 0;
        const sel = ui.shape === sh.id;
        const editing = ui.editingShape === sh.id;
        const f = sh.font ?? {};
        return (
          <div key={sh.id} className={`gshape${sel ? ' sel' : ''}`} role="img" aria-label={sh.alt || sh.text || sh.kind}
            style={{ left: box.left, top: box.top, width: box.width, height: box.height, transform: rot ? `rotate(${rot}deg)` : undefined, pointerEvents: 'auto' }}
            onPointerDown={(e) => begin(e, sh, 'move')} onPointerMove={move} onPointerUp={end} onPointerCancel={() => { drag.current = null; setLive(null); }}
            onDoubleClick={(e) => { e.stopPropagation(); if (!isLine(sh.kind)) setUi(store, { shape: sh.id, editingShape: sh.id }); }}
            onContextMenu={(e) => { e.preventDefault(); e.stopPropagation(); setUi(store, { shape: sh.id }); }}>
            <ShapeSvg shape={sh} w={box.width} h={box.height} zoom={vp.zoom} />
            {!isLine(sh.kind) && (editing
              ? <ShapeText store={store} shape={sh} zoom={vp.zoom} />
              : sh.text ? <div className="gshape-text" style={{ justifyContent: sh.kind === 'textBox' ? 'flex-start' : 'center', textAlign: sh.align ?? 'center', font: `${f.italic ? 'italic ' : ''}${f.bold ? 700 : 400} ${((f.size ?? 11) * 4) / 3 * vp.zoom}px ${fontFamily(f.name)}`, color: f.color ?? inkFor(sh) }}>{sh.text}</div> : null)}
            {sel && !editing && <>
              {HANDLES.map((hd) => <span key={hd} className={`ghandle h-${hd}`} onPointerDown={(e) => begin(e, sh, hd)} onPointerMove={move} onPointerUp={end} />)}
              {!isLine(sh.kind) && <span className="ghandle rot" title="Rotate" onPointerDown={(e) => begin(e, sh, 'rotate')} onPointerMove={move} onPointerUp={end} />}
            </>}
          </div>
        );
      })}
      {(() => {
        const sh = shapes.find((s) => s.id === ui.shape);
        if (!sh || live || ui.editingShape) return null;
        const p0 = place(vp, sh.anchor);
        // Clear the rotated outline: use its bounding box.
        const rad = ((sh.rotation ?? 0) * Math.PI) / 180;
        const bw = Math.abs(p0.width * Math.cos(rad)) + Math.abs(p0.height * Math.sin(rad)), bh = Math.abs(p0.width * Math.sin(rad)) + Math.abs(p0.height * Math.cos(rad));
        const b = { left: p0.left + (p0.width - bw) / 2, top: p0.top + (p0.height - bh) / 2, width: bw, height: bh };
        const top = b.top - 76 < vp.headerH ? b.top + b.height + 14 : b.top - 76;
        return (
          <div className="gshape-bar" role="toolbar" aria-label="Shape" style={{ left: Math.max(vp.headerW + 4, Math.min(b.left, vp.width - 230)), top, pointerEvents: 'auto' }} onPointerDown={(e) => e.stopPropagation()}>
            {!isLine(sh.kind) && <button title="Edit text" aria-label="Edit text" onClick={() => setUi(store, { editingShape: sh.id })}><Pencil size={15} strokeWidth={S} /></button>}
            <button title="Bring forward" aria-label="Bring forward" onClick={() => orderShape(store, sh.id, 'forward')}><ArrowUpToLine size={15} strokeWidth={S} /></button>
            <button title="Send backward" aria-label="Send backward" onClick={() => orderShape(store, sh.id, 'backward')}><ArrowDownToLine size={15} strokeWidth={S} /></button>
            <button title="Alt text" aria-label="Alt text" onClick={() => openDialog(store, 'altText', { type: 'shape', id: sh.id })}><ScanText size={15} strokeWidth={S} /></button>
            <button title="Format" aria-label="Format" className="txt" onClick={() => { store.view.panel = 'insert'; store.emit(); requestAnimationFrame(() => document.querySelector('.shapefmt')?.scrollIntoView({ block: 'nearest' })); }}>Format</button>
            <button title="Delete" aria-label="Delete shape" onClick={() => { deleteShape(store, sh.id); setUi(store, { shape: null }); focusGrid(); }}><Trash2 size={15} strokeWidth={S} /></button>
          </div>
        );
      })()}
    </>
  );
};

const ShapeText = ({ store, shape, zoom }: { store: GridStore; shape: Shape; zoom: number }) => {
  const [text, setText] = useState(shape.text ?? '');
  const ref = useRef<HTMLTextAreaElement>(null);
  useEffect(() => { const t = ref.current; if (t) { t.focus(); t.setSelectionRange(t.value.length, t.value.length); } }, []);
  const done = () => {
    if (ui.editingShape !== shape.id) return;
    ui.editingShape = null;
    if (text !== (shape.text ?? '')) updateShape(store, shape.id, { text: text || undefined }, 'Shape text'); else store.emit();
    focusGrid();
  };
  const f = shape.font ?? {};
  return (
    <textarea ref={ref} className="gshape-edit" value={text} aria-label="Shape text" onChange={(e) => setText(e.target.value)} onBlur={done}
      onPointerDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Escape' || (e.key === 'Enter' && (e.metaKey || e.ctrlKey))) { e.preventDefault(); done(); } }}
      style={{ textAlign: shape.align ?? 'center', font: `${f.italic ? 'italic ' : ''}${f.bold ? 700 : 400} ${((f.size ?? 11) * 4) / 3 * zoom}px ${fontFamily(f.name)}`, color: f.color ?? inkFor(shape) }} />
  );
};

export const shapeIcon = (kind: ShapeKind): ReactElement => (
  <svg viewBox="-2 -2 24 24" width="22" height="22" aria-hidden="true">
    {kind === 'textBox'
      ? <><rect x="1" y="3" width="18" height="14" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.5" strokeDasharray="2.5 2" /><path d="M6 7h8M10 7v7" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" /></>
      : <><path d={shapePath(kind, 20, isLine(kind) ? 20 : kind.endsWith('Arrow') || kind === 'rect' || kind === 'roundRect' || kind === 'callout' ? 14 : 20)} transform={isLine(kind) || !(kind.endsWith('Arrow') || kind === 'rect' || kind === 'roundRect' || kind === 'callout') ? undefined : 'translate(0,3)'}
        fill={isLine(kind) ? 'none' : 'var(--accent-soft)'} stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
        {kind === 'arrowLine' && <path d="M20 20l-6-1.5 4.5-4.5z" fill="currentColor" />}</>}
  </svg>
);
