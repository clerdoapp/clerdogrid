import { useEffect, useRef, useState, type PointerEvent as RPointerEvent } from 'react';
import { Paintbrush, Pencil, Trash2 } from 'lucide-react';
import { openDialog } from '../../app/commands';
import { layoutOf } from '../../app/layout';
import type { GridStore } from '../../app/store';
import type { Chart } from '../../model/workbook';
import { resolveChart } from './chartData';
import { ChartSvg } from './ChartSvg';
import { logicalDx } from '../render';
import './charts.css';
import './palette';
import { N, T } from '../i18n';

type Box = { left: number; top: number; width: number; height: number };
type Corner = 'nw' | 'ne' | 'sw' | 'se';

/** Which chart is selected. Only one at a time, and only on the active sheet. */
let selectedId: string | null = null;
export const selectedChart = (): string | null => selectedId;
export const selectChart = (store: GridStore, id: string | null): void => {
  if (selectedId === id) return;
  selectedId = id;
  store.emit();
};

const MIN_W = 120, MIN_H = 80;

/** Shows the chart format panel. */
export const openFormat = (store: GridStore): void => { store.view.panel = 'chart'; store.emit(); };

export const deleteChart = (store: GridStore, id: string): void => {
  store.change(N('Delete chart'), 'sheet', () => { store.sheet.charts = store.sheet.charts.filter((c) => c.id !== id); });
  selectChart(store, null);
};

/** Writes a new on-screen box back into the chart's cell anchor (unzoomed sheet pixels). */
const commitBox = (store: GridStore, chart: Chart, delta: Box, zoom: number, label: string): void => {
  const layout = layoutOf(store.sheet, store.layoutVersion);
  const a = chart.anchor;
  const x = Math.max(0, layout.cols.start(a.col) + a.dx + logicalDx(store.sheet.view, delta.left, delta.width) / zoom);
  const y = Math.max(0, layout.rows.start(a.row) + a.dy + delta.top / zoom);
  const col = layout.cols.at(x), row = layout.rows.at(y);
  const next = {
    col, row,
    dx: Math.round(x - layout.cols.start(col)), dy: Math.round(y - layout.rows.start(row)),
    width: Math.max(MIN_W, Math.round(a.width + delta.width / zoom)), height: Math.max(MIN_H, Math.round(a.height + delta.height / zoom)),
  };
  if (next.col === a.col && next.row === a.row && next.dx === a.dx && next.dy === a.dy && next.width === a.width && next.height === a.height) return;
  store.change(label, 'sheet', () => {
    const c = store.sheet.charts.find((x) => x.id === chart.id);
    if (c) c.anchor = next;
  });
};

export const ChartView = ({ store, chart, box, zoom }: { store: GridStore; chart: Chart; box: Box; zoom: number }) => {
  const el = useRef<HTMLDivElement>(null);
  const drag = useRef<{ kind: 'move' | Corner; x: number; y: number; id: number; moved: boolean } | null>(null);
  const [delta, setDelta] = useState<Box>({ left: 0, top: 0, width: 0, height: 0 });
  const selected = selectedId === chart.id;

  // Clicking anywhere outside the chart deselects it.
  useEffect(() => {
    if (!selected) return;
    // The tool panel and dialogs work on the selected chart, so clicks there keep it.
    const off = (e: PointerEvent) => {
      const t = e.target as Element;
      if (el.current && !el.current.contains(t) && !t.closest?.('.toolpanel, .backdrop, .rail')) selectChart(store, null);
    };
    document.addEventListener('pointerdown', off, true);
    return () => document.removeEventListener('pointerdown', off, true);
  }, [selected, store]);

  const start = (kind: 'move' | Corner) => (e: RPointerEvent) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    selectChart(store, chart.id);
    el.current?.focus({ preventScroll: true });
    drag.current = { kind, x: e.clientX, y: e.clientY, id: e.pointerId, moved: false };
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
  };

  const move = (e: RPointerEvent) => {
    const d = drag.current;
    if (!d || d.id !== e.pointerId) return;
    const dx = e.clientX - d.x, dy = e.clientY - d.y;
    if (!d.moved && Math.abs(dx) + Math.abs(dy) < 3) return;
    d.moved = true;
    const minW = MIN_W * zoom - box.width, minH = MIN_H * zoom - box.height;
    if (d.kind === 'move') { setDelta({ left: dx, top: dy, width: 0, height: 0 }); return; }
    const west = d.kind === 'nw' || d.kind === 'sw', north = d.kind === 'nw' || d.kind === 'ne';
    const w = Math.max(minW, west ? -dx : dx), h = Math.max(minH, north ? -dy : dy);
    setDelta({ left: west ? -w : 0, top: north ? -h : 0, width: w, height: h });
  };

  const end = (e: RPointerEvent) => {
    const d = drag.current;
    if (!d || d.id !== e.pointerId) return;
    drag.current = null;
    if (d.moved) commitBox(store, chart, delta, zoom, d.kind === 'move' ? N('Move chart') : N('Resize chart'));
    setDelta({ left: 0, top: 0, width: 0, height: 0 });
  };

  const data = resolveChart(store, chart);
  const style = { left: box.left + delta.left, top: box.top + delta.top, width: box.width + delta.width, height: box.height + delta.height };
  const handlers = { onPointerMove: move, onPointerUp: end, onPointerCancel: end };

  return (
    <div ref={el} className={`chartbox${selected ? ' on' : ''}`} style={style} tabIndex={-1} role="figure" aria-label={data.title || T('Chart')} data-chart-id={chart.id}
      onPointerDown={start('move')} {...handlers}
      onDoubleClick={(e) => { e.stopPropagation(); openDialog(store, 'chart', { id: chart.id }); }}
      onContextMenu={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        e.stopPropagation();
        if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteChart(store, chart.id); }
        if (e.key === 'Escape') selectChart(store, null);
        if (e.key === 'Enter') openDialog(store, 'chart', { id: chart.id });
      }}>
      <ChartSvg chart={chart} data={data} width={chart.anchor.width + delta.width / zoom} height={chart.anchor.height + delta.height / zoom} />
      {selected && (
        <>
          {(['nw', 'ne', 'sw', 'se'] as const).map((k) => <div key={k} className={`chandle ${k}`} onPointerDown={start(k)} {...handlers} />)}
          <div className="chartedit" onPointerDown={(e) => e.stopPropagation()}>
            <button className="ico sm" aria-label={T('Format chart')} title={T('Format chart')} onClick={() => openFormat(store)}><Paintbrush size={15} strokeWidth={1.9} /></button>
            <button className="ico sm" aria-label={T('Edit chart')} title={T('Edit chart')} onClick={() => openDialog(store, 'chart', { id: chart.id })}><Pencil size={15} strokeWidth={1.9} /></button>
            <button className="ico sm" aria-label={T('Delete chart')} title={T('Delete chart')} onClick={() => deleteChart(store, chart.id)}><Trash2 size={15} strokeWidth={1.9} /></button>
          </div>
        </>
      )}
    </div>
  );
};
