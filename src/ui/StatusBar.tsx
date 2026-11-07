import { useMemo } from 'react';
import { Minus, Plus } from 'lucide-react';
import { formatValue } from '../formats/numfmt';
import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { run, setZoom } from '../app/commands';
import type { GridStore } from '../app/store';

const S = 1.9;

/** Sum, average and count of the selection, like the bottom bar of other spreadsheets. */
const useStats = (store: GridStore) => useMemo(() => {
  let sum = 0, count = 0, nums = 0, min = Infinity, max = -Infinity, cells = 0;
  const b = store.engine.usedBounds(store.sheetIndex);
  const seen = new Set<number>();
  for (const r of store.sel.ranges) {
    const r2 = Math.min(r.r2, b.rows), c2 = Math.min(r.c2, b.cols);
    for (let row = r.r1; row <= r2; row++) {
      if (store.sheet.rows.get(row)?.hidden) continue;
      for (let col = r.c1; col <= c2; col++) {
        const k = row * MAX_COLS + col;
        if (seen.has(k)) continue;
        seen.add(k);
        if (++cells > 2_000_000) return null;
        const v = store.value(row, col);
        if (v === null) continue;
        count++;
        if (typeof v === 'number') { nums++; sum += v; min = Math.min(min, v); max = Math.max(max, v); }
      }
    }
  }
  return { sum, count, nums, min, max };
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [store.getVersion(), store.layoutVersion]);

export const StatusBar = ({ store }: { store: GridStore }) => {
  const st = useStats(store);
  const r = store.range;
  const multi = store.sel.ranges.length > 1 || r.r1 !== r.r2 || r.c1 !== r.c2;
  const fmt = store.style(store.sel.active.row, store.sel.active.col).numFmt ?? 'General';
  const show = (n: number) => formatValue(+n.toPrecision(15), fmt === '@' ? 'General' : fmt).text;
  const mode = store.editing ? (store.editing.replace ? 'Enter' : 'Edit') : store.marquee ? (store.marquee.cut ? 'Select destination and press Enter or Paste' : 'Select destination and press Enter or Paste') : 'Ready';
  const zoom = store.sheet.view.zoom;
  const circ = store.engine.circular[0];
  return (
    <footer className="status">
      <span>{mode}</span>
      {circ && <button className="pill warn" onClick={() => { store.setActiveSheet(circ.sheet); store.select({ row: circ.row, col: circ.col }); }}>Circular reference</button>}
      {store.sheet.autoFilter?.columns.length ? <span className="pill">Filter mode</span> : null}
      <div className="grow" />
      {multi && st && st.count > 0 && (
        <span className="stats">
          {st.nums > 0 && <span>Average: {show(st.sum / st.nums)}</span>}
          <span>Count: {st.count}</span>
          {st.nums > 0 && <><span>Min: {show(st.min)}</span><span>Max: {show(st.max)}</span><span>Sum: {show(st.sum)}</span></>}
        </span>
      )}
      {multi && r.r1 === 0 && r.r2 === MAX_ROWS - 1 ? null : null}
      <button aria-label="Zoom out" onClick={() => run(store, 'view.zoomOut')}><Minus size={14} strokeWidth={S} /></button>
      <input type="range" min={10} max={400} value={zoom} aria-label="Zoom" onChange={(e) => setZoom(store, +e.target.value)} />
      <button aria-label="Zoom in" onClick={() => run(store, 'view.zoomIn')}><Plus size={14} strokeWidth={S} /></button>
      <button className="zoomv" onClick={() => setZoom(store, 100)} title="Reset zoom">{zoom}%</button>
    </footer>
  );
};
