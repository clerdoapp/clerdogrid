import { useMemo } from 'react';
import { Minus, Plus } from 'lucide-react';
import { formatValue } from '../formats/numfmt';
import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { run, setZoom } from '../app/commands';
import type { GridStore } from '../app/store';
import { localizeNumberText } from '../app/locale';
import { N, T } from './i18n';

const S = 1.9;

export type StatusItem = 'average' | 'count' | 'numCount' | 'min' | 'max' | 'sum';
export const STATUS_ITEMS: [StatusItem, string][] = [['average', N('Average')], ['count', N('Count')], ['numCount', N('Numerical count')], ['min', N('Minimum')], ['max', N('Maximum')], ['sum', N('Sum')]];
/** Which selection figures the status bar shows (File > Options or right-click the status bar). */
export const statusPrefs: { items: StatusItem[]; onMenu: ((store: GridStore, x: number, y: number) => void) | null } = { items: ['average', 'count', 'sum'], onMenu: null };

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
  const show = (n: number) => localizeNumberText(formatValue(+n.toPrecision(15), fmt === '@' ? 'General' : fmt).text, fmt);
  const mode = store.editing ? (store.editing.replace ? T('Enter') : T('Edit')) : store.marquee ? T('Select destination and press Enter or Paste') : T('Ready');
  const zoom = store.sheet.view.zoom;
  const circ = store.engine.circular[0];
  return (
    <footer className="status" onContextMenu={(e) => { if (statusPrefs.onMenu) { e.preventDefault(); statusPrefs.onMenu(store, e.clientX, e.clientY); } }}>
      <span>{mode}</span>
      {circ && <button className="pill warn" onClick={() => { store.setActiveSheet(circ.sheet); store.select({ row: circ.row, col: circ.col }); }}>{T('Circular reference')}</button>}
      {store.sheet.autoFilter?.columns.length ? <span className="pill">{T('Filter mode')}</span> : null}
      <div className="grow" />
      {store.engine.stale && <button className="pill warn" title={T('Calculate now (F9)')} onClick={() => run(store, 'formula.calcNow')}>{T('Calculate')}</button>}
      {multi && st && st.count > 0 && (
        <span className="stats">
          {statusPrefs.items.includes('average') && st.nums > 0 && <span>{T('Average:')}{' '}{show(st.sum / st.nums)}</span>}
          {statusPrefs.items.includes('count') && <span>{T('Count:')}{' '}{st.count}</span>}
          {statusPrefs.items.includes('numCount') && <span>{T('Numerical count:')}{' '}{st.nums}</span>}
          {statusPrefs.items.includes('min') && st.nums > 0 && <span>{T('Min:')}{' '}{show(st.min)}</span>}
          {statusPrefs.items.includes('max') && st.nums > 0 && <span>{T('Max:')}{' '}{show(st.max)}</span>}
          {statusPrefs.items.includes('sum') && st.nums > 0 && <span>{T('Sum:')}{' '}{show(st.sum)}</span>}
        </span>
      )}
      {multi && r.r1 === 0 && r.r2 === MAX_ROWS - 1 ? null : null}
      <button aria-label={T('Zoom out')} onClick={() => run(store, 'view.zoomOut')}><Minus size={14} strokeWidth={S} /></button>
      <input type="range" min={10} max={400} value={zoom} aria-label={T('Zoom')} onChange={(e) => setZoom(store, +e.target.value)} />
      <button aria-label={T('Zoom in')} onClick={() => run(store, 'view.zoomIn')}><Plus size={14} strokeWidth={S} /></button>
      <button className="zoomv" onClick={() => setZoom(store, 100)} title={T('Reset zoom')}>{zoom}%</button>
    </footer>
  );
};
