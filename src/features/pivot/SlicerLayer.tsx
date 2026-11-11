// Slicers and timelines floating over the sheet.
import { useRef, useState, type PointerEvent as RPointerEvent } from 'react';
import { ListChecks, FilterX } from 'lucide-react';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import { fromSerial, toSerial } from '../../engine/date';
import { anchorBox, logicalDx, type Viewport } from '../../ui/render';
import type { Slicer, SlicerCache } from '../../model/pivot';
import { cacheByName, setSlicerHidden, setTimelineRange, slicerItems, timelineBounds, updateSlicer } from './slicers';
import { pui, setPui } from './ui';
import { N, T, Tr, getLang } from '../../ui/i18n';

const S = 1.9;
/** Short month name in the interface language. */
const month = (m: number): string => new Intl.DateTimeFormat(getLang(), { month: 'short', timeZone: 'UTC' }).format(Date.UTC(2000, m - 1, 1));

const place = (vp: Viewport, a: Slicer['anchor']) => anchorBox(vp, a);

/** Anchor for a box moved or resized by some screen pixels. */
const moved = (vp: Viewport, a: Slicer['anchor'], dx: number, dy: number, dw: number, dh: number): Slicer['anchor'] => {
  const x = Math.max(0, vp.layout.cols.start(a.col) + a.dx + logicalDx(vp, dx, dw) / vp.zoom), y = Math.max(0, vp.layout.rows.start(a.row) + a.dy + dy / vp.zoom);
  const col = vp.layout.cols.at(x), row = vp.layout.rows.at(y);
  return { col, row, dx: Math.round(x - vp.layout.cols.start(col)), dy: Math.round(y - vp.layout.rows.start(row)), width: Math.max(100, Math.round(a.width + dw / vp.zoom)), height: Math.max(60, Math.round(a.height + dh / vp.zoom)) };
};

export const SlicerLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const list = store.sheet.slicers;
  if (!list?.length) return null;
  return <>{list.map((s) => <SlicerBox key={s.name} store={store} vp={vp} s={s} />)}</>;
};

const theme = (style: string): string => (/dark/i.test(style) ? 'dark' : /other|medium/i.test(style) ? 'mid' : /Light1$/.test(style) ? 'plain' : 'light');

const SlicerBox = ({ store, vp, s }: { store: GridStore; vp: Viewport; s: Slicer }) => {
  const [drag, setDrag] = useState<{ kind: 'move' | 'size'; x: number; y: number; dx: number; dy: number } | null>(null);
  const [multi, setMulti] = useState(false);
  const sc = cacheByName(store, s.cache);
  if (!sc) return null;
  const box = place(vp, s.anchor);
  const selected = pui.slicer === s.name;
  const shown = drag ? { left: box.left + (drag.kind === 'move' ? drag.dx : 0), top: box.top + (drag.kind === 'move' ? drag.dy : 0), width: Math.max(100, box.width + (drag.kind === 'size' ? drag.dx : 0)), height: Math.max(60, box.height + (drag.kind === 'size' ? drag.dy : 0)) } : box;
  const startDrag = (kind: 'move' | 'size') => (e: RPointerEvent) => {
    if (e.button !== 0) return;
    e.stopPropagation();
    (e.currentTarget as Element).setPointerCapture?.(e.pointerId);
    setPui(store, { slicer: s.name });
    setDrag({ kind, x: e.clientX, y: e.clientY, dx: 0, dy: 0 });
  };
  const onMove = (e: RPointerEvent) => { if (drag) setDrag({ ...drag, dx: e.clientX - drag.x, dy: e.clientY - drag.y }); };
  const onUp = () => {
    if (!drag) return;
    const d = drag;
    setDrag(null);
    if (Math.abs(d.dx) + Math.abs(d.dy) < 2) return;
    const anchor = d.kind === 'move' ? moved(vp, s.anchor, d.dx, d.dy, 0, 0) : moved(vp, s.anchor, 0, 0, d.dx, d.dy);
    updateSlicer(store, s.name, { anchor }, undefined, d.kind === 'move' ? N('Move slicer') : N('Resize slicer'));
  };
  return (
    <div className={`pvslicer t-${theme(s.style)}${selected ? ' sel' : ''}${s.kind === 'timeline' ? ' timeline' : ''}`} style={shown} role="group" aria-label={s.kind === 'timeline' ? T('Timeline {name}', { name: s.caption }) : T('Slicer {name}', { name: s.caption })}
      onPointerDown={(e) => { e.stopPropagation(); if (pui.slicer !== s.name) setPui(store, { slicer: s.name }); }} onPointerMove={onMove} onPointerUp={onUp}
      onDoubleClick={(e) => { e.stopPropagation(); openDialog(store, 'slicerSettings', { name: s.name }); }}>
      {s.showCaption && (
        <div className="pvslicer-h" onPointerDown={startDrag('move')}>
          <span className="cap">{s.caption}</span>
          {s.kind === 'slicer' && <button className={`ico sm${multi ? ' on' : ''}`} aria-pressed={multi} title={T('Select several')} aria-label={T('Select several')} onPointerDown={(e) => e.stopPropagation()} onClick={() => setMulti(!multi)}><ListChecks size={14} strokeWidth={S} /></button>}
          <button className="ico sm" title={T('Clear filter')} aria-label={T('Clear filter')} disabled={s.kind === 'slicer' ? !sc.hidden.length : sc.timeline?.start === undefined} onPointerDown={(e) => e.stopPropagation()}
            onClick={() => (s.kind === 'slicer' ? setSlicerHidden(store, sc.name, [], N('Clear slicer')) : setTimelineRange(store, sc.name, undefined, undefined))}><FilterX size={14} strokeWidth={S} /></button>
        </div>
      )}
      {!s.showCaption && <div className="pvslicer-grip" onPointerDown={startDrag('move')} />}
      {s.kind === 'slicer' ? <SlicerButtons store={store} s={s} sc={sc} multi={multi} zoom={vp.zoom} /> : <TimelineStrip store={store} s={s} sc={sc} />}
      <div className="pvslicer-size" onPointerDown={startDrag('size')} aria-hidden="true" />
    </div>
  );
};

const SlicerButtons = ({ store, s, sc, multi, zoom }: { store: GridStore; s: Slicer; sc: SlicerCache; multi: boolean; zoom: number }) => {
  const items = slicerItems(store, sc);
  const anchorKey = useRef<string | null>(null);
  const click = (key: string, e: { ctrlKey: boolean; metaKey: boolean; shiftKey: boolean }) => {
    const all = items.map((i) => i.key);
    const hidden = new Set(sc.hidden);
    if (e.shiftKey && anchorKey.current) {
      const a = all.indexOf(anchorKey.current), b = all.indexOf(key);
      const [lo, hi] = [Math.min(a, b), Math.max(a, b)];
      const keep = new Set(all.slice(lo, hi + 1));
      setSlicerHidden(store, sc.name, [...new Set([...sc.hidden.filter((k) => !all.includes(k)), ...all.filter((k) => !keep.has(k))])]);
      return;
    }
    anchorKey.current = key;
    if (multi || e.ctrlKey || e.metaKey) {
      if (!sc.hidden.length) { setSlicerHidden(store, sc.name, all.filter((k) => k !== key)); return; }
      if (hidden.has(key)) hidden.delete(key); else hidden.add(key);
      if (all.every((k) => hidden.has(k))) return;
      setSlicerHidden(store, sc.name, [...hidden]);
      return;
    }
    const onlyThis = sc.hidden.length === all.length - 1 && !hidden.has(key);
    setSlicerHidden(store, sc.name, onlyThis ? [] : all.filter((k) => k !== key));
  };
  return (
    <div className="pvslicer-items" style={{ gridTemplateColumns: `repeat(${Math.max(1, s.columns)}, minmax(0, 1fr))` }} role="listbox" aria-multiselectable="true">
      {items.map((it) => (
        <button key={it.key} role="option" aria-selected={it.selected && !!sc.hidden.length} className={`pvsitem${it.selected ? ' on' : ''}${it.hasData ? '' : ' nodata'}${sc.hidden.length ? ' filtered' : ''}`}
          style={{ height: Math.max(18, s.buttonHeight * zoom) }} onPointerDown={(e) => e.stopPropagation()} onClick={(e) => click(it.key, e)} title={it.label}>{it.label}</button>
      ))}
      {!items.length && <span className="muted small">{T('No items')}</span>}
    </div>
  );
};

type Level = 'years' | 'quarters' | 'months' | 'days';
const periodStart = (serial: number, level: Level, d1904: boolean): number => {
  const d = fromSerial(serial, d1904);
  if (level === 'years') return toSerial(d.year, 1, 1, d1904);
  if (level === 'quarters') return toSerial(d.year, Math.floor((d.month - 1) / 3) * 3 + 1, 1, d1904);
  if (level === 'months') return toSerial(d.year, d.month, 1, d1904);
  return serial;
};
const nextPeriod = (serial: number, level: Level, d1904: boolean): number => {
  const d = fromSerial(serial, d1904);
  if (level === 'years') return toSerial(d.year + 1, 1, 1, d1904);
  if (level === 'quarters') return toSerial(d.year, d.month + 3, 1, d1904);
  if (level === 'months') return toSerial(d.year, d.month + 1, 1, d1904);
  return serial + 1;
};
const periodLabel = (serial: number, level: Level, d1904: boolean): { top: string; label: string } => {
  const d = fromSerial(serial, d1904);
  if (level === 'years') return { top: '', label: String(d.year) };
  if (level === 'quarters') return { top: String(d.year), label: `Q${Math.floor((d.month - 1) / 3) + 1}` };
  if (level === 'months') return { top: String(d.year), label: month(d.month) };
  return { top: `${month(d.month)} ${d.year}`, label: String(d.day) };
};

const TimelineStrip = ({ store, s, sc }: { store: GridStore; s: Slicer; sc: SlicerCache }) => {
  const [dragFrom, setDragFrom] = useState<number | null>(null);
  const [hover, setHover] = useState<number | null>(null);
  const bounds = timelineBounds(store, sc);
  const d1904 = store.wb.date1904;
  const level: Level = s.level ?? sc.timeline?.level ?? 'months';
  if (!bounds) return <div className="muted small pad">{T('No dates')}</div>;
  const periods: number[] = [];
  for (let p = periodStart(bounds[0], level, d1904); p <= bounds[1] && periods.length < 2000; p = nextPeriod(p, level, d1904)) periods.push(p);
  const tl = sc.timeline ?? { level };
  const inSel = (p: number) => tl.start !== undefined && tl.end !== undefined && p <= tl.end && nextPeriod(p, level, d1904) - 1 >= tl.start;
  const commit = (a: number, b: number) => {
    const lo = Math.min(a, b), hi = Math.max(a, b);
    setTimelineRange(store, sc.name, periods[lo]!, nextPeriod(periods[hi]!, level, d1904) - 1);
  };
  const fmt = (v: number) => { const d = fromSerial(v, d1904); return new Intl.DateTimeFormat(getLang(), { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' }).format(Date.UTC(d.year, d.month - 1, d.day)); };
  const label = tl.start !== undefined && tl.end !== undefined ? `${fmt(tl.start)} – ${fmt(tl.end)}` : T('All periods');
  return (
    <div className="pvtl" onPointerDown={(e) => e.stopPropagation()}>
      <div className="pvtl-bar">
        {s.showSelectionLabel !== false && <span className="pvtl-label">{Tr(label)}</span>}
        {s.showTimeLevel !== false && (
          <select aria-label={T('Time level')} value={level} onChange={(e) => updateSlicer(store, s.name, { level: e.target.value as Level }, { timeline: { ...tl, level: e.target.value as Level } }, N('Timeline level'))}>
            <option value="years">{T('Years')}</option><option value="quarters">{T('Quarters')}</option><option value="months">{T('Months')}</option><option value="days">{T('Days')}</option>
          </select>
        )}
      </div>
      <div className="pvtl-strip" onPointerLeave={() => setHover(null)}>
        {periods.map((p, i) => {
          const l = periodLabel(p, level, d1904);
          const showTop = i === 0 || periodLabel(periods[i - 1]!, level, d1904).top !== l.top;
          const on = dragFrom !== null && hover !== null ? i >= Math.min(dragFrom, hover) && i <= Math.max(dragFrom, hover) : inSel(p);
          return (
            <div key={p} className="pvtl-cell">
              <span className="pvtl-top">{showTop ? l.top : ''}</span>
              <button className={`pvtl-p${on ? ' on' : ''}`} aria-label={`${l.top} ${l.label}`.trim()} aria-pressed={on}
                onPointerDown={(e) => { (e.currentTarget as Element).releasePointerCapture?.(e.pointerId); if (e.shiftKey && tl.start !== undefined) { const first = periods.findIndex((q) => inSel(q)); commit(first < 0 ? i : first, i); return; } setDragFrom(i); setHover(i); }}
                onPointerEnter={() => { if (dragFrom !== null) setHover(i); }}
                onPointerUp={() => { if (dragFrom !== null) { commit(dragFrom, i); setDragFrom(null); setHover(null); } }}>
                <span>{l.label}</span>
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
};
