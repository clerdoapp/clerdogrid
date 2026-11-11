// Buttons drawn over PivotTable cells: +/- to expand and collapse, and field buttons that open the
// filter and sort popup.
import { useMemo, useState, type ReactElement } from 'react';
import type React from 'react';
import { ChevronDown, Filter, Minus, Plus } from 'lucide-react';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import { anchorBox, paneFor, panesOf, rectIn, screenToSheet, type Viewport } from '../../ui/render';
import { cacheOf, drawnArea, findByKey, layoutOf, pivotAt, type PivotRef } from './apply';
import { normalizeFields, dataFieldName, fieldCaption } from './compute';
import { clearFieldFilters, moveItem, setHiddenItems, setPageItem, sortField, toggleItem } from './ops';
import { pui, setPui } from './ui';
import { T } from '../../ui/i18n';

const S = 2;

export const CellButtons = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const sheet = store.sheet;
  if (!sheet.pivots?.length) return pui.popup ? <FilterPopup store={store} vp={vp} /> : null;
  const panes = panesOf(vp);
  const out: ReactElement[] = [];
  sheet.pivots.forEach((pt, index) => {
    const area = drawnArea(pt);
    const l = layoutOf(store.wb, pt);
    if (!area || !l) return;
    const ref: PivotRef = { sheet: store.sheetIndex, index, pt };
    for (const p of panes) {
      const r1 = Math.max(p.r1, area.r1), r2 = Math.min(p.r2, area.r2), c1 = Math.max(p.c1, area.c1), c2 = Math.min(p.c2, area.c2);
      if (r1 > r2 || c1 > c2) continue;
      for (let row = r1; row <= r2; row++) {
        if (!vp.layout.rows.size(row)) continue;
        for (let col = c1; col <= c2; col++) {
          if (!vp.layout.cols.size(col)) continue;
          const dr = row - pt.anchor.row, dc = col - pt.anchor.col;
          const cell = dr < 0 ? l.pageCells.find((x) => x.dr === dr && x.dc === dc)?.cell : l.cells[dr]?.[dc];
          if (!cell) continue;
          const pane = paneFor(panes, vp, row, col);
          if (pane !== p) continue;
          const rect = rectIn(p, vp, { r1: row, c1: col, r2: row, c2: col });
          const z = vp.zoom;
          if (cell.button && pt.showDrill && cell.field !== undefined && cell.slot !== undefined && cell.field >= 0) {
            const size = Math.max(10, Math.round(11 * z));
            const indent = cell.role === 'rowItem' ? (cell.indent ?? 0) * 3 * 9 * z : 0;
            const field = cell.field, slot = cell.slot;
            out.push(
              <button key={`b${index}:${row}:${col}`} className="pvbtn" aria-label={cell.button === 'collapsed' ? T('Expand') : T('Collapse')} title={cell.button === 'collapsed' ? T('Expand') : T('Collapse')}
                style={{ left: vp.rtl ? rect.x + rect.w - 4 * z - indent - size : rect.x + 4 * z + indent, top: rect.y + (rect.h - size) / 2, width: size, height: size }}
                onPointerDown={(e) => e.stopPropagation()} onClick={() => toggleItem(store, ref, field, slot)}>
                {cell.button === 'collapsed' ? <Plus size={size - 2} strokeWidth={S} /> : <Minus size={size - 2} strokeWidth={S} />}
              </button>,
            );
          }
          if (cell.filter !== undefined && pt.showHeaders) {
            const size = Math.max(14, Math.round(16 * z));
            const filtered = filteredField(ref, cell.filter);
            const f = cell.filter;
            out.push(
              <button key={`f${index}:${row}:${col}`} className={`pvfbtn${filtered ? ' on' : ''}`} aria-label={T('Filter and sort')} title={T('Filter and sort')}
                style={{ left: vp.rtl ? rect.x + 2 : rect.x + rect.w - size - 2, top: rect.y + (rect.h - size) / 2, width: size, height: size }}
                onPointerDown={(e) => e.stopPropagation()}
                onClick={() => setPui(store, { popup: { sheet: store.sheetIndex, name: pt.name, field: f, x: vp.rtl ? rect.x + 120 : rect.x + rect.w, y: rect.y + rect.h, page: cell.role === 'pageLabel' || cell.role === 'pageValue' } })}>
                {filtered ? <Filter size={size - 5} strokeWidth={S} /> : <ChevronDown size={size - 3} strokeWidth={S} />}
              </button>,
            );
          }
        }
      }
    }
  });
  return <>{out}<ItemDrag store={store} vp={vp} />{pui.popup && <FilterPopup store={store} vp={vp} />}</>;
};

const filteredField = (ref: PivotRef, f: number | 'rows' | 'cols'): boolean => {
  const fields = typeof f === 'number' ? [f] : (f === 'rows' ? ref.pt.rows : ref.pt.cols).filter((x) => x >= 0);
  return fields.some((x) => ref.pt.fields[x]?.items.some((it) => it.hidden) || !!ref.pt.fields[x]?.filters?.length || ref.pt.pages.some((p) => p.field === x && p.item !== undefined));
};

/** Sort, filter and item checklist for one field. */
const FilterPopup = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const p = pui.popup!;
  const sh = store.wb.sheets[p.sheet];
  const index = sh?.pivots?.findIndex((x) => x.name === p.name) ?? -1;
  const pt = sh?.pivots?.[index];
  const cache = pt ? cacheOf(store.wb, pt.cacheId) : undefined;
  const choices = pt ? (typeof p.field === 'number' ? [p.field] : (p.field === 'rows' ? pt.rows : pt.cols).filter((x) => x >= 0)) : [];
  const [field, setField] = useState<number>(choices[0] ?? 0);
  const norm = useMemo(() => (pt && cache ? normalizeFields(pt, cache) : []), [pt, cache]);
  const l = pt ? layoutOf(store.wb, pt) : null;
  const pf = norm[field];
  const page = pt?.pages.find((x) => x.field === field);
  const [multi, setMulti] = useState(!!pf?.multiPage || (!!page && page.item === undefined && !!pf?.items.some((it) => it.hidden)));
  const [hidden, setHidden] = useState<Set<number>>(() => new Set(pf?.items.flatMap((it, i) => (it.hidden ? [i] : [])) ?? []));
  const [pageItem, setPageItemState] = useState<number | undefined>(page?.item);
  const [q, setQ] = useState('');
  const [dragOver, setDragOver] = useState<number | null>(null);
  if (!pt || !cache || !pf || !l) return null;
  const ref: PivotRef = { sheet: p.sheet, index, pt };
  const close = () => setPui(store, { popup: null });
  const items = pf.items.map((_, i) => ({ i, text: l.caption(field, i) })).filter((x) => !q || x.text.toLowerCase().includes(q.toLowerCase()));
  const allOn = items.every((x) => !hidden.has(x.i));
  const isPage = !!page;
  const left = vp.rtl ? Math.max(0, Math.min(p.x - 120, vp.width - vp.headerW - 260)) : Math.max(vp.headerW, Math.min(p.x - 120, vp.width - 260));
  const top = Math.min(p.y, Math.max(vp.headerH, vp.height - 440));
  const apply = () => {
    if (isPage) setPageItem(store, ref, field, multi ? undefined : pageItem, multi ? hidden : undefined);
    else setHiddenItems(store, ref, field, hidden);
    close();
  };
  return (
    <div className="pvpop" role="dialog" aria-label={T('Filter and sort')} style={{ left, top }} onPointerDown={(e) => e.stopPropagation()} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Escape') close(); if (e.key === 'Enter') apply(); }}>
      {choices.length > 1 && (
        <label className="field"><span>{T('Field')}</span>
          <select value={field} onChange={(e) => { const f = +e.target.value; setField(f); setHidden(new Set(norm[f]!.items.flatMap((it, i) => (it.hidden ? [i] : [])))); }}>
            {choices.map((f) => <option key={f} value={f}>{fieldCaption(pt, cache, f)}</option>)}
          </select>
        </label>
      )}
      {!isPage && (
        <div className="pvpop-sort">
          <button onClick={() => { sortField(store, ref, field, 'ascending'); close(); }}>{T('Sort A to Z')}</button>
          <button onClick={() => { sortField(store, ref, field, 'descending'); close(); }}>{T('Sort Z to A')}</button>
          {pt.data.length > 0 && (
            <label className="field"><span>{T('Sort by value')}</span>
              <select value={pf.sortByData ?? ''} onChange={(e) => { const v = e.target.value; sortField(store, ref, field, v === '' ? undefined : (pf.sort ?? 'descending'), v === '' ? undefined : +v); close(); }}>
                <option value="">{T('Labels')}</option>
                {pt.data.map((d, i) => <option key={i} value={i}>{dataFieldName(cache, d)}</option>)}
              </select>
            </label>
          )}
          <div className="pvpop-row">
            <button onClick={() => { close(); openDialog(store, 'pivotFilter', { sheet: ref.sheet, name: pt.name, field, kind: 'label' }); }}>{T('Label filter…')}</button>
            <button disabled={!pt.data.length} onClick={() => { close(); openDialog(store, 'pivotFilter', { sheet: ref.sheet, name: pt.name, field, kind: 'value' }); }}>{T('Value filter…')}</button>
          </div>
          <button disabled={!filteredField(ref, field)} onClick={() => { clearFieldFilters(store, ref, field); close(); }}>{T('Clear filter from “{name}”', { name: fieldCaption(pt, cache, field) })}</button>
        </div>
      )}
      <input type="search" placeholder={T('Search')} aria-label={T('Search items')} value={q} onChange={(e) => setQ(e.target.value)} />
      {isPage && <label className="b-check"><input type="checkbox" checked={multi} onChange={(e) => setMulti(e.target.checked)} />{T('Select several items')}</label>}
      <div className="pvpop-list">
        {isPage && !multi ? (
          <>
            <label className="b-check"><input type="radio" name="pvpage" checked={pageItem === undefined} onChange={() => setPageItemState(undefined)} />{T('(All)')}</label>
            {items.map((x) => <label key={x.i} className="b-check"><input type="radio" name="pvpage" checked={pageItem === x.i} onChange={() => setPageItemState(x.i)} />{x.text}</label>)}
          </>
        ) : (
          <>
            <label className="b-check"><input type="checkbox" checked={allOn} onChange={(e) => { const n = new Set(hidden); for (const x of items) { if (e.target.checked) n.delete(x.i); else n.add(x.i); } setHidden(n); }} />{T('(Select all)')}</label>
            {items.map((x) => (
              <label key={x.i} className={`b-check pvdrag${dragOver === x.i ? ' over' : ''}`} draggable={!q} title={T('Drag to change the order')}
                onDragStart={(e) => { e.dataTransfer.setData('text/x-pivot-item', String(x.i)); e.dataTransfer.effectAllowed = 'move'; }}
                onDragOver={(e) => { if ([...e.dataTransfer.types].includes('text/x-pivot-item')) { e.preventDefault(); setDragOver(x.i); } }}
                onDragLeave={() => setDragOver(null)}
                onDrop={(e) => {
                  e.preventDefault(); setDragOver(null);
                  const from = Number(e.dataTransfer.getData('text/x-pivot-item'));
                  if (!Number.isInteger(from) || from === x.i) return;
                  const to = from < x.i ? x.i - 1 : x.i;
                  // Keep the checkboxes on the same items after the move.
                  const order = pf.items.map((_, i) => i);
                  const [moved] = order.splice(from, 1);
                  order.splice(to, 0, moved!);
                  setHidden(new Set(order.flatMap((old, i) => (hidden.has(old) ? [i] : []))));
                  moveItem(store, ref, field, from, to);
                }}>
                <input type="checkbox" checked={!hidden.has(x.i)} onChange={(e) => { const n = new Set(hidden); if (e.target.checked) n.delete(x.i); else n.add(x.i); setHidden(n); }} />{x.text}
              </label>
            ))}
          </>
        )}
      </div>
      <div className="pvpop-foot"><button onClick={close}>{T('Cancel')}</button><button className="primary" disabled={!isPage && pf.items.every((_, i) => hidden.has(i))} onClick={apply}>{T('OK')}</button></div>
    </div>
  );
};

/** Field buttons on a PivotChart: one per row and column field, opening the same filter popup. */
export const PivotChartButtons = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const charts = store.sheet.charts.filter((c) => c.pivot);
  if (!charts.length) return null;
  return (
    <>
      {charts.map((ch) => {
        const hit = findByKey(store.wb, ch.pivot!);
        if (!hit) return null;
        const cache = cacheOf(store.wb, hit.pt.cacheId);
        if (!cache) return null;
        const a = ch.anchor;
        const { left, top } = anchorBox(vp, a);
        const fields = [...hit.pt.pages.map((p) => p.field), ...hit.pt.cols, ...hit.pt.rows].filter((f) => f >= 0);
        return (
          <div key={ch.id} className="pvchartbtns" style={{ left, top: Math.max(0, top - 24), width: a.width * vp.zoom }}>
            {hit.pt.data.map((d, i) => <span key={`d${i}`} className="pvchartbtn data">{dataFieldName(cache, d)}</span>)}
            {fields.map((f) => (
              <button key={f} className={`pvchartbtn${filteredField(hit, f) ? ' on' : ''}`} onPointerDown={(e) => e.stopPropagation()}
                onClick={(e) => { const b = (e.currentTarget as HTMLElement).getBoundingClientRect(); const host = (e.currentTarget.closest('.objlayer') as HTMLElement).getBoundingClientRect(); setPui(store, { popup: { sheet: hit.sheet, name: hit.pt.name, field: f, x: b.right - host.left + vp.headerW, y: b.bottom - host.top + vp.headerH, page: hit.pt.pages.some((p) => p.field === f) } }); }}>
                {fieldCaption(hit.pt, cache, f)}<ChevronDown size={12} strokeWidth={S} />
              </button>
            ))}
          </div>
        );
      })}
    </>
  );
};

/**
 * Dragging an item label by its border changes the manual order: the active label cell gets a
 * grab frame, and a line shows where the item will land.
 */
const ItemDrag = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const [drag, setDrag] = useState<null | { x: number; y: number; target: { slot: number; after: boolean; line: { x: number; y: number; w: number; h: number } } | null }>(null);
  const { row, col } = store.sel.active;
  const ref = pivotAt(store.wb, store.sheetIndex, row, col);
  const l = ref ? layoutOf(store.wb, ref.pt) : null;
  if (!ref || !l || store.editing) return null;
  const pt = ref.pt;
  const cell = l.cells[row - pt.anchor.row]?.[col - pt.anchor.col];
  if (!cell || (cell.role !== 'rowItem' && cell.role !== 'colItem') || cell.total || cell.field === undefined || cell.field < 0 || cell.slot === undefined) return null;
  const isRow = cell.role === 'rowItem';
  const field = cell.field, from = cell.slot;
  const panes = panesOf(vp);
  const box = rectIn(paneFor(panes, vp, row, col), vp, { r1: row, c1: col, r2: row, c2: col });
  const gridPoint = (e: { clientX: number; clientY: number; currentTarget: EventTarget }) => {
    const host = (e.currentTarget as HTMLElement).closest('.objlayer')?.parentElement?.getBoundingClientRect();
    return { x: e.clientX - (host?.left ?? 0), y: e.clientY - (host?.top ?? 0) };
  };
  const targetAt = (x: number, y: number) => {
    const [sx, sy] = screenToSheet(vp, x, y);
    const r = vp.layout.rows.at(Math.max(0, sy)), c = vp.layout.cols.at(Math.max(0, sx));
    const idx = isRow ? r - pt.anchor.row - l.firstDataRow : c - pt.anchor.col - l.firstDataCol;
    const line = (isRow ? l.rowLines : l.colLines)[idx];
    const el2 = line && line.kind !== 'grand' ? line.path.find((p) => p.field === field) : undefined;
    if (!el2) return null;
    const rr = rectIn(paneFor(panes, vp, isRow ? r : row, isRow ? col : c), vp, { r1: isRow ? r : row, c1: isRow ? pt.anchor.col : c, r2: isRow ? r : row, c2: isRow ? pt.anchor.col + Math.max(0, l.firstDataCol - 1) : c });
    const after = isRow ? y > rr.y + rr.h / 2 : vp.rtl ? x < rr.x + rr.w / 2 : x > rr.x + rr.w / 2;
    const lineBox = isRow ? { x: rr.x, y: (after ? rr.y + rr.h : rr.y) - 1.5, w: rr.w, h: 3 } : { x: (after !== !!vp.rtl ? rr.x + rr.w : rr.x) - 1.5, y: rr.y, w: 3, h: rr.h };
    return { slot: el2.slot, after, line: lineBox };
  };
  const t = 5;
  const edge = (style: React.CSSProperties, k: string) => (
    <div key={k} className="pvgrab" style={{ position: 'absolute', ...style }} title={T('Drag to move this item')}
      onPointerDown={(e) => { e.stopPropagation(); (e.currentTarget as Element).setPointerCapture?.(e.pointerId); const p = gridPoint(e); setDrag({ ...p, target: null }); }}
      onPointerMove={(e) => { if (!drag) return; const p = gridPoint(e); setDrag({ ...p, target: targetAt(p.x, p.y) }); }}
      onPointerUp={() => {
        const d = drag;
        setDrag(null);
        const tg = d?.target;
        if (!tg || tg.slot === from) return;
        const to = tg.after ? (tg.slot < from ? tg.slot + 1 : tg.slot) : (tg.slot < from ? tg.slot : tg.slot - 1);
        moveItem(store, ref, field, from, to);
      }} />
  );
  return (
    <>
      {edge({ left: box.x - 2, top: box.y - 2, width: box.w + 4, height: t }, 't')}
      {edge({ left: box.x - 2, top: box.y + box.h - t + 2, width: box.w + 4, height: t }, 'b')}
      {edge({ left: box.x - 2, top: box.y - 2, width: t, height: box.h + 4 }, 'l')}
      {edge({ left: box.x + box.w - t + 2, top: box.y - 2, width: t, height: box.h + 4 }, 'r')}
      {drag?.target && <div className="pvdropline" style={{ left: drag.target.line.x, top: drag.target.line.y, width: drag.target.line.w, height: drag.target.line.h }} />}
    </>
  );
};
