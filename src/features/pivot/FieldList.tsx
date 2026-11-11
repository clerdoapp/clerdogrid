// The PivotTable field list: fields to check or drag into Filters, Columns, Rows and Values.
import { useState, type DragEvent, type ReactNode } from 'react';
import { ChevronDown, Filter, GripVertical, RefreshCw, Rows3, Columns3, Sigma, X } from 'lucide-react';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import type { Viewport } from '../../ui/render';
import { VALUES_FIELD, type PivotTable } from '../../model/pivot';
import { cacheOf, refreshPivots, redrawPivot, type PivotRef } from './apply';
import { dataFieldName, fieldCaption } from './compute';
import { areaOf, moveDataField, placeField, removeField, type Area } from './model';
import { addField, toggleField, update } from './ops';
import { activePivot, currentPivot, pui, setPui } from './ui';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;
type Chip = { field: number; area: Area; index: number };
const MIME = 'application/x-clerdo-pivot-field';

const AREAS: { area: Area; label: string; icon: ReactNode }[] = [
  { area: 'pages', label: N('Filters'), icon: <Filter size={14} strokeWidth={S} /> },
  { area: 'cols', label: N('Columns'), icon: <Columns3 size={14} strokeWidth={S} /> },
  { area: 'rows', label: N('Rows'), icon: <Rows3 size={14} strokeWidth={S} /> },
  { area: 'data', label: N('Values'), icon: <Sigma size={14} strokeWidth={S} /> },
];

const chipsOf = (pt: PivotTable, area: Area): number[] =>
  area === 'rows' ? pt.rows : area === 'cols' ? pt.cols : area === 'pages' ? pt.pages.map((p) => p.field) : pt.data.map((d) => d.field);

/** Moves a field (or data field) from one place to another. */
const moveChip = (store: GridStore, ref: PivotRef, from: Chip | { field: number; area: null; index: -1 }, to: Area, at: number): void => {
  const cache = cacheOf(store.wb, ref.pt.cacheId);
  if (!cache) return;
  if (from.field === VALUES_FIELD && to !== 'rows' && to !== 'cols') return;
  update(store, ref, (pt) => {
    if (from.area === 'data' && to === 'data') return moveDataField(pt, from.index, at > from.index ? at - 1 : at);
    let p = pt;
    if (from.area === 'data') {
      const df = pt.data[from.index];
      p = removeField(pt, from.field, 'data', from.index);
      if (to === 'data' || !df) return p;
    } else if (from.area && to !== 'data') {
      const list = chipsOf(pt, from.area);
      if (from.area === to && list.indexOf(from.field) < at) at--;
    }
    return placeField(p, cache, from.field, to, at);
  }, N('Move field'));
};

export const FieldList = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const [menu, setMenu] = useState<null | { chip: Chip | null; field: number; x: number; y: number }>(null);
  const [drop, setDrop] = useState<{ area: Area; at: number } | null>(null);
  const seen = activePivot(store);
  const seenKey = seen ? `${seen.sheet}!${seen.pt.name}` : '';
  if (seenKey !== pui.lastSeen) {
    pui.lastSeen = seenKey;
    if (seen && !pui.pane && pui.autoPane) queueMicrotask(() => setPui(store, { pane: true, paneTarget: { sheet: seen.sheet, name: seen.pt.name } }));
  }
  if (!pui.pane) return null;
  const ref = currentPivot(store);
  const width = Math.min(300, Math.max(220, vp.width - vp.headerW - 40));
  const style = { left: vp.rtl ? 0 : Math.max(vp.headerW, vp.width - width), top: vp.headerH, height: Math.max(0, vp.height - vp.headerH), width, pointerEvents: 'auto' as const };
  const closePane = () => setPui(store, { pane: false, autoPane: false });
  if (!ref) {
    return (
      <aside className="pvpane" aria-label={T('PivotTable fields')} style={style} onPointerDown={(e) => e.stopPropagation()}>
        <div className="phead"><h2>{T('PivotTable fields')}</h2><button className="ico" aria-label={T('Close field list')} onClick={closePane}><X size={16} strokeWidth={S} /></button></div>
        <p className="muted small pad">{T('Select a cell in a PivotTable to see its fields.')}</p>
      </aside>
    );
  }
  if (pui.paneTarget?.name !== ref.pt.name || pui.paneTarget.sheet !== ref.sheet) queueMicrotask(() => { pui.paneTarget = { sheet: ref.sheet, name: ref.pt.name }; });
  const pt = ref.pt;
  const cache = cacheOf(store.wb, pt.cacheId);
  if (!cache) return null;
  const q = pui.search.trim().toLowerCase();
  const fields = cache.fields.map((f, i) => ({ f, i })).filter(({ f }) => !q || f.name.toLowerCase().includes(q));

  const startDrag = (e: DragEvent, chip: Chip | { field: number; area: null; index: -1 }) => {
    e.dataTransfer.setData(MIME, JSON.stringify(chip));
    e.dataTransfer.effectAllowed = 'move';
  };
  const overArea = (e: DragEvent, area: Area) => {
    if (![...e.dataTransfer.types].includes(MIME)) return;
    e.preventDefault();
    const list = (e.currentTarget as HTMLElement).querySelectorAll('.pvchip');
    let at = list.length;
    for (let i = 0; i < list.length; i++) { const r = list[i]!.getBoundingClientRect(); if (e.clientY < r.top + r.height / 2) { at = i; break; } }
    if (drop?.area !== area || drop.at !== at) setDrop({ area, at });
  };
  const dropOn = (e: DragEvent, area: Area) => {
    e.preventDefault();
    const raw = e.dataTransfer.getData(MIME);
    setDrop(null);
    if (!raw) return;
    try { moveChip(store, ref, JSON.parse(raw), area, drop?.area === area ? drop.at : chipsOf(pt, area).length); } catch { /* ignore bad drops */ }
  };
  const removeOutside = (e: DragEvent) => {
    const raw = e.dataTransfer.getData(MIME);
    if (!raw) return;
    try { const c = JSON.parse(raw) as Chip; if (c.area) update(store, ref, (p) => removeField(p, c.field, c.area, c.area === 'data' ? c.index : undefined), N('Remove field')); } catch { /* ignore */ }
  };

  const chipLabel = (c: Chip) => (c.area === 'data' ? dataFieldName(cache, pt.data[c.index]!) : fieldCaption(pt, cache, c.field));
  const menuItems = (m: NonNullable<typeof menu>): { label: string; run: () => void; disabled?: boolean }[] => {
    const c = m.chip;
    if (!c) {
      return (['pages', 'cols', 'rows', 'data'] as Area[]).map((a) => ({ label: ({ pages: N('Add to filters'), cols: N('Add to columns'), rows: N('Add to rows'), data: N('Add to values') } as Record<Area, string>)[a]!, run: () => addField(store, ref, m.field, a), disabled: a !== 'data' && !!cache.fields[m.field]!.formula }))
        .concat(cache.fields[m.field]!.formula ? [{ label: N('Edit calculated field…'), run: () => openDialog(store, 'pivotCalcField', { field: m.field }), disabled: false }] : []);
    }
    const list = chipsOf(pt, c.area);
    const out: { label: string; run: () => void; disabled?: boolean }[] = [
      { label: N('Move up'), run: () => moveChip(store, ref, c, c.area, c.index - 1), disabled: c.index === 0 },
      { label: N('Move down'), run: () => moveChip(store, ref, c, c.area, c.index + 2), disabled: c.index >= list.length - 1 },
      { label: N('Move to beginning'), run: () => moveChip(store, ref, c, c.area, 0), disabled: c.index === 0 },
      { label: N('Move to end'), run: () => moveChip(store, ref, c, c.area, list.length), disabled: c.index >= list.length - 1 },
    ];
    for (const a of AREAS) if (a.area !== c.area && (c.field !== VALUES_FIELD || a.area === 'rows' || a.area === 'cols')) out.push({ label: ({ pages: N('Move to filters'), cols: N('Move to columns'), rows: N('Move to rows'), data: N('Move to values') } as Record<Area, string>)[a.area]!, run: () => moveChip(store, ref, c, a.area, chipsOf(pt, a.area).length) });
    if (c.field !== VALUES_FIELD) {
      out.push({ label: N('Remove field'), run: () => update(store, ref, (p) => removeField(p, c.field, c.area, c.area === 'data' ? c.index : undefined), N('Remove field')) });
      out.push(c.area === 'data'
        ? { label: N('Value field settings…'), run: () => openDialog(store, 'pivotValueField', { sheet: ref.sheet, name: pt.name, index: c.index }) }
        : { label: N('Field settings…'), run: () => openDialog(store, 'pivotFieldSettings', { sheet: ref.sheet, name: pt.name, field: c.field }) });
    }
    return out;
  };

  return (
    <aside className="pvpane" aria-label={T('PivotTable fields')} style={style} onPointerDown={(e) => { e.stopPropagation(); if (menu && !(e.target as Element).closest('.pvmenu')) setMenu(null); }} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Escape') setMenu(null); }}
      onDragOver={(e) => { if ([...e.dataTransfer.types].includes(MIME)) e.preventDefault(); }}>
      <div className="phead">
        <h2>{T('PivotTable fields')}</h2>
        <button className="ico" aria-label={T('Close field list')} onClick={closePane}><X size={16} strokeWidth={S} /></button>
      </div>
      <div className="pvsearch"><input type="search" placeholder={T('Search fields')} aria-label={T('Search fields')} value={pui.search} onChange={(e) => setPui(store, { search: e.target.value })} /></div>
      <div className="pvfields" role="list" onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); removeOutside(e); }}>
        {fields.map(({ f, i }) => {
          const area = areaOf(pt, i);
          return (
            <div key={i} className="pvfield" role="listitem" draggable onDragStart={(e) => startDrag(e, { field: i, area: null, index: -1 })}>
              <label><input type="checkbox" checked={!!area} onChange={(e) => toggleField(store, ref, i, e.target.checked)} /><span title={f.formula ? `= ${f.formula}` : f.name}>{f.name}{f.formula ? <em className="muted">{' '}{T('(calculated)')}</em> : null}</span></label>
              <button className="ico sm" aria-label={T('More for {name}', { name: f.name })} onClick={(e) => { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setMenu({ chip: null, field: i, x: r.left, y: r.bottom }); }}><ChevronDown size={14} strokeWidth={S} /></button>
            </div>
          );
        })}
        {!fields.length && <p className="muted small pad">{T('No fields match.')}</p>}
      </div>
      <p className="pvhint muted small">{T('Drag fields between the areas below.')}</p>
      <div className="pvareas">
        {AREAS.map(({ area, label, icon }) => {
          const list = chipsOf(pt, area);
          return (
            <div key={area} className={`pvarea${drop?.area === area ? ' over' : ''}`} onDragOver={(e) => overArea(e, area)} onDragLeave={() => setDrop(null)} onDrop={(e) => dropOn(e, area)} aria-label={Tr(label)}>
              <div className="pvarea-h">{icon}<span>{Tr(label)}</span></div>
              <div className="pvarea-list">
                {list.map((field, index) => {
                  const chip: Chip = { field, area, index };
                  return (
                    <div key={`${field}:${index}`} className={`pvchip${drop?.area === area && drop.at === index ? ' before' : ''}`} draggable onDragStart={(e) => startDrag(e, chip)}>
                      <GripVertical size={12} strokeWidth={S} className="grip" />
                      <span className="name">{field === VALUES_FIELD ? T('Σ Values') : chipLabel(chip)}</span>
                      <button className="ico sm" aria-label={(field === VALUES_FIELD ? T('Options for values') : T('Options for {name}', { name: chipLabel(chip) }))} onClick={(e) => { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setMenu({ chip, field, x: r.left, y: r.bottom }); }}><ChevronDown size={14} strokeWidth={S} /></button>
                    </div>
                  );
                })}
                {drop?.area === area && drop.at >= list.length && <div className="pvdrop" />}
              </div>
            </div>
          );
        })}
      </div>
      <div className="pvfoot">
        <label className="b-check"><input type="checkbox" checked={!!pt.deferLayout} onChange={(e) => { const on = e.target.checked; if (on) update(store, ref, (p) => ({ ...p, deferLayout: true }), N('Defer layout update')); else redrawPivot(store, ref.sheet, pt.name, { deferLayout: undefined }); }} />{T('Defer layout update')}</label>
        <div className="grow" />
        {pt.deferLayout && <button className="sm primary" onClick={() => redrawPivot(store, ref.sheet, pt.name)}>{T('Update')}</button>}
        <button className="ico" aria-label={T('Refresh')} title={T('Refresh')} onClick={() => { const err = refreshPivots(store, [pt.cacheId]); if (err) store.notify(err); }}><RefreshCw size={15} strokeWidth={S} /></button>
      </div>
      {menu && (
        <div className="pvmenu menu" role="menu" style={{ position: 'fixed', left: Math.min(menu.x, window.innerWidth - 230), top: Math.min(menu.y, window.innerHeight - 320) }}>
          {menuItems(menu).map((it) => <button key={it.label} role="menuitem" disabled={it.disabled} onClick={() => { setMenu(null); it.run(); }}>{Tr(it.label)}</button>)}
        </div>
      )}
    </aside>
  );
};
