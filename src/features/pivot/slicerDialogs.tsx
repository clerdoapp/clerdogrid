import { useState } from 'react';
import type { GridStore } from '../../app/store';
import { inRange } from '../../app/store';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check } from '../../ui/dialogs/shared';
import { cacheOf } from './apply';
import { activePivot, pui } from './ui';
import {
  SLICER_STYLES, TIMELINE_STYLES, cacheByName, connectable, insertPivotSlicers, insertTableSlicers, insertTimeline, setConnections, updateSlicer, applySlicerCache,
} from './slicers';
import { N, T, Tr } from '../../ui/i18n';

const tableAt = (store: GridStore) => store.sheet.tables.find((t) => inRange(t.range, store.sel.active.row, store.sel.active.col));

export const InsertSlicerDialog = ({ store, close }: DialogProps) => {
  const ref = activePivot(store);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const table = !ref ? tableAt(store) : undefined;
  const [picked, setPicked] = useState<Set<number>>(new Set());
  if (!ref && !table) return <Dialog title={T('Insert slicer')} close={close}><p>{T('Select a cell in a PivotTable or a table first.')}</p></Dialog>;
  const options = cache ? cache.fields.map((f, i) => ({ i, name: f.name, ok: !f.formula })) : table!.columns.map((c, i) => ({ i, name: c.name, ok: true }));
  const ok = () => {
    if (!picked.size) { store.notify(T('Pick at least one field.')); return false; }
    if (ref) insertPivotSlicers(store, ref.sheet, ref.pt.name, [...picked]);
    else insertTableSlicers(store, store.sheetIndex, table!.name, [...picked]);
    return true;
  };
  return (
    <Dialog title={T('Insert slicer')} close={close} width={340} onOk={ok} okLabel={T('Insert')}>
      <p className="muted small">{ref ? T('One slicer per field. Click its buttons to filter the PivotTable.') : T('One slicer per field. Click its buttons to filter the table.')}</p>
      <div className="checks pvlist">
        {options.filter((o) => o.ok).map((o) => <label key={o.i}><input type="checkbox" checked={picked.has(o.i)} onChange={(e) => { const n = new Set(picked); if (e.target.checked) n.add(o.i); else n.delete(o.i); setPicked(n); }} /> {o.name}</label>)}
      </div>
    </Dialog>
  );
};

export const InsertTimelineDialog = ({ store, close }: DialogProps) => {
  const ref = activePivot(store);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const dates = cache ? cache.fields.map((f, i) => ({ f, i })).filter(({ f }) => f.dates && f.database !== false) : [];
  const [pick, setPick] = useState(dates[0]?.i ?? -1);
  if (!ref || !cache) return <Dialog title={T('Insert timeline')} close={close}><p>{T('Select a cell in a PivotTable first.')}</p></Dialog>;
  if (!dates.length) return <Dialog title={T('Insert timeline')} close={close}><p>{T('This PivotTable\'s data has no date fields.')}</p></Dialog>;
  return (
    <Dialog title={T('Insert timeline')} close={close} width={340} onOk={() => { insertTimeline(store, ref.sheet, ref.pt.name, pick); return true; }} okLabel={T('Insert')}>
      <div className="listbox">{dates.map(({ f, i }) => <label key={i}><input type="radio" name="pvtl" checked={pick === i} onChange={() => setPick(i)} /> {f.name}</label>)}</div>
    </Dialog>
  );
};

export const SlicerSettingsDialog = ({ store, close, data }: DialogProps<{ name?: string }>) => {
  const name = data?.name ?? pui.slicer ?? '';
  const s = store.wb.sheets.flatMap((x) => x.slicers ?? []).find((x) => x.name === name);
  const sc = s ? cacheByName(store, s.cache) : undefined;
  const [caption, setCaption] = useState(s?.caption ?? '');
  const [newName, setNewName] = useState(s?.name ?? '');
  const [showCaption, setShowCaption] = useState(s?.showCaption ?? true);
  const [columns, setColumns] = useState(String(s?.columns ?? 1));
  const [height, setHeight] = useState(String(s?.buttonHeight ?? 26));
  const [style, setStyle] = useState(s?.style ?? '');
  const [sort, setSort] = useState(sc?.sort ?? 'ascending');
  const [hideNoData, setHideNoData] = useState(!!sc?.hideNoData);
  const [noDataLast, setNoDataLast] = useState(sc?.showNoDataLast !== false);
  const [showLabel, setShowLabel] = useState(s?.showSelectionLabel !== false);
  const [showLevel, setShowLevel] = useState(s?.showTimeLevel !== false);
  const [conns, setConns] = useState(sc?.source.kind === 'pivot' ? sc.source.pivots : []);
  if (!s || !sc) return <Dialog title={T('Slicer settings')} close={close}><p>{T('Select a slicer or timeline first.')}</p></Dialog>;
  const timeline = s.kind === 'timeline';
  const all = sc.source.kind === 'pivot' ? connectable(store, sc.source.cacheId) : [];
  const ok = () => {
    const n = newName.trim();
    if (!n) { store.notify(T('Enter a name.')); return false; }
    if (n !== s.name && store.wb.sheets.some((x) => x.slicers?.some((y) => y.name.toLowerCase() === n.toLowerCase()))) { store.notify(T('Another slicer has that name.')); return false; }
    const cols = Math.max(1, Math.min(20000, Math.round(+columns || 1)));
    const h = Math.max(16, Math.min(200, Math.round(+height || 26)));
    updateSlicer(store, s.name, { name: n, caption: caption.trim() || s.caption, showCaption, columns: cols, buttonHeight: h, style, showSelectionLabel: showLabel ? undefined : false, showTimeLevel: showLevel ? undefined : false },
      { sort: sort as 'ascending', hideNoData: hideNoData || undefined, showNoDataLast: noDataLast ? undefined : false });
    if (sc.source.kind === 'pivot' && JSON.stringify(conns) !== JSON.stringify(sc.source.pivots)) {
      if (!conns.length) { store.notify(T('Keep at least one PivotTable connected.')); return false; }
      setConnections(store, sc.name, conns);
    } else store.change(N('Slicer settings'), 'book', () => { const c = cacheByName(store, sc.name); if (c) applySlicerCache(store, c); });
    return true;
  };
  return (
    <Dialog title={timeline ? T('Timeline settings') : T('Slicer settings')} close={close} width={440} onOk={ok}>
      <div className="grid2">
        <Field label={T('Caption')}><input value={caption} onChange={(e) => setCaption(e.target.value)} /></Field>
        <Field label={T('Name')}><input value={newName} onChange={(e) => setNewName(e.target.value)} /></Field>
      </div>
      <Check label={timeline ? T('Show header') : T('Show caption')} checked={showCaption} onChange={setShowCaption} />
      {!timeline ? (
        <>
          <div className="grid2">
            <Field label={T('Columns')}><input type="number" min={1} value={columns} onChange={(e) => setColumns(e.target.value)} /></Field>
            <Field label={T('Button height (px)')}><input type="number" min={16} value={height} onChange={(e) => setHeight(e.target.value)} /></Field>
          </div>
          <Field label={T('Sort items')}><select value={sort} onChange={(e) => setSort(e.target.value as 'ascending')}><option value="ascending">{T('A to Z')}</option><option value="descending">{T('Z to A')}</option><option value="none">{T('Data order')}</option></select></Field>
          <Check label={T('Hide items with no data')} checked={hideNoData} onChange={setHideNoData} />
          {!hideNoData && <Check label={T('Show items with no data last')} checked={noDataLast} onChange={setNoDataLast} />}
        </>
      ) : (
        <>
          <Check label={T('Show the selected dates')} checked={showLabel} onChange={setShowLabel} />
          <Check label={T('Show the time level')} checked={showLevel} onChange={setShowLevel} />
        </>
      )}
      <Field label={T('Style')}><select value={style} onChange={(e) => setStyle(e.target.value)}>{(timeline ? TIMELINE_STYLES : SLICER_STYLES).map((x) => <option key={x} value={x}>{styleLabel(x)}</option>)}</select></Field>
      {all.length > 1 && (
        <>
          <div className="seclabel">{T('PivotTables this filters')}</div>
          <div className="checks pvlist">
            {all.map((p) => {
              const on = conns.some((c) => c.sheet === p.sheet && c.name === p.name);
              return <label key={`${p.sheet}!${p.name}`}><input type="checkbox" checked={on} onChange={(e) => setConns(e.target.checked ? [...conns, p] : conns.filter((c) => !(c.sheet === p.sheet && c.name === p.name)))} /> {p.name} <span className="muted">({p.sheet})</span></label>;
            })}
          </div>
        </>
      )}
    </Dialog>
  );
};

const styleLabel = (x: string): string => {
  const base = /Dark/.test(x) ? 'Dark green' : /Other/.test(x) ? 'Medium green' : /Light1$/.test(x) ? 'Plain' : 'Light green';
  return /6$/.test(x) ? T('{style}, bold', { style: Tr(base) }) : Tr(base);
};
