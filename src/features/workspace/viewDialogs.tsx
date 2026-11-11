// Sheet views manager, custom views manager, and the status bar choices menu.
import { useEffect, useRef, useState } from 'react';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check } from '../../ui/dialogs/shared';
import { STATUS_ITEMS, type StatusItem } from '../../ui/StatusBar';
import { activeViewOf, addCustomView, createView, deleteCustomView, deleteView, duplicateView, keepTemporary, renameView, showCustomView, switchView } from './views';
import { saveSettings, settings } from './settings';
import { applyDeviceSettings } from './options';
import { setUi, ui } from './ui';
import { T, Tr } from '../../ui/i18n';

export const SheetViewsDialog = ({ store, close }: DialogProps) => {
  const s = store.sheet;
  const [sel, setSel] = useState<string | null>(activeViewOf(s));
  const [name, setName] = useState('');
  const [editing, setEditing] = useState<string | null>(null);
  const views = s.sheetViews ?? [];
  const current = activeViewOf(s);
  const selected = views.find((v) => v.id === sel);
  useEffect(() => { setName(selected?.name ?? ''); }, [sel]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Dialog title={T('Sheet views: {name}', { name: s.name })} close={close} width={520}>
      <p className="muted small">{T('A sheet view keeps its own filter, sort and hidden rows and columns, so you can look at the data your way without changing the default view.')}</p>
      <div className="ws-viewlist" role="listbox" aria-label={T('Views')}>
        <button role="option" aria-selected={sel === null} className={sel === null ? 'on' : ''} onClick={() => setSel(null)} onDoubleClick={() => { switchView(store, null); close(); }}>{T('Default')}{current === null && <em> · {T('showing')}</em>}
        </button>
        {current === '' && <button role="option" aria-selected={sel === ''} className={sel === '' ? 'on' : ''} onClick={() => setSel('')}>{T('Temporary view')}{' '}<em> · {T('showing')}</em></button>}
        {views.map((v) => (
          <button key={v.id} role="option" aria-selected={sel === v.id} className={sel === v.id ? 'on' : ''} onClick={() => setSel(v.id!)} onDoubleClick={() => { switchView(store, v.id!); close(); }}>
            {editing === v.id
              ? <input autoFocus value={name} onClick={(e) => e.stopPropagation()} onChange={(e) => setName(e.target.value)} onBlur={() => { renameView(store, v.id!, name); setEditing(null); }} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') { renameView(store, v.id!, name); setEditing(null); } if (e.key === 'Escape') setEditing(null); }} />
              : <>{v.name}{current === v.id && <em> · {T('showing')}</em>}<span className="muted small">{v.filter?.columns.length ? ` · ${T('{count, plural, one {# filter} other {# filters}}', { count: v.filter.columns.length })}` : ''}{v.filter?.sort?.length ? ` · ${T('sorted')}` : ''}{v.hiddenRows.length + v.hiddenCols.length ? ` · ${T('{count} hidden', { count: v.hiddenRows.length + v.hiddenCols.length })}` : ''}</span></>}
          </button>
        ))}
      </div>
      <div className="b-chips">
        <button className="primary" disabled={sel === current} onClick={() => { switchView(store, sel); close(); }}>{T('Switch to')}</button>
        <button onClick={() => { const v = createView(store); setSel(v.id!); }}>{T('New')}</button>
        <button disabled={!selected} onClick={() => { if (selected) { setEditing(selected.id!); setName(selected.name); } }}>{T('Rename')}</button>
        <button disabled={!selected} onClick={() => selected && duplicateView(store, selected.id!)}>{T('Duplicate')}</button>
        <button disabled={!selected} onClick={() => { if (selected && confirm(T('Delete the view "{name}"?', { name: selected.name }))) { deleteView(store, selected.id!); setSel(null); } }}>{T('Delete')}</button>
        <button disabled={current === ''} onClick={() => { switchView(store, ''); close(); }}>{T('Temporary view')}</button>
      </div>
      {current === '' && (
        <div className="grid2">
          <Field label={T('Keep the temporary view as')}><input value={name} placeholder={T('View name')} onChange={(e) => setName(e.target.value)} /></Field>
          <div className="ws-bottom"><button onClick={() => { keepTemporary(store, name || T('View')); close(); }}>{T('Keep')}</button></div>
        </div>
      )}
    </Dialog>
  );
};

export const CustomViewsDialog = ({ store, close }: DialogProps) => {
  const views = store.wb.customViews ?? [];
  const [sel, setSel] = useState<string | null>(views[0]?.guid ?? null);
  const [adding, setAdding] = useState(false);
  const [name, setName] = useState('');
  const [print, setPrint] = useState(true);
  const [hidden, setHidden] = useState(true);
  return (
    <Dialog title={T('Custom views')} close={close} width={480}>
      <p className="muted small">{T('A custom view remembers the whole workbook\'s display: which sheets show, zoom, frozen panes, selection, and if you like hidden rows and columns, filters and print settings.')}</p>
      {!adding ? <>
        <div className="ws-viewlist" role="listbox" aria-label={T('Custom views')}>
          {!views.length && <div className="muted small ws-pad">{T('No custom views yet.')}</div>}
          {views.map((v) => (
            <button key={v.guid} role="option" aria-selected={sel === v.guid} className={sel === v.guid ? 'on' : ''} onClick={() => setSel(v.guid)} onDoubleClick={() => { showCustomView(store, v.guid); close(); }}>
              {v.name}<span className="muted small">{v.includePrint ? T(' · print settings') : ''}{v.includeHidden ? T(' · hidden rows, columns and filters') : ''}</span>
            </button>
          ))}
        </div>
        <div className="b-chips">
          <button className="primary" disabled={!sel} onClick={() => { if (sel) { showCustomView(store, sel); close(); } }}>{T('Show')}</button>
          <button onClick={() => setAdding(true)}>{T('Add…')}</button>
          <button disabled={!sel} onClick={() => { const v = views.find((x) => x.guid === sel); if (v && confirm(T('Delete the view "{name}"?', { name: v.name }))) { deleteCustomView(store, v.guid); setSel(null); } }}>{T('Delete')}</button>
        </div>
      </> : <>
        <Field label={T('Name')}><input autoFocus value={name} onChange={(e) => setName(e.target.value)} maxLength={255} /></Field>
        <Check label={T('Include print settings')} checked={print} onChange={setPrint} />
        <Check label={T('Include hidden rows, columns and filter settings')} checked={hidden} onChange={setHidden} />
        <div className="b-chips">
          <button className="primary" onClick={() => { if (!name.trim()) { store.notify(T('Give the view a name.')); return; } addCustomView(store, name, print, hidden); setAdding(false); setSel(store.wb.customViews?.at(-1)?.guid ?? null); }}>{T('Add')}</button>
          <button onClick={() => setAdding(false)}>{T('Back')}</button>
        </div>
      </>}
    </Dialog>
  );
};

/** Right-click menu of the status bar: which selection figures to show. */
export const StatusMenu = ({ store }: { store: GridStore }) => {
  const ref = useRef<HTMLDivElement>(null);
  const m = ui.statusMenu;
  useEffect(() => {
    if (!m) return;
    const f = (e: PointerEvent) => { if (!ref.current?.contains(e.target as Node)) setUi(store, { statusMenu: null }); };
    const k = (e: KeyboardEvent) => { if (e.key === 'Escape') setUi(store, { statusMenu: null }); };
    document.addEventListener('pointerdown', f, true);
    document.addEventListener('keydown', k, true);
    return () => { document.removeEventListener('pointerdown', f, true); document.removeEventListener('keydown', k, true); };
  }, [m, store]);
  if (!m) return null;
  const toggle = (k: StatusItem) => {
    const items = settings.statusItems.includes(k) ? settings.statusItems.filter((x) => x !== k) : [...settings.statusItems, k];
    saveSettings(store, { statusItems: STATUS_ITEMS.map(([x]) => x).filter((x) => items.includes(x)) });
    applyDeviceSettings(store);
    store.emit();
  };
  return (
    <div ref={ref} className="menu ws-statusmenu" role="menu" aria-label={T('Status bar')} style={{ left: Math.min(m.x, window.innerWidth - 230), bottom: window.innerHeight - m.y + 4 }}>
      <div className="ws-errtitle">{T('Show on the status bar')}</div>
      {STATUS_ITEMS.map(([k, l]) => (
        <button key={k} role="menuitemcheckbox" aria-checked={settings.statusItems.includes(k)} onClick={() => toggle(k)}>
          <span className="ws-tick">{settings.statusItems.includes(k) ? '✓' : ''}</span>{Tr(l)}
        </button>
      ))}
    </div>
  );
};
