import type React from 'react';
import { structureLocked } from '../app/protection';
import { useEffect, useRef } from 'react';
import { command, openDialog, run } from '../app/commands';
import { deleteSheet, duplicateSheet, setSheetState } from '../app/ops/structure';
import type { GridStore } from '../app/store';
import { prettyKeys } from './keys';
import { menuProviderItems } from './panelRegistry';
import { N, T, Tr } from './i18n';

type Item = { id: string; label?: string; arg?: unknown } | { action: () => void; label: string; disabled?: boolean } | 'sep';

const cellItems: Item[] = [
  { id: 'edit.cut' }, { id: 'edit.copy' }, { id: 'edit.paste' }, { id: 'edit.pasteValues' }, { id: 'edit.pasteFormats' }, { id: 'edit.pasteTranspose' }, { id: 'edit.pasteSpecial' }, 'sep',
  { id: 'cells.insertDialog', label: N('Insert…') }, { id: 'cells.deleteDialog', label: N('Delete…') }, { id: 'edit.clearContents' }, 'sep',
  { id: 'data.sortAsc' }, { id: 'data.sortDesc' }, { id: 'data.filter' }, 'sep',
  { id: 'insert.note', label: N('Insert note') }, { id: 'insert.link', label: N('Link…') }, { id: 'formula.defineName' }, { id: 'data.validation' }, 'sep',
  { id: 'cells.formatDialog' }, { id: 'insert.conditional' },
];
const colItems: Item[] = [
  { id: 'edit.cut' }, { id: 'edit.copy' }, { id: 'edit.paste' }, 'sep', { id: 'cells.insertCols', label: N('Insert columns') }, { id: 'cells.deleteCols', label: N('Delete columns') }, { id: 'edit.clearContents' }, 'sep',
  { id: 'cells.formatDialog' }, { id: 'cells.colWidth' }, { id: 'cells.autofitCols' }, { id: 'cells.hideCols', label: N('Hide') }, { id: 'cells.unhideCols', label: N('Unhide') }, 'sep', { id: 'data.sortAsc' }, { id: 'data.sortDesc' },
];
const rowItems: Item[] = [
  { id: 'edit.cut' }, { id: 'edit.copy' }, { id: 'edit.paste' }, 'sep', { id: 'cells.insertRows', label: N('Insert rows') }, { id: 'cells.deleteRows', label: N('Delete rows') }, { id: 'edit.clearContents' }, 'sep',
  { id: 'cells.formatDialog' }, { id: 'cells.rowHeight' }, { id: 'cells.autofitRows' }, { id: 'cells.hideRows', label: N('Hide') }, { id: 'cells.unhideRows', label: N('Unhide') },
];

const TAB_COLORS = ['', '#c00000', '#ff8a00', '#ffc000', '#00b050', '#00b0f0', '#0C7A4B', '#7030a0', '#7f7f7f'];

export const ContextMenu = ({ store }: { store: GridStore }) => {
  const ref = useRef<HTMLDivElement>(null);
  const m = store.menu!;
  const close = () => { store.menu = null; store.emit(); };
  useEffect(() => {
    const onDown = (e: PointerEvent) => { if (!ref.current?.contains(e.target as Node)) close(); };
    window.addEventListener('pointerdown', onDown, true);
    ref.current?.querySelector('button')?.focus();
    return () => window.removeEventListener('pointerdown', onDown, true);
  });
  let items: Item[];
  let extra: React.ReactElement | null = null;
  if (m.kind === 'sheet') {
    const data = m.data as { index?: number; list?: boolean };
    if (data.list) {
      items = store.wb.sheets.map((s, i) => ({ label: s.state !== 'visible' ? T('{name} (hidden)', { name: s.name }) : s.name, action: () => { if (s.state !== 'visible') setSheetState(store, i, 'visible'); if (store.wb.sheets[i]?.state === 'visible') store.setActiveSheet(i); } }));
    } else {
      const i = data.index ?? store.sheetIndex;
      const hidden = store.wb.sheets.filter((s) => s.state === 'hidden');
      items = [
        { id: 'sheet.insert', label: N('Insert sheet') }, { label: N('Delete'), action: () => deleteSheet(store, i) }, { label: N('Duplicate'), action: () => duplicateSheet(store, i) },
        { label: N('Rename'), action: () => openDialog(store, 'renameSheet', { index: i }) }, 'sep',
        { label: N('Hide'), action: () => setSheetState(store, i, 'hidden') }, { label: N('Unhide…'), action: () => openDialog(store, 'unhideSheet'), disabled: !hidden.length }, 'sep',
        { id: 'sheet.protect' }, { id: 'review.protectWorkbook' }, { label: N('Select all sheets'), action: () => undefined, disabled: true },
      ];
      extra = (
        <div className="menucolors"><span>{T('Tab color')}</span>{TAB_COLORS.map((c) => (
          <button key={c || 'none'} className={`sw${c ? '' : ' none'}`} style={c ? { background: c } : undefined} aria-label={c || T('No color')}
            onClick={() => { if (!structureLocked(store)) store.change(N('Tab color'), 'book', () => { store.wb.sheets[i]!.tabColor = c || undefined; }); close(); }} />
        ))}</div>
      );
    }
  } else items = [...menuProviderItems(store, m.kind), ...(m.kind === 'col' ? colItems : m.kind === 'row' ? rowItems : cellItems)];
  const x = Math.min(m.x, window.innerWidth - 260), y = Math.min(m.y, window.innerHeight - Math.min(items.length * 30 + 60, window.innerHeight - 10));
  return (
    <div className="menu" ref={ref} role="menu" style={{ left: x, top: Math.max(4, y) }} onKeyDown={(e) => {
      if (e.key === 'Escape') close();
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const btns = [...ref.current!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
        const i = btns.indexOf(document.activeElement as HTMLButtonElement);
        btns[(i + (e.key === 'ArrowDown' ? 1 : -1) + btns.length) % btns.length]?.focus();
      }
    }}>
      {items.map((it, i) => {
        if (it === 'sep') return <div key={i} className="sep" role="separator" />;
        if ('action' in it) return <button key={i} role="menuitem" disabled={it.disabled} onClick={() => { close(); it.action(); }}>{Tr(it.label)}</button>;
        const c = command(it.id);
        if (!c) return null;
        return (
          <button key={i} role="menuitem" disabled={c.enabled ? !c.enabled(store) : false} onClick={() => { close(); run(store, it.id, it.arg); }}>
            <span>{Tr(it.label ?? c.label)}</span>{c.keys && <kbd>{prettyKeys(c.keys)}</kbd>}
          </button>
        );
      })}
      {extra}
    </div>
  );
};
