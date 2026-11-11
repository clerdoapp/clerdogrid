import { useMemo, useState } from 'react';
import { allCommands, run } from '../../app/commands';
import { searchCatalog, signature } from '../../engine/catalog';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { colPx, pxToChars } from '../../app/layout';
import { changeStructure, renameSheet, setSheetState, shiftCells } from '../../app/ops/structure';
import { Dialog, Field, type DialogProps } from './Dialog';
import { prettyKeys } from '../keys';
import { focusGrid } from '../FormulaBar';
import { gotoReference } from '../../app/goto';
import { N, T, Tr } from '../i18n';

export const PaletteDialog = ({ store, close }: DialogProps) => {
  const [q, setQ] = useState('');
  const [sel, setSel] = useState(0);
  const items = useMemo(() => {
    const cmds = allCommands().filter((c) => !q || c.label.toLowerCase().includes(q.toLowerCase()) || Tr(c.label).toLowerCase().includes(q.toLowerCase()) || c.group?.toLowerCase().includes(q.toLowerCase()) || Tr(c.group).toLowerCase().includes(q.toLowerCase()))
      .filter((c) => !c.enabled || c.enabled(store)).slice(0, 30).map((c) => ({ key: c.id, label: Tr(c.label), hint: c.keys ? prettyKeys(c.keys) : Tr(c.group) || '', go: () => run(store, c.id) }));
    const fns = q.length >= 2 ? searchCatalog(q, 8).map((f) => ({ key: `fn:${f.name}`, label: signature(f), hint: Tr(f.category), go: () => store.beginEdit(`=${f.name}(`) })) : [];
    return [...cmds, ...fns];
  }, [q, store]);
  const pick = (i: number) => { const it = items[i]; close(); if (it) { focusGrid(); it.go(); } };
  return (
    <div className="backdrop top" onPointerDown={(e) => { if (e.target === e.currentTarget) close(); }}>
      <div className="dialog palette" role="dialog" aria-label={T('Search commands')}>
        <input autoFocus className="palinput" placeholder={T('Type a command or a function name')} value={q}
          onChange={(e) => { setQ(e.target.value); setSel(0); }}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') { e.preventDefault(); setSel(Math.min(items.length - 1, sel + 1)); }
            if (e.key === 'ArrowUp') { e.preventDefault(); setSel(Math.max(0, sel - 1)); }
            if (e.key === 'Enter') pick(sel);
            if (e.key === 'Escape') close();
          }} />
        <div className="pallist" role="listbox">
          {items.map((it, i) => (
            <button key={it.key} role="option" aria-selected={i === sel} className={i === sel ? 'on' : ''} onMouseEnter={() => setSel(i)} onClick={() => pick(i)}>
              <span>{Tr(it.label)}</span><kbd>{Tr(it.hint)}</kbd>
            </button>
          ))}
          {!items.length && <div className="muted pad">{T('Nothing matches.')}</div>}
        </div>
      </div>
    </div>
  );
};

export const RenameSheetDialog = ({ store, data, close }: DialogProps<{ index: number }>) => {
  const i = data?.index ?? store.sheetIndex;
  const [name, setName] = useState(store.wb.sheets[i]!.name);
  return (
    <Dialog title={T('Rename sheet')} close={close} onOk={() => renameSheet(store, i, name)}>
      <Field label={T('Sheet name')}><input autoFocus value={name} maxLength={31} onChange={(e) => setName(e.target.value)} /></Field>
    </Dialog>
  );
};

export const UnhideSheetDialog = ({ store, close }: DialogProps) => {
  const hidden = store.wb.sheets.map((s, i) => ({ s, i })).filter((x) => x.s.state === 'hidden');
  const [pick, setPick] = useState(hidden[0]?.i ?? -1);
  return (
    <Dialog title={T('Unhide sheet')} close={close} onOk={() => { if (pick >= 0) setSheetState(store, pick, 'visible'); }}>
      <div className="listbox">{hidden.map(({ s, i }) => <label key={i}><input type="radio" checked={pick === i} onChange={() => setPick(i)} /> {s.name}</label>)}</div>
      {!hidden.length && <p className="muted">{T('No sheets are hidden.')}</p>}
    </Dialog>
  );
};

export const InsertCellsDialog = ({ store, data, close }: DialogProps<{ insert: boolean }>) => {
  const insert = data?.insert ?? true;
  const [how, setHow] = useState(insert ? 'down' : 'up');
  const r = store.range;
  const opts = insert ? [['right', 'Shift cells right'], ['down', 'Shift cells down'], ['row', N('Entire row')], ['col', N('Entire column')]] : [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', N('Entire row')], ['col', N('Entire column')]];
  return (
    <Dialog title={insert ? T('Insert') : T('Delete')} close={close} width={300} onOk={() => {
      if (how === 'row') changeStructure(store, 'row', r.r1, (insert ? 1 : -1) * (r.r2 - r.r1 + 1));
      else if (how === 'col') changeStructure(store, 'col', r.c1, (insert ? 1 : -1) * (r.c2 - r.c1 + 1));
      else shiftCells(store, r, how as 'down' | 'right' | 'up' | 'left');
    }}>
      <div className="listbox">{opts.map(([v, l]) => <label key={v}><input type="radio" checked={how === v} onChange={() => setHow(v!)} /> {Tr(l)}</label>)}</div>
    </Dialog>
  );
};

export const SizeDialog = ({ store, data, close }: DialogProps<{ axis: 'row' | 'col' }>) => {
  const axis = data?.axis ?? 'col';
  const r = store.range;
  const cur = axis === 'col' ? store.sheet.cols.get(r.c1)?.width ?? store.sheet.defaultColWidth : store.sheet.rows.get(r.r1)?.height ?? store.sheet.defaultRowHeight;
  const [v, setV] = useState(String(cur));
  return (
    <Dialog title={axis === 'col' ? T('Column width') : T('Row height')} width={300} close={close} onOk={() => {
      const n = Number(v);
      if (!Number.isFinite(n) || n < 0 || n > (axis === 'col' ? 255 : 409)) { store.notify(axis === 'col' ? T('Enter a width from 0 to 255.') : T('Enter a height from 0 to 409.')); return false; }
      store.change(axis === 'col' ? N('Column width') : N('Row height'), 'sheet', () => {
        if (axis === 'col') for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 16383); c++) { const info = store.sheet.cols.get(c) ?? {}; info.width = n; info.hidden = n === 0 || undefined; store.sheet.cols.set(c, info); }
        else for (let row = r.r1; row <= Math.min(r.r2, MAX_ROWS - 1, r.r1 + 100000); row++) { const info = store.sheet.rows.get(row) ?? {}; info.height = n; info.hidden = n === 0 || undefined; store.sheet.rows.set(row, info); }
      });
    }}>
      <Field label={axis === 'col' ? T('Width (characters)') : T('Height (points)')}><input autoFocus type="number" min={0} step={axis === 'col' ? 0.01 : 0.25} value={v} onChange={(e) => setV(e.target.value)} /></Field>
      {axis === 'col' && <p className="muted">{T('{count, plural, one {# pixel} other {# pixels}}', { count: colPx(Number(v) || 0) })}</p>}
      <p className="muted small">{T('Default:')}{' '}{axis === 'col' ? pxToChars(colPx(store.sheet.defaultColWidth)) : store.sheet.defaultRowHeight}</p>
    </Dialog>
  );
};

export const GotoDialog = ({ store, close }: DialogProps) => {
  const [ref, setRef] = useState('');
  return (
    <Dialog title={T('Go to')} width={360} close={close} onOk={() => gotoReference(store, ref)}>
      <div className="listbox">{store.wb.names.map((n) => <button key={n.name} className="tool" onClick={() => setRef(n.name)}>{n.name}</button>)}</div>
      <Field label={T('Reference')}><input autoFocus value={ref} onChange={(e) => setRef(e.target.value)} placeholder={T('A1, Sheet2!B5:C9 or a name')} /></Field>
    </Dialog>
  );
};

export { MAX_COLS };
