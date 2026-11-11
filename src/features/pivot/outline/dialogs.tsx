// Subtotal, Consolidate and outline settings dialogs.
import { useMemo, useState } from 'react';
import type { GridStore } from '../../../app/store';
import { colName, type CellRange } from '../../../model/workbook';
import { Dialog, Field, type DialogProps } from '../../../ui/dialogs/Dialog';
import { Check } from '../../../ui/dialogs/shared';
import { resolveRange, rangeText } from '../../../ui/dialogs/a/util';
import {
  CONSOLIDATE_FUNCS, SUBTOTAL_FUNCS, applySubtotals, consolidate, hasSubtotals, removeSubtotals, subtotalRange, type ConsolidateFn,
} from '../../../app/ops/outline';
import { N, T, Tr } from '../../../ui/i18n';

const header = (store: GridStore, r: CellRange, c: number): string => {
  const v = store.value(r.r1, c);
  return v === null || v === '' ? `Column ${colName(c)}` : String(typeof v === 'object' ? v.code : v);
};

export const SubtotalDialog = ({ store, close }: DialogProps) => {
  const range = useMemo(() => subtotalRange(store), [store]);
  const cols = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i);
  const numericCols = cols.filter((c) => typeof store.value(range.r1 + 1, c) === 'number');
  const [groupBy, setGroupBy] = useState(range.c1);
  const [fn, setFn] = useState(9);
  const [picked, setPicked] = useState<Set<number>>(new Set(numericCols.length ? [numericCols[numericCols.length - 1]!] : [range.c2]));
  const existing = hasSubtotals(store, range);
  const [replace, setReplace] = useState(true);
  const [breaks, setBreaks] = useState(false);
  const [below, setBelow] = useState(store.sheet.outlinePr?.summaryBelow !== false);
  if (range.r2 <= range.r1) return <Dialog title={T('Subtotal')} close={close}><p>{T('Select a list with a header row and some data first.')}</p></Dialog>;
  const ok = () => {
    if (!picked.size) { store.notify(T('Pick at least one column to subtotal.')); return false; }
    applySubtotals(store, range, { groupBy, fn, columns: [...picked].sort((a, b) => a - b), replace, pageBreaks: breaks, summaryBelow: below });
    return true;
  };
  return (
    <Dialog title={T('Subtotal')} close={close} width={420} onOk={ok}
      footer={existing ? <button onClick={() => { removeSubtotals(store, range); close(); }}>{T('Remove all')}</button> : undefined}>
      <p className="muted small">{T('List:')}{' '}{rangeText(store, range)}</p>
      <Field label={T('At each change in')}><select value={groupBy} onChange={(e) => setGroupBy(+e.target.value)}>{cols.map((c) => <option key={c} value={c}>{header(store, range, c)}</option>)}</select></Field>
      <Field label={T('Use function')}><select value={fn} onChange={(e) => setFn(+e.target.value)}>{SUBTOTAL_FUNCS.map(([n, l]) => <option key={n} value={n}>{Tr(l)}</option>)}</select></Field>
      <div className="seclabel">{T('Add subtotal to')}</div>
      <div className="checks pvlist">
        {cols.map((c) => <label key={c}><input type="checkbox" checked={picked.has(c)} onChange={(e) => { const n = new Set(picked); if (e.target.checked) n.add(c); else n.delete(c); setPicked(n); }} /> {header(store, range, c)}</label>)}
      </div>
      {existing && <Check label={T('Replace current subtotals')} checked={replace} onChange={setReplace} />}
      <Check label={T('Page break between groups')} checked={breaks} onChange={setBreaks} />
      <Check label={T('Summary below data')} checked={below} onChange={setBelow} />
    </Dialog>
  );
};

export const ConsolidateDialog = ({ store, close }: DialogProps) => {
  const [fn, setFn] = useState<ConsolidateFn>('sum');
  const [refText, setRefText] = useState('');
  const [refs, setRefs] = useState<string[]>([]);
  const [top, setTop] = useState(false);
  const [left, setLeft] = useState(false);
  const [links, setLinks] = useState(false);
  const [pick, setPick] = useState<number | null>(null);
  const add = () => {
    const t = refText.trim();
    if (!t) return;
    if (!resolveRange(store, t)) { store.notify(T('Enter a range such as Sheet2!A1:D10.')); return; }
    if (!refs.includes(t)) setRefs([...refs, t]);
    setRefText('');
  };
  const ok = () => {
    const list = [...refs, ...(refText.trim() && !refs.includes(refText.trim()) ? [refText.trim()] : [])];
    const resolved = list.map((t) => resolveRange(store, t));
    if (resolved.some((r) => !r)) { store.notify(T('One of the references is not a range.')); return false; }
    const err = consolidate(store, { fn, refs: resolved.map((r) => ({ sheet: r!.sheet, range: r!.range })), topRow: top, leftCol: left, links });
    if (err) { store.notify(err); return false; }
    return true;
  };
  return (
    <Dialog title={T('Consolidate')} close={close} width={460} onOk={ok}>
      <p className="muted small">{T('Results go at the active cell.')}</p>
      <Field label={T('Function')}><select value={fn} onChange={(e) => setFn(e.target.value as ConsolidateFn)}>{CONSOLIDATE_FUNCS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select></Field>
      <Field label={T('Reference')}><input value={refText} onChange={(e) => setRefText(e.target.value)} placeholder={T('Sheet2!A1:D10')} spellCheck={false} onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); add(); } }} /></Field>
      <div className="pvpop-row"><button onClick={add}>{T('Add')}</button><button disabled={pick === null} onClick={() => { setRefs(refs.filter((_, i) => i !== pick)); setPick(null); }}>{T('Delete')}</button></div>
      <div className="listbox pvreflist" role="listbox" aria-label={T('All references')}>
        {refs.map((r, i) => <label key={r}><input type="radio" name="cref" checked={pick === i} onChange={() => { setPick(i); setRefText(r); }} /> {r}</label>)}
        {!refs.length && <span className="muted small">{T('No references yet.')}</span>}
      </div>
      <div className="seclabel">{T('Use labels in')}</div>
      <Check label={T('Top row')} checked={top} onChange={setTop} />
      <Check label={T('Left column')} checked={left} onChange={setLeft} />
      <Check label={T('Create links to source data')} checked={links} onChange={setLinks} />
    </Dialog>
  );
};

export const OutlineSettingsDialog = ({ store, close }: DialogProps) => {
  const [below, setBelow] = useState(store.sheet.outlinePr?.summaryBelow !== false);
  const [right, setRight] = useState(store.sheet.outlinePr?.summaryRight !== false);
  return (
    <Dialog title={T('Outline settings')} close={close} width={360} onOk={() => { store.change(N('Outline settings'), 'sheet', () => { store.sheet.outlinePr = below && right ? undefined : { summaryBelow: below, summaryRight: right }; }); return true; }}>
      <Check label={T('Summary rows below detail')} checked={below} onChange={setBelow} />
      <Check label={T('Summary columns right of detail')} checked={right} onChange={setRight} />
    </Dialog>
  );
};
