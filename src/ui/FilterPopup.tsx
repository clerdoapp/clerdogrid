import { useMemo, useState } from 'react';
import { ArrowDownAZ, ArrowUpAZ, FilterX } from 'lucide-react';
import { filterValues, setFilterColumn, sortRange } from '../app/ops/data';
import type { CfOperator } from '../model/workbook';
import type { GridStore } from '../app/store';

/** Is there a filter (or table header) button at this cell? */
export const filterPopupFor = (store: GridStore, row: number, col: number): boolean => {
  const af = store.sheet.autoFilter;
  if (af && row === af.range.r1 && col >= af.range.c1 && col <= af.range.c2) return true;
  return false;
};

const OPS: [CfOperator, string][] = [['equal', 'equals'], ['notEqual', 'does not equal'], ['greaterThan', 'is greater than'], ['greaterThanOrEqual', 'is greater than or equal to'], ['lessThan', 'is less than'], ['lessThanOrEqual', 'is less than or equal to']];

export const FilterPopup = ({ store, col, x, y, onClose }: { store: GridStore; col: number; x: number; y: number; onClose: () => void }) => {
  const af = store.sheet.autoFilter!;
  const values = useMemo(() => filterValues(store, col), [store, col]);
  const current = af.columns.find((f) => f.col === col);
  const [checked, setChecked] = useState<Set<string>>(() => new Set(current?.values ?? values.map((v) => v.text)));
  const [query, setQuery] = useState('');
  const [op, setOp] = useState<CfOperator>(current?.custom?.rules[0]?.operator ?? 'greaterThan');
  const [opVal, setOpVal] = useState(current?.custom?.rules[0]?.value ?? '');
  const shown = values.filter((v) => !query || v.text.toLowerCase().includes(query.toLowerCase()));
  const all = shown.every((v) => checked.has(v.text));

  const apply = () => {
    if (opVal.trim()) setFilterColumn(store, { col, custom: { rules: [{ operator: op, value: opVal.trim() }] } }, col);
    else if (checked.size === values.length && !query) setFilterColumn(store, null, col);
    else setFilterColumn(store, { col, values: [...checked].filter((v) => !query || shown.some((s) => s.text === v)), blanks: checked.has('') }, col);
    onClose();
  };
  const sort = (descending: boolean) => {
    sortRange(store, af.range, [{ col, descending }], true);
    onClose();
  };
  return (
    <div className="popup filterpop" style={{ left: Math.max(4, x), top: y }} role="dialog" aria-label="Filter" onPointerDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => { if (e.key === 'Escape') onClose(); if (e.key === 'Enter') apply(); }}>
      <button className="tool" onClick={() => sort(false)}><ArrowDownAZ size={15} strokeWidth={1.9} /> Sort smallest to largest</button>
      <button className="tool" onClick={() => sort(true)}><ArrowUpAZ size={15} strokeWidth={1.9} /> Sort largest to smallest</button>
      <button className="tool" disabled={!current} onClick={() => { setFilterColumn(store, null, col); onClose(); }}><FilterX size={15} strokeWidth={1.9} /> Clear filter</button>
      <div className="sep" />
      <div className="row2">
        <select value={op} onChange={(e) => setOp(e.target.value as CfOperator)} aria-label="Condition">{OPS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select>
        <input value={opVal} onChange={(e) => setOpVal(e.target.value)} placeholder="Value (* and ? work)" aria-label="Condition value" />
      </div>
      <input className="search" autoFocus value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search" aria-label="Search values" />
      <div className="checklist">
        <label><input type="checkbox" checked={all} onChange={() => { const n = new Set(checked); for (const v of shown) { if (all) n.delete(v.text); else n.add(v.text); } setChecked(n); }} /> (Select all)</label>
        {shown.slice(0, 2000).map((v) => (
          <label key={v.text}><input type="checkbox" checked={checked.has(v.text)} onChange={() => { const n = new Set(checked); if (n.has(v.text)) n.delete(v.text); else n.add(v.text); setChecked(n); }} /> {v.text === '' ? '(Blanks)' : v.text} <span className="muted">{v.count}</span></label>
        ))}
      </div>
      <div className="actions"><button onClick={onClose}>Cancel</button><button className="primary" onClick={apply}>OK</button></div>
    </div>
  );
};
