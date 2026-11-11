// PivotTable dialogs: value field and field settings, options, grouping, calculated fields and
// items, filters, data source, report filter pages and moving a table.
import { useMemo, useState } from 'react';
import type { GridStore } from '../../app/store';
import { parseAddr, rangeAddr } from '../../model/workbook';
import { quoteSheet } from '../../engine/parser';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check, Radio, Tabs } from '../../ui/dialogs/shared';
import { resolveRange } from '../../ui/dialogs/a/util';
import {
  VALUES_FIELD,
  type DataField, type LabelOp, type PivotAgg, type PivotFilter, type PivotTable, type ShowAs, type Subtotal, type ValueOp,
} from '../../model/pivot';
import { cacheOf, changeSource, movePivot, sourceText, type PivotRef } from './apply';
import { fieldItems } from './cache';
import { AGG_LABEL, dataFieldName, fieldCaption, normalizeFields } from './compute';
import { parsePivotFormula } from './formula';
import { setField, uniqueDataName } from './model';
import { deleteCalcField, deleteCalcItem, groupRange, saveCalcField, saveCalcItem, showFilterPages, update } from './ops';
import { currentPivot } from './ui';
import { N, T, Tr } from '../../ui/i18n';

type Target = { sheet: number; name: string };
const refOf = (store: GridStore, t?: Target): PivotRef | null => {
  if (!t) return currentPivot(store);
  const index = store.wb.sheets[t.sheet]?.pivots?.findIndex((p) => p.name === t.name) ?? -1;
  return index >= 0 ? { sheet: t.sheet, index, pt: store.wb.sheets[t.sheet]!.pivots![index]! } : null;
};

const AGGS: PivotAgg[] = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp', 'distinctCount'];
const SHOW_AS: [ShowAs, string, 'none' | 'field' | 'item'][] = [
  ['normal', N('No calculation'), 'none'], ['percentOfTotal', N('% of grand total'), 'none'], ['percentOfCol', N('% of column total'), 'none'], ['percentOfRow', N('% of row total'), 'none'],
  ['percent', N('% of'), 'item'], ['percentOfParentRow', N('% of parent row total'), 'none'], ['percentOfParentCol', N('% of parent column total'), 'none'],
  ['percentOfParent', N('% of parent total'), 'field'], ['difference', N('Difference from'), 'item'], ['percentDiff', N('% difference from'), 'item'],
  ['runTotal', N('Running total in'), 'field'], ['percentOfRunningTotal', N('% running total in'), 'field'], ['rankAscending', N('Rank smallest to largest'), 'field'],
  ['rankDescending', N('Rank largest to smallest'), 'field'], ['index', N('Index'), 'none'],
];
const FORMATS: [string, string][] = [
  ['', N('Automatic')], [N('General'), N('General')], ['#,##0', N('Number (1,235)')], ['#,##0.00', N('Number (1,234.57)')], ['"$"#,##0.00', N('Currency')],
  ['0%', N('Percent (12%)')], ['0.00%', N('Percent (12.34%)')], ['m/d/yyyy', N('Date')], ['0.00E+00', N('Scientific')],
];

const missing = (close: () => void) => <Dialog title={T('PivotTable')} close={close}><p>{T('Select a cell in a PivotTable first.')}</p></Dialog>;

// ---------------------------------------------------------------- value field settings

export const ValueFieldDialog = ({ store, close, data }: DialogProps<Target & { index: number }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const df0 = ref?.pt.data[data?.index ?? 0];
  const [df, setDf] = useState<DataField | undefined>(df0 ? { ...df0 } : undefined);
  const [tab, setTab] = useState<'sum' | 'show'>('sum');
  if (!ref || !cache || !df) return missing(close);
  const axisFields = [...ref.pt.rows, ...ref.pt.cols].filter((f) => f >= 0);
  const mode = SHOW_AS.find((m) => m[0] === (df.showAs ?? 'normal'))!;
  const baseField = df.baseField ?? axisFields[0] ?? 0;
  const baseItems = useMemo(() => (cache ? normalizeFields(ref.pt, cache)[baseField]?.items ?? [] : []), [cache, ref.pt, baseField]); // eslint-disable-line react-hooks/exhaustive-deps
  const autoName = (agg: PivotAgg) => `${AGG_LABEL[agg]} of ${cache.fields[df.field]?.name ?? ''}`;
  const ok = () => {
    const name = df.name.trim() || autoName(df.agg);
    if (cache.fields.some((f) => f.name.toLowerCase() === name.toLowerCase())) { store.notify(T('That name is already a field name. Choose another one.')); return false; }
    if (ref.pt.data.some((d, i) => i !== data!.index && d.name.toLowerCase() === name.toLowerCase())) { store.notify(T('Another value field already has that name.')); return false; }
    const next: DataField = { ...df, name };
    if (!next.showAs || next.showAs === 'normal') { delete next.showAs; delete next.baseField; delete next.baseItem; }
    else if (mode[2] === 'none') { delete next.baseField; delete next.baseItem; }
    else { next.baseField = baseField; if (mode[2] === 'field') delete next.baseItem; else next.baseItem = df.baseItem ?? 'prev'; }
    if (!next.numFmt) delete next.numFmt;
    update(store, ref, (pt) => ({ ...pt, data: pt.data.map((d, i) => (i === data!.index ? next : d)) }), N('Value field settings'));
    return true;
  };
  return (
    <Dialog title={T('Value field settings')} close={close} width={460} onOk={ok}>
      <p className="muted small">{T('Source field:')}{' '}{cache.fields[df.field]?.name}</p>
      <Field label={T('Custom name')}><input value={df.name} onChange={(e) => setDf({ ...df, name: e.target.value })} /></Field>
      <Tabs tabs={[['sum', N('Summarize values by')], ['show', N('Show values as')]] as const} value={tab} onChange={setTab} />
      {tab === 'sum' ? (
        <div className="listbox pvlist" role="listbox" aria-label={T('Summarize values by')}>
          {AGGS.map((a) => (
            <label key={a}><input type="radio" name="pvagg" checked={df.agg === a} onChange={() => setDf({ ...df, agg: a, name: df.name === autoName(df.agg) || !df.name ? uniqueDataName(ref.pt, cache, autoName(a), data!.index) : df.name })} /> {AGG_LABEL[a]}</label>
          ))}
        </div>
      ) : (
        <>
          <Field label={T('Show values as')}>
            <select value={df.showAs ?? 'normal'} onChange={(e) => setDf({ ...df, showAs: e.target.value as ShowAs })}>{SHOW_AS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
          </Field>
          {mode[2] !== 'none' && (
            <div className="grid2">
              <Field label={T('Base field')}>
                <select value={baseField} onChange={(e) => setDf({ ...df, baseField: +e.target.value, baseItem: 'prev' })}>{axisFields.map((f) => <option key={f} value={f}>{fieldCaption(ref.pt, cache, f)}</option>)}</select>
              </Field>
              {mode[2] === 'item' && (
                <Field label={T('Base item')}>
                  <select value={String(df.baseItem ?? 'prev')} onChange={(e) => setDf({ ...df, baseItem: e.target.value === 'prev' || e.target.value === 'next' ? e.target.value : +e.target.value })}>
                    <option value="prev">{T('(previous)')}</option><option value="next">{T('(next)')}</option>
                    {baseItems.map((it, i) => <option key={i} value={i}>{it.name ?? (it.x >= 0 ? String(fieldItems(cache.fields[baseField]!)[it.x] ?? '(blank)') : '')}</option>)}
                  </select>
                </Field>
              )}
            </div>
          )}
          {mode[2] !== 'none' && !axisFields.length && <p className="muted small">{T('Add a row or column field to pick a base field.')}</p>}
        </>
      )}
      <Field label={T('Number format')}>
        <select value={df.numFmt ?? ''} onChange={(e) => setDf({ ...df, numFmt: e.target.value || undefined })}>
          {FORMATS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}
          {df.numFmt && !FORMATS.some((f) => f[0] === df.numFmt) && <option value={df.numFmt}>{df.numFmt}</option>}
        </select>
      </Field>
      <Field label={T('Or type a format code')}><input value={df.numFmt ?? ''} placeholder="#,##0.00" onChange={(e) => setDf({ ...df, numFmt: e.target.value || undefined })} spellCheck={false} /></Field>
    </Dialog>
  );
};

// ---------------------------------------------------------------- field settings

const SUBS: [Exclude<Subtotal, 'default'>, string][] = [
  ['sum', N('Sum')], ['count', N('Count')], ['average', N('Average')], ['max', N('Max')], ['min', N('Min')], ['product', N('Product')],
  ['countNums', N('Count Numbers')], ['stdDev', N('StdDev')], ['stdDevp', N('StdDevp')], ['var', N('Var')], ['varp', N('Varp')],
];

export const FieldSettingsDialog = ({ store, close, data }: DialogProps<Target & { field: number }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const f = data?.field ?? -1;
  const pf0 = ref?.pt.fields[f];
  const [name, setName] = useState(pf0?.name ?? '');
  const subs0 = pf0?.subtotals ?? ['default'];
  const [subMode, setSubMode] = useState<'auto' | 'none' | 'custom'>(!subs0.length ? 'none' : subs0.length === 1 && subs0[0] === 'default' ? 'auto' : 'custom');
  const [custom, setCustom] = useState<Set<Subtotal>>(new Set(subs0.filter((s) => s !== 'default')));
  const [form, setForm] = useState<'compact' | 'outline' | 'tabular'>(pf0 ? (pf0.outline === false ? 'tabular' : pf0.compact === false ? 'outline' : ref!.pt.layout === 'tabular' && pf0.outline === undefined ? 'tabular' : ref!.pt.layout === 'outline' && pf0.compact === undefined ? 'outline' : 'compact') : 'compact');
  const [top, setTop] = useState(pf0?.subtotalTop !== false);
  const [repeat, setRepeat] = useState(!!pf0?.repeatLabels);
  const [blank, setBlank] = useState(!!pf0?.blankRow);
  const [showAll, setShowAll] = useState(!!pf0?.showAll);
  const [pageBreak, setPageBreak] = useState(!!pf0?.pageBreak);
  const [numFmt, setNumFmt] = useState(pf0?.numFmt ?? '');
  const [tab, setTab] = useState<'sub' | 'layout'>('sub');
  if (!ref || !cache || !pf0) return missing(close);
  const ok = () => {
    const subtotals: Subtotal[] | undefined = subMode === 'auto' ? undefined : subMode === 'none' ? [] : [...custom];
    if (subMode === 'custom' && !custom.size) { store.notify(T('Pick at least one function, or choose None.')); return false; }
    update(store, ref, (pt) => setField(pt, f, {
      name: name.trim() && name.trim() !== cache.fields[f]!.name ? name.trim() : undefined, subtotals,
      compact: form === 'compact', outline: form !== 'tabular', subtotalTop: top ? undefined : false,
      repeatLabels: repeat || undefined, blankRow: blank || undefined, showAll: showAll || undefined, pageBreak: pageBreak || undefined, numFmt: numFmt || undefined,
    }), N('Field settings'));
    return true;
  };
  return (
    <Dialog title={T('Field settings')} close={close} width={460} onOk={ok}>
      <p className="muted small">{T('Source field:')}{' '}{cache.fields[f]!.name}</p>
      <Field label={T('Custom name')}><input value={name} placeholder={cache.fields[f]!.name} onChange={(e) => setName(e.target.value)} /></Field>
      <Tabs tabs={[['sub', N('Subtotals and filters')], ['layout', N('Layout and print')]] as const} value={tab} onChange={setTab} />
      {tab === 'sub' ? (
        <>
          <Radio name="pvsub" value={subMode} onChange={setSubMode} options={[['auto', N('Automatic')], ['none', N('None')], ['custom', N('Custom')]]} />
          {subMode === 'custom' && (
            <div className="checks pvlist">
              {SUBS.map(([v, l]) => <label key={v}><input type="checkbox" checked={custom.has(v)} onChange={(e) => { const n = new Set(custom); if (e.target.checked) n.add(v); else n.delete(v); setCustom(n); }} /> {Tr(l)}</label>)}
            </div>
          )}
          <Check label={T('Include new items in a manual filter')} checked={pf0.includeNew !== false} onChange={(v) => update(store, ref, (pt) => setField(pt, f, { includeNew: v ? undefined : false }), N('Field settings'))} />
        </>
      ) : (
        <>
          <Radio name="pvform" value={form} onChange={setForm} options={[['compact', N('Labels in one column (compact)')], ['outline', N('Labels in their own columns (outline)')], ['tabular', N('Labels in a table (tabular)')]]} />
          {form !== 'tabular' && <Check label={T('Subtotals at the top of each group')} checked={top} onChange={setTop} />}
          <Check label={T('Repeat item labels')} checked={repeat} onChange={setRepeat} />
          <Check label={T('Blank row after each item')} checked={blank} onChange={setBlank} />
          <Check label={T('Show items with no data')} checked={showAll} onChange={setShowAll} />
          <Check label={T('Page break after each item')} checked={pageBreak} onChange={setPageBreak} />
          <Field label={T('Number format for labels')}><input value={numFmt} placeholder={T('Automatic')} onChange={(e) => setNumFmt(e.target.value)} spellCheck={false} /></Field>
        </>
      )}
    </Dialog>
  );
};

// ---------------------------------------------------------------- options

export const PivotOptionsDialog = ({ store, close, data }: DialogProps<Target>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const [o, setO] = useState<PivotTable | undefined>(ref ? { ...ref.pt } : undefined);
  const [refreshOnLoad, setRefreshOnLoad] = useState(!!cache?.refreshOnLoad);
  const [dropMissing, setDropMissing] = useState(!!cache?.dropMissing);
  const [tab, setTab] = useState<'layout' | 'totals' | 'display' | 'data'>('layout');
  if (!ref || !cache || !o) return missing(close);
  const set = (p: Partial<PivotTable>) => setO({ ...o, ...p });
  const ok = () => {
    const name = o.name.trim();
    if (!name) { store.notify(T('Enter a name.')); return false; }
    if (name !== ref.pt.name && store.wb.sheets.some((s) => s.pivots?.some((p) => p.name.toLowerCase() === name.toLowerCase()))) { store.notify(T('Another PivotTable already has that name.')); return false; }
    if (refreshOnLoad !== !!cache.refreshOnLoad || dropMissing !== !!cache.dropMissing) {
      store.change(N('PivotTable options'), 'book', () => {
        store.wb.pivotCaches = (store.wb.pivotCaches ?? []).map((c) => (c.id === cache.id ? { ...c, refreshOnLoad: refreshOnLoad || undefined, dropMissing: dropMissing || undefined } : c));
      });
    }
    const renamed = name !== ref.pt.name;
    update(store, ref, () => ({ ...o, name }), N('PivotTable options'));
    if (renamed) store.change(N('Rename PivotTable'), 'book', () => {
      const s = store.wb.sheets[ref.sheet]!;
      s.pivots = s.pivots?.map((p) => (p.name === ref.pt.name ? { ...p, name } : p));
      for (const sc of store.wb.slicerCaches ?? []) if (sc.source.kind === 'pivot') sc.source.pivots = sc.source.pivots.map((x) => (x.name === ref.pt.name && x.sheet === s.name ? { ...x, name } : x));
      for (const sh of store.wb.sheets) for (const ch of sh.charts) if (ch.pivot === `${s.name}!${ref.pt.name}`) ch.pivot = `${s.name}!${name}`;
    });
    return true;
  };
  return (
    <Dialog title={T('PivotTable options')} close={close} width={480} onOk={ok}>
      <Field label={T('Name')}><input value={o.name} onChange={(e) => set({ name: e.target.value })} /></Field>
      <Tabs tabs={[['layout', N('Layout')], ['totals', N('Totals')], ['display', N('Display')], ['data', N('Data')]] as const} value={tab} onChange={setTab} />
      {tab === 'layout' && (
        <>
          <Check label={T('For empty cells show')} checked={o.showEmpty !== false && o.emptyText !== undefined} onChange={(v) => set({ showEmpty: v, emptyText: v ? o.emptyText ?? '' : undefined })} />
          <input aria-label={T('Text for empty cells')} value={o.emptyText ?? ''} disabled={o.emptyText === undefined} onChange={(e) => set({ emptyText: e.target.value })} />
          <Check label={T('For error values show')} checked={!!o.showError} onChange={(v) => set({ showError: v, errorText: v ? o.errorText ?? '' : o.errorText })} />
          <input aria-label={T('Text for errors')} value={o.errorText ?? ''} disabled={!o.showError} onChange={(e) => set({ errorText: e.target.value })} />
          <Check label={T('Fit column widths on update')} checked={o.autoFit} onChange={(v) => set({ autoFit: v })} />
          <Check label={T('Keep cell formatting on update')} checked={o.preserveFormatting} onChange={(v) => set({ preserveFormatting: v })} />
          <div className="grid2">
            <Field label={T('Report filters')}><select value={o.pageOverThenDown ? 'over' : 'down'} onChange={(e) => set({ pageOverThenDown: e.target.value === 'over' || undefined })}><option value="down">{T('Down, then over')}</option><option value="over">{T('Over, then down')}</option></select></Field>
            <Field label={T('Filters per column or row (0: all)')}><input type="number" min={0} max={255} value={o.pageWrap ?? 0} onChange={(e) => set({ pageWrap: Math.max(0, Math.min(255, +e.target.value || 0)) || undefined })} /></Field>
          </div>
        </>
      )}
      {tab === 'totals' && (
        <>
          <Check label={T('Show grand totals for rows (a total column)')} checked={o.rowGrand} onChange={(v) => set({ rowGrand: v })} />
          <Check label={T('Show grand totals for columns (a total row)')} checked={o.colGrand} onChange={(v) => set({ colGrand: v })} />
          <Field label={T('Grand total caption')}><input value={o.grandTotalCaption ?? ''} placeholder={T('Grand Total')} onChange={(e) => set({ grandTotalCaption: e.target.value || undefined })} /></Field>
        </>
      )}
      {tab === 'display' && (
        <>
          <Check label={T('Show expand and collapse buttons')} checked={o.showDrill} onChange={(v) => set({ showDrill: v })} />
          <Check label={T('Show field captions and filter buttons')} checked={o.showHeaders} onChange={(v) => set({ showHeaders: v })} />
          <div className="grid2">
            <Field label={T('Row labels caption')}><input value={o.rowHeaderCaption ?? ''} placeholder={T('Row Labels')} onChange={(e) => set({ rowHeaderCaption: e.target.value || undefined })} /></Field>
            <Field label={T('Column labels caption')}><input value={o.colHeaderCaption ?? ''} placeholder={T('Column Labels')} onChange={(e) => set({ colHeaderCaption: e.target.value || undefined })} /></Field>
            <Field label={T('Values caption')}><input value={o.dataCaption ?? ''} placeholder={T('Values')} onChange={(e) => set({ dataCaption: e.target.value || undefined })} /></Field>
            <Field label={T('Compact indent (characters)')}><input type="number" min={0} max={127} value={o.indent ?? 1} onChange={(e) => set({ indent: Math.max(0, Math.min(127, +e.target.value || 0)) })} /></Field>
          </div>
        </>
      )}
      {tab === 'data' && (
        <>
          <p className="muted small">{T('Source:')}{' '}{sourceText(cache.source)} · {T('{count, plural, one {# record} other {# records}}', { count: cache.records.length })}</p>
          <Check label={T('Refresh when the file opens')} checked={refreshOnLoad} onChange={setRefreshOnLoad} />
          <Check label={T('Drop items no longer in the data on refresh')} checked={dropMissing} onChange={setDropMissing} />
        </>
      )}
    </Dialog>
  );
};

// ---------------------------------------------------------------- grouping by ranges

const UNITS = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'] as const;

export const GroupDialog = ({ store, close, data }: DialogProps<Target & { field: number }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const f = data?.field ?? -1;
  const cf = cache?.fields[f];
  const nums = (cf?.items ?? []).filter((v): v is number => typeof v === 'number');
  const lo = nums.length ? Math.min(...nums) : 0, hi = nums.length ? Math.max(...nums) : 0;
  const dates = !!cf?.dates;
  const [start, setStart] = useState(''), [end, setEnd] = useState('');
  const [autoStart, setAutoStart] = useState(true), [autoEnd, setAutoEnd] = useState(true);
  const [interval, setInterval] = useState(dates ? '7' : String(Math.max(1, Math.round((hi - lo) / 10)) || 10));
  const [units, setUnits] = useState<Set<string>>(new Set(dates ? ['months'] : []));
  if (!ref || !cache || !cf) return missing(close);
  const ok = () => {
    const s = autoStart ? undefined : Number(start), e = autoEnd ? undefined : Number(end);
    if ((s !== undefined && Number.isNaN(s)) || (e !== undefined && Number.isNaN(e))) { store.notify(T('Enter numbers for the start and end.')); return false; }
    if (dates) {
      if (!units.size) { store.notify(T('Pick at least one unit.')); return false; }
      const toSerial = (t: string | undefined) => { if (t === undefined) return undefined; const d = new Date(t); return Number.isNaN(d.getTime()) ? Number(t) : Math.floor(d.getTime() / 86400000) + 25569; };
      groupRange(store, ref, f, [...units].map((u) => ({ by: u as typeof UNITS[number], start: toSerial(autoStart ? undefined : start), end: toSerial(autoEnd ? undefined : end), interval: u === 'days' && units.size === 1 ? Math.max(1, +interval || 1) : 1 })));
    } else {
      const step = Number(interval);
      if (!(step > 0)) { store.notify(T('Enter a step bigger than 0.')); return false; }
      groupRange(store, ref, f, [{ by: 'range', start: s, end: e, interval: step }]);
    }
    return true;
  };
  return (
    <Dialog title={T('Group')} close={close} width={380} onOk={ok}>
      <div className="grid2">
        <Field label={T('Starting at')}><input value={autoStart ? (dates ? '' : String(lo)) : start} disabled={autoStart} placeholder={dates ? T('YYYY-MM-DD') : String(lo)} onChange={(e) => setStart(e.target.value)} /></Field>
        <Field label={T('Ending at')}><input value={autoEnd ? (dates ? '' : String(hi)) : end} disabled={autoEnd} placeholder={dates ? T('YYYY-MM-DD') : String(hi)} onChange={(e) => setEnd(e.target.value)} /></Field>
      </div>
      <div className="grid2">
        <Check label={T('Automatic start')} checked={autoStart} onChange={setAutoStart} />
        <Check label={T('Automatic end')} checked={autoEnd} onChange={setAutoEnd} />
      </div>
      {dates ? (
        <>
          <div className="seclabel">{T('By')}</div>
          <div className="checks pvlist">
            {[...UNITS].reverse().map((u) => <label key={u}><input type="checkbox" checked={units.has(u)} onChange={(e) => { const n = new Set(units); if (e.target.checked) n.add(u); else n.delete(u); setUnits(n); }} /> {u[0]!.toUpperCase() + u.slice(1)}</label>)}
          </div>
          {units.size === 1 && units.has('days') && <Field label={T('Number of days')}><input type="number" min={1} value={interval} onChange={(e) => setInterval(e.target.value)} /></Field>}
        </>
      ) : <Field label={T('By (step)')}><input type="number" value={interval} onChange={(e) => setInterval(e.target.value)} /></Field>}
    </Dialog>
  );
};

// ---------------------------------------------------------------- calculated fields

export const CalcFieldDialog = ({ store, close, data }: DialogProps<Target & { field?: number }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const calcs = (cache?.fields ?? []).map((f, i) => ({ f, i })).filter(({ f }) => !!f.formula);
  const [pick, setPick] = useState<number | null>(data?.field ?? null);
  const cur = pick !== null ? cache?.fields[pick] : undefined;
  const [name, setName] = useState(cur?.name ?? `Field${calcs.length + 1}`);
  const [formula, setFormula] = useState(cur?.formula ? `=${cur.formula}` : '=0');
  if (!ref || !cache) return missing(close);
  const quoted = (n: string) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  const ok = () => {
    const n = name.trim();
    if (!n) { store.notify(T('Enter a name.')); return false; }
    if (cache.fields.some((f, i) => i !== pick && f.name.toLowerCase() === n.toLowerCase())) { store.notify(T('A field with that name already exists.')); return false; }
    const body = formula.trim().replace(/^=/, '');
    try { parsePivotFormula(body); } catch (e) { store.notify(T('There\'s a problem with this formula. {message}', { message: (e as Error).message })); return false; }
    saveCalcField(store, ref, n, body, pick);
    return true;
  };
  return (
    <Dialog title={T('Calculated field')} close={close} width={460} onOk={ok} okLabel={pick !== null ? T('Save') : T('Add')}
      footer={pick !== null ? <button onClick={() => { deleteCalcField(store, ref, pick); close(); }}>{T('Delete')}</button> : undefined}>
      {calcs.length > 0 && (
        <Field label={T('Calculated fields')}>
          <select value={pick ?? ''} onChange={(e) => { const v = e.target.value === '' ? null : +e.target.value; setPick(v); const f = v !== null ? cache.fields[v] : undefined; setName(f?.name ?? `Field${calcs.length + 1}`); setFormula(f?.formula ? `=${f.formula}` : '=0'); }}>
            <option value="">{T('New field')}</option>{calcs.map(({ f, i }) => <option key={i} value={i}>{f.name}</option>)}
          </select>
        </Field>
      )}
      <Field label={T('Name')}><input value={name} onChange={(e) => setName(e.target.value)} /></Field>
      <Field label={T('Formula (uses the sum of each field)')}><input value={formula} onChange={(e) => setFormula(e.target.value)} spellCheck={false} /></Field>
      <div className="seclabel">{T('Fields (click to insert)')}</div>
      <div className="pvchips">
        {cache.fields.map((f, i) => (f.database === false && !f.formula ? null : <button key={i} className="sm" onClick={() => setFormula((t) => `${t}${/[=+\-*/^(,&]$/.test(t.trim()) ? '' : '+'}${quoted(f.name)}`)}>{f.name}</button>))}
      </div>
    </Dialog>
  );
};

// ---------------------------------------------------------------- calculated items

export const CalcItemDialog = ({ store, close, data }: DialogProps<Target & { field?: number }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const axis = ref ? [...ref.pt.rows, ...ref.pt.cols].filter((f) => f >= 0 && !cache?.fields[f]?.group) : [];
  const [field, setFieldIdx] = useState<number>(data?.field !== undefined && axis.includes(data.field) ? data.field : axis[0] ?? -1);
  const norm = useMemo(() => (ref && cache ? normalizeFields(ref.pt, cache) : []), [ref, cache]);
  const items = norm[field]?.items ?? [];
  const calcSlots = items.map((item, i) => ({ it: item, i })).filter((c) => c.it.x < 0);
  const [slot, setSlot] = useState<number | null>(null);
  const [name, setName] = useState(`Formula${calcSlots.length + 1}`);
  const [formula, setFormula] = useState('=0');
  if (!ref || !cache) return missing(close);
  if (field < 0) return <Dialog title={T('Calculated item')} close={close}><p>{T('Add a row or column field (not a grouped one) to the PivotTable first.')}</p></Dialog>;
  const caption = (i: number) => { const it = items[i]!; return it.name ?? (it.x >= 0 ? String(fieldItems(cache.fields[field]!)[it.x] ?? '(blank)') : ''); };
  const quoted = (n: string) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  const ok = () => {
    const n = name.trim();
    if (!n) { store.notify(T('Enter a name.')); return false; }
    if (items.some((_, i) => i !== slot && caption(i).toLowerCase() === n.toLowerCase())) { store.notify(T('An item with that name already exists.')); return false; }
    const body = formula.trim().replace(/^=/, '');
    try { parsePivotFormula(body); } catch (e) { store.notify(T('There\'s a problem with this formula. {message}', { message: (e as Error).message })); return false; }
    saveCalcItem(store, ref, field, n, body, slot);
    return true;
  };
  return (
    <Dialog title={T('Calculated item')} close={close} width={460} onOk={ok} okLabel={slot !== null ? T('Save') : T('Add')}
      footer={slot !== null ? <button onClick={() => { deleteCalcItem(store, ref, field, slot); close(); }}>{T('Delete')}</button> : undefined}>
      <div className="grid2">
        <Field label={T('Field')}><select value={field} onChange={(e) => { setFieldIdx(+e.target.value); setSlot(null); }}>{axis.map((f) => <option key={f} value={f}>{fieldCaption(ref.pt, cache, f)}</option>)}</select></Field>
        <Field label={T('Item')}>
          <select value={slot ?? ''} onChange={(e) => { const v = e.target.value === '' ? null : +e.target.value; setSlot(v); if (v !== null) { setName(items[v]!.name ?? ''); setFormula(`=${items[v]!.formula ?? '0'}`); } else { setName(`Formula${calcSlots.length + 1}`); setFormula('=0'); } }}>
            <option value="">{T('New item')}</option>{calcSlots.map(({ it, i }) => <option key={i} value={i}>{it.name}</option>)}
          </select>
        </Field>
      </div>
      <Field label={T('Name')}><input value={name} onChange={(e) => setName(e.target.value)} /></Field>
      <Field label={T('Formula (items of this field, e.g. East+West)')}><input value={formula} onChange={(e) => setFormula(e.target.value)} spellCheck={false} /></Field>
      <div className="seclabel">{T('Items (click to insert)')}</div>
      <div className="pvchips">
        {items.map((it, i) => (it.x < 0 ? null : <button key={i} className="sm" onClick={() => setFormula((t) => `${t}${/[=+\-*/^(,&]$/.test(t.trim()) ? '' : '+'}${quoted(caption(i))}`)}>{caption(i)}</button>))}
      </div>
    </Dialog>
  );
};

// ---------------------------------------------------------------- label and value filters

const LABEL_OPS: [LabelOp, string][] = [
  ['equal', 'equals'], ['notEqual', 'does not equal'], ['beginsWith', 'begins with'], ['notBeginsWith', 'does not begin with'], ['endsWith', 'ends with'],
  ['notEndsWith', 'does not end with'], ['contains', 'contains'], ['notContains', 'does not contain'], ['greaterThan', 'is greater than'],
  ['greaterThanOrEqual', 'is greater than or equal to'], ['lessThan', 'is less than'], ['lessThanOrEqual', 'is less than or equal to'], ['between', 'is between'], ['notBetween', 'is not between'],
];
const VALUE_OPS: [ValueOp | 'top', string][] = [
  ['equal', 'equals'], ['notEqual', 'does not equal'], ['greaterThan', 'is greater than'], ['greaterThanOrEqual', 'is greater than or equal to'],
  ['lessThan', 'is less than'], ['lessThanOrEqual', 'is less than or equal to'], ['between', 'is between'], ['notBetween', 'is not between'], ['top', 'top or bottom'],
];

export const FilterDialog = ({ store, close, data }: DialogProps<Target & { field: number; kind: 'label' | 'value' }>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const f = data?.field ?? -1;
  const existing = ref?.pt.fields[f]?.filters?.find((x) => (data?.kind === 'label' ? x.kind === 'label' : x.kind === 'value' || x.kind === 'top'));
  const [op, setOp] = useState<string>(existing ? (existing.kind === 'top' ? 'top' : existing.kind === 'dates' ? 'between' : existing.op) : data?.kind === 'label' ? 'contains' : 'greaterThan');
  const [v1, setV1] = useState(existing && existing.kind !== 'dates' ? String(existing.kind === 'top' ? existing.value : existing.value1) : data?.kind === 'value' ? '0' : '');
  const [v2, setV2] = useState(existing && (existing.kind === 'label' || existing.kind === 'value') ? String(existing.value2 ?? '') : '');
  const [dataField, setDataField] = useState(existing && (existing.kind === 'value' || existing.kind === 'top') ? existing.dataField : 0);
  const [top, setTop] = useState(existing?.kind === 'top' ? existing.top : true);
  const [by, setBy] = useState<'count' | 'percent' | 'sum'>(existing?.kind === 'top' ? existing.by : 'count');
  if (!ref || !cache || !data) return missing(close);
  const between = op === 'between' || op === 'notBetween';
  const ok = () => {
    let flt: PivotFilter;
    if (data.kind === 'label') flt = { kind: 'label', op: op as LabelOp, value1: v1, value2: between ? v2 : undefined };
    else if (op === 'top') {
      const n = Number(v1);
      if (!(n > 0)) { store.notify(T('Enter a number bigger than 0.')); return false; }
      flt = { kind: 'top', top, by, value: n, dataField };
    } else {
      const a = Number(v1), b = Number(v2);
      if (Number.isNaN(a) || (between && Number.isNaN(b))) { store.notify(T('Enter numbers.')); return false; }
      flt = { kind: 'value', op: op as ValueOp, dataField, value1: a, value2: between ? b : undefined };
    }
    update(store, ref, (pt) => {
      const others = (pt.fields[f]?.filters ?? []).filter((x) => (data.kind === 'label' ? x.kind !== 'label' : x.kind === 'label' || x.kind === 'dates'));
      return setField(pt, f, { filters: [...others, flt] });
    }, N('Filter'));
    return true;
  };
  return (
    <Dialog title={T(data.kind === 'label' ? 'Label filter ({field})' : 'Value filter ({field})', { field: fieldCaption(ref.pt, cache, f) })} close={close} width={460} onOk={ok}>
      {data.kind === 'value' && (
        <Field label={T('Value field')}><select value={dataField} onChange={(e) => setDataField(+e.target.value)}>{ref.pt.data.map((d, i) => <option key={i} value={i}>{dataFieldName(cache, d)}</option>)}</select></Field>
      )}
      <Field label={T('Show items whose value')}>
        <select value={op} onChange={(e) => setOp(e.target.value)}>{(data.kind === 'label' ? LABEL_OPS : VALUE_OPS).map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
      </Field>
      {op === 'top' ? (
        <div className="grid2">
          <Field label={T('Show')}><select value={top ? 'top' : 'bottom'} onChange={(e) => setTop(e.target.value === 'top')}><option value="top">{T('Top')}</option><option value="bottom">{T('Bottom')}</option></select></Field>
          <Field label={T('Amount')}><input type="number" min={1} value={v1} onChange={(e) => setV1(e.target.value)} /></Field>
          <Field label={T('Of')}><select value={by} onChange={(e) => setBy(e.target.value as 'count')}><option value="count">{T('Items')}</option><option value="percent">{T('Percent')}</option><option value="sum">{T('Sum')}</option></select></Field>
        </div>
      ) : (
        <div className="grid2">
          <Field label={between ? T('From') : T('Value')}><input value={v1} onChange={(e) => setV1(e.target.value)} /></Field>
          {between && <Field label={T('To')}><input value={v2} onChange={(e) => setV2(e.target.value)} /></Field>}
        </div>
      )}
      {data.kind === 'label' && <p className="muted small">{T('Use ? for one character and * for any run of characters.')}</p>}
    </Dialog>
  );
};

// ---------------------------------------------------------------- source, pages, move

export const ChangeSourceDialog = ({ store, close, data }: DialogProps<Target>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const [text, setText] = useState(cache ? sourceText(cache.source) : '');
  if (!ref || !cache) return missing(close);
  const ok = () => {
    const t = text.trim().replace(/^=/, '');
    const r = resolveRange(store, t);
    const src = r ? { kind: 'range' as const, sheet: store.wb.sheets[r.sheet]!.name, range: r.range } : { kind: 'name' as const, name: t };
    const err = changeSource(store, cache.id, src);
    if (err !== null) { if (err) store.notify(err); return false; }
    return true;
  };
  const sharers = store.wb.sheets.flatMap((s) => (s.pivots ?? []).filter((p) => p.cacheId === cache.id)).length;
  return (
    <Dialog title={T('Change data source')} close={close} width={440} onOk={ok}>
      <Field label={T('Range, table or name')}><input value={text} onChange={(e) => setText(e.target.value)} spellCheck={false} /></Field>
      {sharers > 1 && <p className="muted small">{T('{count} PivotTables share this data and will all change.', { count: sharers })}</p>}
    </Dialog>
  );
};

export const ShowPagesDialog = ({ store, close, data }: DialogProps<Target>) => {
  const ref = refOf(store, data);
  const cache = ref ? cacheOf(store.wb, ref.pt.cacheId) : undefined;
  const [field, setFieldIdx] = useState(ref?.pt.pages[0]?.field ?? -1);
  if (!ref || !cache) return missing(close);
  if (!ref.pt.pages.length) return <Dialog title={T('Show report filter pages')} close={close}><p>{T('Add a field to Filters first.')}</p></Dialog>;
  return (
    <Dialog title={T('Show report filter pages')} close={close} width={360} onOk={() => { const n = showFilterPages(store, ref, field); store.notify(T('Added {n, plural, one {# sheet} other {# sheets}}.', { n })); return true; }}>
      <p className="muted small">{T('Adds one sheet per item, each with this PivotTable filtered to that item.')}</p>
      <div className="listbox">
        {ref.pt.pages.map((p) => <label key={p.field}><input type="radio" name="pvpages" checked={field === p.field} onChange={() => setFieldIdx(p.field)} /> {fieldCaption(ref.pt, cache, p.field)}</label>)}
      </div>
    </Dialog>
  );
};

export const MovePivotDialog = ({ store, close, data }: DialogProps<Target>) => {
  const ref = refOf(store, data);
  const [text, setText] = useState(ref ? `${quoteSheet(store.wb.sheets[ref.sheet]!.name)}!${rangeAddr({ r1: ref.pt.anchor.row, c1: ref.pt.anchor.col, r2: ref.pt.anchor.row, c2: ref.pt.anchor.col })}` : '');
  if (!ref) return missing(close);
  const ok = () => {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text.trim());
    const sheet = m ? store.engine.sheetIndex((m[1] ?? m[2] ?? '').replace(/''/g, "'")) : store.sheetIndex;
    const a = parseAddr((m ? m[3]! : text).replace(/\$/g, '').trim());
    if (sheet < 0 || !a) { store.notify(T('Enter a cell such as Sheet2!A3.')); return false; }
    return movePivot(store, ref.sheet, ref.pt.name, { sheet, row: a.row, col: a.col });
  };
  return (
    <Dialog title={T('Move PivotTable')} close={close} width={360} onOk={ok} okLabel={T('Move')}>
      <Field label={T('New top-left cell')}><input value={text} onChange={(e) => setText(e.target.value)} spellCheck={false} /></Field>
    </Dialog>
  );
};

export { VALUES_FIELD };
