// Select data, move chart and save-as-template dialogs.
import { useState } from 'react';
import { ArrowDown, ArrowUp, Plus, Trash2 } from 'lucide-react';
import type { GridStore } from '../../app/store';
import type { Chart, ChartOptions, ChartSeries } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from '../dialogs/Dialog';
import { parseRef, resolveChart } from './chartData';
import { traits } from './chartTypes';
import { editChart, moveToChartSheet, moveToSheet, saveTemplate, switchRowCol } from './chartOps';
import { parseNumberList } from './trend';
import { N, T } from '../i18n';

const chartOf = (store: GridStore, id?: string): Chart | undefined => (id ? store.sheet.charts.find((c) => c.id === id) : undefined);

/** A range reference, an array constant or nothing (for optional parts). */
const validRef = (s: string, optional: boolean): boolean => {
  const t = s.trim();
  if (!t) return optional;
  return !!parseRef(t) || !!parseNumberList(t);
};

const withSheet = (store: GridStore, s: string): string => {
  const t = s.trim().replace(/^=/, '');
  const p = parseRef(t);
  if (!p || p.sheet !== undefined) return t;
  const q = /^[A-Za-z_][\w.]*$/.test(store.sheet.name) ? store.sheet.name : `'${store.sheet.name.replace(/'/g, "''")}'`;
  return `${q}!${t}`;
};

export const SelectDataDialog = ({ store, data, close }: DialogProps<{ id?: string }>) => {
  const chart = chartOf(store, data?.id);
  const [rows, setRows] = useState<ChartSeries[]>(() => structuredClone(chart?.series ?? []));
  const [sel, setSel] = useState(0);
  const [cats, setCats] = useState(chart?.series[0]?.categories ?? '');
  const [blanks, setBlanks] = useState<NonNullable<ChartOptions['blanks']>>(chart?.options?.blanks ?? 'gap');
  const [hidden, setHidden] = useState(!!chart?.options?.showHidden);
  if (!chart) { queueMicrotask(close); return null; }
  const tr = traits(chart.type);
  const xy = tr.family === 'scatter' || tr.family === 'bubble';
  const resolved = resolveChart(store, { ...chart, series: rows.filter((r) => validRef(r.values, false)) });
  const cur = rows[sel];
  const up = (patch: Partial<ChartSeries>) => setRows((r) => r.map((x, i) => (i === sel ? { ...x, ...patch } : x)));
  const move = (d: number) => {
    const j = sel + d;
    if (j < 0 || j >= rows.length) return;
    setRows((r) => { const n = [...r]; [n[sel], n[j]] = [n[j]!, n[sel]!]; return n; });
    setSel(j);
  };
  const ok = (): boolean => {
    if (!rows.length) { store.notify(T('A chart needs at least one series.')); return false; }
    const bad = rows.findIndex((r) => !validRef(r.values, false) || (r.sizes !== undefined && !validRef(r.sizes, true)) || (xy && !validRef(r.categories ?? '', true)));
    if (bad >= 0) { setSel(bad); store.notify(T('Enter the values as a range such as Sheet1!B2:B9.')); return false; }
    if (!xy && !validRef(cats, true)) { store.notify(T('Enter the category labels as a range such as Sheet1!A2:A9.')); return false; }
    editChart(store, chart.id, N('Select data'), (c) => {
      c.series = rows.map((r) => {
        const s: ChartSeries = { ...r, values: withSheet(store, r.values) };
        if (r.name !== undefined && parseRef(r.name.replace(/^=/, ''))) s.name = withSheet(store, r.name);
        if (!xy) { if (cats.trim()) s.categories = withSheet(store, cats); else delete s.categories; }
        else if (r.categories?.trim()) s.categories = withSheet(store, r.categories); else delete s.categories;
        if (r.sizes?.trim()) s.sizes = withSheet(store, r.sizes); else delete s.sizes;
        if (!s.name) delete s.name;
        return s;
      });
      c.options = { ...c.options, blanks: blanks === 'gap' ? undefined : blanks, showHidden: hidden || undefined };
    });
    return true;
  };
  return (
    <Dialog title={T('Select data')} close={close} onOk={ok} okLabel={T('Save')} width={560}>
      <div className="seldata">
        <div className="serlist">
          <div className="mini">{T('Series')}</div>
          <div className="listbox" role="listbox" aria-label={T('Series')}>
            {rows.map((r, i) => (
              <button key={i} role="option" aria-selected={i === sel} className={`listitem${i === sel ? ' on' : ''}`} onClick={() => setSel(i)}>
                {resolved.series.find((s) => s.src === r || s.index === i)?.name ?? (r.name || T('Series {p0}', { p0: i + 1 }))}
              </button>
            ))}
          </div>
          <div className="iconrow">
            <button className="ico sm" title={T('Add series')} aria-label={T('Add series')} onClick={() => { setRows((r) => [...r, { values: '' }]); setSel(rows.length); }}><Plus size={15} /></button>
            <button className="ico sm" title={T('Remove series')} aria-label={T('Remove series')} disabled={!cur} onClick={() => { setRows((r) => r.filter((_, i) => i !== sel)); setSel(Math.max(0, sel - 1)); }}><Trash2 size={15} /></button>
            <button className="ico sm" title={T('Move up')} aria-label={T('Move up')} disabled={sel === 0} onClick={() => move(-1)}><ArrowUp size={15} /></button>
            <button className="ico sm" title={T('Move down')} aria-label={T('Move down')} disabled={sel >= rows.length - 1} onClick={() => move(1)}><ArrowDown size={15} /></button>
          </div>
        </div>
        <div className="serform">
          {cur ? <>
            <Field label={T('Name (text or a cell like Sheet1!B1)')}><input value={cur.name ?? ''} onChange={(e) => up({ name: e.target.value || undefined })} spellCheck={false} /></Field>
            <Field label={xy ? T('Y values') : T('Values')}><input value={cur.values} aria-invalid={!validRef(cur.values, false)} onChange={(e) => up({ values: e.target.value })} placeholder={T('Sheet1!B2:B9')} spellCheck={false} /></Field>
            {xy && <Field label={T('X values')}><input value={cur.categories ?? ''} aria-invalid={!validRef(cur.categories ?? '', true)} onChange={(e) => up({ categories: e.target.value || undefined })} placeholder={T('Sheet1!A2:A9')} spellCheck={false} /></Field>}
            {tr.family === 'bubble' && <Field label={T('Bubble sizes')}><input value={cur.sizes ?? ''} aria-invalid={!validRef(cur.sizes ?? '', true)} onChange={(e) => up({ sizes: e.target.value || undefined })} placeholder={T('Sheet1!C2:C9')} spellCheck={false} /></Field>}
          </> : <p className="muted">{T('Add a series to chart.')}</p>}
          {!xy && <Field label={T('Category labels (all series)')}><input value={cats} aria-invalid={!validRef(cats, true)} onChange={(e) => setCats(e.target.value)} placeholder={T('Sheet1!A2:A9')} spellCheck={false} /></Field>}
          <Field label={T('Show empty cells as')}>
            <select value={blanks} onChange={(e) => setBlanks(e.target.value as typeof blanks)}>
              <option value="gap">{T('Gaps')}</option><option value="zero">{T('Zero')}</option><option value="span">{T('Join points with a line')}</option>
            </select>
          </Field>
          <label className="checkrow"><input type="checkbox" checked={hidden} onChange={(e) => setHidden(e.target.checked)} />{' '}{T('Show data in hidden rows and columns')}</label>
          {!tr.combo && <button className="tool small" onClick={() => { if (switchRowCol(store, chart.id)) close(); }}>{T('Switch rows and columns')}</button>}
        </div>
      </div>
    </Dialog>
  );
};

export const MoveChartDialog = ({ store, data, close }: DialogProps<{ id?: string }>) => {
  const chart = chartOf(store, data?.id);
  const onChartSheet = store.sheet.kind === 'chart';
  const targets = store.wb.sheets.map((s, i) => ({ s, i })).filter(({ s, i }) => s.kind !== 'chart' && i !== store.sheetIndex);
  const [mode, setMode] = useState<'new' | 'object'>(onChartSheet ? 'object' : 'new');
  const [name, setName] = useState(chart?.title && !chart.title.startsWith('=') ? chart.title.slice(0, 31) : T('Chart'));
  const [target, setTarget] = useState(targets[0]?.i ?? -1);
  if (!chart) { queueMicrotask(close); return null; }
  const ok = (): boolean => {
    if (mode === 'new') {
      if (/[\\/?*[\]:]/.test(name)) { store.notify(T('Sheet names can\'t contain \\ / ? * [ ] or :')); return false; }
      return moveToChartSheet(store, chart.id, name.slice(0, 31));
    }
    if (target < 0) { store.notify(T('There\'s no other worksheet to move the chart to.')); return false; }
    return moveToSheet(store, chart.id, target);
  };
  return (
    <Dialog title={T('Move chart')} close={close} onOk={ok} okLabel={T('Move')}>
      {!onChartSheet && <label className="checkrow"><input type="radio" name="mv" checked={mode === 'new'} onChange={() => setMode('new')} />{' '}{T('To a new chart sheet')}</label>}
      {mode === 'new' && <Field label={T('Sheet name')}><input value={name} maxLength={31} onChange={(e) => setName(e.target.value)} /></Field>}
      <label className="checkrow"><input type="radio" name="mv" checked={mode === 'object'} disabled={!targets.length} onChange={() => setMode('object')} />{' '}{T('Onto a worksheet')}</label>
      {mode === 'object' && <Field label={T('Worksheet')}>
        <select value={target} onChange={(e) => setTarget(+e.target.value)}>{targets.map(({ s, i }) => <option key={i} value={i}>{s.name}</option>)}</select>
      </Field>}
    </Dialog>
  );
};

export const SaveTemplateDialog = ({ store, data, close }: DialogProps<{ id?: string }>) => {
  const chart = chartOf(store, data?.id);
  const [name, setName] = useState(chart?.title && !chart.title.startsWith('=') ? T('{name} style', { name: chart.title }) : T('My chart'));
  if (!chart) { queueMicrotask(close); return null; }
  return (
    <Dialog title={T('Save as template')} close={close} onOk={() => { saveTemplate(store, chart, name); return true; }} okLabel={T('Save')}>
      <Field label={T('Template name')}><input value={name} maxLength={80} onChange={(e) => setName(e.target.value)} /></Field>
      <p className="muted small">{T('Templates keep the chart type, colours, axes and formatting, and are kept in this browser. Use one from Insert chart, under Templates.')}</p>
    </Dialog>
  );
};
