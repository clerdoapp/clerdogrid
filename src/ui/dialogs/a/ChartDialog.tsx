import { useEffect, useMemo, useRef, useState } from 'react';
import type { GridStore } from '../../../app/store';
import { currentRegion } from '../../../app/ops/data';
import type { Chart, ChartType } from '../../../model/workbook';
import { buildSeries, chartOrientation, chartSourceRange, detectHeaders, recommendTypes, resolveChart, type Headers, type Orientation } from '../../charts/chartData';
import { ChartSvg } from '../../charts/ChartSvg';
import { selectChart } from '../../charts/ChartView';
import '../../charts/charts.css';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { clipToUsed, rangeText, readMatrix, resolveRange, type SheetRange } from './util';

export const CHART_TYPES: { type: ChartType; label: string }[] = [
  { type: 'column', label: 'Column' }, { type: 'stackedColumn', label: 'Stacked column' }, { type: 'percentColumn', label: '100% column' },
  { type: 'bar', label: 'Bar' }, { type: 'stackedBar', label: 'Stacked bar' }, { type: 'line', label: 'Line' },
  { type: 'area', label: 'Area' }, { type: 'stackedArea', label: 'Stacked area' }, { type: 'pie', label: 'Pie' },
  { type: 'doughnut', label: 'Doughnut' }, { type: 'scatter', label: 'Scatter' }, { type: 'radar', label: 'Radar' },
  { type: 'combo', label: 'Column and line' },
];
const labelOf = (t: ChartType) => CHART_TYPES.find((x) => x.type === t)?.label ?? t;

const DEFAULT_W = 480, DEFAULT_H = 288;
const newId = () => `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/** The selection, or the block of data around the active cell when one cell is selected. */
const sourceFromSelection = (store: GridStore): SheetRange => {
  const r = store.range;
  const single = r.r1 === r.r2 && r.c1 === r.c2;
  return { sheet: store.sheetIndex, name: store.sheet.name, range: clipToUsed(store, store.sheetIndex, single ? currentRegion(store) : r) };
};

interface Draft { type: ChartType; source: SheetRange; orientation?: Orientation; headers?: Headers; title?: string; legend?: Chart['legend']; axisTitles?: Chart['axisTitles']; dataLabels?: boolean }

const draftChart = (store: GridStore, d: Draft, base?: Chart): Chart | null => {
  const m = readMatrix(store, d.source.sheet, d.source.range, 2000);
  const series = buildSeries(d.source.name, d.source.range, m, { orientation: d.orientation, headers: d.headers, type: d.type });
  if (!series.length) return null;
  return {
    id: base?.id ?? newId(), type: d.type, series, title: d.title,
    anchor: base?.anchor ?? { row: d.source.range.r1, col: d.source.range.c2 + 1, dx: 12, dy: 0, width: DEFAULT_W, height: DEFAULT_H },
    legend: d.legend, axisTitles: d.axisTitles, dataLabels: d.dataLabels,
  };
};

const defaultTitle = (store: GridStore, chart: Chart): string | undefined => {
  if (chart.series.length !== 1 || !chart.series[0]!.name?.includes('!')) return undefined;
  return resolveChart(store, chart).series[0]?.name;
};

const hasNumbers = (store: GridStore, chart: Chart): boolean => resolveChart(store, chart).series.some((s) => s.values.some((v) => v !== null));

const insert = (store: GridStore, chart: Chart): void => {
  store.change('Insert chart', 'sheet', () => { store.sheet.charts.push(chart); });
  selectChart(store, chart.id);
};

/** Inserts a chart from the selection straight away. Returns false with a message if there's nothing to chart. */
export const quickChart = (store: GridStore, type: ChartType): boolean => {
  const source = sourceFromSelection(store);
  const chart = draftChart(store, { type, source });
  if (!chart || !hasNumbers(store, chart)) { store.notify('Select a block of cells with numbers to make a chart.'); return false; }
  chart.title = defaultTitle(store, chart);
  insert(store, chart);
  return true;
};

export const ChartDialog = ({ store, data, close }: DialogProps<{ type?: ChartType; quick?: boolean; id?: string }>) => {
  const editing = data?.id ? store.sheet.charts.find((c) => c.id === data.id) : undefined;
  const ran = useRef(false);
  useEffect(() => {
    if (!data?.quick || ran.current) return;
    ran.current = true;
    quickChart(store, data.type ?? 'column');
    close();
  }, [data, store, close]);
  if (data?.quick) return null;
  if (data?.id && !editing) { queueMicrotask(close); return null; }
  return <ChartEditor store={store} close={close} editing={editing} initialType={data?.type} />;
};

const ChartEditor = ({ store, close, editing, initialType }: { store: GridStore; close: () => void; editing?: Chart; initialType?: ChartType }) => {
  const initial = useMemo(() => {
    if (editing) {
      const src = chartSourceRange(editing);
      const si = src?.sheet ? store.wb.sheets.findIndex((s) => s.name.toLowerCase() === src.sheet!.toLowerCase()) : store.sheetIndex;
      const sheet = si < 0 ? store.sheetIndex : si;
      return { source: src ? { sheet, name: store.wb.sheets[sheet]!.name, range: src.range } : sourceFromSelection(store), orientation: chartOrientation(editing) };
    }
    const source = sourceFromSelection(store);
    const m = readMatrix(store, source.sheet, source.range, 2000);
    return { source, orientation: undefined as Orientation | undefined, recs: recommendTypes(m) };
  }, [editing, store]);

  const [rangeInput, setRangeInput] = useState(rangeText(store, initial.source.range, initial.source.sheet));
  const parsed = resolveRange(store, rangeInput);
  const source = parsed ?? initial.source;
  const matrix = useMemo(() => readMatrix(store, source.sheet, source.range, 2000), [store, source.sheet, source.range.r1, source.range.c1, source.range.r2, source.range.c2]); // eslint-disable-line react-hooks/exhaustive-deps
  const detected = useMemo(() => detectHeaders(matrix), [matrix]);
  const recs = useMemo(() => ('recs' in initial && initial.recs ? initial.recs : recommendTypes(matrix)), [initial, matrix]);

  const [type, setType] = useState<ChartType>(editing?.type ?? initialType ?? recs[0] ?? 'column');
  const [orientation, setOrientation] = useState<Orientation | undefined>(initial.orientation);
  const [headers, setHeaders] = useState<Headers | undefined>(undefined);
  const [title, setTitle] = useState(editing?.title ?? '');
  const [titleTouched, setTitleTouched] = useState(!!editing);
  const [xTitle, setXTitle] = useState(editing?.axisTitles?.x ?? '');
  const [yTitle, setYTitle] = useState(editing?.axisTitles?.y ?? '');
  const [legend, setLegend] = useState<NonNullable<Chart['legend']> | 'auto'>(editing?.legend ?? 'auto');
  const [labels, setLabels] = useState(!!editing?.dataLabels);
  const [tab, setTab] = useState<'rec' | 'all'>(editing ? 'all' : 'rec');
  // In edit mode keep the chart's own series (and colors) until the data settings change.
  const [dataChanged, setDataChanged] = useState(!editing);

  const h = headers ?? detected;
  const settings = { title: title.trim() || undefined, legend: legend === 'auto' ? undefined : legend, axisTitles: xTitle || yTitle ? { x: xTitle || undefined, y: yTitle || undefined } : undefined, dataLabels: labels || undefined };
  const preview: Chart | null = useMemo(() => {
    if (editing && !dataChanged) return { ...editing, type, ...settings };
    return draftChart(store, { type, source, orientation, headers: h, ...settings }, editing);
  }, [editing, dataChanged, type, source, orientation, h.firstRow, h.firstCol, settings.title, settings.legend, xTitle, yTitle, labels, store]); // eslint-disable-line react-hooks/exhaustive-deps
  const autoTitle = preview && !titleTouched ? defaultTitle(store, preview) : undefined;
  const shown = preview ? { ...preview, title: settings.title ?? autoTitle } : null;
  const resolved = useMemo(() => (shown ? resolveChart(store, shown, store.sheetIndex) : null), [shown, store]);

  const thumbs = (types: ChartType[]) => (
    <div className="charttypes" role="listbox" aria-label="Chart types">
      {types.map((t) => {
        const c = draftChart(store, { type: t, source, orientation, headers: h }) ?? (editing ? { ...editing, type: t } : null);
        return (
          <button key={t} role="option" aria-selected={type === t} className={`charttype${type === t ? ' on' : ''}`} onClick={() => setType(t)}>
            <div className="thumb">{c && <ChartSvg chart={c} data={resolveChart(store, c)} width={120} height={64} thumb />}</div>
            {labelOf(t)}
          </button>
        );
      })}
    </div>
  );

  const ok = (): boolean => {
    if (!parsed) { store.notify('Enter a data range such as A1:D10.'); return false; }
    if (!shown || !hasNumbers(store, shown)) { store.notify('The data range has no numbers to chart.'); return false; }
    const finalChart: Chart = { ...shown, title: shown.title?.trim() || undefined };
    if (finalChart.type !== 'combo') finalChart.series = finalChart.series.map(({ type: _t, ...s }) => s);
    else if (finalChart.series.every((s) => !s.type)) finalChart.series = finalChart.series.map((s, i) => ({ ...s, type: i === 0 ? 'column' : 'line' }));
    if (editing) {
      store.change('Edit chart', 'sheet', () => {
        const i = store.sheet.charts.findIndex((c) => c.id === editing.id);
        if (i >= 0) store.sheet.charts[i] = finalChart;
      });
    } else insert(store, finalChart);
    return true;
  };

  const markData = () => setDataChanged(true);

  return (
    <Dialog title={editing ? 'Edit chart' : 'Insert chart'} close={close} wide onOk={ok} okLabel={editing ? 'Save' : 'Insert'}>
      <div className="tabsrow">
        <button className={tab === 'rec' ? 'on' : ''} onClick={() => setTab('rec')}>Suggested</button>
        <button className={tab === 'all' ? 'on' : ''} onClick={() => setTab('all')}>All charts</button>
      </div>
      {thumbs(tab === 'rec' ? recs : CHART_TYPES.map((t) => t.type))}
      <div className="chartdlg">
        <div className="chartpreview">
          {shown && resolved ? <ChartSvg chart={shown} data={resolved} width={460} height={230} /> : <p className="muted pad">Choose a range with numbers to see a preview.</p>}
        </div>
        <div className="listbox">
          <Field label="Data range">
            <input value={rangeInput} onChange={(e) => { setRangeInput(e.target.value); markData(); }} placeholder="A1:D10" aria-invalid={!parsed} spellCheck={false} />
          </Field>
          <div className="checkrow">
            <span className="muted small">Series in</span>
            <div className="segmented" role="radiogroup" aria-label="Series in">
              {(['cols', 'rows'] as const).map((o) => {
                const on = (orientation ?? (editing && !dataChanged ? chartOrientation(editing) : undefined) ?? (matrix.length - (h.firstRow ? 1 : 0) >= (matrix[0]?.length ?? 0) - (h.firstCol ? 1 : 0) ? 'cols' : 'rows')) === o;
                return <button key={o} role="radio" aria-checked={on} className={on ? 'on' : ''} onClick={() => { setOrientation(o); markData(); }}>{o === 'cols' ? 'Columns' : 'Rows'}</button>;
              })}
            </div>
          </div>
          <label className="checkrow"><input type="checkbox" checked={h.firstRow} onChange={(e) => { setHeaders({ ...h, firstRow: e.target.checked }); markData(); }} /> First row has names</label>
          <label className="checkrow"><input type="checkbox" checked={h.firstCol} onChange={(e) => { setHeaders({ ...h, firstCol: e.target.checked }); markData(); }} /> First column has labels</label>
        </div>
      </div>
      <div className="grid2">
        <Field label="Title"><input value={title} placeholder={autoTitle ?? 'Chart title'} onChange={(e) => { setTitle(e.target.value); setTitleTouched(true); }} /></Field>
        <Field label="Legend">
          <select value={legend} onChange={(e) => setLegend(e.target.value as typeof legend)}>
            <option value="auto">Automatic</option><option value="bottom">Bottom</option><option value="top">Top</option>
            <option value="right">Right</option><option value="left">Left</option><option value="none">None</option>
          </select>
        </Field>
        {type !== 'pie' && type !== 'doughnut' && type !== 'radar' && <>
          <Field label="Horizontal axis title"><input value={xTitle} onChange={(e) => setXTitle(e.target.value)} /></Field>
          <Field label="Vertical axis title"><input value={yTitle} onChange={(e) => setYTitle(e.target.value)} /></Field>
        </>}
      </div>
      <label className="checkrow"><input type="checkbox" checked={labels} onChange={(e) => setLabels(e.target.checked)} /> Show values on the chart</label>
    </Dialog>
  );
};
