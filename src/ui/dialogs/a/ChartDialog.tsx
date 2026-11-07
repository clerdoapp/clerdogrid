import { useEffect, useMemo, useRef, useState } from 'react';
import type { GridStore } from '../../../app/store';
import { currentRegion } from '../../../app/ops/data';
import type { Chart, ChartOptions, ChartSeries, ChartType } from '../../../model/workbook';
import { buildSeries, chartOrientation, chartSourceRange, detectHeaders, recommendTypes, resolveChart, secondaryOf, seriesColor, STYLE_COUNT, STYLE_NAMES, type Headers, type Orientation } from '../../charts/chartData';
import { guessTotals } from '../../charts/chartStats';
import { CHART_GROUPS, CHART_TYPES, chartGroup, chartLabel, comboKind, hasAxes, supportsSecondary, traits, type ChartGroup } from '../../charts/chartTypes';
import { ChartSvg } from '../../charts/ChartSvg';
import { selectChart } from '../../charts/ChartView';
import '../../charts/charts.css';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { clipToUsed, rangeText, readMatrix, resolveRange, type SheetRange } from './util';

export { CHART_TYPES };

const DEFAULT_W = 480, DEFAULT_H = 288;
const newId = () => `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

/** The selection, or the block of data around the active cell when one cell is selected. */
const sourceFromSelection = (store: GridStore): SheetRange => {
  const r = store.range;
  const single = r.r1 === r.r2 && r.c1 === r.c2;
  return { sheet: store.sheetIndex, name: store.sheet.name, range: clipToUsed(store, store.sheetIndex, single ? currentRegion(store) : r) };
};

interface Draft { type: ChartType; source: SheetRange; orientation?: Orientation; headers?: Headers; title?: string; legend?: Chart['legend']; axisTitles?: Chart['axisTitles']; dataLabels?: boolean; options?: ChartOptions }

const draftChart = (store: GridStore, d: Draft, base?: Chart): Chart | null => {
  const m = readMatrix(store, d.source.sheet, d.source.range, 2000);
  const series = buildSeries(d.source.name, d.source.range, m, { orientation: d.orientation, headers: d.headers, type: d.type });
  if (!series.length) return null;
  return {
    id: base?.id ?? newId(), type: d.type, series, title: d.title,
    anchor: base?.anchor ?? { row: d.source.range.r1, col: d.source.range.c2 + 1, dx: 12, dy: 0, width: DEFAULT_W, height: DEFAULT_H },
    legend: d.legend, axisTitles: d.axisTitles, dataLabels: d.dataLabels, ...(d.options ? { options: d.options } : {}),
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

/** Series fields that only some chart types use are dropped or filled in for the final type. */
const normalizeSeries = (type: ChartType, series: ChartSeries[]): ChartSeries[] => {
  const tr = traits(type);
  return series.map((s, i) => {
    const { type: kind, sizes, ...rest } = s;
    const out: ChartSeries = { ...rest };
    if (tr.combo) out.type = kind ?? comboKind(type, i, series.length);
    if (tr.family === 'bubble' && sizes) out.sizes = sizes;
    return out;
  });
};

/** Inserts a chart from the selection straight away. Returns false with a message if there's nothing to chart. */
export const quickChart = (store: GridStore, type: ChartType): boolean => {
  const source = sourceFromSelection(store);
  const chart = draftChart(store, { type, source });
  if (!chart || !hasNumbers(store, chart)) { store.notify('Select a block of cells with numbers to make a chart.'); return false; }
  chart.series = normalizeSeries(type, chart.series);
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

const NUM_FORMATS: { code: string; label: string }[] = [
  { code: '', label: 'Same as data' }, { code: 'General', label: 'General' }, { code: '#,##0', label: '1,234' }, { code: '#,##0.00', label: '1,234.56' },
  { code: '0%', label: '12%' }, { code: '0.0%', label: '12.3%' }, { code: '"$"#,##0', label: '$1,234' }, { code: '0.00E+00', label: '1.23E+03' },
];

/** Treemap, sunburst and funnel show their labels unless turned off. */
const labelsOnByDefault = (t: ChartType): boolean => ['treemap', 'sunburst', 'funnel'].includes(traits(t).family);

const numOrUndef = (s: string): number | undefined => {
  const t = s.trim();
  if (!t) return undefined;
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
};

/** CSS colour → #rrggbb for the colour input (variables resolve against the page). */
const toHex = (css: string): string => {
  if (/^#[0-9a-f]{6}$/i.test(css)) return css;
  const v = /^var\((--[\w-]+)\)$/.exec(css);
  const raw = v ? getComputedStyle(document.documentElement).getPropertyValue(v[1]!).trim() : css;
  return /^#[0-9a-f]{6}$/i.test(raw) ? raw : '#12a565';
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
  const [y2Title, setY2Title] = useState(editing?.axisTitles?.y2 ?? '');
  const [legend, setLegend] = useState<NonNullable<Chart['legend']> | 'auto'>(editing?.legend ?? 'auto');
  const [labels, setLabels] = useState<boolean | undefined>(editing?.dataLabels);
  const [opts, setOpts] = useState<ChartOptions>(editing?.options ?? {});
  const [colors, setColors] = useState<(string | undefined)[]>(editing?.series.map((s) => s.color) ?? []);
  const [kinds, setKinds] = useState<(ChartSeries['type'])[]>(editing?.series.map((s) => s.type) ?? []);
  const [minText, setMinText] = useState(editing?.options?.min?.toString() ?? '');
  const [maxText, setMaxText] = useState(editing?.options?.max?.toString() ?? '');
  const [bins, setBins] = useState<'auto' | 'count' | 'width'>(editing?.options?.binWidth ? 'width' : editing?.options?.binCount ? 'count' : 'auto');
  const [binText, setBinText] = useState(String(editing?.options?.binWidth ?? editing?.options?.binCount ?? ''));
  const [tab, setTab] = useState<'rec' | ChartGroup>(editing ? chartGroup(editing.type) : initialType ? chartGroup(initialType) : 'rec');
  const [part, setPart] = useState<'data' | 'format'>('data');
  // In edit mode keep the chart's own series until the data settings change.
  const [dataChanged, setDataChanged] = useState(!editing);

  const tr = traits(type);
  const h = headers ?? detected;
  const setOpt = <K extends keyof ChartOptions>(k: K, v: ChartOptions[K] | undefined) => setOpts((o) => {
    const n = { ...o };
    if (v === undefined || (Array.isArray(v) && !v.length)) delete n[k]; else n[k] = v;
    return n;
  });

  const options: ChartOptions | undefined = useMemo(() => {
    const o: ChartOptions = { ...opts };
    const mn = numOrUndef(minText), mx = numOrUndef(maxText);
    if (mn !== undefined) o.min = mn; else delete o.min;
    if (mx !== undefined) o.max = mx; else delete o.max;
    delete o.binCount; delete o.binWidth;
    const b = numOrUndef(binText);
    if (tr.family === 'histogram' && b !== undefined && b > 0) { if (bins === 'count') o.binCount = Math.max(1, Math.round(b)); if (bins === 'width') o.binWidth = b; }
    if (!supportsSecondary(type) && !tr.combo) delete o.secondary;
    if (tr.family !== 'waterfall') delete o.totals;
    if (tr.family !== 'box') { delete o.hideMean; delete o.hideOutliers; }
    return Object.keys(o).length ? o : undefined;
  }, [opts, minText, maxText, binText, bins, type, tr.family, tr.combo]);

  const settings = {
    title: title.trim() || undefined, legend: legend === 'auto' ? undefined : legend,
    axisTitles: xTitle || yTitle || y2Title ? { ...(xTitle ? { x: xTitle } : {}), ...(yTitle ? { y: yTitle } : {}), ...(y2Title ? { y2: y2Title } : {}) } : undefined,
    dataLabels: labelsOnByDefault(type) ? (labels === false ? false : undefined) : labels || undefined, options,
  };
  const base: Chart | null = useMemo(() => {
    if (editing && !dataChanged) return { ...editing, type };
    return draftChart(store, { type, source, orientation, headers: h }, editing);
  }, [editing, dataChanged, type, source, orientation, h.firstRow, h.firstCol, store]); // eslint-disable-line react-hooks/exhaustive-deps
  const preview: Chart | null = base ? {
    ...base, ...settings,
    series: normalizeSeries(type, base.series.map((s, i) => ({ ...s, color: colors[i], type: kinds[i] ?? (tr.combo ? comboKind(type, i, base.series.length) : undefined) }))),
  } : null;
  const autoTitle = preview && !titleTouched ? defaultTitle(store, preview) : undefined;
  const shown = preview ? { ...preview, title: settings.title ?? autoTitle } : null;
  const resolved = shown ? resolveChart(store, shown, store.sheetIndex) : null;

  const listed: ChartType[] = tab === 'rec' ? recs : CHART_TYPES.filter((t) => t.group === tab).map((t) => t.type);
  const thumbs = (
    <div className="charttypes" role="listbox" aria-label="Chart types">
      {listed.map((t) => {
        const c = draftChart(store, { type: t, source, orientation, headers: h }) ?? (editing ? { ...editing, type: t } : null);
        return (
          <button key={t} role="option" aria-selected={type === t} className={`charttype${type === t ? ' on' : ''}`} onClick={() => setType(t)}>
            <div className="thumb">{c && <ChartSvg chart={c} data={resolveChart(store, c)} width={120} height={64} thumb />}</div>
            {chartLabel(t)}
          </button>
        );
      })}
    </div>
  );

  const ok = (): boolean => {
    if (!parsed) { store.notify('Enter a data range such as A1:D10.'); return false; }
    if (!shown || !hasNumbers(store, shown)) { store.notify('The data range has no numbers to chart.'); return false; }
    const finalChart: Chart = { ...shown, title: shown.title?.trim() || undefined };
    for (const k of ['legend', 'axisTitles', 'dataLabels', 'options'] as const) if (finalChart[k] === undefined) delete finalChart[k];
    if (editing) {
      store.change('Edit chart', 'sheet', () => {
        const i = store.sheet.charts.findIndex((c) => c.id === editing.id);
        if (i >= 0) store.sheet.charts[i] = finalChart;
      });
    } else insert(store, finalChart);
    return true;
  };

  const markData = () => { setDataChanged(true); setColors([]); setKinds([]); };
  const axes = hasAxes(type) || tr.family === 'radar';
  const xyAxes = hasAxes(type);
  const labelModes = ['pie', 'funnel', 'treemap'].includes(tr.family);
  const sec = shown ? secondaryOf(shown) : new Set<number>();
  const perSeries = resolved && !['pie', 'treemap', 'sunburst', 'waterfall', 'pareto'].includes(tr.family) && !(tr.family === 'funnel');
  const styleIdx = opts.style ?? 0;

  return (
    <Dialog title={editing ? 'Edit chart' : 'Insert chart'} close={close} wide onOk={ok} okLabel={editing ? 'Save' : 'Insert'}>
      <div className="chartgroups" role="tablist" aria-label="Chart groups">
        <button role="tab" aria-selected={tab === 'rec'} className={tab === 'rec' ? 'on' : ''} onClick={() => setTab('rec')}>Suggested</button>
        {CHART_GROUPS.map((g) => <button key={g} role="tab" aria-selected={tab === g} className={tab === g ? 'on' : ''} onClick={() => setTab(g)}>{g}</button>)}
      </div>
      {thumbs}
      <div className="chartdlg">
        <div className="chartpreview">
          {shown && resolved ? <ChartSvg chart={shown} data={resolved} width={460} height={230} /> : <p className="muted pad">Choose a range with numbers to see a preview.</p>}
        </div>
        <div className="listbox">
          <div className="segmented" role="tablist" aria-label="Settings">
            <button role="tab" aria-selected={part === 'data'} className={part === 'data' ? 'on' : ''} onClick={() => setPart('data')}>Data</button>
            <button role="tab" aria-selected={part === 'format'} className={part === 'format' ? 'on' : ''} onClick={() => setPart('format')}>Look</button>
          </div>
          {part === 'data' ? <>
            <Field label="Data range">
              <input value={rangeInput} onChange={(e) => { setRangeInput(e.target.value); markData(); }} placeholder="A1:D10" aria-invalid={!parsed} spellCheck={false} />
            </Field>
            {tr.family !== 'treemap' && tr.family !== 'sunburst' && <div className="checkrow">
              <span className="muted small">Series in</span>
              <div className="segmented" role="radiogroup" aria-label="Series in">
                {(['cols', 'rows'] as const).map((o) => {
                  const on = (orientation ?? (editing && !dataChanged ? chartOrientation(editing) : undefined) ?? (matrix.length - (h.firstRow ? 1 : 0) >= (matrix[0]?.length ?? 0) - (h.firstCol ? 1 : 0) ? 'cols' : 'rows')) === o;
                  return <button key={o} role="radio" aria-checked={on} className={on ? 'on' : ''} onClick={() => { setOrientation(o); markData(); }}>{o === 'cols' ? 'Columns' : 'Rows'}</button>;
                })}
              </div>
            </div>}
            <label className="checkrow"><input type="checkbox" checked={h.firstRow} onChange={(e) => { setHeaders({ ...h, firstRow: e.target.checked }); markData(); }} /> First row has names</label>
            <label className="checkrow"><input type="checkbox" checked={h.firstCol} onChange={(e) => { setHeaders({ ...h, firstCol: e.target.checked }); markData(); }} /> First column has labels</label>
            {tr.family === 'bubble' && <p className="muted small">Columns after the x values come in pairs: y values, then bubble sizes.</p>}
            {(tr.family === 'treemap' || tr.family === 'sunburst') && <p className="muted small">Text columns on the left are the levels, outer level first. The last column holds the values.</p>}
            {tr.family === 'histogram' && <>
              <Field label="Bins">
                <select value={bins} onChange={(e) => setBins(e.target.value as typeof bins)}>
                  <option value="auto">Automatic</option><option value="count">Number of bins</option><option value="width">Bin width</option>
                </select>
              </Field>
              {bins !== 'auto' && <Field label={bins === 'count' ? 'Number of bins' : 'Bin width'}><input inputMode="decimal" value={binText} onChange={(e) => setBinText(e.target.value)} /></Field>}
            </>}
            {tr.family === 'box' && <>
              <label className="checkrow"><input type="checkbox" checked={!opts.hideMean} onChange={(e) => setOpt('hideMean', e.target.checked ? undefined : true)} /> Show mean marker</label>
              <label className="checkrow"><input type="checkbox" checked={!opts.hideOutliers} onChange={(e) => setOpt('hideOutliers', e.target.checked ? undefined : true)} /> Show outliers</label>
            </>}
            {tr.family === 'waterfall' && resolved && <>
              <span className="seclabel">Total bars</span>
              <div className="checks">
                {resolved.categories.map((c, i) => {
                  const list = opts.totals ?? guessTotals(resolved.categories);
                  return <label key={i}><input type="checkbox" checked={list.includes(i)} onChange={(e) => setOpt('totals', e.target.checked ? [...list, i].sort((a, b) => a - b) : list.filter((x) => x !== i))} /> {c || `Point ${i + 1}`}</label>;
                })}
              </div>
            </>}
          </> : <>
            <span className="seclabel">Colours</span>
            <div className="stylepick" role="radiogroup" aria-label="Colours">
              {Array.from({ length: STYLE_COUNT }, (_, k) => (
                <button key={k} role="radio" aria-checked={styleIdx === k} className={styleIdx === k ? 'on' : ''} onClick={() => { setOpt('style', k || undefined); setColors([]); }}>
                  <span className="swatches">{Array.from({ length: 5 }, (_, i) => <i key={i} style={{ background: seriesColor(i, undefined, k) }} />)}</span>
                  {STYLE_NAMES[k]}
                </button>
              ))}
            </div>
            {perSeries && resolved.series.length > 0 && <>
              <span className="seclabel">Series</span>
              {resolved.series.map((s, i) => (
                <div key={i} className="seriesrow">
                  <input type="color" aria-label={`Colour of ${s.name}`} value={toHex(s.color)} onChange={(e) => setColors((c) => { const n = [...c]; n[i] = e.target.value.toUpperCase(); return n; })} />
                  <span title={s.name}>{s.name}</span>
                  {tr.combo ? (
                    <select aria-label={`How ${s.name} is drawn`} value={s.type} onChange={(e) => setKinds((k) => { const n = [...k]; n[i] = e.target.value as ChartSeries['type']; return n; })}>
                      <option value="column">Column</option><option value="line">Line</option><option value="area">Area</option>
                    </select>
                  ) : <span />}
                  {(supportsSecondary(type) || tr.combo) && resolved.series.length > 1 ? (
                    <label className="checkrow small"><input type="checkbox" checked={sec.has(i)} onChange={(e) => setOpt('secondary', e.target.checked ? [...sec, i].sort((a, b) => a - b) : [...sec].filter((x) => x !== i))} /> Right axis</label>
                  ) : <span />}
                </div>
              ))}
              {colors.some(Boolean) && <button className="tool small" onClick={() => setColors([])}>Use set colours</button>}
            </>}
          </>}
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
        {xyAxes && <>
          <Field label={tr.horizontal ? 'Vertical axis title' : 'Horizontal axis title'}><input value={xTitle} onChange={(e) => setXTitle(e.target.value)} /></Field>
          <Field label={tr.horizontal ? 'Horizontal axis title' : 'Vertical axis title'}><input value={yTitle} onChange={(e) => setYTitle(e.target.value)} /></Field>
          {sec.size > 0 && <Field label="Right axis title"><input value={y2Title} onChange={(e) => setY2Title(e.target.value)} /></Field>}
        </>}
        {axes && tr.family !== 'histogram' && tr.family !== 'pareto' && <>
          <Field label="Axis minimum"><input inputMode="decimal" value={minText} placeholder="Automatic" onChange={(e) => setMinText(e.target.value)} /></Field>
          <Field label="Axis maximum"><input inputMode="decimal" value={maxText} placeholder="Automatic" onChange={(e) => setMaxText(e.target.value)} /></Field>
        </>}
        {tr.family !== 'histogram' && <Field label="Number format">
          <select value={opts.numFmt ?? ''} onChange={(e) => setOpt('numFmt', e.target.value || undefined)}>
            {NUM_FORMATS.map((f) => <option key={f.code} value={f.code}>{f.label}</option>)}
          </select>
        </Field>}
        {labelModes && <Field label="Labels show">
          <select value={opts.labelMode ?? (tr.family === 'pie' ? 'percent' : 'value')} onChange={(e) => setOpt('labelMode', e.target.value as ChartOptions['labelMode'])}>
            <option value="value">Values</option><option value="percent">Percent</option><option value="both">Values and percent</option>
          </select>
        </Field>}
      </div>
      <div className="checkrow" style={{ gap: 18, flexWrap: 'wrap' }}>
        <label className="checkrow"><input type="checkbox" checked={labels ?? labelsOnByDefault(type)} onChange={(e) => setLabels(e.target.checked)} /> Show values on the chart</label>
        {axes && <label className="checkrow"><input type="checkbox" checked={opts.gridlines !== false} onChange={(e) => setOpt('gridlines', e.target.checked ? undefined : false)} /> Gridlines</label>}
      </div>
    </Dialog>
  );
};
