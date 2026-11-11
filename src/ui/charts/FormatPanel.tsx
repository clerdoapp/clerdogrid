// The chart format panel: every chart setting, grouped, in the tool panel.
import { useEffect, useState, type ReactNode } from 'react';
import { ClipboardCopy, Database, FileStack, Image as ImageIcon, MoveRight, Pencil, Repeat2, Save, Trash2 } from 'lucide-react';
import { openDialog } from '../../app/commands';
import type { GridStore } from '../../app/store';
import type {
  AxisOptions, Chart, ChartOptions, ChartSeries, DataLabelOptions, ErrorBars, FillFormat, LabelPos, LineFormat, MarkerFormat, MarkerSymbol, ShapeFormat, Trendline,
} from '../../model/workbook';
import { resolveChart, seriesColor, STYLE_COUNT, STYLE_NAMES } from './chartData';
import { CHART_GROUPS, CHART_TYPES, hasAxes, is3D, supportsErrorBars, supportsSecondary, supportsTrendlines, traits } from './chartTypes';
import { DISPLAY_UNITS } from './axes';
import { LOOKS } from './svgKit';
import { activeChart, changeType, copyChartPicture, editChart, switchRowCol } from './chartOps';
import { deleteChart } from './ChartView';
import { N, T, Tr } from '../i18n';

const S = 1.9;

const Sec = ({ title, children, open = false }: { title: string; children: ReactNode; open?: boolean }) => (
  <details className="cfsec" open={open}><summary>{Tr(title)}</summary><div className="cfbody">{children}</div></details>
);
const Row = ({ label, children }: { label: string; children: ReactNode }) => <label className="cfrow"><span>{Tr(label)}</span>{children}</label>;
const Chk = ({ label, on, set }: { label: string; on: boolean; set: (v: boolean) => void }) => (
  <label className="cfcheck"><input type="checkbox" checked={on} onChange={(e) => set(e.target.checked)} /> {Tr(label)}</label>
);

/** A text box that commits on Enter or when it loses focus. */
const Txt = ({ value, set, placeholder, label }: { value: string | undefined; set: (v: string | undefined) => void; placeholder?: string; label?: string }) => {
  const [t, setT] = useState(value ?? '');
  useEffect(() => setT(value ?? ''), [value]);
  const commit = () => { const v = t.trim() ? t : undefined; if (v !== value) set(v); };
  return <input aria-label={Tr(label)} value={t} placeholder={Tr(placeholder)} spellCheck={false} onChange={(e) => setT(e.target.value)} onBlur={commit} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') commit(); }} />;
};
const Num = ({ value, set, placeholder = N('Automatic'), min, max, label }: { value: number | undefined; set: (v: number | undefined) => void; placeholder?: string; min?: number; max?: number; label?: string }) => {
  const [t, setT] = useState(value?.toString() ?? '');
  useEffect(() => setT(value?.toString() ?? ''), [value]);
  const commit = () => {
    const s = t.trim();
    if (!s) { if (value !== undefined) set(undefined); return; }
    let n = Number(s);
    if (!Number.isFinite(n)) { setT(value?.toString() ?? ''); return; }
    if (min !== undefined) n = Math.max(min, n);
    if (max !== undefined) n = Math.min(max, n);
    if (n !== value) set(n);
  };
  return <input aria-label={Tr(label)} inputMode="decimal" value={t} placeholder={Tr(placeholder)} onChange={(e) => setT(e.target.value)} onBlur={commit} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') commit(); }} />;
};
const Sel = <T extends string>({ value, set, options, label }: { value: T; set: (v: T) => void; options: [T, string][]; label?: string }) => (
  <select aria-label={Tr(label)} value={value} onChange={(e) => set(e.target.value as T)}>{options.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
);

/** CSS colour → #rrggbb for colour inputs (variables resolve against the page). */
export const toHex = (css: string | undefined, fallback = '#12a565'): string => {
  if (!css) return fallback;
  if (/^#[0-9a-f]{6}$/i.test(css)) return css;
  const v = /^var\((--[\w-]+)/.exec(css);
  const raw = v ? getComputedStyle(document.documentElement).getPropertyValue(v[1]!).trim() : css;
  return /^#[0-9a-f]{6}$/i.test(raw) ? raw : fallback;
};
const Color = ({ value, set, auto, label }: { value: string | undefined; set: (v: string | undefined) => void; auto?: string; label: string }) => (
  <span className="cfcolor">
    <input type="color" aria-label={Tr(label)} value={toHex(value ?? auto)} onChange={(e) => set(e.target.value.toUpperCase())} />
    {value ? <button className="ico sm" title={T('Automatic')} aria-label={T('{label}: automatic', { label })} onClick={() => set(undefined)}>×</button> : <span className="muted small">{T('Auto')}</span>}
  </span>
);

const DASHES: [NonNullable<LineFormat['dash']>, string][] = [['solid', N('Solid')], ['dot', N('Dots')], ['dash', N('Dashes')], ['dashDot', N('Dash dot')], ['longDash', N('Long dashes')], ['longDashDot', N('Long dash dot')]];
const MARKERS: [MarkerSymbol, string][] = [['auto', N('Automatic')], ['none', N('None')], ['circle', N('Circle')], ['square', N('Square')], ['diamond', N('Diamond')], ['triangle', N('Triangle')], ['x', N('Cross')], ['plus', N('Plus')], ['star', N('Star')], ['dash', N('Dash')], ['dot', N('Dot')]];

const FillEditor = ({ fill, set, auto, label }: { fill: FillFormat | undefined; set: (f: FillFormat | undefined) => void; auto?: string; label: string }) => {
  const f = fill ?? {};
  const mode = f.none ? 'none' : f.gradient ? 'gradient' : 'solid';
  const up = (p: FillFormat) => { const n = { ...f, ...p }; for (const k of Object.keys(n) as (keyof FillFormat)[]) if (n[k] === undefined) delete n[k]; set(Object.keys(n).length ? n : undefined); };
  return <>
    <Row label={Tr(label)}><Sel label={Tr(label)} value={mode} set={(m) => (m === 'none' ? set({ none: true }) : m === 'gradient' ? set({ gradient: { from: toHex(f.color ?? auto), to: '#FFFFFF', angle: 90 } }) : set(f.color ? { color: f.color } : undefined))} options={[['solid', N('Solid')], ['gradient', N('Gradient')], ['none', N('None')]]} /></Row>
    {mode === 'solid' && <Row label={T('Colour')}><Color label={T('{label} colour', { label })} value={f.color} auto={auto} set={(c) => up({ color: c })} /></Row>}
    {mode === 'gradient' && <>
      <Row label={T('From')}><Color label={T('Gradient start')} value={f.gradient!.from} set={(c) => up({ gradient: { ...f.gradient!, from: c ?? '#FFFFFF' } })} /></Row>
      <Row label={T('To')}><Color label={T('Gradient end')} value={f.gradient!.to} set={(c) => up({ gradient: { ...f.gradient!, to: c ?? '#FFFFFF' } })} /></Row>
      <Row label={T('Angle')}><Num label={T('Gradient angle')} value={f.gradient!.angle} placeholder="90" min={0} max={360} set={(a) => up({ gradient: { ...f.gradient!, angle: a } })} /></Row>
    </>}
    {mode !== 'none' && <Row label={T('Clear %')}><Num label={T('Transparency')} value={f.transparency} placeholder="0" min={0} max={100} set={(v) => up({ transparency: v })} /></Row>}
  </>;
};

const LineEditor = ({ line, set, auto, label }: { line: LineFormat | undefined; set: (l: LineFormat | undefined) => void; auto?: string; label: string }) => {
  const l = line ?? {};
  const up = (p: LineFormat) => { const n = { ...l, ...p }; for (const k of Object.keys(n) as (keyof LineFormat)[]) if (n[k] === undefined || n[k] === false) delete n[k]; set(Object.keys(n).length ? n : undefined); };
  return <>
    <Chk label={T('None')} on={!!l.none} set={(v) => up({ none: v || undefined })} />
    {!l.none && <>
      <Row label={T('{label} colour', { label })}><Color label={T('{label} colour', { label })} value={l.color} auto={auto} set={(c) => up({ color: c })} /></Row>
      <Row label={T('Width (pt)')}><Num label={T('{label} width', { label })} value={l.width} placeholder={T('Auto')} min={0} max={20} set={(w) => up({ width: w })} /></Row>
      <Row label={T('Dashes')}><Sel label={T('{label} dashes', { label })} value={l.dash ?? 'solid'} set={(d) => up({ dash: d === 'solid' ? undefined : d })} options={DASHES} /></Row>
    </>}
  </>;
};

const ShapeEditor = ({ value, set, label }: { value: ShapeFormat | undefined; set: (s: ShapeFormat | undefined) => void; label: string }) => (
  <>
    <FillEditor label={T('{label} fill', { label })} fill={value?.fill} auto="#ffffff" set={(f) => set(f || value?.line ? { ...value, fill: f } : undefined)} />
    <LineEditor label={T('Border')} line={value?.line} auto="#a6b2ab" set={(l) => set(l || value?.fill ? { ...value, line: l } : undefined)} />
  </>
);

const POSITIONS = (fam: string, stacked: boolean): [LabelPos, string][] => {
  const all: [LabelPos, string][] = [['auto', N('Automatic')]];
  if (fam === 'pie' || fam === 'ofPie') return [...all, ['center', N('Centre')], ['insideEnd', N('Inside end')], ['outsideEnd', N('Outside end')], ['bestFit', N('Best fit')]];
  if (fam === 'cartesian' || fam === 'stock' || fam === 'histogram' || fam === 'waterfall' || fam === 'pareto') {
    const bars: [LabelPos, string][] = [['center', N('Centre')], ['insideEnd', N('Inside end')], ['insideBase', N('Inside base')], ...(stacked ? [] : [['outsideEnd', N('Outside end')]] as [LabelPos, string][])];
    const lines: [LabelPos, string][] = [['above', N('Above')], ['below', N('Below')], ['left', N('Left')], ['right', N('Right')]];
    return [...all, ...bars, ...lines];
  }
  if (fam === 'scatter' || fam === 'bubble') return [...all, ['center', N('Centre')], ['above', N('Above')], ['below', N('Below')], ['left', N('Left')], ['right', N('Right')]];
  return all;
};

const LabelEditor = ({ value, set, fam, stacked, pie }: { value: DataLabelOptions; set: (v: DataLabelOptions) => void; fam: string; stacked: boolean; pie: boolean }) => {
  const up = (p: Partial<DataLabelOptions>) => { const n = { ...value, ...p }; for (const k of Object.keys(n) as (keyof DataLabelOptions)[]) if (n[k] === undefined || n[k] === false && k !== 'leaderLines') delete n[k]; set(n); };
  return <>
    <div className="cfchecks">
      <Chk label={T('Value')} on={!!value.value} set={(v) => up({ value: v })} />
      <Chk label={T('Series name')} on={!!value.seriesName} set={(v) => up({ seriesName: v })} />
      <Chk label={T('Category name')} on={!!value.categoryName} set={(v) => up({ categoryName: v })} />
      {(pie || fam === 'cartesian') && <Chk label={T('Percentage')} on={!!value.percent} set={(v) => up({ percent: v })} />}
      {fam === 'bubble' && <Chk label={T('Bubble size')} on={!!value.bubbleSize} set={(v) => up({ bubbleSize: v })} />}
      {pie && <Chk label={T('Leader lines')} on={value.leaderLines !== false} set={(v) => up({ leaderLines: v ? undefined : false })} />}
    </div>
    <Row label={T('Text from cells')}><Txt label={T('Label text from cells')} value={value.fromCells} placeholder={T('Sheet1!D2:D9')} set={(v) => up({ fromCells: v })} /></Row>
    <Row label={T('Position')}><Sel label={T('Label position')} value={value.position ?? 'auto'} set={(v) => up({ position: v === 'auto' ? undefined : v })} options={POSITIONS(fam, stacked)} /></Row>
    <Row label={T('Separator')}><Txt label={T('Label separator')} value={value.separator} placeholder=", " set={(v) => up({ separator: v })} /></Row>
    <Row label={T('Number format')}><Txt label={T('Label number format')} value={value.numFmt} placeholder={T('Same as data')} set={(v) => up({ numFmt: v })} /></Row>
  </>;
};

const TREND_KINDS: [Trendline['kind'], string][] = [['linear', N('Linear')], ['exp', N('Exponential')], ['log', N('Logarithmic')], ['poly', N('Polynomial')], ['power', N('Power')], ['movingAvg', N('Moving average')]];

const TrendEditor = ({ t, set, remove }: { t: Trendline; set: (t: Trendline) => void; remove: () => void }) => {
  const up = (p: Partial<Trendline>) => { const n = { ...t, ...p }; for (const k of Object.keys(n) as (keyof Trendline)[]) if (n[k] === undefined || n[k] === false) delete n[k]; set(n); };
  return (
    <div className="cfitem">
      <Row label={T('Kind')}><Sel label={T('Trendline kind')} value={t.kind} set={(k) => up({ kind: k, order: k === 'poly' ? t.order ?? 2 : undefined, period: k === 'movingAvg' ? t.period ?? 2 : undefined })} options={TREND_KINDS} /></Row>
      {t.kind === 'poly' && <Row label={T('Order')}><Sel label={T('Order')} value={String(t.order ?? 2)} set={(v) => up({ order: +v })} options={[2, 3, 4, 5, 6].map((n) => [String(n), String(n)])} /></Row>}
      {t.kind === 'movingAvg' && <Row label={T('Period')}><Num label={T('Period')} value={t.period} placeholder="2" min={2} max={255} set={(v) => up({ period: v === undefined ? undefined : Math.round(v) })} /></Row>}
      {t.kind !== 'movingAvg' && <>
        <Row label={T('Forward')}><Num label={T('Forecast forward')} value={t.forward} placeholder="0" min={0} set={(v) => up({ forward: v || undefined })} /></Row>
        <Row label={T('Backward')}><Num label={T('Forecast backward')} value={t.backward} placeholder="0" min={0} set={(v) => up({ backward: v || undefined })} /></Row>
        {['linear', 'exp', 'poly'].includes(t.kind) && <Row label={T('Crosses y at')}><Num label={T('Set intercept')} value={t.intercept} placeholder={T('Automatic')} set={(v) => up({ intercept: v })} /></Row>}
        <div className="cfchecks">
          <Chk label={T('Show equation')} on={!!t.showEquation} set={(v) => up({ showEquation: v })} />
          <Chk label={T('Show R²')} on={!!t.showR2} set={(v) => up({ showR2: v })} />
        </div>
      </>}
      <Row label={T('Name')}><Txt label={T('Trendline name')} value={t.name} placeholder={T('Automatic')} set={(v) => up({ name: v })} /></Row>
      <LineEditor label={T('Line')} line={t.line} set={(l) => up({ line: l })} />
      <button className="tool small" onClick={remove}><Trash2 size={14} strokeWidth={S} /><span>{T('Remove trendline')}</span></button>
    </div>
  );
};

const ERR_KINDS: [ErrorBars['kind'], string][] = [['fixed', N('Fixed value')], ['percent', N('Percentage')], ['stdDev', N('Standard deviation')], ['stdErr', N('Standard error')], ['custom', N('Custom')]];
const ErrEditor = ({ e, set, remove, xy }: { e: ErrorBars; set: (e: ErrorBars) => void; remove: () => void; xy: boolean }) => {
  const up = (p: Partial<ErrorBars>) => { const n = { ...e, ...p }; for (const k of Object.keys(n) as (keyof ErrorBars)[]) if (n[k] === undefined || n[k] === false) delete n[k]; set(n); };
  return (
    <div className="cfitem">
      {xy && <Row label={T('Direction')}><Sel label={T('Error bar direction')} value={e.dir ?? 'y'} set={(d) => up({ dir: d === 'y' ? undefined : d })} options={[['y', N('Vertical')], ['x', N('Horizontal')]]} /></Row>}
      <Row label={T('Amount')}><Sel label={T('Error amount')} value={e.kind} set={(k) => up({ kind: k })} options={ERR_KINDS} /></Row>
      {e.kind !== 'custom' && e.kind !== 'stdErr' && <Row label={e.kind === 'percent' ? T('Percent') : e.kind === 'stdDev' ? T('Times') : T('Value')}><Num label={T('Error amount value')} value={e.value} placeholder={e.kind === 'percent' ? '5' : '1'} min={0} set={(v) => up({ value: v })} /></Row>}
      {e.kind === 'custom' && <>
        <Row label={T('Plus')}><Txt label={T('Plus amounts')} value={e.plus} placeholder={T('Sheet1!E2:E9 or {1,2}')} set={(v) => up({ plus: v })} /></Row>
        <Row label={T('Minus')}><Txt label={T('Minus amounts')} value={e.minus} placeholder={T('Sheet1!F2:F9')} set={(v) => up({ minus: v })} /></Row>
      </>}
      <Row label={T('Show')}><Sel label={T('Error bar ends')} value={e.type ?? 'both'} set={(v) => up({ type: v === 'both' ? undefined : v })} options={[['both', N('Both')], ['plus', N('Plus')], ['minus', N('Minus')]]} /></Row>
      <Chk label={T('End caps')} on={!e.noCap} set={(v) => up({ noCap: !v || undefined })} />
      <LineEditor label={T('Line')} line={e.line} set={(l) => up({ line: l })} />
      <button className="tool small" onClick={remove}><Trash2 size={14} strokeWidth={S} /><span>{T('Remove error bars')}</span></button>
    </div>
  );
};

const AxisEditor = ({ a, set, value, date, title, setTitle, scatter }: { a: AxisOptions; set: (a: AxisOptions) => void; value: boolean; date: boolean; title?: string; setTitle: (t: string | undefined) => void; scatter: boolean }) => {
  const up = (p: Partial<AxisOptions>) => { const n = { ...a, ...p }; for (const k of Object.keys(n) as (keyof AxisOptions)[]) if (n[k] === undefined) delete n[k]; set(n); };
  const crossMode = a.crossesAt !== undefined ? 'at' : a.crosses ?? 'auto';
  return <>
    <Row label={T('Title')}><Txt label={T('Axis title')} value={title} placeholder={T('=Sheet1!A1 links a cell')} set={setTitle} /></Row>
    {value && <>
      <Row label={T('Minimum')}><Num label={T('Minimum')} value={a.min} set={(v) => up({ min: v })} /></Row>
      <Row label={T('Maximum')}><Num label={T('Maximum')} value={a.max} set={(v) => up({ max: v })} /></Row>
      <Row label={T('Major unit')}><Num label={T('Major unit')} value={a.majorUnit} min={0} set={(v) => up({ majorUnit: v || undefined })} /></Row>
      <Row label={T('Minor unit')}><Num label={T('Minor unit')} value={a.minorUnit} min={0} set={(v) => up({ minorUnit: v || undefined })} /></Row>
      <Chk label={T('Logarithmic scale')} on={!!a.logBase} set={(v) => up({ logBase: v ? 10 : undefined })} />
      {a.logBase && <Row label={T('Base')}><Num label={T('Log base')} value={a.logBase} placeholder="10" min={2} max={1000} set={(v) => up({ logBase: v ?? 10 })} /></Row>}
      <Row label={T('Display units')}><Sel label={T('Display units')} value={a.dispUnits ?? 'none'} set={(v) => up({ dispUnits: v === 'none' ? undefined : v })} options={[['none', N('None')], ...DISPLAY_UNITS.map((u): [typeof u.value, string] => [u.value, u.label])]} /></Row>
      {a.dispUnits && <Chk label={T('Show units label')} on={a.dispUnitsLabel !== false} set={(v) => up({ dispUnitsLabel: v ? undefined : false })} />}
      <Row label={T('Number format')}><Txt label={T('Axis number format')} value={a.numFmt} placeholder={T('Same as data')} set={(v) => up({ numFmt: v })} /></Row>
    </>}
    {!value && date && <>
      <Row label={T('Axis type')}><Sel label={T('Axis type')} value={a.dateAxis === true ? 'date' : a.dateAxis === false ? 'text' : 'auto'} set={(v) => up({ dateAxis: v === 'date' ? true : v === 'text' ? false : undefined })} options={[['auto', N('Automatic')], ['text', N('Text')], ['date', N('Dates')]]} /></Row>
      {a.dateAxis !== false && <Row label={T('Base unit')}><Sel label={T('Base unit')} value={a.baseUnit ?? 'auto'} set={(v) => up({ baseUnit: v === 'auto' ? undefined : v })} options={[['auto', N('Automatic')], ['days', N('Days')], ['months', N('Months')], ['years', N('Years')]]} /></Row>}
    </>}
    <Chk label={T('Reverse order')} on={!!a.reverse} set={(v) => up({ reverse: v || undefined })} />
    <Row label={value || scatter ? T('Other axis crosses') : T('Value axis crosses')}>
      <Sel label={T('Crossing')} value={crossMode} set={(v) => up({ crosses: v === 'min' || v === 'max' ? v : undefined, crossesAt: v === 'at' ? a.crossesAt ?? (value || scatter ? 0 : 1) : undefined })}
        options={[['auto', N('Automatic')], ['min', value || scatter ? N('Minimum') : N('First')], ['max', value || scatter ? N('Maximum') : N('Last')], ['at', value || scatter ? N('At value') : N('At category')]]} />
    </Row>
    {crossMode === 'at' && <Row label={value || scatter ? T('Value') : T('Category no.')}><Num label={T('Crosses at')} value={a.crossesAt} placeholder="0" set={(v) => up({ crossesAt: v ?? 0 })} /></Row>}
    <Row label={T('Major ticks')}><Sel label={T('Major tick marks')} value={a.majorTick ?? 'none'} set={(v) => up({ majorTick: v === 'none' ? undefined : v })} options={[['none', N('None')], ['out', N('Outside')], ['in', N('Inside')], ['cross', N('Cross')]]} /></Row>
    <Row label={T('Minor ticks')}><Sel label={T('Minor tick marks')} value={a.minorTick ?? 'none'} set={(v) => up({ minorTick: v === 'none' ? undefined : v })} options={[['none', N('None')], ['out', N('Outside')], ['in', N('Inside')], ['cross', N('Cross')]]} /></Row>
    <Row label={T('Labels')}><Sel label={T('Label position')} value={a.labelPos ?? 'nextTo'} set={(v) => up({ labelPos: v === 'nextTo' ? undefined : v })} options={[['nextTo', N('Next to axis')], ['low', N('Low')], ['high', N('High')], ['none', N('None')]]} /></Row>
    <div className="cfchecks">
      <Chk label={T('Major gridlines')} on={a.majorGrid ?? (value || scatter)} set={(v) => up({ majorGrid: v === (value || scatter) ? undefined : v })} />
      <Chk label={T('Minor gridlines')} on={!!a.minorGrid} set={(v) => up({ minorGrid: v || undefined })} />
      <Chk label={T('Hide axis')} on={!!a.hidden} set={(v) => up({ hidden: v || undefined })} />
    </div>
    <LineEditor label={T('Axis line')} line={a.line} auto="#a6b2ab" set={(l) => up({ line: l })} />
  </>;
};

const ChartFormat = ({ store, chart }: { store: GridStore; chart: Chart }) => {
  const [si, setSi] = useState(0);
  const [pi, setPi] = useState(-1);
  const [axis, setAxis] = useState<'x' | 'y' | 'y2'>('y');
  const id = chart.id;
  const tr = traits(chart.type);
  const o = chart.options ?? {};
  const data = resolveChart(store, chart);
  const sIdx = Math.min(si, Math.max(0, chart.series.length - 1));
  const ser = chart.series[sIdx];
  const rs = data.series[sIdx];
  const opt = <K extends keyof ChartOptions>(k: K, v: ChartOptions[K] | undefined, label = N('Format chart')) => editChart(store, id, label, (c) => { c.options = { ...c.options, [k]: v }; });
  const serUp = (p: Partial<ChartSeries>, label = N('Format series')) => editChart(store, id, label, (c) => {
    const s = c.series[sIdx];
    if (!s) return;
    Object.assign(s, p);
    for (const k of Object.keys(p) as (keyof ChartSeries)[]) if (p[k] === undefined) delete s[k];
  });
  const pie = tr.family === 'pie' || tr.family === 'ofPie';
  const stacked = tr.stack !== 'none';
  const flatAxes = hasAxes(chart.type) && !tr.d3 && tr.family !== 'surface';
  const scatter = tr.family === 'scatter' || tr.family === 'bubble';
  const sec = new Set(o.secondary ?? []);
  const axOpts = (k: 'x' | 'y' | 'y2'): AxisOptions => {
    const a = { ...(o.axes?.[k] ?? {}) };
    if (k === 'y') {
      if (o.min !== undefined) a.min = o.min;
      if (o.max !== undefined) a.max = o.max;
      if (o.numFmt) a.numFmt = o.numFmt;
      if (o.gridlines === false) a.majorGrid = false;
    }
    return a;
  };
  const setAx = (k: 'x' | 'y' | 'y2', a: AxisOptions) => editChart(store, id, N('Format axis'), (c) => {
    const n = { ...a };
    const opts: ChartOptions = { ...c.options };
    // The value axis keeps its bounds, number format and gridline switch in the chart-wide fields.
    if (k === 'y') {
      opts.min = n.min; opts.max = n.max; opts.numFmt = n.numFmt; opts.gridlines = n.majorGrid === false ? false : undefined;
      delete n.min; delete n.max; delete n.numFmt; delete n.majorGrid;
    }
    opts.axes = { ...opts.axes, [k]: n };
    c.options = opts;
  });
  const setTitle = (k: 'x' | 'y' | 'y2', t: string | undefined) => editChart(store, id, N('Axis title'), (c) => { c.axisTitles = { ...c.axisTitles, [k]: t }; });
  const lo: DataLabelOptions = o.labels ?? (pie ? { value: o.labelMode !== undefined && o.labelMode !== 'percent', percent: (o.labelMode ?? 'percent') !== 'value' } : { value: true });
  const v3 = o.view3D ?? {};
  const v3up = (p: Partial<NonNullable<ChartOptions['view3D']>>) => opt('view3D', { ...v3, ...p }, N('Format 3D'));

  return (
    <div className="chartformat">
      <div className="toollist">
        <button className="tool" onClick={() => openDialog(store, 'chart', { id })}><Pencil size={16} strokeWidth={S} /><span>{T('Edit chart…')}</span></button>
        <button className="tool" onClick={() => openDialog(store, 'chartData', { id })}><Database size={16} strokeWidth={S} /><span>{T('Select data…')}</span></button>
        {!tr.combo && tr.family !== 'treemap' && tr.family !== 'sunburst' && <button className="tool" onClick={() => switchRowCol(store, id)}><Repeat2 size={16} strokeWidth={S} /><span>{T('Switch rows and columns')}</span></button>}
        <button className="tool" onClick={() => openDialog(store, 'chartMove', { id })}><MoveRight size={16} strokeWidth={S} /><span>{T('Move chart…')}</span></button>
        <button className="tool" onClick={() => void copyChartPicture(store, id)}><ClipboardCopy size={16} strokeWidth={S} /><span>{T('Copy as picture')}</span></button>
        <button className="tool" onClick={() => openDialog(store, 'chartTemplate', { id })}><Save size={16} strokeWidth={S} /><span>{T('Save as template…')}</span></button>
        <button className="tool" onClick={() => openDialog(store, 'altText', { type: 'chart', id })}><ImageIcon size={16} strokeWidth={S} /><span>{T('Alt text…')}</span></button>
        {store.sheet.kind !== 'chart' && <button className="tool" onClick={() => deleteChart(store, id)}><Trash2 size={16} strokeWidth={S} /><span>{T('Delete chart')}</span></button>}
      </div>

      <Sec title={T('Chart')} open>
        <Row label={T('Type')}>
          <select aria-label={T('Chart type')} value={chart.type} onChange={(e) => changeType(store, id, e.target.value as Chart['type'])}>
            {CHART_GROUPS.map((g) => <optgroup key={g} label={Tr(g)}>{CHART_TYPES.filter((t) => t.group === g).map((t) => <option key={t.type} value={t.type}>{Tr(t.label)}</option>)}</optgroup>)}
          </select>
        </Row>
        <Row label={T('Title')}><Txt label={T('Chart title')} value={chart.title} placeholder={T('=Sheet1!A1 links a cell')} set={(t) => editChart(store, id, N('Chart title'), (c) => { c.title = t; })} /></Row>
        <Row label={T('Colours')}>
          <Sel label={T('Colours')} value={String(o.style ?? 0)} set={(v) => opt('style', +v || undefined, N('Chart colours'))} options={Array.from({ length: STYLE_COUNT }, (_, k) => [String(k), STYLE_NAMES[k]!])} />
        </Row>
        <div className="cfswatch">{Array.from({ length: 6 }, (_, i) => <i key={i} style={{ background: seriesColor(i, undefined, o.style ?? 0) }} />)}</div>
        <Row label={T('Style')}><Sel label={T('Chart style')} value={String(o.look ?? 0)} set={(v) => opt('look', +v || undefined, N('Chart style'))} options={LOOKS.map((l, k) => [String(k), l.name])} /></Row>
        <Row label={T('Legend')}>
          <Sel label={T('Legend')} value={chart.legend ?? 'auto'} set={(v) => editChart(store, id, N('Legend'), (c) => { c.legend = v === 'auto' ? undefined : v; })}
            options={[['auto', N('Automatic')], ['bottom', N('Bottom')], ['top', N('Top')], ['right', N('Right')], ['left', N('Left')], ['none', N('None')]]} />
        </Row>
        <Chk label={T('Legend over the plot')} on={!!o.legendOverlay} set={(v) => opt('legendOverlay', v || undefined, N('Legend'))} />
      </Sec>

      <Sec title={T('Data labels')}>
        <Chk label={T('Show data labels')} on={!!chart.dataLabels || (chart.dataLabels === undefined && ['treemap', 'sunburst', 'funnel'].includes(tr.family))} set={(v) => editChart(store, id, N('Data labels'), (c) => { c.dataLabels = v || (['treemap', 'sunburst', 'funnel'].includes(tr.family) ? false : undefined); })} />
        {chart.dataLabels && <LabelEditor value={lo} fam={tr.family} stacked={stacked} pie={pie} set={(l) => opt('labels', l, N('Data labels'))} />}
      </Sec>

      {chart.series.length > 0 && rs && ser && (
        <Sec title={T('Series')}>
          <Row label={T('Series')}>
            <select aria-label={T('Series')} value={sIdx} onChange={(e) => { setSi(+e.target.value); setPi(-1); }}>
              {data.series.map((s, i) => <option key={i} value={i}>{s.name}</option>)}
            </select>
          </Row>
          {!pie && <FillEditor label={T('Fill')} fill={ser.fill ?? (ser.color ? { color: ser.color } : undefined)} auto={rs.color}
            set={(f) => serUp(f && !f.gradient && !f.none && f.transparency === undefined ? { color: f.color, fill: undefined } : { fill: f, color: f?.gradient ? undefined : ser.color })} />}
          {!pie && <LineEditor label={tr.base === 'line' || tr.family === 'scatter' || tr.family === 'radar' || tr.family === 'stock' ? T('Line') : T('Border')} line={ser.line} auto={rs.color} set={(l) => serUp({ line: l })} />}
          {(tr.base === 'line' || tr.family === 'scatter' || tr.family === 'radar' || tr.family === 'stock' || (tr.combo && rs.type === 'line')) && <>
            <Row label={T('Marker')}><Sel label={T('Marker')} value={ser.marker?.symbol ?? 'auto'} set={(v) => serUp({ marker: { ...ser.marker, symbol: v === 'auto' ? undefined : v } })} options={MARKERS} /></Row>
            <Row label={T('Marker size')}><Num label={T('Marker size')} value={ser.marker?.size} placeholder="5" min={2} max={72} set={(v) => serUp({ marker: { ...ser.marker, size: v } })} /></Row>
            <Row label={T('Marker fill')}><Color label={T('Marker fill')} value={ser.marker?.fill} set={(c) => serUp({ marker: { ...ser.marker, fill: c } as MarkerFormat })} /></Row>
            <Row label={T('Marker line')}><Color label={T('Marker line')} value={ser.marker?.line} auto={rs.color} set={(c) => serUp({ marker: { ...ser.marker, line: c } as MarkerFormat })} /></Row>
            {tr.family !== 'radar' && <Chk label={T('Smooth line')} on={ser.smooth ?? tr.smooth} set={(v) => serUp({ smooth: v === tr.smooth ? undefined : v })} />}
          </>}
          {(tr.base === 'column' || tr.family === 'bar3d') && !pie && tr.family !== 'surface' && <Chk label={T('Invert colours if negative')} on={!!ser.invertIfNegative} set={(v) => serUp({ invertIfNegative: v || undefined })} />}
          {pie && <Row label={T('Pull out %')}><Num label={T('Pull slices out')} value={ser.explosion} placeholder={tr.exploded ? '18' : '0'} min={0} max={400} set={(v) => serUp({ explosion: v })} /></Row>}
          {(supportsSecondary(chart.type) || tr.combo) && chart.series.length > 1 && <Chk label={T('Plot on the right axis')} on={sec.has(sIdx)} set={(v) => opt('secondary', v ? [...sec, sIdx].sort((a, b) => a - b) : [...sec].filter((x) => x !== sIdx))} />}
          {tr.combo && <Row label={T('Draw as')}><Sel label={T('Series kind')} value={rs.type} set={(v) => serUp({ type: v })} options={[['column', N('Column')], ['line', N('Line')], ['area', N('Area')]]} /></Row>}
          <div className="mini">{T('Labels for this series')}</div>
          <Chk label={T('Own label settings')} on={!!ser.labels} set={(v) => serUp({ labels: v ? { ...lo } : undefined })} />
          {ser.labels && <LabelEditor value={ser.labels} fam={tr.family} stacked={stacked} pie={pie} set={(l) => serUp({ labels: l })} />}

          <div className="mini">{T('Point')}</div>
          <Row label={T('Point')}>
            <select aria-label={T('Point')} value={pi} onChange={(e) => setPi(+e.target.value)}>
              <option value={-1}>{T('Choose a point')}</option>
              {data.categories.map((c, i) => <option key={i} value={i}>{c || T('Point {p0}', { p0: i + 1 })}</option>)}
            </select>
          </Row>
          {pi >= 0 && (() => {
            const p = ser.points?.find((x) => x.idx === pi) ?? { idx: pi };
            const setP = (patch: object) => serUp({
              points: [...(ser.points ?? []).filter((x) => x.idx !== pi), { ...p, ...patch }].map((x) => { const n = { ...x } as Record<string, unknown>; for (const k of Object.keys(n)) if (n[k] === undefined) delete n[k]; return n as unknown as typeof x; }).filter((x) => Object.keys(x).length > 1).sort((a, b) => a.idx - b.idx),
            }, N('Format point'));
            return <>
              <FillEditor label={T('Point fill')} fill={p.fill} auto={pie ? seriesColor(pi, undefined, o.style ?? 0) : rs.color} set={(f) => setP({ fill: f })} />
              <LineEditor label={T('Point border')} line={p.line} set={(l) => setP({ line: l })} />
              {pie && <Row label={T('Pull out %')}><Num label={T('Pull point out')} value={p.explosion} placeholder="0" min={0} max={400} set={(v) => setP({ explosion: v })} /></Row>}
              {(tr.base === 'line' || tr.family === 'scatter') && <Row label={T('Marker')}><Sel label={T('Point marker')} value={p.marker?.symbol ?? 'auto'} set={(v) => setP({ marker: v === 'auto' ? undefined : { ...p.marker, symbol: v } })} options={MARKERS} /></Row>}
              <button className="tool small" onClick={() => serUp({ points: (ser.points ?? []).filter((x) => x.idx !== pi) }, N('Reset point'))}>{T('Reset this point')}</button>
            </>;
          })()}

          {supportsTrendlines(chart.type) && <>
            <div className="mini">{T('Trendlines')}</div>
            {(ser.trendlines ?? []).map((t, k) => <TrendEditor key={k} t={t} set={(n) => serUp({ trendlines: ser.trendlines!.map((x, j) => (j === k ? n : x)) }, N('Trendline'))} remove={() => serUp({ trendlines: ser.trendlines!.filter((_, j) => j !== k) }, N('Remove trendline'))} />)}
            <button className="tool small" onClick={() => serUp({ trendlines: [...(ser.trendlines ?? []), { kind: 'linear' }] }, N('Add trendline'))}>{T('Add trendline')}</button>
          </>}
          {supportsErrorBars(chart.type) && <>
            <div className="mini">{T('Error bars')}</div>
            {(ser.errorBars ?? []).map((e, k) => <ErrEditor key={k} e={e} xy={scatter} set={(n) => serUp({ errorBars: ser.errorBars!.map((x, j) => (j === k ? n : x)) }, N('Error bars'))} remove={() => serUp({ errorBars: ser.errorBars!.filter((_, j) => j !== k) }, N('Remove error bars'))} />)}
            <button className="tool small" onClick={() => serUp({ errorBars: [...(ser.errorBars ?? []), { kind: 'stdErr' }] }, N('Add error bars'))}>{T('Add error bars')}</button>
          </>}
        </Sec>
      )}

      {(flatAxes || tr.family === 'radar' || tr.d3 || tr.family === 'surface') && (
        <Sec title={T('Axes')}>
          {flatAxes && <div className="segmented" role="tablist" aria-label={T('Axis')}>
            {(['x', 'y', ...(sec.size || chart.type === 'comboSecondary' || (tr.volume && tr.family === 'stock') ? ['y2'] : [])] as ('x' | 'y' | 'y2')[]).map((k) => (
              <button key={k} role="tab" aria-selected={axis === k} className={axis === k ? 'on' : ''} onClick={() => setAxis(k)}>{k === 'x' ? (tr.horizontal ? T('Vertical') : T('Horizontal')) : k === 'y' ? (tr.horizontal ? T('Horizontal') : T('Vertical')) : T('Right')}</button>
            ))}
          </div>}
          {(() => {
            const k = flatAxes ? axis : 'y';
            const value = k !== 'x' || scatter;
            return <AxisEditor key={k} a={axOpts(k)} set={(a) => setAx(k, a)} value={value} date={k === 'x' && !scatter} scatter={scatter && k === 'x'} title={chart.axisTitles?.[k]} setTitle={(t) => setTitle(k, t)} />;
          })()}
        </Sec>
      )}

      <Sec title={T('Chart elements')}>
        {(tr.family === 'cartesian' || tr.family === 'stock') && !tr.horizontal && <Chk label={T('Data table')} on={!!o.dataTable} set={(v) => opt('dataTable', v ? {} : undefined, N('Data table'))} />}
        {o.dataTable && <div className="cfchecks indent">
          <Chk label={T('Legend keys')} on={o.dataTable.keys !== false} set={(v) => opt('dataTable', { ...o.dataTable, keys: v ? undefined : false })} />
          <Chk label={T('Row lines')} on={o.dataTable.horizontal !== false} set={(v) => opt('dataTable', { ...o.dataTable, horizontal: v ? undefined : false })} />
          <Chk label={T('Column lines')} on={o.dataTable.vertical !== false} set={(v) => opt('dataTable', { ...o.dataTable, vertical: v ? undefined : false })} />
          <Chk label={T('Outline')} on={o.dataTable.outline !== false} set={(v) => opt('dataTable', { ...o.dataTable, outline: v ? undefined : false })} />
        </div>}
        {(tr.base === 'line' || tr.family === 'stock' || tr.combo) && <>
          <Chk label={T('Up/down bars')} on={o.upDownBars ?? (tr.family === 'stock' && tr.open)} set={(v) => opt('upDownBars', v, N('Up/down bars'))} />
          {(o.upDownBars ?? (tr.family === 'stock' && tr.open)) && <>
            <Row label={T('Up colour')}><Color label={T('Up bar colour')} value={o.upFill} auto="#12a565" set={(c) => opt('upFill', c)} /></Row>
            <Row label={T('Down colour')}><Color label={T('Down bar colour')} value={o.downFill} auto="#e03131" set={(c) => opt('downFill', c)} /></Row>
          </>}
          <Chk label={T('High-low lines')} on={o.hiLowLines ?? tr.family === 'stock'} set={(v) => opt('hiLowLines', v, N('High-low lines'))} />
        </>}
        {(tr.base === 'line' || tr.base === 'area' || tr.family === 'stock' || tr.combo) && <Chk label={T('Drop lines')} on={!!o.dropLines} set={(v) => opt('dropLines', v || undefined, N('Drop lines'))} />}
        {tr.family === 'cartesian' && stacked && tr.base === 'column' && <Chk label={T('Series lines')} on={!!o.seriesLines} set={(v) => opt('seriesLines', v || undefined, N('Series lines'))} />}
        {(tr.base === 'column' || tr.family === 'bar3d' || tr.family === 'stock') && !pie && <>
          <Row label={T('Gap width %')}><Num label={T('Gap width')} value={o.gapWidth} placeholder={T('Auto')} min={0} max={500} set={(v) => opt('gapWidth', v)} /></Row>
          {tr.family === 'cartesian' && <Row label={T('Overlap %')}><Num label={T('Overlap')} value={o.overlap} placeholder={T('Auto')} min={-100} max={100} set={(v) => opt('overlap', v)} /></Row>}
        </>}
        {(chart.series.length === 1 || pie) && !['treemap', 'sunburst', 'map', 'funnel', 'waterfall', 'histogram', 'pareto', 'box'].includes(tr.family) && <Chk label={T('Vary colours by point')} on={o.varyColors ?? pie} set={(v) => opt('varyColors', v === pie ? undefined : v, N('Vary colours'))} />}
        {pie && <Row label={T('First slice at °')}><Num label={T('First slice angle')} value={o.firstSliceAngle} placeholder="0" min={0} max={360} set={(v) => opt('firstSliceAngle', v)} /></Row>}
        {tr.family === 'pie' && tr.filled && <Row label={T('Hole size %')}><Num label={T('Hole size')} value={o.holeSize} placeholder="52" min={10} max={90} set={(v) => opt('holeSize', v)} /></Row>}
        {tr.family === 'ofPie' && <>
          <Row label={T('Split by')}><Sel label={T('Split series by')} value={o.ofPie?.split ?? 'pos'} set={(v) => opt('ofPie', { ...o.ofPie, split: v })} options={[['pos', N('Position')], ['val', N('Value')], ['percent', N('Percent')]]} /></Row>
          <Row label={o.ofPie?.split === 'val' ? T('Values below') : o.ofPie?.split === 'percent' ? T('Percent below') : T('Last points')}><Num label={T('Split value')} value={o.ofPie?.splitValue} placeholder={T('Auto')} min={0} set={(v) => opt('ofPie', { ...o.ofPie, splitValue: v })} /></Row>
          <Row label={T('Second plot %')}><Num label={T('Second plot size')} value={o.ofPie?.secondSize} placeholder="75" min={5} max={200} set={(v) => opt('ofPie', { ...o.ofPie, secondSize: v })} /></Row>
          <Row label={T('Gap %')}><Num label={T('Gap between plots')} value={o.ofPie?.gapWidth} placeholder="100" min={0} max={500} set={(v) => opt('ofPie', { ...o.ofPie, gapWidth: v })} /></Row>
        </>}
        {tr.family === 'bubble' && <Row label={T('Bubble size %')}><Num label={T('Bubble size')} value={o.bubbleScale} placeholder="100" min={0} max={300} set={(v) => opt('bubbleScale', v)} /></Row>}
        {tr.family === 'map' && <>
          <Row label={T('Map area')}><Sel label={T('Map area')} value={o.map?.area ?? 'auto'} set={(v) => opt('map', { ...o.map, area: v === 'auto' ? undefined : v })} options={[['auto', N('Regions with data')], ['world', N('World')], ['us', N('United States')]]} /></Row>
          <Row label={T('Map labels')}><Sel label={T('Map labels')} value={o.map?.labels ?? 'none'} set={(v) => opt('map', { ...o.map, labels: v === 'none' ? undefined : v })} options={[['none', N('None')], ['best', N('Where they fit')], ['all', N('All')]]} /></Row>
          <Chk label={T('Three colours (low, middle, high)')} on={!!o.map?.diverging} set={(v) => opt('map', { ...o.map, diverging: v || undefined })} />
          <Row label={T('Low colour')}><Color label={T('Low colour')} value={o.map?.low} set={(c) => opt('map', { ...o.map, low: c })} /></Row>
          {o.map?.diverging && <Row label={T('Middle colour')}><Color label={T('Middle colour')} value={o.map?.mid} auto="#f5f1e4" set={(c) => opt('map', { ...o.map, mid: c })} /></Row>}
          <Row label={T('High colour')}><Color label={T('High colour')} value={o.map?.high} set={(c) => opt('map', { ...o.map, high: c })} /></Row>
        </>}
        {tr.family === 'box' && <>
          <Chk label={T('Mean marker')} on={!o.hideMean} set={(v) => opt('hideMean', v ? undefined : true)} />
          <Chk label={T('Outliers')} on={!o.hideOutliers} set={(v) => opt('hideOutliers', v ? undefined : true)} />
        </>}
      </Sec>

      {is3D(chart.type) && !(tr.family === 'surface' && tr.contour) && (
        <Sec title="3D view">
          <Row label={T('Tilt (X) °')}><Num label={T('Tilt')} value={v3.rotX} placeholder={tr.family === 'pie' ? '30' : '15'} min={tr.family === 'pie' ? 10 : -90} max={90} set={(v) => v3up({ rotX: v })} /></Row>
          {tr.family !== 'pie' && <>
            <Row label={T('Turn (Y) °')}><Num label={T('Turn')} value={v3.rotY} placeholder="20" min={0} max={360} set={(v) => v3up({ rotY: v })} /></Row>
            <Row label={T('Perspective')}><Num label={T('Perspective')} value={v3.perspective} placeholder="30" min={0} max={240} set={(v) => v3up({ perspective: v })} /></Row>
            <Chk label={T('Right-angle axes')} on={!!v3.rightAngle} set={(v) => v3up({ rightAngle: v || undefined })} />
            <Row label={T('Depth %')}><Num label={T('Depth')} value={v3.depthPercent} placeholder="100" min={20} max={2000} set={(v) => v3up({ depthPercent: v })} /></Row>
            {(tr.family === 'bar3d' || tr.family === 'line3d' || tr.family === 'area3d') && <Row label={T('Gap depth %')}><Num label={T('Gap depth')} value={v3.gapDepth} placeholder="150" min={0} max={500} set={(v) => v3up({ gapDepth: v })} /></Row>}
          </>}
          <Row label={T('Height %')}><Num label={T('Height')} value={v3.heightPercent} placeholder="100" min={5} max={500} set={(v) => v3up({ heightPercent: v })} /></Row>
          {tr.family === 'bar3d' && <Row label={T('Shape')}><Sel label={T('Shape')} value={o.shape ?? tr.shape} set={(v) => opt('shape', v === tr.shape ? undefined : v)} options={[['box', N('Box')], ['cylinder', N('Cylinder')], ['cone', N('Cone')], ['pyramid', N('Pyramid')]]} /></Row>}
          {tr.family !== 'pie' && <>
            <FillEditor label={T('Walls')} fill={o.walls?.fill} auto="#f2f4f3" set={(f) => opt('walls', f ? { ...o.walls, fill: f } : undefined)} />
            <FillEditor label={T('Floor')} fill={o.floor?.fill} auto="#eceeed" set={(f) => opt('floor', f ? { ...o.floor, fill: f } : undefined)} />
          </>}
        </Sec>
      )}

      <Sec title={T('Chart and plot area')}>
        <ShapeEditor label={T('Chart area')} value={o.chartArea} set={(v) => opt('chartArea', v, N('Chart area'))} />
        <ShapeEditor label={T('Plot area')} value={o.plotArea} set={(v) => opt('plotArea', v, N('Plot area'))} />
      </Sec>

      <Sec title={T('Empty and hidden cells')}>
        <Row label={T('Empty cells')}>
          <Sel label={T('Show empty cells as')} value={o.blanks ?? 'gap'} set={(v) => opt('blanks', v === 'gap' ? undefined : v)} options={[['gap', N('Gaps')], ['zero', N('Zero')], ['span', N('Join with a line')]]} />
        </Row>
        <Chk label={T('Show data in hidden rows and columns')} on={!!o.showHidden} set={(v) => opt('showHidden', v || undefined)} />
      </Sec>
    </div>
  );
};

export const ChartFormatPanel = ({ store }: { store: GridStore }) => {
  const chart = activeChart(store);
  if (!chart) return <p className="muted small">{T('Select a chart to format it.')}{' '}<FileStack size={13} strokeWidth={S} style={{ verticalAlign: 'middle' }} /></p>;
  return <ChartFormat key={chart.id} store={store} chart={chart} />;
};
