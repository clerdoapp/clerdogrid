// Chart editing operations shared by the format panel, dialogs and commands.
import type { GridStore } from '../../app/store';
import { newSheet, type Chart, type ChartOptions, type ChartSeries, type ChartType } from '../../model/workbook';
import { sanitizeOptions, sanitizeSeriesFormat } from '../../formats/chart-options';
import { buildSeries, chartOrientation, chartSourceRange, type Orientation } from './chartData';
import { isChartType, traits } from './chartTypes';
import { selectChart, selectedChart } from './ChartView';
import { readMatrix, uniqueSheetName } from '../dialogs/a/util';
import { N, T } from '../i18n';

export const CHART_SHEET_ANCHOR = { row: 0, col: 0, dx: 0, dy: 0, width: 910, height: 660 };

/** The chart the format panel works on: the selected one, or the chart of a chart sheet. */
export const activeChart = (store: GridStore): Chart | undefined => {
  const sh = store.sheet;
  if (sh.kind === 'chart') return sh.charts[0];
  const id = selectedChart();
  return id ? sh.charts.find((c) => c.id === id) : undefined;
};

/** Changes a chart as one undo step. */
export const editChart = (store: GridStore, id: string, label: string, fn: (c: Chart) => void): void => {
  store.change(label, 'sheet', () => {
    const c = store.sheet.charts.find((x) => x.id === id);
    if (c) { fn(c); tidy(c); }
  });
};

/** Removes empty option objects so saved files and comparisons stay small. */
export const tidy = (c: Chart): void => {
  const clean = (o: Record<string, unknown>): void => {
    for (const k of Object.keys(o)) {
      const v = o[k];
      if (v === undefined) delete o[k];
      else if (v && typeof v === 'object' && !Array.isArray(v)) { clean(v as Record<string, unknown>); if (!Object.keys(v).length) delete o[k]; }
      else if (Array.isArray(v) && !v.length) delete o[k];
    }
  };
  if (c.options) { clean(c.options as Record<string, unknown>); if (!Object.keys(c.options).length) delete c.options; }
  c.series.forEach((s) => clean(s as unknown as Record<string, unknown>));
  if (c.axisTitles) { clean(c.axisTitles as Record<string, unknown>); if (!Object.keys(c.axisTitles).length) delete c.axisTitles; }
};

export const setOption = <K extends keyof ChartOptions>(store: GridStore, id: string, key: K, value: ChartOptions[K] | undefined, label = N('Format chart')): void =>
  editChart(store, id, label, (c) => { c.options = { ...c.options, [key]: value }; });

export const patchSeries = (store: GridStore, id: string, index: number, patch: Partial<ChartSeries>, label = N('Format series')): void =>
  editChart(store, id, label, (c) => { const s = c.series[index]; if (s) Object.assign(s, patch); });

/** Changes the chart type, keeping what still applies. */
export const changeType = (store: GridStore, id: string, type: ChartType): void => editChart(store, id, N('Change chart type'), (c) => {
  c.type = type;
  const tr = traits(type);
  if (!tr.combo) c.series.forEach((s) => { delete s.type; });
  if (tr.family !== 'bubble') c.series.forEach((s) => { delete s.sizes; });
});

/** Swaps series between rows and columns of the source block. */
export const switchRowCol = (store: GridStore, id: string): boolean => {
  const c = store.sheet.charts.find((x) => x.id === id);
  const src = c && chartSourceRange(c);
  if (!c || !src) { store.notify(T('This chart\'s data isn\'t one block of cells, so rows and columns can\'t be switched.')); return false; }
  const si = src.sheet ? store.wb.sheets.findIndex((s) => s.name.toLowerCase() === src.sheet!.toLowerCase()) : store.sheetIndex;
  if (si < 0) return false;
  const name = store.wb.sheets[si]!.name;
  const next: Orientation = chartOrientation(c) === 'cols' ? 'rows' : 'cols';
  const m = readMatrix(store, si, src.range, 2000);
  const series = buildSeries(name, src.range, m, { orientation: next, type: c.type });
  if (!series.length) return false;
  editChart(store, id, N('Switch rows and columns'), (ch) => {
    // Formatting stays with the series position.
    ch.series = series.map((s, i) => {
      const { values: _v, categories: _c, name: _n, sizes: _z, ...fmt } = ch.series[i] ?? { values: '' };
      return { ...fmt, ...s };
    });
  });
  return true;
};

const fixNames = (store: GridStore, map: (i: number) => number): void => {
  for (const n of store.wb.names) if (n.sheet !== undefined) n.sheet = map(n.sheet);
};

/** Moves a chart to a new chart sheet of its own. */
export const moveToChartSheet = (store: GridStore, id: string, name: string): boolean => {
  const from = store.sheetIndex;
  const c = store.wb.sheets[from]!.charts.find((x) => x.id === id);
  if (!c) return false;
  const nm = uniqueSheetName(store, name.trim() || T('Chart'));
  store.change(N('Move chart'), 'book', () => {
    const src = store.wb.sheets[from]!;
    src.charts = src.charts.filter((x) => x !== c);
    const sh = newSheet(nm);
    sh.kind = 'chart';
    sh.view.showGrid = false;
    sh.view.showHeaders = false;
    sh.charts = [{ ...structuredClone(c), anchor: { ...CHART_SHEET_ANCHOR } }];
    const at = from + 1;
    store.wb.sheets.splice(at, 0, sh);
    fixNames(store, (i) => (i >= at ? i + 1 : i));
    store.wb.activeSheet = at;
    store.engine.sheetsChanged();
  });
  selectChart(store, null);
  return true;
};

/** Moves the chart of a chart sheet (or a floating chart) onto a worksheet. */
export const moveToSheet = (store: GridStore, id: string, target: number): boolean => {
  const from = store.sheetIndex;
  const src = store.wb.sheets[from]!;
  const c = src.charts.find((x) => x.id === id);
  const dest = store.wb.sheets[target];
  if (!c || !dest || dest.kind === 'chart' || target === from) return false;
  store.change(N('Move chart'), 'book', () => {
    const copy = structuredClone(c);
    copy.id = `chart-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    copy.anchor = src.kind === 'chart' ? { row: 1, col: 1, dx: 0, dy: 0, width: 480, height: 288 } : { ...c.anchor };
    dest.charts.push(copy);
    if (src.kind === 'chart') {
      store.wb.sheets.splice(from, 1);
      fixNames(store, (i) => (i > from ? i - 1 : i));
      store.wb.activeSheet = target > from ? target - 1 : target;
    } else {
      src.charts = src.charts.filter((x) => x !== c);
      store.wb.activeSheet = target;
    }
    store.engine.sheetsChanged();
  });
  return true;
};

// ---------- Copy as picture ----------

/** Resolves CSS variables in an SVG copy so it draws the same outside the page. */
const inlineVars = (root: Element, from: Element): void => {
  const cs = getComputedStyle(from);
  const resolve = (s: string): string => s.replace(/var\((--[\w-]+)(?:,\s*([^)]+))?\)/g, (_, name: string, fb?: string) => cs.getPropertyValue(name).trim() || (fb ? resolve(fb.trim()) : ''));
  const walk = (el: Element) => {
    const st = el.getAttribute('style');
    if (st && st.includes('var(')) el.setAttribute('style', resolve(resolve(st)));
    for (const a of ['fill', 'stroke', 'stop-color']) { const v = el.getAttribute(a); if (v?.includes('var(')) el.setAttribute(a, resolve(v)); }
    for (const k of Array.from(el.children)) walk(k);
  };
  walk(root);
};

export const chartSvgElement = (id: string): SVGSVGElement | null => document.querySelector(`[data-chart-id="${CSS.escape(id)}"] svg`);

/** Draws the chart's SVG to a PNG at `scale` times its size. */
export const chartPng = async (svg: SVGSVGElement, scale = 2): Promise<Blob> => {
  const box = svg.viewBox.baseVal;
  const w = Math.max(1, Math.round(box.width)), h = Math.max(1, Math.round(box.height));
  const copy = svg.cloneNode(true) as SVGSVGElement;
  inlineVars(copy, svg);
  copy.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  copy.setAttribute('width', String(w));
  copy.setAttribute('height', String(h));
  copy.setAttribute('style', `font-family: 'Hanken Grotesk', system-ui, sans-serif; background: ${getComputedStyle(svg).getPropertyValue('--chart-bg').trim() || '#ffffff'}`);
  const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bg.setAttribute('width', '100%'); bg.setAttribute('height', '100%');
  bg.setAttribute('fill', getComputedStyle(svg).getPropertyValue('--chart-bg').trim() || '#ffffff');
  copy.insertBefore(bg, copy.firstChild);
  const text = new XMLSerializer().serializeToString(copy);
  const url = URL.createObjectURL(new Blob([text], { type: 'image/svg+xml' }));
  try {
    const img = new Image();
    await new Promise<void>((ok, fail) => { img.onload = () => ok(); img.onerror = () => fail(new Error(T('The chart could not be drawn.'))); img.src = url; });
    const canvas = document.createElement('canvas');
    canvas.width = w * scale; canvas.height = h * scale;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error(T('The chart could not be drawn.'));
    ctx.scale(scale, scale);
    ctx.drawImage(img, 0, 0, w, h);
    return await new Promise<Blob>((ok, fail) => canvas.toBlob((b) => (b ? ok(b) : fail(new Error(T('The chart could not be drawn.')))), 'image/png'));
  } finally {
    URL.revokeObjectURL(url);
  }
};

export const copyChartPicture = async (store: GridStore, id: string): Promise<void> => {
  const svg = chartSvgElement(id);
  if (!svg) { store.notify(T('Show the chart on screen to copy it.')); return; }
  try {
    const png = await chartPng(svg);
    if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': png })]);
      store.notify(T('Chart copied as a picture.'));
    } else {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(png);
      a.download = 'chart.png';
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      store.notify(T('Copying pictures isn\'t available here, so the chart was saved as chart.png.'));
    }
  } catch (e) {
    store.notify(e instanceof Error ? e.message : T('The chart could not be copied.'));
  }
};

// ---------- Templates (kept in this browser) ----------

export interface ChartTemplate {
  id: string;
  name: string;
  type: ChartType;
  legend?: Chart['legend'];
  dataLabels?: boolean;
  options?: ChartOptions;
  /** Formatting for the first series, second series and so on. */
  series: Partial<ChartSeries>[];
}

const KEY = 'clerdogrid.chartTemplates.v1';

export const loadTemplates = (): ChartTemplate[] => {
  try {
    const raw = JSON.parse(localStorage.getItem(KEY) ?? '[]') as unknown;
    if (!Array.isArray(raw)) return [];
    return raw.slice(0, 100).flatMap((t): ChartTemplate[] => {
      if (!t || typeof t !== 'object') return [];
      const r = t as Record<string, unknown>;
      if (!isChartType(r.type) || typeof r.name !== 'string' || typeof r.id !== 'string') return [];
      const legend = ['right', 'bottom', 'top', 'left', 'none'].includes(r.legend as string) ? (r.legend as Chart['legend']) : undefined;
      return [{
        id: r.id.slice(0, 64), name: r.name.slice(0, 80), type: r.type, legend, dataLabels: r.dataLabels === true ? true : undefined,
        options: sanitizeOptions(r.options), series: Array.isArray(r.series) ? r.series.slice(0, 64).map(sanitizeSeriesFormat) : [],
      }];
    });
  } catch {
    return [];
  }
};

const saveAll = (list: ChartTemplate[]): boolean => {
  try { localStorage.setItem(KEY, JSON.stringify(list)); return true; } catch { return false; }
};

export const saveTemplate = (store: GridStore, c: Chart, name: string): void => {
  const series = c.series.map(({ values: _v, categories: _c, name: _n, sizes: _s, labels, errorBars, ...fmt }) => ({
    ...fmt,
    ...(labels ? { labels: { ...labels, fromCells: undefined } } : {}),
    ...(errorBars ? { errorBars: errorBars.filter((e) => e.kind !== 'custom') } : {}),
  }));
  const t: ChartTemplate = { id: `t${Date.now().toString(36)}`, name: name.trim() || T('Chart template'), type: c.type, legend: c.legend, dataLabels: c.dataLabels, options: c.options, series };
  const list = [...loadTemplates().filter((x) => x.name !== t.name), t];
  store.notify(saveAll(list) ? T('Saved the template "{name}".', { name: t.name }) : T('The template could not be saved in this browser.'));
};

export const deleteTemplate = (id: string): void => { saveAll(loadTemplates().filter((t) => t.id !== id)); };

/** A chart with a template's look over the given series. */
export const applyTemplate = (t: ChartTemplate, c: Chart): Chart => ({
  ...c, type: t.type, legend: t.legend ?? c.legend, dataLabels: t.dataLabels ?? c.dataLabels,
  ...(t.options ? { options: structuredClone(t.options) } : {}),
  series: c.series.map((s, i) => ({ ...structuredClone(t.series[i] ?? {}), ...s })),
});
