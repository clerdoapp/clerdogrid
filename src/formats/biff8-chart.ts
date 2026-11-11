// Embedded charts in the 1997-2003 binary format (MS-XLS chart substream): reading into the Chart
// model, and writing a chart substream with its series, groups, axes, legend and titles.
import type { Chart, ChartSeries, ChartType } from '../model/workbook';

export interface ChartRec { type: number; data: Uint8Array }

const u16 = (d: Uint8Array, o: number): number => (d[o] ?? 0) | ((d[o + 1] ?? 0) << 8);

/** Text of a ShortXLUnicodeString-like run: count byte, flags byte, characters. */
const shortText = (d: Uint8Array, o: number): string => {
  const n = d[o] ?? 0, hi = (d[o + 1] ?? 0) & 1;
  let s = '';
  for (let i = 0; i < n; i++) s += String.fromCharCode(hi ? u16(d, o + 2 + i * 2) : d[o + 2 + i] ?? 0);
  return s;
};

interface Group { kind: string; flags: number; donut: number; axis: number; bubbles: boolean }

/** Reads one chart substream (from its BOF to its EOF). `ref` turns a reference formula into text. */
export const readXlsChart = (recs: ChartRec[], ref: (rgce: Uint8Array) => string | undefined): Omit<Chart, 'id' | 'anchor'> | null => {
  const series: (ChartSeries & { group?: number; lines?: boolean; markers?: boolean; order?: number; explode?: boolean; cached?: string })[] = [];
  const groups: Group[] = [];
  let title: string | undefined, xTitle: string | undefined, yTitle: string | undefined;
  let legend: Chart['legend'] = 'none';
  let dataLabels = false;
  let axis = 0;
  let depth = 0;
  // what the current Begin/End block belongs to
  const stack: string[] = [];
  let last = '';
  let cur: (typeof series)[number] | null = null;
  let text: { link?: number; value?: string } | null = null;
  let textDepth = -1;
  let fmt: { series: number } | null = null;
  let defaults: { lines?: boolean; markers?: boolean } | null = null;
  let groupLines: boolean | undefined, groupMarkers: boolean | undefined;
  let titleAuto = false;
  let fmtDepth = -1;

  for (const r of recs) {
    const d = r.data;
    switch (r.type) {
      case 0x1033: stack.push(last); depth++; break;
      case 0x1034: {
        depth--;
        const owner = stack.pop();
        if (owner === 'series') cur = null;
        if (text && depth === textDepth) {
          if (text.value === undefined && text.link === 1) titleAuto = true;
          if (text.value !== undefined) {
            if (text.link === 1) title = text.value;
            else if (text.link === 2) yTitle = text.value;
            else if (text.link === 3) xTitle = text.value;
          }
          text = null; textDepth = -1;
        }
        if ((fmt || defaults) && depth === fmtDepth) { fmt = null; defaults = null; fmtDepth = -1; }
        break;
      }
      case 0x1003: cur = { values: '' }; series.push(cur); last = 'series'; continue;
      case 0x1051: {
        const id = d[0] ?? 0, rt = d[1] ?? 0, cce = u16(d, 6);
        const f = rt === 2 && cce ? ref(d.subarray(8, 8 + cce)) : undefined;
        if (text && depth === textDepth + 1 && id === 0 && f) text.value = `=${f}`;
        else if (cur && stack.at(-1) === 'series' && f) {
          if (id === 0) cur.name = f;
          else if (id === 1) cur.values = f;
          else if (id === 2) cur.categories = f;
          else if (id === 3) cur.sizes = f;
        }
        break;
      }
      case 0x100d: {
        const s = shortText(d, 2);
        if (text && depth === textDepth + 1) text.value ??= s;
        else if (cur && stack.at(-1) === 'series') { cur.cached = s; if (!cur.name) cur.name = s; }
        break;
      }
      case 0x1025: text = {}; textDepth = depth; last = 'text'; continue;
      case 0x1027: if (text) text.link = u16(d, 0); break;
      case 0x1045: if (cur) cur.group = u16(d, 0); break;
      case 0x1041: axis = u16(d, 0); break;
      case 0x1014: groups.push({ kind: '', flags: 0, donut: 0, axis, bubbles: false }); last = 'group'; continue;
      case 0x1017: case 0x1018: case 0x1019: case 0x101a: case 0x101b: case 0x103e: case 0x1040: case 0x103f: case 0x1061: {
        const g = groups.at(-1);
        if (!g || g.kind) break;
        g.kind = ({ 0x1017: 'bar', 0x1018: 'line', 0x1019: 'pie', 0x101a: 'area', 0x101b: 'scatter', 0x103e: 'radar', 0x1040: 'radarArea', 0x103f: 'surface', 0x1061: 'pie' } as Record<number, string>)[r.type]!;
        if (r.type === 0x1017) g.flags = u16(d, 4);
        else if (r.type === 0x1018 || r.type === 0x101a) g.flags = u16(d, 0);
        else if (r.type === 0x1019) g.donut = u16(d, 2);
        else if (r.type === 0x101b) g.bubbles = !!(u16(d, 4) & 1);
        break;
      }
      case 0x1015: {
        const t = d[16];
        legend = t === 0 ? 'bottom' : t === 2 ? 'top' : t === 4 ? 'left' : 'right';
        break;
      }
      case 0x1006: {
        // DataFormat: a whole series (point index 0xFFFF) gets its color from the formats that follow
        if (u16(d, 0) === 0xffff) { fmt = { series: u16(d, 2) }; fmtDepth = depth; const t = series[u16(d, 2)]; if (t) t.order = u16(d, 4); }
        else if (u16(d, 4) === 0xfffd) { defaults = {}; fmtDepth = depth; }
        last = 'format';
        continue;
      }
      case 0x100a: {
        const s = fmt ? series[fmt.series] : undefined;
        if (s && !(u16(d, 10) & 1) && u16(d, 8) === 1) s.color = hexOf(d, 0);
        break;
      }
      case 0x1007: {
        const s = fmt ? series[fmt.series] : undefined;
        if (defaults) { groupLines = u16(d, 4) !== 5; break; }
        if (s) {
          if (u16(d, 4) === 5) s.lines = false;
          else if (!(u16(d, 8) & 1) && !s.color) s.color = hexOf(d, 0);
        }
        break;
      }
      case 0x1009: { if (defaults) { groupMarkers = u16(d, 8) !== 0; break; } const s = fmt ? series[fmt.series] : undefined; if (s) s.markers = u16(d, 8) !== 0; break; }
      case 0x100c: if (u16(d, 0) & 1) dataLabels = true; break;
      case 0x100b: { const s = fmt ? series[fmt.series] : undefined; if (s && u16(d, 0) > 0) s.explode = true; break; }
      default: break;
    }
    last = '';
  }
  const used = series.filter((s) => s.values).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  for (const s of used) { s.lines ??= groupLines; s.markers ??= groupMarkers; }
  if (titleAuto && !title && used.length === 1) title = used[0]!.cached ?? used[0]!.name;
  if (!used.length || !groups.length) return null;

  const kindOf = (g: Group, ss: typeof used): ChartType => {
    const stacked = !!(g.flags & (g.kind === 'bar' ? 2 : 1)), percent = !!(g.flags & (g.kind === 'bar' ? 4 : 2));
    const markers = ss.some((s) => s.markers !== false);
    switch (g.kind) {
      case 'bar': {
        const base = g.flags & 1 ? 'Bar' : 'Column';
        return (percent ? `percent${base}` : stacked ? `stacked${base}` : base.toLowerCase()) as ChartType;
      }
      case 'line': return (percent ? (markers ? 'percentLineMarkers' : 'percentLine') : stacked ? (markers ? 'stackedLineMarkers' : 'stackedLine') : markers ? 'lineMarkers' : 'line');
      case 'area': return percent ? 'percentArea' : stacked ? 'stackedArea' : 'area';
      case 'pie': return g.donut ? (ss.some((s) => s.explode) ? 'doughnutExploded' : 'doughnut') : ss.some((s) => s.explode) ? 'pieExploded' : 'pie';
      case 'scatter': return g.bubbles ? 'bubble' : ss.some((s) => s.lines !== false) ? 'scatterLinesMarkers' : 'scatter';
      case 'radar': return markers ? 'radarMarkers' : 'radar';
      case 'radarArea': return 'radarFilled';
      default: return 'line';
    }
  };
  const inGroup = (gi: number) => used.filter((s) => (s.group ?? 0) === gi);
  let type = kindOf(groups[0]!, inGroup(0));
  const out: Omit<Chart, 'id' | 'anchor'> = { type, series: [], legend };
  if (groups.length > 1) {
    const second = groups.findIndex((_, i) => i > 0 && inGroup(i).length);
    if (second > 0) {
      const g2 = groups[second]!;
      const simple = (g: Group) => (g.kind === 'bar' ? 'column' : g.kind === 'area' ? 'area' : 'line') as 'column' | 'line' | 'area';
      type = g2.axis === 1 ? 'comboSecondary' : groups[0]!.kind === 'area' ? 'comboAreaColumn' : 'combo';
      out.type = type;
      used.forEach((s) => { s.type = simple(groups[s.group ?? 0] ?? groups[0]!); });
      const sec = used.map((s, i) => (groups[s.group ?? 0]?.axis === 1 ? i : -1)).filter((i) => i >= 0);
      if (sec.length) out.options = { secondary: sec };
    }
  }
  out.series = used.map((s) => {
    const o: ChartSeries = { values: s.values };
    if (s.name) o.name = s.name;
    if (s.categories) o.categories = s.categories;
    if (s.sizes && type === 'bubble') o.sizes = s.sizes;
    if (s.color) o.color = s.color;
    if (s.type && out.type.startsWith('combo')) o.type = s.type;
    return o;
  });
  if (title) out.title = title;
  if (xTitle || yTitle) out.axisTitles = { ...(xTitle ? { x: xTitle } : {}), ...(yTitle ? { y: yTitle } : {}) };
  if (dataLabels) out.dataLabels = true;
  return out;
};

const hexOf = (d: Uint8Array, o: number): string => '#' + [d[o], d[o + 1], d[o + 2]].map((x) => (x ?? 0).toString(16).padStart(2, '0')).join('').toUpperCase();

// ---------------------------------------------------------------- writing

class W {
  b: number[] = [];
  u8(v: number): this { this.b.push(v & 0xff); return this; }
  u16(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff); return this; }
  u32(v: number): this { this.b.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); return this; }
  f64(v: number): this { this.b.push(...new Uint8Array(new Float64Array([v]).buffer)); return this; }
  bytes(a: ArrayLike<number>): this { for (let i = 0; i < a.length; i++) this.b.push(a[i]!); return this; }
  zeros(n: number): this { for (let i = 0; i < n; i++) this.b.push(0); return this; }
  text(s: string): this {
    const t = s.slice(0, 255);
    const hi = /[^\x00-\xff]/.test(t);
    this.u8(t.length).u8(hi ? 1 : 0);
    for (let i = 0; i < t.length; i++) { if (hi) this.u16(t.charCodeAt(i)); else this.u8(t.charCodeAt(i)); }
    return this;
  }
  i16raw(v: number): this { return this.u16(v & 0xffff); }
  rgb(c: string | undefined): this { const v = parseInt((c ?? '#000000').slice(1, 7), 16) || 0; return this.u8(v >> 16).u8((v >> 8) & 0xff).u8(v & 0xff).u8(0); }
}

const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

/** Records of a chart substream for `c`, from its BOF to its EOF. `ref` compiles a reference to tokens (or null). */
export const xlsChartRecords = (c: Chart, ref: (text: string) => Uint8Array | null): ChartRec[] => {
  const out: ChartRec[] = [];
  const rec = (type: number, w: W | number[] = []) => out.push({ type, data: new Uint8Array(w instanceof W ? w.b : w) });
  const begin = () => rec(0x1033), end = () => rec(0x1034);
  const kind = chartKind(c.type);
  const brai = (id: number, f: string | undefined, literal?: boolean) => {
    const t = f ? ref(f.replace(/^=/, '')) : null;
    if (t) rec(0x1051, new W().u8(id).u8(2).u16(0).u16(0).u16(t.length).bytes(t));
    else rec(0x1051, new W().u8(id).u8(literal ? 1 : 0).u16(0).u16(0).u16(0));
  };
  const textBlock = (link: number, value: string) => {
    const isRef = value.startsWith('=');
    rec(0x1025, new W().u8(2).u8(2).u16(1).rgb('#000000').u32(0).u32(0).u32(0).u32(0).u16(0x0081).u16(0x004d).u16(0).u16(link === 2 ? 90 : 0));
    begin();
    rec(0x104f, new W().u16(2).u16(2).u32(0).u32(0).u32(0).u32(0));
    if (isRef) brai(0, value);
    else { brai(0, undefined, true); rec(0x100d, new W().u16(0).text(value)); }
    rec(0x1027, new W().u16(link).u16(0).u16(0));
    end();
  };

  rec(0x0809, new W().u16(0x0600).u16(0x0020).u16(0x0dbb).u16(0x07cc).u32(0x000100c1).u32(0x00000406));
  rec(0x0014); rec(0x0015);
  rec(0x0083, new W().u16(0)); rec(0x0084, new W().u16(0));
  rec(0x00a1, new W().u16(9).u16(100).u16(1).u16(1).u16(1).u16(2).u16(600).u16(600).f64(0.5).f64(0.5).u16(1));
  rec(0x1001, new W().u16(0));
  rec(0x1002, new W().u32(0).u32(0).u32(Math.round(c.anchor.width * 0.75) << 16).u32(Math.round(c.anchor.height * 0.75) << 16));
  begin();
  rec(0x1064, new W().u32(0x10000).u32(0x10000));
  rec(0x1032, new W().u16(0).u16(3)); begin();
  rec(0x1007, new W().rgb('#000000').u16(5).i16raw(-1).u16(0x0008).u16(0x4d));
  rec(0x100a, new W().rgb('#FFFFFF').rgb('#000000').u16(1).u16(0x0001).u16(0x4e).u16(0x4d)); end();

  const secondary = new Set(c.options?.secondary ?? []);
  const groupOf = (i: number): number => (kind.combo ? (secondary.has(i) || (c.type !== 'comboSecondary' && c.series[i]?.type && c.series[i]!.type !== (c.type === 'comboAreaColumn' ? 'area' : 'column')) ? 1 : 0) : 0);
  c.series.forEach((s, i) => {
    const pts = (f: string | undefined) => cellCount(f);
    rec(0x1003, new W().u16(1).u16(1).u16(pts(s.categories) || pts(s.values)).u16(pts(s.values)).u16(1).u16(kind.bubbles ? pts(s.sizes) : 0));
    begin();
    const nameRef = s.name && /!/.test(s.name) ? s.name : undefined;
    brai(0, nameRef, !!s.name && !nameRef);
    if (s.name && !nameRef) rec(0x100d, new W().u16(0).text(s.name));
    brai(1, s.values);
    brai(2, s.categories);
    brai(3, kind.kind === 'scatter' && kind.bubbles ? s.sizes : undefined);
    const color = s.color ?? PALETTE[i % PALETTE.length]!;
    rec(0x1006, new W().u16(0xffff).u16(i).u16(i).u16(0)); begin();
    const lineish = kind.kind === 'line' || kind.kind === 'radar' || (kind.kind === 'scatter' && !kind.bubbles) || (kind.combo && s.type === 'line');
    rec(0x1007, new W().rgb(lineish ? color : '#000000').u16(lineish && kind.noLines ? 5 : 0).i16raw(lineish ? 1 : -1).u16(lineish ? 0 : 0x0008).u16(0x4d));
    rec(0x100a, new W().rgb(color).rgb('#FFFFFF').u16(1).u16(0).u16(0x4d).u16(0x4e));
    if (kind.kind === 'pie') rec(0x100b, new W().u16(0));
    if (lineish || kind.kind === 'scatter') rec(0x1009, new W().rgb(color).rgb(color).u16(kind.noMarkers ? 0 : 2).u16(0).u16(0x4d).u16(0x4d).u32(100));
    if (c.dataLabels) rec(0x100c, new W().u16(1));
    end();
    rec(0x1045, new W().u16(groupOf(i)));
    end();
  });

  rec(0x1044, new W().u16(0x000a).u8(0).u8(0));
  if (c.title) textBlock(1, c.title);
  rec(0x1046, new W().u16(kind.combo && c.type === 'comboSecondary' ? 2 : 1));

  const chartGroup = (g: number): void => {
    rec(0x1014, new W().zeros(16).u16(0).u16(g)); begin();
    const k = g === 0 ? kind : chartKind((c.series.find((_, i) => groupOf(i) === 1)?.type ?? 'line') as ChartType);
    switch (k.kind) {
      case 'bar': rec(0x1017, new W().u16(k.stacked || k.percent ? 0xff9c : 0).u16(150).u16((k.horizontal ? 1 : 0) | (k.stacked || k.percent ? 2 : 0) | (k.percent ? 4 : 0))); break;
      case 'line': rec(0x1018, new W().u16((k.stacked || k.percent ? 1 : 0) | (k.percent ? 2 : 0))); break;
      case 'area': rec(0x101a, new W().u16((k.stacked || k.percent ? 1 : 0) | (k.percent ? 2 : 0))); break;
      case 'pie': rec(0x1019, new W().u16(0).u16(k.donut ? 50 : 0).u16(0)); break;
      case 'scatter': rec(0x101b, new W().u16(100).u16(1).u16(k.bubbles ? 1 : 0)); break;
      case 'radar': rec(0x103e, new W().u16(1).u16(0)); break;
      case 'radarArea': rec(0x1040, new W().u16(1).u16(0)); break;
    }
    if (g === 0 && c.legend !== 'none') {
      const pos = c.legend === 'bottom' ? 0 : c.legend === 'top' ? 2 : c.legend === 'left' ? 4 : 3;
      rec(0x1015, new W().u32(0).u32(0).u32(0).u32(0).u8(pos).u8(1).u16(0x001f));
      begin();
      rec(0x104f, new W().u16(5).u16(2).u32(0).u32(0).u32(0).u32(0));
      rec(0x1025, new W().u8(2).u8(2).u16(1).rgb('#000000').u32(0).u32(0).u32(0).u32(0).u16(0x00b1).u16(0x004d).u16(0).u16(0));
      begin(); rec(0x104f, new W().u16(2).u16(2).u32(0).u32(0).u32(0).u32(0)); end();
      end();
    }
    end();
  };
  const axisParent = (g: number, formats: number[]): void => {
    rec(0x1041, new W().u16(g).u32(0).u32(0).u32(0).u32(0));
    begin();
    const hasAxes = kind.kind !== 'pie';
    if (hasAxes && (g === 0 || c.type === 'comboSecondary')) {
      rec(0x104f, new W().u16(2).u16(2).u32(0).u32(0).u32(0).u32(0));
      const axisBlock = (type: number) => {
        rec(0x101d, new W().u16(type).zeros(16)); begin();
        if (type === 0 && !(kind.kind === 'scatter')) rec(0x1020, new W().u16(1).u16(1).u16(1).u16(1));
        else rec(0x101f, new W().f64(0).f64(0).f64(0).f64(0).f64(0).u16(0x001f));
        rec(0x101e, new W().u8(2).u8(0).u8(3).u8(1).rgb('#000000').zeros(16).u16(0x0023).u16(0x4d).u16(0));
        rec(0x1021, new W().u16(0));
        rec(0x1007, new W().rgb('#000000').u16(0).u16(0xffff).u16(0x0009).u16(0x4d));
        if (type === 1 && c.options?.gridlines !== false && g === 0) { rec(0x1021, new W().u16(1)); rec(0x1007, new W().rgb('#D9D9D9').u16(0).u16(0xffff).u16(0x0008).u16(0x16)); }
        end();
      };
      axisBlock(0);
      axisBlock(1);
      if (g === 0) {
        if (c.axisTitles?.x) textBlock(3, c.axisTitles.x);
        if (c.axisTitles?.y) textBlock(2, c.axisTitles.y);
        rec(0x1035, []);
        rec(0x1032, new W().u16(0).u16(3)); begin();
        rec(0x1007, new W().rgb('#808080').u16(5).u16(0xffff).u16(0x0009).u16(0x17));
        rec(0x100a, new W().rgb('#FFFFFF').rgb('#000000').u16(0).u16(0x0003).u16(0x4e).u16(0x4d)); end();
      }
    }
    formats.forEach(chartGroup);
    end();
  };
  if (kind.combo && c.type !== 'comboSecondary') axisParent(0, [0, 1]);
  else { axisParent(0, [0]); if (kind.combo) axisParent(1, [1]); }
  end();
  rec(0x000a);
  return out;
};

/** Number of cells in a reference such as Sheet1!$B$2:$B$6 (0 when it can't be told). */
const cellCount = (f: string | undefined): number => {
  const m = f && /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(f.replace(/^=/, ''));
  if (!m) return 0;
  const col = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const r = m[4] ? Math.abs(+m[4] - +m[2]!) + 1 : 1, c = m[3] ? Math.abs(col(m[3]) - col(m[1]!)) + 1 : 1;
  return Math.min(65535, r * c);
};

interface Kind { kind: string; stacked?: boolean; percent?: boolean; horizontal?: boolean; donut?: boolean; bubbles?: boolean; combo?: boolean; noMarkers?: boolean; noLines?: boolean }

const chartKind = (t: ChartType): Kind => {
  const s = String(t);
  if (s.startsWith('combo')) return { ...chartKind(t === 'comboAreaColumn' ? 'area' : 'column'), combo: true };
  if (/bar/i.test(s) && !/^bar3D/.test(s)) return { kind: 'bar', horizontal: true, stacked: /^stacked/.test(s), percent: /^percent/.test(s) };
  if (/column|^bar3D|cylinder|cone|pyramid|histogram|pareto|waterfall|funnel|boxWhisker/i.test(s)) return { kind: 'bar', horizontal: /^bar3D|funnel/.test(s), stacked: /Stacked|^stacked/.test(s), percent: /Percent|^percent/.test(s) };
  if (/area/i.test(s)) return { kind: 'area', stacked: /stacked/i.test(s), percent: /percent/i.test(s) };
  if (s === 'pie' || s === 'doughnut' || /pie/i.test(s) || s === 'treemap' || s === 'sunburst') return { kind: 'pie', donut: s === 'doughnut' || s === 'sunburst' };
  if (s === 'bubble') return { kind: 'scatter', bubbles: true };
  if (/^scatter/.test(s)) return { kind: 'scatter', noLines: s === 'scatter', noMarkers: s === 'scatterLines' || s === 'scatterSmooth' };
  if (s === 'radarFilled') return { kind: 'radarArea' };
  if (/^radar/.test(s)) return { kind: 'radar', noMarkers: s === 'radar' };
  return { kind: 'line', stacked: /^stacked/.test(s), percent: /^percent/.test(s), noMarkers: !/Markers$/.test(s) && /^(stacked|percent)?[lL]ine$/.test(s) };
};
