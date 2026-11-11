// Sparkline groups in a worksheet's extension list (x14:sparklineGroups, ext uri {05C60535-1F16-4fd2-B633-F4F36F0B64E0}).
import { addr, parseRange, type Sheet, type SparklineGroup } from '../model/workbook';
import { attr, esc, find, findAll, kid, kids, numAttr, type XNode } from './xml';

export const SPARKLINE_EXT_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const NS_XM = 'http://schemas.microsoft.com/office/excel/2006/main';

const argb = (c: string | undefined): string | undefined => (c && /^#?[0-9a-fA-F]{6}$/.test(c) ? `FF${c.replace('#', '').toUpperCase()}` : undefined);
const COLORS: [keyof SparklineGroup, string][] = [
  ['color', 'colorSeries'], ['negativeColor', 'colorNegative'], ['axisColor', 'colorAxis'], ['markersColor', 'colorMarkers'],
  ['firstColor', 'colorFirst'], ['lastColor', 'colorLast'], ['highColor', 'colorHigh'], ['lowColor', 'colorLow'],
];

/** The ext element for a sheet's sparklines, or '' when it has none. */
export const sparklinesExtXml = (s: Sheet): string => {
  const groups = (s.sparklines ?? []).filter((g) => g.sparklines.length);
  if (!groups.length) return '';
  const groupXml = (g: SparklineGroup): string => {
    const a: Record<string, string | number | undefined> = {
      manualMax: g.maxAxis === 'custom' ? g.manualMax : undefined,
      manualMin: g.minAxis === 'custom' ? g.manualMin : undefined,
      lineWeight: g.lineWeight !== undefined && g.lineWeight !== 0.75 ? g.lineWeight : undefined,
      type: g.kind === 'line' ? undefined : g.kind === 'column' ? 'column' : 'stacked',
      dateAxis: g.dateRange ? 1 : undefined,
      displayEmptyCellsAs: g.displayEmptyCellsAs ?? 'gap',
      markers: g.markers ? 1 : undefined, high: g.high ? 1 : undefined, low: g.low ? 1 : undefined, first: g.first ? 1 : undefined, last: g.last ? 1 : undefined,
      negative: g.negative ? 1 : undefined, displayXAxis: g.showAxis ? 1 : undefined, displayHidden: g.displayHidden ? 1 : undefined,
      minAxisType: g.minAxis && g.minAxis !== 'individual' ? g.minAxis : undefined,
      maxAxisType: g.maxAxis && g.maxAxis !== 'individual' ? g.maxAxis : undefined,
      rightToLeft: g.rightToLeft ? 1 : undefined,
    };
    const attrs = Object.entries(a).filter(([, v]) => v !== undefined).map(([k, v]) => ` ${k}="${esc(String(v))}"`).join('');
    const colors = COLORS.map(([k, tag]) => { const c = argb(g[k] as string | undefined); return c ? `<x14:${tag} rgb="${c}"/>` : ''; }).join('');
    const date = g.dateRange ? `<xm:f>${esc(g.dateRange.replace(/^=/, ''))}</xm:f>` : '';
    const sps = g.sparklines.map((sp) => `<x14:sparkline><xm:f>${esc(sp.ref.replace(/^=/, ''))}</xm:f><xm:sqref>${addr(sp.row, sp.col)}</xm:sqref></x14:sparkline>`).join('');
    return `<x14:sparklineGroup${attrs}>${colors}${date}<x14:sparklines>${sps}</x14:sparklines></x14:sparklineGroup>`;
  };
  return `<ext uri="${SPARKLINE_EXT_URI}" xmlns:x14="${NS_X14}"><x14:sparklineGroups xmlns:xm="${NS_XM}">${groups.map(groupXml).join('')}</x14:sparklineGroups></ext>`;
};

const flag = (n: XNode, k: string): boolean => { const v = attr(n, k); return v === '1' || v === 'true'; };
const color = (n: XNode | undefined): string | undefined => {
  const rgb = attr(n, 'rgb');
  return rgb && /^[0-9A-Fa-f]{8}$/.test(rgb) ? '#' + rgb.slice(2).toUpperCase() : rgb && /^[0-9A-Fa-f]{6}$/.test(rgb) ? '#' + rgb.toUpperCase() : undefined;
};

/** Sparkline groups from a worksheet's extLst. Theme and indexed colours fall back to the defaults. */
export const readSparklines = (ws: XNode): SparklineGroup[] => {
  const ext = kids(kid(ws, 'extLst'), 'ext').find((e) => attr(e, 'uri') === SPARKLINE_EXT_URI);
  const out: SparklineGroup[] = [];
  for (const g of findAll(find(ext, 'sparklineGroups'), 'sparklineGroup').slice(0, 10000)) {
    const type = attr(g, 'type');
    const grp: SparklineGroup = { kind: type === 'column' ? 'column' : type === 'stacked' ? 'winLoss' : 'line', sparklines: [] };
    for (const [k, tag] of COLORS) { const c = color(kid(g, tag)); if (c) (grp as unknown as Record<string, unknown>)[k] = c; }
    const lw = numAttr(g, 'lineWeight');
    grp.lineWeight = lw !== undefined && lw > 0 && lw <= 100 ? lw : 0.75;
    for (const k of ['markers', 'high', 'low', 'first', 'last', 'negative'] as const) if (flag(g, k)) grp[k] = true;
    if (flag(g, 'displayXAxis')) grp.showAxis = true;
    if (flag(g, 'displayHidden')) grp.displayHidden = true;
    if (flag(g, 'rightToLeft')) grp.rightToLeft = true;
    const empty = attr(g, 'displayEmptyCellsAs');
    grp.displayEmptyCellsAs = empty === 'zero' || empty === 'span' ? empty : empty === 'gap' ? 'gap' : 'zero';
    for (const [k, mk] of [['minAxis', 'minAxisType'], ['maxAxis', 'maxAxisType']] as const) {
      const v = attr(g, mk);
      if (v === 'group' || v === 'custom') grp[k] = v;
    }
    const mn = numAttr(g, 'manualMin'), mx = numAttr(g, 'manualMax');
    if (grp.minAxis === 'custom' && mn !== undefined && Number.isFinite(mn)) grp.manualMin = mn;
    if (grp.maxAxis === 'custom' && mx !== undefined && Number.isFinite(mx)) grp.manualMax = mx;
    const date = kid(g, 'f')?.text.trim();
    if (date && flag(g, 'dateAxis')) grp.dateRange = date;
    for (const sp of kids(kid(g, 'sparklines'), 'sparkline')) {
      const ref = kid(sp, 'f')?.text.trim();
      const loc = parseRange(kid(sp, 'sqref')?.text.trim().replace(/\$/g, '') ?? '');
      if (ref && loc) grp.sparklines.push({ ref, row: loc.r1, col: loc.c1 });
    }
    if (grp.sparklines.length) out.push(grp);
  }
  return out;
};
