// Chart settings from untrusted sources (files, saved templates): only known fields with the right
// types and ranges are kept.
import type { ChartOptions, ChartSeries } from '../model/workbook';

type Rule =
  | { t: 'num'; min?: number; max?: number }
  | { t: 'int'; min?: number; max?: number }
  | { t: 'bool' }
  | { t: 'true' }
  | { t: 'false' }
  | { t: 'str'; max?: number }
  | { t: 'color' }
  | { t: 'enum'; v: readonly string[] }
  | { t: 'obj'; f: Record<string, Rule>; req?: string[] }
  | { t: 'arr'; of: Rule; max?: number };

const num = (min?: number, max?: number): Rule => ({ t: 'num', min, max });
const int = (min?: number, max?: number): Rule => ({ t: 'int', min, max });
const bool: Rule = { t: 'bool' };
const str = (max = 255): Rule => ({ t: 'str', max });
const color: Rule = { t: 'color' };
const en = (...v: string[]): Rule => ({ t: 'enum', v });
const obj = (f: Record<string, Rule>, req?: string[]): Rule => ({ t: 'obj', f, req });
const arr = (of: Rule, max = 4096): Rule => ({ t: 'arr', of, max });

const DASH = en('solid', 'dot', 'dash', 'dashDot', 'longDash', 'longDashDot');
const MARKER = en('auto', 'none', 'circle', 'square', 'diamond', 'triangle', 'x', 'star', 'plus', 'dash', 'dot');
const FILL = obj({ color, gradient: obj({ from: color, to: color, angle: num(-360, 360) }), none: bool, transparency: num(0, 100) });
const LINE = obj({ color, width: num(0, 100), dash: DASH, none: bool });
const SHAPE = obj({ fill: FILL, line: LINE });
const MARK = obj({ symbol: MARKER, size: num(2, 72), fill: color, line: color });
const LABELS = obj({
  value: bool, seriesName: bool, categoryName: bool, percent: bool, bubbleSize: bool, legendKey: bool, fromCells: str(1000), leaderLines: bool,
  position: en('auto', 'center', 'insideEnd', 'insideBase', 'outsideEnd', 'left', 'right', 'above', 'below', 'bestFit'), separator: str(20), numFmt: str(),
});
const TICK = en('none', 'out', 'in', 'cross');
const AXIS = obj({
  min: num(), max: num(), majorUnit: num(0), minorUnit: num(0), logBase: num(2, 1000), reverse: bool, crosses: en('auto', 'min', 'max'), crossesAt: num(),
  dispUnits: en('hundreds', 'thousands', 'tenThousands', 'hundredThousands', 'millions', 'tenMillions', 'hundredMillions', 'billions', 'trillions'), dispUnitsLabel: bool,
  numFmt: str(), majorTick: TICK, minorTick: TICK, labelPos: en('nextTo', 'low', 'high', 'none'), majorGrid: bool, minorGrid: bool, dateAxis: bool,
  baseUnit: en('days', 'months', 'years'), hidden: bool, line: LINE,
});

export const OPTION_RULES: Record<keyof ChartOptions, Rule> = {
  secondary: arr(int(0, 99999)),
  binCount: int(1, 1000),
  binWidth: num(Number.MIN_VALUE),
  totals: arr(int(0, 99999)),
  labelMode: en('value', 'percent', 'both'),
  gridlines: { t: 'false' },
  min: num(), max: num(),
  numFmt: str(),
  style: int(1, 31),
  hideMean: { t: 'true' }, hideOutliers: { t: 'true' },
  axes: obj({ x: AXIS, y: AXIS, x2: AXIS, y2: AXIS }),
  view3D: obj({ rotX: num(-90, 90), rotY: num(0, 360), perspective: num(0, 240), rightAngle: bool, depthPercent: num(20, 2000), gapDepth: num(0, 500), heightPercent: num(5, 500) }),
  shape: en('box', 'cylinder', 'cone', 'pyramid'),
  gapWidth: num(0, 500), overlap: num(-100, 100), varyColors: bool, firstSliceAngle: num(0, 360), holeSize: num(10, 90),
  ofPie: obj({ split: en('pos', 'val', 'percent'), splitValue: num(), secondSize: num(5, 200), gapWidth: num(0, 500) }),
  upDownBars: bool, upFill: color, downFill: color, hiLowLines: bool, dropLines: bool, seriesLines: bool,
  dataTable: obj({ keys: bool, horizontal: bool, vertical: bool, outline: bool }),
  labels: LABELS, legendOverlay: bool,
  chartArea: SHAPE, plotArea: SHAPE, walls: SHAPE, floor: SHAPE,
  look: int(0, 31),
  blanks: en('gap', 'zero', 'span'), showHidden: bool,
  map: obj({ area: en('auto', 'world', 'us'), labels: en('none', 'best', 'all'), low: color, mid: color, high: color, diverging: bool }),
  bubbleScale: num(0, 300),
};

const TREND = obj({
  kind: en('linear', 'exp', 'log', 'poly', 'power', 'movingAvg'), order: int(2, 6), period: int(2, 255), forward: num(0), backward: num(0), intercept: num(),
  showEquation: bool, showR2: bool, name: str(), line: LINE,
}, ['kind']);
const ERR = obj({ dir: en('x', 'y'), kind: en('fixed', 'percent', 'stdDev', 'stdErr', 'custom'), value: num(0), plus: str(1000), minus: str(1000), type: en('both', 'plus', 'minus'), noCap: bool, line: LINE }, ['kind']);

/** Formatting fields of a series (not its data references). */
export const SERIES_RULES: Record<string, Rule> = {
  color, type: en('column', 'line', 'area'),
  trendlines: arr(TREND, 8), errorBars: arr(ERR, 4), labels: LABELS, fill: FILL, line: LINE, marker: MARK, smooth: bool, invertIfNegative: bool,
  explosion: num(0, 400), points: arr(obj({ idx: int(0, 99999), fill: FILL, line: LINE, marker: MARK, explosion: num(0, 400) }, ['idx']), 4096),
};

const clean = (v: unknown, r: Rule): unknown => {
  switch (r.t) {
    case 'num': case 'int': {
      if (typeof v !== 'number' || !Number.isFinite(v)) return undefined;
      if (r.min !== undefined && v < r.min) return undefined;
      if (r.max !== undefined && v > r.max) return undefined;
      return r.t === 'int' ? (Number.isInteger(v) ? v : undefined) : v;
    }
    case 'bool': return typeof v === 'boolean' ? v : undefined;
    case 'true': return v === true ? true : undefined;
    case 'false': return v === false ? false : undefined;
    case 'str': return typeof v === 'string' && v.length <= (r.max ?? 255) ? v : undefined;
    case 'color': return typeof v === 'string' && /^#[0-9a-fA-F]{6}$/.test(v) ? v : undefined;
    case 'enum': return typeof v === 'string' && r.v.includes(v) ? v : undefined;
    case 'arr': {
      if (!Array.isArray(v)) return undefined;
      const out = v.slice(0, r.max ?? 4096).map((x) => clean(x, r.of)).filter((x) => x !== undefined);
      return out;
    }
    case 'obj': {
      if (!v || typeof v !== 'object' || Array.isArray(v)) return undefined;
      const src = v as Record<string, unknown>;
      const out: Record<string, unknown> = {};
      for (const [k, rule] of Object.entries(r.f)) {
        if (!Object.prototype.hasOwnProperty.call(src, k)) continue;
        const c = clean(src[k], rule);
        if (c !== undefined && !(typeof c === 'object' && c !== null && !Array.isArray(c) && !Object.keys(c).length)) out[k] = c;
      }
      if (r.req?.some((k) => out[k] === undefined)) return undefined;
      return out;
    }
  }
};

/** Options from untrusted JSON: only known fields with the right types are kept. */
export const sanitizeOptions = (raw: unknown): ChartOptions | undefined => {
  const o = clean(raw, { t: 'obj', f: OPTION_RULES }) as ChartOptions | undefined;
  if (!o) return undefined;
  if (o.secondary && !o.secondary.length) delete o.secondary;
  if (o.binCount !== undefined) o.binCount = Math.min(1000, Math.round(o.binCount));
  return Object.keys(o).length ? o : undefined;
};

/** Series formatting from untrusted JSON (data references are not taken from here). */
export const sanitizeSeriesFormat = (raw: unknown): Partial<ChartSeries> => (clean(raw, { t: 'obj', f: SERIES_RULES }) as Partial<ChartSeries> | undefined) ?? {};
