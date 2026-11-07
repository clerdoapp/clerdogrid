import { fromSerial, todaySerial } from '../engine/date';
import { isError, isMatrix, type Scalar } from '../engine/types';
import type { CellRange, CellStyle, CfRule, CfValue, ConditionalFormat } from '../model/workbook';
import { inRange, type GridStore } from './store';
import { compareOp } from './validation';

export interface CfResult {
  style?: CellStyle;
  bar?: { fraction: number; color: string; gradient: boolean; negative?: boolean; axis?: number };
  fill?: string;
  icon?: { set: string; index: number };
  hideValue?: boolean;
}

interface Stats { nums: number[]; sorted: number[]; min: number; max: number; avg: number; sd: number; counts: Map<string, number> }

const lerp = (a: string, b: string, t: number): string => {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16)), pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  return `#${pa.map((x, i) => Math.round(x + (pb[i]! - x) * t).toString(16).padStart(2, '0')).join('')}`;
};

const percentile = (sorted: number[], p: number): number => {
  if (!sorted.length) return 0;
  const k = (sorted.length - 1) * p;
  const f = Math.floor(k);
  return sorted[f]! + (sorted[Math.min(f + 1, sorted.length - 1)]! - sorted[f]!) * (k - f);
};

export const ICON_SETS: Record<string, number> = {
  '3Arrows': 3, '3ArrowsGray': 3, '3Flags': 3, '3TrafficLights1': 3, '3TrafficLights2': 3, '3Signs': 3, '3Symbols': 3, '3Symbols2': 3, '3Stars': 3, '3Triangles': 3,
  '4Arrows': 4, '4ArrowsGray': 4, '4RedToBlack': 4, '4Rating': 4, '4TrafficLights': 4, '5Arrows': 5, '5ArrowsGray': 5, '5Rating': 5, '5Quarters': 5, '5Boxes': 5,
};

/** Works out conditional formats for one sheet, caching range statistics for a render pass. */
export class CfEvaluator {
  private stats = new Map<ConditionalFormat, Stats>();
  private today = todaySerial();
  constructor(private readonly store: GridStore) {}

  private statsFor(cf: ConditionalFormat): Stats {
    let st = this.stats.get(cf);
    if (st) return st;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    for (const r of cf.ranges) {
      const b = this.store.engine.usedBounds(this.store.sheetIndex);
      const r2 = Math.min(r.r2, b.rows), c2 = Math.min(r.c2, b.cols);
      for (let i = r.r1; i <= r2; i++) for (let j = r.c1; j <= c2; j++) {
        const v = this.store.value(i, j);
        if (typeof v === 'number') nums.push(v);
        if (v !== null && !isError(v)) { const k = typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${String(v)}`; counts.set(k, (counts.get(k) ?? 0) + 1); }
      }
    }
    const sorted = [...nums].sort((a, b) => a - b);
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    const sd = nums.length > 1 ? Math.sqrt(nums.reduce((a, b) => a + (b - avg) ** 2, 0) / (nums.length - 1)) : 0;
    st = { nums, sorted, min: sorted[0] ?? 0, max: sorted.at(-1) ?? 0, avg, sd, counts };
    this.stats.set(cf, st);
    return st;
  }

  private threshold(v: CfValue, st: Stats, cf: ConditionalFormat, isMin: boolean): number {
    switch (v.type) {
      case 'min': case 'autoMin': return st.min;
      case 'max': case 'autoMax': return st.max;
      case 'num': return Number(v.value);
      case 'percent': return st.min + ((st.max - st.min) * Number(v.value)) / 100;
      case 'percentile': return percentile(st.sorted, Number(v.value) / 100);
      case 'formula': {
        const r = cf.ranges[0]!;
        const x = this.store.engine.evaluate(String(v.value ?? '').replace(/^=/, ''), this.store.sheetIndex, r.r1, r.c1);
        const n = isMatrix(x) ? x[0]![0] : x;
        return typeof n === 'number' ? n : isMin ? st.min : st.max;
      }
    }
  }

  private formulaTrue(formula: string, cf: ConditionalFormat, row: number, col: number): boolean {
    // Relative references in the rule are written for the top-left cell of the first range.
    const top = cf.ranges[0]!;
    const shifted = shiftForCell(formula, row - top.r1, col - top.c1);
    const v = this.store.engine.evaluate(shifted, this.store.sheetIndex, row, col);
    const x = isMatrix(v) ? v[0]![0] : v;
    return x === true || (typeof x === 'number' && x !== 0);
  }

  evaluate(row: number, col: number): CfResult | null {
    const s = this.store.sheet;
    if (!s.conditionalFormats.length) return null;
    let out: CfResult = {};
    let hit = false;
    const v = this.store.value(row, col);
    const list = s.conditionalFormats.filter((cf) => cf.ranges.some((r) => inRange(r, row, col)));
    list.sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
    for (const cf of list) {
      for (const rule of cf.rules) {
        const res = this.rule(rule, cf, v, row, col);
        if (!res) continue;
        hit = true;
        out = { ...out, ...res, style: res.style ? { ...res.style, ...out.style } : out.style };
        if ('stop' in rule && rule.stop) return out;
      }
    }
    return hit ? out : null;
  }

  private rule(rule: CfRule, cf: ConditionalFormat, v: Scalar, row: number, col: number): CfResult | null {
    switch (rule.type) {
      case 'cellIs': {
        if (v === null && rule.operator !== 'equal' && rule.operator !== 'notEqual') return null;
        const a = this.evalNum(rule.formulas[0], cf, row, col), b = this.evalNum(rule.formulas[1], cf, row, col);
        if (typeof v === 'number' && typeof a === 'number') return compareOp(rule.operator, v, a, typeof b === 'number' ? b : a) ? { style: rule.style } : null;
        if (typeof v === 'string' && typeof a === 'string') {
          const c = v.toLowerCase().localeCompare(a.toLowerCase());
          return compareOp(rule.operator, c, 0, 0) ? { style: rule.style } : null;
        }
        return null;
      }
      case 'expression': return this.formulaTrue(rule.formula, cf, row, col) ? { style: rule.style } : null;
      case 'colorScale': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(cf);
        const pts = rule.stops.map((s, i) => ({ at: this.threshold(s.value, st, cf, i === 0), color: s.color }));
        if (v <= pts[0]!.at) return { fill: pts[0]!.color };
        for (let i = 1; i < pts.length; i++) {
          const p = pts[i - 1]!, q = pts[i]!;
          if (v <= q.at) return { fill: lerp(p.color, q.color, q.at === p.at ? 1 : (v - p.at) / (q.at - p.at)) };
        }
        return { fill: pts.at(-1)!.color };
      }
      case 'dataBar': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(cf);
        const lo = Math.min(this.threshold(rule.min, st, cf, true), 0), hi = this.threshold(rule.max, st, cf, false);
        const span = hi - lo || 1;
        if (lo < 0) {
          const axis = -lo / span;
          return { bar: { fraction: Math.abs(v) / span, color: v < 0 ? '#ff0000' : rule.color, gradient: rule.gradient !== false, negative: v < 0, axis }, hideValue: rule.showValue === false };
        }
        const minT = this.threshold(rule.min, st, cf, true);
        const frac = hi === minT ? 1 : (v - minT) / (hi - minT);
        return { bar: { fraction: Math.max(0.1, Math.min(1, 0.1 + frac * 0.9)), color: rule.color, gradient: rule.gradient !== false }, hideValue: rule.showValue === false };
      }
      case 'iconSet': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(cf);
        const n = ICON_SETS[rule.set] ?? rule.values.length;
        let idx = 0;
        rule.values.forEach((t, i) => { if (i > 0 && v >= this.threshold(t, st, cf, false)) idx = i; });
        if (rule.reverse) idx = n - 1 - idx;
        return { icon: { set: rule.set, index: idx }, hideValue: rule.showValue === false };
      }
      case 'top10': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(cf);
        const n = rule.percent ? Math.max(1, Math.floor((st.nums.length * rule.rank) / 100)) : rule.rank;
        const cut = rule.bottom ? st.sorted[Math.min(n, st.sorted.length) - 1] : st.sorted[st.sorted.length - Math.min(n, st.sorted.length)];
        if (cut === undefined) return null;
        return (rule.bottom ? v <= cut : v >= cut) ? { style: rule.style } : null;
      }
      case 'aboveAverage': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(cf);
        const limit = st.avg + (rule.stdDev ?? 0) * st.sd * (rule.below ? -1 : 1);
        const ok = rule.below ? (rule.equal ? v <= limit : v < limit) : rule.equal ? v >= limit : v > limit;
        return ok ? { style: rule.style } : null;
      }
      case 'duplicateValues': case 'uniqueValues': {
        if (v === null || isError(v)) return null;
        const k = typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${String(v)}`;
        const n = this.statsFor(cf).counts.get(k) ?? 0;
        return (rule.type === 'duplicateValues' ? n > 1 : n === 1) ? { style: rule.style } : null;
      }
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
        const t = v === null ? '' : String(v).toLowerCase(), q = rule.text.toLowerCase();
        const hit = rule.type === 'beginsWith' ? t.startsWith(q) : rule.type === 'endsWith' ? t.endsWith(q) : t.includes(q);
        return (rule.type === 'notContainsText' ? !hit : hit) ? { style: rule.style } : null;
      }
      case 'containsBlanks': return v === null || (typeof v === 'string' && !v.trim()) ? { style: rule.style } : null;
      case 'notContainsBlanks': return v === null || (typeof v === 'string' && !v.trim()) ? null : { style: rule.style };
      case 'containsErrors': return isError(v) ? { style: rule.style } : null;
      case 'notContainsErrors': return isError(v) ? null : { style: rule.style };
      case 'timePeriod': {
        if (typeof v !== 'number') return null;
        const d = Math.floor(v), t = this.today;
        const wd = (x: number) => (x + 6) % 7; // serial → weekday (0 = Sunday) for modern dates
        const weekStart = t - wd(t);
        const now = fromSerial(t), dv = fromSerial(d);
        const monthDiff = (dv.year - now.year) * 12 + dv.month - now.month;
        const ok = {
          today: d === t, yesterday: d === t - 1, tomorrow: d === t + 1, last7Days: d > t - 7 && d <= t,
          thisWeek: d >= weekStart && d < weekStart + 7, lastWeek: d >= weekStart - 7 && d < weekStart, nextWeek: d >= weekStart + 7 && d < weekStart + 14,
          thisMonth: monthDiff === 0, lastMonth: monthDiff === -1, nextMonth: monthDiff === 1,
        }[rule.period];
        return ok ? { style: rule.style } : null;
      }
    }
  }

  private evalNum(formula: string | undefined, cf: ConditionalFormat, row: number, col: number): Scalar {
    if (formula === undefined) return null;
    const top = cf.ranges[0]!;
    const v = this.store.engine.evaluate(shiftForCell(formula.replace(/^=/, ''), row - top.r1, col - top.c1), this.store.sheetIndex, row, col);
    return isMatrix(v) ? v[0]![0]! : v;
  }
}

import { shiftFormula } from '../engine/refactor';
const shiftCache = new Map<string, string>();
const shiftForCell = (f: string, dr: number, dc: number): string => {
  if (!dr && !dc) return f;
  const k = `${dr},${dc},${f}`;
  let s = shiftCache.get(k);
  if (s === undefined) { s = shiftFormula(f, dr, dc); if (shiftCache.size > 20000) shiftCache.clear(); shiftCache.set(k, s); }
  return s;
};

export const unionRange = (rs: CellRange[]): CellRange => rs.reduce((a, r) => ({ r1: Math.min(a.r1, r.r1), c1: Math.min(a.c1, r.c1), r2: Math.max(a.r2, r.r2), c2: Math.max(a.c2, r.c2) }));
