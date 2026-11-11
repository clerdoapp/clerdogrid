import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/statistical';
import { detectSeason, etsSse, fitEts, prepareSeries } from '../src/engine/ets';
import { toSerial } from '../src/engine/date';
import { CellError, type Scalar, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const col = (...xs: Scalar[]): Scalar[][] => xs.map((x) => [x]);
const code = (v: Value): string | null => (v instanceof CellError ? v.code : null);
const n = (v: Value): number => { if (typeof v !== 'number') throw new Error(`not a number: ${String(v)}`); return v; };

// Deterministic noise.
const rng = (seed: number) => () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648 - 0.5; };

// Monthly international airline passengers (thousands), 1949-1951, from Box & Jenkins' classic series.
const AIR = [
  112, 118, 132, 129, 121, 135, 148, 148, 136, 119, 104, 118,
  115, 126, 141, 135, 125, 149, 170, 170, 158, 133, 114, 140,
  145, 150, 178, 163, 172, 178, 199, 199, 184, 162, 146, 166,
];
const AIR_1952 = [171, 180, 193, 181, 183, 218, 230, 242, 209, 191, 172, 194];

describe('FORECAST.ETS', () => {
  it('reproduces a noise-free trend + season exactly', () => {
    const season = [5, -3, 2, -4];
    const t = Array.from({ length: 24 }, (_, i) => i + 1);
    const y = t.map((i) => 100 + 2 * i + season[(i - 1) % 4]!);
    for (let h = 1; h <= 8; h++) {
      const want = 100 + 2 * (24 + h) + season[(23 + h) % 4]!;
      expect(n(call('FORECAST.ETS', 24 + h, col(...y), col(...t)))).toBeCloseTo(want, 6);
    }
    expect(n(call('FORECAST.ETS.STAT', col(...y), col(...t), 6))).toBeLessThan(1e-6);
    expect(n(call('FORECAST.ETS.CONFINT', 30, col(...y), col(...t)))).toBeLessThan(1e-5);
  });

  it('a linear series without a season', () => {
    const t = Array.from({ length: 15 }, (_, i) => i * 7);
    const y = t.map((x) => 3 + 0.5 * x);
    expect(call('FORECAST.ETS.SEASONALITY', col(...y), col(...t))).toBe(1);
    expect(n(call('FORECAST.ETS', 105, col(...y), col(...t)))).toBeCloseTo(55.5, 6);
    expect(n(call('FORECAST.ETS', 101.5, col(...y), col(...t)))).toBeCloseTo(53.75, 6);
    expect(call('FORECAST.ETS.STAT', col(...y), col(...t), 8)).toBe(7);
  });

  it('timeline order does not matter', () => {
    const t = [5, 1, 3, 2, 4, 6, 8, 7];
    const y = t.map((x) => 10 + x);
    expect(n(call('FORECAST.ETS', 9, col(...y), col(...t)))).toBeCloseTo(19, 6);
  });

  it('detects the season of noisy monthly data and forecasts it', () => {
    const r = rng(7);
    const s12 = [10, 6, 2, -3, -8, -10, -7, -2, 1, 4, 3, 4];
    const y = Array.from({ length: 72 }, (_, i) => 200 + 0.8 * i + s12[i % 12]! + 1.5 * r());
    const t = y.map((_, i) => i + 1);
    expect(call('FORECAST.ETS.SEASONALITY', col(...y), col(...t))).toBe(12);
    for (let h = 1; h <= 12; h++) {
      const truth = 200 + 0.8 * (71 + h) + s12[(71 + h) % 12]!;
      expect(Math.abs(n(call('FORECAST.ETS', 72 + h, col(...y), col(...t))) - truth)).toBeLessThan(3);
    }
  });

  it('forecasts the published airline series within a few percent', () => {
    const t = AIR.map((_, i) => toSerial(1949 + Math.floor(i / 12), (i % 12) + 1, 1));
    expect(call('FORECAST.ETS.SEASONALITY', col(...AIR), col(...t))).toBe(12);
    let ape = 0;
    AIR_1952.forEach((actual, k) => {
      const target = toSerial(1952, k + 1, 1);
      const f = n(call('FORECAST.ETS', target, col(...AIR), col(...t)));
      ape += Math.abs(f - actual) / actual;
    });
    expect(ape / 12).toBeLessThan(0.08);
    expect(call('FORECAST.ETS.STAT', col(...AIR), col(...t), 8)).toBe(31);
    const ci = n(call('FORECAST.ETS.CONFINT', toSerial(1952, 1, 1), col(...AIR), col(...t)));
    expect(ci).toBeGreaterThan(1);
    expect(ci).toBeLessThan(40);
  });

  it('explicit seasonality, and seasonality 0 turns it off', () => {
    const season = [4, -4];
    const t = Array.from({ length: 20 }, (_, i) => i);
    const y = t.map((i) => 50 + season[i % 2]!);
    expect(n(call('FORECAST.ETS', 20, col(...y), col(...t), 2))).toBeCloseTo(54, 5);
    const flat = n(call('FORECAST.ETS', 20, col(...y), col(...t), 0));
    expect(Math.abs(flat - 50)).toBeLessThan(4.5);
    expect(code(call('FORECAST.ETS', 20, col(...y), col(...t), 11))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 20, col(...y), col(...t), 8761))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 20, col(...y), col(...t), -1))).toBe('#NUM!');
  });

  it('argument errors', () => {
    const t = [1, 2, 3, 4, 5, 6], y = [1, 2, 3, 4, 5, 6];
    expect(code(call('FORECAST.ETS', 3, col(...y), col(...t)))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 7, col(...y), col(...t.slice(1))))).toBe('#N/A');
    expect(code(call('FORECAST.ETS', 7, col(...y), col(1, 2, 3, 4, 5, 6.5)))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 7, col(...y), col(1, 2, 'x', 4, 5, 6)))).toBe('#VALUE!');
    expect(code(call('FORECAST.ETS', 7, col(...y), col(...t), 1, 2))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 7, col(...y), col(...t), 1, 1, 8))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 7, col(1, 2), col(1, 2)))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS.CONFINT', 7, col(...y), col(...t), 1))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS.CONFINT', 7, col(...y), col(...t), 0))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS.STAT', col(...y), col(...t), 9))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS.STAT', col(...y), col(...t), 0))).toBe('#NUM!');
  });
});

describe('data preparation', () => {
  it('aggregates duplicate timestamps with all seven methods', () => {
    const t = [1, 1, 1, 2, 3, 4];
    const v: Scalar[] = [2, 9, 4, 5, 6, 7];
    const first = (agg: number): number => (prepareSeries(v, t, 1, agg) as { y: number[] }).y[0]!;
    expect(first(0)).toBe(5);
    expect(first(1)).toBe(5);
    expect(first(2)).toBe(3);
    expect(first(3)).toBe(3);
    expect(first(4)).toBe(9);
    expect(first(5)).toBe(4);
    expect(first(6)).toBe(2);
    expect(first(7)).toBe(15);
    const withText = prepareSeries([1, 'x', null, 5, 6, 7], [1, 1, 1, 2, 3, 4], 1, 3) as { y: number[] };
    expect(withText.y[0]).toBe(2);
    expect((prepareSeries([1, 'x', null, 5, 6, 7], [1, 1, 1, 2, 3, 4], 1, 2) as { y: number[] }).y[0]).toBe(1);
  });

  it('completes missing points by interpolation or zeros, up to 30%', () => {
    const t = [1, 2, 4, 5, 6, 7, 8, 9, 10, 11];
    const v = [10, 20, 40, 50, 60, 70, 80, 90, 100, 110];
    expect((prepareSeries(v, t, 1, 0) as { y: number[] }).y[2]).toBe(30);
    expect((prepareSeries(v, t, 0, 0) as { y: number[] }).y[2]).toBe(0);
    const blankValue = prepareSeries([1, null, 3, 4], [1, 2, 3, 4], 1, 0) as { y: number[] };
    expect(blankValue.y).toEqual([1, 2, 3, 4]);
    expect(prepareSeries([1, 2, 3, 4], [1, 2, 3, 10], 1, 0)).toBeInstanceOf(CellError);
    expect(prepareSeries([1, 2, 3], [1, 2, 3], 2, 0)).toBeInstanceOf(CellError);
    // Completion feeds the forecast: a gap in a straight line does not bend it.
    expect(n(call('FORECAST.ETS', 12, col(...v), col(...t)))).toBeCloseTo(120, 5);
  });

  it('understands monthly and month-end date timelines', () => {
    const monthEnds = Array.from({ length: 12 }, (_, i) => toSerial(2023, i + 2, 0));
    const s = prepareSeries(monthEnds.map((_, i) => i), monthEnds, 1, 0);
    expect(s).not.toBeInstanceOf(CellError);
    expect((s as { months: number }).months).toBe(1);
    const quarterly = Array.from({ length: 10 }, (_, i) => toSerial(2020, 1 + 3 * i, 15));
    const y = quarterly.map((_, i) => 100 + 10 * i);
    expect(n(call('FORECAST.ETS', toSerial(2022, 7, 15), col(...y), col(...quarterly), 0))).toBeCloseTo(200, 5);
    expect(n(call('FORECAST.ETS', toSerial(2022, 8, 30), col(...y), col(...quarterly), 0))).toBeCloseTo(200 + 10 * (1 + 15 / 31) / 3, 3);
  });
});

describe('statistics and intervals', () => {
  const r = rng(42);
  const s4 = [8, -2, -9, 3];
  const y = Array.from({ length: 40 }, (_, i) => 50 + 0.3 * i + s4[i % 4]! + 2 * r());
  const t = y.map((_, i) => i + 1);
  const stat = (k: number): number => n(call('FORECAST.ETS.STAT', col(...y), col(...t), k));

  it('all eight statistics are coherent', () => {
    for (const k of [1, 2, 3]) { expect(stat(k)).toBeGreaterThanOrEqual(0); expect(stat(k)).toBeLessThanOrEqual(1); }
    const md = fitEts(y, detectSeason(y));
    const errs = y.map((v, i) => v - md.fitted[i]!);
    expect(stat(6)).toBeCloseTo(errs.reduce((a, e) => a + Math.abs(e), 0) / y.length, 9);
    expect(stat(7)).toBeCloseTo(Math.sqrt(errs.reduce((a, e) => a + e * e, 0) / y.length), 9);
    expect(stat(6)).toBeLessThanOrEqual(stat(7));
    expect(stat(5)).toBeGreaterThan(0);
    expect(stat(5)).toBeLessThan(2);
    expect(stat(4)).toBeGreaterThan(0);
    expect(stat(4)).toBeLessThan(1);
    expect(stat(8)).toBe(1);
  });

  it('the fit beats nearby smoothing weights', () => {
    const md = fitEts(y, 4);
    expect(md.sse).toBeLessThan(fitEts(y, 1).sse);
    expect(etsSse(y, 4, md.alpha, md.beta, md.gamma)).toBeCloseTo(md.sse, 8);
    for (const [da, db, dg] of [[0.05, 0, 0], [-0.05, 0, 0], [0, 0.05, 0], [0, -0.05, 0], [0, 0, 0.05], [0, 0, -0.05]]) {
      const cl = (v: number): number => Math.min(0.999, Math.max(0.001, v));
      expect(etsSse(y, 4, cl(md.alpha + da!), cl(md.beta + db!), cl(md.gamma + dg!))).toBeGreaterThanOrEqual(md.sse * (1 - 1e-9));
    }
  });

  it('intervals widen with the horizon and the confidence level', () => {
    const ci = (h: number, c = 0.95): number => n(call('FORECAST.ETS.CONFINT', 40 + h, col(...y), col(...t), c));
    expect(ci(1)).toBeGreaterThan(0);
    expect(ci(8)).toBeGreaterThan(ci(1));
    expect(ci(4, 0.99)).toBeGreaterThan(ci(4, 0.9));
    expect(ci(4, 0.5)).toBeLessThan(ci(4, 0.9));
    // The one-step half-width is z * RMSE.
    expect(ci(1)).toBeCloseTo(1.959963985 * stat(7), 6);
  });
});
