import { describe, expect, it } from 'vitest';
import { errorRanges, fitTrend, parseNumberList, polyFit, sampleFit, stdev } from '../src/ui/charts/trend';

const pts = (xs: number[], f: (x: number) => number): [number, number][] => xs.map((x) => [x, f(x)]);
const text = (runs: { t: string; sup?: boolean }[]) => runs.map((r) => (r.sup ? `^${r.t}` : r.t)).join('');

describe('trendlines', () => {
  it('linear least squares with R²', () => {
    const f = fitTrend([[1, 2], [2, 4], [3, 5], [4, 4], [5, 5]], { kind: 'linear' })!;
    expect(f.coef[1]).toBeCloseTo(0.6, 12);
    expect(f.coef[0]).toBeCloseTo(2.2, 12);
    expect(f.r2).toBeCloseTo(0.6, 12);
    expect(text(f.equation)).toBe('y = 0.6x + 2.2');
  });
  it('linear through a set intercept', () => {
    const f = fitTrend([[1, 2], [2, 4], [3, 6]], { kind: 'linear', intercept: 0 })!;
    expect(f.coef).toEqual([0, expect.closeTo(2, 12)]);
    expect(f.r2).toBeCloseTo(1, 12);
    const g = fitTrend([[1, 3], [2, 4], [3, 8]], { kind: 'linear', intercept: 1 })!;
    // slope = Σx(y-1) / Σx² = (2 + 6 + 21) / 14
    expect(g.coef[1]).toBeCloseTo(29 / 14, 12);
    expect(g.at(0)).toBeCloseTo(1, 12);
  });
  it('exponential, logarithmic and power fits recover exact curves', () => {
    const xs = [1, 2, 3, 4, 5, 6];
    const e = fitTrend(pts(xs, (x) => 3 * Math.exp(0.5 * x)), { kind: 'exp' })!;
    expect(e.coef[0]).toBeCloseTo(3, 10); expect(e.coef[1]).toBeCloseTo(0.5, 10); expect(e.r2).toBeCloseTo(1, 10);
    const l = fitTrend(pts(xs, (x) => 2 * Math.log(x) + 1), { kind: 'log' })!;
    expect(l.coef[1]).toBeCloseTo(2, 10); expect(l.coef[0]).toBeCloseTo(1, 10);
    const p = fitTrend(pts(xs, (x) => 4 * x ** 1.5), { kind: 'power' })!;
    expect(p.coef[0]).toBeCloseTo(4, 10); expect(p.coef[1]).toBeCloseTo(1.5, 10);
    expect(text(p.equation)).toBe('y = 4x^1.5');
    expect(text(e.equation)).toBe('y = 3e^0.5x');
  });
  it('exponential R² is measured on ln(y), like a linear fit of the logs', () => {
    const data: [number, number][] = [[1, 3], [2, 7], [3, 20], [4, 55]];
    const f = fitTrend(data, { kind: 'exp' })!;
    const lx = data.map(([x]) => x), ly = data.map(([, y]) => Math.log(y));
    const mx = lx.reduce((a, b) => a + b) / 4, my = ly.reduce((a, b) => a + b) / 4;
    const sxy = lx.reduce((s, x, i) => s + (x - mx) * (ly[i]! - my), 0), sxx = lx.reduce((s, x) => s + (x - mx) ** 2, 0), syy = ly.reduce((s, y) => s + (y - my) ** 2, 0);
    expect(f.coef[1]).toBeCloseTo(sxy / sxx, 12);
    expect(f.coef[0]).toBeCloseTo(Math.exp(my - (sxy / sxx) * mx), 12);
    expect(f.r2).toBeCloseTo((sxy * sxy) / (sxx * syy), 12);
  });
  it('polynomial matches a textbook least-squares example', () => {
    const f = fitTrend([[0, 2.1], [1, 7.7], [2, 13.6], [3, 27.2], [4, 40.9], [5, 61.1]], { kind: 'poly', order: 2 })!;
    expect(f.coef[0]).toBeCloseTo(2.47857, 4);
    expect(f.coef[1]).toBeCloseTo(2.35929, 4);
    expect(f.coef[2]).toBeCloseTo(1.86071, 4);
    expect(f.r2).toBeCloseTo(0.99851, 4);
    expect(text(f.equation)).toBe('y = 1.8607x^2 + 2.3593x + 2.4786');
  });
  it('high orders stay stable on large x values', () => {
    const xs = Array.from({ length: 12 }, (_, i) => 2000 + i);
    const c = polyFit(pts(xs, (x) => 0.001 * (x - 2005) ** 3 - 2 * (x - 2005) + 7), 6)!;
    for (const x of xs) expect(c.reduceRight((a, k) => a * x + k, 0)).toBeCloseTo(0.001 * (x - 2005) ** 3 - 2 * (x - 2005) + 7, 4);
  });
  it('polynomial with a set intercept', () => {
    const f = fitTrend(pts([1, 2, 3, 4], (x) => 1 + 2 * x + 3 * x * x), { kind: 'poly', order: 2, intercept: 1 })!;
    expect(f.coef[0]).toBe(1); expect(f.coef[1]).toBeCloseTo(2, 10); expect(f.coef[2]).toBeCloseTo(3, 10);
  });
  it('moving average', () => {
    const f = fitTrend([[1, 1], [2, 2], [3, 3], [4, 4], [5, 8]], { kind: 'movingAvg', period: 3 })!;
    expect(f.points).toEqual([[3, 2], [4, 3], [5, 5]]);
    expect(sampleFit(f, 0, 10)).toBe(f.points);
  });
  it('skips points a curve can\'t use and refuses too little data', () => {
    expect(fitTrend([[1, -1], [2, 2], [3, 4]], { kind: 'exp' })!.coef[0]).toBeGreaterThan(0);
    expect(fitTrend([[1, 1]], { kind: 'linear' })).toBeNull();
    expect(fitTrend([[1, 1], [2, 2]], { kind: 'poly', order: 3 })).toBeNull();
    expect(fitTrend([[1, 1], [2, 2]], { kind: 'movingAvg', period: 3 })).toBeNull();
  });
});

describe('error bars', () => {
  const v = [2, 4, 4, 4, 5, 5, 7, 9];
  it('fixed, percent and both ends', () => {
    expect(errorRanges([10, null, 20], { kind: 'fixed', value: 2 })).toEqual([[8, 12], null, [18, 22]]);
    expect(errorRanges([50], { kind: 'percent', value: 10, type: 'plus' })).toEqual([[50, 55]]);
    expect(errorRanges([50], { kind: 'percent', value: 10, type: 'minus' })).toEqual([[45, 50]]);
  });
  it('standard deviation is centred on the mean; standard error is per point', () => {
    const sd = stdev(v);
    expect(sd).toBeCloseTo(2.13809, 5);
    const r = errorRanges(v, { kind: 'stdDev', value: 1 });
    expect(r[0]![0]).toBeCloseTo(5 - sd, 10); expect(r[7]![1]).toBeCloseTo(5 + sd, 10);
    const se = errorRanges(v, { kind: 'stdErr' });
    expect(se[0]![1] - 2).toBeCloseTo(sd / Math.sqrt(8), 10);
  });
  it('custom amounts', () => {
    expect(errorRanges([10, 20], { kind: 'custom' }, { plus: [1, 2], minus: [3, null] })).toEqual([[7, 11], [20, 22]]);
    expect(parseNumberList('={1,2;3}')).toEqual([1, 2, 3]);
    expect(parseNumberList('A1:A3')).toBeNull();
  });
});
