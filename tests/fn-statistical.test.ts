import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/statistical';
import { CellError, Err, ref, type Matrix, type Scalar, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx([[
  [10, 'Not available', true, null],
  [7, 'x', false, null],
  [9, null, 3, null],
  [2, 5, '4', null],
]]);
const call = (name: string, ...args: Value[]): Value => {
  const d = defs.find((x) => x.name === name);
  if (!d) throw new Error(`missing ${name}`);
  return d.fn(args, ctx);
};
const col = (...xs: Scalar[]): Matrix => xs.map((x) => [x]);
const code = (v: Value): string | null => (v instanceof CellError ? v.code : null);

const NAMES = `AVEDEV AVERAGE AVERAGEA BETA.DIST BETA.INV BETADIST BETAINV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BINOMDIST CHIDIST CHIINV
CHITEST CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONFIDENCE CONFIDENCE.NORM CONFIDENCE.T CORREL COUNT COUNTA COUNTBLANK
COVAR COVARIANCE.P COVARIANCE.S CRITBINOM DEVSQ EXPON.DIST EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FDIST FINV FISHER FISHERINV
FORECAST FORECAST.LINEAR FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FREQUENCY FTEST GAMMA GAMMA.DIST
GAMMA.INV GAMMADIST GAMMAINV GAMMALN GAMMALN.PRECISE GAUSS GEOMEAN GROWTH HARMEAN HYPGEOM.DIST HYPGEOMDIST INTERCEPT KURT LARGE LINEST
LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST MAX MAXA MEDIAN MIN MINA MODE MODE.MULT MODE.SNGL NEGBINOM.DIST NEGBINOMDIST NORM.DIST
NORM.INV NORM.S.DIST NORM.S.INV NORMDIST NORMINV NORMSDIST NORMSINV PEARSON PERCENTILE PERCENTILE.EXC PERCENTILE.INC PERCENTRANK
PERCENTRANK.EXC PERCENTRANK.INC PERMUT PERMUTATIONA PHI POISSON POISSON.DIST PROB QUARTILE QUARTILE.EXC QUARTILE.INC RANK RANK.AVG RANK.EQ
RSQ SKEW SKEW.P SLOPE SMALL STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP STDEVPA STEYX T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST
TDIST TINV TREND TRIMMEAN TTEST VAR VAR.P VAR.S VARA VARP VARPA WEIBULL WEIBULL.DIST Z.TEST ZTEST`.split(/\s+/);

const strength = col(1345, 1301, 1368, 1322, 1310, 1370, 1318, 1350, 1303, 1299);
const ex7 = [4, 5, 8, 7, 11, 4, 3];
const sk = col(3, 4, 5, 2, 3, 4, 5, 6, 4, 7);
const ky = col(2, 3, 9, 1, 8, 7, 5), kx = col(6, 5, 11, 7, 5, 4, 4);

// [function, args, expected, decimal digits]
const cases: [string, Value[], number, number][] = [
  ['AVEDEV', [4, 5, 6, 7, 5, 4, 3], 1.020408163, 8],
  ['AVERAGE', [10, 7, 9, 27, 2], 11, 12],
  ['AVERAGEA', [col(10, 7, 9, 2, 'Not available')], 5.6, 12],
  ['BETA.DIST', [2, 8, 10, true, 1, 3], 0.6854706, 7],
  ['BETA.DIST', [2, 8, 10, false, 1, 3], 1.4837646, 7],
  ['BETA.INV', [0.685470581, 8, 10, 1, 3], 2, 7],
  ['BETADIST', [2, 8, 10, 1, 3], 0.685470581, 8],
  ['BETAINV', [0.685470581, 8, 10, 1, 3], 2, 7],
  ['BINOM.DIST', [6, 10, 0.5, false], 0.205078125, 9],
  ['BINOMDIST', [6, 10, 0.5, false], 0.205078125, 9],
  ['BINOM.DIST', [6, 10, 0.5, true], 0.828125, 9],
  ['BINOM.DIST.RANGE', [60, 0.75, 48], 0.083974967, 8],
  ['BINOM.DIST.RANGE', [60, 0.75, 45, 50], 0.523629793, 8],
  ['BINOM.INV', [6, 0.5, 0.75], 4, 12],
  ['CRITBINOM', [6, 0.5, 0.75], 4, 12],
  ['CHISQ.DIST', [0.5, 1, true], 0.52049988, 8],
  ['CHISQ.DIST', [2, 3, false], 0.20755375, 8],
  ['CHISQ.DIST.RT', [18.307, 10], 0.0500006, 7],
  ['CHIDIST', [18.307, 10], 0.050000589, 9],
  ['CHISQ.INV', [0.93, 1], 3.283020287, 8],
  ['CHISQ.INV', [0.6, 2], 1.832581464, 8],
  ['CHISQ.INV.RT', [0.050001, 10], 18.30697346, 7],
  ['CHIINV', [0.050001, 10], 18.30697346, 7],
  ['CONFIDENCE.NORM', [0.05, 2.5, 50], 0.692951912, 8],
  ['CONFIDENCE', [0.05, 2.5, 50], 0.692951912, 8],
  ['CONFIDENCE.T', [0.05, 1, 50], 0.284196855, 8],
  ['CORREL', [col(3, 2, 4, 5, 6), col(9, 7, 12, 15, 17)], 0.997054486, 9],
  ['PEARSON', [col(9, 7, 5, 3, 1), col(10, 6, 1, 5, 3)], 0.699379, 6],
  ['COVAR', [col(3, 2, 4, 5, 6), col(9, 7, 12, 15, 17)], 5.2, 12],
  ['COVARIANCE.P', [col(3, 2, 4, 5, 6), col(9, 7, 12, 15, 17)], 5.2, 12],
  ['COVARIANCE.S', [col(2, 4, 8), col(5, 11, 12)], 9.666666667, 8],
  ['DEVSQ', ex7, 48, 12],
  ['EXPON.DIST', [0.2, 10, true], 0.86466472, 8],
  ['EXPONDIST', [0.2, 10, false], 1.35335283, 8],
  ['F.DIST', [15.2069, 6, 4, true], 0.99, 6],
  ['F.DIST', [15.2069, 6, 4, false], 0.0012238, 7],
  ['F.DIST.RT', [15.2069, 6, 4], 0.01, 6],
  ['FDIST', [15.2069, 6, 4], 0.01, 6],
  ['F.INV', [0.01, 6, 4], 0.10930991, 8],
  ['F.INV.RT', [0.01, 6, 4], 15.20686, 5],
  ['FINV', [0.01, 6, 4], 15.20686, 5],
  ['F.TEST', [col(6, 7, 9, 15, 21), col(20, 28, 31, 38, 40)], 0.64831785, 8],
  ['FTEST', [col(6, 7, 9, 15, 21), col(20, 28, 31, 38, 40)], 0.64831785, 8],
  ['FISHER', [0.75], 0.972955075, 9],
  ['FISHERINV', [0.972955], 0.75, 6],
  ['FORECAST', [30, col(6, 7, 9, 15, 21), col(20, 28, 31, 38, 40)], 10.607253, 6],
  ['FORECAST.LINEAR', [30, col(6, 7, 9, 15, 21), col(20, 28, 31, 38, 40)], 10.607253, 6],
  ['GAMMA', [2.5], 1.329340388, 9],
  ['GAMMA', [-0.75], -4.834146544, 8],
  ['GAMMA', [5], 24, 12],
  ['GAMMA.DIST', [10.00001131, 9, 2, false], 0.032639, 6],
  ['GAMMA.DIST', [10.00001131, 9, 2, true], 0.068094, 6],
  ['GAMMADIST', [10.00001131, 9, 2, true], 0.068094, 6],
  ['GAMMA.INV', [0.068094, 9, 2], 10.0000112, 5],
  ['GAMMAINV', [0.068094, 9, 2], 10.0000112, 5],
  ['GAMMALN', [4], 1.791759469, 9],
  ['GAMMALN.PRECISE', [4], 1.791759469, 9],
  ['GAUSS', [2], 0.477249868, 9],
  ['GEOMEAN', ex7, 5.476986969, 8],
  ['HARMEAN', ex7, 5.028375962, 8],
  ['HYPGEOM.DIST', [1, 4, 8, 20, true], 0.465428277, 8],
  ['HYPGEOM.DIST', [1, 4, 8, 20, false], 0.363261094, 8],
  ['HYPGEOMDIST', [1, 4, 8, 20], 0.363261094, 8],
  ['INTERCEPT', [col(2, 3, 9, 1, 8), col(6, 5, 11, 7, 5)], 0.048387097, 8],
  ['KURT', [sk], -0.151799637, 9],
  ['LARGE', [[[3, 4], [5, 2], [3, 4], [5, 6], [4, 7]], 3], 5, 12],
  ['LARGE', [[[3, 4], [5, 2], [3, 4], [5, 6], [4, 7]], 7], 4, 12],
  ['LOGNORM.DIST', [4, 3.5, 1.2, true], 0.0390836, 7],
  ['LOGNORM.DIST', [4, 3.5, 1.2, false], 0.0176176, 7],
  ['LOGNORMDIST', [4, 3.5, 1.2], 0.0390836, 7],
  ['LOGNORM.INV', [0.039084, 3.5, 1.2], 4.0000252, 6],
  ['LOGINV', [0.039084, 3.5, 1.2], 4.0000252, 6],
  ['MAX', [col(10, 7, 9, 27, 2), 30], 30, 12],
  ['MAXA', [col(0, 0.2, 0.5, 0.4, true)], 1, 12],
  ['MIN', [col(10, 7, 9, 27, 2), 0], 0, 12],
  ['MINA', [col(false, 0.2, 0.5, 0.4, 0.8)], 0, 12],
  ['MEDIAN', [1, 2, 3, 4, 5], 3, 12],
  ['MEDIAN', [1, 2, 3, 4, 5, 6], 3.5, 12],
  ['MODE.SNGL', [5.6, 4, 4, 3, 2, 4], 4, 12],
  ['MODE', [5.6, 4, 4, 3, 2, 4], 4, 12],
  ['NEGBINOM.DIST', [10, 5, 0.25, true], 0.3135141, 7],
  ['NEGBINOM.DIST', [10, 5, 0.25, false], 0.0550487, 7],
  ['NEGBINOMDIST', [10, 5, 0.25], 0.05504866, 8],
  ['NORM.DIST', [42, 40, 1.5, true], 0.9087888, 7],
  ['NORM.DIST', [42, 40, 1.5, false], 0.10934005, 8],
  ['NORMDIST', [42, 40, 1.5, true], 0.9087888, 7],
  ['NORM.INV', [0.908789, 40, 1.5], 42.000002, 6],
  ['NORMINV', [0.908789, 40, 1.5], 42.000002, 6],
  ['NORM.S.DIST', [1.333333, true], 0.908788726, 9],
  ['NORM.S.DIST', [1.333333, false], 0.164010148, 9],
  ['NORMSDIST', [1.333333], 0.908788726, 9],
  ['NORM.S.INV', [0.908789], 1.3333347, 7],
  ['NORMSINV', [0.908789], 1.3333347, 7],
  ['PERCENTILE.EXC', [col(1, 2, 3, 6, 6, 6, 7, 8, 9), 0.25], 2.5, 12],
  ['PERCENTILE.INC', [col(1, 3, 2, 4), 0.3], 1.9, 12],
  ['PERCENTILE', [col(1, 3, 2, 4), 0.3], 1.9, 12],
  ['PERCENTRANK.EXC', [col(1, 2, 3, 6, 6, 6, 7, 8, 9), 7], 0.7, 12],
  ['PERCENTRANK.EXC', [col(1, 2, 3, 6, 6, 6, 7, 8, 9), 5.43], 0.381, 12],
  ['PERCENTRANK.EXC', [col(1, 2, 3, 6, 6, 6, 7, 8, 9), 5.43, 1], 0.3, 12],
  ['PERCENTRANK.INC', [col(13, 12, 11, 8, 4, 3, 2, 1, 1, 1), 2], 0.333, 12],
  ['PERCENTRANK.INC', [col(13, 12, 11, 8, 4, 3, 2, 1, 1, 1), 4], 0.555, 12],
  ['PERCENTRANK.INC', [col(13, 12, 11, 8, 4, 3, 2, 1, 1, 1), 8], 0.666, 12],
  ['PERCENTRANK', [col(13, 12, 11, 8, 4, 3, 2, 1, 1, 1), 5], 0.583, 12],
  ['PERMUT', [100, 3], 970200, 12],
  ['PERMUT', [3, 2], 6, 12],
  ['PERMUTATIONA', [3, 2], 9, 12],
  ['PERMUTATIONA', [2, 2], 4, 12],
  ['PHI', [0.75], 0.301137432, 9],
  ['POISSON.DIST', [2, 5, true], 0.124652019, 9],
  ['POISSON.DIST', [2, 5, false], 0.084224337, 9],
  ['POISSON', [2, 5, false], 0.084224337, 9],
  ['PROB', [col(0, 1, 2, 3), col(0.2, 0.3, 0.1, 0.4), 2], 0.1, 12],
  ['PROB', [col(0, 1, 2, 3), col(0.2, 0.3, 0.1, 0.4), 1, 3], 0.8, 12],
  ['QUARTILE.EXC', [col(6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49), 1], 15, 12],
  ['QUARTILE.EXC', [col(6, 7, 15, 36, 39, 40, 41, 42, 43, 47, 49), 3], 43, 12],
  ['QUARTILE.INC', [col(1, 2, 4, 7, 8, 9, 10, 12), 1], 3.5, 12],
  ['QUARTILE', [col(1, 2, 4, 7, 8, 9, 10, 12), 1], 3.5, 12],
  ['RANK.AVG', [94, col(89, 88, 92, 101, 94, 97, 95)], 4, 12],
  ['RANK.AVG', [3.5, col(7, 3.5, 3.5, 1, 2), 1], 3.5, 12],
  ['RANK.EQ', [7, col(7, 3.5, 3.5, 1, 2), 1], 5, 12],
  ['RANK.EQ', [3.5, col(7, 3.5, 3.5, 1, 2)], 2, 12],
  ['RANK', [7, col(7, 3.5, 3.5, 1, 2), 1], 5, 12],
  ['RSQ', [ky, kx], 0.05795, 5],
  ['SKEW', [sk], 0.359543071, 9],
  ['SKEW.P', [sk], 0.303193339, 9],
  ['SLOPE', [ky, kx], 0.305555556, 9],
  ['STEYX', [ky, kx], 3.305718950, 9],
  ['SMALL', [col(3, 4, 5, 2, 3, 4, 6, 4, 7), 4], 4, 12],
  ['SMALL', [col(1, 4, 8, 3, 7, 12, 54, 8, 23), 2], 3, 12],
  ['STANDARDIZE', [42, 40, 1.5], 1.333333333, 9],
  ['STDEV.S', [strength], 27.46391572, 8],
  ['STDEV', [strength], 27.46391572, 8],
  ['STDEVA', [strength], 27.46391572, 8],
  ['STDEV.P', [strength], 26.05455814, 8],
  ['STDEVP', [strength], 26.05455814, 8],
  ['STDEVPA', [strength], 26.05455814, 8],
  ['VAR.S', [strength], 754.2666667, 7],
  ['VAR', [strength], 754.2666667, 7],
  ['VARA', [strength], 754.2666667, 7],
  ['VAR.P', [strength], 678.84, 9],
  ['VARP', [strength], 678.84, 9],
  ['VARPA', [strength], 678.84, 9],
  ['T.DIST', [60, 1, true], 0.99469533, 8],
  ['T.DIST', [8, 3, false], 0.00073691, 8],
  ['T.DIST.2T', [1.959999998, 60], 0.054644930, 9],
  ['T.DIST.RT', [1.959999998, 60], 0.027322465, 9],
  ['TDIST', [1.959999998, 60, 2], 0.054644930, 9],
  ['TDIST', [1.959999998, 60, 1], 0.027322465, 9],
  ['T.INV', [0.75, 2], 0.816496581, 9],
  ['T.INV.2T', [0.546449, 60], 0.606533, 6],
  ['TINV', [0.546449, 60], 0.606533, 6],
  ['T.TEST', [col(3, 4, 5, 8, 9, 1, 2, 4, 5), col(6, 19, 3, 2, 14, 4, 5, 17, 1), 2, 1], 0.196015785, 9],
  ['TTEST', [col(3, 4, 5, 8, 9, 1, 2, 4, 5), col(6, 19, 3, 2, 14, 4, 5, 17, 1), 2, 1], 0.196015785, 9],
  ['TRIMMEAN', [col(4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3), 0.2], 3.777777778, 9],
  ['WEIBULL.DIST', [105, 20, 100, true], 0.929581, 6],
  ['WEIBULL.DIST', [105, 20, 100, false], 0.035589, 6],
  ['WEIBULL', [105, 20, 100, true], 0.929581, 6],
  ['Z.TEST', [col(3, 6, 7, 8, 6, 5, 4, 2, 1, 9), 4], 0.090574197, 9],
  ['Z.TEST', [col(3, 6, 7, 8, 6, 5, 4, 2, 1, 9), 6], 0.863043, 6],
  ['ZTEST', [col(3, 6, 7, 8, 6, 5, 4, 2, 1, 9), 4], 0.090574197, 9],
];

describe('statistical functions', () => {
  it('defines every function exactly once', () => {
    const have = defs.map((d) => d.name);
    expect(new Set(have).size).toBe(have.length);
    expect(NAMES.filter((n) => !have.includes(n))).toEqual([]);
    expect(have.length).toBe(NAMES.length);
  });

  it.each(cases)('%s(%j) = %d', (name, args, expected, digits) => {
    const r = call(name, ...args);
    expect(typeof r).toBe('number');
    expect(r as number).toBeCloseTo(expected, digits);
  });

  it('CHISQ.TEST / CHITEST', () => {
    const act = [[58, 35], [11, 25], [10, 23]], exp = [[45.35, 47.65], [17.56, 18.44], [16.09, 16.91]];
    expect(call('CHISQ.TEST', act, exp) as number).toBeCloseTo(0.000308192, 9);
    expect(call('CHITEST', act, exp) as number).toBeCloseTo(0.000308192, 9);
  });

  it('error cases', () => {
    expect(code(call('AVERAGE', col('a', null)))).toBe('#DIV/0!');
    expect(code(call('AVERAGE', 'abc'))).toBe('#VALUE!');
    expect(code(call('STDEV.S', 1))).toBe('#DIV/0!');
    expect(code(call('MEDIAN', col('a')))).toBe('#NUM!');
    expect(code(call('MODE.SNGL', 1, 2, 3))).toBe('#N/A');
    expect(code(call('GAMMA', 0))).toBe('#NUM!');
    expect(code(call('GAMMA', -2))).toBe('#NUM!');
    expect(code(call('NORM.DIST', 1, 0, 0, true))).toBe('#NUM!');
    expect(code(call('NORM.S.INV', 1))).toBe('#NUM!');
    expect(code(call('PERCENTILE.EXC', col(1, 2, 3, 6, 6, 6, 7, 8, 9), 0))).toBe('#NUM!');
    expect(code(call('PERCENTILE.EXC', col(1, 2, 3, 6, 6, 6, 7, 8, 9), 0.01))).toBe('#NUM!');
    expect(code(call('PERCENTILE.EXC', col(1, 2, 3, 6, 6, 6, 7, 8, 9), 2))).toBe('#NUM!');
    expect(code(call('PERCENTRANK.INC', col(1, 2, 3), 5))).toBe('#N/A');
    expect(code(call('LARGE', col(1, 2), 3))).toBe('#NUM!');
    expect(code(call('CORREL', col(1, 2), col(1, 2, 3)))).toBe('#N/A');
    expect(code(call('CORREL', col(1, 1), col(1, 2)))).toBe('#DIV/0!');
    expect(code(call('RANK.EQ', 9, col(1, 2)))).toBe('#N/A');
    expect(code(call('BINOM.DIST', 11, 10, 0.5, false))).toBe('#NUM!');
    expect(code(call('T.TEST', col(1, 2, 3), col(2, 3, 4), 3, 1))).toBe('#NUM!');
    expect(code(call('FISHER', 1))).toBe('#NUM!');
    expect(code(call('GEOMEAN', 1, -1))).toBe('#NUM!');
    expect(code(call('CONFIDENCE.T', 0.05, 1, 1))).toBe('#DIV/0!');
    expect(code(call('AVERAGE', col(1, Err.na)))).toBe('#N/A');
    expect(code(call('PROB', col(0, 1), col(0.5, 0.4), 0))).toBe('#NUM!');
  });

  it('argument coercion: direct vs range', () => {
    expect(call('AVERAGE', true, '3', 2)).toBe(2);
    expect(call('MAX', col(1, true, '9'))).toBe(1);
    expect(call('MAX', '9', 1)).toBe(9);
    expect(call('MAX', col('a'))).toBe(0);
    // Sheet1!A1:D4 holds numbers, text, booleans and blanks
    const r = ref(0, 0, 0, 3, 3);
    expect(call('COUNT', r)).toBe(6);
    expect(call('COUNTA', r)).toBe(11);
    expect(call('COUNTBLANK', r)).toBe(5);
    expect(call('AVERAGE', r)).toBe(36 / 6);
    expect(call('AVERAGEA', r)).toBeCloseTo(37 / 11, 12);
    expect(call('AVERAGEA', ref(0, 0, 0, 3, 0), ref(0, 0, 3, 0, 3))).toBe(7);
    expect(call('COUNT', col('Data', 'Sales', 39790, 19, 22.24, true, Err.div0))).toBe(3);
    expect(call('COUNT', col('Data', 'Sales', 39790, 19, 22.24, true, Err.div0), 2)).toBe(4);
    expect(call('COUNT', true, '5', 'x', null)).toBe(3);
    expect(call('COUNTA', 'x', null, 1)).toBe(3);
  });

  it('MODE.MULT and FREQUENCY return vertical arrays', () => {
    expect(call('MODE.MULT', col(1, 2, 3, 4, 3, 2, 1, 2, 3, 5, 6, 1))).toEqual([[1], [2], [3]]);
    expect(call('FREQUENCY', col(79, 85, 78, 85, 50, 81, 95, 88, 97), col(70, 79, 89))).toEqual([[1], [2], [4], [2]]);
    expect(call('FREQUENCY', col(1, 2, 3), col(2, 1))).toEqual([[1], [1], [1]]);
  });

  it('LINEST simple and with statistics', () => {
    const r = call('LINEST', col(1, 9, 5, 7), col(0, 4, 2, 3)) as number[][];
    expect(r[0]![0]).toBeCloseTo(2, 12);
    expect(r[0]![1]).toBeCloseTo(1, 12);
    const nc = call('LINEST', col(1, 9, 5, 7), col(0, 4, 2, 3), false, true) as Scalar[][];
    expect(nc[0]![1]).toBe(0);
    expect(code(nc[1]![1]!)).toBe('#N/A');
    expect(code(nc[2]![1]!)).toBe(null);
  });

  it('LINEST multiple regression (office building example)', () => {
    const X = [[2310, 2, 2, 20], [2333, 2, 2, 12], [2356, 3, 1.5, 33], [2379, 3, 2, 43], [2402, 2, 3, 53], [2425, 4, 2, 23],
      [2448, 2, 1.5, 99], [2471, 2, 2, 34], [2494, 3, 3, 23], [2517, 4, 4, 55], [2540, 2, 3, 22]];
    const Y = col(142000, 144000, 151000, 150000, 139000, 169000, 126000, 142900, 163000, 169000, 149000);
    const r = call('LINEST', Y, X, true, true) as Scalar[][];
    const exp = [
      [-234.2371645, 2553.21066, 12529.76817, 27.64138737, 52317.83051],
      [13.26801148, 530.6691519, 400.0668382, 5.429374042, 12237.3616],
      [0.996747993, 970.5784629], [459.7536742, 6], [1732393319, 5652135.316],
    ];
    exp.forEach((row, i) => row.forEach((v, j) => expect(r[i]![j] as number).toBeCloseTo(v, Math.max(0, 6 - Math.ceil(Math.log10(Math.abs(v) + 1))))));
    expect(code(r[2]![2]!)).toBe('#N/A');
    expect(code(r[4]![4]!)).toBe('#N/A');
  });

  it('LOGEST and GROWTH', () => {
    const months = col(11, 12, 13, 14, 15, 16), units = col(33100, 47300, 69000, 102000, 150000, 220000);
    const r = call('LOGEST', units, months, true, true) as number[][];
    expect(r[0]![0]).toBeCloseTo(1.463275628, 8);
    expect(r[0]![1]).toBeCloseTo(495.3047702, 5);
    expect(r[1]![0]).toBeCloseTo(0.002633403, 8);
    expect(r[1]![1]).toBeCloseTo(0.035834282, 8);
    expect(r[2]![0]).toBeCloseTo(0.99980862, 8);
    expect(r[2]![1]).toBeCloseTo(0.011016315, 8);
    expect(r[3]![0]).toBeCloseTo(20896.8011, 3);
    expect(r[3]![1]).toBe(4);
    const g = call('GROWTH', units, months) as number[][];
    expect(g[0]![0]).toBeCloseTo(32618.20377, 4);
    const g2 = call('GROWTH', units, months, col(17, 18)) as number[][];
    expect(g2[0]![0]).toBeCloseTo(320196.7184, 3);
    expect(g2[1]![0]).toBeCloseTo(468536.0539, 3);
  });

  it('TREND', () => {
    const m = col(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12);
    const cost = col(133890, 135000, 135790, 137300, 138130, 139100, 139900, 141120, 141890, 143230, 144000, 145290);
    expect((call('TREND', cost, m) as number[][])[0]![0]).toBeCloseTo(133953.3333, 3);
    const t = call('TREND', cost, m, col(13, 14, 15)) as number[][];
    expect(t[0]![0]).toBeCloseTo(146171.5152, 3);
    expect(t[1]![0]).toBeCloseTo(147189.6970, 3);
    expect(t[2]![0]).toBeCloseTo(148207.8788, 3);
  });

  it('FORECAST.ETS family on a clean seasonal series', () => {
    const season = [5, -3, 2, -4];
    const t = Array.from({ length: 24 }, (_, i) => i + 1);
    const y = t.map((i) => 100 + 2 * i + season[(i - 1) % 4]!);
    expect(call('FORECAST.ETS.SEASONALITY', col(...y), col(...t))).toBe(4);
    expect(call('FORECAST.ETS', 25, col(...y), col(...t)) as number).toBeCloseTo(100 + 50 + 5, 1);
    expect(call('FORECAST.ETS', 26, col(...y), col(...t), 4) as number).toBeCloseTo(100 + 52 - 3, 1);
    expect(call('FORECAST.ETS.STAT', col(...y), col(...t), 8)).toBe(1);
    expect(call('FORECAST.ETS.CONFINT', 25, col(...y), col(...t)) as number).toBeLessThan(1);
    expect(code(call('FORECAST.ETS', 10, col(...y), col(...t)))).toBe('#NUM!');
    expect(code(call('FORECAST.ETS', 25, col(...y), col(...t.slice(1))))).toBe('#N/A');
  });

  it('element-wise distribution functions accept the scalar flag', () => {
    const d = defs.find((x) => x.name === 'NORM.DIST')!;
    expect(d.scalar).toBe(true);
    expect(defs.find((x) => x.name === 'LARGE')!.scalar).toEqual([1]);
  });

  it('kernels stay accurate in the tails', () => {
    expect(call('NORM.S.DIST', -10, true) as number).toBeCloseTo(7.61985302416e-24, 33);
    expect(call('NORM.S.INV', 1e-10) as number).toBeCloseTo(-6.361340902, 8);
    expect(call('T.INV', 0.001, 1) as number).toBeCloseTo(-318.3088389, 5);
    expect(call('CHISQ.INV.RT', 1e-10, 3) as number).toBeCloseTo(49.54215593, 6);
    expect(call('GAMMALN', 0.5) as number).toBeCloseTo(0.5723649429, 10);
    expect(call('GAMMALN', 100) as number).toBeCloseTo(359.1342053695754, 9);
  });
});
