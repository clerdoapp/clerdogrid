import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/math';
import { Err, ref, type Matrix, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx([[[1, 2, 'x'], [3, true, null], [4, 5, 6]]]);
const call = (name: string, ...args: Value[]): Value => {
  const d = defs.find((x) => x.name === name);
  if (!d) throw new Error(`missing ${name}`);
  return d.fn(args, ctx);
};
const near = (v: Value, expected: number, digits = 9): void => {
  expect(typeof v).toBe('number');
  expect(v as number).toBeCloseTo(expected, digits);
};

describe('rounding', () => {
  it('ROUND family follows spreadsheet decimal rounding', () => {
    expect(call('ROUND', 2.15, 1)).toBe(2.2);
    expect(call('ROUND', 2.149, 1)).toBe(2.1);
    expect(call('ROUND', -1.475, 2)).toBe(-1.48);
    expect(call('ROUND', 21.5, -1)).toBe(20);
    expect(call('ROUND', 626.3, -3)).toBe(1000);
    expect(call('ROUND', 1.98, -1)).toBe(0);
    expect(call('ROUND', -50.55, -2)).toBe(-100);
    expect(call('ROUND', 2.675, 2)).toBe(2.68);
    expect(call('ROUND', -2.5, 0)).toBe(-3);
    expect(call('ROUND', 1.005, 2)).toBe(1.01);
    expect(call('ROUND', 'abc', 1)).toBe(Err.value);
    expect(call('ROUND', '2.5', 0)).toBe(3);
    expect(call('ROUNDUP', 3.2, 0)).toBe(4);
    expect(call('ROUNDUP', 76.9, 0)).toBe(77);
    expect(call('ROUNDUP', 3.14159, 3)).toBe(3.142);
    expect(call('ROUNDUP', -3.14159, 1)).toBe(-3.2);
    expect(call('ROUNDUP', 31415.92654, -2)).toBe(31500);
    expect(call('ROUNDDOWN', 3.2, 0)).toBe(3);
    expect(call('ROUNDDOWN', 3.14159, 3)).toBe(3.141);
    expect(call('ROUNDDOWN', -3.14159, 1)).toBe(-3.1);
    expect(call('ROUNDDOWN', 31415.92654, -2)).toBe(31400);
    expect(call('TRUNC', 8.9)).toBe(8);
    expect(call('TRUNC', -8.9)).toBe(-8);
    expect(call('TRUNC', 0.45)).toBe(0);
    expect(call('TRUNC', 3.14159, 2)).toBe(3.14);
    expect(call('INT', 8.9)).toBe(8);
    expect(call('INT', -8.9)).toBe(-9);
  });

  it('MROUND, EVEN, ODD', () => {
    expect(call('MROUND', 10, 3)).toBe(9);
    expect(call('MROUND', -10, -3)).toBe(-9);
    expect(call('MROUND', 1.3, 0.2)).toBe(1.4);
    expect(call('MROUND', 5, -2)).toBe(Err.num);
    expect(call('EVEN', 1.5)).toBe(2);
    expect(call('EVEN', 3)).toBe(4);
    expect(call('EVEN', 2)).toBe(2);
    expect(call('EVEN', -1)).toBe(-2);
    expect(call('EVEN', 0)).toBe(0);
    expect(call('ODD', 1.5)).toBe(3);
    expect(call('ODD', 3)).toBe(3);
    expect(call('ODD', 2)).toBe(3);
    expect(call('ODD', -1)).toBe(-1);
    expect(call('ODD', -2)).toBe(-3);
    expect(call('ODD', 0)).toBe(1);
  });

  it('CEILING and FLOOR families', () => {
    expect(call('CEILING', 2.5, 1)).toBe(3);
    expect(call('CEILING', -2.5, -2)).toBe(-4);
    expect(call('CEILING', -2.5, 2)).toBe(-2);
    expect(call('CEILING', 1.5, 0.1)).toBe(1.5);
    expect(call('CEILING', 0.234, 0.01)).toBe(0.24);
    expect(call('CEILING', 2.5, -2)).toBe(Err.num);
    expect(call('FLOOR', 3.7, 2)).toBe(2);
    expect(call('FLOOR', -2.5, -2)).toBe(-2);
    expect(call('FLOOR', 2.5, -2)).toBe(Err.num);
    expect(call('FLOOR', 1.58, 0.1)).toBe(1.5);
    expect(call('FLOOR', 0.234, 0.01)).toBe(0.23);
    expect(call('FLOOR', -2.5, 2)).toBe(-4);
    expect(call('FLOOR', 5, 0)).toBe(Err.div0);
    expect(call('CEILING.MATH', 24.3, 5)).toBe(25);
    expect(call('CEILING.MATH', 6.7)).toBe(7);
    expect(call('CEILING.MATH', -8.1, 2)).toBe(-8);
    expect(call('CEILING.MATH', -5.5, 2, -1)).toBe(-6);
    expect(call('FLOOR.MATH', 24.3, 5)).toBe(20);
    expect(call('FLOOR.MATH', 6.7)).toBe(6);
    expect(call('FLOOR.MATH', -8.1, 2)).toBe(-10);
    expect(call('FLOOR.MATH', -5.5, 2, -1)).toBe(-4);
    expect(call('CEILING.PRECISE', 4.3)).toBe(5);
    expect(call('CEILING.PRECISE', -4.3)).toBe(-4);
    expect(call('CEILING.PRECISE', 4.3, 2)).toBe(6);
    expect(call('CEILING.PRECISE', 4.3, -2)).toBe(6);
    expect(call('CEILING.PRECISE', -4.3, 2)).toBe(-4);
    expect(call('CEILING.PRECISE', -4.3, -2)).toBe(-4);
    expect(call('ISO.CEILING', 4.3)).toBe(5);
    expect(call('ISO.CEILING', -4.3, -2)).toBe(-4);
    expect(call('FLOOR.PRECISE', -3.2, -1)).toBe(-4);
    expect(call('FLOOR.PRECISE', 3.2, 1)).toBe(3);
    expect(call('FLOOR.PRECISE', -3.2, 1)).toBe(-4);
    expect(call('FLOOR.PRECISE', 3.2, -1)).toBe(3);
    expect(call('FLOOR.PRECISE', 3.2)).toBe(3);
  });
});

describe('arithmetic', () => {
  it('MOD sign follows the divisor', () => {
    expect(call('MOD', 3, 2)).toBe(1);
    expect(call('MOD', -3, 2)).toBe(1);
    expect(call('MOD', 3, -2)).toBe(-1);
    expect(call('MOD', -3, -2)).toBe(-1);
    expect(call('MOD', 3, 0)).toBe(Err.div0);
    expect(call('MOD', 0.3, 0.1)).toBe(0);
    expect(call('MOD', 5.5, 2)).toBe(1.5);
  });

  it('QUOTIENT, POWER, SQRT, SIGN, ABS', () => {
    expect(call('QUOTIENT', 5, 2)).toBe(2);
    expect(call('QUOTIENT', 4.5, 3.1)).toBe(1);
    expect(call('QUOTIENT', -10, 3)).toBe(-3);
    expect(call('QUOTIENT', 1, 0)).toBe(Err.div0);
    expect(call('POWER', 5, 2)).toBe(25);
    near(call('POWER', 98.6, 3.2), 2401077.2220695773, 4);
    near(call('POWER', 4, 5 / 4), 5.656854249);
    expect(call('POWER', 0, 0)).toBe(Err.num);
    expect(call('POWER', 0, -1)).toBe(Err.div0);
    expect(call('POWER', -8, 1 / 3)).toBe(Err.num);
    expect(call('SQRT', 16)).toBe(4);
    expect(call('SQRT', -16)).toBe(Err.num);
    expect(call('SQRT', call('ABS', -16))).toBe(4);
    near(call('SQRTPI', 1), 1.772453851);
    near(call('SQRTPI', 2), 2.506628275);
    expect(call('SIGN', 10)).toBe(1);
    expect(call('SIGN', 0)).toBe(0);
    expect(call('SIGN', -0.00001)).toBe(-1);
    expect(call('ABS', -2)).toBe(2);
    expect(call('ABS', null)).toBe(0);
    expect(call('ABS', true)).toBe(1);
    expect(call('ABS', Err.na)).toBe(Err.na);
  });

  it('logarithms and exponentials', () => {
    near(call('EXP', 1), 2.718281828);
    near(call('EXP', 2), 7.389056099);
    near(call('LN', 86), 4.454347296);
    near(call('LN', 2.7182818), 1, 6);
    expect(call('LN', 0)).toBe(Err.num);
    expect(call('LOG', 10)).toBe(1);
    expect(call('LOG', 8, 2)).toBe(3);
    near(call('LOG', 86, 2.7182818), 4.454347, 5);
    expect(call('LOG', 10, 1)).toBe(Err.div0);
    expect(call('LOG', -1)).toBe(Err.num);
    expect(call('LOG10', 86)).toBeCloseTo(1.934498451, 9);
    expect(call('LOG10', 10)).toBe(1);
    expect(call('LOG10', 1e5)).toBe(5);
    expect(call('LOG10', 0)).toBe(Err.num);
  });

  it('FACT, FACTDOUBLE, COMBIN, COMBINA, MULTINOMIAL', () => {
    expect(call('FACT', 5)).toBe(120);
    expect(call('FACT', 1.9)).toBe(1);
    expect(call('FACT', 0)).toBe(1);
    expect(call('FACT', -1)).toBe(Err.num);
    expect(call('FACT', 171)).toBe(Err.num);
    expect(call('FACTDOUBLE', 6)).toBe(48);
    expect(call('FACTDOUBLE', 7)).toBe(105);
    expect(call('FACTDOUBLE', -1)).toBe(1);
    expect(call('FACTDOUBLE', -2)).toBe(Err.num);
    expect(call('COMBIN', 8, 2)).toBe(28);
    expect(call('COMBIN', 2, 8)).toBe(Err.num);
    expect(call('COMBIN', 10.9, 3.2)).toBe(120);
    expect(call('COMBINA', 4, 3)).toBe(20);
    expect(call('COMBINA', 10, 3)).toBe(220);
    expect(call('COMBINA', 0, 0)).toBe(1);
    expect(call('MULTINOMIAL', 2, 3, 4)).toBe(1260);
    expect(call('MULTINOMIAL', 1, -1)).toBe(Err.num);
  });

  it('GCD and LCM', () => {
    expect(call('GCD', 5, 2)).toBe(1);
    expect(call('GCD', 24, 36)).toBe(12);
    expect(call('GCD', 7, 1)).toBe(1);
    expect(call('GCD', 5, 0)).toBe(5);
    expect(call('GCD', -1, 5)).toBe(Err.num);
    expect(call('GCD', 'a', 5)).toBe(Err.value);
    expect(call('GCD', [[12, 18, 24]])).toBe(6);
    expect(call('LCM', 5, 2)).toBe(10);
    expect(call('LCM', 24, 36)).toBe(72);
    expect(call('LCM', 0, 5)).toBe(0);
    expect(call('LCM', 3.9, 4)).toBe(12);
  });
});

describe('trigonometry', () => {
  it('basic functions match reference examples', () => {
    near(call('ACOS', -0.5), 2.094395102, 7);
    near(call('ACOSH', 1), 0, 7);
    near(call('ACOSH', 10), 2.9932228, 6);
    expect(call('ACOSH', 0.5)).toBe(Err.num);
    near(call('ACOT', 2), 0.463647609, 7);
    near(call('ACOTH', 6), 0.168236118, 7);
    expect(call('ACOTH', 0.5)).toBe(Err.num);
    near(call('ASIN', -0.5), -0.523598776, 7);
    expect(call('ASIN', 2)).toBe(Err.num);
    near(call('ASINH', -2.5), -1.647231146, 7);
    near(call('ASINH', 10), 2.99822295, 7);
    near(call('ATAN', 1), 0.785398163, 7);
    near(call('ATAN2', 1, 1), 0.785398163, 7);
    near(call('ATAN2', -1, -1), -2.35619449, 7);
    expect(call('ATAN2', 0, 0)).toBe(Err.div0);
    near(call('ATANH', 0.76159416), 1, 7);
    near(call('ATANH', -0.1), -0.100335348, 7);
    expect(call('ATANH', 1)).toBe(Err.num);
    near(call('COS', 1.047), 0.5001711, 7);
    near(call('COS', (60 * Math.PI) / 180), 0.5, 7);
    near(call('COSH', 4), 27.30823284, 7);
    near(call('COT', 30), -0.156119952, 7);
    expect(call('COT', 0)).toBe(Err.div0);
    near(call('COTH', 2), 1.037314721, 7);
    near(call('CSC', 15), 1.537780562, 7);
    near(call('CSCH', 1.5), 0.469642441, 7);
    near(call('SEC', 45), 1.903594407, 7);
    expect((call('SECH', 45) as number) / 5.725037161098787e-20).toBeCloseTo(1, 12);
    near(call('SIN', Math.PI / 2), 1, 7);
    near(call('SINH', 1), 1.175201194, 7);
    near(call('TAN', 0.785), 0.99920399, 7);
    near(call('TANH', -2), -0.96402758, 7);
    expect(call('SIN', 2 ** 27)).toBe(Err.num);
    expect(call('DEGREES', Math.PI)).toBe(180);
    near(call('RADIANS', 270), 4.712389, 7);
    expect(call('PI')).toBe(Math.PI);
  });
});

describe('aggregates', () => {
  it('SUM coerces typed arguments but skips text and logicals in ranges', () => {
    expect(call('SUM', 3, 2)).toBe(5);
    expect(call('SUM', '5', 15, true)).toBe(21);
    expect(call('SUM', ref(0, 0, 0, 2, 2))).toBe(21);
    expect(call('SUM', 'abc')).toBe(Err.value);
    expect(call('SUM', [[1, Err.na]])).toBe(Err.na);
    expect(call('SUM', 0.1, 0.2)).toBeCloseTo(0.3, 15);
    expect(call('SUM', null)).toBe(0);
  });

  it('PRODUCT and SUMSQ', () => {
    expect(call('PRODUCT', [[5, 15, 30]])).toBe(2250);
    expect(call('PRODUCT', [[5, 15, 30]], 2)).toBe(4500);
    expect(call('PRODUCT', ref(0, 0, 2, 0, 2))).toBe(0);
    expect(call('SUMSQ', 3, 4)).toBe(25);
    expect(call('SUMSQ', ref(0, 0, 0, 1, 1))).toBe(14);
  });

  it('SUMX2MY2, SUMX2PY2, SUMXMY2', () => {
    const a: Matrix = [[2, 3, 9, 1, 8, 7, 5]];
    const b: Matrix = [[6, 5, 11, 7, 5, 4, 4]];
    expect(call('SUMX2MY2', a, b)).toBe(-55);
    expect(call('SUMX2PY2', a, b)).toBe(521);
    expect(call('SUMXMY2', a, b)).toBe(79);
    expect(call('SUMXMY2', [[1, 2]], [[1, 2, 3]])).toBe(Err.na);
    expect(call('SUMXMY2', [[1, 'a']], [[3, 2]])).toBe(4);
  });

  it('SERIESSUM', () => {
    const coeff: Matrix = [[1, -1 / 2, 1 / 24, -1 / 720]];
    near(call('SERIESSUM', Math.PI / 4, 0, 2, coeff), 0.707103, 6);
    near(call('SERIESSUM', 2, 1, 1, [[1, 2, 3]]), 2 + 8 + 24);
  });
});

describe('matrices', () => {
  it('MDETERM', () => {
    expect(call('MDETERM', [[1, 3, 8, 5], [1, 3, 6, 1], [1, 1, 1, 0], [7, 3, 10, 2]])).toBe(88);
    expect(call('MDETERM', [[3, 6], [1, 1]])).toBe(-3);
    expect(call('MDETERM', [[1, 3, 8, 5], [1, 3, 6, 1]])).toBe(Err.value);
    expect(call('MDETERM', [[1, 'a'], [1, 2]])).toBe(Err.value);
  });

  it('MINVERSE', () => {
    expect(call('MINVERSE', [[4, -1], [2, 0]])).toEqual([[0, 0.5], [-1, 2]]);
    const inv = call('MINVERSE', [[1, 2, 1], [3, 4, -1], [0, 2, 0]]) as number[][];
    const expected = [[0.25, 0.25, -0.75], [0, 0, 0.5], [0.75, -0.25, -0.25]];
    inv.forEach((row, i) => row.forEach((x, j) => expect(x).toBeCloseTo(expected[i][j], 12)));
    expect(call('MINVERSE', [[1, 2], [2, 4]])).toBe(Err.num);
  });

  it('MMULT and MUNIT', () => {
    expect(call('MMULT', [[1, 3], [7, 2]], [[2, 0], [0, 2]])).toEqual([[2, 6], [14, 4]]);
    expect(call('MMULT', [[1, 2, 3]], [[4], [5], [6]])).toEqual([[32]]);
    expect(call('MMULT', [[1, 2]], [[1, 2]])).toBe(Err.value);
    expect(call('MUNIT', 3)).toEqual([[1, 0, 0], [0, 1, 0], [0, 0, 1]]);
    expect(call('MUNIT', 0)).toBe(Err.value);
  });
});

describe('numeral systems', () => {
  it('ROMAN in all five forms', () => {
    expect(call('ROMAN', 499, 0)).toBe('CDXCIX');
    expect(call('ROMAN', 499, 1)).toBe('LDVLIV');
    expect(call('ROMAN', 499, 2)).toBe('XDIX');
    expect(call('ROMAN', 499, 3)).toBe('VDIV');
    expect(call('ROMAN', 499, 4)).toBe('ID');
    expect(call('ROMAN', 2013, 0)).toBe('MMXIII');
    expect(call('ROMAN', 499, true)).toBe('CDXCIX');
    expect(call('ROMAN', 499, false)).toBe('ID');
    expect(call('ROMAN', 1999)).toBe('MCMXCIX');
    expect(call('ROMAN', 3999)).toBe('MMMCMXCIX');
    expect(call('ROMAN', 0)).toBe('');
    expect(call('ROMAN', 4000)).toBe(Err.value);
    expect(call('ROMAN', -1)).toBe(Err.value);
    expect(call('ROMAN', 10, 5)).toBe(Err.value);
  });

  it('ARABIC', () => {
    expect(call('ARABIC', 'LVII')).toBe(57);
    expect(call('ARABIC', 'mcmxii')).toBe(1912);
    expect(call('ARABIC', '')).toBe(0);
    expect(call('ARABIC', '-MMXI')).toBe(-2011);
    expect(call('ARABIC', 'ID')).toBe(499);
    expect(call('ARABIC', 'ABC')).toBe(Err.value);
  });

  it('BASE and DECIMAL', () => {
    expect(call('BASE', 7, 2)).toBe('111');
    expect(call('BASE', 100, 16)).toBe('64');
    expect(call('BASE', 15, 2, 10)).toBe('0000001111');
    expect(call('BASE', -1, 2)).toBe(Err.num);
    expect(call('BASE', 10, 37)).toBe(Err.num);
    expect(call('DECIMAL', 'FF', 16)).toBe(255);
    expect(call('DECIMAL', 111, 2)).toBe(7);
    expect(call('DECIMAL', 'zap', 36)).toBe(45745);
    expect(call('DECIMAL', '2', 2)).toBe(Err.num);
  });
});

describe('random', () => {
  it('RAND and RANDBETWEEN stay in range', () => {
    for (let i = 0; i < 50; i++) {
      const r = call('RAND') as number;
      expect(r).toBeGreaterThanOrEqual(0);
      expect(r).toBeLessThan(1);
      const k = call('RANDBETWEEN', -3, 3) as number;
      expect(Number.isInteger(k) && k >= -3 && k <= 3).toBe(true);
    }
    expect(call('RANDBETWEEN', 5, 1)).toBe(Err.num);
    expect(defs.find((d) => d.name === 'RAND')!.volatile).toBe(true);
  });
});

it('every element-wise function is marked scalar', () => {
  const aggregates = new Set(['SUM', 'SUMSQ', 'PRODUCT', 'GCD', 'LCM', 'MULTINOMIAL', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'SERIESSUM', 'MDETERM', 'MINVERSE', 'MMULT', 'MUNIT', 'PI', 'RAND']);
  for (const d of defs) if (!aggregates.has(d.name)) expect(d.scalar, d.name).toBe(true);
  const unique = new Set(defs.map((d) => d.name));
  expect(unique.size).toBe(defs.length);
});
