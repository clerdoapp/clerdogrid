import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/engineering';
import { Err, ref, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx([[['3+4i', '5-3i', null]]]);
const call = (name: string, ...args: Value[]): Value => {
  const d = defs.find((x) => x.name === name);
  if (!d) throw new Error(`missing ${name}`);
  return d.fn(args, ctx);
};
const near = (v: Value, expected: number, digits = 9): void => {
  expect(typeof v).toBe('number');
  expect(v as number).toBeCloseTo(expected, digits);
};
const rel = (v: Value, expected: number, tol = 1e-13): void => {
  expect(typeof v).toBe('number');
  expect(Math.abs((v as number) - expected) / Math.abs(expected)).toBeLessThan(tol);
};

describe('Bessel functions', () => {
  it('match reference examples', () => {
    near(call('BESSELI', 1.5, 1), 0.981666428, 8);
    near(call('BESSELJ', 1.9, 2), 0.329925728, 8);
    near(call('BESSELK', 1.5, 1), 0.277387804, 8);
    near(call('BESSELY', 2.5, 1), 0.145918138, 8);
  });

  it('are accurate across the range', () => {
    rel(call('BESSELJ', 1, 0), 0.7651976865579666);
    rel(call('BESSELJ', 10, 1), 0.04347274616886144, 1e-12);
    rel(call('BESSELJ', 50, 0), 0.05581232766925181, 1e-12);
    rel(call('BESSELJ', 30, 5), -0.14324029551207718, 1e-12);
    rel(call('BESSELJ', 10, 24), 3.46326295114685e-8, 1e-8);
    rel(call('BESSELJ', -10, 1), -0.04347274616886144, 1e-12);
    rel(call('BESSELY', 1, 0), 0.08825696421567697);
    rel(call('BESSELY', 10, 1), 0.24901542420695388);
    rel(call('BESSELY', 50, 0), -0.09806499547007708, 1e-12);
    rel(call('BESSELY', 0.1, 0), -1.5342386513503667);
    rel(call('BESSELY', 0.01, 1), -63.67859628158415, 1e-10);
    rel(call('BESSELK', 1, 0), 0.42102443824070834);
    rel(call('BESSELK', 2, 1), 0.13986588181652243);
    rel(call('BESSELK', 0.01, 0), 4.721244730161, 1e-12);
    rel(call('BESSELI', 1, 0), 1.2660658777520082);
  });

  it('validate arguments', () => {
    expect(call('BESSELJ', 1, -1)).toBe(Err.num);
    expect(call('BESSELY', 0, 1)).toBe(Err.num);
    expect(call('BESSELK', -1, 1)).toBe(Err.num);
    expect(call('BESSELJ', 'x', 1)).toBe(Err.value);
    expect(call('BESSELJ', true, 1)).toBe(Err.value);
    expect(call('BESSELJ', 1.5, 1.9)).toBe(call('BESSELJ', 1.5, 1));
    expect(call('BESSELJ', 0, 0)).toBe(1);
  });
});

describe('base conversion', () => {
  it('binary', () => {
    expect(call('BIN2DEC', '1100100')).toBe(100);
    expect(call('BIN2DEC', '1111111111')).toBe(-1);
    expect(call('BIN2DEC', 1100100)).toBe(100);
    expect(call('BIN2HEX', '11111011', 4)).toBe('00FB');
    expect(call('BIN2HEX', '1110')).toBe('E');
    expect(call('BIN2HEX', '1111111111')).toBe('FFFFFFFFFF');
    expect(call('BIN2OCT', '1001', 3)).toBe('011');
    expect(call('BIN2OCT', '1100100')).toBe('144');
    expect(call('BIN2OCT', '1111111111')).toBe('7777777777');
    expect(call('BIN2DEC', '12')).toBe(Err.num);
    expect(call('BIN2DEC', '11111111111')).toBe(Err.num);
    expect(call('BIN2HEX', '1110', 0)).toBe(Err.num);
    expect(call('BIN2HEX', '11111011', 1)).toBe(Err.num);
    expect(call('BIN2DEC', null)).toBe(0);
    expect(call('BIN2DEC', true)).toBe(Err.value);
  });

  it('decimal', () => {
    expect(call('DEC2BIN', 9, 4)).toBe('1001');
    expect(call('DEC2BIN', -100)).toBe('1110011100');
    expect(call('DEC2BIN', 512)).toBe(Err.num);
    expect(call('DEC2BIN', -513)).toBe(Err.num);
    expect(call('DEC2BIN', 511)).toBe('111111111');
    expect(call('DEC2BIN', -512)).toBe('1000000000');
    expect(call('DEC2BIN', 9.9)).toBe('1001');
    expect(call('DEC2BIN', '9')).toBe('1001');
    expect(call('DEC2BIN', -1, 3)).toBe('1111111111');
    expect(call('DEC2HEX', 100, 4)).toBe('0064');
    expect(call('DEC2HEX', -54)).toBe('FFFFFFFFCA');
    expect(call('DEC2HEX', 28)).toBe('1C');
    expect(call('DEC2HEX', 64, 1)).toBe(Err.num);
    expect(call('DEC2HEX', 549755813887)).toBe('7FFFFFFFFF');
    expect(call('DEC2HEX', 549755813888)).toBe(Err.num);
    expect(call('DEC2OCT', 58, 3)).toBe('072');
    expect(call('DEC2OCT', -100)).toBe('7777777634');
    expect(call('DEC2OCT', 'abc')).toBe(Err.value);
  });

  it('hexadecimal and octal', () => {
    expect(call('HEX2BIN', 'F', 8)).toBe('00001111');
    expect(call('HEX2BIN', 'B7')).toBe('10110111');
    expect(call('HEX2BIN', 'FFFFFFFFFF')).toBe('1111111111');
    expect(call('HEX2BIN', '200')).toBe(Err.num);
    expect(call('HEX2DEC', 'A5')).toBe(165);
    expect(call('HEX2DEC', 'FFFFFFFF5B')).toBe(-165);
    expect(call('HEX2DEC', '3DA408B9')).toBe(1034160313);
    expect(call('HEX2DEC', 'a5')).toBe(165);
    expect(call('HEX2DEC', 'G1')).toBe(Err.num);
    expect(call('HEX2OCT', 'F', 3)).toBe('017');
    expect(call('HEX2OCT', '3B4E')).toBe('35516');
    expect(call('HEX2OCT', 'FFFFFFFF00')).toBe('7777777400');
    expect(call('OCT2BIN', 3, 3)).toBe('011');
    expect(call('OCT2BIN', 7777777000)).toBe('1000000000');
    expect(call('OCT2DEC', 54)).toBe(44);
    expect(call('OCT2DEC', 7777777533)).toBe(-165);
    expect(call('OCT2DEC', 8)).toBe(Err.num);
    expect(call('OCT2HEX', 100, 4)).toBe('0040');
    expect(call('OCT2HEX', 7777777533)).toBe('FFFFFFFF5B');
  });
});

describe('bitwise', () => {
  it('match reference examples', () => {
    expect(call('BITAND', 1, 5)).toBe(1);
    expect(call('BITAND', 13, 25)).toBe(9);
    expect(call('BITOR', 23, 10)).toBe(31);
    expect(call('BITXOR', 5, 3)).toBe(6);
    expect(call('BITLSHIFT', 4, 2)).toBe(16);
    expect(call('BITRSHIFT', 13, 2)).toBe(3);
    expect(call('BITLSHIFT', 13, -2)).toBe(3);
    expect(call('BITRSHIFT', 3, -2)).toBe(12);
  });

  it('validate arguments', () => {
    expect(call('BITAND', -1, 5)).toBe(Err.num);
    expect(call('BITAND', 1.5, 5)).toBe(Err.num);
    expect(call('BITOR', 2 ** 48, 1)).toBe(Err.num);
    expect(call('BITAND', 2 ** 48 - 1, 2 ** 47)).toBe(2 ** 47);
    expect(call('BITLSHIFT', 1, 48)).toBe(Err.num);
    expect(call('BITLSHIFT', 1, 54)).toBe(Err.num);
    expect(call('BITXOR', 'a', 1)).toBe(Err.value);
  });
});

describe('complex numbers', () => {
  it('COMPLEX formats like other spreadsheets', () => {
    expect(call('COMPLEX', 3, 4)).toBe('3+4i');
    expect(call('COMPLEX', 3, 4, 'j')).toBe('3+4j');
    expect(call('COMPLEX', 0, 1)).toBe('i');
    expect(call('COMPLEX', 1, 0)).toBe('1');
    expect(call('COMPLEX', 0, -1)).toBe('-i');
    expect(call('COMPLEX', 0, 0)).toBe('0');
    expect(call('COMPLEX', 1.5, -2.5)).toBe('1.5-2.5i');
    expect(call('COMPLEX', 3, 4, 'I')).toBe(Err.value);
    expect(call('COMPLEX', 'a', 4)).toBe(Err.value);
  });

  it('parses parts', () => {
    expect(call('IMREAL', '6-9i')).toBe(6);
    expect(call('IMAGINARY', '3+4i')).toBe(4);
    expect(call('IMAGINARY', '0-j')).toBe(-1);
    expect(call('IMAGINARY', 4)).toBe(0);
    expect(call('IMAGINARY', 'i')).toBe(1);
    expect(call('IMAGINARY', '-i')).toBe(-1);
    expect(call('IMREAL', '1e+2-3e-1i')).toBe(100);
    expect(call('IMAGINARY', '1e+2-3e-1i')).toBe(-0.3);
    expect(call('IMREAL', '.5+.25j')).toBe(0.5);
    expect(call('IMREAL', 'abc')).toBe(Err.num);
    expect(call('IMREAL', '3+4k')).toBe(Err.num);
    expect(call('IMREAL', '3+4ii')).toBe(Err.num);
    expect(call('IMREAL', true)).toBe(Err.value);
    expect(call('IMREAL', Err.na)).toBe(Err.na);
  });

  it('match reference examples', () => {
    expect(call('IMABS', '5+12i')).toBe(13);
    near(call('IMARGUMENT', '3+4i'), 0.92729522, 8);
    expect(call('IMARGUMENT', '0')).toBe(Err.div0);
    expect(call('IMCONJUGATE', '3+4i')).toBe('3-4i');
    expect(call('IMCOS', '1+i')).toBe('0.833730025131149-0.988897705762865i');
    expect(call('IMCOSH', '4+3i')).toBe('-27.0349456030742+3.85115333481178i');
    expect(call('IMCSC', '4+3i')).toBe('-0.0754898329158637+0.0648774713706355i');
    expect(call('IMCSCH', '4+3i')).toBe('-0.036275889628626-0.0051744731840194i');
    expect(call('IMDIV', '-238+240i', '10+24i')).toBe('5+12i');
    expect(call('IMEXP', '1+i')).toBe('1.46869393991589+2.28735528717884i');
    expect(call('IMLN', '3+4i')).toBe('1.6094379124341+0.927295218001612i');
    expect(call('IMLOG10', '3+4i')).toBe('0.698970004336019+0.402719196273373i');
    expect(call('IMLOG2', '3+4i')).toBe('2.32192809488736+1.33780421245098i');
    expect(call('IMPOWER', '2+3i', 3)).toBe('-46+9.00000000000001i');
    expect(call('IMPRODUCT', '3+4i', '5-3i')).toBe('27+11i');
    expect(call('IMPRODUCT', '1+2i', 30)).toBe('30+60i');
    expect(call('IMSECH', '4+3i')).toBe('-0.0362534969158689-0.00516434460775318i');
    expect(call('IMSIN', '4+3i')).toBe('-7.61923172032141-6.548120040911i');
    expect(call('IMSINH', '4+3i')).toBe('-27.0168132580039+3.85373803791938i');
    expect(call('IMSQRT', '1+i')).toBe('1.09868411346781+0.455089860562227i');
    expect(call('IMSUB', '13+4i', '5+3i')).toBe('8+i');
    expect(call('IMSUM', '3+4i', '5-3i')).toBe('8+i');
    expect((call('IMTAN', '4+3i') as string).startsWith('0.0049082580674')).toBe(true);
    expect((call('IMCOT', '4+3i') as string).endsWith('-0.999266927805902i')).toBe(true);
    expect((call('IMSEC', '4+3i') as string).startsWith('-0.06529402785794')).toBe(true);
  });

  it('handles ranges, units and errors', () => {
    expect(call('IMSUM', ref(0, 0, 0, 0, 2))).toBe('8+i');
    expect(call('IMSUM', '1+j', '2j')).toBe('1+3j');
    expect(call('IMSUM', '1+i', '1+j')).toBe(Err.value);
    expect(call('IMDIV', '1+i', 0)).toBe(Err.num);
    expect(call('IMLN', 0)).toBe(Err.num);
    expect(call('IMPOWER', 0, -1)).toBe(Err.num);
    expect(call('IMPOWER', 'i', 2)).toBe('-1+1.22464679914735E-16i');
    expect(call('IMPOWER', -2, 2)).toBe('4');
    expect(call('IMSUM', '1+i', 'x')).toBe(Err.num);
  });
});

describe('CONVERT', () => {
  it('matches reference examples', () => {
    expect(call('CONVERT', 1, 'lbm', 'kg')).toBeCloseTo(0.4535924, 7);
    expect(call('CONVERT', 68, 'F', 'C')).toBe(20);
    expect(call('CONVERT', 2.5, 'ft', 'sec')).toBe(Err.na);
    expect(call('CONVERT', call('CONVERT', 100, 'ft', 'm'), 'ft', 'm')).toBeCloseTo(9.290304, 9);
  });

  it('covers categories and prefixes', () => {
    expect(call('CONVERT', 6, 'tsp', 'tbs')).toBe(2);
    expect(call('CONVERT', 6, 'gal', 'l')).toBe(22.712470704);
    expect(call('CONVERT', 6, 'mi', 'km')).toBe(9.656064);
    expect(call('CONVERT', 6, 'in', 'ft')).toBe(0.5);
    expect(call('CONVERT', 1, 'kibyte', 'bit')).toBe(8192);
    expect(call('CONVERT', 1, 'Mbyte', 'kbyte')).toBe(1000);
    expect(call('CONVERT', 100, 'ft2', 'm2')).toBe(9.290304);
    expect(call('CONVERT', 1, 'km2', 'm2')).toBe(1e6);
    expect(call('CONVERT', 1, 'cm3', 'ml')).toBe(1);
    expect(call('CONVERT', 1, 'day', 'hr')).toBe(24);
    expect(call('CONVERT', 1, 'atm', 'Pa')).toBe(101325);
    expect(call('CONVERT', 0, 'C', 'K')).toBe(273.15);
    expect(call('CONVERT', 100, 'C', 'F')).toBe(212);
    expect(call('CONVERT', 0, 'K', 'Rank')).toBe(0);
    expect(call('CONVERT', 80, 'Reau', 'C')).toBe(100);
    expect(call('CONVERT', 1000, 'mK', 'K')).toBe(1);
    expect(call('CONVERT', 1, 'kWh', 'J')).toBe(3.6e6);
    expect(call('CONVERT', 1, 'HP', 'W')).toBeCloseTo(745.69987158227, 9);
    expect(call('CONVERT', 1, 'T', 'ga')).toBe(10000);
    expect(call('CONVERT', 1, 'm/s', 'km/h')).toBeCloseTo(3.6, 12);
    expect(call('CONVERT', 1, 'kn', 'm/h')).toBe(1852);
    expect(call('CONVERT', 1, 'LBM', 'kg')).toBe(Err.na);
    expect(call('CONVERT', 1, 'kft', 'm')).toBe(Err.na);
    expect(call('CONVERT', 1, 'kibit', 'bit')).toBe(1024);
    expect(call('CONVERT', 1, 'kim', 'm')).toBe(Err.na);
    expect(call('CONVERT', 'x', 'm', 'ft')).toBe(Err.value);
  });
});

describe('error function and steps', () => {
  it('ERF and ERFC', () => {
    near(call('ERF', 0.745), 0.70792892, 8);
    near(call('ERF', 1), 0.84270079, 8);
    rel(call('ERF', 0.5), 0.5204998778130465, 1e-15);
    rel(call('ERF', 2.5), 0.999593047982555, 1e-15);
    rel(call('ERF', -0.5), -0.5204998778130465, 1e-15);
    rel(call('ERF', 0, 1), 0.8427007929497149, 1e-15);
    rel(call('ERF.PRECISE', 0.745), 0.7079289200957377, 1e-14);
    rel(call('ERFC', 1), 0.15729920705028513, 1e-14);
    rel(call('ERFC', 1.9), 0.0072095707647425325, 1e-13);
    rel(call('ERFC', 3), 2.209049699858544e-5, 1e-14);
    rel(call('ERFC', 5), 1.5374597944280349e-12, 1e-13);
    rel(call('ERFC.PRECISE', 1), 0.15729920705028513, 1e-14);
    rel(call('ERFC', -1), 1.8427007929497148, 1e-15);
    expect(call('ERF', 'a')).toBe(Err.value);
  });

  it('DELTA and GESTEP', () => {
    expect(call('DELTA', 5, 4)).toBe(0);
    expect(call('DELTA', 5, 5)).toBe(1);
    expect(call('DELTA', 0.5, 0)).toBe(0);
    expect(call('DELTA', 0)).toBe(1);
    expect(call('DELTA', 'a', 1)).toBe(Err.value);
    expect(call('GESTEP', 5, 4)).toBe(1);
    expect(call('GESTEP', 5, 5)).toBe(1);
    expect(call('GESTEP', -4, -5)).toBe(1);
    expect(call('GESTEP', -1)).toBe(0);
  });
});
