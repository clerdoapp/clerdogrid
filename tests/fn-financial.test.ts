import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/financial';
import { defs as dt } from '../src/engine/functions/datetime';
import { Err, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const D = (y: number, m: number, d: number): number => dt.find((x) => x.name === 'DATE')!.fn([y, m, d], ctx) as number;
const close = (v: Value, expected: number, digits = 2): void => { expect(typeof v).toBe('number'); expect(v as number).toBeCloseTo(expected, digits); };
const row = (...xs: number[]): Value => [xs];

describe('time value of money', () => {
  it('PMT/FV/PV/NPER/RATE', () => {
    close(call('PMT', 0.08 / 12, 10, 10000), -1037.03);
    close(call('PMT', 0.06 / 12, 18 * 12, 0, 50000), -129.08);
    close(call('PMT', 0, 10, 1000), -100);
    close(call('FV', 0.06 / 12, 10, -200, -500, 1), 2581.4);
    close(call('FV', 0.12 / 12, 12, -1000), 12682.5);
    close(call('PV', 0.08 / 12, 12 * 20, 500, 0, 0), -59777.15);
    close(call('NPER', 0.12 / 12, -100, -1000, 10000, 1), 59.6738657, 6);
    close(call('NPER', 0.12 / 12, -100, -1000, 10000), 60.0821229, 6);
    close(call('NPER', 0.12 / 12, -100, -1000), -9.57859404, 6);
    close(call('RATE', 4 * 12, -200, 8000), 0.00770147, 8);
    expect(call('RATE', 10, 100, 100)).toEqual(Err.num);
    expect(call('PMT', 'abc', 1, 1)).toEqual(Err.value);
  });
  it('IPMT/PPMT/ISPMT/CUMIPMT/CUMPRINC', () => {
    close(call('IPMT', 0.1 / 12, 1, 3 * 12, 8000), -66.67);
    close(call('IPMT', 0.1, 3, 3, 8000), -292.45);
    close(call('PPMT', 0.1 / 12, 1, 24, 2000), -75.62);
    close(call('PPMT', 0.08, 10, 10, 200000), -27598.05);
    close(call('ISPMT', 0.1 / 12, 1, 36, 8000), -64.8148, 4);
    close(call('CUMIPMT', 0.09 / 12, 360, 125000, 13, 24, 0), -11135.23);
    close(call('CUMIPMT', 0.09 / 12, 360, 125000, 1, 1, 0), -937.5);
    close(call('CUMPRINC', 0.09 / 12, 360, 125000, 13, 24, 0), -934.1071234, 6);
    close(call('CUMPRINC', 0.09 / 12, 360, 125000, 1, 1, 0), -68.27827118, 6);
    expect(call('CUMIPMT', 0.09 / 12, 360, 125000, 13, 24, 2)).toEqual(Err.num);
    expect(call('IPMT', 0.1, 4, 3, 8000)).toEqual(Err.num);
  });
  it('EFFECT/NOMINAL/PDURATION/RRI/DOLLARDE/DOLLARFR', () => {
    close(call('EFFECT', 0.0525, 4), 0.053542667, 8);
    close(call('NOMINAL', 0.053543, 4), 0.05250032, 8);
    close(call('PDURATION', 0.025, 2000, 2200), 3.86, 2);
    close(call('RRI', 96, 10000, 11000), 0.0009933, 7);
    close(call('DOLLARDE', 1.02, 16), 1.125, 10);
    close(call('DOLLARDE', 1.1, 32), 1.3125, 10);
    close(call('DOLLARFR', 1.125, 16), 1.02, 10);
    close(call('DOLLARFR', 1.125, 32), 1.04, 10);
    expect(call('DOLLARDE', 1.02, 0)).toEqual(Err.div0);
    expect(call('EFFECT', 0.05, 0)).toEqual(Err.num);
  });
});

describe('cash flows', () => {
  it('NPV/IRR/MIRR', () => {
    close(call('NPV', 0.1, -10000, 3000, 4200, 6800), 1188.44);
    close(call('NPV', 0.08, row(8000, 9200, 10000, 12000, 14500)) as number - 40000, 1922.06);
    close(call('IRR', row(-70000, 12000, 15000, 18000, 21000)), -0.021244848, 8);
    close(call('IRR', row(-70000, 12000, 15000, 18000, 21000, 26000)), 0.086630948, 8);
    close(call('IRR', row(-70000, 12000, 15000), -0.1), -0.443506941, 8);
    expect(call('IRR', row(1, 2, 3))).toEqual(Err.num);
    close(call('MIRR', row(-120000, 39000, 30000, 21000, 37000, 46000), 0.1, 0.12), 0.126094, 6);
    close(call('MIRR', row(-120000, 39000, 30000, 21000), 0.1, 0.12), -0.048044655, 6);
  });
  it('XNPV/XIRR/FVSCHEDULE', () => {
    const v = row(-10000, 2750, 4250, 3250, 2750), d = row(39448, 39508, 39751, 39859, 39904);
    close(call('XNPV', 0.09, v, d), 2086.65);
    close(call('XIRR', v, d), 0.373362535, 8);
    expect(call('XIRR', row(1, 2), row(39448, 39508))).toEqual(Err.num);
    close(call('FVSCHEDULE', 1, row(0.09, 0.11, 0.1)), 1.33089, 5);
  });
});

describe('depreciation', () => {
  it('SLN/SYD/DB/DDB/VDB', () => {
    close(call('SLN', 30000, 7500, 10), 2250);
    close(call('SYD', 30000, 7500, 10, 1), 4090.91);
    close(call('SYD', 30000, 7500, 10, 10), 409.09);
    close(call('DB', 1000000, 100000, 6, 1, 7), 186083.33);
    close(call('DB', 1000000, 100000, 6, 2, 7), 259639.42);
    close(call('DB', 1000000, 100000, 6, 7, 7), 15845.1);
    close(call('DDB', 2400, 300, 3650, 1), 1.32);
    close(call('DDB', 2400, 300, 120, 1, 2), 40);
    close(call('DDB', 2400, 300, 10, 1, 2), 480);
    close(call('DDB', 2400, 300, 10, 2, 1.5), 306);
    close(call('DDB', 2400, 300, 10, 10), 22.12);
    close(call('VDB', 2400, 300, 3650, 0, 1), 1.32);
    close(call('VDB', 2400, 300, 120, 0, 1), 40);
    close(call('VDB', 2400, 300, 10, 0, 1), 480);
    close(call('VDB', 2400, 300, 120, 6, 18), 396.31);
    close(call('VDB', 2400, 300, 120, 6, 18, 1.5), 311.81);
    close(call('VDB', 2400, 300, 10, 0, 0.875, 1.5), 315);
  });
  it('AMORLINC/AMORDEGRC', () => {
    close(call('AMORLINC', 2400, D(2008, 8, 19), D(2008, 12, 31), 300, 1, 0.15, 1), 360);
    close(call('AMORDEGRC', 2400, D(2008, 8, 19), D(2008, 12, 31), 300, 1, 0.15, 1), 776);
  });
});

describe('securities', () => {
  it('coupon functions', () => {
    const a = [D(2011, 1, 25), D(2011, 11, 15), 2, 1] as const;
    expect(call('COUPDAYBS', ...a)).toBe(71);
    expect(call('COUPDAYS', ...a)).toBe(181);
    expect(call('COUPDAYSNC', ...a)).toBe(110);
    expect(call('COUPNCD', ...a)).toBe(D(2011, 5, 15));
    expect(call('COUPNUM', D(2007, 1, 25), D(2008, 11, 15), 2, 1)).toBe(4);
    expect(call('COUPPCD', ...a)).toBe(D(2010, 11, 15));
    expect(call('COUPNUM', D(2011, 1, 25), D(2011, 11, 15), 3, 1)).toEqual(Err.num);
  });
  it('ACCRINT/ACCRINTM', () => {
    close(call('ACCRINT', D(2008, 3, 1), D(2008, 8, 31), D(2008, 5, 1), 0.1, 1000, 2, 0), 16.666667, 5);
    close(call('ACCRINT', D(2008, 3, 5), D(2008, 8, 31), D(2008, 5, 1), 0.1, 1000, 2, 0, true), 15.555556, 5);
    close(call('ACCRINTM', D(2008, 4, 1), D(2008, 6, 15), 0.1, 1000, 3), 20.54794521, 6);
  });
  it('PRICE/YIELD/DURATION/MDURATION', () => {
    close(call('PRICE', D(2008, 2, 15), D(2017, 11, 15), 0.0575, 0.065, 100, 2, 0), 94.63436, 5);
    close(call('YIELD', D(2008, 2, 15), D(2016, 11, 15), 0.0575, 95.04287, 100, 2, 0), 0.065, 6);
    close(call('DURATION', D(2018, 7, 1), D(2048, 1, 1), 0.08, 0.09, 2, 1), 10.9191453, 6);
    close(call('MDURATION', D(2008, 1, 1), D(2016, 1, 1), 0.08, 0.09, 2, 1), 5.73567, 5);
    expect(call('PRICE', D(2018, 1, 1), D(2008, 1, 1), 0.05, 0.05, 100, 2)).toEqual(Err.num);
  });
  it('odd-period securities', () => {
    close(call('ODDFPRICE', D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2, 1), 113.597717, 5);
    close(call('ODDFYIELD', D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0575, 84.5, 100, 2, 0), 0.0772, 4);
    close(call('ODDLPRICE', D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375, 0.0405, 100, 2, 0), 99.87829, 5);
    close(call('ODDLYIELD', D(2008, 4, 20), D(2008, 6, 15), D(2007, 12, 24), 0.0375, 99.875, 100, 2, 0), 0.045192, 6);
  });
  it('discount and maturity securities', () => {
    close(call('DISC', D(2007, 1, 25), D(2007, 6, 15), 97.975, 100, 1), 0.052420213, 8);
    close(call('INTRATE', D(2008, 2, 15), D(2008, 5, 15), 1000000, 1014420, 2), 0.05768, 5);
    close(call('RECEIVED', D(2008, 2, 15), D(2008, 5, 15), 1000000, 0.0575, 2), 1014584.65, 2);
    close(call('PRICEDISC', D(2008, 2, 16), D(2008, 3, 1), 0.0525, 100, 2), 99.79583, 5);
    close(call('PRICEMAT', D(2008, 2, 15), D(2008, 4, 13), D(2007, 11, 11), 0.061, 0.061, 0), 99.98449888, 6);
    close(call('YIELDDISC', D(2008, 2, 16), D(2008, 3, 1), 99.795, 100, 2), 0.052823, 6);
    close(call('YIELDMAT', D(2008, 3, 15), D(2008, 11, 3), D(2007, 11, 8), 0.0625, 100.0123, 0), 0.060954, 6);
    close(call('TBILLPRICE', D(2008, 3, 31), D(2008, 6, 1), 0.09), 98.45, 2);
    close(call('TBILLYIELD', D(2008, 3, 31), D(2008, 6, 1), 98.45), 0.091417, 6);
    close(call('TBILLEQ', D(2008, 3, 31), D(2008, 6, 1), 0.0914), 0.094151, 6);
    expect(call('TBILLPRICE', D(2008, 3, 31), D(2009, 6, 1), 0.09)).toEqual(Err.num);
  });
});
