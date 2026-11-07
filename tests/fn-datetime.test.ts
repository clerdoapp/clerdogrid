import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/datetime';
import { Err, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const D = (y: number, m: number, d: number): number => call('DATE', y, m, d) as number;
const col = (...xs: Value[]): Value => xs.map((x) => [x]) as Value;

describe('DATE and parts', () => {
  it('DATE', () => {
    expect(call('DATE', 2008, 1, 1)).toBe(39448);
    expect(call('DATE', 1900, 1, 1)).toBe(1);
    expect(call('DATE', 1900, 2, 28)).toBe(59);
    expect(call('DATE', 1900, 2, 29)).toBe(60);
    expect(call('DATE', 1900, 3, 1)).toBe(61);
    expect(call('DATE', 108, 1, 2)).toBe(39449);
    expect(call('DATE', 2008, 14, 2)).toBe(D(2009, 2, 2));
    expect(call('DATE', 2008, 1, 35)).toBe(D(2008, 2, 4));
    expect(call('DATE', 2008, -3, 2)).toBe(D(2007, 9, 2));
    expect(call('DATE', 2008, 1, -15)).toBe(D(2007, 12, 16));
    expect(call('DATE', 1900, 1, 0)).toBe(0);
    expect(call('DATE', 10000, 1, 1)).toEqual(Err.num);
    expect(call('DATE', -1, 1, 1)).toEqual(Err.num);
    expect(call('DATE', 1899, 12, 31)).toBe(D(3799, 12, 31));
    expect(call('DATE', '2008', '1', '1')).toBe(39448);
    expect(call('DATE', 'x', 1, 1)).toEqual(Err.value);
  });
  it('YEAR/MONTH/DAY/WEEKDAY', () => {
    expect(call('YEAR', 39448)).toBe(2008);
    expect(call('MONTH', 39448)).toBe(1);
    expect(call('DAY', 60)).toBe(29);
    expect(call('MONTH', 60)).toBe(2);
    expect(call('DAY', 0)).toBe(0);
    expect(call('YEAR', 0)).toBe(1900);
    expect(call('DAY', '2011-04-15')).toBe(15);
    expect(call('YEAR', null)).toBe(1900);
    expect(call('YEAR', -1)).toEqual(Err.num);
    expect(call('YEAR', 'abc')).toEqual(Err.value);
    expect(call('WEEKDAY', 1)).toBe(1);
    expect(call('WEEKDAY', 39492)).toBe(5);
    expect(call('WEEKDAY', 39492, 2)).toBe(4);
    expect(call('WEEKDAY', 39492, 3)).toBe(3);
    expect(call('WEEKDAY', 39492, 11)).toBe(4);
    expect(call('WEEKDAY', 39492, 17)).toBe(5);
    expect(call('WEEKDAY', 39492, 4)).toEqual(Err.num);
    expect(call('WEEKDAY', 61)).toBe(5); // 1900-03-01 was a Thursday
  });
  it('time parts', () => {
    expect(call('HOUR', 0.75)).toBe(18);
    expect(call('HOUR', '7/18/2011 7:45')).toBe(7);
    expect(call('MINUTE', '12:45:00 PM')).toBe(45);
    expect(call('SECOND', '4:48:18 PM')).toBe(18);
    expect(call('SECOND', 0.0000057870)).toBe(0); // 0.5 s rounds... down at the boundary
    expect(call('HOUR', 1.99999999)).toBe(0); // rounds to the next midnight
    expect(call('TIME', 12, 0, 0)).toBe(0.5);
    expect(call('TIME', 16, 48, 10)).toBeCloseTo(0.700115741, 9);
    expect(call('TIME', 27, 0, 0)).toBeCloseTo(0.125, 12);
    expect(call('TIME', 0, -1, 0)).toEqual(Err.num);
    expect(call('TIME', 0, 750, 0)).toBeCloseTo(0.520833333, 9);
    expect(call('TIMEVALUE', '2:24 AM')).toBeCloseTo(0.1, 12);
    expect(call('TIMEVALUE', '22-Aug-2011 6:35 AM')).toBeCloseTo(0.274305556, 9);
    expect(call('TIMEVALUE', 'x')).toEqual(Err.value);
    expect(call('DATEVALUE', '8/22/2011')).toBe(40777);
    expect(call('DATEVALUE', '22-MAY-2011')).toBe(40685);
    expect(call('DATEVALUE', '2011/02/23')).toBe(40597);
    expect(call('DATEVALUE', 40777)).toEqual(Err.value);
    expect(call('DATEVALUE', '2/30/2011')).toEqual(Err.value);
  });
});

describe('date arithmetic', () => {
  it('DATEDIF', () => {
    expect(call('DATEDIF', D(2001, 1, 1), D(2003, 1, 1), 'Y')).toBe(2);
    expect(call('DATEDIF', D(2001, 6, 1), D(2002, 8, 15), 'D')).toBe(440);
    expect(call('DATEDIF', D(2001, 6, 1), D(2002, 8, 15), 'YD')).toBe(75);
    expect(call('DATEDIF', D(2001, 6, 1), D(2002, 8, 15), 'm')).toBe(14);
    expect(call('DATEDIF', D(2001, 6, 1), D(2002, 8, 15), 'YM')).toBe(2);
    expect(call('DATEDIF', D(2001, 6, 1), D(2002, 8, 15), 'MD')).toBe(14);
    expect(call('DATEDIF', D(2011, 1, 31), D(2011, 3, 1), 'MD')).toBe(-2);
    expect(call('DATEDIF', D(2011, 1, 15), D(2011, 3, 10), 'MD')).toBe(23);
    expect(call('DATEDIF', D(2003, 1, 1), D(2001, 1, 1), 'Y')).toEqual(Err.num);
    expect(call('DATEDIF', D(2001, 1, 1), D(2003, 1, 1), 'W')).toEqual(Err.num);
  });
  it('DAYS/DAYS360/EDATE/EOMONTH', () => {
    expect(call('DAYS', '3/15/11', '2/1/11')).toBe(42);
    expect(call('DAYS', D(2011, 12, 31), D(2011, 1, 1))).toBe(364);
    expect(call('DAYS360', D(2011, 1, 30), D(2011, 12, 31))).toBe(330);
    expect(call('DAYS360', D(2011, 1, 30), D(2011, 12, 31), true)).toBe(330);
    expect(call('DAYS360', D(2011, 2, 28), D(2011, 3, 31))).toBe(30);
    expect(call('DAYS360', D(2011, 1, 1), D(2011, 1, 31))).toBe(30);
    expect(call('DAYS360', D(2011, 1, 1), D(2011, 1, 31), true)).toBe(29);
    expect(call('EDATE', D(2011, 1, 15), 1)).toBe(D(2011, 2, 15));
    expect(call('EDATE', D(2011, 1, 15), -1)).toBe(D(2010, 12, 15));
    expect(call('EDATE', D(2011, 1, 31), 1)).toBe(D(2011, 2, 28));
    expect(call('EDATE', D(2011, 1, 15), 2.9)).toBe(D(2011, 3, 15));
    expect(call('EOMONTH', D(2011, 1, 1), 1)).toBe(40602);
    expect(call('EOMONTH', D(2011, 1, 1), -3)).toBe(40482);
    expect(call('EOMONTH', D(1900, 1, 15), 1)).toBe(60);
    expect(call('EOMONTH', D(1900, 1, 15), -2)).toEqual(Err.num);
  });
  it('ISOWEEKNUM/WEEKNUM', () => {
    expect(call('ISOWEEKNUM', D(2012, 3, 9))).toBe(10);
    expect(call('WEEKNUM', D(2012, 3, 9))).toBe(10);
    expect(call('WEEKNUM', D(2012, 3, 9), 2)).toBe(11);
    expect(call('ISOWEEKNUM', D(2021, 1, 1))).toBe(53);
    expect(call('ISOWEEKNUM', D(2019, 12, 30))).toBe(1);
    expect(call('WEEKNUM', D(2021, 1, 1), 21)).toBe(53);
    expect(call('WEEKNUM', D(2021, 1, 3), 1)).toBe(2);
    expect(call('WEEKNUM', D(2021, 1, 3), 2)).toBe(1);
    expect(call('WEEKNUM', D(2021, 1, 3), 17)).toBe(2);
    expect(call('WEEKNUM', D(2021, 1, 3), 3)).toEqual(Err.num);
  });
  it('NETWORKDAYS/WORKDAY', () => {
    const hol = col(D(2012, 11, 22), D(2012, 12, 4), D(2013, 1, 21));
    expect(call('NETWORKDAYS', D(2012, 10, 1), D(2013, 3, 1))).toBe(110);
    expect(call('NETWORKDAYS', D(2012, 10, 1), D(2013, 3, 1), D(2012, 11, 22))).toBe(109);
    expect(call('NETWORKDAYS', D(2012, 10, 1), D(2013, 3, 1), hol)).toBe(107);
    expect(call('NETWORKDAYS', D(2013, 3, 1), D(2012, 10, 1))).toBe(-110);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 1, 31))).toBe(22);
    expect(call('NETWORKDAYS.INTL', D(2006, 2, 28), D(2006, 1, 31))).toBe(-21);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 2, 1), 7, col(D(2006, 1, 2), D(2006, 1, 16)))).toBe(22);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 2, 1), '0010001', col(D(2006, 1, 2), D(2006, 1, 16)))).toBe(20);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 2, 1), 8)).toEqual(Err.num);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 2, 1), '111111')).toEqual(Err.value);
    expect(call('NETWORKDAYS.INTL', D(2006, 1, 1), D(2006, 1, 7), '1111111')).toBe(0);
    expect(call('WORKDAY', D(2008, 10, 1), 151)).toBe(D(2009, 4, 30));
    expect(call('WORKDAY', D(2008, 10, 1), 151, col(D(2008, 11, 26), D(2008, 12, 4), D(2009, 1, 21)))).toBe(D(2009, 5, 5));
    expect(call('WORKDAY', D(2008, 10, 4), 0)).toBe(D(2008, 10, 4));
    expect(call('WORKDAY', D(2008, 10, 6), -1)).toBe(D(2008, 10, 3));
    expect(call('WORKDAY.INTL', D(2012, 1, 1), 30, 0)).toEqual(Err.num);
    expect(call('WORKDAY.INTL', D(2012, 1, 1), 90, 11)).toBe(41013);
    expect(call('WORKDAY.INTL', D(2012, 1, 1), 30, 17)).toBe(D(2012, 2, 5));
    expect(call('WORKDAY.INTL', D(2012, 1, 1), 1, '1111111')).toEqual(Err.value);
  });
  it('YEARFRAC', () => {
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30))).toBeCloseTo(0.58055556, 8);
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30), 1)).toBeCloseTo(0.57650273, 8);
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30), 3)).toBeCloseTo(0.57808219, 8);
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30), 2)).toBeCloseTo(0.58611111, 8);
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30), 4)).toBeCloseTo(0.58055556, 8);
    expect(call('YEARFRAC', D(2011, 1, 1), D(2011, 12, 31), 1)).toBeCloseTo(364 / 365, 12);
    expect(call('YEARFRAC', D(2011, 6, 1), D(2012, 6, 1), 1)).toBeCloseTo(1, 12);
    expect(call('YEARFRAC', D(2012, 6, 1), D(2011, 6, 1), 1)).toBeCloseTo(1, 12);
    expect(call('YEARFRAC', D(2010, 1, 1), D(2013, 1, 1), 1)).toBeCloseTo(1096 / 365.25, 12);
    expect(call('YEARFRAC', D(2011, 3, 1), D(2012, 2, 29), 1)).toBeCloseTo(365 / 366, 12);
    expect(call('YEARFRAC', D(2011, 1, 31), D(2011, 2, 28), 0)).toBeCloseTo(28 / 360, 12);
    expect(call('YEARFRAC', D(2012, 1, 1), D(2012, 7, 30), 5)).toEqual(Err.num);
  });
  it('NOW/TODAY are volatile', () => {
    expect(defs.find((d) => d.name === 'NOW')!.volatile).toBe(true);
    const t = call('TODAY') as number;
    expect(Number.isInteger(t)).toBe(true);
    expect((call('NOW') as number) - t).toBeLessThan(1);
  });
});
