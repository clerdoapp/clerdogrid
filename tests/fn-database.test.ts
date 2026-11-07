import { describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/database';
import { Err, type Matrix, type Scalar, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => {
  const d = defs.find((x) => x.name === name);
  if (!d) throw new Error(`missing ${name}`);
  return d.fn(args, ctx);
};

// The reference example database for the D-functions.
const db: Matrix = [
  ['Tree', 'Height', 'Age', 'Yield', 'Profit'],
  ['Apple', 18, 20, 14, 105],
  ['Pear', 12, 12, 10, 96],
  ['Cherry', 13, 14, 9, 105],
  ['Apple', 14, 15, 10, 75],
  ['Pear', 9, 8, 8, 76.8],
  ['Apple', 8, 9, 6, 45],
];
const crit = (...rows: Scalar[][]): Matrix => rows;
// Apple taller than 10 and shorter than 16, or any pear.
const appleOrPear = crit(['Tree', 'Height', 'Age', 'Yield', 'Profit', 'Height'], ['=Apple', '>10', null, null, null, '<16'], ['=Pear', null, null, null, null, null]);

describe('reference examples', () => {
  it('DCOUNT, DCOUNTA', () => {
    expect(call('DCOUNT', db, 'Age', crit(['Tree', 'Height', 'Height'], ['=Apple', '>10', '<16']))).toBe(1);
    expect(call('DCOUNTA', db, 'Profit', crit(['Tree', 'Height', 'Height'], ['=Apple', '>10', '<16']))).toBe(1);
    expect(call('DCOUNT', db, null, appleOrPear)).toBe(3);
  });

  it('DMAX, DMIN, DSUM, DPRODUCT, DAVERAGE', () => {
    expect(call('DMAX', db, 'Profit', appleOrPear)).toBe(96);
    expect(call('DMIN', db, 'Profit', appleOrPear)).toBe(75);
    expect(call('DSUM', db, 'Profit', crit(['Tree'], ['=Apple']))).toBe(225);
    expect(call('DSUM', db, 'Profit', appleOrPear)).toBe(247.8);
    expect(call('DPRODUCT', db, 'Yield', appleOrPear)).toBe(800);
    expect(call('DAVERAGE', db, 'Yield', crit(['Height'], ['>10']))).toBe(10.75);
    expect(call('DAVERAGE', db, 3, db)).toBe(13);
  });

  it('DSTDEV, DSTDEVP, DVAR, DVARP', () => {
    const trees = crit(['Tree'], ['=Apple'], ['=Pear']);
    expect(call('DSTDEV', db, 'Yield', trees)).toBeCloseTo(2.966479395, 9);
    expect(call('DSTDEVP', db, 'Yield', trees)).toBeCloseTo(2.653299832, 9);
    expect(call('DVAR', db, 'Yield', trees)).toBeCloseTo(8.8, 12);
    expect(call('DVARP', db, 'Yield', trees)).toBeCloseTo(7.04, 12);
    expect(call('DSTDEV', db, 'Yield', appleOrPear)).toBeCloseTo(1.154700538, 9);
  });

  it('DGET', () => {
    expect(call('DGET', db, 'Yield', crit(['Tree', 'Height', 'Height'], ['=Apple', '>10', '<16']))).toBe(10);
    expect(call('DGET', db, 'Yield', crit(['Tree'], ['=Orange']))).toBe(Err.value);
    expect(call('DGET', db, 'Yield', appleOrPear)).toBe(Err.num);
  });
});

describe('criteria rules', () => {
  it('plain text is a case-insensitive prefix match; =text is exact', () => {
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['ap']))).toBe(3);
    expect(call('DCOUNTA', db, 'Tree', crit(['tree'], ['APPLE']))).toBe(3);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['=App']))).toBe(0);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['<>Apple']))).toBe(3);
  });

  it('wildcards and escapes', () => {
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['=*e*']))).toBe(6);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['=?ear']))).toBe(2);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['*rr']))).toBe(1);
    const t: Matrix = [['Name'], ['a*b'], ['axb'], ['a?']];
    expect(call('DCOUNTA', t, 'Name', crit(['Name'], ['=a~*b']))).toBe(1);
    expect(call('DCOUNTA', t, 'Name', crit(['Name'], ['=a*b']))).toBe(2);
    expect(call('DCOUNTA', t, 'Name', crit(['Name'], ['=a~?']))).toBe(1);
  });

  it('numeric comparisons, numbers as text, and text order', () => {
    expect(call('DSUM', db, 'Height', crit(['Height'], ['>=13']))).toBe(45);
    expect(call('DSUM', db, 'Height', crit(['Height'], [12]))).toBe(12);
    expect(call('DSUM', db, 'Height', crit(['Height'], ['12']))).toBe(12);
    expect(call('DSUM', db, 'Height', crit(['Height'], ['<>12']))).toBe(62);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['>c']))).toBe(3);
    expect(call('DCOUNTA', db, 'Tree', crit(['Tree'], ['<B']))).toBe(3);
  });

  it('blank and non-blank tests, and empty criteria cells', () => {
    const t: Matrix = [['K', 'V'], ['a', 1], [null, 2], ['b', 3]];
    expect(call('DSUM', t, 'V', crit(['K'], ['=']))).toBe(2);
    expect(call('DSUM', t, 'V', crit(['K'], ['<>']))).toBe(4);
    expect(call('DSUM', t, 'V', crit(['K'], [null]))).toBe(6);
    expect(call('DSUM', t, 'V', crit(['K', 'V'], ['a', null], [null, '>2']))).toBe(4);
  });

  it('unknown criteria headers never match', () => {
    expect(call('DSUM', db, 'Profit', crit(['Colour'], ['red']))).toBe(0);
    expect(call('DSUM', db, 'Profit', crit(['Colour'], [null]))).toBe(call('DSUM', db, 'Profit', crit(['Tree'], [null])));
  });
});

describe('argument handling', () => {
  it('field by name or index; bad fields are #VALUE!', () => {
    expect(call('DSUM', db, 5, crit(['Tree'], ['Apple']))).toBe(225);
    expect(call('DSUM', db, ' profit ', crit(['Tree'], ['Apple']))).toBe(225);
    expect(call('DSUM', db, 6, crit(['Tree'], ['Apple']))).toBe(Err.value);
    expect(call('DSUM', db, 0, crit(['Tree'], ['Apple']))).toBe(Err.value);
    expect(call('DSUM', db, 'Colour', crit(['Tree'], ['Apple']))).toBe(Err.value);
    expect(call('DSUM', db, null, crit(['Tree'], ['Apple']))).toBe(Err.value);
  });

  it('empty results', () => {
    const none = crit(['Tree'], ['=Orange']);
    expect(call('DSUM', db, 'Profit', none)).toBe(0);
    expect(call('DMAX', db, 'Profit', none)).toBe(0);
    expect(call('DMIN', db, 'Profit', none)).toBe(0);
    expect(call('DPRODUCT', db, 'Profit', none)).toBe(0);
    expect(call('DCOUNT', db, 'Profit', none)).toBe(0);
    expect(call('DAVERAGE', db, 'Profit', none)).toBe(Err.div0);
    expect(call('DVAR', db, 'Profit', crit(['Tree'], ['=Cherry']))).toBe(Err.div0);
    expect(call('DVARP', db, 'Profit', crit(['Tree'], ['=Cherry']))).toBe(0);
    expect(call('DSTDEVP', db, 'Profit', none)).toBe(Err.div0);
  });

  it('DCOUNT counts numbers only, DCOUNTA non-blanks; errors propagate in numeric functions', () => {
    const t: Matrix = [['K', 'V'], ['a', 1], ['a', 'x'], ['a', null], ['a', true], ['b', Err.na]];
    expect(call('DCOUNT', t, 'V', crit(['K'], ['a']))).toBe(1);
    expect(call('DCOUNTA', t, 'V', crit(['K'], ['a']))).toBe(3);
    expect(call('DCOUNT', t, crit(['K'], ['a']))).toBe(4);
    expect(call('DSUM', t, 'V', crit(['K'], ['a']))).toBe(1);
    expect(call('DSUM', t, 'V', crit(['K'], [null]))).toBe(Err.na);
    expect(call('DGET', t, 'V', crit(['K'], ['b']))).toBe(Err.na);
  });

  it('criteria need a header and at least one row', () => {
    expect(call('DSUM', db, 'Profit', crit(['Tree']))).toBe(Err.value);
    expect(call('DSUM', Err.ref, 'Profit', crit(['Tree'], ['a']))).toBe(Err.ref);
  });
});
