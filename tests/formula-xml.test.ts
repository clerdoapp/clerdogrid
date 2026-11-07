import { expect, it } from 'vitest';
import { addFuturePrefixes, stripFuturePrefixes } from '../src/formats/formula-xml';

it('prefixes nested newer functions for the file format', () => {
  expect(addFuturePrefixes('SORT(FILTER(A1:B9,B1:B9>4),2,-1)')).toBe('_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:B9,B1:B9>4),2,-1)');
  expect(addFuturePrefixes('IFS(A1>1,XLOOKUP(1,A:A,B:B),TRUE,0)')).toBe('_xlfn.IFS(A1>1,_xlfn.XLOOKUP(1,A:A,B:B),TRUE,0)');
  expect(stripFuturePrefixes('_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:B9,B1:B9>4))')).toBe('SORT(FILTER(A1:B9,B1:B9>4))');
});
