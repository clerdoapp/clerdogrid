import { expect, it } from 'vitest';
import { Engine } from '../src/engine/engine';
import { registerAllFunctions } from '../src/engine/functions';
import { newWorkbook } from '../src/model/workbook';

it('accepts a function name where a LAMBDA is expected', () => {
  registerAllFunctions();
  const wb = newWorkbook();
  const e = new Engine(wb);
  [['East', 10], ['West', 5], ['East', 7]].forEach(([r, v], i) => { e.setValue(0, i, 0, r as string); e.setValue(0, i, 1, v as number); });
  e.setFormula(0, 0, 3, 'GROUPBY(A1:A3,B1:B3,SUM)');
  e.setFormula(0, 0, 6, 'BYROW(A1:B3,COUNTA)');
  e.recalc();
  expect([e.getValue(0, 0, 3), e.getValue(0, 0, 4), e.getValue(0, 1, 4), e.getValue(0, 0, 6)]).toEqual(['East', 17, 5, 2]);
});
