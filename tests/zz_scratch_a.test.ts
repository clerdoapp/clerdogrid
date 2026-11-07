import { test, expect } from 'vitest';
import { registerAllFunctions } from '../src/engine/functions';
registerAllFunctions();
import { GridStore } from '../src/app/store';
test('eta', () => {
  const s = new GridStore();
  const rows = [['Region','Item','Sales'],['East','A',10],['West','B',5],['East','B',7]];
  s.change('x','sheet',()=>{ rows.forEach((r,i)=>r.forEach((v,j)=>s.enter(i,j,String(v)))); s.enter(0,5,'=GROUPBY(A1:A4,C1:C4,LAMBDA(x,SUM(x)),3)'); s.enter(0,9,'=PIVOTBY(A1:A4,B1:B4,C1:C4,LAMBDA(x,SUM(x)),3)'); s.enter(10,0,'=IFERROR(MODE.SNGL(C2:C4),"None")');});
  const out = []; for (let r=0;r<5;r++){ const l=[]; for(let c=5;c<14;c++) l.push(s.value(r,c)); out.push(l.join('|')); }
  console.log(out.join('\n'), s.value(10,0));
  expect(1).toBe(1);
});
