// Writes the sample workbooks in examples/ with the app's own code: npm run examples
import { mkdirSync, writeFileSync } from 'node:fs';
import { TEMPLATES } from '../src/ui/dialogs/a/templates';
import { writeXlsx } from '../src/formats/xlsx';
import { Engine } from '../src/engine/engine';
import { registerAllFunctions } from '../src/engine/functions';
import { cellKey, internStyle, newSheet, newWorkbook, type Workbook } from '../src/model/workbook';

registerAllFunctions();
mkdirSync('examples', { recursive: true });

const save = (name: string, wb: Workbook): void => {
  new Engine(wb).recalc(true);
  writeFileSync(`examples/${name}.xlsx`, writeXlsx(wb));
  console.log(`examples/${name}.xlsx`);
};

for (const t of TEMPLATES) save(t.id, t.make());

// A tour of the calculation engine: dynamic arrays, LET/LAMBDA, lookups and text tools.
const tour = newWorkbook();
const s = tour.sheets[0]!;
s.name = 'Formula tour';
const bold = internStyle(tour, { font: { bold: true } });
const head = internStyle(tour, { font: { bold: true, color: '#FFFFFF' }, fill: { color: '#0C7A4B' } });
const put = (r: number, c: number, v: string | number, style?: number) => {
  const cell = typeof v === 'string' && v.startsWith('=') ? { v: null, f: v.slice(1), s: style } : { v, s: style };
  s.cells.set(cellKey(r, c), cell);
};
['Product', 'Region', 'Units', 'Price'].forEach((h, i) => put(0, i, h, head));
const rows: [string, string, number, number][] = [['Kettle', 'East', 12, 39.5], ['Lamp', 'West', 30, 24], ['Desk', 'East', 4, 210], ['Chair', 'North', 18, 89], ['Lamp', 'East', 22, 24], ['Kettle', 'North', 9, 39.5], ['Desk', 'West', 6, 210]];
rows.forEach((r, i) => r.forEach((v, j) => put(i + 1, j, v)));
const examples: [string, string][] = [
  ['Sorted unique products', '=SORT(UNIQUE(A2:A8))'],
  ['Revenue by region', '=GROUPBY(B2:B8,C2:C8*D2:D8,SUM)'],
  ['Rows with more than 10 units', '=FILTER(A2:C8,C2:C8>10)'],
  ['Price of a desk', '=XLOOKUP("Desk",A2:A8,D2:D8)'],
  ['Total with LET', '=LET(rev,C2:C8*D2:D8,SUM(rev))'],
  ['Custom LAMBDA (tax at 15%)', '=LAMBDA(x,x*1.15)(SUM(C2:C8*D2:D8))'],
  ['Running total of units', '=SCAN(0,C2:C8,LAMBDA(a,b,a+b))'],
  ['Region list as text', '=TEXTJOIN(", ",TRUE,UNIQUE(B2:B8))'],
];
examples.forEach(([label, f], i) => { put(i * 9 + 0, 6, label, bold); put(i * 9 + 1, 6, f); });
s.cols.set(0, { width: 12 }); s.cols.set(6, { width: 28 }); s.cols.set(7, { width: 14 });
s.view.frozenRows = 1;
save('formula-tour', tour);

// A larger data set for trying sorting, filters, PivotTables and charts.
const big = newWorkbook();
const d = big.sheets[0]!;
d.name = 'Sales data';
['Date', 'Region', 'Rep', 'Product', 'Units', 'Unit price', 'Revenue'].forEach((h, i) => d.cells.set(cellKey(0, i), { v: h, s: internStyle(big, { font: { bold: true } }) }));
const reps = ['Ana', 'Ben', 'Chen', 'Dina', 'Eli', 'Fay'], regions = ['East', 'West', 'North', 'South'], products: [string, number][] = [['Kettle', 39.5], ['Lamp', 24], ['Desk', 210], ['Chair', 89], ['Shelf', 65]];
let seed = 7;
const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
const dateFmt = internStyle(big, { numFmt: 'yyyy-mm-dd' }), money = internStyle(big, { numFmt: '"$"#,##0.00' });
for (let r = 1; r <= 500; r++) {
  const [p, price] = products[Math.floor(rnd() * products.length)]!;
  d.cells.set(cellKey(r, 0), { v: 45658 + Math.floor(rnd() * 300), s: dateFmt });
  d.cells.set(cellKey(r, 1), { v: regions[Math.floor(rnd() * 4)]! });
  d.cells.set(cellKey(r, 2), { v: reps[Math.floor(rnd() * 6)]! });
  d.cells.set(cellKey(r, 3), { v: p });
  d.cells.set(cellKey(r, 4), { v: 1 + Math.floor(rnd() * 40) });
  d.cells.set(cellKey(r, 5), { v: price, s: money });
  d.cells.set(cellKey(r, 6), { v: null, f: `E${r + 1}*F${r + 1}`, s: money });
}
d.view.frozenRows = 1;
d.autoFilter = { range: { r1: 0, c1: 0, r2: 500, c2: 6 }, columns: [] };
[12, 10, 10, 10, 8, 11, 12].forEach((w, i) => d.cols.set(i, { width: w }));
big.sheets.push(newSheet('Summary'));
const sum = big.sheets[1]!;
sum.cells.set(cellKey(0, 0), { v: 'Revenue by region and product', s: internStyle(big, { font: { bold: true, size: 14 } }) });
sum.cells.set(cellKey(2, 0), { v: null, f: "PIVOTBY('Sales data'!B2:B501,'Sales data'!D2:D501,'Sales data'!G2:G501,SUM)" });
save('sales-data', big);
