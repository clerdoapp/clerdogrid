import { Engine } from '../../../engine/engine';
import type { Scalar } from '../../../engine/types';
import { cellKey, colIndex, internStyle, newSheet, newWorkbook, type Cell, type CellStyle, type Chart, type Sheet, type Workbook } from '../../../model/workbook';
import { N } from '../../i18n';

/** A template cell: a value, or a formula written "=...". */
type In = Scalar | undefined;

interface SheetSpec {
  name: string;
  widths: Record<string, number>;
  rows: { at: number; cells: In[]; style?: CellStyle; col?: string }[];
  formats?: { cols: string; from: number; to: number; numFmt: string }[];
  frozen?: number;
  charts?: Chart[];
}

const HEAD: CellStyle = { font: { bold: true, color: '#FFFFFF' }, fill: { color: '#0C7A4B' } };
const TITLE: CellStyle = { font: { bold: true, size: 16 } };
const BOLD: CellStyle = { font: { bold: true } };
const TOTAL: CellStyle = { font: { bold: true }, border: { top: { style: 'thin' } } };
const MONEY = '#,##0.00';
const DATE = 'm/d/yyyy';

const build = (title: string, sheets: SheetSpec[]): Workbook => {
  const wb = newWorkbook();
  wb.sheets = sheets.map((spec) => fillSheet(wb, spec));
  wb.props = { title, created: new Date().toISOString() };
  // Compute every formula once so cached values are right before the store loads it.
  new Engine(wb).recalc(true);
  return wb;
};

const fillSheet = (wb: Workbook, spec: SheetSpec): Sheet => {
  const s = newSheet(spec.name);
  for (const [c, w] of Object.entries(spec.widths)) s.cols.set(colIndex(c), { width: w });
  const styleAt = new Map<number, CellStyle>();
  const addStyle = (k: number, st: CellStyle) => {
    const cur = styleAt.get(k) ?? {};
    styleAt.set(k, { ...cur, ...st, font: cur.font || st.font ? { ...cur.font, ...st.font } : undefined });
  };
  for (const row of spec.rows) {
    const c0 = row.col ? colIndex(row.col) : 0;
    row.cells.forEach((v, j) => {
      if (v === undefined) return;
      const k = cellKey(row.at - 1, c0 + j);
      const cell: Cell = typeof v === 'string' && v.startsWith('=') ? { v: null, f: v.slice(1) } : { v };
      s.cells.set(k, cell);
      if (row.style) addStyle(k, row.style);
    });
  }
  for (const f of spec.formats ?? []) {
    for (const c of f.cols.split(',')) for (let r = f.from; r <= f.to; r++) addStyle(cellKey(r - 1, colIndex(c)), { numFmt: f.numFmt });
  }
  for (const [k, st] of styleAt) {
    const cell = s.cells.get(k) ?? { v: null };
    cell.s = internStyle(wb, st);
    s.cells.set(k, cell);
  }
  if (spec.frozen) s.view.frozenRows = spec.frozen;
  s.charts = spec.charts ?? [];
  return s;
};

const nums = (from: number, to: number): number[] => Array.from({ length: to - from + 1 }, (_, i) => from + i);

// ---------- The templates ----------

const budget = (): Workbook => {
  const income = [['Salary', 4200, 4200], ['Side work', 600, 450], ['Other', 0, 120]];
  const spend = [['Rent', 1400, 1400], ['Groceries', 450, 512], ['Utilities', 180, 164], ['Transport', 160, 141], ['Phone and internet', 90, 90], ['Eating out', 200, 265], ['Health', 80, 40], ['Savings', 800, 800], ['Fun', 150, 188]];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Monthly budget'], style: TITLE },
    { at: 3, cells: ['Income', 'Planned', 'Actual', 'Difference'], style: HEAD },
    ...income.map((r, i) => ({ at: 4 + i, cells: [...r, `=C${4 + i}-B${4 + i}`] })),
    { at: 7, cells: ['Total income', '=SUM(B4:B6)', '=SUM(C4:C6)', '=C7-B7'], style: TOTAL },
    { at: 9, cells: ['Spending', 'Planned', 'Actual', 'Difference'], style: HEAD },
    ...spend.map((r, i) => ({ at: 10 + i, cells: [...r, `=B${10 + i}-C${10 + i}`] })),
    { at: 19, cells: ['Total spending', '=SUM(B10:B18)', '=SUM(C10:C18)', '=B19-C19'], style: TOTAL },
    { at: 21, cells: ['Left over', '=B7-B19', '=C7-C19', '=C21-B21'], style: TOTAL },
  ];
  return build('Monthly budget', [{ name: 'Budget', widths: { A: 24, B: 13, C: 13, D: 13 }, rows, formats: [{ cols: 'B,C,D', from: 4, to: 21, numFmt: MONEY }], frozen: 3 }]);
};

const invoice = (): Workbook => {
  const items = [['Design work', 12, 65], ['Development', 30, 80], ['Hosting (yearly)', 1, 240]];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Invoice'], style: TITLE },
    { at: 3, cells: ['Invoice number', 'INV-0001'] },
    { at: 4, cells: ['Date', '=TODAY()'] },
    { at: 5, cells: ['Due', '=B4+30'] },
    { at: 7, cells: ['Bill to'], style: BOLD },
    { at: 8, cells: ['Customer name'] },
    { at: 9, cells: ['Street address'] },
    { at: 10, cells: ['City, postcode'] },
    { at: 12, cells: ['Description', 'Quantity', 'Unit price', 'Amount'], style: HEAD },
    ...nums(13, 22).map((r) => ({ at: r, cells: [...(items[r - 13] ?? [undefined, undefined, undefined]), `=IF(B${r}="","",B${r}*C${r})`] })),
    { at: 24, cells: ['Subtotal', '=SUM(D13:D22)'], col: 'C', style: BOLD },
    { at: 25, cells: ['Tax rate', 0.1], col: 'C' },
    { at: 26, cells: ['Tax', '=D24*D25'], col: 'C' },
    { at: 27, cells: ['Total due', '=D24+D26'], col: 'C', style: TOTAL },
  ];
  return build('Invoice', [{
    name: 'Invoice', widths: { A: 34, B: 12, C: 14, D: 14 }, rows,
    formats: [{ cols: 'B', from: 4, to: 5, numFmt: DATE }, { cols: 'C,D', from: 13, to: 22, numFmt: MONEY }, { cols: 'D', from: 24, to: 24, numFmt: MONEY }, { cols: 'D', from: 25, to: 25, numFmt: '0%' }, { cols: 'D', from: 26, to: 27, numFmt: MONEY }],
  }]);
};

const loan = (): Workbook => {
  const n = 60;
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Loan calculator'], style: TITLE },
    { at: 3, cells: ['Loan amount', 25000] },
    { at: 4, cells: ['Yearly interest rate', 0.065] },
    { at: 5, cells: ['Years', 5] },
    { at: 6, cells: ['Payments per year', 12] },
    { at: 8, cells: ['Payment', '=PMT(B4/B6,B5*B6,-B3)'], style: BOLD },
    { at: 9, cells: ['Total paid', '=B8*B5*B6'] },
    { at: 10, cells: ['Total interest', '=B9-B3'] },
    { at: 12, cells: ['Number', 'Payment', 'Interest', 'Principal', 'Balance'], style: HEAD },
    ...nums(1, n).map((i) => {
      const r = 12 + i;
      const prev = i === 1 ? '$B$3' : `E${r - 1}`;
      return { at: r, cells: [`=IF(${i}<=$B$5*$B$6,${i},"")`, `=IF(A${r}="","",$B$8)`, `=IF(A${r}="","",${prev}*$B$4/$B$6)`, `=IF(A${r}="","",B${r}-C${r})`, `=IF(A${r}="","",${prev}-D${r})`] as In[] };
    }),
  ];
  return build('Loan calculator', [{
    name: 'Loan', widths: { A: 22, B: 14, C: 14, D: 14, E: 14 }, rows, frozen: 12,
    formats: [{ cols: 'B', from: 3, to: 3, numFmt: MONEY }, { cols: 'B', from: 4, to: 4, numFmt: '0.00%' }, { cols: 'B', from: 8, to: 10, numFmt: MONEY }, { cols: 'B,C,D,E', from: 13, to: 12 + n, numFmt: MONEY }],
  }]);
};

const todo = (): Workbook => {
  const tasks: In[][] = [['Plan the week', 'High', 0, 'Done'], ['Reply to messages', 'Medium', 1, 'In progress'], ['Book appointment', 'Low', 3, 'Not started'], ['Prepare report', 'High', 2, 'Not started']];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Task', 'Priority', 'Due', 'Status', 'Days left'], style: HEAD },
    ...nums(2, 21).map((r) => {
      const t = tasks[r - 2];
      return { at: r, cells: [t?.[0], t?.[1], t ? `=TODAY()+${t[2]}` : undefined, t?.[3], `=IF(OR(C${r}="",D${r}="Done"),"",C${r}-TODAY())`] as In[] };
    }),
    { at: 23, cells: ['Done', '=COUNTIF(D2:D21,"Done")'], style: BOLD },
    { at: 24, cells: ['Open', '=COUNTA(A2:A21)-B23'], style: BOLD },
  ];
  const wb = build('To-do list', [{ name: 'Tasks', widths: { A: 34, B: 11, C: 12, D: 14, E: 10 }, rows, frozen: 1, formats: [{ cols: 'C', from: 2, to: 21, numFmt: DATE }] }]);
  const s = wb.sheets[0]!;
  s.validations.push({ ranges: [{ r1: 1, c1: 3, r2: 20, c2: 3 }], type: 'list', formula1: '"Not started,In progress,Done"', showDropdown: true, allowBlank: true });
  s.validations.push({ ranges: [{ r1: 1, c1: 1, r2: 20, c2: 1 }], type: 'list', formula1: '"High,Medium,Low"', showDropdown: true, allowBlank: true });
  s.conditionalFormats.push({ ranges: [{ r1: 1, c1: 0, r2: 20, c2: 4 }], rules: [{ type: 'expression', formula: '$D2="Done"', style: { font: { strike: true, color: '#8A90A0' } } }] });
  return wb;
};

const sales = (): Workbook => {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const data = [[12400, 9800, 4300], [13100, 10400, 4700], [15800, 11200, 5200], [14900, 12100, 5600], [16700, 12800, 6100], [18200, 13500, 6400], [17600, 14200, 6900], [19300, 14800, 7300], [18800, 15100, 7700], [20500, 16000, 8200], [22800, 17400, 8800], [25600, 19200, 9600]];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Month', 'North', 'South', 'Online', 'Total'], style: HEAD },
    ...months.map((m, i) => ({ at: i + 2, cells: [m, ...data[i]!, `=SUM(B${i + 2}:D${i + 2})`] as In[] })),
    { at: 14, cells: ['Total', '=SUM(B2:B13)', '=SUM(C2:C13)', '=SUM(D2:D13)', '=SUM(E2:E13)'], style: TOTAL },
    { at: 15, cells: ['Share', '=B14/$E$14', '=C14/$E$14', '=D14/$E$14', '=E14/$E$14'] },
  ];
  const ref = (c: string) => `Sales!$${c}$2:$${c}$13`;
  const chart: Chart = {
    id: 'chart-sales', type: 'line', title: N('Sales by month'),
    series: ['B', 'C', 'D'].map((c) => ({ name: `Sales!$${c}$1`, values: ref(c), categories: ref('A') })),
    anchor: { row: 0, col: 6, dx: 12, dy: 0, width: 520, height: 300 }, legend: 'bottom',
  };
  return build('Sales report', [{
    name: 'Sales', widths: { A: 10, B: 12, C: 12, D: 12, E: 13 }, rows, frozen: 1, charts: [chart],
    formats: [{ cols: 'B,C,D,E', from: 2, to: 14, numFmt: '#,##0' }, { cols: 'B,C,D,E', from: 15, to: 15, numFmt: '0.0%' }],
  }]);
};

const attendance = (): Workbook => {
  const days = 20;
  const people = ['Alex', 'Sam', 'Jordan', 'Taylor', 'Morgan', 'Riley'];
  const lastCol = 1 + days;
  const col = (i: number) => { let s = ''; let n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const L = col(lastCol);
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Attendance'], style: TITLE },
    { at: 2, cells: ['Enter P for present, A for absent, L for late.'] },
    { at: 4, cells: ['Name', ...nums(1, days).map((d) => `Day ${d}`), 'Present', 'Absent', 'Late', 'Rate'], style: HEAD },
    ...people.map((p, i) => {
      const r = 5 + i;
      const marks = nums(1, days).map((d) => ((d * 7 + i * 3) % 11 === 0 ? 'A' : (d + i) % 9 === 0 ? 'L' : 'P'));
      return { at: r, cells: [p, ...marks, `=COUNTIF(B${r}:${L}${r},"P")`, `=COUNTIF(B${r}:${L}${r},"A")`, `=COUNTIF(B${r}:${L}${r},"L")`, `=IF(COUNTA(B${r}:${L}${r})=0,"",(${col(lastCol + 1)}${r}+${col(lastCol + 3)}${r})/COUNTA(B${r}:${L}${r}))`] as In[] };
    }),
  ];
  const widths: Record<string, number> = { A: 16 };
  for (let i = 1; i <= days; i++) widths[col(i)] = 6;
  for (let i = 1; i <= 4; i++) widths[col(lastCol + i)] = 9;
  const wb = build('Attendance', [{ name: 'Attendance', widths, rows, frozen: 4, formats: [{ cols: col(lastCol + 4), from: 5, to: 4 + people.length, numFmt: '0%' }] }]);
  wb.sheets[0]!.view.frozenCols = 1;
  return wb;
};

const expenses = (): Workbook => {
  const items: In[][] = [[-6, 'Travel', 'Train tickets', 84.5], [-5, 'Meals', 'Client lunch', 46.2], [-3, 'Lodging', 'Hotel, 1 night', 129], [-2, 'Supplies', 'Printer paper', 18.99]];
  const cats = ['Travel', 'Meals', 'Lodging', 'Supplies', 'Other'];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Expense report'], style: TITLE },
    { at: 2, cells: ['Name', 'Your name'] },
    { at: 3, cells: ['Period', 'This month'] },
    { at: 5, cells: ['Date', 'Category', 'Description', 'Amount'], style: HEAD },
    ...nums(6, 25).map((r) => { const it = items[r - 6]; return { at: r, cells: it ? [`=TODAY()${it[0]}`, it[1], it[2], it[3]] : [] }; }),
    { at: 26, cells: ['Total', undefined, undefined, '=SUM(D6:D25)'], style: TOTAL },
    { at: 28, cells: ['By category', 'Amount'], col: 'F', style: HEAD },
    ...cats.map((c, i) => ({ at: 29 + i, cells: [c, `=SUMIF($B$6:$B$25,F${29 + i},$D$6:$D$25)`], col: 'F' })),
  ];
  const wb = build('Expense report', [{
    name: 'Expenses', widths: { A: 12, B: 14, C: 30, D: 12, E: 3, F: 14, G: 12 }, rows,
    formats: [{ cols: 'A', from: 6, to: 25, numFmt: DATE }, { cols: 'D', from: 6, to: 26, numFmt: MONEY }, { cols: 'G', from: 29, to: 33, numFmt: MONEY }],
  }]);
  wb.sheets[0]!.validations.push({ ranges: [{ r1: 5, c1: 1, r2: 24, c2: 1 }], type: 'list', formula1: `"${cats.join(',')}"`, showDropdown: true, allowBlank: true });
  return wb;
};

const grades = (): Workbook => {
  const students = ['Avery', 'Blake', 'Casey', 'Drew', 'Emery', 'Finley', 'Harper', 'Jules'];
  const scores = [[88, 92, 79, 85], [72, 68, 81, 77], [95, 98, 91, 94], [64, 71, 58, 69], [83, 79, 88, 90], [91, 85, 87, 82], [77, 82, 74, 80], [69, 75, 83, 71]];
  const rows: SheetSpec['rows'] = [
    { at: 1, cells: ['Student', 'Test 1', 'Test 2', 'Project', 'Final', 'Average', 'Grade'], style: HEAD },
    ...students.map((s, i) => {
      const r = i + 2;
      return { at: r, cells: [s, ...scores[i]!, `=AVERAGE(B${r}:E${r})`, `=IF(F${r}>=90,"A",IF(F${r}>=80,"B",IF(F${r}>=70,"C",IF(F${r}>=60,"D","F"))))`] as In[] };
    }),
    { at: 11, cells: ['Class average', '=AVERAGE(B2:B9)', '=AVERAGE(C2:C9)', '=AVERAGE(D2:D9)', '=AVERAGE(E2:E9)', '=AVERAGE(F2:F9)'], style: TOTAL },
    { at: 12, cells: ['Highest', '=MAX(B2:B9)', '=MAX(C2:C9)', '=MAX(D2:D9)', '=MAX(E2:E9)', '=MAX(F2:F9)'] },
    { at: 13, cells: ['Lowest', '=MIN(B2:B9)', '=MIN(C2:C9)', '=MIN(D2:D9)', '=MIN(E2:E9)', '=MIN(F2:F9)'] },
  ];
  return build('Grade book', [{ name: 'Grades', widths: { A: 16, B: 9, C: 9, D: 9, E: 9, F: 10, G: 8 }, rows, frozen: 1, formats: [{ cols: 'B,C,D,E,F', from: 11, to: 11, numFmt: '0.0' }, { cols: 'F', from: 2, to: 9, numFmt: '0.0' }] }]);
};

export interface Template { id: string; name: string; about: string; icon: 'budget' | 'invoice' | 'loan' | 'todo' | 'sales' | 'attendance' | 'expenses' | 'grades'; make: () => Workbook }

export const TEMPLATES: Template[] = [
  { id: 'budget', name: N('Monthly budget'), about: N('Plan income and spending, and see what is left.'), icon: 'budget', make: budget },
  { id: 'invoice', name: N('Invoice'), about: N('Line items with tax and a total due.'), icon: 'invoice', make: invoice },
  { id: 'loan', name: N('Loan calculator'), about: N('Monthly payment and a full payment schedule.'), icon: 'loan', make: loan },
  { id: 'todo', name: N('To-do list'), about: N('Tasks with priority, due date and status.'), icon: 'todo', make: todo },
  { id: 'sales', name: N('Sales report'), about: N('Monthly sales by region with a chart.'), icon: 'sales', make: sales },
  { id: 'attendance', name: N('Attendance'), about: N('Daily marks with totals and a rate.'), icon: 'attendance', make: attendance },
  { id: 'expenses', name: N('Expense report'), about: N('Receipts with totals by category.'), icon: 'expenses', make: expenses },
  { id: 'grades', name: N('Grade book'), about: N('Scores, averages and letter grades.'), icon: 'grades', make: grades },
];
