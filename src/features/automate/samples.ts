export interface SampleScript { id: string; name: string; about: string; code: string }

export const SAMPLES: SampleScript[] = [
  {
    id: 'unhide', name: 'Unhide all rows and columns', about: 'Shows every hidden row and column on this sheet.',
    code: `function main(workbook) {
  const sheet = workbook.getActiveSheet();
  const used = sheet.getUsedRange();
  if (!used) return;
  used.setRowsHidden(false);
  used.setColumnsHidden(false);
  console.log("Unhid rows and columns in " + used.getAddress());
}
`,
  },
  {
    id: 'removeEmptyRows', name: 'Remove empty rows', about: 'Deletes rows with nothing in them inside the used range.',
    code: `function main(workbook) {
  const sheet = workbook.getActiveSheet();
  const used = sheet.getUsedRange();
  if (!used) return;
  const values = used.getValues();
  let removed = 0;
  // Go from the bottom up so row numbers stay correct.
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i].every((v) => v === null || v === "")) {
      sheet.deleteRows(used.getRowIndex() + i, 1);
      removed++;
    }
  }
  console.log("Removed " + removed + " empty rows");
}
`,
  },
  {
    id: 'properCase', name: 'Convert text to proper case', about: 'Capitalizes the first letter of each word in the selected cells.',
    code: `function main(workbook) {
  const range = workbook.getSelectedRange();
  const formulas = range.getFormulas();
  const out = formulas.map((row) => row.map((v) => {
    if (typeof v !== "string" || v.startsWith("=")) return v;
    return v.toLowerCase().replace(/(^|[\\s\\-'(])(\\p{L})/gu, (m, a, b) => a + b.toUpperCase());
  }));
  range.setFormulas(out);
}
`,
  },
  {
    id: 'highlightDuplicates', name: 'Highlight duplicates', about: 'Fills cells in the selection whose value appears more than once.',
    code: `function main(workbook) {
  const range = workbook.getSelectedRange();
  const values = range.getValues();
  const counts = new Map();
  for (const row of values) for (const v of row) {
    if (v === null || v === "") continue;
    const key = String(v).toLowerCase();
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  let marked = 0;
  values.forEach((row, r) => row.forEach((v, c) => {
    if (v !== null && v !== "" && counts.get(String(v).toLowerCase()) > 1) {
      range.getCell(r, c).setFill("#FFE08A");
      marked++;
    }
  }));
  console.log(marked + " cells have duplicates");
}
`,
  },
  {
    id: 'totalRow', name: 'Add a total row', about: 'Adds a bold row of SUM formulas under the data on this sheet.',
    code: `function main(workbook) {
  const sheet = workbook.getActiveSheet();
  const used = sheet.getUsedRange();
  if (!used || used.getRowCount() < 2) return;
  const values = used.getValues();
  const top = used.getRowIndex(), left = used.getColumnIndex();
  const last = top + used.getRowCount() - 1;
  const totalRow = last + 1;
  sheet.getCell(totalRow, left).setValue("Total");
  for (let c = 1; c < used.getColumnCount(); c++) {
    const numbers = values.slice(1).some((row) => typeof row[c] === "number");
    if (!numbers) continue;
    const col = sheet.getCell(0, left + c).getAddress().replace(/\\d+/, "");
    sheet.getCell(totalRow, left + c).setFormula("=SUM(" + col + (top + 2) + ":" + col + (last + 1) + ")");
  }
  sheet.getRange(sheet.getCell(totalRow, left).getAddress() + ":" + sheet.getCell(totalRow, left + used.getColumnCount() - 1).getAddress()).setBold(true);
}
`,
  },
  {
    id: 'splitNames', name: 'Split full names into columns', about: 'Splits names in the selected column into first and last name columns to its right.',
    code: `function main(workbook) {
  const range = workbook.getSelectedRange().getColumn(0);
  const values = range.getValues();
  const out = values.map(([v]) => {
    if (typeof v !== "string" || !v.trim()) return [null, null];
    const parts = v.trim().split(/\\s+/);
    const first = parts.shift();
    return [first, parts.join(" ") || null];
  });
  const sheet = range.getSheet();
  const top = range.getRowIndex(), col = range.getColumnIndex();
  const target = sheet.getCell(top, col + 1).getAddress() + ":" + sheet.getCell(top + out.length - 1, col + 2).getAddress();
  sheet.getRange(target).setValues(out);
  console.log("Split " + out.length + " names into " + target);
}
`,
  },
  {
    id: 'numberRows', name: 'Number the rows', about: 'Puts 1, 2, 3… in a new column to the left of the data.',
    code: `function main(workbook) {
  const sheet = workbook.getActiveSheet();
  const used = sheet.getUsedRange();
  if (!used) return;
  const left = used.getColumnIndex();
  const top = used.getRowIndex();
  const rows = used.getRowCount();
  if (left === 0) {
    console.log("Leave column A empty first, so there's room for the numbers.");
    return;
  }
  const numbers = [["#"]];
  for (let i = 1; i < rows; i++) numbers.push([i]);
  const start = sheet.getCell(top, left - 1).getAddress();
  const end = sheet.getCell(top + rows - 1, left - 1).getAddress();
  sheet.getRange(start + ":" + end).setValues(numbers).setHorizontalAlignment("center");
}
`,
  },
  {
    id: 'summarySheet', name: 'Summary of all sheets', about: 'Makes a Summary sheet with the total of the numbers on every sheet.',
    code: `function main(workbook) {
  let summary = workbook.getSheet("Summary");
  if (!summary) summary = workbook.addSheet("Summary");
  const rows = [["Sheet", "Total"]];
  for (const sheet of workbook.getSheets()) {
    if (sheet.getName() === "Summary") continue;
    const used = sheet.getUsedRange();
    let total = 0;
    if (used) for (const row of used.getValues()) for (const v of row) if (typeof v === "number") total += v;
    rows.push([sheet.getName(), total]);
  }
  summary.getRange("A1:B" + rows.length).setValues(rows);
  summary.getRange("A1:B1").setBold(true).setFill("#E1F3E9");
  summary.getRange("B2:B" + rows.length).setNumberFormat("#,##0.00");
  summary.getRange("A:B").autofitColumns();
  summary.activate();
}
`,
  },
];

export const NEW_SCRIPT = `function main(workbook) {
  const sheet = workbook.getActiveSheet();
  const range = workbook.getSelectedRange();
  console.log("Selected " + range.getAddress() + " on " + sheet.getName());
}
`;
