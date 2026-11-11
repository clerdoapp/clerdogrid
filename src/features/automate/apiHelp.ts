import { N } from '../../ui/i18n';
/** Reference shown next to the script editor. */
export const API_HELP: { title: string; items: [string, string][] }[] = [
  { title: N('Start'), items: [
    ['function main(workbook) { }', N('Every script has a main function. It gets the workbook.')],
    ['console.log(value)', N('Prints to the output below.')],
  ] },
  { title: 'workbook', items: [
    ['getActiveSheet()', N('The sheet that is showing.')],
    ['getSheet(name)', N('A sheet by name, or null.')],
    ['getSheets()', N('All sheets.')],
    ['addSheet(name)', N('Adds a sheet at the end.')],
    ['getSelectedRange()', N('The selected cells.')],
  ] },
  { title: 'sheet', items: [
    ['getRange("A1:C3")', N('A block of cells. Also "A:C" or "2:5".')],
    ['getCell(row, col)', N('One cell, counted from 0.')],
    ['getUsedRange()', N('The block that has data, or null.')],
    ['getName() / setName(name)', N('The sheet name.')],
    ['freezePanes(rows, cols)', N('Keeps rows and columns in view.')],
    ['deleteRows(index, count)', N('Removes whole rows.')],
    ['addTable("A1:D9", name)', N('Turns a block into a table.')],
    ['activate()', N('Shows this sheet when the script ends.')],
  ] },
  { title: 'range', items: [
    ['getValues() / setValues(rows)', N('Values as rows of cells: [[1, 2], [3, 4]].')],
    ['getFormulas() / setFormulas(rows)', N('Formulas such as "=SUM(A1:A9)", or values.')],
    ['getValue() / setValue(v)', N('The first cell. setValue fills the whole range.')],
    ['getAddress()', N('For example "B2:D9".')],
    ['getRowCount() / getColumnCount()', N('Size of the range.')],
    ['getRowIndex() / getColumnIndex()', N('Where it starts, counted from 0.')],
    ['getCell(r, c) / getRow(i) / getColumn(i)', N('Parts of the range.')],
    ['setBold(true) / setItalic(true)', N('Font style.')],
    ['setFontColor("#0C7A4B")', N('Text color.')],
    ['setFill("#E1F3E9")', N('Cell color. setFill(null) removes it.')],
    ['setNumberFormat("0.00%")', N('How numbers show.')],
    ['setHorizontalAlignment("center")', N('left, center, right or general.')],
    ['autofitColumns()', N('Fits column widths to the text.')],
    ['sort(column, ascending)', N('Sorts rows by a column (0 = first in range, or "B").')],
    ['setRowsHidden(b) / setColumnsHidden(b)', N('Hides or shows rows or columns.')],
    ['clear("contents" | "formats")', N('Clears cells. With no argument, clears both.')],
    ['select()', N('Selects the range when the script ends.')],
  ] },
];
