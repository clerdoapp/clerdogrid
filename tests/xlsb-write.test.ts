import { describe, expect, it } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { cellKey, internStyle, newSheet, type Workbook } from '../src/model/workbook';
import { writeXlsb, xlsbLimits } from '../src/formats/xlsb-write';
import { readXlsb, xlsbRecords } from '../src/formats/xlsb-read';
import { workbookFromBytes } from '../src/app/files';

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 4, 0, 9, 251, 3, 253, 227, 85, 242, 156, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

const sample = (): Workbook => {
  const wb: Workbook = { sheets: [newSheet('Data'), newSheet('Other sheet')], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const s = wb.sheets[0]!, o = wb.sheets[1]!;
  const head = internStyle(wb, { font: { bold: true, color: '#C00000', size: 14 }, fill: { color: '#FFFF00' }, border: { bottom: { style: 'double', color: '#0000FF' } }, align: { h: 'center', wrap: true } });
  const kg = internStyle(wb, { numFmt: '#,##0.000 "kg"', fill: { pattern: 'lightGrid', color: '#DDEEFF', patternColor: '#112233' }, protection: { locked: false } });
  s.cells.set(cellKey(0, 0), { v: 'Name', s: head });
  for (let r = 1; r <= 5; r++) { s.cells.set(cellKey(r, 0), { v: `Item ${r}` }); s.cells.set(cellKey(r, 1), { v: r * 1.5, s: kg }); }
  s.cells.set(cellKey(6, 1), { v: 22.5, f: 'SUM(B2:B6)' });
  s.cells.set(cellKey(7, 1), { v: 0.375, f: "B2/'Other sheet'!A1" });
  s.cells.set(cellKey(8, 1), { v: 'yes', f: 'IF(B7>10,"yes","no")' });
  s.cells.set(cellKey(9, 1), { v: 45748, f: 'EDATE(DATE(2025,1,1),3)' });
  s.cells.set(cellKey(10, 1), { v: 3, f: 'IFS(B2>1,3,TRUE,4)' });
  s.cells.set(cellKey(11, 1), { v: 45, f: 'Total*2' });
  s.cells.set(cellKey(12, 0), { v: true });
  s.cells.set(cellKey(12, 1), { v: 'rich', rich: [{ text: 'ri' }, { text: 'ch', font: { italic: true, color: '#00AA00' } }] });
  s.cells.set(cellKey(13, 0), { v: 9, f: 'SUM(B2:B3*2)', arrayRange: { r1: 13, r2: 13, c1: 0, c2: 0 } });
  o.cells.set(0, { v: 4 });
  wb.names.push({ name: 'Total', formula: 'Data!$B$7' }, { name: 'Local', formula: 'Data!$A$1', sheet: 0 });
  s.merges.push({ r1: 15, r2: 16, c1: 0, c2: 2 });
  s.cols.set(0, { width: 20 });
  s.cols.set(3, { width: 12, hidden: true });
  s.rows.set(0, { height: 30 });
  s.view.frozenRows = 1;
  s.view.frozenCols = 1;
  s.view.zoom = 120;
  s.notes.set(0, { author: 'Ann', text: 'A note\nline 2' });
  s.hyperlinks.set(cellKey(1, 0), 'https://example.com/');
  s.hyperlinks.set(cellKey(2, 0), "#'Other sheet'!A1");
  s.conditionalFormats.push({ ranges: [{ r1: 1, r2: 5, c1: 1, c2: 1 }], rules: [
    { type: 'cellIs', operator: 'between', formulas: ['2', '$B$3'], style: { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFC7CE' } } },
    { type: 'expression', formula: 'B2>A$1', style: { fill: { pattern: 'darkGrid', color: '#EEEEEE', patternColor: '#000080' } }, stop: true },
    { type: 'containsText', text: 'Item', style: { font: { italic: true, name: 'Arial' }, numFmt: '0.0%' } },
    { type: 'top10', rank: 2, bottom: true, percent: true, style: { border: { top: { style: 'thin', color: '#00FF00' } } } },
    { type: 'aboveAverage', below: true, equal: true, style: { font: { strike: true } } },
    { type: 'timePeriod', period: 'lastWeek', style: { font: { underline: 'double' } } },
    { type: 'duplicateValues', style: { font: { size: 9 } } },
    { type: 'notContainsBlanks', style: { protection: { hidden: true } } },
    { type: 'colorScale', stops: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'dataBar', min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', showValue: false },
  ] });
  s.validations.push({ ranges: [{ r1: 1, r2: 5, c1: 2, c2: 2 }], type: 'list', formula1: '"Red,Green,Blue"', showDropdown: false, prompt: { title: 'Pick', text: 'Choose a color' } });
  s.validations.push({ ranges: [{ r1: 1, r2: 5, c1: 3, c2: 3 }], type: 'whole', operator: 'between', formula1: '1', formula2: 'B2*10', allowBlank: true, error: { style: 'warning', title: 'Bad', text: 'Out of range' } });
  s.validations.push({ ranges: [{ r1: 1, r2: 5, c1: 4, c2: 4 }, { r1: 8, r2: 9, c1: 4, c2: 4 }], type: 'custom', formula1: 'E2<B2' });
  s.charts.push({ id: 'c1', type: 'column', title: 'Values', series: [{ name: 'Value', values: 'Data!$B$2:$B$6', categories: 'Data!$A$2:$A$6' }], anchor: { row: 2, col: 4, dx: 0, dy: 0, width: 400, height: 240 } });
  s.pictures.push({ id: 'p1', data: PNG, mime: 'image/png', anchor: { row: 20, col: 1, dx: 0, dy: 0, width: 50, height: 50 }, alt: 'Dot' });
  s.shapes = [{ id: 's1', kind: 'rect', text: 'Box', anchor: { row: 20, col: 4, dx: 0, dy: 0, width: 100, height: 60 }, fill: '#AACCEE' }];
  s.cells.set(cellKey(30, 0), { v: 'A' });
  s.cells.set(cellKey(30, 1), { v: 'B' });
  s.tables.push({ name: 'Tbl', range: { r1: 30, r2: 32, c1: 0, c2: 1 }, headerRow: true, totalsRow: false, columns: [{ name: 'A' }, { name: 'B' }] });
  s.autoFilter = { range: { r1: 0, r2: 5, c1: 0, c2: 1 }, columns: [{ col: 0, values: ['Item 1', 'Item 2'] }] };
  s.pageSetup = { orientation: 'landscape', paper: 9, fitWidth: 1, fitHeight: 0, header: '&CTitle', printArea: 'A1:D20', centerH: true, gridlines: true };
  s.protection = { sheet: true, password: 'CC1A', options: { formatCells: false, selectLockedCells: false, sort: true } };
  o.tabColor = '#FF8800';
  o.state = 'hidden';
  wb.calc = { iterate: true, maxIterations: 50, maxChange: 0.01 };
  return wb;
};

describe('.xlsb writing', () => {
  const bytes = writeXlsb(sample());
  const back = workbookFromBytes('book.xlsb', bytes);
  const s = back.sheets[0]!;
  const c = (r: number, col: number) => s.cells.get(cellKey(r, col));

  it('writes a package of binary parts', () => {
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['[Content_Types].xml', '_rels/.rels', 'xl/workbook.bin', 'xl/styles.bin', 'xl/sharedStrings.bin', 'xl/worksheets/sheet1.bin', 'xl/comments1.bin', 'xl/drawings/drawing1.xml', 'xl/charts/chart1.xml']));
    expect(new TextDecoder().decode(files['[Content_Types].xml']!)).toContain('application/vnd.ms-excel.worksheet');
    const types = xlsbRecords(files['xl/worksheets/sheet1.bin']!).map((r) => r.type);
    expect(types[0]).toBe(129);
    expect(types[types.length - 1]).toBe(130);
    expect(types.indexOf(145)).toBeLessThan(types.indexOf(146));
  });
  it('keeps values, rich text and formulas', () => {
    expect(c(0, 0)!.v).toBe('Name');
    expect(c(3, 1)!.v).toBe(4.5);
    expect(c(12, 0)!.v).toBe(true);
    expect(c(12, 1)!.rich).toEqual([{ text: 'ri' }, { text: 'ch', font: { italic: true, color: '#00AA00' } }]);
    expect(c(6, 1)).toMatchObject({ f: 'SUM(B2:B6)', v: 22.5 });
    expect(c(7, 1)!.f).toBe("B2/'Other sheet'!A1");
    expect(c(8, 1)).toMatchObject({ f: 'IF(B7>10,"yes","no")', v: 'yes' });
    expect(c(9, 1)!.f).toBe('EDATE(DATE(2025,1,1),3)');
    expect(c(10, 1)!.f).toBe('IFS(B2>1,3,TRUE,4)');
    expect(c(11, 1)!.f).toBe('Total*2');
    expect(c(13, 0)).toMatchObject({ f: 'SUM(B2:B3*2)', arrayRange: { r1: 13, r2: 13, c1: 0, c2: 0 } });
    expect(back.names).toEqual([{ name: 'Total', formula: 'Data!$B$7' }, { name: 'Local', formula: 'Data!$A$1', sheet: 0 }]);
    expect(back.calc).toMatchObject({ iterate: true, maxIterations: 50, maxChange: 0.01 });
  });
  it('keeps styles', () => {
    expect(back.styles[c(0, 0)!.s!]).toEqual({ font: { size: 14, bold: true, color: '#C00000' }, fill: { color: '#FFFF00' }, border: { bottom: { style: 'double', color: '#0000FF' } }, align: { h: 'center', wrap: true } });
    expect(back.styles[c(1, 1)!.s!]).toEqual({ numFmt: '#,##0.000 "kg"', fill: { pattern: 'lightGrid', color: '#DDEEFF', patternColor: '#112233' }, protection: { locked: false } });
  });
  it('keeps sheet layout, view, notes, links, protection and printing', () => {
    expect(s.merges).toEqual([{ r1: 15, r2: 16, c1: 0, c2: 2 }]);
    expect(s.cols.get(0)!.width).toBeCloseTo(20, 1);
    expect(s.cols.get(3)).toMatchObject({ hidden: true });
    expect(s.rows.get(0)!.height).toBe(30);
    expect(s.view).toMatchObject({ frozenRows: 1, frozenCols: 1, zoom: 120 });
    expect(s.notes.get(0)).toEqual({ author: 'Ann', text: 'A note\nline 2' });
    expect(s.hyperlinks.get(cellKey(1, 0))).toBe('https://example.com/');
    expect(s.hyperlinks.get(cellKey(2, 0))).toBe("#'Other sheet'!A1");
    expect(s.protection).toMatchObject({ sheet: true, password: 'CC1A', options: { formatCells: false, selectLockedCells: false, sort: true, objects: false } });
    expect(s.pageSetup).toMatchObject({ orientation: 'landscape', paper: 9, fitWidth: 1, fitHeight: 0, header: '&CTitle', printArea: 'A1:D20', centerH: true, gridlines: true });
    expect(s.autoFilter).toEqual({ range: { r1: 0, r2: 5, c1: 0, c2: 1 }, columns: [{ col: 0, values: ['Item 1', 'Item 2'] }] });
    const o = back.sheets[1]!;
    expect([o.name, o.state, o.tabColor, o.cells.get(0)!.v]).toEqual(['Other sheet', 'hidden', '#FF8800', 4]);
  });
  it('keeps conditional formats and validation', () => {
    const rules = s.conditionalFormats.flatMap((cf) => cf.rules);
    expect(rules).toEqual(sample().sheets[0]!.conditionalFormats[0]!.rules);
    expect(s.validations).toEqual(sample().sheets[0]!.validations);
  });
  it('keeps charts, pictures, shapes and tables', () => {
    expect(s.charts).toHaveLength(1);
    expect(s.charts[0]).toMatchObject({ type: 'column', title: 'Values', series: [{ name: 'Value', values: 'Data!$B$2:$B$6', categories: 'Data!$A$2:$A$6' }] });
    expect(s.pictures[0]).toMatchObject({ mime: 'image/png', alt: 'Dot', anchor: { row: 20, col: 1 } });
    expect(s.pictures[0]!.data).toEqual(PNG);
    expect(s.shapes?.[0]).toMatchObject({ kind: 'rect', text: 'Box', fill: '#AACCEE' });
    expect(s.tables).toEqual(sample().sheets[0]!.tables);
  });
  it('lists what other apps will not see', () => {
    expect(xlsbLimits(sample())).toEqual(['Tables, color scales, data bars and icon sets are kept for ClerdoGrid only; other apps show the cells without them.']);
  });
  it('reads files with a damaged extras part', () => {
    const files = unzipSync(bytes);
    files['xl/clerdogrid/binaryExtras.json'] = new TextEncoder().encode('{"cf":{"0":[{"ranges":[{"r1":-1}],"rules":[{"type":"colorScale"}]}]},"tables":{"0":[{"name":"x y"}]}}');
    const wb = readXlsb(zipSync(files));
    expect(wb.sheets[0]!.tables).toEqual([]);
    expect(wb.sheets[0]!.conditionalFormats.flatMap((x) => x.rules).some((r) => r.type === 'colorScale')).toBe(false);
  });
  it('keeps custom XML items such as saved data queries', () => {
    const wb = sample();
    const item = new TextEncoder().encode('<queries xmlns="urn:example"/>');
    wb.extraParts = new Map([['customXml/item1.xml', item]]);
    const again = readXlsb(writeXlsb(wb));
    expect(again.extraParts?.get('customXml/item1.xml')).toEqual(item);
  });
});
