import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { cellKey, internStyle, newSheet, newWorkbook, type Workbook } from '../src/model/workbook';
import { CellError } from '../src/engine/types';
import { writeOds } from '../src/formats/ods-write';
import { readOds } from '../src/formats/ods-read';
import { fromOdfFormula, toOdfFormula } from '../src/formats/ods-formula';
import { dataStyleXml, readDataStyle } from '../src/formats/ods-numfmt';
import { parseOdf, findDeep, elems, a } from '../src/formats/ods-xml';

const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEElEQVR4nGP4z8AARwzEcQCukw/x0F8jngAAAABJRU5ErkJggg==', 'base64'));

const sample = (): Workbook => {
  const wb = newWorkbook();
  const s = wb.sheets[0]!;
  s.name = 'Data';
  const other = newSheet('Other Sheet');
  const hidden = newSheet('Hidden');
  hidden.state = 'hidden';
  hidden.tabColor = '#FF0000';
  wb.sheets.push(other, hidden);
  for (let r = 0; r < 10; r++) for (let c = 0; c < 4; c++) s.cells.set(cellKey(r, c), { v: r * 4 + c + 0.5 });
  other.cells.set(0, { v: 'a' });
  hidden.cells.set(0, { v: 42 });
  s.cells.set(cellKey(0, 6), { v: 185, f: 'SUM(A1:A10)' });
  s.cells.set(cellKey(1, 6), { v: 'ax', f: "'Other Sheet'!A1&\"x\"" });
  s.cells.set(cellKey(2, 6), { v: 84, f: 'Hidden!$A$1*2' });
  s.cells.set(cellKey(3, 6), { v: 'err', f: 'IFERROR(1/0,"err")' });
  s.cells.set(cellKey(4, 6), { v: 3, f: 'IFS(A1>0,3,TRUE,4)' });
  s.cells.set(cellKey(5, 6), { v: 1, f: 'myName*2' });
  s.cells.set(cellKey(6, 6), { v: true, f: 'AND(A1>0,NOT(B1=1))' });
  s.cells.set(cellKey(7, 6), { v: 4, f: 'MAX(SUM(A1:B2),{1,2;3,4})' });
  s.cells.set(cellKey(8, 6), { v: 2, f: 'NORMSDIST(0)+STDEV.S(A1:A3)+SUM((A1,B2))+SUM(A1:B2 B1:C3)' });
  s.cells.set(cellKey(9, 6), { v: new CellError('#N/A'), f: 'VLOOKUP("x",A1:D10,2,FALSE)' });
  s.cells.set(cellKey(10, 6), { v: 9, f: 'SUM(A:A)+SUM(1:2)' });
  s.cells.set(cellKey(0, 8), { v: 2, f: 'A1:A2*2', arrayRange: { r1: 0, c1: 8, r2: 1, c2: 8 } });
  wb.names.push({ name: 'myName', formula: 'Data!$A$1:$A$3' }, { name: 'local', formula: 'Data!$B$2', sheet: 1 }, { name: 'rate', formula: '0.05*2' });
  s.merges.push({ r1: 12, c1: 0, r2: 13, c2: 2 });
  s.view.frozenRows = 1;
  s.view.frozenCols = 2;
  s.view.zoom = 125;
  s.view.showGrid = false;
  s.notes.set(cellKey(0, 0), { author: 'Ann', text: 'A note\nline 2' });
  s.hyperlinks.set(cellKey(1, 0), 'https://example.com/x');
  s.hyperlinks.set(cellKey(2, 0), "#'Other Sheet'!A1");
  s.cols.set(0, { width: 20 });
  s.cols.set(3, { hidden: true, width: 5 });
  s.rows.set(4, { height: 30 });
  s.rows.set(5, { hidden: true });
  const st = internStyle(wb, {
    font: { name: 'Arial', size: 14, bold: true, italic: true, color: '#123456', underline: 'double', strike: true }, fill: { color: '#FFFF00' },
    border: { left: { style: 'thin' }, bottom: { style: 'double', color: '#0000FF' }, top: { style: 'medium' }, right: { style: 'dashed', color: '#00FF00' } },
    align: { h: 'center', v: 'top', wrap: true, rotation: 45 }, numFmt: '0.000', protection: { locked: false },
  });
  s.cells.get(cellKey(0, 1))!.s = st;
  s.cells.set(cellKey(20, 0), { v: 'bold it', rich: [{ text: 'bold', font: { bold: true } }, { text: ' it', font: { italic: true, color: '#00FF00' } }] });
  s.cells.set(cellKey(21, 0), { v: true });
  s.cells.set(cellKey(22, 0), { v: 45000.25, s: internStyle(wb, { numFmt: 'yyyy-mm-dd hh:mm' }) });
  s.cells.set(cellKey(23, 0), { v: 'Unicode – ünïcødé 漢字  two spaces\ttab\nnew line' });
  s.cells.set(cellKey(24, 0), { v: 0.125, s: internStyle(wb, { numFmt: '0.0%' }) });
  s.cells.set(cellKey(25, 0), { v: -1234.5, s: internStyle(wb, { numFmt: '"$"#,##0.00;[Red]\\-"$"#,##0.00' }) });
  s.cells.set(cellKey(26, 0), { v: 1.5, s: internStyle(wb, { numFmt: '[h]:mm:ss' }) });
  s.cells.set(cellKey(27, 0), { v: 0.75, s: internStyle(wb, { numFmt: 'h:mm AM/PM' }) });
  s.cells.set(cellKey(28, 0), { v: new CellError('#DIV/0!') });
  s.conditionalFormats.push({ ranges: [{ r1: 0, c1: 1, r2: 9, c2: 1 }], rules: [
    { type: 'cellIs', operator: 'between', formulas: ['10', '20'], style: { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFCCCC' } } },
    { type: 'expression', formula: 'MOD(ROW(),2)=0', style: { fill: { color: '#CCFFCC' } } },
    { type: 'colorScale', stops: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'containsText', text: 'x', style: { font: { italic: true } } },
  ] });
  s.validations.push({ ranges: [{ r1: 0, c1: 9, r2: 5, c2: 9 }], type: 'list', formula1: '"red,green,blue"', showDropdown: true, prompt: { title: 'Pick', text: 'Choose' }, error: { style: 'warning', title: 'Bad', text: 'No' } });
  s.validations.push({ ranges: [{ r1: 0, c1: 10, r2: 5, c2: 10 }], type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true });
  s.validations.push({ ranges: [{ r1: 0, c1: 11, r2: 2, c2: 11 }], type: 'decimal', operator: 'greaterThan', formula1: '$A$1' });
  s.autoFilter = { range: { r1: 0, c1: 0, r2: 9, c2: 3 }, columns: [{ col: 1, custom: { and: true, rules: [{ operator: 'greaterThan', value: '5' }] } }] };
  s.pageSetup = { orientation: 'landscape', paper: 9, printArea: '$A$1:$D$10', printTitleRows: '$1:$1', margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 }, header: '&CTitle', footer: '&LLeft&RRight' };
  s.protection = { sheet: true, password: 'CC3D' };
  s.pictures.push({ id: 'p1', data: PNG, mime: 'image/png', anchor: { row: 15, col: 4, dx: 3, dy: 2, width: 120, height: 80 }, alt: 'logo' });
  s.charts.push({ id: 'c1', type: 'stackedColumn', title: 'Sales', series: [{ name: 'Data!$B$1', values: 'Data!$B$2:$B$10', categories: 'Data!$A$2:$A$10' }, { values: 'Data!$C$2:$C$10', categories: 'Data!$A$2:$A$10' }], anchor: { row: 2, col: 13, dx: 0, dy: 0, width: 400, height: 250 }, legend: 'bottom', axisTitles: { x: 'Month', y: 'Amount' } });
  wb.activeSheet = 0;
  wb.props = { title: 'Report', creator: 'Ann' };
  return wb;
};

describe('.ods save and reopen', () => {
  const back = readOds(writeOds(sample()));
  const s = back.sheets[0]!;
  const c = (r: number, col = 6) => s.cells.get(cellKey(r, col))!;

  it('writes a valid package', () => {
    const files = unzipSync(writeOds(sample()));
    expect(Object.keys(files)[0]).toBe('mimetype');
    expect(strFromU8(files.mimetype!)).toBe('application/vnd.oasis.opendocument.spreadsheet');
    expect(strFromU8(files['META-INF/manifest.xml']!)).toContain('Object 1/');
    const tpl = unzipSync(writeOds(sample(), { template: true }));
    expect(strFromU8(tpl.mimetype!)).toBe('application/vnd.oasis.opendocument.spreadsheet-template');
  });

  it('keeps sheets, states, tab colours and properties', () => {
    expect(back.sheets.map((x) => x.name)).toEqual(['Data', 'Other Sheet', 'Hidden']);
    expect(back.sheets[2]!.state).toBe('hidden');
    expect(back.sheets[2]!.tabColor).toBe('#FF0000');
    expect(back.props).toMatchObject({ title: 'Report', creator: 'Ann' });
  });

  it('keeps values of every type', () => {
    expect(s.cells.get(cellKey(3, 2))!.v).toBe(14.5);
    expect(s.cells.get(cellKey(21, 0))!.v).toBe(true);
    expect(s.cells.get(cellKey(22, 0))!.v).toBeCloseTo(45000.25, 8);
    expect(s.cells.get(cellKey(23, 0))!.v).toBe('Unicode – ünïcødé 漢字  two spaces\ttab\nnew line');
    expect(s.cells.get(cellKey(24, 0))!.v).toBe(0.125);
    expect(s.cells.get(cellKey(25, 0))!.v).toBe(-1234.5);
    expect(s.cells.get(cellKey(26, 0))!.v).toBeCloseTo(1.5, 8);
    expect((s.cells.get(cellKey(28, 0))!.v as CellError).code).toBe('#DIV/0!');
  });

  it('keeps formulas and array formulas', () => {
    expect(c(0).f).toBe('SUM(A1:A10)');
    expect(c(0).v).toBe(185);
    expect(c(1).f).toBe("'Other Sheet'!A1&\"x\"");
    expect(c(1).v).toBe('ax');
    expect(c(2).f).toBe('Hidden!$A$1*2');
    expect(c(3).f).toBe('IFERROR(1/0,"err")');
    expect(c(4).f).toBe('IFS(A1>0,3,TRUE,4)');
    expect(c(5).f).toBe('myName*2');
    expect(c(6).f).toBe('AND(A1>0,NOT(B1=1))');
    expect(c(7).f).toBe('MAX(SUM(A1:B2),{1,2;3,4})');
    expect(c(8).f).toBe('NORMSDIST(0)+STDEV.S(A1:A3)+SUM((A1,B2))+SUM(A1:B2 B1:C3)');
    expect(c(10).f).toBe('SUM(A:A)+SUM(1:2)');
    expect(c(0, 8)).toMatchObject({ f: 'A1:A2*2', arrayRange: { r1: 0, c1: 8, r2: 1, c2: 8 } });
  });

  it('keeps styles, number formats and rich text', () => {
    const st = back.styles[c(0, 1).s!]!;
    expect(st.font).toEqual({ name: 'Arial', size: 14, bold: true, italic: true, color: '#123456', underline: 'double', strike: true });
    expect(st.fill).toEqual({ color: '#FFFF00' });
    expect(st.border).toEqual({ left: { style: 'thin' }, bottom: { style: 'double', color: '#0000FF' }, top: { style: 'medium' }, right: { style: 'dashed', color: '#00FF00' } });
    expect(st.align).toEqual({ h: 'center', v: 'top', wrap: true, rotation: 45 });
    expect(st.numFmt).toBe('0.000');
    expect(st.protection).toEqual({ locked: false });
    const fmt = (r: number) => back.styles[s.cells.get(cellKey(r, 0))!.s!]!.numFmt;
    expect(fmt(22)).toBe('yyyy-mm-dd hh:mm');
    expect(fmt(24)).toBe('0.0%');
    expect(fmt(25)).toBe('"$"#,##0.00;[Red]\\-"$"#,##0.00');
    expect(fmt(26)).toBe('[h]:mm:ss');
    expect(s.cells.get(cellKey(20, 0))!.rich).toEqual([{ text: 'bold', font: { bold: true } }, { text: ' it', font: { italic: true, color: '#00FF00' } }]);
  });

  it('keeps layout, view and names', () => {
    expect(s.merges).toEqual([{ r1: 12, c1: 0, r2: 13, c2: 2 }]);
    expect(s.cols.get(0)!.width).toBeCloseTo(20, 1);
    expect(s.cols.get(3)).toMatchObject({ hidden: true });
    expect(s.rows.get(4)!.height).toBe(30);
    expect(s.rows.get(5)!.hidden).toBe(true);
    expect(s.view).toMatchObject({ frozenRows: 1, frozenCols: 2, zoom: 125, showGrid: false });
    expect(back.names).toEqual(expect.arrayContaining([{ name: 'myName', formula: 'Data!$A$1:$A$3' }, { name: 'local', formula: 'Data!$B$2', sheet: 1 }, { name: 'rate', formula: '0.05*2' }]));
  });

  it('keeps notes, links, conditional formats, validation, filter, protection and page setup', () => {
    expect(s.notes.get(0)).toEqual({ author: 'Ann', text: 'A note\nline 2' });
    expect(s.hyperlinks.get(cellKey(1, 0))).toBe('https://example.com/x');
    expect(s.hyperlinks.get(cellKey(2, 0))).toBe("#'Other Sheet'!A1");
    const rules = s.conditionalFormats[0]!.rules;
    expect(s.conditionalFormats[0]!.ranges).toEqual([{ r1: 0, r2: 9, c1: 1, c2: 1 }]);
    expect(rules[0]).toEqual({ type: 'cellIs', operator: 'between', formulas: ['10', '20'], style: { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFCCCC' } } });
    expect(rules[1]).toEqual({ type: 'expression', formula: 'MOD(ROW(),2)=0', style: { fill: { color: '#CCFFCC' } } });
    expect(rules[2]).toEqual({ type: 'colorScale', stops: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] });
    expect(rules[3]).toEqual({ type: 'containsText', text: 'x', style: { font: { italic: true } } });
    const dv = (col: number) => s.validations.find((v) => v.ranges[0]!.c1 === col)!;
    expect(dv(9)).toEqual({ ranges: [{ r1: 0, r2: 5, c1: 9, c2: 9 }], type: 'list', formula1: '"red,green,blue"', showDropdown: true, prompt: { title: 'Pick', text: 'Choose' }, error: { style: 'warning', title: 'Bad', text: 'No' } });
    expect(dv(10)).toMatchObject({ type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true });
    expect(dv(11)).toMatchObject({ type: 'decimal', operator: 'greaterThan', formula1: '$A$1', ranges: [{ r1: 0, r2: 2, c1: 11, c2: 11 }] });
    expect(s.autoFilter).toEqual({ range: { r1: 0, r2: 9, c1: 0, c2: 3 }, columns: [{ col: 1, custom: { and: true, rules: [{ operator: 'greaterThan', value: '5' }] } }] });
    expect(s.protection).toMatchObject({ sheet: true, password: 'CC3D' });
    expect(s.pageSetup).toMatchObject({ orientation: 'landscape', paper: 9, printArea: '$A$1:$D$10', printTitleRows: '$1:$1', header: '&CTitle', footer: '&LLeft&RRight' });
    expect(s.pageSetup?.margins).toMatchObject({ left: 0.5, right: 0.5, top: 1, bottom: 1 });
  });

  it('keeps pictures and charts', () => {
    expect(s.pictures).toHaveLength(1);
    expect(s.pictures[0]).toMatchObject({ mime: 'image/png', alt: 'logo', anchor: { row: 15, col: 4, dx: 3, dy: 2, width: 120, height: 80 } });
    expect(s.pictures[0]!.data).toEqual(PNG);
    expect(s.charts).toHaveLength(1);
    expect(s.charts[0]).toMatchObject({ type: 'stackedColumn', title: 'Sales', legend: 'bottom', axisTitles: { x: 'Month', y: 'Amount' }, anchor: { row: 2, col: 13, width: 400, height: 250 } });
    expect(s.charts[0]!.series).toEqual([{ name: 'Data!$B$1', values: 'Data!$B$2:$B$10', categories: 'Data!$A$2:$A$10' }, { values: 'Data!$C$2:$C$10', categories: 'Data!$A$2:$A$10' }]);
  });

  it('writes a 1904 date system', () => {
    const w = newWorkbook();
    w.date1904 = true;
    w.sheets[0]!.cells.set(0, { v: 100, s: internStyle(w, { numFmt: 'yyyy-mm-dd' }) });
    const r = readOds(writeOds(w));
    expect(r.date1904).toBe(true);
    expect(r.sheets[0]!.cells.get(0)!.v).toBe(100);
  });
});

describe('OpenFormula', () => {
  const cases: [string, string][] = [
    ['SUM(A1:B2)', 'of:=SUM([.A1:.B2])'],
    ["'My Sheet'!$A$1+Data!B2", "of:=[$'My Sheet'.$A$1]+[$Data.B2]"],
    ['Sheet1:Sheet3!A1', 'of:=[$Sheet1.A1:$Sheet3.A1]'],
    ['IF(A1>1,"a""b",FALSE)', 'of:=IF([.A1]>1;"a""b";FALSE())'],
    ['{1,2;3,4}', 'of:={1;2|3;4}'],
    ['SUM((A1,B2))', 'of:=SUM(([.A1]~[.B2]))'],
    ['A1:B2 B2:C3', 'of:=[.A1:.B2]![.B2:.C3]'],
    ['XLOOKUP(1,A:A,B:B)', 'of:=COM.MICROSOFT.XLOOKUP(1;[.A:.A];[.B:.B])'],
    ['CHIDIST(1,2)', 'of:=LEGACY.CHIDIST(1;2)'],
    ['ERROR.TYPE(A1)', 'of:=ERROR.TYPE([.A1])'],
    ['NORM.S.DIST(1,TRUE)', 'of:=COM.MICROSOFT.NORM.S.DIST(1;TRUE())'],
    ['SUM(2:3)', 'of:=SUM([.2:.3])'],
    ['-A1^2%', 'of:=-[.A1]^2%'],
  ];
  for (const [grid, odf] of cases) {
    it(grid, () => {
      expect(toOdfFormula(grid)).toBe(odf);
      expect(fromOdfFormula(odf)).toBe(grid);
    });
  }
  it('reads forms other apps write', () => {
    expect(fromOdfFormula('of:=SUM([.A1:.A10])')).toBe('SUM(A1:A10)');
    expect(fromOdfFormula('of:=[$Sheet2.A1]&[.B$3]')).toBe('Sheet2!A1&B$3');
    expect(fromOdfFormula('of:=ORG.OPENOFFICE.WEEKS(1;2;0)')).toBe('WEEKS(1,2,0)');
    expect(fromOdfFormula('msoxl:=SUM(A1:A3)')).toBe('SUM(A1:A3)');
    expect(fromOdfFormula('of:=[.#REF!]+1')).toBe('#REF!+1');
    expect(fromOdfFormula("of:=['file:///x.ods'#$Sheet1.A1]")).toBe('#REF!');
    expect(fromOdfFormula('oooc:=SUM([.A1]; 2)')).toBe('SUM(A1, 2)');
  });
});

describe('number styles', () => {
  const roundTrip = (code: string) => {
    const doc = parseOdf(`<r xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">${dataStyleXml('N1', code).replace(/ clerdo:format-code="[^"]*"/g, '')}</r>`);
    const map = new Map(elems(findDeep(doc, 'r')).map((n) => [a(n, 'style:name')!, n]));
    return readDataStyle(map.get('N1')!, map);
  };
  for (const [code, want] of [
    ['0.00', '0.00'], ['#,##0', '#,##0'], ['0%', '0%'], ['0.00E+00', '0.00E+00'], ['# ?/?', '# ?/?'], ['@', '@'], ['yyyy-mm-dd', 'yyyy-mm-dd'],
    ['dddd, mmmm d, yyyy', 'dddd, mmmm d, yyyy'], ['h:mm:ss AM/PM', 'h:mm:ss AM/PM'], ['[h]:mm:ss', '[h]:mm:ss'], ['#,##0;[Red]-#,##0', '#,##0;[Red]-#,##0'],
    ['0.00;(0.00);"zero"', '0.00;(0.00);"zero"'], ['#,##0,', '#,##0,'], ['mm:ss.00', 'mm:ss.00'],
  ] as const) {
    it(`${code} without the stored code`, () => expect(roundTrip(code)).toBe(want));
  }
});

const soffice = ['/usr/bin/soffice', '/usr/local/bin/soffice'].find((p) => existsSync(p));
describe.skipIf(!soffice)('.ods opens in another app', () => {
  it('converts a written file to CSV with the same values', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ods-'));
    const wb = sample();
    wb.sheets = [wb.sheets[0]!];
    wb.names = [{ name: 'myName', formula: 'Data!$A$1:$A$3' }];
    writeFileSync(join(dir, 'a.ods'), writeOds(wb));
    execFileSync(soffice!, ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'a.ods')], { stdio: 'ignore', timeout: 120_000 });
    const rows = readFileSync(join(dir, 'a.csv'), 'utf8').split(/\r?\n/).map((l) => l.split(','));
    expect(rows[0]![0]).toBe('0.5');
    expect(rows[3]![2]).toBe('14.5');
    expect(rows[0]![6]).toBe('185');
    expect(rows[3]![6]).toBe('err');
    expect(rows[4]![6]).toBe('3');
    expect(rows[6]![6]).toBe('TRUE');
  }, 180_000);

  it('reads a file the other app wrote', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ods-'));
    writeFileSync(join(dir, 'b.csv'), 'name,amount\napple,1.5\npear,=B2*2\n');
    execFileSync(soffice!, ['--headless', '--convert-to', 'ods', '--outdir', dir, join(dir, 'b.csv')], { stdio: 'ignore', timeout: 120_000 });
    const s = readOds(new Uint8Array(readFileSync(join(dir, 'b.ods')))).sheets[0]!;
    expect(s.cells.get(cellKey(1, 0))!.v).toBe('apple');
    expect(s.cells.get(cellKey(1, 1))!.v).toBe(1.5);
    expect(s.cells.get(cellKey(2, 1))).toMatchObject({ f: 'B2*2', v: 3 });
  }, 180_000);
});

describe('.ods encryption', () => {
  it('derives keys like the standard PBKDF2', async () => {
    const { pbkdf2Sha1 } = await import('../src/formats/ods-crypto');
    const { pbkdf2Sync } = await import('node:crypto');
    const pw = new TextEncoder().encode('password'), salt = new TextEncoder().encode('salt');
    expect(Buffer.from(pbkdf2Sha1(pw, salt, 4096, 32)).toString('hex')).toBe(pbkdf2Sync(pw, salt, 4096, 32, 'sha1').toString('hex'));
  });
  it('saves and reopens with a password', async () => {
    const { PasswordRequired } = await import('../src/formats/biff8-read');
    const bytes = writeOds(sample(), { password: 'pässword' });
    const files = unzipSync(bytes);
    expect(strFromU8(files['META-INF/manifest.xml']!)).toContain('aes256-cbc');
    expect(strFromU8(files['content.xml']!)).not.toContain('table:table');
    expect(() => readOds(bytes)).toThrow(PasswordRequired);
    expect(() => readOds(bytes, { password: 'nope' })).toThrow(/not right/);
    const back = readOds(bytes, { password: 'pässword' });
    expect(back.sheets[0]!.cells.get(cellKey(0, 6))!.f).toBe('SUM(A1:A10)');
    expect(back.sheets[0]!.pictures[0]!.data).toEqual(PNG);
    expect(back.sheets[0]!.charts).toHaveLength(1);
  });
});

describe('.ods older and newer encryption', () => {
  it('AES-GCM matches the standard', async () => {
    const { aesGcm } = await import('../src/formats/ods-crypto');
    const { createCipheriv } = await import('node:crypto');
    const key = new Uint8Array(32).map((_, i) => i), iv = new Uint8Array(12).fill(9), msg = new TextEncoder().encode('x'.repeat(100) + 'end');
    const c = createCipheriv('aes-256-gcm', key, iv);
    const enc = Buffer.concat([c.update(msg), c.final(), c.getAuthTag()]);
    expect(Buffer.from(aesGcm(key, iv, msg, true)!).equals(enc)).toBe(true);
    expect(aesGcm(key, iv, new Uint8Array(enc))).toEqual(msg);
    enc[3]! ^= 1;
    expect(aesGcm(key, iv, new Uint8Array(enc))).toBeNull();
  });
  it('Argon2id and BLAKE2b match RFC 9106 and RFC 7693', async () => {
    const { argon2id, blake2b } = await import('../src/formats/argon2');
    expect(Buffer.from(blake2b(new TextEncoder().encode('abc'))).toString('hex').slice(0, 16)).toBe('ba80a53f981c4d0d');
    const r = argon2id(new Uint8Array(32).fill(1), new Uint8Array(16).fill(2), { iterations: 3, memoryKiB: 32, lanes: 4, length: 32, secret: new Uint8Array(8).fill(3), ad: new Uint8Array(12).fill(4) });
    expect(Buffer.from(r).toString('hex')).toBe('0d640df58d78766c08c037a34a8b53c9d01ef0452d75b65eb52520e96b01e659');
  });
  it('Blowfish matches the published test vector', async () => {
    const { Blowfish } = await import('../src/formats/blowfish');
    const lr = [0, 0];
    new Blowfish(new Uint8Array(8)).block(lr);
    expect(lr.map((x) => x.toString(16).padStart(8, '0')).join('')).toBe('4ef997456198dd78');
  });
  it('opens files encrypted with Argon2id and AES-GCM, per stream or as one package', async () => {
    const { zipSync, strToU8, deflateSync } = await import('fflate');
    const { aesGcm } = await import('../src/formats/ods-crypto');
    const { argon2id } = await import('../src/formats/argon2');
    const { sha256, toBase64 } = await import('../src/formats/crypto');
    const plainFiles = unzipSync(writeOds(sample()));
    const salt = new Uint8Array(16).fill(5);
    const key = argon2id(sha256(new TextEncoder().encode('pw')), salt, { iterations: 1, memoryKiB: 64, lanes: 1, length: 32 });
    const enc = (b: Uint8Array, n: number) => { const iv = new Uint8Array(12).fill(n); return { data: aesGcm(key, iv, deflateSync(b), true)!, iv }; };
    const entry = (path: string, size: number, iv: Uint8Array) => `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml" manifest:size="${size}"><manifest:encryption-data>`
      + `<manifest:algorithm manifest:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" manifest:initialisation-vector="${toBase64(iv)}"/>`
      + '<manifest:start-key-generation manifest:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" manifest:key-size="32"/>'
      + `<manifest:key-derivation manifest:key-derivation-name="urn:org:documentfoundation:names:experimental:office:manifest:argon2id" manifest:key-size="32" manifest:salt="${toBase64(salt)}" loext:argon2-iterations="1" loext:argon2-memory="64" loext:argon2-lanes="1"/>`
      + '</manifest:encryption-data></manifest:file-entry>';
    const head = '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" manifest:version="1.4"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>';
    // whole package
    const inner = zipSync(Object.fromEntries(Object.entries(plainFiles).filter(([k]) => k !== 'mimetype')));
    const e = enc(inner, 1);
    const pkg = zipSync({ mimetype: [strToU8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }], 'META-INF/manifest.xml': strToU8(head + entry('encrypted-package', inner.length, e.iv) + '</manifest:manifest>'), 'encrypted-package': [e.data, { level: 0 }] });
    expect(readOds(pkg, { password: 'pw' }).sheets[0]!.cells.get(cellKey(0, 6))!.f).toBe('SUM(A1:A10)');
    expect(() => readOds(pkg, { password: 'no' })).toThrow(/not right/);
    // per stream
    const c = enc(plainFiles['content.xml']!, 2);
    const pkg2 = zipSync({ mimetype: [strToU8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }], 'META-INF/manifest.xml': strToU8(head + entry('content.xml', plainFiles['content.xml']!.length, c.iv) + '</manifest:manifest>'), 'content.xml': [c.data, { level: 0 }], 'styles.xml': plainFiles['styles.xml']! });
    expect(readOds(pkg2, { password: 'pw' }).sheets[0]!.name).toBe('Data');
  });
});

describe('.ods first-edition encryption', () => {
  it('opens Blowfish CFB streams', async () => {
    const { zipSync, strToU8, deflateSync } = await import('fflate');
    const { Blowfish } = await import('../src/formats/blowfish');
    const { pbkdf2Sha1 } = await import('../src/formats/ods-crypto');
    const { sha1, toBase64 } = await import('../src/formats/crypto');
    const files = unzipSync(writeOds(sample()));
    const salt = new Uint8Array(16).fill(4), iv = new Uint8Array(8).fill(6);
    const key = pbkdf2Sha1(sha1(new TextEncoder().encode('old')), salt, 1024, 16);
    const packed = deflateSync(files['content.xml']!);
    const bf = new Blowfish(key);
    const out = new Uint8Array(packed.length);
    const reg = iv.slice();
    for (let o = 0; o < packed.length; o += 8) {
      const lr = [((reg[0]! << 24) | (reg[1]! << 16) | (reg[2]! << 8) | reg[3]!) >>> 0, ((reg[4]! << 24) | (reg[5]! << 16) | (reg[6]! << 8) | reg[7]!) >>> 0];
      bf.block(lr);
      const ks = [lr[0]! >>> 24, (lr[0]! >>> 16) & 255, (lr[0]! >>> 8) & 255, lr[0]! & 255, lr[1]! >>> 24, (lr[1]! >>> 16) & 255, (lr[1]! >>> 8) & 255, lr[1]! & 255];
      for (let i = 0; i < 8 && o + i < packed.length; i++) { out[o + i] = packed[o + i]! ^ ks[i]!; reg[i] = out[o + i]!; }
    }
    const manifest = '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
      + `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:size="${files['content.xml']!.length}"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="${toBase64(sha1(packed.subarray(0, 1024)))}">`
      + `<manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="${toBase64(iv)}"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:iteration-count="1024" manifest:salt="${toBase64(salt)}"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>`;
    const pkg = zipSync({ mimetype: [strToU8('application/vnd.oasis.opendocument.spreadsheet'), { level: 0 }], 'META-INF/manifest.xml': strToU8(manifest), 'content.xml': [out, { level: 0 }], 'styles.xml': files['styles.xml']! });
    expect(readOds(pkg, { password: 'old' }).sheets[0]!.cells.get(cellKey(0, 6))!.f).toBe('SUM(A1:A10)');
    expect(() => readOds(pkg, { password: 'new' })).toThrow(/not right/);
  });
  it('writes protection keys other apps accept and reads them back', async () => {
    const { protectionHash } = await import('../src/app/protection');
    const wb = sample();
    wb.sheets[0]!.protection = { sheet: true, password: protectionHash('test') };
    wb.protection = { structure: true, password: protectionHash('book') };
    const content = strFromU8(unzipSync(writeOds(wb))['content.xml']!);
    expect(content).toContain('table:protection-key-digest-algorithm="http://docs.oasis-open.org/office/ns/table/legacy-hash-excel"');
    expect(content).toContain('table:structure-protected="true"');
    expect(content).not.toContain('clerdo:protection-hash');
    const back = readOds(writeOds(wb));
    expect(back.sheets[0]!.protection?.password).toBe(protectionHash('test'));
    expect(back.protection).toEqual({ structure: true, password: protectionHash('book') });
  });
});
