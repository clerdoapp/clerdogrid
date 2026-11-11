import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { cellKey, internStyle, newSheet, newWorkbook, type Workbook } from '../src/model/workbook';
import { CellError } from '../src/engine/types';
import { writeXls, xlsLimits } from '../src/formats/biff8-write';
import { PasswordRequired, readXls } from '../src/formats/biff8-read';
import { readCfb } from '../src/formats/cfb';
import { decodeRgce, encodeFormula, type DecodeCtx, type EncodeCtx } from '../src/formats/biff8-formula';

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
  s.cells.set(cellKey(2, 6), { v: 84, f: 'Hidden!A1*2' });
  s.cells.set(cellKey(3, 6), { v: 'err', f: 'IFERROR(1/0,"err")' });
  s.cells.set(cellKey(4, 6), { v: 45092, f: 'EDATE(45000,3)' });
  s.cells.set(cellKey(5, 6), { v: 1, f: 'myName*2' });
  s.cells.set(cellKey(6, 6), { v: true, f: 'AND(A1>0,NOT(B1=1))' });
  s.cells.set(cellKey(7, 6), { v: 3, f: 'MAX(SUM(A1:B2),{1,2;3,4})' });
  s.cells.set(cellKey(8, 6), { v: 6, f: 'SUMPRODUCT(A1:A3,B1:B3)/(1+2)^2' });
  s.cells.set(cellKey(9, 6), { v: new CellError('#N/A'), f: 'VLOOKUP("x",A1:D10,2,FALSE)' });
  s.cells.set(cellKey(10, 6), { v: 2, f: 'XLOOKUP(4.5,A1:A10,B1:B10)' });
  s.cells.set(cellKey(0, 8), { v: 2, f: 'A1:A2*2', arrayRange: { r1: 0, c1: 8, r2: 1, c2: 8 } });
  s.cells.set(cellKey(1, 8), { v: 9 });
  wb.names.push({ name: 'myName', formula: 'Data!$A$1:$A$3' }, { name: 'local', formula: 'Data!$B$2', sheet: 1, hidden: true });
  s.merges.push({ r1: 12, c1: 0, r2: 13, c2: 2 });
  s.view.frozenRows = 1;
  s.view.frozenCols = 2;
  s.view.zoom = 125;
  s.view.showGrid = false;
  s.notes.set(cellKey(0, 0), { author: 'Ann', text: 'A note\nline 2' });
  s.hyperlinks.set(cellKey(1, 0), 'https://example.com/x');
  s.hyperlinks.set(cellKey(2, 0), "#'Other Sheet'!A1");
  s.hyperlinks.set(cellKey(3, 0), 'mailto:someone@example.com');
  s.cols.set(0, { width: 20 });
  s.cols.set(3, { hidden: true, width: 5 });
  s.rows.set(4, { height: 30 });
  s.rows.set(5, { hidden: true });
  const st = internStyle(wb, {
    font: { name: 'Arial', size: 14, bold: true, italic: true, color: '#123456', underline: 'single' }, fill: { color: '#FFFF00' },
    border: { left: { style: 'thin' }, bottom: { style: 'double', color: '#0000FF' } }, align: { h: 'center', v: 'top', wrap: true, rotation: 45 }, numFmt: '0.000',
    protection: { locked: false },
  });
  s.cells.get(cellKey(0, 1))!.s = st;
  s.cells.set(cellKey(20, 0), { v: 'bold', rich: [{ text: 'bo', font: { bold: true } }, { text: 'ld', font: { italic: true, color: '#00FF00' } }] });
  s.cells.set(cellKey(21, 0), { v: true });
  s.cells.set(cellKey(22, 0), { v: 45000.25, s: internStyle(wb, { numFmt: 'yyyy-mm-dd hh:mm' }) });
  s.cells.set(cellKey(23, 0), { v: 'Unicode – ünïcødé 漢字' });
  s.cells.set(cellKey(24, 0), { v: new CellError('#DIV/0!') });
  s.cells.set(cellKey(25, 0), { v: -1234567.891 });
  s.conditionalFormats.push({ ranges: [{ r1: 0, c1: 1, r2: 9, c2: 1 }], rules: [
    { type: 'cellIs', operator: 'between', formulas: ['10', '20'], style: { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFCCCC' } } },
    { type: 'expression', formula: 'MOD(ROW(),2)=0', style: { fill: { color: '#CCFFCC' } } },
    { type: 'expression', formula: 'B1>$A$1', style: { border: { bottom: { style: 'thin' } } } },
  ] });
  s.validations.push({ ranges: [{ r1: 0, c1: 9, r2: 5, c2: 9 }], type: 'list', formula1: '"red,green,blue"', showDropdown: true, prompt: { title: 'Pick', text: 'Choose' }, error: { style: 'warning', title: 'Bad', text: 'No' } });
  s.validations.push({ ranges: [{ r1: 0, c1: 10, r2: 5, c2: 10 }], type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true });
  s.validations.push({ ranges: [{ r1: 0, c1: 11, r2: 5, c2: 11 }], type: 'list', formula1: '$A$1:$A$5' });
  s.autoFilter = { range: { r1: 0, c1: 0, r2: 9, c2: 3 }, columns: [{ col: 1, custom: { and: true, rules: [{ operator: 'greaterThan', value: '5' }] } }] };
  s.pageSetup = { orientation: 'landscape', paper: 9, printArea: 'A1:D10', printTitleRows: '$1:$1', margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 }, header: '&CTitle', fitWidth: 1, fitHeight: 0, rowBreaks: [20] };
  s.protection = { sheet: true, password: 'CC3D' };
  s.pictures.push({ id: 'p1', data: PNG, mime: 'image/png', anchor: { row: 15, col: 4, dx: 3, dy: 2, width: 120, height: 80 }, alt: 'logo' });
  wb.protection = { structure: true };
  wb.activeSheet = 0;
  return wb;
};

describe('.xls save and reopen', () => {
  const wb = sample();
  const back = readXls(writeXls(wb));
  const s = back.sheets[0]!;

  it('keeps sheets, states and tab colours', () => {
    expect(back.sheets.map((x) => x.name)).toEqual(['Data', 'Other Sheet', 'Hidden']);
    expect(back.sheets[2]!.state).toBe('hidden');
    expect(back.sheets[2]!.tabColor).toBe('#FF0000');
    expect(back.protection?.structure).toBe(true);
  });

  it('keeps values of every type', () => {
    expect(s.cells.get(cellKey(3, 2))!.v).toBe(14.5);
    expect(s.cells.get(cellKey(21, 0))!.v).toBe(true);
    expect(s.cells.get(cellKey(23, 0))!.v).toBe('Unicode – ünïcødé 漢字');
    expect((s.cells.get(cellKey(24, 0))!.v as CellError).code).toBe('#DIV/0!');
    expect(s.cells.get(cellKey(25, 0))!.v).toBe(-1234567.891);
    expect(s.cells.get(cellKey(22, 0))!.v).toBe(45000.25);
    expect(back.styles[s.cells.get(cellKey(22, 0))!.s!]!.numFmt).toBe('yyyy-mm-dd hh:mm');
  });

  it('keeps formulas, cached values and array formulas', () => {
    const f = (r: number, c = 6) => s.cells.get(cellKey(r, c))!;
    expect(f(0).f).toBe('SUM(A1:A10)');
    expect(f(0).v).toBe(185);
    expect(f(1).f).toBe("'Other Sheet'!A1&\"x\"");
    expect(f(1).v).toBe('ax');
    expect(f(2).f).toBe('Hidden!A1*2');
    expect(f(3).f).toBe('IFERROR(1/0,"err")');
    expect(f(4).f).toBe('EDATE(45000,3)');
    expect(f(5).f).toBe('myName*2');
    expect(f(6).f).toBe('AND(A1>0,NOT(B1=1))');
    expect(f(6).v).toBe(true);
    expect(f(7).f).toBe('MAX(SUM(A1:B2),{1,2;3,4})');
    expect(f(8).f).toBe('SUMPRODUCT(A1:A3,B1:B3)/(1+2)^2');
    expect((f(9).v as CellError).code).toBe('#N/A');
    expect(f(10).f).toBe('XLOOKUP(4.5,A1:A10,B1:B10)');
    expect(f(0, 8).f).toBe('A1:A2*2');
    expect(f(0, 8).arrayRange).toEqual({ r1: 0, c1: 8, r2: 1, c2: 8 });
    expect(f(1, 8).f).toBeUndefined();
    expect(f(1, 8).v).toBe(9);
  });

  it('keeps styles and rich text', () => {
    const st = back.styles[s.cells.get(cellKey(0, 1))!.s!]!;
    expect(st.font).toMatchObject({ name: 'Arial', size: 14, bold: true, italic: true, underline: 'single', color: '#123456' });
    expect(st.fill).toEqual({ color: '#FFFF00' });
    expect(st.border?.left).toEqual({ style: 'thin' });
    expect(st.border?.bottom).toEqual({ style: 'double', color: '#0000FF' });
    expect(st.align).toEqual({ h: 'center', v: 'top', wrap: true, rotation: 45 });
    expect(st.numFmt).toBe('0.000');
    expect(st.protection).toEqual({ locked: false });
    expect(s.cells.get(cellKey(20, 0))!.rich).toEqual([{ text: 'bo', font: { bold: true } }, { text: 'ld', font: { italic: true, color: '#00FF00' } }]);
  });

  it('keeps layout, view and names', () => {
    expect(s.merges).toEqual([{ r1: 12, r2: 13, c1: 0, c2: 2 }]);
    expect(s.cols.get(0)!.width).toBeCloseTo(20, 1);
    expect(s.cols.get(3)!.hidden).toBe(true);
    expect(s.rows.get(4)!.height).toBe(30);
    expect(s.rows.get(5)!.hidden).toBe(true);
    expect(s.view).toMatchObject({ frozenRows: 1, frozenCols: 2, zoom: 125, showGrid: false });
    expect(back.names).toEqual([{ name: 'myName', formula: 'Data!$A$1:$A$3' }, { name: 'local', formula: 'Data!$B$2', sheet: 1, hidden: true }]);
  });

  it('keeps notes, links, rules, validation, filter, protection and page setup', () => {
    expect(s.notes.get(0)).toEqual({ author: 'Ann', text: 'A note\nline 2' });
    expect(s.hyperlinks.get(cellKey(1, 0))).toBe('https://example.com/x');
    expect(s.hyperlinks.get(cellKey(2, 0))).toBe("#'Other Sheet'!A1");
    expect(s.hyperlinks.get(cellKey(3, 0))).toBe('mailto:someone@example.com');
    const rules = s.conditionalFormats[0]!.rules;
    expect(rules[0]).toEqual({ type: 'cellIs', operator: 'between', formulas: ['10', '20'], style: { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFCCCC' } } });
    expect(rules[1]).toEqual({ type: 'expression', formula: 'MOD(ROW(),2)=0', style: { fill: { color: '#CCFFCC' } } });
    expect(rules[2]).toMatchObject({ type: 'expression', formula: 'B1>$A$1', style: { border: { bottom: { style: 'thin' } } } });
    expect(s.validations[0]).toEqual({ ranges: [{ r1: 0, r2: 5, c1: 9, c2: 9 }], type: 'list', formula1: '"red,green,blue"', showDropdown: true, prompt: { title: 'Pick', text: 'Choose' }, error: { style: 'warning', title: 'Bad', text: 'No' } });
    expect(s.validations[1]).toMatchObject({ type: 'whole', operator: 'between', formula1: '1', formula2: '10', allowBlank: true });
    expect(s.validations[2]).toMatchObject({ type: 'list', formula1: '$A$1:$A$5' });
    expect(s.autoFilter?.range).toEqual({ r1: 0, r2: 9, c1: 0, c2: 3 });
    expect(s.autoFilter?.columns[0]).toEqual({ col: 1, custom: { and: true, rules: [{ operator: 'greaterThan', value: '5' }] } });
    expect(s.protection).toMatchObject({ sheet: true, password: 'CC3D' });
    expect(s.pageSetup).toMatchObject({ orientation: 'landscape', paper: 9, printArea: '$A$1:$D$10', printTitleRows: '$1:$1', header: '&CTitle', fitWidth: 1, fitHeight: 0, rowBreaks: [20] });
    expect(s.pageSetup?.margins).toEqual({ left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 });
  });

  it('keeps pictures', () => {
    expect(s.pictures).toHaveLength(1);
    expect(s.pictures[0]!.mime).toBe('image/png');
    expect(s.pictures[0]!.data).toEqual(PNG);
    expect(s.pictures[0]!.anchor.row).toBe(15);
    expect(s.pictures[0]!.anchor.col).toBe(4);
    expect(Math.abs(s.pictures[0]!.anchor.width - 120)).toBeLessThan(3);
  });

  it('splits long string tables across continuation records', () => {
    const big = newWorkbook();
    const sh = big.sheets[0]!;
    const long = 'x'.repeat(20000) + '漢'.repeat(5000);
    sh.cells.set(0, { v: long });
    for (let i = 1; i < 3000; i++) sh.cells.set(cellKey(i, 0), { v: `string ${i} ${i % 2 ? 'é' : 'ü漢'}` });
    const r = readXls(writeXls(big)).sheets[0]!;
    expect(r.cells.get(0)!.v).toBe(long);
    expect(r.cells.get(cellKey(2999, 0))!.v).toBe('string 2999 é');
    expect(r.cells.get(cellKey(1500, 0))!.v).toBe('string 1500 ü漢');
  });

  it('writes a 1904 date system and limits', () => {
    const w = newWorkbook();
    w.date1904 = true;
    w.sheets[0]!.cells.set(cellKey(70000, 0), { v: 1 });
    expect(readXls(writeXls(w)).date1904).toBe(true);
    expect(xlsLimits(w)).toContain('Rows after 65,536 are left out.');
  });
});

describe('.xls encryption', () => {
  it('opens files encrypted with the default password without asking', () => {
    const back = readXls(writeXls(sample(), { password: '' }));
    expect(back.sheets[0]!.cells.get(cellKey(0, 6))!.f).toBe('SUM(A1:A10)');
  });
  it('needs the right password', () => {
    const bytes = writeXls(sample(), { password: 'secret' });
    expect(() => readXls(bytes)).toThrow(PasswordRequired);
    expect(() => readXls(bytes, { password: 'nope' })).toThrow(/not right/);
    const back = readXls(bytes, { password: 'secret' });
    expect(back.sheets[0]!.notes.get(0)!.text).toBe('A note\nline 2');
    expect(back.sheets[0]!.cells.get(cellKey(23, 0))!.v).toBe('Unicode – ünïcødé 漢字');
  });
});

describe('binary formula tokens', () => {
  const sheets = ['Data', 'Two Words'];
  const names = ['Rate'];
  const enc: EncodeCtx = { mode: 'cell', sheetIndex: (n) => { const i = sheets.findIndex((s) => s.toUpperCase() === n.toUpperCase()); return i < 0 ? undefined : i; }, ixti: (a, b) => a * 10 + b, nameIndex: (n) => { const i = names.indexOf(n); return i < 0 ? undefined : i + 1; }, addin: () => 99 };
  const dec: DecodeCtx = { xlsb: false, row: 0, col: 0, sheets: (i) => ({ first: sheets[Math.floor(i / 10)]!, last: sheets[i % 10]! }), name: (i) => (i === 99 ? { name: '_xlfn.TEXTJOIN' } : { name: names[i - 1]! }), externName: () => undefined };
  for (const f of ['1+2*3', '(1+2)*3', '-A1^2', '-(A1^2)', 'A1&"x"', 'SUM(A1:B5,C3)', "'Two Words'!$A$1+Data!B$2", "'Data:Two Words'!A1", 'IF(A1>1,"a",IF(A2,1,2))',
    'Rate*2', '{1,2;3,-4.5}', 'SUM(A:A)', 'SUM(3:4)', 'A1:B2 B2:C3', 'SUM((A1,B2))', '50%', 'TEXTJOIN(",",TRUE,A1:A3)', 'ROUND(PI(),2)', 'NOW()', '"a""b"', 'TRUE', '#N/A']) {
    it(f, () => {
      const { rgce, rgcb } = encodeFormula(f, enc);
      expect(decodeRgce(rgce, rgcb, dec)).toBe(f.replace("'Data:Two Words'!A1", "'Data:Two Words'!A1"));
    });
  }
  it('decodes relative tokens from the formula cell', () => {
    const { rgce, rgcb } = encodeFormula('B2+$C$3', { ...enc, mode: 'rel', row: 1, col: 1 });
    expect(decodeRgce(rgce, rgcb, { ...dec, row: 5, col: 5 })).toBe('F6+$C$3');
  });
});

const soffice = ['/usr/bin/soffice', '/usr/local/bin/soffice'].find((p) => existsSync(p));
describe.skipIf(!soffice)('.xls opens in another app', () => {
  it('converts a written file to CSV with the same values', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xls-'));
    const wb = sample();
    wb.sheets = [wb.sheets[0]!];
    wb.names = [{ name: 'myName', formula: 'Data!$A$1:$A$3' }];
    writeFileSync(join(dir, 'a.xls'), writeXls(wb));
    execFileSync(soffice!, ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'a.xls')], { stdio: 'ignore', timeout: 120_000 });
    const rows = readFileSync(join(dir, 'a.csv'), 'utf8').split(/\r?\n/).map((l) => l.split(','));
    expect(rows[0]![0]).toBe('0.5');
    expect(rows[3]![2]).toBe('14.5');
    expect(rows[0]![6]).toBe('185');
    expect(rows[3]![6]).toBe('err');
    expect(rows[4]![6]).toBe('45092');
    expect(rows[23]![0]).toBe('Unicode – ünïcødé 漢字');
  }, 180_000);

  it('reads a file the other app wrote', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xls-'));
    writeFileSync(join(dir, 'b.csv'), 'name,amount\napple,1.5\npear,=B2*2\n');
    execFileSync(soffice!, ['--headless', '--convert-to', 'xls', '--outdir', dir, join(dir, 'b.csv')], { stdio: 'ignore', timeout: 120_000 });
    const s = readXls(new Uint8Array(readFileSync(join(dir, 'b.xls')))).sheets[0]!;
    expect(s.cells.get(cellKey(1, 0))!.v).toBe('apple');
    expect(s.cells.get(cellKey(1, 1))!.v).toBe(1.5);
    expect(readCfb(new Uint8Array(readFileSync(join(dir, 'b.xls')))).has('Workbook')).toBe(true);
  }, 180_000);
});
