import { describe, expect, test } from 'vitest';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { cellKey, internStyle, newSheet, newWorkbook, type Workbook } from '../src/model/workbook';
import { CellError } from '../src/engine/types';
import { readXlsx, writeXlsx, shiftFormula, addFuturePrefixes, stripFuturePrefixes } from '../src/formats/xlsx';
import { parseXml, type XNode } from '../src/formats/xml';
import { applyTint } from '../src/formats/xlsx-styles';

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 15, 4, 0, 9, 251, 3, 253, 227, 85, 242, 156, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130]);

const build = (): Workbook => {
  const wb = newWorkbook();
  wb.props = { title: 'Report', creator: 'Tester', created: '2024-01-02T03:04:05Z', modified: '2024-01-02T03:04:05Z' };
  wb.calc = { iterate: true, maxIterations: 50, maxChange: 0.01 };
  const bold = internStyle(wb, { font: { bold: true, color: '#FF0000' }, fill: { color: '#FFFF00' }, border: { bottom: { style: 'thin', color: '#000000' } } });
  const money = internStyle(wb, { numFmt: '"$"#,##0.00', align: { h: 'right', wrap: true, indent: 1 } });
  const date = internStyle(wb, { numFmt: 'm/d/yyyy' });
  const custom = internStyle(wb, { numFmt: '0.000" kg"', protection: { locked: false } });
  const custom2 = internStyle(wb, { numFmt: '0.000" kg"', font: { italic: true, name: 'Arial', size: 14 } });
  const grad = internStyle(wb, { fill: { gradient: { angle: 90, stops: [{ pos: 0, color: '#FFFFFF' }, { pos: 1, color: '#4472C4' }] } }, align: { rotation: -45, v: 'top' } });
  const pat = internStyle(wb, { fill: { pattern: 'darkGrid', color: '#00FF00', patternColor: '#0000FF' }, font: { underline: 'double', strike: true } });
  const s = wb.sheets[0]!;
  s.name = 'Data & <Stuff>';
  const set = (r: number, c: number, cell: Parameters<typeof s.cells.set>[1]): void => { s.cells.set(cellKey(r, c), cell); };
  set(0, 0, { v: 'Name', s: bold });
  set(0, 1, { v: 'Amount', s: bold });
  set(1, 0, { v: '  padded\ttext\nline2 ' });
  set(1, 1, { v: 1234.5, s: money });
  set(2, 0, { v: 'Ctrl\u0001char & <tag> "q"' });
  set(2, 1, { v: -0.000123 });
  set(3, 0, { v: true });
  set(3, 1, { v: new CellError('#DIV/0!'), f: '1/0' });
  set(4, 0, { v: 45366, s: date });
  set(4, 1, { v: 2, f: 'SUM(B2:B3)', s: custom });
  set(5, 0, { v: 'Rich', rich: [{ text: 'Ri', font: { bold: true } }, { text: 'ch', font: { color: '#00B050', italic: true } }] });
  set(5, 1, { v: 3, f: 'XLOOKUP(A2,A1:A5,B1:B5)', s: custom2 });
  set(6, 0, { v: 1, f: 'SEQUENCE(3)', dyn: true });
  set(6, 1, { v: 10, f: 'B2:B4*2', arrayRange: { r1: 6, c1: 1, r2: 8, c2: 1 } });
  set(7, 2, { v: 'x', f: 'TEXTJOIN(",",TRUE,A1:A3)&"_xlfn.LITERAL"' });
  set(8, 2, { v: 6, f: 'LET(x,2,y,3,x*y)', s: grad });
  set(9, 2, { v: null, s: pat });
  set(10, 0, { v: 7, f: 'SUM(A7#)' });
  set(11, 0, { v: '' });
  s.cols.set(0, { width: 20 });
  s.cols.set(1, { width: 12.5, hidden: true });
  s.cols.set(2, { width: 12.5, hidden: true });
  s.cols.set(5, { s: money });
  s.rows.set(0, { height: 30 });
  s.rows.set(20, { hidden: true });
  s.merges.push({ r1: 12, c1: 0, r2: 13, c2: 2 });
  s.view = { frozenRows: 1, frozenCols: 1, showGrid: false, showHeaders: true, zoom: 125, rtl: true, selection: { r1: 2, c1: 1, r2: 4, c2: 2 } };
  s.tabColor = '#FF8800';
  s.conditionalFormats = [
    { ranges: [{ r1: 1, c1: 1, r2: 5, c2: 1 }], rules: [{ type: 'cellIs', operator: 'greaterThan', formulas: ['100'], style: { font: { color: '#9C0006' }, fill: { color: '#FFC7CE' } }, stop: true }], priority: 1 },
    { ranges: [{ r1: 1, c1: 1, r2: 5, c2: 1 }, { r1: 1, c1: 3, r2: 5, c2: 3 }], rules: [{ type: 'colorScale', stops: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'max' }, color: '#63BE7B' }] }], priority: 2 },
    { ranges: [{ r1: 1, c1: 0, r2: 9, c2: 0 }], rules: [{ type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: '#638EC6', showValue: false }], priority: 3 },
    { ranges: [{ r1: 1, c1: 0, r2: 9, c2: 0 }], rules: [{ type: 'iconSet', set: '3Arrows', values: [{ type: 'percent', value: '0' }, { type: 'percent', value: '33' }, { type: 'percent', value: '67' }], reverse: true }], priority: 4 },
    { ranges: [{ r1: 1, c1: 0, r2: 9, c2: 0 }], rules: [
      { type: 'expression', formula: 'MOD(ROW(),2)=0', style: { fill: { color: '#EEEEEE' } } },
      { type: 'top10', rank: 3, bottom: true, percent: true, style: { font: { bold: true } } },
      { type: 'aboveAverage', below: true, style: { font: { italic: true } } },
      { type: 'duplicateValues', style: { font: { strike: true } } },
      { type: 'containsText', text: 'a"b', style: { font: { color: '#0000FF' } } },
      { type: 'containsBlanks', style: { fill: { color: '#CCCCCC' } } },
      { type: 'timePeriod', period: 'last7Days', style: { font: { underline: 'single' } } },
    ], priority: 5 },
  ];
  s.validations = [
    { ranges: [{ r1: 1, c1: 4, r2: 10, c2: 4 }], type: 'list', formula1: '"Yes,No"', allowBlank: true, showDropdown: true, prompt: { title: 'Pick', text: 'Yes or no' }, error: { style: 'warning', title: 'Oops', text: 'Bad' } },
    { ranges: [{ r1: 1, c1: 5, r2: 10, c2: 5 }], type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
  ];
  s.autoFilter = { range: { r1: 0, c1: 0, r2: 5, c2: 1 }, columns: [{ col: 0, values: ['Rich', 'Name'], blanks: true }, { col: 1, custom: { and: true, rules: [{ operator: 'greaterThan', value: '1' }, { operator: 'lessThan', value: '5000' }] } }], sort: [{ col: 1, descending: true }] };
  s.hyperlinks.set(cellKey(1, 0), 'https://clerdo.app/');
  s.hyperlinks.set(cellKey(2, 0), '#Second!A1');
  s.hyperlinks.set(cellKey(3, 0), 'javascript:alert(1)');
  s.notes.set(cellKey(0, 0), { author: 'Ann', text: 'Header note' });
  s.notes.set(cellKey(4, 1), { text: 'No author <b>' });
  s.protection = { sheet: true, password: 'CC3D', options: { formatCells: false, sort: false } };
  s.pageSetup = { orientation: 'landscape', paper: 9, fitWidth: 1, fitHeight: 0, printArea: '$A$1:$C$10', margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.4, footer: 0.4 } };
  s.charts.push({ id: 'chart1', type: 'column', title: 'Amounts', series: [{ name: 'Amount', values: "'Data & <Stuff>'!$B$2:$B$5", categories: "'Data & <Stuff>'!$A$2:$A$5", color: '#4472C4' }], anchor: { row: 2, col: 4, dx: 0, dy: 0, width: 480, height: 300 }, legend: 'bottom', axisTitles: { x: 'Who', y: 'How much' }, dataLabels: true });
  s.charts.push({ id: 'chart2', type: 'pie', series: [{ values: "'Data & <Stuff>'!$B$2:$B$5" }], anchor: { row: 20, col: 1, dx: 10, dy: 5, width: 300, height: 200 }, legend: 'right' });
  s.charts.push({ id: 'chart3', type: 'combo', title: 'Combo', series: [{ name: 'A', values: "'Data & <Stuff>'!$B$2:$B$5", type: 'column' }, { name: 'B', values: "'Data & <Stuff>'!$B$2:$B$5", type: 'line' }], anchor: { row: 30, col: 1, dx: 0, dy: 0, width: 400, height: 250 }, legend: 'none' });
  s.pictures.push({ id: 'picture1', data: PNG, mime: 'image/png', anchor: { row: 40, col: 0, dx: 3, dy: 4, width: 64, height: 32 } });

  const t = newSheet('Second');
  t.cells.set(cellKey(0, 0), { v: 'Col1' });
  t.cells.set(cellKey(0, 1), { v: 'Col2' });
  t.cells.set(cellKey(1, 0), { v: 1 });
  t.cells.set(cellKey(1, 1), { v: 2, f: 'Table1[@Col1]*2' });
  t.cells.set(cellKey(2, 0), { v: 1, f: 'SUBTOTAL(109,Table1[Col1])' });
  t.tables.push({ name: 'Table1', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, headerRow: true, totalsRow: true, columns: [{ name: 'Col1', totalsFunction: 'sum' }, { name: 'Col2', calculatedFormula: 'Table1[@Col1]*2' }], style: 'TableStyleLight9', bandedRows: true, bandedCols: false });
  t.defaultColWidth = 10;
  t.defaultRowHeight = 18;
  wb.sheets.push(t);
  const h = newSheet('Hidden');
  h.state = 'veryHidden';
  wb.sheets.push(h);
  wb.activeSheet = 1;
  wb.names.push({ name: 'Rate', formula: '0.05', comment: 'Tax' }, { name: 'LocalRange', formula: 'Second!$A$1:$A$3', sheet: 1 }, { name: 'Secret', formula: '42', hidden: true });
  wb.extraParts = new Map([['xl/vbaProject.bin', new Uint8Array([1, 2, 3, 4])]]);
  return wb;
};

const parts = (data: Uint8Array): Record<string, string> => {
  const files = unzipSync(data);
  const out: Record<string, string> = {};
  for (const k of Object.keys(files)) if (/\.(xml|rels|vml)$/.test(k)) out[k] = strFromU8(files[k]!);
  return out;
};

/** Strict well-formedness check: balanced tags, quoted attributes, legal characters and entities. */
const wellFormed = (xml: string): void => {
  // eslint-disable-next-line no-control-regex
  expect(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(xml)).toBe(false);
  expect(/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml)).toBe(false);
  const stack: string[] = [];
  const re = /<(\/?)([A-Za-z_][\w.:-]*)((?:\s+[\w.:-]+\s*=\s*(?:"[^"<]*"|'[^'<]*'))*)\s*(\/?)>|<\?[^?]*\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|(<)/g;
  let m: RegExpExecArray | null;
  let roots = 0;
  while ((m = re.exec(xml))) {
    if (m[5]) throw new Error(`stray < at ${m.index}: ${xml.slice(m.index, m.index + 60)}`);
    if (!m[2]) continue;
    if (m[1]) { expect(stack.pop()).toBe(m[2]); }
    else if (!m[4]) { if (!stack.length) roots++; stack.push(m[2]); }
    else if (!stack.length) roots++;
    const names = [...(m[3] ?? '').matchAll(/([\w.:-]+)\s*=/g)].map((a) => a[1]);
    expect(new Set(names).size).toBe(names.length);
  }
  expect(stack).toEqual([]);
  expect(roots).toBe(1);
};

describe('xlsx round trip', () => {
  const wb = build();
  const data = writeXlsx(wb);
  const back = readXlsx(data);

  test('every part is well-formed XML', () => {
    const p = parts(data);
    expect(Object.keys(p)).toEqual(expect.arrayContaining(['[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml', 'xl/worksheets/sheet1.xml', 'xl/metadata.xml', 'xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/tables/table1.xml', 'xl/charts/chart1.xml', 'xl/drawings/drawing1.xml']));
    for (const [name, xml] of Object.entries(p)) {
      try { wellFormed(xml); } catch (e) { throw new Error(`${name}: ${(e as Error).message}`); }
    }
  });

  test('content types cover every part', () => {
    const p = parts(data);
    const ct = p['[Content_Types].xml']!;
    for (const name of Object.keys(unzipSync(data))) {
      if (name === '[Content_Types].xml') continue;
      const ext = name.slice(name.lastIndexOf('.') + 1);
      expect(ct.includes(`PartName="/${name}"`) || ct.includes(`Extension="${ext}"`), name).toBe(true);
    }
    expect(ct).toContain('macroEnabled');
  });

  test('workbook-level data', () => {
    expect(back.sheets.map((s) => s.name)).toEqual(['Data & <Stuff>', 'Second', 'Hidden']);
    expect(back.sheets.map((s) => s.state)).toEqual(['visible', 'visible', 'veryHidden']);
    expect(back.activeSheet).toBe(1);
    expect(back.date1904).toBe(false);
    expect(back.props).toEqual(wb.props);
    expect(back.calc).toEqual(wb.calc);
    expect(back.names).toEqual(wb.names);
    expect(back.extraParts?.get('xl/vbaProject.bin')).toEqual(new Uint8Array([1, 2, 3, 4]));
  });

  test('cells, formulas and styles', () => {
    const a = wb.sheets[0]!, b = back.sheets[0]!;
    expect([...b.cells.keys()].sort()).toEqual([...a.cells.keys()].sort());
    for (const [k, c] of a.cells) {
      const d = b.cells.get(k)!;
      expect(d.v, `cell ${k}`).toEqual(c.v === '' ? '' : c.v);
      expect(d.f).toEqual(c.f);
      expect(d.dyn).toEqual(c.dyn);
      expect(d.arrayRange).toEqual(c.arrayRange);
      expect(d.rich).toEqual(c.rich);
      expect(back.styles[d.s ?? 0], `style of ${k}`).toEqual(wb.styles[c.s ?? 0]);
    }
  });

  test('sheet structure', () => {
    const a = wb.sheets[0]!, b = back.sheets[0]!;
    expect(b.cols).toEqual(a.cols);
    expect(b.rows).toEqual(a.rows);
    expect(b.merges).toEqual(a.merges);
    expect(b.view).toEqual(a.view);
    expect(b.tabColor).toBe(a.tabColor);
    expect(b.conditionalFormats).toEqual(a.conditionalFormats);
    expect(b.validations).toEqual(a.validations);
    expect(b.autoFilter).toEqual(a.autoFilter);
    expect([...b.hyperlinks]).toEqual([...a.hyperlinks].filter(([, t]) => !t.startsWith('javascript')));
    expect(b.notes).toEqual(a.notes);
    expect(b.protection).toEqual(a.protection);
    expect(b.pageSetup).toEqual(a.pageSetup);
    expect(b.charts).toEqual(a.charts);
    expect(b.pictures).toEqual(a.pictures);
    const s2 = back.sheets[1]!;
    expect(s2.tables).toEqual(wb.sheets[1]!.tables);
    expect(s2.defaultColWidth).toBe(10);
    expect(s2.defaultRowHeight).toBe(18);
    expect(s2.cells).toEqual(wb.sheets[1]!.cells);
  });

  test('write → read → write is stable', () => {
    const again = readXlsx(writeXlsx(back));
    expect(again.sheets[0]!.cells).toEqual(back.sheets[0]!.cells);
    expect(again.styles).toEqual(back.styles);
    expect(again.sheets[0]!.conditionalFormats).toEqual(back.sheets[0]!.conditionalFormats);
  });

  test('formulas are stored with future-function prefixes', () => {
    const xml = parts(data)['xl/worksheets/sheet1.xml']!;
    expect(xml).toContain('<f>_xlfn.XLOOKUP(A2,A1:A5,B1:B5)</f>');
    expect(xml).toContain('_xlfn.SEQUENCE(3)</f>');
    expect(xml).toContain('cm="1"');
    expect(xml).toContain('_xlfn.LET(_xlpm.x,2,_xlpm.y,3,_xlpm.x*_xlpm.y)');
    expect(xml).toContain('SUM(_xlfn.ANCHORARRAY(A7))');
    expect(xml).toContain('&amp;&quot;_xlfn.LITERAL&quot;');
    expect(xml).toContain('<f t="array" ref="B7:B9">B2:B4*2</f>');
  });

  test('styles are deduplicated', () => {
    const xml = parts(data)['xl/styles.xml']!;
    const doc = parseXml(xml).children[0]!;
    const count = (n: string): number => doc.children.find((c) => c.name === n)!.children.length;
    expect(count('numFmts')).toBe(2); // "$"#,##0.00 and 0.000" kg" (m/d/yyyy is built-in 14)
    expect(xml).toContain('numFmtId="164"');
    expect(xml).toContain('numFmtId="14"');
    expect(count('cellXfs')).toBe(wb.styles.length);
    expect(count('borders')).toBe(2);
    // fonts: default, bold red, italic Arial 14, double-underline strike
    expect(count('fonts')).toBe(4);
  });

  test('a workbook with many identical styled cells keeps one font and fill', () => {
    const w = newWorkbook();
    for (let i = 0; i < 50; i++) {
      const st = internStyle(w, { font: { bold: true }, fill: { color: '#112233' }, numFmt: i % 2 ? '0.0' : '0.00%' + 'x'.repeat(i % 3) });
      w.sheets[0]!.cells.set(cellKey(i, 0), { v: i, s: st });
    }
    const doc = parseXml(parts(writeXlsx(w))['xl/styles.xml']!).children[0]!;
    const kidsOf = (n: string): XNode[] => doc.children.find((c) => c.name === n)!.children;
    expect(kidsOf('fonts').length).toBe(2);
    expect(kidsOf('fills').length).toBe(3);
    expect(kidsOf('numFmts').length).toBe(3);
  });
});

const zipParts = (files: Record<string, string>): Uint8Array => {
  const z: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) z[k] = strToU8(v);
  return zipSync(z);
};

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const handMade = (sheet: string, extra: Record<string, string> = {}, wbExtra = ''): Uint8Array => zipParts({
  '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml': `<workbook xmlns="${MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>${wbExtra}</workbook>`,
  'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/s.xml"/>'
    + (extra['xl/styles.xml'] ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' : '')
    + (extra['xl/theme/theme1.xml'] ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' : '')
    + (extra['xl/sharedStrings.xml'] ? '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : '')
    + '</Relationships>',
  'xl/worksheets/s.xml': `<worksheet xmlns="${MAIN}">${sheet}</worksheet>`,
  ...extra,
});

describe('reading hand-written parts', () => {
  test('shared formulas expand with relative refs shifted', () => {
    const wb = readXlsx(handMade(`<sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:C3" si="0">A1*$A$1+SUM(A$1:A1)+'My Sheet'!A1&amp;"A1"+Table1[A1]</f><v>3</v></c><c r="C1"><f t="shared" si="0"/><v>0</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>5</v></c></row>
      <row r="3"><c r="B3"><f t="shared" si="0"/></c><c r="C3"><f t="shared" si="0"/></c></row>
    </sheetData>`));
    const s = wb.sheets[0]!;
    const f = (r: number, c: number): string | undefined => s.cells.get(cellKey(r, c))?.f;
    expect(f(0, 1)).toBe(`A1*$A$1+SUM(A$1:A1)+'My Sheet'!A1&"A1"+Table1[A1]`);
    expect(f(0, 2)).toBe(`B1*$A$1+SUM(B$1:B1)+'My Sheet'!B1&"A1"+Table1[A1]`);
    expect(f(1, 1)).toBe(`A2*$A$1+SUM(A$1:A2)+'My Sheet'!A2&"A1"+Table1[A1]`);
    expect(f(2, 2)).toBe(`B3*$A$1+SUM(B$1:B3)+'My Sheet'!B3&"A1"+Table1[A1]`);
    expect(s.cells.get(cellKey(1, 1))?.v).toBe(5);
    expect(wb.date1904).toBe(true);
  });

  test('cell types, inline strings, ISO dates, errors, theme colours and dxfs', () => {
    const wb = readXlsx(handMade(`<sheetData><row r="1">
      <c r="A1" t="inlineStr"><is><t>inline</t></is></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="e"><v>#N/A</v></c>
      <c r="D1" t="d"><v>1904-01-02T12:00:00</v></c><c r="E1" t="str"><f>_xlfn.CONCAT("a","b")</f><v>ab</v></c><c r="F1" t="s" s="1"><v>0</v></c><c r="G1"><v>1e3</v></c>
      </row></sheetData>
      <conditionalFormatting sqref="A1:A9 C1"><cfRule type="expression" dxfId="0" priority="2"><formula>_xlfn.ISFORMULA(A1)</formula></cfRule></conditionalFormatting>`, {
      'xl/sharedStrings.xml': `<sst xmlns="${MAIN}"><si><r><rPr><b/><color theme="4"/></rPr><t>Bo</t></r><r><t xml:space="preserve">ld </t></r><rPh><t>ignored</t></rPh></si></sst>`,
      'xl/styles.xml': `<styleSheet xmlns="${MAIN}"><fonts><font><sz val="11"/><name val="Calibri"/><color theme="1"/></font><font><b/><sz val="11"/><color theme="4" tint="0.39997558519241921"/><name val="Calibri"/></font></fonts>
        <fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="10"/></patternFill></fill></fills>
        <borders><border/></borders><cellXfs><xf/><xf fontId="1" fillId="2" numFmtId="10"><alignment horizontal="center" textRotation="135"/></xf></cellXfs>
        <dxfs><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`,
      'xl/theme/theme1.xml': '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="112233"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
    }));
    const s = wb.sheets[0]!;
    const v = (c: number): unknown => s.cells.get(cellKey(0, c))?.v;
    expect(v(0)).toBe('inline');
    expect(v(1)).toBe(true);
    expect(v(2)).toEqual(new CellError('#N/A'));
    expect(v(3)).toBe(1.5);
    expect(s.cells.get(cellKey(0, 4))?.f).toBe('CONCAT("a","b")');
    expect(v(5)).toBe('Bold ');
    expect(v(6)).toBe(1000);
    expect(s.cells.get(cellKey(0, 5))?.rich).toEqual([{ text: 'Bo', font: { bold: true, color: '#112233' } }, { text: 'ld ' }]);
    const st = wb.styles[s.cells.get(cellKey(0, 5))!.s!]!;
    expect(st.font).toEqual({ bold: true, color: applyTint('#112233', 0.39997558519241921) });
    expect(st.fill).toEqual({ color: '#FF0000' });
    expect(st.numFmt).toBe('0.00%');
    expect(st.align).toEqual({ h: 'center', rotation: -45 });
    expect(s.conditionalFormats).toEqual([{ ranges: [{ r1: 0, c1: 0, r2: 8, c2: 0 }, { r1: 0, c1: 2, r2: 0, c2: 2 }], rules: [{ type: 'expression', formula: 'ISFORMULA(A1)', style: { font: { color: '#9C0006' }, fill: { color: '#FFC7CE' } } }], priority: 2 }]);
  });

  test('tint matches the documented HLS adjustment', () => {
    expect(applyTint('#4472C4', 0.7999816888943144)).toBe('#DAE3F3');
    expect(applyTint('#4472C4', -0.249977111117893)).toBe('#2F5597');
    expect(applyTint('#FFFFFF', -0.1499984740745262)).toBe('#D9D9D9');
  });

  test('dynamic arrays from metadata; spilled values are dropped', () => {
    const md = `<metadata xmlns="${MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="x"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
    const data = zipParts({
      ...Object.fromEntries(Object.entries(unzipSync(handMade('<sheetData><row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">_xlfn._xlws.SORT(B1:B3)</f><v>1</v></c><c r="B1"><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData>'))).map(([k, v]) => [k, strFromU8(v)])),
      'xl/metadata.xml': md,
      'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/s.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>',
    });
    const s = readXlsx(data).sheets[0]!;
    expect(s.cells.get(cellKey(0, 0))).toEqual({ v: 1, f: 'SORT(B1:B3)', dyn: true });
    expect(s.cells.has(cellKey(1, 0))).toBe(false);
  });

  test('print area and sheet-scoped names', () => {
    const wb = readXlsx(handMade('<sheetData/>', {}, `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S!$A$1:$D$20</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!$A$1:$B$2</definedName><definedName name="Loc" localSheetId="0">S!$B$2</definedName></definedNames>`));
    expect(wb.sheets[0]!.pageSetup?.printArea).toBe('$A$1:$D$20');
    expect(wb.names).toEqual([{ name: 'Loc', formula: 'S!$B$2', sheet: 0 }]);
  });

  test('rejects non-workbooks', () => {
    expect(() => readXlsx(zipParts({ 'a.txt': 'x' }))).toThrow();
  });
});

describe('formula helpers', () => {
  test('shiftFormula', () => {
    expect(shiftFormula('A1+B$2+$C3+$D$4', 1, 1)).toBe('B2+C$2+$C4+$D$4');
    expect(shiftFormula('SUM(A:A)+SUM(1:1)+LOG10(A1)', 2, 1)).toBe('SUM(B:B)+SUM(3:3)+LOG10(B3)');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(shiftFormula('"B2"&Sheet2!B2&\'A b\'!B2', 1, 0)).toBe('"B2"&Sheet2!B3&\'A b\'!B3');
    expect(shiftFormula('MyName+R2D2(1)+1.5E3', 1, 1)).toBe('MyName+R2D2(1)+1.5E3');
  });
  test('future prefixes', () => {
    expect(addFuturePrefixes('IFS(A1>1,"x",TRUE,"y")+STDEV.S(A:A)')).toBe('_xlfn.IFS(A1>1,"x",TRUE,"y")+_xlfn.STDEV.S(A:A)');
    expect(addFuturePrefixes('FILTER(A1:A9,B1:B9>0)')).toBe('_xlfn._xlws.FILTER(A1:A9,B1:B9>0)');
    expect(addFuturePrefixes('LAMBDA(a,b,a+b)(1,2)')).toBe('_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)(1,2)');
    expect(addFuturePrefixes('SUM(A1)+"XLOOKUP("')).toBe('SUM(A1)+"XLOOKUP("');
    expect(stripFuturePrefixes('_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))&"_xlfn.x"')).toBe('SORT(UNIQUE(A1:A3))&"_xlfn.x"');
  });
});

// Optional cross-check with an independent reader when one is installed.
const hasPython = ((): boolean => {
  try { execFileSync('python3', ['-c', 'import openpyxl'], { stdio: 'ignore' }); return true; } catch { return false; }
})();

test.skipIf(!hasPython)('written file loads in openpyxl', () => {
  const dir = mkdtempSync(join(tmpdir(), 'cg-'));
  const file = join(dir, 'out.xlsx');
  writeFileSync(file, writeXlsx(build()));
  const script = `import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.worksheets[0]
print(wb.sheetnames, ws['B2'].value, ws['B2'].number_format, ws['A1'].font.b, ws.freeze_panes, len(ws.conditional_formatting), ws['A1'].comment.text, list(ws.merged_cells.ranges)[0].coord, wb.worksheets[1].tables['Table1'].ref)`;
  const out = execFileSync('python3', ['-I', '-c', script, file], { encoding: 'utf8' }).trim();
  expect(out).toBe(`['Data & <Stuff>', 'Second', 'Hidden'] 1234.5 "$"#,##0.00 True B2 3 Header note A13:C14 A1:B3`);
});
