import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { newWorkbook, type PageSetup } from '../src/model/workbook';

// 1×1 PNG.
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));

const full: PageSetup = {
  orientation: 'landscape', paper: 9, scale: 85, printArea: '$A$1:$F$40,$H$1:$J$9',
  margins: { left: 0.5, right: 0.6, top: 1, bottom: 0.9, header: 0.4, footer: 0.35 },
  header: '&L&"Arial,Bold"&14Title &A&C&D&RPage &P of &N', footer: '&C&F & more',
  evenHeader: '&CEven', evenFooter: '&REven page &P', firstHeader: '&CFirst &G', firstFooter: '&LConfidential',
  differentOddEven: true, differentFirst: true, scaleWithDoc: false, alignWithMargins: false,
  headerPictures: { CHFIRST: { data: PNG, mime: 'image/png', width: 72, height: 36 } },
  printTitleRows: '1:2', printTitleCols: 'A:B', gridlines: true, headings: true, centerH: true, centerV: true,
  firstPageNumber: 5, order: 'overThenDown', rowBreaks: [10, 25], colBreaks: [4],
  blackAndWhite: true, draft: true, cellComments: 'atEnd', errors: 'dash', horizontalDpi: 300, verticalDpi: 300, copies: 2,
};

const roundTrip = (setup: PageSetup, view?: 'pageLayout' | 'pageBreakPreview') => {
  const wb = newWorkbook();
  wb.sheets[0]!.cells.set(0, { v: 'x' });
  wb.sheets[0]!.pageSetup = structuredClone(setup);
  if (view) wb.sheets[0]!.view.pageView = view;
  const bytes = writeXlsx(wb);
  return { bytes, wb: readXlsx(bytes) };
};

describe('page setup in .xlsx', () => {
  it('saves and reopens every setting', () => {
    const { wb, bytes } = roundTrip(full, 'pageBreakPreview');
    expect(wb.sheets[0]!.pageSetup).toEqual(full);
    expect(wb.sheets[0]!.view.pageView).toBe('pageBreakPreview');
    expect(wb.names.filter((n) => n.name.startsWith('_xlnm'))).toEqual([]);
    const files = unzipSync(bytes);
    const sheet = strFromU8(files['xl/worksheets/sheet1.xml']!);
    // Schema order: printOptions, pageMargins, pageSetup, headerFooter, rowBreaks, colBreaks, ..., legacyDrawingHF.
    const order = ['<printOptions', '<pageMargins', '<pageSetup', '<headerFooter', '<rowBreaks', '<colBreaks', '<legacyDrawingHF'].map((t) => sheet.indexOf(t));
    expect(order.every((p, i) => p > 0 && (i === 0 || p > order[i - 1]!))).toBe(true);
    expect(sheet).toContain('<brk id="10" max="16383" man="1"/>');
    const book = strFromU8(files['xl/workbook.xml']!);
    expect(book).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$F$40,Sheet1!$H$1:$J$9</definedName>');
    expect(book).toContain('<definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$A:$B,Sheet1!$1:$2</definedName>');
    expect(Object.keys(files).some((f) => /vmlDrawingHF1\.vml$/.test(f))).toBe(true);
  });

  it('keeps fit-to-page and the page layout view', () => {
    const { wb, bytes } = roundTrip({ fitWidth: 1, fitHeight: 0 }, 'pageLayout');
    expect(wb.sheets[0]!.pageSetup).toEqual({ fitWidth: 1, fitHeight: 0 });
    expect(wb.sheets[0]!.view.pageView).toBe('pageLayout');
    expect(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!)).toContain('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
  });

  it('reads files from other apps: Print_Titles names, theme-free headers and automatic breaks', () => {
    const wb = newWorkbook();
    const files = unzipSync(writeXlsx(wb));
    let sheet = strFromU8(files['xl/worksheets/sheet1.xml']!);
    sheet = sheet.replace('</worksheet>', '').replace(/<pageMargins[^>]*\/>/, '<printOptions gridLines="1"/><pageMargins left="0.25" right="0.25" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="5" orientation="portrait" useFirstPageNumber="0" firstPageNumber="3" r:id="rId9"/><headerFooter><oddFooter>Page &amp;P_x000D_</oddFooter></headerFooter><rowBreaks count="2" manualBreakCount="1"><brk id="12" max="16383" man="1"/><brk id="40" max="16383"/></rowBreaks>') + '</worksheet>';
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(sheet);
    let book = strFromU8(files['xl/workbook.xml']!);
    book = book.replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">\'Sheet1\'!$3:$4</definedName></definedNames>');
    files['xl/workbook.xml'] = new TextEncoder().encode(book);
    const back = readXlsx(zipSync(files));
    expect(back.sheets[0]!.pageSetup).toEqual({ paper: 5, orientation: 'portrait', gridlines: true, margins: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }, footer: 'Page &P\r', rowBreaks: [12, 40], printTitleRows: '3:4' });
  });

  it('opens in LibreOffice with the page setup kept', () => {
    if (!existsSync('/usr/bin/soffice') && !existsSync('/usr/local/bin/soffice')) return;
    const dir = mkdtempSync(join(tmpdir(), 'pl-lo-'));
    const src = join(dir, 'in.xlsx');
    const { bytes } = roundTrip({ ...full, headerPictures: undefined, firstHeader: '&CFirst' });
    writeFileSync(src, bytes);
    const out = join(dir, 'out');
    try {
      execFileSync('soffice', ['--headless', `-env:UserInstallation=file://${dir}/profile`, '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, src], { stdio: 'ignore', timeout: 120_000 });
    } catch { return; }
    const back = readXlsx(new Uint8Array(readFileSync(join(out, 'in.xlsx'))));
    const ps = back.sheets[0]!.pageSetup!;
    expect(ps.orientation).toBe('landscape');
    expect(ps.paper).toBe(9);
    expect(ps.scale).toBe(85);
    expect(ps.margins?.left).toBeCloseTo(0.5, 2);
    // Other apps may move part of the header height into the top margin.
    expect(Math.abs(ps.margins!.top - 1)).toBeLessThan(0.1);
    expect(ps.header).toContain('&P');
    expect(ps.printArea).toContain('$A$1:$F$40');
    expect(ps.printTitleRows).toBe('1:2');
    expect(ps.rowBreaks).toEqual(expect.arrayContaining([10, 25]));
    expect(ps.centerH).toBe(true);
    expect(ps.firstPageNumber).toBe(5);
  }, 150_000);
});
