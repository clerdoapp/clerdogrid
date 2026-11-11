import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8, zipSync } from 'fflate';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { newWorkbook, newSheet, cellKey } from '../src/model/workbook';
import { checkPassword, makePasswordHash, modernHash, protectionHash } from '../src/app/protection';

const roundTrip = (wb: ReturnType<typeof newWorkbook>) => readXlsx(writeXlsx(wb));
const part = (bytes: Uint8Array, p: string) => strFromU8(unzipSync(bytes)[p]!);

describe('workbook protection in xlsx', () => {
  it('writes and reads legacy and modern workbook hashes', () => {
    const wb = newWorkbook();
    wb.protection = { structure: true, windows: true, password: protectionHash('abc'), hash: makePasswordHash('abc', 1000) };
    const bytes = writeXlsx(wb);
    const xml = part(bytes, 'xl/workbook.xml');
    expect(xml).toMatch(/<workbookProtection workbookPassword="[0-9A-F]{4}" workbookAlgorithmName="SHA-512" workbookHashValue="[^"]+" workbookSaltValue="[^"]+" workbookSpinCount="1000" lockStructure="1" lockWindows="1"\/>/);
    expect(xml.indexOf('<workbookProtection')).toBeGreaterThan(xml.indexOf('<workbookPr'));
    expect(xml.indexOf('<workbookProtection')).toBeLessThan(xml.indexOf('<bookViews'));
    const back = readXlsx(bytes);
    expect(back.protection?.structure).toBe(true);
    expect(back.protection?.windows).toBe(true);
    expect(checkPassword('abc', back.protection!)).toBe(true);
    expect(checkPassword('abd', back.protection!)).toBe(false);
  });

  it('verifies a modern-only hash and a legacy-only hash', () => {
    const h = makePasswordHash('pässword', 5);
    expect(checkPassword('pässword', { hash: h })).toBe(true);
    expect(checkPassword('password', { hash: h })).toBe(false);
    expect(checkPassword('abc', { password: protectionHash('abc') })).toBe(true);
    expect(modernHash('x', h.saltValue, 5, 'MD5')).toBeNull();
  });

  it('keeps modern sheet protection hashes and option flags', () => {
    const wb = newWorkbook();
    wb.sheets[0]!.protection = { sheet: true, hash: makePasswordHash('s3cret', 100), options: { formatCells: false, sort: false, selectLockedCells: true } };
    const back = roundTrip(wb);
    const p = back.sheets[0]!.protection!;
    expect(p.hash?.algorithmName).toBe('SHA-512');
    expect(p.options).toMatchObject({ formatCells: false, sort: false, selectLockedCells: true });
    expect(p.options).not.toHaveProperty('hashValue');
    expect(checkPassword('s3cret', p)).toBe(true);
  });
});

describe('views in xlsx', () => {
  it('round-trips named sheet views in the shared part and keeps hidden rows', () => {
    const wb = newWorkbook();
    const s = wb.sheets[0]!;
    s.sheetViews = [
      { name: 'Only north', hiddenRows: [4, 5], hiddenCols: [2], filter: { range: { r1: 0, c1: 0, r2: 9, c2: 2 }, columns: [{ col: 1, values: ['North'] }], sort: [{ col: 2, descending: true }] } },
      { name: 'All', hiddenRows: [], hiddenCols: [] },
    ];
    const bytes = writeXlsx(wb);
    const nsv = part(bytes, 'xl/namedSheetViews/namedSheetView1.xml');
    expect(nsv).toContain('xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews"');
    expect(nsv).toContain('<x:filter val="North"/>');
    expect(part(bytes, 'xl/worksheets/_rels/sheet1.xml.rels')).toContain('relationships/namedSheetView');
    expect(part(bytes, '[Content_Types].xml')).toContain('application/vnd.ms-excel.namedsheetviews+xml');
    const back = readXlsx(bytes).sheets[0]!.sheetViews!;
    expect(back).toHaveLength(2);
    expect(back[0]).toMatchObject({ name: 'Only north', hiddenRows: [4, 5], hiddenCols: [2] });
    expect(back[0]!.filter?.columns[0]).toEqual({ col: 1, values: ['North'] });
    expect(back[0]!.filter?.sort).toEqual([{ col: 2, descending: true }]);
    expect(back[0]!.id).toMatch(/^\{[0-9A-F-]{36}\}$/);
  });

  it('reads named sheet views from the shared part alone', () => {
    const wb = newWorkbook();
    wb.sheets[0]!.sheetViews = [{ name: 'V', hiddenRows: [], hiddenCols: [], filter: { range: { r1: 0, c1: 0, r2: 5, c2: 1 }, columns: [{ col: 0, custom: { rules: [{ operator: 'greaterThan', value: '3' }] } }] } }];
    const files = unzipSync(writeXlsx(wb));
    delete files['xl/clerdogrid/workspace.json'];
    const back = readXlsx(zipSync(files)).sheets[0]!.sheetViews!;
    expect(back[0]!.filter?.columns[0]?.custom?.rules[0]).toEqual({ operator: 'greaterThan', value: '3' });
  });

  it('round-trips custom views with print settings, filters and hidden rows', () => {
    const wb = newWorkbook();
    wb.sheets.push(newSheet('Two'));
    wb.customViews = [{
      name: 'Review', guid: '{11111111-2222-4333-8444-555555555555}', includePrint: true, includeHidden: true, activeSheet: 1,
      sheets: [
        { sheet: 0, zoom: 80, showGrid: false, hiddenRows: [3], hiddenCols: [1, 2], selection: { r1: 1, c1: 1, r2: 2, c2: 2 }, filter: { range: { r1: 0, c1: 0, r2: 4, c2: 1 }, columns: [{ col: 0, values: ['a'] }] }, pageSetup: { orientation: 'landscape', margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 } } },
        { sheet: 1, state: 'hidden', frozenRows: 1 },
      ],
    }];
    const bytes = writeXlsx(wb);
    expect(part(bytes, 'xl/workbook.xml')).toContain('<customWorkbookView name="Review" guid="{11111111-2222-4333-8444-555555555555}"');
    const s1 = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(s1).toContain('<customSheetViews><customSheetView guid="{11111111-2222-4333-8444-555555555555}" scale="80" showGridLines="0"');
    const back = readXlsx(bytes).customViews!;
    expect(back[0]).toMatchObject({ name: 'Review', activeSheet: 1, includePrint: true });
    const v0 = back[0]!.sheets.find((x) => x.sheet === 0)!;
    expect(v0).toMatchObject({ zoom: 80, showGrid: false, hiddenRows: [3], hiddenCols: [1, 2] });
    expect(v0.pageSetup?.orientation).toBe('landscape');
    expect(v0.filter?.columns).toEqual([{ col: 0, values: ['a'] }]);
    expect(back[0]!.sheets.find((x) => x.sheet === 1)).toMatchObject({ state: 'hidden', frozenRows: 1 });
  });

  it('round-trips split panes, ignored errors, calc options and the 1904 date system', () => {
    const wb = newWorkbook();
    wb.date1904 = true;
    wb.calc = { manual: false, autoNoTable: true, precisionAsDisplayed: true, r1c1: true };
    const s = wb.sheets[0]!;
    s.view.split = { x: 200, y: 120, scrollX: 64, scrollY: 40 };
    s.ignoredErrors = [{ range: { r1: 0, c1: 0, r2: 3, c2: 0 }, types: ['numberStoredAsText'] }];
    s.cells.set(cellKey(0, 0), { v: '12' });
    const bytes = writeXlsx(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toMatch(/<pane xSplit="3375" ySplit="2100" activePane="bottomRight"\/>/);
    expect(xml).toContain('<ignoredErrors><ignoredError sqref="A1:A4" numberStoredAsText="1"/></ignoredErrors>');
    expect(part(bytes, 'xl/workbook.xml')).toMatch(/calcMode="autoNoTable" fullPrecision="0" refMode="R1C1"/);
    const back = readXlsx(bytes);
    expect(back.date1904).toBe(true);
    expect(back.calc).toMatchObject({ autoNoTable: true, precisionAsDisplayed: true, r1c1: true });
    expect(back.sheets[0]!.view.split).toEqual({ x: 200, y: 120, scrollX: 64, scrollY: 40 });
    expect(back.sheets[0]!.ignoredErrors).toEqual(s.ignoredErrors);
  });
});
