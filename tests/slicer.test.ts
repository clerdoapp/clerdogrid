import { beforeAll, describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { unzipSync, strFromU8, strToU8, zipSync } from 'fflate';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import type { Scalar } from '../src/engine/types';
import { readXlsx } from '../src/formats/xlsx-read';
import { writeXlsx } from '../src/formats/xlsx-write';
import { createPivot } from '../src/features/pivot/apply';
import { newPivotTable, placeField } from '../src/features/pivot/model';
import {
  connectable, deleteSlicer, insertPivotSlicers, insertTableSlicers, insertTimeline, setConnections, setSlicerHidden, setTimelineRange, slicerItems,
} from '../src/features/pivot/slicers';

beforeAll(() => registerAllFunctions());

const DATA: Scalar[][] = [
  ['Region', 'Product', 'Date', 'Sales'],
  ['East', 'Apple', 45292, 10], ['East', 'Pear', 45383, 20], ['West', 'Apple', 45300, 5], ['West', 'Pear', 45700, 15], ['North', 'Apple', 45400, 7],
];

const book = () => {
  const s = new GridStore();
  s.change('data', 'sheet', () => DATA.forEach((r, i) => r.forEach((v, j) => s.engine.setValue(0, i, j, v))));
  s.selectRange({ r1: 1, c1: 2, r2: 5, c2: 2 });
  s.applyStyle({ numFmt: 'm/d/yyyy' });
  const src = { kind: 'range' as const, sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 5, c2: 3 } };
  const build = (pt: ReturnType<typeof newPivotTable>, c: Parameters<typeof placeField>[1]) => placeField(placeField(pt, c, 1, 'rows'), c, 3, 'data');
  createPivot(s, { source: src, dest: { sheet: 0, row: 0, col: 6 }, build }, newPivotTable);
  createPivot(s, { source: src, dest: { sheet: 0, row: 0, col: 10 }, useCache: 1, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable);
  s.setActiveSheet(0);
  return s;
};
const col = (s: GridStore, c: number, r1: number, r2: number, sheet = 0) => Array.from({ length: r2 - r1 + 1 }, (_, i) => s.value(r1 + i, c, sheet));

describe('slicers', () => {
  it('filters connected PivotTables and shows items with data', () => {
    const s = book();
    insertPivotSlicers(s, 0, 'PivotTable1', [0]);
    const sc = s.wb.slicerCaches![0]!;
    expect(sc.name).toBe('Slicer_Region');
    expect(slicerItems(s, sc).map((i) => i.label)).toEqual(['East', 'North', 'West']);
    // Region is index 0 East, 1 North, 2 West in the cache.
    setSlicerHidden(s, 'Slicer_Region', ['1', '2']);
    expect(col(s, 6, 0, 3)).toEqual(['Row Labels', 'Apple', 'Pear', 'Grand Total']);
    expect(col(s, 7, 1, 3)).toEqual([10, 20, 30]);
    // Not yet connected to the second table.
    expect(s.value(4, 11)).toBe(57);
    expect(connectable(s, 1)).toEqual([{ sheet: 'Sheet1', name: 'PivotTable1' }, { sheet: 'Sheet1', name: 'PivotTable2' }]);
    setConnections(s, 'Slicer_Region', connectable(s, 1));
    expect(col(s, 10, 1, 2)).toEqual(['East', 'Grand Total']);
    s.undo();
    expect(s.value(4, 11)).toBe(57);
    s.redo();
    // A second slicer on Product reports which regions still have data.
    insertPivotSlicers(s, 0, 'PivotTable2', [1]);
    setSlicerHidden(s, 'Slicer_Product', ['0']);
    const regions = slicerItems(s, s.wb.slicerCaches![0]!);
    expect(regions.filter((i) => i.hasData).map((i) => i.label)).toEqual(['East', 'West']);
    deleteSlicer(s, s.sheet.slicers![0]!.name);
    expect(s.wb.slicerCaches!.length).toBe(1);
    expect(s.value(1, 10)).toBe('East');
  });

  it('filters a table by hiding rows', () => {
    const s = new GridStore();
    s.change('data', 'book', () => {
      DATA.forEach((r, i) => r.forEach((v, j) => s.engine.setValue(0, i, j, v)));
      s.sheet.tables.push({ name: 'Sales', range: { r1: 0, c1: 0, r2: 5, c2: 3 }, headerRow: true, totalsRow: false, columns: DATA[0]!.map((n) => ({ name: String(n) })) });
    });
    insertTableSlicers(s, 0, 'Sales', [0, 1]);
    setSlicerHidden(s, 'Slicer_Region', ['East']);
    expect([1, 2, 3].map((r) => !!s.sheet.rows.get(r)?.hidden)).toEqual([true, true, false]);
    setSlicerHidden(s, 'Slicer_Product', ['Pear']);
    expect([3, 4, 5].map((r) => !!s.sheet.rows.get(r)?.hidden)).toEqual([false, true, false]);
    expect(slicerItems(s, s.wb.slicerCaches![1]!).map((i) => [i.label, i.hasData])).toEqual([['Apple', true], ['Pear', true]]);
    setSlicerHidden(s, 'Slicer_Region', []);
    expect([1, 2].map((r) => !!s.sheet.rows.get(r)?.hidden)).toEqual([false, true]);
    // Saved the standard way: the table's own filter holds the selection.
    const bytes = writeXlsx(s.wb);
    const files = unzipSync(bytes);
    const table = strFromU8(files['xl/tables/table1.xml']!);
    expect(table).toContain('<autoFilter ref="A1:D6"><filterColumn colId="1"><filters><filter val="Apple"/></filters></filterColumn></autoFilter>');
    // Without ClerdoGrid's own copy the selection still comes back.
    for (const k of Object.keys(files)) if (k.startsWith('xl/slicerCaches/')) files[k] = strToU8(strFromU8(files[k]!).replace(/<x:ext uri="\{4B7C1E92[^]*?<\/x:ext>/, ''));
    expect(readXlsx(zipSync(files)).slicerCaches!.map((c) => c.hidden)).toEqual([[], ['Pear']]);
    const back = readXlsx(bytes);
    expect(back.slicerCaches!.map((c) => [c.name, c.source, c.hidden])).toEqual([['Slicer_Region', { kind: 'table', table: 'Sales', column: 0 }, []], ['Slicer_Product', { kind: 'table', table: 'Sales', column: 1 }, ['Pear']]]);
    expect(back.sheets[0]!.slicers!.map((x) => x.caption)).toEqual(['Region', 'Product']);
  });

  it('timelines filter by a date range', () => {
    const s = book();
    insertTimeline(s, 0, 'PivotTable1', 2);
    setTimelineRange(s, 'NativeTimeline_Date', 45292, 45382); // Jan to Mar 2024
    expect(col(s, 6, 1, 3)).toEqual(['Apple', 'Grand Total', null]);
    expect(s.value(1, 7)).toBe(15);
    setTimelineRange(s, 'NativeTimeline_Date', undefined, undefined);
    expect(s.value(3, 7)).toBe(57);
  });

  it('saves and reopens slicers, timelines and their filters', () => {
    const s = book();
    insertPivotSlicers(s, 0, 'PivotTable1', [0]);
    setConnections(s, 'Slicer_Region', connectable(s, 1));
    setSlicerHidden(s, 'Slicer_Region', ['1']);
    insertTimeline(s, 0, 'PivotTable1', 2);
    setTimelineRange(s, 'NativeTimeline_Date', 45292, 45657);
    const bytes = writeXlsx(s.wb);
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['xl/slicers/slicer1.xml', 'xl/slicerCaches/slicerCache1.xml', 'xl/timelines/timeline1.xml', 'xl/timelineCaches/timelineCache2.xml']));
    const wbXml = strFromU8(files['xl/workbook.xml']!);
    expect(wbXml).toContain('<x14:slicerCaches><x14:slicerCache r:id=');
    expect(wbXml).toContain('<x15:timelineCacheRefs>');
    expect(wbXml).toContain('<definedName name="Slicer_Region">#N/A</definedName>');
    const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']!);
    expect(sheetXml).toContain('<x14:slicerList><x14:slicer r:id=');
    expect(strFromU8(files['xl/drawings/drawing1.xml']!)).toContain('<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Region"/>');
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('application/vnd.ms-excel.slicerCache+xml');
    const back = readXlsx(bytes);
    expect(back.names.some((n) => n.name === 'Slicer_Region')).toBe(false);
    const sc = back.slicerCaches!.find((c) => c.name === 'Slicer_Region')!;
    expect(sc.hidden).toEqual(['1']);
    expect(sc.source).toEqual({ kind: 'pivot', cacheId: 1, field: 0, pivots: [{ sheet: 'Sheet1', name: 'PivotTable1' }, { sheet: 'Sheet1', name: 'PivotTable2' }] });
    const tl = back.slicerCaches!.find((c) => c.timeline)!;
    expect(tl.timeline).toMatchObject({ start: 45292, end: 45657, level: 'months' });
    expect(back.sheets[0]!.slicers!.map((x) => [x.name, x.kind, x.anchor])).toEqual(s.sheet.slicers!.map((x) => [x.name, x.kind, x.anchor]));
    // The tables still carry the filters and refresh with them.
    const s2 = new GridStore(back);
    expect(s2.value(1, 10)).toBe('East');
    expect(back.sheets[0]!.pivots![0]!.fields[2]!.filters).toEqual([{ kind: 'dates', start: 45292, end: 45657, timeline: 'NativeTimeline_Date' }]);
  });

  const soffice = ['/usr/bin/soffice', '/usr/local/bin/soffice'].find((p) => existsSync(p));
  it.runIf(!!soffice)('a file with slicers opens in another app, with the filter applied', () => {
    const s = book();
    // Product is a row field of PivotTable1: hide Pear (cache item 1).
    insertPivotSlicers(s, 0, 'PivotTable1', [1]);
    setSlicerHidden(s, 'Slicer_Product', ['1']);
    expect(col(s, 6, 1, 2)).toEqual(['Apple', 'Grand Total']);
    const dir = mkdtempSync(join(tmpdir(), 'slicer-'));
    writeFileSync(join(dir, 'book.xlsx'), writeXlsx(s.wb));
    execFileSync(soffice!, ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'book.xlsx')], { timeout: 120000, stdio: 'ignore' });
    const lines = readFileSync(join(dir, 'book.csv'), 'utf8').trim().split(/\r?\n/).map((l) => l.split(','));
    // Columns G:H hold PivotTable1 as the other app rebuilt it.
    expect(lines.slice(1, 3).map((l) => l.slice(6, 8))).toEqual([['Apple', '22'], [expect.stringMatching(/total/i), '22']]);
  }, 180000);
});
