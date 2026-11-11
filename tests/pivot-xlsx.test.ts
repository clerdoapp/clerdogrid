import { beforeAll, describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import type { Scalar } from '../src/engine/types';
import { readXlsx } from '../src/formats/xlsx-read';
import { writeXlsx } from '../src/formats/xlsx-write';
import { createPivot, layoutOf, replaceCache, refreshPivots } from '../src/features/pivot/apply';
import { addCalcField, groupByRange } from '../src/features/pivot/cache';
import { computePivot } from '../src/features/pivot/compute';
import { newPivotTable, placeField, setField } from '../src/features/pivot/model';
import { insertPivotChart } from '../src/features/pivot/ops';
import { VALUES_FIELD } from '../src/model/pivot';

beforeAll(() => registerAllFunctions());

const DATA: Scalar[][] = [
  ['Region', 'Product', 'Date', 'Sales', 'Units'],
  ['East', 'Apple', 45292, 10, 1], ['East', 'Pear', 45383, 20, 2], ['East', 'Apple', 45658, 30, 3],
  ['West', 'Apple', 45300, 5, 1], ['West', 'Pear', 45700, 15, 5], ['West', 'Pear', 45710, 25, 5], ['North', 'Apple', 45400, 7, 7],
];

const book = () => {
  const s = new GridStore();
  s.change('data', 'sheet', () => DATA.forEach((r, i) => r.forEach((v, j) => s.engine.setValue(0, i, j, v))));
  s.selectRange({ r1: 1, c1: 2, r2: 7, c2: 2 });
  s.applyStyle({ numFmt: 'm/d/yyyy' });
  return s;
};

export const cellsOf = (wb: ReturnType<typeof readXlsx>, sheet: number, r1: number, c1: number, r2: number, c2: number) => {
  const s = new GridStore(wb);
  return Array.from({ length: r2 - r1 + 1 }, (_, i) => Array.from({ length: c2 - c1 + 1 }, (_, j) => s.value(r1 + i, c1 + j, sheet)));
};

const soffice = ['/usr/bin/soffice', '/usr/local/bin/soffice'].find((p) => existsSync(p));

describe('pivot file format', () => {
  it('saves and reopens a live table with settings', () => {
    const s = book();
    const res = createPivot(s, {
      source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 4 } }, dest: { sheet: 0, row: 2, col: 8 },
      build: (pt, c) => {
        let p = placeField(pt, c, 0, 'rows');
        p = placeField(p, c, 1, 'rows');
        p = placeField(p, c, 3, 'data');
        p = placeField(p, c, 4, 'data');
        p = setField(p, 0, { items: p.fields[0]!.items.map((it) => (it.x === 1 ? { ...it, collapsed: true } : it)), sort: 'descending' });
        p = setField(p, 1, { filters: [{ kind: 'label', op: 'contains', value1: 'p' }], subtotals: ['sum', 'max'] });
        p = { ...p, data: [p.data[0]!, { ...p.data[1]!, agg: 'average', showAs: 'percentOfParentRow', numFmt: '0.0%' }], style: { ...p.style, name: 'PivotStyleMedium7', bandedRows: true }, emptyText: '-', showEmpty: true };
        return p;
      },
    }, newPivotTable) as { sheet: number; name: string };
    const before = s.wb.sheets[0]!.pivots![0]!;
    const bytes = writeXlsx(s.wb);
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotCache/pivotCacheRecords1.xml', 'xl/pivotTables/pivotTable1.xml']));
    const ct = strFromU8(files['[Content_Types].xml']!);
    expect(ct).toContain('spreadsheetml.pivotTable+xml');
    expect(ct).toContain('spreadsheetml.pivotCacheDefinition+xml');
    expect(strFromU8(files['xl/workbook.xml']!)).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
    expect(strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels']!)).toContain('relationships/pivotTable');
    const back = readXlsx(bytes);
    const pt = back.sheets[0]!.pivots![0]!;
    expect(back.pivotCaches![0]!.records).toEqual(s.wb.pivotCaches![0]!.records);
    expect(back.pivotCaches![0]!.fields[2]!.dates).toBe(true);
    expect(pt.rows).toEqual(before.rows);
    expect(pt.cols).toEqual([VALUES_FIELD]);
    expect(pt.data).toEqual(before.data);
    expect(pt.fields[0]!.sort).toBe('descending');
    expect(pt.fields[0]!.items.find((it) => it.x === 1)!.collapsed).toBe(true);
    expect(pt.fields[1]!.filters).toEqual(before.fields[1]!.filters);
    expect(pt.fields[1]!.subtotals).toEqual(['sum', 'max']);
    expect(pt.style.name).toBe('PivotStyleMedium7');
    expect(pt.emptyText).toBe('-');
    expect(pt.location).toEqual(before.location);
    expect(pt.rowItems).toEqual(before.rowItems);
    expect(pt.colItems).toEqual(before.colItems);
    // Same cells from the reopened definition.
    const a = computePivot(before, s.wb.pivotCaches![0]!).cells.map((r) => r.map((c) => c?.v ?? null));
    const b = computePivot(pt, back.pivotCaches![0]!).cells.map((r) => r.map((c) => c?.v ?? null));
    expect(b).toEqual(a);
    // And it still refreshes.
    const s2 = new GridStore(back);
    expect(layoutOf(s2.wb, pt)).toBeTruthy();
    void res;
  });

  it('round-trips date groups, number groups, manual groups and calculated fields and items', () => {
    const s = book();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 4 } }, dest: null }, newPivotTable) as { sheet: number; name: string };
    let cache = s.wb.pivotCaches![0]!;
    const g = groupByRange(cache, 2, [{ by: 'months' }, { by: 'years' }]);
    cache = addCalcField(g.cache, 'Price', 'Sales/Units');
    replaceCache(s, cache, null, 'Group', { sheet: res.sheet, name: res.name, fn: (pt) => {
      let p = placeField(pt, cache, g.fields[1]!, 'rows');
      p = placeField(p, cache, 2, 'rows');
      p = placeField(p, cache, 0, 'cols');
      p = placeField(p, cache, 5, 'data');
      p = setField(p, 0, { items: [...p.fields[0]!.items, { x: -1, name: 'EastWest', formula: 'East+West' }] });
      return p;
    } });
    const before = s.wb.sheets[res.sheet]!.pivots![0]!;
    const bytes = writeXlsx(s.wb);
    const def = strFromU8(unzipSync(bytes)['xl/pivotCache/pivotCacheDefinition1.xml']!);
    expect(def).toContain('<rangePr groupBy="months"');
    expect(def).toContain('formula="Sales/Units"');
    expect(def).toContain('<calculatedItem formula="East+West">');
    const back = readXlsx(bytes);
    const c2 = back.pivotCaches![0]!;
    expect(c2.fields.map((f) => f.name)).toEqual(cache.fields.map((f) => f.name));
    expect(c2.fields[2]!.group).toEqual(cache.fields[2]!.group);
    expect(c2.fields[g.fields[1]!]!.group).toEqual(cache.fields[g.fields[1]!]!.group);
    expect(c2.fields[0]!.items).toEqual(cache.fields[0]!.items);
    const pt = back.sheets[res.sheet]!.pivots![0]!;
    expect(pt.fields[0]!.items.at(-1)).toEqual({ x: -1, name: 'EastWest', formula: 'East+West' });
    const a = computePivot(before, cache).cells.map((r) => r.map((c) => c?.v ?? null));
    const b = computePivot(pt, c2).cells.map((r) => r.map((c) => c?.v ?? null));
    expect(b).toEqual(a);
    expect(a[1]).toEqual(['Row Labels', 'East', 'North', 'West', 'EastWest', 'Grand Total']);
  });

  it.runIf(!!soffice)('opens in another spreadsheet app with its values', () => {
    const s = book();
    createPivot(s, {
      source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 4 } }, dest: null,
      build: (pt, c) => placeField(placeField(placeField(pt, c, 0, 'rows'), c, 1, 'cols'), c, 3, 'data'),
    }, newPivotTable);
    // Kept formatting goes along as a standard formats element.
    s.selectRange({ r1: 4, c1: 1, r2: 4, c2: 1 });
    s.applyStyle({ fill: { color: '#ffcc00' } });
    refreshPivots(s);
    expect(strFromU8(unzipSync(writeXlsx(s.wb))['xl/pivotTables/pivotTable1.xml']!)).toContain('<formats count="1">');
    const dir = mkdtempSync(join(tmpdir(), 'pivot-'));
    writeFileSync(join(dir, 'book.xlsx'), writeXlsx(s.wb));
    execFileSync(soffice!, ['--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1', '--outdir', dir, join(dir, 'book.xlsx')], { timeout: 120000, stdio: 'ignore' });
    const csv = readFileSync(join(dir, 'book-Sheet2.csv'), 'utf8').trim().split(/\r?\n/);
    // The other app rebuilds the table from the saved definition (its own captions), with the same values.
    const body = csv.slice(csv.findIndex((l) => l.startsWith('East')));
    expect(body.slice(0, 4)).toEqual(['East,40,20,60', 'North,7,,7', 'West,5,40,45', expect.stringMatching(/,52,60,112$/)]);
    expect(csv.some((l) => /^Sum of Sales,/.test(l))).toBe(true);
    // A file saved by the other app opens here as a live table.
    const out = join(dir, 'out');
    execFileSync(soffice!, ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, join(dir, 'book.xlsx')], { timeout: 120000, stdio: 'ignore' });
    const back = readXlsx(new Uint8Array(readFileSync(join(out, 'book.xlsx'))));
    const ref = back.sheets.flatMap((sh, i) => (sh.pivots ?? []).map((pt) => ({ i, pt })))[0]!;
    expect(ref).toBeTruthy();
    const l = computePivot(ref.pt, back.pivotCaches!.find((c) => c.id === ref.pt.cacheId)!);
    const vals = l.cells.slice(l.firstDataRow).map((r) => r.slice(l.firstDataCol).map((c) => c?.v ?? null));
    expect(vals).toEqual([[40, 20, 60], [7, null, 7], [5, 40, 45], [52, 60, 112]]);
  }, 180000);

  it('PivotCharts save their labels on several levels', () => {
    const s = book();
    createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 4 } }, dest: { sheet: 0, row: 0, col: 8 }, build: (pt, c) => placeField(placeField(placeField(pt, c, 0, 'rows'), c, 1, 'rows'), c, 3, 'data') }, newPivotTable);
    insertPivotChart(s, { sheet: 0, index: 0, pt: s.wb.sheets[0]!.pivots![0]! });
    const ser = s.wb.sheets[0]!.charts[0]!.series[0]!;
    expect(ser.multiLevel!.levels).toEqual([['Apple', 'Pear', 'Apple', 'Apple', 'Pear'], ['East', '', 'North', 'West', '']]);
    const bytes = writeXlsx(s.wb);
    const xml = strFromU8(unzipSync(bytes)['xl/charts/chart1.xml']!);
    expect(xml).toContain('<c:cat><c:multiLvlStrRef><c:f>Sheet1!$I$3:$I$9</c:f><c:multiLvlStrCache><c:ptCount val="5"/><c:lvl><c:pt idx="0"><c:v>Apple</c:v></c:pt>');
    expect(xml).toContain('<c:pivotSource><c:name>[Book1]Sheet1!PivotTable1</c:name>');
    const back = readXlsx(bytes).sheets[0]!.charts[0]!;
    expect(back.pivot).toBe('Sheet1!PivotTable1');
    expect(back.series[0]!.multiLevel!.levels).toEqual(ser.multiLevel!.levels);
    if (soffice) {
      const dir = mkdtempSync(join(tmpdir(), 'pchart-'));
      writeFileSync(join(dir, 'c.xlsx'), bytes);
      execFileSync(soffice, ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'c.xlsx')], { timeout: 120000, stdio: 'ignore' });
      expect(readFileSync(join(dir, 'c.csv'), 'utf8')).toMatch(/East/);
    }
  }, 180000);
});
