import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import type { Scalar } from '../src/engine/types';
import { cellKey } from '../src/model/workbook';
import { createPivot, drillDown, refreshPivots, updatePivot, deletePivot, layoutOf, pivotChartSeries } from '../src/features/pivot/apply';
import { newPivotTable, placeField } from '../src/features/pivot/model';
import { moveItem, groupSelection, ungroup, saveCalcField, showFilterPages, setPageItem } from '../src/features/pivot/ops';
import { readXlsx } from '../src/formats/xlsx-read';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { changeStructure, renameSheet } from '../src/app/ops/structure';
import { refreshOnOpen } from '../src/features/pivot/apply';
import { writeXlsx } from '../src/formats/xlsx-write';
import { getPivotDataText } from '../src/features/pivot/getpivotdata';

beforeAll(() => registerAllFunctions());

const DATA: Scalar[][] = [
  ['Region', 'Product', 'Year', 'Sales'],
  ['East', 'Apple', 2023, 10], ['East', 'Pear', 2023, 20], ['East', 'Apple', 2024, 30],
  ['West', 'Apple', 2023, 5], ['West', 'Pear', 2024, 15], ['West', 'Pear', 2024, 25], ['North', 'Apple', 2024, 7],
];

export const setupBook = () => {
  const s = new GridStore();
  s.change('data', 'sheet', () => DATA.forEach((r, i) => r.forEach((v, j) => s.engine.setValue(0, i, j, v))));
  return s;
};

const grid = (s: GridStore, sheet: number, r1: number, c1: number, r2: number, c2: number) =>
  Array.from({ length: r2 - r1 + 1 }, (_, i) => Array.from({ length: c2 - c1 + 1 }, (_, j) => s.value(r1 + i, c1 + j, sheet)));

describe('PivotTables in a workbook', () => {
  it('creates, refreshes, undoes and deletes a table', () => {
    const s = setupBook();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 } }, newPivotTable);
    expect(typeof res).toBe('object');
    const name = (res as { name: string }).name;
    updatePivot(s, 0, name, (pt) => placeField(placeField(pt, s.wb.pivotCaches![0]!, 0, 'rows'), s.wb.pivotCaches![0]!, 3, 'data'));
    expect(grid(s, 0, 0, 6, 4, 7)).toEqual([['Row Labels', 'Sum of Sales'], ['East', 60], ['North', 7], ['West', 45], ['Grand Total', 112]]);
    // Change the source and refresh.
    s.change('edit', 'sheet', () => s.engine.setValue(0, 7, 3, 100));
    expect(refreshPivots(s)).toBeNull();
    expect(s.value(2, 7)).toBe(100);
    expect(s.value(4, 7)).toBe(205);
    s.undo();
    expect(s.value(2, 7)).toBe(7);
    expect(s.wb.pivotCaches![0]!.records[6]![3]).toBe(7);
    s.redo();
    expect(s.value(2, 7)).toBe(100);
    // GETPIVOTDATA reads the live table.
    s.change('f', 'sheet', () => s.engine.setFormula(0, 10, 0, 'GETPIVOTDATA("Sum of Sales",$G$1,"Region","West")'));
    expect(s.value(10, 0)).toBe(45);
    s.change('f', 'sheet', () => s.engine.setFormula(0, 11, 0, 'GETPIVOTDATA("Sales",G1)'));
    expect(s.value(11, 0)).toBe(205);
    s.change('f', 'sheet', () => s.engine.setFormula(0, 12, 0, 'GETPIVOTDATA("Sales",G1,"Region","South")'));
    expect((s.value(12, 0) as { code: string }).code).toBe('#REF!');
    expect(getPivotDataText(s.wb, 0, 1, 7, 0)).toBe('GETPIVOTDATA("Sum of Sales",$G$1,"Region","East")');
    expect(getPivotDataText(s.wb, 0, 1, 7, 1)).toBe('GETPIVOTDATA("Sum of Sales",Sheet1!$G$1,"Region","East")');
    deletePivot(s, 0, name);
    expect(s.value(0, 6)).toBeNull();
    expect(s.wb.pivotCaches).toBeUndefined();
  });

  it('puts a new table on a new sheet and drills down', () => {
    const s = setupBook();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: null, build: (pt, c) => placeField(placeField(placeField(pt, c, 0, 'rows'), c, 2, 'cols'), c, 3, 'data') }, newPivotTable);
    const { sheet } = res as { sheet: number };
    expect(s.wb.sheets.length).toBe(2);
    expect(s.wb.sheets[sheet]!.name).toBe('Sheet2');
    // A3 holds the table; West/2024 is at row 7 (0-based 6), column C.
    expect(grid(s, sheet, 2, 0, 4, 3)).toEqual([['Sum of Sales', 'Column Labels', null, null], ['Row Labels', 2023, 2024, 'Grand Total'], ['East', 30, 30, 60]]);
    expect(drillDown(s, sheet, 6, 2)).toBe(true);
    const at = s.sheetIndex;
    expect(grid(s, at, 0, 0, 2, 3)).toEqual([['Region', 'Product', 'Year', 'Sales'], ['West', 'Pear', 2024, 15], ['West', 'Pear', 2024, 25]]);
    expect(s.wb.sheets[at]!.tables[0]!.range).toEqual({ r1: 0, c1: 0, r2: 2, c2: 3 });
  });

  it('warns before overwriting cells and refuses overlapping tables', () => {
    const s = setupBook();
    s.change('x', 'sheet', () => s.engine.setValue(0, 3, 7, 'keep'));
    (globalThis as { confirm?: (m: string) => boolean }).confirm = () => false;
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 }, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable);
    expect(res).toBe('');
    expect(s.value(3, 7)).toBe('keep');
    // A refresh that would grow over the cell asks too.
    s.change('x', 'sheet', () => s.engine.setValue(0, 3, 7, null));
    expect(typeof createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 8, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 }, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable)).toBe('object');
    s.change('x', 'sheet', () => { s.engine.setValue(0, 5, 6, 'below'); s.engine.setValue(0, 8, 0, 'South'); s.engine.setValue(0, 8, 3, 1); });
    expect(refreshPivots(s)).toBe('');
    expect(s.value(5, 6)).toBe('below');
    (globalThis as { confirm?: (m: string) => boolean }).confirm = () => true;
    expect(refreshPivots(s)).toBeNull();
    expect(s.value(3, 6)).toBe('South');
    expect(s.value(5, 6)).toBe('Grand Total');
    delete (globalThis as { confirm?: unknown }).confirm;
  });

  it('keeps formatting the user added when refreshed', () => {
    const s = setupBook();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 }, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable);
    const { name } = res as { name: string };
    s.selectRange({ r1: 2, c1: 7, r2: 2, c2: 7 });
    s.applyStyle({ fill: { color: '#ff0000' } });
    refreshPivots(s);
    expect(s.style(2, 7).fill?.color).toBe('#ff0000');
    // Sorting descending moves North; its formatting follows the item.
    updatePivot(s, 0, name, (pt) => ({ ...pt, fields: pt.fields.map((f, i) => (i === 0 ? { ...f, sort: 'descending' } : f)) }));
    expect(s.value(2, 6)).toBe('North');
    expect(s.style(2, 7).fill?.color).toBe('#ff0000');
    const pt = s.wb.sheets[0]!.pivots![0]!;
    expect(pivotChartSeries(s, 0, pt, layoutOf(s.wb, pt)!)[0]).toEqual({ name: '=Sheet1!$H$1', values: 'Sheet1!$H$2:$H$4', categories: 'Sheet1!$G$2:$G$4' });
    expect(s.wb.sheets[0]!.cells.get(cellKey(0, 6))).toBeTruthy();
  });

  it('moves items by hand, groups selected items and keeps formats in the file', () => {
    const s = setupBook();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 }, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable) as { sheet: number; name: string };
    const ref = () => ({ sheet: 0, index: 0, pt: s.wb.sheets[0]!.pivots![0]! });
    moveItem(s, ref(), 0, 2, 0);
    expect(grid(s, 0, 1, 6, 3, 6).flat()).toEqual(['West', 'East', 'North']);
    // Select East and North labels and group them.
    s.selectRange({ r1: 2, c1: 6, r2: 3, c2: 6 });
    expect(groupSelection(s)).toBeNull();
    expect(grid(s, 0, 0, 6, 7, 6).flat()).toEqual(['Row Labels', 'Group1', 'East', 'North', 'West', 'West', 'Grand Total', null]);
    s.selectRange({ r1: 3, c1: 6, r2: 3, c2: 6 });
    expect(ungroup(s)).toBeNull();
    expect(grid(s, 0, 1, 6, 3, 6).flat()).toEqual(['West', 'East', 'North']);
    // A calculated field added from the dialog's helper.
    saveCalcField(s, ref(), 'Double', 'Sales*2', null);
    expect(grid(s, 0, 0, 7, 1, 8)).toEqual([['Sum of Sales', 'Sum of Double'], [45, 90]]);
    // Formatting the user set survives a save and a refresh.
    s.selectRange({ r1: 1, c1: 7, r2: 1, c2: 7 });
    s.applyStyle({ fill: { color: '#ffcc00' } });
    refreshPivots(s);
    const bytes = writeXlsx(s.wb);
    // The standard formats element carries it on its own, without ClerdoGrid's extension.
    const files = unzipSync(bytes);
    const table = strFromU8(files['xl/pivotTables/pivotTable1.xml']!);
    expect(table).toMatch(/<formats count="\d+"><format dxfId="\d+"><pivotArea/);
    files['xl/pivotTables/pivotTable1.xml'] = strToU8(table.replace(/<extLst>.*<\/extLst>/, ''));
    const back = readXlsx(zipSync(files));
    const s2 = new GridStore(back);
    refreshPivots(s2);
    expect(s2.style(1, 7).fill?.color?.toLowerCase()).toBe('#ffcc00');
    expect(s2.style(2, 7).fill?.color?.toLowerCase()).not.toBe('#ffcc00');
    expect(s2.value(1, 6)).toBe('West');
    void res;
  });

  it('report filter pages make one sheet per item', () => {
    const s = setupBook();
    createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: null, build: (pt, c) => placeField(placeField(placeField(pt, c, 0, 'pages'), c, 1, 'rows'), c, 3, 'data') }, newPivotTable);
    const at = s.sheetIndex;
    const ref = { sheet: at, index: 0, pt: s.wb.sheets[at]!.pivots![0]! };
    setPageItem(s, ref, 0, undefined, new Set([1]));
    expect(s.value(0, 1, at)).toBe('(Multiple Items)');
    const n = showFilterPages(s, { sheet: at, index: 0, pt: s.wb.sheets[at]!.pivots![0]! }, 0);
    expect(n).toBe(2);
    expect(s.wb.sheets.slice(at, at + 2).map((x) => x.name)).toEqual(['East', 'West']);
    expect(s.value(0, 1, at)).toBe('East');
    expect(s.value(5, 1, at)).toBe(60);
  });

  it('follows inserted rows and renamed sheets, and refreshes on open when asked', () => {
    const s = setupBook();
    const res = createPivot(s, { source: { kind: 'range', sheet: 'Sheet1', range: { r1: 0, c1: 0, r2: 7, c2: 3 } }, dest: { sheet: 0, row: 0, col: 6 }, build: (pt, c) => placeField(placeField(pt, c, 0, 'rows'), c, 3, 'data') }, newPivotTable) as { sheet: number; name: string };
    changeStructure(s, 'row', 0, 2);
    const pt = s.wb.sheets[0]!.pivots![0]!;
    expect(pt.anchor.row).toBe(2);
    expect(s.wb.pivotCaches![0]!.source).toMatchObject({ range: { r1: 2, r2: 9 } });
    expect(refreshPivots(s)).toBeNull();
    expect(s.value(6, 7)).toBe(112);
    renameSheet(s, 0, 'Data');
    expect(refreshPivots(s)).toBeNull();
    // Saved with refresh-on-open; a change to the data shows when it opens again.
    s.change('flag', 'book', () => { s.wb.pivotCaches = s.wb.pivotCaches!.map((c) => ({ ...c, refreshOnLoad: true })); });
    s.change('edit', 'sheet', () => s.engine.setValue(0, 9, 3, 1007));
    const back = readXlsx(writeXlsx(s.wb));
    const s2 = new GridStore(back);
    expect(s2.value(6, 7)).toBe(112);
    refreshOnOpen(s2);
    expect(s2.value(6, 7)).toBe(1112);
    expect(s2.file.dirty).toBe(false);
    void res;
  });
});
