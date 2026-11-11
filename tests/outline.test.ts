import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import type { Scalar } from '../src/engine/types';
import { readXlsx } from '../src/formats/xlsx-read';
import { writeXlsx } from '../src/formats/xlsx-write';
import { copySelection, paste } from '../src/app/ops/clipboard';
import {
  applySubtotals, autoOutline, clearOutline, consolidate, groupLines, maxLevel, outlineGroups, removeSubtotals, selectVisible, setGroupCollapsed, showLevel,
} from '../src/app/ops/outline';

beforeAll(() => registerAllFunctions());

const fill = (s: GridStore, rows: Scalar[][], r0 = 0, c0 = 0, sheet = 0) => s.change('data', 'book', () => rows.forEach((r, i) => r.forEach((v, j) => {
  if (typeof v === 'string' && v.startsWith('=')) s.engine.setFormula(sheet, r0 + i, c0 + j, v.slice(1)); else s.engine.setValue(sheet, r0 + i, c0 + j, v);
})));
const col = (s: GridStore, c: number, r1: number, r2: number) => Array.from({ length: r2 - r1 + 1 }, (_, i) => s.value(r1 + i, c));
const hidden = (s: GridStore) => [...s.sheet.rows].filter(([, i]) => i.hidden).map(([r]) => r).sort((a, b) => a - b);

describe('outline', () => {
  it('groups to 8 levels, collapses and expands with nested groups kept', () => {
    const s = new GridStore();
    groupLines(s, 'row', 1, 6, 1);
    groupLines(s, 'row', 2, 3, 1);
    expect(maxLevel(s.sheet, 'row')).toBe(2);
    const outer = outlineGroups(s.sheet, 'row').find((g) => g.level === 1)!;
    const inner = outlineGroups(s.sheet, 'row').find((g) => g.level === 2)!;
    expect([outer.start, outer.end, outer.summary, inner.start, inner.end, inner.summary]).toEqual([1, 6, 7, 2, 3, 4]);
    setGroupCollapsed(s, 'row', inner, true);
    expect(hidden(s)).toEqual([2, 3]);
    expect(s.sheet.rows.get(4)!.collapsed).toBe(true);
    setGroupCollapsed(s, 'row', outlineGroups(s.sheet, 'row').find((g) => g.level === 1)!, true);
    expect(hidden(s)).toEqual([1, 2, 3, 4, 5, 6]);
    setGroupCollapsed(s, 'row', outlineGroups(s.sheet, 'row').find((g) => g.level === 1)!, false);
    expect(hidden(s)).toEqual([2, 3]);
    showLevel(s, 'row', 3);
    expect(hidden(s)).toEqual([]);
    showLevel(s, 'row', 1);
    expect(hidden(s)).toEqual([1, 2, 3, 4, 5, 6]);
    s.undo();
    expect(hidden(s)).toEqual([]);
    for (let i = 0; i < 5; i++) groupLines(s, 'row', 2, 2, 1);
    expect(maxLevel(s.sheet, 'row')).toBe(7);
    groupLines(s, 'row', 2, 2, 1);
    expect(maxLevel(s.sheet, 'row')).toBe(7);
    clearOutline(s);
    expect(maxLevel(s.sheet, 'row')).toBe(0);
  });

  it('builds an outline from formulas', () => {
    const s = new GridStore();
    fill(s, [['A'], [1], [2], ['=SUM(A2:A3)'], [4], [5], ['=SUM(A5:A6)'], ['=A4+A7']]);
    s.selectRange({ r1: 0, c1: 0, r2: 7, c2: 0 });
    autoOutline(s);
    expect([1, 2, 3, 4, 5, 6, 7].map((r) => s.sheet.rows.get(r)?.outline ?? 0)).toEqual([2, 2, 1, 2, 2, 1, 0]);
  });

  it('adds subtotals, nests them, keeps formulas right and removes them', () => {
    const s = new GridStore();
    fill(s, [['Region', 'Item', 'Sales'], ['East', 'a', 1], ['East', 'b', 2], ['East', 'b', 3], ['West', 'a', 4], ['West', 'a', 5], ['After', null, null]]);
    // The row after the list moves down when rows are inserted.
    s.change('x', 'sheet', () => s.engine.setValue(0, 8, 0, 'below'));
    const range = { r1: 0, c1: 0, r2: 5, c2: 2 };
    applySubtotals(s, range, { groupBy: 0, fn: 9, columns: [2], replace: true, pageBreaks: false, summaryBelow: true });
    expect(col(s, 0, 1, 8)).toEqual(['East', 'East', 'East', 'East Total', 'West', 'West', 'West Total', 'Grand Total']);
    expect(col(s, 2, 1, 8)).toEqual([1, 2, 3, 6, 4, 5, 9, 15]);
    expect(s.cell(4, 2)!.f).toBe('SUBTOTAL(9,C2:C4)');
    expect([1, 4, 7, 8].map((r) => s.sheet.rows.get(r)?.outline ?? 0)).toEqual([2, 1, 1, 0]);
    expect(s.value(11, 0)).toBe('below');
    // Nest a count by item inside each region.
    applySubtotals(s, { r1: 0, c1: 0, r2: 8, c2: 2 }, { groupBy: 1, fn: 3, columns: [2], replace: false, pageBreaks: false, summaryBelow: true });
    expect(col(s, 1, 1, 12)).toEqual(['a', 'a Count', 'b', 'b', 'b Count', null, 'a', 'a', 'a Count', null, null, null]);
    expect(col(s, 0, 1, 13)).toEqual(['East', null, 'East', 'East', null, 'East Total', 'West', 'West', null, 'West Total', 'Grand Total', 'After', null]);
    expect(col(s, 2, 1, 11)).toEqual([1, 1, 2, 3, 2, 6, 4, 5, 2, 9, 15]);
    expect(maxLevel(s.sheet, 'row')).toBe(3);
    removeSubtotals(s, { r1: 0, c1: 0, r2: 11, c2: 2 });
    expect(col(s, 0, 1, 6)).toEqual(['East', 'East', 'East', 'West', 'West', 'After']);
    expect(maxLevel(s.sheet, 'row')).toBe(0);
    expect(s.value(8, 0)).toBe('below');
  });

  it('subtotals above the data', () => {
    const s = new GridStore();
    fill(s, [['K', 'V'], ['x', 1], ['x', 2], ['y', 3]]);
    applySubtotals(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, { groupBy: 0, fn: 9, columns: [1], replace: true, pageBreaks: true, summaryBelow: false });
    expect(col(s, 0, 1, 6)).toEqual(['Grand Total', 'x Total', 'x', 'x', 'y Total', 'y']);
    expect(col(s, 1, 1, 6)).toEqual([6, 3, 1, 2, 3, 3]);
    expect(s.sheet.outlinePr?.summaryBelow).toBe(false);
  });

  it('selects visible cells only', () => {
    const s = new GridStore();
    s.change('h', 'sheet', () => { s.sheet.rows.set(2, { hidden: true }); s.sheet.cols.set(1, { hidden: true }); });
    s.selectRange({ r1: 0, c1: 0, r2: 4, c2: 2 });
    selectVisible(s);
    expect(s.sel.ranges.map((r) => [r.r1, r.r2, r.c1, r.c2]).sort()).toEqual([[0, 1, 0, 0], [0, 1, 2, 2], [3, 4, 0, 0], [3, 4, 2, 2]].sort());
  });

  it('copies only the visible cells', async () => {
    const s = new GridStore();
    fill(s, [[1, 'x', 10], [2, 'x', '=A2*10'], [3, 'x', 30], [4, 'x', '=A4*10']]);
    s.change('h', 'sheet', () => { s.sheet.rows.set(1, { hidden: true }); s.sheet.cols.set(1, { hidden: true }); });
    s.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 });
    selectVisible(s);
    await copySelection(s);
    s.selectRange({ r1: 10, c1: 5, r2: 10, c2: 5 });
    await paste(s);
    const out = [10, 11, 12].map((r) => [s.value(r, 5), s.value(r, 6)]);
    expect(out).toEqual([[1, 10], [3, 30], [4, 0]]);
    // Formulas move by the distance from their own source cell (C4 to G13).
    expect(s.cell(12, 6)!.f).toBe('E13*10');
  });

  it('consolidates by position, by labels and with links', () => {
    const s = new GridStore();
    s.change('sheets', 'book', () => { s.wb.sheets.push({ ...structuredClone(s.wb.sheets[0]!), name: 'B' }); s.engine.sheetsChanged(); });
    fill(s, [['', 'Q1', 'Q2'], ['East', 1, 2], ['West', 3, 4]], 0, 0, 0);
    fill(s, [['', 'Q2', 'Q3'], ['West', 10, 20], ['North', 5, 6]], 0, 0, 1);
    s.selectRange({ r1: 10, c1: 0, r2: 10, c2: 0 });
    expect(consolidate(s, { fn: 'sum', refs: [{ sheet: 0, range: { r1: 0, c1: 0, r2: 2, c2: 2 } }, { sheet: 1, range: { r1: 0, c1: 0, r2: 2, c2: 2 } }], topRow: true, leftCol: true, links: false })).toBeNull();
    const g = (r1: number, r2: number) => Array.from({ length: r2 - r1 + 1 }, (_, i) => [0, 1, 2, 3].map((c) => s.value(r1 + i, c)));
    expect(g(10, 13)).toEqual([[null, 'Q1', 'Q2', 'Q3'], ['East', 1, 2, null], ['West', 3, 14, 20], ['North', null, 5, 6]]);
    s.selectRange({ r1: 20, c1: 0, r2: 20, c2: 0 });
    consolidate(s, { fn: 'max', refs: [{ sheet: 0, range: { r1: 1, c1: 1, r2: 2, c2: 2 } }, { sheet: 1, range: { r1: 1, c1: 1, r2: 2, c2: 2 } }], topRow: false, leftCol: false, links: false });
    expect(g(20, 21).map((r) => r.slice(0, 2))).toEqual([[10, 20], [5, 6]]);
    s.selectRange({ r1: 30, c1: 0, r2: 30, c2: 0 });
    consolidate(s, { fn: 'sum', refs: [{ sheet: 0, range: { r1: 0, c1: 0, r2: 2, c2: 2 } }, { sheet: 1, range: { r1: 0, c1: 0, r2: 2, c2: 2 } }], topRow: true, leftCol: true, links: true });
    // West has detail rows from both sheets, then the summary.
    expect(g(33, 35)).toEqual([['Sheet1', 3, 4, null], ['B', null, 10, 20], ['West', 3, 14, 20]]);
    expect(s.cell(34, 2)!.f).toBe('B!$B$2');
    expect(s.sheet.rows.get(33)).toMatchObject({ outline: 1, hidden: true });
    expect(s.sheet.rows.get(35)?.collapsed).toBe(true);
  });

  it('saves and reopens outline state', () => {
    const s = new GridStore();
    groupLines(s, 'row', 1, 3, 1);
    groupLines(s, 'col', 2, 4, 1);
    setGroupCollapsed(s, 'row', outlineGroups(s.sheet, 'row')[0]!, true);
    s.change('pr', 'sheet', () => { s.sheet.outlinePr = { summaryBelow: true, summaryRight: false }; });
    const back = readXlsx(writeXlsx(s.wb)).sheets[0]!;
    expect(back.rows.get(1)).toMatchObject({ outline: 1, hidden: true });
    expect(back.rows.get(4)?.collapsed).toBe(true);
    expect(back.cols.get(3)?.outline).toBe(1);
    expect(back.outlinePr).toEqual({ summaryBelow: true, summaryRight: false });
  });
});
