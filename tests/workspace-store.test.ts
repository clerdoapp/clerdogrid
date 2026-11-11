import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { addSheet, deleteSheet, moveSheet, renameSheet, setSheetState, duplicateSheet } from '../src/app/ops/structure';
import { CellError } from '../src/engine/types';
import { newSheet, parseAddr } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { panesOf, paneFor, type Viewport } from '../src/ui/render';
import { layoutOf } from '../src/app/layout';
import { roundToDisplay, applyWorkbookSettings } from '../src/features/workspace/options';
import { issuesAt, convertToNumbers, ignoreIssue, extendRange, copyNeighbourFormula } from '../src/features/workspace/errors';
import { completionFor, installAutoComplete } from '../src/features/workspace/autocomplete';
import { contentHash } from '../src/features/workspace/versions';
import { activeViewOf, addCustomView, captureView, createView, settleViewsForSave, showCustomView, switchView } from '../src/features/workspace/views';
import { blocked } from '../src/features/workspace/protect';

beforeAll(() => registerAllFunctions());

const setup = () => {
  const s = new GridStore();
  const at = (a: string) => parseAddr(a)!;
  const type = (a: string, text: string) => { const p = at(a); s.select(p); s.beginEdit(text); expect(s.commitEdit()).toBe(true); };
  const val = (a: string) => { const p = at(a); const v = s.value(p.row, p.col); return v instanceof CellError ? v.code : v; };
  return { s, type, val, at };
};

describe('calculation options', () => {
  it('manual calculation leaves dependents stale until a full recalculation', () => {
    const { s, type, val } = setup();
    type('A1', '1');
    type('B1', '=A1*2');
    s.wb.calc = { manual: true };
    applyWorkbookSettings(s);
    type('A1', '5');
    expect(val('B1')).toBe(2);
    expect(s.engine.stale).toBe(true);
    type('C1', '=A1+1');
    expect(val('C1')).toBe(6);
    s.engine.recalc(true);
    expect(val('B1')).toBe(10);
    expect(s.engine.stale).toBe(false);
  });

  it('precision as displayed rounds formula results to their format', () => {
    const { s, type, val, at } = setup();
    type('A1', '=1/3');
    const p = at('A1');
    s.styleRange({ r1: p.row, c1: p.col, r2: p.row, c2: p.col }, (st) => ({ ...st, numFmt: '0.00' }));
    s.wb.calc = { precisionAsDisplayed: true };
    applyWorkbookSettings(s);
    s.engine.recalc(true);
    expect(val('A1')).toBe(0.33);
    expect(roundToDisplay(1234.5678, '#,##0.0')).toBe(1234.6);
    expect(roundToDisplay(-0.12345, '0.0%')).toBeCloseTo(-0.123, 10);
    expect(roundToDisplay(0.5, 'm/d/yyyy')).toBe(0.5);
    expect(roundToDisplay(2.345, 'General')).toBe(2.345);
  });

  it('R1C1 reference style shows and accepts R1C1 formulas', () => {
    const { s, type, at } = setup();
    s.r1c1 = true;
    type('C3', '=R[-1]C[-1]+R1C1');
    const p = at('C3');
    expect(s.cell(p.row, p.col)?.f).toBe('B2+$A$1');
    expect(s.inputText(p.row, p.col)).toBe('=R[-1]C[-1]+R1C1');
  });

  it('moves the selection after Enter in the chosen direction', () => {
    const { s } = setup();
    s.select({ row: 2, col: 2 });
    s.enterDirection = 'right';
    s.moveAfterEnter(false);
    expect(s.sel.active).toEqual({ row: 2, col: 3 });
    s.moveAfterEnter(true);
    expect(s.sel.active).toEqual({ row: 2, col: 2 });
    s.enterDirection = 'none';
    s.moveAfterEnter(false);
    expect(s.sel.active).toEqual({ row: 2, col: 2 });
  });
});

describe('workbook structure protection', () => {
  it('blocks sheet structure changes while protected', () => {
    const { s } = setup();
    s.wb.sheets.push(newSheet('Two'));
    s.engine.sheetsChanged();
    s.wb.protection = { structure: true };
    addSheet(s);
    duplicateSheet(s);
    deleteSheet(s, 1);
    expect(renameSheet(s, 1, 'Other')).toBe(false);
    moveSheet(s, 0, 1);
    setSheetState(s, 1, 'hidden');
    expect(s.wb.sheets.map((x) => `${x.name}:${x.state}`)).toEqual(['Sheet1:visible', 'Two:visible']);
    delete s.wb.protection;
    addSheet(s);
    expect(s.wb.sheets.length).toBe(3);
  });

  it('reads sheet protection flags with the file format defaults', () => {
    const sh = newSheet('S');
    expect(blocked(sh, 'formatCells')).toBe(false);
    sh.protection = { sheet: true };
    expect(blocked(sh, 'formatCells')).toBe(true);
    expect(blocked(sh, 'selectLockedCells')).toBe(false);
    sh.protection.options = { formatCells: false, selectLockedCells: true };
    expect(blocked(sh, 'formatCells')).toBe(false);
    expect(blocked(sh, 'selectLockedCells')).toBe(true);
  });
});

describe('sheet views', () => {
  it('keeps filters and hidden rows per view and saves the default view in the sheet data', () => {
    const { s, type } = setup();
    ['Region', 'North', 'South', 'North', 'East'].forEach((v, i) => type(`A${i + 1}`, v));
    expect(activeViewOf(s.sheet)).toBeNull();
    const v = createView(s, 'North only');
    s.sheet.autoFilter = { range: { r1: 0, c1: 0, r2: 4, c2: 0 }, columns: [{ col: 0, values: ['North'] }] };
    switchView(s, null);
    expect(s.sheet.autoFilter).toBeUndefined();
    expect([...s.sheet.rows.values()].some((r) => r.hidden)).toBe(false);
    switchView(s, v.id!);
    expect(s.sheet.rows.get(2)?.hidden).toBe(true);
    expect(s.sheet.rows.get(1)?.hidden).toBeFalsy();
    s.sheet.cols.set(3, { hidden: true });
    // Saving writes the default view's state; the view keeps its own.
    const undo = settleViewsForSave(s);
    const back = readXlsx(writeXlsx(s.wb));
    undo();
    expect(back.sheets[0]!.autoFilter).toBeUndefined();
    expect(back.sheets[0]!.cols.get(3)?.hidden).toBeFalsy();
    const saved = back.sheets[0]!.sheetViews![0]!;
    expect(saved.name).toBe('North only');
    expect(saved.hiddenCols).toEqual([3]);
    expect(saved.filter?.columns[0]?.values).toEqual(['North']);
    expect(s.sheet.cols.get(3)?.hidden).toBe(true);
    expect(captureView(s.sheet).hiddenCols).toEqual([3]);
  });

  it('custom views restore zoom, hidden rows and sheet visibility', () => {
    const { s } = setup();
    s.wb.sheets.push(newSheet('Two'));
    s.engine.sheetsChanged();
    s.sheet.view.zoom = 150;
    s.sheet.rows.set(4, { hidden: true });
    addCustomView(s, 'Zoomed', true, true);
    s.sheet.view.zoom = 100;
    s.sheet.rows.delete(4);
    s.wb.sheets[1]!.state = 'hidden';
    showCustomView(s, s.wb.customViews![0]!.guid);
    expect(s.sheet.view.zoom).toBe(150);
    expect(s.sheet.rows.get(4)?.hidden).toBe(true);
    expect(s.wb.sheets[1]!.state).toBe('visible');
    s.undo();
    expect(s.sheet.view.zoom).toBe(100);
  });
});

describe('error checking', () => {
  it('finds numbers stored as text and converts them', () => {
    const { s, type, at } = setup();
    type('A1', "'123");
    const p = at('A1');
    expect(issuesAt(s, p.row, p.col).map((i) => i.rule)).toEqual(['numberStoredAsText']);
    s.select(p);
    convertToNumbers(s);
    expect(s.value(p.row, p.col)).toBe(123);
    expect(issuesAt(s, p.row, p.col)).toEqual([]);
  });

  it('flags inconsistent formulas and copies the neighbouring one', () => {
    const { s, type, at, val } = setup();
    for (let r = 1; r <= 3; r++) { type(`A${r}`, String(r)); type(`B${r}`, String(r * 10)); }
    type('C1', '=A1+B1'); type('C2', '=A2*B2'); type('C3', '=A3+B3');
    const p = at('C2');
    expect(issuesAt(s, p.row, p.col).map((i) => i.rule)).toContain('formula');
    s.select(p);
    copyNeighbourFormula(s);
    expect(val('C2')).toBe(22);
    expect(issuesAt(s, p.row, p.col)).toEqual([]);
  });

  it('flags formulas that leave out adjacent numbers and extends them', () => {
    const { s, type, at } = setup();
    ['1', '2', '3', '4'].forEach((v, i) => type(`A${i + 1}`, v));
    type('A5', '=SUM(A1:A3)');
    const p = at('A5');
    expect(issuesAt(s, p.row, p.col).map((i) => i.rule)).toContain('formulaRange');
    s.select(p);
    extendRange(s);
    expect(s.cell(p.row, p.col)?.f).toBe('SUM(A1:A4)');
  });

  it('ignores an error for a cell and keeps that in the file', () => {
    const { s, type, at } = setup();
    type('B2', "'42");
    const p = at('B2');
    s.select(p);
    ignoreIssue(s, 'numberStoredAsText');
    expect(issuesAt(s, p.row, p.col)).toEqual([]);
    const back = readXlsx(writeXlsx(s.wb));
    expect(back.sheets[0]!.ignoredErrors).toEqual([{ range: { r1: 1, c1: 1, r2: 1, c2: 1 }, types: ['numberStoredAsText'] }]);
  });

  it('flags error results, unlocked formulas and two-digit years', () => {
    const { s, type, at } = setup();
    type('A1', '=1/0');
    type('A2', '1/2/25');
    expect(issuesAt(s, 0, 0).map((i) => i.rule)).toContain('evalError');
    s.engine.setValue(0, 2, 0, '3/4/99');
    expect(issuesAt(s, 2, 0).map((i) => i.rule)).toEqual(['twoDigitTextYear']);
    const p = at('A1');
    s.styleRange({ r1: p.row, c1: p.col, r2: p.row, c2: p.col }, (st) => ({ ...st, protection: { locked: false } }));
    expect(issuesAt(s, p.row, p.col).map((i) => i.rule)).toContain('unlockedFormula');
  });
});

describe('autocomplete', () => {
  it('fills in a unique match from the column, which typing replaces and Enter accepts', () => {
    const { s, type, val } = setup();
    installAutoComplete(s);
    type('A1', 'North');
    type('A2', 'South');
    s.select({ row: 2, col: 0 });
    s.beginEdit('N');
    expect(s.editing?.text).toBe('North');
    s.setEditText('No');
    expect(s.editing?.text).toBe('North');
    s.setEditText('No');
    expect(s.editing?.text).toBe('No');
    expect(completionFor(s)).toBe('North');
    s.setEditText('Nor');
    expect(s.editing?.text).toBe('North');
    expect(s.commitEdit()).toBe(true);
    expect(val('A3')).toBe('North');
    s.select({ row: 3, col: 0 });
    s.beginEdit('Z');
    expect(s.editing?.text).toBe('Z');
    s.cancelEdit();
  });
});

describe('version history', () => {
  it('hashes contents without save times so unchanged work is not stored twice', () => {
    const { s, type } = setup();
    type('A1', 'x');
    const a = writeXlsx({ ...s.wb, props: { modified: '2020-01-01T00:00:00Z' } });
    const b = writeXlsx({ ...s.wb, props: { modified: '2026-01-01T00:00:00Z' } });
    expect(contentHash(a)).toBe(contentHash(b));
    type('A1', 'y');
    expect(contentHash(writeXlsx(s.wb))).not.toBe(contentHash(a));
  });
});

describe('split panes', () => {
  it('lays out four independently scrolled panes and finds the pane showing a cell', () => {
    const s = newSheet('S');
    const layout = layoutOf(s, 0, [{}]);
    const vp: Viewport = { layout, zoom: 1, headerW: 40, headerH: 22, width: 800, height: 600, scrollX: 640, scrollY: 400, frozenRows: 0, frozenCols: 0, split: { w: 200, h: 100, sx: 0, sy: 0 } };
    const panes = panesOf(vp);
    expect(panes).toHaveLength(4);
    expect(panes[0]!.scrollX).toBe(640);
    expect(panes.find((p) => p.y === 22 && p.x === 40)!.scrollY).toBe(0);
    expect(paneFor(panes, vp, 0, 0)).toBe(panes[3]);
    expect(paneFor(panes, vp, 25, 12)).toBe(panes[0]);
  });
});
