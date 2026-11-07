import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { changeStructure, addSheet, renameSheet, deleteSheet } from '../src/app/ops/structure';
import { autoFill, fillDirection, flashFill } from '../src/app/ops/fill';
import { sortRange, removeDuplicates, findAll, goalSeek } from '../src/app/ops/data';
import { moveRange } from '../src/app/ops/move';
import { CellError } from '../src/engine/types';
import { parseAddr } from '../src/model/workbook';

beforeAll(() => registerAllFunctions());

const setup = () => {
  const s = new GridStore();
  const type = (a: string, text: string) => { const p = parseAddr(a)!; s.select({ row: p.row, col: p.col }); s.beginEdit(text); expect(s.commitEdit()).toBe(true); };
  const val = (a: string) => { const p = parseAddr(a)!; const v = s.value(p.row, p.col); return v instanceof CellError ? v.code : v; };
  const f = (a: string) => { const p = parseAddr(a)!; return s.cell(p.row, p.col)?.f; };
  const fmt = (a: string) => { const p = parseAddr(a)!; return s.style(p.row, p.col).numFmt; };
  return { s, type, val, f, fmt };
};

describe('typing', () => {
  it('reads numbers, dates, percents, booleans and formulas like a spreadsheet', () => {
    const { type, val, fmt, f } = setup();
    type('A1', '1,234.5'); type('A2', '15%'); type('A3', '3/15/2024'); type('A4', 'true'); type('A5', "'007"); type('A6', '=sum(a1:a2'); type('A7', '$12.50');
    expect(val('A1')).toBe(1234.5); expect(fmt('A1')).toBe('#,##0.00');
    expect(val('A2')).toBe(0.15); expect(fmt('A2')).toBe('0%');
    expect(val('A3')).toBe(45366); expect(fmt('A3')).toMatch(/yyyy/);
    expect(val('A4')).toBe(true);
    expect(val('A5')).toBe('007');
    expect(f('A6')).toBe('SUM(A1:A2)'); expect(val('A6')).toBeCloseTo(1234.65);
    expect(val('A7')).toBe(12.5);
  });

  it('refuses broken formulas and keeps the editor open', () => {
    const { s } = setup();
    s.beginEdit('=1+*2');
    expect(s.commitEdit()).toBe(false);
    expect(s.editing).not.toBeNull();
  });

  it('undoes and redoes edits', () => {
    const { s, type, val } = setup();
    type('B1', '5'); type('B2', '=B1*2');
    expect(val('B2')).toBe(10);
    type('B1', '7');
    expect(val('B2')).toBe(14);
    s.undo();
    expect(val('B2')).toBe(10);
    s.redo();
    expect(val('B2')).toBe(14);
    s.undo(); s.undo();
    expect(val('B2')).toBe(null);
  });

  it('fills the whole selection with Ctrl+Enter, shifting relative references', () => {
    const { s, type, val } = setup();
    type('A1', '1'); type('A2', '2'); type('A3', '3');
    s.selectRange({ r1: 0, c1: 1, r2: 2, c2: 1 });
    s.beginEdit('=A1*10');
    s.commitEdit(true);
    expect([val('B1'), val('B2'), val('B3')]).toEqual([10, 20, 30]);
  });
});

describe('structure', () => {
  it('moves references when rows and columns are inserted or deleted', () => {
    const { s, type, f, val } = setup();
    type('A1', '1'); type('A2', '2'); type('A3', '3'); type('C1', '=SUM(A1:A3)'); type('C2', '=A3');
    changeStructure(s, 'row', 1, 2);
    expect(f('C1')).toBe('SUM(A1:A5)');
    expect(f('C4')).toBe('A5');
    expect(val('C1')).toBe(6);
    changeStructure(s, 'row', 4, -1);
    expect(f('C4')).toBe('#REF!');
    changeStructure(s, 'col', 0, 1);
    expect(f('D1')).toBe('SUM(B1:B4)');
    s.undo(); s.undo(); s.undo();
    expect(f('C1')).toBe('SUM(A1:A3)');
  });

  it('renames and deletes sheets, fixing formulas', () => {
    const { s, type, f } = setup();
    addSheet(s);
    s.setActiveSheet(0);
    type('A1', '=Sheet2!B2+1');
    renameSheet(s, 1, 'Q1 Data');
    expect(f('A1')).toBe("'Q1 Data'!B2+1");
    deleteSheet(s, 1);
    expect(f('A1')).toBe('#REF!+1');
  });
});

describe('fill', () => {
  it('continues series and shifts formulas', () => {
    const { s, type, val, f } = setup();
    type('A1', '2'); type('A2', '4');
    type('B1', 'Jan'); type('B2', 'Feb');
    type('C1', 'Item 1');
    type('D1', '=A1*2'); type('D2', '=A2*2');
    autoFill(s, { r1: 0, c1: 0, r2: 1, c2: 3 }, { r1: 0, c1: 0, r2: 4, c2: 3 });
    expect([val('A3'), val('A5')]).toEqual([6, 10]);
    expect([val('B3'), val('B5')]).toEqual(['Mar', 'May']);
    expect(f('D5')).toBe('A5*2');
    autoFill(s, { r1: 0, c1: 2, r2: 0, c2: 2 }, { r1: 0, c1: 2, r2: 4, c2: 2 });
    expect(val('C5')).toBe('Item 5');
    s.selectRange({ r1: 0, c1: 4, r2: 3, c2: 4 });
    type('E1', '9');
    s.selectRange({ r1: 0, c1: 4, r2: 3, c2: 4 });
    fillDirection(s, 'down');
    expect(val('E4')).toBe(9);
  });

  it('learns a Flash Fill pattern from one example', () => {
    const { s, type, val } = setup();
    type('A1', 'Ada Lovelace'); type('A2', 'Alan Turing'); type('A3', 'Grace Hopper');
    type('B1', 'Lovelace, Ada');
    s.select({ row: 0, col: 1 });
    expect(flashFill(s)).toBe(true);
    expect([val('B2'), val('B3')]).toEqual(['Turing, Alan', 'Hopper, Grace']);
  });
});

describe('data tools', () => {
  it('sorts with headers, blanks last, formulas following their rows', () => {
    const { s, type, val } = setup();
    type('A1', 'Name'); type('B1', 'Score');
    type('A2', 'Cy'); type('B2', '70');
    type('A3', 'Al'); type('B3', '95');
    type('A4', 'Bo');
    type('C2', '=B2*2'); type('C3', '=B3*2'); type('C4', '=B4*2');
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 2 }, [{ col: 1, descending: true }], true);
    expect([val('A2'), val('A3'), val('A4')]).toEqual(['Al', 'Cy', 'Bo']);
    expect([val('C2'), val('C3'), val('C4')]).toEqual([190, 140, 0]);
  });

  it('removes duplicates, finds text, and seeks goals', () => {
    const { s, type, val } = setup();
    ['x', 'y', 'X', 'z', 'y'].forEach((t, i) => type(`A${i + 1}`, t));
    expect(removeDuplicates(s, { r1: 0, c1: 0, r2: 4, c2: 0 }, [0], false)).toBe(2);
    expect([val('A1'), val('A2'), val('A3'), val('A4')]).toEqual(['x', 'y', 'z', null]);
    expect(findAll(s, { text: 'y*' }).length).toBe(1);
    type('D1', '3'); type('D2', '=D1^2+1');
    const r = goalSeek(s, { row: 1, col: 3 }, 50, { row: 0, col: 3 });
    expect(r.ok).toBe(true);
    expect(val('D1') as number).toBeCloseTo(7, 5);
  });

  it('moves a block and keeps references pointing at it', () => {
    const { s, type, f, val } = setup();
    type('A1', '5'); type('B1', '=A1+1');
    moveRange(s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { row: 4, col: 2 });
    expect(val('C5')).toBe(5);
    expect(f('B1')).toBe('C5+1');
    expect(val('B1')).toBe(6);
  });
});
