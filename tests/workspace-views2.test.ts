import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { newSheet, parseAddr } from '../src/model/workbook';
import { sortRange, sortInterceptor } from '../src/app/ops/data';
import { layoutOf } from '../src/app/layout';
import { RowOrder, withRowOrder } from '../src/app/rowOrder';
import { gotoReference } from '../src/app/goto';
import { panesOf, rectIn, type Viewport } from '../src/ui/render';
import { toScreen, toSheet } from '../src/features/objects/state';
import { createView, sortInView, switchView, carryViews } from '../src/features/workspace/views';
import { blocked } from '../src/features/workspace/protect';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';

beforeAll(() => { registerAllFunctions(); sortInterceptor.current = sortInView; });

const setup = () => {
  const s = new GridStore();
  const type = (a: string, text: string) => { const p = parseAddr(a)!; s.select(p); s.beginEdit(text); expect(s.commitEdit()).toBe(true); };
  return { s, type };
};

describe('row order', () => {
  it('maps display and data rows and splits display ranges into data runs', () => {
    const o = new RowOrder(1, Int32Array.from([3, 1, 2]));
    expect([0, 1, 2, 3, 4].map((p) => o.toData(p))).toEqual([0, 3, 1, 2, 4]);
    expect([0, 1, 2, 3, 4].map((d) => o.toDisplay(d))).toEqual([0, 2, 3, 1, 4]);
    expect(o.dataRuns(0, 4, 0, 1)).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 1 }, { r1: 3, c1: 0, r2: 3, c2: 1 }, { r1: 1, c1: 0, r2: 2, c2: 1 }, { r1: 4, c1: 0, r2: 4, c2: 1 }]);
    expect(o.displaySpan(1, 2)).toEqual([2, 3]);
    expect(o.displaySpan(0, 9)).toEqual([0, 9]);
  });

  it('lays rows out in display order', () => {
    const sh = newSheet('S');
    sh.rows.set(3, { height: 30 });
    const base = layoutOf(sh, 0, [{}]);
    const l = withRowOrder(base, new RowOrder(1, Int32Array.from([3, 1, 2])));
    expect(l.rows.size(1)).toBe(base.rows.size(3));
    expect(l.rows.size(2)).toBe(base.rows.size(1));
    expect(l.rows.start(4)).toBe(base.rows.start(4));
    expect(l.rows.at(l.rows.start(2) + 1)).toBe(2);
    expect(l.rows.at(base.rows.start(10) + 1)).toBe(10);
  });
});

describe('sorting inside a sheet view', () => {
  const data = () => {
    const t = setup();
    [['Name', 'Score'], ['Cy', '3'], ['Al', '1'], ['Bo', '2']].forEach(([a, b], i) => { t.type(`A${i + 1}`, a!); t.type(`B${i + 1}`, b!); });
    t.type('C2', '=B2*10');
    return t;
  };

  it('reorders only the view; the data, formulas and the default view keep their order', () => {
    const { s } = data();
    createView(s, 'By score');
    s.select({ row: 1, col: 1 });
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1 }], true);
    expect([1, 2, 3].map((r) => s.value(r, 0))).toEqual(['Cy', 'Al', 'Bo']);
    expect(s.value(1, 2)).toBe(30);
    const o = s.rowOrderOf()!;
    expect([1, 2, 3].map((p) => s.value(o.toData(p), 0))).toEqual(['Al', 'Bo', 'Cy']);
    // Moving down walks display order.
    s.select({ row: 2, col: 0 });
    s.move(1, 0);
    expect(s.value(s.sel.active.row, 0)).toBe('Bo');
    s.move(1, 0);
    expect(s.value(s.sel.active.row, 0)).toBe('Cy');
    // A Shift selection over display rows becomes runs of data rows.
    s.select({ row: 2, col: 0 });
    s.select({ row: 1, col: 1 }, true);
    expect(s.sel.ranges.flatMap((r) => { const out: string[] = []; for (let i = r.r1; i <= r.r2; i++) out.push(String(s.value(i, 0))); return out; }).sort()).toEqual(['Al', 'Bo', 'Cy']);
    // Ctrl+Down from the first display row goes to the last display row of the block.
    const end = s.edge({ row: 2, col: 0 }, 1, 0);
    expect(s.value(end.row, 0)).toBe('Cy');
    switchView(s, null);
    expect(s.rowOrderOf()).toBeNull();
    expect([1, 2, 3].map((r) => s.value(r, 0))).toEqual(['Cy', 'Al', 'Bo']);
  });

  it('keeps the view sort in the file and shows it again when the view opens', () => {
    const { s } = data();
    const v = createView(s, 'Sorted');
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 0, descending: true }], true);
    switchView(s, null);
    const back = new GridStore(readXlsx(writeXlsx(s.wb)));
    back.load(readXlsx(writeXlsx(s.wb)), 'x');
    expect(back.sheet.sheetViews![0]!.filter?.sort).toEqual([{ col: 0, descending: true }]);
    switchView(back, back.sheet.sheetViews![0]!.id!);
    const o = back.rowOrderOf()!;
    expect([1, 2, 3].map((p) => back.value(o.toData(p), 0))).toEqual(['Cy', 'Bo', 'Al']);
    void v;
  });

  it('sorts the data when no view is shown', () => {
    const { s } = data();
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1 }], true);
    expect([1, 2, 3].map((r) => s.value(r, 0))).toEqual(['Al', 'Bo', 'Cy']);
    expect(s.rowOrderOf()).toBeNull();
  });

  it('draws cells at their display rows', () => {
    const { s } = data();
    createView(s, 'V');
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1 }], true);
    const o = s.rowOrderOf()!;
    const layout = withRowOrder(layoutOf(s.sheet, 0, s.wb.styles), o);
    const vp: Viewport = { layout, zoom: 1, headerW: 40, headerH: 20, width: 800, height: 600, scrollX: 0, scrollY: 0, frozenRows: 0, frozenCols: 0, rowOrder: o };
    const p = panesOf(vp)[0]!;
    // "Al" is data row 2 and shows at display row 1.
    expect(rectIn(p, vp, { r1: 2, c1: 0, r2: 2, c2: 0 }).y).toBe(rectIn(p, { ...vp, rowOrder: null }, { r1: 1, c1: 0, r2: 1, c2: 0 }).y);
  });

  it('keeps the view after undo replaces the sheet', () => {
    const { s, type } = data();
    createView(s, 'V');
    sortRange(s, { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1 }], true);
    const before = [...s.wb.sheets];
    type('D1', 'x');
    s.change('Sheet change', 'sheet', () => { s.sheet.name = s.sheet.name; });
    s.undo();
    carryViews(s, before);
    expect(s.rowOrderOf()).not.toBeNull();
  });
});

describe('selection protection with the keyboard', () => {
  it('skips cells that may not be selected and refuses Go To on them', () => {
    const { s } = setup();
    s.wb.sheets[0]!.protection = { sheet: true, options: { selectLockedCells: true } };
    // Column B is unlocked; everything else is locked and can't be selected.
    s.styleRange({ r1: 0, c1: 1, r2: 20, c2: 1 }, (st) => ({ ...st, protection: { locked: false } }));
    s.styleRange({ r1: 0, c1: 3, r2: 20, c2: 3 }, (st) => ({ ...st, protection: { locked: false } }));
    s.selectable = (r, c) => !blocked(s.sheet, s.isLocked(r, c) ? 'selectLockedCells' : 'selectUnlockedCells');
    s.selectRange({ r1: 0, c1: 1, r2: 0, c2: 1 });
    s.move(0, 1);
    expect(s.sel.active).toEqual({ row: 0, col: 3 });
    s.move(1, 0);
    expect(s.sel.active).toEqual({ row: 1, col: 3 });
    s.move(0, -1);
    expect(s.sel.active).toEqual({ row: 1, col: 1 });
    s.move(0, -1);
    expect(s.sel.active).toEqual({ row: 1, col: 1 });
    gotoReference(s, 'A5');
    expect(s.sel.active).toEqual({ row: 1, col: 1 });
    gotoReference(s, 'D7');
    expect(s.sel.active).toEqual({ row: 6, col: 3 });
  });
});

describe('ink in split panes', () => {
  it('maps sheet points to the pane that shows them and back', () => {
    const sh = newSheet('S');
    const layout = layoutOf(sh, 0, [{}]);
    const vp: Viewport = { layout, zoom: 2, headerW: 40, headerH: 20, width: 1000, height: 700, scrollX: 3000, scrollY: 2000, frozenRows: 0, frozenCols: 0, split: { w: 200, h: 100, sx: 0, sy: 0 } };
    // A point near the origin shows in the top-left pane.
    expect(toScreen(vp, 10, 10)).toEqual([40 + 20, 20 + 20]);
    // A far point shows in the main pane.
    const [x, y] = toScreen(vp, 3050, 2020);
    expect([x, y]).toEqual([40 + 200 + 5 + 100, 20 + 100 + 5 + 40]);
    expect(toSheet(vp, x, y)).toEqual([3050, 2020]);
    expect(toSheet(vp, 60, 40)).toEqual([10, 10]);
  });
});

describe('undo after a restore', () => {
  it('brings back calculation settings and the date system', () => {
    const { s } = setup();
    s.wb.calc = { iterate: true };
    s.change('Restore version', 'book', () => { s.wb.calc = { manual: true }; s.wb.date1904 = true; });
    s.undo();
    expect(s.wb.calc).toEqual({ iterate: true });
    expect(s.wb.date1904).toBe(false);
    s.redo();
    expect(s.wb.calc).toEqual({ manual: true });
    expect(s.wb.date1904).toBe(true);
  });
});
