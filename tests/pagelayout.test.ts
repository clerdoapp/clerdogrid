import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { MAX_COLS } from '../src/engine/types';
import { paginate, pageGrid, paperSize, parseAreas, usedArea } from '../src/features/pagelayout/paginate';
import { fieldText, hfFor, joinHf, parseHf, splitHf } from '../src/features/pagelayout/headerFooter';
import { insertBreak, moveBreak, removeBreak, resetBreaks, setPrintArea, addToPrintArea, clearPrintArea, updateSetup } from '../src/features/pagelayout/setup';
import { planPrint } from '../src/features/pagelayout/print';
import { lengthPt, paperByCode } from '../src/features/pagelayout/paper';
import { SheetLayout } from '../src/app/layout';

beforeAll(() => registerAllFunctions());

const filled = (rows: number, cols: number) => {
  const s = new GridStore();
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) s.enter(r, c, String(r + c));
  return s;
};

describe('paper', () => {
  it('knows sizes in points and turns for landscape', () => {
    expect(paperByCode(1)).toMatchObject({ w: 612, h: 792 });
    expect(paperByCode(9).w).toBeCloseTo(595.28, 1);
    expect(paperSize({ paper: 9, orientation: 'landscape' }).w).toBeCloseTo(841.89, 1);
    expect(lengthPt('210mm')).toBeCloseTo(595.28, 1);
    expect(lengthPt('8.5in')).toBe(612);
    expect(lengthPt('bad')).toBeUndefined();
  });
});

describe('paginate', () => {
  it('splits by default row heights and column widths on Letter paper', () => {
    const s = filled(100, 15);
    const pg = paginate(s.sheet);
    // 684pt of printable height / 15pt rows = 45 rows; 511.2pt / 48pt columns = 10 columns.
    expect(pg.pages.map((p) => [p.rows, p.cols])).toEqual([
      [[0, 44], [0, 9]], [[45, 89], [0, 9]], [[90, 99], [0, 9]],
      [[0, 44], [10, 14]], [[45, 89], [10, 14]], [[90, 99], [10, 14]],
    ]);
    expect(pg.pages.map((p) => p.number)).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it('follows over-then-down order and the first page number', () => {
    const s = filled(100, 15);
    s.sheet.pageSetup = { order: 'overThenDown', firstPageNumber: 7 };
    const pg = paginate(s.sheet);
    expect(pg.pages.slice(0, 3).map((p) => [p.rows[0], p.cols[0]])).toEqual([[0, 0], [0, 10], [45, 0]]);
    expect(pg.pages[0]!.number).toBe(7);
  });

  it('honours manual breaks and hidden rows', () => {
    const s = filled(60, 3);
    s.sheet.pageSetup = { rowBreaks: [10, 30] };
    s.sheet.rows.set(5, { hidden: true });
    const pg = paginate(s.sheet);
    expect(pg.pages.map((p) => p.rows)).toEqual([[0, 9], [10, 29], [30, 59]]);
    expect(pg.bands[0]!.manualRows.has(10)).toBe(true);
    expect(pageGrid(pg, pg.pages[0]!).rows.map((r) => r.index)).not.toContain(5);
  });

  it('scales down to fit pages wide and tall, ignoring manual breaks', () => {
    const s = filled(100, 15);
    s.sheet.pageSetup = { fitWidth: 1, fitHeight: 1, rowBreaks: [20] };
    const pg = paginate(s.sheet);
    expect(pg.pages).toHaveLength(1);
    expect(pg.fitPercent).toBeLessThan(100);
    expect(pg.fitPercent).toBeGreaterThan(40);
    const wide = filled(100, 15);
    wide.sheet.pageSetup = { fitWidth: 1, fitHeight: 0 };
    const pw = paginate(wide.sheet);
    expect(new Set(pw.pages.map((p) => p.cols[0]))).toEqual(new Set([0]));
    expect(pw.pages.length).toBeGreaterThan(1);
  });

  it('uses a fixed scale', () => {
    const s = filled(100, 3);
    s.sheet.pageSetup = { scale: 50 };
    expect(paginate(s.sheet).pages).toHaveLength(2);
  });

  it('prints each print area on its own pages and limits whole columns to the used range', () => {
    const s = filled(30, 6);
    s.sheet.pageSetup = { printArea: 'A1:B5,D10:E12' };
    const pg = paginate(s.sheet);
    expect(pg.pages.map((p) => [p.rows, p.cols])).toEqual([[[0, 4], [0, 1]], [[9, 11], [3, 4]]]);
    s.sheet.pageSetup = { printArea: 'B:C' };
    expect(paginate(s.sheet).areas[0]).toEqual({ r1: 0, c1: 1, r2: 29, c2: 2 });
    expect(parseAreas('$A$1:$B$2,3:4')).toEqual([{ r1: 0, c1: 0, r2: 1, c2: 1 }, { r1: 2, r2: 3, c1: 0, c2: MAX_COLS - 1 }]);
  });

  it('repeats print titles on later pages and leaves room for them', () => {
    const s = filled(100, 3);
    s.sheet.pageSetup = { printTitleRows: '1:2' };
    const pg = paginate(s.sheet);
    expect(pg.pages[0]!.titleRows).toBeUndefined();
    expect(pg.pages[1]!.titleRows).toEqual([0, 1]);
    expect(pg.pages[1]!.rows[1] - pg.pages[1]!.rows[0] + 1).toBe(43);
    const g = pageGrid(pg, pg.pages[1]!);
    expect(g.rows.slice(0, 3).map((r) => r.index)).toEqual([0, 1, 45]);
  });

  it('centers the block and leaves room for headings', () => {
    const s = filled(3, 2);
    s.sheet.pageSetup = { centerH: true, centerV: true, headings: true };
    const pg = paginate(s.sheet);
    const p = pg.pages[0]!;
    expect(pg.headW).toBeGreaterThan(0);
    expect(p.w).toBeCloseTo(pg.headW + 96, 5);
    expect(p.x).toBeCloseTo(50.4 + (511.2 - p.w) / 2, 5);
    expect(p.y).toBeCloseTo(54 + (684 - p.h) / 2, 5);
  });

  it('counts fills and objects in the default print range, not empty cells', () => {
    const s = new GridStore();
    s.enter(2, 2, 'x');
    s.sheet.cells.set(5 * MAX_COLS + 7, { v: null, s: s.wb.styles.push({ fill: { color: '#ff0000' } }) - 1 });
    s.sheet.cells.set(40 * MAX_COLS + 40, { v: null });
    const used = usedArea(s.sheet, s.wb.styles, new SheetLayout(s.sheet));
    expect(used).toEqual({ r1: 2, c1: 2, r2: 5, c2: 7 });
    expect(paginate(new GridStore().sheet).pages).toHaveLength(0);
  });
});

describe('headers and footers', () => {
  const ctx = { page: 3, pages: 9, date: new Date(2026, 0, 2, 14, 5), file: 'Book.xlsx', path: '/x/', sheet: 'Data' };
  it('parses sections, fields and font codes', () => {
    const h = parseHf('&L&"Arial,Bold Italic"&14Left &P+1&C&KFF0000Center &N&R&BRight&B &A && more');
    expect(h.L[0]).toMatchObject({ kind: 'text', text: 'Left ', font: { name: 'Arial', bold: true, italic: true, size: 14 } });
    expect(h.L[1]).toMatchObject({ kind: 'field', field: 'P', offset: 1 });
    expect(h.C[0]).toMatchObject({ kind: 'text', text: 'Center ', font: { color: '#FF0000' } });
    expect(h.R.map((r) => r.kind === 'text' ? r.text : r.kind)).toEqual(['Right', ' ', 'field', ' & more']);
    expect(h.R[0]!.font.bold).toBe(true);
    expect(h.R[1]!.font.bold).toBe(false);
    expect(fieldText('P', 1, ctx)).toBe('4');
    expect(fieldText('A', undefined, ctx)).toBe('Data');
  });
  it('reads theme colors, underline, strike, sub and superscript and pictures', () => {
    const h = parseHf('&C&K04+050&U&E&S&X&Ya&G');
    expect(h.C[0]).toMatchObject({ font: { color: '#A2B9E2', underline: 'double', strike: true, vert: 'sub' } });
    expect(h.C[1]).toMatchObject({ kind: 'picture' });
  });
  it('defaults to the center section and round-trips through the editor', () => {
    expect(parseHf('Page &P').C).toHaveLength(2);
    const raw = '&L&"Tinos,Bold"Left&&&CMiddle&RRight &P';
    const parts = splitHf(raw);
    expect(parts).toEqual({ L: '&"Tinos,Bold"Left&&', C: 'Middle', R: 'Right &P' });
    expect(joinHf(parts)).toBe(raw);
  });
  it('picks first-page and even-page variants', () => {
    const s = { header: 'odd', evenHeader: 'even', firstHeader: 'first', differentFirst: true, differentOddEven: true };
    expect([1, 2, 3, 4].map((p) => hfFor(s, p).header)).toEqual(['first', 'even', 'odd', 'even']);
  });
});

describe('page commands', () => {
  it('sets, adds to and clears the print area with undo', () => {
    const s = filled(10, 10);
    s.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 });
    setPrintArea(s);
    expect(s.sheet.pageSetup?.printArea).toBe('A1:C4');
    s.selectRange({ r1: 6, c1: 6, r2: 7, c2: 7 });
    addToPrintArea(s);
    expect(s.sheet.pageSetup?.printArea).toBe('A1:C4,G7:H8');
    clearPrintArea(s);
    expect(s.sheet.pageSetup).toBeUndefined();
    s.undo();
    expect(s.sheet.pageSetup?.printArea).toBe('A1:C4,G7:H8');
  });
  it('inserts, moves, removes and resets page breaks', () => {
    const s = filled(80, 20);
    s.select({ row: 20, col: 5 });
    insertBreak(s);
    expect(s.sheet.pageSetup).toMatchObject({ rowBreaks: [20], colBreaks: [5] });
    moveBreak(s, 'row', 20, 25, { r1: 0, c1: 0, r2: 79, c2: 19 });
    expect(s.sheet.pageSetup?.rowBreaks).toEqual([25]);
    moveBreak(s, 'row', 30, 35, { r1: 0, c1: 0, r2: 79, c2: 19 });
    expect(s.sheet.pageSetup?.rowBreaks).toEqual([25, 35]);
    s.select({ row: 25, col: 5 });
    removeBreak(s);
    expect(s.sheet.pageSetup?.rowBreaks).toEqual([35]);
    expect(s.sheet.pageSetup?.colBreaks).toBeUndefined();
    resetBreaks(s);
    expect(s.sheet.pageSetup).toBeUndefined();
  });
  it('plans the workbook, the selection and comments at the end', () => {
    const s = filled(5, 3);
    s.sheet.notes.set(1 * MAX_COLS + 1, { author: 'Ann', text: 'Check this' });
    updateSetup(s, 'x', { cellComments: 'atEnd', footer: '&CPage &P of &N' });
    const job = planPrint(s, { what: 'workbook' });
    expect(job.pages).toHaveLength(2);
    expect(job.pages[1]!.notes).toEqual([{ cell: 'B2', text: 'Ann: Check this' }]);
    expect(job.pages.map((p) => p.total)).toEqual([2, 2]);
    s.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
    const sel = planPrint(s, { what: 'selection' });
    expect(sel.pages[0]!.page.rows).toEqual([1, 2]);
  });
});
