import { beforeAll, describe, expect, test } from 'vitest';
import { registerAllFunctions } from '../src/engine/functions';
import { GridStore } from '../src/app/store';
import { cellKey, newWorkbook } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { applyOps, runDirect, snapshot } from '../src/features/automate/apply';
import { buildRuntime, workerSource, type ScriptOp } from '../src/features/automate/runtime';
import { SAMPLES } from '../src/features/automate/samples';

beforeAll(() => registerAllFunctions());

type Fn = (wb: any) => unknown; // eslint-disable-line @typescript-eslint/no-explicit-any

const fill = (s: GridStore, rows: (string | number | null)[][], sheet = 0) =>
  s.change('fill', 'book', () => rows.forEach((r, i) => r.forEach((v, j) => { if (v !== null) s.enter(i, j, String(v), sheet); })));

const values = (s: GridStore, r1: number, c1: number, r2: number, c2: number, sheet = 0) =>
  Array.from({ length: r2 - r1 + 1 }, (_, i) => Array.from({ length: c2 - c1 + 1 }, (_, j) => s.value(r1 + i, c1 + j, sheet)));

const exec = (s: GridStore, fn: Fn) => {
  const r = runDirect(s, fn as never);
  if (!r.ok) throw new Error(r.error);
  const err = applyOps(s, r.ops);
  if (err) throw new Error(err);
  return r;
};

/** Runs sample code text the same way the worker does (wrapped function), but in this process. */
const runSampleCode = (s: GridStore, code: string) => {
  const main = new Function('workbook', 'console', `${code}\nreturn main(workbook);`) as Fn; // test-only
  const r = buildRuntime().run(snapshot(s), (wb, con) => (main as unknown as (a: unknown, b: unknown) => unknown)(wb, con));
  if (!r.ok) throw new Error(r.error);
  const err = applyOps(s, r.ops, 'Run sample');
  if (err) throw new Error(err);
  return r;
};

describe('script API', () => {
  test('reads and writes values and formulas as one undo step', () => {
    const s = new GridStore();
    fill(s, [['Name', 'Qty'], ['a', 2], ['b', 3]]);
    const r = exec(s, (wb) => {
      const sh = wb.getActiveSheet();
      const v = sh.getRange('A1:B3').getValues();
      console.log(v.length);
      sh.getRange('C1:C3').setValues([['Double'], [v[1][1] * 2], [v[2][1] * 2]]);
      sh.getCell(3, 1).setFormula('=SUM(B2:B3)');
      sh.getRange('D1').setValue('=B2*10');
      expect(sh.getRange('B4').getFormulas()).toEqual([['=SUM(B2:B3)']]);
      expect(sh.getUsedRange().getAddress()).toBe('A1:D4');
      expect(sh.getRange('A1:C3').getRowCount()).toBe(3);
      expect(sh.getRange('A1:C3').getColumnCount()).toBe(3);
    });
    expect(r.logs).toEqual([]);
    expect(values(s, 0, 2, 2, 2)).toEqual([['Double'], [4], [6]]);
    expect(s.value(3, 1)).toBe(5);
    expect(s.cell(3, 1)?.f).toBe('SUM(B2:B3)');
    expect(s.value(0, 3)).toBe(20);
    expect(s.undoLabel).toBe('Run script');
    s.undo();
    expect(s.value(1, 2)).toBe(null);
    expect(s.value(3, 1)).toBe(null);
    expect(s.value(0, 0)).toBe('Name');
  });

  test('console.log is captured', () => {
    const s = new GridStore();
    const r = runDirect(s, ((wb: any, ...rest: unknown[]) => { void rest; void wb; }) as never); // eslint-disable-line @typescript-eslint/no-explicit-any
    expect(r.ok).toBe(true);
    const r2 = buildRuntime().run(snapshot(s), (_wb, con) => { (con as { log: (...a: unknown[]) => void }).log('hi', { a: 1 }, [1, 2]); });
    expect(r2.logs).toEqual(['hi {"a":1} [1,2]']);
  });

  test('formats', () => {
    const s = new GridStore();
    fill(s, [['x', 1], ['y', 2]]);
    exec(s, (wb) => {
      wb.getActiveSheet().getRange('A1:B1').setBold(true).setFontColor('#FF0000').setFill('#00FF00').setHorizontalAlignment('center');
      wb.getActiveSheet().getRange('B2').setNumberFormat('0.00%');
    });
    const st = s.style(0, 1);
    expect(st.font?.bold).toBe(true);
    expect(st.font?.color).toBe('#FF0000');
    expect(st.fill?.color).toBe('#00FF00');
    expect(st.align?.h).toBe('center');
    expect(s.style(1, 1).numFmt).toBe('0.00%');
    exec(s, (wb) => { wb.getActiveSheet().getRange('A1:B1').setFill(null).setBold(false); });
    expect(s.style(0, 1).fill).toBeUndefined();
    expect(s.style(0, 1).font?.bold).toBeFalsy();
  });

  test('sorting inside the script and in the store agree', () => {
    const s = new GridStore();
    fill(s, [['Name', 'Score'], ['c', 3], ['a', 1], ['b', 2]]);
    exec(s, (wb) => {
      const r = wb.getActiveSheet().getRange('A2:B4');
      r.sort(1, false);
      expect(r.getValues()).toEqual([['c', 3], ['b', 2], ['a', 1]]);
    });
    expect(values(s, 1, 0, 3, 1)).toEqual([['c', 3], ['b', 2], ['a', 1]]);
    exec(s, (wb) => { wb.getActiveSheet().getRange('A2:B4').sort('A'); });
    expect(values(s, 1, 0, 3, 0)).toEqual([['a'], ['b'], ['c']]);
  });

  test('adds and renames sheets, freezes panes, selection', () => {
    const s = new GridStore();
    s.selectRange({ r1: 1, c1: 1, r2: 2, c2: 2 });
    exec(s, (wb) => {
      expect(wb.getSelectedRange().getAddress()).toBe('B2:C3');
      const n = wb.addSheet('Report');
      n.getRange('A1').setValue('hello');
      n.freezePanes(1, 0);
      wb.getActiveSheet().setName('Data');
      expect(wb.getSheets().map((x: { getName: () => string }) => x.getName())).toEqual(['Data', 'Report']);
      expect(wb.getSheet('report').getRange('A1').getValue()).toBe('hello');
      n.activate();
    });
    expect(s.wb.sheets.map((x) => x.name)).toEqual(['Data', 'Report']);
    expect(s.value(0, 0, 1)).toBe('hello');
    expect(s.wb.sheets[1]!.view.frozenRows).toBe(1);
    expect(s.sheetIndex).toBe(1);
    s.undo();
    expect(s.wb.sheets.map((x) => x.name)).toEqual(['Sheet1']);
  });

  test('errors are reported with nothing applied', () => {
    const s = new GridStore();
    fill(s, [[1]]);
    const r = runDirect(s, ((wb: any) => { wb.getActiveSheet().getRange('A1').setValue(5); wb.getActiveSheet().getRange('nope'); }) as never); // eslint-disable-line @typescript-eslint/no-explicit-any
    expect(r.ok).toBe(false);
    expect(r.error).toMatch(/isn't a valid range/);
    expect(r.ops).toEqual([]);
    expect(s.value(0, 0)).toBe(1);
    const shape = runDirect(s, ((wb: any) => { wb.getActiveSheet().getRange('A1:B2').setValues([[1]]); }) as never); // eslint-disable-line @typescript-eslint/no-explicit-any
    expect(shape.error).toMatch(/1×1 but the range A1:B2 is 2×2/);
  });

  test('bad formulas and malformed edits are refused', () => {
    const s = new GridStore();
    expect(applyOps(s, [{ t: 'formula', s: 0, r: 0, c: 0, f: 'SUM(' }])).toMatch(/can't be read/);
    expect(applyOps(s, [{ t: 'value', s: 0, r: -1, c: 0, v: 1 }])).toMatch(/does not understand/);
    expect(applyOps(s, [{ t: 'nope' } as unknown as ScriptOp])).toMatch(/does not understand/);
    expect(applyOps(s, [{ t: 'value', s: 5, r: 0, c: 0, v: 1 }])).toMatch(/does not exist/);
    expect(s.canUndo).toBe(false);
  });
});

describe('sandbox', () => {
  test('network and storage globals throw after shadowing', () => {
    const scope: Record<string, unknown> = { fetch: () => 'net', XMLHttpRequest: class {}, WebSocket: class {}, importScripts: () => 0, indexedDB: {} };
    buildRuntime().shadow(scope);
    for (const name of ['fetch', 'XMLHttpRequest', 'WebSocket', 'importScripts', 'indexedDB', 'EventSource', 'caches']) {
      expect(() => (scope[name] as () => void)()).toThrow(/isn't available in scripts/);
      expect(() => { scope[name] = () => 'again'; }).toThrow();
    }
  });

  test('shadowing reaches globals defined on the prototype', () => {
    class Proto { fetch() { return 'net'; } }
    const scope = Object.create(Proto.prototype) as { fetch: () => string };
    buildRuntime().shadow(scope);
    expect(() => scope.fetch()).toThrow(/isn't available/);
  });

  test('worker source shadows the network before the script and keeps the code inside a function', () => {
    const { source } = workerSource('function main(workbook) { fetch("https://example.com"); }');
    const prep = source.indexOf('__clerdoRt.prepare(self');
    const user = source.indexOf('function __clerdoUser');
    expect(prep).toBeGreaterThan(0);
    expect(user).toBeGreaterThan(prep);
    expect(source).toContain('__clerdoRt.install(self, __clerdoUser)');
    // Nothing from the person's code runs at the top level.
    expect(source.slice(0, user)).not.toContain('example.com');
  });

  test('the worker file runs in a fake worker scope and blocks fetch', async () => {
    const { source } = workerSource('function main(workbook) { console.log("start"); fetch("https://example.com"); }');
    const posted: unknown[] = [];
    const scope: Record<string, unknown> = { postMessage: (m: unknown) => posted.push(m), fetch: () => 'net' };
    scope.self = scope;
    // Run the generated file with `self` bound to the fake scope (test-only; the app uses a Blob worker).
    new Function('self', source.replace(/^"use strict";\n/, ''))(scope);
    (scope.onmessage as (e: unknown) => void)({ data: { snapshot: { sheets: [{ name: 'Sheet1', cells: [], hiddenRows: [], hiddenCols: [] }], active: 0, selection: { sheet: 0, r1: 0, c1: 0, r2: 0, c2: 0 } } } });
    const r = posted[0] as { ok: boolean; error: string; logs: string[] };
    expect(r.ok).toBe(false);
    expect(r.error).toMatch(/fetch isn't available/);
    expect(r.logs).toEqual(['start']);
    expect(() => (scope.postMessage as () => void)()).toThrow();
  });
});

describe('sample scripts', () => {
  const sample = (id: string) => SAMPLES.find((x) => x.id === id)!.code;

  test('there are eight samples, each with main()', () => {
    expect(SAMPLES.length).toBe(8);
    for (const x of SAMPLES) expect(x.code).toMatch(/function main\(workbook\)/);
  });

  test('unhide all rows and columns', () => {
    const s = new GridStore();
    fill(s, [[1, 2, 3], [4, 5, 6], [7, 8, 9]]);
    s.sheet.rows.set(1, { hidden: true });
    s.sheet.cols.set(2, { hidden: true });
    runSampleCode(s, sample('unhide'));
    expect(s.sheet.rows.get(1)?.hidden).toBeFalsy();
    expect(s.sheet.cols.get(2)?.hidden).toBeFalsy();
  });

  test('remove empty rows', () => {
    const s = new GridStore();
    fill(s, [['a', 1], [null, null], ['b', 2], [null, null], ['c', 3]]);
    const r = runSampleCode(s, sample('removeEmptyRows'));
    expect(values(s, 0, 0, 2, 0)).toEqual([['a'], ['b'], ['c']]);
    expect(r.logs[0]).toBe('Removed 2 empty rows');
  });

  test('proper case', () => {
    const s = new GridStore();
    fill(s, [['JOHN o\'neil'], ['mary-ann SMITH'], ['=UPPER("x")']]);
    s.selectRange({ r1: 0, c1: 0, r2: 2, c2: 0 });
    runSampleCode(s, sample('properCase'));
    expect(values(s, 0, 0, 1, 0)).toEqual([["John O'Neil"], ['Mary-Ann Smith']]);
    expect(s.cell(2, 0)?.f).toBe('UPPER("x")');
  });

  test('highlight duplicates', () => {
    const s = new GridStore();
    fill(s, [['x'], ['y'], ['X'], ['z']]);
    s.selectRange({ r1: 0, c1: 0, r2: 3, c2: 0 });
    runSampleCode(s, sample('highlightDuplicates'));
    expect(s.style(0, 0).fill?.color).toBe('#FFE08A');
    expect(s.style(2, 0).fill?.color).toBe('#FFE08A');
    expect(s.style(1, 0).fill).toBeUndefined();
  });

  test('add a total row', () => {
    const s = new GridStore();
    fill(s, [['Item', 'Qty', 'Note'], ['a', 2, 'x'], ['b', 5, 'y']]);
    runSampleCode(s, sample('totalRow'));
    expect(s.value(3, 0)).toBe('Total');
    expect(s.cell(3, 1)?.f).toBe('SUM(B2:B3)');
    expect(s.value(3, 1)).toBe(7);
    expect(s.value(3, 2)).toBe(null);
    expect(s.style(3, 0).font?.bold).toBe(true);
  });

  test('split names', () => {
    const s = new GridStore();
    fill(s, [['Ada Lovelace'], ['Grace Brewster Hopper'], ['Plato']]);
    s.selectRange({ r1: 0, c1: 0, r2: 2, c2: 0 });
    runSampleCode(s, sample('splitNames'));
    expect(values(s, 0, 1, 2, 2)).toEqual([['Ada', 'Lovelace'], ['Grace', 'Brewster Hopper'], ['Plato', null]]);
  });

  test('number the rows', () => {
    const s = new GridStore();
    s.change('fill', 'book', () => { s.enter(0, 1, 'Name'); s.enter(1, 1, 'a'); s.enter(2, 1, 'b'); });
    runSampleCode(s, sample('numberRows'));
    expect(values(s, 0, 0, 2, 0)).toEqual([['#'], [1], [2]]);
    expect(s.style(1, 0).align?.h).toBe('center');
  });

  test('summary sheet of all sheets', () => {
    const s = new GridStore();
    fill(s, [['a', 1], ['b', 2]]);
    s.change('add', 'book', () => { s.wb.sheets.push({ ...structuredClone({ ...s.wb.sheets[0]!, cells: undefined }), name: 'Two', cells: new Map(), cols: new Map(), rows: new Map(), hyperlinks: new Map(), notes: new Map() } as never); s.engine.sheetsChanged(); });
    s.change('fill2', 'book', () => { s.enter(0, 0, '10', 1); s.enter(1, 0, '5', 1); });
    runSampleCode(s, sample('summarySheet'));
    const i = s.wb.sheets.findIndex((x) => x.name === 'Summary');
    expect(i).toBe(2);
    expect(values(s, 0, 0, 2, 1, i)).toEqual([['Sheet', 'Total'], ['Sheet1', 3], ['Two', 15]]);
    expect(s.sheetIndex).toBe(2);
    expect(s.wb.sheets[i]!.cols.get(0)?.width).toBeGreaterThan(0);
  });
});

describe('xlsx', () => {
  test('scripts round-trip as their own part and never alter cells', () => {
    const wb = newWorkbook();
    wb.sheets[0]!.cells.set(cellKey(0, 0), { v: 1 });
    wb.scripts = [
      { id: 's1', name: 'Total <row> & "more"', code: 'function main(workbook) {\n  // a < b && c > d\n  console.log("hi");\n}\n' },
      { id: 's2', name: 'Second', code: '\tindented\n' },
    ];
    const back = readXlsx(writeXlsx(wb));
    expect(back.scripts).toEqual(wb.scripts);
    expect(back.sheets[0]!.cells.get(cellKey(0, 0))?.v).toBe(1);
    expect(readXlsx(writeXlsx(back)).scripts).toEqual(wb.scripts);
  });

  test('a workbook without scripts has no scripts part', () => {
    const wb = newWorkbook();
    const back = readXlsx(writeXlsx(wb));
    expect(back.scripts).toBeUndefined();
  });

  test('loading a file with scripts does not run them', () => {
    const wb = newWorkbook();
    wb.scripts = [{ id: 'x', name: 'Bad', code: 'function main(workbook) { workbook.getActiveSheet().getRange("A1").setValue("ran"); }' }];
    const s = new GridStore();
    s.load(readXlsx(writeXlsx(wb)), 'a.xlsx');
    expect(s.value(0, 0)).toBe(null);
    expect(s.wb.scripts?.length).toBe(1);
  });
});
