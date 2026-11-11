import { copyFileSync, existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { buildCli, ROOT, runCli } from './cli-helpers';

let bin = '';
let dir = '';
let cleanup = () => {};
const ex = (name: string) => join(ROOT, 'examples', name);
const run = (...args: string[]) => runCli(bin, args);

beforeAll(() => { ({ bin, dir, cleanup } = buildCli()); }, 120_000);
afterAll(() => cleanup());

describe('clerdogrid command-line tool', () => {
  it('starts with a shebang and runs on plain Node', () => {
    expect(readFileSync(bin, 'utf8').startsWith('#!/usr/bin/env node')).toBe(true);
    const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8'));
    expect(run('--version')).toMatchObject({ code: 0, out: `clerdogrid ${pkg.version}\n` });
  });

  it('prints help and rejects unknown commands and options', () => {
    expect(run('--help').out).toContain('Usage: clerdogrid <command>');
    expect(run('help', 'set').out).toContain('--out');
    const none = run();
    expect(none.code).toBe(2);
    expect(none.err).toContain('Usage');
    const bad = run('frobnicate');
    expect(bad.code).toBe(2);
    expect(bad.err).toMatch(/^clerdogrid: Unknown command/);
    const opt = run('info', ex('sales.xlsx'), '--bogus');
    expect(opt.code).toBe(2);
    expect(opt.err).not.toMatch(/\n\s+at /);
  });

  it('reports missing and unreadable files without a stack trace', () => {
    const r = run('info', join(dir, 'missing.xlsx'));
    expect(r.code).toBe(3);
    expect(r.err).toContain("doesn't exist");
    writeFileSync(join(dir, 'junk.xlsx'), 'not a zip');
    const j = run('info', join(dir, 'junk.xlsx'));
    expect(j.code).toBe(3);
    expect(j.err).not.toMatch(/\n\s+at /);
  });

  it('describes every example workbook', () => {
    for (const name of ['attendance', 'budget', 'expenses', 'formula-tour', 'grades', 'invoice', 'loan', 'sales-data', 'sales', 'todo']) {
      const r = run('info', ex(`${name}.xlsx`), '--json');
      expect(r.code, `${name}: ${r.err}`).toBe(0);
      const info = JSON.parse(r.out);
      expect(info.sheets.length).toBeGreaterThan(0);
      expect(info.totals.cells).toBeGreaterThan(0);
    }
    const text = run('info', ex('sales.xlsx')).out;
    expect(text).toContain('Sheet "Sales"');
    expect(text).toMatch(/Chart line "Sales by month"/);
  });

  it('reads cells and ranges as text, raw values, formulas and JSON', () => {
    expect(run('get', ex('sales.xlsx'), 'A1:B2').out).toBe('Month\tNorth\nJan\t12,400\n');
    expect(run('get', ex('sales.xlsx'), 'B2', '--raw').out).toBe('12400\n');
    expect(run('get', ex('sales.xlsx'), 'Sales!E2', '--formulas').out).toBe('=SUM(B2:D2)\n');
    const j = JSON.parse(run('get', ex('sales.xlsx'), "'Sales'!E2:E3", '--json').out);
    expect(j).toMatchObject({ sheet: 'Sales', range: 'E2:E3', values: [[26500], [28200]], formulas: [['=SUM(B2:D2)'], ['=SUM(B2:D3)'.replace('B2', 'B3')]] });
    const bad = run('get', ex('sales.xlsx'), 'Nope!A1');
    expect(bad.code).toBe(2);
    expect(bad.err).toContain('no sheet named "Nope"');
  });

  it('sets a value, recalculates dependents and saves', () => {
    const f = join(dir, 'sales-set.xlsx');
    copyFileSync(ex('sales.xlsx'), f);
    const r = run('set', f, 'B2', '20000');
    expect(r.code, r.err).toBe(0);
    expect(run('get', f, 'E2', '--raw').out).toBe(String(20000 + 9800 + 4300) + '\n');
    const out = join(dir, 'sales-formula.xlsx');
    expect(run('set', f, 'G1', '=B2*2', '-o', out).code).toBe(0);
    expect(run('get', out, 'G1', '--raw').out).toBe('40000\n');
    expect(run('get', f, 'G1', '--raw').out).toBe('\n');
    expect(run('set', f, 'G2', '=SUM((').code).toBe(2);
    expect(run('set', f, 'A1:B2', '1').code).toBe(2);
  });

  it('recalculates and saves cached values', () => {
    const f = join(dir, 'loan-calc.xlsx');
    copyFileSync(ex('loan.xlsx'), f);
    const r = run('calc', f, '--json');
    expect(r.code, r.err).toBe(0);
    expect(JSON.parse(r.out)).toMatchObject({ circular: [] });
    expect(JSON.parse(r.out).formulas).toBeGreaterThan(10);
  });

  it('evaluates formulas on their own and against a workbook', () => {
    expect(run('eval', '=SUM(SEQUENCE(10))').out).toBe('55\n');
    expect(run('eval', 'TEXTJOIN("-",TRUE,"a","b")').out).toBe('a-b\n');
    expect(JSON.parse(run('eval', '=1/0', '--json').out).value).toBe('#DIV/0!');
    expect(run('eval', '=SUM(B2:D2)', '--file', ex('sales.xlsx')).out).toBe('26500\n');
    expect(run('eval', '=SEQUENCE(2,2)').out).toBe('1\t2\n3\t4\n');
    expect(run('eval', '=(').code).toBe(2);
  });

  it('converts between xlsx, csv, tsv and json', () => {
    const csv = join(dir, 'sales.csv'), tsv = join(dir, 'sales.tsv'), json = join(dir, 'sales.json'), back = join(dir, 'back.xlsx');
    expect(run('convert', ex('sales.xlsx'), csv).code).toBe(0);
    expect(readFileSync(csv, 'utf8').split('\r\n')[1]).toBe('Jan,"12,400","9,800","4,300","26,500"');
    expect(run('convert', ex('sales.xlsx'), csv).code).toBe(3);
    expect(run('convert', ex('sales.xlsx'), tsv, '--raw', '--force').code).toBe(0);
    expect(readFileSync(tsv, 'utf8').split('\n')[1]).toBe('Jan\t12400\t9800\t4300\t26500');
    expect(run('convert', ex('sales.xlsx'), json, '--formulas').code).toBe(0);
    const data = JSON.parse(readFileSync(json, 'utf8'));
    expect(data.sheets[0].rows[1][4]).toBe('=SUM(B2:D2)');
    expect(run('convert', json, back).code).toBe(0);
    expect(run('get', back, 'E2', '--raw').out).toBe('26500\n');
    expect(run('convert', csv, join(dir, 'from-csv.xlsx')).code).toBe(0);
    expect(run('get', join(dir, 'from-csv.xlsx'), 'B2', '--raw').out).toBe('12400\n');
    expect(run('convert', ex('sales.xlsx'), join(dir, 'x.unknown')).code).toBe(2);
    expect(run('convert', '--formats').out).toMatch(/Reads:.*xlsx.*csv/);
  });

  it('saves and reopens password-protected workbooks when the build can', () => {
    if (!run('convert', '--formats').out.includes('Password-protected')) return;
    const f = join(dir, 'locked.xlsx');
    expect(run('convert', ex('sales.xlsx'), f, '--encrypt', 'pa55').code).toBe(0);
    expect(readFileSync(f).subarray(0, 4)).toEqual(Buffer.from([0xd0, 0xcf, 0x11, 0xe0]));
    const none = run('info', f);
    expect(none.code).toBe(3);
    expect(none.err).toContain('--password');
    expect(run('info', f, '--password', 'wrong').err).toContain('not right');
    expect(run('get', f, 'E2', '--raw', '--password', 'pa55').out).toBe('26500\n');
    expect(run('convert', ex('sales.xlsx'), join(dir, 'locked.csv'), '--encrypt', 'x').code).toBe(2);
  });

  it('prints a sheet or range as CSV', () => {
    const r = run('csv', ex('sales.xlsx'), '--range', 'A1:C2');
    expect(r.out).toBe('Month,North,South\r\nJan,"12,400","9,800"\r\n');
    expect(run('csv', ex('sales.xlsx'), '--range', 'A1:B2', '--tsv', '--raw').out).toBe('Month\tNorth\nJan\t12400\n');
  });

  it('finds text, values and formulas', () => {
    expect(run('find', ex('sales.xlsx'), 'Jan').out).toBe('Sales!A2\tJan\n');
    const f = JSON.parse(run('find', ex('sales.xlsx'), 'SUM(B2', '--json').out);
    expect(f.map((h: { cell: string }) => h.cell)).toContain('E2');
    expect(run('find', ex('sales.xlsx'), 'no such text').code).toBe(1);
  });

  it('lists and describes functions', () => {
    const list = run('functions', '--search', 'xlookup');
    expect(list.out.split('\n')[0]).toMatch(/^XLOOKUP\s/);
    const all = JSON.parse(run('functions', '--json').out);
    expect(all.length).toBeGreaterThan(400);
    const d = run('describe', 'vlookup');
    expect(d.out).toContain('VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])');
    const missing = run('describe', 'VLOKUP');
    expect(missing.code).toBe(2);
    expect(missing.err).toContain('Did you mean');
  });

  it('makes workbooks from templates', () => {
    const list = run('new', '--list', '--json');
    expect(JSON.parse(list.out).map((t: { id: string }) => t.id)).toEqual(expect.arrayContaining(['blank', 'budget', 'invoice']));
    const f = join(dir, 'new-invoice.xlsx');
    expect(run('new', f, 'invoice').code).toBe(0);
    expect(existsSync(f)).toBe(true);
    expect(JSON.parse(run('info', f, '--json').out).totals.formulas).toBeGreaterThan(0);
    expect(run('new', f, 'invoice').code).toBe(3);
    expect(run('new', join(dir, 'blank.xlsx')).code).toBe(0);
    expect(run('new', join(dir, 'bad.xlsx'), 'nope').code).toBe(2);
  });

  it('validates formulas', () => {
    expect(run('validate', ex('formula-tour.xlsx'))).toMatchObject({ code: 0 });
    const f = join(dir, 'broken.xlsx');
    copyFileSync(ex('sales.xlsx'), f);
    run('set', f, 'H1', '=NOSUCHFN(1)');
    run('set', f, 'H2', '=H3');
    run('set', f, 'H3', '=H2+1');
    run('set', f, 'H4', '=1/0');
    run('set', f, 'H5', '=#REF!+1');
    const r = run('validate', f, '--json');
    expect(r.code).toBe(1);
    const kinds = JSON.parse(r.out).problems.map((p: { cell: string; kind: string }) => `${p.cell}:${p.kind}`);
    expect(kinds).toEqual(expect.arrayContaining(['H1:unknownFunction', 'H2:circular', 'H3:circular', 'H4:error', 'H5:ref']));
    expect(kinds).not.toContain('H1:error');
  });

  it('compares workbooks cell by cell', () => {
    expect(run('diff', ex('sales.xlsx'), ex('sales.xlsx'))).toMatchObject({ code: 0, out: 'No differences\n' });
    const f = join(dir, 'diff.xlsx');
    copyFileSync(ex('sales.xlsx'), f);
    run('set', f, 'B2', '1');
    run('set', f, 'J9', 'new');
    const r = run('diff', ex('sales.xlsx'), f);
    expect(r.code).toBe(1);
    expect(r.out).toContain('~ Sales!B2\t12400\t->\t1');
    expect(r.out).toContain('+ Sales!J9\tnew');
    expect(r.out).toContain('~ Sales!E2');
    const j = JSON.parse(run('diff', ex('sales.xlsx'), ex('budget.xlsx'), '--json').out);
    expect(j.differences.map((d: { kind: string }) => d.kind)).toEqual(expect.arrayContaining(['sheetAdded', 'sheetRemoved']));
  });
});
