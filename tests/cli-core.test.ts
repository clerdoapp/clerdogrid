import { describe, expect, it } from 'vitest';
import { allowOnly, parseArgs } from '../src/cli/args';
import { createChart, diff, find, openStore, parseTarget, readRange, validate, writeCells } from '../src/cli/core';
import { UserError } from '../src/cli/errors';
import { kindOf, readWorkbook, workbookFromJson, writeWorkbook } from '../src/cli/io';
import { check, Server } from '../src/cli/mcp';
import { aboutInfo, parseNotices, repoUrl } from '../src/features/about/info';
import { newWorkbook } from '../src/model/workbook';

describe('argument parsing', () => {
  it('reads flags, values, aliases, repeats and positionals', () => {
    const p = parseArgs(['get', 'a.xlsx', '-sData', '--json', '--range=A1:B2', '-o', 'x.csv', '--root', 'a', '--root', 'b', '-5', '--', '--literal'], ['sheet', 'range', 'out', 'root'], { s: 'sheet', o: 'out' });
    expect(p.positional).toEqual(['get', 'a.xlsx', '-5', '--literal']);
    expect(p.flags.get('sheet')).toBe('Data');
    expect(p.flags.get('json')).toBe(true);
    expect(p.flags.get('range')).toBe('A1:B2');
    expect(p.flags.get('out')).toBe('x.csv');
    expect(p.lists.get('root')).toEqual(['a', 'b']);
    expect(() => parseArgs(['--sheet'], ['sheet'])).toThrow(UserError);
    expect(() => parseArgs(['--json=1'], [])).toThrow(/doesn't take a value/);
    expect(() => allowOnly(parseArgs(['--nope'], []), ['json'])).toThrow(/Unknown option --nope/);
  });
});

describe('workbook files', () => {
  it('tells formats from file names', () => {
    expect(kindOf('a.XLSX')).toBe('xlsx');
    expect(kindOf('dir.v2/a.tsv')).toBe('tsv');
    expect(kindOf('noext')).toBeUndefined();
  });

  it('saves and reopens JSON with values, formulas and several sheets', async () => {
    const store = openStore(newWorkbook());
    writeCells(store, 0, 0, 0, [['Item', 'Cost'], ['Pen', 2.5], ['Pad', 4], ['Total', '=SUM(B2:B3)'], [true, '#N/A']]);
    const bytes = await writeWorkbook(store.wb, 'out.json', { formulas: true });
    const again = openStore(await readWorkbook('out.json', bytes));
    expect(readRange(again, parseTarget(again, 'A1:B5')).values).toEqual([['Item', 'Cost'], ['Pen', 2.5], ['Pad', 4], ['Total', 6.5], [true, '#N/A']]);
    expect(readRange(again, parseTarget(again, 'B4')).formulas).toEqual([['=SUM(B2:B3)']]);
    const plain = JSON.parse(new TextDecoder().decode(await writeWorkbook(store.wb, 'v.json')));
    expect(plain.sheets[0].rows[3][1]).toBe(6.5);
  });

  it('reads arrays of rows and arrays of records', () => {
    const rows = openStore(workbookFromJson('people.json', '[["a","b"],[1,2]]'));
    expect(rows.wb.sheets[0]!.name).toBe('people');
    expect(readRange(rows, parseTarget(rows, 'A1:B2')).values).toEqual([['a', 'b'], [1, 2]]);
    const recs = openStore(workbookFromJson('r.json', '[{"x":1,"y":"=A2*2"},{"x":3,"z":true}]'));
    expect(readRange(recs, parseTarget(recs, 'A1:C3')).values).toEqual([['x', 'y', 'z'], [1, 2, null], [3, null, true]]);
    expect(() => workbookFromJson('bad.json', '{')).toThrow(UserError);
    expect(() => workbookFromJson('bad.json', '{"a":1}')).toThrow(/expected/);
  });

  it('saves and reopens xlsx and csv', async () => {
    const store = openStore(newWorkbook());
    writeCells(store, 0, 0, 0, [['n', 'twice'], [3, '=A2*2'], ['12%', 'x,y']]);
    const x = openStore(await readWorkbook('t.xlsx', await writeWorkbook(store.wb, 't.xlsx')));
    expect(readRange(x, parseTarget(x, 'A1:B3')).values).toEqual([['n', 'twice'], [3, 6], [0.12, 'x,y']]);
    const c = openStore(await readWorkbook('t.csv', await writeWorkbook(store.wb, 't.csv')));
    expect(readRange(c, parseTarget(c, 'A1:B3')).text).toEqual([['n', 'twice'], ['3', '6'], ['12%', 'x,y']]);
    await expect(readWorkbook('t.xlsx', new Uint8Array([1, 2, 3]))).rejects.toThrow(UserError);
  });
});

describe('workbook operations', () => {
  const sample = () => {
    const s = openStore(newWorkbook());
    writeCells(s, 0, 0, 0, [['Month', 'East', 'West'], ['Jan', 10, 20], ['Feb', 30, 40], ['Total', '=SUM(B2:B3)', '=SUM(C2:C3)']]);
    return s;
  };

  it('parses targets: cells, ranges, sheets, whole columns and names', () => {
    const s = sample();
    s.wb.names.push({ name: 'Totals', formula: 'Sheet1!$B$4:$C$4' });
    expect(parseTarget(s, 'b2').range).toEqual({ r1: 1, c1: 1, r2: 1, c2: 1 });
    expect(parseTarget(s, "'Sheet1'!$A$1:B2").range).toEqual({ r1: 0, c1: 0, r2: 1, c2: 1 });
    expect(parseTarget(s, 'B:C').range).toEqual({ r1: 0, c1: 1, r2: 3, c2: 2 });
    expect(parseTarget(s, '2:3').range).toEqual({ r1: 1, c1: 0, r2: 2, c2: 2 });
    expect(parseTarget(s, 'Totals').range).toEqual({ r1: 3, c1: 1, r2: 3, c2: 2 });
    expect(() => parseTarget(s, 'Nope!A1')).toThrow(/no sheet named/);
    expect(() => parseTarget(s, 'what')).toThrow(UserError);
  });

  it('finds, validates and diffs', () => {
    const a = sample(), b = sample();
    expect(find(a, 'feb').map((h) => h.cell)).toEqual(['A3']);
    expect(find(a, 'SUM(C', {}).map((h) => h.cell)).toEqual(['C4']);
    expect(validate(a)).toEqual([]);
    writeCells(b, 0, 1, 1, [[11]]);
    expect(diff(a, b).map((d) => `${d.cell}:${d.kind}`)).toEqual(['B2:value', 'B4:value']);
    writeCells(b, 0, 5, 0, [['=LET(f,LAMBDA(x,x*2),f(4))']]);
    expect(validate(b)).toEqual([]);
  });

  it('builds charts from a data block', () => {
    const s = sample();
    const c = createChart(s, parseTarget(s, 'A1:C3'), { type: 'line', title: 'T' });
    expect(c.series.map((x) => [x.name, x.values, x.categories])).toEqual([
      ['Sheet1!$B$1', 'Sheet1!$B$2:$B$3', 'Sheet1!$A$2:$A$3'],
      ['Sheet1!$C$1', 'Sheet1!$C$2:$C$3', 'Sheet1!$A$2:$A$3'],
    ]);
    expect(() => createChart(s, parseTarget(s, 'A1:C3'), { type: 'nope' })).toThrow(/Unknown chart type/);
  });
});

describe('MCP input checks', () => {
  it('applies the schema subset', () => {
    const schema = { type: 'object', properties: { a: { type: 'integer', minimum: 1 }, b: { type: 'string', enum: ['x'] } }, required: ['a'], additionalProperties: false };
    expect(() => check({ a: 2, b: 'x' }, schema)).not.toThrow();
    expect(() => check({}, schema)).toThrow(/a is required/);
    expect(() => check({ a: 0 }, schema)).toThrow(/at least 1/);
    expect(() => check({ a: 1.5 }, schema)).toThrow(/integer/);
    expect(() => check({ a: 1, b: 'y' }, schema)).toThrow(/one of/);
    expect(() => check({ a: 1, c: 1 }, schema)).toThrow(/known argument/);
  });

  it('refuses invalid JSON-RPC envelopes', async () => {
    const s = new Server({ roots: [process.cwd()] });
    expect(await s.handle({ id: 1, method: 'ping' })).toMatchObject({ error: { code: -32600 } });
    expect(await s.handle({ jsonrpc: '2.0', method: 'notifications/initialized' })).toBeNull();
    expect(() => new Server({ roots: ['/definitely/not/here'] })).toThrow(/isn't a folder/);
  });
});

describe('About information', () => {
  it('reads version, licence, links, contributors and notices', () => {
    const info = aboutInfo();
    expect(info.version).toMatch(/^\d+\.\d+\.\d+/);
    expect(info.license).toBe('MIT OR Apache-2.0');
    expect(info.links.map((l) => l.label)).toEqual(['Website', 'Source code', 'Releases', 'Report a problem']);
    expect(info.links.every((l) => l.url.startsWith('https://'))).toBe(true);
    expect(info.contributors.map((c) => c.name)).toEqual(expect.arrayContaining(['Clerdo', 'Claude (AI assistant)']));
    expect(info.notices.some((n) => n.startsWith('React, fflate'))).toBe(true);
  });

  it('keeps only https links and copes with broken data', () => {
    const info = aboutInfo('{"repository":"javascript:alert(1)","homepage":"http://x"}', '{"contributors":[{"name":"A","url":"javascript:x"},{"nope":1}]}', '', '1.0.0');
    expect(info.links).toEqual([]);
    expect(info.contributors).toEqual([{ name: 'A', role: undefined, url: undefined, github: undefined }]);
    expect(aboutInfo('{', '{', '', '1').contributors).toEqual([]);
    expect(repoUrl({ url: 'git+https://github.com/o/r.git' })).toBe('https://github.com/o/r');
    expect(parseNotices('X\n- one: MIT\n  continued\n- two')).toEqual(['one: MIT continued', 'two']);
  });
});
