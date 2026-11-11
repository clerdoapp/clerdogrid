import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, symlinkSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { buildCli, ROOT } from './cli-helpers';

interface Msg { id?: number | null; result?: Record<string, unknown> & { content?: { text: string }[]; structuredContent?: Record<string, unknown>; isError?: boolean }; error?: { code: number; message: string } }

let bin = '', dir = '', cleanup = () => {};
let proc: ChildProcessWithoutNullStreams;
const pending = new Map<number, (m: Msg) => void>();
const unsolicited: Msg[] = [];
let nextId = 1;

const send = (o: unknown) => proc.stdin.write(JSON.stringify(o) + '\n');
const request = (method: string, params?: unknown): Promise<Msg> => new Promise((res) => {
  const id = nextId++;
  pending.set(id, res);
  send({ jsonrpc: '2.0', id, method, ...(params !== undefined && { params }) });
});
const call = async (name: string, args: Record<string, unknown>) => (await request('tools/call', { name, arguments: args })).result!;
const data = async (name: string, args: Record<string, unknown>) => {
  const r = await call(name, args);
  expect(r.isError, r.content?.[0]?.text).toBe(false);
  expect(JSON.parse(r.content![0]!.text)).toEqual(r.structuredContent);
  return r.structuredContent as Record<string, any>;
};

beforeAll(() => {
  ({ bin, dir, cleanup } = buildCli());
  mkdirSync(join(dir, 'work'));
  copyFileSync(join(ROOT, 'examples/sales.xlsx'), join(dir, 'work/sales.xlsx'));
  writeFileSync(join(dir, 'secret.csv'), 'a,b\n1,2\n');
  symlinkSync(join(dir, 'secret.csv'), join(dir, 'work/link.csv'));
  proc = spawn(process.execPath, [bin, 'mcp', '--root', join(dir, 'work'), '--max-cells', '5000'], { cwd: join(dir, 'work') });
  let buf = '';
  proc.stdout.setEncoding('utf8');
  proc.stdout.on('data', (chunk: string) => {
    buf += chunk;
    let nl: number;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl);
      buf = buf.slice(nl + 1);
      const m = JSON.parse(line) as Msg;
      const cb = typeof m.id === 'number' ? pending.get(m.id) : undefined;
      if (cb) { pending.delete(m.id as number); cb(m); } else unsolicited.push(m);
    }
  });
}, 120_000);

afterAll(() => { proc?.kill(); cleanup(); });

describe('clerdogrid MCP server', () => {
  it('initializes and negotiates the protocol version', async () => {
    const r = await request('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } });
    expect(r.result).toMatchObject({ protocolVersion: '2025-06-18', capabilities: { tools: {} }, serverInfo: { name: 'clerdogrid' } });
    send({ jsonrpc: '2.0', method: 'notifications/initialized' });
    const other = await request('initialize', { protocolVersion: '1999-01-01', capabilities: {} });
    expect(other.result!.protocolVersion).toBe('2025-11-25');
    expect((await request('ping')).result).toEqual({});
  });

  it('lists about twenty tools with JSON Schema inputs', async () => {
    const r = await request('tools/list');
    const tools = r.result!.tools as { name: string; inputSchema: { type: string; properties: object }; annotations: object }[];
    const names = tools.map((t) => t.name);
    expect(names).toEqual(expect.arrayContaining(['open_workbook', 'list_sheets', 'read_range', 'write_range', 'write_formula', 'recalculate', 'evaluate_formula', 'add_sheet', 'rename_sheet', 'delete_sheet', 'insert_rows', 'delete_rows', 'find', 'describe_function', 'search_functions', 'create_chart', 'set_format', 'save_workbook', 'export_csv', 'workbook_summary']));
    expect(tools.length).toBeGreaterThanOrEqual(20);
    for (const t of tools) { expect(t.inputSchema.type).toBe('object'); expect(t.annotations).toBeTruthy(); }
  });

  it('opens, edits, recalculates and saves a workbook', async () => {
    const open = await data('open_workbook', { path: 'sales.xlsx' });
    expect(open.workbook).toBe('wb1');
    expect(open.sheets[0]).toMatchObject({ name: 'Sales', used: 'A1:E15' });
    const read = await data('read_range', { workbook: 'wb1', range: 'A1:B2' });
    expect(read.values).toEqual([['Month', 'North'], ['Jan', 12400]]);
    expect(read.text[1][1]).toBe('12,400');
    await data('write_range', { workbook: 'wb1', start: 'B2', values: [[100, '200'], ['=B2+C2', null]] });
    expect((await data('read_range', { workbook: 'wb1', range: 'B3', include: ['values'] })).values).toEqual([[300]]);
    const wf = await data('write_formula', { workbook: 'wb1', cell: 'G2:G3', formula: 'E2*2' });
    expect(wf).toMatchObject({ formula: '=E2*2', range: 'G2:G3' });
    expect((await data('read_range', { workbook: 'wb1', range: 'G3', include: ['formulas'] })).formulas).toEqual([['=E3*2']]);
    expect((await data('recalculate', { workbook: 'wb1' })).circular).toEqual([]);
    expect((await data('evaluate_formula', { workbook: 'wb1', formula: '=B2+1' })).value).toBe(101);
    expect((await data('evaluate_formula', { formula: 'SUM(1,2,3)' })).value).toBe(6);
    await data('insert_rows', { workbook: 'wb1', before: 2, count: 2 });
    expect((await data('read_range', { workbook: 'wb1', range: 'E4', include: ['formulas'] })).formulas).toEqual([['=SUM(B4:D4)']]);
    await data('delete_rows', { workbook: 'wb1', start: 2, count: 2 });
    await data('insert_rows', { workbook: 'wb1', before: 'B', axis: 'columns' });
    await data('delete_rows', { workbook: 'wb1', start: 'B', axis: 'columns' });
    expect((await data('add_sheet', { workbook: 'wb1', name: 'Notes' })).sheets).toEqual(['Sales', 'Notes']);
    expect((await data('rename_sheet', { workbook: 'wb1', sheet: 'Notes', name: 'Summary' })).sheets).toEqual(['Sales', 'Summary']);
    await data('write_range', { workbook: 'wb1', start: 'Summary!A1', values: [['=Sales!E2']] });
    await data('rename_sheet', { workbook: 'wb1', sheet: 'Sales', name: 'Data' });
    expect((await data('read_range', { workbook: 'wb1', range: 'Summary!A1' })).formulas).toEqual([['=Data!E2']]);
    const chart = await data('create_chart', { workbook: 'wb1', sheet: 'Data', range: 'A1:D13', type: 'column', title: 'By region' });
    expect(chart.series).toHaveLength(3);
    expect(chart.series[0].categories).toBe('Data!$A$2:$A$13');
    await data('set_format', { workbook: 'wb1', range: 'Data!A1:E1', bold: true, fill: '#E1F3E9', numberFormat: '0.0' });
    const found = await data('find', { workbook: 'wb1', text: 'jan' });
    expect(found.matches[0]).toMatchObject({ sheet: 'Data', cell: 'A2' });
    expect((await data('export_csv', { workbook: 'wb1', sheet: 'Data', range: 'A1:B2' })).csv).toBe('Month,North\r\nJan,100\r\n');
    expect((await data('workbook_summary', { workbook: 'wb1' })).sheets[0].charts).toHaveLength(2);
    expect((await data('validate_workbook', { workbook: 'wb1' })).ok).toBe(true);
    await data('delete_sheet', { workbook: 'wb1', sheet: 'Summary' });
    expect((await data('list_sheets', { workbook: 'wb1' })).sheets.map((s: { name: string }) => s.name)).toEqual(['Data']);
    const saved = await data('save_workbook', { workbook: 'wb1', path: 'out.xlsx' });
    expect(saved.saved).toBe(join(dir, 'work', 'out.xlsx').replace(/^\/private/, ''));
    expect(existsSync(join(dir, 'work/out.xlsx'))).toBe(true);
    const re = await data('open_workbook', { path: 'out.xlsx' });
    expect(re.workbook).toBe('wb1');
    await data('close_workbook', { workbook: 'wb1' });
    const fresh = await data('open_workbook', { path: 'out.xlsx' });
    expect((await data('read_range', { workbook: fresh.workbook, range: 'A1:B2' })).values).toEqual([['Month', 'North'], ['Jan', 100]]);
  });

  it('makes new workbooks from templates', async () => {
    const r = await data('open_workbook', { create: true, template: 'budget', path: 'budget-new.xlsx' });
    expect(r.sheets.length).toBeGreaterThan(0);
    await data('save_workbook', { workbook: r.workbook });
    expect(existsSync(join(dir, 'work/budget-new.xlsx'))).toBe(true);
  });

  it('saves and opens password-protected workbooks', async () => {
    const r = await data('open_workbook', { create: true, template: 'loan' });
    const saved = await call('save_workbook', { workbook: r.workbook, path: 'locked.xlsx', password: 'pw' });
    if (saved.isError) { expect(saved.content![0]!.text).toContain("can't encrypt"); return; }
    await data('close_workbook', { workbook: r.workbook });
    expect((await call('open_workbook', { path: 'locked.xlsx' })).content![0]!.text).toContain('password');
    const back = await data('open_workbook', { path: 'locked.xlsx', password: 'pw' });
    expect(back.sheets[0].name).toBe('Loan');
  });

  it('describes and searches functions', async () => {
    expect((await data('describe_function', { name: 'xlookup' })).signature).toMatch(/^XLOOKUP\(/);
    expect((await data('search_functions', { query: 'sum', limit: 3 })).functions[0].name).toBe('SUM');
  });

  it('keeps files inside the allowed folders', async () => {
    for (const path of ['../secret.csv', join(dir, 'secret.csv'), 'link.csv']) {
      const r = await call('open_workbook', { path });
      expect(r.isError).toBe(true);
      expect(r.content![0]!.text).toContain('outside the folders');
    }
    const w = await call('save_workbook', { workbook: (await data('open_workbook', { create: true })).workbook, path: '../escape.xlsx' });
    expect(w.isError).toBe(true);
    expect(existsSync(join(dir, 'escape.xlsx'))).toBe(false);
  });

  it('checks arguments and limits, and reports errors as tool results', async () => {
    const missing = await call('read_range', { workbook: 'sales.xlsx' });
    expect(missing).toMatchObject({ isError: true });
    expect(missing.content![0]!.text).toContain('range is required');
    expect((await call('read_range', { workbook: 'sales.xlsx', range: 'A1', extra: 1 })).content![0]!.text).toContain("isn't a known argument");
    expect((await call('read_range', { workbook: 'sales.xlsx', range: 'A1:Z1000' })).content![0]!.text).toContain('limit is 5000');
    expect((await call('write_formula', { workbook: 'sales.xlsx', cell: 'A1', formula: '=SUM((' })).isError).toBe(true);
    expect((await call('read_range', { workbook: 'wb999', range: 'A1' })).content![0]!.text).toContain('open_workbook');
    expect((await request('tools/call', { name: 'nope', arguments: {} })).error!.code).toBe(-32602);
    expect((await request('no/such/method')).error!.code).toBe(-32601);
  });

  it('answers malformed input with JSON-RPC errors', async () => {
    proc.stdin.write('this is not json\n');
    await new Promise((r) => setTimeout(r, 200));
    expect(unsolicited.some((m) => m.error?.code === -32700)).toBe(true);
  });
});
