// Model Context Protocol server over stdio: newline-delimited JSON-RPC 2.0 messages.
import { existsSync, readFileSync, realpathSync, statSync } from 'node:fs';
import { basename, dirname, isAbsolute, relative, resolve, sep } from 'node:path';
import { createInterface } from 'node:readline';
import type { GridStore } from '../app/store';
import { addr, colIndex, rangeAddr } from '../model/workbook';
import {
  addSheet, CHART_TYPES, changeRows, createChart, deleteSheet, describeFunction, evaluate, find, fnOut, fromTemplate,
  listFunctions, openStore, parseTarget, readRange, recalculate, renameSheet, setFormat, sheetIndex, summary,
  templateList, validate, valueOut, writeCells, type FormatPatch, type Input,
} from './core';
import { UserError } from './errors';
import { writeFileSafely } from './files';
import { kindOf, sheetText, readWorkbook, writeWorkbook } from './io';
import { writeCsv } from '../formats/csv';

export const PROTOCOL_VERSIONS = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05'] as const;
export const LATEST_PROTOCOL = PROTOCOL_VERSIONS[0];

export interface McpOptions {
  roots: string[];
  readOnly?: boolean;
  maxBytes?: number;
  maxCells?: number;
  /** For tests: where responses go (default stdout) and where requests come from (default stdin). */
  write?: (line: string) => void;
  input?: NodeJS.ReadableStream;
}

type Json = null | boolean | number | string | Json[] | { [k: string]: Json };
type Schema = {
  type?: string | string[]; description?: string; enum?: readonly (string | number)[]; properties?: Record<string, Schema>; required?: string[];
  items?: Schema; minimum?: number; maximum?: number; maxLength?: number; minLength?: number; maxItems?: number; additionalProperties?: boolean; default?: Json;
};

// ---------- Input checking ----------

const typeOf = (v: unknown): string => (v === null ? 'null' : Array.isArray(v) ? 'array' : Number.isInteger(v) ? 'integer' : typeof v);

/** Checks a value against the subset of JSON Schema the tool definitions use. */
export const check = (v: unknown, s: Schema, at = 'arguments'): void => {
  if (s.type) {
    const want = Array.isArray(s.type) ? s.type : [s.type];
    const got = typeOf(v);
    if (!want.some((w) => w === got || (w === 'number' && got === 'integer'))) throw new UserError(`${at} must be ${want.join(' or ')}.`);
  }
  if (s.enum && !s.enum.includes(v as string)) throw new UserError(`${at} must be one of: ${s.enum.join(', ')}.`);
  if (typeof v === 'string') {
    if (s.maxLength !== undefined && v.length > s.maxLength) throw new UserError(`${at} is too long (at most ${s.maxLength} characters).`);
    if (s.minLength !== undefined && v.length < s.minLength) throw new UserError(`${at} can't be empty.`);
  }
  if (typeof v === 'number') {
    if (s.minimum !== undefined && v < s.minimum) throw new UserError(`${at} must be at least ${s.minimum}.`);
    if (s.maximum !== undefined && v > s.maximum) throw new UserError(`${at} must be at most ${s.maximum}.`);
  }
  if (Array.isArray(v)) {
    if (s.maxItems !== undefined && v.length > s.maxItems) throw new UserError(`${at} has too many items (at most ${s.maxItems}).`);
    if (s.items) v.forEach((x, i) => check(x, s.items!, `${at}[${i}]`));
  }
  if (v && typeof v === 'object' && !Array.isArray(v) && s.properties) {
    for (const r of s.required ?? []) if ((v as Record<string, unknown>)[r] === undefined) throw new UserError(`${at}.${r} is required.`);
    for (const [k, x] of Object.entries(v)) {
      const ps = s.properties[k];
      if (!ps) { if (s.additionalProperties === false) throw new UserError(`${at}.${k} isn't a known argument.`); continue; }
      if (x !== undefined) check(x, ps, `${at}.${k}`);
    }
  }
};

// ---------- Tool definitions ----------

const WB: Schema = { type: 'string', minLength: 1, maxLength: 4096, description: 'Workbook id from open_workbook, or a file path (opened on first use).' };
const SHEET: Schema = { type: 'string', maxLength: 64, description: 'Sheet name (default: the active sheet).' };
const RANGE: Schema = { type: 'string', minLength: 1, maxLength: 256, description: 'A1-style cell or range, optionally with a sheet ("Data!A1:C10"), a defined name or a table name.' };
const CELL_VALUE: Schema = { type: ['string', 'number', 'boolean', 'null'] };

interface Tool {
  name: string;
  title: string;
  description: string;
  inputSchema: Schema & { type: 'object' };
  readOnly?: boolean;
  destructive?: boolean;
  run: (a: Record<string, unknown>, ctx: Server) => Promise<Record<string, unknown>> | Record<string, unknown>;
}

const obj = (properties: Record<string, Schema>, required: string[] = []): Schema & { type: 'object' } => ({ type: 'object', properties, required, additionalProperties: false });

const TOOLS: Tool[] = [
  {
    name: 'open_workbook', title: 'Open workbook', readOnly: true,
    description: 'Opens a workbook file (.xlsx, .csv, .tsv, .json and other formats this build reads) or makes a new one from a template, and returns its id for the other tools.',
    inputSchema: obj({
      path: { type: 'string', maxLength: 4096, description: 'File to open. With create, where it will be saved.' },
      create: { type: 'boolean', description: 'Make a new workbook instead of opening a file.' },
      template: { type: 'string', maxLength: 32, description: `Template for a new workbook: ${templateList().map((t) => t.id).join(', ')}.` },
      password: { type: 'string', maxLength: 255, description: 'Password for a password-protected file.' },
    }),
    run: async (a, ctx) => {
      if (a.create || !a.path) {
        const path = typeof a.path === 'string' ? ctx.resolvePath(a.path, true) : null;
        const doc = ctx.add(openStore(fromTemplate(typeof a.template === 'string' ? a.template : 'blank')), path);
        return { workbook: doc.id, path, ...brief(doc.store) };
      }
      const doc = await ctx.openPath(a.path as string, a.password as string | undefined);
      return { workbook: doc.id, path: doc.path, ...brief(doc.store) };
    },
  },
  {
    name: 'list_sheets', title: 'List sheets', readOnly: true,
    description: 'Lists the sheets of an open workbook with their used ranges.',
    inputSchema: obj({ workbook: WB }, ['workbook']),
    run: async (a, ctx) => brief((await ctx.doc(a)).store),
  },
  {
    name: 'workbook_summary', title: 'Workbook summary', readOnly: true,
    description: 'Describes a workbook: sheets, used ranges, counts of cells and formulas, tables, charts, defined names and properties.',
    inputSchema: obj({ workbook: WB }, ['workbook']),
    run: async (a, ctx) => summary((await ctx.doc(a)).store),
  },
  {
    name: 'read_range', title: 'Read range', readOnly: true,
    description: 'Reads cells: values, the text each cell shows (with number formats) and formulas.',
    inputSchema: obj({ workbook: WB, range: RANGE, sheet: SHEET, include: { type: 'array', items: { type: 'string', enum: ['values', 'text', 'formulas'] }, description: 'What to return (default: all three).' } }, ['workbook', 'range']),
    run: async (a, ctx) => {
      const { store } = await ctx.doc(a);
      const data = readRange(store, parseTarget(store, a.range as string, a.sheet as string | undefined), ctx.maxCells);
      const inc = (a.include as string[] | undefined) ?? ['values', 'text', 'formulas'];
      return { sheet: data.sheet, range: data.range, ...(inc.includes('values') && { values: data.values }), ...(inc.includes('text') && { text: data.text }), ...(inc.includes('formulas') && { formulas: data.formulas }) };
    },
  },
  {
    name: 'write_range', title: 'Write range',
    description: 'Writes a block of values starting at a cell. Strings are read as typed input: "=SUM(A1:A3)" is a formula, "12%" a number; start with \' to keep text as it is. null clears a cell.',
    inputSchema: obj({ workbook: WB, start: { ...RANGE, description: 'Top-left cell, e.g. "B2" or "Data!B2".' }, sheet: SHEET, values: { type: 'array', maxItems: 100000, items: { type: 'array', maxItems: 16384, items: CELL_VALUE }, description: 'Rows of cell values.' } }, ['workbook', 'start', 'values']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      const rows = a.values as Input[][];
      const cells = rows.reduce((n, r) => n + r.length, 0);
      if (cells > ctx.maxCells) throw new UserError(`That is ${cells} cells; the limit is ${ctx.maxCells}.`);
      for (const r of rows) for (const v of r) if (typeof v === 'string' && v.length > 32767) throw new UserError('A cell can hold at most 32767 characters.');
      const t = parseTarget(doc.store, a.start as string, a.sheet as string | undefined);
      const area = writeCells(doc.store, t.sheet, t.range.r1, t.range.c1, rows);
      doc.dirty = true;
      return { sheet: doc.store.wb.sheets[t.sheet]!.name, range: rangeAddr(area), written: cells };
    },
  },
  {
    name: 'write_formula', title: 'Write formula',
    description: 'Puts a formula in one cell (or the same formula in every cell of a range, with references adjusted as when filling) and returns the result.',
    inputSchema: obj({ workbook: WB, cell: RANGE, sheet: SHEET, formula: { type: 'string', minLength: 1, maxLength: 8192, description: 'Formula, with or without the leading "=".' } }, ['workbook', 'cell', 'formula']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      const t = parseTarget(doc.store, a.cell as string, a.sheet as string | undefined);
      const f = `=${(a.formula as string).trim().replace(/^=/, '')}`;
      const { shiftFormula } = await import('../formats/formula-xml');
      const h = t.range.r2 - t.range.r1 + 1, w = t.range.c2 - t.range.c1 + 1;
      if (h * w > ctx.maxCells) throw new UserError(`That range has ${h * w} cells; the limit is ${ctx.maxCells}.`);
      const rows: Input[][] = Array.from({ length: h }, (_, i) => Array.from({ length: w }, (_, j) => (i || j ? `=${shiftFormula(f.slice(1), i, j)}` : f)));
      writeCells(doc.store, t.sheet, t.range.r1, t.range.c1, rows);
      doc.dirty = true;
      const data = readRange(doc.store, { sheet: t.sheet, range: { ...t.range, r2: t.range.r1, c2: t.range.c1 } });
      return { sheet: data.sheet, cell: addr(t.range.r1, t.range.c1), range: rangeAddr(t.range), formula: data.formulas[0]![0], value: data.values[0]![0], text: data.text[0]![0] };
    },
  },
  {
    name: 'recalculate', title: 'Recalculate', readOnly: true,
    description: 'Recalculates every formula in the workbook and reports circular references.',
    inputSchema: obj({ workbook: WB }, ['workbook']),
    run: async (a, ctx) => {
      const { store } = await ctx.doc(a);
      return { ...recalculate(store), formulas: summary(store).totals.formulas };
    },
  },
  {
    name: 'evaluate_formula', title: 'Evaluate formula', readOnly: true,
    description: 'Calculates a formula without storing it. With a workbook, references point into it.',
    inputSchema: obj({ formula: { type: 'string', minLength: 1, maxLength: 8192 }, workbook: { ...WB, description: 'Optional workbook id or path.' }, sheet: SHEET }, ['formula']),
    run: async (a, ctx) => {
      const store = a.workbook ? (await ctx.doc(a)).store : openStore(fromTemplate('blank'));
      const v = evaluate(store, a.formula as string, sheetIndex(store, a.sheet as string | undefined));
      return { value: valueOut(v) as Json };
    },
  },
  {
    name: 'add_sheet', title: 'Add sheet',
    description: 'Adds a sheet.',
    inputSchema: obj({ workbook: WB, name: { type: 'string', maxLength: 31 }, position: { type: 'integer', minimum: 0, description: '0-based position (default: last).' } }, ['workbook']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      const name = addSheet(doc.store, a.name as string | undefined, a.position as number | undefined);
      doc.dirty = true;
      return { name, sheets: doc.store.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'rename_sheet', title: 'Rename sheet',
    description: 'Renames a sheet and updates formulas that refer to it.',
    inputSchema: obj({ workbook: WB, sheet: { ...SHEET, description: 'Current name.' }, name: { type: 'string', minLength: 1, maxLength: 31 } }, ['workbook', 'sheet', 'name']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      renameSheet(doc.store, sheetIndex(doc.store, a.sheet as string), a.name as string);
      doc.dirty = true;
      return { sheets: doc.store.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'delete_sheet', title: 'Delete sheet', destructive: true,
    description: 'Deletes a sheet. Formulas that referred to it show #REF!.',
    inputSchema: obj({ workbook: WB, sheet: { ...SHEET, description: 'Sheet to delete.' } }, ['workbook', 'sheet']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      deleteSheet(doc.store, sheetIndex(doc.store, a.sheet as string));
      doc.dirty = true;
      return { sheets: doc.store.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'insert_rows', title: 'Insert rows or columns',
    description: 'Inserts blank rows (or columns) before a row number (or column letter). Formulas, tables and charts move with the cells.',
    inputSchema: obj({ workbook: WB, sheet: SHEET, before: { type: ['integer', 'string'], description: '1-based row number, or a column letter with axis "columns".' }, count: { type: 'integer', minimum: 1, maximum: 100000, default: 1 }, axis: { type: 'string', enum: ['rows', 'columns'], default: 'rows' } }, ['workbook', 'before']),
    run: async (a, ctx) => structure(a, ctx, 1),
  },
  {
    name: 'delete_rows', title: 'Delete rows or columns', destructive: true,
    description: 'Deletes rows (or columns) starting at a row number (or column letter). Formulas that pointed into them show #REF!.',
    inputSchema: obj({ workbook: WB, sheet: SHEET, start: { type: ['integer', 'string'], description: '1-based row number, or a column letter with axis "columns".' }, count: { type: 'integer', minimum: 1, maximum: 100000, default: 1 }, axis: { type: 'string', enum: ['rows', 'columns'], default: 'rows' } }, ['workbook', 'start']),
    run: async (a, ctx) => structure(a, ctx, -1),
  },
  {
    name: 'find', title: 'Find', readOnly: true,
    description: 'Finds cells whose shown text, value or formula contains the given text.',
    inputSchema: obj({ workbook: WB, text: { type: 'string', minLength: 1, maxLength: 1024 }, sheet: SHEET, matchCase: { type: 'boolean' }, wholeCell: { type: 'boolean' }, searchFormulas: { type: 'boolean', default: true }, limit: { type: 'integer', minimum: 1, maximum: 10000, default: 200 } }, ['workbook', 'text']),
    run: async (a, ctx) => {
      const { store } = await ctx.doc(a);
      const matches = find(store, a.text as string, { sheet: a.sheet ? sheetIndex(store, a.sheet as string) : undefined, matchCase: !!a.matchCase, whole: !!a.wholeCell, formulas: a.searchFormulas !== false, limit: (a.limit as number | undefined) ?? 200 });
      return { count: matches.length, matches: matches as unknown as Json };
    },
  },
  {
    name: 'validate_workbook', title: 'Check formulas', readOnly: true,
    description: 'Lists cells showing error values, formulas with #REF!, unknown functions, unreadable formulas and circular references.',
    inputSchema: obj({ workbook: WB, limit: { type: 'integer', minimum: 1, maximum: 10000, default: 500 } }, ['workbook']),
    run: async (a, ctx) => {
      const problems = validate((await ctx.doc(a)).store);
      return { ok: !problems.length, count: problems.length, problems: problems.slice(0, (a.limit as number | undefined) ?? 500) as unknown as Json };
    },
  },
  {
    name: 'describe_function', title: 'Describe function', readOnly: true,
    description: "Shows a worksheet function's signature, arguments and what it returns.",
    inputSchema: obj({ name: { type: 'string', minLength: 1, maxLength: 64 } }, ['name']),
    run: (a) => describeFunction(a.name as string) as unknown as Record<string, unknown>,
  },
  {
    name: 'search_functions', title: 'Search functions', readOnly: true,
    description: 'Searches worksheet functions by name or description.',
    inputSchema: obj({ query: { type: 'string', maxLength: 128 }, category: { type: 'string', maxLength: 32 }, limit: { type: 'integer', minimum: 1, maximum: 1000, default: 25 } }),
    run: (a) => {
      const list = listFunctions(a.query as string | undefined, a.category as string | undefined, (a.limit as number | undefined) ?? 25);
      return { count: list.length, functions: list.map((f) => { const o = fnOut(f); return { name: o.name, category: o.category, signature: o.signature, about: o.about }; }) };
    },
  },
  {
    name: 'create_chart', title: 'Create chart',
    description: 'Adds a chart for a block of data. The first row or column holding text is used for series names and categories.',
    inputSchema: obj({ workbook: WB, range: RANGE, sheet: SHEET, type: { type: 'string', enum: CHART_TYPES, default: 'column' }, title: { type: 'string', maxLength: 255 }, at: { type: 'string', maxLength: 16, description: 'Cell for the top-left corner (default: right of the data).' }, seriesInColumns: { type: 'boolean' } }, ['workbook', 'range']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      const c = createChart(doc.store, parseTarget(doc.store, a.range as string, a.sheet as string | undefined), { type: a.type as string | undefined, title: a.title as string | undefined, at: a.at as string | undefined, seriesInColumns: a.seriesInColumns as boolean | undefined });
      doc.dirty = true;
      return { id: c.id, type: c.type, title: c.title ?? null, at: addr(c.anchor.row, c.anchor.col), series: c.series.map((s) => ({ name: s.name ?? null, values: s.values, categories: s.categories ?? null })) };
    },
  },
  {
    name: 'set_format', title: 'Format cells',
    description: 'Changes number format, font, colours, alignment, wrapping or borders of a range. Only the given settings change.',
    inputSchema: obj({
      workbook: WB, range: RANGE, sheet: SHEET,
      numberFormat: { type: 'string', maxLength: 255, description: 'Number format code, e.g. "#,##0.00", "0%", "yyyy-mm-dd".' },
      bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, strike: { type: 'boolean' },
      fontName: { type: 'string', maxLength: 64 }, fontSize: { type: 'number', minimum: 1, maximum: 409 },
      color: { type: 'string', maxLength: 7, description: 'Text colour "#RRGGBB".' },
      fill: { type: ['string', 'null'], maxLength: 7, description: 'Fill colour "#RRGGBB", or null to remove.' },
      align: { type: 'string', enum: ['general', 'left', 'center', 'right', 'justify'] },
      valign: { type: 'string', enum: ['top', 'center', 'bottom'] },
      wrap: { type: 'boolean' },
      border: { type: 'string', enum: ['none', 'thin', 'medium', 'thick', 'double'], description: 'Border on every side of each cell.' },
    }, ['workbook', 'range']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a);
      const t = parseTarget(doc.store, a.range as string, a.sheet as string | undefined);
      const { workbook: _w, range: _r, sheet: _s, ...patch } = a;
      if (!Object.keys(patch).length) throw new UserError('Give at least one format setting.');
      setFormat(doc.store, t, patch as FormatPatch, ctx.maxCells);
      doc.dirty = true;
      return { sheet: doc.store.wb.sheets[t.sheet]!.name, range: rangeAddr(t.range), applied: Object.keys(patch) };
    },
  },
  {
    name: 'save_workbook', title: 'Save workbook', destructive: true,
    description: 'Saves the workbook to its file, or to a new path. The format follows the extension (.xlsx, .csv, .tsv, .json and others this build writes).',
    inputSchema: obj({ workbook: WB, path: { type: 'string', maxLength: 4096, description: 'Where to save (default: the file it came from).' }, sheet: { ...SHEET, description: 'Sheet for one-sheet formats such as CSV.' }, password: { type: 'string', minLength: 1, maxLength: 255, description: 'Save an .xlsx file protected with this password.' } }, ['workbook']),
    run: async (a, ctx) => {
      if (ctx.readOnly) throw new UserError('This server was started with --read-only, so it can\'t save files.');
      const doc = await ctx.doc(a);
      const target = typeof a.path === 'string' ? ctx.resolvePath(a.path, true) : doc.path;
      if (!target) throw new UserError('This workbook has no file yet. Give a path.');
      const data = await writeWorkbook(doc.store.wb, target, { sheet: a.sheet ? sheetIndex(doc.store, a.sheet as string) : undefined, encryptWith: a.password as string | undefined });
      if (data.length > ctx.maxBytes) throw new UserError(`The saved file would be ${data.length} bytes; the limit is ${ctx.maxBytes}.`);
      writeFileSafely(target, data);
      if (kindOf(target) === 'xlsx') { doc.path = target; doc.dirty = false; }
      return { saved: target, bytes: data.length };
    },
  },
  {
    name: 'export_csv', title: 'Export CSV', readOnly: true,
    description: 'Returns one sheet (or a range) as CSV text, using the text each cell shows unless raw is set.',
    inputSchema: obj({ workbook: WB, sheet: SHEET, range: { ...RANGE, description: 'Range to export (default: the used range).' }, raw: { type: 'boolean' }, formulas: { type: 'boolean' } }, ['workbook']),
    run: async (a, ctx) => {
      const { store } = await ctx.doc(a);
      const t = a.range ? parseTarget(store, a.range as string, a.sheet as string | undefined) : { sheet: sheetIndex(store, a.sheet as string | undefined), range: undefined };
      if (t.range && (t.range.r2 - t.range.r1 + 1) * (t.range.c2 - t.range.c1 + 1) > ctx.maxCells) throw new UserError(`That range is larger than ${ctx.maxCells} cells.`);
      const rows = sheetText(store.wb, store.engine, { sheet: t.sheet, range: t.range, raw: !!a.raw, formulas: !!a.formulas });
      if (rows.length * (rows[0]?.length ?? 0) > ctx.maxCells) throw new UserError(`That sheet is larger than ${ctx.maxCells} cells. Give a range.`);
      return { sheet: store.wb.sheets[t.sheet]!.name, rows: rows.length, csv: writeCsv(rows) };
    },
  },
  {
    name: 'close_workbook', title: 'Close workbook', readOnly: true,
    description: 'Forgets an open workbook. Unsaved changes are lost.',
    inputSchema: obj({ workbook: WB }, ['workbook']),
    run: async (a, ctx) => {
      const doc = await ctx.doc(a, false);
      ctx.docs.delete(doc.id);
      return { closed: doc.id, unsavedChanges: doc.dirty };
    },
  },
];

const brief = (store: GridStore) => {
  const s = summary(store);
  return { activeSheet: s.activeSheet, sheets: s.sheets.map((x) => ({ name: x.name, used: x.used, rows: x.rows, cols: x.cols, formulas: x.formulas, ...(x.state !== 'visible' && { state: x.state }) })) };
};

const structure = async (a: Record<string, unknown>, ctx: Server, sign: 1 | -1): Promise<Record<string, unknown>> => {
  const doc = await ctx.doc(a);
  const axis = a.axis === 'columns' ? 'col' : 'row';
  const pos = (sign > 0 ? a.before : a.start) as number | string;
  let at: number;
  if (axis === 'row') {
    if (typeof pos !== 'number' && !/^\d+$/.test(String(pos))) throw new UserError('Give a row number, starting at 1.');
    at = Number(pos) - 1;
  } else {
    at = typeof pos === 'number' ? pos - 1 : /^[A-Za-z]{1,3}$/.test(pos) ? colIndex(pos) : NaN;
    if (Number.isNaN(at)) throw new UserError('Give a column letter, or a column number starting at 1.');
  }
  const count = (a.count as number | undefined) ?? 1;
  const sheet = sheetIndex(doc.store, a.sheet as string | undefined);
  changeRows(doc.store, sheet, axis, at, sign * count);
  doc.dirty = true;
  return { sheet: doc.store.wb.sheets[sheet]!.name, axis: axis === 'row' ? 'rows' : 'columns', at: axis === 'row' ? at + 1 : addr(0, at).replace(/\d+$/, ''), count, inserted: sign > 0 };
};

// ---------- Server ----------

interface Doc { id: string; store: GridStore; path: string | null; dirty: boolean }

export class Server {
  readonly docs = new Map<string, Doc>();
  private next = 1;
  readonly roots: string[];
  readonly readOnly: boolean;
  readonly maxBytes: number;
  readonly maxCells: number;
  protocolVersion: string = LATEST_PROTOCOL;

  constructor(o: McpOptions) {
    if (!o.roots.length) throw new UserError('Give at least one --root folder.');
    this.roots = o.roots.map((r) => {
      const abs = resolve(r);
      if (!existsSync(abs) || !statSync(abs).isDirectory()) throw new UserError(`--root ${r} isn't a folder.`);
      return realpathSync(abs);
    });
    this.readOnly = !!o.readOnly;
    this.maxBytes = o.maxBytes ?? 100 * 1024 * 1024;
    this.maxCells = o.maxCells ?? 100_000;
  }

  private inside(p: string): boolean {
    return this.roots.some((r) => { const rel = relative(r, p); return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel) && !rel.startsWith(sep)); });
  }

  /** Resolves a path against the first root and refuses anything outside the roots, following symlinks. */
  resolvePath(p: string, forWrite = false): string {
    if (p.includes('\0')) throw new UserError('That path isn\'t valid.');
    const abs = resolve(this.roots[0]!, p);
    if (!kindOf(abs)) throw new UserError(`Can't tell the format of "${basename(abs)}". Use .xlsx, .csv, .tsv or .json.`);
    // Resolve symlinks through the deepest folder that exists, and check the roots before saying anything about the file.
    let base = abs;
    const rest: string[] = [];
    while (!existsSync(base)) { rest.unshift(basename(base)); const up = dirname(base); if (up === base) break; base = up; }
    const real = resolve(realpathSync(base), ...rest);
    if (!this.inside(real)) throw new UserError(`${p} is outside the folders this server may use (${this.roots.join(', ')}).`);
    if (rest.length) {
      if (!forWrite) throw new UserError(`${p} doesn't exist.`);
      if (rest.length > 1) throw new UserError(`The folder for ${p} doesn't exist.`);
    } else if (statSync(real).isDirectory()) throw new UserError(`${p} is a folder.`);
    return real;
  }

  add(store: GridStore, path: string | null): Doc {
    const doc: Doc = { id: `wb${this.next++}`, store, path, dirty: false };
    this.docs.set(doc.id, doc);
    return doc;
  }

  async openPath(p: string, password?: string): Promise<Doc> {
    const real = this.resolvePath(p);
    for (const d of this.docs.values()) if (d.path === real) return d;
    const size = statSync(real).size;
    if (size > this.maxBytes) throw new UserError(`${p} is ${size} bytes; the limit is ${this.maxBytes}.`);
    return this.add(openStore(await readWorkbook(basename(real), new Uint8Array(readFileSync(real)), { password })), real);
  }

  async doc(a: Record<string, unknown>, autoOpen = true): Promise<Doc> {
    const id = a.workbook as string;
    const d = this.docs.get(id);
    if (d) return d;
    if (autoOpen && /[./\\]/.test(id)) return this.openPath(id);
    throw new UserError(`No open workbook "${id}". Call open_workbook first.`);
  }

  // ---------- JSON-RPC ----------

  async handle(msg: unknown): Promise<Record<string, unknown> | null> {
    if (!msg || typeof msg !== 'object' || Array.isArray(msg) || (msg as { jsonrpc?: unknown }).jsonrpc !== '2.0') {
      return rpcError(null, -32600, 'Invalid request');
    }
    const { id, method, params } = msg as { id?: string | number | null; method?: unknown; params?: unknown };
    const isNotification = id === undefined;
    if (typeof method !== 'string') return isNotification ? null : rpcError(id ?? null, -32600, 'Invalid request');
    try {
      const result = await this.dispatch(method, (params ?? {}) as Record<string, unknown>);
      return isNotification || result === undefined ? null : { jsonrpc: '2.0', id, result };
    } catch (e) {
      if (isNotification) return null;
      if (e instanceof RpcError) return rpcError(id ?? null, e.code, e.message);
      return rpcError(id ?? null, -32603, e instanceof Error ? e.message : 'Internal error');
    }
  }

  private async dispatch(method: string, params: Record<string, unknown>): Promise<unknown> {
    switch (method) {
      case 'initialize': {
        const asked = params.protocolVersion;
        this.protocolVersion = typeof asked === 'string' && (PROTOCOL_VERSIONS as readonly string[]).includes(asked) ? asked : LATEST_PROTOCOL;
        return {
          protocolVersion: this.protocolVersion,
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: 'clerdogrid', title: 'ClerdoGrid', version: __APP_VERSION__ },
          instructions: `Reads and edits spreadsheet workbooks. Call open_workbook with a file path (relative paths start at ${this.roots[0]}) to get a workbook id, then pass it to the other tools. Changes stay in memory until save_workbook.${this.readOnly ? ' Saving is turned off.' : ''}`,
        };
      }
      case 'notifications/initialized':
      case 'notifications/cancelled':
        return undefined;
      case 'ping':
        return {};
      case 'tools/list':
        return {
          tools: TOOLS.filter((t) => !(this.readOnly && t.name === 'save_workbook')).map((t) => ({
            name: t.name, title: t.title, description: t.description, inputSchema: t.inputSchema,
            annotations: { title: t.title, readOnlyHint: !!t.readOnly, destructiveHint: !!t.destructive, idempotentHint: !!t.readOnly, openWorldHint: false },
          })),
        };
      case 'tools/call': {
        const name = params.name;
        const tool = TOOLS.find((t) => t.name === name);
        if (!tool || (this.readOnly && tool.name === 'save_workbook')) throw new RpcError(-32602, `Unknown tool: ${String(name)}`);
        const args = (params.arguments ?? {}) as Record<string, unknown>;
        try {
          check(args, tool.inputSchema);
          const out = await tool.run(args, this);
          return { content: [{ type: 'text', text: JSON.stringify(out, null, 2) }], structuredContent: out, isError: false };
        } catch (e) {
          // Tool failures go back to the model as results it can read and correct, not protocol errors.
          const message = e instanceof UserError ? e.message : `Unexpected error: ${e instanceof Error ? e.message : String(e)}`;
          return { content: [{ type: 'text', text: message }], isError: true };
        }
      }
      default:
        if (method.startsWith('notifications/')) return undefined;
        throw new RpcError(-32601, `Method not found: ${method}`);
    }
  }
}

class RpcError extends Error {
  constructor(readonly code: number, message: string) { super(message); }
}

const rpcError = (id: string | number | null, code: number, message: string) => ({ jsonrpc: '2.0', id, error: { code, message } });

export const toolNames = (): string[] => TOOLS.map((t) => t.name);

/** Serves until standard input closes. Logs go to standard error so standard output stays protocol-only. */
export const runMcp = async (o: McpOptions): Promise<void> => {
  const server = new Server(o);
  const write = o.write ?? ((line: string) => { process.stdout.write(line + '\n'); });
  const rl = createInterface({ input: o.input ?? process.stdin, crlfDelay: Infinity });
  let chain = Promise.resolve();
  const MAX_LINE = Math.max(16 * 1024 * 1024, server.maxCells * 64);
  process.stderr.write(`clerdogrid MCP server ${__APP_VERSION__} ready; files under ${server.roots.join(', ')}\n`);
  for await (const line of rl) {
    if (!line.trim()) continue;
    if (line.length > MAX_LINE) { write(JSON.stringify(rpcError(null, -32600, 'Message too large'))); continue; }
    let msg: unknown;
    try { msg = JSON.parse(line); } catch { chain = chain.then(() => write(JSON.stringify(rpcError(null, -32700, 'Parse error')))); continue; }
    // Requests run one at a time, in order, because they share open workbooks.
    chain = chain.then(async () => {
      if (Array.isArray(msg)) {
        const out: Record<string, unknown>[] = [];
        for (const m of msg) { const r = await server.handle(m); if (r) out.push(r); }
        if (out.length) write(JSON.stringify(out));
      } else {
        const r = await server.handle(msg);
        if (r) write(JSON.stringify(r));
      }
    });
  }
  await chain;
};
