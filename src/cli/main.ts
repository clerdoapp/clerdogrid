import { existsSync, readFileSync, statSync } from 'node:fs';
import { basename } from 'node:path';
import { writeFileSafely } from './files';
import { writeCsv } from '../formats/csv';
import { addr } from '../model/workbook';
import type { GridStore } from '../app/store';
import { allowOnly, bool, int, parseArgs, str, type Parsed } from './args';
import {
  describeFunction, diff, evaluate, find, fromTemplate, listFunctions, openStore, parseTarget, readRange,
  recalculate, sheetIndex, summary, templateList, validate, valueOut, valueText, writeCells,
} from './core';
import { EXIT, UserError } from './errors';
import { COMMAND_HELP, USAGE } from './help';
import { canEncrypt, kindOf, MAX_FILE_BYTES, readWorkbook, sheetText, supportedKinds, writeWorkbook, type ExportOptions } from './io';
import { runMcp } from './mcp';

export const VERSION = __APP_VERSION__;

export interface Io { out: (s: string) => void; err: (s: string) => void }
const stdio: Io = { out: (s) => process.stdout.write(s), err: (s) => process.stderr.write(s) };

const VALUED = ['password', 'encrypt', 'sheet', 'out', 'file', 'range', 'search', 'category', 'limit', 'root', 'max-bytes', 'max-cells', 'template'];
const ALIASES = { o: 'out', s: 'sheet', h: 'help', V: 'version', f: 'file', r: 'range' };

const readFile = (path: string): Uint8Array => {
  if (!existsSync(path)) throw new UserError(`${path} doesn't exist.`, EXIT.file);
  const st = statSync(path);
  if (st.isDirectory()) throw new UserError(`${path} is a folder, not a file.`, EXIT.file);
  if (st.size > MAX_FILE_BYTES) throw new UserError(`${path} is too large to open.`, EXIT.file);
  return new Uint8Array(readFileSync(path));
};

let password: string | undefined;
const load = async (path: string): Promise<GridStore> => openStore(await readWorkbook(basename(path), readFile(path), { password }));

const save = async (store: GridStore, path: string, o: ExportOptions = {}): Promise<void> => {
  writeFileSafely(path, await writeWorkbook(store.wb, path, o));
};

const need = (p: Parsed, n: number, usage: string): string[] => {
  if (p.positional.length < n) throw new UserError(`Usage: ${usage}`);
  if (p.positional.length > n) throw new UserError(`Too many arguments. Usage: ${usage}`);
  return p.positional;
};

const json = (io: Io, v: unknown): void => io.out(JSON.stringify(v, null, 2) + '\n');

type Handler = (p: Parsed, io: Io) => Promise<number>;

const commands: Record<string, Handler> = {
  async info(p, io) {
    allowOnly(p, ['json']);
    const [file] = need(p, 1, 'clerdogrid info <file>');
    const s = summary(await load(file!));
    if (bool(p, 'json')) { json(io, s); return EXIT.ok; }
    const lines = [`${basename(file!)}: ${s.totals.sheets} sheet${s.totals.sheets === 1 ? '' : 's'}, ${s.totals.cells} cells, ${s.totals.formulas} formulas`];
    if (s.properties.title) lines.push(`Title: ${s.properties.title}`);
    for (const sh of s.sheets) {
      lines.push('', `Sheet "${sh.name}"${sh.state !== 'visible' ? ` (${sh.state})` : ''}${sh.name === s.activeSheet ? ' (active)' : ''}`);
      lines.push(`  Used range: ${sh.used ?? 'empty'}  (${sh.rows} rows x ${sh.cols} columns, ${sh.cells} cells, ${sh.formulas} formulas)`);
      if (sh.frozen) lines.push(`  Frozen: ${sh.frozen.rows} rows, ${sh.frozen.cols} columns`);
      for (const t of sh.tables) lines.push(`  Table ${t.name}: ${t.range} [${t.columns.join(', ')}]`);
      for (const c of sh.charts) lines.push(`  Chart ${c.type}${c.title ? ` "${c.title}"` : ''} at ${c.at}, ${c.series} series`);
      const extra = [sh.pictures && `${sh.pictures} pictures`, sh.merges && `${sh.merges} merged ranges`, sh.conditionalFormats && `${sh.conditionalFormats} conditional formats`, sh.validations && `${sh.validations} validation rules`, sh.protected && 'protected'].filter(Boolean);
      if (extra.length) lines.push(`  ${extra.join(', ')}`);
    }
    if (s.names.length) {
      lines.push('', 'Defined names');
      for (const n of s.names) lines.push(`  ${n.name}${n.scope ? ` (${n.scope})` : ''} ${n.refersTo}`);
    }
    io.out(lines.join('\n') + '\n');
    return EXIT.ok;
  },

  async get(p, io) {
    allowOnly(p, ['sheet', 'json', 'raw', 'formulas']);
    const [file, ref] = need(p, 2, 'clerdogrid get <file> <ref|range>');
    const store = await load(file!);
    const t = parseTarget(store, ref!, str(p, 'sheet'));
    const data = readRange(store, t, 5_000_000);
    if (bool(p, 'json')) { json(io, data); return EXIT.ok; }
    const rows = data.values.map((r, i) => r.map((v, j) => (bool(p, 'formulas') && data.formulas[i]![j]) || (bool(p, 'raw') ? (v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)) : data.text[i]![j]!)));
    io.out(rows.map((r) => r.join('\t')).join('\n') + '\n');
    return EXIT.ok;
  },

  async set(p, io) {
    allowOnly(p, ['sheet', 'out', 'text']);
    const [file, ref, value] = need(p, 3, 'clerdogrid set <file> <ref> <value|=formula>');
    const store = await load(file!);
    const t = parseTarget(store, ref!, str(p, 'sheet'));
    if (t.range.r1 !== t.range.r2 || t.range.c1 !== t.range.c2) throw new UserError('set changes one cell. Give a single cell like B4.');
    const input = bool(p, 'text') ? `'${value}` : value!;
    writeCells(store, t.sheet, t.range.r1, t.range.c1, [[input]]);
    const out = str(p, 'out') ?? file!;
    await save(store, out, { sheet: t.sheet });
    const shown = readRange(store, t);
    io.out(`${store.wb.sheets[t.sheet]!.name}!${addr(t.range.r1, t.range.c1)} = ${shown.text[0]![0]}${shown.formulas[0]![0] ? `  (${shown.formulas[0]![0]})` : ''}\n`);
    return EXIT.ok;
  },

  async calc(p, io) {
    allowOnly(p, ['out', 'json']);
    const [file] = need(p, 1, 'clerdogrid calc <file> [-o out]');
    const store = await load(file!);
    const { circular } = recalculate(store);
    const out = str(p, 'out') ?? file!;
    await save(store, out);
    const formulas = summary(store).totals.formulas;
    if (bool(p, 'json')) json(io, { saved: out, formulas, circular });
    else {
      io.out(`Recalculated ${formulas} formula${formulas === 1 ? '' : 's'} and saved ${out}\n`);
      if (circular.length) io.err(`Circular references: ${circular.slice(0, 20).join(', ')}${circular.length > 20 ? ` and ${circular.length - 20} more` : ''}\n`);
    }
    return EXIT.ok;
  },

  async eval(p, io) {
    allowOnly(p, ['file', 'sheet', 'json']);
    const [formula] = need(p, 1, 'clerdogrid eval "<formula>" [--file workbook]');
    const file = str(p, 'file');
    const store = file ? await load(file) : openStore((await import('../model/workbook')).newWorkbook());
    const v = evaluate(store, formula!, sheetIndex(store, str(p, 'sheet')));
    if (bool(p, 'json')) json(io, { formula: `=${formula!.replace(/^=/, '')}`, value: valueOut(v) });
    else io.out(valueText(v) + '\n');
    return EXIT.ok;
  },

  async convert(p, io) {
    allowOnly(p, ['sheet', 'range', 'raw', 'formulas', 'force', 'formats', 'encrypt']);
    if (bool(p, 'formats')) {
      const k = await supportedKinds();
      io.out(`Reads:  ${k.read.join(', ')}\nWrites: ${k.write.join(', ')}\n${(await canEncrypt()) ? 'Password-protected .xlsx: read with --password, write with --encrypt\n' : ''}`);
      return EXIT.ok;
    }
    const [input, output] = need(p, 2, 'clerdogrid convert <in> <out>');
    if (!kindOf(output!)) throw new UserError(`Can't tell the format from "${output}". Run "clerdogrid convert --formats".`);
    if (existsSync(output!) && !bool(p, 'force')) throw new UserError(`${output} already exists. Add --force to replace it.`, EXIT.file);
    const store = await load(input!);
    const sheetName = str(p, 'sheet');
    const sheet = sheetName !== undefined ? sheetIndex(store, sheetName) : undefined;
    const range = str(p, 'range') ? parseTarget(store, str(p, 'range')!, sheetName) : undefined;
    await save(store, output!, { sheet: range?.sheet ?? sheet, range: range?.range, raw: bool(p, 'raw'), formulas: bool(p, 'formulas'), encryptWith: str(p, 'encrypt') });
    io.err(`Wrote ${output}\n`);
    return EXIT.ok;
  },

  async csv(p, io) {
    allowOnly(p, ['sheet', 'range', 'raw', 'formulas', 'tsv']);
    const [file] = need(p, 1, 'clerdogrid csv <file>');
    const store = await load(file!);
    const t = str(p, 'range') ? parseTarget(store, str(p, 'range')!, str(p, 'sheet')) : { sheet: sheetIndex(store, str(p, 'sheet')), range: undefined };
    const rows = sheetText(store.wb, store.engine, { sheet: t.sheet, range: t.range, raw: bool(p, 'raw'), formulas: bool(p, 'formulas') });
    io.out(bool(p, 'tsv') ? rows.map((r) => r.map((f) => f.replace(/[\t\r\n]+/g, ' ')).join('\t')).join('\n') + (rows.length ? '\n' : '') : writeCsv(rows));
    return EXIT.ok;
  },

  async find(p, io) {
    allowOnly(p, ['sheet', 'match-case', 'whole', 'no-formulas', 'limit', 'json']);
    const [file, text] = need(p, 2, 'clerdogrid find <file> <text>');
    const store = await load(file!);
    const hits = find(store, text!, { sheet: str(p, 'sheet') !== undefined ? sheetIndex(store, str(p, 'sheet')) : undefined, matchCase: bool(p, 'match-case'), whole: bool(p, 'whole'), formulas: !bool(p, 'no-formulas'), limit: int(p, 'limit', 1) });
    if (bool(p, 'json')) json(io, hits);
    else for (const h of hits) io.out(`${h.sheet}!${h.cell}\t${h.text}${h.formula ? `\t${h.formula}` : ''}\n`);
    return hits.length ? EXIT.ok : EXIT.found;
  },

  async validate(p, io) {
    allowOnly(p, ['json']);
    const [file] = need(p, 1, 'clerdogrid validate <file>');
    const problems = validate(await load(file!));
    if (bool(p, 'json')) json(io, { ok: !problems.length, problems });
    else if (!problems.length) io.out(`${basename(file!)}: no problems found\n`);
    else {
      for (const x of problems) io.out(`${x.sheet}!${x.cell}\t${x.kind}\t${x.detail}${x.formula ? `\t${x.formula}` : ''}\n`);
      io.err(`${problems.length} problem${problems.length === 1 ? '' : 's'} found\n`);
    }
    return problems.length ? EXIT.found : EXIT.ok;
  },

  async diff(p, io) {
    allowOnly(p, ['values-only', 'limit', 'json']);
    const [a, b] = need(p, 2, 'clerdogrid diff <a> <b>');
    const d = diff(await load(a!), await load(b!), { valuesOnly: bool(p, 'values-only'), limit: int(p, 'limit', 1) });
    if (bool(p, 'json')) json(io, { same: !d.length, differences: d });
    else {
      for (const x of d) {
        if (x.kind === 'sheetAdded' || x.kind === 'sheetRemoved') io.out(`${x.kind === 'sheetAdded' ? '+' : '-'} sheet "${x.sheet}"\n`);
        else if (x.kind === 'added') io.out(`+ ${x.sheet}!${x.cell}\t${x.b}\n`);
        else if (x.kind === 'removed') io.out(`- ${x.sheet}!${x.cell}\t${x.a}\n`);
        else io.out(`~ ${x.sheet}!${x.cell}\t${x.a}\t->\t${x.b}\n`);
      }
      if (!d.length) io.out('No differences\n');
    }
    return d.length ? EXIT.found : EXIT.ok;
  },

  async new(p, io) {
    allowOnly(p, ['force', 'list', 'template', 'json']);
    if (bool(p, 'list')) {
      const list = templateList();
      if (bool(p, 'json')) json(io, list);
      else for (const t of list) io.out(`${t.id.padEnd(12)}${t.name} - ${t.about}\n`);
      return EXIT.ok;
    }
    if (p.positional.length < 1 || p.positional.length > 2) throw new UserError('Usage: clerdogrid new <out.xlsx> [template]');
    const [out, tpl] = p.positional;
    if (existsSync(out!) && !bool(p, 'force')) throw new UserError(`${out} already exists. Add --force to replace it.`, EXIT.file);
    const wb = fromTemplate(tpl ?? str(p, 'template') ?? 'blank');
    writeFileSafely(out!, await writeWorkbook(wb, out!));
    io.err(`Wrote ${out}\n`);
    return EXIT.ok;
  },

  async functions(p, io) {
    allowOnly(p, ['search', 'category', 'json', 'limit']);
    const q = str(p, 'search') ?? p.positional[0];
    const list = listFunctions(q, str(p, 'category'), int(p, 'limit', 1));
    if (bool(p, 'json')) json(io, list.map((f) => ({ name: f.name, category: f.category, about: f.about })));
    else {
      const w = Math.max(8, ...list.map((f) => f.name.length)) + 2;
      const cw = Math.max(8, ...list.map((f) => f.category.length)) + 2;
      for (const f of list) io.out(`${f.name.padEnd(w)}${f.category.padEnd(cw)}${f.about}\n`);
    }
    return list.length ? EXIT.ok : EXIT.found;
  },

  async describe(p, io) {
    allowOnly(p, ['json']);
    const [name] = need(p, 1, 'clerdogrid describe <FUNCTION>');
    const f = describeFunction(name!);
    if (bool(p, 'json')) { json(io, f); return EXIT.ok; }
    const lines = [f.signature, '', f.about, '', `Category: ${f.category}`];
    if (f.args.length) {
      lines.push('', 'Arguments');
      const w = Math.max(...f.args.map((a) => a.name.length)) + 2;
      for (const a of f.args) lines.push(`  ${a.name.padEnd(w)}${a.about}${a.optional ? ' (optional)' : ''}${a.repeats ? ' (can repeat)' : ''}`);
    }
    if (f.returns) lines.push('', `Returns: ${f.returns}`);
    io.out(lines.join('\n') + '\n');
    return EXIT.ok;
  },

  async mcp(p) {
    allowOnly(p, ['root', 'read-only', 'max-bytes', 'max-cells']);
    if (p.positional.length) throw new UserError('Usage: clerdogrid mcp [--root dir]...');
    await runMcp({
      roots: p.lists.get('root') ?? [process.cwd()],
      readOnly: bool(p, 'read-only'),
      maxBytes: int(p, 'max-bytes', 1024, MAX_FILE_BYTES),
      maxCells: int(p, 'max-cells', 1, 5_000_000),
    });
    return EXIT.ok;
  },
};

export const main = async (argv: string[], io: Io = stdio): Promise<number> => {
  try {
    const p = parseArgs(argv, VALUED, ALIASES);
    const [cmd, ...rest] = p.positional;
    if (p.flags.has('version') && !cmd) { io.out(`clerdogrid ${VERSION}\n`); return EXIT.ok; }
    if (!cmd || cmd === 'help') {
      const topic = rest[0];
      if (topic && !COMMAND_HELP[topic]) throw new UserError(`Unknown command "${topic}".`);
      const asked = !!cmd || p.flags.has('help');
      (asked ? io.out : io.err)((topic ? COMMAND_HELP[topic]! : USAGE) + '\n');
      return asked ? EXIT.ok : EXIT.usage;
    }
    if (cmd === 'version') { io.out(`clerdogrid ${VERSION}\n`); return EXIT.ok; }
    const handler = Object.hasOwn(commands, cmd) ? commands[cmd] : undefined;
    if (!handler) throw new UserError(`Unknown command "${cmd}". Run "clerdogrid help" for the list.`);
    if (p.flags.has('help')) { io.out(COMMAND_HELP[cmd]! + '\n'); return EXIT.ok; }
    password = str(p, 'password') ?? (process.env.CLERDOGRID_PASSWORD || undefined);
    p.flags.delete('password');
    return await handler({ ...p, positional: rest }, io);
  } catch (e) {
    if (e instanceof UserError) { io.err(`clerdogrid: ${e.message}\n`); return e.exitCode; }
    if ((e as NodeJS.ErrnoException)?.code === 'EPIPE') return EXIT.ok;
    io.err(`clerdogrid: unexpected error: ${e instanceof Error ? e.message : String(e)}\n`);
    if (process.env.CLERDOGRID_DEBUG && e instanceof Error) io.err(`${e.stack}\n`);
    return EXIT.internal;
  }
};

declare const __CLI_BUNDLE__: boolean | undefined;
// Runs only in the bundled tool (npm run cli), not when tests import this file.
if (typeof __CLI_BUNDLE__ !== 'undefined' && __CLI_BUNDLE__) {
  process.stdout.on('error', (e: NodeJS.ErrnoException) => { if (e.code === 'EPIPE') process.exit(0); });
  void main(process.argv.slice(2)).then((code) => { process.exitCode = code; });
}

