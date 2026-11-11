import type { GridStore } from '../../app/store';
import { applyOps, snapshot } from './apply';
import { workerSource, type RunResult } from './runtime';
import { N, T, Tr } from '../../ui/i18n';

export const TIME_LIMIT_MS = 10_000;
const MAX_CODE = 1_000_000;

/** Turns a line in the worker file into a line in the person's script, or undefined when it's outside it. */
const userLine = (absolute: number | undefined, start: number, codeLines: number): number | undefined => {
  if (!absolute) return undefined;
  const l = absolute > start ? absolute - start : absolute;
  return l >= 1 && l <= codeLines ? l : undefined;
};

/** Messages the script runtime throws (in English, inside the worker), shown in the interface language. */
const RUNTIME_MESSAGES = [
  N('{name} isn\'t available in scripts.'), N('Building code from text isn\'t available in scripts.'), N('"{text}" isn\'t a valid range. Use an address such as "A1:C3".'),
  N('"{text}" isn\'t a color. Use "#RRGGBB".'), N('The script made too many changes.'), N('That range is too large to read at once.'),
  N('Pass a two-dimensional array, for example [[1, 2], [3, 4]].'), N('The array is {rows}×{cols} but the range {range} is {height}×{width}.'),
  N('getCell needs whole numbers from 0.'), N('Alignment must be "left", "center", "right" or "general".'),
  N('sort needs a column: a number counted from 0 inside the range, or a letter such as "B".'), N('Column {col} is outside the range {range}.'),
  N('"{name}" can\'t be used as a sheet name.'), N('Another sheet is already called "{name}".'), N('getCell needs whole numbers from 0, for example getCell(0, 0) for A1.'),
  N('deleteRows needs a row index from 0 and a count of at least 1.'), N('A table needs a header row and at least one row of data.'), N('postMessage isn\'t available in scripts.'),
  N('Add a function called main(workbook).'),
];
const runtimePatterns = RUNTIME_MESSAGES.map((m) => {
  const names: string[] = [];
  const re = new RegExp(`^${m.replace(/[.*+?^$()|[\]\\]/g, '\\$&').replace(/\\?\{(\w+)\\?\}|\{(\w+)\}/g, (_x, a: string, b: string) => { names.push(a ?? b); return '(.+?)'; })}$`);
  return { m, re, names };
});
/** An English runtime message in the interface language (other text comes back as it is). */
export const localizeRuntime = (msg: string): string => {
  for (const p of runtimePatterns) {
    const hit = p.re.exec(msg);
    if (hit) return Tr(p.m, Object.fromEntries(p.names.map((n, i) => [n, hit[i + 1]!])));
  }
  return msg;
};

/**
 * Runs a script in a throwaway Web Worker made from a Blob: no DOM, network globals replaced with
 * functions that throw, and stopped after the time limit. The worker only ever sees a copy of the cells.
 */
export const runInWorker = (store: GridStore, code: string, limitMs = TIME_LIMIT_MS): Promise<RunResult> => {
  if (code.length > MAX_CODE) return Promise.resolve({ ok: false, ops: [], logs: [], error: T('The script is too long.') });
  if (typeof Worker === 'undefined') return Promise.resolve({ ok: false, ops: [], logs: [], error: T('Scripts need a browser that supports workers.') });
  const { source, userStart } = workerSource(code);
  const lines = code.split('\n').length;
  const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
  return new Promise<RunResult>((resolve) => {
    let done = false;
    let worker: Worker;
    const finish = (r: RunResult) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      worker?.terminate();
      URL.revokeObjectURL(url);
      resolve(r);
    };
    const timer = setTimeout(() => finish({ ok: false, ops: [], logs: [], error: T('The script ran for more than {p0} seconds and was stopped.', { p0: Math.round(limitMs / 1000) }) }), limitMs);
    try {
      worker = new Worker(url, { name: 'clerdogrid-script' });
    } catch (e) {
      finish({ ok: false, ops: [], logs: [], error: T('The script could not start: {p0}', { p0: e instanceof Error ? e.message : String(e) }) });
      return;
    }
    worker.onmessage = (e: MessageEvent<RunResult>) => {
      const r = e.data;
      if (!r || typeof r !== 'object' || !Array.isArray(r.ops) || !Array.isArray(r.logs)) { finish({ ok: false, ops: [], logs: [], error: T('The script returned something unexpected.') }); return; }
      finish({
        ok: !!r.ok, ops: r.ops, logs: r.logs.filter((x) => typeof x === 'string'),
        error: typeof r.error === 'string' ? localizeRuntime(r.error.replace(/^Error:\s+/, '')) : undefined,
        line: typeof r.line === 'number' ? userLine(r.line, 0, lines) : undefined,
      });
    };
    worker.onerror = (e: ErrorEvent) => {
      e.preventDefault();
      const msg = localizeRuntime((e.message || T('The script has an error.')).replace(/^Uncaught\s+/, '').replace(/^Error:\s+/, ''));
      finish({ ok: false, ops: [], logs: [], error: msg, line: userLine(e.lineno, userStart, lines) });
    };
    worker.postMessage({ snapshot: snapshot(store) });
  });
};

export interface ScriptOutcome { ok: boolean; logs: string[]; message: string; line?: number; changes: number }

/** Runs a script and applies its edits as one undo step. Never called when a file opens. */
export const runScript = async (store: GridStore, code: string, _name = 'script'): Promise<ScriptOutcome> => {
  const r = await runInWorker(store, code);
  if (!r.ok) {
    const where = r.line ? ` ${T('(line {line})', { line: r.line })}` : '';
    return { ok: false, logs: r.logs, message: `${r.error ?? T('The script failed.')}${where}`, line: r.line, changes: 0 };
  }
  const err = applyOps(store, r.ops, N('Run script'));
  if (err) return { ok: false, logs: r.logs, message: err, changes: 0 };
  const edits = r.ops.filter((o) => o.t !== 'select' && o.t !== 'activate').length;
  return { ok: true, logs: r.logs, message: edits ? T('Done. {count, plural, one {# change} other {# changes}}.', { count: edits }) : T('Done. No changes.'), changes: edits };
};
