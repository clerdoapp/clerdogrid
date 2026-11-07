import type { GridStore } from '../../app/store';
import { applyOps, snapshot } from './apply';
import { workerSource, type RunResult } from './runtime';

export const TIME_LIMIT_MS = 10_000;
const MAX_CODE = 1_000_000;

/** Turns a line in the worker file into a line in the person's script, or undefined when it's outside it. */
const userLine = (absolute: number | undefined, start: number, codeLines: number): number | undefined => {
  if (!absolute) return undefined;
  const l = absolute > start ? absolute - start : absolute;
  return l >= 1 && l <= codeLines ? l : undefined;
};

/**
 * Runs a script in a throwaway Web Worker made from a Blob: no DOM, network globals replaced with
 * functions that throw, and stopped after the time limit. The worker only ever sees a copy of the cells.
 */
export const runInWorker = (store: GridStore, code: string, limitMs = TIME_LIMIT_MS): Promise<RunResult> => {
  if (code.length > MAX_CODE) return Promise.resolve({ ok: false, ops: [], logs: [], error: 'The script is too long.' });
  if (typeof Worker === 'undefined') return Promise.resolve({ ok: false, ops: [], logs: [], error: 'Scripts need a browser that supports workers.' });
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
    const timer = setTimeout(() => finish({ ok: false, ops: [], logs: [], error: `The script ran for more than ${Math.round(limitMs / 1000)} seconds and was stopped.` }), limitMs);
    try {
      worker = new Worker(url, { name: 'clerdogrid-script' });
    } catch (e) {
      finish({ ok: false, ops: [], logs: [], error: `The script could not start: ${e instanceof Error ? e.message : String(e)}` });
      return;
    }
    worker.onmessage = (e: MessageEvent<RunResult>) => {
      const r = e.data;
      if (!r || typeof r !== 'object' || !Array.isArray(r.ops) || !Array.isArray(r.logs)) { finish({ ok: false, ops: [], logs: [], error: 'The script returned something unexpected.' }); return; }
      finish({
        ok: !!r.ok, ops: r.ops, logs: r.logs.filter((x) => typeof x === 'string'),
        error: typeof r.error === 'string' ? r.error : undefined,
        line: typeof r.line === 'number' ? userLine(r.line, 0, lines) : undefined,
      });
    };
    worker.onerror = (e: ErrorEvent) => {
      e.preventDefault();
      const msg = (e.message || 'The script has an error.').replace(/^Uncaught\s+/, '');
      finish({ ok: false, ops: [], logs: [], error: msg, line: userLine(e.lineno, userStart, lines) });
    };
    worker.postMessage({ snapshot: snapshot(store) });
  });
};

export interface ScriptOutcome { ok: boolean; logs: string[]; message: string; line?: number; changes: number }

/** Runs a script and applies its edits as one undo step. Never called when a file opens. */
export const runScript = async (store: GridStore, code: string, name = 'script'): Promise<ScriptOutcome> => {
  const r = await runInWorker(store, code);
  if (!r.ok) {
    const where = r.line ? ` (line ${r.line})` : '';
    return { ok: false, logs: r.logs, message: `${r.error ?? 'The script failed.'}${where}`, line: r.line, changes: 0 };
  }
  const err = applyOps(store, r.ops, `Run ${name}`);
  if (err) return { ok: false, logs: r.logs, message: err, changes: 0 };
  const edits = r.ops.filter((o) => o.t !== 'select' && o.t !== 'activate').length;
  return { ok: true, logs: r.logs, message: edits ? `Done. ${edits} ${edits === 1 ? 'change' : 'changes'}.` : 'Done. No changes.', changes: edits };
};
