// The script runtime. `buildRuntime` must stay self-contained (no imports, no outer references):
// its source text is copied into the sandboxed worker with Function.prototype.toString().

export type SnapValue = number | string | boolean | null;
/** [row, col, value, formula without "="] */
export type SnapCell = [number, number, SnapValue, string?];
export interface SnapSheet { name: string; cells: SnapCell[]; hiddenRows: number[]; hiddenCols: number[] }
export interface Snapshot { sheets: SnapSheet[]; active: number; selection: { sheet: number; r1: number; c1: number; r2: number; c2: number } }

export interface Area { r1: number; c1: number; r2: number; c2: number }
export interface FormatPatch { bold?: boolean; italic?: boolean; fontColor?: string | null; fill?: string | null; numFmt?: string; hAlign?: string }

export type ScriptOp =
  | { t: 'value'; s: number; r: number; c: number; v: SnapValue }
  | { t: 'formula'; s: number; r: number; c: number; f: string }
  | { t: 'clear'; s: number; a: Area; what: 'contents' | 'formats' | 'all' }
  | { t: 'format'; s: number; a: Area; p: FormatPatch }
  | { t: 'autofit'; s: number; c1: number; c2: number }
  | { t: 'sort'; s: number; a: Area; col: number; asc: boolean }
  | { t: 'freeze'; s: number; rows: number; cols: number }
  | { t: 'addSheet'; name: string }
  | { t: 'renameSheet'; s: number; name: string }
  | { t: 'hide'; s: number; axis: 'row' | 'col'; from: number; to: number; hidden: boolean }
  | { t: 'deleteRows'; s: number; at: number; count: number }
  | { t: 'addTable'; s: number; a: Area; name: string }
  | { t: 'select'; s: number; a: Area }
  | { t: 'activate'; s: number };

export interface RunResult { ok: boolean; ops: ScriptOp[]; logs: string[]; error?: string; line?: number }

export type UserMain = (workbook: unknown, console: unknown) => unknown;

/* eslint-disable @typescript-eslint/no-explicit-any */
export function buildRuntime() {
  const MAX_ROWS = 1048576, MAX_COLS = 16384, MAX_OPS = 2000000, MAX_LOGS = 2000;
  const BLOCKED = ['fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'BroadcastChannel',
    'Worker', 'SharedWorker', 'WebTransport', 'RTCPeerConnection', 'webkitRTCPeerConnection', 'Request', 'navigator', 'open', 'close'];

  /** Replaces network, storage and messaging globals (on the object and its prototypes) with functions that throw. */
  const shadow = (scope: any, sealRealm = false): void => {
    for (const name of BLOCKED) {
      const stub = function blocked() { throw new Error(`${name} isn't available in scripts.`); };
      let o: any = scope;
      for (let depth = 0; o && depth < 8; depth++, o = Object.getPrototypeOf(o)) {
        try { if (Object.prototype.hasOwnProperty.call(o, name) || depth === 0) Object.defineProperty(o, name, { value: stub, writable: false, configurable: false }); } catch { /* not configurable */ }
      }
    }
    if (!sealRealm) return;
    // No building code from text: eval and every function constructor throw.
    const noEval = function blocked() { throw new Error("Building code from text isn't available in scripts."); };
    try { Object.defineProperty(scope, 'eval', { value: noEval, writable: false, configurable: false }); } catch { /* ignore */ }
    const protos = [function () {}, async function () {}, function* () {}, async function* () {}].map((f) => Object.getPrototypeOf(f));
    for (const p of protos) { try { Object.defineProperty(p, 'constructor', { value: noEval, writable: false, configurable: false }); } catch { /* ignore */ } }
    try { Object.defineProperty(scope, 'Function', { value: noEval, writable: false, configurable: false }); } catch { /* ignore */ }
  };

  const colName = (col: number): string => { let s = '', n = col + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colIndex = (name: string): number => { let n = 0; for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const parseArea = (text: string): Area => {
    const t = String(text).trim().replace(/\$/g, '');
    let m = /^([A-Za-z]{1,3})(\d{1,7})(?::([A-Za-z]{1,3})(\d{1,7}))?$/.exec(t);
    if (m) {
      const c1 = colIndex(m[1]!), r1 = +m[2]! - 1, c2 = m[3] ? colIndex(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
      if (r1 >= 0 && r2 >= 0 && c1 < MAX_COLS && c2 < MAX_COLS && r1 < MAX_ROWS && r2 < MAX_ROWS) return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
    }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
    if (m) { const a = colIndex(m[1]!), b = colIndex(m[2]!); if (a < MAX_COLS && b < MAX_COLS) return { r1: 0, c1: Math.min(a, b), r2: MAX_ROWS - 1, c2: Math.max(a, b) }; }
    m = /^(\d{1,7}):(\d{1,7})$/.exec(t);
    if (m) { const a = +m[1]! - 1, b = +m[2]! - 1; if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) return { r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAX_COLS - 1 }; }
    throw new Error(`"${text}" isn't a valid range. Use an address such as "A1:C3".`);
  };
  const addrOf = (a: Area): string => {
    const one = (r: number, c: number) => `${colName(c)}${r + 1}`;
    return a.r1 === a.r2 && a.c1 === a.c2 ? one(a.r1, a.c1) : `${one(a.r1, a.c1)}:${one(a.r2, a.c2)}`;
  };
  const cleanValue = (v: unknown): SnapValue => {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v === 'boolean') return v;
    if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.getTime() / 86400000 + 25569;
    const s = String(v);
    return s.length > 32767 ? s.slice(0, 32767) : s;
  };
  const isColor = (c: unknown): c is string => typeof c === 'string' && /^#?[0-9a-fA-F]{6}$/.test(c);
  const color = (c: unknown): string => {
    if (!isColor(c)) throw new Error(`"${String(c)}" isn't a color. Use "#RRGGBB".`);
    return (c.startsWith('#') ? c : `#${c}`).toUpperCase();
  };
  const fmtArg = (v: unknown): string => {
    if (typeof v === 'string') return v;
    try { return JSON.stringify(v) ?? String(v); } catch { return String(v); }
  };

  /** Runs `main` against an in-memory copy of the workbook and returns the edits to apply. */
  const run = (snap: Snapshot, main: UserMain): RunResult => {
    const ops: ScriptOp[] = [];
    const logs: string[] = [];
    const push = (op: ScriptOp) => { if (ops.length >= MAX_OPS) throw new Error('The script made too many changes.'); ops.push(op); };
    interface MSheet { name: string; cells: Map<number, { v: SnapValue; f?: string }>; hiddenRows: Set<number>; hiddenCols: Set<number>; frozen: [number, number]; tables: number }
    const sheets: MSheet[] = snap.sheets.map((s) => ({
      name: s.name, cells: new Map(s.cells.map(([r, c, v, f]) => [r * MAX_COLS + c, f ? { v, f } : { v }])),
      hiddenRows: new Set(s.hiddenRows), hiddenCols: new Set(s.hiddenCols), frozen: [0, 0], tables: 0,
    }));
    let active = snap.active;
    let selection = snap.selection;
    const key = (r: number, c: number) => r * MAX_COLS + c;
    const bounds = (sh: MSheet): Area | null => {
      let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
      for (const k of sh.cells.keys()) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c; }
      for (const r of sh.hiddenRows) { if (r < r1) r1 = r; if (r > r2) r2 = r; if (c1 === Infinity) { c1 = 0; c2 = 0; } }
      for (const c of sh.hiddenCols) { if (c < c1) c1 = c; if (c > c2) c2 = c; if (r1 === Infinity) { r1 = 0; r2 = 0; } }
      return r2 < 0 ? null : { r1, c1, r2, c2 };
    };
    /** Large whole-row/column ranges are cut to the used area when reading or writing cell by cell. */
    const clip = (sh: MSheet, a: Area): Area => {
      const b = bounds(sh);
      if ((a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) <= 1000000) return a;
      if (!b) return { r1: a.r1, c1: a.c1, r2: a.r1 - 1, c2: a.c1 - 1 };
      return { r1: a.r1, c1: a.c1, r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
    };
    const sheetIndexOf = (sh: MSheet) => sheets.indexOf(sh);

    const makeRange = (sh: MSheet, a: Area): any => {
      const s = () => sheetIndexOf(sh);
      const rows = a.r2 - a.r1 + 1, cols = a.c2 - a.c1 + 1;
      const grid = <T,>(f: (r: number, c: number) => T): T[][] => {
        const b = clip(sh, a);
        if ((b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1) > 1000000) throw new Error('That range is too large to read at once.');
        const out: T[][] = [];
        for (let r = b.r1; r <= b.r2; r++) { const line: T[] = []; for (let c = b.c1; c <= b.c2; c++) line.push(f(r, c)); out.push(line); }
        return out;
      };
      const write = (data: unknown, asFormulas: boolean) => {
        if (!Array.isArray(data) || !data.every(Array.isArray)) throw new Error('Pass a two-dimensional array, for example [[1, 2], [3, 4]].');
        const m = data as unknown[][];
        if (m.length !== rows || m.some((line) => line.length !== cols)) throw new Error(`The array is ${m.length}×${m[0]?.length ?? 0} but the range ${addrOf(a)} is ${rows}×${cols}.`);
        m.forEach((line, i) => line.forEach((v, j) => setOne(a.r1 + i, a.c1 + j, v, asFormulas)));
      };
      /** Text starting with "=" is a formula; with setValues a leading apostrophe keeps it as text. */
      const setOne = (r: number, c: number, raw: unknown, asFormula: boolean) => {
        let v = cleanValue(raw);
        if (!asFormula && typeof v === 'string' && v.startsWith("'")) v = v.slice(1);
        else if (typeof v === 'string' && v.length > 1 && v.startsWith('=')) {
          sh.cells.set(key(r, c), { v: null, f: v.slice(1) });
          push({ t: 'formula', s: s(), r, c, f: v.slice(1) });
          return;
        }
        if (v === null) sh.cells.delete(key(r, c)); else sh.cells.set(key(r, c), { v });
        push({ t: 'value', s: s(), r, c, v });
      };
      const format = (p: FormatPatch) => { push({ t: 'format', s: s(), a: { ...a }, p }); return range; };
      const range: any = {
        getAddress: () => addrOf(a),
        getSheet: () => sheetApi(sh),
        getRowIndex: () => a.r1,
        getColumnIndex: () => a.c1,
        getRowCount: () => rows,
        getColumnCount: () => cols,
        getCell: (r: number, c: number) => {
          if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) throw new Error('getCell needs whole numbers from 0.');
          return makeRange(sh, { r1: a.r1 + r, c1: a.c1 + c, r2: a.r1 + r, c2: a.c1 + c });
        },
        getRow: (i: number) => makeRange(sh, { r1: a.r1 + i, c1: a.c1, r2: a.r1 + i, c2: a.c2 }),
        getColumn: (i: number) => makeRange(sh, { r1: a.r1, c1: a.c1 + i, r2: a.r2, c2: a.c1 + i }),
        getValues: () => grid((r, c) => sh.cells.get(key(r, c))?.v ?? null),
        getFormulas: () => grid((r, c) => { const x = sh.cells.get(key(r, c)); return x?.f ? `=${x.f}` : x?.v ?? null; }),
        getValue: () => sh.cells.get(key(a.r1, a.c1))?.v ?? null,
        getFormula: () => { const x = sh.cells.get(key(a.r1, a.c1)); return x?.f ? `=${x.f}` : x?.v ?? null; },
        setValues: (m: unknown) => { write(m, false); return range; },
        setFormulas: (m: unknown) => { write(m, true); return range; },
        setValue: (v: unknown) => { const b = clip(sh, a); for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) setOne(r, c, v, false); return range; },
        setFormula: (f: unknown) => { const t = String(f); range.setValue(t.startsWith('=') ? t : `=${t}`); return range; },
        clear: (what: unknown = 'all') => {
          const w = what === 'contents' || what === 'formats' ? what : 'all';
          if (w !== 'formats') { const b = clip(sh, a); for (const k of [...sh.cells.keys()]) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2) sh.cells.delete(k); } }
          push({ t: 'clear', s: s(), a: { ...a }, what: w });
          return range;
        },
        setBold: (on: unknown = true) => format({ bold: !!on }),
        setItalic: (on: unknown = true) => format({ italic: !!on }),
        setFontColor: (c: unknown) => format({ fontColor: c === null ? null : color(c) }),
        setFill: (c: unknown) => format({ fill: c === null || c === undefined || c === '' ? null : color(c) }),
        setNumberFormat: (f: unknown) => format({ numFmt: String(f).slice(0, 255) }),
        setHorizontalAlignment: (h: unknown) => {
          const v = String(h).toLowerCase();
          if (!['general', 'left', 'center', 'right'].includes(v)) throw new Error('Alignment must be "left", "center", "right" or "general".');
          return format({ hAlign: v });
        },
        autofitColumns: () => { push({ t: 'autofit', s: s(), c1: a.c1, c2: Math.min(a.c2, a.c1 + 500) }); return range; },
        setRowsHidden: (hidden: unknown = true) => {
          for (let r = a.r1; r <= Math.min(a.r2, a.r1 + 100000); r++) { if (hidden) sh.hiddenRows.add(r); else sh.hiddenRows.delete(r); }
          push({ t: 'hide', s: s(), axis: 'row', from: a.r1, to: a.r2, hidden: !!hidden });
          return range;
        },
        setColumnsHidden: (hidden: unknown = true) => {
          for (let c = a.c1; c <= a.c2; c++) { if (hidden) sh.hiddenCols.add(c); else sh.hiddenCols.delete(c); }
          push({ t: 'hide', s: s(), axis: 'col', from: a.c1, to: a.c2, hidden: !!hidden });
          return range;
        },
        sort: (column: unknown = 0, ascending: unknown = true) => {
          let col: number;
          if (typeof column === 'number' && Number.isInteger(column)) col = a.c1 + column;
          else if (typeof column === 'string' && /^[A-Za-z]{1,3}$/.test(column)) col = colIndex(column);
          else throw new Error('sort needs a column: a number counted from 0 inside the range, or a letter such as "B".');
          if (col < a.c1 || col > a.c2) throw new Error(`Column ${colName(col)} is outside the range ${addrOf(a)}.`);
          const b = clip(sh, a);
          const lines: { v: { v: SnapValue; f?: string } | undefined; cells: ({ v: SnapValue; f?: string } | undefined)[]; i: number }[] = [];
          for (let r = b.r1; r <= b.r2; r++) {
            const cells = []; for (let c = b.c1; c <= b.c2; c++) cells.push(sh.cells.get(key(r, c)));
            lines.push({ v: sh.cells.get(key(r, col)), cells, i: r });
          }
          const rank = (x: SnapValue) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : typeof x === 'boolean' ? 2 : 3);
          const asc = ascending !== false;
          lines.sort((p, q) => {
            const x = p.v?.v ?? null, y = q.v?.v ?? null;
            if (x === null || y === null) return x === y ? p.i - q.i : x === null ? 1 : -1;
            let d = rank(x) - rank(y);
            if (!d) d = typeof x === 'string' ? x.localeCompare(y as string, undefined, { sensitivity: 'base' }) : (x as number) - (y as number);
            return d ? (asc ? d : -d) : p.i - q.i;
          });
          lines.forEach((ln, k) => ln.cells.forEach((cell, j) => {
            const kk = key(b.r1 + k, b.c1 + j);
            if (cell) sh.cells.set(kk, cell); else sh.cells.delete(kk);
          }));
          push({ t: 'sort', s: s(), a: { ...b }, col, asc });
          return range;
        },
        select: () => { selection = { sheet: s(), ...a }; active = s(); push({ t: 'select', s: s(), a: { ...a } }); return range; },
      };
      return range;
    };

    const apiCache = new Map<MSheet, any>();
    const sheetApi = (sh: MSheet): any => {
      const hit = apiCache.get(sh);
      if (hit) return hit;
      const s = () => sheetIndexOf(sh);
      const api = {
        getName: () => sh.name,
        getIndex: () => s(),
        setName: (name: unknown) => {
          const n = String(name).trim();
          if (!n || n.length > 31 || /[\\/?*[\]:]/.test(n)) throw new Error(`"${n}" can't be used as a sheet name.`);
          if (sheets.some((o) => o !== sh && o.name.toLowerCase() === n.toLowerCase())) throw new Error(`Another sheet is already called "${n}".`);
          sh.name = n;
          push({ t: 'renameSheet', s: s(), name: n });
          return api;
        },
        getRange: (address?: unknown) => {
          if (address === undefined) { const b = bounds(sh) ?? { r1: 0, c1: 0, r2: 0, c2: 0 }; return makeRange(sh, b); }
          return makeRange(sh, parseArea(String(address)));
        },
        getCell: (row: number, col: number) => {
          if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) throw new Error('getCell needs whole numbers from 0, for example getCell(0, 0) for A1.');
          return makeRange(sh, { r1: row, c1: col, r2: row, c2: col });
        },
        getUsedRange: () => { const b = bounds(sh); return b ? makeRange(sh, b) : null; },
        freezePanes: (rows: unknown = 0, cols: unknown = 0) => {
          const r = Math.max(0, Math.min(1000, Math.floor(Number(rows) || 0))), c = Math.max(0, Math.min(1000, Math.floor(Number(cols) || 0)));
          sh.frozen = [r, c];
          push({ t: 'freeze', s: s(), rows: r, cols: c });
          return api;
        },
        deleteRows: (index: unknown, count: unknown = 1) => {
          const at = Math.floor(Number(index)), n = Math.floor(Number(count));
          if (!(at >= 0) || !(n >= 1) || at + n > MAX_ROWS) throw new Error('deleteRows needs a row index from 0 and a count of at least 1.');
          const moved = new Map<number, { v: SnapValue; f?: string }>();
          for (const [k, v] of sh.cells) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (r < at) moved.set(k, v); else if (r >= at + n) moved.set(key(r - n, c), v); }
          sh.cells = moved;
          sh.hiddenRows = new Set([...sh.hiddenRows].filter((r) => r < at || r >= at + n).map((r) => (r >= at + n ? r - n : r)));
          push({ t: 'deleteRows', s: s(), at, count: n });
          return api;
        },
        addTable: (address: unknown, name?: unknown) => {
          const a = parseArea(String(address));
          if (a.r2 <= a.r1) throw new Error('A table needs a header row and at least one row of data.');
          sh.tables++;
          const n = name === undefined ? '' : String(name).replace(/[^A-Za-z0-9_.]/g, '_').replace(/^([^A-Za-z_])/, '_$1').slice(0, 255);
          push({ t: 'addTable', s: s(), a, name: n });
          return makeRange(sh, a);
        },
        activate: () => { active = s(); push({ t: 'activate', s: s() }); return api; },
      };
      apiCache.set(sh, api);
      return api;
    };

    const workbook = {
      getActiveSheet: () => sheetApi(sheets[active] ?? sheets[0]!),
      getSheet: (name: unknown) => { const n = String(name).toLowerCase(); const sh = sheets.find((x) => x.name.toLowerCase() === n); return sh ? sheetApi(sh) : null; },
      getSheets: () => sheets.map(sheetApi),
      addSheet: (name?: unknown) => {
        let n = name === undefined ? '' : String(name).trim();
        if (!n) { let i = sheets.length + 1; while (sheets.some((x) => x.name.toLowerCase() === `sheet${i}`)) i++; n = `Sheet${i}`; }
        if (n.length > 31 || /[\\/?*[\]:]/.test(n)) throw new Error(`"${n}" can't be used as a sheet name.`);
        if (sheets.some((x) => x.name.toLowerCase() === n.toLowerCase())) throw new Error(`Another sheet is already called "${n}".`);
        const sh: MSheet = { name: n, cells: new Map(), hiddenRows: new Set(), hiddenCols: new Set(), frozen: [0, 0], tables: 0 };
        sheets.push(sh);
        push({ t: 'addSheet', name: n });
        return sheetApi(sh);
      },
      getSelectedRange: () => {
        const sh = sheets[selection.sheet] ?? sheets[active]!;
        return makeRange(sh, { r1: selection.r1, c1: selection.c1, r2: selection.r2, c2: selection.c2 });
      },
    };
    const log = (...args: unknown[]) => { if (logs.length < MAX_LOGS) logs.push(args.map(fmtArg).join(' ').slice(0, 10000)); };
    const konsole = Object.freeze({ log, info: log, warn: (...a: unknown[]) => log('Warning:', ...a), error: (...a: unknown[]) => log('Error:', ...a) });
    try {
      main(Object.freeze(workbook), konsole);
      return { ok: true, ops, logs };
    } catch (e) {
      const err = e as { message?: string; stack?: string; name?: string };
      return { ok: false, ops: [], logs, error: `${err?.name && err.name !== 'Error' ? `${err.name}: ` : ''}${err?.message ?? String(e)}`, line: lineOf(err?.stack) };
    }
  };

  /** Line number inside the script, from an error stack. Frames are counted from USER_START (set by the worker). */
  let userStart = 0;
  const lineOf = (stack: string | undefined): number | undefined => {
    if (!stack) return undefined;
    for (const ln of stack.split('\n')) {
      const m = /(?:blob:[^\s)]+|<anonymous>|clerdo-script):(\d+):(\d+)\)?\s*$/.exec(ln);
      if (m && +m[1]! > userStart) return +m[1]! - userStart;
    }
    return undefined;
  };

  /** Inside the worker, before the person's code is evaluated: shadows the network and keeps a private postMessage. */
  let post: ((m: unknown) => void) | null = null;
  const prepare = (scope: any, start: number): void => {
    userStart = start;
    post = scope.postMessage.bind(scope);
    shadow(scope, true);
    try { Object.defineProperty(scope, 'postMessage', { value: () => { throw new Error("postMessage isn't available in scripts."); }, configurable: false, writable: false }); } catch { /* ignore */ }
  };
  /** Answers one "run" message with the result. The script's wrapper also gets local stand-ins for the blocked names. */
  const install = (scope: any, user: (...a: unknown[]) => unknown): void => {
    const send = post;
    if (!send) return;
    const stubs = LOCALS.map((name) => function blocked() { throw new Error(`${name} isn't available in scripts.`); });
    scope.onmessage = (e: { data: { snapshot: Snapshot } }) => {
      scope.onmessage = null;
      send(run(e.data.snapshot, (wb, con) => user(wb, con, ...stubs)));
    };
  };
  const LOCALS = [...BLOCKED, 'self', 'globalThis', 'postMessage', 'onmessage'];

  return { run, shadow, prepare, install, parseArea, addrOf, LOCALS };
}

/**
 * The full worker script. The runtime shadows the network first; the person's code is wrapped in a
 * function so it only runs once the data arrives.
 */
export const workerSource = (code: string): { source: string; userStart: number } => {
  const rt = `const __clerdoRt = (${buildRuntime.toString()})();\n`;
  const pre = `"use strict";\n${rt}`;
  const userStart = pre.split('\n').length - 1 + 2;
  const head = `${pre}__clerdoRt.prepare(self, ${userStart});\nfunction __clerdoUser(workbook, console, ${buildRuntime().LOCALS.join(', ')}) {\n`;
  const tail = `\n;if (typeof main === "function") return main(workbook);\nthrow new Error("Add a function called main(workbook).");\n}\n__clerdoRt.install(self, __clerdoUser);\n`;
  return { source: head + code + tail, userStart };
};
