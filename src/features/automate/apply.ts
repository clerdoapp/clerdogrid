import { FormulaError, parseFormula } from '../../engine/parser';
import { CellError, MAX_COLS, MAX_ROWS, type Scalar } from '../../engine/types';
import { changeStructure, renameSheet } from '../../app/ops/structure';
import { sortRange } from '../../app/ops/data';
import { displayText } from '../../app/ops/clipboard';
import { canonical, type GridStore } from '../../app/store';
import { pxToChars } from '../../app/layout';
import { keyCol, keyRow, newSheet, type CellRange, type CellStyle, type TableDef } from '../../model/workbook';
import { styleTable, TABLE_LOOKS } from '../../ui/dialogs/styles';
import { buildRuntime, type FormatPatch, type RunResult, type ScriptOp, type SnapCell, type SnapValue, type Snapshot } from './runtime';

const snapValue = (v: Scalar): SnapValue => (v instanceof CellError ? v.code : v);

/** Copies every sheet's cells (values and formulas) so a script can read them without touching the store. */
export const snapshot = (store: GridStore): Snapshot => {
  const r = store.range;
  return {
    active: store.sheetIndex,
    selection: { sheet: store.sheetIndex, ...r },
    sheets: store.wb.sheets.map((s, si) => {
      const cells: SnapCell[] = [];
      for (const [k, c] of s.cells) {
        const row = keyRow(k), col = keyCol(k);
        const v = snapValue(store.engine.getValue(si, row, col));
        if (c.f) cells.push([row, col, v, c.f]);
        else if (v !== null) cells.push([row, col, v]);
      }
      const hiddenRows = [...s.rows].filter(([, i]) => i.hidden).map(([n]) => n);
      const hiddenCols = [...s.cols].filter(([, i]) => i.hidden).map(([n]) => n);
      return { name: s.name, cells, hiddenRows, hiddenCols };
    }),
  };
};

/** Runs a script without a worker. Only for trusted callers such as tests: `main` is a real function. */
export const runDirect = (store: GridStore, main: (workbook: never) => unknown): RunResult =>
  buildRuntime().run(snapshot(store), (wb) => main(wb as never));

const HEX = /^#[0-9A-F]{6}$/i;
const ALIGNS = new Set(['general', 'left', 'center', 'right']);
const patchStyle = (raw: FormatPatch) => (s: CellStyle): CellStyle => {
  const p: FormatPatch = {
    bold: typeof raw.bold === 'boolean' ? raw.bold : undefined,
    italic: typeof raw.italic === 'boolean' ? raw.italic : undefined,
    fontColor: raw.fontColor === null || (typeof raw.fontColor === 'string' && HEX.test(raw.fontColor)) ? raw.fontColor : undefined,
    fill: raw.fill === null || (typeof raw.fill === 'string' && HEX.test(raw.fill)) ? raw.fill : undefined,
    numFmt: typeof raw.numFmt === 'string' ? raw.numFmt.slice(0, 255) : undefined,
    hAlign: typeof raw.hAlign === 'string' && ALIGNS.has(raw.hAlign) ? raw.hAlign : undefined,
  };
  const out: CellStyle = { ...s };
  if (p.bold !== undefined || p.italic !== undefined || p.fontColor !== undefined) {
    const font = { ...s.font };
    if (p.bold !== undefined) font.bold = p.bold || undefined;
    if (p.italic !== undefined) font.italic = p.italic || undefined;
    if (p.fontColor !== undefined) font.color = p.fontColor ?? undefined;
    out.font = font;
  }
  if (p.fill !== undefined) out.fill = p.fill ? { color: p.fill } : undefined;
  if (p.numFmt !== undefined) out.numFmt = p.numFmt;
  if (p.hAlign !== undefined) out.align = { ...s.align, h: p.hAlign as NonNullable<CellStyle['align']>['h'] };
  return out;
};

const measure = (() => {
  let ctx: CanvasRenderingContext2D | null | undefined;
  return (text: string, st: CellStyle): number => {
    const size = ((st.font?.size ?? 11) * 4) / 3;
    if (ctx === undefined) ctx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
    if (!ctx) return text.length * size * 0.55 * (st.font?.bold ? 1.08 : 1);
    ctx.font = `${st.font?.bold ? '700 ' : ''}${size}px Carlito, Calibri, sans-serif`;
    return ctx.measureText(text).width;
  };
})();

const autofit = (store: GridStore, sheet: number, c1: number, c2: number): void => {
  const s = store.wb.sheets[sheet]!;
  const widths = new Map<number, number>();
  for (const k of s.cells.keys()) {
    const c = keyCol(k), row = keyRow(k);
    if (c < c1 || c > c2 || store.mergeAt(row, c, sheet)) continue;
    if (store.engine.getValue(sheet, row, c) === null) continue;
    widths.set(c, Math.max(widths.get(c) ?? 0, measure(displayText(store, row, c, sheet), store.style(row, c, sheet))));
  }
  for (const [c, w] of widths) { const info = s.cols.get(c) ?? {}; info.width = Math.min(255, pxToChars(Math.ceil(w) + 9)); s.cols.set(c, info); }
};

const uniqueTableName = (store: GridStore, wanted: string): string => {
  const taken = new Set([...store.wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())), ...store.wb.names.map((n) => n.name.toLowerCase())]);
  if (wanted && !taken.has(wanted.toLowerCase()) && !/^[A-Za-z]{1,3}\d+$/.test(wanted)) return wanted;
  for (let i = 1; ; i++) if (!taken.has(`table${i}`)) return `Table${i}`;
};

const addTable = (store: GridStore, sheet: number, a: CellRange, wanted: string): void => {
  const s = store.wb.sheets[sheet]!;
  if (s.tables.some((t) => t.range.r1 <= a.r2 && t.range.r2 >= a.r1 && t.range.c1 <= a.c2 && t.range.c2 >= a.c1)) throw new Error('That range overlaps another table.');
  const names: string[] = [];
  for (let c = a.c1; c <= a.c2; c++) {
    const v = store.engine.getValue(sheet, a.r1, c);
    let n = v === null ? '' : String(snapValue(v)).trim();
    if (!n) n = `Column${c - a.c1 + 1}`;
    const stem = n;
    for (let k = 2; names.some((x) => x.toLowerCase() === n.toLowerCase()); k++) n = `${stem}${k}`;
    names.push(n);
    if (v === null || String(snapValue(v)) !== n) store.engine.setValue(sheet, a.r1, c, n);
  }
  const t: TableDef = { name: uniqueTableName(store, wanted), range: { ...a }, headerRow: true, totalsRow: false, style: 'TableStyleMedium2', bandedRows: true, bandedCols: false, columns: names.map((name) => ({ name })) };
  s.tables.push(t);
  store.engine.sheetsChanged();
  const lk = TABLE_LOOKS.find((l) => l.id === t.style) ?? TABLE_LOOKS[0]!;
  const prev = store.wb.activeSheet;
  store.wb.activeSheet = sheet;
  try { styleTable(store, t, lk); } finally { store.wb.activeSheet = prev; }
};

/** Runs `fn` with `sheet` as the active sheet, for operations that work on the active sheet. */
const onSheet = (store: GridStore, sheet: number, fn: () => void): void => {
  const prev = store.wb.activeSheet;
  store.wb.activeSheet = sheet;
  try { fn(); } finally { store.wb.activeSheet = Math.min(prev, store.wb.sheets.length - 1); }
};

const isInt = (n: unknown, max: number): boolean => typeof n === 'number' && Number.isInteger(n) && n >= 0 && n < max;
const isArea = (a: unknown): boolean => {
  const x = a as CellRange | null;
  return !!x && isInt(x.r1, MAX_ROWS) && isInt(x.r2, MAX_ROWS) && isInt(x.c1, MAX_COLS) && isInt(x.c2, MAX_COLS) && x.r1 <= x.r2 && x.c1 <= x.c2;
};
const OP_TYPES = new Set(['value', 'formula', 'clear', 'format', 'autofit', 'sort', 'freeze', 'addSheet', 'renameSheet', 'hide', 'deleteRows', 'addTable', 'select', 'activate']);

/** The edits come back from the sandbox, so their shape is checked like any other untrusted input. */
const wellFormed = (op: ScriptOp): boolean => {
  if (!op || typeof op !== 'object' || !OP_TYPES.has(op.t)) return false;
  const o = op as unknown as Record<string, unknown>;
  if ('r' in o && !isInt(o.r, MAX_ROWS)) return false;
  if ('c' in o && !isInt(o.c, MAX_COLS)) return false;
  if ('a' in o && !isArea(o.a)) return false;
  if (op.t === 'value' && !(op.v === null || ['number', 'string', 'boolean'].includes(typeof op.v))) return false;
  if ((op.t === 'formula' && typeof op.f !== 'string') || ((op.t === 'addSheet' || op.t === 'renameSheet') && typeof op.name !== 'string')) return false;
  if (op.t === 'autofit' && !(isInt(op.c1, MAX_COLS) && isInt(op.c2, MAX_COLS))) return false;
  if (op.t === 'hide' && !(isInt(op.from, MAX_COLS * 64) && isInt(op.to, MAX_ROWS))) return false;
  if (op.t === 'deleteRows' && !(isInt(op.at, MAX_ROWS) && isInt(op.count, MAX_ROWS + 1))) return false;
  if (op.t === 'freeze' && !(isInt(op.rows, 1001) && isInt(op.cols, 1001))) return false;
  if (op.t === 'format' && (typeof op.p !== 'object' || op.p === null)) return false;
  return true;
};

/** Checks formulas before anything changes, so a bad one stops the run with nothing applied. */
export const checkOps = (store: GridStore, ops: ScriptOp[]): string | null => {
  let sheets = store.wb.sheets.length;
  for (const op of ops) {
    if (!wellFormed(op)) return 'The script returned an edit that ClerdoGrid does not understand.';
    if (op.t === 'addSheet') sheets++;
    if ('s' in op && (op.s < 0 || op.s >= sheets)) return 'The script referred to a sheet that does not exist.';
    if (op.t === 'formula') {
      try { parseFormula(op.f); } catch (e) {
        if (e instanceof FormulaError) return `The formula "=${op.f}" can't be read: ${e.message}`;
        throw e;
      }
    }
  }
  return null;
};

/** Applies a script's edits to the workbook as one undo step. Returns an error message, or null. */
export const applyOps = (store: GridStore, ops: ScriptOp[], label = 'Run script'): string | null => {
  if (!ops.length) return null;
  const bad = checkOps(store, ops);
  if (bad) return bad;
  if (store.wb.protection?.structure && ops.some((o) => o.t === 'addSheet' || o.t === 'renameSheet')) return 'The workbook structure is protected, so the script cannot add or rename sheets.';
  const touched = new Set<number>();
  for (const op of ops) if ('s' in op) touched.add(op.s);
  for (const si of touched) if (store.wb.sheets[si]?.protection?.sheet) return `The sheet "${store.wb.sheets[si]!.name}" is protected. Unprotect it to run this script.`;
  let finalSheet: number | null = null;
  let finalSel: CellRange | null = null;
  let failure: string | null = null;
  store.commitEdit();
  store.change(label, 'book', () => {
    for (const op of ops) {
      try {
        switch (op.t) {
          case 'value': store.engine.setValue(op.s, op.r, op.c, op.v); break;
          case 'formula': store.engine.setFormula(op.s, op.r, op.c, canonical(op.f)); break;
          case 'clear': {
            const s = store.wb.sheets[op.s]!;
            const inA = (k: number) => { const r = keyRow(k), c = keyCol(k); return r >= op.a.r1 && r <= op.a.r2 && c >= op.a.c1 && c <= op.a.c2; };
            const keys = [...s.cells.keys()].filter(inA);
            for (const k of keys) {
              const r = keyRow(k), c = keyCol(k);
              if (op.what === 'all') store.engine.clearCell(op.s, r, c);
              else if (op.what === 'contents') store.engine.setValue(op.s, r, c, null);
              else { const cell = s.cells.get(k); if (cell) { delete cell.s; if (cell.v === null && !cell.f) s.cells.delete(k); store.engine.markChanged(op.s, r, c); } }
            }
            break;
          }
          case 'format': store.styleRange(op.a, patchStyle(op.p), op.s); break;
          case 'autofit': autofit(store, op.s, op.c1, op.c2); break;
          case 'sort': onSheet(store, op.s, () => sortRange(store, op.a, [{ col: op.col, descending: !op.asc }], false)); break;
          case 'freeze': { const v = store.wb.sheets[op.s]!.view; v.frozenRows = op.rows; v.frozenCols = op.cols; break; }
          case 'addSheet': store.wb.sheets.push(newSheet(op.name)); store.engine.sheetsChanged(); break;
          case 'renameSheet': if (!renameSheet(store, op.s, op.name)) throw new Error(`The sheet couldn't be renamed to "${op.name}".`); break;
          case 'hide': {
            const s = store.wb.sheets[op.s]!;
            const map = op.axis === 'row' ? s.rows : s.cols;
            if (op.hidden) for (let i = op.from; i <= Math.min(op.to, op.from + 100000); i++) map.set(i, { ...map.get(i), hidden: true });
            else for (const [i, info] of [...map]) if (i >= op.from && i <= op.to && info.hidden) { delete info.hidden; delete (info as { filtered?: boolean }).filtered; }
            break;
          }
          case 'deleteRows': onSheet(store, op.s, () => changeStructure(store, 'row', op.at, -op.count)); break;
          case 'addTable': addTable(store, op.s, op.a, op.name); break;
          case 'select': finalSheet = op.s; finalSel = op.a; break;
          case 'activate': finalSheet = op.s; break;
        }
      } catch (e) {
        failure = e instanceof Error ? e.message : String(e);
        break;
      }
    }
  });
  if (failure) { store.undo(); return failure; }
  if (finalSheet !== null && finalSheet !== store.sheetIndex) store.setActiveSheet(finalSheet);
  if (finalSel) store.selectRange(finalSel);
  return null;
};

