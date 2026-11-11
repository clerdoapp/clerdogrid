// Applying operations from other people to the open workbook, and to the snapshots in the undo history.
import type { GridStore, Snapshot } from '../../app/store';
import { inRange } from '../../app/store';
import { changeStructure } from '../../app/ops/structure';
import type { Engine } from '../../engine/engine';
import { cellKey, keyCol, keyRow, newSheet, type Cell, type Sheet, type Workbook } from '../../model/workbook';
import { deleteSheetInFormula, renameSheetInFormula } from '../../engine/refactor';
import { sheetOrder } from './sheets';
import type { Op } from './ops';
import { shiftIndex, shiftRange, shiftSelection, xform, type Struct } from './ot';
import { adjustForStructure, type Structural } from '../../engine/refactor';
import { decodeCell, decodeDim, decodeSheet, fromWire, setSheetProp, type WCell } from './wire';

/**
 * Changes the sheet list and puts it back in key order, keeping defined names on their sheets. The result depends
 * only on the sheets present, so every copy ends up with the same list: an empty book gets one blank sheet and a
 * book with only hidden sheets shows its first one.
 */
export const BLANK_SHEET = 's-blank';

/**
 * Deleted sheets are kept (per workbook copy, never saved) and still take other people's edits, so every copy has
 * the same set. If everyone deletes every sheet at once, the deleted sheet last in order is shown again: the shown
 * list is a function of the sheets and the deletions alone, whatever order they arrived in.
 */
const tombs = new WeakMap<Workbook, Map<string, Sheet>>();
export const deletedSheets = (wb: Workbook): Map<string, Sheet> => {
  let m = tombs.get(wb);
  if (!m) { m = new Map(); tombs.set(wb, m); }
  return m;
};

/** Applies the shown-sheets rule after a local change (a new sheet hides a deleted one shown as a stand-in). */
export const settleSheets = (wb: Workbook): void => {
  const active = wb.sheets[wb.activeSheet];
  relist(wb, () => undefined);
  wb.activeSheet = Math.max(0, active ? wb.sheets.indexOf(active) : 0);
};

const relist = (wb: Workbook, fn: () => void): void => {
  const old = [...wb.sheets];
  fn();
  const dead = deletedSheets(wb);
  const real = wb.sheets.filter((s) => !(s.cid && dead.has(s.cid)) && s.cid !== BLANK_SHEET);
  if (real.length) wb.sheets.splice(0, wb.sheets.length, ...real);
  else {
    const last = [...dead.values()].sort(sheetOrder).at(-1);
    if (last) wb.sheets.splice(0, wb.sheets.length, last);
    else if (!wb.sheets.length) { const b = newSheet('Sheet1'); b.cid = BLANK_SHEET; b.cpos = 'i'; wb.sheets.push(b); }
  }
  wb.sheets.sort(sheetOrder);
  if (!wb.sheets.some((s) => s.state === 'visible')) wb.sheets[0]!.state = 'visible';
  wb.names = wb.names.flatMap((n) => {
    if (n.sheet === undefined) return [n];
    const i = old[n.sheet] ? wb.sheets.indexOf(old[n.sheet]!) : -1;
    return i < 0 ? [] : [{ ...n, sheet: i }];
  });
};

const eachFormula = (wb: Workbook, fn: (f: string) => string): void => {
  const safe = (f: string) => { try { return fn(f); } catch { return f; } };
  for (const s of wb.sheets) {
    for (const c of s.cells.values()) if (c.f) c.f = safe(c.f);
    for (const ch of s.charts) for (const ser of ch.series) {
      ser.values = safe(ser.values);
      if (ser.categories) ser.categories = safe(ser.categories);
      if (ser.name) ser.name = safe(ser.name);
    }
  }
  for (const n of wb.names) n.formula = safe(n.formula);
};

/** Sets one cell, telling the engine when there is one. */
const putCell = (wb: Workbook, si: number, row: number, col: number, w: WCell | null, engine?: Engine): void => {
  const sheet = wb.sheets[si]!;
  const k = cellKey(row, col);
  if (!engine) {
    if (w) sheet.cells.set(k, decodeCell(w, wb)); else sheet.cells.delete(k);
    return;
  }
  if (!w) { if (sheet.cells.has(k)) engine.clearCell(si, row, col); return; }
  const c = decodeCell(w, wb);
  if (c.f) {
    try { engine.setFormula(si, row, col, c.f); } catch {
      // A formula this copy can't read is kept as text rather than lost.
      engine.setValue(si, row, col, `=${c.f}`);
    }
  } else engine.setValue(si, row, col, c.v);
  const placed: Cell = sheet.cells.get(k) ?? { v: null };
  if (c.dyn) placed.dyn = true; else delete placed.dyn;
  if (c.arrayRange) placed.arrayRange = c.arrayRange; else delete placed.arrayRange;
  if (c.rich) placed.rich = c.rich; else delete placed.rich;
  if (c.s) placed.s = c.s; else delete placed.s;
  if (placed.v === null && !placed.f && !placed.s && !placed.rich) sheet.cells.delete(k); else sheet.cells.set(k, placed);
};

export interface ApplyCtx { engine?: Engine; store?: GridStore }

/** Applies one op to a workbook. Returns false when it couldn't apply (unknown sheet, or a struct op without a store). */
export const applyOp = (wb: Workbook, op: Op, ctx: ApplyCtx = {}): boolean => {
  const { engine, store } = ctx;
  const si = 's' in op ? wb.sheets.findIndex((s) => s.cid === op.s) : -1;
  // Edits to a deleted sheet still land on its kept copy.
  const dead = si < 0 && 's' in op && op.t !== 'sheetAdd' && op.t !== 'sheetDel' ? deletedSheets(wb).get(op.s) : undefined;
  if (dead) {
    const shell: Workbook = { ...wb, sheets: [dead], names: [], activeSheet: 0 };
    if (op.t === 'struct') structureOnBook(shell, 0, op);
    else if (op.t === 'sheetPos') { const active = wb.sheets[wb.activeSheet]; relist(wb, () => { dead.cpos = op.pos; }); wb.activeSheet = Math.max(0, active ? wb.sheets.indexOf(active) : 0); engine?.sheetsChanged(); }
    else applyOp(shell, op);
    return false;
  }
  const sheet = si >= 0 ? wb.sheets[si]! : undefined;
  const activeRef = wb.sheets[wb.activeSheet];
  const keepActive = () => {
    const i = activeRef ? wb.sheets.indexOf(activeRef) : -1;
    if (i >= 0) { wb.activeSheet = i; return; }
    let a = Math.min(Math.max(0, wb.activeSheet), wb.sheets.length - 1);
    while (a > 0 && wb.sheets[a]?.state !== 'visible') a--;
    wb.activeSheet = a;
    if (store) { const s = wb.sheets[a]!.view.selection; const p = s ? { row: s.r1, col: s.c1 } : { row: 0, col: 0 }; store.sel = { ranges: [s ?? { r1: 0, c1: 0, r2: 0, c2: 0 }], active: p, anchor: p }; store.editing = null; }
  };
  switch (op.t) {
    case 'cells':
      if (!sheet) return false;
      for (const [r, c, w] of op.c) putCell(wb, si, r, c, w, engine);
      return true;
    case 'keyed': {
      if (!sheet) return false;
      const o = sheet as unknown as Record<string, Map<number, unknown> | undefined>;
      const map = o[op.k] ?? new Map<number, unknown>();
      for (const [r, c, v] of op.c) { const val = fromWire(v); if (val === null || val === undefined) map.delete(cellKey(r, c)); else map.set(cellKey(r, c), val); }
      if (op.k === 'comments' && !map.size) delete o.comments; else o[op.k] = map;
      return true;
    }
    case 'dim': {
      if (!sheet) return false;
      const map = (op.a === 'row' ? sheet.rows : sheet.cols) as Map<number, unknown>;
      for (const [i, w] of op.c) if (w) map.set(i, decodeDim(w, wb)); else map.delete(i);
      return true;
    }
    case 'item': {
      if (!sheet) return false;
      const o = sheet as unknown as Record<string, { id: string }[] | undefined>;
      const list = o[op.k] ?? [];
      const at = list.findIndex((x) => x.id === op.id);
      if (op.v === null) { if (at >= 0) list.splice(at, 1); }
      else {
        const v = fromWire(op.v) as { id: string };
        if (at >= 0) list[at] = v; else list.splice(Math.max(0, Math.min(op.i, list.length)), 0, v);
      }
      if (!list.length && (op.k === 'shapes' || op.k === 'ink')) delete o[op.k]; else o[op.k] = list;
      return true;
    }
    case 'prop':
      if (!sheet) return false;
      setSheetProp(sheet, op.k, fromWire(op.v));
      if (engine && (op.k === 'tables' || op.k === 'state')) engine.sheetsChanged();
      return true;
    case 'struct': {
      if (!sheet || !store || store.wb !== wb) return false;
      const prevActive = wb.activeSheet, prevSel = store.sel, prevEdit = store.editing;
      wb.activeSheet = si;
      store.sel = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } };
      try { changeStructure(store, op.a, op.at, op.n); } finally {
        wb.activeSheet = prevActive;
        store.sel = prevActive === si ? shiftSelection(prevSel, op.a, op.at, op.n) : prevSel;
        store.editing = prevEdit;
        if (prevEdit && prevActive === si) {
          const i = op.a === 'row' ? prevEdit.row : prevEdit.col;
          const ni = shiftIndex(i, op.at, op.n);
          store.editing = ni === null ? null : { ...prevEdit, [op.a]: ni };
        }
      }
      return true;
    }
    case 'sheetAdd': {
      // A sheet brought back (undo of a delete) comes back as the kept copy, with every edit made to it meanwhile.
      const kept = deletedSheets(wb).get(op.s) ?? (si >= 0 ? wb.sheets[si] : undefined);
      const s = kept ?? decodeSheet(op.d, wb);
      s.cid = op.s; s.cpos = op.pos; s.name = op.name;
      relist(wb, () => {
        deletedSheets(wb).delete(op.s);
        if (!wb.sheets.includes(s)) wb.sheets.push(s);
      });
      keepActive();
      engine?.sheetsChanged();
      return true;
    }
    case 'sheetDel': {
      if (!sheet || deletedSheets(wb).has(op.s)) return false;
      relist(wb, () => { if (op.s === BLANK_SHEET) wb.sheets.splice(si, 1); else deletedSheets(wb).set(op.s, sheet); });
      eachFormula(wb, (f) => deleteSheetInFormula(f, op.name));
      keepActive();
      engine?.sheetsChanged();
      return true;
    }
    case 'sheetRen':
      if (!sheet) return false;
      sheet.name = op.to;
      eachFormula(wb, (f) => renameSheetInFormula(f, op.from, op.to));
      engine?.sheetsChanged();
      return true;
    case 'sheetPos':
      if (!sheet) return false;
      relist(wb, () => { sheet.cpos = op.pos; });
      keepActive();
      engine?.sheetsChanged();
      return true;
    case 'book': {
      const o = wb as unknown as Record<string, unknown>;
      const v = fromWire(op.v);
      if (op.k === 'names') { wb.names = Array.isArray(v) ? v as Workbook['names'] : []; engine?.sheetsChanged(); return true; }
      if (op.k === 'date1904') { wb.date1904 = v === true; return true; }
      if (v === null || v === undefined) delete o[op.k]; else o[op.k] = v;
      if (op.k === 'calc' && store) {
        const c = wb.calc;
        store.engine.iterate = !!c?.iterate;
        if (c?.maxIterations) store.engine.maxIterations = c.maxIterations;
        if (c?.maxChange) store.engine.maxChange = c.maxChange;
      }
      return true;
    }
  }
};

// ---------- Undo history ----------

/** Runs the real row/column insert/delete on a workbook copy that has no engine or window (an undo snapshot). */
export const structureOnBook = (book: Workbook, si: number, op: Struct): void => {
  const sel = { ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 }, anchor: { row: 0, col: 0 } };
  const shell = {
    wb: book, sel, sheetIndex: si,
    get sheet() { return book.sheets[si]!; },
    get range() { return sel.ranges[0]!; },
    notify: () => undefined, markOp: () => undefined,
    change: (_label: string, _scope: unknown, fn: () => void) => fn(),
    engine: { sheetsChanged: () => undefined, recalc: () => undefined },
  };
  changeStructure(shell as unknown as GridStore, op.a, op.at, op.n);
};

const adjust = (f: string, host: string, st: Structural): string => { try { return adjustForStructure(f, host, st); } catch { return f; } };

/** Formulas on another sheet that point into the sheet whose rows or columns moved. */
const adjustSheetFormulas = (sheet: Sheet, st: Structural): void => {
  for (const c of sheet.cells.values()) if (c.f) c.f = adjust(c.f, sheet.name, st);
  for (const cf of sheet.conditionalFormats) for (const r of cf.rules) {
    if ('formula' in r) r.formula = adjust(r.formula, sheet.name, st);
    if ('formulas' in r) r.formulas = r.formulas.map((f) => adjust(f, sheet.name, st));
  }
  for (const dv of sheet.validations) {
    if (dv.formula1) dv.formula1 = adjust(dv.formula1.replace(/^=/, ''), sheet.name, st);
    if (dv.formula2) dv.formula2 = adjust(dv.formula2.replace(/^=/, ''), sheet.name, st);
  }
  for (const ch of sheet.charts) for (const ser of ch.series) {
    ser.values = adjust(ser.values, sheet.name, st);
    if (ser.categories) ser.categories = adjust(ser.categories, sheet.name, st);
    if (ser.name) ser.name = adjust(ser.name, sheet.name, st);
  }
};

interface StructMeta { kind: 'struct'; sheet: number; name?: string; cid?: string; axis: 'row' | 'col'; at: number; count: number }
const structMeta = (meta: readonly unknown[]): StructMeta | undefined =>
  meta.find((x) => !!x && typeof x === 'object' && (x as { kind?: string }).kind === 'struct') as StructMeta | undefined;

/**
 * Patches the local undo/redo snapshots after someone else's op, so undoing your own step restores your "before"
 * without reverting their change. `namesBefore` are the sheet names before the op was applied.
 *
 * Snapshots of older steps use the positions from before your later row/column inserts and deletes, so the op is
 * moved back through each of those (and forward through the ones on the redo stack) before it is patched in.
 */
export const rebaseHistory = (store: GridStore, op: Op, idsBefore: (string | undefined)[]): void => {
  const wb = store.wb;
  const idsAfter = wb.sheets.map((s) => s.cid);
  const listChanged = op.t === 'sheetAdd' || op.t === 'sheetDel' || op.t === 'sheetPos';
  // Cell and sheet snapshots name their sheet by position in the current list; keep them on the same sheet.
  const remap = (i: number): number | null => {
    if (!listChanged) return i;
    const id = idsBefore[i];
    if (id === undefined) return null;
    const j = idsAfter.indexOf(id);
    return j < 0 ? null : j;
  };
  const moving = { undo: [op] as Op[], redo: [op] as Op[] };

  const patch = (snap: Snapshot, o: Op): boolean => {
    if (snap.kind === 'book') {
      if (o.t === 'struct') {
        const si = snap.data.sheets.findIndex((x) => x.cid === o.s);
        if (si >= 0) structureOnBook(snap.data, si, o);
        else for (const sh of snap.data.sheets) adjustSheetFormulas(sh, { sheet: o.nm, axis: o.a, at: o.at, count: o.n });
        return true;
      }
      applyOp(snap.data, o);
      return true;
    }
    const id = idsAfter[snap.sheet];
    const hostName = wb.sheets[snap.sheet]?.name ?? '';
    if (o.t === 'struct') {
      const st: Structural = { sheet: o.nm, axis: o.a, at: o.at, count: o.n };
      if (snap.kind === 'sheet') {
        if (id === o.s) structureOnBook({ ...wb, sheets: [snap.data], names: [] }, 0, o);
        else adjustSheetFormulas(snap.data, st);
        return true;
      }
      // Unqualified references are on the snapshot's own sheet.
      for (const [, c] of snap.cells) if (c.f) c.f = adjust(c.f, id === o.s ? o.nm : hostName === o.nm ? '\u0000' : hostName, st);
      if (id !== o.s) return true;
      const area = shiftRange(snap.area, o.a, o.at, o.n);
      if (!area) { snap.cells = []; return true; }
      snap.area = area;
      snap.cells = snap.cells.flatMap(([k, c]) => {
        const r = o.a === 'row' ? shiftIndex(keyRow(k), o.at, o.n) : keyRow(k);
        const col = o.a === 'col' ? shiftIndex(keyCol(k), o.at, o.n) : keyCol(k);
        return r === null || col === null ? [] : [[cellKey(r, col), c] as [number, Cell]];
      });
      return true;
    }
    const opSheet = 's' in o && o.t !== 'sheetAdd' ? o.s : null;
    if (o.t === 'sheetRen' || o.t === 'sheetDel') {
      // Formulas in snapshots that name the renamed or deleted sheet.
      const fix = o.t === 'sheetRen' ? (f: string) => renameSheetInFormula(f, o.from, o.to) : (f: string) => deleteSheetInFormula(f, o.name);
      const cells = snap.kind === 'sheet' ? [...snap.data.cells.values()] : snap.cells.map(([, c]) => c);
      for (const c of cells) if (c.f) { try { c.f = fix(c.f); } catch { /* keep */ } }
      if (o.t === 'sheetRen' && snap.kind === 'sheet' && id === o.s) snap.data.name = o.to;
      return true;
    }
    if (opSheet === null || id !== opSheet) return true;
    if (snap.kind === 'sheet') {
      // Live styles: sheet snapshots restore into the same style list.
      const shell: Workbook = { ...wb, sheets: [snap.data] };
      applyOp(shell, { ...o, s: snap.data.cid ?? opSheet } as Op);
      return true;
    }
    if (o.t !== 'cells') return true;
    const hit = new Map<number, WCell | null>();
    for (const [r, c, w] of o.c) if (inRange(snap.area, r, c)) hit.set(cellKey(r, c), w);
    if (!hit.size) return true;
    snap.cells = snap.cells.filter(([x]) => !hit.has(x));
    for (const [k, w] of hit) if (w) snap.cells.push([k, decodeCell(w, wb)]);
    return true;
  };

  store.rebaseHistory((step, stack) => {
    let ops = moving[stack];
    const m = structMeta(step.meta);
    if (m) {
      // This step inserted or deleted lines. Its snapshot (the whole book) shows the sheet at the step's index.
      const book = step.snaps.find((x) => x.kind === 'book');
      if (book?.kind !== 'book') return false;
      const sh = (m.cid !== undefined ? book.data.sheets.find((x) => x.cid === m.cid) : undefined) ?? book.data.sheets[m.sheet];
      if (!sh?.cid) return false;
      // Undo snapshots are from before the step: move the op back through it. Redo snapshots are from after it.
      const through: Struct = { t: 'struct', s: sh.cid, nm: sh.name, a: m.axis, at: m.at, n: stack === 'undo' ? -m.count : m.count };
      ops = ops.flatMap((o) => xform(o, through, true));
      moving[stack] = ops;
    }
    for (const snap of step.snaps) {
      if (snap.kind !== 'book') {
        const j = remap(snap.sheet);
        if (j === null) return false;
        snap.sheet = j;
      }
      for (const o of ops) if (!patch(snap, o)) return false;
    }
    // Where the selection comes back to after undo.
    for (const o of ops) if (o.t === 'struct' && idsAfter[store.sheetIndex] === o.s) step.sel = shiftSelection(step.sel, o.a, o.at, o.n);
    return true;
  });
};
