// Operations sent between copies, and how a finished change (or undo) becomes a list of them.
import type { ChangeEvent, GridStore, Snapshot } from '../../app/store';
import { cloneBook, inRange } from '../../app/store';
import { structureOnBook } from './apply';
import { placeSheets } from './sheets';
import { cellKey, keyCol, keyRow, type CellRange, type CellStyle, type Sheet, type Workbook } from '../../model/workbook';
import {
  ITEMS, KEYED, bookPropKeys, deepEq, encodeCell, encodeDim, encodeSheet, sameCell, sharedProp, sheetPropKeys, toWire,
  type ItemKey, type Json, type KeyedKey, type WCell, type WDim, type WSheet,
} from './wire';

export type CellEntry = [number, number, WCell | null];

/** `s` is the sheet's stable id (sheets.ts), so edits follow a sheet through concurrent renames, moves and inserts. */
export type Op =
  | { t: 'cells'; s: string; c: CellEntry[] }
  | { t: 'keyed'; s: string; k: KeyedKey; c: [number, number, Json][] }
  | { t: 'dim'; s: string; a: 'row' | 'col'; c: [number, WDim | null][] }
  | { t: 'item'; s: string; k: ItemKey; id: string; v: Json; i: number }
  | { t: 'prop'; s: string; k: string; v: Json }
  /** `nm` is the sheet's name, for adjusting formulas that point at it. */
  | { t: 'struct'; s: string; nm: string; a: 'row' | 'col'; at: number; n: number }
  | { t: 'sheetAdd'; s: string; name: string; pos: string; d: WSheet }
  | { t: 'sheetDel'; s: string; name: string }
  | { t: 'sheetRen'; s: string; from: string; to: string }
  | { t: 'sheetPos'; s: string; pos: string }
  | { t: 'book'; k: string; v: Json };

export const OP_TYPES = ['cells', 'keyed', 'dim', 'item', 'prop', 'struct', 'sheetAdd', 'sheetDel', 'sheetRen', 'sheetPos', 'book'] as const;

interface StructMeta { kind: 'struct'; sheet: number; name?: string; cid?: string; axis: 'row' | 'col'; at: number; count: number }
const isStructMeta = (m: unknown): m is StructMeta => !!m && typeof m === 'object' && (m as StructMeta).kind === 'struct';

/** Parts of a sheet captured when a cell-scoped change starts, in case it touched them too (paste can add notes or merges). */
const LIGHT = ['merges', 'conditionalFormats', 'validations', 'tables', 'autoFilter', 'notes', 'hyperlinks', 'comments'] as const;

export const captureLight = (store: GridStore, e: ChangeEvent, scope: { area: CellRange; sheet?: number }[] | 'sheet' | 'book'): void => {
  if (!Array.isArray(scope)) return;
  const light = new Map<Sheet, Partial<Sheet>>();
  for (const s of scope) {
    const sheet = store.wb.sheets[s.sheet ?? store.sheetIndex];
    if (!sheet || light.has(sheet)) continue;
    const part: Record<string, unknown> = {};
    for (const k of LIGHT) part[k] = structuredClone((sheet as unknown as Record<string, unknown>)[k]);
    light.set(sheet, part as Partial<Sheet>);
  }
  e.extra.collabLight = light;
};

interface Base {
  full?: { sheet: Sheet; styles: CellStyle[] };
  cells: Map<number, WCell | null>;
  areas: CellRange[];
}

const covered = (areas: CellRange[], k: number) => areas.some((a) => inRange(a, keyRow(k), keyCol(k)));
const areaSize = (a: CellRange) => (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);

export const opsFromEvent = (store: GridStore, e: ChangeEvent): Op[] => {
  const wb = store.wb;
  // A row/column insert or delete on its own goes as one operation, so others can transform their edits around it.
  if (e.kind === 'change' && e.meta.length === 1 && isStructMeta(e.meta[0]) && e.snaps.length === 1 && e.snaps[0]!.kind === 'book') {
    const m = e.meta[0];
    const sheet = wb.sheets[m.sheet];
    if (sheet?.cid) return [{ t: 'struct', s: sheet.cid, nm: sheet.name, a: m.axis, at: m.at, n: m.count }];
  }
  // Undoing or redoing one: the reverse (or same) insert/delete, then whatever else differs (restored cells and parts).
  if (e.kind !== 'change' && e.meta.length === 1 && isStructMeta(e.meta[0]) && e.snaps.length === 1 && e.snaps[0]!.kind === 'book') {
    const m = e.meta[0];
    const before = e.snaps[0]!.data;
    let si = m.cid !== undefined ? before.sheets.findIndex((x) => x.cid === m.cid) : -1;
    if (si < 0 && m.name !== undefined) si = before.sheets.findIndex((x) => x.name === m.name);
    if (si < 0 && before.sheets[m.sheet]) si = m.sheet;
    if (si >= 0 && before.sheets[si]!.cid) {
      const sh = before.sheets[si]!;
      const struct: Op = { t: 'struct', s: sh.cid!, nm: sh.name, a: m.axis, at: m.at, n: e.kind === 'undo' ? -m.count : m.count };
      const baseline = cloneBook(before);
      structureOnBook(baseline, si, struct as Extract<Op, { t: 'struct' }>);
      const rest = opsFromEvent(store, { kind: 'change', label: e.label, snaps: [{ kind: 'book', data: baseline }], sheetsBefore: baseline.sheets, meta: [], extra: {} });
      return [struct, ...rest];
    }
  }
  const bases = new Map<Sheet, Base>();
  const baseOf = (ref: Sheet | undefined): Base | null => {
    if (!ref) return null;
    let b = bases.get(ref);
    if (!b) { b = { cells: new Map(), areas: [] }; bases.set(ref, b); }
    return b;
  };
  let bookBase: Workbook | null = null;
  const refAt = (i: number) => e.sheetsBefore[i] ?? wb.sheets[i];
  for (const snap of e.snaps) collect(snap);
  function collect(snap: Snapshot) {
    if (snap.kind === 'cells') {
      const b = baseOf(refAt(snap.sheet));
      if (!b || b.full) return;
      for (const [k, c] of snap.cells) if (!b.cells.has(k) && !covered(b.areas, k)) b.cells.set(k, encodeCell(c, wb.styles));
      b.areas.push(snap.area);
    } else if (snap.kind === 'sheet') {
      const b = baseOf(refAt(snap.sheet));
      if (b && !b.full) b.full = { sheet: snap.data, styles: wb.styles };
    } else {
      bookBase ??= snap.data;
      snap.data.sheets.forEach((s, i) => { const b = baseOf(refAt(i)); if (b && !b.full) b.full = { sheet: s, styles: snap.data.styles }; });
    }
  }

  const ops: Op[] = [];
  const light = (e.extra.collabLight as Map<Sheet, Partial<Sheet>> | undefined) ?? new Map();

  // Which sheet now is which sheet before: same object, else same id (undo restores copies), else same name.
  const before = e.sheetsBefore;
  const beforeSheet = (ref: Sheet) => bases.get(ref)?.full?.sheet ?? ref;
  const match = new Map<Sheet, Sheet>(); // after -> before
  const usedBefore = new Set<Sheet>();
  const take = (s: Sheet, b: Sheet) => { match.set(s, b); usedBefore.add(b); };
  for (const s of wb.sheets) if (before.includes(s) && !usedBefore.has(s)) take(s, s);
  for (const s of wb.sheets) {
    if (match.has(s) || !s.cid) continue;
    const b = before.find((x) => !usedBefore.has(x) && beforeSheet(x).cid === s.cid);
    if (b) take(s, b);
  }
  for (const s of wb.sheets) {
    if (match.has(s) || s.cid) continue;
    const b = before.find((x) => !usedBefore.has(x) && !beforeSheet(x).cid && beforeSheet(x).name === s.name);
    if (b) take(s, b);
  }
  if (bookBase) placeSheets(wb.sheets, new Set(match.keys()));

  for (const b of before) if (!usedBefore.has(b) && bookBase && beforeSheet(b).cid) ops.push({ t: 'sheetDel', s: beforeSheet(b).cid!, name: beforeSheet(b).name });
  for (const [after, b] of match) {
    const was = beforeSheet(b);
    if (was.name !== after.name) ops.push({ t: 'sheetRen', s: after.cid!, from: was.name, to: after.name });
  }

  for (const [after, b] of match) {
    if (!after.cid) continue;
    const base = bases.get(b);
    if (base) diffSheet(after, base, wb, ops);
    const l = light.get(b);
    if (l && !base?.full) diffParts(after, l as Sheet, wb.styles, wb, ops, LIGHT as unknown as string[]);
  }

  for (const s of wb.sheets) if (!match.has(s)) ops.push({ t: 'sheetAdd', s: s.cid!, name: s.name, pos: s.cpos!, d: encodeSheet(s, wb.styles) });
  if (bookBase) for (const [after, b] of match) if (after.cpos !== beforeSheet(b).cpos) ops.push({ t: 'sheetPos', s: after.cid!, pos: after.cpos! });

  if (bookBase) {
    const bb = bookBase as Workbook;
    for (const k of bookPropKeys(bb, wb)) {
      if (k === 'extraParts') continue;
      const a = (bb as unknown as Record<string, unknown>)[k], c = (wb as unknown as Record<string, unknown>)[k];
      if (!deepEq(a, c)) ops.push({ t: 'book', k, v: toWire(c) });
    }
  }
  return ops;
};

const diffSheet = (after: Sheet, base: Base, wb: Workbook, ops: Op[]): void => {
  const changes: CellEntry[] = [];
  const check = (k: number) => {
    const now = encodeCell(after.cells.get(k), wb.styles);
    const was = base.cells.has(k) ? base.cells.get(k)! : covered(base.areas, k) ? null : base.full ? encodeCell(base.full.sheet.cells.get(k), base.full.styles) : now;
    if (!sameCell(was, now)) changes.push([keyRow(k), keyCol(k), now]);
  };
  if (base.full) {
    const keys = new Set<number>([...base.full.sheet.cells.keys(), ...after.cells.keys(), ...base.cells.keys()]);
    for (const k of keys) check(k);
  } else {
    const keys = new Set<number>(base.cells.keys());
    for (const a of base.areas) {
      if (areaSize(a) > after.cells.size) { for (const k of after.cells.keys()) if (inRange(a, keyRow(k), keyCol(k))) keys.add(k); }
      else for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) if (after.cells.has(cellKey(r, c))) keys.add(cellKey(r, c));
    }
    for (const k of keys) check(k);
  }
  if (changes.length) ops.push({ t: 'cells', s: after.cid!, c: changes });
  if (base.full) diffParts(after, base.full.sheet, base.full.styles, wb, ops);
};

/** Rows, columns, keyed maps, item lists and props of a sheet against an earlier copy. `only` limits the parts compared. */
const diffParts = (after: Sheet, was: Sheet, wasStyles: CellStyle[], wb: Workbook, ops: Op[], only?: string[]): void => {
  const want = (k: string) => !only || only.includes(k);
  if (want('rows') || want('cols')) {
    for (const a of ['row', 'col'] as const) {
      const m1 = a === 'row' ? was.rows : was.cols, m2 = a === 'row' ? after.rows : after.cols;
      if (!m1 || !m2) continue;
      const c: [number, WDim | null][] = [];
      for (const i of new Set([...m1.keys(), ...m2.keys()])) {
        const x = encodeDim(m1.get(i), wasStyles), y = encodeDim(m2.get(i), wb.styles);
        if (!deepEq(x, y)) c.push([i, y]);
      }
      if (c.length) ops.push({ t: 'dim', s: after.cid!, a, c });
    }
  }
  for (const k of KEYED) {
    if (!want(k)) continue;
    const m1 = (was[k] ?? new Map()) as Map<number, unknown>, m2 = (after[k] ?? new Map()) as Map<number, unknown>;
    const c: [number, number, Json][] = [];
    for (const key of new Set([...m1.keys(), ...m2.keys()])) if (!deepEq(m1.get(key), m2.get(key))) c.push([keyRow(key), keyCol(key), toWire(m2.get(key))]);
    if (c.length) ops.push({ t: 'keyed', s: after.cid!, k, c });
  }
  for (const k of ITEMS) {
    if (!want(k)) continue;
    const l1 = (was[k] ?? []) as { id: string }[], l2 = (after[k] ?? []) as { id: string }[];
    if (deepEq(l1, l2)) continue;
    const ids1 = l1.map((x) => x.id), ids2 = l2.map((x) => x.id);
    const kept1 = ids1.filter((id) => ids2.includes(id)), kept2 = ids2.filter((id) => ids1.includes(id));
    if (kept1.join('\u0000') !== kept2.join('\u0000') || new Set(ids2).size !== ids2.length) { ops.push({ t: 'prop', s: after.cid!, k, v: toWire(l2) }); continue; }
    for (const id of ids1) if (!ids2.includes(id)) ops.push({ t: 'item', s: after.cid!, k, id, v: null, i: 0 });
    l2.forEach((x, i) => { const old = l1.find((y) => y.id === x.id); if (!old || !deepEq(old, x)) ops.push({ t: 'item', s: after.cid!, k, id: x.id, v: toWire(x), i }); });
  }
  for (const k of only ? only.filter((x) => !(KEYED as readonly string[]).includes(x) && x !== 'rows' && x !== 'cols') : sheetPropKeys(was, after)) {
    const a = sharedProp(was, k), b = sharedProp(after, k);
    if (!deepEq(a, b)) ops.push({ t: 'prop', s: after.cid!, k, v: toWire(b) });
  }
};

/** Ops for loading a whole workbook over the current one are never needed: a full snapshot is sent instead. */
export const opsSize = (ops: Op[]): number => JSON.stringify(ops).length;
