// Get Data: importing text files and ranges of other workbooks into this one, keeping a query that can be refreshed.
import { cellKey, colName, internStyle, keyCol, keyRow, newSheet, parseRange, type CellRange, type DataQuery, type Sheet, type Workbook } from '../../model/workbook';
import type { Scalar } from '../../engine/types';
import type { GridStore } from '../../app/store';
import { OPEN_ACCEPT, OPEN_TYPES, desktop, workbookFromBytes } from '../../app/files';
import { importText, type TextImportOptions } from '../../formats/text-import';
import { queriesFromXml, queriesToXml, QUERIES_NS } from '../../formats/queries-xml';
import { N, T } from '../../ui/i18n';

export interface PickedFile { name: string; data: Uint8Array; handle?: unknown; path?: string }

type PickerWindow = Window & { showOpenFilePicker?: (o: unknown) => Promise<{ getFile: () => Promise<File>; queryPermission?: (o: unknown) => Promise<string>; requestPermission?: (o: unknown) => Promise<string> }[]> };

/** Files picked this session, so Refresh can read them again without asking. */
const handles = new Map<string, unknown>();

export const pickFile = async (text: boolean): Promise<PickedFile | null> => {
  if (desktop) {
    const f = await desktop.open();
    return f ? { name: f.name, data: f.data, path: f.path } : null;
  }
  const w = window as PickerWindow;
  if (w.showOpenFilePicker) {
    try {
      const types = text ? [OPEN_TYPES[1]!, OPEN_TYPES[0]!] : OPEN_TYPES;
      const [h] = await w.showOpenFilePicker({ types });
      const file = await h!.getFile();
      return { name: file.name, data: new Uint8Array(await file.arrayBuffer()), handle: h };
    } catch { return null; }
  }
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = text ? '.csv,.tsv,.txt,.prn' : OPEN_ACCEPT;
    input.onchange = async () => {
      const file = input.files?.[0];
      resolve(file ? { name: file.name, data: new Uint8Array(await file.arrayBuffer()) } : null);
    };
    input.oncancel = () => resolve(null);
    input.click();
  });
};

// ---------- Remembered file handles (web) ----------

const HANDLE_DB = 'clerdogrid-sources';
const handleDb = (): Promise<IDBDatabase> => new Promise((res, rej) => {
  const r = indexedDB.open(HANDLE_DB, 1);
  r.onupgradeneeded = () => r.result.createObjectStore('handles');
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
});
const storeHandle = async (id: string, handle: unknown): Promise<void> => {
  if (typeof indexedDB === 'undefined') return;
  try {
    const db = await handleDb();
    await new Promise<void>((res, rej) => { const tx = db.transaction('handles', 'readwrite'); tx.objectStore('handles').put(handle, id); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); });
    db.close();
  } catch { /* handles that can't be stored are kept for this session only */ }
};
const loadHandle = async (id: string): Promise<unknown> => {
  if (typeof indexedDB === 'undefined') return undefined;
  try {
    const db = await handleDb();
    const v = await new Promise<unknown>((res) => { const r = db.transaction('handles').objectStore('handles').get(id); r.onsuccess = () => res(r.result); r.onerror = () => res(undefined); });
    db.close();
    return v;
  } catch { return undefined; }
};

type FileHandle = { getFile: () => Promise<File>; queryPermission?: (o: unknown) => Promise<string>; requestPermission?: (o: unknown) => Promise<string> };

/** Reads the query's file again: the desktop path, or a file handle from this or an earlier session; otherwise asks. */
const sourceAgain = async (q: DataQuery): Promise<PickedFile | null> => {
  if (desktop?.readFile && q.path) {
    try { const f = await desktop.readFile(q.path); if (f) return { name: f.name, data: f.data, path: f.path }; } catch { /* ask instead */ }
  }
  const h = (handles.get(q.id) ?? (await loadHandle(q.id))) as FileHandle | undefined;
  if (h && typeof h.getFile === 'function') {
    try {
      const ok = !h.queryPermission || (await h.queryPermission({ mode: 'read' })) === 'granted' || (await h.requestPermission?.({ mode: 'read' })) === 'granted';
      if (ok) { const f = await h.getFile(); return { name: f.name, data: new Uint8Array(await f.arrayBuffer()), handle: h }; }
    } catch { /* fall back to picking */ }
  }
  return pickFile(q.kind === 'text');
};

export interface Block { values: Scalar[][]; formats: (string | undefined)[][] }

/** Values (not formulas) of a range or table in another workbook. */
export const workbookBlock = (wb: Workbook, sheetName: string, range: string | undefined, values: (sheet: number, r: number, c: number) => Scalar): Block | string => {
  let si = wb.sheets.findIndex((s) => s.name === sheetName);
  let area: CellRange | null = null;
  if (range) {
    for (let i = 0; i < wb.sheets.length && !area; i++) {
      const t = wb.sheets[i]!.tables.find((x) => x.name.toUpperCase() === range.toUpperCase());
      if (t) { si = i; area = t.range; }
    }
    if (!area) {
      const name = wb.names.find((n) => n.name.toUpperCase() === range.toUpperCase());
      const text = name ? name.formula : range;
      const bang = text.lastIndexOf('!');
      if (bang > 0) {
        const sn = text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
        const k = wb.sheets.findIndex((s) => s.name === sn);
        if (k >= 0) si = k;
      }
      area = parseRange(text.slice(bang + 1).replace(/\$/g, ''));
      if (!area) return T('That range or table was not found in the file.');
    }
  }
  const sheet = wb.sheets[si];
  if (!sheet) return T('That sheet was not found in the file.');
  if (!area) {
    let r2 = -1, c2 = -1;
    for (const k of sheet.cells.keys()) { r2 = Math.max(r2, keyRow(k)); c2 = Math.max(c2, keyCol(k)); }
    if (r2 < 0) return T('That sheet is empty.');
    area = { r1: 0, c1: 0, r2, c2 };
  }
  const out: Block = { values: [], formats: [] };
  for (let r = area.r1; r <= area.r2; r++) {
    const vr: Scalar[] = [], fr: (string | undefined)[] = [];
    for (let c = area.c1; c <= area.c2; c++) {
      vr.push(values(si, r, c));
      const cell = sheet.cells.get(cellKey(r, c));
      fr.push(cell?.s ? wb.styles[cell.s]?.numFmt : undefined);
    }
    out.values.push(vr); out.formats.push(fr);
  }
  return out;
};

/** Writes a block at the query's target (clearing what the query wrote last time) as one undo step. */
export const placeBlock = (store: GridStore, q: DataQuery, block: Block, label: string): void => {
  const rows = block.values.length, cols = Math.max(0, ...block.values.map((r) => r.length));
  store.change(label, 'book', () => {
    let si = store.wb.sheets.findIndex((s) => s.name === q.target.sheet);
    if (si < 0) {
      const sheet = newSheet(uniqueName(store.wb, q.target.sheet || q.name));
      store.wb.sheets.push(sheet);
      si = store.wb.sheets.length - 1;
      q.target.sheet = sheet.name;
    }
    const sheet: Sheet = store.wb.sheets[si]!;
    const { row, col } = q.target;
    for (let r = 0; r < q.target.rows; r++) for (let c = 0; c < q.target.cols; c++) sheet.cells.delete(cellKey(row + r, col + c));
    block.values.forEach((vr, r) => vr.forEach((v, c) => {
      if (row + r >= 1_048_576 || col + c >= 16_384) return;
      const f = block.formats[r]?.[c];
      if (v === null && !f) return;
      sheet.cells.set(cellKey(row + r, col + c), f ? { v, s: internStyle(store.wb, { numFmt: f }) } : { v });
    }));
    q.target.rows = rows;
    q.target.cols = cols;
    q.refreshed = new Date().toISOString();
    const list = store.wb.queries ?? [];
    if (!list.some((x) => x.id === q.id)) list.push(q);
    store.wb.queries = list;
    store.wb.activeSheet = si;
    store.engine.sheetsChanged();
  });
  store.selectRange({ r1: q.target.row, c1: q.target.col, r2: q.target.row + Math.max(0, rows - 1), c2: q.target.col + Math.max(0, cols - 1) });
};

const uniqueName = (wb: Workbook, base: string): string => {
  const clean = base.replace(/\.[^.]+$/, '').replace(/[\\/?*[\]:]/g, '_').slice(0, 28) || 'Data';
  let name = clean, n = 2;
  while (wb.sheets.some((s) => s.name.toUpperCase() === name.toUpperCase())) name = `${clean.slice(0, 27)} ${n++}`;
  return name;
};

export const newQueryId = (): string => `q${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

export const rememberHandle = (q: DataQuery, file: PickedFile): void => {
  if (file.path) q.path = file.path;
  if (file.handle) { handles.set(q.id, file.handle); void storeHandle(q.id, file.handle); }
};

/** Values of another workbook as computed by a temporary engine. */
export const computedValues = async (wb: Workbook): Promise<(sheet: number, r: number, c: number) => Scalar> => {
  const { Engine } = await import('../../engine/engine');
  const engine = new Engine(wb);
  engine.recalc(true);
  return (sheet, r, c) => engine.getValue(sheet, r, c);
};

export const refreshQuery = async (store: GridStore, q: DataQuery): Promise<boolean> => {
  const file = await sourceAgain(q);
  if (!file) return false;
  rememberHandle(q, file);
  try {
    if (q.kind === 'text') {
      const o = JSON.parse(q.options ?? '{}') as TextImportOptions;
      placeBlock(store, q, importText(file.data, o), `Refresh ${q.name}`);
    } else {
      const wb = workbookFromBytes(file.name, file.data);
      const block = workbookBlock(wb, q.source?.sheet ?? '', q.source?.range, await computedValues(wb));
      if (typeof block === 'string') { store.notify(block); return false; }
      placeBlock(store, q, block, `Refresh ${q.name}`);
    }
    q.file = file.name;
    store.notify(T('Refreshed {name}.', { name: q.name }));
    return true;
  } catch (e) {
    store.notify(e instanceof Error ? e.message : T('The data could not be refreshed.'));
    return false;
  }
};

export const refreshAll = async (store: GridStore): Promise<void> => {
  for (const q of [...(store.wb.queries ?? [])]) if (!(await refreshQuery(store, q))) break;
};

export const deleteQuery = (store: GridStore, q: DataQuery): void => {
  store.change(N('Remove query'), 'book', () => {
    const list = (store.wb.queries ?? []).filter((x) => x.id !== q.id);
    if (list.length) store.wb.queries = list; else delete store.wb.queries;
  });
};

export const targetText = (q: DataQuery): string => `${q.target.sheet}!${colName(q.target.col)}${q.target.row + 1}`;

// ---------- Saving queries in .xlsx (as a custom XML item the reader keeps) ----------

const isQueriesPart = (bytes: Uint8Array): boolean => new TextDecoder().decode(bytes.subarray(0, 300)).includes(QUERIES_NS);

/** Before an .xlsx save: puts the queries part into the kept package parts; returns the undo. */
export const stashQueries = (store: GridStore): (() => void) | void => {
  const wb = store.wb;
  const parts = wb.extraParts ?? new Map<string, Uint8Array>();
  const before = wb.extraParts ? new Map(wb.extraParts) : undefined;
  let path = [...parts.keys()].find((p) => /^customXml\/item\d+\.xml$/i.test(p) && isQueriesPart(parts.get(p)!));
  if (!wb.queries?.length && !path) return;
  if (!path) {
    let n = 1;
    while (parts.has(`customXml/item${n}.xml`)) n++;
    path = `customXml/item${n}.xml`;
  }
  if (wb.queries?.length) parts.set(path, new TextEncoder().encode(queriesToXml(wb.queries)));
  else parts.delete(path);
  wb.extraParts = parts;
  return () => { if (before) wb.extraParts = before; else delete wb.extraParts; };
};

/** After an .xlsx load: reads queries saved in a custom XML item. */
export const loadQueries = (wb: Workbook): void => {
  if (wb.queries || !wb.extraParts) return;
  for (const [p, bytes] of wb.extraParts) {
    if (!/^customXml\/item\d+\.xml$/i.test(p) || !isQueriesPart(bytes)) continue;
    const list = queriesFromXml(new TextDecoder().decode(bytes));
    if (list.length) wb.queries = list;
    return;
  }
};
