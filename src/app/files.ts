import { readCsv, writeCsv } from '../formats/csv';
import { readXlsx, writeXlsx } from '../formats/xlsx';
import { encryptWorkbook, isEncryptedPackage } from '../formats/encryption';
import { isCfb, writeCfb } from '../formats/cfb';
import { isOldBiff, readOldBiff } from '../formats/biff-old-read';
import { PasswordRequired, readXls } from '../formats/biff8-read';
import { writeXls } from '../formats/biff8-write';
import { writeXlsb } from '../formats/xlsb-write';
import { readOds } from '../formats/ods-read';
import { writeOds } from '../formats/ods-write';
import { readXlsb } from '../formats/xlsb-read';
import { newWorkbook, cellKey, newSheet, internStyle, userSheetDefaults, type Workbook } from '../model/workbook';
import { displayText } from './ops/clipboard';
import type { GridStore } from './store';
import { N, T } from '../ui/i18n';

/** The desktop app's bridge (electron/preload.ts). Absent in the browser. */
interface DesktopBridge {
  open: () => Promise<{ path: string; name: string; data: Uint8Array } | null>;
  save: (path: string | null, name: string, data: Uint8Array) => Promise<{ path: string; name: string } | null>;
  launchFile: () => Promise<{ path: string; name: string; data: Uint8Array } | null>;
  /** Reads a file the person picked before (Get Data sources); null when not allowed. */
  readFile?: (path: string) => Promise<{ path: string; name: string; data: Uint8Array } | null>;
  newWindow: () => void;
  print: () => Promise<void>;
}
export const desktop = (globalThis as { clerdo?: DesktopBridge }).clerdo;

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MAX_BYTES = 512 * 1024 * 1024;

/** Kind of workbook in a zip package, from its parts. */
const zipKind = (data: Uint8Array, lower: string): 'ods' | 'xlsb' | 'xlsx' => {
  // OpenDocument packages start with an uncompressed "mimetype" entry
  const head = new TextDecoder('latin1').decode(data.subarray(0, 120));
  if (/mimetypeapplication\/vnd\.oasis\.opendocument\.spreadsheet/.test(head) || /\.(ods|ots)$/.test(lower)) return 'ods';
  const scan = new TextDecoder('latin1').decode(data.subarray(Math.max(0, data.length - 65536)));
  if (/xl\/workbook\.bin/.test(scan) || lower.endsWith('.xlsb')) return 'xlsb';
  return 'xlsx';
};

export const workbookFromBytes = (name: string, data: Uint8Array, password?: string): Workbook => {
  if (data.length > MAX_BYTES) throw new Error(T('This file is too large to open.'));
  const lower = name.toLowerCase();
  if (/\.(csv|tsv|txt|prn)$/.test(lower) && !isCfb(data) && !(data[0] === 0x50 && data[1] === 0x4b)) {
    const text = new TextDecoder().decode(data);
    const { rows, formats } = readCsv(text, lower.endsWith('.tsv') ? { delimiter: '\t' } : undefined);
    const wb = newWorkbook();
    const sheet = wb.sheets[0]!;
    sheet.name = name.replace(/\.[^.]+$/, '').slice(0, 31).replace(/[\\/?*[\]:]/g, '_') || 'Sheet1';
    rows.forEach((r, i) => r.forEach((v, j) => {
      if (v === null) return;
      const fmt = formats[i]?.[j];
      sheet.cells.set(cellKey(i, j), fmt ? { v, s: internStyle(wb, { numFmt: fmt }) } : { v });
    }));
    return wb;
  }
  if (isCfb(data)) {
    if (isEncryptedPackage(data)) throw new Error(T('This workbook is protected with a password.'));
    return readXls(data, { password });
  }
  if (isOldBiff(data)) return readOldBiff(data);
  // a bare record stream (no compound-file wrapper) of the 1995-2003 formats
  if (data[0] === 0x09 && data[1] === 0x08 && (data[2] === 8 || data[2] === 16) && data[3] === 0) return readXls(writeCfb(new Map([[N('Workbook'), data]])), { password });
  if (data[0] !== 0x50 || data[1] !== 0x4b) throw new Error(T('This file isn\'t a workbook ClerdoGrid can read.'));
  const kind = zipKind(data, lower);
  if (kind === 'ods') return readOds(data, { password });
  if (kind === 'xlsb') return readXlsb(data);
  return readXlsx(data);
};

/** Dialogs the formats feature shows while opening: text import options, and the password of an older workbook. */
export const formatOpenHandlers: {
  text: ((store: GridStore, file: { name: string; data: Uint8Array }) => void) | null;
  password: ((store: GridStore, file: { name: string; data: Uint8Array; path: string | null; handle?: unknown; wrong: boolean }) => void) | null;
} = { text: null, password: null };

/** File types Save can write, with their extension and media type. */
export const SAVE_TYPES: Record<string, { ext: string; mime: string; label: string }> = {
  xlsx: { ext: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', label: N('Workbook') },
  ods: { ext: 'ods', mime: 'application/vnd.oasis.opendocument.spreadsheet', label: N('OpenDocument spreadsheet') },
  ots: { ext: 'ots', mime: 'application/vnd.oasis.opendocument.spreadsheet-template', label: N('OpenDocument template') },
  xls: { ext: 'xls', mime: 'application/vnd.ms-excel', label: N('Older workbook') },
  xlsb: { ext: 'xlsb', mime: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12', label: N('Binary workbook') },
};

/** Asks for the password of an encrypted workbook; set by the workspace feature. */
export const encryptedOpenHandler: { current: ((store: GridStore, file: { name: string; data: Uint8Array; path: string | null; handle?: unknown }) => void) | null } = { current: null };

/** Run just before a workbook is written; each may return a function that undoes its temporary change. */
export const beforeSaveHooks: ((store: GridStore) => (() => void) | void)[] = [];

export const openBytes = (store: GridStore, name: string, data: Uint8Array, path: string | null = null, handle?: unknown): void => {
  if (encryptedOpenHandler.current && isEncryptedPackage(data)) { encryptedOpenHandler.current(store, { name, data, path, handle }); return; }
  if (formatOpenHandlers.text && /\.(txt|prn)$/i.test(name) && !isCfb(data) && !(data[0] === 0x50 && data[1] === 0x4b)) { formatOpenHandlers.text(store, { name, data }); return; }
  try {
    const wb = workbookFromBytes(name, data);
    const plain = /\.(csv|tsv|txt|prn|xltm|xlam)$/i.test(name);
    store.load(wb, name, plain ? null : path, plain ? undefined : handle);
  } catch (e) {
    if (e instanceof PasswordRequired && formatOpenHandlers.password) { formatOpenHandlers.password(store, { name, data, path, handle, wrong: false }); return; }
    store.notify(e instanceof Error ? e.message : T('The file could not be opened.'));
  }
};

export const newWorkbookFile = (store: GridStore): void => {
  if (desktop) { desktop.newWindow(); return; }
  if (store.file.dirty && !confirm(T('Start a new workbook? Unsaved changes in this one will be lost.'))) return;
  const wb = newWorkbook();
  for (const sh of wb.sheets) sh.view.rtl = userSheetDefaults.rtl;
  store.load(wb, 'Book1');
};

type PickerWindow = Window & {
  showOpenFilePicker?: (o: unknown) => Promise<{ getFile: () => Promise<File> }[]>;
  showSaveFilePicker?: (o: unknown) => Promise<{ createWritable: () => Promise<{ write: (d: Blob) => Promise<void>; close: () => Promise<void> }>; name: string }>;
};

export const OPEN_TYPES = [
  { description: N('Workbooks'), accept: {
    [XLSX_TYPE]: ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam'],
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12': ['.xlsb'],
    'application/vnd.ms-excel': ['.xls', '.xlt'],
    'application/vnd.oasis.opendocument.spreadsheet': ['.ods', '.ots'],
  } },
  { description: N('Text'), accept: { 'text/csv': ['.csv', '.tsv'], 'text/plain': ['.txt', '.prn'] } },
];
export const OPEN_ACCEPT = '.xlsx,.xlsm,.xltx,.xltm,.xlam,.xlsb,.xls,.xlt,.ods,.ots,.csv,.tsv,.txt,.prn';

export const openFile = async (store: GridStore): Promise<void> => {
  if (desktop) {
    try {
      const f = await desktop.open();
      if (f) openBytes(store, f.name, f.data, f.path);
    } catch (e) { store.notify(e instanceof Error ? e.message : T('The file could not be opened.')); }
    return;
  }
  const w = window as PickerWindow;
  if (w.showOpenFilePicker) {
    try {
      const [h] = await w.showOpenFilePicker({ types: OPEN_TYPES });
      const file = await h!.getFile();
      openBytes(store, file.name, new Uint8Array(await file.arrayBuffer()), null, h);
    } catch { /* cancelled */ }
    return;
  }
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = OPEN_ACCEPT;
  input.onchange = async () => {
    const file = input.files?.[0];
    if (file) openBytes(store, file.name, new Uint8Array(await file.arrayBuffer()));
  };
  input.click();
};

const download = (name: string, blob: Blob): void => {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30_000);
};

export const workbookBytes = (store: GridStore): Uint8Array => {
  store.sheet.view.selection = store.range;
  store.wb.props = { ...store.wb.props, modified: new Date().toISOString(), creator: store.wb.props?.creator ?? 'Clerdo Grid' };
  const undo = beforeSaveHooks.map((h) => h(store));
  try { return writeXlsx(store.wb); } finally { for (const u of undo.reverse()) u?.(); }
};

/** Bytes of the workbook in a save type (see SAVE_TYPES). */
export const bytesFor = (store: GridStore, format: string): Uint8Array => {
  if (format === 'xlsx') {
    const data = workbookBytes(store);
    return store.file.password ? encryptWorkbook(data, store.file.password) : data;
  }
  store.sheet.view.selection = store.range;
  store.wb.props = { ...store.wb.props, modified: new Date().toISOString(), creator: store.wb.props?.creator ?? 'Clerdo Grid' };
  const undo = beforeSaveHooks.map((h) => h(store));
  try {
    if (format === 'xlsb') { const data = writeXlsb(store.wb); return store.file.password ? encryptWorkbook(data, store.file.password) : data; }
    if (format === 'xls') return writeXls(store.wb, store.file.password ? { password: store.file.password } : {});
    if (format === 'ods' || format === 'ots') return writeOds(store.wb, { template: format === 'ots', ...(store.file.password ? { password: store.file.password } : {}) });
    throw new Error(T('unknown file type'));
  } finally { for (const u of undo.reverse()) u?.(); }
};

/** Lets a feature take over plain Save (the cloud build saves cloud workbooks there). Returns true when it saved. */
export const saveHandlers: { cloud: ((store: GridStore) => Promise<boolean>) | null } = { cloud: null };

/** Saves in the file's own type, or in `format` (which then becomes the file's type). */
export const saveFile = async (store: GridStore, saveAs = false, format?: string): Promise<void> => {
  store.commitEdit();
  if (!saveAs && !format && saveHandlers.cloud && await saveHandlers.cloud(store)) return;
  const type = format ?? (store.file.format && SAVE_TYPES[store.file.format] ? store.file.format : 'xlsx');
  const t = SAVE_TYPES[type]!;
  const changing = type !== (store.file.format ?? 'xlsx');
  let data: Uint8Array;
  try { data = bytesFor(store, type); } catch (e) { store.notify(T('The workbook could not be saved: {p0}', { p0: e instanceof Error ? e.message : T('unknown error') })); return; }
  const name = `${store.file.name || 'Book1'}.${t.ext}`;
  const extRe = new RegExp(`\\.${t.ext}$`, 'i');
  const done = (patch: Partial<typeof store.file>) => {
    store.file = { ...store.file, ...patch, dirty: false, lastSaved: Date.now(), ...(type === 'xlsx' ? { format: undefined } : { format: type }) };
    store.emit();
  };
  if (desktop) {
    try {
      const r = await desktop.save(saveAs || changing ? null : store.file.path, name, data);
      if (!r) return;
      done({ path: r.path, name: r.name.replace(/\.(xlsx|xlsb|xls|ods|ots)$/i, '') });
    } catch (e) { store.notify(e instanceof Error ? e.message : T('The workbook could not be saved.')); }
    return;
  }
  const w = window as PickerWindow;
  const blob = new Blob([data as BlobPart], { type: t.mime });
  if (w.showSaveFilePicker) {
    try {
      let handle = !saveAs && !changing ? (store.file.handle as Awaited<ReturnType<NonNullable<PickerWindow['showSaveFilePicker']>>> | undefined) : undefined;
      if (handle && (!('createWritable' in handle) || !extRe.test(handle.name))) handle = undefined;
      handle ??= await w.showSaveFilePicker({ suggestedName: name, types: [{ description: t.label, accept: { [t.mime]: [`.${t.ext}`] } }] });
      const wr = await handle.createWritable();
      await wr.write(blob);
      await wr.close();
      done({ handle, name: handle.name.replace(/\.(xlsx|xlsb|xls|ods|ots)$/i, '') });
      return;
    } catch (e) {
      if ((e as { name?: string }).name === 'AbortError') return;
    }
  }
  download(name, blob);
  done({});
};

/** Offers a file to save (exports). */
export const saveBlob = async (name: string, blob: Blob, description: string, ext: string): Promise<void> => {
  if (desktop) {
    try { await desktop.save(null, name, new Uint8Array(await blob.arrayBuffer())); } catch { /* cancelled */ }
    return;
  }
  const w = window as PickerWindow;
  if (w.showSaveFilePicker) {
    try {
      const handle = await w.showSaveFilePicker({ suggestedName: name, types: [{ description, accept: { [blob.type.split(';')[0] || 'application/octet-stream']: [`.${ext}`] } }] });
      const wr = await handle.createWritable();
      await wr.write(blob);
      await wr.close();
      return;
    } catch (e) { if ((e as { name?: string }).name === 'AbortError') return; }
  }
  download(name, blob);
};

export const exportCsv = (store: GridStore): void => {
  const b = store.engine.usedBounds(store.sheetIndex);
  const rows: string[][] = [];
  for (let r = 0; r <= b.rows; r++) {
    const line: string[] = [];
    for (let c = 0; c <= b.cols; c++) line.push(displayText(store, r, c));
    rows.push(line);
  }
  const csv = writeCsv(rows);
  download(`${store.file.name || 'Book1'} - ${store.sheet.name}.csv`, new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' }));
};

/** Prints the used range (or selection) of the sheet as an HTML table; "Save as PDF" in the print dialog gives a PDF. */
export const printSheet = (store: GridStore): void => {
  const r = store.range;
  const b = store.engine.usedBounds(store.sheetIndex);
  const area = r.r1 === r.r2 && r.c1 === r.c2 ? { r1: 0, c1: 0, r2: b.rows, c2: b.cols } : r;
  const esc = (s: string) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!);
  const rows: string[] = [];
  for (let row = area.r1; row <= area.r2; row++) {
    if (store.sheet.rows.get(row)?.hidden) continue;
    const cells: string[] = [];
    for (let col = area.c1; col <= area.c2; col++) {
      if (store.sheet.cols.get(col)?.hidden) continue;
      const st = store.style(row, col);
      const v = store.value(row, col);
      const css = [
        st.font?.bold && 'font-weight:700', st.font?.italic && 'font-style:italic', st.font?.color && `color:${st.font.color}`, st.fill?.color && `background:${st.fill.color}`,
        `text-align:${st.align?.h && st.align.h !== 'general' ? st.align.h : typeof v === 'number' ? 'right' : 'left'}`, st.font?.size && `font-size:${st.font.size}pt`,
        ...(['top', 'bottom', 'left', 'right'] as const).map((side) => st.border?.[side] && `border-${side}:${st.border[side]!.style === 'thin' ? 1 : 2}px solid ${st.border[side]!.color ?? '#000'}`),
      ].filter(Boolean).join(';');
      cells.push(`<td style="${css}">${esc(displayText(store, row, col))}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  const grid = store.sheet.view.showGrid ? 'td{border:1px solid #d9d9d9}' : '';
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${esc(store.file.name)} - ${esc(store.sheet.name)}</title><style>
body{font:11pt Carlito,Calibri,Arial,sans-serif;margin:0}table{border-collapse:collapse}td{padding:1px 4px;white-space:pre;${''}}${grid}
@page{size:${store.sheet.pageSetup?.orientation ?? 'portrait'};margin:0.7in}</style></head><body><table>${rows.join('')}</table></body></html>`;
  const frame = document.createElement('iframe');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.appendChild(frame);
  const doc = frame.contentDocument!;
  doc.open(); doc.write(html); doc.close();
  setTimeout(() => {
    frame.contentWindow!.focus();
    frame.contentWindow!.print();
    setTimeout(() => frame.remove(), 1000);
  }, 200);
};

// ---------- Autosave (browser and iPad) ----------

const DB = 'clerdogrid';
const idb = (): Promise<IDBDatabase> => new Promise((res, rej) => {
  const r = indexedDB.open(DB, 1);
  r.onupgradeneeded = () => r.result.createObjectStore('drafts');
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
});

const putDraft = async (data: { name: string; bytes: Uint8Array; at: number } | null): Promise<void> => {
  const db = await idb();
  await new Promise<void>((res, rej) => {
    const tx = db.transaction('drafts', 'readwrite');
    if (data) tx.objectStore('drafts').put(data, 'current'); else tx.objectStore('drafts').delete('current');
    tx.oncomplete = () => res();
    tx.onerror = () => rej(tx.error);
  });
  db.close();
};

const getDraft = async (): Promise<{ name: string; bytes: Uint8Array; at: number } | undefined> => {
  const db = await idb();
  const out = await new Promise<{ name: string; bytes: Uint8Array; at: number } | undefined>((res) => {
    const r = db.transaction('drafts').objectStore('drafts').get('current');
    r.onsuccess = () => res(r.result as never);
    r.onerror = () => res(undefined);
  });
  db.close();
  return out;
};

/** Keeps a copy of unsaved work in the browser so closing the tab or app on iPad doesn't lose it. */
export const startAutosave = (store: GridStore): (() => void) => {
  if (desktop || typeof indexedDB === 'undefined') return () => undefined;
  let last = -1;
  const t = setInterval(() => {
    const v = store.getVersion();
    if (v === last || store.editing) return;
    last = v;
    if (!store.file.dirty || store.file.password || store.readOnly) return;
    try { void putDraft({ name: store.file.name, bytes: workbookBytes(store), at: Date.now() }).catch(() => undefined); } catch { /* keep going */ }
  }, 4000);
  return () => clearInterval(t);
};

export const restoreDraft = async (store: GridStore): Promise<boolean> => {
  if (desktop || typeof indexedDB === 'undefined') return false;
  try {
    const d = await getDraft();
    if (!d) return false;
    store.load(readXlsx(d.bytes), d.name);
    store.file.dirty = true;
    store.notify(T('Restored your unsaved workbook.'));
    return true;
  } catch { return false; }
};

export const forgetDraft = (): void => { void putDraft(null).catch(() => undefined); };

export const openLaunchFile = async (store: GridStore): Promise<boolean> => {
  if (!desktop) return false;
  try {
    const f = await desktop.launchFile();
    if (!f) return false;
    openBytes(store, f.name, f.data, f.path);
    return true;
  } catch (e) { store.notify(e instanceof Error ? e.message : T('The file could not be opened.')); return false; }
};

export { newSheet };
