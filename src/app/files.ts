import { readCsv, writeCsv } from '../formats/csv';
import { readXlsx, writeXlsx } from '../formats/xlsx';
import { newWorkbook, cellKey, newSheet, internStyle, type Workbook } from '../model/workbook';
import { displayText } from './ops/clipboard';
import type { GridStore } from './store';

/** The desktop app's bridge (electron/preload.ts). Absent in the browser. */
interface DesktopBridge {
  open: () => Promise<{ path: string; name: string; data: Uint8Array } | null>;
  save: (path: string | null, name: string, data: Uint8Array) => Promise<{ path: string; name: string } | null>;
  launchFile: () => Promise<{ path: string; name: string; data: Uint8Array } | null>;
  newWindow: () => void;
  print: () => Promise<void>;
}
export const desktop = (globalThis as { clerdo?: DesktopBridge }).clerdo;

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const MAX_BYTES = 512 * 1024 * 1024;

export const workbookFromBytes = (name: string, data: Uint8Array): Workbook => {
  if (data.length > MAX_BYTES) throw new Error('This file is too large to open.');
  const lower = name.toLowerCase();
  if (/\.(csv|tsv|txt)$/.test(lower)) {
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
  if (data[0] !== 0x50 || data[1] !== 0x4b) {
    if (data[0] === 0xd0 && data[1] === 0xcf) throw new Error('This is an older .xls file or a password-protected workbook. Save it as .xlsx in the other app first.');
    throw new Error("This file isn't a workbook ClerdoGrid can read.");
  }
  return readXlsx(data);
};

export const openBytes = (store: GridStore, name: string, data: Uint8Array, path: string | null = null, handle?: unknown): void => {
  try {
    const wb = workbookFromBytes(name, data);
    store.load(wb, name, /\.(csv|tsv|txt)$/i.test(name) ? null : path, /\.(csv|tsv|txt)$/i.test(name) ? undefined : handle);
  } catch (e) {
    store.notify(e instanceof Error ? e.message : 'The file could not be opened.');
  }
};

export const newWorkbookFile = (store: GridStore): void => {
  if (desktop) { desktop.newWindow(); return; }
  if (store.file.dirty && !confirm('Start a new workbook? Unsaved changes in this one will be lost.')) return;
  store.load(newWorkbook(), 'Book1');
};

type PickerWindow = Window & {
  showOpenFilePicker?: (o: unknown) => Promise<{ getFile: () => Promise<File> }[]>;
  showSaveFilePicker?: (o: unknown) => Promise<{ createWritable: () => Promise<{ write: (d: Blob) => Promise<void>; close: () => Promise<void> }>; name: string }>;
};

export const openFile = async (store: GridStore): Promise<void> => {
  if (desktop) {
    try {
      const f = await desktop.open();
      if (f) openBytes(store, f.name, f.data, f.path);
    } catch (e) { store.notify(e instanceof Error ? e.message : 'The file could not be opened.'); }
    return;
  }
  const w = window as PickerWindow;
  if (w.showOpenFilePicker) {
    try {
      const [h] = await w.showOpenFilePicker({ types: [{ description: 'Workbooks', accept: { [XLSX_TYPE]: ['.xlsx', '.xlsm', '.xltx'], 'text/csv': ['.csv', '.tsv', '.txt'] } }] });
      const file = await h!.getFile();
      openBytes(store, file.name, new Uint8Array(await file.arrayBuffer()), null, h);
    } catch { /* cancelled */ }
    return;
  }
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xlsm,.xltx,.csv,.tsv,.txt';
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
  store.wb.props = { ...store.wb.props, modified: new Date().toISOString(), creator: store.wb.props?.creator ?? 'ClerdoGrid' };
  return writeXlsx(store.wb);
};

export const saveFile = async (store: GridStore, saveAs = false): Promise<void> => {
  store.commitEdit();
  let data: Uint8Array;
  try { data = workbookBytes(store); } catch (e) { store.notify(`The workbook could not be saved: ${e instanceof Error ? e.message : 'unknown error'}`); return; }
  const name = `${store.file.name || 'Book1'}.xlsx`;
  if (desktop) {
    try {
      const r = await desktop.save(saveAs ? null : store.file.path, name, data);
      if (!r) return;
      store.file = { ...store.file, path: r.path, name: r.name.replace(/\.xlsx$/i, ''), dirty: false, lastSaved: Date.now() };
      store.emit();
    } catch (e) { store.notify(e instanceof Error ? e.message : 'The workbook could not be saved.'); }
    return;
  }
  const w = window as PickerWindow;
  const blob = new Blob([data as BlobPart], { type: XLSX_TYPE });
  if (w.showSaveFilePicker) {
    try {
      let handle = !saveAs ? (store.file.handle as Awaited<ReturnType<NonNullable<PickerWindow['showSaveFilePicker']>>> | undefined) : undefined;
      if (handle && !('createWritable' in handle)) handle = undefined;
      handle ??= await w.showSaveFilePicker({ suggestedName: name, types: [{ description: 'Workbook', accept: { [XLSX_TYPE]: ['.xlsx'] } }] });
      const wr = await handle.createWritable();
      await wr.write(blob);
      await wr.close();
      store.file = { ...store.file, handle, name: handle.name.replace(/\.xlsx$/i, ''), dirty: false, lastSaved: Date.now() };
      store.emit();
      return;
    } catch (e) {
      if ((e as { name?: string }).name === 'AbortError') return;
    }
  }
  download(name, blob);
  store.file = { ...store.file, dirty: false, lastSaved: Date.now() };
  store.emit();
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
    if (!store.file.dirty) return;
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
    store.notify('Restored your unsaved workbook.');
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
  } catch (e) { store.notify(e instanceof Error ? e.message : 'The file could not be opened.'); return false; }
};

export { newSheet };
