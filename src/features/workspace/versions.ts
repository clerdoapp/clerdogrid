// Version history: automatic local snapshots of the open workbook, kept in IndexedDB per file.
import { unzipSync } from 'fflate';
import { sha256, toBase64 } from '../../formats/crypto';
import { readXlsx, writeXlsx } from '../../formats/xlsx';
import type { GridStore } from '../../app/store';
import type { Workbook } from '../../model/workbook';
import { settings } from './settings';
import { T } from '../../ui/i18n';

export interface VersionMeta {
  id: string;
  file: string;
  at: number;
  size: number;
  hash: string;
  name?: string;
  reason: 'save' | 'auto' | 'restore' | 'manual';
}
interface VersionRecord extends VersionMeta { bytes: Uint8Array }

const DB = 'clerdogrid-versions';
const STORE = 'versions';
const MAX_TOTAL_BYTES = 400 * 1024 * 1024;

let dbPromise: Promise<IDBDatabase> | null = null;
const db = (): Promise<IDBDatabase> => {
  if (typeof indexedDB === 'undefined') return Promise.reject(new Error(T('No local storage')));
  dbPromise ??= new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => {
      const os = r.result.createObjectStore(STORE, { keyPath: 'id' });
      os.createIndex('file', 'file');
    };
    r.onsuccess = () => res(r.result);
    r.onerror = () => { dbPromise = null; rej(r.error); };
  });
  return dbPromise;
};

const tx = async <T>(mode: IDBTransactionMode, fn: (os: IDBObjectStore) => IDBRequest<T> | void): Promise<T | undefined> => {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(STORE, mode);
    const req = fn(t.objectStore(STORE));
    t.oncomplete = () => res(req ? req.result : undefined);
    t.onerror = () => rej(t.error);
    t.onabort = () => rej(t.error);
  });
};

/** Identity of the open file: its path on the desktop, otherwise its name. */
export const fileKey = (store: GridStore): string => store.file.path ? `path:${store.file.path}` : `name:${store.file.name || 'Book1'}`;

/** A hash of the workbook's contents that ignores zip timestamps and save times, so unchanged work isn't stored twice. */
export const contentHash = (bytes: Uint8Array): string => {
  const files = unzipSync(bytes);
  const enc = new TextEncoder();
  const parts: Uint8Array[] = [];
  for (const name of Object.keys(files).sort()) {
    if (name === 'docProps/core.xml' || name === 'docProps/app.xml') continue;
    parts.push(enc.encode(name), files[name]!);
  }
  let n = 0;
  for (const p of parts) n += p.length;
  const all = new Uint8Array(n);
  let o = 0;
  for (const p of parts) { all.set(p, o); o += p.length; }
  return toBase64(sha256(all));
};

export const snapshotBytes = (wb: Workbook): Uint8Array => writeXlsx(wb);

export const listVersions = async (file: string): Promise<VersionMeta[]> => {
  const all = await tx<VersionRecord[]>('readonly', (os) => os.index('file').getAll(file) as IDBRequest<VersionRecord[]>);
  return (all ?? []).map(({ bytes: _b, ...m }) => m).sort((a, b) => b.at - a.at);
};

export const getVersion = async (id: string): Promise<VersionRecord | undefined> => tx<VersionRecord>('readonly', (os) => os.get(id) as IDBRequest<VersionRecord>);

export const deleteVersion = async (id: string): Promise<void> => { await tx('readwrite', (os) => { os.delete(id); }); };

export const renameVersion = async (id: string, name: string): Promise<void> => {
  const v = await getVersion(id);
  if (!v) return;
  const n = name.trim().slice(0, 120);
  await tx('readwrite', (os) => { os.put({ ...v, name: n || undefined }); });
};

/** Drops the oldest unnamed versions beyond the limit for a file, and beyond the overall storage budget. */
export const prune = async (file: string, keep = settings.versionsKeep): Promise<void> => {
  const list = await listVersions(file);
  const unnamed = list.filter((v) => !v.name);
  const extra = unnamed.slice(Math.max(1, keep));
  for (const v of extra) await deleteVersion(v.id);
  const everything = await tx<VersionRecord[]>('readonly', (os) => os.getAll() as IDBRequest<VersionRecord[]>);
  let total = 0;
  const sorted = (everything ?? []).sort((a, b) => b.at - a.at);
  for (const v of sorted) {
    total += v.size;
    if (total > MAX_TOTAL_BYTES && !v.name) await deleteVersion(v.id);
  }
};

/** Stores a snapshot unless it matches the newest one for this file. Returns the new version or null. */
export const addVersion = async (file: string, bytes: Uint8Array, reason: VersionMeta['reason'], name?: string): Promise<VersionMeta | null> => {
  const hash = contentHash(bytes);
  const list = await listVersions(file);
  if (list[0]?.hash === hash && !name) return null;
  const meta: VersionMeta = { id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`, file, at: Date.now(), size: bytes.length, hash, reason, ...(name ? { name } : {}) };
  await tx('readwrite', (os) => { os.put({ ...meta, bytes }); });
  await prune(file);
  return meta;
};

export const loadVersion = async (id: string): Promise<{ meta: VersionMeta; wb: Workbook } | null> => {
  const v = await getVersion(id);
  if (!v) return null;
  const { bytes, ...meta } = v;
  return { meta, wb: readXlsx(bytes) };
};

/** Moves a file's versions to its new identity after Save as. */
export const moveVersions = async (from: string, to: string): Promise<void> => {
  if (from === to) return;
  const all = await tx<VersionRecord[]>('readonly', (os) => os.index('file').getAll(from) as IDBRequest<VersionRecord[]>);
  if (!all?.length) return;
  await tx('readwrite', (os) => { for (const v of all) os.put({ ...v, file: to }); });
};

export const formatSize = (n: number): string => (n < 1024 ? `${n} B` : n < 1024 * 1024 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`);

export const formatWhen = (at: number): string => {
  const d = new Date(at);
  const now = new Date();
  const time = d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  if (d.toDateString() === now.toDateString()) return `Today ${time}`;
  const y = new Date(now); y.setDate(now.getDate() - 1);
  if (d.toDateString() === y.toDateString()) return `Yesterday ${time}`;
  return `${d.toLocaleDateString([], { month: 'short', day: 'numeric', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' })} ${time}`;
};
