import type { GridStore } from '../../app/store';

export interface Script { id: string; name: string; code: string }

const KEY = 'clerdogrid.scriptLibrary';
const MAX_ITEMS = 200;

export const newScriptId = (): string => `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;

const valid = (x: unknown): x is Script => {
  const s = x as Script;
  return !!s && typeof s.id === 'string' && typeof s.name === 'string' && typeof s.code === 'string';
};

/** Scripts kept in this browser, shared by every workbook. */
export const loadLibrary = (): Script[] => {
  try {
    const raw = localStorage.getItem(KEY);
    const list: unknown = raw ? JSON.parse(raw) : [];
    return Array.isArray(list) ? list.filter(valid).slice(0, MAX_ITEMS).map((s) => ({ id: s.id.slice(0, 64), name: s.name.slice(0, 200), code: s.code.slice(0, 1_000_000) })) : [];
  } catch { return []; }
};

export const saveLibrary = (list: Script[]): boolean => {
  try { localStorage.setItem(KEY, JSON.stringify(list.slice(0, MAX_ITEMS))); return true; } catch { return false; }
};

export const uniqueScriptName = (names: string[], base: string): string => {
  const taken = new Set(names.map((n) => n.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
};

/** Workbook scripts travel with the file. Changing them marks the file as changed (not an undo step). */
export const bookScripts = (store: GridStore): Script[] => store.wb.scripts ?? [];
export const setBookScripts = (store: GridStore, list: Script[]): void => {
  store.wb.scripts = list.length ? list : undefined;
  store.file.dirty = true;
  store.emit();
};
