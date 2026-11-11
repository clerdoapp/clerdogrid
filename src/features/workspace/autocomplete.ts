// AutoComplete for typed values: when what's typed starts exactly one entry already in the column,
// the rest of that entry is filled in and selected, so typing on replaces it and Enter accepts it.
import { MAX_ROWS } from '../../engine/types';
import type { GridStore } from '../../app/store';
import { settings } from './settings';

/** Column entries around the cell (the contiguous block above and below), case-insensitive unique. */
const columnEntries = (store: GridStore, row: number, col: number): string[] => {
  const out: string[] = [];
  const seen = new Set<string>();
  const add = (r: number): boolean => {
    const v = store.value(r, col);
    if (v === null) return false;
    if (typeof v === 'string' && v.length > 0) { const k = v.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(v); } }
    return true;
  };
  for (let r = row - 1, n = 0; r >= 0 && n < 5000; r--, n++) if (!add(r)) break;
  for (let r = row + 1, n = 0; r < MAX_ROWS && n < 5000; r++, n++) if (!add(r)) break;
  return out;
};

/** The full entry to suggest for this typed text in this cell, or null. */
export const suggestionFor = (store: GridStore, row: number, col: number, typed: string): string | null => {
  if (!typed || typed.startsWith('=') || /^[\d.,+\-$%()\s]+$/.test(typed) || typed.includes('\n')) return null;
  const lower = typed.toLowerCase();
  const hits = columnEntries(store, row, col).filter((v) => v.length > typed.length && v.toLowerCase().startsWith(lower));
  // Only a single unambiguous completion.
  return hits.length === 1 ? typed + hits[0]!.slice(typed.length) : null;
};

/** Watches typing in a cell and fills in the suggested rest of the entry. */
export const installAutoComplete = (store: GridStore): (() => void) => {
  let last: { row: number; col: number; text: string } | null = null;
  let filled: { typed: string; full: string } | null = null;
  return store.subscribe(() => {
    const e = store.editing;
    if (!e || e.inBar || !e.replace || !settings.autocomplete) { last = null; filled = null; return; }
    const prev = last && last.row === e.row && last.col === e.col ? last.text : '';
    last = { row: e.row, col: e.col, text: e.text };
    if (filled && e.text === filled.full) return;
    // A key that reached the grid before the selection was placed lands after the filled-in part: treat it as typed over it.
    if (filled && e.text.length > filled.full.length && e.text.startsWith(filled.full)) {
      const typed = filled.typed + e.text.slice(filled.full.length);
      filled = null;
      last = { row: e.row, col: e.col, text: typed.slice(0, -1) };
      store.setEditText(typed);
      return;
    }
    // Only after a character was added at the end: Backspace and Delete leave the text as it is.
    const base = filled ? filled.typed : prev;
    const grew = e.text.length > base.length && e.text.startsWith(base);
    if (!grew) { filled = null; return; }
    const sug = suggestionFor(store, e.row, e.col, e.text);
    if (!sug) { filled = null; return; }
    filled = { typed: e.text, full: sug };
    last = { row: e.row, col: e.col, text: sug };
    store.setEditText(sug);
    const from = e.text.length;
    if (typeof requestAnimationFrame !== 'undefined') requestAnimationFrame(() => {
      const ta = document.querySelector<HTMLTextAreaElement>('.celleditor');
      if (ta && ta.value === sug) ta.setSelectionRange(from, sug.length, 'backward');
    });
  });
};

/** For tests: what the current suggestion would be. */
export const completionFor = (store: GridStore): string | null => {
  const e = store.editing;
  if (!e || !settings.autocomplete) return null;
  return suggestionFor(store, e.row, e.col, e.text);
};
