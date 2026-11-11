import { allCommands, run } from '../app/commands';
import { MAX_COLS, MAX_ROWS } from '../engine/types';
import type { GridStore } from '../app/store';
import { T } from './i18n';

const isMac = typeof navigator !== 'undefined' && /Mac|iPad|iPhone/.test(navigator.platform || navigator.userAgent);

/** Normalises an event to the "Mod+Shift+K" form commands use. */
export const keyName = (e: KeyboardEvent): string => {
  const parts: string[] = [];
  if (isMac ? e.metaKey : e.ctrlKey) parts.push('Mod');
  if (isMac && e.ctrlKey) parts.push('Ctrl');
  if (e.altKey) parts.push('Alt');
  if (e.shiftKey) parts.push('Shift');
  let k = e.key;
  // Shifted punctuation: use the physical key so Mod+Shift+; works on any layout.
  if (e.code?.startsWith('Digit') && (e.shiftKey || e.altKey)) k = e.code.slice(5);
  else if (e.code === 'Semicolon') k = ';';
  else if (e.code === 'Equal') k = '=';
  else if (e.code === 'Minus') k = '-';
  else if (e.code === 'Backquote') k = '`';
  else if (e.code === 'Period') k = '.';
  else if (e.code === 'Comma') k = ',';
  else if (e.code === 'Slash') k = '/';
  else if (k === ' ') k = 'Space';
  else if (k.length === 1) k = k.toUpperCase();
  parts.push(k);
  return parts.join('+');
};

export const prettyKeys = (k: string): string => k.replace(/Mod/g, isMac ? '⌘' : T('Ctrl')).replace(/Shift/g, isMac ? '⇧' : 'Shift').replace(/Alt/g, isMac ? '⌥' : 'Alt').replace(/\+/g, isMac ? '' : '+');

/**
 * Row and column step for an arrow key. On a right-to-left sheet the columns run leftwards on screen, so
 * Left moves to the next column (towards B) and Right to the previous one: arrows always move the way they point.
 */
export const arrowDelta = (key: string, rtl: boolean): [number, number] | undefined => {
  switch (key) {
    case 'ArrowUp': return [-1, 0];
    case 'ArrowDown': return [1, 0];
    case 'ArrowLeft': return [0, rtl ? 1 : -1];
    case 'ArrowRight': return [0, rtl ? -1 : 1];
    default: return undefined;
  }
};

let byKey: Map<string, string> | null = null;
const commandFor = (k: string): string | undefined => {
  if (!byKey) { byKey = new Map(); for (const c of allCommands()) if (c.keys) for (const one of c.keys.split(',')) byKey.set(one.trim(), c.id); }
  return byKey.get(k);
};

/** Keys pressed while the grid has focus (not while typing in a cell). Returns true when handled. */
export const handleGridKey = (store: GridStore, e: KeyboardEvent, opts: { pageRows: number }): boolean => {
  const mod = isMac ? e.metaKey : e.ctrlKey;
  const { active } = store.sel;
  const d = arrowDelta(e.key, !!store.sheet.view.rtl);
  if (d && !e.altKey) {
    if (mod) {
      const from = e.shiftKey ? cornerOpposite(store) : active;
      const to = store.edge(from, d[0], d[1]);
      if (e.shiftKey) store.select(to, true); else store.select(to);
      return true;
    }
    if (e.shiftKey) { const c = cornerOpposite(store); store.select({ row: c.row + d[0], col: c.col + d[1] }, true); }
    else store.move(d[0], d[1]);
    return true;
  }
  switch (e.key) {
    case 'Enter':
      if (mod || e.altKey) break;
      store.moveAfterEnter(e.shiftKey);
      return true;
    case 'Tab':
      store.move(0, e.shiftKey ? -1 : 1, true);
      return true;
    case 'F2':
      if (e.shiftKey) break;
      store.beginEdit();
      return true;
    case 'Home':
      if (mod) store.select({ row: store.sheet.view.frozenRows, col: store.sheet.view.frozenCols }, e.shiftKey);
      else store.select({ row: active.row, col: 0 }, e.shiftKey);
      return true;
    case 'End':
      if (mod) { const b = store.engine.usedBounds(store.sheetIndex); store.select({ row: Math.max(0, b.rows), col: Math.max(0, b.cols) }, e.shiftKey); return true; }
      break;
    case 'PageDown': case 'PageUp':
      if (mod) break;
      if (e.altKey) store.select({ row: active.row, col: Math.max(0, Math.min(MAX_COLS - 1, active.col + (e.key === 'PageDown' ? 10 : -10))) }, e.shiftKey);
      else store.select({ row: Math.max(0, Math.min(MAX_ROWS - 1, active.row + (e.key === 'PageDown' ? opts.pageRows : -opts.pageRows))), col: active.col }, e.shiftKey);
      return true;
    case 'Backspace':
      if (mod) break;
      if (store.isLocked(active.row, active.col)) { store.notify(T('This cell is protected.')); return true; }
      run(store, 'edit.clearContents');
      store.beginEdit('');
      return true;
    case 'Delete':
      if (mod) break;
      run(store, 'edit.clearContents');
      return true;
    case 'Escape':
      if (store.menu || store.dialog) { store.menu = null; store.dialog = null; store.emit(); }
      return true;
    case 'ArrowDown':
      break;
  }
  if (e.altKey && e.key === 'ArrowDown') { store.beginEdit(); return true; }
  const k = keyName(e);
  const id = commandFor(k) ?? (k === 'Mod+Shift+=' || k === 'Mod+Shift+Plus' ? 'cells.insertRows' : undefined);
  if (id) {
    // Mod+Shift+= inserts cells for a cell selection, rows for whole rows.
    if (id === 'cells.insertRows' || id === 'cells.deleteRows') {
      const r = store.range;
      const wholeRows = r.c1 === 0 && r.c2 === MAX_COLS - 1, wholeCols = r.r1 === 0 && r.r2 === MAX_ROWS - 1;
      const insert = id === 'cells.insertRows';
      run(store, wholeRows ? id : wholeCols ? (insert ? 'cells.insertCols' : 'cells.deleteCols') : insert ? 'cells.insertDialog' : 'cells.deleteDialog');
      return true;
    }
    run(store, id);
    return true;
  }
  return false;
};

/** The corner of the selection opposite the anchor, which Shift+arrows move. */
const cornerOpposite = (store: GridStore): { row: number; col: number } => {
  const r = store.range, a = store.sel.anchor;
  return { row: a.row === r.r1 ? r.r2 : r.r1, col: a.col === r.c1 ? r.c2 : r.c1 };
};
