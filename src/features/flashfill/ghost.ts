// Flash Fill suggestions: after two examples are typed beside a block of data, the rest of the column
// shows the suggested values in grey. Enter accepts them (one undo step), Esc or any edit dismisses.
import type { GridStore } from '../../app/store';
import { applyFlashPlan, planFlashFill, type FlashPlan } from '../../app/ops/flash';

export const FLASH_AUTO_KEY = 'clerdogrid.flashAuto';

/** Automatic suggestions are on unless turned off in Options (or with the clerdogrid.flashAuto key). */
export const flashAutoOn = (): boolean => {
  try {
    const own = localStorage.getItem(FLASH_AUTO_KEY);
    if (own !== null) return own !== '0' && own !== 'false';
    const raw = localStorage.getItem('clerdogrid.options');
    if (raw) { const o = JSON.parse(raw) as { flashFillAuto?: boolean }; if (typeof o.flashFillAuto === 'boolean') return o.flashFillAuto; }
  } catch { /* no storage */ }
  return true;
};

export const ghost = { plan: null as FlashPlan | null, version: -1 };

let lastEdit: { sheet: number; row: number; col: number } | null = null;
let busy = false;

export const dismiss = (store: GridStore): void => {
  if (!ghost.plan) return;
  ghost.plan = null;
  store.emit();
};

export const accept = (store: GridStore): boolean => {
  const plan = ghost.plan;
  if (!plan || plan.sheet !== store.sheetIndex) return false;
  ghost.plan = null;
  applyFlashPlan(store, plan);
  return true;
};

/** Watches the store for a finished cell entry and works out a suggestion for that column. */
export const watch = (store: GridStore): (() => void) => store.subscribe(() => {
  if (busy) return;
  const e = store.editing;
  if (e) {
    lastEdit = { sheet: store.sheetIndex, row: e.row, col: e.col };
    if (ghost.plan) { ghost.plan = null; }
    return;
  }
  const done = lastEdit;
  lastEdit = null;
  if (ghost.plan && (ghost.plan.sheet !== store.sheetIndex || store.sel.active.col !== ghost.plan.col || [...ghost.plan.fills.keys()].some((r) => store.value(r, ghost.plan!.col) !== null))) {
    ghost.plan = null;
    busy = true; try { store.emit(); } finally { busy = false; }
    return;
  }
  if (!done || done.sheet !== store.sheetIndex || !flashAutoOn()) return;
  if (store.value(done.row, done.col) === null || store.cell(done.row, done.col)?.f) return;
  const plan = planFlashFill(store, done.row, done.col, 2);
  if (typeof plan === 'string') return;
  ghost.plan = plan;
  busy = true; try { store.emit(); } finally { busy = false; }
});
