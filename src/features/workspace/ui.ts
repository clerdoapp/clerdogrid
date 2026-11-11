// UI state for the workspace feature (not saved with the workbook, not undoable).
import type { GridStore } from '../../app/store';

export const ui = {
  focusCell: false,
  navPane: false,
  navSearch: '',
  navHidden: false,
  versionsPane: false,
  fullScreen: false,
  /** Selected cell's error-check menu is open. */
  errorMenu: false,
  /** Status bar choices menu. */
  statusMenu: null as { x: number; y: number } | null,
};

export const setUi = (store: GridStore, patch: Partial<typeof ui>): void => { Object.assign(ui, patch); store.emit(); };
