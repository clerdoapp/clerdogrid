import { DIALOGS } from './dialogs/registry';
import { EXTRA_DIALOGS } from './dialogs/extra';
import type { GridStore } from '../app/store';
import { focusGrid } from './FormulaBar';

export const DialogHost = ({ store }: { store: GridStore }) => {
  const d = store.dialog;
  if (!d) return null;
  const close = () => { store.dialog = null; store.emit(); focusGrid(); };
  const C = DIALOGS[d.id] ?? EXTRA_DIALOGS[d.id];
  if (!C) { queueMicrotask(close); return null; }
  return <C key={d.id} store={store} data={d.data as never} close={close} />;
};
