import type { GridStore } from '../../app/store';
import { initOnline, onlineAllowed, onOnlineSettled, setOnlineAllowed } from '../../engine/online';
import { setImageRedraw } from '../../ui/cellImages';
import { T } from '../../ui/i18n';

let restricted = false;
const bound = new WeakSet<GridStore>();

/**
 * While internet access is off, a second content policy blocks every remote request, so the setting holds
 * even if some code tried to fetch. A policy can't be lifted once added, so turning access on later needs a reload.
 */
export const startOnline = (): void => {
  initOnline();
  if (onlineAllowed() || typeof document === 'undefined' || !import.meta.env.PROD) return;
  const meta = document.createElement('meta');
  meta.httpEquiv = 'Content-Security-Policy';
  meta.content = "connect-src 'self' blob: data:; img-src 'self' data: blob:";
  document.head.appendChild(meta);
  restricted = true;
};

export const needsReload = (): boolean => restricted;

export const bindStore = (store: GridStore): void => {
  if (bound.has(store)) return;
  bound.add(store);
  onOnlineSettled(() => { store.engine.recalc(); store.touch(); });
  setImageRedraw(() => store.touch());
};

export const setAllowed = (store: GridStore, on: boolean): void => {
  setOnlineAllowed(on, true);
  store.engine.recalc(true);
  store.touch();
  if (on && restricted) store.notify(T('Reload the app to finish turning on internet access for functions.'));
};
