import { useEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import { GridStore } from '../app/store';
import { registerCommands } from '../app/commands';
import { registerFeatures } from '../features';
import { registerAllFunctions } from '../engine/functions';
import { openLaunchFile, restoreDraft, startAutosave, openBytes, desktop } from '../app/files';
import { TopBar } from './TopBar';
import { Rail, ToolPanel } from './Rail';
import { FormulaBar } from './FormulaBar';
import { GridView, type GridApi } from './GridView';
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import { ContextMenu } from './ContextMenu';
import { DialogHost } from './DialogHost';
import { runStartups } from './panelRegistry';
import { onLangChange } from './i18n';

const coarse = typeof matchMedia !== 'undefined' && matchMedia('(pointer: coarse)').matches;

export const App = () => {
  const store = useMemo(() => {
    registerAllFunctions();
    registerCommands();
    registerFeatures();
    const s = new GridStore();
    if (coarse || window.innerWidth < 900) s.view.panel = null;
    try { const t = localStorage.getItem('clerdogrid.theme'); if (t === 'light' || t === 'dark') s.view.theme = t; } catch { /* private mode */ }
    return s;
  }, []);
  useSyncExternalStore(store.subscribe, store.getVersion);
  // A new interface language re-renders everything (canvas included) without a reload.
  useEffect(() => onLangChange(() => store.touch()), [store]);
  const api = useRef<GridApi | null>(null);

  useEffect(() => {
    (window as unknown as { __store: GridStore }).__store = store;
    void runStartups(store).then((handled) => { if (!handled) void openLaunchFile(store).then((opened) => { if (!opened) void restoreDraft(store); }); });
    const stop = startAutosave(store);
    const beforeUnload = (e: BeforeUnloadEvent) => { if (desktop && store.file.dirty) { e.preventDefault(); e.returnValue = ''; } };
    window.addEventListener('beforeunload', beforeUnload);
    // Files dropped on the window open as workbooks.
    const onDrop = (e: DragEvent) => {
      const f = e.dataTransfer?.files?.[0];
      if (!f) return;
      e.preventDefault();
      void f.arrayBuffer().then((b) => openBytes(store, f.name, new Uint8Array(b)));
    };
    const onDragOver = (e: DragEvent) => { if (e.dataTransfer?.types.includes('Files')) e.preventDefault(); };
    window.addEventListener('drop', onDrop);
    window.addEventListener('dragover', onDragOver);
    (document.querySelector('.gridwrap') as HTMLElement | null)?.focus();
    return () => { stop(); window.removeEventListener('beforeunload', beforeUnload); window.removeEventListener('drop', onDrop); window.removeEventListener('dragover', onDragOver); };
  }, [store]);

  useEffect(() => {
    const t = store.view.theme;
    if (t === 'system') document.documentElement.removeAttribute('data-theme');
    else document.documentElement.setAttribute('data-theme', t);
  }, [store.view.theme]);

  useEffect(() => {
    document.title = `${store.file.dirty ? '• ' : ''}${store.file.name} - ClerdoGrid`;
  }, [store.file.name, store.file.dirty]);

  return (
    <div className="app">
      <TopBar store={store} />
      <div className="work">
        <Rail store={store} />
        <ToolPanel store={store} />
        <main className="stage">
          <FormulaBar store={store} />
          <GridView store={store} apiRef={api} />
          <SheetTabs store={store} />
        </main>
      </div>
      <StatusBar store={store} />
      {store.menu && <ContextMenu store={store} />}
      <DialogHost store={store} />
      {store.toast && <div className="toast" role="status">{store.toast}</div>}
    </div>
  );
};
