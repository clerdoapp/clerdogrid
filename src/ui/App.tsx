import { useEffect, useMemo, useRef, useSyncExternalStore } from 'react';
import { Button, ClerdoFrame, setFrameTranslator, useTheme } from '@clerdo/frame';
import { FolderOpen, Redo2, Save, Undo2 } from 'lucide-react';
import { GridStore } from '../app/store';
import { registerCommands, run } from '../app/commands';
import { registerFeatures } from '../features';
import { registerAllFunctions } from '../engine/functions';
import { openLaunchFile, restoreDraft, startAutosave, openBytes, desktop } from '../app/files';
import { Rail, ToolPanel } from './Rail';
import { FormulaBar } from './FormulaBar';
import { GridView, type GridApi } from './GridView';
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import { ContextMenu } from './ContextMenu';
import { DialogHost } from './DialogHost';
import { hasAccountSource, runStartups, topBarItems, useAccountSource } from './panelRegistry';
import { onLangChange, T } from './i18n';
import { en } from './i18n/en';

const coarse = typeof matchMedia !== 'undefined' && matchMedia('(pointer: coarse)').matches;
const S = 1.9;

// The shared Clerdo frame speaks Grid's interface language.
setFrameTranslator((text, vars) => (text in en ? T(text as keyof typeof en, vars) : text.replace(/\{(\w+)\}/g, (m, k: string) => (vars && k in vars ? String(vars[k]) : m))));

export const App = () => {
  const store = useMemo(() => {
    registerAllFunctions();
    registerCommands();
    registerFeatures();
    const s = new GridStore();
    if (coarse || window.innerWidth < 900) s.view.panel = null;
    return s;
  }, []);
  useSyncExternalStore(store.subscribe, store.getVersion);
  useEffect(() => onLangChange(() => store.touch()), [store]);
  const api = useRef<GridApi | null>(null);
  const account = useAccountSource();

  // One light/dark setting for every Clerdo app: Grid's canvas follows the frame's choice.
  const { choice, setChoice } = useTheme();
  useEffect(() => {
    try {
      const old = localStorage.getItem('clerdogrid.theme');
      if (old && !localStorage.getItem('clerdo.theme') && (old === 'light' || old === 'dark')) setChoice(old);
    } catch { /* private mode */ }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useEffect(() => { if (store.view.theme !== choice) { store.view.theme = choice; store.touch(); } }, [choice, store]);
  useEffect(() => { if (store.view.theme !== choice) setChoice(store.view.theme); }, [store.view.theme]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    (window as unknown as { __store: GridStore }).__store = store;
    void runStartups(store).then((handled) => { if (!handled) void openLaunchFile(store).then((opened) => { if (!opened) void restoreDraft(store); }); });
    const stop = startAutosave(store);
    const beforeUnload = (e: BeforeUnloadEvent) => { if (desktop && store.file.dirty) { e.preventDefault(); e.returnValue = ''; } };
    window.addEventListener('beforeunload', beforeUnload);
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
    document.title = `${store.file.dirty ? '• ' : ''}${store.file.name} - Grid`;
  }, [store.file.name, store.file.dirty]);

  const f = store.file;
  // Online, the save chip from the cloud feature tells where the workbook is; offline, the local state shows here.
  const status = hasAccountSource() ? undefined : f.dirty ? T('Not saved yet') : f.lastSaved ? T('Saved') : undefined;
  return (
    <ClerdoFrame
      app="grid"
      user={account.user}
      onSignOut={account.signOut}
      onSignIn={hasAccountSource() ? account.signIn : undefined}
      homeHref={hasAccountSource() ? '/' : '#'}
      doc={{
        title: f.name,
        status,
        saved: !f.dirty,
        titleLabel: T('Workbook name'),
        onTitleChange: (name) => { store.file = { ...store.file, name: name.replace(/[\\/:*?"<>|]/g, '') }; store.emit(); },
      }}
      searchPlaceholder={T('Search commands and functions')}
      onSearch={() => run(store, 'edit.palette')}
      actions={
        <>
          {topBarItems().map((t) => <span key={t.id} className="topitem">{t.render(store)}</span>)}
          <button className="ico" aria-label={T('Undo')} title={T('Undo')} disabled={!store.canUndo} onClick={() => store.undo()}><Undo2 size={17} strokeWidth={S} /></button>
          <button className="ico" aria-label={T('Redo')} title={T('Redo')} disabled={!store.canRedo} onClick={() => store.redo()}><Redo2 size={17} strokeWidth={S} /></button>
          <button className="ico" aria-label={T('Open')} title={T('Open')} onClick={() => run(store, 'file.open')}><FolderOpen size={17} strokeWidth={S} /></button>
        </>
      }
      primary={<Button variant="primary" onClick={() => run(store, 'file.save')}><Save size={15} strokeWidth={S} /> {T('Save')}</Button>}
      rail={<Rail store={store} />}
      panel={<ToolPanel store={store} />}
    >
      <div className="app">
        <main className="stage">
          <FormulaBar store={store} />
          <GridView store={store} apiRef={api} />
          <SheetTabs store={store} />
        </main>
        <StatusBar store={store} />
        {store.menu && <ContextMenu store={store} />}
        <DialogHost store={store} />
        {store.toast && <div className="toast" role="status">{store.toast}</div>}
      </div>
    </ClerdoFrame>
  );
};
