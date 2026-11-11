import { FolderOpen, Redo2, Save, Search, Undo2 } from 'lucide-react';
import { run } from '../app/commands';
import type { GridStore } from '../app/store';
import { prettyKeys } from './keys';
import { topBarItems } from './panelRegistry';
import { T, Tr } from './i18n';

const S = 1.9;

export const TopBar = ({ store }: { store: GridStore }) => {
  const f = store.file;
  const state = f.dirty ? T('Not saved yet') : f.lastSaved ? T('Saved') : '';
  return (
    <header className="topbar">
      <span className="wordmark">Clerdo<b>Grid</b></span>
      <input className="docname" aria-label={T('Workbook name')} value={f.name} spellCheck={false}
        onChange={(e) => { store.file = { ...f, name: e.target.value.replace(/[\\/:*?"<>|]/g, '') }; store.emit(); }} />
      <span className="savestate" role="status">{state}</span>
      <div className="grow" />
      <button className="cmdsearch" onClick={() => run(store, 'edit.palette')}>
        <Search size={15} strokeWidth={S} /> <span>{T('Search commands and functions')}</span> <kbd>{prettyKeys('Mod+K')}</kbd>
      </button>
      <div className="grow" />
      {topBarItems().map((t) => <span key={t.id} className="topitem">{t.render(store)}</span>)}
      <button className="ico" aria-label={T('Undo')} title={store.undoLabel ? T('Undo {undoLabel}', { undoLabel: Tr(store.undoLabel) }) : T('Undo')} disabled={!store.canUndo} onClick={() => store.undo()}><Undo2 size={17} strokeWidth={S} /></button>
      <button className="ico" aria-label={T('Redo')} title={T('Redo')} disabled={!store.canRedo} onClick={() => store.redo()}><Redo2 size={17} strokeWidth={S} /></button>
      <button className="ico" aria-label={T('Open')} title={T('Open')} onClick={() => run(store, 'file.open')}><FolderOpen size={17} strokeWidth={S} /></button>
      <button className="primary" onClick={() => run(store, 'file.save')}><Save size={15} strokeWidth={S} />{' '}{T('Save')}</button>
    </header>
  );
};
