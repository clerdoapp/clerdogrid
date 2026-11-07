import { FolderOpen, Redo2, Save, Search, Undo2 } from 'lucide-react';
import { run } from '../app/commands';
import type { GridStore } from '../app/store';
import { prettyKeys } from './keys';

const S = 1.9;

export const TopBar = ({ store }: { store: GridStore }) => {
  const f = store.file;
  const state = f.dirty ? 'Not saved yet' : f.lastSaved ? 'Saved' : '';
  return (
    <header className="topbar">
      <span className="wordmark">Clerdo<b>Grid</b></span>
      <input className="docname" aria-label="Workbook name" value={f.name} spellCheck={false}
        onChange={(e) => { store.file = { ...f, name: e.target.value.replace(/[\\/:*?"<>|]/g, '') }; store.emit(); }} />
      <span className="savestate" role="status">{state}</span>
      <div className="grow" />
      <button className="cmdsearch" onClick={() => run(store, 'edit.palette')}>
        <Search size={15} strokeWidth={S} /> <span>Search commands and functions</span> <kbd>{prettyKeys('Mod+K')}</kbd>
      </button>
      <div className="grow" />
      <button className="ico" aria-label="Undo" title={store.undoLabel ? `Undo ${store.undoLabel}` : 'Undo'} disabled={!store.canUndo} onClick={() => store.undo()}><Undo2 size={17} strokeWidth={S} /></button>
      <button className="ico" aria-label="Redo" title="Redo" disabled={!store.canRedo} onClick={() => store.redo()}><Redo2 size={17} strokeWidth={S} /></button>
      <button className="ico" aria-label="Open" title="Open" onClick={() => run(store, 'file.open')}><FolderOpen size={17} strokeWidth={S} /></button>
      <button className="primary" onClick={() => run(store, 'file.save')}><Save size={15} strokeWidth={S} /> Save</button>
    </header>
  );
};
