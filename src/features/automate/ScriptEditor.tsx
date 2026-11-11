import { useEffect, useRef, useState } from 'react';
import { BookMarked, BookOpen, Copy, FilePlus2, FileText, Library, Pencil, Play, Save, Trash2, X } from 'lucide-react';
import type { DialogProps } from '../../ui/dialogs/Dialog';
import type { GridStore } from '../../app/store';
import { CodeEditor } from './CodeEditor';
import { bookScripts, loadLibrary, newScriptId, saveLibrary, setBookScripts, uniqueScriptName, type Script } from './library';
import { runScript } from './runner';
import { NEW_SCRIPT, SAMPLES } from './samples';
import { API_HELP } from './apiHelp';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;
type Where = 'book' | 'lib' | 'sample';
interface Ref { where: Where; id: string }
export interface EditorData { open?: Ref; tab?: Where; run?: boolean }

interface Output { ok: boolean; message: string; logs: string[]; line?: number; running?: boolean }

export const ScriptEditor = ({ store, data, close }: DialogProps<EditorData>) => {
  const [tab, setTab] = useState<Where>(data?.tab ?? data?.open?.where ?? (bookScripts(store).length ? 'book' : 'sample'));
  const [lib, setLib] = useState<Script[]>(loadLibrary);
  const book = bookScripts(store);
  const [cur, setCur] = useState<Ref | null>(data?.open ?? (book[0] ? { where: 'book', id: book[0].id } : null));
  const find = (r: Ref | null): Script | undefined => !r ? undefined : r.where === 'book' ? book.find((s) => s.id === r.id) : r.where === 'lib' ? lib.find((s) => s.id === r.id) : SAMPLES.find((s) => s.id === r.id);
  const current = find(cur);
  const [draft, setDraft] = useState(current?.code ?? '');
  const [renaming, setRenaming] = useState<string | null>(null);
  const [out, setOut] = useState<Output | null>(null);
  const [help, setHelp] = useState(false);
  const dirty = !!current && cur?.where !== 'sample' && draft !== current.code;
  const root = useRef<HTMLDivElement>(null);

  useEffect(() => { setDraft(current?.code ?? ''); setRenaming(null); setOut(null); }, [cur?.where, cur?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const updateBook = (list: Script[]) => setBookScripts(store, list);
  const updateLib = (list: Script[]) => { setLib(list); if (!saveLibrary(list)) store.notify(T('Your library couldn\'t be saved in this browser.')); };
  const listOf = (w: Where): Script[] => (w === 'book' ? book : w === 'lib' ? lib : SAMPLES);
  const setList = (w: Where, l: Script[]) => (w === 'book' ? updateBook(l) : w === 'lib' ? updateLib(l) : undefined);

  const save = (): Script | undefined => {
    if (!cur || !current || cur.where === 'sample') return current;
    const next = { ...current, code: draft };
    setList(cur.where, listOf(cur.where).map((s) => (s.id === cur.id ? next : s)));
    return next;
  };

  const create = (w: 'book' | 'lib', name: string, code: string) => {
    const s: Script = { id: newScriptId(), name: uniqueScriptName(listOf(w).map((x) => x.name), name), code };
    setList(w, [...listOf(w), s]);
    setTab(w);
    setCur({ where: w, id: s.id });
    return s;
  };

  const run = async () => {
    if (out?.running) return;
    const code = draft;
    if (dirty) save();
    setOut({ ok: true, message: N('Running…'), logs: [], running: true });
    const r = await runScript(store, code, current?.name ?? 'script');
    setOut({ ok: r.ok, message: r.message, logs: r.logs, line: r.line });
  };

  const remove = () => {
    if (!cur || cur.where === 'sample' || !current) return;
    if (!window.confirm(T('Delete "{name}"?', { name: current.name }))) return;
    const l = listOf(cur.where).filter((s) => s.id !== cur.id);
    setList(cur.where, l);
    setCur(l[0] ? { where: cur.where, id: l[0].id } : null);
  };

  const rename = (name: string) => {
    if (!cur || cur.where === 'sample' || !current) return;
    const n = name.trim().slice(0, 200);
    setRenaming(null);
    if (!n || n === current.name) return;
    setList(cur.where, listOf(cur.where).map((s) => (s.id === cur.id ? { ...s, name: uniqueScriptName(listOf(cur.where).filter((x) => x.id !== cur.id).map((x) => x.name), n) } : s)));
  };

  const tryClose = () => {
    if (dirty && !window.confirm(T('Close without saving your changes to this script?'))) return;
    close();
  };

  useEffect(() => { if (data?.run && current) void run(); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const items = listOf(tab);
  const errLine = out && !out.ok ? out.line : undefined;
  const width = Math.min(1080, typeof window !== 'undefined' ? window.innerWidth - 24 : 1080);

  return (
    <div className="backdrop" onPointerDown={(e) => { if (e.target === e.currentTarget) tryClose(); }}>
      <div className="dialog au-dialog" ref={root} role="dialog" aria-modal="true" aria-label={T('Scripts')} style={{ width }}
        onKeyDown={(e) => { if (e.key === 'Escape') { e.stopPropagation(); tryClose(); } }}>
        <div className="dhead"><h2>{T('Scripts')}</h2><button className="ico" aria-label={T('Close')} onClick={tryClose}><X size={16} strokeWidth={S} /></button></div>
        <div className="au-body">
          <aside className="au-side">
            <div className="b-seg au-tabs" role="tablist">
              {([['book', N('Workbook')], ['lib', N('My library')], ['sample', N('Samples')]] as const).map(([w, l]) => (
                <button key={w} role="tab" aria-selected={tab === w} className={tab === w ? 'on' : ''} onClick={() => setTab(w)}>{Tr(l)}</button>
              ))}
            </div>
            <p className="au-hint">{tab === 'book' ? T('Saved in this file.') : tab === 'lib' ? T('Saved in this browser, for any workbook.') : T('Ready-made scripts to try or copy.')}</p>
            {tab === 'sample' ? (
              <div className="au-gallery">
                {SAMPLES.map((s) => (
                  <button key={s.id} className={`au-card${cur?.where === 'sample' && cur.id === s.id ? ' on' : ''}`} onClick={() => setCur({ where: 'sample', id: s.id })}>
                    <b>{Tr(s.name)}</b><span>{Tr(s.about)}</span>
                  </button>
                ))}
              </div>
            ) : (
              <div className="au-list" role="listbox" aria-label={tab === 'book' ? T('Workbook scripts') : T('Library scripts')}>
                {items.map((s) => (
                  <button key={s.id} role="option" aria-selected={cur?.where === tab && cur.id === s.id} className={cur?.where === tab && cur.id === s.id ? 'on' : ''} onClick={() => setCur({ where: tab, id: s.id })}>
                    <FileText size={15} strokeWidth={S} /><span>{s.name}</span>
                  </button>
                ))}
                {!items.length && <p className="au-empty">{T('No scripts yet.')}</p>}
              </div>
            )}
            {tab !== 'sample' && <button className="tool au-new" onClick={() => create(tab, T('New script'), NEW_SCRIPT)}><FilePlus2 size={16} strokeWidth={S} /><span>{T('New script')}</span></button>}
          </aside>

          <section className="au-main">
            {current ? (<>
              <div className="au-bar">
                {renaming !== null ? (
                  <input className="au-name" autoFocus value={renaming} aria-label={T('Script name')} maxLength={200}
                    onChange={(e) => setRenaming(e.target.value)} onBlur={() => rename(renaming)}
                    onKeyDown={(e) => { if (e.key === 'Enter') rename(renaming); if (e.key === 'Escape') { e.stopPropagation(); setRenaming(null); } }} />
                ) : <h3 className="au-title" title={Tr(current.name)}>{Tr(current.name)}{dirty && <i className="au-dot" aria-label={T('Not saved')} />}</h3>}
                <div className="grow" />
                <button className="primary" onClick={() => void run()} disabled={out?.running} title={T('Run (Ctrl+Enter)')}><Play size={15} strokeWidth={S} />{T('Run')}</button>
                {cur?.where !== 'sample' && <button className="tool icon" aria-label={T('Save')} title={T('Save (Ctrl+S)')} disabled={!dirty} onClick={save}><Save size={16} strokeWidth={S} /></button>}
                {cur?.where !== 'sample' && <button className="tool icon" aria-label={T('Rename')} title={T('Rename')} onClick={() => setRenaming(current.name)}><Pencil size={16} strokeWidth={S} /></button>}
                <button className="tool icon" aria-label={cur?.where === 'sample' ? T('Copy to workbook') : T('Duplicate')} title={cur?.where === 'sample' ? T('Copy to workbook') : T('Duplicate')}
                  onClick={() => create(cur?.where === 'lib' ? 'lib' : 'book', cur?.where === 'sample' ? Tr(current.name) : T('{name} copy', { name: current.name }), draft)}><Copy size={16} strokeWidth={S} /></button>
                {cur?.where === 'book' && <button className="tool icon" aria-label={T('Save to my library')} title={T('Save to my library')} onClick={() => { create('lib', current.name, draft); store.notify(T('Saved to your library.')); }}><Library size={16} strokeWidth={S} /></button>}
                {cur?.where === 'lib' && <button className="tool icon" aria-label={T('Add to workbook')} title={T('Add to workbook')} onClick={() => create('book', current.name, draft)}><BookMarked size={16} strokeWidth={S} /></button>}
                {cur?.where !== 'sample' && <button className="tool icon" aria-label={T('Delete')} title={T('Delete')} onClick={remove}><Trash2 size={16} strokeWidth={S} /></button>}
                <button className={`tool icon${help ? ' on' : ''}`} aria-label={T('Script reference')} title={T('Script reference')} aria-pressed={help} onClick={() => setHelp(!help)}><BookOpen size={16} strokeWidth={S} /></button>
              </div>
              <div className="au-work">
                <CodeEditor value={draft} onChange={setDraft} errorLine={errLine} onRun={() => void run()} onSave={save} readOnly={cur?.where === 'sample'} />
                {help && <div className="au-help" aria-label={T('Script reference')}>{API_HELP.map((g) => (
                  <div key={g.title}><h4>{Tr(g.title)}</h4>{g.items.map(([sig, about]) => <p key={sig}><code>{sig}</code><span>{Tr(about)}</span></p>)}</div>
                ))}</div>}
              </div>
              <div className={`au-out${out && !out.ok ? ' bad' : ''}`} aria-live="polite">
                <div className="au-outhead"><b>{T('Output')}</b>{out && <span>{Tr(out.message)}</span>}</div>
                <pre>{out?.logs.length ? out.logs.join('\n') : out ? '' : T('Run the script to see what it prints with console.log.')}</pre>
              </div>
            </>) : (
              <div className="au-blank">
                <p>{tab === 'sample' ? T('Pick a sample to see its code.') : T('Make a new script, or start from a sample.')}</p>
                <div className="b-inline">
                  <button className="primary" onClick={() => create('book', T('New script'), NEW_SCRIPT)}><FilePlus2 size={15} strokeWidth={S} />{T('New script')}</button>
                  <button onClick={() => setTab('sample')}>{T('See samples')}</button>
                </div>
              </div>
            )}
          </section>
        </div>
        <div className="dfoot"><span className="muted au-foot">{T('Scripts run in a sandbox: no internet, no files, and they stop after 10 seconds. You can undo a run.')}</span><div className="grow" /><button className="primary" onClick={tryClose}>{T('Close')}</button></div>
      </div>
    </div>
  );
};

/** Runs a workbook script by id from the panel. */
export const runSavedScript = async (store: GridStore, id: string): Promise<void> => {
  const s = bookScripts(store).find((x) => x.id === id);
  if (!s) return;
  store.notify(T('Running "{name}"…', { name: s.name }));
  const r = await runScript(store, s.code, s.name);
  const log = r.logs.length ? ` Output: ${r.logs.slice(-1)[0]}` : '';
  store.notify(`${s.name}: ${r.message}${log}`.slice(0, 300));
};
