import { useEffect, useState } from 'react';
import { Cloud, CloudOff, Copy, FileSpreadsheet, History, Link2, LogIn, LogOut, RefreshCw, Trash2 } from 'lucide-react';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import { Dialog, type DialogProps } from '../../ui/dialogs/Dialog';
import { T } from '../../ui/i18n';
import { cloud, type CloudMeta, type Version } from './api';
import { currentBinding, openCloud, removeCloud, resolveConflict, restoreVersion, saveToCloud, signIn, signOut, useCloud } from './state';

const when = (iso: string): string => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
};

const copy = async (text: string): Promise<boolean> => { try { await navigator.clipboard.writeText(text); return true; } catch { return false; } };

/** File > Clerdo cloud: sign-in, saving to the cloud, the person's cloud workbooks, versions and a share link. */
export const CloudPanel = ({ store }: { store: GridStore }) => {
  const { session } = useCloud();
  if (session.status === 'unknown') return <p className="mini">{T('Loading…')}</p>;
  if (session.status !== 'signedIn') {
    return (
      <div className="region-panel">
        <p className="mini">{T('Sign in with your Clerdo account to keep workbooks in the cloud and open them anywhere.')}</p>
        <button className="tool" onClick={signIn}><LogIn size={16} /><span>{T('Sign in')}</span></button>
      </div>
    );
  }
  const bound = currentBinding(store);
  return (
    <div className="region-panel">
      <p className="mini">{session.me.user.name}{session.me.user.email ? ` · ${session.me.user.email}` : ''}</p>
      {!bound && <button className="tool" onClick={() => void saveToCloud(store)}><Cloud size={16} /><span>{T('Save to the cloud')}</span></button>}
      {bound && <Versions store={store} id={bound.id} revision={bound.revision} />}
      {bound && <ShareLink store={store} id={bound.id} />}
      <Library store={store} current={bound?.id ?? null} />
      <button className="tool" onClick={() => void signOut()}><LogOut size={16} /><span>{T('Sign out')}</span></button>
    </div>
  );
};

const Library = ({ store, current }: { store: GridStore; current: string | null }) => {
  const [items, setItems] = useState<CloudMeta[] | null>(null);
  const [failed, setFailed] = useState(false);
  const load = () => { setFailed(false); cloud.list().then(setItems, () => setFailed(true)); };
  useEffect(load, [current]);
  return (
    <>
      <p className="mini"><b>{T('Your cloud workbooks')}</b></p>
      {failed && <button className="tool" onClick={load}><RefreshCw size={16} /><span>{T('Try again')}</span></button>}
      {!failed && items === null && <p className="mini">{T('Loading…')}</p>}
      {items?.length === 0 && <p className="mini">{T('No workbooks in the cloud yet.')}</p>}
      {items?.map((p) => (
        <div key={p.id} className="cg-cloud-row">
          <button className="tool" disabled={p.id === current} title={when(p.updatedAt)} onClick={() => void openCloud(store, p.id)}>
            <FileSpreadsheet size={16} /><span>{p.name}</span>
          </button>
          <button className="tool icon" aria-label={T('Delete')} title={T('Delete')} onClick={() => void removeCloud(store, p).then((ok) => { if (ok) load(); })}><Trash2 size={14} /></button>
        </div>
      ))}
    </>
  );
};

const Versions = ({ store, id, revision }: { store: GridStore; id: string; revision: number }) => {
  const [list, setList] = useState<{ current: Version; revisions: Version[] } | null>(null);
  useEffect(() => { setList(null); cloud.versions(id).then(setList, () => setList({ current: { revision, savedAt: '', bytes: 0 }, revisions: [] })); }, [id, revision]);
  return (
    <>
      <p className="mini"><b><History size={13} /> {T('Version history')}</b></p>
      {list === null && <p className="mini">{T('Loading…')}</p>}
      {list && <p className="mini">{T('Current version')}{list.current.savedAt ? ` · ${when(list.current.savedAt)}` : ''}</p>}
      {list?.revisions.map((v) => (
        <button key={v.revision} className="tool" onClick={() => void restoreVersion(store, v.revision)}>
          <span>{when(v.savedAt)}</span><span className="mini"> · {T('Restore')}</span>
        </button>
      ))}
    </>
  );
};

const ShareLink = ({ store, id }: { store: GridStore; id: string }) => {
  const [url, setUrl] = useState<string | null>(null);
  const create = async () => {
    try {
      const r = await cloud.share(id);
      setUrl(r.url);
      if (await copy(r.url)) store.notify(T('Link copied.'));
    } catch (e) { store.notify(e instanceof Error ? e.message : String(e)); }
  };
  const stop = async () => {
    try { await cloud.unshare(id); setUrl(null); } catch (e) { store.notify(e instanceof Error ? e.message : String(e)); }
  };
  return (
    <>
      <p className="mini"><b><Link2 size={13} /> {T('Share link')}</b></p>
      <p className="mini">{T('Anyone with the link can open a copy of this workbook.')}</p>
      {!url && <button className="tool" onClick={() => void create()}><Link2 size={16} /><span>{T('Create a link')}</span></button>}
      {url && (
        <>
          <input className="cg-cloud-link" readOnly value={url} onFocus={(e) => e.currentTarget.select()} />
          <button className="tool" onClick={() => void copy(url).then((ok) => ok && store.notify(T('Link copied.')))}><Copy size={16} /><span>{T('Copy link')}</span></button>
          <button className="tool" onClick={() => void stop()}><span>{T('Stop sharing')}</span></button>
        </>
      )}
    </>
  );
};

/** Top bar: whether the cloud workbook is saved. */
export const SaveChip = ({ store }: { store: GridStore }) => {
  const { binding } = useCloud();
  if (!binding || binding.wb !== store.wb) return null;
  const text = binding.status === 'saved' ? T('Saved to the cloud') : binding.status === 'saving' ? T('Saving…') : binding.status === 'error' ? T('Could not save. Trying again…') : T('Not saved yet');
  const Icon = binding.status === 'error' || binding.status === 'conflict' ? CloudOff : Cloud;
  if (binding.status === 'conflict') return <button className={`cg-cloud-chip ${binding.status}`} onClick={() => openDialog(store, 'cloud.conflict')}><Icon size={14} /><span>{T('Changed somewhere else')}</span></button>;
  return <span className={`cg-cloud-chip ${binding.status}`} role="status" title={text}><Icon size={14} /><span>{text}</span></span>;
};

export const ConflictDialog = ({ store, close }: DialogProps) => (
  <Dialog title={T('Changed somewhere else')} close={close} okLabel={T('Keep mine')} onOk={() => { void resolveConflict(store, 'mine'); }}
    footer={<button onClick={() => { close(); void resolveConflict(store, 'theirs'); }}>{T('Use the other version')}</button>}>
    <p>{T('This workbook was changed in another window or on another device. Which version do you want to keep?')}</p>
  </Dialog>
);
