import { useEffect, useRef, useState } from 'react';
import { Check, MoreHorizontal, Pencil, RotateCcw, Trash2, X } from 'lucide-react';
import { addr, cellKey, keyCol, keyRow, type CommentThread } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import type { Viewport } from '../../ui/render';
import { cellBox, fmtTime, setUi, ui } from './state';
import { addComment, deleteComment, editComment, replyComment, resolveComment, sortedKeys, threadAt } from './ops';

const S = 1.9;
const focusGrid = () => (document.querySelector('.gridwrap') as HTMLElement | null)?.focus({ preventScroll: true });

const initials = (name: string): string => name.split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]!.toUpperCase()).join('') || '?';

const Composer = ({ placeholder, initial = '', onSend, onCancel, label = 'Post', autoFocus }: { placeholder: string; initial?: string; onSend: (t: string) => void; onCancel?: () => void; label?: string; autoFocus?: boolean }) => {
  const [text, setText] = useState(initial);
  const ref = useRef<HTMLTextAreaElement>(null);
  useEffect(() => { if (autoFocus) ref.current?.focus({ preventScroll: true }); }, [autoFocus]);
  const send = () => { if (text.trim()) { onSend(text); setText(''); } };
  return (
    <div className="cmt-compose">
      <textarea ref={ref} rows={2} value={text} placeholder={placeholder} aria-label={placeholder} maxLength={32767} onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); send(); } if (e.key === 'Escape') { e.preventDefault(); onCancel?.(); } }} />
      <div className="cmt-actions">
        {onCancel && <button onClick={onCancel}>Cancel</button>}
        <button className="primary sm" disabled={!text.trim()} onClick={send}>{label}</button>
      </div>
    </div>
  );
};

const Message = ({ store, row, col, index, author, text, time, canEdit }: { store: GridStore; row: number; col: number; index: number; author: string; text: string; time: string; canEdit: boolean }) => {
  const [editing, setEditing] = useState(false);
  const [menu, setMenu] = useState(false);
  return (
    <div className="cmt-msg">
      <div className="cmt-head">
        <span className="cmt-avatar" aria-hidden="true">{initials(author)}</span>
        <div className="cmt-who"><b>{author}</b><span className="muted">{fmtTime(time)}</span></div>
        {canEdit && <button className="ico sm" aria-label="More" title="More" onClick={() => setMenu(!menu)}><MoreHorizontal size={15} strokeWidth={S} /></button>}
      </div>
      {menu && <div className="cmt-menu">
        <button onClick={() => { setMenu(false); setEditing(true); }}><Pencil size={14} strokeWidth={S} />Edit</button>
        <button onClick={() => { setMenu(false); deleteComment(store, row, col, index); }}><Trash2 size={14} strokeWidth={S} />{index < 0 ? 'Delete thread' : 'Delete reply'}</button>
      </div>}
      {editing ? <Composer placeholder="Edit comment" initial={text} label="Save" autoFocus onCancel={() => setEditing(false)} onSend={(t) => { editComment(store, row, col, index, t); setEditing(false); }} />
        : <div className="cmt-text">{text}</div>}
    </div>
  );
};

export const Thread = ({ store, row, col, thread, compact }: { store: GridStore; row: number; col: number; thread: CommentThread; compact?: boolean }) => (
  <div className={`cmt-thread${thread.resolved ? ' resolved' : ''}`}>
    <div className="cmt-bar">
      <span className="cmt-ref">{addr(row, col)}</span>
      {thread.resolved && <span className="cmt-tag">Resolved</span>}
      <div className="grow" />
      {thread.resolved
        ? <button className="ico sm" title="Reopen" aria-label="Reopen" onClick={() => resolveComment(store, row, col, false)}><RotateCcw size={15} strokeWidth={S} /></button>
        : <button className="ico sm" title="Resolve" aria-label="Resolve" onClick={() => resolveComment(store, row, col, true)}><Check size={15} strokeWidth={S} /></button>}
    </div>
    <Message store={store} row={row} col={col} index={-1} author={thread.author} text={thread.text} time={thread.time} canEdit />
    {thread.replies.map((r, i) => <Message key={i} store={store} row={row} col={col} index={i} author={r.author} text={r.text} time={r.time} canEdit />)}
    {!compact && <Composer placeholder={thread.resolved ? 'Reply to reopen' : 'Reply'} label="Reply" onSend={(t) => replyComment(store, row, col, t)} />}
  </div>
);

/** Note and comment popovers next to cells, plus all notes when "Show all notes" is on. */
export const PopoverLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const si = store.sheetIndex;
  const pin = ui.pinned?.sheet === si ? ui.pinned : null;
  const hov = !pin && ui.hover?.sheet === si ? ui.hover : null;
  const target = pin ?? hov;
  const out: React.ReactNode[] = [];
  const pos = (row: number, col: number, w: number) => {
    const b = cellBox(vp, row, col, store);
    let left = b.left + b.width + 6;
    if (left + w > vp.width - 6) left = Math.max(vp.headerW + 4, b.left - w - 6);
    return { left, top: Math.max(vp.headerH + 4, Math.min(b.top, vp.height - 160)) };
  };
  if (ui.showAllNotes) for (const k of store.sheet.notes.keys()) {
    const row = keyRow(k), col = keyCol(k);
    if (target && target.row === row && target.col === col) continue;
    const n = store.sheet.notes.get(k)!;
    out.push(<div key={`n${k}`} className="notepop all" style={{ ...pos(row, col, 200), pointerEvents: 'auto' }} onPointerDown={(e) => e.stopPropagation()} onDoubleClick={() => { store.select({ row, col }); openDialog(store, 'note'); }}>{n.author && <b>{n.author}</b>}<div>{n.text}</div></div>);
  }
  if (target) {
    const { row, col } = target;
    const k = cellKey(row, col);
    const note = store.sheet.notes.get(k);
    const thread = threadAt(store, row, col);
    const compose = pin?.compose && !thread;
    if (note) out.push(
      <div key="note" className="notepop" style={{ ...pos(row, col, 220), pointerEvents: 'auto' }} onPointerDown={(e) => e.stopPropagation()}>
        {note.author && <b>{note.author}</b>}<div>{note.text}</div>
        {pin && <div className="cmt-actions"><button onClick={() => { setUi(store, { pinned: null }); store.select({ row, col }); openDialog(store, 'note'); }}>Edit note</button></div>}
      </div>,
    );
    if (thread || compose) {
      const p = pos(row, col, 300);
      out.push(
        <div key="cmt" className="cmtpop" role="dialog" aria-label={`Comment on ${addr(row, col)}`} style={{ left: p.left, top: note ? p.top + 90 : p.top, pointerEvents: 'auto' }}
          onPointerDown={(e) => { e.stopPropagation(); if (!pin) setUi(store, { pinned: { sheet: si, row, col } }); }}
          onKeyDown={(e) => { if (e.key === 'Escape') { setUi(store, { pinned: null }); focusGrid(); } }}>
          {pin && <button className="ico sm cmt-close" aria-label="Close" onClick={() => { setUi(store, { pinned: null }); focusGrid(); }}><X size={15} strokeWidth={S} /></button>}
          {thread ? <Thread store={store} row={row} col={col} thread={thread} compact={!pin} />
            : <div className="cmt-thread"><div className="cmt-bar"><span className="cmt-ref">{addr(row, col)}</span></div>
              <Composer placeholder="Write a comment" autoFocus onCancel={() => { setUi(store, { pinned: null }); focusGrid(); }}
                onSend={(t) => { addComment(store, row, col, t); setUi(store, { pinned: { sheet: si, row, col } }); }} /></div>}
        </div>,
      );
    }
  }
  return out.length ? <>{out}</> : null;
};

/** Side pane listing every thread on the sheet. */
export const CommentsPane = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  if (!ui.commentsPane) return null;
  const map = store.sheet.comments ?? new Map<number, CommentThread>();
  const keys = sortedKeys(map.keys()).filter((k) => ui.showResolved || !map.get(k)!.resolved);
  const go = (row: number, col: number) => { store.select({ row, col }); setUi(store, { pinned: null }); };
  return (
    <aside className="cmtpane" aria-label="Comments" style={{ left: Math.max(vp.headerW, vp.width - 316), top: vp.headerH, height: Math.max(0, vp.height - vp.headerH), pointerEvents: 'auto' }} onPointerDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
      <div className="phead"><h2>Comments</h2><button className="ico" aria-label="Close comments" onClick={() => { setUi(store, { commentsPane: false }); focusGrid(); }}><X size={16} strokeWidth={S} /></button></div>
      <div className="cmtpane-tools">
        <button className="primary sm" onClick={() => { const { row, col } = store.sel.active; setUi(store, { pinned: { sheet: store.sheetIndex, row, col, compose: true } }); }}>New comment</button>
        <label className="b-check"><input type="checkbox" checked={ui.showResolved} onChange={(e) => setUi(store, { showResolved: e.target.checked })} />Show resolved</label>
      </div>
      <div className="cmtpane-list">
        {!keys.length && <p className="muted small pad">No comments on this sheet. Select a cell and choose New comment.</p>}
        {keys.map((k) => {
          const row = keyRow(k), col = keyCol(k);
          const on = store.sel.active.row === row && store.sel.active.col === col;
          return (
            <div key={k} className={`cmtpane-item${on ? ' on' : ''}`} onClick={(e) => { if (!(e.target as Element).closest('button,textarea')) go(row, col); }}>
              <Thread store={store} row={row} col={col} thread={map.get(k)!} />
            </div>
          );
        })}
      </div>
    </aside>
  );
};
