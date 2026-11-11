// Side panes over the grid: navigation (everything in the workbook, by sheet) and version history.
import { useEffect, useMemo, useState } from 'react';
import { BarChart3, ChevronDown, ChevronRight, Eye, FileSpreadsheet, History, Image, MessageSquare, Pencil, Search, Shapes, StickyNote, Table2, Tag, Trash2, X, Grid3x3, Save } from 'lucide-react';
import { keyCol, keyRow, parseRange, rangeAddr, type CellRange } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import type { Viewport } from '../../ui/render';
import { focusGrid } from '../../ui/FormulaBar';
import { setUi as setObjUi } from '../objects/state';
import { setUi, ui } from './ui';
import { openPreview, preview, refreshVersions, snapshotNow, versionList } from './history';
import { deleteVersion, formatSize, formatWhen, renameVersion } from './versions';
import { settings, saveSettings } from './settings';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;

// ---------- Navigation ----------

interface NavItem { key: string; kind: 'table' | 'chart' | 'picture' | 'shape' | 'comment' | 'note' | 'name' | 'pivot'; label: string; detail?: string; hidden?: boolean; go: () => void }
interface NavSheet { index: number; name: string; state: string; items: NavItem[] }

const ICONS = { table: Table2, chart: BarChart3, picture: Image, shape: Shapes, comment: MessageSquare, note: StickyNote, name: Tag, pivot: Grid3x3 } as const;
const KIND_LABEL: Record<NavItem['kind'], string> = { table: N('Tables'), chart: N('Charts'), picture: N('Pictures'), shape: N('Shapes'), comment: N('Comments'), note: N('Notes'), name: N('Named ranges'), pivot: N('PivotTables') };

const nameTarget = (store: GridStore, formula: string): { sheet: number; range: CellRange } | null => {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$/.exec(formula.trim());
  if (!m) return null;
  const si = store.engine.sheetIndex((m[1] ?? m[2] ?? '').replace(/''/g, "'"));
  const r = parseRange(m[3]!.replace(/\$/g, ''));
  return si >= 0 && r ? { sheet: si, range: r } : null;
};

const collect = (store: GridStore): NavSheet[] => {
  const go = (sheet: number, range: CellRange) => () => {
    if (store.wb.sheets[sheet]!.state !== 'visible') { store.notify(T('That sheet is hidden. Unhide it to go there.')); return; }
    if (sheet !== store.sheetIndex) store.setActiveSheet(sheet);
    store.selectRange(range);
    focusGrid();
  };
  const cell = (row: number, col: number): CellRange => ({ r1: row, c1: col, r2: row, c2: col });
  return store.wb.sheets.map((s, si) => {
    const items: NavItem[] = [];
    for (const t of s.tables) items.push({ key: `t:${t.name}`, kind: 'table', label: t.name, detail: rangeAddr(t.range), go: go(si, t.range) });
    for (const p of s.pivots ?? []) items.push({ key: `p:${p.name}`, kind: 'pivot', label: p.name, detail: p.location ? rangeAddr(p.location.ref) : undefined, go: go(si, p.location?.ref ?? cell(p.anchor.row, p.anchor.col)) });
    s.charts.forEach((c, i) => items.push({ key: `c:${c.id}`, kind: 'chart', label: c.title || c.alt || T('Chart {n}', { n: i + 1 }), go: go(si, cell(c.anchor.row, c.anchor.col)) }));
    s.pictures.forEach((p, i) => items.push({ key: `i:${p.id}`, kind: 'picture', label: p.alt || T('Picture {n}', { n: i + 1 }), go: go(si, cell(p.anchor.row, p.anchor.col)) }));
    (s.shapes ?? []).forEach((sh, i) => items.push({
      key: `s:${sh.id}`, kind: 'shape', label: sh.text?.split('\n')[0]?.slice(0, 40) || sh.alt || (sh.kind === 'textBox' ? T('Text box {n}', { n: i + 1 }) : T('Shape {n}', { n: i + 1 })),
      go: () => { go(si, cell(sh.anchor.row, sh.anchor.col))(); setObjUi(store, { shape: sh.id }); },
    }));
    for (const [k, t] of s.comments ?? []) items.push({ key: `m:${k}`, kind: 'comment', label: t.text.split('\n')[0]!.slice(0, 50) || N('(empty)'), detail: `${rangeAddr(cell(keyRow(k), keyCol(k)))} · ${t.author}${t.resolved ? ` · ${T('resolved')}` : ''}`, go: go(si, cell(keyRow(k), keyCol(k))) });
    for (const [k, n] of s.notes) if (!s.comments?.has(k)) items.push({ key: `n:${k}`, kind: 'note', label: n.text.split('\n')[0]!.slice(0, 50) || N('(empty)'), detail: rangeAddr(cell(keyRow(k), keyCol(k))), go: go(si, cell(keyRow(k), keyCol(k))) });
    for (const n of store.wb.names) {
      const t = nameTarget(store, n.formula);
      if (t?.sheet !== si || (n.sheet !== undefined && n.sheet !== si && n.sheet !== t.sheet)) continue;
      items.push({ key: `d:${n.name}:${n.sheet ?? ''}`, kind: 'name', label: n.name, detail: rangeAddr(t.range), hidden: n.hidden, go: go(si, t.range) });
    }
    return { index: si, name: s.name, state: s.state, items };
  });
};

export const NavPane = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const tree = useMemo(() => (ui.navPane ? collect(store) : []), [store.getVersion(), store.layoutVersion, ui.navPane]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!ui.navPane) return null;
  const q = ui.navSearch.trim().toLowerCase();
  const match = (t: string) => !q || t.toLowerCase().includes(q);
  const sheets = tree.filter((s) => ui.navHidden || s.state === 'visible');
  return (
    <aside className="ws-pane ws-left" aria-label={T('Navigation')} style={{ top: vp.headerH, height: Math.max(0, vp.height - vp.headerH), left: vp.rtl ? Math.max(0, vp.width - vp.headerW - 316) : vp.headerW }} onPointerDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
      <div className="phead"><h2>{T('Navigation')}</h2><button className="ico" aria-label={T('Close navigation')} onClick={() => { setUi(store, { navPane: false }); focusGrid(); }}><X size={16} strokeWidth={S} /></button></div>
      <div className="ws-search"><Search size={14} strokeWidth={S} /><input aria-label={T('Search the workbook')} placeholder={T('Search sheets and objects')} value={ui.navSearch} onChange={(e) => setUi(store, { navSearch: e.target.value })} /></div>
      <label className="b-check ws-pad"><input type="checkbox" checked={ui.navHidden} onChange={(e) => setUi(store, { navHidden: e.target.checked })} />{' '}{T('Show hidden sheets and names')}</label>
      <div className="ws-tree" role="tree">
        {sheets.map((s) => {
          const items = s.items.filter((it) => (ui.navHidden || !it.hidden) && (match(it.label) || match(it.detail ?? '') || match(s.name)));
          if (q && !items.length && !match(s.name)) return null;
          const isOpen = q ? true : open[s.index] ?? s.index === store.sheetIndex;
          const groups = (Object.keys(KIND_LABEL) as NavItem['kind'][]).map((k) => [k, items.filter((it) => it.kind === k)] as const).filter(([, l]) => l.length);
          return (
            <div key={s.index} role="treeitem" aria-expanded={isOpen} aria-selected={s.index === store.sheetIndex}>
              <div className={`ws-node sheet${s.index === store.sheetIndex ? ' on' : ''}`}>
                <button className="ico sm" aria-label={isOpen ? T('Collapse') : T('Expand')} onClick={() => setOpen({ ...open, [s.index]: !isOpen })}>{isOpen ? <ChevronDown size={14} strokeWidth={S} /> : <ChevronRight size={14} strokeWidth={S} />}</button>
                <button className="ws-nodebtn" onClick={() => { if (s.state === 'visible') { store.setActiveSheet(s.index); focusGrid(); } else store.notify(T('That sheet is hidden. Unhide it to go there.')); }}>
                  <FileSpreadsheet size={14} strokeWidth={S} /><span>{s.name}</span>{s.state !== 'visible' && <em className="muted small">{T('hidden')}</em>}<span className="ws-count">{items.length || ''}</span>
                </button>
              </div>
              {isOpen && groups.map(([k, list]) => {
                const Icon = ICONS[k];
                return (
                  <div key={k} role="group" className="ws-group">
                    <div className="ws-grouplabel">{Tr(KIND_LABEL[k])}</div>
                    {list.map((it) => (
                      <button key={it.key} role="treeitem" className="ws-node item" onClick={it.go} title={it.detail ? `${it.label} (${it.detail})` : it.label}>
                        <Icon size={13} strokeWidth={S} /><span className="ws-label">{it.label}</span>{it.detail && <span className="muted small">{it.detail}</span>}
                      </button>
                    ))}
                  </div>
                );
              })}
              {isOpen && !groups.length && <div className="ws-empty muted small">{q ? T('No matches') : T('Nothing else on this sheet')}</div>}
            </div>
          );
        })}
      </div>
    </aside>
  );
};

// ---------- Versions ----------

export const VersionsPane = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const [renaming, setRenaming] = useState<{ id: string; text: string } | null>(null);
  const open = ui.versionsPane;
  useEffect(() => { if (open) void refreshVersions(store); }, [open, store]);
  if (!open) return null;
  const items = versionList.items;
  const encrypted = !!store.file.password;
  const finishRename = async () => { if (!renaming) return; await renameVersion(renaming.id, renaming.text); setRenaming(null); await refreshVersions(store); };
  return (
    <aside className="ws-pane ws-right" aria-label={T('Version history')} style={{ top: vp.headerH, height: Math.max(0, vp.height - vp.headerH), left: vp.rtl ? 0 : Math.max(vp.headerW, vp.width - 330) }} onPointerDown={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
      <div className="phead"><h2>{T('Version history')}</h2><button className="ico" aria-label={T('Close version history')} onClick={() => { setUi(store, { versionsPane: false }); focusGrid(); }}><X size={16} strokeWidth={S} /></button></div>
      <div className="ws-tools">
        <button className="primary sm" disabled={encrypted || !!preview.current} onClick={() => void snapshotNow(store, 'manual', `Saved ${new Date().toLocaleString([], { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })}`).then((m) => store.notify(m ? T('Version saved.') : T('Nothing changed since the last version.')))}>
          <Save size={14} strokeWidth={S} />{' '}{T('Save a version now')}</button>
        <label className="b-check"><input type="checkbox" checked={settings.versionsOn} onChange={(e) => saveSettings(store, { versionsOn: e.target.checked })} />{' '}{T('Automatic')}</label>
      </div>
      <div className="ws-tools small muted">
        <label>{T('Every')}{' '}<input className="ws-num" type="number" min={1} max={240} value={settings.versionsEvery} onChange={(e) => saveSettings(store, { versionsEvery: Math.max(1, Number(e.target.value) || 5) })} /> {T('min')}</label>
        <label>{T('Keep')}{' '}<input className="ws-num" type="number" min={1} max={500} value={settings.versionsKeep} onChange={(e) => saveSettings(store, { versionsKeep: Math.max(1, Math.min(500, Number(e.target.value) || 50)) })} /></label>
      </div>
      <div className="ws-list">
        {encrypted && <p className="muted small ws-pad">{T('Versions aren\'t kept for encrypted workbooks, so the password protects every copy.')}</p>}
        {versionList.error && <p className="muted small ws-pad">{versionList.error}</p>}
        {!encrypted && !versionList.error && !items.length && <p className="muted small ws-pad">{T('No versions yet. A version is saved when you save, and every few minutes while you work.')}</p>}
        {items.map((v) => (
          <div key={v.id} className={`ws-version${preview.current?.meta.id === v.id ? ' on' : ''}`}>
            {renaming?.id === v.id ? (
              <input autoFocus aria-label={T('Version name')} value={renaming.text} maxLength={120} onChange={(e) => setRenaming({ id: v.id, text: e.target.value })}
                onBlur={() => void finishRename()} onKeyDown={(e) => { if (e.key === 'Enter') void finishRename(); if (e.key === 'Escape') setRenaming(null); }} />
            ) : (
              <button className="ws-vmain" onClick={() => void openPreview(store, v.id)} title={T('Preview this version')}>
                <History size={14} strokeWidth={S} />
                <span><b>{v.name || formatWhen(v.at)}</b><span className="muted small">{v.name ? `${formatWhen(v.at)} · ` : ''}{formatSize(v.size)} · {v.reason === 'save' ? 'saved' : v.reason === 'restore' ? T('before restore') : v.reason === 'manual' ? T('saved by you') : 'automatic'}</span></span>
              </button>
            )}
            <div className="ws-vbtns">
              <button className="ico sm" aria-label={T('Preview')} title={T('Preview')} onClick={() => void openPreview(store, v.id)}><Eye size={14} strokeWidth={S} /></button>
              <button className="ico sm" aria-label={T('Name this version')} title={T('Name')} onClick={() => setRenaming({ id: v.id, text: v.name ?? '' })}><Pencil size={14} strokeWidth={S} /></button>
              <button className="ico sm" aria-label={T('Delete this version')} title={T('Delete')} onClick={() => { if (confirm(T('Delete this version?'))) void deleteVersion(v.id).then(() => refreshVersions(store)); }}><Trash2 size={14} strokeWidth={S} /></button>
            </div>
          </div>
        ))}
      </div>
    </aside>
  );
};
