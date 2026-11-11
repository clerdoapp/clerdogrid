// Things drawn over the grid: focus cell bands, error indicators with their menu, value autocomplete,
// split bars, and the version preview banner.
import { useEffect, useRef, useState, type ReactElement } from 'react';
import { AlertTriangle, Copy, History, RotateCcw, X } from 'lucide-react';
import { cellKey, keyCol, keyRow } from '../../model/workbook';
import { openDialog, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { paneFor, panesOf, rectIn, SPLIT_BAR, type Viewport } from '../../ui/render';
import { focusGrid } from '../../ui/FormulaBar';
import { settings } from './settings';
import { setUi, ui } from './ui';
import { convertToNumbers, convertTwoDigitYears, copyNeighbourFormula, extendRange, ignoreIssue, issuesAt, lockCells } from './errors';
import { closePreview, copyPreviewToNew, preview, restorePreview } from './history';
import { T, Tr, getLang } from '../../ui/i18n';

const S = 1.9;

const hexA = (hex: string, a: number): string => {
  const m = /^#?([0-9a-f]{6})/i.exec(hex);
  if (!m) return `rgba(245,197,66,${a})`;
  const n = parseInt(m[1]!, 16);
  return `rgba(${n >> 16},${(n >> 8) & 255},${n & 255},${a})`;
};

// ---------- Focus cell ----------

export const FocusLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  if (!ui.focusCell || store.editing?.inBar === false) return null;
  const { row, col } = store.sel.active;
  const m = store.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
  const color = hexA(settings.focusColor, 0.22);
  const edge = hexA(settings.focusColor, 0.9);
  const out: ReactElement[] = [];
  panesOf(vp).forEach((p, i) => {
    const rowRect = rectIn(p, vp, { r1: m.r1, r2: m.r2, c1: p.c1, c2: p.c2 });
    const colRect = rectIn(p, vp, { r1: p.r1, r2: p.r2, c1: m.c1, c2: m.c2 });
    const cell = rectIn(p, vp, m);
    const clip = { left: p.x, top: p.y, width: p.w, height: p.h };
    out.push(
      <div key={i} className="ws-clip" style={clip}>
        {m.r2 >= p.r1 && m.r1 <= p.r2 && <>
          <div className="ws-band" style={{ left: 0, width: Math.max(0, cell.x - p.x), top: rowRect.y - p.y, height: rowRect.h, background: color, borderTop: `1px solid ${edge}`, borderBottom: `1px solid ${edge}` }} />
          <div className="ws-band" style={{ left: cell.x + cell.w - p.x, right: 0, top: rowRect.y - p.y, height: rowRect.h, background: color, borderTop: `1px solid ${edge}`, borderBottom: `1px solid ${edge}` }} />
        </>}
        {m.c2 >= p.c1 && m.c1 <= p.c2 && <>
          <div className="ws-band" style={{ top: 0, height: Math.max(0, cell.y - p.y), left: colRect.x - p.x, width: colRect.w, background: color, borderLeft: `1px solid ${edge}`, borderRight: `1px solid ${edge}` }} />
          <div className="ws-band" style={{ top: cell.y + cell.h - p.y, bottom: 0, left: colRect.x - p.x, width: colRect.w, background: color, borderLeft: `1px solid ${edge}`, borderRight: `1px solid ${edge}` }} />
        </>}
      </div>,
    );
  });
  return <div className="ws-focus" aria-hidden="true">{out}</div>;
};

// ---------- Error indicators ----------

export const ErrorLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  if (!settings.errorChecking || store.view.showFormulas) return null;
  const sheet = store.sheet;
  const marks: ReactElement[] = [];
  const size = Math.max(5, 6 * vp.zoom);
  for (const p of panesOf(vp)) {
    const cells = (p.r2 - p.r1 + 1) * (p.c2 - p.c1 + 1);
    const visit = (row: number, col: number) => {
      if (marks.length > 3000) return;
      if (!issuesAt(store, row, col).length) return;
      const m = store.mergeAt(row, col);
      if (m && (m.r1 !== row || m.c1 !== col)) return;
      const r = rectIn(p, vp, m ?? { r1: row, c1: col, r2: row, c2: col });
      if (r.x + r.w < p.x || r.x > p.x + p.w || r.y + r.h < p.y || r.y > p.y + p.h) return;
      marks.push(<div key={`${p.x},${p.y},${row},${col}`} className="ws-tri" style={vp.rtl
        ? { left: r.x + r.w - size - 1, top: r.y + 1, borderWidth: `${size}px 0 0 ${size}px`, borderColor: `${settings.errorColor} transparent transparent transparent` }
        : { left: r.x + 1, top: r.y + 1, borderWidth: `${size}px ${size}px 0 0`, borderColor: `${settings.errorColor} transparent transparent transparent` }} />);
    };
    if (cells < sheet.cells.size) { for (let row = p.r1; row <= p.r2; row++) for (let col = p.c1; col <= p.c2; col++) if (sheet.cells.has(cellKey(row, col))) visit(row, col); }
    else for (const k of sheet.cells.keys()) { const row = keyRow(k), col = keyCol(k); if (row >= p.r1 && row <= p.r2 && col >= p.c1 && col <= p.c2) visit(row, col); }
  }
  const { row, col } = store.sel.active;
  const issues = store.editing ? [] : issuesAt(store, row, col);
  let button: ReactElement | null = null;
  if (issues.length) {
    const m = store.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
    const r = rectIn(paneFor(panesOf(vp), vp, m.r1, m.c1), vp, m);
    const left = vp.rtl ? (r.x + r.w + 24 <= vp.width - vp.headerW ? r.x + r.w + 2 : r.x - 24) : r.x - 24 >= vp.headerW ? r.x - 24 : r.x + r.w + 2;
    button = <ErrorButton store={store} left={left} top={r.y} issues={issues} vpWidth={vp.width} />;
  }
  return <div className="ws-errors">{marks}{button}</div>;
};

const ErrorButton = ({ store, left, top, issues, vpWidth }: { store: GridStore; left: number; top: number; issues: ReturnType<typeof issuesAt>; vpWidth: number }) => {
  const [help, setHelp] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!ui.errorMenu) return;
    const f = (e: PointerEvent) => { if (!ref.current?.contains(e.target as Node)) setUi(store, { errorMenu: false }); };
    document.addEventListener('pointerdown', f, true);
    return () => document.removeEventListener('pointerdown', f, true);
  }, [store, ui.errorMenu]);
  const act = (f: () => void) => { setUi(store, { errorMenu: false }); f(); focusGrid(); };
  const first = issues[0]!;
  return (
    <div ref={ref} className="ws-errbtn-wrap" style={{ left, top }} onPointerDown={(e) => e.stopPropagation()}>
      <button className="ws-errbtn" aria-label={T('{title}: show options', { title: first.title })} title={Tr(first.title)} aria-expanded={ui.errorMenu} onClick={() => { setHelp(false); setUi(store, { errorMenu: !ui.errorMenu }); }}>
        <AlertTriangle size={14} strokeWidth={S} />
      </button>
      {ui.errorMenu && (
        <div className="ws-errmenu" role="menu" style={left > vpWidth - 280 ? { right: 0 } : undefined}>
          {issues.map((i) => (
            <div key={i.rule} className="ws-errgroup">
              <div className="ws-errtitle">{Tr(i.title)}</div>
              {i.rule === 'numberStoredAsText' && <button role="menuitem" onClick={() => act(() => convertToNumbers(store))}>{T('Convert to number')}</button>}
              {i.rule === 'formula' && <button role="menuitem" onClick={() => act(() => copyNeighbourFormula(store))}>{T('Copy formula from next to it')}</button>}
              {i.rule === 'formulaRange' && <button role="menuitem" onClick={() => act(() => extendRange(store))}>{T('Update formula to include cells')}</button>}
              {i.rule === 'unlockedFormula' && <button role="menuitem" onClick={() => act(() => lockCells(store))}>{T('Lock cell')}</button>}
              {i.rule === 'twoDigitTextYear' && <>
                <button role="menuitem" onClick={() => act(() => convertTwoDigitYears(store, 2000))}>{T('Convert to 20XX')}</button>
                <button role="menuitem" onClick={() => act(() => convertTwoDigitYears(store, 1900))}>{T('Convert to 19XX')}</button>
              </>}
              {(i.rule === 'evalError' || i.rule === 'formula' || i.rule === 'formulaRange') && <button role="menuitem" onClick={() => act(() => run(store, 'formula.evaluate'))}>{T('Show calculation steps')}</button>}
              {i.rule === 'evalError' && <button role="menuitem" onClick={() => act(() => run(store, 'formula.tracePrecedents'))}>{T('Trace precedents')}</button>}
              <button role="menuitem" onClick={() => act(() => ignoreIssue(store, i.rule))}>{T('Ignore error')}</button>
            </div>
          ))}
          <button role="menuitem" onClick={() => setHelp(!help)}>{help ? T('Hide help') : T('Help on this error')}</button>
          {help && <p className="ws-errhelp">{issues.map((i) => i.help).join(' ')}</p>}
          <button role="menuitem" onClick={() => act(() => openDialog(store, 'workspace.options', { tab: 'proofing' }))}>{T('Error checking options…')}</button>
        </div>
      )}
    </div>
  );
};

// ---------- Split bars ----------

export const SplitBars = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const sp = store.sheet.view.split;
  const drag = useRef<{ axis: 'x' | 'y'; start: number; orig: number } | null>(null);
  if (!sp || !vp.split) return null;
  const down = (axis: 'x' | 'y') => (e: React.PointerEvent) => {
    e.stopPropagation(); e.preventDefault();
    (e.target as Element).setPointerCapture(e.pointerId);
    drag.current = { axis, start: axis === 'x' ? e.clientX : e.clientY, orig: axis === 'x' ? sp.x : sp.y };
  };
  const move = (e: React.PointerEvent) => {
    const d = drag.current;
    if (!d) return;
    const delta = ((d.axis === 'x' ? e.clientX : e.clientY) - d.start) / vp.zoom * (d.axis === 'x' && vp.rtl ? -1 : 1);
    const max = d.axis === 'x' ? (vp.width - vp.headerW - 40) / vp.zoom : (vp.height - vp.headerH - 40) / vp.zoom;
    sp[d.axis] = Math.max(0, Math.min(max, d.orig + delta));
    store.emit();
  };
  const up = () => {
    const d = drag.current;
    drag.current = null;
    if (!d) return;
    if (sp[d.axis] < 12) sp[d.axis] = 0;
    if (!sp.x && !sp.y) delete store.sheet.view.split;
    store.file.dirty = true;
    store.touch();
  };
  const remove = (axis: 'x' | 'y') => () => { sp[axis] = 0; if (!sp.x && !sp.y) delete store.sheet.view.split; store.touch(); };
  return (
    <>
      {vp.split.h > 0 && <div className="ws-splitbar h" role="separator" aria-orientation="horizontal" aria-label={T('Split bar')} title={T('Drag to move the split; double-click to remove it')}
        style={{ top: vp.headerH + vp.split.h, height: SPLIT_BAR, left: 0, right: 0 }} onPointerDown={down('y')} onPointerMove={move} onPointerUp={up} onDoubleClick={remove('y')} />}
      {vp.split.w > 0 && <div className="ws-splitbar v" role="separator" aria-orientation="vertical" aria-label={T('Split bar')} title={T('Drag to move the split; double-click to remove it')}
        style={{ left: vp.rtl ? vp.width - vp.headerW - vp.split.w - SPLIT_BAR : vp.headerW + vp.split.w, width: SPLIT_BAR, top: 0, bottom: 0 }} onPointerDown={down('x')} onPointerMove={move} onPointerUp={up} onDoubleClick={remove('x')} />}
    </>
  );
};

// ---------- Version preview banner ----------

export const PreviewBanner = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const p = preview.current;
  if (!p) return null;
  return (
    <div className="ws-banner" role="region" aria-label={T('Version preview')} style={{ left: vp.headerW + 8, width: Math.max(200, vp.width - vp.headerW - 16 - (ui.versionsPane ? 330 : 0)), top: Math.max(vp.headerH + 8, vp.height - 54) }} onPointerDown={(e) => e.stopPropagation()}>
      <History size={16} strokeWidth={S} />
      <span><b>{p.meta.name || T('Earlier version')}</b> {T('from {time}', { time: new Date(p.meta.at).toLocaleString(getLang()) })} · {T('read-only')}</span>
      <button className="primary sm" onClick={() => void restorePreview(store)}><RotateCcw size={14} strokeWidth={S} />{' '}{T('Restore')}</button>
      <button className="sm" onClick={() => void copyPreviewToNew(store)}><Copy size={14} strokeWidth={S} />{' '}{T('Copy to new workbook')}</button>
      <button className="ico" aria-label={T('Close preview')} onClick={() => closePreview(store)}><X size={16} strokeWidth={S} /></button>
    </div>
  );
};

