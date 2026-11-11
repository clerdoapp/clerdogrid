// Page break preview (drawn over the grid) and page layout view (pages instead of the grid).
import { useEffect, useMemo, useRef, useState, type PointerEvent as RPointerEvent, type ReactElement } from 'react';
import type { CellRange, Sheet } from '../../model/workbook';
import { inRange, type GridStore } from '../../app/store';
import { run } from '../../app/commands';
import { panesOf, paneFor, rectIn, type Viewport } from '../../ui/render';
import { pageGrid, type Pagination } from './paginate';
import { loadFonts, pageToSvg, paginateSheet, type PrintPage } from './print';
import { browserEnv } from './browser';
import { moveBreak, updateSetup } from './setup';
import { joinHf, splitHf } from './headerFooter';
import { isRtlLang, N, T, Tr } from '../../ui/i18n';

const cache = new WeakMap<Sheet, { v: number; pg: Pagination }>();
export const sheetPagination = (store: GridStore): Pagination => {
  const hit = cache.get(store.sheet);
  if (hit && hit.v === store.getVersion()) return hit.pg;
  const pg = paginateSheet(store, store.sheetIndex);
  cache.set(store.sheet, { v: store.getVersion(), pg });
  return pg;
};

// ---------- Page break preview ----------

type Drag = { kind: 'row' | 'col'; from: number; area: number; edge?: 'r1' | 'r2' | 'c1' | 'c2'; pos: number; id: number };

export const BreakLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const [drag, setDrag] = useState<Drag | null>(null);
  if (store.sheet.view.pageView !== 'pageBreakPreview') return null;
  const pg = sheetPagination(store);
  const panes = panesOf(vp);
  const main = panes[0]!;
  const rectOf = (r: CellRange) => rectIn(paneFor(panes, vp, r.r1, r.c1), vp, r);
  const W = vp.width, H = vp.height;
  const L = vp.layout;
  // Pointer position → nearest row or column boundary.
  const boundary = (kind: 'row' | 'col', px: number): number => {
    const axis = kind === 'row' ? L.rows : L.cols;
    const sheetPx = (kind === 'row' ? main.scrollY + (px - main.y) / vp.zoom : main.scrollX + (vp.rtl ? main.x + main.w - px : px - main.x) / vp.zoom);
    const i = axis.at(Math.max(0, sheetPx));
    return sheetPx - axis.start(i) > axis.size(i) / 2 ? i + 1 : i;
  };
  const start = (d: Omit<Drag, 'id' | 'pos'>, pos: number) => (e: RPointerEvent) => {
    if (e.button !== 0) return;
    e.stopPropagation(); e.preventDefault();
    (e.currentTarget as Element).setPointerCapture(e.pointerId);
    setDrag({ ...d, pos, id: e.pointerId });
  };
  const wrapRect = () => (document.querySelector('.gridwrap') as HTMLElement | null)?.getBoundingClientRect();
  const move = (e: RPointerEvent) => {
    if (!drag || drag.id !== e.pointerId) return;
    const r = wrapRect();
    if (r) setDrag({ ...drag, pos: drag.kind === 'row' ? e.clientY - r.top : e.clientX - r.left });
  };
  const end = (e: RPointerEvent) => {
    if (!drag || drag.id !== e.pointerId) return;
    const to = boundary(drag.kind, drag.pos);
    const area = pg.areas[drag.area]!;
    setDrag(null);
    if (drag.edge) {
      const next = { ...area };
      if (drag.edge === 'r1') next.r1 = Math.min(to, area.r2);
      if (drag.edge === 'r2') next.r2 = Math.max(to - 1, area.r1);
      if (drag.edge === 'c1') next.c1 = Math.min(to, area.c2);
      if (drag.edge === 'c2') next.c2 = Math.max(to - 1, area.c1);
      const list = pg.areas.map((a, i) => (i === drag.area ? next : a));
      updateSetup(store, N('Print area'), { printArea: list.map((a) => rangeText(a)).join(',') });
      return;
    }
    if (to !== drag.from) moveBreak(store, drag.kind, drag.from, to, area);
  };
  const handlers = { onPointerMove: move, onPointerUp: end, onPointerCancel: () => setDrag(null) };

  let shade = `M0 0H${W}V${H}H0Z`;
  const lines: ReactElement[] = [];
  const marks: ReactElement[] = [];
  const grips: ReactElement[] = [];
  pg.bands.forEach((b, ai) => {
    const ar = rectOf(b.area);
    shade += `M${ar.x} ${ar.y}h${ar.w}v${ar.h}h${-ar.w}Z`;
    lines.push(<rect key={`a${ai}`} x={ar.x + 1.5} y={ar.y + 1.5} width={Math.max(0, ar.w - 3)} height={Math.max(0, ar.h - 3)} className="pb-area" />);
    const edge = (e: 'r1' | 'r2' | 'c1' | 'c2') => {
      const horiz = e === 'r1' || e === 'r2';
      const pos = e === 'r1' ? ar.y : e === 'r2' ? ar.y + ar.h : (e === 'c1') !== !!vp.rtl ? ar.x : ar.x + ar.w;
      grips.push(<div key={`e${ai}${e}`} className={`pb-grip ${horiz ? 'h' : 'v'}`} title={T('Drag to change the print area')}
        style={horiz ? { left: ar.x, width: ar.w, top: pos - 4 } : { top: ar.y, height: ar.h, left: pos - 4 }}
        onPointerDown={start({ kind: horiz ? 'row' : 'col', from: -1, area: ai, edge: e }, pos)} {...handlers} />);
    };
    edge('r1'); edge('r2'); edge('c1'); edge('c2');
    b.rowStarts.slice(1).forEach((r) => {
      const y = rectOf({ r1: r, c1: b.area.c1, r2: r, c2: b.area.c1 }).y;
      if (y < 0 || y > H) return;
      lines.push(<line key={`r${ai}-${r}`} x1={ar.x} x2={ar.x + ar.w} y1={y} y2={y} className={b.manualRows.has(r) ? 'pb-man' : 'pb-auto'} />);
      grips.push(<div key={`gr${ai}-${r}`} className="pb-grip h" title={T('Drag to move the page break')} style={{ left: ar.x, width: ar.w, top: y - 4 }} onPointerDown={start({ kind: 'row', from: r, area: ai }, y)} {...handlers} />);
    });
    b.colStarts.slice(1).forEach((c) => {
      const cr = rectOf({ r1: b.area.r1, c1: c, r2: b.area.r1, c2: c });
      const x = vp.rtl ? cr.x + cr.w : cr.x;
      if (x < 0 || x > W) return;
      lines.push(<line key={`c${ai}-${c}`} y1={ar.y} y2={ar.y + ar.h} x1={x} x2={x} className={b.manualCols.has(c) ? 'pb-man' : 'pb-auto'} />);
      grips.push(<div key={`gc${ai}-${c}`} className="pb-grip v" title={T('Drag to move the page break')} style={{ top: ar.y, height: ar.h, left: x - 4 }} onPointerDown={start({ kind: 'col', from: c, area: ai }, x)} {...handlers} />);
    });
  });
  for (const p of pg.pages) {
    const r = rectOf({ r1: p.rows[0], c1: p.cols[0], r2: p.rows[1], c2: p.cols[1] });
    if (r.x > W || r.y > H || r.x + r.w < vp.headerW || r.y + r.h < vp.headerH) continue;
    const size = Math.max(10, Math.min(72, r.h / 4, r.w / 4));
    const cx = Math.max(Math.min(r.x + r.w / 2, W - size * 2), r.x + size * 2), cy = Math.max(Math.min(r.y + r.h / 2, H - size), r.y + size);
    marks.push(<text key={`p${p.index}`} x={cx} y={cy} className="pb-mark" fontSize={size} textAnchor="middle" dominantBaseline="middle">{T('Page {page}', { page: p.number })}</text>);
  }
  return (
    <div className="pb-layer" style={{ position: 'absolute', left: 0, top: 0, width: W, height: H, pointerEvents: 'none' }}>
      <svg width={W} height={H} style={{ position: 'absolute', left: 0, top: 0 }} aria-hidden="true">
        <path d={shade} fillRule="evenodd" className="pb-shade" />
        {marks}
        {lines}
        {drag && (drag.kind === 'row' ? <line x1={0} x2={W} y1={drag.pos} y2={drag.pos} className="pb-drag" /> : <line y1={0} y2={H} x1={drag.pos} x2={drag.pos} className="pb-drag" />)}
      </svg>
      {grips}
    </div>
  );
};

const rangeText = (a: CellRange): string => {
  const c = (n: number) => { let s = ''; let x = n + 1; while (x > 0) { const m = (x - 1) % 26; s = String.fromCharCode(65 + m) + s; x = Math.floor((x - 1) / 26); } return s; };
  return a.r1 === a.r2 && a.c1 === a.c2 ? `${c(a.c1)}${a.r1 + 1}` : `${c(a.c1)}${a.r1 + 1}:${c(a.c2)}${a.r2 + 1}`;
};

// ---------- Page layout view ----------

const PX = 4 / 3;
type Zone = { kind: 'H' | 'F'; sec: 'L' | 'C' | 'R' };

export const PageLayoutView = ({ store }: { store: GridStore }) => {
  const [ready, setReady] = useState(0);
  const [editHf, setEditHf] = useState<(Zone & { page: number; text: string }) | null>(null);
  const [ruler, setRuler] = useState(() => { try { return localStorage.getItem('clerdogrid.ruler') !== '0'; } catch { return true; } });
  const [scroll, setScroll] = useState({ top: 0, h: 800 });
  const box = useRef<HTMLDivElement>(null);
  const version = store.getVersion();
  const pg = sheetPagination(store);
  const zoom = store.sheet.view.zoom / 100;
  const env = useMemo(() => browserEnv(store), [store]);
  const pages: PrintPage[] = useMemo(() => pg.pages.map((page) => ({ sheet: store.sheetIndex, pg, page, total: pg.pages.length, position: page.index + 1 })), [pg, store.sheetIndex]);
  useEffect(() => {
    let live = true;
    void loadFonts(store, { pages, title: '' }, env).then(() => { if (live) setReady((x) => x + 1); });
    return () => { live = false; };
  }, [pages, env, store]);
  useEffect(() => { box.current?.focus({ preventScroll: true }); }, []);

  // Pages in a grid: one column per band of columns, like the printout's order.
  const cols = pg.bands.length === 1 ? Math.max(1, pg.bands[0]!.colStarts.length) : 1;
  const pageW = pg.paperW * PX * zoom, pageH = pg.paperH * PX * zoom;
  const gap = 28, rulerH = ruler ? 18 : 0;
  const place = (i: number) => {
    const p = pages[i]!.page;
    if (cols === 1) return { x: 0, y: i * (pageH + gap + rulerH) };
    const b = pg.bands[0]!;
    const col = b.colStarts.indexOf(p.cols[0]), row = b.rowStarts.indexOf(p.rows[0]);
    return { x: (store.sheet.view.rtl ? cols - 1 - col : col) * (pageW + gap), y: row * (pageH + gap + rulerH) };
  };
  const rowsN = cols === 1 ? pages.length : pg.bands[0]!.rowStarts.length;
  const svgs = useMemo(() => new Map<number, string>(), [version, ready, pg]);
  const svgOf = (i: number) => {
    let s = svgs.get(i);
    if (s === undefined) { s = pageToSvg(store, pages[i]!, env, '100%', '100%'); svgs.set(i, s); }
    return s;
  };

  const sel = store.sel;
  const cellAt = (i: number, ex: number, ey: number): { row: number; col: number } | null => {
    const p = pages[i]!.page;
    const g = pageGrid(pg, p);
    let bx = (ex / (PX * zoom) - p.x) / pg.scale, by = (ey / (PX * zoom) - p.y) / pg.scale;
    if (store.sheet.view.rtl) bx = p.w - bx;
    const c = g.cols.find((q) => bx >= q.x && bx < q.x + q.w), r = g.rows.find((q) => by >= q.y && by < q.y + q.h);
    return c && r ? { row: r.index, col: c.index } : null;
  };
  const rectsOn = (i: number, ranges: CellRange[]) => {
    const p = pages[i]!.page;
    const g = pageGrid(pg, p);
    const out: { x: number; y: number; w: number; h: number }[] = [];
    for (const r of ranges) {
      const cs = g.cols.filter((q) => q.index >= r.c1 && q.index <= r.c2), rs = g.rows.filter((q) => q.index >= r.r1 && q.index <= r.r2);
      if (!cs.length || !rs.length) continue;
      let x = cs[0]!.x;
      const w = cs.at(-1)!.x + cs.at(-1)!.w - x;
      if (store.sheet.view.rtl) x = p.w - x - w;
      out.push({ x: (p.x + x * pg.scale) * PX * zoom, y: (p.y + rs[0]!.y * pg.scale) * PX * zoom, w: w * pg.scale * PX * zoom, h: (rs.at(-1)!.y + rs.at(-1)!.h - rs[0]!.y) * pg.scale * PX * zoom });
    }
    return out;
  };

  const editing = store.editing;
  const commitHf = () => {
    if (!editHf) return;
    const pp = pages[editHf.page]!;
    const s = pg.setup;
    const variant = s.differentFirst && pp.position === 1 ? 'first' : s.differentOddEven && pp.position % 2 === 0 ? 'even' : '';
    const key = (variant === 'first' ? (editHf.kind === 'H' ? 'firstHeader' : 'firstFooter') : variant === 'even' ? (editHf.kind === 'H' ? 'evenHeader' : 'evenFooter') : editHf.kind === 'H' ? 'header' : 'footer') as 'header';
    const parts = splitHf(s[key]);
    parts[editHf.sec] = editHf.text;
    const next = joinHf(parts);
    if (next !== (s[key] ?? '')) updateSetup(store, editHf.kind === 'H' ? N('Header') : N('Footer'), { [key]: next || undefined });
    setEditHf(null);
  };
  const openHf = (i: number, z: Zone) => {
    const pp = pages[i]!;
    const s = pg.setup;
    const src = s.differentFirst && pp.position === 1 ? (z.kind === 'H' ? s.firstHeader : s.firstFooter) : s.differentOddEven && pp.position % 2 === 0 ? (z.kind === 'H' ? s.evenHeader : s.evenFooter) : z.kind === 'H' ? s.header : s.footer;
    setEditHf({ ...z, page: i, text: splitHf(src)[z.sec] });
  };

  const visible = (i: number) => { const { y } = place(i); return y + pageH + rulerH > scroll.top - 400 && y < scroll.top + scroll.h + 400; };
  const active = pages.findIndex((pp) => inRange({ r1: pp.page.rows[0], c1: pp.page.cols[0], r2: pp.page.rows[1], c2: pp.page.cols[1] }, sel.active.row, sel.active.col));

  return (
    <div className="plv" ref={box} tabIndex={-1} role="region" aria-label={T('Page layout')} dir={store.sheet.view.rtl ? 'rtl' : 'ltr'}
      onScroll={(e) => setScroll({ top: e.currentTarget.scrollTop, h: e.currentTarget.clientHeight })}
      onKeyDown={(e) => {
        if (e.target instanceof HTMLTextAreaElement || e.target instanceof HTMLInputElement) { e.stopPropagation(); return; }
        if ((e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) || e.key === 'F2' || e.key === 'Enter') {
          e.stopPropagation(); e.preventDefault();
          if (e.key === 'Enter' && !e.shiftKey && editing) return;
          store.beginEdit(e.key.length === 1 ? e.key : undefined, true);
        }
        // Everything else (arrows, shortcuts, Delete) bubbles to the grid's own key handling.
      }}>
      <div className="plv-bar" dir={isRtlLang() ? 'rtl' : 'ltr'}>
        <span>{T('Page layout')}</span>
        <label className="b-check"><input type="checkbox" checked={ruler} onChange={(e) => { setRuler(e.target.checked); try { localStorage.setItem('clerdogrid.ruler', e.target.checked ? '1' : '0'); } catch { /* private mode */ } }} /> <span>{T('Ruler')}</span></label>
        <button className="tool" onClick={() => run(store, 'page.setup')}>{T('Page setup…')}</button>
        <button className="tool" onClick={() => run(store, 'page.viewNormal')}>{T('Back to normal view')}</button>
      </div>
      {!pages.length && <div className="plv-empty">{T('This sheet has nothing to print yet. Type in the grid in normal view, or set a print area.')}</div>}
      <div className="plv-pages" style={{ width: cols * (pageW + gap), height: rowsN * (pageH + gap + rulerH) }}>
        {pages.map((_pp, i) => {
          const { x, y } = place(i);
          const show = visible(i);
          const m = pg.margins;
          const z = PX * zoom;
          return (
            <div key={i} className="plv-wrap" style={{ left: x, top: y, width: pageW }}>
              {ruler && <Ruler width={pageW} left={m.left * z} right={(pg.paperW - m.right) * z} zoom={zoom} />}
              <div className={`plv-page${i === active ? ' on' : ''}`} style={{ width: pageW, height: pageH }}
                onPointerDown={(e) => {
                  if ((e.target as Element).closest('.plv-hf, .plv-hfedit')) return;
                  const r = e.currentTarget.getBoundingClientRect();
                  const c = cellAt(i, e.clientX - r.left, e.clientY - r.top);
                  if (!c) return;
                  if (store.editing) store.commitEdit();
                  store.select(c, e.shiftKey);
                  box.current?.focus({ preventScroll: true });
                }}
                onDoubleClick={(e) => {
                  const r = e.currentTarget.getBoundingClientRect();
                  if (cellAt(i, e.clientX - r.left, e.clientY - r.top)) store.beginEdit(undefined, true);
                }}>
                {show ? <div className="pp-svg" dangerouslySetInnerHTML={{ __html: svgOf(i) }} /> : null}
                {show && rectsOn(i, sel.ranges).map((r, k) => <div key={k} className="plv-sel" style={{ left: r.x, top: r.y, width: r.w, height: r.h }} />)}
                {show && rectsOn(i, [{ r1: sel.active.row, c1: sel.active.col, r2: sel.active.row, c2: sel.active.col }]).map((r, k) => (
                  <div key={`a${k}`} className="plv-active" style={{ left: r.x, top: r.y, width: r.w, height: r.h }}>
                    {editing && editing.row === sel.active.row && editing.col === sel.active.col && (
                      <textarea className="plv-editor" autoFocus value={editing.text} spellCheck={false} aria-label={T('Cell editor')}
                        onChange={(e) => store.setEditText(e.target.value)}
                        onKeyDown={(e) => {
                          e.stopPropagation();
                          if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); if (store.commitEdit()) { store.move(e.shiftKey ? -1 : 1, 0); box.current?.focus(); } }
                          else if (e.key === 'Tab') { e.preventDefault(); if (store.commitEdit()) { store.move(0, e.shiftKey ? -1 : 1); box.current?.focus(); } }
                          else if (e.key === 'Escape') { store.cancelEdit(); box.current?.focus(); }
                        }} />
                    )}
                  </div>
                ))}
                {show && (['H', 'F'] as const).map((kind) => (['L', 'C', 'R'] as const).map((sec) => {
                  const top = kind === 'H' ? 0 : (pg.paperH - m.bottom) * z;
                  const h = (kind === 'H' ? m.top : m.bottom) * z;
                  const left = sec === 'L' ? m.left * z : sec === 'C' ? pageW / 3 : (pg.paperW * 2 / 3) * z;
                  const w = sec === 'L' ? pageW / 3 - m.left * z : sec === 'C' ? pageW / 3 : pageW / 3 - m.right * z;
                  const open = editHf && editHf.page === i && editHf.kind === kind && editHf.sec === sec;
                  return (
                    <div key={kind + sec} className={`plv-hf${open ? ' on' : ''}`} style={{ left, top, width: Math.max(20, w), height: Math.max(12, h) }}
                      title={kind === 'H' ? T('Click to edit the header') : T('Click to edit the footer')} onClick={() => !open && openHf(i, { kind, sec })}>
                      {open && (
                        <div className="plv-hfedit">
                          <textarea autoFocus rows={2} value={editHf.text} spellCheck={false} aria-label={kind === 'H' ? T('Header') : T('Footer')}
                            onChange={(e) => setEditHf({ ...editHf, text: e.target.value })}
                            onBlur={commitHf}
                            onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Escape') setEditHf(null); if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitHf(); } }} />
                          <div className="plv-codes">
                            {[['&P', N('Page')], ['&N', N('Pages')], ['&D', N('Date')], ['&T', N('Time')], ['&F', N('File')], ['&A', N('Sheet')]].map(([c, l]) => (
                              <button key={c} type="button" onMouseDown={(e) => { e.preventDefault(); setEditHf({ ...editHf, text: editHf.text + c }); }}>{Tr(l)}</button>
                            ))}
                          </div>
                        </div>
                      )}
                      {!open && !splitHf(kind === 'H' ? pg.setup.header : pg.setup.footer)[sec] && sec === 'C' && <span className="plv-add">{kind === 'H' ? T('Add header') : T('Add footer')}</span>}
                    </div>
                  );
                }))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

const Ruler = ({ width, left, right, zoom }: { width: number; left: number; right: number; zoom: number }) => {
  const ticks: ReactElement[] = [];
  const inch = 72 * PX * zoom;
  for (let i = 0; i * inch / 4 <= width; i++) {
    const x = (i * inch) / 4;
    ticks.push(<line key={i} x1={x} x2={x} y1={i % 4 ? 12 : 4} y2={18} />);
    if (!(i % 4) && i) ticks.push(<text key={`t${i}`} x={x + 2} y={10}>{i / 4}</text>);
  }
  return (
    <svg className="plv-ruler" width={width} height={18} aria-hidden="true">
      <rect x={0} y={0} width={left} height={18} className="m" /><rect x={right} y={0} width={Math.max(0, width - right)} height={18} className="m" />
      {ticks}
    </svg>
  );
};
