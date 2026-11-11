import { useEffect, useMemo, useRef, useState } from 'react';
import { ChevronLeft, ChevronRight, ExternalLink, FileDown, Printer, Settings2, X } from 'lucide-react';
import type { DialogProps } from '../../ui/dialogs/Dialog';
import { openDialog } from '../../app/commands';
import { COMMON_PAPERS, MARGIN_PRESETS, paperByCode } from './paper';
import { DEFAULT_MARGINS } from './paginate';
import { loadFonts, pageToSvg, planPrint, type PrintOptions } from './print';
import { browserEnv, canOpenPdf, openPdf, printJob, savePdf } from './browser';
import { setupOf, updateSetup } from './setup';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;
type What = PrintOptions['what'];

const scalingOf = (s: ReturnType<typeof setupOf>): string => {
  if (s.fitWidth === undefined && s.fitHeight === undefined) return (s.scale ?? 100) === 100 ? 'none' : 'custom';
  if (s.fitWidth === 1 && s.fitHeight === 1) return 'page';
  if (s.fitWidth === 1 && s.fitHeight === 0) return 'cols';
  if (s.fitWidth === 0 && s.fitHeight === 1) return 'rows';
  return 'custom';
};

/** File → Print: page previews, settings beside them and the print and PDF buttons. */
export const PrintPane = ({ store, close, data }: DialogProps<{ what?: What }>) => {
  const [what, setWhat] = useState<What>(data?.what ?? 'sheet');
  const [ignoreArea, setIgnoreArea] = useState(false);
  const [index, setIndex] = useState(0);
  const [ready, setReady] = useState(0);
  const [busy, setBusy] = useState('');
  const version = store.getVersion();
  const opts: PrintOptions = { what, ignorePrintArea: ignoreArea };
  const job = useMemo(() => planPrint(store, opts), [store, what, ignoreArea, version]);
  const env = useMemo(() => browserEnv(store), [store]);
  const n = job.pages.length;
  const cur = Math.min(index, Math.max(0, n - 1));

  useEffect(() => {
    let live = true;
    void loadFonts(store, job, env).then(() => { if (live) setReady((x) => x + 1); });
    return () => { live = false; };
  }, [job, env, store]);

  const svgs = useMemo(() => new Map<number, string>(), [job, ready]);
  const svgOf = (i: number) => {
    let s = svgs.get(i);
    if (s === undefined) { const pp = job.pages[i]; s = pp ? pageToSvg(store, pp, env, '100%', '100%') : ''; svgs.set(i, s); }
    return s;
  };

  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => { ref.current?.querySelector('.pp-thumb.on')?.scrollIntoView({ block: 'nearest' }); }, [cur]);

  const s = setupOf(store);
  const m = { ...DEFAULT_MARGINS, ...s.margins };
  const marginPreset = MARGIN_PRESETS.find((p) => (Object.keys(p.m) as (keyof typeof m)[]).every((k) => Math.abs(p.m[k] - m[k]) < 0.001))?.id ?? 'custom';
  const sheetIdx = store.sheetIndex;
  const run = async (label: string, fn: () => Promise<void>) => { setBusy(label); try { await fn(); } finally { setBusy(''); } };
  const pp = job.pages[cur];

  return (
    <div className="backdrop pp-backdrop" onKeyDown={(e) => {
      if (e.key === 'Escape') { e.stopPropagation(); close(); }
      if (e.key === 'PageDown' || (e.key === 'ArrowRight' && !(e.target instanceof HTMLInputElement || e.target instanceof HTMLSelectElement))) setIndex(Math.min(n - 1, cur + 1));
      if (e.key === 'PageUp' || (e.key === 'ArrowLeft' && !(e.target instanceof HTMLInputElement || e.target instanceof HTMLSelectElement))) setIndex(Math.max(0, cur - 1));
    }}>
      <div className="pp" role="dialog" aria-modal="true" aria-label={T('Print')}>
        <aside className="pp-side">
          <div className="pp-title"><h2>{T('Print')}</h2><button className="ico" aria-label={T('Close')} onClick={close}><X size={16} strokeWidth={S} /></button></div>
          <button className="primary pp-big" disabled={!!busy || !n} onClick={() => void run(N('Preparing…'), () => printJob(store, opts))}><Printer size={16} strokeWidth={S} /> {busy === 'Preparing…' ? Tr(busy) : T('Print')}</button>
          <div className="pp-row">
            <button disabled={!!busy} onClick={() => void run(N('Saving…'), () => savePdf(store, opts))}><FileDown size={15} strokeWidth={S} /> {busy === 'Saving…' ? Tr(busy) : T('Save as PDF')}</button>
            <button disabled={!!busy} onClick={() => void run(N('Opening…'), () => openPdf(store, opts))} title={canOpenPdf() ? T('Open in the PDF viewer') : T('Open the PDF in a new tab')}><ExternalLink size={15} strokeWidth={S} /> {busy === 'Opening…' ? Tr(busy) : T('Open PDF')}</button>
          </div>
          <label className="field"><span>{T('What to print')}</span>
            <select value={what} onChange={(e) => { setWhat(e.target.value as What); setIndex(0); }}>
              <option value="sheet">{T('Active sheet')}</option><option value="workbook">{T('Entire workbook')}</option><option value="selection">{T('Selection')}</option>
            </select>
          </label>
          <label className="b-check"><input type="checkbox" checked={ignoreArea} onChange={(e) => setIgnoreArea(e.target.checked)} /> <span>{T('Ignore print area')}</span></label>
          <label className="field"><span>{T('Orientation')}</span>
            <select value={s.orientation ?? 'portrait'} onChange={(e) => updateSetup(store, N('Orientation'), { orientation: e.target.value as 'portrait' | 'landscape' }, sheetIdx)}>
              <option value="portrait">{T('Portrait')}</option><option value="landscape">{T('Landscape')}</option>
            </select>
          </label>
          <label className="field"><span>{T('Paper')}</span>
            <select value={String(s.paper ?? 1)} onChange={(e) => updateSetup(store, N('Paper size'), { paper: +e.target.value, paperWidth: undefined, paperHeight: undefined }, sheetIdx)}>
              {[...new Set([...COMMON_PAPERS, s.paper ?? 1])].map((c) => <option key={c} value={c}>{Tr(paperByCode(c).name)}</option>)}
            </select>
          </label>
          <label className="field"><span>{T('Margins')}</span>
            <select value={marginPreset} onChange={(e) => { const p = MARGIN_PRESETS.find((x) => x.id === e.target.value); if (p) updateSetup(store, N('Margins'), { margins: { ...p.m } }, sheetIdx); else openDialog(store, 'pageSetup', { tab: 'margins' }); }}>
              {MARGIN_PRESETS.map((p) => <option key={p.id} value={p.id}>{Tr(p.label)}</option>)}
              <option value="custom">{marginPreset === 'custom' ? T('Custom') : T('Custom…')}</option>
            </select>
          </label>
          <label className="field"><span>{T('Scaling')}</span>
            <select value={scalingOf(s)} onChange={(e) => {
              const v = e.target.value;
              if (v === 'none') updateSetup(store, N('Scaling'), { scale: undefined, fitWidth: undefined, fitHeight: undefined }, sheetIdx);
              else if (v === 'page') updateSetup(store, N('Scaling'), { scale: undefined, fitWidth: 1, fitHeight: 1 }, sheetIdx);
              else if (v === 'cols') updateSetup(store, N('Scaling'), { scale: undefined, fitWidth: 1, fitHeight: 0 }, sheetIdx);
              else if (v === 'rows') updateSetup(store, N('Scaling'), { scale: undefined, fitWidth: 0, fitHeight: 1 }, sheetIdx);
              else openDialog(store, 'pageSetup', { tab: 'page' });
            }}>
              <option value="none">{T('Actual size')}</option><option value="page">{T('Fit sheet on one page')}</option><option value="cols">{T('Fit all columns on one page')}</option><option value="rows">{T('Fit all rows on one page')}</option>
              <option value="custom">{scalingOf(s) === 'custom' ? (s.fitWidth !== undefined || s.fitHeight !== undefined ? T('Custom (fit)') : T('Custom ({p0})', { p0: `${s.scale}%` })) : T('Custom…')}</option>
            </select>
          </label>
          <button className="tool" onClick={() => openDialog(store, 'pageSetup')}><Settings2 size={16} strokeWidth={S} /><span>{T('Page setup…')}</span></button>
          {pp?.pg.fitPercent !== undefined && <div className="muted pl-note">{T('Printed at {percent}% to fit.', { percent: pp.pg.fitPercent })}</div>}
        </aside>
        <main className="pp-main">
          <div className="pp-page" style={pp ? { aspectRatio: `${pp.pg.paperW} / ${pp.pg.paperH}` } : undefined}>
            {pp ? <div className="pp-svg" dangerouslySetInnerHTML={{ __html: svgOf(cur) }} /> : <div className="pp-empty">{T('There is nothing to print.')}</div>}
          </div>
          <div className="pp-nav">
            <button className="ico" aria-label={T('Previous page')} disabled={cur <= 0} onClick={() => setIndex(cur - 1)}><ChevronLeft size={16} strokeWidth={S} /></button>
            <span>{T('Page')}{' '}<input aria-label={T('Page')} value={n ? cur + 1 : 0} onChange={(e) => { const v = parseInt(e.target.value, 10); if (v >= 1 && v <= n) setIndex(v - 1); }} /> {T('of {count}', { count: n })}</span>
            <button className="ico" aria-label={T('Next page')} disabled={cur >= n - 1} onClick={() => setIndex(cur + 1)}><ChevronRight size={16} strokeWidth={S} /></button>
          </div>
        </main>
        <div className="pp-thumbs" ref={ref} aria-label={T('Pages')}>
          {job.pages.map((p, i) => (
            <button key={i} className={`pp-thumb${i === cur ? ' on' : ''}`} aria-label={T('Page {p0}', { p0: i + 1 })} onClick={() => setIndex(i)}>
              <div className="pp-tpage" style={{ aspectRatio: `${p.pg.paperW} / ${p.pg.paperH}` }}>
                {Math.abs(i - cur) <= 12 ? <div className="pp-svg" dangerouslySetInnerHTML={{ __html: svgOf(i) }} /> : null}
              </div>
              <span>{i + 1}{what === 'workbook' && (i === 0 || job.pages[i - 1]!.sheet !== p.sheet) ? ` · ${store.wb.sheets[p.sheet]!.name}` : ''}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};
