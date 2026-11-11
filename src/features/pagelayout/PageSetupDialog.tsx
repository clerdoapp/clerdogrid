import { useMemo, useRef, useState } from 'react';
import type { HeaderPicture, PageSetup } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check, Radio, Tabs } from '../../ui/dialogs/shared';
import { openDialog } from '../../app/commands';
import { COMMON_PAPERS, MARGIN_PRESETS, PAPERS, paperByCode } from './paper';
import { DEFAULT_MARGINS, paperSize, parseAreas, parseTitleCols, parseTitleRows } from './paginate';
import { hfPresets, joinHf, parseHf, sectionText, splitHf, type HfContext } from './headerFooter';
import { setupOf, updateSetup } from './setup';
import { N, T, Tr } from '../../ui/i18n';

type Tab = 'page' | 'margins' | 'hf' | 'sheet';
const TABS = [['page', N('Page')], ['margins', N('Margins')], ['hf', N('Header and footer')], ['sheet', N('Sheet')]] as const;

const numIn = (v: string, lo: number, hi: number): number | undefined => {
  const n = parseFloat(v);
  return Number.isFinite(n) ? Math.max(lo, Math.min(hi, n)) : undefined;
};

export const PageSetupDialog = ({ store, close, data }: DialogProps<{ tab?: Tab }>) => {
  const [tab, setTab] = useState<Tab>(data?.tab ?? 'page');
  const [s, setS] = useState<PageSetup>(() => structuredClone(setupOf(store)));
  const [error, setError] = useState('');
  const set = (p: Partial<PageSetup>) => setS((x) => ({ ...x, ...p }));

  const ok = (): boolean => {
    for (const a of (s.printArea ?? '').split(',').filter((x) => x.trim())) if (!parseAreas(a).length) { setTab('sheet'); setError(T('"{p0}" isn\'t a range.', { p0: a.trim() })); return false; }
    if (s.printTitleRows && !parseTitleRows(s.printTitleRows)) { setTab('sheet'); setError(T('Rows to repeat must look like 1:2.')); return false; }
    if (s.printTitleCols && !parseTitleCols(s.printTitleCols)) { setTab('sheet'); setError(T('Columns to repeat must look like A:B.')); return false; }
    const m = { ...DEFAULT_MARGINS, ...s.margins };
    const { w, h } = paperSize(s);
    if ((m.left + m.right) * 72 >= w - 36 || (m.top + m.bottom) * 72 >= h - 36) { setTab('margins'); setError(T('The margins don\'t leave room on the paper.')); return false; }
    const clean: PageSetup = {
      ...s,
      printArea: s.printArea?.replace(/\s+/g, '').replace(/\$/g, '') || undefined,
      printTitleRows: s.printTitleRows?.replace(/\$/g, '').trim() || undefined,
      printTitleCols: s.printTitleCols?.replace(/\$/g, '').trim().toUpperCase() || undefined,
    };
    updateSetup(store, N('Page setup'), () => clean);
    return true;
  };

  return (
    <Dialog title={T('Page setup')} close={close} onOk={ok} wide
      footer={<button onClick={() => { if (ok()) openDialog(store, 'printPreview'); }}>{T('Print preview')}</button>}>
      <Tabs tabs={TABS} value={tab} onChange={(t) => { setTab(t); setError(''); }} />
      {tab === 'page' && <PageTab s={s} set={set} />}
      {tab === 'margins' && <MarginsTab s={s} set={set} />}
      {tab === 'hf' && <HfTab s={s} set={set} author={store.wb.props?.creator} ctx={{ page: 1, pages: 1, date: new Date(), file: `${store.file.name || 'Book1'}.xlsx`, path: '', sheet: store.sheet.name }} />}
      {tab === 'sheet' && <SheetTab s={s} set={set} />}
      {error && <div className="pl-error" role="alert">{error}</div>}
    </Dialog>
  );
};

const PageTab = ({ s, set }: { s: PageSetup; set: (p: Partial<PageSetup>) => void }) => {
  const fit = s.fitWidth !== undefined || s.fitHeight !== undefined;
  const papers = useMemo(() => [...COMMON_PAPERS.map((c) => paperByCode(c)), ...PAPERS.filter((p) => !COMMON_PAPERS.includes(p.code)).sort((a, b) => a.name.localeCompare(b.name))], []);
  const fmt = (p: { w: number; h: number; code: number }) => (p.code === 8 || p.code === 9 || p.code === 11 || p.code === 66 || p.code >= 27 && p.code <= 36 ? `${Math.round(p.w / 72 * 25.4)} × ${Math.round(p.h / 72 * 25.4)} mm` : `${+(p.w / 72).toFixed(3)} × ${+(p.h / 72).toFixed(3)} in`);
  return (
    <>
      <Field label={T('Orientation')}><Radio name="orient" row value={s.orientation ?? 'portrait'} options={[['portrait', N('Portrait')], ['landscape', N('Landscape')]]} onChange={(v) => set({ orientation: v })} /></Field>
      <Field label={T('Scaling')}>
        <div className="pl-scale">
          <label><input type="radio" name="scale" checked={!fit} onChange={() => set({ fitWidth: undefined, fitHeight: undefined, scale: s.scale ?? 100 })} />{' '}{T('Adjust to')}</label>
          <input type="number" min={10} max={400} aria-label={T('Percent of normal size')} value={s.scale ?? 100} disabled={fit} onChange={(e) => set({ scale: numIn(e.target.value, 10, 400) ?? 100 })} /> <span>{T('% of normal size')}</span>
        </div>
        <div className="pl-scale">
          <label><input type="radio" name="scale" checked={fit} onChange={() => set({ fitWidth: 1, fitHeight: 1, scale: undefined })} />{' '}{T('Fit to')}</label>
          <input type="number" min={0} max={32767} aria-label={T('Pages wide')} value={fit ? s.fitWidth ?? 1 : 1} disabled={!fit} onChange={(e) => set({ fitWidth: Math.round(numIn(e.target.value, 0, 32767) ?? 0) })} /> <span>{T('page(s) wide by')}</span>
          <input type="number" min={0} max={32767} aria-label={T('Pages tall')} value={fit ? s.fitHeight ?? 1 : 1} disabled={!fit} onChange={(e) => set({ fitHeight: Math.round(numIn(e.target.value, 0, 32767) ?? 0) })} /> <span>{T('page(s) tall')}</span>
        </div>
        <div className="muted pl-note">{T('0 means as many pages as needed in that direction.')}</div>
      </Field>
      <div className="grid2">
        <Field label={T('Paper size')}>
          <select value={s.paperWidth ? 'custom' : String(s.paper ?? 1)} onChange={(e) => set({ paper: +e.target.value, paperWidth: undefined, paperHeight: undefined })}>
            {s.paperWidth && <option value="custom">{T('Custom ({size})', { size: `${s.paperWidth} × ${s.paperHeight}` })}</option>}
            {papers.map((p) => <option key={p.code} value={p.code}>{Tr(p.name)} ({fmt(p)})</option>)}
          </select>
        </Field>
        <Field label={T('Print quality')}>
          <select value={String(s.horizontalDpi ?? 600)} onChange={(e) => set({ horizontalDpi: +e.target.value === 600 ? undefined : +e.target.value, verticalDpi: +e.target.value === 600 ? undefined : +e.target.value })}>
            {[150, 300, 600, 1200].map((d) => <option key={d} value={d}>{T('{dpi} dpi', { dpi: d })}</option>)}
          </select>
        </Field>
        <Field label={T('First page number')}>
          <input placeholder={T('Automatic')} value={s.firstPageNumber ?? ''} onChange={(e) => { const v = e.target.value.trim(); set({ firstPageNumber: v === '' || !/^-?\d+$/.test(v) ? undefined : Math.max(-32765, Math.min(32767, +v)) }); }} />
        </Field>
        <Field label={T('Copies')}><input type="number" min={1} max={999} value={s.copies ?? 1} onChange={(e) => set({ copies: Math.round(numIn(e.target.value, 1, 999) ?? 1) === 1 ? undefined : Math.round(numIn(e.target.value, 1, 999)!) })} /></Field>
      </div>
    </>
  );
};

const MarginsTab = ({ s, set }: { s: PageSetup; set: (p: Partial<PageSetup>) => void }) => {
  const m = { ...DEFAULT_MARGINS, ...s.margins };
  const { w, h } = paperSize(s);
  const preset = MARGIN_PRESETS.find((p) => (Object.keys(p.m) as (keyof typeof m)[]).every((k) => Math.abs(p.m[k] - m[k]) < 0.001))?.id ?? 'custom';
  const field = (k: keyof typeof m, label: string) => (
    <Field label={Tr(label)}><input type="number" step={0.05} min={0} max={20} value={+m[k].toFixed(3)} onChange={(e) => set({ margins: { ...m, [k]: numIn(e.target.value, 0, 20) ?? 0 } })} /></Field>
  );
  const sc = 150 / Math.max(w, h);
  return (
    <div className="pl-margins">
      <div className="pl-mfields">
        <Field label={T('Preset')}>
          <select value={preset} onChange={(e) => { const p = MARGIN_PRESETS.find((x) => x.id === e.target.value); if (p) set({ margins: { ...p.m } }); }}>
            {MARGIN_PRESETS.map((p) => <option key={p.id} value={p.id}>{Tr(p.label)}</option>)}
            {preset === 'custom' && <option value="custom">{T('Custom')}</option>}
          </select>
        </Field>
        <div className="grid2">{field('top', N('Top (in)'))}{field('bottom', N('Bottom (in)'))}{field('left', N('Left (in)'))}{field('right', N('Right (in)'))}{field('header', N('Header (in)'))}{field('footer', N('Footer (in)'))}</div>
        <Field label={T('Center on page')}><div className="b-checks"><Check label={T('Horizontally')} checked={!!s.centerH} onChange={(v) => set({ centerH: v })} /><Check label={T('Vertically')} checked={!!s.centerV} onChange={(v) => set({ centerV: v })} /></div></Field>
      </div>
      <svg className="pl-mpreview" viewBox={`0 0 ${w * sc + 4} ${h * sc + 4}`} width={w * sc + 4} height={h * sc + 4} aria-hidden="true">
        <rect x={2} y={2} width={w * sc} height={h * sc} className="paper" />
        <rect x={2 + m.left * 72 * sc} y={2 + m.top * 72 * sc} width={Math.max(0, (w - (m.left + m.right) * 72) * sc)} height={Math.max(0, (h - (m.top + m.bottom) * 72) * sc)} className="area" />
        <line x1={2} x2={2 + w * sc} y1={2 + m.header * 72 * sc} y2={2 + m.header * 72 * sc} className="hf" />
        <line x1={2} x2={2 + w * sc} y1={2 + h * sc - m.footer * 72 * sc} y2={2 + h * sc - m.footer * 72 * sc} className="hf" />
      </svg>
    </div>
  );
};

type Variant = '' | 'EVEN' | 'FIRST';
const KEYS: Record<Variant, { h: 'header' | 'evenHeader' | 'firstHeader'; f: 'footer' | 'evenFooter' | 'firstFooter' }> = {
  '': { h: 'header', f: 'footer' }, EVEN: { h: 'evenHeader', f: 'evenFooter' }, FIRST: { h: 'firstHeader', f: 'firstFooter' },
};

const HfTab = ({ s, set, author, ctx }: { s: PageSetup; set: (p: Partial<PageSetup>) => void; author?: string; ctx: HfContext }) => {
  const [variant, setVariant] = useState<Variant>('');
  const presets = hfPresets(author);
  const variants: [Variant, string][] = [['', s.differentOddEven ? N('Odd pages') : s.differentFirst ? N('Other pages') : N('All pages')]];
  if (s.differentOddEven) variants.push(['EVEN', N('Even pages')]);
  if (s.differentFirst) variants.push(['FIRST', N('First page')]);
  const v = variants.some(([id]) => id === variant) ? variant : '';
  const pictures = s.headerPictures ?? {};
  const setPicture = (slot: string, pic: HeaderPicture | null) => {
    const next = { ...pictures };
    if (pic) next[slot] = pic; else delete next[slot];
    set({ headerPictures: next });
  };
  return (
    <>
      <div className="b-checks">
        <Check label={T('Different odd and even pages')} checked={!!s.differentOddEven} onChange={(x) => set({ differentOddEven: x })} />
        <Check label={T('Different first page')} checked={!!s.differentFirst} onChange={(x) => set({ differentFirst: x })} />
        <Check label={T('Scale with the sheet')} checked={s.scaleWithDoc !== false} onChange={(x) => set({ scaleWithDoc: x })} />
        <Check label={T('Align with page margins')} checked={s.alignWithMargins !== false} onChange={(x) => set({ alignWithMargins: x })} />
      </div>
      {variants.length > 1 && <Tabs tabs={variants} value={v} onChange={setVariant} />}
      {(['h', 'f'] as const).map((part) => {
        const key = KEYS[v][part];
        const value = s[key] ?? '';
        const label = part === 'h' ? T('Header') : T('Footer');
        return (
          <div key={part} className="pl-hfblock">
            <div className="pl-hfhead">
              <b>{Tr(label)}</b>
              <select aria-label={T('{label} presets', { label })} value={presets.find(([, c]) => c === value)?.[1] ?? '__custom'} onChange={(e) => { if (e.target.value !== '__custom') set({ [key]: e.target.value || undefined }); }}>
                {presets.map(([l, c]) => <option key={l} value={c}>{l}</option>)}
                <option value="__custom">{T('Custom')}</option>
              </select>
            </div>
            <HfEditor value={value} onChange={(x) => set({ [key]: x || undefined })} ctx={ctx} kind={part === 'h' ? 'H' : 'F'} variant={v} pictures={pictures} setPicture={setPicture} />
          </div>
        );
      })}
      <div className="muted pl-note">{T('Codes: &P page, &N pages, &D date, &T time, &F file, &Z path, &A sheet, &G picture, &B bold, &I italic, &U underline, &12 size, &KFF0000 color, &"Font,Bold" font.')}</div>
    </>
  );
};

const CODES: [string, string][] = [['&P', N('Page')], ['&N', N('Pages')], ['&D', N('Date')], ['&T', N('Time')], ['&F', N('File')], ['&Z', N('Path')], ['&A', N('Sheet')], ['&B', N('Bold')], ['&I', N('Italic')], ['&U', N('Underline')]];

const HfEditor = ({ value, onChange, ctx, kind, variant, pictures, setPicture }: {
  value: string; onChange: (v: string) => void; ctx: HfContext; kind: 'H' | 'F'; variant: Variant;
  pictures: Record<string, HeaderPicture>; setPicture: (slot: string, p: HeaderPicture | null) => void;
}) => {
  const parts = splitHf(value);
  const [focus, setFocus] = useState<'L' | 'C' | 'R'>('C');
  const refs = { L: useRef<HTMLTextAreaElement>(null), C: useRef<HTMLTextAreaElement>(null), R: useRef<HTMLTextAreaElement>(null) };
  const file = useRef<HTMLInputElement>(null);
  const update = (sec: 'L' | 'C' | 'R', text: string) => onChange(joinHf({ ...parts, [sec]: text }));
  const insert = (code: string) => {
    const el = refs[focus].current;
    const cur = parts[focus];
    const at = el ? el.selectionStart : cur.length;
    const end = el ? el.selectionEnd : cur.length;
    update(focus, cur.slice(0, at) + code + cur.slice(end));
    requestAnimationFrame(() => { el?.focus(); el?.setSelectionRange(at + code.length, at + code.length); });
  };
  const slot = (sec: 'L' | 'C' | 'R') => `${sec}${kind}${variant}`;
  const pick = async (f: File) => {
    if (!/^image\/(png|jpeg|gif|bmp)$/.test(f.type) || f.size > 20 * 1024 * 1024) return;
    const data = new Uint8Array(await f.arrayBuffer());
    let w = 96, h = 48;
    try { const bmp = await createImageBitmap(f); const sc = Math.min(1, 144 / bmp.width, 72 / bmp.height); w = bmp.width * sc * 0.75; h = bmp.height * sc * 0.75; bmp.close(); } catch { /* keep default size */ }
    setPicture(slot(focus), { data, mime: f.type, width: Math.round(w), height: Math.round(h) });
    if (!parts[focus].includes('&G')) insert('&G');
  };
  const secs = parseHf(value);
  return (
    <div className="pl-hfedit">
      <div className="pl-codes">
        {CODES.map(([c, l]) => <button key={c} type="button" className="stylechip" onMouseDown={(e) => e.preventDefault()} onClick={() => insert(c)}>{Tr(l)}</button>)}
        <button type="button" className="stylechip" onMouseDown={(e) => e.preventDefault()} onClick={() => insert('&12')}>{T('Size')}</button>
        <button type="button" className="stylechip" onMouseDown={(e) => e.preventDefault()} onClick={() => insert('&KC00000')}>{T('Color')}</button>
        <button type="button" className="stylechip" onMouseDown={(e) => e.preventDefault()} onClick={() => file.current?.click()}>{T('Picture…')}</button>
        {pictures[slot(focus)] && <button type="button" className="stylechip" onClick={() => { setPicture(slot(focus), null); update(focus, parts[focus].replace(/&G/g, '')); }}>{T('Remove picture')}</button>}
        <input ref={file} type="file" accept="image/png,image/jpeg,image/gif,image/bmp" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) void pick(f); e.target.value = ''; }} />
      </div>
      <div className="pl-hfsecs">
        {(['L', 'C', 'R'] as const).map((sec) => (
          <label key={sec} className="field">
            <span>{sec === 'L' ? T('Left') : sec === 'C' ? T('Center') : T('Right')}</span>
            <textarea ref={refs[sec]} rows={2} value={parts[sec]} spellCheck={false} onFocus={() => setFocus(sec)} onChange={(e) => update(sec, e.target.value)} />
            <span className="pl-hfprev" style={{ textAlign: sec === 'L' ? 'left' : sec === 'C' ? 'center' : 'right' }}>{sectionText(secs[sec], ctx) || ' '}</span>
          </label>
        ))}
      </div>
    </div>
  );
};

const SheetTab = ({ s, set }: { s: PageSetup; set: (p: Partial<PageSetup>) => void }) => (
  <>
    <Field label={T('Print area')}><input placeholder={T('Whole sheet')} value={s.printArea ?? ''} onChange={(e) => set({ printArea: e.target.value || undefined })} /></Field>
    <div className="grid2">
      <Field label={T('Rows to repeat at top')}><input placeholder="e.g. 1:1" value={s.printTitleRows ?? ''} onChange={(e) => set({ printTitleRows: e.target.value || undefined })} /></Field>
      <Field label={T('Columns to repeat at left')}><input placeholder="e.g. A:A" value={s.printTitleCols ?? ''} onChange={(e) => set({ printTitleCols: e.target.value || undefined })} /></Field>
    </div>
    <Field label={T('Print')}>
      <div className="b-checks">
        <Check label={T('Gridlines')} checked={!!s.gridlines} onChange={(v) => set({ gridlines: v })} />
        <Check label={T('Row and column headings')} checked={!!s.headings} onChange={(v) => set({ headings: v })} />
        <Check label={T('Black and white')} checked={!!s.blackAndWhite} onChange={(v) => set({ blackAndWhite: v })} />
        <Check label={T('Draft quality')} checked={!!s.draft} onChange={(v) => set({ draft: v })} />
      </div>
    </Field>
    <div className="grid2">
      <Field label={T('Comments and notes')}>
        <select value={s.cellComments ?? 'none'} onChange={(e) => set({ cellComments: e.target.value as PageSetup['cellComments'] })}>
          <option value="none">{T('None')}</option><option value="atEnd">{T('At end of sheet')}</option><option value="asDisplayed">{T('As shown on the sheet')}</option>
        </select>
      </Field>
      <Field label={T('Cell errors as')}>
        <select value={s.errors ?? 'displayed'} onChange={(e) => set({ errors: e.target.value as PageSetup['errors'] })}>
          <option value="displayed">{T('Shown')}</option><option value="blank">{T('Blank')}</option><option value="dash">--</option><option value="NA">#N/A</option>
        </select>
      </Field>
    </div>
    <Field label={T('Page order')}><Radio name="order" value={s.order ?? 'downThenOver'} options={[['downThenOver', N('Down, then over')], ['overThenDown', N('Over, then down')]]} onChange={(v) => set({ order: v })} /></Field>
  </>
);
