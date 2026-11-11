// Printing and PDF in the app: bundled font files, pictures PDF can't hold, the print frame,
// and saving or opening the PDF (desktop dialog, file picker or download).
import type { GridStore } from '../../app/store';
import { desktop } from '../../app/files';
import { FontBook, type FontSource } from '../../formats/pdf-text';
import { isJpeg, isPng } from '../../formats/pdf-image';
import { fontFamily } from '../../ui/render';
import { jobToPdf, loadFonts, pageToSvg, planPrint, type PrintEnv, type PrintJob, type PrintOptions } from './print';
import type { TextStyle } from './painter';
import { T } from '../../ui/i18n';

// Only the weights the app already ships, so nothing extra is downloaded.
const FONT_URLS = import.meta.glob([
  '/node_modules/@fontsource/{carlito,caladea,arimo,tinos,cousine,hanken-grotesk}/files/*-{400,700}-{normal,italic}.woff',
  '/node_modules/@fontsource/{inter,lato,montserrat,open-sans,roboto,roboto-mono,noto-naskh-arabic}/files/*-{400,700}-normal.woff',
  // Chinese, Japanese and Korean: regular weight only (bold is drawn heavier); chunks are fetched on demand.
  '/node_modules/@fontsource/{noto-sans-sc,noto-sans-jp,noto-sans-kr}/files/*-400-normal.woff',
  '!/node_modules/@fontsource/*/files/*-{chinese-simplified,japanese,korean}-*.woff',
], { eager: true, query: '?url', import: 'default', exhaustive: true } as never) as unknown as Record<string, string>;

/** Which characters each font chunk holds, so export fetches only the chunks a sheet needs. */
const RANGES = import.meta.glob('/node_modules/@fontsource/{carlito,caladea,arimo,tinos,cousine,hanken-grotesk,inter,lato,montserrat,open-sans,roboto,roboto-mono,noto-naskh-arabic,noto-sans-sc,noto-sans-jp,noto-sans-kr}/unicode.json',
  { eager: true, import: 'default', exhaustive: true } as never) as unknown as Record<string, Record<string, string>>;

const browserFonts = (): FontSource => {
  const byName = new Map(Object.entries(FONT_URLS).map(([p, url]) => [p.slice(p.lastIndexOf('/') + 1), url]));
  return {
    ranges: () => Object.fromEntries(Object.entries(RANGES).map(([p, r]) => [p.split('/').at(-2)!, r])),
    files: () => [...byName.keys()],
    load: async (f) => {
      const url = byName.get(f);
      if (!url) throw new Error(T('Unknown font.'));
      const r = await fetch(url);
      if (!r.ok) throw new Error(T('Font not available.'));
      return new Uint8Array(await r.arrayBuffer());
    },
  };
};

let book: FontBook | null = null;
let measureCtx: CanvasRenderingContext2D | null | undefined;
const measureMissing = (text: string, req: { family?: string; bold?: boolean; italic?: boolean }): number | null => {
  if (measureCtx === undefined) measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
  if (!measureCtx) return null;
  measureCtx.font = `${req.italic ? 'italic ' : ''}${req.bold ? '700 ' : '400 '}1000px ${fontFamily(req.family)}`;
  return measureCtx.measureText(text).width;
};
export const fontBook = (): FontBook => (book ??= new FontBook(browserFonts(), { measureMissing }));

/** Text no bundled font covers (for example Chinese) is drawn by the system as a picture. */
const rasterize = (text: string, st: TextStyle) => {
  if (typeof document === 'undefined') return null;
  const px = 100;
  const c = document.createElement('canvas');
  const ctx = c.getContext('2d');
  if (!ctx) return null;
  const font = `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : '400 '}${px}px ${fontFamily(st.family)}`;
  ctx.font = font;
  const m = ctx.measureText(text);
  const ascent = Math.ceil(m.actualBoundingBoxAscent || px * 0.8) + 4, descent = Math.ceil(m.actualBoundingBoxDescent || px * 0.25) + 4;
  const w = Math.max(1, Math.ceil(m.width));
  const scale = 3;
  c.width = Math.min(16000, w * scale); c.height = (ascent + descent) * scale;
  ctx.scale(scale, scale);
  ctx.font = font;
  ctx.fillStyle = st.color;
  ctx.textBaseline = 'alphabetic';
  ctx.fillText(text, 0, ascent);
  const b64 = c.toDataURL('image/png').split(',')[1] ?? '';
  const bin = atob(b64);
  const png = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) png[i] = bin.charCodeAt(i);
  return { png, width: w, ascent, height: ascent + descent };
};

/** GIF, BMP, WebP and SVG pictures as PNG for the PDF. */
const convertPictures = async (store: GridStore, job: PrintJob): Promise<Map<Uint8Array, Uint8Array>> => {
  const out = new Map<Uint8Array, Uint8Array>();
  const sheets = new Set(job.pages.map((p) => p.sheet));
  const pics = [...sheets].flatMap((i) => [
    ...store.wb.sheets[i]!.pictures.map((p) => ({ data: p.data, mime: p.mime })),
    ...Object.values(store.wb.sheets[i]!.pageSetup?.headerPictures ?? {}).map((p) => ({ data: p.data, mime: p.mime })),
  ]);
  for (const p of pics) {
    if (isPng(p.data) || isJpeg(p.data) || out.has(p.data)) continue;
    try {
      const url = URL.createObjectURL(new Blob([p.data as BlobPart], { type: p.mime }));
      const img = new Image();
      img.src = url;
      await img.decode();
      const c = document.createElement('canvas');
      c.width = Math.max(1, img.naturalWidth || 300); c.height = Math.max(1, img.naturalHeight || 150);
      c.getContext('2d')!.drawImage(img, 0, 0, c.width, c.height);
      URL.revokeObjectURL(url);
      const blob = await new Promise<Blob | null>((res) => c.toBlob(res, 'image/png'));
      if (blob) out.set(p.data, new Uint8Array(await blob.arrayBuffer()));
    } catch { /* drawn as an empty box */ }
  }
  return out;
};

const urls = new WeakMap<Uint8Array, string>();
const objectUrl = (data: Uint8Array, mime: string): string => {
  let u = urls.get(data);
  if (!u) { u = URL.createObjectURL(new Blob([data as BlobPart], { type: mime })); urls.set(data, u); }
  return u;
};

export const browserEnv = (store: GridStore): PrintEnv => ({
  fonts: fontBook(),
  file: `${store.file.name || 'Book1'}.xlsx`,
  path: store.file.path ? store.file.path.replace(/[^\\/]*$/, '') : '',
  raster: rasterize,
  imageUrl: objectUrl,
});

export const makePdf = async (store: GridStore, o: PrintOptions): Promise<Uint8Array> => {
  store.commitEdit();
  const job = planPrint(store, o);
  const env = browserEnv(store);
  env.images = await convertPictures(store, job);
  await loadFonts(store, job, env);
  return jobToPdf(store, job, env);
};

type PickerWindow = Window & { showSaveFilePicker?: (o: unknown) => Promise<{ createWritable: () => Promise<{ write: (d: Blob) => Promise<void>; close: () => Promise<void> }> }> };
interface PdfBridge { savePdf?: (name: string, data: Uint8Array) => Promise<{ path: string } | null>; openPdf?: (name: string, data: Uint8Array) => Promise<boolean> }
const bridge = (): PdfBridge | undefined => (desktop as unknown as PdfBridge | undefined);

export const pdfName = (store: GridStore, o: PrintOptions): string => {
  const base = (store.file.name || 'Book1').replace(/[\\/:*?"<>|]/g, '_');
  return `${o.what === 'workbook' ? base : `${base} - ${store.wb.sheets[o.sheet ?? store.sheetIndex]?.name ?? 'Sheet'}`.replace(/[\\/:*?"<>|]/g, '_')}.pdf`;
};

export const canOpenPdf = (): boolean => !!bridge()?.openPdf;

/** Saves the PDF: desktop save dialog, the browser's file picker, or a download. */
export const savePdf = async (store: GridStore, o: PrintOptions): Promise<void> => {
  let data: Uint8Array;
  try { data = await makePdf(store, o); } catch (e) { store.notify(T('The PDF could not be made: {p0}', { p0: e instanceof Error ? e.message : T('unknown error') })); return; }
  const name = pdfName(store, o);
  const b = bridge();
  if (b?.savePdf) {
    try { const r = await b.savePdf(name, data); if (r) store.notify(T('Saved as PDF.')); } catch (e) { store.notify(e instanceof Error ? e.message : T('The PDF could not be saved.')); }
    return;
  }
  const blob = new Blob([data as BlobPart], { type: 'application/pdf' });
  const w = window as PickerWindow;
  if (w.showSaveFilePicker) {
    try {
      const h = await w.showSaveFilePicker({ suggestedName: name, types: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }] });
      const wr = await h.createWritable();
      await wr.write(blob); await wr.close();
      store.notify(T('Saved as PDF.'));
      return;
    } catch (e) { if ((e as { name?: string }).name === 'AbortError') return; }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 60_000);
};

/** Opens the PDF: the system viewer on the desktop, a new tab in the browser. */
export const openPdf = async (store: GridStore, o: PrintOptions): Promise<void> => {
  const tab = bridge()?.openPdf ? null : window.open('', '_blank');
  let data: Uint8Array;
  try { data = await makePdf(store, o); } catch (e) { tab?.close(); store.notify(T('The PDF could not be made: {p0}', { p0: e instanceof Error ? e.message : T('unknown error') })); return; }
  const b = bridge();
  if (b?.openPdf) { try { await b.openPdf(pdfName(store, o), data); } catch (e) { store.notify(e instanceof Error ? e.message : T('The PDF could not be opened.')); } return; }
  const url = URL.createObjectURL(new Blob([data as BlobPart], { type: 'application/pdf' }));
  if (tab) tab.location.href = url; else window.open(url, '_blank');
  setTimeout(() => URL.revokeObjectURL(url), 10 * 60_000);
};

/** Prints through the system dialog: each page as a vector picture in a hidden frame. */
export const printJob = async (store: GridStore, o: PrintOptions): Promise<void> => {
  store.commitEdit();
  const job = planPrint(store, o);
  if (!job.pages.length) { store.notify(T('There is nothing to print on this sheet.')); return; }
  const env = browserEnv(store);
  await loadFonts(store, job, env);
  const sizes = new Map<string, string>();
  const body = job.pages.map((pp) => {
    const key = `${Math.round(pp.pg.paperW * 100)}x${Math.round(pp.pg.paperH * 100)}`;
    if (!sizes.has(key)) sizes.set(key, `s${sizes.size}`);
    return `<div class="pg" style="page:${sizes.get(key)};width:${pp.pg.paperW}pt;height:${pp.pg.paperH}pt">${pageToSvg(store, pp, env)}</div>`;
  }).join('');
  const pageCss = [...sizes].map(([k, n]) => { const [w, h] = k.split('x').map((v) => +v / 100); return `@page ${n}{size:${w}pt ${h}pt;margin:0}`; }).join('');
  const frame = document.createElement('iframe');
  frame.setAttribute('aria-hidden', 'true');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0';
  document.body.appendChild(frame);
  const doc = frame.contentDocument!;
  doc.open();
  doc.write(`<!doctype html><html><head><meta charset="utf-8"><title>${job.title.replace(/[<&]/g, '')}</title></head><body></body></html>`);
  doc.close();
  // Same fonts as the app.
  for (const el of document.querySelectorAll('link[rel="stylesheet"], style')) doc.head.appendChild(doc.importNode(el, true));
  const st = doc.createElement('style');
  st.textContent = `${pageCss}@page{margin:0}html,body{margin:0;padding:0;background:#fff}.pg{break-after:page;overflow:hidden}.pg:last-child{break-after:auto}.pg svg{display:block;width:100%;height:100%}`;
  doc.head.appendChild(st);
  doc.body.innerHTML = body;
  try { await (doc as Document & { fonts?: FontFaceSet }).fonts?.ready; } catch { /* print anyway */ }
  await new Promise((r) => setTimeout(r, 150));
  frame.contentWindow!.focus();
  frame.contentWindow!.print();
  setTimeout(() => frame.remove(), 60_000);
};
