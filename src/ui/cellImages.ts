// Draws pictures returned by =IMAGE() inside their cells. Images load in the background (only when the
// person allowed functions to use the internet) and are cached; the grid redraws when one arrives.
import { cellImage, fetchImageViaRelay, onlineAllowed, onOnlineChanged } from '../engine/online';
import type { GridStore } from '../app/store';

interface Loaded { img: HTMLImageElement; ok: boolean | null }
const loaded = new Map<string, Loaded>();
let redraw: (() => void) | null = null;
// Turning internet access on or off starts again, so pictures refused while off can load.
onOnlineChanged(() => loaded.clear());

/** Set by the app so a finished download repaints the grid. */
export const setImageRedraw = (f: () => void): void => { redraw = f; };

const get = (url: string): Loaded | null => {
  if (typeof Image === 'undefined') return null;
  let l = loaded.get(url);
  if (!l) {
    if (loaded.size > 500) loaded.delete(loaded.keys().next().value!);
    const img = new Image();
    img.decoding = 'async';
    img.referrerPolicy = 'no-referrer';
    l = { img, ok: null };
    const entry = l;
    img.onload = () => { entry.ok = true; redraw?.(); };
    img.onerror = () => { entry.ok = false; redraw?.(); };
    loaded.set(url, l);
    // Desktop: bytes come through the main process and show from a local blob address.
    fetchImageViaRelay(url).then((r) => {
      if (r) img.src = URL.createObjectURL(new Blob([r.data as BlobPart], { type: r.mime }));
      else if (onlineAllowed()) img.src = url;
      else { entry.ok = false; redraw?.(); }
    }, () => { entry.ok = false; redraw?.(); });
  }
  return l;
};

/** Draws the cell's image; returns true when the picture took the place of the cell text. */
export const drawCellImage = (ctx: CanvasRenderingContext2D, store: GridStore, rect: { x: number; y: number; w: number; h: number }, row: number, col: number, zoom: number): boolean => {
  const ci = cellImage(store.sheetIndex, row, col);
  if (!ci) return false;
  const anchor = store.cell(ci.anchor.row, ci.anchor.col, ci.anchor.sheet);
  if (!anchor?.f || !/IMAGE\s*\(/i.test(anchor.f)) return false;
  const l = get(ci.url);
  if (!l?.ok) return false;
  const iw = l.img.naturalWidth || 1, ih = l.img.naturalHeight || 1;
  const pad = 1;
  const bw = Math.max(1, rect.w - 2 * pad), bh = Math.max(1, rect.h - 2 * pad);
  let w: number, h: number;
  switch (ci.sizing) {
    case 1: w = bw; h = bh; break;
    case 2: w = iw * zoom; h = ih * zoom; break;
    case 3: w = (ci.width ?? iw) * zoom; h = (ci.height ?? ih) * zoom; break;
    default: { const s = Math.min(bw / iw, bh / ih); w = iw * s; h = ih * s; }
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.w, rect.h);
  ctx.clip();
  ctx.drawImage(l.img, rect.x + pad + (bw - w) / 2, rect.y + pad + (bh - h) / 2, w, h);
  ctx.restore();
  return true;
};
