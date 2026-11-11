// Share links: the whole workbook, deflated, in the link's #hash (never sent to any server). Larger workbooks are
// shared as a file instead.
import type { GridStore } from '../../app/store';
import { workbookFromBytes } from '../../app/files';
import { writeXlsx } from '../../formats/xlsx';
import { packBytes, unpackBytes } from './b64';
import { T } from '../../ui/i18n';

/** Longest hash we put in a link. Browsers accept links of a few megabytes; chat apps often cut them shorter. */
export const LINK_LIMIT = 2 * 1024 * 1024;
const MAX_FILE = 64 * 1024 * 1024;
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const appUrl = (): string => (typeof location !== 'undefined' && /^https?:$/.test(location.protocol) ? `${location.origin}${location.pathname}` : 'https://clerdoapp.github.io/clerdogrid/');

export const fileBytes = (store: GridStore): Uint8Array => {
  store.commitEdit();
  return writeXlsx(store.wb);
};

export interface ShareLink { url: string; size: number }

/** Builds the link, or returns null with the encoded size when the workbook is too big for one. */
export const makeShareLink = (store: GridStore, readOnly: boolean): ShareLink | { tooBig: number } => {
  const packed = packBytes(fileBytes(store));
  if (packed.length > LINK_LIMIT) return { tooBig: packed.length };
  const name = encodeURIComponent((store.file.name || 'Book1').slice(0, 120));
  return { url: `${appUrl()}#wb=${packed}${readOnly ? '&ro=1' : ''}&n=${name}`, size: packed.length };
};

export interface SharedBook { bytes: Uint8Array; readOnly: boolean; name: string }

/** Reads "#wb=...&ro=1&n=..." from a link. Returns null when the hash isn't a share link; throws when it is broken. */
export const parseShareHash = (hash: string): SharedBook | null => {
  const h = hash.replace(/^#/, '');
  if (!h.startsWith('wb=')) return null;
  const params = new Map<string, string>();
  for (const part of h.split('&')) { const i = part.indexOf('='); if (i > 0) params.set(part.slice(0, i), part.slice(i + 1)); }
  const data = params.get('wb') ?? '';
  if (!/^[A-Za-z0-9_-]+$/.test(data) || data.length > LINK_LIMIT * 2) throw new Error(T('This share link isn\'t complete.'));
  let name = T('Shared workbook');
  try { name = decodeURIComponent(params.get('n') ?? '').replace(/[\\/:*?"<>|]/g, '').slice(0, 120) || name; } catch { /* keep default */ }
  let bytes: Uint8Array;
  try { bytes = unpackBytes(data, MAX_FILE); } catch { throw new Error(T('This share link isn\'t complete.')); }
  return { bytes, readOnly: params.get('ro') === '1', name };
};

/** Opens a workbook from the page's share link, through the normal (untrusted-input) file reader. */
export const openShareHash = (store: GridStore, hash: string): boolean => {
  let shared: SharedBook | null;
  try { shared = parseShareHash(hash); } catch (e) { store.notify(e instanceof Error ? e.message : T('This share link couldn\'t be opened.')); return false; }
  if (!shared) return false;
  try {
    const wb = workbookFromBytes(`${shared.name}.xlsx`, shared.bytes);
    store.load(wb, shared.name);
    store.readOnly = shared.readOnly;
    // Kept as a draft like any unsaved workbook, so a reload doesn't lose edits.
    store.file.dirty = !shared.readOnly;
    store.notify(shared.readOnly ? T('Opened a read-only copy from a link.') : T('Opened a copy from a link. Save it to keep it.'));
    return true;
  } catch (e) {
    store.notify(e instanceof Error ? e.message : T('This share link couldn\'t be opened.'));
    return false;
  }
};

/** Shares the .xlsx with the system share sheet when it can take files, else downloads it. */
export const shareFile = async (store: GridStore): Promise<'shared' | 'downloaded' | 'cancelled'> => {
  const name = `${store.file.name || 'Book1'}.xlsx`;
  const bytes = fileBytes(store);
  const nav = navigator as Navigator & { canShare?: (d: { files: File[] }) => boolean };
  if (typeof File !== 'undefined' && nav.share && nav.canShare) {
    const file = new File([bytes as BlobPart], name, { type: XLSX_TYPE });
    if (nav.canShare({ files: [file] })) {
      try { await nav.share({ files: [file], title: store.file.name }); return 'shared'; } catch (e) { if ((e as { name?: string }).name === 'AbortError') return 'cancelled'; }
    }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: XLSX_TYPE }));
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30_000);
  return 'downloaded';
};
