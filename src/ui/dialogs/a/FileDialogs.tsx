import { useEffect, useMemo, useRef, useState } from 'react';
import { FileUp, ImagePlus } from 'lucide-react';
import { readCsv } from '../../../formats/csv';
import { cellKey, internStyle, mergeStyle, rangeAddr, type Picture } from '../../../model/workbook';
import { MAX_COLS, MAX_ROWS, type Scalar } from '../../../engine/types';
import '../../charts/charts.css';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { BOLD, hasContent } from './util';

const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
const MAX_TEXT_BYTES = 64 * 1024 * 1024;
const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp', 'image/svg+xml'];

const FilePick = ({ accept, label, icon, file, onPick }: { accept: string; label: string; icon: 'image' | 'file'; file: string | null; onPick: (f: File) => void }) => {
  const input = useRef<HTMLInputElement>(null);
  const Icon = icon === 'image' ? ImagePlus : FileUp;
  return (
    <div className="filepick">
      <button type="button" autoFocus onClick={() => input.current?.click()}><Icon size={16} strokeWidth={1.9} /> {label}</button>
      <span className="muted">{file ?? 'No file chosen'}</span>
      <input ref={input} type="file" accept={accept} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) onPick(f); e.target.value = ''; }} />
    </div>
  );
};

// ---------- Picture ----------

interface PickedImage { name: string; data: Uint8Array; mime: string; url: string; w: number; h: number }

const loadImage = (url: string): Promise<{ w: number; h: number }> => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve({ w: img.naturalWidth || 300, h: img.naturalHeight || 200 });
  img.onerror = () => reject(new Error('bad image'));
  img.src = url;
});

export const PictureDialog = ({ store, close }: DialogProps) => {
  const [pic, setPic] = useState<PickedImage | null>(null);
  useEffect(() => () => { if (pic) URL.revokeObjectURL(pic.url); }, [pic]);

  const pick = async (f: File) => {
    if (f.size > MAX_IMAGE_BYTES) { store.notify('That picture is too large. Choose one under 20 MB.'); return; }
    const mime = IMAGE_TYPES.includes(f.type) ? f.type : '';
    if (!mime) { store.notify('Choose a PNG, JPEG, GIF, WebP, BMP or SVG picture.'); return; }
    try {
      const data = new Uint8Array(await f.arrayBuffer());
      const url = URL.createObjectURL(new Blob([data as BlobPart], { type: mime }));
      const { w, h } = await loadImage(url);
      setPic({ name: f.name, data, mime, url, w, h });
    } catch {
      store.notify("This picture couldn't be read.");
    }
  };

  const ok = () => {
    if (!pic) { store.notify('Choose a picture first.'); return false; }
    const scale = Math.min(1, 480 / pic.w);
    const { row, col } = store.sel.active;
    const p: Picture = { id: `pic-${Date.now().toString(36)}`, data: pic.data, mime: pic.mime, anchor: { row, col, dx: 0, dy: 0, width: Math.round(pic.w * scale), height: Math.round(pic.h * scale) } };
    store.change('Insert picture', 'sheet', () => { store.sheet.pictures.push(p); });
    return true;
  };

  return (
    <Dialog title="Insert picture" close={close} onOk={ok} okLabel="Insert">
      <FilePick accept="image/*" label="Choose picture" icon="image" file={pic?.name ?? null} onPick={(f) => void pick(f)} />
      {pic && <img className="imgprev" src={pic.url} alt="" />}
      <p className="muted small">The picture is placed at {rangeAddr({ r1: store.sel.active.row, c1: store.sel.active.col, r2: store.sel.active.row, c2: store.sel.active.col })}, up to 480 pixels wide.</p>
    </Dialog>
  );
};

// ---------- Text / CSV import ----------

const DELIMS: { value: string; label: string }[] = [
  { value: '', label: 'Detect automatically' }, { value: ',', label: 'Comma' }, { value: ';', label: 'Semicolon' }, { value: '\t', label: 'Tab' }, { value: '|', label: 'Vertical bar' },
];

export const ImportCsvDialog = ({ store, close }: DialogProps) => {
  const [file, setFile] = useState<{ name: string; text: string } | null>(null);
  const [delim, setDelim] = useState('');
  const [header, setHeader] = useState(true);

  const parsed = useMemo(() => (file ? readCsv(file.text, delim ? { delimiter: delim } : {}) : null), [file, delim]);

  const pick = async (f: File) => {
    if (f.size > MAX_TEXT_BYTES) { store.notify('That file is too large to import.'); return; }
    try {
      const text = await f.text();
      setFile({ name: f.name, text });
      if (/\.tsv$/i.test(f.name)) setDelim('\t');
    } catch {
      store.notify("This file couldn't be read.");
    }
  };

  const ok = () => {
    if (!parsed || !parsed.rows.length) { store.notify('Choose a text or CSV file with some data.'); return false; }
    const { row, col } = store.sel.active;
    const rows = parsed.rows.slice(0, MAX_ROWS - row);
    const width = Math.min(MAX_COLS - col, Math.max(...rows.map((r) => r.length)));
    const area = { r1: row, c1: col, r2: row + rows.length - 1, c2: col + width - 1 };
    if (hasContent(store, store.sheetIndex, area) && !confirm('Replace the data in the cells where the file goes?')) return false;
    const sheet = store.sheetIndex;
    store.change('Import text', [{ area }], () => {
      rows.forEach((line, i) => {
        for (let j = 0; j < width; j++) {
          const v: Scalar = line[j] ?? null;
          store.engine.setValue(sheet, row + i, col + j, v);
          const fmt = parsed.formats[i]?.[j];
          const bold = header && i === 0;
          if (!fmt && !bold) continue;
          const k = cellKey(row + i, col + j);
          const base = store.style(row + i, col + j);
          const id = internStyle(store.wb, mergeStyle(base, { ...(fmt ? { numFmt: fmt } : {}), ...(bold ? BOLD : {}) }));
          const c = store.sheet.cells.get(k);
          if (c) c.s = id; else store.sheet.cells.set(k, { v: null, s: id });
        }
      });
    });
    store.selectRange(area, { row, col });
    if (parsed.rows.length > rows.length) store.notify('Some rows did not fit on the sheet and were left out.');
    return true;
  };

  const preview = parsed?.rows.slice(0, 12) ?? [];
  const cols = Math.min(20, Math.max(0, ...preview.map((r) => r.length)));
  return (
    <Dialog title="Import text or CSV" close={close} wide onOk={ok} okLabel="Import">
      <FilePick accept=".csv,.tsv,.txt,text/csv,text/plain,text/tab-separated-values" label="Choose file" icon="file" file={file?.name ?? null} onPick={(f) => void pick(f)} />
      <div className="grid2">
        <Field label="Separator">
          <select value={delim} onChange={(e) => setDelim(e.target.value)}>{DELIMS.map((d) => <option key={d.label} value={d.value}>{d.label}</option>)}</select>
        </Field>
        <label className="checkrow" style={{ alignSelf: 'end', minHeight: 32 }}><input type="checkbox" checked={header} onChange={(e) => setHeader(e.target.checked)} /> First row is headers</label>
      </div>
      {parsed && (
        <>
          <div className="csvprev">
            <table><tbody>
              {preview.map((r, i) => <tr key={i} className={header && i === 0 ? 'hdr' : ''}>{Array.from({ length: cols }, (_, j) => <td key={j}>{cellText(r[j] ?? null)}</td>)}</tr>)}
            </tbody></table>
          </div>
          <p className="muted small">{parsed.rows.length.toLocaleString()} rows. Goes into the sheet at {rangeAddr({ r1: store.sel.active.row, c1: store.sel.active.col, r2: store.sel.active.row, c2: store.sel.active.col })}.</p>
        </>
      )}
    </Dialog>
  );
};

const cellText = (v: Scalar): string => (v === null ? '' : typeof v === 'object' ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
