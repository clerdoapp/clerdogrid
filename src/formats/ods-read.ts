// OpenDocument spreadsheet reader (.ods / .ots, ODF 1.3).
import { unzipSync } from 'fflate';
import {
  cellKey, internStyle, keyCol, keyRow, newSheet, type AutoFilter, type BorderSide, type BorderStyle, type Cell, type CellRange, type CellStyle,
  type CfOperator, type CfRule, type DataValidation, type FilterColumn, type RichRun, type Sheet, type Workbook,
} from '../model/workbook';
import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../engine/types';
import { toSerial } from '../engine/date';
import { a, child, children, elems, findDeep, odfText, parseOdf, type ONode } from './ods-xml';
import { fromOdfFormula, gridRef, splitAddresses } from './ods-formula';
import { readDataStyle } from './ods-numfmt';
import { readChartObject } from './ods-chart';
import { shownColWidth } from './xlsx-chart';
import { queriesFromXml } from './queries-xml';
import { decryptStream, type EntryEncryption } from './ods-crypto';
import { fromBase64, sha1, sha256 } from './crypto';
import { PasswordRequired } from './biff8-read';

const MAX_PART = 1 << 30;
const MAX_CELLS = 20_000_000;
const SAFE_LINK = /^(https:|mailto:|#)/i;
const MAX_ROW = 1_048_576, MAX_COL = 16_384;

/** A length ("2.5cm", "0.3in", "12pt", "4mm", "10px") in points. */
export const toPt = (s: string | undefined): number | undefined => {
  if (!s) return undefined;
  const m = /^(-?[\d.]+)\s*(cm|mm|in|inch|pt|pc|px)?$/.exec(s.trim());
  if (!m) return undefined;
  const v = +m[1]!;
  switch (m[2]) {
    case 'cm': return (v / 2.54) * 72;
    case 'mm': return (v / 25.4) * 72;
    case 'in': case 'inch': return v * 72;
    case 'pc': return v * 12;
    case 'px': return v * 0.75;
    default: return v;
  }
};

const ptToChars = (pt: number): number => Math.max(0, Math.round(((pt * 96) / 72 - 5) / 7 * 100) / 100);

const isoToSerial = (s: string, date1904: boolean): number | undefined => {
  const m = /^(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s);
  if (!m) return undefined;
  const day = toSerial(+m[1]!, +m[2]!, +m[3]!, date1904);
  return day + ((+(m[4] ?? 0)) * 3600 + (+(m[5] ?? 0)) * 60 + +(m[6] ?? 0)) / 86400;
};
const durationToDays = (s: string): number | undefined => {
  const m = /^(-)?P(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(s.trim());
  if (!m) return undefined;
  const v = (+(m[2] ?? 0)) + ((+(m[3] ?? 0)) * 3600 + (+(m[4] ?? 0)) * 60 + +(m[5] ?? 0)) / 86400;
  return m[1] ? -v : v;
};

const border = (s: string | undefined): BorderSide | undefined => {
  if (!s || s === 'none' || s === 'hidden') return undefined;
  const parts = s.trim().split(/\s+/);
  const w = toPt(parts.find((p) => /^[\d.]+(pt|cm|mm|in|px)$/.test(p))) ?? 0.75;
  const kind = (parts.find((p) => /^(solid|dotted|dashed|double(-thin)?|thin-double|dash-dot|dash-dot-dot|fine-dashed|wave|groove|ridge|inset|outset)$/.test(p)) ?? 'solid').replace(/^(double-thin|thin-double)$/, 'double');
  const color = parts.find((p) => /^#[0-9a-f]{6}$/i.test(p));
  let style: BorderStyle;
  if (kind === 'double') style = 'double';
  else if (kind === 'dotted') style = w < 0.3 ? 'hair' : 'dotted';
  else if (kind === 'dashed') style = w > 1.1 ? 'mediumDashed' : 'dashed';
  else if (kind === 'fine-dashed') style = 'hair';
  else if (kind === 'dash-dot') style = w > 1.1 ? 'mediumDashDot' : 'dashDot';
  else if (kind === 'dash-dot-dot') style = w > 1.1 ? 'mediumDashDotDot' : 'dashDotDot';
  else style = w < 0.3 ? 'hair' : w <= 1.1 ? 'thin' : w <= 2 ? 'medium' : 'thick';
  const c = color?.toUpperCase();
  return c && c !== '#000000' ? { style, color: c } : { style };
};

/** Cells listed one by one → covering rectangles (runs per row, then joined downwards). */
export const toRanges = (keys: number[]): CellRange[] => {
  const byRow = new Map<number, number[]>();
  for (const k of keys) { const r = keyRow(k); let l = byRow.get(r); if (!l) byRow.set(r, (l = [])); l.push(keyCol(k)); }
  const runs: CellRange[] = [];
  for (const [r, cols] of [...byRow].sort((x, y) => x[0] - y[0])) {
    cols.sort((x, y) => x - y);
    let s = cols[0]!, p = s;
    for (const c of [...cols.slice(1), -2]) {
      if (c === p + 1) { p = c; continue; }
      runs.push({ r1: r, r2: r, c1: s, c2: p });
      s = p = c;
    }
  }
  const out: CellRange[] = [];
  for (const run of runs) {
    const prev = out.find((o) => o.c1 === run.c1 && o.c2 === run.c2 && o.r2 === run.r1 - 1);
    if (prev) prev.r2 = run.r2; else out.push({ ...run });
  }
  return out;
};

/** ODF cell-range address ("Sheet1.A1:Sheet1.B2", "$'My sheet'.$A$1:.$B$2") → sheet name and range. */
export const parseRangeAddress = (addr: string | undefined): { sheet?: string; range: CellRange } | undefined => {
  if (!addr) return undefined;
  const ref = gridRef(splitAddresses(addr)[0] ?? '');
  if (ref === '#REF!') return undefined;
  const bang = ref.lastIndexOf('!');
  let sheet = bang >= 0 ? ref.slice(0, bang) : undefined;
  if (sheet?.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  if (sheet?.includes(':')) sheet = sheet.split(':')[0];
  const body = ref.slice(bang + 1).replace(/\$/g, '');
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(body);
  if (!m) {
    const cols = /^([A-Z]+):([A-Z]+)$/.exec(body);
    if (cols) return { sheet, range: { r1: 0, r2: MAX_ROW - 1, c1: colNum(cols[1]!), c2: colNum(cols[2]!) } };
    const rows = /^(\d+):(\d+)$/.exec(body);
    if (rows) return { sheet, range: { r1: +rows[1]! - 1, r2: +rows[2]! - 1, c1: 0, c2: MAX_COL - 1 } };
    return undefined;
  }
  const c1 = colNum(m[1]!), r1 = +m[2]! - 1;
  const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
  return { sheet, range: { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) } };
};
const colNum = (s: string): number => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** Splits "a;b" / "a,b" at the top level (outside quotes, brackets and parentheses). */
const splitArgs = (s: string): string[] => {
  const out: string[] = [];
  let depth = 0, buf = '', q = '';
  for (const ch of s) {
    if (q) { buf += ch; if (ch === q) q = ''; continue; }
    if (ch === '"' || ch === "'") { q = ch; buf += ch; continue; }
    if (ch === '(' || ch === '[' || ch === '{') depth++;
    if (ch === ')' || ch === ']' || ch === '}') depth--;
    if ((ch === ';' || ch === ',') && depth === 0) { out.push(buf); buf = ''; continue; }
    buf += ch;
  }
  out.push(buf);
  return out.map((x) => x.trim());
};
/** "name(args)" → [name, inner]. */
const callOf = (s: string): [string, string] | undefined => {
  const m = /^([a-z-]+)\((.*)\)$/s.exec(s.trim());
  return m ? [m[1]!, m[2]!] : undefined;
};
const expr = (s: string): string | undefined => fromOdfFormula(`of:=${s}`);

const PICTURE_MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };

/** Reads an .ods or .ots file; an encrypted one needs its password (PasswordRequired otherwise). */
export const readOds = (data: Uint8Array, opts: { password?: string } = {}): Workbook => {
  let total = 0;
  const files = unzipSync(data, {
    filter: (f) => { total += f.originalSize; if (f.originalSize > MAX_PART || total > 2 ** 31) throw new Error('The file is too large to open.'); return true; },
  });
  decryptPackage(files, opts.password);
  const text = (p: string): string | undefined => (files[p] ? new TextDecoder().decode(files[p]) : undefined);
  const contentText = text('content.xml');
  if (!contentText) throw new Error("This file isn't a spreadsheet ClerdoGrid can read.");
  const content = parseOdf(contentText);
  const stylesDoc = text('styles.xml') ? parseOdf(text('styles.xml')!) : undefined;
  const settings = text('settings.xml') ? parseOdf(text('settings.xml')!) : undefined;
  const meta = text('meta.xml') ? parseOdf(text('meta.xml')!) : undefined;

  const wb: Workbook = { sheets: [], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  const spreadsheet = findDeep(content, 'office:spreadsheet');
  if (!spreadsheet) throw new Error("This file isn't a spreadsheet ClerdoGrid can read.");

  // ---------- Styles ----------
  const fontFaces = new Map<string, string>();
  for (const doc of [stylesDoc, content]) {
    for (const f of children(findDeep(doc, 'office:font-face-decls'), 'style:font-face')) {
      const fam = (a(f, 'svg:font-family') ?? a(f, 'style:name') ?? '').replace(/^['"]|['"]$/g, '');
      fontFaces.set(a(f, 'style:name') ?? '', fam);
    }
  }
  const styleNodes = new Map<string, ONode>();
  const dataStyles = new Map<string, ONode>();
  let defaultCell: ONode | undefined;
  for (const doc of [stylesDoc, content]) {
    for (const sec of ['office:styles', 'office:automatic-styles']) {
      for (const s of elems(findDeep(doc, sec))) {
        if (s.name === 'style:style') styleNodes.set(`${a(s, 'style:family')}/${a(s, 'style:name')}`, s);
        else if (s.name.startsWith('number:') && s.name.endsWith('-style')) dataStyles.set(a(s, 'style:name') ?? '', s);
        else if (s.name === 'style:default-style' && a(s, 'style:family') === 'table-cell') defaultCell = s;
      }
    }
  }
  const dataCode = new Map<string, string | undefined>();
  const codeOf = (name: string | undefined): string | undefined => {
    if (!name) return undefined;
    if (!dataCode.has(name)) { const n = dataStyles.get(name); dataCode.set(name, n ? readDataStyle(n, dataStyles) : undefined); }
    return dataCode.get(name);
  };

  const applyProps = (st: CellStyle, node: ONode): void => {
    const tc = child(node, 'style:table-cell-properties');
    const pp = child(node, 'style:paragraph-properties');
    const tp = child(node, 'style:text-properties');
    if (tp) {
      const f = { ...st.font };
      const fn = a(tp, 'style:font-name');
      const fam = (fn ? fontFaces.get(fn) ?? fn : undefined) ?? a(tp, 'fo:font-family')?.replace(/^['"]|['"]$/g, '');
      if (fam) f.name = fam;
      const size = toPt(a(tp, 'fo:font-size'));
      if (size) f.size = Math.round(size * 2) / 2;
      const w = a(tp, 'fo:font-weight');
      if (w) f.bold = w === 'bold' || +w >= 600 || undefined;
      const it = a(tp, 'fo:font-style');
      if (it) f.italic = it === 'italic' || it === 'oblique' || undefined;
      const ul = a(tp, 'style:text-underline-style');
      if (ul) f.underline = ul === 'none' ? undefined : a(tp, 'style:text-underline-type') === 'double' ? 'double' : 'single';
      const lt = a(tp, 'style:text-line-through-style');
      if (lt) f.strike = lt !== 'none' || undefined;
      const c = a(tp, 'fo:color');
      if (c && /^#[0-9a-f]{6}$/i.test(c)) f.color = c.toUpperCase() === '#000000' ? undefined : c.toUpperCase();
      const pos = a(tp, 'style:text-position');
      if (pos) f.vertAlign = /^super|^\d/.test(pos) && !/^0%/.test(pos) ? 'superscript' : /^sub|^-/.test(pos) ? 'subscript' : undefined;
      st.font = f;
    }
    if (tc) {
      const bg = a(tc, 'fo:background-color');
      if (bg && /^#[0-9a-f]{6}$/i.test(bg)) st.fill = { color: bg.toUpperCase() };
      else if (bg === 'transparent') delete st.fill;
      const all = a(tc, 'fo:border');
      const b = { ...st.border };
      const set = (k: 'left' | 'right' | 'top' | 'bottom', v: string | undefined) => { if (v !== undefined) { const x = border(v); if (x) b[k] = x; else delete b[k]; } };
      if (all !== undefined) for (const k of ['left', 'right', 'top', 'bottom'] as const) set(k, all);
      set('left', a(tc, 'fo:border-left')); set('right', a(tc, 'fo:border-right')); set('top', a(tc, 'fo:border-top')); set('bottom', a(tc, 'fo:border-bottom'));
      const d1 = border(a(tc, 'style:diagonal-tl-br')), d2 = border(a(tc, 'style:diagonal-bl-tr'));
      if (d1) b.diagonalDown = d1;
      if (d2) b.diagonalUp = d2;
      st.border = b;
      const al = { ...st.align };
      const va = a(tc, 'style:vertical-align');
      if (va) al.v = va === 'top' ? 'top' : va === 'middle' ? 'center' : va === 'justify' ? 'justify' : undefined;
      const wrap = a(tc, 'fo:wrap-option');
      if (wrap) al.wrap = wrap === 'wrap' || undefined;
      const rot = a(tc, 'style:rotation-angle');
      if (rot !== undefined) {
        let deg = parseFloat(rot);
        if (/rad$/.test(rot)) deg = (deg * 180) / Math.PI; else if (/grad$/.test(rot)) deg = deg * 0.9;
        deg = ((Math.round(deg) % 360) + 360) % 360;
        al.rotation = deg === 0 ? undefined : deg <= 90 ? deg : deg >= 270 ? deg - 360 : undefined;
      }
      if (a(tc, 'style:direction') === 'ttb') al.rotation = 255;
      const shrink = a(tc, 'style:shrink-to-fit');
      if (shrink) al.shrink = shrink === 'true' || undefined;
      st.align = al;
      const prot = a(tc, 'style:cell-protect');
      if (prot) {
        const locked = /protected/.test(prot);
        const hidden = /hidden/.test(prot);
        st.protection = { ...(locked ? {} : { locked: false }), ...(hidden ? { hidden: true } : {}) };
      }
    }
    if (pp) {
      const al = { ...st.align };
      const ta = a(pp, 'fo:text-align');
      const src = a(tc, 'style:text-align-source');
      if (ta && src !== 'value-type') al.h = ta === 'start' || ta === 'left' ? 'left' : ta === 'end' || ta === 'right' ? 'right' : ta === 'center' ? 'center' : ta === 'justify' ? 'justify' : undefined;
      if (src === 'value-type') al.h = undefined;
      const ml = toPt(a(pp, 'fo:margin-left'));
      if (ml !== undefined) al.indent = ml >= 5 ? Math.min(15, Math.round(ml / 10)) || undefined : undefined;
      if (a(pp, 'style:writing-mode') === 'rl-tb') al.rtl = true;
      st.align = al;
    }
    const ds = a(node, 'style:data-style-name');
    if (ds) { const code = codeOf(ds); if (code && code !== 'General') st.numFmt = code; else delete st.numFmt; }
  };

  const clean = (st: CellStyle): CellStyle => {
    const out: CellStyle = {};
    const f = Object.fromEntries(Object.entries(st.font ?? {}).filter(([, v]) => v !== undefined && v !== false)) as NonNullable<CellStyle['font']>;
    if (f.name === 'Calibri') delete f.name;
    if (f.size === 11) delete f.size;
    if (Object.keys(f).length) out.font = f;
    if (st.fill?.color) out.fill = st.fill;
    if (st.border && Object.keys(st.border).length) out.border = st.border;
    const al = Object.fromEntries(Object.entries(st.align ?? {}).filter(([, v]) => v !== undefined && v !== false)) as NonNullable<CellStyle['align']>;
    if (Object.keys(al).length) out.align = al;
    if (st.numFmt) out.numFmt = st.numFmt;
    if (st.protection && Object.keys(st.protection).length) out.protection = st.protection;
    return out;
  };

  const resolved = new Map<string, CellStyle>();
  const resolveCell = (name: string | undefined, depth = 0): CellStyle => {
    const key = name ?? '';
    const hit = resolved.get(key);
    if (hit) return hit;
    let st: CellStyle = {};
    if (!name) { if (defaultCell) applyProps(st, defaultCell); }
    else {
      const node = styleNodes.get(`table-cell/${name}`);
      const parent = node ? a(node, 'style:parent-style-name') : undefined;
      st = structuredClone(depth < 8 ? resolveCell(parent ?? (name === 'Default' ? undefined : 'Default'), depth + 1) : {});
      if (node) applyProps(st, node);
    }
    resolved.set(key, st);
    return st;
  };
  const styleIds = new Map<string, number>();
  const styleId = (name: string | undefined): number => {
    const key = name ?? '';
    let id = styleIds.get(key);
    if (id === undefined) { id = internStyle(wb, clean(resolveCell(name ?? 'Default'))); styleIds.set(key, id); }
    return id;
  };
  const textStyle = (name: string | undefined): RichRun['font'] | undefined => {
    const node = styleNodes.get(`text/${name}`);
    if (!node) return undefined;
    const st: CellStyle = {};
    applyProps(st, node);
    const f = clean(st).font;
    return f && Object.keys(f).length ? f : undefined;
  };

  if (a(spreadsheet, 'table:structure-protected') === 'true') {
    const hash = legacyHash(spreadsheet);
    wb.protection = { structure: true, ...(hash ? { password: hash } : {}) };
  }
  const calc = child(spreadsheet, 'table:calculation-settings');
  const nullDate = a(child(calc, 'table:null-date'), 'table:date-value');
  if (nullDate?.startsWith('1904-01-01')) wb.date1904 = true;
  const iter = child(calc, 'table:iteration');
  if (a(iter, 'table:status') === 'enable') wb.calc = { iterate: true, maxIterations: +(a(iter, 'table:steps') ?? 100) || 100, maxChange: +(a(iter, 'table:minimum-difference') ?? 0.001) || 0.001 };

  // content validations by name
  const validations = new Map<string, Omit<DataValidation, 'ranges'>>();
  for (const v of children(child(spreadsheet, 'table:content-validations'), 'table:content-validation')) {
    const dv = readValidation(v);
    if (dv) validations.set(a(v, 'table:name') ?? '', dv);
  }

  // ---------- Tables ----------
  const tableNodes = children(spreadsheet, 'table:table');
  let cellCount = 0;
  const pending: { sheet: Sheet; node: ONode }[] = [];
  const sheetNames = new Set<string>();
  tableNodes.forEach((t, ti) => {
    let name = (a(t, 'table:name') ?? `Sheet${ti + 1}`).slice(0, 31);
    while (sheetNames.has(name.toUpperCase())) name = `${name.slice(0, 28)}_${ti}`;
    sheetNames.add(name.toUpperCase());
    const sheet = newSheet(name);
    wb.sheets.push(sheet);
    pending.push({ sheet, node: t });
  });

  pending.forEach(({ sheet, node: t }, sheetIdx) => {
    const ts = styleNodes.get(`table/${a(t, 'table:style-name')}`);
    const tp = child(ts, 'style:table-properties');
    if (a(tp, 'table:display') === 'false') sheet.state = 'hidden';
    const tab = a(tp, 'tableooo:tab-color') ?? a(tp, 'table:tab-color');
    if (tab && /^#[0-9a-f]{6}$/i.test(tab)) sheet.tabColor = tab.toUpperCase();
    if (a(tp, 'style:writing-mode') === 'rl-tb') sheet.view.rtl = true;
    if (a(t, 'table:protected') === 'true') {
      const hash = a(t, 'clerdo:protection-hash') ?? legacyHash(t);
      sheet.protection = { sheet: true, ...(hash && /^[0-9A-F]{4}$/i.test(hash) ? { password: hash.toUpperCase() } : {}) };
    }
    const pr = a(t, 'table:print-ranges');
    if (pr) {
      const parts = splitAddresses(pr).map((p) => parseRangeAddress(p)).filter((x): x is NonNullable<typeof x> => !!x);
      if (parts.length) sheet.pageSetup = { ...sheet.pageSetup, printArea: parts.map((p) => rangeText(p.range)).join(',') };
    }
    readPageLayout(sheet, a(ts, 'style:master-page-name'));

    // columns
    let col = 0;
    const colStyle: (string | undefined)[] = [];
    const walkCols = (n: ONode): void => {
      for (const c of elems(n)) {
        if (c.name === 'table:table-column') {
          const rep = Math.min(MAX_COL - col, Math.max(1, +(a(c, 'table:number-columns-repeated') ?? 1) || 1));
          const cs = styleNodes.get(`table-column/${a(c, 'table:style-name')}`);
          const w = toPt(a(child(cs, 'style:table-column-properties'), 'style:column-width'));
          const hidden = a(c, 'table:visibility') === 'collapse';
          const dcs = a(c, 'table:default-cell-style-name');
          for (let i = 0; i < rep && col < MAX_COL; i++, col++) {
            colStyle[col] = dcs;
            if (rep > 1024 && !hidden && dcs === undefined) continue;
            const info: NonNullable<ReturnType<Sheet['cols']['get']>> = {};
            if (w !== undefined) info.width = ptToChars(w);
            if (hidden) info.hidden = true;
            if (dcs && dcs !== 'Default') { const s = styleId(dcs); if (s) info.s = s; }
            if (Object.keys(info).length) sheet.cols.set(col, info);
          }
          if (rep > 1024) col = Math.min(MAX_COL, col);
        } else if (/^table:table-(column-group|header-columns|columns)$/.test(c.name)) walkCols(c);
      }
    };
    walkCols(t);
    // most common width becomes the default
    const freq = new Map<number, number>();
    for (const c of sheet.cols.values()) if (c.width !== undefined && !c.hidden && !c.s) freq.set(c.width, (freq.get(c.width) ?? 0) + 1);
    const common = [...freq].sort((x, y) => y[1] - x[1])[0];
    if (common && common[1] > 32) {
      sheet.defaultColWidth = common[0];
      for (const [k, c] of [...sheet.cols]) if (c.width === common[0] && !c.hidden && !c.s) sheet.cols.delete(k);
    }

    // rows
    let row = 0;
    const validationCells = new Map<string, number[]>();
    const mapCells = new Map<string, number[]>();
    const headerRows: [number, number][] = [];
    const walkRows = (n: ONode, header: boolean): void => {
      for (const r of elems(n)) {
        if (r.name === 'table:table-row') {
          const rep = Math.min(MAX_ROW - row, Math.max(1, +(a(r, 'table:number-rows-repeated') ?? 1) || 1));
          const rs = styleNodes.get(`table-row/${a(r, 'table:style-name')}`);
          const rp = child(rs, 'style:table-row-properties');
          const h = toPt(a(rp, 'style:row-height'));
          const custom = a(rp, 'style:use-optimal-row-height') === 'false';
          const vis = a(r, 'table:visibility');
          const cells = readRowCells(r);
          const empty = !cells.length;
          const start = row;
          for (let i = 0; i < rep && row < MAX_ROW; i++, row++) {
            if (empty && rep > 1000 && vis !== 'collapse' && vis !== 'filter') { row = start + rep; break; }
            const info: NonNullable<ReturnType<Sheet['rows']['get']>> = {};
            if (h !== undefined && (custom || Math.abs(h - sheet.defaultRowHeight) > 0.6)) info.height = Math.round(h * 4) / 4;
            if (vis === 'collapse') info.hidden = true;
            if (vis === 'filter') { info.hidden = true; info.filtered = true; }
            if (Object.keys(info).length) sheet.rows.set(row, info);
            for (const c of cells) placeCell(c, row);
            if (cellCount > MAX_CELLS) throw new Error('This file has too many cells to open.');
          }
          if (header) headerRows.push([start, row - 1]);
        } else if (r.name === 'table:table-header-rows') walkRows(r, true);
        else if (/^table:table-(row-group|rows)$/.test(r.name)) walkRows(r, false);
      }
    };

    interface Raw { col: number; rep: number; node: ONode; covered: boolean }
    const readRowCells = (r: ONode): Raw[] => {
      const out: Raw[] = [];
      let c = 0;
      for (const k of elems(r)) {
        if (k.name !== 'table:table-cell' && k.name !== 'table:covered-table-cell') continue;
        const rep = Math.max(1, +(a(k, 'table:number-columns-repeated') ?? 1) || 1);
        const meaningful = elems(k).length > 0 || a(k, 'office:value-type') !== undefined || a(k, 'table:formula') !== undefined || a(k, 'table:number-columns-spanned') !== undefined
          || a(k, 'table:content-validation-name') !== undefined || (a(k, 'table:style-name') !== undefined && rep <= 1024);
        if (meaningful && c < MAX_COL) out.push({ col: c, rep: Math.min(rep, MAX_COL - c), node: k, covered: k.name === 'table:covered-table-cell' });
        c += rep;
        if (c >= MAX_COL) break;
      }
      return out;
    };

    const placeCell = (raw: Raw, r: number): void => {
      const k = raw.node;
      for (let i = 0; i < raw.rep; i++) {
        const c = raw.col + i;
        const cell = readCell(k, r, c);
        const vn = a(k, 'table:content-validation-name');
        if (vn) { let l = validationCells.get(vn); if (!l) validationCells.set(vn, (l = [])); l.push(cellKey(r, c)); }
        const sn = a(k, 'table:style-name');
        if (sn && child(styleNodes.get(`table-cell/${sn}`), 'style:map')) { let l = mapCells.get(sn); if (!l) mapCells.set(sn, (l = [])); l.push(cellKey(r, c)); }
        if (cell) { sheet.cells.set(cellKey(r, c), cell); cellCount++; }
        if (i === 0) {
          const cs = +(a(k, 'table:number-columns-spanned') ?? 1) || 1, rs = +(a(k, 'table:number-rows-spanned') ?? 1) || 1;
          if (cs > 1 || rs > 1) sheet.merges.push({ r1: r, c1: c, r2: Math.min(MAX_ROW - 1, r + rs - 1), c2: Math.min(MAX_COL - 1, c + cs - 1) });
          for (const f of children(k, 'draw:frame')) readFrame(f, sheet, sheetIdx, { row: r, col: c });
        }
      }
    };

    const readCell = (k: ONode, r: number, c: number): Cell | undefined => {
      const type = a(k, 'calcext:value-type') === 'error' ? 'error' : a(k, 'office:value-type') ?? a(k, 'calcext:value-type');
      const ps = children(k, 'text:p');
      const s = a(k, 'table:style-name') ?? colStyle[c];
      let v: Scalar = null;
      switch (type) {
        case 'float': case 'percentage': case 'currency': v = Number(a(k, 'office:value') ?? 0); if (!Number.isFinite(v)) v = 0; break;
        case 'date': v = isoToSerial(a(k, 'office:date-value') ?? '', wb.date1904) ?? null; break;
        case 'time': v = durationToDays(a(k, 'office:time-value') ?? '') ?? null; break;
        case 'boolean': v = a(k, 'office:boolean-value') === 'true'; break;
        case 'string': v = a(k, 'office:string-value') ?? ps.map(odfText).join('\n'); break;
        case 'error': { const t = ps.map(odfText).join(''); v = new CellError((ERROR_CODES.includes(t as ErrorCode) ? t : '#VALUE!') as ErrorCode); break; }
        default: if (ps.length && a(k, 'table:formula') === undefined) v = ps.map(odfText).join('\n') || null;
      }
      const cell: Cell = { v };
      const fRaw = a(k, 'table:formula');
      if (fRaw) {
        const f = fromOdfFormula(fRaw);
        if (f) {
          cell.f = f;
          const mr = +(a(k, 'table:number-matrix-rows-spanned') ?? 0), mc = +(a(k, 'table:number-matrix-columns-spanned') ?? 0);
          if (mr > 0 && mc > 0) cell.arrayRange = { r1: r, c1: c, r2: r + mr - 1, c2: c + mc - 1 };
        }
        if (type === undefined && ps.length) { const t = ps.map(odfText).join(''); cell.v = ERROR_CODES.includes(t as ErrorCode) ? new CellError(t as ErrorCode) : t; }
      }
      if (s && s !== 'Default') { const id = styleId(s); if (id) cell.s = id; }
      if (type === 'percentage' && !wb.styles[cell.s ?? 0]?.numFmt) cell.s = internStyle(wb, { ...wb.styles[cell.s ?? 0], numFmt: '0%' });
      if (type === 'date' && !wb.styles[cell.s ?? 0]?.numFmt) cell.s = internStyle(wb, { ...wb.styles[cell.s ?? 0], numFmt: 'yyyy-mm-dd' });
      if (type === 'time' && !wb.styles[cell.s ?? 0]?.numFmt) cell.s = internStyle(wb, { ...wb.styles[cell.s ?? 0], numFmt: 'h:mm:ss' });
      // rich text and links
      if (ps.length) {
        const runs: RichRun[] = [];
        let link: string | undefined;
        const walk = (n: ONode, font: RichRun['font'] | undefined) => {
          for (const kid of n.kids) {
            if (typeof kid === 'string') { if (kid) runs.push(font ? { text: kid, font } : { text: kid }); continue; }
            if (kid.name === 'text:span') walk(kid, { ...font, ...textStyle(a(kid, 'text:style-name')) });
            else if (kid.name === 'text:a') { link ??= a(kid, 'xlink:href'); walk(kid, font); }
            else if (kid.name === 'office:annotation' || kid.name === 'draw:frame') continue;
            else { const t = odfText({ name: '', attrs: {}, kids: [kid] }); if (t) runs.push(font ? { text: t, font } : { text: t }); }
          }
        };
        ps.forEach((p, i) => { if (i) runs.push({ text: '\n' }); walk(p, undefined); });
        const merged: RichRun[] = [];
        for (const run of runs) {
          const last = merged.at(-1);
          if (last && JSON.stringify(last.font ?? null) === JSON.stringify(run.font ?? null)) last.text += run.text;
          else merged.push({ ...run });
        }
        if (typeof cell.v === 'string' && merged.some((m) => m.font && Object.keys(m.font).length) && merged.map((m) => m.text).join('') === cell.v) cell.rich = merged.map((m) => (m.font && Object.keys(m.font).length ? m : { text: m.text }));
        if (link && SAFE_LINK.test(link)) sheet.hyperlinks.set(cellKey(r, c), link.startsWith('#') ? `#${odfLocation(link.slice(1))}` : link);
      }
      const ann = child(k, 'office:annotation');
      if (ann) {
        const author = odfText(child(ann, 'dc:creator') ?? { name: '', attrs: {}, kids: [] }).trim();
        const t = children(ann, 'text:p').map(odfText).join('\n');
        sheet.notes.set(cellKey(r, c), author ? { author, text: t } : { text: t });
      }
      if (cell.v === null && !cell.f && !cell.s) return undefined;
      return cell;
    };

    walkRows(t, false);
    if (headerRows.length) {
      const [r1, r2] = [Math.min(...headerRows.map((h) => h[0])), Math.max(...headerRows.map((h) => h[1]))];
      sheet.pageSetup = { ...sheet.pageSetup, printTitleRows: `$${r1 + 1}:$${r2 + 1}` };
    }
    for (const [name, keys] of validationCells) {
      const dv = validations.get(name);
      if (dv) sheet.validations.push({ ...dv, ranges: toRanges(keys) });
    }
    // ODF conditional styles on cell styles (files that also carry the extended form use that instead)
    const hasExtended = !!child(t, 'calcext:conditional-formats');
    for (const [name, keys] of hasExtended ? [] : mapCells) {
      const node = styleNodes.get(`table-cell/${name}`)!;
      const rules: CfRule[] = [];
      for (const m of children(node, 'style:map')) {
        const st = clean(resolveCell(a(m, 'style:apply-style-name')));
        const rule = conditionRule(a(m, 'style:condition') ?? '', st, true);
        if (rule) rules.push(rule);
      }
      if (rules.length) sheet.conditionalFormats.push({ ranges: toRanges(keys), rules });
    }
    for (const f of children(child(t, 'table:shapes'), 'draw:frame')) readFrame(f, sheet, sheetIdx);
    for (const ne of elems(child(t, 'table:named-expressions'))) readName(ne, sheetIdx);
    readConditionalFormats(t, sheet);
  });

  for (const ne of elems(child(spreadsheet, 'table:named-expressions'))) readName(ne);
  readDatabaseRanges(child(spreadsheet, 'table:database-ranges'));
  readSettings();

  if (meta) {
    const m = findDeep(meta, 'office:meta');
    const get = (n: string) => { const x = child(m, n); return x ? odfText(x).trim() || undefined : undefined; };
    const props: NonNullable<Workbook['props']> = {};
    const title = get('dc:title'), creator = get('meta:initial-creator') ?? get('dc:creator'), created = get('meta:creation-date'), modified = get('dc:date');
    if (title) props.title = title;
    if (creator) props.creator = creator;
    if (created) props.created = created;
    if (modified) props.modified = modified;
    if (Object.keys(props).length) wb.props = props;
  }
  const q = text('clerdo/queries.xml');
  if (q) { const list = queriesFromXml(q); if (list.length) wb.queries = list; }
  if (!wb.sheets.length) wb.sheets.push(newSheet('Sheet1'));
  return wb;

  // ---------- helpers that need the closure ----------

  function rangeText(r: CellRange): string {
    const cl = (c: number) => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    return r.r1 === r.r2 && r.c1 === r.c2 ? `$${cl(r.c1)}$${r.r1 + 1}` : `$${cl(r.c1)}$${r.r1 + 1}:$${cl(r.c2)}$${r.r2 + 1}`;
  }

  function odfLocation(loc: string): string {
    // "Sheet1.A1" or "'My sheet'.A1:B2" → "Sheet1!A1"
    const p = parseRangeAddress(loc);
    if (!p) return loc;
    const sheetPart = p.sheet !== undefined ? (/^[A-Za-z_][A-Za-z0-9_]*$/.test(p.sheet) ? p.sheet : `'${p.sheet.replace(/'/g, "''")}'`) + '!' : '';
    return sheetPart + rangeText(p.range).replace(/\$/g, '');
  }

  function readPageLayout(sheet: Sheet, master: string | undefined): void {
    if (!master || !stylesDoc) return;
    const mp = children(findDeep(stylesDoc, 'office:master-styles'), 'style:master-page').find((m) => a(m, 'style:name') === master);
    const layoutName = a(mp, 'style:page-layout-name');
    const layout = elems(findDeep(stylesDoc, 'office:automatic-styles')).find((x) => x.name === 'style:page-layout' && a(x, 'style:name') === layoutName);
    const pp = child(layout, 'style:page-layout-properties');
    const ps: NonNullable<Sheet['pageSetup']> = {};
    if (pp) {
      if (a(pp, 'style:print-orientation') === 'landscape') ps.orientation = 'landscape';
      const inch = (v: string | undefined) => { const p = toPt(v); return p === undefined ? undefined : Math.round((p / 72) * 1000) / 1000; };
      const l = inch(a(pp, 'fo:margin-left')), r = inch(a(pp, 'fo:margin-right')), t2 = inch(a(pp, 'fo:margin-top')), b = inch(a(pp, 'fo:margin-bottom'));
      // the header and footer areas sit inside the page margins in ODF
      const area = (n: string) => {
        const hp = child(elems(layout).find((x) => x.name === n), 'style:header-footer-properties');
        if (!hp) return 0;
        return (inch(a(hp, 'fo:min-height')) ?? 0) + (inch(a(hp, n === 'style:header-style' ? 'fo:margin-bottom' : 'fo:margin-top')) ?? 0);
      };
      const ha = area('style:header-style'), fa = area('style:footer-style');
      const rd = (v: number) => Math.round(v * 1000) / 1000;
      if (l !== undefined || r !== undefined || t2 !== undefined || b !== undefined) ps.margins = { left: l ?? 0.7, right: r ?? 0.7, top: rd((t2 ?? 0.75) + ha), bottom: rd((b ?? 0.75) + fa), header: ha ? t2 ?? 0.3 : 0.3, footer: fa ? b ?? 0.3 : 0.3 };
      const scale = +(a(pp, 'style:scale-to') ?? '').replace('%', '');
      if (scale && scale !== 100) ps.scale = scale;
      const fx = a(pp, 'style:scale-to-X'), fy = a(pp, 'style:scale-to-Y'), pages = a(pp, 'style:scale-to-pages');
      if (fx !== undefined || fy !== undefined) { ps.fitWidth = +(fx ?? 0); ps.fitHeight = +(fy ?? 0); }
      else if (pages) { ps.fitWidth = +pages; ps.fitHeight = +pages; }
      const print = a(pp, 'style:print') ?? '';
      if (/\bgrid\b/.test(print)) ps.gridlines = true;
      if (/\bheaders\b/.test(print)) ps.headings = true;
      const ctr = a(pp, 'style:table-centering');
      if (ctr === 'horizontal' || ctr === 'both') ps.centerH = true;
      if (ctr === 'vertical' || ctr === 'both') ps.centerV = true;
      const first = a(pp, 'style:first-page-number');
      if (first && first !== 'continue' && +first !== 1) ps.firstPageNumber = +first;
      const w = toPt(a(pp, 'fo:page-width')), h = toPt(a(pp, 'fo:page-height'));
      if (w && h) {
        const [sw, sh] = [Math.min(w, h), Math.max(w, h)];
        const paper = Math.abs(sw - 595) < 6 && Math.abs(sh - 842) < 6 ? 9 : Math.abs(sw - 612) < 6 && Math.abs(sh - 792) < 6 ? 1 : Math.abs(sw - 612) < 6 && Math.abs(sh - 1008) < 6 ? 5 : undefined;
        if (paper) ps.paper = paper;
      }
    }
    const hf = (n: ONode | undefined): string | undefined => {
      if (!n || a(n, 'style:display') === 'false') return undefined;
      const part = (tagName: string, code: string) => { const x = child(n, tagName); return x ? code + children(x, 'text:p').map(odfText).join('\n') : ''; };
      const s = part('style:region-left', '&L') + part('style:region-center', '&C') + part('style:region-right', '&R') || (children(n, 'text:p').map(odfText).join('\n') ? '&C' + children(n, 'text:p').map(odfText).join('\n') : '');
      return s.replace(/^&[LCR]$/, '') || undefined;
    };
    const header = hf(child(mp, 'style:header')), footer = hf(child(mp, 'style:footer'));
    if (header) ps.header = header;
    if (footer) ps.footer = footer;
    if (Object.keys(ps).length) sheet.pageSetup = { ...sheet.pageSetup, ...ps };
  }

  function readFrame(f: ONode, sheet: Sheet, sheetIdx: number, at?: { row: number; col: number }): void {
    const w = toPt(a(f, 'svg:width')) ?? 0, h = toPt(a(f, 'svg:height')) ?? 0;
    const x = toPt(a(f, 'svg:x')) ?? 0, y = toPt(a(f, 'svg:y')) ?? 0;
    const px = (pt: number) => Math.round((pt * 96) / 72);
    let anchor: Sheet['pictures'][number]['anchor'];
    if (at) anchor = { row: at.row, col: at.col, dx: px(x), dy: px(y), width: px(w), height: px(h) };
    else {
      // page anchored: walk the grid from the top-left corner
      let rx = px(x), ry = px(y), c = 0, r = 0;
      const cw = (i: number) => (sheet.cols.get(i)?.hidden ? 0 : Math.round(((sheet.cols.get(i)?.width ?? sheet.defaultColWidth) * 7) + 5));
      const rh = (i: number) => (sheet.rows.get(i)?.hidden ? 0 : ((sheet.rows.get(i)?.height ?? sheet.defaultRowHeight) * 96) / 72);
      while (c < MAX_COL - 1 && rx >= cw(c) && cw(c) > 0) rx -= cw(c++);
      while (r < MAX_ROW - 1 && ry >= rh(r) && rh(r) > 0) ry -= rh(r++);
      anchor = { row: r, col: c, dx: Math.round(rx), dy: Math.round(ry), width: px(w), height: px(h) };
    }
    const alt = child(f, 'svg:desc') ? odfText(child(f, 'svg:desc')!) : child(f, 'svg:title') ? odfText(child(f, 'svg:title')!) : undefined;
    const img = child(f, 'draw:image');
    const obj = child(f, 'draw:object');
    if (obj) {
      const href = (a(obj, 'xlink:href') ?? '').replace(/^\.\//, '').replace(/\/$/, '');
      const doc = text(`${href}/content.xml`);
      if (doc && !href.includes('..')) {
        const chart = readChartObject(parseOdf(doc), `ods-chart-${sheetIdx}-${sheet.charts.length + 1}`, anchor, a(f, 'clerdo:chart'));
        if (chart) { if (alt) chart.alt = alt; sheet.charts.push(chart); return; }
      }
    }
    if (img) {
      const href = (a(img, 'xlink:href') ?? '').replace(/^\.\//, '');
      const bytes = !href.includes('..') && !/^[a-z]+:/i.test(href) ? files[href] : undefined;
      const ext = href.split('.').pop()?.toLowerCase() ?? '';
      const mime = a(img, 'draw:mime-type') ?? a(img, 'loext:mime-type') ?? PICTURE_MIME[ext];
      if (bytes && mime && /^image\//.test(mime)) {
        const pic: Sheet['pictures'][number] = { id: `ods-${sheetIdx}-${sheet.pictures.length + 1}`, data: bytes, mime, anchor };
        if (alt) pic.alt = alt;
        sheet.pictures.push(pic);
      }
    }
  }

  function readName(n: ONode, local?: number): void {
    const name = a(n, 'table:name');
    if (!name) return;
    let formula: string | undefined;
    if (n.name === 'table:named-range') {
      const addr = a(n, 'table:cell-range-address');
      if (!addr) return;
      formula = splitAddresses(addr).map((p) => gridRef(p)).map((r) => r.replace(/(^|!|:)(\$?[A-Z]+)(\$?)(\d+)/g, '$1$2$3$4')).join(',');
      if (/print-range/.test(a(n, 'table:range-usable-as') ?? '') && /Print_Area/i.test(name)) {
        const p = parseRangeAddress(addr);
        const k = p?.sheet !== undefined ? wb.sheets.findIndex((s) => s.name === p.sheet) : local ?? -1;
        if (p && k >= 0) { wb.sheets[k]!.pageSetup = { ...wb.sheets[k]!.pageSetup, printArea: rangeText(p.range) }; return; }
      }
      if (!formula || formula.includes('#REF!')) return;
    } else if (n.name === 'table:named-expression') {
      formula = fromOdfFormula(a(n, 'table:expression') ?? '');
    }
    if (!formula) return;
    if (/^_xlnm\.Print_Area$/i.test(name)) return;
    const d: Workbook['names'][number] = { name, formula };
    if (local !== undefined) d.sheet = local;
    if (a(n, 'clerdo:hidden') === 'true') d.hidden = true;
    wb.names.push(d);
  }

  function readDatabaseRanges(n: ONode | undefined): void {
    for (const db of children(n, 'table:database-range')) {
      if (a(db, 'table:display-filter-buttons') !== 'true') continue;
      const p = parseRangeAddress(a(db, 'table:target-range-address'));
      const k = p?.sheet !== undefined ? wb.sheets.findIndex((s) => s.name === p.sheet) : -1;
      if (!p || k < 0) continue;
      const af: AutoFilter = { range: p.range, columns: [] };
      const filter = child(db, 'table:filter');
      const visit = (g: ONode, and: boolean): void => {
        for (const c of elems(g)) {
          if (c.name === 'table:filter-and') visit(c, true);
          else if (c.name === 'table:filter-or') visit(c, false);
          else if (c.name === 'table:filter-condition') {
            const field = +(a(c, 'table:field-number') ?? 0);
            let fc = af.columns.find((x) => x.col === p.range.c1 + field);
            if (!fc) { fc = { col: p.range.c1 + field }; af.columns.push(fc); }
            const op = a(c, 'table:operator') ?? '=';
            const value = a(c, 'table:value') ?? '';
            const items = children(c, 'table:filter-set-item').map((x) => a(x, 'table:value') ?? '');
            if (items.length) { fc.values = items; continue; }
            if (op === 'empty') { fc.blanks = true; continue; }
            if (/^(top|bottom) (values|percent)$/.test(op)) { fc.top10 = { top: op.startsWith('top'), percent: op.endsWith('percent'), value: +value || 10 }; continue; }
            const map: Record<string, CfOperator> = { '=': 'equal', '!=': 'notEqual', '<': 'lessThan', '>': 'greaterThan', '<=': 'lessThanOrEqual', '>=': 'greaterThanOrEqual' };
            let operator = map[op];
            let v = value;
            if (!operator) {
              if (op === '!empty') { operator = 'notEqual'; v = ''; }
              else if (op === 'contains') { operator = 'equal'; v = `*${value}*`; }
              else if (op === '!contains') { operator = 'notEqual'; v = `*${value}*`; }
              else if (op === 'begins') { operator = 'equal'; v = `${value}*`; }
              else if (op === 'ends') { operator = 'equal'; v = `*${value}`; }
              else continue;
            }
            fc.custom ??= { and, rules: [] };
            fc.custom.rules.push({ operator, value: v });
          }
        }
      };
      if (filter) visit(filter, true);
      wb.sheets[k]!.autoFilter = af;
    }
  }

  function readConditionalFormats(t: ONode, sheet: Sheet): void {
    for (const cf of children(child(t, 'calcext:conditional-formats'), 'calcext:conditional-format')) {
      const ranges = splitAddresses(a(cf, 'calcext:target-range-address') ?? '').map((x) => parseRangeAddress(x)?.range).filter((x): x is CellRange => !!x);
      if (!ranges.length) continue;
      const rules: CfRule[] = [];
      for (const c of elems(cf)) {
        if (c.name === 'calcext:condition') {
          const st = clean(resolveCell(a(c, 'calcext:apply-style-name')));
          const rule = conditionRule(a(c, 'calcext:value') ?? '', st, false);
          if (rule) rules.push(rule);
        } else if (c.name === 'calcext:color-scale') {
          const stops = children(c, 'calcext:color-scale-entry').map((e) => ({ value: cfValue(e), color: (a(e, 'calcext:color') ?? '#000000').toUpperCase() }));
          if (stops.length >= 2) rules.push({ type: 'colorScale', stops });
        } else if (c.name === 'calcext:data-bar') {
          const entries = children(c, 'calcext:formatting-entry');
          rules.push({ type: 'dataBar', min: cfValue(entries[0]), max: cfValue(entries[1]), color: (a(c, 'calcext:positive-color') ?? '#638EC6').toUpperCase(), gradient: a(c, 'calcext:gradient') !== 'false' });
        } else if (c.name === 'calcext:icon-set') {
          const ICONS: Record<string, string> = { '3Arrows': '3Arrows', '3ArrowsGray': '3ArrowsGray', '3Flags': '3Flags', '3TrafficLights1': '3TrafficLights1', '3TrafficLights2': '3TrafficLights2', '3Signs': '3Signs', '3Symbols': '3Symbols', '3Symbols2': '3Symbols2', '4Arrows': '4Arrows', '4ArrowsGray': '4ArrowsGray', '4RedToBlack': '4RedToBlack', '4Rating': '4Rating', '4TrafficLights': '4TrafficLights', '5Arrows': '5Arrows', '5ArrowsGray': '5ArrowsGray', '5Rating': '5Rating', '5Quarters': '5Quarters' };
          const set = ICONS[a(c, 'calcext:icon-set-type') ?? ''] ?? '3TrafficLights1';
          rules.push({ type: 'iconSet', set, values: children(c, 'calcext:formatting-entry').map(cfValue), ...(a(c, 'calcext:show-value') === 'false' ? { showValue: false } : {}) });
        } else if (c.name === 'calcext:date-is') {
          const st = clean(resolveCell(a(c, 'calcext:style')));
          const d = a(c, 'calcext:date') ?? 'today';
          const map: Record<string, Extract<CfRule, { type: 'timePeriod' }>['period']> = { today: 'today', yesterday: 'yesterday', tomorrow: 'tomorrow', 'last-7-days': 'last7Days', 'this-week': 'thisWeek', 'last-week': 'lastWeek', 'next-week': 'nextWeek', 'this-month': 'thisMonth', 'last-month': 'lastMonth', 'next-month': 'nextMonth' };
          if (map[d]) rules.push({ type: 'timePeriod', period: map[d]!, style: st });
        }
      }
      if (rules.length) sheet.conditionalFormats.push({ ranges, rules });
    }
  }

  function readSettings(): void {
    if (!settings) return;
    const views = findDeep(settings, 'config:config-item-map-indexed');
    const entry = child(views, 'config:config-item-map-entry');
    const item = (n: ONode | undefined, name: string): string | undefined => { const x = children(n, 'config:config-item').find((i) => a(i, 'config:name') === name); return x ? odfText(x) : undefined; };
    const active = item(entry, 'ActiveTable');
    if (active) { const k = wb.sheets.findIndex((s) => s.name === active); if (k >= 0) wb.activeSheet = k; }
    const tables = children(entry, 'config:config-item-map-named').find((m) => a(m, 'config:name') === 'Tables');
    for (const t of children(tables, 'config:config-item-map-entry')) {
      const sheet = wb.sheets.find((s) => s.name === a(t, 'config:name'));
      if (!sheet) continue;
      if (item(t, 'HorizontalSplitMode') === '2') sheet.view.frozenCols = +(item(t, 'HorizontalSplitPosition') ?? 0) || 0;
      if (item(t, 'VerticalSplitMode') === '2') sheet.view.frozenRows = +(item(t, 'VerticalSplitPosition') ?? 0) || 0;
      const zoom = +(item(t, 'ZoomValue') ?? 0);
      if (zoom >= 10 && zoom <= 400) sheet.view.zoom = zoom;
      const grid = item(t, 'ShowGrid') ?? item(entry, 'ShowGrid');
      if (grid === 'false') sheet.view.showGrid = false;
      if (item(t, 'HasColumnRowHeaders') === 'false' || item(entry, 'HasColumnRowHeaders') === 'false') sheet.view.showHeaders = false;
      const cx = +(item(t, 'CursorPositionX') ?? -1), cy = +(item(t, 'CursorPositionY') ?? -1);
      if (cx >= 0 && cy >= 0 && cx < MAX_COL && cy < MAX_ROW) sheet.view.selection = { r1: cy, r2: cy, c1: cx, c2: cx };
    }
  }
};

/** Decrypts streams the manifest marks as encrypted, in place. */
const decryptPackage = (files: Record<string, Uint8Array>, password: string | undefined): void => {
  const m = files['META-INF/manifest.xml'];
  if (!m) return;
  const doc = parseOdf(new TextDecoder().decode(m));
  const list = children(findDeep(doc, 'manifest:manifest'), 'manifest:file-entry').filter((e) => child(e, 'manifest:encryption-data'));
  if (!list.length) return;
  if (password === undefined) throw new PasswordRequired(false);
  for (const e of list) {
    const path = a(e, 'manifest:full-path') ?? '';
    const bytes = files[path];
    if (!bytes) continue;
    const ed = child(e, 'manifest:encryption-data')!;
    const alg = child(ed, 'manifest:algorithm'), kd = child(ed, 'manifest:key-derivation'), sk = child(ed, 'manifest:start-key-generation');
    const info: EntryEncryption = {
      algorithm: a(alg, 'manifest:algorithm-name') ?? '', iv: a(alg, 'manifest:initialisation-vector') ?? '', salt: a(kd, 'manifest:salt') ?? '',
      iterations: Math.min(10_000_000, +(a(kd, 'manifest:iteration-count') ?? 1024) || 1024), keySize: +(a(kd, 'manifest:key-size') ?? 16) || 16,
      startKey: a(sk, 'manifest:start-key-generation-name') ?? 'SHA1', checksumType: a(ed, 'manifest:checksum-type') ?? '', checksum: a(ed, 'manifest:checksum') ?? '',
    };
    const kdName = a(kd, 'manifest:key-derivation-name') ?? '';
    if (/argon2id/i.test(kdName)) {
      const num = (suffix: string, dflt: number) => { const k = Object.keys(kd!.attrs).find((x) => x.endsWith(suffix)); const v = k ? +kd!.attrs[k]! : NaN; return Number.isFinite(v) && v > 0 ? v : dflt; };
      info.argon2 = { iterations: num('argon2-iterations', 3), memoryKiB: num('argon2-memory', 65536), lanes: num('argon2-lanes', 4) };
      if (!a(kd, 'manifest:key-size')) info.keySize = 32;
    }
    if (!sk) info.noStartKey = true;
    try { files[path] = decryptStream(bytes, info, password); }
    catch (err) { if (/encryption method|too much memory/.test((err as Error).message)) throw err; throw new PasswordRequired(true); }
  }
  // whole-package encryption: the real document is a zip inside "encrypted-package"
  const inner = files['encrypted-package'];
  if (inner && list.some((e) => a(e, 'manifest:full-path') === 'encrypted-package') && inner[0] === 0x50 && inner[1] === 0x4b) {
    const unpacked = unzipSync(inner);
    delete files['encrypted-package'];
    for (const [k, v] of Object.entries(unpacked)) files[k] = v;
    if (unpacked['META-INF/manifest.xml']) decryptPackage(files, password);
  }
};

/** The 16-bit password hash behind a legacy-hash protection key (found by trying every value). */
export const legacyHash = (n: ONode | undefined): string | undefined => {
  const key = a(n, 'table:protection-key');
  const alg = a(n, 'table:protection-key-digest-algorithm') ?? '';
  const alg2 = a(n, 'loext:protection-key-digest-algorithm-2') ?? '';
  if (!key || !/legacy-hash-excel/.test(alg)) return undefined;
  let want: Uint8Array;
  try { want = fromBase64(key); } catch { return undefined; }
  const h = /sha256/.test(alg2) ? sha256 : sha1;
  const b = new Uint8Array(2);
  for (let v = 0; v < 65536; v++) {
    b[0] = v >> 8; b[1] = v & 0xff;
    const d = h(b);
    if (d.length === want.length && d.every((x, i) => x === want[i])) return v.toString(16).toUpperCase().padStart(4, '0');
  }
  return undefined;
};

const cfValue = (e: ONode | undefined): { type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'; value?: string } => {
  const t = a(e, 'calcext:type') ?? 'minimum';
  const v = a(e, 'calcext:value');
  switch (t) {
    case 'minimum': return { type: 'min' };
    case 'maximum': return { type: 'max' };
    case 'auto-minimum': return { type: 'autoMin' };
    case 'auto-maximum': return { type: 'autoMax' };
    case 'percent': return { type: 'percent', value: v ?? '0' };
    case 'percentile': return { type: 'percentile', value: v ?? '50' };
    case 'formula': return { type: 'formula', value: expr(v ?? '0') ?? v ?? '0' };
    default: return { type: 'num', value: v ?? '0' };
  }
};

/** A conditional-format condition (calcext:value or style:condition) as a rule. */
const conditionRule = (raw: string, style: CellStyle, odfStd: boolean): CfRule | undefined => {
  let s = raw.trim();
  if (odfStd) s = s.replace(/^of:/, '').replace(/^cell-content\(\)\s*/, '').replace(/^cell-content-is-between\(/, 'between(').replace(/^cell-content-is-not-between\(/, 'not-between(').replace(/^is-true-formula\(/, 'formula-is(');
  const call = callOf(s);
  if (call) {
    const [fn, inner] = call;
    const args = splitArgs(inner);
    switch (fn) {
      case 'between': case 'not-between': {
        const f = args.map(expr);
        if (f.every((x): x is string => !!x) && f.length === 2) return { type: 'cellIs', operator: fn === 'between' ? 'between' : 'notBetween', formulas: f, style };
        return undefined;
      }
      case 'formula-is': { const f = expr(inner); return f ? { type: 'expression', formula: f, style } : undefined; }
      case 'begins-with': case 'ends-with': case 'contains-text': case 'not-contains-text': {
        const t = args[0]?.replace(/^"|"$/g, '').replace(/""/g, '"') ?? '';
        const type = ({ 'begins-with': 'beginsWith', 'ends-with': 'endsWith', 'contains-text': 'containsText', 'not-contains-text': 'notContainsText' } as const)[fn];
        return { type, text: t, style };
      }
      case 'top-elements': case 'bottom-elements': case 'top-percent': case 'bottom-percent':
        return { type: 'top10', rank: +(args[0] ?? 10) || 10, bottom: fn.startsWith('bottom') || undefined, percent: fn.endsWith('percent') || undefined, style };
      default: return undefined;
    }
  }
  switch (s) {
    case 'duplicate': return { type: 'duplicateValues', style };
    case 'unique': return { type: 'uniqueValues', style };
    case 'above-average': return { type: 'aboveAverage', style };
    case 'below-average': return { type: 'aboveAverage', below: true, style };
    case 'above-equal-average': return { type: 'aboveAverage', equal: true, style };
    case 'below-equal-average': return { type: 'aboveAverage', below: true, equal: true, style };
    case 'is-error': return { type: 'containsErrors', style };
    case 'is-no-error': return { type: 'notContainsErrors', style };
    case 'is-blank': return { type: 'containsBlanks', style };
    case 'is-no-blank': return { type: 'notContainsBlanks', style };
    default: break;
  }
  const m = /^(<=|>=|!=|<>|=|<|>)(.*)$/s.exec(s);
  if (m) {
    const f = expr(m[2]!);
    const op: Record<string, CfOperator> = { '<=': 'lessThanOrEqual', '>=': 'greaterThanOrEqual', '!=': 'notEqual', '<>': 'notEqual', '=': 'equal', '<': 'lessThan', '>': 'greaterThan' };
    return f ? { type: 'cellIs', operator: op[m[1]!]!, formulas: [f], style } : undefined;
  }
  return undefined;
};

const readValidation = (v: ONode): Omit<DataValidation, 'ranges'> | undefined => {
  const cond = (a(v, 'table:condition') ?? '').replace(/^[a-z]+:/, '');
  const dv: Omit<DataValidation, 'ranges'> = { type: 'any' };
  const parts = cond.split(/\s+and\s+/);
  const typeFn: Record<string, DataValidation['type']> = {
    'cell-content-is-whole-number()': 'whole', 'cell-content-is-decimal-number()': 'decimal', 'cell-content-is-date()': 'date', 'cell-content-is-time()': 'time',
  };
  for (const p of parts) {
    const t = p.trim();
    if (typeFn[t]) { dv.type = typeFn[t]!; continue; }
    let m: RegExpExecArray | null;
    if ((m = /^cell-content-is-in-list\((.*)\)$/s.exec(t))) {
      dv.type = 'list';
      const args = splitArgs(m[1]!);
      if (args.length === 1 && /^\[.*\]$/.test(args[0]!)) dv.formula1 = expr(args[0]!);
      else if (args.every((x) => /^".*"$/s.test(x))) dv.formula1 = `"${args.map((x) => x.slice(1, -1).replace(/""/g, '"')).join(',')}"`;
      else { const f = expr(m[1]!); if (f) dv.formula1 = f; }
      continue;
    }
    if ((m = /^cell-content-text-length\(\)\s*(<=|>=|!=|=|<|>)(.*)$/s.exec(t))) { dv.type = 'textLength'; setOp(dv, m[1]!, m[2]!); continue; }
    if ((m = /^cell-content-text-length-is-(not-)?between\((.*)\)$/s.exec(t))) { dv.type = 'textLength'; setBetween(dv, !!m[1], m[2]!); continue; }
    if ((m = /^cell-content-is-(not-)?between\((.*)\)$/s.exec(t))) { setBetween(dv, !!m[1], m[2]!); continue; }
    if ((m = /^cell-content\(\)\s*(<=|>=|!=|=|<|>)(.*)$/s.exec(t))) { setOp(dv, m[1]!, m[2]!); continue; }
    if ((m = /^is-true-formula\((.*)\)$/s.exec(t))) { dv.type = 'custom'; const f = expr(m[1]!); if (f) dv.formula1 = f; continue; }
  }
  if (dv.type === 'any' && !cond) return undefined;
  if (a(v, 'table:allow-empty-cell') !== 'false') dv.allowBlank = true;
  if (dv.type === 'list') dv.showDropdown = a(v, 'table:display-list') !== 'none';
  const help = child(v, 'table:help-message');
  if (help && a(help, 'table:display') !== 'false') {
    const t = children(help, 'text:p').map(odfText).join('\n');
    if (t || a(help, 'table:title')) dv.prompt = { ...(a(help, 'table:title') ? { title: a(help, 'table:title') } : {}), text: t };
  }
  const err = child(v, 'table:error-message');
  if (err && a(err, 'table:display') !== 'false') {
    const kind = a(err, 'table:message-type');
    dv.error = { style: kind === 'warning' ? 'warning' : kind === 'information' ? 'information' : 'stop', ...(a(err, 'table:title') ? { title: a(err, 'table:title') } : {}), text: children(err, 'text:p').map(odfText).join('\n') };
  }
  return dv;
};

const setOp = (dv: Omit<DataValidation, 'ranges'>, op: string, value: string): void => {
  const map: Record<string, CfOperator> = { '<=': 'lessThanOrEqual', '>=': 'greaterThanOrEqual', '!=': 'notEqual', '=': 'equal', '<': 'lessThan', '>': 'greaterThan' };
  dv.operator = map[op];
  const f = expr(value);
  if (f) dv.formula1 = f;
  if (dv.type === 'any') dv.type = 'decimal';
};
const setBetween = (dv: Omit<DataValidation, 'ranges'>, not: boolean, inner: string): void => {
  const [x, y] = splitArgs(inner);
  dv.operator = not ? 'notBetween' : 'between';
  const f1 = expr(x ?? ''), f2 = expr(y ?? '');
  if (f1) dv.formula1 = f1;
  if (f2) dv.formula2 = f2;
  if (dv.type === 'any') dv.type = 'decimal';
};

export type { FilterColumn };
export { shownColWidth };
