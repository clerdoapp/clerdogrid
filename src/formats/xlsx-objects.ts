// Floating shapes (DrawingML xdr:sp, ECMA-376 Part 1, 20.5 and 20.1.9), protected ranges, and the
// ClerdoGrid part that keeps comment threads, checkbox cells and ink exactly as the app holds them.
import {
  cellKey, internStyle, keyCol, keyRow, addr, parseAddr, parseRange, rangeAddr,
  type Chart, type CommentThread, type InkStroke, type Note, type Shape, type ShapeKind, type Sheet, type Workbook,
} from '../model/workbook';
import { attr, esc, find, kid, kids, numAttr, type XNode } from './xml';
import { EMU_PER_PX } from './xlsx-chart';

/** Relationship type and content type of the ClerdoGrid objects part. */
export const OBJECTS_REL = 'https://schemas.clerdo.app/grid/2026/relationships/clerdogridObjects';
export const OBJECTS_REL_SHORT = 'clerdogridObjects';
export const OBJECTS_PATH = 'xl/clerdogrid/objects.json';
export const OBJECTS_CT = 'application/vnd.clerdo.grid.objects+json';

const PRST: Record<ShapeKind, string> = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', leftArrow: 'leftArrow',
  upArrow: 'upArrow', downArrow: 'downArrow', line: 'line', arrowLine: 'line', star5: 'star5', diamond: 'diamond', pentagon: 'pentagon',
  hexagon: 'hexagon', callout: 'wedgeRectCallout', textBox: 'rect',
};
const FROM_PRST: Record<string, ShapeKind> = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rightArrow: 'rightArrow', leftArrow: 'leftArrow',
  upArrow: 'upArrow', downArrow: 'downArrow', line: 'line', straightConnector1: 'line', star5: 'star5', diamond: 'diamond', pentagon: 'pentagon',
  hexagon: 'hexagon', wedgeRectCallout: 'callout', wedgeRoundRectCallout: 'callout',
};

const hex = (c: string | undefined): string | undefined => {
  const m = /^#?([0-9a-f]{6})/i.exec(c ?? '');
  return m ? m[1]!.toUpperCase() : undefined;
};
const solid = (c: string): string => `<a:solidFill><a:srgbClr val="${hex(c) ?? '000000'}"/></a:solidFill>`;
const isLine = (k: ShapeKind): boolean => k === 'line' || k === 'arrowLine';

/** One shape as an xdr:sp element. */
export const shapeXml = (id: number, s: Shape): string => {
  const box = s.kind === 'textBox';
  const nameBase = box ? 'TextBox' : isLine(s.kind) ? 'Line' : 'Shape';
  const rot = s.rotation ? ` rot="${Math.round(((s.rotation % 360) + 360) % 360 * 60000)}"` : '';
  const xfrm = `<a:xfrm${rot}><a:off x="0" y="0"/><a:ext cx="${Math.round(s.anchor.width * EMU_PER_PX)}" cy="${Math.round(s.anchor.height * EMU_PER_PX)}"/></a:xfrm>`;
  const fill = isLine(s.kind) ? '' : s.fill ? solid(s.fill) : '<a:noFill/>';
  let ln = '<a:ln><a:noFill/></a:ln>';
  if (s.line && s.line.width > 0) {
    const dash = s.line.dash === 'dash' ? '<a:prstDash val="dash"/>' : s.line.dash === 'dot' ? '<a:prstDash val="sysDot"/>' : '';
    ln = `<a:ln w="${Math.round(s.line.width * EMU_PER_PX)}">${solid(s.line.color)}${dash}${s.kind === 'arrowLine' ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>`;
  }
  let tx = '';
  if (!isLine(s.kind)) {
    const algn = s.align === 'left' ? 'l' : s.align === 'right' ? 'r' : 'ctr';
    const f = s.font ?? {};
    const rPr = `<a:rPr lang="en-US" sz="${Math.round((f.size ?? 11) * 100)}"${f.bold ? ' b="1"' : ''}${f.italic ? ' i="1"' : ''}>${f.color ? solid(f.color) : ''}${f.name ? `<a:latin typeface="${esc(f.name)}"/>` : ''}</a:rPr>`;
    const paras = (s.text ?? '').split('\n').map((line) => `<a:p><a:pPr algn="${algn}"/>${line ? `<a:r>${rPr}<a:t>${esc(line)}</a:t></a:r>` : `<a:endParaRPr lang="en-US" sz="${Math.round((f.size ?? 11) * 100)}"/>`}</a:p>`).join('');
    tx = `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="${box ? 't' : 'ctr'}"/><a:lstStyle/>${paras}</xdr:txBody>`;
  }
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${nameBase} ${id}"${s.alt ? ` descr="${esc(s.alt)}"` : ''}/><xdr:cNvSpPr${box ? ' txBox="1"' : ''}/></xdr:nvSpPr>`
    + `<xdr:spPr>${xfrm}<a:prstGeom prst="${PRST[s.kind]}"><a:avLst/></a:prstGeom>${fill}${ln}</xdr:spPr>${tx}</xdr:sp>`;
};

const colorIn = (n: XNode | undefined): string | undefined => {
  const c = kid(n, 'srgbClr');
  const v = attr(c, 'val');
  return v && /^[0-9a-f]{6}$/i.test(v) ? `#${v.toUpperCase()}` : undefined;
};

/** Reads an xdr:sp element. Shapes with custom geometry come back as rectangles. */
export const readShape = (sp: XNode, anchor: Chart['anchor'], id: string): Shape => {
  const spPr = kid(sp, 'spPr');
  const prst = attr(kid(spPr, 'prstGeom'), 'prst') ?? 'rect';
  const nv = kid(sp, 'nvSpPr');
  const txBox = attr(kid(nv, 'cNvSpPr'), 'txBox') === '1';
  const lnNode = kid(spPr, 'ln');
  let kind: ShapeKind = txBox ? 'textBox' : FROM_PRST[prst] ?? 'rect';
  if (kind === 'line' && (kid(lnNode, 'tailEnd') && attr(kid(lnNode, 'tailEnd'), 'type') !== 'none')) kind = 'arrowLine';
  const shape: Shape = { id, kind, anchor };
  const solidFill = kid(spPr, 'solidFill');
  if (solidFill) shape.fill = colorIn(solidFill) ?? '#D6EFE2';
  else if (!kid(spPr, 'noFill') && !txBox && !kind.includes('ine')) shape.fill = '#D6EFE2';
  if (lnNode && !kid(lnNode, 'noFill')) {
    const w = numAttr(lnNode, 'w');
    const dash = attr(kid(lnNode, 'prstDash'), 'val');
    shape.line = { color: colorIn(kid(lnNode, 'solidFill')) ?? '#0C7A4B', width: w !== undefined ? Math.max(0.5, Math.round((w / EMU_PER_PX) * 4) / 4) : 1, dash: !dash || dash === 'solid' ? 'solid' : /dot/i.test(dash) ? 'dot' : 'dash' };
  } else if (!lnNode && !txBox) shape.line = { color: '#0C7A4B', width: 1, dash: 'solid' };
  const rot = numAttr(kid(spPr, 'xfrm'), 'rot');
  if (rot) shape.rotation = Math.round(rot / 60000);
  const descr = attr(kid(nv, 'cNvPr'), 'descr');
  if (descr) shape.alt = descr;
  const tb = kid(sp, 'txBody');
  if (tb) {
    const ps = kids(tb, 'p');
    const lines = ps.map((p) => kids(p, 'r').map((r) => kid(r, 't')?.text ?? '').join(''));
    const text = lines.join('\n');
    if (text) shape.text = text.slice(0, 32767);
    const algn = attr(kid(ps[0], 'pPr'), 'algn');
    if (algn === 'l') shape.align = 'left';
    else if (algn === 'r') shape.align = 'right';
    else if (algn === 'ctr') shape.align = 'center';
    const rPr = find(tb, 'rPr');
    if (rPr) {
      const font: NonNullable<Shape['font']> = {};
      const sz = numAttr(rPr, 'sz');
      if (sz) font.size = Math.min(400, Math.max(1, sz / 100));
      if (attr(rPr, 'b') === '1') font.bold = true;
      if (attr(rPr, 'i') === '1') font.italic = true;
      const c = colorIn(kid(rPr, 'solidFill'));
      if (c) font.color = c;
      const face = attr(kid(rPr, 'latin'), 'typeface');
      if (face && !face.startsWith('+')) font.name = face;
      if (Object.keys(font).length) shape.font = font;
    }
  }
  return shape;
};

/** Adds alternative text to the first cNvPr of a drawing object. */
export const withDescr = (xml: string, alt: string | undefined): string =>
  alt ? xml.replace(/<xdr:cNvPr ([^>]*?)(\/?)>/, (_, a: string, close: string) => `<xdr:cNvPr ${a} descr="${esc(alt)}"${close}>`) : xml;

/**
 * Notes plus comment threads as legacy comments, so other apps show them. A thread becomes its
 * first message with each reply on its own line. A cell's note wins over its thread.
 */
export const legacyNotes = (s: Sheet): Map<number, Note> => {
  if (!s.comments?.size) return s.notes;
  const out = new Map(s.notes);
  for (const [k, t] of s.comments) {
    if (out.has(k)) continue;
    const lines = [`${t.author}: ${t.text}`, ...t.replies.map((r) => `${r.author}: ${r.text}`)];
    if (t.resolved) lines.push('(Resolved)');
    out.set(k, { author: t.author, text: lines.join('\n') });
  }
  return out;
};

const sanitizeEditRanges = (list: unknown): Sheet['editRanges'] => {
  if (!Array.isArray(list)) return undefined;
  const out: NonNullable<Sheet['editRanges']> = [];
  for (const e of list.slice(0, 1000)) {
    const range = typeof e?.range === 'string' ? parseRange(e.range) : null;
    if (!range) continue;
    out.push({ title: String(e.title ?? 'Range').slice(0, 255), range, ...(typeof e.password === 'string' && /^[0-9A-F]{4}$/i.test(e.password) ? { password: e.password.toUpperCase() } : {}) });
  }
  return out.length ? out : undefined;
};

/** Worksheet protectedRanges element (ranges people may edit while the sheet is protected). */
export const protectedRangesXml = (s: Sheet): string => {
  if (!s.editRanges?.length) return '';
  return `<protectedRanges>${s.editRanges.map((e) => `<protectedRange${e.password ? ` password="${esc(e.password)}"` : ''} sqref="${rangeAddr(e.range)}" name="${esc(e.title.replace(/[^\w.]/g, '_') || 'Range')}"/>`).join('')}</protectedRanges>`;
};

export const readProtectedRanges = (ws: XNode, sheet: Sheet): void => {
  const list = kids(kid(ws, 'protectedRanges'), 'protectedRange').map((p) => ({ title: attr(p, 'name') ?? 'Range', range: (attr(p, 'sqref') ?? '').split(/\s+/)[0], password: attr(p, 'password') }));
  const r = sanitizeEditRanges(list);
  if (r) sheet.editRanges = r;
};

interface SheetObjects {
  comments?: { ref: string; thread: CommentThread }[];
  /** Cells whose legacy comment was written from a thread (not a real note). */
  threadNotes?: string[];
  checkboxes?: string[];
  ink?: InkStroke[];
  /** Titles kept as typed (the file's protectedRange names allow fewer characters). */
  editRangeTitles?: string[];
}
interface ObjectsDoc { app: 'ClerdoGrid'; version: 1; sheets: Record<string, SheetObjects> }

/** JSON for the ClerdoGrid objects part, or null when there is nothing to keep. */
export const objectsPart = (wb: Workbook): string | null => {
  const doc: ObjectsDoc = { app: 'ClerdoGrid', version: 1, sheets: {} };
  wb.sheets.forEach((s, i) => {
    const o: SheetObjects = {};
    if (s.comments?.size) {
      o.comments = [...s.comments].sort((a, b) => a[0] - b[0]).map(([k, thread]) => ({ ref: addr(keyRow(k), keyCol(k)), thread }));
      const tn = [...s.comments.keys()].filter((k) => !s.notes.has(k)).map((k) => addr(keyRow(k), keyCol(k)));
      if (tn.length) o.threadNotes = tn;
    }
    const boxes: string[] = [];
    for (const [k, c] of s.cells) if (c.s && wb.styles[c.s]?.checkbox) boxes.push(addr(keyRow(k), keyCol(k)));
    if (boxes.length) o.checkboxes = boxes;
    if (s.ink?.length) o.ink = s.ink;
    if (s.editRanges?.length) o.editRangeTitles = s.editRanges.map((e) => e.title);
    if (Object.keys(o).length) doc.sheets[String(i)] = o;
  });
  return Object.keys(doc.sheets).length ? JSON.stringify(doc) : null;
};

const str = (v: unknown, max = 32767): string => (typeof v === 'string' ? v.slice(0, max) : '');
const num = (v: unknown): number => (typeof v === 'number' && Number.isFinite(v) ? v : 0);

const cleanThread = (t: Partial<CommentThread> | undefined, i: number): CommentThread | null => {
  if (!t || typeof t !== 'object') return null;
  return {
    id: str(t.id, 64) || `c${i + 1}`, author: str(t.author, 255) || 'You', text: str(t.text), time: str(t.time, 40),
    ...(t.resolved ? { resolved: true } : {}),
    replies: Array.isArray(t.replies) ? t.replies.slice(0, 1000).map((r) => ({ author: str(r?.author, 255) || 'You', text: str(r?.text), time: str(r?.time, 40) })) : [],
  };
};

const cleanStroke = (s: Partial<InkStroke> | undefined, i: number): InkStroke | null => {
  if (!s || !Array.isArray(s.points)) return null;
  const tool = s.tool === 'highlighter' || s.tool === 'marker' ? s.tool : 'pen';
  const points = s.points.slice(0, 20000).filter((p) => Array.isArray(p)).map((p) => [num(p[0]), num(p[1])] as [number, number]);
  if (!points.length) return null;
  const color = /^#[0-9a-f]{6}$/i.test(str(s.color, 7)) ? str(s.color, 7) : '#000000';
  return { id: str(s.id, 64) || `ink${i + 1}`, tool, color, width: Math.min(100, Math.max(0.5, num(s.width) || 2)), points };
};

/** Restores the ClerdoGrid objects part into a workbook just read. Bad data is skipped, never trusted. */
export const applyObjectsPart = (wb: Workbook, text: string | undefined): void => {
  if (!text || text.length > 64 * 1024 * 1024) return;
  let doc: Partial<ObjectsDoc>;
  try { doc = JSON.parse(text) as Partial<ObjectsDoc>; } catch { return; }
  if (!doc || typeof doc !== 'object' || !doc.sheets || typeof doc.sheets !== 'object') return;
  for (const [idx, o] of Object.entries(doc.sheets)) {
    const s = wb.sheets[Number(idx)];
    if (!s || !o || typeof o !== 'object') continue;
    if (Array.isArray(o.comments)) {
      const map = new Map<number, CommentThread>();
      o.comments.slice(0, 100000).forEach((c, i) => {
        const p = parseAddr(str(c?.ref, 12));
        const t = cleanThread(c?.thread, i);
        if (p && t) map.set(cellKey(p.row, p.col), t);
      });
      if (map.size) s.comments = map;
    }
    if (Array.isArray(o.threadNotes)) for (const ref of o.threadNotes) {
      const p = parseAddr(str(ref, 12));
      if (p && s.comments?.has(cellKey(p.row, p.col))) s.notes.delete(cellKey(p.row, p.col));
    }
    if (Array.isArray(o.checkboxes)) for (const ref of o.checkboxes.slice(0, 1_000_000)) {
      const p = parseAddr(str(ref, 12));
      if (!p) continue;
      const k = cellKey(p.row, p.col);
      const c = s.cells.get(k) ?? { v: false };
      c.s = internStyle(wb, { ...(wb.styles[c.s ?? 0] ?? {}), checkbox: true });
      if (c.v === null) c.v = false;
      s.cells.set(k, c);
    }
    if (Array.isArray(o.ink)) {
      const ink = o.ink.slice(0, 50000).map(cleanStroke).filter((x): x is InkStroke => !!x);
      if (ink.length) s.ink = ink;
    }
    if (Array.isArray(o.editRangeTitles) && s.editRanges) o.editRangeTitles.forEach((t, i) => { if (s.editRanges![i] && typeof t === 'string') s.editRanges![i]!.title = t.slice(0, 255); });
  }
};
