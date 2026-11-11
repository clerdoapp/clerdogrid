// Page setup parts of a worksheet (ECMA-376 Part 1, 18.3.1): printOptions, pageMargins, pageSetup,
// headerFooter, rowBreaks/colBreaks, the fitToPage flag, the Print_Area / Print_Titles names and
// header/footer pictures (legacyDrawingHF, a VML part).
import { colIndex, colName, type HeaderPicture, type PageSetup } from '../model/workbook';
import { attr, boolAttr, decodeXstr, el, escXstr, kid, kids, numAttr, parseRoot, esc, type XNode } from './xml';

const num = (n: XNode | undefined, a: string): number | undefined => {
  const v = numAttr(n, a);
  return v !== undefined && Number.isFinite(v) ? v : undefined;
};

const ERRORS: Record<string, PageSetup['errors']> = { displayed: 'displayed', blank: 'blank', dash: 'dash', NA: 'NA' };
const COMMENTS: Record<string, PageSetup['cellComments']> = { none: 'none', atEnd: 'atEnd', asDisplayed: 'asDisplayed' };

const hfText = (n: XNode | undefined): string | undefined => {
  if (!n) return undefined;
  const t = decodeXstr(n.text);
  return t ? t.slice(0, 2000) : undefined;
};

/** Reads the print settings of a worksheet node. `pr` is its sheetPr element. */
export const readPageSetup = (ws: XNode, pr: XNode | undefined): PageSetup | undefined => {
  const ps = kid(ws, 'pageSetup');
  const pm = kid(ws, 'pageMargins');
  const po = kid(ws, 'printOptions');
  const hf = kid(ws, 'headerFooter');
  const rb = kid(ws, 'rowBreaks');
  const cb = kid(ws, 'colBreaks');
  if (!ps && !pm && !po && !hf && !rb && !cb) return undefined;
  const s: PageSetup = {};
  if (ps) {
    const o = attr(ps, 'orientation');
    if (o === 'portrait' || o === 'landscape') s.orientation = o;
    const paper = num(ps, 'paperSize');
    if (paper && paper > 0 && paper < 256) s.paper = paper;
    const pw = attr(ps, 'paperWidth'), ph = attr(ps, 'paperHeight');
    if (pw && /^\d+(\.\d+)?(mm|cm|in|pt|pc|pi)$/.test(pw)) s.paperWidth = pw;
    if (ph && /^\d+(\.\d+)?(mm|cm|in|pt|pc|pi)$/.test(ph)) s.paperHeight = ph;
    const scale = num(ps, 'scale');
    if (scale !== undefined && scale !== 100) s.scale = Math.max(10, Math.min(400, Math.round(scale)));
    if (boolAttr(kid(pr, 'pageSetUpPr'), 'fitToPage')) {
      s.fitWidth = Math.max(0, Math.min(32767, num(ps, 'fitToWidth') ?? 1));
      s.fitHeight = Math.max(0, Math.min(32767, num(ps, 'fitToHeight') ?? 1));
    }
    const first = num(ps, 'firstPageNumber');
    if (first !== undefined && boolAttr(ps, 'useFirstPageNumber')) s.firstPageNumber = Math.round(first);
    if (attr(ps, 'pageOrder') === 'overThenDown') s.order = 'overThenDown';
    if (boolAttr(ps, 'blackAndWhite')) s.blackAndWhite = true;
    if (boolAttr(ps, 'draft')) s.draft = true;
    const cc = COMMENTS[attr(ps, 'cellComments') ?? ''];
    if (cc && cc !== 'none') s.cellComments = cc;
    const er = ERRORS[attr(ps, 'errors') ?? ''];
    if (er && er !== 'displayed') s.errors = er;
    const hd = num(ps, 'horizontalDpi'), vd = num(ps, 'verticalDpi'), copies = num(ps, 'copies');
    if (hd && hd !== 600) s.horizontalDpi = hd;
    if (vd && vd !== 600) s.verticalDpi = vd;
    if (copies && copies > 1) s.copies = Math.min(999, Math.round(copies));
  }
  if (pm) {
    const m = (a: string, d: number) => { const v = num(pm, a); return v !== undefined && v >= 0 && v < 50 ? v : d; };
    const mg = { left: m('left', 0.7), right: m('right', 0.7), top: m('top', 0.75), bottom: m('bottom', 0.75), header: m('header', 0.3), footer: m('footer', 0.3) };
    if ((Object.keys(mg) as (keyof typeof mg)[]).some((k) => Math.abs(mg[k] - DEFAULT_MARGINS[k]) > 1e-9)) s.margins = mg;
  }
  if (po) {
    if (boolAttr(po, 'gridLines')) s.gridlines = true;
    if (boolAttr(po, 'headings')) s.headings = true;
    if (boolAttr(po, 'horizontalCentered')) s.centerH = true;
    if (boolAttr(po, 'verticalCentered')) s.centerV = true;
  }
  if (hf) {
    const set = (k: 'header' | 'footer' | 'evenHeader' | 'evenFooter' | 'firstHeader' | 'firstFooter', tag: string) => { const t = hfText(kid(hf, tag)); if (t) s[k] = t; };
    set('header', 'oddHeader'); set('footer', 'oddFooter');
    set('evenHeader', 'evenHeader'); set('evenFooter', 'evenFooter');
    set('firstHeader', 'firstHeader'); set('firstFooter', 'firstFooter');
    if (boolAttr(hf, 'differentOddEven')) s.differentOddEven = true;
    if (boolAttr(hf, 'differentFirst')) s.differentFirst = true;
    if (attr(hf, 'scaleWithDoc') !== undefined && !boolAttr(hf, 'scaleWithDoc', true)) s.scaleWithDoc = false;
    if (attr(hf, 'alignWithMargins') !== undefined && !boolAttr(hf, 'alignWithMargins', true)) s.alignWithMargins = false;
  }
  const breaks = (n: XNode | undefined, max: number): number[] | undefined => {
    const ids = kids(n, 'brk').filter((b) => attr(b, 'man') === undefined || boolAttr(b, 'man')).map((b) => num(b, 'id') ?? 0).filter((v) => v > 0 && v < max);
    return ids.length ? [...new Set(ids)].sort((a, b) => a - b).slice(0, 1026) : undefined;
  };
  const r = breaks(rb, 1048576), c = breaks(cb, 16384);
  if (r) s.rowBreaks = r;
  if (c) s.colBreaks = c;
  return Object.keys(s).length ? s : undefined;
};

const isFit = (s: PageSetup): boolean => s.fitWidth !== undefined || s.fitHeight !== undefined;

/** sheetPr child for fit-to-page. */
export const pageSetUpPrXml = (s: PageSetup | undefined): string => (s && isFit(s) ? '<pageSetUpPr fitToPage="1"/>' : '');

export const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 } as const;

/** printOptions + pageMargins + pageSetup + headerFooter + rowBreaks + colBreaks, in schema order. */
export const pageSetupXml = (s: PageSetup | undefined): string => {
  let x = '';
  if (s && (s.gridlines || s.headings || s.centerH || s.centerV)) {
    x += el('printOptions', { horizontalCentered: s.centerH || undefined, verticalCentered: s.centerV || undefined, headings: s.headings || undefined, gridLines: s.gridlines || undefined });
  }
  x += el('pageMargins', { ...DEFAULT_MARGINS, ...s?.margins });
  if (s) {
    const fit = isFit(s);
    const a: Record<string, string | number | boolean | undefined> = {
      paperSize: s.paper && s.paper !== 1 ? s.paper : undefined,
      paperHeight: s.paperHeight, paperWidth: s.paperWidth,
      scale: s.scale !== undefined && s.scale !== 100 ? s.scale : undefined,
      firstPageNumber: s.firstPageNumber,
      fitToWidth: fit && s.fitWidth !== 1 ? s.fitWidth ?? 1 : undefined,
      fitToHeight: fit && s.fitHeight !== 1 ? s.fitHeight ?? 1 : undefined,
      pageOrder: s.order === 'overThenDown' ? 'overThenDown' : undefined,
      orientation: s.orientation,
      blackAndWhite: s.blackAndWhite || undefined,
      draft: s.draft || undefined,
      cellComments: s.cellComments && s.cellComments !== 'none' ? s.cellComments : undefined,
      useFirstPageNumber: s.firstPageNumber !== undefined || undefined,
      errors: s.errors && s.errors !== 'displayed' ? s.errors : undefined,
      horizontalDpi: s.horizontalDpi, verticalDpi: s.verticalDpi, copies: s.copies,
    };
    if (Object.values(a).some((v) => v !== undefined)) x += el('pageSetup', a);
    const parts: [string, string | undefined][] = [['oddHeader', s.header], ['oddFooter', s.footer], ['evenHeader', s.evenHeader], ['evenFooter', s.evenFooter], ['firstHeader', s.firstHeader], ['firstFooter', s.firstFooter]];
    const inner = parts.filter(([, t]) => t).map(([tag, t]) => `<${tag}>${escXstr(t!)}</${tag}>`).join('');
    if (inner || s.differentOddEven || s.differentFirst || s.scaleWithDoc === false || s.alignWithMargins === false) {
      x += el('headerFooter', {
        differentOddEven: s.differentOddEven || undefined, differentFirst: s.differentFirst || undefined,
        scaleWithDoc: s.scaleWithDoc === false ? false : undefined, alignWithMargins: s.alignWithMargins === false ? false : undefined,
      }, inner);
    }
    const brk = (tag: string, ids: number[] | undefined, max: number) => {
      const list = [...new Set(ids ?? [])].filter((v) => v > 0 && v <= max).sort((p, q) => p - q);
      if (!list.length) return '';
      return `<${tag} count="${list.length}" manualBreakCount="${list.length}">${list.map((id) => el('brk', { id, max, man: true })).join('')}</${tag}>`;
    };
    x += brk('rowBreaks', s.rowBreaks, 16383);
    x += brk('colBreaks', s.colBreaks, 1048575);
  }
  return x;
};

// ---------- Print titles ----------

/** "$1:$2" (rows) / "$A:$B" (columns) parts of a Print_Titles formula, without the sheet name. */
export const parsePrintTitles = (formula: string): { rows?: string; cols?: string } => {
  const out: { rows?: string; cols?: string } = {};
  for (const part of formula.split(',')) {
    const ref = part.slice(part.lastIndexOf('!') + 1).replace(/\$/g, '').trim();
    let m = /^(\d+):(\d+)$/.exec(ref);
    if (m) { out.rows = `${Math.min(+m[1]!, +m[2]!)}:${Math.max(+m[1]!, +m[2]!)}`; continue; }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(ref);
    if (m) {
      const a = colIndex(m[1]!), b = colIndex(m[2]!);
      out.cols = `${colName(Math.min(a, b))}:${colName(Math.max(a, b))}`;
    }
  }
  return out;
};

/** Print_Titles formula for a sheet reference such as "Sheet1" or "'My sheet'". */
export const printTitlesFormula = (sheetRef: string, s: PageSetup | undefined): string | undefined => {
  const parts: string[] = [];
  const c = s?.printTitleCols && /^([A-Z]{1,3}):([A-Z]{1,3})$/i.exec(s.printTitleCols);
  if (c) parts.push(`${sheetRef}!$${c[1]!.toUpperCase()}:$${c[2]!.toUpperCase()}`);
  const r = s?.printTitleRows && /^(\d+):(\d+)$/.exec(s.printTitleRows);
  if (r) parts.push(`${sheetRef}!$${r[1]}:$${r[2]}`);
  return parts.length ? parts.join(',') : undefined;
};

// ---------- Header and footer pictures (VML) ----------

const SLOT = /^[LCR][HF](FIRST|EVEN)?$/;
const toPt = (v: string | undefined): number | undefined => {
  const m = v && /^\s*(-?\d+(?:\.\d+)?)\s*(pt|in|px|mm|cm)?\s*$/.exec(v);
  if (!m) return undefined;
  const n = +m[1]!;
  switch (m[2]) { case 'in': return n * 72; case 'px': return n * 0.75; case 'mm': return (n * 72) / 25.4; case 'cm': return (n * 72) / 2.54; default: return n; }
};
const MIME_OF: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

/** Reads the pictures of a legacyDrawingHF part. `image(relId)` returns the bytes and path of a related image. */
export const readHeaderPictures = (vml: string, image: (relId: string) => { data: Uint8Array; path: string } | undefined): Record<string, HeaderPicture> | undefined => {
  let root: XNode;
  try { root = parseRoot(vml); } catch { return undefined; }
  const out: Record<string, HeaderPicture> = {};
  const walk = (n: XNode) => {
    for (const c of n.children) {
      if (c.name === 'shape') {
        const id = (attr(c, 'id') ?? '').toUpperCase();
        const data = kid(c, 'imagedata');
        const rel = attr(data, 'o:relid') ?? attr(data, 'r:id') ?? attr(data, 'relid');
        if (SLOT.test(id) && rel) {
          const img = image(rel);
          const ext = img?.path.slice(img.path.lastIndexOf('.') + 1).toLowerCase() ?? '';
          const style = attr(c, 'style') ?? '';
          const get = (k: string) => new RegExp(`(?:^|;)\\s*${k}\\s*:\\s*([^;]+)`, 'i').exec(style)?.[1];
          if (img && MIME_OF[ext]) out[id] = { data: img.data, mime: MIME_OF[ext]!, width: toPt(get('width')) ?? 72, height: toPt(get('height')) ?? 36 };
        }
      } else walk(c);
    }
  };
  walk(root);
  return Object.keys(out).length ? out : undefined;
};

/** VML for header/footer pictures; `relIds[i]` is the relationship id of pictures[i]'s image. */
export const headerPicturesVml = (pictures: [string, HeaderPicture][], relIds: string[], block: number): string => {
  const shapes = pictures.map(([slot, p], i) => `<v:shape id="${esc(slot)}" o:spid="_x0000_s${block * 1024 + i + 1}" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:${+p.width.toFixed(2)}pt;height:${+p.height.toFixed(2)}pt;z-index:${i + 1}">`
    + `<v:imagedata o:relid="${relIds[i]}" o:title="${esc(slot)}"/><o:lock v:ext="edit" rotation="t"/></v:shape>`).join('');
  return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
    + `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${block}"/></o:shapelayout>`
    + '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/>'
    + '<v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>'
    + '<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'
    + shapes + '</xml>';
};

export const HF_IMG_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };
