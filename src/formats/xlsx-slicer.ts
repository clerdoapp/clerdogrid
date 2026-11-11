// Slicer and timeline parts (MS-XLSX extensions to ECMA-376): x14 slicer caches and slicers, x15 table
// slicer caches, x15 timeline caches and timelines, their worksheet and workbook extension lists, and
// the drawing anchors that place them on the sheet.
import { fromSerial, toSerial } from '../engine/date';
import { cellKey, type TableDef, type Workbook } from '../model/workbook';
import { formatValue } from './numfmt';
import { isError } from '../engine/types';
import type { PivotTable, Slicer, SlicerCache } from '../model/pivot';
import { anchorXml, readAnchor } from './xlsx-chart';
import { attr, boolAttr, el, esc, kid, kids, numAttr, type XNode } from './xml';
import type { PartSource } from './xlsx-pivot';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const NS_X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

export const SLICER_RT = {
  cache: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
  slicer: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
  timelineCache: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
  timeline: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
};
export const SLICER_CT = {
  cache: 'application/vnd.ms-excel.slicerCache+xml',
  slicer: 'application/vnd.ms-excel.slicer+xml',
  timelineCache: 'application/vnd.ms-excel.timelineCache+xml',
  timeline: 'application/vnd.ms-excel.timeline+xml',
};
const EXT = {
  wbSlicerCaches: '{BBE1A952-AA13-448e-AADC-164F8A28A991}',
  wbTableSlicerCaches: '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
  wbTimelineCaches: '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}',
  wsSlicers: '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
  wsTableSlicers: '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
  wsTimelines: '{7E03D99C-DC04-49d9-9315-930204A7B6E9}',
  tableSlicerCache: '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}',
  hideNoData: '{470722E0-AACD-4C17-9CDC-17EF765DBC7E}',
  clerdo: '{4B7C1E92-6D3A-4F58-8C2E-1A9D5F3B7E60}',
};
const NS_CLERDO = 'https://clerdo.app/ns/grid/slicer/1';
const EMU_PER_PX = 9525;
const LEVELS = ['years', 'quarters', 'months', 'days'] as const;

const isoOf = (serial: number, date1904: boolean): string => {
  const d = fromSerial(Math.floor(serial), date1904);
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  return `${p(d.year, 4)}-${p(d.month)}-${p(d.day)}T00:00:00`;
};
const serialOf = (iso: string | undefined, date1904: boolean): number | undefined => {
  const m = iso ? /^(\d{4})-(\d{2})-(\d{2})/.exec(iso) : null;
  return m ? toSerial(+m[1]!, +m[2]!, +m[3]!, date1904) : undefined;
};

// ---------------------------------------------------------------- writing

export interface SlicerSink { put: (path: string, text: string, ct?: string) => void }

/** Writes slicers and timelines. Sheets call `sheetParts`; the workbook calls `workbookParts`. */
export class SlicerWriter {
  private n = { slicer: 0, timeline: 0 };
  /** Table name to file table id (1-based, in sheet order, as the table writer numbers them). */
  private tableIds = new Map<string, { id: number; sheet: number }>();

  constructor(private readonly wb: Workbook, private readonly sink: SlicerSink) {
    let id = 0;
    wb.sheets.forEach((s, si) => s.tables.forEach((t) => this.tableIds.set(t.name.toLowerCase(), { id: ++id, sheet: si })));
  }

  get has(): boolean { return !!this.wb.slicerCaches?.length; }

  private caches(): SlicerCache[] {
    return (this.wb.slicerCaches ?? []).filter((c) => (c.source.kind === 'table' ? this.tableIds.has(c.source.table.toLowerCase()) : this.wb.pivotCaches?.some((p) => p.id === (c.source as { cacheId: number }).cacheId)));
  }

  /**
   * The table's autoFilter columns for its slicers: the items still shown, which is where table
   * slicers keep their selection (the x15 tableSlicerCache points at the column).
   */
  tableFilterXml(t: TableDef): string {
    const at = this.tableIds.get(t.name.toLowerCase());
    if (!at) return '';
    const caches = this.caches().filter((c) => c.source.kind === 'table' && c.source.table.toLowerCase() === t.name.toLowerCase() && c.hidden.length);
    const byCol = new Map<number, Set<string>>();
    for (const c of caches) {
      const col = (c.source as { column: number }).column;
      const set = byCol.get(col) ?? new Set<string>();
      for (const h of c.hidden) set.add(h);
      byCol.set(col, set);
    }
    return [...byCol].sort((a, b) => a[0] - b[0]).map(([col, hidden]) => {
      const shown = columnTexts(this.wb, at.sheet, t, col).filter((x) => !hidden.has(x));
      const blank = shown.includes('');
      return el('filterColumn', { colId: col }, el('filters', { blank: blank || undefined }, shown.filter((x) => x !== '').map((v) => el('filter', { val: v })).join('')));
    }).join('');
  }

  /** Defined names other apps expect for each cache (formula #N/A). */
  definedNames(): { name: string; formula: string }[] {
    return this.caches().map((c) => ({ name: c.name, formula: '#N/A' }));
  }

  /** Writes slicer parts for a sheet; returns its extension list items and drawing anchors. */
  sheetParts(si: number, addRel: (type: string, target: string) => string): { ext: string; anchors: string[] } {
    const s = this.wb.sheets[si]!;
    const list = (s.slicers ?? []).filter((x) => this.caches().some((c) => c.name === x.cache));
    if (!list.length) return { ext: '', anchors: [] };
    const cacheOf = (x: Slicer) => this.caches().find((c) => c.name === x.cache)!;
    const pivotSlicers = list.filter((x) => x.kind === 'slicer' && cacheOf(x).source.kind === 'pivot');
    const tableSlicers = list.filter((x) => x.kind === 'slicer' && cacheOf(x).source.kind === 'table');
    const timelines = list.filter((x) => x.kind === 'timeline');
    let ext = '';
    const slicerPart = (items: Slicer[]): string => {
      const n = ++this.n.slicer;
      const body = items.map((x) => el('slicer', {
        name: x.name, cache: x.cache, caption: x.caption, columnCount: x.columns > 1 ? x.columns : undefined,
        showCaption: x.showCaption ? undefined : false, style: x.style || undefined, rowHeight: Math.round(x.buttonHeight * EMU_PER_PX),
      })).join('');
      this.sink.put(`xl/slicers/slicer${n}.xml`, `${XML_HEAD}<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">${body}</slicers>`, SLICER_CT.slicer);
      return addRel(SLICER_RT.slicer, `../slicers/slicer${n}.xml`);
    };
    if (pivotSlicers.length) ext += `<ext uri="${EXT.wsSlicers}" xmlns:x14="${NS_X14}"><x14:slicerList><x14:slicer r:id="${slicerPart(pivotSlicers)}"/></x14:slicerList></ext>`;
    if (tableSlicers.length) ext += `<ext uri="${EXT.wsTableSlicers}" xmlns:x15="${NS_X15}"><x14:slicerList xmlns:x14="${NS_X14}"><x14:slicer r:id="${slicerPart(tableSlicers)}"/></x14:slicerList></ext>`;
    if (timelines.length) {
      const n = ++this.n.timeline;
      const body = timelines.map((x) => {
        const c = cacheOf(x);
        const lvl = LEVELS.indexOf(x.level ?? c.timeline?.level ?? 'months');
        return el('timeline', {
          name: x.name, cache: x.cache, caption: x.caption, level: lvl, selectionLevel: lvl,
          scrollPosition: c.timeline?.start !== undefined ? isoOf(c.timeline.start, this.wb.date1904) : undefined,
          showHeader: x.showCaption ? undefined : false, showSelectionLabel: x.showSelectionLabel === false ? false : undefined,
          showTimeLevel: x.showTimeLevel === false ? false : undefined, showHorizontalScrollbar: x.showScrollbar === false ? false : undefined,
          style: x.style || undefined,
        });
      }).join('');
      this.sink.put(`xl/timelines/timeline${n}.xml`, `${XML_HEAD}<timelines xmlns="${NS_X15}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">${body}</timelines>`, SLICER_CT.timeline);
      const rid = addRel(SLICER_RT.timeline, `../timelines/timeline${n}.xml`);
      ext += `<ext uri="${EXT.wsTimelines}" xmlns:x15="${NS_X15}"><x15:timelineRefs><x15:timelineRef r:id="${rid}"/></x15:timelineRefs></ext>`;
    }
    let shapeId = 3000;
    const anchors = list.map((x) => anchorXml(x.anchor, s, frameXml(x, ++shapeId)));
    return { ext, anchors };
  }

  /** Writes cache parts; returns the workbook extension list items. */
  workbookParts(addRel: (type: string, target: string) => string): string {
    const pivotRefs: string[] = [], tableRefs: string[] = [], timelineRefs: string[] = [];
    this.caches().forEach((c, i) => {
      const n = i + 1;
      if (c.timeline) {
        this.sink.put(`xl/timelineCaches/timelineCache${n}.xml`, this.timelineCacheXml(c), SLICER_CT.timelineCache);
        timelineRefs.push(addRel(SLICER_RT.timelineCache, `timelineCaches/timelineCache${n}.xml`));
        return;
      }
      this.sink.put(`xl/slicerCaches/slicerCache${n}.xml`, this.slicerCacheXml(c), SLICER_CT.cache);
      const rid = addRel(SLICER_RT.cache, `slicerCaches/slicerCache${n}.xml`);
      (c.source.kind === 'table' ? tableRefs : pivotRefs).push(rid);
    });
    let ext = '';
    if (pivotRefs.length) ext += `<ext uri="${EXT.wbSlicerCaches}" xmlns:x14="${NS_X14}"><x14:slicerCaches>${pivotRefs.map((r) => `<x14:slicerCache r:id="${r}"/>`).join('')}</x14:slicerCaches></ext>`;
    if (tableRefs.length) ext += `<ext uri="${EXT.wbTableSlicerCaches}" xmlns:x15="${NS_X15}"><x15:slicerCaches xmlns:x14="${NS_X14}">${tableRefs.map((r) => `<x14:slicerCache r:id="${r}"/>`).join('')}</x15:slicerCaches></ext>`;
    if (timelineRefs.length) ext += `<ext uri="${EXT.wbTimelineCaches}" xmlns:x15="${NS_X15}"><x15:timelineCacheRefs>${timelineRefs.map((r) => `<x15:timelineCacheRef r:id="${r}"/>`).join('')}</x15:timelineCacheRefs></ext>`;
    return ext;
  }

  private pivotTablesOf(c: SlicerCache): string {
    if (c.source.kind !== 'pivot') return '';
    return c.source.pivots.map((p) => {
      const si = this.wb.sheets.findIndex((s) => s.name.toLowerCase() === p.sheet.toLowerCase());
      return si >= 0 && this.wb.sheets[si]!.pivots?.some((x) => x.name === p.name) ? el('pivotTable', { tabId: si + 1, name: p.name }) : '';
    }).join('');
  }

  private slicerCacheXml(c: SlicerCache): string {
    const src = c.source;
    const sortOrder = c.sort === 'descending' ? 'descending' : undefined;
    const crossFilter = c.hideNoData ? undefined : c.showNoDataLast === false ? 'showItemsWithNoData' : undefined;
    if (src.kind === 'table') {
      const t = this.tableIds.get(src.table.toLowerCase())!;
      const table = this.wb.sheets[t.sheet]!.tables.find((x) => x.name.toLowerCase() === src.table.toLowerCase())!;
      const own = c.hidden.length ? `<x:ext uri="${EXT.clerdo}" xmlns:cg="${NS_CLERDO}"><cg:hidden>${c.hidden.map((h) => `<cg:i v="${esc(h)}"/>`).join('')}</cg:hidden></x:ext>` : '';
      return `${XML_HEAD}<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x x15" xmlns:x="${NS_MAIN}" xmlns:x15="${NS_X15}" name="${esc(c.name)}" sourceName="${esc(table.columns[src.column]?.name ?? '')}">`
        + `<extLst><x:ext uri="${EXT.tableSlicerCache}" xmlns:x15="${NS_X15}">${el('x15:tableSlicerCache', { tableId: t.id, column: src.column + 1, sortOrder, crossFilter })}</x:ext>${own}</extLst></slicerCacheDefinition>`;
    }
    const cache = this.wb.pivotCaches!.find((p) => p.id === src.cacheId)!;
    const field = cache.fields[src.field];
    const n = field ? (field.group?.range ? field.group.items.length : field.group ? field.group.items.length : field.items.length) : 0;
    const hidden = new Set(c.hidden);
    const items = Array.from({ length: n }, (_, x) => el('i', { x, s: hidden.has(String(x)) ? undefined : true })).join('');
    return `${XML_HEAD}<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(c.name)}" sourceName="${esc(field?.name ?? '')}">`
      + `<pivotTables>${this.pivotTablesOf(c)}</pivotTables>`
      + `<data><tabular pivotCacheId="${src.cacheId}"${sortOrder ? ` sortOrder="${sortOrder}"` : ''}${crossFilter ? ` crossFilter="${crossFilter}"` : ''}><items count="${n}">${items}</items></tabular></data>`
      + (c.hideNoData ? `<extLst><x:ext uri="${EXT.hideNoData}" xmlns:x15="${NS_X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext></extLst>` : '')
      + '</slicerCacheDefinition>';
  }

  private timelineCacheXml(c: SlicerCache): string {
    const src = c.source;
    if (src.kind !== 'pivot') return '';
    const cache = this.wb.pivotCaches!.find((p) => p.id === src.cacheId)!;
    const field = cache.fields[src.field];
    const base = field?.group?.base ?? src.field;
    const di = cache.fields.slice(0, base).filter((f) => f.database !== false).length;
    let lo = Infinity, hi = -Infinity;
    for (const r of cache.records) { const v = r[di]; if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); } }
    const d1904 = this.wb.date1904;
    const tl = c.timeline!;
    const sel = tl.start !== undefined && tl.end !== undefined ? el('selection', { startDate: isoOf(tl.start, d1904), endDate: isoOf(tl.end + 1, d1904) }) : '';
    const bounds = lo <= hi ? el('bounds', { startDate: isoOf(lo, d1904), endDate: isoOf(hi + 1, d1904) }) : '';
    return `${XML_HEAD}<timelineCacheDefinition xmlns="${NS_X15}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(c.name)}" sourceName="${esc(cache.fields[base]?.name ?? '')}">`
      + `<pivotTables>${this.pivotTablesOf(c)}</pivotTables>`
      + `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${src.cacheId}" filterType="${sel ? 'dateBetween' : 'unknown'}">${sel}${bounds}</state>`
      + '</timelineCacheDefinition>';
  }
}

/** Distinct shown texts of a table column, in sheet order (as a slicer lists them). */
export const columnTexts = (wb: Workbook, sheet: number, t: TableDef, col: number): string[] => {
  const s = wb.sheets[sheet]!;
  const r1 = t.range.r1 + (t.headerRow ? 1 : 0), r2 = t.range.r2 - (t.totalsRow ? 1 : 0);
  const c = t.range.c1 + col;
  const out = new Set<string>();
  for (let r = r1; r <= r2; r++) {
    const cell = s.cells.get(cellKey(r, c));
    const v = cell?.v ?? null;
    const sid = cell?.s ?? s.rows.get(r)?.s ?? s.cols.get(c)?.s ?? 0;
    out.add(v === null ? '' : isError(v) ? v.code : formatValue(v, wb.styles[sid]?.numFmt ?? 'General', { date1904: wb.date1904 }).text);
  }
  return [...out];
};

const frameXml = (x: Slicer, id: number): string => {
  const isTimeline = x.kind === 'timeline';
  const req = isTimeline ? 'tsle' : 'a14';
  const nsDecl = isTimeline ? 'xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer"' : 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"';
  const data = isTimeline
    ? `<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2012/timeslicer"><tsle:timeslicer name="${esc(x.name)}"/></a:graphicData>`
    : `<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(x.name)}"/></a:graphicData>`;
  const cx = Math.round(x.anchor.width * EMU_PER_PX), cy = Math.round(x.anchor.height * EMU_PER_PX);
  return `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice ${nsDecl} Requires="${req}">`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(x.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>${data}</a:graphic></xdr:graphicFrame>`
    + `</mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>`
    + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="0C7A4B"/></a:solidFill></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>${isTimeline ? 'Timeline' : 'Slicer'}: ${esc(x.caption)}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>`;
};

// ---------------------------------------------------------------- reading

const extsOf = (n: XNode | undefined): XNode[] => kids(kid(n, 'extLst'), 'ext');
const findDeep = (n: XNode | undefined, local: string): XNode[] => {
  const out: XNode[] = [];
  const walk = (x: XNode) => { if (x.name === local || x.name.endsWith(`:${local}`)) out.push(x); x.children.forEach(walk); };
  if (n) walk(n);
  return out;
};
const rid = (n: XNode): string | undefined => attr(n, 'r:id') ?? attr(n, 'id');

/** Reads slicer and timeline caches and the slicers placed on each sheet. Call after pivot tables are read. */
export const readSlicerParts = (pkg: PartSource, wb: Workbook, wbx: XNode, wbPath: string, sheetPaths: string[]): void => {
  const rels = new Map(pkg.rels(wbPath).map((r) => [r.id, r]));
  const caches: SlicerCache[] = [];
  // tabId is the sheetId from workbook.xml, not the position.
  const sheetIds = kids(kid(wbx, 'sheets'), 'sheet').filter((s) => { const r = rels.get(rid(s) ?? ''); return r && (r.type === 'worksheet' || r.type.endsWith('/worksheet')); }).map((s) => numAttr(s, 'sheetId') ?? 0);
  const pivotByTab = (tabId: number, name: string): { sheet: string; pt: PivotTable } | null => {
    const at = sheetIds.indexOf(tabId);
    const s = wb.sheets[at >= 0 ? at : tabId - 1];
    const pt = s?.pivots?.find((p) => p.name === name);
    return s && pt ? { sheet: s.name, pt } : null;
  };
  const tables: { name: string; sheet: number }[] = [];
  wb.sheets.forEach((s, si) => s.tables.forEach((t) => tables.push({ name: t.name, sheet: si })));
  const d1904 = wb.date1904;
  for (const ext of extsOf(wbx)) {
    const uri = attr(ext, 'uri');
    const refs = uri === EXT.wbSlicerCaches || uri === EXT.wbTableSlicerCaches ? findDeep(ext, 'slicerCache') : uri === EXT.wbTimelineCaches ? findDeep(ext, 'timelineCacheRef') : [];
    for (const r of refs) {
      const rel = rels.get(rid(r) ?? '');
      if (!rel || rel.external) continue;
      const x = pkg.xml(rel.target);
      if (!x) continue;
      try {
        const name = attr(x, 'name') ?? `Slicer_${caches.length + 1}`;
        const sourceName = (attr(x, 'sourceName') ?? '').toLowerCase();
        const pivots = findDeep(kid(x, 'pivotTables') ?? x, 'pivotTable').map((p) => pivotByTab(numAttr(p, 'tabId') ?? 0, attr(p, 'name') ?? '')).filter((p): p is { sheet: string; pt: PivotTable } => !!p);
        const fieldIn = (pt: PivotTable) => wb.pivotCaches?.find((c) => c.id === pt.cacheId)?.fields.findIndex((f) => f.name.toLowerCase() === sourceName) ?? -1;
        if (x.name.endsWith('timelineCacheDefinition')) {
          const p = pivots[0];
          if (!p) continue;
          const state = kid(x, 'state');
          const sel = kid(state, 'selection');
          const start = serialOf(attr(sel, 'startDate'), d1904), end0 = serialOf(attr(sel, 'endDate'), d1904);
          caches.push({ name, source: { kind: 'pivot', cacheId: p.pt.cacheId, field: fieldIn(p.pt), pivots: pivots.map((q) => ({ sheet: q.sheet, name: q.pt.name })) }, hidden: [], sort: 'ascending', timeline: { start, end: end0 !== undefined ? end0 - 1 : undefined, level: 'months' } });
          continue;
        }
        const tsc = findDeep(x, 'tableSlicerCache')[0];
        if (tsc) {
          const t = tables[(numAttr(tsc, 'tableId') ?? 0) - 1];
          if (!t) continue;
          const column = Math.max(0, (numAttr(tsc, 'column') ?? 1) - 1);
          // The selection lives in the table's own filter; ClerdoGrid's copy is only a fallback.
          let hidden = tableFilterHidden(pkg, wb, sheetPaths, t, column);
          if (hidden === null) hidden = findDeep(x, 'hidden').flatMap((h) => h.children.map((i) => attr(i, 'v') ?? ''));
          caches.push({ name, source: { kind: 'table', table: t.name, column }, hidden, sort: attr(tsc, 'sortOrder') === 'descending' ? 'descending' : 'ascending', hideNoData: attr(tsc, 'crossFilter') === 'none' ? false : undefined });
          continue;
        }
        const tab = findDeep(x, 'tabular')[0];
        const p = pivots[0];
        if (!tab || !p) continue;
        const field = fieldIn(p.pt);
        if (field < 0) continue;
        const hidden = findDeep(tab, 'i').filter((i) => !boolAttr(i, 's')).map((i) => String(numAttr(i, 'x') ?? 0));
        caches.push({
          name, source: { kind: 'pivot', cacheId: p.pt.cacheId, field, pivots: pivots.map((q) => ({ sheet: q.sheet, name: q.pt.name })) }, hidden,
          sort: attr(tab, 'sortOrder') === 'descending' ? 'descending' : 'ascending',
          hideNoData: findDeep(x, 'slicerCacheHideItemsWithNoData').length > 0 || undefined,
          showNoDataLast: attr(tab, 'crossFilter') === 'showItemsWithNoData' ? false : undefined,
        });
      } catch { /* skip a broken cache */ }
    }
  }
  if (!caches.length) return;
  wb.slicerCaches = caches;
  // Defined names that only stand for slicer caches.
  const names = new Set(caches.map((c) => c.name.toLowerCase()));
  wb.names = wb.names.filter((n) => !(names.has(n.name.toLowerCase()) && /^#N\/A$/i.test(n.formula.trim())));

  sheetPaths.forEach((sp, si) => {
    if (!sp) return;
    const sheet = wb.sheets[si]!;
    const srels = pkg.rels(sp);
    const parts = srels.filter((r) => !r.external && (r.type === 'slicer' || r.type === 'timeline' || r.type.endsWith('/slicer') || r.type.endsWith('/timeline')));
    if (!parts.length) return;
    // Positions come from the drawing's graphic frames.
    const pos = new Map<string, Slicer['anchor']>();
    const drawing = srels.find((r) => r.type === 'drawing' || r.type.endsWith('/drawing'));
    const dx = drawing ? pkg.xml(drawing.target) : undefined;
    for (const a of dx?.children ?? []) {
      if (!a.name.endsWith('Anchor')) continue;
      const frame = findDeep(a, 'slicer')[0] ?? findDeep(a, 'timeslicer')[0];
      const nm = frame ? attr(frame, 'name') : undefined;
      if (nm) pos.set(nm, readAnchor(a, sheet));
    }
    for (const r of parts) {
      const x = pkg.xml(r.target);
      for (const n of x?.children ?? []) {
        const timeline = n.name.endsWith('timeline');
        if (!timeline && !n.name.endsWith('slicer')) continue;
        const name = attr(n, 'name') ?? '';
        const cache = attr(n, 'cache') ?? '';
        if (!caches.some((c) => c.name === cache)) continue;
        const anchor = pos.get(name) ?? { row: 1, col: 1, dx: 0, dy: 0, width: timeline ? 360 : 180, height: timeline ? 140 : 240 };
        const lvl = numAttr(n, 'level');
        const s: Slicer = {
          name, cache, caption: attr(n, 'caption') ?? name, kind: timeline ? 'timeline' : 'slicer', anchor,
          columns: numAttr(n, 'columnCount') ?? 1, buttonHeight: Math.max(16, Math.round((numAttr(n, 'rowHeight') ?? 241300) / EMU_PER_PX)),
          style: attr(n, 'style') ?? '', showCaption: timeline ? attr(n, 'showHeader') !== '0' : attr(n, 'showCaption') !== '0',
        };
        if (timeline) {
          s.level = LEVELS[lvl ?? 2] ?? 'months';
          if (attr(n, 'showSelectionLabel') === '0') s.showSelectionLabel = false;
          if (attr(n, 'showTimeLevel') === '0') s.showTimeLevel = false;
          if (attr(n, 'showHorizontalScrollbar') === '0') s.showScrollbar = false;
          const c = caches.find((q) => q.name === cache);
          if (c?.timeline) c.timeline.level = s.level;
        }
        sheet.slicers = [...(sheet.slicers ?? []), s];
      }
    }
  });
};

/** Items a table's autoFilter leaves out of one column, or null when the column isn't filtered. */
const tableFilterHidden = (pkg: PartSource, wb: Workbook, sheetPaths: string[], t: { name: string; sheet: number }, column: number): string[] | null => {
  const sp = sheetPaths[t.sheet];
  const def = wb.sheets[t.sheet]?.tables.find((x) => x.name === t.name);
  if (!sp || !def) return null;
  for (const r of pkg.rels(sp)) {
    if (r.external || !(r.type === 'table' || r.type.endsWith('/table'))) continue;
    const tx = pkg.xml(r.target);
    if (!tx || (attr(tx, 'name') ?? attr(tx, 'displayName')) !== t.name) continue;
    const fc = kids(kid(tx, 'autoFilter'), 'filterColumn').find((f) => numAttr(f, 'colId') === column);
    const filters = kid(fc, 'filters');
    if (!filters) return null;
    const shown = new Set(kids(filters, 'filter').map((f) => attr(f, 'val') ?? ''));
    if (boolAttr(filters, 'blank')) shown.add('');
    const lower = new Set([...shown].map((v) => v.toLowerCase()));
    return columnTexts(wb, t.sheet, def, column).filter((v) => !lower.has(v.toLowerCase()));
  }
  return null;
};

export const SLICER_PARTS = /^xl\/(slicers|slicerCaches|timelines|timelineCaches)\//i;
export { NS_R };
