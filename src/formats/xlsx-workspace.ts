// Workbook and sheet protection, custom views, named sheet views, split panes and ignored errors
// in .xlsx (ECMA-376 Part 1: 18.2.29 workbookProtection, 18.3.1.85 sheetProtection, 18.2.4/18.3.1.24
// custom views, 18.3.1.66 pane, 18.3.1.51 ignoredErrors) plus the named sheet views part
// (namespace http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews).
import {
  addr, parseAddr, parseRange, rangeAddr,
  type AutoFilter, type CfOperator, type CustomView, type ErrorCheckType, type FilterColumn, type NamedSheetView, type PasswordHash, type Sheet, type Workbook,
} from '../model/workbook';
import { attr, boolAttr, el, esc, kid, kids, numAttr, parseXml, XML_HEAD, type XNode } from './xml';

export const NSV_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews';
export const NSV_REL = 'http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView';
export const NSV_CT = 'application/vnd.ms-excel.namedsheetviews+xml';
export const WORKSPACE_REL = 'https://schemas.clerdo.app/grid/2026/relationships/workspace';
export const WORKSPACE_PATH = 'xl/clerdogrid/workspace.json';
export const WORKSPACE_CT = 'application/vnd.clerdo.grid.workspace+json';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const guidRe = /^\{[0-9A-Fa-f-]{36}\}$/;
export const newGuid = (): string => {
  const b = new Uint8Array(16);
  globalThis.crypto.getRandomValues(b);
  b[6] = (b[6]! & 0x0f) | 0x40; b[8] = (b[8]! & 0x3f) | 0x80;
  const h = [...b].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
};

// ---------- Protection ----------

const hashAttrs = (h: PasswordHash | undefined, prefix = ''): Record<string, string | number | undefined> => {
  if (!h) return {};
  const k = (n: string) => (prefix ? prefix + n[0]!.toUpperCase() + n.slice(1) : n);
  return { [k('algorithmName')]: h.algorithmName, [k('hashValue')]: h.hashValue, [k('saltValue')]: h.saltValue, [k('spinCount')]: h.spinCount };
};

const readHash = (n: XNode, prefix = ''): PasswordHash | undefined => {
  const k = (x: string) => (prefix ? prefix + x[0]!.toUpperCase() + x.slice(1) : x);
  const hashValue = attr(n, k('hashValue')), saltValue = attr(n, k('saltValue'));
  if (!hashValue || !saltValue) return undefined;
  return { algorithmName: attr(n, k('algorithmName')) ?? 'SHA-512', hashValue, saltValue, spinCount: Math.max(0, Math.min(10_000_000, numAttr(n, k('spinCount')) ?? 100000)) };
};

export const workbookProtectionXml = (wb: Workbook): string => {
  const p = wb.protection;
  if (!p || (!p.structure && !p.windows)) return '';
  return el('workbookProtection', { workbookPassword: p.password, ...hashAttrs(p.hash, 'workbook'), lockStructure: p.structure || undefined, lockWindows: p.windows || undefined });
};

export const readWorkbookProtection = (wbx: XNode, wb: Workbook): void => {
  const n = kid(wbx, 'workbookProtection');
  if (!n) return;
  const structure = boolAttr(n, 'lockStructure'), windows = boolAttr(n, 'lockWindows');
  if (!structure && !windows) return;
  const p: NonNullable<Workbook['protection']> = { structure };
  if (windows) p.windows = true;
  const pw = attr(n, 'workbookPassword');
  if (pw && /^[0-9A-Fa-f]{1,4}$/.test(pw)) p.password = pw.toUpperCase().padStart(4, '0');
  const h = readHash(n, 'workbook');
  if (h) p.hash = h;
  wb.protection = p;
};

/** Attributes for a sheet's sheetProtection element: legacy hash, modern hash and option flags. */
export const sheetProtectionAttrs = (s: Sheet): Record<string, string | number | boolean | undefined> | null => {
  const p = s.protection;
  if (!p?.sheet) return null;
  return { password: p.password, ...hashAttrs(p.hash), sheet: true, ...p.options };
};

export const readSheetProtectionHash = (n: XNode | undefined): PasswordHash | undefined => (n ? readHash(n) : undefined);

// ---------- Filter columns (shared by named sheet views and custom views) ----------

const filterColumnInner = (c: FilterColumn, pfx: string): string => {
  const e = (name: string, a: Record<string, string | number | boolean | undefined>, inner?: string) => el(pfx + name, a, inner);
  if (c.values || c.blanks) return e('filters', { blank: c.blanks || undefined }, (c.values ?? []).map((val) => e('filter', { val })).join(''));
  if (c.custom) return e('customFilters', { and: c.custom.and || undefined }, c.custom.rules.map((r) => e('customFilter', { operator: r.operator === 'equal' ? undefined : r.operator, val: r.value })).join(''));
  if (c.top10) return e('top10', { top: c.top10.top ? undefined : false, percent: c.top10.percent || undefined, val: c.top10.value });
  return '';
};

const readFilterColumn = (fc: XNode, col: number): FilterColumn => {
  const c: FilterColumn = { col };
  const filters = kid(fc, 'filters');
  if (filters) {
    c.values = kids(filters, 'filter').map((f) => attr(f, 'val') ?? '');
    if (boolAttr(filters, 'blank')) c.blanks = true;
  }
  const custom = kid(fc, 'customFilters');
  if (custom) {
    c.custom = { rules: kids(custom, 'customFilter').map((f) => ({ operator: (attr(f, 'operator') ?? 'equal') as CfOperator, value: attr(f, 'val') ?? '' })) };
    if (boolAttr(custom, 'and')) c.custom.and = true;
  }
  const top = kid(fc, 'top10');
  if (top) c.top10 = { top: attr(top, 'top') !== '0', percent: boolAttr(top, 'percent'), value: numAttr(top, 'val') ?? 10 };
  return c;
};

const autoFilterXml = (af: AutoFilter): string => {
  const cols = af.columns.map((c) => el('filterColumn', { colId: c.col }, filterColumnInner(c, ''))).join('');
  const sort = af.sort?.length
    ? el('sortState', { ref: rangeAddr({ ...af.range, r1: af.range.r1 + 1 }) }, af.sort.map((srt) => el('sortCondition', {
      descending: srt.descending || undefined, ref: rangeAddr({ r1: af.range.r1 + 1, r2: af.range.r2, c1: af.range.c1 + srt.col, c2: af.range.c1 + srt.col }),
    })).join(''))
    : '';
  return el('autoFilter', { ref: rangeAddr(af.range) }, cols + sort);
};

const readAutoFilterNode = (af: XNode | undefined): AutoFilter | undefined => {
  const range = af && parseRange(attr(af, 'ref') ?? '');
  if (!af || !range) return undefined;
  const out: AutoFilter = { range, columns: kids(af, 'filterColumn').map((fc) => readFilterColumn(fc, numAttr(fc, 'colId') ?? 0)) };
  const sort = kids(kid(af, 'sortState'), 'sortCondition').map((s) => {
    const r = parseRange(attr(s, 'ref') ?? '');
    return { col: r ? r.c1 - range.c1 : 0, ...(boolAttr(s, 'descending') ? { descending: true } : {}) };
  });
  if (sort.length) out.sort = sort;
  return out;
};

// ---------- Named sheet views part ----------

export const namedSheetViewsXml = (s: Sheet): string | null => {
  const views = s.sheetViews?.filter((v) => v.name.trim());
  if (!views?.length) return null;
  const body = views.map((v) => {
    const id = v.id && guidRe.test(v.id) ? v.id : newGuid();
    v.id = id;
    let inner = '';
    if (v.filter) {
      const af = v.filter;
      const cols = af.columns.map((c) => el('columnFilter', { colId: c.col, id: newGuid() }, el('filter', { colId: c.col }, filterColumnInner(c, 'x:')))).join('');
      const sort = af.sort?.length
        ? `<sortRules>${af.sort.map((srt) => el('sortRule', { colId: srt.col, id: newGuid() }, el('sortCondition', {
          descending: srt.descending || undefined, ref: rangeAddr({ r1: af.range.r1 + 1, r2: af.range.r2, c1: af.range.c1 + srt.col, c2: af.range.c1 + srt.col }),
        }))).join('')}</sortRules>`
        : '';
      inner = el('nsvFilter', { filterId: newGuid(), ref: rangeAddr(af.range) }, cols + sort);
    }
    return el('namedSheetView', { name: v.name, id }, inner);
  }).join('');
  return `${XML_HEAD}<namedSheetViews xmlns="${NSV_NS}" xmlns:x="${NS_MAIN}">${body}</namedSheetViews>`;
};

export const readNamedSheetViews = (text: string | undefined): NamedSheetView[] => {
  if (!text) return [];
  const root = parseXml(text).children[0];
  if (!root || root.name !== 'namedSheetViews') return [];
  return kids(root, 'namedSheetView').slice(0, 1000).map((n) => {
    const v: NamedSheetView = { name: (attr(n, 'name') ?? 'View').slice(0, 255), hiddenRows: [], hiddenCols: [] };
    const id = attr(n, 'id');
    if (id && guidRe.test(id)) v.id = id;
    const f = kid(n, 'nsvFilter');
    const range = f && parseRange(attr(f, 'ref') ?? '');
    if (f && range) {
      const af: AutoFilter = { range, columns: kids(f, 'columnFilter').map((cf) => readFilterColumn(kid(cf, 'filter') ?? cf, numAttr(cf, 'colId') ?? 0)) };
      const sort = kids(kid(f, 'sortRules'), 'sortRule').map((sr) => {
        const sc = kid(sr, 'sortCondition') ?? kid(sr, 'richSortCondition');
        return { col: numAttr(sr, 'colId') ?? 0, ...(boolAttr(sc, 'descending') ? { descending: true } : {}) };
      });
      if (sort.length) af.sort = sort;
      v.filter = af;
    }
    return v;
  });
};

// ---------- Custom views ----------

export const customWorkbookViewsXml = (wb: Workbook): string => {
  const views = wb.customViews?.filter((v) => v.name.trim());
  if (!views?.length) return '';
  return `<customWorkbookViews>${views.map((v) => el('customWorkbookView', {
    name: v.name, guid: guidRe.test(v.guid) ? v.guid : (v.guid = newGuid()), includePrintSettings: v.includePrint ? undefined : false, includeHiddenRowCol: v.includeHidden ? undefined : false,
    maximized: true, windowWidth: 1400, windowHeight: 900, activeSheetId: Math.max(0, Math.min(v.activeSheet, wb.sheets.length - 1)) + 1,
  })).join('')}</customWorkbookViews>`;
};

const pageSetupXml = (ps: NonNullable<Sheet['pageSetup']>): string => {
  let x = '';
  if (ps.margins) x += el('pageMargins', { ...ps.margins });
  const o: Record<string, string | number | boolean | undefined> = { paperSize: ps.paper, scale: ps.scale, fitToWidth: ps.fitWidth, fitToHeight: ps.fitHeight, orientation: ps.orientation, firstPageNumber: ps.firstPageNumber, useFirstPageNumber: ps.firstPageNumber !== undefined || undefined, pageOrder: ps.order };
  if (Object.values(o).some((v) => v !== undefined)) x += el('pageSetup', o);
  if (ps.header || ps.footer) x += el('headerFooter', {}, (ps.header ? `<oddHeader>${esc(ps.header)}</oddHeader>` : '') + (ps.footer ? `<oddFooter>${esc(ps.footer)}</oddFooter>` : ''));
  return x;
};

/** customSheetViews element for one sheet (written between autoFilter and mergeCells). */
export const customSheetViewsXml = (wb: Workbook, si: number): string => {
  const items = (wb.customViews ?? []).flatMap((v) => {
    const sv = v.sheets.find((x) => x.sheet === si);
    if (!sv) return [];
    const pane = sv.frozenRows || sv.frozenCols
      ? el('pane', { xSplit: sv.frozenCols || undefined, ySplit: sv.frozenRows || undefined, topLeftCell: addr(sv.frozenRows ?? 0, sv.frozenCols ?? 0), activePane: sv.frozenRows && sv.frozenCols ? 'bottomRight' : sv.frozenRows ? 'bottomLeft' : 'topRight', state: 'frozen' })
      : '';
    const sel = sv.selection ? el('selection', { activeCell: addr(sv.selection.r1, sv.selection.c1), sqref: rangeAddr(sv.selection) }) : '';
    const ps = v.includePrint && sv.pageSetup ? pageSetupXml(sv.pageSetup) : '';
    const af = sv.filter ? autoFilterXml(sv.filter) : '';
    return [el('customSheetView', {
      guid: v.guid, scale: sv.zoom && sv.zoom !== 100 ? sv.zoom : undefined, showGridLines: sv.showGrid === false ? false : undefined, showRowCol: sv.showHeaders === false ? false : undefined,
      filter: sv.filter ? true : undefined, showAutoFilter: sv.filter ? true : undefined, hiddenRows: sv.hiddenRows?.length ? true : undefined, hiddenColumns: sv.hiddenCols?.length ? true : undefined,
      state: sv.state && sv.state !== 'visible' ? sv.state : undefined, printArea: v.includePrint && sv.pageSetup?.printArea ? true : undefined,
    }, pane + sel + ps + af)];
  });
  return items.length ? `<customSheetViews>${items.join('')}</customSheetViews>` : '';
};

export const readCustomViews = (wbx: XNode, sheetXml: (i: number) => XNode | undefined, sheetCount: number): CustomView[] => {
  const out: CustomView[] = [];
  for (const n of kids(kid(wbx, 'customWorkbookViews'), 'customWorkbookView').slice(0, 500)) {
    const guid = attr(n, 'guid') ?? '';
    if (!guidRe.test(guid)) continue;
    out.push({
      name: (attr(n, 'name') ?? 'View').slice(0, 255), guid, includePrint: attr(n, 'includePrintSettings') !== '0' && attr(n, 'includePrintSettings') !== 'false',
      includeHidden: attr(n, 'includeHiddenRowCol') !== '0' && attr(n, 'includeHiddenRowCol') !== 'false',
      activeSheet: Math.max(0, Math.min(sheetCount - 1, (numAttr(n, 'activeSheetId') ?? 1) - 1)), sheets: [],
    });
  }
  if (!out.length) return out;
  const byGuid = new Map(out.map((v) => [v.guid.toUpperCase(), v]));
  for (let i = 0; i < sheetCount; i++) {
    for (const csv of kids(kid(sheetXml(i), 'customSheetViews'), 'customSheetView')) {
      const v = byGuid.get((attr(csv, 'guid') ?? '').toUpperCase());
      if (!v) continue;
      const sv: CustomView['sheets'][number] = { sheet: i };
      const scale = numAttr(csv, 'scale');
      if (scale) sv.zoom = scale;
      if (attr(csv, 'showGridLines') === '0') sv.showGrid = false;
      if (attr(csv, 'showRowCol') === '0') sv.showHeaders = false;
      const st = attr(csv, 'state');
      if (st === 'hidden' || st === 'veryHidden') sv.state = st;
      const pane = kid(csv, 'pane');
      if (pane && /frozen/.test(attr(pane, 'state') ?? '')) { sv.frozenCols = Math.round(numAttr(pane, 'xSplit') ?? 0); sv.frozenRows = Math.round(numAttr(pane, 'ySplit') ?? 0); }
      const sel = kid(csv, 'selection');
      const r = sel && parseRange((attr(sel, 'sqref') ?? attr(sel, 'activeCell') ?? '').split(' ')[0] ?? '');
      if (r) sv.selection = r;
      const af = readAutoFilterNode(kid(csv, 'autoFilter'));
      if (af) sv.filter = af;
      const ps = kid(csv, 'pageSetup'), pm = kid(csv, 'pageMargins');
      if (ps || pm) {
        const setup: NonNullable<Sheet['pageSetup']> = {};
        const o = attr(ps, 'orientation');
        if (o === 'portrait' || o === 'landscape') setup.orientation = o;
        const paper = numAttr(ps, 'paperSize'); if (paper) setup.paper = paper;
        const scale2 = numAttr(ps, 'scale'); if (scale2) setup.scale = scale2;
        if (pm) setup.margins = { left: numAttr(pm, 'left') ?? 0.7, right: numAttr(pm, 'right') ?? 0.7, top: numAttr(pm, 'top') ?? 0.75, bottom: numAttr(pm, 'bottom') ?? 0.75, header: numAttr(pm, 'header') ?? 0.3, footer: numAttr(pm, 'footer') ?? 0.3 };
        sv.pageSetup = setup;
      }
      v.sheets.push(sv);
    }
  }
  return out;
};

// ---------- Split panes ----------

const HEADER_W = 25, HEADER_H = 20;
/** pane element for a split (not frozen) window; sizes are written in twips including the headers. */
export const splitPaneXml = (s: Sheet): { pane: string; activePane: string; topLeft?: string } | null => {
  const sp = s.view.split;
  if (!sp || s.view.frozenRows || s.view.frozenCols || (sp.x <= 0 && sp.y <= 0)) return null;
  const activePane = sp.x > 0 && sp.y > 0 ? 'bottomRight' : sp.y > 0 ? 'bottomLeft' : 'topRight';
  return {
    pane: el('pane', { xSplit: sp.x > 0 ? Math.round((sp.x + HEADER_W) * 15) : undefined, ySplit: sp.y > 0 ? Math.round((sp.y + HEADER_H) * 15) : undefined, activePane, state: undefined }),
    activePane,
  };
};

export const readSplitPane = (pane: XNode | undefined, sheet: Sheet): void => {
  if (!pane) return;
  const state = attr(pane, 'state') ?? 'split';
  if (state !== 'split') return;
  const xs = numAttr(pane, 'xSplit') ?? 0, ys = numAttr(pane, 'ySplit') ?? 0;
  const x = xs > 0 ? Math.max(20, xs / 15 - HEADER_W) : 0, y = ys > 0 ? Math.max(10, ys / 15 - HEADER_H) : 0;
  if (x > 0 || y > 0) sheet.view.split = { x, y, scrollX: 0, scrollY: 0 };
};

// ---------- Ignored errors ----------

const ERROR_TYPES: ErrorCheckType[] = ['numberStoredAsText', 'formula', 'formulaRange', 'unlockedFormula', 'twoDigitTextYear', 'evalError', 'emptyCellReference', 'calculatedColumn', 'listDataValidation'];

export const ignoredErrorsXml = (s: Sheet): string => {
  if (!s.ignoredErrors?.length) return '';
  // Group ranges that ignore the same set of rules.
  const groups = new Map<string, string[]>();
  for (const ie of s.ignoredErrors) {
    const key = [...new Set(ie.types)].sort().join(',');
    if (!key) continue;
    (groups.get(key) ?? groups.set(key, []).get(key)!).push(rangeAddr(ie.range));
  }
  if (!groups.size) return '';
  return `<ignoredErrors>${[...groups].map(([k, refs]) => el('ignoredError', { sqref: refs.join(' '), ...Object.fromEntries(k.split(',').map((t) => [t, true])) })).join('')}</ignoredErrors>`;
};

export const readIgnoredErrors = (ws: XNode, sheet: Sheet): void => {
  const out: NonNullable<Sheet['ignoredErrors']> = [];
  for (const n of kids(kid(ws, 'ignoredErrors'), 'ignoredError')) {
    const types = ERROR_TYPES.filter((t) => boolAttr(n, t));
    if (!types.length) continue;
    for (const ref of (attr(n, 'sqref') ?? '').split(/\s+/).slice(0, 10000)) {
      const r = parseRange(ref);
      if (r) out.push({ range: r, types });
    }
  }
  if (out.length) sheet.ignoredErrors = out;
};

// ---------- ClerdoGrid workspace part ----------

interface WorkspaceDoc {
  app: 'ClerdoGrid';
  version: 1;
  /** Named sheet views by sheet index, with hidden rows and columns (which the shared part can't hold). */
  views?: Record<string, { name: string; id?: string; filter?: AutoFilter; hiddenRows: number[]; hiddenCols: number[] }[]>;
  /** Exact hidden rows and columns of custom views, by guid and sheet. */
  customHidden?: Record<string, { sheet: number; rows: number[]; cols: number[] }[]>;
  /** Top-left pane scroll offsets of split windows. */
  splits?: Record<string, { scrollX: number; scrollY: number; x: number; y: number }>;
}

export const workspacePart = (wb: Workbook): string | null => {
  const doc: WorkspaceDoc = { app: 'ClerdoGrid', version: 1 };
  wb.sheets.forEach((s, i) => {
    if (s.sheetViews?.length) (doc.views ??= {})[String(i)] = s.sheetViews.map((v) => ({ name: v.name, id: v.id, filter: v.filter, hiddenRows: v.hiddenRows, hiddenCols: v.hiddenCols }));
    if (s.view.split && !s.view.frozenRows && !s.view.frozenCols) (doc.splits ??= {})[String(i)] = { ...s.view.split };
  });
  for (const v of wb.customViews ?? []) {
    const hidden = v.sheets.filter((x) => x.hiddenRows?.length || x.hiddenCols?.length).map((x) => ({ sheet: x.sheet, rows: x.hiddenRows ?? [], cols: x.hiddenCols ?? [] }));
    if (hidden.length) (doc.customHidden ??= {})[v.guid] = hidden;
  }
  return doc.views || doc.customHidden || doc.splits ? JSON.stringify(doc) : null;
};

const nums = (v: unknown, max = 1_048_576): number[] => (Array.isArray(v) ? v.filter((x): x is number => Number.isInteger(x) && x >= 0 && x < max).slice(0, 1_048_576) : []);
const cleanFilter = (f: unknown): AutoFilter | undefined => {
  if (!f || typeof f !== 'object') return undefined;
  const o = f as Partial<AutoFilter>;
  const r = o.range;
  if (!r || ![r.r1, r.c1, r.r2, r.c2].every((x) => Number.isInteger(x) && x >= 0)) return undefined;
  const columns = Array.isArray(o.columns) ? o.columns.filter((c) => c && Number.isInteger(c.col)).slice(0, 16384).map((c) => {
    const out: FilterColumn = { col: c.col };
    if (Array.isArray(c.values)) out.values = c.values.filter((x) => typeof x === 'string').slice(0, 100000);
    if (c.blanks) out.blanks = true;
    if (c.custom && Array.isArray(c.custom.rules)) out.custom = { and: !!c.custom.and, rules: c.custom.rules.slice(0, 2).map((x) => ({ operator: x.operator, value: String(x.value ?? '') })) };
    if (c.top10) out.top10 = { top: !!c.top10.top, percent: !!c.top10.percent, value: Number(c.top10.value) || 10 };
    if (typeof c.color === 'string' && /^#[0-9a-f]{6,8}$/i.test(c.color)) out.color = c.color;
    return out;
  }) : [];
  const af: AutoFilter = { range: { r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 }, columns };
  if (Array.isArray(o.sort)) af.sort = o.sort.filter((x) => x && Number.isInteger(x.col)).slice(0, 64).map((x) => ({ col: x.col, ...(x.descending ? { descending: true } : {}) }));
  return af;
};

export const applyWorkspacePart = (wb: Workbook, text: string | undefined): void => {
  if (!text || text.length > 64 * 1024 * 1024) return;
  let doc: Partial<WorkspaceDoc>;
  try { doc = JSON.parse(text) as Partial<WorkspaceDoc>; } catch { return; }
  if (!doc || typeof doc !== 'object') return;
  for (const [idx, list] of Object.entries(doc.views ?? {})) {
    const s = wb.sheets[Number(idx)];
    if (!s || !Array.isArray(list)) continue;
    s.sheetViews = list.slice(0, 1000).filter((v) => v && typeof v.name === 'string').map((v) => {
      const out: NamedSheetView = { name: v.name.slice(0, 255), hiddenRows: nums(v.hiddenRows), hiddenCols: nums(v.hiddenCols, 16384) };
      if (typeof v.id === 'string' && guidRe.test(v.id)) out.id = v.id;
      const f = cleanFilter(v.filter);
      if (f) out.filter = f;
      return out;
    });
  }
  for (const [guid, list] of Object.entries(doc.customHidden ?? {})) {
    const v = wb.customViews?.find((x) => x.guid.toUpperCase() === guid.toUpperCase());
    if (!v || !Array.isArray(list)) continue;
    for (const h of list) {
      if (!h || !Number.isInteger(h.sheet)) continue;
      let sv = v.sheets.find((x) => x.sheet === h.sheet);
      if (!sv && wb.sheets[h.sheet]) v.sheets.push((sv = { sheet: h.sheet }));
      if (!sv) continue;
      sv.hiddenRows = nums(h.rows); sv.hiddenCols = nums(h.cols, 16384);
    }
  }
  for (const [idx, sp] of Object.entries(doc.splits ?? {})) {
    const s = wb.sheets[Number(idx)];
    if (!s || !sp || typeof sp !== 'object') continue;
    const n = (x: unknown) => (typeof x === 'number' && Number.isFinite(x) && x >= 0 ? Math.min(x, 1e7) : 0);
    if (n(sp.x) > 0 || n(sp.y) > 0) s.view.split = { x: n(sp.x), y: n(sp.y), scrollX: n(sp.scrollX), scrollY: n(sp.scrollY) };
  }
};

export { parseAddr };
