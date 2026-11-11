// PivotTable parts of an .xlsx package (ECMA-376 Part 1 §18.10): pivot cache definitions with shared
// items and field groups, cache records, PivotTable definitions, and the workbook's pivotCaches list.
// Extensions follow the published MS-XLSX extension list (x14 data field "show as" and pivot fields).
import { fromSerial, timeOf, toSerial } from '../engine/date';
import { CellError, ERROR_CODES, isError, type ErrorCode, type Scalar } from '../engine/types';
import { internStyle, parseRange, rangeAddr, type CellRange, type CellStyle, type Workbook } from '../model/workbook';
import {
  VALUES_FIELD,
  type AxisItem, type CacheField, type DataField, type DateGroupBy, type FieldGroup, type ItemType, type LabelOp, type PivotAgg,
  type PivotCache, type PivotField, type PivotFilter, type PivotItem, type PivotTable, type ShowAs, type Subtotal, type ValueOp,
} from '../model/pivot';
import { builtinFormat } from './numfmt';
import { attr, boolAttr, el, esc, kid, kids, numAttr, type XNode } from './xml';

export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const EXT_DATAFIELD = '{E15A36E0-9728-4e99-A89B-3F7291B0FE68}';
const EXT_PIVOTFIELD = '{2946ED86-A175-432a-8AC1-64E0C546D7DE}';
/** ClerdoGrid's own pivot settings that have no place in the standard schema. */
const EXT_CLERDO = '{8E5D2C71-3F4A-4B9E-A6D2-7C1B5E9F0A34}';
const NS_CLERDO = 'https://clerdo.app/ns/grid/pivot/1';

export const PIVOT_RT = { cacheDef: RT + 'pivotCacheDefinition', records: RT + 'pivotCacheRecords', table: RT + 'pivotTable' };
export const PIVOT_CT = { cacheDef: CT + 'pivotCacheDefinition+xml', records: CT + 'pivotCacheRecords+xml', table: CT + 'pivotTable+xml' };

// ---------------------------------------------------------------- values

const isoOf = (serial: number, date1904: boolean): string => {
  const d = fromSerial(Math.floor(serial), date1904);
  const t = timeOf(serial - Math.floor(serial));
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  return `${p(d.year, 4)}-${p(d.month)}-${p(d.day)}T${p(t.hour)}:${p(t.minute)}:${p(t.second)}`;
};
const serialOf = (iso: string, date1904: boolean): number | null => {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?/.exec(iso);
  if (!m) return null;
  const day = toSerial(+m[1]!, +m[2]!, +m[3]!, date1904);
  return day + ((+(m[4] ?? 0)) * 3600 + (+(m[5] ?? 0)) * 60 + (+(m[6] ?? 0))) / 86400;
};

const itemXml = (v: Scalar, dates: boolean, date1904: boolean): string => {
  if (v === null || v === '') return '<m/>';
  if (typeof v === 'number') return dates ? `<d v="${isoOf(v, date1904)}"/>` : `<n v="${Number.isFinite(v) ? v : 0}"/>`;
  if (typeof v === 'boolean') return `<b v="${v ? 1 : 0}"/>`;
  if (isError(v)) return `<e v="${esc(v.code)}"/>`;
  return `<s v="${esc(v)}"/>`;
};

const readItem = (n: XNode, date1904: boolean): Scalar => {
  const v = attr(n, 'v');
  switch (n.name) {
    case 'n': return v !== undefined && Number.isFinite(+v) ? +v : 0;
    case 'b': return v === '1' || v === 'true';
    case 'e': return new CellError((ERROR_CODES.includes(v as ErrorCode) ? v : '#N/A') as ErrorCode);
    case 'd': return v ? serialOf(v, date1904) : null;
    case 's': return v ?? '';
    default: return null;
  }
};

const sharedItemsXml = (items: Scalar[], dates: boolean, date1904: boolean): string => {
  const strs = items.some((v) => typeof v === 'string' && v !== '');
  const numbers = items.filter((v): v is number => typeof v === 'number');
  const blank = items.some((v) => v === null || v === '');
  const other = items.some((v) => typeof v === 'boolean' || isError(v));
  const types = [strs, numbers.length > 0, other].filter(Boolean).length;
  const a: Record<string, string | number | boolean | undefined> = {};
  if (!strs) { a.containsSemiMixedTypes = false; a.containsString = false; }
  if (dates && numbers.length) {
    a.containsNonDate = false; a.containsDate = true;
    a.minDate = isoOf(Math.min(...numbers), date1904); a.maxDate = isoOf(Math.max(...numbers) + (numbers.every(Number.isInteger) ? 1 : 0), date1904);
  } else if (numbers.length) {
    a.containsNumber = true;
    if (numbers.every(Number.isInteger)) a.containsInteger = true;
    a.minValue = Math.min(...numbers); a.maxValue = Math.max(...numbers);
  }
  if (blank) a.containsBlank = true;
  if (types > 1) a.containsMixedTypes = true;
  if (items.some((v) => typeof v === 'string' && v.length > 255)) a.longText = true;
  a.count = items.length;
  return el('sharedItems', a, items.map((v) => itemXml(v, dates, date1904)).join(''));
};

// ---------------------------------------------------------------- writing caches

const GROUP_BY: Record<string, string> = { range: 'range', years: 'years', quarters: 'quarters', months: 'months', days: 'days', hours: 'hours', minutes: 'minutes', seconds: 'seconds' };

const fieldGroupXml = (g: FieldGroup, date1904: boolean): string => {
  let inner = '';
  if (g.range) {
    const r = g.range;
    inner += el('rangePr', {
      autoStart: r.autoStart ? undefined : false, autoEnd: r.autoEnd ? undefined : false, groupBy: r.by === 'range' ? undefined : GROUP_BY[r.by],
      startNum: r.dates ? undefined : r.start, endNum: r.dates ? undefined : r.end,
      startDate: r.dates ? isoOf(r.start, date1904) : undefined, endDate: r.dates ? isoOf(r.end + (r.by === 'range' ? 0 : 1), date1904) : undefined,
      groupInterval: r.interval !== 1 ? r.interval : undefined,
    });
  }
  if (g.discrete) inner += el('discretePr', { count: g.discrete.length }, g.discrete.map((x) => `<x v="${x}"/>`).join(''));
  inner += el('groupItems', { count: g.items.length }, g.items.map((s) => `<s v="${esc(s)}"/>`).join(''));
  return el('fieldGroup', { par: g.parent, base: g.base }, inner);
};

/** Calculated items live in the cache: their names are extra shared items of the field. */
interface CalcItemOut { field: number; index: number; formula: string }

const cacheDefXml = (wb: Workbook, cache: PivotCache, numFmtId: (code: string | undefined) => number, calc: CalcItemOut[], extraItems: Map<number, string[]>): string => {
  const src = cache.source;
  const ws = src.kind === 'range' ? el('worksheetSource', { ref: rangeAddr(src.range), sheet: src.sheet }) : el('worksheetSource', { name: src.name });
  const fields = cache.fields.map((f, i) => {
    const a: Record<string, string | number | boolean | undefined> = { name: f.name, numFmtId: numFmtId(f.numFmt ?? (f.dates ? 'm/d/yyyy' : undefined)) };
    if (f.formula) { a.formula = f.formula; a.databaseField = false; }
    else if (f.database === false) a.databaseField = false;
    const extra = extraItems.get(i) ?? [];
    const shared = f.database === false && !f.formula ? '' : f.formula ? '' : sharedItemsXml([...f.items, ...extra], !!f.dates, wb.date1904);
    const group = f.group ? fieldGroupXml(f.group, wb.date1904) : '';
    return el('cacheField', a, (f.database === false ? (extra.length ? sharedItemsXml(extra, false, wb.date1904) : '') : shared) + group);
  }).join('');
  const calcXml = calc.length
    ? el('calculatedItems', { count: calc.length }, calc.map((c) => el('calculatedItem', { formula: c.formula },
      `<pivotArea cacheIndex="1" outline="0" fieldPosition="0"><references count="1"><reference field="${c.field}" count="1"><x v="${c.index}"/></reference></references></pivotArea>`)).join(''))
    : '';
  return `${XML_HEAD}<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" r:id="rId1"`
    + `${cache.refreshOnLoad ? ' refreshOnLoad="1"' : ''} refreshedBy="${esc(cache.refreshedBy ?? 'ClerdoGrid')}" refreshedDate="${refreshedSerial(cache)}"`
    + ` createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${cache.records.length}"${cache.dropMissing ? ' missingItemsLimit="0"' : ''}>`
    + `<cacheSource type="worksheet">${ws}</cacheSource>`
    + el('cacheFields', { count: cache.fields.length }, fields)
    + calcXml
    + `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cache.id}"/></ext></extLst>`
    + '</pivotCacheDefinition>';
};

const refreshedSerial = (c: PivotCache): number => {
  const ms = c.refreshedDate ?? Date.now();
  return Math.round((ms / 86400000 + 25569) * 1e6) / 1e6;
};

const recordsXml = (wb: Workbook, cache: PivotCache): string => {
  const db = cache.fields.filter((f) => f.database !== false);
  const maps = db.map((f) => { const m = new Map<string, number>(); f.items.forEach((v, i) => { const k = key(v); if (!m.has(k)) m.set(k, i); }); return m; });
  const parts: string[] = [];
  for (const rec of cache.records) {
    let r = '<r>';
    for (let j = 0; j < db.length; j++) {
      const v = rec[j] ?? null;
      const i = maps[j]!.get(key(v));
      r += i !== undefined ? `<x v="${i}"/>` : itemXml(v, !!db[j]!.dates, wb.date1904);
    }
    parts.push(r + '</r>');
  }
  return `${XML_HEAD}<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_R}" count="${cache.records.length}">${parts.join('')}</pivotCacheRecords>`;
};

const key = (v: Scalar): string => (v === null || v === '' ? 'b' : typeof v === 'number' ? `n${v}` : typeof v === 'boolean' ? `l${v}` : isError(v) ? `e${v.code}` : `s${v.toLowerCase()}`);

// ---------------------------------------------------------------- writing tables

const SUB_ATTR: Record<Exclude<Subtotal, 'default'>, string> = {
  sum: 'sumSubtotal', count: 'countASubtotal', average: 'avgSubtotal', max: 'maxSubtotal', min: 'minSubtotal', product: 'productSubtotal',
  countNums: 'countSubtotal', stdDev: 'stdDevSubtotal', stdDevp: 'stdDevPSubtotal', var: 'varSubtotal', varp: 'varPSubtotal',
};
const SUB_ITEM: Record<Subtotal, ItemType> = {
  default: 'default', sum: 'sum', count: 'countA', average: 'avg', max: 'max', min: 'min', product: 'product', countNums: 'count',
  stdDev: 'stdDev', stdDevp: 'stdDevP', var: 'var', varp: 'varP',
};
const AGG_FILE: Record<PivotAgg, string> = {
  sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums',
  stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp', distinctCount: 'count',
};
const X14_SHOW: Partial<Record<ShowAs, string>> = {
  percentOfParentRow: 'percentOfParentRow', percentOfParentCol: 'percentOfParentCol', percentOfParent: 'percentOfParent',
  percentOfRunningTotal: 'percentOfRunningTotal', rankAscending: 'rankAscending', rankDescending: 'rankDescending',
};
const LABEL_TYPE: Record<LabelOp, string> = {
  equal: 'captionEqual', notEqual: 'captionNotEqual', beginsWith: 'captionBeginsWith', notBeginsWith: 'captionNotBeginsWith', endsWith: 'captionEndsWith',
  notEndsWith: 'captionNotEndsWith', contains: 'captionContains', notContains: 'captionNotContains', greaterThan: 'captionGreaterThan',
  greaterThanOrEqual: 'captionGreaterThanOrEqual', lessThan: 'captionLessThan', lessThanOrEqual: 'captionLessThanOrEqual', between: 'captionBetween', notBetween: 'captionNotBetween',
};
const VALUE_TYPE: Record<ValueOp, string> = {
  equal: 'valueEqual', notEqual: 'valueNotEqual', greaterThan: 'valueGreaterThan', greaterThanOrEqual: 'valueGreaterThanOrEqual',
  lessThan: 'valueLessThan', lessThanOrEqual: 'valueLessThanOrEqual', between: 'valueBetween', notBetween: 'valueNotBetween',
};
const CUSTOM_OP: Record<string, string | undefined> = {
  equal: undefined, notEqual: 'notEqual', greaterThan: 'greaterThan', greaterThanOrEqual: 'greaterThanOrEqual', lessThan: 'lessThan', lessThanOrEqual: 'lessThanOrEqual',
};

const filterXml = (f: PivotFilter, fld: number, id: number): string => {
  let type: string, inner: string;
  const a: Record<string, string | number | boolean | undefined> = { fld, evalOrder: -1, id };
  if (f.kind === 'label') {
    type = LABEL_TYPE[f.op];
    a.stringValue1 = f.value1; a.stringValue2 = f.value2;
    const pat = f.op.includes('egins') ? `${f.value1}*` : f.op.includes('nds') ? `*${f.value1}` : f.op.includes('ontains') ? `*${f.value1}*` : f.value1;
    const neg = f.op.startsWith('not') && f.op !== 'notBetween';
    const custom = f.op === 'between' ? `<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${esc(f.value1)}"/><customFilter operator="lessThanOrEqual" val="${esc(f.value2 ?? '')}"/></customFilters>`
      : f.op === 'notBetween' ? `<customFilters><customFilter operator="lessThan" val="${esc(f.value1)}"/><customFilter operator="greaterThan" val="${esc(f.value2 ?? '')}"/></customFilters>`
      : `<customFilters>${el('customFilter', { operator: neg ? 'notEqual' : CUSTOM_OP[f.op], val: pat })}</customFilters>`;
    inner = custom;
  } else if (f.kind === 'value') {
    type = VALUE_TYPE[f.op];
    a.iMeasureFld = f.dataField;
    const custom = f.op === 'between' ? `<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${f.value1}"/><customFilter operator="lessThanOrEqual" val="${f.value2 ?? f.value1}"/></customFilters>`
      : f.op === 'notBetween' ? `<customFilters><customFilter operator="lessThan" val="${f.value1}"/><customFilter operator="greaterThan" val="${f.value2 ?? f.value1}"/></customFilters>`
      : `<customFilters>${el('customFilter', { operator: CUSTOM_OP[f.op], val: f.value1 })}</customFilters>`;
    inner = custom;
  } else if (f.kind === 'dates') {
    type = 'dateBetween';
    a.name = f.timeline;
    inner = `<customFilters and="1"><customFilter operator="greaterThanOrEqual" val="${f.start}"/><customFilter operator="lessThanOrEqual" val="${f.end}"/></customFilters>`;
  } else {
    type = f.by;
    a.iMeasureFld = f.dataField;
    inner = el('top10', { top: f.top ? undefined : false, percent: f.by === 'percent' || undefined, val: f.value, filterVal: f.value });
  }
  return el('filter', { ...a, type }, `<autoFilter ref="A1"><filterColumn colId="0">${inner}</filterColumn></autoFilter>`);
};

const axisItemsXml = (tag: string, list: AxisItem[] | undefined): string => {
  const items = list?.length ? list : [{ x: [] }];
  return el(tag, { count: items.length }, items.map((it) => el('i', { t: it.t && it.t !== 'data' ? it.t : undefined, r: it.r || undefined, i: it.i || undefined },
    it.x.map((v) => (v ? `<x v="${v}"/>` : '<x/>')).join(''))).join(''));
};

const tableXml = (pt: PivotTable, cache: PivotCache, sink: PartSink, calcIndex: (field: number, slot: number) => number): string => {
  const numFmtId = sink.numFmtId;
  const loc = pt.location ?? { ref: { r1: pt.anchor.row, c1: pt.anchor.col, r2: pt.anchor.row + 1, c2: pt.anchor.col }, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 0 };
  const compact = pt.layout === 'compact', tabular = pt.layout === 'tabular';
  const a: Record<string, string | number | boolean | undefined> = {
    name: pt.name, cacheId: pt.cacheId,
    applyNumberFormats: false, applyBorderFormats: false, applyFontFormats: false, applyPatternFormats: false, applyAlignmentFormats: false, applyWidthHeightFormats: true,
    dataCaption: pt.dataCaption ?? 'Values', grandTotalCaption: pt.grandTotalCaption, errorCaption: pt.errorText, showError: pt.showError || undefined,
    missingCaption: pt.emptyText, showMissing: pt.showEmpty === false ? false : undefined,
    updatedVersion: 6, minRefreshableVersion: 3, showDrill: pt.showDrill ? undefined : false, useAutoFormatting: pt.autoFit || undefined,
    itemPrintTitles: true, createdVersion: 6, indent: pt.indent ?? 0, preserveFormatting: pt.preserveFormatting ? undefined : false,
    showHeaders: pt.showHeaders ? undefined : false, rowGrandTotals: pt.rowGrand ? undefined : false, colGrandTotals: pt.colGrand ? undefined : false,
    compact: compact ? undefined : false, compactData: compact ? undefined : false, outline: tabular ? undefined : true, outlineData: tabular ? undefined : true,
    gridDropZones: undefined, multipleFieldFilters: false,
    rowHeaderCaption: pt.rowHeaderCaption, colHeaderCaption: pt.colHeaderCaption, pageWrap: pt.pageWrap || undefined, pageOverThenDown: pt.pageOverThenDown || undefined,
  };
  const pages = pt.pages;
  const pageCols = pages.length ? (pt.pageOverThenDown ? Math.min(pages.length, pt.pageWrap || pages.length) : Math.ceil(pages.length / (pt.pageWrap || pages.length))) : 0;
  const pageRowsN = pages.length ? (pt.pageOverThenDown ? Math.ceil(pages.length / (pt.pageWrap || pages.length)) : Math.min(pages.length, pt.pageWrap || pages.length)) : 0;
  let x = `${XML_HEAD}<pivotTableDefinition xmlns="${NS_MAIN}"${attrsOf(a)}>`;
  x += el('location', { ref: rangeAddr(loc.ref), firstHeaderRow: loc.firstHeaderRow, firstDataRow: loc.firstDataRow, firstDataCol: loc.firstDataCol, rowPageCount: pageRowsN || undefined, colPageCount: pageCols || undefined });
  const axisOf = (f: number) => (pt.rows.includes(f) ? 'axisRow' : pt.cols.includes(f) ? 'axisCol' : pages.some((p) => p.field === f) ? 'axisPage' : undefined);
  const fields = cache.fields.map((cf, f) => {
    const pf: PivotField = pt.fields[f] ?? { items: [] };
    const axis = axisOf(f);
    const isData = pt.data.some((d) => d.field === f);
    const subs = pf.subtotals ?? ['default'];
    const fa: Record<string, string | number | boolean | undefined> = {
      name: pf.name, axis, dataField: isData || undefined, showAll: pf.showAll ? undefined : false,
      compact: (pf.compact ?? compact) ? undefined : false, outline: (pf.outline ?? !tabular) ? undefined : false,
      subtotalTop: pf.subtotalTop === false ? false : undefined, insertBlankRow: pf.blankRow || undefined, insertPageBreak: pf.pageBreak || undefined,
      multipleItemSelectionAllowed: pf.multiPage || undefined, numFmtId: pf.numFmt ? numFmtId(pf.numFmt) : undefined,
      sortType: pf.sort, defaultSubtotal: subs.includes('default') ? undefined : false, includeNewItemsInFilter: pf.includeNew === false ? undefined : undefined,
    };
    for (const s of subs) if (s !== 'default') fa[SUB_ATTR[s]] = true;
    let inner = '';
    if ((axis || pf.items.some((it) => it.hidden)) && !cf.formula) {
      const groupCount = cf.group?.range ? cf.group.items.length : Infinity;
      const its = pf.items.filter((it) => it.x < 0 ? !!it.formula : it.x < groupCount).map((it, s) => el('item', {
        n: it.name, h: it.hidden || undefined, sd: it.collapsed ? false : undefined, f: it.x < 0 || undefined,
        x: it.x < 0 ? calcIndex(f, s) : it.x,
      }));
      for (const s of subs) its.push(el('item', { t: SUB_ITEM[s] }));
      inner += el('items', { count: its.length }, its.join(''));
    }
    if (pf.sortByData !== undefined && pf.sort) {
      inner += `<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="${pf.sortByData}"/></reference></references></pivotArea></autoSortScope>`;
    }
    if (pf.repeatLabels) inner += `<extLst><ext uri="${EXT_PIVOTFIELD}" xmlns:x14="${NS_X14}"><x14:pivotField fillDownLabels="1"/></ext></extLst>`;
    return el('pivotField', fa, inner || undefined);
  }).join('');
  x += el('pivotFields', { count: cache.fields.length }, fields);
  if (pt.rows.length) {
    x += el('rowFields', { count: pt.rows.length }, pt.rows.map((f) => `<field x="${f}"/>`).join(''));
    x += axisItemsXml('rowItems', pt.rowItems);
  } else if (pt.rowItems?.length) x += axisItemsXml('rowItems', pt.rowItems);
  if (pt.cols.length) x += el('colFields', { count: pt.cols.length }, pt.cols.map((f) => `<field x="${f}"/>`).join(''));
  x += axisItemsXml('colItems', pt.colItems);
  if (pages.length) x += el('pageFields', { count: pages.length }, pages.map((p) => el('pageField', { fld: p.field, item: p.item, hier: -1 })).join(''));
  if (pt.data.length) {
    x += el('dataFields', { count: pt.data.length }, pt.data.map((d) => dataFieldXml(d, cache, numFmtId)).join(''));
  }
  if (pt.formats?.length && sink.dxf && sink.styles) {
    const items = pt.formats.flatMap((f) => { const a = areaXml(f.key, pt); return a ? [el('format', { dxfId: sink.dxf!(sink.styles![f.style] ?? {}) }, a)] : []; });
    if (items.length) x += el('formats', { count: items.length }, items.join(''));
  }
  x += el('pivotTableStyleInfo', { name: pt.style.name || undefined, showRowHeaders: pt.style.rowHeaders, showColHeaders: pt.style.colHeaders, showRowStripes: pt.style.bandedRows, showColStripes: pt.style.bandedCols, showLastColumn: pt.style.lastColumn ?? true });
  const filters: string[] = [];
  let fid = 1;
  pt.fields.forEach((pf, f) => { for (const flt of pf.filters ?? []) filters.push(filterXml(flt, f, fid++)); });
  if (filters.length) x += el('filters', { count: filters.length }, filters.join(''));
  const own = clerdoExt(pt);
  if (own) x += `<extLst><ext uri="${EXT_CLERDO}" xmlns:cg="${NS_CLERDO}">${own}</ext></extLst>`;
  return x + '</pivotTableDefinition>';
};

const clerdoExt = (pt: PivotTable): string => {
  const distinct = pt.data.map((d, i) => (d.agg === 'distinctCount' ? i : -1)).filter((i) => i >= 0);
  let out = '';
  if (distinct.length) out += `<cg:distinctCount fields="${distinct.join(' ')}"/>`;
  if (pt.deferLayout) out += '<cg:defer/>';
  for (const f of pt.formats ?? []) out += `<cg:format key="${esc(f.key)}" s="${f.style}"/>`;
  return out;
};

const dataFieldXml = (d: DataField, cache: PivotCache, numFmtId: (code: string | undefined) => number): string => {
  const x14 = d.showAs ? X14_SHOW[d.showAs] : undefined;
  const showDataAs = !d.showAs || d.showAs === 'normal' || x14 ? undefined : d.showAs;
  const usesBase = !!d.showAs && d.showAs !== 'normal' && !['percentOfTotal', 'percentOfRow', 'percentOfCol', 'index', 'percentOfParentRow', 'percentOfParentCol'].includes(d.showAs);
  const baseItem = d.baseItem === 'prev' ? 1048828 : d.baseItem === 'next' ? 1048829 : d.baseItem;
  void cache;
  return el('dataField', {
    name: d.name, fld: d.field, subtotal: d.agg === 'sum' ? undefined : AGG_FILE[d.agg], showDataAs,
    baseField: usesBase ? d.baseField ?? 0 : 0, baseItem: usesBase ? baseItem ?? 0 : 0, numFmtId: d.numFmt ? numFmtId(d.numFmt) : undefined,
  }, x14 ? `<extLst><ext uri="${EXT_DATAFIELD}" xmlns:x14="${NS_X14}"><x14:dataField pivotShowAs="${x14}"/></ext></extLst>` : undefined);
};

const attrsOf = (o: Record<string, string | number | boolean | undefined>): string => {
  let s = '';
  for (const [k, v] of Object.entries(o)) if (v !== undefined) s += ` ${k}="${typeof v === 'boolean' ? (v ? 1 : 0) : esc(String(v))}"`;
  return s;
};

export interface PartSink {
  put: (path: string, text: string, ct?: string) => void;
  numFmtId: (code: string | undefined) => number;
  /** Differential format id for a style (for kept formatting). */
  dxf?: (style: CellStyle) => number;
  styles?: readonly CellStyle[];
}

// ---------------------------------------------------------------- kept formatting as pivot areas
//
// Formatting kept over a table is keyed by the line the cell sits on: "d|row|col" for values,
// "r|line|field" / "c|line|field" for labels and "role|field|data" for captions, where a line is
// "kind:subtotal:data:field.slot,...". These map to §18.3.1.68 pivotArea references and back.

const SIGMA = 4294967294;
interface LineK { kind: string; fn: string; data: number; path: [number, number][] }
const parseLine = (t: string): LineK => {
  const [kind = 'item', fn = '', data = '-1', path = ''] = t.split(':');
  return { kind, fn, data: Number(data), path: path ? path.split(',').map((e) => e.split('.').map(Number) as [number, number]) : [] };
};
const lineText = (l: LineK): string => `${l.kind}:${l.fn}:${l.data}:${l.path.map(([f, x]) => `${f}.${x}`).join(',')}`;
const SUB_FN_ATTR: Record<string, string> = { default: 'defaultSubtotal', sum: 'sumSubtotal', count: 'countASubtotal', average: 'avgSubtotal', max: 'maxSubtotal', min: 'minSubtotal', product: 'productSubtotal', countNums: 'countSubtotal', stdDev: 'stdDevSubtotal', stdDevp: 'stdDevPSubtotal', var: 'varSubtotal', varp: 'varPSubtotal' };

const refsXml = (l: LineK): string[] => l.path.map(([f, x], i) => el('reference', {
  field: f < 0 ? SIGMA : f, count: 1, selected: false, ...(l.kind === 'sub' && i === l.path.length - 1 ? { [SUB_FN_ATTR[l.fn || 'default'] ?? 'defaultSubtotal']: true } : {}),
}, `<x v="${x}"/>`));

const areaXml = (key: string, pt: PivotTable): string | null => {
  const parts = key.split('|');
  const head = parts[0]!;
  const area = (a: Record<string, string | number | boolean | undefined>, refs: string[]) =>
    el('pivotArea', { outline: false, fieldPosition: 0, ...a }, refs.length ? el('references', { count: refs.length }, refs.join('')) : undefined);
  if (head === 'd') {
    const row = parseLine(parts[1] ?? ''), col = parseLine(parts[2] ?? '');
    const refs = [...(row.kind === 'grand' ? [] : refsXml(row)), ...(col.kind === 'grand' ? [] : refsXml(col))];
    const hasSigma = [...row.path, ...col.path].some(([f]) => f === VALUES_FIELD);
    if (!hasSigma && pt.data.length) refs.push(`<reference field="${SIGMA}" count="1" selected="0"><x v="${Math.max(0, row.data, col.data)}"/></reference>`);
    return area({ grandRow: row.kind === 'grand' || undefined, grandCol: col.kind === 'grand' || undefined }, refs);
  }
  if (head === 'r' || head === 'c') {
    const l = parseLine(parts[1] ?? '');
    const field = Number(parts[2]);
    const grand = l.kind === 'grand';
    return area({ dataOnly: false, labelOnly: true, axis: head === 'r' ? 'axisRow' : 'axisCol', field: Number.isFinite(field) ? (field < 0 ? SIGMA : field) : undefined, [head === 'r' ? 'grandRow' : 'grandCol']: grand || undefined }, grand ? [] : refsXml(l));
  }
  const field = parts[1] ? Number(parts[1]) : undefined, data = parts[2] ? Number(parts[2]) : undefined;
  switch (head) {
    case 'dataCaption': return area({ type: 'origin', dataOnly: false, labelOnly: true }, data !== undefined ? [`<reference field="${SIGMA}" count="1" selected="0"><x v="${data}"/></reference>`] : []);
    case 'rowCaption': return area({ type: 'button', dataOnly: false, labelOnly: true, axis: 'axisRow', field }, []);
    case 'colCaption': return area({ type: 'button', dataOnly: false, labelOnly: true, axis: 'axisCol', field }, []);
    case 'pageLabel': return area({ type: 'button', dataOnly: false, labelOnly: true, axis: 'axisPage', field }, []);
    case 'pageValue': return area({ dataOnly: false, labelOnly: true, axis: 'axisPage', field }, []);
  }
  return null;
};

/** The key of a pivotArea read from a file, or null when it isn't a single cell this app keys. */
const areaKey = (a: XNode, pt: PivotTable): string | null => {
  const type = attr(a, 'type') ?? 'normal';
  const axis = attr(a, 'axis');
  const fld = numAttr(a, 'field');
  const field = fld === SIGMA ? VALUES_FIELD : fld;
  const labelOnly = boolAttr(a, 'labelOnly');
  const refs = kids(kid(a, 'references'), 'reference').map((r) => {
    const f = numAttr(r, 'field') ?? 0;
    const sub = Object.entries(SUB_FN_ATTR).find(([, n]) => boolAttr(r, n))?.[0];
    return { f: f === SIGMA || f === -2 ? VALUES_FIELD : f, x: numAttr(kid(r, 'x'), 'v') ?? 0, sub };
  });
  const sigmaRef = refs.find((r) => r.f === VALUES_FIELD);
  const line = (fieldsOnAxis: number[], grand: boolean): LineK => {
    if (grand) return { kind: 'grand', fn: '', data: -1, path: [] };
    const sigmaHere = fieldsOnAxis.includes(VALUES_FIELD);
    const path: [number, number][] = [];
    let sub: string | undefined;
    for (const f of fieldsOnAxis) {
      const r = refs.find((q) => q.f === f);
      if (!r) break;
      path.push([f, r.x]);
      if (r.sub) sub = r.sub;
    }
    return { kind: sub ? 'sub' : 'item', fn: sub ?? '', data: sigmaHere && sigmaRef ? sigmaRef.x : -1, path };
  };
  if (type === 'origin') return `dataCaption||${sigmaRef ? sigmaRef.x : pt.data.length ? 0 : ''}`;
  if (type === 'button') return `${axis === 'axisCol' ? 'colCaption' : axis === 'axisPage' ? 'pageLabel' : 'rowCaption'}|${field ?? ''}|`;
  if (axis === 'axisPage') return `pageValue|${field ?? ''}|`;
  if (type !== 'normal') return null;
  if (labelOnly && (axis === 'axisRow' || axis === 'axisCol')) {
    const isRow = axis === 'axisRow';
    const l = line(isRow ? pt.rows : pt.cols, boolAttr(a, isRow ? 'grandRow' : 'grandCol'));
    return `${isRow ? 'r' : 'c'}|${lineText(l)}|${field ?? l.path.at(-1)?.[0] ?? ''}`;
  }
  if (!boolAttr(a, 'dataOnly', true)) return null;
  return `d|${lineText(line(pt.rows, boolAttr(a, 'grandRow')))}|${lineText(line(pt.cols, boolAttr(a, 'grandCol')))}`;
};

/** Writes pivot parts for a workbook. Sheets call `sheetParts`, the workbook calls `workbookParts`. */
export class PivotWriter {
  private calc = new Map<number, { items: CalcItemOut[]; extra: Map<number, string[]>; slots: Map<string, number> }>();
  private tableCount = 0;
  readonly caches: PivotCache[];

  constructor(private readonly wb: Workbook, private readonly sink: PartSink) {
    const used = new Set<number>();
    for (const s of wb.sheets) for (const p of s.pivots ?? []) used.add(p.cacheId);
    for (const sc of wb.slicerCaches ?? []) if (sc.source.kind === 'pivot') used.add(sc.source.cacheId);
    this.caches = (wb.pivotCaches ?? []).filter((c) => used.has(c.id));
    // Calculated items become extra shared items in their cache.
    for (const c of this.caches) {
      const info = { items: [] as CalcItemOut[], extra: new Map<number, string[]>(), slots: new Map<string, number>() };
      for (const s of wb.sheets) for (const pt of s.pivots ?? []) {
        if (pt.cacheId !== c.id) continue;
        pt.fields.forEach((pf, f) => pf.items.forEach((it, slot) => {
          if (it.x >= 0 || !it.formula || !c.fields[f]) return;
          const extra = info.extra.get(f) ?? [];
          const name = it.name ?? `Formula${extra.length + 1}`;
          let i = extra.findIndex((e) => e.toLowerCase() === name.toLowerCase());
          if (i < 0) { extra.push(name); i = extra.length - 1; info.items.push({ field: f, index: baseCount(c, f) + i, formula: it.formula }); }
          info.extra.set(f, extra);
          info.slots.set(`${pt.name}|${f}|${slot}`, baseCount(c, f) + i);
        }));
      }
      this.calc.set(c.id, info);
    }
  }

  get has(): boolean { return this.caches.length > 0; }

  cachePath(id: number): string { return `xl/pivotCache/pivotCacheDefinition${this.caches.findIndex((c) => c.id === id) + 1}.xml`; }

  /** Writes the sheet's PivotTable parts and adds their relationships. */
  sheetParts(si: number, addRel: (type: string, target: string) => string): void {
    const s = this.wb.sheets[si]!;
    for (const pt of s.pivots ?? []) {
      const cache = this.caches.find((c) => c.id === pt.cacheId);
      if (!cache) continue;
      const n = ++this.tableCount;
      const info = this.calc.get(cache.id)!;
      const xml = tableXml(pt, cache, this.sink, (f, slot) => info.slots.get(`${pt.name}|${f}|${slot}`) ?? 0);
      this.sink.put(`xl/pivotTables/pivotTable${n}.xml`, xml, PIVOT_CT.table);
      const ci = this.caches.indexOf(cache) + 1;
      this.sink.put(`xl/pivotTables/_rels/pivotTable${n}.xml.rels`, relsXml([[PIVOT_RT.cacheDef, `../pivotCache/pivotCacheDefinition${ci}.xml`]]));
      addRel(PIVOT_RT.table, `../pivotTables/pivotTable${n}.xml`);
    }
  }

  /** Writes cache parts; returns the <pivotCaches> element for workbook.xml. */
  workbookParts(addRel: (type: string, target: string) => string): string {
    if (!this.caches.length) return '';
    const list = this.caches.map((c, i) => {
      const n = i + 1;
      const info = this.calc.get(c.id)!;
      this.sink.put(`xl/pivotCache/pivotCacheDefinition${n}.xml`, cacheDefXml(this.wb, c, this.sink.numFmtId, info.items, info.extra), PIVOT_CT.cacheDef);
      this.sink.put(`xl/pivotCache/pivotCacheRecords${n}.xml`, recordsXml(this.wb, c), PIVOT_CT.records);
      this.sink.put(`xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`, relsXml([[PIVOT_RT.records, `pivotCacheRecords${n}.xml`]]));
      const rid = addRel(PIVOT_RT.cacheDef, `pivotCache/pivotCacheDefinition${n}.xml`);
      return el('pivotCache', { cacheId: c.id, 'r:id': rid });
    });
    return el('pivotCaches', {}, list.join(''));
  }
}

const baseCount = (c: PivotCache, f: number): number => (c.fields[f]?.database === false ? 0 : c.fields[f]?.items.length ?? 0);

export const relsXml = (list: [string, string][]): string =>
  `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([t, target], i) => el('Relationship', { Id: `rId${i + 1}`, Type: t, Target: target })).join('')}</Relationships>`;

// ---------------------------------------------------------------- reading

export interface PartSource {
  xml(path: string): XNode | undefined;
  rels(part: string): { id: string; type: string; target: string; external: boolean }[];
}

const numFmtsOf = (pkg: PartSource, stylesPath: string | undefined): Map<number, string> => {
  const m = new Map<number, string>();
  const root = stylesPath ? pkg.xml(stylesPath) : undefined;
  for (const n of kids(kid(root, 'numFmts'), 'numFmt')) { const id = numAttr(n, 'numFmtId'); const code = attr(n, 'formatCode'); if (id !== undefined && code) m.set(id, code); }
  return m;
};

const fmtOf = (fmts: Map<number, string>, id: number | undefined): string | undefined => {
  if (!id) return undefined;
  const code = fmts.get(id) ?? builtinFormat(id);
  return code && code !== 'General' ? code : undefined;
};

const BY_FILE: Record<string, FieldGroup['range'] extends infer T ? T extends { by: infer B } ? B : never : never> = {
  range: 'range', years: 'years', quarters: 'quarters', months: 'months', days: 'days', hours: 'hours', minutes: 'minutes', seconds: 'seconds',
};

const readCache = (pkg: PartSource, path: string, id: number, fmts: Map<number, string>, wb: Workbook): { cache: PivotCache; calc: { field: number; index: number; formula: string }[] } | null => {
  const def = pkg.xml(path);
  if (!def || def.name !== 'pivotCacheDefinition') return null;
  const src = kid(kid(def, 'cacheSource'), 'worksheetSource');
  let source: PivotCache['source'];
  const sname = attr(src, 'name');
  const ref = attr(src, 'ref');
  if (sname) source = { kind: 'name', name: sname };
  else {
    const r = ref ? parseRange(ref.replace(/\$/g, '')) : null;
    if (!r) return null;
    source = { kind: 'range', sheet: attr(src, 'sheet') ?? wb.sheets[0]?.name ?? 'Sheet1', range: r };
  }
  const date1904 = wb.date1904;
  const fields: CacheField[] = kids(kid(def, 'cacheFields'), 'cacheField').map((cf) => {
    const si = kid(cf, 'sharedItems');
    const items = si ? si.children.map((n) => readItem(n, date1904)) : [];
    const f: CacheField = { name: attr(cf, 'name') ?? 'Field', items };
    const fmt = fmtOf(fmts, numAttr(cf, 'numFmtId'));
    if (fmt) f.numFmt = fmt;
    if (boolAttr(si, 'containsDate') && !boolAttr(si, 'containsNonDate', true)) f.dates = true;
    else if (si && si.children.some((n) => n.name === 'd')) f.dates = true;
    const formula = attr(cf, 'formula');
    if (formula) f.formula = formula.replace(/^=/, '');
    if (attr(cf, 'databaseField') === '0') f.database = false;
    const fg = kid(cf, 'fieldGroup');
    if (fg) {
      const g: FieldGroup = {
        base: numAttr(fg, 'base') ?? 0,
        items: (kid(fg, 'groupItems')?.children ?? []).map((n) => { const v = readItem(n, date1904); return v === null ? '(blank)' : isError(v) ? v.code : String(v); }),
      };
      const par = numAttr(fg, 'par');
      if (par !== undefined) g.parent = par;
      const rp = kid(fg, 'rangePr');
      if (rp) {
        const by = BY_FILE[attr(rp, 'groupBy') ?? 'range'] ?? 'range';
        const sd = attr(rp, 'startDate'), ed = attr(rp, 'endDate');
        const dates = by !== 'range' || !!sd;
        const start = dates ? (sd ? serialOf(sd, date1904) ?? 0 : 0) : numAttr(rp, 'startNum') ?? 0;
        let end = dates ? (ed ? serialOf(ed, date1904) ?? 0 : 0) : numAttr(rp, 'endNum') ?? 0;
        if (dates && by !== 'range') end = Math.max(start, end - 1);
        g.range = { by: by as 'range' | DateGroupBy, start, end, autoStart: attr(rp, 'autoStart') !== '0', autoEnd: attr(rp, 'autoEnd') !== '0', interval: numAttr(rp, 'groupInterval') ?? 1, dates };
      }
      const dp = kid(fg, 'discretePr');
      if (dp) g.discrete = kids(dp, 'x').map((x) => numAttr(x, 'v') ?? 0);
      f.group = g;
    }
    return f;
  });
  const calc = kids(kid(def, 'calculatedItems'), 'calculatedItem').flatMap((ci) => {
    const ref0 = kid(kid(kid(ci, 'pivotArea'), 'references'), 'reference');
    const field = numAttr(ref0, 'field'), index = numAttr(kid(ref0, 'x'), 'v');
    const formula = attr(ci, 'formula');
    return field !== undefined && index !== undefined && formula ? [{ field, index, formula: formula.replace(/^=/, '') }] : [];
  });
  // Records.
  const recRel = pkg.rels(path).find((r) => r.type.endsWith('/pivotCacheRecords') || r.type === 'pivotCacheRecords');
  const recX = recRel ? pkg.xml(recRel.target) : undefined;
  const db = fields.filter((f) => f.database !== false);
  const records: Scalar[][] = kids(recX, 'r').map((r) => {
    const out: Scalar[] = [];
    r.children.forEach((n, j) => {
      const f = db[j];
      if (!f) return;
      out.push(n.name === 'x' ? f.items[numAttr(n, 'v') ?? 0] ?? null : readItem(n, date1904));
    });
    while (out.length < db.length) out.push(null);
    return out;
  });
  // Shared items that are only calculated item names leave the cache.
  const calcByField = new Map<number, Set<number>>();
  for (const c of calc) { const s = calcByField.get(c.field) ?? new Set(); s.add(c.index); calcByField.set(c.field, s); }
  const cache: PivotCache = { id, source, fields, records };
  if (boolAttr(def, 'refreshOnLoad')) cache.refreshOnLoad = true;
  const by = attr(def, 'refreshedBy');
  if (by) cache.refreshedBy = by;
  if (attr(def, 'missingItemsLimit') === '0') cache.dropMissing = true;
  for (const [f, set] of calcByField) {
    const cf = fields[f];
    if (!cf) continue;
    cf.items = cf.items.filter((_, i) => !set.has(i));
  }
  return { cache, calc };
};

const FILE_AGG: Record<string, PivotAgg> = { sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };
const TYPE_LABEL = Object.fromEntries(Object.entries(LABEL_TYPE).map(([k, v]) => [v, k])) as Record<string, LabelOp>;
const TYPE_VALUE = Object.fromEntries(Object.entries(VALUE_TYPE).map(([k, v]) => [v, k])) as Record<string, ValueOp>;
const SHOW_AS = new Set<ShowAs>(['normal', 'percentOfTotal', 'percentOfCol', 'percentOfRow', 'percent', 'difference', 'percentDiff', 'runTotal', 'index']);

const readTable = (x: XNode, cache: PivotCache, calc: { field: number; index: number; formula: string }[], fmts: Map<number, string>, origItems: Map<number, number>, dxfStyle?: (id: number) => number | undefined): PivotTable | null => {
  if (x.name !== 'pivotTableDefinition') return null;
  const locN = kid(x, 'location');
  const ref = parseRange((attr(locN, 'ref') ?? '').replace(/\$/g, ''));
  if (!ref) return null;
  const tabular = attr(x, 'outline') !== '1' && attr(x, 'compact') === '0';
  const compactT = attr(x, 'compact') !== '0';
  const layout: PivotTable['layout'] = compactT ? 'compact' : tabular ? 'tabular' : 'outline';
  const firstHeaderRow = numAttr(locN, 'firstHeaderRow') ?? 1, firstDataRow = numAttr(locN, 'firstDataRow') ?? 1, firstDataCol = numAttr(locN, 'firstDataCol') ?? 0;
  const pageRowCount = numAttr(locN, 'rowPageCount') ?? 0;
  const calcFor = (f: number, x0: number) => calc.find((c) => c.field === f && c.index === x0);
  const fields: PivotField[] = kids(kid(x, 'pivotFields'), 'pivotField').map((pfN, f) => {
    const pf: PivotField = { items: [] };
    const n = attr(pfN, 'name');
    if (n) pf.name = n;
    const its = kids(kid(pfN, 'items'), 'item');
    const subs: Subtotal[] = [];
    if (attr(pfN, 'defaultSubtotal') !== '0') subs.push('default');
    for (const [k, attrName] of Object.entries(SUB_ATTR)) if (boolAttr(pfN, attrName)) subs.push(k as Subtotal);
    if (!(subs.length === 1 && subs[0] === 'default')) pf.subtotals = subs;
    const removedBefore = (x0: number) => { let c = 0; for (const cx of calc) if (cx.field === f && cx.index < x0) c++; return c; };
    for (const it of its) {
      const t = attr(it, 't');
      if (t && t !== 'data') continue;
      const xi = numAttr(it, 'x') ?? 0;
      const item: PivotItem = { x: xi };
      const ci = calcFor(f, xi);
      if (ci || attr(it, 'f') === '1') {
        item.x = -1;
        item.formula = ci?.formula ?? '0';
        item.name = attr(it, 'n') ?? `Formula${pf.items.filter((p) => p.x < 0).length + 1}`;
      } else item.x = xi - removedBefore(xi);
      if (attr(it, 'n') && item.x >= 0) item.name = attr(it, 'n');
      if (boolAttr(it, 'h')) item.hidden = true;
      if (attr(it, 'sd') === '0') item.collapsed = true;
      pf.items.push(item);
    }
    if (!its.length) pf.items = [];
    if (attr(pfN, 'compact') === '0') pf.compact = false; else if (layout !== 'compact') pf.compact = true;
    if (attr(pfN, 'outline') === '0') pf.outline = false; else if (layout === 'tabular') pf.outline = true;
    if (attr(pfN, 'subtotalTop') === '0') pf.subtotalTop = false;
    if (boolAttr(pfN, 'insertBlankRow')) pf.blankRow = true;
    if (boolAttr(pfN, 'insertPageBreak')) pf.pageBreak = true;
    if (boolAttr(pfN, 'showAll', true) && attr(pfN, 'showAll') === '1') pf.showAll = true;
    if (boolAttr(pfN, 'multipleItemSelectionAllowed')) pf.multiPage = true;
    const st = attr(pfN, 'sortType');
    if (st === 'ascending' || st === 'descending') pf.sort = st;
    const fmt = fmtOf(fmts, numAttr(pfN, 'numFmtId'));
    if (fmt) pf.numFmt = fmt;
    const scope = kid(kid(kid(kid(pfN, 'autoSortScope'), 'pivotArea'), 'references'), 'reference');
    if (scope && attr(scope, 'field') === '4294967294') pf.sortByData = numAttr(kid(scope, 'x'), 'v') ?? 0;
    for (const ext of kids(kid(pfN, 'extLst'), 'ext')) for (const c of ext.children) if (c.name.endsWith('pivotField') && attr(c, 'fillDownLabels') === '1') pf.repeatLabels = true;
    void origItems;
    return pf;
  });
  while (fields.length < cache.fields.length) fields.push({ items: [] });
  const fieldList = (tag: string) => kids(kid(x, tag), 'field').map((n) => numAttr(n, 'x') ?? 0).filter((f) => f === VALUES_FIELD || (f >= 0 && f < cache.fields.length));
  const axisItems = (tag: string): AxisItem[] => kids(kid(x, tag), 'i').map((n) => {
    const it: AxisItem = { x: kids(n, 'x').map((v) => numAttr(v, 'v') ?? 0) };
    const t = attr(n, 't');
    if (t) it.t = t as ItemType;
    const r = numAttr(n, 'r');
    if (r) it.r = r;
    const i = numAttr(n, 'i');
    if (i) it.i = i;
    return it;
  });
  const data: DataField[] = kids(kid(x, 'dataFields'), 'dataField').map((d) => {
    const field = numAttr(d, 'fld') ?? 0;
    const df: DataField = { name: attr(d, 'name') ?? '', field, agg: FILE_AGG[attr(d, 'subtotal') ?? 'sum'] ?? 'sum' };
    const sa = attr(d, 'showDataAs') as ShowAs | undefined;
    if (sa && SHOW_AS.has(sa) && sa !== 'normal') df.showAs = sa;
    for (const ext of kids(kid(d, 'extLst'), 'ext')) for (const c of ext.children) if (c.name.endsWith('dataField')) { const p = attr(c, 'pivotShowAs'); if (p) df.showAs = p as ShowAs; }
    if (df.showAs && !['percentOfTotal', 'percentOfRow', 'percentOfCol', 'index', 'percentOfParentRow', 'percentOfParentCol'].includes(df.showAs)) {
      df.baseField = numAttr(d, 'baseField') ?? 0;
      const bi = numAttr(d, 'baseItem') ?? 0;
      df.baseItem = bi === 1048828 ? 'prev' : bi === 1048829 ? 'next' : bi;
    }
    const fmt = fmtOf(fmts, numAttr(d, 'numFmtId'));
    if (fmt) df.numFmt = fmt;
    return df;
  });
  const si = kid(x, 'pivotTableStyleInfo');
  const pt: PivotTable = {
    name: attr(x, 'name') ?? 'PivotTable1',
    cacheId: cache.id,
    anchor: { row: ref.r1, col: ref.c1 },
    location: { ref, firstHeaderRow, firstDataRow, firstDataCol, pageRows: pageRowCount ? pageRowCount + 1 : undefined },
    fields,
    rows: fieldList('rowFields'),
    cols: fieldList('colFields'),
    pages: kids(kid(x, 'pageFields'), 'pageField').map((p) => ({ field: numAttr(p, 'fld') ?? 0, item: numAttr(p, 'item') })).filter((p) => p.field >= 0 && p.field < cache.fields.length),
    data,
    layout,
    rowGrand: attr(x, 'rowGrandTotals') !== '0',
    colGrand: attr(x, 'colGrandTotals') !== '0',
    style: { name: attr(si, 'name') ?? '', rowHeaders: boolAttr(si, 'showRowHeaders'), colHeaders: boolAttr(si, 'showColHeaders'), bandedRows: boolAttr(si, 'showRowStripes'), bandedCols: boolAttr(si, 'showColStripes'), lastColumn: boolAttr(si, 'showLastColumn', true) },
    preserveFormatting: attr(x, 'preserveFormatting') !== '0',
    autoFit: boolAttr(x, 'useAutoFormatting'),
    showDrill: attr(x, 'showDrill') !== '0',
    showHeaders: attr(x, 'showHeaders') !== '0',
    rowItems: axisItems('rowItems'),
    colItems: axisItems('colItems'),
  };
  for (const p of pt.pages) if (p.item !== undefined && !fields[p.field]?.items[p.item]) delete p.item;
  const opt = (k: keyof PivotTable, a: string) => { const v = attr(x, a); if (v !== undefined) (pt as unknown as Record<string, unknown>)[k] = v; };
  opt('dataCaption', 'dataCaption'); opt('grandTotalCaption', 'grandTotalCaption'); opt('rowHeaderCaption', 'rowHeaderCaption'); opt('colHeaderCaption', 'colHeaderCaption');
  const err = attr(x, 'errorCaption');
  if (err !== undefined) pt.errorText = err;
  if (boolAttr(x, 'showError')) pt.showError = true;
  const miss = attr(x, 'missingCaption');
  if (miss !== undefined) { pt.emptyText = miss; pt.showEmpty = attr(x, 'showMissing') !== '0'; }
  const ind = numAttr(x, 'indent');
  if (ind !== undefined) pt.indent = ind;
  const pw = numAttr(x, 'pageWrap');
  if (pw) pt.pageWrap = pw;
  if (boolAttr(x, 'pageOverThenDown')) pt.pageOverThenDown = true;
  // Filters.
  for (const fN of kids(kid(x, 'filters'), 'filter')) {
    const fld = numAttr(fN, 'fld') ?? -1;
    const pf = fields[fld];
    if (!pf) continue;
    const type = attr(fN, 'type') ?? '';
    let flt: PivotFilter | null = null;
    if (TYPE_LABEL[type]) flt = { kind: 'label', op: TYPE_LABEL[type]!, value1: attr(fN, 'stringValue1') ?? '', value2: attr(fN, 'stringValue2') };
    else if (TYPE_VALUE[type]) {
      const cfs = kids(kid(kid(kid(fN, 'autoFilter'), 'filterColumn'), 'customFilters'), 'customFilter');
      const v1 = Number(attr(cfs[0], 'val') ?? 0), v2 = cfs[1] ? Number(attr(cfs[1], 'val')) : undefined;
      flt = { kind: 'value', op: TYPE_VALUE[type]!, dataField: numAttr(fN, 'iMeasureFld') ?? 0, value1: v1, value2: v2 };
    } else if (type === 'dateBetween') {
      const cfs = kids(kid(kid(kid(fN, 'autoFilter'), 'filterColumn'), 'customFilters'), 'customFilter');
      const v1 = Number(attr(cfs[0], 'val') ?? 0), v2 = Number(attr(cfs[1], 'val') ?? v1);
      flt = { kind: 'dates', start: v1, end: v2, timeline: attr(fN, 'name') };
    } else if (type === 'count' || type === 'percent' || type === 'sum') {
      const t10 = kid(kid(kid(fN, 'autoFilter'), 'filterColumn'), 'top10');
      flt = { kind: 'top', top: attr(t10, 'top') !== '0', by: type, value: numAttr(t10, 'val') ?? 10, dataField: numAttr(fN, 'iMeasureFld') ?? 0 };
    }
    if (flt) pf.filters = [...(pf.filters ?? []), flt];
  }
  if (dxfStyle) {
    for (const f of kids(kid(x, 'formats'), 'format')) {
      const id = numAttr(f, 'dxfId');
      const area = kid(f, 'pivotArea');
      if (id === undefined || !area || attr(f, 'action') === 'blank') continue;
      const k = areaKey(area, pt);
      const st = dxfStyle(id);
      if (k && st !== undefined) pt.formats = [...(pt.formats ?? []).filter((q) => q.key !== k), { key: k, style: st }];
    }
  }
  for (const ext of kids(kid(x, 'extLst'), 'ext')) {
    if (attr(ext, 'uri') !== EXT_CLERDO) continue;
    for (const c of ext.children) {
      if (c.name.endsWith('distinctCount')) for (const i of (attr(c, 'fields') ?? '').split(/\s+/).filter(Boolean).map(Number)) { if (pt.data[i]) pt.data[i]!.agg = 'distinctCount'; }
      if (c.name.endsWith('defer')) pt.deferLayout = true;
      // Our own copy is only a fallback for areas the standard formats didn't give.
      if (c.name.endsWith('format')) { const k = attr(c, 'key'), st = numAttr(c, 's'); if (k && st !== undefined && st >= 0 && !pt.formats?.some((f) => f.key === k)) pt.formats = [...(pt.formats ?? []), { key: k, style: st }]; }
    }
  }
  return pt;
};

/**
 * Reads pivot caches and tables. `sheetPaths[i]` is the part of sheet i. Tables whose cache is
 * missing are skipped (their cells stay as plain values).
 */
export const readPivotParts = (pkg: PartSource, wb: Workbook, wbx: XNode, wbPath: string, sheetPaths: string[], stylesPath?: string, dxfs: readonly CellStyle[] = []): void => {
  const dxfStyle = (id: number): number | undefined => (dxfs[id] ? internStyle(wb, dxfs[id]!) : undefined);
  const pcs = kids(kid(wbx, 'pivotCaches'), 'pivotCache');
  if (!pcs.length) return;
  const fmts = numFmtsOf(pkg, stylesPath);
  const rels = new Map(pkg.rels(wbPath).map((r) => [r.id, r]));
  const byPath = new Map<string, { cache: PivotCache; calc: { field: number; index: number; formula: string }[] }>();
  const caches: PivotCache[] = [];
  for (const pc of pcs) {
    const id = numAttr(pc, 'cacheId');
    const rel = rels.get(attr(pc, 'id') ?? '');
    if (id === undefined || !rel || rel.external) continue;
    try {
      const got = readCache(pkg, rel.target, id, fmts, wb);
      if (!got) continue;
      byPath.set(rel.target.toLowerCase(), got);
      caches.push(got.cache);
    } catch { /* a broken cache leaves its tables as plain cells */ }
  }
  if (!caches.length) return;
  wb.pivotCaches = caches;
  sheetPaths.forEach((sp, si) => {
    if (!sp) return;
    for (const r of pkg.rels(sp)) {
      if (r.external || !(r.type === 'pivotTable' || r.type.endsWith('/pivotTable'))) continue;
      const x = pkg.xml(r.target);
      const cRel = pkg.rels(r.target).find((q) => q.type === 'pivotCacheDefinition' || q.type.endsWith('/pivotCacheDefinition'));
      const got = cRel ? byPath.get(cRel.target.toLowerCase()) : undefined;
      if (!x || !got) continue;
      try {
        const pt = readTable(x, got.cache, got.calc, fmts, new Map(), dxfStyle);
        if (pt) { pt.cacheId = got.cache.id; const s = wb.sheets[si]!; s.pivots = [...(s.pivots ?? []), pt]; }
      } catch { /* keep the cells */ }
    }
  });
};

export const PIVOT_PARTS = /^xl\/(pivotCache|pivotTables)\//i;
export type { CellRange };
