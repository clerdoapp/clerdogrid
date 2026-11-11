// Creating and changing PivotTable definitions (pure functions over the model).
import {
  DEFAULT_PIVOT_STYLE, VALUES_FIELD,
  type DataField, type PivotAgg, type PivotCache, type PivotField, type PivotTable,
} from '../../model/pivot';
import { fieldItems } from './cache';
import { AGG_LABEL, normalizeFields } from './compute';

export const newPivotTable = (name: string, cache: PivotCache, anchor: { row: number; col: number }): PivotTable => ({
  name,
  cacheId: cache.id,
  anchor,
  fields: cache.fields.map((f) => ({ items: f.formula ? [] : fieldItems(f).map((_, x) => ({ x })) })),
  rows: [], cols: [], pages: [], data: [],
  layout: 'compact', rowGrand: true, colGrand: true,
  style: { ...DEFAULT_PIVOT_STYLE },
  preserveFormatting: true, autoFit: true, showDrill: true, showHeaders: true,
});

export type Area = 'rows' | 'cols' | 'pages' | 'data';

/** The default summary for a field: sum for numbers, count otherwise. */
export const defaultAgg = (cache: PivotCache, field: number): PivotAgg => {
  const f = cache.fields[field];
  if (!f) return 'count';
  if (f.formula) return 'sum';
  if (f.group || f.dates) return 'count';
  const nonBlank = f.items.filter((v) => v !== null && v !== '');
  return nonBlank.length && nonBlank.every((v) => typeof v === 'number') ? 'sum' : 'count';
};

export const uniqueDataName = (pt: PivotTable, cache: PivotCache, base: string, except = -1): string => {
  const taken = new Set([...pt.data.filter((_, i) => i !== except).map((d) => d.name.toLowerCase()), ...cache.fields.map((f) => f.name.toLowerCase())]);
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
};

export const makeDataField = (pt: PivotTable, cache: PivotCache, field: number, agg = defaultAgg(cache, field)): DataField => ({
  name: uniqueDataName(pt, cache, `${AGG_LABEL[agg]} of ${cache.fields[field]?.name ?? ''}`),
  field, agg,
});

const without = (pt: PivotTable, field: number): PivotTable => ({
  ...pt,
  rows: pt.rows.filter((f) => f !== field),
  cols: pt.cols.filter((f) => f !== field),
  pages: pt.pages.filter((p) => p.field !== field),
});

/**
 * Puts a field into an area at a position (or at the end). Data fields can repeat; other areas hold a
 * field once. Moving Σ Values (VALUES_FIELD) between rows and columns is allowed.
 */
export const placeField = (pt: PivotTable, cache: PivotCache, field: number, area: Area, at?: number): PivotTable => {
  if (area === 'data') {
    if (field === VALUES_FIELD) return pt;
    const data = [...pt.data];
    data.splice(at ?? data.length, 0, makeDataField(pt, cache, field));
    return withValuesField({ ...pt, data });
  }
  if (field !== VALUES_FIELD && cache.fields[field]?.formula) return pt;
  const next = without(pt, field);
  const ins = <T,>(list: T[], v: T) => { const l = [...list]; l.splice(Math.min(at ?? l.length, l.length), 0, v); return l; };
  if (area === 'rows') next.rows = ins(next.rows, field);
  else if (area === 'cols') next.cols = ins(next.cols, field);
  else if (field !== VALUES_FIELD) next.pages = ins(next.pages, { field });
  return withValuesField(next);
};

export const removeField = (pt: PivotTable, field: number, area?: Area, dataIndex?: number): PivotTable => {
  if (area === 'data' && dataIndex !== undefined) return withValuesField({ ...pt, data: pt.data.filter((_, i) => i !== dataIndex) });
  if (area === 'data') return withValuesField({ ...pt, data: pt.data.filter((d) => d.field !== field) });
  const next = without(pt, field);
  if (!area) next.data = pt.data.filter((d) => d.field !== field);
  return withValuesField(next);
};

export const moveDataField = (pt: PivotTable, from: number, to: number): PivotTable => {
  const data = [...pt.data];
  const [d] = data.splice(from, 1);
  if (!d) return pt;
  data.splice(Math.max(0, Math.min(to, data.length)), 0, d);
  return { ...pt, data };
};

/** Keeps the Σ Values field in rows or columns only while there are two or more data fields. */
export const withValuesField = (pt: PivotTable): PivotTable => {
  const has = pt.rows.includes(VALUES_FIELD) || pt.cols.includes(VALUES_FIELD);
  if (pt.data.length > 1 && !has) return { ...pt, cols: [...pt.cols, VALUES_FIELD] };
  if (pt.data.length <= 1 && has) return { ...pt, rows: pt.rows.filter((f) => f !== VALUES_FIELD), cols: pt.cols.filter((f) => f !== VALUES_FIELD) };
  return pt;
};

export const areaOf = (pt: PivotTable, field: number): Area | null => {
  if (pt.rows.includes(field)) return 'rows';
  if (pt.cols.includes(field)) return 'cols';
  if (pt.pages.some((p) => p.field === field)) return 'pages';
  if (pt.data.some((d) => d.field === field)) return 'data';
  return null;
};

export const setField = (pt: PivotTable, field: number, patch: Partial<PivotField>): PivotTable => {
  const fields = [...pt.fields];
  fields[field] = { ...(fields[field] ?? { items: [] }), ...patch };
  return { ...pt, fields };
};

/** Applies a layout form to the table and all its fields. */
export const setLayout = (pt: PivotTable, layout: PivotTable['layout']): PivotTable => ({
  ...pt, layout,
  fields: pt.fields.map((f) => ({ ...f, compact: layout === 'compact', outline: layout !== 'tabular' })),
});

/** Expands or collapses every item of a field (or all row/column fields). */
export const setCollapsed = (pt: PivotTable, cache: PivotCache, fields: number[], collapsed: boolean): PivotTable => {
  const norm = normalizeFields(pt, cache);
  const next = [...pt.fields];
  for (const f of fields) if (norm[f]) next[f] = { ...norm[f], items: norm[f].items.map((it) => ({ ...it, collapsed: collapsed || undefined })) };
  return { ...pt, fields: next };
};

/** Remaps a table after its cache changed (refresh, grouping): field and item indexes. */
export const remapTable = (pt: PivotTable, fieldMap: number[], itemMaps: number[][], newCache: PivotCache): PivotTable => {
  const mf = (f: number) => (f === VALUES_FIELD ? f : fieldMap[f] ?? -1);
  const fields: PivotField[] = newCache.fields.map(() => ({ items: [] }));
  pt.fields.forEach((pf, f) => {
    const nf = mf(f);
    if (nf < 0 || !newCache.fields[nf]) return;
    const im = itemMaps[f];
    const items = im ? pf.items.flatMap((it) => (it.x < 0 ? [it] : (im[it.x] ?? -1) >= 0 ? [{ ...it, x: im[it.x]! }] : [])) : pf.items;
    fields[nf] = {
      ...pf, items,
      sortByData: pf.sortByData,
      filters: pf.filters,
    };
  });
  const remapItem = (f: number, slot: number | undefined): number | undefined => {
    if (slot === undefined) return undefined;
    const old = pt.fields[f]?.items[slot];
    if (!old) return undefined;
    const nf = mf(f);
    const nx = old.x < 0 ? -1 : itemMaps[f]?.[old.x] ?? old.x;
    const idx = fields[nf]?.items.findIndex((it) => (old.x < 0 ? it.formula === old.formula : it.x === nx)) ?? -1;
    return idx >= 0 ? idx : undefined;
  };
  return {
    ...pt,
    fields,
    rows: pt.rows.map(mf).filter((f) => f === VALUES_FIELD || f >= 0),
    cols: pt.cols.map(mf).filter((f) => f === VALUES_FIELD || f >= 0),
    pages: pt.pages.flatMap((p) => (mf(p.field) >= 0 ? [{ field: mf(p.field), item: remapItem(p.field, p.item) }] : [])),
    data: pt.data.flatMap((d) => (mf(d.field) >= 0 ? [{ ...d, field: mf(d.field), baseField: d.baseField !== undefined ? (mf(d.baseField) >= 0 ? mf(d.baseField) : undefined) : undefined }] : [])),
  };
};

export { fieldItems };
