// Slicers and timelines: filter state in slicer caches, applied to connected PivotTables or to a table.
import type { GridStore } from '../../app/store';
import { displayText } from '../../app/ops/clipboard';
import { fromSerial } from '../../engine/date';
import { isError, type Scalar } from '../../engine/types';
import type { Sheet, TableDef } from '../../model/workbook';
import type { Slicer, SlicerCache, PivotTable } from '../../model/pivot';
import { cacheOf, drawPivot, layoutOf, findByKey } from './apply';
import { compareItems, fieldItems, recordItems } from './cache';
import { computePivot, normalizeFields } from './compute';
import { N } from '../../ui/i18n';

export const SLICER_STYLES = ['SlicerStyleLight1', 'SlicerStyleLight6', 'SlicerStyleOther1', 'SlicerStyleDark1', 'SlicerStyleDark6'];
export const TIMELINE_STYLES = ['TimeSlicerStyleLight1', 'TimeSlicerStyleLight6', 'TimeSlicerStyleDark1', 'TimeSlicerStyleDark6'];

const cleanName = (s: string): string => s.replace(/[^\p{L}\p{N}_]/gu, '_').replace(/^(\d)/, '_$1') || 'Field';

export const uniqueCacheName = (store: GridStore, base: string): string => {
  const taken = new Set([...(store.wb.slicerCaches ?? []).map((c) => c.name.toLowerCase()), ...store.wb.names.map((n) => n.name.toLowerCase())]);
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 1; ; i++) if (!taken.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
};
const uniqueSlicerName = (store: GridStore, base: string): string => {
  const taken = new Set(store.wb.sheets.flatMap((s) => (s.slicers ?? []).map((x) => x.name.toLowerCase())));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 1; ; i++) if (!taken.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
};

export const cacheByName = (store: GridStore, name: string): SlicerCache | undefined => store.wb.slicerCaches?.find((c) => c.name === name);

/** A place near the table for a new slicer, cascading so several don't stack exactly. */
const placeNear = (store: GridStore, row: number, col: number, w: number, h: number): Slicer['anchor'] => {
  const n = store.sheet.slicers?.length ?? 0;
  return { row, col, dx: 8 + (n % 5) * 24, dy: 8 + (n % 5) * 24, width: w, height: h };
};

export const insertPivotSlicers = (store: GridStore, sheet: number, ptName: string, fields: number[]): void => {
  const pt = store.wb.sheets[sheet]?.pivots?.find((p) => p.name === ptName);
  const cache = pt ? cacheOf(store.wb, pt.cacheId) : undefined;
  if (!pt || !cache || !fields.length) return;
  const area = pt.location?.ref ?? { r1: pt.anchor.row, c1: pt.anchor.col, r2: pt.anchor.row, c2: pt.anchor.col };
  store.change(N('Insert slicer'), 'book', () => {
    const made: string[] = [];
    fields.forEach((f, i) => {
      const fname = cache.fields[f]!.name;
      const name = uniqueCacheName(store, `Slicer_${cleanName(fname)}`);
      const sc: SlicerCache = { name, source: { kind: 'pivot', cacheId: cache.id, field: f, pivots: [{ sheet: store.wb.sheets[sheet]!.name, name: pt.name }] }, hidden: [], sort: 'ascending' };
      // Start from the table's current filter on that field.
      const norm = normalizeFields(pt, cache)[f]!;
      sc.hidden = norm.items.filter((it) => it.hidden && it.x >= 0).map((it) => String(it.x));
      store.wb.slicerCaches = [...(store.wb.slicerCaches ?? []), sc];
      const s = store.sheet;
      const slicer: Slicer = { name: uniqueSlicerName(store, fname), cache: name, caption: fname, kind: 'slicer', anchor: placeNear(store, area.r1, area.c2 + 2 + i, 180, 230), columns: 1, buttonHeight: 26, style: 'SlicerStyleLight6', showCaption: true };
      s.slicers = [...(s.slicers ?? []), slicer];
      made.push(slicer.name);
    });
  });
};

export const insertTableSlicers = (store: GridStore, sheet: number, table: string, columns: number[]): void => {
  const t = store.wb.sheets[sheet]?.tables.find((x) => x.name === table);
  if (!t || !columns.length) return;
  store.change(N('Insert slicer'), 'book', () => {
    columns.forEach((c, i) => {
      const cname = t.columns[c]?.name ?? `Column${c + 1}`;
      const name = uniqueCacheName(store, `Slicer_${cleanName(cname)}`);
      store.wb.slicerCaches = [...(store.wb.slicerCaches ?? []), { name, source: { kind: 'table', table: t.name, column: c }, hidden: [], sort: 'ascending' }];
      const s = store.sheet;
      s.slicers = [...(s.slicers ?? []), { name: uniqueSlicerName(store, cname), cache: name, caption: cname, kind: 'slicer', anchor: placeNear(store, t.range.r1, t.range.c2 + 2 + i, 180, 230), columns: 1, buttonHeight: 26, style: 'SlicerStyleLight6', showCaption: true }];
    });
  });
};

export const insertTimeline = (store: GridStore, sheet: number, ptName: string, field: number): void => {
  const pt = store.wb.sheets[sheet]?.pivots?.find((p) => p.name === ptName);
  const cache = pt ? cacheOf(store.wb, pt.cacheId) : undefined;
  if (!pt || !cache) return;
  const fname = cache.fields[field]!.name;
  const area = pt.location?.ref ?? { r1: pt.anchor.row, c1: pt.anchor.col, r2: pt.anchor.row, c2: pt.anchor.col };
  store.change(N('Insert timeline'), 'book', () => {
    const name = uniqueCacheName(store, `NativeTimeline_${cleanName(fname)}`);
    store.wb.slicerCaches = [...(store.wb.slicerCaches ?? []), { name, source: { kind: 'pivot', cacheId: cache.id, field, pivots: [{ sheet: store.wb.sheets[sheet]!.name, name: pt.name }] }, hidden: [], sort: 'ascending', timeline: { level: 'months' } }];
    const s = store.sheet;
    s.slicers = [...(s.slicers ?? []), { name: uniqueSlicerName(store, fname), cache: name, caption: fname, kind: 'timeline', anchor: placeNear(store, area.r2 + 2, area.c1, 420, 140), columns: 1, buttonHeight: 26, style: 'TimeSlicerStyleLight1', showCaption: true, level: 'months' }];
  });
};

// ---------------------------------------------------------------- items

export interface SlicerItem { key: string; label: string; value: Scalar; selected: boolean; hasData: boolean }

const tableOf = (store: GridStore, name: string): { sheet: number; t: TableDef } | null => {
  for (let i = 0; i < store.wb.sheets.length; i++) { const t = store.wb.sheets[i]!.tables.find((x) => x.name.toLowerCase() === name.toLowerCase()); if (t) return { sheet: i, t }; }
  return null;
};

const tableRows = (t: TableDef): [number, number] => [t.range.r1 + (t.headerRow ? 1 : 0), t.range.r2 - (t.totalsRow ? 1 : 0)];

export const slicerItems = (store: GridStore, sc: SlicerCache): SlicerItem[] => {
  const hidden = new Set(sc.hidden);
  let items: SlicerItem[] = [];
  if (sc.source.kind === 'pivot') {
    const cache = cacheOf(store.wb, sc.source.cacheId);
    const cf = cache?.fields[sc.source.field];
    if (!cache || !cf) return [];
    const vals = fieldItems(cf).slice(0, cf.group?.range ? cf.group.items.length : undefined);
    // Items with data under the connected tables' other filters.
    // An item has data when every connected table still has records for it under its other filters.
    let has = null as Set<number> | null;
    const f = sc.source.field;
    const ri = recordItems(cache, store.wb.date1904)[f]!;
    for (const p of sc.source.pivots) {
      const ref = findByKey(store.wb, `${p.sheet}!${p.name}`);
      if (!ref) continue;
      const norm = normalizeFields(ref.pt, cache);
      const pt: PivotTable = { ...ref.pt, fields: norm.map((x, i) => (i === f ? { ...x, items: x.items.map((it) => ({ ...it, hidden: undefined })), filters: x.filters?.filter((y) => y.kind !== 'dates') } : x)), pages: ref.pt.pages.map((pg) => (pg.field === f ? { field: f } : pg)) };
      const here = new Set<number>();
      for (const r of computePivot(pt, cache, store.wb.date1904).records) here.add(ri[r]!);
      const prev: Set<number> | null = has;
      has = prev ? new Set([...prev].filter((x) => here.has(x))) : here;
    }
    if (!has) { has = new Set(); for (let x = 0; x < vals.length; x++) has.add(x); }
    const hasData = has;
    items = vals.map((v, x) => ({ key: String(x), label: labelOf(v, cf.dates && !cf.group ? (cf.numFmt ?? 'm/d/yyyy') : undefined, store.wb.date1904), value: v, selected: !hidden.has(String(x)), hasData: hasData.has(x) }));
  } else {
    const src = sc.source;
    const at = tableOf(store, src.table);
    if (!at) return [];
    const [r1, r2] = tableRows(at.t);
    const col = at.t.range.c1 + src.column;
    const others = (store.wb.slicerCaches ?? []).filter((c) => c !== sc && c.source.kind === 'table' && c.source.table.toLowerCase() === src.table.toLowerCase() && c.hidden.length);
    const seen = new Map<string, SlicerItem>();
    for (let r = r1; r <= r2; r++) {
      const text = displayText(store, r, col, at.sheet);
      const key = text;
      const visible = others.every((o) => !o.hidden.includes(displayText(store, r, at.t.range.c1 + (o.source as { column: number }).column, at.sheet)));
      const cur = seen.get(key);
      if (cur) { if (visible) cur.hasData = true; continue; }
      seen.set(key, { key, label: text === '' ? N('(blank)') : text, value: store.value(r, col, at.sheet), selected: !hidden.has(key), hasData: visible });
    }
    items = [...seen.values()];
  }
  const dir = sc.sort === 'descending' ? -1 : 1;
  if (sc.sort !== 'none') items.sort((a, b) => dir * (compareItems(a.value, b.value) || a.label.localeCompare(b.label)));
  if (sc.hideNoData) items = items.filter((i) => i.hasData);
  else if (sc.showNoDataLast !== false) items = [...items.filter((i) => i.hasData), ...items.filter((i) => !i.hasData)];
  return items;
};

const labelOf = (v: Scalar, fmt: string | undefined, d1904: boolean): string => {
  if (v === null || v === '') return '(blank)';
  if (isError(v)) return v.code;
  if (typeof v === 'number' && fmt) { const d = fromSerial(Math.floor(v), d1904); return `${d.month}/${d.day}/${d.year}`; }
  return String(v);
};

// ---------------------------------------------------------------- applying filters

/** Applies one cache's state to what it filters. Call inside store.change(). */
export const applySlicerCache = (store: GridStore, sc: SlicerCache): void => {
  if (sc.source.kind === 'table') { applyTableFilters(store, sc.source.table); return; }
  const cache = cacheOf(store.wb, sc.source.cacheId);
  if (!cache) return;
  const f = sc.source.field;
  for (const p of sc.source.pivots) {
    const ref = findByKey(store.wb, `${p.sheet}!${p.name}`);
    if (!ref || ref.pt.cacheId !== cache.id) continue;
    const norm = normalizeFields(ref.pt, cache);
    let pt: PivotTable;
    if (sc.timeline) {
      const tl = sc.timeline;
      const others = (norm[f]!.filters ?? []).filter((x) => !(x.kind === 'dates' && x.timeline === sc.name));
      const filters = tl.start !== undefined && tl.end !== undefined ? [...others, { kind: 'dates' as const, start: tl.start, end: tl.end, timeline: sc.name }] : others;
      pt = { ...ref.pt, fields: norm.map((x, i) => (i === f ? { ...x, filters: filters.length ? filters : undefined } : x)) };
    } else {
      const hidden = new Set(sc.hidden);
      pt = { ...ref.pt, fields: norm.map((x, i) => (i === f ? { ...x, items: x.items.map((it) => ({ ...it, hidden: it.x >= 0 && hidden.has(String(it.x)) ? true : undefined })) } : x)) };
      // A report filter on the same field shows the slicer's choice.
      pt = { ...pt, pages: pt.pages.map((pg) => (pg.field === f ? { field: f } : pg)) };
    }
    const s = store.wb.sheets[ref.sheet]!;
    s.pivots = (s.pivots ?? []).map((x) => (x.name === ref.pt.name ? drawPivot(store, ref.sheet, pt) : x));
  }
};

/** Hides table rows that any slicer on the table leaves out. */
export const applyTableFilters = (store: GridStore, table: string): void => {
  const at = tableOf(store, table);
  if (!at) return;
  const sheet: Sheet = store.wb.sheets[at.sheet]!;
  const [r1, r2] = tableRows(at.t);
  const caches = (store.wb.slicerCaches ?? []).filter((c) => c.source.kind === 'table' && c.source.table.toLowerCase() === table.toLowerCase());
  for (let r = r1; r <= r2; r++) {
    const out = caches.some((c) => c.hidden.length && c.hidden.includes(displayText(store, r, at.t.range.c1 + (c.source as { column: number }).column, at.sheet)));
    const info = sheet.rows.get(r) ?? {};
    if (out) sheet.rows.set(r, { ...info, hidden: true, filtered: true });
    else if (info.filtered && info.hidden) { const n = { ...info }; delete n.hidden; delete n.filtered; sheet.rows.set(r, n); }
  }
};

export const setSlicerHidden = (store: GridStore, name: string, hidden: string[], label = N('Slicer')): void => {
  store.change(label, 'book', () => {
    const sc = cacheByName(store, name);
    if (!sc) return;
    sc.hidden = hidden;
    applySlicerCache(store, sc);
  });
};

export const setTimelineRange = (store: GridStore, name: string, start: number | undefined, end: number | undefined): void => {
  store.change(N('Timeline'), 'book', () => {
    const sc = cacheByName(store, name);
    if (!sc?.timeline) return;
    sc.timeline = { ...sc.timeline, start, end };
    applySlicerCache(store, sc);
  });
};

export const updateSlicer = (store: GridStore, name: string, patch: Partial<Slicer>, cachePatch?: Partial<SlicerCache>, label = N('Slicer settings')): void => {
  store.change(label, 'book', () => {
    for (const s of store.wb.sheets) s.slicers = s.slicers?.map((x) => (x.name === name ? { ...x, ...patch } : x));
    const sl = store.wb.sheets.flatMap((s) => s.slicers ?? []).find((x) => x.name === (patch.name ?? name));
    if (cachePatch && sl) {
      const sc = cacheByName(store, sl.cache);
      if (sc) { Object.assign(sc, cachePatch); if (cachePatch.timeline && sc.timeline) sc.timeline = { ...sc.timeline }; }
    }
  });
};

/** Moving a slicer, without an undo step per pixel: the caller commits once with `updateSlicer`. */
export const slicerOf = (store: GridStore, name: string): Slicer | undefined => store.sheet.slicers?.find((x) => x.name === name);

export const deleteSlicer = (store: GridStore, name: string): void => {
  store.change(N('Delete slicer'), 'book', () => {
    let cacheName = '';
    for (const s of store.wb.sheets) { const hit = s.slicers?.find((x) => x.name === name); if (hit) cacheName = hit.cache; s.slicers = s.slicers?.filter((x) => x.name !== name); if (s.slicers && !s.slicers.length) delete s.slicers; }
    const still = store.wb.sheets.some((s) => s.slicers?.some((x) => x.cache === cacheName));
    if (!still) {
      const sc = cacheByName(store, cacheName);
      if (sc) { sc.hidden = []; if (sc.timeline) sc.timeline = { ...sc.timeline, start: undefined, end: undefined }; applySlicerCache(store, sc); }
      store.wb.slicerCaches = (store.wb.slicerCaches ?? []).filter((c) => c.name !== cacheName);
      if (!store.wb.slicerCaches.length) delete store.wb.slicerCaches;
    }
  });
};

/** Report connections: which tables on the same data the slicer filters. */
export const setConnections = (store: GridStore, cacheName: string, pivots: { sheet: string; name: string }[]): void => {
  store.change(N('Report connections'), 'book', () => {
    const sc = cacheByName(store, cacheName);
    if (!sc || sc.source.kind !== 'pivot') return;
    const before = sc.source.pivots;
    // Tables no longer connected lose this slicer's filter.
    const dropped = before.filter((b) => !pivots.some((p) => p.sheet === b.sheet && p.name === b.name));
    if (dropped.length) {
      const clear: SlicerCache = { ...sc, hidden: [], timeline: sc.timeline ? { ...sc.timeline, start: undefined, end: undefined } : undefined, source: { ...sc.source, pivots: dropped } };
      applySlicerCache(store, clear);
    }
    sc.source = { ...sc.source, pivots };
    applySlicerCache(store, sc);
  });
};

/** Tables that share a cache and could be connected. */
export const connectable = (store: GridStore, cacheId: number): { sheet: string; name: string }[] =>
  store.wb.sheets.flatMap((s) => (s.pivots ?? []).filter((p) => p.cacheId === cacheId).map((p) => ({ sheet: s.name, name: p.name })));

/** Date bounds of a timeline's field. */
export const timelineBounds = (store: GridStore, sc: SlicerCache): [number, number] | null => {
  if (sc.source.kind !== 'pivot') return null;
  const cache = cacheOf(store.wb, sc.source.cacheId);
  if (!cache) return null;
  const base = cache.fields[sc.source.field]?.group?.base ?? sc.source.field;
  const di = cache.fields.slice(0, base).filter((f) => f.database !== false).length;
  let lo = Infinity, hi = -Infinity;
  for (const r of cache.records) { const v = r[di]; if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); } }
  return lo <= hi ? [Math.floor(lo), Math.floor(hi)] : null;
};

export { layoutOf };
