// PivotTable actions behind the commands, menus, field list and dialogs.
import type { GridStore } from '../../app/store';
import { newSheet, type Chart } from '../../model/workbook';
import { VALUES_FIELD, type PivotCache, type PivotTable, type Subtotal } from '../../model/pivot';
import { cacheOf, drawPivot, layoutOf, pivotChartSeries, pivotKey, refreshPivots, replaceCache, updatePivot, type PivotRef } from './apply';
import { addCalcField, groupByRange, groupItems, ungroupRange, type GroupRange } from './cache';
import { normalizeFields } from './compute';
import { placeField, setField, setLayout, setCollapsed, type Area, removeField, remapTable } from './model';
import { focusField, focusOf, currentPivot } from './ui';
import { N, T } from '../../ui/i18n';

export const update = (store: GridStore, ref: PivotRef, fn: (pt: PivotTable) => PivotTable, label = N('Change PivotTable')): boolean => updatePivot(store, ref.sheet, ref.pt.name, fn, label);

const withCache = (store: GridStore, ref: PivotRef): PivotCache | null => cacheOf(store.wb, ref.pt.cacheId) ?? null;

export const addField = (store: GridStore, ref: PivotRef, field: number, area: Area, at?: number): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => placeField(pt, cache, field, area, at), N('Add field'));
};

export const dropField = (store: GridStore, ref: PivotRef, field: number, area?: Area, dataIndex?: number): void => {
  update(store, ref, (pt) => removeField(pt, field, area, dataIndex), N('Remove field'));
};

/** Checking a field in the list: text and dates go to rows, numbers to values. */
export const toggleField = (store: GridStore, ref: PivotRef, field: number, on: boolean): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  if (!on) { update(store, ref, (pt) => removeField(pt, field), N('Remove field')); return; }
  const f = cache.fields[field]!;
  const numeric = !!f.formula || (!f.dates && !f.group && f.items.length > 0 && f.items.every((v) => v === null || typeof v === 'number'));
  addField(store, ref, field, numeric ? 'data' : 'rows');
};

export const setTableLayout = (store: GridStore, ref: PivotRef, layout: PivotTable['layout']): void => { update(store, ref, (pt) => setLayout(pt, layout), N('Report layout')); };

export const setSubtotalPlacement = (store: GridStore, ref: PivotRef, where: 'top' | 'bottom' | 'none'): void => {
  update(store, ref, (pt) => ({
    ...pt,
    fields: pt.fields.map((f) => ({ ...f, subtotals: where === 'none' ? [] : (f.subtotals?.length ? f.subtotals : undefined), subtotalTop: where === 'bottom' ? false : undefined })),
  }), N('Subtotals'));
};

export const setGrandTotals = (store: GridStore, ref: PivotRef, which: 'both' | 'rows' | 'cols' | 'none'): void => {
  update(store, ref, (pt) => ({ ...pt, rowGrand: which === 'both' || which === 'rows', colGrand: which === 'both' || which === 'cols' }), N('Grand totals'));
};

export const setAllFields = (store: GridStore, ref: PivotRef, patch: (f: PivotTable['fields'][number]) => PivotTable['fields'][number], label: string): void => {
  update(store, ref, (pt) => ({ ...pt, fields: pt.fields.map((f, i) => (pt.rows.includes(i) || pt.cols.includes(i) ? patch(f) : f)) }), label);
};

export const setStyle = (store: GridStore, ref: PivotRef, patch: Partial<PivotTable['style']>): void => {
  update(store, ref, (pt) => ({ ...pt, style: { ...pt.style, ...patch } }), N('PivotTable style'));
};

/** Expands or collapses one item (the +/- button). */
export const toggleItem = (store: GridStore, ref: PivotRef, field: number, slot: number): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    const items = norm[field]!.items.map((it, i) => (i === slot ? { ...it, collapsed: !it.collapsed || undefined } : it));
    return setField({ ...pt, fields: norm }, field, { items });
  }, N('Expand or collapse'));
};

export const expandField = (store: GridStore, ref: PivotRef, field: number | null, collapsed: boolean): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  const fields = field !== null ? [field] : [...ref.pt.rows, ...ref.pt.cols].filter((f) => f >= 0);
  update(store, ref, (pt) => setCollapsed(pt, cache, fields, collapsed), collapsed ? N('Collapse field') : N('Expand field'));
};

/** Sorts a field by its labels, or by a data field's values. */
export const sortField = (store: GridStore, ref: PivotRef, field: number, dir: 'ascending' | 'descending' | undefined, byData?: number): void => {
  update(store, ref, (pt) => setField(pt, field, { sort: dir, sortByData: byData }), N('Sort'));
};

/** Moves an item to another position in manual order (dragging a label). */
export const moveItem = (store: GridStore, ref: PivotRef, field: number, from: number, to: number): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    const items = [...norm[field]!.items];
    const [it] = items.splice(from, 1);
    if (!it) return pt;
    items.splice(Math.max(0, Math.min(items.length, to)), 0, it);
    return setField({ ...pt, fields: norm }, field, { items, sort: undefined, sortByData: undefined });
  }, N('Move item'));
};

export const setHiddenItems = (store: GridStore, ref: PivotRef, field: number, hidden: Set<number>, label = N('Filter')): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    return setField({ ...pt, fields: norm }, field, { items: norm[field]!.items.map((it, i) => ({ ...it, hidden: hidden.has(i) || undefined })) });
  }, label);
};

export const clearFieldFilters = (store: GridStore, ref: PivotRef, field: number | null): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    const fields = norm.map((f, i) => (field === null || i === field ? { ...f, filters: undefined, items: f.items.map((it) => ({ ...it, hidden: undefined })) } : f));
    return { ...pt, fields, pages: pt.pages.map((p) => (field === null || p.field === field ? { field: p.field } : p)) };
  }, N('Clear filter'));
};

export const setPageItem = (store: GridStore, ref: PivotRef, field: number, item: number | undefined, hidden?: Set<number>): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    const fields = norm.map((f, i) => (i === field ? { ...f, multiPage: hidden ? true : f.multiPage, items: f.items.map((it, s) => ({ ...it, hidden: hidden?.has(s) || undefined })) } : f));
    return { ...pt, fields, pages: pt.pages.map((p) => (p.field === field ? { field, item } : p)) };
  }, N('Report filter'));
};

export const setSubtotals = (store: GridStore, ref: PivotRef, field: number, subs: Subtotal[]): void => {
  update(store, ref, (pt) => setField(pt, field, { subtotals: subs.length === 1 && subs[0] === 'default' ? undefined : subs }), N('Subtotals'));
};

// ---------------------------------------------------------------- grouping

/** Groups the selected items (manual groups), or a number/date field by ranges. */
export const groupSelection = (store: GridStore): string | null => {
  const f = focusOf(store);
  if (!f) return T('Select a cell in a PivotTable.');
  const field = focusField(f);
  if (field === null) return T('Select items in a row or column field.');
  const cache = f.cache;
  const cf = cache.fields[field]!;
  const base = cf.group?.discrete ? cf.group.base : field;
  // Collect the selected items of that field.
  const slots = new Set<number>();
  const { row: R, col: C } = f.ref.pt.anchor;
  for (const r of store.sel.ranges) for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++) for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 200); col++) {
    const cell = f.layout.cells[row - R]?.[col - C];
    if (cell && cell.field === field && cell.slot !== undefined && (cell.role === 'rowItem' || cell.role === 'colItem') && !cell.total) slots.add(cell.slot);
  }
  const numericOrDate = !cf.group && cf.items.some((v) => typeof v === 'number') && cf.items.every((v) => v === null || typeof v === 'number');
  if (slots.size <= 1 && numericOrDate) return 'range';
  if (!slots.size) return T('Select the items to group.');
  if (field !== base) return T('Ungroup this field first.');
  const pf = normalizeFields(f.ref.pt, cache)[field]!;
  const items = [...slots].map((s) => pf.items[s]!.x).filter((x) => x >= 0);
  const existing = cache.fields.filter((x) => x.group?.discrete && x.group.base === field);
  const name = `Group${existing.length ? existing[0]!.group!.items.filter((n) => /^Group\d+$/.test(n)).length + 1 : 1}`;
  if (existing.length) {
    // Add another group to the existing grouping field.
    const gi = cache.fields.indexOf(existing[0]!);
    const g = existing[0]!.group!;
    const itemsList = [...g.items, name];
    const ni = itemsList.length - 1;
    const discrete = g.discrete!.map((d, x) => (items.includes(x) ? ni : d));
    const used = new Set(discrete);
    const keep = itemsList.map((_, i) => used.has(i));
    const remap: number[] = [];
    const finalItems: string[] = [];
    itemsList.forEach((n, i) => { if (keep[i]) { remap[i] = finalItems.length; finalItems.push(n); } else remap[i] = -1; });
    const nc: PivotCache = { ...cache, fields: cache.fields.map((x, i) => (i === gi ? { ...x, group: { ...g, items: finalItems, discrete: discrete.map((d) => remap[d]!) } } : x)) };
    const itemMaps = cache.fields.map((_, i) => (i === gi ? g.items.map((_, j) => remap[j] ?? -1) : fieldIdentity(cache, i)));
    replaceCache(store, nc, { fields: cache.fields.map((_, i) => i), items: itemMaps }, N('Group'));
    return null;
  }
  const res = groupItems(cache, field, [{ name, items }]);
  replaceCache(store, res.cache, null, N('Group'), {
    sheet: f.ref.sheet, name: f.ref.pt.name,
    fn: (pt) => {
      const area = pt.rows.includes(field) ? 'rows' : 'cols';
      const list = area === 'rows' ? pt.rows : pt.cols;
      return placeField(pt, res.cache, res.field, area, list.indexOf(field));
    },
  });
  return null;
};

const fieldIdentity = (cache: PivotCache, f: number): number[] => {
  const cf = cache.fields[f]!;
  const n = cf.group ? cf.group.items.length + (cf.group.range ? cf.items.length : 0) : cf.items.length;
  return Array.from({ length: n }, (_, i) => i);
};

export const groupRange = (store: GridStore, ref: PivotRef, field: number, units: GroupRange[]): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  const cleared = cache.fields[field]!.group?.range ? ungroupRange(cache, field) : { cache, remap: cache.fields.map((_, i) => i) };
  const g = groupByRange(cleared.cache, field, units, store.wb.date1904);
  const fieldMap = cleared.remap;
  const itemMaps = cache.fields.map((_, i) => (i === field ? [] : fieldIdentity(cache, i)));
  replaceCache(store, g.cache, { fields: fieldMap, items: itemMaps }, N('Group'), {
    sheet: ref.sheet, name: ref.pt.name,
    fn: (pt) => {
      // Put the coarser units next to the field, outermost first.
      let p = setField(pt, field, { items: [] });
      const area = p.rows.includes(field) ? 'rows' : p.cols.includes(field) ? 'cols' : null;
      if (!area) return p;
      const at = (area === 'rows' ? p.rows : p.cols).indexOf(field);
      for (const extra of g.fields.slice(1)) p = placeField(p, g.cache, extra, area, at);
      return p;
    },
  });
};

export const ungroup = (store: GridStore): string | null => {
  const f = focusOf(store);
  let field = focusField(f);
  if (!f || field === null) return T('Select a grouped field in a PivotTable.');
  let cf = f.cache.fields[field]!;
  // On an item of a grouped field, ungroup the grouping field over it.
  if (!cf.group) {
    const over = f.cache.fields.findIndex((x) => x.group?.discrete && x.group.base === field);
    if (over >= 0) { field = over; cf = f.cache.fields[over]!; }
  }
  if (cf.group?.discrete) {
    // Remove the manual grouping field.
    const remap = f.cache.fields.map((_, i) => (i === field ? -1 : i < field ? i : i - 1));
    const nc: PivotCache = { ...f.cache, fields: f.cache.fields.filter((_, i) => i !== field).map((x) => (x.group ? { ...x, group: { ...x.group, base: remap[x.group.base] ?? x.group.base, parent: x.group.parent !== undefined ? remap[x.group.parent] : undefined } } : x)) };
    replaceCache(store, nc, { fields: remap, items: f.cache.fields.map((_, i) => fieldIdentity(f.cache, i)) }, N('Ungroup'));
    return null;
  }
  const baseField = cf.group?.range ? (cf.database === false ? cf.group.base : field) : -1;
  if (baseField < 0) return T('This field isn’t grouped.');
  const u = ungroupRange(f.cache, baseField);
  replaceCache(store, u.cache, { fields: u.remap, items: f.cache.fields.map((_, i) => (i === baseField ? [] : fieldIdentity(f.cache, i))) }, N('Ungroup'), {
    sheet: f.ref.sheet, name: f.ref.pt.name, fn: (pt) => setField(pt, u.remap[baseField]!, { items: [] }),
  });
  return null;
};

// ---------------------------------------------------------------- calculated fields and items

export const saveCalcField = (store: GridStore, ref: PivotRef, name: string, formula: string, existing: number | null): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  if (existing !== null) {
    const nc: PivotCache = { ...cache, fields: cache.fields.map((f, i) => (i === existing ? { ...f, name, formula } : f)) };
    replaceCache(store, nc, null, N('Calculated field'));
    return;
  }
  const nc = addCalcField(cache, name, formula);
  replaceCache(store, nc, null, N('Calculated field'), { sheet: ref.sheet, name: ref.pt.name, fn: (pt) => placeField(pt, nc, nc.fields.length - 1, 'data') });
};

export const deleteCalcField = (store: GridStore, ref: PivotRef, field: number): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  const remap = cache.fields.map((_, i) => (i === field ? -1 : i < field ? i : i - 1));
  const nc: PivotCache = { ...cache, fields: cache.fields.filter((_, i) => i !== field).map((x) => (x.group ? { ...x, group: { ...x.group, base: remap[x.group.base] ?? x.group.base } } : x)) };
  replaceCache(store, nc, { fields: remap, items: cache.fields.map((_, i) => fieldIdentity(cache, i)) }, N('Delete calculated field'));
};

export const saveCalcItem = (store: GridStore, ref: PivotRef, field: number, name: string, formula: string, slot: number | null): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    const items = [...norm[field]!.items];
    if (slot !== null && items[slot]) items[slot] = { ...items[slot]!, name, formula };
    else items.push({ x: -1, name, formula });
    return setField({ ...pt, fields: norm }, field, { items });
  }, N('Calculated item'));
};

export const deleteCalcItem = (store: GridStore, ref: PivotRef, field: number, slot: number): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => {
    const norm = normalizeFields(pt, cache);
    return setField({ ...pt, fields: norm }, field, { items: norm[field]!.items.filter((_, i) => i !== slot) });
  }, N('Delete calculated item'));
};

// ---------------------------------------------------------------- report filter pages, charts, refresh

/** One new sheet per item of a report filter field, each with a copy of the table filtered to that item. */
export const showFilterPages = (store: GridStore, ref: PivotRef, field: number): number => {
  const cache = withCache(store, ref);
  if (!cache) return 0;
  const norm = normalizeFields(ref.pt, cache);
  const pf = norm[field]!;
  const l = layoutOf(store.wb, ref.pt)!;
  const slots = pf.items.map((it, s) => ({ it, s })).filter(({ it }) => !it.hidden && it.x >= 0);
  let made = 0;
  store.change(N('Show report filter pages'), 'book', () => {
    const at = ref.sheet;
    for (const { s } of [...slots].reverse()) {
      const caption = l.caption(field, s).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31).trim() || 'Page';
      const taken = new Set(store.wb.sheets.map((x) => x.name.toLowerCase()));
      let name = caption, i = 2;
      while (taken.has(name.toLowerCase())) name = `${caption.slice(0, 28)} (${i++})`;
      store.wb.sheets.splice(at, 0, newSheet(name));
      store.engine.sheetsChanged();
      const pt: PivotTable = { ...structuredClone(ref.pt), fields: norm, name: `${ref.pt.name}_${made + 1}`, pages: ref.pt.pages.map((p) => (p.field === field ? { field, item: s } : p)), location: undefined, formats: undefined };
      store.wb.sheets[at]!.pivots = [drawPivot(store, at, pt)];
      made++;
    }
    store.wb.activeSheet = at;
  });
  return made;
};

let chartSeq = 0;
/** A chart linked to the table: it follows the table's layout whenever the table changes. */
export const insertPivotChart = (store: GridStore, ref: PivotRef, type: Chart['type'] = 'column'): void => {
  const l = layoutOf(store.wb, ref.pt);
  if (!l) return;
  const series = pivotChartSeries(store, ref.sheet, ref.pt, l);
  const area = ref.pt.location?.ref ?? { r1: ref.pt.anchor.row, c1: ref.pt.anchor.col, r2: ref.pt.anchor.row, c2: ref.pt.anchor.col };
  const chart: Chart = {
    id: `pc${Date.now().toString(36)}${chartSeq++}`, type, series,
    anchor: { row: area.r1, col: area.c2 + 2, dx: 0, dy: 0, width: 480, height: 290 },
    legend: series.length > 1 ? 'right' : 'none', pivot: pivotKey(store.wb, ref.sheet, ref.pt),
    title: series.length === 1 ? undefined : undefined,
  };
  store.change(N('PivotChart'), 'sheet', () => { store.wb.sheets[ref.sheet]!.charts.push(chart); });
};

export const refreshCurrent = (store: GridStore): void => {
  const ref = currentPivot(store);
  if (!ref) { store.notify(T('Select a cell in a PivotTable.')); return; }
  const err = refreshPivots(store, [ref.pt.cacheId]);
  if (err) store.notify(err);
};

export const refreshAll = (store: GridStore): void => {
  if (!store.wb.pivotCaches?.length) { store.notify(T('This workbook has no PivotTables.')); return; }
  const err = refreshPivots(store, undefined, N('Refresh all'));
  if (err) store.notify(err);
};

/** Puts Σ Values in rows or columns. */
export const moveValues = (store: GridStore, ref: PivotRef, area: 'rows' | 'cols'): void => {
  const cache = withCache(store, ref);
  if (!cache) return;
  update(store, ref, (pt) => placeField(pt, cache, VALUES_FIELD, area), N('Move values'));
};

export { remapTable };
