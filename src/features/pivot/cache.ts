// Pivot caches: the snapshot of source records a PivotTable summarizes, plus grouping.
import { compareValues } from '../../engine/coerce';
import { fromSerial, timeOf, toSerial } from '../../engine/date';
import { CellError, isError, type Scalar } from '../../engine/types';
import type { CacheField, DateGroupBy, FieldGroup, PivotCache, PivotSource } from '../../model/pivot';

/** Items compare without regard to case, like other apps do. */
export const itemKey = (v: Scalar): string => {
  if (v === null || v === '') return 'b';
  if (typeof v === 'number') return `n${v}`;
  if (typeof v === 'boolean') return v ? 'T' : 'F';
  if (isError(v)) return `e${v.code}`;
  return `s${v.toLowerCase()}`;
};

export const compareItems = (a: Scalar, b: Scalar): number => {
  if (isError(a) || isError(b)) return isError(a) === isError(b) ? (isError(a) && isError(b) ? a.code.localeCompare(b.code) : 0) : isError(a) ? 1 : -1;
  return compareValues(a === '' ? null : a, b === '' ? null : b);
};

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const dateText = (serial: number, date1904 = false): string => {
  const d = fromSerial(Math.floor(serial), date1904);
  return `${d.month}/${d.day}/${d.year}`;
};

const trimNum = (n: number): string => String(+n.toPrecision(12));

/** Group captions for a range grouping, including the "before start" and "after end" items. */
export const rangeItems = (r: NonNullable<FieldGroup['range']>, date1904 = false): string[] => {
  const lo = r.dates ? `<${dateText(r.start, date1904)}` : `<${trimNum(r.start)}`;
  const hi = r.dates ? `>${dateText(r.end, date1904)}` : `>${trimNum(r.end)}`;
  const mid: string[] = [];
  switch (r.by) {
    case 'years': {
      const a = fromSerial(Math.floor(r.start), date1904).year, b = fromSerial(Math.floor(r.end), date1904).year;
      for (let y = a; y <= b && mid.length < 10000; y++) mid.push(String(y));
      break;
    }
    case 'quarters': mid.push('Qtr1', 'Qtr2', 'Qtr3', 'Qtr4'); break;
    case 'months': mid.push(...MONTHS); break;
    case 'days':
      if (r.interval > 1) {
        for (let d = Math.floor(r.start); d <= r.end && mid.length < 10000; d += r.interval) mid.push(`${dateText(d, date1904)} - ${dateText(Math.min(d + r.interval - 1, Math.floor(r.end)), date1904)}`);
      } else {
        // Every day of a leap year, so each date has a slot.
        for (let m = 0; m < 12; m++) for (let d = 1; d <= [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m]!; d++) mid.push(`${d}-${MONTHS[m]}`);
      }
      break;
    case 'hours': for (let h = 0; h < 24; h++) mid.push(`${h % 12 || 12} ${h < 12 ? 'AM' : 'PM'}`); break;
    case 'minutes': for (let m = 0; m < 60; m++) mid.push(`:${String(m).padStart(2, '0')}`); break;
    case 'seconds': for (let s = 0; s < 60; s++) mid.push(`:${String(s).padStart(2, '0')}`); break;
    default: {
      const step = r.interval > 0 ? r.interval : 1;
      const ints = Number.isInteger(r.start) && Number.isInteger(step);
      for (let a = r.start; a <= r.end && mid.length < 10000; a += step) mid.push(`${trimNum(a)}-${trimNum(ints ? a + step - 1 : a + step)}`);
    }
  }
  return [lo, ...mid, hi];
};

const DAY_OF_YEAR_START = [0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335];

/** Index of a source value in a range grouping's items (0 = before start, last = after end, -1 = blank/text). */
export const rangeIndex = (r: NonNullable<FieldGroup['range']>, v: Scalar, count: number, date1904 = false): number => {
  if (typeof v !== 'number') return -1;
  if (r.dates ? Math.floor(v) < Math.floor(r.start) : v < r.start) return 0;
  if (r.dates ? Math.floor(v) > Math.floor(r.end) : v > r.end) return count - 1;
  switch (r.by) {
    case 'years': return 1 + fromSerial(Math.floor(v), date1904).year - fromSerial(Math.floor(r.start), date1904).year;
    case 'quarters': return 1 + Math.floor((fromSerial(Math.floor(v), date1904).month - 1) / 3);
    case 'months': return fromSerial(Math.floor(v), date1904).month;
    case 'days': {
      if (r.interval > 1) return 1 + Math.floor((Math.floor(v) - Math.floor(r.start)) / r.interval);
      const d = fromSerial(Math.floor(v), date1904);
      return 1 + DAY_OF_YEAR_START[d.month - 1]! + d.day - 1;
    }
    case 'hours': return 1 + timeOf(v).hour;
    case 'minutes': return 1 + timeOf(v).minute;
    case 'seconds': return 1 + timeOf(v).second;
    default: {
      const step = r.interval > 0 ? r.interval : 1;
      return Math.min(count - 2, 1 + Math.floor((v - r.start) / step + 1e-9));
    }
  }
};

/** Source values of a database field (records hold only database fields, in order). */
export const dbIndex = (cache: PivotCache, field: number): number => {
  let n = 0;
  for (let i = 0; i < field; i++) if (cache.fields[i]!.database !== false) n++;
  return cache.fields[field]?.database === false ? -1 : n;
};

const indexCache = new WeakMap<PivotCache, Int32Array[]>();

/**
 * For each cache field, each record's item index (into the field's shared items or its group items).
 * -1 for calculated fields. Memoized per cache object (caches are never changed in place).
 */
export const recordItems = (cache: PivotCache, date1904 = false): Int32Array[] => {
  const hit = indexCache.get(cache);
  if (hit) return hit;
  const n = cache.records.length;
  const out: Int32Array[] = [];
  const raw = (f: number): Int32Array => {
    const field = cache.fields[f]!;
    const di = dbIndex(cache, f);
    const a = new Int32Array(n).fill(-1);
    if (di < 0) return a;
    const map = new Map<string, number>();
    field.items.forEach((v, i) => { const k = itemKey(v); if (!map.has(k)) map.set(k, i); });
    for (let r = 0; r < n; r++) a[r] = map.get(itemKey(cache.records[r]![di] ?? null)) ?? -1;
    return a;
  };
  const rawCache = new Map<number, Int32Array>();
  const rawOf = (f: number) => { let a = rawCache.get(f); if (!a) { a = raw(f); rawCache.set(f, a); } return a; };
  cache.fields.forEach((field, f) => {
    const g = field.group;
    if (!g) { out[f] = rawOf(f); return; }
    const a = new Int32Array(n).fill(-1);
    const baseField = cache.fields[g.base];
    const di = dbIndex(cache, g.base);
    if (g.range && baseField && di >= 0) {
      const count = g.items.length;
      for (let r = 0; r < n; r++) {
        const v = cache.records[r]![di] ?? null;
        const i = rangeIndex(g.range, v, count, date1904);
        a[r] = i < 0 ? -1 : Math.max(0, Math.min(count - 1, i));
      }
      // Blank and text values get their own items after the groups when present.
      if (field.items.length) {
        const map = new Map<string, number>();
        field.items.forEach((v, i) => map.set(itemKey(v), i));
        for (let r = 0; r < n; r++) if (a[r]! < 0) { const k = map.get(itemKey(cache.records[r]![di] ?? null)); if (k !== undefined) a[r] = count + k; }
      }
    } else if (g.discrete) {
      const base = baseField?.group && g.base !== f ? out[g.base] ?? rawOf(g.base) : rawOf(g.base);
      for (let r = 0; r < n; r++) { const b = base[r]!; a[r] = b >= 0 ? g.discrete[b] ?? -1 : -1; }
    }
    out[f] = a;
  });
  indexCache.set(cache, out);
  return out;
};

/** Captions shown for a field's items: group items when grouped, otherwise shared items. */
export const fieldItems = (field: CacheField): Scalar[] => {
  if (field.group?.range) return [...field.group.items, ...field.items];
  if (field.group) return field.group.items;
  return field.items;
};

const looksDate = (fmt: string | undefined): boolean => !!fmt && /[dmy]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '')) && !/^[#0.,%E+\-\s]*$/.test(fmt);

export interface SourceData { header: Scalar[]; rows: Scalar[][]; formats: (string | undefined)[] }

const uniqueName = (name: string, used: Set<string>): string => {
  let n = name.trim() || 'Column';
  if (used.has(n.toLowerCase())) { let i = 2; while (used.has(`${n}${i}`.toLowerCase())) i++; n = `${n}${i}`; }
  used.add(n.toLowerCase());
  return n;
};

const headerText = (v: Scalar, i: number): string => (v === null || v === '' ? `Column${i + 1}` : isError(v) ? v.code : String(v));

const distinctSorted = (vals: Scalar[]): Scalar[] => {
  const seen = new Map<string, Scalar>();
  for (const v of vals) { const k = itemKey(v); if (!seen.has(k)) seen.set(k, v === '' ? null : v); }
  return [...seen.values()].sort(compareItems);
};

/** A new cache from a block of source data. */
export const buildCache = (id: number, source: PivotSource, data: SourceData): PivotCache => {
  const used = new Set<string>();
  const fields: CacheField[] = data.header.map((h, i) => {
    const col = data.rows.map((r) => r[i] ?? null);
    const fmt = data.formats[i];
    const nums = col.filter((v) => v !== null && v !== '');
    const dates = looksDate(fmt) && nums.length > 0 && nums.every((v) => typeof v === 'number');
    const f: CacheField = { name: uniqueName(headerText(h, i), used), items: distinctSorted(col) };
    if (fmt && fmt !== 'General') f.numFmt = fmt;
    if (dates) f.dates = true;
    return f;
  });
  return { id, source, fields, records: data.rows.map((r) => data.header.map((_, i) => normalize(r[i] ?? null))), refreshedDate: Date.now() };
};

const normalize = (v: Scalar): Scalar => (v === '' ? null : v);

export interface Remap {
  /** Old field index to new field index (-1 when the field is gone). */
  fields: number[];
  /** Per old field: old item index to new item index (-1 when gone). */
  items: number[][];
}

/**
 * Re-reads the source: keeps fields by name and items by value so PivotTables keep their settings,
 * adds new items after the old ones, and rebuilds groupings.
 */
export const refreshCache = (old: PivotCache, data: SourceData, date1904 = false): { cache: PivotCache; remap: Remap } => {
  const fresh = buildCache(old.id, old.source, data);
  const byName = new Map(old.fields.map((f, i) => [f.name.toLowerCase(), i]));
  const fields: CacheField[] = [];
  const remapFields = old.fields.map(() => -1);
  const remapItems: number[][] = old.fields.map(() => []);
  fresh.fields.forEach((nf) => {
    const oi = byName.get(nf.name.toLowerCase());
    const of = oi !== undefined ? old.fields[oi] : undefined;
    if (of && of.database !== false) {
      const freshKeys = new Set(nf.items.map(itemKey));
      const keep = old.dropMissing ? of.items.filter((v) => freshKeys.has(itemKey(v))) : of.items.slice();
      const keepKeys = new Set(keep.map(itemKey));
      const items = [...keep, ...nf.items.filter((v) => !keepKeys.has(itemKey(v)))];
      const pos = new Map(items.map((v, i) => [itemKey(v), i]));
      remapItems[oi!] = of.items.map((v) => pos.get(itemKey(v)) ?? -1);
      remapFields[oi!] = fields.length;
      fields.push({ ...nf, items, numFmt: nf.numFmt ?? of.numFmt, dates: nf.dates ?? of.dates, group: of.group });
    } else fields.push(nf);
  });
  // Grouped and calculated fields go after the database fields.
  old.fields.forEach((f, i) => {
    if (f.database !== false) return;
    remapFields[i] = fields.length;
    fields.push({ ...f });
  });
  const cache: PivotCache = { ...old, fields, records: fresh.records, refreshedDate: Date.now() };
  // Fix group references and rebuild groups against the new items.
  cache.fields = cache.fields.map((f) => {
    if (!f.group) return f;
    const base = remapFields[f.group.base] ?? -1;
    if (base < 0) return { ...f, group: undefined };
    const parent = f.group.parent !== undefined ? remapFields[f.group.parent] : undefined;
    return { ...f, group: { ...f.group, base, parent: parent !== undefined && parent >= 0 ? parent : undefined } };
  });
  for (let i = 0; i < cache.fields.length; i++) {
    const f = cache.fields[i]!;
    const g = f.group;
    if (!g) continue;
    if (g.range) cache.fields[i] = { ...f, group: rebuildRange(cache, g, date1904) };
    else if (g.discrete) {
      const oldIdx = old.fields.findIndex((x) => x.name === f.name);
      const oldBaseIdx = oldIdx >= 0 ? old.fields[oldIdx]!.group?.base : undefined;
      const oldBase = oldBaseIdx !== undefined ? old.fields[oldBaseIdx] : undefined;
      const oldMap = new Map<string, number>();
      if (oldBase && oldIdx >= 0) fieldItems(oldBase).forEach((v, j) => { const gi = old.fields[oldIdx]!.group!.discrete![j]; if (gi !== undefined) oldMap.set(itemKey(v), gi); });
      const baseItems = fieldItems(cache.fields[g.base]!);
      const items = g.items.slice();
      const discrete = baseItems.map((v) => {
        const hit = oldMap.get(itemKey(v));
        if (hit !== undefined) return hit;
        items.push(v === null ? '(blank)' : isError(v) ? v.code : String(v));
        return items.length - 1;
      });
      cache.fields[i] = { ...f, group: { ...g, items, discrete } };
    }
  }
  return { cache, remap: { fields: remapFields, items: remapItems } };
};

const rebuildRange = (cache: PivotCache, g: FieldGroup, date1904: boolean): FieldGroup => {
  const r = { ...g.range! };
  const di = dbIndex(cache, g.base);
  if (di >= 0 && (r.autoStart || r.autoEnd)) {
    let lo = Infinity, hi = -Infinity;
    for (const rec of cache.records) { const v = rec[di]; if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); } }
    if (lo <= hi) {
      if (r.autoStart) r.start = r.dates ? Math.floor(lo) : lo;
      if (r.autoEnd) r.end = r.dates ? Math.floor(hi) : hi;
    }
  }
  return { ...g, range: r, items: rangeItems(r, date1904) };
};

export interface GroupRange { by: 'range' | DateGroupBy; start?: number; end?: number; interval?: number }

/**
 * Groups a field by ranges: dates by one or more calendar units, or numbers by an interval. Returns a
 * new cache plus the indexes of the fields that now hold each unit (finest unit stays on the field).
 */
export const groupByRange = (cache: PivotCache, field: number, units: GroupRange[], date1904 = false): { cache: PivotCache; fields: number[] } => {
  const base = cache.fields[field]!;
  const di = dbIndex(cache, field);
  let lo = Infinity, hi = -Infinity;
  if (di >= 0) for (const rec of cache.records) { const v = rec[di]; if (typeof v === 'number') { lo = Math.min(lo, v); hi = Math.max(hi, v); } }
  if (lo > hi) { lo = 0; hi = 0; }
  const dates = !!base.dates && units.some((u) => u.by !== 'range');
  const order: (GroupRange['by'])[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years'];
  const sorted = [...units].sort((a, b) => order.indexOf(a.by) - order.indexOf(b.by));
  const fields = cache.fields.map((f) => ({ ...f }));
  // Drop old derived range groups of this field.
  const mk = (u: GroupRange): FieldGroup['range'] => {
    const start = u.start ?? (dates ? Math.floor(lo) : lo);
    const end = u.end ?? (dates ? Math.floor(hi) : hi);
    return { by: u.by, start, end, autoStart: u.start === undefined, autoEnd: u.end === undefined, interval: u.interval ?? (u.by === 'range' ? 10 : 1), dates };
  };
  const out: number[] = [];
  const firstRange = mk(sorted[0]!)!;
  fields[field] = { ...base, group: { base: field, range: firstRange, items: rangeItems(firstRange, date1904) } };
  out.push(field);
  let prev = field;
  for (const u of sorted.slice(1)) {
    const r = mk(u)!;
    const name = uniqueFieldName(fields, u.by === 'range' ? `${base.name}2` : u.by[0]!.toUpperCase() + u.by.slice(1));
    fields.push({ name, items: [], database: false, group: { base: field, range: r, items: rangeItems(r, date1904) } });
    const idx = fields.length - 1;
    fields[prev] = { ...fields[prev]!, group: { ...fields[prev]!.group!, parent: idx } };
    prev = idx;
    out.push(idx);
  }
  return { cache: { ...cache, fields }, fields: out };
};

export const uniqueFieldName = (fields: { name: string }[], base: string): string => {
  const taken = new Set(fields.map((f) => f.name.toLowerCase()));
  if (!taken.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!taken.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
};

/** Removes range grouping from a field and drops the derived unit fields. Returns the new cache and a field remap. */
export const ungroupRange = (cache: PivotCache, field: number): { cache: PivotCache; remap: number[] } => {
  const drop = new Set<number>();
  for (let p = cache.fields[field]?.group?.parent; p !== undefined; p = cache.fields[p]?.group?.parent) drop.add(p);
  cache.fields.forEach((f, i) => { if (f.database === false && f.group?.range && f.group.base === field) drop.add(i); });
  const remap: number[] = [];
  const fields: CacheField[] = [];
  cache.fields.forEach((f, i) => { if (drop.has(i)) { remap.push(-1); return; } remap.push(fields.length); fields.push(i === field ? { ...f, group: undefined } : { ...f }); });
  return { cache: { ...cache, fields: fields.map((f) => (f.group ? { ...f, group: { ...f.group, base: remap[f.group.base] ?? f.group.base, parent: f.group.parent !== undefined && remap[f.group.parent]! >= 0 ? remap[f.group.parent] : undefined } } : f)) }, remap };
};

/** Adds a manual grouping field over `field`: the given item sets become named groups, other items stay alone. */
export const groupItems = (cache: PivotCache, field: number, groups: { name: string; items: number[] }[]): { cache: PivotCache; field: number } => {
  const base = cache.fields[field]!;
  const baseItems = fieldItems(base);
  const items: string[] = [];
  const discrete = baseItems.map(() => -1);
  for (const g of groups) { items.push(g.name); const gi = items.length - 1; for (const x of g.items) discrete[x] = gi; }
  baseItems.forEach((v, i) => { if (discrete[i]! < 0) { items.push(v === null ? '(blank)' : isError(v) ? v.code : String(v)); discrete[i] = items.length - 1; } });
  const name = uniqueFieldName(cache.fields, `${base.name}2`);
  const fields = [...cache.fields, { name, items: [], database: false, group: { base: field, discrete, items } } as CacheField];
  return { cache: { ...cache, fields }, field: fields.length - 1 };
};

/** Adds a calculated field. */
export const addCalcField = (cache: PivotCache, name: string, formula: string): PivotCache => ({
  ...cache, fields: [...cache.fields, { name: uniqueFieldName(cache.fields, name), items: [], database: false, formula }],
});

export const groupLabel = (v: Scalar): string => (v === null ? '(blank)' : isError(v) ? v.code : String(v));

export const dateOf = (y: number, m: number, d: number, date1904 = false): number => toSerial(y, m, d, date1904);
export const errorItem = (code: string): CellError => new CellError(code as CellError['code']);
