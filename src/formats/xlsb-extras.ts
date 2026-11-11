// Reads ClerdoGrid's own part in .xlsb files: settings with no binary record written here (color
// scales, data bars, icon sets, tables and filter criteria). The JSON is untrusted and checked field by field.
import type { AutoFilter, CellRange, CfOperator, CfRule, CfValue, ConditionalFormat, FilterColumn, TableDef, Workbook } from '../model/workbook';

const MAX = 1_048_576;
const str = (v: unknown, max = 32767): string | undefined => (typeof v === 'string' ? v.slice(0, max) : undefined);
const int = (v: unknown, max = MAX): number | undefined => (Number.isInteger(v) && (v as number) >= 0 && (v as number) < max ? (v as number) : undefined);
const colorOk = (v: unknown): string | undefined => (typeof v === 'string' && /^#[0-9a-f]{6}([0-9a-f]{2})?$/i.test(v) ? v : undefined);
const obj = (v: unknown): Record<string, unknown> | undefined => (v && typeof v === 'object' && !Array.isArray(v) ? (v as Record<string, unknown>) : undefined);
const list = (v: unknown, max: number): unknown[] => (Array.isArray(v) ? v.slice(0, max) : []);

const range = (v: unknown): CellRange | undefined => {
  const o = obj(v);
  if (!o) return undefined;
  const r1 = int(o.r1), r2 = int(o.r2), c1 = int(o.c1, 16384), c2 = int(o.c2, 16384);
  return r1 !== undefined && r2 !== undefined && c1 !== undefined && c2 !== undefined && r1 <= r2 && c1 <= c2 ? { r1, r2, c1, c2 } : undefined;
};

const CFVO_TYPES = new Set(['min', 'max', 'num', 'percent', 'percentile', 'formula', 'autoMin', 'autoMax']);
const cfvo = (v: unknown): CfValue | undefined => {
  const o = obj(v);
  if (!o || typeof o.type !== 'string' || !CFVO_TYPES.has(o.type)) return undefined;
  const value = str(o.value, 8192);
  return { type: o.type as CfValue['type'], ...(value !== undefined ? { value } : {}) };
};

const rule = (v: unknown): CfRule | undefined => {
  const o = obj(v);
  if (!o) return undefined;
  if (o.type === 'colorScale') {
    const stops = list(o.stops, 3).map((s) => { const x = obj(s); const value = cfvo(x?.value), color = colorOk(x?.color); return value && color ? { value, color } : undefined; });
    return stops.length >= 2 && stops.every(Boolean) ? { type: 'colorScale', stops: stops as { value: CfValue; color: string }[] } : undefined;
  }
  if (o.type === 'dataBar') {
    const min = cfvo(o.min), max = cfvo(o.max), color = colorOk(o.color);
    if (!min || !max || !color) return undefined;
    return { type: 'dataBar', min, max, color, ...(o.gradient === false ? { gradient: false } : {}), ...(o.showValue === false ? { showValue: false } : {}) };
  }
  if (o.type === 'iconSet') {
    const set = str(o.set, 40);
    const values = list(o.values, 5).map(cfvo);
    if (!set || !/^[0-9A-Za-z]+$/.test(set) || !values.every(Boolean)) return undefined;
    return { type: 'iconSet', set, values: values as CfValue[], ...(o.reverse === true ? { reverse: true } : {}), ...(o.showValue === false ? { showValue: false } : {}) };
  }
  return undefined;
};

const table = (v: unknown): TableDef | undefined => {
  const o = obj(v);
  const name = str(o?.name, 255);
  const r = range(o?.range);
  if (!o || !name || !/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(name) || !r) return undefined;
  const columns = list(o.columns, 16384).map((c) => {
    const x = obj(c);
    const n = str(x?.name, 255);
    if (n === undefined) return undefined;
    const col: TableDef['columns'][number] = { name: n };
    const tf = str(x?.totalsFunction, 8192), cf = str(x?.calculatedFormula, 8192);
    if (tf) col.totalsFunction = tf;
    if (cf) col.calculatedFormula = cf;
    return col;
  });
  if (!columns.every(Boolean) || columns.length !== r.c2 - r.c1 + 1) return undefined;
  const style = str(o.style, 255);
  return {
    name, range: r, headerRow: o.headerRow !== false, totalsRow: o.totalsRow === true, columns: columns as TableDef['columns'],
    ...(style ? { style } : {}), ...(typeof o.bandedRows === 'boolean' ? { bandedRows: o.bandedRows } : {}), ...(typeof o.bandedCols === 'boolean' ? { bandedCols: o.bandedCols } : {}),
  };
};

const OPS = new Set(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']);
const filter = (v: unknown, base: CellRange): AutoFilter | undefined => {
  const o = obj(v);
  if (!o) return undefined;
  const columns = list(o.columns, 16384).map((c): FilterColumn | undefined => {
    const x = obj(c);
    const col = int(x?.col, 16384);
    if (!x || col === undefined) return undefined;
    const out: FilterColumn = { col };
    if (Array.isArray(x.values)) out.values = x.values.filter((s): s is string => typeof s === 'string').slice(0, 100000).map((s) => s.slice(0, 32767));
    if (x.blanks === true) out.blanks = true;
    const cu = obj(x.custom);
    if (cu) {
      const rules = list(cu.rules, 2).map((r) => { const y = obj(r); return y && typeof y.operator === 'string' && OPS.has(y.operator) ? { operator: y.operator as CfOperator, value: String(y.value ?? '').slice(0, 255) } : undefined; });
      if (rules.length && rules.every(Boolean)) out.custom = { and: cu.and === true, rules: rules as NonNullable<FilterColumn['custom']>['rules'] };
    }
    const t = obj(x.top10);
    if (t && typeof t.value === 'number' && Number.isFinite(t.value)) out.top10 = { top: t.top !== false, percent: t.percent === true, value: t.value };
    const color = colorOk(x.color);
    if (color) out.color = color;
    return out;
  });
  const sort = list(o.sort, 64).map((s) => { const x = obj(s); const col = int(x?.col, 16384); return col === undefined ? undefined : { col, ...(x?.descending === true ? { descending: true } : {}) }; });
  return {
    range: base, columns: columns.filter((c): c is FilterColumn => !!c),
    ...(sort.length && sort.every(Boolean) ? { sort: sort as NonNullable<AutoFilter['sort']> } : {}),
  };
};

/** Applies the part to a workbook just read; anything malformed is dropped. */
export const readExtras = (wb: Workbook, text: string | undefined): void => {
  if (!text || text.length > 64 * 1024 * 1024) return;
  let doc: Record<string, unknown> | undefined;
  try { doc = obj(JSON.parse(text)); } catch { return; }
  if (!doc) return;
  for (const [idx, v] of Object.entries(obj(doc.cf) ?? {})) {
    const s = wb.sheets[Number(idx)];
    if (!s) continue;
    for (const c of list(v, 10000)) {
      const o = obj(c);
      const ranges = list(o?.ranges, 512).map(range);
      const rules = list(o?.rules, 256).map(rule);
      if (!ranges.length || !ranges.every(Boolean) || !rules.length || !rules.every(Boolean)) continue;
      const cf: ConditionalFormat = { ranges: ranges as CellRange[], rules: rules as CfRule[] };
      s.conditionalFormats.push(cf);
    }
  }
  for (const [idx, v] of Object.entries(obj(doc.tables) ?? {})) {
    const s = wb.sheets[Number(idx)];
    if (!s) continue;
    for (const t of list(v, 1000).map(table)) if (t && !wb.sheets.some((x) => x.tables.some((y) => y.name.toUpperCase() === t.name.toUpperCase()))) s.tables.push(t);
  }
  for (const [idx, v] of Object.entries(obj(doc.filters) ?? {})) {
    const s = wb.sheets[Number(idx)];
    if (!s?.autoFilter) continue;
    const f = filter(v, s.autoFilter.range);
    if (f) s.autoFilter = f;
  }
};
