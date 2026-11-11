import type { CSSProperties, ReactNode } from 'react';
import { quoteSheet } from '../../engine/parser';
import { isError, isMatrix, MAX_COLS, MAX_ROWS, type Matrix, type Scalar } from '../../engine/types';
import { colName, parseRange, type CellRange, type CellStyle } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { T, Tr } from '../i18n';

export const PALETTE = [
  '#000000', '#3f3f46', '#71717a', '#a1a1aa', '#d4d4d8', '#ffffff',
  '#b91c1c', '#ef4444', '#f97316', '#f59e0b', '#facc15', '#84cc16',
  '#16a34a', '#0d9488', '#0891b2', '#2563eb', '#2347ff', '#7c3aed',
  '#fde2e2', '#ffedd5', '#fef9c3', '#dcfce7', '#dbeafe', '#ede9fe',
];

/** Small segmented tab strip used inside dialogs. */
export const Tabs = <T extends string>({ tabs, value, onChange }: { tabs: readonly (readonly [T, string])[]; value: T; onChange: (v: T) => void }) => (
  <div className="tabsrow" role="tablist">
    {tabs.map(([id, label]) => (
      <button key={id} role="tab" aria-selected={value === id} className={value === id ? 'on' : ''} onClick={() => onChange(id)}>{Tr(label)}</button>
    ))}
  </div>
);

export const Check = ({ label, checked, onChange, disabled }: { label: ReactNode; checked: boolean; onChange: (v: boolean) => void; disabled?: boolean }) => (
  <label className="b-check"><input type="checkbox" checked={checked} disabled={disabled} onChange={(e) => onChange(e.target.checked)} /> <span>{Tr(label)}</span></label>
);

export const Radio = <T extends string>({ name, value, options, onChange, row }: { name: string; value: T; options: readonly (readonly [T, string])[]; onChange: (v: T) => void; row?: boolean }) => (
  <div className={row ? 'b-radios row' : 'b-radios'} role="radiogroup">
    {options.map(([v, l]) => (
      <label key={v}><input type="radio" name={name} checked={value === v} onChange={() => onChange(v)} /> {Tr(l)}</label>
    ))}
  </div>
);

/** Swatch row with an optional "no color" button and a custom color picker. */
export const ColorPick = ({ value, onChange, none, label }: { value: string | undefined; onChange: (c: string | undefined) => void; none?: string; label?: string }) => (
  <div className="b-colors" aria-label={Tr(label)}>
    {none !== undefined && <button type="button" className={`sw none${!value ? ' on' : ''}`} title={Tr(none)} aria-label={Tr(none)} onClick={() => onChange(undefined)} />}
    {PALETTE.map((c) => (
      <button key={c} type="button" className={`sw${value?.toLowerCase() === c ? ' on' : ''}`} style={{ background: c }} title={c} aria-label={c} onClick={() => onChange(c)} />
    ))}
    <label className="sw custom" title={T('More colors')} aria-label={T('More colors')}>
      <input type="color" value={value && /^#[0-9a-f]{6}$/i.test(value) ? value : '#000000'} onChange={(e) => onChange(e.target.value)} />
    </label>
  </div>
);

/** CSS for previewing a cell style in the dialogs. */
export const cssOf = (s: CellStyle, scaleFont = true): CSSProperties => {
  const f = s.font ?? {};
  const deco = [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  const g = s.fill?.gradient;
  const side = (b?: { style: string; color?: string }) => (b ? `${b.style === 'thick' ? 3 : b.style.startsWith('medium') ? 2 : 1}px ${b.style === 'double' ? 'double' : /dash|slant/i.test(b.style) ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid'} ${b.color ?? '#000'}` : undefined);
  return {
    fontFamily: f.name ? `'${f.name}', sans-serif` : undefined,
    fontSize: f.size && scaleFont ? `${Math.max(8, Math.min(f.size, 26)) * (4 / 3)}px` : undefined,
    fontWeight: f.bold ? 700 : undefined,
    fontStyle: f.italic ? 'italic' : undefined,
    textDecoration: deco || undefined,
    textDecorationStyle: f.underline?.startsWith('double') ? 'double' : undefined,
    color: f.color,
    verticalAlign: f.vertAlign === 'superscript' ? 'super' : f.vertAlign === 'subscript' ? 'sub' : undefined,
    background: g ? `linear-gradient(${g.angle + 90}deg, ${g.stops.map((x) => `${x.color} ${x.pos * 100}%`).join(', ')})` : s.fill?.color,
    borderTop: side(s.border?.top), borderBottom: side(s.border?.bottom), borderLeft: side(s.border?.left), borderRight: side(s.border?.right),
  };
};

/** "$A$1" or "$A$1:$B$4". */
export const absAddr = (r: CellRange): string => {
  const a = (row: number, col: number) => `$${colName(col)}$${row + 1}`;
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `$${colName(r.c1)}:$${colName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `$${r.r1 + 1}:$${r.r2 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a(r.r1, r.c1) : `${a(r.r1, r.c1)}:${a(r.r2, r.c2)}`;
};

export const sheetRef = (store: GridStore, sheet: number, r: CellRange): string => `${quoteSheet(store.wb.sheets[sheet]!.name)}!${absAddr(r)}`;

/** Reads "A1:B4", "$A$1", "Sheet2!B5:C9", "'My sheet'!A1" and comma-separated lists of them. */
export const parseRefs = (store: GridStore, text: string): { sheet: number; ranges: CellRange[] } | null => {
  const parts = text.replace(/^=/, '').split(',').map((p) => p.trim()).filter(Boolean);
  if (!parts.length) return null;
  let sheet = store.sheetIndex;
  const ranges: CellRange[] = [];
  for (const p of parts) {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(p);
    let body = p;
    if (m) {
      const si = store.engine.sheetIndex((m[1] ?? m[2] ?? '').replace(/''/g, "'"));
      if (si < 0) return null;
      sheet = si;
      body = m[3]!;
    }
    const clean = body.replace(/\$/g, '').toUpperCase();
    let r = parseRange(clean);
    if (!r && /^[A-Z]{1,3}:[A-Z]{1,3}$/.test(clean)) { const [a, b] = clean.split(':'); const x = parseRange(`${a}1`), y = parseRange(`${b}1`); if (x && y) r = { r1: 0, r2: MAX_ROWS - 1, c1: Math.min(x.c1, y.c1), c2: Math.max(x.c1, y.c1) }; }
    if (!r && /^\d+:\d+$/.test(clean)) { const [a, b] = clean.split(':').map(Number); if (a! > 0 && b! > 0) r = { r1: Math.min(a!, b!) - 1, r2: Math.max(a!, b!) - 1, c1: 0, c2: MAX_COLS - 1 }; }
    if (!r) return null;
    ranges.push(r);
  }
  return { sheet, ranges };
};

export const goTo = (store: GridStore, sheet: number, row: number, col: number): void => {
  if (sheet !== store.sheetIndex) store.setActiveSheet(sheet);
  store.selectRange({ r1: row, c1: col, r2: row, c2: col });
};

export const scalarText = (v: Scalar): string => {
  if (v === null) return '';
  if (isError(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(+v.toPrecision(15));
  return v;
};

/** Short display of a formula result, including arrays. */
export const resultText = (v: Scalar | Matrix): string => {
  if (!isMatrix(v)) return typeof v === 'string' ? `"${v}"` : scalarText(v);
  const rows = v.slice(0, 4).map((r) => r.slice(0, 6).map((x) => (typeof x === 'string' ? `"${x}"` : scalarText(x))).join(', '));
  const more = v.length > 4 || (v[0]?.length ?? 0) > 6 ? '; …' : '';
  return `{${rows.join('; ')}${more}}`;
};

export const overlaps = (a: CellRange, b: CellRange): boolean => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;

/** Parts of `a` that lie outside `b` (up to four rectangles). */
export const subtractRange = (a: CellRange, b: CellRange): CellRange[] => {
  if (!overlaps(a, b)) return [a];
  const out: CellRange[] = [];
  if (a.r1 < b.r1) out.push({ ...a, r2: b.r1 - 1 });
  if (a.r2 > b.r2) out.push({ ...a, r1: b.r2 + 1 });
  const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2);
  if (a.c1 < b.c1) out.push({ r1, r2, c1: a.c1, c2: b.c1 - 1 });
  if (a.c2 > b.c2) out.push({ r1, r2, c1: b.c2 + 1, c2: a.c2 });
  return out;
};

/** Removes undefined fields and empty groups so equal styles intern to the same entry. */
export const cleanStyle = (s: CellStyle): CellStyle => {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(s)) {
    if (v === undefined) continue;
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      const inner = Object.fromEntries(Object.entries(v as Record<string, unknown>).filter(([, x]) => x !== undefined));
      if (Object.keys(inner).length) out[k] = inner;
    } else out[k] = v;
  }
  return out as CellStyle;
};

export const rangeLabel = (r: CellRange): string => absAddr(r).replace(/\$/g, '');

/** "Column B" or the header text of a column. */
export const columnLabel = (store: GridStore, r: CellRange, col: number, header: boolean): string => {
  if (header) { const v = store.value(r.r1, col); if (v !== null && scalarText(v).trim()) return scalarText(v); }
  return `Column ${colName(col)}`;
};
