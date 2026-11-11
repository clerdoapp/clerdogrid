// PivotTable style presets in ClerdoGrid's green palette. Files store the preset by name
// (pivotTableStyleInfo); cells also get explicit formatting so every app shows the same look.
import type { CellStyle } from '../../model/workbook';
import type { PivotStyleInfo } from '../../model/pivot';
import type { PCell } from './compute';
import { N } from '../../ui/i18n';

export interface PivotPreset {
  /** Name written to files. */
  name: string;
  label: string;
  header: CellStyle;
  rowHeader: CellStyle;
  sub: CellStyle;
  grand: CellStyle;
  band: CellStyle;
  page: CellStyle;
}

const thin = (color: string) => ({ style: 'thin' as const, color });
const medium = (color: string) => ({ style: 'medium' as const, color });

const GREENS = {
  ink: '#0b3d26', deep: '#0c5c39', main: '#0c7a4b', mid: '#5fb387', soft: '#b9e0c9', pale: '#e1f3e9', faint: '#f1f9f4',
};

const preset = (name: string, label: string, p: Omit<PivotPreset, 'name' | 'label'>): PivotPreset => ({ name, label, ...p });

export const PRESETS: PivotPreset[] = [
  preset('PivotStyleLight1', N('Plain'), {
    header: { font: { bold: true }, border: { bottom: thin('#9aa3b8') } }, rowHeader: { font: { bold: true } },
    sub: { font: { bold: true } }, grand: { font: { bold: true }, border: { top: thin('#9aa3b8') } }, band: {}, page: { font: { bold: true } },
  }),
  preset('PivotStyleLight16', N('Light green'), {
    header: { font: { bold: true, color: GREENS.main }, border: { bottom: thin(GREENS.mid) } },
    rowHeader: { font: { bold: true } },
    sub: { font: { bold: true } },
    grand: { font: { bold: true }, border: { top: thin(GREENS.mid) } },
    band: { fill: { color: GREENS.faint } },
    page: { font: { bold: true, color: GREENS.main } },
  }),
  preset('PivotStyleLight21', N('Light green, filled header'), {
    header: { font: { bold: true, color: GREENS.ink }, fill: { color: GREENS.pale }, border: { bottom: thin(GREENS.main) } },
    rowHeader: { font: { bold: true } },
    sub: { font: { bold: true }, border: { bottom: thin(GREENS.soft) } },
    grand: { font: { bold: true }, fill: { color: GREENS.pale }, border: { top: medium(GREENS.main) } },
    band: { fill: { color: GREENS.faint } },
    page: { font: { bold: true }, fill: { color: GREENS.pale } },
  }),
  preset('PivotStyleMedium7', N('Medium green'), {
    header: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.main } },
    rowHeader: { font: { bold: true } },
    sub: { font: { bold: true }, fill: { color: GREENS.pale } },
    grand: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.deep } },
    band: { fill: { color: GREENS.faint } },
    page: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.main } },
  }),
  preset('PivotStyleMedium14', N('Medium green, soft'), {
    header: { font: { bold: true, color: GREENS.ink }, fill: { color: GREENS.soft } },
    rowHeader: { font: { bold: true, color: GREENS.ink } },
    sub: { font: { bold: true }, fill: { color: GREENS.pale } },
    grand: { font: { bold: true, color: GREENS.ink }, fill: { color: GREENS.soft }, border: { top: thin(GREENS.main) } },
    band: { fill: { color: GREENS.faint } },
    page: { font: { bold: true }, fill: { color: GREENS.soft } },
  }),
  preset('PivotStyleMedium28', N('Medium green, lined'), {
    header: { font: { bold: true, color: GREENS.deep }, border: { top: medium(GREENS.main), bottom: medium(GREENS.main) } },
    rowHeader: { font: { bold: true } },
    sub: { font: { bold: true }, border: { top: thin(GREENS.soft) } },
    grand: { font: { bold: true }, border: { top: medium(GREENS.main), bottom: medium(GREENS.main) } },
    band: { fill: { color: GREENS.pale } },
    page: { font: { bold: true, color: GREENS.deep } },
  }),
  preset('PivotStyleDark7', N('Dark green'), {
    header: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.ink } },
    rowHeader: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.deep } },
    sub: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.main } },
    grand: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.ink } },
    band: { fill: { color: GREENS.pale } },
    page: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.ink } },
  }),
  preset('PivotStyleDark14', N('Dark green, soft rows'), {
    header: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.deep } },
    rowHeader: { font: { bold: true }, fill: { color: GREENS.soft } },
    sub: { font: { bold: true }, fill: { color: GREENS.soft } },
    grand: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.deep } },
    band: { fill: { color: GREENS.faint } },
    page: { font: { bold: true, color: '#ffffff' }, fill: { color: GREENS.deep } },
  }),
];

export const NO_STYLE = 'None';

/** Preset for a style name; names from other files map to the nearest look by their family. */
export const presetFor = (name: string): PivotPreset | null => {
  if (!name || name === NO_STYLE) return null;
  const hit = PRESETS.find((p) => p.name === name);
  if (hit) return hit;
  if (/dark/i.test(name)) return PRESETS.find((p) => p.name === 'PivotStyleDark7')!;
  if (/medium/i.test(name)) return PRESETS.find((p) => p.name === 'PivotStyleMedium7')!;
  return PRESETS.find((p) => p.name === 'PivotStyleLight16')!;
};

const merge = (a: CellStyle, b: CellStyle): CellStyle => ({
  ...a, ...b,
  font: a.font || b.font ? { ...a.font, ...b.font } : undefined,
  fill: b.fill ?? a.fill,
  border: a.border || b.border ? { ...a.border, ...b.border } : undefined,
});

/** Style for one table cell. `bandRow`/`bandCol` say which band the cell is in. */
export const cellStyleFor = (info: PivotStyleInfo, cell: PCell, bandRow: boolean, bandCol: boolean, drill = true): CellStyle => {
  const p = presetFor(info.name);
  let s: CellStyle = {};
  if (cell.numFmt) s.numFmt = cell.numFmt;
  const indent = cell.role === 'rowItem' ? (cell.indent ?? 0) * (drill ? 3 : 2) + (cell.button ? 2 : 0) : 0;
  if (indent) s.align = { indent, h: 'left' };
  else if (cell.role === 'rowItem' && typeof cell.v === 'number') s.align = { h: 'left' };
  if (!p) {
    if (cell.total || cell.role === 'rowCaption' || cell.role === 'colCaption' || cell.role === 'pageLabel') s = merge(s, { font: { bold: true } });
    return clean(s);
  }
  const header = cell.role === 'rowCaption' || cell.role === 'colCaption' || cell.role === 'dataCaption' || (cell.role === 'colItem' && info.colHeaders);
  if (cell.role === 'pageLabel' || cell.role === 'pageValue') s = merge(s, p.page);
  else if (cell.total === 'grand') s = merge(s, p.grand);
  else if (header) s = merge(s, p.header);
  else if (cell.total === 'sub') s = merge(s, p.sub);
  else if (cell.role === 'rowItem' && info.rowHeaders) s = merge(s, p.rowHeader.fill ? p.rowHeader : { font: p.rowHeader.font && cell.button ? p.rowHeader.font : undefined });
  else if (cell.role === 'data' && ((info.bandedRows && bandRow) || (info.bandedCols && bandCol))) s = merge(s, p.band);
  if (cell.role === 'pageValue') s = merge(s, { font: { bold: false } });
  return clean(s);
};

const clean = (s: CellStyle): CellStyle => {
  const out: CellStyle = {};
  for (const [k, v] of Object.entries(s)) if (v !== undefined && !(typeof v === 'object' && v && !Object.keys(v).length)) (out as Record<string, unknown>)[k] = v;
  if (out.font) out.font = Object.fromEntries(Object.entries(out.font).filter(([, v]) => v !== undefined && v !== false)) as CellStyle['font'];
  if (out.font && !Object.keys(out.font).length) delete out.font;
  return out;
};

/** Small swatch description for the style picker. */
export const swatch = (p: PivotPreset | null): { head: string; headInk: string; band: string; total: string } => ({
  head: p?.header.fill?.color ?? 'transparent', headInk: p?.header.font?.color ?? '#142019', band: p?.band.fill?.color ?? 'transparent', total: p?.grand.fill?.color ?? p?.sub.fill?.color ?? 'transparent',
});
