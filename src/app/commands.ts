import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { cellKey, colName, mergeStyle, rangeAddr, type BorderSide, type CellRange, type CellStyle } from '../model/workbook';
import { colPx, pxToChars } from './layout';
import { copySelection, paste, displayText } from './ops/clipboard';
import { fillDirection, flashFill } from './ops/fill';
import { changeStructure, addSheet, deleteSheet, duplicateSheet } from './ops/structure';
import { dataRange, looksLikeHeader, sortRange, toggleFilter, clearFilters, applyFilter, groupLines, setLinesHidden, currentRegion } from './ops/data';
import type { GridStore } from './store';
import { canonical } from './store';

export interface Command {
  id: string;
  label: string;
  group?: string;
  keys?: string;
  run: (store: GridStore, arg?: unknown) => void;
  enabled?: (store: GridStore) => boolean;
  active?: (store: GridStore) => boolean;
}

const registry = new Map<string, Command>();
export const command = (id: string): Command | undefined => registry.get(id);
export const allCommands = (): Command[] => [...registry.values()];
export const run = (store: GridStore, id: string, arg?: unknown): void => {
  const c = registry.get(id);
  if (!c || (c.enabled && !c.enabled(store))) return;
  if (store.editing && !['edit.undo', 'edit.redo'].includes(id) && !id.startsWith('format.')) { if (!store.commitEdit()) return; }
  c.run(store, arg);
};
const add = (c: Command) => registry.set(c.id, c);
/** Feature modules add their own commands with this. */
export const registerCommand = (c: Command): void => { registry.set(c.id, c); };

const activeStyle = (s: GridStore): CellStyle => s.style(s.sel.active.row, s.sel.active.col);
const toggleFont = (key: 'bold' | 'italic' | 'strike', label: string, keys: string) => add({
  id: `format.${key}`, label, group: 'Font', keys,
  active: (s) => !!activeStyle(s).font?.[key],
  run: (s) => { const on = !activeStyle(s).font?.[key]; s.applyStyle((st) => ({ ...st, font: { ...st.font, [key]: on || undefined } }), label); },
});

const NUMBER_FORMATS: [string, string, string][] = [
  ['general', 'General', 'General'], ['number', 'Number', '0.00'], ['comma', 'Comma style', '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'],
  ['currency', 'Currency', '"$"#,##0.00'], ['accounting', 'Accounting', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
  ['shortDate', 'Short date', 'm/d/yyyy'], ['longDate', 'Long date', 'dddd, mmmm d, yyyy'], ['time', 'Time', 'h:mm:ss AM/PM'],
  ['percent', 'Percent style', '0%'], ['fraction', 'Fraction', '# ?/?'], ['scientific', 'Scientific', '0.00E+00'], ['text', 'Text', '@'],
];

export const registerCommands = (): void => {
  if (registry.size) return;
  // ---- Edit ----
  add({ id: 'edit.undo', label: 'Undo', keys: 'Mod+Z', enabled: (s) => s.canUndo, run: (s) => s.undo() });
  add({ id: 'edit.redo', label: 'Redo', keys: 'Mod+Y', enabled: (s) => s.canRedo, run: (s) => s.redo() });
  add({ id: 'edit.copy', label: 'Copy', keys: 'Mod+C', run: (s) => void copySelection(s) });
  add({ id: 'edit.cut', label: 'Cut', keys: 'Mod+X', run: (s) => void copySelection(s, true) });
  add({ id: 'edit.paste', label: 'Paste', keys: 'Mod+V', run: (s) => void paste(s) });
  add({ id: 'edit.pasteValues', label: 'Paste values', group: 'Paste', keys: 'Mod+Shift+V', run: (s) => void paste(s, { mode: 'values' }) });
  add({ id: 'edit.pasteFormats', label: 'Paste formatting', group: 'Paste', run: (s) => void paste(s, { mode: 'formats' }) });
  add({ id: 'edit.pasteFormulas', label: 'Paste formulas', group: 'Paste', run: (s) => void paste(s, { mode: 'formulas' }) });
  add({ id: 'edit.pasteTranspose', label: 'Paste transposed', group: 'Paste', run: (s) => void paste(s, { transpose: true }) });
  add({ id: 'edit.pasteLink', label: 'Paste link', group: 'Paste', run: (s) => void paste(s, { mode: 'link' }) });
  add({ id: 'edit.pasteSpecial', label: 'Paste special…', group: 'Paste', keys: 'Mod+Alt+V', run: (s) => openDialog(s, 'pasteSpecial') });
  add({ id: 'edit.selectAll', label: 'Select all', keys: 'Mod+A', run: (s) => {
    const r = s.range;
    const region = currentRegion(s);
    const isRegion = region.r1 === r.r1 && region.c1 === r.c1 && region.r2 === r.r2 && region.c2 === r.c2;
    if ((region.r1 === region.r2 && region.c1 === region.c2) || isRegion) s.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, s.sel.active);
    else s.selectRange(region, s.sel.active);
  } });
  add({ id: 'edit.selectColumn', label: 'Select column', keys: 'Mod+Space', run: (s) => { const r = s.range; s.selectRange({ r1: 0, c1: r.c1, r2: MAX_ROWS - 1, c2: r.c2 }, s.sel.active); } });
  add({ id: 'edit.selectRow', label: 'Select row', keys: 'Shift+Space', run: (s) => { const r = s.range; s.selectRange({ r1: r.r1, c1: 0, r2: r.r2, c2: MAX_COLS - 1 }, s.sel.active); } });
  add({ id: 'edit.clearContents', label: 'Clear contents', group: 'Clear', keys: 'Delete', run: (s) => clear(s, 'contents') });
  add({ id: 'edit.clearFormats', label: 'Clear formats', group: 'Clear', run: (s) => clear(s, 'formats') });
  add({ id: 'edit.clearAll', label: 'Clear all', group: 'Clear', run: (s) => clear(s, 'all') });
  add({ id: 'edit.clearNotes', label: 'Clear notes', group: 'Clear', run: (s) => clear(s, 'notes') });
  add({ id: 'edit.clearLinks', label: 'Clear hyperlinks', group: 'Clear', run: (s) => clear(s, 'links') });
  add({ id: 'edit.fillDown', label: 'Fill down', group: 'Fill', keys: 'Mod+D', run: (s) => fillDirection(s, 'down') });
  add({ id: 'edit.fillRight', label: 'Fill right', group: 'Fill', keys: 'Mod+R', run: (s) => fillDirection(s, 'right') });
  add({ id: 'edit.fillUp', label: 'Fill up', group: 'Fill', run: (s) => fillDirection(s, 'up') });
  add({ id: 'edit.fillLeft', label: 'Fill left', group: 'Fill', run: (s) => fillDirection(s, 'left') });
  add({ id: 'edit.series', label: 'Series…', group: 'Fill', run: (s) => openDialog(s, 'series') });
  add({ id: 'edit.flashFill', label: 'Flash Fill', group: 'Fill', keys: 'Mod+E', run: (s) => void flashFill(s) });
  add({ id: 'edit.find', label: 'Find…', keys: 'Mod+F', run: (s) => openDialog(s, 'find', { tab: 'find' }) });
  add({ id: 'edit.replace', label: 'Replace…', keys: 'Mod+H', run: (s) => openDialog(s, 'find', { tab: 'replace' }) });
  add({ id: 'edit.goto', label: 'Go to…', keys: 'Mod+G', run: (s) => openDialog(s, 'goto') });
  add({ id: 'edit.gotoSpecial', label: 'Go to special…', run: (s) => openDialog(s, 'gotoSpecial') });
  add({ id: 'edit.palette', label: 'Search commands', keys: 'Mod+K', run: (s) => openDialog(s, 'palette') });
  add({ id: 'edit.insertDate', label: "Insert today's date", keys: 'Mod+;', run: (s) => { const d = new Date(); s.beginEdit(`${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`); s.commitEdit(); } });
  add({ id: 'edit.insertTime', label: 'Insert the current time', keys: 'Mod+Shift+;', run: (s) => { const d = new Date(); const h = d.getHours() % 12 || 12; s.beginEdit(`${h}:${String(d.getMinutes()).padStart(2, '0')} ${d.getHours() < 12 ? 'AM' : 'PM'}`); s.commitEdit(); } });
  add({ id: 'edit.formatPainter', label: 'Format painter', active: () => !!painter.style, run: (s) => { painter.style = painter.style ? null : structuredClone(activeStyle(s)); s.emit(); } });

  // ---- Font & alignment ----
  toggleFont('bold', 'Bold', 'Mod+B');
  toggleFont('italic', 'Italic', 'Mod+I');
  toggleFont('strike', 'Strikethrough', 'Mod+5');
  add({ id: 'format.underline', label: 'Underline', group: 'Font', keys: 'Mod+U', active: (s) => !!activeStyle(s).font?.underline,
    run: (s) => { const u = activeStyle(s).font?.underline ? undefined : 'single' as const; s.applyStyle((st) => ({ ...st, font: { ...st.font, underline: u } }), 'Underline'); } });
  add({ id: 'format.doubleUnderline', label: 'Double underline', group: 'Font', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, underline: st.font?.underline === 'double' ? undefined : 'double' } }), 'Underline') });
  add({ id: 'format.superscript', label: 'Superscript', group: 'Font', active: (s) => activeStyle(s).font?.vertAlign === 'superscript', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, vertAlign: st.font?.vertAlign === 'superscript' ? undefined : 'superscript' } }), 'Superscript') });
  add({ id: 'format.subscript', label: 'Subscript', group: 'Font', active: (s) => activeStyle(s).font?.vertAlign === 'subscript', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, vertAlign: st.font?.vertAlign === 'subscript' ? undefined : 'subscript' } }), 'Subscript') });
  add({ id: 'format.fontName', label: 'Font', group: 'Font', run: (s, name) => s.applyStyle((st) => ({ ...st, font: { ...st.font, name: name as string } }), 'Font') });
  add({ id: 'format.fontSize', label: 'Font size', group: 'Font', run: (s, size) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: Math.max(1, Math.min(409, Number(size))) } }), 'Font size') });
  add({ id: 'format.growFont', label: 'Increase font size', group: 'Font', keys: 'Mod+Shift+.', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: stepSize(st.font?.size ?? 11, 1) } }), 'Font size') });
  add({ id: 'format.shrinkFont', label: 'Decrease font size', group: 'Font', keys: 'Mod+Shift+,', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: stepSize(st.font?.size ?? 11, -1) } }), 'Font size') });
  add({ id: 'format.fontColor', label: 'Font color', group: 'Font', run: (s, c) => s.applyStyle((st) => ({ ...st, font: { ...st.font, color: (c as string) || undefined } }), 'Font color') });
  add({ id: 'format.fillColor', label: 'Fill color', group: 'Font', run: (s, c) => s.applyStyle((st) => ({ ...st, fill: c ? { color: c as string } : undefined }), 'Fill color') });
  for (const h of ['left', 'center', 'right'] as const) add({ id: `format.align.${h}`, label: `Align ${h}`, group: 'Alignment', keys: h === 'center' ? 'Mod+Shift+E' : undefined, active: (s) => activeStyle(s).align?.h === h, run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, h: st.align?.h === h ? undefined : h } }), 'Align') });
  for (const v of ['top', 'center', 'bottom'] as const) add({ id: `format.valign.${v}`, label: v === 'center' ? 'Middle align' : `${v[0]!.toUpperCase()}${v.slice(1)} align`, group: 'Alignment', active: (s) => (activeStyle(s).align?.v ?? 'bottom') === v, run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, v: v === 'bottom' ? undefined : v } }), 'Align') });
  add({ id: 'format.wrap', label: 'Wrap text', group: 'Alignment', active: (s) => !!activeStyle(s).align?.wrap, run: (s) => { const on = !activeStyle(s).align?.wrap; s.applyStyle((st) => ({ ...st, align: { ...st.align, wrap: on || undefined } }), 'Wrap text'); } });
  add({ id: 'format.indent', label: 'Increase indent', group: 'Alignment', run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, indent: Math.min(250, (st.align?.indent ?? 0) + 1), h: st.align?.h === 'right' ? 'right' : st.align?.h === 'center' ? 'left' : st.align?.h ?? 'left' } }), 'Indent') });
  add({ id: 'format.outdent', label: 'Decrease indent', group: 'Alignment', run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, indent: Math.max(0, (st.align?.indent ?? 0) - 1) || undefined } }), 'Indent') });
  add({ id: 'format.rotate', label: 'Orientation', group: 'Alignment', run: (s, deg) => s.applyStyle((st) => ({ ...st, align: { ...st.align, rotation: (deg as number) || undefined } }), 'Orientation') });
  add({ id: 'format.mergeCenter', label: 'Merge & center', group: 'Alignment', active: (s) => !!s.mergeAt(s.sel.active.row, s.sel.active.col), run: (s) => mergeCells(s, 'center') });
  add({ id: 'format.mergeAcross', label: 'Merge across', group: 'Alignment', run: (s) => mergeCells(s, 'across') });
  add({ id: 'format.merge', label: 'Merge cells', group: 'Alignment', run: (s) => mergeCells(s, 'plain') });
  add({ id: 'format.unmerge', label: 'Unmerge cells', group: 'Alignment', run: (s) => mergeCells(s, 'unmerge') });

  // ---- Number ----
  for (const [id, label, code] of NUMBER_FORMATS) add({ id: `number.${id}`, label, group: 'Number', active: (s) => (activeStyle(s).numFmt ?? 'General') === code, run: (s) => s.applyStyle((st) => ({ ...st, numFmt: code === 'General' ? undefined : code }), 'Number format') });
  add({ id: 'number.format', label: 'Number format', group: 'Number', run: (s, code) => s.applyStyle((st) => ({ ...st, numFmt: (code as string) === 'General' ? undefined : (code as string) }), 'Number format') });
  add({ id: 'number.moreDecimals', label: 'Increase decimal', group: 'Number', run: (s) => changeDecimals(s, 1) });
  add({ id: 'number.lessDecimals', label: 'Decrease decimal', group: 'Number', run: (s) => changeDecimals(s, -1) });
  const keyFormats: [string, string, string][] = [['Mod+Shift+1', 'number.comma2', '#,##0.00'], ['Mod+Shift+4', 'number.currency2', '"$"#,##0.00'], ['Mod+Shift+5', 'number.percent0', '0%'], ['Mod+Shift+3', 'number.date2', 'd-mmm-yy'], ['Mod+Shift+2', 'number.time2', 'h:mm AM/PM'], ['Mod+Shift+6', 'number.sci2', '0.00E+00'], ['Mod+Shift+`', 'number.general2', 'General']];
  for (const [k, id, code] of keyFormats) add({ id, label: `Apply ${code}`, keys: k, run: (s) => s.applyStyle((st) => ({ ...st, numFmt: code === 'General' ? undefined : code }), 'Number format') });

  // ---- Borders ----
  const borderCmd = (id: string, label: string, f: (r: CellRange, row: number, col: number, side: BorderSide) => CellStyle['border'] | null) => add({
    id: `border.${id}`, label, group: 'Borders',
    run: (s, styleArg) => {
      const side: BorderSide = (styleArg as BorderSide) ?? { style: id === 'thickBox' || id === 'thickBottom' ? 'medium' : id === 'doubleBottom' ? 'double' : 'thin' };
      s.change(label, s.sel.ranges.map((area) => ({ area: { r1: Math.max(0, area.r1 - 1), c1: Math.max(0, area.c1 - 1), r2: area.r2 + 1, c2: area.c2 + 1 } })), () => {
        for (const r of s.sel.ranges) for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 5000); row++) for (let col = r.c1; col <= Math.min(r.c2, r.c1 + 500); col++) {
          const b = f(r, row, col, side);
          if (b === null) continue;
          s.styleRange({ r1: row, c1: col, r2: row, c2: col }, (st) => ({ ...st, border: id === 'none' ? undefined : { ...st.border, ...b } }));
        }
      });
    },
  });
  borderCmd('all', 'All borders', (_r, _row, _col, sd) => ({ top: sd, bottom: sd, left: sd, right: sd }));
  borderCmd('outside', 'Outside borders', (r, row, col, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }), ...(col === r.c1 && { left: sd }), ...(col === r.c2 && { right: sd }) }));
  borderCmd('thickBox', 'Thick outside borders', (r, row, col, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }), ...(col === r.c1 && { left: sd }), ...(col === r.c2 && { right: sd }) }));
  borderCmd('inside', 'Inside borders', (r, row, col, sd) => ({ ...(row > r.r1 && { top: sd }), ...(row < r.r2 && { bottom: sd }), ...(col > r.c1 && { left: sd }), ...(col < r.c2 && { right: sd }) }));
  borderCmd('top', 'Top border', (r, row, _c, sd) => (row === r.r1 ? { top: sd } : null));
  borderCmd('bottom', 'Bottom border', (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('left', 'Left border', (r, _row, col, sd) => (col === r.c1 ? { left: sd } : null));
  borderCmd('right', 'Right border', (r, _row, col, sd) => (col === r.c2 ? { right: sd } : null));
  borderCmd('thickBottom', 'Thick bottom border', (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('doubleBottom', 'Bottom double border', (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('topBottom', 'Top and bottom border', (r, row, _c, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }) }));
  borderCmd('none', 'No border', () => ({}));

  // ---- Cells, rows, columns ----
  add({ id: 'cells.insertRows', label: 'Insert sheet rows', group: 'Cells', keys: 'Mod+Shift+=', run: (s) => { const r = s.range; changeStructure(s, 'row', r.r1, r.r2 - r.r1 + 1); } });
  add({ id: 'cells.insertCols', label: 'Insert sheet columns', group: 'Cells', run: (s) => { const r = s.range; changeStructure(s, 'col', r.c1, r.c2 - r.c1 + 1); } });
  add({ id: 'cells.deleteRows', label: 'Delete sheet rows', group: 'Cells', keys: 'Mod+-', run: (s) => { const r = s.range; changeStructure(s, 'row', r.r1, -(r.r2 - r.r1 + 1)); } });
  add({ id: 'cells.deleteCols', label: 'Delete sheet columns', group: 'Cells', run: (s) => { const r = s.range; changeStructure(s, 'col', r.c1, -(r.c2 - r.c1 + 1)); } });
  add({ id: 'cells.insertDialog', label: 'Insert cells…', group: 'Cells', run: (s) => openDialog(s, 'insertCells', { insert: true }) });
  add({ id: 'cells.deleteDialog', label: 'Delete cells…', group: 'Cells', run: (s) => openDialog(s, 'insertCells', { insert: false }) });
  add({ id: 'cells.hideRows', label: 'Hide rows', group: 'Cells', keys: 'Mod+9', run: (s) => setLinesHidden(s, 'row', s.range.r1, s.range.r2, true) });
  add({ id: 'cells.unhideRows', label: 'Unhide rows', group: 'Cells', keys: 'Mod+Shift+9', run: (s) => setLinesHidden(s, 'row', Math.max(0, s.range.r1 - (s.range.r1 === s.range.r2 ? 0 : 0)), s.range.r2, false) });
  add({ id: 'cells.hideCols', label: 'Hide columns', group: 'Cells', keys: 'Mod+0', run: (s) => setLinesHidden(s, 'col', s.range.c1, s.range.c2, true) });
  add({ id: 'cells.unhideCols', label: 'Unhide columns', group: 'Cells', keys: 'Mod+Shift+0', run: (s) => setLinesHidden(s, 'col', s.range.c1, s.range.c2, false) });
  add({ id: 'cells.autofitCols', label: 'AutoFit column width', group: 'Cells', run: (s) => autofit(s, 'col') });
  add({ id: 'cells.autofitRows', label: 'AutoFit row height', group: 'Cells', run: (s) => autofit(s, 'row') });
  add({ id: 'cells.colWidth', label: 'Column width…', group: 'Cells', run: (s) => openDialog(s, 'size', { axis: 'col' }) });
  add({ id: 'cells.rowHeight', label: 'Row height…', group: 'Cells', run: (s) => openDialog(s, 'size', { axis: 'row' }) });
  add({ id: 'cells.formatDialog', label: 'Format cells…', group: 'Cells', keys: 'Mod+1', run: (s) => openDialog(s, 'formatCells') });
  add({ id: 'cells.styles', label: 'Cell styles', group: 'Cells', run: (s) => openDialog(s, 'cellStyles') });
  add({ id: 'cells.formatTable', label: 'Format as table', group: 'Cells', keys: 'Mod+T', run: (s) => openDialog(s, 'table') });
  add({ id: 'cells.group', label: 'Group', group: 'Outline', keys: 'Alt+Shift+ArrowRight', run: (s) => { const r = s.range; if (r.c1 === 0 && r.c2 === MAX_COLS - 1 || r.c1 === r.c2) groupLines(s, 'row', r.r1, r.r2, 1); else groupLines(s, 'col', r.c1, r.c2, 1); } });
  add({ id: 'cells.ungroup', label: 'Ungroup', group: 'Outline', keys: 'Alt+Shift+ArrowLeft', run: (s) => { const r = s.range; if (r.c1 === 0 && r.c2 === MAX_COLS - 1 || r.c1 === r.c2) groupLines(s, 'row', r.r1, r.r2, -1); else groupLines(s, 'col', r.c1, r.c2, -1); } });

  // ---- Sheets ----
  add({ id: 'sheet.insert', label: 'Insert sheet', group: 'Sheet', keys: 'Shift+F11', run: (s) => addSheet(s) });
  add({ id: 'sheet.delete', label: 'Delete sheet', group: 'Sheet', run: (s) => deleteSheet(s) });
  add({ id: 'sheet.duplicate', label: 'Duplicate sheet', group: 'Sheet', run: (s) => duplicateSheet(s) });
  add({ id: 'sheet.rename', label: 'Rename sheet', group: 'Sheet', run: (s) => openDialog(s, 'renameSheet', { index: s.sheetIndex }) });
  add({ id: 'sheet.next', label: 'Next sheet', keys: 'Mod+PageDown', run: (s) => { let i = s.sheetIndex + 1; while (s.wb.sheets[i] && s.wb.sheets[i]!.state !== 'visible') i++; if (s.wb.sheets[i]) s.setActiveSheet(i); } });
  add({ id: 'sheet.prev', label: 'Previous sheet', keys: 'Mod+PageUp', run: (s) => { let i = s.sheetIndex - 1; while (i >= 0 && s.wb.sheets[i]!.state !== 'visible') i--; if (i >= 0) s.setActiveSheet(i); } });
  add({ id: 'sheet.protect', label: 'Protect sheet', group: 'Review', active: (s) => !!s.sheet.protection?.sheet, run: (s) => openDialog(s, 'protect') });

  // ---- Formulas ----
  for (const fn of ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN']) add({ id: `formula.auto${fn}`, label: fn === 'SUM' ? 'AutoSum' : fn === 'COUNT' ? 'Count numbers' : fn[0] + fn.slice(1).toLowerCase(), group: 'Formulas', keys: fn === 'SUM' ? 'Alt+=' : undefined, run: (s) => autoSum(s, fn) });
  add({ id: 'formula.insertFunction', label: 'Insert function…', group: 'Formulas', keys: 'Shift+F3', run: (s) => openDialog(s, 'insertFunction') });
  add({ id: 'formula.names', label: 'Name manager', group: 'Formulas', keys: 'Mod+F3', run: (s) => openDialog(s, 'names') });
  add({ id: 'formula.defineName', label: 'Define name…', group: 'Formulas', run: (s) => openDialog(s, 'defineName') });
  add({ id: 'formula.createFromSelection', label: 'Create names from selection', group: 'Formulas', keys: 'Mod+Shift+F3', run: (s) => createNamesFromSelection(s) });
  add({ id: 'formula.showFormulas', label: 'Show formulas', group: 'Formulas', keys: 'Mod+`', active: (s) => s.view.showFormulas, run: (s) => { s.view.showFormulas = !s.view.showFormulas; s.touch(); } });
  add({ id: 'formula.calcNow', label: 'Calculate now', group: 'Formulas', keys: 'F9', run: (s) => { s.engine.recalc(true); s.touch(); } });
  add({ id: 'formula.iterative', label: 'Iterative calculation', group: 'Formulas', active: (s) => s.engine.iterate, run: (s) => { s.engine.iterate = !s.engine.iterate; s.wb.calc = { ...s.wb.calc, iterate: s.engine.iterate }; s.engine.recalc(true); s.touch(); } });
  add({ id: 'formula.evaluate', label: 'Evaluate formula', group: 'Formulas', run: (s) => openDialog(s, 'evaluate') });
  add({ id: 'formula.tracePrecedents', label: 'Trace precedents', group: 'Formulas', run: (s) => openDialog(s, 'trace', { dir: 'precedents' }) });
  add({ id: 'formula.traceDependents', label: 'Trace dependents', group: 'Formulas', run: (s) => openDialog(s, 'trace', { dir: 'dependents' }) });
  add({ id: 'formula.errorCheck', label: 'Error checking', group: 'Formulas', run: (s) => openDialog(s, 'errors') });
  add({ id: 'formula.arrayEnter', label: 'Enter as array formula', keys: 'Mod+Shift+Enter', run: (s) => { if (s.editing) s.commitEdit(false, true); } });

  // ---- Data ----
  add({ id: 'data.sortAsc', label: 'Sort A to Z', group: 'Data', run: (s) => quickSort(s, false) });
  add({ id: 'data.sortDesc', label: 'Sort Z to A', group: 'Data', run: (s) => quickSort(s, true) });
  add({ id: 'data.sort', label: 'Custom sort…', group: 'Data', run: (s) => openDialog(s, 'sort') });
  add({ id: 'data.filter', label: 'Filter', group: 'Data', keys: 'Mod+Shift+L', active: (s) => !!s.sheet.autoFilter, run: (s) => toggleFilter(s) });
  add({ id: 'data.clearFilter', label: 'Clear filters', group: 'Data', enabled: (s) => !!s.sheet.autoFilter?.columns.length, run: (s) => clearFilters(s) });
  add({ id: 'data.reapply', label: 'Reapply filter', group: 'Data', enabled: (s) => !!s.sheet.autoFilter, run: (s) => s.change('Reapply', 'sheet', () => applyFilter(s)) });
  add({ id: 'data.removeDuplicates', label: 'Remove duplicates…', group: 'Data', run: (s) => openDialog(s, 'removeDuplicates') });
  add({ id: 'data.textToColumns', label: 'Text to columns…', group: 'Data', run: (s) => openDialog(s, 'textToColumns') });
  add({ id: 'data.validation', label: 'Data validation…', group: 'Data', run: (s) => openDialog(s, 'validation') });
  add({ id: 'data.goalSeek', label: 'Goal Seek…', group: 'Data', run: (s) => openDialog(s, 'goalSeek') });
  add({ id: 'data.pivot', label: 'PivotTable…', group: 'Data', run: (s) => openDialog(s, 'pivot') });
  add({ id: 'data.subtotal', label: 'Subtotal…', group: 'Data', run: (s) => openDialog(s, 'subtotals') });
  add({ id: 'data.whatIfTable', label: 'Data table…', group: 'Data', run: (s) => openDialog(s, 'dataTable') });
  add({ id: 'data.importCsv', label: 'From text/CSV…', group: 'Data', run: (s) => openDialog(s, 'importCsv') });
  add({ id: 'data.stats', label: 'Descriptive statistics', group: 'Data', run: (s) => openDialog(s, 'stats') });

  // ---- Insert ----
  add({ id: 'insert.chart', label: 'Chart…', group: 'Insert', keys: 'Alt+F1', run: (s) => openDialog(s, 'chart') });
  add({ id: 'insert.picture', label: 'Picture…', group: 'Insert', run: (s) => openDialog(s, 'picture') });
  add({ id: 'insert.link', label: 'Link…', group: 'Insert', keys: 'Mod+K', run: (s) => openDialog(s, 'link') });
  add({ id: 'insert.note', label: 'Note', group: 'Insert', keys: 'Shift+F2', run: (s) => openDialog(s, 'note') });
  add({ id: 'insert.symbol', label: 'Symbol…', group: 'Insert', run: (s) => openDialog(s, 'symbol') });
  add({ id: 'insert.conditional', label: 'Conditional formatting…', group: 'Styles', run: (s) => openDialog(s, 'conditional') });
  add({ id: 'insert.clearRules', label: 'Clear rules from selection', group: 'Styles', run: (s) => clearRules(s, false) });
  add({ id: 'insert.clearAllRules', label: 'Clear rules from sheet', group: 'Styles', run: (s) => clearRules(s, true) });

  // ---- View ----
  add({ id: 'view.freezePanes', label: 'Freeze panes', group: 'View', active: (s) => !!(s.sheet.view.frozenRows || s.sheet.view.frozenCols), run: (s) => freeze(s, 'selection') });
  add({ id: 'view.freezeTopRow', label: 'Freeze top row', group: 'View', run: (s) => freeze(s, 'row') });
  add({ id: 'view.freezeFirstCol', label: 'Freeze first column', group: 'View', run: (s) => freeze(s, 'col') });
  add({ id: 'view.unfreeze', label: 'Unfreeze panes', group: 'View', run: (s) => freeze(s, 'none') });
  add({ id: 'view.gridlines', label: 'Gridlines', group: 'View', active: (s) => s.sheet.view.showGrid, run: (s) => { s.sheet.view.showGrid = !s.sheet.view.showGrid; s.file.dirty = true; s.touch(); } });
  add({ id: 'view.headings', label: 'Headings', group: 'View', active: (s) => s.sheet.view.showHeaders, run: (s) => { s.sheet.view.showHeaders = !s.sheet.view.showHeaders; s.touch(); } });
  add({ id: 'view.zoomIn', label: 'Zoom in', group: 'View', keys: 'Mod+=', run: (s) => setZoom(s, s.sheet.view.zoom + 10) });
  add({ id: 'view.zoomOut', label: 'Zoom out', group: 'View', keys: 'Mod+Shift+-', run: (s) => setZoom(s, s.sheet.view.zoom - 10) });
  add({ id: 'view.zoom100', label: 'Zoom to 100%', group: 'View', run: (s) => setZoom(s, 100) });
  add({ id: 'view.zoomSelection', label: 'Zoom to selection', group: 'View', run: (s) => openDialog(s, 'zoomToSelection') });
  add({ id: 'view.theme', label: 'Dark mode', group: 'View', active: (s) => s.view.theme === 'dark' || (s.view.theme === 'system' && matchMedia('(prefers-color-scheme: dark)').matches), run: (s) => { const dark = s.view.theme === 'dark' || (s.view.theme === 'system' && matchMedia('(prefers-color-scheme: dark)').matches); s.view.theme = dark ? 'light' : 'dark'; try { localStorage.setItem('clerdogrid.theme', s.view.theme); } catch { /* private mode */ } s.touch(); } });
  add({ id: 'view.rtl', label: 'Sheet right to left', group: 'View', active: (s) => s.sheet.view.rtl, run: (s) => { s.sheet.view.rtl = !s.sheet.view.rtl; s.touch(); } });

  // ---- File ----
  add({ id: 'file.new', label: 'New workbook', group: 'File', keys: 'Mod+N', run: (s) => void import('./files').then((m) => m.newWorkbookFile(s)) });
  add({ id: 'file.open', label: 'Open…', group: 'File', keys: 'Mod+O', run: (s) => void import('./files').then((m) => m.openFile(s)) });
  add({ id: 'file.save', label: 'Save', group: 'File', keys: 'Mod+S', run: (s) => void import('./files').then((m) => m.saveFile(s)) });
  add({ id: 'file.saveAs', label: 'Save as…', group: 'File', keys: 'Mod+Shift+S', run: (s) => void import('./files').then((m) => m.saveFile(s, true)) });
  add({ id: 'file.exportCsv', label: 'Export sheet as CSV', group: 'File', run: (s) => void import('./files').then((m) => m.exportCsv(s)) });
  add({ id: 'file.exportPdf', label: 'Export as PDF / print', group: 'File', keys: 'Mod+P', run: (s) => void import('./files').then((m) => m.printSheet(s)) });
  add({ id: 'file.templates', label: 'Templates…', group: 'File', run: (s) => openDialog(s, 'templates') });
  add({ id: 'file.properties', label: 'Workbook properties', group: 'File', run: (s) => openDialog(s, 'properties') });
  add({ id: 'help.shortcuts', label: 'Keyboard shortcuts', group: 'Help', keys: 'Mod+/', run: (s) => openDialog(s, 'shortcuts') });
  add({ id: 'help.about', label: 'About ClerdoGrid', group: 'Help', run: (s) => openDialog(s, 'about') });
};

export const painter: { style: CellStyle | null } = { style: null };

export const openDialog = (s: GridStore, id: string, data?: unknown): void => { s.dialog = { id, data }; s.emit(); };

const stepSize = (size: number, dir: 1 | -1): number => {
  const steps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  if (dir > 0) return steps.find((x) => x > size) ?? Math.min(409, size + 10);
  return [...steps].reverse().find((x) => x < size) ?? Math.max(1, size - 1);
};

const clear = (s: GridStore, what: 'contents' | 'formats' | 'all' | 'notes' | 'links'): void => {
  const whole = s.sel.ranges.some((r) => (r.r1 === 0 && r.r2 === MAX_ROWS - 1) || (r.c1 === 0 && r.c2 === MAX_COLS - 1));
  const b = s.engine.usedBounds(s.sheetIndex);
  const areas = s.sel.ranges.map((r) => ({ r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, b.rows)), c2: Math.min(r.c2, Math.max(r.c1, b.cols)) }));
  if (areas.some((r) => { for (let row = r.r1; row <= r.r2 && row - r.r1 < 2000; row++) for (let col = r.c1; col <= r.c2 && col - r.c1 < 200; col++) if (s.isLocked(row, col)) return true; return false; })) {
    s.notify('Some of these cells are protected. Unprotect the sheet to change them.');
    return;
  }
  s.change(what === 'contents' ? 'Clear' : `Clear ${what}`, whole ? 'sheet' : areas.map((area) => ({ area })), () => {
    for (const r of areas) {
      for (const [k, c] of [...s.sheet.cells]) {
        const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
        if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
        if (what === 'contents' || what === 'all') { if (c.f || c.v !== null) s.engine.setValue(s.sheetIndex, row, col, null); }
        if (what === 'formats' || what === 'all') { const cc = s.sheet.cells.get(k); if (cc) { delete cc.s; if (cc.v === null && !cc.f) s.sheet.cells.delete(k); } }
        if (what === 'notes' || what === 'all') s.sheet.notes.delete(k);
        if (what === 'links' || what === 'all') s.sheet.hyperlinks.delete(k);
      }
      if (what === 'notes' || what === 'links') for (const k of [...(what === 'notes' ? s.sheet.notes : s.sheet.hyperlinks).keys()]) {
        const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
        if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2) (what === 'notes' ? s.sheet.notes : s.sheet.hyperlinks).delete(k);
      }
    }
  });
};

const mergeCells = (s: GridStore, how: 'center' | 'across' | 'plain' | 'unmerge'): void => {
  const r = s.range;
  const existing = s.sheet.merges.filter((m) => m.r1 <= r.r2 && m.r2 >= r.r1 && m.c1 <= r.c2 && m.c2 >= r.c1);
  if (how === 'unmerge' || (how === 'center' && existing.length && existing.every((m) => m.r1 >= r.r1 && m.r2 <= r.r2 && m.c1 >= r.c1 && m.c2 <= r.c2) && existing.some((m) => m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2))) {
    s.change('Unmerge', 'sheet', () => { s.sheet.merges = s.sheet.merges.filter((m) => !existing.includes(m)); });
    return;
  }
  if (r.r1 === r.r2 && r.c1 === r.c2) return;
  const rows = how === 'across' ? Array.from({ length: r.r2 - r.r1 + 1 }, (_, i) => ({ ...r, r1: r.r1 + i, r2: r.r1 + i })) : [r];
  let lost = 0;
  for (const m of rows) for (let row = m.r1; row <= m.r2; row++) for (let col = m.c1; col <= m.c2; col++) if ((row !== m.r1 || col !== m.c1) && s.value(row, col) !== null) lost++;
  if (lost && !confirm('Merging cells only keeps the upper-left value and discards the other values.')) return;
  s.change('Merge', 'sheet', () => {
    s.sheet.merges = s.sheet.merges.filter((m) => !existing.includes(m));
    for (const m of rows) {
      for (let row = m.r1; row <= m.r2; row++) for (let col = m.c1; col <= m.c2; col++) if (row !== m.r1 || col !== m.c1) s.engine.setValue(s.sheetIndex, row, col, null);
      s.sheet.merges.push(m);
      if (how === 'center') s.styleRange({ r1: m.r1, c1: m.c1, r2: m.r1, c2: m.c1 }, (st) => mergeStyle(st, { align: { ...st.align, h: 'center' } }));
    }
  });
  s.selectRange(r);
};

const changeDecimals = (s: GridStore, d: 1 | -1): void => {
  const { row, col } = s.sel.active;
  const v = s.value(row, col);
  s.applyStyle((st) => {
    let fmt = st.numFmt ?? 'General';
    if (fmt === 'General') {
      const shown = typeof v === 'number' ? String(+v.toPrecision(10)) : '0';
      const dec = (shown.split('.')[1] ?? '').length;
      fmt = dec + d > 0 ? `0.${'0'.repeat(dec + d)}` : '0';
      return { ...st, numFmt: fmt };
    }
    const sections = fmt.split(';').map((sec) => {
      if (d > 0) {
        if (/0\.0*/.test(sec)) return sec.replace(/(0\.0*)/, '$10');
        return sec.replace(/(^|[^0#?.])((?:#,##)?0)(?![0.])/, '$1$2.0');
      }
      return sec.replace(/0\.(0*)0/, (_m, z: string) => (z.length ? `0.${z}` : '0'));
    });
    return { ...st, numFmt: sections.join(';') };
  }, d > 0 ? 'Increase decimal' : 'Decrease decimal');
};

const autofit = (s: GridStore, axis: 'col' | 'row'): void => {
  const r = s.range;
  const canvas = document.createElement('canvas').getContext('2d')!;
  s.change('AutoFit', 'sheet', () => {
    const b = s.engine.usedBounds(s.sheetIndex);
    if (axis === 'col') {
      for (let c = r.c1; c <= Math.min(r.c2, b.cols); c++) {
        let max = 0;
        for (let row = 0; row <= b.rows; row++) {
          if (s.value(row, c) === null || s.mergeAt(row, c)) continue;
          const st = s.style(row, c);
          canvas.font = `${st.font?.bold ? '700 ' : ''}${((st.font?.size ?? 11) * 4) / 3}px Carlito, Calibri, sans-serif`;
          max = Math.max(max, canvas.measureText(displayText(s, row, c)).width);
        }
        if (max > 0) { const info = s.sheet.cols.get(c) ?? {}; info.width = Math.min(255, pxToChars(Math.ceil(max) + 9)); s.sheet.cols.set(c, info); }
      }
    } else {
      for (let row = r.r1; row <= Math.min(r.r2, b.rows); row++) {
        let maxPt = 0;
        for (let c = 0; c <= b.cols; c++) {
          const v = s.value(row, c);
          const st = s.style(row, c);
          const size = st.font?.size ?? 11;
          if (v === null) { maxPt = Math.max(maxPt, size * 1.36); continue; }
          let lines = 1;
          if (st.align?.wrap) {
            const w = colPx(s.sheet.cols.get(c)?.width ?? s.sheet.defaultColWidth) - 6;
            canvas.font = `${(size * 4) / 3}px Carlito, Calibri, sans-serif`;
            const text = displayText(s, row, c);
            lines = text.split('\n').reduce((n, para) => n + Math.max(1, Math.ceil(canvas.measureText(para).width / Math.max(1, w))), 0);
          } else lines = String(v).split('\n').length > 1 && st.align?.wrap ? String(v).split('\n').length : 1;
          maxPt = Math.max(maxPt, size * 1.36 * lines);
        }
        const info = s.sheet.rows.get(row) ?? {};
        if (maxPt) info.height = Math.min(409, Math.round(maxPt * 4) / 4); else delete info.height;
        s.sheet.rows.set(row, info);
      }
    }
  });
};

const autoSum = (s: GridStore, fn: string): void => {
  const { row, col } = s.sel.active;
  const r = s.range;
  const isNum = (rr: number, cc: number) => typeof s.value(rr, cc) === 'number';
  if (r.r1 !== r.r2 || r.c1 !== r.c2) {
    // A block: put totals under each column (and to the right of each row when there's room).
    s.change(fn, [{ area: { ...r, r2: r.r2 + 1 } }], () => {
      for (let c = r.c1; c <= r.c2; c++) s.enter(r.r2 + 1, c, `=${fn}(${colName(c)}${r.r1 + 1}:${colName(c)}${r.r2 + 1})`);
    });
    return;
  }
  let top = row - 1;
  while (top >= 0 && (isNum(top, col) || (top === row - 1 && s.value(top, col) === null))) top--;
  let range: string;
  if (top < row - 1 && isNum(row - 1, col)) range = `${colName(col)}${top + 2}:${colName(col)}${row}`;
  else {
    let left = col - 1;
    while (left >= 0 && isNum(row, left)) left--;
    range = left < col - 1 ? `${colName(left + 1)}${row + 1}:${colName(col - 1)}${row + 1}` : '';
  }
  s.beginEdit(`=${fn}(${range})`);
};

const quickSort = (s: GridStore, desc: boolean): void => {
  const r = dataRange(s);
  sortRange(s, r, [{ col: s.sel.active.col, descending: desc }], looksLikeHeader(s, r));
};

const freeze = (s: GridStore, how: 'selection' | 'row' | 'col' | 'none'): void => {
  const v = s.sheet.view;
  if (how === 'selection' && (v.frozenRows || v.frozenCols)) how = 'none';
  const { row, col } = s.sel.active;
  if (how === 'none') { v.frozenRows = 0; v.frozenCols = 0; }
  else if (how === 'row') { v.frozenRows = 1; v.frozenCols = 0; }
  else if (how === 'col') { v.frozenRows = 0; v.frozenCols = 1; }
  else { v.frozenRows = row; v.frozenCols = col; }
  s.file.dirty = true;
  s.touch();
};

export const setZoom = (s: GridStore, z: number): void => {
  s.sheet.view.zoom = Math.max(10, Math.min(400, Math.round(z)));
  s.touch();
};

const createNamesFromSelection = (s: GridStore): void => {
  const r = s.range;
  if (r.r1 === r.r2) return;
  s.change('Create names', 'book', () => {
    for (let c = r.c1; c <= r.c2; c++) {
      const label = s.value(r.r1, c);
      if (typeof label !== 'string' || !label.trim()) continue;
      const name = label.trim().replace(/[^A-Za-z0-9_.]/g, '_').replace(/^(\d)/, '_$1');
      if (s.wb.names.some((n) => n.name.toLowerCase() === name.toLowerCase())) continue;
      s.wb.names.push({ name, formula: `'${s.sheet.name.replace(/'/g, "''")}'!$${colName(c)}$${r.r1 + 2}:$${colName(c)}$${r.r2 + 1}` });
    }
    s.engine.sheetsChanged();
    s.engine.recalc(true);
  });
  s.notify('Names created from the top row.');
};

const clearRules = (s: GridStore, all: boolean): void => {
  s.change('Clear rules', 'sheet', () => {
    if (all) { s.sheet.conditionalFormats = []; return; }
    const r = s.range;
    s.sheet.conditionalFormats = s.sheet.conditionalFormats.filter((cf) => !cf.ranges.some((x) => x.r1 <= r.r2 && x.r2 >= r.r1 && x.c1 <= r.c2 && x.c2 >= r.c1));
  });
};

/** Applies the format painter's style to the new selection. */
export const applyPainter = (s: GridStore): void => {
  if (!painter.style) return;
  const st = painter.style;
  painter.style = null;
  s.applyStyle(() => st, 'Format painter');
};

export { canonical, rangeAddr, cellKey };
