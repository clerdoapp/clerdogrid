import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { cellKey, colName, mergeStyle, rangeAddr, type BorderSide, type CellRange, type CellStyle } from '../model/workbook';
import { colPx, pxToChars } from './layout';
import { copySelection, paste, displayText } from './ops/clipboard';
import { fillDirection, flashFill } from './ops/fill';
import { changeStructure, addSheet, deleteSheet, duplicateSheet } from './ops/structure';
import { dataRange, looksLikeHeader, sortRange, toggleFilter, clearFilters, applyFilter, groupLines, setLinesHidden, currentRegion } from './ops/data';
import type { GridStore } from './store';
import { canonical } from './store';
import { regionalFormat } from './locale';
import { N, T } from '../ui/i18n';

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
  id: `format.${key}`, label, group: N('Font'), keys,
  active: (s) => !!activeStyle(s).font?.[key],
  run: (s) => { const on = !activeStyle(s).font?.[key]; s.applyStyle((st) => ({ ...st, font: { ...st.font, [key]: on || undefined } }), label); },
});

const NUMBER_FORMATS: [string, string, string][] = [
  ['general', N('General'), N('General')], ['number', N('Number'), '0.00'], ['comma', N('Comma style'), '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'],
  ['currency', N('Currency'), '"$"#,##0.00'], ['accounting', N('Accounting'), '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
  ['shortDate', N('Short date'), 'm/d/yyyy'], ['longDate', N('Long date'), 'dddd, mmmm d, yyyy'], ['time', N('Time'), 'h:mm:ss AM/PM'],
  ['percent', N('Percent style'), '0%'], ['fraction', N('Fraction'), '# ?/?'], ['scientific', N('Scientific'), '0.00E+00'], ['text', N('Text'), '@'],
];

export const registerCommands = (): void => {
  if (registry.size) return;
  // ---- Edit ----
  add({ id: 'edit.undo', label: N('Undo'), keys: 'Mod+Z', enabled: (s) => s.canUndo, run: (s) => s.undo() });
  add({ id: 'edit.redo', label: N('Redo'), keys: 'Mod+Y', enabled: (s) => s.canRedo, run: (s) => s.redo() });
  add({ id: 'edit.copy', label: N('Copy'), keys: 'Mod+C', run: (s) => void copySelection(s) });
  add({ id: 'edit.cut', label: N('Cut'), keys: 'Mod+X', run: (s) => void copySelection(s, true) });
  add({ id: 'edit.paste', label: N('Paste'), keys: 'Mod+V', run: (s) => void paste(s) });
  add({ id: 'edit.pasteValues', label: N('Paste values'), group: N('Paste'), keys: 'Mod+Shift+V', run: (s) => void paste(s, { mode: 'values' }) });
  add({ id: 'edit.pasteFormats', label: N('Paste formatting'), group: N('Paste'), run: (s) => void paste(s, { mode: 'formats' }) });
  add({ id: 'edit.pasteFormulas', label: N('Paste formulas'), group: N('Paste'), run: (s) => void paste(s, { mode: 'formulas' }) });
  add({ id: 'edit.pasteTranspose', label: N('Paste transposed'), group: N('Paste'), run: (s) => void paste(s, { transpose: true }) });
  add({ id: 'edit.pasteLink', label: N('Paste link'), group: N('Paste'), run: (s) => void paste(s, { mode: 'link' }) });
  add({ id: 'edit.pasteSpecial', label: N('Paste special…'), group: N('Paste'), keys: 'Mod+Alt+V', run: (s) => openDialog(s, 'pasteSpecial') });
  add({ id: 'edit.selectAll', label: N('Select all'), keys: 'Mod+A', run: (s) => {
    const r = s.range;
    const region = currentRegion(s);
    const isRegion = region.r1 === r.r1 && region.c1 === r.c1 && region.r2 === r.r2 && region.c2 === r.c2;
    if ((region.r1 === region.r2 && region.c1 === region.c2) || isRegion) s.selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, s.sel.active);
    else s.selectRange(region, s.sel.active);
  } });
  add({ id: 'edit.selectColumn', label: N('Select column'), keys: 'Mod+Space', run: (s) => { const r = s.range; s.selectRange({ r1: 0, c1: r.c1, r2: MAX_ROWS - 1, c2: r.c2 }, s.sel.active); } });
  add({ id: 'edit.selectRow', label: N('Select row'), keys: 'Shift+Space', run: (s) => { const r = s.range; s.selectRange({ r1: r.r1, c1: 0, r2: r.r2, c2: MAX_COLS - 1 }, s.sel.active); } });
  add({ id: 'edit.clearContents', label: N('Clear contents'), group: N('Clear'), keys: 'Delete', run: (s) => clear(s, 'contents') });
  add({ id: 'edit.clearFormats', label: N('Clear formats'), group: N('Clear'), run: (s) => clear(s, 'formats') });
  add({ id: 'edit.clearAll', label: N('Clear all'), group: N('Clear'), run: (s) => clear(s, 'all') });
  add({ id: 'edit.clearNotes', label: N('Clear notes'), group: N('Clear'), run: (s) => clear(s, 'notes') });
  add({ id: 'edit.clearLinks', label: N('Clear hyperlinks'), group: N('Clear'), run: (s) => clear(s, 'links') });
  add({ id: 'edit.fillDown', label: N('Fill down'), group: N('Fill'), keys: 'Mod+D', run: (s) => fillDirection(s, 'down') });
  add({ id: 'edit.fillRight', label: N('Fill right'), group: N('Fill'), keys: 'Mod+R', run: (s) => fillDirection(s, 'right') });
  add({ id: 'edit.fillUp', label: N('Fill up'), group: N('Fill'), run: (s) => fillDirection(s, 'up') });
  add({ id: 'edit.fillLeft', label: N('Fill left'), group: N('Fill'), run: (s) => fillDirection(s, 'left') });
  add({ id: 'edit.series', label: N('Series…'), group: N('Fill'), run: (s) => openDialog(s, 'series') });
  add({ id: 'edit.flashFill', label: N('Flash Fill'), group: N('Fill'), keys: 'Mod+E', run: (s) => void flashFill(s) });
  add({ id: 'edit.find', label: N('Find…'), keys: 'Mod+F', run: (s) => openDialog(s, 'find', { tab: 'find' }) });
  add({ id: 'edit.replace', label: N('Replace…'), keys: 'Mod+H', run: (s) => openDialog(s, 'find', { tab: 'replace' }) });
  add({ id: 'edit.goto', label: N('Go to…'), keys: 'Mod+G', run: (s) => openDialog(s, 'goto') });
  add({ id: 'edit.gotoSpecial', label: N('Go to special…'), run: (s) => openDialog(s, 'gotoSpecial') });
  add({ id: 'edit.palette', label: N('Search commands'), keys: 'Mod+K', run: (s) => openDialog(s, 'palette') });
  add({ id: 'edit.insertDate', label: N('Insert today\'s date'), keys: 'Mod+;', run: (s) => { const d = new Date(); s.beginEdit(`${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`); s.commitEdit(); } });
  add({ id: 'edit.insertTime', label: N('Insert the current time'), keys: 'Mod+Shift+;', run: (s) => { const d = new Date(); const h = d.getHours() % 12 || 12; s.beginEdit(`${h}:${String(d.getMinutes()).padStart(2, '0')} ${d.getHours() < 12 ? 'AM' : 'PM'}`); s.commitEdit(); } });
  add({ id: 'edit.formatPainter', label: N('Format painter'), active: () => !!painter.style, run: (s) => { painter.style = painter.style ? null : structuredClone(activeStyle(s)); s.emit(); } });

  // ---- Font & alignment ----
  toggleFont('bold', N('Bold'), 'Mod+B');
  toggleFont('italic', N('Italic'), 'Mod+I');
  toggleFont('strike', N('Strikethrough'), 'Mod+5');
  add({ id: 'format.underline', label: N('Underline'), group: N('Font'), keys: 'Mod+U', active: (s) => !!activeStyle(s).font?.underline,
    run: (s) => { const u = activeStyle(s).font?.underline ? undefined : 'single' as const; s.applyStyle((st) => ({ ...st, font: { ...st.font, underline: u } }), N('Underline')); } });
  add({ id: 'format.doubleUnderline', label: N('Double underline'), group: N('Font'), run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, underline: st.font?.underline === 'double' ? undefined : 'double' } }), N('Underline')) });
  add({ id: 'format.superscript', label: N('Superscript'), group: N('Font'), active: (s) => activeStyle(s).font?.vertAlign === 'superscript', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, vertAlign: st.font?.vertAlign === 'superscript' ? undefined : 'superscript' } }), N('Superscript')) });
  add({ id: 'format.subscript', label: N('Subscript'), group: N('Font'), active: (s) => activeStyle(s).font?.vertAlign === 'subscript', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, vertAlign: st.font?.vertAlign === 'subscript' ? undefined : 'subscript' } }), N('Subscript')) });
  add({ id: 'format.fontName', label: N('Font'), group: N('Font'), run: (s, name) => s.applyStyle((st) => ({ ...st, font: { ...st.font, name: name as string } }), N('Font')) });
  add({ id: 'format.fontSize', label: N('Font size'), group: N('Font'), run: (s, size) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: Math.max(1, Math.min(409, Number(size))) } }), N('Font size')) });
  add({ id: 'format.growFont', label: N('Increase font size'), group: N('Font'), keys: 'Mod+Shift+.', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: stepSize(st.font?.size ?? 11, 1) } }), N('Font size')) });
  add({ id: 'format.shrinkFont', label: N('Decrease font size'), group: N('Font'), keys: 'Mod+Shift+,', run: (s) => s.applyStyle((st) => ({ ...st, font: { ...st.font, size: stepSize(st.font?.size ?? 11, -1) } }), N('Font size')) });
  add({ id: 'format.fontColor', label: N('Font color'), group: N('Font'), run: (s, c) => s.applyStyle((st) => ({ ...st, font: { ...st.font, color: (c as string) || undefined } }), N('Font color')) });
  add({ id: 'format.fillColor', label: N('Fill color'), group: N('Font'), run: (s, c) => s.applyStyle((st) => ({ ...st, fill: c ? { color: c as string } : undefined }), N('Fill color')) });
  for (const h of ['left', 'center', 'right'] as const) add({ id: `format.align.${h}`, label: ({ left: N('Align left'), center: N('Align center'), right: N('Align right') })[h], group: N('Alignment'), keys: h === 'center' ? 'Mod+Shift+E' : undefined, active: (s) => activeStyle(s).align?.h === h, run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, h: st.align?.h === h ? undefined : h } }), N('Align')) });
  for (const v of ['top', 'center', 'bottom'] as const) add({ id: `format.valign.${v}`, label: ({ top: N('Top align'), center: N('Middle align'), bottom: N('Bottom align') })[v], group: N('Alignment'), active: (s) => (activeStyle(s).align?.v ?? 'bottom') === v, run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, v: v === 'bottom' ? undefined : v } }), N('Align')) });
  add({ id: 'format.wrap', label: N('Wrap text'), group: N('Alignment'), active: (s) => !!activeStyle(s).align?.wrap, run: (s) => { const on = !activeStyle(s).align?.wrap; s.applyStyle((st) => ({ ...st, align: { ...st.align, wrap: on || undefined } }), N('Wrap text')); } });
  add({ id: 'format.indent', label: N('Increase indent'), group: N('Alignment'), run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, indent: Math.min(250, (st.align?.indent ?? 0) + 1), h: st.align?.h === 'right' ? 'right' : st.align?.h === 'center' ? 'left' : st.align?.h ?? 'left' } }), N('Indent')) });
  add({ id: 'format.outdent', label: N('Decrease indent'), group: N('Alignment'), run: (s) => s.applyStyle((st) => ({ ...st, align: { ...st.align, indent: Math.max(0, (st.align?.indent ?? 0) - 1) || undefined } }), N('Indent')) });
  add({ id: 'format.rotate', label: N('Orientation'), group: N('Alignment'), run: (s, deg) => s.applyStyle((st) => ({ ...st, align: { ...st.align, rotation: (deg as number) || undefined } }), N('Orientation')) });
  add({ id: 'format.mergeCenter', label: N('Merge & center'), group: N('Alignment'), active: (s) => !!s.mergeAt(s.sel.active.row, s.sel.active.col), run: (s) => mergeCells(s, 'center') });
  add({ id: 'format.mergeAcross', label: N('Merge across'), group: N('Alignment'), run: (s) => mergeCells(s, 'across') });
  add({ id: 'format.merge', label: N('Merge cells'), group: N('Alignment'), run: (s) => mergeCells(s, 'plain') });
  add({ id: 'format.unmerge', label: N('Unmerge cells'), group: N('Alignment'), run: (s) => mergeCells(s, 'unmerge') });

  // ---- Number ----
  for (const [id, label, code] of NUMBER_FORMATS) add({ id: `number.${id}`, label, group: N('Number'), active: (s) => (activeStyle(s).numFmt ?? 'General') === regionalFormat(code), run: (s) => s.applyStyle((st) => ({ ...st, numFmt: code === 'General' ? undefined : regionalFormat(code) }), N('Number format')) });
  add({ id: 'number.format', label: N('Number format'), group: N('Number'), run: (s, code) => s.applyStyle((st) => ({ ...st, numFmt: (code as string) === 'General' ? undefined : (code as string) }), N('Number format')) });
  add({ id: 'number.moreDecimals', label: N('Increase decimal'), group: N('Number'), run: (s) => changeDecimals(s, 1) });
  add({ id: 'number.lessDecimals', label: N('Decrease decimal'), group: N('Number'), run: (s) => changeDecimals(s, -1) });
  const keyFormats: [string, string, string, string][] = [['Mod+Shift+1', 'number.comma2', '#,##0.00', N('Number with separators')], ['Mod+Shift+4', 'number.currency2', '"$"#,##0.00', N('Currency, 2 decimals')], ['Mod+Shift+5', 'number.percent0', '0%', N('Percent, no decimals')], ['Mod+Shift+3', 'number.date2', 'd-mmm-yy', N('Date: day, month, year')], ['Mod+Shift+2', 'number.time2', 'h:mm AM/PM', N('Time: hours and minutes')], ['Mod+Shift+6', 'number.sci2', '0.00E+00', N('Scientific, 2 decimals')], ['Mod+Shift+`', 'number.general2', N('General'), N('General number format')]];
  for (const [k, id, code, label] of keyFormats) add({ id, label, keys: k, run: (s) => s.applyStyle((st) => ({ ...st, numFmt: code === 'General' ? undefined : regionalFormat(code) }), N('Number format')) });

  // ---- Borders ----
  const borderCmd = (id: string, label: string, f: (r: CellRange, row: number, col: number, side: BorderSide) => CellStyle['border'] | null) => add({
    id: `border.${id}`, label, group: N('Borders'),
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
  borderCmd('all', N('All borders'), (_r, _row, _col, sd) => ({ top: sd, bottom: sd, left: sd, right: sd }));
  borderCmd('outside', N('Outside borders'), (r, row, col, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }), ...(col === r.c1 && { left: sd }), ...(col === r.c2 && { right: sd }) }));
  borderCmd('thickBox', N('Thick outside borders'), (r, row, col, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }), ...(col === r.c1 && { left: sd }), ...(col === r.c2 && { right: sd }) }));
  borderCmd('inside', N('Inside borders'), (r, row, col, sd) => ({ ...(row > r.r1 && { top: sd }), ...(row < r.r2 && { bottom: sd }), ...(col > r.c1 && { left: sd }), ...(col < r.c2 && { right: sd }) }));
  borderCmd('top', N('Top border'), (r, row, _c, sd) => (row === r.r1 ? { top: sd } : null));
  borderCmd('bottom', N('Bottom border'), (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('left', N('Left border'), (r, _row, col, sd) => (col === r.c1 ? { left: sd } : null));
  borderCmd('right', N('Right border'), (r, _row, col, sd) => (col === r.c2 ? { right: sd } : null));
  borderCmd('thickBottom', N('Thick bottom border'), (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('doubleBottom', N('Bottom double border'), (r, row, _c, sd) => (row === r.r2 ? { bottom: sd } : null));
  borderCmd('topBottom', N('Top and bottom border'), (r, row, _c, sd) => ({ ...(row === r.r1 && { top: sd }), ...(row === r.r2 && { bottom: sd }) }));
  borderCmd('none', N('No border'), () => ({}));

  // ---- Cells, rows, columns ----
  add({ id: 'cells.insertRows', label: N('Insert sheet rows'), group: N('Cells'), keys: 'Mod+Shift+=', run: (s) => { const r = s.range; changeStructure(s, 'row', r.r1, r.r2 - r.r1 + 1); } });
  add({ id: 'cells.insertCols', label: N('Insert sheet columns'), group: N('Cells'), run: (s) => { const r = s.range; changeStructure(s, 'col', r.c1, r.c2 - r.c1 + 1); } });
  add({ id: 'cells.deleteRows', label: N('Delete sheet rows'), group: N('Cells'), keys: 'Mod+-', run: (s) => { const r = s.range; changeStructure(s, 'row', r.r1, -(r.r2 - r.r1 + 1)); } });
  add({ id: 'cells.deleteCols', label: N('Delete sheet columns'), group: N('Cells'), run: (s) => { const r = s.range; changeStructure(s, 'col', r.c1, -(r.c2 - r.c1 + 1)); } });
  add({ id: 'cells.insertDialog', label: N('Insert cells…'), group: N('Cells'), run: (s) => openDialog(s, 'insertCells', { insert: true }) });
  add({ id: 'cells.deleteDialog', label: N('Delete cells…'), group: N('Cells'), run: (s) => openDialog(s, 'insertCells', { insert: false }) });
  add({ id: 'cells.hideRows', label: N('Hide rows'), group: N('Cells'), keys: 'Mod+9', run: (s) => setLinesHidden(s, 'row', s.range.r1, s.range.r2, true) });
  add({ id: 'cells.unhideRows', label: N('Unhide rows'), group: N('Cells'), keys: 'Mod+Shift+9', run: (s) => setLinesHidden(s, 'row', Math.max(0, s.range.r1 - (s.range.r1 === s.range.r2 ? 0 : 0)), s.range.r2, false) });
  add({ id: 'cells.hideCols', label: N('Hide columns'), group: N('Cells'), keys: 'Mod+0', run: (s) => setLinesHidden(s, 'col', s.range.c1, s.range.c2, true) });
  add({ id: 'cells.unhideCols', label: N('Unhide columns'), group: N('Cells'), keys: 'Mod+Shift+0', run: (s) => setLinesHidden(s, 'col', s.range.c1, s.range.c2, false) });
  add({ id: 'cells.autofitCols', label: N('AutoFit column width'), group: N('Cells'), run: (s) => autofit(s, 'col') });
  add({ id: 'cells.autofitRows', label: N('AutoFit row height'), group: N('Cells'), run: (s) => autofit(s, 'row') });
  add({ id: 'cells.colWidth', label: N('Column width…'), group: N('Cells'), run: (s) => openDialog(s, 'size', { axis: 'col' }) });
  add({ id: 'cells.rowHeight', label: N('Row height…'), group: N('Cells'), run: (s) => openDialog(s, 'size', { axis: 'row' }) });
  add({ id: 'cells.formatDialog', label: N('Format cells…'), group: N('Cells'), keys: 'Mod+1', run: (s) => openDialog(s, 'formatCells') });
  add({ id: 'cells.styles', label: N('Cell styles'), group: N('Cells'), run: (s) => openDialog(s, 'cellStyles') });
  add({ id: 'cells.formatTable', label: N('Format as table'), group: N('Cells'), keys: 'Mod+T', run: (s) => openDialog(s, 'table') });
  add({ id: 'cells.group', label: N('Group'), group: N('Outline'), keys: 'Alt+Shift+ArrowRight', run: (s) => { const r = s.range; if (r.c1 === 0 && r.c2 === MAX_COLS - 1 || r.c1 === r.c2) groupLines(s, 'row', r.r1, r.r2, 1); else groupLines(s, 'col', r.c1, r.c2, 1); } });
  add({ id: 'cells.ungroup', label: N('Ungroup'), group: N('Outline'), keys: 'Alt+Shift+ArrowLeft', run: (s) => { const r = s.range; if (r.c1 === 0 && r.c2 === MAX_COLS - 1 || r.c1 === r.c2) groupLines(s, 'row', r.r1, r.r2, -1); else groupLines(s, 'col', r.c1, r.c2, -1); } });

  // ---- Sheets ----
  add({ id: 'sheet.insert', label: N('Insert sheet'), group: N('Sheet'), keys: 'Shift+F11', run: (s) => addSheet(s) });
  add({ id: 'sheet.delete', label: N('Delete sheet'), group: N('Sheet'), run: (s) => deleteSheet(s) });
  add({ id: 'sheet.duplicate', label: N('Duplicate sheet'), group: N('Sheet'), run: (s) => duplicateSheet(s) });
  add({ id: 'sheet.rename', label: N('Rename sheet'), group: N('Sheet'), run: (s) => openDialog(s, 'renameSheet', { index: s.sheetIndex }) });
  add({ id: 'sheet.next', label: N('Next sheet'), keys: 'Mod+PageDown', run: (s) => { let i = s.sheetIndex + 1; while (s.wb.sheets[i] && s.wb.sheets[i]!.state !== 'visible') i++; if (s.wb.sheets[i]) s.setActiveSheet(i); } });
  add({ id: 'sheet.prev', label: N('Previous sheet'), keys: 'Mod+PageUp', run: (s) => { let i = s.sheetIndex - 1; while (i >= 0 && s.wb.sheets[i]!.state !== 'visible') i--; if (i >= 0) s.setActiveSheet(i); } });
  add({ id: 'sheet.protect', label: N('Protect sheet'), group: N('Review'), active: (s) => !!s.sheet.protection?.sheet, run: (s) => openDialog(s, 'protect') });

  // ---- Formulas ----
  for (const fn of ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN']) add({ id: `formula.auto${fn}`, label: ({ SUM: N('AutoSum'), AVERAGE: N('Average'), COUNT: N('Count numbers'), MAX: N('Max'), MIN: N('Min') } as Record<string, string>)[fn]!, group: N('Formulas'), keys: fn === 'SUM' ? 'Alt+=' : undefined, run: (s) => autoSum(s, fn) });
  add({ id: 'formula.insertFunction', label: N('Insert function…'), group: N('Formulas'), keys: 'Shift+F3', run: (s) => openDialog(s, 'insertFunction') });
  add({ id: 'formula.names', label: N('Name manager'), group: N('Formulas'), keys: 'Mod+F3', run: (s) => openDialog(s, 'names') });
  add({ id: 'formula.defineName', label: N('Define name…'), group: N('Formulas'), run: (s) => openDialog(s, 'defineName') });
  add({ id: 'formula.createFromSelection', label: N('Create names from selection'), group: N('Formulas'), keys: 'Mod+Shift+F3', run: (s) => createNamesFromSelection(s) });
  add({ id: 'formula.showFormulas', label: N('Show formulas'), group: N('Formulas'), keys: 'Mod+`', active: (s) => s.view.showFormulas, run: (s) => { s.view.showFormulas = !s.view.showFormulas; s.touch(); } });
  add({ id: 'formula.calcNow', label: N('Calculate now'), group: N('Formulas'), keys: 'F9', run: (s) => { s.engine.recalc(true); s.touch(); } });
  add({ id: 'formula.iterative', label: N('Iterative calculation'), group: N('Formulas'), active: (s) => s.engine.iterate, run: (s) => { s.engine.iterate = !s.engine.iterate; s.wb.calc = { ...s.wb.calc, iterate: s.engine.iterate }; s.engine.recalc(true); s.touch(); } });
  add({ id: 'formula.evaluate', label: N('Evaluate formula'), group: N('Formulas'), run: (s) => openDialog(s, 'evaluate') });
  add({ id: 'formula.tracePrecedents', label: N('Trace precedents'), group: N('Formulas'), run: (s) => openDialog(s, 'trace', { dir: 'precedents' }) });
  add({ id: 'formula.traceDependents', label: N('Trace dependents'), group: N('Formulas'), run: (s) => openDialog(s, 'trace', { dir: 'dependents' }) });
  add({ id: 'formula.errorCheck', label: N('Error checking'), group: N('Formulas'), run: (s) => openDialog(s, 'errors') });
  add({ id: 'formula.arrayEnter', label: N('Enter as array formula'), keys: 'Mod+Shift+Enter', run: (s) => { if (s.editing) s.commitEdit(false, true); } });

  // ---- Data ----
  add({ id: 'data.sortAsc', label: N('Sort A to Z'), group: N('Data'), run: (s) => quickSort(s, false) });
  add({ id: 'data.sortDesc', label: N('Sort Z to A'), group: N('Data'), run: (s) => quickSort(s, true) });
  add({ id: 'data.sort', label: N('Custom sort…'), group: N('Data'), run: (s) => openDialog(s, 'sort') });
  add({ id: 'data.filter', label: N('Filter'), group: N('Data'), keys: 'Mod+Shift+L', active: (s) => !!s.sheet.autoFilter, run: (s) => toggleFilter(s) });
  add({ id: 'data.clearFilter', label: N('Clear filters'), group: N('Data'), enabled: (s) => !!s.sheet.autoFilter?.columns.length, run: (s) => clearFilters(s) });
  add({ id: 'data.reapply', label: N('Reapply filter'), group: N('Data'), enabled: (s) => !!s.sheet.autoFilter, run: (s) => s.change(N('Reapply'), 'sheet', () => applyFilter(s)) });
  add({ id: 'data.removeDuplicates', label: N('Remove duplicates…'), group: N('Data'), run: (s) => openDialog(s, 'removeDuplicates') });
  add({ id: 'data.textToColumns', label: N('Text to columns…'), group: N('Data'), run: (s) => openDialog(s, 'textToColumns') });
  add({ id: 'data.validation', label: N('Data validation…'), group: N('Data'), run: (s) => openDialog(s, 'validation') });
  add({ id: 'data.goalSeek', label: N('Goal Seek…'), group: N('Data'), run: (s) => openDialog(s, 'goalSeek') });
  add({ id: 'data.pivot', label: N('PivotTable…'), group: N('Data'), run: (s) => openDialog(s, 'pivot') });
  add({ id: 'data.subtotal', label: N('Subtotal…'), group: N('Data'), run: (s) => openDialog(s, 'subtotals') });
  add({ id: 'data.whatIfTable', label: N('Data table…'), group: N('Data'), run: (s) => openDialog(s, 'dataTable') });
  add({ id: 'data.importCsv', label: N('From text/CSV…'), group: N('Data'), run: (s) => openDialog(s, 'importCsv') });
  add({ id: 'data.stats', label: N('Descriptive statistics'), group: N('Data'), run: (s) => openDialog(s, 'stats') });

  // ---- Insert ----
  add({ id: 'insert.chart', label: N('Chart…'), group: N('Insert'), keys: 'Alt+F1', run: (s) => openDialog(s, 'chart') });
  add({ id: 'insert.picture', label: N('Picture…'), group: N('Insert'), run: (s) => openDialog(s, 'picture') });
  add({ id: 'insert.link', label: N('Link…'), group: N('Insert'), keys: 'Mod+K', run: (s) => openDialog(s, 'link') });
  add({ id: 'insert.note', label: N('Note'), group: N('Insert'), keys: 'Shift+F2', run: (s) => openDialog(s, 'note') });
  add({ id: 'insert.symbol', label: N('Symbol…'), group: N('Insert'), run: (s) => openDialog(s, 'symbol') });
  add({ id: 'insert.conditional', label: N('Conditional formatting…'), group: N('Styles'), run: (s) => openDialog(s, 'conditional') });
  add({ id: 'insert.clearRules', label: N('Clear rules from selection'), group: N('Styles'), run: (s) => clearRules(s, false) });
  add({ id: 'insert.clearAllRules', label: N('Clear rules from sheet'), group: N('Styles'), run: (s) => clearRules(s, true) });

  // ---- View ----
  add({ id: 'view.freezePanes', label: N('Freeze panes'), group: N('View'), active: (s) => !!(s.sheet.view.frozenRows || s.sheet.view.frozenCols), run: (s) => freeze(s, 'selection') });
  add({ id: 'view.freezeTopRow', label: N('Freeze top row'), group: N('View'), run: (s) => freeze(s, 'row') });
  add({ id: 'view.freezeFirstCol', label: N('Freeze first column'), group: N('View'), run: (s) => freeze(s, 'col') });
  add({ id: 'view.unfreeze', label: N('Unfreeze panes'), group: N('View'), run: (s) => freeze(s, 'none') });
  add({ id: 'view.gridlines', label: N('Gridlines'), group: N('View'), active: (s) => s.sheet.view.showGrid, run: (s) => { s.sheet.view.showGrid = !s.sheet.view.showGrid; s.file.dirty = true; s.touch(); } });
  add({ id: 'view.headings', label: N('Headings'), group: N('View'), active: (s) => s.sheet.view.showHeaders, run: (s) => { s.sheet.view.showHeaders = !s.sheet.view.showHeaders; s.touch(); } });
  add({ id: 'view.zoomIn', label: N('Zoom in'), group: N('View'), keys: 'Mod+=', run: (s) => setZoom(s, s.sheet.view.zoom + 10) });
  add({ id: 'view.zoomOut', label: N('Zoom out'), group: N('View'), keys: 'Mod+Shift+-', run: (s) => setZoom(s, s.sheet.view.zoom - 10) });
  add({ id: 'view.zoom100', label: N('Zoom to 100%'), group: N('View'), run: (s) => setZoom(s, 100) });
  add({ id: 'view.zoomSelection', label: N('Zoom to selection'), group: N('View'), run: (s) => openDialog(s, 'zoomToSelection') });
  add({ id: 'view.theme', label: N('Dark mode'), group: N('View'), active: (s) => s.view.theme === 'dark' || (s.view.theme === 'system' && matchMedia('(prefers-color-scheme: dark)').matches), run: (s) => { const dark = s.view.theme === 'dark' || (s.view.theme === 'system' && matchMedia('(prefers-color-scheme: dark)').matches); s.view.theme = dark ? 'light' : 'dark'; try { localStorage.setItem('clerdogrid.theme', s.view.theme); } catch { /* private mode */ } s.touch(); } });
  add({ id: 'view.rtl', label: N('Sheet right to left'), group: N('View'), active: (s) => s.sheet.view.rtl, run: (s) => { s.sheet.view.rtl = !s.sheet.view.rtl; s.file.dirty = true; s.touch(); } });

  // ---- File ----
  add({ id: 'file.new', label: N('New workbook'), group: N('File'), keys: 'Mod+N', run: (s) => void import('./files').then((m) => m.newWorkbookFile(s)) });
  add({ id: 'file.open', label: N('Open…'), group: N('File'), keys: 'Mod+O', run: (s) => void import('./files').then((m) => m.openFile(s)) });
  add({ id: 'file.save', label: N('Save'), group: N('File'), keys: 'Mod+S', run: (s) => void import('./files').then((m) => m.saveFile(s)) });
  add({ id: 'file.saveAs', label: N('Save as…'), group: N('File'), keys: 'Mod+Shift+S', run: (s) => void import('./files').then((m) => m.saveFile(s, true)) });
  add({ id: 'file.exportCsv', label: N('Export sheet as CSV'), group: N('File'), run: (s) => void import('./files').then((m) => m.exportCsv(s)) });
  add({ id: 'file.exportPdf', label: N('Export as PDF / print'), group: N('File'), keys: 'Mod+P', run: (s) => void import('./files').then((m) => m.printSheet(s)) });
  add({ id: 'file.templates', label: N('Templates…'), group: N('File'), run: (s) => openDialog(s, 'templates') });
  add({ id: 'file.properties', label: N('Workbook properties'), group: N('File'), run: (s) => openDialog(s, 'properties') });
  add({ id: 'help.shortcuts', label: N('Keyboard shortcuts'), group: N('Help'), keys: 'Mod+/', run: (s) => openDialog(s, 'shortcuts') });
  add({ id: 'help.about', label: N('About ClerdoGrid'), group: N('Help'), run: (s) => openDialog(s, 'about') });
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
    s.notify(T('Some of these cells are protected. Unprotect the sheet to change them.'));
    return;
  }
  s.change(({ contents: N('Clear'), formats: N('Clear formats'), all: N('Clear all'), notes: N('Clear notes'), links: N('Clear hyperlinks') })[what], whole ? 'sheet' : areas.map((area) => ({ area })), () => {
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
    s.change(N('Unmerge'), 'sheet', () => { s.sheet.merges = s.sheet.merges.filter((m) => !existing.includes(m)); });
    return;
  }
  if (r.r1 === r.r2 && r.c1 === r.c2) return;
  const rows = how === 'across' ? Array.from({ length: r.r2 - r.r1 + 1 }, (_, i) => ({ ...r, r1: r.r1 + i, r2: r.r1 + i })) : [r];
  let lost = 0;
  for (const m of rows) for (let row = m.r1; row <= m.r2; row++) for (let col = m.c1; col <= m.c2; col++) if ((row !== m.r1 || col !== m.c1) && s.value(row, col) !== null) lost++;
  if (lost && !confirm(T('Merging cells only keeps the upper-left value and discards the other values.'))) return;
  s.change(N('Merge'), 'sheet', () => {
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
  }, d > 0 ? N('Increase decimal') : N('Decrease decimal'));
};

const autofit = (s: GridStore, axis: 'col' | 'row'): void => {
  const r = s.range;
  const canvas = document.createElement('canvas').getContext('2d')!;
  s.change(N('AutoFit'), 'sheet', () => {
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
  s.change(N('Create names'), 'book', () => {
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
  s.notify(T('Names created from the top row.'));
};

const clearRules = (s: GridStore, all: boolean): void => {
  s.change(N('Clear rules'), 'sheet', () => {
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
  s.applyStyle(() => st, N('Format painter'));
};

export { canonical, rangeAddr, cellKey };
