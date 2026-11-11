// Comments, notes, shapes, checkboxes, ink and review tools.
import { createElement } from 'react';
import { registerCommand, openDialog } from '../../app/commands';
import { cellKey, type ShapeKind } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { registerGridHook, registerOverlay, registerPanelSection } from '../../ui/panelRegistry';
import { registerDialog } from '../../ui/dialogs/extra';
import { selectedChart } from '../../ui/charts/ChartView';
import { anchorAt, lastVp, setUi, ui, visibleOrigin, cellBox } from './state';
import {
  addShape, clearInk, defaultShape, deleteComment, deleteNote, deleteShape, insertCheckboxes, isCheckbox, removeCheckboxes, resolveComment,
  stepTo, threadAt, toggleCheckboxes,
} from './ops';
import { ShapeLayer } from './ShapeLayer';
import { InkLayer } from './InkLayer';
import { CommentsPane, PopoverLayer } from './CommentLayer';
import { AccessibilityDialog, AltTextDialog, EditRangesDialog, PerformanceDialog, SpellingDialog, StatsDialog } from './dialogs';
import { checkSection, checkboxSection, commentsSection, drawInkSection, drawStyleSection, drawToolsSection, notesSection, shapesSection } from './panels';
import './objects.css';
import { N, T } from '../../ui/i18n';

const at = (s: GridStore) => s.sel.active;
const pinAt = (s: GridStore, row: number, col: number, compose = false) => setUi(s, { pinned: { sheet: s.sheetIndex, row, col, compose }, hover: null });

const insertShape = (s: GridStore, kind: ShapeKind) => {
  const vp = lastVp.current;
  const w = kind === 'textBox' ? 180 : 140, h = kind === 'textBox' ? 60 : kind.endsWith('Arrow') && (kind === 'rightArrow' || kind === 'leftArrow') ? 70 : 110;
  const n = s.sheet.shapes?.length ?? 0;
  let anchor = { row: s.sel.active.row, col: s.sel.active.col + 1, dx: 8, dy: 8, width: w, height: h };
  if (vp) {
    const o = visibleOrigin(vp);
    anchor = anchorAt(vp, o.x + 60 + (n % 6) * 18, o.y + 40 + (n % 6) * 18, w, h);
  }
  const shape = defaultShape(kind, anchor);
  ui.tool = 'select';
  addShape(s, shape);
  setUi(s, { shape: shape.id, editingShape: kind === 'textBox' ? shape.id : null, pinned: null, hover: null });
};

export const register = (): void => {
  // ---------- Comments ----------
  registerCommand({ id: 'comment.new', label: N('New comment'), group: N('Review'), keys: 'Mod+Alt+M', run: (s) => { const { row, col } = at(s); pinAt(s, row, col, !threadAt(s, row, col)); } });
  registerCommand({ id: 'comment.pane', label: N('Show comments'), group: N('Review'), active: () => ui.commentsPane, run: (s) => setUi(s, { commentsPane: !ui.commentsPane }) });
  registerCommand({ id: 'comment.next', label: N('Next comment'), group: N('Review'), enabled: (s) => !!s.sheet.comments?.size, run: (s) => { const p = stepTo(s, s.sheet.comments?.keys() ?? [], 1); if (p) { s.select(p); pinAt(s, p.row, p.col); } } });
  registerCommand({ id: 'comment.prev', label: N('Previous comment'), group: N('Review'), enabled: (s) => !!s.sheet.comments?.size, run: (s) => { const p = stepTo(s, s.sheet.comments?.keys() ?? [], -1); if (p) { s.select(p); pinAt(s, p.row, p.col); } } });
  registerCommand({ id: 'comment.resolve', label: N('Resolve or reopen'), group: N('Review'), enabled: (s) => !!threadAt(s, at(s).row, at(s).col), run: (s) => { const { row, col } = at(s); resolveComment(s, row, col, !threadAt(s, row, col)!.resolved); } });
  registerCommand({ id: 'comment.delete', label: N('Delete comment'), group: N('Review'), enabled: (s) => !!threadAt(s, at(s).row, at(s).col), run: (s) => { const { row, col } = at(s); deleteComment(s, row, col); setUi(s, { pinned: null }); } });

  // ---------- Notes ----------
  registerCommand({ id: 'note.delete', label: N('Delete note'), group: N('Review'), enabled: (s) => s.sheet.notes.size > 0, run: (s) => deleteNote(s) });
  registerCommand({ id: 'note.showAll', label: N('Show all notes'), group: N('Review'), active: () => ui.showAllNotes, run: (s) => setUi(s, { showAllNotes: !ui.showAllNotes }) });
  registerCommand({ id: 'note.next', label: N('Next note'), group: N('Review'), enabled: (s) => s.sheet.notes.size > 0, run: (s) => { const p = stepTo(s, s.sheet.notes.keys(), 1); if (p) { s.select(p); pinAt(s, p.row, p.col); } } });
  registerCommand({ id: 'note.prev', label: N('Previous note'), group: N('Review'), enabled: (s) => s.sheet.notes.size > 0, run: (s) => { const p = stepTo(s, s.sheet.notes.keys(), -1); if (p) { s.select(p); pinAt(s, p.row, p.col); } } });

  // ---------- Checkboxes, shapes, ink ----------
  registerCommand({ id: 'insert.checkbox', label: N('Checkbox'), group: N('Insert'), run: insertCheckboxes });
  registerCommand({ id: 'insert.removeCheckbox', label: N('Remove checkboxes'), group: N('Insert'), run: removeCheckboxes });
  registerCommand({ id: 'insert.shape', label: N('Shape'), group: N('Insert'), run: (s, kind) => insertShape(s, (typeof kind === 'string' ? kind : 'rect') as ShapeKind) });
  registerCommand({ id: 'insert.textBox', label: N('Text box'), group: N('Insert'), run: (s) => insertShape(s, 'textBox') });
  registerCommand({ id: 'draw.clear', label: N('Clear ink on this sheet'), group: N('Draw'), enabled: (s) => !!s.sheet.ink?.length, run: clearInk });
  registerCommand({ id: 'draw.select', label: N('Select tool'), group: N('Draw'), run: (s) => setUi(s, { tool: 'select' }) });
  registerCommand({ id: 'draw.pen', label: N('Pen'), group: N('Draw'), run: (s) => setUi(s, { tool: 'pen' }) });
  registerCommand({ id: 'draw.highlighter', label: N('Highlighter'), group: N('Draw'), run: (s) => setUi(s, { tool: 'highlighter' }) });
  registerCommand({ id: 'draw.eraser', label: N('Eraser'), group: N('Draw'), run: (s) => setUi(s, { tool: 'eraser' }) });

  // ---------- Review tools ----------
  registerCommand({ id: 'review.spelling', label: N('Spelling'), group: N('Review'), keys: 'F7', run: (s) => openDialog(s, 'spelling') });
  registerCommand({ id: 'review.stats', label: N('Workbook statistics'), group: N('Review'), run: (s) => openDialog(s, 'workbookStats') });
  registerCommand({ id: 'review.accessibility', label: N('Check accessibility'), group: N('Review'), run: (s) => openDialog(s, 'accessibility') });
  registerCommand({ id: 'review.performance', label: N('Check performance'), group: N('Review'), run: (s) => openDialog(s, 'performance') });
  registerCommand({ id: 'review.editRanges', label: N('Allow edit ranges'), group: N('Review'), run: (s) => openDialog(s, 'editRanges') });
  registerCommand({
    id: 'review.altText', label: N('Alt text'), group: N('Review'),
    enabled: (s) => !!ui.shape || !!selectedChart() || s.sheet.pictures.length > 0,
    run: (s) => {
      if (ui.shape) openDialog(s, 'altText', { type: 'shape', id: ui.shape });
      else if (selectedChart()) openDialog(s, 'altText', { type: 'chart', id: selectedChart() });
      else if (s.sheet.pictures[0]) openDialog(s, 'altText', { type: 'picture', id: s.sheet.pictures[0].id });
    },
  });

  registerDialog('spelling', SpellingDialog);
  registerDialog('workbookStats', StatsDialog);
  registerDialog('accessibility', AccessibilityDialog);
  registerDialog('performance', PerformanceDialog);
  registerDialog('editRanges', EditRangesDialog);
  registerDialog('altText', AltTextDialog);

  // ---------- Panels ----------
  registerPanelSection({ tab: 'insert', title: N('Shapes'), order: 10, render: shapesSection });
  registerPanelSection({ tab: 'insert', title: N('Checkboxes'), order: 11, render: checkboxSection });
  registerPanelSection({ tab: 'insert', title: N('Comments'), order: 12, render: (s) => commentsSection(s) });
  registerPanelSection({ tab: 'review', title: N('Comments'), order: 1, render: commentsSection });
  registerPanelSection({ tab: 'review', title: N('Note tools'), order: 2, render: notesSection });
  registerPanelSection({ tab: 'review', title: N('Review tools'), order: 3, render: checkSection });
  registerPanelSection({ tab: 'draw', title: N('Tools'), order: 1, render: drawToolsSection });
  registerPanelSection({ tab: 'draw', title: N('Style'), order: 2, render: drawStyleSection });
  registerPanelSection({ tab: 'draw', title: N('Ink'), order: 3, render: drawInkSection });

  // ---------- Overlays (order: shapes under ink under popovers) ----------
  registerOverlay({ id: 'objects.shapes', order: 10, render: (s, vp) => createElement(ShapeLayer, { store: s, vp }) });
  registerOverlay({ id: 'objects.ink', order: 20, render: (s, vp) => createElement(InkLayer, { store: s, vp }) });
  registerOverlay({ id: 'objects.popovers', order: 30, render: (s, vp) => createElement(PopoverLayer, { store: s, vp }) });
  registerOverlay({ id: 'objects.pane', order: 40, render: (s, vp) => createElement(CommentsPane, { store: s, vp }) });

  // ---------- Grid input ----------
  registerGridHook({
    id: 'objects',
    cellTap: (s, h) => {
      if (ui.shape || ui.editingShape) setUi(s, { shape: null, editingShape: null });
      const k = cellKey(h.row, h.col);
      const vp = lastVp.current;
      // A tap on a checkbox's square toggles it; elsewhere in the cell just selects.
      if (isCheckbox(s, h.row, h.col) && typeof s.value(h.row, h.col) === 'boolean' && vp) {
        const b = cellBox(vp, h.row, h.col, s);
        const size = Math.max(16, Math.min(14 * vp.zoom, b.height - 4) + 8);
        if (Math.abs(h.x - (b.left + b.width / 2)) <= size / 2 + 2 && Math.abs(h.y - (b.top + b.height / 2)) <= size / 2 + 2) {
          s.select({ row: h.row, col: h.col });
          if (!toggleCheckboxes(s, [{ row: h.row, col: h.col }]) && s.isLocked(h.row, h.col)) s.notify(T('This cell is protected.'));
          return true;
        }
      }
      // Touch has no hover: a tap on a cell with a note or comment opens its popover.
      if (h.pointerType !== 'mouse' && (s.sheet.notes.has(k) || s.sheet.comments?.has(k))) { s.select({ row: h.row, col: h.col }); pinAt(s, h.row, h.col); return true; }
      if (ui.pinned && (ui.pinned.row !== h.row || ui.pinned.col !== h.col)) setUi(s, { pinned: null });
      return false;
    },
    hover: (s, h) => {
      const k = h ? cellKey(h.row, h.col) : -1;
      const has = !!h && (s.sheet.notes.has(k) || !!s.sheet.comments?.has(k));
      const next = has ? { sheet: s.sheetIndex, row: h!.row, col: h!.col } : null;
      const cur = ui.hover;
      if ((cur?.row ?? -1) === (next?.row ?? -1) && (cur?.col ?? -1) === (next?.col ?? -1) && cur?.sheet === next?.sheet) return;
      hoverTimer && clearTimeout(hoverTimer);
      if (!next) { hoverTimer = setTimeout(() => setUi(s, { hover: null }), 250); return; }
      hoverTimer = setTimeout(() => setUi(s, { hover: next }), 120);
    },
    key: (s, ev) => {
      if ((ev.key === 'Delete' || ev.key === 'Backspace') && ui.shape && !ui.editingShape) { deleteShape(s, ui.shape); setUi(s, { shape: null }); return true; }
      if (ev.key === 'Escape' && (ui.shape || ui.pinned || ui.tool !== 'select')) { setUi(s, { shape: null, pinned: null, tool: ui.shape || ui.pinned ? ui.tool : 'select' }); return true; }
      if (ev.key === 'Enter' && ui.shape && !ui.editingShape) { setUi(s, { editingShape: ui.shape }); return true; }
      if (ev.key === ' ' && !ev.shiftKey && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ui.shape) {
        const { row, col } = s.sel.active;
        if (isCheckbox(s, row, col)) return toggleCheckboxes(s) || true;
      }
      return false;
    },
  });
};

let hoverTimer: ReturnType<typeof setTimeout> | null = null;
