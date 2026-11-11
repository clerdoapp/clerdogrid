// PivotTables, PivotCharts, slicers, timelines, GETPIVOTDATA and outlines.
import { createElement } from 'react';
import { command, openDialog, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerGridHook, registerMenuProvider, registerOverlay, registerPanelSection, type MenuItem } from '../../ui/panelRegistry';
import { autoOutline, clearOutline, groupLines, groupTarget, selectVisible, showDetail } from '../../app/ops/outline';
import { deletePivot, drawnArea, drillDown, pivotCellAt, refreshOnOpen } from './apply';
import { dataFieldName, fieldCaption } from './compute';
import { getPivotDataText } from './getpivotdata';
import {
  clearFieldFilters, dropField, expandField, groupSelection, insertPivotChart, moveItem, refreshAll, refreshCurrent, setSubtotals, sortField, toggleItem, ungroup,
} from './ops';
import { activePivot, currentPivot, focusField, focusOf, pui, save, setPui } from './ui';
import { CellButtons, PivotChartButtons } from './CellButtons';
import { FieldList } from './FieldList';
import { SlicerLayer } from './SlicerLayer';
import {
  CalcFieldDialog, CalcItemDialog, ChangeSourceDialog, FieldSettingsDialog, FilterDialog, GroupDialog, MovePivotDialog, PivotOptionsDialog, ShowPagesDialog, ValueFieldDialog,
} from './dialogs';
import { InsertSlicerDialog, InsertTimelineDialog, SlicerSettingsDialog } from './slicerDialogs';
import { deleteSlicer } from './slicers';
import { ConsolidateDialog, OutlineSettingsDialog, SubtotalDialog } from './outline/dialogs';
import { insertPivotSection, outlineSection, pivotCalcSection, pivotLayoutSection, pivotSection, pivotStyleSection, slicerSection } from './panels';
import './pivot.css';
import { N, T } from '../../ui/i18n';

const target = (s: GridStore) => { const r = currentPivot(s); return r ? { sheet: r.sheet, name: r.pt.name } : undefined; };
const need = (s: GridStore): boolean => { if (currentPivot(s)) return true; s.notify(T('Select a cell in a PivotTable.')); return false; };
const inPivot = (s: GridStore) => !!currentPivot(s);

const openFieldSettings = (s: GridStore): void => {
  const f = focusOf(s);
  if (!f) { s.notify(T('Select a cell in a PivotTable.')); return; }
  const c = f.cell;
  if (c && (c.role === 'data' || c.role === 'dataCaption' || (c.field === -2 && c.slot !== undefined)) && (c.data !== undefined || c.slot !== undefined)) {
    openDialog(s, 'pivotValueField', { sheet: f.ref.sheet, name: f.ref.pt.name, index: c.data ?? c.slot ?? 0 });
    return;
  }
  const field = focusField(f);
  if (field === null) { if (f.ref.pt.data.length) openDialog(s, 'pivotValueField', { sheet: f.ref.sheet, name: f.ref.pt.name, index: 0 }); else s.notify(T('Select a field label in the PivotTable.')); return; }
  openDialog(s, 'pivotFieldSettings', { sheet: f.ref.sheet, name: f.ref.pt.name, field });
};

export const register = (): void => {
  // ---------- PivotTables ----------
  registerCommand({ id: 'pivot.fieldList', label: N('Field list'), group: N('PivotTable'), active: () => pui.pane, run: (s) => setPui(s, { pane: !pui.pane }) });
  registerCommand({ id: 'pivot.refresh', label: N('Refresh'), group: N('PivotTable'), keys: 'Alt+F5', enabled: inPivot, run: refreshCurrent });
  registerCommand({ id: 'pivot.refreshAll', label: N('Refresh all'), group: N('PivotTable'), keys: 'Mod+Alt+F5', enabled: (s) => !!s.wb.pivotCaches?.length, run: refreshAll });
  registerCommand({ id: 'pivot.changeSource', label: N('Change data source…'), group: N('PivotTable'), enabled: inPivot, run: (s) => need(s) && openDialog(s, 'pivotSource', target(s)) });
  registerCommand({ id: 'pivot.options', label: N('PivotTable options…'), group: N('PivotTable'), enabled: inPivot, run: (s) => need(s) && openDialog(s, 'pivotOptions', target(s)) });
  registerCommand({ id: 'pivot.fieldSettings', label: N('Field settings…'), group: N('PivotTable'), enabled: inPivot, run: openFieldSettings });
  registerCommand({ id: 'pivot.select', label: N('Select PivotTable'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const r = currentPivot(s); const a = r && drawnArea(r.pt); if (a) s.selectRange(a, { row: r!.pt.anchor.row, col: r!.pt.anchor.col }); } });
  registerCommand({ id: 'pivot.move', label: N('Move PivotTable…'), group: N('PivotTable'), enabled: inPivot, run: (s) => need(s) && openDialog(s, 'pivotMove', target(s)) });
  registerCommand({ id: 'pivot.delete', label: N('Delete PivotTable'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const r = currentPivot(s); if (r) deletePivot(s, r.sheet, r.pt.name); } });
  registerCommand({ id: 'pivot.gpd', label: N('Make GETPIVOTDATA formulas'), group: N('PivotTable'), active: () => pui.gpd, run: (s) => { save('clerdogrid.pivot.gpd', !pui.gpd); setPui(s, { gpd: !pui.gpd }); } });
  registerCommand({ id: 'pivot.expandAll', label: N('Expand entire field'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const f = focusOf(s); if (f) expandField(s, f.ref, focusField(f), false); } });
  registerCommand({ id: 'pivot.collapseAll', label: N('Collapse entire field'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const f = focusOf(s); if (f) expandField(s, f.ref, focusField(f), true); } });
  registerCommand({
    id: 'pivot.group', label: N('Group selection'), group: N('PivotTable'), enabled: inPivot,
    run: (s) => {
      const r = groupSelection(s);
      if (r === 'range') { const f = focusOf(s); const field = focusField(f); if (f && field !== null) openDialog(s, 'pivotGroup', { sheet: f.ref.sheet, name: f.ref.pt.name, field }); }
      else if (r) s.notify(r);
    },
  });
  registerCommand({ id: 'pivot.ungroup', label: N('Ungroup'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const e = ungroup(s); if (e) s.notify(e); } });
  registerCommand({ id: 'pivot.calcField', label: N('Calculated field…'), group: N('PivotTable'), enabled: inPivot, run: (s) => need(s) && openDialog(s, 'pivotCalcField', target(s)) });
  registerCommand({ id: 'pivot.calcItem', label: N('Calculated item…'), group: N('PivotTable'), enabled: inPivot, run: (s) => { if (!need(s)) return; const f = focusOf(s); openDialog(s, 'pivotCalcItem', { ...target(s)!, field: focusField(f) ?? undefined }); } });
  registerCommand({ id: 'pivot.showPages', label: N('Show report filter pages…'), group: N('PivotTable'), enabled: (s) => !!currentPivot(s)?.pt.pages.length, run: (s) => openDialog(s, 'pivotShowPages', target(s)) });
  registerCommand({ id: 'pivot.drillDown', label: N('Show details'), group: N('PivotTable'), enabled: (s) => focusOf(s)?.cell?.role === 'data', run: (s) => { if (!drillDown(s, s.sheetIndex, s.sel.active.row, s.sel.active.col)) s.notify(T('Select a value in a PivotTable.')); } });
  registerCommand({ id: 'pivot.clearFilters', label: N('Clear all filters'), group: N('PivotTable'), enabled: inPivot, run: (s) => { const r = currentPivot(s); if (r) clearFieldFilters(s, r, null); } });
  registerCommand({ id: 'insert.pivotChart', label: N('PivotChart'), group: N('Insert'), enabled: inPivot, run: (s) => { const r = currentPivot(s); if (r) insertPivotChart(s, r); else s.notify(T('Select a cell in a PivotTable.')); } });
  registerCommand({ id: 'insert.slicer', label: N('Slicer…'), group: N('Insert'), run: (s) => openDialog(s, 'slicerInsert') });
  registerCommand({ id: 'insert.timeline', label: N('Timeline…'), group: N('Insert'), run: (s) => openDialog(s, 'timelineInsert') });
  registerCommand({ id: 'slicer.settings', label: N('Slicer settings…'), group: N('Slicer'), enabled: () => !!pui.slicer, run: (s) => openDialog(s, 'slicerSettings', { name: pui.slicer }) });
  registerCommand({ id: 'slicer.delete', label: N('Delete slicer'), group: N('Slicer'), enabled: () => !!pui.slicer, run: (s) => { if (pui.slicer) { deleteSlicer(s, pui.slicer); setPui(s, { slicer: null }); } } });

  // Sorting inside a table sorts the field rather than the cells.
  for (const [id, dir] of [['data.sortAsc', 'ascending'], ['data.sortDesc', 'descending']] as const) {
    const orig = command(id);
    if (!orig) continue;
    registerCommand({
      ...orig,
      run: (s, arg) => {
        const f = focusOf(s);
        if (!f || !activePivot(s)) { orig.run(s, arg); return; }
        if (f.cell?.role === 'data' && f.cell.data !== undefined) {
          const rl = f.layout.rowLines[f.cell.rowLine!];
          const field = rl?.path.filter((p) => p.field >= 0).at(-1)?.field ?? f.ref.pt.rows.find((x) => x >= 0);
          if (field !== undefined) { sortField(s, f.ref, field, dir, f.cell.data); return; }
        }
        const field = focusField(f);
        if (field === null) { s.notify(T('Select a label or value in the PivotTable to sort.')); return; }
        sortField(s, f.ref, field, dir);
      },
    });
  }

  // ---------- Outline ----------
  const groupRun = (delta: 1 | -1) => (s: GridStore) => {
    if (activePivot(s)) { if (delta > 0) command('pivot.group')!.run(s); else command('pivot.ungroup')!.run(s); return; }
    const t = groupTarget(s);
    groupLines(s, t.axis, t.from, t.to, delta);
  };
  registerCommand({ id: 'cells.group', label: N('Group'), group: N('Outline'), keys: 'Alt+Shift+ArrowRight', run: groupRun(1) });
  registerCommand({ id: 'cells.ungroup', label: N('Ungroup'), group: N('Outline'), keys: 'Alt+Shift+ArrowLeft', run: groupRun(-1) });
  registerCommand({ id: 'outline.auto', label: N('Auto outline'), group: N('Outline'), run: (s) => { autoOutline(s); } });
  registerCommand({ id: 'outline.clear', label: N('Clear outline'), group: N('Outline'), run: clearOutline });
  registerCommand({ id: 'outline.showDetail', label: N('Show detail'), group: N('Outline'), run: (s) => { const f = focusOf(s); if (activePivot(s) && f?.cell?.button && f.cell.field !== undefined && f.cell.slot !== undefined) toggleItem(s, f.ref, f.cell.field, f.cell.slot); else showDetail(s, true); } });
  registerCommand({ id: 'outline.hideDetail', label: N('Hide detail'), group: N('Outline'), run: (s) => { const f = focusOf(s); if (activePivot(s) && f?.cell?.button && f.cell.field !== undefined && f.cell.slot !== undefined) toggleItem(s, f.ref, f.cell.field, f.cell.slot); else showDetail(s, false); } });
  registerCommand({ id: 'outline.settings', label: N('Outline settings…'), group: N('Outline'), run: (s) => openDialog(s, 'outlineSettings') });
  registerCommand({ id: 'data.subtotal', label: N('Subtotal…'), group: N('Data'), run: (s) => openDialog(s, 'subtotalOutline') });
  registerCommand({ id: 'data.consolidate', label: N('Consolidate…'), group: N('Data'), run: (s) => openDialog(s, 'consolidate') });
  registerCommand({ id: 'edit.selectVisible', label: N('Select visible cells'), group: N('Edit'), keys: 'Alt+;', run: selectVisible });

  // ---------- Dialogs ----------
  registerDialog('pivotValueField', ValueFieldDialog);
  registerDialog('pivotFieldSettings', FieldSettingsDialog);
  registerDialog('pivotOptions', PivotOptionsDialog);
  registerDialog('pivotGroup', GroupDialog);
  registerDialog('pivotCalcField', CalcFieldDialog);
  registerDialog('pivotCalcItem', CalcItemDialog);
  registerDialog('pivotFilter', FilterDialog);
  registerDialog('pivotSource', ChangeSourceDialog);
  registerDialog('pivotShowPages', ShowPagesDialog);
  registerDialog('pivotMove', MovePivotDialog);
  registerDialog('slicerInsert', InsertSlicerDialog);
  registerDialog('timelineInsert', InsertTimelineDialog);
  registerDialog('slicerSettings', SlicerSettingsDialog);
  registerDialog('subtotalOutline', SubtotalDialog);
  registerDialog('consolidate', ConsolidateDialog);
  registerDialog('outlineSettings', OutlineSettingsDialog);

  // ---------- Panels ----------
  registerPanelSection({ tab: 'data', title: N('PivotTable'), order: 20, render: pivotSection });
  registerPanelSection({ tab: 'data', title: N('PivotTable layout'), order: 21, render: pivotLayoutSection });
  registerPanelSection({ tab: 'data', title: N('Group and calculate'), order: 22, render: pivotCalcSection });
  registerPanelSection({ tab: 'data', title: N('PivotTable style'), order: 23, render: pivotStyleSection });
  registerPanelSection({ tab: 'data', title: N('Slicers and timelines'), order: 24, render: slicerSection });
  registerPanelSection({ tab: 'data', title: N('Outline and subtotals'), order: 25, render: outlineSection });
  registerPanelSection({ tab: 'insert', title: N('PivotTables'), order: 5, render: insertPivotSection });

  // ---------- Overlays ----------
  registerOverlay({ id: 'pivot.cells', order: 15, render: (s, vp) => { queueMicrotask(() => refreshOnOpen(s)); return createElement(CellButtons, { store: s, vp }); } });
  registerOverlay({ id: 'pivot.slicers', order: 16, render: (s, vp) => createElement(SlicerLayer, { store: s, vp }) });
  registerOverlay({ id: 'pivot.chartButtons', order: 17, render: (s, vp) => createElement(PivotChartButtons, { store: s, vp }) });
  registerOverlay({ id: 'pivot.fieldList', order: 45, render: (s, vp) => createElement(FieldList, { store: s, vp }) });

  // ---------- Grid input ----------
  registerGridHook({
    id: 'pivot',
    cellTap: (s) => {
      if (pui.popup || pui.slicer) setPui(s, { popup: null, slicer: null });
      return false;
    },
    cellDoubleTap: (s, h) => {
      const hit = pivotCellAt(s.wb, s.sheetIndex, h.row, h.col);
      if (!hit?.cell) return false;
      const c = hit.cell;
      s.select({ row: h.row, col: h.col });
      if (c.role === 'data' && c.data !== undefined && c.v !== null) return drillDown(s, s.sheetIndex, h.row, h.col) || true;
      if ((c.role === 'rowItem' || c.role === 'colItem') && c.button && c.field !== undefined && c.slot !== undefined) { toggleItem(s, hit.ref, c.field, c.slot); return true; }
      if (c.role === 'dataCaption' || (c.field === -2 && c.slot !== undefined)) { openDialog(s, 'pivotValueField', { sheet: hit.ref.sheet, name: hit.ref.pt.name, index: c.data ?? c.slot ?? 0 }); return true; }
      if (c.field !== undefined && c.field >= 0) { openDialog(s, 'pivotFieldSettings', { sheet: hit.ref.sheet, name: hit.ref.pt.name, field: c.field }); return true; }
      return true;
    },
    refText: (s, sheet, row, col, editSheet) => (pui.gpd ? getPivotDataText(s.wb, sheet, row, col, editSheet) : null),
    key: (s, ev) => {
      if (ev.key === 'Escape' && (pui.popup || pui.slicer)) { setPui(s, { popup: null, slicer: null }); return true; }
      if ((ev.key === 'Delete' || ev.key === 'Backspace') && pui.slicer) { deleteSlicer(s, pui.slicer); setPui(s, { slicer: null }); return true; }
      // Cells of a PivotTable can't be typed over.
      const typing = (ev.key.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) || ev.key === 'Delete' || ev.key === 'Backspace' || ev.key === 'F2';
      if (typing && !s.editing && activePivot(s)) { s.notify(T('You can\'t change this part of a PivotTable. Use the field list instead.')); return true; }
      return false;
    },
  });

  registerMenuProvider({
    id: 'pivot',
    items: (s, kind) => {
      if (kind !== 'cell') return null;
      const f = focusOf(s);
      if (!f || !activePivot(s)) return null;
      const { ref, cache, cell } = f;
      const field = focusField(f);
      const items: MenuItem[] = [{ label: N('Refresh'), action: () => refreshCurrent(s) }];
      if (field !== null) {
        const name = fieldCaption(ref.pt, cache, field);
        items.push({ label: N('Sort A to Z'), action: () => sortField(s, ref, field, 'ascending') }, { label: N('Sort Z to A'), action: () => sortField(s, ref, field, 'descending') });
        items.push({ label: T('Clear filter from “{name}”', { name }), action: () => clearFieldFilters(s, ref, field) });
        const subs = ref.pt.fields[field]?.subtotals;
        if (ref.pt.rows.includes(field) || ref.pt.cols.includes(field)) items.push({ label: subs?.length === 0 ? T('Show subtotal for “{name}”', { name }) : T('Hide subtotal for “{name}”', { name }), action: () => setSubtotals(s, ref, field, subs?.length === 0 ? ['default'] : []) });
        if (cell?.button && cell.field !== undefined && cell.slot !== undefined) items.push({ label: cell.button === 'collapsed' ? N('Expand') : N('Collapse'), action: () => toggleItem(s, ref, cell.field!, cell.slot!) });
        items.push({ label: N('Expand entire field'), action: () => expandField(s, ref, field, false) }, { label: N('Collapse entire field'), action: () => expandField(s, ref, field, true) });
        if ((cell?.role === 'rowItem' || cell?.role === 'colItem') && !cell.total && cell.slot !== undefined && cell.field === field) {
          // Manual order: move the item among the items shown.
          const lines = cell.role === 'rowItem' ? f.layout.rowLines : f.layout.colLines;
          const shown = [...new Set(lines.flatMap((l) => l.path.filter((p) => p.field === field).map((p) => p.slot)))];
          const at = shown.indexOf(cell.slot);
          const label = String(cell.v ?? '');
          const mv = (to: number) => moveItem(s, ref, field, cell.slot!, to);
          if (at > 0) items.push({ label: T('Move “{label}” up', { label }), action: () => mv(shown[at - 1]!) }, { label: T('Move “{label}” to the beginning', { label }), action: () => mv(Math.min(...shown)) });
          if (at >= 0 && at < shown.length - 1) items.push({ label: T('Move “{label}” down', { label }), action: () => mv(shown[at + 1]!) }, { label: T('Move “{label}” to the end', { label }), action: () => mv(Math.max(...shown)) });
        }
        items.push({ label: N('Group…'), action: () => command('pivot.group')!.run(s) }, { label: N('Ungroup'), action: () => { const e = ungroup(s); if (e) s.notify(e); } });
        items.push({ label: T('Remove “{name}”', { name }), action: () => dropField(s, ref, field) });
        items.push({ label: N('Field settings…'), action: () => openDialog(s, 'pivotFieldSettings', { sheet: ref.sheet, name: ref.pt.name, field }) });
      }
      if (cell?.role === 'data' && cell.data !== undefined) {
        const d = cell.data;
        items.push({ label: N('Show details'), action: () => drillDown(s, s.sheetIndex, s.sel.active.row, s.sel.active.col) });
        items.push({ label: T('Remove “{name}”', { name: dataFieldName(cache, ref.pt.data[d]!) }), action: () => dropField(s, ref, ref.pt.data[d]!.field, 'data', d) });
        items.push({ label: N('Value field settings…'), action: () => openDialog(s, 'pivotValueField', { sheet: ref.sheet, name: ref.pt.name, index: d }) });
      }
      items.push({ label: N('PivotTable options…'), action: () => openDialog(s, 'pivotOptions', { sheet: ref.sheet, name: ref.pt.name }) });
      items.push({ label: pui.pane ? N('Hide field list') : N('Show field list'), action: () => setPui(s, { pane: !pui.pane }) });
      return items;
    },
  });
};

