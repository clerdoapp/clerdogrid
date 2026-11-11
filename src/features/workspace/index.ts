// Workspace: version history, view tools (focus cell, navigation, full screen, sheet and custom views,
// split windows, new windows), workbook protection and encryption, and File > Options.
import { createElement } from 'react';
import { openDialog, registerCommand, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { beforeSaveHooks, desktop, encryptedOpenHandler } from '../../app/files';
import { flashFill } from '../../app/ops/fill';
import { flashPreview } from '../../app/ops/flash';
import { newWorkbook, type Workbook } from '../../model/workbook';
import { registerGridHook, registerOverlay, registerPanelSection, registerStartup } from '../../ui/panelRegistry';
import { registerDialog } from '../../ui/dialogs/extra';
import { statusPrefs } from '../../ui/StatusBar';
import { activeGrid } from '../../ui/GridView';
import { setZoom } from '../../app/commands';
import { settings, saveSettings } from './settings';
import { setUi, ui } from './ui';
import { desktopWindow } from './desktop';
import { closePreview, consumeHandoff, preview, startVersionTimer } from './history';
import { detachMirror, mirrorState, openMirrorWindow, startMirrorIfAsked, startMirrorSource } from './mirror';
import { activeViewOf, carryViews, createView, resetViews, settleViewsForSave, sortInView, switchView, toggleFullScreen, toggleSplit, updateRowOrder } from './views';
import { sortInterceptor } from '../../app/ops/data';
import { installAutoComplete } from './autocomplete';
import { applyDeviceSettings, applyNewWorkbookDefaults, applyWorkbookSettings, OptionsDialog } from './options';
import { blocked, EncryptDialog, guardCommands, OpenPasswordDialog, ProtectSheetDialog, ProtectWorkbookDialog } from './protect';
import { CustomViewsDialog, SheetViewsDialog, StatusMenu } from './viewDialogs';
import { ErrorLayer, FocusLayer, PreviewBanner, SplitBars } from './layers';
import { NavPane, VersionsPane } from './panes';
import { calcSection, fileSection, focusSection, protectSection, sheetViewSection, showSection, windowSection } from './panels';
import { MirrorBanner } from './MirrorBanner';
import './workspace.css';
import { N, T } from '../../ui/i18n';

const fitSelection = (s: GridStore): void => {
  const vp = activeGrid.current?.viewport?.();
  if (!vp) { openDialog(s, 'zoomToSelection'); return; }
  const r = s.range;
  const w = vp.layout.cols.start(Math.min(r.c2 + 1, r.c1 + 2000)) - vp.layout.cols.start(r.c1);
  const h = vp.layout.rows.start(Math.min(r.r2 + 1, r.r1 + 20000)) - vp.layout.rows.start(r.r1);
  const availW = vp.width - vp.headerW - 24, availH = vp.height - vp.headerH - 24;
  const z = Math.floor(Math.min(availW / Math.max(1, w), availH / Math.max(1, h)) * 100);
  setZoom(s, Math.max(10, Math.min(400, z)));
  requestAnimationFrame(() => activeGrid.current?.scrollToPx?.(vp.layout.cols.start(r.c1) - vp.layout.cols.start(s.sheet.view.frozenCols), vp.layout.rows.start(r.r1) - vp.layout.rows.start(s.sheet.view.frozenRows)));
};

const autoFlashFill = (s: GridStore): void => {
  // With the grey suggestion preview (features/flashfill) the person accepts with Enter instead.
  if (!settings.flashFillAuto || flashPreview.enabled) return;
  const { row, col } = s.sel.active;
  // The entry just typed is the cell above the active one (Enter moved down).
  const r = row - 1;
  if (r < 1 || col < 1) return;
  const v = s.value(r, col);
  if (typeof v !== 'string' || !v || s.cell(r, col)?.f) return;
  // Needs two or more typed examples in a row at the top of a column next to data, with blanks below.
  let top = r;
  while (top > 0 && typeof s.value(top - 1, col) === 'string' && s.value(top - 1, col - 1) !== null) top--;
  if (r - top < 1 || s.value(r + 1, col) !== null || s.value(r + 1, col - 1) === null) return;
  if (top > 0 && s.value(top - 1, col - 1) !== null && s.value(top - 1, col) === null) return;
  const notify = s.notify;
  let filled = false;
  s.notify = () => undefined;
  try {
    const keep = s.sel;
    s.sel = { ranges: [{ r1: r, c1: col, r2: r, c2: col }], active: { row: r, col }, anchor: { row: r, col } };
    filled = flashFill(s);
    s.sel = keep;
  } finally { s.notify = notify; }
  if (filled) s.notify(T('Flash Fill filled in the rest of the column. Undo to remove it.'));
};

const newBook = (): Workbook => { const wb = newWorkbook(); applyNewWorkbookDefaults(wb); return wb; };

export const register = (): void => {
  // ---------- View ----------
  registerCommand({ id: 'view.focusCell', label: N('Focus cell'), group: N('View'), active: () => ui.focusCell, run: (s) => setUi(s, { focusCell: !ui.focusCell }) });
  registerCommand({ id: 'view.navPane', label: N('Navigation'), group: N('View'), active: () => ui.navPane, run: (s) => setUi(s, { navPane: !ui.navPane }) });
  registerCommand({ id: 'view.fullScreen', label: N('Full screen'), group: N('View'), keys: 'Mod+Shift+F1', active: () => ui.fullScreen, run: (s) => void toggleFullScreen(s) });
  registerCommand({ id: 'view.split', label: N('Split window'), group: N('View'), active: (s) => !!s.sheet.view.split, run: toggleSplit });
  registerCommand({ id: 'view.newWindow', label: N('New window'), group: N('View'), run: openMirrorWindow });
  registerCommand({ id: 'view.arrange', label: N('Arrange windows'), group: N('View'), enabled: () => !!desktopWindow?.arrangeWindows, run: () => void desktopWindow?.arrangeWindows?.('tiled') });
  registerCommand({ id: 'view.zoomFit', label: N('Fit selection'), group: N('View'), run: fitSelection });
  registerCommand({ id: 'view.sheetViews', label: N('Sheet views…'), group: N('View'), run: (s) => openDialog(s, 'workspace.sheetViews') });
  registerCommand({ id: 'view.newSheetView', label: N('New sheet view'), group: N('View'), run: (s) => { const v = createView(s); s.notify(T('Now in "{name}". Filters, sorting and hidden rows here don\'t change the default view.', { name: v.name })); } });
  registerCommand({ id: 'view.tempView', label: N('Temporary view'), group: N('View'), active: (s) => activeViewOf(s.sheet) === '', run: (s) => switchView(s, activeViewOf(s.sheet) === '' ? null : '') });
  registerCommand({ id: 'view.defaultView', label: N('Default view'), group: N('View'), run: (s) => switchView(s, null) });
  registerCommand({ id: 'view.customViews', label: N('Custom views…'), group: N('View'), run: (s) => openDialog(s, 'workspace.customViews') });
  registerCommand({ id: 'view.formulaBar', label: N('Formula bar'), group: N('View'), active: () => settings.showFormulaBar, run: (s) => { saveSettings(s, { showFormulaBar: !settings.showFormulaBar }); } });
  registerCommand({ id: 'view.statusBar', label: N('Status bar'), group: N('View'), active: () => settings.showStatusBar, run: (s) => { saveSettings(s, { showStatusBar: !settings.showStatusBar }); } });

  // ---------- Review and protection ----------
  registerCommand({ id: 'sheet.protect', label: N('Protect sheet'), group: N('Review'), active: (s) => !!s.sheet.protection?.sheet, run: (s) => openDialog(s, 'workspace.protectSheet') });
  registerCommand({ id: 'review.protectWorkbook', label: N('Protect workbook'), group: N('Review'), active: (s) => !!s.wb.protection?.structure, run: (s) => openDialog(s, 'workspace.protectWorkbook') });
  registerCommand({ id: 'file.encrypt', label: N('Protect with password'), group: N('File'), active: (s) => !!s.file.password, run: (s) => openDialog(s, 'workspace.encrypt') });

  // ---------- File ----------
  registerCommand({ id: 'file.versions', label: N('Version history'), group: N('File'), active: () => ui.versionsPane, run: (s) => setUi(s, { versionsPane: !ui.versionsPane }) });
  registerCommand({ id: 'file.options', label: N('Options'), group: N('File'), run: (s) => openDialog(s, 'workspace.options') });
  registerCommand({
    id: 'file.new', label: N('New workbook'), group: N('File'), keys: 'Mod+N',
    run: (s) => {
      if (desktop) { desktop.newWindow(); return; }
      if (s.file.dirty && !confirm(T('Start a new workbook? Unsaved changes in this one will be lost.'))) return;
      if (preview.current) closePreview(s, false);
      s.load(newBook(), 'Book1');
    },
  });

  // ---------- Calculation ----------
  registerCommand({ id: 'formula.calcSheet', label: N('Calculate sheet'), group: N('Formulas'), keys: 'Shift+F9', run: (s) => { s.engine.refreshSheet(s.sheetIndex); s.engine.recalc(); s.touch(); } });
  registerCommand({
    id: 'formula.calcManual', label: N('Manual calculation'), group: N('Formulas'), active: (s) => !!s.wb.calc?.manual,
    run: (s) => {
      const manual = !s.wb.calc?.manual;
      s.change(manual ? N('Manual calculation') : N('Automatic calculation'), 'book', () => {
        s.wb.calc = { ...s.wb.calc, manual: manual || undefined, autoNoTable: undefined };
        applyWorkbookSettings(s);
        if (!manual) s.engine.recalc(true);
      });
      s.notify(manual ? T('Calculation is manual. Press F9 to calculate.') : T('Calculation is automatic.'));
    },
  });
  registerCommand({ id: 'formula.calcOptions', label: N('Calculation options…'), group: N('Formulas'), run: (s) => openDialog(s, 'workspace.options', { tab: 'formulas' }) });
  registerCommand({
    id: 'formula.iterative', label: N('Iterative calculation'), group: N('Formulas'), active: (s) => s.engine.iterate,
    run: (s) => s.change(N('Iterative calculation'), 'book', () => { s.wb.calc = { ...s.wb.calc, iterate: !s.engine.iterate || undefined }; applyWorkbookSettings(s); s.engine.recalc(true); }),
  });

  // ---------- Dialogs ----------
  registerDialog('workspace.options', OptionsDialog);
  registerDialog('workspace.protectSheet', ProtectSheetDialog);
  registerDialog('workspace.protectWorkbook', ProtectWorkbookDialog);
  registerDialog('workspace.encrypt', EncryptDialog);
  registerDialog('workspace.openPassword', OpenPasswordDialog);
  registerDialog('workspace.sheetViews', SheetViewsDialog);
  registerDialog('workspace.customViews', CustomViewsDialog);

  // ---------- Panels ----------
  registerPanelSection({ tab: 'view', title: N('Window'), order: 20, render: windowSection });
  registerPanelSection({ tab: 'view', title: N('Sheet view'), order: 21, render: sheetViewSection });
  registerPanelSection({ tab: 'view', title: N('Focus cell'), order: 22, render: focusSection });
  registerPanelSection({ tab: 'view', title: N('Bars'), order: 23, render: showSection });
  registerPanelSection({ tab: 'review', title: N('Protection'), order: 20, render: protectSection });
  registerPanelSection({ tab: 'formulas', title: N('Calculation options'), order: 20, render: calcSection });
  registerPanelSection({ tab: 'file', title: N('Workbook tools'), order: 20, render: fileSection });

  // ---------- Overlays ----------
  registerOverlay({ id: 'workspace.focus', order: 5, render: (s, vp) => createElement(FocusLayer, { store: s, vp }) });
  registerOverlay({ id: 'workspace.errors', order: 25, render: (s, vp) => createElement(ErrorLayer, { store: s, vp }) });
  registerOverlay({ id: 'workspace.split', order: 27, render: (s, vp) => createElement(SplitBars, { store: s, vp }) });
  registerOverlay({ id: 'workspace.nav', order: 41, render: (s, vp) => createElement(NavPane, { store: s, vp }) });
  registerOverlay({ id: 'workspace.versions', order: 42, render: (s, vp) => createElement(VersionsPane, { store: s, vp }) });
  registerOverlay({ id: 'workspace.banner', order: 50, render: (s, vp) => createElement(PreviewBanner, { store: s, vp }) });
  registerOverlay({ id: 'workspace.mirror', order: 51, render: (s, vp) => createElement(MirrorBanner, { store: s, vp }) });
  registerOverlay({ id: 'workspace.statusMenu', order: 60, render: (s) => createElement(StatusMenu, { store: s }) });

  // ---------- Grid input ----------
  registerGridHook({
    id: 'workspace',
    cellTap: (s, h) => {
      if (ui.errorMenu) setUi(s, { errorMenu: false });
      // Protected sheets can stop locked (or all) cells from being selected.
      const sheet = s.sheet;
      if (sheet.protection?.sheet && blocked(sheet, s.isLocked(h.row, h.col) ? 'selectLockedCells' : 'selectUnlockedCells')) return true;
      return false;
    },
    key: (s, ev) => {
      if (ev.key === 'Escape' && preview.current) { closePreview(s); return true; }
      if (ev.key === 'Escape' && ui.errorMenu) { setUi(s, { errorMenu: false }); return true; }
      return false;
    },
  });

  // ---------- Start-up ----------
  registerStartup(async (store) => {
    applyDeviceSettings(store);
    installAutoComplete(store);
    sortInterceptor.current = sortInView;
    // Protected sheets can forbid selecting locked or unlocked cells, by mouse, keyboard, Go To or Find.
    store.selectable = (row, col) => {
      const sh = store.sheet;
      if (!sh.protection?.sheet) return true;
      return !blocked(sh, store.isLocked(row, col) ? 'selectLockedCells' : 'selectUnlockedCells');
    };
    guardCommands();
    statusPrefs.onMenu = (s, x, y) => setUi(s, { statusMenu: { x, y } });
    encryptedOpenHandler.current = (s, file) => { if (preview.current) closePreview(s, false); openDialog(s, 'workspace.openPassword', file); };
    beforeSaveHooks.push((s) => { if (preview.current) throw new Error(T('close the version preview first')); return settleViewsForSave(s); });

    // Workbook settings follow whichever workbook is loaded.
    let lastWb = store.wb;
    let calcSig = JSON.stringify([store.wb.calc ?? null, store.wb.date1904]);
    applyWorkbookSettings(store);
    store.subscribe(() => {
      if (store.wb === lastWb || preview.current?.saved.wb === lastWb) return;
      lastWb = store.wb;
      applyWorkbookSettings(store);
      calcSig = JSON.stringify([store.wb.calc ?? null, store.wb.date1904]);
      if (store.engine.manual) store.engine.stale = false;
      resetViews(store);
    });
    store.observe({ end: (e) => {
      if (e.kind === 'change' && e.label === 'Typing') queueMicrotask(() => autoFlashFill(store));
      if (e.kind !== 'change') carryViews(store, e.sheetsBefore);
      // A view's row order follows its filter range when rows are inserted or deleted.
      const o = store.rowOrderOf(), af = store.sheet.autoFilter;
      if ((o && (!af || o.r1 !== af.range.r1 + 1 || o.r2 !== af.range.r2)) || (!o && activeViewOf(store.sheet) !== null && af?.sort?.length)) updateRowOrder(store, store.sheet);
      // Undo and redo can bring back other calculation settings or the date system.
      const sig = JSON.stringify([store.wb.calc ?? null, store.wb.date1904]);
      if (e.kind !== 'change' && sig !== calcSig) { applyWorkbookSettings(store); store.engine.recalc(true); }
      calcSig = sig;
    } });

    if (typeof document !== 'undefined') {
      // Escape leaves full screen wherever focus is (the browser handles its own full screen mode).
      document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape' || !ui.fullScreen || store.dialog || store.editing || document.fullscreenElement) return;
        e.preventDefault();
        void toggleFullScreen(store);
      });
      document.addEventListener('fullscreenchange', () => {
        const on = !!document.fullscreenElement;
        if (on !== ui.fullScreen && !desktopWindow?.setFullScreen) { document.documentElement.toggleAttribute('data-fullscreen', on); setUi(store, { fullScreen: on }); }
      });
    }

    const handled = startMirrorIfAsked(store) || await consumeHandoff(store);
    if (!handled) {
      applyNewWorkbookDefaults(store.wb);
      store.engine.sheetsChanged();
      store.touch();
    }
    startVersionTimer(store);
    startMirrorSource(store);
    return handled;
  });
};

export { detachMirror, mirrorState, run };
