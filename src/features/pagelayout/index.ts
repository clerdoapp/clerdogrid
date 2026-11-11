// Page layout, printing and PDF export.
import { createElement } from 'react';
import { openDialog, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { registerOverlay, registerPanelSection, registerStageView } from '../../ui/panelRegistry';
import { registerDialog } from '../../ui/dialogs/extra';
import { PageSetupDialog } from './PageSetupDialog';
import { PrintPane } from './PrintPane';
import { BreakLayer, PageLayoutView } from './views';
import { addToPrintArea, clearPrintArea, insertBreak, removeBreak, resetBreaks, setPageView, setPrintArea, setupOf, updateSetup } from './setup';
import { areaSection, breaksSection, previewSection, printSection, setupSection, sheetOptionsSection, viewsSection } from './panels';
import './pagelayout.css';
import { N } from '../../ui/i18n';

const view = (s: GridStore) => s.sheet.view.pageView ?? 'normal';

export const register = (): void => {
  registerCommand({ id: 'page.setup', label: N('Page setup…'), group: N('Page'), run: (s, arg) => openDialog(s, 'pageSetup', arg) });
  registerCommand({ id: 'page.portrait', label: N('Portrait'), group: N('Page'), active: (s) => setupOf(s).orientation !== 'landscape', run: (s) => updateSetup(s, N('Orientation'), { orientation: 'portrait' }) });
  registerCommand({ id: 'page.landscape', label: N('Landscape'), group: N('Page'), active: (s) => setupOf(s).orientation === 'landscape', run: (s) => updateSetup(s, N('Orientation'), { orientation: 'landscape' }) });
  registerCommand({ id: 'page.printAreaSet', label: N('Set print area'), group: N('Page'), run: (s) => setPrintArea(s) });
  registerCommand({ id: 'page.printAreaAdd', label: N('Add to print area'), group: N('Page'), enabled: (s) => !!setupOf(s).printArea, run: (s) => addToPrintArea(s) });
  registerCommand({ id: 'page.printAreaClear', label: N('Clear print area'), group: N('Page'), enabled: (s) => !!setupOf(s).printArea, run: clearPrintArea });
  registerCommand({ id: 'page.titles', label: N('Print titles…'), group: N('Page'), run: (s) => openDialog(s, 'pageSetup', { tab: 'sheet' }) });
  registerCommand({ id: 'page.headerFooter', label: N('Header and footer…'), group: N('Page'), run: (s) => openDialog(s, 'pageSetup', { tab: 'hf' }) });
  registerCommand({ id: 'page.breakInsert', label: N('Insert page break'), group: N('Page'), run: insertBreak });
  registerCommand({ id: 'page.breakRemove', label: N('Remove page break'), group: N('Page'), enabled: (s) => !!(setupOf(s).rowBreaks?.length || setupOf(s).colBreaks?.length), run: removeBreak });
  registerCommand({ id: 'page.breakReset', label: N('Reset all page breaks'), group: N('Page'), enabled: (s) => !!(setupOf(s).rowBreaks?.length || setupOf(s).colBreaks?.length), run: resetBreaks });
  registerCommand({ id: 'page.printGridlines', label: N('Print gridlines'), group: N('Page'), active: (s) => !!setupOf(s).gridlines, run: (s) => updateSetup(s, N('Print gridlines'), { gridlines: !setupOf(s).gridlines }) });
  registerCommand({ id: 'page.printHeadings', label: N('Print headings'), group: N('Page'), active: (s) => !!setupOf(s).headings, run: (s) => updateSetup(s, N('Print headings'), { headings: !setupOf(s).headings }) });
  registerCommand({ id: 'page.viewNormal', label: N('Normal view'), group: N('View'), active: (s) => view(s) === 'normal', run: (s) => setPageView(s, 'normal') });
  registerCommand({ id: 'page.viewBreaks', label: N('Page break preview'), group: N('View'), active: (s) => view(s) === 'pageBreakPreview', run: (s) => setPageView(s, 'pageBreakPreview') });
  registerCommand({ id: 'page.viewLayout', label: N('Page layout view'), group: N('View'), active: (s) => view(s) === 'pageLayout', run: (s) => setPageView(s, 'pageLayout') });
  // Replaces the old print-the-table command; same id and shortcut.
  registerCommand({ id: 'file.exportPdf', label: N('Print…'), group: N('File'), keys: 'Mod+P', run: (s) => openDialog(s, 'printPreview') });
  registerCommand({ id: 'file.print', label: N('Print preview'), group: N('File'), run: (s) => openDialog(s, 'printPreview') });
  registerCommand({ id: 'file.savePdf', label: N('Save as PDF…'), group: N('File'), run: (s, arg) => void import('./browser').then((m) => m.savePdf(s, (arg as never) ?? { what: 'sheet' })) });
  registerCommand({ id: 'file.openPdf', label: N('Open as PDF'), group: N('File'), run: (s) => void import('./browser').then((m) => m.openPdf(s, { what: 'sheet' })) });

  registerDialog('pageSetup', PageSetupDialog);
  registerDialog('printPreview', PrintPane);

  registerPanelSection({ tab: 'page', title: N('Page setup'), order: 1, render: setupSection });
  registerPanelSection({ tab: 'page', title: N('Print area and titles'), order: 2, render: areaSection });
  registerPanelSection({ tab: 'page', title: N('Page breaks'), order: 3, render: breaksSection });
  registerPanelSection({ tab: 'page', title: N('Sheet options'), order: 4, render: sheetOptionsSection });
  registerPanelSection({ tab: 'page', title: N('Views'), order: 5, render: viewsSection });
  registerPanelSection({ tab: 'page', title: N('Print'), order: 6, render: previewSection });
  registerPanelSection({ tab: 'file', title: N('Print'), order: 1, render: printSection });
  registerPanelSection({ tab: 'view', title: N('Workbook views'), order: 1, render: viewsSection });

  registerOverlay({ id: 'pagelayout.breaks', order: 5, render: (s, vp) => createElement(BreakLayer, { store: s, vp }) });
  registerStageView({ id: 'pagelayout.layout', active: (s) => s.sheet.view.pageView === 'pageLayout', render: (s) => createElement(PageLayoutView, { store: s }) });
};
