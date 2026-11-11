// Hooks the chart format panel, chart dialogs, chart sheets and chart commands into the app.
import { createElement } from 'react';
import { registerCommand, openDialog } from '../../app/commands';
import { registerDialog } from '../dialogs/extra';
import { registerGridHook, registerOverlay, registerPanelSection } from '../panelRegistry';
import { activeChart, copyChartPicture, moveToChartSheet, switchRowCol } from './chartOps';
import { MoveChartDialog, SaveTemplateDialog, SelectDataDialog } from './ChartDialogs';
import { ChartFormatPanel } from './FormatPanel';
import { ChartSheetView } from './ChartSheetView';
import { openFormat } from './ChartView';
import { N } from '../i18n';

let done = false;

export const registerCharts = (): void => {
  if (done) return;
  done = true;
  registerPanelSection({ tab: 'chart', title: N('Format chart'), render: (s) => createElement(ChartFormatPanel, { store: s }) });
  registerDialog('chartData', SelectDataDialog);
  registerDialog('chartMove', MoveChartDialog);
  registerDialog('chartTemplate', SaveTemplateDialog);
  registerOverlay({ id: 'chartsheet', order: 50, render: (s, vp) => (s.sheet.kind === 'chart' ? createElement(ChartSheetView, { store: s, vp }) : null) });
  // A chart sheet has no cells to select or type into.
  registerGridHook({
    id: 'chartsheet',
    cellTap: (s) => s.sheet.kind === 'chart',
    key: (s, ev) => s.sheet.kind === 'chart' && !(ev.ctrlKey || ev.metaKey) && (ev.key.length === 1 || ['Delete', 'Backspace', 'Enter', 'F2', 'Tab'].includes(ev.key) || ev.key.startsWith('Arrow')),
  });
  const has = (s: Parameters<typeof activeChart>[0]) => !!activeChart(s);
  registerCommand({ id: 'chart.format', label: N('Format chart'), group: N('Chart'), enabled: has, run: (s) => openFormat(s) });
  registerCommand({ id: 'chart.selectData', label: N('Select chart data…'), group: N('Chart'), enabled: has, run: (s) => openDialog(s, 'chartData', { id: activeChart(s)!.id }) });
  registerCommand({ id: 'chart.switchRowCol', label: N('Switch chart rows and columns'), group: N('Chart'), enabled: has, run: (s) => { switchRowCol(s, activeChart(s)!.id); } });
  registerCommand({ id: 'chart.move', label: N('Move chart…'), group: N('Chart'), enabled: has, run: (s) => openDialog(s, 'chartMove', { id: activeChart(s)!.id }) });
  registerCommand({ id: 'chart.toSheet', label: N('Move chart to its own sheet'), group: N('Chart'), enabled: (s) => has(s) && s.sheet.kind !== 'chart', run: (s) => { const c = activeChart(s)!; moveToChartSheet(s, c.id, c.title && !c.title.startsWith('=') ? c.title.slice(0, 31) : 'Chart'); } });
  registerCommand({ id: 'chart.copyPicture', label: N('Copy chart as picture'), group: N('Chart'), enabled: has, run: (s) => { void copyChartPicture(s, activeChart(s)!.id); } });
  registerCommand({ id: 'chart.saveTemplate', label: N('Save chart as template…'), group: N('Chart'), enabled: has, run: (s) => openDialog(s, 'chartTemplate', { id: activeChart(s)!.id }) });
};
