// More file types: save as .ods/.ots/.xls/.xlsb, export web page / XML / JSON, text import, and Get Data queries.
import { openDialog, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { beforeSaveHooks, saveFile, workbookFromBytes } from '../../app/files';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerPanelSection, registerStartup } from '../../ui/panelRegistry';
import { xlsLimits } from '../../formats/biff8-write';
import { xlsbLimits } from '../../formats/xlsb-write';
import { ExportDialog, TextImportDialog, WorkbookImportDialog, XlsbLimitsDialog, XlsLimitsDialog, XlsPasswordDialog, installOpenHandlers } from './dialogs';
import { exportDataSection, getDataSection, saveTypesSection } from './panels';
import { loadQueries, pickFile, refreshAll, stashQueries } from './ops';
import './formats.css';
import { N, T } from '../../ui/i18n';

const saveAsXls = (s: GridStore): void => {
  if (xlsLimits(s.wb).length) openDialog(s, 'formats.xlsLimits');
  else void saveFile(s, true, 'xls');
};

const saveAsXlsb = (s: GridStore): void => {
  if (xlsbLimits(s.wb).length) openDialog(s, 'formats.xlsbLimits');
  else void saveFile(s, true, 'xlsb');
};

const fromText = async (s: GridStore): Promise<void> => {
  const file = await pickFile(true);
  if (file) openDialog(s, 'formats.textImport', { file, mode: 'insert' });
};

const fromWorkbook = async (s: GridStore): Promise<void> => {
  const file = await pickFile(false);
  if (!file) return;
  try {
    const wb = workbookFromBytes(file.name, file.data);
    loadQueries(wb);
    openDialog(s, 'formats.fromWorkbook', { file, wb });
  } catch (e) {
    s.notify(e instanceof Error ? e.message : T('The file could not be opened.'));
  }
};

export const register = (): void => {
  registerCommand({ id: 'file.saveAsXlsx', label: N('Save as workbook (.xlsx)'), group: N('File'), run: (s) => void saveFile(s, true, 'xlsx') });
  registerCommand({ id: 'file.saveAsOds', label: N('Save as OpenDocument (.ods)'), group: N('File'), run: (s) => void saveFile(s, true, 'ods') });
  registerCommand({ id: 'file.saveAsOts', label: N('Save as OpenDocument template (.ots)'), group: N('File'), run: (s) => void saveFile(s, true, 'ots') });
  registerCommand({ id: 'file.saveAsXls', label: N('Save as older workbook (.xls)'), group: N('File'), run: saveAsXls });
  registerCommand({ id: 'file.saveAsXlsb', label: N('Save as binary workbook (.xlsb)'), group: N('File'), run: saveAsXlsb });
  registerCommand({ id: 'file.exportHtml', label: N('Export as web page'), group: N('File'), run: (s) => openDialog(s, 'formats.export', { kind: 'html' }) });
  registerCommand({ id: 'file.exportXml', label: N('Export as XML data'), group: N('File'), run: (s) => openDialog(s, 'formats.export', { kind: 'xml' }) });
  registerCommand({ id: 'file.exportJson', label: N('Export as JSON data'), group: N('File'), run: (s) => openDialog(s, 'formats.export', { kind: 'json' }) });
  registerCommand({ id: 'data.fromText', label: N('From text or CSV'), group: N('Get data'), run: (s) => void fromText(s) });
  // the built-in Get data button uses the full text import with a refreshable query
  registerCommand({ id: 'data.importCsv', label: N('From text/CSV…'), group: N('Data'), run: (s) => void fromText(s) });
  registerCommand({ id: 'data.fromWorkbook', label: N('From workbook'), group: N('Get data'), run: (s) => void fromWorkbook(s) });
  registerCommand({ id: 'data.refreshAll', label: N('Refresh all'), group: N('Get data'), keys: 'Mod+Alt+F5', enabled: (s) => !!s.wb.queries?.length, run: (s) => void refreshAll(s) });

  registerDialog('formats.textImport', TextImportDialog);
  registerDialog('formats.fromWorkbook', WorkbookImportDialog);
  registerDialog('formats.password', XlsPasswordDialog);
  registerDialog('formats.xlsLimits', XlsLimitsDialog);
  registerDialog('formats.xlsbLimits', XlsbLimitsDialog);
  registerDialog('formats.export', ExportDialog);

  registerPanelSection({ tab: 'file', title: N('Save as type'), order: 5, render: saveTypesSection });
  registerPanelSection({ tab: 'file', title: N('Export data'), order: 6, render: exportDataSection });
  registerPanelSection({ tab: 'data', title: N('Import and refresh'), order: 1, render: getDataSection });

  beforeSaveHooks.push(stashQueries);
  installOpenHandlers(openDialog);

  registerStartup((store) => {
    let last = store.wb;
    loadQueries(store.wb);
    store.subscribe(() => {
      if (store.wb === last) return;
      last = store.wb;
      loadQueries(store.wb);
    });
  });
};
