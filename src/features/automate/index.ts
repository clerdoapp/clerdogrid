import { createElement } from 'react';
import { registerCommand, openDialog } from '../../app/commands';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerPanelSection } from '../../ui/panelRegistry';
import { ScriptEditor, type EditorData } from './ScriptEditor';
import { InsightsDialog } from './InsightsDialog';
import { AutomatePanel, InsightsPanel } from './panels';
import './automate.css';
import { N } from '../../ui/i18n';

export const register = (): void => {
  registerDialog<EditorData>('scripts', ScriptEditor);
  registerDialog('insights', InsightsDialog);
  registerCommand({ id: 'automate.scripts', label: N('Scripts…'), group: N('Automate'), run: (s, arg) => openDialog(s, 'scripts', arg) });
  registerCommand({ id: 'automate.newScript', label: N('New script'), group: N('Automate'), run: (s) => openDialog(s, 'scripts', { tab: 'book' }) });
  registerCommand({ id: 'automate.samples', label: N('Sample scripts'), group: N('Automate'), run: (s) => openDialog(s, 'scripts', { tab: 'sample' }) });
  registerCommand({ id: 'data.analyze', label: N('Analyze data'), group: N('Data'), run: (s) => openDialog(s, 'insights') });
  registerPanelSection({ tab: 'automate', title: N('Scripts'), order: 0, render: (store) => createElement(AutomatePanel, { store }) });
  registerPanelSection({ tab: 'data', title: N('Insights'), order: 10, render: (store) => createElement(InsightsPanel, { store }) });
};
