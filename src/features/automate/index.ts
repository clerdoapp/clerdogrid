import { createElement } from 'react';
import { registerCommand, openDialog } from '../../app/commands';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerPanelSection } from '../../ui/panelRegistry';
import { ScriptEditor, type EditorData } from './ScriptEditor';
import { InsightsDialog } from './InsightsDialog';
import { AutomatePanel, InsightsPanel } from './panels';
import './automate.css';

export const register = (): void => {
  registerDialog<EditorData>('scripts', ScriptEditor);
  registerDialog('insights', InsightsDialog);
  registerCommand({ id: 'automate.scripts', label: 'Scripts…', group: 'Automate', run: (s, arg) => openDialog(s, 'scripts', arg) });
  registerCommand({ id: 'automate.newScript', label: 'New script', group: 'Automate', run: (s) => openDialog(s, 'scripts', { tab: 'book' }) });
  registerCommand({ id: 'automate.samples', label: 'Sample scripts', group: 'Automate', run: (s) => openDialog(s, 'scripts', { tab: 'sample' }) });
  registerCommand({ id: 'data.analyze', label: 'Analyze data', group: 'Data', run: (s) => openDialog(s, 'insights') });
  registerPanelSection({ tab: 'automate', title: 'Scripts', order: 0, render: (store) => createElement(AutomatePanel, { store }) });
  registerPanelSection({ tab: 'data', title: 'Insights', order: 10, render: (store) => createElement(InsightsPanel, { store }) });
};
