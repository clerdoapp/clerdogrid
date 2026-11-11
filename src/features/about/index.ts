import { createElement } from 'react';
import { openDialog, registerCommand } from '../../app/commands';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerPanelSection } from '../../ui/panelRegistry';
import { AboutDialog } from './AboutDialog';
import { AboutPanel } from './AboutPanel';
import './about.css';
import { N } from '../../ui/i18n';

export const register = (): void => {
  registerDialog('about.full', AboutDialog);
  // Replaces the short built-in About with the full one (contributors, links, notices).
  registerCommand({ id: 'help.about', label: N('About ClerdoGrid'), group: N('Help'), run: (s) => openDialog(s, 'about.full') });
  registerPanelSection({ tab: 'file', title: N('About ClerdoGrid'), order: 100, render: (store) => createElement(AboutPanel, { store }) });
};
