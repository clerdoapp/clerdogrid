import { createElement } from 'react';
import { registerPanelSection, registerStartup, registerTopBarItem } from '../../ui/panelRegistry';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerCommand } from '../../app/commands';
import { saveHandlers } from '../../app/files';
import { N } from '../../ui/i18n';
import { CloudPanel, ConflictDialog, SaveChip } from './ui';
import { cloudEnabled, cloudSave, postToKelvo, saveToCloud, startCloud } from './state';
import './cloud.css';

export const register = (): void => {
  if (!cloudEnabled()) return;
  saveHandlers.cloud = cloudSave;
  registerStartup(async (store) => {
    const opened = await startCloud(store);
    postToKelvo({ type: 'ready' });
    return opened;
  });
  registerPanelSection({ tab: 'file', title: N('Clerdo cloud'), order: 5, render: (store) => createElement(CloudPanel, { store }) });
  registerTopBarItem({ id: 'cloud.status', order: 5, render: (store) => createElement(SaveChip, { store }) });
  registerDialog('cloud.conflict', ConflictDialog);
  registerCommand({ id: 'cloud.save', label: N('Save to the cloud'), group: N('File'), run: (s) => void saveToCloud(s) });
};
