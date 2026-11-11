import { createElement } from 'react';
import { registerOverlay, registerPanelSection } from '../../ui/panelRegistry';
import { OnlinePanel } from './OnlinePanel';
import { bindStore, startOnline } from './state';
import { N } from '../../ui/i18n';

export const register = (): void => {
  startOnline();
  registerPanelSection({ tab: 'file', title: N('Online functions'), order: 41, render: (store) => { bindStore(store); return createElement(OnlinePanel, { store }); } });
  // The grid's overlay pass is the first place a store is available; it wires recalculation on answers.
  registerOverlay({ id: 'online-bridge', order: -100, render: (store) => { bindStore(store); return null; } });
};
