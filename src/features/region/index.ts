import { createElement } from 'react';
import { registerPanelSection } from '../../ui/panelRegistry';
import { initLocale, setSeparators } from '../../app/locale';
import { RegionPanel } from './RegionPanel';
import './region.css';
import { N } from '../../ui/i18n';

/** Separator override from File > Options, when that setting is saved on this device. */
const optionSeparators = (): void => {
  try {
    const raw = localStorage.getItem('clerdogrid.options');
    if (!raw) return;
    const o = JSON.parse(raw) as { decimalSep?: string; thousandsSep?: string };
    if (o.decimalSep || o.thousandsSep) setSeparators(o.decimalSep ?? '', o.thousandsSep ?? '');
  } catch { /* no storage */ }
};

export const register = (): void => {
  initLocale();
  optionSeparators();
  registerPanelSection({ tab: 'file', title: N('Region'), order: 40, render: (store) => createElement(RegionPanel, { store }) });
};
