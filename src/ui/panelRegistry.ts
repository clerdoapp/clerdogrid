import type { ReactNode } from 'react';
import type { GridStore } from '../app/store';

/** Tool-panel sections contributed by feature modules, shown under a rail tab after its built-in groups. */
export interface PanelSection { tab: string; title: string; order?: number; render: (store: GridStore) => ReactNode }

const sections: PanelSection[] = [];
export const registerPanelSection = (s: PanelSection): void => { sections.push(s); };
export const panelSections = (tab: string): PanelSection[] => sections.filter((s) => s.tab === tab).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

/** Things drawn over the grid by feature modules (shapes, ink, comment markers). */
export interface OverlayLayer { id: string; order?: number; render: (store: GridStore, vp: import('./render').Viewport) => ReactNode }
const layers: OverlayLayer[] = [];
export const registerOverlay = (l: OverlayLayer): void => { if (!layers.some((x) => x.id === l.id)) layers.push(l); };
export const overlays = (): OverlayLayer[] => [...layers].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
