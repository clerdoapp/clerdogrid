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

/** Grid input hooks from feature modules (checkbox clicks, comment popovers, keys). Returning true stops the grid's own handling. */
export interface GridHit { row: number; col: number; x: number; y: number; pointerType: string }
export interface GridHook {
  id: string;
  cellTap?: (store: GridStore, hit: GridHit) => boolean;
  hover?: (store: GridStore, hit: GridHit | null) => void;
  key?: (store: GridStore, ev: KeyboardEvent) => boolean;
}
const hooks: GridHook[] = [];
export const registerGridHook = (h: GridHook): void => { if (!hooks.some((x) => x.id === h.id)) hooks.push(h); };
export const gridHooks = (): readonly GridHook[] => hooks;
