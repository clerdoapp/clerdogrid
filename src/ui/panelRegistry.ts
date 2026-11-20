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
  /** Double-click or double-tap on a cell. */
  cellDoubleTap?: (store: GridStore, hit: GridHit) => boolean;
  /** Text to insert instead of a plain reference when a cell is clicked while typing a formula. */
  refText?: (store: GridStore, sheet: number, row: number, col: number, editSheet: number) => string | null;
}
const hooks: GridHook[] = [];
export const registerGridHook = (h: GridHook): void => { if (!hooks.some((x) => x.id === h.id)) hooks.push(h); };
export const gridHooks = (): readonly GridHook[] => hooks;

/** Small items shown in the top bar by feature modules (for example who is in a live session). */
export interface TopBarItem { id: string; order?: number; render: (store: GridStore) => ReactNode }
const topItems: TopBarItem[] = [];
export const registerTopBarItem = (t: TopBarItem): void => { if (!topItems.some((x) => x.id === t.id)) topItems.push(t); };
export const topBarItems = (): TopBarItem[] => [...topItems].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

/**
 * Start-up hooks from feature modules, run once the store exists. A hook that opens a workbook itself
 * (for example a second window) returns true so the launch file and saved draft are skipped.
 */
export type StartupHook = (store: GridStore) => boolean | Promise<boolean> | void;
const startups: StartupHook[] = [];
export const registerStartup = (h: StartupHook): void => { startups.push(h); };
export const runStartups = async (store: GridStore): Promise<boolean> => {
  let opened = false;
  for (const h of startups) { try { if (await h(store)) opened = true; } catch { /* a failing hook mustn't stop the app */ } }
  return opened;
};

/** Views that cover the grid area (the page layout view). The grid stays mounted underneath. */
export interface StageView { id: string; active: (store: GridStore) => boolean; render: (store: GridStore) => ReactNode }
const stageViews: StageView[] = [];
export const registerStageView = (v: StageView): void => { if (!stageViews.some((x) => x.id === v.id)) stageViews.push(v); };
export const activeStageView = (store: GridStore): StageView | undefined => stageViews.find((v) => v.active(store));

/** Extra context-menu items from feature modules, shown first in the cell menu. */
export type MenuItem = { label: string; action: () => void; disabled?: boolean } | 'sep';
export interface MenuProvider { id: string; items: (store: GridStore, kind: 'cell' | 'row' | 'col') => MenuItem[] | null }
const menuProviders: MenuProvider[] = [];
export const registerMenuProvider = (p: MenuProvider): void => { if (!menuProviders.some((x) => x.id === p.id)) menuProviders.push(p); };
export const menuProviderItems = (store: GridStore, kind: 'cell' | 'row' | 'col'): MenuItem[] => menuProviders.flatMap((p) => { const it = p.items(store, kind); return it?.length ? [...it, 'sep' as const] : []; });

/**
 * Who is signed in, for the account menu in the top bar. Only the online build (grid.clerdo.app) registers one;
 * the desktop and offline builds have no account.
 */
export interface AccountSource { useAccount: () => { user: import('@clerdo/frame').ClerdoUser | null; signIn: () => void; signOut: () => void } }
const noAccount: AccountSource = { useAccount: () => ({ user: null, signIn: () => undefined, signOut: () => undefined }) };
let accountSource: AccountSource = noAccount;
let accountRegistered = false;
export const registerAccountSource = (a: AccountSource): void => { accountSource = a; accountRegistered = true; };
export const useAccountSource = () => accountSource.useAccount();
export const hasAccountSource = (): boolean => accountRegistered;
