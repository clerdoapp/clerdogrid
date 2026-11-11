import { createElement } from 'react';
import { registerGridHook, registerOverlay } from '../../ui/panelRegistry';
import type { GridStore } from '../../app/store';
import { GhostLayer } from './GhostLayer';
import { accept, dismiss, ghost, watch } from './ghost';
import { flashPreview } from '../../app/ops/flash';
import './flashfill.css';

const watched = new WeakSet<GridStore>();

export const register = (): void => {
  flashPreview.enabled = true;
  registerOverlay({
    id: 'flashfill-ghost',
    order: 50,
    render: (store, vp) => {
      if (!watched.has(store)) { watched.add(store); watch(store); }
      return createElement(GhostLayer, { store, vp });
    },
  });
  registerGridHook({
    id: 'flashfill-keys',
    key: (store, ev) => {
      if (!ghost.plan || store.editing) return false;
      if (ev.key === 'Enter' && !ev.shiftKey && !ev.ctrlKey && !ev.metaKey && !ev.altKey) { ev.preventDefault(); return accept(store); }
      if (ev.key === 'Escape') { dismiss(store); return true; }
      return false;
    },
  });
};
