// Feature modules register their commands, panel sections, dialogs and overlays here.
// Each module exports `register()`; App calls registerFeatures() once at start.
const modules = import.meta.glob<{ register?: () => void }>('./*/index.ts', { eager: true });

export const registerFeatures = (): void => {
  for (const m of Object.values(modules)) m.register?.();
};
