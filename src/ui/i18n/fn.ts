// Function descriptions and argument hints in the interface language. Function names stay in English
// everywhere; only the descriptions are translated. Each language is its own lazy chunk keyed by the
// English text from src/engine/catalog.ts; text without a translation shows in English.
const loaders = import.meta.glob<{ default: Record<string, string> }>('./fn/*.ts');

let current: Record<string, string> = {};

/** Loads the descriptions for `code` (called when the interface language changes). */
export const loadFnText = async (code: string): Promise<void> => {
  const load = loaders[`./fn/${code}.ts`];
  if (!load) { current = {}; return; }
  try { current = (await load()).default; } catch { current = {}; }
};

/** A catalog description or argument hint in the interface language. */
export const fnText = (english: string | undefined): string => (english ? current[english] ?? english : '');
