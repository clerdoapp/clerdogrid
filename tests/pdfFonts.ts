// Node font source for PDF tests: reads the bundled font files from node_modules.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { FontSource } from '../src/formats/pdf-text';

const root = join(__dirname, '..', 'node_modules', '@fontsource');
export const nodeFontSource = (): FontSource => {
  const paths = new Map<string, string>();
  const ranges: Record<string, Record<string, string>> = {};
  for (const pkg of readdirSync(root)) {
    try { for (const f of readdirSync(join(root, pkg, 'files'))) if (f.endsWith('.woff')) paths.set(f, join(root, pkg, 'files', f)); } catch { /* not a font package */ }
    try { ranges[pkg] = JSON.parse(readFileSync(join(root, pkg, 'unicode.json'), 'utf8')); } catch { /* no coverage data */ }
  }
  return { files: () => [...paths.keys()], ranges: () => ranges, load: async (f) => new Uint8Array(readFileSync(paths.get(f)!)) };
};
