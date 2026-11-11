// PDF export of a whole workbook object, for the command-line tool and other callers without the app.
// Fonts come from `fonts`; under Node the bundled font files are read from node_modules.
import type { Workbook } from '../model/workbook';
import { GridStore } from '../app/store';
import { FontBook, type FontSource } from './pdf-text';
import { buildPdf, type PrintOptions } from '../features/pagelayout/print';

export interface PdfExportOptions {
  /** Print only this sheet (0-based); otherwise every visible sheet. */
  sheet?: number;
  what?: PrintOptions['what'];
  ignorePrintArea?: boolean;
  fonts?: FontSource;
  /** File name shown by the &F header code. */
  fileName?: string;
}

const nodeFonts = async (): Promise<FontSource> => {
  const fs = await import('node:fs');
  const path = await import('node:path');
  const roots = [path.join(process.cwd(), 'node_modules', '@fontsource')];
  try { roots.push(path.join(path.dirname(new URL(import.meta.url).pathname), '..', '..', 'node_modules', '@fontsource')); } catch { /* not a file URL */ }
  const files = new Map<string, string>();
  const ranges: Record<string, Record<string, string>> = {};
  for (const root of roots) {
    let pkgs: string[] = [];
    try { pkgs = fs.readdirSync(root); } catch { continue; }
    for (const pkg of pkgs) {
      try { for (const f of fs.readdirSync(path.join(root, pkg, 'files'))) if (f.endsWith('.woff') && !files.has(f)) files.set(f, path.join(root, pkg, 'files', f)); } catch { /* not a font package */ }
      try { ranges[pkg] ??= JSON.parse(fs.readFileSync(path.join(root, pkg, 'unicode.json'), 'utf8')); } catch { /* no coverage data */ }
    }
  }
  return { files: () => [...files.keys()], ranges: () => ranges, load: async (f) => new Uint8Array(fs.readFileSync(files.get(f)!)) };
};

export const exportPdf = async (wb: Workbook, o: PdfExportOptions = {}): Promise<Uint8Array> => {
  const store = new GridStore(wb);
  if (o.sheet !== undefined && wb.sheets[o.sheet]) store.wb.activeSheet = o.sheet;
  const fonts = o.fonts ?? (typeof process !== 'undefined' && process.versions?.node ? await nodeFonts() : { files: () => [], load: async () => { throw new Error('No fonts.'); } });
  const what = o.what ?? (o.sheet !== undefined ? 'sheet' : 'workbook');
  return buildPdf(store, { what, sheet: o.sheet, ignorePrintArea: o.ignorePrintArea }, { fonts: new FontBook(fonts), file: o.fileName });
};
