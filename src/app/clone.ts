import type { Sheet } from '../model/workbook';

/** Deep copy of a sheet (Maps and typed arrays included). */
export const cloneSheet = (s: Sheet): Sheet => structuredClone(s);
