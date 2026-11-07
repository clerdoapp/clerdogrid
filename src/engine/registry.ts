import type { FunctionDef } from './types';

const table = new Map<string, FunctionDef>();

export const registerFunctions = (defs: readonly FunctionDef[]): void => {
  for (const d of defs) table.set(d.name.toUpperCase(), d);
};

export const getFunction = (name: string): FunctionDef | undefined => table.get(name.toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, ''));

export const functionNames = (): string[] => [...table.keys()].sort();

/** Functions handled inside the evaluator because they control how their arguments are evaluated. */
export const SPECIAL_FORMS = [
  'IF', 'IFS', 'IFERROR', 'IFNA', 'SWITCH', 'CHOOSE', 'LET', 'LAMBDA', 'ISOMITTED', 'INDIRECT',
] as const;

export const VOLATILE_SPECIAL = new Set(['INDIRECT']);
