import { registerFunctions } from '../registry';
import type { FunctionDef } from '../types';

const modules = import.meta.glob<{ defs?: FunctionDef[] }>('./*.ts', { eager: true });

/** Registers every worksheet function module in this folder. */
export const registerAllFunctions = (): void => {
  for (const [path, mod] of Object.entries(modules)) if (!path.endsWith('/index.ts') && mod.defs) registerFunctions(mod.defs);
};
