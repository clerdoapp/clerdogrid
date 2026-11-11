import { renameSync, unlinkSync, writeFileSync } from 'node:fs';
import { basename, dirname, join } from 'node:path';
import { EXIT, UserError } from './errors';

/** Writes next to the target first, so a failed save never leaves half a file. */
export const writeFileSafely = (path: string, data: Uint8Array): void => {
  const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.tmp`);
  try {
    writeFileSync(tmp, data);
    renameSync(tmp, path);
  } catch (e) {
    try { unlinkSync(tmp); } catch { /* not created */ }
    throw new UserError(`Couldn't save ${path}: ${(e as NodeJS.ErrnoException).code ?? (e as Error).message}`, EXIT.file);
  }
};
