// Builds the command-line bundle into a temporary folder and runs it with plain Node, as users do.
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

export const ROOT = resolve(__dirname, '..');

export const buildCli = (): { bin: string; dir: string; cleanup: () => void } => {
  const dir = mkdtempSync(join(tmpdir(), 'clerdogrid-cli-'));
  const bin = join(dir, 'clerdogrid.mjs');
  execFileSync(process.execPath, [join(ROOT, 'src/cli/build.mjs'), '--outfile', bin], { cwd: ROOT, stdio: 'pipe' });
  return { bin, dir, cleanup: () => rmSync(dir, { recursive: true, force: true }) };
};

export interface Run { code: number; out: string; err: string }

export const runCli = (bin: string, args: string[], opts: { cwd?: string; input?: string } = {}): Run => {
  const r = spawnSync(process.execPath, [bin, ...args], { cwd: opts.cwd ?? ROOT, input: opts.input, encoding: 'utf8', env: { ...process.env, VITEST: '' }, timeout: 60_000 });
  return { code: r.status ?? -1, out: r.stdout, err: r.stderr };
};
