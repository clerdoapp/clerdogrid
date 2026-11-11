// Bundles the command-line tool into one file that runs on plain Node: npm run cli
import { readFileSync, readdirSync, statSync, mkdirSync, chmodSync, existsSync } from 'node:fs';
import { dirname, join, relative, resolve, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const outArg = process.argv.indexOf('--outfile');
const outfile = outArg > 0 && process.argv[outArg + 1] ? resolve(process.argv[outArg + 1]) : join(root, 'dist-cli', 'clerdogrid.mjs');

/** "./*.ts", "../formats/*.ts", "./*\/index.ts" or "{a,b}" alternatives, matched against real files. */
const globToRegex = (pattern) => {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '*') re += pattern[i + 1] === '*' ? (i++, '.*') : '[^/]*';
    else if (ch === '{') {
      const end = pattern.indexOf('}', i);
      re += `(?:${pattern.slice(i + 1, end).split(',').map((s) => s.replace(/[.+^$()|[\]\\]/g, '\\$&')).join('|')})`;
      i = end;
    } else if (ch === '?') re += '[^/]';
    else re += ch.replace(/[.+^$()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}$`);
};

const listFiles = (dir, depth = 3) => {
  const out = [];
  if (!existsSync(dir)) return out;
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) { if (depth > 0) out.push(...listFiles(p, depth - 1)); } else out.push(p);
  }
  return out;
};

// Vite's import.meta.glob, done at bundle time so the same sources build with esbuild.
// With skipLazy, lazy globs (the optional format modules) come out empty.
const globPlugin = (skipLazy) => ({
  name: 'import-meta-glob',
  setup(b) {
    b.onLoad({ filter: /\.(ts|tsx)$/ }, (args) => {
      const src = readFileSync(args.path, 'utf8');
      if (!src.includes('import.meta.glob')) return undefined;
      const dir = dirname(args.path);
      let n = 0;
      const imports = [];
      const code = src.replace(/import\.meta\.glob(?:<[^>]*>)?\(\s*(['"])([^'"]+)\1\s*(?:,\s*(\{[^}]*\}))?\s*\)/g, (_m, _q, pattern, opts) => {
        const eager = !!opts && /eager\s*:\s*true/.test(opts);
        if (!eager && skipLazy) return '({})';
        // Interface languages are for the app only; the command-line tool stays in English.
        if (!eager && /[\\/]ui[\\/]i18n[\\/]/.test(args.path)) return '({})';
        const re = globToRegex(pattern.startsWith('.') ? pattern : `./${pattern}`);
        const files = listFiles(resolve(dir, pattern.split('*')[0].replace(/[^/]*$/, '') || '.'))
          .map((f) => { const r = relative(dir, f).split('\\').join('/'); return r.startsWith('.') ? r : `./${r}`; })
          .filter((r) => re.test(r) && posix.normalize(r) !== posix.normalize(`./${relative(dir, args.path)}`))
          .sort();
        const entries = files.map((f) => {
          if (!eager) return `${JSON.stringify(f)}: () => import(${JSON.stringify(f)})`;
          const id = `__glob_${n++}`;
          imports.push(`import * as ${id} from ${JSON.stringify(f)};`);
          return `${JSON.stringify(f)}: ${id}`;
        });
        return `({${entries.join(', ')}})`;
      });
      return { contents: imports.join('\n') + '\n' + code, loader: args.path.endsWith('x') ? 'tsx' : 'ts' };
    });
  },
});

mkdirSync(dirname(outfile), { recursive: true });
const options = (skipLazy) => ({
  entryPoints: [join(root, 'src/cli/main.ts')],
  outfile,
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node20',
  splitting: false,
  legalComments: 'none',
  banner: { js: "#!/usr/bin/env node\nimport { createRequire as __cgRequire } from 'node:module'; const require = __cgRequire(import.meta.url);" },
  define: { __APP_VERSION__: JSON.stringify(pkg.version), __CLI_BUNDLE__: 'true', 'import.meta.env.DEV': 'false', 'import.meta.env.PROD': 'true' },
  loader: { '.json': 'json' },
  logLevel: 'warning',
  plugins: [globPlugin(skipLazy)],
});

try {
  await build({ ...options(false), logLevel: process.argv.includes('--strict') ? 'warning' : 'silent' });
} catch (e) {
  // A format module that doesn't build yet shouldn't stop the tool: build without the optional ones.
  if (process.argv.includes('--strict')) throw e;
  await build({ ...options(true), logLevel: 'silent' });
  console.warn(`Built without the optional format modules, which failed to build:\n${e.errors?.map((x) => `  ${x.location?.file ?? ''}:${x.location?.line ?? ''} ${x.text}`).join('\n') ?? e.message}`);
}
chmodSync(outfile, 0o755);
console.log(relative(root, outfile));
