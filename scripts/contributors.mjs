// Refreshes contributors/contributors.json from git history and, when reachable, the GitHub API.
// Maintainers run it before a release: npm run contributors
//   --offline     use git only
//   --check       exit 1 if the file would change (for CI), without writing
// GITHUB_TOKEN (optional) raises the API rate limit. Nothing is fetched at app run time; the app
// only reads the JSON file this writes. Emails are never stored.
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const file = join(root, 'contributors', 'contributors.json');
const args = new Set(process.argv.slice(2));
const offline = args.has('--offline');
const checkOnly = args.has('--check');

const current = JSON.parse(readFileSync(file, 'utf8'));
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

const BOT = /\[bot\]$|^dependabot|^github-actions|^renovate/i;
// Different spellings of the same contributor end up in one entry.
const ALIASES = [
  { match: /^claude\b/i, name: 'Claude (AI assistant)' },
  { match: /^clerdo\b/i, name: 'Clerdo' },
];
const canonical = (name) => ALIASES.find((a) => a.match.test(name))?.name ?? name.trim();

const git = (argv) => {
  try {
    return execFileSync('git', argv, { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
  } catch {
    return '';
  }
};

/** name -> commits, from authors and Co-authored-by trailers. */
const fromGit = () => {
  const counts = new Map();
  const add = (out) => {
    for (const line of out.split('\n')) {
      const m = /^\s*(\d+)\s+(.+?)(?:\s+<[^>]*>)?\s*$/.exec(line);
      if (!m || BOT.test(m[2])) continue;
      const name = canonical(m[2]);
      counts.set(name, (counts.get(name) ?? 0) + Number(m[1]));
    }
  };
  add(git(['shortlog', '-sn', '--no-merges', 'HEAD']));
  add(git(['shortlog', '-sn', '--no-merges', '--group=trailer:co-authored-by', 'HEAD']));
  return counts;
};

/** login -> { name, url, commits } from the repository's contributor list, or null when offline. */
const fromGitHub = async () => {
  if (offline) return null;
  const m = /github\.com[/:]([^/]+)\/([^/.]+)/.exec(pkg.repository?.url ?? '');
  if (!m) return null;
  const headers = { accept: 'application/vnd.github+json', 'user-agent': 'clerdogrid-contributors-script' };
  if (process.env.GITHUB_TOKEN) headers.authorization = `Bearer ${process.env.GITHUB_TOKEN}`;
  const out = [];
  try {
    for (let page = 1; page <= 10; page++) {
      const ctl = new AbortController();
      const timer = setTimeout(() => ctl.abort(), 10_000);
      const res = await fetch(`https://api.github.com/repos/${m[1]}/${m[2]}/contributors?per_page=100&page=${page}`, { headers, signal: ctl.signal });
      clearTimeout(timer);
      if (!res.ok) { console.warn(`GitHub API answered ${res.status}; using git history only.`); return null; }
      const list = await res.json();
      if (!Array.isArray(list) || !list.length) break;
      for (const u of list) if (u?.type !== 'Bot' && typeof u?.login === 'string' && !BOT.test(u.login)) out.push({ login: u.login, url: `https://github.com/${u.login}`, commits: u.contributions ?? 0 });
      if (list.length < 100) break;
    }
  } catch (e) {
    console.warn(`GitHub API not reachable (${e?.name === 'AbortError' ? 'timed out' : e?.message ?? e}); using git history only.`);
    return null;
  }
  return out;
};

const merge = (gitCounts, gh) => {
  const byName = new Map(current.contributors.map((c) => [c.name.toLowerCase(), { ...c }]));
  for (const [name, commits] of gitCounts) {
    const key = name.toLowerCase();
    const e = byName.get(key) ?? { name, role: 'Contributor' };
    e.commits = commits;
    byName.set(key, e);
  }
  for (const u of gh ?? []) {
    const hit = [...byName.values()].find((c) => c.github?.toLowerCase() === u.login.toLowerCase() || c.name.toLowerCase() === u.login.toLowerCase());
    if (hit) { hit.github = u.login; hit.url ??= u.url; hit.commits = Math.max(hit.commits ?? 0, u.commits); continue; }
    byName.set(u.login.toLowerCase(), { name: u.login, github: u.login, url: u.url, role: 'Contributor', commits: u.commits });
  }
  return [...byName.values()]
    .map(({ email: _drop, ...c }) => c)
    .sort((a, b) => Number(!!b.pinned) - Number(!!a.pinned) || (b.commits ?? 0) - (a.commits ?? 0) || a.name.localeCompare(b.name));
};

const gitCounts = fromGit();
if (!gitCounts.size) console.warn('No git history found; keeping the listed contributors.');
const contributors = merge(gitCounts, await fromGitHub());
const next = { ...current, contributors };
const same = JSON.stringify(next.contributors) === JSON.stringify(current.contributors);
if (checkOnly) {
  console.log(same ? 'contributors.json is up to date.' : 'contributors.json is out of date: run npm run contributors.');
  process.exit(same ? 0 : 1);
}
if (same) {
  console.log(`contributors.json is up to date (${contributors.length} contributors).`);
} else {
  next.updated = new Date().toISOString().slice(0, 10);
  writeFileSync(file, JSON.stringify(next, null, 2) + '\n');
  console.log(`Wrote contributors/contributors.json (${contributors.length} contributors).`);
}
