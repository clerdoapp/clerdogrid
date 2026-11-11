import pkgText from '../../../package.json?raw';
import contributorsText from '../../../contributors/contributors.json?raw';
import noticeText from '../../../NOTICE?raw';
import { N } from '../../ui/i18n';

export interface Contributor { name: string; role?: string; url?: string; github?: string }

export interface AboutInfo {
  name: string;
  version: string;
  license: string;
  homepage?: string;
  repo?: string;
  links: { label: string; url: string }[];
  contributors: Contributor[];
  notices: string[];
}

const safeUrl = (u: unknown): string | undefined => (typeof u === 'string' && /^https:\/\/[^\s"<>]+$/i.test(u) ? u : undefined);

const parse = <T,>(text: string, fallback: T): T => {
  try { return JSON.parse(text) as T; } catch { return fallback; }
};

/** Splits NOTICE into its third-party entries (lines starting with "- ", continued by indented lines). */
export const parseNotices = (text: string): string[] => {
  const out: string[] = [];
  for (const line of text.split(/\r?\n/)) {
    if (/^\s*-\s+/.test(line)) out.push(line.replace(/^\s*-\s+/, '').trim());
    else if (/^\s{2,}\S/.test(line) && out.length) out[out.length - 1] += ` ${line.trim()}`;
  }
  return out;
};

export const repoUrl = (repository: unknown): string | undefined => {
  const raw = typeof repository === 'string' ? repository : (repository as { url?: unknown } | undefined)?.url;
  if (typeof raw !== 'string') return undefined;
  const m = /github\.com[/:]([\w.-]+)\/([\w.-]+?)(?:\.git)?$/i.exec(raw.trim());
  return m ? `https://github.com/${m[1]}/${m[2]}` : safeUrl(raw);
};

export const aboutInfo = (pkgJson = pkgText, contributorsJson = contributorsText, notice = noticeText, version = __APP_VERSION__): AboutInfo => {
  const pkg = parse<Record<string, unknown>>(pkgJson, {});
  const repo = repoUrl(pkg.repository);
  const homepage = safeUrl(pkg.homepage);
  const list = parse<{ contributors?: unknown }>(contributorsJson, {}).contributors;
  const contributors = (Array.isArray(list) ? list : [])
    .filter((c): c is Record<string, unknown> => !!c && typeof c === 'object' && typeof (c as { name?: unknown }).name === 'string')
    .map((c) => ({ name: String(c.name).slice(0, 80), role: typeof c.role === 'string' ? c.role.slice(0, 80) : undefined, url: safeUrl(c.url), github: typeof c.github === 'string' ? c.github : undefined }));
  const links = [
    homepage && { label: N('Website'), url: homepage },
    repo && { label: N('Source code'), url: repo },
    repo && { label: N('Releases'), url: `${repo}/releases` },
    repo && { label: N('Report a problem'), url: `${repo}/issues` },
  ].filter((x): x is { label: string; url: string } => !!x);
  return {
    name: typeof pkg.productName === 'string' ? pkg.productName : 'ClerdoGrid',
    version,
    license: typeof pkg.license === 'string' ? pkg.license : 'MIT OR Apache-2.0',
    homepage,
    repo,
    links,
    contributors,
    notices: parseNotices(notice),
  };
};
