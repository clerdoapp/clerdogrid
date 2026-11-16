// Client for the Clerdo cloud API served next to the app at grid.clerdo.app (server/grid_app).
// Sign-in is a server-side OAuth flow: the browser only holds an HttpOnly session cookie, and every
// write carries the CSRF token from /api/me.

export type CloudErrorCode = 'unauthenticated' | 'csrf' | 'not_found' | 'conflict' | 'validation' | 'limit' | 'too_large' | 'rate_limited' | 'server' | 'network' | 'unknown';

export class CloudError extends Error {
  constructor(readonly code: CloudErrorCode, readonly status: number, message: string, readonly payload?: unknown) {
    super(message);
    this.name = 'CloudError';
  }
}

export interface CloudUser { id: string; name: string; email: string; avatar: string | null }
export interface Me { user: CloudUser; csrf: string | null; limits: { projects: number | null; docBytes: number } }
export interface CloudMeta { id: string; name: string; revision: number; updatedAt: string; createdAt: string; bytes: number; sheets: number; starred: boolean; archived: boolean }
export interface Version { revision: number; savedAt: string; bytes: number }
export interface Conflict { project: CloudMeta; workbook: string }

const KNOWN = new Set<CloudErrorCode>(['unauthenticated', 'csrf', 'not_found', 'conflict', 'validation', 'limit', 'too_large', 'rate_limited', 'server']);

let csrf: string | null = null;
let mePromise: Promise<Me> | null = null;

const request = async <T>(method: string, path: string, body?: unknown, retried = false): Promise<T> => {
  const headers: Record<string, string> = { Accept: 'application/json' };
  const init: RequestInit = { method, credentials: 'same-origin', headers };
  if (body !== undefined) {
    headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(body);
  }
  if (method !== 'GET') {
    if (csrf === null && !path.startsWith('/api/share/')) await me().catch(() => null);
    if (csrf) headers['X-CSRF-Token'] = csrf;
  }
  let res: Response;
  try { res = await fetch(path, init); } catch { throw new CloudError('network', 0, 'offline'); }
  if (res.status === 204) return undefined as T;
  let data: unknown = null;
  try { data = await res.json(); } catch { data = null; }
  if (res.ok) return data as T;
  const err = (data ?? {}) as { error?: string; message?: string };
  const code: CloudErrorCode = err.error && KNOWN.has(err.error as CloudErrorCode) ? err.error as CloudErrorCode : res.status >= 500 ? 'server' : 'unknown';
  if (code === 'csrf' && !retried) {
    csrf = null;
    mePromise = null;
    return request<T>(method, path, body, true);
  }
  throw new CloudError(code, res.status, typeof err.message === 'string' ? err.message : code, data);
};

export const me = (force = false): Promise<Me> => {
  if (!mePromise || force) {
    mePromise = request<Me>('GET', '/api/me').then((m) => { csrf = m.csrf; return m; });
    mePromise.catch(() => { mePromise = null; });
  }
  return mePromise;
};

export const signInUrl = (next: string): string => `/auth/login?next=${encodeURIComponent(next.startsWith('/') && !next.startsWith('//') ? next : '/')}`;

export const cloud = {
  logout: async (): Promise<void> => { await request<void>('POST', '/auth/logout', {}); csrf = null; mePromise = null; },
  list: () => request<{ projects: CloudMeta[] }>('GET', '/api/projects').then((r) => r.projects),
  create: (name: string, workbook: string, sheets: number) => request<{ project: CloudMeta }>('POST', '/api/projects', { name, workbook, sheets }).then((r) => r.project),
  get: (id: string) => request<{ project: CloudMeta; workbook: string }>('GET', `/api/projects/${encodeURIComponent(id)}`),
  save: (id: string, baseRevision: number, workbook: string, sheets: number, name: string) =>
    request<{ project: CloudMeta }>('PUT', `/api/projects/${encodeURIComponent(id)}`, { baseRevision, workbook, sheets, name }).then((r) => r.project),
  remove: (id: string) => request<void>('DELETE', `/api/projects/${encodeURIComponent(id)}`),
  versions: (id: string) => request<{ current: Version; revisions: Version[] }>('GET', `/api/projects/${encodeURIComponent(id)}/revisions`),
  restore: (id: string, revision: number) => request<{ project: CloudMeta; workbook: string }>('POST', `/api/projects/${encodeURIComponent(id)}/revisions/${revision}/restore`, {}),
  share: (id: string) => request<{ token: string; url: string }>('POST', `/api/projects/${encodeURIComponent(id)}/share`, { mode: 'view' }),
  unshare: (id: string) => request<void>('DELETE', `/api/projects/${encodeURIComponent(id)}/share`),
  shared: (token: string) => request<{ name: string; workbook: string }>('GET', `/api/share/${encodeURIComponent(token)}`),
};

/** Standard base64 (with padding), as the server's strict base64_decode expects. */
export const toBase64 = (bytes: Uint8Array): string => {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
};

export const fromBase64 = (text: string): Uint8Array => {
  const bin = atob(text);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
};
