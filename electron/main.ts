import { app, BrowserWindow, dialog, ipcMain, Menu, screen, session, shell } from 'electron';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { existsSync, statSync } from 'node:fs';
import { basename, extname, join, resolve } from 'node:path';
import { relayFetch, setRelayAllowed } from './netRelay';

const MAX_FILE = 512 * 1024 * 1024;
// File types the app may write, by extension, with the save dialog's filter name.
const SAVE_TYPES: Record<string, string> = {
  xlsx: 'Workbook', ods: 'OpenDocument spreadsheet', ots: 'OpenDocument template', xls: 'Older workbook', xlsb: 'Binary workbook',
  html: 'Web page', xml: 'XML data', json: 'JSON data', csv: 'Comma-separated values', txt: 'Text',
};
const OPEN_FILTERS = [{ name: 'Workbooks', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xlam', 'xlsb', 'xls', 'xlt', 'ods', 'ots', 'csv', 'tsv', 'txt', 'prn'] }];
const OPENABLE = new Set(OPEN_FILTERS[0]!.extensions.map((e) => `.${e}`));
// Only paths the person picked in a dialog, or opened from the system, can be written.
const allowed = new Set<string>();
const launch = new Map<number, string>();
const queued: string[] = [];

const workbooksIn = (argv: string[], cwd = process.cwd()): string[] =>
  argv.filter((a) => !a.startsWith('-')).map((a) => resolve(cwd, a)).filter((p) => OPENABLE.has(extname(p).toLowerCase()) && existsSync(p) && statSync(p).isFile());

const readChecked = async (path: string) => {
  if ((await stat(path)).size > MAX_FILE) throw new Error('This file is too large to open.');
  allowed.add(path);
  return { path, name: basename(path), data: new Uint8Array(await readFile(path)) };
};

const createWindow = (file?: string): void => {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 640,
    minHeight: 480,
    title: 'ClerdoGrid',
    backgroundColor: '#eceef3',
    webPreferences: { preload: join(__dirname, 'preload.cjs'), contextIsolation: true, sandbox: true, nodeIntegration: false, spellcheck: false },
  });
  win.setMenuBarVisibility(false);
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:|^mailto:/i.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.webContents.on('will-prevent-unload', (e) => {
    const choice = dialog.showMessageBoxSync(win, {
      type: 'question', buttons: ['Keep editing', 'Close without saving'], defaultId: 0, cancelId: 0,
      message: 'This workbook has unsaved changes.',
    });
    if (choice === 1) e.preventDefault();
  });
  if (file) launch.set(win.webContents.id, file);
  void win.loadFile(join(__dirname, '..', 'dist', 'index.html'));
};

const openWorkbooks = (paths: string[]): void => {
  for (const p of paths) createWindow(p);
  BrowserWindow.getAllWindows().at(-1)?.focus();
};

ipcMain.handle('file:launch', async (e) => {
  const path = launch.get(e.sender.id);
  launch.delete(e.sender.id);
  return path ? readChecked(path) : null;
});

// Files the person picked to import data from (Get Data), kept across launches so Refresh can read
// them again. Only paths in this list can be read back, and only workbook and text files.
const sourcesFile = () => join(app.getPath('userData'), 'data-sources.json');
let sources: Set<string> | null = null;
const loadSources = async (): Promise<Set<string>> => {
  if (sources) return sources;
  try { const list = JSON.parse(await readFile(sourcesFile(), 'utf8')) as unknown; sources = new Set(Array.isArray(list) ? list.filter((x): x is string => typeof x === 'string').slice(-500) : []); }
  catch { sources = new Set(); }
  return sources;
};
const rememberSource = async (path: string): Promise<void> => {
  const set = await loadSources();
  set.delete(path); set.add(path);
  try { await mkdir(app.getPath('userData'), { recursive: true }); await writeFile(sourcesFile(), JSON.stringify([...set].slice(-500))); } catch { /* best effort */ }
};

ipcMain.handle('file:read', async (_e, path: unknown) => {
  if (typeof path !== 'string' || path.length > 4096) return null;
  const full = resolve(path);
  if (full !== path || !OPENABLE.has(extname(full).toLowerCase())) return null;
  if (!(await loadSources()).has(full) && !allowed.has(full)) return null;
  if (!existsSync(full) || !statSync(full).isFile()) return null;
  return readChecked(full);
});

ipcMain.handle('file:open', async (e) => {
  const win = BrowserWindow.fromWebContents(e.sender) ?? undefined;
  const r = win ? await dialog.showOpenDialog(win, { properties: ['openFile'], filters: OPEN_FILTERS }) : await dialog.showOpenDialog({ properties: ['openFile'], filters: OPEN_FILTERS });
  const path = r.filePaths[0];
  if (r.canceled || !path) return null;
  void rememberSource(path);
  return readChecked(path);
});

ipcMain.handle('file:save', async (e, path: unknown, name: unknown, data: unknown) => {
  if (!(data instanceof Uint8Array)) throw new Error('Nothing to save.');
  const wanted = typeof name === 'string' ? extname(name).slice(1).toLowerCase() : 'xlsx';
  const ext = SAVE_TYPES[wanted] ? wanted : 'xlsx';
  const sameType = (p: string) => extname(p).slice(1).toLowerCase() === ext;
  let target = typeof path === 'string' && allowed.has(path) && sameType(path) ? path : null;
  if (!target) {
    const win = BrowserWindow.fromWebContents(e.sender) ?? undefined;
    const opts = { defaultPath: typeof name === 'string' ? basename(name) : `Book1.${ext}`, filters: [{ name: SAVE_TYPES[ext]!, extensions: [ext] }] };
    const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
    if (r.canceled || !r.filePath) return null;
    target = sameType(r.filePath) ? r.filePath : `${r.filePath}.${ext}`;
    allowed.add(target);
  }
  await writeFile(target, data);
  return { path: target, name: basename(target) };
});

ipcMain.on('window:new', () => createWindow());

// Internet access for worksheet functions (see netRelay.ts). Only the app's own pages may ask.
const fromApp = (e: Electron.IpcMainInvokeEvent): boolean => (e.senderFrame?.url ?? '').startsWith('file:');
ipcMain.handle('net:setAllowed', (e, on: unknown) => (fromApp(e) ? setRelayAllowed(on) : false));
ipcMain.handle('net:fetch', (e, req: unknown) => (fromApp(e) ? relayFetch(req) : { ok: false, error: 'not allowed' }));

// PDF export: the page sends finished PDF bytes; only real PDFs of a sane size are written.
const MAX_PDF = 512 * 1024 * 1024;
const checkPdf = (data: unknown): Uint8Array => {
  if (!(data instanceof Uint8Array) || data.length < 8 || data.length > MAX_PDF) throw new Error('Nothing to save.');
  if (String.fromCharCode(...data.subarray(0, 5)) !== '%PDF-') throw new Error('Not a PDF.');
  return data;
};
const pdfBase = (name: unknown): string => (typeof name === 'string' ? basename(name) : 'Book1.pdf').replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').replace(/\.pdf$/i, '').slice(0, 120) || 'Book1';

ipcMain.handle('pdf:save', async (e, name: unknown, data: unknown) => {
  const bytes = checkPdf(data);
  const win = BrowserWindow.fromWebContents(e.sender) ?? undefined;
  const opts = { defaultPath: `${pdfBase(name)}.pdf`, filters: [{ name: 'PDF', extensions: ['pdf'] }] };
  const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return null;
  const target = /\.pdf$/i.test(r.filePath) ? r.filePath : `${r.filePath}.pdf`;
  await writeFile(target, bytes);
  return { path: target };
});

ipcMain.handle('pdf:open', async (_e, name: unknown, data: unknown) => {
  const bytes = checkPdf(data);
  const dir = join(app.getPath('temp'), 'ClerdoGrid');
  await mkdir(dir, { recursive: true });
  const target = join(dir, `${pdfBase(name)}-${Date.now().toString(36)}.pdf`);
  await writeFile(target, bytes);
  const err = await shell.openPath(target);
  if (err) throw new Error(err);
  return true;
});

ipcMain.handle('window:fullscreen', (e, on: unknown) => {
  const win = BrowserWindow.fromWebContents(e.sender);
  if (!win) return false;
  win.setFullScreen(on === true);
  return win.isFullScreen();
});
ipcMain.handle('window:isFullScreen', (e) => BrowserWindow.fromWebContents(e.sender)?.isFullScreen() ?? false);
ipcMain.handle('window:count', () => BrowserWindow.getAllWindows().length);
// Lays every ClerdoGrid window out on the display of the window that asked.
ipcMain.handle('window:arrange', (e, how: unknown) => {
  const wins = BrowserWindow.getAllWindows().filter((w) => !w.isMinimized());
  const me = BrowserWindow.fromWebContents(e.sender);
  if (!wins.length || !me) return;
  const area = screen.getDisplayMatching(me.getBounds()).workArea;
  const n = wins.length;
  wins.forEach((w, i) => {
    if (w.isFullScreen()) w.setFullScreen(false);
    if (w.isMaximized()) w.unmaximize();
    let b: Electron.Rectangle;
    if (how === 'vertical') { const wd = Math.floor(area.width / n); b = { x: area.x + i * wd, y: area.y, width: wd, height: area.height }; }
    else if (how === 'horizontal') { const ht = Math.floor(area.height / n); b = { x: area.x, y: area.y + i * ht, width: area.width, height: ht }; }
    else if (how === 'cascade') { const off = 32 * i; b = { x: area.x + off, y: area.y + off, width: Math.max(640, area.width - 32 * (n - 1)), height: Math.max(480, area.height - 32 * (n - 1)) }; }
    else {
      const cols = Math.ceil(Math.sqrt(n)), rows = Math.ceil(n / cols);
      const wd = Math.floor(area.width / cols), ht = Math.floor(area.height / rows);
      b = { x: area.x + (i % cols) * wd, y: area.y + Math.floor(i / cols) * ht, width: wd, height: ht };
    }
    w.setBounds(b);
  });
  me.focus();
});
// Read-only mirror windows of a workbook: messages are relayed to every other window.
ipcMain.on('mirror:post', (e, msg: unknown) => {
  for (const w of BrowserWindow.getAllWindows()) if (w.webContents.id !== e.sender.id) w.webContents.send('mirror:message', msg);
});

if (!app.requestSingleInstanceLock()) app.quit();
app.on('second-instance', (_e, argv, cwd) => {
  const files = workbooksIn(argv.slice(1), cwd);
  if (files.length) openWorkbooks(files);
  else BrowserWindow.getAllWindows()[0]?.focus();
});
app.on('open-file', (e, path) => {
  e.preventDefault();
  if (app.isReady()) openWorkbooks([path]); else queued.push(path);
});

void app.whenReady().then(() => {
  // The page never needs the camera, location or other permissions.
  session.defaultSession.setPermissionRequestHandler((_wc, permission, cb) => cb(permission === 'clipboard-sanitized-write' || permission === 'clipboard-read'));
  if (process.platform === 'darwin') {
    Menu.setApplicationMenu(Menu.buildFromTemplate([
      { role: 'appMenu' },
      { label: 'File', submenu: [{ label: 'New Window', accelerator: 'CmdOrCtrl+Shift+N', click: () => createWindow() }, { role: 'close' }] },
      { role: 'editMenu' }, { role: 'viewMenu' }, { role: 'windowMenu' },
    ]));
  } else Menu.setApplicationMenu(null);
  const files = [...queued, ...workbooksIn(process.argv.slice(app.isPackaged ? 1 : 2))];
  if (files.length) openWorkbooks(files); else createWindow();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
