import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('clerdo', {
  open: () => ipcRenderer.invoke('file:open'),
  launchFile: () => ipcRenderer.invoke('file:launch'),
  readFile: (path: string) => ipcRenderer.invoke('file:read', path),
  save: (path: string | null, name: string, data: Uint8Array) => ipcRenderer.invoke('file:save', path, name, data),
  newWindow: () => ipcRenderer.send('window:new'),
  savePdf: (name: string, data: Uint8Array) => ipcRenderer.invoke('pdf:save', name, data),
  openPdf: (name: string, data: Uint8Array) => ipcRenderer.invoke('pdf:open', name, data),
  setFullScreen: (on: boolean) => ipcRenderer.invoke('window:fullscreen', on),
  isFullScreen: () => ipcRenderer.invoke('window:isFullScreen'),
  arrangeWindows: (how: string) => ipcRenderer.invoke('window:arrange', how),
  windowCount: () => ipcRenderer.invoke('window:count'),
  mirrorPost: (msg: unknown) => ipcRenderer.send('mirror:post', msg),
  netSetAllowed: (on: boolean) => ipcRenderer.invoke('net:setAllowed', on === true),
  netFetch: (req: { url: string; kind: 'text' | 'binary'; allowed: boolean; maxBytes?: number; timeoutMs?: number }) => ipcRenderer.invoke('net:fetch', req),
  onMirror: (fn: (msg: unknown) => void) => { ipcRenderer.on('mirror:message', (_e, msg) => fn(msg)); },
});
