import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('clerdo', {
  open: () => ipcRenderer.invoke('file:open'),
  launchFile: () => ipcRenderer.invoke('file:launch'),
  save: (path: string | null, name: string, data: Uint8Array) => ipcRenderer.invoke('file:save', path, name, data),
  newWindow: () => ipcRenderer.send('window:new'),
});
