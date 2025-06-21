import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
    processFile: (filePath: string) => ipcRenderer.invoke('process-file', filePath),
    selectFile: () => ipcRenderer.invoke('select-file')
});