import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
    processFile: (filePath: string) => ipcRenderer.invoke('process-file', filePath),
    selectFile: () => ipcRenderer.invoke('select-file'),
    openConfig: () => ipcRenderer.invoke('open-config'),
    getMappings: () => ipcRenderer.invoke('get-mappings'),
    updateMapping: (oldChar: string, newChar: string) => ipcRenderer.invoke('update-mapping', oldChar, newChar),
    deleteMapping: (char: string) => ipcRenderer.invoke('delete-mapping', char)
});