const { contextBridge, ipcRenderer } = require('electron');

// 向渲染进程暴露安全的API
contextBridge.exposeInMainWorld('electronAPI', {
  // 文件选择相关
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  saveDocument: (options) => ipcRenderer.invoke('save-document', options),
  
  // 文件处理相关
  scanFiles: (paths) => ipcRenderer.invoke('scan-files', paths),
  parseCode: (files, options) => ipcRenderer.invoke('parse-code', files, options),
  generateDocument: (data, options) => ipcRenderer.invoke('generate-document', data, options)
});
