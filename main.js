const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs-extra');

// 保持对窗口对象的全局引用
let mainWindow;

function createWindow() {
  // 创建浏览器窗口
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    titleBarStyle: 'hiddenInset', // macOS风格标题栏
    show: false // 先不显示，等加载完成后再显示
  });

  // 加载应用的 index.html
  mainWindow.loadFile('src/renderer/index.html');

  // 窗口加载完成后显示
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // 当窗口被关闭时，取消引用 window 对象
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Electron 初始化完成后创建窗口
app.whenReady().then(createWindow);

// 当所有窗口都被关闭时退出应用
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

// IPC 处理程序
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: '代码文件', extensions: ['js', 'ts', 'py', 'java', 'cpp', 'c', 'cs', 'php', 'go', 'rs', 'vue', 'jsx', 'tsx'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });
  return result;
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'multiSelections']
  });
  return result;
});

ipcMain.handle('save-document', async (event, options) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存Word文档',
    defaultPath: `源程序文档_${new Date().toISOString().slice(0, 10)}.docx`,
    filters: [
      { name: 'Word文档', extensions: ['docx'] }
    ]
  });
  return result;
});

// 导入文件处理模块
const fileHandler = require('./src/main/fileHandler');
const codeParser = require('./src/main/codeParser');
const documentGenerator = require('./src/main/documentGenerator');

// 注册文件处理相关的IPC处理程序
ipcMain.handle('scan-files', async (event, paths) => {
  return await fileHandler.scanFiles(paths);
});

ipcMain.handle('parse-code', async (event, files, options) => {
  return await codeParser.parseCode(files, options);
});

ipcMain.handle('generate-document', async (event, options) => {
  return await documentGenerator.generateDocument(options);
});
