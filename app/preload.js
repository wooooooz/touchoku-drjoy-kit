'use strict';
// preload.js — safe bridge between the sandboxed renderer and the Node-privileged main process.
//   20261106
const { contextBridge, ipcRenderer, webUtils, webFrame } = require('electron');

contextBridge.exposeInMainWorld('td', {
  loadConfig: () => ipcRenderer.invoke('td:loadConfig'),
  loadMaster: (q) => ipcRenderer.invoke('td:loadMaster', q),
  saveDept: (q) => ipcRenderer.invoke('td:saveDept', q),
  // 前月の 全日タイプ／昼夜タイプ（空の月の「昼夜共通」スイッチの初期状態。20260922）
  loadPrevTypes: (q) => ipcRenderer.invoke('td:loadPrevTypes', q),
  exportXlsx: (q) => ipcRenderer.invoke('td:exportXlsx', q),
  importXlsx: () => ipcRenderer.invoke('td:importXlsx'),
  // ドラッグ＆ドロップ: File → 実パス（Electron 32 以降は File.path が無いので webUtils で取る）
  importXlsxPath: (p) => ipcRenderer.invoke('td:importXlsxPath', { path: p }),
  pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch (_) { return ''; } },
  // 未提出の変更の有無（20260922）: main が「閉じる」前の確認に使う
  setDirty: (flag) => ipcRenderer.send('td:setDirty', !!flag),
  // 早見表を画像として保存（20261106 管理者）: rect = 画面の中の位置（CSS px）
  saveImage: (q) => ipcRenderer.invoke('td:saveImage', q),
  // 画面の拡大（文字を大きく・Ctrl＋ホイール。20260927）
  setZoom: (f) => { try { webFrame.setZoomFactor(+f || 1); } catch (_) { /* 拡大できなくても入力は続けられる */ } },
});
