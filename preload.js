const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('api', {
  create: (d) => ipcRenderer.invoke('create', d),
  login: () => ipcRenderer.invoke('login'),
  open: (u) => ipcRenderer.invoke('open', u),
  searchPeople: (q) => ipcRenderer.invoke('searchPeople', q),
  checkAuth: () => ipcRenderer.invoke('checkAuth'),
  syncSite: (d) => ipcRenderer.invoke('syncSite', d),
  listRecentSites: () => ipcRenderer.invoke('listRecentSites'),
  checkSiteExists: (d) => ipcRenderer.invoke('checkSiteExists', d),
  checkSynced: (d) => ipcRenderer.invoke('checkSynced', d),
});
