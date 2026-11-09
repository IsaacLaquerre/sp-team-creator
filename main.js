// App entry: main window, lifecycle and module registration. Features live in lib/.
const { app, BrowserWindow, Menu, screen } = require('electron');
const path = require('path');
const authSession = require('./lib/session');
const teams = require('./lib/teams');
const onedrive = require('./lib/onedrive');
const installs = require('./lib/installs');
const updater = require('./lib/updater');
const { loadWindowState, saveWindowState } = require('./lib/windowState');

// Strips the Electron and app tokens from the user agent.
app.userAgentFallback = app.userAgentFallback.replace(/ (Electron|sp-team-creator|Équipe de projet)\/\S+/g, '');

let mainWindow = null;
const getWindow = () => mainWindow;

authSession.register();
teams.register();
onedrive.register({ getWindow });
installs.register();
updater.register({ getWindow });

// Opens DevTools only for SharePoint-confirmed admins. Removes the default menu and its shortcuts.
function restrictDevTools(win) {
  Menu.setApplicationMenu(null);
  win.webContents.on('before-input-event', (e, input) => {
    const k = (input.key || '').toLowerCase();
    const isDevToolsKey = k === 'f12' || (input.control && input.shift && ['i', 'j', 'c'].includes(k));
    if (input.type !== 'keyDown' || !isDevToolsKey) return;
    e.preventDefault();
    if (installs.isCurrentUserAdmin()) win.webContents.toggleDevTools();
  });
  win.webContents.on('devtools-opened', () => {
    if (!installs.isCurrentUserAdmin()) win.webContents.closeDevTools();
  });
}

app.whenReady().then(() => {
  const { height: screenHeight } = screen.getPrimaryDisplay().workAreaSize;
  const saved = loadWindowState();
  const win = new BrowserWindow({
    width: 600, height: Math.round(screenHeight * 0.9), minWidth: 420, minHeight: 480,
    ...(saved ? { width: saved.width, height: saved.height, x: saved.x, y: saved.y } : {}),
    icon: path.join(__dirname, 'assets', process.platform === 'win32' ? 'icon.ico' : 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true },
  });
  if (saved && saved.maximized) win.maximize();
  win.on('close', () => saveWindowState(win));
  // Keeps the window on index.html: blocks navigation and new windows.
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  if (app.isPackaged) restrictDevTools(win);
  mainWindow = win;
  win.setMenuBarVisibility(false);
  win.loadFile('index.html');
  authSession.watchSessionCookies();
  updater.checkForUpdates();
});

app.on('window-all-closed', async () => {
  await authSession.flushCookies();
  app.quit();
  setTimeout(() => app.exit(0), 1500); // force-quits if still running
});
