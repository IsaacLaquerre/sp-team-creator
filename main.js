// App entry: main window, lifecycle and module registration. Features live in lib/.
const { app, BrowserWindow, screen } = require('electron');
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
