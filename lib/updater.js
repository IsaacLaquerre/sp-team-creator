// Auto-update from the public repo's GitHub Releases. Checks on launch and downloads in the background.
const { app, ipcMain, shell } = require('electron');
const { autoUpdater } = require('electron-updater');

const GITHUB_OWNER = 'IsaacLaquerre';
const GITHUB_REPO = 'sp-team-creator';

autoUpdater.autoDownload = true;
autoUpdater.autoInstallOnAppQuit = false;
autoUpdater.allowPrerelease = true; // TEMPORARY, remove for stable releases
// electron-builder only publishes latest.yml. Without this, electron-updater
// derives the channel from the version's prerelease tag (e.g. "alpha") and
// requests alpha.yml instead, which is never published.
autoUpdater.channel = 'latest';
autoUpdater.on('error', (err) => console.log('[autoUpdater] error:', err && err.message));

// Newest published (non-draft) release version from GitHub, prereleases included. Cached 10 minutes.
let latestReleaseCache = null; // { ts, version }
async function latestReleaseVersion() {
  if (latestReleaseCache && Date.now() - latestReleaseCache.ts < 10 * 60 * 1000) return latestReleaseCache.version;
  const r = await fetch(`https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/releases?per_page=10`, {
    headers: { Accept: 'application/vnd.github+json', 'User-Agent': GITHUB_REPO },
  });
  if (!r.ok) throw new Error(`GitHub HTTP ${r.status}`);
  const release = (await r.json()).find((x) => !x.draft);
  const version = release ? String(release.tag_name || '').replace(/^v/, '') : '';
  latestReleaseCache = { ts: Date.now(), version };
  return version;
}

function checkForUpdates() {
  if (!app.isPackaged) return;
  autoUpdater.checkForUpdates().catch((err) => console.log('[autoUpdater] check failed:', err && err.message));
}

function register({ getWindow }) {
  autoUpdater.on('update-downloaded', (info) => {
    const version = info && info.version;
    const url = `https://github.com/${GITHUB_OWNER}/${GITHUB_REPO}/releases/tag/v${version}`;
    const win = getWindow();
    if (win && !win.isDestroyed()) win.webContents.send('updateReady', { version, url });
  });

  ipcMain.handle('quitAndInstall', () => autoUpdater.quitAndInstall());
  ipcMain.handle('openExternal', (_e, url) => {
    try { if (new URL(url).origin === 'https://github.com') shell.openExternal(url); } catch {}
  });
}

module.exports = { register, checkForUpdates, latestReleaseVersion };
