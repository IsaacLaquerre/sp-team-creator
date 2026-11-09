// Sign-in window, auth check and SharePoint cookie persistence.
const { BrowserWindow, ipcMain } = require('electron');
const { TENANT, PARTITION, sp, getDigest } = require('./sharepoint');

const cookieUrl = (c) => `http${c.secure ? 's' : ''}://${c.domain.replace(/^\./, '')}${c.path}`;
const isSharePointCookie = (c) => /(^|\.)sharepoint\.com$/i.test(c.domain);

// Gives SharePoint session cookies a 90 day expiration. Host-only cookies are set without a domain.
function persistCookie(c) {
  const farFuture = Math.floor(Date.now() / 1000) + 60 * 60 * 24 * 90; // 90 days
  return sp().cookies.set({
    url: cookieUrl(c), name: c.name, value: c.value, path: c.path,
    ...(c.hostOnly ? {} : { domain: c.domain }),
    secure: c.secure, httpOnly: c.httpOnly, sameSite: c.sameSite || 'unspecified',
    expirationDate: farFuture,
  }).catch(() => {});
}
async function persistSessionCookies() {
  try {
    const cookies = await sp().cookies.get({});
    await Promise.all(cookies.filter((c) => c.session && isSharePointCookie(c)).map(persistCookie));
  } catch {}
}

// Persists SharePoint session cookies 2s after their last change.
function watchSessionCookies() {
  let timer;
  sp().cookies.on('changed', (_e, c) => {
    if (!isSharePointCookie(c)) return;
    clearTimeout(timer);
    timer = setTimeout(persistSessionCookies, 2000);
  });
}

// Writes cookies to disk.
async function flushCookies() {
  try { await sp().cookies.flushStore(); } catch {}
}

// Microsoft session cookies cleared before sign-in.
const STALE_MS_COOKIES = new Set(['ESTSAUTH', 'ESTSAUTHLIGHT', 'x-ms-gateway-slice', 'stsservicecookie']);

// Clears SharePoint cookies and STALE_MS_COOKIES. Keeps ESTSAUTHPERSISTENT.
async function clearStaleCookies() {
  try {
    const cookies = await sp().cookies.get({});
    const stale = cookies.filter((c) => isSharePointCookie(c) || STALE_MS_COOKIES.has(c.name));
    await Promise.all(stale.map((c) => sp().cookies.remove(cookieUrl(c), c.name).catch(() => {})));
  } catch {}
}

function register() {
  // Opens a Microsoft 365 sign-in window, hidden for the first 3s.
  ipcMain.handle('login', () => clearStaleCookies().then(() => new Promise((resolve) => {
    const w = new BrowserWindow({ width: 520, height: 720, title: 'SharePoint', show: false, webPreferences: { partition: PARTITION } });
    w.setMenuBarVisibility(false);
    const showTimer = setTimeout(() => { if (!w.isDestroyed()) w.show(); }, 3000);
    let done = false;
    const finish = (v) => { if (!done) { done = true; clearTimeout(showTimer); if (v) persistSessionCookies().then(() => resolve(v)); else resolve(v); if (!w.isDestroyed()) w.close(); } };
    w.webContents.on('did-finish-load', () => {
      try { if (new URL(w.webContents.getURL()).origin === TENANT) finish(true); } catch {}
    });
    w.on('closed', () => finish(false));
    w.loadURL(TENANT);
  })));

  // Checks auth silently, no sign-in window.
  ipcMain.handle('checkAuth', async () => {
    try {
      await getDigest();
      persistSessionCookies();
      return { ok: true };
    } catch (e) { return { ok: false, auth: !!e.auth, error: e.message }; }
  });
}

module.exports = { register, watchSessionCookies, flushCookies };
