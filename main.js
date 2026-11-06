const { app, BrowserWindow, ipcMain, session, shell, screen } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileAsync = promisify(execFile);
const { autoUpdater } = require('electron-updater');

// Strips the Electron and app tokens from the user agent.
app.userAgentFallback = app.userAgentFallback.replace(/ (Electron|sp-team-creator|Équipe de projet)\/\S+/g, '');

const TENANT = 'https://consstellaire.sharepoint.com';
const PARTITION = 'persist:sharepoint'; // keeps your Microsoft sign-in between launches
const sp = () => session.fromPartition(PARTITION);
// Fetches without following redirects. Returns a 401 on a redirect.
const spFetch = async (url, opts = {}) => {
  const unauth = () => new Response(null, { status: 401 });
  try {
    const r = await sp().fetch(url, { ...opts, redirect: 'manual' });
    return r.status >= 300 && r.status < 400 ? unauth() : r;
  } catch (e) {
    if (/redirect/i.test(e && e.message)) return unauth();
    throw e;
  }
};

// Turns the group behind a site into a Teams team via EnsureTeamForGroupEx.
// Runs after creation in the background, retrying a few times.
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
ipcMain.handle('teamifyGroup', async (_e, { siteUrl }) => {
  if (!siteUrl) return { ok: false, error: 'No site URL.' };
  const base = siteUrl.replace(/\/$/, '');
  const attempt = async () => {
    const digest = await getDigestFor(base);
    const r = await spFetch(`${base}/_api/GroupSiteManager/EnsureTeamForGroupEx`, {
      method: 'POST', credentials: 'include',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
      },
    });
    if (r.ok) return { ok: true };
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch {}
    const msg = (j && j.error && j.error.message && j.error.message.value) || `HTTP ${r.status}`;
    const err = Object.assign(new Error(msg), { auth: r.status === 401 || r.status === 403 });
    throw err;
  };
  let lastErr;
  for (const delay of [0, 4000, 10000]) {
    if (delay) await sleep(delay);
    try {
      return await attempt();
    } catch (e) {
      lastErr = e;
      if (e.auth) break;
    }
  }
  return { ok: false, auth: !!lastErr.auth, error: lastErr.message };
});

// Unused. Removes memberEmails from the group's Owners, matched by mail.
// Self demotion is commented out below.
ipcMain.handle('demoteAfterCreate', async (_e, { siteUrl, groupId, memberEmails, protectedEmails }) => {
  if (!siteUrl || !groupId) return { ok: false, error: 'Missing site or group.' };
  const base = siteUrl.replace(/\/$/, '');
  try {
    // Self demotion: start
    // const meRes = await spFetch(`${base}/_api/web/currentuser?$select=Email,UserPrincipalName`, {
    //   credentials: 'include',
    //   headers: { Accept: 'application/json;odata=nometadata' },
    // });
    // const me = await meRes.json();
    // const myEmail = (me.Email || me.UserPrincipalName || '').toLowerCase();
    // const protectedSet = new Set((protectedEmails || []).map((e) => e.toLowerCase()));
    // Self demotion: end

    const toDemote = new Set((memberEmails || []).map((e) => e.toLowerCase()));
    // if (myEmail && !protectedSet.has(myEmail)) toDemote.add(myEmail); // Self demotion
    if (!toDemote.size) return { ok: true, removed: [] };

    const ownersRes = await spFetch(
      `${base}/_api/SP.Directory.DirectorySession/Group('${groupId}')/Owners?$top=50&$select=PrincipalName,Id,DisplayName,Mail`,
      { credentials: 'include', headers: { Accept: 'application/json;odata=verbose' } }
    );
    const ownersText = await ownersRes.text();
    let ownersJson;
    try { ownersJson = JSON.parse(ownersText); } catch {}
    if (!ownersRes.ok || !ownersJson) return { ok: false, error: `HTTP ${ownersRes.status} listing owners.` };
    const owners = (ownersJson.d && ownersJson.d.results) || [];

    const digest = await getDigestFor(base);
    const removed = [];
    const failed = [];
    for (const owner of owners) {
      const mail = (owner.mail || '').toLowerCase();
      if (!toDemote.has(mail)) continue;
      const r = await spFetch(
        `${base}/_api/SP.Directory.DirectorySession/Group('${groupId}')/Owners/Remove('${owner.id}')`,
        {
          method: 'POST', credentials: 'include',
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': digest,
          },
        }
      );
      if (r.ok) removed.push(mail);
      else failed.push(mail);
    }
    return { ok: failed.length === 0, removed, failed, error: failed.length ? `Failed: ${failed.join(', ')}` : undefined };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// Gets a user's Entra object ID from their user profile.
async function getObjectId({ email, key }) {
  const claim = key && key.includes('|') ? key : `i:0#.f|membership|${email}`;
  const r = await spFetch(
    `${TENANT}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(claim.replace(/'/g, "''"))}'`,
    { credentials: 'include', headers: { Accept: 'application/json;odata=nometadata' } }
  );
  if (r.status === 401 || r.status === 403) throw Object.assign(new Error(`HTTP ${r.status}`), { auth: true });
  let j; try { j = await r.json(); } catch {}
  const prop = ((j && j.UserProfileProperties) || []).find((p) => p.Key === 'msOnline-ObjectId');
  return (prop && prop.Value) || null;
}

// Adds members to the group's Members, with retries. Falls back to the tenant root without a site URL.
ipcMain.handle('addMembers', async (_e, { siteUrl, groupId, members }) => {
  if (!groupId) return { ok: false, error: 'Missing group.' };
  if (!Array.isArray(members) || !members.length) return { ok: true, added: [] };
  const base = (siteUrl || TENANT).replace(/\/$/, '');
  const added = [];
  let pending = members.map((m) => ({ ...m, error: '' }));
  try {
    for (const delay of [0, 4000, 10000, 20000]) {
      if (delay) await sleep(delay);
      const digest = await getDigestFor(base);
      const still = [];
      for (const m of pending) {
        try {
          const oid = m.oid || (m.oid = await getObjectId(m));
          if (!oid) { m.error = 'user not found'; still.push(m); continue; }
          const r = await spFetch(
            `${base}/_api/SP.Directory.DirectorySession/Group('${groupId}')/Members/Add(objectId='${oid}', principalName='')`,
            {
              method: 'POST', credentials: 'include',
              headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'X-RequestDigest': digest,
              },
            }
          );
          if (r.ok) { added.push(m.email); continue; }
          if (r.status === 401) throw Object.assign(new Error('Not signed in'), { auth: true });
          m.error = `HTTP ${r.status}`; still.push(m);
        } catch (e) {
          if (e.auth) throw e;
          m.error = e.message; still.push(m);
        }
      }
      pending = still;
      if (!pending.length) break;
    }
  } catch (e) {
    return { ok: false, auth: !!e.auth, added, error: e.message };
  }
  const failed = pending.map((m) => `${m.email} (${m.error})`);
  return { ok: !failed.length, added, failed, error: failed.length ? failed.join(', ') : undefined };
});

// --- Install tracking --------------------------------------------------------
// Emails that see the installations panel. Each also needs owner rights on the AppVersion site.
const ADMIN_EMAILS = [
  'ilaquerre@stellaireconstruction.com',
];
const isAdmin = (email) => ADMIN_EMAILS.some((a) => a.toLowerCase() === email);
const INSTALLS_SITE = `${TENANT}/sites/AppVersion`;
const INSTALLS_LIST = '/sites/AppVersion/Lists/sp_team_creator';
const listApi = (p, q = '') => `${INSTALLS_SITE}/_api/web/GetList(@l)${p}?@l='${INSTALLS_LIST}'${q ? '&' + q : ''}`;
const odataStr = (s) => String(s).replace(/'/g, "''");
const JSON_NM = 'application/json;odata=nometadata';

// Signed-in user's email and name.
async function whoAmI() {
  const r = await spFetch(`${TENANT}/_api/web/currentuser?$select=Email,Title,UserPrincipalName`, { credentials: 'include', headers: { Accept: JSON_NM } });
  if (!r.ok) throw Object.assign(new Error(`HTTP ${r.status}`), { auth: r.status === 401 || r.status === 403 });
  const j = await r.json();
  return { email: (j.Email || j.UserPrincipalName || '').toLowerCase(), name: j.Title || '' };
}
ipcMain.handle('whoAmI', async () => {
  try { const me = await whoAmI(); return { ok: true, ...me, isAdmin: isAdmin(me.email) }; }
  catch (e) { return { ok: false, auth: !!e.auth, error: e.message }; }
});

// Maps the list's column display names to internal names, cached.
let installFieldsCache = null;
async function installFields() {
  if (installFieldsCache) return installFieldsCache;
  const r = await spFetch(listApi('/fields', '$select=InternalName,Title&$filter=Hidden eq false'), { credentials: 'include', headers: { Accept: JSON_NM } });
  if (!r.ok) throw new Error(`HTTP ${r.status} reading list columns`);
  const byTitle = {};
  ((await r.json()).value || []).forEach((f) => { byTitle[f.Title.toLowerCase()] = f.InternalName; });
  const pick = (title) => byTitle[title.toLowerCase()] || title;
  installFieldsCache = { userName: pick('UserName'), computer: pick('Computer'), appVersion: pick('AppVersion'), lastSeen: pick('LastSeen') };
  return installFieldsCache;
}

// Creates or updates this user's row for this computer.
ipcMain.handle('reportInstall', async () => {
  try {
    const me = await whoAmI();
    if (!me.email) return { ok: false, error: 'No email.' };
    const f = await installFields();
    const computer = os.hostname();
    const digest = await getDigestFor(INSTALLS_SITE);
    const find = await spFetch(
      listApi('/items', `$select=Id&$top=1&$filter=${encodeURIComponent(`Title eq '${odataStr(me.email)}' and ${f.computer} eq '${odataStr(computer)}'`)}`),
      { credentials: 'include', headers: { Accept: JSON_NM } }
    );
    if (!find.ok) return { ok: false, error: `HTTP ${find.status} finding row` };
    const existing = ((await find.json()).value || [])[0];
    const body = JSON.stringify({
      Title: me.email, [f.userName]: me.name, [f.computer]: computer,
      [f.appVersion]: app.getVersion(), [f.lastSeen]: new Date().toISOString(),
    });
    const headers = { Accept: JSON_NM, 'Content-Type': JSON_NM, 'X-RequestDigest': digest };
    const r = existing
      ? await spFetch(listApi(`/items(${existing.Id})`), { method: 'POST', credentials: 'include', headers: { ...headers, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body })
      : await spFetch(listApi('/items'), { method: 'POST', credentials: 'include', headers, body });
    return r.ok ? { ok: true } : { ok: false, error: `HTTP ${r.status} ${(await r.text()).slice(0, 200)}` };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// Lists every install row. Admin only.
ipcMain.handle('listInstalls', async () => {
  try {
    const me = await whoAmI();
    if (!isAdmin(me.email)) return { ok: false, error: 'Not allowed.' };
    const f = await installFields();
    const r = await spFetch(
      listApi('/items', `$select=Title,${f.userName},${f.computer},${f.appVersion},${f.lastSeen}&$top=500&$orderby=Modified desc`),
      { credentials: 'include', headers: { Accept: JSON_NM } }
    );
    if (!r.ok) return { ok: false, auth: r.status === 401, error: `HTTP ${r.status}` };
    const rows = ((await r.json()).value || []).map((it) => ({
      email: it.Title || '', name: it[f.userName] || '', computer: it[f.computer] || '',
      version: it[f.appVersion] || '', lastSeen: it[f.lastSeen] || '',
    }));
    return { ok: true, rows, currentVersion: app.getVersion() };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// --- Auto-update -----------------------------------------------------------
// Checks GitHub Releases on launch, downloads updates in the background,
// installs on next quit. GITHUB_UPDATE_TOKEN is a fine-grained PAT scoped
// to this repo only, Contents: Read-only.
const GITHUB_UPDATE_TOKEN = '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';
const GITHUB_OWNER = 'IsaacLaquerre';
const GITHUB_REPO = 'sp-team-creator';
autoUpdater.autoDownload = true;
autoUpdater.autoInstallOnAppQuit = true;
autoUpdater.allowPrerelease = true; // TEMPORARY, remove for stable releases
// electron-builder only publishes latest.yml. Without this, electron-updater
// derives the channel from the version's prerelease tag (e.g. "alpha") and
// requests alpha.yml instead, which is never published.
autoUpdater.channel = 'latest';
if (GITHUB_UPDATE_TOKEN) autoUpdater.requestHeaders = { authorization: `token ${GITHUB_UPDATE_TOKEN}` };
autoUpdater.on('error', (err) => console.log('[autoUpdater] error:', err && err.message));

autoUpdater.on('update-downloaded', (info) => {
  const version = info && info.version;
  const url = `https://github.com/${GITHUB_OWNER}/${GITHUB_REPO}/releases/tag/v${version}`;
  if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('updateReady', { version, url });
});

function checkForUpdates() {
  if (!app.isPackaged) return;
  autoUpdater.checkForUpdates().catch((err) => console.log('[autoUpdater] check failed:', err && err.message));
}

ipcMain.handle('quitAndInstall', () => autoUpdater.quitAndInstall());
ipcMain.handle('openExternal', (_e, url) => {
  try { if (new URL(url).origin === 'https://github.com') shell.openExternal(url); } catch {}
});

// Saved window size, position and maximized state.
const WINDOW_STATE_FILE = path.join(app.getPath('userData'), 'window-state.json');
function loadWindowState() {
  try {
    const s = JSON.parse(fs.readFileSync(WINDOW_STATE_FILE, 'utf8'));
    if (!(s.width > 0 && s.height > 0)) return null;
    // Drops the position if under 100px of the window is on any display.
    const onScreen = Number.isFinite(s.x) && Number.isFinite(s.y) && screen.getAllDisplays().some(({ workArea: a }) =>
      Math.min(s.x + s.width, a.x + a.width) - Math.max(s.x, a.x) >= 100 &&
      Math.min(s.y + s.height, a.y + a.height) - Math.max(s.y, a.y) >= 100);
    if (!onScreen) { delete s.x; delete s.y; }
    return s;
  } catch { return null; }
}
function saveWindowState(win) {
  try {
    const { x, y, width, height } = win.getNormalBounds();
    fs.writeFileSync(WINDOW_STATE_FILE, JSON.stringify({ x, y, width, height, maximized: win.isMaximized() }));
  } catch {}
}

let mainWindow = null;
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
  watchSessionCookies();
  checkForUpdates();
});
app.on('window-all-closed', async () => {
  try { await sp().cookies.flushStore(); } catch {}
  app.quit();
  setTimeout(() => app.exit(0), 1500); // force-quits if still running
});

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

// Requests a site-scoped X-RequestDigest.
async function getDigestFor(siteUrl) {
  const r = await spFetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST', credentials: 'include',
    headers: { Accept: 'application/json;odata=nometadata' },
  });
  let j;
  try { j = await r.json(); } catch { j = null; }
  if (!r.ok || !j || !j.FormDigestValue) throw Object.assign(new Error('Not signed in'), { auth: true });
  return j.FormDigestValue;
}
const getDigest = () => getDigestFor(TENANT);

// Checks auth silently, no sign-in window.
ipcMain.handle('checkAuth', async () => {
  try {
    await getDigest();
    persistSessionCookies();
    return { ok: true };
  } catch (e) { return { ok: false, auth: !!e.auth, error: e.message }; }
});

// Resolves an email to its UPN. Falls back to the email.
async function resolveUpn(email, digest) {
  try {
    const r = await spFetch(`${TENANT}/_api/web/ensureuser`, {
      method: 'POST', credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest },
      body: JSON.stringify({ logonName: email }),
    });
    const j = await r.json();
    return j.UserPrincipalName || (j.LoginName || '').split('|').pop() || email;
  } catch { return email; }
}

ipcMain.handle('create', async (_e, { displayName, alias, owners }) => {
  try {
    const digest = await getDigest();
    owners = await Promise.all((Array.isArray(owners) ? owners : []).map((o) => resolveUpn(o, digest)));
    const r = await spFetch(`${TENANT}/_api/GroupSiteManager/CreateGroupEx`, {
      method: 'POST', credentials: 'include',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
      },
      body: JSON.stringify({
        displayName, alias, isPublic: false,
        optionalParams: {
          Description: '', Classification: '',
          Owners: { results: Array.isArray(owners) ? owners : [] },
        },
      }),
    });
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch {}
    if (!r.ok) return { ok: false, error: (j && j.error && j.error.message && j.error.message.value) || `HTTP ${r.status}` };
    const res = j && j.d && j.d.CreateGroupEx;
    return { ok: true, status: res && res.SiteStatus, url: res && res.SiteUrl, groupId: res && res.GroupId };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// Gets a group site's status. SiteStatus: 0 not found, 1 provisioning, 2 ready, 3 error.
ipcMain.handle('siteStatus', async (_e, { groupId }) => {
  try {
    if (!groupId) return { ok: false, error: 'No group.' };
    const digest = await getDigest();
    const r = await spFetch(`${TENANT}/_api/GroupSiteManager/GetSiteStatus`, {
      method: 'POST', credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest },
      body: JSON.stringify({ groupId }),
    });
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch {}
    if (!r.ok || !j) return { ok: false, auth: r.status === 401 || r.status === 403, error: `HTTP ${r.status}` };
    // Deleted group: SiteStatus 3 with an all-zero GroupId.
    const gone = j.SiteStatus === 0 || (j.SiteStatus === 3 && /^[0-]+$/.test(j.GroupId || ''));
    return { ok: true, status: j.SiteStatus, url: j.SiteUrl || '', gone };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// Site deletion isn't offered.

// Checks whether a site still exists.
ipcMain.handle('checkSiteExists', async (_e, { siteUrl }) => {
  try {
    if (!siteUrl) return { ok: false, error: 'No site URL.' };
    const base = siteUrl.replace(/\/$/, '');
    const r = await spFetch(`${base}/_api/web/currentuser?$select=Id`, {
      credentials: 'include', headers: { Accept: 'application/json;odata=nometadata' },
    });
    if (r.ok) return { ok: true, exists: true };
    if (r.status === 401 || r.status === 403) return { ok: false, auth: true, error: `HTTP ${r.status}` };
    if (r.status === 404) return { ok: true, exists: false };
    return { ok: false, error: `HTTP ${r.status}` };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

// Opens OneDrive's "odopen://sync" protocol link to start a sync for the site.
ipcMain.handle('syncSite', async (_e, { siteUrl }) => {
  try {
    if (!siteUrl) return { ok: false, error: 'No site URL.' };
    const base = siteUrl.replace(/\/$/, '');
    const getJson = async (path) => {
      const r = await spFetch(`${base}${path}`, { credentials: 'include', headers: { Accept: 'application/json;odata=nometadata' } });
      if (!r.ok) throw new Error(`HTTP ${r.status} for ${path}`);
      return r.json();
    };
    const [site, web, list, me] = await Promise.all([
      getJson(`/_api/site?$select=Id`),
      getJson(`/_api/web?$select=Id,Title`),
      getJson(`/_api/web/lists/getbytitle('Documents')?$select=Id`),
      getJson(`/_api/web/currentuser?$select=Email,UserPrincipalName`),
    ]);
    // Encodes spaces as %20 instead of URLSearchParams's "+".
    const fields = {
      siteId: site.Id,
      webId: web.Id,
      listId: list.Id,
      webTitle: web.Title || '',
      userEmail: me.Email || me.UserPrincipalName || '',
      webUrl: base,
    };
    const qs = Object.entries(fields).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join('&');
    shell.openExternal(`odopen://sync/?${qs}`);
    return { ok: true };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// Searches people for the owners/members fields.
ipcMain.handle('searchPeople', async (_e, query) => {
  try {
    if (!query || query.length < 2) return { ok: true, people: [] };
    const digest = await getDigest();
    const r = await spFetch(`${TENANT}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`, {
      method: 'POST', credentials: 'include',
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
      },
      body: JSON.stringify({
        queryParams: {
          __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          MaximumEntitySuggestions: 8,
          PrincipalSource: 15, // all sources
          PrincipalType: 1,    // users only (no groups)
          QueryString: query,
        },
      }),
    });
    const text = await r.text();
    let j; try { j = JSON.parse(text); } catch { j = null; }
    const raw = j && j.d && j.d.ClientPeoplePickerSearchUser;
    if (!raw) return { ok: true, people: [] };
    let entities; try { entities = JSON.parse(raw); } catch { entities = []; }
    const people = (entities || [])
      .map((e) => ({
        name: e.DisplayText || (e.EntityData && e.EntityData.Title) || e.Key,
        email: (e.EntityData && e.EntityData.Email) || (e.IsResolved ? e.Key : ''),
        key: e.Key || '',
      }))
      .filter((p) => p.email);
    return { ok: true, people };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message, people: [] };
  }
});

// Lists group-backed sites from SharePoint's recent/frequent sites API.
ipcMain.handle('listRecentSites', async () => {
  try {
    const url = `${TENANT}/_api/v2.1/insights/recentSites?$expand=insights,activityItems($expand=analytics($expand=allTime($expand=activities)))&$top=81`;
    const r = await spFetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
    if (!r.ok) {
      if (r.status === 401 || r.status === 403) throw Object.assign(new Error(`HTTP ${r.status}`), { auth: true });
      return { ok: false, error: `HTTP ${r.status}` };
    }
    const j = await r.json();
    const items = j.value || [];
    const sites = items
      .filter((it) => it.groupId)
      .map((it) => ({ url: it.webUrl || '', title: it.title || '', groupId: it.groupId || null }))
      .filter((s) => s.url);
    const seen = new Set();
    const unique = sites.filter((s) => {
      const key = s.url.replace(/\/$/, '').toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key); return true;
    });
    return { ok: true, sites: unique };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message };
  }
});

// Lists folder names under a directory via PowerShell, cached briefly.
// Used only as a fallback by checkSynced.
const DIR_CACHE_TTL_MS = 8000;
const dirCache = new Map(); // path -> { ts, promise }
function listDirNames(dir) {
  const now = Date.now();
  const cached = dirCache.get(dir);
  if (cached && now - cached.ts < DIR_CACHE_TTL_MS) return cached.promise;
  const promise = (async () => {
    try {
      const psPath = dir.replace(/'/g, "''");
      const { stdout } = await execFileAsync(
        'powershell.exe',
        [
          '-NoProfile', '-NonInteractive', '-Command',
          `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-ChildItem -LiteralPath '${psPath}' -Directory -Name -ErrorAction SilentlyContinue`,
        ],
        { windowsHide: true }
      );
      return stdout.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    } catch {
      return [];
    }
  })();
  dirCache.set(dir, { ts: now, promise });
  return promise;
}

// Parses a ClientPolicy_*.ini file (%LOCALAPPDATA%\Microsoft\OneDrive\
// settings\Business*\) for its site URL, built from RootTenantUrl and
// ScopeBaseUrl. One file exists per currently synced library.
const ONEDRIVE_SETTINGS_CACHE_TTL_MS = 30000;
let oneDriveSettingsCache = null; // { ts, found, sites: [{ url }] }

// Reads a text file, detecting UTF-16. OneDrive's ClientPolicy ini files are
// UTF-16LE with no BOM, so plain ASCII bytes are followed by a zero byte.
function readIniText(filePath) {
  const buf = fs.readFileSync(filePath);
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) return buf.slice(2).toString('utf16le');
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) return buf.slice(2).swap16().toString('utf16le');
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.slice(3).toString('utf8');
  if (buf.length >= 4 && buf[1] === 0x00 && buf[3] === 0x00) return buf.toString('utf16le');
  return buf.toString('utf8');
}

function parseClientPolicyFile(text) {
  const read = (key) => {
    const m = text.match(new RegExp('^' + key + '\\s*=\\s*(.*)$', 'm'));
    return m ? m[1].trim() : '';
  };
  const root = read('RootTenantUrl');
  const scope = read('ScopeBaseUrl');
  const url = root && scope ? scope.replace('{RootUrl}', root) : '';
  const title = read('SiteTitle');
  if (!url && !title) return null;
  return { url: url || '', title: title || '' };
}

// Watches the settings folder. On any change, invalidates the cache and notifies the renderer.
const watchedDirs = new Set();
function watchOneDriveSettingsDir(dirPath) {
  if (watchedDirs.has(dirPath)) return;
  watchedDirs.add(dirPath);
  let debounce;
  try {
    fs.watch(dirPath, () => {
      clearTimeout(debounce);
      debounce = setTimeout(() => {
        oneDriveSettingsCache = null;
        if (mainWindow && !mainWindow.isDestroyed()) mainWindow.webContents.send('oneDriveSettingsChanged');
      }, 500);
    });
  } catch {}
}

// Returns { found, sites }. found is true if the settings folder was reachable.
async function getOneDriveSyncedSites() {
  const now = Date.now();
  if (oneDriveSettingsCache && now - oneDriveSettingsCache.ts < ONEDRIVE_SETTINGS_CACHE_TTL_MS) return oneDriveSettingsCache;
  let sites = [];
  let found = false;
  try {
    const root = process.env.LOCALAPPDATA;
    if (root) {
      const settingsDir = path.join(root, 'Microsoft', 'OneDrive', 'settings');
      const subdirs = fs.readdirSync(settingsDir).filter((d) => /^Business/i.test(d));
      for (const sub of subdirs) {
        const subDir = path.join(settingsDir, sub);
        let files;
        try { files = fs.readdirSync(subDir); } catch { continue; }
        found = true;
        for (const f of files) {
          if (!/^ClientPolicy_.*\.ini$/i.test(f)) continue;
          try {
            const text = readIniText(path.join(subDir, f));
            const site = parseClientPolicyFile(text);
            if (site) sites.push(site);
          } catch {}
        }
        watchOneDriveSettingsDir(subDir);
      }
    }
  } catch {}
  oneDriveSettingsCache = { ts: now, found, sites };
  return oneDriveSettingsCache;
}

ipcMain.handle('checkSynced', async (_e, { siteUrl, num, title, alias }) => {
  const norm = (s) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  try {
    if (siteUrl || title) {
      const { found, sites: synced } = await getOneDriveSyncedSites();
      if (found) {
        const target = siteUrl ? siteUrl.replace(/\/$/, '').toLowerCase() : '';
        const targetTitle = norm(title);
        const synced_ = synced.some((s) => {
          if (target && s.url && s.url.replace(/\/$/, '').toLowerCase() === target) return true;
          if (targetTitle && s.title && norm(s.title) === targetTitle) return true;
          return false;
        });
        return { ok: true, synced: synced_ };
      }
    }
    const envRoots = Object.keys(process.env)
      .filter((k) => /^OneDrive/i.test(k))
      .map((k) => process.env[k])
      .filter(Boolean);
    // Derives the shared-library root from OneDriveCommercial's folder name.
    const derivedRoots = envRoots.map((p) => {
      const base = path.basename(p);
      const org = base.replace(/^OneDrive\s*-\s*/i, '').trim();
      return org && org !== base ? path.join(os.homedir(), org) : null;
    }).filter(Boolean);
    const roots = [...new Set([...envRoots, ...derivedRoots])];
    const needles = [num, title, alias].filter(Boolean).map(norm).filter(Boolean);
    const matches = (name) => needles.some((n) => name.includes(n) || n.includes(name));

    for (const root of roots) {
      const top = await listDirNames(root);
      for (const name of top) {
        if (matches(norm(name))) return { ok: true, synced: true };
      }
      // One level deeper.
      for (const name of top) {
        const sub = await listDirNames(path.join(root, name));
        for (const s of sub) {
          if (matches(norm(s))) return { ok: true, synced: true };
        }
      }
    }
    return { ok: true, synced: false };
  } catch {
    return { ok: true, synced: false };
  }
});

ipcMain.handle('open', (_e, url) => {
  try { if (new URL(url).origin === TENANT) shell.openExternal(url); } catch {}
});
