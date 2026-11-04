const { app, BrowserWindow, ipcMain, session, shell, screen } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileAsync = promisify(execFile);
const { autoUpdater } = require('electron-updater');

const TENANT = 'https://consstellaire.sharepoint.com';
const PARTITION = 'persist:sharepoint'; // keeps your Microsoft sign-in between launches
const sp = () => session.fromPartition(PARTITION);

// --- Auto-update -----------------------------------------------------------
// Checks the GitHub Releases configured in package.json's build.publish for
// a newer version on every launch, downloads it in the background if found,
// and installs it automatically the next time the app quits — nobody has to
// go looking for updates or click anything; they're just always on the
// latest version a launch or two after a new one is published.
//
// IMPORTANT — this points at a PRIVATE repo ("private": true in
// package.json), and GitHub requires authentication to download release
// assets from a private repo even via electron-updater's normal download
// URLs. GITHUB_UPDATE_TOKEN below needs to be a GitHub fine-grained Personal
// Access Token scoped to ONLY this one repo, with "Contents: Read-only" and
// nothing else — it ships inside the .exe handed to colleagues, so it must
// never be a token with any broader access, and should be rotated/revoked
// if the app is ever redistributed outside the company. If embedding any
// token at all isn't acceptable, the simpler alternative is making the
// GitHub repo public instead (nothing secret lives in this app's source —
// sign-in is interactive per user) and removing this token plus
// "private": true in package.json.
const GITHUB_UPDATE_TOKEN = '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';
autoUpdater.autoDownload = true;
autoUpdater.autoInstallOnAppQuit = true;
// TEMPORARY, while shipping alpha builds (versions like "1.0.1-alpha.1"):
// explicitly accept prerelease versions, so alpha installs keep picking up
// newer alphas. Remove this line (or set it to false) once you're back to
// shipping plain stable versions — otherwise a stable install could later
// pick up an alpha build by mistake.
autoUpdater.allowPrerelease = true;
if (GITHUB_UPDATE_TOKEN) autoUpdater.requestHeaders = { authorization: `token ${GITHUB_UPDATE_TOKEN}` };
autoUpdater.on('error', (err) => console.log('[autoUpdater] error:', err && err.message));
function checkForUpdates() {
  // Only meaningful in a packaged (installed) build — running via "npm
  // start"/"electron ." in development has no installer to update and
  // autoUpdater would just fail its own internal checks.
  if (!app.isPackaged) return;
  autoUpdater.checkForUpdates().catch((err) => console.log('[autoUpdater] check failed:', err && err.message));
}

app.whenReady().then(() => {
  const { height: screenHeight } = screen.getPrimaryDisplay().workAreaSize;
  const win = new BrowserWindow({
    width: 500, height: Math.round(screenHeight * 0.9), minWidth: 420, minHeight: 480,
    icon: path.join(__dirname, 'assets', process.platform === 'win32' ? 'icon.ico' : 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true },
  });
  win.setMenuBarVisibility(false);
  win.loadFile('index.html');
  checkForUpdates();
});
app.on('window-all-closed', () => {
  app.quit();
  // Safety net: electron-updater's background update check can leave an
  // open network handle that keeps Node's event loop alive even after
  // app.quit() runs, which silently prevents the process from actually
  // exiting (the window disappears, but electron.exe stays in Task
  // Manager). If the app hasn't actually terminated shortly after a normal
  // quit, force it — this only ever fires if something is still hanging
  // around; a clean quit finishes well before this timeout and the process
  // is already gone by the time it would run.
  setTimeout(() => app.exit(0), 1500);
});

// Electron only writes a cookie to disk (so it survives the app being fully
// closed and relaunched) if the cookie itself has an expiration date. Most of
// the cookies Microsoft's sign-in actually sets (the SharePoint "FedAuth"/
// "rtFa" session cookies in particular) are plain *session* cookies — no
// expiration — because the real session lifetime is enforced server-side,
// not by the cookie. A normal browser keeps those alive because the browser
// process itself doesn't exit between visits; this app's process does, once
// per launch, so Chromium drops them and the next launch has nothing to
// resume. This doesn't change what the server allows — Microsoft still
// expires/invalidates the underlying session on its own schedule and a stale
// cookie just gets a 401, same as always — it only stops Electron's own
// in-memory-only handling of session cookies from throwing away a session
// that's still otherwise good, between one launch and the next.
async function persistSessionCookies() {
  try {
    const cookies = await sp().cookies.get({});
    const farFuture = Math.floor(Date.now() / 1000) + 60 * 60 * 24 * 90; // 90 days
    await Promise.all(cookies.filter((c) => c.session).map((c) => {
      const url = `http${c.secure ? 's' : ''}://${c.domain.replace(/^\./, '')}${c.path}`;
      return sp().cookies.set({
        url, name: c.name, value: c.value, domain: c.domain, path: c.path,
        secure: c.secure, httpOnly: c.httpOnly, sameSite: c.sameSite || 'unspecified',
        expirationDate: farFuture,
      }).catch(() => {});
    }));
  } catch {}
}

// Sign-in window: you log in with your Microsoft 365 account, the session is reused for requests.
ipcMain.handle('login', () => new Promise((resolve) => {
  const w = new BrowserWindow({ width: 520, height: 720, title: 'SharePoint', webPreferences: { partition: PARTITION } });
  w.setMenuBarVisibility(false);
  let done = false;
  const finish = (v) => { if (!done) { done = true; if (v) persistSessionCookies().then(() => resolve(v)); else resolve(v); if (!w.isDestroyed()) w.close(); } };
  w.webContents.on('did-finish-load', () => {
    try { if (new URL(w.webContents.getURL()).origin === TENANT) finish(true); } catch {}
  });
  w.on('closed', () => finish(false));
  w.loadURL(TENANT);
}));

async function getDigest() {
  const r = await sp().fetch(`${TENANT}/_api/contextinfo`, {
    method: 'POST', credentials: 'include',
    headers: { Accept: 'application/json;odata=nometadata' },
  });
  let j;
  try { j = await r.json(); } catch { j = null; }
  if (!r.ok || !j || !j.FormDigestValue) throw Object.assign(new Error('Not signed in'), { auth: true });
  return j.FormDigestValue;
}

// Silent check only — no sign-in window, no site creation. Used on startup to
// see if the persisted session (see PARTITION above) still works, so the
// user is only bugged with a sign-in popup when it doesn't.
ipcMain.handle('checkAuth', async () => {
  try {
    await getDigest();
    // Microsoft sometimes reissues a session cookie (no expiration) on an
    // otherwise-valid request, which would quietly undo the persistence
    // below — so re-apply it on every successful check, not just at login.
    persistSessionCookies();
    return { ok: true };
  } catch (e) { return { ok: false, auth: !!e.auth, error: e.message }; }
});

ipcMain.handle('create', async (_e, { displayName, alias, owners }) => {
  try {
    const digest = await getDigest();
    const r = await sp().fetch(`${TENANT}/_api/GroupSiteManager/CreateGroupEx`, {
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
          // Email addresses of people to add as owners, in addition to the signed-in user.
          // odata=verbose requires collection properties wrapped as { results: [...] },
          // not a bare array — a bare array is what threw the StartArray/PrimitiveValue error.
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

// Deleting sites isn't offered by this app at all — scrapped rather than
// shipped half-working, since the real delete endpoint was never confirmed.

// Lightweight existence check for a site, used only to notice when a site
// shown in the list (most often one surfaced via "recent sites") has
// actually been deleted elsewhere, so it can be quietly dropped from view.
ipcMain.handle('checkSiteExists', async (_e, { siteUrl }) => {
  try {
    if (!siteUrl) return { ok: false, error: 'No site URL.' };
    const base = siteUrl.replace(/\/$/, '');
    const r = await sp().fetch(`${base}/_api/web/currentuser?$select=Id`, {
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

// Resolves the site/web/list IDs and the signed-in user's email, then opens
// the "odopen://sync" link OneDrive's desktop client registers — a local
// protocol launch, not an HTTP call, and only works if OneDrive is installed
// and signed in on this machine. This now actually launches OneDrive
// successfully (confirmed) — the previous version of this handler left out
// webTitle, which turned out to control the local folder name: OneDrive
// fell back to the tenant/org name for every site without it, which is also
// why synced folders piled up as "<org name> - Documents (1)", "(2)", etc.
// all colliding on the same name.
ipcMain.handle('syncSite', async (_e, { siteUrl }) => {
  try {
    if (!siteUrl) return { ok: false, error: 'No site URL.' };
    const base = siteUrl.replace(/\/$/, '');
    const getJson = async (path) => {
      const r = await sp().fetch(`${base}${path}`, { credentials: 'include', headers: { Accept: 'application/json;odata=nometadata' } });
      if (!r.ok) throw new Error(`HTTP ${r.status} for ${path}`);
      return r.json();
    };
    const [site, web, list, me] = await Promise.all([
      getJson(`/_api/site?$select=Id`),
      getJson(`/_api/web?$select=Id,Title`),
      getJson(`/_api/web/lists/getbytitle('Documents')?$select=Id`),
      getJson(`/_api/web/currentuser?$select=Email,UserPrincipalName`),
    ]);
    // Built by hand rather than with URLSearchParams: URLSearchParams
    // encodes spaces as "+" (form-encoding), but OneDrive's handler just
    // percent-decodes without turning "+" back into a space, which was
    // showing up literally in the synced folder name. encodeURIComponent
    // uses %20 instead, which decodes correctly.
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

// People search for the owners field, using SharePoint's own people-picker API.
// This runs over the same signed-in cookie session as everything else here — it
// does not need a separate Microsoft Graph token or app registration.
ipcMain.handle('searchPeople', async (_e, query) => {
  try {
    if (!query || query.length < 2) return { ok: true, people: [] };
    const digest = await getDigest();
    const r = await sp().fetch(`${TENANT}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`, {
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
      }))
      .filter((p) => p.email);
    return { ok: true, people };
  } catch (e) {
    return { ok: false, auth: !!e.auth, error: e.message, people: [] };
  }
});

// Lists sites from SharePoint's own "recent/frequent sites" insights API — the
// same one the SharePoint start page uses. Confirmed against a real response
// from the tenant: each top-level entry in "value" is one site, with fields
// webUrl, title, and — for sites backed by a Microsoft 365 group (template
// "GROUP", which is what this app creates) — groupId directly on the item.
// We only keep group-backed sites: the tenant's root communication site and
// personal "Mon espace de travail" container also show up here but aren't
// project sites in the sense this app deals with, and they carry no groupId.
// Nested activityItems/insights/analytics (recently viewed files, access
// counts) aren't used — only the site-level fields are.
ipcMain.handle('listRecentSites', async () => {
  try {
    const url = `${TENANT}/_api/v2.1/insights/recentSites?$expand=insights,activityItems($expand=analytics($expand=allTime($expand=activities)))&$top=81`;
    const r = await sp().fetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
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

// Best-effort check for whether a site is already synced to this computer's
// OneDrive, for sites that were synced some other way (OneDrive's own "Sync"
// button, or before this app existed) rather than through this app's sync
// button. There's no official API for "is this library synced" that a
// regular desktop app can call — OneDrive doesn't expose that state to
// anyone else. This instead looks for a local folder, under OneDrive's known
// root(s) and one level into them (OneDrive often nests synced libraries
// under a tenant/org folder), whose name matches the site's project number,
// title or alias — the number prefix (e.g. "430") is checked first since
// it's the most distinctive part and the least likely to be mangled by
// OneDrive's own folder-naming (it often appends " - Documents" etc., which
// a full-title match can be thrown off by).
// NOTE: heuristic, not certain — a folder that was renamed after syncing, or
// a name that happens to collide, can give a wrong answer either way.
// Shells out rather than using Node's fs module directly. Confirmed: on this
// machine, fs.readdirSync() from inside this app's process was silently
// returning an incomplete listing of these exact OneDrive folders (missing
// nearly all of them) while a manually opened cmd.exe window, reading the
// identical path, saw everything correctly — strongly suggesting something
// (Controlled Folder Access, or antivirus/EDR software) is specifically
// filtering what this app's own process can see via direct file APIs.
// Spawning a separate, trusted process to do the listing sidesteps that.
// Uses PowerShell's Get-ChildItem rather than cmd.exe's "dir": dir prints a
// literal "File Not Found" line to its own output for every folder with no
// subfolders (most of them), which was leaking into the terminal;
// Get-ChildItem just returns nothing for an empty folder. It also avoids
// cmd.exe's OEM-codepage text output, which was mangling accented folder
// names (Québec, Réunions, etc.) into garbage when decoded as UTF-8.
//
// Called with execFileAsync (not execFileSync): a synchronous spawn would
// block Electron's entire main process — every window, every IPC call — for
// as long as PowerShell takes to start, which is much slower than cmd.exe
// was and made the whole app feel frozen once several sites were each
// triggering their own checkSynced call.
//
// Results are cached per directory for a few seconds (dirCache below) so
// that checking several sites back-to-back — which all need to list the
// same OneDrive roots — reuses one scan instead of each site re-spawning
// PowerShell over the same folders from scratch.
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

ipcMain.handle('checkSynced', async (_e, { num, title, alias }) => {
  try {
    const envRoots = Object.keys(process.env)
      .filter((k) => /^OneDrive/i.test(k))
      .map((k) => process.env[k])
      .filter(Boolean);
    // A synced *shared library* (as opposed to the user's own personal
    // OneDrive) lands in a separate root that isn't exposed by any
    // environment variable: "<home>\<Org Name>\", where "<Org Name>" is the
    // same org name OneDriveCommercial uses, just without its
    // "OneDrive - " prefix (confirmed: OneDriveCommercial was
    // "...\OneDrive - GESTION E.S.C INC" and the real sync root was
    // "...\GESTION E.S.C INC").
    const derivedRoots = envRoots.map((p) => {
      const base = path.basename(p);
      const org = base.replace(/^OneDrive\s*-\s*/i, '').trim();
      return org && org !== base ? path.join(os.homedir(), org) : null;
    }).filter(Boolean);
    const roots = [...new Set([...envRoots, ...derivedRoots])];
    const norm = (s) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const needles = [num, title, alias].filter(Boolean).map(norm).filter(Boolean);
    const matches = (name) => needles.some((n) => name.includes(n) || n.includes(name));

    for (const root of roots) {
      const top = await listDirNames(root);
      for (const name of top) {
        if (matches(norm(name))) return { ok: true, synced: true };
      }
      // One level deeper (e.g. "<OneDrive root>\<Organization>\<Site> - Documents")
      for (const name of top) {
        const sub = await listDirNames(path.join(root, name));
        for (const s of sub) {
          if (matches(norm(s))) return { ok: true, synced: true };
        }
      }
    }
    return { ok: true, synced: false };
  } catch {
    return { ok: true, synced: false }; // heuristic only — never blocks the UI
  }
});

ipcMain.handle('open', (_e, url) => {
  try { if (new URL(url).origin === TENANT) shell.openExternal(url); } catch {}
});
