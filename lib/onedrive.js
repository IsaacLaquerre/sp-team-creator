// OneDrive sync detection and the odopen:// sync link.
const { ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { promisify } = require('util');
const { JSON_NM, spFetch } = require('./sharepoint');

const execFileAsync = promisify(execFile);

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
let getWindow = () => null;
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
        const win = getWindow();
        if (win && !win.isDestroyed()) win.webContents.send('oneDriveSettingsChanged');
      }, 500);
    });
  } catch {}
}

// Returns { found, sites }. found is true if the settings folder was reachable.
async function getOneDriveSyncedSites() {
  const now = Date.now();
  if (oneDriveSettingsCache && now - oneDriveSettingsCache.ts < ONEDRIVE_SETTINGS_CACHE_TTL_MS) return oneDriveSettingsCache;
  const sites = [];
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
            const site = parseClientPolicyFile(readIniText(path.join(subDir, f)));
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

function register(opts) {
  getWindow = opts.getWindow;

  ipcMain.handle('checkSynced', async (_e, { siteUrl, num, title, alias }) => {
    const norm = (s) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    try {
      if (siteUrl || title) {
        const { found, sites: synced } = await getOneDriveSyncedSites();
        if (found) {
          const target = siteUrl ? siteUrl.replace(/\/$/, '').toLowerCase() : '';
          const targetTitle = norm(title);
          const isSynced = synced.some((s) => {
            if (target && s.url && s.url.replace(/\/$/, '').toLowerCase() === target) return true;
            if (targetTitle && s.title && norm(s.title) === targetTitle) return true;
            return false;
          });
          return { ok: true, synced: isSynced };
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

  // Opens OneDrive's "odopen://sync" protocol link to start a sync for the site.
  ipcMain.handle('syncSite', async (_e, { siteUrl }) => {
    try {
      if (!siteUrl) return { ok: false, error: 'No site URL.' };
      const base = siteUrl.replace(/\/$/, '');
      const getJson = async (p) => {
        const r = await spFetch(`${base}${p}`, { credentials: 'include', headers: { Accept: JSON_NM } });
        if (!r.ok) throw new Error(`HTTP ${r.status} for ${p}`);
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
}

module.exports = { register };
