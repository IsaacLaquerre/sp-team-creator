// Install tracking in a SharePoint list and the admin installations view.
const { app, ipcMain } = require('electron');
const os = require('os');
const { TENANT, JSON_NM, odataStr, spFetch, getDigestFor, whoAmI } = require('./sharepoint');
const { latestReleaseVersion } = require('./updater');

// Emails that see the installations panel. Each also needs owner rights on the AppVersion site.
const ADMIN_EMAILS = [
  'ilaquerre@stellaireconstruction.com',
];
const isAdmin = (email) => ADMIN_EMAILS.some((a) => a.toLowerCase() === email);
// Last SharePoint-confirmed admin status of the signed-in user.
let currentIsAdmin = false;
async function whoAmIChecked() {
  const me = await whoAmI();
  currentIsAdmin = isAdmin(me.email);
  return me;
}
const INSTALLS_SITE = `${TENANT}/sites/AppVersion`;
const INSTALLS_LIST = '/sites/AppVersion/Lists/sp_team_creator';
const listApi = (p, q = '') => `${INSTALLS_SITE}/_api/web/GetList(@l)${p}?@l='${INSTALLS_LIST}'${q ? '&' + q : ''}`;

// Maps the list's column display names to internal names, cached.
let installFieldsCache = null;
async function installFields() {
  if (installFieldsCache) return installFieldsCache;
  const r = await spFetch(listApi('/fields', '$select=InternalName,Title&$filter=Hidden eq false'), { credentials: 'include', headers: { Accept: JSON_NM } });
  if (!r.ok) throw new Error(`HTTP ${r.status} reading list columns`);
  const byTitle = {};
  ((await r.json()).value || []).forEach((f) => { byTitle[f.Title.toLowerCase()] = f.InternalName; });
  const pick = (title) => byTitle[title.toLowerCase()] || title;
  installFieldsCache = {
    userName: pick('UserName'), computer: pick('Computer'), appVersion: pick('AppVersion'), lastSeen: pick('LastSeen'),
    role: byTitle.role || null, // optional column
  };
  return installFieldsCache;
}

function register() {
  ipcMain.handle('whoAmI', async () => {
    try { const me = await whoAmIChecked(); return { ok: true, ...me, isAdmin: currentIsAdmin }; }
    catch (e) { return { ok: false, auth: !!e.auth, error: e.message }; }
  });

  // Creates or updates this user's row for this computer.
  ipcMain.handle('reportInstall', async () => {
    try {
      const me = await whoAmIChecked();
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
        ...(f.role ? { [f.role]: isAdmin(me.email) ? 'Admin' : 'User' } : {}),
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
      const me = await whoAmIChecked();
      if (!isAdmin(me.email)) return { ok: false, error: 'Not allowed.' };
      const f = await installFields();
      const r = await spFetch(
        listApi('/items', `$select=Title,${f.userName},${f.computer},${f.appVersion},${f.lastSeen}${f.role ? ',' + f.role : ''}&$top=500&$orderby=Modified desc`),
        { credentials: 'include', headers: { Accept: JSON_NM } }
      );
      if (!r.ok) return { ok: false, auth: r.status === 401, error: `HTTP ${r.status}` };
      const rows = ((await r.json()).value || []).map((it) => ({
        email: it.Title || '', name: it[f.userName] || '', computer: it[f.computer] || '',
        version: it[f.appVersion] || '', lastSeen: it[f.lastSeen] || '', role: (f.role && it[f.role]) || '',
      }));
      let latestError = '';
      const latestVersion = await latestReleaseVersion().catch((e) => { latestError = e.message; return ''; });
      return { ok: true, rows, latestVersion, latestError };
    } catch (e) {
      return { ok: false, auth: !!e.auth, error: e.message };
    }
  });
}

module.exports = { register, isCurrentUserAdmin: () => currentIsAdmin };
