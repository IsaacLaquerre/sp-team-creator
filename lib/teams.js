// Team creation, membership, people search and site lookups.
const { ipcMain, shell } = require('electron');
const { TENANT, JSON_NM, sleep, spFetch, getDigestFor, getDigest, resolveUpn, getObjectId } = require('./sharepoint');

const VERBOSE = 'application/json;odata=verbose';

function register() {
  ipcMain.handle('create', async (_e, { displayName, alias, owners }) => {
    try {
      const digest = await getDigest();
      owners = await Promise.all((Array.isArray(owners) ? owners : []).map((o) => resolveUpn(o, digest)));
      const r = await spFetch(`${TENANT}/_api/GroupSiteManager/CreateGroupEx`, {
        method: 'POST', credentials: 'include',
        headers: { Accept: VERBOSE, 'Content-Type': VERBOSE, 'X-RequestDigest': digest },
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

  // Turns the group behind a site into a Teams team via EnsureTeamForGroupEx.
  // Runs after creation in the background, retrying a few times.
  ipcMain.handle('teamifyGroup', async (_e, { siteUrl }) => {
    if (!siteUrl) return { ok: false, error: 'No site URL.' };
    const base = siteUrl.replace(/\/$/, '');
    const attempt = async () => {
      const digest = await getDigestFor(base);
      const r = await spFetch(`${base}/_api/GroupSiteManager/EnsureTeamForGroupEx`, {
        method: 'POST', credentials: 'include',
        headers: { Accept: VERBOSE, 'Content-Type': VERBOSE, 'X-RequestDigest': digest },
      });
      if (r.ok) return { ok: true };
      const text = await r.text();
      let j; try { j = JSON.parse(text); } catch {}
      const msg = (j && j.error && j.error.message && j.error.message.value) || `HTTP ${r.status}`;
      throw Object.assign(new Error(msg), { auth: r.status === 401 || r.status === 403 });
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
              { method: 'POST', credentials: 'include', headers: { Accept: VERBOSE, 'Content-Type': VERBOSE, 'X-RequestDigest': digest } }
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
        { credentials: 'include', headers: { Accept: VERBOSE } }
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
          { method: 'POST', credentials: 'include', headers: { Accept: VERBOSE, 'Content-Type': VERBOSE, 'X-RequestDigest': digest } }
        );
        if (r.ok) removed.push(mail);
        else failed.push(mail);
      }
      return { ok: failed.length === 0, removed, failed, error: failed.length ? `Failed: ${failed.join(', ')}` : undefined };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  });

  // Searches people for the owners/members fields.
  ipcMain.handle('searchPeople', async (_e, query) => {
    try {
      if (!query || query.length < 2) return { ok: true, people: [] };
      const digest = await getDigest();
      const r = await spFetch(`${TENANT}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`, {
        method: 'POST', credentials: 'include',
        headers: { Accept: VERBOSE, 'Content-Type': VERBOSE, 'X-RequestDigest': digest },
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

  // Gets a group site's status. SiteStatus: 0 not found, 1 provisioning, 2 ready, 3 error.
  ipcMain.handle('siteStatus', async (_e, { groupId }) => {
    try {
      if (!groupId) return { ok: false, error: 'No group.' };
      const digest = await getDigest();
      const r = await spFetch(`${TENANT}/_api/GroupSiteManager/GetSiteStatus`, {
        method: 'POST', credentials: 'include',
        headers: { Accept: JSON_NM, 'Content-Type': JSON_NM, 'X-RequestDigest': digest },
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

  // Checks whether a site still exists. Site deletion isn't offered.
  ipcMain.handle('checkSiteExists', async (_e, { siteUrl }) => {
    try {
      if (!siteUrl) return { ok: false, error: 'No site URL.' };
      const base = siteUrl.replace(/\/$/, '');
      const r = await spFetch(`${base}/_api/web/currentuser?$select=Id`, { credentials: 'include', headers: { Accept: JSON_NM } });
      if (r.ok) return { ok: true, exists: true };
      if (r.status === 401 || r.status === 403) return { ok: false, auth: true, error: `HTTP ${r.status}` };
      if (r.status === 404) return { ok: true, exists: false };
      return { ok: false, error: `HTTP ${r.status}` };
    } catch (e) {
      return { ok: false, error: e.message };
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
      const sites = (j.value || [])
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

  // Opens a tenant URL in the default browser.
  ipcMain.handle('open', (_e, url) => {
    try { if (new URL(url).origin === TENANT) shell.openExternal(url); } catch {}
  });
}

module.exports = { register };
