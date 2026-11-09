// SharePoint tenant, session partition, REST fetch and shared lookups.
const { session } = require('electron');

const TENANT = 'https://consstellaire.sharepoint.com';
const PARTITION = 'persist:sharepoint'; // keeps Microsoft sign-in between launches
const JSON_NM = 'application/json;odata=nometadata';

const sp = () => session.fromPartition(PARTITION);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const odataStr = (s) => String(s).replace(/'/g, "''");

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

// Requests a site-scoped X-RequestDigest.
async function getDigestFor(siteUrl) {
  const r = await spFetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST', credentials: 'include',
    headers: { Accept: JSON_NM },
  });
  let j;
  try { j = await r.json(); } catch { j = null; }
  if (!r.ok || !j || !j.FormDigestValue) throw Object.assign(new Error('Not signed in'), { auth: true });
  return j.FormDigestValue;
}
const getDigest = () => getDigestFor(TENANT);

// Signed-in user's email and name.
async function whoAmI() {
  const r = await spFetch(`${TENANT}/_api/web/currentuser?$select=Email,Title,UserPrincipalName`, { credentials: 'include', headers: { Accept: JSON_NM } });
  if (!r.ok) throw Object.assign(new Error(`HTTP ${r.status}`), { auth: r.status === 401 || r.status === 403 });
  const j = await r.json();
  return { email: (j.Email || j.UserPrincipalName || '').toLowerCase(), name: j.Title || '' };
}

// Resolves an email to its UPN. Falls back to the email.
async function resolveUpn(email, digest) {
  try {
    const r = await spFetch(`${TENANT}/_api/web/ensureuser`, {
      method: 'POST', credentials: 'include',
      headers: { Accept: JSON_NM, 'Content-Type': JSON_NM, 'X-RequestDigest': digest },
      body: JSON.stringify({ logonName: email }),
    });
    const j = await r.json();
    return j.UserPrincipalName || (j.LoginName || '').split('|').pop() || email;
  } catch { return email; }
}

// Gets a user's Entra object ID from their user profile.
async function getObjectId({ email, key }) {
  const claim = key && key.includes('|') ? key : `i:0#.f|membership|${email}`;
  const r = await spFetch(
    `${TENANT}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(odataStr(claim))}'`,
    { credentials: 'include', headers: { Accept: JSON_NM } }
  );
  if (r.status === 401 || r.status === 403) throw Object.assign(new Error(`HTTP ${r.status}`), { auth: true });
  let j; try { j = await r.json(); } catch {}
  const prop = ((j && j.UserProfileProperties) || []).find((p) => p.Key === 'msOnline-ObjectId');
  return (prop && prop.Value) || null;
}

module.exports = { TENANT, PARTITION, JSON_NM, sp, sleep, odataStr, spFetch, getDigestFor, getDigest, whoAmI, resolveUpn, getObjectId };
