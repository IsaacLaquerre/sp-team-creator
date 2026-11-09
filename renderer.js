// Renderer for index.html: UI, translations, people pickers, sites list and admin panel.
const ICON_PIN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5"/><path d="M9 10.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24V16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V7a1 1 0 0 1 1-1 2 2 0 0 0 0-4H8a2 2 0 0 0 0 4 1 1 0 0 1 1 1z"/></svg>';
const ICON_SYNC ='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16"/><path d="M16 16h5v5"/></svg>';

const S = {
  fr: { heading:"Créer une équipe de projet", number:"Numéro", title:"Titre du projet",
    teamName:"Nom de l’équipe", alias:"Adresse du site (alias)", create:"Créer l’équipe", creating:"Création en cours…",
    language:"Langue", theme:"Thème", system:"Système", light:"Clair", dark:"Sombre", settings:"Paramètres",
    noApi:"Cette page doit être lancée par l’application Electron (run.bat ou npm start), pas ouverte dans un navigateur.", errNum:"Le numéro de projet doit comporter exactement 3 chiffres.", errTitle:"Le titre du projet est requis.",
    errAlias:"Le titre ne contient aucun caractère utilisable pour l’adresse.",
    signin:"Connexion à SharePoint requise…", signinFail:"Connexion annulée. Réessayez.",
    created:"Équipe créée.", prov:"Création lancée, le site sera prêt dans quelques instants.",
    exists:"Une équipe avec cette adresse existe déjà.", failed:"Échec : ", openSite:"Ouvrir le site",
    consentTitle:"Autoriser la connexion à Microsoft 365 ?", consentIntro:"Cette application doit se connecter à votre compte Microsoft pour continuer. En poursuivant, vous acceptez que l’application :",
    consentItem1Title:"Vous connecte", consentItem1:"ouvre une fenêtre de connexion Microsoft officielle et garde la session active pour les prochaines utilisations.",
    consentItem2Title:"Crée des sites", consentItem2:"envoie des requêtes à consstellaire.sharepoint.com en votre nom pour créer des équipes de projet.",
    consentItem3Title:"Lit l’annuaire", consentItem3:"consulte la liste des utilisateurs de votre organisation (nom, courriel) pour suggérer des propriétaires.",
    consentAgree:"J’accepte et je continue", consentDecline:"Annuler",
    createTab:"Créer", sitesTab:"Sites",
    owners:"Propriétaires", members:"Membres",
    ownersSearching:"Recherche…", ownersEmpty:"Aucun résultat.", ownersAuth:"Connexion requise pour chercher des utilisateurs.",
    historyEmpty:"Aucun site créé pour l’instant.",
    histOpen:"Ouvrir le site", histPending:"Site en cours de création",
    sortAsc:"Nom, A → Z", sortDesc:"Nom, Z → A", sortTitle:"Trier les sites", pin:"Épingler en haut", unpin:"Détacher",
    syncTitle:"Synchroniser avec OneDrive", syncFailed:"Échec de la synchronisation : ",
    syncLaunched:"Demande envoyée à OneDrive — vérifiez si une fenêtre OneDrive s’est ouverte. Rien ne garantit que la synchronisation a réellement démarré.",
    teamifyWorking:"Ajout du clavardage d’équipe en temps réel…", teamifyOk:"Clavardage d’équipe ajouté.", teamifyFail:"Impossible d’ajouter le clavardage d’équipe : ",
    membersWorking:"Ajout des membres…", membersOk:"Membres ajoutés.", membersFail:"Impossible d’ajouter certains membres : ",
    updateTitle:"Mise à jour prête", updateBody:"La version {v} a été téléchargée. L’application va redémarrer pour l’installer.",
    updateLink:"Voir la publication sur GitHub", updateLater:"Plus tard", updateRestart:"Redémarrer maintenant",
    updateNow:"Mettre à jour maintenant", updateNowTitle:"Installer la version {v} et redémarrer",
    adminOpen:"Installations", adminTitle:"Installations de l’application", adminRefresh:"Actualiser", adminClose:"Fermer",
    adminName:"Nom", adminEmail:"Courriel", adminRole:"Rôle", adminComputer:"Ordinateur", adminVersion:"Version", adminLastSeen:"Dernière activité",
    adminLoading:"Chargement…", adminEmpty:"Aucune installation signalée pour l’instant.", adminFail:"Impossible de lire la liste : ",
    adminSummary:"{n} installation(s), {o} à mettre à jour (dernière version : {v})" },
  en: { heading:"Create a project team", number:"Number", title:"Project title",
    teamName:"Team name", alias:"Site address (alias)", create:"Create team", creating:"Creating…",
    language:"Language", theme:"Theme", system:"System", light:"Light", dark:"Dark", settings:"Settings",
    noApi:"This page must be launched by the Electron app (run.bat or npm start), not opened in a browser.", errNum:"The project number must be exactly 3 digits.", errTitle:"The project title is required.",
    errAlias:"The title has no characters usable in the site address.",
    signin:"SharePoint sign-in required…", signinFail:"Sign-in cancelled. Try again.",
    created:"Team created.", prov:"Creation started, the site will be ready shortly.",
    exists:"A team with this address already exists.", failed:"Failed: ", openSite:"Open site",
    consentTitle:"Allow connection to Microsoft 365?", consentIntro:"This app needs to sign in to your Microsoft account to continue. By proceeding, you agree that the app will:",
    consentItem1Title:"Sign you in", consentItem1:"open an official Microsoft sign-in window and keep that session active for future use.",
    consentItem2Title:"Create sites", consentItem2:"send requests to consstellaire.sharepoint.com on your behalf to create project teams.",
    consentItem3Title:"Read the directory", consentItem3:"look up your organization's users (name, email) to suggest owners.",
    consentAgree:"I agree, continue", consentDecline:"Cancel",
    createTab:"Create", sitesTab:"Sites",
    owners:"Owners", members:"Members",
    ownersSearching:"Searching…", ownersEmpty:"No results.", ownersAuth:"Sign-in needed to search people.",
    historyEmpty:"No sites created yet.",
    histOpen:"Open site", histPending:"Site still provisioning",
    sortAsc:"Name, A → Z", sortDesc:"Name, Z → A", sortTitle:"Sort sites", pin:"Pin to top", unpin:"Unpin",
    syncTitle:"Sync with OneDrive", syncFailed:"Sync failed: ",
    syncLaunched:"Request sent to OneDrive — check whether a OneDrive window opened. There's no way to confirm from here that syncing actually started.",
    teamifyWorking:"Adding real-time team chat…", teamifyOk:"Team chat added.", teamifyFail:"Couldn't add team chat: ",
    membersWorking:"Adding members…", membersOk:"Members added.", membersFail:"Couldn't add some members: ",
    updateTitle:"Update ready", updateBody:"Version {v} has been downloaded. The app will restart to install it.",
    updateLink:"View the release on GitHub", updateLater:"Later", updateRestart:"Restart now",
    updateNow:"Update now", updateNowTitle:"Install version {v} and restart",
    adminOpen:"Installations", adminTitle:"App installations", adminRefresh:"Refresh", adminClose:"Close",
    adminName:"Name", adminEmail:"Email", adminRole:"Role", adminComputer:"Computer", adminVersion:"Version", adminLastSeen:"Last seen",
    adminLoading:"Loading…", adminEmpty:"No installs reported yet.", adminFail:"Couldn't read the list: ",
    adminSummary:"{n} install(s), {o} out of date (latest: {v})" },
};
const $ = (id) => document.getElementById(id);
const store = { get:(k) => { try { return localStorage.getItem(k) || 'system'; } catch { return 'system'; } },
                set:(k,v) => { try { localStorage.setItem(k,v); } catch {} } };
let langPref = store.get('lang'), themePref = store.get('theme');
const lang = () => langPref !== 'system' ? langPref : (navigator.language || 'fr').toLowerCase().startsWith('fr') ? 'fr' : 'en';
const t = (k) => S[lang()][k];

function applyTheme(){ const r = document.documentElement; themePref === 'system' ? r.removeAttribute('data-theme') : r.setAttribute('data-theme', themePref); }
function applyLang(){
  document.documentElement.lang = lang();
  document.querySelectorAll('[data-i]').forEach((el) => el.textContent = t(el.dataset.i));
  $('gear').title = $('gear').ariaLabel = t('settings');
  $('go').textContent = $('go').disabled ? t('creating') : t('create');
  renderHistory();
  renderMsg();
  Object.keys(statusState).forEach(renderStatus);
  if (adminData) renderInstalls();
  renderUpdateNow();
}

// ---------- tabs ----------
function showTab(name){
  const isCreate = name === 'create';
  $('tabCreateBtn').classList.toggle('active', isCreate);
  $('tabHistoryBtn').classList.toggle('active', !isCreate);
  $('viewCreate').classList.toggle('active', isCreate);
  $('viewHistory').classList.toggle('active', !isCreate);
  if (!isCreate) {
    renderHistory(); // shows last known states instantly, no flicker
    recheckAllSynced(); // background refresh, only re-renders sites whose result actually changed
    ensureRecentSites();
  }
}
$('tabCreateBtn').addEventListener('click', () => showTab('create'));
$('tabHistoryBtn').addEventListener('click', () => showTab('history'));

// Matches the wide layout media query. Both panels are visible there.
const wideLayout = matchMedia('(min-width:1134px)');
const sitesVisible = () => wideLayout.matches || $('viewHistory').classList.contains('active');
wideLayout.addEventListener('change', () => { if (wideLayout.matches) { renderHistory(); recheckAllSynced(); ensureRecentSites(); } });

// ---------- naming rules ----------
// The title keeps only letters, digits and underscores (accents are flattened: é → e).
const ALIAS_SEPARATOR = '-'; // site addresses can't contain spaces, so " - " becomes "-"
const cleanTitle = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9_]/g, '');
const names = () => {
  const num = $('num').value, title = $('ttl').value.trim();
  return { num, title, displayName: `${num} - ${title}`, alias: `${num}${ALIAS_SEPARATOR}${cleanTitle(title)}`.slice(0, 64) };
};
function preview(){ const n = names(); $('pName').textContent = n.title || n.num ? n.displayName : '—'; $('pAlias').textContent = n.title || n.num ? n.alias : '—'; }

function say(html, cls){ const m = $('msg'); m.className = cls || ''; m.textContent = ''; if (typeof html === 'string') m.textContent = html; else m.append(html); }

// #msg remembers what it's showing, so a language switch re-translates it.
let msgState = null; // { kind: 'key'|'failed'|'success', ...details } or null
function setMsg(kind, opts){ msgState = { kind, ...(opts || {}) }; renderMsg(); }
function clearMsg(){ msgState = null; say('', ''); setStatus('teamifyMsg', null); setStatus('membersMsg', null); }

// Status lines under #msg. Re-rendered by applyLang.
const statusState = {};
function setStatus(id, key, error){ statusState[id] = key ? { key, error } : null; renderStatus(id); }
function renderStatus(id){
  const el = $(id); if (!el) return;
  const s = statusState[id];
  el.textContent = s ? t(s.key) + (s.error || '') : '';
  el.className = s && s.error ? 'hist-meta err' : 'hist-meta';
}
function renderMsg(){
  if (!msgState) return say('', '');
  if (msgState.kind === 'success') {
    const box = document.createElement('span');
    box.append(t(msgState.prov ? 'prov' : 'created') + ' ');
    if (msgState.url) {
      const a = document.createElement('a'); a.textContent = t('openSite');
      a.onclick = () => window.api.open(msgState.url);
      box.append(a);
    }
    return say(box, 'ok');
  }
  if (msgState.kind === 'failed') return say(t('failed') + msgState.error, 'err');
  if (msgState.kind === 'syncFailed') return say(t('syncFailed') + msgState.error, 'err');
  say(t(msgState.kind), msgState.cls || ''); // cls: '' neutral, 'err' error
}

// Runs after creation, with its own status line separate from #msg.
function teamifyInBackground(siteUrl){
  if (!siteUrl || !window.api) return setStatus('teamifyMsg', null);
  setStatus('teamifyMsg', 'teamifyWorking');
  (async () => {
    const res = await window.api.teamifyGroup({ siteUrl });
    if (res.ok) setStatus('teamifyMsg', 'teamifyOk');
    else setStatus('teamifyMsg', 'teamifyFail', res.error || '?');
  })();
}

// Adds members to the group after creation.
function addMembersAfterCreate(siteUrl, groupId, memberList){
  if (!memberList.length || !groupId || !window.api) return setStatus('membersMsg', null);
  setStatus('membersMsg', 'membersWorking');
  (async () => {
    const res = await window.api.addMembers({ siteUrl, groupId, members: memberList });
    if (res.ok) setStatus('membersMsg', 'membersOk');
    else setStatus('membersMsg', 'membersFail', res.error || '?');
  })();
}

// ---------- consent gate ----------
// Requires explicit consent before sign-in, remembered locally. Shares one
// loginPromise across simultaneous callers.
let loginPromise = null;
function doLogin(){
  if (!window.api) return Promise.resolve(false);
  if (!loginPromise) loginPromise = window.api.login().finally(() => { loginPromise = null; });
  return loginPromise;
}

function hasConsent(){ try { return localStorage.getItem('msConsent') === '1'; } catch { return false; } }
function setConsent(v){ try { localStorage.setItem('msConsent', v ? '1' : '0'); } catch {} }
function askConsent(){
  return new Promise((resolve) => {
    if (hasConsent()) return resolve(true);
    applyLang();
    $('overlay').classList.add('open');
    const cleanup = () => { $('overlay').classList.remove('open'); $('consentAgree').onclick = null; $('consentDecline').onclick = null; };
    $('consentAgree').onclick = () => { setConsent(true); cleanup(); resolve(true); };
    $('consentDecline').onclick = () => { cleanup(); resolve(false); };
  });
}

// ---------- people pickers: shared search + chips logic ----------
// One factory, two instances (owners, members). list is mutated in place.
// Picking someone already in otherList moves them out of it.
function makePicker(list, ids, otherList){
  let seq = 0;
  function render(){
    const box = $(ids.chips);
    box.textContent = '';
    list.forEach((p, i) => {
      const chip = document.createElement('span'); chip.className = 'chip';
      const label = document.createElement('span'); label.textContent = p.name || p.email;
      const rm = document.createElement('button'); rm.type = 'button'; rm.textContent = '×'; rm.title = p.email;
      rm.addEventListener('click', () => { list.splice(i, 1); render(); });
      chip.append(label, rm); box.append(chip);
    });
  }
  function closeResults(){ $(ids.results).classList.remove('open'); $(ids.results).textContent = ''; }
  function showResultsMessage(text){
    const box = $(ids.results);
    box.textContent = '';
    const d = document.createElement('div'); d.className = 'result empty'; d.textContent = text;
    box.append(d); box.classList.add('open');
  }
  function showResults(people){
    const box = $(ids.results);
    box.textContent = '';
    const fresh = people.filter((p) => !list.some((o) => o.email.toLowerCase() === p.email.toLowerCase()));
    if (!fresh.length) return showResultsMessage(t('ownersEmpty'));
    fresh.forEach((p) => {
      const btn = document.createElement('button'); btn.type = 'button'; btn.className = 'result';
      const nm = document.createElement('span'); nm.className = 'rname'; nm.textContent = p.name;
      const em = document.createElement('span'); em.className = 'remail'; em.textContent = p.email;
      btn.append(nm, em);
      btn.addEventListener('click', () => {
        const dup = otherList.findIndex((o) => o.email.toLowerCase() === p.email.toLowerCase());
        if (dup >= 0) otherList.splice(dup, 1);
        list.push(p); renderChips();
        $(ids.input).value = ''; closeResults(); $(ids.input).focus();
      });
      box.append(btn);
    });
    box.classList.add('open');
  }
  let debounce;
  async function onInput(){
    const q = $(ids.input).value.trim();
    clearTimeout(debounce);
    if (q.length < 2) return closeResults();
    showResultsMessage(t('ownersSearching'));
    const mySeq = ++seq;
    debounce = setTimeout(async () => {
      if (!window.api) return;
      if (!(await askConsent())) return closeResults();
      let res = await window.api.searchPeople(q);
      if (!res.ok && res.auth) {
        setMsg('signin');
        if (!(await doLogin())) { clearMsg(); return showResultsMessage(t('ownersAuth')); }
        clearMsg();
        res = await window.api.searchPeople(q);
      }
      if (mySeq !== seq) return; // a newer query superseded this one
      if (!res.ok) return showResultsMessage(res.error || t('ownersEmpty'));
      showResults(res.people || []);
    }, 250);
  }
  $(ids.input).addEventListener('input', onInput);
  document.addEventListener('click', (e) => { if (!e.target.closest('.' + ids.field)) closeResults(); });
  return { render, closeResults };
}

// Pre-filled on every new team. Removable via the chip's x button.
const DEFAULT_OWNERS = [
  { email: 'edionne@stellaireconstruction.com' },
  { email: 'sdionne@stellaireconstruction.com' },
  { email: 'cdionne@stellaireconstruction.com' },
  { email: 'claquerre@stellaireconstruction.com' },
  { email: 'vlamontagne@stellaireconstruction.com' },
];
// Default owner names and keys, cached locally by lowercase email.
const OWNER_NAMES_KEY = 'defaultOwnerNames';
function applyOwnerNames(map){
  [DEFAULT_OWNERS, owners, members].forEach((l) => l.forEach((o) => {
    const n = map[o.email.toLowerCase()];
    if (n && !o.name) { o.name = n.name; o.key = n.key; }
  }));
  renderChips();
}
// Looks up each default owner's name with the people search.
async function resolveDefaultOwnerNames(){
  if (!window.api) return;
  let map = {};
  try { map = JSON.parse(localStorage.getItem(OWNER_NAMES_KEY) || '{}'); } catch {}
  for (const o of DEFAULT_OWNERS) {
    const res = await window.api.searchPeople(o.email);
    const p = res.ok && (res.people || []).find((x) => x.email.toLowerCase() === o.email.toLowerCase());
    if (p) map[o.email.toLowerCase()] = { name: p.name, key: p.key || '' };
  }
  try { localStorage.setItem(OWNER_NAMES_KEY, JSON.stringify(map)); } catch {}
  applyOwnerNames(map);
}

let owners = DEFAULT_OWNERS.map((o) => ({ ...o })); // { name, email, key }
let members = []; // added as plain members, not owners

const ownersPicker = makePicker(owners, { chips: 'ownerChips', input: 'ownerInput', results: 'ownerResults', field: 'owners-field' }, members);
const membersPicker = makePicker(members, { chips: 'memberChips', input: 'memberInput', results: 'memberResults', field: 'members-field' }, owners);
function renderChips(){ ownersPicker.render(); membersPicker.render(); }
try { applyOwnerNames(JSON.parse(localStorage.getItem(OWNER_NAMES_KEY) || '{}')); } catch {}
function closeAllResults(){ ownersPicker.closeResults(); membersPicker.closeResults(); }

// ---------- sites (history) ----------
// App-created sites, plus (below) ones pulled from SharePoint's "recent sites" API.
function loadHistory(){ try { return JSON.parse(localStorage.getItem('teamHistory') || '[]'); } catch { return []; } }
function saveHistory(list){ try { localStorage.setItem('teamHistory', JSON.stringify(list.slice(0, 200))); } catch {} }
function addHistory(entry){
  // Drops any existing record for this URL first.
  const list = loadHistory().filter((e) => e.url !== entry.url);
  list.unshift(entry);
  saveHistory(list);
}
function removeHistory(url){ saveHistory(loadHistory().filter((e) => e.url !== url)); }

// Looks up entries with no URL by group ID, every 10s while provisioning.
// Fills in the URL when ready, removes the entry when the group is gone.
const pendingChecked = new Set();
function resolvePending(entry){
  if (entry.url || !entry.groupId || pendingChecked.has(entry.groupId) || !window.api) return;
  pendingChecked.add(entry.groupId);
  (async () => {
    const res = await window.api.siteStatus({ groupId: entry.groupId });
    if (!res.ok) return;
    const list = loadHistory();
    if (res.status === 2 && res.url) list.forEach((e) => { if (e.groupId === entry.groupId) e.url = res.url; });
    else if (res.gone) list.splice(0, list.length, ...list.filter((e) => e.groupId !== entry.groupId));
    else {
      if (res.status === 1) setTimeout(() => { pendingChecked.delete(entry.groupId); renderHistory(); }, 10000);
      return;
    }
    saveHistory(list);
    renderHistory();
  })();
}
const normalizeUrl = (u) => (u || '').replace(/\/$/, '').toLowerCase();

// One-time cleanup of an old "synced" flag. No longer written.
try { localStorage.removeItem('syncedSites'); } catch {}

function aliasFromUrl(url){
  try { const u = new URL(url); const segs = u.pathname.split('/').filter(Boolean); return segs[segs.length - 1] || u.hostname; }
  catch { return url; }
}

// Supplement to the app-created list, loaded once per launch. null means not loaded yet.
let recentSites = null;
async function ensureRecentSites(){
  if (recentSites !== null || !window.api) return;
  let res = await window.api.listRecentSites();
  if (!res.ok && res.auth) {
    if (!(await askConsent())) { recentSites = []; return; }
    if (!(await doLogin())) { recentSites = []; return; }
    res = await window.api.listRecentSites();
  }
  recentSites = res.ok ? res.sites : [];
  renderHistory();
}

// Ghost-site detection: drops stale "recent sites" entries whose site was deleted elsewhere.
// Map<normalizedUrl, 'checking' | 'exists' | 'gone' | 'error'>
const existsCache = new Map();
function ensureExists(url){
  const key = normalizeUrl(url);
  if (!url || existsCache.has(key) || !window.api) return;
  existsCache.set(key, 'checking');
  (async () => {
    let res = await window.api.checkSiteExists({ siteUrl: url });
    if (!res.ok && res.auth) {
      if (!(await doLogin())) { existsCache.set(key, 'error'); renderHistory(); return; }
      res = await window.api.checkSiteExists({ siteUrl: url });
    }
    if (res.ok) existsCache.set(key, res.exists ? 'exists' : 'gone');
    else existsCache.set(key, 'error');
    renderHistory();
  })();
}

// Holds resolved sync-check results. In-flight checks are tracked separately in syncHeuristicInFlight.
// Map<normalizedUrl, boolean>
const syncHeuristicCache = new Map();
const syncHeuristicInFlight = new Set();
function ensureSyncHeuristic(url, num, title, alias, force){
  const key = normalizeUrl(url);
  if (!url || !window.api) return;
  if (!force && syncHeuristicCache.has(key)) return;
  if (syncHeuristicInFlight.has(key)) return;
  syncHeuristicInFlight.add(key);
  (async () => {
    const res = await window.api.checkSynced({ siteUrl: url, num, title, alias });
    syncHeuristicInFlight.delete(key);
    const next = !!(res.ok && res.synced);
    const changed = syncHeuristicCache.get(key) !== next;
    syncHeuristicCache.set(key, next);
    if (changed) renderHistory();
  })();
}

// Sites currently on screen.
let lastRenderedList = [];

// Rechecks sync status for sites on screen, updating only those whose result changed.
function recheckAllSynced(){
  if (!sitesVisible()) return;
  for (const e of lastRenderedList) {
    ensureSyncHeuristic(e.url, e.num, e.displayName, e.alias, true);
  }
}
setInterval(recheckAllSynced, 20000);

// Fires immediately when main.js's file watcher sees a change.
if (window.api && window.api.onSyncSettingsChanged) window.api.onSyncSettingsChanged(recheckAllSynced);

// Downloaded update waiting to install, or null. "Later" reveals Update now in settings.
let pendingUpdate = null;
function renderUpdateNow(){
  const b = $('updateNow');
  b.hidden = !(pendingUpdate && pendingUpdate.postponed);
  b.textContent = t('updateNow');
  b.title = pendingUpdate ? t('updateNowTitle').replace('{v}', pendingUpdate.version) : '';
}
function showUpdateReady(version, url){
  pendingUpdate = { version, url, postponed: false };
  $('updateBody').textContent = t('updateBody').replace('{v}', version);
  $('updateLink').onclick = () => window.api.openExternal(url);
  $('updateLater').onclick = () => { $('updateOverlay').classList.remove('open'); pendingUpdate.postponed = true; renderUpdateNow(); };
  $('updateRestart').onclick = () => window.api.quitAndInstall();
  $('updateOverlay').classList.add('open');
  renderUpdateNow();
}
$('updateNow').addEventListener('click', () => window.api && window.api.quitAndInstall());
if (window.api && window.api.onUpdateReady) window.api.onUpdateReady((_e, { version, url }) => showUpdateReady(version, url));

// ---------- installations panel (admins only) ----------
// Compares "1.0.6-alpha" style versions. A release ranks above its prereleases.
function compareVersions(a, b){
  const parse = (v) => { const [n, pre] = String(v || '').split('-'); return { n: n.split('.').map((x) => parseInt(x, 10) || 0), pre: pre || '' }; };
  const A = parse(a), B = parse(b);
  for (let i = 0; i < 3; i++) if ((A.n[i] || 0) !== (B.n[i] || 0)) return (A.n[i] || 0) - (B.n[i] || 0);
  if (A.pre === B.pre) return 0;
  if (!A.pre) return 1;
  if (!B.pre) return -1;
  return A.pre.localeCompare(B.pre, undefined, { numeric: true });
}
let adminData = null; // { rows, latestVersion } or { error }
function renderInstalls(){
  const body = $('adminTable').tBodies[0]; body.textContent = '';
  if (!adminData) return;
  if (adminData.error) { $('adminStatus').textContent = t('adminFail') + adminData.error; $('adminStatus').className = 'hist-meta err'; return; }
  const rows = adminData.rows.slice();
  // Latest GitHub release, or the newest reported version if GitHub was unreachable.
  const latest = adminData.latestVersion || rows.reduce((m, r) => compareVersions(r.version, m) > 0 ? r.version : m, '');
  rows.sort((a, b) => compareVersions(b.version, a.version) || String(b.lastSeen).localeCompare(String(a.lastSeen)));
  const old = rows.filter((r) => compareVersions(r.version, latest) < 0).length;
  $('adminStatus').className = 'hist-meta';
  $('adminStatus').textContent = rows.length ? t('adminSummary').replace('{n}', rows.length).replace('{o}', old).replace('{v}', latest) : t('adminEmpty');
  rows.forEach((r) => {
    const tr = document.createElement('tr');
    const seen = r.lastSeen ? new Date(r.lastSeen).toLocaleString(lang() === 'fr' ? 'fr-CA' : 'en-CA', { dateStyle: 'medium', timeStyle: 'short' }) : '';
    [r.name, r.email, r.role, r.computer, r.version, seen].forEach((v, i) => {
      const td = document.createElement('td'); td.textContent = v;
      if (i === 4 && compareVersions(r.version, latest) < 0) td.className = 'old';
      tr.append(td);
    });
    body.append(tr);
  });
}
async function loadInstalls(){
  adminData = null; renderInstalls();
  $('adminStatus').className = 'hist-meta'; $('adminStatus').textContent = t('adminLoading');
  const res = await window.api.listInstalls();
  adminData = res.ok ? { rows: res.rows, latestVersion: res.latestVersion } : { error: res.error || '?' };
  renderInstalls();
}
function closeAdmin(){ $('adminOverlay').classList.remove('open'); }
$('adminOpen').addEventListener('click', () => {
  $('settings').classList.remove('open'); $('gear').setAttribute('aria-expanded', 'false');
  $('adminOverlay').classList.add('open'); $('adminClose').focus();
  loadInstalls();
});
$('adminRefresh').addEventListener('click', loadInstalls);
$('adminClose').addEventListener('click', closeAdmin);
$('adminOverlay').addEventListener('click', (e) => { if (e.target === $('adminOverlay')) closeAdmin(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && $('adminOverlay').classList.contains('open')) closeAdmin(); });

// Reports this install now and every 4 hours. Errors are ignored.
function startInstallReporting(){
  if (!window.api || !window.api.reportInstall) return;
  const report = () => window.api.reportInstall().catch(() => {});
  report();
  setInterval(report, 4 * 60 * 60 * 1000);
  window.api.whoAmI().then((me) => { if (me.ok && me.isAdmin) $('adminOpen').hidden = false; }).catch(() => {});
}

function iconButton(cls, svg, title, onClick, disabled){
  const b = document.createElement('button'); b.type = 'button'; b.className = 'icon-btn ' + cls;
  b.innerHTML = svg; b.title = title; b.setAttribute('aria-label', title);
  if (disabled) b.disabled = true; else b.addEventListener('click', onClick);
  return b;
}

// Sort direction ('asc' or 'desc') and pinned site URLs, stored locally.
let histSort = (() => { try { return localStorage.getItem('histSort') === 'desc' ? 'desc' : 'asc'; } catch { return 'asc'; } })();
let pinned = (() => { try { return new Set(JSON.parse(localStorage.getItem('pinnedSites') || '[]')); } catch { return new Set(); } })();
function savePinned(){ try { localStorage.setItem('pinnedSites', JSON.stringify([...pinned])); } catch {} }
function togglePin(key){ if (pinned.has(key)) pinned.delete(key); else pinned.add(key); savePinned(); renderHistory(); }
// ---------- sort dropdown ----------
const SORTS = [['asc', 'sortAsc'], ['desc', 'sortDesc']];
const ICON_CHECK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 6 9 17l-5-5"/></svg>';
// Builds the options once, then only updates labels and selection.
SORTS.forEach(([v]) => {
  const o = document.createElement('button'); o.type = 'button'; o.className = 'dd-opt'; o.role = 'option'; o.dataset.sort = v;
  o.append(document.createElement('span')); o.insertAdjacentHTML('beforeend', ICON_CHECK);
  o.addEventListener('click', () => { setSort(v); closeSortMenu(); $('histSort').focus(); });
  $('histSortMenu').append(o);
});
function renderSortButton(){
  $('histSortLabel').textContent = t(SORTS.find(([v]) => v === histSort)[1]);
  $('histSort').title = t('sortTitle');
  $('histSortMenu').querySelectorAll('.dd-opt').forEach((o) => {
    o.firstChild.textContent = t(SORTS.find(([v]) => v === o.dataset.sort)[1]);
    o.setAttribute('aria-selected', String(o.dataset.sort === histSort));
  });
}
function setSort(v){
  histSort = v;
  try { localStorage.setItem('histSort', v); } catch {}
  renderHistory();
}
function openSortMenu(){
  $('histSortMenu').classList.add('open'); $('histSort').setAttribute('aria-expanded', 'true');
  const sel = $('histSortMenu').querySelector('[aria-selected=true]'); if (sel) sel.focus();
}
function closeSortMenu(){ $('histSortMenu').classList.remove('open'); $('histSort').setAttribute('aria-expanded', 'false'); }
$('histSort').addEventListener('click', () => $('histSortMenu').classList.contains('open') ? closeSortMenu() : openSortMenu());
$('histSortMenu').addEventListener('keydown', (e) => {
  const opts = [...$('histSortMenu').querySelectorAll('.dd-opt')];
  const i = opts.indexOf(document.activeElement);
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); opts[(i + (e.key === 'ArrowDown' ? 1 : -1) + opts.length) % opts.length].focus(); }
  else if (e.key === 'Escape') { closeSortMenu(); $('histSort').focus(); }
  else if (e.key === 'Tab') closeSortMenu();
});
document.addEventListener('click', (e) => { if (!e.target.closest('.dd')) closeSortMenu(); });

// Pinned sites first, then by name in the chosen direction.
function sortSites(list){
  const dir = histSort === 'desc' ? -1 : 1;
  return list.slice().sort((a, b) => {
    const pa = pinned.has(normalizeUrl(a.url)), pb = pinned.has(normalizeUrl(b.url));
    if (pa !== pb) return pa ? -1 : 1;
    return dir * (a.displayName || '').localeCompare(b.displayName || '', lang(), { numeric: true, sensitivity: 'base' });
  });
}

function renderHistory(){
  renderSortButton();
  // Dedupes by URL, keeping the first occurrence.
  const seenUrls = new Set();
  const own = loadHistory().filter((e) => {
    const key = normalizeUrl(e.url);
    if (!key || seenUrls.has(key)) return !key; // keep entries with no url (pending creates) as-is
    seenUrls.add(key);
    return true;
  });
  const ownUrls = new Set(own.map((e) => normalizeUrl(e.url)));
  // "recent sites" entries have no known creation date (alias shown instead).
  const extra = (recentSites || [])
    .filter((s) => !ownUrls.has(normalizeUrl(s.url)))
    .map((s) => ({ displayName: s.title || aliasFromUrl(s.url), alias: aliasFromUrl(s.url), url: s.url, date: null }));
  const list = sortSites(own.concat(extra));
  lastRenderedList = list;

  $('histEmpty').style.display = list.length ? 'none' : '';
  const box = $('histList'); box.textContent = '';
  list.forEach((e) => {
    const item = document.createElement('div'); item.className = 'hist-item';
    const main = document.createElement('div'); main.className = 'hist-main';
    const title = document.createElement('div'); title.className = 'hist-title'; title.textContent = e.displayName;
    const meta = document.createElement('div'); meta.className = 'hist-meta';

    resolvePending(e);
    ensureExists(e.url);
    ensureSyncHeuristic(e.url, e.num, e.displayName, e.alias);
    const key = normalizeUrl(e.url);

    // Shows the alias, never a timestamp.
    meta.textContent = e.alias;
    main.append(title, meta);

    const actions = document.createElement('div');
    actions.style.cssText = 'display:flex;gap:6px;align-items:center';
    if (existsCache.get(key) === 'gone') {
      // Drops a confirmed-gone (404) site from the list.
      removeHistory(e.url);
      return;
    }
    if (key) {
      const isPinned = pinned.has(key);
      const pinBtn = iconButton('pin' + (isPinned ? ' on' : ''), ICON_PIN, t(isPinned ? 'unpin' : 'pin'), () => togglePin(key));
      pinBtn.setAttribute('aria-pressed', String(isPinned));
      actions.append(pinBtn);
    }

    const heuristicSynced = syncHeuristicCache.get(key);
    const isSynced = heuristicSynced === true;
    // Button stays hidden until the sync check resolves.
    const checkSettled = heuristicSynced === true || heuristicSynced === false;

    if (!isSynced && checkSettled) {
      actions.append(iconButton('sync', ICON_SYNC, t('syncTitle'), async () => {
        if (!window.api) return;
        const res = await window.api.syncSite({ siteUrl: e.url });
        if (!res.ok) return setMsg('syncFailed', { error: res.error });
        // Retries the sync check a few times.
        setMsg('syncLaunched');
        [4000, 10000, 20000].forEach((delay) => setTimeout(() => {
          syncHeuristicCache.delete(key);
          renderHistory();
        }, delay));
      }));
    }

    const openBtn = document.createElement('button'); openBtn.type = 'button'; openBtn.className = 'hist-open';
    if (e.url) { openBtn.textContent = t('histOpen'); openBtn.addEventListener('click', () => window.api && window.api.open(e.url)); }
    else { openBtn.textContent = t('histPending'); openBtn.disabled = true; }
    actions.append(openBtn);

    item.append(main, actions); box.append(item);
  });
}

// ---------- submit ----------
async function submit(e){
  e.preventDefault();
  if (!window.api) return setMsg('noApi', { cls: 'err' });
  const n = names();
  if (!/^\d{3}$/.test(n.num)) return setMsg('errNum', { cls: 'err' });
  if (!n.title) return setMsg('errTitle', { cls: 'err' });
  if (!cleanTitle(n.title)) return setMsg('errAlias', { cls: 'err' });
  $('go').disabled = true; applyLang(); clearMsg();
  const ownerEmails = owners.map((o) => o.email);
  // Skips members who are also owners.
  const ownerSet = new Set(ownerEmails.map((e) => e.toLowerCase()));
  const memberList = members
    .filter((m) => !ownerSet.has(m.email.toLowerCase()))
    .map((m) => ({ email: m.email, key: m.key || '' }));
  try {
    if (!(await askConsent())) return setMsg('signinFail', { cls: 'err' });
    let res = await window.api.create({ displayName: n.displayName, alias: n.alias, owners: ownerEmails });
    if (!res.ok && res.auth) {
      setMsg('signin');
      if (!(await doLogin())) return setMsg('signinFail', { cls: 'err' });
      res = await window.api.create({ displayName: n.displayName, alias: n.alias, owners: ownerEmails });
    }
    if (!res.ok) return setMsg('failed', { error: res.error });
    if (res.status === 4) return setMsg('exists', { cls: 'err' });
    if (res.status === 3 || res.status === 0) return setMsg('failed', { error: 'SiteStatus ' + res.status });
    addHistory({ num: n.num, title: n.title, displayName: n.displayName, alias: n.alias, url: res.url || '', groupId: res.groupId || '', date: new Date().toISOString() });
    setMsg('success', { prov: res.status === 1, url: res.url });
    teamifyInBackground(res.url);
    addMembersAfterCreate(res.url, res.groupId, memberList);
    owners.length = 0; owners.push(...DEFAULT_OWNERS.map((o) => ({ ...o })));
    members.length = 0;
    renderChips();
  } finally { $('go').disabled = false; applyLang(); }
}

$('num').addEventListener('input', () => { $('num').value = $('num').value.replace(/\D/g, '').slice(0, 3); preview(); });
$('ttl').addEventListener('input', preview);
$('f').addEventListener('submit', submit);
$('gear').addEventListener('click', () => { const o = $('settings').classList.toggle('open'); $('gear').setAttribute('aria-expanded', o); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { $('settings').classList.remove('open'); closeAllResults(); $('consentDecline').click(); } });
$('lang').addEventListener('change', (e) => { langPref = e.target.value; store.set('lang', langPref); applyLang(); });
$('theme').addEventListener('change', (e) => { themePref = e.target.value; store.set('theme', themePref); applyTheme(); });

$('lang').value = langPref; $('theme').value = themePref;
applyTheme(); applyLang(); preview(); renderChips(); renderHistory();

// ---------- startup: ask consent right away, then sign in ----------
// Pops up a sign-in window only if the stored session has expired.
(async () => {
  if (!window.api) return; // noApi message is only relevant once the user tries to act
  if (!(await askConsent())) return;
  setMsg('signin');
  const check = await window.api.checkAuth();
  if (!check.ok && !(await doLogin())) { setMsg('signinFail', { cls: 'err' }); return; }
  clearMsg();
  resolveDefaultOwnerNames();
  startInstallReporting();
  if (wideLayout.matches) { recheckAllSynced(); ensureRecentSites(); } // Sites panel is already visible
})();
