/**
 * SEBN TN2 — Système RH de suivi des réclamations  |  BACKEND (Google Apps Script)
 * ---------------------------------------------------------------------------
 * Ce script est l'UNIQUE backend du système : il sert à la fois
 *   (a) le formulaire de réclamation  (actions "verify" et "submit")
 *   (b) l'application RH               (login, poll, take, close, ...)
 * Tout passe par le même Google Sheet :
 *
 *   formulaire  --(verify/submit : ajoute une ligne)-->  feuille "Reclamations"
 *   ce script   --(détecte la ligne, lui donne un ID REC-000123)-->  "RH_Dossiers"
 *   application RH  <--(lecture / actions)-->  ce script
 *   chaque action   --(horodatée côté serveur)-->  "RH_Journal"
 *
 * Fuseau horaire de référence : Africa/Tunis (toutes les heures sont celles du serveur,
 * jamais celles de l'ordinateur de l'agent).
 */

// ======================= CONFIGURATION (à adapter) ========================
var CONFIG = {
  // ID du Google Sheet qui contient les réclamations (la partie entre /d/ et /edit dans l'URL)
  SPREADSHEET_ID: 'COLLER_ICI_L_ID_DU_GOOGLE_SHEET',

  // Nom de la feuille alimentée par le formulaire. Laisser '' pour la détecter
  // automatiquement (feuille dont l'en-tête contient "Matricule" et "Zone de problème").
  SOURCE_SHEET_NAME: '',

  // Agents RH autorisés (matricule -> nom, d'après le fichier master data).
  AGENTS: {
    '01756': 'Hajji Khouloud',
    '04526': 'Ammar Sonia',
    '00600': 'Romdhani Ghada'
  },

  // Feuille de la liste des employés (export "Etat personnel"). Détectée par ses titres
  // "Code matricule" + "CIN" ; peut être forcée ici si besoin.
  MASTER_SHEET_NAME: 'MasterData',
  MAX_VERIFY_FAILS: 5,            // essais Matricule+CIN erronés avant blocage
  VERIFY_LOCK_SECONDS: 600,       // durée du blocage (10 min)

  TZ: 'Africa/Tunis',
  SESSION_SECONDS: 6 * 60 * 60,   // durée d'une session (6 h, maximum d'Apps Script)
  MIN_TEXT: 10,                   // longueur minimale du texte de clôture / rejet

  // Statut donné aux lignes DÉJÀ présentes dans la feuille au moment de l'installation
  // (anciennes réclamations). Elles reçoivent un ID mais n'encombrent pas la file "Nouvelle".
  LEGACY_STATUS: 'Historique',

  SHEET_DOSSIERS: 'RH_Dossiers',
  SHEET_JOURNAL: 'RH_Journal'
};

var STATUS = { NEW: 'Nouvelle', WIP: 'En cours', CLOSED: 'Clôturée', REJECTED: 'Rejetée' };
var PRIORITIES = ['Basse', 'Normale', 'Haute'];

var DOSSIER_HEADERS = [
  'ID', 'Date de soumission', 'Matricule', 'Département', 'Type', 'Shift', 'Area', 'Sub Area',
  'Qualification', 'Zone de problème', 'Catégorie', 'Sous-catégorie', 'Date et heure incident',
  'Description', 'Statut', 'Priorité', 'Agent responsable', 'Prise en charge le',
  'Clôture / rejet le', 'Clôture / rejet par', 'Traitement / motif', 'Dernière modification',
  'Ligne source', 'Enregistrée dans le système le'
];
var COL = {}; DOSSIER_HEADERS.forEach(function (h, i) { COL[h] = i; });
var JOURNAL_HEADERS = [
  'Horodatage (Tunis)', 'Horodatage (epoch ms)', 'ID réclamation', 'Matricule agent',
  'Action', 'Détail', 'Ancienne valeur', 'Nouvelle valeur', 'Commentaire', 'Appareil'
];
var CLIENT_LOG_ACTIONS = {
  PAGE_TABLEAU_DE_BORD: 1, PAGE_HISTORIQUE: 1, PAGE_JOURNAL: 1, OUVERTURE_DETAIL: 1,
  OUVERTURE_PAGE_CLOTURE: 1, OUVERTURE_PAGE_REJET: 1, ANNULATION: 1, EXPORT_HISTORIQUE: 1,
  EXPORT_JOURNAL: 1, FILTRE: 1, CLIC: 1
};

// ============================ POINTS D'ENTRÉE =============================
function doGet() {
  return json_({ ok: true, service: 'SEBN TN2 RH', time: nowStr_() });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Requête invalide.' }); }
  try {
    switch (req.action) {
      case 'verify':   return json_(verify_(req));
      case 'submit':   return json_(submit_(req));
      case 'login':    return json_(login_(req));
      case 'logout':   return json_(logout_(req));
      case 'poll':     return json_(poll_(req));
      case 'detail':   return json_(detail_(req));
      case 'take':     return json_(take_(req));
      case 'priority': return json_(setPriority_(req));
      case 'close':    return json_(finish_(req, STATUS.CLOSED));
      case 'reject':   return json_(finish_(req, STATUS.REJECTED));
      case 'history':  return json_(history_(req));
      case 'journal':  return json_(journal_(req));
      case 'log':      return json_(clientLog_(req));
      default:         return json_({ ok: false, error: 'Action inconnue.' });
    }
  } catch (err) {
    if (err && err.userMessage) return json_({ ok: false, error: err.userMessage, code: err.code });
    console.error(err && err.stack || err);
    return json_({ ok: false, error: 'Erreur serveur. Réessayez.' });
  }
}

/** À exécuter UNE fois à la main (menu ▶ Exécuter) : crée les feuilles, donne un ID à toutes les
 *  réclamations existantes et installe un déclencheur qui surveille le formulaire chaque minute. */
function installer() {
  var ss = ss_();
  var ctx = ctx_(ss);
  ensureSource_(ss);
  ingest_(ctx, true);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'surveiller') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('surveiller').timeBased().everyMinutes(1).create();
  Logger.log('Installation terminée. Réclamations enregistrées : ' + (ctx.dossiers.getLastRow() - 1));
}

/** Déclencheur chaque minute : même si personne n'est connecté, chaque nouvelle réclamation reçoit son ID. */
function surveiller() {
  var ctx = ctx_(ss_());
  ingest_(ctx, false);
}

// ============================== OUTILS ====================================
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fail_(msg, code) { var e = new Error(msg); e.userMessage = msg; e.code = code; return e; }
function nowStr_(d) { return Utilities.formatDate(d || new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'); }
function ss_() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
}
/** Un texte commençant par = + - @ serait interprété comme une formule par Google Sheets : on le force en texte. */
function protect_(v) { return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v; }
function cell_(v) {
  if (v instanceof Date) return nowStr_(v);
  return v == null ? '' : String(v);
}

function ctx_(ss) {
  var dossiers = ss.getSheetByName(CONFIG.SHEET_DOSSIERS);
  if (!dossiers) {
    dossiers = ss.insertSheet(CONFIG.SHEET_DOSSIERS);
    dossiers.getRange(1, 1, 1, DOSSIER_HEADERS.length).setValues([DOSSIER_HEADERS]).setFontWeight('bold').setBackground('#10205C').setFontColor('#FFFFFF');
    dossiers.setFrozenRows(1);
    dossiers.getRange(1, 1, dossiers.getMaxRows(), DOSSIER_HEADERS.length).setNumberFormat('@');
  }
  var journal = ss.getSheetByName(CONFIG.SHEET_JOURNAL);
  if (!journal) {
    journal = ss.insertSheet(CONFIG.SHEET_JOURNAL);
    journal.getRange(1, 1, 1, JOURNAL_HEADERS.length).setValues([JOURNAL_HEADERS]).setFontWeight('bold').setBackground('#2C1272').setFontColor('#FFFFFF');
    journal.setFrozenRows(1);
    journal.getRange(1, 1, journal.getMaxRows(), JOURNAL_HEADERS.length).setNumberFormat('@');
  }
  return { ss: ss, dossiers: dossiers, journal: journal };
}

function findSource_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i], n = sh.getName();
    if (n === CONFIG.SHEET_DOSSIERS || n === CONFIG.SHEET_JOURNAL) continue;
    if (CONFIG.SOURCE_SHEET_NAME) { if (n === CONFIG.SOURCE_SHEET_NAME) return sh; continue; }
    if (sh.getLastColumn() < 5) continue;
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
    if (h.indexOf('matricule') > -1 && h.indexOf('zonedeprobleme') > -1) return sh;
  }
  throw fail_('Feuille source des réclamations introuvable.', 'NO_SOURCE');
}

/** Associe chaque champ à sa colonne dans la feuille source d'après le titre de colonne. */
function sourceMap_(sh) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  function idx(pred) { for (var i = 0; i < h.length; i++) if (pred(h[i])) return i; return -1; }
  var m = {
    'Date de soumission': idx(function (x) { return x === 'date'; }),
    'Matricule': idx(function (x) { return x === 'matricule'; }),
    'Département': idx(function (x) { return x === 'departement'; }),
    'Type': idx(function (x) { return x === 'type'; }),
    'Shift': idx(function (x) { return x === 'shift'; }),
    'Area': idx(function (x) { return x === 'area'; }),
    'Sub Area': idx(function (x) { return x === 'subarea'; }),
    'Qualification': idx(function (x) { return x === 'qualification'; }),
    'Zone de problème': idx(function (x) { return x === 'zonedeprobleme'; }),
    'Catégorie': idx(function (x) { return x === 'categorie'; }),
    'Sous-catégorie': idx(function (x) { return x === 'souscategorie'; }),
    'Date et heure incident': idx(function (x) { return x.indexOf('incident') > -1; }),
    'Description': idx(function (x) { return x === 'description'; })
  };
  if (m['Matricule'] < 0 || m['Date de soumission'] < 0) throw fail_('Colonnes "Date" / "Matricule" introuvables dans la feuille source.', 'BAD_SOURCE');
  return m;
}

// ====================== JOURNAL (chaque clic est tracé) ===================
function log_(ctx, entries) {
  if (!entries.length) return;
  var d = new Date(), t = nowStr_(d), ms = d.getTime();
  var rows = entries.map(function (e) {
    return [t, ms, e.id || '', e.agent || '', e.action, protect_(e.detail || ''), protect_(e.old || ''), protect_(e.neu || ''), protect_(e.comment || ''), protect_(e.device || '')];
  });
  var sh = ctx.journal;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, JOURNAL_HEADERS.length).setValues(rows);
}

// ============== ENREGISTREMENT DES NOUVELLES RÉCLAMATIONS (IDs) ===========
function seq_() { return Number(PropertiesService.getScriptProperties().getProperty('seq') || 0); }
function bump_() { var p = PropertiesService.getScriptProperties(); var v = Number(p.getProperty('seq') || 0) + 1; p.setProperty('seq', String(v)); return v; }

/** Copie dans RH_Dossiers les lignes du formulaire pas encore vues et leur attribue un ID. */
function ingest_(ctx, firstInstall) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var props = PropertiesService.getScriptProperties();
    var src;
    try { src = findSource_(ctx.ss); } catch (e) { if (e.code === 'NO_SOURCE') return 0; throw e; }
    var last = src.getLastRow();
    var done = Number(props.getProperty('lastSourceRow') || 1);
    if (last <= done) return 0;
    var map = sourceMap_(src);
    var width = src.getLastColumn();
    var data = src.getRange(done + 1, 1, last - done, width).getValues();
    var nextNum = ctx.dossiers.getLastRow(); // ligne 1 = en-tête => ID suivant = nb de dossiers + 1
    var existing = nextNum - 1;
    var legacy = (firstInstall || props.getProperty('installed') !== '1') && existing === 0 && (last - 1) > 0;
    var now = nowStr_();
    var out = [], logs = [], r, rowObj;
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.every(function (c) { return c === '' || c == null; })) { continue; }
      rowObj = new Array(DOSSIER_HEADERS.length).fill('');
      var id = 'REC-' + ('000000' + (existing + out.length + 1)).slice(-6);
      rowObj[COL['ID']] = id;
      Object.keys(map).forEach(function (k) { if (map[k] > -1) rowObj[COL[k]] = protect_(cell_(row[map[k]])); });
      rowObj[COL['Statut']] = legacy ? CONFIG.LEGACY_STATUS : STATUS.NEW;
      rowObj[COL['Dernière modification']] = now;
      rowObj[COL['Ligne source']] = String(done + 1 + i);
      rowObj[COL['Enregistrée dans le système le']] = now;
      out.push(rowObj);
      if (!legacy) logs.push({ id: id, agent: 'SYSTEME', action: 'RECEPTION', detail: 'Réclamation reçue du formulaire (ligne ' + (done + 1 + i) + ') et enregistrée', neu: STATUS.NEW });
    }
    if (out.length) {
      ctx.dossiers.getRange(ctx.dossiers.getLastRow() + 1, 1, out.length, DOSSIER_HEADERS.length).setValues(out);
      if (legacy) logs.push({ id: '', agent: 'SYSTEME', action: 'IMPORT_INITIAL', detail: out.length + ' réclamations existantes importées avec un ID (statut ' + CONFIG.LEGACY_STATUS + ')' });
      log_(ctx, logs);
      bump_();
    }
    props.setProperty('lastSourceRow', String(last));
    props.setProperty('installed', '1');
    return out.length;
  } finally { lock.releaseLock(); }
}

// ============================== SESSIONS ==================================
function login_(req) {
  var ctx = ctx_(ss_());
  var m = String(req.matricule || '').trim();
  var dev = String(req.device || '').slice(0, 200);
  if (!CONFIG.AGENTS.hasOwnProperty(m)) {
    log_(ctx, [{ agent: m.slice(0, 20), action: 'CONNEXION_REFUSEE', detail: 'Matricule non autorisé', device: dev }]);
    throw fail_("Matricule non autorisé. Accès réservé aux agents RH habilités.", 'DENIED');
  }
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + token, m, CONFIG.SESSION_SECONDS);
  log_(ctx, [{ agent: m, action: 'CONNEXION', detail: 'Connexion réussie', device: dev }]);
  return { ok: true, token: token, matricule: m, name: CONFIG.AGENTS[m], agents: agentList_(), now: nowStr_(), nowMs: Date.now() };
}

function agentList_() {
  return Object.keys(CONFIG.AGENTS).map(function (k) { return { matricule: k, name: CONFIG.AGENTS[k] }; });
}

function auth_(req) {
  var cache = CacheService.getScriptCache();
  var m = req.token ? cache.get('s_' + req.token) : null;
  if (!m) throw fail_('Session expirée. Veuillez vous reconnecter.', 'SESSION');
  cache.put('s_' + req.token, m, CONFIG.SESSION_SECONDS);
  return m;
}

function logout_(req) {
  var cache = CacheService.getScriptCache();
  var m = req.token ? cache.get('s_' + req.token) : null;
  if (m) {
    cache.remove('s_' + req.token);
    log_(ctx_(ss_()), [{ agent: m, action: 'DECONNEXION', detail: req.reason === 'timeout' ? 'Déconnexion automatique' : 'Déconnexion volontaire', device: String(req.device || '').slice(0, 200) }]);
  }
  return { ok: true };
}

// =============================== LECTURE ==================================
function toObj_(row) {
  var o = {};
  DOSSIER_HEADERS.forEach(function (h, i) { o[h] = row[i] == null ? '' : String(row[i]); });
  return o;
}

function readAll_(ctx) {
  var n = ctx.dossiers.getLastRow();
  if (n < 2) return [];
  return ctx.dossiers.getRange(2, 1, n - 1, DOSSIER_HEADERS.length).getValues();
}

/** Appelé toutes les quelques secondes par l'application : renvoie du neuf seulement si quelque chose a changé. */
function poll_(req) {
  auth_(req);
  var ctx = ctx_(ss_());
  var cache = CacheService.getScriptCache();
  if (!cache.get('ingest_recent')) { cache.put('ingest_recent', '1', 2); ingest_(ctx, false); }
  var seq = seq_();
  var base = { ok: true, seq: seq, now: nowStr_(), nowMs: Date.now() };
  if (Number(req.seq) === seq) { base.changed = false; return base; }
  var rows = readAll_(ctx);
  var stats = { total: rows.length, Nouvelle: 0, 'En cours': 0, 'Clôturée': 0, 'Rejetée': 0, Historique: 0 };
  var open = [], i, st;
  for (i = 0; i < rows.length; i++) {
    st = rows[i][COL['Statut']];
    if (stats[st] != null) stats[st]++;
    if (st === STATUS.NEW || st === STATUS.WIP) open.push(i);
  }
  var recentFrom = Math.max(0, rows.length - 60);
  var pick = {};
  for (i = recentFrom; i < rows.length; i++) pick[i] = 1;
  open.forEach(function (k) { pick[k] = 1; });
  var items = Object.keys(pick).map(Number).sort(function (a, b) { return b - a; }).map(function (k) { return toObj_(rows[k]); });
  base.changed = true; base.stats = stats; base.items = items; base.agents = agentList_();
  return base;
}

function findRow_(ctx, id) {
  var m = /^REC-(\d{6})$/.exec(String(id || ''));
  if (!m) throw fail_('Identifiant invalide.', 'BAD_ID');
  var rowNum = Number(m[1]) + 1;
  if (rowNum <= ctx.dossiers.getLastRow() && ctx.dossiers.getRange(rowNum, 1).getValue() === id) return rowNum;
  var f = ctx.dossiers.getRange(2, 1, Math.max(1, ctx.dossiers.getLastRow() - 1), 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!f) throw fail_('Réclamation introuvable.', 'NOT_FOUND');
  return f.getRow();
}

function events_(ctx, id) {
  var n = ctx.journal.getLastRow();
  if (n < 2) return [];
  var ids = ctx.journal.getRange(2, 3, n - 1, 1).getValues();
  var hits = [];
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) hits.push(i + 2);
  return hits.map(function (r) {
    var v = ctx.journal.getRange(r, 1, 1, JOURNAL_HEADERS.length).getValues()[0];
    return { t: String(v[0]), ms: Number(v[1]), agent: v[3], action: v[4], detail: v[5], old: v[6], neu: v[7], comment: v[8] };
  });
}

function detail_(req) {
  var agent = auth_(req);
  var ctx = ctx_(ss_());
  var row = findRow_(ctx, req.id);
  var o = toObj_(ctx.dossiers.getRange(row, 1, 1, DOSSIER_HEADERS.length).getValues()[0]);
  log_(ctx, [{ id: req.id, agent: agent, action: 'OUVERTURE_DETAIL', detail: 'Fiche de la réclamation consultée', device: String(req.device || '').slice(0, 200) }]);
  return { ok: true, dossier: o, events: events_(ctx, req.id), now: nowStr_(), nowMs: Date.now() };
}

function history_(req) {
  auth_(req);
  var ctx = ctx_(ss_());
  var all = readAll_(ctx);
  var total = all.length;
  var off = Math.max(0, Number(req.offset) || 0), lim = Math.min(3000, Number(req.limit) || 1000);
  // du plus récent au plus ancien
  var from = Math.max(0, total - off - lim), to = Math.max(0, total - off);
  var items = all.slice(from, to).reverse().map(toObj_);
  return { ok: true, total: total, offset: off, items: items, now: nowStr_(), nowMs: Date.now() };
}

function journal_(req) {
  auth_(req);
  var ctx = ctx_(ss_());
  var n = ctx.journal.getLastRow() - 1;
  var off = Math.max(0, Number(req.offset) || 0), lim = Math.min(3000, Number(req.limit) || 1000);
  if (n < 1) return { ok: true, total: 0, items: [] };
  var to = n - off, from = Math.max(0, to - lim);
  if (to <= 0) return { ok: true, total: n, items: [] };
  var vals = ctx.journal.getRange(from + 2, 1, to - from, JOURNAL_HEADERS.length).getValues().reverse();
  var items = vals.map(function (v) {
    return { t: String(v[0]), ms: Number(v[1]), id: v[2], agent: v[3], action: v[4], detail: v[5], old: v[6], neu: v[7], comment: v[8], device: v[9] };
  });
  return { ok: true, total: n, offset: off, items: items };
}

// =============================== ACTIONS ==================================
function mutate_(req, fn) {
  var agent = auth_(req);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ctx = ctx_(ss_());
    var row = findRow_(ctx, req.id);
    var rng = ctx.dossiers.getRange(row, 1, 1, DOSSIER_HEADERS.length);
    var vals = rng.getValues()[0].map(function (v) { return v == null ? '' : String(v); });
    var logs = [];
    fn(vals, logs, agent);
    var now = nowStr_();
    vals[COL['Dernière modification']] = now;
    rng.setValues([vals.map(protect_)]);
    logs.forEach(function (l) { l.id = req.id; l.agent = agent; l.device = String(req.device || '').slice(0, 200); });
    log_(ctx, logs);
    bump_();
    return { ok: true, dossier: toObj_(vals), seq: seq_(), now: now, nowMs: Date.now() };
  } finally { lock.releaseLock(); }
}

function take_(req) {
  var assigned = String(req.agent || '');
  var prio = String(req.priority || '');
  if (!CONFIG.AGENTS.hasOwnProperty(assigned)) throw fail_("Choisissez l'agent RH qui ouvre la réclamation.", 'BAD_AGENT');
  if (PRIORITIES.indexOf(prio) < 0) throw fail_('Choisissez une priorité (Basse / Normale / Haute).', 'BAD_PRIO');
  return mutate_(req, function (v, logs) {
    if (v[COL['Statut']] !== STATUS.NEW) {
      throw fail_('Cette réclamation a déjà été prise en charge (' + v[COL['Statut']] + (v[COL['Agent responsable']] ? ' par ' + v[COL['Agent responsable']] : '') + ').', 'CONFLICT');
    }
    var now = nowStr_();
    v[COL['Statut']] = STATUS.WIP;
    v[COL['Priorité']] = prio;
    v[COL['Agent responsable']] = assigned;
    v[COL['Prise en charge le']] = now;
    logs.push({ action: 'AGENT_ASSIGNE', detail: 'Agent qui ouvre la réclamation : ' + assigned, neu: assigned });
    logs.push({ action: 'PRIORITE_DEFINIE', detail: 'Priorité définie', neu: prio });
    logs.push({ action: 'STATUT_CHANGE', detail: 'Réclamation prise en charge', old: STATUS.NEW, neu: STATUS.WIP });
  });
}

function setPriority_(req) {
  var prio = String(req.priority || '');
  if (PRIORITIES.indexOf(prio) < 0) throw fail_('Priorité invalide.', 'BAD_PRIO');
  return mutate_(req, function (v, logs) {
    if (v[COL['Statut']] !== STATUS.WIP) throw fail_("La priorité ne peut être modifiée que pour une réclamation « En cours ».", 'BAD_STATE');
    var old = v[COL['Priorité']];
    if (old === prio) return;
    v[COL['Priorité']] = prio;
    logs.push({ action: 'PRIORITE_MODIFIEE', detail: 'Priorité modifiée', old: old, neu: prio });
  });
}

function finish_(req, target) {
  var text = String(req.text || '').trim();
  var isClose = target === STATUS.CLOSED;
  if (text.length < CONFIG.MIN_TEXT) {
    throw fail_((isClose ? 'Décrivez comment la réclamation a été traitée et résolue' : 'Indiquez le motif du rejet') + ' (' + CONFIG.MIN_TEXT + ' caractères minimum).', 'BAD_TEXT');
  }
  if (text.length > 4000) throw fail_('Texte trop long (4000 caractères maximum).', 'BAD_TEXT');
  return mutate_(req, function (v, logs, agent) {
    if (v[COL['Statut']] !== STATUS.WIP) {
      throw fail_(v[COL['Statut']] === STATUS.NEW
        ? "La réclamation doit d'abord être prise en charge (« En cours »)."
        : 'Cette réclamation est déjà ' + v[COL['Statut']].toLowerCase() + '.', 'BAD_STATE');
    }
    var old = v[COL['Statut']];
    v[COL['Statut']] = target;
    v[COL['Clôture / rejet le']] = nowStr_();
    v[COL['Clôture / rejet par']] = agent;
    v[COL['Traitement / motif']] = text;
    logs.push({ action: isClose ? 'CLOTURE' : 'REJET', detail: isClose ? 'Réclamation clôturée — traitement décrit' : 'Réclamation rejetée — motif déclaré', old: old, neu: target, comment: text });
  });
}

function clientLog_(req) {
  var agent = auth_(req);
  var list = (req.events || []).slice(0, 20);
  var entries = [];
  list.forEach(function (e) {
    if (!CLIENT_LOG_ACTIONS[e.action]) return;
    entries.push({ id: /^REC-\d{6}$/.test(e.id || '') ? e.id : '', agent: agent, action: e.action, detail: String(e.detail || '').slice(0, 500), device: String(req.device || '').slice(0, 200) });
  });
  log_(ctx_(ss_()), entries);
  return { ok: true };
}


// ================== FORMULAIRE : vérification + enregistrement ===============
var SRC_HEADERS = ['Date', 'Matricule', 'Département', 'Type', 'Shift', 'Area', 'Sub Area', 'Qualification',
  'Zone de problème', 'Catégorie', 'Sous-catégorie', 'Date et heure incident', 'Description'];

/** Retourne la feuille source ; la crée (avec ses titres) si elle n'existe pas encore. */
function ensureSource_(ss) {
  try { return findSource_(ss); } catch (e) {
    if (e.code !== 'NO_SOURCE') throw e;
    var sh = ss.insertSheet(CONFIG.SOURCE_SHEET_NAME || 'Reclamations');
    sh.getRange(1, 1, 1, SRC_HEADERS.length).setValues([SRC_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
}

function digits_(v, n) { var d = String(v == null ? '' : v).replace(/\D/g, ''); return d ? ('00000000' + d).slice(-n) : ''; }
function clean_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

/** Cherche l'employé (matricule + CIN) dans la feuille MasterData. Retourne son profil ou null. */
function findEmployee_(ss, matricule, cin) {
  var sh = null, sheets = ss.getSheets(), i;
  if (CONFIG.MASTER_SHEET_NAME) sh = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (!sh) throw fail_("Liste des employés introuvable : importez le fichier « master data » dans une feuille « MasterData ».", 'NO_MASTER');
  var last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 2) return null;
  var all = sh.getRange(1, 1, last, width).getValues(), hr = -1, h;
  for (i = 0; i < Math.min(all.length, 30); i++) {
    h = all[i].map(norm_);
    if (h.indexOf('codematricule') > -1 && h.indexOf('cin') > -1) { hr = i; break; }
  }
  if (hr < 0) throw fail_('Titres « Code matricule » / « CIN » introuvables dans MasterData.', 'BAD_MASTER');
  h = all[hr].map(norm_);
  function col(pred) { for (var k = 0; k < h.length; k++) if (pred(h[k])) return k; return -1; }
  var c = {
    mat: col(function (x) { return x === 'codematricule'; }), cin: col(function (x) { return x === 'cin'; }),
    nom: col(function (x) { return x.indexOf('nomet') === 0 || x === 'nom'; }),
    dep: col(function (x) { return x.indexOf('departement') === 0; }),
    qual: col(function (x) { return x === 'qualification'; }),
    type: col(function (x) { return x === 'emploi'; }),
    area: col(function (x) { return x === 'area'; }),
    sub: col(function (x) { return x === 'subarea'; })
  };
  var m = digits_(matricule, 5), n = digits_(cin, 8);
  if (!m || !n) return null;
  for (i = hr + 1; i < all.length; i++) {
    if (digits_(all[i][c.mat], 5) === m) {
      if (digits_(all[i][c.cin], 8) !== n) return null;
      var g = function (k) { return k > -1 ? clean_(all[i][k]) : ''; };
      return { matricule: m, nom: g(c.nom), departement: g(c.dep), type: g(c.type), area: g(c.area), subArea: g(c.sub), qualification: g(c.qual) };
    }
  }
  return null;
}

/** Limite les essais Matricule+CIN erronés (anti-devinette du CIN). */
function checkVerifyLock_(matricule) {
  var cache = CacheService.getScriptCache(), k = 'vf_' + digits_(matricule, 5);
  var n = Number(cache.get(k) || 0);
  if (n >= CONFIG.MAX_VERIFY_FAILS) throw fail_('Trop de tentatives. Réessayez dans quelques minutes.', 'LOCKED');
  return { cache: cache, k: k, n: n };
}

function verify_(req) {
  var lk = checkVerifyLock_(req.matricule);
  var emp = findEmployee_(ss_(), req.matricule, req.cin);
  if (!emp) {
    lk.cache.put(lk.k, String(lk.n + 1), CONFIG.VERIFY_LOCK_SECONDS);
    throw fail_('Matricule ou CIN incorrect.', 'BAD_CREDENTIALS');
  }
  lk.cache.remove(lk.k);
  return { ok: true, nom: emp.nom, departement: emp.departement, type: emp.type, area: emp.area, subArea: emp.subArea, qualification: emp.qualification };
}

function safeCell_(v, max) {
  v = String(v == null ? '' : v).replace(/\r/g, '').trim().slice(0, max || 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;     // empêche l'injection de formules dans Google Sheets
}

function submit_(req) {
  var lk = checkVerifyLock_(req.matricule);
  var ss = ss_();
  var emp = findEmployee_(ss, req.matricule, req.cin);
  if (!emp) { lk.cache.put(lk.k, String(lk.n + 1), CONFIG.VERIFY_LOCK_SECONDS); throw fail_('Matricule ou CIN incorrect.', 'BAD_CREDENTIALS'); }
  var r = req.record || {};
  var need = [['shift', 'Shift'], ['zone', 'Zone de problème'], ['categorie', 'Catégorie'], ['sousCategorie', 'Sous-catégorie'], ['incidentDate', "Date et heure de l'incident"]];
  for (var i = 0; i < need.length; i++) if (!clean_(r[need[i][0]])) throw fail_('Champ obligatoire manquant : ' + need[i][1] + '.', 'MISSING');
  var inc = new Date(String(r.incidentDate).slice(0, 16) + ':00+01:00');     // saisie en heure de Tunis
  if (isNaN(inc.getTime())) throw fail_("Date et heure de l'incident invalides.", 'BAD_DATE');
  if (inc.getTime() > Date.now() + 5 * 60000) throw fail_("L'incident ne peut pas être dans le futur.", 'BAD_DATE');

  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var src = ensureSource_(ss);
    var map = sourceMap_(src);
    var row = new Array(src.getLastColumn()).fill('');
    var vals = {
      'Date de soumission': new Date(), 'Matricule': emp.matricule, 'Département': emp.departement, 'Type': emp.type,
      'Shift': safeCell_(r.shift, 30), 'Area': emp.area, 'Sub Area': emp.subArea, 'Qualification': emp.qualification,
      'Zone de problème': safeCell_(r.zone, 80), 'Catégorie': safeCell_(r.categorie, 120), 'Sous-catégorie': safeCell_(r.sousCategorie, 160),
      'Date et heure incident': inc, 'Description': safeCell_(r.description, 4000)
    };
    Object.keys(vals).forEach(function (k) { if (map[k] > -1) row[map[k]] = vals[k]; });
    var at = src.getLastRow() + 1;
    src.getRange(at, 1, 1, row.length).setValues([row]);
    [map['Date de soumission'], map['Date et heure incident']].forEach(function (cIdx) {
      if (cIdx > -1) src.getRange(at, cIdx + 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    });
  } finally { lock.releaseLock(); }
  // L'ID est attribué immédiatement (sans attendre le déclencheur d'une minute)
  var ctx = ctx_(ss); ingest_(ctx, false);
  var last = ctx.dossiers.getLastRow();
  return { ok: true, id: String(ctx.dossiers.getRange(last, 1).getValue()) };
}
