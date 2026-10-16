/**
 * SEBN TN2 — Reclamations backend (Google Apps Script)
 *
 * TWO SEPARATE FILES:
 *  1) MASTER file (this script is attached to it) -> tab "Employees". Read-only for this script.
 *  2) RECLAMATIONS file (a different spreadsheet)  -> tab "Reclamations". All submissions go here.
 *
 * The Employees data never leaves this script: the browser only receives ok / not ok.
 */

var EMP_SHEET = 'Employees';
var REC_SHEET = 'Reclamations';

// Paste the ID of the separate "Reclamations" spreadsheet here.
// (The ID is the long text in its URL:  docs.google.com/spreadsheets/d/<THIS_PART>/edit)
var REC_SPREADSHEET_ID = 'PASTE_RECLAMATIONS_FILE_ID_HERE';
var MAX_FAILS = 5;          // wrong attempts allowed per matricule
var BLOCK_SECONDS = 900;    // lock duration (15 min)

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var check = verify_(req.matricule, req.cin);
    if (!check.ok) return out_(check);

    if (req.action === 'verify') {
      return out_({ ok: true, nom: check.nom, departement: check.departement,
                   type: check.type, area: check.area, subArea: check.subArea,
                   qualification: check.qualification });
    }
    if (req.action === 'submit') {
      if (!REC_SPREADSHEET_ID || REC_SPREADSHEET_ID.indexOf('PASTE') === 0 ||
          REC_SPREADSHEET_ID === SpreadsheetApp.getActive().getId()) {
        return out_({ ok: false, error: 'Fichier des réclamations non configuré.' });
      }
      save_(check, req.record || {});
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Action inconnue.' });
  } catch (err) {
    return out_({ ok: false, error: 'Erreur serveur.' });
  }
}

function doGet() {
  return out_({ ok: true, info: 'SEBN TN2 API active' });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// lower-case, no accents, no spaces/quotes -> easy header matching
function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]/g, '');
}
// compare ids ignoring spaces and leading zeros (05344 == 5344)
function id_(v) {
  return String(v == null ? '' : v).trim().replace(/^0+/, '');
}

function verify_(matricule, cin) {
  var m = id_(matricule), c = id_(cin);
  if (!m || !c) return { ok: false, error: 'Matricule et CIN obligatoires.' };

  var cache = CacheService.getScriptCache();
  var key = 'fail_' + m;
  var fails = parseInt(cache.get(key) || '0', 10);
  if (fails >= MAX_FAILS) {
    return { ok: false, error: 'Trop de tentatives. Réessayez dans 15 minutes.' };
  }

  var sh = SpreadsheetApp.getActive().getSheetByName(EMP_SHEET);
  var data = sh.getDataRange().getDisplayValues();

  // find the header row (the one containing "Code matricule")
  var h = -1;
  for (var i = 0; i < Math.min(data.length, 15); i++) {
    if (data[i].some(function (x) { return norm_(x) === 'codematricule'; })) { h = i; break; }
  }
  if (h < 0) return { ok: false, error: 'Configuration invalide (en-têtes introuvables).' };

  var head = data[h].map(norm_);
  var cM = head.indexOf('codematricule');
  var cC = head.indexOf('cin');
  var cN = head.findIndex(function (x) { return x.indexOf('nometprenom') === 0; });
  var cD = head.findIndex(function (x) { return x.indexOf('departement') === 0; });
  var cE = head.indexOf('emploi');
  var cA = head.indexOf('area');
  var cS = head.indexOf('subarea');
  var cQ = head.indexOf('qualification');
  if (cM < 0 || cC < 0) return { ok: false, error: 'Configuration invalide (colonnes).' };

  for (var r = h + 1; r < data.length; r++) {
    if (id_(data[r][cM]) === m && id_(data[r][cC]) === c) {
      cache.remove(key);
      return {
        ok: true,
        matricule: String(data[r][cM]).trim(),
        nom: cN >= 0 ? data[r][cN] : '',
        departement: cD >= 0 ? data[r][cD] : '',
        type: cE >= 0 ? String(data[r][cE]).trim() : '',
        area: cA >= 0 ? String(data[r][cA]).trim() : '',
        subArea: cS >= 0 ? String(data[r][cS]).trim() : '',
        qualification: cQ >= 0 ? String(data[r][cQ]).trim() : ''
      };
    }
  }
  cache.put(key, String(fails + 1), BLOCK_SECONDS);
  return { ok: false, error: 'Matricule ou CIN incorrect.' };
}

function save_(who, rec) {
  var ss = SpreadsheetApp.openById(REC_SPREADSHEET_ID); // separate file, never the master
  var sh = ss.getSheetByName(REC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REC_SHEET);
    sh.appendRow(['Date', 'Matricule', 'Département', 'Type', 'Shift', 'Area',
                  'Sub Area', 'Qualification', 'Zone', 'Catégorie', 'Sous-catégorie',
                  'Date et heure incident', 'Description']);
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    new Date(), "'" + who.matricule, rec.departement || '', who.type,
    rec.shift || '', who.area, who.subArea, who.qualification,
    rec.zone || '', rec.categorie || '', rec.sousCategorie || '',
    rec.incidentDate ? String(rec.incidentDate).replace('T', ' ') : '',
    rec.description || ''
  ]);
}
