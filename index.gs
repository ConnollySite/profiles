/**
 * PROFILES - backend (Google Apps Script)
 * Connolly Site - Bidvest Noonan
 *
 * Shared staff profiles for the supervisors: notes (Check Points), Informal Counseling,
 * verbal / written warnings, accident reports and probation reviews.
 *
 * SETUP
 *  1. Create a NEW Google Sheet (name it "Profiles"). Extensions > Apps Script. Paste this file.
 *  2. Set the PIN of each supervisor in SUPERVISORS below (4 to 8 digits). An empty PIN = that person cannot log in.
 *  3. Run setupSheets() once (authorise when asked).
 *  4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL.
 *  5. To change this code later: Deploy > Manage deployments > pencil > New version (the /exec URL stays).
 *
 * Every request needs the supervisor name + PIN. Nothing is readable without them.
 */

var SHEET_ID = '';          // leave empty when the script was created from inside the sheet (Extensions > Apps Script)
var TZ = 'Europe/Dublin';

var SUPERVISORS = {
  'Leandro Rocha': '',
  'Michal Rogala': ''
};

var EMP_SHEET = 'Employees';
var EMP_HEADERS = ['ID', 'Name', 'Badge', 'Team', 'Status', 'Probation Start', 'Probation End', 'Probation Status', 'Notes', 'Created', 'Updated'];
var ENT_SHEET = 'Entries';
var ENT_HEADERS = ['ID', 'Created', 'Updated', 'Date', 'Time', 'Type', 'Employee IDs', 'Employee Names', 'Teams', 'Category', 'Title', 'Text', 'Data', 'Author', 'Updated By', 'Deleted', 'Source'];
var SIG_SHEET = 'Signatures';
var SIG_HEADERS = ['Entry ID', 'Image'];

var TYPES = ['NOTE', 'COUNSELING', 'VERBAL', 'WRITTEN', 'ACCIDENT', 'PROBATION'];

/* ---------------- helpers ---------------- */
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@');
    sh.appendRow(headers);
  }
  return sh;
}
function empSheet_() { return sheet_(EMP_SHEET, EMP_HEADERS); }
function entSheet_() { return sheet_(ENT_SHEET, ENT_HEADERS); }
function sigSheet_() { return sheet_(SIG_SHEET, SIG_HEADERS); }

function setupSheets() {
  var ss = ss_();
  empSheet_(); entSheet_(); sigSheet_();
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return 'Ready';
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }
function s_(v, max) { return String(v == null ? '' : v).slice(0, max || 200); }
function normName_(n) { return String(n || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function newId_(p) { return p + new Date().getTime().toString(36) + Math.floor(Math.random() * 46656).toString(36); }
function validDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')); }
function validTime_(t) { return t === '' || /^\d{2}:\d{2}$/.test(String(t || '')); }

/* ---------------- login (no lock-out: a wrong PIN can be retried any number of times) ---------------- */
function auth_(user, pin) {
  user = String(user || '');
  if (!SUPERVISORS.hasOwnProperty(user)) return { error: 'Unknown supervisor' };
  var good = SUPERVISORS[user];
  if (!good) return { error: 'No PIN set for this supervisor in the script' };
  if (String(pin || '') !== String(good)) return { error: 'Incorrect PIN' };
  return { ok: true, user: user };
}

/* ---------------- data access ---------------- */
function empRows_() {
  var v = empSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!String(v[i][0]).trim()) continue;
    out.push({
      id: String(v[i][0]), name: String(v[i][1]), badge: String(v[i][2]), team: String(v[i][3]),
      status: String(v[i][4]) || 'active', probStart: String(v[i][5]), probEnd: String(v[i][6]),
      probStatus: String(v[i][7]), notes: String(v[i][8]), created: String(v[i][9]), updated: String(v[i][10])
    });
  }
  return out;
}
function writeEmps_(rows) {
  var sh = empSheet_();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, EMP_HEADERS.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, EMP_HEADERS.length).setValues(rows.map(function (r) {
      return [r.id, r.name, r.badge, r.team, r.status, r.probStart, r.probEnd, r.probStatus, r.notes, r.created, r.updated];
    }));
  }
}
function entRows_(withDeleted) {
  var v = entSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    if (!String(v[i][0]).trim()) continue;
    var deleted = String(v[i][15]);
    if (deleted && !withDeleted) continue;
    var data = {};
    try { data = JSON.parse(String(v[i][12]) || '{}'); } catch (e) { data = {}; }
    out.push({
      id: String(v[i][0]), created: String(v[i][1]), updated: String(v[i][2]), date: String(v[i][3]), time: String(v[i][4]),
      type: String(v[i][5]), employeeIds: String(v[i][6]).split('|').filter(String), employeeNames: String(v[i][7]).split('|').filter(String),
      teams: String(v[i][8]).split('|').filter(String), category: String(v[i][9]), title: String(v[i][10]), text: String(v[i][11]),
      data: data, author: String(v[i][13]), updatedBy: String(v[i][14]), deleted: deleted, source: String(v[i][16])
    });
  }
  return out;
}
function entToRow_(e) {
  return [e.id, e.created, e.updated, e.date, e.time, e.type, e.employeeIds.join('|'), e.employeeNames.join('|'), e.teams.join('|'),
    e.category, e.title, e.text, JSON.stringify(e.data || {}), e.author, e.updatedBy, e.deleted || '', e.source];
}
function writeEnts_(rows) {
  var sh = entSheet_();
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, ENT_HEADERS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, ENT_HEADERS.length).setValues(rows.map(entToRow_));
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- actions ---------------- */
function listAll_(user) {
  return {
    ok: true, user: user, users: Object.keys(SUPERVISORS),
    employees: empRows_(), entries: entRows_(false), serverTime: nowStr_()
  };
}

function saveEmployee_(user, e) {
  e = e || {};
  var name = s_(e.name, 80).trim().replace(/\s+/g, ' ');
  if (!name) return { error: 'Name is required' };
  var status = e.status === 'left' ? 'left' : 'active';
  var ps = String(e.probStatus || '');
  if (['', 'in_progress', 'approved', 'extended', 'failed'].indexOf(ps) < 0) ps = '';
  var pStart = validDate_(e.probStart) ? e.probStart : '';
  var pEnd = validDate_(e.probEnd) ? e.probEnd : '';
  return withLock_(function () {
    var rows = empRows_();
    var row = null;
    rows.forEach(function (r) { if (e.id && r.id === String(e.id)) row = r; });
    if (!row) {
      var dup = null;
      rows.forEach(function (r) { if (normName_(r.name) === normName_(name)) dup = r; });
      if (dup) return { error: 'There is already an employee called ' + dup.name };
      row = { id: newId_('E'), created: nowStr_() };
      rows.push(row);
    }
    row.name = name; row.badge = s_(e.badge, 30); row.team = s_(e.team, 60); row.status = status;
    row.probStart = pStart; row.probEnd = pEnd; row.probStatus = ps; row.notes = s_(e.notes, 1000); row.updated = nowStr_();
    writeEmps_(rows);
    return { ok: true, id: row.id };
  });
}

function importEmployees_(user, list) {
  list = (list || []).slice(0, 500);
  return withLock_(function () {
    var rows = empRows_();
    var known = {};
    rows.forEach(function (r) { known[normName_(r.name)] = true; });
    var added = 0, skipped = 0;
    list.forEach(function (it) {
      var name = s_(it && it.name, 80).trim().replace(/\s+/g, ' ');
      if (!name) return;
      if (known[normName_(name)]) { skipped++; return; }
      known[normName_(name)] = true;
      rows.push({ id: newId_('E'), name: name, badge: s_(it.badge, 30), team: s_(it.team, 60), status: 'active', probStart: '', probEnd: '', probStatus: '', notes: '', created: nowStr_(), updated: nowStr_() });
      added++;
    });
    if (added) writeEmps_(rows);
    return { ok: true, added: added, skipped: skipped };
  });
}

function cleanEntry_(user, e, source) {
  e = e || {};
  var id = s_(e.id, 60).replace(/[^A-Za-z0-9_-]/g, '');
  if (!id) return { error: 'Missing id' };
  if (TYPES.indexOf(e.type) < 0) return { error: 'Bad type' };
  if (!validDate_(e.date)) return { error: 'Bad date' };
  if (!validTime_(e.time || '')) return { error: 'Bad time' };
  var ids = (e.employeeIds || []).map(function (x) { return s_(x, 40); }).filter(String).slice(0, 10);
  var names = (e.employeeNames || []).map(function (x) { return s_(x, 80).replace(/\|/g, ' '); }).filter(String).slice(0, 10);
  var data = e.data && typeof e.data === 'object' ? e.data : {};
  var dj = JSON.stringify(data);
  if (dj.length > 40000) return { error: 'Entry too large' };
  return {
    ok: true,
    entry: {
      id: id, type: e.type, date: e.date, time: e.time || '', employeeIds: ids, employeeNames: names,
      teams: (e.teams || []).map(function (x) { return s_(x, 60).replace(/\|/g, ' '); }).filter(String).slice(0, 10),
      category: s_(e.category, 60), title: s_(e.title, 200), text: String(e.text || '').slice(0, 8000), data: data,
      source: source
    }
  };
}

function saveEntry_(user, e, sig) {
  var c = cleanEntry_(user, e, 'profiles');
  if (c.error) return c;
  var ent = c.entry;
  if (!ent.employeeIds.length) return { error: 'Pick an employee' };
  if (sig && (typeof sig !== 'string' || sig.indexOf('data:image/png;base64,') !== 0 || sig.length > 45000)) return { error: 'Bad signature image' };
  return withLock_(function () {
    var rows = entRows_(true);
    var cur = null;
    rows.forEach(function (r) { if (r.id === ent.id) cur = r; });
    if (cur) {
      ent.created = cur.created; ent.author = cur.author; ent.source = cur.source || 'profiles'; ent.deleted = cur.deleted;
    } else {
      ent.created = nowStr_(); ent.author = user; ent.deleted = '';
    }
    ent.updated = nowStr_(); ent.updatedBy = user;
    var out = rows.filter(function (r) { return r.id !== ent.id; });
    out.push(ent);
    writeEnts_(out);
    if (sig) saveSig_(ent.id, sig);
    applyProbation_(ent);
    return { ok: true, id: ent.id, updated: ent.updated };
  });
}

// A probation review with a decision updates the employee's probation status.
function applyProbation_(ent) {
  if (ent.type !== 'PROBATION' || ent.deleted) return;
  var d = ent.data || {};
  var map = { approve: 'approved', extend: 'extended', fail: 'failed' };
  var st = map[d.decision];
  if (!st) return;
  var rows = empRows_();
  var changed = false;
  rows.forEach(function (r) {
    if (ent.employeeIds.indexOf(r.id) >= 0) {
      r.probStatus = st; r.updated = nowStr_(); changed = true;
      if (st === 'extended' && validDate_(d.newEnd)) r.probEnd = d.newEnd;
    }
  });
  if (changed) writeEmps_(rows);
}

function saveSig_(entryId, sig) {
  var sh = sigSheet_();
  var v = sh.getDataRange().getValues();
  var rowN = -1;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === entryId) { rowN = i + 1; break; }
  if (rowN > 0) sh.getRange(rowN, 1, 1, 2).setValues([[entryId, sig]]);
  else sh.appendRow([entryId, sig]);
}
function getSig_(id) {
  var v = sigSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return { ok: true, image: String(v[i][1]) };
  return { ok: true, image: '' };
}

function deleteEntry_(user, id) {
  return withLock_(function () {
    var rows = entRows_(true);
    var found = false;
    rows.forEach(function (r) { if (r.id === String(id) && !r.deleted) { r.deleted = user + ' ' + nowStr_(); found = true; } });
    if (!found) return { error: 'Entry not found' };
    writeEnts_(rows);
    return { ok: true };
  });
}

/* ---- Check Points app: names list and note upload ---- */
function employeeNames_() {
  return { ok: true, names: empRows_().filter(function (e) { return e.status !== 'left'; }).map(function (e) { return e.name; }) };
}

function syncNotes_(user, notes, deleted) {
  notes = (notes || []).slice(0, 300);
  deleted = (deleted || []).slice(0, 300).map(function (x) { return s_(x, 60); });
  return withLock_(function () {
    var emps = empRows_();
    var byName = {};
    emps.forEach(function (e) { byName[normName_(e.name)] = e; });
    var created = [];
    var rows = entRows_(true);
    var byId = {};
    rows.forEach(function (r) { byId[r.id] = r; });
    var saved = 0, rejected = [];
    var empChanged = false;
    notes.forEach(function (n) {
      var names = (n.employees || []).map(function (x) { return s_(x, 80).trim().replace(/\s+/g, ' '); }).filter(String).slice(0, 10);
      if (!names.length) { rejected.push(String(n && n.id)); return; }
      var ids = [];
      names.forEach(function (nm) {
        var e = byName[normName_(nm)];
        if (!e) {
          e = { id: newId_('E'), name: nm, badge: '', team: '', status: 'active', probStart: '', probEnd: '', probStatus: '', notes: 'Added from Check Points', created: nowStr_(), updated: nowStr_() };
          emps.push(e); byName[normName_(nm)] = e; created.push(nm); empChanged = true;
        }
        ids.push(e.id);
      });
      var c = cleanEntry_(user, {
        id: n.id, type: 'NOTE', date: n.date, time: n.time, employeeIds: ids, employeeNames: names.map(function (nm) { return byName[normName_(nm)].name; }),
        teams: n.teams, category: n.category, title: '', text: n.text,
        data: { kind: n.type === 'positive' ? 'positive' : 'fault' }
      }, 'checkpoints');
      if (c.error) { rejected.push(String(n.id)); return; }
      var ent = c.entry;
      var cur = byId[ent.id];
      if (cur) {
        if (cur.source !== 'checkpoints') { rejected.push(ent.id); return; }   // never overwrite a Profiles form
        ent.created = cur.created; ent.deleted = cur.deleted;
      } else { ent.created = nowStr_(); ent.deleted = ''; }
      ent.author = s_(n.supervisor, 80) || user; ent.updatedBy = user; ent.updated = nowStr_();
      byId[ent.id] = ent;
      saved++;
    });
    deleted.forEach(function (id) {
      var r = byId[id];
      if (r && r.source === 'checkpoints' && !r.deleted) r.deleted = user + ' ' + nowStr_();
    });
    if (empChanged) writeEmps_(emps);
    var all = Object.keys(byId).map(function (k) { return byId[k]; });
    var order = {};
    rows.forEach(function (r, i) { order[r.id] = i; });
    all.sort(function (a, b) { return (order[a.id] == null ? 1e9 : order[a.id]) - (order[b.id] == null ? 1e9 : order[b.id]); });
    if (saved || deleted.length) writeEnts_(all);
    return { ok: true, saved: saved, rejected: rejected, createdEmployees: created, names: emps.filter(function (e) { return e.status !== 'left'; }).map(function (e) { return e.name; }) };
  });
}

/* ---------------- web entry points ---------------- */
function doGet() {
  return jsonOut_({ ok: true, app: 'Profiles' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return jsonOut_({ error: 'Bad request' }); }
  var a = auth_(req.user, req.pin);
  if (a.error) return jsonOut_(a);
  var user = a.user;
  var act = req.action;
  try {
    if (act === 'login') return jsonOut_({ ok: true, user: user, users: Object.keys(SUPERVISORS) });
    if (act === 'list') return jsonOut_(listAll_(user));
    if (act === 'saveEmployee') return jsonOut_(saveEmployee_(user, req.employee));
    if (act === 'importEmployees') return jsonOut_(importEmployees_(user, req.list));
    if (act === 'saveEntry') return jsonOut_(saveEntry_(user, req.entry, req.sig));
    if (act === 'getSig') return jsonOut_(getSig_(req.id));
    if (act === 'deleteEntry') return jsonOut_(deleteEntry_(user, req.id));
    if (act === 'names') return jsonOut_(employeeNames_());
    if (act === 'syncNotes') return jsonOut_(syncNotes_(user, req.notes, req.deleted));
    return jsonOut_({ error: 'Unknown action' });
  } catch (err) {
    return jsonOut_({ error: 'Server error: ' + (err && err.message ? err.message : err) });
  }
}
