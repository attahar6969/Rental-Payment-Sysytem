function myFunction() {
  
}
/** WebApp.gs — web front end entry points. Add as a NEW file in the SAME script project. Reuses all existing logic. */

function doGet() {
  return HtmlService.createTemplateFromFile('App').evaluate()
    .setTitle('Rent Payment System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function webInit() {
  ensureReady_();
  return { email: userEmail_(), admin: isAdmin_(), business: getSystemConfig().business.name };
}

/** which = 'log' | 'ledger' | 'history'. Returns display text exactly as shown in the sheet. */
function webTable(which) {
  ensureReady_();
  var name = { log: SHEETS.LOG, ledger: SHEETS.LEDGER, history: SHEETS.HISTORY }[which];
  if (!name) throw new Error('Unknown table.');
  var sh = getSheet_(name), lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return { headers: [], rows: [] };
  var v = sh.getRange(1, 1, lr, lc).getDisplayValues();
  var drop = which === 'log' ? L.JSON : -1;
  function cut(r) { return r.filter(function (c, i) { return i !== drop; }); }
  var rows = v.slice(1).filter(function (r) { return r[0] !== ''; }).map(cut);
  if (which === 'log') rows.reverse();
  return { headers: cut(v[0]), rows: rows };
}

function webGetSettings() {
  if (!isAdmin_()) throw new Error('Only an administrator can open Settings.');
  var sh = getSheet_(SHEETS.SETTINGS), n = sh.getLastRow() - SETTINGS_FIRST_ROW + 1;
  return sh.getRange(SETTINGS_FIRST_ROW, 1, n, 3).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { key: String(r[0]), value: String(r[1]), note: String(r[2]) }; });
}

function webSaveSettings(list) {
  if (!isAdmin_()) throw new Error('Only an administrator can change Settings.');
  return withLock_(function () {
    var sh = getSheet_(SHEETS.SETTINGS), n = sh.getLastRow() - SETTINGS_FIRST_ROW + 1;
    var keys = sh.getRange(SETTINGS_FIRST_ROW, 1, n, 1).getValues();
    var map = {};
    (list || []).forEach(function (s) { map[String(s.key)] = String(s.value); });
    keys.forEach(function (k, i) {
      var key = String(k[0]);
      if (map.hasOwnProperty(key)) sh.getRange(SETTINGS_FIRST_ROW + i, 2).setValue(map[key]);
    });
    SpreadsheetApp.flush();
    getSystemConfig(true);
    clearTenantCache_();
    return true;
  });
}
