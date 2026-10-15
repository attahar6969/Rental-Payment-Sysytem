/** Main.gs — menu, setup and orchestration. */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏠 RENTAL PAYMENT')
    .addItem('New Payment Entry', 'showPaymentEntry')
    .addSeparator()
    .addItem('View Rent Log', 'viewRentLog')
    .addItem('View Rent Ledger', 'viewRentLedger')
    .addItem('View Rent History', 'viewRentHistory')
    .addItem('Search / Reprint / Delete Receipt', 'showReceiptSearch')
    .addSeparator()
    .addItem('Refresh System', 'menuRefreshSystem')
    .addItem('Review Rent Revisions', 'showRevisionReview')
    .addItem('Rebuild Ledger', 'menuRebuildLedger')
    .addItem('Open Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('System Settings', 'openSettings')
    .addItem('First-time Setup (Initialize)', 'menuInitialize')
    .addToUi();
}

function alert_(msg) { SpreadsheetApp.getUi().alert(msg); }

function ensureReady_() {
  var ss = SpreadsheetApp.getActive();
  [SHEETS.LOG, SHEETS.LEDGER, SHEETS.HISTORY, SHEETS.DASH, SHEETS.SETTINGS].forEach(function (n) {
    if (!ss.getSheetByName(n)) throw new Error('Sheet "' + n + '" is missing. Use menu > First-time Setup (Initialize).');
  });
}

function showPaymentEntry() {
  try {
    ensureReady_();
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Entry').setWidth(1020).setHeight(780), 'New Rent Payment');
  } catch (e) { logError_('showPaymentEntry', e); alert_(e.message); }
}

function showReceiptSearch() {
  try { ensureReady_(); SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Search').setWidth(820).setHeight(640), 'Search / Reprint Receipt'); }
  catch (e) { alert_(e.message); }
}

function showDashboard() {
  try { ensureReady_(); SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Dashboard').setWidth(1280).setHeight(800), 'Rent Dashboard'); }
  catch (e) { alert_(e.message); }
}

function showRevisionReview() {
  try { ensureReady_(); SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Revision').setWidth(780).setHeight(620), 'Rent Revisions Awaiting Review'); }
  catch (e) { alert_(e.message); }
}

function activateSheet_(name) { var sh = getSheet_(name); SpreadsheetApp.getActive().setActiveSheet(sh); }
function viewRentLog() { try { activateSheet_(SHEETS.LOG); } catch (e) { alert_(e.message); } }
function viewRentLedger() { try { activateSheet_(SHEETS.LEDGER); } catch (e) { alert_(e.message); } }
function viewRentHistory() { try { activateSheet_(SHEETS.HISTORY); } catch (e) { alert_(e.message); } }
function openSettings() { try { activateSheet_(SHEETS.SETTINGS); } catch (e) { alert_(e.message); } }
function openDashboardSheet() { activateSheet_(SHEETS.DASH); }

/** Refresh: fetch parent, flag changes, rebuild ledger + dashboard. Never adds payments. */
function refreshSystem() {
  ensureReady_();
  clearTenantCache_();
  var sync = syncAgreementHistory_();
  var rows = rebuildLedgerInternal_();
  refreshDashboard();
  var pending = readHistory_().filter(function (h) { return h.status === 'PENDING'; }).length;
  return { baselines: sync.baselines, flagged: sync.flagged, pending: pending, ledgerRows: rows };
}

function menuRefreshSystem() {
  try {
    var r = refreshSystem(), msg = 'Refresh complete.\nLedger rows: ' + r.ledgerRows + '\nNew tenants recorded: ' + r.baselines;
    if (r.pending) {
      msg += '\n\n' + r.pending + ' agreement change(s) need administrator review (rent, dates). Until confirmed, the old terms stay in force.';
      alert_(msg);
      showRevisionReview();
    } else alert_(msg);
  } catch (e) { logError_('menuRefreshSystem', e); alert_('Refresh failed: ' + e.message); }
}

function menuRebuildLedger() {
  try {
    requireAdmin_();
    var ui = SpreadsheetApp.getUi();
    if (ui.alert('Rebuild Ledger', 'Rebuilds RENT LEDGER from RENT LOG and RENT HISTORY. No payments or receipts are changed. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    var n = rebuildRentLedger();
    refreshDashboard();
    alert_('Ledger rebuilt: ' + n + ' rows.');
  } catch (e) { logError_('menuRebuildLedger', e); alert_(e.message); }
}

function menuInitialize() {
  try { alert_(initializeSystem()); } catch (e) { logError_('initializeSystem', e); alert_('Setup failed: ' + e.message); }
}

/** RUN THIS ONCE from the Apps Script editor (or menu). Safe to run again: it never clears data. */
function initializeSystem() {
  var ss = SpreadsheetApp.getActive();
  function ensure(name, headers) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (headers && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#1f2a44').setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
      sh.setFrozenRows(1);
    }
    return sh;
  }
  var log = ensure(SHEETS.LOG, LOG_COLS), led = ensure(SHEETS.LEDGER, LEDGER_COLS), his = ensure(SHEETS.HISTORY, HIST_COLS);
  var dash = ensure(SHEETS.DASH, null), set = ensure(SHEETS.SETTINGS, null), err = ensure(SHEETS.ERRORS, ['Time', 'User', 'Where', 'Message']);

  // SETTINGS
  if (set.getLastRow() < SETTINGS_FIRST_ROW) {
    set.getRange('A1').setValue('SYSTEM SETTINGS').setFontSize(14).setFontWeight('bold');
    set.getRange('A2').setValue('Edit the Value column (B). Opening balances for tenants with dues before the ledger start go in the table at F3.').setFontColor('#777');
    set.getRange(3, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']]).setBackground('#1f2a44').setFontColor('#fff').setFontWeight('bold');
    set.getRange(SETTINGS_FIRST_ROW, 2, SETTING_DEFS.length, 1).setNumberFormat('@');
    set.getRange(SETTINGS_FIRST_ROW, 1, SETTING_DEFS.length, 3).setValues(SETTING_DEFS);
    set.getRange(3, OPENING_COL, 1, 4).setValues([['Tenant ID', 'Opening Outstanding (Rs)', 'Opening Advance (Rs)', 'Proration Override (optional)']])
      .setBackground('#1f2a44').setFontColor('#fff').setFontWeight('bold').setWrap(true);
    set.getRange(SETTINGS_FIRST_ROW, OPENING_COL, 60, 1).setNumberFormat('@');
    set.setColumnWidth(1, 220); set.setColumnWidth(2, 260); set.setColumnWidth(3, 520);
    set.setColumnWidths(OPENING_COL, 4, 150);
  }

  // formats
  log.setColumnWidths(1, LOG_COLS.length, 130); log.setColumnWidth(22, 200);
  led.setColumnWidths(1, LEDGER_COLS.length, 130);
  his.setColumnWidths(1, HIST_COLS.length, 130);
  his.getRange(2, 3, 500, 2).setNumberFormat(CURRENCY_FORMAT);
  his.getRange(2, 5, 500, 3).setNumberFormat(DATE_FORMAT);
  his.getRange(2, 11, 500, 2).setNumberFormat(DATE_FORMAT);
  led.setFrozenColumns(2);

  // status colours on ledger + log
  led.setConditionalFormatRules(statusRules_(led.getRange('J2:J5000')));
  log.setConditionalFormatRules(statusRules_(log.getRange('Q2:Q20000')));
  his.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDING').setBackground('#fbe6cf').setRanges([his.getRange('J2:J5000')]).build()]);

  // warn before hand-editing derived/permanent sheets (warning-only so the script can still write as any operator)
  [log, led, his, dash].forEach(function (s) {
    s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    s.protect().setDescription(s.getName() + ' — managed by the system').setWarningOnly(true);
  });

  getSystemConfig(true);
  var r = refreshSystem();
  return 'Setup complete.\n\nTenants recorded: ' + r.baselines + '\nPending reviews: ' + r.pending +
    '\n\nNext: open SETTINGS and set BUSINESS_NAME, then use menu > New Payment Entry.';
}

function statusRules_(range) {
  function rule(text, color) { return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build(); }
  return [rule('PAID', '#d8f0e2'), rule('PARTIAL', '#fbe6cf'), rule('UNPAID', '#f6d5d1'), rule('ADVANCE', '#d6e6f7')];
}
