/**
 * Config.gs — all configurable values live here or in the SETTINGS sheet.
 * Nothing in this file touches money logic.
 */

var SHEETS = {
  LOG: 'RENT LOG',
  LEDGER: 'RENT LEDGER',
  HISTORY: 'RENT HISTORY',
  DASH: 'DASHBOARD',
  SETTINGS: 'SETTINGS',
  ERRORS: 'ERROR LOG'
};

var CURRENCY_FORMAT = '[>=10000000]"₹"##\\,##\\,##\\,##0.00;[>=100000]"₹"##\\,##\\,##0.00;"₹"##,##0.00';
var DATE_FORMAT = 'dd-MMM-yyyy';
var DATETIME_FORMAT = 'dd-MMM-yyyy HH:mm:ss';
var PRORATION_MODES = ['DAILY', 'FLAT30', 'FULL_MONTH_IF_ANY', 'NONE'];

/* ---------- RENT LOG (append-only) ---------- */
var LOG_COLS = ['Receipt ID', 'Payment Date', 'Tenant ID', 'Tenant Name', 'Phone Number',
  'Agreement Rent Snapshot', 'Agreement Start Date Snapshot', 'Agreement End Date Snapshot',
  'Amount Received', 'Rent From', 'Rent Till', 'Outstanding Before Payment',
  'Amount Allocated to Previous Dues', 'Amount Allocated to Current Month', 'Advance Amount',
  'Outstanding After Payment', 'Payment Status', 'Entry Timestamp', 'Entered By',
  'Transaction ID', 'Allocation Reference', 'Allocation Detail (JSON)', 'Property / Room', 'Waived Off Amount', 'Waiver Reason', 'Payment Mode'];
var L = { RECEIPT: 0, DATE: 1, TID: 2, NAME: 3, PHONE: 4, RENT: 5, AS: 6, AE: 7, AMT: 8, RFROM: 9,
  RTILL: 10, OUTB: 11, ALPREV: 12, ALCUR: 13, ADV: 14, OUTA: 15, STATUS: 16, TS: 17, BY: 18,
  TXN: 19, ALREF: 20, JSON: 21, PROP: 22, WAIVE: 23, WREASON: 24, PMODE: 25 };

/* ---------- RENT HISTORY ---------- */
var HIST_COLS = ['Tenant ID', 'Tenant Name', 'Previous Rent', 'Revised Rent', 'Effective From',
  'Effective Till', 'Revision Date', 'Revision Reference', 'Updated By', 'Revision Status',
  'Agreement Start', 'Agreement End', 'Change Type'];
var H = { TID: 0, NAME: 1, PREV: 2, NEW: 3, FROM: 4, TILL: 5, REVDATE: 6, REF: 7, BY: 8, STATUS: 9,
  AS: 10, AE: 11, TYPE: 12 };
// Revision status values: BASELINE, PENDING, CONFIRMED, APPLIED, REJECTED, SUPERSEDED
// Change types: BASELINE, RENT_REVISION, NEW_AGREEMENT, END_DATE_CHANGE

/* ---------- RENT LEDGER (derived) ---------- */
var LEDGER_COLS = ['Tenant ID', 'Tenant Name', 'Rent Month', 'Applicable Agreement Rent',
  'Opening Outstanding', 'Monthly Rent Due', 'Total Payment Allocated', 'Closing Outstanding',
  'Advance Credit', 'Payment Status', 'Last Payment Date', 'Last Receipt ID', 'Updated Timestamp',
  'Cumulative Outstanding'];

/* ---------- SETTINGS sheet ---------- */
var SETTINGS_FIRST_ROW = 4;      // key/value rows start here (A=key, B=value, C=note)
var OPENING_COL = 6;             // opening-balance table starts in column F
var SETTING_DEFS = [
  ['PARENT_SPREADSHEET_ID', '115wFO57oP85vMKVKGxx4j8piEbcb2P1zACXGgaZ8GXk', 'ID of the RENTAL AGREEMENT REGISTER workbook (text between /d/ and /edit in its URL).'],
  ['PARENT_SHEET_NAME', 'Property Status', 'Tab in the parent workbook that holds current agreements.'],
  ['BUSINESS_NAME', 'YOUR BUSINESS NAME', 'Printed on every receipt (saved with each receipt at the time of payment).'],
  ['BUSINESS_ADDRESS', 'Raipur, Chhattisgarh', 'Printed on receipts.'],
  ['BUSINESS_PHONE', '', 'Printed on receipts.'],
  ['RECEIPT_PREFIX', 'RENT', 'Receipt IDs look like RENT-2026-001.'],
  ['RECEIPT_YEAR_MODE', 'CALENDAR', 'CALENDAR = RENT-2026-001 (restarts 1 Jan). FINANCIAL = RENT-2026-27-001 (restarts 1 Apr).'],
  ['PRORATION_MODE', 'DAILY', 'Partial-month rent: DAILY (rent x days / days in month), FLAT30 (rent x days / 30), FULL_MONTH_IF_ANY, NONE (no rent for partial months).'],
  ['LEDGER_START_MONTH', '2026-09', 'First month tracked by this system (yyyy-MM). Earlier dues go in the Opening Balances table.'],
  ['ADMIN_EMAILS', '', 'Comma-separated Google accounts allowed to confirm revisions and rebuild. Blank = workbook owner only.'],
  ['MAX_ADVANCE_MONTHS', 12, 'How many months ahead a payment may be applied.'],
  ['COL_TENANT_ID', 'A', 'Parent column: Tenant ID (property Code).'],
  ['COL_PROPERTY', 'B', 'Parent column: Property address / room.'],
  ['COL_LANDLORD', 'D', 'Parent column: Landlord (optional).'],
  ['COL_TENANT_NAME', 'E', 'Parent column: Tenant name.'],
  ['COL_PHONE', 'F', 'Parent column: Mobile.'],
  ['COL_START', 'G', 'Parent column: Agreement start date.'],
  ['COL_END', 'H', 'Parent column: Agreement end date.'],
  ['COL_RENT', 'J', 'Parent column: Monthly rent.'],
  ['COL_STATUS', 'L', 'Parent column: Status (contains the word Active for live tenants).']
];

/* Parent column definitions. `kw` = header keywords used to VERIFY the configured column. */
var PARENT_FIELDS = {
  id:       { setting: 'COL_TENANT_ID',   kw: ['code', 'tenant id'] },
  property: { setting: 'COL_PROPERTY',    kw: ['property address', 'address', 'room'] },
  landlord: { setting: 'COL_LANDLORD',    kw: ['makan malik', 'landlord'], optional: true },
  name:     { setting: 'COL_TENANT_NAME', kw: ['kirayedar', 'tenant'] },
  phone:    { setting: 'COL_PHONE',       kw: ['mobile', 'phone'] },
  start:    { setting: 'COL_START',       kw: ['shuru', 'start'] },
  end:      { setting: 'COL_END',         kw: ['khatam', 'samapt', 'end'] },
  rent:     { setting: 'COL_RENT',        kw: ['masik', 'kiraya', 'rent'] },
  status:   { setting: 'COL_STATUS',      kw: ['sthiti', 'status'] }
};

var _cfgCache = null;

/** Reads SETTINGS once per execution. */
function getSystemConfig(force) {
  if (_cfgCache && !force) return _cfgCache;
  var kv = {};
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  if (sh && sh.getLastRow() >= SETTINGS_FIRST_ROW) {
    sh.getRange(SETTINGS_FIRST_ROW, 1, sh.getLastRow() - SETTINGS_FIRST_ROW + 1, 2).getValues()
      .forEach(function (r) { if (r[0] !== '') kv[String(r[0]).trim()] = r[1]; });
  }
  var defs = {};
  SETTING_DEFS.forEach(function (d) { defs[d[0]] = d[1]; });
  function g(k) { return (kv[k] === undefined || kv[k] === '') ? defs[k] : kv[k]; }

  var mode = String(g('PRORATION_MODE')).trim().toUpperCase();
  if (PRORATION_MODES.indexOf(mode) < 0) mode = 'DAILY';

  var startVal = g('LEDGER_START_MONTH');
  var startDate = (startVal instanceof Date) ? toDate_(startVal) : toDate_(String(startVal).trim() + '-01');
  if (!startDate) startDate = toDate_('2026-09-01');

  var cols = {};
  Object.keys(PARENT_FIELDS).forEach(function (f) {
    cols[f] = String(g(PARENT_FIELDS[f].setting) || '').trim().toUpperCase();
  });

  _cfgCache = {
    parentId: String(g('PARENT_SPREADSHEET_ID')).trim(),
    parentSheet: String(g('PARENT_SHEET_NAME')).trim(),
    business: { name: String(g('BUSINESS_NAME')), address: String(g('BUSINESS_ADDRESS') || ''), phone: String(g('BUSINESS_PHONE') || '') },
    prefix: String(g('RECEIPT_PREFIX')).trim() || 'RENT',
    yearMode: String(g('RECEIPT_YEAR_MODE')).trim().toUpperCase() === 'FINANCIAL' ? 'FINANCIAL' : 'CALENDAR',
    proration: mode,
    ledgerStartIdx: monthIdx_(startDate),
    admins: String(g('ADMIN_EMAILS') || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String),
    maxAdvanceMonths: Number(g('MAX_ADVANCE_MONTHS')) || 12,
    cols: cols
  };
  return _cfgCache;
}

/** Opening balances table (SETTINGS F:I). Returns { tenantId: {outP, advP, mode} } in paise. */
function getOpeningBalances_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  var map = {};
  if (!sh || sh.getLastRow() < SETTINGS_FIRST_ROW) return map;
  sh.getRange(SETTINGS_FIRST_ROW, OPENING_COL, sh.getLastRow() - SETTINGS_FIRST_ROW + 1, 4).getValues()
    .forEach(function (r) {
      var id = String(r[0]).trim();
      if (!id) return;
      var mode = String(r[3] || '').trim().toUpperCase();
      map[id] = { outP: Math.max(0, toPaise_(parseMoney_(r[1]) || 0)), advP: Math.max(0, toPaise_(parseMoney_(r[2]) || 0)),
        mode: PRORATION_MODES.indexOf(mode) >= 0 ? mode : '' };
    });
  return map;
}

/* ---------- admin control ---------- */
function isAdmin_() {
  var me = userEmail_().toLowerCase();
  var cfg = getSystemConfig();
  if (cfg.admins.length) return cfg.admins.indexOf(me) >= 0;
  try {
    var owner = SpreadsheetApp.getActive().getOwner();
    if (owner) return owner.getEmail().toLowerCase() === me;
  } catch (e) { /* shared drive: no owner */ }
  return false;
}

function requireAdmin_() {
  if (!isAdmin_()) throw new Error('Only an administrator can do this. Add your Google account to ADMIN_EMAILS in SETTINGS.');
}
