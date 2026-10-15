/** Utils.gs — shared helpers. Pure functions here are unit-tested. */

var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function today_() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

function mkDate_(y, m, d) {
  var dt = new Date(y, m, d);
  return (dt.getFullYear() === y && dt.getMonth() === m && dt.getDate() === d) ? dt : null;
}

/** Any date-like value -> local midnight Date (or null). */
function toDate_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return mkDate_(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[A-Za-z]*[-\/ ,]*(\d{4})/))) {
    var mi = MONTHS.indexOf(m[2].charAt(0).toUpperCase() + m[2].slice(1).toLowerCase());
    return mi >= 0 ? mkDate_(+m[3], mi, +m[1]) : null;
  }
  if ((m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))) return mkDate_(+m[3], +m[2] - 1, +m[1]);
  return null;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function isoDate_(d) { return d ? d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate()) : ''; }
function fmtDate_(d) { return d ? pad2_(d.getDate()) + '-' + MONTHS[d.getMonth()] + '-' + d.getFullYear() : ''; }

function dayNum_(d) { return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000); }
function monthIdx_(d) { return d.getFullYear() * 12 + d.getMonth(); }
function idxStart_(idx) { return new Date(Math.floor(idx / 12), idx % 12, 1); }
function idxEnd_(idx) { return new Date(Math.floor(idx / 12), idx % 12 + 1, 0); }
function idxLabel_(idx) { return MONTHS[idx % 12] + ' ' + Math.floor(idx / 12); }
function idxIso_(idx) { return Math.floor(idx / 12) + '-' + pad2_(idx % 12 + 1); }

function toPaise_(n) { return Math.round(Number(n) * 100); }
function rupees_(p) { return Math.round(p) / 100; }

function parseMoney_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? null : n;
}

/** Indian digit grouping: ₹1,23,456.50 */
function formatINR_(n, dec) {
  n = Number(n) || 0;
  var neg = n < 0; n = Math.abs(n);
  if (dec === undefined) dec = (Math.abs(n - Math.round(n)) < 0.005) ? 0 : 2;
  var parts = n.toFixed(dec).split('.');
  var i = parts[0], last3 = i.slice(-3), rest = i.slice(0, -3);
  if (rest) last3 = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3;
  return (neg ? '-' : '') + '₹' + last3 + (parts[1] ? '.' + parts[1] : '');
}

var ONES_ = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve',
  'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
var TENS_ = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

function wordsBelow1000_(n) {
  var s = '';
  if (n >= 100) { s += ONES_[Math.floor(n / 100)] + ' Hundred'; n %= 100; if (n) s += ' '; }
  if (n >= 20) { s += TENS_[Math.floor(n / 10)]; if (n % 10) s += ' ' + ONES_[n % 10]; }
  else if (n > 0) s += ONES_[n];
  return s;
}

function numberToWordsIndian_(n) {
  if (n === 0) return 'Zero';
  var parts = [], units = [[10000000, 'Crore'], [100000, 'Lakh'], [1000, 'Thousand']];
  units.forEach(function (u) {
    if (n >= u[0]) { parts.push(wordsBelow1000_(Math.floor(n / u[0])) + ' ' + u[1]); n %= u[0]; }
  });
  if (n > 0) parts.push(wordsBelow1000_(n));
  return parts.join(' ');
}

/** 20000.5 -> "Rupees Twenty Thousand and Fifty Paise Only" */
function amountInWords_(amount) {
  var p = Math.round(Number(amount) * 100), r = Math.floor(p / 100), ps = p % 100;
  return 'Rupees ' + numberToWordsIndian_(r) + (ps ? ' and ' + numberToWordsIndian_(ps) + ' Paise' : '') + ' Only';
}

function userEmail_() {
  try { var e = Session.getActiveUser().getEmail(); if (e) return e; } catch (x) { }
  try { return Session.getEffectiveUser().getEmail() || 'unknown'; } catch (y) { return 'unknown'; }
}

function normText_(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim(); }

function letterToIdx_(letters) {
  var n = 0, s = String(letters).toUpperCase();
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

function getSheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run initializeSystem() first.');
  return sh;
}

function logError_(where, err) {
  try {
    console.error(where + ': ' + (err && err.stack ? err.stack : err));
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ERRORS);
    if (sh) sh.appendRow([new Date(), userEmail_(), where, String(err && err.message ? err.message : err)]);
  } catch (e) { /* never let logging break a transaction */ }
}

/** Runs fn under the document lock. */
function withLock_(fn, timeoutMs) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(timeoutMs || 30000)) throw new Error('The system is busy with another entry. Please try again in a few seconds.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function groupBy_(arr, keyFn) {
  var out = {};
  arr.forEach(function (x) { var k = keyFn(x); (out[k] = out[k] || []).push(x); });
  return out;
}

/** Natural sort compare so P2 comes before P10. */
function natCompare_(a, b) {
  return String(a).localeCompare(String(b), 'en', { numeric: true, sensitivity: 'base' });
}
