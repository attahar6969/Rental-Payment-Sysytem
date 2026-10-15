/** LedgerService.gs — monthly liability + the derived RENT LEDGER sheet. */

/**
 * Rent due for calendar month `idx` given the tenant's billing intervals.
 * Full month covered by one interval -> full rent. Partial coverage follows the proration mode.
 * Returns paise: { dueP, rentP } (rentP = the contractual monthly rent that applies).
 */
function calculateMonthlyLiability(idx, segments, mode) {
  var mStart = dayNum_(idxStart_(idx)), mEnd = dayNum_(idxEnd_(idx)), dim = mEnd - mStart + 1;
  var covers = [], covered = 0;
  segments.forEach(function (s) {
    var a = Math.max(s.startN, mStart), b = Math.min(s.endN, mEnd);
    if (b >= a) { covers.push({ s: s, days: b - a + 1 }); covered += b - a + 1; }
  });
  if (!covers.length) return { dueP: 0, rentP: 0 };
  var main = covers.reduce(function (best, c) { return c.days >= best.days ? c : best; });
  var rentP = main.s.rentP, due = 0;
  if (covered === dim && covers.length === 1) {
    due = rentP;
  } else if (covered === dim) {
    covers.forEach(function (c) { due += c.s.rentP * c.days / dim; });          // mid-month revision: split by days
  } else if (mode === 'DAILY') {
    covers.forEach(function (c) { due += c.s.rentP * c.days / dim; });
  } else if (mode === 'FLAT30') {
    covers.forEach(function (c) { due += c.s.rentP * c.days / 30; });
    due = Math.min(due, covers[covers.length - 1].s.rentP);
  } else if (mode === 'FULL_MONTH_IF_ANY') {
    due = covers[covers.length - 1].s.rentP;
  } else {
    due = 0;                                                                     // NONE
  }
  return { dueP: Math.round(due), rentP: rentP };
}

function calculateClosingOutstanding(openP, dueP, allocP) { return openP + dueP - allocP; }

/** Loads everything the engine needs in a handful of sheet reads. */
function buildContext_() {
  var cfg = getSystemConfig(), hist = readHistory_(), log = readLog_(), byT = {}, segCache = {};
  log.forEach(function (p) { (byT[p.tid] = byT[p.tid] || []).push({ id: p.receipt, date: p.date, rentFrom: p.rentFrom, rentTill: p.rentTill, amountP: p.amountP, waiveP: p.waiveP }); });
  return { cfg: cfg, hist: hist, opening: getOpeningBalances_(), logByTenant: byT,
    segsFor: function (tid) { return segCache[tid] || (segCache[tid] = getSegments_(tid, hist)); } };
}

/** All tenants' engine states + ledger rows. */
function computeAllLedgers_() {
  var ctx = buildContext_(), names = {}, order = [], asOf = monthIdx_(today_());
  ctx.hist.forEach(function (h) { if (!names[h.tid]) order.push(h.tid); names[h.tid] = h.name; });
  try {
    order = [];
    getParentTenants_(false).forEach(function (t) { names[t.id] = t.name; order.push(t.id); });
    ctx.hist.forEach(function (h) { if (order.indexOf(h.tid) < 0) order.push(h.tid); });
  } catch (e) { logError_('computeAllLedgers_.parent', e); }
  var out = [];
  order.forEach(function (tid) {
    if (!ctx.segsFor(tid).length && !(ctx.opening[tid])) return;
    var run = runTenantEngine_(ctx, tid, null, true);
    out.push({ tid: tid, name: names[tid] || tid, st: run.st, rows: buildLedgerRows_(run.st, asOf) });
  });
  return { ctx: ctx, tenants: out };
}

function ledgerRowToArray_(tid, name, r, ts) {
  return [tid, name, idxStart_(r.idx), rupees_(r.rentP), rupees_(r.openP), rupees_(r.dueP), rupees_(r.allocP),
    rupees_(r.closeP), rupees_(r.advP), r.status, r.lastDate || '', r.lastReceipt, ts, rupees_(r.cumP)];
}

function writeLedgerSheet_(values) {
  var sh = getSheet_(SHEETS.LEDGER);
  var last = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, last - 1, LEDGER_COLS.length).clearContent();
  if (values.length) {
    sh.getRange(2, 1, values.length, LEDGER_COLS.length).setValues(values);
    sh.getRange(2, 3, values.length, 1).setNumberFormat('mmm yyyy');
    sh.getRange(2, 4, values.length, 5).setNumberFormat(CURRENCY_FORMAT);
    sh.getRange(2, 9, values.length, 1).setNumberFormat(CURRENCY_FORMAT);
    sh.getRange(2, 11, values.length, 1).setNumberFormat(DATE_FORMAT);
    sh.getRange(2, 13, values.length, 1).setNumberFormat(DATETIME_FORMAT);
    sh.getRange(2, 14, values.length, 1).setNumberFormat(CURRENCY_FORMAT);
  }
}

/** Deterministic full rebuild from RENT LOG + RENT HISTORY. Safe to run repeatedly. No admin check (used by Refresh). */
function rebuildLedgerInternal_() {
  return withLock_(function () {
    var all = computeAllLedgers_(), ts = new Date(), values = [];
    all.tenants.forEach(function (t) { t.rows.forEach(function (r) { values.push(ledgerRowToArray_(t.tid, t.name, r, ts)); }); });
    writeLedgerSheet_(values);
    SpreadsheetApp.flush();
    return values.length;
  });
}

/** Administrator-only rebuild (menu). */
function rebuildRentLedger() {
  requireAdmin_();
  return rebuildLedgerInternal_();
}

/** Recomputes one tenant's rows and replaces only those rows. */
function updateTenantLedger(tenantId) {
  var tid = String(tenantId).trim();
  var ctx = buildContext_(), ts = new Date();
  var run = runTenantEngine_(ctx, tid, null, true);
  var name = tid;
  try { var t = getTenantDetails(tid); if (t) name = t.name; } catch (e) { }
  var mine = buildLedgerRows_(run.st, monthIdx_(today_())).map(function (r) { return ledgerRowToArray_(tid, name, r, ts); });
  var sh = getSheet_(SHEETS.LEDGER), lr = sh.getLastRow();
  var existing = lr >= 2 ? sh.getRange(2, 1, lr - 1, LEDGER_COLS.length).getValues() : [];
  var others = existing.filter(function (r) { return String(r[0]).trim() !== tid && String(r[0]).trim() !== ''; });
  var merged = others.concat(mine);
  merged.sort(function (a, b) {
    var c = natCompare_(a[0], b[0]);
    if (c) return c;
    return toDate_(a[2]).getTime() - toDate_(b[2]).getTime();
  });
  writeLedgerSheet_(merged);
  SpreadsheetApp.flush();
}
