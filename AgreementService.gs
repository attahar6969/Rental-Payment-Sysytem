/** AgreementService.gs — rent revision history. RENT HISTORY is the permanent source of agreement terms. */

function readHistory_() {
  var sh = getSheet_(SHEETS.HISTORY), lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, HIST_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, tid: String(r[H.TID]).trim(), name: r[H.NAME],
      prevP: r[H.PREV] === '' ? null : toPaise_(r[H.PREV]), newP: toPaise_(r[H.NEW] || 0),
      from: toDate_(r[H.FROM]), till: toDate_(r[H.TILL]), revDate: r[H.REVDATE], ref: String(r[H.REF]),
      by: r[H.BY], status: String(r[H.STATUS]).trim(), as: toDate_(r[H.AS]), ae: toDate_(r[H.AE]),
      type: String(r[H.TYPE]).trim() };
  }).filter(function (h) { return h.tid; });
}

function isSegmentRow_(h) {
  return (h.status === 'BASELINE' || h.status === 'CONFIRMED') &&
    (h.type === 'BASELINE' || h.type === 'RENT_REVISION' || h.type === 'NEW_AGREEMENT') && h.from && h.as && h.ae;
}

/**
 * Turns confirmed history rows into billing intervals:
 * each interval runs from max(effective-from, agreement start) to min(agreement end, day before next revision).
 */
function getSegments_(tenantId, histRows) {
  var rows = histRows.filter(function (h) { return h.tid === String(tenantId).trim() && isSegmentRow_(h); });
  rows.sort(function (a, b) { return (dayNum_(a.from) - dayNum_(b.from)) || (a.row - b.row); });
  var segs = [];
  rows.forEach(function (r, i) {
    var start = dayNum_(r.from) > dayNum_(r.as) ? r.from : r.as;
    var endN = dayNum_(r.ae);
    if (i + 1 < rows.length) endN = Math.min(endN, dayNum_(rows[i + 1].from) - 1);
    var startN = dayNum_(start);
    if (endN < startN) return;
    segs.push({ rentP: r.newP, startN: startN, endN: endN, start: start, end: new Date(start.getFullYear(), start.getMonth(), start.getDate() + (endN - startN)), ref: r.ref });
  });
  return segs;
}

function getRentHistory(tenantId) {
  return readHistory_().filter(function (h) { return h.tid === String(tenantId).trim(); }).map(function (h) {
    return { ref: h.ref, type: h.type, status: h.status, prev: h.prevP === null ? null : rupees_(h.prevP), rent: rupees_(h.newP),
      from: isoDate_(h.from), till: isoDate_(h.till), start: isoDate_(h.as), end: isoDate_(h.ae) };
  });
}

/** Monthly rent applicable on a given date (0 if no agreement covers it). */
function getApplicableRent(tenantId, date) {
  var n = dayNum_(toDate_(date)), rent = 0;
  getSegments_(tenantId, readHistory_()).forEach(function (s) { if (n >= s.startN && n <= s.endN) rent = rupees_(s.rentP); });
  return rent;
}

/**
 * Compares the Parent Workbook with the last recorded agreement for every tenant.
 * Returns the list of differences (does not write anything).
 */
function detectRentRevision(parentTenants, hist) {
  var byT = groupBy_(hist, function (h) { return h.tid; });
  var out = [];
  parentTenants.forEach(function (t) {
    if (t.invalid) return;
    var rows = byT[t.id] || [];
    var segRows = rows.filter(isSegmentRow_);
    if (!segRows.length) { out.push({ tenant: t, type: 'BASELINE' }); return; }
    segRows.sort(function (a, b) { return (dayNum_(a.from) - dayNum_(b.from)) || (a.row - b.row); });
    var last = segRows[segRows.length - 1], type = null;
    if (dayNum_(last.as) !== dayNum_(t.start)) type = 'NEW_AGREEMENT';
    else if (last.newP !== toPaise_(t.rent)) type = 'RENT_REVISION';
    else if (dayNum_(last.ae) !== dayNum_(t.end)) type = 'END_DATE_CHANGE';
    out.push({ tenant: t, type: type, last: last, pending: rows.filter(function (r) { return r.status === 'PENDING'; }) });
  });
  return out;
}

/**
 * Writes baselines for new tenants and flags changes as PENDING. Never touches confirmed amounts.
 * Returns { baselines, flagged }.
 */
function syncAgreementHistory_() {
  return withLock_(function () {
    var tenants = getParentTenants_(true), hist = readHistory_();
    var changes = detectRentRevision(tenants, hist);
    var now = new Date(), by = userEmail_(), add = [], supersede = [], flagged = 0, baselines = 0, seq = 0;
    changes.forEach(function (c) {
      var t = c.tenant;
      if (c.type === 'BASELINE') {
        add.push([t.id, t.name, '', t.rent, t.start, '', now, 'BASE-' + t.id, by, 'BASELINE', t.start, t.end, 'BASELINE']);
        baselines++; return;
      }
      if (!c.type) { c.pending.forEach(function (p) { supersede.push(p.row); }); return; }
      var same = c.pending.filter(function (p) {
        return p.type === c.type && p.as && p.ae && dayNum_(p.as) === dayNum_(t.start) && dayNum_(p.ae) === dayNum_(t.end) && p.newP === toPaise_(t.rent);
      });
      if (same.length) return;
      c.pending.forEach(function (p) { supersede.push(p.row); });
      var ref = 'REV-' + t.id + '-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + (++seq);
      add.push([t.id, t.name, rupees_(c.last.newP), t.rent, c.type === 'NEW_AGREEMENT' ? t.start : '', '', now, ref, by, 'PENDING', t.start, t.end, c.type]);
      flagged++;
    });
    var sh = getSheet_(SHEETS.HISTORY);
    supersede.forEach(function (row) { sh.getRange(row, H.STATUS + 1).setValue('SUPERSEDED'); });
    if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, HIST_COLS.length).setValues(add);
    SpreadsheetApp.flush();
    return { baselines: baselines, flagged: flagged };
  });
}

function getPendingRevisions() {
  var hist = readHistory_(), parent = {};
  try { getParentTenants_(false).forEach(function (t) { parent[t.id] = t; }); } catch (e) { }
  return hist.filter(function (h) { return h.status === 'PENDING'; }).map(function (h) {
    var t = parent[h.tid];
    return { ref: h.ref, tid: h.tid, name: t ? t.name : h.name, property: t ? t.property : '', type: h.type,
      prev: h.prevP === null ? null : rupees_(h.prevP), rent: rupees_(h.newP),
      start: isoDate_(h.as), end: isoDate_(h.ae), suggested: isoDate_(h.from), isAdmin: isAdmin_() };
  });
}

/** Rules for an effective date. Returns { ok, message, needsConfirm }. */
function validateEffectiveDate(pending, eff, lastSeg) {
  var cfg = getSystemConfig();
  if (!eff) return { ok: false, message: 'Enter a valid effective date.' };
  if (dayNum_(eff) < dayNum_(pending.as)) return { ok: false, message: 'Effective date cannot be before the agreement start (' + fmtDate_(pending.as) + ').' };
  if (dayNum_(eff) > dayNum_(pending.ae)) return { ok: false, message: 'Effective date cannot be after the agreement end (' + fmtDate_(pending.ae) + ').' };
  if (lastSeg && dayNum_(eff) <= dayNum_(lastSeg.from)) return { ok: false, message: 'Effective date must be after the previous revision (' + fmtDate_(lastSeg.from) + ').' };
  if (monthIdx_(eff) > monthIdx_(today_()) + cfg.maxAdvanceMonths) return { ok: false, message: 'Effective date is too far in the future.' };
  var curStart = idxStart_(monthIdx_(today_()));
  if (dayNum_(eff) < dayNum_(curStart)) {
    return { ok: true, needsConfirm: true,
      message: 'This date is in a past period. Ledger months from ' + idxLabel_(monthIdx_(eff)) + ' to ' + idxLabel_(monthIdx_(today_())) +
        ' will be recalculated at the new rent. Old receipts stay unchanged.' };
  }
  return { ok: true, needsConfirm: false };
}

/** Admin confirms a flagged change. effIso = yyyy-MM-dd (not needed for END_DATE_CHANGE). */
function recordRentRevision(ref, effIso, confirmPast) {
  requireAdmin_();
  var result = withLock_(function () {
    var hist = readHistory_();
    var p = hist.filter(function (h) { return h.ref === ref && h.status === 'PENDING'; })[0];
    if (!p) return { ok: false, message: 'This change is no longer pending. Run Refresh System.' };
    var t = getTenantDetails(p.tid);
    if (!t || dayNum_(t.start) !== dayNum_(p.as) || dayNum_(t.end) !== dayNum_(p.ae) || toPaise_(t.rent) !== p.newP) {
      return { ok: false, message: 'The Parent Workbook changed again after this was flagged. Run Refresh System first.' };
    }
    var sh = getSheet_(SHEETS.HISTORY), now = new Date(), by = userEmail_();
    if (p.type === 'END_DATE_CHANGE') {
      hist.forEach(function (h) {
        if (h.tid === p.tid && h.as && dayNum_(h.as) === dayNum_(p.as) && isSegmentRow_(h)) sh.getRange(h.row, H.AE + 1).setValue(p.ae);
      });
      sh.getRange(p.row, H.REVDATE + 1, 1, 4).setValues([[now, p.ref, by, 'APPLIED']]);
    } else {
      var eff = effIso ? toDate_(effIso) : (p.type === 'NEW_AGREEMENT' ? p.as : null);
      var lastSeg = hist.filter(function (h) { return h.tid === p.tid && isSegmentRow_(h); })
        .sort(function (a, b) { return (dayNum_(a.from) - dayNum_(b.from)) || (a.row - b.row); }).pop();
      var v = validateEffectiveDate(p, eff, lastSeg);
      if (!v.ok) return { ok: false, message: v.message };
      if (v.needsConfirm && !confirmPast) return { ok: false, needsConfirm: true, message: v.message };
      sh.getRange(p.row, H.FROM + 1).setValue(eff);
      sh.getRange(p.row, H.REVDATE + 1, 1, 4).setValues([[now, p.ref, by, 'CONFIRMED']]);
      if (lastSeg) {
        var newTill = new Date(eff.getFullYear(), eff.getMonth(), eff.getDate() - 1);
        sh.getRange(lastSeg.row, H.TILL + 1).setValue(newTill);
      }
    }
    SpreadsheetApp.flush();
    return { ok: true, tid: p.tid };
  });
  if (result.ok) { try { rebuildLedgerInternal_(); } catch (e) { logError_('recordRentRevision.rebuild', e); } }
  return result;
}

function rejectRentRevision(ref) {
  requireAdmin_();
  return withLock_(function () {
    var p = readHistory_().filter(function (h) { return h.ref === ref && h.status === 'PENDING'; })[0];
    if (!p) return { ok: false, message: 'This change is no longer pending.' };
    getSheet_(SHEETS.HISTORY).getRange(p.row, H.STATUS + 1).setValue('REJECTED');
    return { ok: true, note: 'Rejected. The Parent Workbook still differs from the recorded terms, so it will be flagged again on the next refresh unless you correct the Parent Workbook.' };
  });
}
