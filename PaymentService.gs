/** PaymentService.gs — validation and the only code path that writes to RENT LOG. */

function readLog_() {
  var sh = getSheet_(SHEETS.LOG), lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2, 1, lr - 1, LOG_COLS.length).getValues().map(function (r, i) {
    return { row: i + 2, receipt: String(r[L.RECEIPT]).trim(), date: toDate_(r[L.DATE]), tid: String(r[L.TID]).trim(),
      name: r[L.NAME], phone: r[L.PHONE], amountP: toPaise_(r[L.AMT] || 0), rentFrom: toDate_(r[L.RFROM]),
      rentTill: toDate_(r[L.RTILL]), outB: r[L.OUTB], outA: r[L.OUTA], status: r[L.STATUS], ts: r[L.TS],
      txn: String(r[L.TXN]).trim(), waiveP: toPaise_(r[L.WAIVE] || 0), raw: r };
  }).filter(function (p) { return p.receipt && p.date; });
}

/** Data the Entry form needs on open. */
function getEntryInitData() {
  var cfg = getSystemConfig();
  return { tenants: getActiveTenants(), today: isoDate_(today_()), business: cfg.business.name,
    maxAdvanceMonths: cfg.maxAdvanceMonths, isAdmin: isAdmin_() };
}

/** Server-side validation of a submission. Returns { ok, errors[], tenant, payDate, rentFrom, amount }. */
function validatePayment(p) {
  var errors = [], cfg = getSystemConfig();
  var v = validateTenant(p && p.tenantId);
  if (!v.ok) errors.push(v.message);
  var payDate = toDate_(p && p.paymentDate), rentFrom = toDate_(p && p.rentFrom), rentTill = toDate_(p && p.rentTill);
  if (!payDate) errors.push('Payment date is not valid.');
  else if (dayNum_(payDate) > dayNum_(today_())) errors.push('Payment date cannot be in the future.');
  if (!rentFrom) errors.push('Rent From date is not valid.');
  else if (monthIdx_(rentFrom) > monthIdx_(today_()) + cfg.maxAdvanceMonths) errors.push('Rent period is more than ' + cfg.maxAdvanceMonths + ' months ahead.');
  if (!rentTill) errors.push('Rent Till date is not valid.');
  else if (rentFrom && dayNum_(rentTill) < dayNum_(rentFrom)) errors.push('Rent Till cannot be before Rent From.');
  else if (monthIdx_(rentTill) > monthIdx_(today_()) + cfg.maxAdvanceMonths) errors.push('Rent period is more than ' + cfg.maxAdvanceMonths + ' months ahead.');
  var amount = Number(p && p.amount);
  if (p === null || p === undefined || p.amount === '' || !isFinite(amount) || amount <= 0) errors.push('Amount received must be greater than zero.');
  else if (Math.abs(amount * 100 - Math.round(amount * 100)) > 1e-6) errors.push('Amount can have at most 2 decimal places.');
  else if (amount > 100000000) errors.push('Amount is unrealistically large.');
  var waive = Number(p && p.waive) || 0, waiveReason = String((p && p.waiveReason) || '').trim();
if (waive < 0) errors.push('Waive off amount cannot be negative.');
else if (waive > 0) {
  if (!isAdmin_()) errors.push('Only an administrator can waive off an amount.');
  if (waiveReason.length < 3) errors.push('Waive off reason is compulsory.');
}
  if (!p || !/^TXN-[A-Za-z0-9\-]{8,64}$/.test(String(p.token || ''))) errors.push('Missing submission token. Close and reopen the form.');
  return { ok: errors.length === 0, errors: errors, tenant: v.tenant, payDate: payDate, rentFrom: rentFrom, rentTill: rentTill, amount: amount, waive: waive, waiveReason: waiveReason };
}

/** Year token used in receipt IDs. */
function receiptYearToken_(date, mode) {
  var y = date.getFullYear();
  if (mode === 'FINANCIAL') {
    var start = date.getMonth() >= 3 ? y : y - 1;
    return start + '-' + pad2_((start + 1) % 100);
  }
  return String(y);
}

/** Next RENT-YYYY-NNN. Must be called while holding the lock. */
function generateReceiptId(cfg, payDate) {
  var sh = getSheet_(SHEETS.LOG), lr = sh.getLastRow();
  var head = cfg.prefix + '-' + receiptYearToken_(payDate, cfg.yearMode) + '-';
  var max = 0;
  if (lr >= 2) {
    sh.getRange(2, 1, lr - 1, 1).getValues().forEach(function (r) {
      var s = String(r[0]);
      if (s.indexOf(head) === 0) { var n = parseInt(s.slice(head.length), 10); if (n > max) max = n; }
    });
  }
  var n2 = max + 1;
  return head + (n2 < 10 ? '00' : n2 < 100 ? '0' : '') + n2;
}

/** Outstanding for a tenant as of today (info helper). */
function calculateOutstanding(tenantId) {
  var ctx = buildContext_(), run = runTenantEngine_(ctx, String(tenantId).trim(), null, true);
  var total = 0;
  Object.keys(run.st.months).forEach(function (k) { total += Math.max(0, monthOut_(run.st.months[k])); });
  return { outstanding: rupees_(total), advance: rupees_(calculateAdvanceCredit(run.st)) };
}

/** Appends exactly one row to RENT LOG in a single write. Caller holds the lock. */
function savePaymentTransaction(row) {
  var sh = getSheet_(SHEETS.LOG), target = sh.getLastRow() + 1;
  sh.getRange(target, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
  if (String(sh.getRange(target, 1).getValue()) !== String(row[L.RECEIPT])) throw new Error('The payment row could not be verified after saving. Nothing was recorded.');
  sh.getRange(target, L.DATE + 1).setNumberFormat(DATE_FORMAT);
  [L.RENT, L.AMT, L.OUTB, L.ALPREV, L.ALCUR, L.ADV, L.OUTA].forEach(function (c) { sh.getRange(target, c + 1).setNumberFormat(CURRENCY_FORMAT); });
  [L.AS, L.AE, L.RFROM, L.RTILL].forEach(function (c) { sh.getRange(target, c + 1).setNumberFormat(DATE_FORMAT); });
  sh.getRange(target, L.TS + 1).setNumberFormat(DATETIME_FORMAT);
  return target;
}

/**
 * Records a payment. Everything is recomputed on the server; the client's preview signature is only used
 * to detect that the numbers changed since the operator reviewed them.
 */
function submitPayment(payload) {
  var v;
  try {
    v = validatePayment(payload);
    if (!v.ok) return { ok: false, errors: v.errors, message: v.errors.join(' ') };
    var payMode = String(payload.paymentMode || '').trim();
    if (['Cash', 'Online'].indexOf(payMode) < 0) return { ok: false, message: 'Please select Payment Mode (Cash or Online).' };
    return withLock_(function () {
      var cfg = getSystemConfig(true), sh = getSheet_(SHEETS.LOG);
      // duplicate submission protection: same token => return the receipt already created
      var existing = readLog_().filter(function (p) { return p.txn === payload.token; })[0];
      if (existing) return { ok: true, duplicate: true, receiptId: existing.receipt, message: 'This payment was already recorded as ' + existing.receipt + '.' };

      var t = v.tenant, ctx = buildContext_();
      if (!ctx.segsFor(t.id).length) return { ok: false, message: 'No confirmed agreement record for this tenant. Run Refresh System and confirm pending revisions.' };
      var receiptId = generateReceiptId(cfg, v.payDate);
      var run = runTenantEngine_(ctx, t.id, { id: receiptId, date: v.payDate, rentFrom: v.rentFrom, rentTill: v.rentTill, amountP: toPaise_(v.amount), waiveP: toPaise_(v.waive) }, false);
      var r = run.result, preview = formatPreview_(r, v.amount);
      if (r.waiveExcessP > 0) return { ok: false, message: 'Waive off cannot be more than the remaining balance.' };

      var sig = [preview.totals.allocated, preview.totals.outstandingAfter, preview.totals.advanceCreated].join('|');
      if (payload.previewSig && payload.previewSig !== sig) {
        return { ok: false, stale: true, preview: preview, message: 'Figures changed since you reviewed the allocation (another entry may have been saved). Please review the updated allocation and submit again.' };
      }
      var now = new Date(), rentTill = v.rentTill;
      var detail = { v: 1, biz: cfg.business, targetMonth: preview.targetMonth, rows: preview.rows, totals: preview.totals };
      var row = [receiptId, v.payDate, t.id, t.name, t.phone, t.rent, t.start, t.end, v.amount, v.rentFrom, rentTill,
        rupees_(r.dueBeforeP - r.poolAppliedP), rupees_(r.prevAllocP), rupees_(r.curAllocP), rupees_(r.advanceCreatedP),
        rupees_(r.outstandingAfterP), r.status, now, userEmail_(), payload.token, 'ALLOC-' + receiptId, JSON.stringify(detail), t.property, v.waive || 0, v.waiveReason, payMode];
      savePaymentTransaction(row);

      var warning = '';
      try { updateTenantLedger(t.id); } catch (e) { logError_('submitPayment.ledger', e); warning = 'Payment saved. The ledger will fully update on the next Refresh System.'; }
      return { ok: true, receiptId: receiptId, status: r.status, amount: v.amount, outstandingAfter: preview.totals.outstandingAfter,
        advance: preview.totals.advanceCreated, tenantName: t.name, warning: warning };
    }, 30000);
  } catch (e) {
    logError_('submitPayment', e);
    return { ok: false, message: e.message || String(e) };
  }
}
