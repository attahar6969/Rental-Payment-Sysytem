/**
 * AllocationService.gs — the payment allocation engine.
 * All money is handled in integer PAISE. Payments are replayed in Rent Log order (entry order),
 * so a stored receipt can never be changed by later activity. Oldest outstanding month is always settled first.
 */

function newEngineState_(segments, mode, startIdx, openOutP, openAdvP) {
  var st = { segments: segments, mode: mode, startIdx: startIdx, months: {}, pool: [], billedThrough: startIdx - 1 };
  if (openOutP > 0) monthRec_(st, startIdx - 1).openP = openOutP;   // opening arrears sit on the month before ledger start
  if (openAdvP > 0) st.pool.push({ src: 'OPENING', date: null, amtP: openAdvP });
  return st;
}

function monthRec_(st, idx) {
  var m = st.months[idx];
  if (!m) {
    var lia = idx < st.startIdx ? { dueP: 0, rentP: 0 } : calculateMonthlyLiability(idx, st.segments, st.mode);
    m = st.months[idx] = { idx: idx, rentP: lia.rentP, dueP: lia.dueP, openP: 0, allocP: 0, allocs: [] };
  }
  return m;
}

function monthOut_(m) { return m.openP + m.dueP - m.allocP; }

/** Creates (bills) every month up to idx. Returns the set of months billed for the first time. */
function billThrough_(st, idx) {
  var newly = {};
  for (var i = st.billedThrough + 1; i <= idx; i++) { monthRec_(st, i); newly[i] = true; }
  if (idx > st.billedThrough) st.billedThrough = idx;
  return newly;
}

/** RULE: settle the oldest outstanding month first, up to (and including) upToIdx. */
function allocateOldestOutstandingFirst(st, upToIdx, amountP, src, date) {
  var byMonth = {}, left = amountP;
  for (var i = st.startIdx - 1; i <= upToIdx && left > 0; i++) {
    var m = st.months[i];
    if (!m) continue;
    var out = monthOut_(m);
    if (out <= 0) continue;
    var take = Math.min(out, left);
    m.allocP += take;
    m.allocs.push({ src: src, date: date, amtP: take });
    byMonth[i] = (byMonth[i] || 0) + take;
    left -= take;
  }
  return { byMonth: byMonth, allocatedP: amountP - left, leftP: left };
}

/** Uses banked advance credit (oldest chunk first) against outstanding months up to upToIdx. */
function applyAdvancePool_(st, upToIdx) {
  var byMonth = {}, total = 0;
  while (st.pool.length) {
    var ch = st.pool[0];
    var r = allocateOldestOutstandingFirst(st, upToIdx, ch.amtP, ch.src, ch.date);
    Object.keys(r.byMonth).forEach(function (k) { byMonth[k] = (byMonth[k] || 0) + r.byMonth[k]; });
    total += r.allocatedP;
    ch.amtP = r.leftP;
    if (ch.amtP === 0) st.pool.shift(); else break;
  }
  return { byMonth: byMonth, totalP: total };
}

/** Total outstanding of months strictly before beforeIdx. */
function calculateCarryForward(st, beforeIdx) {
  var t = 0;
  Object.keys(st.months).forEach(function (k) { if (+k < beforeIdx) t += Math.max(0, monthOut_(st.months[k])); });
  return t;
}

function calculateAdvanceCredit(st) {
  return st.pool.reduce(function (a, c) { return a + c.amtP; }, 0);
}

/**
 * Applies one payment { id, date, rentFrom, amountP } to the state and returns its breakdown.
 * The horizon month H = latest of (payment month, selected rent month): everything up to H is due.
 */
function processPayment_(st, pay) {
  var H = Math.max(monthIdx_(pay.date), monthIdx_(pay.rentTill || pay.rentFrom), st.startIdx);
  var newly = billThrough_(st, H), before = {}, i, m;
  for (i = st.startIdx - 1; i <= H; i++) { m = st.months[i]; if (m) before[i] = monthOut_(m); }

  var pool = applyAdvancePool_(st, H);
  var cash = allocateOldestOutstandingFirst(st, H, pay.amountP, pay.id, pay.date);
  if (cash.leftP > 0) st.pool.push({ src: pay.id, date: pay.date, amtP: cash.leftP });
  // waive off: settles the oldest outstanding rent without any cash (cannot exceed what is outstanding)
  var wv = pay.waiveP > 0 ? allocateOldestOutstandingFirst(st, H, pay.waiveP, pay.id, pay.date) : { allocatedP: 0, leftP: 0 };

  var rows = [], dueBefore = 0, prevOut = 0, after = 0, prevAlloc = 0, curAlloc = 0;
  for (i = st.startIdx - 1; i <= H; i++) {
    m = st.months[i];
    if (!m) continue;
    var b = before[i], a = cash.byMonth[i] || 0, close = monthOut_(m);
    dueBefore += b; after += close;
    if (i < H) { prevOut += b; prevAlloc += a; } else { curAlloc += a; }
    if (b > 0) rows.push({ idx: i, rentP: m.rentP, openP: newly[i] ? 0 : b, dueP: newly[i] ? m.dueP : 0,
      advP: pool.byMonth[i] || 0, allocP: a, closeP: close });
  }
  var status = after > 0 ? (cash.allocatedP > 0 ? 'PARTIAL' : 'UNPAID') : (cash.leftP > 0 ? 'ADVANCE' : 'PAID');
  return { H: H, rows: rows, dueBeforeP: dueBefore, prevOutP: prevOut, currentDueP: before[H] || 0,
    poolAppliedP: pool.totalP, waivedP: wv.allocatedP, waiveExcessP: wv.leftP, allocP: cash.allocatedP, prevAllocP: prevAlloc, curAllocP: curAlloc,
    advanceCreatedP: cash.leftP, outstandingAfterP: after, advanceBalanceP: calculateAdvanceCredit(st), status: status };
}

/** Brings the state up to today: bills months that have since started and spends banked advance. */
function finalizeSweep_(st, asOfIdx) {
  var H = Math.max(st.billedThrough, asOfIdx, st.startIdx);
  billThrough_(st, H);
  applyAdvancePool_(st, H);
}

/** Builds RENT LEDGER rows (paise) for one tenant from a finalized state. */
function buildLedgerRows_(st, asOfIdx) {
  var idxs = Object.keys(st.months).map(Number).sort(function (a, b) { return a - b; });
  var rows = [], cum = 0;
  idxs.forEach(function (i) {
    var m = st.months[i];
    if (m.openP + m.dueP === 0 && m.allocP === 0) return;
    var out = monthOut_(m), last = m.allocs.length ? m.allocs[m.allocs.length - 1] : null;
    cum += out;
    rows.push({ idx: i, rentP: m.rentP, openP: m.openP, dueP: m.dueP, allocP: m.allocP, closeP: out, advP: 0,
      status: out === 0 ? 'PAID' : (m.allocP > 0 ? 'PARTIAL' : 'UNPAID'),
      lastDate: last ? last.date : null, lastReceipt: last ? last.src : '', cumP: cum });
  });
  var adv = calculateAdvanceCredit(st);
  if (adv > 0) {
    if (!rows.length) rows.push({ idx: asOfIdx, rentP: 0, openP: 0, dueP: 0, allocP: 0, closeP: 0, advP: 0, status: 'PAID', lastDate: null, lastReceipt: '', cumP: 0 });
    rows[rows.length - 1].advP = adv;
  }
  return rows;
}

/** Runs a tenant's full history through the engine. extraPay (optional) is appended as the newest payment. */
function runTenantEngine_(ctx, tid, extraPay, sweep) {
  var op = ctx.opening[tid] || {};
  var st = newEngineState_(ctx.segsFor(tid), op.mode || ctx.cfg.proration, ctx.cfg.ledgerStartIdx, op.outP || 0, op.advP || 0);
  (ctx.logByTenant[tid] || []).forEach(function (p) { processPayment_(st, p); });
  var res = extraPay ? processPayment_(st, extraPay) : null;
  if (sweep) finalizeSweep_(st, monthIdx_(today_()));
  return { st: st, result: res };
}

/**
 * Read-only preview for the payment form. Never writes. req = { tenantId, paymentDate, amount, rentFrom } (ISO strings).
 */
function calculateAllocationPreview(req) {
  var errors = [];
  var v = validateTenant(req.tenantId);
  if (!v.ok) return { ok: false, errors: [v.message] };
  var payDate = toDate_(req.paymentDate), rentFrom = toDate_(req.rentFrom), rentTill = toDate_(req.rentTill);
  if (!payDate) errors.push('Enter a valid payment date.');
  if (!rentFrom) errors.push('Enter a valid Rent From date.');
  if (errors.length) return { ok: false, errors: errors };
  var amt = Number(req.amount);
  if (req.amount === '' || req.amount === null || req.amount === undefined || !isFinite(amt) || amt <= 0) {
    errors.push('Enter an amount greater than zero.'); amt = 0;
  }
  var ctx = buildContext_();
  if (!ctx.segsFor(v.tenant.id).length) return { ok: false, errors: ['No confirmed agreement record for this tenant. Run Refresh System, then confirm any pending revisions.'] };
  var run = runTenantEngine_(ctx, v.tenant.id, { id: 'PREVIEW', date: payDate, rentFrom: rentFrom, rentTill: rentTill, amountP: toPaise_(amt), waiveP: toPaise_(Number(req.waive) || 0) }, false);
  if (Number(req.waive) > 0 && !isAdmin_()) errors.push('Only an administrator can waive off an amount.');
  if (run.result.waiveExcessP > 0) errors.push('Waive off cannot be more than the remaining balance.');
  return { ok: errors.length === 0, errors: errors, preview: formatPreview_(run.result, amt) };
}

function formatPreview_(r, amount) {
  return {
    targetMonth: idxLabel_(r.H),
    rows: r.rows.map(function (x) {
      return { month: idxLabel_(x.idx), rent: rupees_(x.rentP), opening: rupees_(x.openP), due: rupees_(x.dueP),
        advance: rupees_(x.advP), allocated: rupees_(x.allocP), closing: rupees_(x.closeP) };
    }),
    totals: { payment: amount, allocated: rupees_(r.allocP), advanceUsed: rupees_(r.poolAppliedP),
      previousOutstanding: rupees_(r.prevOutP), waived: rupees_(r.waivedP), currentMonthDue: rupees_(r.currentDueP),
      outstandingBefore: rupees_(r.dueBeforeP - r.poolAppliedP), advanceCreated: rupees_(r.advanceCreatedP),
      outstandingAfter: rupees_(r.outstandingAfterP), advanceBalance: rupees_(r.advanceBalanceP) },
    status: r.status
  };
}
