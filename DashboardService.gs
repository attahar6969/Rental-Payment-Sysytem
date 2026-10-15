/** DashboardService.gs — derived reporting. Reads RENT LOG, ledger engine output and the parent register. */

function monthFromFilter_(s) { var d = toDate_(String(s || '') + '-01'); return d ? monthIdx_(d) : monthIdx_(today_()); }

function getTenantSummary(filters, all) {
  filters = filters || {};
  all = all || computeAllLedgers_();
  var M = filters.month ? monthFromFilter_(filters.month) : monthIdx_(today_()), parent = {};
  try { getParentTenants_(false).forEach(function (t) { parent[t.id] = t; }); } catch (e) { }
  var out = [];
  all.tenants.forEach(function (t) {
    var p = parent[t.tid];
    var agr = p ? (p.active ? 'ACTIVE' : 'INACTIVE') : 'INACTIVE';
    if (filters.tenantId && filters.tenantId !== t.tid) return;
    if (filters.agreementStatus && filters.agreementStatus !== agr) return;
    var due = 0, alloc = 0, rem = 0, cur = 0, curAlloc = 0;
    t.rows.forEach(function (r) {
      if (r.idx > M) return;
      due += r.openP + r.dueP; alloc += r.allocP; rem += r.closeP;
      if (r.idx === M) { cur = r.dueP; curAlloc = r.allocP; }
    });
    var prev = calculateCarryForward(t.st, M);
    var adv = calculateAdvanceCredit(t.st);
    var status = rem <= 0 ? (adv > 0 ? 'ADVANCE' : 'PAID') : (curAlloc > 0 ? 'PARTIAL' : 'UNPAID');
    if (filters.status && filters.status !== status && !(filters.status === 'PAID' && status === 'ADVANCE')) return;
    var lastDate = null;
    (all.ctx.logByTenant[t.tid] || []).forEach(function (x) { if (!lastDate || x.date > lastDate) lastDate = x.date; });
    var rentNow = 0;
    var segs = all.ctx.segsFor(t.tid), today = dayNum_(today_());
    segs.forEach(function (s) { if (today >= s.startN && today <= s.endN) rentNow = s.rentP; });
    if (!rentNow && p) rentNow = toPaise_(p.rent || 0);
    out.push({ tid: t.tid, name: t.name, phone: p ? p.phone : '', agreementRent: rupees_(rentNow), currentMonthRent: rupees_(cur),
      previousOutstanding: rupees_(prev), totalDue: rupees_(due), totalReceived: rupees_(alloc), remaining: rupees_(rem),
      advance: rupees_(adv), status: status, lastPaymentDate: lastDate ? isoDate_(lastDate) : '',
      agreementEnd: p && p.end ? isoDate_(p.end) : '', agreementStatus: agr });
  });
  return out;
}

function getCollectionSummary(filters, all) {
  filters = filters || {};
  var from = toDate_(filters.from), to = toDate_(filters.to);
  var logs = readLog_().filter(function (p) {
    if (from && dayNum_(p.date) < dayNum_(from)) return false;
    if (to && dayNum_(p.date) > dayNum_(to)) return false;
    if (filters.tenantId && p.tid !== filters.tenantId) return false;
    if (filters.month && monthIdx_(p.rentFrom || p.date) !== monthFromFilter_(filters.month)) return false;
    if (filters.status && String(p.status) !== filters.status) return false;
    return true;
  });
  var byMonth = {}, byTenant = {};
  logs.forEach(function (p) {
    var k = idxIso_(monthIdx_(p.date));
    byMonth[k] = (byMonth[k] || 0) + p.amountP;
    byTenant[p.name] = (byTenant[p.name] || 0) + p.amountP;
  });
  var months = Object.keys(byMonth).sort(), cum = 0;
  var outByMonth = {};
  (all || computeAllLedgers_()).tenants.forEach(function (t) {
    if (filters.tenantId && filters.tenantId !== t.tid) return;
    t.rows.forEach(function (r) { var k = idxIso_(r.idx); outByMonth[k] = (outByMonth[k] || 0) + r.closeP; });
  });
  return {
    logs: logs, monthlyCollection: months.map(function (m) { return { month: m, amount: rupees_(byMonth[m]) }; }),
    trend: months.map(function (m) { cum += byMonth[m]; return { month: m, amount: rupees_(cum) }; }),
    monthlyOutstanding: Object.keys(outByMonth).sort().map(function (m) { return { month: m, amount: rupees_(outByMonth[m]) }; }),
    tenantCollection: Object.keys(byTenant).sort().map(function (n) { return { name: n, amount: rupees_(byTenant[n]) }; })
  };
}

function getDashboardData(filters) {
  filters = filters || {};
  var all = computeAllLedgers_();
  var tenants = getTenantSummary(filters, all), coll = getCollectionSummary(filters, all);
  var parentActive = 0;
  try { parentActive = getParentTenants_(false).filter(function (t) { return t.active; }).length; } catch (e) { }
  var M = filters.month ? monthFromFilter_(filters.month) : monthIdx_(today_());
  var expected = 0;
  all.tenants.forEach(function (t) {
    if (filters.tenantId && filters.tenantId !== t.tid) return;
    t.rows.forEach(function (r) { if (r.idx <= M) expected += r.openP + r.dueP; });
  });
  var sum = function (k) { return rupees_(tenants.reduce(function (a, x) { return a + toPaise_(x[k]); }, 0)); };
  var received = coll.logs.reduce(function (a, p) { return a + p.amountP; }, 0);
  var count = function (s) { return tenants.filter(function (x) { return x.status === s; }).length; };
  var tx = coll.logs.slice().sort(function (a, b) { return b.row - a.row; }).slice(0, 100).map(function (p) {
    var mon = ''; try { var d = JSON.parse(p.raw[L.JSON] || '{}'); mon = (d.rows || []).filter(function (r) { return r.allocated > 0; }).map(function (r) { return r.month; }).join(', '); } catch (e) { }
    return { receiptId: p.receipt, date: isoDate_(p.date), tenant: p.name, amount: rupees_(p.amountP), allocation: mon || idxLabel_(monthIdx_(p.rentFrom)),
      remaining: Number(p.outA), status: p.status };
  });
  var tenantOptions = tenants.map(function (t) { return { id: t.tid, name: t.name }; });
  return {
    generated: fmtDate_(today_()),
    kpis: { activeTenants: parentActive, expected: rupees_(expected), received: rupees_(received), outstanding: sum('remaining'),
      advance: sum('advance'), paid: count('PAID') + count('ADVANCE'), partial: count('PARTIAL'), unpaid: count('UNPAID') },
    tenants: tenants, transactions: tx,
    charts: { monthlyCollection: coll.monthlyCollection, monthlyOutstanding: coll.monthlyOutstanding, tenantCollection: coll.tenantCollection,
      trend: coll.trend, status: { Paid: count('PAID') + count('ADVANCE'), Partial: count('PARTIAL'), Unpaid: count('UNPAID') } },
    tenantOptions: tenantOptions
  };
}

/** Rewrites the DASHBOARD sheet (values, colours and native charts). */
function refreshDashboard() {
  var sh = getSheet_(SHEETS.DASH), d = getDashboardData({});
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1').setValue('RENT COLLECTION DASHBOARD').setFontSize(16).setFontWeight('bold').setFontColor('#1f2a44');
  sh.getRange('A2').setValue('Updated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATETIME_FORMAT) + '  |  use menu > Refresh System to update').setFontColor('#777');

  var k = d.kpis;
  var cards = [['Active Tenants', k.activeTenants, '#1f2a44'], ['Total Expected Rent', k.expected, '#1f2a44'], ['Total Rent Received', k.received, '#1b7f4b'],
    ['Total Outstanding', k.outstanding, '#c0392b'], ['Advance Credit', k.advance, '#2563a8'], ['Fully Paid', k.paid, '#1b7f4b'],
    ['Partial Payment', k.partial, '#d9822b'], ['Unpaid', k.unpaid, '#c0392b']];
  cards.forEach(function (c, i) {
    var col = i + 1;
    sh.getRange(4, col).setValue(c[0]).setFontColor('#fff').setBackground(c[2]).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    var v = sh.getRange(5, col).setValue(c[1]).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(c[2]).setBackground('#f3f5f9');
    if (i >= 1 && i <= 4) v.setNumberFormat(CURRENCY_FORMAT);
  });
  sh.setRowHeight(4, 34);

  var head = ['Tenant Name', 'Phone Number', 'Agreement Rent', 'Current Month Rent', 'Previous Outstanding', 'Total Due', 'Total Received',
    'Remaining Balance', 'Advance Credit', 'Payment Status', 'Last Payment Date', 'Agreement End Date'];
  sh.getRange(7, 1, 1, head.length).setValues([head]).setBackground('#1f2a44').setFontColor('#fff').setFontWeight('bold').setWrap(true);
  var rows = d.tenants.map(function (t) {
    return [t.name, t.phone, t.agreementRent, t.currentMonthRent, t.previousOutstanding, t.totalDue, t.totalReceived, t.remaining, t.advance,
      t.status, t.lastPaymentDate ? toDate_(t.lastPaymentDate) : '', t.agreementEnd ? toDate_(t.agreementEnd) : ''];
  });
  if (rows.length) {
    sh.getRange(8, 1, rows.length, head.length).setValues(rows);
    sh.getRange(8, 3, rows.length, 7).setNumberFormat(CURRENCY_FORMAT);
    sh.getRange(8, 11, rows.length, 2).setNumberFormat(DATE_FORMAT);
    var colors = { PAID: '#d8f0e2', ADVANCE: '#d6e6f7', PARTIAL: '#fbe6cf', UNPAID: '#f6d5d1' };
    sh.getRange(8, 10, rows.length, 1).setBackgrounds(rows.map(function (r) { return [colors[r[9]] || '#fff']; }));
    sh.getRange(8, 8, rows.length, 1).setFontColor('#c0392b');
    sh.getRange(8, 7, rows.length, 1).setFontColor('#1b7f4b');
    sh.getRange(8, 9, rows.length, 1).setFontColor('#2563a8');
  }

  var t0 = 8 + Math.max(rows.length, 1) + 2;
  sh.getRange(t0, 1).setValue('RECENT TRANSACTIONS').setFontWeight('bold');
  var th = ['Receipt ID', 'Payment Date', 'Tenant', 'Amount Received', 'Rent Month Allocation', 'Remaining Outstanding', 'Payment Status'];
  sh.getRange(t0 + 1, 1, 1, th.length).setValues([th]).setBackground('#1f2a44').setFontColor('#fff').setFontWeight('bold');
  var tx = d.transactions.slice(0, 50).map(function (x) { return [x.receiptId, toDate_(x.date), x.tenant, x.amount, x.allocation, x.remaining, x.status]; });
  if (tx.length) {
    sh.getRange(t0 + 2, 1, tx.length, th.length).setValues(tx);
    sh.getRange(t0 + 2, 2, tx.length, 1).setNumberFormat(DATE_FORMAT);
    sh.getRange(t0 + 2, 4, tx.length, 1).setNumberFormat(CURRENCY_FORMAT).setFontColor('#1b7f4b');
    sh.getRange(t0 + 2, 6, tx.length, 1).setNumberFormat(CURRENCY_FORMAT).setFontColor('#c0392b');
  }
  sh.setColumnWidths(1, 12, 130); sh.setColumnWidth(1, 170); sh.setColumnWidth(3, 170);

  // chart data (columns AA onward)
  var c = d.charts, base = 27;
  function block(col, title, hdr, list, mapFn) {
    sh.getRange(1, col, 1, 2).setValues([[title, '']]).setFontColor('#999');
    sh.getRange(2, col, 1, 2).setValues([hdr]).setFontColor('#999');
    if (list.length) sh.getRange(3, col, list.length, 2).setValues(list.map(mapFn));
    return list.length;
  }
  var n1 = block(base, 'Monthly Rent Collection', ['Month', 'Collected'], c.monthlyCollection, function (x) { return [x.month, x.amount]; });
  var n2 = block(base + 3, 'Monthly Outstanding', ['Month', 'Outstanding'], c.monthlyOutstanding, function (x) { return [x.month, x.amount]; });
  var n3 = block(base + 6, 'Tenant-wise Collection', ['Tenant', 'Collected'], c.tenantCollection, function (x) { return [x.name, x.amount]; });
  var st = [['Paid', c.status.Paid], ['Partial', c.status.Partial], ['Unpaid', c.status.Unpaid]];
  block(base + 9, 'Paid vs Partial vs Unpaid', ['Status', 'Tenants'], st, function (x) { return x; });
  var n5 = block(base + 12, 'Collection Trend', ['Month', 'Cumulative'], c.trend, function (x) { return [x.month, x.amount]; });

  function add(type, col, n, title, row, ccol, color) {
    if (!n) return;
    var b = sh.newChart().setChartType(type).addRange(sh.getRange(2, col, n + 1, 2)).setPosition(row, ccol, 0, 0)
      .setOption('title', title).setOption('width', 470).setOption('height', 260).setOption('legend', { position: 'none' });
    if (color) b.setOption('colors', color);
    sh.insertChart(b.build());
  }
  add(Charts.ChartType.COLUMN, base, n1, 'Monthly Rent Collection', 3, 14, ['#1b7f4b']);
  add(Charts.ChartType.COLUMN, base + 3, n2, 'Monthly Outstanding', 3, 21, ['#c0392b']);
  add(Charts.ChartType.BAR, base + 6, n3, 'Tenant-wise Collection', 17, 14, ['#2563a8']);
  add(Charts.ChartType.PIE, base + 9, 3, 'Paid vs Partial vs Unpaid', 17, 21);
  add(Charts.ChartType.LINE, base + 12, n5, 'Collection Trend', 31, 14, ['#1f2a44']);
  return d.kpis;
}
