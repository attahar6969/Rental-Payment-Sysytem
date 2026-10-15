/** TenantService.gs — reads tenant/agreement master data from the Parent Workbook (read-only). */

/** Verifies each configured column against its header; falls back to a header-keyword search. */
function resolveParentColumns_(headers, cfg) {
  var norm = headers.map(function (h) { return normText_(h).toLowerCase(); });
  var map = {};
  Object.keys(PARENT_FIELDS).forEach(function (f) {
    var def = PARENT_FIELDS[f];
    function matches(i) { return i >= 0 && i < norm.length && def.kw.some(function (k) { return norm[i].indexOf(k) >= 0; }); }
    var idx = cfg.cols[f] ? letterToIdx_(cfg.cols[f]) : -1;
    if (!matches(idx)) {
      idx = -1;
      for (var i = 0; i < norm.length && idx < 0; i++) if (matches(i)) idx = i;
    }
    if (idx < 0) {
      if (def.optional) return;
      throw new Error('Parent sheet column for "' + f + '" not found. Expected header containing: ' + def.kw.join(' / ') +
        '. Check the COL_* rows in SETTINGS.');
    }
    map[f] = idx;
  });
  return map;
}

function isActiveStatus_(s) {
  var t = normText_(s).toLowerCase();
  return t.indexOf('active') >= 0 && t.indexOf('inactive') < 0;
}

/** All rows of the parent sheet as tenant objects. Cached briefly. */
function getParentTenants_(fresh) {
  var cache = CacheService.getScriptCache();
  if (!fresh) {
    var c = cache.get('PARENT_TENANTS_V1');
    if (c) return JSON.parse(c).map(rehydrateTenant_);
  }
  var cfg = getSystemConfig();
  if (!cfg.parentId) throw new Error('PARENT_SPREADSHEET_ID is empty in SETTINGS.');
  var sh;
  try { sh = SpreadsheetApp.openById(cfg.parentId).getSheetByName(cfg.parentSheet); }
  catch (e) { throw new Error('Cannot open the Parent Workbook. Check the ID and that your account can view it. (' + e.message + ')'); }
  if (!sh) throw new Error('Parent sheet "' + cfg.parentSheet + '" not found.');
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];
  var data = sh.getRange(1, 1, lr, lc).getValues();
  var m = resolveParentColumns_(data[0], cfg);
  var seen = {}, out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i], id = normText_(r[m.id]);
    if (!id) continue;
    if (seen[id]) throw new Error('Duplicate Code "' + id + '" in the parent sheet (row ' + (i + 1) + '). Each property Code must be unique.');
    seen[id] = true;
    var rent = parseMoney_(r[m.rent]), start = toDate_(r[m.start]), end = toDate_(r[m.end]);
    var t = {
      id: id, name: normText_(r[m.name]), phone: normText_(r[m.phone]),
      rent: rent, start: start, end: end,
      statusText: normText_(r[m.status]), active: isActiveStatus_(r[m.status]),
      property: normText_(r[m.property]), landlord: m.landlord !== undefined ? normText_(r[m.landlord]) : '',
      row: i + 1, invalid: ''
    };
    if (!t.name) t.invalid = 'Tenant name is blank';
    else if (rent === null || rent <= 0) t.invalid = 'Rent is missing or zero';
    else if (!start || !end) t.invalid = 'Agreement start/end date is missing';
    else if (dayNum_(end) < dayNum_(start)) t.invalid = 'Agreement end is before start';
    out.push(t);
  }
  try { cache.put('PARENT_TENANTS_V1', JSON.stringify(out.map(dehydrateTenant_)), 120); } catch (e) { }
  return out;
}

function dehydrateTenant_(t) {
  var c = JSON.parse(JSON.stringify(t));
  c.start = isoDate_(t.start); c.end = isoDate_(t.end); return c;
}
function rehydrateTenant_(c) { c.start = toDate_(c.start); c.end = toDate_(c.end); return c; }

function clearTenantCache_() { CacheService.getScriptCache().remove('PARENT_TENANTS_V1'); }

/** Active tenants for the payment form (client-safe: dates as yyyy-MM-dd). */
function getActiveTenants() {
  var pendingIds = {};
  readHistory_().forEach(function (h) { if (h.status === 'PENDING') pendingIds[h.tid] = true; });
  return getParentTenants_(false).filter(function (t) { return t.active && !t.invalid; }).map(function (t) {
    return { id: t.id, name: t.name, phone: t.phone, rent: t.rent, start: isoDate_(t.start), end: isoDate_(t.end),
      property: t.property, pendingReview: !!pendingIds[t.id] };
  });
}

function getTenantDetails(tenantId) {
  var found = getParentTenants_(false).filter(function (t) { return t.id === String(tenantId).trim(); })[0];
  return found || null;
}

/** Server-side gate used by preview and submit. */
function validateTenant(tenantId) {
  var t = getTenantDetails(tenantId);
  if (!t) return { ok: false, message: 'Tenant not found in the Rental Agreement Register.' };
  if (!t.active) return { ok: false, message: t.name + ' is not an active tenant (status: ' + (t.statusText || 'blank') + ').' };
  if (t.invalid) return { ok: false, message: 'Agreement data for ' + t.name + ' is incomplete: ' + t.invalid + '.' };
  return { ok: true, tenant: t };
}

/** Current agreement as recorded in RENT HISTORY plus live parent values. */
function getAgreementDetails(tenantId) {
  var t = getTenantDetails(tenantId);
  var segs = getSegments_(tenantId, readHistory_());
  return { parent: t ? dehydrateTenant_(t) : null,
    segments: segs.map(function (s) { return { rent: rupees_(s.rentP), from: isoDate_(s.start), till: isoDate_(s.end) }; }) };
}
