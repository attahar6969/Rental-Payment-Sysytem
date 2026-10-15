/** ReceiptService.gs — receipts are rebuilt from the frozen RENT LOG row, so they never change after later revisions. */

function esc_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function getReceiptById(receiptId) {
  var id = String(receiptId || '').trim().toUpperCase();
  if (!id) return null;
  var sh = getSheet_(SHEETS.LOG), lr = sh.getLastRow();
  if (lr < 2) return null;
  var ids = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim().toUpperCase() === id) {
      var r = sh.getRange(i + 2, 1, 1, LOG_COLS.length).getValues()[0], d = {};
      try { d = JSON.parse(r[L.JSON] || '{}'); } catch (e) { }
      return {
        receiptId: String(r[L.RECEIPT]), paymentDate: toDate_(r[L.DATE]), tenantId: r[L.TID], tenantName: r[L.NAME], phone: r[L.PHONE],
        property: r[L.PROP], agreementRent: Number(r[L.RENT]), rentFrom: toDate_(r[L.RFROM]), rentTill: toDate_(r[L.RTILL]),
        amount: Number(r[L.AMT]), outstandingBefore: Number(r[L.OUTB]), outstandingAfter: Number(r[L.OUTA]),
        advance: Number(r[L.ADV]), waived: Number(r[L.WAIVE]) || 0, waiveReason: String(r[L.WREASON] || ''), paymentMode: String(r[L.PMODE] || ''), status: r[L.STATUS], generated: r[L.TS], enteredBy: r[L.BY],
        rows: d.rows || [], totals: d.totals || {}, biz: d.biz || getSystemConfig().business
      };
    }
  }
  return null;
}

/** Standalone HTML fragment with inline CSS (used for on-screen view, print and PDF). */
function generateReceipt(receiptId) {
  var r = getReceiptById(receiptId);
  if (!r) throw new Error('Receipt ' + receiptId + ' not found.');
  var color = r.status === 'PAID' ? '#1b7f4b' : r.status === 'ADVANCE' ? '#2563a8' : '#d9822b';
  var th = 'style="text-align:right;padding:6px 8px;border-bottom:1px solid #999;font-weight:600;font-size:12px"';
  var td = 'style="text-align:right;padding:6px 8px;border-bottom:1px solid #ddd;font-size:12px"';
  var body = r.rows.map(function (x) {
    return '<tr><td style="padding:6px 8px;border-bottom:1px solid #ddd;font-size:12px">' + esc_(x.month) + '</td>' +
      '<td ' + td + '>' + formatINR_(x.opening) + '</td><td ' + td + '>' + formatINR_(x.due) + '</td>' +
      (x.advance ? '<td ' + td + '>' + formatINR_(x.advance) + '</td>' : '<td ' + td + '>-</td>') +
      '<td ' + td + '><b>' + formatINR_(x.allocated) + '</b></td><td ' + td + '>' + formatINR_(x.closing) + '</td></tr>';
  }).join('');
  function line(k, v) { return '<tr><td style="padding:4px 0;color:#555;width:38%;font-size:13px">' + k + '</td><td style="padding:4px 0;font-size:13px"><b>' + v + '</b></td></tr>'; }
  var gen = r.generated instanceof Date ? fmtDate_(r.generated) + ' ' + pad2_(r.generated.getHours()) + ':' + pad2_(r.generated.getMinutes()) : esc_(r.generated);
  return '<div style="font-family:Arial,Helvetica,sans-serif;color:#1f2a44;max-width:680px;margin:0 auto;padding:22px;border:2px solid #1f2a44">' +
    '<div style="text-align:center;border-bottom:2px solid #1f2a44;padding-bottom:10px"><div style="font-size:20px;font-weight:700">' + esc_(r.biz.name) + '</div>' +
    '<div style="font-size:12px;color:#555">' + esc_(r.biz.address || '') + (r.biz.phone ? ' | ' + esc_(r.biz.phone) : '') + '</div>' +
    '<div style="font-size:16px;font-weight:700;margin-top:8px;letter-spacing:1px">RENT PAYMENT RECEIPT</div></div>' +
    '<table style="width:100%;margin-top:12px;border-collapse:collapse"><tr><td style="vertical-align:top;width:50%"><table style="width:100%">' +
    line('Receipt No.', esc_(r.receiptId)) + line('Payment Date', fmtDate_(r.paymentDate)) + line('Rent Period', fmtDate_(r.rentFrom) + ' to ' + fmtDate_(r.rentTill))+ line('Payment Mode', esc_(r.paymentMode || '-'))+
    '</table></td><td style="vertical-align:top;width:50%"><table style="width:100%">' +
    line('Received From', esc_(r.tenantName)) + line('Phone', esc_(r.phone)) + line('Property', esc_(r.property)) +
    '</table></td></tr></table>' +
    '<table style="width:100%;margin-top:6px"><tr><td style="padding:4px 0;color:#555;font-size:13px;width:38%">Agreement Rent (monthly)</td><td style="font-size:13px"><b>' + formatINR_(r.agreementRent) + '</b></td></tr></table>' +
    '<div style="margin:14px 0;padding:12px;background:#f3f5f9;text-align:center"><div style="font-size:12px;color:#555">AMOUNT RECEIVED</div>' +
    '<div style="font-size:28px;font-weight:700">' + formatINR_(r.amount, 2) + '</div>' +
    '<div style="font-size:12px;font-style:italic">' + esc_(amountInWords_(r.amount)) + '</div></div>' +
    '<table style="width:100%;border-collapse:collapse"><tr><th style="text-align:left;padding:6px 8px;border-bottom:1px solid #999;font-size:12px">Rent Month</th>' +
    '<th ' + th + '>Outstanding</th><th ' + th + '>Rent Due</th><th ' + th + '>Advance Used</th><th ' + th + '>Paid Now</th><th ' + th + '>Balance</th></tr>' + body + '</table>' +
    '<table style="width:100%;margin-top:12px">' +
    line('Previous Outstanding (before this payment)', formatINR_(r.outstandingBefore, 2)) +
    (r.waived > 0 ? line('Waived Off', formatINR_(r.waived, 2)) + line('Waiver Reason', esc_(r.waiveReason)) : '') +
    line('Remaining Outstanding', formatINR_(r.outstandingAfter, 2)) +
    (r.advance > 0 ? line('Advance Credit Carried', formatINR_(r.advance, 2)) : '') +
    '<tr><td style="padding:4px 0;color:#555;font-size:13px">Payment Status</td><td><b style="font-size:13px;color:' + color + '">' + esc_(r.status) + '</b></td></tr></table>' +
    '<div style="margin-top:22px;display:table;width:100%"><div style="display:table-cell;font-size:11px;color:#777">Generated: ' + gen + '</div>' +
    '<div style="display:table-cell;text-align:right;font-size:12px;padding-top:24px">Authorised Signatory</div></div></div>';
}

/** Opens the receipt viewer (View / Print / Export PDF). */
function showReceiptDialog(receiptId) {
  var t = HtmlService.createTemplateFromFile('Receipt');
  t.receiptId = String(receiptId);
  t.body = generateReceipt(receiptId);
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(760).setHeight(720), 'Receipt ' + receiptId);
}

/** Print is done by the browser from the viewer; this returns the printable HTML page. */
function printReceipt(receiptId) {
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc_(receiptId) + '</title></head><body onload="window.print()">' + generateReceipt(receiptId) + '</body></html>';
}

/** Creates a PDF, saves it in Drive (folder "Rent Receipts") and returns it for download. */
function exportReceiptPDF(receiptId) {
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + generateReceipt(receiptId) + '</body></html>';
  var name = String(receiptId) + '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  var url = '';
  try {
    var props = PropertiesService.getDocumentProperties(), fid = props.getProperty('RECEIPT_FOLDER_ID'), folder = null;
    if (fid) { try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = null; } }
    if (!folder) { folder = DriveApp.createFolder('Rent Receipts'); props.setProperty('RECEIPT_FOLDER_ID', folder.getId()); }
    url = folder.createFile(blob).getUrl();
  } catch (e) { logError_('exportReceiptPDF.drive', e); }
  return { name: name, base64: Utilities.base64Encode(blob.getBytes()), url: url };
}

/** Search by Receipt ID, tenant name or Tenant ID (newest first, max 50). */
function searchReceipts(query) {
  var q = String(query || '').trim().toLowerCase();
  var rows = readLog_().filter(function (p) {
    return !q || p.receipt.toLowerCase().indexOf(q) >= 0 || String(p.name).toLowerCase().indexOf(q) >= 0 || p.tid.toLowerCase() === q;
  });
  rows.sort(function (a, b) { return b.row - a.row; });
  return rows.slice(0, 50).map(function (p) {
    return { receiptId: p.receipt, date: fmtDate_(p.date), tenant: p.name, amount: rupees_(p.amountP), status: p.status };
  });
}
