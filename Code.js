// ================================================================
// DAR PROCUREMENT MONITORING SYSTEM — LOG-BASED WORKFLOW (v9)
// ================================================================

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

var SHEETS = {
  TRANSACTIONS: 'TRANSACTIONS',
  HISTORY: 'TRANSACTION_HISTORY',
  USERS: 'USERS',
  SUPPLIERS: 'SUPPLIERS',
  END_USERS: 'END_USERS',
  AUDIT: 'AUDIT_LOGS'
};

var END_USER_ROLES = ['HR','GSS','LTID','LEGAL','FINANCE','IT','ADMIN DIVISION','DARPO'];

var TRANSACTIONS_HEADERS = [
  'TRACKING NO.','P.R NO.','END-USER / REQUESTED BY','CATEGORY','PHIGEPS POSTED','APPROVED ABC','RFQ','COMPLETE RFQ RETURNED',
  'RFQ DATE OF COMPLETION','ABSTRACT OF AWARD','AWARDED TO','DATE OF AWARD','MODE OF PROCUREMENT','BAC RESO','BAC RESO STATUS','NOA','NTP',
  'P.O. NO.','P.O. DATE','DOCUMENT TYPE','SUPPLIER','PARTICULARS','P.O. AMOUNT','DATE TRANSMITTED TO SUPPLIER',
  'DATE RECEIVED BY SUPPLIER','COA DATE RECEIVED','I.A.R  NO.','INVOICE  NO.','ORS NO.','BUPR NO.','ORS AMOUNT','DV NO.',
  'NET AMOUNT','TAX AMOUNT','CHEQUE NO. / LDDAP','DATE OF CHEQUE','RCAO','ARDA','CURRENT DEPT','CURRENT STATUS'
];
var HISTORY_HEADERS = ['TRACKING NO.','FROM DEPT','TO DEPT','ACTION','STATUS','REMARKS','USER','TIMESTAMP'];

var TX_FIELD_ALIASES = {
  'P.R #': ['P.R #','P.R NO.'],
  'P.R NO.': ['P.R NO.','P.R #'],
  'P.O. No.': ['P.O. No.','P.O. NO.'],
  'P.O. NO.': ['P.O. NO.','P.O. No.'],
  'AMOUNT': ['AMOUNT','P.O. AMOUNT','ORS AMOUNT'],
  'P.O. AMOUNT': ['P.O. AMOUNT','AMOUNT'],
  'I.A.R #': ['I.A.R #','I.A.R  NO.'],
  'I.A.R  NO.': ['I.A.R  NO.','I.A.R #'],
  'INVOICE #': ['INVOICE #','INVOICE  NO.'],
  'INVOICE  NO.': ['INVOICE  NO.','INVOICE #'],
  'DV No.': ['DV No.','DV NO.'],
  'DV NO.': ['DV NO.','DV No.'],
  'Net Amount': ['Net Amount','NET AMOUNT'],
  'NET AMOUNT': ['NET AMOUNT','Net Amount'],
  'TAX Amount': ['TAX Amount','TAX AMOUNT'],
  'TAX AMOUNT': ['TAX AMOUNT','TAX Amount']
};

var DEPT_ALIASES = {
  'BAC': 'BAC',
  'SUPPLY': 'SUPPLY',
  'SUPPLY AND PROPERTY': 'SUPPLY',
  'BUDGET': 'BUDGET',
  'ACCOUNTING': 'ACCOUNTING',
  'CASH': 'CASH',
  'CASHIER': 'CASH',
  'RCAO': 'RCAO',
  'ARDA': 'ARDA',
  'END USER': 'END USER',
  'REQUESTING OFFICE': 'END USER',
  'COMPLETED': 'COMPLETED'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('DAR RO V — Procurement Monitoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
}

function formatCellForClient(header, value) {
  if (value === null || value === undefined) return '';
  if (!(value instanceof Date)) return String(value);
  var key = String(header || '').trim().toUpperCase();
  var withTime = key.indexOf('DATE') !== -1 || key === 'TIMESTAMP';
  var pattern = withTime ? 'MMM dd, yyyy hh:mm a' : 'MMM dd, yyyy';
  return Utilities.formatDate(value, Session.getScriptTimeZone(), pattern);
}

function zeroPad(num, size) {
  var s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

function normalizeDepartmentName(name) {
  var key = String(name || '').trim().toUpperCase();
  return DEPT_ALIASES[key] || key;
}

function displayDepartmentName(name) {
  var n = normalizeDepartmentName(name);
  if (n === 'SUPPLY') return 'Supply';
  if (n === 'CASH') return 'Cashier';
  if (n === 'END USER') return 'Requesting Office';
  return n;
}

function normalizeSupplierStatus(status) {
  var v = String(status || '').trim().toUpperCase();
  if (v === 'INACTIVE' || v === 'BLOCKED') return v;
  return 'ACTIVE';
}

function normalizeSupplierName(name) {
  return String(name || '').trim().replace(/\s+/g, ' ');
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getHeaderMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || '').trim().toUpperCase()] = i;
  }
  return map;
}

function ensureSheetHeaders(sheet, expectedHeaders) {
  var lc = sheet.getLastColumn();
  if (sheet.getLastRow() === 0 || lc === 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return;
  }

  var current = sheet.getRange(1, 1, 1, lc).getValues()[0].map(function(h){ return String(h || '').trim(); });
  var currentUpper = current.map(function(h){ return h.toUpperCase(); });

  for (var i = 0; i < expectedHeaders.length; i++) {
    if (currentUpper.indexOf(expectedHeaders[i].toUpperCase()) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(expectedHeaders[i]);
      current.push(expectedHeaders[i]);
      currentUpper.push(expectedHeaders[i].toUpperCase());
    }
  }
}

function ensureCoreSheets(ss) {
  var tx = getOrCreateSheet(ss, SHEETS.TRANSACTIONS);
  ensureSheetHeaders(tx, TRANSACTIONS_HEADERS);

  var hist = getOrCreateSheet(ss, SHEETS.HISTORY);
  ensureSheetHeaders(hist, HISTORY_HEADERS);

  var users = getOrCreateSheet(ss, SHEETS.USERS);
  ensureSheetHeaders(users, ['FIRST NAME','LAST NAME','USERNAME','PASSWORD','ROLE','END USER','STATUS']);

  var suppliers = getOrCreateSheet(ss, SHEETS.SUPPLIERS);
  ensureSheetHeaders(suppliers, ['SUPPLIER NAME','STATUS']);

  var eus = getOrCreateSheet(ss, SHEETS.END_USERS);
  ensureSheetHeaders(eus, ['END USER ROLE','STATUS']);

  var audit = getOrCreateSheet(ss, SHEETS.AUDIT);
  ensureSheetHeaders(audit, ['TIMESTAMP','USER','ACTION','DEPARTMENT','RECORD ID']);
}

function getTransactionsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.TRANSACTIONS);
  ensureSheetHeaders(sheet, TRANSACTIONS_HEADERS);
  return sheet;
}

function getHistorySheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.HISTORY);
  ensureSheetHeaders(sheet, HISTORY_HEADERS);
  return sheet;
}

function getTransactionsRows(ss) {
  var sheet = getTransactionsSheet(ss);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(hdrs);
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var trackingNo = String(row[hm['TRACKING NO.']] || '').trim();
    if (!trackingNo) continue;
    out.push({
      _rowNum: i + 1,
      trackingNo: trackingNo,
      prNo: hm['P.R NO.'] !== undefined ? String(row[hm['P.R NO.']] || '').trim() : '',
      requestingOffice: hm['END-USER / REQUESTED BY'] !== undefined ? String(row[hm['END-USER / REQUESTED BY']] || '').trim() : '',
      category: hm['CATEGORY'] !== undefined ? String(row[hm['CATEGORY']] || '').trim() : '',
      currentDept: normalizeDepartmentName(row[hm['CURRENT DEPT']]),
      currentStatus: String(row[hm['CURRENT STATUS']] || '').trim() || 'PROCESSING'
    });
  }
  return out;
}

function getTxCandidateKeys(key) {
  var k = String(key || '').trim();
  return TX_FIELD_ALIASES[k] || [k];
}

function getTxValueByCandidates(record, candidates) {
  if (!record) return '';
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    if (record.hasOwnProperty(c) && String(record[c] || '').trim() !== '') return record[c];
  }
  return '';
}

function setTransactionFieldsByTracking(ss, trackingNo, data) {
  if (!data) return;
  var sheet = getTransactionsSheet(ss);
  var all = sheet.getDataRange().getValues();
  if (all.length < 2) return;
  var hdrs = all[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(hdrs);
  var targetRow = 0;
  var tnIdx = hm['TRACKING NO.'];
  if (tnIdx === undefined) return;

  for (var i = 1; i < all.length; i++) {
    if (String(all[i][tnIdx] || '').trim() === String(trackingNo || '').trim()) {
      targetRow = i + 1;
      break;
    }
  }
  if (!targetRow) return;

  var protectedKeys = {'TRACKING NO.':1,'CURRENT DEPT':1,'CURRENT STATUS':1};
  for (var key in data) {
    if (!data.hasOwnProperty(key)) continue;
    var normalizedKey = String(key || '').trim();
    if (!normalizedKey || protectedKeys[normalizedKey]) continue;
    var candidates = getTxCandidateKeys(normalizedKey);
    var col = -1;
    for (var c = 0; c < candidates.length; c++) {
      var idx = hm[String(candidates[c]).trim().toUpperCase()];
      if (idx !== undefined) {
        col = idx + 1;
        break;
      }
    }
    if (col === -1) continue;

    if (normalizedKey === 'SUPPLIER' || normalizedKey === 'AWARDED TO') {
      ensureSupplierExists(ss, data[key]);
    }

    sheet.getRange(targetRow, col).setValue(data[key]);
  }
}

function getHistoryRows(ss) {
  var sheet = getHistorySheet(ss);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(hdrs);

  function pick(row, key, fallback) {
    var idx = hm[key];
    if (idx === undefined && fallback) idx = hm[fallback];
    return idx === undefined ? '' : row[idx];
  }

  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var trackingNo = String(pick(row, 'TRACKING NO.')).trim();
    if (!trackingNo) continue;
    var isNotifiedVal = pick(row, 'IS NOTIFIED');
    var isNotified = isNotifiedVal === '' ? undefined : (String(isNotifiedVal || '').toUpperCase() === 'TRUE');
    out.push({
      _rowNum: i + 1,
      trackingNo: trackingNo,
      fromDept: normalizeDepartmentName(pick(row, 'FROM DEPT')),
      toDept: normalizeDepartmentName(pick(row, 'TO DEPT')),
      action: String(pick(row, 'ACTION')).trim(),
      status: String(pick(row, 'STATUS')).trim(),
      remarks: String(pick(row, 'REMARKS')).trim(),
      processedBy: String(pick(row, 'USER', 'PROCESSED BY')).trim(),
      timestamp: formatCellForClient('TIMESTAMP', pick(row, 'TIMESTAMP')),
      isNotified: isNotified
    });
  }
  return out;
}

function appendHistoryLog(ss, payload) {
  var sheet = getHistorySheet(ss);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(headers);

  var row = new Array(headers.length);
  for (var i = 0; i < row.length; i++) row[i] = '';

  function set(col, val) {
    var idx = hm[col];
    if (idx !== undefined) row[idx] = val;
  }

  set('TRACKING NO.', payload.trackingNo || '');
  set('FROM DEPT', normalizeDepartmentName(payload.fromDept));
  set('TO DEPT', normalizeDepartmentName(payload.toDept));
  set('ACTION', payload.action || '');
  set('STATUS', payload.status || '');
  set('REMARKS', payload.remarks || '');
  if (hm['PROCESSED BY'] !== undefined) set('PROCESSED BY', payload.processedBy || 'SYSTEM');
  if (hm['USER'] !== undefined) set('USER', payload.processedBy || 'SYSTEM');
  set('TIMESTAMP', getTimestamp());
  if (hm['IS NOTIFIED'] !== undefined) set('IS NOTIFIED', 'FALSE');

  sheet.appendRow(row);
}

function upsertTransactionSummary(ss, trackingNo, currentDept, currentStatus) {
  var txSheet = getTransactionsSheet(ss);
  var data = txSheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(hdrs);

  var tnCol = hm['TRACKING NO.'] + 1;
  var deptCol = hm['CURRENT DEPT'] + 1;
  var statusCol = hm['CURRENT STATUS'] + 1;
  var wanted = String(trackingNo || '').trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][tnCol - 1] || '').trim() === wanted) {
      txSheet.getRange(i + 1, deptCol).setValue(normalizeDepartmentName(currentDept));
      txSheet.getRange(i + 1, statusCol).setValue(String(currentStatus || '').trim() || 'PROCESSING');
      return i + 1;
    }
  }

  var newRow = new Array(hdrs.length);
  for (var j = 0; j < newRow.length; j++) newRow[j] = '';
  newRow[tnCol - 1] = wanted;
  newRow[deptCol - 1] = normalizeDepartmentName(currentDept);
  newRow[statusCol - 1] = String(currentStatus || '').trim() || 'PROCESSING';
  txSheet.appendRow(newRow);
  return txSheet.getLastRow();
}

function getTrackingNoByRowNum(ss, rowNum) {
  var tx = getTransactionsSheet(ss);
  var hdrs = tx.getRange(1,1,1,tx.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
  var hm = getHeaderMap(hdrs);
  var tnIdx = hm['TRACKING NO.'];
  if (tnIdx === undefined) return '';
  if (rowNum < 2 || rowNum > tx.getLastRow()) return '';
  return String(tx.getRange(rowNum, tnIdx + 1).getValue() || '').trim();
}

function generateTrackingNo(ss) {
  var now = new Date();
  var yyyy = now.getFullYear();
  var mm = zeroPad(now.getMonth() + 1, 2);

  var rows = getTransactionsRows(ss);
  var prefix = 'DAR-' + yyyy + '-' + mm + '-';
  var maxNum = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].trackingNo.indexOf(prefix) !== 0) continue;
    var n = parseInt(rows[i].trackingNo.substring(prefix.length), 10) || 0;
    if (n > maxNum) maxNum = n;
  }
  return prefix + zeroPad(maxNum + 1, 4);
}

function logAudit(username, action, department, recordId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);
    var sheet = ss.getSheetByName(SHEETS.AUDIT);
    sheet.appendRow([getTimestamp(), username || 'SYSTEM', action || '', department || '', recordId || '']);
  } catch (e) {
    // Keep audit failures non-blocking.
  }
}

function logTransactionHistory(trackingNo, prNo, action, fromDept, toDept, username, remarks) {
  var _ = prNo;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);
  appendHistoryLog(ss, {
    trackingNo: trackingNo,
    fromDept: fromDept,
    toDept: toDept,
    action: action,
    status: action === 'Completed' ? 'COMPLETED' : '',
    remarks: remarks || '',
    processedBy: username || 'SYSTEM'
  });
}

function getTransactionHistory(trackingNo) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);
  var rows = getHistoryRows(ss).filter(function(r){ return r.trackingNo === trackingNo; });
  return rows.map(function(r){
    return {
      'TRACKING NO.': r.trackingNo,
      'USER': r.processedBy,
      'FROM': displayDepartmentName(r.fromDept),
      'TO': displayDepartmentName(r.toDept),
      'ACTION': r.action,
      'STATUS': r.status,
      'REMARKS': r.remarks,
      'TIMESTAMP': r.timestamp
    };
  });
}

function parseMetaFromCreateLog(remarks) {
  var text = String(remarks || '').trim();
  if (text.indexOf('META:') !== 0) return {};
  try {
    return JSON.parse(text.substring(5));
  } catch (e) {
    return {};
  }
}

function getTrackingMetaMap(ss) {
  var out = {};
  var rows = getHistoryRows(ss);
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r.action || '').toUpperCase() !== 'CREATED') continue;
    if (out[r.trackingNo]) continue;
    out[r.trackingNo] = parseMetaFromCreateLog(r.remarks);
  }
  return out;
}

function buildDepartmentRows(deptName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);

  var dept = normalizeDepartmentName(deptName);
  var txSheet = getTransactionsSheet(ss);
  var txData = txSheet.getDataRange().getValues();
  if (txData.length < 2) return [];
  var txHeaders = txData[0].map(function(h){ return String(h || '').trim(); });
  var txMap = getHeaderMap(txHeaders);

  var txRows = [];
  for (var r = 1; r < txData.length; r++) {
    var rowVals = txData[r];
    var tracking = String(rowVals[txMap['TRACKING NO.']] || '').trim();
    if (!tracking) continue;
    var record = {};
    for (var h = 0; h < txHeaders.length; h++) {
      if (!txHeaders[h]) continue;
      record[txHeaders[h]] = formatCellForClient(txHeaders[h], rowVals[h]);
    }
    txRows.push({
      _rowNum: r + 1,
      trackingNo: tracking,
      currentDept: normalizeDepartmentName(rowVals[txMap['CURRENT DEPT']]),
      currentStatus: String(rowVals[txMap['CURRENT STATUS']] || '').trim() || 'PROCESSING',
      record: record
    });
  }

  var histRows = getHistoryRows(ss);
  var metaMap = getTrackingMetaMap(ss);

  var byTracking = {};
  for (var i = 0; i < histRows.length; i++) {
    var h = histRows[i];
    byTracking[h.trackingNo] = byTracking[h.trackingNo] || [];
    byTracking[h.trackingNo].push(h);
  }

  var out = [];
  for (var j = 0; j < txRows.length; j++) {
    var t = txRows[j];
    if (dept !== 'ALL' && t.currentDept !== dept) continue;

    var txRecord = t.record || {};
    var meta = metaMap[t.trackingNo] || {};
    var history = byTracking[t.trackingNo] || [];
    var latest = history.length ? history[history.length - 1] : null;

    function pickTx(key, fallback) {
      var candidates = getTxCandidateKeys(key).slice();
      if (fallback) candidates.push(fallback);
      var val = getTxValueByCandidates(txRecord, candidates);
      return val === '' ? '' : val;
    }

    var prNo = pickTx('P.R NO.') || meta.prNo || '';
    var office = pickTx('END-USER / REQUESTED BY') || meta.requestingOffice || '';
    var category = pickTx('CATEGORY') || meta.category || '';

    var receivedAt = '';
    var completedAt = '';
    var cancelledAt = '';
    var forwardedFromDept = '';
    var returnedFromDept = '';
    var returnedFromDate = '';
    var returnedFromRemarks = '';
    var statusUpper = String(t.currentStatus || '').trim().toUpperCase();
    var allowNotifiedFromHistory = statusUpper !== 'PROCESSING' && statusUpper !== 'RETURN TO PROCESSING' && statusUpper !== 'UNDO';
    var displayStatus = statusUpper === 'UNDO' ? 'PROCESSING' : (t.currentStatus || 'PROCESSING');
    var isNewNotified = false;
    var latestForwardIdx = -1;
    for (var z = history.length - 1; z >= 0; z--) {
      var hz = history[z];
      if (!completedAt && String(hz.action || '').toUpperCase() === 'COMPLETED') completedAt = hz.timestamp;
      if (!cancelledAt && String(hz.action || '').toUpperCase() === 'CANCELLED') cancelledAt = hz.timestamp;
      if (!receivedAt && normalizeDepartmentName(hz.toDept) === t.currentDept && String(hz.action || '').toUpperCase() === 'RECEIVED') {
        receivedAt = hz.timestamp;
      }
      if (!returnedFromDept && String(hz.action || '').toUpperCase() === 'RETURNED' && normalizeDepartmentName(hz.toDept) === t.currentDept) {
        returnedFromDept = displayDepartmentName(normalizeDepartmentName(hz.fromDept));
        returnedFromDate = hz.timestamp || '';
        returnedFromRemarks = hz.remarks || '';
      }
      if (!isNewNotified && allowNotifiedFromHistory && String(hz.action || '').toUpperCase() === 'FORWARD') {
        isNewNotified = hz.isNotified === false || String(hz.isNotified || '').toLowerCase() === 'false';
        latestForwardIdx = z;
        if (!forwardedFromDept && normalizeDepartmentName(hz.toDept) === t.currentDept) {
          forwardedFromDept = displayDepartmentName(normalizeDepartmentName(hz.fromDept));
        }
        break;
      }
      if (completedAt && cancelledAt && receivedAt) break;
    }
    // A transaction is "returning" to this dept if it was previously processed here
    // (i.e., this dept appears as fromDept in any history entry before the latest forward)
    var isReturnBack = false;
    var returnBackFromDept = ''; // the dept that forwarded it back (e.g. END USER)
    if (isNewNotified && latestForwardIdx > 0) {
      for (var k = 0; k < latestForwardIdx; k++) {
        if (normalizeDepartmentName(history[k].fromDept) === t.currentDept) {
          isReturnBack = true;
        }
        // Track the most recent Returned event where this dept returned it to someone —
        // that recipient is who eventually forwarded it back
        if (isReturnBack && String(history[k].action || '').toUpperCase() === 'RETURNED' &&
            normalizeDepartmentName(history[k].fromDept) === t.currentDept) {
          returnBackFromDept = history[k].toDept;
        }
      }
    }

    out.push({
      _rowNum: t._rowNum,
      'TRACKING NO.': t.trackingNo,
      'P.R #': prNo,
      'P.R NO.': prNo,
      'END-USER / REQUESTED BY': office,
      'END USER': office,
      'CATEGORY': category,
      'PHIGEPS POSTED': pickTx('PHIGEPS POSTED'),
      'APPROVED ABC': pickTx('APPROVED ABC'),
      'RFQ': pickTx('RFQ'),
      'COMPLETE RFQ RETURNED': pickTx('COMPLETE RFQ RETURNED'),
      'RFQ DATE OF COMPLETION': pickTx('RFQ DATE OF COMPLETION'),
      'ABSTRACT OF AWARD': pickTx('ABSTRACT OF AWARD'),
      'AWARDED TO': pickTx('AWARDED TO'),
      'DATE OF AWARD': pickTx('DATE OF AWARD'),
      'MODE OF PROCUREMENT': pickTx('MODE OF PROCUREMENT'),
      'BAC RESO': pickTx('BAC RESO'),
      'BAC RESO STATUS': pickTx('BAC RESO STATUS'),
      'NOA': pickTx('NOA'),
      'NTP': pickTx('NTP'),
      'P.O. No.': pickTx('P.O. NO.'),
      'P.O. NO.': pickTx('P.O. NO.'),
      'P.O. DATE': pickTx('P.O. DATE'),
      'DOCUMENT TYPE': pickTx('DOCUMENT TYPE'),
      'SUPPLIER': pickTx('SUPPLIER'),
      'PARTICULARS': pickTx('PARTICULARS'),
      'AMOUNT': pickTx('P.O. AMOUNT') || pickTx('ORS AMOUNT'),
      'P.O. AMOUNT': pickTx('P.O. AMOUNT'),
      'DATE TRANSMITTED TO SUPPLIER': pickTx('DATE TRANSMITTED TO SUPPLIER'),
      'DATE RECEIVED BY SUPPLIER': pickTx('DATE RECEIVED BY SUPPLIER'),
      'COA DATE RECEIVED': pickTx('COA DATE RECEIVED'),
      'I.A.R #': pickTx('I.A.R  NO.'),
      'I.A.R  NO.': pickTx('I.A.R  NO.'),
      'INVOICE #': pickTx('INVOICE  NO.'),
      'INVOICE  NO.': pickTx('INVOICE  NO.'),
      'ORS NO.': pickTx('ORS NO.'),
      'BUPR NO.': pickTx('BUPR NO.'),
      'ORS AMOUNT': pickTx('ORS AMOUNT'),
      'DV No.': pickTx('DV NO.'),
      'DV NO.': pickTx('DV NO.'),
      'Net Amount': pickTx('NET AMOUNT'),
      'NET AMOUNT': pickTx('NET AMOUNT'),
      'TAX Amount': pickTx('TAX AMOUNT'),
      'TAX AMOUNT': pickTx('TAX AMOUNT'),
      'CHEQUE NO. / LDDAP': pickTx('CHEQUE NO. / LDDAP'),
      'DATE OF CHEQUE': pickTx('DATE OF CHEQUE'),
      'RCAO': pickTx('RCAO'),
      'ARDA': pickTx('ARDA'),
      'STATUS': displayStatus,
      'CURRENT DEPT': t.currentDept,
      'FORWARDED FROM': forwardedFromDept,
      'FORWARD REMARKS': latest ? latest.remarks : '',
      'COMPLETED AT': completedAt,
      'CANCELLED AT': cancelledAt,
      'RETURN TO': '',
      'RETURNED FROM': isReturnBack ? displayDepartmentName(returnBackFromDept) : returnedFromDept,
      'RETURN REMARKS': returnedFromRemarks,
      'RETURNED DATE': isReturnBack && latest ? latest.timestamp : returnedFromDate,
      'RETURN RECEIVED DATE': '',
      'RETURN RECEIVED REMARKS': '',
      'IS NOTIFIED': isNewNotified ? 'TRUE' : '',
      'IS NEW': isNewNotified ? 'TRUE' : '',
      'IS RETURN BACK': isReturnBack ? 'TRUE' : '',
      'DATE RECEIVED': receivedAt || (latest ? latest.timestamp : '')
    });
  }

  return out;
}

function buildDepartmentRowsMapFromAllRows(allRows) {
  var map = { bac:[], supply:[], budget:[], accounting:[], cash:[], rcao:[], arda:[], enduser:[] };
  (allRows || []).forEach(function(r) {
    var dept = normalizeDepartmentName(r['CURRENT DEPT']);
    if (dept === 'BAC') map.bac.push(r);
    else if (dept === 'SUPPLY') map.supply.push(r);
    else if (dept === 'BUDGET') map.budget.push(r);
    else if (dept === 'ACCOUNTING') map.accounting.push(r);
    else if (dept === 'CASH') map.cash.push(r);
    else if (dept === 'RCAO') map.rcao.push(r);
    else if (dept === 'ARDA') map.arda.push(r);
    else if (dept === 'END USER') map.enduser.push(r);
  });
  return map;
}

function getDepartmentRowsMap() {
  return buildDepartmentRowsMapFromAllRows(buildDepartmentRows('ALL'));
}

// Returns rows for a single dept including forwarded-away completed copies.
function getDeptPageData(deptName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);
  var dept = normalizeDepartmentName(deptName);
  var allRows = buildDepartmentRows('ALL');
  var histRows = getHistoryRows(ss);

  var histByTracking = {};
  for (var i = 0; i < histRows.length; i++) {
    var hh = histRows[i];
    histByTracking[hh.trackingNo] = histByTracking[hh.trackingNo] || [];
    histByTracking[hh.trackingNo].push(hh);
  }

  var rows = allRows.filter(function(row) { return normalizeDepartmentName(row['CURRENT DEPT']) === dept; });

  for (var j = 0; j < allRows.length; j++) {
    var row = allRows[j];
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (!trackingNo) continue;
    if (normalizeDepartmentName(row['CURRENT DEPT']) === dept) continue;

    var history = histByTracking[trackingNo] || [];
    var completedAt = '';
    for (var k = history.length - 1; k >= 0; k--) {
      var entry = history[k];
      var actionUpper = String(entry.action || '').toUpperCase();
      if (normalizeDepartmentName(entry.fromDept) !== dept) continue;
      if (actionUpper === 'FORWARD' || actionUpper === 'COMPLETED') {
        completedAt = entry.timestamp || '';
        break;
      }
    }
    if (!completedAt) continue;

    var copy = Object.assign({}, row);
    copy['COMPLETED AT'] = completedAt;
    copy['IS NOTIFIED'] = '';
    copy['IS NEW'] = '';
    copy['RETURNED DATE'] = '';
    copy['RETURN RECEIVED DATE'] = '';
    copy['RETURN TO'] = '';
    rows.push(copy);
  }

  return rows;
}

// Returns page-data for all depts in a single pass (efficient for dashboard bundle).
function getDeptPageDataAll() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);
  var allRows = buildDepartmentRows('ALL');
  var histRows = getHistoryRows(ss);

  var histByTracking = {};
  for (var i = 0; i < histRows.length; i++) {
    var hh = histRows[i];
    histByTracking[hh.trackingNo] = histByTracking[hh.trackingNo] || [];
    histByTracking[hh.trackingNo].push(hh);
  }

  var DEPT_KEYS = { 'BAC': 'bac', 'SUPPLY': 'supply', 'BUDGET': 'budget', 'ACCOUNTING': 'accounting', 'CASH': 'cash', 'RCAO': 'rcao', 'ARDA': 'arda', 'END USER': 'enduser' };
  var map = { bac:[], supply:[], budget:[], accounting:[], cash:[], rcao:[], arda:[], enduser:[] };

  // Step 1: bucket current rows by dept
  (allRows || []).forEach(function(row) {
    var key = DEPT_KEYS[normalizeDepartmentName(row['CURRENT DEPT'])];
    if (key) map[key].push(row);
  });

  // Step 2: add forwarded-away completed copies for each past dept
  (allRows || []).forEach(function(row) {
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (!trackingNo) return;
    var currentKey = DEPT_KEYS[normalizeDepartmentName(row['CURRENT DEPT'])];
    var history = histByTracking[trackingNo] || [];

    // Find last completion-causing entry per source dept.
    // FORWARD keeps the old behavior, RETURNED from enduser keeps enduser behavior,
    // and COMPLETED supports cashier-paid rows appearing in cashier done tab.
    var lastCompletionEntry = {};
    for (var k = 0; k < history.length; k++) {
      var hz = history[k];
      var hzAction = String(hz.action || '').toUpperCase();
      var fromKey = DEPT_KEYS[normalizeDepartmentName(hz.fromDept)];
      // Include FORWARD from any dept, and RETURNED from END USER (end users use "Returned" to send transactions on)
      var isEnduserReturn = hzAction === 'RETURNED' && fromKey === 'enduser';
      var isCompleted = hzAction === 'COMPLETED';
      if (hzAction !== 'FORWARD' && !isEnduserReturn && !isCompleted) continue;
      // Allow enduser entries unless the transaction is still at END USER (to avoid duplicates with Step 1)
      if (!fromKey || (fromKey === 'enduser' && currentKey === 'enduser')) continue;
      lastCompletionEntry[fromKey] = hz;
    }

    var fromKeys = Object.keys(lastCompletionEntry);
    for (var fi = 0; fi < fromKeys.length; fi++) {
      var fromKey = fromKeys[fi];
      if (fromKey === currentKey) continue;
      var hz2 = lastCompletionEntry[fromKey];
      var copy = Object.assign({}, row);
      copy['COMPLETED AT'] = hz2.timestamp || copy['COMPLETED AT'] || '';
      copy['IS NOTIFIED'] = '';
      copy['IS NEW'] = '';
      copy['RETURNED DATE'] = '';
      copy['RETURN RECEIVED DATE'] = '';
      copy['RETURN TO'] = '';
      if (fromKey === 'enduser') {
        copy['COMPLETED TO'] = displayDepartmentName(hz2.toDept || '') || '—';
        copy['FORWARD REMARKS'] = hz2.remarks || '';
      }
      map[fromKey].push(copy);
    }
  });

  return map;
}

function loginUser(username, password) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) return {success:false, message:'USERS sheet not found.'};

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return {success:false, message:'No users found.'};

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var uCol = hm['USERNAME'];
    var pCol = hm['PASSWORD'];
    var rCol = hm['ROLE'];
    var euCol = hm['END USER'];
    var stCol = hm['STATUS'];

    if (uCol === undefined || pCol === undefined || rCol === undefined) {
      return {success:false, message:'USERS sheet missing required columns.'};
    }

    for (var i = 1; i < data.length; i++) {
      var uname = String(data[i][uCol] || '').trim();
      var pwd = String(data[i][pCol] || '').trim();
      var role = String(data[i][rCol] || '').trim().toUpperCase();
      var isEU = euCol !== undefined ? String(data[i][euCol] || '').trim().toUpperCase() === 'TRUE' : false;
      var status = stCol !== undefined ? String(data[i][stCol] || '').trim().toUpperCase() : 'ACTIVE';

      if (uname.toLowerCase() !== String(username || '').trim().toLowerCase()) continue;
      if (pwd !== String(password || '').trim()) continue;
      if (status === 'INACTIVE' || status === 'FALSE' || status === 'BLOCKED') {
        return {success:false, message:'Your account is inactive.'};
      }

      logAudit(uname, 'LOGIN', role || '-', '-');
      return {success:true, username:uname, role:role, isEndUser:isEU};
    }

    return {success:false, message:'Invalid username or password.'};
  } catch (e) {
    return {success:false, message:'Login error: ' + e.message};
  }
}

function getBACData()        { return buildDepartmentRows('BAC'); }
function getBACPageData()    {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);

  var allRows = buildDepartmentRows('ALL');
  var rows = allRows.filter(function(row){ return normalizeDepartmentName(row['CURRENT DEPT']) === 'BAC'; });
  var histRows = getHistoryRows(ss);

  var histByTracking = {};
  for (var i = 0; i < histRows.length; i++) {
    var hist = histRows[i];
    histByTracking[hist.trackingNo] = histByTracking[hist.trackingNo] || [];
    histByTracking[hist.trackingNo].push(hist);
  }

  for (var j = 0; j < allRows.length; j++) {
    var row = allRows[j];
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (!trackingNo) continue;
    if (normalizeDepartmentName(row['CURRENT DEPT']) === 'BAC') continue;

    var history = histByTracking[trackingNo] || [];
    var forwardedAt = '';
    var forwardedFromBac = false;
    for (var k = history.length - 1; k >= 0; k--) {
      var entry = history[k];
      if (String(entry.action || '').toUpperCase() !== 'FORWARD') continue;
      if (normalizeDepartmentName(entry.fromDept) !== 'BAC') continue;
      forwardedFromBac = true;
      forwardedAt = entry.timestamp || '';
      break;
    }
    if (!forwardedFromBac) continue;

    var copy = Object.assign({}, row);
    copy['COMPLETED AT'] = forwardedAt || copy['COMPLETED AT'] || '';
    copy['SOURCE DEPT'] = 'BAC';
    rows.push(copy);
  }

  return rows;
}
function getSupplyData()     { return buildDepartmentRows('SUPPLY'); }
function getBudgetData()     { return buildDepartmentRows('BUDGET'); }
function getAccountingData() { return buildDepartmentRows('ACCOUNTING'); }
function getCashData()       { return buildDepartmentRows('CASH'); }
function getRCAOData()       { return buildDepartmentRows('RCAO'); }
function getARDAData()       { return buildDepartmentRows('ARDA'); }

// Page-data variants include forwarded-away completed copies (for done tab).
function getSupplyPageData()     { return getDeptPageData('SUPPLY'); }
function getBudgetPageData()     { return getDeptPageData('BUDGET'); }
function getAccountingPageData() { return getDeptPageData('ACCOUNTING'); }
function getCashPageData()       { return getDeptPageData('CASH'); }
function getRCAOPageData()       { return getDeptPageData('RCAO'); }
function getARDAPageData()       { return getDeptPageData('ARDA'); }
function getEndUserData()    { return []; }

function getDepartmentPageCounts() {
  var map = getDepartmentRowsMap();
  return {
    bac: map.bac.length,
    supply: map.supply.length,
    budget: map.budget.length,
    accounting: map.accounting.length,
    cash: map.cash.length,
    rcao: map.rcao.length,
    arda: map.arda.length
  };
}

function getMyBACData(role) {
  var r = String(role || '').trim().toUpperCase();
  return getBACData().filter(function(x){
    return String(x['END-USER / REQUESTED BY'] || '').trim().toUpperCase() === r;
  });
}

function getRequestingOffices() {
  var list = getEndUsers();
  return (list || []).slice().sort();
}

function getSuppliers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);
    var s = ss.getSheetByName(SHEETS.SUPPLIERS);
    var data = s.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);
    var nameCol = hm['SUPPLIER NAME'];
    var statusCol = hm['STATUS'];
    var out = [];

    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][nameCol] || '').trim();
      if (!name) continue;
      var st = statusCol !== undefined ? normalizeSupplierStatus(data[i][statusCol]) : 'ACTIVE';
      if (st !== 'ACTIVE') continue;
      out.push(name);
    }

    return out.sort();
  } catch (e) {
    return [];
  }
}

function ensureSupplierExists(ss, supplierName) {
  var name = normalizeSupplierName(supplierName);
  if (!name) return {success:true, added:false, updated:false, supplierName:''};

  var sheet = getOrCreateSheet(ss, SHEETS.SUPPLIERS);
  ensureSheetHeaders(sheet, ['SUPPLIER NAME','STATUS']);

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    sheet.appendRow([name, 'ACTIVE']);
    return {success:true, added:true, updated:false, supplierName:name};
  }

  var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
  var hm = getHeaderMap(hdrs);
  var nameCol = hm['SUPPLIER NAME'];
  var statusCol = hm['STATUS'];
  if (nameCol === undefined) return {success:false, message:'SUPPLIERS sheet missing required columns.'};

  var wanted = name.toUpperCase();
  for (var i = 1; i < data.length; i++) {
    var existing = normalizeSupplierName(data[i][nameCol]);
    if (!existing || existing.toUpperCase() !== wanted) continue;
    if (statusCol !== undefined && normalizeSupplierStatus(data[i][statusCol]) !== 'ACTIVE') {
      sheet.getRange(i + 1, statusCol + 1).setValue('ACTIVE');
    }
    return {success:true, added:false, updated:true, supplierName:existing};
  }

  var row = new Array(hdrs.length);
  for (var j = 0; j < row.length; j++) row[j] = '';
  row[nameCol] = name;
  if (statusCol !== undefined) row[statusCol] = 'ACTIVE';
  sheet.appendRow(row);
  return {success:true, added:true, updated:false, supplierName:name};
}

function getDashboardStats() {
  var tx = getTransactionsRows(SpreadsheetApp.openById(SPREADSHEET_ID));
  var total = tx.length;
  var completed = tx.filter(function(r){ return r.currentStatus === 'COMPLETED'; }).length;
  var cancelled = tx.filter(function(r){ return r.currentStatus === 'CANCELLED'; }).length;
  var active = Math.max(0, total - completed - cancelled);
  return { total: total, processing: active, completed: completed, cancelled: cancelled };
}

function getMyDashboardStats(role) {
  var _ = role;
  return getDashboardStats();
}

function getMyTransactions(role) {
  var r = String(role || '').trim().toUpperCase();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var metaMap = getTrackingMetaMap(ss);
  var tx = getTransactionsRows(ss);
  return tx.filter(function(x){
    var m = metaMap[x.trackingNo] || {};
    var office = String(x.requestingOffice || m.requestingOffice || '').trim().toUpperCase();
    return office === r;
  }).map(function(x){
    var m = metaMap[x.trackingNo] || {};
    return {
      trackingNo: x.trackingNo,
      prNo: x.prNo || m.prNo || '',
      office: x.requestingOffice || m.requestingOffice || '',
      currentSection: x.currentDept,
      currentStatus: x.currentStatus
    };
  });
}

function getMyRequestDepartmentBundle(role) {
  var rows = getMyTransactions(role);
  var wanted = {};
  for (var i = 0; i < rows.length; i++) {
    wanted[rows[i].trackingNo] = true;
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureCoreSheets(ss);
  var allRows = buildDepartmentRows('ALL');
  var histRows = getHistoryRows(ss);
  var workflowDeptMap = {
    'BAC': 'bac',
    'SUPPLY': 'supply',
    'BUDGET': 'budget',
    'ACCOUNTING': 'accounting',
    'CASH': 'cash',
    'RCAO': 'rcao',
    'ARDA': 'arda',
    'END USER': 'enduser'
  };

  var historyByTracking = {};
  for (var h = 0; h < histRows.length; h++) {
    var hr = histRows[h];
    historyByTracking[hr.trackingNo] = historyByTracking[hr.trackingNo] || [];
    historyByTracking[hr.trackingNo].push(hr);
  }

  function getReturnContext(trackingNo, targetDept) {
    var list = historyByTracking[trackingNo] || [];
    var target = normalizeDepartmentName(targetDept);
    for (var i = list.length - 1; i >= 0; i--) {
      var item = list[i];
      var action = String(item.action || '').trim().toUpperCase();
      if (action !== 'RETURNED') continue;

      var toDept = normalizeDepartmentName(item.toDept);
      if (toDept !== target) continue;

      var fromDept = normalizeDepartmentName(item.fromDept);
      var deptKey = workflowDeptMap[fromDept];
      if (!deptKey) continue;

      var receivedDate = '';
      var receivedRemarks = '';
      for (var j = i + 1; j < list.length; j++) {
        var nextItem = list[j];
        var nextToDept = normalizeDepartmentName(nextItem.toDept);
        if (nextToDept !== target) continue;
        if (String(nextItem.action || '').trim().toUpperCase() === 'RECEIVED') {
          receivedDate = nextItem.timestamp || '';
          receivedRemarks = nextItem.remarks || '';
          break;
        }
      }

      return {
        deptKey: deptKey,
        returnedFrom: displayDepartmentName(fromDept),
        returnTo: displayDepartmentName(target),
        returnRemarks: item.remarks || '',
        returnedDate: item.timestamp || '',
        returnReceivedDate: receivedDate,
        returnReceivedRemarks: receivedRemarks
      };
    }
    return null;
  }

  function getLatestReturnedEvent(trackingNo) {
    var list = historyByTracking[trackingNo] || [];
    for (var i = list.length - 1; i >= 0; i--) {
      var item = list[i];
      if (String(item.action || '').trim().toUpperCase() !== 'RETURNED') continue;
      return item;
    }
    return null;
  }

  var departmentRows = { bac:[], supply:[], budget:[], accounting:[], cash:[], rcao:[], arda:[], enduser:[] };
  for (var r = 0; r < allRows.length; r++) {
    var row = allRows[r];
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (!trackingNo) continue;

    var currentDept = normalizeDepartmentName(row['CURRENT DEPT']);
    var rowStatus = String(row['STATUS'] || '').trim().toUpperCase();
    var includeByOwner = !!wanted[trackingNo];
    var includeByEndUserReturn = currentDept === 'END USER' && rowStatus === 'RETURNED';
    if (!includeByOwner && !includeByEndUserReturn) continue;

    var directDeptKey = workflowDeptMap[currentDept];
    if (directDeptKey) {
      departmentRows[directDeptKey].push(row);
      // Also add completed-copy rows to depts that have already processed this transaction
      var txHistory = historyByTracking[trackingNo] || [];
      var addedPastDepts = {};
      addedPastDepts[directDeptKey] = true;
      for (var h2 = 0; h2 < txHistory.length; h2++) {
        var hItem = txHistory[h2];
        var hAct = String(hItem.action || '').toUpperCase();
        var hFrom = normalizeDepartmentName(hItem.fromDept || '');
        if (hAct !== 'FORWARD' && !(hAct === 'RETURNED' && hFrom === 'END USER')) continue;
        var pastKey = workflowDeptMap[hFrom];
        if (!pastKey || addedPastDepts[pastKey]) continue;
        addedPastDepts[pastKey] = true;
        var pastCopy = Object.assign({}, row, {
          'IS NOTIFIED': '',
          'IS NEW': '',
          'COMPLETED AT': hItem.timestamp || ''
        });
        if (pastKey === 'enduser') {
          pastCopy['COMPLETED TO'] = displayDepartmentName(hItem.toDept || '') || '—';
          pastCopy['FORWARD REMARKS'] = hItem.remarks || '';
        }
        departmentRows[pastKey].push(pastCopy);
      }
      continue;
    }

    // Fully completed transactions (CURRENT DEPT = COMPLETED): add done copies to past depts
    if (currentDept === 'COMPLETED') {
      var txHistC = historyByTracking[trackingNo] || [];
      var addedCompletedDepts = {};
      for (var hc = 0; hc < txHistC.length; hc++) {
        var hItemC = txHistC[hc];
        var hActC = String(hItemC.action || '').toUpperCase();
        var hFromC = normalizeDepartmentName(hItemC.fromDept || '');
        if (hActC !== 'FORWARD' && !(hActC === 'RETURNED' && hFromC === 'END USER')) continue;
        var completedPastKey = workflowDeptMap[hFromC];
        if (!completedPastKey || addedCompletedDepts[completedPastKey]) continue;
        addedCompletedDepts[completedPastKey] = true;
        var completedCopy = Object.assign({}, row, {
          'IS NOTIFIED': '',
          'IS NEW': '',
          'COMPLETED AT': hItemC.timestamp || ''
        });
        if (completedPastKey === 'enduser') {
          completedCopy['COMPLETED TO'] = displayDepartmentName(hItemC.toDept || '') || '—';
          completedCopy['FORWARD REMARKS'] = hItemC.remarks || '';
        }
        departmentRows[completedPastKey].push(completedCopy);
      }
      continue;
    }

    var returnCtx = getReturnContext(trackingNo, currentDept);
    if (!returnCtx) {
      var latestReturned = getLatestReturnedEvent(trackingNo);
      var fallbackFrom = latestReturned ? normalizeDepartmentName(latestReturned.fromDept) : '';
      var fallbackDeptKey = workflowDeptMap[fallbackFrom];
      if (!fallbackDeptKey) continue;

      returnCtx = {
        deptKey: fallbackDeptKey,
        returnedFrom: displayDepartmentName(fallbackFrom),
        returnTo: displayDepartmentName(currentDept),
        returnRemarks: latestReturned ? (latestReturned.remarks || '') : '',
        returnedDate: latestReturned ? (latestReturned.timestamp || '') : '',
        returnReceivedDate: '',
        returnReceivedRemarks: ''
      };
    }

    var mappedRow = Object.assign({}, row, {
      'RETURN TO': returnCtx.returnTo,
      'RETURNED FROM': returnCtx.returnedFrom,
      'RETURN REMARKS': returnCtx.returnRemarks,
      'RETURNED DATE': returnCtx.returnedDate,
      'RETURN RECEIVED DATE': returnCtx.returnReceivedDate,
      'RETURN RECEIVED REMARKS': returnCtx.returnReceivedRemarks,
      'STATUS': String(row['STATUS'] || '').trim() || 'RETURNED'
    });
    departmentRows[returnCtx.deptKey].push(mappedRow);
  }

  return {
    requests: rows,
    departmentRows: departmentRows
  };
}

function getDashboardLoadBundle() {
  var deptMap = getDeptPageDataAll();
  return {
    bac: deptMap.bac,
    supply: deptMap.supply,
    budget: deptMap.budget,
    accounting: deptMap.accounting,
    cash: deptMap.cash,
    rcao: deptMap.rcao,
    arda: deptMap.arda,
    enduser: deptMap.enduser,
    latestActivityByTracking: getLatestActivityByTracking(8000)
  };
}

function findRowNumByTracking(trackingNo) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tx = getTransactionsRows(ss);
  for (var i = 0; i < tx.length; i++) {
    if (tx[i].trackingNo === trackingNo) return tx[i]._rowNum;
  }
  return 0;
}

function updateEndUserSupplyFields(rowNum, data, username) {
  var _ = rowNum;
  var _2 = data;
  logAudit(username || 'SYSTEM', 'END USER EDIT', 'SUPPLY', 'N/A');
  return {success:true};
}

function createTransaction(data, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = generateTrackingNo(ss);
    var prNo = String(data['P.R #'] || data['P.R NO.'] || '').trim();
    var office = String(data['END-USER / REQUESTED BY'] || data['REQUESTING OFFICE'] || '').trim();
    var category = String(data['CATEGORY'] || '').trim();

    upsertTransactionSummary(ss, trackingNo, 'BAC', 'PROCESSING');
    setTransactionFieldsByTracking(ss, trackingNo, {
      'P.R NO.': prNo,
      'END-USER / REQUESTED BY': office,
      'CATEGORY': category,
      'PHIGEPS POSTED': data['PHIGEPS POSTED'] || '',
      'APPROVED ABC': data['APPROVED ABC'] || '',
      'RFQ': data['RFQ'] || '',
      'COMPLETE RFQ RETURNED': data['COMPLETE RFQ RETURNED'] || '',
      'RFQ DATE OF COMPLETION': data['RFQ DATE OF COMPLETION'] || '',
      'ABSTRACT OF AWARD': data['ABSTRACT OF AWARD'] || '',
      'AWARDED TO': data['AWARDED TO'] || '',
      'DATE OF AWARD': data['DATE OF AWARD'] || '',
      'MODE OF PROCUREMENT': data['MODE OF PROCUREMENT'] || '',
      'BAC RESO': data['BAC RESO'] || '',
      'BAC RESO STATUS': data['BAC RESO STATUS'] || '',
      'NOA': data['NOA'] || '',
      'NTP': data['NTP'] || ''
    });

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: '',
      toDept: 'BAC',
      action: 'Created',
      status: 'PROCESSING',
      remarks: '',
      processedBy: username || 'SYSTEM'
    });

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: '',
      toDept: 'BAC',
      action: 'Received',
      status: 'PROCESSING',
      remarks: '',
      processedBy: username || 'SYSTEM'
    });

    logAudit(username || 'SYSTEM', 'CREATE', 'TRANSACTIONS', trackingNo);
    return {success:true, trackingNo:trackingNo, bacId:'BAC-' + trackingNo};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updateTransaction(sheetName, rowNum, data, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var dept = normalizeDepartmentName(sheetName);
    var status = String((data && data['STATUS']) || '').trim() || 'PROCESSING';
    var remarks = String((data && data['FORWARD REMARKS']) || '').trim();

    setTransactionFieldsByTracking(ss, trackingNo, data || {});

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: dept,
      toDept: dept,
      action: 'Updated',
      status: status,
      remarks: remarks,
      processedBy: username || 'SYSTEM'
    });

    upsertTransactionSummary(ss, trackingNo, dept, status);
    logAudit(username || 'SYSTEM', 'UPDATE', dept, trackingNo);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function forwardTransaction(sheetName, rowNum, targetSheetName, extraData, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var fromDept = normalizeDepartmentName(sheetName);
    var toDept = normalizeDepartmentName(targetSheetName);
    if (!toDept) throw new Error('Invalid forward target.');

    var forwardedStatus = 'FORWARDED TO ' + displayDepartmentName(toDept).toUpperCase();
    var status = String((extraData && extraData['STATUS']) || forwardedStatus).trim() || forwardedStatus;
    var remarks = String((extraData && extraData['FORWARD REMARKS']) || '').trim();
    var timestamp = getTimestamp();

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: fromDept,
      toDept: toDept,
      action: 'Forward',
      status: status,
      remarks: remarks,
      processedBy: username || 'SYSTEM'
    });

    setTransactionFieldsByTracking(ss, trackingNo, {
      'COMPLETED AT': timestamp
    });

    upsertTransactionSummary(ss, trackingNo, toDept, status);
    logAudit(username || 'SYSTEM', 'FORWARD', fromDept, trackingNo + ' -> ' + toDept);
    return {success:true, trackingNo:trackingNo};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function smartForwardTransaction(sheetName, rowNum, targetSheetName, extraData, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var toDept = normalizeDepartmentName(targetSheetName);
    if (!toDept) throw new Error('Invalid forward target.');

    // END USER is always a return; for all other depts check if they've received it before
    var history = getHistoryRows(ss).filter(function(r) { return r.trackingNo === trackingNo; });
    var isReturn = toDept === 'END USER' || history.some(function(h) { return normalizeDepartmentName(h.toDept) === toDept; });

    if (isReturn) {
      var remarks = String((extraData && extraData['FORWARD REMARKS']) || '').trim();
      return returnTransaction(sheetName, rowNum, targetSheetName, remarks, username);
    } else {
      return forwardTransaction(sheetName, rowNum, targetSheetName, extraData, username);
    }
  } catch (e) {
    return {success: false, message: e.message};
  }
}

function receiveTransaction(sheetName, rowNum, remarks, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var dept = normalizeDepartmentName(sheetName);
    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: '',
      toDept: dept,
      action: 'Received',
      status: 'PROCESSING',
      remarks: String(remarks || '').trim(),
      processedBy: username || 'SYSTEM'
    });

    upsertTransactionSummary(ss, trackingNo, dept, 'PROCESSING');
    logAudit(username || 'SYSTEM', 'RECEIVE', dept, trackingNo);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function returnTransaction(sheetName, rowNum, returnTo, remarks, username, markCompleted) {
  var shouldMarkCompleted = markCompleted === true || String(markCompleted || '').toLowerCase() === 'true';
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var fromDept = normalizeDepartmentName(sheetName);
    var toDept = normalizeDepartmentName(returnTo);
    if (!toDept) throw new Error('Return target is required.');

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: fromDept,
      toDept: toDept,
      action: 'Returned',
      status: 'RETURNED',
      remarks: String(remarks || '').trim(),
      processedBy: username || 'SYSTEM'
    });

    if (shouldMarkCompleted) {
      appendHistoryLog(ss, {
        trackingNo: trackingNo,
        fromDept: toDept,
        toDept: 'COMPLETED',
        action: 'Completed',
        status: 'COMPLETED',
        remarks: '',
        processedBy: username || 'SYSTEM'
      });
      upsertTransactionSummary(ss, trackingNo, 'COMPLETED', 'COMPLETED');
      logAudit(username || 'SYSTEM', 'RETURN+COMPLETE', fromDept, trackingNo + ' -> ' + toDept);
      return {success:true, target:'COMPLETED', completed:true};
    }

    upsertTransactionSummary(ss, trackingNo, toDept, 'RETURNED');
    logAudit(username || 'SYSTEM', 'RETURN', fromDept, trackingNo + ' -> ' + toDept);
    return {success:true, target:toDept};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function completeReturnedTransaction(sheetName, rowNum, remarks, username) {
  return receiveTransaction(sheetName, rowNum, remarks, username);
}

function receiveAndForwardReturnedTransaction(sheetName, rowNum, forwardTo, forwardRemarks, username) {
  var first = receiveTransaction(sheetName, rowNum, forwardRemarks, username);
  if (!first.success) return first;
  return forwardTransaction(sheetName, rowNum, forwardTo, {'FORWARD REMARKS': forwardRemarks}, username);
}

function markCompleted(rowNum, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var txRows = getTransactionsRows(ss);
    var currentDept = '';
    var currentStatus = '';
    for (var i = 0; i < txRows.length; i++) {
      if (txRows[i].trackingNo === trackingNo) {
        currentDept = txRows[i].currentDept;
        currentStatus = txRows[i].currentStatus;
        break;
      }
    }

    if (normalizeDepartmentName(currentDept) === 'COMPLETED' || String(currentStatus || '').trim().toUpperCase() === 'COMPLETED') {
      return {success:true, alreadyCompleted:true};
    }

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: currentDept,
      toDept: 'COMPLETED',
      action: 'Completed',
      status: 'COMPLETED',
      remarks: '',
      processedBy: username || 'SYSTEM'
    });

    upsertTransactionSummary(ss, trackingNo, 'COMPLETED', 'COMPLETED');
    logAudit(username || 'SYSTEM', 'COMPLETE', 'CASH', trackingNo);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function cancelTransaction(sheetName, rowNum, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var trackingNo = getTrackingNoByRowNum(ss, rowNum);
    if (!trackingNo) throw new Error('Transaction row not found.');

    var dept = normalizeDepartmentName(sheetName);
    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: dept,
      toDept: dept,
      action: 'Cancelled',
      status: 'CANCELLED',
      remarks: '',
      processedBy: username || 'SYSTEM'
    });

    upsertTransactionSummary(ss, trackingNo, dept, 'CANCELLED');
    logAudit(username || 'SYSTEM', 'CANCEL', dept, trackingNo);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updatePrNumber(trackingNo, newPrNo, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    appendHistoryLog(ss, {
      trackingNo: trackingNo,
      fromDept: '',
      toDept: '',
      action: 'Updated',
      status: '',
      remarks: 'P.R NO UPDATED: ' + String(newPrNo || '').trim(),
      processedBy: username || 'SYSTEM'
    });

    logAudit(username || 'SYSTEM', 'UPDATE P.R #', 'TRANSACTIONS', trackingNo + ' -> ' + (newPrNo || '(cleared)'));
    return {success:true, updatedCount:1};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function searchTransaction(query) {
  try {
    var q = String(query || '').trim().toUpperCase();
    if (!q) return null;

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var tx = getTransactionsRows(ss);
    var match = null;
    for (var i = 0; i < tx.length; i++) {
      if (tx[i].trackingNo.toUpperCase().indexOf(q) !== -1) {
        match = tx[i];
        break;
      }
    }
    if (!match) return null;

    var history = getHistoryRows(ss).filter(function(r){ return r.trackingNo === match.trackingNo; });
    var timeline = ['BAC','SUPPLY','BUDGET','ACCOUNTING','CASH','RCAO','ARDA'].map(function(d){
      var reached = history.some(function(h){ return h.toDept === d || h.fromDept === d; });
      var last = null;
      for (var k = history.length - 1; k >= 0; k--) {
        if (history[k].toDept === d || history[k].fromDept === d) { last = history[k]; break; }
      }
      return {
        section: d === 'SUPPLY' ? 'SUPPLY AND PROPERTY' : d,
        reached: reached,
        status: last ? (last.status || '') : '',
        completed: match.currentStatus === 'COMPLETED' && normalizeDepartmentName(match.currentDept) === 'COMPLETED',
        completedAt: (match.currentStatus === 'COMPLETED' && last) ? last.timestamp : '',
        returnedFrom: '',
        returnRemarks: '',
        returnReceivedRemarks: '',
        returnedDate: ''
      };
    });

    var meta = getTrackingMetaMap(ss)[match.trackingNo] || {};
    return {
      trackingNo: match.trackingNo,
      prNo: match.prNo || meta.prNo || '',
      office: match.requestingOffice || meta.requestingOffice || '',
      item: '',
      currentSection: match.currentDept === 'SUPPLY' ? 'SUPPLY AND PROPERTY' : match.currentDept,
      currentStatus: match.currentStatus,
      currentReturnRemarks: '',
      currentReturnReceivedRemarks: '',
      timeline: timeline,
      history: history.map(function(h){
        return {
          'ACTION': h.action,
          'FROM': displayDepartmentName(h.fromDept),
          'TO': displayDepartmentName(h.toDept),
          'USER': h.processedBy,
          'REMARKS': h.remarks,
          'TIMESTAMP': h.timestamp
        };
      })
    };
  } catch (e) {
    return {error:e.message};
  }
}

function getSourceDeptRowByTracking(trackingNo, sourceDeptName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dept = normalizeDepartmentName(sourceDeptName);
  var rows = buildDepartmentRows(dept);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['TRACKING NO.'] || '').trim() === String(trackingNo || '').trim()) return rows[i];
  }
  return null;
}

function getUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.USERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var username = String(data[i][hm['USERNAME']] || '').trim();
      if (!username) continue;
      var firstName = hm['FIRST NAME'] !== undefined ? String(data[i][hm['FIRST NAME']] || '').trim() : '';
      var lastName = hm['LAST NAME'] !== undefined ? String(data[i][hm['LAST NAME']] || '').trim() : '';
      var role = String(data[i][hm['ROLE']] || '').trim().toUpperCase();
      var eu = hm['END USER'] !== undefined ? String(data[i][hm['END USER']] || '').trim().toUpperCase() === 'TRUE' : false;
      var status = hm['STATUS'] !== undefined ? String(data[i][hm['STATUS']] || '').trim().toUpperCase() : 'ACTIVE';

      out.push({
        _rowNum: i + 1,
        userId: 'USER' + zeroPad(i, 4),
        firstName: firstName,
        lastName: lastName,
        username: username,
        role: role,
        isEndUser: eu,
        isActive: status !== 'INACTIVE' && status !== 'FALSE' && status !== 'BLOCKED'
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function addUser(userData, adminUsername) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.USERS);
    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
    var hm = getHeaderMap(hdrs);

    var rows = getUsers();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].username.toLowerCase() === String(userData.username || '').trim().toLowerCase()) {
        return {success:false, message:'Username already exists.'};
      }
    }

    var row = new Array(hdrs.length);
    for (var j = 0; j < row.length; j++) row[j] = '';
    if (hm['FIRST NAME'] !== undefined) row[hm['FIRST NAME']] = String(userData.firstName || '').trim();
    if (hm['LAST NAME'] !== undefined) row[hm['LAST NAME']] = String(userData.lastName || '').trim();
    if (hm['USERNAME'] !== undefined) row[hm['USERNAME']] = String(userData.username || '').trim();
    if (hm['PASSWORD'] !== undefined) row[hm['PASSWORD']] = String(userData.password || '1234').trim();
    if (hm['ROLE'] !== undefined) row[hm['ROLE']] = String(userData.role || '').trim().toUpperCase();
    if (hm['END USER'] !== undefined) row[hm['END USER']] = userData.isEndUser ? 'TRUE' : '';
    if (hm['STATUS'] !== undefined) row[hm['STATUS']] = userData.isActive === false ? 'INACTIVE' : 'ACTIVE';

    sheet.appendRow(row);
    logAudit(adminUsername || 'ADMIN', 'ADD USER', 'USERS', String(userData.username || ''));
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updateUser(rowNum, userData, adminUsername) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) throw new Error('USERS sheet not found.');

    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    function set(key, val) {
      if (hm[key] !== undefined) sheet.getRange(rowNum, hm[key] + 1).setValue(val);
    }

    if (userData.firstName !== undefined) set('FIRST NAME', String(userData.firstName || '').trim());
    if (userData.lastName !== undefined) set('LAST NAME', String(userData.lastName || '').trim());
    if (userData.username !== undefined) set('USERNAME', String(userData.username || '').trim());
    if (userData.password) set('PASSWORD', String(userData.password));
    if (userData.role !== undefined) set('ROLE', String(userData.role || '').trim().toUpperCase());
    if (userData.isEndUser !== undefined) set('END USER', userData.isEndUser ? 'TRUE' : '');
    if (userData.isActive !== undefined) set('STATUS', userData.isActive ? 'ACTIVE' : 'INACTIVE');

    logAudit(adminUsername || 'ADMIN', 'UPDATE USER', 'USERS', 'Row ' + rowNum);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function deleteUser(rowNum, targetUsername, adminUsername) {
  try {
    if (String(targetUsername || '').toLowerCase() === String(adminUsername || '').toLowerCase()) {
      return {success:false, message:'You cannot delete your own account.'};
    }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) throw new Error('USERS sheet not found.');
    sheet.deleteRow(rowNum);
    logAudit(adminUsername || 'ADMIN', 'DELETE USER', 'USERS', targetUsername || ('Row ' + rowNum));
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updateMyCredentials(currentUsername, currentPassword, profileData) {
  try {
    var username = String(currentUsername || '').trim();
    var password = String(currentPassword || '').trim();
    var requestedUsername = String((profileData && profileData.username) || '').trim();
    var requestedNewPassword = String((profileData && profileData.newPassword) || '').trim();

    if (!username) return {success:false, message:'Current user is required.'};
    if (!requestedUsername) return {success:false, message:'Username is required.'};

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) return {success:false, message:'USERS sheet not found.'};

    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var uCol = hm['USERNAME'];
    var pCol = hm['PASSWORD'];
    var rCol = hm['ROLE'];
    var eCol = hm['END USER'];

    var targetRow = -1;
    var role = '';
    var isEU = false;

    for (var i = 1; i < data.length; i++) {
      var rowU = String(data[i][uCol] || '').trim();
      if (rowU.toLowerCase() !== username.toLowerCase()) continue;
      var rowP = String(data[i][pCol] || '').trim();
      if (requestedNewPassword && rowP !== password) return {success:false, message:'Current password is incorrect.'};
      targetRow = i + 1;
      role = String(data[i][rCol] || '').trim().toUpperCase();
      isEU = eCol !== undefined ? String(data[i][eCol] || '').trim().toUpperCase() === 'TRUE' : false;
      break;
    }

    if (targetRow === -1) return {success:false, message:'User account was not found.'};

    sheet.getRange(targetRow, uCol + 1).setValue(requestedUsername);
    if (requestedNewPassword) sheet.getRange(targetRow, pCol + 1).setValue(requestedNewPassword);
    
    var fNameCol = hm['FIRST NAME'];
    var lNameCol = hm['LAST NAME'];
    if (fNameCol !== undefined && profileData && profileData.firstName) sheet.getRange(targetRow, fNameCol + 1).setValue(String(profileData.firstName).trim());
    if (lNameCol !== undefined && profileData && profileData.lastName) sheet.getRange(targetRow, lNameCol + 1).setValue(String(profileData.lastName).trim());

    logAudit(requestedUsername, 'UPDATE MY PROFILE', 'USERS', 'Row ' + targetRow);
    return {success:true, username:requestedUsername, firstName:(profileData && profileData.firstName)||'', lastName:(profileData && profileData.lastName)||'', role:role, isEndUser:isEU};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function getAllSuppliers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);
    var nameCol = hm['SUPPLIER NAME'];
    var stCol = hm['STATUS'];

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][nameCol] || '').trim();
      if (!name) continue;
      out.push({
        _rowNum: i + 1,
        supplierId: 'SUPP' + zeroPad(i, 4),
        supplierName: name,
        status: stCol !== undefined ? normalizeSupplierStatus(data[i][stCol]) : 'ACTIVE'
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function addSupplier(supplierData, adminUsername) {
  try {
    var name = String((supplierData && supplierData.supplierName) || '').trim();
    if (!name) return {success:false, message:'Supplier name is required.'};

    var rows = getAllSuppliers();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].supplierName.toLowerCase() === name.toLowerCase()) return {success:false, message:'Supplier already exists.'};
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var row = new Array(hdrs.length);
    for (var j = 0; j < row.length; j++) row[j] = '';
    if (hm['SUPPLIER NAME'] !== undefined) row[hm['SUPPLIER NAME']] = name;
    if (hm['STATUS'] !== undefined) row[hm['STATUS']] = normalizeSupplierStatus(supplierData && supplierData.status);

    sheet.appendRow(row);
    logAudit(adminUsername || 'ADMIN', 'ADD SUPPLIER', 'SUPPLIERS', name);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updateSupplier(rowNum, supplierData, adminUsername) {
  try {
    var name = String((supplierData && supplierData.supplierName) || '').trim();
    if (!name) return {success:false, message:'Supplier name is required.'};

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet) throw new Error('SUPPLIERS sheet not found.');

    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    if (hm['SUPPLIER NAME'] !== undefined) sheet.getRange(rowNum, hm['SUPPLIER NAME'] + 1).setValue(name);
    if (hm['STATUS'] !== undefined) sheet.getRange(rowNum, hm['STATUS'] + 1).setValue(normalizeSupplierStatus(supplierData && supplierData.status));

    logAudit(adminUsername || 'ADMIN', 'UPDATE SUPPLIER', 'SUPPLIERS', 'Row ' + rowNum);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function deleteSupplier(rowNum, supplierName, adminUsername) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.SUPPLIERS);
    if (!sheet) throw new Error('SUPPLIERS sheet not found.');
    sheet.deleteRow(rowNum);
    logAudit(adminUsername || 'ADMIN', 'DELETE SUPPLIER', 'SUPPLIERS', supplierName || ('Row ' + rowNum));
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function ensureEndUsersSheet(sheet) {
  ensureSheetHeaders(sheet, ['END USER ROLE','STATUS']);
}

function getEndUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.END_USERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);
    var roleCol = hm['END USER ROLE'];
    var statusCol = hm['STATUS'];

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][roleCol] || '').trim();
      if (!role) continue;
      var st = statusCol !== undefined ? String(data[i][statusCol] || '').trim().toUpperCase() : 'ACTIVE';
      if (st === 'INACTIVE') continue;
      out.push(role);
    }
    return out.sort();
  } catch (e) {
    return [];
  }
}

function getAllEndUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    var sheet = ss.getSheetByName(SHEETS.END_USERS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var role = String(data[i][hm['END USER ROLE']] || '').trim();
      if (!role) continue;
      out.push({
        _rowNum: i + 1,
        endUserId: 'EU' + zeroPad(i, 4),
        endUserRole: role,
        status: hm['STATUS'] !== undefined ? String(data[i][hm['STATUS']] || '').trim().toUpperCase() : 'ACTIVE'
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function addEndUser(endUserData, adminUsername) {
  try {
    var role = String((endUserData && endUserData.endUserRole) || '').trim();
    if (!role) return {success:false, message:'End User Role is required.'};

    var rows = getAllEndUsers();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].endUserRole.toLowerCase() === role.toLowerCase()) return {success:false, message:'End User Role already exists.'};
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.END_USERS);
    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    var row = new Array(hdrs.length);
    for (var j = 0; j < row.length; j++) row[j] = '';
    if (hm['END USER ROLE'] !== undefined) row[hm['END USER ROLE']] = role;
    if (hm['STATUS'] !== undefined) row[hm['STATUS']] = String((endUserData && endUserData.status) || 'ACTIVE').trim().toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE';

    sheet.appendRow(row);
    logAudit(adminUsername || 'ADMIN', 'ADD END USER', 'END_USERS', role);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function updateEndUser(rowNum, endUserData, adminUsername) {
  try {
    var role = String((endUserData && endUserData.endUserRole) || '').trim();
    if (!role) return {success:false, message:'End User Role is required.'};

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.END_USERS);
    if (!sheet) throw new Error('END_USERS sheet not found.');

    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
    var hm = getHeaderMap(hdrs);

    if (hm['END USER ROLE'] !== undefined) sheet.getRange(rowNum, hm['END USER ROLE'] + 1).setValue(role);
    if (hm['STATUS'] !== undefined) {
      var st = String((endUserData && endUserData.status) || 'ACTIVE').trim().toUpperCase();
      sheet.getRange(rowNum, hm['STATUS'] + 1).setValue(st === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE');
    }

    logAudit(adminUsername || 'ADMIN', 'UPDATE END USER', 'END_USERS', 'Row ' + rowNum);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function deleteEndUser(rowNum, endUserName, adminUsername) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.END_USERS);
    if (!sheet) throw new Error('END_USERS sheet not found.');
    sheet.deleteRow(rowNum);
    logAudit(adminUsername || 'ADMIN', 'DELETE END USER', 'END_USERS', endUserName || ('Row ' + rowNum));
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function getAuditLogs(limitRows) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);

    // Build username → full name lookup from USERS sheet
    var nameMap = {};
    try {
      var uSheet = ss.getSheetByName(SHEETS.USERS);
      if (uSheet && uSheet.getLastRow() >= 2) {
        var uData = uSheet.getDataRange().getValues();
        var uHdrs = uData[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
        var uHm = getHeaderMap(uHdrs);
        for (var u = 1; u < uData.length; u++) {
          var uname = String(uData[u][uHm['USERNAME']] || '').trim().toLowerCase();
          if (!uname) continue;
          var fn = uHm['FIRST NAME'] !== undefined ? String(uData[u][uHm['FIRST NAME']] || '').trim() : '';
          var ln = uHm['LAST NAME']  !== undefined ? String(uData[u][uHm['LAST NAME']]  || '').trim() : '';
          nameMap[uname] = (fn + ' ' + ln).trim() || uname;
        }
      }
    } catch(eu) {}

    var sheet = ss.getSheetByName(SHEETS.AUDIT);
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getDataRange().getValues();
    var out = [];
    var limit = Math.max(1, parseInt(limitRows, 10) || 500);

    for (var i = data.length - 1; i >= 1 && out.length < limit; i--) {
      var rawUser = String(data[i][1] || '');
      var fullName = nameMap[rawUser.toLowerCase()] || rawUser;
      out.push({
        rowNum: i + 1,
        _rowNum: i + 1,
        timestamp: formatCellForClient('TIMESTAMP', data[i][0]),
        username: rawUser,
        fullName: fullName,
        action: String(data[i][2] || ''),
        department: String(data[i][3] || ''),
        recordId: String(data[i][4] || '')
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function deleteAuditLogs(keys) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);
    var sheet = ss.getSheetByName(SHEETS.AUDIT);
    if (!sheet || sheet.getLastRow() < 2) return {success:true, deletedCount:0};

    var keyList = Array.isArray(keys) ? keys : [];
    if (!keyList.length) return {success:false, message:'No audit keys provided.'};

    var numericRows = keyList
      .map(function(k){ return parseInt(k, 10); })
      .filter(function(n){ return !isNaN(n) && n >= 2; });

    var deleted = 0;

    if (numericRows.length) {
      var seen = {};
      var rows = numericRows.filter(function(n){
        if (seen[n]) return false;
        seen[n] = true;
        return true;
      }).sort(function(a,b){ return b-a; });

      for (var i = 0; i < rows.length; i++) {
        var rn = rows[i];
        if (rn <= sheet.getLastRow()) {
          sheet.deleteRow(rn);
          deleted++;
        }
      }
      return {success:true, deletedCount:deleted};
    }

    // Fallback: match by timestamp string if row numbers are not provided.
    var wantedTs = {};
    keyList.forEach(function(k){ wantedTs[String(k || '').trim()] = true; });
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var r = 1; r < data.length; r++) {
      var ts = formatCellForClient('TIMESTAMP', data[r][0]);
      if (wantedTs[ts]) rowsToDelete.push(r + 1);
    }
    rowsToDelete.sort(function(a,b){ return b-a; }).forEach(function(rn){
      sheet.deleteRow(rn);
      deleted++;
    });

    return {success:true, deletedCount:deleted};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function getLatestActivityByTracking(limitRows) {
  try {
    var rows = getHistoryRows(SpreadsheetApp.openById(SPREADSHEET_ID));
    var out = {};
    var limit = Math.max(1, parseInt(limitRows, 10) || 5000);

    for (var i = rows.length - 1; i >= 0; i--) {
      var r = rows[i];
      if (!out[r.trackingNo]) {
        out[r.trackingNo] = {
          timestamp: r.timestamp,
          action: r.action,
          fromDept: displayDepartmentName(r.fromDept),
          toDept: displayDepartmentName(r.toDept),
          status: r.status,
          remarks: r.remarks,
          user: r.processedBy
        };
      }
      if (Object.keys(out).length >= limit) break;
    }

    return out;
  } catch (e) {
    return {};
  }
}

function setupSheets() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureCoreSheets(ss);
    return {success:true};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

function continueReturnedTransaction(sheetName, rowNum, username) {
  return receiveTransaction(sheetName, rowNum, 'Reopened from return', username);
}

function continueReturnedNew(sheetName, rowNum, username) {
  return receiveTransaction(sheetName, rowNum, 'Continue returned transaction', username);
}