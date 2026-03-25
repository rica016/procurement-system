// ================================================================
//  DAR PROCUREMENT MONITORING SYSTEM — Code.gs  v8
//  Department of Agrarian Reform, Regional Office V
//
//  SHEET HEADERS (exact, from DAR_PROCUREMENT.xlsx):
//  ─────────────────────────────────────────────────────────────
//  END USER  : END USER ID | TRACKING NO. | P.R # | REQUESTING OFFICE | STATUS | DATE CREATED | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  BAC       : BAC ID | TRACKING NO. | P.R # | DATE RECEIVED | END-USER / REQUESTED BY | CATEGORY | PHIGEPS POSTED | APPROVED ABC | RFQ | RFQ DATE OF COMPLETION | ABSTRACT OF AWARD | AWARDED TO | DATE OF AWARD | MODE OF PROCUREMENT | BAC RESO | NOA | NTP | STATUS | COMPLETED AT | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  SUPPLY    : SP ID | TRACKING NO. | DATE RECEIVED | P.R # | END USER | P.O. No. | P.O. DATE | DOCUMENT TYPE | SUPPLIER | PARTICULARS | AMOUNT | DATE TRANSMITTED TO SUPPLIER | DATE RECEIVED BY SUPPLIER | COA DATE RECEIVED | I.A.R # | INVOICE # | STATUS | COMPLETED AT | DATE RETURNED TO SUPPLY | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  BUDGET    : BUDGET ID | TRACKING NO. | DATE RECEIVED | P.R # | ORS NO. | BUPR NO. | AMOUNT | STATUS | COMPLETED AT | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  ACCOUNTING: ACCOUNTING ID | TRACKING NO. | DATE RECEIVED | P.R # | DV No. | Net Amount | TAX Amount | STATUS | COMPLETED AT | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  CASH      : CASH ID | TRACKING NO. | DATE RECEIVED | P.R # | CHEQUE NO. / LDDAP | DATE OF CHEQUE | STATUS | COMPLETED AT | IS NEW | FORWARD REMARKS | RETURN TO | RETURNED FROM | RETURN REMARKS | RETURNED DATE | RETURN RECEIVED DATE | RETURN RECEIVED REMARKS
//  SUPPLIERS : SUPPLIER ID | SUPPLIER NAME | STATUS
//  USERS     : USER ID | USERNAME | PASSWORD | ROLE | END USER | ACTIVE
//  AUDIT_LOGS: TIMESTAMP | USER | ACTION | DEPARTMENT | RECORD ID
//
//  FLOW (v8):
//   BAC is the START of all transactions — only BAC/ADMIN can create.
//   End users see rows where END USER matches their role (Supply sheet primary).
//   End users may edit: DATE TRANSMITTED TO SUPPLIER, DATE RECEIVED BY SUPPLIER, COA DATE RECEIVED on their Supply row.
//
//  END USER STATUS FLOW (auto-updated on every dept action):
//   BAC creates     → FORWARDED TO BAC (set on BAC row creation)
//   BAC receives    → RECEIVED BY BAC
//   BAC forwards    → FORWARDED TO SUPPLY
//   Supply receives → RECEIVED BY SUPPLY
//   Supply forwards → FORWARDED TO BUDGET
//   Budget receives → RECEIVED BY BUDGET
//   Budget forwards → FORWARDED TO ACCOUNTING
//   Accounting rcv  → RECEIVED BY ACCOUNTING
//   Accounting fwd  → FORWARDED TO CASHIER
//   Cash receives   → RECEIVED BY CASHIER
//   Cash marks paid → PAID
//   Any cancel      → CANCELLED
// ================================================================

var SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

var END_USER_ROLES = ['HR','GSS','LTID','LEGAL','FINANCE','IT','ADMIN DIVISION','DARPO'];

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('DAR RO V — Procurement Monitoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
}
function formatCellForClient(header, value) {
  if (value === null || value === undefined) return '';
  if (!(value instanceof Date)) return String(value);

  var key = String(header || '').trim().toUpperCase();
  var withTime = (key === 'DATE RECEIVED' || key === 'COMPLETED AT' || key === 'TIMESTAMP' || key === 'RETURNED DATE' || key === 'RETURN RECEIVED DATE');
  var pattern = withTime ? 'MMM dd, yyyy hh:mm a' : 'MMM dd, yyyy';
  return Utilities.formatDate(value, Session.getScriptTimeZone(), pattern);
}
function zeroPad(num, size) {
  var s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

function normalizeSupplierStatus(status) {
  var v = String(status || '').trim().toUpperCase();
  if (v === 'INACTIVE' || v === 'BLOCKED') return v;
  return 'ACTIVE';
}

function ensureSuppliersSheet(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['SUPPLIER ID','SUPPLIER NAME','STATUS']);
    return;
  }
  var hdrs = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0]
    .map(function(h){ return String(h || '').trim().toUpperCase(); });
  if (hdrs.indexOf('SUPPLIER ID') === -1 && hdrs.indexOf('SUPPLIER NAME') === -1) {
    sheet.getRange(1,1,1,3).setValues([['SUPPLIER ID','SUPPLIER NAME','STATUS']]);
    return;
  }
  if (hdrs.indexOf('STATUS') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('STATUS');
  }
}

// ── EXACT SHEET HEADERS ───────────────────────────────────────────────────────
var SHEET_HEADERS = {
  'END USER': [
    'END USER ID','TRACKING NO.','P.R #','REQUESTING OFFICE','STATUS','DATE CREATED','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ],
  'BAC': [
    'BAC ID','TRACKING NO.','P.R #','DATE RECEIVED',
    'END-USER / REQUESTED BY','CATEGORY','PHIGEPS POSTED','APPROVED ABC','RFQ','RFQ DATE OF COMPLETION',
    'ABSTRACT OF AWARD','AWARDED TO','DATE OF AWARD','MODE OF PROCUREMENT',
    'BAC RESO','NOA','NTP','STATUS','COMPLETED AT','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ],
  'SUPPLY AND PROPERTY': [
    'SP ID','TRACKING NO.','DATE RECEIVED','P.R #','END USER',
    'P.O. No.','P.O. DATE','DOCUMENT TYPE','SUPPLIER','PARTICULARS','AMOUNT',
    'DATE TRANSMITTED TO SUPPLIER','DATE RECEIVED BY SUPPLIER',
    'COA DATE RECEIVED','I.A.R #','INVOICE #',
    'STATUS','COMPLETED AT','DATE RETURNED TO SUPPLY','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ],
  'BUDGET': [
    'BUDGET ID','TRACKING NO.','DATE RECEIVED','P.R #',
    'ORS NO.','BUPR NO.','AMOUNT','STATUS','COMPLETED AT','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ],
  'ACCOUNTING': [
    'ACCOUNTING ID','TRACKING NO.','DATE RECEIVED','P.R #',
    'DV No.','Net Amount','TAX Amount','STATUS','COMPLETED AT','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ],
  'CASH': [
    'CASH ID','TRACKING NO.','DATE RECEIVED','P.R #',
    'CHEQUE NO. / LDDAP','DATE OF CHEQUE','STATUS','COMPLETED AT','IS NEW',
    'FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS'
  ]
};

var ID_CONFIG = {
  'END USER':            { col:'END USER ID',   prefix:'EU'   },
  'BAC':                 { col:'BAC ID',         prefix:'BAC'  },
  'SUPPLY AND PROPERTY': { col:'SP ID',          prefix:'SP'   },
  'BUDGET':              { col:'BUDGET ID',      prefix:'BUD'  },
  'ACCOUNTING':          { col:'ACCOUNTING ID',  prefix:'ACC'  },
  'CASH':                { col:'CASH ID',        prefix:'CASH' }
};

// These columns are NEVER overwritten by updateTransaction()
var PROTECTED_COLS = [
  'BAC ID','SP ID','BUDGET ID','ACCOUNTING ID','CASH ID','END USER ID',
  'TRACKING NO.','DATE RECEIVED','DATE CREATED','COMPLETED AT'
];

// ── END USER STATUS LABELS ─────────────────────────────────────────────────────
var DEPT_RECEIVE_LABELS = {
  'BAC':                 'RECEIVED BY BAC',
  'SUPPLY AND PROPERTY': 'RECEIVED BY SUPPLY',
  'BUDGET':              'RECEIVED BY BUDGET',
  'ACCOUNTING':          'RECEIVED BY ACCOUNTING',
  'CASH':                'RECEIVED BY CASHIER'
};

var DEPT_FORWARD_LABELS = {
  'BAC':                 'FORWARDED TO SUPPLY',
  'SUPPLY AND PROPERTY': 'FORWARDED TO BUDGET',
  'BUDGET':              'FORWARDED TO ACCOUNTING',
  'ACCOUNTING':          'FORWARDED TO CASHIER'
};

var DEPT_STATUS_OPTIONS = {
  'BAC':                 ['PROCESSING','FOR CANVAS','FOR REVIEW','FOR SIGNATURE','RETURNED TO REQUESTING OFFICE','FORWARDED TO SUPPLY','RETURNED','CANCELLED'],
  'SUPPLY AND PROPERTY': ['PROCESSING','RECEIVED','FORWARDED TO BUDGET','RETURNED TO REQUESTING OFFICE','RETURNED','CANCELLED'],
  'BUDGET':              ['PROCESSING','RECEIVED','FORWARDED TO ACCOUNTING','RETURNED TO REQUESTING OFFICE','RETURNED','CANCELLED'],
  'ACCOUNTING':          ['PROCESSING','RECEIVED','FORWARDED TO CASHIER','RETURNED TO REQUESTING OFFICE','RETURNED','CANCELLED'],
  'CASH':                ['PROCESSING','RECEIVED','LACKING OF ATTACHMENTS','UNPAID','RETURNED','CANCELLED']
};

function syncStatusValidationForRow(sheet, sheetName, rowNum, hdrs, overrideOptions) {
  var options = overrideOptions || DEPT_STATUS_OPTIONS[sheetName];
  if (!sheet || !options || !options.length) return;
  var statusCol = hdrs.indexOf('STATUS');
  if (statusCol === -1) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(rowNum, statusCol + 1).setDataValidation(rule);
}

// ── UPDATE END USER STATUS (syncs END USER sheet) ────────────────────────────
function updateEndUserStatus(ss, trackingNo, newStatus) {
  try {
    if (!trackingNo) return;
    var sheet = ss.getSheetByName('END USER');
    if (!sheet || sheet.getLastRow() < 2) return;
    var data  = sheet.getDataRange().getValues();
    var hdrs  = data[0].map(function(h){ return String(h||'').trim(); });
    var tnIdx = hdrs.indexOf('TRACKING NO.');
    var stIdx = hdrs.indexOf('STATUS');
    if (tnIdx===-1||stIdx===-1) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][tnIdx]||'').trim() === trackingNo) {
        sheet.getRange(i+1, stIdx+1).setValue(newStatus);
        return;
      }
    }
  } catch(e) { /* silent */ }
}

// ── ENSURE OPTIONAL COLUMNS ───────────────────────────────────────────────────
function ensureColumn(sheet, colName) {
  if (!sheet || sheet.getLastColumn()===0) return;
  var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                  .map(function(h){ return String(h||'').trim(); });
  if (hdrs.indexOf(colName)===-1)
    sheet.getRange(1, sheet.getLastColumn()+1).setValue(colName);
}
function ensureIsNew(s)       { ensureColumn(s,'IS NEW'); }
function ensureParticulars(s) { ensureColumn(s,'PARTICULARS'); }
function ensureWorkflowColumns(s) {
  ['FORWARDED FROM','FORWARD REMARKS','RETURN TO','RETURNED FROM','RETURN REMARKS','RETURNED DATE','RETURN RECEIVED DATE','RETURN RECEIVED REMARKS']
    .forEach(function(col){ ensureColumn(s, col); });
}
function normalizeDepartmentName(name) {
  var raw = String(name || '').trim().toUpperCase();
  var map = {
    'BAC':'BAC',
    'SUPPLY':'SUPPLY AND PROPERTY',
    'SUPPLY AND PROPERTY':'SUPPLY AND PROPERTY',
    'BUDGET':'BUDGET',
    'ACCOUNTING':'ACCOUNTING',
    'CASH':'CASH',
    'CASHIER':'CASH',
    'END USER':'END USER',
    'REQUESTING OFFICE':'END USER'
  };
  return map[raw] || raw;
}
function displayDepartmentName(name) {
  var normalized = normalizeDepartmentName(name);
  var map = {
    'BAC':'BAC',
    'SUPPLY AND PROPERTY':'Supply',
    'BUDGET':'Budget',
    'ACCOUNTING':'Accounting',
    'CASH':'Cashier',
    'END USER':'Requesting Office'
  };
  return map[normalized] || normalized;
}
function isReturnPendingRecord(record) {
  return !!record && String(record['RETURNED DATE'] || '').trim() !== '' && String(record['RETURN RECEIVED DATE'] || '').trim() === '';
}
function isReturnResolvedRecord(record) {
  return !!record && String(record['RETURNED DATE'] || '').trim() !== '' && String(record['RETURN RECEIVED DATE'] || '').trim() !== '';
}
function isOpenWorkflowRecord(record) {
  var status = String((record && record['STATUS']) || '').trim().toUpperCase();
  return !!record && !isReturnPendingRecord(record) && String(record['COMPLETED AT'] || '').trim() === '' && status !== 'CANCELLED';
}
function getRowRecord(sheet, rowNum) {
  var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
  var vals = sheet.getRange(rowNum,1,1,hdrs.length).getValues()[0];
  var out = {_rowNum:rowNum};
  for (var i = 0; i < hdrs.length; i++) {
    if (!hdrs[i]) continue;
    out[hdrs[i]] = formatCellForClient(hdrs[i], vals[i]);
  }
  return {headers:hdrs, values:vals, record:out};
}
function getUserContextByUsername(ss, username) {
  var uname = String(username || '').trim();
  if (!uname) return null;
  var usersSheet = ss.getSheetByName('USERS');
  if (!usersSheet) return null;
  var data = usersSheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var hdrs = data[0].map(function(h){ return String(h || '').trim().toUpperCase(); });
  var uCol = hdrs.indexOf('USERNAME');
  var rCol = hdrs.indexOf('ROLE');
  var eCol = hdrs.indexOf('END USER');
  if (uCol === -1 || rCol === -1) return null;

  for (var i = 1; i < data.length; i++) {
    var rowUname = String(data[i][uCol] || '').trim();
    if (rowUname.toLowerCase() !== uname.toLowerCase()) continue;
    var role = String(data[i][rCol] || '').trim().toUpperCase();
    var isEU = eCol !== -1 ? String(data[i][eCol] || '').trim().toUpperCase() === 'TRUE' : false;
    return { username: rowUname, role: role, isEndUser: isEU };
  }
  return null;
}
function canUserCompleteReturnedItem(userCtx, currentDepartment, returnTo) {
  if (!userCtx) return false;
  var role = String(userCtx.role || '').trim().toUpperCase();
  if (role === 'ADMIN') return true;

  if (returnTo === 'END USER') {
    return !!userCtx.isEndUser || END_USER_ROLES.indexOf(role) !== -1;
  }

  var roleDeptMap = {
    'BAC':'BAC',
    'SUPPLY':'SUPPLY AND PROPERTY',
    'BUDGET':'BUDGET',
    'ACCOUNTING':'ACCOUNTING',
    'CASH':'CASH',
    'CASHIER':'CASH'
  };
  return roleDeptMap[role] === currentDepartment;
}
function setRecordValuesByHeader(sheet, rowNum, headers, currentValues, updates) {
  var next = currentValues.slice();
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || '').trim();
    if (!key || !updates.hasOwnProperty(key)) continue;
    next[i] = updates[key];
  }
  sheet.getRange(rowNum,1,1,headers.length).setValues([next]);
}
function normalizeTrackingKey(v) {
  return String(v || '').replace(/\s+/g,'').trim().toUpperCase();
}
function findLatestRowNumByTracking(sheet, trackingNo) {
  if (!sheet || !trackingNo || sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var tnIdx = hdrs.indexOf('TRACKING NO.');
  if (tnIdx === -1) return 0;
  var wanted = normalizeTrackingKey(trackingNo);
  for (var i = data.length - 1; i >= 1; i--) {
    if (normalizeTrackingKey(data[i][tnIdx]) === wanted) return i + 1;
  }
  return 0;
}
function findLatestRowNumByIdentity(sheet, trackingNo, prNo) {
  var rowNum = findLatestRowNumByTracking(sheet, trackingNo);
  if (rowNum) return rowNum;
  if (!sheet || !prNo || sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var prIdx = hdrs.indexOf('P.R #');
  if (prIdx === -1) return 0;
  var wantedPr = String(prNo || '').trim();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][prIdx] || '').trim() === wantedPr) return i + 1;
  }
  return 0;
}
function findPendingReturnSourceRow(sheet, trackingNo, targetLabel) {
  if (!sheet || !trackingNo || sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var hdrs = data[0].map(function(h){ return String(h || '').trim(); });
  var tnIdx = hdrs.indexOf('TRACKING NO.');
  var rtIdx = hdrs.indexOf('RETURN TO');
  var rdIdx = hdrs.indexOf('RETURNED DATE');
  var rrIdx = hdrs.indexOf('RETURN RECEIVED DATE');
  if (tnIdx === -1 || rtIdx === -1 || rdIdx === -1 || rrIdx === -1) return findLatestRowNumByTracking(sheet, trackingNo);

  var wantedTarget = normalizeDepartmentName(targetLabel);
  var fallbackPendingAnyTarget = 0;
  var fallbackReturnedAnyState = 0;
  var fallbackAnyTracking = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][tnIdx] || '').trim() !== trackingNo) continue;

    if (!fallbackAnyTracking) fallbackAnyTracking = i + 1;

    if (String(data[i][rdIdx] || '').trim() === '') continue;

    if (!fallbackReturnedAnyState) fallbackReturnedAnyState = i + 1;

    var isPending = String(data[i][rrIdx] || '').trim() === '';
    if (!isPending) continue;

    var rowTarget = normalizeDepartmentName(String(data[i][rtIdx] || '').trim());
    if (wantedTarget && rowTarget === wantedTarget) return i + 1;
    if (!fallbackPendingAnyTarget) fallbackPendingAnyTarget = i + 1;
  }

  if (fallbackPendingAnyTarget) return fallbackPendingAnyTarget;
  if (fallbackReturnedAnyState) return fallbackReturnedAnyState;
  return fallbackAnyTracking;
}

// ── AUTO-GENERATE ROW ID ──────────────────────────────────────────────────────
function generateNextId(sheet, sheetName) {
  var cfg = ID_CONFIG[sheetName];
  if (!cfg) return '';
  var lastRow = sheet.getLastRow();
  var hdrs = lastRow>0
    ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();})
    : [];
  var colIdx = hdrs.indexOf(cfg.col);
  if (colIdx===-1) return cfg.prefix+'0001';
  var maxNum = 0;
  if (lastRow>=2) {
    sheet.getRange(2,colIdx+1,lastRow-1,1).getValues().forEach(function(r){
      var id=String(r[0]||'');
      if (id.indexOf(cfg.prefix)===0){
        var n=parseInt(id.substring(cfg.prefix.length))||0;
        if(n>maxNum) maxNum=n;
      }
    });
  }
  return cfg.prefix+zeroPad(maxNum+1,4);
}

// ── AUTO-GENERATE TRACKING NO. ────────────────────────────────────────────────
function generateTrackingNo(ss) {
  var now  = new Date();
  var base = 'DAR-'+now.getFullYear()+'-'+zeroPad(now.getMonth()+1,2)+'-';
  var maxNum = 0;
  ['BAC'].forEach(function(sname){
    var ws=ss.getSheetByName(sname);
    if (!ws||ws.getLastRow()<2) return;
    var hdrs=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
    var tnCol=hdrs.indexOf('TRACKING NO.');
    if (tnCol===-1) return;
    ws.getRange(2,tnCol+1,ws.getLastRow()-1,1).getValues().forEach(function(r){
      var tn=String(r[0]||'');
      if (tn.indexOf(base)===0){
        var n=parseInt(tn.substring(base.length))||0;
        if(n>maxNum) maxNum=n;
      }
    });
  });
  return base+zeroPad(maxNum+1,4);
}

// ── AUDIT LOG ─────────────────────────────────────────────────────────────────
function logAudit(username, action, department, recordId) {
  try {
    var s=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('AUDIT_LOGS');
    if(s) s.appendRow([getTimestamp(),username||'SYSTEM',action,department,recordId]);
  } catch(e){}
}

function logTransactionHistory(trackingNo, prNo, action, fromDept, toDept, username, remarks) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('TRANSACTION_HISTORY');
    if (!sheet) {
      sheet = ss.insertSheet('TRANSACTION_HISTORY');
      sheet.appendRow(['TIMESTAMP','TRACKING NO.','P.R #','ACTION','FROM','TO','USER','REMARKS']);
    }
    sheet.appendRow([getTimestamp(), trackingNo||'', prNo||'', action||'', fromDept||'', toDept||'', username||'SYSTEM', remarks||'']);
  } catch(e){}
}

function getTransactionHistory(trackingNo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('TRANSACTION_HISTORY');
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim(); });
    var tnIdx = hdrs.indexOf('TRACKING NO.');
    var q = String(trackingNo||'').trim();
    if (!q || tnIdx === -1) return [];
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][tnIdx]||'').trim() !== q) continue;
      var obj = {};
      hdrs.forEach(function(h, j) { if (h) obj[h] = formatCellForClient(h, data[i][j]); });
      out.push(obj);
    }
    return out;
  } catch(e){ return []; }
}

// ── SHEET → OBJECTS ───────────────────────────────────────────────────────────
function sheetToObjects(sheetName) {
  var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  var ws=ss.getSheetByName(sheetName);
  if (!ws) return [];
  var data=ws.getDataRange().getValues();
  if (data.length<2) return [];
  var headers=data[0].map(function(h){return String(h||'').trim();});
  var rows=[];
  for (var i=1;i<data.length;i++){
    var row=data[i];
    if (row.every(function(c){return c===''||c===null||c===undefined;})) continue;
    var obj={_rowNum:i+1};
    for (var j=0;j<headers.length;j++){
      if (!headers[j]) continue;
      var val=row[j];
      obj[headers[j]]=formatCellForClient(headers[j], val);
    }
    rows.push(obj);
  }
  return rows;
}

// ── AUTH ──────────────────────────────────────────────────────────────────────
function loginUser(username, password) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('USERS');
    if (!sheet) return {success:false,message:'USERS sheet not found.'};
    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var uCol=hdrs.indexOf('USERNAME'),pCol=hdrs.indexOf('PASSWORD'),
        rCol=hdrs.indexOf('ROLE'),eCol=hdrs.indexOf('END USER');
    if (uCol===-1||pCol===-1||rCol===-1)
      return {success:false,message:'USERS sheet missing USERNAME/PASSWORD/ROLE columns.'};
    for (var i=1;i<data.length;i++){
      var uname=String(data[i][uCol]||'').trim();
      var pwd  =String(data[i][pCol]||'').trim();
      var role =String(data[i][rCol]||'').trim().toUpperCase();
      var isEU =eCol!==-1?String(data[i][eCol]||'').trim().toUpperCase()==='TRUE':false;
      if (uname.toLowerCase()===username.toLowerCase()&&pwd===password){
        logAudit(uname,'LOGIN',role||'—','—');
        return {success:true,username:uname,role:role,isEndUser:isEU};
      }
    }
    return {success:false,message:'Invalid username or password.'};
  } catch(e){return {success:false,message:'Login error: '+e.message};}
}

// ── DATA GETTERS ──────────────────────────────────────────────────────────────
function getBACData()        { return sheetToObjects('BAC'); }
function getSupplyData()     { return sheetToObjects('SUPPLY AND PROPERTY'); }
function getBudgetData()     { return sheetToObjects('BUDGET'); }
function getAccountingData() { return sheetToObjects('ACCOUNTING'); }
function getCashData()       { return sheetToObjects('CASH'); }
function getEndUserData()    { return sheetToObjects('END USER'); }

function getDepartmentPageCounts() {
  function openCount(rows) {
    return (rows || []).filter(function(r){
      return isOpenWorkflowRecord(r);
    }).length;
  }

  return {
    bac: openCount(getBACData()),
    supply: openCount(getSupplyData()),
    budget: openCount(getBudgetData()),
    accounting: openCount(getAccountingData()),
    cash: openCount(getCashData())
  };
}

function getMyBACData(role) {
  var roleLower=String(role||'').toLowerCase();
  return getBACData().filter(function(r){
    return String(r['END-USER / REQUESTED BY']||'').toLowerCase().indexOf(roleLower)!==-1;
  });
}

function getRequestingOffices() {
  try {
    var seen={},out=[];
    getBACData().forEach(function(r){
      var v=String(r['END-USER / REQUESTED BY']||'').trim();
      if(v&&!seen[v]){seen[v]=true;out.push(v);}
    });
    return out.sort();
  } catch(e){return [];}
}

function getSuppliers() {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('SUPPLIERS');
    if (!sheet) return [];
    ensureSuppliersSheet(sheet);
    var data=sheet.getDataRange().getValues();
    if (data.length<2) return [];
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var col=hdrs.indexOf('SUPPLIER NAME');
    var statusCol=hdrs.indexOf('STATUS');
    if (col===-1) col=1;
    var seen={},out=[];
    for (var i=1;i<data.length;i++){
      var status = statusCol===-1 ? 'ACTIVE' : normalizeSupplierStatus(data[i][statusCol]);
      if (status !== 'ACTIVE') continue;
      var v=String(data[i][col]||'').trim();
      if(v&&!seen[v]){seen[v]=true;out.push(v);}
    }
    return out.sort();
  } catch(e){return [];}
}

// ── DASHBOARD STATS ───────────────────────────────────────────────────────────
function getDashboardStats() {
  var bac=getBACData(),sup=getSupplyData(),bud=getBudgetData(),
      acc=getAccountingData(),cash=getCashData();
  function isPending(r){
    return (!r['COMPLETED AT']||r['COMPLETED AT'].trim()==='')&&
           String(r['STATUS']||'').trim().toUpperCase()!=='CANCELLED';
  }
  var allTN={};
  bac.forEach(function(r){if(r['TRACKING NO.'])allTN[r['TRACKING NO.']]=true;});
  return {
    total:        Object.keys(allTN).length,
    atBAC:        bac.filter(isPending).length,
    atSupply:     sup.filter(isPending).length,
    atBudget:     bud.filter(isPending).length,
    atAccounting: acc.filter(isPending).length,
    atCash:       cash.filter(isPending).length,
    completed:    cash.filter(function(r){return String(r['STATUS']||'').trim().toUpperCase()==='PAID';}).length
  };
}

function getMyDashboardStats(role) {
  var myBac=getMyBACData(role);
  function isPending(r){
    return (!r['COMPLETED AT']||r['COMPLETED AT'].trim()==='')&&
           String(r['STATUS']||'').trim().toUpperCase()!=='CANCELLED';
  }
  var myTNs={};
  myBac.forEach(function(r){if(r['TRACKING NO.'])myTNs[r['TRACKING NO.']]=true;});
  function filterTN(rows){return rows.filter(function(r){return myTNs[r['TRACKING NO.']];});}
  var myCash=filterTN(getCashData());
  return {
    total:        Object.keys(myTNs).length,
    atBAC:        myBac.filter(isPending).length,
    atSupply:     filterTN(getSupplyData()).filter(isPending).length,
    atBudget:     filterTN(getBudgetData()).filter(isPending).length,
    atAccounting: filterTN(getAccountingData()).filter(isPending).length,
    atCash:       myCash.filter(isPending).length,
    completed:    myCash.filter(function(r){return String(r['STATUS']||'').trim().toUpperCase()==='PAID';}).length
  };
}

// ── GET MY TRANSACTIONS (v8: Supply sheet as primary END USER source) ─────────
function getMyTransactions(role) {
  try {
    var SECTIONS=['BAC','SUPPLY AND PROPERTY','BUDGET','ACCOUNTING','CASH'];
    var roleLower=String(role||'').toLowerCase();

    // PRIMARY: identify transactions from Supply sheet END USER column
    var myTNs={};
    sheetToObjects('SUPPLY AND PROPERTY').forEach(function(r){
      var eu=String(r['END USER']||'').toLowerCase();
      if(eu.indexOf(roleLower)!==-1&&r['TRACKING NO.'])
        myTNs[r['TRACKING NO.']]=true;
    });

    // FALLBACK: BAC sheet for transactions not yet forwarded to Supply
    sheetToObjects('BAC').forEach(function(r){
      var eu=String(r['END-USER / REQUESTED BY']||'').toLowerCase();
      if(eu.indexOf(roleLower)!==-1&&r['TRACKING NO.'])
        myTNs[r['TRACKING NO.']]=true;
    });

    var bySheet={};
    SECTIONS.forEach(function(sname){
      bySheet[sname]={};
      sheetToObjects(sname).forEach(function(r){
        var tn=r['TRACKING NO.'];
        if(tn&&myTNs[tn]){
          if(!bySheet[sname][tn]) bySheet[sname][tn]=[];
          bySheet[sname][tn].push(r);
        }
      });
    });

    var results=[];
    Object.keys(myTNs).forEach(function(tn){
      var timeline=SECTIONS.map(function(sname){
        var rows=bySheet[sname][tn]||[];
        if(!rows.length) return {section:sname,reached:false,status:'',completedAt:''};
        var latest=rows[rows.length-1];
        return {
          section:sname,reached:true,
          completed:!!(latest['COMPLETED AT']&&latest['COMPLETED AT'].trim()!==''),
          status:String(latest['STATUS']||'').trim(),
          completedAt:latest['COMPLETED AT']||'',
          returnedDate:latest['RETURNED DATE']||'',
          returnReceivedDate:latest['RETURN RECEIVED DATE']||'',
          returnedFrom:latest['RETURNED FROM']||'',
          returnTo:latest['RETURN TO']||'',
          returnRemarks:latest['RETURN REMARKS']||'',
          returnReceivedRemarks:latest['RETURN RECEIVED REMARKS']||''
        };
      });

      var currentSection='BAC';
      var currentStatus='';
      for(var i=0;i<timeline.length;i++){
        var pending = timeline[i].returnedDate && !timeline[i].returnReceivedDate;
        if(pending){
          currentSection=timeline[i].section;
          currentStatus='RETURNED';
          break;
        }
      }
      if(!currentStatus){
        for(var j=0;j<timeline.length;j++){
          if(timeline[j].reached && !timeline[j].completed){
            currentSection=timeline[j].section;
            currentStatus=timeline[j].status;
            break;
          }
        }
      }
      if(!currentStatus){
        for(var k=timeline.length-1;k>=0;k--){
          if(timeline[k].reached){currentSection=timeline[k].section;currentStatus=timeline[k].status;break;}
        }
      }

      var bacRows=bySheet['BAC'][tn]||[];
      var supRows=bySheet['SUPPLY AND PROPERTY'][tn]||[];
      var firstBac=bacRows.length?bacRows[0]:{};
      var firstSup=supRows.length?supRows[0]:{};
      var hasSupply=supRows.length>0;

      results.push({
        trackingNo:                tn,
        prNo:                      firstBac['P.R #']||'',
        office:                    firstSup['END USER']||firstBac['END-USER / REQUESTED BY']||'',
        currentSection:            currentSection,
        currentStatus:             currentStatus,
        timeline:                  timeline,
        hasSupply:                 hasSupply,
        // Supply row reference — null if not yet at Supply
        supRowNum:                 hasSupply?(firstSup._rowNum||null):null,
        // Supply date fields (blank until transaction reaches Supply)
        dateTransmittedToSupplier: firstSup['DATE TRANSMITTED TO SUPPLIER']||'',
        dateReceivedBySupplier:    firstSup['DATE RECEIVED BY SUPPLIER']||'',
        coaDateReceived:           firstSup['COA DATE RECEIVED']||''
      });
    });

    results.sort(function(a,b){return a.trackingNo<b.trackingNo?1:-1;});
    return results;
  } catch(e){return {error:e.message};}
}

function getMyRequestDepartmentBundle(role) {
  try {
    var requests = getMyTransactions(role);
    if (requests && requests.error) return requests;

    var trackingSet = {};
    (requests || []).forEach(function(r) {
      if (r && r.trackingNo) trackingSet[r.trackingNo] = true;
    });

    function filterByTracking(rows) {
      return rows.filter(function(r) {
        return r['TRACKING NO.'] && trackingSet[r['TRACKING NO.']];
      });
    }

    return {
      requests: requests || [],
      departmentRows: {
        bac:        filterByTracking(getBACData()),
        supply:     filterByTracking(getSupplyData()),
        budget:     filterByTracking(getBudgetData()),
        accounting: filterByTracking(getAccountingData()),
        cash:       filterByTracking(getCashData())
      }
    };
  } catch (e) {
    return {error:e.message};
  }
}

// ── UPDATE END USER SUPPLY FIELDS (end-user only: 3 date fields on Supply row) ──
function updateEndUserSupplyFields(rowNum, data, username) {
  try {
    var ALLOWED = ['DATE TRANSMITTED TO SUPPLIER','DATE RECEIVED BY SUPPLIER','COA DATE RECEIVED'];
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('SUPPLY AND PROPERTY');
    if(!sheet) throw new Error('SUPPLY AND PROPERTY sheet not found.');

    var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                    .map(function(h){ return String(h||'').trim(); });

    var tnIdx = hdrs.indexOf('TRACKING NO.');
    var tn = tnIdx!==-1?String(sheet.getRange(rowNum,tnIdx+1).getValue()||'').trim():'';

    for(var c=0;c<hdrs.length;c++){
      var h=hdrs[c];
      if(ALLOWED.indexOf(h)!==-1&&data.hasOwnProperty(h)&&data[h]!==''){
        sheet.getRange(rowNum,c+1).setValue(data[h]);
      }
    }

    logAudit(username||'SYSTEM','END USER EDIT','SUPPLY AND PROPERTY','Row '+rowNum+(tn?' · '+tn:''));
    return {success:true};
  } catch(e){ return {success:false,message:e.message}; }
}

// ── CREATE TRANSACTION (BAC/ADMIN only in v8) ─────────────────────────────────
function createTransaction(data, username) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!data['END-USER / REQUESTED BY'] && !data['REQUESTING OFFICE'])
      throw new Error('End User / Requesting Office is required.');

    var endUser = data['END-USER / REQUESTED BY'] || data['REQUESTING OFFICE'] || '';
    var ts=getTimestamp();
    var trackingNo=generateTrackingNo(ss);
    var prNo = String(data['P.R #']||'').trim();

    var bacSheet=ss.getSheetByName('BAC')||ss.insertSheet('BAC');
    ensureIsNew(bacSheet);
  ensureWorkflowColumns(bacSheet);
    var bacHdrs=bacSheet.getRange(1,1,1,bacSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
    var bacId=generateNextId(bacSheet,'BAC');

    var BAC_FILLABLE = [
      'P.R #','END-USER / REQUESTED BY','CATEGORY','PHIGEPS POSTED','APPROVED ABC','RFQ','RFQ DATE OF COMPLETION',
      'ABSTRACT OF AWARD','AWARDED TO','DATE OF AWARD','MODE OF PROCUREMENT',
      'BAC RESO','NOA','NTP','STATUS'
    ];

    bacSheet.appendRow(bacHdrs.map(function(h){
      if(h==='BAC ID')        return bacId;
      if(h==='TRACKING NO.')  return trackingNo;
      if(h==='P.R #')         return prNo;
      if(h==='DATE RECEIVED') return ts;
      if(h==='COMPLETED AT')  return '';
      if(h==='IS NEW')        return 'FALSE';
      if(h==='STATUS')        return data['STATUS']||'PROCESSING';
      if(BAC_FILLABLE.indexOf(h)!==-1 && data[h]!==undefined && data[h]!=='') return data[h];
      return '';
    }));

    logAudit(username||'SYSTEM','CREATE','BAC',bacId+' · '+trackingNo+' · '+endUser);
    logTransactionHistory(trackingNo, prNo, 'CREATED', '', 'BAC', username, '');
    return {success:true, trackingNo:trackingNo, bacId:bacId};
  } catch(e){return {success:false, message:e.message};}
}

function returnTransaction(sheetName, rowNum, returnTo, remarks, username, markCompleted) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sourceName = normalizeDepartmentName(sheetName);
    var targetName = normalizeDepartmentName(returnTo);
    if (!targetName) throw new Error('Return target is required.');
    if (targetName === sourceName) throw new Error('Cannot return to the same department.');

    var sourceSheet = ss.getSheetByName(sourceName);
    if (!sourceSheet) throw new Error('Sheet not found: '+sourceName);
    ensureWorkflowColumns(sourceSheet);

    var sourceState = getRowRecord(sourceSheet, rowNum);
    var hdrs = sourceState.headers;
    var vals = sourceState.values;
    var row = sourceState.record;
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    var prNo = String(row['P.R #'] || '').trim();
    var ts = getTimestamp();
    var targetLabel = displayDepartmentName(targetName);
    var sourceLabel = displayDepartmentName(sourceName);
    var cleanRemarks = String(remarks || '').trim();
    var shouldMarkCompleted = (markCompleted === false || String(markCompleted).toLowerCase() === 'false') ? false : true;

    syncStatusValidationForRow(sourceSheet, sourceName, rowNum, hdrs);
    setRecordValuesByHeader(sourceSheet, rowNum, hdrs, vals, {
      'STATUS':'RETURNED',
      'COMPLETED AT':shouldMarkCompleted ? ts : '',
      'IS NEW':'',
      'RETURN TO':targetLabel,
      'RETURNED FROM':sourceLabel,
      'RETURN REMARKS':cleanRemarks,
      'RETURNED DATE':ts,
      'RETURN RECEIVED DATE':'',
      'RETURN RECEIVED REMARKS':''
    });

    if (targetName === 'END USER') {
      if (trackingNo) updateEndUserStatus(ss, trackingNo, 'RETURNED TO REQUESTING OFFICE');
      logAudit(username||'SYSTEM', 'RETURN → REQUESTING OFFICE', sourceName, trackingNo||('Row '+rowNum));
      logTransactionHistory(trackingNo, prNo, 'RETURNED', displayDepartmentName(sourceName), 'Requesting Office', username, cleanRemarks);
      return {success:true, target:'END USER'};
    }

    var targetSheet = ss.getSheetByName(targetName) || ss.insertSheet(targetName);
    ensureIsNew(targetSheet);
    ensureWorkflowColumns(targetSheet);
    if (targetName === 'SUPPLY AND PROPERTY') ensureParticulars(targetSheet);

    var targetRowNum = findLatestRowNumByIdentity(targetSheet, trackingNo, prNo);
    if (targetRowNum) {
      var targetState = getRowRecord(targetSheet, targetRowNum);
      setRecordValuesByHeader(targetSheet, targetRowNum, targetState.headers, targetState.values, {
        'STATUS':'RETURNED',
        'COMPLETED AT': shouldMarkCompleted ? ts : '',
        'IS NEW': shouldMarkCompleted ? '' : 'TRUE',
        'RETURN TO':targetLabel,
        'RETURNED FROM':sourceLabel,
        'RETURN REMARKS':cleanRemarks,
        'RETURNED DATE':ts,
        'RETURN RECEIVED DATE':'',
        'RETURN RECEIVED REMARKS':''
      });
    } else {
      var targetHdrs = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues()[0].map(function(h){ return String(h||'').trim(); });
      var idCfg = ID_CONFIG[targetName];
      var nextId = generateNextId(targetSheet, targetName);
      var office = row['END-USER / REQUESTED BY'] || row['END USER'] || '';
      var amount = (targetName === 'SUPPLY AND PROPERTY') ? '' : (row['AMOUNT'] || row['APPROVED ABC'] || '');
      targetSheet.appendRow(targetHdrs.map(function(h){
        if (idCfg && h === idCfg.col) return nextId;
        if (h === 'TRACKING NO.') return trackingNo;
        if (h === 'P.R #') return prNo;
        if (h === 'DATE RECEIVED') return ts;
        if (h === 'COMPLETED AT') return shouldMarkCompleted ? ts : '';
        if (h === 'IS NEW') return shouldMarkCompleted ? '' : 'TRUE';
        if (h === 'STATUS') return 'RETURNED';
        if (h === 'END-USER / REQUESTED BY' || h === 'END USER' || h === 'REQUESTING OFFICE') return office;
        if (h === 'AMOUNT' || h === 'APPROVED ABC') return amount;
        if (h === 'RETURN TO') return targetLabel;
        if (h === 'RETURNED FROM') return sourceLabel;
        if (h === 'RETURN REMARKS') return cleanRemarks;
        if (h === 'RETURNED DATE') return ts;
        if (h === 'RETURN RECEIVED DATE') return '';
        if (h === 'RETURN RECEIVED REMARKS') return '';
        return '';
      }));
      targetRowNum = targetSheet.getLastRow();
    }

    if (trackingNo) updateEndUserStatus(ss, trackingNo, 'RETURNED TO '+targetLabel.toUpperCase());
    logAudit(username||'SYSTEM', 'RETURN → '+targetLabel.toUpperCase(), sourceName, trackingNo||('Row '+rowNum));
    logTransactionHistory(trackingNo, prNo, 'RETURNED', displayDepartmentName(sourceName), targetLabel, username, cleanRemarks);
    return {success:true, target:targetName, rowNum:targetRowNum};
  } catch(e) {
    return {success:false, message:e.message};
  }
}

function completeReturnedTransaction(sheetName, rowNum, remarks, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var currentName = normalizeDepartmentName(sheetName);
    var currentSheet = ss.getSheetByName(currentName);
    if (!currentSheet) throw new Error('Sheet not found: '+currentName);
    ensureWorkflowColumns(currentSheet);

    var currentState = getRowRecord(currentSheet, rowNum);
    var currentHdrs = currentState.headers;
    var currentVals = currentState.values;
    var row = currentState.record;
    if (!isReturnPendingRecord(row)) throw new Error('This item is not waiting in the return tab.');
    var cleanRemarks = String(remarks || '').trim();
    if (!cleanRemarks) throw new Error('Completion remarks are required.');

    var ts = getTimestamp();
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    var returnTo = normalizeDepartmentName(row['RETURN TO']);
    var actor = getUserContextByUsername(ss, username);
    if (!canUserCompleteReturnedItem(actor, currentName, returnTo)) {
      throw new Error('You are not allowed to complete returned items for this department.');
    }
    var sourceName = (returnTo === 'END USER') ? currentName : normalizeDepartmentName(row['RETURNED FROM']);
    if (!sourceName) throw new Error('Return source could not be determined.');

    var targetLabel = displayDepartmentName(returnTo === 'END USER' ? 'END USER' : currentName);
    var sourceSheet = ss.getSheetByName(sourceName);
    if (!sourceSheet) throw new Error('Source sheet not found: '+sourceName);
    ensureWorkflowColumns(sourceSheet);

    var sourceRowNum = (returnTo === 'END USER')
      ? rowNum
      : findPendingReturnSourceRow(sourceSheet, trackingNo, currentName);
    if (!sourceRowNum) throw new Error('Source return row not found.');
    var sourceState = getRowRecord(sourceSheet, sourceRowNum);
    var sourceHdrs = sourceState.headers;
    var sourceVals = sourceState.values;
    var activeStatus = sourceName === 'BAC' ? 'PROCESSING' : 'RECEIVED';
    syncStatusValidationForRow(sourceSheet, sourceName, sourceRowNum, sourceHdrs);
    setRecordValuesByHeader(sourceSheet, sourceRowNum, sourceHdrs, sourceVals, {
      'STATUS':activeStatus,
      'COMPLETED AT':'',
      'IS NEW':'',
      'RETURN RECEIVED DATE':ts,
      'RETURN RECEIVED REMARKS':cleanRemarks
    });

    if (!(returnTo === 'END USER')) {
      var targetUpdates = {
        'RETURN RECEIVED DATE':ts,
        'RETURN RECEIVED REMARKS':cleanRemarks,
        'IS NEW':''
      };
      var currentCompletedAt = String(row['COMPLETED AT'] || '').trim();
      if (currentCompletedAt === '') targetUpdates['COMPLETED AT'] = ts;
      setRecordValuesByHeader(currentSheet, rowNum, currentHdrs, currentVals, targetUpdates);
    }

    var euStatus = DEPT_RECEIVE_LABELS[sourceName] || activeStatus;
    if (trackingNo) updateEndUserStatus(ss, trackingNo, euStatus);
    logAudit(username||'SYSTEM', 'RETURN DONE', currentName, trackingNo||('Row '+rowNum));
    logTransactionHistory(trackingNo, '', 'RETURN DONE', displayDepartmentName(currentName), displayDepartmentName(sourceName), username, cleanRemarks);
    return {success:true};
  } catch(e) {
    return {success:false, message:e.message};
  }
}

function receiveAndForwardReturnedTransaction(sheetName, rowNum, forwardTo, forwardRemarks, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var currentName = normalizeDepartmentName(sheetName);
    var targetName = normalizeDepartmentName(forwardTo);
    if (!currentName) throw new Error('Current department is required.');
    if (!targetName) throw new Error('Forward target is required.');
    var currentSheet = ss.getSheetByName(currentName);
    if (!currentSheet) throw new Error('Sheet not found: '+currentName);
    ensureWorkflowColumns(currentSheet);

    var currentState = getRowRecord(currentSheet, rowNum);
    var currentHdrs = currentState.headers;
    var currentVals = currentState.values;
    var row = currentState.record;
    if (!isReturnPendingRecord(row)) throw new Error('This item is not waiting in the return tab.');

    var cleanRemarks = String(forwardRemarks || '').trim();
    if (!cleanRemarks) throw new Error('Forward remarks are required.');

    var returnTo = normalizeDepartmentName(row['RETURN TO']);
    var actor = getUserContextByUsername(ss, username);
    if (!canUserCompleteReturnedItem(actor, currentName, returnTo)) {
      throw new Error('You are not allowed to receive returned items for this department.');
    }

    var ts = getTimestamp();
    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    var prNo = String(row['P.R #'] || '').trim();

    setRecordValuesByHeader(currentSheet, rowNum, currentHdrs, currentVals, {
      'RETURN RECEIVED DATE': ts,
      'RETURN RECEIVED REMARKS': cleanRemarks,
      'FORWARD REMARKS': cleanRemarks
    });

    if (targetName === currentName) {
      var reopenStatus = currentName === 'BAC' ? 'PROCESSING' : 'RECEIVED';
      syncStatusValidationForRow(currentSheet, currentName, rowNum, currentHdrs);
      setRecordValuesByHeader(currentSheet, rowNum, currentHdrs, currentVals, {
        'STATUS': reopenStatus,
        'COMPLETED AT': '',
        'IS NEW': 'TRUE',
        'FORWARDED FROM': 'Requesting Office',
        'FORWARD REMARKS': cleanRemarks,
        'RETURN RECEIVED DATE': ts,
        'RETURN RECEIVED REMARKS': cleanRemarks
      });

      if (trackingNo) updateEndUserStatus(ss, trackingNo, 'FORWARDED TO '+displayDepartmentName(currentName).toUpperCase());
      logAudit(username||'SYSTEM', 'RECEIVE & FORWARD → '+displayDepartmentName(currentName).toUpperCase(), currentName, trackingNo||('Row '+rowNum));
      logTransactionHistory(trackingNo, prNo, 'FORWARDED', 'Requesting Office', displayDepartmentName(currentName), username, cleanRemarks);
      return {success:true, target:currentName};
    }

    if (targetName === 'END USER') {
      syncStatusValidationForRow(currentSheet, currentName, rowNum, currentHdrs);
      setRecordValuesByHeader(currentSheet, rowNum, currentHdrs, currentVals, {
        'STATUS': 'FORWARDED TO REQUESTING OFFICE',
        'COMPLETED AT': ts,
        'IS NEW': '',
        'RETURN RECEIVED DATE': ts,
        'RETURN RECEIVED REMARKS': cleanRemarks,
        'FORWARD REMARKS': cleanRemarks
      });

      if (trackingNo) updateEndUserStatus(ss, trackingNo, 'FORWARDED TO REQUESTING OFFICE');
      logAudit(username||'SYSTEM', 'RECEIVE & FORWARD → REQUESTING OFFICE', currentName, trackingNo||('Row '+rowNum));
      logTransactionHistory(trackingNo, prNo, 'FORWARDED', displayDepartmentName(currentName), 'Requesting Office', username, cleanRemarks);
      return {success:true, target:'END USER'};
    }

    var fwdRes = forwardTransaction(currentName, rowNum, targetName, {'FORWARD REMARKS': cleanRemarks}, username);
    if (!fwdRes || !fwdRes.success) throw new Error((fwdRes && fwdRes.message) || 'Unable to forward transaction.');
    return {success:true, target:targetName};
  } catch(e) {
    return {success:false, message:e.message};
  }
}

// ── UPDATE TRANSACTION ────────────────────────────────────────────────────────
function updateTransaction(sheetName, rowNum, data, username) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName(sheetName);
    if(!sheet) throw new Error('Sheet not found: '+sheetName);
    ensureWorkflowColumns(sheet);

    var hdrs   =sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var rowVals=sheet.getRange(rowNum,1,1,hdrs.length).getValues()[0];

    var isNewIdx=-1;
    for(var i=0;i<hdrs.length;i++){if(String(hdrs[i]).trim()==='IS NEW'){isNewIdx=i;break;}}
    var wasNew=isNewIdx!==-1&&String(rowVals[isNewIdx]||'').trim().toUpperCase()==='TRUE';

    var tnIdx=-1;
    for(var i=0;i<hdrs.length;i++){if(String(hdrs[i]).trim()==='TRACKING NO.'){tnIdx=i;break;}}
    var trackingNo=tnIdx!==-1?String(rowVals[tnIdx]||'').trim():'';

    syncStatusValidationForRow(sheet, sheetName, rowNum, hdrs.map(function(h){ return String(h||'').trim(); }));

    var newVals=rowVals.slice();
    for(var c=0;c<hdrs.length;c++){
      var h=String(hdrs[c]).trim();
      if(!h) continue;
      if(h==='COMPLETED AT' && data.hasOwnProperty(h)) { newVals[c]=data[h]; continue; }
      if(PROTECTED_COLS.indexOf(h)!==-1) continue;
      if(h==='P.R #') continue;
      if(h==='IS NEW')  { newVals[c]=''; continue; }
      if(h==='STATUS'&&wasNew&&sheetName!=='BAC') { newVals[c]='PROCESSING'; continue; }
      if(data.hasOwnProperty(h)) newVals[c]=data[h];
    }
    sheet.getRange(rowNum,1,1,hdrs.length).setValues([newVals]);

    if(wasNew&&trackingNo){
      var euStatus=DEPT_RECEIVE_LABELS[sheetName];
      if(euStatus){
        updateEndUserStatus(ss,trackingNo,euStatus);
        logAudit(username||'SYSTEM',euStatus,'END USER',trackingNo);
      }
    }

    logAudit(username||'SYSTEM','UPDATE',sheetName,'Row '+rowNum+(trackingNo?' · '+trackingNo:''));
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

// ── UPDATE P.R # (ADMIN ONLY) ─────────────────────────────────────────────────
function updatePrNumber(trackingNo, newPrNo, username) {
  try {
    if (!trackingNo) throw new Error('Tracking number is required.');
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetNames = ['BAC', 'SUPPLY AND PROPERTY', 'BUDGET', 'ACCOUNTING', 'CASH'];
    var updatedCount = 0;
    sheetNames.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;
      var data = sheet.getDataRange().getValues();
      var hdrs = data[0].map(function(h) { return String(h || '').trim(); });
      var tnIdx = hdrs.indexOf('TRACKING NO.');
      var prIdx = hdrs.indexOf('P.R #');
      if (tnIdx === -1 || prIdx === -1) return;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][tnIdx] || '').trim() === trackingNo) {
          sheet.getRange(i + 1, prIdx + 1).setValue(newPrNo);
          updatedCount++;
        }
      }
    });
    logAudit(username || 'SYSTEM', 'UPDATE P.R #', 'ALL', trackingNo + ' → ' + (newPrNo || '(cleared)'));
    return { success: true, updatedCount: updatedCount };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── FORWARD CONFIG ────────────────────────────────────────────────────────────
var FORWARD_CONFIG = {
  'BAC': {
    forwardStatus: 'FORWARDED TO SUPPLY',
    nextSheet:     'SUPPLY AND PROPERTY',
    colMap: { 'END USER':'END-USER / REQUESTED BY', 'P.R #':'P.R #' }
  },
  'SUPPLY AND PROPERTY': {
    forwardStatus: 'FORWARDED TO BUDGET',
    nextSheet:     'BUDGET',
    colMap: {}
  },
  'BUDGET': {
    forwardStatus: 'FORWARDED TO ACCOUNTING',
    nextSheet:     'ACCOUNTING',
    colMap: {}
  },
  'ACCOUNTING': {
    forwardStatus: 'FORWARDED TO CASHIER',
    nextSheet:     'CASH',
    colMap: {}
  }
};

// ── FORWARD TRANSACTION ───────────────────────────────────────────────────────
function forwardTransaction(sheetName, rowNum, targetSheetName, extraData, username) {
  try {
    var ss    =SpreadsheetApp.openById(SPREADSHEET_ID);
    var sourceName = normalizeDepartmentName(sheetName);
    var sheet =ss.getSheetByName(sourceName);
    if(!sheet) throw new Error('Sheet not found: '+sourceName);
    ensureWorkflowColumns(sheet);
    var config=FORWARD_CONFIG[sourceName];
    if(!config) throw new Error('No forwarding config for: '+sourceName);
    var targetName = normalizeDepartmentName(targetSheetName || config.nextSheet);
    if(!targetName) throw new Error('Forward target is required.');
    if(targetName === sourceName) throw new Error('Cannot forward to the same department.');
    if(targetName === 'END USER') throw new Error('Use Return to send to Requesting Office.');

    var allData=sheet.getDataRange().getValues();
    var hdrs   =allData[0].map(function(h){return String(h||'').trim();});
    var rowData=allData[rowNum-1];
    var src={};
    for(var j=0;j<hdrs.length;j++){
      if(!hdrs[j]) continue;
      var v=rowData[j];
      src[hdrs[j]]=formatCellForClient(hdrs[j], v);
    }

    var ts=getTimestamp();
    var trackingNo=src['TRACKING NO.']||'';
    var targetLabel = displayDepartmentName(targetName);
    var dynamicForwardStatus = 'FORWARDED TO ' + targetLabel.toUpperCase();

    var sourceStatusOptions = (DEPT_STATUS_OPTIONS[sourceName] || []).slice();
    if (sourceStatusOptions.indexOf(dynamicForwardStatus) === -1) sourceStatusOptions.push(dynamicForwardStatus);
    syncStatusValidationForRow(sheet, sourceName, rowNum, hdrs, sourceStatusOptions);

    var statusCol = hdrs.indexOf('STATUS');
    var completedCol = hdrs.indexOf('COMPLETED AT');
    var isNewCol = hdrs.indexOf('IS NEW');
    var forwardRemarksCol = hdrs.indexOf('FORWARD REMARKS');
    if (statusCol !== -1) {
      var statusCell = sheet.getRange(rowNum, statusCol + 1);
      statusCell.setValue(dynamicForwardStatus);
    }
    if (completedCol !== -1) sheet.getRange(rowNum, completedCol + 1).setValue(ts);
    if (isNewCol !== -1) sheet.getRange(rowNum, isNewCol + 1).setValue('');
    if (forwardRemarksCol !== -1) sheet.getRange(rowNum, forwardRemarksCol + 1).setValue(String((extraData && extraData['FORWARD REMARKS']) || '').trim());

    var nextSheet=ss.getSheetByName(targetName)||ss.insertSheet(targetName);
    ensureIsNew(nextSheet);
    ensureWorkflowColumns(nextSheet);
    var nextHdrs=nextSheet.getRange(1,1,1,nextSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});

    var targetRowNum = findLatestRowNumByIdentity(nextSheet, trackingNo, src['P.R #']||'');
    if (targetRowNum) {
      var targetState = getRowRecord(nextSheet, targetRowNum);
      setRecordValuesByHeader(nextSheet, targetRowNum, targetState.headers, targetState.values, {
        'DATE RECEIVED': ts,
        'COMPLETED AT': '',
        'STATUS': (extraData&&extraData['STATUS'])?extraData['STATUS']:'RECEIVED',
        'IS NEW': 'TRUE',
        'FORWARDED FROM': displayDepartmentName(sourceName),
        'FORWARD REMARKS': String((extraData&&extraData['FORWARD REMARKS'])||'').trim(),
        'RETURN TO': '',
        'RETURNED FROM': '',
        'RETURN REMARKS': '',
        'RETURNED DATE': '',
        'RETURN RECEIVED DATE': '',
        'RETURN RECEIVED REMARKS': ''
      });
    } else {
      var nextId  =generateNextId(nextSheet,targetName);
      nextSheet.appendRow(nextHdrs.map(function(h){
        var idCfg=ID_CONFIG[targetName];
        if(idCfg&&h===idCfg.col)   return nextId;
        if(h==='DATE RECEIVED')     return ts;
        if(h==='COMPLETED AT')      return '';
        if(h==='STATUS')            return (extraData&&extraData['STATUS'])?extraData['STATUS']:'RECEIVED';
        if(h==='IS NEW')            return 'TRUE';
        if(h==='FORWARDED FROM')    return displayDepartmentName(sourceName);
        if(h==='FORWARD REMARKS')   return String((extraData&&extraData['FORWARD REMARKS'])||'').trim();
        if(h==='RETURN TO')         return '';
        if(h==='RETURNED FROM')     return '';
        if(h==='RETURN REMARKS')    return '';
        if(h==='RETURNED DATE')     return '';
        if(h==='RETURN RECEIVED DATE') return '';
        if(h==='RETURN RECEIVED REMARKS') return '';
        if(extraData&&extraData[h]!==undefined&&extraData[h]!=='') return extraData[h];
        var mapped=config.colMap[h];
        if(mapped&&src[mapped]!==undefined&&src[mapped]!=='') return src[mapped];
        if(src[h]!==undefined&&src[h]!=='') return src[h];
        return '';
      }));
    }

    if(trackingNo){
      var euStatus=dynamicForwardStatus;
      if(euStatus){
        updateEndUserStatus(ss,trackingNo,euStatus);
        logAudit(username||'SYSTEM',euStatus,'END USER',trackingNo);
      }
    }

    logAudit(username||'SYSTEM','FORWARD → '+targetName,sourceName,trackingNo||'Row '+rowNum);
    logTransactionHistory(trackingNo, src['P.R #']||'', 'FORWARDED', displayDepartmentName(sourceName), displayDepartmentName(targetName), username, String((extraData&&extraData['FORWARD REMARKS'])||'').trim());
    return {success:true,trackingNo:trackingNo};
  } catch(e){return {success:false,message:e.message};}
}

// ── MARK PAID ─────────────────────────────────────────────────────────────────
function markCompleted(rowNum, username) {
  try {
    var ss   =SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('CASH');
    if(!sheet) throw new Error('CASH sheet not found.');
    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim();});
    var tnIdx=hdrs.indexOf('TRACKING NO.');
    var tn=tnIdx!==-1?String(data[rowNum-1][tnIdx]||''):'';
    var ts=getTimestamp();

    syncStatusValidationForRow(sheet, 'CASH', rowNum, hdrs, ['RECEIVED','LACKING OF ATTACHMENTS','PAID','UNPAID','CANCELLED']);

    for(var c=0;c<hdrs.length;c++){
      var h=hdrs[c];
      if(h==='STATUS')       sheet.getRange(rowNum,c+1).setValue('PAID');
      if(h==='COMPLETED AT') sheet.getRange(rowNum,c+1).setValue(ts);
      if(h==='IS NEW')       sheet.getRange(rowNum,c+1).setValue('');
    }

    if(tn) updateEndUserStatus(ss,tn,'PAID');
    logAudit(username||'SYSTEM','MARK PAID','CASH',tn||'Row '+rowNum);
    logTransactionHistory(tn, '', 'PAID', 'Cashier', '', username, '');
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

// ── CANCEL TRANSACTION ────────────────────────────────────────────────────────
function cancelTransaction(sheetName, rowNum, username) {
  try {
    var ss   =SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName(sheetName);
    if(!sheet) throw new Error('Sheet not found: '+sheetName);
    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim();});
    var tnIdx=hdrs.indexOf('TRACKING NO.');
    var tn=tnIdx!==-1?String(data[rowNum-1][tnIdx]||''):'';
    var ts=getTimestamp();

    syncStatusValidationForRow(sheet, sheetName, rowNum, hdrs);

    for(var c=0;c<hdrs.length;c++){
      var h=String(hdrs[c]).trim();
      if(h==='STATUS')       sheet.getRange(rowNum,c+1).setValue('CANCELLED');
      if(h==='COMPLETED AT') sheet.getRange(rowNum,c+1).setValue(ts);
      if(h==='IS NEW')       sheet.getRange(rowNum,c+1).setValue('');
    }

    if(tn) updateEndUserStatus(ss,tn,'CANCELLED');
    logAudit(username||'SYSTEM','CANCEL',sheetName,tn||'Row '+rowNum);
    logTransactionHistory(tn, '', 'CANCELLED', displayDepartmentName(normalizeDepartmentName(sheetName)), '', username, '');
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

// ── SEARCH TRANSACTION ────────────────────────────────────────────────────────
function searchTransaction(query) {
  try {
    var SECTIONS=['BAC','SUPPLY AND PROPERTY','BUDGET','ACCOUNTING','CASH'];
    var q=String(query||'').trim().toLowerCase();
    if(!q) return null;

    var found=null,bySheet={};
    SECTIONS.forEach(function(sname){
      sheetToObjects(sname).forEach(function(r){
        var tn=String(r['TRACKING NO.']||'').toLowerCase();
        var pr=String(r['P.R #']||'').toLowerCase();
        if((tn&&(tn===q||tn.indexOf(q)!==-1))||(pr&&(pr===q||pr.indexOf(q)!==-1))){
          if(!found) found={
            trackingNo:r['TRACKING NO.']||'',
            prNo:r['P.R #']||'',
            office:r['END-USER / REQUESTED BY']||r['END USER']||'',
            item:r['PARTICULARS']||''
          };
          if(!bySheet[sname]) bySheet[sname]=[];
          bySheet[sname].push(r);
        }
      });
    });

    if(!found) return null;

    var timeline=SECTIONS.map(function(sname){
      var rows=bySheet[sname]||[];
      if(!rows.length) return {section:sname,reached:false,status:'',completedAt:''};
      var latest=rows[rows.length-1];
      return {
        section:sname,reached:true,
        completed:!!(latest['COMPLETED AT']&&latest['COMPLETED AT'].trim()!==''),
        status:String(latest['STATUS']||'').trim(),
        completedAt:latest['COMPLETED AT']||'',
        returnedDate:latest['RETURNED DATE']||'',
        returnReceivedDate:latest['RETURN RECEIVED DATE']||'',
        returnedFrom:latest['RETURNED FROM']||'',
        returnRemarks:latest['RETURN REMARKS']||'',
        returnReceivedRemarks:latest['RETURN RECEIVED REMARKS']||''
      };
    });

    var currentSection='BAC',currentStatus='';
    for(var i=0;i<timeline.length;i++){
      var pending = timeline[i].returnedDate && !timeline[i].returnReceivedDate;
      if(pending){ currentSection=timeline[i].section; currentStatus='RETURNED'; break; }
    }
    if(!currentStatus){
      for(var j=0;j<timeline.length;j++){
        if(timeline[j].reached && !timeline[j].completed){ currentSection=timeline[j].section; currentStatus=timeline[j].status; break; }
      }
    }
    if(!currentStatus){
      for(var k=timeline.length-1;k>=0;k--){
        if(timeline[k].reached){currentSection=timeline[k].section;currentStatus=timeline[k].status;break;}
      }
    }

    return {
      trackingNo: found.trackingNo,
      prNo:       found.prNo,
      office:     found.office,
      item:       found.item,
      currentSection: currentSection,
      currentStatus:  currentStatus,
      currentReturnRemarks: (timeline.find(function(t){ return t.section===currentSection; })||{}).returnRemarks || '',
      currentReturnReceivedRemarks: (timeline.find(function(t){ return t.section===currentSection; })||{}).returnReceivedRemarks || '',
      timeline:   timeline,
      history:    getTransactionHistory(found.trackingNo)
    };
  } catch(e){return {error:e.message};}
}

function getSourceDeptRowByTracking(trackingNo, sourceDeptName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var normalized = normalizeDepartmentName(sourceDeptName);
    var sheet = ss.getSheetByName(normalized);
    if (!sheet) return null;
    var rowNum = findLatestRowNumByTracking(sheet, trackingNo);
    if (!rowNum) return null;
    return getRowRecord(sheet, rowNum).record;
  } catch(e) { return null; }
}

// ════════════════════════════════════════════════════════════════════════════════
//  USER MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════════
function getUsers() {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('USERS');
    if(!sheet) return [];
    var data=sheet.getDataRange().getValues();
    if(data.length<2) return [];
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var uCol=hdrs.indexOf('USERNAME'),rCol=hdrs.indexOf('ROLE'),
        eCol=hdrs.indexOf('END USER'),idCol=hdrs.indexOf('USER ID'),aCol=hdrs.indexOf('ACTIVE');
    var rows=[];
    for(var i=1;i<data.length;i++){
      var row=data[i];
      if(!row[uCol]&&!row[rCol]) continue;
      rows.push({
        _rowNum:i+1,
        userId:  idCol!==-1?String(row[idCol]||''):'',
        username:String(row[uCol]||''),
        role:    String(row[rCol]||'').toUpperCase(),
        isEndUser:eCol!==-1?String(row[eCol]||'').trim().toUpperCase()==='TRUE':false,
        isActive:aCol!==-1?(String(row[aCol]||'').trim().toUpperCase()!=='FALSE'&&String(row[aCol]||'').trim()!==''):true
      });
    }
    return rows;
  } catch(e){return [];}
}

function getAllSuppliers() {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('SUPPLIERS');
    if(!sheet) return [];
    ensureSuppliersSheet(sheet);
    var data=sheet.getDataRange().getValues();
    if(data.length<2) return [];
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var idCol=hdrs.indexOf('SUPPLIER ID');
    var nameCol=hdrs.indexOf('SUPPLIER NAME');
    var statusCol=hdrs.indexOf('STATUS');
    var rows=[];
    for(var i=1;i<data.length;i++){
      var row=data[i];
      var supplierName=String(nameCol!==-1?row[nameCol]:'').trim();
      if(!supplierName) continue;
      rows.push({
        _rowNum:i+1,
        supplierId:idCol!==-1?String(row[idCol]||'').trim():'',
        supplierName:supplierName,
        status:normalizeSupplierStatus(statusCol!==-1?row[statusCol]:'ACTIVE')
      });
    }
    return rows;
  } catch(e){return [];}
}

function addUser(userData, adminUsername) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('USERS');
    if(!sheet) throw new Error('USERS sheet not found.');
    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var uCol=hdrs.indexOf('USERNAME'),idCol=hdrs.indexOf('USER ID');

    for(var i=1;i<data.length;i++){
      if(String(data[i][uCol]||'').trim().toLowerCase()===userData.username.toLowerCase())
        return {success:false,message:'Username "'+userData.username+'" already exists.'};
    }

    var maxNum=0;
    if(idCol!==-1){
      for(var j=1;j<data.length;j++){
        var uid=String(data[j][idCol]||'');
        if(uid.indexOf('USER')===0){
          var n=parseInt(uid.substring(4))||0;
          if(n>maxNum) maxNum=n;
        }
      }
    }
    var newId='USER'+zeroPad(maxNum+1,4);
    var euRolesUp=END_USER_ROLES.map(function(r){return r.toUpperCase();});
    var roleUpper=String(userData.role||'').toUpperCase();
    var isEU=userData.isEndUser===true||userData.isEndUser==='true'||euRolesUp.indexOf(roleUpper)!==-1;

    sheet.appendRow(hdrs.map(function(h){
      if(h==='USER ID')  return newId;
      if(h==='USERNAME') return userData.username||'';
      if(h==='PASSWORD') return userData.password||'1234';
      if(h==='ROLE')     return roleUpper;
      if(h==='END USER') return isEU?'TRUE':'';
      if(h==='ACTIVE')   return userData.isActive!==false?'TRUE':'';
      return '';
    }));
    logAudit(adminUsername||'ADMIN','ADD USER','USERS',newId+' · '+userData.username+' · '+roleUpper);
    return {success:true,userId:newId};
  } catch(e){return {success:false,message:e.message};}
}

function updateUser(rowNum, userData, adminUsername) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('USERS');
    if(!sheet) throw new Error('USERS sheet not found.');
    var hdrs=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var euRolesUp=END_USER_ROLES.map(function(r){return r.toUpperCase();});
    var roleUpper=String(userData.role||'').toUpperCase();
    var isEU=userData.isEndUser===true||userData.isEndUser==='true'||euRolesUp.indexOf(roleUpper)!==-1;

    for(var c=0;c<hdrs.length;c++){
      var h=hdrs[c];
      if(h==='USER ID')  continue;
      if(h==='USERNAME'  &&userData.username!==undefined)                sheet.getRange(rowNum,c+1).setValue(userData.username);
      if(h==='PASSWORD'  &&userData.password&&userData.password!=='')    sheet.getRange(rowNum,c+1).setValue(userData.password);
      if(h==='ROLE'      &&userData.role!==undefined)                    sheet.getRange(rowNum,c+1).setValue(roleUpper);
      if(h==='END USER')                                                  sheet.getRange(rowNum,c+1).setValue(isEU?'TRUE':'');
      if(h==='ACTIVE'    &&userData.isActive!==undefined)                sheet.getRange(rowNum,c+1).setValue(userData.isActive?'TRUE':'');
    }
    logAudit(adminUsername||'ADMIN','UPDATE USER','USERS','Row '+rowNum+' · '+userData.username);
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

function deleteUser(rowNum, targetUsername, adminUsername) {
  try {
    if(targetUsername.toLowerCase()===adminUsername.toLowerCase())
      return {success:false,message:'You cannot delete your own account.'};
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('USERS');
    if(!sheet) throw new Error('USERS sheet not found.');
    sheet.deleteRow(rowNum);
    logAudit(adminUsername||'ADMIN','DELETE USER','USERS',targetUsername);
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

function updateMyCredentials(currentUsername, currentPassword, profileData) {
  try {
    var cleanCurrentUsername = String(currentUsername || '').trim();
    var cleanCurrentPassword = String(currentPassword || '').trim();
    var requestedUsername = String((profileData && profileData.username) || '').trim();
    var requestedNewPassword = String((profileData && profileData.newPassword) || '').trim();
    var wantsPasswordChange = requestedNewPassword !== '';

    if (!cleanCurrentUsername) return {success:false,message:'Current user is required.'};
    if (!requestedUsername) return {success:false,message:'Username is required.'};
    if (wantsPasswordChange && !cleanCurrentPassword) return {success:false,message:'Current password is required.'};
    if (wantsPasswordChange && requestedNewPassword.length < 4) {
      return {success:false,message:'New password must be at least 4 characters.'};
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('USERS');
    if (!sheet) return {success:false,message:'USERS sheet not found.'};

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return {success:false,message:'No user records found.'};

    var hdrs = data[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
    var uCol = hdrs.indexOf('USERNAME');
    var pCol = hdrs.indexOf('PASSWORD');
    var rCol = hdrs.indexOf('ROLE');
    var eCol = hdrs.indexOf('END USER');
    if (uCol === -1 || pCol === -1 || rCol === -1) {
      return {success:false,message:'USERS sheet missing USERNAME/PASSWORD/ROLE columns.'};
    }

    var targetRowNum = -1;
    var role = '';
    var isEU = false;

    for (var i=1; i<data.length; i++) {
      var uname = String(data[i][uCol] || '').trim();
      if (uname.toLowerCase() !== cleanCurrentUsername.toLowerCase()) continue;

      var pwd = String(data[i][pCol] || '').trim();
      if (wantsPasswordChange && pwd !== cleanCurrentPassword) {
        return {success:false,message:'Current password is incorrect.'};
      }

      targetRowNum = i + 1;
      role = String(data[i][rCol] || '').trim().toUpperCase();
      isEU = eCol !== -1 ? String(data[i][eCol] || '').trim().toUpperCase() === 'TRUE' : false;
      break;
    }

    if (targetRowNum === -1) {
      return {success:false,message:'User account was not found.'};
    }

    for (var j=1; j<data.length; j++) {
      if (j + 1 === targetRowNum) continue;
      var otherUsername = String(data[j][uCol] || '').trim();
      if (otherUsername && otherUsername.toLowerCase() === requestedUsername.toLowerCase()) {
        return {success:false,message:'Username "'+requestedUsername+'" already exists.'};
      }
    }

    sheet.getRange(targetRowNum, uCol + 1).setValue(requestedUsername);
    if (requestedNewPassword) {
      sheet.getRange(targetRowNum, pCol + 1).setValue(requestedNewPassword);
    }

    var auditDetail = 'Row '+targetRowNum+' · '+requestedUsername+(requestedNewPassword ? ' · PASSWORD UPDATED' : '');
    logAudit(requestedUsername, 'UPDATE MY PROFILE', 'USERS', auditDetail);

    return {
      success:true,
      username:requestedUsername,
      role:role,
      isEndUser:isEU,
      message:'Profile updated successfully.'
    };
  } catch(e) {
    return {success:false,message:e.message};
  }
}

function addSupplier(supplierData, adminUsername) {
  try {
    var name=String((supplierData&&supplierData.supplierName)||'').trim();
    if(!name) return {success:false,message:'Supplier name is required.'};

    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('SUPPLIERS') || ss.insertSheet('SUPPLIERS');
    ensureSuppliersSheet(sheet);

    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var idCol=hdrs.indexOf('SUPPLIER ID');
    var nameCol=hdrs.indexOf('SUPPLIER NAME');
    var statusCol=hdrs.indexOf('STATUS');

    for(var i=1;i<data.length;i++){
      if(String(data[i][nameCol]||'').trim().toLowerCase()===name.toLowerCase())
        return {success:false,message:'Supplier "'+name+'" already exists.'};
    }

    var maxNum=0;
    for(var j=1;j<data.length;j++){
      var supplierId=String(idCol!==-1?data[j][idCol]:'').trim();
      var match=supplierId.match(/(\d+)$/);
      var n=match?parseInt(match[1],10):0;
      if(n>maxNum) maxNum=n;
    }
    var newId='SUPP'+zeroPad(maxNum+1,4);
    var status=normalizeSupplierStatus(supplierData&&supplierData.status);

    sheet.appendRow(hdrs.map(function(h){
      if(h==='SUPPLIER ID') return newId;
      if(h==='SUPPLIER NAME') return name;
      if(h==='STATUS') return status;
      return '';
    }));
    logAudit(adminUsername||'ADMIN','ADD SUPPLIER','SUPPLIERS',newId+' · '+name+' · '+status);
    return {success:true,supplierId:newId};
  } catch(e){return {success:false,message:e.message};}
}

function updateSupplier(rowNum, supplierData, adminUsername) {
  try {
    var name=String((supplierData&&supplierData.supplierName)||'').trim();
    if(!name) return {success:false,message:'Supplier name is required.'};

    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('SUPPLIERS');
    if(!sheet) throw new Error('SUPPLIERS sheet not found.');
    ensureSuppliersSheet(sheet);

    var data=sheet.getDataRange().getValues();
    var hdrs=data[0].map(function(h){return String(h||'').trim().toUpperCase();});
    var nameCol=hdrs.indexOf('SUPPLIER NAME');
    var statusCol=hdrs.indexOf('STATUS');

    for(var i=1;i<data.length;i++){
      if(i+1===rowNum) continue;
      if(String(data[i][nameCol]||'').trim().toLowerCase()===name.toLowerCase())
        return {success:false,message:'Supplier "'+name+'" already exists.'};
    }

    sheet.getRange(rowNum, nameCol+1).setValue(name);
    if(statusCol!==-1) sheet.getRange(rowNum, statusCol+1).setValue(normalizeSupplierStatus(supplierData&&supplierData.status));
    logAudit(adminUsername||'ADMIN','UPDATE SUPPLIER','SUPPLIERS','Row '+rowNum+' · '+name);
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

function deleteSupplier(rowNum, supplierName, adminUsername) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('SUPPLIERS');
    if(!sheet) throw new Error('SUPPLIERS sheet not found.');
    ensureSuppliersSheet(sheet);
    sheet.deleteRow(rowNum);
    logAudit(adminUsername||'ADMIN','DELETE SUPPLIER','SUPPLIERS',supplierName||('Row '+rowNum));
    return {success:true};
  } catch(e){return {success:false,message:e.message};}
}

// ════════════════════════════════════════════════════════════════════════════════
//  END USER MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════════
function ensureEndUsersSheet(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['END USER ID','END USER ROLE','STATUS']);
    return;
  }
  var hdrs = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0]
    .map(function(h){ return String(h||'').trim().toUpperCase(); });
  if (hdrs.indexOf('END USER ID') === -1 && hdrs.indexOf('END USER ROLE') === -1) {
    sheet.clearContents();
    sheet.appendRow(['END USER ID','END USER ROLE','STATUS']);
  }
}

function getEndUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('END USERS');
    if (!sheet) return [];
    ensureEndUsersSheet(sheet);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var hdrs = data[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
    var roleCol = hdrs.indexOf('END USER ROLE');
    var statusCol = hdrs.indexOf('STATUS');
    if (roleCol === -1) return [];
    var seen = {}, out = [];
    for (var i = 1; i < data.length; i++) {
      var status = statusCol === -1 ? 'ACTIVE' : String(data[i][statusCol]||'').trim().toUpperCase();
      if (status === 'INACTIVE') continue;
      var v = String(data[i][roleCol]||'').trim();
      if (v && !seen[v]) { seen[v] = true; out.push(v); }
    }
    return out.sort();
  } catch(e) { return []; }
}

function getAllEndUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('END USERS');
    if (!sheet) return [];
    ensureEndUsersSheet(sheet);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var hdrs = data[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
    var idCol = hdrs.indexOf('END USER ID');
    var roleCol = hdrs.indexOf('END USER ROLE');
    var statusCol = hdrs.indexOf('STATUS');
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var endUserRole = String(data[i][roleCol]||'').trim();
      if (!endUserRole) continue;
      rows.push({
        _rowNum: i + 1,
        endUserId: idCol !== -1 ? String(data[i][idCol]||'').trim() : '',
        endUserRole: endUserRole,
        status: String(statusCol !== -1 ? (data[i][statusCol]||'') : 'ACTIVE').trim().toUpperCase() || 'ACTIVE'
      });
    }
    return rows;
  } catch(e) { return []; }
}

function addEndUser(endUserData, adminUsername) {
  try {
    var endUserRole = String((endUserData && endUserData.endUserRole) || '').trim();
    if (!endUserRole) return {success:false, message:'End User Role is required.'};

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('END USERS') || ss.insertSheet('END USERS');
    ensureEndUsersSheet(sheet);

    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
    var idCol = hdrs.indexOf('END USER ID');
    var roleCol = hdrs.indexOf('END USER ROLE');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][roleCol]||'').trim().toLowerCase() === endUserRole.toLowerCase())
        return {success:false, message:'End User Role "'+endUserRole+'" already exists.'};
    }

    var maxNum = 0;
    for (var j = 1; j < data.length; j++) {
      var euId = String(idCol !== -1 ? data[j][idCol] : '').trim();
      var match = euId.match(/(\d+)$/);
      var n = match ? parseInt(match[1], 10) : 0;
      if (n > maxNum) maxNum = n;
    }
    var newId = 'EU' + zeroPad(maxNum + 1, 4);
    var status = String(endUserData.status || 'ACTIVE').trim().toUpperCase();
    if (status !== 'INACTIVE') status = 'ACTIVE';

    sheet.appendRow(hdrs.map(function(h) {
      if (h === 'END USER ID')   return newId;
      if (h === 'END USER ROLE') return endUserRole;
      if (h === 'STATUS')        return status;
      return '';
    }));
    logAudit(adminUsername||'ADMIN','ADD END USER','END USERS', newId+' · '+endUserRole);
    return {success:true, endUserId:newId};
  } catch(e) { return {success:false, message:e.message}; }
}

function updateEndUser(rowNum, endUserData, adminUsername) {
  try {
    var endUserRole = String((endUserData && endUserData.endUserRole) || '').trim();
    if (!endUserRole) return {success:false, message:'End User Role is required.'};

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('END USERS');
    if (!sheet) throw new Error('END USERS sheet not found.');
    ensureEndUsersSheet(sheet);

    var data = sheet.getDataRange().getValues();
    var hdrs = data[0].map(function(h){ return String(h||'').trim().toUpperCase(); });
    var roleCol = hdrs.indexOf('END USER ROLE');
    var statusCol = hdrs.indexOf('STATUS');

    for (var i = 1; i < data.length; i++) {
      if (i + 1 === rowNum) continue;
      if (String(data[i][roleCol]||'').trim().toLowerCase() === endUserRole.toLowerCase())
        return {success:false, message:'End User Role "'+endUserRole+'" already exists.'};
    }

    var status = String(endUserData.status || 'ACTIVE').trim().toUpperCase();
    if (status !== 'INACTIVE') status = 'ACTIVE';

    if (roleCol !== -1) sheet.getRange(rowNum, roleCol + 1).setValue(endUserRole);
    if (statusCol !== -1) sheet.getRange(rowNum, statusCol + 1).setValue(status);
    logAudit(adminUsername||'ADMIN','UPDATE END USER','END USERS','Row '+rowNum+' · '+endUserRole);
    return {success:true};
  } catch(e) { return {success:false, message:e.message}; }
}

function deleteEndUser(rowNum, endUserName, adminUsername) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('END USERS');
    if (!sheet) throw new Error('END USERS sheet not found.');
    ensureEndUsersSheet(sheet);
    sheet.deleteRow(rowNum);
    logAudit(adminUsername||'ADMIN','DELETE END USER','END USERS', endUserName||('Row '+rowNum));
    return {success:true};
  } catch(e) { return {success:false, message:e.message}; }
}

// ── AUDIT LOGS ────────────────────────────────────────────────────────────────
function getAuditLogs(limitRows) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet=ss.getSheetByName('AUDIT_LOGS');
    if(!sheet||sheet.getLastRow()<2) return [];
    var limit=limitRows||500;
    var data=sheet.getDataRange().getValues();
    var rows=[];
    for(var i=data.length-1;i>=1&&rows.length<limit;i--){
      var r=data[i];
      rows.push({
        timestamp: formatCellForClient('TIMESTAMP', r[0]),
        username:  String(r[1]||''),
        action:    String(r[2]||''),
        department:String(r[3]||''),
        recordId:  String(r[4]||'')
      });
    }
    return rows;
  } catch(e){return [];}
}

function getLatestActivityByTracking(limitRows) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('AUDIT_LOGS');
    if (!sheet || sheet.getLastRow() < 2) return {};

    var limit = Math.max(1, parseInt(limitRows, 10) || 5000);
    var data = sheet.getDataRange().getValues();
    var out = {};
    var tnPattern = /DAR-\d{4}-\d{2}-\d{4}/i;

    // Walk newest to oldest so first hit per tracking no. is its latest activity.
    for (var i = data.length - 1, seen = 0; i >= 1 && seen < limit; i--, seen++) {
      var ts = data[i][0];
      var recordId = String(data[i][4] || '');
      var m = recordId.match(tnPattern);
      if (!m) continue;
      var trackingNo = String(m[0] || '').toUpperCase();
      if (!trackingNo || out[trackingNo]) continue;
      out[trackingNo] = formatCellForClient('TIMESTAMP', ts);
    }
    return out;
  } catch (e) {
    return {};
  }
}

// ── SETUP (run once) ──────────────────────────────────────────────────────────
function setupSheets() {
  var ss=SpreadsheetApp.openById(SPREADSHEET_ID);

  ['BAC','SUPPLY AND PROPERTY','BUDGET','ACCOUNTING','CASH'].forEach(function(name){
    var s=ss.getSheetByName(name)||ss.insertSheet(name);
    ensureIsNew(s);
    ensureWorkflowColumns(s);
  });

  var endUserSheet = ss.getSheetByName('END USER');
  if (endUserSheet) ensureWorkflowColumns(endUserSheet);

  ['USERS','SUPPLIERS'].forEach(function(n){if(!ss.getSheetByName(n)) ss.insertSheet(n);});

  var auditSheet=ss.getSheetByName('AUDIT_LOGS')||ss.insertSheet('AUDIT_LOGS');
  if(auditSheet.getLastRow()===0)
    auditSheet.appendRow(['TIMESTAMP','USER','ACTION','DEPARTMENT','RECORD ID']);

  var usersSheet=ss.getSheetByName('USERS');
  if(usersSheet&&usersSheet.getLastRow()>0){
    var hdrs=usersSheet.getRange(1,1,1,usersSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim().toUpperCase();});
    if(hdrs.indexOf('END USER')===-1)
      usersSheet.getRange(1,usersSheet.getLastColumn()+1).setValue('END USER');
  }

  ensureSuppliersSheet(ss.getSheetByName('SUPPLIERS'));

  var endUsersSheet = ss.getSheetByName('END USERS') || ss.insertSheet('END USERS');
  ensureEndUsersSheet(endUsersSheet);

  return 'Setup complete — all sheets and columns verified.';
}

function continueReturnedTransaction(sheetName, rowNum, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var currentName = normalizeDepartmentName(sheetName);
    var sheet = ss.getSheetByName(currentName);
    if (!sheet) throw new Error('Sheet not found: '+currentName);
    ensureWorkflowColumns(sheet);

    var rowState = getRowRecord(sheet, rowNum);
    var hdrs = rowState.headers;
    var vals = rowState.values;
    var row = rowState.record;
    if (!isReturnResolvedRecord(row)) throw new Error('This item is not in Return Done state.');

    var activeStatus = 'PROCESSING';
    syncStatusValidationForRow(sheet, currentName, rowNum, hdrs, [
      'PROCESSING','RECEIVED','FORWARDED TO SUPPLY','FORWARDED TO BUDGET','FORWARDED TO ACCOUNTING','FORWARDED TO CASHIER',
      'FOR CANVAS','FOR REVIEW','FOR SIGNATURE','LACKING OF ATTACHMENTS','UNPAID','RETURNED','RETURNED TO REQUESTING OFFICE','CANCELLED'
    ]);

    setRecordValuesByHeader(sheet, rowNum, hdrs, vals, {
      'STATUS':activeStatus,
      'COMPLETED AT':'',
      'IS NEW':'',
      'RETURN TO':'',
      'RETURNED FROM':'',
      'RETURN REMARKS':'',
      'RETURNED DATE':'',
      'RETURN RECEIVED DATE':'',
      'RETURN RECEIVED REMARKS':''
    });

    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (trackingNo) {
      var euStatus = currentName === 'BAC' ? 'FORWARDED TO BAC' : (DEPT_RECEIVE_LABELS[currentName] || activeStatus);
      updateEndUserStatus(ss, trackingNo, euStatus);
    }

    logAudit(username||'SYSTEM', 'RETURN CONTINUE', currentName, trackingNo||('Row '+rowNum));
    logTransactionHistory(trackingNo, '', 'RETURN CONTINUE', displayDepartmentName(currentName), '', username, '');
    return {success:true};
  } catch(e) {
    return {success:false, message:e.message};
  }
}

function continueReturnedNew(sheetName, rowNum, username) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var currentName = normalizeDepartmentName(sheetName);
    var sheet = ss.getSheetByName(currentName);
    if (!sheet) throw new Error('Sheet not found: '+currentName);
    ensureWorkflowColumns(sheet);

    var state = getRowRecord(sheet, rowNum);
    var hdrs = state.headers;
    var vals = state.values;
    var row = state.record;

    var isNew = String(row['IS NEW']||'').trim().toUpperCase() === 'TRUE';
    var hasReturn = String(row['RETURNED DATE']||'').trim() !== '';
    if (!isNew || !hasReturn) throw new Error('This row is not a returned NEW transaction.');

    var activeStatus = currentName === 'BAC' ? 'PROCESSING' : 'RECEIVED';
    syncStatusValidationForRow(sheet, currentName, rowNum, hdrs);
    setRecordValuesByHeader(sheet, rowNum, hdrs, vals, {
      'STATUS': activeStatus,
      'COMPLETED AT': '',
      'IS NEW': '',
      'RETURN TO': '',
      'RETURNED FROM': '',
      'RETURN REMARKS': '',
      'RETURNED DATE': '',
      'RETURN RECEIVED DATE': '',
      'RETURN RECEIVED REMARKS': ''
    });

    var trackingNo = String(row['TRACKING NO.'] || '').trim();
    if (trackingNo) {
      var euStatus = DEPT_RECEIVE_LABELS[currentName] || activeStatus;
      updateEndUserStatus(ss, trackingNo, euStatus);
    }
    logAudit(username||'SYSTEM', 'CONTINUE RETURNED NEW', currentName, trackingNo||('Row '+rowNum));
    return {success:true};
  } catch(e) {
    return {success:false, message:e.message};
  }
}