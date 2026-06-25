// ============================================================
// Transborder Logistics — Google Apps Script Backend v2
// Proper sheets: Clients, Airlines, Orders, Tickets, Invoices,
// Reimbursements — each with their own columns and rows
// ============================================================

var SPREADSHEET_ID = '1ey_YJm-gvDJ8SRitulfxw5BYyDYgnYOAYtDNHuz_QNI';
var DRIVE_FOLDER_ID = '1zIRQ2sHBzNpDOmTcKnMNIxeV-_4g63Lt';

// Sheet definitions — name + columns in order
var SHEETS = {
  meta:             { name: 'Meta',             cols: ['key','value'] },
  clients:          { name: 'Clients',          cols: ['id','name','address','type','tier','contact','email','phone','createdAt'] },
  airlines:         { name: 'Airlines',         cols: ['id','name','code'] },
  orders:           { name: 'Orders',           cols: ['id','clientId','orderRef','notes','createdAt'] },
  tickets:          { name: 'Tickets',          cols: ['id','orderId','passengerName','airlineId','route','bookingRef','travelDate','ticketCost','serviceCharge','status','usageFlag','invoiceId','attachmentUrl','attachmentName','createdAt'] },
  invoices:         { name: 'Invoices',         cols: ['id','clientId','invoiceNo','issueDate','dueDate','status','subtotal','vat','total','createdAt'] },
  reimbursements:   { name: 'Reimbursements',   cols: ['id','clientId','invoiceIds','amountExpected','notes','createdAt'] },
  payments:         { name: 'Payments',         cols: ['id','reimbursementId','amount','date','note','createdAt'] },
};

// ============================================================
// SERVE THE APP — doGet handles BOTH serving the HTML and
// API calls (via ?payload=... query param) so CORS is never
// an issue — GitHub Pages uses fetch GET with the payload.
// ============================================================
function doGet(e) {
  // If a payload param is present this is an API call, not a page load
  if (e && e.parameter && e.parameter.payload) {
    var result;
    try {
      var payload = JSON.parse(decodeURIComponent(e.parameter.payload));
      var action = payload.action;
      if      (action === 'loadAll')    result = loadAll();
      else if (action === 'saveAll')    result = saveAll(payload.db);
      else if (action === 'uploadFile') result = uploadFile(payload.fileName, payload.mimeType, payload.base64Data);
      else result = { error: 'Unknown action: ' + action };
    } catch(err) {
      result = { error: err.toString() };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Otherwise serve the HTML app
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Transborder Logistics — Ticketing & Invoicing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// HANDLE DATA REQUESTS via POST (used when served via Apps Script URL)
// ============================================================
function doPost(e) {
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    if      (action === 'loadAll')    result = loadAll();
    else if (action === 'saveAll')    result = saveAll(payload.db);
    else if (action === 'uploadFile') result = uploadFile(payload.fileName, payload.mimeType, payload.base64Data);
    else result = { error: 'Unknown action: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CALLED BY google.script.run FROM HTML
// ============================================================
function handleRequest(payloadStr) {
  try {
    var payload = JSON.parse(payloadStr);
    var action = payload.action;
    if      (action === 'loadAll')    return loadAll();
    else if (action === 'saveAll')    return saveAll(payload.db);
    else if (action === 'uploadFile') return uploadFile(payload.fileName, payload.mimeType, payload.base64Data);
    else return { error: 'Unknown action: ' + action };
  } catch(err) {
    return { error: err.toString() };
  }
}

// ============================================================
// SETUP — run once manually to create all sheets
// ============================================================
function setup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Rename old 'data' sheet if exists (can't delete last sheet)
  var oldSheet = ss.getSheetByName('data');

  // Create all sheets with headers
  for (var key in SHEETS) {
    var def = SHEETS[key];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) sheet = ss.insertSheet(def.name);
    // Write headers in row 1
    sheet.clearContents();
    var headers = def.cols;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // Format header row
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0B2547');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
  }

  // Set meta values
  var metaSheet = ss.getSheetByName('Meta');
  metaSheet.getRange(2, 1).setValue('nextInvoiceSeq');
  metaSheet.getRange(2, 2).setValue('1');

  Logger.log('Setup complete — all sheets created!');
}

// ============================================================
// LOAD ALL DATA — read each sheet into DB object
// ============================================================
function loadAll() {
  try {
  Logger.log('loadAll started, SPREADSHEET_ID: ' + SPREADSHEET_ID);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('Spreadsheet opened: ' + ss.getName());
  var allSheets = ss.getSheets().map(function(s){ return s.getName(); });
  Logger.log('All sheets: ' + allSheets.join(', '));
  var db = {
    clients: [], airlines: [], orders: [], tickets: [],
    invoices: [], reimbursements: [], payments: [],
    meta: { nextInvoiceSeq: 1 }
  };

  for (var key in SHEETS) {
    var def = SHEETS[key];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue; // only headers, no data

    var headers = data[0];

    if (key === 'meta') {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === 'nextInvoiceSeq') {
          db.meta.nextInvoiceSeq = parseInt(data[i][1]) || 1;
        }
      }
      continue;
    }

    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // skip empty rows
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var val = row[j];
        // Convert Date objects back to YYYY-MM-DD strings
        if (val instanceof Date) {
          var yyyy = val.getFullYear();
          var mm = String(val.getMonth() + 1).padStart(2, '0');
          var dd = String(val.getDate()).padStart(2, '0');
          val = yyyy + '-' + mm + '-' + dd;
        }
        // Parse booleans
        else if (val === 'TRUE' || val === true) val = true;
        else if (val === 'FALSE' || val === false) val = false;
        // Parse numbers stay as numbers
        else if (typeof val === 'number') val = val;
        // Parse JSON arrays (invoiceIds)
        else if (typeof val === 'string' && val.startsWith('[')) {
          try { val = JSON.parse(val); } catch(e) {}
        }
        // Empty string to null for nullable fields
        else if (val === '') val = null;
        record[headers[j]] = val;
      }
      records.push(record);
    }
    if (key === 'payments') db.payments = records;
    else db[key] = records;
  }

  Logger.log('loadAll result — clients: ' + db.clients.length + ', tickets: ' + db.tickets.length + ', orders: ' + db.orders.length);
  if (db.clients.length > 0) Logger.log('First client: ' + JSON.stringify(db.clients[0]));
  return { db: db };
  } catch(err) {
    Logger.log('ERROR in loadAll: ' + err.toString());
    throw err;
  }
}

// ============================================================
// SAVE ALL DATA — write each collection to its sheet
// ============================================================
function saveAll(db) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Ensure sheets exist
  for (var key in SHEETS) {
    var def = SHEETS[key];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.cols.length).setValues([def.cols]);
      var hr = sheet.getRange(1, 1, 1, def.cols.length);
      hr.setBackground('#0B2547');
      hr.setFontColor('#ffffff');
      hr.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }

  var collections = {
    clients: db.clients || [],
    airlines: db.airlines || [],
    orders: db.orders || [],
    tickets: db.tickets || [],
    invoices: db.invoices || [],
    reimbursements: db.reimbursements || [],
    payments: db.payments || [],
  };

  for (var key in collections) {
    var def = SHEETS[key];
    if (!def) continue;
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) continue;

    var records = collections[key];
    var cols = def.cols;

    // Clear existing data (keep header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, cols.length).clearContent();

    if (records.length === 0) continue;

    // Write all records
    var rows = records.map(function(record) {
      return cols.map(function(col) {
        var val = record[col];
        if (val === null || val === undefined) return '';
        if (Array.isArray(val)) return JSON.stringify(val);
        // Prefix date strings with apostrophe to prevent Sheets auto-converting
        if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val)) return "'" + val;
        return val;
      });
    });

    sheet.getRange(2, 1, rows.length, cols.length).setValues(rows);
    sheet.autoResizeColumns(1, cols.length);
  }

  // Save meta
  var metaSheet = ss.getSheetByName('Meta');
  if (metaSheet && db.meta) {
    metaSheet.getRange(2, 1).setValue('nextInvoiceSeq');
    metaSheet.getRange(2, 2).setValue(db.meta.nextInvoiceSeq || 1);
  }

  return { ok: true, savedAt: new Date().toISOString() };
}

// ============================================================
// FILE UPLOAD TO DRIVE
// ============================================================
function uploadFile(fileName, mimeType, base64Data) {
  var folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch(e) {
    folder = DriveApp.createFolder('Transborder Ticket Attachments');
  }
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType, fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    url: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    fileId: file.getId(),
    name: fileName
  };
}

// ============================================================
// DEBUG — run this manually to see what sheets exist
// ============================================================
function debugSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  sheets.forEach(function(s) {
    var lastRow = s.getLastRow();
    var lastCol = s.getLastColumn();
    Logger.log('Sheet: "' + s.getName() + '" — rows: ' + lastRow + ', cols: ' + lastCol);
    if (lastRow > 0 && lastCol > 0) {
      var headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
      Logger.log('  Headers: ' + headers.join(', '));
      if (lastRow > 1) {
        var firstRow = s.getRange(2, 1, 1, lastCol).getValues()[0];
        Logger.log('  First row: ' + firstRow.join(', '));
      }
    }
  });
}
