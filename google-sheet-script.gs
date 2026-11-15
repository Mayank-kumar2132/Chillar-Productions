/**
 * Chillar Productions – website form → Google Sheet
 * Paste this into: Google Sheet → Extensions → Apps Script, then Deploy as a Web app.
 * Every form submission on the website is added as a new row in the first tab.
 */

var HEADERS = [
  'Timestamp', 'Form', 'Name', 'Phone / WhatsApp', 'Email',
  'Instagram / Brand', 'I am a', 'Needs', 'Message', 'Page'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    var p = (e && e.parameter) || {};
    sheet.appendRow([
      new Date(),
      clean(p.form),
      clean(p.name),
      "'" + String(p.phone || '').slice(0, 40),   // keep +91 numbers as text
      clean(p.email),
      clean(p.handle),
      clean(p.role),
      clean(p.needs),
      clean(p.message),
      clean(p.page)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Stops anyone typing a spreadsheet formula into the form, and limits length.
function clean(v) {
  var s = String(v || '').slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// Optional: run once from the editor to create the header row.
function setupHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}
