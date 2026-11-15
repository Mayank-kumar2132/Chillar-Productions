/**
 * Chillar Productions – "Curious? Talk to us" 3-step form → Google Sheet
 * Paste into: Google Sheet → Extensions → Apps Script (replace everything), Save.
 * Then Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy.
 * (Who has access must stay "Anyone".)
 */

var HEADERS = [
  'Timestamp', 'What do you do', 'Looking for', 'Name', 'Contact number', 'Email',
  'Existing channel?', 'Channel link', 'OK with ₹50k/month retainer?', 'Time frame', 'Page'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    var p = (e && e.parameter) || {};
    sheet.appendRow([
      new Date(),
      clean(p.what_do),
      clean(p.looking_for),
      clean(p.name),
      "'" + String(p.phone || '').slice(0, 40),
      clean(p.email),
      clean(p.has_channel),
      clean(p.channel_link),
      clean(p.retainer),
      clean(p.duration),
      clean(p.page)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Chillar form endpoint is working');
}

function clean(v) {
  var s = String(v || '').slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
