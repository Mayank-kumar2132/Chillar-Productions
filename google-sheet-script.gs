/**
 * Chillar Productions – "Curious? Talk to us" form → Google Sheet
 * Saves after every step (Next, Next, Submit). One row per person, updated as they go.
 */

var HEADERS = [
  'Timestamp', 'Lead ID', 'Status', 'What do you do', 'Looking for', 'Name', 'Contact number',
  'Email', 'Existing channel?', 'Channel link', 'OK with ₹50k/month retainer?', 'Time frame', 'Page'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    var p = (e && e.parameter) || {};
    var row = [
      new Date(),
      clean(p.lead_id),
      clean(p.status),
      clean(p.what_do),
      clean(p.looking_for),
      clean(p.name),
      p.phone ? "'" + String(p.phone).slice(0, 40) : '',
      clean(p.email),
      clean(p.has_channel),
      clean(p.channel_link),
      clean(p.retainer),
      clean(p.duration),
      clean(p.page)
    ];
    // same person (same Lead ID) → update their row, otherwise add a new row
    var target = 0, last = sheet.getLastRow();
    if (p.lead_id && last > 1) {
      var ids = sheet.getRange(2, 2, last - 1, 1).getValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(p.lead_id)) { target = i + 2; break; }
      }
    }
    if (target) sheet.getRange(target, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Chillar form endpoint v3 is working');
}

function clean(v) {
  var s = String(v || '').slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
