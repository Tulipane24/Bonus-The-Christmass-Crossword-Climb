/**
 * Lead collector for "The Christmas Crossword Climb" bonus page.
 * Paste this file as-is into Extensions > Apps Script of your Google Sheet.
 * Leads are written to the tab named "Lead" (created automatically).
 */
var SHARED_SECRET = 'NeUEzBBzcx_ewg1nOim1gFdN5BTpxUWR';
var SPREADSHEET_ID = '1W88iqzqoAuHNdblR9mF7uVBldOxzeAxRZ_K4Jg7c-DA';
var SHEET_NAME = 'Lead';
var HEADERS = ['Data e ora', 'Nome', 'Cognome', 'Email', 'Libro', 'Invii'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss;
  try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch (e) { ss = SpreadsheetApp.getActiveSpreadsheet(); }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.secret !== SHARED_SECRET) return json_({ ok: false, error: 'unauthorized' });
    var email = String(data.email || '').trim().toLowerCase();
    var nome = String(data.nome || '').trim();
    var cognome = String(data.cognome || '').trim();
    if (!email || !nome || !cognome) return json_({ ok: false, error: 'missing fields' });

    lock.waitLock(10000);
    var sh = getSheet_();
    var last = sh.getLastRow();
    var row = -1;
    if (last > 1) {
      var emails = sh.getRange(2, 4, last - 1, 1).getValues();
      for (var i = 0; i < emails.length; i++) {
        if (String(emails[i][0]).trim().toLowerCase() === email) { row = i + 2; break; }
      }
    }
    var when = data.timestamp ? new Date(data.timestamp) : new Date();
    if (row > 0) {
      var count = Number(sh.getRange(row, 6).getValue()) || 1;
      sh.getRange(row, 1, 1, 6).setValues([[when, nome, cognome, email, data.libro || '', count + 1]]);
    } else {
      sh.appendRow([when, nome, cognome, email, data.libro || '', 1]);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet() {
  return json_({ ok: true, info: 'Lead collector is running.' });
}
