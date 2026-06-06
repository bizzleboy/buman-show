// ─────────────────────────────────────────────────────────────────────────────
//  Paste this entire file into the Google Apps Script editor.
//  File → Save → Deploy → New deployment → Web app
//
//  Creates two sheet tabs automatically:
//    "Responses" — radio button answers (Date, Heard Via)
//    "Emails"    — email sign-ups (Date, Email)
// ─────────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();

    if (data.type === 'email') {
      var emailSheet = ss.getSheetByName('Emails');
      if (!emailSheet) {
        emailSheet = ss.insertSheet('Emails');
        emailSheet.appendRow(['Date', 'Email']);
        emailSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      }
      emailSheet.appendRow([data.date || '', data.email || '']);

    } else if (data.type === 'response') {
      var responseSheet = ss.getSheetByName('Responses');
      if (!responseSheet) {
        responseSheet = ss.insertSheet('Responses');
        responseSheet.appendRow(['Date', 'Heard Via']);
        responseSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      }
      responseSheet.appendRow([data.date || '', data.source || '']);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
