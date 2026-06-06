// ─────────────────────────────────────────────────────────────────────────────
//  1. Replace SPREADSHEET_ID below with your Google Sheet ID.
//     (The long string in your sheet's URL between /d/ and /edit)
//  2. Paste into Apps Script editor → Save → Deploy → New deployment → Web app
//     Execute as: Me | Who has access: Anyone
// ─────────────────────────────────────────────────────────────────────────────

var SPREADSHEET_ID = '1cB-RJJdbxgrls6IyYG_m_i2Uft2U_5CN0W_eD_OFQhg';

function doGet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ContentService
    .createTextOutput('Connected to: ' + ss.getName())
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (data.type === 'response') {
      var sheet = ss.getSheetByName('Responses');
      if (!sheet) {
        sheet = ss.insertSheet('Responses');
        sheet.appendRow(['Date', 'Heard Via']);
        sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      }
      sheet.appendRow([data.date || '', data.source || '']);

    } else if (data.type === 'email') {
      var sheet = ss.getSheetByName('Emails');
      if (!sheet) {
        sheet = ss.insertSheet('Emails');
        sheet.appendRow(['Date', 'Email']);
        sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      }
      sheet.appendRow([data.date || '', data.email || '']);
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
