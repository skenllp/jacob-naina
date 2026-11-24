// Deploy as: Web App → Execute as: Me → Who has access: Anyone
// After deploying, paste the /exec URL into CONFIG.RSVP_ENDPOINT in script.js

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var raw = e.postData ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'RSVPs';
    var sheet = ss.getSheetByName(sheetName);

    // Create the sheet and header row if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'Timestamp', 'ID', 'Name', 'Attendance',
        'Guests', 'Events', 'Message', 'Submitted At'
      ]);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.id || '',
      data.name || '',
      data.attendance || '',
      data.guests || 0,
      (data.events || []).join(', '),
      data.message || '',
      data.submittedAt || ''
    ]);

    output.setContent(JSON.stringify({ ok: true }));
  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return output;
}

// Test by running this manually in the Apps Script editor
function testSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet: ' + ss.getName());
  var sheet = ss.getSheetByName('RSVPs');
  Logger.log('Sheet exists: ' + (sheet !== null));
}
