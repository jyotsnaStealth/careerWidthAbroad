function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();

    // Parse JSON body or form parameters
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    // Current formatted timestamp (e.g., 10/08/2026 11:35:00)
    var timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'dd/MM/yyyy HH:mm:ss'
    );

    // Row ordered according to your Google Sheet columns:
    // A: Timestamp
    // B: Full Name
    // C: Phone
    // D: Email
    // E: Preferred Country
    // F: Current Qualification
    // G: Any Exam Given
    // H: City
    // I: Interested Course
    // (Optional extra tracking columns J onwards)
    var newRow = [
      timestamp,
      data.name || '',
      data.phone || '',
      data.email || '',
      data.country || '',
      data.qualification || '',
      data.exam || '',
      data.city || '',
      data.course || '',
      data.source || '',
      data.utmSource || '',
      data.utmMedium || '',
      data.utmCampaign || '',
      data.pageUrl || ''
    ];

    sheet.appendRow(newRow);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Google Apps Script Lead Webhook is active.");
}
