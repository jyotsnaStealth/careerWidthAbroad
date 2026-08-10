function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getActiveSheet();

    // Parse JSON payload or URL-encoded form parameters
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

    // Formatted timestamp (e.g. 10/08/2026 13:05:00)
    var timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'dd/MM/yyyy HH:mm:ss'
    );

    // Standardized lead data
    var payload = {
      timestamp: timestamp,
      name: data.name || '',
      phone: data.phone || '',
      email: data.email || '',
      country: data.country || '',
      qualification: data.qualification || '',
      exam: data.exam || '',
      city: data.city || '',
      course: data.course || '',
      utmSource: data.utmSource || data.utm_source || '',
      utmMedium: data.utmMedium || data.utm_medium || '',
      utmCampaign: data.utmCampaign || data.utm_campaign || '',
      utmTerm: data.utmTerm || data.utm_term || '',
      utmContent: data.utmContent || data.utm_content || '',
      formSource: data.formName || data.formSource || ''
    };

    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      // Dynamically map values to match your sheet's exact column headers in Row 1
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var newRow = [];

      for (var i = 0; i < headers.length; i++) {
        var h = (headers[i] || '').toString().trim().toLowerCase();

        if (h.indexOf('timestamp') !== -1 || h.indexOf('date') !== -1 || h.indexOf('time') !== -1) {
          newRow.push(payload.timestamp);
        } else if (h.indexOf('name') !== -1 && h.indexOf('form') === -1) {
          newRow.push(payload.name);
        } else if (h.indexOf('phone') !== -1 || h.indexOf('mobile') !== -1 || h.indexOf('contact') !== -1) {
          newRow.push(payload.phone);
        } else if (h.indexOf('email') !== -1) {
          newRow.push(payload.email);
        } else if (h.indexOf('country') !== -1 || h.indexOf('destination') !== -1) {
          newRow.push(payload.country);
        } else if (h.indexOf('qual') !== -1) {
          newRow.push(payload.qualification);
        } else if (h.indexOf('exam') !== -1) {
          newRow.push(payload.exam);
        } else if (h.indexOf('city') !== -1 || h.indexOf('location') !== -1) {
          newRow.push(payload.city);
        } else if (h.indexOf('course') !== -1 || h.indexOf('program') !== -1) {
          newRow.push(payload.course);
        } else if (h.indexOf('utm_source') !== -1 || h.indexOf('utm source') !== -1) {
          newRow.push(payload.utmSource);
        } else if (h.indexOf('utm_medium') !== -1 || h.indexOf('utm medium') !== -1) {
          newRow.push(payload.utmMedium);
        } else if (h.indexOf('utm_campaign') !== -1 || h.indexOf('utm campaign') !== -1) {
          newRow.push(payload.utmCampaign);
        } else if (h.indexOf('utm_term') !== -1 || h.indexOf('utm term') !== -1) {
          newRow.push(payload.utmTerm);
        } else if (h.indexOf('utm_content') !== -1 || h.indexOf('utm content') !== -1) {
          newRow.push(payload.utmContent);
        } else if (h.indexOf('form') !== -1) {
          newRow.push(payload.formSource);
        } else if (h.indexOf('source') !== -1) {
          // Plain 'Source' header maps to UTM Source
          newRow.push(payload.utmSource);
        } else if (h.indexOf('medium') !== -1) {
          newRow.push(payload.utmMedium);
        } else if (h.indexOf('campaign') !== -1) {
          newRow.push(payload.utmCampaign);
        } else {
          newRow.push('');
        }
      }

      sheet.appendRow(newRow);
    } else {
      // Fallback row if sheet has no headers
      sheet.appendRow([
        payload.timestamp,
        payload.name,
        payload.phone,
        payload.email,
        payload.country,
        payload.qualification,
        payload.exam,
        payload.city,
        payload.course,
        payload.utmSource,
        payload.utmMedium,
        payload.utmCampaign,
        payload.formSource
      ]);
    }

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
