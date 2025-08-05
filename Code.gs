function doGet(e) {
  const action = e.parameter.action;
  const path = e.pathInfo || '';
  
  switch(action) {
    case 'getWOs':
      return ContentService.createTextOutput(
        JSON.stringify(getWOs())
      ).setMimeType(ContentService.MimeType.JSON);
      
    case 'exportLookupAsExcel':
      const wo = e.parameter.wo;
      const fileObj = exportLookupAsExcel(wo);
      return ContentService.createTextOutput(
        JSON.stringify(fileObj)
      ).setMimeType(ContentService.MimeType.JSON);
      
    default:
      return HtmlService.createHtmlOutput("Invalid action");
  }
}

// Keep your existing functions:
// getWOs(), exportLookupAsExcel(), sendLookupExcelByEmail(), etc.

  if (path === 'manifest.webmanifest') {
    return ContentService
      .createTextOutput(HtmlService.createHtmlOutputFromFile('manifest.webmanifest').getContent())
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (path === 'service-worker.js') {
    return ContentService
      .createTextOutput(HtmlService.createHtmlOutputFromFile('service-worker.js').getContent())
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return HtmlService.createHtmlOutputFromFile('Index');
}

// Get unique WO list from "PO DATABASE" column V
function getWOs() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('PO DATABASE');
  const values = sheet.getRange('V2:V' + sheet.getLastRow()).getValues();
  return [...new Set(values.flat().filter(v => v))];
}

// Create & return Excel as base64, then delete from Drive
function exportLookupAsExcel(wo) {
  const ss = SpreadsheetApp.getActive();
  const lookupSheet = ss.getSheetByName('LOOKUP');
  lookupSheet.getRange('B2').setValue(wo);

  const tempSheet = lookupSheet.copyTo(ss).setName('TEMP_EXPORT');
  const range = tempSheet.getDataRange();
  range.copyTo(range, { contentsOnly: true });

  const tempFile = SpreadsheetApp.create('TEMP_' + wo);
  tempSheet.copyTo(tempFile).setName('LOOKUP');
  tempFile.deleteSheet(tempFile.getSheets()[0]);

  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + tempFile.getId() +
    '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });

  const blob = response.getBlob().setName(wo + '_LOOKUP.xlsx');
  const base64Data = Utilities.base64Encode(blob.getBytes());

  // Delete temp files from Drive
  DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  ss.deleteSheet(tempSheet);

  return { filename: wo + '_LOOKUP.xlsx', base64: base64Data };
}

// Send Excel via email
function sendLookupExcelByEmail(wo, email) {
  const ss = SpreadsheetApp.getActive();
  const lookupSheet = ss.getSheetByName('LOOKUP');
  lookupSheet.getRange('B2').setValue(wo);

  const tempSheet = lookupSheet.copyTo(ss).setName('TEMP_EXPORT');
  const range = tempSheet.getDataRange();
  range.copyTo(range, { contentsOnly: true });

  const tempFile = SpreadsheetApp.create('TEMP_' + wo);
  tempSheet.copyTo(tempFile).setName('LOOKUP');
  tempFile.deleteSheet(tempFile.getSheets()[0]);

  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + tempFile.getId() +
    '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });

  const blob = response.getBlob().setName(wo + '_LOOKUP.xlsx');
  MailApp.sendEmail({
    to: email,
    subject: `WO ${wo} - LOOKUP Sheet Export`,
    body: `Dear User,\n\nPlease find attached the LOOKUP sheet for WO ${wo}.\n\n- Nilambar`,
    attachments: [blob]
  });

  // Delete temp files
  DriveApp.getFileById(tempFile.getId()).setTrashed(true);
  ss.deleteSheet(tempSheet);

  return `✔ Excel file for WO ${wo} has been sent to ${email}`;
}

// Log usage
function logAnalytics(event, detail) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Analytics') || SpreadsheetApp.getActive().insertSheet('Analytics');
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), event, detail]);
}
