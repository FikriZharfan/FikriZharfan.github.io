// Copy this into Extensions -> Apps Script in your Google Spreadsheet project
// Then Deploy -> New deployment -> Select "Web app"
// Set "Who has access" to "Anyone" (or "Anyone, even anonymous") and deploy.

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ready", message: "Apps Script Web App is deployed." })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var payload = {};

    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        // fallback to form data
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    var name = payload.name || "";
    var email = payload.email || "";
    var message = payload.message || "";
    sheet.appendRow([new Date(), name, email, message]);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
