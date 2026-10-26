/**
 * Receives applications from the membership form and writes them to the first
 * sheet of this spreadsheet. Each application has an ID in column A: saving the
 * same application again updates its row, a new application adds a row.
 *
 * Setup: open the Google Sheet → Extensions → Apps Script, paste this file,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web app URL into SHEET_SCRIPT_URL in Pilgrim_Preceptors_Membership_Application.html.
 * To check the setup, choose setup in the function menu and press Run (do not run doPost).
 */

/** Safe to Run from the editor: asks for permission and checks the sheet has its header row. */
function setup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].filter(String);
  Logger.log('Sheet "%s" is ready with %s columns (first: %s).', sheet.getName(), headers.length, headers[0]);
}

/** Opening the web app URL in a browser shows this, to confirm the deployment works. */
function doGet() {
  return ContentService.createTextOutput('Pilgrim Preceptors sheet connection is working.');
}

function doPost(e) {
  if (!e || !e.postData) {
    // Pressing Run in the editor calls doPost with no data; run setup() instead.
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No data received' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const data = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const last = sheet.getLastRow();
    const ids = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    (data.rows || []).forEach(row => {
      const values = headers.map(h => (row[h] == null ? '' : String(row[h])));
      let i = ids.indexOf(String(row.ID));
      if (i < 0) { ids.push(String(row.ID)); i = ids.length - 1; }
      // Keep entries as typed text (dates, phone numbers with leading zeros).
      const range = sheet.getRange(i + 2, 1, 1, headers.length);
      range.setNumberFormat('@');
      range.setValues([values]);
    });
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
