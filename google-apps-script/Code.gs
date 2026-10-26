/**
 * Receives applications from the membership form and writes them to the first
 * sheet of this spreadsheet. Each application has an ID in column A: saving the
 * same application again updates its row, a new application adds a row.
 *
 * Setup: open the Google Sheet → Extensions → Apps Script, paste this file,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web app URL into SHEET_SCRIPT_URL in Pilgrim_Preceptors_Membership_Application.html.
 */
function doPost(e) {
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
