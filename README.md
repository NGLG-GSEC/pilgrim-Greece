# Pilgrim Preceptors – Membership Application

Live form: https://nglg-gsec.github.io/pilgrim-greece/ (once GitHub Pages is enabled)

- **Print / Save PDF** produces the two-page A4 form laid out exactly like the Word form ΑΙΤΗΣΗ, filled with the typed data and blue signatures.
- **ADMIN** (code required) shows the applications saved on this device, the Excel export, and **Excel data**, which opens the Google Sheet with all registrations for editing.

## Connecting the Google Sheet

The sheet *Pilgrim Preceptors – Membership Applications* is in the owner's Google Drive.
To have every saved application written into it:

1. Open the sheet → **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with [`google-apps-script/Code.gs`](google-apps-script/Code.gs) and save.
3. **Deploy → New deployment → Web app**: *Execute as* **Me**, *Who has access* **Anyone**. Authorise when asked.
4. Copy the web app URL and paste it into `SHEET_SCRIPT_URL` near the top of the script in `Pilgrim_Preceptors_Membership_Application.html`.

After that, **Save application** sends the entry to the sheet, and **Send all to Google Sheet** (in ADMIN) sends everything saved on that device.
