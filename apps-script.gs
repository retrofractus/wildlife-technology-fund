/* Wildlife Technology Fund — "Join us" form collector
   ------------------------------------------------------------------
   Writes each submission from Join-us.dc.html into a Google Sheet.

   Setup (one time):
   1. Create a Google Sheet to hold responses.
   2. In the Sheet: Extensions → Apps Script. Delete the placeholder code
      and paste in this whole file. Save.
   3. Deploy → New deployment → type "Web app".
        Execute as: Me
        Who has access: Anyone
      Authorise it when prompted.
   4. Copy the web app URL (ends in /exec) and paste it into ENDPOINT at
      the top of the <script> in Join-us.dc.html.
   5. Submit a test response on the site. A "responses" tab with headers
      appears in the Sheet.

   After editing this script, redeploy with Deploy → Manage deployments →
   Edit → Version: New version, or the old code keeps running.            */

const SHEET_NAME = 'responses';

/* Optional: an address to email whenever someone submits. Leave blank for
   no notifications. */
const NOTIFY_EMAIL = '';

const HEADERS = [
  'submitted_at',
  'name',
  'institution',
  'email',
  'expertise',
  'linkedin',
  'interest',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);

    // Hidden "website" field: real people never see it, bots fill it in.
    // Report success so the bot moves on, but don't store anything.
    if (data.website) return json({ ok: true });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }

    const row = HEADERS.map((h) => clean(data[h]));
    sheet.appendRow(row);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New Join us submission: ' + clean(data.name),
        body: HEADERS.map((h, i) => h + ': ' + row[i]).join('\n\n'),
      });
    }

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* Trim, cap length, and stop anything that starts like a formula from being
   run as one when the Sheet opens it. */
function clean(value) {
  let v = value == null ? '' : String(value).trim().slice(0, 5000);
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
