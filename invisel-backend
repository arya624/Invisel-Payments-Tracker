/**
 * Invisel Payment Ledger — backend
 *
 * SETUP
 * 1. Go to https://sheets.google.com and create a new, blank Google Sheet.
 *    (Name it whatever you like, e.g. "Invisel Payment Ledger Data".)
 * 2. In that Sheet: Extensions -> Apps Script.
 * 3. Delete anything in the editor and paste this whole file in.
 * 4. Click the disk icon (or Ctrl/Cmd+S) to save. Give the project a name.
 * 5. Click "Deploy" -> "New deployment".
 *    - Click the gear icon next to "Select type" -> choose "Web app".
 *    - Description: anything.
 *    - Execute as: "Me".
 *    - Who has access: "Anyone".
 *    - Click "Deploy". Authorize when Google asks (click through the
 *      "unsafe" warning — this is your own script, it's fine).
 * 6. Copy the "Web app URL" it gives you — it should look like
 *      https://script.google.com/macros/s/XXXXXXXX/exec
 *    (or, on a Workspace domain like invisel.in, it may look like
 *      https://script.google.com/a/macros/invisel.in/s/XXXXXXXX/exec —
 *    either format is fine, just copy exactly what Google shows you.)
 * 7. Open invisel-payment-tracker.html in a text editor, find the line:
 *        const APPS_SCRIPT_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
 *    and replace the placeholder text with the URL you copied.
 * 8. Save the HTML file and re-upload it to GitHub Pages (or wherever
 *    you're hosting it). Open it — the sidebar should say
 *    "Synced to your Google Sheet."
 *
 * NOTE: every time you change this script's code later, you must create
 * a NEW deployment version (Deploy -> Manage deployments -> edit -> New
 * version) for the changes to take effect on the live URL.
 *
 * NOTE: "Execute as" must be "Me", not "User accessing the web app" —
 * the latter causes a "Sorry, unable to open the file" error for anyone
 * who doesn't personally have Drive access to this Sheet.
 *
 * NOTE ON ATTACHMENTS: uploaded invoice files are saved to a Google Drive
 * folder named "Invisel Payment Ledger Attachments" (auto-created on
 * first upload, in the Drive of whoever deployed this script) and shared
 * as "anyone with the link can view". That's the same security model as
 * the exec URL itself — not gated by Google sign-in, protected only by
 * the link being long and unguessable. Don't attach anything you wouldn't
 * want reachable by anyone who somehow obtained that specific file link.
 */

function doGet(e) {
  var sheet = getSheet_();
  var rows = sheet.getDataRange().getValues();
  var data = {};
  for (var i = 1; i < rows.length; i++) {
    var key = rows[i][0];
    var value = rows[i][1];
    if (key) data[key] = value;
  }
  return jsonOut_({ ok: true, data: data });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'uploadFile') return handleUpload_(body);

    var key = body.key;
    var value = body.value;
    if (!key) throw new Error('Missing key');

    var sheet = getSheet_();
    var rows = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) { rowIndex = i + 1; break; } // +1: sheet rows are 1-indexed
    }
    var now = new Date();
    if (rowIndex === -1) {
      sheet.appendRow([key, value, now]);
    } else {
      sheet.getRange(rowIndex, 2).setValue(value);
      sheet.getRange(rowIndex, 3).setValue(now);
    }
    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function handleUpload_(body) {
  try {
    if (!body.base64 || !body.filename) throw new Error('Missing file data');
    var bytes = Utilities.base64Decode(body.base64);
    var blob = Utilities.newBlob(bytes, body.mimeType || 'application/octet-stream', body.filename);
    var folder = getAttachmentsFolder_();
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return jsonOut_({ ok: true, url: file.getUrl(), fileId: file.getId(), name: body.filename });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getAttachmentsFolder_() {
  var name = 'Invisel Payment Ledger Attachments';
  var existing = DriveApp.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(name);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  if (!sheet) {
    sheet = ss.insertSheet('Data');
    sheet.appendRow(['Key', 'Value', 'UpdatedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
