/**
 * PILIHAN: kumpul markah dari semua peranti murid ke dalam satu Google Sheet.
 *
 * Cara pasang:
 * 1. Buka Google Sheet baharu > Extensions > Apps Script.
 * 2. Padam kod sedia ada, tampal semua kod ini, Save.
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me      Who has access: Anyone
 * 4. Salin URL Web App (berakhir dengan /exec).
 * 5. Dalam index.html, tampal URL itu pada CONFIG.SHEET_URL.
 */
const SHEET_NAME = 'Results';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['id', 'Nama', 'Kelas', 'Markah', 'Masa', 'json']);
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const v = sheet_().getDataRange().getValues().slice(1);
  const records = [];
  v.forEach(r => { if (r[0] && r[5]) { try { records.push(JSON.parse(r[5])); } catch (e) {} } });
  return out_({ records: records });
}
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_();
    const ids = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
    if (b.action === 'save') {
      const r = b.record;
      const row = [r.id, r.name, r.cls, r.total, new Date(r.ts), JSON.stringify(r)];
      const i = ids.indexOf(String(r.id));
      if (i >= 0) s.getRange(i + 2, 1, 1, 6).setValues([row]); else s.appendRow(row);
    } else if (b.action === 'delete') {
      const i = ids.indexOf(String(b.id));
      if (i >= 0) s.deleteRow(i + 2);
    } else if (b.action === 'clear') {
      if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    }
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
