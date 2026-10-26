/**
 * Coconut Test Logger – Google Sheets backend
 * Paste this whole file into Extensions > Apps Script (replace everything there).
 * Change PIN below, then Deploy > New deployment > Web app.
 */
const PIN = 'change-me';          // <-- choose your own PIN (must match index.html)
const SHEET_NAME = 'Test Data';

const HEAD = ['ID', 'Date', 'Whole coconut', 'water', 'shell & husk', 'Black Kernel', 'Testa',
  'Diced (after cut)', 'Tombong', 'White kernel', 'Testa %', 'Diced % (after cut)',
  'Tombong %', 'White kernel %', 'Tested by', 'Remarks', 'Created'];

function doGet(e) {
  return out_(handle_(e.parameter || {}));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return out_(handle_(body));
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  if (String(p.pin) !== PIN) return { ok: false, error: 'Wrong PIN' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (p.action === 'list') return { ok: true, records: list_(sh) };
    if (p.action === 'save') { save_(sh, p.record); return { ok: true }; }
    if (p.action === 'delete') { del_(sh, p.id); return { ok: true }; }
    return { ok: false, error: 'Unknown action' };
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length)
      .setFontWeight('bold').setBackground('#1f6f4a').setFontColor('#ffffff');
    sh.getRange(1, 11, 1, 4).setBackground('#3b4a8f');
    sh.setColumnWidths(1, HEAD.length, 110);
    sh.hideColumns(1);   // ID column (internal) – unhide if you want to see it
  }
  return sh;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function save_(sh, r) {
  const row = findRow_(sh, r.id);
  const parts = String(r.date).split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
  let t = row;
  if (!t) {
    t = sh.getLastRow() + 1;
    sh.getRange(t, 17).setValue(new Date());
  }
  sh.getRange(t, 1).setNumberFormat('@');
  sh.getRange(t, 1, 1, 10).setValues([[String(r.id), date, +r.coconut, +r.water, +r.shell,
    +r.black, +r.testa, +r.diced, +r.tombong, +r.white]]);
  sh.getRange(t, 2).setNumberFormat('d-mmm');
  sh.getRange(t, 11, 1, 4).setFormulas([[
    '=IFERROR(G' + t + '/F' + t + ',"")',
    '=IFERROR(H' + t + '/F' + t + ',"")',
    '=IFERROR(I' + t + '/F' + t + ',"")',
    '=IFERROR(J' + t + '/F' + t + ',"")'
  ]]).setNumberFormat('0%');
  sh.getRange(t, 15, 1, 2).setValues([[r.worker || '', r.remarks || '']]);
  sh.getRange(t, 17).setNumberFormat('yyyy-mm-dd hh:mm');
}

function del_(sh, id) {
  const row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
}

function list_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(2, 1, last - 1, 16).getValues();
  return vals.filter(v => v[0] !== '').map(v => ({
    id: String(v[0]),
    date: (v[1] instanceof Date) ? Utilities.formatDate(v[1], tz, 'yyyy-MM-dd') : String(v[1]),
    coconut: +v[2], water: +v[3], shell: +v[4], black: +v[5],
    testa: +v[6], diced: +v[7], tombong: +v[8], white: +v[9],
    worker: String(v[14] || ''), remarks: String(v[15] || '')
  }));
}
