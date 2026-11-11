/**
 * Jembatan web -> Google Sheet.
 * Tempel di Sheet: Extensions > Apps Script, lalu Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Salin URL /exec ke SHEET_URL di index.html.
 */
const SPREADSHEET_ID = '1bRAMgZXH2lLnyAR4pMtNbDDVJxcNAeGD7AxACAZoIjc';
const SHEET_NAME = 'Requests';
const HEAD = ['Kode','Dikirim','Requester','Divisi','Judul','Jenis design','Channel','Output','Link brief',
              'Deadline','Prioritas','Catatan','PIC','Status','Catatan supervisor','Update terakhir'];
const STATUS = {not_started:'Not started', in_progress:'On progress', finished:'Finished'};

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); sh.getRange(1,1,1,HEAD.length).setFontWeight('bold'); }
    const codes = sh.getRange(2, 1, Math.max(sh.getLastRow()-1, 1), 1).getValues().flat();
    let row = codes.indexOf(p.code); row = row >= 0 ? row + 2 : -1;

    if (p.action === 'delete') { if (row > 0) sh.deleteRow(row); return out('deleted'); }

    if (p.action === 'request') {
      const vals = [p.code, p.createdAt, p.requester, p.division, p.title, p.type, p.channel, p.output, p.briefLink,
                    p.deadline, p.priority, p.notes, '', STATUS.not_started, '', ''];
      if (row > 0) sh.getRange(row, 1, 1, 12).setValues([vals.slice(0, 12)]); else sh.appendRow(vals);
      return out('ok');
    }

    if (p.action === 'assign') {
      if (row < 0) return out('not-found');
      sh.getRange(row, 13, 1, 4).setValues([[p.pic || '', STATUS[p.status] || p.status, p.note || '', p.updatedAt || '']]);
      return out('ok');
    }
    return out('unknown');
  } finally { lock.releaseLock(); }
}
function out(s){ return ContentService.createTextOutput(s); }
