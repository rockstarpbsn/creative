/**
 * Jembatan web -> Google Sheet, tab "NOV" (Designer Task Tracker).
 * Pasang: Sheet > Extensions > Apps Script > tempel file ini > Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Salin URL /exec ke SHEET_URL di index.html.
 *
 * Kolom tracker (baris header = 5, data mulai baris 6):
 *  I No. | J Task Name | K Design Category | L Brand | M Deadline | N Days Remaining (rumus, tidak disentuh)
 *  O Priority | P Work Status | Q Approval Status | R Asset Link | S Technical Draw Link | T Production Status | U Designer
 *  V Kode Request (kunci, dipakai script untuk menemukan baris; boleh di-hide)
 */
const SPREADSHEET_ID = '1bRAMgZXH2lLnyAR4pMtNbDDVJxcNAeGD7AxACAZoIjc';
const SHEET_NAME = 'NOV';          // ganti ke 'DES' dst. saat bulan berganti
const HEADER_ROW = 5, FIRST_ROW = 6;
const COL = {NAME:10, CAT:11, BRAND:12, DUE:13, DAYS:14, PRIO:15, WORK:16, APPR:17, ASSET:18, TECH:19, PROD:20, DESIGNER:21, KEY:22};
const WORK = {not_started:'Not Started', in_progress:'In Progress', on_hold:'On Hold', finished:'Completed'};

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sh) return out('no-sheet');
    if (!sh.getRange(HEADER_ROW, COL.KEY).getValue()) sh.getRange(HEADER_ROW, COL.KEY).setValue('Kode Request').setFontColor('#999999');

    const footer = footerRow_(sh);                       // baris "xx Deliverables"
    const n = footer - FIRST_ROW;                        // jumlah baris data
    const keys = sh.getRange(FIRST_ROW, COL.KEY, n, 1).getValues().flat();
    const names = sh.getRange(FIRST_ROW, COL.NAME, n, 1).getValues().flat();
    let row = keys.indexOf(p.code); row = row >= 0 ? row + FIRST_ROW : -1;

    if (p.action === 'delete') { if (row > 0) clearRow_(sh, row); return out('deleted'); }

    if (p.action === 'request') {
      if (row < 0) row = emptyRow_(sh, names, footer);
      const brand = (p.brand || '').trim();
      const title = /^\s*\[/.test(p.title) ? p.title : '[' + brand + '] ' + p.title;
      sh.getRange(row, COL.NAME, 1, 4).setValues([[title, p.type, brand, toDate_(p.deadline)]]);
      sh.getRange(row, COL.DUE).setNumberFormat('yyyy-mm-dd');
      sh.getRange(row, COL.PRIO, 1, 7).setValues([[p.priority || 'Medium', 'Not Started', '', link_(p.briefLink, 'BRIEF'), '', 'Pending', '']]);
      sh.getRange(row, COL.KEY).setValue(p.code);
      return out('ok');
    }

    if (p.action === 'assign') {
      if (row < 0) return out('not-found');
      sh.getRange(row, COL.WORK, 1, 2).setValues([[WORK[p.status] || p.status, p.approval || '']]);
      if (p.assetLink) sh.getRange(row, COL.ASSET).setValue(p.assetLink);
      sh.getRange(row, COL.TECH).setValue(p.techLink || '');
      sh.getRange(row, COL.PROD, 1, 2).setValues([[p.production || 'Pending', p.pic || '']]);
      return out('ok');
    }
    return out('unknown');
  } finally { lock.releaseLock(); }
}

function footerRow_(sh) {
  const v = sh.getRange(FIRST_ROW, COL.NAME, sh.getLastRow() - FIRST_ROW + 1, 1).getValues().flat();
  const i = v.findIndex(x => /deliverables\s*$/i.test(String(x)));
  return i >= 0 ? FIRST_ROW + i : sh.getLastRow() + 1;
}
// pakai baris kosong yang sudah disiapkan; kalau habis, sisipkan baris baru di atas baris footer
function emptyRow_(sh, names, footer) {
  const i = names.findIndex(x => String(x).trim() === '');
  if (i >= 0) return FIRST_ROW + i;
  const last = footer - 1;
  sh.insertRowAfter(last - 1);                           // sisip di dalam rentang supaya rumus ringkasan ikut melebar
  const row = last;
  sh.getRange(last - 1, 1, 1, sh.getLastColumn()).copyTo(sh.getRange(row, 1, 1, sh.getLastColumn()));
  sh.getRange(row, COL.NAME, 1, 4).clearContent();
  sh.getRange(row, COL.PRIO, 1, 8).clearContent();
  return row;
}
function clearRow_(sh, row) {                            // kosongkan isi (rumus & dropdown tetap), bukan hapus baris
  sh.getRange(row, COL.NAME, 1, 4).clearContent();
  sh.getRange(row, COL.PRIO, 1, 8).clearContent();
  sh.getRange(row, COL.PROD).setValue('Pending');
}
function toDate_(s) { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, m - 1, d); }
function link_(url, label) { return url ? '=HYPERLINK("' + String(url).replace(/"/g, '""') + '","' + label + '")' : ''; }
function out(s) { return ContentService.createTextOutput(s); }
