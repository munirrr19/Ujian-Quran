// Jambatan antara laman web (GitHub Pages) dan Google Sheet.
// Tampal dalam Extensions > Apps Script pada spreadsheet, tukar PASSCODE, kemudian Deploy > Web app.
const PASSCODE = 'TUKAR-KOD-INI';
const SHEET = 'ISI MARKAH';
const WSHEET = 'BERTULIS'; // markah mentah bertulis /100 di kolum D
const FIRST_ROW = 5; // baris pelajar pertama

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function wMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(WSHEET);
  const m = {};
  if (!sh) return m;
  const n = sh.getLastRow() - FIRST_ROW + 1;
  if (n < 1) return m;
  sh.getRange(FIRST_ROW, 2, n, 3).getValues().forEach(r => { if (r[0]) m[String(r[0])] = r[2]; });
  return m;
}

// Baca senarai pelajar, markah syafawi & markah bertulis
function doGet(e) {
  if ((e.parameter.pass || '') !== PASSCODE) return out_({ error: 'Kod salah' });
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const n = sh.getLastRow() - FIRST_ROW + 1;
  const v = sh.getRange(FIRST_ROW, 1, n, 17).getValues();
  const w = wMap_();
  const students = v.filter(r => r[1] && r[2]).map(r => ({
    no: String(r[1]), nama: String(r[2]),
    h: r.slice(3, 11), tj: r[12], fs: r[13], cat: r[16] || '',
    w: w[String(r[1])] === undefined ? '' : w[String(r[1])]
  }));
  return out_({ students });
}

// Simpan markah (hanya kolum input; formula jumlah tak disentuh)
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (d.pass !== PASSCODE) return out_({ error: 'Kod salah' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET);
    const n = sh.getLastRow() - FIRST_ROW + 1;

    if (d.action === 'clear') { // kosongkan semua markah (kolum input sahaja)
      const ids2 = sh.getRange(FIRST_ROW, 2, n, 1).getValues();
      const rows = ids2.map((r, k) => r[0] ? FIRST_ROW + k : 0).filter(Boolean);
      if (rows.length) {
        const cnt = Math.max.apply(null, rows) - FIRST_ROW + 1;
        sh.getRange(FIRST_ROW, 4, cnt, 8).clearContent();  // D:K
        sh.getRange(FIRST_ROW, 13, cnt, 2).clearContent(); // M:N
        sh.getRange(FIRST_ROW, 17, cnt, 1).clearContent(); // Q
        const ws = ss.getSheetByName(WSHEET);
        if (ws) ws.getRange(FIRST_ROW, 4, cnt, 1).clearContent(); // BERTULIS D
      }
      return out_({ ok: true });
    }

    if (d.action === 'written') { // simpan markah mentah bertulis
      const ws = ss.getSheetByName(WSHEET);
      const wn = ws.getLastRow() - FIRST_ROW + 1;
      const wids = ws.getRange(FIRST_ROW, 2, wn, 1).getValues().map(r => String(r[0]));
      const wi = wids.indexOf(String(d.no));
      if (wi < 0) return out_({ error: 'Pelajar tidak dijumpai' });
      ws.getRange(FIRST_ROW + wi, 4).setValue(d.w === '' || d.w == null ? '' : Number(d.w));
      return out_({ ok: true });
    }

    const ids = sh.getRange(FIRST_ROW, 2, n, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(d.no));
    if (i < 0) return out_({ error: 'Pelajar tidak dijumpai' });
    const row = FIRST_ROW + i;
    sh.getRange(row, 4, 1, 8).setValues([d.h]);           // D:K hafazan 1-8
    sh.getRange(row, 13, 1, 2).setValues([[d.tj, d.fs]]); // M:N tajwid, fasahah
    sh.getRange(row, 17).setValue(d.cat || '');           // Q catatan
    return out_({ ok: true });
  } finally { lock.releaseLock(); }
}
