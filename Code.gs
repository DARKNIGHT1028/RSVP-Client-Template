/**
 * Reusable RSVP + Ucapan template backend — JSONP edition.
 * Tab "Client RSVP": A Timestamp | B Nama Tetamu | C Kehadiran | D ID
 * Tab "Ucapan":      A Timestamp | B Nama        | C Ucapan    | D ID
 */
const SHEET_ID = '1aLM34s3qc0xHtS4gi2TbF48j908uDJVTFbjlw7x8ik8';
function adminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
}
const TAB = 'Client RSVP';
const HEAD = ['Timestamp', 'Nama Tetamu', 'Kehadiran', 'ID'];

// --- NEW: Ucapan (wishes) ---
const TAB_W = 'Ucapan';
const HEAD_W = ['Timestamp', 'Nama', 'Ucapan', 'ID'];
const WISH_LIMIT = 200; // max wishes returned to the public page

function sheetOf_(tab, head) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(tab);
  if (!sh) sh = ss.insertSheet(tab);
  if (sh.getRange(1, 1, 1, 4).getValues()[0].join('|') !== head.join('|')) {
    sh.getRange(1, 1, 1, 4).setValues([head]);
  }
  return sh;
}
function sheet_()  { return sheetOf_(TAB, HEAD); }
function wsheet_() { return sheetOf_(TAB_W, HEAD_W); }

function output_(data, callback) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function safe_(value) {
  const s = String(value || '');
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** Delete a row by "rowN" id or by the ID column value. */
function deleteRow_(sh, id) {
  const n = Math.max(0, sh.getLastRow() - 1);
  const ids = n ? sh.getRange(2, 4, n, 1).getValues().map(r => String(r[0])) : [];
  const m = /^row(\d+)$/.exec(id);
  if (m) {
    const rowNumber = Number(m[1]);
    if (rowNumber >= 2 && rowNumber <= sh.getLastRow()) sh.deleteRow(rowNumber);
  } else {
    const idx = ids.indexOf(id);
    if (idx >= 0) sh.deleteRow(idx + 2);
  }
}

/** GET handles JSONP requests from GitHub Pages.
 *  Actions: health, rsvp, list, delete, wish, wishes, wish_list, wish_delete */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  try {
    if (p.action === 'health') {
      return output_({ok:true, message:'RSVP template API is running.'}, cb);
    }

    /* ---------- RSVP ---------- */
    if (p.action === 'rsvp') {
      const name = String(p.name || '').trim().slice(0, 80);
      const status = String(p.status || '');
      const id = String(p.id || '');
      if (!id || !name || ['Hadir', 'Tidak hadir'].indexOf(status) < 0) return output_({ok:false, error:'invalid'}, cb);
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        sheet_().appendRow([new Date(), safe_(name), status, id]);
      } finally { lock.releaseLock(); }
      return output_({ok:true}, cb);
    }
    if (p.action === 'list') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const values = sheet_().getDataRange().getValues().slice(1);
      const rows = [];
      values.forEach((r, i) => {
        if (!r[1] && !r[2]) return;
        rows.push({id:'row' + (i + 2), at:r[0] ? new Date(r[0]).getTime() : 0, name:String(r[1]), status:String(r[2])});
      });
      rows.sort((a, b) => b.at - a.at);
      return output_({ok:true, rows:rows}, cb);
    }
    if (p.action === 'delete') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { deleteRow_(sheet_(), String(p.id || '')); } finally { lock.releaseLock(); }
      return output_({ok:true}, cb);
    }

    /* ---------- UCAPAN (NEW) ---------- */
    // Guest submits a wish:  ?action=wish&name=...&msg=...&id=...&callback=...
    if (p.action === 'wish') {
      const name = String(p.name || '').trim().slice(0, 60);
      const msg  = String(p.msg  || '').trim().slice(0, 300);
      const id   = String(p.id   || '');
      if (!id || !name || !msg) return output_({ok:false, error:'invalid'}, cb);
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        wsheet_().appendRow([new Date(), safe_(name), safe_(msg), id]);
      } finally { lock.releaseLock(); }
      return output_({ok:true}, cb);
    }
    // Public list shown in "Senarai Ucapan" (no key needed, newest first)
    if (p.action === 'wishes') {
      const values = wsheet_().getDataRange().getValues().slice(1);
      const rows = [];
      values.forEach(r => {
        if (!r[1] || !r[2]) return;
        rows.push({at: r[0] ? new Date(r[0]).getTime() : 0, n: String(r[1]), m: String(r[2])});
      });
      rows.sort((a, b) => b.at - a.at);
      return output_({ok:true, rows: rows.slice(0, WISH_LIMIT)}, cb);
    }
    // Admin list (with row ids, for moderation)
    if (p.action === 'wish_list') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const values = wsheet_().getDataRange().getValues().slice(1);
      const rows = [];
      values.forEach((r, i) => {
        if (!r[1] && !r[2]) return;
        rows.push({id:'row' + (i + 2), at:r[0] ? new Date(r[0]).getTime() : 0, name:String(r[1]), msg:String(r[2])});
      });
      rows.sort((a, b) => b.at - a.at);
      return output_({ok:true, rows:rows}, cb);
    }
    // Admin delete a wish
    if (p.action === 'wish_delete') {
      if (p.key !== adminKey_()) return output_({ok:false, error:'auth'}, cb);
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try { deleteRow_(wsheet_(), String(p.id || '')); } finally { lock.releaseLock(); }
      return output_({ok:true}, cb);
    }

    return output_({ok:true, message:'RSVP template API is running.'}, cb);
  } catch (err) { return output_({ok:false, error:'server'}, cb); }
}

/** Retained for compatibility with direct POST clients. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'rsvp') {
      const name = String(b.name || '').trim().slice(0, 80);
      if (!b.id || !name || ['Hadir', 'Tidak hadir'].indexOf(b.status) < 0) return output_({ok:false, error:'invalid'});
      sheet_().appendRow([new Date(), safe_(name), b.status, String(b.id)]);
      return output_({ok:true});
    }
    if (b.action === 'wish') {
      const name = String(b.name || '').trim().slice(0, 60);
      const msg  = String(b.msg  || '').trim().slice(0, 300);
      if (!b.id || !name || !msg) return output_({ok:false, error:'invalid'});
      wsheet_().appendRow([new Date(), safe_(name), safe_(msg), String(b.id)]);
      return output_({ok:true});
    }
    if (b.action === 'delete') {
      if (b.key !== adminKey_()) return output_({ok:false, error:'auth'});
      deleteRow_(sheet_(), String(b.id || ''));
      return output_({ok:true});
    }
    if (b.action === 'wish_delete') {
      if (b.key !== adminKey_()) return output_({ok:false, error:'auth'});
      deleteRow_(wsheet_(), String(b.id || ''));
      return output_({ok:true});
    }
    return output_({ok:false, error:'unknown'});
  } catch (err) { return output_({ok:false, error:'server'}); }
  finally { lock.releaseLock(); }
}

/** Run this from Apps Script. The script owner needs Editor access to the spreadsheet. */
function testConnection() {
  const sh = sheet_();
  const wh = wsheet_(); // creates the "Ucapan" tab if missing
  Logger.log('Spreadsheet connected successfully.');
  Logger.log('RSVP sheet: ' + sh.getName() + ' | last row: ' + sh.getLastRow());
  Logger.log('Ucapan sheet: ' + wh.getName() + ' | last row: ' + wh.getLastRow());
}
