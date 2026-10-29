// Timesheet backend: one styled tab per month ("Oct 2026") plus a "Masters" tab.
// Paste into Extensions > Apps Script, then Deploy > Manage deployments > Edit > New version.
// Your old single "Timesheet" tab is moved into month tabs automatically (and kept as "Timesheet (backup)").

const HEAD = ['ID','SL.no','Month','Date','Code','Source','Product','Transaction Start Time','Transaction End time','Total Duration','CONTACT NO','Business Name','Issue/Request','Ticket No','Status','Remarks'];
const KEYS = ['id','sl','month','date','code','src','prod','s','e','dur','contact','biz','issue','ticket','stat','rem'];
const WIDTHS = [40,48,78,125,64,84,110,70,70,76,110,190,100,86,100,330];
const ALIGN  = ['c','c','c','c','c','c','l','c','c','c','c','l','c','c','c','l'];
const HDR = 4, FIRST = 5, N = HEAD.length;          // title rows 1-2, header row 4, data from row 5
const MHEAD = ['Type','Name','Code'];
const MON  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const LONG = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DEFAULTS = { stat: ['Open','In Progress','Resolved','Closed'], src: ['Call','Email','WhatsApp','Chat'] };
const NAVY = '#1f3864', BLUE = '#2f5597', PALE = '#eaf0fa', LINE = '#d0d7e2';

// ---------- helpers ----------
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function str_(v, n) { return String(v == null ? '' : v).slice(0, n || 2000); }
function v_(v, k) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), k === 'date' ? 'yyyy-MM-dd' : 'HH:mm');
  return String(v);
}
function isMonth_(n) { return /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}$/.test(n); }
function monthKey_(n) { const p = n.split(' '); return +p[1] * 12 + MON.indexOf(p[0]); }
function monthName_(d) { const p = String(d).split('-'); return MON[+p[1] - 1] + ' ' + p[0]; }
function monthSheets_() {
  return SpreadsheetApp.getActive().getSheets().filter(s => isMonth_(s.getName())).sort((a, b) => monthKey_(a.getName()) - monthKey_(b.getName()));
}

// ---------- styling ----------
function style_(sh, name) {
  const p = name.split(' '), title = 'Timesheet  |  ' + LONG[MON.indexOf(p[0])] + ' ' + p[1];
  sh.getRange(HDR, 1, 1, N).setValues([HEAD]);                         // header first, so data rows are safe even if styling fails
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  const rows = sh.getMaxRows() - FIRST + 1;
  step(() => sh.setHiddenGridlines(true));
  step(() => sh.getRange(1, 1, sh.getMaxRows(), N).setFontFamily('Calibri').setFontSize(10).setVerticalAlignment('top'));
  step(() => sh.getRange(FIRST, 1, rows, N).setNumberFormat('@'));      // keep dates and times as typed
  step(() => { const t = sh.getRange(1, 2, 1, N - 1).merge(); t.setValue(title).setBackground(NAVY).setFontColor('#ffffff').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle'); sh.setRowHeight(1, 38); });
  step(() => {
    sh.getRange(2, 2, 1, N - 1).setBackground(PALE).setVerticalAlignment('middle');
    sh.setRowHeight(2, 28);
    const lab = (c, t) => sh.getRange(2, c).setValue(t).setFontColor('#5b6b86').setFontWeight('bold').setHorizontalAlignment('right');
    const val = (c, f) => sh.getRange(2, c).setFormula(f).setFontColor(NAVY).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('left');
    lab(2, 'Entries');    val(3, '=COUNTA(D5:D)');
    lab(5, 'Total time'); val(6, '=LET(m,ARRAYFORMULA(SUM(IFERROR(VALUE(LEFT(J5:J,FIND(":",J5:J)-1)),0)*60+IFERROR(VALUE(MID(J5:J,FIND(":",J5:J)+1,2)),0))),INT(m/60)&"h "&TEXT(MOD(m,60),"00")&"m")');
    lab(9, 'Resolved');   val(10, '=COUNTIF(O5:O,"*esolv*")+COUNTIF(O5:O,"*losed*")');
    lab(12, 'Open');      val(13, '=COUNTIF(O5:O,"Open")');
    sh.setRowHeight(3, 8);
  });
  step(() => {
    sh.getRange(HDR, 2, 1, N - 1).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
      .setBorder(true, true, true, true, true, false, LINE, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(HDR, 34);
  });
  step(() => { WIDTHS.forEach((w, i) => sh.setColumnWidth(i + 1, w)); sh.hideColumns(1); });
  step(() => ALIGN.forEach((a, i) => { if (i) sh.getRange(FIRST, i + 1, rows, 1).setHorizontalAlignment(a === 'c' ? 'center' : 'left'); }));
  step(() => { sh.getRange(FIRST, 12, rows, 1).setWrap(true); sh.getRange(FIRST, 16, rows, 1).setWrap(true); });
  step(() => sh.setFrozenRows(HDR));
  step(() => {
    const st = sh.getRange(FIRST, 15, rows, 1), band = sh.getRange(FIRST, 2, rows, N - 1);
    const R = SpreadsheetApp.newConditionalFormatRule;
    sh.setConditionalFormatRules([
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"resolv|clos|done|complet")').setBackground('#c6efce').setFontColor('#006100').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"progress|pending|hold|wait|wip|working")').setBackground('#ffeb9c').setFontColor('#7f6000').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=REGEXMATCH(LOWER($O5),"open|new")').setBackground('#ffc7ce').setFontColor('#9c0006').setBold(true).setRanges([st]).build(),
      R().whenFormulaSatisfied('=AND($D5<>"",ISEVEN(ROW()))').setBackground('#f6f8fc').setRanges([band]).build()
    ]);
  });
  step(() => sh.setTabColor(NAVY));
}
function styleMasters_(sh) {
  const step = f => { try { f(); } catch (e) { Logger.log('style: ' + e); } };
  step(() => sh.getRange(1, 1, 1, 3).setBackground(BLUE).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center'));
  step(() => { [110, 230, 380].forEach((w, i) => sh.setColumnWidth(i + 1, w)); sh.setFrozenRows(1); sh.setTabColor('#f59e0b'); });
  step(() => sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setFontColor('#6b7280').setFontSize(9));
}

// ---------- sheets ----------
function monthSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  try {
    const earlier = monthSheets_().filter(s => s.getName() !== name && monthKey_(s.getName()) < monthKey_(name)).length;
    ss.setActiveSheet(sh); ss.moveActiveSheet(earlier + 1);
  } catch (e) {}
  style_(sh, name);
  return sh;
}
function mastersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Masters');
  if (!sh) {
    sh = ss.insertSheet('Masters');
    sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([MHEAD]);
    styleMasters_(sh);
  }
  return sh;
}
function find_(id) {
  for (const sh of monthSheets_()) {
    const n = sh.getLastRow() - HDR;
    if (n < 1) continue;
    const ids = sh.getRange(FIRST, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return { sh: sh, row: FIRST + i };
  }
  return null;
}
function renumber_(sh) {
  const n = sh.getLastRow() - HDR;
  if (n > 0) sh.getRange(FIRST, 2, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
}
function ensureRows_(sh, rowNo) {
  const need = rowNo - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), Math.max(need, 100));
}
function border_(sh, row, rows) {
  try { sh.getRange(row, 2, rows, N - 1).setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID); } catch (e) {}
}

// ---------- data ----------
function upsert_(r) {
  if (!r || !r.id || !/^\d{4}-\d{2}-\d{2}$/.test(String(r.date || ''))) throw 'Invalid entry.';
  const id = str_(r.id, 40), target = monthName_(r.date), found = find_(id);
  let sh, rowNo;
  if (found && found.sh.getName() === target) { sh = found.sh; rowNo = found.row; }
  else {
    if (found) { found.sh.deleteRow(found.row); renumber_(found.sh); }   // the date moved to another month
    sh = monthSheet_(target); rowNo = Math.max(sh.getLastRow() + 1, FIRST);
  }
  const row = {};
  KEYS.forEach(k => row[k] = str_(r[k]));
  row.id = id; row.sl = '';
  row.month = Utilities.formatDate(new Date(r.date + 'T00:00:00'), tz_(), 'MMM yyyy');
  ensureRows_(sh, rowNo);
  sh.getRange(rowNo, 1, 1, N).setNumberFormat('@').setValues([KEYS.map(k => row[k])]);
  border_(sh, rowNo, 1);
  renumber_(sh);
}
function del_(id) {
  const f = find_(String(id));
  if (f) { f.sh.deleteRow(f.row); renumber_(f.sh); }
}
function masters_() {
  const ms = mastersSheet_();
  if (ms.getLastRow() < 2) {
    const seed = [];
    for (const t in DEFAULTS) DEFAULTS[t].forEach(n => seed.push([t, n, '']));
    ms.getRange(2, 1, seed.length, 3).setValues(seed);
  }
  const m = { prod: [], biz: [], stat: [], src: [], tpl: [] };
  ms.getRange(2, 1, ms.getLastRow() - 1, 3).getValues().forEach(r => {
    const t = r[0]; if (!m[t]) return;
    let x = String(r[1]);
    if (t === 'biz') x = { name: x, code: String(r[2]) };
    else if (t === 'tpl') { try { x = JSON.parse(String(r[2])); } catch (e) { return; } }
    m[t].push(x);
  });
  return m;
}
function setMasters_(masters) {
  const ms = mastersSheet_(), rows = [];
  ['prod', 'biz', 'stat', 'src', 'tpl'].forEach(t => (masters && masters[t] || []).slice(0, 2000).forEach(x => {
    if (t === 'biz') rows.push([t, str_(x.name, 200), str_(x.code, 30)]);
    else if (t === 'tpl') rows.push([t, str_(x.name, 100), JSON.stringify(x).slice(0, 4000)]);
    else rows.push([t, str_(x, 200), '']);
  }));
  if (ms.getLastRow() > 1) ms.getRange(2, 1, ms.getLastRow() - 1, 3).clearContent();
  if (rows.length) ms.getRange(2, 1, rows.length, 3).setValues(rows);
}

// Moves the old single "Timesheet" tab into month tabs (runs once, keeps the old tab as a backup).
function migrate_() {
  const ss = SpreadsheetApp.getActive(), old = ss.getSheetByName('Timesheet');
  if (!old) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const n = old.getLastRow() - 1;
    if (n > 0) {
      const head = old.getRange(1, 1, 1, old.getLastColumn()).getValues()[0].map(String);
      const idx = HEAD.map(h => head.indexOf(h));
      const groups = {};
      old.getRange(2, 1, n, old.getLastColumn()).getValues().forEach(r => {
        const o = {};
        KEYS.forEach((k, i) => o[k] = idx[i] >= 0 ? v_(r[idx[i]], k) : '');
        if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date)) return;
        o.month = MON[+o.date.slice(5, 7) - 1] + ' ' + o.date.slice(0, 4);
        (groups[monthName_(o.date)] = groups[monthName_(o.date)] || []).push(o);
      });
      Object.keys(groups).forEach(name => {
        const sh = monthSheet_(name), have = new Set(sh.getLastRow() >= FIRST ? sh.getRange(FIRST, 1, sh.getLastRow() - HDR, 1).getValues().map(x => String(x[0])) : []);
        const add = groups[name].filter(o => o.id && !have.has(o.id));
        if (!add.length) return;
        const start = Math.max(sh.getLastRow() + 1, FIRST);
        ensureRows_(sh, start + add.length - 1);
        sh.getRange(start, 1, add.length, N).setNumberFormat('@').setValues(add.map(o => KEYS.map(k => o[k])));
        border_(sh, start, add.length);
        renumber_(sh);
      });
    }
    old.setName('Timesheet (backup)');
    try { old.setTabColor('#9aa3b5'); } catch (e) {}
  } finally { lock.releaseLock(); }
}

// Run this from the editor if you ever want to re-apply the styling to all month tabs.
function restyleAll() {
  monthSheets_().forEach(sh => style_(sh, sh.getName()));
  styleMasters_(mastersSheet_());
}
function migrateToMonthly() { migrate_(); }

// ---------- web app ----------
function load_() {
  if (SpreadsheetApp.getActive().getSheetByName('Timesheet')) migrate_();
  const m = masters_(), rows = [];
  monthSheets_().forEach(sh => {
    const n = sh.getLastRow() - HDR;
    if (n < 1) return;
    sh.getRange(FIRST, 1, n, N).getValues().forEach(r => {
      if (!String(r[0])) return;
      const o = {}; KEYS.forEach((k, i) => o[k] = v_(r[i], k)); rows.push(o);
    });
  });
  const cur = monthSheet_(monthName_(Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd')));
  return { masters: m, rows: rows, link: SpreadsheetApp.getActive().getUrl() + '#gid=' + cur.getSheetId() };
}
function doGet() {
  try { return out_(load_()); } catch (err) { return out_({ error: String(err) }); }
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'upsertRow') upsert_(b.row);
    else if (b.action === 'deleteRow') del_(b.id);
    else if (b.action === 'setMasters') setMasters_(b.masters);
    return out_({ ok: true });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
