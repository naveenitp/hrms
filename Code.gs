// Timesheet backend. Paste into Extensions > Apps Script of your Google Sheet.
const TS = 'Timesheet', MS = 'Masters';
const HEAD = ['ID','SL.no','Month','Date','Code','Source','Product','Transaction Start Time','Transaction End time','Total Duration','CONTACT NO','Business Name','Issue/Request','Ticket No','Status','Remarks'];
const KEYS = ['id','sl','month','date','code','src','prod','s','e','dur','contact','biz','issue','ticket','stat','rem'];
const DEFAULTS = { stat: ['Open','In Progress','Resolved','Closed'], src: ['Call','Email','WhatsApp','Chat'] };

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@'); // keep dates/times as text
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ids_(ts) {
  return ts.getLastRow() > 1 ? ts.getRange(2, 1, ts.getLastRow() - 1, 1).getValues().map(x => String(x[0])) : [];
}
function renumber_(ts) {
  const n = ts.getLastRow() - 1;
  if (n > 0) ts.getRange(2, 2, n, 1).setValues(Array.from({ length: n }, (_, i) => [i + 1]));
}

function doGet() {
  try {
    const ts = sheet_(TS, HEAD), ms = sheet_(MS, ['Type', 'Name', 'Code']);
    if (ms.getLastRow() < 2) {
      const seed = [];
      for (const t in DEFAULTS) DEFAULTS[t].forEach(n => seed.push([t, n, '']));
      ms.getRange(2, 1, seed.length, 3).setValues(seed);
    }
    const m = { prod: [], biz: [], stat: [], src: [] };
    ms.getRange(2, 1, ms.getLastRow() - 1, 3).getValues().forEach(r => {
      if (m[r[0]]) m[r[0]].push(r[0] === 'biz' ? { name: String(r[1]), code: String(r[2]) } : String(r[1]));
    });
    const n = ts.getLastRow() - 1;
    const rows = n > 0 ? ts.getRange(2, 1, n, HEAD.length).getValues().map(r => {
      const o = {}; KEYS.forEach((k, i) => o[k] = String(r[i])); return o;
    }) : [];
    return out_({ masters: m, rows: rows });
  } catch (err) { return out_({ error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents), ts = sheet_(TS, HEAD);
    if (b.action === 'upsertRow') {
      const r = b.row, i = ids_(ts).indexOf(r.id);
      r.month = Utilities.formatDate(new Date(r.date + 'T00:00:00'), Session.getScriptTimeZone(), 'MMM yyyy');
      r.sl = '';
      ts.getRange(i < 0 ? ts.getLastRow() + 1 : i + 2, 1, 1, HEAD.length).setValues([KEYS.map(k => r[k] == null ? '' : r[k])]);
      renumber_(ts);
    } else if (b.action === 'deleteRow') {
      const i = ids_(ts).indexOf(b.id);
      if (i >= 0) ts.deleteRow(i + 2);
      renumber_(ts);
    } else if (b.action === 'setMasters') {
      const ms = sheet_(MS, ['Type', 'Name', 'Code']), rows = [];
      for (const t in b.masters) b.masters[t].forEach(x => rows.push(t === 'biz' ? [t, x.name, x.code] : [t, x, '']));
      if (ms.getLastRow() > 1) ms.getRange(2, 1, ms.getLastRow() - 1, 3).clearContent();
      if (rows.length) ms.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
