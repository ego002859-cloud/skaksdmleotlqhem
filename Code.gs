// MY OS v3 — Apps Script Backend
// 탭: schedule / todo / log / notes (없으면 자동 생성)
// 배포: 기존 프로젝트에 붙여넣기 → 배포 관리 → 편집 → 새 버전 → 배포 (URL 유지)

const SHEETS = {
  schedule: ['id','title','date','time','category','memo','done','created_at'],
  todo:     ['id','title','category','due','priority','done','memo','created_at','done_at'],
  log:      ['id','content','category','date','created_at'],
  notes:    ['id','title','body','tags','created_at','updated_at'],
};

function getSheet(name) {
  const headers = SHEETS[name];
  if (!headers) throw new Error('unknown sheet: ' + name);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), headers.length).setNumberFormat('@'); // 날짜/시간 자동변환 방지
  }
  return s;
}

function doGet(e) {
  try {
    const data = {};
    Object.keys(SHEETS).forEach(n => data[n] = sheetToJson(n));
    return json({ ok: true, data });
  } catch (err) { return json({ ok: false, error: err.message }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, sheet, row, id } = body;
    if (action === 'add')    return json({ ok: true, id: appendRow(sheet, row || {}) });
    if (action === 'update') return json({ ok: updateRow(sheet, id, row || {}) });
    if (action === 'delete') return json({ ok: deleteRow(sheet, id) });
    return json({ ok: false, error: 'unknown action' });
  } catch (err) { return json({ ok: false, error: err.message }); }
}

function fmt(v, h) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (h === 'time') return Utilities.formatDate(v, tz, 'HH:mm');
    if (h === 'date' || h === 'due') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return v.toISOString();
  }
  return v === '' ? null : v;
}

function sheetToJson(name) {
  const rows = getSheet(name).getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = fmt(r[i], h));
    if ('done' in o) o.done = o.done === true || o.done === 'TRUE' || o.done === 'true' || o.done === 1;
    return o;
  });
}

function appendRow(name, row) {
  const s = getSheet(name);
  const headers = SHEETS[name];
  if (!row.id) row.id = Utilities.getUuid();          // 단축어에서 올 때는 id 없음
  if (!row.created_at) row.created_at = new Date().toISOString();
  if (headers.includes('done') && row.done === undefined) row.done = false;
  s.appendRow(headers.map(h => row[h] === undefined || row[h] === null ? '' : String(row[h])));
  return row.id;
}

function findRow(s, id) {
  const col = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (col[i][0] === id) return i + 1;
  return -1;
}

function updateRow(name, id, updates) {
  const s = getSheet(name), headers = SHEETS[name], r = findRow(s, id);
  if (r < 0) return false;
  headers.forEach((h, i) => { if (h in updates) s.getRange(r, i + 1).setValue(updates[h] === null ? '' : String(updates[h])); });
  return true;
}

function deleteRow(name, id) {
  const s = getSheet(name), r = findRow(s, id);
  if (r < 0) return false;
  s.deleteRow(r);
  return true;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
