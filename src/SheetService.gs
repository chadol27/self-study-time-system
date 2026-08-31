function spreadsheet_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function initializeSheets_() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = spreadsheet_();
    let roster = ss.getSheetByName(APP.SHEETS.ROSTER);
    if (!roster) {
      const first = ss.getSheets()[0];
      if (first && first.getLastRow() === 0 && first.getLastColumn() === 0) {
        first.setName(APP.SHEETS.ROSTER); roster = first;
      } else roster = ss.insertSheet(APP.SHEETS.ROSTER);
    }
    const log = ss.getSheetByName(APP.SHEETS.LOG) || ss.insertSheet(APP.SHEETS.LOG);
    const settings = ss.getSheetByName(APP.SHEETS.SETTINGS) || ss.insertSheet(APP.SHEETS.SETTINGS);
    if (roster.getLastRow() === 0) setupRoster_(roster);
    if (log.getLastRow() === 0) setupSimpleSheet_(log, APP.LOG_HEADERS);
    if (settings.getLastRow() === 0) setupSimpleSheet_(settings, APP.SETTINGS_HEADERS);
  } finally { lock.releaseLock(); }
  ensureStudentKeys_();
}

function setupRoster_(sheet) {
  sheet.getRange(1, 1, 2, 17).setValues([APP.ROSTER_HEADERS_1.slice(), APP.ROSTER_HEADERS_2.slice()]);
  sheet.setFrozenRows(2);
  sheet.hideColumns(APP.STUDENT_KEY_COL);
  sheet.getRange('A:D').setNumberFormat('@');
  sheet.getRange(1, 1, 2, 17).setFontWeight('bold').setHorizontalAlignment('center');
}

function setupSimpleSheet_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function ensureStudentKeys_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
    if (!sheet || sheet.getLastRow() < APP.ROSTER_FIRST_DATA_ROW) return;
    const count = sheet.getLastRow() - 2;
    const rows = sheet.getRange(3, 1, count, 17).getDisplayValues();
    const keys = new Set(rows.map(function (r) { return r[16]; }).filter(Boolean));
    const writes = [];
    rows.forEach(function (row, i) {
      if (row.slice(0, 16).some(function (v) { return v !== ''; }) && !row[16]) {
        let key; do { key = Utilities.getUuid(); } while (keys.has(key));
        keys.add(key); writes.push({ row: i + 3, key: key });
      }
    });
    writes.forEach(function (w) { sheet.getRange(w.row, 17).setValue(w.key); });
    if (!sheet.isColumnHiddenByUser(17)) sheet.hideColumns(17);
  } finally { lock.releaseLock(); }
}

function headersValid_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  if (!sheet || sheet.getMaxColumns() < 17 || sheet.getMaxRows() < 2) return false;
  const values = sheet.getRange(1, 1, 2, 17).getDisplayValues();
  return APP.ROSTER_HEADERS_1.every(function (v, i) { return values[0][i] === v; }) &&
    APP.ROSTER_HEADERS_2.every(function (v, i) { return values[1][i] === v; });
}

function getAttendanceColumns_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  const last = sheet.getLastColumn();
  const result = [];
  if (last < APP.ATTENDANCE_FIRST_COL) return result;
  const width = last - APP.ATTENDANCE_FIRST_COL + 1;
  const values = sheet.getRange(1, APP.ATTENDANCE_FIRST_COL, 2, width).getValues();
  for (let i = 0; i < width; i += 3) {
    const raw = values[0][i];
    const key = raw instanceof Date && !isNaN(raw.getTime()) ? dateKey_(raw) : '';
    result.push({ key: key, col: APP.ATTENDANCE_FIRST_COL + i, raw: raw, periods: values[1].slice(i, i + 3) });
  }
  return result;
}

function ensureDateColumns_(key) {
  const existing = getAttendanceColumns_();
  const found = existing.filter(function (x) { return x.key === key; });
  if (found.length === 1) return found[0].col;
  if (found.length > 1) throw userError_('날짜 출결 열이 중복되어 있습니다.', 'INVALID_HEADERS');
  if (existing.some(function (x) { return !x.key || x.periods.join('|') !== '1교시|2교시|3교시'; })) {
    throw userError_('출결 날짜 헤더 구조가 올바르지 않습니다.', 'INVALID_HEADERS');
  }
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  let col = sheet.getLastColumn() + 1;
  const later = existing.find(function (x) { return x.key > key; });
  if (later) { col = later.col; sheet.insertColumnsBefore(col, 3); }
  else if (sheet.getMaxColumns() < col + 2) sheet.insertColumnsAfter(sheet.getMaxColumns(), col + 2 - sheet.getMaxColumns());
  const date = parseDateKey_(key);
  sheet.getRange(1, col, 1, 3).setValues([[date, date, date]]).setNumberFormat('M/d').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, col, 1, 3).setValues([['1교시', '2교시', '3교시']]).setFontWeight('bold').setHorizontalAlignment('center');
  return col;
}

function readRoster_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  const width = Math.max(17, sheet.getLastColumn());
  const raw = sheet.getRange(3, 1, lastRow - 2, width).getValues();
  const display = sheet.getRange(3, 1, lastRow - 2, width).getDisplayValues();
  return raw.map(function (r, i) {
    return {
      row: i + 3, studentId: display[i][0].trim(), name: display[i][1].trim(), seatRaw: r[2],
      seat: Number(r[2]), pin: display[i][3], applications: r.slice(4, 16), key: display[i][16],
      attendance: r.slice(17), active: r.slice(4, 16).some(function (v) { return v !== ''; })
    };
  }).filter(function (s) { return s.studentId || s.name || s.key || s.active; });
}

function getClosedDates_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return new Set();
  return new Set(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function (r) {
    return r[0] instanceof Date && !isNaN(r[0].getTime()) ? dateKey_(r[0]) : '';
  }).filter(Boolean));
}

function isOperatingDate_(key) { return isWeekdayOperating_(key) && !getClosedDates_().has(key); }
function applicationIndex_(key, period) { return (weekdayIndex_(key) - 1) * 3 + period - 1; }
function isApplied_(student, key, period) { return Number(student.applications[applicationIndex_(key, period)]) === 1; }
function normalizeStatus_(value) { return value === '' || value === null || Number(value) === 0 ? '' : String(Number(value)); }

function appendAudit_(role, student, key, period, before, after) {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
  sheet.appendRow([now_(), role, student.key, student.studentId, parseDateKey_(key), period, normalizeStatus_(before), normalizeStatus_(after)]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('yyyy-MM-dd');
}

function validateAll_() {
  const errors = [];
  const config = getConfig_();
  if (!config.ok) errors.push('Script Properties 오류: ' + config.errors.join(', '));
  [APP.SHEETS.ROSTER, APP.SHEETS.LOG, APP.SHEETS.SETTINGS].forEach(function (n) { if (!spreadsheet_().getSheetByName(n)) errors.push(n + ' 시트가 없습니다.'); });
  if (!headersValid_()) errors.push('명부 시트의 고정 헤더가 올바르지 않습니다.');
  const logSheet = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
  const settingsSheet = spreadsheet_().getSheetByName(APP.SHEETS.SETTINGS);
  if (logSheet && (logSheet.getMaxColumns() < APP.LOG_HEADERS.length || logSheet.getRange(1, 1, 1, APP.LOG_HEADERS.length).getDisplayValues()[0].some(function (v, i) { return v !== APP.LOG_HEADERS[i]; }))) errors.push('기록 시트 헤더가 올바르지 않습니다.');
  if (settingsSheet && (settingsSheet.getMaxColumns() < APP.SETTINGS_HEADERS.length || settingsSheet.getRange(1, 1, 1, APP.SETTINGS_HEADERS.length).getDisplayValues()[0].some(function (v, i) { return v !== APP.SETTINGS_HEADERS[i]; }))) errors.push('설정 시트 헤더가 올바르지 않습니다.');
  if (!config.ok || !headersValid_()) return { errors: errors, excludedKeys: [] };
  const students = readRoster_();
  const excluded = new Set();
  const activeIds = {};
  const keys = {};
  students.forEach(function (s) {
    if (s.key) { if (keys[s.key]) { errors.push('내부 학생 키 중복: ' + s.studentId); excluded.add(s.key); excluded.add(keys[s.key]); } else keys[s.key] = s.key; }
    else { errors.push('내부 학생 키 누락: ' + s.studentId); excluded.add(s.key); }
    if (s.active) {
      if (!activeIds[s.studentId]) activeIds[s.studentId] = [];
      activeIds[s.studentId].push(s);
    }
    if (!/^\d{4}$/.test(s.pin)) { errors.push('전화번호 뒤 4자리 형식 오류: ' + s.studentId); excluded.add(s.key); }
    if (!Number.isInteger(s.seat) || s.seat < 1 || s.seat > config.totalSeats) { errors.push('좌석번호 범위 오류: ' + s.studentId + ' (' + s.seatRaw + ')'); excluded.add(s.key); }
    s.applications.forEach(function (v, i) { if (!(v === '' || Number(v) === 0 || Number(v) === 1)) { errors.push('신청값 오류: ' + s.studentId + ', ' + (Math.floor(i / 3) + 1) + '요일 ' + (i % 3 + 1) + '교시'); excluded.add(s.key); } });
    s.attendance.forEach(function (v) { if (!(v === '' || [0, 1, 2, 3].includes(Number(v)))) { errors.push('출결값 오류: ' + s.studentId); excluded.add(s.key); } });
  });
  Object.keys(activeIds).forEach(function (id) { if (id && activeIds[id].length > 1) { errors.push('활성 학번 중복: ' + id); activeIds[id].forEach(function (s) { excluded.add(s.key); }); } });
  for (let day = 0; day < 4; day++) for (let period = 0; period < 3; period++) {
    const seats = {};
    students.filter(function (s) { return s.active && Number(s.applications[day * 3 + period]) === 1; }).forEach(function (s) {
      if (!seats[s.seat]) seats[s.seat] = [];
      seats[s.seat].push(s);
    });
    Object.keys(seats).forEach(function (seat) { if (seats[seat].length > 1) { errors.push('좌석 중복: ' + (day + 1) + '요일 ' + (period + 1) + '교시 ' + seat + '번 (' + seats[seat].map(function (s) { return s.studentId; }).join(', ') + ')'); seats[seat].forEach(function (s) { excluded.add(s.key); }); } });
  }
  const cols = getAttendanceColumns_();
  cols.forEach(function (x) { if (!x.key || x.periods.join('|') !== '1교시|2교시|3교시') errors.push('출결 날짜/교시 헤더 구조 오류: ' + x.col + '열'); });
  for (let i = 1; i < cols.length; i++) if (cols[i - 1].key >= cols[i].key) errors.push('출결 날짜 헤더가 중복되었거나 시간순이 아닙니다.');
  return { errors: errors, excludedKeys: Array.from(excluded) };
}
