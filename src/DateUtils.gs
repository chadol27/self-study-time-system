function now_() { return new Date(); }
/** @param {*} date Apps Script accepts native Date values despite its ambient Date type. */
function dateKey_(date) { return Utilities.formatDate(date, APP.TZ, 'yyyy-MM-dd'); }
function todayKey_() { return dateKey_(now_()); }
function parseDateKey_(key) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(key))) throw userError_('날짜 형식이 올바르지 않습니다.', 'INVALID_DATE');
  const parts = String(key).split('-').map(Number);
  const noon = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 3));
  if (dateKey_(noon) !== key) throw userError_('존재하지 않는 날짜입니다.', 'INVALID_DATE');
  return noon;
}
function addDays_(key, days) { const d = parseDateKey_(key); d.setUTCDate(d.getUTCDate() + days); return dateKey_(d); }
function compareDateKeys_(a, b) { return a < b ? -1 : a > b ? 1 : 0; }
function weekdayIndex_(key) { return Number(Utilities.formatDate(parseDateKey_(key), APP.TZ, 'u')); }
function isWeekdayOperating_(key) { const day = weekdayIndex_(key); return day >= 1 && day <= 4; }
function assertFutureRange_(key, allowPast) {
  const today = todayKey_();
  if ((!allowPast && key < today) || key > addDays_(today, APP.MAX_FUTURE_DAYS)) {
    throw userError_('선택 가능한 날짜 범위를 벗어났습니다.', 'DATE_OUT_OF_RANGE');
  }
}
function currentMinutes_() { return Number(Utilities.formatDate(now_(), APP.TZ, 'H')) * 60 + Number(Utilities.formatDate(now_(), APP.TZ, 'm')); }
function timeToMinutes_(time) { const p = time.split(':').map(Number); return p[0] * 60 + p[1]; }
function defaultPeriod_(config) {
  const minutes = currentMinutes_();
  return minutes < timeToMinutes_(config.periodTimes[1]) ? 1 : minutes < timeToMinutes_(config.periodTimes[2]) ? 2 : 3;
}
function isTodayStudentClosed_(config) { return currentMinutes_() >= timeToMinutes_(config.periodTimes[0]); }
