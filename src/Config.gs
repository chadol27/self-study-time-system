const APP = Object.freeze({
  TZ: 'Asia/Seoul',
  SHEETS: Object.freeze({ ROSTER: '명부', LOG: '기록', SETTINGS: '설정' }),
  ROSTER_FIRST_DATA_ROW: 3,
  ATTENDANCE_FIRST_COL: 18,
  STUDENT_KEY_COL: 17,
  MAX_FUTURE_DAYS: 30,
  PERIODS: 3,
  REQUIRED_PROPERTIES: Object.freeze([
    'TEACHER_PASSWORD', 'TOTAL_SEATS', 'PERIOD_1_START_TIME',
    'PERIOD_2_START_TIME', 'PERIOD_3_START_TIME'
  ]),
  ROSTER_HEADERS_1: Object.freeze([
    '학번', '이름', '좌석번호', '전화번호 뒤 4자리',
    '월요일', '', '', '화요일', '', '', '수요일', '', '', '목요일', '', '', '내부 학생 키'
  ]),
  ROSTER_HEADERS_2: Object.freeze([
    '학번', '이름', '좌석번호', '전화번호 뒤 4자리',
    '1교시', '2교시', '3교시', '1교시', '2교시', '3교시',
    '1교시', '2교시', '3교시', '1교시', '2교시', '3교시', '내부 학생 키'
  ]),
  LOG_HEADERS: Object.freeze(['변경 시각', '역할', '내부 학생 키', '학번', '대상 날짜', '교시', '이전 상태', '새 상태']),
  SETTINGS_HEADERS: Object.freeze(['날짜', '사유'])
});

function getConfig_() {
  const raw = PropertiesService.getScriptProperties().getProperties();
  const errors = [];
  APP.REQUIRED_PROPERTIES.forEach(function (key) {
    if (!Object.prototype.hasOwnProperty.call(raw, key) || String(raw[key]).trim() === '') errors.push(key);
  });
  const seats = Number(raw.TOTAL_SEATS);
  if (raw.TOTAL_SEATS && (!Number.isInteger(seats) || seats < 1 || seats > 500)) errors.push('TOTAL_SEATS');
  const times = [raw.PERIOD_1_START_TIME, raw.PERIOD_2_START_TIME, raw.PERIOD_3_START_TIME];
  times.forEach(function (v, i) { if (v && !/^([01]\d|2[0-3]):[0-5]\d$/.test(v)) errors.push('PERIOD_' + (i + 1) + '_START_TIME'); });
  if (times.every(Boolean) && !(times[0] < times[1] && times[1] < times[2])) {
    errors.push('PERIOD_START_TIME_ORDER');
  }
  return {
    ok: errors.length === 0,
    errors: Array.from(new Set(errors)),
    teacherPassword: raw.TEACHER_PASSWORD || '',
    totalSeats: seats,
    periodTimes: times
  };
}

function requireConfig_() {
  const config = getConfig_();
  if (!config.ok) throw userError_('필수 운영 설정이 올바르지 않습니다.', 'CONFIG_INVALID', { keys: config.errors });
  return config;
}

function userError_(message, code, details) {
  /** @type {Error & { code?: string, details?: * }} */
  const error = new Error(message);
  error.name = 'UserError';
  error.code = code || 'INVALID_REQUEST';
  error.details = details || null;
  return error;
}

function publicCall_(callback) {
  try {
    const value = callback();
    return { ok: true, data: value };
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return {
      ok: false,
      error: {
        code: error && error.name === 'UserError' ? error.code : 'INTERNAL_ERROR',
        message: error && error.name === 'UserError' ? error.message : '처리 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.',
        details: error && error.name === 'UserError' ? error.details : null
      }
    };
  }
}
