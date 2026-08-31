function teacherBootstrap(token) {
  return publicCall_(function () {
    const config = requireConfig_();
    requireTeacher_(token);
    initializeSheets_();
    return {
      today: todayKey_(),
      maxDate: addDays_(todayKey_(), 30),
      defaultDate: latestOperatingDate_(),
      defaultPeriod: defaultPeriod_(config),
      totalSeats: config.totalSeats,
    };
  });
}
function latestOperatingDate_() {
  let key = todayKey_();
  for (let i = 0; i < 370; i++, key = addDays_(key, -1))
    if (isOperatingDate_(key)) return key;
  return todayKey_();
}
function teacherSchedule(token) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    const historical = getAttendanceColumns_()
      .map(function (x) {
        return x.key;
      })
      .filter(function (key) {
        return key && isOperatingDate_(key);
      });
    const future = [];
    for (
      let key = todayKey_(), max = addDays_(key, 30);
      key <= max;
      key = addDays_(key, 1)
    )
      if (isOperatingDate_(key)) future.push(key);
    return Array.from(new Set(historical.concat(future))).sort();
  });
}
function getTeacherSeats(token, key, period) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    return withWriteLock_(function () {
      requireTeacher_(token);
      return teacherSeats_(String(key), Number(period));
    });
  });
}
function teacherSeats_(key, period) {
  const config = requireConfig_();
  if (!Number.isInteger(period) || period < 1 || period > 3)
    throw userError_("교시가 올바르지 않습니다.", "INVALID_PERIOD");
  parseDateKey_(key);
  if (!isOperatingDate_(key))
    throw userError_("미운영일은 조회할 수 없습니다.", "CLOSED_DATE");
  if (key > addDays_(todayKey_(), 30))
    throw userError_(
      "조회 가능한 미래 날짜를 벗어났습니다.",
      "DATE_OUT_OF_RANGE",
    );
  let info = getAttendanceColumns_().find(function (x) {
    return x.key === key;
  });
  if (!info && key >= todayKey_())
    info = { key: key, col: ensureDateColumns_(key) };
  const report = validateAll_();
  const excluded = new Set(report.excludedKeys);
  const seats = /** @type {any[]} */ (
    Array.from({ length: config.totalSeats }, function (_, i) {
      return { seat: i + 1, student: null, status: "empty", label: "미신청" };
    })
  );
  readRoster_()
    .filter(function (s) {
      return s.active && !excluded.has(s.key) && isApplied_(s, key, period);
    })
    .forEach(function (s) {
      const raw = info
        ? normalizeStatus_(attendanceCell_(s, info.col, period).getValue())
        : "";
      const status =
        raw === "2"
          ? "student-absence"
          : raw === "3"
            ? "teacher-absence"
            : "present";
      const label =
        raw === "2"
          ? "사전 결석"
          : raw === "3"
            ? "교사 지정 결석"
            : "출석 대상";
      seats[s.seat - 1] = {
        seat: s.seat,
        student: { key: s.key, studentId: s.studentId, name: s.name },
        status: status,
        label: label,
        appliedPeriods: [1, 2, 3].filter(function (p) {
          return isApplied_(s, key, p);
        }),
      };
    });
  return {
    date: key,
    period: period,
    seats: seats,
    errors: report.errors,
    totalSeats: config.totalSeats,
  };
}
function latestTeacherBefore_(student, key, period) {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.LOG);
  if (sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const rowKey = rows[i][4] instanceof Date ? dateKey_(rows[i][4]) : "";
    if (
      rows[i][1] === "교사" &&
      String(rows[i][2]) === student.key &&
      rowKey === key &&
      Number(rows[i][5]) === period &&
      normalizeStatus_(rows[i][7]) === "3"
    )
      return normalizeStatus_(rows[i][6]);
  }
  return null;
}
function teacherSetAbsence(token, studentKey, key, period) {
  return teacherChange_(token, studentKey, key, period, true);
}
function teacherRestore(token, studentKey, key, period) {
  return teacherChange_(token, studentKey, key, period, false);
}
function teacherChange_(token, studentKey, key, period, absent) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    key = String(key);
    period = Number(period);
    parseDateKey_(key);
    assertFutureRange_(key, true);
    if (!isOperatingDate_(key))
      throw userError_("미운영일은 변경할 수 없습니다.", "CLOSED_DATE");
    if (!Number.isInteger(period) || period < 1 || period > 3)
      throw userError_("교시가 올바르지 않습니다.", "INVALID_PERIOD");
    return withWriteLock_(function () {
      requireTeacher_(token);
      const student = readRoster_().find(function (s) {
        return s.active && s.key === studentKey;
      });
      if (!student) throw userError_("학생을 찾을 수 없습니다.", "NOT_FOUND");
      const col = ensureDateColumns_(key);
      let changed = 0;
      const audits = [];
      const range = attendanceCell_(student, col, 1).offset(0, 0, 1, 3);
      const values = range.getValues()[0];
      for (let p = period; p <= 3; p++)
        if (isApplied_(student, key, p)) {
          const current = normalizeStatus_(values[p - 1]);
          let next = current;
          if (absent) next = "3";
          else {
            if (current === "3") {
              const prior = latestTeacherBefore_(student, key, p);
              const fallback = wasAutoProcessed_(key) ? "1" : "";
              next = prior === null || prior === "3" ? fallback : prior;
            }
          }
          if (current !== normalizeStatus_(next)) {
            values[p - 1] = next === "" ? "" : Number(next);
            audits.push([
              now_(),
              "교사",
              student.key,
              student.studentId,
              parseDateKey_(key),
              p,
              current,
              normalizeStatus_(next),
            ]);
            changed++;
          }
        }
      if (changed) range.setValues([values]);
      appendAudits_(audits);
      return { changed: changed, view: teacherSeats_(key, period) };
    });
  });
}
