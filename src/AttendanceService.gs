function withWriteLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}
function attendanceCell_(student, col, period) {
  return spreadsheet_()
    .getSheetByName(APP.SHEETS.ROSTER)
    .getRange(student.row, col + period - 1);
}
function wasAutoProcessed_(key) {
  if (key < todayKey_()) return true;
  if (key > todayKey_()) return false;
  return (
    PropertiesService.getScriptProperties().getProperty(
      "AUTO_PROCESSED_" + key,
    ) === "1"
  );
}
function processOperatingDate_(key) {
  if (wasAutoProcessed_(key)) return { date: key, processed: 0 };
  if (!isOperatingDate_(key)) return { date: key, processed: 0 };
  return withWriteLock_(function () {
    if (wasAutoProcessed_(key)) return { date: key, processed: 0 };
    return processOperatingDateUnlocked_(key);
  });
}
function processOperatingDateUnlocked_(key) {
  if (wasAutoProcessed_(key)) return { date: key, processed: 0 };
  const col = ensureDateColumns_(key);
  let count = 0;
  readRoster_()
    .filter(function (s) {
      return s.active;
    })
    .forEach(function (s) {
      for (let p = 1; p <= 3; p++)
        if (isApplied_(s, key, p)) {
          const cell = attendanceCell_(s, col, p);
          const status = normalizeStatus_(cell.getValue());
          if (status === "") {
            cell.setValue(1);
            count++;
          }
        }
    });
  PropertiesService.getScriptProperties().setProperty(
    "AUTO_PROCESSED_" + key,
    "1",
  );
  return { date: key, processed: count };
}
function syncFutureAttendanceColumns_() {
  const today = todayKey_();
  const max = addDays_(today, APP.MAX_FUTURE_DAYS);
  return withWriteLock_(function () {
    const closedDates = getClosedDates_();
    const operatingDates = [];
    for (let key = today; key <= max; key = addDays_(key, 1)) {
      if (isWeekdayOperating_(key) && !closedDates.has(key))
        operatingDates.push(key);
    }

    const columns = getAttendanceColumns_();
    if (
      columns.some(function (x) {
        return !x.key || x.periods.join("|") !== "1교시|2교시|3교시";
      })
    ) {
      throw userError_(
        "출결 날짜 헤더 구조가 올바르지 않습니다.",
        "INVALID_HEADERS",
      );
    }
    const operatingSet = new Set(operatingDates);
    const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
    columns
      .filter(function (x) {
        return x.key >= today && x.key <= max && !operatingSet.has(x.key);
      })
      .sort(function (a, b) {
        return b.col - a.col;
      })
      .forEach(function (x) {
        sheet.deleteColumns(x.col, 3);
      });
    operatingDates.forEach(function (key) {
      ensureDateColumns_(key);
    });

    const attendance = operatingSet.has(today)
      ? processOperatingDateUnlocked_(today)
      : { date: today, processed: 0 };
    return {
      createdThrough: max,
      operatingDates: operatingDates.length,
      processed: attendance.processed,
    };
  });
}
function catchUpToday_() {
  if (isOperatingDate_(todayKey_())) return processOperatingDate_(todayKey_());
  return { date: todayKey_(), processed: 0 };
}
function dailyAttendanceJob() {
  initializeSheets_();
  requireConfig_();
  return syncFutureAttendanceColumns_();
}

function refreshStudentAttendance_(student) {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  const width = sheet.getLastColumn() - APP.ATTENDANCE_FIRST_COL + 1;
  student.attendance =
    width > 0
      ? sheet
          .getRange(student.row, APP.ATTENDANCE_FIRST_COL, 1, width)
          .getValues()[0]
      : [];
  return student;
}
function studentView_(key, config, student) {
  config = config || requireConfig_();
  student = student || requireStudent_(key);
  const today = todayKey_();
  const max = addDays_(today, 30);
  const cols = getAttendanceColumns_();
  const absences = [];
  cols
    .filter(function (x) {
      return x.key >= today && x.key <= max;
    })
    .forEach(function (x) {
      const periods = [];
      for (let p = 1; p <= 3; p++)
        if (
          normalizeStatus_(
            student.attendance[x.col - APP.ATTENDANCE_FIRST_COL + p - 1],
          ) === "4"
        )
          periods.push(p);
      if (periods.length)
        absences.push({
          date: x.key,
          periods: periods,
          cancellable: x.key > today || !isTodayStudentClosed_(config),
        });
    });
  return {
    studentKey: student.key,
    studentId: student.studentId,
    name: student.name,
    today: today,
    maxDate: max,
    todayClosed: isTodayStudentClosed_(config),
    absences: absences,
  };
}
function getStudentView(studentKey) {
  return publicCall_(function () {
    requireConfig_();
    return studentView_(studentKey);
  });
}
function registerStudentAbsence(studentKey, startKey, endKey) {
  return publicCall_(function () {
    const config = requireConfig_();
    const start = String(startKey);
    const end = String(endKey || startKey);
    assertFutureRange_(start, false);
    assertFutureRange_(end, false);
    if (end < start)
      throw userError_(
        "종료일은 시작일보다 빠를 수 없습니다.",
        "INVALID_DATE_RANGE",
      );
    const dates = [];
    const closedDates = getClosedDates_();
    for (let key = start; key <= end; key = addDays_(key, 1))
      if (isOperatingDate_(key, closedDates)) dates.push(key);
    if (!dates.length)
      throw userError_(
        "선택한 범위에 적용 가능한 운영일이 없습니다.",
        "NO_OPERATING_DATES",
      );
    return withWriteLock_(function () {
      const student = requireStudent_(studentKey);
      const appliedDates = [];
      const audits = [];
      dates.forEach(function (key) {
        if (key === todayKey_() && isTodayStudentClosed_(config))
          throw userError_("당일 등록이 마감되었습니다.", "TODAY_CLOSED");
        let changed = false;
        const col = ensureDateColumns_(key);
        const range = attendanceCell_(student, col, 1).offset(0, 0, 1, 3);
        const values = range.getValues()[0];
        for (let p = 1; p <= 3; p++)
          if (isApplied_(student, key, p)) {
            const before = normalizeStatus_(values[p - 1]);
            if (before !== "2" && before !== "3" && before !== "4") {
              values[p - 1] = 4;
              audits.push([
                now_(),
                "학생",
                student.key,
                student.studentId,
                parseDateKey_(key),
                p,
                before,
                "4",
              ]);
              changed = true;
            }
          }
        if (changed) range.setValues([values]);
        if (changed) appliedDates.push(key);
      });
      appendAudits_(audits);
      const view = studentView_(
        studentKey,
        config,
        refreshStudentAttendance_(student),
      );
      view.appliedDates = appliedDates;
      view.appliedCount = appliedDates.length;
      return view;
    });
  });
}
function cancelStudentAbsence(studentKey, key) {
  return publicCall_(function () {
    const config = requireConfig_();
    key = String(key);
    assertFutureRange_(key, false);
    if (key === todayKey_() && isTodayStudentClosed_(config))
      throw userError_("당일 등록이 마감되었습니다.", "TODAY_CLOSED");
    return withWriteLock_(function () {
      const student = requireStudent_(studentKey);
      const info = getAttendanceColumns_().find(function (x) {
        return x.key === key;
      });
      if (!info) throw userError_("취소할 사전 결석이 없습니다.", "NOT_FOUND");
      let count = 0;
      const restore = wasAutoProcessed_(key) ? 1 : "";
      const audits = [];
      const range = attendanceCell_(student, info.col, 1).offset(0, 0, 1, 3);
      const values = range.getValues()[0];
      for (let p = 1; p <= 3; p++)
        if (normalizeStatus_(values[p - 1]) === "4") {
          values[p - 1] = restore;
          audits.push([
            now_(),
            "학생",
            student.key,
            student.studentId,
            parseDateKey_(key),
            p,
            "4",
            normalizeStatus_(restore),
          ]);
          count++;
        }
      if (!count)
        throw userError_("취소할 수 있는 사전 결석이 없습니다.", "NOT_FOUND");
      range.setValues([values]);
      appendAudits_(audits);
      return studentView_(
        studentKey,
        config,
        refreshStudentAttendance_(student),
      );
    });
  });
}
