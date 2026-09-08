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
  const closedDates = getClosedDates_();
  for (let i = 0; i < 370; i++, key = addDays_(key, -1))
    if (isOperatingDate_(key, closedDates)) return key;
  return todayKey_();
}
function teacherSchedule(token) {
  return publicCall_(function () {
    requireConfig_();
    requireTeacher_(token);
    const today = todayKey_();
    const closedDates = getClosedDates_();
    const historical = getAttendanceColumns_()
      .map(function (x) {
        return x.key;
      })
      .concat(
        readExtraSheet_().blocks.map(function (x) {
          return x.key;
        }),
      )
      .filter(function (key) {
        return key && key < today && isOperatingDate_(key, closedDates);
      });
    const future = [];
    for (
      let key = today, max = addDays_(key, 30);
      key <= max;
      key = addDays_(key, 1)
    )
      if (isOperatingDate_(key, closedDates)) future.push(key);
    return Array.from(new Set(historical.concat(future))).sort();
  });
}
function teacherGetTodayName(token) {
  return publicCall_(function () {
    requireTeacher_(token);
    const sheet = spreadsheet_().getSheetByName(APP.SHEETS.TEACHERS);
    if (!simpleHeadersValid_(sheet, APP.TEACHER_HEADERS))
      throw userError_("교사 시트 헤더를 확인해 주세요.", "INVALID_HEADERS");
    const today = todayKey_();
    if (sheet.getLastRow() < 2) return { name: "" };
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const row = rows.find(function (values) {
      return (
        values[0] instanceof Date &&
        !isNaN(values[0].getTime()) &&
        dateKey_(values[0]) === today
      );
    });
    return { name: row ? String(row[1] == null ? "" : row[1]) : "" };
  });
}
function teacherSaveTodayName(token, name) {
  return publicCall_(function () {
    if (typeof name !== "string")
      throw userError_("교사 정보를 확인해 주세요.", "INVALID_TEACHER_NAME");
    const trimmed = name.trim();
    if (trimmed.length > 100)
      throw userError_(
        "교사 정보는 100자 이하로 입력해 주세요.",
        "INVALID_TEACHER_NAME",
      );
    return withWriteLock_(function () {
      requireTeacher_(token);
      const sheet = spreadsheet_().getSheetByName(APP.SHEETS.TEACHERS);
      if (!simpleHeadersValid_(sheet, APP.TEACHER_HEADERS))
        throw userError_("교사 시트 헤더를 확인해 주세요.", "INVALID_HEADERS");
      const today = todayKey_();
      const date = parseDateKey_(today);
      const count = Math.max(0, sheet.getLastRow() - 1);
      const rows = count ? sheet.getRange(2, 1, count, 2).getValues() : [];
      const index = rows.findIndex(function (values) {
        return (
          values[0] instanceof Date &&
          !isNaN(values[0].getTime()) &&
          dateKey_(values[0]) === today
        );
      });
      if (index >= 0) {
        sheet.getRange(index + 2, 2).setValue(trimmed);
      } else {
        const row = sheet.getLastRow() + 1;
        if (sheet.getMaxRows() < row)
          sheet.insertRowsAfter(sheet.getMaxRows(), 1);
        sheet.getRange(row, 1, 1, 2).setValues([[date, trimmed]]);
        sheet.getRange(row, 1).setNumberFormat("yyyy-MM-dd");
      }
      return { name: trimmed };
    });
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
function teacherSeatSnapshot_(key, period, writable = false) {
  const config = requireConfig_();
  if (!Number.isInteger(period) || period < 1 || period > 3)
    throw userError_("교시가 올바르지 않습니다.", "INVALID_PERIOD");
  parseDateKey_(key);
  const today = todayKey_();
  if (writable && key !== today)
    throw userError_("오늘 날짜만 수정할 수 있습니다.", "DATE_READ_ONLY");
  if (writable && !headersValid_())
    throw userError_(
      "명부 시트의 고정 헤더를 확인해 주세요.",
      "INVALID_HEADERS",
    );
  const columns = getAttendanceColumns_();
  let info = columns.find(function (x) {
    return x.key === key;
  });
  if (!isOperatingDate_(key))
    throw userError_("미운영일은 조회할 수 없습니다.", "CLOSED_DATE");
  if (key > addDays_(today, 30))
    throw userError_(
      "조회 가능한 미래 날짜를 벗어났습니다.",
      "DATE_OUT_OF_RANGE",
    );
  if (
    writable &&
    columns.some(function (x, i) {
      return (
        !x.key ||
        x.periods.join("|") !== "1교시|2교시|3교시" ||
        (i > 0 && columns[i - 1].key >= x.key)
      );
    })
  )
    throw userError_("출결 날짜 헤더 구조를 확인해 주세요.", "INVALID_HEADERS");
  if (!info && key >= today) {
    ensureDateColumns_(key, columns);
    info = columns.find(function (x) {
      return x.key === key;
    });
  }
  const snapshot = {
    config: config,
    today: today,
    columns: columns,
    info: info,
    roster: readRoster_(),
    report: null,
  };
  snapshot.report = validateAll_(false, snapshot);
  return snapshot;
}
function teacherSeats_(key, period, snapshot = null) {
  snapshot = snapshot || teacherSeatSnapshot_(key, period);
  const config = snapshot.config;
  const info = snapshot.info;
  const report = snapshot.report;
  const readOnly = key !== snapshot.today;
  const excluded = new Set(report.excludedKeys);
  const seats = /** @type {any[]} */ (
    Array.from({ length: config.totalSeats }, function (_, i) {
      return {
        seat: i + 1,
        student: null,
        status: "empty",
        label: "미신청",
        raw: "",
        checked: false,
      };
    })
  );
  snapshot.roster
    .filter(function (s) {
      return s.active && !excluded.has(s.key) && isAppliedOnDate_(s, key);
    })
    .forEach(function (s) {
      const applicationSummary = ["월", "화", "수", "목"]
        .map(function (day, dayIndex) {
          const periods = [1, 2, 3].filter(function (p) {
            return Number(s.applications[dayIndex * 3 + p - 1]) === 1;
          });
          if (!periods.length) return "";
          if (periods.length === 3) return day;
          if (periods.length === 2 && periods[1] === periods[0] + 1)
            return day + periods[0] + "-" + periods[1];
          return day + periods.join(",");
        })
        .filter(Boolean)
        .join(", ");
      const index = info
        ? info.col - APP.ATTENDANCE_FIRST_COL + period - 1
        : -1;
      const raw = index >= 0 ? normalizeStatus_(s.attendance[index]) : "";
      const applied = isApplied_(s, key, period);
      const previous =
        info && period > 1 ? normalizeStatus_(s.attendance[index - 1]) : "";
      const checked =
        raw === "2" ||
        (!readOnly &&
          raw !== "3" &&
          raw !== "4" &&
          applied &&
          (period === 1 || (previous !== "3" && previous !== "4")));
      const status =
        raw === "2"
          ? "attended"
          : raw === "3"
            ? "absent"
            : raw === "4"
              ? "pre-absence"
              : applied
                ? "applied"
                : "not-applied";
      const label =
        raw === "2"
          ? "출석"
          : raw === "3"
            ? "결석"
            : raw === "4"
              ? "결석 예정"
              : applied
                ? "신청"
                : "해당교시 미신청";
      seats[s.seat - 1] = {
        seat: s.seat,
        student: { key: s.key, studentId: s.studentId, name: s.name },
        status: status,
        label: label,
        applied: applied,
        raw: raw,
        checked: checked,
        applicationSummary: applicationSummary,
      };
    });
  return {
    date: key,
    period: period,
    readOnly: readOnly,
    seats: seats,
    errors: report.errors,
    totalSeats: config.totalSeats,
  };
}
function teacherSaveSeats(token, selections, key, period) {
  return publicCall_(function () {
    return withWriteLock_(function () {
      requireTeacher_(token);
      key = String(key);
      period = Number(period);
      const snapshot = teacherSeatSnapshot_(key, period, true);
      if (
        !Array.isArray(selections) ||
        selections.length > snapshot.config.totalSeats
      )
        throw userError_(
          "저장할 학생 정보를 확인해 주세요.",
          "INVALID_SELECTION",
        );
      const selected = new Map();
      selections.forEach(function (selection) {
        if (
          !selection ||
          typeof selection.studentKey !== "string" ||
          !selection.studentKey ||
          typeof selection.checked !== "boolean" ||
          selected.has(selection.studentKey)
        )
          throw userError_(
            "저장할 학생 정보를 확인해 주세요.",
            "INVALID_SELECTION",
          );
        selected.set(selection.studentKey, selection.checked);
      });
      const excluded = new Set(snapshot.report.excludedKeys);
      const students = snapshot.roster.filter(function (student) {
        return (
          student.active &&
          !excluded.has(student.key) &&
          isAppliedOnDate_(student, key) &&
          selected.has(student.key)
        );
      });
      if (students.length !== selected.size)
        throw userError_(
          "선택한 학생 정보를 다시 확인해 주세요.",
          "INVALID_SELECTION",
        );
      const col = snapshot.info.col + period - 1;
      const index = col - APP.ATTENDANCE_FIRST_COL;
      const ranges = new Map();
      const audits = [];
      const timestamp = now_();
      const date = parseDateKey_(key);
      students.forEach(function (student) {
        const current = normalizeStatus_(student.attendance[index]);
        // A stale checkbox must never overwrite a newly registered pre-absence.
        if (current === "4") return;
        const next = selected.get(student.key)
          ? "2"
          : isApplied_(student, key, period)
            ? "3"
            : "";
        if (current === next) return;
        if (!ranges.has(next)) ranges.set(next, []);
        ranges.get(next).push(student.row);
        student.attendance[index] = next === "" ? "" : Number(next);
        audits.push([
          timestamp,
          "교사",
          student.key,
          student.studentId,
          date,
          period,
          current,
          next,
        ]);
      });
      teacherWriteSeatRanges_(col, ranges);
      appendAudits_(audits);
      return {
        changed: audits.length,
        view: teacherSeats_(key, period, snapshot),
      };
    });
  });
}
function teacherWriteSeatRanges_(col, ranges) {
  if (!ranges.size) return;
  let columnName = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26))
    columnName = String.fromCharCode(65 + ((n - 1) % 26)) + columnName;
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.ROSTER);
  // Address only changed cells, preserving formulas and unknown cells in gaps.
  ranges.forEach(function (rows, next) {
    const addresses = rows.map(function (row) {
      return columnName + row;
    });
    sheet.getRangeList(addresses).setValue(next === "" ? "" : Number(next));
  });
}
function teacherBatchChange(token, studentKeys, key, period, action) {
  return publicCall_(function () {
    const config = requireConfig_();
    requireTeacher_(token);
    key = String(key);
    period = Number(period);
    action = String(action);
    parseDateKey_(key);
    if (key !== todayKey_())
      throw userError_("오늘 날짜만 수정할 수 있습니다.", "DATE_READ_ONLY");
    assertFutureRange_(key, false);
    if (!isOperatingDate_(key))
      throw userError_("미운영일은 변경할 수 없습니다.", "CLOSED_DATE");
    if (!Number.isInteger(period) || period < 1 || period > 3)
      throw userError_("교시가 올바르지 않습니다.", "INVALID_PERIOD");
    if (!["present", "absent", "restore"].includes(action))
      throw userError_("변경 작업이 올바르지 않습니다.", "INVALID_ACTION");
    if (!Array.isArray(studentKeys))
      throw userError_("학생을 선택해 주세요.", "INVALID_SELECTION");
    const keys = Array.from(
      new Set(
        studentKeys.map(function (studentKey) {
          return String(studentKey);
        }),
      ),
    );
    if (!keys.length || keys.length > config.totalSeats)
      throw userError_("선택한 학생을 확인해 주세요.", "INVALID_SELECTION");
    return withWriteLock_(function () {
      requireTeacher_(token);
      const snapshot = teacherSeatSnapshot_(key, period, true);
      if (keys.length > snapshot.config.totalSeats)
        throw userError_("선택한 학생을 확인해 주세요.", "INVALID_SELECTION");
      const selectedSet = new Set(keys);
      const excluded = new Set(snapshot.report.excludedKeys);
      const students = snapshot.roster.filter(function (s) {
        return s.active && !excluded.has(s.key) && selectedSet.has(s.key);
      });
      if (
        students.length !== keys.length ||
        students.some(function (student) {
          return !isAppliedOnDate_(student, key);
        })
      )
        throw userError_(
          "선택한 학생 정보를 다시 확인해 주세요.",
          "INVALID_SELECTION",
        );
      const col = snapshot.info.col + period - 1;
      const index = col - APP.ATTENDANCE_FIRST_COL;
      const ranges = new Map();
      let changed = 0;
      const audits = [];
      students.forEach(function (student) {
        const current = normalizeStatus_(student.attendance[index]);
        const applied = isApplied_(student, key, period);
        if (
          action === "absent" &&
          current !== "2" &&
          !(applied && current === "1")
        )
          return;
        if (action === "restore" && current === "4") return;
        const next =
          action === "present"
            ? "2"
            : action === "absent"
              ? "3"
              : applied
                ? "1"
                : "";
        if (current === next) return;
        if (!ranges.has(next)) ranges.set(next, []);
        ranges.get(next).push(student.row);
        student.attendance[index] = next === "" ? "" : Number(next);
        audits.push([
          now_(),
          "교사",
          student.key,
          student.studentId,
          parseDateKey_(key),
          period,
          current,
          next,
        ]);
        changed++;
      });
      teacherWriteSeatRanges_(col, ranges);
      appendAudits_(audits);
      return { changed: changed, view: teacherSeats_(key, period, snapshot) };
    });
  });
}
