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
      const raw = info
        ? normalizeStatus_(attendanceCell_(s, info.col, period).getValue())
        : "";
      const status =
        raw === "2"
          ? "attended"
          : raw === "3"
            ? "absent"
            : raw === "4"
              ? "pre-absence"
              : "applied";
      const label =
        raw === "2"
          ? "출석"
          : raw === "3"
            ? "결석"
            : raw === "4"
              ? "결석 예정"
              : "신청함";
      seats[s.seat - 1] = {
        seat: s.seat,
        student: { key: s.key, studentId: s.studentId, name: s.name },
        status: status,
        label: label,
        applicationSummary: applicationSummary,
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
function teacherBatchChange(token, studentKeys, key, period, action) {
  return publicCall_(function () {
    const config = requireConfig_();
    requireTeacher_(token);
    key = String(key);
    period = Number(period);
    action = String(action);
    parseDateKey_(key);
    assertFutureRange_(key, true);
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
      const selectedSet = new Set(keys);
      const students = readRoster_().filter(function (s) {
        return s.active && selectedSet.has(s.key);
      });
      if (
        students.length !== keys.length ||
        students.some(function (student) {
          return !isApplied_(student, key, period);
        })
      )
        throw userError_(
          "선택한 학생 정보를 다시 확인해 주세요.",
          "INVALID_SELECTION",
        );
      const col = ensureDateColumns_(key);
      let changed = 0;
      const audits = [];
      students.forEach(function (student) {
        const range = attendanceCell_(student, col, 1).offset(0, 0, 1, 3);
        const values = range.getValues()[0];
        let studentChanged = false;
        for (let p = period; p <= 3; p++)
          if (isApplied_(student, key, p)) {
            const current = normalizeStatus_(values[p - 1]);
            const next =
              action === "present" ? "2" : action === "absent" ? "3" : "1";
            if (current !== normalizeStatus_(next)) {
              values[p - 1] = Number(next);
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
              studentChanged = true;
            }
          }
        if (studentChanged) range.setValues([values]);
      });
      appendAudits_(audits);
      return { changed: changed, view: teacherSeats_(key, period) };
    });
  });
}
