const STUDENT_DIRECTORY_CACHE_KEY = "STUDENT_DIRECTORY_V1";

function readStudentDirectorySheet_() {
  const sheet = spreadsheet_().getSheetByName(APP.SHEETS.STUDENT_DIRECTORY);
  if (!simpleHeadersValid_(sheet, APP.STUDENT_DIRECTORY_HEADERS)) {
    return {
      students: {},
      errors: ["학생명단 시트 헤더가 올바르지 않습니다."],
    };
  }
  const students = {};
  const errors = [];
  const duplicates = new Set();
  const seenIds = new Set();
  if (sheet.getLastRow() < 2) return { students: students, errors: errors };
  sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 3)
    .getDisplayValues()
    .forEach(function (row, index) {
      const studentId = row[0].trim();
      const name = row[1].trim();
      const pin = row[2].trim();
      const rowNumber = index + 2;
      if (!studentId && !name && !pin) return;
      if (!studentId) {
        errors.push("학생명단 학번 누락: " + rowNumber + "행");
        return;
      }
      if (seenIds.has(studentId)) duplicates.add(studentId);
      seenIds.add(studentId);
      if (!name || !/^\d{4}$/.test(pin)) {
        if (!name) errors.push("학생명단 이름 누락: " + studentId);
        if (!/^\d{4}$/.test(pin))
          errors.push("학생명단 전화번호 뒤 4자리 형식 오류: " + studentId);
        return;
      }
      students[studentId] = { name: name, pin: pin };
    });
  duplicates.forEach(function (studentId) {
    delete students[studentId];
    errors.push("학생명단 학번 중복: " + studentId);
  });
  return { students: students, errors: errors };
}

function refreshStudentDirectoryCache_() {
  const directory = readStudentDirectorySheet_();
  const serialized = JSON.stringify(directory);
  try {
    CacheService.getScriptCache().put(
      STUDENT_DIRECTORY_CACHE_KEY,
      serialized,
      21600,
    );
  } catch (error) {
    throw userError_(
      "학생명단이 너무 커서 캐시에 저장할 수 없습니다.",
      "DIRECTORY_CACHE_FAILED",
    );
  }
  return directory;
}

function getStudentDirectory_() {
  const cached = CacheService.getScriptCache().get(STUDENT_DIRECTORY_CACHE_KEY);
  if (!cached) return refreshStudentDirectoryCache_();
  try {
    return JSON.parse(cached);
  } catch (error) {
    return refreshStudentDirectoryCache_();
  }
}

function getStudentDirectoryForIds_(studentIds) {
  let directory = getStudentDirectory_();
  const missing = Array.from(new Set(studentIds)).some(function (studentId) {
    return studentId && !directory.students[studentId];
  });
  if (missing) directory = refreshStudentDirectoryCache_();
  return directory;
}

function applyStudentDirectory() {
  initializeSheets_();
  const directory = refreshStudentDirectoryCache_();
  const count = Object.keys(directory.students).length;
  SpreadsheetApp.getUi().alert(
    directory.errors.length
      ? "유효한 학생 " +
          count +
          "명을 적용했습니다.\n\n" +
          directory.errors.join("\n")
      : "학생 명단 " + count + "명을 적용했습니다.",
  );
}
