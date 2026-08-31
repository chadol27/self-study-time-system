function installDailyTrigger() {
  requireConfig_();
  const matches = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "dailyAttendanceJob";
  });
  matches.slice(1).forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  if (!matches.length)
    ScriptApp.newTrigger("dailyAttendanceJob")
      .timeBased()
      .atHour(0)
      .nearMinute(5)
      .everyDays(1)
      .inTimezone(APP.TZ)
      .create();
  return "일일 자동 출석 트리거가 설치되어 있습니다.";
}
