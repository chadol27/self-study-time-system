function doGet(e) {
  initializeSheets_();
  const config = getConfig_();
  if (config.ok) catchUpToday_();
  const requested = e && e.parameter ? String(e.parameter.page || '') : '';
  const page = requested === 'student' ? 'Student' : requested === 'teacher' ? 'Teacher' : 'Index';
  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = ScriptApp.getService().getUrl();
  template.bootstrap = JSON.stringify({
    configOk: config.ok,
    configErrors: config.errors,
    page: page.toLowerCase()
  }).replace(/</g, '\\u003c');
  return template.evaluate()
    .setTitle('스공시 좌석 및 출결 관리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPublicBootstrap() {
  return publicCall_(function () {
    initializeSheets_();
    const config = getConfig_();
    return { configOk: config.ok, configErrors: config.errors };
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('스공시 관리')
    .addItem('시트 초기화 및 검사', 'initializeAndValidate')
    .addItem('일일 트리거 설치', 'installDailyTrigger')
    .addToUi();
}

function initializeAndValidate() {
  initializeSheets_();
  const report = validateAll_();
  SpreadsheetApp.getUi().alert(report.errors.length ? report.errors.join('\n') : '검사가 완료되었습니다. 오류가 없습니다.');
}
