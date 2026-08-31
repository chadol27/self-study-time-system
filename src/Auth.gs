function hashToken_(token) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8)).replace(/=+$/, '');
}
function sessionPropertyKey_(hash) { return 'TEACHER_SESSION_' + hash; }
function cleanupSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const today = todayKey_();
  Object.keys(all).filter(function (k) { return k.indexOf('TEACHER_SESSION_') === 0 && all[k] !== today; }).forEach(function (k) { props.deleteProperty(k); });
}
function requireTeacher_(token) {
  if (!token || typeof token !== 'string' || token.length < 32) throw userError_('교사 인증이 필요합니다.', 'AUTH_REQUIRED');
  cleanupSessions_();
  const expiry = PropertiesService.getScriptProperties().getProperty(sessionPropertyKey_(hashToken_(token)));
  if (expiry !== todayKey_()) throw userError_('교사 인증이 만료되었습니다. 다시 로그인해 주세요.', 'AUTH_EXPIRED');
}
function teacherLogin(password) {
  return publicCall_(function () {
    const config = requireConfig_(); cleanupSessions_();
    if (typeof password !== 'string' || password !== config.teacherPassword) throw userError_('비밀번호가 올바르지 않습니다.', 'LOGIN_FAILED');
    const token = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(sessionPropertyKey_(hashToken_(token)), todayKey_());
    return { token: token, expires: todayKey_() };
  });
}
function teacherLogout(token) {
  return publicCall_(function () { if (token) PropertiesService.getScriptProperties().deleteProperty(sessionPropertyKey_(hashToken_(String(token)))); return true; });
}
function studentLogin(studentId, pin) {
  return publicCall_(function () {
    requireConfig_(); initializeSheets_();
    const id = String(studentId == null ? '' : studentId).trim();
    const secret = String(pin == null ? '' : pin);
    if (!id || !/^\d{4}$/.test(secret)) throw userError_('학번과 전화번호 뒤 4자리를 확인해 주세요.', 'LOGIN_FAILED');
    const active = readRoster_().filter(function (s) { return s.active && s.studentId === id; });
    if (active.length !== 1 || !/^\d{4}$/.test(active[0].pin) || active[0].pin !== secret) throw userError_('로그인 정보가 올바르지 않거나 명부를 확인해야 합니다.', 'LOGIN_FAILED');
    return studentView_(active[0].key);
  });
}
function requireStudent_(key) {
  if (!key || typeof key !== 'string') throw userError_('다시 로그인해 주세요.', 'AUTH_REQUIRED');
  const matches = readRoster_().filter(function (s) { return s.active && s.key === key; });
  if (matches.length !== 1) throw userError_('학생 정보를 확인할 수 없습니다. 다시 로그인해 주세요.', 'AUTH_REQUIRED');
  return matches[0];
}
