/**
 * 基翔點名-雲端系統 - 後端 (Google Apps Script)
 * ============================================
 * 部署方式：
 * 1. 建立一個新的 Google 試算表，複製網址列裡的試算表 ID 或整段網址，
 *    貼到下面的 SPREADSHEET_ID（若這份程式是從該試算表「擴充功能 > Apps Script」開啟的，可留空）。
 * 2. 把這個檔案全部內容貼到 Apps Script 編輯器（先刪掉原本的程式碼）。
 * 3. 執行一次 setup()（第一次會跳出授權視窗，點允許）。
 *    它會自動建立/補齊所有分頁與欄位，建立管理員帳號 admin / 1234，並建立每日清理排程。
 *    之後更新程式碼時，只有在說明需要時才要再執行 setup()；分頁格式整理會在每天凌晨 3 點自動進行。
 * 4. 「部署 > 新增部署作業」→ 類型「網頁應用程式」→ 執行身分「我」、存取權限「任何人」。
 *    取得 /exec 網址後貼到 index.html 最上面的 CONFIG.API_URL。
 *    （修改程式碼後要到「部署 > 管理部署作業 > 編輯 > 版本：新版本」重新部署，網址不變）
 */

// ---------- 基本設定 ----------
var SPREADSHEET_ID = '1Xk5mxGTzRDmsLlbag6dkHn6420n6ZLUP59z6QefTjKo';      // 試算表 ID 或整段網址 (留空則使用目前綁定的試算表)
var RETENTION_DAYS = 40;      // 點名紀錄保留天數：超過 40 天的舊資料於每天凌晨 3 點自動清除
var SESSION_HOURS = 12;       // 登入 token 最長有效時數 (前端另外有閒置 60 分鐘自動登出)
var TZ = 'Asia/Taipei';
var SCHEMA_VERSION = 'schema_v8';
var DEFAULT_CONTACT_TYPES = ['爸爸', '媽媽', '住家'];
var VALID_STATUS = ['present', 'selfDrop', 'selfWalk', 'leave', 'fixedOff', ''];

var SHEET_ADMINS = 'Admins';
var SHEET_TEACHERS = 'Teachers';
var SHEET_SCHOOLS = 'Schools';
var SHEET_STUDENTS = 'Students';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS = 'Sessions';
var SHEET_SETTINGS = 'Settings';
var SHEET_SUBMISSIONS = 'Submissions';
var SHEET_SPECIAL = 'SpecialDays';
var SHEET_BOOKINGS = 'LeaveBookings';
var SHEET_EXTRAS = 'ExtraClasses';

var HEADERS = {
  Admins: ['id', 'username', 'password', 'name'],
  Teachers: ['id', 'username', 'password', 'name', 'active', 'createdAt'],
  Schools: ['id', 'name', 'createdAt', 'active'],
  Students: ['id', 'schoolId', 'name', 'grade', 'contacts', 'fixedOffDays', 'active', 'createdAt', 'note'],
  // auto=true：這筆狀態是系統預設值（預約請假 / 固定不進班），老師沒有改過
  Attendance: ['id', 'date', 'schoolId', 'studentId', 'status', 'selfDrop', 'selfWalk', 'teacherId', 'teacherName', 'submittedAt', 'wasFixedOff', 'auto'],
  Sessions: ['token', 'userId', 'role', 'name', 'username', 'createdAt', 'expiresAt'],
  Settings: ['key', 'value'],
  // 每次送出點名的紀錄：同一天同一校可能上午 A 老師、下午 B 老師分批送出
  Submissions: ['id', 'date', 'schoolId', 'teacherId', 'teacherName', 'submittedAt', 'count'],
  // 點名日設定：workday=補班補課（週末可點名）、holiday=國定假日放假（不能點名）
  SpecialDays: ['date', 'type', 'reason', 'createdBy', 'createdAt'],
  // 預約請假：家長事先告知的請假期間（管理員與老師都可以登記）
  LeaveBookings: ['id', 'schoolId', 'studentId', 'startDate', 'endDate', 'reason', 'createdBy', 'createdRole', 'createdAt', 'updatedBy', 'updatedAt'],
  // 臨時進班：固定不進班的學生因臨時或考前加強要進班，期間內固定不進班的日子改為正常進班
  ExtraClasses: ['id', 'schoolId', 'studentId', 'startDate', 'endDate', 'reason', 'createdBy', 'createdRole', 'createdAt', 'updatedBy', 'updatedAt']
};

// ---------- 初始化 ----------
function setup() {
  ensureSchema_(true);
  var def = ss_().getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss_().deleteSheet(def);
  createDailyCleanupTrigger();
  return 'setup done';
}

function createDailyCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupOldRecords') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupOldRecords').timeBased().everyDays(1).atHour(3).create();
}

// 建立缺少的分頁、補齊缺少的欄位、建立預設管理員。
// force=true（執行 setup() 或每天凌晨 3 點排程）才會把整張分頁設成純文字格式；平常只做快速檢查，避免拖慢使用者
function ensureSchema_(force) {
  var cache = CacheService.getScriptCache();
  if (!force && cache.get(SCHEMA_VERSION)) return;
  var ss = ss_();
  var existing = {};
  ss.getSheets().forEach(function (sh) { existing[sh.getName()] = sh; });
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = existing[name];
    var isNew = !sheet;
    if (isNew) sheet = ss.insertSheet(name);
    var lastCol = sheet.getLastColumn();
    var have = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
    var added = false;
    HEADERS[name].forEach(function (h) {
      if (have.indexOf(h) === -1) { have.push(h); added = true; }
    });
    if (added) sheet.getRange(1, 1, 1, have.length).setNumberFormat('@').setValues([have]);
    if (isNew || force) {
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, sheet.getMaxRows(), have.length).setNumberFormat('@');
    }
  });
  if (sheet_(SHEET_ADMINS).getLastRow() < 2) {
    appendObject(SHEET_ADMINS, { id: genId('A'), username: 'admin', password: hashPwd('1234'), name: '管理員' });
  }
  cache.put(SCHEMA_VERSION, '1', 21600);
}

// ---------- 工具函式 ----------
var _ssCache = null;

function extractSpreadsheetId_(raw) {
  var s = String(raw || '').trim();
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return s.replace(/^\/+|\/+$/g, '').replace(/\/edit.*$/, '');
}

function ss_() {
  if (_ssCache) return _ssCache;
  var ss;
  if (SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(extractSpreadsheetId_(SPREADSHEET_ID));
    } catch (e) {
      throw new Error('無法開啟試算表，請確認 SPREADSHEET_ID 是否正確：' + e.message);
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) throw new Error('找不到試算表：請在程式碼最上方的 SPREADSHEET_ID 填入試算表 ID');
  _ssCache = ss;
  return ss;
}

function sheet_(name) {
  var sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('找不到工作表「' + name + '」：請先執行一次 setup()');
  return sheet;
}

function genId(prefix) {
  return (prefix || 'X') + Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

function hashPwd(pwd) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pwd), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function isTrue_(v) { return String(v).toLowerCase() === 'true'; }
function isActive_(v) { return String(v).toLowerCase() !== 'false'; }

function toCell_(v) {
  if (v === true) return 'true';
  if (v === false) return 'false';
  if (v === undefined || v === null) return '';
  return v;
}

// 全域寫入鎖：多位管理員/老師同時操作時，所有寫入依序執行，避免互相覆蓋。
// 可重入（同一次執行內巢狀呼叫不會卡死）
var _lockDepth = 0;
function withLock_(fn) {
  if (_lockDepth > 0) return fn();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { throw new Error('BUSY'); }
  _lockDepth++;
  try { return fn(); } finally { _lockDepth--; lock.releaseLock(); }
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function dateStrDaysAgo_(n) {
  return Utilities.formatDate(new Date(Date.now() - n * 86400000), TZ, 'yyyy-MM-dd');
}

// 1(週一) ~ 7(週日)，與指令碼時區無關
function weekdayOf(dateStr) {
  return parseInt(Utilities.formatDate(new Date(dateStr + 'T12:00:00+08:00'), TZ, 'u'), 10);
}

// 試算表若不小心把文字轉成日期，讀取時轉回字串
function fromDateCell_(v, header) {
  if (header === 'date' || header === 'startDate' || header === 'endDate') return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return v.toISOString();
}

// 同一次請求內讀過的分頁先記住，不重複讀整張表（有寫入時自動作廢）
var _memo = {};
function touch_(name) { delete _memo[name]; }
function sheetToObjects(name) {
  if (_memo[name]) return _memo[name].map(function (o) { return Object.assign({}, o); });
  var out = readSheetObjects_(name);
  _memo[name] = out;
  return out.map(function (o) { return Object.assign({}, o); });
}
function readSheetObjects_(name) {
  var rows = sheet_(name).getDataRange().getValues();
  if (rows.length < 1) return [];
  var headers = rows[0];
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var v = rows[i][j];
      if (v instanceof Date) v = fromDateCell_(v, headers[j]);
      obj[headers[j]] = v;
    }
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function headersOf_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// 整列重寫前，把被試算表自動轉成日期型別的儲存格轉回字串
function plainRow_(headers, row) {
  return row.map(function (v, j) { return v instanceof Date ? fromDateCell_(v, headers[j]) : v; });
}

function appendObject(name, obj) { appendObjects(name, [obj]); }

function appendObjects(name, objs) {
  if (!objs.length) return;
  touch_(name);
  var sheet = sheet_(name);
  var headers = headersOf_(sheet);
  var data = objs.map(function (o) {
    return headers.map(function (h) { return toCell_(o[h]); });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, headers.length).setNumberFormat('@').setValues(data);
}

function updateObjectByRow(name, rowIndex, obj) {
  touch_(name);
  var sheet = sheet_(name);
  var headers = headersOf_(sheet);
  var row = headers.map(function (h) { return toCell_(obj[h]); });
  sheet.getRange(rowIndex, 1, 1, headers.length).setNumberFormat('@').setValues([row]);
}

// ---------- Settings (key/value) ----------
function getSetting_(key, def) {
  var rows = sheetToObjects(SHEET_SETTINGS);
  for (var i = 0; i < rows.length; i++) if (rows[i].key === key) return rows[i].value;
  return def;
}

function setSetting_(key, value) {
  var rows = sheetToObjects(SHEET_SETTINGS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].key === key) { updateObjectByRow(SHEET_SETTINGS, rows[i]._row, { key: key, value: value }); return; }
  }
  appendObject(SHEET_SETTINGS, { key: key, value: value });
}

// ---------- Session ----------
function createSession(user, role) {
  var sheet = sheet_(SHEET_SESSIONS);
  if (sheet.getLastRow() > 300) cleanupExpiredSessions_();
  var token = Utilities.getUuid();
  var now = new Date();
  var sess = {
    token: token, userId: user.id, role: role, name: user.name, username: user.username,
    createdAt: now.toISOString(), expiresAt: new Date(now.getTime() + SESSION_HOURS * 3600 * 1000).toISOString()
  };
  appendObject(SHEET_SESSIONS, sess);
  cacheSession_(sess);
  return token;
}

// 登入身分放進快取：每個動作不必再讀整張「登入紀錄」表，速度快很多
function cacheSession_(s) {
  CacheService.getScriptCache().put('sess_' + s.token, JSON.stringify({
    token: s.token, userId: s.userId, role: s.role, name: s.name, username: s.username, expiresAt: s.expiresAt
  }), 21600);
}

function getSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var hit = cache.get('sess_' + token);
  if (hit) {
    var c = JSON.parse(hit);
    return new Date(c.expiresAt).getTime() < Date.now() ? null : c;
  }
  var sessions = sheetToObjects(SHEET_SESSIONS);
  for (var i = sessions.length - 1; i >= 0; i--) {
    if (sessions[i].token === token) {
      if (new Date(sessions[i].expiresAt).getTime() < Date.now()) return null;
      cacheSession_(sessions[i]);
      return sessions[i];
    }
  }
  return null;
}

function requireSession(token, role) {
  var s = getSession(token);
  if (!s) throw new Error('AUTH_EXPIRED');
  if (role && s.role !== role) throw new Error('FORBIDDEN');
  // 帳號被刪除後，已登入的裝置也立即失效（刪除帳號時會一併清掉這個快取）
  var cache = CacheService.getScriptCache();
  if (cache.get('user_' + s.userId)) return s;
  var users = sheetToObjects(s.role === 'admin' ? SHEET_ADMINS : SHEET_TEACHERS);
  var u = users.filter(function (x) { return x.id === s.userId; })[0];
  if (!u || (s.role === 'teacher' && !isActive_(u.active))) throw new Error('AUTH_EXPIRED');
  s.name = u.name;
  cache.put('user_' + s.userId, '1', 600);
  return s;
}

function cleanupExpiredSessions_() {
  withLock_(function () {
    var sheet = sheet_(SHEET_SESSIONS);
    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;
    var expCol = rows[0].indexOf('expiresAt');
    var keep = [rows[0]];
    for (var i = 1; i < rows.length; i++) {
      if (new Date(rows[i][expCol]).getTime() >= Date.now()) keep.push(plainRow_(rows[0], rows[i]));
    }
    if (keep.length < rows.length) {
      touch_(SHEET_SESSIONS);
      sheet.getRange(2, 1, rows.length - 1, rows[0].length).clearContent();
      if (keep.length > 1) sheet.getRange(2, 1, keep.length - 1, rows[0].length).setNumberFormat('@').setValues(keep.slice(1));
    }
  });
}

// ---------- HTTP 入口 ----------
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '基翔點名-雲端系統 API 運作中' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    result = route(body);
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function route(body) {
  _memo = {};
  var action = body.action;
  if (action === 'ping') {
    ss_();
    ensureSchema_(false);
    return { ok: true, time: new Date().toISOString() };
  }
  try { ensureSchema_(false); } catch (e) { /* 交給後面的動作丟出明確錯誤 */ }
  var ACTIONS = {
    login: actionLogin, logout: actionLogout, listLoginAccounts: actionListLoginAccounts,
    bootstrap: actionBootstrap, changePwd: actionChangePwd,
    getContactTypes: actionGetContactTypes, addContactType: actionAddContactType,
    listSchools: actionListSchools, addSchool: actionAddSchool, addSchoolsBatch: actionAddSchoolsBatch,
    editSchool: actionEditSchool, deleteSchool: actionDeleteSchool, getDayLeaves: actionGetDayLeaves,
    listStudents: actionListStudents, addStudent: actionAddStudent, addStudentsBatch: actionAddStudentsBatch,
    editStudent: actionEditStudent, deleteStudent: actionDeleteStudent, deleteStudentsBatch: actionDeleteStudentsBatch,
    listTeachers: actionListTeachers, addTeacher: actionAddTeacher, editTeacher: actionEditTeacher,
    deleteTeacher: actionDeleteTeacher, deleteTeachersBatch: actionDeleteTeachersBatch,
    getRollCall: actionGetRollCall, submitAttendance: actionSubmitAttendance, saveMarks: actionSaveMarks, setPresence: actionSetPresence,
    getDayInfo: actionGetDayInfo, listSpecialDays: actionListSpecialDays,
    setSpecialDay: actionSetSpecialDay, deleteSpecialDay: actionDeleteSpecialDay,
    setSpecialDaysBatch: actionSetSpecialDaysBatch, editSpecialDay: actionEditSpecialDay,
    listBookings: actionListBookings, saveBooking: actionSaveBooking, saveBookingsBatch: actionSaveBookingsBatch, deleteBooking: actionDeleteBooking,
    getOverview: actionGetOverview, getSchoolDayDetail: actionGetSchoolDayDetail,
    listLeaveStreaks: actionListLeaveStreaks
  };
  // 會寫入試算表的動作：一律在全域鎖內執行（多人同時操作也不會互相覆蓋）
  var WRITES = ['login', 'logout', 'changePwd', 'addContactType', 'addSchool', 'addSchoolsBatch', 'editSchool', 'deleteSchool',
    'addStudent', 'addStudentsBatch', 'editStudent', 'deleteStudent', 'deleteStudentsBatch',
    'addTeacher', 'editTeacher', 'deleteTeacher', 'deleteTeachersBatch', 'submitAttendance', 'saveMarks', 'setSpecialDay', 'deleteSpecialDay',
    'setSpecialDaysBatch', 'editSpecialDay', 'saveBooking', 'saveBookingsBatch', 'deleteBooking'];
  var fn = ACTIONS[action];
  if (!fn) return { ok: false, error: 'UNKNOWN_ACTION' };
  if (WRITES.indexOf(action) !== -1) return withLock_(function () { return fn(body); });
  return fn(body);
}

// ---------- 登入 / 帳號 ----------
function actionLogin(body) {
  var username = String(body.username || '').trim();
  var hashed = hashPwd(String(body.password || ''));
  var admins = sheetToObjects(SHEET_ADMINS);
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].username === username && admins[i].password === hashed) {
      return { ok: true, token: createSession(admins[i], 'admin'), role: 'admin', name: admins[i].name, today: today_() };
    }
  }
  var teachers = sheetToObjects(SHEET_TEACHERS);
  for (var j = 0; j < teachers.length; j++) {
    var t = teachers[j];
    if (isActive_(t.active) && t.username === username && t.password === hashed) {
      return { ok: true, token: createSession(t, 'teacher'), role: 'teacher', name: t.name, today: today_() };
    }
  }
  return { ok: false, error: 'INVALID_LOGIN' };
}

function actionLogout(body) {
  if (!body.token) return { ok: true };
  var sess = getSession(body.token);
  if (sess) CacheService.getScriptCache().remove('pres_' + sess.userId);   // 登出後不再顯示「點名中」
  CacheService.getScriptCache().remove('sess_' + body.token);
  withLock_(function () {
    var sheet = sheet_(SHEET_SESSIONS);
    var rows = sheetToObjects(SHEET_SESSIONS);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (rows[i].token === body.token) { sheet.deleteRow(rows[i]._row); break; }
    }
  });
  return { ok: true };
}

function actionListLoginAccounts(body) {
  var admins = sheetToObjects(SHEET_ADMINS).map(function (a) {
    return { username: a.username, name: a.name, role: 'admin' };
  });
  var teachers = sheetToObjects(SHEET_TEACHERS).filter(function (t) { return isActive_(t.active); })
    .map(function (t) { return { username: t.username, name: t.name, role: 'teacher' }; });
  teachers.sort(function (a, b) { return String(a.username).localeCompare(String(b.username)); });
  return { ok: true, accounts: admins.concat(teachers) };
}

function actionBootstrap(body) {
  var s = requireSession(body.token);
  return { ok: true, role: s.role, name: s.name, today: today_() };
}

function actionChangePwd(body) {
  var s = requireSession(body.token);
  var newPassword = String(body.newPassword || '');
  if (!newPassword) return { ok: false, error: 'PASSWORD_REQUIRED' };
  var sheetName = s.role === 'admin' ? SHEET_ADMINS : SHEET_TEACHERS;
  var rows = sheetToObjects(sheetName);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === s.userId) {
      if (rows[i].password !== hashPwd(body.oldPassword || '')) return { ok: false, error: 'WRONG_OLD_PASSWORD' };
      rows[i].password = hashPwd(newPassword);
      updateObjectByRow(sheetName, rows[i]._row, rows[i]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'USER_NOT_FOUND' };
}

// ---------- 老師帳號 ----------
// 取最小的未使用編號：Teacher02 被刪除後，下一位新老師會重新使用 Teacher02
function nextTeacherUsername_(activeTeachers) {
  var used = {};
  activeTeachers.forEach(function (t) {
    var m = String(t.username).match(/^Teacher(\d+)$/i);
    if (m) used[parseInt(m[1], 10)] = true;
  });
  var n = 1;
  while (used[n]) n++;
  return 'Teacher' + (n < 10 ? '0' + n : String(n));
}

function activeTeachers_() {
  return sheetToObjects(SHEET_TEACHERS).filter(function (t) { return isActive_(t.active); });
}

// ---------- 老師目前在哪間學校（只存在快取，不寫試算表）----------
// 老師端畫面開著時每 30 秒回報一次，每次點名也會更新；超過 PRESENCE_TTL 秒沒有任何回報就視為離開
// 關掉頁面、登出會立即清除
var PRESENCE_TTL = 600;   // 最後一次操作後 10 分鐘內都算在這間學校（鎖螢幕等學生時不會消失）
function presenceList_(excludeUserId) {
  var cache = CacheService.getScriptCache();
  var ids = activeTeachers_().map(function (t) { return t.id; });
  if (!ids.length) return [];
  var got = cache.getAll(ids.map(function (id) { return 'pres_' + id; })) || {};
  var schoolMap = {};
  sheetToObjects(SHEET_SCHOOLS).forEach(function (sc) { schoolMap[sc.id] = sc.name; });
  var now = Date.now(), out = [];
  Object.keys(got).forEach(function (k) {
    try {
      var p = JSON.parse(got[k]);
      if (!p || p.teacherId === excludeUserId || now - p.at > PRESENCE_TTL * 1000) return;
      p.schoolName = schoolMap[p.schoolId] || '';
      p.ago = Math.max(0, Math.round((now - p.at) / 1000));
      out.push(p);
    } catch (e) { /* 壞掉的快取略過 */ }
  });
  out.sort(function (a, b) { return String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant') || b.at - a.at; });
  return out;
}

function putPresence_(s, schoolId, mode) {
  CacheService.getScriptCache().put('pres_' + s.userId, JSON.stringify({
    teacherId: s.userId, name: s.name, schoolId: String(schoolId), mode: mode === 'edit' ? 'edit' : 'view', at: Date.now()
  }), PRESENCE_TTL);
}
// 老師回報目前所在學校與狀態（view=查看、edit=點名中）；schoolId 空白＝離開。回傳其他老師的狀態
function actionSetPresence(body) {
  var s = requireSession(body.token);
  var cache = CacheService.getScriptCache();
  if (s.role === 'teacher') {
    if (body.schoolId) putPresence_(s, body.schoolId, body.mode);
    else cache.remove('pres_' + s.userId);
  }
  return { ok: true, presence: presenceList_(s.userId) };
}

function actionListTeachers(body) {
  requireSession(body.token, 'admin');
  var teachers = activeTeachers_();
  teachers.sort(function (a, b) { return String(a.username).localeCompare(String(b.username)); });
  return {
    ok: true,
    nextUsername: nextTeacherUsername_(teachers),
    teachers: teachers.map(function (t) { return { id: t.id, username: t.username, name: t.name, active: true }; })
  };
}

function actionAddTeacher(body) {
  requireSession(body.token, 'admin');
  return withLock_(function () {
    var all = sheetToObjects(SHEET_TEACHERS);
    var active = all.filter(function (t) { return isActive_(t.active); });
    var name = String(body.name || '').trim();
    if (!name) return { ok: false, error: 'NAME_REQUIRED' };
    var username = String(body.username || '').trim();
    if (!username) return { ok: false, error: 'USERNAME_REQUIRED' };
    var lower = username.toLowerCase();
    var taken = active.concat(sheetToObjects(SHEET_ADMINS)).some(function (u) { return String(u.username).toLowerCase() === lower; });
    if (taken) return { ok: false, error: 'USERNAME_TAKEN' };
    // 舊版「停用」留下的同名帳號列，先移除以便重新使用編號
    deleteRows_(SHEET_TEACHERS, all.filter(function (t) { return !isActive_(t.active) && String(t.username).toLowerCase() === lower; }));
    var id = genId('TC');
    appendObject(SHEET_TEACHERS, {
      id: id, username: username, password: hashPwd(body.password || '0000'),
      name: name, active: true, createdAt: new Date().toISOString()
    });
    return { ok: true, id: id, username: username };
  });
}

// 編輯老師帳號：姓名、登入帳號、重設密碼（選填）
// 重設密碼 → 這位老師所有已登入的裝置都會被登出；只改姓名/帳號 → 已登入的裝置繼續使用並套用新姓名
function actionEditTeacher(body) {
  requireSession(body.token, 'admin');
  var all = sheetToObjects(SHEET_TEACHERS);
  var t = all.filter(function (x) { return x.id === body.teacherId && isActive_(x.active); })[0];
  if (!t) return { ok: false, error: 'TEACHER_NOT_FOUND' };
  var name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'NAME_REQUIRED' };
  var username = String(body.username || '').trim();
  if (!username) return { ok: false, error: 'USERNAME_REQUIRED' };
  var lower = username.toLowerCase();
  var taken = all.filter(function (x) { return x.id !== t.id && isActive_(x.active); }).concat(sheetToObjects(SHEET_ADMINS))
    .some(function (u) { return String(u.username).toLowerCase() === lower; });
  if (taken) return { ok: false, error: 'USERNAME_TAKEN' };
  var newPwd = String(body.password || '').trim();
  var row = Object.assign({}, t, { name: name, username: username });
  if (newPwd) row.password = hashPwd(newPwd);
  delete row._row;
  updateObjectByRow(SHEET_TEACHERS, t._row, row);
  // 已登入的裝置：重設密碼就登出，否則更新登入資料中的姓名與帳號
  var cache = CacheService.getScriptCache();
  var sessions = sheetToObjects(SHEET_SESSIONS).filter(function (x) { return x.userId === t.id; });
  cache.removeAll(sessions.map(function (x) { return 'sess_' + x.token; }).concat(['user_' + t.id]));
  if (newPwd) deleteRows_(SHEET_SESSIONS, sessions);
  else sessions.forEach(function (x) {
    var o = Object.assign({}, x, { name: name, username: username });
    delete o._row;
    updateObjectByRow(SHEET_SESSIONS, x._row, o);
  });
  return { ok: true, passwordReset: !!newPwd, loggedOut: newPwd ? sessions.length : 0 };
}

function deleteRows_(sheetName, rowObjs) {
  touch_(sheetName);
  var sheet = sheet_(sheetName);
  rowObjs.map(function (o) { return o._row; }).sort(function (a, b) { return b - a; })
    .forEach(function (r) { sheet.deleteRow(r); });
}

// 真正刪除老師帳號（點名紀錄裡的老師姓名仍保留），編號釋出可再使用
function actionDeleteTeachersBatch(body) {
  requireSession(body.token, 'admin');
  var ids = {};
  (body.teacherIds || []).forEach(function (id) { ids[id] = true; });
  var targets = sheetToObjects(SHEET_TEACHERS).filter(function (t) { return ids[t.id]; });
  deleteRows_(SHEET_TEACHERS, targets);
  CacheService.getScriptCache().removeAll(targets.map(function (t) { return 'user_' + t.id; }));
  return { ok: true, deleted: targets.length };
}

function actionDeleteTeacher(body) {
  var r = actionDeleteTeachersBatch({ token: body.token, teacherIds: [body.teacherId] });
  return r.deleted ? { ok: true } : { ok: false, error: 'TEACHER_NOT_FOUND' };
}

// ---------- 聯絡人類型 (預設 + 自訂) ----------
function customContactTypes_() {
  try {
    var arr = JSON.parse(getSetting_('contactTypes', '[]'));
    return Array.isArray(arr) ? arr.map(String) : [];
  } catch (e) { return []; }
}

function contactTypesResult_() {
  var custom = customContactTypes_().filter(function (c) { return DEFAULT_CONTACT_TYPES.indexOf(c) === -1; });
  return { ok: true, types: DEFAULT_CONTACT_TYPES.concat(custom), custom: custom };
}

function actionGetContactTypes(body) {
  requireSession(body.token);
  return contactTypesResult_();
}

function actionAddContactType(body) {
  requireSession(body.token, 'admin');
  var name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'NAME_REQUIRED' };
  if (name.length > 12) return { ok: false, error: 'NAME_TOO_LONG' };
  withLock_(function () {
    var custom = customContactTypes_();
    if (DEFAULT_CONTACT_TYPES.indexOf(name) === -1 && custom.indexOf(name) === -1) {
      custom.push(name);
      setSetting_('contactTypes', JSON.stringify(custom));
    }
  });
  var r = contactTypesResult_();
  r.added = name;
  return r;
}

// ---------- 學校 ----------
function activeSchools_() {
  return sheetToObjects(SHEET_SCHOOLS).filter(function (sc) { return isActive_(sc.active); });
}

function actionListSchools(body) {
  requireSession(body.token);
  var counts = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (isActive_(st.active)) counts[st.schoolId] = (counts[st.schoolId] || 0) + 1;
  });
  var out = activeSchools_().map(function (sc) {
    return { id: sc.id, name: sc.name, studentCount: counts[sc.id] || 0 };
  });
  out.sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'zh-Hant'); });
  return { ok: true, schools: out };
}

function actionAddSchool(body) {
  requireSession(body.token, 'admin');
  if (!String(body.name || '').trim()) return { ok: false, error: 'NAME_REQUIRED' };
  var r = actionAddSchoolsBatch({ token: body.token, names: [body.name] });
  if (!r.created.length) return { ok: false, error: 'SCHOOL_EXISTS' };
  return { ok: true };
}

function actionAddSchoolsBatch(body) {
  requireSession(body.token, 'admin');
  var names = (body.names || []).map(function (n) { return String(n || '').trim(); }).filter(Boolean);
  if (!names.length) return { ok: false, error: 'NAME_REQUIRED' };
  return withLock_(function () {
    var existing = activeSchools_().map(function (s) { return s.name; });
    var created = [], skipped = [], rows = [];
    names.forEach(function (name) {
      if (existing.indexOf(name) !== -1 || created.indexOf(name) !== -1) { skipped.push(name); return; }
      rows.push({ id: genId('S'), name: name, createdAt: new Date().toISOString(), active: true });
      created.push(name);
    });
    appendObjects(SHEET_SCHOOLS, rows);
    return { ok: true, created: created, skipped: skipped };
  });
}

// 修改校名：不可空白、不可與其他學校同名
function actionEditSchool(body) {
  requireSession(body.token, 'admin');
  var name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'NAME_REQUIRED' };
  var all = sheetToObjects(SHEET_SCHOOLS);
  var target = all.filter(function (sc) { return sc.id === body.schoolId && isActive_(sc.active); })[0];
  if (!target) return { ok: false, error: 'SCHOOL_NOT_FOUND' };
  var dup = all.some(function (sc) { return sc.id !== target.id && isActive_(sc.active) && sc.name === name; });
  if (dup) return { ok: false, error: 'SCHOOL_EXISTS' };
  target.name = name;
  updateObjectByRow(SHEET_SCHOOLS, target._row, target);
  return { ok: true };
}

// 刪除學校：標記刪除（歷史點名紀錄仍可查看校名），該校學生一併標記刪除
function actionDeleteSchool(body) {
  requireSession(body.token, 'admin');
  var target = sheetToObjects(SHEET_SCHOOLS).filter(function (sc) { return sc.id === body.schoolId && isActive_(sc.active); })[0];
  if (!target) return { ok: false, error: 'SCHOOL_NOT_FOUND' };
  target.active = false;
  updateObjectByRow(SHEET_SCHOOLS, target._row, target);
  var ids = activeStudentsOf_(target.id).map(function (st) { return st.id; });
  var r = ids.length ? actionDeleteStudentsBatch({ token: body.token, studentIds: ids }) : { deleted: 0 };
  return { ok: true, deletedStudents: r.deleted };
}

// ---------- 學生 ----------
// 年級班級字串(如 "306")：百位數為年級，個位/十位為班級；單純數字 1~9 視為年級
function gradeNum_(g) {
  var n = parseInt(String(g).trim(), 10);
  if (isNaN(n)) return 0;
  return n >= 100 ? Math.floor(n / 100) : n;
}
function classNum_(g) {
  var n = parseInt(String(g).trim(), 10);
  if (isNaN(n)) return 0;
  return n >= 100 ? n % 100 : 0;
}

function sortStudents(list) {
  list.sort(function (a, b) {
    var ga = gradeNum_(a.grade), gb = gradeNum_(b.grade);
    if (ga !== gb) return ga - gb;
    var ca = classNum_(a.grade), cb = classNum_(b.grade);
    if (ca !== cb) return ca - cb;
    return String(a.name).localeCompare(String(b.name), 'zh-Hant');
  });
  return list;
}

// contacts: [{rel, phone, primary}]，最多 3 組，恰有一組主要聯絡人
function normalizeContacts(raw) {
  var arr = (Array.isArray(raw) ? raw : []).map(function (c) {
    c = c || {};
    return {
      rel: String(c.rel !== undefined ? c.rel : (c.name || '')).trim(),
      phone: String(c.phone || '').trim(),
      primary: c.primary === true
    };
  }).filter(function (c) { return c.rel || c.phone; }).slice(0, 3);
  if (arr.length) {
    var firstPrimary = -1;
    arr.forEach(function (c, i) { if (c.primary && firstPrimary === -1) firstPrimary = i; });
    if (firstPrimary === -1) firstPrimary = 0;
    arr.forEach(function (c, i) { c.primary = (i === firstPrimary); });
  }
  return arr;
}

function parseContacts(row) {
  var contacts = [];
  if (row.contacts) {
    try { contacts = JSON.parse(row.contacts); } catch (e) { contacts = []; }
  }
  if ((!contacts || !contacts.length) && (row.contactName || row.contactPhone)) {
    contacts = [{ rel: row.contactName || '', phone: row.contactPhone || '', primary: true }];
  }
  return normalizeContacts(contacts);
}

function normalizeDays_(arr) {
  var seen = {};
  (Array.isArray(arr) ? arr : []).forEach(function (d) {
    d = parseInt(d, 10);
    if (d >= 1 && d <= 5) seen[d] = true;   // 固定不進班只有星期一到五
  });
  return Object.keys(seen).map(Number).sort();
}

function parseDays_(v) {
  return String(v === undefined || v === null ? '' : v).split(/[,|]/).filter(String).map(Number);
}

function studentOut_(x) {
  return {
    id: x.id, name: x.name, grade: String(x.grade), gradeNum: gradeNum_(x.grade),
    contacts: parseContacts(x), fixedOffDays: parseDays_(x.fixedOffDays), note: String(x.note || '')
  };
}

// 學生必填：姓名、年級班級、至少一組「聯絡人 + 聯絡方式」
function validateStudent_(name, grade, contacts) {
  if (!String(name || '').trim()) return 'NAME_REQUIRED';
  if (!String(grade || '').trim()) return 'GRADE_REQUIRED';
  var ok = normalizeContacts(contacts).some(function (c) { return c.rel && c.phone; });
  if (!ok) return 'CONTACT_REQUIRED';
  return null;
}

function actionListStudents(body) {
  requireSession(body.token);
  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) {
    return x.schoolId === body.schoolId && isActive_(x.active);
  }).map(studentOut_);
  sortStudents(students);
  return { ok: true, students: students };
}

function studentRow_(schoolId, st) {
  return {
    id: genId('T'), schoolId: schoolId, name: String(st.name || '').trim(), grade: String(st.grade || '').trim(),
    contacts: JSON.stringify(normalizeContacts(st.contacts)),
    fixedOffDays: normalizeDays_(st.fixedOffDays).join('|'), active: true, createdAt: new Date().toISOString(),
    note: String(st.note || '').trim()
  };
}

function actionAddStudent(body) {
  requireSession(body.token, 'admin');
  var err = validateStudent_(body.name, body.grade, body.contacts);
  if (err) return { ok: false, error: err };
  var row = studentRow_(body.schoolId, body);
  appendObject(SHEET_STUDENTS, row);
  return { ok: true, id: row.id };
}

function actionAddStudentsBatch(body) {
  requireSession(body.token, 'admin');
  var list = (body.students || []).filter(function (st) { return st && String(st.name || '').trim(); });
  // 任何一列不完整就整批不寫入，並回報第幾列
  var bad = [];
  list.forEach(function (st, i) {
    var err = validateStudent_(st.name, st.grade, st.contacts);
    if (err) bad.push({ index: i, name: String(st.name).trim(), error: err });
  });
  if (bad.length) return { ok: false, error: 'INVALID_ROWS', rows: bad };
  return withLock_(function () {
    var existing = {};
    sheetToObjects(SHEET_STUDENTS).forEach(function (x) {
      if (x.schoolId === body.schoolId && isActive_(x.active)) existing[x.name + '|' + String(x.grade)] = true;
    });
    var rows = [], skipped = [];
    list.forEach(function (st) {
      var name = String(st.name || '').trim();
      var key = name + '|' + String(st.grade || '').trim();
      if (existing[key]) { skipped.push(name); return; }
      existing[key] = true;
      rows.push(studentRow_(body.schoolId, st));
    });
    appendObjects(SHEET_STUDENTS, rows);
    return { ok: true, created: rows.length, skipped: skipped };
  });
}

function actionEditStudent(body) {
  requireSession(body.token, 'admin');
  var students = sheetToObjects(SHEET_STUDENTS);
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === body.studentId) {
      var u = students[i];
      var name = body.name !== undefined ? String(body.name).trim() : u.name;
      var grade = body.grade !== undefined ? String(body.grade).trim() : String(u.grade);
      var contacts = body.contacts !== undefined ? body.contacts : parseContacts(u);
      var err = validateStudent_(name, grade, contacts);
      if (err) return { ok: false, error: err };
      u.name = name;
      u.grade = grade;
      u.contacts = JSON.stringify(normalizeContacts(contacts));
      if (body.fixedOffDays !== undefined) u.fixedOffDays = normalizeDays_(body.fixedOffDays).join('|');
      if (body.note !== undefined) u.note = String(body.note || '').trim();
      updateObjectByRow(SHEET_STUDENTS, u._row, u);
      return { ok: true };
    }
  }
  return { ok: false, error: 'STUDENT_NOT_FOUND' };
}

// 刪除學生：標記為已刪除（保留歷史點名紀錄可對照），一次寫回整欄
function actionDeleteStudentsBatch(body) {
  requireSession(body.token, 'admin');
  var ids = {};
  (body.studentIds || []).forEach(function (id) { ids[id] = true; });
  var sheet = sheet_(SHEET_STUDENTS);
  var values = sheet.getDataRange().getValues();
  var idCol = values[0].indexOf('id'), actCol = values[0].indexOf('active');
  var col = [], n = 0;
  for (var i = 1; i < values.length; i++) {
    var v = values[i][actCol];
    if (ids[values[i][idCol]] && isActive_(v)) { v = 'false'; n++; }
    col.push([toCell_(v)]);
  }
  if (n) { touch_(SHEET_STUDENTS); sheet.getRange(2, actCol + 1, col.length, 1).setNumberFormat('@').setValues(col); }
  return { ok: true, deleted: n };
}

function actionDeleteStudent(body) {
  var r = actionDeleteStudentsBatch({ token: body.token, studentIds: [body.studentId] });
  return r.deleted ? { ok: true } : { ok: false, error: 'STUDENT_NOT_FOUND' };
}

// ---------- 點名日：例假日 / 補班補課 / 國定假日 ----------
var WEEKDAY_NAMES = ['', '一', '二', '三', '四', '五', '六', '日'];

function specialDays_() {
  return sheetToObjects(SHEET_SPECIAL).map(function (x) {
    return { date: String(x.date), type: String(x.type), reason: String(x.reason || ''), createdBy: x.createdBy, createdAt: String(x.createdAt || '') };
  });
}

// 某一天能不能點名：補班補課 > 國定假日 > 週六日例假日 > 平日
function dayInfo_(date, list) {
  var sp = (list || specialDays_()).filter(function (x) { return x.date === date; })[0];
  var wd = weekdayOf(date);
  var base = { date: date, weekday: wd, reason: sp ? sp.reason : '' };
  if (sp && sp.type === 'workday') return Object.assign(base, { open: true, type: 'workday', label: '補班補課' });
  if (sp && sp.type === 'holiday') return Object.assign(base, { open: false, type: 'holiday', label: '國定假日放假' });
  if (wd >= 6) return Object.assign(base, { open: false, type: 'weekend', label: '例假日（星期' + WEEKDAY_NAMES[wd] + '）放假' });
  return Object.assign(base, { open: true, type: 'normal', label: '' });
}

function actionGetDayInfo(body) {
  requireSession(body.token);
  var date = body.date && /^\d{4}-\d{2}-\d{2}$/.test(body.date) ? body.date : today_();
  return { ok: true, today: today_(), day: dayInfo_(date) };
}

function actionListSpecialDays(body) {
  requireSession(body.token, 'admin');
  var list = specialDays_().sort(function (a, b) { return b.date.localeCompare(a.date); });
  list.forEach(function (x) { x.weekday = weekdayOf(x.date); });
  return { ok: true, today: today_(), todayInfo: dayInfo_(today_(), list), days: list };
}

function actionSetSpecialDay(body) {
  var s = requireSession(body.token, 'admin');
  var date = String(body.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'DATE_REQUIRED' };
  var type = body.type;
  if (type !== 'workday' && type !== 'holiday') return { ok: false, error: 'UNKNOWN_ACTION' };
  var wd = weekdayOf(date);
  if (type === 'workday' && wd < 6) return { ok: false, error: 'NOT_WEEKEND' };   // 平日本來就可以點名
  if (type === 'holiday' && wd >= 6) return { ok: false, error: 'NOT_WEEKDAY' };  // 週末本來就放假
  var rows = sheetToObjects(SHEET_SPECIAL).filter(function (x) { return String(x.date) === date; });
  deleteRows_(SHEET_SPECIAL, rows);
  appendObject(SHEET_SPECIAL, { date: date, type: type, reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString() });
  return { ok: true, day: dayInfo_(date) };
}

function actionDeleteSpecialDay(body) {
  requireSession(body.token, 'admin');
  var rows = sheetToObjects(SHEET_SPECIAL).filter(function (x) { return String(x.date) === String(body.date); });
  if (!rows.length) return { ok: false, error: 'DAY_NOT_FOUND' };
  deleteRows_(SHEET_SPECIAL, rows);
  return { ok: true, day: dayInfo_(String(body.date)) };
}

// 依星期自動決定設定種類：週六日 → 補班補課；平日 → 國定假日放假
function autoDayType_(date) { return weekdayOf(date) >= 6 ? 'workday' : 'holiday'; }
function isDateStr_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')) && !isNaN(Date.parse(String(d) + 'T00:00:00Z')); }

// 批次新增點名日：days = [{date, type(可省略，依星期自動判斷), reason}]；同一天已有設定會覆蓋
function actionSetSpecialDaysBatch(body) {
  var s = requireSession(body.token, 'admin');
  var input = Array.isArray(body.days) ? body.days : [];
  var map = {}, order = [], bad = [];
  input.forEach(function (d, i) {
    var date = String((d && d.date) || '');
    if (!isDateStr_(date)) { bad.push({ row: i + 1, error: 'DATE_REQUIRED' }); return; }
    var wd = weekdayOf(date);
    var type = d.type || autoDayType_(date);
    if (type !== 'workday' && type !== 'holiday') { bad.push({ row: i + 1, date: date, error: 'UNKNOWN_ACTION' }); return; }
    if (type === 'workday' && wd < 6) { bad.push({ row: i + 1, date: date, error: 'NOT_WEEKEND' }); return; }
    if (type === 'holiday' && wd >= 6) { bad.push({ row: i + 1, date: date, error: 'NOT_WEEKDAY' }); return; }
    if (!map[date]) order.push(date);
    map[date] = { type: type, reason: String(d.reason || '').trim().slice(0, 30) };   // 重複日期：以最後一筆為準
  });
  if (bad.length) return { ok: false, error: 'INVALID_DAYS', bad: bad };
  if (!order.length) return { ok: false, error: 'DATE_REQUIRED' };
  var existing = sheetToObjects(SHEET_SPECIAL).filter(function (x) { return map[String(x.date)]; });
  var updated = {};
  existing.forEach(function (x) { updated[String(x.date)] = true; });
  deleteRows_(SHEET_SPECIAL, existing);
  var now = new Date().toISOString();
  order.forEach(function (date) {
    appendObject(SHEET_SPECIAL, { date: date, type: map[date].type, reason: map[date].reason, createdBy: s.name, createdAt: now });
  });
  var nUpd = Object.keys(updated).length;
  return { ok: true, created: order.length - nUpd, updated: nUpd };
}

// 修改點名日：可改日期與原因；種類依新日期的星期自動決定
function actionEditSpecialDay(body) {
  var s = requireSession(body.token, 'admin');
  var date = String(body.date || '');
  var newDate = String(body.newDate || date);
  if (!isDateStr_(newDate)) return { ok: false, error: 'DATE_REQUIRED' };
  var all = sheetToObjects(SHEET_SPECIAL);
  var rows = all.filter(function (x) { return String(x.date) === date; });
  if (!rows.length) return { ok: false, error: 'DAY_NOT_FOUND' };
  if (newDate !== date && all.some(function (x) { return String(x.date) === newDate; })) return { ok: false, error: 'DAY_EXISTS' };
  deleteRows_(SHEET_SPECIAL, rows);
  appendObject(SHEET_SPECIAL, { date: newDate, type: autoDayType_(newDate), reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString() });
  return { ok: true, day: dayInfo_(newDate) };
}

// ---------- 預約請假 ----------
var BOOKING_MAX_DAYS = 120;   // 一次最多登記 120 天
// kind：'leave' = 預約請假（LeaveBookings）、'attend' = 臨時進班（ExtraClasses），兩者共用同一套登記流程
function bkKind_(k) { return k === 'attend' ? 'attend' : 'leave'; }
function bkSheet_(kind) { return kind === 'attend' ? SHEET_EXTRAS : SHEET_BOOKINGS; }
function bookings_(kind) {
  return sheetToObjects(bkSheet_(bkKind_(kind))).map(function (b) {
    return { id: String(b.id), schoolId: String(b.schoolId), studentId: String(b.studentId), startDate: String(b.startDate),
      endDate: String(b.endDate), reason: String(b.reason || ''), createdBy: String(b.createdBy || ''), createdRole: String(b.createdRole || ''),
      createdAt: String(b.createdAt || ''), updatedBy: String(b.updatedBy || ''), _row: b._row };
  });
}
function bookingBrief_(b) { return { id: b.id, startDate: b.startDate, endDate: b.endDate, reason: b.reason }; }
// 預約期間實際請假的上課日（扣除例假日、國定假日、學生固定不進班）
// 臨時進班：期間內「原本固定不進班」且可以點名的日子
function bookingDays_(ctx, sid, start, end, kind) {
  var out = [];
  for (var d = start; d <= end && out.length <= 400; d = nextDate_(d)) {
    if (!ctx.isOpen(d)) continue;
    if (kind === 'attend' ? ctx.isFixedBase(sid, d) : !ctx.isFixed(sid, d)) out.push(d);
  }
  return out;
}
function bookingOut_(ctx, b, schoolMap, kind) {
  var st = ctx.stMap[b.studentId];
  var so = st ? studentOut_(st) : { name: '(已刪除的學生)', grade: '' };
  var today = ctx.today;
  var state = b.endDate < today ? 'past' : (b.startDate <= today ? 'ongoing' : 'upcoming');
  return Object.assign(bookingBrief_(b), {
    schoolId: b.schoolId, schoolName: schoolMap[b.schoolId] || '(已移除的學校)', studentId: b.studentId,
    name: so.name, grade: String(so.grade), days: bookingDays_(ctx, b.studentId, b.startDate, b.endDate, kind).length,
    createdBy: b.createdBy, createdRole: b.createdRole, updatedBy: b.updatedBy, state: state
  });
}

// 列出預約請假 / 臨時進班（管理員與老師都可以看）；body: { kind?, schoolId?, includePast? }
function actionListBookings(body) {
  requireSession(body.token);
  var kind = bkKind_(body.kind);
  var ctx = statusCtx_();
  var schoolMap = {};
  sheetToObjects(SHEET_SCHOOLS).forEach(function (sc) { schoolMap[sc.id] = sc.name; });
  var items = bookings_(kind).filter(function (b) {
    var st = ctx.stMap[b.studentId];
    if (!st || !isActive_(st.active)) return false;
    if (body.schoolId && b.schoolId !== body.schoolId) return false;
    return body.includePast || b.endDate >= ctx.today;
  }).map(function (b) { return bookingOut_(ctx, b, schoolMap, kind); });
  var rank = { ongoing: 0, upcoming: 1, past: 2 };
  items.sort(function (a, b) {
    return (rank[a.state] - rank[b.state]) ||
      (a.state === 'past' ? b.endDate.localeCompare(a.endDate) : a.startDate.localeCompare(b.startDate)) ||
      String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant') || (gradeNum_(a.grade) - gradeNum_(b.grade)) ||
      (classNum_(a.grade) - classNum_(b.grade)) || String(a.name).localeCompare(String(b.name), 'zh-Hant');
  });
  return { ok: true, kind: kind, today: ctx.today, oldest: dateStrDaysAgo_(RETENTION_DAYS), items: items };
}

// 檢查一筆預約請假 / 臨時進班；all = 同類型既有的登記（含本次批次已通過的）
// 回傳 { error, clash? } 或 { st, start, end, reason, days, old }
function checkBooking_(s, ctx, kind, input, all) {
  var isAdmin = s.role === 'admin';
  var today = ctx.today;
  var start = String(input.startDate || ''), end = String(input.endDate || input.startDate || '');
  if (!isDateStr_(start) || !isDateStr_(end)) return { error: 'DATE_REQUIRED' };
  if (end < start) return { error: 'BOOKING_RANGE' };
  var st = ctx.stMap[String(input.studentId || '')];
  if (!st || !isActive_(st.active)) return { error: 'STUDENT_NOT_FOUND' };
  if (!isAdmin && input.schoolId && st.schoolId !== input.schoolId) return { error: 'STUDENT_NOT_FOUND' };
  if (kind === 'attend' && !parseDays_(st.fixedOffDays).length) return { error: 'EXTRA_NOT_FIXED' };
  var old = input.id ? all.filter(function (b) { return b.id === String(input.id); })[0] : null;
  if (input.id && !old) return { error: 'BOOKING_NOT_FOUND' };
  if (!isAdmin) {
    if (old && old.startDate < today) {
      // 已經開始、還沒結束的預約：老師可以調整結束日期與原因（開始日期與學生不能改，過去的紀錄要保留）
      if (old.endDate < today) return { error: 'BOOKING_STARTED' };
      if (start !== old.startDate || st.id !== old.studentId) return { error: 'BOOKING_STARTED' };
      if (end < today) return { error: 'BOOKING_PAST' };
    } else if (start < today) return { error: 'BOOKING_PAST' };
  } else if (start < dateStrDaysAgo_(RETENTION_DAYS)) return { error: 'BOOKING_TOO_OLD' };
  if ((Date.parse(end) - Date.parse(start)) / 86400000 + 1 > BOOKING_MAX_DAYS) return { error: 'BOOKING_TOO_LONG' };
  var clash = all.filter(function (b) {
    return b.studentId === st.id && (!old || b.id !== old.id) && b.startDate <= end && start <= b.endDate;
  })[0];
  if (clash) return { error: kind === 'attend' ? 'EXTRA_OVERLAP' : 'BOOKING_OVERLAP', clash: bookingBrief_(clash) };
  var days = bookingDays_(ctx, st.id, start, end, kind);
  if (!days.length) return { error: kind === 'attend' ? 'EXTRA_NO_DAYS' : 'BOOKING_NO_DAYS' };
  return { st: st, start: start, end: end, reason: String(input.reason || '').trim().slice(0, 30), days: days.length, old: old };
}
function bookingRow_(s, kind, c, now) {
  var old = c.old;
  return {
    id: old ? old.id : genId(kind === 'attend' ? 'EX' : 'BK'), schoolId: c.st.schoolId, studentId: c.st.id, startDate: c.start, endDate: c.end,
    reason: c.reason,
    createdBy: old ? old.createdBy : s.name, createdRole: old ? old.createdRole : s.role, createdAt: old ? old.createdAt : now,
    updatedBy: old ? s.name : '', updatedAt: old ? now : ''
  };
}

// 新增或修改預約請假 / 臨時進班；body: { kind?, id?, studentId, startDate, endDate, reason }
// 老師只能登記今天以後的日期；已經開始但還沒結束的預約，老師可以修改結束日期與原因（開始日期不能動），已結束的要請管理員處理
function actionSaveBooking(body) {
  var s = requireSession(body.token);
  var kind = bkKind_(body.kind);
  var ctx = statusCtx_();
  var c = checkBooking_(s, ctx, kind, body, bookings_(kind));
  if (c.error) return { ok: false, error: c.error, clash: c.clash };
  var row = bookingRow_(s, kind, c, new Date().toISOString());
  if (c.old) updateObjectByRow(bkSheet_(kind), c.old._row, row);
  else appendObject(bkSheet_(kind), row);
  return { ok: true, id: row.id, days: c.days, startDate: c.start, endDate: c.end };
}

// 批次登記：每一列各自的學生與期間；body: { kind?, schoolId?(老師端限定學校), items: [{ studentId, startDate, endDate, reason }] }
// 全部檢查通過才一起寫入；有問題的列回傳 bad: [{ row, error, clash? }]
var BOOKING_BATCH_MAX = 100;
function actionSaveBookingsBatch(body) {
  var s = requireSession(body.token);
  var kind = bkKind_(body.kind);
  var items = Array.isArray(body.items) ? body.items : [];
  if (!items.length) return { ok: false, error: 'BATCH_EMPTY' };
  if (items.length > BOOKING_BATCH_MAX) return { ok: false, error: 'BATCH_TOO_MANY' };
  var ctx = statusCtx_();
  var all = bookings_(kind);
  var now = new Date().toISOString();
  var bad = [], rows = [], totalDays = 0;
  items.forEach(function (it, i) {
    var input = { studentId: it && it.studentId, startDate: it && it.startDate, endDate: it && it.endDate, reason: it && it.reason, schoolId: body.schoolId };
    var c = checkBooking_(s, ctx, kind, input, all);
    if (c.error) {
      var inBatch = !!c.clash && rows.some(function (r) { return r.id === c.clash.id; });
      bad.push({ row: i + 1, error: c.error, clash: c.clash || null, inBatch: inBatch });
      return;
    }
    var row = bookingRow_(s, kind, c, now);
    rows.push(row);
    all.push(row);   // 同一批裡同一位學生的期間也不能重疊
    totalDays += c.days;
  });
  if (bad.length) return { ok: false, error: 'INVALID_BOOKINGS', bad: bad };
  appendObjects(bkSheet_(kind), rows);
  return { ok: true, created: rows.length, days: totalDays };
}

function actionDeleteBooking(body) {
  var s = requireSession(body.token);
  var kind = bkKind_(body.kind);
  var b = bookings_(kind).filter(function (x) { return x.id === String(body.id || ''); })[0];
  if (!b) return { ok: false, error: 'BOOKING_NOT_FOUND' };
  var today = today_();
  if (s.role !== 'admin' && b.startDate < today) {
    // 已經開始的預約：老師「取消」＝ 從今天起不再請假（結束日期改為昨天），之前已過去的日子照舊保留
    if (b.endDate < today) return { ok: false, error: 'BOOKING_STARTED' };
    var row = bookingRow_(s, kind, { old: b, st: { id: b.studentId, schoolId: b.schoolId }, start: b.startDate, end: prevDate_(today), reason: b.reason }, new Date().toISOString());
    updateObjectByRow(bkSheet_(kind), b._row, row);
    return { ok: true, truncated: true };
  }
  deleteRows_(bkSheet_(kind), [b]);
  return { ok: true };
}

// ---------- 點名 ----------
// 自送/自走本身就是一種出席狀態；相容舊資料(狀態=到班 + 自送/自走旗標)
function effectiveStatus_(a) {
  var s = String(a.status || '');
  if (!s || s === 'present') {
    if (isTrue_(a.selfDrop)) return 'selfDrop';
    if (isTrue_(a.selfWalk)) return 'selfWalk';
  }
  return s;
}

// 連續請假天數：從 date 往前數「可以點名的日子」，跳過例假日/國定假日與學生固定不進班的日子；
// 當天必須是請假，往前遇到出席或沒有點名紀錄就停止。回傳 { studentId: 天數 }
var LEAVE_STREAK_ALERT = 3;
function prevDate_(d) { return Utilities.formatDate(new Date(Date.parse(d + 'T12:00:00+08:00') - 86400000), TZ, 'yyyy-MM-dd'); }
// ---------- 每位學生每天的狀態（全系統共用同一套判斷）----------
// 1. 老師點過的狀態（到班、自送、自走、請假）優先
// 2. 預約請假期間（只算可以點名、且不是學生固定不進班的日子）→ 請假，老師沒送出也算
// 3. 今天：跟著學生目前的固定不進班設定；過去的日子：照當天的紀錄
// 臨時進班期間：原本固定不進班的日子改為正常進班（視同不是固定不進班）
function statusCtx_() {
  var today = today_();
  var rec = {};
  sheetToObjects(SHEET_ATTENDANCE).forEach(function (a) { rec[String(a.date) + '|' + a.studentId] = a; });
  var fixed = {}, stMap = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) { fixed[st.id] = parseDays_(st.fixedOffDays); stMap[st.id] = st; });
  var sp = specialDays_(), openCache = {}, wdCache = {};
  var wdOf = function (d) { if (!(d in wdCache)) wdCache[d] = weekdayOf(d); return wdCache[d]; };
  var isOpen = function (d) { if (!(d in openCache)) openCache[d] = dayInfo_(d, sp).open; return openCache[d]; };
  var isFixedBase = function (sid, d) { return (fixed[sid] || []).indexOf(wdOf(d)) !== -1; };   // 每週固定不進班設定
  var ex = {};
  bookings_('attend').forEach(function (b) { (ex[b.studentId] = ex[b.studentId] || []).push(b); });
  var extraOn = function (sid, d) {   // 這天有登記臨時進班（只在原本固定不進班、且可以點名的日子生效）
    if (!isFixedBase(sid, d) || !isOpen(d)) return null;
    var list = ex[sid] || [];
    for (var i = 0; i < list.length; i++) if (list[i].startDate <= d && d <= list[i].endDate) return list[i];
    return null;
  };
  var isFixed = function (sid, d) { return isFixedBase(sid, d) && !extraOn(sid, d); };
  var bk = {};
  bookings_().forEach(function (b) { (bk[b.studentId] = bk[b.studentId] || []).push(b); });
  var bookingOn = function (sid, d) {
    var list = bk[sid] || [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].startDate <= d && d <= list[i].endDate) return (isOpen(d) && !isFixed(sid, d)) ? list[i] : null;
    }
    return null;
  };
  var teacherSet = function (a) {   // 老師親自點的狀態（不含系統預設值）
    if (!a) return '';
    var v = effectiveStatus_(a);
    return v && v !== 'fixedOff' && !isTrue_(a.auto) ? v : '';
  };
  var status = function (sid, d) {
    var a = rec[d + '|' + sid];
    var mine = teacherSet(a);
    if (mine) return mine;
    if (bookingOn(sid, d)) return 'leave';
    if (d === today) return isFixed(sid, d) ? 'fixedOff' : '';
    var saved = a ? effectiveStatus_(a) : '';
    if (saved === 'fixedOff') return extraOn(sid, d) ? '' : 'fixedOff';   // 事後補登臨時進班 → 改回未點名
    if (a && isTrue_(a.auto) && saved === 'leave') return '';   // 預約請假後來被取消
    return saved;
  };
  // 系統預設值：預約請假 > 固定不進班
  var defaultOf = function (sid, d) { return bookingOn(sid, d) ? 'leave' : (isFixed(sid, d) ? 'fixedOff' : ''); };
  return { today: today, rec: rec, fixed: fixed, stMap: stMap, isOpen: isOpen, isFixed: isFixed, isFixedBase: isFixedBase, wdOf: wdOf,
    extras: ex, extraOn: extraOn, bookings: bk, bookingOn: bookingOn, teacherSet: teacherSet, status: status, defaultOf: defaultOf };
}
function leaveStreakCtx_() { return statusCtx_(); }
function leaveStreaks_(date, studentIds, ctx) {
  var out = {};
  if (!studentIds.length) return out;
  ctx = ctx || statusCtx_();
  var oldest = dateStrDaysAgo_(RETENTION_DAYS + 1);
  studentIds.forEach(function (sid) {
    if (ctx.status(sid, date) !== 'leave') return;
    var n = 0, d = date;
    while (d >= oldest) {
      var st = ctx.status(sid, d);
      if (st === 'leave') n++;
      else if (st === 'fixedOff') { /* 固定不進班：不中斷也不計入 */ }
      else if (!st && (!ctx.isOpen(d) || ctx.isFixed(sid, d))) { /* 放假日或固定不進班日沒點名：跳過 */ }
      else break;
      d = prevDate_(d);
    }
    out[sid] = n;
  });
  return out;
}

function nextDate_(d) { return Utilities.formatDate(new Date(Date.parse(d + 'T12:00:00+08:00') + 86400000), TZ, 'yyyy-MM-dd'); }

// 管理端「連續請假名單」：列出保留期間內每一段連續請假（給管理員評估退費）
// body: { schoolId?, minDays?(預設 3), from?, to? } → 期間有重疊的才列出
function actionListLeaveStreaks(body) {
  requireSession(body.token, 'admin');
  var today = today_();
  var minDays = Math.max(1, Number(body.minDays) || LEAVE_STREAK_ALERT);
  var oldest = dateStrDaysAgo_(RETENTION_DAYS);
  var from = isDateStr_(body.from) ? String(body.from) : oldest;
  var to = isDateStr_(body.to) ? String(body.to) : today;
  var ctx = statusCtx_();
  var schoolMap = {};
  sheetToObjects(SHEET_SCHOOLS).forEach(function (sc) { schoolMap[sc.id] = sc.name; });
  // 有出現過請假紀錄或預約請假的學生才需要計算
  var cand = {};
  sheetToObjects(SHEET_ATTENDANCE).forEach(function (a) {
    if (effectiveStatus_(a) !== 'leave' || String(a.date) < oldest) return;
    if (body.schoolId && a.schoolId !== body.schoolId) return;
    cand[a.studentId] = a.schoolId;
  });
  Object.keys(ctx.bookings).forEach(function (sid) {
    ctx.bookings[sid].forEach(function (b) {
      if (b.endDate < oldest || b.startDate > today) return;
      if (body.schoolId && b.schoolId !== body.schoolId) return;
      cand[sid] = cand[sid] || b.schoolId;
    });
  });
  var stMap = ctx.stMap;
  var dates = [];
  for (var d = oldest; d <= today; d = nextDate_(d)) dates.push(d);
  var out = [];
  Object.keys(cand).forEach(function (sid) {
    var run = [], booked = [];
    var close = function (ongoing) {
      if (run.length >= minDays && run[run.length - 1] >= from && run[0] <= to) {
        var st = stMap[sid];
        var so = st ? studentOut_(st) : { name: '(已刪除的學生)', grade: '' };
        out.push({
          studentId: sid, schoolId: cand[sid], schoolName: schoolMap[cand[sid]] || '(已移除的學校)',
          name: so.name, grade: String(so.grade), start: run[0], end: run[run.length - 1],
          days: run.length, dates: run.slice(), bookedDates: booked.slice(), ongoing: !!ongoing
        });
      }
      run = []; booked = [];
    };
    dates.forEach(function (d) {
      var s = ctx.status(sid, d);
      if (s === 'leave') { run.push(d); if (ctx.bookingOn(sid, d) && !ctx.teacherSet(ctx.rec[d + '|' + sid])) booked.push(d); }
      else if (s === 'fixedOff') { /* 固定不進班：不中斷 */ }
      else if (!s && (d === today || !ctx.isOpen(d) || ctx.isFixed(sid, d))) { /* 今天還沒點名、放假日、固定不進班日：不中斷 */ }
      else close(false);
    });
    close(true);
  });
  out.sort(function (a, b) {
    return (b.ongoing - a.ongoing) || (b.days - a.days) ||
      String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant') ||
      (gradeNum_(a.grade) - gradeNum_(b.grade)) || (classNum_(a.grade) - classNum_(b.grade)) ||
      String(a.name).localeCompare(String(b.name), 'zh-Hant') || a.start.localeCompare(b.start);
  });
  return { ok: true, today: today, oldest: oldest, from: from, to: to, minDays: minDays, items: out };
}

// 今天的狀態要跟著學生「目前」的固定不進班設定：
// 還沒點名（含先送出部分名單時存下的空白）且今天固定不進班 → 固定不進班；
// 存的是固定不進班（老師沒改過的預設值）但管理員已取消今天的設定 → 改回未點名
// （老師端按鈕沒有「固定不進班」，所以存成 fixedOff 一定是系統預設值，可以安全調整）
function todayStatus_(status, isFixedNow) {
  status = status || '';
  if (!status && isFixedNow) return 'fixedOff';
  if (status === 'fixedOff' && !isFixedNow) return '';
  return status;
}

function activeStudentsOf_(schoolId) {
  return sheetToObjects(SHEET_STUDENTS).filter(function (x) {
    return x.schoolId === schoolId && isActive_(x.active);
  });
}

// 送出點名的老師清單：同一位老師多次送出只列一次（取最後一次時間），不同老師依時間先後全部列出
// 回傳 { 'date|schoolId': [{ name, at }] }
function submittersIndex_(cutoff, attendanceRows) {
  var idx = {};
  var add = function (key, tid, name, at) {
    if (!name) return;
    var m = idx[key] = idx[key] || {};
    var k = tid || name;
    if (!m[k] || at > m[k].at) m[k] = { name: name, at: at };
  };
  var logged = {};
  sheetToObjects(SHEET_SUBMISSIONS).forEach(function (l) {
    var d = String(l.date);
    if (cutoff && d < cutoff) return;
    var key = d + '|' + l.schoolId;
    logged[key] = true;
    add(key, l.teacherId, l.teacherName, String(l.submittedAt));
  });
  // 舊資料沒有送出紀錄時，從點名資料推算
  (attendanceRows || []).forEach(function (a) {
    var key = String(a.date) + '|' + a.schoolId;
    if (logged[key] || (cutoff && String(a.date) < cutoff)) return;
    add(key, a.teacherId, a.teacherName, String(a.submittedAt));
  });
  var out = {};
  Object.keys(idx).forEach(function (key) {
    out[key] = Object.keys(idx[key]).map(function (k) { return idx[key][k]; })
      .sort(function (a, b) { return a.at.localeCompare(b.at); });
  });
  return out;
}

function actionGetRollCall(body) {
  requireSession(body.token);
  var date = today_();
  var schoolId = body.schoolId;
  var wd = weekdayOf(date);
  var attendance = sheetToObjects(SHEET_ATTENDANCE).filter(function (a) {
    return a.schoolId === schoolId && String(a.date) === date;
  });
  var attMap = {};
  attendance.forEach(function (a) { attMap[a.studentId] = a; });

  var ctx = statusCtx_();
  var out = activeStudentsOf_(schoolId).map(function (st) {
    var so = studentOut_(st);
    var isFixedOff = ctx.isFixed(st.id, date);   // 臨時進班的日子不算固定不進班
    var exb = ctx.extraOn(st.id, date);
    var existing = attMap[st.id];
    var primary = so.contacts.filter(function (c) { return c.primary; })[0] || null;
    var status = ctx.status(st.id, date);
    var byTeacher = !!ctx.teacherSet(existing);   // 系統自動判斷的不顯示點名老師
    var bk = ctx.bookingOn(st.id, date);
    return {
      id: so.id, name: so.name, grade: so.grade, isFixedOff: isFixedOff, primaryContact: primary,
      contacts: so.contacts, note: so.note, status: status, defaultStatus: ctx.defaultOf(st.id, date), fixedOffDays: so.fixedOffDays,
      booking: bk ? bookingBrief_(bk) : null, extra: exb ? bookingBrief_(exb) : null,
      markedBy: byTeacher ? existing.teacherName : '', markedAt: byTeacher ? String(existing.submittedAt) : ''
    };
  });
  sortStudents(out);

  var submitters = submittersIndex_(date, attendance)[date + '|' + schoolId] || [];
  var last = submitters[submitters.length - 1] || null;
  return {
    ok: true, date: date, day: dayInfo_(date), students: out, submitted: submitters.length > 0, submitters: submitters,
    submittedAt: last ? last.at : null, submittedBy: last ? last.name : null
  };
}

// 點一下就自動儲存：老師每點一位學生就送出（前端會把短時間內的多次點選合併成一次）
function actionSaveMarks(body) {
  var r = actionSubmitAttendance(Object.assign({}, body, { autosave: true }));
  if (r && r.ok) { try { putPresence_(requireSession(body.token, 'teacher'), body.schoolId, 'edit'); } catch (e) {} }
  return r;
}

function actionSubmitAttendance(body) {
  var s = requireSession(body.token, 'teacher');
  var date = today_();
  var schoolId = body.schoolId;
  var now = new Date().toISOString();
  var day = dayInfo_(date);
  if (!day.open) return { ok: false, error: 'DAY_CLOSED', day: day };   // 例假日 / 國定假日不能點名

  // 整段在全域鎖內：兩位老師同時送出同一校時，第二位會看到「已被他人更新」而不是默默覆蓋
  var result = withLock_(function () {
    var students = activeStudentsOf_(schoolId);
    // 同一位學生若重複出現，以最後一筆為準
    var latestOf = {};
    (body.records || []).forEach(function (r) { if (r) latestOf[r.studentId] = r; });
    // 允許先送出部分名單（有的學生中午接、有的下午接），未點名的之後可再補
    var ctx = statusCtx_();
    var valid = {}, fixedToday = {};
    students.forEach(function (st) {
      valid[st.id] = true;
      if (ctx.isFixed(st.id, date)) fixedToday[st.id] = true;   // 臨時進班的日子不算固定不進班
    });

    var sheet = sheet_(SHEET_ATTENDANCE);
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var idx = {};
    headers.forEach(function (h, i) { idx[h] = i; });
    var rowOf = {}, todays = [];
    for (var i = 1; i < values.length; i++) {
      var d = values[i][idx.date];
      if (d instanceof Date) d = fromDateCell_(d, 'date');
      if (String(d) === date && values[i][idx.schoolId] === schoolId) {
        rowOf[values[i][idx.studentId]] = i;
        var at = values[i][idx.submittedAt];
        todays.push({ date: date, schoolId: schoolId, teacherId: values[i][idx.teacherId], teacherName: values[i][idx.teacherName],
          submittedAt: at instanceof Date ? at.toISOString() : String(at || '') });
      }
    }
    // 載入後若已有別人送出（或更新）過，先提醒，不直接覆蓋
    var subs = submittersIndex_(date, todays)[date + '|' + schoolId] || [];
    var last = subs[subs.length - 1];
    var latest = last ? last.at : '';
    // 自動儲存：每位學生各自更新（最後點的為準），不做整校衝突檢查
    if (!body.autosave && !body.force && latest !== String(body.baseSubmittedAt || '')) {
      return { ok: false, error: 'CONFLICT', submittedAt: latest, submittedBy: last ? last.name : '' };
    }
    // 只更新「狀態有變動」的學生：每位學生保留真正幫他點名的老師與時間
    var changedRows = [], fresh = [], n = 0;
    Object.keys(latestOf).map(function (k) { return latestOf[k]; }).forEach(function (r) {
      if (!valid[r.studentId]) return;
      var status = VALID_STATUS.indexOf(r.status) === -1 ? '' : (r.status || '');
      var rec = {
        date: date, schoolId: schoolId, studentId: r.studentId, status: status,
        selfDrop: status === 'selfDrop', selfWalk: status === 'selfWalk',
        teacherId: s.userId, teacherName: s.name, submittedAt: now,
        wasFixedOff: !!fixedToday[r.studentId],   // 今天本來是固定不進班（之後若改為到班，明細會標示）
        auto: false
      };
      // 與系統預設值（預約請假 / 固定不進班）相同 → 標記為預設值，之後取消預約或設定時會跟著更新
      var def = ctx.defaultOf(r.studentId, date);
      rec.auto = !!def && status === def;
      if (rowOf[r.studentId] !== undefined) {
        var row = values[rowOf[r.studentId]];
        var old = {};
        headers.forEach(function (h, c) { old[h] = row[c]; });
        if (effectiveStatus_(old) === status && isTrue_(old.auto) === rec.auto) return;
        rec.id = row[idx.id];
        headers.forEach(function (h, c) { row[c] = toCell_(rec[h]); });
        changedRows.push(rowOf[r.studentId]);
      } else {
        rec.id = genId('R');
        fresh.push(rec);
      }
      n++;
    });
    // 只寫回有變動的列（相鄰的列合併成一段）：不再整張點名表重寫，資料越多也不會變慢
    if (changedRows.length) {
      touch_(SHEET_ATTENDANCE);
      changedRows.sort(function (a, b) { return a - b; });
      var start = changedRows[0], prev = start;
      var writeSeg = function (from, to) {
        var seg = values.slice(from, to + 1).map(function (row) { return plainRow_(headers, row); });
        sheet.getRange(from + 1, 1, seg.length, headers.length).setNumberFormat('@').setValues(seg);
      };
      for (var ci = 1; ci < changedRows.length; ci++) {
        if (changedRows[ci] === prev + 1) { prev = changedRows[ci]; continue; }
        writeSeg(start, prev);
        start = prev = changedRows[ci];
      }
      writeSeg(start, prev);
    }
    appendObjects(SHEET_ATTENDANCE, fresh);
    if (body.autosave) {
      // 自動儲存：每天每校每位老師只保留一筆送出紀錄，更新時間與筆數（不讓紀錄表暴增）
      if (n) {
        var mine = sheetToObjects(SHEET_SUBMISSIONS).filter(function (l) {
          return String(l.date) === date && l.schoolId === schoolId && l.teacherId === s.userId;
        })[0];
        if (mine) {
          var upd = Object.assign({}, mine, { teacherName: s.name, submittedAt: now, count: Number(mine.count || 0) + n });
          delete upd._row;
          updateObjectByRow(SHEET_SUBMISSIONS, mine._row, upd);
        } else appendObject(SHEET_SUBMISSIONS, { id: genId('L'), date: date, schoolId: schoolId, teacherId: s.userId, teacherName: s.name, submittedAt: now, count: n });
      }
    } else {
      appendObject(SHEET_SUBMISSIONS, { id: genId('L'), date: date, schoolId: schoolId, teacherId: s.userId, teacherName: s.name, submittedAt: now, count: n });
    }
    return { ok: true, submittedAt: now, count: n };
  });
  // 舊資料清理只交給每天凌晨 3 點的排程，不讓老師送出時等待
  return result;
}

// 每日每校統計（總覽、過去 40 天小卡、單日請假名單共用）
function overviewData_(schoolFilter, cutoff) {
  var today = today_();
  var schoolMap = {}, activeIds = [];
  sheetToObjects(SHEET_SCHOOLS).forEach(function (sc) {
    schoolMap[sc.id] = sc.name;
    if (isActive_(sc.active)) activeIds.push(sc.id);
  });
  var totals = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (isActive_(st.active)) totals[st.schoolId] = (totals[st.schoolId] || 0) + 1;
  });
  var attendance = sheetToObjects(SHEET_ATTENDANCE);
  var subIdx = submittersIndex_(cutoff, attendance);

  var ctx = statusCtx_();
  var byDate = {}, leaveIds = {}, hasRec = {};
  var bucket = function (d, sid) {
    byDate[d] = byDate[d] || {};
    return byDate[d][sid] = byDate[d][sid] || { present: 0, selfDrop: 0, selfWalk: 0, leave: 0, fixedOff: 0, marked: 0 };
  };
  var count = function (b, d, studentId, schoolId, st) {
    if (st) { b[st] = (b[st] || 0) + 1; b.marked++; }
    if (st === 'leave') (leaveIds[d] = leaveIds[d] || []).push({ sid: studentId, school: schoolId });
  };
  attendance.forEach(function (a) {
    var d = String(a.date);
    if (d < cutoff || d > today) return;
    if (schoolFilter && a.schoolId !== schoolFilter) return;
    hasRec[d + '|' + a.schoolId] = true;
    count(bucket(d, a.schoolId), d, a.studentId, a.schoolId, ctx.status(a.studentId, d));
  });
  // 預約請假：當天沒有點名紀錄也算請假
  Object.keys(ctx.bookings).forEach(function (sid) {
    var st = ctx.stMap[sid];
    if (!st || !isActive_(st.active) || (schoolFilter && st.schoolId !== schoolFilter)) return;
    ctx.bookings[sid].forEach(function (bk) {
      var from = bk.startDate < cutoff ? cutoff : bk.startDate, to = bk.endDate > today ? today : bk.endDate;
      for (var d = from; d <= to; d = nextDate_(d)) {
        if (ctx.rec[d + '|' + sid] || d === today || !ctx.bookingOn(sid, d)) continue;   // 今天另外計算
        count(bucket(d, st.schoolId), d, sid, st.schoolId, 'leave');
      }
    });
  });
  // 今天：依學生目前的設定與預約重新計算
  var schoolStudents = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (isActive_(st.active)) (schoolStudents[st.schoolId] = schoolStudents[st.schoolId] || []).push(st);
  });
  activeIds.forEach(function (sid) {
    if (schoolFilter && schoolFilter !== sid) return;
    var list = schoolStudents[sid] || [];
    var anyBooked = list.some(function (st) { return ctx.bookingOn(st.id, today); });
    if (!hasRec[today + '|' + sid] && !anyBooked) return;
    var b = { present: 0, selfDrop: 0, selfWalk: 0, leave: 0, fixedOff: 0, marked: 0 };
    leaveIds[today] = (leaveIds[today] || []).filter(function (x) { return x.school !== sid; });
    list.forEach(function (st) { count(b, today, st.id, sid, ctx.status(st.id, today)); });
    byDate[today] = byDate[today] || {};
    byDate[today][sid] = b;
  });
  // 每天各校「連續請假 3 天以上」的人數
  var longLeave = {};
  Object.keys(leaveIds).forEach(function (d) {
    var streaks = leaveStreaks_(d, leaveIds[d].map(function (x) { return x.sid; }), ctx);
    leaveIds[d].forEach(function (x) {
      if (streaks[x.sid] >= LEAVE_STREAK_ALERT) longLeave[d + '|' + x.school] = (longLeave[d + '|' + x.school] || 0) + 1;
    });
  });
  byDate[today] = byDate[today] || {};

  var dates = Object.keys(byDate).sort(function (a, b) { return b.localeCompare(a); });
  var result = dates.map(function (date) {
    var ids = Object.keys(byDate[date]);
    if (date === today) {
      activeIds.forEach(function (sid) {
        if (ids.indexOf(sid) === -1 && (!schoolFilter || schoolFilter === sid)) ids.push(sid);
      });
    }
    var schools = ids.map(function (sid) {
      var b = byDate[date][sid];
      var total = totals[sid] || 0;
      var subs = subIdx[date + '|' + sid] || [];
      var last = subs[subs.length - 1] || null;
      var item = {
        schoolId: sid, schoolName: schoolMap[sid] || '(已移除的學校)', total: total, submitted: !!hasRec[date + '|' + sid],
        present: 0, selfDrop: 0, selfWalk: 0, leave: 0, fixedOff: 0, attended: 0,
        unmarked: total, submitters: subs, submittedAt: last ? last.at : null, submittedBy: last ? last.name : null,
        longLeave: longLeave[date + '|' + sid] || 0
      };
      if (b) {
        item.present = b.present; item.selfDrop = b.selfDrop; item.selfWalk = b.selfWalk;
        item.leave = b.leave; item.fixedOff = b.fixedOff;
        item.attended = b.present + b.selfDrop + b.selfWalk;
        item.unmarked = Math.max(0, total - b.marked);
      }
      return item;
    });
    schools.sort(function (a, b) { return String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant'); });
    return { date: date, schools: schools };
  });
  return { today: today, dates: result, attendance: attendance, schoolMap: schoolMap, ctx: ctx };
}

function actionGetOverview(body) {
  requireSession(body.token, 'admin');
  var d = overviewData_(body.schoolId || null, dateStrDaysAgo_(Number(body.days) || RETENTION_DAYS));
  var sp = specialDays_();
  var special = {};
  sp.forEach(function (x) { special[x.date] = { type: x.type, reason: x.reason }; });
  return { ok: true, today: d.today, todayInfo: dayInfo_(d.today, sp), special: special, dates: d.dates, presence: presenceList_(null) };
}

// 某一天的請假名單（全部學校或單一學校）＋ 當天各校摘要
function actionGetDayLeaves(body) {
  requireSession(body.token, 'admin');
  var date = String(body.date || '');
  var schoolFilter = body.schoolId || null;
  var d = overviewData_(schoolFilter, date);
  var day = d.dates.filter(function (x) { return x.date === date; })[0] || { date: date, schools: [] };
  var ctx = d.ctx, stMap = ctx.stMap;
  // 當天有紀錄的學生 + 預約請假的學生
  var cand = {};
  d.attendance.forEach(function (a) {
    if (String(a.date) === date && (!schoolFilter || a.schoolId === schoolFilter)) cand[a.studentId] = a.schoolId;
  });
  Object.keys(ctx.bookings).forEach(function (sid) {
    var st = stMap[sid];
    if (st && isActive_(st.active) && (!schoolFilter || st.schoolId === schoolFilter) && ctx.bookingOn(sid, date)) cand[sid] = cand[sid] || st.schoolId;
  });
  if (date === d.today) {   // 今天：跟著學生目前的設定
    Object.keys(stMap).forEach(function (sid) {
      var st = stMap[sid];
      if (isActive_(st.active) && (!schoolFilter || st.schoolId === schoolFilter)) cand[sid] = cand[sid] || st.schoolId;
    });
  }
  var leaves = Object.keys(cand).filter(function (sid) { return ctx.status(sid, date) === 'leave'; }).map(function (sid) {
    var st = stMap[sid], a = ctx.rec[date + '|' + sid];
    var so = st ? studentOut_(st) : { name: '(已刪除的學生)', grade: '', contacts: [], note: '' };
    var mine = !!ctx.teacherSet(a);
    var bk = mine ? null : ctx.bookingOn(sid, date);
    return {
      schoolId: cand[sid], schoolName: d.schoolMap[cand[sid]] || '(已移除的學校)',
      id: sid, name: so.name, grade: String(so.grade), contacts: so.contacts, note: so.note,
      markedBy: mine ? a.teacherName : '', markedAt: mine ? String(a.submittedAt) : '', booking: bk ? bookingBrief_(bk) : null
    };
  });
  var streaks = leaveStreaks_(date, leaves.map(function (x) { return x.id; }), ctx);
  leaves.forEach(function (x) { x.leaveStreak = streaks[x.id] || 0; });
  leaves.sort(function (a, b) {
    return String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant') ||
      (gradeNum_(a.grade) - gradeNum_(b.grade)) || (classNum_(a.grade) - classNum_(b.grade)) ||
      String(a.name).localeCompare(String(b.name), 'zh-Hant');
  });
  return { ok: true, date: date, schools: day.schools.filter(function (s) { return s.submitted; }), leaves: leaves };
}

function actionGetSchoolDayDetail(body) {
  requireSession(body.token, 'admin');
  var attMap = {}, rows = [];
  sheetToObjects(SHEET_ATTENDANCE).forEach(function (a) {
    if (a.schoolId === body.schoolId && String(a.date) === String(body.date)) { attMap[a.studentId] = a; rows.push(a); }
  });
  var date = String(body.date);
  var wd = weekdayOf(date);
  var isToday = date === today_();
  var ctx = statusCtx_();
  var out = activeStudentsOf_(body.schoolId).map(function (st) {
    var a = attMap[st.id];
    var so = studentOut_(st);
    var fixedNow = ctx.isFixed(st.id, date);
    var exb = ctx.extraOn(st.id, date);
    var status = ctx.status(st.id, date);
    var fixed = isToday || exb ? fixedNow : ((a && String(a.wasFixedOff) !== '') ? isTrue_(a.wasFixedOff) : fixedNow);
    var mine = !!ctx.teacherSet(a);
    var bk = !mine && status === 'leave' ? ctx.bookingOn(st.id, date) : null;
    return {
      id: st.id, name: st.name, grade: String(st.grade), status: status,
      wasFixedOff: fixed, contacts: so.contacts, note: so.note, booking: bk ? bookingBrief_(bk) : null, extra: exb ? bookingBrief_(exb) : null,
      markedBy: mine ? a.teacherName : '', markedAt: mine ? String(a.submittedAt) : ''
    };
  });
  var streaks = leaveStreaks_(date, out.map(function (x) { return x.id; }), ctx);
  out.forEach(function (x) { x.leaveStreak = streaks[x.id] || 0; });
  sortStudents(out);
  var submitters = submittersIndex_(String(body.date), rows)[String(body.date) + '|' + body.schoolId] || [];
  return { ok: true, students: out, submitters: submitters };
}

// ---------- 定期清理 (點名紀錄與送出紀錄超過 40 天才清除；順便清掉過期登入) ----------
function pruneByDate_(sheetName, cutoff) {
  var sheet = sheet_(sheetName);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;
  var dateCol = rows[0].indexOf('date');
  var keep = [];
  for (var i = 1; i < rows.length; i++) {
    var d = rows[i][dateCol];
    if (d instanceof Date) d = fromDateCell_(d, 'date');
    if (String(d) >= cutoff) keep.push(plainRow_(rows[0], rows[i]));
  }
  if (keep.length < rows.length - 1) {
    touch_(sheetName);
    sheet.getRange(2, 1, rows.length - 1, rows[0].length).clearContent();
    if (keep.length) sheet.getRange(2, 1, keep.length, rows[0].length).setNumberFormat('@').setValues(keep);
  }
}

function cleanupOldRecords() {
  // 每天凌晨 3 點：順便整理所有分頁格式（設成純文字，避免日期、數字被試算表自動轉換），不影響白天使用速度
  try { ensureSchema_(true); } catch (e) { /* 格式整理失敗不影響清理 */ }
  withLock_(function () {
    var cutoff = dateStrDaysAgo_(RETENTION_DAYS);
    pruneByDate_(SHEET_ATTENDANCE, cutoff);
    pruneByDate_(SHEET_SUBMISSIONS, cutoff);
    pruneByDate_(SHEET_SPECIAL, cutoff);
    // 預約請假：結束日超過保留期限才清除
    var oldB = sheetToObjects(SHEET_BOOKINGS).filter(function (b) { return String(b.endDate) < cutoff; });
    if (oldB.length) deleteRows_(SHEET_BOOKINGS, oldB);
    var oldE = sheetToObjects(SHEET_EXTRAS).filter(function (b) { return String(b.endDate) < cutoff; });
    if (oldE.length) deleteRows_(SHEET_EXTRAS, oldE);
  });
  cleanupExpiredSessions_();
}
