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
var MAKEUP_DAYS = 7;          // 老師可以補點名的天數：昨天往前算幾天內（不含今天，今天走一般點名）
var RETENTION_DAYS = 40;      // 點名紀錄保留天數：超過 40 天的舊資料於每天凌晨 3 點自動清除
var SESSION_HOURS = 12;       // 登入 token 最長有效時數 (前端另外有閒置 60 分鐘自動登出)
var TZ = 'Asia/Taipei';
var SCHEMA_VERSION = 'schema_v11';
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
var SHEET_PICKUP = 'PickupOverrides';
var SHEET_ALLNOON = 'AllNoonDays';

// ---------- 接送時段規則（中午接 / 下午接）----------
// 依年級決定「進班(接)」的學生每週哪幾天中午接，沒列到的平日一律下午接。
// 要改規則只要改這裡（grades＝適用年級，noon＝中午接的星期：1=週一 … 5=週五），改完重新部署即可。
// 學生因社團等原因要調整，在學生資料裡設定；休業式／結業式在「點名日設定」設定為全部中午接。
var PICKUP_RULES = [
  { grades: [1, 2], noon: [1, 3, 4, 5] },   // 一、二年級：週一三四五中午接，週二下午接
  { grades: [3, 4], noon: [3, 5] },         // 三、四年級：週三五中午接，週一二四下午接
  { grades: [5, 6], noon: [3] }             // 五、六年級：週三中午接，週一二四五下午接
];
// 個別學校的特例：用學校名稱對應，列到的年級改用這裡的規則，沒列到的年級仍用上面的預設
var PICKUP_SCHOOL_RULES = {
  '國光國小': [{ grades: [1, 2], noon: [1, 2, 3, 5] }]   // 一、二年級：週一二三五中午接，週四下午接
};

var HEADERS = {
  Admins: ['id', 'username', 'password', 'name'],
  Teachers: ['id', 'username', 'password', 'name', 'active', 'createdAt'],
  Schools: ['id', 'name', 'createdAt', 'active'],
  // fixedOffDays 固定不進班 / selfDropDays 固定自送 / selfWalkDays 固定自走：星期一到五（1–5），同一天只會出現在其中一個
  // pickupOverride 學生每週接送時段調整（例如社團）：格式 "1:pm|3:noon"＝週一改下午接、週三改中午接；沒有填就依年級規則
  Students: ['id', 'schoolId', 'name', 'grade', 'contacts', 'fixedOffDays', 'active', 'createdAt', 'note', 'selfDropDays', 'selfWalkDays', 'pickupOverride'],
  // auto=true：這筆狀態是系統預設值（預約請假 / 固定不進班），老師沒有改過
  Attendance: ['id', 'date', 'schoolId', 'studentId', 'status', 'selfDrop', 'selfWalk', 'teacherId', 'teacherName', 'submittedAt', 'wasFixedOff', 'auto'],
  Sessions: ['token', 'userId', 'role', 'name', 'username', 'createdAt', 'expiresAt'],
  Settings: ['key', 'value'],
  // 每次送出點名的紀錄：同一天同一校可能上午 A 老師、下午 B 老師分批送出
  Submissions: ['id', 'date', 'schoolId', 'teacherId', 'teacherName', 'submittedAt', 'count'],
  // 點名日設定：workday=補班補課（週末可點名）、holiday=國定假日放假（不能點名）
  SpecialDays: ['date', 'type', 'reason', 'createdBy', 'createdAt', 'followDay'],
  // 預約請假：家長事先告知的請假期間（管理員與老師都可以登記）
  LeaveBookings: ['id', 'schoolId', 'studentId', 'startDate', 'endDate', 'reason', 'createdBy', 'createdRole', 'createdAt', 'updatedBy', 'updatedAt'],
  // 臨時進班：固定不進班的學生因臨時或考前加強要進班，期間內固定不進班的日子改為正常進班
  ExtraClasses: ['id', 'schoolId', 'studentId', 'startDate', 'endDate', 'reason', 'createdBy', 'createdRole', 'createdAt', 'updatedBy', 'updatedAt'],
  // 單日例外：某位學生某一天改成中午接（noon）或下午接（pm），優先於年級規則、每週調整與「全部中午接」
  PickupOverrides: ['id', 'schoolId', 'studentId', 'date', 'slot', 'reason', 'createdBy', 'createdAt'],
  // 全部中午接的日子（休業式、結業式）：這天所有年級都中午接，個別例外用單日例外調整
  AllNoonDays: ['date', 'reason', 'createdBy', 'createdAt']
};

// ---------- 初始化 ----------
function setup() {
  ensureSchema_(true);
  var def = ss_().getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss_().deleteSheet(def);
  createDailyCleanupTrigger();
  createLineReminderTrigger();
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

// 學生建立日（台北時區的 yyyy-MM-dd）；舊資料沒有建立時間 → 回傳空字串（視為一直都在）
function createdDateOf_(st) {
  var c = st && st.createdAt;
  if (!c) return '';
  var t = c instanceof Date ? c : new Date(String(c));
  if (isNaN(t.getTime())) return '';
  return Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
}
// 這位學生在 date 那天是否已經在名單內：新增日之後才算；舊日期不補列之後才新增的學生
function studentExistsOn_(st, date) {
  var c = createdDateOf_(st);
  return !c || c <= String(date);
}

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
    if (body && Array.isArray(body.events)) {   // LINE Webhook（綁定通知群組用），不是前端的 API 呼叫
      try { handleLineWebhook_(body); } catch (err2) { /* 不影響回應 */ }
      return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
    }
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
    listMakeup: actionListMakeup, getMakeupRollCall: actionGetMakeupRollCall, saveMakeup: actionSaveMakeup,
    getRollCall: actionGetRollCall, submitAttendance: actionSubmitAttendance, saveMarks: actionSaveMarks, setPresence: actionSetPresence,
    getDayInfo: actionGetDayInfo, listSpecialDays: actionListSpecialDays,
    setSpecialDay: actionSetSpecialDay, deleteSpecialDay: actionDeleteSpecialDay,
    setSpecialDaysBatch: actionSetSpecialDaysBatch, editSpecialDay: actionEditSpecialDay,
    listBookings: actionListBookings, saveBooking: actionSaveBooking, saveBookingsBatch: actionSaveBookingsBatch, deleteBooking: actionDeleteBooking,
    getOverview: actionGetOverview, getSchoolDayDetail: actionGetSchoolDayDetail,
    listLeaveStreaks: actionListLeaveStreaks,
    setPickupSlot: actionSetPickupSlot, setAllNoonDay: actionSetAllNoonDay, deleteAllNoonDay: actionDeleteAllNoonDay
  };
  // 會寫入試算表的動作：一律在全域鎖內執行（多人同時操作也不會互相覆蓋）
  var WRITES = ['login', 'logout', 'changePwd', 'addContactType', 'addSchool', 'addSchoolsBatch', 'editSchool', 'deleteSchool',
    'addStudent', 'addStudentsBatch', 'editStudent', 'deleteStudent', 'deleteStudentsBatch',
    'addTeacher', 'editTeacher', 'deleteTeacher', 'deleteTeachersBatch', 'submitAttendance', 'saveMarks', 'saveMakeup', 'setSpecialDay', 'deleteSpecialDay',
    'setSpecialDaysBatch', 'editSpecialDay', 'saveBooking', 'saveBookingsBatch', 'deleteBooking',
    'setPickupSlot', 'setAllNoonDay', 'deleteAllNoonDay'];
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
  // 順便回傳今天各校的需進班／未點名人數，讓老師的學校卡片隨心跳更新
  return { ok: true, presence: presenceList_(s.userId), counts: s.role === 'teacher' ? todayCounts_() : null };
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

// 老師端學校卡片用：今天各校「需進班人數」與「未點名人數」
// 需進班 = 在學學生 − 今天固定不進班 − 今天請假（預約或老師點的）；未點名 = 需進班裡還沒有狀態的人
// 今天不能點名（例假日／國定假日）時回傳 null
function todayCounts_() {
  var ctx = statusCtx_(), today = ctx.today;
  if (!ctx.isOpen(today)) return null;
  var pk = pickupCtx_();
  var out = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (!isActive_(st.active)) return;
    var c = out[st.schoolId] = out[st.schoolId] || { need: 0, unmarked: 0, noon: 0, pm: 0, noonUn: 0, pmUn: 0, noneUn: 0 };
    var s = ctx.status(st.id, today);
    if (s === 'fixedOff' || s === 'leave') return;
    c.need++;
    if (!s) c.unmarked++;
    var p = pickupSlotFor_(st, today, ctx, pk);   // 需要接的人數：中午接、下午接各幾位（自走、自送的不算）
    if (p) { c[p.slot]++; if (!s) c[p.slot + 'Un']++; }   // 中午接、下午接各自還沒點名的人數
    else if (!s) c.noneUn++;   // 今天固定自走／自送、不用接但還沒點名的人（算在未點名，但不屬於中午或下午）
  });
  return out;
}

function actionListSchools(body) {
  requireSession(body.token);
  var counts = {};
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (isActive_(st.active)) counts[st.schoolId] = (counts[st.schoolId] || 0) + 1;
  });
  var today = todayCounts_();
  var out = activeSchools_().map(function (sc) {
    var t = today ? (today[sc.id] || { need: 0, unmarked: 0, noon: 0, pm: 0, noonUn: 0, pmUn: 0, noneUn: 0 }) : null;
    return {
      id: sc.id, name: sc.name, studentCount: counts[sc.id] || 0, todayNeed: t ? t.need : null, todayUnmarked: t ? t.unmarked : null,
      todayNoon: t ? t.noon : null, todayPm: t ? t.pm : null, todayNoonUn: t ? t.noonUn : null, todayPmUn: t ? t.pmUn : null, todayNoneUn: t ? t.noneUn : null
    };
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
  var m = weekModes_({ off: parseDays_(x.fixedOffDays), drop: parseDays_(x.selfDropDays), walk: parseDays_(x.selfWalkDays) });
  return {
    id: x.id, name: x.name, grade: String(x.grade), gradeNum: gradeNum_(x.grade),
    contacts: parseContacts(x), fixedOffDays: m.off, selfDropDays: m.drop, selfWalkDays: m.walk, note: String(x.note || ''),
    pickupOverride: parsePickupOverride_(x.pickupOverride)   // 每週接送時段調整：{ 星期: 'noon'|'pm' }，沒有的星期依年級規則
  };
}

// 每週固定設定：星期一到五每天只會有一種（不進班 / 自送 / 自走 / 預設進班需要接）。
// 若資料出現同一天重複設定，依「固定不進班 > 自送 > 自走」保留一個
function weekModes_(raw) {
  var off = normalizeDays_(raw.off);
  var drop = normalizeDays_(raw.drop).filter(function (d) { return off.indexOf(d) === -1; });
  var walk = normalizeDays_(raw.walk).filter(function (d) { return off.indexOf(d) === -1 && drop.indexOf(d) === -1; });
  return { off: off, drop: drop, walk: walk };
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
  // 單日例外（今天起）：讓管理員在學生視窗看到、可取消
  var today = today_(), byStudent = {};
  sheetToObjects(SHEET_PICKUP).forEach(function (o) {
    if (String(o.date) < today) return;
    (byStudent[o.studentId] = byStudent[o.studentId] || []).push({ date: String(o.date), slot: String(o.slot), reason: String(o.reason || '') });
  });
  students.forEach(function (st) {
    st.pickupDates = (byStudent[st.id] || []).sort(function (a, b) { return a.date.localeCompare(b.date); });
  });
  return { ok: true, students: students, pickupRules: pickupRulesOfSchool_(body.schoolId) };
}

function studentRow_(schoolId, st) {
  var m = weekModes_({ off: st.fixedOffDays, drop: st.selfDropDays, walk: st.selfWalkDays });
  return {
    id: genId('T'), schoolId: schoolId, name: String(st.name || '').trim(), grade: String(st.grade || '').trim(),
    contacts: JSON.stringify(normalizeContacts(st.contacts)),
    fixedOffDays: m.off.join('|'), selfDropDays: m.drop.join('|'), selfWalkDays: m.walk.join('|'),
    pickupOverride: formatPickupOverride_(normalizePickupOverride_(st.pickupOverride, st.grade, pickupRulesOfSchool_(schoolId))),
    active: true, createdAt: new Date().toISOString(),
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
      if (body.fixedOffDays !== undefined || body.selfDropDays !== undefined || body.selfWalkDays !== undefined) {
        var cur = studentOut_(u);   // 沒傳的欄位沿用原本的設定
        var m = weekModes_({
          off: body.fixedOffDays !== undefined ? body.fixedOffDays : cur.fixedOffDays,
          drop: body.selfDropDays !== undefined ? body.selfDropDays : cur.selfDropDays,
          walk: body.selfWalkDays !== undefined ? body.selfWalkDays : cur.selfWalkDays
        });
        u.fixedOffDays = m.off.join('|'); u.selfDropDays = m.drop.join('|'); u.selfWalkDays = m.walk.join('|');
      }
      if (body.pickupOverride !== undefined) {
        u.pickupOverride = formatPickupOverride_(normalizePickupOverride_(body.pickupOverride, grade, pickupRulesOfSchool_(u.schoolId)));
      } else if (u.pickupOverride) {   // 只改了年級：把變成和新年級規則一樣的調整清掉
        u.pickupOverride = formatPickupOverride_(normalizePickupOverride_(u.pickupOverride, grade, pickupRulesOfSchool_(u.schoolId)));
      }
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

// ---------- 接送時段：中午接 / 下午接 ----------
// 某位學生某一天的時段，依優先順序決定：
// 1. 單日例外（PickupOverrides）  2. 全部中午接日（休業式／結業式）  3. 學生每週調整（社團）  4. 年級規則（PICKUP_RULES）
// 只有「進班(接)」的日子才有時段；自走、自送、固定不進班的日子不用接，時段留空。
// 某間學校實際適用的規則：學校特例放前面，其餘用預設
function pickupRulesOfSchool_(schoolId) {
  var sc = sheetToObjects(SHEET_SCHOOLS).filter(function (x) { return x.id === schoolId; })[0];
  var extra = (sc && PICKUP_SCHOOL_RULES[String(sc.name).trim()]) || [];
  return extra.concat(PICKUP_RULES);
}
function pickupBase_(grade, wd, rules) {
  rules = rules || PICKUP_RULES;
  if (wd < 1 || wd > 5) return 'noon';   // 補班補課的週末沒有年級規則，一律中午接
  var g = gradeNum_(grade), rule = rules[rules.length - 1];
  if (g < 1) rule = PICKUP_RULES[0];
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].grades.indexOf(g) !== -1) { rule = rules[i]; break; }
  }
  return rule.noon.indexOf(wd) !== -1 ? 'noon' : 'pm';
}

// "1:pm|3:noon" → { 1: 'pm', 3: 'noon' }
function parsePickupOverride_(v) {
  var out = {};
  String(v === undefined || v === null ? '' : v).split('|').forEach(function (p) {
    var m = p.match(/^([1-5]):(noon|pm)$/);
    if (m) out[m[1]] = m[2];
  });
  return out;
}
function formatPickupOverride_(o) {
  return Object.keys(o).sort().map(function (k) { return k + ':' + o[k]; }).join('|');
}
// 只留下「和年級規則不同」的星期，存起來的就一定是真正的調整
function normalizePickupOverride_(raw, grade, rules) {
  var src = (raw && typeof raw === 'object') ? raw : parsePickupOverride_(raw);
  var out = {};
  for (var d = 1; d <= 5; d++) {
    var v = src[d];
    if ((v === 'noon' || v === 'pm') && v !== pickupBase_(grade, d, rules)) out[d] = v;
  }
  return out;
}

function pickupCtx_() {
  var ov = {}, noon = {};
  sheetToObjects(SHEET_PICKUP).forEach(function (o) { ov[o.studentId + '|' + String(o.date)] = o; });
  sheetToObjects(SHEET_ALLNOON).forEach(function (o) { noon[String(o.date)] = o; });
  var follow = {};   // 補班補課日「補星期幾」→ 當天各年級照那天的接送時段
  specialDays_().forEach(function (x) { if (x.type === 'workday' && x.followDay) follow[x.date] = x.followDay; });
  var rc = {};
  return { overrides: ov, allNoon: noon, follow: follow, rulesOf: function (sid) { return rc[sid] || (rc[sid] = pickupRulesOfSchool_(sid)); } };
}

// 回傳 { slot: 'noon'|'pm', src: 'date'|'allNoon'|'week'|'grade', reason }
function pickupInfo_(st, date, pk) {
  var o = pk.overrides[st.id + '|' + date];
  if (o && (o.slot === 'noon' || o.slot === 'pm')) return { slot: o.slot, src: 'date', reason: String(o.reason || '') };
  if (pk.allNoon[date]) return { slot: 'noon', src: 'allNoon', reason: String(pk.allNoon[date].reason || '') };
  var wd = (pk.follow && pk.follow[date]) || weekdayOf(date), base = pickupBase_(st.grade, wd, pk.rulesOf(st.schoolId));
  var mine = wd >= 1 && wd <= 5 ? parsePickupOverride_(st.pickupOverride)[wd] : '';
  if (mine && mine !== base) return { slot: mine, src: 'week', reason: '' };
  return { slot: base, src: 'grade', reason: '' };
}

// 這天需要「接」這位學生才有時段；每週設定為自走／自送／固定不進班（沒有臨時進班）→ 回傳 null
function pickupSlotFor_(st, date, ctx, pk) {
  var wd = ctx.wdOf(date);
  if (wd >= 1 && wd <= 5) {
    if (parseDays_(st.selfDropDays).indexOf(wd) !== -1 || parseDays_(st.selfWalkDays).indexOf(wd) !== -1) return null;
    if (ctx.isFixed(st.id, date)) return null;
  }
  return pickupInfo_(st, date, pk);
}

// 單日例外：某位學生某一天改成中午接或下午接（slot 留空＝取消例外）。管理員可設今天起的任何日期；老師只能調整今天
function actionSetPickupSlot(body) {
  var s = requireSession(body.token);
  var today = today_();
  var date = body.date ? String(body.date) : today;
  if (!isDateStr_(date)) return { ok: false, error: 'DATE_REQUIRED' };
  if (date < today) return { ok: false, error: 'DATE_PAST' };
  if (s.role !== 'admin' && date !== today) return { ok: false, error: 'FORBIDDEN' };
  var slot = String(body.slot || '');
  if (slot !== '' && slot !== 'noon' && slot !== 'pm') return { ok: false, error: 'UNKNOWN_ACTION' };
  var st = sheetToObjects(SHEET_STUDENTS).filter(function (x) { return x.id === body.studentId && isActive_(x.active); })[0];
  if (!st) return { ok: false, error: 'STUDENT_NOT_FOUND' };
  deleteRows_(SHEET_PICKUP, sheetToObjects(SHEET_PICKUP).filter(function (o) { return o.studentId === st.id && String(o.date) === date; }));
  var def = pickupInfo_(st, date, pickupCtx_());   // 沒有單日例外時的時段
  if (slot && slot !== def.slot) {
    appendObject(SHEET_PICKUP, {
      id: genId('P'), schoolId: st.schoolId, studentId: st.id, date: date, slot: slot,
      reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString()
    });
    return { ok: true, date: date, slot: slot, src: 'date' };
  }
  return { ok: true, date: date, slot: def.slot, src: def.src };
}

// 全部中午接的日子（休業式、結業式）
function allNoonDays_() {
  return sheetToObjects(SHEET_ALLNOON).map(function (x) {
    return { date: String(x.date), reason: String(x.reason || ''), weekday: weekdayOf(String(x.date)) };
  }).sort(function (a, b) { return b.date.localeCompare(a.date); });
}
function actionSetAllNoonDay(body) {
  var s = requireSession(body.token, 'admin');
  var date = String(body.date || '');
  if (!isDateStr_(date)) return { ok: false, error: 'DATE_REQUIRED' };
  deleteRows_(SHEET_ALLNOON, sheetToObjects(SHEET_ALLNOON).filter(function (x) { return String(x.date) === date; }));
  appendObject(SHEET_ALLNOON, { date: date, reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString() });
  return { ok: true };
}
function actionDeleteAllNoonDay(body) {
  requireSession(body.token, 'admin');
  var rows = sheetToObjects(SHEET_ALLNOON).filter(function (x) { return String(x.date) === String(body.date); });
  if (!rows.length) return { ok: false, error: 'DAY_NOT_FOUND' };
  deleteRows_(SHEET_ALLNOON, rows);
  return { ok: true };
}

// ---------- 點名日：例假日 / 補班補課 / 國定假日 ----------
var WEEKDAY_NAMES = ['', '一', '二', '三', '四', '五', '六', '日'];

// 補班補課「補星期幾」：1～5 才有效，其餘（空白、週末）＝0＝不指定（一律中午接）
function followDay_(v) {
  var n = Number(v);
  return (n >= 1 && n <= 5 && n === Math.floor(n)) ? n : 0;
}
function specialDays_() {
  return sheetToObjects(SHEET_SPECIAL).map(function (x) {
    return { date: String(x.date), type: String(x.type), reason: String(x.reason || ''), createdBy: x.createdBy, createdAt: String(x.createdAt || ''),
      followDay: String(x.type) === 'workday' ? followDay_(x.followDay) : 0 };
  });
}

// 某一天能不能點名：補班補課 > 國定假日 > 週六日例假日 > 平日
function dayInfo_(date, list) {
  var sp = (list || specialDays_()).filter(function (x) { return x.date === date; })[0];
  var wd = weekdayOf(date);
  var base = { date: date, weekday: wd, reason: sp ? sp.reason : '' };
  if (sp && sp.type === 'workday') return Object.assign(base, { open: true, type: 'workday', label: '補班補課', followDay: sp.followDay || 0 });
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
  return { ok: true, today: today_(), todayInfo: dayInfo_(today_(), list), days: list, allNoonDays: allNoonDays_() };
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
  appendObject(SHEET_SPECIAL, { date: date, type: type, reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString(),
    followDay: type === 'workday' ? (followDay_(body.followDay) || '') : '' });
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
    map[date] = { type: type, reason: String(d.reason || '').trim().slice(0, 30), followDay: type === 'workday' ? (followDay_(d.followDay) || '') : '' };   // 重複日期：以最後一筆為準
  });
  if (bad.length) return { ok: false, error: 'INVALID_DAYS', bad: bad };
  if (!order.length) return { ok: false, error: 'DATE_REQUIRED' };
  var existing = sheetToObjects(SHEET_SPECIAL).filter(function (x) { return map[String(x.date)]; });
  var updated = {};
  existing.forEach(function (x) { updated[String(x.date)] = true; });
  deleteRows_(SHEET_SPECIAL, existing);
  var now = new Date().toISOString();
  order.forEach(function (date) {
    appendObject(SHEET_SPECIAL, { date: date, type: map[date].type, reason: map[date].reason, createdBy: s.name, createdAt: now, followDay: map[date].followDay });
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
  var newType = autoDayType_(newDate);
  var keep = body.followDay === undefined ? followDay_(rows[0].followDay) : followDay_(body.followDay);   // 沒帶就沿用原本的
  deleteRows_(SHEET_SPECIAL, rows);
  appendObject(SHEET_SPECIAL, { date: newDate, type: newType, reason: String(body.reason || '').trim().slice(0, 30), createdBy: s.name, createdAt: new Date().toISOString(),
    followDay: newType === 'workday' ? (keep || '') : '' });
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
  else {
    appendObject(bkSheet_(kind), row);
    notifyBookingsLine_(s, kind, [{ row: row, days: c.days, st: c.st }]);   // 新增才通知老師群組（修改不通知）
  }
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
  var bad = [], rows = [], totalDays = 0, notes = [];
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
    notes.push({ row: row, days: c.days, st: c.st });
    all.push(row);   // 同一批裡同一位學生的期間也不能重疊
    totalDays += c.days;
  });
  if (bad.length) return { ok: false, error: 'INVALID_BOOKINGS', bad: bad };
  appendObjects(bkSheet_(kind), rows);
  notifyBookingsLine_(s, kind, notes);
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
    var mk = isMakeupAt_(key.split('|')[0], at);   // 隔天（含）以後才送出的點名＝補點名，和當天點的分開列
    var k = (tid || name) + (mk ? '|mk' : '');
    if (!m[k] || at > m[k].at) m[k] = { name: name, at: at, makeup: mk };
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

// ---------- 補點名 ----------
// 送出時間的台北日期 > 點名日期 ＝ 補點名（不用多存欄位，舊資料也自然適用）
function dayOfIso_(iso) { var t = Date.parse(iso); return isNaN(t) ? '' : Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd'); }
function isMakeupAt_(date, iso) { var d = dayOfIso_(iso); return !!d && d > String(date); }

// 補點名的日期要在「昨天往前 MAKEUP_DAYS 天」內、而且當天可以點名；回傳錯誤代碼或空字串
function checkMakeupDate_(date) {
  date = String(date || '');
  if (!isDateStr_(date)) return 'DATE_REQUIRED';
  if (date >= today_() || date < dateStrDaysAgo_(MAKEUP_DAYS)) return 'MAKEUP_RANGE';
  if (!dayInfo_(date).open) return 'DAY_CLOSED';
  return '';
}

// 過去的日子各校還沒點名的人數：{ 日期: { 學校id: 人數 } }。
// 沒有狀態、也不是固定不進班的學生才算；從有點名紀錄的第一天算起（系統上線前不算），今天不算（還在點）
function missingByDay_(fromDate, toDate, ctx) {
  var today = today_();
  if (toDate >= today) toDate = prevDate_(today);
  var first = '';
  Object.keys(ctx.rec).forEach(function (k) { var d = k.slice(0, 10); if (!first || d < first) first = d; });
  if (!first) return {};
  if (fromDate < first) fromDate = first;
  var activeIds = {};
  activeSchools_().forEach(function (sc) { activeIds[sc.id] = true; });
  var students = sheetToObjects(SHEET_STUDENTS).filter(function (st) { return isActive_(st.active) && activeIds[st.schoolId]; });
  var out = {};
  for (var d = fromDate; d <= toDate; d = nextDate_(d)) {
    if (!ctx.isOpen(d)) continue;
    students.forEach(function (st) {
      if (!studentExistsOn_(st, d) || ctx.status(st.id, d) || ctx.isFixed(st.id, d)) return;
      var m = out[d] = out[d] || {};
      m[st.schoolId] = (m[st.schoolId] || 0) + 1;
    });
  }
  return out;
}
function missingList_(map, schoolMap, schoolFilter) {
  return Object.keys(map || {}).filter(function (sid) { return !schoolFilter || sid === schoolFilter; }).map(function (sid) {
    return { schoolId: sid, schoolName: schoolMap[sid] || '', n: map[sid] };
  }).sort(function (a, b) { return String(a.schoolName).localeCompare(String(b.schoolName), 'zh-Hant'); });
}

// 老師端：哪幾天、哪幾間學校還有學生沒點名（可以補點名）
function actionListMakeup(body) {
  requireSession(body.token);
  var ctx = statusCtx_();
  var schoolMap = {};
  activeSchools_().forEach(function (sc) { schoolMap[sc.id] = sc.name; });
  var miss = missingByDay_(dateStrDaysAgo_(MAKEUP_DAYS), prevDate_(today_()), ctx);
  var items = [];
  Object.keys(miss).sort(function (a, b) { return b.localeCompare(a); }).forEach(function (d) {
    missingList_(miss[d], schoolMap, null).forEach(function (x) { items.push({ date: d, schoolId: x.schoolId, schoolName: x.schoolName, unmarked: x.n }); });
  });
  return { ok: true, days: MAKEUP_DAYS, items: items };
}

// 老師端：某天某校的補點名名單。固定不進班（沒有紀錄）的學生不用點，預約請假（系統判斷）的學生只讀
function actionGetMakeupRollCall(body) {
  requireSession(body.token, 'teacher');
  var date = String(body.date || '');
  var err = checkMakeupDate_(date);
  if (err) return { ok: false, error: err };
  var ctx = statusCtx_();
  var wd = weekdayOf(date);
  var recs = {};
  sheetToObjects(SHEET_ATTENDANCE).forEach(function (a) { if (a.schoolId === body.schoolId && String(a.date) === date) recs[a.studentId] = a; });
  var out = [];
  activeStudentsOf_(body.schoolId).forEach(function (st) {
    if (!studentExistsOn_(st, date) && !recs[st.id]) return;
    var status = ctx.status(st.id, date);
    if (status === 'fixedOff' || (!status && ctx.isFixed(st.id, date))) return;   // 固定不進班：不用點
    var so = studentOut_(st), a = recs[st.id], mine = !!ctx.teacherSet(a);
    out.push({
      id: so.id, name: so.name, grade: so.grade, status: status, locked: status === 'leave' && !mine,
      fixedMode: so.selfDropDays.indexOf(wd) !== -1 ? 'selfDrop' : (so.selfWalkDays.indexOf(wd) !== -1 ? 'selfWalk' : ''),
      markedBy: mine ? a.teacherName : '', markedAt: mine ? String(a.submittedAt) : '', makeup: mine && isMakeupAt_(date, a.submittedAt)
    });
  });
  sortStudents(out);
  return { ok: true, date: date, students: out };
}

// 老師端：補點名儲存（每點一位就送一次）。沿用一般點名的寫入流程，紀錄裡的送出時間晚於點名日＝補點名
function actionSaveMakeup(body) {
  var err = checkMakeupDate_(body.date);
  if (err) { requireSession(body.token, 'teacher'); return { ok: false, error: err }; }
  var records = (body.records || []).filter(function (r) { return r && PICK_STATUS_.indexOf(r.status) !== -1; });
  return actionSubmitAttendance(Object.assign({}, body, { records: records, autosave: true, makeupDate: String(body.date) }));
}
var PICK_STATUS_ = ['present', 'selfDrop', 'selfWalk', 'leave'];

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
  var pk = pickupCtx_();
  var out = activeStudentsOf_(schoolId).map(function (st) {
    var so = studentOut_(st);
    var isFixedOff = ctx.isFixed(st.id, date);   // 臨時進班的日子不算固定不進班
    var exb = ctx.extraOn(st.id, date);
    var existing = attMap[st.id];
    var primary = so.contacts.filter(function (c) { return c.primary; })[0] || null;
    var status = ctx.status(st.id, date);
    var byTeacher = !!ctx.teacherSet(existing);   // 系統自動判斷的不顯示點名老師
    var bk = ctx.bookingOn(st.id, date);
    var pick = pickupSlotFor_(st, date, ctx, pk);   // 今天中午接／下午接；自走、自送、固定不進班的人不用接 → 空白
    return {
      id: so.id, name: so.name, grade: so.grade, isFixedOff: isFixedOff, primaryContact: primary,
      pickup: pick ? pick.slot : '', pickupSrc: pick ? pick.src : '', pickupReason: pick ? pick.reason : '',
      contacts: so.contacts, note: so.note, status: status, defaultStatus: ctx.defaultOf(st.id, date), fixedOffDays: so.fixedOffDays,
      // 今天是這位學生固定的自送／自走日：只當提示，不自動預填狀態（要老師確認到班後再點）
      fixedMode: so.selfDropDays.indexOf(wd) !== -1 ? 'selfDrop' : (so.selfWalkDays.indexOf(wd) !== -1 ? 'selfWalk' : ''),
      booking: bk ? bookingBrief_(bk) : null, extra: exb ? bookingBrief_(exb) : null,
      markedBy: byTeacher ? existing.teacherName : '', markedAt: byTeacher ? String(existing.submittedAt) : ''
    };
  });
  sortStudents(out);

  var submitters = submittersIndex_(date, attendance)[date + '|' + schoolId] || [];
  var last = submitters[submitters.length - 1] || null;
  return {
    ok: true, date: date, day: dayInfo_(date), students: out, submitted: submitters.length > 0, submitters: submitters,
    allNoon: pk.allNoon[date] ? { reason: String(pk.allNoon[date].reason || '') } : null,   // 今天是休業式／結業式：全部中午接
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
  if (body.makeupDate) {   // 補點名：補過去幾天（由 actionSaveMakeup 帶入，這裡再檢查一次）
    var mkErr = checkMakeupDate_(body.makeupDate);
    if (mkErr) return { ok: false, error: mkErr };
    date = String(body.makeupDate);
  }
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
    // 這次真的有存到學生 → 順便把該校「今天固定不進班」、還沒有紀錄（或只有空白）的學生補寫成固定不進班（系統預設值），
    // 這樣隔天回頭看歷史時不會變成未點名；不算進送出筆數，也不會被當成老師點的
    if (n) {
      students.forEach(function (st) {
        if (!fixedToday[st.id] || latestOf[st.id]) return;   // 這次有送出的學生已經在上面處理過
        var ri = rowOf[st.id];
        var fx = {
          date: date, schoolId: schoolId, studentId: st.id, status: 'fixedOff', selfDrop: false, selfWalk: false,
          teacherId: s.userId, teacherName: s.name, submittedAt: now, wasFixedOff: true, auto: true
        };
        if (ri === undefined) { fx.id = genId('R'); fresh.push(fx); return; }
        var frow = values[ri], fold = {};
        headers.forEach(function (h, c) { fold[h] = frow[c]; });
        if (effectiveStatus_(fold)) return;   // 已經有狀態就不動
        fx.id = frow[idx.id];
        headers.forEach(function (h, c) { frow[c] = toCell_(fx[h]); });
        changedRows.push(ri);
      });
    }
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
  var totals = {}, activeStudents = [];
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (!isActive_(st.active)) return;
    totals[st.schoolId] = (totals[st.schoolId] || 0) + 1;
    activeStudents.push(st);
  });
  // 某天某校的「總人數」：過去的日子不算之後才新增的學生（今天一律用目前人數）
  var totalOn = function (sid, d) {
    if (d === today) return totals[sid] || 0;
    var n = 0;
    activeStudents.forEach(function (st) { if (st.schoolId === sid && studentExistsOn_(st, d)) n++; });
    return n;
  };
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
    // 「已送出」＝老師真的點了至少一位學生的狀態。系統補寫的固定不進班、或老師按「完成點名」時一起寫入的空白列（全部還是未點名）都不算
    if (String(a.teacherId) !== 'system' && ctx.teacherSet(a)) hasRec[d + '|' + a.schoolId] = true;
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
      var total = totalOn(sid, date);
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
  // 今天各校的中午接／下午接人數（今天不能點名時不帶）
  var tc = todayCounts_();
  if (tc) d.dates.forEach(function (day) {
    if (day.date !== d.today) return;
    day.schools.forEach(function (sc) {
      var c = tc[sc.schoolId] || { noon: 0, pm: 0, noonUn: 0, pmUn: 0, noneUn: 0 };
      sc.noon = c.noon; sc.pm = c.pm; sc.noonUn = c.noonUn; sc.pmUn = c.pmUn; sc.noneUn = c.noneUn;
    });
  });
  // 過去每天各校還沒點名的人數（含整天都沒人點的學校）→ 管理端標示「有老師未點名」
  var days = Number(body.days) || RETENTION_DAYS;
  var miss = missingByDay_(dateStrDaysAgo_(days), prevDate_(d.today), d.ctx), missing = {};
  Object.keys(miss).forEach(function (dt) { missing[dt] = missingList_(miss[dt], d.schoolMap, body.schoolId || null); });
  return { ok: true, today: d.today, todayInfo: dayInfo_(d.today, sp), special: special, dates: d.dates, missing: missing, presence: presenceList_(null) };
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
  var missing = [];
  if (date < d.today) {
    var mm = missingByDay_(date, date, ctx);
    missing = missingList_(mm[date], d.schoolMap, schoolFilter);
  }
  return { ok: true, date: date, schools: day.schools.filter(function (s) { return s.submitted; }), leaves: leaves, missing: missing };
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
  var out = activeStudentsOf_(body.schoolId).filter(function (st) {
    // 過去的日子不列出「之後才新增」的學生（當天已有點名紀錄的一律保留）
    return isToday || attMap[st.id] || studentExistsOn_(st, date);
  }).map(function (st) {
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
      markedBy: mine ? a.teacherName : '', markedAt: mine ? String(a.submittedAt) : '', makeup: mine && isMakeupAt_(date, a.submittedAt)
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

// 把 fromDate～toDate（含）每個「可以點名」的日子，補寫還沒有紀錄（或只有空白）的固定不進班學生為「固定不進班」。
// 系統補寫的列 teacherId 是 'system'、老師姓名留空：不算老師點的，也不會讓該校顯示成「已送出」。
// 只補過去的日子（今天仍然即時依設定計算）；依「執行當下」學生的固定不進班設定與新增日期判斷。
function backfillFixedOff_(fromDate, toDate) {
  var today = today_();
  if (toDate >= today) toDate = prevDate_(today);
  if (fromDate > toDate) return 0;
  return withLock_(function () {
    var ctx = statusCtx_();
    var now = new Date().toISOString();
    var students = sheetToObjects(SHEET_STUDENTS).filter(function (st) { return isActive_(st.active); });
    var fresh = [], n = 0;
    for (var d = fromDate; d <= toDate; d = nextDate_(d)) {
      if (!ctx.isOpen(d)) continue;
      students.forEach(function (st) {
        if (!studentExistsOn_(st, d) || !ctx.isFixed(st.id, d)) return;
        var cur = ctx.rec[d + '|' + st.id];
        if (cur && effectiveStatus_(cur)) return;   // 已經有狀態就不動
        var fx = {
          date: d, schoolId: st.schoolId, studentId: st.id, status: 'fixedOff', selfDrop: false, selfWalk: false,
          teacherId: 'system', teacherName: '', submittedAt: now, wasFixedOff: true, auto: true
        };
        if (cur) {
          var upd = Object.assign({}, cur, fx, { id: cur.id });
          delete upd._row;
          updateObjectByRow(SHEET_ATTENDANCE, cur._row, upd);
        } else {
          fx.id = genId('R');
          fresh.push(fx);
        }
        n++;
      });
    }
    if (fresh.length) appendObjects(SHEET_ATTENDANCE, fresh);
    return n;
  });
}
// 一次性：補回保留期限內（40 天）到昨天所有缺少的固定不進班紀錄。在 Apps Script 編輯器手動執行一次即可
function backfillFixedOffHistory() {
  var n = backfillFixedOff_(dateStrDaysAgo_(RETENTION_DAYS), prevDate_(today_()));
  return 'backfilled ' + n;
}

function cleanupOldRecords() {
  // 每天凌晨 3 點：順便整理所有分頁格式（設成純文字，避免日期、數字被試算表自動轉換），不影響白天使用速度
  try { ensureSchema_(true); } catch (e) { /* 格式整理失敗不影響清理 */ }
  // 先補寫昨天固定不進班的學生（沒人點名的學校也不會漏），再清舊資料
  try { backfillFixedOff_(prevDate_(today_()), prevDate_(today_())); } catch (e) { /* 補寫失敗不影響清理 */ }
  withLock_(function () {
    var cutoff = dateStrDaysAgo_(RETENTION_DAYS);
    pruneByDate_(SHEET_ATTENDANCE, cutoff);
    pruneByDate_(SHEET_SUBMISSIONS, cutoff);
    pruneByDate_(SHEET_SPECIAL, cutoff);
    pruneByDate_(SHEET_PICKUP, cutoff);
    pruneByDate_(SHEET_ALLNOON, cutoff);
    // 預約請假：結束日超過保留期限才清除
    var oldB = sheetToObjects(SHEET_BOOKINGS).filter(function (b) { return String(b.endDate) < cutoff; });
    if (oldB.length) deleteRows_(SHEET_BOOKINGS, oldB);
    var oldE = sheetToObjects(SHEET_EXTRAS).filter(function (b) { return String(b.endDate) < cutoff; });
    if (oldE.length) deleteRows_(SHEET_EXTRAS, oldE);
  });
  cleanupExpiredSessions_();
}

// =====================================================================
// ---------- LINE 通知（只發給老師群組，不發給家長）----------
// 1) 新增「預約請假」「臨時進班」時，立刻推播
// 2) 每天 13:00 推播「今天中午接、還沒點名」的學生；16:30 推播「今天下午接、還沒點名」的學生
//
// 設定：Apps Script「專案設定 > 指令碼屬性」新增兩個屬性（金鑰不要寫在程式裡，也不會被推到 GitHub）
//   LINE_CHANNEL_TOKEN  Messaging API 的 Channel access token (long-lived)
//   LINE_GROUP_ID       老師群組的 groupId（C 開頭）；可用逗號放多個收件對象
// 沒有 groupId 時：把官方帳號 @098stubd 邀進老師群組，Webhook URL 填這個 Web App 的 /exec 網址，
//   然後在群組裡傳一句「綁定點名通知」，系統會自動把這個群組存成 LINE_GROUP_ID 並回覆確認。
// 第一次使用：在編輯器執行一次 setupLine()（授權 + 建立排程 + 傳測試訊息）。
// =====================================================================
var LINE_NOTIFY_ADMIN_BOOKINGS = true;   // 管理員新增的預約請假／臨時進班也通知老師群組；不要的話改成 false
var LINE_BIND_KEYWORD = '綁定點名通知';
var LINE_REMINDER_GRACE_MIN = 30;        // 排程每 5 分鐘檢查一次；整點後 30 分鐘內補發都算，過了就不發
var LINE_REMINDERS = [
  { slot: 'noon', label: '中午接', hour: 13, minute: 0 },
  { slot: 'pm', label: '下午接', hour: 16, minute: 30 }
];

function lineProps_() { return PropertiesService.getScriptProperties(); }
function lineTargets_() {
  return String(lineProps_().getProperty('LINE_GROUP_ID') || '').split(',').map(function (x) { return x.trim(); }).filter(String);
}
function lineConfigured_() { return !!lineProps_().getProperty('LINE_CHANNEL_TOKEN') && lineTargets_().length > 0; }

// 推播文字（超過 LINE 單則上限時自動拆成多則）。回傳 { ok, error? }；失敗只記錄，不丟錯誤，避免影響點名
function linePush_(text) {
  var p = lineProps_(), token = p.getProperty('LINE_CHANNEL_TOKEN'), to = lineTargets_();
  if (!token || !to.length) return { ok: false, error: 'LINE_NOT_CONFIGURED' };
  var chunks = [], cur = '';
  String(text).split('\n').forEach(function (line) {
    if ((cur + '\n' + line).length > 4500 && cur) { chunks.push(cur); cur = line; }
    else cur = cur ? cur + '\n' + line : line;
  });
  if (cur) chunks.push(cur);
  var err = '';
  to.forEach(function (target) {
    chunks.forEach(function (chunk) {
      try {
        var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { Authorization: 'Bearer ' + token },
          payload: JSON.stringify({ to: target, messages: [{ type: 'text', text: chunk }] })
        });
        if (res.getResponseCode() !== 200) err = res.getResponseCode() + ' ' + res.getContentText().slice(0, 300);
      } catch (e) { err = String(e && e.message || e); }
    });
  });
  p.setProperty('LINE_LAST_ERROR', err ? new Date().toISOString() + ' ' + err : '');
  return err ? { ok: false, error: err } : { ok: true };
}

function lineDateLabel_(d) {
  var w = ['', '一', '二', '三', '四', '五', '六', '日'][weekdayOf(d)];
  return parseInt(d.slice(5, 7), 10) + '/' + parseInt(d.slice(8, 10), 10) + '(' + w + ')';
}

// 新增預約請假／臨時進班 → 通知老師群組；entries: [{ row, days, st }]
function notifyBookingsLine_(s, kind, entries) {
  try {
    if (!entries || !entries.length || !lineConfigured_()) return;
    if (s.role === 'admin' && !LINE_NOTIFY_ADMIN_BOOKINGS) return;
    var schoolName = {};
    sheetToObjects(SHEET_SCHOOLS).forEach(function (sc) { schoolName[sc.id] = sc.name; });
    var bySchool = {}, order = [];
    entries.forEach(function (e) {
      var sid = e.row.schoolId;
      if (!bySchool[sid]) { bySchool[sid] = []; order.push(sid); }
      bySchool[sid].push(e);
    });
    var lines = [kind === 'attend' ? '➕ 新增臨時進班' : '📝 新增預約請假',
      '登記人：' + s.name + (s.role === 'admin' ? '（管理員）' : '（老師）')];
    order.forEach(function (sid) {
      lines.push('');
      lines.push('【' + (schoolName[sid] || '學校') + '】');
      var sorted = sortStudents(bySchool[sid].map(function (e) { return { grade: e.st.grade, name: e.st.name, e: e }; }));
      sorted.forEach(function (x) {
        var e = x.e, r = e.row, st = e.st;
        var span = r.startDate === r.endDate ? lineDateLabel_(r.startDate) : lineDateLabel_(r.startDate) + '～' + lineDateLabel_(r.endDate);
        lines.push('• ' + st.name + '（' + st.grade + '）' + span + '，共 ' + e.days + ' 天' + (r.reason ? '｜' + r.reason : ''));
      });
    });
    linePush_(lines.join('\n'));
  } catch (err) { /* 通知失敗不影響登記 */ }
}

// 今天「還沒點名、而且要接」的學生，依學校與接送時段分組；今天不能點名（例假日／國定假日）回傳 null
// 排除：已點名（到班、自送、自走、請假）、預約請假、今天固定不進班、固定自走／自送（不用接，不屬於中午／下午）
function unmarkedBySlot_() {
  var ctx = statusCtx_(), today = ctx.today;
  if (!ctx.isOpen(today)) return null;
  var pk = pickupCtx_();
  var schools = {};
  activeSchools_().forEach(function (sc) { schools[sc.id] = { name: sc.name, noon: [], pm: [], walk: [] }; });
  var wd = ctx.wdOf(today);
  sheetToObjects(SHEET_STUDENTS).forEach(function (st) {
    if (!isActive_(st.active) || !schools[st.schoolId]) return;
    if (ctx.status(st.id, today)) return;
    var p = pickupSlotFor_(st, today, ctx, pk);
    if (!p) {   // 今天固定自走／自送、不用接但老師還沒點：下午接的提醒一起列出（標註自走或自送）
      var tag = parseDays_(st.selfDropDays).indexOf(wd) !== -1 ? '自送' : '自走';
      schools[st.schoolId].walk.push(Object.assign({}, st, { tag: tag }));
      return;
    }
    schools[st.schoolId][p.slot].push(st);
  });
  return { date: today, schools: schools };
}

// 組出提醒文字；沒有未點名的學生回傳空字串
function buildReminderText_(r, data) {
  var lines = [], schools = 0, total = 0;
  Object.keys(data.schools).map(function (k) { return data.schools[k]; })
    .sort(function (a, b) { return String(a.name).localeCompare(String(b.name), 'zh-Hant'); })
    .forEach(function (sc) {
      var list = sortStudents(sc[r.slot].slice());
      var walk = r.slot === 'pm' ? sortStudents(sc.walk.slice()) : [];   // 固定自走／自送的只加在下午接提醒，放在該校名單最後面
      if (!list.length && !walk.length) return;
      schools++; total += list.length + walk.length;
      var nm = function (st) { return st.name + '(' + st.grade + ')'; };
      lines.push('');
      lines.push('【' + sc.name + '】' + (list.length + walk.length) + ' 位');
      if (list.length) lines.push(list.map(nm).join('、'));
      walk.forEach(function (st) { lines.push('📌 ' + nm(st) + '[固定' + st.tag + ']'); });   // 固定自走／自送：每人一行，前面加 📌（和點名畫面的提示一致）
    });
  if (!total) return '';
  var hm = ('0' + r.hour).slice(-2) + ':' + ('0' + r.minute).slice(-2);
  return '🔔 ' + r.label + ' 還沒點名（' + hm + '）\n' + lineDateLabel_(data.date) + '\n共 ' + schools + ' 校、' + total + ' 位' +
    lines.join('\n') + '\n\n請負責的老師盡快點名';
}

// 送出一個時段的提醒；force=true（手動測試）時沒有未點名也會回覆一則說明
function sendReminder_(r, force) {
  _memo = {};
  var data = unmarkedBySlot_();
  if (!data) return force ? linePush_('今天不是上課日（例假日／國定假日），不發送「' + r.label + '」提醒') : { ok: true, skipped: 'CLOSED' };
  var text = buildReminderText_(r, data);
  if (!text) return force ? linePush_('✅ ' + r.label + '的學生都已點名完成，目前沒有未點名') : { ok: true, skipped: 'ALL_MARKED' };
  return linePush_(text);
}

// 每 5 分鐘執行：到了 13:00 / 16:30（之後 30 分鐘內）就發一次，同一天同一時段只發一次
function lineReminderTick() {
  if (!lineConfigured_()) return;
  var mins = parseInt(Utilities.formatDate(new Date(), TZ, 'H'), 10) * 60 + parseInt(Utilities.formatDate(new Date(), TZ, 'm'), 10);
  var today = today_(), p = lineProps_();
  LINE_REMINDERS.forEach(function (r) {
    var at = r.hour * 60 + r.minute;
    if (mins < at || mins >= at + LINE_REMINDER_GRACE_MIN) return;
    var key = 'LINE_SENT_' + r.slot;
    var claimed = withLock_(function () {
      if (p.getProperty(key) === today) return false;
      p.setProperty(key, today);   // 先登記，避免重複發送
      return true;
    });
    if (!claimed) return;
    var res = sendReminder_(r, false);
    if (res && res.ok === false) p.deleteProperty(key);   // 發送失敗 → 下一輪（寬限時間內）再試
  });
}

function createLineReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'lineReminderTick') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('lineReminderTick').timeBased().everyMinutes(5).create();
}

// LINE Webhook：在老師群組傳「綁定點名通知」→ 把該群組存成通知對象
function handleLineWebhook_(body) {
  (body.events || []).forEach(function (ev) {
    if (!ev || ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
    if (String(ev.message.text).trim() !== LINE_BIND_KEYWORD) return;
    var src = ev.source || {};
    var id = src.groupId || src.roomId;
    if (!id) return;   // 只綁定群組／多人聊天室，不綁定一對一聊天
    lineProps_().setProperty('LINE_GROUP_ID', id);
    linePush_('✅ 已綁定這個群組為「基翔點名」通知對象\n之後會在這裡通知：新增預約請假／臨時進班、13:00 中午接未點名、16:30 下午接未點名');
  });
}

// ---------- 手動執行用（在編輯器選函式 > 執行）----------
// 第一次使用：授權、建立每 5 分鐘的排程、傳一則測試訊息
function setupLine() {
  createLineReminderTrigger();
  var r = lineConfigured_() ? linePush_('✅ 基翔點名 LINE 通知測試成功') : { ok: false, error: '尚未設定指令碼屬性 LINE_CHANNEL_TOKEN / LINE_GROUP_ID' };
  Logger.log(JSON.stringify(r));
  return r;
}
// 立刻各發一次「中午接」「下午接」的未點名提醒（不受時間與每日一次限制，測試用）
function testLineReminders() {
  LINE_REMINDERS.forEach(function (r) { Logger.log(r.label + ': ' + JSON.stringify(sendReminder_(r, true))); });
}
