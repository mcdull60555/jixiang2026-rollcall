/**
 * 學生點名系統 - 後端 (Google Apps Script)
 * ============================================
 * 部署方式：
 * 1. 建立一個新的 Google 試算表 (Google Sheet)，並複製網址列裡的試算表 ID
 *    （網址長得像 https://docs.google.com/spreadsheets/d/【這一長串就是ID】/edit），
 *    貼到下面的 SPREADSHEET_ID。
 * 2. 在試算表選單「擴充功能 > Apps Script」，把編輯器裡原本的程式碼全部刪除，
 *    貼上這個檔案全部內容(含你剛填好的 SPREADSHEET_ID)。
 *    （若你是另外在 script.google.com 開獨立專案寫的，也完全沒問題，
 *      因為程式改用 SPREADSHEET_ID 指定試算表，不依賴「目前綁定哪個檔案」）
 * 3. 執行一次 setup() 函式（上方選單選 setup，再按執行；第一次會跳出權限授權視窗，點允許），
 *    它會自動建立所有需要的工作表分頁，並建立管理員帳號。
 *    管理員帳號密碼：帳號 admin / 密碼 1234 (可登入後在系統裡自行修改)
 * 4. 點「部署 > 新增部署作業」，類型選「網頁應用程式」：
 *    - 執行身分：我 (your account)
 *    - 存取權限：任何人
 *    部署後會得到一個 /exec 結尾的網址，把它貼到 index.html 最上面的 API_URL。
 *    （之後若修改這份程式碼，要讓網址生效，須到「部署 > 管理部署作業」，
 *      點編輯(鉛筆)圖示、版本選「新版本」、再按部署，/exec 網址不會變）
 * 5. (建議) 選單執行 createDailyCleanupTrigger() 一次，
 *    之後系統會每天自動清除超過 60 天的點名紀錄。
 */

// ---------- 基本設定 ----------
var SPREADSHEET_ID = '';      // 貼上你的 Google 試算表 ID 或整個網址都可以 (留空則嘗試用目前綁定的試算表)
var RETENTION_DAYS = 60;      // 點名紀錄保留天數
var SESSION_HOURS = 12;       // 登入 token 有效時數

var SHEET_ADMINS = 'Admins';
var SHEET_TEACHERS = 'Teachers';
var SHEET_SCHOOLS = 'Schools';
var SHEET_STUDENTS = 'Students';
var SHEET_ATTENDANCE = 'Attendance';
var SHEET_SESSIONS = 'Sessions';

var HEADERS = {
  Admins: ['id', 'username', 'password', 'name'],
  Teachers: ['id', 'username', 'password', 'name', 'active', 'createdAt'],
  Schools: ['id', 'name', 'createdAt'],
  Students: ['id', 'schoolId', 'name', 'grade', 'contactName', 'contactPhone', 'fixedOffDays', 'active', 'createdAt'],
  Attendance: ['id', 'date', 'schoolId', 'studentId', 'status', 'selfDrop', 'selfWalk', 'teacherId', 'teacherName', 'submittedAt'],
  Sessions: ['token', 'userId', 'role', 'name', 'username', 'createdAt', 'expiresAt']
};

// ---------- 初始化 ----------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[name]);
      sheet.setFrozenRows(1);
    }
  });
  // 刪除預設的 Sheet1 (若還存在且是空的)
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  // 建立預設管理員帳號
  var adminSheet = ss.getSheetByName(SHEET_ADMINS);
  if (adminSheet.getLastRow() < 2) {
    adminSheet.appendRow([genId('A'), 'admin', hashPwd('1234'), '管理員']);
  }
  createDailyCleanupTrigger();
  return 'setup done';
}

function createDailyCleanupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cleanupOldRecords') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupOldRecords').timeBased().everyDays(1).atHour(3).create();
}

// ---------- 工具函式 ----------
function extractSpreadsheetId_(raw) {
  var s = String(raw || '').trim();
  // 如果貼的是整個網址，自動擷取 /d/ 和下一個 / 之間的那一段
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  // 否則假設使用者已經直接貼了純 ID，去掉可能誤貼的前後斜線或 edit 字樣
  return s.replace(/^\/+|\/+$/g, '').replace(/\/edit.*$/, '');
}

function ss_() {
  var ss;
  if (SPREADSHEET_ID) {
    var id = extractSpreadsheetId_(SPREADSHEET_ID);
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      throw new Error('無法開啟試算表，請確認 SPREADSHEET_ID 是否正確（只需要網址中 /d/ 和 /edit 之間那一段）：' + e.message);
    }
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error('找不到試算表：請在程式碼最上方的 SPREADSHEET_ID 填入你的 Google 試算表 ID');
  }
  return ss;
}
function sheet_(name) {
  var sheet = ss_().getSheetByName(name);
  if (!sheet) {
    throw new Error('找不到工作表「' + name + '」：請先執行一次 setup() 函式建立所有分頁');
  }
  return sheet;
}

function genId(prefix) {
  return (prefix || 'X') + Utilities.getUuid().replace(/-/g, '').substring(0, 10);
}

function hashPwd(pwd) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pwd), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function sheetToObjects(name) {
  var sheet = sheet_(name);
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = rows[i][j];
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function appendObject(name, obj) {
  var sheet = sheet_(name);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return (obj[h] !== undefined ? obj[h] : ''); });
  sheet.appendRow(row);
}

function updateObjectByRow(name, rowIndex, obj) {
  var sheet = sheet_(name);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return (obj[h] !== undefined ? obj[h] : ''); });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function fmtDate(d) {
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd');
}

function today_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
}

function weekdayOf(dateStr) {
  // 回傳 1(一)~7(日)
  var d = new Date(dateStr + 'T00:00:00+08:00');
  var wd = d.getDay(); // 0=Sun..6=Sat
  return wd === 0 ? 7 : wd;
}

// ---------- Session ----------
function createSession(user, role) {
  var token = Utilities.getUuid();
  var now = new Date();
  var expires = new Date(now.getTime() + SESSION_HOURS * 3600 * 1000);
  appendObject(SHEET_SESSIONS, {
    token: token, userId: user.id, role: role, name: user.name,
    username: user.username, createdAt: now.toISOString(), expiresAt: expires.toISOString()
  });
  return token;
}

function getSession(token) {
  if (!token) return null;
  var sessions = sheetToObjects(SHEET_SESSIONS);
  for (var i = sessions.length - 1; i >= 0; i--) {
    var s = sessions[i];
    if (s.token === token) {
      if (new Date(s.expiresAt).getTime() < Date.now()) return null;
      return s;
    }
  }
  return null;
}

function requireSession(token, role) {
  var s = getSession(token);
  if (!s) throw new Error('AUTH_EXPIRED');
  if (role && s.role !== role) throw new Error('FORBIDDEN');
  return s;
}

// ---------- HTTP 入口 ----------
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '點名系統 API 運作中' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    result = route(body);
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function route(body) {
  var action = body.action;
  switch (action) {
    case 'login': return actionLogin(body);
    case 'bootstrap': return actionBootstrap(body);
    case 'listSchools': return actionListSchools(body);
    case 'addSchool': return actionAddSchool(body);
    case 'listStudents': return actionListStudents(body);
    case 'addStudent': return actionAddStudent(body);
    case 'editStudent': return actionEditStudent(body);
    case 'deleteStudent': return actionDeleteStudent(body);
    case 'listTeachers': return actionListTeachers(body);
    case 'addTeacher': return actionAddTeacher(body);
    case 'deleteTeacher': return actionDeleteTeacher(body);
    case 'changePwd': return actionChangePwd(body);
    case 'getRollCall': return actionGetRollCall(body);
    case 'submitAttendance': return actionSubmitAttendance(body);
    case 'getOverview': return actionGetOverview(body);
    case 'getSchoolDayDetail': return actionGetSchoolDayDetail(body);
    default: return { ok: false, error: 'UNKNOWN_ACTION' };
  }
}

// ---------- Actions ----------
function actionLogin(body) {
  var username = String(body.username || '').trim();
  var password = String(body.password || '');
  var hashed = hashPwd(password);

  var admins = sheetToObjects(SHEET_ADMINS);
  for (var i = 0; i < admins.length; i++) {
    if (admins[i].username === username && admins[i].password === hashed) {
      var token = createSession(admins[i], 'admin');
      return { ok: true, token: token, role: 'admin', name: admins[i].name };
    }
  }
  var teachers = sheetToObjects(SHEET_TEACHERS);
  for (var j = 0; j < teachers.length; j++) {
    var t = teachers[j];
    if (t.username === username && t.password === hashed) {
      if (String(t.active) === 'false') return { ok: false, error: 'ACCOUNT_DISABLED' };
      var token2 = createSession(t, 'teacher');
      return { ok: true, token: token2, role: 'teacher', name: t.name };
    }
  }
  return { ok: false, error: 'INVALID_LOGIN' };
}

function actionBootstrap(body) {
  var s = requireSession(body.token);
  return { ok: true, role: s.role, name: s.name, today: today_() };
}

function actionListSchools(body) {
  requireSession(body.token);
  var schools = sheetToObjects(SHEET_SCHOOLS);
  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) { return String(x.active) !== 'false'; });
  var out = schools.map(function (sc) {
    var count = students.filter(function (st) { return st.schoolId === sc.id; }).length;
    return { id: sc.id, name: sc.name, studentCount: count };
  });
  out.sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hant'); });
  return { ok: true, schools: out };
}

function actionAddSchool(body) {
  requireSession(body.token, 'admin');
  var name = String(body.name || '').trim();
  if (!name) return { ok: false, error: 'NAME_REQUIRED' };
  var id = genId('S');
  appendObject(SHEET_SCHOOLS, { id: id, name: name, createdAt: new Date().toISOString() });
  return { ok: true, id: id };
}

function sortStudents(list) {
  list.sort(function (a, b) {
    var ga = Number(a.grade) || 0, gb = Number(b.grade) || 0;
    if (ga !== gb) return ga - gb;
    return String(a.name).localeCompare(String(b.name), 'zh-Hant');
  });
  return list;
}

function actionListStudents(body) {
  requireSession(body.token);
  var schoolId = body.schoolId;
  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) {
    return x.schoolId === schoolId && String(x.active) !== 'false';
  });
  var out = students.map(function (x) {
    return {
      id: x.id, name: x.name, grade: x.grade, contactName: x.contactName,
      contactPhone: x.contactPhone,
      fixedOffDays: x.fixedOffDays ? String(x.fixedOffDays).split(',').filter(String).map(Number) : []
    };
  });
  sortStudents(out);
  return { ok: true, students: out };
}

function actionAddStudent(body) {
  requireSession(body.token, 'admin');
  var id = genId('T');
  appendObject(SHEET_STUDENTS, {
    id: id, schoolId: body.schoolId, name: body.name, grade: body.grade,
    contactName: body.contactName || '', contactPhone: body.contactPhone || '',
    fixedOffDays: (body.fixedOffDays || []).join(','), active: true,
    createdAt: new Date().toISOString()
  });
  return { ok: true, id: id };
}

function actionEditStudent(body) {
  requireSession(body.token, 'admin');
  var students = sheetToObjects(SHEET_STUDENTS);
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === body.studentId) {
      var updated = students[i];
      if (body.name !== undefined) updated.name = body.name;
      if (body.grade !== undefined) updated.grade = body.grade;
      if (body.contactName !== undefined) updated.contactName = body.contactName;
      if (body.contactPhone !== undefined) updated.contactPhone = body.contactPhone;
      if (body.fixedOffDays !== undefined) updated.fixedOffDays = (body.fixedOffDays || []).join(',');
      updateObjectByRow(SHEET_STUDENTS, students[i]._row, updated);
      return { ok: true };
    }
  }
  return { ok: false, error: 'STUDENT_NOT_FOUND' };
}

function actionDeleteStudent(body) {
  requireSession(body.token, 'admin');
  var students = sheetToObjects(SHEET_STUDENTS);
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === body.studentId) {
      var updated = students[i];
      updated.active = false;
      updateObjectByRow(SHEET_STUDENTS, students[i]._row, updated);
      return { ok: true };
    }
  }
  return { ok: false, error: 'STUDENT_NOT_FOUND' };
}

function actionListTeachers(body) {
  requireSession(body.token, 'admin');
  var teachers = sheetToObjects(SHEET_TEACHERS);
  var out = teachers.map(function (t) {
    return { id: t.id, username: t.username, name: t.name, active: String(t.active) !== 'false' };
  });
  return { ok: true, teachers: out };
}

function actionAddTeacher(body) {
  requireSession(body.token, 'admin');
  var username = String(body.username || '').trim();
  if (!username) return { ok: false, error: 'USERNAME_REQUIRED' };
  var teachers = sheetToObjects(SHEET_TEACHERS);
  for (var i = 0; i < teachers.length; i++) {
    if (teachers[i].username === username) return { ok: false, error: 'USERNAME_TAKEN' };
  }
  var id = genId('TC');
  appendObject(SHEET_TEACHERS, {
    id: id, username: username, password: hashPwd(body.password || '0000'),
    name: body.name || username, active: true, createdAt: new Date().toISOString()
  });
  return { ok: true, id: id };
}

function actionDeleteTeacher(body) {
  requireSession(body.token, 'admin');
  var teachers = sheetToObjects(SHEET_TEACHERS);
  for (var i = 0; i < teachers.length; i++) {
    if (teachers[i].id === body.teacherId) {
      var updated = teachers[i];
      updated.active = false;
      updateObjectByRow(SHEET_TEACHERS, teachers[i]._row, updated);
      return { ok: true };
    }
  }
  return { ok: false, error: 'TEACHER_NOT_FOUND' };
}

function actionChangePwd(body) {
  var s = requireSession(body.token);
  var sheetName = s.role === 'admin' ? SHEET_ADMINS : SHEET_TEACHERS;
  var rows = sheetToObjects(sheetName);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].id === s.userId) {
      if (rows[i].password !== hashPwd(body.oldPassword || '')) return { ok: false, error: 'WRONG_OLD_PASSWORD' };
      var updated = rows[i];
      updated.password = hashPwd(body.newPassword || '');
      updateObjectByRow(sheetName, rows[i]._row, updated);
      return { ok: true };
    }
  }
  return { ok: false, error: 'USER_NOT_FOUND' };
}

function actionGetRollCall(body) {
  var s = requireSession(body.token);
  var date = body.date || today_();
  var schoolId = body.schoolId;
  var wd = weekdayOf(date);

  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) {
    return x.schoolId === schoolId && String(x.active) !== 'false';
  });
  var attendance = sheetToObjects(SHEET_ATTENDANCE).filter(function (a) {
    return a.schoolId === schoolId && a.date === date;
  });
  var attMap = {};
  attendance.forEach(function (a) { attMap[a.studentId] = a; });

  var out = students.map(function (st) {
    var offDays = st.fixedOffDays ? String(st.fixedOffDays).split(',').filter(String).map(Number) : [];
    var isFixedOff = offDays.indexOf(wd) !== -1;
    var existing = attMap[st.id];
    return {
      id: st.id, name: st.name, grade: st.grade, isFixedOff: isFixedOff,
      status: existing ? existing.status : (isFixedOff ? 'fixedOff' : ''),
      selfDrop: existing ? (String(existing.selfDrop) === 'true') : false,
      selfWalk: existing ? (String(existing.selfWalk) === 'true') : false
    };
  });
  sortStudents(out);

  var submitted = attendance.length > 0;
  var submittedAt = submitted ? attendance[0].submittedAt : null;
  var submittedBy = submitted ? attendance[0].teacherName : null;

  return { ok: true, date: date, students: out, submitted: submitted, submittedAt: submittedAt, submittedBy: submittedBy };
}

function actionSubmitAttendance(body) {
  var s = requireSession(body.token, 'teacher');
  var date = body.date || today_();
  var schoolId = body.schoolId;
  var records = body.records || [];
  var now = new Date().toISOString();

  var existing = sheetToObjects(SHEET_ATTENDANCE).filter(function (a) {
    return a.schoolId === schoolId && a.date === date;
  });
  var existingMap = {};
  existing.forEach(function (a) { existingMap[a.studentId] = a; });

  records.forEach(function (r) {
    var row = {
      date: date, schoolId: schoolId, studentId: r.studentId,
      status: r.status || '', selfDrop: !!r.selfDrop, selfWalk: !!r.selfWalk,
      teacherId: s.userId, teacherName: s.name, submittedAt: now
    };
    var old = existingMap[r.studentId];
    if (old) {
      row.id = old.id;
      updateObjectByRow(SHEET_ATTENDANCE, old._row, row);
    } else {
      row.id = genId('R');
      appendObject(SHEET_ATTENDANCE, row);
    }
  });

  cleanupOldRecords();
  return { ok: true, submittedAt: now };
}

function actionGetOverview(body) {
  requireSession(body.token, 'admin');
  var schoolFilter = body.schoolId || null;
  var days = body.days || RETENTION_DAYS;

  var schools = sheetToObjects(SHEET_SCHOOLS);
  var schoolMap = {};
  schools.forEach(function (sc) { schoolMap[sc.id] = sc.name; });

  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) { return String(x.active) !== 'false'; });
  var studentCountBySchool = {};
  students.forEach(function (st) {
    studentCountBySchool[st.schoolId] = (studentCountBySchool[st.schoolId] || 0) + 1;
  });

  var attendance = sheetToObjects(SHEET_ATTENDANCE);
  var cutoff = new Date(Date.now() - days * 86400000);

  var byDateSchool = {}; // date -> schoolId -> {present,leave,fixedOff,selfDrop,selfWalk,submittedAt,submittedBy}
  attendance.forEach(function (a) {
    var d = new Date(a.date);
    if (d.getTime() < cutoff.getTime()) return;
    if (schoolFilter && a.schoolId !== schoolFilter) return;
    byDateSchool[a.date] = byDateSchool[a.date] || {};
    var bucket = byDateSchool[a.date][a.schoolId] = byDateSchool[a.date][a.schoolId] || {
      present: 0, leave: 0, fixedOff: 0, selfDrop: 0, selfWalk: 0, submittedAt: a.submittedAt, submittedBy: a.teacherName
    };
    if (a.status === 'present') bucket.present++;
    else if (a.status === 'leave') bucket.leave++;
    else if (a.status === 'fixedOff') bucket.fixedOff++;
    if (String(a.selfDrop) === 'true') bucket.selfDrop++;
    if (String(a.selfWalk) === 'true') bucket.selfWalk++;
    if (new Date(a.submittedAt) > new Date(bucket.submittedAt)) {
      bucket.submittedAt = a.submittedAt;
      bucket.submittedBy = a.teacherName;
    }
  });

  var dates = Object.keys(byDateSchool).sort(function (a, b) { return b.localeCompare(a); });
  var result = dates.map(function (date) {
    var schoolIds = Object.keys(byDateSchool[date]);
    var schoolsOut = schoolIds.map(function (sid) {
      var b = byDateSchool[date][sid];
      return {
        schoolId: sid, schoolName: schoolMap[sid] || '(未知學校)',
        total: studentCountBySchool[sid] || 0,
        present: b.present, leave: b.leave, fixedOff: b.fixedOff,
        selfDrop: b.selfDrop, selfWalk: b.selfWalk,
        submittedAt: b.submittedAt, submittedBy: b.submittedBy
      };
    });
    schoolsOut.sort(function (a, b) { return a.schoolName.localeCompare(b.schoolName, 'zh-Hant'); });
    return { date: date, schools: schoolsOut };
  });

  return { ok: true, dates: result };
}

function actionGetSchoolDayDetail(body) {
  requireSession(body.token, 'admin');
  var schoolId = body.schoolId;
  var date = body.date;
  var students = sheetToObjects(SHEET_STUDENTS).filter(function (x) {
    return x.schoolId === schoolId && String(x.active) !== 'false';
  });
  var attendance = sheetToObjects(SHEET_ATTENDANCE).filter(function (a) {
    return a.schoolId === schoolId && a.date === date;
  });
  var attMap = {};
  attendance.forEach(function (a) { attMap[a.studentId] = a; });

  var out = students.map(function (st) {
    var a = attMap[st.id];
    return {
      id: st.id, name: st.name, grade: st.grade,
      status: a ? a.status : '', selfDrop: a ? String(a.selfDrop) === 'true' : false,
      selfWalk: a ? String(a.selfWalk) === 'true' : false
    };
  });
  sortStudents(out);
  return { ok: true, students: out };
}

// ---------- 定期清理 (超過 60 天的點名紀錄) ----------
function cleanupOldRecords() {
  var sheet = sheet_(SHEET_ATTENDANCE);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;
  var headers = rows[0];
  var dateCol = headers.indexOf('date');
  var cutoff = new Date(Date.now() - RETENTION_DAYS * 86400000);
  var keep = [headers];
  for (var i = 1; i < rows.length; i++) {
    var d = new Date(rows[i][dateCol]);
    if (d.getTime() >= cutoff.getTime()) keep.push(rows[i]);
  }
  if (keep.length < rows.length) {
    sheet.clearContents();
    sheet.getRange(1, 1, keep.length, headers.length).setValues(keep);
    sheet.setFrozenRows(1);
  }
}
