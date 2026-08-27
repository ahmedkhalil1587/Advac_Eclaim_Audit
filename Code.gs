/**
 * Claims Audit — Auth Backend (Google Apps Script)
 * =================================================
 * منطق كامل لتسجيل المستخدمين، تفعيل الإيميل بكود، موافقة الأدمن،
 * تسجيل الدخول، نسيان كلمة المرور، وتعطيل/تفعيل الحسابات.
 *
 * الإعداد (مرة واحدة):
 * 1) افتح Google Sheet جديد -> Extensions > Apps Script -> الصق الكود ده.
 * 2) من قائمة Run اختار الدالة initSheet وشغّلها مرة واحدة (هتطلب صلاحيات، وافق).
 * 3) من File > Project Properties أو Project Settings > Script Properties ضيف:
 *      PEPPER            = أي نص عشوائي طويل وسري (يُستخدم في تشفير الباسورد)
 *      ADMIN_SETUP_KEY   = مفتاح سري تستخدمه مرة واحدة بس لإنشاء أول حساب أدمن
 * 4) Deploy > New deployment > Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    وانسخ الرابط (Web app URL) وحطه في index.html في APPS_SCRIPT_URL
 * 5) أنشئ أول أدمن (مرة واحدة) عن طريق استدعاء action=setupFirstAdmin بالمفتاح السري
 *    (فيه زرار جاهز لده في صفحة الإعداد بـ index.html)
 * 6) بعد ما تعمل أول أدمن، امسح أو غيّر ADMIN_SETUP_KEY عشان محدش يقدر يستخدمه تاني.
 */

const USERS_SHEET = 'Users';
const SESSIONS_SHEET = 'Sessions';
const CODE_EXPIRY_MIN = 15;          // صلاحية كود التفعيل / إعادة التعيين بالدقايق
const SESSION_EXPIRY_DAYS = 7;       // صلاحية جلسة الدخول بالأيام

// ---------- Setup ----------

function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let users = ss.getSheetByName(USERS_SHEET);
  if (!users) users = ss.insertSheet(USERS_SHEET);
  users.clear();
  users.appendRow([
    'ID', 'Name', 'Email', 'PasswordHash', 'Salt', 'Role', 'Status',
    'EmailVerified', 'VerifyCode', 'VerifyExpiry', 'ResetCode', 'ResetExpiry',
    'CreatedAt', 'ApprovedBy', 'ApprovedAt'
  ]);
  users.setFrozenRows(1);

  let sessions = ss.getSheetByName(SESSIONS_SHEET);
  if (!sessions) sessions = ss.insertSheet(SESSIONS_SHEET);
  sessions.clear();
  sessions.appendRow(['Token', 'Email', 'Role', 'ExpiresAt']);
  sessions.setFrozenRows(1);

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Logger.log('تم إعداد الشيتات بنجاح.');
}

// ---------- HTTP entry points ----------

function doGet(e) {
  return jsonOut({ ok: true, message: 'Claims Audit Auth API is running' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    switch (action) {
      case 'register':          result = registerUser(body); break;
      case 'verifyEmail':       result = verifyEmail(body); break;
      case 'login':              result = loginUser(body); break;
      case 'forgotPassword':    result = forgotPassword(body); break;
      case 'resetPassword':     result = resetPassword(body); break;
      case 'listPending':       result = listPending(body); break;
      case 'approveUser':       result = approveUser(body); break;
      case 'rejectUser':        result = rejectUser(body); break;
      case 'disableUser':       result = setUserStatus(body, 'disabled'); break;
      case 'enableUser':        result = setUserStatus(body, 'active'); break;
      case 'listUsers':         result = listUsers(body); break;
      case 'setupFirstAdmin':   result = setupFirstAdmin(body); break;
      case 'validateSession':   result = validateSessionAction(body); break;
      default:                  result = { ok: false, error: 'إجراء غير معروف' };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Helpers ----------

function getUsersSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
}
function getSessionsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
}

function generateSalt() {
  return Utilities.getUuid();
}

function generateCode() {
  return String(Math.floor(100000 + Math.random() * 900000)); // 6 digits
}

function hashPassword(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PEPPER') || '';
  const raw = salt + '::' + password + '::' + pepper;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => ((b < 0 ? b + 256 : b).toString(16).padStart(2, '0'))).join('');
}

function findUserRow(email) {
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  const norm = String(email || '').trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim().toLowerCase() === norm) {
      return { rowIndex: i + 1, row: data[i], headers: data[0] };
    }
  }
  return null;
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, idx) => obj[h] = row[idx]);
  return obj;
}

function sendMail(to, subject, body) {
  MailApp.sendEmail(to, subject, body);
}

// ---------- Session ----------

function createSession(email, role) {
  const token = Utilities.getUuid();
  const expires = new Date(Date.now() + SESSION_EXPIRY_DAYS * 24 * 60 * 60 * 1000);
  getSessionsSheet().appendRow([token, email, role, expires]);
  return token;
}

function validateSession(token) {
  if (!token) return null;
  const sheet = getSessionsSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      const expires = new Date(data[i][3]);
      if (expires.getTime() < Date.now()) return null;
      return { email: data[i][1], role: data[i][2] };
    }
  }
  return null;
}

function validateSessionAction(body) {
  const session = validateSession(body.token);
  if (!session) return { ok: false, error: 'الجلسة منتهية، سجّل دخول تاني' };
  return { ok: true, email: session.email, role: session.role };
}

function requireAdmin(token) {
  const session = validateSession(token);
  if (!session) throw new Error('الجلسة منتهية، سجّل دخول تاني');
  if (session.role !== 'admin') throw new Error('الصلاحية دي للأدمن بس');
  return session;
}

function invalidateSessionsForEmail(email) {
  const sheet = getSessionsSheet();
  const data = sheet.getDataRange().getValues();
  const norm = String(email).trim().toLowerCase();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1] || '').trim().toLowerCase() === norm) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ---------- Registration & verification ----------

function registerUser(body) {
  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');

  if (!name || !email || !password) return { ok: false, error: 'البيانات ناقصة' };
  if (password.length < 6) return { ok: false, error: 'كلمة المرور لازم تكون 6 حروف/أرقام على الأقل' };

  const existing = findUserRow(email);
  if (existing) {
    const status = existing.row[6];
    if (status === 'active' || status === 'pending_verification' || status === 'pending_approval') {
      return { ok: false, error: 'الإيميل ده مسجل بالفعل' };
    }
  }

  const salt = generateSalt();
  const hash = hashPassword(password, salt);
  const code = generateCode();
  const expiry = new Date(Date.now() + CODE_EXPIRY_MIN * 60 * 1000);
  const now = new Date();

  const sheet = getUsersSheet();
  if (existing) {
    // إعادة استخدام صف مرفوض سابقًا
    const r = existing.rowIndex;
    sheet.getRange(r, 1, 1, 15).setValues([[
      Utilities.getUuid(), name, email, hash, salt, 'user', 'pending_verification',
      false, code, expiry, '', '', now, '', ''
    ]]);
  } else {
    sheet.appendRow([
      Utilities.getUuid(), name, email, hash, salt, 'user', 'pending_verification',
      false, code, expiry, '', '', now, '', ''
    ]);
  }

  sendMail(email, 'كود تفعيل حسابك — Claims Audit',
    'أهلاً ' + name + ',\n\nكود تفعيل حسابك هو: ' + code +
    '\nالكود صالح لمدة ' + CODE_EXPIRY_MIN + ' دقيقة.\n\nلو مطلبتش التسجيل، تجاهل الإيميل ده.');

  return { ok: true, message: 'تم إرسال كود التفعيل على إيميلك' };
}

function verifyEmail(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const code = String(body.code || '').trim();

  const found = findUserRow(email);
  if (!found) return { ok: false, error: 'الحساب غير موجود' };

  const row = found.row;
  if (row[6] !== 'pending_verification') return { ok: false, error: 'الحساب مش في حالة انتظار تفعيل' };
  if (String(row[8]) !== code) return { ok: false, error: 'الكود غلط' };
  if (new Date(row[9]).getTime() < Date.now()) return { ok: false, error: 'الكود منتهي الصلاحية' };

  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 7).setValue('pending_approval'); // Status
  sheet.getRange(found.rowIndex, 8).setValue(true);               // EmailVerified
  sheet.getRange(found.rowIndex, 9).setValue('');                 // clear code
  sheet.getRange(found.rowIndex, 10).setValue('');

  return { ok: true, message: 'تم تفعيل الإيميل، طلبك دلوقتي في انتظار موافقة الأدمن' };
}

// ---------- Login ----------

function loginUser(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');

  const found = findUserRow(email);
  if (!found) return { ok: false, error: 'بيانات الدخول غلط' };

  const row = found.row;
  const status = row[6];
  if (status === 'pending_verification') return { ok: false, error: 'لازم تفعّل إيميلك الأول بالكود اللي اتبعتلك' };
  if (status === 'pending_approval') return { ok: false, error: 'طلبك لسه في انتظار موافقة الأدمن' };
  if (status === 'rejected') return { ok: false, error: 'طلب التسجيل اتراض' };
  if (status === 'disabled') return { ok: false, error: 'الحساب ده متعطّل، كلم الأدمن' };
  if (status !== 'active') return { ok: false, error: 'الحساب مش نشط' };

  const salt = row[4];
  const expectedHash = row[3];
  const actualHash = hashPassword(password, salt);
  if (actualHash !== expectedHash) return { ok: false, error: 'بيانات الدخول غلط' };

  const role = row[5];
  const token = createSession(email, role);
  return { ok: true, token: token, role: role, name: row[1], email: email };
}

// ---------- Forgot / reset password ----------

function forgotPassword(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const found = findUserRow(email);
  // رسالة عامة دايمًا عشان محدش يعرف يتأكد مين مسجل ومين لأ
  const generic = { ok: true, message: 'لو الإيميل ده مسجل عندنا، هيوصلك كود إعادة تعيين كلمة المرور' };
  if (!found) return generic;
  if (found.row[6] !== 'active') return generic;

  const code = generateCode();
  const expiry = new Date(Date.now() + CODE_EXPIRY_MIN * 60 * 1000);
  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 11).setValue(code);
  sheet.getRange(found.rowIndex, 12).setValue(expiry);

  sendMail(email, 'كود إعادة تعيين كلمة المرور — Claims Audit',
    'كود إعادة تعيين كلمة المرور هو: ' + code + '\nصالح لمدة ' + CODE_EXPIRY_MIN + ' دقيقة.');

  return generic;
}

function resetPassword(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const code = String(body.code || '').trim();
  const newPassword = String(body.newPassword || '');

  if (newPassword.length < 6) return { ok: false, error: 'كلمة المرور لازم تكون 6 حروف/أرقام على الأقل' };

  const found = findUserRow(email);
  if (!found) return { ok: false, error: 'طلب غير صالح' };
  const row = found.row;
  if (!row[10] || String(row[10]) !== code) return { ok: false, error: 'الكود غلط' };
  if (new Date(row[11]).getTime() < Date.now()) return { ok: false, error: 'الكود منتهي الصلاحية' };

  const salt = generateSalt();
  const hash = hashPassword(newPassword, salt);
  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 4).setValue(hash);  // PasswordHash
  sheet.getRange(found.rowIndex, 5).setValue(salt);  // Salt
  sheet.getRange(found.rowIndex, 11).setValue('');   // clear reset code
  sheet.getRange(found.rowIndex, 12).setValue('');

  invalidateSessionsForEmail(email); // اخرج من كل الجلسات القديمة بعد تغيير الباسورد

  return { ok: true, message: 'تم تغيير كلمة المرور بنجاح' };
}

// ---------- Admin actions ----------

function listPending(body) {
  requireAdmin(body.token);
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'pending_approval') {
      out.push({ name: data[i][1], email: data[i][2], createdAt: data[i][12] });
    }
  }
  return { ok: true, users: out };
}

function approveUser(body) {
  const admin = requireAdmin(body.token);
  const found = findUserRow(body.email);
  if (!found) return { ok: false, error: 'المستخدم غير موجود' };
  if (found.row[6] !== 'pending_approval') return { ok: false, error: 'الحساب مش في حالة انتظار موافقة' };

  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 7).setValue('active');
  sheet.getRange(found.rowIndex, 14).setValue(admin.email);
  sheet.getRange(found.rowIndex, 15).setValue(new Date());

  sendMail(found.row[2], 'تم قبول حسابك — Claims Audit', 'أهلاً ' + found.row[1] + ',\n\nتم قبول حساب وتقدر تسجل دخول دلوقتي.');
  return { ok: true, message: 'تم قبول المستخدم' };
}

function rejectUser(body) {
  requireAdmin(body.token);
  const found = findUserRow(body.email);
  if (!found) return { ok: false, error: 'المستخدم غير موجود' };

  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 7).setValue('rejected');

  sendMail(found.row[2], 'طلب التسجيل — Claims Audit', 'أهلاً ' + found.row[1] + ',\n\nنأسف، تم رفض طلب تسجيلك. لو معتقدش ده صح كلم الأدمن.');
  return { ok: true, message: 'تم رفض الطلب' };
}

function setUserStatus(body, newStatus) {
  const admin = requireAdmin(body.token);
  const found = findUserRow(body.email);
  if (!found) return { ok: false, error: 'المستخدم غير موجود' };
  if (String(found.row[2]).toLowerCase() === admin.email.toLowerCase() && newStatus === 'disabled') {
    return { ok: false, error: 'مينفعش تعطّل حسابك انت' };
  }

  const sheet = getUsersSheet();
  sheet.getRange(found.rowIndex, 7).setValue(newStatus);
  if (newStatus === 'disabled') invalidateSessionsForEmail(found.row[2]);

  return { ok: true, message: newStatus === 'disabled' ? 'تم تعطيل الحساب' : 'تم تفعيل الحساب' };
}

function listUsers(body) {
  requireAdmin(body.token);
  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    out.push({
      name: data[i][1], email: data[i][2], role: data[i][5],
      status: data[i][6], createdAt: data[i][12]
    });
  }
  return { ok: true, users: out };
}

// ---------- One-time first-admin bootstrap ----------

function setupFirstAdmin(body) {
  const setupKey = PropertiesService.getScriptProperties().getProperty('ADMIN_SETUP_KEY');
  if (!setupKey) return { ok: false, error: 'ADMIN_SETUP_KEY مش متظبط في Script Properties' };
  if (String(body.setupKey || '') !== setupKey) return { ok: false, error: 'مفتاح الإعداد غلط' };

  const name = String(body.name || '').trim();
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  if (!name || !email || password.length < 6) return { ok: false, error: 'بيانات ناقصة' };

  const salt = generateSalt();
  const hash = hashPassword(password, salt);
  const sheet = getUsersSheet();

  const existing = findUserRow(email);
  const rowValues = [
    Utilities.getUuid(), name, email, hash, salt, 'admin', 'active',
    true, '', '', '', '', new Date(), 'system', new Date()
  ];
  if (existing) {
    sheet.getRange(existing.rowIndex, 1, 1, 15).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { ok: true, message: 'تم إنشاء حساب الأدمن. غيّر أو امسح ADMIN_SETUP_KEY دلوقتي.' };
}
