/**
 * אתר הזמנות - נבחרת צימרי מבואות החרמון
 * Google Apps Script - מחובר לגיליון (Extensions > Apps Script מתוך הגיליון)
 *
 * חשוב: אחרי כל שינוי בקוד הזה צריך פריסה חדשה:
 * Deploy > Manage deployments > עיפרון > Version: New version > Deploy
 *
 * שלב 2: הקמת הגיליון + קריאת קטלוג ומלאי (doGet?action=catalog)
 */

// ---------- שמות הלשוניות ----------
var SH = {
  PRODUCTS:  'מוצרים',
  ORDERS:    'הזמנות',
  ITEMS:     'פריטי הזמנה',
  MOVES:     'תנועות מלאי',
  SETTINGS:  'הגדרות',
  TOTALS:    'ריכוז כמויות',
  BY_ZIMMER: 'סיכום לפי צימר'
};

var HEADERS = {};
HEADERS[SH.PRODUCTS] = ['מזהה', 'קטגוריה', 'שם קטגוריה', 'קוד', 'מק"ט', 'שם', 'תיאור מורחב', 'צבע', 'מחיר',
  'מלאי התחלתי', 'יתרה', 'ללא הגבלה', 'תמונה', 'סוג', 'גוון', 'קוטר', 'מתנה', 'פעיל', 'הערה פנימית', 'סדר'];
HEADERS[SH.ORDERS] = ['מספר הזמנה', 'תאריך', 'שם הצימר', 'ישוב', 'איש קשר', 'טלפון', 'מייל', 'כתובת לאספקה',
  'יום הגעה לאולם', 'סה"כ', 'תקציב', 'חריגה', 'אישור חריגה', 'סטטוס', 'סטטוס מייל', 'קובץ Excel', 'הערות צוות'];
HEADERS[SH.ITEMS] = ['מספר הזמנה', 'תאריך', 'שם הצימר', 'מזהה', 'קוד', 'מק"ט', 'שם', 'צבע', 'פירוט', 'כמות',
  'מחיר יחידה', 'סה"כ שורה', 'סטטוס הזמנה'];
HEADERS[SH.MOVES] = ['תאריך', 'מי', 'מזהה', 'מק"ט', 'שינוי', 'יתרה אחרי', 'סוג', 'הפניה', 'הערה'];

var DEFAULT_SETTINGS = [
  ['תקציב', 9500, 'כולל מע"מ'],
  ['נמעני מייל צוות', '', 'כתובות מופרדות בפסיק'],
  ['סף מלאי נמוך', 3, 'מתחת לזה - סימון כתום במסכי הניהול'],
  ['שם התוכנית', 'נבחרת צימרי מבואות החרמון', ''],
  ['פתוח להזמנות', 'כן', 'כן / לא']
];

var HEADER_COLOR = '#5C6B3A'; // ירוק זית

// ---------- הקמה (להריץ פעם אחת מהעורך) ----------
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // שומרים את מזהה הגיליון, כדי שאפליקציית האינטרנט תפתח אותו לפי מזהה (openById)
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  SpreadsheetApp.openById(ss.getId()); // מבקש את הרשאת הגיליונות המלאה כבר עכשיו
  ss.setSpreadsheetLocale('iw_IL');
  [SH.PRODUCTS, SH.ORDERS, SH.ITEMS, SH.MOVES].forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    styleSheet_(sh);
  });
  var st = ss.getSheetByName(SH.SETTINGS) || ss.insertSheet(SH.SETTINGS);
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 1, 3).setValues([['הגדרה', 'ערך', 'הסבר']]);
    st.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  }
  styleSheet_(st);
  // מחיקת לשונית ברירת המחדל הריקה אם קיימת
  ['Sheet1', 'גיליון1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  // יתרה ריקה = מלאי התחלתי
  var p = readTable_(SH.PRODUCTS);
  var colRem = p.col['יתרה'], colInit = p.col['מלאי התחלתי'];
  p.rows.forEach(function (r, i) {
    if (r[colRem] === '' && r[colInit] !== '') p.sheet.getRange(i + 2, colRem + 1).setValue(r[colInit]);
  });
  Logger.log('הגיליון מוכן. מוצרים: ' + p.rows.length);
}

function styleSheet_(sh) {
  sh.setRightToLeft(true);
  sh.setFrozenRows(1);
  var lc = Math.max(sh.getLastColumn(), 1);
  sh.getRange(1, 1, 1, lc).setBackground(HEADER_COLOR).setFontColor('#ffffff').setFontWeight('bold');
}

// בדיקת החשבון והמכסה היומית של מיילים (להריץ מהעורך ולקרוא ב-Execution log)
function checkAccount() {
  var email = Session.getEffectiveUser().getEmail();
  var isGmail = /@gmail\.com$/i.test(email);
  Logger.log('הסקריפט רץ מהחשבון: ' + email);
  Logger.log('סוג חשבון: ' + (isGmail ? 'Gmail רגיל - עד 100 נמענים ביום' : 'Workspace - עד 1,500 נמענים ביום'));
  Logger.log('נמענים שנותרו היום: ' + MailApp.getRemainingDailyQuota());
}

// ---------- כלי עזר ----------
// הגיליון: לפי המזהה שנשמר ב-setupSheet (עובד גם כשהקוד רץ כאפליקציית אינטרנט)
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('צריך להריץ פעם אחת את setupSheet מהעורך');
  return SpreadsheetApp.openById(id);
}

function readTable_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('חסרה לשונית: ' + name);
  var values = sh.getDataRange().getValues();
  var head = values.shift() || [];
  var col = {};
  head.forEach(function (h, i) { col[String(h).trim()] = i; });
  return { sheet: sh, head: head, col: col, rows: values };
}

function getSettings_() {
  var t = readTable_(SH.SETTINGS), s = {};
  t.rows.forEach(function (r) { if (r[0] !== '') s[String(r[0]).trim()] = r[1]; });
  return s;
}

function isYes_(v) { return v === true || String(v).trim() === 'כן'; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- קטלוג ----------
function getCatalog_() {
  var t = readTable_(SH.PRODUCTS), c = t.col, s = getSettings_();
  var products = [];
  t.rows.forEach(function (r) {
    if (!r[c['מזהה']] || !isYes_(r[c['פעיל']])) return;
    var unlimited = isYes_(r[c['ללא הגבלה']]);
    products.push({
      id: String(r[c['מזהה']]),
      cat: r[c['קטגוריה']], catLabel: r[c['שם קטגוריה']],
      code: String(r[c['קוד']]), sku: String(r[c['מק"ט']]),
      name: r[c['שם']], desc: r[c['תיאור מורחב']], color: r[c['צבע']],
      price: Number(r[c['מחיר']]) || 0,
      stock: unlimited ? null : Math.max(0, Number(r[c['יתרה']]) || 0),
      img: r[c['תמונה']], kind: r[c['סוג']] || 'item',
      tone: r[c['גוון']], diameter: r[c['קוטר']] === '' ? '' : Number(r[c['קוטר']]),
      gift: r[c['מתנה']], order: Number(r[c['סדר']]) || 0
    });
  });
  products.sort(function (a, b) { return a.order - b.order; });
  return {
    ok: true,
    budget: Number(s['תקציב']) || 9500,
    program: s['שם התוכנית'] || '',
    open: isYes_(s['פתוח להזמנות']),
    products: products,
    time: new Date().toISOString()
  };
}

// ---------- נקודות כניסה ----------
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'catalog';
  try {
    if (action === 'catalog') return json_(getCatalog_());
    return json_({ ok: false, error: 'פעולה לא מוכרת: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    var p = parsePostBody(e);
    // פעולות ההזמנה והניהול יתווספו בשלבים 4-5
    return json_({ ok: false, error: 'פעולה לא מוכרת: ' + (p.action || '') });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * קריאת גוף הבקשה. עוקף את מגבלת 8190 התווים של e.parameter - לא להסיר.
 * תומך גם ב-JSON (נשלח כ-Content-Type: text/plain כדי שהתשובה תהיה קריאה בדפדפן)
 * וגם בטופס רגיל (key=value&...).
 */
function parsePostBody(e) {
  var p = {};
  if (e && e.parameter) {
    for (var k in e.parameter) { p[k] = e.parameter[k]; }
  }
  if (e && e.postData && e.postData.contents) {
    var body = e.postData.contents;
    var trimmed = body.replace(/^\s+/, '');
    if (trimmed.charAt(0) === '{') {
      var obj = JSON.parse(trimmed);
      for (var j in obj) { p[j] = obj[j]; }
      return p;
    }
    body.split('&').forEach(function (pair) {
      var idx = pair.indexOf('=');
      if (idx > 0) {
        try {
          var key = decodeURIComponent(pair.substring(0, idx).replace(/\+/g, ' '));
          var val = decodeURIComponent(pair.substring(idx + 1).replace(/\+/g, ' '));
          p[key] = val;
        } catch (ex) {}
      }
    });
  }
  return p;
}
