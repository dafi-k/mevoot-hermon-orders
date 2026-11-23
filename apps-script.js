/**
 * אתר הזמנות - נבחרת צימרי מבואות החרמון
 * Google Apps Script - מחובר לגיליון (תוספים > Apps Script מתוך הגיליון)
 *
 * חשוב: אחרי כל שינוי בקוד הזה צריך פריסה חדשה:
 * פריסה > ניהול פריסות > עיפרון > גרסה: גרסה חדשה > פריסה
 *
 * שלב 2: הקמת הגיליון + קריאת קטלוג ומלאי (doGet?action=catalog)
 * שלב 4: שליחת הזמנה עם בדיקת מלאי בתוך נעילה (submitOrder) + מייל בבקשה נפרדת (sendOrderEmail)
 *
 * עקרונות:
 * - מהאתר מתקבלים רק מזהים וכמויות. מחירים, סה"כ וחריגה מחושבים כאן, לפי הגיליון בלבד.
 * - בדיקת המלאי וכתיבת ההזמנה נעשות בתוך LockService, כך ששני לקוחות לא יקבלו את אותה יחידה אחרונה.
 * - אם חסר מלאי - שום דבר לא נכתב.
 * - כשל במייל לא מבטל הזמנה. הסטטוס נרשם בעמודה "סטטוס מייל".
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
  ['פתוח להזמנות', 'לא', 'כן / לא. כש"לא" - רק הזמנות בדיקה עם קוד בדיקה'],
  ['קוד בדיקה', '', 'לשליחת הזמנות בדיקה כשהאתר סגור: לפתוח את האתר עם ?test=הקוד'],
  ['תיקיית קבצים', '', 'מזהה תיקיית Drive לקבצי ה-Excel של ההזמנות (נוצרת אוטומטית)']
];

var HEADER_COLOR = '#5C6B3A'; // ירוק זית
var TZ = 'Asia/Jerusalem';
var DAYS = ['1.12.26', '2.12.26', '3.12.26'];
var ST = { NEW: 'חדשה', APPROVED: 'אושרה', SUPPLIED: 'סופקה', CANCELLED: 'בוטלה' };
var MAIL = { PENDING: 'ממתין', SENT: 'נשלח', FAILED: 'לא נשלח' };
// כתובת האתר הציבורי - ממנה נלקחות תמונות מוקטנות (pic/thumbs) למייל ול-Excel שנבנה בשרת
var SITE_URL = 'https://dafi-k.github.io/mevoot-hermon-orders/';
var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var WATCHDOG_MINUTES = 2; // אחרי כמה דקות ב"ממתין" השרת שולח את המייל בעצמו

// =====================================================================
// הקמה (להריץ מהעורך - בטוח להריץ שוב, לא משכפל ולא מוחק נתונים)
// =====================================================================
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // שומרים את מזהה הגיליון, כדי שאפליקציית האינטרנט תפתח אותו לפי מזהה (openById)
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  SpreadsheetApp.openById(ss.getId()); // מבקש את הרשאת הגיליונות המלאה כבר עכשיו
  ss.setSpreadsheetLocale('iw_IL');
  [SH.PRODUCTS, SH.ORDERS, SH.ITEMS, SH.MOVES].forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    ensureHeaders_(sh, HEADERS[name]);
    styleSheet_(sh);
    // עמודת התאריך: תאריך ושעה, כולל שורות קיימות
    var di = HEADERS[name].indexOf('תאריך');
    if (di >= 0) sh.getRange(2, di + 1, sh.getMaxRows() - 1, 1).setNumberFormat(DATE_FMT);
  });
  setupSummaryTabs_(ss);
  // טלפון ומספר הזמנה כטקסט (שלא ייעלם האפס בהתחלה)
  var o = ss.getSheetByName(SH.ORDERS);
  o.getRange('A:A').setNumberFormat('@');
  o.getRange(1, HEADERS[SH.ORDERS].indexOf('טלפון') + 1, o.getMaxRows(), 1).setNumberFormat('@');

  var st = ss.getSheetByName(SH.SETTINGS) || ss.insertSheet(SH.SETTINGS);
  if (st.getLastRow() === 0) st.getRange(1, 1, 1, 3).setValues([['הגדרה', 'ערך', 'הסבר']]);
  ensureSettings_(st);
  styleSheet_(st);
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
  // קוד בדיקה ותיקיית קבצים
  var s = getSettings_();
  if (!String(s['קוד בדיקה'] || '').trim()) setSetting_('קוד בדיקה', 'test' + Math.floor(100000 + Math.random() * 900000));
  getFilesFolder_(); // יוצר את התיקייה ב-Drive אם צריך (ומבקש הרשאת Drive)
  s = getSettings_();
  Logger.log('הגיליון מוכן. מוצרים: ' + p.rows.length);
  Logger.log('פתוח להזמנות: ' + s['פתוח להזמנות'] + ' | קוד בדיקה: ' + s['קוד בדיקה']);
  Logger.log('נמעני צוות: ' + (s['נמעני מייל צוות'] || '(ריק - צריך למלא בלשונית הגדרות)'));
}

// לשוניות מחושבות (נוסחאות QUERY) - בלי הזמנות שבוטלו ובלי הזמנות בדיקה. לא לערוך ידנית.
function setupSummaryTabs_(ss) {
  var totals = ss.getSheetByName(SH.TOTALS) || ss.insertSheet(SH.TOTALS);
  totals.clear();
  totals.getRange('A1').setFormula("=QUERY('" + SH.ITEMS + "'!A:M,\"select D, E, F, G, H, sum(J), sum(L) where D is not null and M <> '" + ST.CANCELLED +
    "' and not A starts with 'TEST' group by D, E, F, G, H order by sum(J) desc label D 'מזהה', E 'קוד', F 'מק\"\"ט', G 'פריט', H 'צבע', sum(J) 'כמות שהוזמנה', sum(L) 'סכום ₪'\",1)");
  var byz = ss.getSheetByName(SH.BY_ZIMMER) || ss.insertSheet(SH.BY_ZIMMER);
  byz.clear();
  byz.getRange('A1').setFormula("=QUERY('" + SH.ORDERS + "'!A:Q,\"select C, D, E, F, A, B, J, K, K-J, L, N where A is not null and N <> '" + ST.CANCELLED +
    "' and not A starts with 'TEST' order by C label C 'שם הצימר', D 'ישוב', E 'איש קשר', F 'טלפון', A 'מספר הזמנה', B 'תאריך', J 'סה\"\"כ', K 'תקציב', K-J 'יתרה', L 'חריגה', N 'סטטוס' format B 'dd/mm/yyyy hh:mm'\",1)");
  [totals, byz].forEach(function (sh) {
    sh.setRightToLeft(true); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 11).setBackground(HEADER_COLOR).setFontColor('#ffffff').setFontWeight('bold');
  });
}

function ensureHeaders_(sh, headers) {
  var lc = Math.max(sh.getLastColumn(), 1);
  var cur = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (h) { return String(h).trim(); });
  headers.forEach(function (h) {
    if (cur.indexOf(h) < 0) { cur.push(h); sh.getRange(1, cur.length).setValue(h); }
  });
}

function ensureSettings_(st) {
  var vals = st.getDataRange().getValues();
  var have = vals.map(function (r) { return String(r[0]).trim(); });
  DEFAULT_SETTINGS.forEach(function (d) {
    if (have.indexOf(d[0]) < 0) st.appendRow(d);
  });
}

function styleSheet_(sh) {
  sh.setRightToLeft(true);
  sh.setFrozenRows(1);
  var lc = Math.max(sh.getLastColumn(), 1);
  sh.getRange(1, 1, 1, lc).setBackground(HEADER_COLOR).setFontColor('#ffffff').setFontWeight('bold');
}

// בדיקת החשבון והמכסה היומית של מיילים (להריץ מהעורך ולקרוא ביומן הביצוע)
function checkAccount() {
  var email = Session.getEffectiveUser().getEmail();
  var isGmail = /@gmail\.com$/i.test(email);
  Logger.log('הסקריפט רץ מהחשבון: ' + email);
  Logger.log('סוג חשבון: ' + (isGmail ? 'Gmail רגיל - עד 100 נמענים ביום' : 'Workspace - עד 1,500 נמענים ביום'));
  Logger.log('נמענים שנותרו היום: ' + MailApp.getRemainingDailyQuota());
  Logger.log('מיילים של הזמנות שנשלחו היום: ' + mailCountToday_());
}

// =====================================================================
// כלי עזר
// =====================================================================
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

// שורה חדשה לפי שמות העמודות בגיליון (לא תלוי בסדר העמודות)
function rowFor_(t, obj) {
  var row = t.head.map(function () { return ''; });
  for (var k in obj) { if (t.col[k] !== undefined) row[t.col[k]] = obj[k]; }
  return row;
}

function appendRows_(t, rows) {
  if (!rows.length) return;
  var start = t.sheet.getLastRow() + 1;
  t.sheet.getRange(start, 1, rows.length, t.head.length).setValues(rows);
  // תאריך ושעה (לא רק תאריך)
  if (t.col['תאריך'] !== undefined) t.sheet.getRange(start, t.col['תאריך'] + 1, rows.length, 1).setNumberFormat(DATE_FMT);
}
var DATE_FMT = 'dd/MM/yyyy HH:mm';

function getSettings_() {
  var t = readTable_(SH.SETTINGS), s = {};
  t.rows.forEach(function (r) { if (r[0] !== '') s[String(r[0]).trim()] = r[1]; });
  return s;
}

function setSetting_(key, value) {
  var t = readTable_(SH.SETTINGS);
  for (var i = 0; i < t.rows.length; i++) {
    if (String(t.rows[i][0]).trim() === key) { t.sheet.getRange(i + 2, 2).setValue(value); return; }
  }
  t.sheet.appendRow([key, value, '']);
}

function isYes_(v) { return v === true || String(v).trim() === 'כן'; }
function normPhone_(v) { var d = String(v || '').replace(/\D/g, ''); return d.length === 9 ? '0' + d : d; }
function money_(n) { return '₪' + Math.round(n).toLocaleString('en-US'); }
function esc_(s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function now_() { return new Date(); }
function fmtDate_(d) { return Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// קטלוג
// =====================================================================
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

// מפת מוצרים לפי מזהה: { id: {i (אינדקס שורה), row, ...} }
function productMap_(t) {
  var c = t.col, map = {};
  t.rows.forEach(function (r, i) {
    var id = String(r[c['מזהה']] || '').trim();
    if (!id) return;
    map[id] = {
      i: i, id: id, active: isYes_(r[c['פעיל']]), unlimited: isYes_(r[c['ללא הגבלה']]),
      remaining: Number(r[c['יתרה']]) || 0, price: Number(r[c['מחיר']]) || 0,
      code: String(r[c['קוד']]), sku: String(r[c['מק"ט']]), name: String(r[c['שם']]),
      color: String(r[c['צבע']] || ''), kind: String(r[c['סוג']] || 'item'), tone: String(r[c['גוון']] || ''),
      diameter: r[c['קוטר']], gift: String(r[c['מתנה']] || ''), img: String(r[c['תמונה']] || '')
    };
  });
  return map;
}

// =====================================================================
// שליחת הזמנה
// =====================================================================
function cleanCustomer_(c) {
  c = c || {};
  var out = {};
  ['zimmer', 'town', 'contact', 'phone', 'email', 'address', 'day'].forEach(function (k) {
    out[k] = String(c[k] === undefined || c[k] === null ? '' : c[k]).trim().slice(0, 200);
  });
  out.phone = normPhone_(out.phone);
  if (out.zimmer.length < 2) return { error: 'חסר שם הצימר' };
  if (out.town.length < 2) return { error: 'חסר ישוב' };
  if (out.contact.length < 2) return { error: 'חסר איש קשר' };
  if (!/^05\d{8}$/.test(out.phone)) return { error: 'מספר הטלפון לא תקין' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(out.email)) return { error: 'כתובת המייל לא תקינה' };
  if (out.address.length < 5) return { error: 'חסרה כתובת לאספקה' };
  if (DAYS.indexOf(out.day) < 0) return { error: 'לא נבחר יום הגעה לאולם' };
  return out;
}

// שורות מהאתר: [{type:'item', id, qty}] או [{type:'table', leg, top, qty}] - רק מזהים וכמויות
function cleanLines_(lines) {
  if (!Array.isArray(lines) || !lines.length) return { error: 'העגלה ריקה' };
  if (lines.length > 200) return { error: 'יותר מדי שורות בהזמנה' };
  var out = [];
  for (var i = 0; i < lines.length; i++) {
    var l = lines[i] || {};
    var qty = Number(l.qty);
    if (!(qty >= 1 && qty <= 999 && Math.floor(qty) === qty)) return { error: 'כמות לא תקינה' };
    if (l.type === 'table') out.push({ type: 'table', leg: String(l.leg || ''), top: String(l.top || ''), qty: qty, lid: l.lid });
    else out.push({ type: 'item', id: String(l.id || ''), qty: qty, lid: l.lid });
  }
  return { lines: out };
}

function submitOrder_(p) {
  var s = getSettings_();
  var testCode = String(s['קוד בדיקה'] || '').trim();
  var isTest = !!testCode && String(p.testCode || '').trim() === testCode;
  if (!isYes_(s['פתוח להזמנות']) && !isTest) {
    return { ok: false, code: 'closed', error: 'שליחת הזמנות סגורה כרגע.' };
  }
  var c = cleanCustomer_(p.customer);
  if (c.error) return { ok: false, code: 'invalid', error: c.error };
  var cl = cleanLines_(p.lines);
  if (cl.error) return { ok: false, code: 'invalid', error: cl.error };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    return { ok: false, code: 'busy', error: 'המערכת עמוסה, נסו שוב בעוד רגע.' };
  }
  try {
    // 1. הזמנה אחת לטלפון (הזמנות שבוטלו לא נספרות)
    var o = readTable_(SH.ORDERS), oc = o.col;
    var dup = o.rows.some(function (r) {
      return normPhone_(r[oc['טלפון']]) === c.phone && String(r[oc['סטטוס']]).trim() !== ST.CANCELLED;
    });
    if (dup) return { ok: false, code: 'duplicate', error: 'כבר נשלחה הזמנה ממספר הטלפון הזה. לשינויים יש לפנות לצוות.' };

    // 2. מלאי ומחירים מהגיליון
    var t = readTable_(SH.PRODUCTS), pm = productMap_(t);
    var need = {};
    var add = function (id, q) { need[id] = (need[id] || 0) + q; };
    cl.lines.forEach(function (l) {
      if (l.type === 'table') { add(l.leg, l.qty); add(l.top, l.qty); } else add(l.id, l.qty);
    });
    var shortages = [], stock = {};
    Object.keys(need).forEach(function (id) {
      var P = pm[id];
      var bad = !P || !P.active ||
        (cl.lines.some(function (l) { return l.type === 'table' && ((l.leg === id && P.kind !== 'leg') || (l.top === id && P.kind !== 'top')); }));
      if (bad) { shortages.push({ id: id, requested: need[id], available: 0 }); stock[id] = 0; return; }
      if (P.unlimited) { stock[id] = null; return; }
      stock[id] = Math.max(0, P.remaining);
      if (need[id] > P.remaining) shortages.push({ id: id, requested: need[id], available: Math.max(0, P.remaining) });
    });
    if (shortages.length) return { ok: false, code: 'shortage', error: 'חלק מהפריטים כבר לא זמינים בכמות שביקשת.', shortages: shortages, stock: stock };

    // 3. חישוב מחירים - רק מהגיליון
    var out = [], total = 0;
    cl.lines.forEach(function (l) {
      var x;
      if (l.type === 'table') { x = tableLine_(pm[l.leg], pm[l.top], l.qty); x.leg = pm[l.leg]; x.top = pm[l.top]; }
      else { x = itemLine_(pm[l.id], l.qty, pm[l.id].price, pm[l.id].gift ? 'במתנה: ' + pm[l.id].gift : ''); x.p = pm[l.id]; }
      x.lid = l.lid;
      out.push(x);
      total += x.subtotal;
    });
    var budget = Number(s['תקציב']) || 9500;
    var over = Math.max(0, total - budget);
    if (over > 0 && p.overOk !== true) {
      return { ok: false, code: 'overConfirm', error: 'ההזמנה מעל התקציב. יש לסמן אישור לתשלום ההפרש.', total: total, over: over };
    }

    // 4. כתיבה
    var now = now_();
    var seq = ('000' + (o.rows.length + 1)).slice(-3);
    var orderId = (isTest ? 'TEST-' : 'MH-') + Utilities.formatDate(now, TZ, 'yyMMdd') + '-' + seq;
    appendRows_(o, [rowFor_(o, {
      'מספר הזמנה': orderId, 'תאריך': now, 'שם הצימר': c.zimmer, 'ישוב': c.town, 'איש קשר': c.contact,
      'טלפון': c.phone, 'מייל': c.email, 'כתובת לאספקה': c.address, 'יום הגעה לאולם': c.day,
      'סה"כ': total, 'תקציב': budget, 'חריגה': over, 'אישור חריגה': over > 0 ? 'כן' : '',
      'סטטוס': ST.NEW, 'סטטוס מייל': MAIL.PENDING, 'הערות צוות': isTest ? 'הזמנת בדיקה' : ''
    })]);

    var it = readTable_(SH.ITEMS), itemRows = [];
    var itemRow = function (P, qty, price, detail) {
      return rowFor_(it, {
        'מספר הזמנה': orderId, 'תאריך': now, 'שם הצימר': c.zimmer, 'מזהה': P.id, 'קוד': P.code, 'מק"ט': P.sku,
        'שם': P.name, 'צבע': P.color, 'פירוט': detail, 'כמות': qty, 'מחיר יחידה': price, 'סה"כ שורה': price * qty,
        'סטטוס הזמנה': ST.NEW
      });
    };
    out.forEach(function (l, n) {
      if (l.kind === 'table') {
        // שולחן = 2 שורות (רגל + פלטה) עם אותו "פירוט", כדי שהמלאי והריכוזים יהיו לפי מק"ט
        itemRows.push(itemRow(l.leg, l.qty, l.leg.price, l.name));
        itemRows.push(itemRow(l.top, l.qty, l.top.price, l.name));
      } else itemRows.push(itemRow(l.p, l.qty, l.price, l.detail));
    });
    appendRows_(it, itemRows);

    // 5. עדכון יתרות + יומן תנועות
    var mv = readTable_(SH.MOVES), moveRows = [], colRem = t.col['יתרה'];
    Object.keys(need).forEach(function (id) {
      var P = pm[id];
      if (P.unlimited) return;
      var after = P.remaining - need[id];
      t.sheet.getRange(P.i + 2, colRem + 1).setValue(after);
      moveRows.push(rowFor_(mv, { 'תאריך': now, 'מי': isTest ? 'אתר (בדיקה)' : 'אתר', 'מזהה': id, 'מק"ט': P.sku,
        'שינוי': -need[id], 'יתרה אחרי': after, 'סוג': 'הזמנה', 'הפניה': orderId, 'הערה': c.zimmer }));
    });
    appendRows_(mv, moveRows);
    SpreadsheetApp.flush();

    return {
      ok: true, orderId: orderId, test: isTest, date: fmtDate_(now), customer: c,
      total: total, budget: budget, over: over,
      lines: out.map(publicLine_)
    };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// שורות תצוגה של הזמנה (אתר, Excel, PDF, מייל, מסכי ניהול)
// שולחן מוצג כשורה אחת: תמונת רגל + תמונת פלטה, תיאור ברור, מחיר רגל ומחיר פלטה בנפרד וסה"כ לשולחן.
// =====================================================================
function tableTitle_(legCode, tone, dia) { return 'שולחן: רגל ' + legCode + ' + פלטה ' + tone + ' קוטר ' + dia; }
function tableDetail_(lg, tp) {
  return 'רגל ' + lg.code + ': ' + money_(lg.price) + ' · פלטה ' + tp.tone + ' ' + tp.diameter + ' ס"מ: ' + money_(tp.price) +
    ' · סה"כ לשולחן: ' + money_(lg.price + tp.price);
}
// lg/tp: {code, sku, price, img, tone, diameter}
function tableLine_(lg, tp, qty) {
  var price = lg.price + tp.price;
  return { kind: 'table', code: lg.code + ' + ' + tp.code, sku: lg.sku + ' + ' + tp.sku, name: tableTitle_(lg.code, tp.tone, tp.diameter),
    color: tp.tone, detail: tableDetail_(lg, tp), qty: qty, price: price, subtotal: price * qty,
    legCode: lg.code, legPrice: lg.price, topPrice: tp.price, imgs: [lg.img, tp.img].filter(String) };
}
function itemLine_(P, qty, price, detail) {
  return { kind: 'item', code: P.code, sku: P.sku, name: P.name, color: P.color, detail: detail || '', qty: qty, price: price,
    subtotal: price * qty, imgs: P.img ? [P.img] : [] };
}
function publicLine_(l) {
  return { kind: l.kind, lid: l.lid, code: l.code, sku: l.sku, name: l.name, color: l.color, detail: l.detail, qty: l.qty,
    price: l.price, subtotal: l.subtotal, imgs: l.imgs, legPrice: l.legPrice, topPrice: l.topPrice,
    img: l.imgs[l.imgs.length - 1] || '' };
}
// שורות הפריטים של הזמנה מהגיליון -> שורות תצוגה. רגל + פלטה עוקבות עם אותו "פירוט" של שולחן = שולחן אחד.
// המחירים נלקחים מהשורות שנשמרו בהזמנה (לא מהמחיר הנוכחי בגיליון המוצרים).
function orderLines_(orderId, pm, it) {
  it = it || readTable_(SH.ITEMS);
  var ic = it.col;
  var rows = it.rows.filter(function (x) { return String(x[ic['מספר הזמנה']]) === String(orderId); });
  var part = function (r) {
    var id = String(r[ic['מזהה']]), P = pm[id] || {};
    return { id: id, kind: P.kind || '', code: String(r[ic['קוד']]), sku: String(r[ic['מק"ט']]), name: String(r[ic['שם']]),
      color: String(r[ic['צבע']] || ''), price: Number(r[ic['מחיר יחידה']]) || 0, qty: Number(r[ic['כמות']]) || 0,
      detail: String(r[ic['פירוט']] || ''), img: P.img || '', tone: P.tone || String(r[ic['צבע']] || ''), diameter: P.diameter };
  };
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var a = part(rows[i]), b = rows[i + 1] ? part(rows[i + 1]) : null;
    if (a.kind === 'leg' && b && b.kind === 'top' && /^שולחן/.test(a.detail) && a.detail === b.detail && a.qty === b.qty) {
      out.push(tableLine_(a, b, a.qty)); i++;
    } else out.push(itemLine_(a, a.qty, a.price, a.detail));
  }
  return out;
}
function thumbName_(img) { return String(img || '').replace(/\.[a-z0-9]+$/i, '') + '.jpg'; }
function thumbUrl_(img) { return SITE_URL + 'pic/thumbs/' + encodeURIComponent(thumbName_(img)); }

// =====================================================================
// מייל (בקשה נפרדת - כשל כאן לא מבטל את ההזמנה)
// =====================================================================
// "תפיסת" שליחה של הזמנה - כדי שהאתר, ניסיון חוזר והגיבוי האוטומטי לא ישלחו אותו מייל פעמיים במקביל
function claimMail_(orderId) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return false;
  try {
    var cache = CacheService.getScriptCache(), k = 'mailclaim-' + orderId;
    if (cache.get(k)) return false;
    cache.put(k, '1', 300);
    return true;
  } finally { lock.releaseLock(); }
}
function releaseMail_(orderId) { CacheService.getScriptCache().remove('mailclaim-' + orderId); }
function findOrder_(orderId) {
  var o = readTable_(SH.ORDERS), c = o.col;
  for (var i = 0; i < o.rows.length; i++) {
    if (String(o.rows[i][c['מספר הזמנה']]) === String(orderId)) return { t: o, i: i, row: o.rows[i], c: c };
  }
  return null;
}

function setOrderField_(f, field, value) {
  f.t.sheet.getRange(f.i + 2, f.c[field] + 1).setValue(value);
}

function getFilesFolder_() {
  var s = getSettings_(), id = String(s['תיקיית קבצים'] || '').trim();
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var folder = DriveApp.createFolder('הזמנות מבואות החרמון - קבצים');
  setSetting_('תיקיית קבצים', folder.getId());
  return folder;
}

function sendOrderEmail_(p) {
  var f = findOrder_(p.orderId);
  if (!f) return { ok: false, error: 'הזמנה לא נמצאה' };
  if (normPhone_(f.row[f.c['טלפון']]) !== normPhone_(p.phone)) return { ok: false, error: 'פרטי ההזמנה לא תואמים' };
  if (String(f.row[f.c['סטטוס מייל']]).indexOf(MAIL.SENT) === 0) return { ok: true, emailed: true, already: true };
  // שליחה אחרת של אותה הזמנה כבר רצה (ניסיון קודם שהתשובה שלו לא הגיעה, או הגיבוי האוטומטי)
  if (!claimMail_(p.orderId)) return { ok: true, emailed: false, pending: true };
  try {
    // שמירת ה-Excel מהאתר ב-Drive (רק אם עוד אין קובץ להזמנה - ניסיון חוזר לא יוצר כפילות)
    var fileId = '';
    if (!String(f.row[f.c['קובץ Excel']] || '')) {
      try {
        var b64 = String(p.xlsxBase64 || '');
        if (b64 && b64.length < 8 * 1024 * 1024) {
          var name = String(p.xlsxName || (p.orderId + '.xlsx')).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 120);
          var file = getFilesFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), XLSX_MIME, name));
          fileId = file.getId();
          setOrderField_(f, 'קובץ Excel', file.getUrl());
        }
      } catch (e) { /* אם אין קובץ מהאתר - השרת יבנה Excel בעצמו */ }
    }
    var res = sendEmailForOrder_(p.orderId, fileId);
    return { ok: true, emailed: res.ok, error: res.error || '' };
  } finally { releaseMail_(p.orderId); }
}

// קובץ ה-Excel של ההזמנה ב-Drive. אם אין (למשל האתר לא הספיק להעלות) - השרת בונה אותו מהגיליון ושומר.
function ensureOrderXlsx_(f, orderId) {
  var url = String(f.row[f.c['קובץ Excel']] || '');
  var m = url.match(/[-\w]{25,}/);
  if (m) { try { return DriveApp.getFileById(m[0]); } catch (e) { /* הקובץ נמחק - בונים מחדש */ } }
  var file = getFilesFolder_().createFile(buildOrderXlsx_(orderId));
  setOrderField_(f, 'קובץ Excel', file.getUrl());
  f.row[f.c['קובץ Excel']] = file.getUrl();
  return file;
}

// שולח את מייל ההזמנה ומעדכן "סטטוס מייל". משמש גם לשליחה חוזרת ולגיבוי האוטומטי.
function sendEmailForOrder_(orderId, fileId, via) {
  var f = findOrder_(orderId);
  if (!f) return { ok: false, error: 'הזמנה לא נמצאה' };
  var c = f.c, r = f.row, s = getSettings_();
  try {
    var staff = String(s['נמעני מייל צוות'] || '').split(/[,;\s]+/).filter(function (x) { return /@/.test(x); });
    var customer = String(r[c['מייל']] || '').trim();
    var recipients = staff.length + (customer ? 1 : 0);
    if (!recipients) throw new Error('אין נמענים');
    if (MailApp.getRemainingDailyQuota() < recipients) throw new Error('מכסת המיילים היומית נגמרה');

    // Excel מצורף: הקובץ מהאתר, ואם אין - קובץ שהשרת בונה מהגיליון
    var attachments = [], xlsxNote = '';
    try {
      var file = fileId ? DriveApp.getFileById(fileId) : ensureOrderXlsx_(f, orderId);
      attachments.push(file.getBlob());
    } catch (e) { xlsxNote = ' (בלי Excel: ' + String(e && e.message || e).slice(0, 80) + ')'; }

    var total = Number(r[c['סה"כ']]) || 0, budget = Number(r[c['תקציב']]) || 0, over = Number(r[c['חריגה']]) || 0;
    var isTest = String(orderId).indexOf('TEST-') === 0;
    var subject = (isTest ? '[בדיקה] ' : '') + 'הזמנת ריהוט ' + orderId + ' - ' + r[c['שם הצימר']] + ' (' + r[c['ישוב']] + ')';
    var inline = {}; // תמונות מוקטנות בתוך המייל (cid:) - Gmail לא חוסם אותן כמו קישור חיצוני
    var html = buildEmailHtml_(orderId, f, total, budget, over, isTest, inline);
    var opts = { subject: subject, htmlBody: html, name: s['שם התוכנית'] || 'נבחרת צימרי מבואות החרמון', attachments: attachments };
    if (Object.keys(inline).length) opts.inlineImages = inline;
    if (staff.length) { opts.to = staff.join(','); if (customer) opts.cc = customer; opts.replyTo = staff[0]; }
    else opts.to = customer;
    MailApp.sendEmail(opts);
    bumpMailCount_(recipients);
    var note = MAIL.SENT + ' ' + fmtDate_(now_()) + (staff.length ? '' : ' (ללקוח בלבד - אין נמעני צוות בהגדרות)') +
      (via ? ' (' + via + ')' : '') + xlsxNote;
    setOrderField_(f, 'סטטוס מייל', note);
    return { ok: true, attached: attachments.length > 0 };
  } catch (err) {
    var msg = String(err && err.message || err);
    setOrderField_(f, 'סטטוס מייל', MAIL.FAILED + ': ' + msg);
    return { ok: false, error: msg };
  }
}

// inline: אובייקט שמתמלא בתמונות המוקטנות { img0: blob, ... } ל-MailApp inlineImages
function buildEmailHtml_(orderId, f, total, budget, over, isTest, inline) {
  var c = f.c, r = f.row;
  var lines = orderLines_(orderId, productMap_(readTable_(SH.PRODUCTS)));
  // טעינת התמונות המוקטנות (JPEG 160×120, כמה KB כל אחת) בשרת. תמונה שחוזרת (למשל אותה רגל) מצורפת פעם אחת.
  var files = [];
  lines.forEach(function (l) { l.imgs.forEach(function (im) { if (files.indexOf(im) < 0) files.push(im); }); });
  var thumbs = fetchThumbs_(files), cid = {};
  files.forEach(function (im, i) {
    if (!thumbs[im] || !inline) return;
    cid[im] = 'img' + i;
    inline[cid[im]] = Utilities.newBlob(thumbs[im], 'image/jpeg', thumbName_(im));
  });
  var td = 'padding:8px 10px;border-bottom:1px solid #E6DDC8;text-align:right;vertical-align:middle;';
  lines = lines.map(function (l) {
    // שולחן: רגל ואז פלטה. תמונה שלא נטענה פשוט לא מוצגת (בלי קישור חיצוני)
    var pics = l.imgs.filter(function (im) { return cid[im]; }).map(function (im) {
      return '<img src="cid:' + cid[im] + '" width="64" height="48" alt="" style="width:64px;height:48px;border:1px solid #E6DDC8;border-radius:6px;margin:0 0 0 4px;vertical-align:middle">';
    }).join('');
    return '<tr><td style="' + td + 'white-space:nowrap">' + pics + '</td><td style="' + td + '"><b>' + esc_(l.name) + '</b>' +
      (l.kind === 'item' && l.color ? ' · ' + esc_(l.color) : '') +
      '<br><span style="color:#6E6A5E;font-size:12px">' + esc_(l.code) + (l.detail ? ' · ' + esc_(l.detail) : '') + '</span>' +
      '</td><td style="' + td + 'text-align:center">' + esc_(l.qty) + '</td><td style="' + td + '">' + money_(l.price) +
      '</td><td style="' + td + 'font-weight:bold">' + money_(l.subtotal) + '</td></tr>';
  }).join('');
  var info = function (k, v) { return '<tr><td style="padding:3px 0;color:#6E6A5E;width:130px">' + k + '</td><td style="padding:3px 0"><b>' + esc_(v) + '</b></td></tr>'; };
  return '<div dir="rtl" style="font-family:Arial,sans-serif;color:#2F3A26;max-width:680px">' +
    (isTest ? '<div style="background:#F8E4DA;color:#B23A2E;padding:10px;border-radius:8px;margin-bottom:10px;font-weight:bold">הזמנת בדיקה - לא להעביר לביצוע</div>' : '') +
    '<div style="background:#5C6B3A;color:#fff;padding:18px;border-radius:10px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold">נבחרת צימרי מבואות החרמון</div><div style="margin-top:4px">אישור הזמנת ריהוט ' + esc_(orderId) + '</div></div>' +
    '<table style="margin:16px 0;font-size:14px">' +
    info('צימר:', r[c['שם הצימר']] + ', ' + r[c['ישוב']]) + info('איש קשר:', r[c['איש קשר']]) + info('טלפון:', r[c['טלפון']]) +
    info('מייל:', r[c['מייל']]) + info('כתובת לאספקה:', r[c['כתובת לאספקה']]) + info('הגעה לאולם:', r[c['יום הגעה לאולם']]) +
    info('תאריך הזמנה:', r[c['תאריך']] instanceof Date ? fmtDate_(r[c['תאריך']]) : r[c['תאריך']]) + '</table>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;border:1px solid #E6DDC8">' +
    '<tr style="background:#5C6B3A;color:#fff"><th style="padding:8px 10px"></th><th style="padding:8px 10px;text-align:right">פריט</th><th style="padding:8px 10px">כמות</th><th style="padding:8px 10px;text-align:right">מחיר</th><th style="padding:8px 10px;text-align:right">סה"כ</th></tr>' +
    lines + '</table>' +
    '<div style="margin-top:14px;padding:12px 16px;border:1px solid #E6DDC8;border-radius:10px;font-size:15px">' +
    '<div><b>סה"כ הזמנה: ' + money_(total) + '</b></div><div style="color:#6E6A5E">תקציב: ' + money_(budget) + '</div>' +
    (over > 0 ? '<div style="color:#B23A2E;font-weight:bold;margin-top:4px">מעל התקציב - לתשלום על ידי בעל/ת הצימר: ' + money_(over) + '</div>'
      : '<div style="color:#5C6B3A;margin-top:4px">יתרת תקציב: ' + money_(budget - total) + '</div>') + '</div>' +
    '<p style="font-size:13px;color:#6E6A5E;margin-top:16px">הובלה והרכבה כלולות בתוכנית. לשינויים בהזמנה יש לפנות לצוות.</p></div>';
}

function mailKey_() { return 'mail-' + Utilities.formatDate(now_(), TZ, 'yyyyMMdd'); }
function mailCountToday_() { return Number(PropertiesService.getScriptProperties().getProperty(mailKey_()) || 0); }
function bumpMailCount_(n) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(mailKey_(), String(mailCountToday_() + n));
}

// שליחה חוזרת של כל המיילים שלא נשלחו (אפשר להריץ מהעורך; במסכי הניהול יש כפתור "שלח שוב")
function resendFailedEmails() {
  var o = readTable_(SH.ORDERS), c = o.col, n = 0, ok = 0;
  o.rows.forEach(function (r) {
    var st = String(r[c['סטטוס מייל']] || ''), id = String(r[c['מספר הזמנה']]);
    if (!id || st.indexOf(MAIL.SENT) === 0 || String(r[c['סטטוס']]) === ST.CANCELLED) return;
    if (!claimMail_(id)) return;
    n++;
    try {
      var res = sendEmailForOrder_(id, '');
      if (res.ok) ok++;
      Logger.log(id + ': ' + (res.ok ? 'נשלח' : 'נכשל - ' + res.error));
    } finally { releaseMail_(id); }
  });
  Logger.log('נשלחו ' + ok + ' מתוך ' + n);
}

// =====================================================================
// גיבוי אוטומטי: הזמנה שנשארה "ממתין" יותר מ-2 דקות (למשל הדפדפן נסגר או נקטע אחרי השמירה)
// - השרת בונה את ה-Excel מהגיליון ושולח את המייל בעצמו. רץ כל דקה (טריגר).
// =====================================================================
// להריץ פעם אחת מהעורך: מתקין את הטריגר (ומבקש את ההרשאות החדשות). בטוח להריץ שוב.
function installMailWatchdog() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'mailWatchdog') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mailWatchdog').timeBased().everyMinutes(1).create();
  // רק הזמנות מעכשיו והלאה (הזמנות בדיקה ישנות שנשארו "ממתין" לא יישלחו פתאום)
  PropertiesService.getScriptProperties().setProperty('WATCHDOG_SINCE', String(now_().getTime()));
  UrlFetchApp.fetch(SITE_URL + 'pic/thumbs/LG-LP1.jpg', { muteHttpExceptions: true }); // מבקש הרשאת גישה לאתר (תמונות ל-Excel)
  Logger.log('הגיבוי האוטומטי פועל: כל דקה נבדקות הזמנות שנשארו "ממתין" יותר מ-' + WATCHDOG_MINUTES + ' דקות.');
}

function mailWatchdog() {
  var since = Number(PropertiesService.getScriptProperties().getProperty('WATCHDOG_SINCE') || 0);
  var o = readTable_(SH.ORDERS), c = o.col, nowMs = now_().getTime(), sent = 0;
  for (var i = 0; i < o.rows.length && sent < 10; i++) {
    var r = o.rows[i], id = String(r[c['מספר הזמנה']] || '');
    if (!id || String(r[c['סטטוס מייל']] || '').trim() !== MAIL.PENDING || String(r[c['סטטוס']]) === ST.CANCELLED) continue;
    var d = r[c['תאריך']], t = d instanceof Date ? d.getTime() : 0;
    if (!t || t < since || nowMs - t < WATCHDOG_MINUTES * 60 * 1000) continue;
    if (!claimMail_(id)) continue;
    try {
      var res = sendEmailForOrder_(id, '', 'גיבוי אוטומטי');
      sent++;
      Logger.log(id + ': ' + (res.ok ? 'נשלח (גיבוי אוטומטי)' : 'נכשל - ' + res.error));
    } finally { releaseMail_(id); }
  }
}

// =====================================================================
// Excel שנבנה בשרת מהגיליון (כשהאתר לא העלה קובץ) - קובץ xlsx אמיתי, עם תמונות מוקטנות
// =====================================================================
function xmlEsc_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; });
}
function textBlob_(str, name, type) {
  var b = Utilities.newBlob('', type || 'application/xml', name);
  b.setDataFromString(str, 'UTF-8');
  return b;
}
// תמונות מוקטנות (JPEG) מהאתר: { 'LG-LP1.webp': bytes }
function fetchThumbs_(files) {
  var out = {};
  if (!files.length) return out;
  try {
    var res = UrlFetchApp.fetchAll(files.map(function (f) { return { url: thumbUrl_(f), muteHttpExceptions: true }; }));
    res.forEach(function (x, i) { if (x.getResponseCode() === 200) out[files[i]] = x.getContent(); });
  } catch (e) { /* בלי תמונות - הקובץ עדיין נבנה */ }
  return out;
}

function buildOrderXlsx_(orderId) {
  var f = findOrder_(orderId);
  if (!f) throw new Error('הזמנה לא נמצאה');
  var r = f.row, c = f.c;
  var lines = orderLines_(orderId, productMap_(readTable_(SH.PRODUCTS)));
  var files = [];
  lines.forEach(function (l) { l.imgs.forEach(function (im) { if (files.indexOf(im) < 0) files.push(im); }); });
  var thumbs = fetchThumbs_(files);
  var date = r[c['תאריך']] instanceof Date ? fmtDate_(r[c['תאריך']]) : String(r[c['תאריך']]);
  var order = {
    id: orderId, test: String(orderId).indexOf('TEST-') === 0,
    total: Number(r[c['סה"כ']]) || 0, budget: Number(r[c['תקציב']]) || 0, over: Number(r[c['חריגה']]) || 0,
    info: [['שם הצימר', r[c['שם הצימר']] + ', ' + r[c['ישוב']]], ['איש קשר', r[c['איש קשר']]], ['טלפון', r[c['טלפון']]],
      ['מייל', r[c['מייל']]], ['כתובת לאספקה', r[c['כתובת לאספקה']]], ['הגעה לאולם', r[c['יום הגעה לאולם']]], ['תאריך הזמנה', date]]
  };
  var parts = xlsxParts_(order, lines, thumbs);
  var blobs = parts.map(function (p) {
    return typeof p.data === 'string' ? textBlob_(p.data, p.name) : Utilities.newBlob(p.data, 'image/jpeg', p.name);
  });
  var name = ('הזמנה_' + orderId + '_' + r[c['שם הצימר']]).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80) + '.xlsx';
  var zip = Utilities.zip(blobs, name);
  zip.setContentType(XLSX_MIME);
  return zip;
}

// בונה את קבצי ה-XML של ה-xlsx (נבדק מקומית עם ExcelJS ב-tools/test-server.js)
function xlsxParts_(order, lines, thumbs) {
  var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  var NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var COLS = 'ABCDEFGHI';
  var rowsXml = [], merges = [], anchors = [], media = [], rowN = 0;
  var cell = function (ci, v, st) {
    var ref = COLS.charAt(ci) + rowN;
    if (typeof v === 'number') return '<c r="' + ref + '" s="' + st + '"><v>' + v + '</v></c>';
    return '<c r="' + ref + '" s="' + st + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc_(v) + '</t></is></c>';
  };
  var addRow = function (cells, ht) {
    rowN++;
    var xml = cells.map(function (x, i) { return x ? cell(i, x[0], x[1]) : ''; }).join('');
    rowsXml.push('<row r="' + rowN + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + xml + '</row>');
    return rowN;
  };
  // סגנונות: 1 כותרת, 2 תווית, 3 כותרת טבלה, 4 טקסט, 5 כמות, 6 כסף, 7 תווית סיכום, 8 סכום, 9 סכום אדום, 10 הערה, 11 שם פריט, 12 סכום ירוק
  var t = addRow([['נבחרת צימרי מבואות החרמון - הזמנה ' + order.id + (order.test ? ' (בדיקה)' : ''), 1]], 26);
  merges.push('A' + t + ':I' + t);
  order.info.forEach(function (kv) {
    var n = addRow([[kv[0], 2], [String(kv[1] === undefined ? '' : kv[1]), 0]]);
    merges.push('B' + n + ':E' + n);
  });
  addRow([]);
  addRow(['תמונה', 'קוד', 'מק"ט', 'פריט', 'צבע', 'פירוט', 'כמות', 'מחיר יחידה', 'סה"כ'].map(function (h) { return [h, 3]; }), 22);
  lines.forEach(function (l) {
    var n = addRow([['', 4], [l.code, 4], [l.sku, 4], [l.name, 11], [l.color || '', 4], [l.detail || '', 4],
      [l.qty, 5], [l.price, 6], [l.subtotal, 6]], 46);
    l.imgs.forEach(function (im, k) {
      if (!thumbs[im]) return;
      var mi = media.indexOf(im);
      if (mi < 0) { media.push(im); mi = media.length - 1; }
      anchors.push({ row: n - 1, k: k, media: mi });
    });
  });
  addRow([]);
  var tot = function (label, val, st) { addRow([null, null, null, null, null, null, null, [label, 7], [val, st]]); };
  tot('סה"כ הזמנה:', order.total, 8);
  tot('תקציב:', order.budget, 8);
  if (order.over > 0) tot('לתשלום ע"י הצימר:', order.over, 9); else tot('יתרת תקציב:', order.budget - order.total, 12);
  var nt = addRow([['הובלה והרכבה כלולות בתוכנית. לשינויים בהזמנה יש לפנות לצוות.', 10]]);
  merges.push('A' + nt + ':I' + nt);

  var widths = [20, 14, 22, 34, 12, 46, 8, 13, 13];
  var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="' + NS + '" xmlns:r="' + NSR + '"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
    '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
    '<cols>' + widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>' +
    '<sheetData>' + rowsXml.join('') + '</sheetData>' +
    '<mergeCells count="' + merges.length + '">' + merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join('') + '</mergeCells>' +
    '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
    '<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
    (anchors.length ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';

  var font = function (b, i, sz, color) { return '<font>' + (b ? '<b/>' : '') + (i ? '<i/>' : '') + '<sz val="' + sz + '"/><color rgb="' + color + '"/><name val="Arial"/></font>'; };
  var xf = function (numFmt, fontId, fillId, borderId, align) {
    return '<xf numFmtId="' + numFmt + '" fontId="' + fontId + '" fillId="' + fillId + '" borderId="' + borderId + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
      (align ? '<alignment ' + align + '/>' : '') + '</xf>';
  };
  var R = 'horizontal="right" vertical="center"';
  var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="' + NS + '">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;₪&quot;#,##0"/></numFmts>' +
    '<fonts count="8">' + font(0, 0, 11, 'FF2F3A26') + font(1, 0, 11, 'FF6E6A5E') + font(1, 0, 11, 'FFFFFFFF') + font(1, 0, 15, 'FF5C6B3A') +
    font(1, 0, 11, 'FFB23A2E') + font(0, 1, 11, 'FF6E6A5E') + font(1, 0, 11, 'FF2F3A26') + font(1, 0, 11, 'FF5C6B3A') + '</fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF5C6B3A"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left/><right/><top/><bottom style="thin"><color rgb="FFE6DDC8"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="13">' +
    xf(0, 0, 0, 0, R) + xf(0, 3, 0, 0, R) + xf(0, 1, 0, 0, R) + xf(0, 2, 2, 0, R) +
    xf(0, 0, 0, 1, R + ' wrapText="1"') + xf(0, 0, 0, 1, 'horizontal="center" vertical="center"') + xf(164, 0, 0, 1, R) +
    xf(0, 6, 0, 0, R) + xf(164, 6, 0, 0, R) + xf(164, 4, 0, 0, R) + xf(0, 5, 0, 0, R) + xf(0, 6, 0, 1, R + ' wrapText="1"') + xf(164, 7, 0, 0, R) +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  var rel = function (id, type, target) { return '<Relationship Id="' + id + '" Type="' + NSR + '/' + type + '" Target="' + target + '"/>'; };
  var rels = function (inner) { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + inner + '</Relationships>'; };
  var CT = 'application/vnd.openxmlformats-officedocument.';
  var parts = [
    { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="' + CT + 'spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + CT + 'spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="' + CT + 'spreadsheetml.styles+xml"/>' +
      (anchors.length ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="' + CT + 'drawing+xml"/>' : '') + '</Types>' },
    { name: '_rels/.rels', data: rels(rel('rId1', 'officeDocument', 'xl/workbook.xml')) },
    { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + NS + '" xmlns:r="' + NSR + '">' +
      '<sheets><sheet name="הזמנה" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: rels(rel('rId1', 'worksheet', 'worksheets/sheet1.xml') + rel('rId2', 'styles', 'styles.xml')) },
    { name: 'xl/styles.xml', data: styles },
    { name: 'xl/worksheets/sheet1.xml', data: sheet }
  ];
  if (anchors.length) {
    var EMU = 9525, W = 64, H = 48;
    var drawing = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NSR + '">' +
      anchors.map(function (a, i) {
        return '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>' + (4 + a.k * (W + 4)) * EMU + '</xdr:colOff><xdr:row>' + a.row + '</xdr:row><xdr:rowOff>' + 7 * EMU + '</xdr:rowOff></xdr:from>' +
          '<xdr:ext cx="' + W * EMU + '" cy="' + H * EMU + '"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 2) + '" name="Picture ' + (i + 1) + '"/>' +
          '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImg' + (a.media + 1) + '"/>' +
          '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + W * EMU + '" cy="' + H * EMU + '"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
      }).join('') + '</xdr:wsDr>';
    parts.push({ name: 'xl/worksheets/_rels/sheet1.xml.rels', data: rels(rel('rId1', 'drawing', '../drawings/drawing1.xml')) });
    parts.push({ name: 'xl/drawings/drawing1.xml', data: drawing });
    parts.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: rels(media.map(function (m, i) { return rel('rIdImg' + (i + 1), 'image', '../media/image' + (i + 1) + '.jpeg'); }).join('')) });
    media.forEach(function (m, i) { parts.push({ name: 'xl/media/image' + (i + 1) + '.jpeg', data: thumbs[m] }); });
  }
  return parts;
}

// =====================================================================
// מסכי ניהול (שלב 5) - כל פעולה נבדקת כאן בשרת מול אסימון כניסה
// הסיסמה נשמרת במאפייני הסקריפט (ADMIN_PASSWORD), לא בקוד.
// =====================================================================
var MOVE = { ORDER: 'הזמנה', CANCEL: 'ביטול', RECEIVE: 'כניסת סחורה', ADJUST: 'תיקון ידני' };

function adminLogin_(p) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('login-fails') || 0);
  if (fails >= 10) return { ok: false, error: 'יותר מדי ניסיונות כניסה. נסו שוב בעוד 10 דקות.' };
  var pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pass) return { ok: false, error: 'לא הוגדרה סיסמה (ADMIN_PASSWORD במאפייני הסקריפט).' };
  var name = String(p.name || '').trim().slice(0, 40);
  if (name.length < 2) return { ok: false, error: 'נא לכתוב שם' };
  if (String(p.password || '') !== pass) {
    cache.put('login-fails', String(fails + 1), 600);
    Utilities.sleep(800);
    return { ok: false, error: 'סיסמה שגויה' };
  }
  var token = Utilities.getUuid();
  cache.put('adm-' + token, name, 8 * 3600);
  return { ok: true, token: token, name: name };
}

// מחזיר את שם העובד, או זורק שגיאה אם האסימון לא תקף
function adminName_(p) {
  var name = p.token ? CacheService.getScriptCache().get('adm-' + p.token) : null;
  if (!name) { var e = new Error('נדרשת כניסה מחדש'); e.auth = true; throw e; }
  return name;
}

function cellVal_(v) { return v instanceof Date ? fmtDate_(v) : v; }

function adminData_(p) {
  var name = adminName_(p);
  var s = getSettings_();
  var t = readTable_(SH.PRODUCTS), c = t.col;
  var it = readTable_(SH.ITEMS), ic = it.col;
  var mv = readTable_(SH.MOVES), mc = mv.col;
  var o = readTable_(SH.ORDERS), oc = o.col;

  // סיכומי תנועות לכל מוצר
  var agg = {};
  var A = function (id) { return agg[id] || (agg[id] = { ordered: 0, received: 0, adjusted: 0, net: 0 }); };
  mv.rows.forEach(function (r) {
    var id = String(r[mc['מזהה']] || ''); if (!id) return;
    var ch = Number(r[mc['שינוי']]) || 0, type = String(r[mc['סוג']]);
    var a = A(id); a.net += ch;
    if (type === MOVE.RECEIVE) a.received += ch;
    if (type === MOVE.ADJUST) a.adjusted += ch;
  });
  // הוזמן = כמויות בהזמנות פעילות (כולל פריטים ללא הגבלת מלאי)
  it.rows.forEach(function (r) {
    var id = String(r[ic['מזהה']] || ''); if (!id) return;
    if (String(r[ic['סטטוס הזמנה']]) === ST.CANCELLED) return;
    A(id).ordered += Number(r[ic['כמות']]) || 0;
  });
  var products = t.rows.filter(function (r) { return r[c['מזהה']]; }).map(function (r) {
    var id = String(r[c['מזהה']]), a = agg[id] || { ordered: 0, received: 0, adjusted: 0, net: 0 };
    var unlimited = isYes_(r[c['ללא הגבלה']]);
    var initial = Number(r[c['מלאי התחלתי']]) || 0, remaining = Number(r[c['יתרה']]) || 0;
    return {
      id: id, code: String(r[c['קוד']]), sku: String(r[c['מק"ט']]), name: r[c['שם']], color: r[c['צבע']],
      cat: r[c['קטגוריה']], catLabel: r[c['שם קטגוריה']], price: Number(r[c['מחיר']]) || 0, img: r[c['תמונה']],
      kind: r[c['סוג']], tone: r[c['גוון']], diameter: r[c['קוטר']], active: isYes_(r[c['פעיל']]), unlimited: unlimited,
      initial: unlimited ? null : initial, remaining: unlimited ? null : remaining,
      ordered: a.ordered, received: a.received, adjusted: a.adjusted,
      // בדיקת עקביות: מלאי התחלתי + כל התנועות = יתרה (אם לא - מישהו שינה יתרה ידנית בגיליון)
      mismatch: !unlimited && (initial + a.net !== remaining)
    };
  });
  var items = {};
  it.rows.forEach(function (r) {
    var id = String(r[ic['מספר הזמנה']]); if (!id) return;
    (items[id] = items[id] || []).push({ id: r[ic['מזהה']], code: r[ic['קוד']], sku: r[ic['מק"ט']], name: r[ic['שם']], color: r[ic['צבע']],
      detail: r[ic['פירוט']], qty: Number(r[ic['כמות']]) || 0, price: Number(r[ic['מחיר יחידה']]) || 0, subtotal: Number(r[ic['סה"כ שורה']]) || 0 });
  });
  var pm = productMap_(t);
  var orders = o.rows.filter(function (r) { return r[oc['מספר הזמנה']]; }).map(function (r) {
    var id = String(r[oc['מספר הזמנה']]);
    return {
      id: id, test: id.indexOf('TEST-') === 0, date: cellVal_(r[oc['תאריך']]), ts: r[oc['תאריך']] instanceof Date ? r[oc['תאריך']].getTime() : 0,
      zimmer: r[oc['שם הצימר']], town: r[oc['ישוב']], contact: r[oc['איש קשר']], phone: normPhone_(r[oc['טלפון']]), email: r[oc['מייל']],
      address: r[oc['כתובת לאספקה']], day: r[oc['יום הגעה לאולם']], total: Number(r[oc['סה"כ']]) || 0, budget: Number(r[oc['תקציב']]) || 0,
      over: Number(r[oc['חריגה']]) || 0, status: String(r[oc['סטטוס']] || ''), mail: String(r[oc['סטטוס מייל']] || ''),
      excel: String(r[oc['קובץ Excel']] || ''), notes: String(r[oc['הערות צוות']] || ''), items: items[id] || [],
      lines: items[id] ? orderLines_(id, pm, it) : []
    };
  }).reverse();
  var moves = mv.rows.filter(function (r) { return r[mc['מזהה']]; }).map(function (r) {
    return { date: cellVal_(r[mc['תאריך']]), who: r[mc['מי']], id: r[mc['מזהה']], sku: r[mc['מק"ט']], change: Number(r[mc['שינוי']]) || 0,
      after: r[mc['יתרה אחרי']], type: r[mc['סוג']], ref: r[mc['הפניה']], note: r[mc['הערה']] };
  }).reverse().slice(0, 1000);
  return {
    ok: true, name: name, products: products, orders: orders, moves: moves,
    settings: { low: Number(s['סף מלאי נמוך']) || 3, open: isYes_(s['פתוח להזמנות']), budget: Number(s['תקציב']) || 9500 },
    mail: { remaining: MailApp.getRemainingDailyQuota(), sentToday: mailCountToday_() },
    time: fmtDate_(now_())
  };
}

// שינוי יתרה של מוצר + רישום ביומן. חייב לרוץ בתוך נעילה.
function changeStock_(t, pm, id, delta, type, who, ref, note, mvRows, mv) {
  var P = pm[id];
  var after = P.remaining + delta;
  t.sheet.getRange(P.i + 2, t.col['יתרה'] + 1).setValue(after);
  P.remaining = after;
  mvRows.push(rowFor_(mv, { 'תאריך': now_(), 'מי': who, 'מזהה': id, 'מק"ט': P.sku, 'שינוי': delta, 'יתרה אחרי': after,
    'סוג': type, 'הפניה': ref || '', 'הערה': note || '' }));
  return after;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, code: 'busy', error: 'המערכת עמוסה, נסו שוב בעוד רגע.' };
  try { return fn(); } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

// כניסת סחורה: מוסיף כמות למק"ט קיים
function receiveStock_(p) {
  var who = adminName_(p);
  var qty = Number(p.qty);
  if (!(qty >= 1 && qty <= 100000 && Math.floor(qty) === qty)) return { ok: false, error: 'כמות לא תקינה' };
  var note = String(p.note || '').trim().slice(0, 300);
  var date = String(p.date || '').trim().slice(0, 20);
  return withLock_(function () {
    var t = readTable_(SH.PRODUCTS), pm = productMap_(t), P = pm[String(p.id)];
    if (!P) return { ok: false, error: 'מוצר לא נמצא' };
    if (P.unlimited) return { ok: false, error: 'למוצר הזה אין ניהול מלאי (ללא הגבלה)' };
    var mv = readTable_(SH.MOVES), rows = [];
    var after = changeStock_(t, pm, P.id, qty, MOVE.RECEIVE, who, date ? 'תאריך קבלה ' + date : '', note, rows, mv);
    appendRows_(mv, rows);
    return { ok: true, id: P.id, remaining: after };
  });
}

// תיקון ידני: קובע יתרה חדשה (חובה לכתוב סיבה)
function adjustStock_(p) {
  var who = adminName_(p);
  var target = Number(p.remaining);
  if (!(target >= 0 && target <= 100000 && Math.floor(target) === target)) return { ok: false, error: 'יתרה לא תקינה' };
  var note = String(p.note || '').trim().slice(0, 300);
  if (note.length < 2) return { ok: false, error: 'חובה לכתוב סיבה לתיקון' };
  return withLock_(function () {
    var t = readTable_(SH.PRODUCTS), pm = productMap_(t), P = pm[String(p.id)];
    if (!P) return { ok: false, error: 'מוצר לא נמצא' };
    if (P.unlimited) return { ok: false, error: 'למוצר הזה אין ניהול מלאי (ללא הגבלה)' };
    var delta = target - P.remaining;
    if (delta === 0) return { ok: true, id: P.id, remaining: target, unchanged: true };
    var mv = readTable_(SH.MOVES), rows = [];
    changeStock_(t, pm, P.id, delta, MOVE.ADJUST, who, '', note, rows, mv);
    appendRows_(mv, rows);
    return { ok: true, id: P.id, remaining: target };
  });
}

// שינוי סטטוס הזמנה. ביטול מחזיר את הכמויות למלאי. הזמנה שבוטלה לא חוזרת (צריך הזמנה חדשה).
function setOrderStatus_(p) {
  var who = adminName_(p);
  var status = String(p.status || '');
  var allowed = [ST.NEW, ST.APPROVED, ST.SUPPLIED, ST.CANCELLED];
  if (allowed.indexOf(status) < 0) return { ok: false, error: 'סטטוס לא מוכר' };
  return withLock_(function () {
    var f = findOrder_(p.orderId);
    if (!f) return { ok: false, error: 'הזמנה לא נמצאה' };
    var cur = String(f.row[f.c['סטטוס']]);
    if (cur === status) return { ok: true, unchanged: true };
    if (cur === ST.CANCELLED) return { ok: false, error: 'הזמנה שבוטלה לא חוזרת. צריך לשלוח הזמנה חדשה מהאתר.' };
    var it = readTable_(SH.ITEMS), ic = it.col, returned = [];
    if (status === ST.CANCELLED) {
      var t = readTable_(SH.PRODUCTS), pm = productMap_(t), mv = readTable_(SH.MOVES), rows = [], back = {};
      it.rows.forEach(function (r) {
        if (String(r[ic['מספר הזמנה']]) !== String(p.orderId)) return;
        var id = String(r[ic['מזהה']]);
        back[id] = (back[id] || 0) + (Number(r[ic['כמות']]) || 0);
      });
      Object.keys(back).forEach(function (id) {
        if (!pm[id] || pm[id].unlimited) return;
        changeStock_(t, pm, id, back[id], MOVE.CANCEL, who, p.orderId, String(p.note || '').slice(0, 300) || f.row[f.c['שם הצימר']], rows, mv);
        returned.push({ id: id, qty: back[id] });
      });
      appendRows_(mv, rows);
    }
    // סטטוס בהזמנה ובפריטים
    setOrderField_(f, 'סטטוס', status);
    var note = String(p.note || '').trim();
    if (note) {
      var prev = String(f.row[f.c['הערות צוות']] || '');
      setOrderField_(f, 'הערות צוות', (prev ? prev + ' | ' : '') + status + ' (' + who + '): ' + note.slice(0, 300));
    }
    it.rows.forEach(function (r, i) {
      if (String(r[ic['מספר הזמנה']]) === String(p.orderId)) it.sheet.getRange(i + 2, ic['סטטוס הזמנה'] + 1).setValue(status);
    });
    return { ok: true, status: status, returned: returned };
  });
}

function resendEmail_(p) {
  adminName_(p);
  var f = findOrder_(p.orderId);
  if (!f) return { ok: false, error: 'הזמנה לא נמצאה' };
  if (!claimMail_(p.orderId)) return { ok: false, error: 'המייל של ההזמנה הזו נשלח ממש עכשיו. נסו שוב בעוד דקה.' };
  try {
    // אם אין קובץ Excel ב-Drive - השרת בונה אותו מהגיליון ומצרף
    var res = sendEmailForOrder_(p.orderId, '', 'שליחה חוזרת - ' + adminName_(p));
    return { ok: res.ok, error: res.error || '', attached: !!res.attached };
  } finally { releaseMail_(p.orderId); }
}

// =====================================================================
// נקודות כניסה
// =====================================================================
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
    if (p.action === 'submitOrder') return json_(submitOrder_(p));
    if (p.action === 'sendOrderEmail') return json_(sendOrderEmail_(p));
    var admin = { adminLogin: adminLogin_, adminData: adminData_, receiveStock: receiveStock_, adjustStock: adjustStock_,
      setOrderStatus: setOrderStatus_, resendEmail: resendEmail_ };
    if (admin[p.action]) return json_(admin[p.action](p));
    return json_({ ok: false, error: 'פעולה לא מוכרת: ' + (p.action || '') });
  } catch (err) {
    if (err && err.auth) return json_({ ok: false, code: 'auth', error: err.message });
    return json_({ ok: false, code: 'error', error: 'שגיאת שרת: ' + String(err && err.message || err) });
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
