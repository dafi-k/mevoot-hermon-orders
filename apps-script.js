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
  });
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
}

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
      if (l.type === 'table') {
        var lg = pm[l.leg], tp = pm[l.top], price = lg.price + tp.price;
        out.push({ kind: 'table', lid: l.lid, leg: lg, top: tp, qty: l.qty, price: price, subtotal: price * l.qty,
          code: lg.code + ' + ' + tp.code, sku: lg.sku + ' + ' + tp.sku, name: 'שולחן', color: tp.tone,
          detail: 'רגל ' + lg.code + ' (' + money_(lg.price) + ') + פלטה ' + tp.tone + ' ' + tp.diameter + ' ס"מ (' + money_(tp.price) + ')',
          img: tp.img, img2: lg.img });
      } else {
        var P = pm[l.id];
        out.push({ kind: 'item', lid: l.lid, p: P, qty: l.qty, price: P.price, subtotal: P.price * l.qty,
          code: P.code, sku: P.sku, name: P.name, color: P.color, detail: P.gift ? 'במתנה: ' + P.gift : '', img: P.img });
      }
      total += out[out.length - 1].subtotal;
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
        var tag = 'שולחן: ' + l.leg.code + ' + פלטה ' + l.top.tone + ' ' + l.top.diameter;
        itemRows.push(itemRow(l.leg, l.qty, l.leg.price, tag));
        itemRows.push(itemRow(l.top, l.qty, l.top.price, tag));
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
      lines: out.map(function (l) {
        return { kind: l.kind, lid: l.lid, code: l.code, sku: l.sku, name: l.name, color: l.color, detail: l.detail,
          qty: l.qty, price: l.price, subtotal: l.subtotal, img: l.img, img2: l.img2 || '' };
      })
    };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// מייל (בקשה נפרדת - כשל כאן לא מבטל את ההזמנה)
// =====================================================================
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

  // שמירת ה-Excel ב-Drive (כדי שאפשר יהיה לשלוח שוב בלי הלקוח)
  var fileId = '';
  try {
    var b64 = String(p.xlsxBase64 || '');
    if (b64 && b64.length < 8 * 1024 * 1024) {
      var name = String(p.xlsxName || (p.orderId + '.xlsx')).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 120);
      var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', name);
      var file = getFilesFolder_().createFile(blob);
      fileId = file.getId();
      setOrderField_(f, 'קובץ Excel', file.getUrl());
    }
  } catch (e) { /* ממשיכים לשלוח גם בלי קובץ */ }

  var res = sendEmailForOrder_(p.orderId, fileId);
  return { ok: true, emailed: res.ok, error: res.error || '' };
}

// שולח את מייל ההזמנה ומעדכן "סטטוס מייל". משמש גם לשליחה חוזרת.
function sendEmailForOrder_(orderId, fileId) {
  var f = findOrder_(orderId);
  if (!f) return { ok: false, error: 'הזמנה לא נמצאה' };
  var c = f.c, r = f.row, s = getSettings_();
  try {
    var staff = String(s['נמעני מייל צוות'] || '').split(/[,;\s]+/).filter(function (x) { return /@/.test(x); });
    var customer = String(r[c['מייל']] || '').trim();
    var recipients = staff.length + (customer ? 1 : 0);
    if (!recipients) throw new Error('אין נמענים');
    if (MailApp.getRemainingDailyQuota() < recipients) throw new Error('מכסת המיילים היומית נגמרה');

    if (!fileId) {
      var url = String(r[c['קובץ Excel']] || '');
      var m = url.match(/[-\w]{25,}/);
      if (m) fileId = m[0];
    }
    var attachments = [];
    if (fileId) { try { attachments.push(DriveApp.getFileById(fileId).getBlob()); } catch (e) {} }

    var total = Number(r[c['סה"כ']]) || 0, budget = Number(r[c['תקציב']]) || 0, over = Number(r[c['חריגה']]) || 0;
    var isTest = String(orderId).indexOf('TEST-') === 0;
    var subject = (isTest ? '[בדיקה] ' : '') + 'הזמנת ריהוט ' + orderId + ' - ' + r[c['שם הצימר']] + ' (' + r[c['ישוב']] + ')';
    var html = buildEmailHtml_(orderId, f, total, budget, over, isTest);
    var opts = { subject: subject, htmlBody: html, name: s['שם התוכנית'] || 'נבחרת צימרי מבואות החרמון', attachments: attachments };
    if (staff.length) { opts.to = staff.join(','); if (customer) opts.cc = customer; opts.replyTo = staff[0]; }
    else opts.to = customer;
    MailApp.sendEmail(opts);
    bumpMailCount_(recipients);
    var note = MAIL.SENT + ' ' + fmtDate_(now_()) + (staff.length ? '' : ' (ללקוח בלבד - אין נמעני צוות בהגדרות)');
    setOrderField_(f, 'סטטוס מייל', note);
    return { ok: true };
  } catch (err) {
    var msg = String(err && err.message || err);
    setOrderField_(f, 'סטטוס מייל', MAIL.FAILED + ': ' + msg);
    return { ok: false, error: msg };
  }
}

function buildEmailHtml_(orderId, f, total, budget, over, isTest) {
  var c = f.c, r = f.row;
  var it = readTable_(SH.ITEMS), ic = it.col;
  var rows = it.rows.filter(function (x) { return String(x[ic['מספר הזמנה']]) === String(orderId); });
  var td = 'padding:8px 10px;border-bottom:1px solid #E6DDC8;text-align:right;';
  var lines = rows.map(function (x) {
    return '<tr><td style="' + td + '"><b>' + esc_(x[ic['קוד']]) + '</b></td><td style="' + td + '">' + esc_(x[ic['שם']]) +
      (x[ic['צבע']] ? ' · ' + esc_(x[ic['צבע']]) : '') + (x[ic['פירוט']] ? '<br><span style="color:#6E6A5E;font-size:12px">' + esc_(x[ic['פירוט']]) + '</span>' : '') +
      '</td><td style="' + td + 'text-align:center">' + esc_(x[ic['כמות']]) + '</td><td style="' + td + '">' + money_(x[ic['מחיר יחידה']]) +
      '</td><td style="' + td + 'font-weight:bold">' + money_(x[ic['סה"כ שורה']]) + '</td></tr>';
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
    '<tr style="background:#5C6B3A;color:#fff"><th style="padding:8px 10px;text-align:right">קוד</th><th style="padding:8px 10px;text-align:right">פריט</th><th style="padding:8px 10px">כמות</th><th style="padding:8px 10px;text-align:right">מחיר</th><th style="padding:8px 10px;text-align:right">סה"כ</th></tr>' +
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

// שליחה חוזרת של כל המיילים שלא נשלחו (להריץ מהעורך עד שיהיה כפתור במסכי הניהול - שלב 5)
function resendFailedEmails() {
  var o = readTable_(SH.ORDERS), c = o.col, n = 0, ok = 0;
  o.rows.forEach(function (r) {
    var st = String(r[c['סטטוס מייל']] || '');
    if (st.indexOf(MAIL.SENT) === 0 || String(r[c['סטטוס']]) === ST.CANCELLED) return;
    n++;
    var res = sendEmailForOrder_(r[c['מספר הזמנה']], '');
    if (res.ok) ok++;
    Logger.log(r[c['מספר הזמנה']] + ': ' + (res.ok ? 'נשלח' : 'נכשל - ' + res.error));
  });
  Logger.log('נשלחו ' + ok + ' מתוך ' + n);
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
    return json_({ ok: false, error: 'פעולה לא מוכרת: ' + (p.action || '') });
  } catch (err) {
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
