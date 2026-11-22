// בדיקת קוד השרת (apps-script.js) מקומית, עם הדמיה של שירותי Google (גיליון, נעילה, מייל, Drive).
// הרצה: node test-server.js
const { makeEnv } = require('./mock-gas');

let fails = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fails++; };

const cust = (phone = '0521111111') => ({ zimmer: 'צימר הגפן', town: 'מסעדה', contact: 'רות', phone, email: 'ruth@example.com', address: 'הגפן 3, מסעדה', day: '2.12.26' });
const col = (sh, name) => sh.data[0].indexOf(name);
const cell = (sh, id, name) => { const r = sh.data.find(r => r[0] === id); return r ? r[col(sh, name)] : undefined; };
const setSetting = (env, k, v) => env.ctx.setSetting_(k, v);
const post = (env, body) => JSON.parse(env.ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s);

// ===================================================================
console.log('--- הקמה');
let env = makeEnv();
const s0 = env.ctx.getSettings_();
ok(s0['פתוח להזמנות'] === 'לא', 'ברירת מחדל: "פתוח להזמנות" = לא');
ok(/^test\d{6}$/.test(s0['קוד בדיקה']), 'נוצר קוד בדיקה: ' + s0['קוד בדיקה']);
ok(s0['תיקיית קבצים'] === 'FOLDER1', 'נוצרה תיקיית קבצים ב-Drive');
ok(cell(env.sheets['מוצרים'], 'AR-A6', 'יתרה') === 2, 'יתרה התחלתית = מלאי התחלתי');
const CODE = s0['קוד בדיקה'];

console.log('--- סגור להזמנות');
let r = post(env, { action: 'submitOrder', customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'closed', 'סגור: הזמנה רגילה נדחית');
r = post(env, { action: 'submitOrder', customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }], testCode: 'wrong' });
ok(!r.ok && r.code === 'closed', 'סגור: קוד בדיקה שגוי נדחה');

console.log('--- הזמנת בדיקה + מחירים מהשרת');
r = post(env, { action: 'submitOrder', testCode: CODE, overOk: false, customer: cust(),
  lines: [{ type: 'item', id: 'AR-A6', qty: 2, price: 1 }, { type: 'table', leg: 'LG-LP1', top: 'TOP-P11', qty: 1, price: 1 }, { type: 'item', id: 'OD27', qty: 1 }] });
ok(!r.ok && r.code === 'overConfirm', 'מעל התקציב בלי אישור: נדחה (מחיר שנשלח מהאתר התעלם ממנו)');
ok(env.sheets['הזמנות'].data.length === 1, 'שום דבר לא נכתב כשנדחה');
r = post(env, { action: 'submitOrder', testCode: CODE, overOk: true, customer: cust(),
  lines: [{ type: 'item', id: 'AR-A6', qty: 2, price: 1 }, { type: 'table', leg: 'LG-LP1', top: 'TOP-P11', qty: 1, price: 1, subtotal: 1 }, { type: 'item', id: 'OD27', qty: 1 }], total: 5 });
const expTotal = 899 * 2 + (799 + 3599) + 5600;
ok(r.ok && r.test && /^TEST-\d{6}-001$/.test(r.orderId), 'הזמנת בדיקה נקלטה: ' + r.orderId);
ok(r.total === expTotal && r.over === expTotal - 9500, `סה"כ ${r.total} וחריגה ${r.over} לפי מחירי הגיליון`);
const O = env.sheets['הזמנות'], IT = env.sheets['פריטי הזמנה'], MV = env.sheets['תנועות מלאי'];
ok(O.data.length === 2 && O.data[1][col(O, 'הערות צוות')] === 'הזמנת בדיקה' && O.data[1][col(O, 'סטטוס')] === 'חדשה' && O.data[1][col(O, 'סטטוס מייל')] === 'ממתין', 'שורת הזמנה: סטטוס חדשה, מייל ממתין, מסומנת כבדיקה');
ok(O.data[1][col(O, 'טלפון')] === '0521111111', 'טלפון נשמר עם 0');
ok(IT.data.length === 5, 'פריטי הזמנה: כורסא + רגל + פלטה + סט חוץ = 4 שורות');
ok(cell(env.sheets['מוצרים'], 'AR-A6', 'יתרה') === 0 && cell(env.sheets['מוצרים'], 'LG-LP1', 'יתרה') === 4 && cell(env.sheets['מוצרים'], 'TOP-P11', 'יתרה') === 3, 'יתרות ירדו: A6 2→0, LP1 5→4, P11 4→3');
ok(MV.data.length === 4 && MV.data.slice(1).every(m => m[col(MV, 'הפניה')] === r.orderId), 'יומן תנועות: 3 תנועות (סט החוץ ללא הגבלה - בלי תנועה)');
const firstId = r.orderId;

console.log('--- הזמנה אחת לטלפון');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'duplicate', 'אותו טלפון שוב: נדחה');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('052-111-1111'), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'duplicate', 'אותו טלפון בפורמט אחר (עם מקפים): נדחה');
O.data[1][col(O, 'סטטוס')] = 'בוטלה';
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(r.ok, 'אחרי שההזמנה בוטלה: אותו טלפון יכול להזמין שוב');

console.log('--- חוסר מלאי');
const before = JSON.stringify([O.data.length, IT.data.length, MV.data.length, cell(env.sheets['מוצרים'], 'LG-LP1', 'יתרה')]);
r = post(env, { action: 'submitOrder', testCode: CODE, overOk: true, customer: cust('0522222222'),
  lines: [{ type: 'item', id: 'AR-A6', qty: 1 }, { type: 'table', leg: 'LG-LP1', top: 'TOP-P1', qty: 3 }, { type: 'table', leg: 'LG-LP1', top: 'TOP-P4', qty: 2 }, { type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'shortage', 'חוסר: נדחה עם קוד shortage');
const sm = Object.fromEntries((r.shortages || []).map(x => [x.id, x]));
ok(sm['AR-A6'] && sm['AR-A6'].available === 0, 'A6: אזל (זמין 0)');
ok(sm['LG-LP1'] && sm['LG-LP1'].requested === 5 && sm['LG-LP1'].available === 4, 'רגל LP1 משותפת לשני שולחנות: ביקשו 5, זמין 4');
ok(!sm['CH-P1'] && !sm['TOP-P1'], 'פריטים שיש מהם מספיק לא מסומנים');
ok(JSON.stringify([O.data.length, IT.data.length, MV.data.length, cell(env.sheets['מוצרים'], 'LG-LP1', 'יתרה')]) === before, 'בחוסר - שום דבר לא נכתב');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('0522222222'), lines: [{ type: 'item', id: 'NOPE', qty: 1 }] });
ok(!r.ok && r.code === 'shortage' && r.shortages[0].id === 'NOPE', 'מזהה לא קיים: מסומן כחסר');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('0522222222'), lines: [{ type: 'table', leg: 'CH-P1', top: 'TOP-P1', qty: 1 }] });
ok(!r.ok && r.code === 'shortage', 'שולחן עם "רגל" שהיא לא רגל: נדחה');

console.log('--- ולידציה');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: { ...cust('0523333333'), phone: '123' }, lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'invalid', 'טלפון לא תקין: נדחה');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('0523333333'), lines: [{ type: 'item', id: 'CH-P1', qty: 1.5 }] });
ok(!r.ok && r.code === 'invalid', 'כמות לא שלמה: נדחתה');
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('0523333333'), lines: [] });
ok(!r.ok && r.code === 'invalid', 'עגלה ריקה: נדחתה');

console.log('--- נעילה תפוסה');
env.setLockBusy(true);
r = post(env, { action: 'submitOrder', testCode: CODE, customer: cust('0523333333'), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
ok(!r.ok && r.code === 'busy' && r.error === 'המערכת עמוסה, נסו שוב בעוד רגע.', 'נעילה תפוסה: "המערכת עמוסה, נסו שוב בעוד רגע."');
env.setLockBusy(false);

console.log('--- פתוח להזמנות');
setSetting(env, 'פתוח להזמנות', 'כן');
r = post(env, { action: 'submitOrder', customer: cust('0524444444'), lines: [{ type: 'item', id: 'CH-P1', qty: 2 }] });
ok(r.ok && /^MH-/.test(r.orderId) && !r.test, 'פתוח: הזמנה רגילה נקלטת ' + r.orderId);
const realId = r.orderId;

console.log('--- מייל');
setSetting(env, 'נמעני מייל צוות', 'a@contrast.test, b@contrast.test');
r = post(env, { action: 'sendOrderEmail', orderId: realId, phone: '0529999999', xlsxBase64: 'AAAA' });
ok(!r.ok, 'מייל עם טלפון לא תואם: נדחה');
r = post(env, { action: 'sendOrderEmail', orderId: realId, phone: '0524444444', xlsxBase64: Buffer.from('xlsx').toString('base64'), xlsxName: 'הזמנה.xlsx' });
ok(r.ok && r.emailed, 'מייל נשלח');
const m = env.mails[0];
ok(m && m.to === 'a@contrast.test,b@contrast.test' && m.cc === 'ruth@example.com' && m.attachments.length === 1, 'נמענים: צוות ב-To, הלקוח ב-CC, Excel מצורף');
ok(/^נשלח /.test(cell(O, realId, 'סטטוס מייל')) && /drive\.google\.com/.test(cell(O, realId, 'קובץ Excel')), 'סטטוס מייל "נשלח" + קישור לקובץ ב-Drive');
ok(env.ctx.mailCountToday_() === 3, 'ספירת נמענים יומית: 3');
r = post(env, { action: 'sendOrderEmail', orderId: realId, phone: '0524444444', xlsxBase64: 'AAAA' });
ok(r.ok && r.already && env.mails.length === 1, 'שליחה כפולה של אותו מייל: לא נשלח פעמיים');

console.log('--- כשל במייל לא מבטל הזמנה');
let env2 = makeEnv({ mailThrows: 'Service invoked too many times for one day: email.' });
env2.ctx.setSetting_('נמעני מייל צוות', 'a@contrast.test');
const C2 = env2.ctx.getSettings_()['קוד בדיקה'];
r = post(env2, { action: 'submitOrder', testCode: C2, customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
const id2 = r.orderId;
r = post(env2, { action: 'sendOrderEmail', orderId: id2, phone: '0521111111', xlsxBase64: 'AAAA' });
const O2 = env2.sheets['הזמנות'];
ok(r.ok && r.emailed === false, 'כשל במייל: התשובה מודיעה שלא נשלח');
ok(cell(O2, id2, 'סטטוס') === 'חדשה' && /^לא נשלח: /.test(cell(O2, id2, 'סטטוס מייל')), 'ההזמנה נשארה "חדשה", וסטטוס המייל "לא נשלח: <סיבה>"');
let env3 = makeEnv({ quota: 1 });
env3.ctx.setSetting_('נמעני מייל צוות', 'a@contrast.test, b@contrast.test');
const C3 = env3.ctx.getSettings_()['קוד בדיקה'];
r = post(env3, { action: 'submitOrder', testCode: C3, customer: cust(), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
const id3 = r.orderId;
post(env3, { action: 'sendOrderEmail', orderId: id3, phone: '0521111111', xlsxBase64: 'AAAA' });
ok(/מכסת המיילים היומית נגמרה/.test(cell(env3.sheets['הזמנות'], id3, 'סטטוס מייל')), 'מכסה יומית נגמרה: נרשם "לא נשלח: מכסת המיילים היומית נגמרה"');
// שליחה חוזרת
env2.ctx.MailApp.sendEmail = o => env2.mails.push(o);
env2.ctx.resendFailedEmails();
ok(/^נשלח /.test(cell(O2, id2, 'סטטוס מייל')) && env2.mails.length === 1, 'resendFailedEmails שולח שוב ומעדכן ל"נשלח"');

// ===================================================================
console.log('--- ניהול: כניסה');
const A = makeEnv();
const AC = A.ctx.getSettings_()['קוד בדיקה'];
const AP = b => JSON.parse(A.ctx.doPost({ postData: { contents: JSON.stringify(b) } }).s);
r = AP({ action: 'adminLogin', name: 'דפי', password: 'x' });
ok(!r.ok && /לא הוגדרה סיסמה/.test(r.error), 'בלי ADMIN_PASSWORD: כניסה נחסמת');
A.props.ADMIN_PASSWORD = 'Sod-2026!';
r = AP({ action: 'adminLogin', name: 'דפי', password: 'wrong' });
ok(!r.ok && r.error === 'סיסמה שגויה', 'סיסמה שגויה נדחית');
r = AP({ action: 'adminData', token: 'fake' });
ok(!r.ok && r.code === 'auth', 'בלי אסימון תקף: אין גישה לנתונים');
r = AP({ action: 'receiveStock', token: 'fake', id: 'CH-P1', qty: 5 });
ok(!r.ok && r.code === 'auth', 'בלי אסימון תקף: אין כניסת סחורה');
const login = AP({ action: 'adminLogin', name: 'דפי', password: 'Sod-2026!' });
ok(login.ok && login.token, 'סיסמה נכונה: מתקבל אסימון');
const TK = login.token;
for (let i = 0; i < 10; i++) AP({ action: 'adminLogin', name: 'פורץ', password: 'bad' + i });
r = AP({ action: 'adminLogin', name: 'דפי', password: 'Sod-2026!' });
ok(!r.ok && /יותר מדי ניסיונות/.test(r.error), 'אחרי 10 ניסיונות שגויים: חסימה זמנית');

console.log('--- ניהול: הזמנה, ביטול ומלאי');
const PR = A.sheets['מוצרים'], OR = A.sheets['הזמנות'], ITm = A.sheets['פריטי הזמנה'], MVm = A.sheets['תנועות מלאי'];
const rem = id => cell(PR, id, 'יתרה');
const o1 = AP({ action: 'submitOrder', testCode: AC, overOk: true, customer: cust('0527777777'),
  lines: [{ type: 'item', id: 'DC1', qty: 3 }, { type: 'table', leg: 'LG-LP3', top: 'TOP-P2', qty: 2 }, { type: 'item', id: 'OD5', qty: 1 }] });
ok(o1.ok && rem('DC1') === 9 && rem('LG-LP3') === 9 && rem('TOP-P2') === 5, 'הזמנה הורידה מלאי (DC1 12→9, LP3 11→9, P2 7→5)');
let d = AP({ action: 'adminData', token: TK });
ok(d.ok && d.products.length === 160 && d.orders.length === 1 && d.orders[0].items.length === 4, 'נתוני ניהול: 160 מוצרים, הזמנה אחת עם 4 שורות');
const pDC1 = d.products.find(p => p.id === 'DC1');
ok(pDC1.initial === 12 && pDC1.ordered === 3 && pDC1.remaining === 9 && !pDC1.mismatch, 'מלאי DC1: התחלתי 12, הוזמן 3, יתרה 9, עקבי');
ok(d.products.find(p => p.id === 'OD5').ordered === 1 && d.products.find(p => p.id === 'OD5').remaining === null, 'ריהוט חוץ: נספר כהוזמן, בלי יתרה');
ok(/^\d\d\/\d\d\/\d{4} \d\d:\d\d$/.test(d.orders[0].date) && /^\d\d\/\d\d\/\d{4} \d\d:\d\d$/.test(d.moves[0].date), 'תאריך ושעה בהזמנות וביומן: ' + d.orders[0].date);

r = AP({ action: 'receiveStock', token: TK, id: 'DC1', qty: 10, date: '25.11.26', note: 'משלוח מהספק' });
ok(r.ok && rem('DC1') === 19, 'כניסת סחורה: DC1 9→19');
const lastMove = MVm.data[MVm.data.length - 1];
ok(lastMove[col(MVm, 'סוג')] === 'כניסת סחורה' && lastMove[col(MVm, 'מי')] === 'דפי' && lastMove[col(MVm, 'שינוי')] === 10 && /25\.11\.26/.test(lastMove[col(MVm, 'הפניה')]), 'יומן: כניסת סחורה, מי=דפי, +10, תאריך קבלה');
r = AP({ action: 'receiveStock', token: TK, id: 'OD5', qty: 1 });
ok(!r.ok, 'כניסת סחורה למוצר ללא הגבלה: נדחית');
r = AP({ action: 'receiveStock', token: TK, id: 'DC1', qty: -2 });
ok(!r.ok, 'כמות שלילית בכניסת סחורה: נדחית');
r = AP({ action: 'adjustStock', token: TK, id: 'DC1', remaining: 17 });
ok(!r.ok && /סיבה/.test(r.error), 'תיקון ידני בלי סיבה: נדחה');
r = AP({ action: 'adjustStock', token: TK, id: 'DC1', remaining: 17, note: 'ספירת מלאי - 2 פגומים' });
ok(r.ok && rem('DC1') === 17 && MVm.data[MVm.data.length - 1][col(MVm, 'שינוי')] === -2, 'תיקון ידני: DC1 19→17, ביומן -2');

r = AP({ action: 'setOrderStatus', token: TK, orderId: o1.orderId, status: 'אושרה' });
ok(r.ok && cell(OR, o1.orderId, 'סטטוס') === 'אושרה' && rem('DC1') === 17, 'אישור הזמנה: הסטטוס משתנה, המלאי לא');
r = AP({ action: 'setOrderStatus', token: TK, orderId: o1.orderId, status: 'בוטלה', note: 'בדיקה' });
ok(r.ok && rem('DC1') === 20 && rem('LG-LP3') === 11 && rem('TOP-P2') === 7, 'ביטול: הכמויות חזרו למלאי (DC1 +3, LP3 +2, P2 +2)');
ok(ITm.data.slice(1).every(x => x[col(ITm, 'סטטוס הזמנה')] === 'בוטלה'), 'ביטול: כל פריטי ההזמנה מסומנים "בוטלה"');
ok(MVm.data.filter(x => x[col(MVm, 'סוג')] === 'ביטול').length === 3, 'ביטול: 3 תנועות "ביטול" ביומן (בלי ריהוט החוץ)');
ok(/בוטלה \(דפי\): בדיקה/.test(cell(OR, o1.orderId, 'הערות צוות')), 'ביטול: נרשם בהערות צוות מי ביטל ולמה');
r = AP({ action: 'setOrderStatus', token: TK, orderId: o1.orderId, status: 'בוטלה' });
ok(r.ok && r.unchanged && rem('DC1') === 20, 'ביטול כפול: לא מחזיר מלאי פעמיים');
r = AP({ action: 'setOrderStatus', token: TK, orderId: o1.orderId, status: 'חדשה' });
ok(!r.ok && /לא חוזרת/.test(r.error), 'הזמנה שבוטלה לא חוזרת לחדשה');
d = AP({ action: 'adminData', token: TK });
const pDC1b = d.products.find(p => p.id === 'DC1');
ok(pDC1b.ordered === 0 && pDC1b.received === 10 && pDC1b.adjusted === -2 && pDC1b.remaining === 20 && !pDC1b.mismatch, 'אחרי הכל: הוזמן 0, נכנס 10, תיקון -2, יתרה 20 - עקבי');
PR.data.find(x => x[0] === 'DC2')[col(PR, 'יתרה')] = 99;
d = AP({ action: 'adminData', token: TK });
ok(d.products.find(p => p.id === 'DC2').mismatch, 'שינוי ידני של יתרה בגיליון מזוהה כ"לא עקבי"');

console.log('--- ניהול: שליחה חוזרת של מייל');
A.ctx.setSetting_('נמעני מייל צוות', 'a@contrast.test');
const o2 = AP({ action: 'submitOrder', testCode: AC, customer: cust('0528888888'), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
r = AP({ action: 'resendEmail', token: 'fake', orderId: o2.orderId });
ok(!r.ok && r.code === 'auth', 'שליחה חוזרת בלי כניסה: נחסמת');
r = AP({ action: 'resendEmail', token: TK, orderId: o2.orderId });
ok(r.ok && /^נשלח /.test(cell(OR, o2.orderId, 'סטטוס מייל')) && A.mails.length === 1, 'שליחה חוזרת: נשלח ועודכן');
ok(/QUERY\('פריטי הזמנה'!A:M/.test(A.sheets['ריכוז כמויות'].formula) && /not A starts with 'TEST'/.test(A.sheets['סיכום לפי צימר'].formula), 'לשוניות ריכוז כמויות וסיכום לפי צימר: נוסחאות QUERY בלי הזמנות בדיקה');

console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל בדיקות השרת עברו');
process.exit(fails ? 1 : 0);
