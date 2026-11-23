// בדיקות גיבוי המייל בשרת: Excel שנבנה מהגיליון, טריגר "ממתין" אחרי 2 דקות, שליחה חוזרת עם Excel,
// מניעת שליחה כפולה, והצגת שולחן (רגל + פלטה) במייל, ב-Excel ובמסכי הניהול.
// הרצה: node test-mail-backup.js
const JSZip = require('jszip');
const ExcelJS = require('exceljs');
const { makeEnv } = require('./mock-gas');
let fails = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fails++; };
const col = (sh, n) => sh.data[0].indexOf(n);
const cell = (sh, id, n) => { const r = sh.data.find(r => r[0] === id); return r ? r[col(sh, n)] : undefined; };
const setCell = (sh, id, n, v) => { sh.data.find(r => r[0] === id)[col(sh, n)] = v; };
const cust = phone => ({ zimmer: 'צימר הגפן', town: 'מסעדה', contact: 'רות', phone, email: 'ruth@example.com', address: 'הגפן 3, מסעדה', day: '1.12.26' });
const TABLE = { type: 'table', leg: 'LG-LP2', top: 'TOP-P8', qty: 1 };

// zip מדומה -> קובץ xlsx אמיתי -> ExcelJS
async function openXlsx(blob) {
  const z = new JSZip();
  blob.entries.forEach(b => z.file(b.name, b.bytes));
  const buf = await z.generateAsync({ type: 'nodebuffer' });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return { wb, ws: wb.worksheets[0], size: buf.length };
}
// תמונות בתוך המייל: כל src הוא cid: שקיים ב-inlineImages כ-JPEG, ואין קישור חיצוני לתמונה
const inlineCheck = (m, title) => {
  const srcs = [...m.htmlBody.matchAll(/<img[^>]+src="([^"]+)"/g)].map(x => x[1]);
  const inl = m.inlineImages || {};
  const allCid = srcs.length > 0 && srcs.every(s => /^cid:img[0-9]+$/.test(s) && inl[s.slice(4)] && inl[s.slice(4)].type === 'image/jpeg' && inl[s.slice(4)].bytes.length > 500);
  const noExternal = !/<img[^>]+src="https?:/.test(m.htmlBody);
  const row = title ? (m.htmlBody.split('<tr>').find(r => r.includes(title)) || '') : '';
  const tablePics = title ? (row.match(/src="cid:/g) || []).length : 0;
  const kb = Object.values(inl).reduce((a, b) => a + b.bytes.length, 0) / 1024;
  return { allCid, noExternal, tablePics, n: srcs.length, kb };
};
const rowsText = ws => { const out = []; ws.eachRow(r => out.push(r.values.slice(1).map(v => v && v.richText ? v.richText.map(x => x.text).join('') : v))); return out; };

(async () => {
  const env = makeEnv();
  const E = env.ctx, P = b => JSON.parse(E.doPost({ postData: { contents: JSON.stringify(b) } }).s);
  E.setSetting_('נמעני מייל צוות', 'a@contrast.test, b@contrast.test');
  E.setSetting_('פתוח להזמנות', 'כן');
  env.props.ADMIN_PASSWORD = 'Sod-2026!';
  const O = env.sheets['הזמנות'];
  const lp2 = env.sheets['מוצרים'].data.find(r => r[0] === 'LG-LP2'), p11 = env.sheets['מוצרים'].data.find(r => r[0] === 'TOP-P8');
  const legPrice = lp2[col(env.sheets['מוצרים'], 'מחיר')], topPrice = p11[col(env.sheets['מוצרים'], 'מחיר')];
  const tone = p11[col(env.sheets['מוצרים'], 'גוון')], dia = p11[col(env.sheets['מוצרים'], 'קוטר')];
  const title = `שולחן: רגל LP2 + פלטה ${tone} קוטר ${dia}`;

  console.log('--- תשובת השרת לשולחן');
  let r = P({ action: 'submitOrder', customer: cust('0521000001'), overOk: true, lines: [{ type: 'item', id: 'CH-P1', qty: 2 }, { ...TABLE, qty: 2 }] });
  ok(r.ok, 'הזמנה עם שולחן נקלטה: ' + r.orderId);
  const t = r.lines.find(l => l.kind === 'table');
  ok(t.name === title, `שם השולחן: "${t.name}"`);
  ok(t.legPrice === legPrice && t.topPrice === topPrice && t.price === legPrice + topPrice && t.subtotal === (legPrice + topPrice) * 2, 'מחיר רגל, מחיר פלטה וסה"כ לשולחן בנפרד');
  ok(t.imgs.length === 2 && t.imgs[0] === 'LG-LP2.webp' && /^TOP-/.test(t.imgs[1]), 'שתי תמונות: רגל ואז פלטה (' + t.imgs.join(', ') + ')');
  ok(t.detail.includes('רגל LP2') && t.detail.includes('סה"כ לשולחן'), 'פירוט: ' + t.detail);
  const id1 = r.orderId;

  console.log('--- מייל בלי Excel מהאתר: השרת בונה Excel');
  r = P({ action: 'sendOrderEmail', orderId: id1, phone: '0521000001' }); // כמו אייפד שלא העלה קובץ
  ok(r.ok && r.emailed, 'המייל נשלח');
  const m1 = env.mails[env.mails.length - 1];
  ok(m1.attachments.length === 1 && /\.xlsx$/.test(m1.attachments[0].name), 'מצורף Excel שנבנה בשרת: ' + m1.attachments[0].name);
  ok(/^https:\/\/drive/.test(cell(O, id1, 'קובץ Excel')), 'הקובץ נשמר ב-Drive ונרשם בהזמנה');
  const x1 = await openXlsx(m1.attachments[0]);
  const rows = rowsText(x1.ws);
  ok(x1.ws.views[0].rightToLeft, 'Excel מימין לשמאל');
  ok(String(rows[0][0]).includes(id1), 'כותרת עם מספר ההזמנה');
  ok(rows.some(rw => rw[0] === 'שם הצימר' && String(rw[1]).includes('צימר הגפן')), 'פרטי הצימר');
  const tr = rows.find(rw => rw[3] === title);
  ok(!!tr, 'שורת שולחן אחת עם התיאור המלא');
  ok(tr && tr[6] === 2 && tr[7] === legPrice + topPrice && tr[8] === (legPrice + topPrice) * 2 && String(tr[5]).includes('סה"כ לשולחן'), 'שולחן ב-Excel: כמות, מחיר לשולחן, סה"כ, ופירוט מחירים');
  ok(!rows.some(rw => rw[3] === lp2[col(env.sheets['מוצרים'], 'שם')] && rw[1] === 'LP2'), 'הרגל לא מופיעה כשורה נפרדת');
  const imgs = x1.ws.getImages();
  let trNum = 0;
  x1.ws.eachRow((rw, n) => { if (rw.values[4] === title) trNum = n; });
  const tableImgs = imgs.filter(im => im.range.tl.nativeRow === trNum - 1);
  ok(imgs.length === 3 && tableImgs.length === 2, `תמונות ב-Excel: ${imgs.length} (2 בשורת השולחן - רגל ופלטה)`);
  ok(rows.some(rw => rw[7] === 'סה"כ הזמנה:' && rw[8] === cell(O, id1, 'סה"כ')), 'סה"כ הזמנה ב-Excel = הגיליון');
  ok(env.fetches.every(u => /^https:\/\/dafi-k\.github\.io\/mevoot-hermon-orders\/pic\/thumbs\/.+\.jpg$/.test(u)), 'תמונות מוקטנות מהאתר הציבורי (jpg)');

  console.log('--- המייל: שולחן עם 2 תמונות');
  const i1 = inlineCheck(m1, title);
  ok(m1.htmlBody.includes(title) && i1.tablePics === 2, 'במייל: תיאור השולחן + תמונת רגל + תמונת פלטה');
  ok(i1.allCid && i1.noExternal, `במייל: ${i1.n} תמונות בתוך המייל (cid), בלי קישור חיצוני`);
  ok(i1.kb < 60, `במייל: התמונות קלות - ${i1.kb.toFixed(1)}KB בסך הכל`);
  ok(m1.htmlBody.includes('סה&quot;כ לשולחן'), 'במייל: מחיר רגל ומחיר פלטה בנפרד');

  console.log('--- ניסיון חוזר לא יוצר כפילות');
  const nFiles = Object.keys(env.files).length, nMails = env.mails.length;
  r = P({ action: 'sendOrderEmail', orderId: id1, phone: '0521000001', xlsxBase64: 'AAAA' });
  ok(r.ok && r.already && env.mails.length === nMails && Object.keys(env.files).length === nFiles, 'הזמנה שכבר נשלחה: אין מייל נוסף ואין קובץ נוסף');

  console.log('--- מניעת שליחה כפולה במקביל');
  r = P({ action: 'submitOrder', customer: cust('0521000002'), lines: [{ type: 'item', id: 'CH-P1', qty: 1 }] });
  const id2 = r.orderId;
  env.cache['mailclaim-' + id2] = '1'; // שליחה אחרת "רצה" כרגע
  r = P({ action: 'sendOrderEmail', orderId: id2, phone: '0521000002', xlsxBase64: 'AAAA' });
  ok(r.ok && r.pending && !r.emailed && cell(O, id2, 'סטטוס מייל') === 'ממתין', 'בזמן ששליחה אחרת רצה: תשובה "בטיפול", בלי מייל כפול');
  delete env.cache['mailclaim-' + id2];
  r = P({ action: 'sendOrderEmail', orderId: id2, phone: '0521000002', xlsxBase64: Buffer.from('xlsx-from-site').toString('base64'), xlsxName: 'a.xlsx' });
  ok(r.emailed && env.mails[env.mails.length - 1].attachments[0].name === 'a.xlsx', 'Excel מהאתר נשמר ומצורף (השרת לא בונה במקומו)');
  ok(!env.cache['mailclaim-' + id2], 'התפיסה משתחררת אחרי השליחה');

  console.log('--- גיבוי אוטומטי (טריגר כל דקה)');
  E.installMailWatchdog();
  E.installMailWatchdog();
  ok(env.triggers.length === 1 && env.triggers[0].fn === 'mailWatchdog' && env.triggers[0].m === 1, 'טריגר אחד בלבד, כל דקה (גם אחרי התקנה כפולה)');
  ok(Math.abs(Number(env.props.WATCHDOG_SINCE) - Date.now()) < 5000, 'נשמר זמן ההתקנה');
  env.props.WATCHDOG_SINCE = String(Date.now() - 10 * 60000); // כאילו הותקן לפני 10 דקות
  const since = Number(env.props.WATCHDOG_SINCE);
  const GDate = E.now_().constructor; // תאריך מאותו "עולם" של קוד השרת (כמו בגיליון אמיתי)
  const mk = (phone, minsAgo) => {
    const x = P({ action: 'submitOrder', customer: cust(phone), lines: [TABLE] });
    setCell(O, x.orderId, 'תאריך', new GDate(Date.now() - minsAgo * 60000));
    return x.orderId;
  };
  const old = mk('0521000003', 3);      // ממתין 3 דקות -> נשלח
  const fresh = mk('0521000004', 1);    // רק דקה -> מחכים
  const legacy = mk('0521000005', 3);   // לפני התקנת הטריגר -> לא נוגעים
  setCell(O, legacy, 'תאריך', new GDate(since - 60 * 60000));
  const cancelled = mk('0521000006', 3);
  setCell(O, cancelled, 'סטטוס', 'בוטלה');
  const before = env.mails.length;
  E.mailWatchdog();
  const st = id => String(cell(O, id, 'סטטוס מייל'));
  ok(/^נשלח .*\(גיבוי אוטומטי\)/.test(st(old)), 'הזמנה "ממתין" מעל 2 דקות: נשלחה - ' + st(old));
  const mw = env.mails[env.mails.length - 1];
  const iw = inlineCheck(mw, title);
  ok(iw.allCid && iw.noExternal && iw.tablePics === 2, 'גיבוי אוטומטי: תמונות בתוך המייל, רגל + פלטה בשורת השולחן');
  ok(env.mails.length === before + 1 && mw.attachments.length === 1 && mw.subject.includes(old), 'בדיוק מייל אחד, עם Excel מצורף');
  ok(st(fresh) === 'ממתין', 'הזמנה בת דקה: עדיין מחכה לאתר');
  ok(st(legacy) === 'ממתין', 'הזמנה ישנה מלפני התקנת הגיבוי: לא נשלחת פתאום');
  ok(st(cancelled) === 'ממתין', 'הזמנה שבוטלה: לא נשלחת');
  E.mailWatchdog();
  ok(env.mails.length === before + 1, 'הרצה נוספת: לא שולחת שוב');
  setCell(O, fresh, 'תאריך', new GDate(Date.now() - 3 * 60000));
  E.mailWatchdog();
  ok(/^נשלח/.test(st(fresh)), 'אחרי שעברו 2 דקות: גם היא נשלחה');

  console.log('--- שליחה חוזרת ממסך הניהול (כמו MH-261123-006/007)');
  const tk = P({ action: 'adminLogin', name: 'דפי', password: 'Sod-2026!' }).token;
  r = P({ action: 'submitOrder', customer: cust('0521000007'), lines: [TABLE, { type: 'item', id: 'DC1', qty: 1 }] });
  const id7 = r.orderId;
  setCell(O, id7, 'סטטוס מייל', 'נשלח 23/11/2026 10:00'); // נשלח בעבר בלי Excel
  r = P({ action: 'resendEmail', token: tk, orderId: id7 });
  const m7 = env.mails[env.mails.length - 1];
  const i7 = inlineCheck(m7, title);
  ok(i7.allCid && i7.noExternal && i7.tablePics === 2 && i7.n === 3, 'שלח שוב: תמונות בתוך המייל (רגל + פלטה + כיסא)');
  ok(r.ok && r.attached && m7.subject.includes(id7) && m7.attachments.length === 1, 'שליחה חוזרת כשאין Excel: השרת בונה ומצרף');
  ok(/שליחה חוזרת - דפי/.test(st(id7)) && /^https:\/\/drive/.test(cell(O, id7, 'קובץ Excel')), 'הסטטוס מציין שליחה חוזרת ומי שלח, והקובץ נשמר');
  const nF = Object.keys(env.files).length;
  r = P({ action: 'resendEmail', token: tk, orderId: id7 });
  ok(r.ok && Object.keys(env.files).length === nF && env.mails[env.mails.length - 1].attachments.length === 1, 'שליחה חוזרת שנייה: משתמשת בקובץ הקיים');
  env.cache['mailclaim-' + id7] = '1';
  r = P({ action: 'resendEmail', token: tk, orderId: id7 });
  ok(!r.ok && /עכשיו/.test(r.error), 'שליחה חוזרת בזמן שליחה אחרת: נחסמת עם הסבר');
  delete env.cache['mailclaim-' + id7];

  console.log('--- מסכי ניהול: שולחן כשורה אחת');
  const d = P({ action: 'adminData', token: tk });
  const o7 = d.orders.find(o => o.id === id7);
  const tl = o7.lines.find(l => l.kind === 'table');
  ok(o7.items.length === 3 && o7.lines.length === 2, 'פריטים גולמיים 3 (לריכוזים), שורות תצוגה 2');
  ok(tl && tl.name === title && tl.imgs.length === 2 && tl.legPrice === legPrice && tl.topPrice === topPrice, 'שורת שולחן: תיאור, 2 תמונות, מחירים בנפרד');

  console.log('--- הזמנות ישנות (תיאור שולחן בפורמט הקודם)');
  r = P({ action: 'submitOrder', customer: cust('0521000008'), lines: [TABLE] });
  const IT = env.sheets['פריטי הזמנה'];
  IT.data.filter(x => x[0] === r.orderId).forEach(x => { x[col(IT, 'פירוט')] = `שולחן: LP2 + פלטה ${tone} ${dia}`; });
  const lines8 = E.orderLines_(r.orderId, E.productMap_(E.readTable_('מוצרים')));
  ok(lines8.length === 1 && lines8[0].name === title, 'גם הזמנה מלפני העדכון מוצגת כשולחן אחד');

  console.log('--- בלי גישה לתמונות');
  const env2 = makeEnv({ noFetch: true });
  env2.ctx.setSetting_('נמעני מייל צוות', 'a@contrast.test');
  env2.ctx.setSetting_('פתוח להזמנות', 'כן');
  const P2 = b => JSON.parse(env2.ctx.doPost({ postData: { contents: JSON.stringify(b) } }).s);
  r = P2({ action: 'submitOrder', customer: cust('0521000009'), lines: [TABLE] });
  P2({ action: 'sendOrderEmail', orderId: r.orderId, phone: '0521000009' });
  const x2 = await openXlsx(env2.mails[0].attachments[0]);
  ok(!/<img/.test(env2.mails[0].htmlBody) && !env2.mails[0].inlineImages && env2.mails[0].htmlBody.includes(title), 'תמונות לא נטענו: המייל נשלח בלי תמונות ובלי קישורים');
  ok(x2.ws.getImages().length === 0 && rowsText(x2.ws).some(rw => rw[3] === title), 'Excel נבנה ונפתח גם בלי תמונות');

  console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל בדיקות הגיבוי עברו');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
