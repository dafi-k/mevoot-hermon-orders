// בדיקת שליחת הזמנה מקצה לקצה: דפדפן אמיתי + קוד השרת האמיתי (apps-script.js) בתוך הדמיה של Google.
// דורש: node serve.js 8080 ברקע. הרצה: node test-submit.js [תיקייה-לצילומים]
const { chromium } = require('playwright-core');
const ExcelJS = require('exceljs');
const path = require('path'), fs = require('fs');
const { makeEnv } = require('./mock-gas');
const OUT = process.argv[2] || path.join(__dirname, 'shots');
fs.mkdirSync(OUT, { recursive: true });
const SITE = 'http://localhost:8080/';
let fails = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fails++; };
const col = (sh, n) => sh.data[0].indexOf(n);
const cell = (sh, id, n) => { const r = sh.data.find(r => r[0] === id); return r ? r[col(sh, n)] : undefined; };

const CUST = phone => ({ zimmer: 'צימר הגפן', town: 'מסעדה', contact: 'רות לוי', phone, email: 'ruth@example.com', address: 'הגפן 3, מסעדה', day: '2.12.26' });

// פותח דף עם שרת מדומה. hooks.beforePost(body) יכול להחזיר 'abort' / 'abort-after' כדי לדמות ניתוק
async function openPage(browser, env, { test = true, cart, phone = '0521111111', hooks = {}, width = 390 } = {}) {
  const ctx = await browser.newContext({ viewport: { width, height: 844 }, acceptDownloads: true, locale: 'he-IL' });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route(/script\.google\.com/, async route => {
    const req = route.request();
    if (req.method() === 'GET') {
      const out = env.ctx.doGet({ parameter: { action: 'catalog' } }).s;
      return route.fulfill({ status: 200, contentType: 'application/json', body: out });
    }
    const body = req.postData();
    const h = hooks.beforePost && hooks.beforePost(JSON.parse(body));
    if (h === 'abort') return route.abort('connectionreset');
    const out = env.ctx.doPost({ postData: { contents: body } }).s;
    if (h === 'abort-after') return route.abort('connectionreset'); // ההזמנה נשמרה אבל התשובה לא הגיעה
    return route.fulfill({ status: 200, contentType: 'application/json', body: out });
  });
  await page.goto(SITE);
  await page.evaluate(([c, cart]) => {
    sessionStorage.clear();
    localStorage.clear();
    localStorage.setItem('mh-orders-v1', JSON.stringify({ customer: c, cart, screen: 4, tab: null }));
  }, [CUST(phone), cart]);
  const code = env.ctx.getSettings_()['קוד בדיקה'];
  await page.goto(SITE + (test ? '?test=' + code : ''));
  await page.waitForFunction(() => S.live === true, null, { timeout: 15000 });
  await page.waitForSelector('#s4.active');
  return { ctx, page, errors };
}
const L = (id, qty, lid) => ({ lid: lid || Math.floor(Math.random() * 1e9), type: 'item', id, qty });
const T = (leg, top, qty) => ({ lid: Math.floor(Math.random() * 1e9), type: 'table', leg, top, qty });

(async () => {
  const browser = await chromium.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe' });

  // ===== 1. סגור להזמנות בלי קוד בדיקה
  console.log('--- סגור להזמנות');
  let env = makeEnv();
  env.ctx.setSetting_('נמעני מייל צוות', 'staff1@contrast.test, staff2@contrast.test');
  let { ctx, page, errors } = await openPage(browser, env, { test: false, cart: [L('CH-P1', 1)] });
  ok(await page.locator('#sendBtn').isDisabled(), 'בלי קוד בדיקה: כפתור השליחה נעול');
  ok((await page.textContent('#sumBanner')).includes('סגורה'), 'מוצגת הודעה שהשליחה סגורה');
  await ctx.close();

  // ===== 2. הזמנת בדיקה מוצלחת
  console.log('--- הזמנת בדיקה מוצלחת');
  ({ ctx, page, errors } = await openPage(browser, env, { cart: [L('AR-A6', 2), T('LG-LP1', 'TOP-P11', 1), L('OD27', 1), L('DC4', 3)] }));
  ok((await page.textContent('#sumBanner')).includes('מצב בדיקה'), 'מצב בדיקה מוצג בסיכום');
  ok(await page.locator('#sendBtn').isDisabled(), 'מעל התקציב: חסום עד אישור');
  await page.check('#overOk');
  const downloads = [];
  page.on('download', d => downloads.push(d));
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  await page.waitForFunction(() => S.lastOrder && (S.lastOrder.mail === 'sent' || S.lastOrder.mail === 'failed'), null, { timeout: 30000 });
  const order = await page.evaluate(() => ({ id: S.lastOrder.orderId, mail: S.lastOrder.mail, total: S.lastOrder.total }));
  ok(/^TEST-/.test(order.id), 'מסך אישור עם מספר הזמנה ' + order.id);
  const expTotal = 899 * 2 + 799 + 3599 + 5600 + 549 * 3;
  ok(order.total === expTotal, `סה"כ מהשרת ${order.total} (צפוי ${expTotal})`);
  ok(order.mail === 'sent' && (await page.textContent('#okMail')).includes('נשלח'), 'המייל נשלח ומוצג אישור');
  await page.waitForTimeout(600);
  await page.screenshot({ path: path.join(OUT, 'submit-1-success.png'), fullPage: true });
  ok(downloads.length === 2, `ירדו 2 קבצים (${downloads.map(d => d.suggestedFilename()).join(', ')})`);
  const xl = downloads.find(d => d.suggestedFilename().endsWith('.xlsx'));
  const pdf = downloads.find(d => d.suggestedFilename().endsWith('.pdf'));
  if (xl) {
    const p = path.join(OUT, 'order.xlsx'); await xl.saveAs(p);
    const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p);
    const ws = wb.worksheets[0];
    ok(ws.views[0].rightToLeft === true, 'Excel: גיליון מימין לשמאל');
    ok(ws.getImages().length === 4, `Excel: ${ws.getImages().length} תמונות מוקטנות מוטמעות (4 שורות)`);
    let found = 0, tRow = 0;
    ws.eachRow((r, n) => { if (r.values.some(v => /^שולחן: רגל LP\d \+ פלטה .+ קוטר \d+$/.test(v))) { found++; tRow = n; } if (r.values.includes(expTotal)) found++; });
    ok(found >= 2, 'Excel: שורת שולחן ("שולחן: רגל ... + פלטה ... קוטר ...") וסה"כ נכון');
    const ti = ws.getImages().find(im => im.range.tl.nativeRow === tRow - 1);
    ok(ti && ti.range.ext.width > ti.range.ext.height * 2, 'Excel: בשורת השולחן תמונת רגל + תמונת פלטה זו לצד זו');
  }
  if (pdf) {
    const p = path.join(OUT, 'order.pdf'); await pdf.saveAs(p);
    const buf = fs.readFileSync(p);
    ok(buf.slice(0, 4).toString() === '%PDF' && buf.length > 20000, `PDF תקין (${Math.round(buf.length / 1024)}KB)`);
  }
  const mail = env.mails[env.mails.length - 1];
  ok(mail && mail.to === 'staff1@contrast.test,staff2@contrast.test' && mail.cc === 'ruth@example.com' && /\[בדיקה\]/.test(mail.subject) && mail.attachments.length === 1, 'מייל: צוות + CC ללקוח, מסומן [בדיקה], Excel מצורף');
  const st = await page.evaluate(() => ({ stored: localStorage.getItem('mh-orders-v1'), cart: S.cart.length, cust: Object.keys(S.customer).length, field: document.querySelector('#f-zimmer').value }));
  ok(!st.stored && st.cart === 0 && st.cust === 0 && st.field === '', 'אחרי הצלחה: העגלה ופרטי הרישום נמחקו מהמכשיר');
  await page.click('#newCust');
  ok(await page.locator('#s1.active').isVisible(), '"סיום - לקוח חדש" חוזר לרישום');
  ok(await page.evaluate(() => S.testCode === '' && !sessionStorage.getItem('mh-test')), '"לקוח חדש" מבטל גם את מצב הבדיקה במכשיר');
  ok(errors.length === 0, 'אין שגיאות JavaScript' + (errors.length ? ': ' + errors.join(' | ') : ''));
  await ctx.close();

  // ===== 3. הזמנה כפולה מאותו טלפון
  console.log('--- הזמנה אחת לטלפון');
  ({ ctx, page } = await openPage(browser, env, { cart: [L('CH-P1', 1)] }));
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.sumKind === 'dup', null, { timeout: 15000 });
  ok((await page.textContent('#sumBanner')).includes('כבר נשלחה הזמנה'), 'אותו טלפון: "כבר נשלחה הזמנה ממספר הטלפון הזה"');
  ok(await page.evaluate(() => S.cart.length) === 1, 'העגלה נשמרה');
  await ctx.close();

  // ===== 4. חוסר מלאי - מישהו הזמין בינתיים
  console.log('--- חוסר מלאי בזמן השליחה');
  ({ ctx, page } = await openPage(browser, env, { phone: '0522222222', cart: [L('DC4', 5), L('CH-P2', 1), T('LG-LP2', 'TOP-P1', 2)] }));
  const P = env.sheets['מוצרים'], rem = col(P, 'יתרה');
  P.data.find(r => r[0] === 'DC4')[rem] = 2;    // נשארו 2
  P.data.find(r => r[0] === 'LG-LP2')[rem] = 0; // הרגל אזלה
  const ordersBefore = env.sheets['הזמנות'].data.length;
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.sumKind === 'shortage', null, { timeout: 15000 });
  const txt = await page.textContent('#s4');
  ok(txt.includes('חלק מהפריטים כבר לא זמינים'), 'מוצגת הודעת חוסר');
  ok(await page.locator('.line.short').count() === 2, 'בדיוק 2 שורות מסומנות באדום (כיסא DC4 והשולחן)');
  ok(txt.includes('הכמות שביקשת כבר לא זמינה במלאי') && txt.includes('אזל מהמלאי - יש להסיר'), 'הודעות לפי שורה: חלקי / אזל');
  ok(!/נותרו|זמינות: \d|available/.test(txt), 'לא מוצגים מספרי מלאי ללקוח');
  ok(env.sheets['הזמנות'].data.length === ordersBefore, 'בחוסר: שום דבר לא נכתב לגיליון');
  await page.waitForTimeout(500);
  await page.screenshot({ path: path.join(OUT, 'submit-2-shortage.png'), fullPage: true });
  await page.click('[data-fixshort]');
  const cart = await page.evaluate(() => S.cart.map(l => [l.id || 'table', l.qty]));
  ok(JSON.stringify(cart) === JSON.stringify([['DC4', 2], ['CH-P2', 1]]), 'תיקון אוטומטי: DC4 ירד ל-2, השולחן הוסר: ' + JSON.stringify(cart));
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  ok(cell(P, 'DC4', 'יתרה') === 0, 'אחרי תיקון ושליחה: ההזמנה נקלטה, יתרת DC4 = 0');
  await ctx.close();

  // ===== 5. מערכת עמוסה
  console.log('--- נעילה תפוסה');
  ({ ctx, page } = await openPage(browser, env, { phone: '0523333333', cart: [L('CH-P1', 1)] }));
  env.setLockBusy(true);
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.sumKind === 'busy', null, { timeout: 15000 });
  ok((await page.textContent('#sumBanner')).includes('המערכת עמוסה, נסו שוב בעוד רגע'), 'נעילה תפוסה: "המערכת עמוסה, נסו שוב בעוד רגע"');
  ok(await page.evaluate(() => S.cart.length) === 1 && !(await page.locator('#sendBtn').isDisabled()), 'העגלה נשמרה ואפשר לנסות שוב');
  env.setLockBusy(false);
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  ok(true, 'ניסיון חוזר הצליח');
  await ctx.close();

  // ===== 6. ניתוק: ההזמנה נשמרה אבל התשובה לא הגיעה
  console.log('--- ניתוק באמצע');
  let first = true;
  ({ ctx, page } = await openPage(browser, env, { phone: '0524444444', cart: [L('CH-P1', 1)],
    hooks: { beforePost: b => (b.action === 'submitOrder' && first ? (first = false, 'abort-after') : null) } }));
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.sumKind === 'net', null, { timeout: 15000 });
  ok((await page.textContent('#sumBanner')).includes('לא התקבלה תשובה'), 'ניתוק: "לא התקבלה תשובה מהשרת"');
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.sumKind === 'dup', null, { timeout: 15000 });
  ok((await page.textContent('#sumBanner')).includes('כבר נקלטה'), 'שליחה חוזרת אחרי ניתוק: "ההזמנה כבר נקלטה" - בלי הזמנה כפולה');
  ok(env.sheets['הזמנות'].data.filter(r => r[col(env.sheets['הזמנות'], 'טלפון')] === '0524444444').length === 1, 'בגיליון יש הזמנה אחת בלבד לטלפון הזה');
  await ctx.close();

  // ===== 7. כשל במייל לא מבטל הזמנה
  console.log('--- כשל במייל');
  const env2 = makeEnv({ mailThrows: 'Service invoked too many times for one day: email.' });
  env2.ctx.setSetting_('נמעני מייל צוות', 'staff1@contrast.test');
  ({ ctx, page } = await openPage(browser, env2, { phone: '0525555555', cart: [L('CH-P1', 1)] }));
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  await page.waitForFunction(() => S.lastOrder.mail === 'failed', null, { timeout: 30000 });
  const id2 = await page.evaluate(() => S.lastOrder.orderId);
  ok((await page.textContent('#okMail')).includes('המייל לא נשלח כרגע. ההזמנה נקלטה'), 'מסך אישור: "המייל לא נשלח כרגע. ההזמנה נקלטה"');
  const O2 = env2.sheets['הזמנות'];
  ok(cell(O2, id2, 'סטטוס') === 'חדשה' && /^לא נשלח: /.test(cell(O2, id2, 'סטטוס מייל')) && /drive/.test(cell(O2, id2, 'קובץ Excel')), 'בגיליון: הזמנה "חדשה", "לא נשלח: ..." וה-Excel שמור ב-Drive');
  await page.waitForTimeout(500);
  await page.screenshot({ path: path.join(OUT, 'submit-3-mailfail.png'), fullPage: true });
  await ctx.close();

  await browser.close();
  console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל בדיקות השליחה עברו');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
