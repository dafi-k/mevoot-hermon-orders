// בדיקת מסכי הניהול מקצה לקצה: דפדפן + קוד השרת האמיתי בתוך הדמיה של Google.
// דורש: node serve.js 8080 ברקע. הרצה: node test-admin.js [תיקייה-לצילומים]
const { chromium } = require('playwright-core');
const path = require('path'), fs = require('fs');
const { makeEnv } = require('./mock-gas');
const OUT = process.argv[2] || path.join(__dirname, 'shots');
fs.mkdirSync(OUT, { recursive: true });
let fails = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fails++; };
const col = (sh, n) => sh.data[0].indexOf(n);
const cell = (sh, id, n) => { const r = sh.data.find(r => r[0] === id); return r ? r[col(sh, n)] : undefined; };
const cust = (phone, zimmer) => ({ zimmer, town: 'מסעדה', contact: 'רות', phone, email: 'r@example.com', address: 'הגפן 3, מסעדה', day: '1.12.26' });

(async () => {
  // הכנת נתונים: 3 הזמנות (אחת בדיקה), אחת עם מייל שנכשל
  const env = makeEnv({ mailThrows: 'quota' });
  env.props.ADMIN_PASSWORD = 'Sod-2026!';
  env.ctx.setSetting_('נמעני מייל צוות', 'a@contrast.test');
  const code = env.ctx.getSettings_()['קוד בדיקה'];
  const P = b => JSON.parse(env.ctx.doPost({ postData: { contents: JSON.stringify(b) } }).s);
  const o1 = P({ action: 'submitOrder', testCode: code, customer: cust('0521000001', 'צימר הגפן'), overOk: true, lines: [{ type: 'item', id: 'AR-A6', qty: 2 }, { type: 'table', leg: 'LG-LP1', top: 'TOP-P11', qty: 2 }] });
  env.ctx.setSetting_('פתוח להזמנות', 'כן');
  const o2 = P({ action: 'submitOrder', customer: cust('0521000002', 'צימר האלון'), lines: [{ type: 'item', id: 'CH-P1', qty: 4 }, { type: 'item', id: 'OD27', qty: 1 }] });
  P({ action: 'sendOrderEmail', orderId: o2.orderId, phone: '0521000002', xlsxBase64: 'AAAA' }); // נכשל (quota)
  ok(o1.ok && o2.ok, 'הוכנו הזמנות: ' + o1.orderId + ', ' + o2.orderId);

  const browser = await chromium.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe' });
  for (const width of [1280, 390]) {
    console.log(`\n===== ${width}px =====`);
    const W = `[${width}]`;
    const ctx = await browser.newContext({ viewport: { width, height: width > 500 ? 900 : 844 }, locale: 'he-IL' });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.route(/script\.google\.com/, async route => {
      const out = env.ctx.doPost({ postData: { contents: route.request().postData() } }).s;
      route.fulfill({ status: 200, contentType: 'application/json', body: out });
    });
    await page.goto('http://localhost:8080/admin.html');

    // כניסה
    await page.fill('#lName', 'דפי'); await page.fill('#lPass', 'wrong'); await page.click('#lBtn');
    await page.waitForFunction(() => document.querySelector('#lErr').textContent.length > 0);
    ok((await page.textContent('#lErr')) === 'סיסמה שגויה', `${W} סיסמה שגויה: הודעה ואין כניסה`);
    await page.fill('#lPass', 'Sod-2026!'); await page.click('#lBtn');
    await page.waitForSelector('#appView:not(.hidden) .kpi');
    ok((await page.textContent('#meta')).includes('דפי'), `${W} כניסה מוצלחת, שם העובד בכותרת`);
    const k = await page.textContent('#kpis');
    ok(k.includes('מיילים שלא נשלחו'), `${W} מדדים מוצגים`);
    await page.waitForTimeout(400);
    await page.screenshot({ path: path.join(OUT, `admin-${width}-1-stock.png`) });

    // מלאי: סינון "אזל"
    await page.click('[data-stock="out"]');
    const outRows = await page.locator('#view tbody tr').count();
    if (width === 1280) ok(outRows >= 1 && (await page.textContent('#view')).includes('A6'), `${W} סינון "אזל" מציג את A6 (אזל אחרי ההזמנה)`);
    await page.click('[data-stock="all"]');
    await page.fill('[data-f="q"]', 'LP1');
    ok(await page.locator('#view tbody tr').count() === 1, `${W} חיפוש LP1 מוצא שורה אחת`);
    const lp1Row = await page.textContent('#view tbody tr');
    if (width === 1280) ok(/5/.test(lp1Row) && /3/.test(lp1Row), `${W} LP1: התחלתי 5, יתרה 3`);

    if (width === 1280) {
      // כניסת סחורה מהשורה
      await page.click('[data-recv="LG-LP1"]');
      await page.fill('#rQty', '4'); await page.fill('#rNote', 'משלוח בדיקה');
      await page.click('#rBtn');
      await page.waitForFunction(() => document.querySelector('.toast') && document.querySelector('.toast').textContent.includes('יתרה חדשה'));
      ok(cell(env.sheets['מוצרים'], 'LG-LP1', 'יתרה') === 7, `${W} כניסת סחורה: LP1 3→7 בגיליון`);
      await page.waitForFunction(() => document.querySelector('#view').textContent.includes('משלוח בדיקה'));
      ok(true, `${W} הכניסה מופיעה ב"כניסות אחרונות"`);
      await page.waitForTimeout(300);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-2-receive.png`) });

      // תיקון ידני
      await page.click('[data-tab="stock"]');
      await page.fill('[data-f="q"]', 'DC1');
      await page.click('[data-adj="DC1"]');
      await page.fill('#aRem', '10');
      await page.click('#aBtn');
      ok((await page.textContent('.toast')).includes('סיבה'), `${W} תיקון בלי סיבה: נחסם`);
      await page.fill('#aNote', 'ספירת מלאי');
      await page.click('#aBtn');
      await page.waitForFunction(() => !document.querySelector('#modalRoot').innerHTML);
      ok(cell(env.sheets['מוצרים'], 'DC1', 'יתרה') === 10, `${W} תיקון ידני: DC1 12→10`);

      // הזמנות: שליחה חוזרת + ביטול
      await page.click('[data-tab="orders"]');
      ok(await page.locator('.order').count() === 2, `${W} 2 הזמנות ברשימה`);
      ok(await page.locator('[data-resend]').count() === 2, `${W} "שלח שוב" מופיע בהזמנות שהמייל שלהן לא נשלח`);
      env.ctx.MailApp.sendEmail = o => env.mails.push(o);
      await page.click(`[data-resend="${o2.orderId}"]`);
      await page.waitForFunction(() => document.querySelector('.toast') && document.querySelector('.toast').textContent.includes('המייל נשלח'));
      ok(/^נשלח /.test(cell(env.sheets['הזמנות'], o2.orderId, 'סטטוס מייל')), `${W} שליחה חוזרת: הסטטוס בגיליון "נשלח"`);
      await page.waitForTimeout(800);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-3-orders.png`), fullPage: true });

      await page.selectOption(`select[data-status="${o1.orderId}"]`, 'בוטלה');
      await page.waitForSelector('#sBtn');
      ok((await page.textContent('.modal')).includes('הכמויות יחזרו למלאי'), `${W} ביטול: חלון אישור עם הסבר`);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-4-cancel.png`) });
      await page.fill('#sNote', 'הזמנת בדיקה');
      await page.click('#sBtn');
      await page.waitForFunction(() => document.querySelector('.toast') && document.querySelector('.toast').textContent.includes('חזרו למלאי'));
      const PR = env.sheets['מוצרים'];
      ok(cell(PR, 'AR-A6', 'יתרה') === 2 && cell(PR, 'LG-LP1', 'יתרה') === 9 && cell(PR, 'TOP-P11', 'יתרה') === 4, `${W} ביטול: A6 0→2, LP1 7→9, P11 2→4`);
      ok(cell(env.sheets['הזמנות'], o1.orderId, 'סטטוס') === 'בוטלה', `${W} ביטול: סטטוס "בוטלה" בגיליון`);

      // סיכומים
      await page.click('[data-tab="totals"]');
      await page.uncheck('#incTest');
      const tot = await page.textContent('#view');
      ok(tot.includes('P1') && !tot.includes('A6'), `${W} ריכוז כמויות: רק הזמנות פעילות (בלי ההזמנה שבוטלה)`);
      await page.click('[data-tab="zimmers"]');
      const z = await page.textContent('#view');
      ok(z.includes('צימר האלון') && !z.includes('צימר הגפן'), `${W} סיכום לפי צימר: בלי הזמנות שבוטלו`);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-5-zimmers.png`) });
      await page.click('[data-tab="moves"]');
      await page.click('[data-mtype="ביטול"]');
      ok(await page.locator('#view tbody tr').count() === 3, `${W} יומן: 3 תנועות ביטול`);
      const mv = await page.textContent('#view');
      ok(/\d\d\/\d\d\/\d{4} \d\d:\d\d/.test(mv) && mv.includes('דפי'), `${W} יומן: תאריך ושעה + שם העובד`);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-6-moves.png`) });
    } else {
      const over = await page.evaluate(() => document.documentElement.scrollWidth > document.documentElement.clientWidth + 1);
      ok(!over, `${W} אין גלילה אופקית של הדף בטלפון`);
      await page.click('[data-tab="orders"]');
      await page.waitForTimeout(300);
      await page.screenshot({ path: path.join(OUT, `admin-${width}-orders.png`), fullPage: true });
    }

    // רענון שומר כניסה, יציאה מוחקת
    await page.reload();
    await page.waitForSelector('#appView:not(.hidden) .kpi');
    ok(true, `${W} רענון: נשאר מחובר`);
    await page.click('#logoutBtn');
    ok(await page.locator('#loginView:not(.hidden)').isVisible() && !(await page.evaluate(() => sessionStorage.getItem('mh-admin-token'))), `${W} יציאה: חזרה למסך כניסה, האסימון נמחק`);
    ok(errors.length === 0, `${W} אין שגיאות JavaScript` + (errors.length ? ': ' + errors.join(' | ') : ''));
    await ctx.close();
  }
  await browser.close();
  console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל בדיקות הניהול עברו');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
