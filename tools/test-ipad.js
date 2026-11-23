// בדיקת שליחה מאייפד: WebKit (המנוע של Safari) בגודל ובזיהוי של אייפד, מול קוד השרת האמיתי בהדמיה של Google.
// מדמה את התקלה: בספארי באייפד הורדת קובץ (Excel/PDF) נפתחת במקום הדף - וכל מה שהדף עוד היה אמור לשלוח, לא נשלח.
// דורש: node serve.js 8080 ברקע, ו-WebKit מותקן (npx playwright-core install webkit). הרצה: node test-ipad.js
const { webkit, chromium, devices } = require('playwright-core');
const path = require('path'), fs = require('fs'), { execSync } = require('child_process');
const { makeEnv } = require('./mock-gas');
const OUT = path.join(__dirname, 'shots');
fs.mkdirSync(OUT, { recursive: true });
const ROOT = path.join(__dirname, '..');
const SITE = 'http://localhost:8080/';
let fails = 0;
const ok = (c, m) => { console.log((c ? 'PASS ' : 'FAIL ') + m); if (!c) fails++; };
const col = (sh, n) => sh.data[0].indexOf(n);
const cell = (sh, id, n) => { const r = sh.data.find(r => r[0] === id); return r ? r[col(sh, n)] : undefined; };
const CUST = phone => ({ zimmer: 'צימר הגפן', town: 'מסעדה', contact: 'רות לוי', phone, email: 'ruth@example.com', address: 'הגפן 3, מסעדה', day: '2.12.26' });
const CART = () => [
  { lid: 11, type: 'item', id: 'CH-P1', qty: 2 },
  { lid: 12, type: 'table', leg: 'LG-LP2', top: 'TOP-P8', qty: 1 },
  { lid: 13, type: 'item', id: 'DC1', qty: 1 }
];
const IPAD = devices['iPad (gen 7)'];
// אייפד במצב "אתר למחשב" (ברירת המחדל בספארי באייפד): מזדהה כ-Mac, עם מסך מגע
// מה שאייפד אמיתי מדווח במצב הזה (WebKit ב-Windows מדווח Win32 בלי מגע)
const IPAD_DESKTOP_NAV = () => { Object.defineProperty(Navigator.prototype, 'platform', { get: () => 'MacIntel' }); Object.defineProperty(Navigator.prototype, 'maxTouchPoints', { get: () => 5 }); };
const IPAD_DESKTOP_UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15';

// הדמיית ספארי באייפד: לחיצה על קישור הורדה בלי לשונית חדשה = הקובץ נפתח במקום הדף
const SAFARI_DOWNLOAD = () => {
  const orig = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    if (this.hasAttribute('download')) {
      window.__dl(JSON.stringify({ name: this.download, target: this.target || '' }));
      if (this.target !== '_blank') location.href = 'about:blank';
      return;
    }
    return orig.call(this);
  };
};

async function openIpad(browser, env, { file = '', phone, ctxOpts = {}, hooks = {}, safari = true, init = null }) {
  const ctx = await browser.newContext({ ...IPAD, locale: 'he-IL', ...ctxOpts });
  const page = await ctx.newPage();
  const log = [], errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.exposeFunction('__dl', s => log.push({ t: Date.now(), type: 'download', ...JSON.parse(s) }));
  if (init) await page.addInitScript(init);
  await page.addInitScript(safari ? SAFARI_DOWNLOAD : () => { const o = HTMLAnchorElement.prototype.click; HTMLAnchorElement.prototype.click = function () { if (this.hasAttribute('download')) window.__dl(JSON.stringify({ name: this.download, target: this.target || '' })); return o.call(this); }; });
  await page.route(/script\.google\.com/, async route => {
    const req = route.request();
    if (req.method() === 'GET') return route.fulfill({ status: 200, contentType: 'application/json', body: env.ctx.doGet({ parameter: { action: 'catalog' } }).s });
    const body = JSON.parse(req.postData());
    log.push({ t: Date.now(), type: 'api', action: body.action, xlsx: !!body.xlsxBase64 });
    const h = hooks.beforePost && hooks.beforePost(body);
    if (h === 'abort') return route.abort('connectionreset');
    if (h === 'hang') return; // הבקשה "נתקעת" (הדף נסגר לפני שהגיעה)
    route.fulfill({ status: 200, contentType: 'application/json', body: env.ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s });
  });
  await page.goto(SITE + file);
  await page.evaluate(([c, cart]) => { sessionStorage.clear(); localStorage.clear();
    localStorage.setItem('mh-orders-v1', JSON.stringify({ customer: c, cart, screen: 4, tab: null })); }, [CUST(phone), CART()]);
  await page.goto(SITE + file);
  await page.waitForFunction(() => S.live === true, null, { timeout: 20000 });
  await page.waitForSelector('#s4.active');
  return { ctx, page, log, errors };
}
const newEnv = () => {
  const env = makeEnv();
  env.ctx.setSetting_('נמעני מייל צוות', 'staff1@contrast.test, staff2@contrast.test');
  env.ctx.setSetting_('פתוח להזמנות', 'כן');
  return env;
};
const orderOf = (env, phone) => { const r = env.sheets['הזמנות'].data.find(x => x[col(env.sheets['הזמנות'], 'טלפון')] === phone); return r && r[0]; };

(async () => {
  const browser = await webkit.launch();
  console.log('דפדפן: WebKit ' + browser.version() + ' | מכשיר: ' + IPAD.viewport.width + '×' + IPAD.viewport.height);

  // ===== 1. שחזור התקלה בגרסה הקודמת של האתר
  console.log('--- הגרסה הקודמת (לפני התיקון) באייפד');
  fs.writeFileSync(path.join(ROOT, '_ipad_old.html'), execSync('git show fd76c2b:index.html', { cwd: ROOT }));
  try {
    const env = newEnv();
    const { ctx, page, log } = await openIpad(browser, env, { file: '_ipad_old.html', phone: '0521000001' });
    await page.click('#sendBtn');
    await page.waitForTimeout(6000);
    const id = orderOf(env, '0521000001');
    ok(!!id && cell(env.sheets['הזמנות'], id, 'סטטוס מייל') === 'ממתין' && !log.some(x => x.action === 'sendOrderEmail'),
      `שוחזר: ההזמנה ${id} נשמרה, הדף "הוריד" Excel ונעלם - בקשת המייל לא יצאה, הסטטוס "ממתין", אין Excel ב-Drive`);
    await ctx.close();
  } finally { fs.unlinkSync(path.join(ROOT, '_ipad_old.html')); }

  // ===== 2. הגרסה המתוקנת
  console.log('--- הגרסה המתוקנת באייפד');
  let env = newEnv();
  let { ctx, page, log, errors } = await openIpad(browser, env, { phone: '0521000002' });
  ok(await page.evaluate(() => IS_IOS), 'האתר מזהה אייפד');
  // הצגת השולחן בסיכום
  const tl = page.locator('.line', { hasText: 'שולחן: רגל LP2' });
  ok(await tl.count() === 1, 'סיכום: "שולחן: רגל LP2 + פלטה אגוז קוטר 80"');
  ok(await tl.locator('.duo img').count() === 2, 'סיכום: תמונת רגל + תמונת פלטה');
  const tTxt = await tl.textContent();
  ok(/רגל LP2: ₪[\d,]+/.test(tTxt) && /פלטה אגוז 80 ס"מ: ₪[\d,]+/.test(tTxt) && /סה"כ לשולחן: ₪[\d,]+/.test(tTxt), 'סיכום: מחיר רגל, מחיר פלטה וסה"כ לשולחן');
  await page.screenshot({ path: path.join(OUT, 'ipad-1-summary.png') });
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  await page.waitForFunction(() => S.lastOrder && S.lastOrder.mail !== 'sending', null, { timeout: 30000 });
  await page.waitForTimeout(1500);
  const id = orderOf(env, '0521000002');
  const mailReq = log.find(x => x.action === 'sendOrderEmail');
  ok(!!mailReq && mailReq.xlsx, 'בקשת המייל יצאה, עם קובץ ה-Excel מהאתר');
  ok(!log.some(x => x.type === 'download'), 'באייפד אין הורדה אוטומטית שעלולה להעלים את הדף');
  ok(/^נשלח /.test(cell(env.sheets['הזמנות'], id, 'סטטוס מייל')) && /^https:\/\/drive/.test(cell(env.sheets['הזמנות'], id, 'קובץ Excel')), `בגיליון: ${id} "נשלח" וקובץ Excel ב-Drive`);
  ok(env.mails.length === 1 && env.mails[0].attachments.length === 1, 'מייל אחד, עם Excel מצורף');
  ok((await page.textContent('#okMail')).includes('נשלח'), 'מסך האישור: "אישור נשלח במייל"');
  ok(await page.locator('#dlX').isEnabled() && (await page.textContent('#okMail')).includes('בכפתורים'), 'כפתורי ההורדה פעילים, עם הסבר');
  await page.screenshot({ path: path.join(OUT, 'ipad-2-success.png') });
  await page.click('#dlX');
  await page.waitForTimeout(1500);
  const dl = log.find(x => x.type === 'download');
  ok(dl && dl.target === '_blank' && /\.xlsx$/.test(dl.name) && await page.locator('#s5.active').count() === 1, 'לחיצה על "הורדת Excel": נפתח בלשונית חדשה, דף האישור נשאר');
  await page.click('#dlP');
  await page.waitForTimeout(2500);
  ok(log.filter(x => x.type === 'download').some(x => /\.pdf$/.test(x.name) && x.target === '_blank'), 'הורדת PDF: גם בלשונית חדשה');
  ok(errors.length === 0, 'אין שגיאות JavaScript' + (errors.length ? ': ' + errors.join(' | ') : ''));
  await ctx.close();

  // ===== 3. אייפד במצב "אתר למחשב" (מזדהה כ-Mac)
  console.log('--- אייפד שמזדהה כ-Mac');
  env = newEnv();
  ({ ctx, page, log } = await openIpad(browser, env, { phone: '0521000003', ctxOpts: { userAgent: IPAD_DESKTOP_UA }, init: IPAD_DESKTOP_NAV }));
  const det = await page.evaluate(() => ({ ios: IS_IOS, ua: navigator.platform, tp: navigator.maxTouchPoints }));
  ok(det.ios, `זוהה כאייפד גם כשמזדהה כ-Mac (platform ${det.ua}, נקודות מגע ${det.tp})`);
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.lastOrder && S.lastOrder.mail !== 'sending', null, { timeout: 30000 });
  ok(env.mails.length === 1 && !log.some(x => x.type === 'download'), 'המייל נשלח, בלי הורדה אוטומטית');
  await ctx.close();

  // ===== 4. ניתוק ברשת בבקשת המייל -> ניסיון חוזר
  console.log('--- ניסיון חוזר');
  env = newEnv();
  let n = 0;
  ({ ctx, page, log } = await openIpad(browser, env, { phone: '0521000004', hooks: { beforePost: b => b.action === 'sendOrderEmail' && ++n === 1 ? 'abort' : '' } }));
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.lastOrder && S.lastOrder.mail !== 'sending', null, { timeout: 40000 });
  ok(log.filter(x => x.action === 'sendOrderEmail').length === 2 && env.mails.length === 1 && (await page.evaluate(() => S.lastOrder.mail)) === 'sent',
    'הבקשה הראשונה נקטעה, השנייה הצליחה - מייל אחד בלבד');
  await ctx.close();

  // ===== 5. הדף נסגר לפני שבקשת המייל הגיעה -> הגיבוי בשרת
  console.log('--- הדף נסגר באמצע - גיבוי בשרת');
  env = newEnv();
  env.ctx.installMailWatchdog();
  env.props.WATCHDOG_SINCE = String(Date.now() - 10 * 60000);
  ({ ctx, page, log } = await openIpad(browser, env, { phone: '0521000005', hooks: { beforePost: b => b.action === 'sendOrderEmail' ? 'hang' : '' } }));
  await page.click('#sendBtn');
  await page.waitForSelector('#s5.active', { timeout: 20000 });
  await ctx.close(); // הטאבלט נסגר / עבר לאפליקציה אחרת
  const id5 = orderOf(env, '0521000005');
  ok(cell(env.sheets['הזמנות'], id5, 'סטטוס מייל') === 'ממתין' && env.mails.length === 0, 'אחרי הסגירה: "ממתין", אין מייל');
  env.ctx.mailWatchdog();
  ok(cell(env.sheets['הזמנות'], id5, 'סטטוס מייל') === 'ממתין', 'פחות מ-2 דקות: הגיבוי עוד מחכה');
  const O = env.sheets['הזמנות'];
  O.data.find(r => r[0] === id5)[col(O, 'תאריך')] = new (env.ctx.now_().constructor)(Date.now() - 2.5 * 60000);
  env.ctx.mailWatchdog();
  ok(/^נשלח .*גיבוי אוטומטי/.test(cell(O, id5, 'סטטוס מייל')) && env.mails.length === 1 && env.mails[0].attachments.length === 1,
    'אחרי 2 דקות: השרת בנה Excel מהגיליון ושלח את המייל');
  await browser.close();

  // ===== 6. מחשב (Chrome): הבקשה לשרת לפני ההורדות, וההורדות עדיין אוטומטיות
  console.log('--- מחשב: סדר הפעולות');
  const chrome = await chromium.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe' });
  env = newEnv();
  ({ ctx, page, log } = await openIpad(chrome, env, { phone: '0521000006', safari: false, ctxOpts: { userAgent: undefined, isMobile: false, hasTouch: false, viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 } }));
  ok(!(await page.evaluate(() => IS_IOS)), 'מחשב לא מזוהה כאייפד');
  await page.click('#sendBtn');
  await page.waitForFunction(() => S.lastOrder && S.lastOrder.mail !== 'sending' && S.lastOrder.pdf, null, { timeout: 40000 });
  await page.waitForTimeout(500);
  const tMail = (log.find(x => x.action === 'sendOrderEmail') || {}).t, dls = log.filter(x => x.type === 'download');
  ok(tMail && dls.length === 2 && dls.every(d => d.t > tMail), 'בקשת המייל יצאה לפני שתי ההורדות (Excel ו-PDF)');
  await chrome.close();

  console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל בדיקות האייפד עברו');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e); try { fs.unlinkSync(path.join(ROOT, '_ipad_old.html')); } catch (x) {} process.exit(1); });
