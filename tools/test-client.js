// בדיקה אוטומטית של צד הלקוח (שלב 3). דורש: node serve.js 8080 ברקע.
// הרצה: node test-client.js [תיקייה-לצילומי-מסך]
const { chromium } = require('playwright-core');
const path = require('path');
const OUT = process.argv[2] || path.join(__dirname, 'shots');
require('fs').mkdirSync(OUT, { recursive: true });
const URL = process.env.SITE_URL || 'http://localhost:8080/';
let fails = 0;
const ok = (cond, msg) => { console.log((cond ? 'PASS ' : 'FAIL ') + msg); if (!cond) fails++; };
const shot = async (page, name, full) => { await page.waitForTimeout(900); await page.screenshot({ path: path.join(OUT, name + '.png'), fullPage: !!full }); };

// בדיקות פריסה שחוזרות בכל מסך
async function layoutChecks(page, label) {
  const r = await page.evaluate(() => {
    const vw = document.documentElement.clientWidth;
    const overflow = document.documentElement.scrollWidth > vw + 1;
    // טקסט קטן מדי (פחות מ-12px) בכל אלמנט גלוי עם טקסט
    const small = [...document.querySelectorAll('body *')].filter(el => {
      if (!el.offsetParent || !el.childNodes.length) return false;
      const hasText = [...el.childNodes].some(n => n.nodeType === 3 && n.textContent.trim());
      return hasText && parseFloat(getComputedStyle(el).fontSize) < 12;
    }).map(el => el.textContent.trim().slice(0, 20));
    // כפתורים קטנים מדי ללחיצה באצבע (חוץ מקישורי טקסט)
    const tiny = [...document.querySelectorAll('button')].filter(b => b.offsetParent && !b.classList.contains('linkbtn'))
      .filter(b => { const x = b.getBoundingClientRect(); return x.height < 36 || x.width < 36; }).map(b => b.textContent.trim().slice(0, 15) || b.getAttribute('aria-label'));
    return { overflow, small, tiny };
  });
  ok(!r.overflow, `${label}: אין גלילה אופקית של הדף`);
  ok(r.small.length === 0, `${label}: אין טקסט קטן מ-12px` + (r.small.length ? ' -> ' + r.small.slice(0, 5).join(' | ') : ''));
  ok(r.tiny.length === 0, `${label}: כל הכפתורים לפחות 36px` + (r.tiny.length ? ' -> ' + r.tiny.slice(0, 5).join(' | ') : ''));
}
async function noStockNumbers(page, label) {
  const txt = await page.evaluate(() => document.body.innerText);
  ok(!/נותרו|במלאי:\s*\d|אפשר עד|אפשר להזמין עד/.test(txt), `${label}: לא מוצגות ללקוח כמויות במלאי`);
}

async function run(browser, width) {
  const W = `[${width}px]`;
  const ctx = await browser.newContext({ viewport: { width, height: width < 500 ? 812 : 1024 }, deviceScaleFactor: 2, isMobile: true, hasTouch: true, locale: 'he-IL' });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });

  let release;
  const gate = new Promise(r => { release = r; });
  await page.route(/script\.google\.com/, async route => { await gate; await route.continue(); });

  await page.goto(URL);
  // --- רישום
  await page.click('#toGuide');
  ok(await page.locator('.inp.invalid').count() === 6, `${W} שדות ריקים מסומנים`);
  await page.fill('#f-zimmer', 'צימר בדיקה');
  await page.fill('#f-town', 'מסעדה');
  await page.fill('#f-contact', 'ישראלה ישראלי');
  await page.fill('#f-phone', '052-1234567');
  await page.fill('#f-email', 'test@example.com');
  await page.fill('#f-address', 'הרחוב 1, מסעדה');
  await page.click('.day[data-v="2.12.26"]');
  await shot(page, `${width}-1-register`, true);
  await layoutChecks(page, `${W} רישום`);
  await page.click('#toGuide');
  await shot(page, `${width}-2-guide`, true);
  await layoutChecks(page, `${W} הנחיות`);
  await page.click('#s2 [data-go="3"]');

  // --- לפני מלאי חי
  await page.waitForSelector('.card');
  ok(await page.locator('.add-btn').count() === 0 && await page.locator('.stepper.wait').count() > 0, `${W} לפני מלאי חי: "בודק מלאי..." והכפתורים נעולים`);
  release();
  await page.waitForFunction(() => document.querySelector('#stockPill').textContent.includes('מעודכן'), null, { timeout: 30000 });
  await shot(page, `${width}-3-catalog`);
  await noStockNumbers(page, `${W} קטלוג`);
  await layoutChecks(page, `${W} קטלוג`);

  // --- פס הקטגוריות
  const tabInfo = await page.evaluate(() => {
    const tabs = document.querySelector('#tabs'), r = tabs.getBoundingClientRect();
    const all = [...tabs.querySelectorAll('.tab')];
    const visible = all.filter(t => { const x = t.getBoundingClientRect(); return x.left >= r.left - 1 && x.right <= r.right + 1; }).length;
    return { total: all.length, visible, scrollable: tabs.scrollWidth > tabs.clientWidth };
  });
  if (width >= 700) {
    ok(tabInfo.visible === tabInfo.total && !tabInfo.scrollable, `${W} כל ${tabInfo.total} הקטגוריות גלויות בבת אחת`);
  } else {
    ok(tabInfo.scrollable, `${W} פס הקטגוריות נגלל אופקית`);
    ok(await page.locator('#edgeEnd.show').count() === 1, `${W} רמז בקצה שיש עוד קטגוריות`);
    // בחירת הקטגוריה האחרונה - צריכה להיגלל לתוך התצוגה
    await page.locator('#edgeEnd button').click();
    await page.waitForTimeout(500);
    ok(await page.locator('#edgeStart.show').count() === 1, `${W} אחרי גלילה מופיע רמז גם בקצה השני`);
    await page.evaluate(() => { S.tab = 'misc'; renderTabs(); renderCatalog(); });
    await page.waitForTimeout(600);
    const centered = await page.evaluate(() => {
      const tabs = document.querySelector('#tabs').getBoundingClientRect(), a = document.querySelector('.tab.active').getBoundingClientRect();
      return a.left >= tabs.left - 1 && a.right <= tabs.right + 1;
    });
    ok(centered, `${W} הקטגוריה הפעילה נגללת אוטומטית לתוך התצוגה`);
    await shot(page, `${width}-4-tabs-scrolled`);
  }
  // דבוק למעלה
  await page.click('.tab[data-tab="plastic"]');
  await page.mouse.wheel(0, 1500);
  await page.evaluate(() => window.scrollBy(0, 1500));
  await page.waitForTimeout(400);
  const stick = await page.evaluate(() => {
    const hdr = document.querySelector('.topbar').getBoundingClientRect().bottom, tb = document.querySelector('#tabbar').getBoundingClientRect();
    return Math.abs(tb.top - hdr) < 3;
  });
  ok(stick, `${W} פס הקטגוריות נשאר דבוק מתחת לכותרת בגלילה`);
  await shot(page, `${width}-5-sticky`);
  await page.evaluate(() => window.scrollTo(0, 0));

  // --- מקסימום מלאי בלי מספרים
  await page.click('.tab[data-tab="armchair"]');
  const a6 = page.locator('.card', { hasText: 'A6' });
  const a6stock = await page.evaluate(() => S.byId['AR-A6'].stock);
  await a6.locator('.add-btn').click();
  for (let i = 1; i < a6stock; i++) await a6.locator('[data-inc]').click();
  ok(await a6.locator('[data-inc]').isDisabled(), `${W} כפתור + ננעל בכמות המקסימלית`);
  ok((await a6.innerText()).includes('זו הכמות המקסימלית הזמינה'), `${W} מופיעה ההודעה "זו הכמות המקסימלית הזמינה"`);
  ok(await page.locator('.toast', { hasText: 'זו הכמות המקסימלית הזמינה' }).count() === 1, `${W} הודעה קופצת בהגעה למקסימום`);
  // פריט שאזל
  await page.evaluate(() => { S.byId['AR-A5'].stock = 0; renderCatalog(); });
  const a5 = page.locator('.card', { hasText: 'A5' });
  ok((await a5.innerText()).includes('אזל מהמלאי') && await a5.locator('.add-btn').isDisabled(), `${W} פריט שאזל: "אזל מהמלאי" והכפתור נעול`);
  await shot(page, `${width}-6-armchairs`);
  await noStockNumbers(page, `${W} כורסאות`);

  // --- שולחן
  await page.click('.tab[data-tab="tables"]');
  await page.click('[data-build]');
  await page.click('[data-leg="LG-LP1"]');
  await page.click('[data-tone="אגוז"]');
  await page.click('[data-dia="120"]');
  for (let i = 0; i < 20; i++) { const b = page.locator('[data-tq="1"]'); if (await b.isDisabled()) break; await b.click(); }
  ok((await page.locator('.modal').innerText()).includes('זו הכמות המקסימלית הזמינה'), `${W} מודאל שולחן: הודעת מקסימום בלי מספר`);
  await noStockNumbers(page, `${W} מודאל שולחן`);
  await shot(page, `${width}-7-table-modal`);
  await layoutChecks(page, `${W} מודאל שולחן`);
  await page.click('[data-tsave]');

  // --- סיכום
  await page.click('#toSummary');
  await page.check('#overOk');
  await shot(page, `${width}-8-summary`, true);
  await noStockNumbers(page, `${W} סיכום`);
  await layoutChecks(page, `${W} סיכום`);

  // --- לקוח חדש
  await page.click('#restartBtn');
  ok(await page.locator('[data-reset]').isVisible(), `${W} "לקוח חדש" מבקש אישור`);
  await shot(page, `${width}-9-reset-dialog`);
  await page.click('[data-reset]');
  const after = await page.evaluate(() => ({ cart: S.cart.length, name: document.querySelector('#f-zimmer').value, screen: S.screen, stored: localStorage.getItem('mh-orders-v1') }));
  ok(after.cart === 0 && after.name === '' && after.screen === 1 && !after.stored, `${W} "לקוח חדש" מנקה עגלה, פרטים ואחסון וחוזר לרישום`);
  await page.reload();
  ok(await page.evaluate(() => S.cart.length === 0 && !S.customer.zimmer), `${W} אחרי רענון - המכשיר נקי`);

  ok(errors.length === 0, `${W} אין שגיאות JavaScript` + (errors.length ? ': ' + errors.join(' | ') : ''));
  await ctx.close();
}

(async () => {
  const browser = await chromium.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe' });
  for (const w of [375, 768]) { console.log(`\n===== ${w}px =====`); await run(browser, w); }
  await browser.close();
  console.log(fails ? `\n${fails} בדיקות נכשלו` : '\nכל הבדיקות עברו');
  process.exit(fails ? 1 : 0);
})().catch(e => { console.error(e.message.split('\n')[0]); process.exit(1); });
