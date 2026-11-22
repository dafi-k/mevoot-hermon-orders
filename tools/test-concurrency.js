// בדיקת הזמנות בו זמנית מול השרת האמיתי: N הזמנות באותו רגע על פריט עם יתרה קטנה.
// הרצה: node test-concurrency.js <מזהה-פריט> <קוד-בדיקה> [N]
// יוצר הזמנות בדיקה (TEST) בלי מייל. אחרי הבדיקה מבטלים אותן במסך הניהול.
const API = 'https://script.google.com/macros/s/AKfycbynZDBWrzSYNvW6bsdgW3KMWGpIV8maveVTKufv3JilsuD_LYD7d5OxK5h4vuzVMQnMIg/exec';
const [id, code, nArg] = process.argv.slice(2);
const N = Number(nArg) || 10;
if (!id || !code) { console.log('שימוש: node test-concurrency.js <מזהה-פריט> <קוד-בדיקה> [N]'); process.exit(1); }
const stockOf = async () => { const j = await (await fetch(API + '?action=catalog&t=' + Date.now())).json(); return j.products.find(p => p.id === id).stock; };
(async () => {
  const before = await stockOf();
  console.log(`פריט ${id}: יתרה לפני הבדיקה ${before}. שולחים ${N} הזמנות במקביל...`);
  const t0 = Date.now();
  const results = await Promise.all(Array.from({ length: N }, (_, i) => {
    const body = { action: 'submitOrder', testCode: code, overOk: true,
      customer: { zimmer: `בדיקת מקביליות ${i + 1}`, town: 'בדיקה', contact: 'בדיקה אוטומטית', phone: '05399000' + String(i + 1).padStart(2, '0'),
        email: 'concurrency-test@example.com', address: 'כתובת בדיקה 1', day: '1.12.26' },
      lines: [{ type: 'item', id, qty: 1 }] };
    const s = Date.now();
    return fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) })
      .then(r => r.json()).then(j => ({ i: i + 1, ms: Date.now() - s, ...j })).catch(e => ({ i: i + 1, ok: false, code: 'net', error: e.message }));
  }));
  results.forEach(r => console.log(`  #${String(r.i).padStart(2)} ${String(r.ms).padStart(5)}ms  ${r.ok ? 'נקלטה ' + r.orderId : (r.code + ': ' + r.error)}`));
  const okN = results.filter(r => r.ok).length, short = results.filter(r => r.code === 'shortage').length, other = N - okN - short;
  const after = await stockOf();
  console.log(`\nנקלטו: ${okN} | אזל: ${short} | אחר: ${other} | זמן כולל: ${Date.now() - t0}ms`);
  console.log(`יתרה אחרי: ${after}`);
  const pass = okN === Math.min(before, N) && after === before - okN && after >= 0 && other === 0;
  console.log(pass ? `\nPASS בדיוק ${okN} הזמנות נקלטו (כמו היתרה), שאר ${short} קיבלו "אזל", היתרה ${after} ולא שלילית`
                   : '\nFAIL התוצאה לא תואמת לציפייה');
  console.log('הזמנות לביטול במסך הניהול: ' + results.filter(r => r.ok).map(r => r.orderId).join(', '));
  process.exit(pass ? 0 : 1);
})();
