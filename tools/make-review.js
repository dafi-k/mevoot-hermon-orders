// מייצר data/review.html - דף בדיקה של כל המוצרים עם התמונות (לפתוח בדפדפן בלחיצה כפולה)
const fs = require('fs');
const path = require('path');
const DATA = path.resolve(__dirname, '..', 'data');
const products = JSON.parse(fs.readFileSync(path.join(DATA, 'products.json'), 'utf8'));
const report = fs.readFileSync(path.join(DATA, 'import-report.txt'), 'utf8');
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

let html = `<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>בדיקת ייבוא מוצרים</title>
<style>
body{font-family:Assistant,Arial,sans-serif;background:#F7F3EA;color:#2F3A26;margin:0;padding:16px}
h1{color:#5C6B3A}h2{background:#5C6B3A;color:#fff;padding:6px 12px;border-radius:8px;margin-top:28px}
pre{background:#fff;padding:12px;border-radius:8px;white-space:pre-wrap;font-size:13px}
table{border-collapse:collapse;width:100%;background:#fff;font-size:14px}
td,th{border:1px solid #E2D9C3;padding:6px;vertical-align:middle;text-align:right}
th{background:#EFE6D0}img{width:90px;height:90px;object-fit:contain;background:#fff}
.warn{background:#FCE4D6}.inf{color:#B5651D;font-weight:bold}
</style></head><body><h1>בדיקת ייבוא - ${products.length} מוצרים</h1>
<p>שורות בכתום = יש הערה פנימית מהאקסל. "∞" = ללא הגבלת מלאי.</p>
<pre>${esc(report)}</pre>`;
let cat = '';
for (const p of products) {
  if (p.cat !== cat) {
    if (cat) html += '</table>';
    cat = p.cat;
    html += `<h2>${esc(p.catLabel)}</h2><table><tr><th>תמונה</th><th>מזהה</th><th>קוד</th><th>מק"ט</th><th>שם</th><th>צבע/גוון</th><th>קוטר</th><th>מחיר</th><th>מלאי</th><th>מתנה</th><th>תיאור מורחב / הערה</th></tr>`;
  }
  html += `<tr class="${p.note ? 'warn' : ''}"><td>${p.img ? `<img src="../pic/${esc(p.img)}" loading="lazy">` : '<b>אין</b>'}</td>
<td>${esc(p.id)}</td><td>${esc(p.code)}</td><td>${esc(p.sku)}</td><td>${esc(p.name)}</td><td>${esc(p.color)}</td>
<td>${esc(p.diameter)}</td><td>₪${esc(p.price)}</td><td>${p.unlimited ? '<span class="inf">∞</span>' : esc(p.stock)}</td>
<td>${esc(p.gift)}</td><td>${esc(p.desc)}${p.note ? `<br><i>הערה: ${esc(p.note)}</i>` : ''}</td></tr>`;
}
html += '</table></body></html>';
fs.writeFileSync(path.join(DATA, 'review.html'), html, 'utf8');
console.log('נוצר: data/review.html');
