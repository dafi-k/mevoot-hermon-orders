// ייבוא חד-פעמי: קורא את קובץ האקסל של המוצרים, ממיר תמונות ל-webp לתוך pic/
// ומייצר data/products.csv (לגיליון) + data/products.json + data/import-report.txt
// הרצה: cd tools && node import-catalog.js
// את התיקיות המקוריות לא משנים - רק קוראים מהן.

const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const sharp = require('sharp');

const SRC_DIR = path.resolve(__dirname, '..', '..');               // תיקיית "קטלוג מבואות חרמון"
const XLSX = path.join(SRC_DIR, 'אתר הזמנות מבואות החרמון - מעודכן תמונות.xlsx');
const REPO = path.resolve(__dirname, '..');
const PIC = path.join(REPO, 'pic');
const DATA = path.join(REPO, 'data');
const MAX_BYTES = 500 * 1024;

// לשונית באקסל -> קטגוריה באתר. prefix משמש לבניית מזהה ייחודי.
const SHEETS = [
  { sheet: 'כיסאות פלסטיק',      cat: 'plastic',  label: 'כיסאות פלסטיק',     prefix: 'CH'  },
  { sheet: 'כורסאות',            cat: 'armchair', label: 'כורסאות',           prefix: 'AR'  },
  { sheet: 'כיסאות פינת אוכל',   cat: 'dining',   label: 'כיסאות פינת אוכל',  prefix: 'DC'  },
  { sheet: 'כיסאות בר',          cat: 'bar',      label: 'כיסאות בר',         prefix: 'BC'  },
  { sheet: 'שולחנות קפה',        cat: 'coffee',   label: 'שולחנות קפה',       prefix: 'CT'  },
  { sheet: 'מראות',              cat: 'mirror',   label: 'מראות',             prefix: 'MR'  },
  { sheet: 'רגלי ברזל+פלטות',    cat: 'tables',   label: 'בנה שולחן',         prefix: 'TB'  },
  { sheet: 'שולחנות חוץ_פנים',   cat: 'otable',   label: 'שולחנות חוץ/פנים',  prefix: 'OT'  },
  { sheet: 'אחר',                cat: 'misc',     label: 'שונות',             prefix: 'MS'  },
  { sheet: 'ריהוט חוץ',          cat: 'outdoor',  label: 'ריהוט חוץ',         prefix: 'OD'  },
];

// תיקונים שאושרו אחרי בדיקה, בלי לשנות את קובץ האקסל המקורי (מזהה -> שדות לדריסה)
const CORRECTIONS = {
  'CH-P3': { color: 'ירוק', note: '' }, // אושר 22.11.26: הכיסא ירוק כמו בתמונה
};

const TOP_TONES ={ 'שחור': 'black', 'אגוז': 'walnut', 'טבעי': 'natural' };

// רק מקף רגיל, בלי רווחים כפולים
function clean(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.richText) v = v.richText.map(t => t.text).join('');
  if (typeof v === 'object' && v.result !== undefined) v = v.result;
  return String(v).replace(/[–—־]/g, '-').replace(/\s+/g, ' ').trim();
}
function num(v) { const s = clean(v); if (s === '') return null; const n = Number(s.replace(/[^\d.]/g, '')); return isNaN(n) ? null : n; }

function makeId(prefix, code) {
  return code.toUpperCase().startsWith(prefix) ? code.toUpperCase() : `${prefix}-${code.toUpperCase()}`;
}

async function toWebp(srcRel, outName, report) {
  const src = path.join(SRC_DIR, srcRel.replace(/\\/g, path.sep));
  if (!fs.existsSync(src)) { report.push(`!! תמונה לא נמצאה: ${srcRel}`); return ''; }
  const out = path.join(PIC, outName);
  let size = 1200, q = 82, buf;
  for (let i = 0; i < 8; i++) {
    buf = await sharp(src).rotate().resize({ width: size, height: size, fit: 'inside', withoutEnlargement: true })
      .flatten({ background: '#ffffff' }).webp({ quality: q }).toBuffer();
    if (buf.length <= MAX_BYTES) break;
    q = Math.max(50, q - 8); size = Math.round(size * 0.85);
  }
  fs.writeFileSync(out, buf);
  return outName;
}

(async () => {
  fs.mkdirSync(PIC, { recursive: true });
  fs.mkdirSync(DATA, { recursive: true });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX);
  const report = [];
  const products = [];
  const imgDone = {}; // מקור -> שם קובץ (לתמונות משותפות)
  let order = 0;

  for (const def of SHEETS) {
    const ws = wb.getWorksheet(def.sheet);
    if (!ws) { report.push(`!! לשונית חסרה: ${def.sheet}`); continue; }
    let rows = 0;
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const sku = clean(row.getCell('A').value);
      const code = clean(row.getCell('B').value);
      if (!sku && !code) continue;
      rows++;
      const descC = clean(row.getCell('C').value);
      let color = clean(row.getCell('D').value);
      if (color === 'לא צוין') color = '';
      const qty = num(row.getCell('E').value);
      const price = num(row.getCell('F').value);
      const imgRel = clean(row.getCell('I').value);
      const note = clean(row.getCell('J').value);
      const descK = clean(row.getCell('K').value);

      const p = {
        id: '', cat: def.cat, catLabel: def.label, code, sku,
        name: descC, desc: '', color, price: price || 0,
        stock: qty === null ? '' : qty, unlimited: qty === null ? 'כן' : '',
        img: '', kind: 'item', tone: '', diameter: '', gift: '', active: 'כן', note, order: ++order,
      };

      if (def.cat === 'tables') {
        if (/^LP/i.test(code)) { p.kind = 'leg'; p.id = makeId('LG', code); }
        else {
          p.kind = 'top'; p.id = makeId('TOP', code);
          const tone = Object.keys(TOP_TONES).find(t => descC.includes(t)) || '';
          p.tone = tone; p.diameter = num(descC.match(/(\d{2,3})\s*$/)?.[1]) || '';
          p.color = tone;
        }
      } else if (def.cat === 'outdoor') {
        p.id = code.toUpperCase();
        p.desc = descK;
        // מתנה: "+ X במתנה" או "+ מתנה X"
        const g = descK.match(/\+\s*([^+]*?)\s*במתנה/) || descK.match(/\+\s*מתנה\s+(.+)$/);
        if (g) p.gift = g[1].trim();
        // שם קצר: עד " - ", אחרת עד הפסיק הראשון, בלי חלק המתנה
        let nm = descK;
        const dash = nm.indexOf(' - ');
        if (dash > 0) nm = nm.slice(0, dash);
        else if (nm.includes(', ')) nm = nm.split(', ')[0];
        if (g && dash < 0) nm = nm.replace(g[0], '').trim();
        p.name = nm;
        if (!descK) report.push(`!! אין תיאור K: ${code}`);
      } else {
        p.id = makeId(def.prefix, code);
      }

      if (CORRECTIONS[p.id]) Object.assign(p, CORRECTIONS[p.id]);

      // תמונה
      if (imgRel && !/חסרה/.test(imgRel)) {
        let outName;
        if (p.kind === 'top') outName = `TOP-${TOP_TONES[p.tone] || 'x'}.webp`;
        else outName = `${p.id}.webp`;
        if (!imgDone[outName]) imgDone[outName] = await toWebp(imgRel, outName, report);
        p.img = imgDone[outName];
      } else report.push(`!! אין תמונה: ${p.id}`);

      products.push(p);
    }
    report.push(`${def.sheet}: ${rows} שורות`);
  }

  // בדיקות
  const ids = {}, skus = {};
  products.forEach(p => { (ids[p.id] = ids[p.id] || []).push(p.code); (skus[p.sku] = skus[p.sku] || []).push(p.id); });
  Object.entries(ids).filter(([, v]) => v.length > 1).forEach(([k]) => report.push(`!! מזהה כפול (חובה לתקן): ${k}`));
  Object.entries(skus).filter(([, v]) => v.length > 1).forEach(([k, v]) => report.push(`אזהרה - מק"ט כפול: ${k} -> ${v.join(', ')}`));
  report.push(`סה"כ מוצרים: ${products.length}`);
  report.push(`ללא הגבלת מלאי: ${products.filter(p => p.unlimited).map(p => p.id).join(', ')}`);

  // לוגו
  const logoSrc = path.resolve(SRC_DIR, '..', 'GMA', 'gma-site', 'client-logos', 'Logo contrast.png');
  if (fs.existsSync(logoSrc)) {
    await sharp(logoSrc).resize({ width: 600, withoutEnlargement: true }).png().toFile(path.join(PIC, 'logo-contrast.png'));
  } else report.push('!! לוגו CONTRAST לא נמצא');

  // פלט
  const COLS = [
    ['id', 'מזהה'], ['cat', 'קטגוריה'], ['catLabel', 'שם קטגוריה'], ['code', 'קוד'], ['sku', 'מק"ט'],
    ['name', 'שם'], ['desc', 'תיאור מורחב'], ['color', 'צבע'], ['price', 'מחיר'],
    ['stock', 'מלאי התחלתי'], ['stock', 'יתרה'], ['unlimited', 'ללא הגבלה'], ['img', 'תמונה'], ['kind', 'סוג'],
    ['tone', 'גוון'], ['diameter', 'קוטר'], ['gift', 'מתנה'], ['active', 'פעיל'], ['note', 'הערה פנימית'], ['order', 'סדר'],
  ];
  const csvEsc = v => { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const csv = [COLS.map(c => csvEsc(c[1])).join(',')]
    .concat(products.map(p => COLS.map(c => csvEsc(p[c[0]])).join(','))).join('\r\n');
  fs.writeFileSync(path.join(DATA, 'products.csv'), '﻿' + csv, 'utf8');
  fs.writeFileSync(path.join(DATA, 'products.json'), JSON.stringify(products, null, 1), 'utf8');

  // עותק שמור לאתר: רק שדות ציבוריים (בלי הערות פנימיות), באותו מבנה שהשרת מחזיר
  const snapshot = products.map(p => ({
    id: p.id, cat: p.cat, catLabel: p.catLabel, code: p.code, sku: p.sku, name: p.name, desc: p.desc,
    color: p.color, price: p.price, stock: p.unlimited ? null : Number(p.stock) || 0, img: p.img,
    kind: p.kind, tone: p.tone, diameter: p.diameter === '' ? '' : Number(p.diameter), gift: p.gift, order: p.order,
  }));
  fs.writeFileSync(path.join(DATA, 'catalog-snapshot.json'),
    JSON.stringify({ ok: true, budget: 9500, products: snapshot }), 'utf8');

  const sizes = fs.readdirSync(PIC).map(f => fs.statSync(path.join(PIC, f)).size);
  report.push(`תמונות ב-pic: ${sizes.length}, הגדולה ביותר: ${Math.round(Math.max(...sizes) / 1024)}KB, סה"כ ${Math.round(sizes.reduce((a, b) => a + b, 0) / 1024)}KB`);
  fs.writeFileSync(path.join(DATA, 'import-report.txt'), report.join('\n'), 'utf8');
  console.log(report.join('\n'));
})().catch(e => { console.error(e); process.exit(1); });
