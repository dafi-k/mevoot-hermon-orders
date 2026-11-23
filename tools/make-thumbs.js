// תמונות מוקטנות JPEG (160×120, רקע לבן) לכל תמונה ב-pic/ -> pic/thumbs/<שם>.jpg
// משמשות במייל ההזמנה (webp לא נתמך בכל תוכנות המייל) וב-Excel שהשרת בונה.
// הרצה: node make-thumbs.js  (להריץ שוב אחרי כל ייבוא קטלוג שמשנה תמונות)
const sharp = require('sharp');
const fs = require('fs'), path = require('path');
const PIC = path.join(__dirname, '..', 'pic'), OUT = path.join(PIC, 'thumbs');
fs.mkdirSync(OUT, { recursive: true });
(async () => {
  const files = fs.readdirSync(PIC).filter(f => /\.webp$/i.test(f));
  for (const f of files) {
    await sharp(path.join(PIC, f))
      .resize(160, 120, { fit: 'contain', background: '#ffffff' })
      .flatten({ background: '#ffffff' })
      .jpeg({ quality: 80 })
      .toFile(path.join(OUT, f.replace(/\.webp$/i, '.jpg')));
  }
  console.log(`נוצרו ${files.length} תמונות מוקטנות ב-pic/thumbs`);
})();
