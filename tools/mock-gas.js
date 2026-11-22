// הדמיה של שירותי Google Apps Script (גיליון, נעילה, מייל, Drive) להרצת apps-script.js מקומית
const fs = require('fs'), path = require('path'), vm = require('vm');
const SRC_PATH = path.join(__dirname, '..', 'apps-script.js');
const snapshot = require('../data/catalog-snapshot.json').products;

function makeEnv(opts = {}) {
  // ---------- גיליון מדומה ----------
  const sheets = {};
  const mkSheet = name => {
    const data = [];
    const sh = {
      name, data,
      getLastRow: () => data.length,
      getLastColumn: () => Math.max(0, ...data.map(r => r.length)),
      getMaxRows: () => 1000,
      getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
      getRange: (a, b, c, d) => {
        if (typeof a === 'string') return { setNumberFormat() {} };
        const r0 = a - 1, c0 = b - 1, nr = c || 1, nc = d || 1;
        return {
          setValues(v) { for (let i = 0; i < nr; i++) { data[r0 + i] = data[r0 + i] || []; for (let j = 0; j < nc; j++) data[r0 + i][c0 + j] = v[i][j]; } return this; },
          setValue(v) { data[r0] = data[r0] || []; data[r0][c0] = v; return this; },
          getValues() { const o = []; for (let i = 0; i < nr; i++) { o.push([]); for (let j = 0; j < nc; j++) o[i].push((data[r0 + i] || [])[c0 + j] ?? ''); } return o; },
          setNumberFormat() { return this; }, setBackground() { return this; }, setFontColor() { return this; }, setFontWeight() { return this; }
        };
      },
      appendRow: r => data.push(r.slice()),
      setRightToLeft() {}, setFrozenRows() {}
    };
    // מרפדים שורות קצרות בערך ריק, כמו בגיליון אמיתי
    const orig = sh.getDataRange;
    sh.getDataRange = () => { const w = sh.getLastColumn(); return { getValues: () => data.map(r => { const x = r.slice(); while (x.length < w) x.push(''); return x.map(v => v === undefined ? '' : v); }) }; };
    sheets[name] = sh;
    return sh;
  };
  const ss = {
    getId: () => 'SHEET1',
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => mkSheet(n),
    getSheets: () => Object.values(sheets),
    deleteSheet: s => delete sheets[s.name],
    setSpreadsheetLocale() {}
  };
  // לשונית מוצרים מהעותק השמור
  const P = mkSheet('מוצרים');
  const H = ['מזהה', 'קטגוריה', 'שם קטגוריה', 'קוד', 'מק"ט', 'שם', 'תיאור מורחב', 'צבע', 'מחיר', 'מלאי התחלתי', 'יתרה', 'ללא הגבלה', 'תמונה', 'סוג', 'גוון', 'קוטר', 'מתנה', 'פעיל', 'הערה פנימית', 'סדר'];
  P.data.push(H);
  snapshot.forEach(p => P.data.push([p.id, p.cat, p.catLabel, p.code, p.sku, p.name, p.desc, p.color, p.price,
    p.stock === null ? '' : p.stock, '', p.stock === null ? 'כן' : '', p.img, p.kind, p.tone, p.diameter, p.gift, 'כן', '', p.order]));

  const props = {};
  const mails = [];
  const files = {};
  let lockBusy = false;
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (lockBusy) return false; lockBusy = true; return true; }, releaseLock: () => { lockBusy = false; } }) },
    Utilities: {
      formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyyMMdd', `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`).replace('yyMMdd', `${String(d.getFullYear()).slice(2)}${p(d.getMonth() + 1)}${p(d.getDate())}`).replace('dd/MM/yyyy HH:mm', `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`); },
      base64Decode: s => Buffer.from(s, 'base64'), newBlob: (b, t, n) => ({ bytes: b, type: t, name: n })
    },
    MailApp: { getRemainingDailyQuota: () => opts.quota ?? 1500, sendEmail: o => { if (opts.mailThrows) throw new Error(opts.mailThrows); mails.push(o); } },
    DriveApp: {
      createFolder: n => ({ getId: () => 'FOLDER1', createFile: b => { const id = 'FILE_' + 'x'.repeat(30) + Object.keys(files).length; files[id] = b; return { getId: () => id, getUrl: () => 'https://drive.google.com/file/d/' + id + '/view' }; } }),
      getFolderById: id => ctx.DriveApp.createFolder(), getFileById: id => ({ getBlob: () => files[id] })
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'x@contrast-il.com' }) },
    Logger: { log: m => (opts.log ? console.log('   log:', m) : null) },
    ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(SRC_PATH, "utf8"), ctx);
  ctx.setupSheet();
  return { ctx, sheets, mails, files, setLockBusy: v => { lockBusy = v; }, props };
}


module.exports = { makeEnv };
