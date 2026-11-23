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
        if (typeof a === 'string') return { setNumberFormat() { return this; }, setFormula(f) { sh.formula = f; return this; } };
        const r0 = a - 1, c0 = b - 1, nr = c || 1, nc = d || 1;
        return {
          setValues(v) { for (let i = 0; i < nr; i++) { data[r0 + i] = data[r0 + i] || []; for (let j = 0; j < nc; j++) data[r0 + i][c0 + j] = v[i][j]; } return this; },
          setValue(v) { data[r0] = data[r0] || []; data[r0][c0] = v; return this; },
          getValues() { const o = []; for (let i = 0; i < nr; i++) { o.push([]); for (let j = 0; j < nc; j++) o[i].push((data[r0 + i] || [])[c0 + j] ?? ''); } return o; },
          setNumberFormat() { return this; }, setBackground() { return this; }, setFontColor() { return this; }, setFontWeight() { return this; }
        };
      },
      appendRow: r => data.push(r.slice()),
      setRightToLeft() {}, setFrozenRows() {}, clear() { data.length = 0; }
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
  const cache = {};
  const mails = [];
  const files = {};
  const fetches = [];
  const triggers = [];
  let lockBusy = false;
  const mkBlob = (b, t, n) => {
    const o = { bytes: b, type: t, name: n,
      setDataFromString(s) { o.bytes = Buffer.from(s, 'utf8'); return o; },
      setContentType(x) { o.type = x; return o; }, setName(x) { o.name = x; return o; },
      getBytes: () => o.bytes, getName: () => o.name };
    return o;
  };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (lockBusy) return false; lockBusy = true; return true; }, releaseLock: () => { lockBusy = false; } }) },
    Utilities: {
      formatDate: (d, tz, f) => { const p = n => String(n).padStart(2, '0'); return f.replace('yyyyMMdd', `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`).replace('yyMMdd', `${String(d.getFullYear()).slice(2)}${p(d.getMonth() + 1)}${p(d.getDate())}`).replace('dd/MM/yyyy HH:mm', `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`); },
      base64Decode: s => Buffer.from(s, 'base64'), newBlob: (b, t, n) => mkBlob(b, t, n), getUuid: () => require('crypto').randomUUID(), sleep() {},
      // zip: מחזיר "בלוב" עם רשימת הקבצים; test-server.js מרכיב ממנו zip אמיתי ופותח עם ExcelJS
      zip: (blobs, name) => { const z = mkBlob(null, 'application/zip', name); z.entries = blobs; return z; }
    },
    UrlFetchApp: {
      fetchAll: reqs => { fetches.push(...reqs.map(r => r.url)); return reqs.map(r => {
        const f = path.join(__dirname, '..', 'pic', 'thumbs', decodeURIComponent(r.url.split('/').pop()));
        const ok = !opts.noFetch && fs.existsSync(f);
        return { getResponseCode: () => ok ? 200 : 404, getContent: () => ok ? fs.readFileSync(f) : null };
      }); },
      fetch: () => ({ getResponseCode: () => 200 })
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: fn => ({ timeBased: () => ({ everyMinutes: m => ({ create: () => { const t = { fn, m, getHandlerFunction: () => fn }; triggers.push(t); return t; } }) }) })
    },
    MailApp: { getRemainingDailyQuota: () => opts.quota ?? 1500, sendEmail: o => { if (opts.mailThrows) throw new Error(opts.mailThrows); mails.push(o); } },
    DriveApp: {
      createFolder: n => ({ getId: () => 'FOLDER1', createFile: b => { const id = 'FILE_' + 'x'.repeat(30) + Object.keys(files).length; files[id] = b; return { getId: () => id, getUrl: () => 'https://drive.google.com/file/d/' + id + '/view', getBlob: () => b }; } }),
      getFolderById: id => ctx.DriveApp.createFolder(),
      getFileById: id => { if (!files[id]) throw new Error('file not found'); return { getBlob: () => files[id] }; }
    },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'x@contrast-il.com' }) },
    Logger: { log: m => (opts.log ? console.log('   log:', m) : null) },
    ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(SRC_PATH, "utf8"), ctx);
  ctx.setupSheet();
  return { ctx, sheets, mails, files, fetches, triggers, cache, setLockBusy: v => { lockBusy = v; }, props };
}


module.exports = { makeEnv };
