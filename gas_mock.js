// Apps Script चे बनावट (mock) वातावरण — Node `vm` मध्ये Code.gs चालवण्यासाठी. खऱ्या Sheet/Drive/Apps Script वर चाचणी नाही.
const vm = require('vm'), fs = require('fs'), crypto = require('crypto');

class Range {
  constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getRow() { return this.r; }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.rows[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined || v === null ? '' : v); } out.push(row); } return out; }
  getValue() { return this.getValues()[0][0]; }
  setValues(vals) {
    if (vals.length !== this.nr || (vals[0] && vals[0].length !== this.nc)) throw new Error('setValues: आकार जुळत नाही');
    if (this.r - 1 + this.nr > this.sh.maxRows) throw new Error('Range बाहेर: ओळी Sheet च्या maxRows पेक्षा जास्त (ensureRows_ गहाळ?)');
    if (this.c - 1 + this.nc > this.sh.maxCols) throw new Error('Range बाहेर: स्तंभ maxCols पेक्षा जास्त');
    vals.forEach((row, i) => { const ri = this.r - 1 + i; while (this.sh.rows.length <= ri) this.sh.rows.push([]); row.forEach((v, j) => { this.sh.rows[ri][this.c - 1 + j] = v; }); });
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) { const row = this.sh.rows[this.r - 1 + i]; if (row) row[this.c - 1 + j] = ''; } return this; }
  createTextFinder(key) {
    const self = this; let whole = false;
    return { matchEntireCell(b) { whole = b; return this; }, matchCase() { return this; },
      findAll() { const out = [], k = String(key).toLowerCase(); self.getValues().forEach((row, i) => row.forEach(v => { const s = String(v).toLowerCase(); if (whole ? s === k : s.includes(k)) out.push({ getRow: () => self.r + i }); })); return out; } };
  }
}
class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.maxRows = 1000; this.maxCols = 26; this.hidden = false; }
  getName() { return this.name; }
  setName(n) { this.name = n; }
  getLastRow() { let n = this.rows.length; while (n > 0 && this.rows[n - 1].every(v => v === '' || v === null || v === undefined)) n--; return n; }
  getLastColumn() { let m = 0; this.rows.forEach(r => r.forEach((v, j) => { if (v !== '' && v !== undefined && v !== null) m = Math.max(m, j + 1); })); return m; }
  getMaxRows() { return this.maxRows; } getMaxColumns() { return this.maxCols; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc); }
  appendRow(arr) { const lr = this.getLastRow(); if (lr + 1 > this.maxRows) this.maxRows = lr + 1; while (this.rows.length < lr) this.rows.push([]); this.rows[lr] = arr.slice(); this.rows.length = lr + 1; }
  insertRowsAfter(n, k) { this.maxRows += k; } insertColumnsAfter(n, k) { this.maxCols += k; }
  deleteRows(start, k) { this.rows.splice(start - 1, k); this.maxRows -= k; } deleteRow(r) { this.deleteRows(r, 1); }
  deleteColumns(s, k) { this.maxCols -= k; }
  setFrozenRows() {} hideSheet() { this.hidden = true; } showSheet() { this.hidden = false; }
  copyTo() { return new Sheet('copy'); }
}
class Spreadsheet {
  constructor() { this.sheets = []; }
  getName() { return 'MockSS'; } getId() { return 'mock-ss-id'; }
  getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(n) { if (this.getSheetByName(n)) throw new Error('Sheet आधीच आहे: ' + n); const s = new Sheet(n); this.sheets.push(s); return s; }
  deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
}
function pad(n, w = 2) { return String(n).padStart(w, '0'); }
function fmtDate(d, tz, pat) {
  const parts = new Intl.DateTimeFormat('en-GB', { timeZone: tz || 'Asia/Kolkata', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(d).reduce((a, p) => (a[p.type] = p.value, a), {});
  if (parts.hour === '24') parts.hour = '00';
  return pat.replace(/yyyy/g, parts.year).replace(/MM/g, parts.month).replace(/dd/g, parts.day).replace(/HH/g, parts.hour).replace(/mm/g, parts.minute).replace(/ss/g, parts.second);
}
function makeEnv(opts = {}) {
  const ss = new Spreadsheet(); const cache = {}, props = {}; const env = { ss, cache, props, locked: false, lockFail: false, logs: [] };
  const b64w = b => Buffer.from(b).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  const ctx = {
    console, JSON, Math, Date, parseInt, parseFloat, isNaN, String, Number, Object, Array, RegExp, Error, encodeURIComponent, decodeURIComponent, escape, unescape,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {}, openById() { throw new Error('openById: mock मध्ये नाही'); } },
    CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = String(v); }, remove: k => { delete cache[k]; }, getAll: ks => { const o = {}; ks.forEach(k => { if (k in cache) o[k] = cache[k]; }); return o; }, putAll: o => Object.assign(cache, o), removeAll: ks => ks.forEach(k => delete cache[k]) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; }, getProperties: () => Object.assign({}, props), setProperties: o => Object.assign(props, o) }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (env.lockFail) return false; env.locked = true; return true; }, waitLock: () => { if (env.lockFail) throw new Error('Lock timeout'); env.locked = true; }, releaseLock: () => { env.locked = false; } }) },
    ContentService: { MimeType: { JSON: 'json', JAVASCRIPT: 'js' }, createTextOutput: t => ({ content: t, setMimeType() { return this; }, getContent() { return t; } }) },
    Logger: { log: m => env.logs.push(m) },
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ onWeekDay: () => ({ atHour: () => ({ create() {} }) }) }) }), WeekDay: { SUNDAY: 'SUNDAY' }, deleteTrigger() {} },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      formatDate: (d, tz, pat) => fmtDate(d, tz, pat),
      getUuid: () => crypto.randomUUID(),
      base64EncodeWebSafe: (x) => b64w(typeof x === 'string' ? Buffer.from(x, 'utf8') : Buffer.from(x)),
      base64DecodeWebSafe: s => Array.from(Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64')),
      computeHmacSha256Signature: (msg, key) => Array.from(crypto.createHmac('sha256', key).update(msg).digest()),
      newBlob: bytes => ({ getDataAsString: () => Buffer.from(bytes).toString('utf8') }),
      sleep() {},
    },
  };
  vm.createContext(ctx);
  env.ctx = ctx;
  env.load = (file) => { vm.runInContext(fs.readFileSync(file, 'utf8'), ctx, { filename: file }); };
  env.run = (code) => vm.runInContext(code, ctx);
  env.sheet = (n) => ss.getSheetByName(n);
  env.seed = (name, headers, rows) => { const s = ss.insertSheet(name); s.maxCols = headers.length + 3; s.appendRow(headers); /* खऱ्या ॲपमध्ये trimNewSheet_ हेच करतो */ (rows || []).forEach(r => s.appendRow(r)); return s; };
  env.token = (u, r, extra) => ctx.signToken_(Object.assign({ u, r, exp: Date.now() + 3600e3 }, extra || {}));
  env.get = (params) => { const out = ctx.doGet({ parameter: Object.assign({ callback: 'cb' }, params) }); const t = out.getContent(); const m = t.match(/^cb\(([\s\S]*)\);$/); return JSON.parse(m ? m[1] : t); };
  env.post = (params) => { const out = ctx.doPost({ parameter: params }); return JSON.parse(out.getContent()); };
  env.allText = (sheetName) => { const s = ss.getSheetByName(sheetName); return s ? JSON.stringify(s.rows) : ''; };
  env.snapshot = (exclude) => { const o = {}; ss.sheets.forEach(s => { if (!(exclude || []).includes(s.name)) o[s.name] = JSON.stringify(s.rows); }); return o; };
  return env;
}
module.exports = { makeEnv, Sheet };
