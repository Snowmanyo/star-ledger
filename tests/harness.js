// 模擬 LINE 對話：載入真正的 Code.gs，假的試算表、快取、LINE API 與 Gemini
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const CODE = fs.readFileSync(process.argv[2], 'utf8');
const SEED = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  _get() { const out = []; for (let i = 0; i < this.nr; i++) { const row = this.sh.rows[this.r - 1 + i] || []; out.push(Array.from({ length: this.nc }, (_, j) => row[this.c - 1 + j] ?? '')); } return out; }
  getValues() { return this._get(); } getDisplayValues() { return this._get().map(r => r.map(String)); }
  setValues(v) { v.forEach((row, i) => { const R = this.sh.rows[this.r - 1 + i] = this.sh.rows[this.r - 1 + i] || []; row.forEach((x, j) => { R[this.c - 1 + j] = x; }); }); return this; }
  setFontWeight() { return this; } setNumberFormat() { return this; }
  clearContent() { return this; }
}
class Sheet {
  constructor() { this.rows = []; }
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
  getLastRow() { let n = this.rows.length; while (n && !(this.rows[n - 1] || []).some(x => x !== '' && x != null)) n--; return n; }
  getMaxRows() { return 1000; } appendRow(line) { this.rows[this.getLastRow()] = line.slice(); } deleteRow(i) { this.rows.splice(i - 1, 1); } setFrozenRows() {}
}
const sheets = {};
const ss = { getId: () => 'SHEET1', getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet()) };
const drive = { folders: {}, seq: 0 };
const iter = list => { let i = 0; return { hasNext: () => i < list.length, next: () => list[i++] }; };
const mkFolder = name => (drive.folders[name] = drive.folders[name] || { name, files: [], getFiles() { return iter(this.files.filter(f => !f.trashed)); } });
const props = { GEMINI_API_KEY: 'g', SHARED_KEY: 'pw123', LINE_TOKEN: 'lt' };
const cache = {};
const QUOTA = { used: 10 };
let NOW_ISO = '2026-10-06T09:00:00+08:00';
class FakeDate extends Date { constructor(...a) { a.length ? super(...a) : super(NOW_ISO); } static now() { return new Date(NOW_ISO).getTime(); } }
const setNow = (day, hm) => { NOW_ISO = day + 'T' + hm + ':00+08:00'; };
const NOW = { day: '2026-10-06', hm: '09:00', wd: '2' }; const triggers = []; const pushes = []; const lineCalls = []; const lineMenus = [];
const replies = []; const geminiCalls = []; const loading = [];
const aiByImage = {}; const editQueue = []; const geminiPrompts = [];
const ctx = {
  console, JSON, Date: FakeDate, Math, String, Object, Array, Number, isFinite, encodeURIComponent, decodeURIComponent,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
  CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  DriveApp: {
    getFoldersByName: n => iter(drive.folders[n] ? [drive.folders[n]] : []), createFolder: n => mkFolder(n),
    getFileById: id => ({ makeCopy: (name, folder) => { const t = ++drive.seq; folder.files.push({ name, trashed: false, getDateCreated: () => new Date(2026, 0, 1, 0, 0, t), setTrashed(v) { this.trashed = v; } }); } }),
  },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: { formatDate: (d, tz, f) => { const t = new Date(d.getTime() + (tz === 'UTC' ? 0 : 8 * 3600000)); const p = n => String(n).padStart(2, '0'); return f.replace('yyyy', t.getUTCFullYear()).replace('MM', p(t.getUTCMonth() + 1)).replace('dd', p(t.getUTCDate())).replace('HH', p(t.getUTCHours())).replace('mm', p(t.getUTCMinutes())).replace('ss', p(t.getUTCSeconds())).replace(/'/g, '').replace(' u', ' ' + (t.getUTCDay() || 7)); }, sleep: () => {}, base64Encode: b => Buffer.from(b).toString('base64'), base64Decode: s => s },
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/X/exec' }), getProjectTriggers: () => triggers, deleteTrigger: t => triggers.splice(triggers.indexOf(t), 1), newTrigger: h => ({ timeBased: () => ({ everyMinutes: n => ({ create: () => triggers.push({ h, n, getHandlerFunction: () => h }) }) }) }) },
  ContentService: { MimeType: { JSON: 1, ICAL: 2 }, createTextOutput: t => ({ text: t, setMimeType() { return this; } }) },
  Logger: { log: console.log },
  UrlFetchApp: { fetch(url, o = {}) {
    const R = (code, body) => ({ getResponseCode: () => code, getContentText: () => typeof body === 'string' ? body : JSON.stringify(body) });
    if (url.includes('/message/reply')) { const p = JSON.parse(o.payload); replies.push(p); return R(200, '{}'); }
    if (url.includes('/message/multicast')) { const pp = JSON.parse(o.payload); pp.messages.forEach(mm => { if (typeof mm !== 'object' || !mm.type) throw new Error('推播訊息格式錯誤：' + JSON.stringify(mm)); }); pushes.push(pp); return R(200, '{}'); }
    if (url.includes('/message/quota/consumption')) return R(200, { totalUsage: QUOTA.used });
    if (url.includes('/message/quota')) return R(200, { type: 'limited', value: 200 });
    if (url.includes('richmenu') || url.includes('snowmanyo.github.io')) { lineCalls.push((o.method || 'get') + ' ' + url); if (/bot\/richmenu$/.test(url)) lineMenus.push(o.payload); return { getResponseCode: () => 200, getContentText: () => '{"richMenuId":"rm1"}', getBlob: () => ({ getBytes: () => [1, 2] }) }; }
    if (url.includes('/chat/loading')) { loading.push(JSON.parse(o.payload).chatId); return R(200, '{}'); }
    if (url.includes('api-data.line.me')) { const id = url.match(/message\/([^/]+)\/content/)[1]; return { getResponseCode: () => 200, getBlob: () => ({ getContentType: () => 'image/jpeg', getBytes: () => Buffer.from(id) }) }; }
    if (url.includes('open.er-api.com')) return R(200, { rates: { KRW: 42.1, JPY: 4.6, USD: 0.031 } });
    if (url.includes('generativelanguage')) {
      const body = JSON.parse(o.payload); const parts = body.contents[0].parts;
      const imgs = parts.filter(p => p.inline_data).map(p => Buffer.from(p.inline_data.data, 'base64').toString());
      geminiCalls.push(imgs.length ? imgs : ['TEXT']); geminiPrompts.push(parts[0].text || '');
      let data;
      if (imgs.length) { data = {}; imgs.forEach(id => Object.assign(data, aiByImage[id] || {})); }
      else data = editQueue.shift() || { understood: false };
      return R(200, { candidates: [{ content: { parts: [{ text: JSON.stringify(data) }] } }] });
    }
    throw new Error('unexpected fetch ' + url);
  } },
};
vm.createContext(ctx);
vm.runInContext(CODE + '\n;this.doPost=doPost;', ctx);
// 每次模擬的請求都是新的執行：清掉同一次執行內的快取
vm.runInContext("['doPost','doGet','reminderTick','setupBot'].forEach(function (n) { var f = this[n]; this[n] = function () { Object.keys(TABLE_MEMO_).forEach(function (k) { delete TABLE_MEMO_[k]; }); return f.apply(this, arguments); }; }, this);", ctx);

// 試算表：放入正式 seed（新欄位）
vm.runInContext('setup_()', ctx);
const T = vm.runInContext('TABLES', ctx);
for (const t of ['events', 'ledger']) (SEED[t] || []).forEach(r => sheets[t].rows.push(T[t].map(c => { const v = r[c]; return v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v); })));
const table = name => { const sh = sheets[name]; const H = sh.rows[0]; return sh.rows.slice(1).filter(r => r[0]).map(r => Object.fromEntries(H.map((h, i) => [h, r[i]]))); };

let tok = 0;
const hook = (events, key = 'pw123') => { replies.length = 0; ctx.doPost({ parameter: { line: key }, postData: { contents: JSON.stringify({ events }) } }); return replies.map(r => r.messages); };
const src = uid => ({ type: 'user', userId: uid });
const text = (uid, t) => hook([{ type: 'message', replyToken: 'r' + (++tok), source: src(uid), message: { type: 'text', id: 't' + tok, text: t } }]);
const img = (uid, id, set) => hook([{ type: 'message', replyToken: 'r' + (++tok), source: src(uid), message: Object.assign({ type: 'image', id }, set ? { imageSet: set } : {}) }]);
const pb = (uid, data) => hook([{ type: 'postback', replyToken: 'r' + (++tok), source: src(uid), postback: { data } }]);
const flexTexts = o => { const out = []; const walk = x => { if (x === null) throw new Error('Flex 裡有空值'); if (typeof x !== 'object') return; if (Array.isArray(x) && x.some(y => y === null || y === undefined)) throw new Error('Flex 陣列裡有空值'); if (x.type === 'text') { if (!String(x.text).length) throw new Error('Flex 文字不能是空的'); out.push(x.text); } Object.values(x).forEach(walk); }; walk(o); return out; };
const flat = msgs => msgs.flat().map(m => (m.type === 'flex' ? '[卡片] ' + m.altText + '\n' + flexTexts(m.contents).join('\n') : m.text) + (m.quickReply ? ' {' + m.quickReply.items.map(i => i.action.label).join('|') + '}' : '')).join(' / ');
const qrData = (msgs, label) => msgs.flat().find(m => m.quickReply).quickReply.items.find(i => i.action.label.startsWith(label)).action.data;
const log = (name, msgs) => console.log('• ' + name.padEnd(14) + flat(msgs));

