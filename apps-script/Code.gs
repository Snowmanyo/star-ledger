// 追星總帳 Google Sheets API
// 部署方式：擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業 → 網頁應用程式
// 「執行身分：我」、「誰可以存取：知道連結的任何人」→ 複製網頁應用程式網址貼到 App 設定頁。

// 共用密碼與 Gemini 金鑰放在「專案設定 → 指令碼屬性」，不要寫在程式裡（這份程式碼是公開的）：
//   SHARED_KEY      共用密碼，App 設定頁需填相同的值；不設表示不驗證
//   GEMINI_API_KEY  Google AI Studio 申請的金鑰，掃票讀圖用
const PROPS = PropertiesService.getScriptProperties();
const SHARED_KEY = PROPS.getProperty('SHARED_KEY') || '';
const GEMINI_MODELS = ['gemini-3.8-flash', 'gemini-3.5-flash-lite']; // 第一個額度用完或忙線時改用下一個
const SHEET_ID = ''; // 從試算表「擴充功能→Apps Script」開的專案留空；獨立專案填試算表網址 /d/ 後面那串 ID

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

const TABLES = {
  orders: ['id', 'orderNumber', 'channel', 'orderDate', 'estimatedShipDate', 'actualShipDate', 'currency', 'domesticShipping', 'internationalShippingTwd', 'internationalShippingRateTwdPerKg', 'discountAmount', 'weightGrams', 'exchangeRate', 'chargedTwd', 'payer', 'paymentMethod', 'paymentDetail', 'settled', 'notes'],
  items: ['id', 'orderId', 'name', 'variant', 'unitPrice', 'quantity', 'ownership', 'proxyFor', 'arrived', 'sorted', 'proxyPaid', 'salePriceTwd', 'soldQuantity'],
  sales: ['id', 'sourceOrderId', 'sourceItemId', 'sourceOrderNumber', 'sourceChannel', 'name', 'variant', 'sourceCurrency', 'unitOriginalPrice', 'unitCostTwd', 'quantity', 'salePriceTwd', 'soldQuantity', 'managedByOwnership', 'createdAt'],
  events: ['id', 'name', 'artist', 'city', 'venue', 'startDate', 'endDate', 'eventNumber', 'originalDate', 'eventType', 'liveTour', 'seriesEvent', 'seat', 'ticketPriceTwd', 'guest', 'payer', 'settled', 'notes', 'createdAt', 'coverUrl', 'startTime'],
  ledger: ['id', 'type', 'category', 'date', 'title', 'eventId', 'amountTwd', 'currency', 'originalAmount', 'exchangeRate', 'payer', 'paymentMethod', 'paymentDetail', 'counterparty', 'expectedReceivableTwd', 'receivedTwd', 'notes', 'ticketType', 'ticketArea', 'ticketRow', 'ticketSeat', 'attendee', 'ticketStatus', 'createdAt', 'settled', 'ticketFaceTwd', 'ticketBenefitTwd', 'ticketFeeTwd', 'ticketPlatform', 'ticketAccount', 'ticketCount', 'splits', 'ticketPickupDate', 'ticketPickedUp', 'ticketOrderNumber', 'ticketPickupMethod'],
  transfers: ['id', 'date', 'eventId', 'kind', 'person', 'ticketCount', 'ticketArea', 'ticketRow', 'ticketSeat', 'costTwd', 'amountTwd', 'settled', 'notes', 'createdAt', 'title', 'feeTwd'],
};

function setup_() {
  const ss = ss_();
  Object.keys(TABLES).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const cols = TABLES[name];
    const first = sheet.getRange(1, 1, 1, cols.length).getValues()[0];
    if (cols.some(function (c, i) { return String(first[i]) !== c; })) {
      sheet.getRange(1, 1, sheet.getMaxRows(), cols.length).setNumberFormat('@'); // 全文字格式，避免長訂單編號被轉成數字失去精度
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });
}

function readTable_(name) {
  const sheet = ss_().getSheetByName(name);
  const cols = TABLES[name];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, cols.length).getDisplayValues();
  const rows = [];
  values.forEach(function (v) {
    if (!v[0]) return;
    const row = {};
    cols.forEach(function (c, i) { row[c] = v[i]; });
    rows.push(row);
  });
  return rows;
}

function writeRows_(name, rows) {
  const sheet = ss_().getSheetByName(name);
  const cols = TABLES[name];
  const last = sheet.getLastRow();
  const ids = {};
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (v, i) { ids[v[0]] = i + 2; });
  }
  rows.forEach(function (row) {
    const line = cols.map(function (c) {
      const v = row[c];
      return v === null || v === undefined ? '' : String(v);
    });
    const at = ids[String(row.id)];
    if (at) sheet.getRange(at, 1, 1, cols.length).setValues([line]);
    else sheet.appendRow(line);
  });
}

function deleteRows_(name, idList) {
  const sheet = ss_().getSheetByName(name);
  const last = sheet.getLastRow();
  if (last < 2) return;
  const values = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  const wanted = {};
  idList.forEach(function (id) { wanted[String(id)] = true; });
  for (let i = values.length - 1; i >= 0; i--) {
    if (wanted[values[i][0]]) sheet.deleteRow(i + 2);
  }
}

function replaceAll_(data) {
  Object.keys(TABLES).forEach(function (name) {
    if (!data[name]) return;
    const sheet = ss_().getSheetByName(name);
    const last = sheet.getLastRow();
    if (last >= 2) sheet.getRange(2, 1, last - 1, TABLES[name].length).clearContent();
    writeRows_(name, data[name]);
  });
}

function uploadImage_(req) {
  const it = DriveApp.getFoldersByName('追星總帳封面');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('追星總帳封面');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(req.dataBase64),
    req.mimeType || 'image/jpeg',
    req.filename || ('cover-' + Date.now() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1600' };
}

const SCAN_PROMPT = [
  '你是購票訂單截圖的資料擷取助手。圖片是同一筆演唱會／活動購票訂單的一張或多張截圖，請合併讀取，只輸出 JSON。',
  '規則：',
  '- 看不到或無法判斷的欄位填空字串或 0，不要猜。',
  '- date 用 YYYY-MM-DD；截圖沒寫年份時，用今天（{TODAY}）之後最近的那一天推算。time 用 24 小時制 HH:MM。',
  '- eventName 用截圖上的活動名稱原文；artist 是表演者。city 是城市（例：台北、首爾），venue 是場館。',
  '- currency 用 TWD、KRW、JPY、USD 其中之一。unitFace 是單張票面價，unitFee 是單張手續費（只有總手續費就除以張數），totalPaid 是這筆訂單實際付款總額。數字不要千分位逗號。',
  '- seats 每張票一筆：area 區域、row 排、seat 座號（只填數字或代號，不要加「排」「號」）。ticketCount 是張數。',
  '- platform 是售票平台，符合以下其一就用這個寫法：拓元、KKTIX、ibon、年代、寬宏、遠大、NOL、Melon、YES24；Interpark 也寫 NOL；都不是就寫看到的名稱。',
  '- account 是購買人的帳號、姓名、信箱或電話（截圖上看得到才填）。orderNumber 是訂單編號。',
  '- unitBenefit 是單張福利或加購費用（例如福利包、應援福利），沒有就填 0。',
  '- eventType：只有一組表演者的個人演唱會填「專場」；多組表演者同台的音樂節、頒獎典禮、聯合演出填「拼盤」；看不出來填空字串。',
  '- pickupMethod 是取票方式，用「電子票」「超商取票」「現場取票」「宅配」其中之一，看不出來填空字串。',
  '- pickupDate 是可取票日期（YYYY-MM-DD）；如果只寫「演出前 N 天可取票」，pickupDate 填空字串、pickupDaysBefore 填 N。',
  '- payMethod 是付款方式，用 credit_card、bank_transfer、mobile_payment、cash 其中之一；payDetail 是卡別或支付平台名稱（例如 永豐、LINE Pay）。看不出來填空字串。',
  '- uncertain 列出你沒把握的欄位名稱。',
].join('\n');

const SCAN_SCHEMA = {
  type: 'OBJECT',
  properties: {
    eventName: { type: 'STRING' }, artist: { type: 'STRING' },
    date: { type: 'STRING' }, time: { type: 'STRING' },
    city: { type: 'STRING' }, venue: { type: 'STRING' },
    ticketType: { type: 'STRING' }, ticketCount: { type: 'NUMBER' },
    seats: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
      area: { type: 'STRING' }, row: { type: 'STRING' }, seat: { type: 'STRING' } } } },
    currency: { type: 'STRING' }, unitFace: { type: 'NUMBER' }, unitBenefit: { type: 'NUMBER' }, unitFee: { type: 'NUMBER' }, totalPaid: { type: 'NUMBER' },
    platform: { type: 'STRING' }, account: { type: 'STRING' }, eventType: { type: 'STRING' },
    orderNumber: { type: 'STRING' }, pickupDate: { type: 'STRING' }, pickupDaysBefore: { type: 'NUMBER' }, pickupMethod: { type: 'STRING' },
    payMethod: { type: 'STRING' }, payDetail: { type: 'STRING' },
    uncertain: { type: 'ARRAY', items: { type: 'STRING' } },
  },
  required: ['eventName', 'date', 'seats', 'currency', 'totalPaid', 'ticketCount'],
};

const today_ = function () { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); };

function scanTicket_(req) {
  if (req.test === 'quota') return { error: '今天的 AI 免費次數用完了（這是模擬測試），可以明天再試，或先手動填寫' };
  return readTicketImages_(req.images || []);
}

function readTicketImages_(images) {
  const parts = [{ text: SCAN_PROMPT.replace('{TODAY}', today_()) }];
  images.forEach(function (img) {
    parts.push({ inline_data: { mime_type: img.mimeType || 'image/jpeg', data: img.dataBase64 } });
  });
  return gemini_(parts, SCAN_SCHEMA);
}

function gemini_(parts, schema) {
  const apiKey = PROPS.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: '還沒設定 Gemini 金鑰：請到 Apps Script「專案設定 → 指令碼屬性」新增 GEMINI_API_KEY' };
  // 讀票不需要深度思考；思考太久會讓手機瀏覽器等超過 60 秒而斷線
  const config = { responseMimeType: 'application/json', responseSchema: schema, temperature: 0, thinkingConfig: { thinkingLevel: 'low' } };
  const call = function (model, cfg) {
    return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
      method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: cfg }),
      headers: { 'x-goog-api-key': apiKey }, muteHttpExceptions: true,
    });
  };
  let lastCode = 0;
  let lastMsg = '';
  for (let i = 0; i < GEMINI_MODELS.length; i++) {
    let res;
    try {
      res = call(GEMINI_MODELS[i], config);
      if (res.getResponseCode() === 400) { // 模型不支援 thinkingConfig 時，拿掉再試一次
        const plain = Object.assign({}, config);
        delete plain.thinkingConfig;
        res = call(GEMINI_MODELS[i], plain);
      }
    } catch (err) {
      lastCode = -1;
      lastMsg = String(err.message || err);
      continue; // 逾時或連線失敗，換下一個模型
    }
    lastCode = res.getResponseCode();
    if (lastCode !== 200) lastMsg = res.getContentText().slice(0, 300);
    if (lastCode === 200) {
      const out = JSON.parse(res.getContentText());
      const cand = (out.candidates || [])[0];
      const text = cand && cand.content ? cand.content.parts.filter(function (p) { return p.text && !p.thought; })
        .map(function (p) { return p.text; }).join('') : '';
      if (!text) return { error: 'AI 沒有讀出內容，請換一張清楚的截圖或手動填寫' };
      return { ok: true, model: GEMINI_MODELS[i], data: JSON.parse(text) };
    }
    if (lastCode === 401 || lastCode === 403) break; // 金鑰有問題，換模型也沒用
  }
  if (lastCode === 429) return { error: '今天的 AI 免費次數用完了，可以明天再試，或先手動填寫' };
  if (lastCode === 401 || lastCode === 403) return { error: 'Gemini 金鑰無效或沒有權限（代碼 ' + lastCode + '），請確認指令碼屬性 GEMINI_API_KEY' };
  return { error: 'AI 暫時無法使用（代碼 ' + lastCode + '）：' + lastMsg };
}

// 檢查工具：在 Apps Script 編輯器選 testGemini →「執行」，下方「執行記錄」會顯示金鑰與模型是否可用
function testGemini() {
  const apiKey = PROPS.getProperty('GEMINI_API_KEY');
  if (!apiKey) { Logger.log('找不到指令碼屬性 GEMINI_API_KEY（名稱要完全一樣，大寫、底線）'); return; }
  Logger.log('金鑰長度 ' + apiKey.length + '，開頭 ' + apiKey.slice(0, 4));
  GEMINI_MODELS.forEach(function (m) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m, {
      headers: { 'x-goog-api-key': apiKey }, muteHttpExceptions: true,
    });
    Logger.log(m + ' → ' + res.getResponseCode() + (res.getResponseCode() === 200 ? ' 可用' : ' ' + res.getContentText().slice(0, 300)));
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  return !SHARED_KEY || key === SHARED_KEY;
}

function doGet(e) {
  if (!checkKey_(e.parameter.key)) return json_({ error: 'bad key' });
  setup_();
  const out = {};
  Object.keys(TABLES).forEach(function (name) { out[name] = readTable_(name); });
  return json_(out);
}

function doPost(e) {
  if (e.parameter && e.parameter.line !== undefined) return lineWebhook_(e); // LINE 機器人
  const req = JSON.parse(e.postData.contents);
  if (!checkKey_(req.key)) return json_({ error: 'bad key' });
  if (req.action === 'scanTicket') { // 不寫試算表，不用排隊
    try {
      return json_(scanTicket_(req));
    } catch (err) {
      return json_({ error: '讀圖時發生錯誤：' + String(err.message || err) });
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setup_();
    if (req.action === 'upsert') writeRows_(req.table, req.rows);
    else if (req.action === 'delete') deleteRows_(req.table, req.ids);
    else if (req.action === 'replaceAll') replaceAll_(req.data);
    else if (req.action === 'uploadImage') return json_(uploadImage_(req));
    else return json_({ error: 'unknown action' });
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ================= LINE 機器人 ================= */
// 指令碼屬性 LINE_TOKEN：LINE Developers 後台的 Channel access token
// LINE 後台 Webhook URL：網頁應用程式網址 + ?line=共用密碼（例：…/exec?line=abc123）
// 朋友加好友後輸入共用密碼（邀請碼）即可使用；名單存在指令碼屬性 LINE_USERS
const LINE_PLATFORMS = ['拓元', 'KKTIX', 'ibon', '年代', '寬宏', '遠大', 'NOL', 'Melon', 'YES24'];
const LINE_PLATFORM_ALIAS = { interpark: 'NOL', 'nol ticket': 'NOL', 'nol interpark': 'NOL', tixcraft: '拓元', kham: '寬宏', 'era ticket': '年代' };
const PAY_LABEL_ = { credit_card: '信用卡', bank_transfer: '轉帳', mobile_payment: '行動支付', cash: '現金' };
const PICKUP_METHODS_ = ['電子票', '超商取票', '現場取票', '宅配'];
const LINE_STEPS = ['purpose', 'eventType', 'platform', 'account', 'payMethod', 'payDetail', 'payer', 'pickupMethod', 'pickup'];
const OPTIONAL_STEPS = ['account', 'payDetail', 'pickup']; // 可以跳過的題目
const AI_FIELDS = ['eventName', 'artist', 'date', 'time', 'city', 'venue', 'eventType', 'ticketType', 'area', 'row', 'seat',
  'ticketCount', 'currency', 'unitFace', 'unitBenefit', 'unitFee', 'totalPaid', 'orderNumber', 'pickupDate', 'pickupMethod',
  'platform', 'account', 'payMethod', 'payDetail'];
const LINE_HELP = '傳購票截圖給我（同一筆訂單可一次傳 2～3 張）→ 回答幾個問題 → 確認卡片按「確認建檔」就完成了 ✦\n隨時輸入「取消」可以放棄目前這筆。';

const num_ = function (v) {
  const n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
};
const str_ = function (v) { return String(v == null ? '' : v).trim(); };
const newId_ = function () { return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8); };
const cache_ = function () { return CacheService.getScriptCache(); };
const normName_ = function (v) { return String(v || '').toLowerCase().replace(/[\s\-–—_:：·・,，.。!！?？'"「」『』()（）[\]【】<>＜＞]/g, ''); };

function lineSession_(uid) {
  const raw = cache_().get('ls_' + uid);
  return raw ? JSON.parse(raw) : null;
}
function saveLineSession_(uid, s) {
  cache_().put('ls_' + uid, JSON.stringify(s), 21600); // 6 小時沒動作自動清掉
  PROPS.setProperty('ld_' + uid, '1');
}
function clearLineSession_(uid) {
  cache_().remove('ls_' + uid);
  PROPS.deleteProperty('ld_' + uid);
}
function lineUsers_() { return JSON.parse(PROPS.getProperty('LINE_USERS') || '{}'); }
function saveLineUsers_(users) { PROPS.setProperty('LINE_USERS', JSON.stringify(users)); }

function lineApi_(url, payload) {
  return UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('LINE_TOKEN') }, muteHttpExceptions: true,
  });
}
function lineReply_(token, messages) {
  messages = (Array.isArray(messages) ? messages : [messages]).filter(Boolean).slice(0, 5)
    .map(function (m) { return typeof m === 'string' ? { type: 'text', text: m } : m; });
  lineApi_('https://api.line.me/v2/bot/message/reply', { replyToken: token, messages: messages });
}
function lineLoading_(uid) {
  lineApi_('https://api.line.me/v2/bot/chat/loading/start', { chatId: uid, loadingSeconds: 30 });
}
function lineImage_(id) {
  const res = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/message/' + id + '/content', {
    headers: { Authorization: 'Bearer ' + PROPS.getProperty('LINE_TOKEN') }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('讀不到 LINE 上的圖片（代碼 ' + res.getResponseCode() + '）');
  const blob = res.getBlob();
  return { mimeType: blob.getContentType() || 'image/jpeg', dataBase64: Utilities.base64Encode(blob.getBytes()) };
}
function ask_(text, step, options) {
  const m = { type: 'text', text: text };
  if (options && options.length) {
    m.quickReply = { items: options.slice(0, 13).map(function (o) {
      const label = String(o[1]).slice(0, 20);
      return { type: 'action', action: { type: 'postback', label: label, data: 'a=' + step + '&v=' + encodeURIComponent(o[0]), displayText: label } };
    }) };
  }
  return m;
}

function lineWebhook_(e) {
  if (SHARED_KEY && e.parameter.line !== SHARED_KEY) return json_({ ok: false });
  const body = JSON.parse((e.postData && e.postData.contents) || '{}');
  (body.events || []).forEach(function (ev) {
    if (ev.deliveryContext && ev.deliveryContext.isRedelivery) return;
    try {
      handleLineEvent_(ev);
    } catch (err) {
      if (ev.replyToken) lineReply_(ev.replyToken, '出了點問題：' + String(err.message || err) + '\n可以再試一次，或改用網站掃票。');
    }
  });
  return json_({ ok: true });
}

function handleLineEvent_(ev) {
  const uid = ev.source && ev.source.userId;
  if (!uid || ev.source.type !== 'user') return;
  const users = lineUsers_();
  const me = users[uid];
  const token = ev.replyToken;
  const text = ev.type === 'message' && ev.message.type === 'text' ? ev.message.text.trim() : '';
  const pb = ev.type === 'postback' ? parsePostback_(ev.postback.data) : null;

  if (ev.type === 'follow') {
    lineReply_(token, me ? '歡迎回來 ✦\n' + LINE_HELP : '你好！這是追星記票機器人 ✦\n請先輸入邀請碼才能開始使用。');
    return;
  }
  if (!me) {
    if (text && (!SHARED_KEY || text === SHARED_KEY)) {
      users[uid] = { name: '', joined: today_() };
      saveLineUsers_(users);
      saveLineSession_(uid, { step: 'name' });
      lineReply_(token, ask_('邀請碼正確 ✦ 請問怎麼稱呼你？（分帳時用來認出你的那一份，也可以直接打字）', 'name', topPayers_().map(function (p) { return [p, p]; })));
    } else if (token) {
      lineReply_(token, '請先輸入邀請碼才能使用喔（向帳號主人索取）。');
    }
    return;
  }

  if (pb && pb.a === 'name') return setName_(token, uid, users, pb.v);
  if (ev.type === 'message' && ev.message.type === 'image') return onLineImage_(ev, uid, me);
  if (pb) return onLinePostback_(token, uid, me, pb);
  if (text) return onLineText_(token, uid, me, users, text);
}

function parsePostback_(data) {
  const out = {};
  String(data || '').split('&').forEach(function (kv) {
    const i = kv.indexOf('=');
    if (i > 0) out[kv.slice(0, i)] = decodeURIComponent(kv.slice(i + 1));
  });
  return out;
}

function setName_(token, uid, users, name) {
  name = str_(name).slice(0, 20);
  if (!name) return lineReply_(token, '請輸入你的名字。');
  users[uid].name = name;
  saveLineUsers_(users);
  clearLineSession_(uid);
  lineReply_(token, '設定完成，' + name + ' ✦\n' + LINE_HELP);
}

function onLineImage_(ev, uid, me) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let s;
  let ready = false;
  try {
    s = lineSession_(uid) || {};
    if (s.step === 'name') s = {};
    const target = s.d ? 'pending' : 'imgs'; // 已經在問答中 → 當作補傳
    const set = ev.message.imageSet;
    if (set) {
      if (!s.set || s.set.id !== set.id) s.set = { id: set.id, total: set.total, items: [] };
      if (!s.set.items.some(function (x) { return x.id === ev.message.id; })) s.set.items.push({ id: ev.message.id, index: set.index || 0 });
      if (s.set.items.length >= s.set.total) {
        s.set.items.sort(function (a, b) { return a.index - b.index; });
        s[target] = (s[target] || []).concat(s.set.items.map(function (x) { return x.id; }));
        delete s.set;
        ready = true;
      }
    } else {
      s[target] = (s[target] || []).concat([ev.message.id]);
      ready = true;
    }
    saveLineSession_(uid, s);
  } finally {
    lock.releaseLock();
  }
  if (!ready) return; // 同一批還有圖片沒到，等最後一張再一起讀
  lineLoading_(uid);
  if (s.d) checkSupplement_(ev.replyToken, uid, s);
  else startDraft_(ev.replyToken, uid, s);
}

function readLineImages_(ids) {
  let images;
  try {
    images = ids.map(lineImage_);
  } catch (err) {
    return { error: String(err.message || err) };
  }
  return readTicketImages_(images);
}

function startDraft_(token, uid, s) {
  const r = readLineImages_(s.imgs || []);
  if (r.error) {
    clearLineSession_(uid);
    return lineReply_(token, r.error + '\n請再傳一次截圖，或改用網站掃票。');
  }
  if (!str_(r.data.eventName) && !str_(r.data.date) && !num_(r.data.totalPaid)) {
    clearLineSession_(uid);
    return lineReply_(token, '看不出這是購票截圖耶，換一張試試看 🙏');
  }
  s.d = draftFrom_(r.data);
  s.asked = {};
  s.edited = [];
  lineReply_(token, ['讀到了 ✦\n' + summaryText_(s.d), nextMessage_(uid, s)]);
}

function draftFrom_(x) {
  const seats = Array.isArray(x.seats) ? x.seats : [];
  const uniq = function (k) {
    const out = [];
    seats.forEach(function (st) { const v = str_(st[k]); if (v && out.indexOf(v) < 0) out.push(v); });
    return out.join('、');
  };
  const cur = str_(x.currency).toUpperCase();
  const plat = str_(x.platform);
  const d = {
    eventName: str_(x.eventName), artist: str_(x.artist), date: str_(x.date), time: str_(x.time),
    city: str_(x.city), venue: str_(x.venue), ticketType: str_(x.ticketType), orderNumber: str_(x.orderNumber),
    area: uniq('area'), row: uniq('row'),
    seat: seats.map(function (st) { return str_(st.seat); }).filter(Boolean).join('、'),
    ticketCount: num_(x.ticketCount) || seats.length || 1,
    currency: ['TWD', 'KRW', 'JPY', 'USD'].indexOf(cur) >= 0 ? cur : 'TWD',
    unitFace: num_(x.unitFace), unitBenefit: num_(x.unitBenefit), unitFee: num_(x.unitFee),
    eventType: ['專場', '拼盤'].indexOf(str_(x.eventType)) >= 0 ? str_(x.eventType) : '',
    pickupMethod: PICKUP_METHODS_.indexOf(str_(x.pickupMethod)) >= 0 ? str_(x.pickupMethod) : '',
    pickupDate: str_(x.pickupDate),
    platform: plat ? (LINE_PLATFORM_ALIAS[plat.toLowerCase()] || LINE_PLATFORMS.filter(function (p) { return p.toLowerCase() === plat.toLowerCase(); })[0] || plat) : '',
    account: str_(x.account),
    payMethod: PAY_LABEL_[str_(x.payMethod)] ? str_(x.payMethod) : '',
    payDetail: str_(x.payDetail),
    purpose: '', payer: '',
  };
  d.totalPaid = num_(x.totalPaid) || (d.unitFace + d.unitBenefit + d.unitFee) * d.ticketCount;
  if (!d.pickupDate && num_(x.pickupDaysBefore)) d.pickupDate = minusDays_(d.date, num_(x.pickupDaysBefore));
  fillTwd_(d);
  return d;
}

function minusDays_(date, days) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !days) return '';
  const t = new Date(date + 'T00:00:00Z');
  t.setUTCDate(t.getUTCDate() - days);
  return t.toISOString().slice(0, 10);
}

// 匯率慣例：KRW、JPY 為 TWD 1 = 原幣 X；USD 為 USD 1 = TWD X
function fxRate_(cur) {
  const hit = cache_().get('fx_' + cur);
  if (hit) return Number(hit);
  try {
    const d = JSON.parse(UrlFetchApp.fetch('https://open.er-api.com/v6/latest/TWD', { muteHttpExceptions: true }).getContentText());
    const x = d.rates && d.rates[cur];
    if (!x) return 0;
    const rate = Number((cur === 'USD' ? 1 / x : x).toFixed(4));
    cache_().put('fx_' + cur, String(rate), 21600);
    return rate;
  } catch (err) {
    return 0;
  }
}
function fillTwd_(d) {
  if (d.currency === 'TWD') {
    d.amountTwd = d.totalPaid;
    d.exchangeRate = '';
    d.twdEstimated = false;
    return;
  }
  const rate = fxRate_(d.currency);
  d.exchangeRate = rate || '';
  d.amountTwd = rate ? Math.round(d.currency === 'USD' ? d.totalPaid * rate : d.totalPaid / rate) : '';
  d.twdEstimated = true;
}

const money_ = function (cur, n) { return (cur || 'TWD') + ' ' + num_(n).toLocaleString('en-US'); };
const seatText_ = function (l) {
  const part = function (v, suf) { v = str_(v); return v ? (v.slice(-1) === suf ? v : v + suf) : ''; };
  return [str_(l.ticketArea), part(l.ticketRow, '排'), part(l.ticketSeat, '號')].filter(Boolean).join(' ');
};
function summaryText_(d) {
  return [
    d.eventName || '（沒讀到活動名稱）',
    [d.date, d.time].filter(Boolean).join(' ') + (d.venue ? '｜' + [d.city, d.venue].filter(Boolean).join(' ') : ''),
    [seatText_({ ticketArea: d.area, ticketRow: d.row, ticketSeat: d.seat }), d.ticketCount + ' 張', money_(d.currency, d.totalPaid)].filter(Boolean).join('｜'),
  ].filter(Boolean).join('\n');
}

function stepDone_(s, step) {
  if (OPTIONAL_STEPS.indexOf(step) >= 0 && s.asked[step]) return true;
  if (step === 'pickup') return !!s.d.pickupDate;
  return !!s.d[step];
}
function nextMessage_(uid, s) {
  for (let i = 0; i < LINE_STEPS.length; i++) {
    if (!stepDone_(s, LINE_STEPS[i])) {
      s.step = LINE_STEPS[i];
      saveLineSession_(uid, s);
      return question_(s.step, s.d);
    }
  }
  s.step = 'confirm';
  saveLineSession_(uid, s);
  return card_(s.d);
}
function question_(step, d) {
  const skip = [['__skip', '跳過']];
  const pairs = function (list) { return list.map(function (x) { return [x, x]; }); };
  if (step === 'purpose') return ask_('這張是自己要去的，還是要轉賣？', step, [['self', '自己去'], ['resale', '要轉賣']]);
  if (step === 'eventType') return ask_('專場還是拼盤？', step, [['專場', '專場'], ['拼盤', '拼盤']]);
  if (step === 'platform') return ask_('在哪個平台買的？（其他平台可以直接打字）', step, pairs(LINE_PLATFORMS));
  if (step === 'account') return ask_('用哪個帳號買的？可以直接打字（名字、信箱、帳號或電話）', step, pairs(knownValues_('ticketAccount', d.platform)).concat(skip));
  if (step === 'payMethod') return ask_('付款方式？', step, Object.keys(PAY_LABEL_).map(function (k) { return [k, PAY_LABEL_[k]]; }));
  if (step === 'payDetail') return ask_('哪張卡或哪個支付平台？可以直接打字（例：永豐、LINE Pay）', step, pairs(knownValues_('paymentDetail')).concat(skip));
  if (step === 'payer') return ask_('付款人是誰？（其他人可以直接打字名字）', step, pairs(topPayers_()));
  if (step === 'pickupMethod') return ask_('取票方式？（其他方式可以直接打字）', step, pairs(PICKUP_METHODS_));
  return ask_('什麼時候可以取票？可以直接打字，例如「演出前5天」或「11/1」', 'pickup',
    [['d3', '演出前 3 天'], ['d7', '演出前 7 天'], ['any', '隨時'], ['unknown', '不確定']]);
}

// 從總帳統計常用值，越常用越前面
function countValues_(rows, key) {
  const cnt = {};
  rows.forEach(function (r) { const v = str_(r[key]); if (v) cnt[v] = (cnt[v] || 0) + 1; });
  return Object.keys(cnt).sort(function (a, b) { return cnt[b] - cnt[a]; });
}
function ledgerRows_() {
  const sheet = ss_().getSheetByName('ledger');
  return sheet ? readTable_('ledger') : [];
}
function knownValues_(key, platform) {
  const rows = ledgerRows_().filter(function (l) { return !platform || l.ticketPlatform === platform; });
  return countValues_(rows, key).slice(0, 6);
}
function topPayers_() {
  return countValues_(ledgerRows_().filter(function (l) { return l.category === 'ticket'; }), 'payer').slice(0, 3);
}

function parsePickup_(text, d) {
  const m = text.match(/(\d+)\s*[天日]/);
  if (m && /前/.test(text)) return minusDays_(d.date, Number(m[1]));
  let y, mo, da;
  const full = text.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  const short = text.match(/(\d{1,2})[\/\-.月](\d{1,2})/);
  if (full) { y = Number(full[1]); mo = Number(full[2]); da = Number(full[3]); }
  else if (short) {
    mo = Number(short[1]); da = Number(short[2]);
    y = Number((d.date || today_()).slice(0, 4));
  } else return '';
  if (mo < 1 || mo > 12 || da < 1 || da > 31) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + da).slice(-2);
}

// 回答一題；回傳 false 表示看不懂要重問
function applyAnswer_(s, step, v, typed) {
  const d = s.d;
  v = str_(v);
  if (!v) return false;
  if (step === 'purpose') {
    if (typed) v = /轉|賣|讓/.test(v) ? 'resale' : /自己|去|我/.test(v) ? 'self' : '';
    if (v !== 'self' && v !== 'resale') return false;
    d.purpose = v;
  } else if (step === 'eventType') {
    if (typed) v = /拼/.test(v) ? '拼盤' : /專/.test(v) ? '專場' : '';
    if (!v) return false;
    d.eventType = v;
  } else if (step === 'payMethod') {
    if (typed) v = /信用|刷卡|卡/.test(v) ? 'credit_card' : /轉帳|匯款|ATM/i.test(v) ? 'bank_transfer'
      : /行動|pay|支付/i.test(v) ? 'mobile_payment' : /現金/.test(v) ? 'cash' : '';
    if (!PAY_LABEL_[v]) return false;
    d.payMethod = v;
  } else if (step === 'pickup') {
    if (v === 'd3' || v === 'd7') d.pickupDate = minusDays_(d.date, v === 'd3' ? 3 : 7);
    else if (v === 'any' || v === 'unknown' || /隨時|不確定|不知道|跳過/.test(v)) d.pickupDate = '';
    else {
      const date = parsePickup_(v, d);
      if (!date) return false;
      d.pickupDate = date;
    }
  } else if (step === 'pickupMethod') {
    d.pickupMethod = PICKUP_METHODS_.filter(function (m) { return m.indexOf(v) >= 0 || v.indexOf(m.slice(0, 2)) >= 0; })[0] || v;
  } else if (step === 'platform') {
    d.platform = LINE_PLATFORM_ALIAS[v.toLowerCase()] || LINE_PLATFORMS.filter(function (p) { return p.toLowerCase() === v.toLowerCase(); })[0] || v;
  } else if (v === '__skip') {
    if (OPTIONAL_STEPS.indexOf(step) < 0) return false;
  } else {
    d[step] = v;
  }
  s.asked[step] = true;
  return true;
}

function card_(d) {
  const row = function (k, v) {
    return { type: 'box', layout: 'baseline', spacing: 'md', contents: [
      { type: 'text', text: k, size: 'sm', color: '#9C917F', flex: 2 },
      { type: 'text', text: str_(v) || '—', size: 'sm', color: '#40382E', flex: 5, wrap: true },
    ] };
  };
  const unit = ['票面 ' + money_(d.currency, d.unitFace)];
  if (num_(d.unitBenefit)) unit.push('福利 ' + money_(d.currency, d.unitBenefit));
  if (num_(d.unitFee)) unit.push('手續費 ' + money_(d.currency, d.unitFee));
  const total = money_(d.currency, d.totalPaid) + (d.currency !== 'TWD'
    ? '\n約 TWD ' + num_(d.amountTwd).toLocaleString('en-US') + (d.twdEstimated ? '（估算）' : '') : '');
  const button = function (label, v, style) {
    return { type: 'button', style: style, height: 'sm', color: style === 'primary' ? '#A67B5B' : undefined,
      action: { type: 'postback', label: label, data: 'a=card&v=' + v, displayText: label } };
  };
  return {
    type: 'flex', altText: '請確認：' + (d.eventName || '購票紀錄'),
    contents: {
      type: 'bubble',
      body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
        { type: 'text', text: d.purpose === 'resale' ? '要轉賣' : '自己去', size: 'xs', color: '#A67B5B', weight: 'bold' },
        { type: 'text', text: d.eventName || '（沒有活動名稱）', weight: 'bold', size: 'lg', wrap: true, color: '#40382E' },
        { type: 'text', text: [d.date, d.time].filter(Boolean).join(' ') + (d.venue ? '｜' + [d.city, d.venue].filter(Boolean).join(' ') : '') || '—', size: 'sm', color: '#9C917F', wrap: true },
        { type: 'separator', margin: 'md' },
        row('類型', [d.eventType, d.artist].filter(Boolean).join('｜')),
        row('票種', d.ticketType),
        row('座位', seatText_({ ticketArea: d.area, ticketRow: d.row, ticketSeat: d.seat })),
        row('張數', d.ticketCount + ' 張'),
        row('單張', unit.join('\n')),
        row('總額', total),
        { type: 'separator', margin: 'md' },
        row('平台', [d.platform, d.account].filter(Boolean).join('｜')),
        row('付款', [PAY_LABEL_[d.payMethod] || '', d.payDetail].filter(Boolean).join(' ')),
        row('付款人', d.payer),
        row('取票', [d.pickupMethod, d.pickupDate || '隨時'].filter(Boolean).join('｜')),
        row('訂單', d.orderNumber),
        { type: 'text', text: '要修改直接打字告訴我，例如「座位改成5排18號」', size: 'xxs', color: '#9C917F', wrap: true, margin: 'md' },
      ] },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
        button('確認建檔', 'confirm', 'primary'), button('要修改', 'edit', 'secondary'), button('取消這筆', 'cancel', 'link'),
      ] },
    },
  };
}

const EDIT_SCHEMA = {
  type: 'OBJECT',
  properties: {
    understood: { type: 'BOOLEAN' },
    eventName: { type: 'STRING' }, artist: { type: 'STRING' }, date: { type: 'STRING' }, time: { type: 'STRING' },
    city: { type: 'STRING' }, venue: { type: 'STRING' }, eventType: { type: 'STRING' }, ticketType: { type: 'STRING' },
    area: { type: 'STRING' }, row: { type: 'STRING' }, seat: { type: 'STRING' }, ticketCount: { type: 'NUMBER' },
    currency: { type: 'STRING' }, unitFace: { type: 'NUMBER' }, unitBenefit: { type: 'NUMBER' }, unitFee: { type: 'NUMBER' },
    totalPaid: { type: 'NUMBER' }, amountTwd: { type: 'NUMBER' }, orderNumber: { type: 'STRING' },
    pickupDate: { type: 'STRING' }, pickupMethod: { type: 'STRING' }, platform: { type: 'STRING' }, account: { type: 'STRING' },
    payMethod: { type: 'STRING' }, payDetail: { type: 'STRING' }, payer: { type: 'STRING' }, purpose: { type: 'STRING' },
  },
  required: ['understood'],
};
function editDraft_(d, text) {
  const view = {};
  Object.keys(EDIT_SCHEMA.properties).forEach(function (k) { if (k !== 'understood') view[k] = d[k]; });
  const prompt = [
    '以下是一筆購票紀錄（JSON）：', JSON.stringify(view),
    '使用者說：「' + text + '」',
    '請輸出 understood: true，並只包含使用者要修改的欄位，沒提到的欄位不要輸出。',
    'purpose 只能是 self（自己去）或 resale（要轉賣）；payMethod 只能是 credit_card、bank_transfer、mobile_payment、cash；',
    'eventType 只能是 專場 或 拼盤；日期用 YYYY-MM-DD，時間用 HH:MM；座位的排、號只填數字或代號。',
    '如果「演出前 N 天取票」，pickupDate 用演出日期往前推 N 天。看不懂使用者要改什麼，就只輸出 understood: false。',
  ].join('\n');
  const r = gemini_([{ text: prompt }], EDIT_SCHEMA);
  if (r.error) return { error: r.error };
  const changes = r.data || {};
  if (!changes.understood) return { changed: [] };
  const changed = [];
  Object.keys(changes).forEach(function (k) {
    if (k === 'understood' || !(k in view)) return;
    d[k] = typeof changes[k] === 'number' ? changes[k] : str_(changes[k]);
    changed.push(k);
  });
  if (changed.some(function (k) { return ['currency', 'totalPaid'].indexOf(k) >= 0; }) && changed.indexOf('amountTwd') < 0) fillTwd_(d);
  if (changed.indexOf('amountTwd') >= 0) d.twdEstimated = false;
  return { changed: changed };
}

function onLineText_(token, uid, me, users, text) {
  let s = lineSession_(uid);
  if (/^(取消|cancel)$/i.test(text)) {
    clearLineSession_(uid);
    return lineReply_(token, s && s.d ? '已取消這筆，沒有建檔。' : '目前沒有進行中的紀錄。');
  }
  if (/^(說明|help|\?|？)$/i.test(text)) return lineReply_(token, LINE_HELP);
  if (!s) {
    if (PROPS.getProperty('ld_' + uid)) {
      PROPS.deleteProperty('ld_' + uid);
      return lineReply_(token, '上一筆放太久，已經自動清除了，請重新傳截圖 🙏');
    }
    return lineReply_(token, '傳購票截圖給我就可以開始記錄 ✦');
  }
  if (s.step === 'name') return setName_(token, uid, users, text);
  if (s.step === 'supplement') return lineReply_(token, ask_('請先選擇這張新截圖要怎麼處理：', 'supp', [['merge', '補充到這筆'], ['new', '當成新的一筆']]));
  if (s.step === 'confirm' || s.step === 'dupe') {
    lineLoading_(uid);
    const r = editDraft_(s.d, text);
    if (r.error) return lineReply_(token, r.error);
    if (!r.changed.length) return lineReply_(token, '看不太懂要改哪裡，可以說得更具體一點，例如「座位改成5排18號」「付款人改成 Jhen」。');
    s.edited = (s.edited || []).concat(r.changed);
    delete s.eventChoice;
    return lineReply_(token, ['已修改 ✦', nextMessage_(uid, s)]);
  }
  if (LINE_STEPS.indexOf(s.step) >= 0) {
    if (applyAnswer_(s, s.step, text, true)) return lineReply_(token, nextMessage_(uid, s));
    const again = question_(s.step, s.d);
    again.text = '請點下面的按鈕選擇，或換個說法 🙏\n' + again.text;
    return lineReply_(token, again);
  }
  lineReply_(token, '還在等同一批的其他截圖，稍等一下喔。');
}

function onLinePostback_(token, uid, me, pb) {
  const s = lineSession_(uid);
  if (!s || !s.d) return lineReply_(token, '這筆已經結束或放太久被清除了，請重新傳截圖 🙏');
  if (LINE_STEPS.indexOf(pb.a) >= 0) {
    if (!applyAnswer_(s, pb.a, pb.v, false)) return lineReply_(token, question_(pb.a, s.d));
    return lineReply_(token, nextMessage_(uid, s));
  }
  if (pb.a === 'supp') {
    if (pb.v === 'merge') return mergeSupplement_(token, uid, s, s.pendingData);
    const fresh = { imgs: s.pending || [] };
    saveLineSession_(uid, fresh);
    lineLoading_(uid);
    return startDraft_(token, uid, fresh);
  }
  if (pb.a === 'card') {
    if (pb.v === 'cancel') {
      clearLineSession_(uid);
      return lineReply_(token, '已取消這筆，沒有建檔。');
    }
    if (pb.v === 'edit') return lineReply_(token, '直接打字告訴我要改什麼就好，例如：\n「座位改成5排18號」\n「付款人改成 Jhen」\n「演出前5天取票」');
    if (s.step !== 'confirm') return lineReply_(token, nextMessage_(uid, s));
    return confirmDraft_(token, uid, me, s);
  }
  if (pb.a === 'dupe') {
    s.eventChoice = pb.v;
    return confirmDraft_(token, uid, me, s);
  }
}

function checkSupplement_(token, uid, s) {
  const r = readLineImages_(s.pending || []);
  if (r.error) {
    s.pending = [];
    saveLineSession_(uid, s);
    return lineReply_(token, r.error + '\n這張補充截圖沒有讀到，目前這筆不受影響。');
  }
  const x = r.data;
  const same = (str_(x.orderNumber) && str_(x.orderNumber) === s.d.orderNumber)
    || (str_(x.eventName) && normName_(x.eventName) === normName_(s.d.eventName) && (!str_(x.date) || str_(x.date) === s.d.date));
  if (same) return mergeSupplement_(token, uid, s, x);
  s.pendingData = x;
  s.stepBefore = s.step;
  s.step = 'supplement';
  saveLineSession_(uid, s);
  lineReply_(token, ask_('這張截圖和目前這筆（' + (s.d.eventName || '未命名') + '）看起來不太一樣，要怎麼處理？', 'supp',
    [['merge', '補充到這筆'], ['new', '當成新的一筆']]));
}

// 補傳的截圖只補上原本空著的欄位，你回答過或修改過的不動
function mergeSupplement_(token, uid, s, x) {
  const nd = draftFrom_(x || {});
  const keep = (s.edited || []).concat(Object.keys(s.asked || {}));
  const filled = [];
  AI_FIELDS.forEach(function (k) {
    if (keep.indexOf(k) >= 0) return;
    const empty = s.d[k] === '' || s.d[k] === 0 || s.d[k] == null || (k === 'ticketCount' && s.d[k] === 1);
    if (empty && nd[k] !== '' && nd[k] !== 0 && nd[k] != null) { s.d[k] = nd[k]; filled.push(k); }
  });
  if (filled.indexOf('totalPaid') >= 0 || filled.indexOf('currency') >= 0) fillTwd_(s.d);
  s.imgs = (s.imgs || []).concat(s.pending || []);
  s.pending = [];
  delete s.pendingData;
  if (s.step === 'supplement') s.step = s.stepBefore;
  delete s.stepBefore;
  lineReply_(token, [(filled.length ? '已合併，補上了新的資訊 ✦\n' : '已合併 ✦（沒有新增的資訊）\n') + summaryText_(s.d), nextMessage_(uid, s)]);
}

function confirmDraft_(token, uid, me, s) {
  const d = s.d;
  if (!d.eventName) return lineReply_(token, '還缺活動名稱，直接打字告訴我，例如「活動名稱是 DAY6 首爾場」。');
  if (d.purpose === 'self' && !/^\d{4}-\d{2}-\d{2}$/.test(d.date)) return lineReply_(token, '還缺演出日期，直接打字告訴我，例如「日期是 11/8」。');
  if (d.purpose === 'self' && !s.eventChoice) {
    const dupes = sameEvents_(d);
    if (dupes.length) {
      s.step = 'dupe';
      saveLineSession_(uid, s);
      return lineReply_(token, ask_('總帳裡已經有同一天的場次，要加到既有場次，還是另建新的？', 'dupe',
        dupes.slice(0, 3).map(function (e) { return [e.id, '加到 ' + e.name]; }).concat([['new', '另建新場次']])));
    }
  }
  const result = commitLine_(d, me, s.eventChoice && s.eventChoice !== 'new' ? s.eventChoice : '');
  clearLineSession_(uid);
  lineReply_(token, result.kind === 'event'
    ? '已建檔 ✦ ' + result.title + '\n可以到網站「活動」頁查看。'
    : '已建檔 ✦ 記到「讓票」頁了：' + result.title);
}

function sameEvents_(d) {
  const n = normName_(d.eventName), ar = normName_(d.artist);
  return readTable_('events').filter(function (e) {
    if (!d.date || e.startDate !== d.date) return false;
    const en = normName_(e.name);
    return (n && en && (en.indexOf(n) >= 0 || n.indexOf(en) >= 0)) || (ar && normName_(e.artist) === ar);
  });
}

function ownShare_(l, myName) {
  let sp = [];
  try { sp = JSON.parse(l.splits || '[]'); } catch (err) { sp = []; }
  const sum = function (list) { return list.reduce(function (t, x) { return t + num_(x.amountTwd); }, 0); };
  if (sp.length) {
    const mine = sp.filter(function (x) { return myName && x.name === myName; });
    if (mine.length) return sum(mine);
    return Math.max(0, num_(l.amountTwd) - sum(sp.filter(function (x) { return x.name !== myName; })));
  }
  return Math.max(0, num_(l.amountTwd) - num_(l.expectedReceivableTwd));
}

// 與網站相同：活動依日期排序後重新編號，回傳編號有變的活動
function renumber_(events) {
  const n = function (e) { return num_(e.eventNumber); };
  const numbered = events.filter(function (e) { return n(e) && e.startDate; });
  const effKey = function (e) {
    if (e.startDate) return String(e.startDate);
    if (!n(e)) return '9999';
    let best = '';
    numbered.forEach(function (x) { if (n(x) <= n(e) && String(x.startDate) > best) best = String(x.startDate); });
    return best || '0000';
  };
  const sorted = events.slice().sort(function (a, b) {
    return effKey(a).localeCompare(effKey(b)) || (n(a) - n(b)) || String(a.createdAt || '').localeCompare(String(b.createdAt || ''));
  });
  const changed = [];
  sorted.forEach(function (e, i) {
    const k = String(i + 1);
    if (String(e.eventNumber) !== k) { e.eventNumber = k; changed.push(e); }
  });
  return changed;
}

function commitLine_(d, me, existingId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setup_();
    const foreign = d.currency !== 'TWD';
    const count = Math.max(1, num_(d.ticketCount) || 1);
    const amountTwd = Math.round(num_(foreign ? d.amountTwd : d.totalPaid));
    const notes = [];
    if (foreign) {
      notes.push('單張票面 ' + money_(d.currency, d.unitFace) + (num_(d.unitBenefit) ? '＋福利 ' + money_(d.currency, d.unitBenefit) : '')
        + (num_(d.unitFee) ? '＋手續費 ' + money_(d.currency, d.unitFee) : ''));
      if (d.twdEstimated) notes.push('台幣實付為估算，尚未對帳單確認');
    }
    notes.push('LINE 建檔：' + (me.name || '未命名'));

    if (d.purpose === 'resale') {
      const title = d.artist && d.eventName.indexOf(d.artist) < 0 ? d.artist + ' - ' + d.eventName : d.eventName;
      const linked = sameEvents_(d)[0];
      const info = [
        d.platform && '平台：' + d.platform, d.account && '帳號：' + d.account,
        (d.payMethod || d.payDetail) && '付款：' + [PAY_LABEL_[d.payMethod] || '', d.payDetail].filter(Boolean).join(' '),
        d.payer && '付款人：' + d.payer, d.orderNumber && '訂單：' + d.orderNumber,
        d.time && '開演 ' + d.time, d.venue && [d.city, d.venue].filter(Boolean).join(' '),
        d.ticketType && '票種：' + d.ticketType, d.eventType,
        d.pickupMethod && '取票：' + d.pickupMethod, d.pickupDate && '領票日 ' + d.pickupDate,
      ].filter(Boolean);
      writeRows_('transfers', [{
        id: newId_(), date: d.date || today_(), eventId: linked ? linked.id : '', kind: '轉賣', person: '', title: title,
        ticketCount: count, ticketArea: d.area, ticketRow: d.row, ticketSeat: d.seat,
        costTwd: amountTwd || '', amountTwd: '', feeTwd: '', settled: false, notes: info.concat(notes).join('｜'), createdAt: today_(),
      }]);
      return { kind: 'transfer', title: title };
    }

    const events = readTable_('events');
    let ev = existingId ? events.filter(function (e) { return e.id === existingId; })[0] : null;
    if (ev) {
      if (!ev.startTime && d.time) ev.startTime = d.time;
      if (!ev.venue && d.venue) ev.venue = d.venue;
      if (!ev.city && d.city) ev.city = d.city;
      if (!ev.artist && d.artist) ev.artist = d.artist;
      if (!ev.eventType && d.eventType) ev.eventType = d.eventType;
    } else {
      ev = {
        id: newId_(), name: d.eventName, artist: d.artist, city: d.city, venue: d.venue, startDate: d.date, startTime: d.time,
        endDate: '', eventNumber: '', originalDate: '', eventType: d.eventType, liveTour: '', seriesEvent: '', seat: '',
        ticketPriceTwd: '', guest: '', payer: '', settled: false, notes: '', createdAt: today_(), coverUrl: '',
      };
      events.push(ev);
    }
    const ledger = {
      id: newId_(), type: 'expense', category: 'ticket', date: d.date || today_(), title: '票券 - ' + ev.name, eventId: ev.id,
      amountTwd: amountTwd || '', currency: foreign ? d.currency : '', originalAmount: foreign ? num_(d.totalPaid) || '' : '',
      exchangeRate: foreign ? num_(d.exchangeRate) || '' : '', payer: d.payer, paymentMethod: d.payMethod, paymentDetail: d.payDetail,
      counterparty: '', expectedReceivableTwd: '', receivedTwd: '', notes: notes.join('｜'),
      ticketType: d.ticketType, ticketArea: d.area, ticketRow: d.row, ticketSeat: d.seat, attendee: '', ticketStatus: '',
      createdAt: today_(), settled: '',
      ticketFaceTwd: foreign ? '' : num_(d.unitFace) || '', ticketBenefitTwd: foreign ? '' : num_(d.unitBenefit) || '',
      ticketFeeTwd: foreign ? '' : num_(d.unitFee) || '', ticketPlatform: d.platform, ticketAccount: d.account, ticketCount: count,
      splits: JSON.stringify(me.name && amountTwd ? [{ name: me.name, count: 1, amountTwd: Math.round(amountTwd / count), settled: false }] : []),
      ticketPickupDate: d.pickupDate, ticketPickedUp: false, ticketOrderNumber: d.orderNumber, ticketPickupMethod: d.pickupMethod,
    };
    writeRows_('ledger', [ledger]);
    // 活動的票價與座位由底下的票券彙整，和網站的規則相同
    const tickets = readTable_('ledger').filter(function (l) { return l.eventId === ev.id && l.category === 'ticket'; });
    ev.ticketPriceTwd = tickets.reduce(function (t, l) { return t + ownShare_(l, me.name); }, 0);
    const seats = tickets.map(seatText_).filter(Boolean);
    if (seats.length) ev.seat = seats.join(' ');
    const changed = renumber_(events);
    if (changed.indexOf(ev) < 0) changed.push(ev);
    writeRows_('events', changed);
    return { kind: 'event', title: ev.name };
  } finally {
    lock.releaseLock();
  }
}

// 檢查工具：在編輯器選 testLine →「執行」，確認 LINE 金鑰可用
function testLine() {
  const token = PROPS.getProperty('LINE_TOKEN');
  if (!token) { Logger.log('找不到指令碼屬性 LINE_TOKEN'); return; }
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  Logger.log('LINE → ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
  Logger.log('已加入的使用者：' + JSON.stringify(lineUsers_()));
}
