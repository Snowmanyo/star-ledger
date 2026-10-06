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

// 同一次執行裡只開一次試算表、每個分頁只讀一次（寫入後才重讀），回覆會快很多
let SS_ = null;
const TABLE_MEMO_ = {};
function ss_() {
  if (!SS_) SS_ = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  return SS_;
}

const TABLES = {
  orders: ['id', 'orderNumber', 'channel', 'orderDate', 'estimatedShipDate', 'actualShipDate', 'currency', 'domesticShipping', 'internationalShippingTwd', 'internationalShippingRateTwdPerKg', 'discountAmount', 'weightGrams', 'exchangeRate', 'chargedTwd', 'payer', 'paymentMethod', 'paymentDetail', 'settled', 'notes'],
  items: ['id', 'orderId', 'name', 'variant', 'unitPrice', 'quantity', 'ownership', 'proxyFor', 'arrived', 'sorted', 'proxyPaid', 'salePriceTwd', 'soldQuantity'],
  sales: ['id', 'sourceOrderId', 'sourceItemId', 'sourceOrderNumber', 'sourceChannel', 'name', 'variant', 'sourceCurrency', 'unitOriginalPrice', 'unitCostTwd', 'quantity', 'salePriceTwd', 'soldQuantity', 'managedByOwnership', 'createdAt'],
  events: ['id', 'name', 'artist', 'city', 'venue', 'startDate', 'endDate', 'eventNumber', 'originalDate', 'eventType', 'liveTour', 'seriesEvent', 'seat', 'ticketPriceTwd', 'guest', 'payer', 'settled', 'notes', 'createdAt', 'coverUrl', 'startTime'],
  ledger: ['id', 'type', 'category', 'date', 'title', 'eventId', 'amountTwd', 'currency', 'originalAmount', 'exchangeRate', 'payer', 'paymentMethod', 'paymentDetail', 'counterparty', 'expectedReceivableTwd', 'receivedTwd', 'notes', 'ticketType', 'ticketArea', 'ticketRow', 'ticketSeat', 'attendee', 'ticketStatus', 'createdAt', 'settled', 'ticketFaceTwd', 'ticketBenefitTwd', 'ticketFeeTwd', 'ticketPlatform', 'ticketAccount', 'ticketCount', 'splits', 'ticketPickupDate', 'ticketPickedUp', 'ticketOrderNumber', 'ticketPickupMethod'],
  transfers: ['id', 'date', 'eventId', 'kind', 'person', 'ticketCount', 'ticketArea', 'ticketRow', 'ticketSeat', 'costTwd', 'amountTwd', 'settled', 'notes', 'createdAt', 'title', 'feeTwd',
    'artist', 'eventName', 'startTime', 'venue', 'city', 'faceTwd', 'platform', 'account', 'orderNumber', 'pickupDate', 'pickupMethod', 'pickedUp', 'ticketFeeTwd'],
  aliases: ['id', 'field', 'from', 'to', 'createdBy', 'createdAt'], // 習慣記法：讀到 from 一律記成 to
  // 搶票提醒：一波開賣一列；watchers／done 是 LINE 使用者 id 的 JSON 陣列
  onsales: ['id', 'groupId', 'title', 'artist', 'venue', 'city', 'showDates', 'phase', 'saleAt', 'platform', 'price', 'notes',
    'watchers', 'done', 'sentEve', 'sentLead', 'createdBy', 'createdAt'],
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
  if (TABLE_MEMO_[name]) return TABLE_MEMO_[name];
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
  TABLE_MEMO_[name] = rows;
  return rows;
}

function writeRows_(name, rows) {
  delete TABLE_MEMO_[name];
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
  delete TABLE_MEMO_[name];
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
  '- docType：已經買好的訂單或票券填 order；主辦單位的售票公告（還沒買，列出開賣時間）填 onsale。',
  '- 如果是 onsale：eventName、artist、venue、city 照上面規則填；showDates 列出所有演出場次（YYYY-MM-DD HH:MM，沒有時間就 YYYY-MM-DD）；',
  '  sales 每一波開賣一筆：phase（例如會員預售、全面開賣、抽選登記）、saleAt（YYYY-MM-DD HH:MM，24 小時制）、platform（同上面的平台寫法）；',
  '  priceInfo 是票價摘要（例如 5880/4880/3880）；notice 是注意事項（例如實名制、每人限購 4 張）。訂單相關欄位留空或 0。',
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
    docType: { type: 'STRING' }, showDates: { type: 'ARRAY', items: { type: 'STRING' } },
    sales: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
      phase: { type: 'STRING' }, saleAt: { type: 'STRING' }, platform: { type: 'STRING' } } } },
    priceInfo: { type: 'STRING' }, notice: { type: 'STRING' },
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
  const parts = [{ text: SCAN_PROMPT.replace('{TODAY}', today_()) + knownNamesText_() }];
  images.forEach(function (img) {
    parts.push({ inline_data: { mime_type: img.mimeType || 'image/jpeg', data: img.dataBase64 } });
  });
  const r = gemini_(parts, SCAN_SCHEMA);
  if (r.ok) {
    applyAliases_(r.data);
    r.duplicate = findOrder_(r.data.orderNumber);
  }
  return r;
}

const ALIAS_FIELDS = { venue: '場館', city: '城市', artist: '表演者', eventName: '活動名稱' };
const sheetRows_ = function (name) { return ss_().getSheetByName(name) ? readTable_(name) : []; };

// 把總帳裡用過的寫法交給 AI，讓它盡量沿用
function knownNamesText_() {
  const events = sheetRows_('events').slice().sort(function (a, b) { return String(b.startDate).localeCompare(String(a.startDate)); });
  const pick = function (key) {
    const out = [];
    events.forEach(function (e) { const v = String(e[key] || '').trim(); if (v && out.indexOf(v) < 0 && out.length < 40) out.push(v); });
    return out.join('、');
  };
  const venues = pick('venue'), artists = pick('artist');
  if (!venues && !artists) return '';
  return '\n- 以下是使用者習慣的寫法，讀到同一個場館或表演者時請沿用這些寫法：\n  場館：' + venues + '\n  表演者：' + artists;
}
function applyAliases_(data) {
  sheetRows_('aliases').forEach(function (a) {
    if (ALIAS_FIELDS[a.field] && data[a.field] && normName_(data[a.field]) === normName_(a.from)) data[a.field] = a.to;
  });
  return data;
}
function findOrder_(orderNumber) {
  const no = String(orderNumber || '').trim();
  if (!no) return null;
  const hit = sheetRows_('ledger').filter(function (l) { return l.ticketOrderNumber === no; })[0];
  if (hit) return { title: hit.title, date: hit.date };
  const tr = sheetRows_('transfers').filter(function (t) { return String(t.notes || '').indexOf('訂單：' + no) >= 0; })[0];
  return tr ? { title: tr.title + '（轉賣）', date: tr.date } : null;
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
  if (e.parameter.ics !== undefined) return calendarFeed_(e.parameter.ics); // Apple 行事曆訂閱
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
const LINE_HELP = [
  '📸 記票：傳購票截圖給我（同一筆訂單可一次傳 2～3 張）→ 回答幾個問題 → 確認卡片按「確認建檔」。隨時輸入「取消」可以放棄目前這筆。',
  '🎫 搶票：傳主辦單位的售票公告截圖給我，確認後會在開賣前一天晚上和開賣前 30 分鐘提醒要搶的人。',
  '🔎 查詢：用下方選單，或直接問我，例如「11月有什麼場」「DAY6 今年看了幾場」「今年花多少」。',
  '⏰ 提醒：可取票當天會提醒大家；打「提醒設定」看時間，打「取票提醒改成 9:30」「轉賣提醒改成週五 21:00」「轉賣提醒關掉」修改。',
  '🗓 行事曆：按選單的「行事曆」看 Apple 行事曆訂閱方式。',
  '📝 換售資訊：選售票或換票、挑要放的票，就會產生可以直接複製的文章；打「售票備註」「換票備註」看或修改固定備註。',
  '💰 花費：打「本月花費」或直接問「今年花多少」。',
].join('\n\n');

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
  if (me) me.uid = uid;
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
    const target = s.d || s.o ? 'pending' : 'imgs'; // 已經在問答中 → 當作補傳
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
  if (s.d || s.o) checkSupplement_(ev.replyToken, uid, s);
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
  if (r.data.docType === 'onsale') return startOnsale_(token, uid, s, r.data);
  if (!str_(r.data.eventName) && !str_(r.data.date) && !num_(r.data.totalPaid)) {
    clearLineSession_(uid);
    return lineReply_(token, '看不出這是購票截圖耶，換一張試試看 🙏');
  }
  s.d = draftFrom_(r.data);
  s.asked = {};
  s.edited = [];
  if (r.duplicate) {
    s.step = 'dupOrder';
    saveLineSession_(uid, s);
    return lineReply_(token, ['讀到了 ✦\n' + summaryText_(s.d), ask_('⚠ 這筆訂單（編號 ' + s.d.orderNumber + '）好像已經記過了：\n'
      + r.duplicate.title + (r.duplicate.date ? '（' + r.duplicate.date + '）' : ''), 'dupord', [['go', '還是要記'], ['cancel', '取消']])]);
  }
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
    [['d3', '演出前 3 天'], ['d5', '演出前 5 天'], ['d7', '演出前 7 天'], ['any', '隨時'], ['unknown', '不確定']]);
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
    if (/^d\d+$/.test(v)) d.pickupDate = minusDays_(d.date, Number(v.slice(1)));
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
  const offers = [];
  Object.keys(changes).forEach(function (k) {
    if (k === 'understood' || !(k in view)) return;
    const before = d[k];
    d[k] = typeof changes[k] === 'number' ? changes[k] : str_(changes[k]);
    changed.push(k);
    if (ALIAS_FIELDS[k] && str_(before) && d[k] && normName_(before) !== normName_(d[k])) offers.push({ f: k, from: str_(before), to: d[k] });
  });
  if (changed.some(function (k) { return ['currency', 'totalPaid'].indexOf(k) >= 0; }) && changed.indexOf('amountTwd') < 0) fillTwd_(d);
  if (changed.indexOf('amountTwd') >= 0) d.twdEstimated = false;
  return { changed: changed, offers: offers };
}

function onLineText_(token, uid, me, users, text) {
  let s = lineSession_(uid);
  if (/^(取消|cancel)$/i.test(text)) {
    if (tradeSession_(uid)) { clearTradeSession_(uid); return lineReply_(token, '已取消，沒有產生文章。'); }
    clearLineSession_(uid);
    return lineReply_(token, s && s.d ? '已取消這筆，沒有建檔。' : '目前沒有進行中的紀錄。');
  }
  if (/^(說明|help|\?|？)$/i.test(text)) return lineReply_(token, LINE_HELP);
  const cmd = lineCommand_(text, me);
  if (cmd) return lineReply_(token, cmd);
  const trade = tradeSession_(uid);
  if (trade) return onTradeText_(token, uid, trade, text);
  if (!s) {
    if (PROPS.getProperty('ld_' + uid)) {
      PROPS.deleteProperty('ld_' + uid);
      return lineReply_(token, '上一筆放太久，已經自動清除了，請重新傳截圖 🙏');
    }
    lineLoading_(uid);
    return lineReply_(token, answerQuestion_(text, me));
  }
  if (s.step === 'dupOrder') return lineReply_(token, ask_('這筆訂單好像已經記過了，要繼續記嗎？', 'dupord', [['go', '還是要記'], ['cancel', '取消']]));
  if (s.step === 'name') return setName_(token, uid, users, text);
  if (s.step === 'supplement') return lineReply_(token, ask_('請先選擇這張新截圖要怎麼處理：', 'supp', [['merge', '補充到這筆'], ['new', '當成新的一筆']]));
  if (s.step === 'oconfirm') {
    lineLoading_(uid);
    const r = editOnsale_(s.o, text);
    if (r.error) return lineReply_(token, r.error);
    if (!r.changed) return lineReply_(token, '看不太懂要改哪裡，可以說得更具體一點，例如「全面開賣改成 11/4 12:00 拓元」。');
    saveLineSession_(uid, s);
    return lineReply_(token, ['已修改 ✦', onsaleCard_(s.o)]);
  }
  if (s.step === 'confirm' || s.step === 'dupe') {
    lineLoading_(uid);
    const r = editDraft_(s.d, text);
    if (r.error) return lineReply_(token, r.error);
    if (!r.changed.length) return lineReply_(token, '看不太懂要改哪裡，可以說得更具體一點，例如「座位改成5排18號」「付款人改成 Jhen」。');
    s.edited = (s.edited || []).concat(r.changed);
    delete s.eventChoice;
    const out = ['已修改 ✦', nextMessage_(uid, s)];
    if (r.offers.length) {
      out.push(ask_('以後都這樣記嗎？\n' + r.offers.map(function (o) { return ALIAS_FIELDS[o.f] + '：' + o.from + ' → ' + o.to; }).join('\n'),
        'alias', [[JSON.stringify(r.offers), '好'], ['no', '只有這次']]));
    }
    return lineReply_(token, out);
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
  if (pb.a === 'picked') return markPicked_(token, pb.v);
  if (pb.a === 'alias') return saveAliases_(token, pb.v, me);
  if (pb.a === 'ogot' && pb.v.indexOf('miss:') === 0) return onsaleMark_(token, uid, 'omiss', pb.v.slice(5));
  if (['watch', 'unwatch', 'ogot', 'omiss'].indexOf(pb.a) >= 0) return onsaleMark_(token, uid, pb.a, pb.v);
  if (pb.a === 'onotify') return onsaleNotify_(token, uid, me, pb.v);
  if (pb.a === 'trade') return tradePick_(token, uid, pb.v === 'swap' ? 'swap' : 'sell');
  if (pb.a === 'tnote' || pb.a === 'tdeliv' || pb.a === 'tplace') {
    const tr = tradeSession_(uid);
    if (!tr) return lineReply_(token, '這篇已經結束或放太久了，請重新按「換售資訊」。');
    if (pb.a === 'tdeliv') return tradeSetDelivery_(token, uid, tr, pb.v === 'meet' ? 'meet' : 'code');
    if (pb.a === 'tplace') return tradeSetPlace_(token, uid, tr, pb.v);
    return tradeFinish_(token, uid, tr, '');
  }
  const s = lineSession_(uid);
  if (!s || (!s.d && !s.o)) return lineReply_(token, '這筆已經結束或放太久被清除了，請重新傳截圖 🙏');
  if (pb.a === 'ocard') {
    if (pb.v === 'cancel') { clearLineSession_(uid); return lineReply_(token, '已取消，沒有建立搶票提醒。'); }
    if (pb.v === 'edit') return lineReply_(token, '直接打字告訴我要改什麼就好，例如：\n「全面開賣改成 11/4 12:00 拓元」\n「場館改成 TICC」');
    return confirmOnsale_(token, uid, me, s);
  }
  if (LINE_STEPS.indexOf(pb.a) >= 0) {
    if (!applyAnswer_(s, pb.a, pb.v, false)) return lineReply_(token, question_(pb.a, s.d));
    return lineReply_(token, nextMessage_(uid, s));
  }
  if (pb.a === 'dupord') {
    if (pb.v !== 'go') {
      clearLineSession_(uid);
      return lineReply_(token, '已取消這筆，沒有建檔。');
    }
    return lineReply_(token, nextMessage_(uid, s));
  }
  if (pb.a === 'supp') {
    if (pb.v === 'merge') return s.o ? mergeOnsale_(token, uid, s, s.pendingData) : mergeSupplement_(token, uid, s, s.pendingData);
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
  if (s.o && x.docType === 'onsale') return mergeOnsale_(token, uid, s, x);
  const same = !s.o && (str_(x.orderNumber) && str_(x.orderNumber) === s.d.orderNumber)
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
  const got = markOnsaleGot_(uid, d);
  lineReply_(token, (result.kind === 'event'
    ? '已建檔 ✦ ' + result.title + '\n可以到網站「活動」頁查看。'
    : '已建檔 ✦ 記到「讓票」頁了：' + result.title)
    + (got.length ? '\n\n也把搶票提醒「' + got.join('、') + '」標成已搶到，不會再提醒你。' : ''));
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
        artist: d.artist, eventName: d.eventName, startTime: d.time, venue: d.venue, city: d.city,
        faceTwd: foreign ? '' : num_(d.unitFace) || '', platform: d.platform, account: d.account, orderNumber: d.orderNumber,
        pickupDate: d.pickupDate, pickupMethod: d.pickupMethod, pickedUp: false, ticketFeeTwd: foreign ? '' : num_(d.unitFee) || '',
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

/* ---------- 習慣記法、已取票 ---------- */
function saveAliases_(token, v, me) {
  if (v === 'no') return lineReply_(token, '好，只有這次。');
  let offers = [];
  try { offers = JSON.parse(v); } catch (err) { offers = []; }
  if (!offers.length) return lineReply_(token, '沒有要記住的寫法。');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setup_();
    const rows = readTable_('aliases');
    const out = offers.map(function (o) {
      const old = rows.filter(function (a) { return a.field === o.f && normName_(a.from) === normName_(o.from); })[0];
      return { id: old ? old.id : newId_(), field: o.f, from: o.from, to: o.to, createdBy: me.name || '', createdAt: today_() };
    });
    writeRows_('aliases', out);
  } finally {
    lock.releaseLock();
  }
  lineReply_(token, '記住了 ✦ 以後會自動換成：\n' + offers.map(function (o) { return o.from + ' → ' + o.to; }).join('\n')
    + '\n（對照表在試算表「aliases」分頁，可以直接修改或刪除）');
}

function markPicked_(token, id) {
  const resale = id.indexOf('t:') === 0;
  const table = resale ? 'transfers' : 'ledger';
  if (resale) id = id.slice(2);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let row;
  try {
    setup_();
    row = readTable_(table).filter(function (l) { return l.id === id; })[0];
    if (row) {
      row[resale ? 'pickedUp' : 'ticketPickedUp'] = 'true';
      writeRows_(table, [row]);
    }
  } finally {
    lock.releaseLock();
  }
  lineReply_(token, row ? '已標記「' + row.title + '」已取票 ✦' : '找不到這張票，可能已經被刪除了。');
}

/* ---------- 查詢 ---------- */
const CAT_LABEL_ = { ticket: '票券', transport: '交通', lodging: '住宿', food: '餐飲', merch: '周邊', other: '其他' };
const WEEK_ = '日一二三四五六';
const md_ = function (date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return date || '日期未定';
  return Number(date.slice(5, 7)) + '/' + Number(date.slice(8, 10)) + '（' + WEEK_[new Date(date + 'T00:00:00Z').getUTCDay()] + '）';
};
const eventTitle_ = function (e) {
  return (e.eventType === '拼盤' || !e.artist || String(e.name).indexOf(e.artist) >= 0) ? e.name : e.artist + ' - ' + e.name;
};
const isPicked_ = function (l) { return String(l.ticketPickedUp) === 'true'; };
function matchEvent_(e, q) {
  const hay = normName_([e.name, e.artist, e.venue, e.liveTour, e.city].join(' '));
  const keys = [q.artist, q.keyword].map(normName_).filter(Boolean);
  return keys.every(function (k) { return hay.indexOf(k) >= 0; });
}
function inPeriod_(date, q) {
  date = String(date || '');
  if (q.year && date.slice(0, 4) !== String(q.year)) return false;
  if (q.month && Number(date.slice(5, 7)) !== Number(q.month)) return false;
  return true;
}
function pendingTickets_(ledger, events, today) {
  const byId = {};
  events.forEach(function (e) { byId[e.id] = e; });
  return ledger.filter(function (l) {
    if (l.category !== 'ticket' || isPicked_(l)) return false;
    const ev = byId[l.eventId];
    return String(ev ? ev.startDate || '9999' : l.date || '9999') >= today;
  }).map(function (l) { return { l: l, ev: byId[l.eventId] }; });
}

/* ---------- 清單卡片（手機上好讀的排版） ---------- */
// 每筆左邊是日期欄，右邊是內容；一張卡片放 6 筆，多的往右滑
const C_INK = '#40382E', C_MUTED = '#9C917F', C_ACCENT = '#A67B5B', C_OK = '#7C9A6D', C_LINE = '#E7DECF';
const ftext_ = function (text, size, color, extra) {
  return Object.assign({ type: 'text', text: String(text), size: size || 'sm', color: color || C_INK, wrap: true }, extra || {});
};
// e：{ date, time, top, title, lines: [], status, statusColor, section }
function agendaRow_(e) {
  if (e.section) return ftext_(e.section, 'sm', C_ACCENT, { weight: 'bold', margin: 'lg' });
  const ok = /^\d{4}-\d{2}-\d{2}$/.test(e.date || '');
  const left = { type: 'box', layout: 'vertical', width: '58px', flex: 0, contents: ok ? [
    ftext_(e.date.slice(0, 4), 'xxs', C_MUTED),
    ftext_(e.date.slice(5, 7) + '/' + e.date.slice(8, 10), 'md', C_INK, { weight: 'bold' }),
    ftext_('週' + WEEK_[new Date(e.date + 'T00:00:00Z').getUTCDay()] + (e.time ? '\n' + e.time : ''), 'xs', C_MUTED),
  ] : [ftext_('日期\n未定', 'xs', C_MUTED)] };
  const right = { type: 'box', layout: 'vertical', flex: 1, spacing: 'xs', contents: [
    e.top ? ftext_(e.top, 'xs', C_ACCENT, { weight: 'bold' }) : null,
    ftext_(e.title || '—', 'sm', C_INK, { weight: 'bold', maxLines: 3 }),
  ].concat((e.lines || []).filter(Boolean).map(function (x) { return ftext_(x, 'xs', C_MUTED); }))
    .concat(e.status ? [ftext_(e.status, 'xs', e.statusColor || C_ACCENT, { weight: 'bold' })] : []).filter(Boolean) };
  return { type: 'box', layout: 'horizontal', spacing: 'md', margin: 'lg', contents: [left, right] };
}
function agendaMessage_(title, entries, opts) {
  opts = opts || {};
  const per = 6;
  const pages = [];
  let cur = [];
  entries.forEach(function (e) {
    if (cur.filter(function (x) { return !x.section; }).length >= per && !e.section) { pages.push(cur); cur = []; }
    cur.push(e);
  });
  if (cur.length) pages.push(cur);
  const bubbles = pages.slice(0, 12).map(function (pg, pi) {
    const body = [ftext_(title, 'md', C_INK, { weight: 'bold' })];
    if (opts.sub) body.push(ftext_(opts.sub, 'xs', C_MUTED));
    pg.forEach(function (e, i) {
      if (i > 0 && !e.section && !pg[i - 1].section) body.push({ type: 'separator', margin: 'lg', color: C_LINE });
      body.push(agendaRow_(e));
    });
    if (pages.length > 1) body.push(ftext_((pi + 1) + ' / ' + Math.min(pages.length, 12) + (pi < pages.length - 1 ? '　往左滑看更多 →' : ''), 'xxs', C_MUTED, { margin: 'xl', align: 'end' }));
    if (opts.foot && pi === pages.length - 1) body.push(ftext_(opts.foot, 'xxs', C_MUTED, { margin: 'xl' }));
    return { type: 'bubble', size: 'giga', body: { type: 'box', layout: 'vertical', spacing: 'none', contents: body } };
  });
  const msg = { type: 'flex', altText: title, contents: bubbles.length === 1 ? bubbles[0] : { type: 'carousel', contents: bubbles } };
  if (opts.quick && opts.quick.length) msg.quickReply = ask_('', opts.step, opts.quick).quickReply;
  return msg;
}
const placeText_ = function (it) { return [it.city, it.venue].filter(Boolean).join(' '); };
const pre_ = function (icon, v) { return str_(v) ? icon + ' ' + str_(v) : ''; };

function upcomingText_(q) {
  const today = today_();
  const list = sheetRows_('events').filter(function (e) { return String(e.startDate) >= today && matchEvent_(e, q) && inPeriod_(e.startDate, q); })
    .sort(function (a, b) { return String(a.startDate + a.startTime).localeCompare(String(b.startDate + b.startTime)); });
  if (!list.length) return '沒有符合的未來場次。';
  return agendaMessage_('接下來的場次', list.slice(0, 72).map(function (e) {
    return { date: e.startDate, time: e.startTime, top: e.artist, title: e.name, lines: [pre_('📍', placeText_(e))] };
  }), { sub: '共 ' + list.length + ' 場' });
}

// 自己的票（總帳）與轉賣的票（讓票頁）整理成同一種格式；一筆紀錄＝一張訂單
const dateW_ = function (date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return date || '日期未定';
  return date.replace(/-/g, '/') + '(' + WEEK_[new Date(date + 'T00:00:00Z').getUTCDay()] + ')';
};
const oneLine_ = function (v) { return String(v || '').replace(/\s*\n\s*/g, ' ').trim(); };
const noteField_ = function (notes, re) { const m = String(notes || '').match(re); return m ? m[1].trim() : ''; };
function ownItem_(l, ev) {
  const count = Math.max(1, num_(l.ticketCount) || 1);
  return {
    kind: 'own', id: l.id, date: ev ? ev.startDate : l.date, time: ev ? ev.startTime : '', artist: ev ? ev.artist : '',
    name: ev ? ev.name : String(l.title || '').replace(/^票券 - /, ''), city: ev ? ev.city : '', venue: ev ? ev.venue : '',
    area: oneLine_(l.ticketArea), row: oneLine_(l.ticketRow), seat: oneLine_(seatText_(l)), count: count,
    face: num_(l.ticketFaceTwd) || Math.round(num_(l.amountTwd) / count), fee: num_(l.ticketFaceTwd) ? num_(l.ticketFeeTwd) : 0,
    pickupDate: l.ticketPickupDate,
    pickupMethod: l.ticketPickupMethod, platform: l.ticketPlatform, account: l.ticketAccount, picked: isPicked_(l),
  };
}
function transferItem_(t, ev) {
  const count = Math.max(1, num_(t.ticketCount) || 1);
  const split = String(t.title || '').split(' - ');
  return {
    kind: 'resale', id: t.id, date: t.date || (ev ? ev.startDate : ''),
    time: t.startTime || (ev ? ev.startTime : '') || noteField_(t.notes, /開演 (\d{1,2}:\d{2})/),
    artist: t.artist || (ev ? ev.artist : '') || (split.length > 1 ? split[0] : ''),
    name: t.eventName || (ev ? ev.name : '') || (split.length > 1 ? split.slice(1).join(' - ') : t.title),
    city: t.city || (ev ? ev.city : ''), venue: t.venue || (ev ? ev.venue : ''),
    area: oneLine_(t.ticketArea), row: oneLine_(t.ticketRow), seat: oneLine_(seatText_(t)), count: count,
    face: num_(t.faceTwd) || Math.round(num_(t.costTwd) / count), fee: num_(t.faceTwd) ? num_(t.ticketFeeTwd) : 0,
    pickupDate: t.pickupDate || noteField_(t.notes, /領票日 (\d{4}-\d{2}-\d{2})/),
    pickupMethod: t.pickupMethod || noteField_(t.notes, /取票：([^｜]+)/),
    platform: t.platform || noteField_(t.notes, /平台：([^｜]+)/), account: t.account || noteField_(t.notes, /帳號：([^｜]+)/),
    picked: String(t.pickedUp) === 'true', sold: String(t.settled) === 'true' || num_(t.amountTwd) > 0,
  };
}
// 每筆：YYYY/MM/DD(星期) 時間／演出者｜活動名稱／地點／座位 ×張數
function itemBlock_(it) {
  return [
    dateW_(it.date) + (it.time ? ' ' + it.time : ''),
    [it.artist, it.name].filter(Boolean).join('｜'),
    [it.city, it.venue].filter(Boolean).join(' '),
    [it.seat, it.count > 1 ? '×' + it.count + ' 張' : ''].filter(Boolean).join(' '),
  ].filter(Boolean).join('\n');
}
function pickupStatus_(it, today) {
  const p = it.pickupDate;
  return [!p || p <= today ? '✅ 現在可取' : '⏳ ' + dateW_(p) + ' 起可取', it.pickupMethod, [it.platform, it.account].filter(Boolean).join(' ')].filter(Boolean).join('｜');
}
function pendingPickups_(today) {
  const events = sheetRows_('events');
  const byId = {};
  events.forEach(function (e) { byId[e.id] = e; });
  const own = pendingTickets_(sheetRows_('ledger'), events, today).map(function (x) { return ownItem_(x.l, x.ev); });
  const resale = sheetRows_('transfers').filter(function (t) { return t.kind !== '退票'; })
    .map(function (t) { return transferItem_(t, byId[t.eventId]); })
    .filter(function (it) { return !it.picked && String(it.date || '9999') >= today; });
  const byPickup = function (a, b) { return String(a.pickupDate || '0000').localeCompare(String(b.pickupDate || '0000')) || String(a.date).localeCompare(String(b.date)); };
  return { own: own.sort(byPickup), resale: resale.sort(byPickup) };
}
function pickupListMessage_() {
  const today = today_();
  const p = pendingPickups_(today);
  if (!p.own.length && !p.resale.length) return '目前沒有待取的票 ✦';
  const entry = function (it) {
    const ready = !it.pickupDate || it.pickupDate <= today;
    return {
      date: it.date, time: it.time, top: it.artist, title: it.name,
      lines: [pre_('📍', placeText_(it)), pre_('🎫', [it.seat, it.count > 1 ? '×' + it.count + ' 張' : ''].filter(Boolean).join('　')),
        [it.pickupMethod, [it.platform, it.account].filter(Boolean).join(' ')].filter(Boolean).join('｜')],
      status: ready ? '✅ 現在可取' : '⏳ ' + dateW_(it.pickupDate) + ' 起可取', statusColor: ready ? C_OK : C_ACCENT,
    };
  };
  const entries = (p.own.length ? [{ section: '🎫 自己的票（' + p.own.length + '）' }].concat(p.own.map(entry)) : [])
    .concat(p.resale.length ? [{ section: '🔁 轉賣的票（' + p.resale.length + '）' }].concat(p.resale.map(entry)) : []);
  // 按鈕最多 13 個：演出日期近的排前面，標籤用「日期＋演出者」比較短
  const buttons = p.own.concat(p.resale).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }).slice(0, 13).map(function (it) {
    return [(it.kind === 'resale' ? 't:' : '') + it.id, (it.kind === 'resale' ? '已取(轉) ' : '已取 ') + shortMD_(it.date) + ' ' + (it.artist || it.name)];
  });
  return agendaMessage_('待取票', entries, { foot: '取完票可以點下方按鈕標記。', step: 'picked', quick: buttons });
}

function spendText_(q, me) {
  const today = today_();
  if (!q.year && !q.month) { q.year = Number(today.slice(0, 4)); q.month = Number(today.slice(5, 7)); }
  const events = sheetRows_('events');
  const byId = {};
  events.forEach(function (e) { byId[e.id] = e; });
  const rows = sheetRows_('ledger').filter(function (l) {
    if (l.type && l.type !== 'expense') return false;
    if (!inPeriod_(l.date, q)) return false;
    if (q.artist || q.keyword) { const ev = byId[l.eventId]; return ev ? matchEvent_(ev, q) : matchEvent_({ name: l.title }, q); }
    return true;
  });
  const label = (q.year ? q.year + ' 年' : '') + (q.month ? ' ' + q.month + ' 月' : '') + (q.artist || q.keyword ? '（' + (q.artist || q.keyword) + '）' : '');
  if (!rows.length) return label + '沒有花費紀錄。';
  const total = rows.reduce(function (t, l) { return t + num_(l.amountTwd); }, 0);
  const mine = me.name ? rows.reduce(function (t, l) { return t + ownShare_(l, me.name); }, 0) : 0;
  const byCat = {};
  rows.forEach(function (l) { const c = CAT_LABEL_[l.category] || '其他'; byCat[c] = (byCat[c] || 0) + num_(l.amountTwd); });
  return label + '花費\n\n總額 TWD ' + total.toLocaleString('en-US') + '（' + rows.length + ' 筆）'
    + (me.name ? '\n其中 ' + me.name + ' 的份 TWD ' + mine.toLocaleString('en-US') : '') + '\n\n'
    + Object.keys(byCat).map(function (c) { return c + '　TWD ' + byCat[c].toLocaleString('en-US'); }).join('\n');
}

function countText_(q) {
  const today = today_();
  const events = sheetRows_('events').filter(function (e) { return matchEvent_(e, q) && inPeriod_(e.startDate, q); });
  const seen = events.filter(function (e) { return e.startDate && e.startDate <= today; }).sort(function (a, b) { return String(b.startDate).localeCompare(String(a.startDate)); });
  const future = events.filter(function (e) { return e.startDate > today; });
  const label = (q.artist || q.keyword || '全部') + (q.year ? '（' + q.year + ' 年）' : '');
  const sum = label + '：已經看過 ' + seen.length + ' 場' + (future.length ? '，還有 ' + future.length + ' 場在後面' : '');
  if (!seen.length) return sum;
  return agendaMessage_(sum, seen.slice(0, 12).map(function (e) {
    return { date: e.startDate, time: e.startTime, top: e.artist, title: e.name, lines: [pre_('📍', placeText_(e)), pre_('🎫', oneLine_(e.seat))] };
  }), { sub: '最近看過的' + Math.min(12, seen.length) + ' 場' });
}

function unsoldTransfers_() {
  const today = today_();
  return sheetRows_('transfers').filter(function (t) {
    return String(t.settled) !== 'true' && !num_(t.amountTwd) && String(t.date || '9999') >= today && ['退票', '換票'].indexOf(t.kind) < 0;
  }).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
}
function resaleText_(title) {
  const list = unsoldTransfers_();
  if (!list.length) return '目前沒有還沒賣出的票 ✦';
  const byId = {};
  sheetRows_('events').forEach(function (e) { byId[e.id] = e; });
  return agendaMessage_(title || '還沒賣出的票', list.map(function (t) {
    const it = transferItem_(t, byId[t.eventId]);
    return { date: it.date, time: it.time, top: it.artist, title: it.name,
      lines: [pre_('📍', placeText_(it)), pre_('🎫', [it.seat, '×' + it.count + ' 張'].filter(Boolean).join('　')), it.face ? '💰 票面 ' + it.face + (it.fee ? '＋系統服務費 ' + it.fee : '') : ''] };
  }), { sub: '共 ' + list.length + ' 筆', foot: '賣出後到網站「讓票」頁填上對方付的金額，就不會再出現在這裡。' });
}

function calendarText_() {
  let base = PROPS.getProperty('WEBAPP_URL') || '';
  try { base = ScriptApp.getService().getUrl() || base; } catch (err) { /* 未授權時用 setupBot 存下的網址 */ }
  if (!base) return '行事曆網址還沒設定好，請先在 Apps Script 執行一次 setupBot。';
  const url = base + '?ics=' + encodeURIComponent(SHARED_KEY);
  return '🗓 Apple 行事曆訂閱\n\n網址：\n' + url + '\n\niPhone：設定 → App → 行事曆 → 行事曆帳號 → 加入帳號 → 其他 → 加入訂閱的行事曆 → 貼上網址 → 下一步 → 儲存。\n\n'
    + '之後新增的場次會自動出現（Apple 大約每小時更新一次）。網址裡有共用密碼，請不要公開分享。';
}

const QUERY_SCHEMA = {
  type: 'OBJECT',
  properties: {
    intent: { type: 'STRING' }, artist: { type: 'STRING' }, keyword: { type: 'STRING' },
    year: { type: 'NUMBER' }, month: { type: 'NUMBER' },
  },
  required: ['intent'],
};
function answerQuestion_(text, me) {
  const prompt = [
    '今天是 ' + today_() + '。使用者在追星記帳機器人問了一句話，請判斷要查什麼，只輸出 JSON：',
    '「' + text + '」',
    'intent 只能是：upcoming（未來場次）、pickup（待取票）、spend（花費）、count（看過幾場）、resale（還沒賣出的轉賣票）、onsale（即將開賣、要搶票的節目）、unknown（其他或看不懂）。',
    'artist 是提到的表演者，keyword 是提到的其他關鍵字（例如場館、活動名稱），沒有就填空字串。',
    'year、month 是提到的年份與月份（例如「11月」→ month 11，「今年」→ 今年的年份，「上個月」→ 換算成對應年月），沒提到填 0。',
  ].join('\n');
  const r = gemini_([{ text: prompt }], QUERY_SCHEMA);
  const q = r.ok ? r.data : { intent: 'unknown' };
  q.artist = str_(q.artist);
  q.keyword = str_(q.keyword);
  q.year = num_(q.year);
  q.month = num_(q.month);
  if (q.intent === 'upcoming') return upcomingText_(q);
  if (q.intent === 'pickup') return pickupListMessage_();
  if (q.intent === 'spend') return spendText_(q, me);
  if (q.intent === 'count') return countText_(q);
  if (q.intent === 'resale') return resaleText_();
  if (q.intent === 'onsale') return onsaleListMessage_(me.uid, q);
  return '我可以幫你查：未來場次、待取票、即將開賣、花費、看過幾場、還沒賣出的票，例如「11月有什麼場」「下週有什麼要搶」。\n要記新的票或搶票公告，直接傳截圖給我 ✦';
}

function lineCommand_(text, me) {
  const t = text.replace(/\s/g, '');
  if (t === '未來場次') return upcomingText_({});
  if (t === '待取票') return pickupListMessage_();
  if (t === '本月花費') return spendText_({}, me);
  if (t === '轉賣中') return resaleText_();
  if (t === '即將開賣') return onsaleListMessage_(me.uid, {});
  if (t === '換售資訊') return tradeStart_();
  if (/^(售票|換票)備註$/.test(t)) return tradeNotesText_(t.indexOf('售') === 0 ? 'sell' : 'swap');
  if (/^(售票|換票)備註(改成|改為)?[：:]/.test(t)) return setTradeNotes_(t.indexOf('售') === 0 ? 'sell' : 'swap', text);
  if (t === '行事曆') return calendarText_();
  if (t === '提醒設定') return settingsText_();
  if (t === '測試提醒') return testReminder_();
  if (/取票提醒/.test(t)) return setReminder_('pickup', t);
  if (/搶票預告/.test(t)) return setReminder_('onsaleEve', t);
  if (/搶票提醒/.test(t)) return setReminder_('onsaleLead', t);
  if (/轉賣(提醒|追蹤)/.test(t)) return setReminder_('resale', t);
  return null;
}

/* ---------- 提醒 ---------- */
// REMIND_PICKUP：'11:50' 或 'off'；REMIND_RESALE：'4 20:00'（星期幾 1=一…7=日）或 'off'
// REMIND_ONSALE_EVE：開賣前一天幾點預告（'21:00' 或 'off'）；REMIND_ONSALE_LEAD：開賣前幾分鐘（'30' 或 'off'）
function reminderSettings_() {
  return {
    pickup: PROPS.getProperty('REMIND_PICKUP') || '11:50', resale: PROPS.getProperty('REMIND_RESALE') || '4 20:00',
    onsaleEve: PROPS.getProperty('REMIND_ONSALE_EVE') || '21:00', onsaleLead: PROPS.getProperty('REMIND_ONSALE_LEAD') || '30',
  };
}
function settingsText_() {
  const c = reminderSettings_();
  const rs = c.resale === 'off' ? '關閉' : '每週' + '一二三四五六日'[Number(c.resale.split(' ')[0]) - 1] + ' ' + c.resale.split(' ')[1];
  return '⏰ 提醒設定（所有人共用）\n\n取票提醒：' + (c.pickup === 'off' ? '關閉' : '可取票當天 ' + c.pickup)
    + '\n轉賣提醒：' + rs
    + '\n搶票預告：' + (c.onsaleEve === 'off' ? '關閉' : '開賣前一天 ' + c.onsaleEve)
    + '\n搶票提醒：' + (c.onsaleLead === 'off' ? '關閉' : '開賣前 ' + c.onsaleLead + ' 分鐘')
    + '\n\n修改方式：打「取票提醒改成 9:30」「轉賣提醒改成週五 21:00」「搶票預告改成 20:00」「搶票提醒改成前 15 分鐘」，或「…關掉」。\n實際送出時間會比設定晚 0～5 分鐘。';
}
function parseTime_(t) {
  const m = t.match(/(\d{1,2})[:：點時](\d{1,2})?(半)?/) || t.match(/(\d{1,2})()()$/);
  if (!m) return '';
  let h = Number(m[1]);
  const min = m[3] ? 30 : Number(m[2] || 0);
  if (/(晚上|下午|傍晚)/.test(t) && h < 12) h += 12;
  if (h > 23 || min > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + min).slice(-2);
}
function setReminder_(kind, t) {
  const key = { pickup: 'REMIND_PICKUP', resale: 'REMIND_RESALE', onsaleEve: 'REMIND_ONSALE_EVE', onsaleLead: 'REMIND_ONSALE_LEAD' }[kind];
  const name = { pickup: '取票提醒', resale: '轉賣提醒', onsaleEve: '搶票預告', onsaleLead: '搶票提醒' }[kind];
  if (/(關|停|取消|不要)/.test(t)) {
    PROPS.setProperty(key, 'off');
    return '已關閉' + name + '。\n\n' + settingsText_();
  }
  if (kind === 'onsaleLead') {
    const m = t.match(/(\d+)\s*分/) || t.match(/(\d+)$/);
    if (!m || Number(m[1]) < 5 || Number(m[1]) > 720) return '看不懂，可以這樣說：「搶票提醒改成前 15 分鐘」（5～720 分鐘）。';
    PROPS.setProperty(key, String(Number(m[1])));
    return '已修改 ✦\n\n' + settingsText_();
  }
  const time = parseTime_(t.replace(/.*(提醒|追蹤|預告)/, ''));
  if (kind === 'onsaleEve') {
    if (!time) return '看不懂時間，可以這樣說：「搶票預告改成 20:00」。';
    PROPS.setProperty(key, time);
  } else if (kind === 'pickup') {
    if (!time) return '看不懂時間，可以這樣說：「取票提醒改成 9:30」。';
    PROPS.setProperty(key, time);
  } else {
    const cur = reminderSettings_().resale;
    const wm = t.match(/(週|星期|禮拜)([一二三四五六日天])/);
    const day = wm ? '一二三四五六日天'.indexOf(wm[2]) % 7 + 1 : Number(cur === 'off' ? 4 : cur.split(' ')[0]);
    const at = time || (cur === 'off' ? '20:00' : cur.split(' ')[1]);
    if (!wm && !time) return '看不懂時間，可以這樣說：「轉賣提醒改成週五 21:00」。';
    PROPS.setProperty(key, day + ' ' + at);
  }
  return '已修改 ✦\n\n' + settingsText_();
}

function pickupBubbles_(items) {
  return items.slice(0, 12).map(function (it) {
    const lines = itemBlock_(it).split('\n').slice(2).concat(['取票：' + (it.pickupMethod || '未填'), [it.platform, it.account].filter(Boolean).join(' ')]).filter(Boolean);
    return {
      type: 'bubble', size: 'kilo',
      body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
        { type: 'text', text: it.kind === 'resale' ? '今天可以取票（轉賣）' : '今天可以取票', size: 'xs', color: '#A67B5B', weight: 'bold' },
        { type: 'text', text: dateW_(it.date) + (it.time ? ' ' + it.time : ''), size: 'sm', color: '#40382E' },
        { type: 'text', text: [it.artist, it.name].filter(Boolean).join('｜') || '票券', weight: 'bold', wrap: true, color: '#40382E' },
      ].concat(lines.map(function (x) { return { type: 'text', text: x, size: 'sm', color: '#9C917F', wrap: true }; })) },
      footer: { type: 'box', layout: 'vertical', contents: [
        { type: 'button', style: 'primary', color: '#A67B5B', height: 'sm',
          action: { type: 'postback', label: '已取票', data: 'a=picked&v=' + (it.kind === 'resale' ? 't:' : '') + it.id, displayText: '已取票' } },
      ] },
    };
  });
}
function pickupReminderMessage_(today) {
  const p = pendingPickups_(today);
  const items = p.own.concat(p.resale).filter(function (it) { return it.pickupDate === today; });
  if (!items.length) return null;
  return { type: 'flex', altText: '今天可以取票：' + items.map(function (it) { return it.name; }).join('、'),
    contents: { type: 'carousel', contents: pickupBubbles_(items) } };
}
function resaleReminderMessage_() {
  return unsoldTransfers_().length ? resaleText_('🔔 每週轉賣提醒：還沒賣出的票') : null;
}

function testReminder_() {
  const today = today_();
  const p = pendingPickups_(today);
  const soon = p.own.concat(p.resale).filter(function (it) { return it.pickupDate; })
    .sort(function (a, b) { return String(a.pickupDate).localeCompare(String(b.pickupDate)); }).slice(0, 3);
  const out = ['這是提醒的範例（只傳給你）：'];
  out.push(soon.length ? { type: 'flex', altText: '取票提醒範例', contents: { type: 'carousel', contents: pickupBubbles_(soon) } } : '（目前沒有設定取票日的票，取票提醒會長得像一張卡片，附「已取票」按鈕）');
  out.push(resaleReminderMessage_() || '（目前沒有未賣出的轉賣票，週提醒不會傳）');
  return out;
}

// 主動推播給所有加入的人
function linePushAll_(messages) {
  return linePush_(Object.keys(lineUsers_()), messages);
}

// 排程每 5 分鐘執行一次（由 setupBot 建立），到了設定時間才真的傳
function reminderTick() {
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm u').split(' ');
  const today = now[0], hm = now[1], weekday = now[2];
  onsaleTick_(today + ' ' + hm);
  const c = reminderSettings_();
  if (c.pickup !== 'off' && hm >= c.pickup && PROPS.getProperty('SENT_PICKUP') !== today) {
    PROPS.setProperty('SENT_PICKUP', today);
    const m = pickupReminderMessage_(today);
    if (m) linePushAll_([m]);
  }
  if (c.resale !== 'off') {
    const r = c.resale.split(' ');
    if (weekday === r[0] && hm >= r[1] && PROPS.getProperty('SENT_RESALE') !== today) {
      PROPS.setProperty('SENT_RESALE', today);
      const m = resaleReminderMessage_();
      if (m) linePushAll_([m]);
    }
  }
}

/* ---------- Apple 行事曆訂閱 ---------- */
function calendarFeed_(key) {
  if (SHARED_KEY && key !== SHARED_KEY) return ContentService.createTextOutput('bad key');
  const esc = function (v) { return String(v || '').replace(/\\/g, '\\\\').replace(/;/g, '\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); };
  const fold = function (line) { // 每行最多 75 字元，超過換行接續
    const out = [];
    while (line.length > 70) { out.push(line.slice(0, 70)); line = ' ' + line.slice(70); }
    out.push(line);
    return out.join('\r\n');
  };
  const ledger = sheetRows_('ledger');
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//star-ledger//ticket//ZH', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'X-WR-CALNAME:追星場次', 'X-WR-TIMEZONE:Asia/Taipei', 'REFRESH-INTERVAL;VALUE=DURATION:PT1H',
    'BEGIN:VTIMEZONE', 'TZID:Asia/Taipei', 'BEGIN:STANDARD', 'DTSTART:19700101T000000', 'TZOFFSETFROM:+0800', 'TZOFFSETTO:+0800', 'TZNAME:CST', 'END:STANDARD', 'END:VTIMEZONE'];
  sheetRows_('events').forEach(function (e) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(e.startDate)) return;
    const d = e.startDate.replace(/-/g, '');
    const tickets = ledger.filter(function (l) { return l.eventId === e.id && l.category === 'ticket'; });
    const desc = [e.seat && '座位：' + e.seat, e.eventType].concat(tickets.map(function (l) {
      return [l.ticketPickupMethod && '取票：' + l.ticketPickupMethod, l.ticketPickupDate && '取票日 ' + l.ticketPickupDate, isPicked_(l) ? '已取票' : ''].filter(Boolean).join('｜');
    })).filter(Boolean).join('\n');
    lines.push('BEGIN:VEVENT', 'UID:' + e.id + '@star-ledger', 'DTSTAMP:' + stamp);
    if (/^\d{1,2}:\d{2}$/.test(e.startTime)) {
      const hm = e.startTime.split(':');
      const start = d + 'T' + ('0' + hm[0]).slice(-2) + hm[1] + '00';
      const endH = Math.min(23, Number(hm[0]) + 3);
      lines.push('DTSTART;TZID=Asia/Taipei:' + start, 'DTEND;TZID=Asia/Taipei:' + d + 'T' + ('0' + endH).slice(-2) + hm[1] + '00');
    } else {
      const next = new Date(e.startDate + 'T00:00:00Z');
      next.setUTCDate(next.getUTCDate() + 1);
      lines.push('DTSTART;VALUE=DATE:' + d, 'DTEND;VALUE=DATE:' + next.toISOString().slice(0, 10).replace(/-/g, ''));
    }
    lines.push(fold('SUMMARY:' + esc(eventTitle_(e))), fold('LOCATION:' + esc([e.city, e.venue].filter(Boolean).join(' '))));
    if (desc) lines.push(fold('DESCRIPTION:' + esc(desc)));
    lines.push('END:VEVENT');
  });
  sheetRows_('onsales').forEach(function (r) {
    if (!hasTime_(r.saleAt)) return;
    const st = r.saleAt.replace(/[-:]/g, '').replace(' ', 'T') + '00';
    const en = shiftTime_(r.saleAt, 30).replace(/[-:]/g, '').replace(' ', 'T') + '00';
    lines.push('BEGIN:VEVENT', 'UID:' + r.id + '@star-ledger-onsale', 'DTSTAMP:' + stamp,
      'DTSTART;TZID=Asia/Taipei:' + st, 'DTEND;TZID=Asia/Taipei:' + en,
      fold('SUMMARY:' + esc('🎫 開賣：' + r.title + '｜' + r.phase)),
      fold('DESCRIPTION:' + esc([r.platform && '平台：' + r.platform, r.price && '票價：' + r.price, r.notes].filter(Boolean).join('\n'))),
      'END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return ContentService.createTextOutput(lines.join('\r\n')).setMimeType(ContentService.MimeType.ICAL);
}

/* ---------- 一次性設定：在編輯器選 setupBot →「執行」 ---------- */
const RICHMENU_IMAGE = 'https://snowmanyo.github.io/star-ledger/line-richmenu.png';
function setupBot() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'reminderTick') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reminderTick').timeBased().everyMinutes(5).create();
  PROPS.setProperty('WEBAPP_URL', ScriptApp.getService().getUrl() || '');
  Logger.log('已建立提醒排程（每 5 分鐘檢查一次）');
  Logger.log(setupRichMenu_());
}
function setupRichMenu_() {
  const head = { Authorization: 'Bearer ' + PROPS.getProperty('LINE_TOKEN') };
  const old = PROPS.getProperty('RICHMENU_ID');
  if (old) UrlFetchApp.fetch('https://api.line.me/v2/bot/richmenu/' + old, { method: 'delete', headers: head, muteHttpExceptions: true });
  const labels = ['未來場次', '待取票', '即將開賣', '換售資訊', '轉賣中', '行事曆', '提醒設定', '說明'];
  const menu = {
    size: { width: 2500, height: 843 }, selected: true, name: '追星記票', chatBarText: '選單',
    areas: labels.map(function (label, i) {
      return { bounds: { x: (i % 4) * 625, y: i < 4 ? 0 : 421, width: 625, height: i < 4 ? 421 : 422 }, action: { type: 'message', text: label } };
    }),
  };
  const res = lineApi_('https://api.line.me/v2/bot/richmenu', menu);
  if (res.getResponseCode() !== 200) return '建立選單失敗：' + res.getContentText();
  const id = JSON.parse(res.getContentText()).richMenuId;
  const img = UrlFetchApp.fetch(RICHMENU_IMAGE).getBlob();
  const up = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/' + id + '/content', {
    method: 'post', contentType: 'image/png', payload: img.getBytes(), headers: head, muteHttpExceptions: true,
  });
  if (up.getResponseCode() !== 200) return '上傳選單圖片失敗：' + up.getContentText();
  UrlFetchApp.fetch('https://api.line.me/v2/bot/user/all/richmenu/' + id, { method: 'post', headers: head, muteHttpExceptions: true });
  PROPS.setProperty('RICHMENU_ID', id);
  return '已建立聊天室下方選單 ✦';
}

/* ---------- 搶票提醒 ---------- */
// 'YYYY-MM-DD HH:MM'（台北時間）加減分鐘
function shiftTime_(at, minutes) {
  const t = new Date(at.replace(' ', 'T') + ':00+08:00');
  t.setTime(t.getTime() + minutes * 60000);
  return Utilities.formatDate(t, 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
}
const nowTaipei_ = function () { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'); };
function normDateTime_(v) {
  const m = str_(v).match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?(?:[ T]*(\d{1,2})[:：點](\d{2})?)?/);
  if (!m) return '';
  const p2 = function (n) { return ('0' + Number(n || 0)).slice(-2); };
  return m[1] + '-' + p2(m[2]) + '-' + p2(m[3]) + (m[4] ? ' ' + p2(m[4]) + ':' + p2(m[5]) : '');
}
const hasTime_ = function (at) { return /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(at); };
const normPlatform_ = function (p) {
  p = str_(p);
  return p ? (LINE_PLATFORM_ALIAS[p.toLowerCase()] || LINE_PLATFORMS.filter(function (x) { return x.toLowerCase() === p.toLowerCase(); })[0] || p) : '';
};
const atText_ = function (at) { return md_(at.slice(0, 10)) + (hasTime_(at) ? ' ' + at.slice(11) : ''); };
const jsonList_ = function (v) { try { const a = JSON.parse(v || '[]'); return Array.isArray(a) ? a : []; } catch (err) { return []; } };

function onsaleFrom_(x) {
  return {
    title: str_(x.eventName), artist: str_(x.artist), venue: str_(x.venue), city: str_(x.city),
    showDates: (Array.isArray(x.showDates) ? x.showDates : []).map(str_).filter(Boolean).join('、'),
    sales: (Array.isArray(x.sales) ? x.sales : []).map(function (sl) {
      return { phase: str_(sl.phase) || '開賣', saleAt: normDateTime_(sl.saleAt), platform: normPlatform_(sl.platform) };
    }).filter(function (sl) { return sl.saleAt; }),
    price: str_(x.priceInfo), notes: str_(x.notice),
  };
}
function startOnsale_(token, uid, s, data) {
  s.o = onsaleFrom_(data);
  s.step = 'oconfirm';
  saveLineSession_(uid, s);
  lineReply_(token, ['這是售票公告 ✦ 確認一下要建立的搶票提醒：', onsaleCard_(s.o)]);
}
function mergeOnsale_(token, uid, s, x) {
  const n = onsaleFrom_(x || {});
  ['title', 'artist', 'venue', 'city', 'showDates', 'price', 'notes'].forEach(function (k) { if (!s.o[k] && n[k]) s.o[k] = n[k]; });
  n.sales.forEach(function (sl) {
    if (!s.o.sales.some(function (o) { return o.saleAt === sl.saleAt && o.phase === sl.phase; })) s.o.sales.push(sl);
  });
  s.o.sales.sort(function (a, b) { return a.saleAt.localeCompare(b.saleAt); });
  s.imgs = (s.imgs || []).concat(s.pending || []);
  s.pending = [];
  delete s.pendingData;
  s.step = 'oconfirm';
  saveLineSession_(uid, s);
  lineReply_(token, ['已合併 ✦', onsaleCard_(s.o)]);
}

function onsaleCard_(o) {
  const row = function (k, v) {
    return { type: 'box', layout: 'baseline', spacing: 'md', contents: [
      { type: 'text', text: k, size: 'sm', color: '#9C917F', flex: 2 },
      { type: 'text', text: str_(v) || '—', size: 'sm', color: '#40382E', flex: 6, wrap: true },
    ] };
  };
  const sales = o.sales.length ? o.sales.map(function (sl) {
    return '・' + sl.phase + ' ' + atText_(sl.saleAt) + (hasTime_(sl.saleAt) ? '' : '（時間未定）') + (sl.platform ? '｜' + sl.platform : '');
  }).join('\n') : '沒讀到開賣時間，請直接打字告訴我，例如「全面開賣 11/3 12:00 拓元」';
  const button = function (label, v, style) {
    return { type: 'button', style: style, height: 'sm', color: style === 'primary' ? '#A67B5B' : undefined,
      action: { type: 'postback', label: label, data: 'a=ocard&v=' + v, displayText: label } };
  };
  return {
    type: 'flex', altText: '搶票提醒：' + (o.title || '售票公告'),
    contents: { type: 'bubble',
      body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
        { type: 'text', text: '搶票提醒', size: 'xs', color: '#A67B5B', weight: 'bold' },
        { type: 'text', text: o.title || '（沒有節目名稱）', weight: 'bold', size: 'lg', wrap: true, color: '#40382E' },
        { type: 'separator', margin: 'md' },
        row('表演者', o.artist),
        row('演出', [o.showDates, [o.city, o.venue].filter(Boolean).join(' ')].filter(Boolean).join('\n')),
        row('開賣', sales),
        row('票價', o.price),
        row('注意', o.notes),
        { type: 'text', text: '要修改直接打字告訴我，例如「全面開賣改成 11/4 12:00 拓元」', size: 'xxs', color: '#9C917F', wrap: true, margin: 'md' },
      ] },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
        button('建立提醒', 'confirm', 'primary'), button('要修改', 'edit', 'secondary'), button('取消', 'cancel', 'link'),
      ] },
    },
  };
}

const ONSALE_EDIT_SCHEMA = {
  type: 'OBJECT',
  properties: {
    understood: { type: 'BOOLEAN' }, title: { type: 'STRING' }, artist: { type: 'STRING' }, venue: { type: 'STRING' },
    city: { type: 'STRING' }, showDates: { type: 'STRING' }, price: { type: 'STRING' }, notes: { type: 'STRING' },
    sales: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
      phase: { type: 'STRING' }, saleAt: { type: 'STRING' }, platform: { type: 'STRING' } } } },
  },
  required: ['understood'],
};
function editOnsale_(o, text) {
  const prompt = [
    '今天是 ' + today_() + '。以下是一筆搶票提醒（JSON）：', JSON.stringify(o),
    '使用者說：「' + text + '」',
    '請輸出 understood: true，並只包含要修改的欄位。如果改到開賣時間，sales 要輸出修改後「完整」的每一波開賣（沒改的也要列出）；',
    'saleAt 用 YYYY-MM-DD HH:MM（24 小時制，沒寫年份就用今天之後最近的日期）。看不懂就只輸出 understood: false。',
  ].join('\n');
  const r = gemini_([{ text: prompt }], ONSALE_EDIT_SCHEMA);
  if (r.error) return { error: r.error };
  const c = r.data || {};
  if (!c.understood) return { changed: false };
  let changed = false;
  ['title', 'artist', 'venue', 'city', 'showDates', 'price', 'notes'].forEach(function (k) {
    if (c[k] !== undefined) { o[k] = str_(c[k]); changed = true; }
  });
  if (Array.isArray(c.sales)) {
    o.sales = onsaleFrom_({ sales: c.sales }).sales.sort(function (a, b) { return a.saleAt.localeCompare(b.saleAt); });
    changed = true;
  }
  return { changed: changed };
}

function confirmOnsale_(token, uid, me, s) {
  const o = s.o;
  if (!o.title) return lineReply_(token, '還缺節目名稱，直接打字告訴我，例如「節目是 DAY6 台北場」。');
  if (!o.sales.some(function (sl) { return hasTime_(sl.saleAt); })) return lineReply_(token, '還缺開賣時間，直接打字告訴我，例如「全面開賣 11/3 12:00 拓元」。');
  const groupId = newId_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setup_();
    writeRows_('onsales', o.sales.map(function (sl) {
      return {
        id: newId_(), groupId: groupId, title: o.title, artist: o.artist, venue: o.venue, city: o.city, showDates: o.showDates,
        phase: sl.phase, saleAt: sl.saleAt, platform: sl.platform, price: o.price, notes: o.notes,
        watchers: JSON.stringify([uid]), done: '[]', sentEve: '', sentLead: '', createdBy: me.name || '', createdAt: today_(),
      };
    }));
  } finally {
    lock.releaseLock();
  }
  clearLineSession_(uid);
  const c = reminderSettings_();
  const next = o.sales.filter(function (sl) { return hasTime_(sl.saleAt) && sl.saleAt > nowTaipei_(); })[0];
  const when = next ? [
    c.onsaleEve !== 'off' ? atText_(shiftTime_(next.saleAt.slice(0, 10) + ' 12:00', -1440).slice(0, 10) + ' ' + c.onsaleEve) + ' 預告' : '',
    c.onsaleLead !== 'off' ? atText_(shiftTime_(next.saleAt, -Number(c.onsaleLead))) + ' 提醒' : '',
  ].filter(Boolean).join('、') : '';
  lineReply_(token, ask_('已建立搶票提醒 ✦ ' + o.title + (when ? '\n會在 ' + when + '你。' : '\n（開賣時間都已經過了，不會再提醒）')
    + '\n\n想揪朋友一起搶，可以按下面的按鈕通知大家。', 'onotify', [[groupId, '通知大家一起搶']]));
}

function onsaleGroups_() {
  const groups = {};
  sheetRows_('onsales').forEach(function (r) {
    (groups[r.groupId] = groups[r.groupId] || []).push(r);
  });
  return Object.keys(groups).map(function (id) {
    const rows = groups[id].sort(function (a, b) { return String(a.saleAt).localeCompare(String(b.saleAt)); });
    return { id: id, rows: rows, first: rows[0], watchers: jsonList_(rows[0].watchers), done: jsonList_(rows[0].done) };
  });
}
function salesLines_(rows) {
  return rows.map(function (r) { return '・' + r.phase + ' ' + atText_(r.saleAt) + (r.platform ? '｜' + r.platform : ''); }).join('\n');
}

function onsaleNotify_(token, uid, me, groupId) {
  const g = onsaleGroups_().filter(function (x) { return x.id === groupId; })[0];
  if (!g) return lineReply_(token, '找不到這個搶票提醒，可能已經被刪除了。');
  const others = Object.keys(lineUsers_()).filter(function (id) { return id !== uid && g.watchers.indexOf(id) < 0; });
  if (!others.length) return lineReply_(token, '其他人都已經在這個搶票提醒裡了（或目前只有你加入機器人）。');
  const msg = ask_('📣 ' + (me.name || '有人') + ' 建立了新的搶票：' + g.first.title + '\n' + salesLines_(g.rows)
    + (g.first.notes ? '\n注意：' + g.first.notes : '') + '\n\n要一起搶的話按「我也要搶」，開賣前會提醒你。', 'watch', [[groupId, '我也要搶']]);
  const sent = linePush_(others, [msg]);
  lineReply_(token, sent ? '已通知 ' + others.length + ' 個人 ✦' : '本月 LINE 主動提醒額度不夠，沒有送出通知。');
}

function onsaleMark_(token, uid, action, groupId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let title = '';
  try {
    const rows = sheetRows_('onsales').filter(function (r) { return r.groupId === groupId; });
    if (!rows.length) return lineReply_(token, '找不到這個搶票提醒，可能已經被刪除了。');
    title = rows[0].title;
    rows.forEach(function (r) {
      const key = action === 'watch' || action === 'unwatch' ? 'watchers' : 'done';
      let list = jsonList_(r[key]).filter(function (id) { return id !== uid; });
      if (action !== 'unwatch') list = list.concat([uid]);
      r[key] = JSON.stringify(list);
    });
    writeRows_('onsales', rows);
  } finally {
    lock.releaseLock();
  }
  const msg = {
    watch: '好 ✦ 開賣前會提醒你：' + title, unwatch: '好，不會再提醒你：' + title,
    ogot: '恭喜搶到 🎉 ' + title + '\n記得把購票截圖傳給我建檔。', omiss: '辛苦了 🥲 ' + title + ' 後面的提醒不會再傳給你。',
  }[action];
  lineReply_(token, msg);
}

// 建檔一張票時，順便把同一個節目的搶票提醒標成這個人已搶到
function markOnsaleGot_(uid, d) {
  const now = nowTaipei_();
  const n = normName_(d.eventName), ar = normName_(d.artist);
  const hits = onsaleGroups_().filter(function (g) {
    if (g.watchers.indexOf(uid) < 0 || g.done.indexOf(uid) >= 0 || String(g.first.saleAt) > now) return false;
    const t = normName_(g.first.title), ga = normName_(g.first.artist);
    return (n && t && (t.indexOf(n) >= 0 || n.indexOf(t) >= 0)) || (ar && ga === ar);
  });
  if (!hits.length) return [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    hits.forEach(function (g) {
      g.rows.forEach(function (r) { r.done = JSON.stringify(jsonList_(r.done).concat([uid])); });
      writeRows_('onsales', g.rows);
    });
  } finally {
    lock.releaseLock();
  }
  return hits.map(function (g) { return g.first.title; });
}

function onsaleListMessage_(uid, q) {
  const now = nowTaipei_();
  const list = onsaleGroups_().filter(function (g) {
    return g.rows.some(function (r) { return String(r.saleAt) >= now.slice(0, 10); })
      && matchEvent_({ name: g.first.title, artist: g.first.artist, venue: g.first.venue }, q || {});
  }).map(function (g) {
    g.upcoming = g.rows.filter(function (r) { return String(r.saleAt) >= now.slice(0, 10); });
    return g;
  }).sort(function (a, b) { return String(a.upcoming[0].saleAt).localeCompare(String(b.upcoming[0].saleAt)); });
  if (!list.length) return '目前沒有即將開賣的節目。傳售票公告截圖給我就可以建立搶票提醒 ✦';
  const entries = list.slice(0, 36).map(function (g) {
    const next = g.upcoming[0];
    const mine = g.watchers.indexOf(uid) >= 0;
    return {
      date: next.saleAt.slice(0, 10), time: next.saleAt.slice(11), top: g.first.artist, title: g.first.title,
      lines: g.upcoming.map(function (r) { return '🎫 ' + r.phase + '　' + shortMD_(r.saleAt.slice(0, 10)) + ' ' + r.saleAt.slice(11) + (r.platform ? '｜' + r.platform : ''); })
        .concat([g.first.notes ? '⚠ ' + g.first.notes : '']),
      status: mine ? (g.done.indexOf(uid) >= 0 ? '已結束' : '👀 你要搶') : '', statusColor: C_ACCENT,
    };
  });
  const join = list.filter(function (g) { return g.watchers.indexOf(uid) < 0; }).slice(0, 13);
  return agendaMessage_('即將開賣', entries, { sub: '左邊是最近一波的開賣時間', foot: join.length ? '想一起搶的可以點下方按鈕。' : '',
    step: 'watch', quick: join.map(function (g) { return [g.id, '我也要搶：' + g.first.title]; }) });
}

// 推播給指定的人；額度不夠就不傳（回傳 false），快用完時附上提醒
function linePush_(to, messages) {
  to = to.filter(Boolean);
  messages = messages.filter(Boolean).map(function (m) { return typeof m === 'string' ? { type: 'text', text: m } : m; });
  if (!to.length || !messages.length) return true;
  const head = { Authorization: 'Bearer ' + PROPS.getProperty('LINE_TOKEN') };
  try {
    const quota = JSON.parse(UrlFetchApp.fetch('https://api.line.me/v2/bot/message/quota', { headers: head, muteHttpExceptions: true }).getContentText());
    const used = JSON.parse(UrlFetchApp.fetch('https://api.line.me/v2/bot/message/quota/consumption', { headers: head, muteHttpExceptions: true }).getContentText());
    if (quota.type === 'limited') {
      const left = num_(quota.value) - num_(used.totalUsage);
      if (left < to.length) { Logger.log('LINE 本月推播額度不足，剩 ' + left + ' 則'); return false; }
      if (left - to.length < 20) messages = messages.concat([{ type: 'text', text: '⚠ 本月 LINE 主動提醒額度只剩 ' + (left - to.length) + ' 則（每月 200 則，下個月重算）。' }]);
    }
  } catch (err) {
    Logger.log('查詢額度失敗：' + err);
  }
  for (let i = 0; i < to.length; i += 500) {
    lineApi_('https://api.line.me/v2/bot/message/multicast', { to: to.slice(i, i + 500), messages: messages.slice(0, 5) });
  }
  return true;
}

function onsaleTick_(now) {
  const c = reminderSettings_();
  if (c.onsaleEve === 'off' && c.onsaleLead === 'off') return;
  const users = lineUsers_();
  const perUser = {}; // uid → { eve: [], lead: [] }
  const touched = [];
  onsaleGroups_().forEach(function (g) {
    g.rows.forEach(function (r) {
      if (!hasTime_(r.saleAt) || r.saleAt < shiftTime_(now, -10)) return;
      const leadAt = c.onsaleLead === 'off' ? r.saleAt : shiftTime_(r.saleAt, -Number(c.onsaleLead));
      const eveAt = shiftTime_(r.saleAt.slice(0, 10) + ' 12:00', -1440).slice(0, 10) + ' ' + c.onsaleEve;
      let kind = '';
      if (c.onsaleLead !== 'off' && !r.sentLead && now >= leadAt && now <= r.saleAt) { kind = 'lead'; r.sentLead = now; }
      else if (c.onsaleEve !== 'off' && !r.sentEve && now >= eveAt && now < leadAt) { kind = 'eve'; r.sentEve = now; }
      if (!kind) return;
      touched.push(r);
      g.watchers.filter(function (id) { return users[id] && g.done.indexOf(id) < 0; }).forEach(function (id) {
        (perUser[id] = perUser[id] || { eve: [], lead: [] })[kind].push(r);
      });
    });
  });
  if (!touched.length) return;
  writeRows_('onsales', touched);
  Object.keys(perUser).forEach(function (id) {
    const x = perUser[id];
    const msgs = [];
    const line = function (r) {
      return r.title + '\n' + r.phase + ' ' + (r.saleAt.slice(0, 10) === now.slice(0, 10) ? '今天' : '明天') + ' ' + r.saleAt.slice(11)
        + (r.platform ? '｜' + r.platform : '') + (r.notes ? '\n注意：' + r.notes : '');
    };
    if (x.eve.length) msgs.push('🔔 搶票預告\n\n' + x.eve.map(line).join('\n\n'));
    if (x.lead.length) {
      msgs.push(ask_('⏰ 快開賣了！\n\n' + x.lead.map(function (r) {
        return r.title + '\n' + r.phase + ' ' + r.saleAt.slice(11) + ' 開賣' + (r.platform ? '｜' + r.platform : '');
      }).join('\n\n') + '\n\n搶完可以點下面告訴我結果。', 'ogot',
        x.lead.slice(0, 6).map(function (r) { return [r.groupId, '搶到了：' + r.title]; })
          .concat(x.lead.slice(0, 6).map(function (r) { return ['miss:' + r.groupId, '沒搶到：' + r.title]; }))));
    }
    linePush_([id], msgs);
  });
}

/* ---------- 換售資訊：產生售票／換票文章 ---------- */
const TRADE_NOTES_DEFAULT = {
  sell: ['請站內信留Line ID，請確定要購買再來信', '需先匯款全額，前5天可取票時給序號自取', '板上有多次讓票紀錄供參考，若有疑慮請勿來信，謝謝'],
  swap: ['互補差價，請站內信留Line ID', '匯款後待可取票時提供序號自取', '板上有多次讓換票紀錄供參考，謝謝'],
};
const TRADE_LABEL = { sell: '售票', swap: '換票' };
function tradeNotes_(mode) {
  const v = PROPS.getProperty(mode === 'sell' ? 'TRADE_NOTES_SELL' : 'TRADE_NOTES_SWAP');
  return v ? JSON.parse(v) : TRADE_NOTES_DEFAULT[mode];
}
function tradeNotesText_(mode) {
  return TRADE_LABEL[mode] + '固定備註：\n' + tradeNotes_(mode).map(function (n, i) { return (i + 1) + '.' + n; }).join('\n')
    + '\n\n要修改就打「' + TRADE_LABEL[mode] + '備註改成：」，換行後一行一條。';
}
function setTradeNotes_(mode, text) {
  const body = text.replace(/^[^：:]*[：:]/, '');
  const list = body.split(/\r?\n/).map(function (x) { return x.replace(/^\s*\d+\s*[.、)）]\s*/, '').trim(); }).filter(Boolean);
  if (!list.length) return '沒有讀到備註內容。可以這樣打：\n' + TRADE_LABEL[mode] + '備註改成：\n請站內信留Line ID\n需先匯款全額';
  PROPS.setProperty(mode === 'sell' ? 'TRADE_NOTES_SELL' : 'TRADE_NOTES_SWAP', JSON.stringify(list));
  return '已修改 ✦\n\n' + tradeNotesText_(mode);
}
const tradeSession_ = function (uid) { const raw = cache_().get('lt_' + uid); return raw ? JSON.parse(raw) : null; };
const saveTradeSession_ = function (uid, t) { cache_().put('lt_' + uid, JSON.stringify(t), 1800); };
const clearTradeSession_ = function (uid) { cache_().remove('lt_' + uid); };

function tradeStart_() {
  return ask_('要產生哪一種文章？', 'trade', [['sell', '售票'], ['swap', '換票']]);
}
function tradeCandidates_() {
  const byId = {};
  sheetRows_('events').forEach(function (e) { byId[e.id] = e; });
  return unsoldTransfers_().map(function (t) { return transferItem_(t, byId[t.eventId]); })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }).slice(0, 30);
}

const shortMD_ = function (date) { return /^\d{4}-\d{2}-\d{2}$/.test(date) ? Number(date.slice(5, 7)) + '/' + Number(date.slice(8, 10)) : date; };
const tradeLine_ = function (it) { return shortMD_(it.date) + ' ' + (it.artist || it.name) + '\n   ' + [it.area, it.row ? it.row + '排' : ''].join('') + ' ×' + it.count; };

function tradePick_(token, uid, mode) {
  const cands = tradeCandidates_();
  if (!cands.length) return lineReply_(token, '轉賣中沒有還沒賣出的票。要放進換售資訊的票，記票時選「要轉賣」就會出現在這裡。');
  saveTradeSession_(uid, { mode: mode, step: 'pick', cands: cands });
  lineReply_(token, '要' + TRADE_LABEL[mode] + '的是哪幾張？回覆編號（可多選，例如：1 3），或打「全部」。\n\n'
    + cands.map(function (it, i) { return (i + 1) + '. ' + tradeLine_(it); }).join('\n'));
}

function onTradeText_(token, uid, t, text) {
  text = text.trim();
  if (t.step === 'pick') {
    const nums = /全部/.test(text) ? t.cands.map(function (x, i) { return i + 1; }) : (text.match(/\d+/g) || []).map(Number);
    const picked = nums.filter(function (n, i) { return n >= 1 && n <= t.cands.length && nums.indexOf(n) === i; }).map(function (n) { return t.cands[n - 1]; });
    if (!picked.length) return lineReply_(token, '請回覆清單上的編號，例如：1 3');
    t.picked = picked;
    if (t.mode === 'swap') {
      t.step = 'want';
      saveTradeSession_(uid, t);
      return lineReply_(token, '想換什麼？例如「5/17 孫燕姿 3880以下*2」');
    }
    return tradeAskDelivery_(token, uid, t);
  }
  if (t.step === 'want') {
    t.want = text;
    return tradeAskDelivery_(token, uid, t);
  }
  if (t.step === 'deliver') {
    if (/面交/.test(text)) return tradeSetDelivery_(token, uid, t, 'meet');
    if (/序號|自取/.test(text)) return tradeSetDelivery_(token, uid, t, 'code');
    return lineReply_(token, tradeDeliveryQuestion_());
  }
  if (t.step === 'place') return tradeSetPlace_(token, uid, t, text);
  if (t.step === 'note') return tradeFinish_(token, uid, t, /^(不用|不要|沒有|無|免|跳過|no|skip)$/i.test(text) ? '' : text);
}

function tradeDeliveryQuestion_() {
  return ask_('交易方式？', 'tdeliv', [['code', '給序號自取'], ['meet', '面交']]);
}
function tradeAskDelivery_(token, uid, t) {
  t.step = 'deliver';
  saveTradeSession_(uid, t);
  lineReply_(token, tradeDeliveryQuestion_());
}
function tradeSetDelivery_(token, uid, t, how) {
  t.delivery = how;
  if (how === 'meet') {
    t.step = 'place';
    saveTradeSession_(uid, t);
    return lineReply_(token, ask_('在哪裡面交？（其他地點可以直接打字）', 'tplace', [['台北市', '台北市']]));
  }
  t.step = 'note';
  saveTradeSession_(uid, t);
  lineReply_(token, tradeNoteQuestion_());
}
function tradeSetPlace_(token, uid, t, place) {
  t.place = str_(place) || '台北市';
  t.step = 'note';
  saveTradeSession_(uid, t);
  lineReply_(token, tradeNoteQuestion_());
}
function tradeNoteQuestion_() {
  return ask_('這次還要加其他備註嗎？要的話直接打字；不用的話點下面的「不用」。', 'tnote', [['__none', '不用']]);
}

// 票價只寫票面，系統服務費另外註明：6980+系統服務費200
const tradePrice_ = function (it) { return (it.face || '') + (it.fee ? '+系統服務費' + it.fee : ''); };
// 距離演出幾天可以取票；已經可以取（或沒有設定取票日）回傳 0
function pickupDaysLeft_(it, today) {
  if (!it.pickupDate || it.pickupDate <= today || !/^\d{4}-\d{2}-\d{2}$/.test(it.date)) return 0;
  return Math.round((new Date(it.date + 'T00:00:00Z') - new Date(it.pickupDate + 'T00:00:00Z')) / 86400000);
}
// 交易方式那一條備註：依每張票真正的取票時間產生
function deliveryNote_(t, today) {
  const label = function (it) { return it.artist || it.name; };
  const waiting = t.picked.filter(function (it) { return pickupDaysLeft_(it, today) > 0; });
  let when = '';
  if (waiting.length) {
    const days = waiting.map(function (it) { return pickupDaysLeft_(it, today); });
    const same = days.every(function (x) { return x === days[0]; }) && waiting.length === t.picked.length;
    when = same ? '前' + days[0] + '天可取票時' : waiting.map(function (it, i) { return label(it) + '前' + days[i] + '天'; }).join('、') + '可取票時';
  }
  if (t.delivery === 'meet') return (t.place || '台北市') + '面交' + (when ? '（' + when + '）' : '');
  if (t.mode === 'sell') return '需先匯款全額，' + (when ? when + '給序號自取' : '匯款後給序號自取');
  return when ? '匯款後待' + when + '提供序號自取' : '匯款後提供序號自取';
}

function tradeFinish_(token, uid, t, extra) {
  const today = today_();
  const fixed = tradeNotes_(t.mode);
  const dn = deliveryNote_(t, today);
  // 固定備註裡講交易方式的那一條，換成依這幾張票產生的內容
  let notes = fixed.map(function (n) { return /序號|面交|自取/.test(n) ? dn : n; })
    .filter(function (n, i, all) { return n !== dn || all.indexOf(dn) === i; });
  if (notes.indexOf(dn) < 0) notes.splice(Math.min(1, notes.length), 0, dn);
  if (t.mode === 'swap') { // 有系統服務費的票，自動補一條說明
    const byFee = {};
    t.picked.forEach(function (it) {
      if (!it.fee) return;
      const k = String(it.fee), name = it.artist || it.name;
      byFee[k] = byFee[k] || [];
      if (byFee[k].indexOf(name) < 0) byFee[k].push(name);
    });
    Object.keys(byFee).forEach(function (fee) { notes.push(byFee[fee].join('、') + '的票價每張皆需再+系統服務費' + fee); });
  }
  if (extra && extra !== '__none') notes.push(extra);
  const notesText = '備註:\n' + notes.map(function (n, i) { return (i + 1) + '.' + n; }).join('\n');
  const timeText = function (it) {
    return /^\d{4}-\d{2}-\d{2}$/.test(it.date) ? it.date.replace(/-/g, '/') + ' (' + WEEK_[new Date(it.date + 'T00:00:00Z').getUTCDay()] + ')' + (it.time ? ' ' + it.time : '') : it.date;
  };
  let posts;
  if (t.mode === 'sell') {
    posts = t.picked.slice(0, 4).map(function (it) {
      const one = Object.assign({}, t, { picked: [it] });
      const itsNotes = notes.map(function (n) { return n === dn ? deliveryNote_(one, today) : n; });
      return [
        shortMD_(it.date) + ' ' + (it.artist || it.name) + ' ' + String(it.area || '').replace(/區$/, '') + '*' + it.count,
        '節目：' + it.name,
        '地點：' + (it.venue || it.city || ''),
        '時間：' + timeText(it),
        '位置：' + String(it.area || '') + (it.row ? it.row + '排' : ''),
        '票價：' + tradePrice_(it),
        '張數：' + it.count,
        '',
        '備註:\n' + itsNotes.map(function (n, i) { return (i + 1) + '.' + n; }).join('\n'),
      ].join('\n');
    });
  } else {
    const heads = [];
    t.picked.forEach(function (it) { const h = shortMD_(it.date) + ' ' + (it.artist || it.name); if (heads.indexOf(h) < 0) heads.push(h); });
    const wantHead = String(t.want || '').replace(/\s*\d{3,}.*$/, '').trim() || t.want;
    posts = [[
      heads.join('、') + ' 換 ' + wantHead,
      '',
      '持有',
      t.picked.map(function (it) {
        return shortMD_(it.date) + ' ' + (it.artist || it.name) + String(it.area || '') + (it.row ? it.row + '排' : '') + ' ' + (it.face || '') + '*' + it.count;
      }).join('\n'),
      '',
      '欲換',
      t.want,
      '',
      notesText,
    ].join('\n')];
  }
  clearTradeSession_(uid);
  lineReply_(token, ['以下是' + TRADE_LABEL[t.mode] + '文章，長按就能複製 ✦（打「' + TRADE_LABEL[t.mode] + '備註」可以看或修改固定備註）'].concat(posts));
}
