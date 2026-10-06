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
  ledger: ['id', 'type', 'category', 'date', 'title', 'eventId', 'amountTwd', 'currency', 'originalAmount', 'exchangeRate', 'payer', 'paymentMethod', 'paymentDetail', 'counterparty', 'expectedReceivableTwd', 'receivedTwd', 'notes', 'ticketType', 'ticketArea', 'ticketRow', 'ticketSeat', 'attendee', 'ticketStatus', 'createdAt', 'settled', 'ticketFaceTwd', 'ticketBenefitTwd', 'ticketFeeTwd', 'ticketPlatform', 'ticketAccount', 'ticketCount', 'splits', 'ticketPickupDate', 'ticketPickedUp', 'ticketOrderNumber'],
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
  '- account 是購買人的帳號、姓名、信箱或電話（截圖上看得到才填）。orderNumber 是訂單編號。pickupDate 是可取票日期（YYYY-MM-DD）。',
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
    currency: { type: 'STRING' }, unitFace: { type: 'NUMBER' }, unitFee: { type: 'NUMBER' }, totalPaid: { type: 'NUMBER' },
    platform: { type: 'STRING' }, account: { type: 'STRING' },
    orderNumber: { type: 'STRING' }, pickupDate: { type: 'STRING' },
    uncertain: { type: 'ARRAY', items: { type: 'STRING' } },
  },
  required: ['eventName', 'date', 'seats', 'currency', 'totalPaid', 'ticketCount'],
};

function scanTicket_(req) {
  if (req.test === 'quota') return { error: '今天的 AI 免費次數用完了（這是模擬測試），可以明天再試，或先手動填寫' };
  const apiKey = PROPS.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { error: '還沒設定 Gemini 金鑰：請到 Apps Script「專案設定 → 指令碼屬性」新增 GEMINI_API_KEY' };
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const parts = [{ text: SCAN_PROMPT.replace('{TODAY}', today) }];
  (req.images || []).forEach(function (img) {
    parts.push({ inline_data: { mime_type: img.mimeType || 'image/jpeg', data: img.dataBase64 } });
  });
  // 讀票不需要深度思考；思考太久會讓手機瀏覽器等超過 60 秒而斷線
  const config = { responseMimeType: 'application/json', responseSchema: SCAN_SCHEMA, temperature: 0, thinkingConfig: { thinkingLevel: 'low' } };
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
