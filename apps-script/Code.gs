// 追星總帳 Google Sheets API
// 部署方式：擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業 → 網頁應用程式
// 「執行身分：我」、「誰可以存取：知道連結的任何人」→ 複製網頁應用程式網址貼到 App 設定頁。

const SHARED_KEY = ''; // 可自訂一組密碼，App 設定頁需填相同的值；留空表示不驗證
const SHEET_ID = ''; // 從試算表「擴充功能→Apps Script」開的專案留空；獨立專案填試算表網址 /d/ 後面那串 ID

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

const TABLES = {
  orders: ['id', 'orderNumber', 'channel', 'orderDate', 'estimatedShipDate', 'actualShipDate', 'currency', 'domesticShipping', 'internationalShippingTwd', 'internationalShippingRateTwdPerKg', 'discountAmount', 'weightGrams', 'exchangeRate', 'chargedTwd', 'payer', 'paymentMethod', 'paymentDetail', 'settled', 'notes'],
  items: ['id', 'orderId', 'name', 'variant', 'unitPrice', 'quantity', 'ownership', 'proxyFor', 'arrived', 'sorted', 'proxyPaid', 'salePriceTwd', 'soldQuantity'],
  sales: ['id', 'sourceOrderId', 'sourceItemId', 'sourceOrderNumber', 'sourceChannel', 'name', 'variant', 'sourceCurrency', 'unitOriginalPrice', 'unitCostTwd', 'quantity', 'salePriceTwd', 'soldQuantity', 'managedByOwnership', 'createdAt'],
  events: ['id', 'name', 'artist', 'city', 'venue', 'startDate', 'endDate', 'eventNumber', 'originalDate', 'eventType', 'liveTour', 'seriesEvent', 'seat', 'ticketPriceTwd', 'guest', 'payer', 'settled', 'notes', 'createdAt', 'coverUrl'],
  ledger: ['id', 'type', 'category', 'date', 'title', 'eventId', 'amountTwd', 'currency', 'originalAmount', 'exchangeRate', 'payer', 'paymentMethod', 'paymentDetail', 'counterparty', 'expectedReceivableTwd', 'receivedTwd', 'notes', 'ticketType', 'ticketArea', 'ticketRow', 'ticketSeat', 'attendee', 'ticketStatus', 'createdAt', 'settled', 'ticketFaceTwd', 'ticketBenefitTwd', 'ticketFeeTwd', 'ticketPlatform', 'ticketAccount', 'ticketCount', 'splits', 'ticketPickupDate', 'ticketPickedUp'],
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
      sheet.getRange('A:Z').setNumberFormat('@'); // 全文字格式，避免長訂單編號被轉成數字失去精度
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
