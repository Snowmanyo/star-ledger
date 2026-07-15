/* 追星總帳 mobile — Google Sheets 版 */
'use strict';

/* ---------- 工具 ---------- */
const $ = (sel, el) => (el || document).querySelector(sel);
const $$ = (sel, el) => Array.from((el || document).querySelectorAll(sel));
const esc = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const uid = () => Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
const num = v => {
  const n = Number(String(v == null ? '' : v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
};
const toBool = v => v === true || v === 'true' || v === 'TRUE';
const today = () => new Date().toISOString().slice(0, 10);
const fmtInt = n => Math.round(num(n)).toLocaleString('en-US');
const fmtTwd = n => 'TWD ' + fmtInt(n);
const fmtMoney = (cur, n) => {
  if (!cur || cur === 'TWD') return fmtTwd(n);
  if (cur === 'USD') return 'USD ' + num(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return cur + ' ' + fmtInt(n);
};

/* ---------- 常數 ---------- */
const THEMES = [
  { id: 'milktea', label: '奶茶', bg: '#F6F1E9', accent: '#A67B5B' },
  { id: 'matcha', label: '抹茶', bg: '#F1F4EC', accent: '#7E9B6E' },
  { id: 'seasalt', label: '海鹽', bg: '#EEF3F5', accent: '#5F8CA3' },
  { id: 'sakura', label: '櫻花', bg: '#F9F1F1', accent: '#C4798F' },
  { id: 'night', label: '夜幕', bg: '#201F1D', accent: '#C9A05F' },
];
const CURRENCIES = ['KRW', 'JPY', 'USD', 'TWD'];
const OWN_LABEL = { self: '自留', proxy: '代購', stock: '現貨', pending: '待補' };
const PAY_LABEL = { cash: '現金', credit_card: '信用卡', bank_transfer: '轉帳', mobile_payment: '行動支付', '': '—' };
const CAT_LABEL = { ticket: '票券', transport: '交通', lodging: '住宿', food: '餐飲', merch: '周邊', other: '其他' };
const TAB_TITLE = { orders: '訂單', sales: '販售', events: '活動', ledger: '總帳', settings: '設定' };
const ICONS = {
  copy: '<svg viewBox="0 0 24 24"><rect x="9" y="9" width="11" height="11" rx="2" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M5 15V6a2 2 0 0 1 2-2h9" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M10 7V5h4v2M6.5 7l1 13h9l1-13M10 11v5.5M14 11v5.5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  close: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
  back: '<svg viewBox="0 0 24 24"><path d="M14.5 5.5 8 12l6.5 6.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>',
};
function sheetTitleHtml(title, withDelete, delId, withBack) {
  return `<div class="sheet-title">
    <span style="display:flex;align-items:center;gap:10px;min-width:0">
      ${withBack ? `<button class="icon-mini" id="sh-back" aria-label="返回">${ICONS.back}</button>` : ''}
      <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${title}</span>
    </span>
    <span class="sheet-title-actions">
    ${withDelete ? `<button class="icon-mini danger" id="${delId}" aria-label="刪除">${ICONS.trash}</button>` : ''}
    <button class="icon-mini" id="sh-close" aria-label="關閉">${ICONS.close}</button>
  </span></div>`;
}
function seatText(l) {
  const part = (v, suf) => v ? (String(v).endsWith(suf) ? String(v) : v + suf) : '';
  return [l.ticketArea || '', part(l.ticketRow, '排'), part(l.ticketSeat, '號')].filter(Boolean).join(' ');
}

const NUMF = {
  orders: ['domesticShipping', 'internationalShippingTwd', 'internationalShippingRateTwdPerKg', 'discountAmount', 'weightGrams', 'exchangeRate', 'chargedTwd'],
  items: ['unitPrice', 'quantity', 'salePriceTwd', 'soldQuantity'],
  sales: ['unitOriginalPrice', 'unitCostTwd', 'quantity', 'salePriceTwd', 'soldQuantity'],
  events: ['ticketPriceTwd'],
  ledger: ['amountTwd', 'originalAmount', 'exchangeRate', 'expectedReceivableTwd', 'receivedTwd'],
};
const BOOLF = {
  orders: ['settled'],
  items: ['arrived', 'sorted', 'proxyPaid'],
  sales: ['managedByOwnership'],
  events: ['settled'],
  ledger: [],
};

/* ---------- 設定與快取 ---------- */
let CFG = Object.assign({ theme: 'milktea', apiUrl: '', key: '', sheetUrl: '', lastSync: '' },
  JSON.parse(localStorage.getItem('sl-config') || '{}'));
const saveCfg = () => localStorage.setItem('sl-config', JSON.stringify(CFG));

let DB = JSON.parse(localStorage.getItem('sl-db') || 'null') || { orders: [], sales: [], events: [], ledger: [] };
const saveDB = () => localStorage.setItem('sl-db', JSON.stringify(DB));
const isDirty = () => localStorage.getItem('sl-dirty') === '1';
const markDirty = () => localStorage.setItem('sl-dirty', '1');
const clearDirty = () => localStorage.removeItem('sl-dirty');

function applyTheme() {
  document.documentElement.dataset.theme = CFG.theme;
  const t = THEMES.find(t => t.id === CFG.theme) || THEMES[0];
  $('meta[name="theme-color"]').setAttribute('content', t.bg);
}

/* ---------- Sheet 資料正規化 ---------- */
function normRow(table, row) {
  const out = Object.assign({}, row);
  NUMF[table].forEach(k => { out[k] = out[k] === '' || out[k] == null ? '' : num(out[k]); });
  BOOLF[table].forEach(k => { out[k] = toBool(out[k]); });
  return out;
}
function assembleDB(raw) {
  const orders = (raw.orders || []).map(r => normRow('orders', r));
  const byOrder = {};
  orders.forEach(o => { o.items = []; byOrder[o.id] = o; });
  (raw.items || []).forEach(r => {
    const it = normRow('items', r);
    const o = byOrder[it.orderId];
    if (o) o.items.push(it);
  });
  return {
    orders,
    sales: (raw.sales || []).map(r => normRow('sales', r)),
    events: (raw.events || []).map(r => normRow('events', r)),
    ledger: (raw.ledger || []).map(r => normRow('ledger', r)),
  };
}
function orderItemRows(order) {
  return (order.items || []).map(it => Object.assign({}, it, { orderId: order.id }));
}
function splitDB(db) {
  const items = [];
  db.orders.forEach(o => items.push(...orderItemRows(o)));
  return { orders: db.orders, items, sales: db.sales, events: db.events, ledger: db.ledger };
}

/* ---------- 雲端 API ---------- */
let syncing = false;
function setSyncIcon(cls) {
  const b = $('#sync-btn');
  b.className = 'icon-btn' + (cls ? ' ' + cls : '');
}
async function apiGet() {
  const url = CFG.apiUrl + (CFG.apiUrl.includes('?') ? '&' : '?') + 'key=' + encodeURIComponent(CFG.key);
  const res = await fetch(url, { redirect: 'follow' });
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}
async function apiPost(body) {
  const res = await fetch(CFG.apiUrl, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ key: CFG.key }, body)),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}
async function doSync(silent) {
  if (!CFG.apiUrl) { toast('尚未設定連線，請到「設定」填入 Apps Script 網址'); setTab('settings'); return; }
  if (isDirty() && !confirm('本機有尚未上傳的變更，同步會以雲端資料覆蓋本機。繼續？')) return;
  if (syncing) return;
  syncing = true;
  setSyncIcon('spin');
  try {
    const raw = await apiGet();
    DB = assembleDB(raw);
    saveDB();
    clearDirty();
    CFG.lastSync = new Date().toLocaleString('zh-TW', { hour12: false });
    saveCfg();
    render();
    setSyncIcon('ok');
    if (!silent) toast('同步完成');
  } catch (err) {
    setSyncIcon('err');
    if (!silent) toast('同步失敗：' + err.message);
  } finally {
    syncing = false;
    setTimeout(() => setSyncIcon(''), 2500);
  }
}
async function pushOps(ops) {
  if (!CFG.apiUrl) { markDirty(); return; }
  setSyncIcon('spin');
  try {
    for (const op of ops) await apiPost(op);
    setSyncIcon('ok');
  } catch (err) {
    markDirty();
    setSyncIcon('err');
    toast('上傳失敗，資料已存在本機：' + err.message);
  } finally {
    setTimeout(() => setSyncIcon(''), 2500);
  }
}
async function uploadAll() {
  if (!CFG.apiUrl) { toast('請先設定 Apps Script 網址'); return; }
  setSyncIcon('spin');
  try {
    await apiPost({ action: 'replaceAll', data: splitDB(DB) });
    clearDirty();
    CFG.lastSync = new Date().toLocaleString('zh-TW', { hour12: false });
    saveCfg();
    setSyncIcon('ok');
    toast('已全部上傳到試算表');
    render();
  } catch (err) {
    setSyncIcon('err');
    toast('上傳失敗：' + err.message);
  } finally {
    setTimeout(() => setSyncIcon(''), 2500);
  }
}

/* ---------- 訂單金額邏輯 ---------- */
const itemsTotal = o => (o.items || []).reduce((s, it) => s + num(it.unitPrice) * num(it.quantity), 0);
const orderTotal = o => itemsTotal(o) + num(o.domesticShipping) - num(o.discountAmount);
function rateLabel(o) {
  const r = num(o.exchangeRate);
  if (!r) return '';
  if (o.currency === 'USD') return 'USD 1 = TWD ' + r.toFixed(2);
  if (o.currency === 'TWD' || !o.currency) return '';
  return 'TWD 1 = ' + o.currency + ' ' + r.toFixed(2);
}
function backCalcRate(o) {
  const total = orderTotal(o), charged = num(o.chargedTwd);
  if (!total || !charged || o.currency === 'TWD') return '';
  return o.currency === 'USD' ? charged / total : total / charged;
}
function orderTwdBase(o) {
  if (num(o.chargedTwd)) return num(o.chargedTwd);
  const total = orderTotal(o), r = num(o.exchangeRate);
  if (o.currency === 'TWD' || !o.currency) return total;
  if (!r) return 0;
  return o.currency === 'USD' ? total * r : total / r;
}
function computeUnitCostTwd(order, item) {
  const base = orderTwdBase(order);
  if (!base) return 0;
  const totalTwd = base + num(order.internationalShippingTwd);
  const itemsSum = itemsTotal(order);
  if (!itemsSum || !num(item.quantity)) return 0;
  const share = (num(item.unitPrice) * num(item.quantity)) / itemsSum;
  return Math.round(totalTwd * share / num(item.quantity));
}

/* ---------- 現貨販售同步 ---------- */
function syncSalesForOrder(order) {
  const ops = { upserts: [], deletes: [] };
  const stockIds = {};
  (order.items || []).forEach(it => {
    if (it.ownership !== 'stock') return;
    stockIds[it.id] = true;
    let sale = DB.sales.find(s => s.sourceItemId === it.id);
    if (!sale) {
      sale = { id: uid(), sourceOrderId: order.id, sourceItemId: it.id, managedByOwnership: true, createdAt: today(), salePriceTwd: num(it.salePriceTwd) || '', soldQuantity: num(it.soldQuantity) || 0 };
      DB.sales.push(sale);
    }
    Object.assign(sale, {
      sourceOrderId: order.id,
      sourceOrderNumber: order.orderNumber,
      sourceChannel: order.channel,
      name: it.name,
      variant: it.variant,
      sourceCurrency: order.currency,
      unitOriginalPrice: num(it.unitPrice),
      unitCostTwd: computeUnitCostTwd(order, it),
      quantity: num(it.quantity),
    });
    ops.upserts.push(sale);
  });
  DB.sales = DB.sales.filter(s => {
    if (s.sourceOrderId === order.id && toBool(s.managedByOwnership) && !stockIds[s.sourceItemId]) {
      ops.deletes.push(s.id);
      return false;
    }
    return true;
  });
  return ops;
}
function recalcAllCosts() {
  const upserts = [];
  DB.sales.forEach(sale => {
    const order = DB.orders.find(o => o.id === sale.sourceOrderId);
    if (!order) return;
    const item = (order.items || []).find(it => it.id === sale.sourceItemId);
    if (!item) return;
    sale.unitCostTwd = computeUnitCostTwd(order, item);
    upserts.push(sale);
  });
  saveDB();
  pushOps([{ action: 'upsert', table: 'sales', rows: upserts }]);
  render();
  toast('已重算 ' + upserts.length + ' 筆成本');
}

/* ---------- 儲存操作 ---------- */
function saveOrder(order, deletedItemIds) {
  const i = DB.orders.findIndex(o => o.id === order.id);
  if (i >= 0) DB.orders[i] = order; else DB.orders.unshift(order);
  const saleOps = syncSalesForOrder(order);
  saveDB();
  const ops = [
    { action: 'upsert', table: 'orders', rows: [order] },
    { action: 'upsert', table: 'items', rows: orderItemRows(order) },
  ];
  if (deletedItemIds && deletedItemIds.length) ops.push({ action: 'delete', table: 'items', ids: deletedItemIds });
  if (saleOps.upserts.length) ops.push({ action: 'upsert', table: 'sales', rows: saleOps.upserts });
  if (saleOps.deletes.length) ops.push({ action: 'delete', table: 'sales', ids: saleOps.deletes });
  pushOps(ops);
}
function deleteOrder(order) {
  DB.orders = DB.orders.filter(o => o.id !== order.id);
  const saleIds = DB.sales.filter(s => s.sourceOrderId === order.id && toBool(s.managedByOwnership)).map(s => s.id);
  DB.sales = DB.sales.filter(s => !saleIds.includes(s.id));
  saveDB();
  const ops = [
    { action: 'delete', table: 'orders', ids: [order.id] },
    { action: 'delete', table: 'items', ids: (order.items || []).map(it => it.id) },
  ];
  if (saleIds.length) ops.push({ action: 'delete', table: 'sales', ids: saleIds });
  pushOps(ops);
}
function copyOrder(src) {
  const copy = JSON.parse(JSON.stringify(src));
  copy.id = uid();
  copy.orderNumber = '';
  copy.actualShipDate = '';
  (copy.items || []).forEach(it => {
    it.id = uid();
    it.arrived = false;
    it.sorted = false;
    it.proxyPaid = false;
    it.soldQuantity = 0;
  });
  return copy;
}
function saveSale(sale) {
  saveDB();
  const ops = [{ action: 'upsert', table: 'sales', rows: [sale] }];
  const order = DB.orders.find(o => o.id === sale.sourceOrderId);
  const item = order && (order.items || []).find(it => it.id === sale.sourceItemId);
  if (item) {
    item.salePriceTwd = sale.salePriceTwd;
    item.soldQuantity = sale.soldQuantity;
    saveDB();
    ops.push({ action: 'upsert', table: 'items', rows: [Object.assign({}, item, { orderId: order.id })] });
  }
  pushOps(ops);
}
function renumberEvents() {
  const numbered = DB.events.filter(e => num(e.eventNumber) && e.startDate);
  // 沒日期的場次：用原編號在已編號場次中的相對位置推出等效日期，維持原本的前後順序
  const effKey = e => {
    if (e.startDate) return String(e.startDate);
    const n = num(e.eventNumber);
    if (!n) return '9999';
    let best = '';
    numbered.forEach(x => { if (num(x.eventNumber) <= n && String(x.startDate) > best) best = String(x.startDate); });
    return best || '0000';
  };
  const sorted = DB.events.slice().sort((a, b) =>
    effKey(a).localeCompare(effKey(b)) || (num(a.eventNumber) - num(b.eventNumber)) ||
    String(a.createdAt || '').localeCompare(String(b.createdAt || '')));
  const changed = [];
  sorted.forEach((e, i) => {
    const n = String(i + 1);
    if (String(e.eventNumber) !== n) { e.eventNumber = n; changed.push(e); }
  });
  return changed;
}
function saveEvent(ev) {
  const i = DB.events.findIndex(e => e.id === ev.id);
  if (i >= 0) DB.events[i] = ev; else DB.events.unshift(ev);
  const changed = renumberEvents();
  if (!changed.some(e => e.id === ev.id)) changed.push(ev);
  saveDB();
  pushOps([{ action: 'upsert', table: 'events', rows: changed }]);
}
function copyEvent(src) {
  const copy = JSON.parse(JSON.stringify(src));
  copy.id = uid();
  copy.createdAt = today();
  saveEvent(copy);
  return copy;
}
function deleteEvent(ev) {
  DB.events = DB.events.filter(e => e.id !== ev.id);
  const orphans = DB.ledger.filter(l => l.eventId === ev.id);
  orphans.forEach(l => { l.eventId = ''; });
  const changed = renumberEvents();
  saveDB();
  const ops = [{ action: 'delete', table: 'events', ids: [ev.id] }];
  if (changed.length) ops.push({ action: 'upsert', table: 'events', rows: changed });
  if (orphans.length) ops.push({ action: 'upsert', table: 'ledger', rows: orphans });
  pushOps(ops);
}
// 活動的票價與座位由掛在它底下的票券流水彙整而來
function syncEventFromTickets(eventId) {
  const ev = DB.events.find(e => e.id === eventId);
  if (!ev) return null;
  const tickets = DB.ledger.filter(l => l.eventId === eventId && l.category === 'ticket');
  if (!tickets.length) return null;
  ev.ticketPriceTwd = tickets.reduce((s, l) => s + num(l.amountTwd), 0);
  const seats = tickets.map(seatText).filter(Boolean);
  if (seats.length) ev.seat = seats.join(' ');
  return ev;
}
function saveLedger(entry) {
  const i = DB.ledger.findIndex(l => l.id === entry.id);
  if (i >= 0) DB.ledger[i] = entry; else DB.ledger.unshift(entry);
  const ops = [{ action: 'upsert', table: 'ledger', rows: [entry] }];
  if (entry.eventId) {
    const ev = syncEventFromTickets(entry.eventId);
    if (ev) ops.push({ action: 'upsert', table: 'events', rows: [ev] });
  }
  saveDB();
  pushOps(ops);
}
function deleteLedger(entry) {
  DB.ledger = DB.ledger.filter(l => l.id !== entry.id);
  const ops = [{ action: 'delete', table: 'ledger', ids: [entry.id] }];
  if (entry.eventId) {
    const ev = syncEventFromTickets(entry.eventId);
    if (ev) ops.push({ action: 'upsert', table: 'events', rows: [ev] });
  }
  saveDB();
  pushOps(ops);
}

/* ---------- 總帳彙整 ---------- */
function entrySettle(l) {
  const exp = num(l.expectedReceivableTwd);
  if (exp) {
    const got = num(l.receivedTwd);
    return got >= exp ? 'settled' : got > 0 ? 'partial' : 'unsettled';
  }
  if (l.settled === undefined || l.settled === null || l.settled === '') {
    const ev = DB.events.find(e => e.id === l.eventId);
    return ev ? (ev.settled ? 'settled' : 'unsettled') : null;
  }
  return toBool(l.settled) ? 'settled' : 'unsettled';
}
function groupSettle(entries) {
  const states = entries.map(entrySettle).filter(Boolean);
  if (!states.length) return null;
  if (states.every(s => s === 'settled')) return 'settled';
  if (states.every(s => s === 'unsettled')) return 'unsettled';
  return 'partial';
}
const settleBadge = s => s === 'settled' ? '<span class="badge ok">已結清</span>'
  : s === 'partial' ? '<span class="badge warn">部分結清</span>'
  : s === 'unsettled' ? '<span class="badge danger">未結清</span>' : '';

/* ---------- 畫面 ---------- */
const state = { tab: 'orders', seg: 'list', search: '', evSearch: '', cat: '', settle: '', openGroups: {}, oSearch: '', oChannel: '', oPayer: '', oStatus: '', evSeg: 'list', statYear: '' };

function setTab(tab) {
  state.tab = tab;
  state.openGroups = {};
  render(true);
}
function render(scrollTop) {
  $('#page-title').textContent = TAB_TITLE[state.tab];
  $$('#tabbar button').forEach(b => b.classList.toggle('on', b.dataset.tab === state.tab));
  $('#fab').classList.toggle('hide', state.tab === 'sales' || state.tab === 'settings');
  const view = $('#view');
  const chipsPos = $$('.chips', view).map(c => c.scrollLeft);
  const y = window.scrollY;
  if (state.tab === 'orders') renderOrdersTab(view);
  else if (state.tab === 'sales') renderSales(view);
  else if (state.tab === 'events') renderEvents(view);
  else if (state.tab === 'ledger') renderLedger(view);
  else renderSettings(view);
  if (scrollTop === true) window.scrollTo(0, 0);
  else {
    $$('.chips', view).forEach((c, i) => { if (chipsPos[i]) c.scrollLeft = chipsPos[i]; });
    if (y) window.scrollTo(0, y);
  }
}

/* ----- 訂單頁（含待到貨、代購） ----- */
function renderOrdersTab(view) {
  const segs = [['list', '訂單'], ['arrivals', '待到貨'], ['proxy', '代購對象']];
  let html = '<div class="seg">' + segs.map(([k, l]) =>
    `<button data-seg="${k}" class="${state.seg === k ? 'on' : ''}">${l}</button>`).join('') + '</div>';
  if (state.seg === 'list') html += ordersListHtml();
  else if (state.seg === 'arrivals') html += arrivalsHtml();
  else html += proxyHtml();
  view.innerHTML = html;
  $$('[data-seg]', view).forEach(b => b.onclick = () => { state.seg = b.dataset.seg; state.openGroups = {}; render(); });
  if (state.seg === 'list') bindOrdersList(view);
  else if (state.seg === 'arrivals') bindArrivals(view);
  else bindProxy(view);
}

function orderSeqMap() {
  const m = {};
  DB.orders.slice()
    .sort((a, b) => String(a.orderDate).localeCompare(String(b.orderDate)) || String(a.id).localeCompare(String(b.id)))
    .forEach((o, i) => { m[o.id] = i + 1; });
  return m;
}
function ordersListHtml() {
  if (!DB.orders.length) return emptyHtml('尚無訂單，點右下角＋新增');
  const seq = orderSeqMap();
  const channels = [...new Set(DB.orders.map(o => o.channel).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const payers = [...new Set(DB.orders.map(o => o.payer).filter(Boolean))].sort();
  const q = state.oSearch.trim().toLowerCase();
  const statusOpts = [['', '全部狀態'], ['unshipped', '未出貨'], ['shipped', '已出貨'], ['awaiting', '有待到貨'], ['unsettled', '未結清'], ['settled', '已結清']];
  const sorted = DB.orders.filter(o => {
    if (state.oChannel && o.channel !== state.oChannel) return false;
    if (state.oPayer && o.payer !== state.oPayer) return false;
    if (state.oStatus === 'unshipped' && o.actualShipDate) return false;
    if (state.oStatus === 'shipped' && !o.actualShipDate) return false;
    if (state.oStatus === 'awaiting' && !(o.items || []).some(it => !it.arrived)) return false;
    if (state.oStatus === 'unsettled' && o.settled) return false;
    if (state.oStatus === 'settled' && !o.settled) return false;
    if (q) {
      const hay = ['#' + seq[o.id], o.orderNumber, o.channel, o.payer, o.paymentDetail, o.notes]
        .concat((o.items || []).flatMap(it => [it.name, it.variant, it.proxyFor]))
        .join(' ').toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  }).sort((a, b) => String(b.orderDate).localeCompare(String(a.orderDate)) || (seq[b.id] - seq[a.id]));
  let html = `<div class="searchbar">
    <svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="m16.5 16.5 4 4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
    <input id="o-search" placeholder="搜尋編號、通路、品名…" value="${esc(state.oSearch)}">
  </div>
  <div class="filter-row">
    <select id="o-channel" class="${state.oChannel ? 'active' : ''}"><option value="">全部通路</option>${channels.map(c => `<option ${state.oChannel === c ? 'selected' : ''} value="${esc(c)}">${esc(c)}</option>`).join('')}</select>
    <select id="o-payer" class="${state.oPayer ? 'active' : ''}"><option value="">全部付款人</option>${payers.map(p => `<option ${state.oPayer === p ? 'selected' : ''} value="${esc(p)}">${esc(p)}</option>`).join('')}</select>
    <select id="o-status" class="${state.oStatus ? 'active' : ''}">${statusOpts.map(([v, l]) => `<option ${state.oStatus === v ? 'selected' : ''} value="${v}">${l}</option>`).join('')}</select>
  </div>
  <div class="section-note">${sorted.length} 筆訂單</div>`;
  if (!sorted.length) html += emptyHtml('沒有符合條件的訂單');
  return html + sorted.map(o => {
    const pending = (o.items || []).filter(it => !it.arrived).length;
    const badges = [
      o.actualShipDate ? '<span class="badge ok">已出貨</span>' : '<span class="badge">未出貨</span>',
      pending ? `<span class="badge warn">待到貨 ${pending}</span>` : '',
      o.settled ? '' : '<span class="badge danger">未結清</span>',
    ].join(' ');
    const money = num(o.chargedTwd) ? fmtTwd(o.chargedTwd) : fmtMoney(o.currency, orderTotal(o));
    return `<div class="card tappable" data-order="${esc(o.id)}">
      <div class="row-head">
        <div class="row-title"><span style="color:var(--accent);font-family:var(--serif)">#${seq[o.id]}</span> ${esc(o.channel || '未填通路')}<span class="sub">${esc(o.orderNumber)}</span></div>
        <div class="amount accent">${money}</div>
      </div>
      <div class="row-meta">
        <span>${esc(o.orderDate || '—')}</span>
        <span>${(o.items || []).length} 項品項</span>
        ${o.estimatedShipDate ? `<span>預計 ${esc(o.estimatedShipDate)}</span>` : ''}
        ${o.payer ? `<span>付款 ${esc(o.payer)}</span>` : ''}
      </div>
      <div style="margin-top:6px;display:flex;justify-content:space-between;align-items:center;gap:8px">
        <div style="min-width:0">${badges}</div>
        <button class="icon-mini" data-copy-order="${esc(o.id)}" aria-label="複製訂單">${ICONS.copy}</button>
      </div>
    </div>`;
  }).join('');
}
function bindOrdersList(view) {
  $$('[data-order]', view).forEach(el => el.onclick = () => {
    const o = DB.orders.find(x => x.id === el.dataset.order);
    if (o) openOrderForm(o);
  });
  $$('[data-copy-order]', view).forEach(b => b.onclick = e => {
    e.stopPropagation();
    const src = DB.orders.find(x => x.id === b.dataset.copyOrder);
    if (!src) return;
    const copy = copyOrder(src);
    saveOrder(copy, []);
    render();
    toast('已複製整筆訂單，可點進去修改');
  });
  const si = $('#o-search', view);
  if (!si) return;
  si.oninput = () => { state.oSearch = si.value; render(); const el = $('#o-search'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); };
  $('#o-channel', view).onchange = e => { state.oChannel = e.target.value; render(); };
  $('#o-payer', view).onchange = e => { state.oPayer = e.target.value; render(); };
  $('#o-status', view).onchange = e => { state.oStatus = e.target.value; render(); };
}

/* ----- 待到貨 ----- */
function arrivalsHtml() {
  const rows = [];
  DB.orders.forEach(o => (o.items || []).forEach(it => {
    if (!it.arrived) rows.push({ o, it });
  }));
  if (!rows.length) return emptyHtml('所有品項都到貨了 ✦');
  const groups = {};
  rows.forEach(r => {
    const k = r.it.name || '未命名品項';
    (groups[k] = groups[k] || []).push(r);
  });
  const names = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const totalQty = rows.reduce((s, r) => s + num(r.it.quantity), 0);
  let html = `<div class="stat-strip" style="grid-template-columns:repeat(2,1fr)">
    <div class="stat"><div class="n">${names.length}</div><div class="l">待到貨品項</div></div>
    <div class="stat"><div class="n">${fmtInt(totalQty)}</div><div class="l">待到貨件數</div></div>
  </div>`;
  html += names.map(name => {
    const list = groups[name];
    const qty = list.reduce((s, r) => s + num(r.it.quantity), 0);
    const open = state.openGroups['a:' + name] ? ' open' : '';
    const body = list.map(({ o, it }) => `<div class="group-row">
        <div class="main">
          <div style="font-size:13px">${esc(it.variant || '—')} <span class="badge">${esc(OWN_LABEL[it.ownership] || it.ownership || '')}</span></div>
          <div class="row-meta"><span>${esc(o.channel)}</span><span>${esc(o.orderNumber)}</span>${o.estimatedShipDate ? `<span>預計 ${esc(o.estimatedShipDate)}</span>` : ''}</div>
        </div>
        <span style="font-size:13px;color:var(--muted)">×${fmtInt(it.quantity)}</span>
        <input type="checkbox" data-arrive-o="${esc(o.id)}" data-arrive-i="${esc(it.id)}" title="標記到貨">
      </div>`).join('');
    return `<div class="group${open}" data-group="a:${esc(name)}">
      <button class="group-head"><span class="row-title">${esc(name)}</span>
        <span style="display:flex;align-items:center;gap:8px"><span class="badge accent">共 ${fmtInt(qty)} 件</span><span class="caret">›</span></span>
      </button>
      <div class="group-body">${body}</div>
    </div>`;
  }).join('');
  return html;
}
function bindArrivals(view) {
  bindGroups(view);
  $$('[data-arrive-o]', view).forEach(cb => cb.onclick = e => {
    e.stopPropagation();
    const o = DB.orders.find(x => x.id === cb.dataset.arriveO);
    const it = o && o.items.find(x => x.id === cb.dataset.arriveI);
    if (!it) return;
    it.arrived = true;
    saveOrder(o, []);
    toast('已標記到貨：' + it.name);
    render();
  });
}

/* ----- 代購對象 ----- */
function proxyHtml() {
  const rows = [];
  DB.orders.forEach(o => (o.items || []).forEach(it => {
    if (it.ownership === 'proxy') rows.push({ o, it });
  }));
  if (!rows.length) return emptyHtml('目前沒有代購品項');
  const groups = {};
  rows.forEach(r => {
    const k = r.it.proxyFor || '未填對象';
    (groups[k] = groups[k] || []).push(r);
  });
  const people = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  let unpaidTotal = 0;
  const cards = people.map(person => {
    const list = groups[person];
    let sub = 0, unpaid = 0;
    const body = list.map(({ o, it }) => {
      const cost = computeUnitCostTwd(o, it) * num(it.quantity);
      sub += cost;
      if (!it.proxyPaid) unpaid += cost;
      return `<div class="group-row">
        <div class="main">
          <div style="font-size:13px">${esc(it.name)}${it.variant ? ` <span class="sub" style="color:var(--muted);font-size:12px">${esc(it.variant)}</span>` : ''}</div>
          <div class="row-meta"><span>${esc(o.channel)}</span><span>×${fmtInt(it.quantity)}</span><span>${cost ? '約 ' + fmtTwd(cost) : '成本待算'}</span></div>
        </div>
        <label class="check-row" style="padding:0;font-size:12px;color:var(--muted)">付清
          <input type="checkbox" data-paid-o="${esc(o.id)}" data-paid-i="${esc(it.id)}" ${it.proxyPaid ? 'checked' : ''}>
        </label>
      </div>`;
    }).join('');
    unpaidTotal += unpaid;
    const open = state.openGroups['p:' + person] ? ' open' : '';
    const badge = unpaid ? `<span class="badge danger">未收 ${fmtTwd(unpaid)}</span>` : '<span class="badge ok">已收齊</span>';
    return `<div class="group${open}" data-group="p:${esc(person)}">
      <button class="group-head"><span class="row-title">${esc(person)}<span class="sub">${list.length} 項</span></span>
        <span style="display:flex;align-items:center;gap:8px">${badge}<span class="caret">›</span></span>
      </button>
      <div class="group-body">${body}</div>
    </div>`;
  }).join('');
  return `<div class="stat-strip" style="grid-template-columns:repeat(2,1fr)">
    <div class="stat"><div class="n">${people.length}</div><div class="l">代購對象</div></div>
    <div class="stat"><div class="n">${fmtInt(unpaidTotal)}</div><div class="l">未收金額 TWD</div></div>
  </div>` + cards;
}
function bindProxy(view) {
  bindGroups(view);
  $$('[data-paid-o]', view).forEach(cb => cb.onclick = e => {
    e.stopPropagation();
    const o = DB.orders.find(x => x.id === cb.dataset.paidO);
    const it = o && o.items.find(x => x.id === cb.dataset.paidI);
    if (!it) return;
    it.proxyPaid = cb.checked;
    saveOrder(o, []);
    render();
  });
}

function bindGroups(view) {
  $$('.group-head', view).forEach(btn => btn.onclick = () => {
    const g = btn.closest('.group');
    const k = g.dataset.group;
    state.openGroups[k] = !state.openGroups[k];
    g.classList.toggle('open');
  });
}

/* ----- 販售 ----- */
function renderSales(view) {
  if (!DB.sales.length) { view.innerHTML = emptyHtml('還沒有現貨販售，把訂單品項歸屬設成「現貨」就會自動出現'); return; }
  const totalCost = DB.sales.reduce((s, x) => s + num(x.unitCostTwd) * num(x.quantity), 0);
  const soldIncome = DB.sales.reduce((s, x) => s + num(x.salePriceTwd) * num(x.soldQuantity), 0);
  const soldQty = DB.sales.reduce((s, x) => s + num(x.soldQuantity), 0);
  let html = `<div class="stat-strip">
    <div class="stat"><div class="n">${DB.sales.length}</div><div class="l">品項</div></div>
    <div class="stat"><div class="n">${fmtInt(totalCost)}</div><div class="l">總成本 TWD</div></div>
    <div class="stat"><div class="n">${fmtInt(soldIncome)}</div><div class="l">已售收入 TWD</div></div>
  </div>
  <div class="section-note" style="display:flex;justify-content:space-between;align-items:center">
    <span>已售出 ${fmtInt(soldQty)} 件</span>
    <button class="btn line small" id="recalc-btn">重算成本</button>
  </div>`;
  html += DB.sales.map(s => {
    const sold = num(s.soldQuantity), qty = num(s.quantity);
    const profit = num(s.salePriceTwd) ? (num(s.salePriceTwd) - num(s.unitCostTwd)) * sold : 0;
    return `<div class="card">
      <div class="row-head">
        <div class="row-title">${esc(s.name)}${s.variant ? `<span class="sub">${esc(s.variant)}</span>` : ''}</div>
        <span class="badge ${sold >= qty && qty ? 'ok' : ''}">${fmtInt(sold)} / ${fmtInt(qty)}</span>
      </div>
      <div class="row-meta">
        <span>${esc(s.sourceChannel)}</span>
        <span>原價 ${esc(fmtMoney(s.sourceCurrency, s.unitOriginalPrice))}</span>
        <span>成本 ${s.unitCostTwd ? fmtTwd(s.unitCostTwd) : '待算'}</span>
        ${profit ? `<span class="amount ${profit > 0 ? 'ok' : 'danger'}" style="font-size:12px">毛利 ${fmtInt(profit)}</span>` : ''}
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:10px;gap:12px">
        <div style="flex:1;display:flex;align-items:center;gap:8px;font-size:12px;color:var(--muted)">
          售價<input type="number" inputmode="numeric" style="width:90px;padding:5px 8px" value="${esc(num(s.salePriceTwd) || '')}" data-price="${esc(s.id)}" placeholder="TWD">
        </div>
        <div class="qty-step">
          <button data-sold="${esc(s.id)}" data-d="-1">−</button>
          <span class="q">${fmtInt(sold)}</span>
          <button data-sold="${esc(s.id)}" data-d="1">＋</button>
        </div>
      </div>
    </div>`;
  }).join('');
  view.innerHTML = html;
  $('#recalc-btn').onclick = recalcAllCosts;
  $$('[data-price]', view).forEach(inp => inp.onchange = () => {
    const s = DB.sales.find(x => x.id === inp.dataset.price);
    s.salePriceTwd = inp.value === '' ? '' : num(inp.value);
    saveSale(s);
    render();
  });
  $$('[data-sold]', view).forEach(btn => btn.onclick = () => {
    const s = DB.sales.find(x => x.id === btn.dataset.sold);
    const next = num(s.soldQuantity) + Number(btn.dataset.d);
    s.soldQuantity = Math.max(0, Math.min(num(s.quantity) || next, next));
    saveSale(s);
    render();
  });
}

/* ----- 活動 ----- */
function renderEvents(view) {
  const segs = [['list', '場次'], ['stats', '統計']];
  let html = '<div class="seg">' + segs.map(([k, l]) =>
    `<button data-eseg="${k}" class="${state.evSeg === k ? 'on' : ''}">${l}</button>`).join('') + '</div>';
  html += state.evSeg === 'stats' ? eventStatsHtml() : eventListHtml();
  view.innerHTML = html;
  $$('[data-eseg]', view).forEach(b => b.onclick = () => { state.evSeg = b.dataset.eseg; render(); });
  if (state.evSeg === 'stats') {
    $$('[data-year]', view).forEach(b => b.onclick = () => { state.statYear = b.dataset.year; render(); });
    return;
  }
  const si = $('#ev-search');
  si.oninput = () => { state.evSearch = si.value; render(); const el = $('#ev-search'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); };
  $$('[data-event]', view).forEach(el => el.onclick = () => {
    const ev = DB.events.find(x => x.id === el.dataset.event);
    if (ev) openEventForm(ev);
  });
  $$('[data-copy-event]', view).forEach(b => b.onclick = e => {
    e.stopPropagation();
    const src = DB.events.find(x => x.id === b.dataset.copyEvent);
    if (!src) return;
    copyEvent(src);
    render();
    toast('已複製整場活動，可點進去修改');
  });
}
function eventListHtml() {
  const q = state.evSearch.trim().toLowerCase();
  let list = DB.events.slice();
  if (q) list = list.filter(e => [e.name, e.artist, e.venue, e.city, e.liveTour, e.seriesEvent, e.seat, e.guest]
    .some(v => String(v || '').toLowerCase().includes(q)));
  const numbered = DB.events.filter(e => num(e.eventNumber) && e.startDate);
  const sortKey = e => {
    const n = num(e.eventNumber);
    if (n) return n;
    const d = String(e.startDate || '');
    if (!d) return 0;
    let best = 0;
    numbered.forEach(x => { if (String(x.startDate) <= d) best = Math.max(best, num(x.eventNumber)); });
    return best + 0.5;
  };
  list.sort((a, b) => (sortKey(b) - sortKey(a)) || String(b.startDate).localeCompare(String(a.startDate)));
  let html = `<div class="searchbar">
    <svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="m16.5 16.5 4 4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
    <input id="ev-search" placeholder="搜尋活動、歌手、場地…" value="${esc(state.evSearch)}">
  </div>
  <div class="section-note">${list.length} 場活動</div>`;
  if (!list.length) html += emptyHtml('沒有符合的活動');
  const evTitle = e => (e.eventType === '拼盤' || !e.artist || String(e.name).includes(e.artist))
    ? e.name : e.artist + ' - ' + e.name;
  html += list.map(e => `<div class="card tappable" data-event="${esc(e.id)}">
    <div class="row-head">
      <div class="row-title"><span style="color:var(--accent);font-family:var(--serif)">#${esc(e.eventNumber || '–')}</span> ${esc(evTitle(e))}</div>
    </div>
    <div class="row-meta">
      <span>${esc(e.startDate || '—')}</span>
      ${e.venue ? `<span>${esc(e.city)} ${esc(e.venue)}</span>` : ''}
      ${e.seat ? `<span>${esc(e.seat)}</span>` : ''}
    </div>
    <div style="margin-top:6px;display:flex;justify-content:space-between;align-items:center;gap:8px">
      <div style="min-width:0">
        ${e.eventType ? `<span class="badge accent">${esc(e.eventType)}</span>` : ''}
        ${e.liveTour ? `<span class="badge">${esc(e.liveTour)}</span>` : ''}
        ${e.settled ? '<span class="badge ok">已結清</span>' : '<span class="badge danger">未結清</span>'}
      </div>
      <button class="icon-mini" data-copy-event="${esc(e.id)}" aria-label="複製活動">${ICONS.copy}</button>
    </div>
  </div>`).join('');
  return html;
}
function topCount(list, keyFn) {
  const m = {};
  list.forEach(e => {
    keyFn(e).forEach(k => {
      if (!k) return;
      m[k] = m[k] || { n: 0, spend: 0 };
      m[k].n++;
      m[k].spend += num(e.ticketPriceTwd);
    });
  });
  return Object.entries(m).map(([label, v]) => ({ label, n: v.n, spend: v.spend }))
    .sort((a, b) => b.n - a.n || b.spend - a.spend);
}
function barListHtml(title, rows, opts) {
  if (!rows.length) return '';
  const max = rows[0].n;
  return `<div class="card"><div class="chart-title">${title}${opts && opts.mini ? `<span class="mini">${opts.mini}</span>` : ''}</div>` +
    rows.map(r => `<div class="bar-row">
      <span class="bl">${esc(r.label)}</span>
      <span class="track"><i style="width:${Math.max(4, Math.round(r.n / max * 100))}%"></i></span>
      <span class="bv">${fmtInt(r.n)} 場${r.spend ? ` <span class="sub2">${fmtInt(r.spend)} 元</span>` : ''}</span>
    </div>`).join('') + '</div>';
}
function colChartHtml(title, cols) {
  const max = Math.max(1, ...cols.map(c => c.n));
  return `<div class="card"><div class="chart-title">${title}</div>
    <div class="cols">${cols.map(c => `<div class="col"><em>${c.n || ''}</em><i style="height:${Math.round(c.n / max * 100)}%"></i></div>`).join('')}</div>
    <div class="cols-labels">${cols.map(c => `<span>${c.label}</span>`).join('')}</div>
  </div>`;
}
function eventStatsHtml() {
  if (!DB.events.length) return emptyHtml('還沒有活動紀錄');
  const years = [...new Set(DB.events.map(e => String(e.startDate || '').slice(0, 4)).filter(y => /^\d{4}$/.test(y)))].sort().reverse();
  const y = state.statYear;
  const evs = DB.events.filter(e => !y || String(e.startDate || '').startsWith(y));
  const priced = evs.filter(e => num(e.ticketPriceTwd) > 0);
  const spend = priced.reduce((s, e) => s + num(e.ticketPriceTwd), 0);
  const artists = new Set(evs.map(e => e.artist).filter(Boolean));
  const cities = new Set(evs.map(e => e.city).filter(Boolean));
  const most = priced.slice().sort((a, b) => num(b.ticketPriceTwd) - num(a.ticketPriceTwd))[0];
  const splitGuests = e => String(e.guest || '').split(/[、,，/／]/).map(s => s.trim()).filter(s => s && s !== '無' && s !== '-');

  let html = `<div class="chips">
    <button class="chip ${!y ? 'on' : ''}" data-year="">所有年份</button>
    ${years.map(yy => `<button class="chip ${y === yy ? 'on' : ''}" data-year="${yy}">${yy}</button>`).join('')}
  </div>
  <div class="stat-strip">
    <div class="stat"><div class="n">${evs.length}</div><div class="l">總場次</div></div>
    <div class="stat"><div class="n">${fmtInt(spend)}</div><div class="l">總花費 TWD</div></div>
    <div class="stat"><div class="n">${artists.size}</div><div class="l">不同藝人</div></div>
  </div>
  <div class="stat-strip">
    <div class="stat"><div class="n">${cities.size}</div><div class="l">跑過城市</div></div>
    <div class="stat"><div class="n">${priced.length ? fmtInt(spend / priced.length) : '—'}</div><div class="l">平均票價</div></div>
    <div class="stat"><div class="n">${most ? fmtInt(most.ticketPriceTwd) : '—'}</div><div class="l">最貴場次</div></div>
  </div>
  ${most ? `<div class="section-note" style="text-align:center">最貴的場次：${esc(most.name)}（${esc(most.startDate)}）</div>` : ''}`;

  if (!y && years.length > 1) {
    const cols = years.slice().reverse().map(yy => ({ label: yy.slice(2), n: DB.events.filter(e => String(e.startDate || '').startsWith(yy)).length }));
    html += colChartHtml('歷年場次趨勢', cols);
  }
  const months = Array.from({ length: 12 }, (_, i) => {
    const mm = String(i + 1).padStart(2, '0');
    return { label: i + 1, n: evs.filter(e => String(e.startDate || '').slice(5, 7) === mm).length };
  });
  html += colChartHtml(y ? y + ' 年月度場次' : '月度場次分布', months);
  html += barListHtml('表演者 Top 10', topCount(evs, e => [e.artist]).slice(0, 10));
  html += barListHtml('場館 Top 10', topCount(evs, e => [e.venue]).slice(0, 10));
  html += barListHtml('城市分布', topCount(evs, e => [e.city]).slice(0, 10));
  html += barListHtml('嘉賓統計 Top 10', topCount(evs, splitGuests).slice(0, 10).map(r => ({ label: r.label, n: r.n })));
  html += barListHtml('活動類型', topCount(evs, e => [e.eventType]).slice(0, 10).map(r => ({ label: r.label, n: r.n })));
  return html;
}

/* ----- 總帳 ----- */
function groupStatus(g) {
  return groupSettle(g.entries) || (g.ev ? (g.ev.settled ? 'settled' : 'unsettled') : null);
}
function renderLedger(view) {
  const catOpts = [['', '全部分類']].concat(Object.entries(CAT_LABEL));
  const setOpts = [['', '全部狀態'], ['settled', '已結清'], ['partial', '部分結清'], ['unsettled', '未結清']];
  const q = state.search.trim();
  let entries = DB.ledger.slice();
  if (q) entries = entries.filter(l => {
    const ev = DB.events.find(e => e.id === l.eventId);
    return [l.title, l.counterparty, l.payer, l.notes, l.attendee, ev && ev.name].some(v => String(v || '').toLowerCase().includes(q.toLowerCase()));
  });
  if (state.cat) entries = entries.filter(l => state.cat === 'other'
    ? !['ticket', 'transport', 'lodging'].includes(l.category)
    : l.category === state.cat);

  const incomes = entries.filter(l => l.type === 'income');
  const expenses = entries.filter(l => l.type !== 'income');
  const groups = {};
  expenses.forEach(l => { (groups[l.eventId || ''] = groups[l.eventId || ''] || []).push(l); });
  let groupList = Object.keys(groups).map(eid => {
    const ev = DB.events.find(e => e.id === eid);
    return { eid, ev, entries: groups[eid], date: ev ? ev.startDate : (groups[eid][0].date || '') };
  });
  if (state.settle) groupList = groupList.filter(g => groupStatus(g) === state.settle);
  groupList.sort((a, b) => String(b.date).localeCompare(String(a.date)));

  const totalSpend = expenses.reduce((s, l) => s + num(l.amountTwd), 0);
  const totalIncome = incomes.reduce((s, l) => s + num(l.amountTwd), 0);
  const unreceived = DB.ledger.reduce((s, l) => s + Math.max(0, num(l.expectedReceivableTwd) - num(l.receivedTwd)), 0);

  let html = `<div class="searchbar">
    <svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="m16.5 16.5 4 4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg>
    <input id="lg-search" placeholder="搜尋標題、活動、對象…" value="${esc(state.search)}">
  </div>
  <div class="filter-row two">
    <select id="lg-cat" class="${state.cat ? 'active' : ''}">${catOpts.map(([k, l]) => `<option ${state.cat === k ? 'selected' : ''} value="${k}">${l}</option>`).join('')}</select>
    <select id="lg-settle" class="${state.settle ? 'active' : ''}">${setOpts.map(([k, l]) => `<option ${state.settle === k ? 'selected' : ''} value="${k}">${l}</option>`).join('')}</select>
  </div>
  <div class="stat-strip">
    <div class="stat"><div class="n">${fmtInt(totalSpend)}</div><div class="l">支出 TWD</div></div>
    <div class="stat"><div class="n">${fmtInt(totalIncome)}</div><div class="l">收入 TWD</div></div>
    <div class="stat"><div class="n">${fmtInt(unreceived)}</div><div class="l">未收 TWD</div></div>
  </div>`;

  if (incomes.length && !state.settle) {
    const body = incomes.map(ledgerRowHtml).join('');
    const open = state.openGroups['g:income'] ? ' open' : '';
    html += `<div class="group${open}" data-group="g:income">
      <button class="group-head"><span class="row-title">收入</span>
        <span style="display:flex;align-items:center;gap:8px"><span class="amount ok">+${fmtInt(totalIncome)}</span><span class="caret">›</span></span>
      </button><div class="group-body">${body}</div></div>`;
  }
  if (!groupList.length && !incomes.length) html += emptyHtml('沒有符合的紀錄');
  html += groupList.map(g => {
    const sum = g.entries.reduce((s, l) => s + num(l.amountTwd), 0);
    const name = g.ev ? g.ev.name : '未指定活動';
    const open = state.openGroups['g:' + g.eid] ? ' open' : '';
    return `<div class="group${open}" data-group="g:${esc(g.eid)}">
      <button class="group-head">
        <span class="main" style="text-align:left;min-width:0">
          <span class="row-title" style="display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(name)}</span>
          <span class="row-meta">${esc(g.date || '')} ${settleBadge(groupStatus(g))}</span>
        </span>
        <span style="display:flex;align-items:center;gap:8px;flex:0 0 auto"><span class="amount">${fmtInt(sum)}</span><span class="caret">›</span></span>
      </button>
      <div class="group-body">${g.entries.map(ledgerRowHtml).join('')}</div>
    </div>`;
  }).join('');
  view.innerHTML = html;

  const si = $('#lg-search');
  si.oninput = () => { state.search = si.value; render(); const el = $('#lg-search'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); };
  $('#lg-cat', view).onchange = e => { state.cat = e.target.value; render(); };
  $('#lg-settle', view).onchange = e => { state.settle = e.target.value; render(); };
  bindGroups(view);
  $$('[data-ledger]', view).forEach(el => el.onclick = e => {
    e.stopPropagation();
    const l = DB.ledger.find(x => x.id === el.dataset.ledger);
    if (l) openLedgerForm(l);
  });
}
function ledgerRowHtml(l) {
  const st = entrySettle(l);
  const seat = seatText(l);
  return `<div class="group-row tappable" data-ledger="${esc(l.id)}">
    <div class="main">
      <div style="font-size:13.5px">${esc(l.title || CAT_LABEL[l.category] || '未命名')}
        <span class="badge">${esc(CAT_LABEL[l.category] || l.category || '')}</span> ${settleBadge(st)}
      </div>
      <div class="row-meta">
        <span>${esc(l.date || '')}</span>
        ${l.payer ? `<span>付款 ${esc(l.payer)}</span>` : ''}
        ${seat ? `<span>${esc(seat)}</span>` : ''}
        ${l.attendee ? `<span>${esc(l.attendee)}</span>` : ''}
        ${num(l.expectedReceivableTwd) ? `<span>應結 ${fmtInt(l.expectedReceivableTwd)}／已結 ${fmtInt(l.receivedTwd)}${l.counterparty ? '（' + esc(l.counterparty) + '）' : ''}</span>` : ''}
      </div>
    </div>
    <span class="amount ${l.type === 'income' ? 'ok' : ''}">${l.type === 'income' ? '+' : ''}${fmtInt(l.amountTwd)}</span>
  </div>`;
}

/* ----- 設定 ----- */
function renderSettings(view) {
  const c = DB;
  const itemCount = c.orders.reduce((s, o) => s + (o.items || []).length, 0);
  const swatches = THEMES.map(t => `<button class="swatch ${CFG.theme === t.id ? 'on' : ''}" data-theme-pick="${t.id}">
      <div class="dot" style="background:linear-gradient(135deg, ${t.bg} 50%, ${t.accent} 50%)"></div>${t.label}
    </button>`).join('');
  view.innerHTML = `
  <div class="form-section" style="margin-top:6px">主題色</div>
  <div class="swatches">${swatches}</div>

  <div class="form-section">雲端同步</div>
  <div class="field"><label>APPS SCRIPT 網址</label><input id="cfg-api" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(CFG.apiUrl)}"></div>
  <div class="field"><label>金鑰（選填，需與 Code.gs 的 SHARED_KEY 一致）</label><input id="cfg-key" value="${esc(CFG.key)}"></div>
  <div class="field"><label>GOOGLE 試算表網址（選填，方便快速開啟）</label><input id="cfg-sheet" placeholder="https://docs.google.com/spreadsheets/d/…" value="${esc(CFG.sheetUrl)}"></div>
  <div class="btn-row">
    <button class="btn line" id="cfg-save">儲存並測試連線</button>
    <button class="btn primary" id="cfg-sync">立即同步</button>
  </div>
  <div class="section-note">
    ${CFG.lastSync ? '上次同步：' + esc(CFG.lastSync) : '尚未同步過'}
    ${isDirty() ? '<br><span style="color:var(--danger)">⚠ 本機有變更尚未上傳，請按「全部上傳」</span>' : ''}
    ${CFG.sheetUrl ? `<br><a href="${esc(CFG.sheetUrl)}" target="_blank" rel="noopener">開啟 Google 試算表 ↗</a>` : ''}
  </div>

  <div class="form-section">資料</div>
  <button class="btn line" id="btn-upload-all">全部上傳到試算表</button>
  <button class="btn line" id="btn-export">匯出 JSON 備份</button>
  <button class="btn line" id="btn-import">匯入 JSON 備份</button>
  <input type="file" id="import-file" accept=".json,application/json" style="display:none">
  <div class="section-note">目前資料：訂單 ${c.orders.length}｜品項 ${itemCount}｜販售 ${c.sales.length}｜活動 ${c.events.length}｜流水 ${c.ledger.length}</div>
  <div class="divider"></div>
  <div class="section-note" style="text-align:center">STAR LEDGER ✦ 追星總帳<br>資料儲存於你的 Google 試算表</div>`;

  $$('[data-theme-pick]', view).forEach(b => b.onclick = () => {
    CFG.theme = b.dataset.themePick;
    saveCfg();
    applyTheme();
    render();
  });
  $('#cfg-save').onclick = async () => {
    CFG.apiUrl = $('#cfg-api').value.trim();
    CFG.key = $('#cfg-key').value.trim();
    CFG.sheetUrl = $('#cfg-sheet').value.trim();
    saveCfg();
    if (!CFG.apiUrl) { toast('已儲存（未設定 API 網址）'); return; }
    toast('測試連線中…');
    try {
      await apiGet();
      toast('連線成功 ✦');
    } catch (err) {
      toast('連線失敗：' + err.message);
    }
    render();
  };
  $('#cfg-sync').onclick = () => doSync(false);
  $('#btn-upload-all').onclick = () => {
    if (confirm('會以本機資料覆蓋試算表全部內容，確定上傳？')) uploadAll();
  };
  $('#btn-export').onclick = () => {
    const blob = new Blob([JSON.stringify(Object.assign({ app: 'Fan Ledger', version: 7 }, DB), null, 1)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'star-ledger-backup-' + today() + '.json';
    a.click();
    URL.revokeObjectURL(a.href);
  };
  $('#btn-import').onclick = () => $('#import-file').click();
  $('#import-file').onchange = async e => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (!data.orders && !data.events && !data.ledger) throw new Error('不是有效的備份檔');
      if (!confirm(`備份含 訂單 ${(data.orders || []).length}｜販售 ${(data.sales || []).length}｜活動 ${(data.events || []).length}｜流水 ${(data.ledger || []).length}，將取代目前所有資料，繼續？`)) return;
      DB = {
        orders: (data.orders || []).map(o => Object.assign({ items: [] }, o)),
        sales: data.sales || [],
        events: data.events || [],
        ledger: data.ledger || [],
      };
      saveDB();
      markDirty();
      render();
      toast('匯入完成');
      if (CFG.apiUrl && confirm('要立刻把匯入的資料上傳到試算表嗎？')) uploadAll();
    } catch (err) {
      toast('匯入失敗：' + err.message);
    }
    e.target.value = '';
  };
}

/* ---------- 底部彈窗 ---------- */
function openSheet(html) {
  const sheet = $('#sheet');
  sheet.innerHTML = '<div class="sheet-handle"></div>' + html;
  sheet.classList.add('show');
  $('#sheet-backdrop').classList.add('show');
  sheet.scrollTop = 0;
  return sheet;
}
function closeSheet() {
  $('#sheet').classList.remove('show');
  $('#sheet-backdrop').classList.remove('show');
}
$('#sheet-backdrop').onclick = closeSheet;

function emptyHtml(msg) {
  return `<div class="empty"><div class="glyph">✦</div>${esc(msg)}</div>`;
}
let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
}

/* ---------- 訂單表單 ---------- */
function fieldHtml(label, inner) {
  return `<div class="field"><label>${label}</label>${inner}</div>`;
}
function selectHtml(id, options, value) {
  return `<select id="${id}">` + options.map(([v, l]) =>
    `<option value="${esc(v)}" ${String(value) === String(v) ? 'selected' : ''}>${esc(l)}</option>`).join('') + '</select>';
}
function datalistOptions(values) {
  return [...new Set(values.filter(Boolean))].map(v => `<option value="${esc(v)}">`).join('');
}

function openOrderForm(existing) {
  const isNew = !existing;
  const o = existing ? JSON.parse(JSON.stringify(existing)) : {
    id: uid(), orderNumber: '', channel: '', orderDate: today(), estimatedShipDate: '', actualShipDate: '',
    currency: 'KRW', domesticShipping: '', internationalShippingTwd: '', internationalShippingRateTwdPerKg: '',
    discountAmount: '', weightGrams: '', exchangeRate: '', chargedTwd: '', payer: '', paymentMethod: '',
    paymentDetail: '', settled: false, notes: '', items: [],
  };
  if (!o.items.length) o.items.push(newItem());
  const originalItemIds = existing ? existing.items.map(it => it.id) : [];

  function newItem() {
    return { id: uid(), name: '', variant: '', unitPrice: '', quantity: 1, ownership: 'self', proxyFor: '', arrived: false, sorted: false, proxyPaid: false, salePriceTwd: '', soldQuantity: 0 };
  }
  function itemBlock(it, i) {
    const ownOpts = Object.entries(OWN_LABEL).map(([v, l]) => [v, l]);
    return `<div class="item-block" data-item="${i}">
      <div class="item-block-head"><span class="no">品項 ${i + 1}</span>
        <span style="display:flex;gap:8px">
          <button class="icon-mini" data-copy-item="${i}" aria-label="複製品項">${ICONS.copy}</button>
          <button class="icon-mini danger" data-del-item="${i}" aria-label="移除品項">${ICONS.trash}</button>
        </span></div>
      <div class="field"><label>品名</label><input data-k="name" data-i="${i}" value="${esc(it.name)}"></div>
      <div class="field-row">
        <div class="field"><label>版本／規格</label><input data-k="variant" data-i="${i}" value="${esc(it.variant)}"></div>
        <div class="field" style="flex:0 0 90px"><label>數量</label><input type="number" inputmode="numeric" data-k="quantity" data-i="${i}" value="${esc(it.quantity)}"></div>
      </div>
      <div class="field-row">
        <div class="field"><label>單價（原幣）</label><input type="number" inputmode="decimal" data-k="unitPrice" data-i="${i}" value="${esc(it.unitPrice)}"></div>
        <div class="field"><label>歸屬</label><select data-k="ownership" data-i="${i}">${ownOpts.map(([v, l]) => `<option value="${v}" ${it.ownership === v ? 'selected' : ''}>${l}</option>`).join('')}</select></div>
      </div>
      <div class="field" data-proxy-wrap="${i}" style="${it.ownership === 'proxy' ? '' : 'display:none'}">
        <label>代購對象</label><input data-k="proxyFor" data-i="${i}" list="dl-proxy" value="${esc(it.proxyFor)}">
      </div>
      <div style="display:flex;gap:18px;padding-bottom:8px">
        <label class="check-row">已到貨 <input type="checkbox" data-k="arrived" data-i="${i}" ${it.arrived ? 'checked' : ''}></label>
        <label class="check-row">已分貨 <input type="checkbox" data-k="sorted" data-i="${i}" ${it.sorted ? 'checked' : ''}></label>
        <label class="check-row" data-paid-wrap="${i}" style="${it.ownership === 'proxy' ? '' : 'display:none'}">已付清 <input type="checkbox" data-k="proxyPaid" data-i="${i}" ${it.proxyPaid ? 'checked' : ''}></label>
      </div>
    </div>`;
  }

  const payOpts = [['', '—'], ['cash', '現金'], ['credit_card', '信用卡'], ['bank_transfer', '轉帳'], ['mobile_payment', '行動支付']];
  const curOpts = CURRENCIES.map(c => [c, c]);
  const html = `
  ${sheetTitleHtml(isNew ? '新增訂單' : '編輯訂單', !isNew, 'order-del')}
  <datalist id="dl-channel">${datalistOptions(DB.orders.map(x => x.channel))}</datalist>
  <datalist id="dl-payer">${datalistOptions(DB.orders.map(x => x.payer).concat(DB.ledger.map(x => x.payer)))}</datalist>
  <datalist id="dl-proxy">${datalistOptions(DB.orders.flatMap(x => (x.items || []).map(it => it.proxyFor)))}</datalist>
  <div class="field-row">
    ${fieldHtml('通路', `<input id="f-channel" list="dl-channel" value="${esc(o.channel)}">`)}
    ${fieldHtml('訂單編號', `<input id="f-orderNumber" value="${esc(o.orderNumber)}">`)}
  </div>
  <div class="field-row">
    ${fieldHtml('訂購日期', `<input id="f-orderDate" type="date" value="${esc(o.orderDate)}">`)}
    ${fieldHtml('幣別', selectHtml('f-currency', curOpts, o.currency))}
  </div>
  <div class="field-row">
    ${fieldHtml('預計出貨', `<input id="f-estimatedShipDate" placeholder="2026-07-02 或區間" value="${esc(o.estimatedShipDate)}">`)}
    ${fieldHtml('實際出貨', `<input id="f-actualShipDate" type="date" value="${esc(o.actualShipDate)}">`)}
  </div>

  <div class="form-section">品項 <button class="btn line small" id="add-item">＋ 加一項</button></div>
  <div id="items-wrap">${o.items.map(itemBlock).join('')}</div>

  <div class="form-section">金額</div>
  <div class="field-row">
    ${fieldHtml('商城／國內運費', `<input id="f-domesticShipping" type="number" inputmode="decimal" value="${esc(o.domesticShipping)}">`)}
    ${fieldHtml('折抵（點數/購物金）', `<input id="f-discountAmount" type="number" inputmode="decimal" value="${esc(o.discountAmount)}">`)}
  </div>
  <div class="field-row">
    ${fieldHtml('實刷台幣', `<input id="f-chargedTwd" type="number" inputmode="decimal" value="${esc(o.chargedTwd)}">`)}
    ${fieldHtml('匯率', `<input id="f-exchangeRate" type="number" step="any" inputmode="decimal" value="${esc(o.exchangeRate)}">`)}
  </div>
  <div class="hint" id="rate-hint"></div>
  <div class="field-row">
    ${fieldHtml('重量（公克）', `<input id="f-weightGrams" type="number" inputmode="numeric" value="${esc(o.weightGrams)}">`)}
    ${fieldHtml('集運費率 TWD/KG', `<input id="f-intlRate" type="number" inputmode="decimal" value="${esc(o.internationalShippingRateTwdPerKg)}">`)}
  </div>
  ${fieldHtml('國際運費（台幣）', `<input id="f-internationalShippingTwd" type="number" inputmode="decimal" value="${esc(o.internationalShippingTwd)}">`)}
  <div class="hint" id="intl-hint"></div>
  <div class="total-box" id="total-box"></div>

  <div class="form-section">付款</div>
  <div class="field-row">
    ${fieldHtml('付款人', `<input id="f-payer" list="dl-payer" value="${esc(o.payer)}">`)}
    ${fieldHtml('付款方式', selectHtml('f-paymentMethod', payOpts, o.paymentMethod))}
  </div>
  ${fieldHtml('付款工具（卡片／平台）', `<input id="f-paymentDetail" placeholder="例：永豐、LINE Pay" value="${esc(o.paymentDetail)}">`)}
  <label class="check-row">已結清 <input type="checkbox" id="f-settled" ${o.settled ? 'checked' : ''}></label>
  ${fieldHtml('備註', `<textarea id="f-notes" rows="2">${esc(o.notes)}</textarea>`)}

  <div class="sheet-actions">
    <button class="btn primary" id="order-save">儲存訂單</button>
  </div>`;

  const sheet = openSheet(html);
  $('#sh-close').onclick = closeSheet;

  function readItems() {
    $$('#items-wrap [data-k]').forEach(inp => {
      const it = o.items[Number(inp.dataset.i)];
      if (!it) return;
      const k = inp.dataset.k;
      if (inp.type === 'checkbox') it[k] = inp.checked;
      else if (k === 'unitPrice' || k === 'quantity') it[k] = inp.value === '' ? '' : num(inp.value);
      else it[k] = inp.value;
    });
  }
  function readOrder() {
    readItems();
    ['channel', 'orderNumber', 'orderDate', 'estimatedShipDate', 'actualShipDate', 'payer', 'paymentDetail', 'notes'].forEach(k => { o[k] = $('#f-' + k).value.trim(); });
    o.currency = $('#f-currency').value;
    o.paymentMethod = $('#f-paymentMethod').value;
    o.settled = $('#f-settled').checked;
    ['domesticShipping', 'discountAmount', 'chargedTwd', 'exchangeRate', 'weightGrams', 'internationalShippingTwd'].forEach(k => {
      const v = $('#f-' + k).value;
      o[k] = v === '' ? '' : num(v);
    });
    const ir = $('#f-intlRate').value;
    o.internationalShippingRateTwdPerKg = ir === '' ? '' : num(ir);
  }
  function refreshTotals() {
    readOrder();
    const t = itemsTotal(o), total = orderTotal(o);
    const rows = [
      ['品項合計', fmtMoney(o.currency, t)],
      num(o.domesticShipping) ? ['＋ 國內運費', fmtMoney(o.currency, o.domesticShipping)] : null,
      num(o.discountAmount) ? ['− 折抵', fmtMoney(o.currency, o.discountAmount)] : null,
      ['訂單總額', fmtMoney(o.currency, total)],
      num(o.chargedTwd) ? ['實刷台幣', fmtTwd(o.chargedTwd)] : null,
      num(o.internationalShippingTwd) ? ['國際運費', fmtTwd(o.internationalShippingTwd)] : null,
    ].filter(Boolean);
    $('#total-box').innerHTML = rows.map(([l, v], idx) =>
      `<div class="tr ${idx === rows.length - 1 && rows.length > 1 ? '' : ''}${l === '訂單總額' ? ' big' : ''}"><span>${l}</span><span class="amount">${v}</span></div>`).join('');
    const back = backCalcRate(o);
    $('#rate-hint').textContent = rateLabel(o) + (back && !num(o.exchangeRate) ? '（依實刷回算：' + (o.currency === 'USD' ? 'USD 1 = TWD ' + back.toFixed(3) : 'TWD 1 = ' + o.currency + ' ' + back.toFixed(3)) + '，儲存時自動帶入）' : '');
    const est = num(o.weightGrams) && num(o.internationalShippingRateTwdPerKg)
      ? Math.round(num(o.weightGrams) / 1000 * num(o.internationalShippingRateTwdPerKg)) : 0;
    $('#intl-hint').textContent = est ? `依重量估算：約 ${fmtTwd(est)}${num(o.internationalShippingTwd) ? '' : '（儲存時自動帶入）'}` : '';
  }
  function rebindItems() {
    $$('#items-wrap [data-copy-item]').forEach(b => b.onclick = () => {
      readItems();
      const i = Number(b.dataset.copyItem);
      const copy = Object.assign({}, o.items[i], { id: uid() });
      o.items.splice(i + 1, 0, copy);
      $('#items-wrap').innerHTML = o.items.map(itemBlock).join('');
      rebindItems();
      refreshTotals();
    });
    $$('#items-wrap [data-del-item]').forEach(b => b.onclick = () => {
      readItems();
      o.items.splice(Number(b.dataset.delItem), 1);
      if (!o.items.length) o.items.push(newItem());
      $('#items-wrap').innerHTML = o.items.map(itemBlock).join('');
      rebindItems();
      refreshTotals();
    });
    $$('#items-wrap select[data-k="ownership"]').forEach(sel => sel.onchange = () => {
      const i = sel.dataset.i;
      $(`[data-proxy-wrap="${i}"]`).style.display = sel.value === 'proxy' ? '' : 'none';
      $(`[data-paid-wrap="${i}"]`).style.display = sel.value === 'proxy' ? '' : 'none';
    });
  }
  rebindItems();
  refreshTotals();
  sheet.oninput = refreshTotals;

  $('#add-item').onclick = () => {
    readItems();
    o.items.push(newItem());
    $('#items-wrap').innerHTML = o.items.map(itemBlock).join('');
    rebindItems();
    refreshTotals();
  };
  $('#order-save').onclick = () => {
    readOrder();
    if (!num(o.exchangeRate)) {
      const back = backCalcRate(o);
      if (back) o.exchangeRate = Number(back.toFixed(4));
    }
    if (!num(o.internationalShippingTwd) && num(o.weightGrams) && num(o.internationalShippingRateTwdPerKg)) {
      o.internationalShippingTwd = Math.round(num(o.weightGrams) / 1000 * num(o.internationalShippingRateTwdPerKg));
    }
    const deleted = originalItemIds.filter(id => !o.items.some(it => it.id === id));
    saveOrder(o, deleted);
    closeSheet();
    render();
    toast('訂單已儲存');
  };
  if (!isNew) {
    $('#order-del').onclick = () => {
      if (!confirm('刪除這筆訂單與所有品項？相關現貨販售紀錄也會移除。')) return;
      deleteOrder(o);
      closeSheet();
      render();
      toast('訂單已刪除');
    };
  }
}

/* ---------- 活動表單 ---------- */
function openEventForm(existing) {
  const isNew = !existing;
  const ev = existing ? JSON.parse(JSON.stringify(existing)) : {
    id: uid(), name: '', artist: '', city: '', venue: '', startDate: today(), endDate: '',
    eventNumber: '', originalDate: '', eventType: '', liveTour: '', seriesEvent: '',
    seat: '', ticketPriceTwd: '', guest: '', payer: '', settled: false, notes: '', createdAt: today(),
  };
  const related = DB.ledger.filter(l => l.eventId === ev.id);
  const relatedHtml = related.length ? related.map(l => `
    <div class="group-row tappable" data-rel-ledger="${esc(l.id)}" style="border:1px solid var(--line);border-radius:12px;margin-bottom:8px">
      <div class="main">
        <div style="font-size:13px">${esc(l.title || CAT_LABEL[l.category] || '')} <span class="badge">${esc(CAT_LABEL[l.category] || '')}</span> ${settleBadge(entrySettle(l))}</div>
        <div class="row-meta"><span>${esc(l.date)}</span>${l.payer ? `<span>${esc(l.payer)}</span>` : ''}${seatText(l) ? `<span>${esc(seatText(l))}</span>` : ''}</div>
      </div>
      <span class="amount">${fmtInt(l.amountTwd)}</span>
    </div>`).join('') : '<div class="section-note">尚無相關流水</div>';

  const html = `
  ${sheetTitleHtml(isNew ? '新增活動' : '編輯活動 #' + esc(ev.eventNumber || '–'), !isNew, 'event-del')}
  <datalist id="dl-artist">${datalistOptions(DB.events.map(x => x.artist))}</datalist>
  <datalist id="dl-type">${datalistOptions(DB.events.map(x => x.eventType).concat(['專場', '演唱會', '簽售', '見面會', '快閃店']))}</datalist>
  <datalist id="dl-payer2">${datalistOptions(DB.events.map(x => x.payer).concat(DB.ledger.map(x => x.payer)))}</datalist>
  ${fieldHtml('活動名稱', `<input id="e-name" value="${esc(ev.name)}">`)}
  ${fieldHtml('表演者', `<input id="e-artist" list="dl-artist" value="${esc(ev.artist)}">`)}
  ${fieldHtml('活動日期', `<input id="e-startDate" type="date" value="${esc(ev.startDate)}">`)}
  <div class="hint">場次編號會依活動日期自動編排</div>
  <div class="field-row">
    ${fieldHtml('城市', `<input id="e-city" value="${esc(ev.city)}">`)}
    ${fieldHtml('場地', `<input id="e-venue" value="${esc(ev.venue)}">`)}
  </div>
  <div class="field-row">
    ${fieldHtml('活動類型', `<input id="e-eventType" list="dl-type" value="${esc(ev.eventType)}">`)}
    ${fieldHtml('LIVE TOUR', `<input id="e-liveTour" value="${esc(ev.liveTour)}">`)}
  </div>
  ${fieldHtml('系列場次', `<input id="e-seriesEvent" placeholder="例：NO.68" value="${esc(ev.seriesEvent)}">`)}
  ${num(ev.ticketPriceTwd) || ev.seat ? `<div class="hint">座位與票價由下方「相關流水」的票券彙整：${ev.seat ? esc(ev.seat) + '｜' : ''}${num(ev.ticketPriceTwd) ? fmtTwd(ev.ticketPriceTwd) : ''}</div>` : ''}
  <div class="field-row">
    ${fieldHtml('嘉賓', `<input id="e-guest" value="${esc(ev.guest)}">`)}
    ${fieldHtml('付款人', `<input id="e-payer" list="dl-payer2" value="${esc(ev.payer)}">`)}
  </div>
  <label class="check-row">已結清 <input type="checkbox" id="e-settled" ${ev.settled ? 'checked' : ''}></label>
  ${fieldHtml('備註', `<textarea id="e-notes" rows="2">${esc(ev.notes)}</textarea>`)}
  ${isNew ? '' : `
  <div class="form-section">相關流水
    <span>
      <button class="btn line small" data-add-cat="ticket">＋票券</button>
      <button class="btn line small" data-add-cat="transport">＋交通</button>
      <button class="btn line small" data-add-cat="lodging">＋住宿</button>
    </span>
  </div>
  <div class="hint">新增後會同步寫入追星總帳；票券金額與座位會回寫此活動。</div>
  ${relatedHtml}`}
  <div class="sheet-actions">
    <button class="btn primary" id="event-save">儲存活動</button>
  </div>`;

  openSheet(html);
  $('#sh-close').onclick = closeSheet;

  function readEvent() {
    ['name', 'artist', 'startDate', 'city', 'venue', 'eventType', 'liveTour', 'seriesEvent', 'guest', 'payer', 'notes'].forEach(k => { ev[k] = $('#e-' + k).value.trim(); });
    ev.settled = $('#e-settled').checked;
  }
  $('#event-save').onclick = () => {
    readEvent();
    saveEvent(ev);
    closeSheet();
    render();
    toast('活動已儲存');
  };
  if (!isNew) {
    $('#event-del').onclick = () => {
      if (!confirm('刪除這場活動？相關總帳流水會保留但解除連結。')) return;
      deleteEvent(ev);
      closeSheet();
      render();
      toast('活動已刪除');
    };
    $$('[data-add-cat]').forEach(b => b.onclick = () => {
      readEvent();
      saveEvent(ev);
      openLedgerForm(null, {
        category: b.dataset.addCat, eventId: ev.id, date: ev.startDate,
        title: ev.name + '｜' + CAT_LABEL[b.dataset.addCat], payer: ev.payer,
      }, ev.id);
    });
    $$('[data-rel-ledger]').forEach(el => el.onclick = () => {
      const l = DB.ledger.find(x => x.id === el.dataset.relLedger);
      if (l) openLedgerForm(l, null, ev.id);
    });
  }
}

/* ---------- 總帳表單 ---------- */
function openLedgerForm(existing, preset, backEventId) {
  const isNew = !existing;
  const l = existing ? JSON.parse(JSON.stringify(existing)) : Object.assign({
    id: uid(), type: 'expense', category: 'ticket', date: today(), title: '', eventId: '',
    amountTwd: '', currency: '', originalAmount: '', exchangeRate: '', payer: '', paymentMethod: '',
    paymentDetail: '', counterparty: '', expectedReceivableTwd: '', receivedTwd: '', settled: false, notes: '',
    ticketType: '', ticketArea: '', ticketRow: '', ticketSeat: '', attendee: '', ticketStatus: '', createdAt: today(),
  }, preset || {});
  const settledChecked = (l.settled === undefined || l.settled === null || l.settled === '')
    ? !!(DB.events.find(e => e.id === l.eventId) || {}).settled
    : toBool(l.settled);

  const evOpts = [['', '不指定活動']].concat(
    DB.events.slice().sort((a, b) => String(b.startDate).localeCompare(String(a.startDate)))
      .map(e => [e.id, `${e.startDate || ''} ${e.name}`]));
  const catOpts = Object.entries(CAT_LABEL);
  const payOpts = [['', '—'], ['cash', '現金'], ['credit_card', '信用卡'], ['bank_transfer', '轉帳'], ['mobile_payment', '行動支付']];
  const curOpts = [['', '—']].concat(CURRENCIES.map(c => [c, c]));

  const html = `
  ${sheetTitleHtml(isNew ? '新增流水' : '編輯流水', !isNew, 'ledger-del', !!backEventId)}
  <div class="seg" id="l-type">
    <button data-type="expense" class="${l.type !== 'income' ? 'on' : ''}">支出</button>
    <button data-type="income" class="${l.type === 'income' ? 'on' : ''}">收入</button>
  </div>
  <div class="field-row">
    ${fieldHtml('分類', selectHtml('l-category', catOpts, l.category))}
    ${fieldHtml('日期', `<input id="l-date" type="date" value="${esc(l.date)}">`)}
  </div>
  ${fieldHtml('標題', `<input id="l-title" value="${esc(l.title)}">`)}
  <div class="field" id="l-event-wrap" style="${l.type === 'income' ? 'display:none' : ''}">
    <label>活動場次</label>${selectHtml('l-eventId', evOpts, l.eventId)}
  </div>
  ${fieldHtml('金額（台幣）', `<input id="l-amountTwd" type="number" inputmode="numeric" value="${esc(l.amountTwd)}">`)}
  <div class="field-row">
    ${fieldHtml('原幣', selectHtml('l-currency', curOpts, l.currency))}
    ${fieldHtml('原幣金額', `<input id="l-originalAmount" type="number" inputmode="decimal" value="${esc(l.originalAmount)}">`)}
    ${fieldHtml('匯率', `<input id="l-exchangeRate" type="number" step="any" inputmode="decimal" value="${esc(l.exchangeRate)}">`)}
  </div>
  <div class="field-row">
    ${fieldHtml('付款人', `<input id="l-payer" list="dl-payer3" value="${esc(l.payer)}">`)}
    ${fieldHtml('付款方式', selectHtml('l-paymentMethod', payOpts, l.paymentMethod))}
  </div>
  <datalist id="dl-payer3">${datalistOptions(DB.ledger.map(x => x.payer).concat(DB.orders.map(x => x.payer)))}</datalist>
  ${fieldHtml('付款工具', `<input id="l-paymentDetail" value="${esc(l.paymentDetail)}">`)}
  <label class="check-row">已結清 <input type="checkbox" id="l-settled" ${settledChecked ? 'checked' : ''}></label>

  <div class="form-section">分帳</div>
  ${fieldHtml('對象', `<input id="l-counterparty" value="${esc(l.counterparty)}">`)}
  <div class="field-row">
    ${fieldHtml('應結金額', `<input id="l-expectedReceivableTwd" type="number" inputmode="numeric" value="${esc(l.expectedReceivableTwd)}">`)}
    ${fieldHtml('已結金額', `<input id="l-receivedTwd" type="number" inputmode="numeric" value="${esc(l.receivedTwd)}">`)}
  </div>

  <div id="ticket-wrap" style="${l.category === 'ticket' ? '' : 'display:none'}">
    <div class="form-section">票券資訊</div>
    <div class="field-row">
      ${fieldHtml('票種', `<input id="l-ticketType" placeholder="例：全席指定" value="${esc(l.ticketType)}">`)}
      ${fieldHtml('入場人', `<input id="l-attendee" value="${esc(l.attendee)}">`)}
    </div>
    <div class="field-row">
      ${fieldHtml('區域', `<input id="l-ticketArea" value="${esc(l.ticketArea)}">`)}
      ${fieldHtml('排', `<input id="l-ticketRow" value="${esc(l.ticketRow)}">`)}
      ${fieldHtml('座號', `<input id="l-ticketSeat" value="${esc(l.ticketSeat)}">`)}
    </div>
    ${fieldHtml('票券狀態', `<input id="l-ticketStatus" list="dl-tstatus" value="${esc(l.ticketStatus)}"><datalist id="dl-tstatus"><option value="已使用"><option value="未使用"><option value="轉讓"><option value="退票"></datalist>`)}
  </div>
  ${fieldHtml('備註', `<textarea id="l-notes" rows="2">${esc(l.notes)}</textarea>`)}
  <div class="sheet-actions">
    <button class="btn primary" id="ledger-save">儲存流水</button>
  </div>`;

  openSheet(html);
  $('#sh-close').onclick = closeSheet;
  const goBack = () => {
    const ev = DB.events.find(e => e.id === backEventId);
    if (ev) openEventForm(ev); else closeSheet();
  };
  if (backEventId) $('#sh-back').onclick = goBack;

  $$('#l-type button').forEach(b => b.onclick = () => {
    l.type = b.dataset.type;
    $$('#l-type button').forEach(x => x.classList.toggle('on', x === b));
    $('#l-event-wrap').style.display = l.type === 'income' ? 'none' : '';
  });
  $('#l-category').onchange = () => {
    $('#ticket-wrap').style.display = $('#l-category').value === 'ticket' ? '' : 'none';
  };
  $('#ledger-save').onclick = () => {
    ['date', 'title', 'payer', 'paymentDetail', 'counterparty', 'notes', 'ticketType', 'ticketArea', 'ticketRow', 'ticketSeat', 'attendee', 'ticketStatus'].forEach(k => { l[k] = $('#l-' + k).value.trim(); });
    l.category = $('#l-category').value;
    l.eventId = l.type === 'income' ? '' : $('#l-eventId').value;
    l.paymentMethod = $('#l-paymentMethod').value;
    l.currency = $('#l-currency').value;
    l.settled = $('#l-settled').checked;
    ['amountTwd', 'originalAmount', 'exchangeRate', 'expectedReceivableTwd', 'receivedTwd'].forEach(k => {
      const v = $('#l-' + k).value;
      l[k] = v === '' ? '' : num(v);
    });
    saveLedger(l);
    render();
    toast('流水已儲存');
    if (backEventId) goBack(); else closeSheet();
  };
  if (!isNew) {
    $('#ledger-del').onclick = () => {
      if (!confirm('刪除這筆流水？')) return;
      deleteLedger(l);
      render();
      toast('流水已刪除');
      if (backEventId) goBack(); else closeSheet();
    };
  }
}

/* ---------- 啟動 ---------- */
$$('#tabbar button').forEach(b => b.onclick = () => setTab(b.dataset.tab));
$('#fab').onclick = () => {
  if (state.tab === 'orders') openOrderForm(null);
  else if (state.tab === 'events') openEventForm(null);
  else if (state.tab === 'ledger') openLedgerForm(null);
};
$('#sync-btn').onclick = () => doSync(false);
applyTheme();
render();
if (CFG.apiUrl && !isDirty()) doSync(true);
