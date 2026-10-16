// ===== Google 忙線（503）=====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-16', '23:00');
let m;
const realFetch = ctx.UrlFetchApp.fetch;
let busy = 99; const calls = [];
ctx.UrlFetchApp.fetch = function (url, o) {
  if (url.includes('generativelanguage') && busy > 0) {
    busy--; calls.push(url.match(/models\/([^:]+)/)[1]);
    return { getResponseCode: () => 503, getContentText: () => '{"error":{"code":503,"message":"This model is currently experiencing high demand.","status":"UNAVAILABLE"}}' };
  }
  return realFetch.apply(this, arguments);
};
aiByImage.g1 = { docType: 'order', eventType: '球賽', artist: '台灣大賽味全龍 vs 中信兄弟', eventName: '中華職棒37年台灣大賽', date: '2026-10-21', time: '18:35', city: '台北', venue: '臺北大巨蛋', currency: 'TWD', ticketCount: 2, totalPaid: 1600, unitFace: 800, seats: [{ area: 'L4內野410區', row: '4', seat: '11' }, { area: 'L4內野410區', row: '4', seat: '12' }], platform: '拓元', orderNumber: 'CT2610160001977' };
// 1. 都忙線 → 每個模型重試一次、白話訊息、有「再試一次」
m = img(A, 'g1'); log('忙線', m);
assert.strictEqual(calls.length, 4, '兩個模型各試兩次'); 
assert(/Google 那邊忙線中/.test(flat(m)) && !/UNAVAILABLE|error/.test(flat(m)) && /\{再試一次\}/.test(flat(m)));
assert(cache['gslow_gemini-3.8-flash'] && cache['gslow_gemini-3.5-flash-lite'], '忙線的模型先排後面');
// 2. 忙線過了，按「再試一次」→ 不用重傳就讀到
busy = 0;
m = pb(A, 'a=retry&v=1'); log('再試一次', m);
assert(/讀到了/.test(flat(m)) && /台灣大賽味全龍 vs 中信兄弟/.test(flat(m)));
// 3. 第一次忙線、重試成功
m = text(A, '取消'); delete cache['gslow_gemini-3.8-flash']; delete cache['gslow_gemini-3.5-flash-lite'];
busy = 1; calls.length = 0;
m = img(A, 'g1'); assert(/讀到了/.test(flat(m)) && calls.length === 1, '重試一次就好，不用換模型');
m = text(A, '取消');
// 4. 放太久
m = pb(A, 'a=retry&v=1'); assert(/放太久了/.test(flat(m)));
console.log('\nALL PASSED');
