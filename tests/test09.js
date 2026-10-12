// ===== AI 鬼打牆、回覆失敗 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
let m;
const realFetch = ctx.UrlFetchApp.fetch;
const models = [];
let mode = 'trunc-then-ok';
ctx.UrlFetchApp.fetch = function (url, o) {
  if (url.includes('generativelanguage')) {
    const model = url.match(/models\/([^:]+)/)[1]; models.push(model);
    const cfg = JSON.parse(o.payload).generationConfig;
    assert(cfg.maxOutputTokens > 0 && cfg.temperature === undefined, '有上限、不設溫度');
    const R = body => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify(body) });
    if (mode === 'trunc-then-ok' && model === 'gemini-3.8-flash') return R({ candidates: [{ finishReason: 'MAX_TOKENS', content: { parts: [{ text: '{"eventName":"N.Flying","seats":[{"area":"C1"},{"area":"C1"},{"are' }] } }] });
    if (mode === 'all-trunc') return R({ candidates: [{ finishReason: 'MAX_TOKENS', content: { parts: [{ text: '{"eventName":' }] } }] });
    return R({ candidates: [{ finishReason: 'STOP', content: { parts: [{ text: JSON.stringify({ eventName: "&CON5: into REM'", artist: 'N.Flying', date: '2026-09-20', time: '16:00', city: '新北', venue: '新北工商展覽中心', currency: 'TWD', ticketCount: 2, totalPaid: 7600, unitFace: 3800, seats: [{ area: 'C1區', row: '34', seat: '77' }, { area: 'C1區', row: '34', seat: '79' }], platform: 'ibon', docType: 'order' }) }] } }] });
  }
  return realFetch.apply(this, arguments);
};
// 1. 第一個模型輸出被截斷 → 換備用模型讀成功
m = img(A, 'ibon1'); log('截斷後換模型', m);
assert.deepStrictEqual(models, ['gemini-3.8-flash', 'gemini-3.5-flash-lite']);
assert(/N\.Flying/.test(flat(m)) && /2 張/.test(flat(m)));
m = text(A, '取消');
// 2. 兩個模型都壞 → 清楚的錯誤訊息，不會卡住
mode = 'all-trunc'; models.length = 0;
m = img(A, 'ibon2'); log('都壞', m);
assert(/AI 這次讀不出來/.test(flat(m)));
// 3. 回覆失敗（權杖過期）→ 改用推播
mode = 'ok';
const realReply = ctx.UrlFetchApp.fetch;
ctx.UrlFetchApp.fetch = function (url, o) {
  if (url.includes('/message/reply')) return { getResponseCode: () => 400, getContentText: () => '{"message":"Invalid reply token"}' };
  return realReply.apply(this, arguments);
};
pushes.length = 0;
text(A, '說明');
assert(pushes.length === 1 && pushes[0].to[0] === A && /記票/.test(JSON.stringify(pushes[0])), '改用推播送出');
console.log('\nALL PASSED');
