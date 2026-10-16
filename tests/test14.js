// ===== 球賽 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-16', '10:00');
let m;
assert(/球賽/.test(vm.runInContext('SCAN_PROMPT', ctx)) && /A vs B/.test(vm.runInContext('SCAN_PROMPT', ctx)));
const base = { docType: 'order', currency: 'TWD', platform: 'ibon', ticketCount: 2, seats: [{ area: 'L4內野一般區410區', row: '3', seat: '1' }, { area: 'L4內野一般區410區', row: '3', seat: '2' }], unitFace: 1000, totalPaid: 2000, payMethod: 'credit_card', payDetail: '永豐', pickupMethod: '電子票' };
// 1. AI 讀出球賽 → 不再問類型，卡片顯示球賽
aiByImage.b1 = Object.assign({}, base, { eventType: '球賽', artist: '統一7-ELEVEn獅 vs 樂天桃猿', eventName: '中華職棒 例行賽', date: '2026-11-30', time: '16:05', city: '台北', venue: '臺北大巨蛋', orderNumber: 'B-1' });
m = img(A, 'b1'); assert(/統一7-ELEVEn獅 vs 樂天桃猿｜中華職棒 例行賽/.test(flat(m)), flat(m));
m = pb(A, 'a=purpose&v=self');
for (let i = 0; i < 10 && !/請確認/.test(flat(m)); i++) { assert(!/專場、拼盤，還是球賽/.test(flat(m)), '讀到球賽就不再問'); const q = m.flat().find(x => x.quickReply); m = pb(A, q.quickReply.items[0].action.data); }
assert(/球賽｜統一7-ELEVEn獅 vs 樂天桃猿/.test(flat(m)));
m = pb(A, 'a=card&v=confirm'); assert(/已建檔/.test(flat(m)));
const ev = table('events').find(e => e.artist === '統一7-ELEVEn獅 vs 樂天桃猿');
assert(ev && ev.eventType === '球賽' && ev.name === '中華職棒 例行賽');
// 2. 看不出類型 → 按鈕有球賽、打字也認得
aiByImage.b2 = Object.assign({}, base, { eventType: '', artist: '中信兄弟 vs 富邦悍將', eventName: '中信兄弟 vs 富邦悍將', date: '2026-12-01', city: '台北', venue: '天母棒球場', orderNumber: 'B-2' });
m = img(A, 'b2'); m = pb(A, 'a=purpose&v=self');
for (let i = 0; i < 10 && !/專場、拼盤，還是球賽/.test(flat(m)); i++) { const q = m.flat().find(x => x.quickReply); m = pb(A, q.quickReply.items[0].action.data); }
assert(/專場\|拼盤\|球賽/.test(flat(m)));
m = text(A, '棒球比賽'); assert(!/專場、拼盤，還是球賽/.test(flat(m)), '打字認得球賽');
assert.strictEqual(JSON.parse(cache['ls_' + A]).d.eventType, '球賽');
console.log('\nALL PASSED');
