// ===== 讀圖中：不同時跑兩次、可以取消 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-17', '10:00');
let m;
const now = () => new Date(NOW_ISO).getTime();
aiByImage.k1 = { docType: 'order', eventType: '專場', artist: 'DAY6', eventName: 'FOREVER', date: '2026-12-05', currency: 'TWD', ticketCount: 1, totalPaid: 4580, seats: [{ area: 'A' }], orderNumber: 'K-1' };
// 1. 正在讀時再傳圖 → 提醒還在讀＋取消鈕，不會開第二次
const before = geminiCalls.length;
cache['lread_' + A] = JSON.stringify({ id: 'r1', at: now() - 40000, imgs: ['k0'] });
m = img(A, 'k1'); log('還在讀', m);
assert(/上一張還在讀（已經 40 秒）/.test(flat(m)) && /\{取消這次\}/.test(flat(m)) && geminiCalls.length === before);
m = img(A, 'k1'); assert.strictEqual(m.flat().length, 0, '連續傳只提醒一次');
delete cache['lbusymsg_' + A];
// 2. 讀太久按再試一次也一樣
cache['lretry_' + A] = JSON.stringify(['k1']);
m = pb(A, 'a=retry&v=1'); assert(/上一張還在讀/.test(flat(m)) && geminiCalls.length === before);
delete cache['lbusymsg_' + A];
// 3. 按取消 → 那一次之後讀完也不回，可以再試一次
m = pb(A, 'a=rcancel&v=1'); log('取消', m);
assert(/已取消/.test(flat(m)) && /\{再試一次\}/.test(flat(m)) && cache['lcx_r1'] && !cache['lread_' + A]);
assert.deepStrictEqual(JSON.parse(cache['lretry_' + A]), ['k0']);
assert.strictEqual(vm.runInContext("readEnd_('Uaaa', 'r1')", ctx), false, '被取消的那次讀完不回覆');
// 4. 取消後重新傳 → 正常讀
m = img(A, 'k1'); assert(/讀到了/.test(flat(m)));
assert(!cache['lread_' + A], '讀完清掉狀態');
m = text(A, '取消'); assert(/已取消這筆/.test(flat(m)));
// 5. 打「取消」也能取消讀圖
cache['lread_' + A] = JSON.stringify({ id: 'r2', at: now() - 5000, imgs: ['k1'] });
m = text(A, '取消'); assert(/已取消，這次讀到的結果不會記/.test(flat(m)));
// 6. 卡太久（超過 150 秒，例如被 Google 中斷）不會一直擋住
cache['lread_' + A] = JSON.stringify({ id: 'r3', at: now() - 200000, imgs: ['k1'] });
m = img(A, 'k1'); assert(/讀到了/.test(flat(m)));
console.log('\nALL PASSED');
