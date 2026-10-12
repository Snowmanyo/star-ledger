// ===== 補登、備份 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-07', '10:00');
let m;
const H = n => sheets[n].rows[0];
const addRow = (n, o) => sheets[n].rows.push(H(n).map(h => o[h] ?? ''));
addRow('orders', { id: 'o1', orderNumber: 'W-100', channel: 'Weverse Shop', orderDate: '2026-10-05', currency: 'KRW', domesticShipping: '3000', discountAmount: '2000', internationalShippingTwd: '0', notes: '預計出貨：11/2～11/9\n國際運費尚未確認' });
addRow('items', { id: 'i1', orderId: 'o1', name: '專輯', variant: 'A', unitPrice: '20000', quantity: '3', ownership: 'self', arrived: 'false' });
addRow('items', { id: 'i2', orderId: 'o1', name: '手燈', unitPrice: '45000', quantity: '1', ownership: 'stock', arrived: 'false' });
addRow('sales', { id: 's1', sourceOrderId: 'o1', sourceItemId: 'i2', name: '手燈', unitCostTwd: '1000', quantity: '1', managedByOwnership: 'true' });
addRow('orders', { id: 'o2', orderNumber: 'J-1', channel: 'JYP JAPAN', orderDate: '2026-10-07', currency: 'JPY', chargedTwd: '1000', notes: '國際運費尚未確認' });
addRow('items', { id: 'i3', orderId: 'o2', name: 'T-shirt', unitPrice: '4600', quantity: '1', ownership: 'self' });
addRow('orders', { id: 'o3', orderNumber: 'TP-9', channel: '快閃店', orderDate: '2026-10-06', currency: 'TWD' });
const ord = id => table('orders').find(x => x.id === id);
// 1. 補登清單
m = text(A, '補登'); log('補登清單', m);
let f = flat(m);
assert(/📝 還沒補完的周邊訂單/.test(f) && /缺：實刷台幣、國際運費/.test(f) && /JYP JAPAN｜J-1/.test(f) && /缺：國際運費/.test(f) && !/TP-9/.test(f));
assert(f.indexOf('J-1') < f.indexOf('W-100'), '新的在前');
// 2. 點一筆補實刷
m = pb(A, JSON.stringify(m).match(/a=ofill&v=o1/)[0]); assert(/補登周邊訂單/.test(flat(m)) && /還缺：實刷台幣、國際運費/.test(flat(m)) && /重算現貨成本/.test(JSON.stringify(m)));
m = text(A, '實刷 2,580'); log('補實刷', m);
assert.strictEqual(ord('o1').chargedTwd, '2580'); assert.strictEqual(ord('o1').exchangeRate, String(Number((106000 / 2580).toFixed(4))));
assert(/TWD 1 = KRW 41\.0853/.test(flat(m)) && /還缺：國際運費/.test(flat(m)) && /按「重算現貨成本」/.test(flat(m)));
assert.strictEqual(table('sales')[0].unitCostTwd, '1000', '沒按就不改成本');
// 3. 重量＋費率
m = text(A, '重量 1.2kg 費率 120'); assert.deepStrictEqual([ord('o1').weightGrams, ord('o1').internationalShippingRateTwdPerKg, ord('o1').internationalShippingTwd], ['1200', '120', '144']);
assert(/都補完了/.test(flat(m))); assert.strictEqual(ord('o1').notes, '預計出貨：11/2～11/9', '拿掉尚未確認，其他備註保留');
// 4. 重算現貨成本
m = pb(A, 'a=ocost&v=o1'); log('重算成本', m);
assert.strictEqual(table('sales')[0].unitCostTwd, String(Math.round((2580 + 144) * 45000 / 105000)));
assert(/已重算現貨成本/.test(flat(m)));
m = pb(A, 'a=xend&v=1');
// 5. 一句話補登＋出貨
m = text(A, 'j-1 國際運費 150 出貨 10/20'); log('一句話', m);
assert.deepStrictEqual([ord('o2').internationalShippingTwd, ord('o2').notes, ord('o2').actualShipDate], ['150', '', '2025-10-20'], '未來日期當成去年');
m = text(A, '出貨了'); assert.strictEqual(ord('o2').actualShipDate, '2026-10-07', '修改中可以繼續補');
m = text(A, '你好嗎'); assert(/看不懂要補什麼/.test(flat(m)));
m = text(A, '取消'); assert(/結束修改/.test(flat(m)));
m = text(A, '補登'); assert(/都補完了/.test(flat(m)));
// 6. 不是訂單編號的話照常回答
m = text(A, 'ABC 實刷 100'); assert(!/補登/.test(flat(m)));
// 7. 手動備份、只留 8 份
m = text(A, '備份'); log('備份', m); assert(/已備份 ✦ 2026-10-07/.test(flat(m)));
const folder = drive.folders['追星總帳備份']; assert.strictEqual(folder.files.length, 1); assert.strictEqual(folder.files[0].name, '追星總帳 備份 2026-10-07');
for (let i = 0; i < 9; i++) text(A, '備份');
assert.strictEqual(folder.files.filter(x => !x.trashed).length, 8); assert.strictEqual(folder.files.filter(x => x.trashed).length, 2, '舊的進垃圾桶');
assert(folder.files[0].trashed && folder.files[1].trashed, '丟掉的是最舊的');
// 8. 每週日 03:00 自動備份一次
const before = folder.files.length;
setNow('2026-10-10', '03:05'); vm.runInContext('reminderTick()', ctx); assert.strictEqual(folder.files.length, before, '週六不備份');
setNow('2026-10-11', '02:55'); vm.runInContext('reminderTick()', ctx); assert.strictEqual(folder.files.length, before);
setNow('2026-10-11', '03:00'); vm.runInContext('reminderTick()', ctx); assert.strictEqual(folder.files.length, before + 1);
setNow('2026-10-11', '03:05'); vm.runInContext('reminderTick()', ctx); assert.strictEqual(folder.files.length, before + 1, '同一天只備份一次');
// 9. 錯誤訊息翻成中文
const fe = e => vm.runInContext('friendlyError_(' + JSON.stringify(e) + ')', ctx);
assert(/重新授權/.test(fe('DriveApp.getFoldersByName を呼び出す権限がありません。')));
assert(/執行太久/.test(fe('Exceeded maximum execution time')));
assert.strictEqual(fe('讀不到 LINE 上的圖片（代碼 404）'), '讀不到 LINE 上的圖片（代碼 404）');
const realBackup = ctx.DriveApp.getFoldersByName; ctx.DriveApp.getFoldersByName = () => { throw new Error('DriveApp.getFoldersByName を呼び出す権限がありません。'); };
m = text(A, '備份'); assert(/備份失敗：程式還沒有被允許/.test(flat(m))); ctx.DriveApp.getFoldersByName = realBackup;
console.log('\nALL PASSED');
