// ===== 周邊訂單 =====
const A = 'Uaaa', B = 'Ubbb';
text(A, 'pw123'); text(A, 'Chi'); text(B, 'pw123'); text(B, 'Jhen');
setNow('2026-10-07', '10:00');
let m;
// 既有訂單：通路寫法與重複
const H = n => sheets[n].rows[0];
sheets.orders.rows.push(H('orders').map(h => ({ id: 'o0', orderNumber: 'OLD-1', channel: 'Weverse Shop', orderDate: '2026-09-01', payer: 'Chi', currency: 'KRW' })[h] ?? ''));
sheets.items.rows.push(H('items').map(h => ({ id: 'i0', orderId: 'o0', name: 'DAY6 專輯', variant: 'A版', unitPrice: '20000', quantity: '1', ownership: 'proxy', proxyFor: '小美', arrived: 'false' })[h] ?? ''));
const W = { docType: 'merch', channel: 'WEVERSE SHOP', orderNumber: 'W-100', orderDate: '2026-10-05', currency: 'KRW',
  items: [{ name: 'DAY6 專輯', variant: 'A版', unitPrice: 20000, quantity: 2 }, { name: 'DAY6 專輯', variant: 'B版', unitPrice: 20000, quantity: 1 }, { name: '手燈', variant: '', unitPrice: 45000, quantity: 1 }],
  domesticShipping: 3000, discountAmount: 2000, totalPaid: 106000, shipFrom: '2026-11-02', shipText: '11/2～11/9 依序出貨', payMethod: 'credit_card', payDetail: '',
  seats: [], ticketCount: 0, eventName: '', date: '' };
aiByImage.w1 = W;
// 1. 自動判斷周邊、通路沿用舊寫法
m = img(A, 'w1'); log('讀周邊', m);
assert(/周邊訂單｜Weverse Shop W-100/.test(flat(m)) && /KRW 106,000｜3 項/.test(flat(m)), '讀到通路與合計');
assert(/全部自留\|全部代購\|全部現貨\|每項分開選/.test(flat(m)));
assert(/周邊通路：Weverse Shop/.test(geminiPrompts.at(-1)), 'AI 有拿到習慣的通路寫法');
// 2. 全部代購 → 給誰
m = pb(A, 'a=mown&v=proxy'); assert(/這筆代購給誰/.test(flat(m)) && /小美/.test(flat(m)), '代購對象選項含以前用過的');
m = pb(A, 'a=mproxy&v=Jhen'); assert(/哪張卡/.test(flat(m)), '付款方式讀到了就跳過');
m = pb(A, 'a=mdetail&v=永豐'); assert(/付款人/.test(flat(m))); m = text(A, 'Chi');
assert(/實際刷了多少台幣/.test(flat(m))); m = text(A, '2580');
log('確認卡', m);
const f = flat(m);
assert(/請確認：Weverse Shop 訂單/.test(f) && /代購・Jhen/.test(f) && /＋運費 KRW 3,000/.test(f) && /－折抵 KRW 2,000/.test(f) && /合計 KRW 106,000/.test(f) && /實刷 TWD 2,580/.test(f));
assert(/11\/2～11\/9 依序出貨/.test(f) && /國際運費尚未確認/.test(f) && !/⚠/.test(f));
// 3. 打字修改
editQueue.push({ understood: true, items: [{ name: 'DAY6 專輯', variant: 'A版', unitPrice: 20000, quantity: 2, ownership: 'proxy', proxyFor: 'Jhen' }, { name: 'DAY6 專輯', variant: 'B版', unitPrice: 20000, quantity: 1, ownership: 'proxy', proxyFor: 'Jhen' }, { name: '手燈', variant: '', unitPrice: 45000, quantity: 1, ownership: 'stock', proxyFor: '' }] });
m = text(A, '第3項改成現貨'); assert(/已修改/.test(flat(m)) && /現貨/.test(flat(m)));
editQueue.push({ understood: true, items: [{ name: 'DAY6 專輯', variant: 'A版', unitPrice: 20000, quantity: 3, ownership: 'proxy', proxyFor: 'Jhen' }, { name: 'DAY6 專輯', variant: 'B版', unitPrice: 20000, quantity: 1, ownership: 'proxy', proxyFor: 'Jhen' }, { name: '手燈', variant: '', unitPrice: 45000, quantity: 1, ownership: 'stock', proxyFor: '' }] });
m = text(A, '第1項數量改成3'); assert(/⚠ 截圖上的總額是 KRW 106,000/.test(flat(m)), '總額不一致提醒');
editQueue.push({ understood: true, items: [{ name: 'DAY6 專輯', variant: 'A版', unitPrice: 20000, quantity: 2, ownership: 'proxy', proxyFor: 'Jhen' }, { name: 'DAY6 專輯', variant: 'B版', unitPrice: 20000, quantity: 1, ownership: 'proxy', proxyFor: 'Jhen' }, { name: '手燈', variant: '', unitPrice: 45000, quantity: 1, ownership: 'stock', proxyFor: '' }] });
m = text(A, '第1項數量改回2'); assert(!/⚠/.test(flat(m)));
const itemIdsBefore = JSON.parse(cache['ls_' + A]).m.items.map(i => i.id);
// 4. 確認建檔
m = pb(A, 'a=mcard&v=confirm'); log('建檔', m);
assert(/已建檔 ✦ Weverse Shop 訂單（3 項、4 件）/.test(flat(m)) && /現貨 1 項已自動加到「販售」/.test(flat(m)));
const o = table('orders').find(x => x.orderNumber === 'W-100');
assert.deepStrictEqual([o.channel, o.orderDate, o.currency, o.domesticShipping, o.discountAmount, o.chargedTwd, o.estimatedShipDate, o.internationalShippingTwd, o.payer, o.paymentMethod, o.paymentDetail, o.settled],
  ['Weverse Shop', '2026-10-05', 'KRW', '3000', '2000', '2580', '2026-11-02', '0', 'Chi', 'credit_card', '永豐', 'false']);
assert.strictEqual(o.exchangeRate, String(Number((106000 / 2580).toFixed(4))), '匯率依實刷回算');
assert(/預計出貨：11\/2～11\/9 依序出貨/.test(o.notes) && /國際運費尚未確認/.test(o.notes));
const its = table('items').filter(x => x.orderId === o.id);
assert.strictEqual(its.length, 3); assert.deepStrictEqual(its.map(x => x.id), itemIdsBefore, '修改後品項 id 保持');
assert.deepStrictEqual(its.map(x => [x.ownership, x.proxyFor, x.arrived, x.proxyPaid]), [['proxy', 'Jhen', 'false', 'false'], ['proxy', 'Jhen', 'false', 'false'], ['stock', '', 'false', 'false']]);
assert(!its.some(x => /運費/.test(x.name)), '運費不是品項');
const sale = table('sales').find(x => x.sourceOrderId === o.id);
assert(sale && sale.sourceItemId === its[2].id && sale.managedByOwnership === 'true' && sale.name === '手燈');
assert.strictEqual(sale.unitCostTwd, String(Math.round(2580 * 45000 / 105000)), '成本依實刷按原價比例分攤');
// 5. 重複訂單
m = img(A, 'w1'); assert(/好像已經記過了/.test(flat(m))); m = pb(A, 'a=mdup&v=cancel'); assert(/已取消/.test(flat(m)));
// 6. 每項分開選、多張截圖合併、待補、台幣訂單
aiByImage.p1 = { docType: 'merch', channel: '台北快閃店', orderNumber: 'TP-9', orderDate: '2026-10-06', currency: 'TWD', items: [{ name: '扇子', variant: 'Wonpil', unitPrice: 450, quantity: 1 }, { name: '貼紙', variant: '', unitPrice: 150, quantity: 2 }], domesticShipping: 0, discountAmount: 0, totalPaid: 0, shipFrom: '', shipText: '', payMethod: '', payDetail: '', seats: [], ticketCount: 0 };
aiByImage.p2 = { docType: 'merch', channel: '台北快閃店', orderNumber: 'TP-9', currency: 'TWD', items: [{ name: '扇子', variant: 'Wonpil', unitPrice: 450, quantity: 1 }, { name: '徽章', variant: 'Young K', unitPrice: 300, quantity: 1 }], domesticShipping: 80, totalPaid: 1130, seats: [], ticketCount: 0 };
m = img(A, 'p1'); assert(/TWD 750｜2 項/.test(flat(m)));
m = img(A, 'p2'); log('補傳合併', m);
assert(/已合併/.test(flat(m)) && /台北快閃店 TP-9/.test(flat(m)) && /TWD 1,130｜3 項/.test(flat(m)) && /全部自留/.test(flat(m)), '補傳同一張訂單：合併、重複品項不重算、運費補上');
m = pb(A, 'a=mown&v=each'); assert(/第 1 \/ 3 項：扇子｜Wonpil ×1/.test(flat(m)));
m = pb(A, 'a=mitem&v=self'); assert(/第 2 \/ 3 項：貼紙 ×2/.test(flat(m)));
m = text(A, '代購'); assert(/第 2 項代購給誰/.test(flat(m))); m = text(A, '阿明');
m = pb(A, 'a=mitem&v=pending'); assert(/付款方式/.test(flat(m)));
m = pb(A, 'a=mown&v=self'); assert(/付款方式/.test(flat(m)), '按到舊按鈕 → 重問目前這題');
m = text(A, '現金'); m = pb(A, 'a=mdetail&v=__skip'); m = pb(A, 'a=mpayer&v=Jhen');
assert(/請確認/.test(flat(m)) && !/實際刷了/.test(flat(m)), '台幣不問實刷'); assert(!/國際運費/.test(flat(m)));
m = pb(A, 'a=mcard&v=confirm');
const o2 = table('orders').find(x => x.orderNumber === 'TP-9');
assert.deepStrictEqual(table('items').filter(x => x.orderId === o2.id).map(x => [x.name, x.ownership, x.proxyFor]), [['扇子', 'self', ''], ['貼紙', 'proxy', '阿明'], ['徽章', 'pending', '']]);
assert.deepStrictEqual([o2.internationalShippingTwd, o2.exchangeRate, o2.notes, o2.estimatedShipDate], ['', '', '', '']);
// 7. 外幣跳過實刷 → 估算匯率
aiByImage.j1 = { docType: 'merch', channel: 'JYP JAPAN', orderNumber: 'J-1', orderDate: '2026-10-07', currency: 'JPY', items: [{ name: 'T-shirt', variant: 'L', unitPrice: 4600, quantity: 1 }], totalPaid: 4600, seats: [], ticketCount: 0 };
m = img(B, 'j1'); m = pb(B, 'a=mown&v=self'); m = pb(B, 'a=mpay&v=credit_card'); m = pb(B, 'a=mdetail&v=__skip'); m = pb(B, 'a=mpayer&v=Jhen');
m = pb(B, 'a=mcharged&v=__skip'); assert(/約 TWD 1,000（估算）/.test(flat(m)));
m = pb(B, 'a=mcard&v=confirm'); assert(/記得到網站補「實刷台幣」/.test(flat(m)));
assert.strictEqual(table('orders').find(x => x.orderNumber === 'J-1').exchangeRate, '4.6');
// 8. 取消
m = img(B, 'j1'); assert(/已經記過/.test(flat(m))); m = text(B, '取消'); assert(/已取消這筆/.test(flat(m)));
// 9. 待到貨：依品名彙總、點一個標記到貨
m = text(A, '待到貨'); log('待到貨', m);
let fa = flat(m);
assert(/📦 待到貨/.test(fa) && /DAY6 專輯/.test(fa) && /共 4 件/.test(fa), '同品名不同版本合併（舊 1＋新 3）');
assert(/A版｜×2｜代購・Jhen｜Weverse Shop/.test(fa) && /T-shirt/.test(fa));
const tapA = JSON.stringify(m).match(/a=marr&v=([^"]+)/g).find(x => x.includes(its[0].id) || true);
const albumTap = JSON.stringify(m).match(/"data":"(a=marr&v=[^"]+)","displayText":"到貨：DAY6 專輯"/)[1];
m = pb(A, albumTap); assert(/「DAY6 專輯」哪些到貨了/.test(flat(m)) && /全部到貨（3 筆）/.test(flat(m)));
m = pb(A, qrData(m, 'B版')); assert(/已標記到貨 ✦ DAY6 專輯｜B版 ×1/.test(flat(m)));
assert.strictEqual(table('items').find(x => x.id === its[1].id).arrived, 'true');
m = pb(A, albumTap); m = pb(A, qrData(m, '全部到貨')); assert(/A版 ×2/.test(flat(m)) && /A版 ×1/.test(flat(m)));
m = text(A, '待到貨'); assert(!/DAY6 專輯/.test(flat(m)) && /手燈/.test(flat(m)));
// 10. 記票仍正常、周邊不影響轉賣／未來場次
const base = { eventType: '專場', currency: 'TWD', pickupMethod: '電子票', payMethod: 'credit_card', payDetail: '永豐' };
aiByImage.t1 = Object.assign({}, base, { docType: 'order', eventName: 'TOUR', artist: 'ITZY', date: '2026-12-27', time: '19:00', city: '台北', venue: '北流', platform: 'KKTIX', orderNumber: 'R-1', ticketCount: 1, seats: [{ area: 'A' }], unitFace: 4380, totalPaid: 4380 });
m = img(A, 't1'); assert(/這張是自己要去的/.test(flat(m)), '購票截圖照常走記票');
m = text(A, '取消');
// 11. 周邊進行中補傳不同訂單 → 詢問；當成新的一筆
m = img(A, 'w1'); m = pb(A, 'a=mdup&v=go'); m = img(A, 'j1'); assert(/看起來不太一樣/.test(flat(m)));
m = pb(A, 'a=supp&v=new'); assert(/JYP JAPAN/.test(flat(m)) && /已經記過/.test(flat(m)));
m = text(A, '取消');
// 12. 選單上傳截圖
m = text(A, '上傳截圖'); assert.deepStrictEqual(m.flat()[0].quickReply.items.map(i => i.action.type), ['cameraRoll', 'camera']);
vm.runInContext('setupBot()', ctx); const areas = JSON.parse(lineMenus.at(-1)).areas;
assert.strictEqual(areas.length, 6); assert.strictEqual(areas.reduce((t, a) => t + a.bounds.width * a.bounds.height, 0), 2500 * 1686); assert.deepStrictEqual([areas[0].action.type, areas[0].action.uri], ['uri', 'https://line.me/R/nv/cameraRoll/multi']);
console.log('\nALL PASSED');
