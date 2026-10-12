// ===== 演出日期讀成訂購日：卡片警告 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-12', '10:00');
let m;
assert(vm.runInContext('SCAN_PROMPT', ctx).includes('不是訂購時間'), '讀圖指示有說明演出日期');
const finish = (uid, msgs) => { for (let i = 0; i < 12 && !/請確認/.test(flat(msgs)); i++) { const q = msgs.flat().find(x => x.quickReply); msgs = pb(uid, q.quickReply.items[0].action.data); } return msgs; };
const base = { eventType: '專場', currency: 'TWD', pickupMethod: '超商取票', payMethod: 'credit_card', payDetail: '永豐', docType: 'order', platform: '拓元', ticketCount: 1, seats: [{ area: '3樓5880 (VIP 2)', row: '22', seat: '44' }], unitFace: 5880, unitFee: 200, totalPaid: 6080 };
aiByImage.y1 = Object.assign({}, base, { eventName: 'Solo Tour <YOUNGEST>', artist: 'Young K', date: '2026-09-01', time: '15:24', orderNumber: '42736345' });
m = img(A, 'y1'); m = pb(A, 'a=purpose&v=resale'); m = finish(A, m);
assert(/⚠ 演出日期已經過了/.test(flat(m)), '過去日期有警告');
m = text(A, '取消');
aiByImage.y2 = Object.assign({}, base, { eventName: 'Solo Tour <YOUNGEST>', artist: 'Young K', date: '2026-11-28', time: '19:00', orderNumber: '42736346' });
m = img(A, 'y2'); m = pb(A, 'a=purpose&v=resale'); m = finish(A, m);
assert(!/⚠ 演出日期已經過了/.test(flat(m)), '未來日期沒有警告');
console.log('\nALL PASSED');
