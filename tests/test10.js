// ===== 轉賣進度：打字、聊天截圖 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-10', '10:00');
let m;
const finish = (uid, msgs) => { for (let i = 0; i < 12 && !/請確認/.test(flat(msgs)); i++) { const q = msgs.flat().find(x => x.quickReply); msgs = pb(uid, q.quickReply.items[0].action.data); } return msgs; };
const base = { eventType: '專場', currency: 'TWD', pickupMethod: '電子票', payMethod: 'credit_card', payDetail: '永豐', docType: 'order' };
aiByImage.r1 = Object.assign({}, base, { eventName: 'FOREVER YOUNG', artist: 'DAY6', date: '2026-12-05', time: '19:00', city: '台北', venue: '北流', platform: 'KKTIX', orderNumber: 'R-1', ticketCount: 1, seats: [{ area: 'A區', row: '5', seat: '18' }], unitFace: 4380, unitFee: 200, totalPaid: 4580 });
aiByImage.r2 = Object.assign({}, base, { eventName: 'TUNNEL VISION', artist: 'ITZY', date: '2026-12-20', time: '19:00', city: '台北', venue: '北流', platform: '拓元', orderNumber: 'R-2', ticketCount: 2, seats: [{ area: '橙207區' }, { area: '橙207區' }], unitFace: 3000, totalPaid: 6000 });
for (const id of ['r1', 'r2']) { m = img(A, id); m = pb(A, 'a=purpose&v=resale'); m = finish(A, m); m = pb(A, 'a=card&v=confirm'); }
const TR = () => table('transfers');
const day6 = TR().find(t => /DAY6|FOREVER/.test(t.title + t.artist)).id, itzy = TR().find(t => /ITZY|TUNNEL/.test(t.title + t.artist)).id;
// 1. 打字：一句話 → 確認卡片 → 更新
editQueue.push({ intent: 'resaleUpdate' });
editQueue.push({ understood: true, ids: [day6], person: '小美', buyerContact: 'LINE abc123', amountTwd: 4580, receivedTwd: 1000, delivered: false, summary: '小美要買，談好 4580，已收訂金 1000' });
m = text(A, 'DAY6 12/5 那張賣給小美 4580，收訂金1000，她 LINE abc123'); log('打字', m);
let f = flat(m);
assert(/確認轉賣進度/.test(f) && /FOREVER YOUNG/.test(f) && /買家\n— → 小美/.test(f) && /已收\n— → 1,000/.test(f) && /更新後：🟡 收訂金/.test(f) && !/給票/.test(f));
assert(geminiPrompts.at(-1).includes(day6) && geminiPrompts.at(-1).includes(itzy), 'AI 有拿到轉賣中的票');
assert.strictEqual(TR().find(t => t.id === day6).person, '', '還沒確認不會改');
m = pb(A, 'a=rsok&v=' + day6); log('確認', m);
let t = TR().find(x => x.id === day6);
assert.deepStrictEqual([t.person, t.buyerContact, t.amountTwd, t.receivedTwd], ['小美', 'LINE abc123', '4580', '1000']);
assert(/🟡 收訂金/.test(flat(m)));
m = pb(A, 'a=xend&v=1');
// 2. 聊天截圖：看不出哪張 → 選票 → 確認
aiByImage.c1 = { docType: 'chat', seats: [], ticketCount: 0, totalPaid: 0, currency: 'TWD', eventName: '', date: '', understood: true, ids: [], person: '阿明', buyerContact: 'IG @aming', amountTwd: 3200, summary: '阿明想用 3200 買' };
m = img(A, 'c1'); log('聊天截圖', m);
assert(/阿明想用 3200 買/.test(flat(m)) && /對不上是哪一張/.test(flat(m)) && /ITZY/.test(flat(m)));
m = pb(A, 'a=rspick&v=' + itzy); assert(/買家\n— → 阿明/.test(flat(m)) && /🟠 洽談中/.test(flat(m)) === false || /洽談中/.test(flat(m)));
// 3. 不是這張 → 換選 → 取消
m = pb(A, 'a=rsno&v=' + itzy); assert(/那是哪一張/.test(flat(m)) && /DAY6/.test(flat(m)) && !/ITZY/.test(flat(m)));
m = pb(A, 'a=rspick&v=no'); assert(/沒有更新/.test(flat(m)));
assert.strictEqual(TR().find(x => x.id === itzy).person, '');
// 4. 再傳一次、選對、更新 → 洽談中
m = img(A, 'c1'); m = pb(A, 'a=rspick&v=' + itzy); m = pb(A, 'a=rsok&v=' + itzy);
t = TR().find(x => x.id === itzy); assert.deepStrictEqual([t.person, t.buyerContact, t.amountTwd], ['阿明', 'IG @aming', '3200']); assert(/🟠 洽談中/.test(flat(m)));
m = pb(A, 'a=xend&v=1');
// 5. 收尾款＋給票 → 完成
editQueue.push({ intent: 'resaleUpdate' });
editQueue.push({ understood: true, ids: [day6], receivedTwd: 4580, delivered: true, person: '小美', summary: '小美付清尾款，票已給' });
m = text(A, '小美尾款收到了，序號也給她了'); f = flat(m);
assert(/已收\n1,000 → 4,580/.test(f) && /給票\n還沒 → 已給/.test(f) && !/買家/.test(f.split('確認轉賣進度')[1].split('更新後')[0]), '沒變的買家不列出');
m = pb(A, 'a=rsok&v=' + day6); assert(/完成 ✦/.test(flat(m)));
// 6. 一般問題照常回答
editQueue.push({ intent: 'unknown' }); m = text(A, '你好'); assert(/我可以幫你查/.test(flat(m)));
// 7. 沒提到價格買家 → 提示
editQueue.push({ intent: 'resaleUpdate' }); editQueue.push({ understood: true, ids: [itzy], summary: '在談' });
m = text(A, 'ITZY 那張有人在問'); log('沒內容', m); assert(/沒讀到買家、價格或收款/.test(flat(m)));
// 8. 放太久
cache['lr_' + A] = undefined; delete cache['lr_' + A]; m = pb(A, 'a=rsok&v=' + itzy); assert(/放太久/.test(flat(m)));
console.log('\nALL PASSED');
