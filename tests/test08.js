// ===== 活動網頁 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
setNow('2026-10-07', '10:00');
let m;
const H = n => sheets[n].rows[0];
const addRow = (n, o) => sheets[n].rows.push(H(n).map(h => o[h] ?? ''));
addRow('events', { id: 'ev1', name: '3RD WORLD TOUR', artist: 'DAY6', city: '台北', venue: '北流', startDate: '2026-12-05', startTime: '19:00', eventNumber: '900' });
addRow('events', { id: 'ev2', name: 'TUNNEL VISION', artist: 'ITZY', city: '台北', venue: '北流', startDate: '2026-12-20', eventNumber: '901' });
addRow('onsales', { id: 'os1', groupId: 'g1', title: 'IVE TAIPEI', artist: 'IVE', city: '台北', saleAt: '2026-10-07 12:00', phase: '全面開賣', platform: '拓元', watchers: JSON.stringify([A]), done: '[]' });
const uris = msgs => JSON.stringify(msgs).match(/"uri":"[^"]+"/g) || [];
// 1. 未來場次：沒存網址 → 搜尋連結
m = text(A, '未來場次全部'); log('未來場次', m);
assert(/🔍 找活動網頁/.test(flat(m)));
assert(uris(m).some(u => u.includes('google.com/search?q=' + encodeURIComponent('DAY6 3RD WORLD TOUR 台北 售票'))));
// 2. 貼售票頁網址 → 自動判斷是售票頁、選場次
m = text(A, 'https://kktix.com/events/day6-tp'); log('貼網址', m);
assert(/這是哪一場的售票頁網址/.test(flat(m)) && /12\/5 DAY6/.test(flat(m)) && /開賣 IVE/.test(flat(m)));
m = pb(A, 'a=lurl&v=e:ev1'); assert(/已存成「DAY6 3RD WORLD TOUR」的🎫 售票頁/.test(flat(m)) || /已存成/.test(flat(m)), flat(m));
assert.strictEqual(table('events').find(e => e.id === 'ev1').ticketUrl, 'https://kktix.com/events/day6-tp');
// 3. 改成公告
m = pb(A, JSON.stringify(m).match(/a=lkind&v=[^"]+/)[0]); assert(/📢 公告/.test(flat(m)));
let ev = table('events').find(e => e.id === 'ev1'); assert.deepStrictEqual([ev.ticketUrl, ev.noticeUrl], ['', 'https://kktix.com/events/day6-tp']);
// 4. 貼公告網址＋關鍵字
m = text(A, 'ITZY https://www.instagram.com/p/abc'); assert(/哪一場的公告網址/.test(flat(m)) && /ITZY/.test(flat(m)) && !/DAY6/.test(flat(m)), '關鍵字縮小');
m = pb(A, 'a=lurl&v=e:ev2'); assert.strictEqual(table('events').find(e => e.id === 'ev2').noticeUrl, 'https://www.instagram.com/p/abc');
// 5. 未來場次顯示存好的連結
m = text(A, '未來場次全部'); assert(/📢 公告/.test(flat(m)) && /🔍 搜尋/.test(flat(m)) && uris(m).some(u => u.includes('instagram.com/p/abc')));
// 6. 貼了不存、打字取消
m = text(A, 'https://tixcraft.com/x'); m = pb(A, 'a=lurl&v=no'); assert(/沒有存/.test(flat(m)));
m = text(A, 'https://tixcraft.com/x'); m = text(A, '取消'); assert(/網址沒有存/.test(flat(m)));
// 7. 等選場次時打別的話 → 當一般訊息
m = text(A, 'https://tixcraft.com/x'); editQueue.push({ intent: 'unknown' }); m = text(A, '今天天氣好嗎'); assert(/我可以幫你查/.test(flat(m)));
// 8. 即將開賣：存售票頁、修改中貼網址直接存
m = text(A, '即將開賣'); assert(/🔍 找活動網頁/.test(flat(m)));
m = pb(A, 'a=oedit&v=g1'); m = text(A, 'https://tixcraft.com/activity/ive'); assert(/已存成「IVE TAIPEI」的🎫 售票頁/.test(flat(m)));
assert.strictEqual(table('onsales')[0].ticketUrl, 'https://tixcraft.com/activity/ive'); m = pb(A, 'a=xend&v=1');
m = text(A, '即將開賣'); assert(/🎫 售票頁/.test(flat(m)));
// 9. 修改即將開賣後網址保留
editQueue.push({ understood: true, notes: '實名制' }); m = pb(A, 'a=oedit&v=g1'); m = text(A, '注意事項改成實名制');
assert.strictEqual(table('onsales')[0].ticketUrl, 'https://tixcraft.com/activity/ive', '改其他欄位不會弄丟網址'); m = pb(A, 'a=xend&v=1');
// 10. 開賣前 30 分鐘提醒附售票頁按鈕
setNow('2026-10-07', '11:31'); pushes.length = 0; vm.runInContext('reminderTick()', ctx);
const p = pushes.find(x => JSON.stringify(x).includes('快開賣了')); assert(p, '有提醒');
assert(JSON.stringify(p).includes('tixcraft.com/activity/ive'), '附售票頁');
assert(p.messages.at(-1).quickReply, '搶到了按鈕留在最後一則');
console.log('\nALL PASSED');
