// ===== 模型順序、塞車避開、自己的記錄 =====
const A = 'Uaaa';
text(A, 'pw123'); text(A, 'Chi');
const order = hasImage => JSON.parse(vm.runInContext('JSON.stringify(geminiOrder_(' + hasImage + '))', ctx));
assert.deepStrictEqual(order(false), ['gemini-3.5-flash-lite', 'gemini-3.8-flash'], '文字先用快的');
assert.deepStrictEqual(order(true), ['gemini-3.8-flash', 'gemini-3.5-flash-lite'], '讀圖先用準的');
cache['gslow_gemini-3.8-flash'] = '1';
assert.deepStrictEqual(order(true), ['gemini-3.5-flash-lite', 'gemini-3.8-flash'], '塞車的排最後');
delete cache['gslow_gemini-3.8-flash'];
cache['gfly_gemini-3.8-flash'] = String(new Date(NOW_ISO).getTime() - 30000);
assert.deepStrictEqual(order(true), ['gemini-3.5-flash-lite', 'gemini-3.8-flash'], '上一次還沒回來也當塞車');
delete cache['gfly_gemini-3.8-flash'];
// 文字問題實際打到的模型
const models = []; const realFetch = ctx.UrlFetchApp.fetch;
ctx.UrlFetchApp.fetch = function (url, o) { if (url.includes('generativelanguage')) models.push(url.match(/models\/([^:]+)/)[1]); return realFetch.apply(this, arguments); };
editQueue.push({ intent: 'unknown' }); text(A, '你好');
assert.strictEqual(models[0], 'gemini-3.5-flash-lite');
assert(!cache['gfly_gemini-3.5-flash-lite'], '呼叫完清掉進行中標記');
const dbg = JSON.parse(cache['dbg']);
assert(dbg.some(l => /收到 message text 你好/.test(l)) && dbg.some(l => /gemini gemini-3.5-flash-lite/.test(l)) && dbg.some(l => /回覆 200/.test(l)), dbg.join('\n'));
console.log('\nALL PASSED');
