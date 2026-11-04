// node tests/backend.test.js — 測 Api.gs 的 doPost 包裝和 Code.gs 的 API
const assert = require('assert');
const { createSim, fakeEntries } = require('./gas-sim');

const sim = createSim();
const today = sim.run("Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd')");

// 通行碼錯
let r = sim.api('load', {}, 'wrong');
assert.strictEqual(r.ok, false);
assert.match(r.error, /通行碼/);

// 不是 JSON
r = JSON.parse(sim.post('not json'));
assert.strictEqual(r.ok, false);

// ping
r = sim.api('ping');
assert.deepStrictEqual([r.ok, r.data.pong, r.data.today], [true, true, today]);

// add + load
r = sim.api('add', { entries: fakeEntries(today) });
assert.strictEqual(r.ok, true, JSON.stringify(r));
r = sim.api('load');
assert.strictEqual(r.ok, true);
assert.strictEqual(r.data.entries.length, 10);
assert.strictEqual(r.data.today, today);
assert.ok(Array.isArray(r.data.budgets));

// update / settle / remove
r = sim.api('update', { entry: Object.assign({}, fakeEntries(today)[4], { amt: 60 }) });
assert.strictEqual(r.ok, true, JSON.stringify(r));
r = sim.api('settle', { ids: ['fake5', 'fake6'], value: true });
assert.strictEqual(r.ok, true);
r = sim.api('remove', { id: 'fake9' });
assert.strictEqual(r.ok, true);
const data = sim.api('load').data;
assert.strictEqual(data.entries.length, 9);
assert.strictEqual(data.entries.find((e) => e.id === 'fake4').amt, 60);
assert.strictEqual(data.entries.find((e) => e.id === 'fake6').settled, true);

// parse
r = sim.api('parse', { text: '10/21\n測試早餐 50 現金 伙食 H\n測試飲料 60 刷卡 休閒', today });
assert.strictEqual(r.ok, true);
assert.strictEqual(r.data.entries.length, 2);

// 未知動作
r = sim.api('nope');
assert.strictEqual(r.ok, false);

// setupApp 連結公式
sim.run('setupApp()');
const f12 = sim.run("__ss.getSheetByName('設定').getRange('B12').getValue()");
assert.match(f12, /me=M&k=.*ENCODEURL\(\$B\$5\)/);
assert.strictEqual(sim.run("__ss.getSheetByName('設定').getRange('B11').getValue()"), 'https://mikadd-tw.github.io/mh-ledger/');

console.log('backend tests passed');
