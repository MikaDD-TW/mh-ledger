// 在 Node 裡模擬 Apps Script：gas-mock.js（假的 SpreadsheetApp 等）+ apps-script/Code.gs + Api.gs
const fs = require('fs'), vm = require('vm'), path = require('path');

function createSim() {
  const ctx = vm.createContext({ console, Intl, Date, Math, JSON, Object, Array, String, Number, RegExp, Error, isFinite, isNaN, Set, Map, Proxy });
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'gas-mock.js'), 'utf8'), ctx);
  vm.runInContext(`
    var ContentService = {
      MimeType: { JSON: 'application/json', TEXT: 'text/plain' },
      createTextOutput: function (s) {
        var o = { _c: String(s), _m: 'text/plain' };
        o.setMimeType = function (m) { o._m = m; return o; };
        o.getContent = function () { return o._c; };
        return o;
      }
    };`, ctx);
  for (const f of ['Code.gs', 'Api.gs']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'), ctx, { filename: f });
  }
  ctx.setup();
  const run = (code) => vm.runInContext(code, ctx);
  const key = run("PropertiesService.getScriptProperties().getProperty('KEY')");
  return {
    ctx, run, key,
    post(bodyString) { return ctx.doPost({ postData: { contents: bodyString, type: 'text/plain' } }).getContent(); },
    api(action, p, k) { return JSON.parse(this.post(JSON.stringify({ k: k === undefined ? key : k, action, p }))); },
    rows() {
      const sh = run("__ss.getSheetByName('明細')");
      const lr = sh.getLastRow();
      return lr < 2 ? [] : sh.getRange(2, 1, lr - 1, 10).getValues();
    },
  };
}

// 測試用的假帳：這個月和上個月各幾筆
function fakeEntries(today) {
  const [y, m] = today.split('-').map(Number);
  const ym = (yy, mm) => yy + '-' + String(mm).padStart(2, '0');
  const cur = ym(y, m), prev = m === 1 ? ym(y - 1, 12) : ym(y, m - 1);
  const d = (mo, day) => mo + '-' + String(day).padStart(2, '0');
  const list = [
    [d(prev, 3), '測試課程', 300, '現金', '教育', 'M', true],
    [d(prev, 5), '測試超市', 800, '刷卡', '伙食', '', true],
    [d(prev, 12), '測試加油', 1000, '刷卡', '交通', '', true],
    [d(prev, 20), '測試電影', 500, '現金', '休閒', 'H', false],
    [d(cur, 1), '測試早餐', 50, '現金', '伙食', 'M', false],
    [d(cur, 1), '測試午餐', 600, '刷卡', '伙食', '', false],
    [d(cur, 2), '測試咖啡', 250, '現金', '休閒', 'M', false],
    [d(cur, 2), '測試量販', 4000, '現金', '伙食', 'H', false],
    [d(cur, 3), '測試晚餐', 150, '現金', '伙食', 'M', true],
    [d(cur, 3), '測試麵包', 180, '刷卡', '雜費', 'H', false],
  ];
  return list.map((r, i) => ({ id: 'fake' + i, date: r[0], desc: r[1], amt: r[2], pay: r[3], cat: r[4], who: r[5], settled: r[6], by: 'M' }));
}

module.exports = { createSim, fakeEntries };
