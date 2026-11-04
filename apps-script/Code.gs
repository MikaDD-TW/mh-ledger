/**
 * ⚠️ 這是「參考用」的 Code.gs（公開 repo，所以拿掉了真實帳目）：
 *   - SEED（舊 Excel 10 月明細）、SEED_BUDGET 的金額與備註、SEED_HISTORY 都換成空的。
 *   - 你 Apps Script 裡的 Code.gs 才是正式版，不要把這份整份貼過去蓋掉。
 *   - 新 App 需要的東西都在 Api.gs，Code.gs 不用改。
 * tests/ 會拿這份加上 Api.gs 在本機模擬 Apps Script 跑測試。
 */
/**
 * @OnlyCurrentDoc
 * 只要求存取「這一份」試算表的權限（不碰你其他檔案），避免被 Google 當成要讀敏感資料而封鎖。
 */
/**
 * MH 記帳本：Google Apps Script 後端
 *
 * 第一次使用：在編輯器上方選 setup → 執行，授權後會建立「明細／額度／月報／設定」四個分頁，
 * 並匯入 2026 年 10 月舊 Excel 的資料。再部署成「網頁應用程式」。
 * setup 可以重複執行，不會清掉已經有的資料。
 */

var SHEET_TX = '明細';
var SHEET_BUDGET = '額度';
var SHEET_REPORT = '月報';
var SHEET_CFG = '設定';
var TZ = 'Asia/Taipei';
var CATS = ['伙食', '休閒', '教育', '交通', '雜費'];
var PAYS = ['現金', '刷卡'];
var WHOS = ['M', 'H'];
var HEAD = ['日期', '用途', '金額', '付款', '分類', '誰出', '鐵盒', '記錄人', '建立時間', 'ID'];
var COL = { date: 1, desc: 2, amt: 3, pay: 4, cat: 5, who: 6, settled: 7, by: 8, created: 9, id: 10 };
var KEEP_DAYS = 400; // App 讀取最近幾天的明細（未結清的不論多久都會讀）

// 舊 Excel「伙食2萬」分頁 10 月明細：日期, 用途, 金額, 付款, 分類, 誰出, 鐵盒(1=已結清)
var SEED = [];

// 舊 Excel 第 2 列的額度（F2:J2），備註是 B1 的薪資分配
var SEED_BUDGET = ['2026-10', 0, 0, 0, 0, 0, ''];

// 舊 Excel A2 儲存格的每月紀錄（照原本數字，沒有 6 月）
var SEED_HISTORY = [];

// ---------------------------------------------------------------- 網頁
function doGet(e) {
  var p = (e && e.parameter) || {};
  var key = PropertiesService.getScriptProperties().getProperty('KEY');
  var me = String(p.me || '').toUpperCase();
  var cfg = {
    k: String(p.k || ''),
    me: WHOS.indexOf(me) >= 0 ? me : '',
    ok: !!key && p.k === key
  };
  var t = HtmlService.createTemplateFromFile('Index');
  t.cfg = JSON.stringify(cfg).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('MH 記帳本')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** App 唯一的呼叫入口。k 是通行碼。 */
function api(k, action, p) {
  checkKey_(k);
  p = p || {};
  switch (action) {
    case 'load': return load_();
    case 'add': return add_(p.entries);
    case 'update': return update_(p.entry);
    case 'remove': return remove_(p.id);
    case 'settle': return settle_(p.ids, p.value);
    case 'parse': return parse_(p.text, p.today);
  }
  throw new Error('不支援的動作：' + action);
}

// ---------------------------------------------------------------- 初始化
function setup() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  TZ_CACHE = TZ;
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('KEY');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    props.setProperty('KEY', key);
  }

  // 明細
  var tx = ss.getSheetByName(SHEET_TX) || ss.insertSheet(SHEET_TX, 0);
  if (tx.getLastRow() === 0) {
    tx.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    styleHeader_(tx.getRange(1, 1, 1, HEAD.length));
    tx.setFrozenRows(1);
    [92, 200, 80, 60, 60, 56, 52, 64, 132, 110].forEach(function (w, i) { tx.setColumnWidth(i + 1, w); });
  }
  applyColumnRules_(tx);
  if (tx.getLastRow() === 1 && SEED.length) {
    var now = new Date();
    var rows = SEED.map(function (r) {
      return entryToRow_({ date: r[0], desc: r[1], amt: r[2], pay: r[3], cat: r[4], who: r[5] || '',
        settled: !!r[6], by: 'Excel', id: newId_() }, now);
    });
    ensureRows_(tx, 1 + rows.length);
    tx.getRange(2, 1, rows.length, HEAD.length).setValues(rows);
  }

  // 額度
  var bs = ss.getSheetByName(SHEET_BUDGET);
  if (!bs) {
    bs = ss.insertSheet(SHEET_BUDGET);
    bs.getRange('A:A').setNumberFormat('@');
    bs.getRange(1, 1, 1, CATS.length + 2).setValues([['月份'].concat(CATS, ['備註'])]);
    styleHeader_(bs.getRange(1, 1, 1, CATS.length + 2));
    bs.getRange(2, 1, 1, CATS.length + 2).setValues([SEED_BUDGET]);
    bs.getRange('B:F').setNumberFormat('#,##0');
    bs.setFrozenRows(1);
    bs.setColumnWidth(CATS.length + 2, 520);
    bs.getRange('A1').setNote('每列一個月份（格式 2026-11），由舊到新往下排。\n某個月沒有列，就沿用前一個有填的月份。');
  }

  // 月報
  if (!ss.getSheetByName(SHEET_REPORT)) buildReport_(ss.insertSheet(SHEET_REPORT));

  // 設定
  var cfg = ss.getSheetByName(SHEET_CFG);
  if (!cfg) { cfg = ss.insertSheet(SHEET_CFG); buildCfg_(cfg); }
  cfg.getRange('B3').setNumberFormat('@').setValue(key);

  // 刪掉預設的空白分頁
  var mine = [SHEET_TX, SHEET_BUDGET, SHEET_REPORT, SHEET_CFG];
  ss.getSheets().forEach(function (sh) {
    if (mine.indexOf(sh.getName()) < 0 && sh.getLastRow() === 0 && sh.getLastColumn() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });
  Logger.log('設定完成。通行碼：' + key + '。下一步：部署 → 新增部署作業 → 網頁應用程式。');
}

function buildReport_(rp) {
  var dateRange = '\'明細\'!$A:$A,">="&$B$2,\'明細\'!$A:$A,"<="&$C$2';
  rp.getRange('B1').setNumberFormat('@');
  rp.getRange('A1:C1').setValues([['月份', Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), '← 改這格看其他月份，例如 2026-11']]);
  rp.getRange('A2').setValue('期間');
  rp.getRange('B2').setFormula('=DATE(VALUE(LEFT($B$1,4)),VALUE(RIGHT($B$1,2)),1)');
  rp.getRange('C2').setFormula('=EOMONTH(B2,0)');
  rp.getRange('B2:C2').setNumberFormat('yyyy/mm/dd');

  rp.getRange('A4:E4').setValues([['分類', '額度', '已花', '剩餘', '使用率']]);
  styleHeader_(rp.getRange('A4:E4'));
  var budgetCol = ['B', 'C', 'D', 'E', 'F'];
  CATS.forEach(function (c, i) {
    var r = 5 + i;
    rp.getRange(r, 1).setValue(c);
    rp.getRange(r, 2).setFormula('=IFERROR(INDEX(\'額度\'!$' + budgetCol[i] + '$2:$' + budgetCol[i] + ',MATCH($B$1,\'額度\'!$A$2:$A,1)),0)');
    rp.getRange(r, 3).setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$E:$E,$A' + r + ',' + dateRange + ')');
    rp.getRange(r, 4).setFormula('=B' + r + '-C' + r);
    rp.getRange(r, 5).setFormula('=IF(B' + r + '>0,C' + r + '/B' + r + ',"")');
  });
  rp.getRange('A10').setValue('合計');
  rp.getRange('B10').setFormula('=SUM(B5:B9)');
  rp.getRange('C10').setFormula('=SUM(C5:C9)');
  rp.getRange('D10').setFormula('=B10-C10');
  rp.getRange('E10').setFormula('=IF(B10>0,C10/B10,"")');
  rp.getRange('A10:E10').setFontWeight('bold');

  rp.getRange('A12:B12').setValues([['本月付款', '金額']]);
  styleHeader_(rp.getRange('A12:B12'));
  rp.getRange('A13:A16').setValues([['現金'], ['刷卡'], ['M 出的現金'], ['H 出的現金']]);
  rp.getRange('B13').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",' + dateRange + ')');
  rp.getRange('B14').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"刷卡",' + dateRange + ')');
  rp.getRange('B15').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",\'明細\'!$F:$F,"M",' + dateRange + ')');
  rp.getRange('B16').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",\'明細\'!$F:$F,"H",' + dateRange + ')');

  rp.getRange('A18:B18').setValues([['鐵盒還沒結清（所有月份）', '金額']]);
  styleHeader_(rp.getRange('A18:B18'));
  rp.getRange('A19:A22').setValues([['M 出的現金'], ['H 出的現金'], ['現金沒填人'], ['刷卡']]);
  rp.getRange('B19').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",\'明細\'!$F:$F,"M",\'明細\'!$G:$G,FALSE)');
  rp.getRange('B20').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",\'明細\'!$F:$F,"H",\'明細\'!$G:$G,FALSE)');
  rp.getRange('B21').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"現金",\'明細\'!$G:$G,FALSE)-B19-B20');
  rp.getRange('B22').setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$D:$D,"刷卡",\'明細\'!$G:$G,FALSE)');

  // 每月一覽：從 2026-10 起 24 個月
  var head = ['月份'].concat(CATS, ['合計', '額度合計', '差額']);
  rp.getRange(24, 1, 1, head.length).setValues([head]);
  styleHeader_(rp.getRange(24, 1, 1, head.length));
  rp.getRange('A23').setValue('每月一覽').setFontWeight('bold');
  rp.getRange('A25:A48').setNumberFormat('@');
  for (var i = 0; i < 24; i++) {
    var r = 25 + i;
    var y = 2026 + Math.floor((9 + i) / 12), m = ((9 + i) % 12) + 1;
    rp.getRange(r, 1).setValue(y + '-' + (m < 10 ? '0' : '') + m);
    var start = 'DATE(VALUE(LEFT($A' + r + ',4)),VALUE(RIGHT($A' + r + ',2)),1)';
    var range = '\'明細\'!$A:$A,">="&' + start + ',\'明細\'!$A:$A,"<="&EOMONTH(' + start + ',0)';
    for (var j = 0; j < CATS.length; j++) {
      rp.getRange(r, 2 + j).setFormula('=SUMIFS(\'明細\'!$C:$C,\'明細\'!$E:$E,' + colLetter_(2 + j) + '$24,' + range + ')');
    }
    rp.getRange(r, 7).setFormula('=SUM(B' + r + ':F' + r + ')');
    rp.getRange(r, 8).setFormula('=IFERROR(SUM(INDEX(\'額度\'!$B$2:$F,MATCH($A' + r + ',\'額度\'!$A$2:$A,1),0)),0)');
    rp.getRange(r, 9).setFormula('=H' + r + '-G' + r);
  }

  rp.getRange('K23').setValue('舊 Excel 的每月紀錄（照原本 A2 的數字）').setFontWeight('bold');
  rp.getRange('K24:L24').setValues([['月份', '金額']]);
  styleHeader_(rp.getRange('K24:L24'));
  if (SEED_HISTORY.length) {
    rp.getRange('K25:K' + (24 + SEED_HISTORY.length)).setNumberFormat('@');
    rp.getRange(25, 11, SEED_HISTORY.length, 2).setValues(SEED_HISTORY);
  }

  rp.getRange('B5:D16').setNumberFormat('#,##0');
  rp.getRange('B19:B22').setNumberFormat('#,##0');
  rp.getRange('B25:I48').setNumberFormat('#,##0');
  rp.getRange('L25:L48').setNumberFormat('#,##0');
  rp.getRange('E5:E10').setNumberFormat('0%');
  rp.setColumnWidth(1, 170);
  rp.setColumnWidth(3, 110);
  rp.setColumnWidth(11, 90);
}

function buildCfg_(cfg) {
  cfg.getRange('A1').setValue('MH 記帳本設定').setFontWeight('bold').setFontSize(14);
  cfg.getRange('A3').setValue('通行碼');
  cfg.getRange('A5').setValue('把部署後拿到的網址（結尾是 /exec）貼在右邊 →');
  cfg.getRange('B5').setBackground('#FFF8D6');
  cfg.getRange('A6').setValue('M 的 App 連結');
  cfg.getRange('B6').setFormula('=IF(B5="","",B5&"?k="&B3&"&me=M")');
  cfg.getRange('A7').setValue('H 的 App 連結');
  cfg.getRange('B7').setFormula('=IF(B5="","",B5&"?k="&B3&"&me=H")');
  cfg.getRange('A9').setValue('連結外流的話：Apps Script → 專案設定 → 指令碼屬性 → 把 KEY 改成新的，再到這裡把 B3 改成一樣，兩人用新連結重開。');
  cfg.setColumnWidth(1, 300);
  cfg.setColumnWidth(2, 620);
}

// ---------------------------------------------------------------- 讀取
var TZ_CACHE = null;
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return TZ_CACHE || (TZ_CACHE = ss_().getSpreadsheetTimeZone() || TZ); }
function pad2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function colLetter_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function toISO_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var m = String(v == null ? '' : v).match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) : '';
}

function shiftDays_(iso, n) {
  var p = iso.split('-');
  var t = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return t.getUTCFullYear() + '-' + pad2_(t.getUTCMonth() + 1) + '-' + pad2_(t.getUTCDate());
}

function txSheet_() {
  var sh = ss_().getSheetByName(SHEET_TX);
  if (!sh) throw new Error('找不到「明細」分頁，請先在 Apps Script 執行 setup');
  return sh;
}

function rowToEntry_(r) {
  var g = r[6];
  return {
    date: toISO_(r[0]),
    desc: String(r[1] == null ? '' : r[1]).trim(),
    amt: Number(r[2]) || 0,
    pay: String(r[3] || '').trim(),
    cat: String(r[4] || '').trim(),
    who: String(r[5] || '').trim().toUpperCase(),
    settled: g === true || /^(v|true)$/i.test(String(g || '').trim()),
    by: String(r[7] || ''),
    id: String(r[9] || '')
  };
}

function readAll_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEAD.length).getValues().map(function (r, i) {
    return { row: i + 2, e: rowToEntry_(r) };
  });
}

function readBudgets_() {
  var sh = ss_().getSheetByName(SHEET_BUDGET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1 + CATS.length).getValues().map(function (r) {
    var raw = Object.prototype.toString.call(r[0]) === '[object Date]'
      ? Utilities.formatDate(r[0], tz_(), 'yyyy-MM') : String(r[0]).trim();
    var mm = raw.match(/^(\d{4})[-\/.](\d{1,2})$/);
    if (!mm) return null;
    var o = { month: mm[1] + '-' + pad2_(mm[2]) };
    CATS.forEach(function (c, i) { o[c] = Number(r[i + 1]) || 0; });
    return o;
  }).filter(Boolean).sort(function (a, b) { return a.month < b.month ? -1 : a.month > b.month ? 1 : 0; });
}

function load_() {
  var sh = txSheet_();
  var rows = readAll_(sh);
  var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var since = shiftDays_(today, -KEEP_DAYS);

  // 在試算表手動加的列沒有 ID，補上
  var missing = rows.filter(function (x) { return !x.e.id && x.e.desc && x.e.date; });
  if (missing.length) {
    withLock_(function () {
      missing.forEach(function (x) { x.e.id = newId_(); sh.getRange(x.row, COL.id).setValue(x.e.id); });
    });
  }
  var entries = rows.map(function (x) { return x.e; }).filter(function (e) {
    return e.date && e.desc && e.id && (e.date >= since || !e.settled);
  });
  return { today: today, since: since, entries: entries, budgets: readBudgets_() };
}

// ---------------------------------------------------------------- 寫入
function cleanEntry_(e) {
  e = e || {};
  var out = {
    id: String(e.id || '').replace(/[^\w-]/g, '').slice(0, 40) || newId_(),
    date: String(e.date || ''),
    desc: String(e.desc == null ? '' : e.desc).replace(/[\s　]+/g, ''),
    amt: Number(String(e.amt == null ? '' : e.amt).replace(/,/g, '')),
    pay: String(e.pay || ''),
    cat: String(e.cat || ''),
    who: String(e.who || '').toUpperCase(),
    settled: e.settled === true,
    by: String(e.by || '').toUpperCase().slice(0, 8)
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(out.date)) throw new Error('日期格式不對：' + out.date);
  if (!out.desc) throw new Error('用途是空的');
  if (String(e.amt).trim() === '' || !isFinite(out.amt) || out.amt < 0 || Math.floor(out.amt) !== out.amt) {
    throw new Error('金額要是 0 以上的整數：' + e.amt);
  }
  if (PAYS.indexOf(out.pay) < 0) throw new Error('付款要是現金或刷卡：' + out.pay);
  if (CATS.indexOf(out.cat) < 0) throw new Error('分類要是 ' + CATS.join('、') + ' 其中一個：' + out.cat);
  if (out.pay !== '現金') out.who = '';
  if (out.who && WHOS.indexOf(out.who) < 0) throw new Error('誰出要是 M 或 H：' + out.who);
  return out;
}

function entryToRow_(e, created) {
  return [e.date, e.desc, e.amt, e.pay, e.cat, e.who, !!e.settled, e.by || '', created, e.id];
}

function idIndex_(sh) {
  var last = sh.getLastRow();
  var map = {};
  if (last < 2) return map;
  sh.getRange(2, COL.id, last - 1, 1).getValues().forEach(function (r, i) {
    if (r[0] !== '' && r[0] != null) map[String(r[0])] = i + 2;
  });
  return map;
}

function add_(list) {
  if (!Array.isArray(list) || !list.length) return { added: [], ok: [] };
  if (list.length > 500) throw new Error('一次最多 500 筆');
  var clean = list.map(cleanEntry_);
  return withLock_(function () {
    var sh = txSheet_();
    var ids = idIndex_(sh);
    var fresh = [];
    clean.forEach(function (e) { if (!ids[e.id]) { ids[e.id] = -1; fresh.push(e); } });
    if (fresh.length) {
      var start = sh.getLastRow() + 1;
      ensureRows_(sh, start + fresh.length - 1);
      var now = new Date();
      sh.getRange(start, 1, fresh.length, HEAD.length).setValues(fresh.map(function (e) { return entryToRow_(e, now); }));
    }
    return { added: fresh.map(function (e) { return e.id; }), ok: clean.map(function (e) { return e.id; }) };
  });
}

function update_(e) {
  var c = cleanEntry_(e);
  return withLock_(function () {
    var sh = txSheet_();
    var row = idIndex_(sh)[c.id];
    if (!row) throw new Error('找不到這筆，可能已經被刪掉了');
    sh.getRange(row, 1, 1, 6).setValues([[c.date, c.desc, c.amt, c.pay, c.cat, c.who]]);
    return { ok: c.id };
  });
}

function remove_(id) {
  id = String(id || '');
  return withLock_(function () {
    var sh = txSheet_();
    var row = idIndex_(sh)[id];
    if (!row) throw new Error('找不到這筆，可能已經被刪掉了');
    var e = rowToEntry_(sh.getRange(row, 1, 1, HEAD.length).getValues()[0]);
    sh.deleteRow(row);
    return { removed: e };
  });
}

function settle_(ids, value) {
  if (!Array.isArray(ids) || !ids.length) return { changed: [] };
  return withLock_(function () {
    var sh = txSheet_();
    var idx = idIndex_(sh);
    var changed = [], cells = [];
    ids.forEach(function (id) { var r = idx[String(id)]; if (r) { changed.push(String(id)); cells.push('G' + r); } });
    if (cells.length) sh.getRangeList(cells).setValue(value !== false);
    return { changed: changed };
  });
}

// ---------------------------------------------------------------- 解析 LINE 文字
/**
 * 支援兩種寫法：
 *   10/21            ← 日期一行（後面可接星期）
 *   早餐飯糰 55 現金 伙食 H
 * 或每行帶日期：10/21 早餐飯糰 55 現金 伙食 H
 * 用途裡有空白也可以，會以「現金／刷卡」的位置來切。
 */
function parse_(text, today) {
  today = /^\d{4}-\d{2}-\d{2}$/.test(String(today || '')) ? today : Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var limit = shiftDays_(today, 31);
  var entries = [], errors = [], cur = '';
  var toHalf = function (s) {
    return s.replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
      .replace(/[\s　]+/g, ' ').trim();
  };
  var mkDate = function (mo, d) {
    var y = +today.slice(0, 4);
    var iso = y + '-' + pad2_(mo) + '-' + pad2_(d);
    if (iso > limit) iso = (y - 1) + '-' + pad2_(mo) + '-' + pad2_(d);
    var chk = new Date(Date.UTC(+iso.slice(0, 4), mo - 1, d));
    return chk.getUTCMonth() === mo - 1 ? iso : '';
  };
  String(text || '').split(/\r?\n/).forEach(function (raw, i) {
    var line = toHalf(raw);
    if (!line) return;
    var tok = line.split(' ');
    var dm = tok[0].match(/^(\d{1,2})[\/.\-](\d{1,2})(?!\d)/);
    if (dm) {
      var dt = mkDate(+dm[1], +dm[2]);
      if (!dt) { if (tok.length < 4) cur = ''; errors.push({ line: i + 1, text: raw.trim(), msg: '日期不存在' }); return; }
      if (tok.length < 4) { cur = dt; return; }   // 日期列，例如「10/18星期日」
      tok = tok.slice(1);
      var r = parseTokens_(tok);
      if (r.error) errors.push({ line: i + 1, text: raw.trim(), msg: r.error });
      else { r.entry.date = dt; entries.push(r.entry); }
      return;
    }
    if (!cur) { errors.push({ line: i + 1, text: raw.trim(), msg: '上面沒有日期，請先加一行日期（例如 10/21）' }); return; }
    var r2 = parseTokens_(tok);
    if (r2.error) errors.push({ line: i + 1, text: raw.trim(), msg: r2.error });
    else { r2.entry.date = cur; entries.push(r2.entry); }
  });
  return { entries: entries, errors: errors };
}

function parseTokens_(tok) {
  var pi = -1;
  for (var i = 1; i < tok.length; i++) if (PAYS.indexOf(tok[i]) >= 0) { pi = i; break; }
  if (pi < 2) return { error: '找不到「金額 現金/刷卡」' };
  var amtStr = tok[pi - 1].replace(/[,$]/g, '');
  if (!/^\d+$/.test(amtStr)) return { error: '金額不是數字：' + tok[pi - 1] };
  var desc = tok.slice(0, pi - 1).join('');
  var cat = tok[pi + 1] || '';
  if (CATS.indexOf(cat) < 0) return { error: cat ? '分類「' + cat + '」不在 ' + CATS.join('、') + ' 裡' : '少了分類' };
  var who = (tok[pi + 2] || '').toUpperCase();
  if (who && WHOS.indexOf(who) < 0) return { error: '最後一欄要是 M 或 H：' + tok[pi + 2] };
  if (tok.length > pi + 3) return { error: '多了看不懂的字：' + tok.slice(pi + 3).join(' ') };
  var pay = tok[pi];
  return { entry: { desc: desc, amt: +amtStr, pay: pay, cat: cat, who: pay === '現金' ? who : '' } };
}

// ---------------------------------------------------------------- 工具
function checkKey_(k) {
  var key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key || String(k || '') !== key) throw new Error('通行碼不對，請用試算表「設定」分頁裡的連結重新開啟');
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function ensureRows_(sh, lastRowNeeded) {
  var max = sh.getMaxRows();
  if (lastRowNeeded > max) {
    sh.insertRowsAfter(max, lastRowNeeded - max + 200);
    applyColumnRules_(sh);
  }
}

function applyColumnRules_(tx) {
  var list = function (vals) {
    return SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build();
  };
  tx.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  tx.getRange('B2:B').setNumberFormat('@');
  tx.getRange('C2:C').setNumberFormat('#,##0');
  tx.getRange('D2:F').setNumberFormat('@');
  tx.getRange('H2:H').setNumberFormat('@');
  tx.getRange('I2:I').setNumberFormat('yyyy/mm/dd hh:mm');
  tx.getRange('J2:J').setNumberFormat('@');
  tx.getRange('D2:D').setDataValidation(list(PAYS));
  tx.getRange('E2:E').setDataValidation(list(CATS));
  tx.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(WHOS, true).setAllowInvalid(false).build());
  tx.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

function styleHeader_(range) {
  range.setFontWeight('bold').setBackground('#E7ECF5').setFontColor('#17202B');
}
