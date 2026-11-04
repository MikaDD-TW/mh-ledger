/**
 * MH 記帳本：給新網頁 App（GitHub Pages 上的 PWA）用的 JSON 入口。
 *
 * 安裝（不用改 Code.gs）：
 *   1. Apps Script 編輯器左側「檔案」旁的 ＋ → 指令碼，檔名打 Api，把這整份貼上，儲存。
 *   2. 上方函式選 setupApp → 執行一次。「設定」分頁會多出 B11–B13：新 App 網址和 M、H 各自的連結。
 *   3. 部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署。/exec 網址不會變。
 *
 * App 用 POST 送 {k, action, p}（Content-Type: text/plain，避開 CORS 預檢），
 * 這裡轉給 Code.gs 的 api()，回傳 {ok:true, data} 或 {ok:false, error}。
 */

var APP_URL_DEFAULT = 'https://mikadd-tw.github.io/mh-ledger/';

function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents) || '';
    var req = body ? JSON.parse(body) : {};
    if (req.action === 'ping') {
      checkKey_(req.k);
      out = { ok: true, data: { pong: true, today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd') } };
    } else {
      out = { ok: true, data: api(req.k, req.action, req.p) };
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 在「設定」分頁加上新 App 的連結。可以重複執行。 */
function setupApp() {
  var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CFG);
  if (!cfg) throw new Error('找不到「設定」分頁，請先執行 setup');
  cfg.getRange('A11').setValue('新 App 網址（GitHub Pages）').setFontWeight('bold');
  if (!cfg.getRange('B11').getValue()) cfg.getRange('B11').setValue(APP_URL_DEFAULT);
  cfg.getRange('B11').setBackground('#FFF8D6');
  cfg.getRange('A12').setValue('M 的新 App 連結');
  cfg.getRange('B12').setFormula('=IF(OR($B$5="",$B$11=""),"",$B$11&"?me=M&k="&$B$3&"&u="&ENCODEURL($B$5))');
  cfg.getRange('A13').setValue('H 的新 App 連結');
  cfg.getRange('B13').setFormula('=IF(OR($B$5="",$B$11=""),"",$B$11&"?me=H&k="&$B$3&"&u="&ENCODEURL($B$5))');
  cfg.getRange('A14').setValue('用 iPhone 的 Safari 打開自己的連結 → 分享 → 加入主畫面。B5 要先貼好 /exec 網址。');
  Logger.log('設定分頁已加上新 App 連結。記得：部署 → 管理部署作業 → 編輯 → 新版本 → 部署。');
}
