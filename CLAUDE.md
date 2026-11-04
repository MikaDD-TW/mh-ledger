# MH 記帳本（mh-ledger）

M 和 H 兩個人共用的記帳 App。前端是放在 GitHub Pages 的 PWA，資料存在一份 Google 試算表，後端是綁在那份試算表上的 Apps Script。

使用者 Mika 用繁體中文溝通。App 裡所有文字都用台灣繁體中文。

## 架構

```
iPhone 主畫面（PWA） ──POST JSON──▶ Apps Script 網頁應用程式（/exec） ──▶ Google 試算表
   index.html + sw.js                 apps-script/Api.gs → Code.gs 的 api()
   GitHub Pages：https://mikadd-tw.github.io/mh-ledger/
```

- `index.html`：整個 App（HTML＋CSS＋JS 一個檔，沒有 build）。手機寬度 < 900px 時是下方四個分頁（記帳／本月／鐵盒／更多），電腦版是兩欄。
- `sw.js`：service worker。畫面檔案「先給快取、背景更新」，所以發佈後第二次打開才是新版。外部字型 cache-first。
- `manifest.webmanifest`、`icons/`：加入主畫面用。manifest **故意不寫 start_url**，讓 iPhone 記住含設定的網址。
- `apps-script/Api.gs`：`doPost`（JSON 入口）和 `setupApp()`（在「設定」分頁產生新 App 連結）。
- `apps-script/Code.gs`：**參考用**。真實帳目（SEED 等）已經拿掉；正式版在使用者的 Apps Script 專案裡。
- `tests/`：在 Node 裡模擬 Apps Script（`gas-mock.js` + `gas-sim.js`），`server.js` 模擬 GitHub Pages 和會 302 轉址的 /exec。

## 連結與設定

App 連結長這樣：`https://mikadd-tw.github.io/mh-ledger/?me=M&k=<通行碼>&u=<encodeURIComponent(/exec 網址)>`

- 由試算表「設定」分頁 B12（M）、B13（H）的公式產生（B3 通行碼、B5 /exec 網址、B11 App 網址）。
- 頁面把設定存進 localStorage（`mh-pwa-cfg-v1`），**不要把網址上的參數清掉**：iPhone「加入主畫面」時會記住當下網址，主畫面 App 的儲存空間跟 Safari 是分開的。
- 沒設定時顯示「第一次使用」畫面，可以貼整條連結。

## API（`POST` 到 /exec，`Content-Type: text/plain`，body `{k, action, p}`）

回應 `{ok:true, data}` 或 `{ok:false, error}`。通行碼錯時 error 含「通行碼」。

| action | p | data |
|---|---|---|
| `ping` | — | `{pong, today}` |
| `load` | — | `{today, since, entries[], budgets[]}`（最近 400 天＋所有未結清） |
| `add` | `{entries[]}` | 依 id 去重，可重送 |
| `update` | `{entry}` | |
| `remove` | `{id}` | `{removed}` |
| `settle` | `{ids[], value}` | `{changed[]}` |
| `parse` | `{text, today}` | `{entries[], errors[]}`（LINE 文字） |

entry：`{id, date:'yyyy-MM-dd', desc, amt, pay:'現金'|'刷卡', cat:'伙食'|'休閒'|'教育'|'交通'|'雜費', who:'M'|'H'|'', settled, by}`

試算表「明細」欄位：A 日期｜B 用途｜C 金額｜D 付款｜E 分類｜F 誰出｜G 鐵盒（勾選＝已結清）｜H 記錄人｜I 建立時間｜J ID。A–G 對應舊 Excel，不要改欄位順序。

## 規則

- **公開 repo**：不要放任何帳目、金額、通行碼、/exec 網址或個人資料。測試只用 `tests/gas-sim.js` 的假資料。
- 改了 `index.html` 就更新 `APP_VERSION`（日期格式）。新增畫面要用的檔案時，加進 `sw.js` 的 `SHELL` 並把 `SHELL_CACHE` 的版本號加一。
- 前端改動推到 `main` 就會由 GitHub Pages 發佈，不用其他步驟。
- **後端不會自動部署**。Apps Script 要使用者自己貼程式、再「部署 → 管理部署作業 → 編輯 → 新版本」。所以：
  - 能只改前端就只改前端。
  - 後端改動優先放在 `Api.gs` 或新的 .gs 檔，不要叫使用者整份覆蓋 Code.gs（正式版有 SEED 資料，覆蓋會掉）。
  - 改了後端要明確告訴使用者：貼哪個檔、要不要執行哪個函式、要部署新版本。
- 權限維持 `spreadsheets.currentonly`（appsscript.json）。之前用廣的 spreadsheets 權限被 Google 擋過授權。
- 離線：新增的帳先進 localStorage 佇列（`mh-app-queue-v1`）再送；上次讀到的資料存在 `mh-pwa-cache-v1`，App 一打開先秀快取。

## 測試

```bash
npm install && npx playwright install chromium   # 第一次
node tests/backend.test.js                       # 後端（不用瀏覽器）
node tests/e2e.js                                # 手機尺寸完整流程，截圖在 tests/shots/
node tests/server.js                             # 手動試：印出可以打開的本機連結
```

改完要跑這兩個測試都過才推。
