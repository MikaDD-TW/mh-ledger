// node tests/e2e.js [截圖資料夾]
// 用 Playwright（Chromium，iPhone 尺寸）對假的 Apps Script 跑完整流程。
const assert = require('assert');
const path = require('path');
const fs = require('fs');
const { chromium, devices } = require('playwright');
const { start } = require('./server');
const { fakeEntries } = require('./gas-sim');

const SHOTS = process.argv[2] || path.join(__dirname, 'shots');
fs.mkdirSync(SHOTS, { recursive: true });
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitFor(fn, what, ms = 8000) {
  const t0 = Date.now();
  for (;;) {
    let v; try { v = await fn(); } catch (e) { v = false; }
    if (v) return v;
    if (Date.now() - t0 > ms) throw new Error('等太久：' + what);
    await sleep(100);
  }
}

(async () => {
  const srv = await start();
  const { sim, state } = srv;
  const today = sim.run("Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd')");
  sim.api('add', { entries: fakeEntries(today) });
  const link = (k = sim.key, me = 'M') => srv.webUrl + '?me=' + me + '&k=' + k + '&u=' + encodeURIComponent(srv.execUrl);
  const row = (desc) => sim.rows().find((r) => r[1] === desc);

  const browser = await chromium.launch();
  const phone = devices['iPhone 13'];
  const errors = [];
  const watch = (page, name) => {
    page.on('pageerror', (e) => errors.push(name + ': ' + e.message));
    page.on('console', (m) => { if (m.type() === 'error' && !/fonts\.(googleapis|gstatic)|ERR_INTERNET_DISCONNECTED|Failed to load resource/.test(m.text())) errors.push(name + ' console: ' + m.text()); });
  };
  // 外部字型在測試環境連不到，直接擋掉，免得拖慢
  const noFonts = (ctx) => ctx.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());

  // ---------- 1. 用連結打開 ----------
  const ctx = await browser.newContext({ ...phone, colorScheme: 'light' });
  await noFonts(ctx);
  const page = await ctx.newPage(); watch(page, 'main');
  await page.goto(link());
  await page.locator('#list .entry', { hasText: '測試量販' }).waitFor();
  assert.ok(await page.locator('#installTip').isVisible(), 'iPhone Safari 要顯示加入主畫面提示');
  assert.ok(await page.locator('#tabbar').isVisible(), '手機要有下方分頁');
  assert.ok(!(await page.locator('#sumCard').isVisible()), '記帳分頁不顯示本月');
  assert.strictEqual(await page.locator('#tinBadge').textContent(), '6');
  await page.screenshot({ path: path.join(SHOTS, 'phone-add.png') });

  // ---------- 2. 記一筆 ----------
  await page.fill('#desc', '測試 便當');
  await page.fill('#amt', '120');
  await page.click('#catSeg button[data-v="伙食"]');
  await page.click('#submitBtn');
  await page.locator('#toast', { hasText: '已加入：測試便當' }).waitFor();
  await waitFor(() => row('測試便當'), '新增寫進試算表');
  const r1 = row('測試便當');
  assert.deepStrictEqual([r1[2], r1[3], r1[4], r1[5], r1[7]], [120, '現金', '伙食', 'M', 'M']);
  await page.screenshot({ path: path.join(SHOTS, 'phone-added.png') });

  // ---------- 3. 修改 ----------
  await page.locator('.entry', { hasText: '測試便當' }).locator('.entry-main').click();
  assert.ok(await page.locator('#editFlag').isVisible());
  await page.fill('#amt', '130');
  await page.click('#submitBtn');
  await waitFor(() => row('測試便當') && row('測試便當')[2] === 130, '修改寫進試算表');

  // ---------- 4. 刪除再復原 ----------
  await page.locator('.entry', { hasText: '測試便當' }).locator('.del').click();
  await page.locator('#toastUndo').waitFor();
  await waitFor(() => !row('測試便當'), '刪除寫進試算表');
  await page.click('#toastUndo');
  await waitFor(() => row('測試便當'), '復原寫回試算表');

  // ---------- 5. 鐵盒單筆結清 ----------
  await page.locator('.entry', { hasText: '測試咖啡' }).locator('.tin').click();
  await waitFor(() => row('測試咖啡')[6] === true, '結清寫進試算表');

  // ---------- 6. 分頁 ----------
  for (const t of ['month', 'tin', 'more']) {
    await page.click(`#tabbar button[data-t="${t}"]`);
    await sleep(150);
    await page.screenshot({ path: path.join(SHOTS, 'phone-' + t + '.png') });
  }
  await page.click('#tabbar button[data-t="month"]');
  assert.ok(await page.locator('#sumCard').isVisible());
  assert.ok(!(await page.locator('#entry').isVisible()));
  await page.click('#tabbar button[data-t="more"]');
  assert.match(await page.locator('#mConn').textContent(), /已連線/);
  await page.click('#tabbar button[data-t="add"]');

  // ---------- 7. 斷線時記帳，恢復後自動補送 ----------
  state.fail = true;
  await page.fill('#desc', '晚餐');
  await page.fill('#amt', '200');
  await page.click('#catSeg button[data-v="伙食"]');
  await page.click('#submitBtn');
  await page.locator('.entry', { hasText: '晚餐' }).locator('.pill.sync', { hasText: '未同步' }).waitFor();
  await page.locator('#banner', { hasText: '還沒存進試算表' }).waitFor();
  assert.ok(!row('晚餐'));
  await page.screenshot({ path: path.join(SHOTS, 'phone-offline-queue.png') });
  state.fail = false;
  await page.click('#bannerBtn');
  await waitFor(() => row('晚餐'), '補送寫進試算表');
  await page.locator('.entry', { hasText: '晚餐' }).locator('.pill.sync').waitFor({ state: 'detached' });

  // ---------- 8. 不帶連結重開（像主畫面 App）：先秀快取，再更新 ----------
  sim.api('add', { entries: [{ id: 'fromH', date: today, desc: 'H記的早餐', amt: 80, pay: '現金', cat: '伙食', who: 'H', by: 'H' }] });
  state.delay = 2500;
  await page.goto(srv.webUrl);
  await page.locator('#list .entry', { hasText: '測試量販' }).waitFor({ timeout: 1500 }); // 快取，不用等試算表
  assert.ok(!(await page.locator('#list .entry', { hasText: 'H記的早餐' }).count()), '這時還是快取');
  await page.locator('#list .entry', { hasText: 'H記的早餐' }).waitFor({ timeout: 6000 });
  state.delay = 0;

  // ---------- 9. 完全離線重開：service worker 給畫面，頁面給快取 ----------
  await waitFor(() => page.evaluate(() => navigator.serviceWorker.controller !== null), 'service worker 接管');
  await ctx.setOffline(true);
  await page.reload();
  await page.locator('#list .entry', { hasText: '測試量販' }).waitFor({ timeout: 3000 });
  await page.locator('#banner', { hasText: '連不上試算表，先顯示' }).waitFor();
  await page.screenshot({ path: path.join(SHOTS, 'phone-offline-open.png') });
  await ctx.setOffline(false);
  await page.click('#bannerBtn');
  await page.locator('#banner').waitFor({ state: 'hidden' });

  // ---------- 10. 深色模式 ----------
  const dctx = await browser.newContext({ ...phone, colorScheme: 'dark' });
  await noFonts(dctx);
  const dp = await dctx.newPage(); watch(dp, 'dark');
  await dp.goto(link(sim.key, 'H'));
  await dp.locator('#list .entry', { hasText: '測試量販' }).waitFor();
  await dp.screenshot({ path: path.join(SHOTS, 'phone-dark.png') });
  assert.strictEqual(await dp.locator('#whoSeg button[aria-pressed="true"]').getAttribute('data-v'), 'H');

  // ---------- 11. 通行碼錯 ----------
  const bctx = await browser.newContext(phone);
  await noFonts(bctx);
  const bp = await bctx.newPage(); watch(bp, 'badkey');
  await bp.goto(link('wrongkey'));
  await bp.locator('#badLink').waitFor();
  await bp.screenshot({ path: path.join(SHOTS, 'phone-badkey.png') });
  await bp.click('#badFix');
  assert.ok(await bp.locator('#setup').isVisible());

  // ---------- 12. 沒有連結：設定畫面，貼上連結 ----------
  const sctx = await browser.newContext(phone);
  await noFonts(sctx);
  const sp = await sctx.newPage(); watch(sp, 'setup');
  await sp.goto(srv.webUrl);
  assert.ok(await sp.locator('#setup').isVisible());
  assert.ok(!(await sp.locator('#tabbar').isVisible()));
  await sp.screenshot({ path: path.join(SHOTS, 'phone-setup.png') });
  await sp.fill('#setupIn', 'abc');
  await sp.click('#setupGo');
  assert.ok(await sp.locator('#setupErr').isVisible());
  await sp.fill('#setupIn', link());
  await sp.click('#setupGo');
  await sp.locator('#list .entry', { hasText: '測試量販' }).waitFor();

  // ---------- 13. 電腦版 ----------
  const wctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  await noFonts(wctx);
  const wp = await wctx.newPage(); watch(wp, 'desktop');
  await wp.goto(link());
  await wp.locator('#list .entry', { hasText: '測試量販' }).waitFor();
  assert.ok(!(await wp.locator('#tabbar').isVisible()));
  assert.ok(await wp.locator('#sumCard').isVisible() && await wp.locator('#settleCard').isVisible() && await wp.locator('#moreCard').isVisible());
  await wp.screenshot({ path: path.join(SHOTS, 'desktop.png'), fullPage: true });

  // ---------- 14. 沒有水平捲動 ----------
  for (const [p, n] of [[page, 'main'], [dp, 'dark'], [sp, 'setup']]) {
    const over = await p.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(over <= 0, n + ' 有水平捲動 ' + over);
  }

  await browser.close();
  srv.close();
  if (errors.length) { console.error(errors.join('\n')); process.exit(1); }
  console.log('e2e passed; screenshots in ' + SHOTS);
})().catch((e) => { console.error(e); process.exit(1); });
