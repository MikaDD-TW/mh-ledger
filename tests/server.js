// 本機測試伺服器：
//   - 靜態網站：http://localhost:<webPort>/mh-ledger/  （模擬 GitHub Pages 的子路徑）
//   - 假的 Apps Script：POST http://127.0.0.1:<gasPort>/macros/s/TEST/exec
//     跟真的一樣先 302 轉址到另一個網域（localhost:<gasPort>/echo），再回 JSON。
const http = require('http'), fs = require('fs'), path = require('path');
const { createSim } = require('./gas-sim');

const ROOT = path.join(__dirname, '..');
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.webmanifest': 'application/manifest+json', '.png': 'image/png', '.json': 'application/json', '.css': 'text/css' };

function start({ webPort = 0, gasPort = 0 } = {}) {
  const sim = createSim();
  const state = { fail: false, delay: 0, calls: [] };
  const results = new Map();
  let seq = 0;

  const web = http.createServer((req, res) => {
    const u = new URL(req.url, 'http://x');
    if (!u.pathname.startsWith('/mh-ledger/')) { res.writeHead(302, { Location: '/mh-ledger/' }); return res.end(); }
    let p = decodeURIComponent(u.pathname.slice('/mh-ledger/'.length)) || 'index.html';
    const file = path.join(ROOT, p);
    if (!file.startsWith(ROOT) || /(^|\/)(tests|apps-script|\.git)(\/|$)/.test(p) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); return res.end('not found'); }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    fs.createReadStream(file).pipe(res);
  });

  const gas = http.createServer((req, res) => {
    const u = new URL(req.url, 'http://x');
    if (req.method === 'POST' && u.pathname === '/macros/s/TEST/exec') {
      let body = '';
      req.on('data', (c) => { body += c; });
      req.on('end', () => {
        if (state.fail) { req.socket.destroy(); return; }
        setTimeout(() => {
          try { state.calls.push(JSON.parse(body).action); } catch (e) { state.calls.push('?'); }
          const id = String(++seq);
          results.set(id, sim.post(body));
          res.writeHead(302, { Location: 'http://localhost:' + gas.address().port + '/echo?id=' + id, 'Access-Control-Allow-Origin': '*' });
          res.end();
        }, state.delay);
      });
      return;
    }
    if (req.method === 'GET' && u.pathname === '/echo') {
      const out = results.get(u.searchParams.get('id'));
      results.delete(u.searchParams.get('id'));
      res.writeHead(out ? 200 : 404, { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*' });
      return res.end(out || '{}');
    }
    res.writeHead(404, { 'Access-Control-Allow-Origin': '*' }); res.end();
  });

  return new Promise((resolve) => {
    web.listen(webPort, () => gas.listen(gasPort, () => {
      const W = web.address().port, G = gas.address().port;
      resolve({
        sim, state,
        webUrl: 'http://localhost:' + W + '/mh-ledger/',
        execUrl: 'http://127.0.0.1:' + G + '/macros/s/TEST/exec',
        close: () => { web.close(); gas.close(); },
      });
    }));
  });
}

module.exports = { start };

if (require.main === module) {
  // node tests/server.js  → 開一個可以用手機/瀏覽器手動試的環境
  start({ webPort: 8080, gasPort: 8081 }).then((s) => {
    const { fakeEntries } = require('./gas-sim');
    const today = s.sim.run("Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd')");
    s.sim.api('add', { entries: fakeEntries(today) });
    console.log('打開：' + s.webUrl + '?me=M&k=' + s.sim.key + '&u=' + encodeURIComponent(s.execUrl));
  });
}
