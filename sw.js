// MH 記帳本 service worker：讓 App 秒開、沒網路也打得開。
// 畫面檔案用「先給快取、背景更新」：發佈新版後，第二次打開就是新版。
// 記帳資料不經過這裡（是 POST 到 Apps Script），離線時由頁面自己暫存。
const SHELL_CACHE = 'mh-ledger-shell-v1';
const FONT_CACHE = 'mh-ledger-fonts-v1';
const SHELL = ['./', 'manifest.webmanifest', 'icons/apple-touch-icon.png', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/favicon-32.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(SHELL_CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('mh-ledger-') && k !== SHELL_CACHE && k !== FONT_CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function staleWhileRevalidate(event, key, networkUrl) {
  return caches.open(SHELL_CACHE).then((cache) =>
    cache.match(key, { ignoreSearch: true }).then((cached) => {
      const network = fetch(networkUrl, { cache: 'no-cache' })
        .then((res) => { if (res && res.ok) cache.put(key, res.clone()); return res; })
        .catch(() => null);
      event.waitUntil(network);
      return cached || network.then((res) => res || new Response('離線中，也還沒有快取', { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8' } }));
    })
  );
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (url.origin === self.location.origin) {
    if (url.pathname.endsWith('/sw.js')) return;
    if (req.mode === 'navigate') {
      // 網址上的 ?me=&k=&u= 只給頁面自己讀，不送去伺服器，也不當快取的 key
      const root = new URL('./', self.registration.scope).href;
      event.respondWith(staleWhileRevalidate(event, root, root));
      return;
    }
    const clean = url.origin + url.pathname;
    event.respondWith(staleWhileRevalidate(event, clean, clean));
    return;
  }

  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.open(FONT_CACHE).then((cache) =>
        cache.match(req).then((hit) => hit || fetch(req).then((res) => { if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone()); return res; }))
      )
    );
  }
});
