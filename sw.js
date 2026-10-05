// 走行メーター帳：オフラインでも開けるようにするキャッシュ
// 画面を更新したら CACHE の番号を上げる（例 v2）と、各iPhoneに新しい版が届きます。
const CACHE = 'meterbook-v14';
const FILES = ['./', 'index.html', 'manifest.webmanifest', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'Code.gs'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  // 画面本体は「ネット優先」：つながれば最新版（ブラウザの控えも使わない）、圏外ならキャッシュ
  if (req.mode === 'navigate') {
    e.respondWith(fetch(url.pathname + url.search, { cache: 'no-store', credentials: 'same-origin' }).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./', copy)); return res; })
      .catch(() => caches.match('./')));
    return;
  }
  // Code.gs は表示用なので常に最新、それ以外（アイコン等）はキャッシュ優先
  if (url.pathname.endsWith('.gs')) { e.respondWith(fetch(req, { cache: 'no-store' }).catch(() => caches.match(req))); return; }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});
