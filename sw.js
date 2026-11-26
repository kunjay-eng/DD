// Deposit Digital Service Worker — แก้ VERSION ทุกครั้งที่ปล่อยเวอร์ชันใหม่ เพื่อให้เบราว์เซอร์ตรวจพบการอัปเดต
const VERSION = '1.12.3';
const CACHE = 'dd-' + VERSION;
const SHELL = ['./', 'index.html', 'tw.css', 'manifest.json', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  // ไม่เรียก skipWaiting เอง: รอให้ผู้ใช้กด "อัปเดต" ในแอป (ข้อความ SKIP_WAITING)
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' })))));
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k.startsWith('dd-') && k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('message', e => {
  if (e.data && e.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== location.origin) return; // ไม่แตะคำขอไป Google Apps Script
  e.respondWith(
    caches.match(r, { ignoreSearch: true }).then(hit => hit ||
      fetch(r).then(res => {
        if (res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(r, cp)); }
        return res;
      }).catch(() => r.mode === 'navigate' ? caches.match('index.html') : Response.error()))
  );
});
