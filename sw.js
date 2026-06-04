const CACHE_VERSION = 'ninja-shooting-v8.1';
self.addEventListener('install', event => { self.skipWaiting(); });
self.addEventListener('activate', event => {
  event.waitUntil((async()=>{
    const keys = await caches.keys();
    await Promise.all(keys.filter(k=>k.startsWith('ninja-shooting-') && k!==CACHE_VERSION).map(k=>caches.delete(k)));
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request, {cache:'no-store'}).catch(() => caches.match(event.request)));
});
