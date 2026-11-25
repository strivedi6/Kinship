/* Kinship service worker: push (OneSignal) + instant opening.
   The app shell is served from the device straight away and refreshed in the background,
   so an update shows the next time the app is opened. Data calls to Google are never cached. */
try { importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js"); } catch (e) {}
const CACHE = 'kinship-v68';
const SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('kinship-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(c => c.match(req, { ignoreSearch: true }).then(hit => {
    const fresh = fetch(req).then(res => { if (res.ok) c.put(req, res.clone()); return res; });
    if (hit) { e.waitUntil(fresh.catch(() => {})); return hit; }
    return fresh.catch(() => c.match('./index.html'));
  })));
});
