// Minimal service worker — PWA "Ana Ekrana Ekle" için gerekli.
// Network-first stratejisi: güncel index.html'i her zaman sunucudan al,
// çevrimdışı kalındığında cache'ten servis et.
const CACHE = 'bmk-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-512.png',
  './logo.png',
  './html2pdf.bundle.min.js',
  './qrcode.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).catch(() => {}));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Apps Script çağrıları cache'lenmesin (origin farklı olsa da güvenlik için)
  if (url.host.includes('script.google.com') || url.host.includes('googleusercontent.com')) return;
  // Network-first: güncel sürüm için
  e.respondWith(
    fetch(req)
      .then(resp => {
        if (resp && resp.status === 200 && resp.type === 'basic') {
          const clone = resp.clone();
          caches.open(CACHE).then(c => c.put(req, clone)).catch(() => {});
        }
        return resp;
      })
      .catch(() => caches.match(req).then(r => r || caches.match('./index.html')))
  );
});
