// Excel差分チェッカー Service Worker
// オフラインでも動作するようにアプリのアセットをキャッシュする

const CACHE_NAME = 'excel-diff-checker-v1';
const BASE_PATH = '/excel-diff-checker';

// インストール時にキャッシュするリソース
const PRECACHE_URLS = [
  BASE_PATH + '/',
  BASE_PATH + '/index.html',
  BASE_PATH + '/manifest.json',
  BASE_PATH + '/icons/icon-192x192.png',
  BASE_PATH + '/icons/icon-512x512.png',
];

// インストールイベント: 基本リソースをキャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_URLS).catch((err) => {
        console.warn('Pre-cache failed for some resources:', err);
      });
    }).then(() => self.skipWaiting())
  );
});

// アクティベートイベント: 古いキャッシュを削除
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter((name) => name !== CACHE_NAME)
          .map((name) => caches.delete(name))
      );
    }).then(() => self.clients.claim())
  );
});

// フェッチイベント: キャッシュファースト戦略
self.addEventListener('fetch', (event) => {
  // chrome-extension や非HTTPリクエストは無視
  if (!event.request.url.startsWith('http')) return;

  // POST リクエストはキャッシュしない
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // キャッシュがあればそれを返しつつ、バックグラウンドで更新
        const fetchPromise = fetch(event.request).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseClone = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseClone);
            });
          }
          return networkResponse;
        }).catch(() => cachedResponse);

        return cachedResponse;
      }

      // キャッシュがなければネットワークから取得してキャッシュに保存
      return fetch(event.request).then((networkResponse) => {
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type === 'opaque') {
          return networkResponse;
        }

        const responseClone = networkResponse.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(event.request, responseClone);
        });

        return networkResponse;
      }).catch(() => {
        // オフライン時: HTMLリクエストならindex.htmlを返す
        if (event.request.headers.get('accept') && event.request.headers.get('accept').includes('text/html')) {
          return caches.match(BASE_PATH + '/index.html');
        }
      });
    })
  );
});
