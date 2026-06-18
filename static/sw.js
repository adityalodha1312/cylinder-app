// CylinderScan Service Worker v1
const CACHE_NAME = 'cylscan-v1';
const OFFLINE_URL = '/scan';

// Assets to cache on install
const PRECACHE_ASSETS = [
  '/scan',
  '/static/manifest.json',
  'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&display=swap',
];

// ── Install: precache core assets ──────────────────────────────
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS).catch(() => {
        // Non-critical if some assets fail (e.g. external fonts offline)
      });
    }).then(() => self.skipWaiting())
  );
});

// ── Activate: clean old caches ─────────────────────────────────
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: network-first, fallback to cache ────────────────────
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Let POST requests (scan submissions) pass through normally
  if (event.request.method !== 'GET') return;

  // For navigation requests (the scan page itself): network first, then cache
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          // Cache the fresh page
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(OFFLINE_URL))
    );
    return;
  }

  // For other GET requests: stale-while-revalidate
  event.respondWith(
    caches.match(event.request).then((cached) => {
      const networkFetch = fetch(event.request).then((response) => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => cached);
      return cached || networkFetch;
    })
  );
});

// ── Offline Scan Queue (IndexedDB) ────────────────────────────
const DB_NAME    = 'cylscan-queue';
const DB_VERSION = 1;
const STORE_NAME = 'pending-scans';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      e.target.result.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror   = (e) => reject(e.target.error);
  });
}

// Receive message from page to queue a scan
self.addEventListener('message', async (event) => {
  if (event.data && event.data.type === 'QUEUE_SCAN') {
    const db = await openDB();
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).add({
      ...event.data.payload,
      queued_at: new Date().toISOString()
    });
    await new Promise(r => { tx.oncomplete = r; tx.onerror = r; });
    // Notify the page
    const clients = await self.clients.matchAll();
    clients.forEach(c => c.postMessage({ type: 'SCAN_QUEUED' }));
  }

  if (event.data && event.data.type === 'FLUSH_QUEUE') {
    await flushQueue();
  }
});

async function flushQueue() {
  const db   = await openDB();
  const tx   = db.transaction(STORE_NAME, 'readonly');
  const all  = await new Promise((res, rej) => {
    const req = tx.objectStore(STORE_NAME).getAll();
    req.onsuccess = () => res(req.result);
    req.onerror   = () => rej(req.error);
  });

  for (const scan of all) {
    try {
      // The original payload is in 'scan', minus id/queued_at
      const payload = { ...scan };
      delete payload.id;
      delete payload.queued_at;

      const res = await fetch('/submit', { 
        method: 'POST', 
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (res.ok || res.redirected) {
        // Remove from queue
        const delTx = db.transaction(STORE_NAME, 'readwrite');
        delTx.objectStore(STORE_NAME).delete(scan.id);
      }
    } catch (e) {
      // Still offline, stop flushing
      break;
    }
  }

  // Notify page
  const clients = await self.clients.matchAll();
  clients.forEach(c => c.postMessage({ type: 'QUEUE_FLUSHED' }));
}

// Auto-flush when back online
self.addEventListener('online', flushQueue);
