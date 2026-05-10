/**
 * Dispatch v2 — sw.js (Service Worker)
 *
 * Strategy:
 *   App shell (HTML, CSS, JS, fonts) — Cache First
 *     Served instantly from cache on every load.
 *     Updated in background when online.
 *
 *   SAP data (Excel file from SharePoint/N8N) — Network Only
 *     Never cached — always fetched fresh.
 *     The calendar data must always be current.
 *
 *   CDN libraries (SheetJS) — Cache First with long TTL
 *     Fetched once, served from cache indefinitely.
 *     Only re-fetched when cache version bumps.
 *
 * Update strategy:
 *   Bump CACHE_VERSION when deploying new JS/CSS.
 *   Old cache deleted automatically on activate.
 *   Users get fresh code on next visit after background update.
 */

'use strict';

const CACHE_VERSION = 'dispatch-v2-1';

// App shell — everything needed to run offline
const SHELL_FILES = [
  '/',
  '/index.html',
  '/styles.css',
  '/js/parser.js',
  '/js/api.js',
  '/js/app.js',
  '/js/calendar.js',
  '/js/filters.js',
  '/js/cm-lookup.json',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png',
];

// CDN libraries — cached on first fetch
const CDN_ORIGINS = [
  'cdnjs.cloudflare.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com',
];

// N8N webhook — never cache, always network
const N8N_PATTERN = /n8n|webhook|sharepoint/i;

// ─── Install ─────────────────────────────────────────────────────────────────

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(SHELL_FILES))
      .then(() => self.skipWaiting())  // activate immediately
      .catch(err => console.error('[SW] Install failed:', err))
  );
});

// ─── Activate ────────────────────────────────────────────────────────────────

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys
          .filter(key => key !== CACHE_VERSION)
          .map(key => {
            console.log('[SW] Deleting old cache:', key);
            return caches.delete(key);
          })
      ))
      .then(() => self.clients.claim())  // take control of open pages
  );
});

// ─── Fetch ───────────────────────────────────────────────────────────────────

self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // Non-GET — pass through
  if (request.method !== 'GET') return;

  // N8N / SharePoint calls — always network, never cache
  if (N8N_PATTERN.test(url.href)) {
    event.respondWith(fetch(request));
    return;
  }

  // CDN libraries — cache first
  if (CDN_ORIGINS.some(origin => url.hostname.includes(origin))) {
    event.respondWith(cacheFirst(request));
    return;
  }

  // App shell — cache first, update in background
  if (url.origin === self.location.origin) {
    event.respondWith(staleWhileRevalidate(request));
    return;
  }

  // Everything else — network with cache fallback
  event.respondWith(networkWithCacheFallback(request));
});

// ─── Strategies ──────────────────────────────────────────────────────────────

/**
 * Cache First — serve from cache, only fetch if not cached.
 * Used for CDN libraries that rarely change.
 */
async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_VERSION);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response('Network error', { status: 503 });
  }
}

/**
 * Stale While Revalidate — serve from cache immediately,
 * fetch update in background for next visit.
 * Used for app shell — instant load, always up to date.
 */
async function staleWhileRevalidate(request) {
  const cache  = await caches.open(CACHE_VERSION);
  const cached = await cache.match(request);

  const fetchPromise = fetch(request)
    .then(response => {
      if (response.ok) cache.put(request, response.clone());
      return response;
    })
    .catch(() => null);

  return cached ?? await fetchPromise ?? new Response('Offline', { status: 503 });
}

/**
 * Network with cache fallback — try network first, fall back to cache.
 * For requests not explicitly handled above.
 */
async function networkWithCacheFallback(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_VERSION);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    return cached ?? new Response('Offline', { status: 503 });
  }
}

// ─── Message handling ─────────────────────────────────────────────────────────

/**
 * Listen for messages from the app.
 * 'SKIP_WAITING' — used to force activate a waiting service worker.
 */
self.addEventListener('message', (event) => {
  if (event.data?.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
