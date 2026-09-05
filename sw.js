var CACHE_NAME = 'chantier-v2';
var ASSETS = [
  '/chantiers-dah/',
  '/chantiers-dah/index.html',
  '/chantiers-dah/manifest.json'
];

// Install — cache app shell
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(names) {
      return Promise.all(
        names.filter(function(n) { return n !== CACHE_NAME; })
             .map(function(n) { return caches.delete(n); })
      );
    })
  );
  self.clients.claim();
});

// Fetch — network first, fallback to cache
self.addEventListener('fetch', function(e) {
  // Always go to network for API calls
  if (e.request.url.indexOf('script.google.com') !== -1) {
    return;
  }

  e.respondWith(
    fetch(e.request).then(function(res) {
      // Cache successful responses
      if (res.status === 200) {
        var clone = res.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(e.request, clone);
        });
      }
      return res;
    }).catch(function() {
      return caches.match(e.request);
    })
  );
});
