/* WITH US Society Register — service worker
   Bump CACHE_VERSION whenever you publish a new index.html so phones pick it up. */
var CACHE_VERSION = 'withus-v2';
var SHELL = ['./', 'index.html', 'manifest.webmanifest',
             'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png',
             'icons/apple-touch-icon.png', 'icons/favicon-32.png'];
var EXTERNAL = [
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.5.0/chart.umd.min.js'
];

self.addEventListener('install', function(e){
  e.waitUntil(caches.open(CACHE_VERSION).then(function(cache){
    var jobs = SHELL.map(function(u){ return cache.add(u); });
    EXTERNAL.forEach(function(u){
      jobs.push(fetch(new Request(u, { mode:'no-cors' })).then(function(r){ return cache.put(u, r); }));
    });
    return Promise.allSettled(jobs);   // one failed download never blocks installation
  }).then(function(){ return self.skipWaiting(); }));
});

self.addEventListener('activate', function(e){
  e.waitUntil(caches.keys().then(function(keys){
    return Promise.all(keys.filter(function(k){ return k !== CACHE_VERSION; }).map(function(k){ return caches.delete(k); }));
  }).then(function(){ return self.clients.claim(); }));
});

self.addEventListener('fetch', function(e){
  var req = e.request;
  if(req.method !== 'GET') return;
  var url = new URL(req.url);

  // Page itself: try the network first (so updates arrive), fall back to the saved copy offline
  if(req.mode === 'navigate'){
    e.respondWith(fetch(req).then(function(res){
      var copy = res.clone();
      caches.open(CACHE_VERSION).then(function(c){ c.put('index.html', copy); });
      return res;
    }).catch(function(){
      return caches.match('index.html').then(function(r){ return r || caches.match('./'); });
    }));
    return;
  }

  // Everything else (icons, libraries, fonts): saved copy first, refresh quietly in the background
  e.respondWith(caches.match(req).then(function(cached){
    var refresh = fetch(req).then(function(res){
      if(res && (res.status === 200 || res.type === 'opaque')){
        var copy = res.clone();
        caches.open(CACHE_VERSION).then(function(c){ c.put(req, copy); });
      }
      return res;
    }).catch(function(){ return cached; });
    return cached || refresh;
  }));
});
