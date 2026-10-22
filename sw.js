var CACHE = 'withus-v31';
var SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'logo-128.png', 'icon-maskable-512.png'];
self.addEventListener('install', function(e){ e.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(SHELL); }).then(function(){ return self.skipWaiting(); })); });
self.addEventListener('activate', function(e){
  e.waitUntil(caches.keys().then(function(keys){ return Promise.all(keys.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); })); }).then(function(){ return self.clients.claim(); }));
});
self.addEventListener('fetch', function(e){
  var req = e.request;
  if(req.method !== 'GET') return;
  var host = new URL(req.url).hostname;
  if(host === 'script.google.com' || host.slice(-18) === 'googleusercontent.com') return; // shared data: always live
  if(req.mode === 'navigate'){
    // Open at once from the saved page; fetch the newest copy in the background and tell the page when it differs.
    e.respondWith(caches.match('index.html').then(function(hit){
      var net = fetch(req).then(function(r){
        if(r && r.ok){
          var cp = r.clone(), cp2 = r.clone();
          caches.open(CACHE).then(function(c){ c.put('index.html', cp); });
          if(hit){ Promise.all([hit.clone().text(), cp2.text()]).then(function(p){ if(p[0] !== p[1]) self.clients.matchAll().then(function(cl){ cl.forEach(function(c){ c.postMessage({ type:'updated' }); }); }); }).catch(function(){}); }
        }
        return r;
      }).catch(function(){ return hit; });
      return hit || net;
    }));
    return;
  }
  e.respondWith(caches.match(req).then(function(hit){
    var net = fetch(req).then(function(r){ if(r && (r.ok || r.type === 'opaque')){ var cp = r.clone(); caches.open(CACHE).then(function(c){ c.put(req, cp); }); } return r; }).catch(function(){ return hit; });
    return hit || net;
  }));
});
