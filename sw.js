// Cache version - increment this to force all clients to update
var CACHE = "ch-v23";
var FILES = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){ return c.addAll(FILES); })
  );
  // Force this service worker to activate immediately
  self.skipWaiting();
});

self.addEventListener("activate", function(e){
  e.waitUntil(
    // Delete ALL old caches
    caches.keys().then(function(keys){
      return Promise.all(
        keys.map(function(k){
          if(k !== CACHE){
            console.log("Deleting old cache:", k);
            return caches.delete(k);
          }
        })
      );
    })
  );
  // Take control of all open pages immediately
  self.clients.claim();
});

self.addEventListener("fetch", function(e){
  if(e.request.url.indexOf("script.google.com") >= 0 ||
     e.request.url.indexOf("googleapis.com") >= 0) return;
  e.respondWith(
    caches.match(e.request).then(function(r){
      return r || fetch(e.request).catch(function(){
        return caches.match("./index.html");
      });
    })
  );
});
