var V="gastos-v1",FILES=["./","index.html","icon.png"];
self.addEventListener("install",function(e){e.waitUntil(caches.open(V).then(function(c){return c.addAll(FILES)}).then(function(){return self.skipWaiting()}))});
self.addEventListener("activate",function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.filter(function(k){return k!==V}).map(function(k){return caches.delete(k)}))}).then(function(){return self.clients.claim()}))});
self.addEventListener("fetch",function(e){
  var r=e.request;
  if(r.method!=="GET"||new URL(r.url).origin!==self.location.origin)return;
  e.respondWith(new Promise(function(ok){
    var done=false;
    function fromCache(){return caches.match(r,{ignoreSearch:true}).then(function(m){return m||caches.match("index.html")})}
    var t=setTimeout(function(){if(!done){fromCache().then(function(m){if(m&&!done){done=true;ok(m)}})}},3000);
    fetch(r).then(function(res){
      clearTimeout(t);
      if(res&&res.ok){var cp=res.clone();caches.open(V).then(function(c){c.put(r,cp)})}
      if(!done){done=true;ok(res)}
    },function(){
      clearTimeout(t);
      if(!done)fromCache().then(function(m){done=true;ok(m||Response.error())});
    });
  }));
});
