/* build 20261118-140327. Mise en cache complète à la 1re ouverture (hors ligne ensuite) ; mise à jour en tâche de fond. */
const V='catalogue-20261118-140327',FILES=["./", "index.html", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png", "bg.webp", "bgl.webp", "data/cat.bin", "data/img0.bin", "data/img1.bin", "data/img2.bin", "data/img3.bin", "data/img4.bin", "data/img5.bin", "data/img6.bin", "data/img7.bin", "data/img8.bin", "data/img9.bin", "data/img10.bin", "data/img11.bin", "data/img12.bin", "data/img13.bin"];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>Promise.all(FILES.map(f=>fetch(f,{cache:'reload'}).then(r=>{if(!r.ok)throw new Error(f+' '+r.status);return c.put(f==='./'?'index.html':f,r)})))).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;const u=new URL(r.url);if(u.origin!==location.origin)return;
 let f=u.pathname.replace(/^.*\/(?=data\/)/,'');if(!f.startsWith('data/'))f=u.pathname.split('/').pop()||'index.html';
 if(f==='config.json')return;if(!FILES.includes(f))return;
 e.respondWith(caches.open(V).then(c=>c.match(f).then(hit=>{const net=fetch(r).then(x=>{if(x&&x.ok)c.put(f,x.clone());return x}).catch(()=>hit);return hit||net})))});
