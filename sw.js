/* Service worker de Union: guarda las pantallas para abrir más rápido.
   Los datos NUNCA se guardan: siempre vienen de la base. */
const V='union-v1';
const ARCH=['login.html','inicio.html','auditorio.html','cafe.html','reservas.html','reportes.html','css/app.css','js/app.js','js/inicio.js','js/auditorio.js','js/cafe.js','js/reservas.js','js/reportes.js','manifest.json','icons/icon-192.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(ARCH)).then(()=>self.skipWaiting()).catch(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==V).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
 const r=e.request;
 if(r.method!=='GET'||new URL(r.url).origin!==location.origin)return; // la base y las fuentes no se tocan
 e.respondWith(fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(V).then(c=>c.put(r,cp))}return res}).catch(()=>caches.match(r).then(m=>m||caches.match('login.html'))));
});
