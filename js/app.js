/* ===== PWA: ícono, color y service worker ===== */
(()=>{const a=(t,o)=>document.head.appendChild(Object.assign(document.createElement(t),o));
 a('link',{rel:'manifest',href:'manifest.json'});a('link',{rel:'apple-touch-icon',href:'icons/icon-180.png'});a('meta',{name:'theme-color',content:'#12303a'});
 if('serviceWorker'in navigator)addEventListener('load',()=>navigator.serviceWorker.register('sw.js').catch(()=>{}))})();
/* ===== CONFIG: cambiar al conectar la base ===== */
const CFG={users:{admin:'cafereal'}}; // Provisional. Se reemplaza por el login real.
const CATS=['Nómina','Energía','Agua e internet','Arriendo','Mantenimiento','Otros'];
const SUB={inicio:'Cómo va el mes. Anota aquí un ingreso o un gasto del auditorio.',auditorio:'Aquí anotas cada pago que recibes por alquilar el salón y los gastos que pagas con esa plata.',cafe:'Ventas y egresos de la caja del café. Solo se consulta, no se edita aquí.',reservas:'Eventos del auditorio y entradas generadas. Solo se consulta.',reportes:'Descarga el reporte del mes en PDF o Excel.'};
const $=(s,r=document)=>r.querySelector(s);
const cop=n=>'$'+Math.round(n).toLocaleString('es-CO');
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sum=a=>a.reduce((t,x)=>t+x.v,0);
const inMes=(a,m)=>a.filter(x=>x.f.startsWith(m));
const byF=(a,b)=>a.f<b.f?-1:a.f>b.f?1:0;
const hoy=()=>{const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')};
const mes=()=>localStorage.getItem('cr_mes')||hoy().slice(0,7);
const diaDef=m=>hoy().startsWith(m)?hoy():m+'-01';
const mesTxt=m=>{const t=new Date(m+'-15').toLocaleDateString('es-CO',{month:'long',year:'numeric'});return t[0].toUpperCase()+t.slice(1)};
const st={get(k,d){try{return JSON.parse(localStorage.getItem(k))??d}catch{return d}},set(k,v){localStorage.setItem(k,JSON.stringify(v))}};
function toast(t){const e=document.createElement('div');e.className='toast';e.setAttribute('role','status');e.textContent=t;document.body.appendChild(e);setTimeout(()=>e.style.opacity=0,1800);setTimeout(()=>e.remove(),2300)}
function bajar(nombre,texto,tipo){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([texto],{type:tipo}));a.download=nombre;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
/* Excel real (.xlsx). Carga la librería la primera vez: requiere internet. */
function libXLSX(){return window.XLSX?Promise.resolve():new Promise((ok,no)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';s.onload=ok;s.onerror=no;document.head.appendChild(s)})}
async function xlsx(nombre,hojas){
 try{await libXLSX();const wb=XLSX.utils.book_new();
  for(const[n,r]of Object.entries(hojas)){const ws=XLSX.utils.aoa_to_sheet(r);ws['!cols']=r[0]?r[0].map((_,i)=>({wch:Math.min(40,Math.max(12,...r.map(f=>String(f[i]??'').length)))})):[];XLSX.utils.book_append_sheet(wb,ws,n.slice(0,31))}
  XLSX.writeFile(wb,nombre+'.xlsx');toast('Excel descargado')}
 catch{alert('No se pudo preparar el Excel. Revisa la conexión a internet e intenta de nuevo.')}
}

/* ===== CONEXIÓN ORDS (solo lectura). Misma base, un esquema por app ===== */
const HOST='https://g473fddcae420b1-cafereal.adb.sa-bogota-1.oraclecloudapps.com/ords/';
const CAFE=HOST+'dev_usuario/',JORGE=HOST+'admin/'; // si el alias de ADMIN es otro, cámbialo aquí
const dia=f=>String(f||'').slice(0,10),sum2=(a,k)=>a.reduce((t,x)=>t+(+x[k]||0),0);
const _c={};
async function oAll(base,tabla,m,col){
 const k=base+tabla+(m||'');if(_c[k])return _c[k];
 const q=m&&col?'&q='+encodeURIComponent(JSON.stringify({[col]:{$gte:{$date:m+'-01T00:00:00Z'},$lt:{$date:new Date(+m.slice(0,4),+m.slice(5)-0,1).toISOString().slice(0,10)+'T00:00:00Z'}}})):'';
 let out=[],off=0;
 for(;;){const r=await fetch(`${base}${tabla}/?limit=500&offset=${off}${q}`,{headers:{Accept:'application/json'}});
  if(!r.ok)throw new Error(tabla+' '+r.status);const d=await r.json();out=out.concat(d.items||[]);if(!d.hasMore)break;off+=500}
 setTimeout(()=>delete _c[k],60000);return _c[k]=out}

/* ===== CAPA DE DATOS =====
   Todo viene de Oracle por ORDS: auditorio (esquema ADMIN), café (DEV_USUARIO) y reservas (ADMIN). */
const AUD=JORGE; // tablas del auditorio: esquema ADMIN
const iso=f=>f+'T00:00:00Z';
const nomItem=x=>x.producto_nombre||x.nombre_producto||x.nombre||x.producto||x.descripcion||'';
const prods=it=>{const o={};it.forEach(x=>{const r=o[x.n]=o[x.n]||{n:x.n,u:0,t:0};r.u+=x.u;r.t+=x.t});return Object.values(o).sort((a,b)=>b.u-a.u)};
const limpia=()=>Object.keys(_c).forEach(k=>delete _c[k]);
async function jr(t,m,o,id){const h={Accept:'application/json'};if(o)h['Content-Type']='application/json';const r=await fetch(AUD+t+'/'+(id??''),{method:m,headers:h,body:o?JSON.stringify(o):undefined});if(!r.ok){let x='';try{x=(await r.text()).slice(0,200)}catch{}throw new Error(t+' '+r.status+' '+x)}limpia()}
const rIng=x=>({id:x.id,f:dia(x.fecha),v:+x.valor,d:x.descripcion||''});
const rGas=x=>({...rIng(x),c:x.categoria,p:x.para});
const DATA={
 ingresos:async()=>(await oAll(AUD,'auditorio_ingresos')).map(rIng),
 gastos:async()=>(await oAll(AUD,'auditorio_gastos')).map(rGas),
 saveIngreso:async o=>jr('auditorio_ingresos',o.id?'PUT':'POST',{fecha:iso(o.f),valor:o.v,descripcion:o.d},o.id),
 saveGasto:async o=>jr('auditorio_gastos',o.id?'PUT':'POST',{fecha:iso(o.f),valor:o.v,categoria:o.c,para:o.p,descripcion:o.d},o.id),
 delIngreso:async id=>jr('auditorio_ingresos','DELETE',null,id),
 delGasto:async id=>jr('auditorio_gastos','DELETE',null,id),
 /* SOLO LECTURA: Caja Café (esquema DEV_USUARIO) */
 async cafe(m){
  const V=await oAll(CAFE,'ventas',m,'fecha'),M=await oAll(CAFE,'movimientos_caja',m,'fecha');
  const vp=V.filter(x=>x.estado==='pagada'&&dia(x.fecha).startsWith(m)),ids=new Set(vp.map(x=>x.id));
  const IT=ids.size?(await oAll(CAFE,'venta_items')).filter(x=>ids.has(x.venta_id)):[];
  const mm=M.filter(x=>dia(x.fecha).startsWith(m)),map={};
  vp.forEach(x=>{const d=dia(x.fecha);(map[d]=map[d]||{f:d,v:0,e:0}).v+=+x.total||0});
  mm.filter(x=>x.tipo==='egreso').forEach(x=>{const d=dia(x.fecha);(map[d]=map[d]||{f:d,v:0,e:0}).e+=+x.monto||0});
  const PN={};if(IT.some(x=>!nomItem(x))){try{(await oAll(CAFE,'productos')).forEach(p=>PN[p.id]=p.nombre||p.name||'')}catch{}}
  const vf={};vp.forEach(x=>vf[x.id]=dia(x.fecha));
  const items=IT.map(x=>({f:vf[x.venta_id],n:nomItem(x)||PN[x.producto_id]||('Producto '+(x.producto_id??'?')),u:+x.cantidad||0,t:+(x.subtotal??x.total??(+(x.precio_unitario??x.precio)||0)*(+x.cantidad||0))||0}));
  return{ventas:sum2(vp,'total'),egresos:sum2(mm.filter(x=>x.tipo==='egreso'),'monto'),productos:sum2(IT,'cantidad'),items,dias:Object.values(map).sort(byF)}},
 /* SOLO LECTURA: Jorge Web (esquema ADMIN) */
 async reservas(m){
  const E=(await oAll(JORGE,'eventos')).filter(x=>dia(x.fecha_inicio).startsWith(m)&&x.estado!=='CANCELADO'),R=await oAll(JORGE,'reservas_cupos'),ahora=new Date().toISOString();
  return E.map(x=>({f:dia(x.fecha_inicio),n:x.titulo,e:R.filter(r=>r.evento_id===x.id&&r.estado==='CONFIRMADA').reduce((t,r)=>t+(+r.num_personas||0),0),s:x.fecha_inicio<ahora?'Realizado':'Próximo'})).sort(byF)}
};

/* Ventanita para anotar o editar un ingreso (t='i') o un gasto (t='g') */
function abrirMov(t,o={}){
 const g=t==='g',d=document.createElement('dialog');
 const op=(a,v)=>a.map(x=>`<option${x===v?' selected':''}>${x}</option>`).join('');
 d.innerHTML=`<h2>${o.id?'Editar':'Nuevo'} ${g?'gasto':'ingreso'}</h2><form id="mf">
 <label>Fecha<input type="date" id="m1" value="${o.f||diaDef(mes())}" required></label>
 <label>Valor<input id="m2" inputmode="numeric" autocomplete="off" placeholder="0" value="${o.v?Number(o.v).toLocaleString('es-CO'):''}" required></label>
 ${g?`<label>Categoría<select id="m3">${op(CATS,o.c)}</select></label><label>¿Para quién fue?<select id="m4">${op(['Café','Auditorio'],o.p||'Café')}</select></label>`:''}
 <label>Descripción (opcional)<input id="m5" value="${esc(o.d)}" placeholder="${g?'Ej: pago de empleadas':'Ej: evento del sábado'}"></label>
 <div class="bt" style="margin:8px 0 0"><button class="p">Guardar</button><button type="button" class="g" id="mc">Cancelar</button>${o.id?'<button type="button" class="x" id="md">Borrar</button>':''}</div></form>`;
 document.body.appendChild(d);
 const v=$('#m2',d);
 v.oninput=()=>{const n=v.value.replace(/\D/g,'');v.value=n?(+n).toLocaleString('es-CO'):''};
 $('#mc',d).onclick=()=>d.close();
 d.onclose=()=>d.remove();
 if(o.id)$('#md',d).onclick=async()=>{if(confirm('¿Borrar este '+(g?'gasto':'ingreso')+'?')){try{if(g)await DATA.delGasto(o.id);else await DATA.delIngreso(o.id);d.close();toast('Borrado');render()}catch(err){alert('No se pudo borrar: '+(err&&err.message?err.message:err))}}};
 $('#mf',d).onsubmit=async e=>{e.preventDefault();const n=+v.value.replace(/\D/g,'');if(!n){v.focus();return}
  const b={id:o.id,f:$('#m1',d).value,v:n,d:$('#m5',d).value.trim()};
  try{if(g)await DATA.saveGasto({...b,c:$('#m3',d).value,p:$('#m4',d).value});else await DATA.saveIngreso(b);
  d.close();toast(g?'Gasto guardado':'Ingreso guardado');render()}catch(err){alert('No se pudo guardar: '+(err&&err.message?err.message:err))}};
 d.showModal();v.focus();
}

function cambiaMes(n){const[y,m]=mes().split('-').map(Number),d=new Date(y,m-1+n,1);localStorage.setItem('cr_mes',d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));$('#mt').textContent=mesTxt(mes());render()}
const IC=(()=>{const f=d=>`<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
 return{inicio:f('<path d="M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>'),auditorio:f('<path d="M4 20h16M6 20V10M10 20V10M14 20V10M18 20V10M3 10l9-6 9 6z"/>'),cafe:f('<path d="M5 8h12v5a5 5 0 0 1-5 5h-2a5 5 0 0 1-5-5zM17 9h1.5a2.5 2.5 0 0 1 0 5H17M7 3v2M11 3v2M15 3v2"/>'),reservas:f('<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>'),reportes:f('<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>')}})();
function shell(page,title){
 if(!sessionStorage.cr_user){location.replace('login.html');return false}
 const L=[['inicio','Inicio'],['auditorio','Auditorio'],['cafe','Café'],['reservas','Reservas'],['reportes','Reportes']];
 document.body.innerHTML=`<aside><div class="mk">Café Real<small>Centro de Negocios</small></div>
 ${L.map(([k,t])=>`<a href="${k}.html" class="${k===page?'on':''}">${IC[k]}<span>${t}</span></a>`).join('')}
 <button class="out" id="out">Cerrar sesión</button></aside>
 <div class="w"><header><div class="ht"><div><h1>${title}</h1><p class="sb">${SUB[page]||''}</p></div><button class="g sal" id="out2">Salir</button></div>
 <div class="mn"><button id="mp" aria-label="Mes anterior">‹</button><b id="mt">${mesTxt(mes())}</b><button id="ms" aria-label="Mes siguiente">›</button><button class="g" id="mh">Este mes</button></div></header><main id="c"></main></div>`;
 $('#mp').onclick=()=>cambiaMes(-1);$('#ms').onclick=()=>cambiaMes(1);
 $('#mh').onclick=()=>{localStorage.removeItem('cr_mes');$('#mt').textContent=mesTxt(mes());render()};
 $('#out').onclick=$('#out2').onclick=()=>{sessionStorage.clear();location.replace('login.html')};
 return true;
}
