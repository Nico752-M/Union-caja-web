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

/* ===== CAPA DE DATOS =====
   Hoy: auditorio se guarda en este navegador (localStorage).
   Después: cada función llama a la base (ORDS) en vez de st.get/st.set. */
const M0=hoy().slice(0,7);
if(!st.get('cr_ing')){
 st.set('cr_ing',[{id:1,f:M0+'-03',v:1000000,d:'Evento corporativo'},{id:2,f:M0+'-05',v:800000,d:'Taller de emprendimiento'},{id:3,f:M0+'-12',v:1200000,d:'Cumpleaños de 15 años'}]);
 st.set('cr_gas',[{id:4,f:M0+'-04',v:900000,c:'Nómina',p:'Café',d:'Pago de empleadas'},{id:5,f:M0+'-08',v:310000,c:'Energía',p:'Café',d:'Factura del mes anterior'},{id:6,f:M0+'-10',v:85000,c:'Mantenimiento',p:'Auditorio',d:'Aseo del salón'}]);
}
const up=(k,o)=>{const a=st.get(k,[]);if(o.id)st.set(k,a.map(x=>x.id===o.id?o:x));else{o.id=Date.now();st.set(k,[...a,o])}};
const DATA={
 ingresos:async()=>st.get('cr_ing',[]),
 gastos:async()=>st.get('cr_gas',[]),
 saveIngreso:async o=>up('cr_ing',o),
 saveGasto:async o=>up('cr_gas',o),
 delIngreso:async id=>st.set('cr_ing',st.get('cr_ing',[]).filter(x=>x.id!==id)),
 delGasto:async id=>st.set('cr_gas',st.get('cr_gas',[]).filter(x=>x.id!==id)),
 respaldo:()=>({ing:st.get('cr_ing',[]),gas:st.get('cr_gas',[])}),
 restaurar:o=>{st.set('cr_ing',o.ing||[]);st.set('cr_gas',o.gas||[])},
 /* SOLO LECTURA, de la app Caja */
 async cafe(m){return{ventas:6420000,egresos:1150000,productos:412,dias:[{f:m+'-12',v:540000,e:60000},{f:m+'-13',v:610000,e:45000},{f:m+'-14',v:720000,e:80000}]}},
 /* SOLO LECTURA, del panel de reservas */
 async reservas(m){return[{f:m+'-03',n:'Evento corporativo',e:80,s:'Realizado'},{f:m+'-12',n:'Cumpleaños de 15 años',e:120,s:'Realizado'},{f:m+'-22',n:'Concierto acústico',e:150,s:'Próximo'}]}
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
  if(g)await DATA.saveGasto({...b,c:$('#m3',d).value,p:$('#m4',d).value});else await DATA.saveIngreso(b);
  d.close();toast(g?'Gasto guardado':'Ingreso guardado');render()};
 d.showModal();v.focus();
}

function cambiaMes(n){const[y,m]=mes().split('-').map(Number),d=new Date(y,m-1+n,1);localStorage.setItem('cr_mes',d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));$('#mt').textContent=mesTxt(mes());render()}
function shell(page,title){
 if(!sessionStorage.cr_user){location.replace('login.html');return false}
 const L=[['inicio','Inicio'],['auditorio','Auditorio'],['cafe','Café'],['reservas','Reservas'],['reportes','Reportes']];
 document.body.innerHTML=`<aside><div class="mk">Café Real<small>Centro de Negocios</small></div>
 ${L.map(([k,t])=>`<a href="${k}.html" class="${k===page?'on':''}">${t}</a>`).join('')}
 <button class="out" id="out">Cerrar sesión</button></aside>
 <div class="w"><header><div><h1>${title}</h1><p class="sb">${SUB[page]||''}</p></div>
 <div class="mn"><button id="mp" aria-label="Mes anterior">‹</button><b id="mt">${mesTxt(mes())}</b><button id="ms" aria-label="Mes siguiente">›</button><button class="g" id="mh">Este mes</button></div></header><main id="c"></main></div>`;
 $('#mp').onclick=()=>cambiaMes(-1);$('#ms').onclick=()=>cambiaMes(1);
 $('#mh').onclick=()=>{localStorage.removeItem('cr_mes');$('#mt').textContent=mesTxt(mes());render()};
 $('#out').onclick=()=>{sessionStorage.clear();location.replace('login.html')};
 return true;
}
