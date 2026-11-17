async function render(){
 const m=mes(),ing=inMes(await DATA.ingresos(),m),gas=inMes(await DATA.gastos(),m),caf=await DATA.cafe(m),rs=(await DATA.reservas(m)).filter(x=>x.f.startsWith(m));
 const s=sum(ing)-sum(gas),n=caf.ventas-caf.egresos;
 const mv=[...ing.map(x=>({...x,t:'i'})),...gas.map(x=>({...x,t:'g'}))].sort((a,b)=>byF(b,a)).slice(0,6);
 $('#c').innerHTML=`<div class="bt"><button class="p big" id="ni">+ Ingreso</button><button class="g big" id="ng">+ Gasto</button><a class="g big" href="reportes.html">Ver reporte del mes</a></div>
 <div class="cols">
 <section class="sec"><h2>Auditorio</h2><p>Alquiler del salón y gastos pagados con esa plata.</p>
  <table><tr><td>Entró</td><td class="n ok">${cop(sum(ing))}</td></tr><tr><td>Se gastó</td><td class="n ng">${cop(sum(gas))}</td></tr><tr><td><b>Quedó</b></td><td class="n ${s>=0?'ok':'ng'}">${cop(s)}</td></tr></table>
  <p style="margin:14px 0 0"><a class="lk" href="auditorio.html">Abrir auditorio</a></p></section>
 <section class="sec"><h2>Café</h2><p>Lo que registra la caja. <span class="ro">Solo lectura</span></p>
  <table><tr><td>Ventas</td><td class="n ok">${cop(caf.ventas)}</td></tr><tr><td>Egresos de caja</td><td class="n ng">${cop(caf.egresos)}</td></tr><tr><td><b>Neto</b></td><td class="n ${n>=0?'ok':'ng'}">${cop(n)}</td></tr></table>
  <p style="margin:14px 0 0"><a class="lk" href="cafe.html">Abrir café</a></p></section>
 <section class="sec"><h2>Reservas</h2><p>Eventos y entradas. <span class="ro">Solo lectura</span></p>
  <table><tr><td>Eventos del mes</td><td class="n">${rs.length}</td></tr><tr><td>Entradas generadas</td><td class="n">${rs.reduce((t,x)=>t+x.e,0)}</td></tr><tr><td>Próximos</td><td class="n">${rs.filter(x=>x.s==='Próximo').length}</td></tr></table>
  <p style="margin:14px 0 0"><a class="lk" href="reservas.html">Abrir reservas</a></p></section></div>
 <section class="sec" style="margin-top:22px"><h2>Últimos movimientos del auditorio</h2><p>${mesTxt(m)}</p>
  <table>${mv.map(x=>`<tr><td style="width:44px">${x.f.slice(8)}</td><td>${esc(x.t==='i'?x.d||'Ingreso':(x.c+(x.d?' · '+x.d:'')))}</td><td class="n ${x.t==='i'?'ok':'ng'}">${x.t==='i'?'+':'−'}${cop(x.v)}</td></tr>`).join('')||'<tr><td class="empty">Aún no hay movimientos este mes.</td></tr>'}</table></section>`;
 $('#ni').onclick=()=>abrirMov('i');$('#ng').onclick=()=>abrirMov('g');
}
if(shell('inicio','Inicio'))render();