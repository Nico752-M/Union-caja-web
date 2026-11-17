let D=null,H=null,LM=null;
const finMes=m=>m+'-'+String(new Date(+m.slice(0,4),+m.slice(5),0).getDate()).padStart(2,'0');
async function render(){
 const m=mes();if(m!==LM){D=m+'-01';H=finMes(m);LM=m}
 const y=D.slice(0,4),r=x=>x.f>=D&&x.f<=H,meses=[];
 let c=D.slice(0,7);while(c<=H.slice(0,7)&&meses.length<24){meses.push(c);const[a,b]=c.split('-').map(Number),n=new Date(a,b,1);c=n.getFullYear()+'-'+String(n.getMonth()+1).padStart(2,'0')}
 const I=await DATA.ingresos(),G=await DATA.gastos(),ing=I.filter(r).sort(byF),gas=G.filter(r).sort(byF);
 const cs=await Promise.all(meses.map(x=>DATA.cafe(x))),dias=cs.flatMap(x=>x.dias).filter(r),items=cs.flatMap(x=>x.items||[]).filter(r),P=prods(items);
 const caf={ventas:dias.reduce((t,x)=>t+x.v,0),egresos:dias.reduce((t,x)=>t+x.e,0),productos:items.reduce((t,x)=>t+x.u,0)};
 const rs=(await Promise.all(meses.map(x=>DATA.reservas(x)))).flat().filter(r),ent=rs.reduce((t,x)=>t+x.e,0);
 const tI=sum(ing),tG=sum(gas),tC=sum(gas.filter(x=>(x.p||'Café')==='Café')),an=[];
 for(let k=1;k<=12;k++){const mm=y+'-'+String(k).padStart(2,'0'),a=sum(inMes(I,mm)),b=sum(inMes(G,mm));an.push([mesTxt(mm).split(' ')[0],a,b,a-b])}
 const T=(h,rows,nc=[])=>`<table><tr>${h.map((x,i)=>`<th class="${nc.includes(i)?'n':''}">${x}</th>`).join('')}</tr>${rows}</table>`;
 const TOT=(span,v)=>`<tr><td colspan="${span}"><b>TOTAL</b></td><td class="n">${cop(v)}</td></tr>`;
 $('#c').innerHTML=`<div class="bt noprint" style="align-items:flex-end"><label class="mu">Desde<br><input type="date" id="d1" value="${D}"></label><label class="mu">Hasta<br><input type="date" id="d2" value="${H}"></label><button class="p big" id="pdf">Descargar PDF</button><button class="g big" id="xl">Descargar Excel</button></div>
 <p class="mu noprint">Por defecto es el mes completo; cambia «Desde» y «Hasta» para otro periodo. En el PDF elige «Guardar como PDF» como destino.</p>
 <h2 style="font:600 24px var(--serif);margin:0 0 2px">Café Real - Reporte del auditorio</h2>
 <p class="mu" style="margin:0 0 16px">Periodo: ${D} a ${H}</p>
 <section class="sec"><h2>Resumen</h2>${T(['Ingresos del salón','Gastos pagados','Saldo','Gastos para el café','Gastos del auditorio'],`<tr><td class="n ok">${cop(tI)}</td><td class="n ng">${cop(tG)}</td><td class="n">${cop(tI-tG)}</td><td class="n">${cop(tC)}</td><td class="n">${cop(tG-tC)}</td></tr>`,[0,1,2,3,4])}</section>
 <section class="sec"><h2>Ingresos del salón</h2>${T(['Fecha','Descripción','Valor'],ing.map(x=>`<tr><td>${x.f}</td><td>${esc(x.d)}</td><td class="n">${cop(x.v)}</td></tr>`).join('')+(ing.length?TOT(2,tI):'<tr><td colspan="3" class="empty">Sin ingresos en este periodo.</td></tr>'),[2])}</section>
 <section class="sec"><h2>Gastos</h2>${T(['Fecha','Gasto / motivo','Categoría','Para','Monto'],gas.map(x=>`<tr><td>${x.f}</td><td>${esc(x.d||x.c)}</td><td>${esc(x.c)}</td><td>${esc(x.p||'Café')}</td><td class="n">${cop(x.v)}</td></tr>`).join('')+(gas.length?TOT(4,tG):'<tr><td colspan="5" class="empty">Sin gastos en este periodo.</td></tr>'),[4])}</section>
 <section class="sec"><h2>Café (Caja)</h2><table><tr><td>Ventas</td><td class="n ok">${cop(caf.ventas)}</td></tr><tr><td>Egresos de caja</td><td class="n ng">${cop(caf.egresos)}</td></tr><tr><td><b>Neto</b></td><td class="n">${cop(caf.ventas-caf.egresos)}</td></tr><tr><td>Productos vendidos</td><td class="n">${caf.productos}</td></tr></table>
  <h3>Productos vendidos</h3>${T(['Producto','Unidades','Total vendido'],P.map(x=>`<tr><td>${esc(x.n)}</td><td class="n">${x.u}</td><td class="n">${cop(x.t)}</td></tr>`).join('')||'<tr><td colspan="3" class="empty">Sin ventas en este periodo.</td></tr>',[1,2])}</section>
 <section class="sec"><h2>Reservas</h2><table><tr><td>Eventos</td><td class="n">${rs.length}</td></tr><tr><td>Entradas generadas</td><td class="n">${ent}</td></tr></table></section>
 <section class="sec noprint"><h2>Auditorio mes a mes en ${y}</h2>${T(['Mes','Ingresos','Gastos','Saldo'],an.map(x=>`<tr><td>${x[0]}</td><td class="n">${cop(x[1])}</td><td class="n">${cop(x[2])}</td><td class="n ${x[3]>=0?'ok':'ng'}">${cop(x[3])}</td></tr>`).join('')+`<tr><td><b>Total</b></td><td class="n">${cop(an.reduce((t,x)=>t+x[1],0))}</td><td class="n">${cop(an.reduce((t,x)=>t+x[2],0))}</td><td class="n">${cop(an.reduce((t,x)=>t+x[3],0))}</td></tr>`,[1,2,3])}</section>`;
 const cam=()=>{const a=$('#d1').value,b=$('#d2').value;if(!a||!b)return;if(a>b){alert('«Desde» no puede ser después de «Hasta».');return}D=a;H=b;render()};
 $('#d1').onchange=cam;$('#d2').onchange=cam;
 $('#pdf').onclick=()=>print();
 $('#xl').onclick=()=>xlsx('Reporte-'+D+'_a_'+H,{Resumen:[['Café Real - Reporte del auditorio'],['Periodo',D+' a '+H],[],['Auditorio'],['Ingresos del salón',tI],['Gastos pagados',tG],['Saldo',tI-tG],['Gastos para el café',tC],['Gastos del auditorio',tG-tC],[],['Café (Caja)'],['Ventas',caf.ventas],['Egresos de caja',caf.egresos],['Neto',caf.ventas-caf.egresos],['Productos vendidos',caf.productos],[],['Reservas'],['Eventos',rs.length],['Entradas',ent]],Ingresos:[['Fecha','Descripción','Valor'],...ing.map(x=>[x.f,x.d||'',x.v]),['TOTAL','',tI]],Gastos:[['Fecha','Gasto / motivo','Categoría','Para','Monto'],...gas.map(x=>[x.f,x.d||x.c,x.c,x.p||'Café',x.v]),['TOTAL','','','',tG]],Productos:[['Producto','Unidades','Total vendido'],...P.map(x=>[x.n,x.u,x.t])],Café:[['Fecha','Ventas','Egresos'],...dias.map(d=>[d.f,d.v,d.e])],Reservas:[['Fecha','Evento','Entradas','Estado'],...rs.map(x=>[x.f,x.n,x.e,x.s])],['Año '+y]:[['Mes','Ingresos','Gastos','Saldo'],...an]});
}
if(shell('reportes','Reportes'))render();
