async function render(){
 const m=mes(),ing=inMes(await DATA.ingresos(),m),gas=inMes(await DATA.gastos(),m);
 const tI=sum(ing),tG=sum(gas),tC=sum(gas.filter(x=>(x.p||'Café')==='Café')),cat={};gas.forEach(g=>cat[g.c]=(cat[g.c]||0)+g.v);
 const mv=[...ing.map(x=>({...x,t:'i'})),...gas.map(x=>({...x,t:'g'}))].sort((a,b)=>byF(b,a));
 $('#c').innerHTML=`<div class="band"><div><span>Entró por alquiler</span><b class="g2">${cop(tI)}</b></div><div><span>Se gastó</span><b class="r2">${cop(tG)}</b></div><div><span>Quedó</span><b>${cop(tI-tG)}</b></div></div>
 <div class="bt"><button class="p big" id="ni">+ Ingreso</button><button class="g big" id="ng">+ Gasto</button><button class="g big" id="xl">Excel del mes</button><a class="g big" href="reportes.html">Reporte en PDF</a></div>
 <section class="sec"><h2>Movimientos de ${mesTxt(m)}</h2><p>Lo más reciente va arriba. Toca «Editar» para corregir un valor.</p>
 <table>${mv.map(x=>`<tr><td style="width:44px">${x.f.slice(8)}</td><td>${x.t==='i'?esc(x.d||'Ingreso'):esc(x.c)+' <span class="ro">'+esc(x.p||'Café')+'</span>'+(x.d?'<div class="mu">'+esc(x.d)+'</div>':'')}</td><td class="n ${x.t==='i'?'ok':'ng'}">${x.t==='i'?'+':'−'}${cop(x.v)}</td><td><button class="e" data-t="${x.t}" data-id="${x.id}">Editar</button></td></tr>`).join('')||'<tr><td class="empty">Aún no hay movimientos este mes. Usa «+ Ingreso» cuando recibas un pago.</td></tr>'}</table></section>
 ${tG?`<section class="sec"><h2>Para qué se usó la plata</h2><p>De todo lo que se gastó este mes.</p><table><tr><td>Sostener el café</td><td class="n">${cop(tC)}</td></tr><tr><td>Gastos del auditorio</td><td class="n">${cop(tG-tC)}</td></tr></table><h3>Por categoría</h3>`+Object.entries(cat).map(([c,t])=>`<div style="margin-top:8px"><div style="display:flex;justify-content:space-between"><span>${esc(c)}</span><b>${cop(t)}</b></div><div class="bar"><i style="width:${Math.round(t/tG*100)}%"></i></div></div>`).join('')+'</section>':''}`;
 $('#ni').onclick=()=>abrirMov('i');$('#ng').onclick=()=>abrirMov('g');
 document.querySelectorAll('[data-id]').forEach(b=>b.onclick=()=>abrirMov(b.dataset.t,(b.dataset.t==='i'?ing:gas).find(x=>x.id===+b.dataset.id)));
 $('#xl').onclick=()=>xlsx('Auditorio-'+m,{Ingresos:[['Fecha','Descripción','Valor'],...ing.sort(byF).map(x=>[x.f,x.d||'',x.v])],Gastos:[['Fecha','Categoría','Para','Descripción','Valor'],...gas.sort(byF).map(x=>[x.f,x.c,x.p||'Café',x.d||'',x.v])]});
}
if(shell('auditorio','Auditorio'))render();