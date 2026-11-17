async function render(){
 const m=mes(),c=await DATA.cafe(m),mx=Math.max(...c.dias.map(d=>d.v),1);
 $('#c').innerHTML=`<div class="band"><div><span>Ventas</span><b class="g2">${cop(c.ventas)}</b></div><div><span>Egresos de caja</span><b class="r2">${cop(c.egresos)}</b></div><div><span>Productos vendidos</span><b>${c.productos}</b></div></div>
 <section class="sec"><h2>Movimiento por día</h2><p>Viene de la app de Caja. <span class="ro">Solo lectura</span></p>
 <table><tr><th>Día</th><th></th><th class="n">Ventas</th><th class="n">Egresos</th></tr>${c.dias.map(d=>`<tr><td>${d.f.slice(8)}</td><td style="width:35%"><div class="bar"><i style="width:${Math.round(d.v/mx*100)}%"></i></div></td><td class="n ok">${cop(d.v)}</td><td class="n ng">${cop(d.e)}</td></tr>`).join('')}</table></section>
 <div class="bt"><button class="g big" id="xl">Excel del café</button><button class="g big" id="pdf">PDF del café</button></div>`;
 $('#pdf').onclick=()=>print();
 $('#xl').onclick=()=>xlsx('Cafe-'+m,{Resumen:[['Concepto','Valor'],['Ventas',c.ventas],['Egresos de caja',c.egresos],['Neto',c.ventas-c.egresos],['Productos vendidos',c.productos]],Días:[['Fecha','Ventas','Egresos'],...c.dias.map(d=>[d.f,d.v,d.e])]});
}
if(shell('cafe','Café'))render();