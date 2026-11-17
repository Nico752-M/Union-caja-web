async function render(){
 const m=mes(),r=(await DATA.reservas(m)).filter(x=>x.f.startsWith(m));
 $('#c').innerHTML=`<div class="band"><div><span>Eventos del mes</span><b>${r.length}</b></div><div><span>Entradas generadas</span><b>${r.reduce((t,x)=>t+x.e,0)}</b></div><div><span>Próximos</span><b>${r.filter(x=>x.s==='Próximo').length}</b></div></div>
 <section class="sec"><h2>Eventos</h2><p>Viene del panel de reservas. <span class="ro">Solo lectura</span></p>
 <table><tr><th>Día</th><th>Evento</th><th class="n">Entradas</th><th>Estado</th></tr>${r.map(x=>`<tr><td>${x.f.slice(8)}</td><td>${esc(x.n)}</td><td class="n">${x.e}</td><td>${esc(x.s)}</td></tr>`).join('')||'<tr><td colspan="4" class="empty">No hay eventos este mes.</td></tr>'}</table></section>
 <div class="bt"><button class="g big" id="xl">Excel de reservas</button><button class="g big" id="pdf">PDF de reservas</button></div>`;
 $('#pdf').onclick=()=>print();
 $('#xl').onclick=()=>xlsx('Reservas-'+m,{Reservas:[['Fecha','Evento','Entradas','Estado'],...r.map(x=>[x.f,x.n,x.e,x.s])]});
}
if(shell('reservas','Reservas'))render();