// ═══════════════════════════════════════════════════════════════════════════
// MÓDULO: DIRECTORIO — claves y nombres de centros y almacenes
// Depende de helpers globales de index.html: $, almList, DIST
// ═══════════════════════════════════════════════════════════════════════════
/* ============ DIRECTORIO ============ */
let dirScope="recurrentes";
function modDirectorio(){
  $("#moduleView").innerHTML=`
    <div class="controls">
      <div class="seg" id="dirSeg"><button data-s="recurrentes" class="on">Recurrentes</button><button data-s="todos">Todos</button></div>
      <input type="search" id="dirSearch" placeholder="Buscar clave, nombre o centro…">
    </div>
    <div class="panel"><div class="panel-head"><h2>Centros y almacenes</h2><span class="pill" id="dirCount"></span></div>
      <div class="scroll" id="dirList"></div></div>`;
  $("#dirSeg").querySelectorAll("button").forEach(b=> b.onclick=()=>{ dirScope=b.dataset.s;
    $("#dirSeg").querySelectorAll("button").forEach(x=>x.classList.toggle("on",x===b)); pintarDir(); });
  $("#dirSearch").oninput=pintarDir; pintarDir();
}
function pintarDir(){
  const q=$("#dirSearch").value.trim().toLowerCase(); let alms=almList();
  if(dirScope==="recurrentes") alms=alms.filter(a=>a.recurrente);
  if(q) alms=alms.filter(a=> a.clave.toLowerCase().includes(q)||a.desc.toLowerCase().includes(q)||(a.centro_desc||"").toLowerCase().includes(q)||(a.centro||"").toLowerCase().includes(q));
  const byC={}; alms.forEach(a=> (byC[a.centro] ||= {desc:a.centro_desc,items:[]}).items.push(a));
  $("#dirCount").textContent=`${alms.length} almacenes`;
  $("#dirList").innerHTML=Object.keys(byC).sort().map(c=>{ const g=byC[c];
    return `<div class="dir-centro">${g.desc||c} <span>${c}</span></div>`+
      g.items.sort((x,y)=>x.desc.localeCompare(y.desc)).map(a=>{ const dist=a.clave===DIST();
        return `<div class="alm-row"><span class="alm-key ${dist?'dist':''}">${a.clave}</span><span>${a.desc}</span>
          ${dist?'<span class="badge">Distribuidor</span>':(a.recurrente&&dirScope==='todos'?'<span class="badge">Recurrente</span>':'')}</div>`;}).join("");
  }).join("") || `<div class="empty">Sin coincidencias.</div>`;
}

