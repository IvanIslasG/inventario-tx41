/* ============ INVENTARIO (D041 por área, con tarjetas) ============ */
const AREA_ICON={Herramientas:"🔧",Misceláneos:"🔩",Papelería:"📄",Cables:"🔌","Ropa y Calzado":"👔",Baja:"📦"};
const TODO="__TODO__";
let invSel=null, invSort={col:"exist",dir:-1};
let invAlmacenSel=null; // almacén elegido para consultar (null = aún no se ha elegido en esta sesión)

// existencia del distribuidor para un catálogo (null si no aparece)
function existDist(cat){ const e=DB.existencias[cat]; return e && DIST() in e ? e[DIST()] : null; }
// traslado del distribuidor (0 si no hay)
function trasDist(cat){ const t=DB.traslados&&DB.traslados[cat]; return t && DIST() in t ? t[DIST()] : 0; }
// lotes reales del distribuidor (lista, vacío si no hay)
function lotesDe(cat){ return (DB.lotes&&DB.lotes[cat])||[]; }

// existencia/traslado en CUALQUIER almacén de la red (no solo el distribuidor)
function existEnAlm(cat, alm){ const e=DB.existencias[cat]; return e && alm in e ? e[alm] : null; }
function trasEnAlm(cat, alm){ const t=DB.traslados&&DB.traslados[cat]; return t && alm in t ? t[alm] : 0; }

// Construye filas desde el MAESTRO para un área (o TODO), sobre el almacén seleccionado.
// Incluye materiales sin existencia (exist=null). Ubicación y lotes solo aplican al distribuidor D041.
function matsArea(area){
  const alm = invAlmacenSel || DIST();
  const esDist = alm === DIST();
  const out=[];
  for(const [cat,m] of Object.entries(DB.materiales)){
    if(area!==TODO && (m.area||"")!==area) continue;
    out.push({cat,desc:m.desc||"",um:m.um||"",area:m.area||"Sin clasificar",ubic: esDist ? (m.ubic||"") : "",
              exist:existEnAlm(cat,alm),tras:trasEnAlm(cat,alm),lotes: esDist ? lotesDe(cat) : []});
  }
  return out;
}

function modInventario(){
  if(invAlmacenSel===null) return renderSelectorAlmacen();
  if(invSel===null) return renderAreasInv();
  return renderDetalleInv();
}

/* ---- Pantalla de selección de almacén ---- */
function renderSelectorAlmacen(){
  $("#backBtn").onclick=mostrarMenu;
  const conData=almacenesConExistencia();
  const codigos=Object.keys(DB.directorio.almacenes||{})
    .filter(c=>c===DIST()||conData.has(c))
    .sort();
  const sinDist=codigos.filter(c=>c!==DIST());
  const ordenados=[DIST(), ...sinDist];

  const cards=ordenados.map(c=>{
    const info=DB.directorio.almacenes[c]||{};
    const esDist=c===DIST();
    return `<div class="alm-card ${esDist?'alm-dist':''}" data-alm="${c}">
      ${esDist?'<span class="alm-badge">Distribuidor</span>':''}
      <div class="alm-code">${c}</div>
      <div class="alm-name">${info.desc||c}</div>
    </div>`;
  }).join("");

  $("#moduleView").innerHTML=`
    <div class="menu-h" style="margin-bottom:14px">
      <h1>Existencias</h1>
      <p>¿Qué almacén quieres consultar? <span style="color:var(--muted)">(${ordenados.length} con datos en el TOTAL cargado)</span></p>
    </div>
    <div class="controls" style="margin-bottom:16px">
      <input type="search" id="almSearch" placeholder="Buscar almacén por clave o nombre…" style="width:100%">
    </div>
    <div class="alm-grid" id="almGrid">${cards}</div>`;

  $("#almSearch").oninput=()=>{
    const q=($("#almSearch").value||"").trim().toLowerCase();
    $("#almGrid").querySelectorAll(".alm-card").forEach(el=>{
      const c=el.dataset.alm, info=DB.directorio.almacenes[c]||{};
      const match=!q || c.toLowerCase().includes(q) || (info.desc||"").toLowerCase().includes(q);
      el.style.display=match?"":"none";
    });
  };
  $("#almGrid").querySelectorAll(".alm-card").forEach(el=>{
    el.onclick=()=>{ invAlmacenSel=el.dataset.alm; invSel=null; invSort={col:"exist",dir:-1}; modInventario(); window.scrollTo(0,0); };
  });
}

/* ---- Pantalla de tarjetas de área ---- */
function renderAreasInv(){
  $("#backBtn").onclick=()=>{ invAlmacenSel=null; modInventario(); window.scrollTo(0,0); };
  const alm=invAlmacenSel||DIST();
  const esDist=alm===DIST();
  const almInfo=DB.directorio.almacenes[alm]||{};
  // conteos por área: total en maestro y con existencia en el almacén seleccionado
  const tot={}, conx={};
  for(const [cat,m] of Object.entries(DB.materiales)){
    const a=m.area||"Sin clasificar"; tot[a]=(tot[a]||0)+1;
    if(existEnAlm(cat,alm)!==null) conx[a]=(conx[a]||0)+1;
  }
  const totalConx=Object.values(conx).reduce((s,n)=>s+n,0);
  const totalMaestro=Object.keys(DB.materiales).length;
  const cards=AREAS.map(a=>`
    <div class="area-card" data-a="${a}">
      <div class="ai">${AREA_ICON[a]||"📦"}</div>
      <div class="an">${a}</div>
      <div class="ae">${nfmt(conx[a]||0)} con existencia · ${nfmt(tot[a]||0)} en maestro</div>
      <span class="cnt">${nfmt(conx[a]||0)}</span>
      <button class="exp" data-exp="${a}">⬇ Excel</button>
    </div>`).join("");
  $("#moduleView").innerHTML=`
    <button class="linkish" id="invCambiarAlm" style="margin-bottom:10px">‹ Cambiar almacén</button>
    <div class="menu-h" style="margin-bottom:14px">
      <h1>Existencias · ${alm}${esDist?" (Distribuidor)":""}</h1>
      <p>${almInfo.desc||alm}. Elige un área para consultar, contar o exportar.</p>
    </div>
    <div class="area-grid">
      ${cards}
      <div class="area-card todo" data-a="${TODO}">
        <div class="ai">📊</div>
        <div class="an">Todo el inventario</div>
        <div class="ae">${nfmt(totalConx)} con existencia · ${nfmt(totalMaestro)} en maestro</div>
        <span class="cnt" style="color:var(--primary)">${nfmt(totalConx)}</span>
        <button class="exp" data-exp="${TODO}">⬇ Exportar…</button>
      </div>
    </div>`;
  $("#invCambiarAlm").onclick=()=>{ invAlmacenSel=null; modInventario(); window.scrollTo(0,0); };
  // abrir detalle
  $("#moduleView").querySelectorAll(".area-card").forEach(c=> c.onclick=()=>{ invSel=c.dataset.a; invSort={col:"exist",dir:-1}; renderDetalleInv(); window.scrollTo(0,0); });
  // export rápido por tarjeta
  $("#moduleView").querySelectorAll(".exp").forEach(b=> b.onclick=(ev)=>{ ev.stopPropagation();
    const a=b.dataset.exp; if(a===TODO) abrirModalExp(); else exportarArea(a,false); });
}

/* ---- Detalle de un área ---- */
function renderDetalleInv(){
  const esTodo=invSel===TODO, label=esTodo?"Todo el inventario":invSel;
  // back vuelve a las tarjetas (no al menú)
  $("#backBtn").onclick=()=>{ invSel=null; modInventario(); window.scrollTo(0,0); };
  // ubicaciones del área (para el select)
  const base=matsArea(invSel);
  const ubis=[...new Set(base.map(m=>m.ubic).filter(Boolean))].sort();
  const areasEnTodo=esTodo?[...new Set(base.map(m=>m.area).filter(Boolean))].sort():[];
  $("#moduleView").innerHTML=`
    <button class="linkish" id="invVolver" style="margin-bottom:10px">‹ Volver a áreas</button>
    <div class="controls">
      <input type="search" id="invSearch" placeholder="Buscar catálogo o descripción…">
      ${esTodo
        ? `<select id="invArea"><option value="">Todas las áreas</option>${areasEnTodo.map(a=>`<option>${a}</option>`).join("")}</select>`
        : `<select id="invUbi"><option value="">Todas las ubicaciones</option>${ubis.map(u=>`<option>${u}</option>`).join("")}</select>`}
      <label class="chk"><input type="checkbox" id="invCeros" checked> Ocultar sin existencia</label>
      <label class="chk"><input type="checkbox" id="invTras"> Solo con traslado</label>
      <label class="chk"><input type="checkbox" id="invLotes"> Solo con lotes</label>
      <label class="chk"><input type="checkbox" id="invLotesCeros" checked> Ocultar lotes en 0</label>
      <button class="btn" id="invConteo">🖨 Lista de conteo</button>
      <button class="btn" id="invExport">⬇ Exportar Excel</button>
    </div>
    <div class="panel"><div class="panel-head">
      <h2>${esTodo?"📊 ":AREA_ICON[invSel]?AREA_ICON[invSel]+" ":""}${label} <span class="pill dist">${invAlmacenSel||DIST()}</span></h2>
      <span class="pill" id="invCount"></span></div>
      <div class="scroll"><table id="invTable"></table></div></div>`;
  $("#invVolver").onclick=()=>{ invSel=null; modInventario(); window.scrollTo(0,0); };
  $("#invSearch").oninput=pintarInv;
  if(esTodo) $("#invArea").onchange=pintarInv; else $("#invUbi").onchange=pintarInv;
  $("#invCeros").onchange=pintarInv;
  $("#invTras").onchange=pintarInv;
  $("#invLotes").onchange=pintarInv;
  $("#invLotesCeros").onchange=pintarInv;
  $("#invExport").onclick=()=> exportarFiltrado();
  $("#invConteo").onclick=()=> imprimirConteo();
  pintarInv();
}

// aplica los filtros del detalle a las filas del área actual
function filasInv(){
  const q=($("#invSearch")?.value||"").trim().toLowerCase();
  const ceros=$("#invCeros")?.checked;
  const soloTras=$("#invTras")?.checked;
  const ubiSel=$("#invUbi")?.value||"";
  const areaSel=$("#invArea")?.value||"";
  let rows=matsArea(invSel).filter(m=>{
    if(q && !(m.cat.toLowerCase().includes(q)||m.desc.toLowerCase().includes(q))) return false;
    if(ubiSel && m.ubic!==ubiSel) return false;
    if(areaSel && m.area!==areaSel) return false;
    if(ceros && !(m.exist!==null && m.exist>0)) return false;
    if(soloTras && !(m.tras>0)) return false;
    if($("#invLotes")?.checked && !(m.lotes&&m.lotes.length>0)) return false;
    return true;
  });
  rows.sort((a,b)=>{const c=invSort.col,k=invSort.dir;
    if(c==="exist") return ((a.exist??-1)-(b.exist??-1))*k;
    if(c==="tras") return ((a.tras||0)-(b.tras||0))*k;
    return String(a[c]||"").localeCompare(String(b[c]||""))*k;});
  return rows;
}
function pintarInv(){
  const esTodo=invSel===TODO, rows=filasInv();
  const esDist=(invAlmacenSel||DIST())===DIST();
  const conx=rows.filter(r=>r.exist!==null&&r.exist>0).length;
  $("#invCount").textContent=`${nfmt(rows.length)} materiales · ${nfmt(conx)} con existencia`;
  const colUbic = esDist ? `<th data-c="ubic">Ubicación</th>` : "";
  const cols=esTodo
    ? `<th data-c="cat">Catálogo</th><th data-c="desc">Descripción</th><th data-c="area">Área</th>${colUbic}<th data-c="exist" class="r">Existencia</th><th data-c="tras" class="r">Traslado</th>`
    : `<th data-c="cat">Catálogo</th><th data-c="desc">Descripción</th><th data-c="um">U.M.</th>${colUbic}<th data-c="exist" class="r">Existencia</th><th data-c="tras" class="r">Traslado</th>`;
  const ncols = esDist?6:5;
  $("#invTable").innerHTML=`<thead><tr>${cols}</tr></thead><tbody>${
    rows.length? rows.map(r=>{
      const ex = r.exist===null ? `<span style="color:var(--zero);font-style:italic">—</span>`
                : `<span class="stk ${stkClass(r.exist)}">${nfmt(r.exist)}</span>`;
      const tr = r.tras>0 ? `<span class="stk tras">${nfmt(r.tras)}</span>` : `<span style="color:var(--zero)">—</span>`;
      const ocultarLotesCeros = $("#invLotesCeros")?.checked;
      const lotesVisibles = ocultarLotesCeros
        ? (r.lotes||[]).filter(l=> (l.lib>0) || (l.tras>0))
        : (r.lotes||[]);
      const tieneLotes = lotesVisibles.length>0;
      const badge = tieneLotes ? ` <span class="lote-badge" data-lt="${r.cat}">📦 ${lotesVisibles.length} lote${lotesVisibles.length>1?'s':''}</span>` : "";
      const tdUbic = esDist ? `<td>${r.ubic||"—"}</td>` : "";
      const main = esTodo
        ? `<tr><td class="cat num">${r.cat}</td><td class="desc">${r.desc||"—"}${badge}</td>
             <td><span class="area-tag">${r.area||"—"}</span></td>${tdUbic}
             <td class="r">${ex} <small style="color:var(--muted)">${r.um}</small></td><td class="r">${tr}</td></tr>`
        : `<tr><td class="cat num">${r.cat}</td><td class="desc">${r.desc||"—"}${badge}</td>
             <td>${r.um||"—"}</td>${tdUbic}
             <td class="r">${ex} <small style="color:var(--muted)">${r.um}</small></td><td class="r">${tr}</td></tr>`;
      const filaLotes = tieneLotes
        ? `<tr class="lotes-row" id="lt-${r.cat}" hidden><td colspan="${ncols}" style="padding:0;background:#fffaf2">
             <table class="lotes-sub"><thead><tr><th>Lote</th><th class="r">Libre</th><th class="r">Traslado</th></tr></thead>
             <tbody>${lotesVisibles.map(l=>`<tr><td class="num">${l.lote}</td><td class="r num">${nfmt(l.lib)}</td>
               <td class="r num" style="color:var(--low)">${nfmt(l.tras||0)}</td></tr>`).join("")}</tbody></table></td></tr>`
        : "";
      return main+filaLotes;
    }).join("") : `<tr><td colspan="${ncols}" class="empty">Sin coincidencias.</td></tr>`}</tbody>`;
  $("#invTable").querySelectorAll("th[data-c]").forEach(th=> th.onclick=()=>{
    const c=th.dataset.c; invSort={col:c,dir:invSort.col===c?-invSort.dir:(c==="exist"||c==="tras"?-1:1)}; pintarInv(); });
  $("#invTable").querySelectorAll(".lote-badge").forEach(b=> b.onclick=()=>{
    const row=document.getElementById("lt-"+b.dataset.lt); if(row) row.hidden=!row.hidden; });
}

/* ---- Exportaciones de Inventario ---- */
function fechaTag(){ return new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'2-digit',year:'numeric'}).replace(/\//g,'-'); }
function expFilas(mats){
  const aoa=[["Área","Catálogo","Descripción","U.M.","Ubicación","Existencia D041","Traslado","Lote"]];
  mats.forEach(m=>{
    aoa.push([m.area,m.cat,m.desc,m.um,m.ubic, m.exist===null?"":m.exist, m.tras||0, ""]);
    if(m.lotes && m.lotes.length){
      m.lotes.forEach(l=> aoa.push(["","  ↳ lote","","","", l.lib, l.tras||0, l.lote]));
    }
  });
  return aoa;
}

// export directo de un área (desde la tarjeta)
function exportarArea(area,soloExist){
  let mats=matsArea(area); if(soloExist) mats=mats.filter(m=>m.exist!==null&&m.exist>0);
  if(!mats.length){ alert("No hay materiales para exportar en "+area); return; }
  mats.sort((a,b)=>(b.exist??-1)-(a.exist??-1));
  descargarXLSX(expFilas(mats), (area===TODO?"Todo":area).substring(0,28), `Inventario_${area===TODO?"Todo":area.replace(/ /g,"_")}_${fechaTag()}`);
}
// export del detalle respetando filtros actuales
function exportarFiltrado(){
  const mats=filasInv(); if(!mats.length){ alert("No hay materiales con los filtros actuales."); return; }
  const nom=invSel===TODO?"Todo":invSel.replace(/ /g,"_");
  descargarXLSX(expFilas(mats), (invSel===TODO?"Todo":invSel).substring(0,28), `Inventario_${nom}_${fechaTag()}`);
}

/* ---- Modal de exportación multi-área ---- */
function abrirModalExp(){
  $("#expAreas").innerHTML=AREAS.map(a=>`<label><input type="checkbox" data-a="${a}" checked> ${a}</label>`).join("");
  $("#expSoloExist").checked=true;
  $("#expModal").classList.add("on");
}
function cerrarModalExp(){ $("#expModal").classList.remove("on"); }
$("#expCancel").onclick=cerrarModalExp;
$("#expAll").onclick=()=>{ const cbs=$("#expAreas").querySelectorAll("input"); const any=[...cbs].some(c=>!c.checked); cbs.forEach(c=>c.checked=any); };
$("#expGo").onclick=()=>{
  const sel=[...$("#expAreas").querySelectorAll("input:checked")].map(c=>c.dataset.a);
  if(!sel.length){ alert("Selecciona al menos un área."); return; }
  const soloExist=$("#expSoloExist").checked;
  let mats=[]; sel.forEach(a=>{ mats=mats.concat(matsArea(a)); });
  if(soloExist) mats=mats.filter(m=>m.exist!==null&&m.exist>0);
  if(!mats.length){ alert("No hay materiales con esos criterios."); return; }
  mats.sort((a,b)=> a.area===b.area ? (b.exist??-1)-(a.exist??-1) : a.area.localeCompare(b.area));
  const nom=sel.length===AREAS.length?"Todo":sel.join("_").substring(0,28);
  descargarXLSX(expFilas(mats), "Inventario", `Inventario_${nom}_${fechaTag()}`);
  cerrarModalExp();
};

/* ---- Modal de ubicaciones para conteo ---- */
function imprimirConteo(){
  const mats=filasInv();
  if(!mats.length){ alert("No hay materiales con los filtros actuales."); return; }
  // agrupar por ubicación
  const grupos={};
  mats.forEach(m=>{ const u=m.ubic||"Sin asignar"; (grupos[u] ||= []).push(m); });
  const ubis=Object.keys(grupos).sort();
  if(ubis.length===1){
    // si solo hay una ubicación, imprimir directo
    _generarImpresion(grupos, ubis);
    return;
  }
  // modal de selección
  const modal=document.createElement("div"); modal.className="modal on"; modal.id="conteoModal";
  modal.innerHTML=`<div class="modal-box">
    <h3>🖨 Lista de conteo — Elegir ubicaciones</h3>
    <div class="modal-body">
      <div style="font-size:12px;font-weight:700;color:var(--muted);margin-bottom:7px;text-transform:uppercase;letter-spacing:.4px">
        Ubicaciones disponibles
        <button class="linkish" id="cntAll" style="float:right;text-transform:none;letter-spacing:0">Todas / ninguna</button>
      </div>
      <div style="display:grid;gap:5px;max-height:240px;overflow-y:auto" id="cntChecks">
        ${ubis.map(u=>`<label style="display:flex;align-items:center;gap:7px;font-size:13px;padding:5px 8px;border-radius:7px;background:#f7f9fc;border:1px solid var(--line);cursor:pointer">
          <input type="checkbox" data-u="${u}" checked>
          <span><b>${u}</b> <span style="color:var(--muted)">(${grupos[u].length} materiales)</span></span>
        </label>`).join("")}
      </div>
    </div>
    <div class="modal-foot">
      <button class="btn" id="cntCancel">Cancelar</button>
      <button class="btn-prim" id="cntOk">🖨 Imprimir selección</button>
    </div>
  </div>`;
  document.body.appendChild(modal);
  $("#cntAll").onclick=()=>{ const cbs=modal.querySelectorAll("input"); const any=[...cbs].some(c=>!c.checked); cbs.forEach(c=>c.checked=any); };
  $("#cntCancel").onclick=()=> modal.remove();
  $("#cntOk").onclick=()=>{
    const sel=[...modal.querySelectorAll("input:checked")].map(c=>c.dataset.u);
    if(!sel.length){ alert("Selecciona al menos una ubicación."); return; }
    modal.remove();
    _generarImpresion(grupos, sel);
  };
}
function _generarImpresion(grupos, ubis){
  const ocultarLotesCeros = $("#invLotesCeros")?.checked;
  const hoy=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
  const label=invSel===TODO?"TODO EL INVENTARIO":invSel.toUpperCase();
  const pa=$("#printArea"); pa.innerHTML="";
  ubis.forEach(u=>{
    const lista=(grupos[u]||[]).slice().sort((a,b)=>+a.cat - +b.cat); if(!lista.length) return;
    const div=document.createElement("div"); div.className="pg";
    div.innerHTML=`
      <div class="ph">
        <div class="ph-top">TELMEX — ALMACÉN DISTRIBUIDOR D041 PUEBLA · ÁREA: ${label}</div>
        <div class="ph-sub">Inventario físico · Lista de conteo</div>
        <div class="ph-ubi">Ubicación: ${u}</div>
        <div class="ph-meta">Fecha: <u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u>&nbsp;&nbsp;&nbsp;Responsable: <u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></div>
      </div>
      <table class="pt"><thead><tr>
        <th style="width:78px">Catálogo</th><th>Descripción</th>
        <th style="width:34px;text-align:center">U.M.</th>
        <th style="width:58px;text-align:center">Existencia</th>
        <th style="width:50px;text-align:center">Traslado</th>
        <th style="width:56px;text-align:center">Físico</th>
      </tr></thead><tbody>
        ${lista.map(m=>{
          const lotesVisibles = ocultarLotesCeros
            ? (m.lotes||[]).filter(l=> (l.lib>0) || (l.tras>0))
            : (m.lotes||[]);
          const base=`<tr>
            <td style="font-family:monospace">${m.cat}</td>
            <td>${m.desc||"—"}${lotesVisibles.length?` <span style="font-size:7.5pt;color:#9a5a00">[${lotesVisibles.length} lote${lotesVisibles.length>1?'s':''}]</span>`:""}</td>
            <td style="text-align:center">${m.um||"—"}</td>
            <td style="text-align:center">${m.exist===null?"—":nfmt(m.exist)}</td>
            <td style="text-align:center">${m.tras>0?nfmt(m.tras):"—"}</td>
            <td class="col-fis">&nbsp;</td></tr>`;
          const subs=lotesVisibles.length? lotesVisibles.map(l=>`<tr style="background:#fffaf2">
            <td style="font-family:monospace;font-size:7.5pt;padding-left:16px;color:#9a5a00">↳ ${l.lote}</td>
            <td style="font-size:7.5pt;color:#777;font-style:italic">Lote</td><td></td>
            <td style="text-align:center;font-size:7.5pt">${nfmt(l.lib)}</td>
            <td style="text-align:center;font-size:7.5pt">${nfmt(l.tras||0)}</td>
            <td class="col-fis">&nbsp;</td></tr>`).join(""):"";
          return base+subs;
        }).join("")}
      </tbody></table>
      <div class="pf"><span>D041 · ${invSel===TODO?"Todo":invSel} · ${u} · ${lista.length} materiales</span><span>Generado: ${hoy}</span></div>`;
    pa.appendChild(div);
  });
  setTimeout(()=>{ window.print(); setTimeout(()=>{ pa.innerHTML=""; },800); },250);
}
