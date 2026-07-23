// ═══════════════════════════════════════════════════════════════════════════
// MÓDULO: ANÁLISIS — existencias por catálogo y por almacén, todos los centros
// Depende de helpers globales de index.html: $, DB, almList, almName, mat, DIST,
// almacenesConExistencia, stkClass, nfmt, descargarXLSX
// ═══════════════════════════════════════════════════════════════════════════
/* ============ ANÁLISIS (por catálogo / por almacén) ============ */
// Nombres canónicos de almacenes (sigla vieja → nombre)
const AN_NOMBRES={
  D041:'Distribuidor Puebla', A084:'Teziutlán',    A085:'Puerto Escondido',
  A08A:'San Pedro',           A08B:'Zaragoza',      A08C:'Orizaba',
  A08D:'Tlaxcala',            A08E:'Tehuacán',      A08F:'San Bruno',
  A08G:'Xalapa Cristal',      A08H:'Mocambo',       A08I:'Córdoba Dos Caminos',
  A08J:'Aquiles Serdán',      A08K:'Altamirano',    A08L:'Lerdo',
  A08M:'Peñuela',             A08X:'Atlixco',       A08Y:'Mayorazgo',
  A09O:'Oaxaca I Belisario',  A09P:'Pinotepa Nacional',
  A09Q:'Oaxaca II',           A09R:'Huajuapan',     A09T:'Cholula',
  A09U:'Amalucan',
};
function anNombre(alm){ return AN_NOMBRES[alm]||(DB.directorio?.almacenes?.[alm]?.desc)||alm; }

let anTab_="hub", anSortCatCol="dist", anSortCatDir=-1, anSortAlmCol="dist", anSortAlmDir=-1;
let anDatosCat=[], anDatosAlm=[], anCatSel=null;
let anDatosCrit=[], anSortCritCol="rest", anSortCritDir=1, anCriticoAlmSel=null;
let anBuscarCatSel=null;
let _anFijadas=[]; // filas fijadas para tabla custom

function modAnalisis(){
  const conCons=Object.keys(DB.consumos||{});
  const almsCons=conCons;
  const porCentro={}; const conEx=almacenesConExistencia();
  almList().filter(a=>conEx.has(a.clave)).forEach(a=> (porCentro[a.centro_desc||a.centro] ||= []).push(a));
  const optg=Object.keys(porCentro).sort().map(c=>`<optgroup label="${c}">${
    porCentro[c].sort((x,y)=>x.desc.localeCompare(y.desc)).map(a=>{
      const isCons=almsCons.includes(a.clave); const d=a.clave===DIST()?" — Distribuidor":"";
      return `<option value="${a.clave}" ${isCons?'style="font-weight:700"':''}>${a.clave} · ${a.desc}${d}${isCons?" ★":""}</option>`;}).join("")
  }</optgroup>`).join("");
  const mStk=DB.meta?.meses_stock||1.5;
  const almsConsCombos=almsCons.map(a=>`<option value="${a}">${a} · ${anNombre(a)}</option>`).join("");
  const tieneConsumos=almsCons.length>0;

  // ---- HUB: menú animado con las 3 vistas ----
  if(anTab_==="hub"){
    $("#moduleView").innerHTML=`
    <style>
      @keyframes anCardIn{from{opacity:0;transform:translateY(14px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
      .an-hub{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;padding:10px 2px 4px}
      .an-hub-card{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:26px 22px;
        cursor:pointer;box-shadow:var(--shadow);animation:anCardIn .38s ease both;transition:transform .15s,box-shadow .15s}
      .an-hub-card:hover{transform:translateY(-3px);box-shadow:0 12px 26px rgba(15,27,45,.14)}
      .an-hub-card .ico{font-size:34px;margin-bottom:10px;line-height:1}
      .an-hub-card h3{margin:0 0 6px;font-size:16.5px}
      .an-hub-card p{margin:0;font-size:12.5px;color:var(--muted);line-height:1.4}
    </style>
    <div class="an-hub">
      <div class="an-hub-card" data-v="cat" style="animation-delay:.02s">
        <div class="ico">📊</div><h3>Distribución por catálogo</h3>
        <p>Ver la distribución de un material entre todos los almacenes.</p>
      </div>
      <div class="an-hub-card" data-v="alm" style="animation-delay:.08s">
        <div class="ico">🏭</div><h3>Distribución por almacén</h3>
        <p>Ver qué necesita repartir D041 a un almacén específico.</p>
      </div>
      <div class="an-hub-card" data-v="critico" style="animation-delay:.14s">
        <div class="ico">⚠️</div><h3>Críticos</h3>
        <p>Materiales con menos de 1 mes de consumo propio — elige el almacén.</p>
      </div>
      <div class="an-hub-card" data-v="buscar" style="animation-delay:.20s">
        <div class="ico">🔎</div><h3>Encontrar material</h3>
        <p>Busca un catálogo y ve en qué almacenes hay existencia.</p>
      </div>
    </div>`;
    $("#moduleView").querySelectorAll(".an-hub-card").forEach(c=>c.onclick=()=>{
      anTab_=c.dataset.v;
      if(anTab_==="critico") anCriticoAlmSel=null;
      if(anTab_==="buscar") anBuscarCatSel=null;
      modAnalisis();
    });
    return;
  }


  const anTituloVista={cat:"📊 Distribución por catálogo",alm:"🏭 Distribución por almacén",critico:"⚠️ Críticos",buscar:"🔎 Encontrar material"}[anTab_]||"";
  $("#moduleView").innerHTML=`
    <div class="controls">
      <button class="btn" id="anBtnHome" title="Volver al menú">🏠 Menú</button>
      <h3 style="margin:0;font-size:15.5px;align-self:center">${anTituloVista}</h3>
      <label class="chk" id="anStkWrap" ${anTab_==="critico"?'style="display:none"':""}>Stock obj. <input type="number" id="anMStk" value="${mStk}" min="0.5" max="12" step="0.5" style="width:58px"></label>
      <span style="font-size:11.5px;color:var(--muted);align-self:center;margin-left:auto">CPM: 6 meses (12 para D041) — automático</span>
    </div>
    <!-- Panel de filas fijadas -->
    <div class="fijadas-bar" id="an-fijadas-bar">
      <h4>📌 Tabla personalizada
        <button class="btn" id="an-exp-custom" style="margin-left:auto">⬇ Exportar tabla</button>
        <button class="linkish" id="an-clear-custom" style="color:#c0392b">Limpiar todo</button>
      </h4>
      <div class="fijadas-chips" id="an-fijadas-chips"></div>
    </div>
    <!-- Panel POR CATÁLOGO -->
    <div id="an-panel-cat" ${anTab_!=="cat"?"hidden":""}>
      <div class="controls" style="margin-top:4px">
        <div class="an-cat-wrap" style="flex:1;min-width:200px">
          <input id="an-cat-input" class="an-cat-input" placeholder="Buscar catálogo por número o descripción…" autocomplete="off">
          <div id="an-cat-sugs" class="an-cat-sugs"></div>
        </div>
        <input type="search" id="an-cat-filtro" placeholder="Filtrar almacén…">
        <label class="chk"><input type="checkbox" id="an-cat-solo"> Solo con demanda</label>
        <label class="chk"><input type="checkbox" id="an-cat-rec"> Almacenes recurrentes</label>
        <button class="btn" id="an-btn-exp-cat" style="display:none">⬇ Exportar</button>
      </div>
      <div id="an-cat-desc" class="an-desc" style="font-size:13px;color:var(--muted);padding:4px 2px"></div>
      <div class="panel"><div class="panel-head"><h2 id="an-cat-title">Distribución por almacén</h2><span class="pill" id="an-cat-count"></span></div>
        <div class="scroll"><table id="an-tabla-cat"></table></div></div>
    </div>
    <!-- Panel POR ALMACÉN -->
    <div id="an-panel-alm" ${anTab_!=="alm"?"hidden":""}>
      ${!tieneConsumos?`<div class="panel"><div style="padding:16px;color:var(--muted)">
          Carga archivos de consumo en Configuración para ver el análisis por almacén.</div></div>`:`
      <div class="controls" style="margin-top:4px">
        <div class="an-cat-wrap" style="min-width:200px">
          <input id="an-alm-input" class="an-cat-input" placeholder="Buscar almacén…" autocomplete="off">
          <div id="an-alm-sugs" class="an-cat-sugs"></div>
          <input type="hidden" id="an-alm-sel">
        </div>
        <input type="search" id="an-alm-buscar" placeholder="Buscar catálogo o descripción…">
        <label class="chk"><input type="checkbox" id="an-solo-pos" checked> Solo con demanda</label>
        <button class="btn" id="an-btn-exp-alm" style="display:none">⬇ Exportar</button>
      </div>
      <div class="panel"><div class="panel-head"><h2 id="an-alm-title">Selecciona un almacén</h2><span class="pill" id="an-alm-count"></span></div>
        <div class="scroll" id="an-tabla-alm"><div class="empty" style="padding:24px">Selecciona un almacén con consumos (★) para ver el análisis.</div></div></div>`}
    </div>
    <!-- Panel CRÍTICOS (por almacén) -->
    <div id="an-panel-critico" ${anTab_!=="critico"?"hidden":""}>
      ${!tieneConsumos?`<div class="panel"><div style="padding:16px;color:var(--muted)">
          Carga archivos de consumo en Configuración para ver esta vista.</div></div>`
      : !anCriticoAlmSel ? `
      <div style="padding:2px 2px 10px;color:var(--muted);font-size:13px">Elige el almacén para ver sus materiales críticos (existencia &lt; 1 mes de su propio consumo):</div>
      <div id="an-crit-picker" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:10px"></div>`
      : `
      <div class="controls" style="margin-top:4px">
        <button class="linkish" id="an-crit-cambiar">← Cambiar almacén</button>
        <span style="font-weight:700;padding:0 4px">${anCriticoAlmSel} · ${anNombre(anCriticoAlmSel)}${anCriticoAlmSel===DIST()?" — Distribuidor":""}</span>
        <input type="search" id="an-crit-buscar" placeholder="Buscar catálogo o descripción…">
        <label class="chk"><input type="checkbox" id="an-crit-solo"> Solo críticos</label>
        <label class="chk"><input type="checkbox" id="an-crit-solo-compra"> Solo requieren compra</label>
        <button class="btn" id="an-btn-exp-crit" style="display:none;margin-left:auto">⬇ Exportar</button>
      </div>
      <div style="font-size:12px;color:var(--muted);padding:2px 2px 4px">⚠ Crítico = existencia menor a 1 mes de consumo propio (umbral fijo, no usa Stock obj.) · 🔄 Cubierto por sustituto = crítico individual, pero el Nombre Genérico completo alcanza ≥1 mes · 🛒 Requiere compra = sigue faltando aún sumando sustitutos</div>
      <div class="panel"><div class="panel-head"><h2>Materiales · Consumo propio</h2><span class="pill" id="an-crit-count"></span></div>
        <div class="scroll" id="an-tabla-crit"></div></div>`}
    </div>
    <!-- Panel ENCONTRAR MATERIAL -->
    <div id="an-panel-buscar" ${anTab_!=="buscar"?"hidden":""}>
      <div class="controls" style="margin-top:4px">
        <div class="an-cat-wrap" style="flex:1;min-width:220px">
          <input id="an-buscar-input" class="an-cat-input" placeholder="Buscar catálogo por número o descripción…" autocomplete="off">
          <div id="an-buscar-sugs" class="an-cat-sugs"></div>
        </div>
        <button class="btn" id="an-btn-exp-buscar" style="display:none">⬇ Exportar</button>
      </div>
      <div id="an-buscar-desc" class="an-desc" style="font-size:13px;color:var(--muted);padding:4px 2px">Escribe un catálogo para ver en qué almacenes hay existencia.</div>
      <div class="panel"><div class="panel-head"><h2>Almacenes con existencia</h2><span class="pill" id="an-buscar-count"></span></div>
        <div class="scroll" id="an-tabla-buscar"></div></div>
    </div>`;

  // ---- Volver al menú ----
  $("#anBtnHome").onclick=()=>{ anTab_="hub"; modAnalisis(); };
  // Panel fijadas
  _renderFijadas();
  document.getElementById("an-clear-custom")?.addEventListener("click",()=>{ _anFijadas=[]; _renderFijadas(); });
  document.getElementById("an-exp-custom")?.addEventListener("click", _exportarCustom);
  // ---- Params ----
  const repintarCat=()=>{ if(anCatSel) anRenderTablaCat(anSortCatCol); };
  const repintarAlm=()=>{ if($("#an-alm-sel")?.value) anRenderAlm(); };
  $("#anMStk").oninput=()=>{ repintarCat(); repintarAlm(); };
  // ---- Picker de almacén para Críticos ----
  if(anTab_==="critico"){
    if(!anCriticoAlmSel){
      const rec=almsCons.filter(a=>almList().find(x=>x.clave===a)?.recurrente);
      const lista=rec.length?rec:almsCons; // si ninguno de los que tienen consumo está marcado recurrente, se muestran todos
      const picker=document.getElementById("an-crit-picker");
      const conD=lista.includes(DIST());
      const resto=lista.filter(a=>a!==DIST()).sort((a,b)=>anNombre(a).localeCompare(anNombre(b)));
      const ordenada=conD?[DIST(),...resto]:resto;
      if(picker) picker.innerHTML=`
        <style>
          @keyframes anCardIn{from{opacity:0;transform:translateY(12px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
          .an-alm-card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px 16px 14px;
            cursor:pointer;box-shadow:var(--shadow);animation:anCardIn .3s ease both;position:relative;overflow:hidden;
            transition:transform .15s,box-shadow .15s,border-color .15s}
          .an-alm-card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--line)}
          .an-alm-card:hover{transform:translateY(-2px);box-shadow:0 10px 22px rgba(15,27,45,.14);border-color:var(--primary)}
          .an-alm-card .code{font-weight:800;font-size:15px;color:var(--primary)}
          .an-alm-card .name{font-size:12.5px;color:var(--muted);margin-top:3px}
          .an-alm-card.principal{background:linear-gradient(135deg,var(--primary-d),var(--primary) 70%);border-color:var(--primary-d)}
          .an-alm-card.principal::before{background:#ffd54a}
          .an-alm-card.principal .code,.an-alm-card.principal .name{color:#fff}
          .an-alm-card.principal .name{color:rgba(255,255,255,.85)}
          .an-alm-card .tag{position:absolute;top:8px;right:10px;font-size:10px;font-weight:700;
            background:#ffd54a;color:#5c3d00;padding:2px 7px;border-radius:20px}
        </style>
        ${ordenada.map((a,i)=>`
          <div class="an-alm-card ${a===DIST()?"principal":""}" data-a="${a}" style="animation-delay:${(i*.03).toFixed(2)}s">
            ${a===DIST()?'<span class="tag">★ Principal</span>':""}
            <div class="code">🏬 ${a}</div>
            <div class="name">${anNombre(a)}${a===DIST()?" — Distribuidor":""}</div>
          </div>`).join("")}`;
      picker?.querySelectorAll("[data-a]").forEach(el=>el.onclick=()=>{ anCriticoAlmSel=el.dataset.a; modAnalisis(); });
    } else {
      anRenderCriticos(anCriticoAlmSel);
      $("#an-crit-cambiar")?.addEventListener("click",()=>{ anCriticoAlmSel=null; modAnalisis(); });
      $("#an-crit-buscar")?.addEventListener("input", ()=>_pintarTablaCriticos());
      $("#an-crit-solo")?.addEventListener("change", ()=>_pintarTablaCriticos());
      $("#an-crit-solo-compra")?.addEventListener("change", ()=>_pintarTablaCriticos());
      document.getElementById("an-btn-exp-crit")?.addEventListener("click", anExportarCriticos);
    }
  }
  // ---- Buscador catálogo ----
  const catInp=$("#an-cat-input"), catSugs=$("#an-cat-sugs"), catDesc=$("#an-cat-desc");
  catInp.oninput=()=>{
    const q=catInp.value.trim().toLowerCase(); anCatSel=null; $("#an-btn-exp-cat").style.display="none";
    if(!q||q.length<2){ catSugs.style.display="none"; catDesc.textContent="Escribe para buscar"; return; }
    const hits=Object.entries(DB.materiales).filter(([c,m])=>c.startsWith(q)||m.desc.toLowerCase().includes(q)).slice(0,14);
    catSugs.innerHTML=hits.map(([c,m])=>`<div class="an-cat-sug" data-c="${c}"><b>${c}</b> · ${m.desc} <span style="color:var(--muted)">${m.um||""}</span></div>`).join("");
    catSugs.style.display=hits.length?"block":"none";
  };
  catSugs.addEventListener("pointerdown", e=>{
    const t=e.target.closest("[data-c]"); if(!t) return;
    anCatSel=t.dataset.c; const m=mat(anCatSel);
    catInp.value=anCatSel+" — "+m.desc; catSugs.style.display="none";
    catDesc.textContent=m.desc+(m.um?" · "+m.um:""); catDesc.className="an-desc ok";
    $("#an-btn-exp-cat").style.display="inline-flex";
    anDatosCat=[]; anRenderTablaCat(anSortCatCol);
  });
  document.addEventListener("click", e=>{ if(!catSugs.contains(e.target)&&e.target!==catInp) catSugs.style.display="none"; },{once:false,capture:false});
  $("#an-cat-filtro")?.addEventListener("input",()=>anRenderTablaCat(anSortCatCol));
  $("#an-cat-solo")?.addEventListener("change",()=>anRenderTablaCat(anSortCatCol));
  $("#an-cat-rec")?.addEventListener("change",()=>anRenderTablaCat(anSortCatCol));
  $("#an-btn-exp-cat")?.addEventListener("click",anExportarCat);
  // ---- Buscador "Encontrar material" ----
  if(anTab_==="buscar"){
    const bInp=$("#an-buscar-input"), bSugs=$("#an-buscar-sugs"), bDesc=$("#an-buscar-desc");
    bInp.oninput=()=>{
      const q=bInp.value.trim().toLowerCase(); anBuscarCatSel=null; $("#an-btn-exp-buscar").style.display="none";
      if(!q||q.length<2){ bSugs.style.display="none"; bDesc.textContent="Escribe un catálogo para ver en qué almacenes hay existencia."; return; }
      const hits=Object.entries(DB.materiales).filter(([c,m])=>c.startsWith(q)||m.desc.toLowerCase().includes(q)).slice(0,14);
      bSugs.innerHTML=hits.map(([c,m])=>`<div class="an-cat-sug" data-c="${c}"><b>${c}</b> · ${m.desc} <span style="color:var(--muted)">${m.um||""}</span></div>`).join("");
      bSugs.style.display=hits.length?"block":"none";
    };
    bSugs.addEventListener("pointerdown", e=>{
      const t=e.target.closest("[data-c]"); if(!t) return;
      anBuscarCatSel=t.dataset.c; const m=mat(anBuscarCatSel);
      bInp.value=anBuscarCatSel+" — "+m.desc; bSugs.style.display="none";
      bDesc.textContent=m.desc+(m.um?" · "+m.um:""); bDesc.className="an-desc ok";
      anRenderTablaBuscar();
    });
    document.addEventListener("click", e=>{ if(!bSugs.contains(e.target)&&e.target!==bInp) bSugs.style.display="none"; },{once:false,capture:false});
    document.getElementById("an-btn-exp-buscar")?.addEventListener("click", anExportarBuscar);
  }
  // ---- Buscador almacén ----
  if(tieneConsumos){
    const almInp=$("#an-alm-input"), almSugs=$("#an-alm-sugs"), almHid=$("#an-alm-sel");
    almInp.oninput=()=>{
      const q=almInp.value.trim().toLowerCase();
      const hits=almsCons.filter(a=> a.toLowerCase().includes(q)||anNombre(a).toLowerCase().includes(q));
      almSugs.innerHTML=hits.map(a=>`<div class="an-cat-sug" data-a="${a}"><b style="color:var(--primary)">${a}</b> · ${anNombre(a)}</div>`).join("");
      almSugs.style.display=hits.length?"block":"none";
    };
    almInp.addEventListener("focus",()=>{ almInp.oninput(); });
    almSugs.addEventListener("pointerdown",e=>{ const t=e.target.closest("[data-a]"); if(!t) return;
      almHid.value=t.dataset.a; almInp.value=t.dataset.a+" — "+anNombre(t.dataset.a); almSugs.style.display="none"; anRenderAlm(); });
    $("#an-alm-buscar").oninput=()=>{ if(almHid.value) anRenderAlm(); };
    $("#an-solo-pos").onchange=()=>{ if(almHid.value) anRenderAlm(); };
    $("#an-btn-exp-alm").onclick=anExportarAlm;
  }
}

// CPM matemáticamente correcto y automático, sin depender de que alguien lo recuerde ajustar:
// el consumo total de cada archivo se divide entre el número real de meses que representa.
// Hoy en día: D041 = 12 meses (su archivo trae 12 columnas de mes), todos los demás = 6 meses.
function anPerAlm(alm){ return alm===DIST() ? 12 : 6; }
function anMStk(){ return Math.max(0.1,parseFloat($("#anMStk")?.value)||1.5); }

/* ---- Análisis por catálogo (distribución entre almacenes) ---- */
function anRenderTablaCat(col){
  if(!anCatSel) return;
  if(anSortCatCol===col) anSortCatDir=-anSortCatDir; else { anSortCatCol=col; anSortCatDir=col==="alm"?1:-1; }
  const obj=anMStk(), cat=anCatSel, m=mat(cat);
  const qFilt=($("#an-cat-filtro")?.value||"").toLowerCase(), soloDem=$("#an-cat-solo")?.checked, soloRec=$("#an-cat-rec")?.checked;
  const alms=almList().map(a=>a.clave);
  anDatosCat=alms.map(alm=>{
    const cons=(DB.consumos?.[alm]?.[cat])||0;
    const ex=(DB.existencias?.[cat]?.[alm])||0;
    const per=anPerAlm(alm);
    const mens=cons/per, nec=mens*obj, dist=Math.max(0,nec-ex);
    const exD=(DB.existencias?.[cat]?.[DIST()])||0;
    return {alm,cons,mens,ex,nec,dist,exD};
  }).filter(r=>{
    if(qFilt && !(r.alm.toLowerCase().includes(qFilt)||anNombre(r.alm).toLowerCase().includes(qFilt))) return false;
    if(soloDem && r.dist<0.5) return false;
    if(soloRec && !((DB.directorio?.almacenes?.[r.alm])?.recurrente)) return false;
    return true;
  });
  anDatosCat.sort((a,b)=>{
    if(anSortCatCol==="alm") return a.alm.localeCompare(b.alm)*anSortCatDir;
    return ((a[anSortCatCol]||0)-(b[anSortCatCol]||0))*anSortCatDir;
  });
  const ic=c=>c===anSortCatCol?(anSortCatDir===1?" ▲":" ▼"):" ⇅";
  const anFmt=(n,d=0)=>n==null||isNaN(n)?"-":Number(n).toLocaleString("es-MX",{minimumFractionDigits:d,maximumFractionDigits:d});
  $("#an-cat-title").innerHTML=`${cat} · ${m.desc} <span style="color:var(--muted);font-size:13px;font-weight:400">${m.um||""}</span>`;
  $("#an-cat-count").textContent=`${anDatosCat.length} almacenes`;
  const tbl=$("#an-tabla-cat");
  tbl.innerHTML=`<thead><tr>
    <th onclick="_anSortCat('alm')">Almacén${ic("alm")}</th>
    <th class="r" onclick="_anSortCat('cons')">Consumo total${ic("cons")}</th>
    <th class="r" onclick="_anSortCat('mens')">Cons. mensual${ic("mens")}</th>
    <th class="r" onclick="_anSortCat('ex')">Existencia${ic("ex")}</th>
    <th class="r" onclick="_anSortCat('nec')">Stock obj.(${obj.toFixed(1)}m)${ic("nec")}</th>
    <th class="r" onclick="_anSortCat('dist')">Distribución${ic("dist")}</th>
    <th>Estado</th><th></th>
  </tr></thead><tbody>${anDatosCat.length? anDatosCat.map(({alm,cons,mens,ex,nec,dist})=>{
    const nc=cons===0;
    const badge=nc?`<span class="an-badge g">Sin consumo</span>`
      :dist<0.5?`<span class="an-badge v">✓ Suficiente</span>`
      :dist<=mens?`<span class="an-badge a">↑ Bajo</span>`
      :`<span class="an-badge r">⚠ Déficit</span>`;
    const yaFijada=_anFijadas.some(f=>f.cat===cat&&f.alm===alm);
    return `<tr>
      <td><div class="an-alm-code">${alm}</div><div class="an-alm-name">${anNombre(alm)}</div></td>
      <td class="r num">${nc?"—":anFmt(cons)}</td>
      <td class="r num">${nc?"—":anFmt(mens,1)}</td>
      <td class="r num">${anFmt(ex)}</td>
      <td class="r num">${nc?"—":anFmt(nec,1)}</td>
      <td class="r num" style="font-weight:700;color:${dist>0.5?"var(--low)":"var(--ok)"}">${dist>0.5?anFmt(dist,1):"—"}</td>
      <td>${badge}</td>
      <td><button class="btn an-fijar" data-cat="${cat}" data-alm="${alm}"
          data-desc="${(m.desc||"").replace(/"/g,"")}" data-um="${m.um||""}"
          data-cons="${cons}" data-ex="${ex}" data-dist="${dist}"
          style="${yaFijada?"background:var(--primary);color:#fff":""}"
          title="${yaFijada?"Ya fijada":"Fijar en tabla custom"}">📌</button></td>
    </tr>`;}).join(""):`<tr><td colspan="8" class="empty">Sin datos con los filtros actuales.</td></tr>`}</tbody>`;
  tbl.querySelectorAll(".an-fijar").forEach(btn=>btn.onclick=()=>{
    const {cat,alm,desc,um,cons,ex,dist}=btn.dataset;
    _fijarFila({cat,alm,desc,um,cons:+cons,ex:+ex,dist:+dist});
    anRenderTablaCat(anSortCatCol);
  });
}
function _fijarFila(fila){
  const idx=_anFijadas.findIndex(f=>f.cat===fila.cat&&f.alm===fila.alm);
  if(idx>=0) _anFijadas.splice(idx,1); // toggle: si ya existe, quitar
  else _anFijadas.push(fila);
  _renderFijadas();
}
function _renderFijadas(){
  const bar=document.getElementById("an-fijadas-bar");
  const chips=document.getElementById("an-fijadas-chips");
  if(!bar||!chips) return;
  if(!_anFijadas.length){ bar.classList.remove("on"); return; }
  bar.classList.add("on");
  chips.innerHTML=_anFijadas.map((f,i)=>`
    <div class="fijada-chip">
      <span><b>${f.cat}</b> · ${f.alm} <span style="color:var(--muted);font-size:11px">${(f.desc||"").substring(0,25)}</span></span>
      <button onclick="_quitarFijada(${i})" title="Quitar">×</button>
    </div>`).join("");
}
function _quitarFijada(i){ _anFijadas.splice(i,1); _renderFijadas(); }
function _exportarCustom(){
  if(!_anFijadas.length){ alert("No hay filas fijadas."); return; }
  const obj=anMStk();
  const aoa=[
    [`Tabla personalizada · Stock obj. ${obj}m · CPM: 6 meses (12 para D041)`],
    [],
    ["Catálogo","Descripción","U.M.","Almacén","Nombre almacén","Consumo total",
     "Cons. mensual","Existencia",`Stock obj.(${obj}m)`,"Distribución","D041 disponible","Meses restantes","Crítico"],
    ..._anFijadas.map(f=>{
      const m=mat(f.cat);
      const cons=(DB.consumos?.[f.alm]?.[f.cat])||0;
      const per=anPerAlm(f.alm);
      const mens=cons/per, nec=mens*obj;
      const ex=DB.existencias?.[f.cat]?.[f.alm]||0;
      const dist=Math.max(0,nec-ex);
      const exD=(DB.existencias?.[f.cat]?.[DIST()])||0;
      const restan=mens>0 ? ex/mens : null;
      return [f.cat,m.desc||f.desc,m.um||f.um,f.alm,anNombre(f.alm),
              cons,+(mens.toFixed(2)),ex,
              +(nec.toFixed(2)),+(dist.toFixed(1)),exD,
              restan==null?"":+(restan.toFixed(2)), restan!=null&&restan<1?"Sí":"No"];
    })
  ];
  descargarXLSX(aoa,"Distribución",`Analisis_Custom_${fechaTag()}`);
}
function _anSortCat(col){ anRenderTablaCat(col); }

/* ---- Análisis por almacén ---- */
function anRenderAlm(){
  const alm=$("#an-alm-sel")?.value; if(!alm) return;
  const per=anPerAlm(alm), obj=anMStk();
  const cAlm=DB.consumos?.[alm]||{}, eAlm={};
  // existencias del almacén
  for(const [cat,e] of Object.entries(DB.existencias)) if(alm in e) eAlm[cat]=e[alm];
  const q=($("#an-alm-buscar")?.value||"").toLowerCase();
  const soloDem=$("#an-solo-pos")?.checked;
  const D=DIST(), eD={};
  for(const [cat,e] of Object.entries(DB.existencias)) if(D in e) eD[cat]=e[D];
  anDatosAlm=Object.keys(cAlm).filter(cat=>{
    const m=mat(cat);
    if(q && !(cat.includes(q)||(m.desc||"").toLowerCase().includes(q))) return false;
    const dist=Math.max(0,(cAlm[cat]/per)*obj-(eAlm[cat]||0));
    if(soloDem && dist<0.5) return false;
    return true;
  }).map(cat=>{
    const m=mat(cat), cons=cAlm[cat]||0, ex=eAlm[cat]||0;
    const mens=cons/per, nec=mens*obj, dist=Math.max(0,nec-ex), exD=eD[cat]||0;
    return {cat,desc:m.desc,um:m.um,area:m.area,cons,mens,ex,nec,dist,exD};
  }).sort((a,b)=>b.dist-a.dist);
  anSortAlmCol="dist"; anSortAlmDir=-1;
  $("#an-alm-title").innerHTML=`${alm} · ${anNombre(alm)}`;
  _pintarTablaAlm(obj);
}
function _pintarTablaAlm(obj){
  const col_=anSortAlmCol, dir_=anSortAlmDir;
  const ic=c=>c===col_?(dir_===1?" ▲":" ▼"):" ⇅";
  const anFmt=(n,d=0)=>n==null||isNaN(n)?"-":Number(n).toLocaleString("es-MX",{minimumFractionDigits:d,maximumFractionDigits:d});
  const o=obj||anMStk();
  const filas=[...anDatosAlm].sort((a,b)=>{
    if(["cat","desc","area"].includes(col_)) return String(a[col_]||"").localeCompare(String(b[col_]||""))*dir_;
    return ((a[col_]||0)-(b[col_]||0))*dir_;
  });
  const alm=$("#an-alm-sel")?.value||"";
  const wrap=document.getElementById("an-tabla-alm");
  const btnExp=document.getElementById("an-btn-exp-alm");
  if(!filas.length){ wrap.innerHTML=`<div class="empty" style="padding:24px">Sin coincidencias.</div>`; if(btnExp) btnExp.style.display="none"; return; }
  $("#an-alm-count").textContent=`${filas.length} catálogos`;
  wrap.innerHTML=`<div class="scroll"><table>
    <thead><tr>
      <th onclick="_anSortAlm('cat')">Catálogo${ic("cat")}</th>
      <th onclick="_anSortAlm('desc')">Descripción${ic("desc")}</th>
      <th class="r">UM</th>
      <th class="r" onclick="_anSortAlm('mens')">Cons. mensual${ic("mens")}</th>
      <th class="r" onclick="_anSortAlm('ex')">Existencia ${alm}${ic("ex")}</th>
      <th class="r" onclick="_anSortAlm('nec')">Stock obj.(${o.toFixed(1)}m)${ic("nec")}</th>
      <th class="r" onclick="_anSortAlm('dist')">Distribución${ic("dist")}</th>
      <th class="r" onclick="_anSortAlm('exD')">D041 disponible${ic("exD")}</th>
      <th>D041</th>
    </tr></thead><tbody>${filas.map(({cat,desc,um,mens,ex,nec,dist,exD})=>{
      const badge=dist<0.5?""
        :exD>=dist?`<span class="an-badge v">✓</span>`:`<span class="an-badge r">Insuf.</span>`;
      return `<tr>
        <td class="cat num">${cat}</td>
        <td style="font-size:12px;max-width:220px">${desc||""}</td>
        <td class="r" style="font-size:11px;color:var(--muted)">${um||""}</td>
        <td class="r num">${anFmt(mens,1)}</td>
        <td class="r num">${anFmt(ex)}</td>
        <td class="r num">${anFmt(nec,1)}</td>
        <td class="r num" style="font-weight:700;color:${dist>0.5?"var(--low)":"var(--ok)"}">${dist>0.5?anFmt(dist,1):"—"}</td>
        <td class="r num">${anFmt(exD)}</td>
        <td>${badge}</td></tr>`;
    }).join("")}</tbody></table></div>`;
  if(btnExp) btnExp.style.display="inline-flex";
}
function _anSortAlm(col){
  if(col===anSortAlmCol) anSortAlmDir=-anSortAlmDir; else { anSortAlmCol=col; anSortAlmDir=["cat","desc","area"].includes(col)?1:-1; }
  _pintarTablaAlm(anMStk());
}
function anExportarCat(){
  if(!anCatSel) return; const cat=anCatSel, m=mat(cat), obj=anMStk();
  // Usar exactamente los mismos filtros y datos que anRenderTablaCat (r.mens ya viene calculado con el CPM correcto por almacén)
  const rows=anDatosCat.map(r=>[r.alm,anNombre(r.alm),r.cons,+(r.mens.toFixed(2)),r.ex,+(r.nec.toFixed(2)),+(r.dist.toFixed(1)),r.exD]);
  descargarXLSX([
    [`Catálogo: ${cat} — ${m.desc||""}`,""],[`CPM: 6 meses (12 para D041) | Stock obj: ${obj}m`,""],
    [],["Almacén","Nombre","Consumo total","Cons. mensual","Existencia",`Stock obj.(${obj}m)`,"Distribución","D041 disponible"],
    ...rows
  ],"Distribución",`Dist_${cat}_${fechaTag()}`);
}
/* ---- Consumo propio por almacén: global de materiales, críticos marcados (<1 mes) ---- */
function anRenderCriticos(alm){
  const per=anPerAlm(alm);
  const cAlm=DB.consumos?.[alm]||{};
  const eAlm={};
  for(const [cat,e] of Object.entries(DB.existencias)) if(alm in e) eAlm[cat]=e[alm];

  // Existencia total por Nombre Genérico (grupo de sustitutos) en este almacén, para poder
  // restarle la propia y saber cuánto hay disponible en sus sustitutos.
  const totalPorNG={};
  for(const c of Object.keys(DB.materiales)){
    const ngc=ngDe(c);
    if(!ngc) continue; // sin NG = no tiene sustitutos registrados
    totalPorNG[ngc]=(totalPorNG[ngc]||0)+(eAlm[c]||0);
  }

  // Base = TODO el maestro de materiales (no solo los que tienen consumo registrado para este almacén),
  // así se ve el global y se marca claramente "Sin consumo" donde no haya dato en vez de simplemente omitirlos.
  anDatosCrit=Object.keys(DB.materiales).map(cat=>{
    const m=mat(cat);
    const tieneCons=cAlm[cat]!==undefined;
    const cons=tieneCons?Math.max(0,cAlm[cat]||0):0; // consumos negativos (ajustes/errores del origen) se tratan como sin consumo, no como demanda
    const ex=eAlm[cat]||0;
    const mens=tieneCons ? cons/per : null; // null = sin archivo de consumo para este material en este almacén
    const restan=(mens!==null && mens>0) ? ex/mens : null; // null = sin consumo o consumo en 0, no aplica "crítico"
    const critico=restan!==null && restan<1; // umbral fijo: menos de 1 mes de consumo propio en existencia

    const ng=ngDe(cat)||"";
    const exGrupoTotal=ng ? (totalPorNG[ng]||0) : ex; // sin NG, su "grupo" es solo él mismo
    const exSustitutos=ng ? Math.max(0,exGrupoTotal-ex) : 0;
    const restanConSustituto=(mens!==null && mens>0) ? exGrupoTotal/mens : null;
    // Sigue siendo crítico DESPUÉS de contar sustitutos = de verdad hay que comprarlo.
    // Si es crítico solo (ex propia baja) pero el sustituto lo cubre, no urge comprar.
    const necesitaComprar=critico && (restanConSustituto===null || restanConSustituto<1);

    return {cat,desc:m.desc,um:m.um,area:m.area,tieneCons,cons,mens,ex,restan,critico,
            ng,exSustitutos,restanConSustituto,necesitaComprar};
  });
  anSortCritCol="rest"; anSortCritDir=1;
  _pintarTablaCriticos();
}
function _filasCriticos(){
  const col_=anSortCritCol, dir_=anSortCritDir;
  const q=($("#an-crit-buscar")?.value||"").toLowerCase();
  const soloCrit=$("#an-crit-solo")?.checked;
  const soloCompra=$("#an-crit-solo-compra")?.checked;
  return anDatosCrit.filter(r=>{
    if(q && !(r.cat.includes(q) || (r.desc||"").toLowerCase().includes(q))) return false;
    if(soloCrit && !r.critico) return false;
    if(soloCompra && !r.necesitaComprar) return false;
    return true;
  }).sort((a,b)=>{
    if(["cat","desc","area"].includes(col_)) return String(a[col_]||"").localeCompare(String(b[col_]||""))*dir_;
    const k=col_==="rest"?"restan":col_;
    const av=a[k]==null?Infinity:a[k], bv=b[k]==null?Infinity:b[k];
    return (av-bv)*dir_;
  });
}
function _pintarTablaCriticos(){
  const col_=anSortCritCol, dir_=anSortCritDir;
  const ic=c=>c===col_?(dir_===1?" ▲":" ▼"):" ⇅";
  const anFmt=(n,d=0)=>n==null||isNaN(n)?"-":Number(n).toLocaleString("es-MX",{minimumFractionDigits:d,maximumFractionDigits:d});
  const wrap=document.getElementById("an-tabla-crit");
  const btnExp=document.getElementById("an-btn-exp-crit");
  if(!wrap) return;

  const filas=_filasCriticos();
  const nCrit=anDatosCrit.filter(r=>r.critico).length;
  const nCompra=anDatosCrit.filter(r=>r.necesitaComprar).length;
  const alm=anCriticoAlmSel;

  if(!filas.length){
    wrap.innerHTML=`<div class="empty" style="padding:24px">Sin coincidencias.</div>`;
    if(btnExp) btnExp.style.display="none";
    return;
  }
  $("#an-crit-count").textContent=`${filas.length} mostrados · ${nCrit} críticos · ${nCompra} requieren compra · ${Object.keys(DB.materiales).length} en el maestro`;
  wrap.innerHTML=`<div class="scroll"><table>
    <thead><tr>
      <th onclick="_anSortCrit('cat')">Catálogo${ic("cat")}</th>
      <th onclick="_anSortCrit('desc')">Descripción${ic("desc")}</th>
      <th class="r">UM</th>
      <th class="r" onclick="_anSortCrit('mens')">Cons. mensual ${alm}${ic("mens")}</th>
      <th class="r" onclick="_anSortCrit('ex')">Existencia ${alm}${ic("ex")}</th>
      <th class="r" onclick="_anSortCrit('rest')">Meses restantes${ic("rest")}</th>
      <th class="r">Sustitutos (exist.)</th>
      <th></th><th></th>
    </tr></thead><tbody>${filas.map(({cat,desc,um,cons,mens,ex,restan,critico,tieneCons,ng,exSustitutos,necesitaComprar})=>{
      const yaFijada=_anFijadas.some(f=>f.cat===cat&&f.alm===alm);
      let badge="";
      if(necesitaComprar) badge='<span class="an-badge r">🛒 Comprar</span>';
      else if(critico) badge='<span class="an-badge" style="background:var(--dist-bg,#fdf6e3);color:var(--dist,#b8860b)">🔄 Cubierto por sustituto</span>';
      return `<tr ${necesitaComprar?'style="background:var(--low-bg)"':""}>
        <td class="cat num">${cat}</td>
        <td style="font-size:12px;max-width:260px">${desc||""}</td>
        <td class="r" style="font-size:11px;color:var(--muted)">${um||""}</td>
        <td class="r num" style="${tieneCons?"":"color:var(--muted);font-style:italic"}">${tieneCons?anFmt(mens,1):"Sin consumo"}</td>
        <td class="r num">${anFmt(ex)}</td>
        <td class="r num" style="${necesitaComprar?"font-weight:700;color:var(--low)":(tieneCons?"":"color:var(--muted);font-style:italic")}">${!tieneCons?"Sin consumo":(restan==null?"—":anFmt(restan,2))}</td>
        <td class="r num" style="${!ng?"color:var(--muted);font-style:italic":""}">${!ng?"Sin sustituto":anFmt(exSustitutos)}</td>
        <td>${badge}</td>
        <td><button class="btn an-fijar-crit" data-cat="${cat}" data-alm="${alm}"
            data-desc="${(desc||"").replace(/"/g,"")}" data-um="${um||""}" data-cons="${cons}" data-ex="${ex}"
            style="${yaFijada?"background:var(--primary);color:#fff":""}"
            title="${yaFijada?"Ya fijada":"Fijar en tabla personalizada"}">📌</button></td>
      </tr>`;}).join("")}</tbody></table></div>`;
  if(btnExp) btnExp.style.display="inline-flex";
  wrap.querySelectorAll(".an-fijar-crit").forEach(btn=>btn.onclick=()=>{
    const {cat,alm,desc,um,cons,ex}=btn.dataset;
    _fijarFila({cat,alm,desc,um,cons:+cons,ex:+ex,dist:0});
    _pintarTablaCriticos();
  });
}
function _anSortCrit(col){
  if(col===anSortCritCol) anSortCritDir=-anSortCritDir; else { anSortCritCol=col; anSortCritDir=["cat","desc","area"].includes(col)?1:-1; }
  _pintarTablaCriticos();
}
function anExportarCriticos(){
  if(!anDatosCrit.length) return;
  const per=anPerAlm(anCriticoAlmSel), alm=anCriticoAlmSel;
  const soloCrit=$("#an-crit-solo")?.checked, soloCompra=$("#an-crit-solo-compra")?.checked;
  const filas=_filasCriticos();
  descargarXLSX([
    [`Materiales ${alm} · ${anNombre(alm)} — Consumo propio${soloCompra?" (solo requieren compra)":soloCrit?" (solo críticos)":" (global)"}`,""],
    [`CPM: ${per} meses (${alm===DIST()?"D041 → 12 meses":"6 meses"}) | Crítico = existencia < 1 mes de consumo propio | Requiere compra = sigue faltando aún sumando existencia de sustitutos (mismo Nombre Genérico)`,""],
    [],["Catálogo","Descripción","UM","Consumo mensual","Existencia","Meses restantes","Crítico","Existencia sustitutos","Requiere compra"],
    ...filas.map(r=>[r.cat,r.desc,r.um,
      r.tieneCons?+(r.mens.toFixed(2)):"Sin consumo",
      r.ex,
      !r.tieneCons?"Sin consumo":(r.restan==null?"":+(r.restan.toFixed(2))),
      r.critico?"Sí":"No",
      r.ng?r.exSustitutos:"Sin sustituto",
      r.necesitaComprar?"Sí":"No"])
  ],"Consumo propio",`Criticos_${alm}_${fechaTag()}`);
}

/* ---- Encontrar material: en qué almacenes hay existencia de un catálogo ---- */
let anDatosBuscar=[];
function anRenderTablaBuscar(){
  const cat=anBuscarCatSel; if(!cat) return;
  const eCat=DB.existencias?.[cat]||{};
  anDatosBuscar=Object.entries(eCat).filter(([alm,ex])=>ex>0)
    .map(([alm,ex])=>({alm,nombre:anNombre(alm),ex}))
    .sort((a,b)=>b.ex-a.ex);
  _pintarTablaBuscar();
}
function _pintarTablaBuscar(){
  const wrap=document.getElementById("an-tabla-buscar");
  const btnExp=document.getElementById("an-btn-exp-buscar");
  if(!wrap) return;
  if(!anDatosBuscar.length){
    wrap.innerHTML=`<div class="empty" style="padding:24px">Sin existencia registrada en ningún almacén para este catálogo.</div>`;
    if(btnExp) btnExp.style.display="none";
    return;
  }
  const total=anDatosBuscar.reduce((s,r)=>s+r.ex,0);
  const anFmt=(n)=>Number(n).toLocaleString("es-MX");
  $("#an-buscar-count").textContent=`${anDatosBuscar.length} almacenes · ${anFmt(total)} en total`;
  wrap.innerHTML=`<div class="scroll"><table>
    <thead><tr><th>Almacén</th><th>Nombre</th><th class="r">Existencia</th></tr></thead>
    <tbody>${anDatosBuscar.map(({alm,nombre,ex})=>`
      <tr ${alm===DIST()?'style="background:var(--ok-bg,#e7f4ec)"':""}>
        <td class="cat num" style="font-weight:${alm===DIST()?700:400}">${alm}${alm===DIST()?" ★":""}</td>
        <td style="font-size:12.5px">${nombre}${alm===DIST()?" — Distribuidor":""}</td>
        <td class="r num" style="font-weight:700">${anFmt(ex)}</td>
      </tr>`).join("")}</tbody></table></div>`;
  if(btnExp) btnExp.style.display="inline-flex";
}
function anExportarBuscar(){
  if(!anDatosBuscar.length || !anBuscarCatSel) return;
  const m=mat(anBuscarCatSel);
  descargarXLSX([
    [`Catálogo: ${anBuscarCatSel} — ${m.desc||""}`,""],
    [],["Almacén","Nombre","Existencia"],
    ...anDatosBuscar.map(r=>[r.alm,r.nombre,r.ex])
  ],"Existencias",`Buscar_${anBuscarCatSel}_${fechaTag()}`);
}


