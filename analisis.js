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

let anTab_="cat", anSortCatCol="dist", anSortCatDir=-1, anSortAlmCol="dist", anSortAlmDir=-1;
let anDatosCat=[], anDatosAlm=[], anCatSel=null;
let anDatosCritD041=[], anSortCritCol="rest", anSortCritDir=1;
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
  const mCPM=DB.meta?.meses_cpm||6, mStk=DB.meta?.meses_stock||1.5;
  const tieneConsD041=!!(DB.consumos?.[DIST()]);
  const almsConsCombos=almsCons.map(a=>`<option value="${a}">${a} · ${anNombre(a)}</option>`).join("");
  const tieneConsumos=almsCons.length>0;

  $("#moduleView").innerHTML=`
    <div class="controls">
      <div class="seg" id="anSeg">
        <button data-v="cat" class="${anTab_==="cat"?"on":""}">Por catálogo</button>
        <button data-v="alm" class="${anTab_==="alm"?"on":""}">Por almacén</button>
        <button data-v="critD041" class="${anTab_==="critD041"?"on":""}">⚠ Críticos D041</button>
      </div>
      <label class="chk">Meses CPM <input type="number" id="anMCPM" value="${mCPM}" min="1" max="24" step="1" style="width:58px"></label>
      <label class="chk" id="anStkWrap" ${anTab_==="critD041"?'style="display:none"':""}>Stock obj. <input type="number" id="anMStk" value="${mStk}" min="0.5" max="12" step="0.5" style="width:58px"></label>
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
    <!-- Panel CRÍTICOS D041 -->
    <div id="an-panel-critD041" ${anTab_!=="critD041"?"hidden":""}>
      ${!tieneConsD041?`<div class="panel"><div style="padding:16px;color:var(--muted)">
          Carga el archivo de consumos de <b>D041</b> en Configuración para ver esta vista (nombra el archivo con "D041").</div></div>`:`
      <div class="controls" style="margin-top:4px">
        <input type="search" id="an-crit-buscar" placeholder="Buscar catálogo o descripción…">
        <label class="chk"><input type="checkbox" id="an-crit-solo"> Solo críticos</label>
        <button class="btn" id="an-btn-exp-crit" style="display:none;margin-left:auto">⬇ Exportar</button>
      </div>
      <div style="font-size:12px;color:var(--muted);padding:2px 2px 4px">⚠ Crítico = existencia D041 menor a 1 mes de su propio consumo (umbral fijo, no usa Stock obj.)</div>
      <div class="panel"><div class="panel-head"><h2>Materiales D041 · Consumo propio</h2><span class="pill" id="an-crit-count"></span></div>
        <div class="scroll" id="an-tabla-crit"></div></div>`}
    </div>`;

  // ---- Tab ----
  $("#anSeg").querySelectorAll("button").forEach(b=> b.onclick=()=>{
    anTab_=b.dataset.v;
    ["cat","alm","critD041"].forEach(v=>{ document.getElementById("an-panel-"+v).hidden=anTab_!==v; });
    $("#anSeg").querySelectorAll("button").forEach(x=>x.classList.toggle("on",x===b));
    const stkWrap=document.getElementById("anStkWrap"); if(stkWrap) stkWrap.style.display=anTab_==="critD041"?"none":"";
    if(anTab_==="critD041" && tieneConsD041) anRenderCriticosD041();
  });
  if(anTab_==="critD041" && tieneConsD041) anRenderCriticosD041();
  // Panel fijadas
  _renderFijadas();
  document.getElementById("an-clear-custom")?.addEventListener("click",()=>{ _anFijadas=[]; _renderFijadas(); });
  document.getElementById("an-exp-custom")?.addEventListener("click", _exportarCustom);
  // ---- Params ----
  const repintarCat=()=>{ if(anCatSel) anRenderTablaCat(anSortCatCol); };
  const repintarAlm=()=>{ if($("#an-alm-sel")?.value) anRenderAlm(); };
  const repintarCrit=()=>{ if(tieneConsD041 && anTab_==="critD041") anRenderCriticosD041(); };
  $("#anMCPM").oninput=()=>{ repintarCat(); repintarAlm(); repintarCrit(); };
  $("#anMStk").oninput=()=>{ repintarCat(); repintarAlm(); };
  $("#an-crit-buscar")?.addEventListener("input", ()=>_pintarTablaCriticosD041());
  $("#an-crit-solo")?.addEventListener("change", ()=>_pintarTablaCriticosD041());
  document.getElementById("an-btn-exp-crit")?.addEventListener("click", anExportarCriticosD041);
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

function anMCPM(){ return Math.max(1,parseFloat($("#anMCPM")?.value)||6); }
function anMStk(){ return Math.max(0.1,parseFloat($("#anMStk")?.value)||1.5); }

/* ---- Análisis por catálogo (distribución entre almacenes) ---- */
function anRenderTablaCat(col){
  if(!anCatSel) return;
  if(anSortCatCol===col) anSortCatDir=-anSortCatDir; else { anSortCatCol=col; anSortCatDir=col==="alm"?1:-1; }
  const per=anMCPM(), obj=anMStk(), cat=anCatSel, m=mat(cat);
  const qFilt=($("#an-cat-filtro")?.value||"").toLowerCase(), soloDem=$("#an-cat-solo")?.checked, soloRec=$("#an-cat-rec")?.checked;
  const alms=almList().map(a=>a.clave);
  anDatosCat=alms.map(alm=>{
    const cons=(DB.consumos?.[alm]?.[cat])||0;
    const ex=(DB.existencias?.[cat]?.[alm])||0;
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
  const per=anMCPM(), obj=anMStk();
  const aoa=[
    [`Tabla personalizada · CPM ${per}m · Stock obj. ${obj}m`],
    [],
    ["Catálogo","Descripción","U.M.","Almacén","Nombre almacén","Consumo total",
     "Cons. mensual","Existencia",`Stock obj.(${obj}m)`,"Distribución","D041 disponible"],
    ..._anFijadas.map(f=>{
      const m=mat(f.cat);
      const cons=(DB.consumos?.[f.alm]?.[f.cat])||0;
      const mens=cons/per, nec=mens*obj;
      const dist=Math.max(0,nec-(DB.existencias?.[f.cat]?.[f.alm]||0));
      const exD=(DB.existencias?.[f.cat]?.[DIST()])||0;
      return [f.cat,m.desc||f.desc,m.um||f.um,f.alm,anNombre(f.alm),
              cons,+(mens.toFixed(2)),DB.existencias?.[f.cat]?.[f.alm]||0,
              +(nec.toFixed(2)),+(dist.toFixed(1)),exD];
    })
  ];
  descargarXLSX(aoa,"Distribución",`Analisis_Custom_${fechaTag()}`);
}
function _anSortCat(col){ anRenderTablaCat(col); }

/* ---- Análisis por almacén ---- */
function anRenderAlm(){
  const alm=$("#an-alm-sel")?.value; if(!alm) return;
  const per=anMCPM(), obj=anMStk();
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
  if(!anCatSel) return; const cat=anCatSel, m=mat(cat), per=anMCPM(), obj=anMStk();
  // Usar exactamente los mismos filtros y datos que anRenderTablaCat
  const rows=anDatosCat.map(r=>[r.alm,anNombre(r.alm),r.cons,+(r.mens.toFixed(2)),r.ex,+(r.nec.toFixed(2)),+(r.dist.toFixed(1)),r.exD]);
  descargarXLSX([
    [`Catálogo: ${cat} — ${m.desc||""}`,""],[`Periodo: ${per}m | Stock obj: ${obj}m`,""],
    [],["Almacén","Nombre","Consumo total","Cons. mensual","Existencia",`Stock obj.(${obj}m)`,"Distribución","D041 disponible"],
    ...rows
  ],"Distribución",`Dist_${cat}_${fechaTag()}`);
}
/* ---- Consumo propio D041: global de materiales, críticos marcados (<1 mes) ---- */
function anRenderCriticosD041(){
  const D=DIST(); // "D041"
  const per=anMCPM();
  const cD=DB.consumos?.[D]||{};
  const eD={};
  for(const [cat,e] of Object.entries(DB.existencias)) if(D in e) eD[cat]=e[D];

  anDatosCritD041=Object.keys(cD).map(cat=>{
    const m=mat(cat);
    const cons=Math.max(0,cD[cat]||0); // consumos negativos (ajustes/errores del origen) se tratan como sin consumo, no como demanda
    const ex=eD[cat]||0;
    const mens=cons/per;
    const restan=mens>0 ? ex/mens : null; // null = sin consumo registrado en el periodo, no aplica "crítico"
    const critico=restan!==null && restan<1; // umbral fijo: menos de 1 mes de consumo propio en existencia
    return {cat,desc:m.desc,um:m.um,area:m.area,cons,mens,ex,restan,critico};
  });
  anSortCritCol="rest"; anSortCritDir=1;
  _pintarTablaCriticosD041();
}
function _filasCriticosD041(){
  const col_=anSortCritCol, dir_=anSortCritDir;
  const q=($("#an-crit-buscar")?.value||"").toLowerCase();
  const soloCrit=$("#an-crit-solo")?.checked;
  return anDatosCritD041.filter(r=>{
    if(q && !(r.cat.includes(q) || (r.desc||"").toLowerCase().includes(q))) return false;
    if(soloCrit && !r.critico) return false;
    return true;
  }).sort((a,b)=>{
    if(["cat","desc","area"].includes(col_)) return String(a[col_]||"").localeCompare(String(b[col_]||""))*dir_;
    const k=col_==="rest"?"restan":col_;
    const av=a[k]==null?Infinity:a[k], bv=b[k]==null?Infinity:b[k];
    return (av-bv)*dir_;
  });
}
function _pintarTablaCriticosD041(){
  const col_=anSortCritCol, dir_=anSortCritDir;
  const ic=c=>c===col_?(dir_===1?" ▲":" ▼"):" ⇅";
  const anFmt=(n,d=0)=>n==null||isNaN(n)?"-":Number(n).toLocaleString("es-MX",{minimumFractionDigits:d,maximumFractionDigits:d});
  const wrap=document.getElementById("an-tabla-crit");
  const btnExp=document.getElementById("an-btn-exp-crit");
  if(!wrap) return;

  const filas=_filasCriticosD041();
  const nCrit=anDatosCritD041.filter(r=>r.critico).length;

  if(!filas.length){
    wrap.innerHTML=`<div class="empty" style="padding:24px">Sin coincidencias.</div>`;
    if(btnExp) btnExp.style.display="none";
    return;
  }
  $("#an-crit-count").textContent=`${filas.length} catálogos · ${nCrit} críticos`;
  wrap.innerHTML=`<div class="scroll"><table>
    <thead><tr>
      <th onclick="_anSortCrit('cat')">Catálogo${ic("cat")}</th>
      <th onclick="_anSortCrit('desc')">Descripción${ic("desc")}</th>
      <th class="r">UM</th>
      <th class="r" onclick="_anSortCrit('mens')">Cons. mensual D041${ic("mens")}</th>
      <th class="r" onclick="_anSortCrit('ex')">Existencia D041${ic("ex")}</th>
      <th class="r" onclick="_anSortCrit('rest')">Meses restantes${ic("rest")}</th>
      <th></th>
    </tr></thead><tbody>${filas.map(({cat,desc,um,mens,ex,restan,critico})=>`
      <tr ${critico?'style="background:var(--low-bg)"':""}>
        <td class="cat num">${cat}</td>
        <td style="font-size:12px;max-width:260px">${desc||""}</td>
        <td class="r" style="font-size:11px;color:var(--muted)">${um||""}</td>
        <td class="r num">${anFmt(mens,1)}</td>
        <td class="r num">${anFmt(ex)}</td>
        <td class="r num" style="${critico?"font-weight:700;color:var(--low)":""}">${restan==null?"—":anFmt(restan,2)}</td>
        <td>${critico?'<span class="an-badge r">⚠ Crítico</span>':""}</td>
      </tr>`).join("")}</tbody></table></div>`;
  if(btnExp) btnExp.style.display="inline-flex";
}
function _anSortCrit(col){
  if(col===anSortCritCol) anSortCritDir=-anSortCritDir; else { anSortCritCol=col; anSortCritDir=["cat","desc","area"].includes(col)?1:-1; }
  _pintarTablaCriticosD041();
}
function anExportarCriticosD041(){
  if(!anDatosCritD041.length) return;
  const per=anMCPM();
  const soloCrit=$("#an-crit-solo")?.checked;
  const filas=_filasCriticosD041();
  descargarXLSX([
    [`Materiales D041 — Consumo propio${soloCrit?" (solo críticos)":" (global)"}`,""],
    [`Periodo consumo: ${per}m | Crítico = existencia < 1 mes de consumo propio`,""],
    [],["Catálogo","Descripción","UM","Consumo mensual D041","Existencia D041","Meses restantes","Crítico"],
    ...filas.map(r=>[r.cat,r.desc,r.um,+(r.mens.toFixed(2)),r.ex,r.restan==null?"":+(r.restan.toFixed(2)),r.critico?"Sí":"No"])
  ],"Consumo D041",`Consumo_D041_${fechaTag()}`);
}


