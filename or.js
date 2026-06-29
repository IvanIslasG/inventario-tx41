/* ============================================================
   or.js  -  Módulo Orden de Reabasto · Sistema TX41 Puebla
   Depende de: DB, mat(), ngDe(), ngLista(), almName(),
               almList(), DIST(), showMod(), $(), AREAS,
               XLSX (SheetJS)
   ============================================================ */

/* ============ ORDEN DE REABASTO ============ */
let orAlm="", orSort={col:"cat",dir:1}, _orManual={};
function consumosDisp(){ return DB.consumos ? Object.keys(DB.consumos).sort() : []; }

// ── Pantalla bienvenida OR ───────────────────────────────────────────────────
let _orAreaSel = "";  // área seleccionada desde el menú inicial

function modOR(){
  _orAreaSel = "";
  _mostrarBienvenidaOR();
}

function _mostrarBienvenidaOR(){
  $("#moduleView").innerHTML = `
    <div id="or-bienvenida" style="
      display:flex;flex-direction:column;align-items:center;justify-content:center;
      min-height:70vh;padding:32px 20px;text-align:center">

      <!-- Animación de entrada -->
      <div id="or-anim-wrap" style="opacity:0;transform:translateY(20px);transition:all .6s ease">
        <div style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:2px;
                    color:var(--muted);margin-bottom:12px">Sistema de Inventario TX41</div>
        <div style="font-size:22px;font-weight:800;color:var(--primary);line-height:1.3;
                    margin-bottom:8px">Orden de Reabasto</div>
        <div style="font-size:15px;color:var(--text);font-weight:500;margin-bottom:6px">
          Telmex RNUM
        </div>
        <div style="font-size:13px;color:var(--muted)">Almacén Distribuidor Puebla · D041</div>
      </div>

      <!-- Menú de áreas -->
      <div id="or-menu-areas" style="
        opacity:0;transform:translateY(16px);transition:all .5s ease;
        margin-top:40px;width:100%;max-width:480px">

        <div style="font-size:13px;font-weight:600;color:var(--muted);
                    margin-bottom:16px;text-transform:uppercase;letter-spacing:.5px">
          ¿Qué OR quieres trabajar?
        </div>

        <div style="display:flex;flex-direction:column;gap:10px">
          ${AREAS.concat(["Sin clasificar"]).map(area => {
            const nMats = (() => {
              const cons = DB.consumos?.[orAlm] || {};
              const criticos = DB.criticos || [];
              const todos = new Set([...Object.keys(cons), ...criticos]);
              return [...todos].filter(cat => (mat(cat).area || "Sin clasificar") === area).length;
            })();
            return `<button onclick="_iniciarOR('${area.replace(/'/g,"\'")}') "
              style="display:flex;align-items:center;justify-content:space-between;
                     padding:14px 20px;background:white;border:1.5px solid var(--line);
                     border-radius:12px;cursor:pointer;text-align:left;width:100%;
                     font-family:inherit;transition:all .15s;font-size:14px;font-weight:600;
                     color:var(--text)"
              onmouseover="this.style.borderColor='var(--primary)';this.style.color='var(--primary)'"
              onmouseout="this.style.borderColor='var(--line)';this.style.color='var(--text)'">
              <span>${area}</span>
              <span style="font-size:11px;font-weight:400;color:var(--muted)">${nMats} materiales</span>
            </button>`;
          }).join("")}

          <button onclick="_iniciarOR('')"
            style="display:flex;align-items:center;justify-content:space-between;
                   padding:14px 20px;background:var(--primary);border:none;
                   border-radius:12px;cursor:pointer;text-align:left;width:100%;
                   font-family:inherit;font-size:14px;font-weight:700;color:white;margin-top:4px">
            <span>Todas las áreas</span>
            <span style="font-size:11px;font-weight:400;opacity:.8">Vista completa</span>
          </button>
        </div>
      </div>
    </div>
  `;

  // Animación secuencial
  requestAnimationFrame(() => {
    setTimeout(() => {
      const anim = document.getElementById("or-anim-wrap");
      if(anim){ anim.style.opacity="1"; anim.style.transform="translateY(0)"; }
    }, 80);
    setTimeout(() => {
      const menu = document.getElementById("or-menu-areas");
      if(menu){ menu.style.opacity="1"; menu.style.transform="translateY(0)"; }
    }, 500);
  });
}

function _iniciarOR(area){
  _orAreaSel = area;
  _modORPanel();
}

function _modORPanel(){
  const alms=consumosDisp();
  if(!alms.length){
    $("#moduleView").innerHTML=`<div class="panel"><div class="soon">
      <h2>Orden de Reabasto</h2>
      <p>Aún no hay archivos de consumo cargados.</p>
      <p>Ve a <b>Configuración → Archivos de consumo</b>, sube el export de consumo (ej. <b>TX8A</b>) y publica.</p>
    </div></div>`; return;
  }
  if(!orAlm || !alms.includes(orAlm)) orAlm=alms[0];
  const m=DB.meta, mCPM=m?.meses_cpm||6, mStk=m?.meses_stock||1.5;
  const opts=alms.map(a=>`<option value="${a}" ${a===orAlm?"selected":""}>${a} · ${almName(a)}</option>`).join("");
  const areasOR=AREAS.concat(["Sin clasificar"]);
  const areaOpts=areasOR.map(a=>`<option value="${a}">${a}</option>`).join("");

  $("#moduleView").innerHTML=`
    <div class="controls" style="flex-wrap:wrap">
      <label class="chk" style="font-weight:700;color:var(--ink)">Almacén
        <select id="orAlm">${opts}</select></label>
      <label class="chk">CPM (meses) <input type="number" id="orMCPM" value="${mCPM}" min="1" max="24" step="1" style="width:58px"></label>
      <label class="chk">Stock obj. <input type="number" id="orMS" value="${mStk}" min="0.5" max="12" step="0.5" style="width:58px"></label>
      <select id="orFArea"><option value="">Todas las áreas</option>${areaOpts}</select>
      <input type="search" id="orSearch" placeholder="Buscar catálogo / descripción…" style="min-width:180px">
    </div>
    <div class="controls" style="flex-wrap:wrap;margin-top:6px">
      <!-- Nombre genérico con autocomplete -->
      <div class="an-cat-wrap" style="min-width:230px;flex:1;max-width:320px">
        <input id="orNGInput" class="an-cat-input" placeholder="Nombre genérico… (escribe para filtrar)" autocomplete="off">
        <div id="orNGSugs" class="an-cat-sugs"></div>
        <input type="hidden" id="orNGVal">
      </div>
      <label class="chk"><input type="checkbox" id="orSolo"> Solo Cálc. surtir &gt; 0</label>
      <label class="chk"><input type="checkbox" id="orSoloX"> Solo X Surtir &gt; 0</label>
      <label class="chk"><input type="checkbox" id="orExced"> Solo excedentes</label>
      <label class="chk"><input type="checkbox" id="orSoloD041"> Solo con existencia D041</label>
      <button class="btn-prim" id="orExport">⬇ Exportar OR</button>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin:8px 0;font-size:12px">
      <span class="pill" id="orResumen"></span>
      <span style="color:var(--muted)">CPM = consumo ÷ meses · Cálc. surtir = (meses stock × CPM) − existencia aux · <b>X Surtir</b> editable</span>
    </div>
    <div id="orBandaNG"></div>
    <div class="panel"><div class="panel-head" style="gap:8px">
      <h2 id="orTitle"></h2><span class="pill" id="orCount"></span></div>
      <div style="overflow-x:auto"><table id="orTable" style="min-width:1200px"></table></div></div>`;
  $("#orAlm").onchange=e=>{ orAlm=e.target.value; _actualizarNGOpts(); pintarOR(); };
  ["orMCPM","orMS"].forEach(id=>{ const el=$("#"+id); el.oninput=pintarOR; el.onchange=pintarOR; });
  ["orSearch","orFArea"].forEach(id=>{ const el=$("#"+id); if(el){ el.oninput=()=>{ _actualizarNGOpts(); pintarOR(); }; }});
  ["orSolo","orSoloX","orExced","orSoloD041"].forEach(id=>{ const el=$("#"+id); if(el) el.onchange=pintarOR; });
  $("#orExport").onclick=abrirModalOR;

  // Autocomplete nombre genérico — filtra por área activa y muestra "Sin sustituto" incluido
  function _actualizarNGOpts(){
    const fArea=$("#orFArea")?.value||"";
    // calcular NGs activos desde consumos + críticos (igual que calcularOR)
    const cons=DB.consumos?.[orAlm]||{};
    const criticos=DB.criticos||[];
    const ngsActivos=new Set();
    const todosOR=new Set([...Object.keys(cons), ...criticos]);
    for(const cat of todosOR){
      const info=mat(cat);
      const area=info.area||"Sin clasificar";
      if(fArea && area!==fArea) continue;
      const ng=ngDe(cat);
      ngsActivos.add(ng||"SIN SUSTITUTO");
    }
    // guardar lista para el autocomplete
    $("#orNGInput")._ngs=[...ngsActivos].sort();
  }
  _actualizarNGOpts();

  const ngInp=$("#orNGInput"), ngSugs=$("#orNGSugs"), ngVal=$("#orNGVal");
  function _abrirNGSugs(q){
    const lista=(ngInp._ngs||[]);
    const qL=(q||"").toLowerCase();
    const hits=qL ? lista.filter(n=>n.toLowerCase().includes(qL)) : lista;
    ngSugs.innerHTML=[{v:"",label:"— Todos los nombres genéricos —"},...hits.map(n=>({v:n,label:n}))].map(
      ({v,label})=>`<div class="an-cat-sug" data-v="${v}">${label}</div>`
    ).join("");
    ngSugs.style.display="block";
  }
  ngInp.oninput=()=>{ ngVal.value=""; _abrirNGSugs(ngInp.value); pintarOR(); };
  // Punto 3: selectAll en el input de NG al enfocar
  ngInp.addEventListener("focus",()=>{ ngInp.select(); _abrirNGSugs(ngInp.value); });
  ngSugs.addEventListener("pointerdown",e=>{
    const t=e.target.closest("[data-v]"); if(!t) return;
    ngVal.value=t.dataset.v;
    ngInp.value=t.dataset.v||"";
    ngSugs.style.display="none";
    pintarOR();
  });
  document.addEventListener("click",e=>{ if(!ngSugs.contains(e.target)&&e.target!==ngInp) ngSugs.style.display="none"; });

  // Pre-seleccionar área elegida en el menú
  if(_orAreaSel){
    const sel = $("#orFArea");
    if(sel) sel.value = _orAreaSel;
    _actualizarNGOpts();
    pintarOR();
  } else {
    pintarOR();
  }
}

function calcularOR(){
  const D=DIST(), cons=DB.consumos?.[orAlm]||{};
  const mCPM=Math.max(1,+$("#orMCPM")?.value||6), mStk=Math.max(0.1,+$("#orMS")?.value||1.5);
  const out=[], vistos=new Set();

  const calcFila=(cat, consumo)=>{
    const info=mat(cat), ng=ngDe(cat)||"SIN SUSTITUTO";
    const cpm=consumo/mCPM;
    const exAux=(DB.existencias?.[cat]?.[orAlm])||0;
    const exD=(DB.existencias?.[cat]?.[D])||0;
    const calcSurtir=mStk*cpm-exAux;
    const man=_orManual[orAlm+"|"+cat]||{};
    const esSinSust=ng==="SIN SUSTITUTO";
    const sugerido=esSinSust ? Math.max(0,Math.min(Math.round(calcSurtir),exD)) : 0;
    const xsurtir=man.xs!=null ? man.xs : sugerido;
    const mdInv=cpm>0?exAux/cpm:null;
    const excedente=(3*cpm-exAux)<0;
    return {cat,desc:info.desc,um:info.um,area:info.area||"Sin clasificar",ng,consumo,cpm,exAux,exD,calcSurtir,xsurtir,mdInv,excedente,obs:man.obs||""};
  };

  // 1. Materiales con consumo en el almacén
  for(const [cat,consumo] of Object.entries(cons)){
    out.push(calcFila(cat, consumo));
    vistos.add(cat);
  }

  // 2. Críticos sin consumo — siempre deben aparecer
  for(const cat of (DB.criticos||[])){
    if(vistos.has(cat)) continue;
    out.push(calcFila(cat, 0));
  }

  return out;
}

function filasOR(){
  const q=($("#orSearch")?.value||"").trim().toLowerCase();
  const fArea=$("#orFArea")?.value||"";
  const ngHid=$("#orNGVal")?.value||"";
  const ngText=($("#orNGInput")?.value||"").trim().toLowerCase();
  const soloCal=$("#orSolo")?.checked;
  const soloX=$("#orSoloX")?.checked;
  const soloExc=$("#orExced")?.checked;

  let rows=calcularOR().filter(r=>{
    if(q && !(r.cat.toLowerCase().includes(q)||(r.desc||"").toLowerCase().includes(q))) return false;
    if(fArea && r.area!==fArea) return false;
    if(ngHid){ if(r.ng!==ngHid) return false; }
    else if(ngText){ if(!r.ng.toLowerCase().includes(ngText)) return false; }
    if(soloCal && r.calcSurtir<=0) return false;
    if(soloX && r.xsurtir<=0) return false;
    if(soloExc && !r.excedente) return false;
    if($("#orSoloD041")?.checked && (r.exD||0) <= 0) return false;
    return true;
  });

  // Punto 2: orden por defecto catálogo asc; respeta el sort manual si el usuario cambió columna
  const c=orSort.col, k=orSort.dir;
  const numCols=["consumo","cpm","exAux","exD","calcSurtir","xsurtir","mdInv"];
  rows.sort((a,b)=> numCols.includes(c)?((a[c]??-1)-(b[c]??-1))*k:String(a[c]||"").localeCompare(String(b[c]||""))*k);
  return rows;
}

const fmt1=n=>n==null?"S/C":(Math.round(n*10)/10).toLocaleString("es-MX");

function pintarOR(){
  const rows=filasOR();
  const nNec=rows.filter(r=>r.calcSurtir>0.5).length;
  const totalX=rows.reduce((s,r)=>s+r.xsurtir,0);
  $("#orTitle").innerHTML=`OR · ${orAlm} · ${almName(orAlm)}`;
  $("#orCount").textContent=`${rows.length} catálogos`;
  $("#orResumen").textContent=`${nNec} a surtir · Total X surtir: ${nfmt(totalX)}`;

  // Punto 4: sumatoria por nombre genérico — siempre actualizar el contenedor fijo
  const sumNG={};
  rows.forEach(r=>{
    if(r.ng==="SIN SUSTITUTO") return;
    if(!sumNG[r.ng]) sumNG[r.ng]={calc:0,xs:0,cats:[]};
    sumNG[r.ng].calc+=r.calcSurtir;
    sumNG[r.ng].xs+=r.xsurtir;
    sumNG[r.ng].cats.push(r.cat);
  });
  const bandaEl=$("#orBandaNG");
  const ngsActivos=Object.keys(sumNG);
  if(bandaEl){
    if(ngsActivos.length){
      bandaEl.style.display="";
      bandaEl.innerHTML=_htmlBandaNG(sumNG, false); // cerrada por defecto
    } else {
      bandaEl.style.display="none";
      bandaEl.innerHTML="";
    }
  }

  const thEx=`style="background:#e8f1fb;color:#0a4ea3;font-weight:800"`;
  const thExD=`style="background:#1a3a6a;color:#fff;font-weight:800"`;
  $("#orTable").innerHTML=`
  <colgroup>
    <col style="width:90px"><col style="width:220px"><col style="width:110px">
    <col style="width:140px"><col style="width:50px"><col style="width:80px">
    <col style="width:70px"><col style="width:60px"><col style="width:90px">
    <col style="width:90px"><col style="width:85px"><col style="width:82px">
    <col style="width:55px"><col>
  </colgroup>
  <thead><tr>
    <th data-c="cat">Catálogo</th><th data-c="desc">Descripción</th>
    <th data-c="area">Área</th><th data-c="ng">Nombre genérico</th><th>UM</th>
    <th class="r" data-c="consumo">Consumo</th><th class="r" data-c="cpm">CPM</th>
    <th class="r" data-c="mdInv">mdInv</th>
    <th class="r" data-c="exAux" ${thEx}>Exist. ${orAlm}</th>
    <th class="r" data-c="exD" ${thExD}>Exist. D041</th>
    <th class="r" data-c="calcSurtir">Cálc. surtir</th>
    <th class="r" data-c="xsurtir">X Surtir</th>
    <th>Exced.</th><th>Observaciones</th>
  </tr></thead><tbody>${
  rows.length? rows.map(r=>{
    const nec=r.calcSurtir>0.5;
    const esSinSust=r.ng==="SIN SUSTITUTO";
    const rowBg=nec?`style="background:#fdf1f1"`:(r.excedente?`style="background:var(--low-bg)"`:"");
    // resaltar ng cuando es sustituto (no SIN SUSTITUTO)
    const ngStyle=esSinSust?"color:var(--muted);font-size:12px":"color:#7a5c00;font-size:12px;font-weight:700";
    return `<tr ${rowBg}>
      <td class="cat num">${r.cat}</td>
      <td class="desc">${r.desc||"—"}</td>
      <td><span class="area-tag">${r.area}</span></td>
      <td style="${ngStyle}">${r.ng}</td>
      <td>${r.um||""}</td>
      <td class="r num">${nfmt(r.consumo)}</td>
      <td class="r num">${fmt1(r.cpm)}</td>
      <td class="r num">${fmt1(r.mdInv)}</td>
      <td class="r num" style="background:#eef5ff;font-weight:700">${nfmt(r.exAux)}</td>
      <td class="r num" style="background:#1a3a6a;color:#fff;font-weight:700">${nfmt(r.exD)}</td>
      <td class="r num" style="font-weight:700;color:${nec?"#c0392b":"var(--muted)"}">${fmt1(r.calcSurtir)}</td>
      <td class="r">
        <input type="number" class="or-xs" data-cat="${r.cat}" value="${r.xsurtir}" min="0"
          style="width:70px;text-align:right;padding:4px 6px;border:1px solid var(--line);border-radius:6px">
      </td>
      <td style="text-align:center">${r.excedente?"⚠️":""}</td>
      <td><input type="text" class="or-obs" data-cat="${r.cat}" value="${r.obs.replace(/"/g,"&quot;")}" placeholder="…"
          style="min-width:100px;padding:4px 6px;border:1px solid var(--line);border-radius:6px"></td>
    </tr>`;
  }).join("") : `<tr><td colspan="14" class="empty">Sin catálogos con estos filtros.</td></tr>`}</tbody>`;

  // Orden por cabecera
  $("#orTable").querySelectorAll("th[data-c]").forEach(th=>th.onclick=()=>{
    const c=th.dataset.c;
    orSort={col:c, dir:orSort.col===c?-orSort.dir:(["cat","desc","area","ng"].includes(c)?1:-1)};
    pintarOR();
  });

  // Guardar cambios en X Surtir y recalcular banda NG
  $("#orTable").querySelectorAll(".or-xs").forEach(inp=>{
    // Punto 3: selectAll al enfocar
    inp.addEventListener("focus",()=>inp.select());
    // Guardar y repintar banda NG al cambiar
    inp.addEventListener("change",()=>{
      (_orManual[orAlm+"|"+inp.dataset.cat]||={xs:0,obs:""}).xs=+inp.value||0;
      _repintarBandaNG();
    });
    // Punto 3: navegar con ↑ ↓ entre celdas X surtir
    inp.addEventListener("keydown",e=>{
      if(e.key==="ArrowDown"||e.key==="ArrowUp"){
        e.preventDefault();
        const all=[...$("#orTable").querySelectorAll(".or-xs")];
        const i=all.indexOf(inp);
        const next=all[e.key==="ArrowDown"?i+1:i-1];
        if(next){ next.focus(); next.select(); }
      }
    });
  });

  // Punto 3: selectAll en obs al enfocar
  $("#orTable").querySelectorAll(".or-obs").forEach(inp=>{
    inp.addEventListener("focus",()=>inp.select());
    inp.addEventListener("input",()=>{ (_orManual[orAlm+"|"+inp.dataset.cat]||={xs:0,obs:""}).obs=inp.value; });
  });
}

function _htmlBandaNG(sumNG, open=true){
  const entries=Object.entries(sumNG).sort((a,b)=>b[1].calc-a[1].calc);
  const n=entries.length;
  // Resumen compacto para mostrar en el summary (siempre visible)
  const resumen=entries.slice(0,4).map(([ng,v])=>
    `<span style="font-weight:700;color:#5a3e00">${ng}</span> <span style="color:#c0392b;margin-left:3px">Cálc: <b>${fmt1(v.calc)}</b></span>`
  ).join(' · ')+(n>4?` <span style="color:var(--muted)">+${n-4} más</span>`:"");
  // Chips completos dentro del body del details
  const chips=entries.map(([ng,v])=>`
    <div title="${v.cats.join(', ')}" style="background:#fff;border:1px solid #e8d080;border-radius:8px;padding:4px 10px;font-size:12px;cursor:default;white-space:nowrap">
      <span style="font-weight:700;color:#5a3e00">${ng}</span>
      <span style="color:#7a5c00;margin-left:5px">Cálc: <b>${fmt1(v.calc)}</b></span>
      <span style="color:#0a4ea3;margin-left:5px">X: <b>${nfmt(v.xs)}</b></span>
      <span style="color:var(--muted);font-size:10.5px;margin-left:4px">${v.cats.length} cat.</span>
    </div>`).join("");
  return `<details ${open?"open":""} style="background:#fffbe8;border:1px solid #f0d060;border-radius:10px;margin-bottom:8px">
    <summary style="list-style:none;padding:7px 12px;cursor:pointer;display:flex;align-items:center;gap:10px;user-select:none">
      <span style="font-size:12px;font-weight:800;color:#7a5c00;white-space:nowrap">∑ Sustitutos</span>
      <span style="font-size:11px;color:#b8960a;white-space:nowrap">${n} grupo${n!==1?"s":""}</span>
      <span style="flex:1"></span>
      <span class="ng-resumen" style="font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right">${resumen}</span>
      <span class="ng-arrow" style="font-size:13px;color:#b8960a;flex-shrink:0;margin-left:8px">▾</span>
    </summary>
    <div style="padding:6px 10px 10px;display:flex;flex-wrap:wrap;gap:5px">${chips}</div>
  </details>`;
}
function _repintarBandaNG(){
  const rows=filasOR();
  const sumNG={};
  rows.forEach(r=>{
    if(r.ng==="SIN SUSTITUTO") return;
    if(!sumNG[r.ng]) sumNG[r.ng]={calc:0,xs:0,cats:[]};
    sumNG[r.ng].calc+=r.calcSurtir;
    const man=_orManual[orAlm+"|"+r.cat];
    sumNG[r.ng].xs+=(man?.xs!=null?man.xs:r.xsurtir);
    sumNG[r.ng].cats.push(r.cat);
  });
  const totalX=rows.reduce((s,r)=>{
    const man=_orManual[orAlm+"|"+r.cat]; return s+(man?.xs!=null?man.xs:r.xsurtir);
  },0);
  $("#orResumen").textContent=`${rows.filter(r=>r.calcSurtir>0.5).length} a surtir · Total X surtir: ${nfmt(totalX)}`;
  const bandaEl=$("#orBandaNG"); if(!bandaEl) return;
  if(!Object.keys(sumNG).length){ bandaEl.style.display="none"; bandaEl.innerHTML=""; return; }
  bandaEl.style.display="";
  const wasOpen=bandaEl.querySelector("details")?.open===true;
  bandaEl.innerHTML=_htmlBandaNG(sumNG, wasOpen);
}

/* ---- Modal exportación OR (captura Nº Reabasto y Genera OR) ---- */
function abrirModalOR(){
  const rows=filasOR().filter(r=>r.xsurtir>0||r.calcSurtir>0.5);
  if(!rows.length){ alert("No hay catálogos para exportar (todos con X Surtir = 0)."); return; }
  const areasPresentes=[...new Set(rows.map(r=>r.area))].sort();
  const mCPM=+$("#orMCPM")?.value||6, mStk=+$("#orMS")?.value||1.5;
  const modal=document.createElement("div"); modal.className="modal on"; modal.id="orModal";
  modal.innerHTML=`<div class="modal-box" style="max-width:400px">
    <h3>⬇ Exportar Orden de Reabasto</h3>
    <div class="modal-body" style="gap:14px">
      <div>
        <label style="display:block;font-size:12px;font-weight:700;color:var(--muted);margin-bottom:5px">Número de Reabasto</label>
        <input type="number" id="orNumReabasto" min="1" placeholder="Ej. 5" style="width:100%">
      </div>
      <div>
        <label style="display:block;font-size:12px;font-weight:700;color:var(--muted);margin-bottom:5px">¿Quién genera la OR?</label>
        <input type="text" id="orGenerador" placeholder="Nombre del responsable" style="width:100%">
      </div>
      <div>
        <label style="display:block;font-size:12px;font-weight:700;color:var(--muted);margin-bottom:5px">Áreas a exportar</label>
        <div style="display:flex;flex-wrap:wrap;gap:6px">
          ${areasPresentes.map(a=>`<label class="chk" style="background:#f7f9fc;padding:5px 9px;border:1px solid var(--line);border-radius:8px">
            <input type="checkbox" class="or-area-chk" data-a="${a}" checked> ${a}</label>`).join("")}
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin:0">${nfmt(rows.length)} catálogos · CPM ${mCPM}m · Stock obj. ${mStk}m</p>
    </div>
    <div class="modal-foot" style="gap:8px">
      <button class="btn" id="orModalCancel">Cancelar</button>
      <button class="btn-prim" id="orModalOk">⬇ Exportar Excel</button>
    </div>
  </div>`;
  document.body.appendChild(modal);
  $("#orModalCancel").onclick=()=>modal.remove();

  const getParams=()=>{
    const num=$("#orNumReabasto").value.trim()||"—";
    const gen=$("#orGenerador").value.trim()||"";
    const selAreas=[...modal.querySelectorAll(".or-area-chk:checked")].map(c=>c.dataset.a);
    return {num, gen, selAreas};
  };

  // Excel con formato completo
  $("#orModalOk").onclick=()=>{
    const {num,gen,selAreas}=getParams(); // capturar ANTES de remove
    if(!selAreas.length){ alert("Selecciona al menos un área."); return; }
    const rowsFilt=rows.filter(r=>selAreas.includes(r.area));
    modal.remove(); // ahora sí remover
    try{
      _exportarORSimple(rowsFilt, num, gen, selAreas);
    }catch(e){
      console.error("Error al exportar OR:",e);
      alert("Error al generar el Excel: "+e.message);
    }
  };
}

/* ---- Export JSON para exportar_or.py ---- */
function _exportarORJson(rows, numReabasto, generador, areasSeleccionadas){
  const hoy=new Date().toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"});
  const mCPM=+$("#orMCPM")?.value||6, mStk=+$("#orMS")?.value||1.5;
  const areas={};
  areasSeleccionadas.forEach(a=>{ areas[a]=[]; });
  rows.forEach(r=>{ if(areas[r.area]) areas[r.area].push({cat:r.cat,desc:r.desc,um:r.um,xs:r.xsurtir,obs:r.obs}); });
  const payload={
    meta:{
      alm_sigla:orAlm, alm_nombre:almName(orAlm),
      num_reabasto:numReabasto, generador, fecha:hoy,
      meses_cpm:mCPM, meses_stock:mStk
    },
    areas
  };
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url;
  a.download=`OR_${orAlm}_Reb${numReabasto}_${hoy.replace(/\//g,"-")}_datos.json`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url),3000);
}

/* ---- Excel OR directo (sin servidor externo) ---- */
async function _exportarORSimple(rows, numReabasto, generador, areasSeleccionadas){
  const hoy=new Date().toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"});
  const alm=orAlm, almNom=almName(alm);
  console.log("[OR export] orAlm:", orAlm, "rows:", rows.length, "areas:", areasSeleccionadas);
  const mCPM=+document.getElementById("orMCPM")?.value||6;
  const mStk=+document.getElementById("orMS")?.value||1.5;
  const filename=`OR_${alm}_Reb${numReabasto}_${hoy.replace(/\//g,"-")}.xlsx`;
  const porArea={}; rows.forEach(r=>(porArea[r.area]||=[]).push(r));
  const wb=XLSX.utils.book_new();

  // ---- Hojas visuales (una por área, sin cambios) ----
  areasSeleccionadas.forEach(area=>{
    const filas=(porArea[area]||[]).filter(r=>r.xsurtir>0);
    if(!filas.length) return;
    const aoa=[
      [`Almacén: ${almNom} - ${alm}`,"","","","","","","","No. Guía",""],
      [`OR ${area}`],
      [`Reabasto: ${numReabasto}`,"","","","Pedido","","","","Siatel",""],
      [`Fecha: ${hoy}`],
      [],
      ["","","CATÁLOGO","DESCRIPCIÓN","Unidad","X SURTIR","","OBSERVs"],
      ...filas.map(f=>["","",f.cat,f.desc,f.um,f.xsurtir,"",f.obs||""]),
      [],
      ["","","","","HLog",`=SUM(F7:F${6+filas.length})`],
      [],[],[],
      ["","","Despacha:"],
      ["","","","Nombre"],
      [],[],[],
      ["","","Genera OR:",generador||""],
      [],
      ["","",`CPM ${mCPM} meses · Stock obj. ${mStk} meses`],
    ];
    const ws=XLSX.utils.aoa_to_sheet(aoa);
    ws["!merges"]=[
      {s:{r:0,c:0},e:{r:0,c:7}},{s:{r:0,c:8},e:{r:0,c:9}},
      {s:{r:1,c:0},e:{r:1,c:7}},
      {s:{r:2,c:0},e:{r:2,c:3}},{s:{r:2,c:4},e:{r:2,c:7}},{s:{r:2,c:8},e:{r:2,c:9}},
      {s:{r:3,c:0},e:{r:3,c:3}},
    ];
    ws["!cols"]=[{wch:3},{wch:3},{wch:13},{wch:40},{wch:8},{wch:11},{wch:3},{wch:32},{wch:14},{wch:14}];
    ws["!rows"]=[{hpt:22},{hpt:18},{hpt:20},{hpt:18},{hpt:6},{hpt:16}];
    XLSX.utils.book_append_sheet(wb, ws, area.substring(0,28));
  });

  // ---- Hoja SAP pegable ----
  const filasSAP=rows.filter(r=>r.xsurtir>0);
  if(filasSAP.length){
    // Detectar centro destino (todas las filas son del mismo centro)
    const siglaDestino=orAlm; // orAlm es el almacén auxiliar = destino de la OR
    const entryDir=DB.directorio?.almacenes?.[siglaDestino];
    // El bundle puede tener el centro como string directo o como objeto {centro, desc, ...}
    const centroDestino=(typeof entryDir==="string" ? entryDir : entryDir?.centro)||"";
    const esRN58=centroDestino==="RN58";

    let wsaSAP, colWidths;
    if(esRN58){
      // Rama MIGO 313 — mismo centro
      const hdr=["Txt. Breve mat.","Ctd. En UME","Lote","Lote,traspaso","Cl. Valoracion","Ce. Traspaso","Almacén, traspaso","Cl.valor., traspaso"];
      const datRows=filasSAP.map(f=>[f.cat, f.xsurtir, "", "", "", "", siglaDestino, "NUEVO"]);
      wsaSAP=XLSX.utils.aoa_to_sheet([hdr,...datRows]);
      colWidths=[{wch:18},{wch:13},{wch:8},{wch:14},{wch:15},{wch:13},{wch:18},{wch:18}];
    } else {
      // Rama ME21N — diferente centro
      const hdr=["Material","Txt.brv.","Ctd.Pedido","Lote","Almacén Procedencia","Centro","Almacén"];
      const datRows=filasSAP.map(f=>[f.cat, "", f.xsurtir, "", "D041", centroDestino, siglaDestino]);
      wsaSAP=XLSX.utils.aoa_to_sheet([hdr,...datRows]);
      wsaSAP["!cols"]=[{wch:13},{wch:8},{wch:13},{wch:8},{wch:20},{wch:10},{wch:10}];
      XLSX.utils.book_append_sheet(wb, wsaSAP, "SAP ME21N");

      // ---- Hoja MIGO 351 — salida de mercancías D041 ----
      const hdr351=["Material","Ctd. en UME","Lote","Cl. Valoracion"];
      const datRows351=filasSAP.map(f=>{
        const lotesD041=DB.lotes?.[f.cat];
        let lote="", clVal="";
        if(lotesD041 && lotesD041.length){
          lote=lotesD041[0].lote;
          clVal=lotesD041[0].lote;
        } else {
          const exD041=(DB.existencias?.[f.cat]?.["D041"])||0;
          if(exD041>0) clVal="NUEVO";
        }
        return [f.cat, f.xsurtir, lote, clVal];
      }).filter(r=>r[3]!=="");
      if(datRows351.length){
        const ws351=XLSX.utils.aoa_to_sheet([hdr351,...datRows351]);
        ws351["!cols"]=[{wch:13},{wch:13},{wch:18},{wch:16}];
        XLSX.utils.book_append_sheet(wb, ws351, "SAP MIGO 351");
      }
    }
  }

  if(!wb.SheetNames.length){ alert("Sin datos para exportar."); return; }
  XLSX.writeFile(wb, filename);
}

/* ---- Fallback: xlsx plano sin formato ---- */
function _exportarORPlano(rows, numReabasto, generador, areasSeleccionadas, hoy, filename){
  const alm=orAlm, almNom=almName(alm);
  const mCPM=+document.getElementById("orMCPM")?.value||6;
  const mStk=+document.getElementById("orMS")?.value||1.5;
  const wb=XLSX.utils.book_new();
  const porArea={}; rows.forEach(r=>(porArea[r.area]||=[]).push(r));
  areasSeleccionadas.forEach(area=>{
    const filas=(porArea[area]||[]).filter(r=>r.xsurtir>0);
    if(!filas.length) return;
    const aoa=[
      [`Almacén: ${almNom} - ${alm}`],
      [`OR ${area}`],
      [`Reabasto: ${numReabasto}`,"","","","Pedido","","","","Siatel"],
      [`Fecha: ${hoy}`],[], 
      ["","","CATÁLOGO","DESCRIPCIÓN","Unidad","X SURTIR","","OBSERVs"],
      ...filas.map(f=>["","",f.cat,f.desc,f.um,f.xsurtir,"",f.obs||""]),
      [],[],[],[],[`Genera OR: ${generador}`],
    ];
    const ws=XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"]=[{wch:3},{wch:3},{wch:13},{wch:40},{wch:8},{wch:11},{wch:3},{wch:32}];
    XLSX.utils.book_append_sheet(wb, ws, area.substring(0,28));
  });
  if(wb.SheetNames.length) XLSX.writeFile(wb, filename||"OR.xlsx");
}


