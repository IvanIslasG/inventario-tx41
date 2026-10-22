// ═══════════════════════════════════════════════════════════════════════════
// MÓDULO: CONTEO FÍSICO — captura de inventario físico con diferencias en vivo
// Depende de helpers globales de index.html: $, DB, mat, nfmt, DIST
// ═══════════════════════════════════════════════════════════════════════════
/* ============ CONTEO FÍSICO ============ */
let _conteoData = {};   // cat (o cat__lote) → físico capturado
let _conteoLista = [];  // lista filtrada actual
let _conteoMats  = [];  // materiales del módulo actual (D041)

// ── Pendientes de captura (se cuentan como Físico; se descontarán del stock al capturarse) ──
let _conteoPend = {};   // cat (o cat__lote) → cantidad pendiente de captura en SAP
let _cPendLog = [];      // bitácora de lo agregado en esta sesión (más reciente primero)
let _cPendUltimos = [];  // folios agregados en la última carga
let _cPendModoSel = 'folio';
let _conteoFolios = {}; // folio → { cat (o cat__lote) → cantidad }
function _cPendRecalc(){ _conteoPend={}; Object.values(_conteoFolios).forEach(f=>Object.keys(f).forEach(k=>{ _conteoPend[k]=(_conteoPend[k]||0)+f[k]; })); }
function _cPendCargar(){
  try{
    _conteoFolios = JSON.parse(localStorage.getItem("conteo_tx41_pend_folios")||"null");
    if(!_conteoFolios){ // migrar lista anterior sin folios
      const viejo = JSON.parse(localStorage.getItem("conteo_tx41_pendientes")||"{}");
      _conteoFolios = Object.keys(viejo).length ? {"Sin folio":viejo} : {};
    }
  }catch(e){ _conteoFolios = {}; }
  _cPendRecalc();
}
function _cPendGuardar(){ _cPendRecalc(); try{ localStorage.setItem("conteo_tx41_pend_folios", JSON.stringify(_conteoFolios)); }catch(e){} }
function _cPend(key){ return _conteoPend[key]||0; }
function _cEsc(t){ return String(t==null?"":t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
// Filas planas (una por material, o por lote) para conciliar
function _cFilas(mats){
  const out=[];
  mats.forEach(m=>{
    if(m.lotes && m.lotes.length>0){
      m.lotes.forEach(l=>{ const key=m.catalogo+"__"+l.lote; out.push({key,cat:m.catalogo,lote:l.lote,desc:m.descripcion,um:m.um,area:m.area,ubic:m.ubicacion,tot:_cTot(l.lib,l.tras,l.transitoPt),lib:l.lib||0,tras:l.tras||0}); });
    } else {
      out.push({key:m.catalogo,cat:m.catalogo,lote:"",desc:m.descripcion,um:m.um,area:m.area,ubic:m.ubicacion,tot:_cTot(m.existencia,m.traslado,m.transitoPt),lib:m.existencia||0,tras:m.traslado||0});
    }
  });
  return out;
}


// ── Tránsito PT: se suma al Total SAP porque va a llegar (casilla activable) ──
function _cIncTrans(){ try{ return localStorage.getItem("conteo_tx41_inc_transito")!=="0"; }catch(e){ return true; } }
function _cTot(lib,tras,trans){ return (lib||0)+(tras||0)+(_cIncTrans()?(trans||0):0); }
function _cToggleTrans(chk){ try{ localStorage.setItem("conteo_tx41_inc_transito", chk.checked?"1":"0"); }catch(e){} filtrarConteo(); }


// ── Clave localStorage por área ──────────────────────────────────────────────
function _conteoKey(area){ return "conteo_tx41_" + (area||"todo").replace(/\s+/g,"_").toLowerCase(); }

function _conteoCargarLS(area){
  try{ return JSON.parse(localStorage.getItem(_conteoKey(area)) || "{}"); }
  catch(e){ return {}; }
}

function _conteoGuardarLS(){
  try{ localStorage.setItem(_conteoKey(_conteoAreaActual), JSON.stringify(_conteoData)); }
  catch(e){}
}

function _conteoLimpiarLS(area){
  try{ localStorage.removeItem(_conteoKey(area)); }
  catch(e){}
}

// ── Menú de selección de área ─────────────────────────────────────────────────
let _conteoAreaActual = "";

function modConteo(){
  _cPendCargar();
  // Construir lista completa de materiales de D041
  _conteoMats = [];
  const mats = DB.materiales || {};
  const exs  = DB.existencias || {};
  const tras  = DB.traslados || {};
  const transito = DB.transitoPT || {};
  const lotes = DB.lotes || {};

  Object.keys(mats).forEach(cat => {
    const m   = mats[cat];
    const lib = (exs[cat] && exs[cat]["D041"]) || 0;
    const tra = (tras[cat] && tras[cat]["D041"]) || 0;
    const trPT = (transito[cat] && transito[cat]["D041"]) || 0;
    const ls  = (lotes[cat] || []).filter(l => (l.lib||0) + (l.tras||0) + (l.transitoPt||0) > 0); // se quitan lotes en 0 — solo estorban
    _conteoMats.push({
      catalogo:    cat,
      descripcion: m.desc || m.descripcion || "",
      um:          m.um || "",
      area:        m.area || "",
      ubicacion:   m.ubic || m.ubicacion || "",
      existencia:  lib,
      traslado:    tra,
      transitoPt:  trPT, // se suma al Total SAP (casilla 'Contar tránsito PT'): va a llegar a D041
      lotes:       ls,
    });
  });
  _conteoMats.sort((a,b) => a.catalogo.localeCompare(b.catalogo));

  // Obtener áreas únicas
  const areas = ["Todo el inventario",
    ...[...new Set(_conteoMats.map(m=>m.area).filter(Boolean))].sort()
  ];

  // Detectar conteos en progreso en localStorage
  const enProgreso = areas.filter(a => {
    const key = _conteoKey(a);
    try{ const d = JSON.parse(localStorage.getItem(key)||"{}"); return Object.keys(d).length > 0; }
    catch(e){ return false; }
  });

  $("#moduleView").innerHTML=`
    <div style="max-width:520px;margin:0 auto;padding-top:8px">
      <div style="margin-bottom:24px">
        <h2 style="margin:0 0 4px;font-size:20px">Inventario Físico</h2>
        <p style="margin:0;color:var(--muted);font-size:13px">D041 · Almacén Distribuidor Puebla</p>
      </div>

      <p style="font-size:14px;font-weight:600;margin:0 0 12px">¿De qué área vas a hacer inventario?</p>

      <div style="display:flex;flex-direction:column;gap:8px">
        ${areas.map(a => {
          const enProg = enProgreso.includes(a);
          const nMats  = a === "Todo el inventario"
            ? _conteoMats.length
            : _conteoMats.filter(m=>m.area===a).length;
          let capturados = 0;
          if(enProg){
            try{
              const d = JSON.parse(localStorage.getItem(_conteoKey(a))||"{}");
              capturados = Object.keys(d).length;
            }catch(e){}
          }
          return `
          <button onclick="iniciarConteo('${a.replace(/'/g,"\'")}') "
            style="display:flex;align-items:center;gap:14px;padding:14px 16px;
                   background:${enProg?"#f0fdf4":"white"};
                   border:2px solid ${enProg?"#86efac":"var(--borde)"};
                   border-radius:12px;cursor:pointer;text-align:left;width:100%;
                   transition:border-color .15s">
            <div style="flex:1">
              <div style="font-size:14px;font-weight:600;color:var(--texto)">${a}</div>
              <div style="font-size:11px;color:var(--muted);margin-top:2px">${nMats} materiales</div>
            </div>
            ${enProg ? `<div style="text-align:right">
              <div style="font-size:11px;font-weight:600;color:#16a34a">EN PROGRESO</div>
              <div style="font-size:11px;color:var(--muted)">${capturados} capturados</div>
            </div>` : ""}
            <div style="font-size:18px;color:var(--muted)">›</div>
          </button>`;
        }).join("")}
      </div>

      ${enProgreso.length > 0 ? `
      <div style="margin-top:20px;padding:12px 14px;background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;font-size:12px;color:#9a3412">
        💾 Los conteos en progreso están guardados en este dispositivo y continuarán donde los dejaste.
      </div>` : ""}
    </div>
  `;
}

function iniciarConteo(area){
  _conteoAreaActual = area;
  _cPendCargar();
  // Cargar datos guardados de localStorage
  _conteoData = _conteoCargarLS(area);
  const hayDatos = Object.keys(_conteoData).length > 0;

  // Filtrar materiales según área seleccionada
  const esTodo = area === "Todo el inventario";
  const matsArea = esTodo ? _conteoMats : _conteoMats.filter(m=>m.area===area);

  $("#moduleView").innerHTML=`
    <div style="margin-bottom:12px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <button class="btn btn-outline btn-sm" onclick="modConteo()" style="padding:4px 10px">‹ Áreas</button>
      <div style="flex:1">
        <h2 style="margin:0;font-size:17px">Inventario · ${area}</h2>
        <div id="conteo-subtitulo" style="font-size:12px;color:var(--muted);margin-top:2px">${matsArea.length} materiales</div>
      </div>
      <button class="btn btn-outline btn-sm" onclick="exportarConteo()">⬇ Exportar</button>
      <button class="btn btn-outline btn-sm" onclick="importarConteo()">⬆ Importar</button>
      <button class="btn btn-outline btn-sm" id="cPendBtn" onclick="_cPendAbrir()" style="border-color:#f59e0b;color:#b45309">⏳ Pendientes de captura${Object.keys(_conteoPend).length?` (${Object.keys(_conteoPend).length})`:""}</button>
      <button class="btn btn-outline btn-sm" style="color:var(--rojo,#dc2626)" onclick="limpiarConteo()">🗑 Limpiar</button>
      <input type="file" id="importarConteoInput" accept=".xlsx" style="display:none" onchange="procesarImportConteo(this)">
    </div>

    <div class="cap-stats">
      <div class="cap-stat"><div class="cap-stat-n" id="cc-cap">0</div><div class="cap-stat-l">Capturados</div></div>
      <div class="cap-stat"><div class="cap-stat-n dif-pending" id="cc-pen">${matsArea.length}</div><div class="cap-stat-l">Pendientes</div></div>
      <div class="cap-stat"><div class="cap-stat-n dif-ok" id="cc-ok">0</div><div class="cap-stat-l">Sin diferencia</div></div>
      <div class="cap-stat"><div class="cap-stat-n dif-neg" id="cc-falt">0</div><div class="cap-stat-l">Faltantes</div></div>
      <div class="cap-stat"><div class="cap-stat-n dif-pos" id="cc-sob">0</div><div class="cap-stat-l">Sobrantes</div></div>
    </div>

    ${hayDatos ? `
    <div style="padding:10px 14px;background:#f0fdf4;border:1px solid #86efac;border-radius:10px;font-size:12px;color:#15803d;margin-bottom:12px">
      💾 Conteo en progreso recuperado — ${Object.keys(_conteoData).length} valores guardados en este dispositivo.
    </div>` : ""}

    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px">
      <input id="conteo-q" type="search" placeholder="Buscar catálogo o descripción..."
        style="flex:1;min-width:180px" oninput="filtrarConteo()">
      <select id="conteo-dif" onchange="filtrarConteo()" style="min-width:140px">
        <option value="">Todos</option>
        <option value="pendiente">⏳ Pendientes</option>
        <option value="capturado">✏️ Solo capturados</option>
        <option value="ok">✓ Sin diferencia</option>
        <option value="faltante">🔴 Faltantes</option>
        <option value="sobrante">🟡 Sobrantes</option>
        <option value="conpend">⏳ Con pendiente de captura</option>
      </select>
      <select id="conteo-ubic" onchange="filtrarConteo()" style="min-width:140px">
        <option value="">Todas las ubicaciones</option>
      </select>
      <label class="chk"><input type="checkbox" id="conteo-solo-ex" onchange="filtrarConteo()"> Solo con existencia</label>
      <label class="chk" title="El tránsito PT va a llegar a D041, por eso se suma al Total SAP"><input type="checkbox" id="conteo-inc-tr" ${_cIncTrans()?"checked":""} onchange="_cToggleTrans(this)"> Contar tránsito PT en el total</label>
    </div>

    <div id="conteo-lista"></div>
  `;
  // Botones flotantes para móvil
  _conteoFabInit();

  // Poblar select de ubicaciones
  const ubics = [...new Set(matsArea.map(m=>m.ubicacion).filter(Boolean))].sort();
  const selUbic = $("#conteo-ubic");
  if(selUbic){
    selUbic.innerHTML = `<option value="">Todas las ubicaciones</option>` +
      ubics.map(u=>`<option value="${u}">${u}</option>`).join("");
  }

  // Filtrar materiales del área y renderizar
  _conteoLista = matsArea;
  filtrarConteo();
}

function filtrarConteo(){
  const q   = ($("#conteo-q")?.value||"").toLowerCase();
  const dif    = $("#conteo-dif")?.value||"";
  const soloEx = $("#conteo-solo-ex")?.checked;
  const ubic   = $("#conteo-ubic")?.value||"";
  const esTodo = _conteoAreaActual === "Todo el inventario";
  const base = esTodo ? _conteoMats : _conteoMats.filter(m=>m.area===_conteoAreaActual);

  _conteoLista = base.filter(m=>{
    const mQ = !q || m.catalogo.includes(q) || m.descripcion.toLowerCase().includes(q);
    if(soloEx && (m.existencia||0) <= 0) return false;
    if(ubic && m.ubicacion !== ubic) return false;
    let mD = true;
    if(dif){
      const fisico    = _conteoData[m.catalogo];
      const cap       = fisico !== undefined;
      const total     = _cTot(m.existencia, m.traslado, m.transitoPt);
      const pendM = (m.lotes&&m.lotes.length) ? m.lotes.reduce((a,l)=>a+_cPend(m.catalogo+"__"+l.lote),0)+_cPend(m.catalogo) : _cPend(m.catalogo);
      const diferencia = cap ? fisico + _cPend(m.catalogo) - total : null;
      if(dif==="conpend")       mD = pendM>0;
      else if(dif==="pendiente")     mD = !cap;
      else if(dif==="capturado") mD = cap;
      else if(dif==="ok")        mD = cap && diferencia===0;
      else if(dif==="faltante")  mD = cap && diferencia<0;
      else if(dif==="sobrante")  mD = cap && diferencia>0;
    }
    return mQ && mD;
  });

  const sub = $("#conteo-subtitulo");
  if(sub) sub.textContent = `${_conteoLista.length} visibles · ${Object.keys(_conteoData).length} capturados`;

  renderConteo();
  actualizarTotalesConteo();
}

function renderConteo(){
  const lista = $("#conteo-lista");
  if(!lista) return;
  lista.innerHTML = "";

  _conteoLista.forEach(m=>{
    const tieneLotes = m.lotes && m.lotes.length > 0;

    if(tieneLotes){
      const div = document.createElement("div");
      div.id = `cc-${m.catalogo}`;
      div.style.cssText = "background:white;border-radius:12px;border:2px solid var(--borde);overflow:hidden;margin-bottom:8px";
      div.innerHTML = `
        <div style="padding:10px 14px;background:#fff7ed;border-bottom:1px solid #fed7aa;display:flex;align-items:center;gap:10px">
          <span class="cap-cat" style="color:#9a3412">${m.catalogo}</span>
          <span style="font-size:10px;background:#fed7aa;color:#9a3412;padding:2px 6px;border-radius:8px">📦 ${m.lotes.length} lotes</span>
          <div style="flex:1;font-size:12px;color:var(--texto)">${m.descripcion}</div>
          <span style="font-size:10px;background:#e0e7ff;color:#3730a3;padding:2px 6px;border-radius:8px" title="Va a llegar a D041 — se suma al Total si la casilla 'Contar tránsito PT' está activa">🚚 Tránsito PT: ${(m.transitoPt||0).toLocaleString()}</span>
          <div style="font-size:10px;color:var(--muted)">${m.ubicacion||"Sin ubicación"} · ${m.um||"—"}</div>
        </div>`;

      const tbody = document.createElement("div");
      tbody.style.padding = "8px";

      m.lotes.forEach((lote, idx)=>{
        const key  = `${m.catalogo}__${lote.lote}`;
        const fis  = _conteoData[key];
        const cap  = fis !== undefined;
        const tot  = _cTot(lote.lib, lote.tras, lote.transitoPt);
        const penL = _cPend(key);
        const dif  = cap ? fis + penL - tot : null;
        let difCls = "dif-pending", difTxt = "—", bg = "transparent";
        if(cap){
          if(dif===0)    {difCls="dif-ok";  difTxt="✓ 0"; bg="#f0fdf4";}
          else if(dif<0) {difCls="dif-neg"; difTxt=dif;   bg="#fff5f5";}
          else           {difCls="dif-pos"; difTxt="+"+dif; bg="#fffdf0";}
        }
        const row = document.createElement("div");
        row.style.cssText = `display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;margin-bottom:4px;background:${bg}`;
        row.innerHTML = `
          <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#9a3412;font-weight:600;min-width:80px">${lote.lote}</div>
          <div class="cap-siatel"><div class="cap-siatel-n">${(lote.lib||0).toLocaleString()}</div><div class="cap-siatel-l">Libre</div></div>
          <div class="cap-siatel"><div class="cap-siatel-n" style="color:var(--naranja,#ea580c)">${(lote.tras||0).toLocaleString()}</div><div class="cap-siatel-l">Traslado</div></div>
          <div class="cap-siatel" title="Va a llegar a D041 — se suma al Total si la casilla está activa"><div class="cap-siatel-n" style="color:#3730a3">${(lote.transitoPt||0).toLocaleString()}</div><div class="cap-siatel-l">Tránsito PT</div></div>
          ${penL>0?`<div class="cap-siatel" title="Pendiente de captura: cuenta como físico"><div class="cap-siatel-n" style="color:#b45309">+${penL.toLocaleString()}</div><div class="cap-siatel-l">Pend.</div></div>`:""}
          <div style="flex:1"></div>
          <div class="cap-input-wrap">
            <input class="cap-input" type="number" inputmode="numeric" min="0" placeholder="—"
              value="${cap?fis:""}"
              data-cat="${m.catalogo}" data-lote="${lote.lote}" data-idx="${idx}"
              data-libre="${lote.lib||0}" data-traslado="${lote.tras||0}" data-transito="${lote.transitoPt||0}"
              onkeyup="if(event.key==='Enter'||event.keyCode===13)_conteoEnterLoteC(this)" onblur="guardarFisicoLoteC(this,false)" onfocus="_conteoFocoInput(this)" enterkeyhint="next" inputmode="numeric"
              onkeydown="navegarInputLoteC(event,'${m.catalogo}',${idx})">
            <div class="cap-input-l">Físico</div>
          </div>
          <div class="cap-dif"><div class="cap-dif-n ${difCls}" id="ccdif-${m.catalogo}-${idx}">${difTxt}</div><div class="cap-dif-l">Dif.</div></div>`;
        tbody.appendChild(row);
      });
      div.appendChild(tbody);
      lista.appendChild(div);

    } else {
      const fis  = _conteoData[m.catalogo];
      const cap  = fis !== undefined;
      const lib  = m.existencia||0;
      const tra  = m.traslado||0;
      const tot  = _cTot(lib, tra, m.transitoPt);
      const pen  = _cPend(m.catalogo);
      const dif  = cap ? fis + pen - tot : null;
      let cardCls = "cap-card", difCls = "dif-pending", difTxt = "—";
      if(cap){
        if(dif===0)    {cardCls+=" capturado"; difCls="dif-ok";  difTxt="✓ 0";}
        else if(dif<0) {cardCls+=" faltante";  difCls="dif-neg"; difTxt=dif;}
        else           {cardCls+=" sobrante";  difCls="dif-pos"; difTxt="+"+dif;}
      }
      const div = document.createElement("div");
      div.className = cardCls;
      div.id = `cc-${m.catalogo}`;
      div.innerHTML = `
        <div class="cap-cat">${m.catalogo}</div>
        <div style="flex:1">
          <div class="cap-desc">${m.descripcion}</div>
          <div class="cap-ubi">${m.ubicacion||"Sin ubicación"} · ${m.um||"—"}</div>
        </div>
        <div class="cap-siatel"><div class="cap-siatel-n">${lib.toLocaleString()}</div><div class="cap-siatel-l">Libre</div></div>
        <div class="cap-siatel"><div class="cap-siatel-n" style="color:var(--naranja,#ea580c)">${tra.toLocaleString()}</div><div class="cap-siatel-l">Traslado</div></div>
        <div class="cap-siatel" title="Va a llegar a D041 — se suma al Total si la casilla 'Contar tránsito PT' está activa"><div class="cap-siatel-n" style="color:#3730a3">${(m.transitoPt||0).toLocaleString()}</div><div class="cap-siatel-l">Tránsito PT</div></div>
        ${pen>0?`<div class="cap-siatel" title="Pendiente de captura: cuenta como físico"><div class="cap-siatel-n" style="color:#b45309">+${pen.toLocaleString()}</div><div class="cap-siatel-l">Pend.</div></div>`:""}
        <div class="cap-input-wrap">
          <input class="cap-input" type="number" inputmode="numeric" min="0" placeholder="—"
            value="${cap?fis:""}"
            data-cat="${m.catalogo}" data-libre="${lib}" data-traslado="${tra}" data-transito="${m.transitoPt||0}"
            onkeyup="if(event.key==='Enter'||event.keyCode===13)_conteoEnterC(this)" onblur="guardarFisicoC(this,false)" onfocus="_conteoFocoInput(this)" enterkeyhint="next" inputmode="numeric"
            onkeydown="navegarInputC(event,'${m.catalogo}')">
          <div class="cap-input-l">Físico</div>
        </div>
        <div class="cap-dif"><div class="cap-dif-n ${difCls}" id="ccdif-${m.catalogo}">${difTxt}</div><div class="cap-dif-l">Diferencia</div></div>`;
      lista.appendChild(div);
    }
  });
}

function guardarFisicoC(input, avanzar=true){
  const cat = input.dataset.cat;
  const lib = parseInt(input.dataset.libre)||0;
  const tra = parseInt(input.dataset.traslado)||0;
  const tot = _cTot(lib, tra, parseInt(input.dataset.transito)||0);
  const val = input.value.trim();
  if(val===""||isNaN(val)) delete _conteoData[cat];
  else _conteoData[cat] = parseInt(val);
  const fis = _conteoData[cat];
  const cap = fis !== undefined;
  const dif = cap ? fis + _cPend(cat) - tot : null;
  const card = document.getElementById(`cc-${cat}`);
  if(card){
    card.className = "cap-card" + (cap?(dif===0?" capturado":dif<0?" faltante":" sobrante"):"");
    const difEl = card.querySelector(".cap-dif-n");
    if(difEl){
      difEl.className  = "cap-dif-n " + (cap?(dif===0?"dif-ok":dif<0?"dif-neg":"dif-pos"):"dif-pending");
      difEl.textContent = cap?(dif===0?"✓ 0":dif<0?dif:"+"+dif):"—";
    }
  }
  _conteoGuardarLS();
  actualizarTotalesConteo();
  if(!avanzar) return; // solo guardar, no mover el foco
  const idx = _conteoLista.findIndex(m=>m.catalogo===cat);
  const sig = _conteoLista[idx+1];
  if(sig){
    const next=document.querySelector(`input[data-cat="${sig.catalogo}"]`);
    if(next){ _conteoFocoInput(next); }
  }
}

function _conteoEnterC(input){
  guardarFisicoC(input, true);
}

function guardarFisicoLoteC(input, avanzar=true){
  const cat = input.dataset.cat;
  const lote= input.dataset.lote;
  const idx = parseInt(input.dataset.idx);
  const lib = parseFloat(input.dataset.libre)||0;
  const tra = parseFloat(input.dataset.traslado)||0;
  const tot = _cTot(lib, tra, parseFloat(input.dataset.transito)||0);
  const key = `${cat}__${lote}`;
  const val = input.value.trim();
  if(val===""||isNaN(val)) delete _conteoData[key];
  else _conteoData[key] = parseInt(val);
  const fis = _conteoData[key];
  const cap = fis !== undefined;
  const dif = cap ? fis + _cPend(key) - tot : null;
  const difEl = document.getElementById(`ccdif-${cat}-${idx}`);
  if(difEl){
    difEl.className  = "cap-dif-n " + (cap?(dif===0?"dif-ok":dif<0?"dif-neg":"dif-pos"):"dif-pending");
    difEl.textContent = cap?(dif===0?"✓ 0":dif<0?dif:"+"+dif):"—";
  }
  _conteoGuardarLS();
  actualizarTotalesConteo();
}

function _conteoEnterLoteC(input){
  guardarFisicoLoteC(input, true);
  const cat = input.dataset.cat;
  const idx = parseInt(input.dataset.idx);
  const m = _conteoMats.find(m=>m.catalogo===cat);
  if(m && idx < m.lotes.length-1){
    const next = document.querySelector(`input[data-cat="${cat}"][data-idx="${idx+1}"]`);
    if(next){ _conteoFocoInput(next); }
  } else {
    const mi = _conteoLista.findIndex(m=>m.catalogo===cat);
    const sig = _conteoLista[mi+1];
    if(sig){ const next=document.querySelector(`input[data-cat="${sig.catalogo}"]`); if(next){ _conteoFocoInput(next); } }
  }
}

function _conteoFocoInput(el){
  el.focus();
  try{ el.setSelectionRange(0, el.value.length); } catch(e){}
  // Solo hacer scroll si el elemento está fuera del viewport visible
  // Sin smooth para evitar loop en móvil con teclado abierto
  const rect = el.getBoundingClientRect();
  const vh = window.innerHeight;
  const margen = 120; // espacio para el teclado
  if(rect.top < 60 || rect.bottom > vh - margen){
    el.scrollIntoView({block:"center"});
  }
}

function navegarInputC(event, cat){
  // Manejado por onkeyup directamente
}

function navegarInputLoteC(event, cat, idx){
  if(event.key==="Enter"||event.key==="Tab"){
    event.preventDefault();
    const input = event.target;
    if(input) guardarFisicoLoteC(input);
    const m = _conteoMats.find(m=>m.catalogo===cat);
    if(m && idx < m.lotes.length-1){
      const next = document.querySelector(`input[data-cat="${cat}"][data-idx="${idx+1}"]`);
      if(next){ _conteoFocoInput(next); }
    } else {
      const mi = _conteoLista.findIndex(m=>m.catalogo===cat);
      const sig = _conteoLista[mi+1];
      if(sig){ const next=document.querySelector(`input[data-cat="${sig.catalogo}"]`); if(next){ _conteoFocoInput(next); } }
    }
  }
}

function actualizarTotalesConteo(){
  // Usar _conteoLista para respetar filtros activos (área + ubicación + búsqueda)
  const base = (_conteoLista && _conteoLista.length > 0) ? _conteoLista : _conteoMats;
  let cap=0,pen=0,ok=0,falt=0,sob=0;
  base.forEach(m=>{
    if(m.lotes && m.lotes.length>0){
      m.lotes.forEach(lote=>{
        const key=`${m.catalogo}__${lote.lote}`;
        const fis=_conteoData[key];
        const tot=_cTot(lote.lib,lote.tras,lote.transitoPt);
        if(fis===undefined){pen++;return;}
        cap++;
        const dif=fis+_cPend(key)-tot;
        if(dif===0)ok++;else if(dif<0)falt++;else sob++;
      });
    } else {
      const fis=_conteoData[m.catalogo];
      const tot=_cTot(m.existencia,m.traslado,m.transitoPt);
      if(fis===undefined){pen++;return;}
      cap++;
      const dif=fis+_cPend(m.catalogo)-tot;
      if(dif===0)ok++;else if(dif<0)falt++;else sob++;
    }
  });
  ["cc-cap","cc-pen","cc-ok","cc-falt","cc-sob"].forEach(id=>{
    const el=document.getElementById(id);
    if(!el) return;
    const n=id==="cc-cap"?cap:id==="cc-pen"?pen:id==="cc-ok"?ok:id==="cc-falt"?falt:sob;
    el.textContent=n;
  });
}

// ── Botones flotantes móvil ──────────────────────────────────────────────────
function _conteoFabInit(){
  // Quitar FAB anterior si existe
  const old = document.getElementById('conteo-fab-group');
  if(old) old.remove();

  const g = document.createElement('div');
  g.id = 'conteo-fab-group';
  g.className = 'fab-group';
  g.innerHTML = `
    <button class="fab secondary" title="Ir arriba" onclick="_conteoFabArriba()" style="position:relative">
      <span style="font-size:14px">↑</span>
      <span class="fab-label">Ir arriba</span>
    </button>
    <button class="fab secondary" title="Último capturado" onclick="_conteoFabUltimo()" style="position:relative">
      <span style="font-size:14px">✎</span>
      <span class="fab-label">Último capturado</span>
    </button>
    <button class="fab secondary" title="Ir abajo" onclick="_conteoFabAbajo()" style="position:relative">
      <span style="font-size:14px">↓</span>
      <span class="fab-label">Ir abajo</span>
    </button>
    <button class="fab" title="Siguiente pendiente" onclick="_conteoFabPendiente()" style="position:relative">
      <span style="font-size:14px">⏭</span>
      <span class="fab-label">Siguiente pendiente</span>
    </button>
  `;
  document.body.appendChild(g);

  // Mostrar/ocultar según scroll
  let scrollTimer;
  window.addEventListener('scroll', ()=>{
    g.classList.add('visible');
    clearTimeout(scrollTimer);
    scrollTimer = setTimeout(()=>{ g.classList.remove('visible'); }, 6000);
  }, {passive:true});

  // Mostrar siempre en móvil si se toca la pantalla
  document.addEventListener('touchstart', ()=>{
    g.classList.add('visible');
    clearTimeout(scrollTimer);
    scrollTimer = setTimeout(()=>{ g.classList.remove('visible'); }, 6000);
  }, {passive:true, once:false});
}

function _conteoFabArriba(){
  window.scrollTo({top:0, behavior:'smooth'});
}

function _conteoFabAbajo(){
  window.scrollTo({top:document.body.scrollHeight, behavior:'smooth'});
}

function _conteoFabUltimo(){
  // Ir al último input con valor capturado
  const inputs = [...document.querySelectorAll('#conteo-lista input.cap-input')];
  const capturados = inputs.filter(i => i.value !== '');
  if(!capturados.length){ 
    alert('No hay materiales capturados aún.'); 
    return; 
  }
  const ultimo = capturados[capturados.length - 1];
  ultimo.scrollIntoView({block:'center'});
  ultimo.focus();
  try{ ultimo.setSelectionRange(0, ultimo.value.length); }catch(e){}
}

function _conteoFabPendiente(){
  // Ir al siguiente input sin valor
  const inputs = [...document.querySelectorAll('#conteo-lista input.cap-input')];
  // Buscar el primero visible sin capturar
  const pendiente = inputs.find(i => i.value === '');
  if(!pendiente){
    alert('¡Todos los materiales visibles están capturados!');
    return;
  }
  pendiente.scrollIntoView({block:'center'});
  pendiente.focus();
  try{ pendiente.setSelectionRange(0, pendiente.value.length); }catch(e){}
}

function importarConteo(){
  document.getElementById('importarConteoInput').click();
}

function procesarImportConteo(input){
  const file = input.files[0];
  if(!file){ return; }
  input.value = ""; // reset para poder importar el mismo archivo otra vez

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:'array'});
      // Buscar hoja "Inventario" (la que genera exportarConteo)
      const wsName = wb.SheetNames.find(s => s === "Inventario") || wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1});

      // Encontrar fila de encabezados (tiene "Catálogo")
      let hdrIdx = -1;
      for(let i=0; i<rows.length; i++){
        if(rows[i] && rows[i].some(c => String(c||'').toLowerCase().includes('catálogo') || String(c||'').toLowerCase().includes('catalogo'))){
          hdrIdx = i; break;
        }
      }
      if(hdrIdx < 0){ alert("No encontré la columna de Catálogo en el archivo."); return; }

      const hdrs = rows[hdrIdx].map(h => String(h||'').toLowerCase());
      const iCat   = hdrs.findIndex(h => h.includes('catálogo') || h.includes('catalogo'));
      const iFis   = hdrs.findIndex(h => h.includes('físico') || h.includes('fisico'));
      const iLote  = hdrs.findIndex(h => h === 'lote');

      if(iCat < 0 || iFis < 0){ alert("El archivo no tiene las columnas de Catálogo y Físico."); return; }

      // Leer datos
      const importados = {};
      for(let i = hdrIdx+1; i < rows.length; i++){
        const r = rows[i];
        if(!r || !r[iCat]) continue;
        const cat  = String(r[iCat]).trim();
        const lote = iLote >= 0 ? String(r[iLote]||'').trim() : '';
        const fis  = r[iFis];
        if(fis === '' || fis === null || fis === undefined || isNaN(fis)) continue;
        const key = lote ? cat+'__'+lote : cat;
        importados[key] = parseInt(fis);
      }

      const nImport = Object.keys(importados).length;
      if(nImport === 0){ alert("El archivo no tiene datos físicos capturados."); return; }

      // Preguntar modo
      const hayDatos = Object.keys(_conteoData).length > 0;
      let modo = 'reemplazar';
      if(hayDatos){
        const resp = confirm(
          `Se encontraron ${nImport} registros en el archivo.

` +
          `Ya tienes ${Object.keys(_conteoData).length} valores capturados en este dispositivo.

` +
          `¿Cómo quieres importar?
` +
          `• Aceptar → Fusionar (solo llena los vacíos, respeta lo que ya capturaste)
` +
          `• Cancelar → Reemplazar todo`
        );
        modo = resp ? 'fusionar' : 'reemplazar';
      }

      if(modo === 'reemplazar'){
        _conteoData = importados;
      } else {
        // Fusionar: solo llenar los que no tienen valor
        Object.entries(importados).forEach(([k,v])=>{
          if(_conteoData[k] === undefined) _conteoData[k] = v;
        });
      }

      _conteoGuardarLS();
      filtrarConteo();
      actualizarTotalesConteo();

      const nFinal = Object.keys(_conteoData).length;
      alert(`✓ Importación completada (${modo}).
${nFinal} registros cargados en este dispositivo.`);

    } catch(err){
      alert("Error al leer el archivo: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ===== PENDIENTES DE CAPTURA / CONCILIACIÓN ===== */
const _C_EST = {
  cubierto:   {l:"✅ Cuadra gracias a los pendientes", c:"#15803d", bg:"#f0fdf4", o:1},
  parcial:    {l:"🟠 Pendientes cubren solo una parte", c:"#c2410c", bg:"#fff7ed", o:2},
  sinresp:    {l:"🔴 Faltante sin respaldo",             c:"#b91c1c", bg:"#fef2f2", o:3},
  exceso:     {l:"🟡 Pendientes de más (sobra)",        c:"#a16207", bg:"#fefce8", o:4},
  sobrante:   {l:"🟡 Sobrante sin pendientes",          c:"#a16207", bg:"#fefce8", o:5},
  sincontar:  {l:"⏳ Pendiente, material aún sin contar", c:"#475569", bg:"#f1f5f9", o:6},
  sinlote:    {l:"⚠ Pendiente sin lote (material con lotes)", c:"#9a3412", bg:"#fff7ed", o:7},
  noexiste:   {l:"❓ Catálogo no existe en D041",       c:"#6b21a8", bg:"#faf5ff", o:8},
  fueraarea:  {l:"↪ Pendiente de otra área",            c:"#64748b", bg:"#f8fafc", o:9}
};
function _cConciliar(){
  const esTodo = _conteoAreaActual==="Todo el inventario" || !_conteoAreaActual;
  const enArea = esTodo ? _conteoMats : _conteoMats.filter(m=>m.area===_conteoAreaActual);
  const filasArea = _cFilas(enArea);
  const todas = _cFilas(_conteoMats);
  const porKey = {}; todas.forEach(f=>porKey[f.key]=f);
  const keysArea = new Set(filasArea.map(f=>f.key));
  const matPorCat = {}; _conteoMats.forEach(m=>matPorCat[m.catalogo]=m);
  const out=[]; const usados=new Set();
  filasArea.forEach(f=>{
    const pen=_cPend(f.key); const fis=_conteoData[f.key]; const cap=fis!==undefined;
    usados.add(f.key);
    if(!cap && pen<=0) return;
    let est, nota="", difFinal=null;
    if(!cap){ est="sincontar"; nota="Cuéntalo para poder conciliar (pendiente: "+pen+")"; }
    else {
      const base = fis - f.tot; difFinal = base + pen;
      if(pen>0){
        if(base<0 && difFinal===0) { est="cubierto"; nota="Faltaban "+(-base)+" y los pendientes lo cubren"; }
        else if(base<0 && difFinal<0){ est="parcial"; nota="Aún faltan "+(-difFinal)+" sin respaldo"; }
        else if(difFinal>0){ est="exceso"; nota="Pendientes exceden por "+difFinal+" — revisar captura"; }
        else { est="cubierto"; nota="Pendientes completan el stock"; }
      } else {
        if(base<0){ est="sinresp"; nota="Faltan "+(-base)+" sin pendiente que lo explique"; }
        else if(base>0){ est="sobrante"; nota="Sobran "+base; }
        else return;
      }
    }
    out.push({key:f.key,cat:f.cat,lote:f.lote,desc:f.desc,um:f.um,tot:f.tot,fis,pen,difFinal,est,estLabel:_C_EST[est].l,nota});
  });
  // pendientes que no cayeron en el área
  Object.keys(_conteoPend).forEach(k=>{
    if(usados.has(k)) return;
    const pen=_conteoPend[k]; const cat=k.split("__")[0]; const lote=k.split("__")[1]||"";
    const m=matPorCat[cat];
    let est, nota;
    if(!m){ est="noexiste"; nota="Revisa el catálogo en tu tabla"; }
    else if(porKey[k] && !keysArea.has(k)){ est="fueraarea"; nota="Pertenece al área "+(m.area||"—")+" (cámbiate a esa área para conciliarlo)"; }
    else if(m.lotes && m.lotes.length && !lote){ est="sinlote"; nota="Indica el lote en la tabla de pendientes"; }
    else if(m.lotes && m.lotes.length && lote){ est="noexiste"; nota="El lote "+lote+" no existe en D041"; }
    else { est="fueraarea"; nota="Pertenece al área "+(m.area||"—"); }
    out.push({key:k,cat,lote,desc:m?m.descripcion:"(no encontrado)",um:m?m.um:"",tot:m?(porKey[k]?porKey[k].tot:0):0,fis:undefined,pen,difFinal:null,est,estLabel:_C_EST[est].l,nota});
  });
  out.sort((a,b)=>_C_EST[a.est].o-_C_EST[b.est].o || a.cat.localeCompare(b.cat));
  const cnt={}; out.forEach(f=>cnt[f.est]=(cnt[f.est]||0)+1);
  return {filas:out,cnt};
}

function _cPendAbrir(){
  _cPendCargar();
  let ov=document.getElementById("cPendModal"); if(ov) ov.remove();
  ov=document.createElement("div"); ov.id="cPendModal";
  ov.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:flex-start;justify-content:center;padding:16px;overflow:auto";
  const n=Object.keys(_conteoPend).length, nf=Object.keys(_conteoFolios).length;
  const ult=_cPendLog[0];
  const banner = ult ? `<div style="background:${ult.noEnc?"#fffbeb":"#f0fdf4"};border:2px solid ${ult.noEnc?"#f59e0b":"#22c55e"};border-radius:10px;padding:10px 12px;margin-bottom:10px">
      <div style="font-size:14px;font-weight:700;color:${ult.noEnc?"#92400e":"#15803d"}">✅ Agregado: ${ult.folios.map(_cEsc).join(", ")}</div>
      <div style="font-size:12.5px;color:#334155;margin-top:2px">${ult.partidas} partida${ult.partidas===1?"":"s"} · ${ult.piezas.toLocaleString()} pzas · ${ult.hora}${ult.omit?` · ${ult.omit} fila${ult.omit===1?"":"s"} omitida${ult.omit===1?"":"s"} (sin catálogo o cantidad)`:""}</div>
      ${ult.noEnc?`<div style="font-size:12px;color:#92400e;margin-top:3px">⚠ ${ult.noEnc} catálogo${ult.noEnc===1?"":"s"} no existe${ult.noEnc===1?"":"n"} en D041 (aparecerán marcados en la conciliación).</div>`:""}
      <div style="font-size:12px;color:#475569;margin-top:4px">Ya puedes cargar el siguiente folio abajo.</div>
    </div>` : "";
  const log = _cPendLog.length>1 ? `<div style="font-size:12px;color:var(--muted);margin:0 0 10px"><b>Agregado en esta sesión:</b>`+_cPendLog.map(e=>`<div>• ${e.hora} — ${e.folios.map(_cEsc).join(", ")}: ${e.partidas} partidas · ${e.piezas.toLocaleString()} pzas</div>`).join("")+`</div>` : "";
  ov.innerHTML=`<div style="background:white;border-radius:14px;max-width:760px;width:100%;padding:18px 18px 16px;margin:auto">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
      <h3 style="margin:0;font-size:17px;flex:1">⏳ Pendientes de captura</h3>
      <button class="btn btn-sm" onclick="document.getElementById('cPendModal').remove()">Listo</button>
    </div>
    <p style="margin:0 0 10px;font-size:12.5px;color:var(--muted);line-height:1.45">Material que ya salió o está por descontarse del stock pero aún no se captura en SAP. Se <b>suma al Físico</b>: <code>Diferencia = Físico + Pendiente − Total SAP</code>. Puedes cargar <b>varios folios seguidos</b>: llena el folio, pega la tabla, procesa y repite.</p>
    ${banner}${log}
    <div style="border:1px dashed var(--borde);border-radius:10px;padding:12px;margin-bottom:10px">
      <div style="font-size:13px;font-weight:600;margin-bottom:6px">Cargar un folio (Catálogo y Cantidad; Lote opcional)</div>
      <input id="cPendFolio" type="text" placeholder="Nombre / número de folio (ej. Ropa 4500123456). Si la tabla trae columna Folio, se usa esa." style="width:100%;box-sizing:border-box;font-size:12.5px;margin-bottom:8px">
      <textarea id="cPendTxt" rows="5" placeholder="Pega aquí la tabla copiada del correo/Excel (con o sin encabezados)" style="width:100%;box-sizing:border-box;font-size:12px;font-family:'IBM Plex Mono',monospace"></textarea>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:8px;font-size:12.5px">
        <span style="color:var(--muted)">o archivo:</span><input type="file" id="cPendFile" accept=".xlsx,.xls,.csv" onchange="_cPendDesdeArchivo(this)" style="font-size:12px;max-width:220px">
        <div style="flex:1"></div>
        <button class="btn btn-sm" onclick="_cPendProcesarTexto()">➕ Agregar folio</button>
      </div>
      <div style="display:flex;gap:14px;flex-wrap:wrap;margin-top:8px;font-size:12px;color:var(--muted)">
        <label title="Si el folio ya estaba cargado, se reemplaza (no se duplica)"><input type="radio" name="cPendModo" value="folio" ${_cPendModoSel==="folio"?"checked":""} onchange="_cPendModoSel=this.value"> Agregar (si el folio existe, lo reemplaza)</label>
        <label><input type="radio" name="cPendModo" value="sumar" ${_cPendModoSel==="sumar"?"checked":""} onchange="_cPendModoSel=this.value"> Sumar al folio</label>
        <label><input type="radio" name="cPendModo" value="todo" ${_cPendModoSel==="todo"?"checked":""} onchange="_cPendModoSel=this.value"> Reemplazar todo</label>
      </div>
      <div id="cPendMsg" style="font-size:12px;margin-top:8px"></div>
    </div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:8px">
      <div style="font-size:13px;font-weight:600;flex:1">Folios cargados: ${nf} · ${n} material${n===1?"":"es"} (totales sumados)</div>
      <button class="btn btn-outline btn-sm" onclick="_cConciliacionVer()" ${n?"":"disabled"}>📊 Ver conciliación</button>
      <button class="btn btn-outline btn-sm" style="color:#dc2626" onclick="_cPendLimpiar()" ${n?"":"disabled"}>🗑 Vaciar todo</button>
    </div>
    <div id="cPendCuerpo"></div>
  </div>`;
  document.body.appendChild(ov);
  _cPendRenderLista();
  setTimeout(()=>{ const f=document.getElementById("cPendFolio"); if(f && ult) f.focus(); },30);
}
function _cPendRenderLista(){
  const el=document.getElementById("cPendCuerpo"); if(!el) return;
  const folios=Object.keys(_conteoFolios).sort();
  if(!folios.length){ el.innerHTML='<div style="font-size:12px;color:var(--muted)">Aún no hay pendientes cargados.</div>'; return; }
  const matPorCat={}; _conteoMats.forEach(m=>matPorCat[m.catalogo]=m);
  el.innerHTML = folios.map(fo=>{
    const items=_conteoFolios[fo]; const ks=Object.keys(items).sort();
    const tot=ks.reduce((a,k)=>a+items[k],0);
    const nuevo=_cPendUltimos.includes(fo);
    return `<details ${nuevo?"open":""} style="border:1px solid ${nuevo?"#22c55e":"var(--borde)"};border-radius:8px;margin-bottom:6px;${nuevo?"background:#f0fdf4":""}">
      <summary style="padding:7px 10px;cursor:pointer;font-size:13px;display:flex;gap:10px;align-items:center;list-style-position:inside">
        <b style="flex:1">📄 ${_cEsc(fo)}${nuevo?' <span style="background:#22c55e;color:white;font-size:10px;padding:1px 6px;border-radius:8px;margin-left:4px">NUEVO</span>':""}</b><span style="color:var(--muted);font-size:12px">${ks.length} partidas · ${tot.toLocaleString()} pzas</span>
        <button class="btn btn-outline btn-sm" style="padding:0 8px;color:#dc2626" onclick="event.preventDefault();_cPendQuitarFolio(decodeURIComponent('${encodeURIComponent(fo)}'))">🗑 Quitar folio</button>
      </summary>
      <div style="max-height:220px;overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:12px"><tbody>`+
      ks.map(k=>{ const cat=k.split("__")[0], lote=k.split("__")[1]||""; const m=matPorCat[cat];
        return `<tr style="border-top:1px solid var(--borde)"><td style="padding:4px 8px;font-family:'IBM Plex Mono',monospace">${_cEsc(cat)}</td><td>${_cEsc(m?m.descripcion:"— no existe en D041 —")}</td><td>${lote?"L - "+_cEsc(lote):""}</td><td style="text-align:right;padding-right:8px;font-weight:600">${items[k].toLocaleString()}</td><td><button class="btn btn-outline btn-sm" style="padding:0 6px" onclick="_cPendQuitar(decodeURIComponent('${encodeURIComponent(fo)}'),decodeURIComponent('${encodeURIComponent(k)}'))">×</button></td></tr>`; }).join("")+
      `</tbody></table></div></details>`;
  }).join("");
}
function _cPendQuitar(folio,k){ _cPendUltimos=[]; _cPendLog=[]; if(_conteoFolios[folio]){ delete _conteoFolios[folio][k]; if(!Object.keys(_conteoFolios[folio]).length) delete _conteoFolios[folio]; } _cPendGuardar(); _cPendRefrescar(); }
function _cPendQuitarFolio(folio){ _cPendUltimos=[]; _cPendLog=[]; if(!confirm('¿Quitar el folio "'+folio+'" completo ('+Object.keys(_conteoFolios[folio]||{}).length+' partidas)?')) return; delete _conteoFolios[folio]; _cPendGuardar(); _cPendRefrescar(); }
function _cPendLimpiar(){ if(!confirm("¿Vaciar TODOS los folios de pendientes de captura?")) return; _conteoFolios={}; _cPendGuardar(); _cPendRefrescar(); }
function _cPendRefrescar(){
  if(document.getElementById("cPendModal")) _cPendAbrir();
  const b=document.getElementById("cPendBtn"); const n=Object.keys(_conteoPend).length;
  if(b) b.textContent="⏳ Pendientes de captura"+(n?` (${n})`:"");
  if(document.getElementById("conteo-lista")) filtrarConteo();
}
function _cPendNum(x){
  const t=String(x==null?"":x).trim().replace(/,/g,"");
  const m=t.match(/^(-?\d+(?:\.\d+)?)\s*(?:[a-záéíóúñ.\/]*)$/i);
  return m? parseFloat(m[1]) : null;
}
function _cPendParsearFilas(rows){
  const clean=rows.map(r=>(r||[]).map(c=>String(c==null?"":c).trim()));
  // buscar encabezado
  let h=-1, iC=-1, iQ=-1, iL=-1, iF=-1;
  for(let i=0;i<Math.min(clean.length,8);i++){
    const low=clean[i].map(c=>c.toLowerCase());
    const c=low.findIndex(x=>/^(cat[aá]logo|material|c[oó]digo|sku)/.test(x));
    if(c>=0){
      h=i; iC=c;
      iQ=low.findIndex(x=>/pend|cant|f[ií]sico|pzas|piezas|total|surtir/.test(x));
      iL=low.findIndex(x=>x==="lote");
      iF=low.findIndex(x=>/^(folio|documento|doc\.?)/.test(x));
      break;
    }
  }
  const res=[]; let omit=0;
  for(let i=h+1;i<clean.length;i++){
    const r=clean[i]; if(!r.some(x=>x)) continue;
    let cat="", cant=null, lote="", folio="";
    if(h>=0){
      cat=(r[iC]||"").replace(/\D/g,"");
      if(iQ>=0) cant=_cPendNum(r[iQ]);
      if(cant===null){ for(let j=r.length-1;j>iC;j--){ if(j===iL) continue; const v=_cPendNum(r[j]); if(v!==null){cant=v;break;} } }
      if(iL>=0) lote=r[iL]||"";
      if(iF>=0) folio=r[iF]||"";
    } else {
      const ic=r.findIndex(x=>/^\d{6,10}$/.test(x));
      if(ic<0){ omit++; continue; }
      cat=r[ic];
      for(let j=r.length-1;j>ic;j--){ const v=_cPendNum(r[j]); if(v!==null){cant=v;break;} }
    }
    if(!cat||cant===null||!(cant>0)){ omit++; continue; }
    res.push({cat,cant,lote,folio});
  }
  return {res,omit};
}
function _cPendTextoAFilas(txt){
  return txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>{
    if(l.includes("\t")) return l.split("\t");
    const m=l.trim().match(/^(\d{6,10})\s+(.*?)\s+([\d.,]+)\s*([a-záéíóúñ.\/]*)$/i);
    if(m) return [m[1],m[2],m[3]+(m[4]?" "+m[4]:"")];
    return l.trim().split(/\s{2,}|;/);
  });
}
function _cPendAplicar(res, omit){
  const modo=(document.querySelector('input[name="cPendModo"]:checked')||{}).value||_cPendModoSel||"folio";
  const base=((document.getElementById("cPendFolio")||{}).value||"").trim()||"Sin folio";
  if(modo==="todo") _conteoFolios={};
  const tocados=[]; let piezas=0;
  const existen=new Set(_conteoMats.map(m=>m.catalogo));
  const noEncSet=new Set();
  res.forEach(r=>{
    const fo=(r.folio||"").trim()||base;
    if(modo==="folio" && !tocados.includes(fo)){ delete _conteoFolios[fo]; }
    if(!tocados.includes(fo)) tocados.push(fo);
    const F=_conteoFolios[fo]||(_conteoFolios[fo]={});
    const k=r.lote?r.cat+"__"+r.lote:r.cat; F[k]=(F[k]||0)+r.cant;
    piezas+=r.cant;
    if(!existen.has(r.cat)) noEncSet.add(r.cat);
  });
  _cPendGuardar();
  _cPendUltimos=tocados;
  const hora=new Date().toLocaleTimeString("es-MX",{hour:"2-digit",minute:"2-digit"});
  _cPendLog.unshift({hora,folios:tocados,partidas:res.length,piezas,omit,noEnc:noEncSet.size});
  _cPendRefrescar();
  if(typeof _guiasToast==="function") try{ _guiasToast(`✓ Folio ${tocados.join(", ")} agregado (${res.length} partidas)`,"ok"); }catch(e){}
}
function _cPendProcesarTexto(){
  const txt=(document.getElementById("cPendTxt")||{}).value||"";
  const msg=document.getElementById("cPendMsg");
  if(!txt.trim()){ if(msg){msg.style.color="#b91c1c"; msg.textContent="Pega una tabla o elige un archivo.";} return; }
  const {res,omit}=_cPendParsearFilas(_cPendTextoAFilas(txt));
  if(!res.length){ if(msg){msg.style.color="#b91c1c"; msg.textContent="No encontré partidas con catálogo y cantidad. Revisa el formato.";} return; }
  _cPendAplicar(res,omit);
}
function _cPendDesdeArchivo(input){
  const f=input.files[0]; if(!f) return;
  const rd=new FileReader();
  rd.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:"array"});
      let rows=[]; 
      for(const sn of wb.SheetNames){ rows=rows.concat(XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,defval:""})); }
      const {res,omit}=_cPendParsearFilas(rows);
      const msg=document.getElementById("cPendMsg");
      if(!res.length){ if(msg){msg.style.color="#b91c1c"; msg.textContent="No encontré partidas con catálogo y cantidad en el archivo.";} return; }
      const fi=document.getElementById("cPendFolio"); if(fi && !fi.value.trim()) fi.value=f.name.replace(/\.[^.]+$/,"");
      _cPendAplicar(res,omit);
    }catch(err){ const msg=document.getElementById("cPendMsg"); if(msg){msg.style.color="#b91c1c"; msg.textContent="Error al leer el archivo: "+err.message;} }
  };
  rd.readAsArrayBuffer(f);
}

function _cConciliacionVer(){
  const R=_cConciliar();
  let ov=document.getElementById("cConcModal"); if(ov) ov.remove();
  ov=document.createElement("div"); ov.id="cConcModal";
  ov.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:10000;display:flex;align-items:flex-start;justify-content:center;padding:16px;overflow:auto";
  const contamos=(R.cnt.cubierto||0), noContamos=(R.cnt.parcial||0)+(R.cnt.sinresp||0);
  const chips=Object.keys(_C_EST).filter(k=>R.cnt[k]).map(k=>`<span style="background:${_C_EST[k].bg};color:${_C_EST[k].c};border:1px solid ${_C_EST[k].c}33;border-radius:14px;padding:3px 10px;font-size:12px;font-weight:600">${_C_EST[k].l}: ${R.cnt[k]}</span>`).join(" ");
  ov.innerHTML=`<div style="background:white;border-radius:14px;max-width:900px;width:100%;padding:18px;margin:auto">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:4px">
      <h3 style="margin:0;font-size:17px;flex:1">📊 Conciliación · ${_cEsc(_conteoAreaActual||"Todo el inventario")}</h3>
      <button class="btn btn-outline btn-sm" onclick="document.getElementById('cConcModal').remove()">✕</button>
    </div>
    <p style="margin:0 0 10px;font-size:12.5px;color:var(--muted)">Con qué <b style="color:#15803d">contamos</b>: <b>${contamos}</b> partida${contamos===1?"":"s"} que cuadra${contamos===1?"":"n"} al sumar los pendientes. Con qué <b style="color:#b91c1c">no contamos</b>: <b>${noContamos}</b> partida${noContamos===1?"":"s"} con faltante que los pendientes no explican${noContamos?" (por completo)":""}.</p>
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px">${chips||'<span style="font-size:12px;color:var(--muted)">Sin partidas para conciliar todavía. Captura físicos en esta área.</span>'}</div>
    <div style="max-height:62vh;overflow:auto;border:1px solid var(--borde);border-radius:8px"><table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead style="position:sticky;top:0;background:#f8fafc"><tr style="text-align:left"><th style="padding:6px 8px">Estado</th><th>Catálogo</th><th>Descripción</th><th style="text-align:right">SAP</th><th style="text-align:right">Físico</th><th style="text-align:right">Pend.</th><th style="text-align:right;padding-right:8px">Dif. final</th></tr></thead>
      <tbody>${R.filas.map(f=>{ const e=_C_EST[f.est]; return `<tr style="border-top:1px solid var(--borde);background:${e.bg}">
        <td style="padding:5px 8px;color:${e.c};font-weight:600;white-space:nowrap">${e.l}</td>
        <td style="font-family:'IBM Plex Mono',monospace">${_cEsc(f.cat)}${f.lote?`<div style="font-size:10px;color:#9a3412">L - ${_cEsc(f.lote)}</div>`:""}</td>
        <td>${_cEsc(f.desc)}<div style="font-size:11px;color:var(--muted)">${_cEsc(f.nota)}</div></td>
        <td style="text-align:right">${f.tot.toLocaleString()}</td>
        <td style="text-align:right">${f.fis===undefined?"—":f.fis.toLocaleString()}</td>
        <td style="text-align:right;color:#b45309;font-weight:600">${f.pen?("+"+f.pen.toLocaleString()):"—"}</td>
        <td style="text-align:right;padding-right:8px;font-weight:700;color:${f.difFinal===null?"#64748b":f.difFinal===0?"#15803d":f.difFinal<0?"#b91c1c":"#a16207"}">${f.difFinal===null?"—":(f.difFinal>0?"+":"")+f.difFinal}</td></tr>`; }).join("")}</tbody></table></div>
    <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:10px"><button class="btn btn-sm" onclick="exportarConteo()">⬇ Exportar Excel (incluye esta conciliación)</button></div>
  </div>`;
  document.body.appendChild(ov);
}

function limpiarConteo(){
  const n = Object.keys(_conteoData).length;
  if(n === 0){ alert("No hay datos capturados que limpiar."); return; }
  if(!confirm(
    `⚠️ ¿Estás seguro de que quieres limpiar el conteo?\n\n` +
    `Se borrarán ${n} registros capturados en este dispositivo.\n` +
    `Esta acción NO se puede deshacer.`
  )) return;
  _conteoData = {};
  _conteoLimpiarLS(_conteoAreaActual);
  filtrarConteo();
}

function exportarConteo(){
  if(Object.keys(_conteoData).length===0){ alert("No hay datos capturados."); return; }
  const hoy = new Date().toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"}).replace(/\//g,"-");
  const wb  = XLSX.utils.book_new();
  const filas = [];
  let capOk=0,capFalt=0,capSob=0,capPend=0;

  // Exportar solo los materiales de la lista filtrada actual, ordenados por catálogo
  const matExport = (_conteoLista.length > 0 ? _conteoLista : _conteoMats)
    .slice().sort((a,b)=>+a.catalogo - +b.catalogo);
  matExport.forEach(m=>{
    if(m.lotes && m.lotes.length>0){
      m.lotes.forEach(lote=>{
        const key=`${m.catalogo}__${lote.lote}`;
        const fis=_conteoData[key];
        const tot=_cTot(lote.lib,lote.tras,lote.transitoPt);
        const pn=_cPend(key);
        const dif=fis!==undefined?fis+pn-tot:"";
        if(fis!==undefined){if(dif===0)capOk++;else if(dif<0)capFalt++;else capSob++;}else capPend++;
        filas.push([m.catalogo,m.descripcion,m.um||"",m.ubicacion||"",lote.lote,lote.lib||0,lote.tras||0,lote.transitoPt||0,tot,fis!==undefined?fis:"",pn||"",fis!==undefined?fis+pn:"",dif]);
      });
    } else {
      const fis=_conteoData[m.catalogo];
      const lib=m.existencia||0,tra=m.traslado||0,tot=_cTot(lib,tra,m.transitoPt);
      const pn=_cPend(m.catalogo);
      const dif=fis!==undefined?fis+pn-tot:"";
      if(fis!==undefined){if(dif===0)capOk++;else if(typeof dif==="number"&&dif<0)capFalt++;else if(typeof dif==="number")capSob++;}else capPend++;
      filas.push([m.catalogo,m.descripcion,m.um||"",m.ubicacion||"","",lib,tra,m.transitoPt||0,tot,fis!==undefined?fis:"",pn||"",fis!==undefined?fis+pn:"",dif]);
    }
  });

  const wsData=[
    ["INVENTARIO FÍSICO — ALMACÉN D041 PUEBLA","","","","","","","","","","","",""],
    [`Fecha: ${hoy}   ·   Materiales: ${_conteoMats.length}   ·   Capturados: ${capOk+capFalt+capSob}   ·   Pendientes: ${capPend}   ·   Tránsito PT: ${_cIncTrans()?"incluido en el total":"no incluido"}`,"","","","","","","","","","","",""],
    [""],
    ["Catálogo","Descripción","U.M.","Ubicación","Lote","Lib. Utiliz.","Traslado","Tránsito PT",_cIncTrans()?"Total SAP (incl. tránsito PT)":"Total SAP","Físico","Pend. captura","Físico + pend.","Diferencia"],
    ...filas,
    [""],
    ["RESUMEN","","","","","","","","","",""],
    ["Sin diferencia",capOk,"","","Faltantes",capFalt,"","","",""],
    ["Sobrantes",capSob,"","","Pendientes",capPend,"","","",""],
  ];
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:12}},{s:{r:1,c:0},e:{r:1,c:12}}];
  ws["!cols"]=[{wch:12},{wch:40},{wch:7},{wch:16},{wch:14},{wch:12},{wch:10},{wch:12},{wch:14},{wch:10},{wch:13},{wch:13},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws,"Inventario");

  // Hoja solo diferencias
  const difData=[
    ["DIFERENCIAS — D041 — "+hoy,"","","","","","","","","","","",""],
    ["Catálogo","Descripción","U.M.","Ubicación","Lote","Lib. Utiliz.","Traslado","Tránsito PT",_cIncTrans()?"Total SAP (incl. tránsito PT)":"Total SAP","Físico","Pend. captura","Físico + pend.","Diferencia"],
    ...filas.filter(f=>f[12]!==""&&f[12]!==0).sort((a,b)=>(a[12]||0)-(b[12]||0))
  ];
  if(difData.length===2) difData.push(["Sin diferencias","","","","","","","","","","","",""]);
  const wsDif=XLSX.utils.aoa_to_sheet(difData);
  wsDif["!cols"]=ws["!cols"];
  XLSX.utils.book_append_sheet(wb,wsDif,"Solo Diferencias");

  // Hoja de conciliación de pendientes de captura
  if(Object.keys(_conteoPend).length){
    const R=_cConciliar();
    const cData=[["CONCILIACIÓN DE PENDIENTES DE CAPTURA — D041 — "+hoy,"","","","","","","",""],
      ["Estado","Catálogo","Descripción","Lote","Total SAP","Físico","Pend. captura","Diferencia final","Qué falta"],
      ...R.filas.map(f=>[f.estLabel,f.cat,f.desc,f.lote,f.tot,f.fis===undefined?"":f.fis,f.pen,f.difFinal===null?"":f.difFinal,f.nota])];
    const wsC=XLSX.utils.aoa_to_sheet(cData);
    wsC["!cols"]=[{wch:26},{wch:12},{wch:40},{wch:14},{wch:11},{wch:10},{wch:13},{wch:14},{wch:46}];
    XLSX.utils.book_append_sheet(wb,wsC,"Conciliación pendientes");
  }

  XLSX.writeFile(wb,`Inventario_D041_${hoy}.xlsx`);
}
