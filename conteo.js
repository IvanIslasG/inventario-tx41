// ═══════════════════════════════════════════════════════════════════════════
// MÓDULO: CONTEO FÍSICO — captura de inventario físico con diferencias en vivo
// Depende de helpers globales de index.html: $, DB, mat, nfmt, DIST
// ═══════════════════════════════════════════════════════════════════════════
/* ============ CONTEO FÍSICO ============ */
let _conteoData = {};   // cat (o cat__lote) → físico capturado
let _conteoLista = [];  // lista filtrada actual
let _conteoMats  = [];  // materiales del módulo actual (D041)

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
  // Construir lista completa de materiales de D041
  _conteoMats = [];
  const mats = DB.materiales || {};
  const exs  = DB.existencias || {};
  const tras  = DB.traslados || {};
  const lotes = DB.lotes || {};

  Object.keys(mats).forEach(cat => {
    const m   = mats[cat];
    const lib = (exs[cat] && exs[cat]["D041"]) || 0;
    const tra = (tras[cat] && tras[cat]["D041"]) || 0;
    const ls  = lotes[cat] || [];
    _conteoMats.push({
      catalogo:    cat,
      descripcion: m.desc || m.descripcion || "",
      um:          m.um || "",
      area:        m.area || "",
      ubicacion:   m.ubic || m.ubicacion || "",
      existencia:  lib,
      traslado:    tra,
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
      </select>
      <select id="conteo-ubic" onchange="filtrarConteo()" style="min-width:140px">
        <option value="">Todas las ubicaciones</option>
      </select>
      <label class="chk"><input type="checkbox" id="conteo-solo-ex" onchange="filtrarConteo()"> Solo con existencia</label>
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
      const total     = (m.existencia||0) + (m.traslado||0);
      const diferencia = cap ? fisico - total : null;
      if(dif==="pendiente")     mD = !cap;
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
          <div style="font-size:10px;color:var(--muted)">${m.ubicacion||"Sin ubicación"} · ${m.um||"—"}</div>
        </div>`;

      const tbody = document.createElement("div");
      tbody.style.padding = "8px";

      m.lotes.forEach((lote, idx)=>{
        const key  = `${m.catalogo}__${lote.lote}`;
        const fis  = _conteoData[key];
        const cap  = fis !== undefined;
        const tot  = (lote.lib||0) + (lote.tras||0);
        const dif  = cap ? fis - tot : null;
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
          <div style="flex:1"></div>
          <div class="cap-input-wrap">
            <input class="cap-input" type="number" inputmode="numeric" min="0" placeholder="—"
              value="${cap?fis:""}"
              data-cat="${m.catalogo}" data-lote="${lote.lote}" data-idx="${idx}"
              data-libre="${lote.lib||0}" data-traslado="${lote.tras||0}"
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
      const tot  = lib + tra;
      const dif  = cap ? fis - tot : null;
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
        <div class="cap-input-wrap">
          <input class="cap-input" type="number" inputmode="numeric" min="0" placeholder="—"
            value="${cap?fis:""}"
            data-cat="${m.catalogo}" data-libre="${lib}" data-traslado="${tra}"
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
  const tot = lib + tra;
  const val = input.value.trim();
  if(val===""||isNaN(val)) delete _conteoData[cat];
  else _conteoData[cat] = parseInt(val);
  const fis = _conteoData[cat];
  const cap = fis !== undefined;
  const dif = cap ? fis - tot : null;
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
  const tot = lib + tra;
  const key = `${cat}__${lote}`;
  const val = input.value.trim();
  if(val===""||isNaN(val)) delete _conteoData[key];
  else _conteoData[key] = parseInt(val);
  const fis = _conteoData[key];
  const cap = fis !== undefined;
  const dif = cap ? fis - tot : null;
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
        const tot=(lote.lib||0)+(lote.tras||0);
        if(fis===undefined){pen++;return;}
        cap++;
        const dif=fis-tot;
        if(dif===0)ok++;else if(dif<0)falt++;else sob++;
      });
    } else {
      const fis=_conteoData[m.catalogo];
      const tot=(m.existencia||0)+(m.traslado||0);
      if(fis===undefined){pen++;return;}
      cap++;
      const dif=fis-tot;
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
    scrollTimer = setTimeout(()=>{ g.classList.remove('visible'); }, 2500);
  }, {passive:true});

  // Mostrar siempre en móvil si se toca la pantalla
  document.addEventListener('touchstart', ()=>{
    g.classList.add('visible');
    clearTimeout(scrollTimer);
    scrollTimer = setTimeout(()=>{ g.classList.remove('visible'); }, 3000);
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
        const tot=(lote.lib||0)+(lote.tras||0);
        const dif=fis!==undefined?fis-tot:"";
        if(fis!==undefined){if(dif===0)capOk++;else if(dif<0)capFalt++;else capSob++;}else capPend++;
        filas.push([m.catalogo,m.descripcion,m.um||"",m.ubicacion||"",lote.lote,lote.lib||0,lote.tras||0,tot,fis!==undefined?fis:"",dif]);
      });
    } else {
      const fis=_conteoData[m.catalogo];
      const lib=m.existencia||0,tra=m.traslado||0,tot=lib+tra;
      const dif=fis!==undefined?fis-tot:"";
      if(fis!==undefined){if(dif===0)capOk++;else if(typeof dif==="number"&&dif<0)capFalt++;else if(typeof dif==="number")capSob++;}else capPend++;
      filas.push([m.catalogo,m.descripcion,m.um||"",m.ubicacion||"","",lib,tra,tot,fis!==undefined?fis:"",dif]);
    }
  });

  const wsData=[
    ["INVENTARIO FÍSICO — ALMACÉN D041 PUEBLA","","","","","","","","",""],
    [`Fecha: ${hoy}   ·   Materiales: ${_conteoMats.length}   ·   Capturados: ${capOk+capFalt+capSob}   ·   Pendientes: ${capPend}`,"","","","","","","","",""],
    [""],
    ["Catálogo","Descripción","U.M.","Ubicación","Lote","Lib. Utiliz.","Traslado","Total SAP","Físico","Diferencia"],
    ...filas,
    [""],
    ["RESUMEN","","","","","","","","",""],
    ["Sin diferencia",capOk,"","","Faltantes",capFalt,"","",""],
    ["Sobrantes",capSob,"","","Pendientes",capPend,"","",""],
  ];
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:9}},{s:{r:1,c:0},e:{r:1,c:9}}];
  ws["!cols"]=[{wch:12},{wch:40},{wch:7},{wch:16},{wch:14},{wch:12},{wch:10},{wch:12},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws,"Inventario");

  // Hoja solo diferencias
  const difData=[
    ["DIFERENCIAS — D041 — "+hoy,"","","","","","","","",""],
    ["Catálogo","Descripción","U.M.","Ubicación","Lote","Lib. Utiliz.","Traslado","Total SAP","Físico","Diferencia"],
    ...filas.filter(f=>f[9]!==""&&f[9]!==0).sort((a,b)=>(a[9]||0)-(b[9]||0))
  ];
  if(difData.length===2) difData.push(["Sin diferencias","","","","","","","",""]);
  const wsDif=XLSX.utils.aoa_to_sheet(difData);
  wsDif["!cols"]=ws["!cols"];
  XLSX.utils.book_append_sheet(wb,wsDif,"Solo Diferencias");

  XLSX.writeFile(wb,`Inventario_D041_${hoy}.xlsx`);
}
