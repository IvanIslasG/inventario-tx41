/* ============================================================
   or.js  -  Módulo Orden de Reabasto · Sistema TX41 Puebla
   Depende de: DB, mat(), ngDe(), ngLista(), almName(),
               almList(), DIST(), showMod(), $(), AREAS,
               XLSX (SheetJS)
   ============================================================ */

/* ============ ORDEN DE REABASTO ============ */
let orAlm="", orSort={col:"cat",dir:1}, _orManual={};

// Estados posibles por partida en el OR, con su color de referencia
const ESTADOS_OR = {
  pendiente: { label: "Pendiente",  color: "#6b7280", bg: "#f1f2f4" },
  validado:  { label: "Validado",   color: "#15803d", bg: "#e7f4ec" },
  parcial:   { label: "Parcial",    color: "#b45309", bg: "#fbf0df" },
  no_surtir: { label: "No surtir",  color: "#475569", bg: "#e2e8f0" },
  revisar:   { label: "Revisar",    color: "#c0392b", bg: "#fdeaea" }
};

// ── Persistencia OR en localStorage ──────────────────────────────────────────
const _OR_LS_KEY = () => `or_tx41_${orAlm}`;
const _OR_CFG_KEY = "or_tx41_config";

function _orGuardar(){
  try{
    localStorage.setItem(_OR_LS_KEY(), JSON.stringify(_orManual));
    // Guardar también almacén y parámetros
    const cfg = {
      alm:  orAlm,
      mCPM: +$("#orMCPM")?.value || 6,
      mStk: +$("#orMS")?.value   || 1.5,
      area: _orAreaSel,
    };
    localStorage.setItem(_OR_CFG_KEY, JSON.stringify(cfg));
  } catch(e){}
}

function _orCargar(){
  try{
    const raw = localStorage.getItem(_OR_LS_KEY());
    if(raw) _orManual = JSON.parse(raw);
    else _orManual = {};
  } catch(e){ _orManual = {}; }
}

function _orLimpiarLS(){
  try{
    localStorage.removeItem(_OR_LS_KEY());
    localStorage.removeItem(_OR_CFG_KEY);
  } catch(e){}
}

function _orTieneAvance(alm){
  try{
    const raw = localStorage.getItem(`or_tx41_${alm}`);
    if(!raw) return 0;
    return Object.keys(JSON.parse(raw)).length;
  } catch(e){ return 0; }
}

function _orCargarConfig(){
  try{ return JSON.parse(localStorage.getItem(_OR_CFG_KEY)||"{}"); }
  catch(e){ return {}; }
}
function consumosDisp(){ return DB.consumos ? Object.keys(DB.consumos).sort() : []; }

// ── Pantalla bienvenida OR ───────────────────────────────────────────────────
let _orAreaSel = "";  // área seleccionada desde el menú inicial

function modOR(){
  _orAreaSel = "";
  orAlm = "";
  _mostrarPaso1Alm();
}

// ── PASO 1: Selección de almacén ─────────────────────────────────────────────
function _tplPaso1Alm(alms){
  var tarjetas = "";
  for(var i=0; i<alms.length; i++){
    var alm = alms[i];
    var avance = _orTieneAvance(alm);
    var bg = avance ? "#f0fdf4" : "white";
    var bc = avance ? "#86efac" : "var(--line)";
    var subTexto = avance ? ("EN PROGRESO &middot; " + avance + " editados") : almName(alm);
    var subColor = avance ? "#16a34a" : "var(--muted)";

    tarjetas +=
      "<button data-alm=\"" + alm + "\" onclick=\"_mostrarBienvenidaOR('" + alm + "')\"" +
      " style=\"display:flex;flex-direction:column;align-items:flex-start;" +
      "gap:3px;padding:12px 14px;background:" + bg + ";border:1.5px solid " + bc + ";" +
      "border-radius:11px;cursor:pointer;text-align:left;width:100%;" +
      "font-family:inherit;transition:border-color .15s\"" +
      " onmouseover=\"this.style.borderColor='var(--primary)'\"" +
      " onmouseout=\"this.style.borderColor='" + bc + "'\">" +
      "<span style=\"font-size:14px;font-weight:700;color:var(--text)\">" + alm + "</span>" +
      "<span style=\"font-size:10px;color:" + subColor + "\">" + subTexto + "</span>" +
      "</button>";
  }

  return (
    "<div style=\"display:flex;flex-direction:column;align-items:center;justify-content:center;" +
    "min-height:70vh;padding:32px 20px;text-align:center\">" +
    "<div id=\"or-anim-wrap\" style=\"opacity:0;transform:translateY(20px);transition:all .6s ease\">" +
    "<div style=\"font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;" +
    "color:var(--muted);margin-bottom:10px\">Sistema de Inventario TX41</div>" +
    "<div style=\"font-size:22px;font-weight:800;color:var(--primary);margin-bottom:6px\">Orden de Reabasto</div>" +
    "<div style=\"font-size:14px;color:var(--text);font-weight:500;margin-bottom:4px\">Telmex RNUM</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">Almacen Distribuidor Puebla &middot; D041</div>" +
    "</div>" +
    "<div id=\"or-menu-alm\" style=\"opacity:0;transform:translateY(16px);transition:all .5s ease;" +
    "margin-top:28px;width:100%;max-width:720px\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);margin-bottom:12px;" +
    "text-transform:uppercase;letter-spacing:.5px\">&iquest;Con que almacen vas a trabajar?</div>" +
    "<input type=\"search\" id=\"or-alm-search\" placeholder=\"Buscar almacen...\"" +
    " style=\"width:100%;margin-bottom:14px;padding:9px 14px;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:13px;font-family:inherit;outline:none\"" +
    " oninput=\"_filtrarAlmMenu(this.value)\">" +
    "<div id=\"or-alm-grid\" style=\"display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px\">" +
    tarjetas +
    "</div></div></div>"
  );
}

function _mostrarPaso1Alm(){
  const alms = consumosDisp();
  const cfg  = _orCargarConfig();

  $("#moduleView").innerHTML = _tplPaso1Alm(alms);

  requestAnimationFrame(() => {
    setTimeout(() => {
      const a = document.getElementById("or-anim-wrap");
      if(a){ a.style.opacity="1"; a.style.transform="translateY(0)"; }
    }, 80);
    setTimeout(() => {
      const m = document.getElementById("or-menu-alm");
      if(m){ m.style.opacity="1"; m.style.transform="translateY(0)"; }
    }, 450);
  });
}

function _orAreaTieneAvance(area){
  // Verifica si hay entradas en _orManual para materiales de esta área
  return Object.keys(_orManual).some(key => {
    const cat = key.split("|")[1];
    if(!cat) return false;
    const m = mat(cat);
    const areaM = m?.area || "Sin clasificar";
    return !area || areaM === area;
  });
}

function _filtrarAlmMenu(q){
  const q2 = q.toLowerCase();
  document.querySelectorAll("#or-alm-grid button[data-alm]").forEach(btn => {
    const alm  = btn.dataset.alm.toLowerCase();
    const nom  = almName(btn.dataset.alm).toLowerCase();
    btn.style.display = (!q2 || alm.includes(q2) || nom.includes(q2)) ? "" : "none";
  });
}

// ── PASO 2: Selección de área ─────────────────────────────────────────────────
function _mostrarBienvenidaOR(almSeleccionado){
  if(almSeleccionado) orAlm = almSeleccionado;
  _orCargar();

  const cons    = DB.consumos?.[orAlm] || {};
  const criticos= DB.criticos || [];
  const todos   = new Set([...Object.keys(cons), ...criticos]);

  var areasList = AREAS.concat(["Sin clasificar"]);
  var areasHtml = "";
  for(var ai=0; ai<areasList.length; ai++){
    var area = areasList[ai];
    var nMats = 0;
    todos.forEach(function(cat){
      var info = mat(cat);
      var a = info.area || "Sin clasificar";
      if(a === area) nMats++;
    });
    var enProg = _orAreaTieneAvance(area);
    var bg = enProg ? "#f0fdf4" : "white";
    var bc = enProg ? "#86efac" : "var(--line)";
    var areaEsc = area.replace(/'/g, "&#39;");
    var rightLabel = enProg ? "EN PROGRESO" : (nMats + " materiales");
    var rightColor = enProg ? "#16a34a" : "var(--muted)";

    areasHtml += "<button onclick=\"_iniciarOR('" + areaEsc + "')\"" +
      " style=\"display:flex;align-items:center;justify-content:space-between;" +
      "gap:12px;padding:12px 18px;background:" + bg + ";border:1.5px solid " + bc + ";" +
      "border-radius:11px;cursor:pointer;text-align:left;width:100%;" +
      "font-family:inherit;transition:all .15s;font-size:14px;font-weight:600;color:var(--text)\"" +
      " onmouseover=\"this.style.borderColor='var(--primary)';this.style.color='var(--primary)'\"" +
      " onmouseout=\"this.style.borderColor='" + bc + "';this.style.color='var(--text)'\">" +
      "<span>" + area + "</span>" +
      "<span style=\"font-size:11px;font-weight:400;color:" + rightColor + "\">" + rightLabel + "</span>" +
      "</button>";
  }

  var headerHtml =
    "<div id=\"or-bienvenida\" style=\"display:flex;flex-direction:column;align-items:center;" +
    "min-height:70vh;padding:32px 20px;text-align:center\">" +
    "<div style=\"align-self:flex-start;margin-bottom:16px\">" +
    "<button onclick=\"_mostrarPaso1Alm()\" style=\"background:none;border:none;cursor:pointer;" +
    "font-size:13px;color:var(--muted);font-family:inherit;padding:4px 0\">&lsaquo; Cambiar almacen</button>" +
    "</div>" +
    "<div id=\"or-anim-wrap\" style=\"opacity:0;transform:translateY(20px);transition:all .6s ease\">" +
    "<div style=\"font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:2px;" +
    "color:var(--muted);margin-bottom:10px\">Orden de Reabasto &middot; " + orAlm + "</div>" +
    "<div style=\"font-size:20px;font-weight:800;color:var(--primary);margin-bottom:4px\">" +
    almName(orAlm) + "</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">Telmex RNUM &middot; Almacen Distribuidor Puebla</div>" +
    "</div>" +
    "<div id=\"or-menu-areas\" style=\"opacity:0;transform:translateY(16px);transition:all .5s ease;" +
    "margin-top:32px;width:100%;max-width:480px\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);margin-bottom:14px;" +
    "text-transform:uppercase;letter-spacing:.5px\">&iquest;Que area quieres trabajar?</div>" +
    "<div style=\"display:flex;flex-direction:column;gap:8px\">" +
    areasHtml +
    "<button onclick=\"_iniciarOR('')\" style=\"display:flex;align-items:center;" +
    "justify-content:space-between;padding:12px 18px;background:var(--primary);border:none;" +
    "border-radius:11px;cursor:pointer;text-align:left;width:100%;font-family:inherit;" +
    "font-size:14px;font-weight:700;color:white;margin-top:4px\">" +
    "<span>Todas las areas</span>" +
    "<span style=\"font-size:11px;font-weight:400;opacity:.8\">Vista completa</span>" +
    "</button>" +
    "</div></div></div>";

  $("#moduleView").innerHTML = headerHtml;

  requestAnimationFrame(function(){
    setTimeout(function(){
      var a = document.getElementById("or-anim-wrap");
      if(a){ a.style.opacity="1"; a.style.transform="translateY(0)"; }
    }, 80);
    setTimeout(function(){
      var m = document.getElementById("or-menu-areas");
      if(m){ m.style.opacity="1"; m.style.transform="translateY(0)"; }
    }, 450);
  });
}

function _iniciarOR(area){
  _orAreaSel = area;
  _modORPanel();
}

function _tplPanelOR(orAlm, mCPM, mStk, areaSel){
  return (
    "<div class=\"controls\" style=\"flex-wrap:wrap;align-items:center\">" +
    "<button class=\"btn btn-outline btn-sm\" onclick=\"_mostrarPaso1Alm()\">&lsaquo; Almacenes</button>" +
    "<button class=\"btn btn-outline btn-sm\" onclick=\"_mostrarBienvenidaOR()\">&lsaquo; Areas</button>" +
    "<span style=\"font-weight:700;color:var(--primary);font-size:13px\">" + orAlm + " &middot; " + almName(orAlm) + "</span>" +
    "<button class=\"btn btn-outline btn-sm\" style=\"color:var(--rojo,#dc2626)\" onclick=\"_orReiniciar()\">&#8634; Reiniciar</button>" +
    "</div>" +
    "<div class=\"controls\" style=\"flex-wrap:wrap;margin-top:6px;align-items:center\">" +
    "<label class=\"chk\">CPM (meses) <input type=\"number\" id=\"orMCPM\" value=\"" + mCPM + "\" min=\"1\" max=\"24\" step=\"1\" style=\"width:58px\"></label>" +
    "<label class=\"chk\">Stock obj. <input type=\"number\" id=\"orMS\" value=\"" + mStk + "\" min=\"0.5\" max=\"12\" step=\"0.5\" style=\"width:58px\"></label>" +
    "<span style=\"font-size:12px;font-weight:600;color:var(--primary);padding:4px 10px;" +
    "background:var(--lite,#f0f4ff);border-radius:8px\">Area: " + (areaSel || "Todas") + "</span>" +
    "<input type=\"search\" id=\"orSearch\" placeholder=\"Buscar catalogo / descripcion...\" style=\"min-width:180px\">" +
    "</div>" +
    "<div class=\"controls\" style=\"flex-wrap:wrap;margin-top:6px\">" +
    "<label class=\"chk\"><input type=\"checkbox\" id=\"orSolo\"> Solo Calc. surtir &gt; 0</label>" +
    "<label class=\"chk\"><input type=\"checkbox\" id=\"orSoloX\"> Solo X Surtir &gt; 0</label>" +
    "<label class=\"chk\"><input type=\"checkbox\" id=\"orExced\"> Solo excedentes</label>" +
    "<label class=\"chk\"><input type=\"checkbox\" id=\"orSoloD041\"> Solo con existencia D041</label>" +
    "<select id=\"orEstadoFiltro\" style=\"padding:5px 8px;border:1px solid var(--line);border-radius:7px;font-family:inherit;font-size:12.5px\">" +
    "<option value=\"\">Todos los estados</option>" +
    Object.keys(ESTADOS_OR).map(k=>`<option value="${k}">${ESTADOS_OR[k].label}</option>`).join("") +
    "</select>" +
    "<button class=\"btn-prim\" id=\"orExport\">&#8681; Exportar OR</button>" +
    "</div>" +
    "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin:8px 0;font-size:12px\">" +
    "<span class=\"pill\" id=\"orResumen\"></span>" +
    "<span style=\"color:var(--muted)\">CPM = consumo &divide; meses &middot; Calc. surtir = (meses stock &times; CPM) &minus; existencia aux &middot; <b>X Surtir</b> editable</span>" +
    "</div>" +
    "<div style=\"display:flex;gap:0;align-items:flex-start;margin-top:8px\">" +
    "<div id=\"orNGSidebar\" style=\"width:200px;flex-shrink:0;position:sticky;top:8px;" +
    "max-height:calc(100vh - 120px);overflow-y:auto;background:white;border:1px solid var(--line);" +
    "border-radius:12px;margin-right:12px;display:flex;flex-direction:column\">" +
    "<div style=\"padding:10px 12px;border-bottom:1px solid var(--line);display:flex;" +
    "align-items:center;justify-content:space-between;position:sticky;top:0;background:white;" +
    "z-index:1;border-radius:12px 12px 0 0\">" +
    "<span style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px\">Grupos</span>" +
    "<div style=\"display:flex;gap:6px\">" +
    "<button onclick=\"_orNGTodos(true)\" title=\"Marcar todos\" style=\"border:none;background:none;" +
    "cursor:pointer;font-size:11px;color:var(--primary);font-family:inherit;padding:0\">&check; Todos</button>" +
    "<span style=\"color:var(--line)\">|</span>" +
    "<button onclick=\"_orNGTodos(false)\" title=\"Desmarcar todos\" style=\"border:none;background:none;" +
    "cursor:pointer;font-size:11px;color:var(--muted);font-family:inherit;padding:0\">&times; Ninguno</button>" +
    "</div></div>" +
    "<div style=\"padding:6px 8px;border-bottom:1px solid var(--line)\">" +
    "<input type=\"search\" id=\"orNGSearch\" placeholder=\"Buscar grupo...\" oninput=\"_orNGFiltrar(this.value)\"" +
    " style=\"width:100%;padding:5px 8px;border:1px solid var(--line);border-radius:7px;" +
    "font-size:11px;font-family:inherit;outline:none\">" +
    "</div>" +
    "<div id=\"orNGList\" style=\"padding:6px 4px;flex:1;overflow-y:auto\"></div>" +
    "</div>" +
    "<div id=\"orNGToggleWrap\" style=\"display:none;margin-bottom:8px;width:100%\">" +
    "<button onclick=\"_orNGToggleMobile()\" style=\"width:100%;padding:8px 14px;background:white;" +
    "border:1.5px solid var(--line);border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;" +
    "font-family:inherit;color:var(--primary);text-align:left\">&#9660; Filtrar por grupo (NG)</button>" +
    "<div id=\"orNGMobilePanel\" style=\"display:none;background:white;border:1px solid var(--line);" +
    "border-radius:0 0 10px 10px;padding:8px 4px;max-height:200px;overflow-y:auto\"></div>" +
    "</div>" +
    "<div style=\"flex:1;min-width:0\">" +
    "<div id=\"orBandaNG\"></div>" +
    "<div class=\"panel\"><div class=\"panel-head\" style=\"gap:8px\">" +
    "<h2 id=\"orTitle\"></h2><span class=\"pill\" id=\"orCount\"></span></div>" +
    "<div style=\"overflow-x:auto\"><table id=\"orTable\" style=\"min-width:1100px\"></table></div>" +
    "</div></div></div>" +
    "<style>@media(max-width:768px){#orNGSidebar{display:none!important}#orNGToggleWrap{display:block!important}}</style>"
  );
}

function _modORPanel(){
  // Cargar avance guardado del almacén actual
  _orCargar();
  const hayAvance = Object.keys(_orManual).length > 0;
  const alms=consumosDisp();
  if(!alms.length){
    $("#moduleView").innerHTML="<div class=\"panel\"><div class=\"soon\">" +
      "<h2>Orden de Reabasto</h2>" +
      "<p>Aun no hay archivos de consumo cargados.</p>" +
      "<p>Ve a <b>Configuracion &rarr; Archivos de consumo</b>, sube el export de consumo (ej. <b>TX8A</b>) y publica.</p>" +
      "</div></div>";
    return;
  }
  // Respetar orAlm elegido en el menú; solo usar fallback si no hay ninguno
  if(!orAlm) orAlm = alms[0];
  const m=DB.meta, mCPM=m?.meses_cpm||6, mStk=m?.meses_stock||1.5;

  $("#moduleView").innerHTML = _tplPanelOR(orAlm, mCPM, mStk, _orAreaSel);

  ["orMCPM","orMS"].forEach(id=>{ const el=$("#"+id); el.oninput=pintarOR; el.onchange=pintarOR; });
  const orSearchEl=$("#orSearch"); if(orSearchEl){ orSearchEl.oninput=()=>{ _actualizarNGOpts(); pintarOR(); }; }
  ["orSolo","orSoloX","orExced","orSoloD041","orEstadoFiltro"].forEach(id=>{ const el=$("#"+id); if(el) el.onchange=pintarOR; });
  $("#orExport").onclick=_mostrarResumenOR;

  // Autocomplete nombre genérico — filtra por área activa
  function _actualizarNGOpts(){
    const fArea=_orAreaSel||"";
    const ngsActivos=new Set();
    for(const cat of Object.keys(DB.materiales||{})){
      const info=mat(cat);
      const area=info.area||"Sin clasificar";
      if(fArea && area!==fArea) continue;
      const ng=ngDe(cat);
      ngsActivos.add(ng||"SIN SUSTITUTO");
    }
  }
  _actualizarNGOpts();
  _orNGIniciado = false;
  setTimeout(_orNGRender, 100);

  pintarOR();
}

function _orReiniciar(){
  const n = Object.keys(_orManual).length;
  if(n === 0){ alert("No hay datos editados que reiniciar."); return; }
  if(!confirm("Reiniciar la OR de " + almName(orAlm) + "? Se borraran " + n + " valores editados. Esta accion no se puede deshacer.")) return;
  _orManual = {};
  _orLimpiarLS();
  pintarOR();
}

// Sidebar NG
let _orNGSeleccionados = new Set();
let _orNGIniciado = false;

function _orNGRender(){
  var fArea = _orAreaSel || "";
  var ngs = new Set();
  var cats = Object.keys(DB.materiales || {});
  for(var i=0; i<cats.length; i++){
    var cat = cats[i];
    var info = mat(cat);
    var area = info.area || "Sin clasificar";
    if(fArea && area !== fArea) continue;
    ngs.add(ngDe(cat) || "SIN SUSTITUTO");
  }
  var lista = Array.from(ngs).sort();
  if(!_orNGIniciado){
    _orNGSeleccionados = new Set(lista);
    _orNGIniciado = true;
  }

  // ¿Qué nombres genéricos tienen al menos un catálogo que necesita surtido? (mismo umbral que el resto del módulo)
  var necesitaMaterial = new Set();
  var todasFilas = calcularOR();
  for(var k=0; k<todasFilas.length; k++){
    var fr = todasFilas[k];
    if(fArea && fr.area !== fArea) continue;
    if(fr.calcSurtir > 0.5) necesitaMaterial.add(fr.ng || "SIN SUSTITUTO");
  }

  // Los que necesitan surtido van primero, para no tener que buscarlos entre los demás
  lista.sort(function(a,b){
    var na=necesitaMaterial.has(a), nb=necesitaMaterial.has(b);
    if(na!==nb) return na?-1:1;
    return a.localeCompare(b);
  });

  function pintar(containerId){
    var el = document.getElementById(containerId);
    if(!el) return;
    el.innerHTML = "";
    for(var j=0; j<lista.length; j++){
      var ng = lista[j];
      var necesita = necesitaMaterial.has(ng);

      var label = document.createElement("label");
      label.style.display = "flex";
      label.style.alignItems = "center";
      label.style.gap = "6px";
      label.style.padding = "5px 10px";
      label.style.cursor = "pointer";
      label.style.borderRadius = "7px";
      label.style.fontSize = "12px";
      if(necesita){
        label.style.background = "#fff1e6";
        label.style.borderLeft = "3px solid #e8590c";
      }

      var input = document.createElement("input");
      input.type = "checkbox";
      input.checked = _orNGSeleccionados.has(ng);
      input.setAttribute("data-ng", ng);
      input.style.accentColor = "var(--primary)";
      input.style.width = "14px";
      input.style.height = "14px";
      input.addEventListener("change", function(){ _orNGCambio(this); });

      var span = document.createElement("span");
      span.textContent = ng;
      if(necesita){ span.style.fontWeight = "700"; span.style.color = "#a8390f"; }

      label.appendChild(input);
      label.appendChild(span);
      if(necesita){
        var dot = document.createElement("span");
        dot.textContent = "●";
        dot.title = "Tiene material por surtir";
        dot.style.color = "#e8590c";
        dot.style.fontSize = "9px";
        dot.style.marginLeft = "auto";
        label.appendChild(dot);
      }
      el.appendChild(label);
    }
  }

  pintar("orNGList");
  pintar("orNGMobilePanel");
}

function _orNGCambio(chk){
  var ng = chk.getAttribute("data-ng");
  var all = document.querySelectorAll('input[data-ng="' + ng.replace(/"/g, '\\"') + '"]');
  for(var i=0; i<all.length; i++){ all[i].checked = chk.checked; }
  if(chk.checked) _orNGSeleccionados.add(ng);
  else _orNGSeleccionados.delete(ng);
  pintarOR();
}

function _orNGTodos(marcar){
  var fArea = _orAreaSel || "";
  var ngs = new Set();
  var cats = Object.keys(DB.materiales || {});
  for(var i=0; i<cats.length; i++){
    var cat = cats[i];
    var info = mat(cat);
    var area = info.area || "Sin clasificar";
    if(fArea && area !== fArea) continue;
    ngs.add(ngDe(cat) || "SIN SUSTITUTO");
  }
  if(marcar) _orNGSeleccionados = new Set(Array.from(ngs));
  else _orNGSeleccionados = new Set();
  _orNGRender();
  pintarOR();
}

function _orNGFiltrar(q){
  var qn = (q || "").toLowerCase();
  var labels = document.querySelectorAll("#orNGList label, #orNGMobilePanel label");
  for(var i=0; i<labels.length; i++){
    var lbl = labels[i];
    var input = lbl.querySelector("input[data-ng]");
    var ng = input ? (input.getAttribute("data-ng") || "") : "";
    var match = !qn || ng.toLowerCase().includes(qn);
    lbl.style.display = match ? "flex" : "none";
  }
}

function _orNGToggleMobile(){
  var panel = document.getElementById("orNGMobilePanel");
  var btn = document.querySelector("#orNGToggleWrap button");
  if(!panel) return;
  var open = panel.style.display === "none";
  panel.style.display = open ? "block" : "none";
  if(btn) btn.textContent = (open ? "▲" : "▼") + " Filtrar por grupo (NG)";
}

function calcularOR(){
  const D=DIST(), cons=DB.consumos?.[orAlm]||{};
  const mCPM=Math.max(1,+$("#orMCPM")?.value||6), mStk=Math.max(0.1,+$("#orMS")?.value||1.5);
  const out=[];

  const calcFila=(cat)=>{
    const consumo = cons[cat] || 0;
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
    return {cat,desc:info.desc,um:info.um,area:info.area||"Sin clasificar",ng,consumo,cpm,exAux,exD,calcSurtir,xsurtir,mdInv,excedente,obs:man.obs||"",estado:man.estado||"pendiente"};
  };

  // Todos los materiales del maestro — el consumo se usa para calcular, no para filtrar
  for(const cat of Object.keys(DB.materiales||{})){
    out.push(calcFila(cat));
  }

  return out;
}

function filasOR(){
  const q=($("#orSearch")?.value||"").trim().toLowerCase();
  const fArea=_orAreaSel||"";

  const soloCal=$("#orSolo")?.checked;
  const soloX=$("#orSoloX")?.checked;
  const soloExc=$("#orExced")?.checked;
  const fEstado=$("#orEstadoFiltro")?.value||"";

  let rows=calcularOR().filter(r=>{
    if(q && !(r.cat.toLowerCase().includes(q)||(r.desc||"").toLowerCase().includes(q))) return false;
    if(fArea && r.area!==fArea) return false;
    // Filtro sidebar NG — si hay selección parcial, filtrar
    if(_orNGSeleccionados.size > 0 && !_orNGSeleccionados.has(r.ng)) return false;

    if(soloCal && r.calcSurtir<=0) return false;
    if(soloX && r.xsurtir<=0) return false;
    if(soloExc && !r.excedente) return false;
    if(fEstado && r.estado!==fEstado) return false;
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
  const conteoEstados={};
  rows.forEach(r=>{ conteoEstados[r.estado]=(conteoEstados[r.estado]||0)+1; });
  const resumenEstados=Object.keys(ESTADOS_OR)
    .filter(k=>conteoEstados[k])
    .map(k=>`<span style="color:${ESTADOS_OR[k].color};font-weight:700">${conteoEstados[k]} ${ESTADOS_OR[k].label.toLowerCase()}</span>`)
    .join(" · ");
  $("#orResumen").innerHTML=`${nNec} a surtir · Total X surtir: ${nfmt(totalX)}` + (resumenEstados?` · ${resumenEstados}`:"");

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
    <col style="width:55px"><col style="width:120px"><col>
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
    <th>Exced.</th><th data-c="estado">Estado</th><th>Observaciones</th>
  </tr></thead><tbody>${
  rows.length? rows.map(r=>{
    const nec=r.calcSurtir>0.5;
    const esSinSust=r.ng==="SIN SUSTITUTO";
    const rowBg=nec?`style="background:#fdf1f1"`:(r.excedente?`style="background:var(--low-bg)"`:"");
    // resaltar ng cuando es sustituto (no SIN SUSTITUTO)
    const ngStyle=esSinSust?"color:var(--muted);font-size:12px":"color:#7a5c00;font-size:12px;font-weight:700";
    const est=ESTADOS_OR[r.estado]||ESTADOS_OR.pendiente;
    const estadoOpts=Object.keys(ESTADOS_OR).map(k=>`<option value="${k}" ${k===r.estado?"selected":""}>${ESTADOS_OR[k].label}</option>`).join("");
    return `<tr ${rowBg} style="border-left:4px solid ${est.color}">
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
      <td style="text-align:center">${r.excedente?"":""}</td>
      <td>
        <select class="or-estado" data-cat="${r.cat}" style="padding:4px 6px;border:1.5px solid ${est.color};
          border-radius:6px;background:${est.bg};color:${est.color};font-weight:700;font-size:11.5px;font-family:inherit">
          ${estadoOpts}
        </select>
      </td>
      <td><input type="text" class="or-obs" data-cat="${r.cat}" value="${r.obs.replace(/"/g,"&quot;")}" placeholder="…"
          style="min-width:100px;padding:4px 6px;border:1px solid var(--line);border-radius:6px"></td>
    </tr>`;
  }).join("") : `<tr><td colspan="15" class="empty">Sin catálogos con estos filtros.</td></tr>`}</tbody>`;

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
      _orGuardar();
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
    inp.addEventListener("input",()=>{ (_orManual[orAlm+"|"+inp.dataset.cat]||={xs:0,obs:""}).obs=inp.value; _orGuardar(); });
  });

  // Guardar cambios de estado por partida
  $("#orTable").querySelectorAll(".or-estado").forEach(sel=>{
    sel.addEventListener("change",()=>{
      (_orManual[orAlm+"|"+sel.dataset.cat]||={xs:0,obs:""}).estado=sel.value;
      _orGuardar();
      pintarOR();
    });
  });
}

function _htmlBandaNG(sumNG, open=true){
  const entries=Object.entries(sumNG).sort((a,b)=>b[1].calc-a[1].calc);
  const n=entries.length;

  var resumenParts = [];
  for(var i=0; i<Math.min(4,entries.length); i++){
    var ng = entries[i][0], v = entries[i][1];
    resumenParts.push(
      "<span style=\"font-weight:700;color:#5a3e00\">" + ng + "</span> " +
      "<span style=\"color:#c0392b;margin-left:3px\">Calc: <b>" + fmt1(v.calc) + "</b></span>"
    );
  }
  var resumen = resumenParts.join(" &middot; ");
  if(n>4) resumen += " <span style=\"color:var(--muted)\">+" + (n-4) + " mas</span>";

  var chipsHtml = "";
  for(var j=0; j<entries.length; j++){
    var ngc = entries[j][0], vc = entries[j][1];
    chipsHtml +=
      "<div title=\"" + vc.cats.join(", ") + "\" style=\"background:#fff;border:1px solid #e8d080;" +
      "border-radius:8px;padding:4px 10px;font-size:12px;cursor:default;white-space:nowrap\">" +
      "<span style=\"font-weight:700;color:#5a3e00\">" + ngc + "</span>" +
      "<span style=\"color:#7a5c00;margin-left:5px\">Calc: <b>" + fmt1(vc.calc) + "</b></span>" +
      "<span style=\"color:#0a4ea3;margin-left:5px\">X: <b>" + nfmt(vc.xs) + "</b></span>" +
      "<span style=\"color:var(--muted);font-size:10.5px;margin-left:4px\">" + vc.cats.length + " cat.</span>" +
      "</div>";
  }

  return (
    "<details " + (open?"open":"") + " style=\"background:#fffbe8;border:1px solid #f0d060;" +
    "border-radius:10px;margin-bottom:8px\">" +
    "<summary style=\"list-style:none;padding:7px 12px;cursor:pointer;display:flex;" +
    "align-items:center;gap:10px;user-select:none\">" +
    "<span style=\"font-size:12px;font-weight:800;color:#7a5c00;white-space:nowrap\">&Sigma; Sustitutos</span>" +
    "<span style=\"font-size:11px;color:#b8960a;white-space:nowrap\">" + n + " grupo" + (n!==1?"s":"") + "</span>" +
    "<span style=\"flex:1\"></span>" +
    "<span class=\"ng-resumen\" style=\"font-size:13px;overflow:hidden;text-overflow:ellipsis;" +
    "white-space:nowrap;text-align:right\">" + resumen + "</span>" +
    "<span class=\"ng-arrow\" style=\"font-size:13px;color:#b8960a;flex-shrink:0;margin-left:8px\">&#9662;</span>" +
    "</summary>" +
    "<div style=\"padding:6px 10px 10px;display:flex;flex-wrap:wrap;gap:5px\">" + chipsHtml + "</div>" +
    "</details>"
  );
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

  var areasChecksHtml = "";
  for(var i=0; i<areasPresentes.length; i++){
    var a = areasPresentes[i];
    areasChecksHtml +=
      "<label class=\"chk\" style=\"background:#f7f9fc;padding:5px 9px;border:1px solid var(--line);" +
      "border-radius:8px\"><input type=\"checkbox\" class=\"or-area-chk\" data-a=\"" + a + "\" checked> " + a + "</label>";
  }

  const modal=document.createElement("div"); modal.className="modal on"; modal.id="orModal";
  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:400px\">" +
    "<h3>&#8681; Exportar Orden de Reabasto</h3>" +
    "<div class=\"modal-body\" style=\"gap:14px\">" +
    "<div><label style=\"display:block;font-size:12px;font-weight:700;color:var(--muted);" +
    "margin-bottom:5px\">Numero de Reabasto</label>" +
    "<input type=\"number\" id=\"orNumReabasto\" min=\"1\" placeholder=\"Ej. 5\" style=\"width:100%\"></div>" +
    "<div><label style=\"display:block;font-size:12px;font-weight:700;color:var(--muted);" +
    "margin-bottom:5px\">&iquest;Quien genera la OR?</label>" +
    "<input type=\"text\" id=\"orGenerador\" placeholder=\"Nombre del responsable\" style=\"width:100%\"></div>" +
    "<div><label style=\"display:block;font-size:12px;font-weight:700;color:var(--muted);" +
    "margin-bottom:5px\">Areas a exportar</label>" +
    "<div style=\"display:flex;flex-wrap:wrap;gap:6px\">" + areasChecksHtml + "</div></div>" +
    "<p style=\"font-size:12px;color:var(--muted);margin:0\">" + nfmt(rows.length) + " catalogos &middot; CPM " + mCPM + "m &middot; Stock obj. " + mStk + "m</p>" +
    "</div>" +
    "<div class=\"modal-foot\" style=\"gap:8px\">" +
    "<button class=\"btn\" id=\"orModalCancel\">Cancelar</button>" +
    "<button class=\"btn-prim\" id=\"orModalOk\">&#8681; Exportar Excel</button>" +
    "</div></div>";

  document.body.appendChild(modal);
  $("#orModalCancel").onclick=()=>modal.remove();

  const getParams=()=>{
    const num=$("#orNumReabasto").value.trim()||"—";
    const gen=$("#orGenerador").value.trim()||"";
    const selAreas=[...modal.querySelectorAll(".or-area-chk:checked")].map(c=>c.dataset.a);
    return {num, gen, selAreas};
  };

  $("#orModalOk").onclick=()=>{
    const {num,gen,selAreas}=getParams();
    if(!selAreas.length){ alert("Selecciona al menos un área."); return; }
    const rowsFilt=rows.filter(r=>selAreas.includes(r.area));
    modal.remove();
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
      wsaSAP["!cols"]=colWidths;
      XLSX.utils.book_append_sheet(wb, wsaSAP, "SAP MIGO 313");
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

function _ejecutarExportOR(numOR, generador){
  const n = parseInt(numOR)||1;
  const g = (generador||"").trim() || "—";
  // Llamar directamente a exportar sin el modal
  const rows = filasOR().filter(r=>r.xsurtir>0);
  const areas = [...new Set(rows.map(r=>r.area))];
  if(!rows.length){ alert("Sin materiales con X Surtir > 0."); return; }
  _exportarORSimple(rows, n, g, areas);
}

/* ── Resumen antes de exportar ── */
function _tplFilaMatResumen(r){
  var obs = r.obs || "&mdash;";
  return (
    "<tr style=\"border-bottom:1px solid var(--lite,#f4f6fb)\">" +
    "<td style=\"padding:7px 10px;font-size:12px;font-family:monospace;" +
    "font-weight:700;color:var(--primary);white-space:nowrap\">" + r.cat + "</td>" +
    "<td style=\"padding:7px 10px;font-size:12px;color:var(--text)\">" + r.desc + "</td>" +
    "<td style=\"padding:7px 10px;font-size:13px;font-weight:700;text-align:center;" +
    "color:var(--primary);white-space:nowrap\">" + r.xsurtir + "</td>" +
    "<td style=\"padding:7px 10px;font-size:11px;color:var(--muted);font-style:italic\">" + obs + "</td>" +
    "</tr>"
  );
}

function _tplStatCard(n, l){
  return (
    "<div style=\"background:white;border:1px solid var(--line);border-radius:12px;" +
    "padding:14px 16px;text-align:center\">" +
    "<div style=\"font-size:24px;font-weight:800;color:var(--primary)\">" + n + "</div>" +
    "<div style=\"font-size:11px;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;margin-top:2px\">" + l + "</div>" +
    "</div>"
  );
}

function _tplResumenOR(orAlm, totalPzas, totalCats, ramaLabel, matRowsHtml){
  return (
    "<div style=\"max-width:680px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:24px\">" +
    "<button onclick=\"_modORPanel()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Regresar</button>" +
    "<div><div style=\"font-size:18px;font-weight:800;color:var(--primary)\">Resumen de la OR</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">" + orAlm + " &middot; " + almName(orAlm) + "</div></div>" +
    "</div>" +
    "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px\">" +
    _tplStatCard(totalPzas.toLocaleString(), "Total piezas") +
    _tplStatCard(totalCats, "Catalogos") +
    "</div>" +
    "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:20px\">" +
    "<div><label style=\"font-size:11px;font-weight:700;color:var(--muted);" +
    "text-transform:uppercase;letter-spacing:.4px;display:block;margin-bottom:6px\">No. de Reabasto</label>" +
    "<input id=\"orNumResum\" type=\"number\" min=\"1\" placeholder=\"Ej. 5\" style=\"width:100%;" +
    "padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;font-size:16px;" +
    "font-weight:700;font-family:inherit;color:var(--primary);outline:none\"" +
    " onfocus=\"this.style.borderColor='var(--primary)'\" onblur=\"this.style.borderColor='var(--line)'\">" +
    "</div>" +
    "<div><label style=\"font-size:11px;font-weight:700;color:var(--muted);" +
    "text-transform:uppercase;letter-spacing:.4px;display:block;margin-bottom:6px\">Genera OR</label>" +
    "<input id=\"orGenResum\" type=\"text\" placeholder=\"Nombre completo\" style=\"width:100%;" +
    "padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;font-size:14px;" +
    "font-family:inherit;color:var(--text);outline:none\"" +
    " onfocus=\"this.style.borderColor='var(--primary)'\" onblur=\"this.style.borderColor='var(--line)'\">" +
    "</div></div>" +
    "<div style=\"background:#f0f4ff;border:1.5px solid #c7d7f0;border-radius:10px;" +
    "padding:12px 16px;margin-bottom:20px;display:flex;align-items:center;gap:10px\">" +
    "<span style=\"font-size:18px\">&#128203;</span>" +
    "<div><div style=\"font-size:12px;font-weight:700;color:var(--primary)\">Ruta SAP</div>" +
    "<div style=\"font-size:12px;color:var(--text)\">" + ramaLabel + "</div></div>" +
    "</div>" +
    "<div style=\"background:white;border:1px solid var(--line);border-radius:12px;" +
    "overflow:hidden;margin-bottom:24px\">" +
    "<div style=\"padding:12px 16px;border-bottom:1px solid var(--line);font-size:12px;" +
    "font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.4px\">" +
    "Materiales a surtir &middot; ordenados por catalogo</div>" +
    "<div style=\"max-height:360px;overflow-y:auto\">" +
    "<table style=\"width:100%;border-collapse:collapse\">" +
    "<thead style=\"position:sticky;top:0;background:var(--lite,#f4f6fb);z-index:1\"><tr>" +
    "<th style=\"padding:8px 10px;font-size:11px;color:var(--muted);text-align:left;" +
    "font-weight:600;white-space:nowrap\">CATALOGO</th>" +
    "<th style=\"padding:8px 10px;font-size:11px;color:var(--muted);text-align:left;" +
    "font-weight:600\">DESCRIPCION</th>" +
    "<th style=\"padding:8px 10px;font-size:11px;color:var(--muted);text-align:center;" +
    "font-weight:600;white-space:nowrap\">X SURTIR</th>" +
    "<th style=\"padding:8px 10px;font-size:11px;color:var(--muted);text-align:left;" +
    "font-weight:600\">OBSERVACIONES</th>" +
    "</tr></thead><tbody>" + matRowsHtml + "</tbody></table></div></div>" +
    "<button onclick=\"_ejecutarExportOR($('#orNumResum').value, $('#orGenResum').value)\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;" +
    "transition:background .15s\"" +
    " onmouseover=\"this.style.background='#0033AA'\" onmouseout=\"this.style.background='var(--primary)'\">" +
    "&#8681; Confirmar y Exportar OR</button>" +
    "</div>"
  );
}

function _mostrarResumenOR(){
  const rows = filasOR();
  const conXS = rows.filter(r => r.xsurtir > 0);
  if(!conXS.length){ alert("No hay materiales con X Surtir > 0. Agrega cantidades antes de exportar."); return; }

  const totalCats = conXS.length;
  const totalPzas = conXS.reduce((s,r) => s + r.xsurtir, 0);

  const siglaDestino = orAlm;
  const centroDestino = (DB.directorio?.almacenes?.[siglaDestino]?.centro) || "";
  const esRN58 = centroDestino === "RN58";
  const ramaLabel = esRN58 ? "MIGO 313 (mismo centro)" : "ME21N + MIGO 351";

  var matRowsHtml = "";
  conXS.sort((a,b) => a.cat.localeCompare(b.cat)).forEach(function(r){
    matRowsHtml += _tplFilaMatResumen(r);
  });

  $("#moduleView").innerHTML = _tplResumenOR(orAlm, totalPzas, totalCats, ramaLabel, matRowsHtml);
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


