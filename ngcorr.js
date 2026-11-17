/* ============================================================
   ngcorr.js — Correcciones de "Nombre genérico" dentro de la app
   ------------------------------------------------------------
   · Las correcciones se guardan en el repo: ng_correcciones.json
   · Se aplican POR ENCIMA del Excel (NombresGenericos.xlsx): ngDe()/ngLista()
     devuelven el valor corregido; todo el sistema (OR, análisis…) lo respeta.
   · Almacenamiento con ADAPTADORES (NGC_STORES). Hoy: "github" (token).
     Mañana: se agrega un adaptador "db" (Firebase/Supabase/API propia) con
     la misma interfaz y se cambia NGC_STORE_ACTIVO — nada más cambia.
       interfaz:  leer()  -> Promise<{version,actualizado,por,correcciones}>
                  cambiar(cambios, mensaje) -> Promise<doc nuevo>
                  (cambios = { catalogo: entrada | null }, null = revertir)
                  listo()  -> bool (¿se puede escribir?)
   ============================================================ */
(function(){
"use strict";

const NGC_ARCHIVO = "ng_correcciones.json";
const K = { cache:"ngc_cache_v1", cola:"ngc_cola_v1", user:"ngc_usuario", tok:"ngc_token_v1" };
const NGC_STORE_ACTIVO = "github";

/* ---------- utilidades ---------- */
const lsGet=(k,d)=>{ try{ const v=localStorage.getItem(k); return v==null?d:JSON.parse(v); }catch(e){ return d; } };
const lsSet=(k,v)=>{ try{ localStorage.setItem(k,JSON.stringify(v)); }catch(e){} };
const lsDel=k=>{ try{ localStorage.removeItem(k); }catch(e){} };
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const ahora=()=>new Date().toISOString();
const docVacio=()=>({version:1,actualizado:"",por:"",correcciones:{}});
const norm=s=>{ s=String(s==null?"":s).trim().replace(/\s+/g," "); return s.toUpperCase()==="SIN SUSTITUTO"?"":s; };

function b64enc(bytes){ let s=""; for(let i=0;i<bytes.length;i+=0x8000) s+=String.fromCharCode.apply(null,bytes.subarray(i,i+0x8000)); return btoa(s); }
function b64dec(b64){ const s=atob(b64.replace(/\s/g,"")); const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++) u[i]=s.charCodeAt(i); return u; }
const utf8enc=t=>new TextEncoder().encode(t);
const utf8dec=u=>new TextDecoder().decode(u);

/* ---------- estado ---------- */
let cache = lsGet(K.cache, null) || docVacio();
let cola  = lsGet(K.cola, {})  || {};           // cambios aún no subidos
let tokenSesion = null;                          // token descifrado en memoria
let estado = { msg:"", tipo:"", sync:false };
let _vers = 0, _listaCache = null, _listaVers = -1, _listaRef = null;

function corrEfectivas(){
  const o = Object.assign({}, cache.correcciones||{});
  for(const c in cola){ if(cola[c]===null) delete o[c]; else o[c]=cola[c]; }
  return o;
}
function bump(){ _vers++; }

/* ---------- reemplazo de ngDe / ngLista ---------- */
const ngDeOrig   = (typeof window.ngDe==="function")   ? window.ngDe   : (c=>(DB.nombresGenericos?.mapa||{})[c]||"");
const ngListaOrig= (typeof window.ngLista==="function")? window.ngLista: (()=>DB.nombresGenericos?.lista||[]);

function ngDeNuevo(cat){
  const c = corrEfectivas()[cat];
  if(c) return c.ng||"";
  return ngDeOrig(cat);
}
function ngListaNueva(){
  const base = ngListaOrig();
  if(_listaCache && _listaVers===_vers && _listaRef===base) return _listaCache;
  const s = new Set(base);
  const ef = corrEfectivas();
  for(const c in ef){ if(ef[c].ng) s.add(ef[c].ng); }
  // quitar nombres que quedaron sin ningún catálogo
  const usados = new Set();
  const mapa = DB.nombresGenericos?.mapa||{};
  for(const c in mapa){ if(!(c in ef) && mapa[c]) usados.add(mapa[c]); }
  for(const c in ef){ if(ef[c].ng) usados.add(ef[c].ng); }
  const res=[...s].filter(n=>n && n!=="SIN SUSTITUTO" && (usados.has(n)||!Object.keys(mapa).length));
  res.sort((a,b)=>a.localeCompare(b,"es"));
  _listaCache=res; _listaVers=_vers; _listaRef=base;
  return res;
}
window.ngDe = ngDeNuevo;
window.ngLista = ngListaNueva;
// el código existente llama a ngDe()/ngLista() como globales -> también función global declarada

/* ---------- repo ---------- */
function repoInfo(){
  let owner="IvanIslasG", repo="inventario-tx41";
  try{
    const h=location.hostname||"";
    if(/\.github\.io$/i.test(h)){
      owner=h.split(".")[0];
      const p=(location.pathname||"").split("/").filter(Boolean)[0];
      if(p && !/\.(html|js)$/i.test(p)) repo=p;
    }
  }catch(e){}
  return { owner, repo, rama:"main" };
}

/* ---------- token (con PIN opcional) ---------- */
async function derivar(pin, salt){
  const km=await crypto.subtle.importKey("raw",utf8enc(pin),"PBKDF2",false,["deriveKey"]);
  return crypto.subtle.deriveKey({name:"PBKDF2",salt,iterations:200000,hash:"SHA-256"},km,{name:"AES-GCM",length:256},false,["encrypt","decrypt"]);
}
async function guardarToken(token, pin){
  if(!pin){ lsSet(K.tok,{t:token}); tokenSesion=token; return; }
  const salt=crypto.getRandomValues(new Uint8Array(16)), iv=crypto.getRandomValues(new Uint8Array(12));
  const key=await derivar(pin,salt);
  const ct=new Uint8Array(await crypto.subtle.encrypt({name:"AES-GCM",iv},key,utf8enc(token)));
  lsSet(K.tok,{enc:1,salt:b64enc(salt),iv:b64enc(iv),ct:b64enc(ct)});
  tokenSesion=token;
}
async function abrirToken(pin){
  const r=lsGet(K.tok,null); if(!r) return null;
  if(!r.enc) return r.t||null;
  const key=await derivar(pin,b64dec(r.salt));
  const pt=await crypto.subtle.decrypt({name:"AES-GCM",iv:b64dec(r.iv)},key,b64dec(r.ct));
  return utf8dec(new Uint8Array(pt));
}
const hayTokenGuardado=()=>!!lsGet(K.tok,null);
const tokenCifrado=()=>!!(lsGet(K.tok,null)||{}).enc;

async function obtenerToken(){
  if(tokenSesion) return tokenSesion;
  const r=lsGet(K.tok,null); if(!r) return null;
  if(!r.enc){ tokenSesion=r.t; return tokenSesion; }
  for(let i=0;i<3;i++){
    const pin=await pedirTexto("PIN para usar el token","Escribe tu PIN (se pide una vez por sesión).",true);
    if(pin==null) return null;
    try{ tokenSesion=await abrirToken(pin); return tokenSesion; }
    catch(e){ aviso("PIN incorrecto","err"); }
  }
  return null;
}

/* ---------- adaptador GitHub ---------- */
const GH = {
  async leer(){
    const t=tokenSesion||null;
    const ri=repoInfo();
    // lectura pública (relativa al sitio); si hay token se usa la API para ver lo último sin esperar al CDN
    if(t){
      try{
        const r=await fetch(`https://api.github.com/repos/${ri.owner}/${ri.repo}/contents/${NGC_ARCHIVO}?ref=${ri.rama}`,
          {headers:{Authorization:"Bearer "+t,Accept:"application/vnd.github+json"},cache:"no-store"});
        if(r.ok){ const j=await r.json(); return JSON.parse(utf8dec(b64dec(j.content))); }
        if(r.status===404) return docVacio();
      }catch(e){}
    }
    const r=await fetch(NGC_ARCHIVO+"?_="+Date.now(),{cache:"no-store"});
    if(r.status===404) return docVacio();
    if(!r.ok) throw new Error("HTTP "+r.status);
    return await r.json();
  },
  listo(){ return !!tokenSesion || hayTokenGuardado(); },
  async cambiar(cambios, mensaje){
    const t=await obtenerToken(); if(!t) throw new Error("Sin token");
    const ri=repoInfo();
    const url=`https://api.github.com/repos/${ri.owner}/${ri.repo}/contents/${NGC_ARCHIVO}`;
    const H={Authorization:"Bearer "+t,Accept:"application/vnd.github+json"};
    let ultimo=null;
    for(let intento=0;intento<4;intento++){
      let sha=null, doc=docVacio();
      const g=await fetch(url+"?ref="+ri.rama+"&_="+Date.now(),{headers:H,cache:"no-store"});
      if(g.ok){ const j=await g.json(); sha=j.sha; try{ doc=JSON.parse(utf8dec(b64dec(j.content))); }catch(e){ doc=docVacio(); } }
      else if(g.status===401||g.status===403) throw new Error("Token sin permiso ("+g.status+")");
      else if(g.status!==404) throw new Error("GitHub "+g.status);
      doc.version=1; doc.correcciones=doc.correcciones||{};
      for(const c in cambios){ if(cambios[c]===null) delete doc.correcciones[c]; else doc.correcciones[c]=cambios[c]; }
      doc.actualizado=ahora(); doc.por=lsGet(K.user,"")||"";
      const body={message:mensaje, branch:ri.rama,
        content:b64enc(utf8enc(JSON.stringify(doc,null,1)+"\n"))};
      if(sha) body.sha=sha;
      const p=await fetch(url,{method:"PUT",headers:Object.assign({"Content-Type":"application/json"},H),body:JSON.stringify(body)});
      if(p.ok) return doc;
      ultimo=p.status;
      if(p.status===409||p.status===422){ await new Promise(r=>setTimeout(r,600*(intento+1))); continue; }
      if(p.status===401||p.status===403||p.status===404) throw new Error("Token sin permiso de escritura ("+p.status+")");
      throw new Error("GitHub "+p.status);
    }
    throw new Error("Conflicto al guardar ("+ultimo+"), intenta de nuevo");
  },
  async verificar(token){
    const ri=repoInfo();
    const r=await fetch(`https://api.github.com/repos/${ri.owner}/${ri.repo}`,{headers:{Authorization:"Bearer "+token,Accept:"application/vnd.github+json"}});
    if(!r.ok) throw new Error(r.status===401?"Token inválido":"No se pudo acceder al repo ("+r.status+")");
    const j=await r.json();
    if(!(j.permissions&&j.permissions.push)) throw new Error("El token no tiene permiso de escritura (Contents: Read and write)");
    return ri.owner+"/"+ri.repo;
  }
};
// Para una futura base de datos: NGC_STORES.db = { leer(){…}, cambiar(cambios,msg){…}, listo(){…} }
const NGC_STORES = { github: GH };
window.NGC_STORES = NGC_STORES;
const store=()=>NGC_STORES[NGC_STORE_ACTIVO];

/* ---------- refresco de pantalla ---------- */
function refrescar(){
  bump();
  try{ if(typeof _orNGIniciado!=="undefined") _orNGIniciado=false; }catch(e){}
  try{ if($("#orTable")){ if(typeof _orNGRender==="function") _orNGRender(); if(typeof pintarOR==="function") pintarOR(); } }catch(e){ console.warn("ngcorr refrescar",e); }
  try{ if(typeof _repintarBandaNG==="function" && $("#orTable")) _repintarBandaNG(); }catch(e){}
  pintarPanelSiAbierto();
}

/* ---------- sincronización ---------- */
async function adoptarRemoto(doc){
  if(!doc||typeof doc!=="object") return false;
  if((doc.actualizado||"")>=(cache.actualizado||"")){
    const antes=JSON.stringify(cache.correcciones||{});
    cache={version:1,actualizado:doc.actualizado||"",por:doc.por||"",correcciones:doc.correcciones||{}};
    lsSet(K.cache,cache);
    return antes!==JSON.stringify(cache.correcciones);
  }
  return false;
}
async function cargarRemoto(silencioso){
  try{
    const doc=await store().leer();
    const cambio=await adoptarRemoto(doc);
    if(cambio) refrescar();
    return true;
  }catch(e){ if(!silencioso) aviso("No se pudo leer correcciones: "+e.message,"err"); return false; }
}
async function sincronizar(){
  const cats=Object.keys(cola);
  if(!cats.length){ await cargarRemoto(false); pintarPanelSiAbierto(); return; }
  if(!store().listo()){ aviso("Conecta el token para subir los cambios pendientes","err"); return; }
  estado.sync=true; pintarPanelSiAbierto();
  const envio={}; cats.forEach(c=>envio[c]=cola[c]);
  const por=lsGet(K.user,"")||"alguien";
  const msg = cats.length===1
    ? "NG: "+cats[0]+" → "+(envio[cats[0]]===null?"(revertido)":(envio[cats[0]].ng||"SIN SUSTITUTO"))+" (por "+por+")"
    : "NG: "+cats.length+" correcciones (por "+por+")";
  try{
    const doc=await store().cambiar(envio,msg);
    cache={version:1,actualizado:doc.actualizado,por:doc.por,correcciones:doc.correcciones};
    lsSet(K.cache,cache);
    cats.forEach(c=>{ if(JSON.stringify(cola[c])===JSON.stringify(envio[c])) delete cola[c]; });
    lsSet(K.cola,cola);
    bump();
    aviso("Guardado en el repo ✔","ok");
  }catch(e){
    aviso("Quedó pendiente (sin subir): "+e.message,"err");
  }
  estado.sync=false;
  pintarPanelSiAbierto();
}

/* ---------- API de cambio ---------- */
function ngOriginal(cat){ return ngDeOrig(cat)||""; }
async function aplicarCambio(cat, nuevoNg){
  nuevoNg=norm(nuevoNg);
  const orig=ngOriginal(cat);
  if(nuevoNg && nuevoNg.toUpperCase()!==orig.toUpperCase()){
    // respetar la ortografía de un nombre ya existente
    const ex=ngListaNueva().find(n=>n.toUpperCase()===nuevoNg.toUpperCase());
    nuevoNg = ex || nuevoNg.toUpperCase();
  }
  if(nuevoNg===orig){ cola[cat]=null; }   // igual al Excel => revertir
  else cola[cat]={ ng:nuevoNg, por:lsGet(K.user,"")||"", fecha:ahora(), anterior:orig };
  // si la corrección no existía en el repo y se revierte, no hace falta subir nada
  if(cola[cat]===null && !(cache.correcciones||{})[cat]) delete cola[cat];
  lsSet(K.cola,cola);
  refrescar();
  await sincronizar();
}

/* ---------- UI: estilos y utilidades ---------- */
function estilos(){
  if(document.getElementById("ngcStyle")) return;
  const s=document.createElement("style"); s.id="ngcStyle";
  s.textContent=`
  .ngc-lapiz{border:none;background:none;cursor:pointer;color:#1a3a6a;font-size:13px;padding:0 4px;margin-left:4px;opacity:.55}
  .ngc-lapiz:hover{opacity:1}
  .ngc-corr{color:#b45309;font-size:11px;margin-left:2px;cursor:help}
  .ngc-ov{position:fixed;inset:0;background:rgba(15,27,45,.5);display:flex;align-items:center;justify-content:center;z-index:400;padding:16px}
  .ngc-box{background:#fff;border-radius:14px;max-width:560px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 18px 50px rgba(0,0,0,.3);font-size:13.5px}
  .ngc-box h3{margin:0;padding:14px 18px;border-bottom:1px solid #dde3ee;font-size:15px}
  .ngc-b{padding:14px 18px;display:grid;gap:10px}
  .ngc-f{display:flex;gap:10px;justify-content:flex-end;padding:12px 18px;border-top:1px solid #dde3ee;background:#f7f9fc;flex-wrap:wrap}
  .ngc-b input[type=text],.ngc-b input[type=password]{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #c5cfe0;border-radius:8px;font:inherit}
  .ngc-b label{font-weight:600;font-size:12px;color:#44506a;display:block;margin-bottom:3px}
  .ngc-mut{color:#6b7790;font-size:12px}
  .ngc-ok{color:#166534}.ngc-err{color:#b91c1c}
  .ngc-tb{width:100%;border-collapse:collapse;font-size:12.5px}
  .ngc-tb th,.ngc-tb td{border-bottom:1px solid #e6ebf3;padding:5px 6px;text-align:left;vertical-align:top}
  .ngc-toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#12263f;color:#fff;padding:10px 16px;border-radius:10px;z-index:500;font-size:13px;box-shadow:0 6px 20px rgba(0,0,0,.3)}
  .ngc-toast.err{background:#9b1c1c}.ngc-toast.ok{background:#166534}
  `;
  document.head.appendChild(s);
}
let _toastT=null;
function aviso(txt,tipo){
  estado.msg=txt; estado.tipo=tipo||"";
  const old=document.getElementById("ngcToast"); if(old) old.remove();
  const d=document.createElement("div"); d.id="ngcToast"; d.className="ngc-toast "+(tipo||""); d.textContent=txt;
  document.body.appendChild(d); clearTimeout(_toastT); _toastT=setTimeout(()=>d.remove(),4500);
}
function overlay(html, id){
  estilos();
  const prev=document.getElementById(id); if(prev) prev.remove();
  const ov=document.createElement("div"); ov.className="ngc-ov"; ov.id=id;
  ov.innerHTML='<div class="ngc-box">'+html+'</div>';
  ov.addEventListener("mousedown",e=>{ if(e.target===ov && id!=="ngcPin") ov.remove(); });
  document.body.appendChild(ov);
  return ov;
}
function pedirTexto(titulo, ayuda, esPass){
  return new Promise(res=>{
    const ov=overlay(`<h3>${esc(titulo)}</h3><div class="ngc-b"><div class="ngc-mut">${esc(ayuda||"")}</div>
      <input type="${esPass?"password":"text"}" id="ngcPinIn" autocomplete="off"></div>
      <div class="ngc-f"><button class="btn btn-outline btn-sm" id="ngcPinNo">Cancelar</button><button class="btn btn-sm" id="ngcPinSi">Aceptar</button></div>`,"ngcPin");
    const inp=ov.querySelector("#ngcPinIn"); inp.focus();
    const fin=v=>{ ov.remove(); res(v); };
    ov.querySelector("#ngcPinNo").onclick=()=>fin(null);
    ov.querySelector("#ngcPinSi").onclick=()=>fin(inp.value);
    inp.onkeydown=e=>{ if(e.key==="Enter") fin(inp.value); if(e.key==="Escape") fin(null); };
  });
}
function pedirUsuarioSiFalta(){
  return new Promise(async res=>{
    if(lsGet(K.user,"")) return res(true);
    const n=await pedirTexto("¿Quién hace la corrección?","Tu nombre (queda registrado en cada cambio).",false);
    if(!n||!n.trim()) return res(false);
    lsSet(K.user,n.trim()); res(true);
  });
}

/* ---------- UI: edición de un catálogo ---------- */
function cuentaGrupo(ng){
  let n=0; const ef=corrEfectivas();
  for(const c of Object.keys(DB.materiales||{})){ if((ngDeNuevo(c)||"")===ng) n++; }
  return n;
}
async function ngcEditar(cat){
  if(!(await pedirUsuarioSiFalta())) return;
  const info=(typeof mat==="function"?mat(cat):{})||{};
  const orig=ngOriginal(cat), act=ngDeNuevo(cat), ef=corrEfectivas()[cat];
  const lista=ngListaNueva();
  const ov=overlay(`<h3>Corregir nombre genérico</h3>
   <div class="ngc-b">
    <div><b>${esc(cat)}</b> — ${esc(info.desc||info.descripcion||"")}</div>
    <div class="ngc-mut">En el Excel: <b>${esc(orig||"SIN SUSTITUTO")}</b>${ef?` · Ahora (corregido): <b>${esc(act||"SIN SUSTITUTO")}</b>`:""}</div>
    <div><label>Nuevo nombre genérico</label>
      <input type="text" id="ngcNuevo" list="ngcDL" value="${esc(act)}" placeholder="Escribe o elige · vacío = SIN SUSTITUTO" autocomplete="off">
      <datalist id="ngcDL">${lista.map(n=>`<option value="${esc(n)}">`).join("")}<option value="SIN SUSTITUTO"></datalist></div>
    <div class="ngc-mut" id="ngcInfo"></div>
    <div class="ngc-mut">Se aplica a todos los usuarios al guardarse en el repo (la página puede tardar unos minutos en reflejarlo en otros equipos).</div>
   </div>
   <div class="ngc-f">
     ${ef?'<button class="btn btn-outline btn-sm" id="ngcRev" style="margin-right:auto">↺ Volver al del Excel</button>':""}
     <button class="btn btn-outline btn-sm" id="ngcNo">Cancelar</button>
     <button class="btn btn-sm" id="ngcSi">Guardar</button>
   </div>`,"ngcEdit");
  const inp=ov.querySelector("#ngcNuevo"), inf=ov.querySelector("#ngcInfo");
  const pinta=()=>{
    const v=norm(inp.value);
    const ex=lista.find(n=>n.toUpperCase()===v.toUpperCase());
    inf.textContent = !v ? "Quedará sin grupo (SIN SUSTITUTO)."
      : ex ? `Se unirá al grupo «${ex}» (${cuentaGrupo(ex)} catálogos ahora).`
      : `Se creará el grupo nuevo «${v.toUpperCase()}».`;
  };
  inp.oninput=pinta; pinta(); inp.focus(); inp.select();
  const guardar=async()=>{ const v=inp.value; ov.remove(); await aplicarCambio(cat,v); };
  ov.querySelector("#ngcSi").onclick=guardar;
  inp.onkeydown=e=>{ if(e.key==="Enter") guardar(); if(e.key==="Escape") ov.remove(); };
  ov.querySelector("#ngcNo").onclick=()=>ov.remove();
  const rv=ov.querySelector("#ngcRev"); if(rv) rv.onclick=async()=>{ ov.remove(); await aplicarCambio(cat,orig); };
}

/* ---------- UI: panel de correcciones ---------- */
function fFecha(s){ try{ return new Date(s).toLocaleString("es-MX",{dateStyle:"short",timeStyle:"short"}); }catch(e){ return s||""; } }
function pintarPanelSiAbierto(){ if(document.getElementById("ngcPanel")) ngcPanel(true); }
function ngcPanel(sinFocus){
  const ef=corrEfectivas(), cats=Object.keys(ef).sort();
  const pend=Object.keys(cola).length;
  const conectado=hayTokenGuardado();
  const ri=repoInfo();
  const filas=cats.map(c=>{
    const e=ef[c], pe=(c in cola);
    return `<tr><td><b>${esc(c)}</b><div class="ngc-mut">${esc((typeof mat==="function"?mat(c).desc:"")||"")}</div></td>
      <td>${esc(e.anterior||"SIN SUSTITUTO")} → <b>${esc(e.ng||"SIN SUSTITUTO")}</b>${pe?' <span class="ngc-err">(pendiente)</span>':""}</td>
      <td class="ngc-mut">${esc(e.por||"")}<br>${esc(fFecha(e.fecha))}</td>
      <td><button class="btn btn-outline btn-sm" onclick="ngcRevertir('${esc(c)}')">↺</button></td></tr>`;
  }).join("");
  const prev=document.getElementById("ngcPanel");
  const scroll=prev?prev.querySelector(".ngc-box").scrollTop:0;
  const ov=overlay(`<h3>🏷 Correcciones de nombre genérico</h3>
   <div class="ngc-b">
     <div class="ngc-mut">Repo: <b>${esc(ri.owner)}/${esc(ri.repo)}</b> · Última sincronización: ${esc(cache.actualizado?fFecha(cache.actualizado)+(cache.por?" por "+cache.por:""):"—")}</div>
     <div>${conectado?'<span class="ngc-ok">● Token conectado</span>'+(tokenCifrado()?" (protegido con PIN)":""):'<span class="ngc-err">● Sin token: los cambios quedan pendientes en este equipo</span>'}
       ${pend?` · <b class="ngc-err">${pend} pendiente(s) de subir</b>`:""}
       ${lsGet(K.user,"")?` · Usuario: <b>${esc(lsGet(K.user,""))}</b>`:""}</div>
     ${cats.length?`<table class="ngc-tb"><thead><tr><th>Catálogo</th><th>Cambio</th><th>Quién / cuándo</th><th></th></tr></thead><tbody>${filas}</tbody></table>`
       :'<div class="ngc-mut">Aún no hay correcciones. Usa el ✎ junto al nombre genérico en la tabla.</div>'}
   </div>
   <div class="ngc-f">
     <button class="btn btn-outline btn-sm" onclick="ngcConectar()">${conectado?"Cambiar token":"Conectar token"}</button>
     ${conectado?'<button class="btn btn-outline btn-sm" onclick="ngcDesconectar()">Desconectar</button>':""}
     <button class="btn btn-outline btn-sm" onclick="ngcNombre()">Mi nombre</button>
     <button class="btn btn-sm" ${estado.sync?"disabled":""} onclick="ngcSync()">${estado.sync?"Sincronizando…":(pend?"Subir pendientes":"Actualizar")}</button>
     <button class="btn btn-outline btn-sm" onclick="document.getElementById('ngcPanel').remove()">Cerrar</button>
   </div>`,"ngcPanel");
  ov.querySelector(".ngc-box").scrollTop=scroll;
}
async function ngcRevertir(cat){ if(!(await pedirUsuarioSiFalta())) return; await aplicarCambio(cat, ngOriginal(cat)); }
async function ngcSync(){ await sincronizar(); }
async function ngcNombre(){
  const n=await pedirTexto("Tu nombre","Queda registrado en cada corrección.",false);
  if(n&&n.trim()){ lsSet(K.user,n.trim()); pintarPanelSiAbierto(); }
}
function ngcDesconectar(){
  if(!confirm("¿Quitar el token de este equipo?")) return;
  lsDel(K.tok); tokenSesion=null; pintarPanelSiAbierto();
}
function ngcConectar(){
  const ri=repoInfo();
  const ov=overlay(`<h3>Conectar con GitHub</h3>
   <div class="ngc-b">
     <div class="ngc-mut">1) En GitHub: Settings → Developer settings → Personal access tokens → <b>Fine-grained tokens</b> → Generate.<br>
       2) Repository access: <b>Only select repositories</b> → <b>${esc(ri.owner)}/${esc(ri.repo)}</b>.<br>
       3) Permissions → Repository → <b>Contents: Read and write</b>. Ponle vencimiento.<br>
       4) Pega el token aquí. Se guarda solo en este navegador.</div>
     <div><label>Token</label><input type="password" id="ngcTk" placeholder="github_pat_…" autocomplete="off"></div>
     <div><label>PIN (opcional, recomendado en equipos compartidos)</label><input type="password" id="ngcPn" placeholder="Cifra el token; se pide una vez por sesión" autocomplete="off"></div>
     <div class="ngc-mut" id="ngcTkMsg"></div>
   </div>
   <div class="ngc-f"><button class="btn btn-outline btn-sm" id="ngcTkNo">Cancelar</button><button class="btn btn-sm" id="ngcTkSi">Verificar y guardar</button></div>`,"ngcTok");
  ov.querySelector("#ngcTkNo").onclick=()=>ov.remove();
  ov.querySelector("#ngcTkSi").onclick=async()=>{
    const t=ov.querySelector("#ngcTk").value.trim(), pin=ov.querySelector("#ngcPn").value;
    const m=ov.querySelector("#ngcTkMsg");
    if(!t){ m.textContent="Pega el token."; return; }
    m.textContent="Verificando…";
    try{
      const r=await GH.verificar(t);
      await guardarToken(t,pin);
      ov.remove(); aviso("Conectado a "+r,"ok");
      if(!lsGet(K.user,"")) await pedirUsuarioSiFalta();
      pintarPanelSiAbierto();
      if(Object.keys(cola).length) sincronizar();
    }catch(e){ m.innerHTML='<span class="ngc-err">'+esc(e.message)+"</span>"; }
  };
}

/* ---------- ganchos para or.js ---------- */
function ngcLapiz(cat){
  const e=corrEfectivas()[cat];
  const c=esc(cat).replace(/'/g,"");
  return (e?`<span class="ngc-corr" title="Corregido${e.por?" por "+esc(e.por):""} (Excel: ${esc(e.anterior||"SIN SUSTITUTO")})">●</span>`:"")
    +`<button class="ngc-lapiz" title="Corregir nombre genérico" onclick="ngcEditar('${c}')">✎</button>`;
}
Object.assign(window,{ngcEditar,ngcPanel,ngcRevertir,ngcSync,ngcNombre,ngcConectar,ngcDesconectar,ngcLapiz});

/* ---------- arranque ---------- */
function iniciar(){ bump(); cargarRemoto(true); if(Object.keys(cola).length && hayTokenGuardado() && !tokenCifrado()) { tokenSesion=(lsGet(K.tok,{})||{}).t||null; sincronizar(); } }
if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",iniciar); else iniciar();
// los datos (DB) llegan después: reintenta refresco cuando existan
let _n=0; const _t=setInterval(()=>{ if(typeof DB!=="undefined"&&DB&&DB.nombresGenericos){ clearInterval(_t); refrescar(); } else if(++_n>60) clearInterval(_t); },500);
})();
