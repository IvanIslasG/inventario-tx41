// ═══════════════════════════════════════════════════════════════════════════
// MÓDULO: CONFIGURACIÓN / ADMIN — parámetros, generación de paquete, consumos
// Depende de helpers globales de index.html: $, DB, SOBRE, mat, nfmt, b64, _K,
// funciones de criptografía (cifrar/descifrar del bloque GATE)
// ═══════════════════════════════════════════════════════════════════════════
/* ============ CONFIGURACIÓN / ADMIN ============ */
let _adminTotal=null, _adminConsumos=null;
function modConfig(){
  const m=DB.meta;
  const consActuales=DB.consumos?Object.keys(DB.consumos):[];
  $("#moduleView").innerHTML=`
  <div class="cfg">
    <div class="panel"><div class="panel-head"><h2>Datos del paquete actual</h2></div>
      <div style="padding:16px"><div class="cfg">
        <div class="row"><b>Paquete generado</b><span>${new Date(m.generado).toLocaleString('es-MX')}</span></div>
        <div class="row"><b>Almacén distribuidor</b><span>${m.almacen_distribuidor} · PUEBLA</span></div>
        <div class="row"><b>Materiales en maestro</b><span class="num">${nfmt(m.n_materiales)}</span></div>
        <div class="row"><b>Almacenes en directorio</b><span class="num">${m.n_almacenes}</span></div>
        <div class="row"><b>Materiales con existencia</b><span class="num">${nfmt(m.n_con_existencia)}</span></div>
        <div class="row"><b>Consumos cargados</b><span>${consActuales.length?consActuales.join(", "):"ninguno"}</span></div>
      </div></div>
    </div>

    <div class="panel"><div class="panel-head"><h2>🔄 Actualizar existencias</h2><span class="pill">Admin</span></div>
      <div style="padding:16px;display:grid;gap:14px">
        <p style="margin:0;color:var(--muted);font-size:13px">Sube el export <b>TOTAL</b> de S/4HANA (.xlsx). Se recalculan existencias, traslados y lotes; el directorio y el maestro se conservan.</p>
        <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
          <label class="btn" style="cursor:pointer">📁 Elegir TOTAL.xlsx<input type="file" id="upTotal" accept=".xlsx" hidden></label>
          <span id="upTotalName" style="font-size:13px;color:var(--muted)">Ningún archivo</span>
          <button class="btn-prim" id="upProcesar" disabled>Procesar</button>
        </div>
        <div id="upResumen" style="font-size:13px"></div>
      </div>
    </div>

    <div class="panel"><div class="panel-head"><h2>📈 Archivos de consumo</h2><span class="pill">Para OR</span></div>
      <div style="padding:16px;display:grid;gap:12px">
        <p style="margin:0;color:var(--muted);font-size:13px">Sube los consumos por almacén auxiliar (ej. <b>TX8A</b>). El nombre del archivo identifica el almacén (TX8A → A08A). Columna B = catálogo, columna C = consumo. Para el Distribuidor Puebla, nombra el archivo con "D041" — soporta el formato con columnas por mes y "Sumatoria".</p>
        <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center">
          <label class="btn" style="cursor:pointer">📁 Agregar consumos<input type="file" id="upCons" multiple hidden></label>
          <span id="upConsName" style="font-size:13px;color:var(--muted)">Ninguno</span>
        </div>
        <div id="upConsList" style="font-size:13px;display:grid;gap:6px"></div>
      </div>
    </div>

    <div class="panel"><div class="panel-head"><h2>💾 Guardar y publicar</h2></div>
      <div style="padding:16px;display:grid;gap:12px">
        <div id="genResumen" style="font-size:13px;color:var(--muted)">Procesa un TOTAL o agrega consumos para generar un paquete nuevo. También puedes re-publicar el actual.</div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <button class="btn" id="genDescargar">⬇ Descargar datos.enc</button>
          <button class="btn" id="genUsar">↻ Usar ahora (esta sesión)</button>
        </div>
        <div style="margin-top:16px;padding-top:16px;border-top:1px solid var(--line)">
          <div style="font-weight:600;color:var(--primary);margin-bottom:10px">🚀 Publicar en GitHub Pages</div>
          <div style="display:grid;gap:8px;max-width:420px">
            <input id="ghToken" type="password" placeholder="GitHub token (fine-grained, Contents R/W)">
            <button class="btn-prim" id="ghPublicar">🚀 Publicar en GitHub Pages</button>
            <div id="ghStatus" style="font-size:12.5px;color:var(--muted)"></div>
            <p style="font-size:11.5px;color:var(--muted);margin:0">Token <b>fine-grained</b> limitado a este repo · Contents: Read and write.</p>
          </div>
        </div>
      </div>
    </div>

    <div class="panel"><div style="padding:16px"><button class="btn" id="cfgOut">Recargar datos</button></div></div>
  </div>`;

  // Actualizar existencias
  let totalFile=null;
  $("#upTotal").onchange=e=>{ totalFile=e.target.files[0]||null;
    $("#upTotalName").textContent=totalFile?totalFile.name:"Ningún archivo";
    $("#upProcesar").disabled=!totalFile; };
  $("#upProcesar").onclick=async()=>{
    if(!totalFile) return;
    $("#upResumen").innerHTML=`<span style="color:var(--muted)">Procesando ${totalFile.name}…</span>`;
    try{
      const buf=await totalFile.arrayBuffer();
      const parsed=parseTotalJS(buf);
      _adminTotal=construirBundleJS(parsed);
      const bm=_adminTotal.meta;
      $("#upResumen").innerHTML=`<div style="background:var(--ok-bg);color:var(--ok);padding:10px 12px;border-radius:8px">
        ✅ Procesado: <b>${nfmt(bm.n_con_existencia)}</b> con existencia · <b>${parsed.nAlmacenes}</b> almacenes ·
        <b>${nfmt(parsed.nTraslados)}</b> con traslado · <b>${nfmt(parsed.nLotesD041)}</b> con lotes en D041.</div>`;
      actualizarGenResumen();
    }catch(err){ console.error(err);
      $("#upResumen").innerHTML=`<div style="background:#fde8e8;color:#c0392b;padding:10px 12px;border-radius:8px">Error: ${err.message}</div>`; }
  };

  // Consumos
  $("#upCons").onchange=async e=>{
    _adminConsumos=_adminConsumos||{};
    for(const f of e.target.files){
      try{
        const {alm,mapa,n}=await parseConsumoFile(f);
        _adminConsumos[alm]=mapa;
        renderConsLine(alm,f.name,n,Object.values(mapa).reduce((s,v)=>s+v,0));
      }catch(err){ renderConsLine("?",f.name+" — error: "+err.message,0,0,true); }
    }
    $("#upConsName").textContent=Object.keys(_adminConsumos).length+" almacén(es)";
    actualizarGenResumen();
  };

  // Generar / publicar
  $("#genDescargar").onclick=async()=>{
    const btn=$("#genDescargar");
    btn.disabled=true; btn.textContent="⏳ Cifrando…";
    try{
      const clave=`${_K}`;
      if(!clave){ alert("No se encontró la clave de sesión. Cierra sesión y vuelve a entrar."); return; }
      const sobre=await cifrarJS(bundleFinal(), clave);
      const json=JSON.stringify(sobre);
      const blob=new Blob([json],{type:"application/octet-stream"});
      const url=URL.createObjectURL(blob);
      const a=document.createElement("a"); a.href=url; a.download="datos.enc";
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(()=>URL.revokeObjectURL(url), 5000);
      btn.textContent="✅ Descargado";
    }catch(err){
      console.error("Error al cifrar/descargar:",err);
      alert("Error al generar el archivo: "+err.message);
      btn.textContent="⬇ Descargar datos.enc";
    } finally {
      setTimeout(()=>{ btn.disabled=false; btn.textContent="⬇ Descargar datos.enc"; }, 3000);
    }
  };
  $("#genUsar").onclick=()=>{
    DB=bundleFinal(); SOBRE=null; _adminTotal=null; _adminConsumos=null;
    // Actualizar el meta en el header
    const m=DB.meta;
    $("#meta").innerHTML=`Generado ${new Date(m.generado).toLocaleString('es-MX')}<br>${m.n_materiales.toLocaleString('es-MX')} materiales · ${m.n_almacenes} almacenes`;
    mostrarMenu();
    setTimeout(()=>alert("✅ Datos actualizados en esta sesión.\nPara que todos lo vean, descarga y publica el datos.enc."),100);
  };
  $("#ghPublicar").onclick=()=> publicarGitHub();
  $("#cfgOut").onclick=()=>{ location.reload(); };
}
function renderConsLine(alm,info,n,suma,err){
  const div=document.createElement("div");
  div.style.cssText="display:flex;gap:8px;align-items:center;padding:6px 10px;border-radius:8px;background:"+(err?"#fde8e8":"var(--ok-bg)");
  div.innerHTML=err?`<span style="color:#c0392b">⚠ ${info}</span>`
    :`<span style="color:var(--ok)">✅ <b>${alm}</b> · ${almName(alm)} — ${nfmt(n)} catálogos · consumo ${nfmt(suma)}</span>`;
  $("#upConsList").appendChild(div);
}
function actualizarGenResumen(){
  const partes=[];
  if(_adminTotal) partes.push("existencias nuevas");
  if(_adminConsumos&&Object.keys(_adminConsumos).length) partes.push(Object.keys(_adminConsumos).length+" consumo(s)");
  $("#genResumen").innerHTML = partes.length
    ? `Listo para generar: <b>${partes.join(" + ")}</b>. Descarga, usa o publica.`
    : "Procesa un TOTAL o agrega consumos para generar un paquete nuevo. También puedes re-publicar el actual.";
}
// Combina: TOTAL nuevo (o DB actual) + consumos (nuevos sobre los existentes)
function fechaLocalISO(){
  const d=new Date();
  const pad=n=>String(n).padStart(2,"0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function bundleFinal(){
  const b = _adminTotal ? _adminTotal : JSON.parse(JSON.stringify(DB));
  const cons = Object.assign({}, b.consumos||{}, _adminConsumos||{});
  if(Object.keys(cons).length) b.consumos = cons;
  b.meta = Object.assign({}, b.meta, { generado:fechaLocalISO(),
    n_materiales:Object.keys(b.materiales).length, n_almacenes:Object.keys(b.directorio.almacenes).length,
    n_con_existencia:Object.keys(b.existencias).length });
  return b;
}

/* ---- Parser de consumos (TSV de SAP o xlsx) ---- */
function claveDeNombre(nombre){
  const base=nombre.replace(/\.[^.]+$/,"").trim().toUpperCase();
  if(base.includes("D041")) return "D041"; // archivo propio del Distribuidor Puebla, formato distinto
  return /^TX/.test(base) ? "A0"+base.slice(2) : base;
}
async function parseConsumoFile(file){
  const alm=claveDeNombre(file.name);
  const buf=await file.arrayBuffer(); const bytes=new Uint8Array(buf);
  let mapa={};
  if(bytes[0]===0x50 && bytes[1]===0x4B){ // xlsx (zip)
    const wb=XLSX.read(buf,{type:"array"}); const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
    // Formato con encabezado propio (ej. D041): columnas "Catálogo" ... "Sumatoria"
    const hdr=(rows[0]||[]).map(h=>_txtJS(h).toLowerCase().replace(/á/g,"a"));
    const iCat=hdr.findIndex(h=>h.includes("catalogo")||h.includes("material"));
    const iSum=hdr.findIndex(h=>h.includes("sumatoria")||h.includes("suma")||h.includes("total"));
    if(iCat>=0 && iSum>=0){
      // Se guarda la suma tal cual viene en el archivo (sin normalizar a ningún número de meses fijo).
      // El divisor real se aplica después, en Análisis, con el campo "Meses CPM" que el usuario controla.
      const vistos=new Set();
      for(let i=1;i<rows.length;i++){
        const r=rows[i]; if(!r) continue;
        const cat=_txtJS(r[iCat]);
        if(!/^\d+$/.test(cat)) continue;
        if(vistos.has(cat)) break; // catálogo repetido = inicio de un bloque de ajustes/errores al final del archivo; se descarta todo lo que sigue
        vistos.add(cat);
        mapa[cat]=_numJS(r[iSum]);
      }
    } else {
      // Formato simple sin encabezado: columna B = catálogo, columna C = consumo
      rows.forEach(r=>{ const cat=_txtJS(r[1]); if(/^\d+$/.test(cat) && !(cat in mapa)) mapa[cat]=_numJS(r[2]); });
    }
  } else { // texto TSV
    const txt=new TextDecoder("iso-8859-1").decode(buf);
    txt.split(/\r?\n/).forEach(ln=>{ const p=ln.split("\t"); if(p.length<3) return;
      const cat=_txtJS(p[1]); if(/^\d+$/.test(cat)) mapa[cat]=_numJS(p[2]); });
  }
  if(!Object.keys(mapa).length) throw new Error("sin filas válidas (revisa columnas B y C, o encabezados Catálogo/Sumatoria)");
  return {alm, mapa, n:Object.keys(mapa).length};
}

/* ---- Parser del TOTAL en el navegador (replica parser_total.py) ---- */
function _txtJS(v){ return v===null||v===undefined?"":String(v).trim(); }
function _numJS(v){ if(v===null||v===undefined) return 0; const s=String(v).trim().replace(/,/g,""); if(!s) return 0; const f=parseFloat(s); return isNaN(f)?0:f; }
function parseTotalJS(arrayBuffer){
  const wb=XLSX.read(arrayBuffer,{type:"array"});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""});
  let hr=-1;
  for(let i=0;i<Math.min(rows.length,10);i++){ if(rows[i].some(c=>_txtJS(c).toLowerCase()==="material")){ hr=i; break; } }
  if(hr<0) throw new Error("No se encontró la fila de encabezados (columna 'Material').");
  const H=rows[hr]; const col={};
  H.forEach((h,i)=>{ h=_txtJS(h).toLowerCase();
    if(h==="ce."||h.startsWith("centro")) col.centro=i;
    else if(h==="alm."||h.startsWith("almac")) col.almacen=i;
    else if(h==="material") col.material=i;
    else if(h.includes("texto breve")) col.desc=i;
    else if(h==="umb") col.umb=i;
    else if(h==="lote") col.lote=i;
    else if(h.startsWith("librutiliz")||h.includes("libre")) col.libre=i;
    else if(h.startsWith("transytras")||h.includes("traslad")||h.includes("transy")) col.tras=i;
  });
  ["centro","almacen","material","libre"].forEach(k=>{ if(!(k in col)) throw new Error("Falta columna: "+k); });
  const D="D041"; const materiales={}; const almacenes={};
  for(let i=hr+1;i<rows.length;i++){
    const r=rows[i]; if(!r) continue;
    const ce=_txtJS(r[col.centro]), alm=_txtJS(r[col.almacen]), cat=_txtJS(r[col.material]);
    if(!(ce||alm||cat)) continue; if(!(cat&&alm)) continue;
    const libre=_numJS(r[col.libre]); const tras="tras"in col?_numJS(r[col.tras]):0;
    const lote="lote"in col?_txtJS(r[col.lote]):""; const desc="desc"in col?_txtJS(r[col.desc]):""; const umb="umb"in col?_txtJS(r[col.umb]):"";
    almacenes[alm]=ce;
    let m=materiales[cat]; if(!m){ m=materiales[cat]={desc:"",umb:"",almacenes:{},traslados:{},lotes:{}}; }
    if(desc&&!m.desc) m.desc=desc; if(umb&&!m.umb) m.umb=umb;
    const loteReal=lote&&lote.toUpperCase()!=="NUEVO";
    if(loteReal){ (m.lotes[alm]||(m.lotes[alm]=[])).push({lote,lib:libre,tras}); }
    m.almacenes[alm]=(m.almacenes[alm]||0)+libre;
    if(tras) m.traslados[alm]=(m.traslados[alm]||0)+tras;
  }
  return {materiales,almacenes,D,
    nAlmacenes:Object.keys(almacenes).length,
    nTraslados:Object.values(materiales).filter(m=>m.traslados[D]).length,
    nLotesD041:Object.values(materiales).filter(m=>m.lotes[D]).length};
}
function construirBundleJS(parsed){
  const D=parsed.D;
  const materiales={}; for(const [c,v] of Object.entries(DB.materiales)) materiales[c]={...v};
  for(const [cat,m] of Object.entries(parsed.materiales)){
    if(!materiales[cat]) materiales[cat]={um:m.umb,desc:m.desc,area:"",ubic:""};
    else { if(!materiales[cat].desc) materiales[cat].desc=m.desc; if(!materiales[cat].um) materiales[cat].um=m.umb; }
  }
  const existencias={}, traslados={}, lotes={}, valoracion=[];
  for(const [cat,m] of Object.entries(parsed.materiales)){
    const ex={}; for(const[a,v]of Object.entries(m.almacenes)) if(v>0) ex[a]=v;
    if(Object.keys(ex).length) existencias[cat]=ex;
    if(Object.keys(m.traslados).length) traslados[cat]=m.traslados;
    if(m.lotes[D]&&m.lotes[D].length) lotes[cat]=m.lotes[D];
    if(Object.keys(m.lotes).length) valoracion.push(cat);
  }
  valoracion.sort();
  const out={ meta:{ generado:fechaLocalISO(), almacen_distribuidor:D,
      n_materiales:Object.keys(materiales).length, n_almacenes:Object.keys(DB.directorio.almacenes).length,
      n_con_existencia:Object.keys(existencias).length,
      meses_cpm:DB.meta?.meses_cpm||6, meses_stock:DB.meta?.meses_stock||1.5 },
    directorio:DB.directorio, materiales, existencias, traslados, lotes, valoracion };
  if(DB.consumos) out.consumos=DB.consumos;
  if(DB.nombresGenericos) out.nombresGenericos=DB.nombresGenericos;
  return out;
}
/* ---- Cifrado en navegador (replica publicar.py) ---- */
function _b64enc(buf){
  // fromCharCode con spread falla en buffers grandes; procesar en chunks de 8KB
  const u8=new Uint8Array(buf), chunk=8192;
  let s="";
  for(let i=0;i<u8.length;i+=chunk){
    s+=String.fromCharCode.apply(null, u8.subarray(i, i+chunk));
  }
  return btoa(s);
}
async function _gzip(bytes){
  const cs=new CompressionStream("gzip");
  const r=new Response(new Blob([bytes]).stream().pipeThrough(cs));
  return new Uint8Array(await r.arrayBuffer());
}
async function cifrarJS(bundle, clave){
  const plano=new TextEncoder().encode(JSON.stringify(bundle));
  const comp=await _gzip(plano);
  const salt=crypto.getRandomValues(new Uint8Array(16));
  const iv=crypto.getRandomValues(new Uint8Array(12));
  const iter=200000;
  const baseKey=await crypto.subtle.importKey("raw", new TextEncoder().encode(clave), "PBKDF2", false, ["deriveKey"]);
  const key=await crypto.subtle.deriveKey({name:"PBKDF2",salt,iterations:iter,hash:"SHA-256"}, baseKey,
    {name:"AES-GCM",length:256}, false, ["encrypt"]);
  const ct=await crypto.subtle.encrypt({name:"AES-GCM",iv}, key, comp);
  return { v:1, alg:"AES-GCM", kdf:"PBKDF2", hash:"SHA-256", iter, gzip:true,
    salt:_b64enc(salt), iv:_b64enc(iv), ct:_b64enc(ct) };
}
async function publicarGitHub(){
  const st=$("#ghStatus"), btn=$("#ghPublicar");
  const token=$("#ghToken").value.trim();
  if(!token){ st.style.color="#c0392b"; st.textContent="Ingresa el token de GitHub."; return; }
  btn.disabled=true; btn.textContent="⏳ Publicando…";
  st.style.color="var(--muted)"; st.textContent="Cifrando y subiendo…";
  const repo="IvanIslasG/inventario-tx41", path="datos.enc";
  const api=`https://api.github.com/repos/${repo}/contents/${path}`;
  const headers={"Authorization":"Bearer "+token,"Accept":"application/vnd.github+json","Content-Type":"application/json"};
  try{
    const sobre=await cifrarJS(bundleFinal(), _K);
    let sha=null;
    const g=await fetch(api,{headers}); if(g.ok){ sha=(await g.json()).sha; }
    const body={message:"Actualización "+fechaLocalISO(), content:btoa(JSON.stringify(sobre))};
    if(sha) body.sha=sha;
    const p=await fetch(api,{method:"PUT",headers,body:JSON.stringify(body)});
    if(p.ok){ st.style.color="var(--ok)"; st.textContent="✅ Publicado. Visible en ~1 min en GitHub Pages."; }
    else{ const e=await p.json(); st.style.color="#c0392b"; st.textContent="Error "+p.status+": "+(e.message||""); }
  }catch(err){ st.style.color="#c0392b"; st.textContent="Error: "+err.message; }
  finally{ btn.disabled=false; btn.textContent="🚀 Publicar en GitHub Pages"; }
}


