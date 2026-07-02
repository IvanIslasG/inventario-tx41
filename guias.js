/* ============================================================
   guias.js  -  Módulo Guías de Embarque · Sistema TX41 Puebla
   Depende de: DB, mat(), almName(), $(), XLSX (SheetJS)
   ============================================================ */

// ── BD de empaques (seed inicial — crece orgánicamente en localStorage) ──────
const _GUIAS_BD_SEED = [{"cat": "1000320", "um": "PZ", "tipo": "Caja", "cont": 192, "freq": 3, "patio": false}, {"cat": "1000372", "um": "PZ", "tipo": "Caja", "cont": 15, "freq": 1, "patio": false}, {"cat": "1000416", "um": "PZ", "tipo": "Caja", "cont": 150, "freq": 1, "patio": false}, {"cat": "1000430", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1000430", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1000381", "um": "PZ", "tipo": "Costal", "cont": 50, "freq": 1, "patio": false}, {"cat": "1000418", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1000600", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1000713", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1000715", "um": "PZ", "tipo": "Caja", "cont": 3000, "freq": 1, "patio": false}, {"cat": "1000902", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1000903", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1001224", "um": "M", "tipo": "Caja", "cont": 250, "freq": 1, "patio": false}, {"cat": "1026180", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": true}, {"cat": "1028665", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1028954", "um": "PZ", "tipo": "Caja", "cont": 20000, "freq": 1, "patio": false}, {"cat": "1034256", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1034325", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1035461", "um": "PZ", "tipo": "Caja", "cont": 12, "freq": 1, "patio": false}, {"cat": "1036183", "um": "PZ", "tipo": "Caja", "cont": 140, "freq": 1, "patio": false}, {"cat": "1036183", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1036713", "um": "PZ", "tipo": "Caja", "cont": 6000, "freq": 1, "patio": false}, {"cat": "1038272", "um": "PZ", "tipo": "Caja", "cont": 150, "freq": 1, "patio": false}, {"cat": "1038272", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1045853", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1046965", "um": "PZ", "tipo": "Caja", "cont": 30, "freq": 1, "patio": false}, {"cat": "1052980", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1054083", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1054594", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 2, "patio": false}, {"cat": "1054594", "um": "PZ", "tipo": "Caja", "cont": 1000, "freq": 1, "patio": false}, {"cat": "1000291", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1000295", "um": "PZ", "tipo": "Caja", "cont": 24, "freq": 1, "patio": false}, {"cat": "1000296", "um": "PZ", "tipo": "Caja", "cont": 24, "freq": 1, "patio": false}, {"cat": "1000337", "um": "PZ", "tipo": "Caja", "cont": 10000, "freq": 1, "patio": false}, {"cat": "1000371", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1000674", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1002447", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": true}, {"cat": "1002449", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002470", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002502", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002447", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002535", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Pieza", "cont": 120, "freq": 3, "patio": false}, {"cat": "1002433", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002521", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002522", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002554", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002556", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1034253", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1034254", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Caja", "cont": 12, "freq": 2, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Caja", "cont": 50, "freq": 1, "patio": false}, {"cat": "1000508", "um": "M", "tipo": "Tubo", "cont": 6, "freq": 1, "patio": true}, {"cat": "1002353", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002413", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002551", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002734", "um": "PZ", "tipo": "Caja", "cont": 10, "freq": 1, "patio": false}, {"cat": "1010457", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1025372", "um": "PZ", "tipo": "Caja", "cont": 40, "freq": 1, "patio": false}, {"cat": "1026092", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1038152", "um": "PZ", "tipo": "Caja", "cont": 10, "freq": 1, "patio": false}];
const _GUIAS_ALM_SEED = {"D041": {"nombre": "ALMACEN DISTRIBUIDOR PUEBLA", "atiende": "GERMAN PEREZ PORRAS", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "A08A": {"nombre": "ALMACEN AUXILIAR SAN PEDRO", "atiende": "ASSENETH ALAVARADO HERRERA", "domicilio": "25 NORTE No. 3617  COL NUEVA AURORA", "ciudad": "PUEBLA, PUE.", "tel": "01 222 2463193 , 01 222 2424783 FAX", "cp": "72070"}, "A08B": {"nombre": "ALMACEN AUXILIAR ZARAGOZA", "atiende": "ROBERTO COYOTZI MONTES", "domicilio": "CALZ IGNACIO ZARAGOZA No 247 COL TEPEYAC", "ciudad": "PUEBLA, PUE.", "tel": "12222352031", "cp": ""}, "A08C": {"nombre": "ALMACEN AUXILIAR ORIZABA", "atiende": "VANIA MARTINEZ JUAREZ", "domicilio": "31 ORIENTE No 308 CIRCUNVALACION", "ciudad": "ORIZABA, VERACRUZ", "tel": "", "cp": ""}, "A08D": {"nombre": "ALMACEN AUXILIAR TLAXCALA", "atiende": "CARLOS VERGARA", "domicilio": "AV INSTITUTO POLITECNICO NACIONAL S/N", "ciudad": "TLAXCALA, TLAXCALA", "tel": "", "cp": ""}, "A08E": {"nombre": "ALMACEN AUXILIAR TEHUACAN", "atiende": "OMAR EURESTI MARTINEZ", "domicilio": "16 NORTE No 413 COL SERDAN", "ciudad": "TEHUACAN, PUEBLA", "tel": "12383821393", "cp": "75750"}, "A08F": {"nombre": "ALMACEN AUXILIAR SAN BRUNO", "atiende": "GONZALO SILCA POZOS", "domicilio": "MARTIREZ 28 DE AGOSTO", "ciudad": "XALAPA, VERACRUZ", "tel": "12288153631", "cp": "91020"}, "A08G": {"nombre": "ALMACEN AUXILIAR CRISTAL", "atiende": "LAURA MANZO SILVESTRE, OMAR BAIZABAL", "domicilio": "ANTONIO CHEDRAHUI CARAM No 250", "ciudad": "XALAPA, VERACRUZ", "tel": "12288150444", "cp": "91180"}, "A08H": {"nombre": "ALMACEN AUXILIAR MOCAMBO", "atiende": "HIRAM SALAMANCA", "domicilio": "AV PALMERAS No 305 COL JARDINES DE VIRGINIA", "ciudad": "BOCA DEL RIO, VERACRUZ", "tel": "12299214122", "cp": ""}, "A08I": {"nombre": "ALMACEN AUXILIAR DOS CAMINOS", "atiende": "DIEGO GARCÍA", "domicilio": "AV 11 No 2627", "ciudad": "CORDOBA, VERACRUZ", "tel": "12717165433", "cp": ""}, "A08J": {"nombre": "ALMACEN AUXILIAR AQUILES SERDAN", "atiende": "JOSÉ ANTONIO LÓPEZ C", "domicilio": "35 PONIENTE No 723", "ciudad": "PUEBLA, PUE.", "tel": "12222433928", "cp": ""}, "A08K": {"nombre": "ALMACEN AUXILIAR ALTAMIRANO", "atiende": "LILIA CONTRERAS", "domicilio": "ALTAMIRANO No 1226", "ciudad": "VERACRUZ, VERACRUZ", "tel": "12299204842", "cp": "91700"}, "A08L": {"nombre": "ALMACEN AUXILIAR LERDO", "atiende": "JULIO CÉSAR GONZÁLEZ", "domicilio": "MARIANO ARISTA No 4424", "ciudad": "VERACRUZ, VERACRUZ", "tel": "12299204842", "cp": "91726"}, "A08M": {"nombre": "ALMACEN AUXILIAR PEÑUELA", "atiende": "JUAN LUNA", "domicilio": "KM 343 BOULEVARD S/N", "ciudad": "CORDOBA, VERACRUZ", "tel": "12717166707", "cp": "94501"}, "A08T": {"nombre": "ALMACEN AUXILIAR TULANCINGO", "atiende": "PEDRO MORALES LIRA", "domicilio": "RIVA PALACIOS 203 COL LOS ALAMOS", "ciudad": "TULANCINGO, HIDALGO", "tel": "17757536701", "cp": "43640"}, "A08V": {"nombre": "ALMACEN AUXILIAR PACHUCA", "atiende": "HIRAM ROSALES", "domicilio": "SAN MARTIN DE PORRES No 407 CFE", "ciudad": "PACHUCA DE SOTO, HIDALGO", "tel": "17717143677", "cp": "42090"}, "A08X": {"nombre": "ALMACEN AUXILIAR ATLIXCO", "atiende": "SERGIO NOE MILIAN PINTLE", "domicilio": "CALZ. OAXACA No 2710", "ciudad": "ATLIXCO, PUEBLA", "tel": "12444451868", "cp": "74294"}, "A08Y": {"nombre": "ALMACEN AUXILIAR MAYORAZGO", "atiende": "GLORIA MARCELA SILVA T", "domicilio": "CALLE BENITO JUAREZ No 12311", "ciudad": "PUEBLA, PUE.", "tel": "12222192800", "cp": ""}, "A08Z": {"nombre": "ALMACEN AUXILIAR POZA RICA", "atiende": "ENRIQUE LEANDRO BLANCO", "domicilio": "POZO 2 No 48 COL DIVISION NORTE", "ciudad": "POZA RICA, VERACRUZ", "tel": "17828233004", "cp": "93350"}, "A09C": {"nombre": "ALMACEN AUXILIAR COATZACOALCOS", "atiende": "JUANA MARIA RAMOS", "domicilio": "AV. JACINTO LEMARROY S/N FRACC. RANCHO ALEGRE 2", "ciudad": "COATZACOALCOS, VERACRUZ", "tel": "19212186550", "cp": "96558"}, "A09S": {"nombre": "ALMACEN AUXILIAR MINATITLAN", "atiende": "SERGIO VAZQUEZ", "domicilio": "EMILIANO ZAPATA No. 25 COL. INSURGENTES SUR", "ciudad": "MINATITLAN, VERACRUZ", "tel": "19222211979", "cp": "96710"}, "A09T": {"nombre": "ALMACEN AUXILIAR CHOLULA", "atiende": "ILIANA DE ITA, CARLOS GUTIERREZ", "domicilio": "22 ORIENTE No 606", "ciudad": "CHOLULA, PUEBLA", "tel": "12222474746", "cp": ""}, "A09U": {"nombre": "ALMACEN AUXILIAR AMALUCAN", "atiende": "JOSEFINA ONOFRE M", "domicilio": "BLVRD ATEOPAN ESQ 17A", "ciudad": "PUEBLA, PUE.", "tel": "12222868207", "cp": ""}, "A09V": {"nombre": "ALMACEN AUXILIAR TULA", "atiende": "PILAR HERNANDEZ, VERONICA MEZA", "domicilio": "AV SUR ESQUINA OTE No 25 COL CENTRO", "ciudad": "TULA DE ALLENDE, HIDALGO", "tel": "17737322888", "cp": "42800"}, "A0X4": {"nombre": "ALMACEN TELECO", "atiende": "", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "A0Y4": {"nombre": "ALMACEN CYCSA PUEBLA", "atiende": "ERNESTO CENTENO", "domicilio": "CALLE C No. 7 PARQUE INDUSTRIAL 2000", "ciudad": "PUEBLA, PUE.", "tel": "", "cp": ""}, "A082": {"nombre": "ALMACEN AUXILIAR TUXPAN", "atiende": "JAVIER CRUZ YAÑEZ", "domicilio": "DEMETRIO RUIZ MALERVA S/N COL CENTRO", "ciudad": "TUXPAN VERACRUZ", "tel": "17838340234", "cp": ""}, "A081": {"nombre": "ALMACEN AUXILIAR PACHOACAN", "atiende": "FELIPE OLANDES MORA", "domicilio": "AV FRNACISCO SARABIA No 100 COL CARLOS ROVIROSA", "ciudad": "PACHUCA DE SOTO, HIDALGO", "tel": "1771148651", "cp": "42082"}, "A083": {"nombre": "ALMACEN AUXILIAR SAN MARTIN", "atiende": "AGUSTIN BALBUENA SCHIAFINI", "domicilio": "AV. JUVENTUD S/N COL. LOS DICIOS", "ciudad": "SAN MARTIN TEXMELUCAN, PUEBLA", "tel": "", "cp": ""}, "D032": {"nombre": "ALMACEN DISTRIBUIDOR CELAYA", "atiende": "JOSE TORADO", "domicilio": "AV LAS FUENTES 10904 COL LAS FUENTES", "ciudad": "CELAYA, GTO", "tel": "", "cp": ""}, "D022": {"nombre": "ALMACEN DISTRIBUIDOR HERMOSILLO", "atiende": "BRUNO GENDA FERNANDEZ", "domicilio": "BLVRD JESUS GARCIA MORALES KM 5 No 145 COL EL LLANO", "ciudad": "HERMOSILLO, SONORA", "tel": "16622188701", "cp": "83210"}, "A09P": {"nombre": "ALMACEN AUXILIAR PINOTEPA", "atiende": "JUAN PEDRO GONZALEZ MARTINEZ", "domicilio": "AV AGUIRRE PALANCARES Y BAÑOS AGUIRRE", "ciudad": "PINOTEPA NACIONAL", "tel": "19545433906", "cp": "71600"}, "A09O": {"nombre": "ALMACEN AUXILIAR OAXACA 1", "atiende": "CARLOS TEJEDA", "domicilio": "HEROICO COLEGIO MILITAR No 1013 COL REFORMA", "ciudad": "OAXACA, OAX", "tel": "19515127100", "cp": "68050"}, "A091": {"nombre": "ALMACEN AUXILIAR APIZACO", "atiende": "LOURDES ESPINOSA", "domicilio": "VENUSTIANO CARRANZA ESQ JOSE A", "ciudad": "APIZACO, TLAXCALA", "tel": "12414174555", "cp": "90350"}, "A084": {"nombre": "ALMACEN AUXILIAR TEZIUTLAN", "atiende": "RAMON GALINDO BECERRA", "domicilio": "AV ENCINO Y AVELLANO SN", "ciudad": "TEZIUTLAN, PUEBLA", "tel": "12313130715", "cp": "73890"}, "A09Q": {"nombre": "ALMACEN AUXILIAR OAXACA II", "atiende": "LEON RUIS MATADAMAS", "domicilio": "ESMERALDA No 201 COL BUGAMBILIAS", "ciudad": "OAXACA, OAX", "tel": "19515127100", "cp": "68010"}, "A09R": {"nombre": "ALMACEN AUXILIAR HUAJUAPAN", "atiende": "J CARLOS GARCIA, VEDA CAROLINA OSORIO", "domicilio": "PROLONGACION DE MINA No 120", "ciudad": "H. CIUDAD DE HUAJUAPAN DE LEON, OAX", "tel": "19535324388", "cp": "69007"}, "D011": {"nombre": "ALMACEN DISTRIBUIDOR MONTERREY", "atiende": "ANTONIO AGUILAR", "domicilio": "CORDILLERA DE LOS ANDES No 701 JARDIN DE LAS PTES S/N0", "ciudad": "SN NICOLAS DE LOS GARZA, NUEVO LEON", "tel": "18183505572", "cp": "66460"}, "A0LE": {"nombre": "ALMACEN AUXILIAR LEGARIA", "atiende": "HERNESTO HERNANDEZ", "domicilio": "FELIPE CARRILLO PUERTO No. 750 TORRE BLANCA  MIGUEL HIDALGO", "ciudad": "MEXICO, DF", "tel": "", "cp": ""}, "D043": {"nombre": "ALMACEN DISTRIBUIDOR VILLAHERMOSA", "atiende": "EDUARDO TABOADA", "domicilio": "AV. ACERO S/N ESQUINA COBRE CD.INDUSTRIAL", "ciudad": "VILLAHERMOSA. TABASCO", "tel": "", "cp": ""}, "A09X": {"nombre": "ALMACEN AUXILIAR CHETUMAL", "atiende": "ARTURO ALONSO RAMIREZ", "domicilio": "AVENIDA 4 DE MARZO NO.30 COL.FIDEL VELAZQUEZ", "ciudad": "CHETUMAL QUINTANA ROO", "tel": "19838372365", "cp": "77080"}, "D021": {"nombre": "ALMACEN DISTRIBUIDOR GUADALAJARA", "atiende": "ENRIQUE MEDINA LOPEZ", "domicilio": "TRATADO DE TLALTELOCO No. 4114 COL. PARQUE AUDITORIO", "ciudad": "ZAPOPAN, JALISCO", "tel": "3336601554", "cp": ""}, "D013": {"nombre": "ALMACEN DISTRIBUIDOR CHIHUAHUA", "atiende": "CARLOS DURAN", "domicilio": "MIGUEL BARRAGAN No. 6903 COL. EL PARRAL", "ciudad": "CHIHUAHUA, CHIHUAHUA", "tel": "", "cp": ""}, "D007": {"nombre": "ALMACEN GENERAL LA PERLA", "atiende": "ROGER CANO", "domicilio": "CALLE NUEVA ESQUINA NEGRA MODELO COL. INDUSTRIAL", "ciudad": "NAUCALPAN EDO DE MEXICO", "tel": "", "cp": ""}, "MTZ": {"nombre": "ALMACEN AUXILIAR MARTINEZ DE LA TORRE", "atiende": "ROLANDO RAZGADO DE JESUS", "domicilio": "BOULEVARD ALFINO FLORES S/N COL. ADOLFO RUIZ CORTINA", "ciudad": "MARTINEZ DE LA TORRE, VER.", "tel": "", "cp": ""}, "C0TA": {"nombre": "CARSO TLAXCALA", "atiende": "", "domicilio": "CALLE AV.OCOTLAN S/N CONTRA ESQ. CALLE CONSTRUCTORES", "ciudad": "SANTA ANA CHIAUTEMPAN", "tel": ",012464620646", "cp": ""}, "ACTOPAN": {"nombre": "NUEVO ACTOPAN", "atiende": "PAREDES TENORIO FRANCISCA", "domicilio": "PEDRO MORENO No 49", "ciudad": "NUEVO ACTOPAN, HGO", "tel": "17727273187", "cp": "42500"}, "A09I": {"nombre": "MERIDA BUENAVISTA", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "MERIDA", "tel": "-", "cp": "-"}, "SAHAGUN": {"nombre": "CT SAHAGUN", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "CD SAHAGÚN", "tel": "-", "cp": "-"}, "TECA": {"nombre": "CT TECAMACHALCO", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "TECAMACHALCO", "tel": "-", "cp": "-"}, "A0AS": {"nombre": "ALMACEN ABASTOS", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "&", "tel": "&", "cp": "&"}, "A0CD": {"nombre": "ALMACEN AUXILIAR  DIANA", "atiende": "LUIS FELIPE ESTROP", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "SLPI": {"nombre": "ALMACEN AUXILIAR PINO", "atiende": "", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}};
const _GUIAS_LS_BD    = "guias_bd_empaques_v1";
const _GUIAS_LS_HIST  = "guias_historial_v1";

// ── Cargar/guardar BD de empaques ────────────────────────────────────────────
function _guiasBDCargar(){
  try{
    const raw = localStorage.getItem(_GUIAS_LS_BD);
    if(raw) return JSON.parse(raw);
  }catch(e){}
  // Primera vez — usar seed
  localStorage.setItem(_GUIAS_LS_BD, JSON.stringify(_GUIAS_BD_SEED));
  return JSON.parse(JSON.stringify(_GUIAS_BD_SEED));
}

function _guiasBDGuardar(bd){
  try{ localStorage.setItem(_GUIAS_LS_BD, JSON.stringify(bd)); }catch(e){}
}

function _guiasBDActualizarEmpaque(cat, tipo, cont, um){
  var bd = _guiasBDCargar();
  // Buscar si ya existe este catálogo+contenido
  var idx = bd.findIndex(function(e){ return e.cat === cat && e.cont === cont; });
  if(idx >= 0){
    bd[idx].freq = (bd[idx].freq || 1) + 1;
    bd[idx].tipo = tipo;
  } else {
    bd.push({cat: cat, um: um || '', tipo: tipo, cont: cont, freq: 1, patio: false});
  }
  _guiasBDGuardar(bd);
}

function _guiasBDOpciones(cat){
  var bd = _guiasBDCargar();
  return bd.filter(function(e){ return e.cat === cat; })
           .sort(function(a,b){ return (b.freq||0)-(a.freq||0); });
}

// ── Historial de guías ────────────────────────────────────────────────────────
function _guiasHistCargar(){
  try{ return JSON.parse(localStorage.getItem(_GUIAS_LS_HIST)||"[]"); }catch(e){ return []; }
}
function _guiasHistGuardar(hist){ try{ localStorage.setItem(_GUIAS_LS_HIST, JSON.stringify(hist)); }catch(e){} }
function _guiasHistAgregar(guia){
  var hist = _guiasHistCargar();
  // Reemplazar si ya existe el mismo folio+area
  var idx = hist.findIndex(function(g){ return g.folio===guia.folio && g.area===guia.area; });
  if(idx >= 0) hist[idx] = guia;
  else hist.unshift(guia);
  if(hist.length > 50) hist = hist.slice(0,50);
  _guiasHistGuardar(hist);
}

// ── Info de almacén (BD propia + directorio DB) ───────────────────────────────
function _guiasAlmInfo(sigla){
  // Preferir BD_ALMACENES propia (tiene más datos: domicilio, atiende, tel)
  if(_GUIAS_ALM_SEED[sigla]) return _GUIAS_ALM_SEED[sigla];
  // Fallback al directorio de DB
  var d = DB.directorio?.almacenes?.[sigla];
  if(d) return {nombre: d.desc||sigla, atiende:'', domicilio:'', ciudad:'', cp:'', tel:''};
  return {nombre: sigla, atiende:'', domicilio:'', ciudad:'', cp:'', tel:''};
}

// ── Estado de la guía en progreso ─────────────────────────────────────────────
var _guiaActual = null; // { destino, area, folio, fecha, lineas:[], transporte:'' }

// ── PANTALLA 1: Menú principal ────────────────────────────────────────────────
function modGuias(){
  var hist = _guiasHistCargar();

  // Agrupar historial por área
  var areas = ["Herramientas","Misceláneos","Papelería","Cables","Ropa y Calzado","General"];
  var porArea = {};
  areas.forEach(function(a){ porArea[a] = []; });
  hist.forEach(function(g,i){ var a = g.area||"General"; if(!porArea[a]) porArea[a]=[]; porArea[a].push({g:g,i:i}); });

  var histHtml = "";
  if(hist.length === 0){
    histHtml = "<div style=\"color:var(--muted);font-size:13px;padding:16px 0\">No hay guías generadas aún.</div>";
  } else {
    areas.forEach(function(area){
      var items = porArea[area] || [];
      if(!items.length) return;
      histHtml +=
        "<details open style=\"margin-bottom:12px\">" +
        "<summary style=\"font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
        "letter-spacing:.4px;cursor:pointer;padding:6px 0;list-style:none\">" +
        area + " <span style=\"color:var(--primary)\">(" + items.length + ")</span></summary>" +
        "<div style=\"display:flex;flex-direction:column;gap:6px;margin-top:8px\">";
      items.forEach(function(item){
        var g=item.g, i=item.i;
        histHtml +=
          "<div style=\"display:flex;align-items:center;gap:10px;padding:10px 14px;" +
          "background:white;border:1px solid var(--line);border-radius:10px\">" +
          "<div style=\"flex:1;cursor:pointer\" onclick=\"_guiasAbrirHistorial(" + i + ")\">" +
          "<div style=\"font-size:13px;font-weight:700\">No. " + g.folio + " &mdash; " + (g.destino||'') + "</div>" +
          "<div style=\"font-size:11px;color:var(--muted)\">" + (g.fecha||'') + " &middot; " + (g.lineas||0) + " materiales</div>" +
          "</div>" +
          "<button onclick=\"_guiasBorrarHistorial(" + i + ")\" title=\"Borrar\"" +
          " style=\"background:none;border:none;color:#dc2626;cursor:pointer;font-size:16px;padding:0\">&times;</button>" +
          "</div>";
      });
      histHtml += "</div></details>";
    });
  }

  $("#moduleView").innerHTML =
    "<div style=\"max-width:600px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"margin-bottom:24px\">" +
    "<h2 style=\"margin:0 0 4px;font-size:20px\">Guías de Embarque</h2>" +
    "<p style=\"margin:0;color:var(--muted);font-size:13px\">D041 &middot; Almacén Distribuidor Puebla</p>" +
    "</div>" +
    "<button onclick=\"_guiasNueva()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;" +
    "margin-bottom:20px\">+ Nueva guía de embarque</button>" +
    "<div style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:10px\">Guías recientes</div>" +
    "<div style=\"display:flex;flex-direction:column;gap:8px\">" + histHtml + "</div>" +
    "</div>";
}

// ── PANTALLA 2: Nueva guía — datos de cabecera ────────────────────────────────
function _guiasNueva(){
  // Calcular próximo folio por área (el usuario puede cambiarlo)
  var hist = _guiasHistCargar();

  // Obtener áreas disponibles
  var areas = ["Herramientas","Misceláneos","Papelería","Cables","Ropa y Calzado","General"];

  var areasHtml = areas.map(function(a){
    return "<option value=\"" + a + "\">" + a + "</option>";
  }).join("");

  // Almacenes para el select de destino (del directorio DB)
  var almsDb = Object.keys(DB.directorio?.almacenes||{}).sort();
  var almsHtml = almsDb.map(function(k){
    var info = _guiasAlmInfo(k);
    return "<option value=\"" + k + "\">" + k + " &mdash; " + info.nombre + "</option>";
  }).join("");

  $("#moduleView").innerHTML =
    "<div style=\"max-width:560px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:24px\">" +
    "<button onclick=\"modGuias()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Cancelar</button>" +
    "<h2 style=\"margin:0;font-size:18px\">Nueva guía</h2>" +
    "</div>" +

    // Área
    "<div style=\"margin-bottom:16px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Área</label>" +
    "<select id=\"gArea\" style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:14px;font-family:inherit\">" + areasHtml + "</select>" +
    "</div>" +

    // Folio
    "<div style=\"margin-bottom:16px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">No. de Guía (folio)</label>" +
    "<input id=\"gFolio\" type=\"number\" min=\"1\" placeholder=\"Ej. 47\"" +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:16px;font-weight:700;font-family:inherit;color:var(--primary)\">" +
    "</div>" +

    // Destino
    "<div style=\"margin-bottom:16px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Destino</label>" +
    "<input id=\"gDestino\" list=\"gDestinoList\" placeholder=\"Escribe o selecciona almacén...\"" +
    " oninput=\"_guiasActualizarDestinatario()\"" +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:14px;font-family:inherit\">" +
    "<datalist id=\"gDestinoList\">" + almsHtml + "</datalist>" +
    "</div>" +

    // Info destinatario (se llena automático)
    "<div id=\"gDestinatarioInfo\" style=\"background:var(--lite,#f0f4ff);border-radius:10px;" +
    "padding:12px 16px;margin-bottom:16px;font-size:12px;color:var(--muted);display:none\">" +
    "</div>" +

    // Fecha
    "<div style=\"margin-bottom:16px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Fecha</label>" +
    "<input id=\"gFecha\" type=\"date\"" +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:14px;font-family:inherit\">" +
    "</div>" +

    // Transporte
    "<div style=\"margin-bottom:24px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Línea de transporte (opcional)</label>" +
    "<input id=\"gTransporte\" type=\"text\" placeholder=\"Ej. ESTAFETA, DHL, Transporte propio\"" +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:14px;font-family:inherit\">" +
    "</div>" +

    // Botón continuar
    "<button onclick=\"_guiasContinuarMateriales()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit\">" +
    "Continuar &rarr; Capturar materiales</button>" +
    "</div>";

  // Poner fecha de hoy por default
  var hoy = new Date();
  var yyyy = hoy.getFullYear();
  var mm = String(hoy.getMonth()+1).padStart(2,'0');
  var dd = String(hoy.getDate()).padStart(2,'0');
  document.getElementById("gFecha").value = yyyy+"-"+mm+"-"+dd;
}

function _guiasActualizarDestinatario(){
  var raw = document.getElementById("gDestino").value.trim();
  // Aceptar sigla directa o "SIGLA — NOMBRE"
  var sigla = raw.split(" ")[0].split("—")[0].trim().toUpperCase();
  var info = document.getElementById("gDestinatarioInfo");
  if(!sigla){ info.style.display="none"; return; }
  var d = _guiasAlmInfo(sigla);
  info.innerHTML =
    "<b>" + d.nombre + "</b><br>" +
    (d.atiende ? "Atiende: " + d.atiende + "<br>" : "") +
    (d.domicilio ? d.domicilio + "<br>" : "") +
    (d.ciudad ? d.ciudad + (d.cp?" &nbsp; C.P. "+d.cp:"") : "") +
    (d.tel ? "<br>Tel. " + d.tel : "");
  info.style.display = "block";
}

function _guiasContinuarMateriales(){
  var area    = document.getElementById("gArea").value;
  var folio   = document.getElementById("gFolio").value.trim();
  var raw = document.getElementById("gDestino").value.trim();
  var destino = raw.split(" ")[0].split("—")[0].trim().toUpperCase();
  var fecha   = document.getElementById("gFecha").value;
  var transp  = document.getElementById("gTransporte").value.trim();

  if(!folio){ alert("Ingresa el número de guía."); return; }
  if(!destino){ alert("Selecciona el almacén destino."); return; }

  var almInfo = _guiasAlmInfo(destino);
  _guiaActual = {
    area: area, folio: parseInt(folio), destino: destino,
    almInfo: almInfo, fecha: fecha, transporte: transp, lineas: [],
    generador: '', surtio: '', operador: '', placas: '', tipoVeh: ''
  };

  _guiasCapturaMateriales();
}

// ── PANTALLA 3: Captura de materiales ─────────────────────────────────────────
function _guiasCapturaMateriales(){
  var lineas = _guiaActual.lineas;

  var lineasHtml = "";
  for(var i=0; i<lineas.length; i++){
    var l = lineas[i];
    lineasHtml += _tplLineaGuia(l, i);
  }

  var footerBtns = lineas.length > 0 ?
    "<button onclick=\"_guiasRevision()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;" +
    "margin-top:16px\">Revisar y generar guía &rarr;</button>" : "";

  $("#moduleView").innerHTML =
    "<div style=\"max-width:700px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:16px\">" +
    "<button onclick=\"_guiasNueva()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Datos</button>" +
    "<div><div style=\"font-size:17px;font-weight:800;color:var(--primary)\">" +
    "Guía " + _guiaActual.area + " No. " + _guiaActual.folio + "</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">" + _guiaActual.destino +
    " &middot; " + _guiaActual.almInfo.nombre + "</div></div>" +
    "</div>" +

    // Panel de entrada de material
    "<div style=\"background:white;border:1.5px solid var(--line);border-radius:12px;" +
    "padding:16px;margin-bottom:16px\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:10px\">Agregar material</div>" +
    "<div style=\"display:flex;gap:8px;flex-wrap:wrap\">" +
    "<input id=\"gCatInput\" type=\"text\" placeholder=\"Catálogo SAP (escanea o escribe)\"" +
    " oninput=\"_guiasCatBuscar(this.value)\"" +
    " onkeydown=\"if(event.key==='Enter')_guiasCatConfirmar()\"" +
    " style=\"flex:1;min-width:160px;padding:10px 14px;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:14px;font-family:inherit;font-weight:700;color:var(--primary)\">" +
    "<button onclick=\"_guiasCatConfirmar()\"" +
    " style=\"padding:10px 18px;background:var(--primary);color:white;border:none;" +
    "border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit\">" +
    "Agregar</button>" +
    "</div>" +
    "<div id=\"gCatInfo\" style=\"margin-top:8px;font-size:12px;color:var(--muted)\"></div>" +
    "</div>" +

    // Opciones de origen (OR o manual)
    "<div style=\"display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap\">" +
    "<button onclick=\"_guiasCargarOR()\"" +
    " style=\"flex:1;padding:10px;background:white;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;" +
    "color:var(--primary)\">&#8679; Importar desde OR (Excel)</button>" +
    "<input type=\"file\" id=\"gORFile\" accept=\".xlsx\" style=\"display:none\"" +
    " onchange=\"_guiasProcesarOR(this)\">" +
    "</div>" +

    // Lista de líneas capturadas
    "<div id=\"gLineasLista\" style=\"display:flex;flex-direction:column;gap:8px\">" +
    lineasHtml +
    "</div>" +
    footerBtns +
    "</div>";
}

function _tplLineaGuia(l, idx){
  // Determinar si va a granel (cant < contEmp) o es caja
  var esGranel = (!l.patio) && (l.contEmp <= 1 || l.cant < l.contEmp);
  var esPatio  = l.patio;
  var bgColor  = esPatio ? "#fff8e1" : esGranel ? "#f0fdf4" : "white";
  var badgeHtml = esPatio
    ? "<span style=\"background:#f59e0b;color:white;font-size:9px;font-weight:700;" +
      "padding:1px 6px;border-radius:8px;margin-left:6px\">PATIO</span>"
    : esGranel
    ? "<span style=\"background:#16a34a;color:white;font-size:9px;font-weight:700;" +
      "padding:1px 6px;border-radius:8px;margin-left:6px\">GRANEL</span>"
    : "";
  return (
    "<div style=\"background:" + bgColor + ";border:1px solid var(--line);border-radius:10px;" +
    "padding:12px 14px;display:flex;align-items:flex-start;gap:10px\">" +
    "<div style=\"flex:1\">" +
    "<div style=\"font-size:12px;font-weight:700;font-family:monospace;color:var(--primary)\">" + l.cat + badgeHtml + "</div>" +
    "<div style=\"font-size:12px;color:var(--text);margin:2px 0\">" + l.desc + "</div>" +
    "<div style=\"font-size:11px;color:var(--muted)\">" +
    l.cant + " " + l.um + " &mdash; " + l.bultos + " " + l.tipoEmp +
    (l.contEmp > 1 ? " de " + l.contEmp + " " + l.um : "") +
    (l.patio ? " &mdash; <b>PATIO</b>" : "") +
    "</div>" +
    "</div>" +
    "<button onclick=\"_guiasEliminarLinea(" + idx + ")\"" +
    " style=\"background:none;border:none;color:#dc2626;cursor:pointer;font-size:18px;" +
    "padding:0;line-height:1\">&times;</button>" +
    "</div>"
  );
}

function _guiasCatBuscar(val){
  var info = document.getElementById("gCatInfo");
  if(!val){ info.textContent = ""; return; }
  var m = mat(val.trim());
  if(m.desc){
    info.innerHTML = "<b>" + m.desc + "</b> &middot; " + m.um;
  } else {
    info.innerHTML = "<span style=\"color:var(--muted)\">Catálogo no encontrado en maestro &mdash; puedes agregarlo manualmente</span>";
  }
}

function _guiasCatConfirmar(){
  var catInput = document.getElementById("gCatInput");
  var cat = (catInput.value || "").trim();
  if(!cat){ alert("Ingresa un catálogo."); return; }

  var m = mat(cat);
  var desc = m.desc || "";
  var um   = m.um   || "";

  // Si no está en el maestro, pedir descripción manual
  if(!desc){
    desc = prompt("Descripción del material " + cat + ":", "") || "";
    um   = prompt("Unidad de medida:", "PZ") || "PZ";
  }

  // Buscar opciones de empaque en BD
  var opciones = _guiasBDOpciones(cat);
  _guiasPedirEmpaque(cat, desc, um, opciones);
}

function _guiasPedirEmpaque(cat, desc, um, opciones){
  var modal = document.createElement("div");
  modal.className = "modal on";

  var catEsc  = cat.replace(/'/g,"&#39;");
  var descEsc = desc.replace(/'/g,"&#39;");
  var esPatio = (mat(cat).ubic || "").toLowerCase() === "patio";

  // Si es patio — preguntar solo cantidad y agregar directo
  if(esPatio){
    var cant = prompt("Cantidad de " + cat + " (va a PATIO):", "");
    if(!cant || isNaN(cant) || parseInt(cant) < 1) return;
    cant = parseInt(cant);
    _guiaActual.lineas.push({
      cat: cat, desc: desc, um: um,
      cant: cant, tipoEmp: "Patio", contEmp: 1,
      bultos: cant, patio: true, granel: false
    });
    var inp = document.getElementById("gCatInput");
    if(inp){ inp.value = ""; document.getElementById("gCatInfo").textContent = ""; }
    _guiasRefrescarLineas();
    return;
  }

  // Botones de empaques conocidos
  var btnsConocidos = "";
  for(var i=0; i<opciones.length; i++){
    var op = opciones[i];
    btnsConocidos +=
      "<button onclick=\"_guiasSeleccionarEmpaque('" + catEsc + "','" + descEsc + "','" + um + "','" +
      op.tipo + "'," + op.cont + ")\"" +
      " style=\"text-align:left;padding:10px 14px;background:#f0f4ff;" +
      "border:1.5px solid var(--primary);border-radius:8px;cursor:pointer;" +
      "font-family:inherit;font-size:13px;width:100%;margin-bottom:6px\">" +
      "<b>" + op.tipo + "</b> de " + op.cont + " " + um +
      " <span style=\"color:var(--muted);font-size:11px\">(x" + (op.freq||1) + ")</span>" +
      "</button>";
  }

  // Columna izquierda — conocidos + granel
  var izqHtml =
    "<div style=\"flex:1;min-width:0;padding:16px\">" +
    (opciones.length > 0 ?
      "<div style=\"font-size:11px;font-weight:700;color:var(--primary);text-transform:uppercase;" +
      "letter-spacing:.4px;margin-bottom:10px\">&#10003; Conocidos</div>" +
      btnsConocidos +
      "<div style=\"border-top:1px dashed var(--line);margin-top:12px;padding-top:12px\">" : "") +
    "<div style=\"font-size:11px;font-weight:700;color:#16a34a;text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:6px\">&#9744; A granel</div>" +
    "<p style=\"font-size:11px;color:var(--muted);margin:0 0 8px\">Va a la caja colectiva del área</p>" +
    "<div style=\"display:flex;gap:6px\">" +
    "<input id=\"gGranelCant\" type=\"number\" min=\"1\" placeholder=\"Cantidad\"" +
    " style=\"flex:1;padding:8px;border:1.5px solid #16a34a;border-radius:8px;font-family:inherit\"/>" +
    "<button onclick=\"_guiasAGranelModal('" + catEsc + "','" + descEsc + "','" + um + "')\"" +
    " style=\"padding:8px 14px;background:#16a34a;color:white;border:none;border-radius:8px;" +
    "font-weight:700;cursor:pointer;font-family:inherit\">Granel</button>" +
    "</div>" +
    (opciones.length > 0 ? "</div>" : "") +
    "</div>";

  // Columna derecha — nuevo empaque (desplegable)
  var derechaHtml =
    "<div style=\"flex:1;min-width:0;padding:16px;border-left:1px solid var(--line)\">" +
    "<details" + (opciones.length === 0 ? " open" : "") + ">" +
    "<summary style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;cursor:pointer;padding:4px 0;list-style:none\">" +
    (opciones.length > 0 ? "+ Nuevo empaque diferente" : "Captura el empaque") +
    "</summary>" +
    "<div style=\"margin-top:10px;display:flex;flex-direction:column;gap:8px\">" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Tipo de empaque</label>" +
    "<input id=\"gEmpTipo\" type=\"text\" value=\"Caja\" placeholder=\"Caja, Costal...\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;margin-top:3px\"></div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Contenido por empaque</label>" +
    "<input id=\"gEmpCont\" type=\"number\" min=\"1\" value=\"1\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;margin-top:3px\"></div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Cantidad total</label>" +
    "<input id=\"gEmpCant\" type=\"number\" min=\"1\" value=\"\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:15px;font-weight:700;color:var(--primary);margin-top:3px\"></div>" +
    "</div></details>" +
    "</div>";

  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:660px;padding:0\">" +
    "<div style=\"padding:16px 20px;border-bottom:1px solid var(--line)\">" +
    "<h3 style=\"margin:0 0 4px\">Empaque &mdash; " + cat + "</h3>" +
    "<div style=\"font-size:12px;color:var(--muted)\">" + desc + "</div>" +
    "</div>" +
    "<div style=\"display:flex;align-items:stretch\">" +
    izqHtml +
    derechaHtml +
    "</div>" +
    "<div style=\"padding:12px 20px;border-top:1px solid var(--line);display:flex;" +
    "justify-content:space-between;align-items:center\">" +
    "<button class=\"btn\" onclick=\"this.closest('.modal').remove()\">Cancelar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarEmpaque('" + catEsc + "','" + descEsc + "','" + um + "')\">Agregar empaque</button>" +
    "</div></div>";

  document.body.appendChild(modal);
  setTimeout(function(){
    var el = document.getElementById(opciones.length === 0 ? "gEmpCant" : "gGranelCant");
    if(el) el.focus();
  }, 100);
}

function _guiasSeleccionarEmpaque(cat, desc, um, tipo, cont){
  cont = parseInt(cont);
  var cant = prompt("Cantidad total de " + cat + ":", "");
  if(!cant || isNaN(cant) || parseInt(cant) < 1) return;
  cant = parseInt(cant);
  document.querySelector(".modal")?.remove();

  var nCajas  = Math.floor(cant / cont);
  var residuo = cant % cont;

  if(nCajas > 0 && residuo > 0){
    // División automática: cajas cerradas + granel
    _guiasAgregarLineaSilente(cat, desc, um, tipo, cont, nCajas * cont, false);
    _guiasAgregarLineaSilente(cat, desc, um, "Granel", 0, residuo, true);
    _limpiarCatInput();
    _guiasRefrescarLineas();
  } else {
    // Todo en cajas cerradas o todo a granel
    var esGranel = (nCajas === 0);
    _guiasAgregarLinea(cat, desc, um, tipo, cont, cant, esGranel);
  }
  _guiasBDActualizarEmpaque(cat, tipo, cont, um);
}

function _guiasAPatioModal(cat, desc, um){
  var cant = parseInt(document.getElementById("gPatioCant")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa la cantidad de patio."); return; }
  document.querySelector(".modal")?.remove();
  var bultos = cant;
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: "Patio", contEmp: 1,
    bultos: bultos, patio: true, granel: false
  });
  var inp = document.getElementById("gCatInput");
  if(inp){ inp.value = ""; document.getElementById("gCatInfo").textContent = ""; }
  _guiasRefrescarLineas();
}

function _guiasAGranelModal(cat, desc, um){
  var cant = parseInt(document.getElementById("gGranelCant")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa la cantidad."); return; }
  document.querySelector(".modal")?.remove();
  _guiasAgregarLinea(cat, desc, um, "Granel", 0, cant, true);
}

function _guiasAGranel(cat, desc, um){
  var cant = prompt("Cantidad de " + cat + " que va a granel:", "");
  if(!cant || isNaN(cant) || parseInt(cant) < 1) return;
  cant = parseInt(cant);
  document.querySelector(".modal")?.remove();
  // contEmp=0 indica granel explícito
  _guiasAgregarLinea(cat, desc, um, "Granel", 0, cant, true);
}

function _guiasConfirmarEmpaque(cat, desc, um){
  var tipo = (document.getElementById("gEmpTipo")?.value || "Caja").trim();
  var cont = parseInt(document.getElementById("gEmpCont")?.value || "1");
  var cant = parseInt(document.getElementById("gEmpCant")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa la cantidad total."); return; }
  document.querySelector(".modal")?.remove();

  var nCajas  = cont > 1 ? Math.floor(cant / cont) : 0;
  var residuo = cont > 1 ? cant % cont : 0;

  if(nCajas > 0 && residuo > 0){
    _guiasAgregarLineaSilente(cat, desc, um, tipo, cont, nCajas * cont, false);
    _guiasAgregarLineaSilente(cat, desc, um, "Granel", 0, residuo, true);
    _limpiarCatInput();
    _guiasRefrescarLineas();
  } else {
    var esGranel = (nCajas === 0 && cont > 1);
    _guiasAgregarLinea(cat, desc, um, tipo, cont, cant, esGranel);
  }
  _guiasBDActualizarEmpaque(cat, tipo, cont, um);
}

function _limpiarCatInput(){
  var inp = document.getElementById("gCatInput");
  if(inp){ inp.value = ""; document.getElementById("gCatInfo").textContent = ""; }
}

function _guiasAgregarLineaSilente(cat, desc, um, tipoEmp, contEmp, cant, esGranel){
  // Agrega sin limpiar input ni refrescar (para llamadas múltiples)
  var bultos = contEmp > 1 ? Math.floor(cant / contEmp) : cant;
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: tipoEmp, contEmp: contEmp,
    bultos: bultos, patio: false, granel: esGranel || false
  });
}

function _guiasAgregarLinea(cat, desc, um, tipoEmp, contEmp, cant, esGranel){
  var bultos = contEmp > 1 ? Math.ceil(cant / contEmp) : cant;
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: tipoEmp, contEmp: contEmp,
    bultos: bultos, patio: false, granel: esGranel || false
  });
  // Limpiar input y cerrar modal
  _limpiarCatInput();
  document.querySelector(".modal")?.remove();
  _guiasRefrescarLineas();
}

function _guiasEliminarLinea(idx){
  _guiaActual.lineas.splice(idx, 1);
  _guiasRefrescarLineas();
}

function _guiasRefrescarLineas(){
  var lista = document.getElementById("gLineasLista");
  if(!lista) return;
  var lineas = _guiaActual.lineas;
  var html = "";
  for(var i=0; i<lineas.length; i++) html += _tplLineaGuia(lineas[i], i);
  lista.innerHTML = html;
  // Mostrar u ocultar botón continuar
  var footer = lista.nextElementSibling;
  if(footer && footer.tagName === "BUTTON"){
    footer.style.display = lineas.length > 0 ? "" : "none";
  } else if(lineas.length > 0 && !document.querySelector("#gLineasLista + button")){
    var btn = document.createElement("button");
    btn.onclick = _guiasRevision;
    btn.style.cssText = "width:100%;padding:14px;background:var(--primary);color:white;border:none;border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;margin-top:16px";
    btn.textContent = "Revisar y generar guía →";
    lista.parentNode.appendChild(btn);
  }
}

function _guiasCargarOR(){
  document.getElementById("gORFile")?.click();
}

function _guiasProcesarOR(input){
  var file = input.files[0];
  if(!file) return;
  input.value = "";
  var reader = new FileReader();
  reader.onload = function(e){
    try{
      var wb = XLSX.read(e.target.result, {type:"array"});
      // Buscar primera hoja que no sea SAP
      var sheetName = wb.SheetNames.find(function(s){ return !s.startsWith("SAP"); });
      if(!sheetName){ alert("No se encontró hoja de datos en el archivo."); return; }
      var ws = wb.Sheets[sheetName];
      var rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
      // Buscar fila de encabezados
      var hdrIdx = -1;
      for(var i=0; i<rows.length; i++){
        if(rows[i].some(function(c){ return String(c||"").toUpperCase().includes("CATALOGO") || String(c||"").toUpperCase().includes("CATÁLOGO"); })){
          hdrIdx = i; break;
        }
      }
      if(hdrIdx < 0){ alert("No se encontró columna de Catálogo."); return; }
      var hdrs = rows[hdrIdx].map(function(h){ return String(h||"").toLowerCase(); });
      var iCat = hdrs.findIndex(function(h){ return h.includes("cat"); });
      var iXS  = hdrs.findIndex(function(h){ return h.includes("surtir") || h.includes("x surtir"); });
      if(iCat < 0 || iXS < 0){ alert("No se encontraron columnas de Catálogo y X Surtir."); return; }
      var importados = 0;
      for(var r=hdrIdx+1; r<rows.length; r++){
        var row = rows[r];
        var cat = String(row[iCat]||"").trim();
        var xs  = parseInt(row[iXS]||0);
        if(!cat || !xs || xs <= 0) continue;
        var m = mat(cat);
        var opciones = _guiasBDOpciones(cat);
        if(opciones.length === 1){
          var op = opciones[0];
          var esGranel = op.cont <= 1 || xs < op.cont;
          _guiasAgregarLinea(cat, m.desc||cat, m.um||"PZ", op.tipo, op.cont, xs, esGranel);
          importados++;
        } else if(opciones.length > 1){
          // Múltiples opciones — pendiente, el usuario elige
          _guiaActual.lineas.push({cat:cat, desc:m.desc||cat, um:m.um||"PZ", cant:xs,
            tipoEmp:"?", contEmp:0, bultos:xs, patio:false, granel:true, pendiente:true});
          importados++;
        } else {
          // Sin BD — agregar como granel por default
          _guiaActual.lineas.push({cat:cat, desc:m.desc||cat, um:m.um||"PZ", cant:xs,
            tipoEmp:"Granel", contEmp:0, bultos:xs, patio:false, granel:true, pendiente:false});
          importados++;
        }
      }
      _guiasRefrescarLineas();
      if(_guiaActual.lineas.some(function(l){ return l.pendiente; })){
        alert(importados + " materiales importados.\n\nAlgunos materiales no tienen empaque definido (marcados con ?).\nTócalos para completar el empaque antes de generar la guía.");
      } else {
        alert(importados + " materiales importados correctamente.");
      }
    }catch(err){
      alert("Error al leer el archivo: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── PANTALLA 4: Revisión y datos de firma ─────────────────────────────────────
function _guiasRevision(){
  if(_guiaActual.lineas.length === 0){ alert("No hay materiales capturados."); return; }
  if(_guiaActual.lineas.some(function(l){ return l.pendiente; })){
    alert("Hay materiales con empaque pendiente (?). Defínelos antes de continuar.");
    return;
  }

  var _gruposBultos = _guiasAgruparLineas(_guiaActual.lineas, _guiaActual.area);
  var totalBultos = 0;
  _gruposBultos.cajas.forEach(function(l){ totalBultos += Math.floor(l.cant/l.contEmp); });
  _gruposBultos.patio.forEach(function(l){ totalBultos += l.bultos || 1; });
  if(_gruposBultos.granel.length > 0) totalBultos += 1; // 1 caja colectiva

  var filasHtml = "";
  for(var i=0; i<_guiaActual.lineas.length; i++){
    var l = _guiaActual.lineas[i];
    var descEmp = l.bultos + " " + l.tipoEmp + (l.contEmp>1?" de "+l.contEmp+" "+l.um:"");
    filasHtml +=
      "<tr style=\"border-bottom:1px solid var(--lite,#f4f6fb)\">" +
      "<td style=\"padding:6px 8px;font-size:12px;font-weight:700;font-family:monospace;" +
      "color:var(--primary)\">" + l.cat + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px\">" + l.desc + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px;text-align:center\">" + l.cant + " " + l.um + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px;text-align:center\">" + descEmp + "</td>" +
      "<td style=\"padding:6px 8px;text-align:center\">" +
      "<button onclick=\"_guiasEditarLineaCompleta(" + i + ")\"" +
      " style=\"background:none;border:none;cursor:pointer;color:var(--primary);font-size:14px\">&#9998;</button>" +
      "</td>" +
      "</tr>";
  }

  $("#moduleView").innerHTML =
    "<div style=\"max-width:720px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:20px\">" +
    "<button onclick=\"_guiasCapturaMateriales()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Materiales</button>" +
    "<div><div style=\"font-size:18px;font-weight:800;color:var(--primary)\">" +
    "Guía " + _guiaActual.area + " No. " + _guiaActual.folio + " &mdash; Revisión</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">" + _guiaActual.destino + " &middot; " +
    _guiaActual.almInfo.nombre + " &middot; " + totalBultos + " bultos total</div></div>" +
    "</div>" +

    // Tabla de materiales
    "<div style=\"background:white;border:1px solid var(--line);border-radius:12px;" +
    "overflow:hidden;margin-bottom:20px\">" +
    "<div style=\"overflow-x:auto\">" +
    "<table style=\"width:100%;border-collapse:collapse\">" +
    "<thead style=\"background:var(--lite,#f4f6fb)\"><tr>" +
    "<th style=\"padding:8px;font-size:11px;color:var(--muted);text-align:left\">CATÁLOGO</th>" +
    "<th style=\"padding:8px;font-size:11px;color:var(--muted);text-align:left\">DESCRIPCIÓN</th>" +
    "<th style=\"padding:8px;font-size:11px;color:var(--muted);text-align:center\">CANTIDAD</th>" +
    "<th style=\"padding:8px;font-size:11px;color:var(--muted);text-align:center\">EMPAQUE</th>" +
    "<th style=\"padding:8px;width:32px\"></th>" +
    "</tr></thead><tbody>" + filasHtml + "</tbody>" +
    "</table></div></div>" +

    // Datos de firma
    "<div style=\"background:white;border:1px solid var(--line);border-radius:12px;" +
    "padding:16px;margin-bottom:20px\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:12px\">Datos para la guía</div>" +
    "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:12px\">" +
    _tplCampoFirma("gPedido", "No. Pedido", _guiaActual.pedido || "0") +
    _tplCampoFirma("gSiatel", "Siatel", _guiaActual.siatel || "0") +
    _tplCampoFirma("gSurtio", "Surtió (quien despacha)", _guiaActual.surtio) +
    _tplCampoFirma("gGenerador", "Capturista SAP", _guiaActual.generador) +
    _tplCampoFirma("gTransporteRev", "Línea de transporte", _guiaActual.transporte) +
    _tplCampoFirma("gOperador", "Operador", _guiaActual.operador) +
    _tplCampoFirma("gTipoVeh", "Tipo de vehículo", _guiaActual.tipoVeh) +
    _tplCampoFirma("gPlacas", "Placas", _guiaActual.placas) +
    "</div></div>" +

    "<button onclick=\"_guiasGenerar()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit\">" +
    "&#128203; Generar e imprimir guía</button>" +
    "</div>";
}

function _tplCampoFirma(id, label, valor){
  return "<div><label style=\"font-size:11px;color:var(--muted);display:block;margin-bottom:4px\">" +
    label + "</label>" +
    "<input id=\"" + id + "\" type=\"text\" value=\"" + (valor||"") + "\"" +
    " style=\"width:100%;padding:8px 12px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:13px\"></div>";
}

function _guiasEditarLineaCompleta(idx){
  var l = _guiaActual.lineas[idx];
  // Abrir modal de empaque con los datos actuales precargados
  var opciones = _guiasBDOpciones(l.cat);
  _guiasPedirEmpaque(l.cat, l.desc, l.um, opciones);
  // Cuando se confirme, reemplazará como línea nueva — primero quitamos la vieja
  _guiaActual.lineas.splice(idx, 1);
}

function _guiasEditarLinea(idx){
  var l = _guiaActual.lineas[idx];
  var cant = prompt("Nueva cantidad para " + l.cat + ":", l.cant);
  if(cant === null) return;
  cant = parseInt(cant);
  if(!cant || cant < 1){ alert("Cantidad inválida."); return; }
  l.cant = cant;
  l.bultos = l.contEmp > 1 ? Math.ceil(cant / l.contEmp) : cant;
  _guiasRevision(); // refrescar
}


var _GUIAS_COLECTIVO = {
  "Misceláneos":   "Material Misceláneo",
  "Herramientas":  "Herramienta en General",
  "Papelería":     "Papelería en General",
  "Ropa y Calzado":"Ropa y Calzado",
  "General":       "Material en General",
};


function _guiasAgruparLineas(lineas, area){
  var cajas   = [];
  var patio   = [];
  var granel  = [];

  for(var i=0; i<lineas.length; i++){
    var l = lineas[i];
    if(l.patio){
      patio.push(l);
    } else if(l.granel || l.contEmp === 0 || l.contEmp === 1 || l.cant < l.contEmp){
      // granel explícito, sin empaque definido, o cantidad menor al contenido del empaque
      granel.push(l);
    } else {
      cajas.push(l);
    }
  }

  // Ordenar cada grupo por catálogo ascendente
  var sortCat = function(a,b){ return a.cat.localeCompare(b.cat); };
  cajas.sort(sortCat);
  patio.sort(sortCat);
  granel.sort(sortCat);

  return { cajas: cajas, patio: patio, granel: granel };
}

function _guiasFilasImpresion(lineas, area){
  var grupos = _guiasAgruparLineas(lineas, area);
  var filas = [];
  var colectivo = _GUIAS_COLECTIVO[area] || "Material en General";

  // GRUPO 1: Cajas cerradas
  grupos.cajas.forEach(function(l){
    var nCajas = Math.floor(l.cant / l.contEmp);
    var descEmp = "C/C " + l.contEmp + " " + l.um;
    filas.push({
      cant:    nCajas,
      um:      descEmp,
      desc:    l.desc,
      cat:     l.cat,
      total:   l.cant + " " + l.um,
      tipo:    "caja",
      patio:   false
    });
  });

  // GRUPO 2: Patio
  grupos.patio.forEach(function(l){
    filas.push({
      cant:    l.bultos || l.cant,
      um:      l.um,
      desc:    l.desc,
      cat:     l.cat,
      total:   l.cant + " " + l.um,
      tipo:    "patio",
      patio:   true
    });
  });

  // GRUPO 3: Granel (solo si hay)
  if(grupos.granel.length > 0){
    // Renglón en blanco + encabezado colectivo
    filas.push({ tipo: "separador" });
    filas.push({ tipo: "colectivo", desc: colectivo });

    grupos.granel.forEach(function(l){
      filas.push({
        cant:    l.cant,
        um:      l.um,
        desc:    l.desc,
        cat:     l.cat,
        total:   l.cant + " " + l.um,
        tipo:    "granel",
        patio:   false
      });
    });
  }

  return filas;
}

// ── PANTALLA 5: Generar guía imprimible ───────────────────────────────────────
function _guiasGenerar(){
  // Capturar datos de firma
  _guiaActual.pedido     = document.getElementById("gPedido")?.value.trim() || "0";
  _guiaActual.siatel     = document.getElementById("gSiatel")?.value.trim() || "0";
  _guiaActual.surtio     = document.getElementById("gSurtio")?.value || "";
  _guiaActual.generador  = document.getElementById("gGenerador")?.value || "";
  _guiaActual.transporte = document.getElementById("gTransporteRev")?.value || "";
  _guiaActual.operador   = document.getElementById("gOperador")?.value || "";
  _guiaActual.tipoVeh    = document.getElementById("gTipoVeh")?.value || "";
  _guiaActual.placas     = document.getElementById("gPlacas")?.value || "";

  var alm  = _guiaActual.almInfo;
  var hoy  = _guiaActual.fecha ? new Date(_guiaActual.fecha+"T12:00:00") : new Date();
  var fechaStr = hoy.toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"});
  var _gruposBultos = _guiasAgruparLineas(_guiaActual.lineas, _guiaActual.area);
  var totalBultos = 0;
  _gruposBultos.cajas.forEach(function(l){ totalBultos += Math.floor(l.cant/l.contEmp); });
  _gruposBultos.patio.forEach(function(l){ totalBultos += l.bultos || 1; });
  if(_gruposBultos.granel.length > 0) totalBultos += 1; // 1 caja colectiva

  // Generar QR con datos de la guía (como URL de datos JSON)
  var qrData = JSON.stringify({
    folio: _guiaActual.folio, area: _guiaActual.area,
    destino: _guiaActual.destino, fecha: fechaStr,
    bultos: totalBultos, mats: _guiaActual.lineas.length
  });

  var _filas = _guiasFilasImpresion(_guiaActual.lineas, _guiaActual.area);
  var filasHtml = "";
  for(var fi=0; fi<_filas.length; fi++){
    var f = _filas[fi];
    if(f.tipo === "separador"){
      filasHtml += "<tr><td colspan=\"5\" style=\"padding:6px 0\">&nbsp;</td></tr>";
    } else if(f.tipo === "colectivo"){
      filasHtml += "<tr><td colspan=\"5\" style=\"padding:5px 8px;font-weight:700;" +
        "font-size:12px;border-top:1px solid #ccc\">" + f.desc + "</td></tr>";
    } else if(f.tipo === "patio"){
      filasHtml +=
        "<tr style=\"background:#fff8e1\">" +
        "<td class=\"col-cant\">" + f.cant + " - Patio</td>" +
        "<td class=\"col-emp\">" + f.um + "</td>" +
        "<td class=\"col-desc\">" + f.desc + "</td>" +
        "<td class=\"col-cat\">" + f.cat + "</td>" +
        "<td class=\"col-tot\">" + f.total + "</td></tr>";
    } else {
      filasHtml +=
        "<tr>" +
        "<td class=\"col-cant\">" + f.cant + "</td>" +
        "<td class=\"col-emp\">" + (f.tipo==="caja" ? f.um : f.um) + "</td>" +
        "<td class=\"col-desc\">" + f.desc + "</td>" +
        "<td class=\"col-cat\">" + f.cat + "</td>" +
        "<td class=\"col-tot\">" + f.total + "</td></tr>";
    }
  }

  var guiaHtml =
    "<!DOCTYPE html><html lang=\"es\"><head><meta charset=\"UTF-8\">" +
    "<title>Guía " + _guiaActual.area + " No." + _guiaActual.folio + " &mdash; D041</title>" +
    "<style>" +
    "body{font-family:Arial,sans-serif;font-size:11px;margin:0;padding:12px;color:#000}" +
    "h1{font-size:13px;margin:0}.hdr{display:grid;grid-template-columns:1fr auto;gap:12px;margin-bottom:8px}" +
    ".empresa{font-size:13px;font-weight:bold}.sub{font-size:11px}" +
    ".folio-box{border:2px solid #000;padding:6px 12px;text-align:right;min-width:160px}" +
    ".folio-box .label{font-size:10px;color:#555}.folio-box .num{font-size:16px;font-weight:bold}" +
    ".dest{border:1px solid #aaa;padding:6px 10px;margin-bottom:8px;font-size:11px}" +
    ".dest table{width:100%;border-collapse:collapse}" +
    ".dest td{padding:2px 4px;vertical-align:top}" +
    "table.items{width:100%;border-collapse:collapse;margin:8px 0}" +
    "table.items th{background:#001E6E;color:#fff;padding:5px 6px;font-size:10px;text-align:center;border:1px solid #001E6E}" +
    "table.items td{border:1px solid #ccc;padding:4px 6px;vertical-align:middle}" +
    ".col-cant{text-align:center;width:60px}.col-emp{width:160px}" +
    ".col-desc{}.col-cat{text-align:center;width:90px;font-weight:bold;font-family:monospace}" +
    ".col-tot{text-align:center;width:80px}" +
    ".firmas{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:16px}" +
    ".firma-box{border-top:1px solid #000;padding-top:4px;font-size:10px}" +
    ".transp{border:1px solid #aaa;padding:6px;font-size:11px;margin-top:8px}" +
    ".transp table{width:100%;border-collapse:collapse}" +
    ".transp td{padding:3px 6px;border-bottom:1px solid #eee}" +
    ".sello-box{border:1px dashed #999;height:60px;display:flex;align-items:center;" +
    "justify-content:center;color:#bbb;font-size:10px;margin-top:8px}" +
    "@media print{body{padding:0}button{display:none}}" +
    "</style></head><body>" +

    // Encabezado
    "<div class=\"hdr\">" +
    "<div style=\"display:flex;align-items:center;gap:12px\">" +
    "<img src=\"data:image/png;base64,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\" style=\"height:48px;object-fit:contain\">" +
    "<div>" +
    "<div class=\"empresa\">TELÉFONOS DE MÉXICO, S.A.B. DE C.V.</div>" +
    "<div class=\"sub\">ALMACÉN DISTRIBUIDOR PUEBLA</div>" +
    "<div class=\"sub\">Esquina Tepeyac, Calz. Ignacio Zaragoza, Los Pinos &nbsp; Tel. 2-57-73-90</div>" +
    "<div class=\"sub\">C.P. 72103 Heroica Puebla de Zaragoza, Pue.</div>" +
    "</div></div>" +
    "<div class=\"folio-box\">" +
    "<div style=\"text-align:center;margin-bottom:6px\">" +
    "<img src=\"data:image/png;base64,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\" style=\"height:34px;object-fit:contain\"></div>" +
    "<div class=\"label\">Guía de embarque</div>" +
    "<div class=\"label\">No. Guía &nbsp; <b>" + _guiaActual.folio + "</b></div>" +
    "<div class=\"label\">Pedido: <b>" + (_guiaActual.pedido||"0") + "</b></div>" +
    "<div class=\"label\">Siatel: <b>" + (_guiaActual.siatel||"0") + "</b></div>" +
    "<div class=\"label\">Área: " + _guiaActual.area + "</div>" +
    "</div>" +
    "</div>" +

    // Destinatario
    "<div class=\"dest\"><table>" +
    "<tr><td><b>Destinatario:</b></td><td>" + alm.nombre + "</td>" +
    "<td><b>Tel.:</b></td><td>" + (alm.tel||"") + "</td></tr>" +
    "<tr><td><b>Atiende:</b></td><td>" + (alm.atiende||"") + "</td>" +
    "<td><b>C.P.:</b></td><td>" + (alm.cp||"") + "</td></tr>" +
    "<tr><td><b>Domicilio:</b></td><td colspan=\"3\">" + (alm.domicilio||"") + " &nbsp; " + (alm.ciudad||"") + "</td></tr>" +
    "</table></div>" +

    // Tabla de materiales
    "<table class=\"items\">" +
    "<thead><tr>" +
    "<th class=\"col-cant\">Cantidad</th>" +
    "<th class=\"col-emp\">Unidad de embarque</th>" +
    "<th class=\"col-desc\">Descripción</th>" +
    "<th class=\"col-cat\">Catálogo</th>" +
    "<th class=\"col-tot\">Total</th>" +
    "</tr></thead><tbody>" + filasHtml + "</tbody>" +
    "<tfoot><tr><td colspan=\"5\" style=\"text-align:right;padding:5px 6px;font-weight:bold;font-size:12px\">" +
    "Total bultos: " + totalBultos + "</td></tr></tfoot>" +
    "</table>" +

    // Transporte
    "<div class=\"transp\"><table>" +
    "<tr><td><b>Surtió:</b></td><td>" + (_guiaActual.surtio||"") + "</td>" +
    "<td><b>Línea de transporte:</b></td><td>" + (_guiaActual.transporte||"") + "</td></tr>" +
    "<tr><td></td><td></td>" +
    "<td><b>Operador:</b></td><td>" + (_guiaActual.operador||"") + "</td></tr>" +
    "<tr><td><b>Recibe:</b></td><td style=\"border-bottom:1px solid #000;min-width:200px\">&nbsp;</td>" +
    "<td><b>Tipo de vehículo:</b></td><td>" + (_guiaActual.tipoVeh||"") + "</td></tr>" +
    "<tr><td></td><td></td>" +
    "<td><b>Placas:</b></td><td>" + (_guiaActual.placas||"") + "</td></tr>" +
    "</table></div>" +

    // Firmas
    "<div class=\"firmas\">" +
    "<div>" +
    "<div class=\"firma-box\">Nombre y firma &nbsp;&mdash;&nbsp; " + (_guiaActual.generador||"Capturista SAP") + "</div>" +
    "<div style=\"font-size:10px;margin-top:4px\">Fecha de embarque: " + fechaStr + "</div>" +
    "</div>" +
    "<div>" +
    "<div class=\"firma-box\">Fecha y firma &nbsp;&mdash;&nbsp; Transportista</div>" +
    "<div class=\"sello-box\">SELLO INSTITUCIONAL</div>" +
    "</div>" +
    "</div>" +

    "<div style=\"text-align:center;font-size:9px;color:#999;margin-top:8px\">" +
    "Sistema TX41 &middot; D041 Puebla &middot; Generado: " + new Date().toLocaleString("es-MX") +
    "</div>" +
    "<div style=\"margin-top:12px;display:flex;gap:10px\">" +
    "<button onclick=\"window.print()\" style=\"padding:10px 24px;background:#001E6E;" +
    "color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer\">" +
    "&#128424; Imprimir</button>" +
    "<button onclick=\"_guiasDescargarPDF()\" style=\"padding:10px 24px;background:#16a34a;" +
    "color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer\">" +
    "&#11015; Descargar PDF</button>" +
    "</div>" +
    "<script>" +
    "function _guiasDescargarPDF(){" +
    "  var opt = {margin:8, filename:'Guia_" + _guiaActual.area + "_No" + _guiaActual.folio + ".pdf'," +
    "  image:{type:'jpeg',quality:0.98}," +
    "  html2canvas:{scale:2,useCORS:true}," +
    "  jsPDF:{unit:'mm',format:'letter',orientation:'portrait'}};" +
    "  var el = document.body.cloneNode(true);" +
    "  el.querySelectorAll('button').forEach(function(b){b.remove();});" +
    "  if(typeof html2pdf !== 'undefined'){" +
    "    html2pdf().set(opt).from(el).save();" +
    "  } else {" +
    "    alert('Descargando... usa Imprimir → Guardar como PDF si no inicia la descarga.');" +
    "    window.print();" +
    "  }" +
    "}" +
    "<\/script>" +
    "<script src=\"https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js\"><\/script>" +
    "</body></html>";

  // Abrir en ventana nueva
  var win = window.open("", "_blank");
  win.document.write(guiaHtml);
  win.document.close();

  // Guardar en historial
  _guiasHistAgregar({
    folio: _guiaActual.folio, area: _guiaActual.area,
    destino: _guiaActual.destino, fecha: fechaStr,
    lineas: _guiaActual.lineas.length,
    datos: JSON.parse(JSON.stringify(_guiaActual))
  });
}

function _guiasBorrarHistorial(idx){
  if(!confirm("¿Borrar esta guía del historial? No se puede deshacer.")) return;
  var hist = _guiasHistCargar();
  hist.splice(idx, 1);
  _guiasHistGuardar(hist);
  modGuias();
}

function _guiasAbrirHistorial(idx){
  var hist = _guiasHistCargar();
  var g = hist[idx];
  if(!g) return;
  if(g.datos){
    // Tiene datos completos — reimprimir directamente
    _guiaActual = g.datos;
    _guiasGenerar();
  } else {
    alert("Guía " + g.area + " No. " + g.folio +
      "\nDestino: " + g.destino +
      "\nFecha: " + g.fecha +
      "\n\nEsta guía se generó con una versión anterior y no tiene datos para reimprimir.");
  }
}
