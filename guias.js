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
  // Mostrar modal de empaque
  var modal = document.createElement("div");
  modal.className = "modal on";

  var opcionesHtml = "";
  if(opciones.length > 0){
    opcionesHtml = ""; // ya no se usa directamente
    opcionesInnerHtml = "";
    
    for(var i=0; i<opciones.length; i++){
      var op = opciones[i];
      opcionesInnerHtml +=
        "<button onclick=\"_guiasSeleccionarEmpaque('" + cat + "','" +
        desc.replace(/'/g,"&#39;") + "','" + um + "','" +
        op.tipo + "'," + op.cont + ")\"" +
        " style=\"text-align:left;padding:10px 14px;background:#f0f4ff;border:1.5px solid var(--primary);" +
        "border-radius:8px;cursor:pointer;font-family:inherit;font-size:13px;width:100%\">" +
        "<b>" + op.tipo + "</b> de " + op.cont + " " + um +
        " <span style=\"color:var(--muted);font-size:11px\">(x" + (op.freq||1) + ")</span>" +
        "</button>";
    }
  }

  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:420px\">" +
    "<h3>Empaque &mdash; " + cat + "</h3>" +
    "<div class=\"modal-body\" style=\"gap:12px\">" +
    "<div style=\"font-size:13px;color:var(--muted)\">" + desc + "</div>" +
    opcionesHtml +
    "<div style=\"border-top:2px dashed var(--line);padding-top:14px;margin-top:4px\">" +
    "<div style=\"font-size:11px;font-weight:700;color:var(--muted);margin-bottom:8px\">" +
    (opciones.length > 0 ? "+ NUEVO EMPAQUE (diferente al conocido)" : "CAPTURA EL EMPAQUE") + "</div>" +
    "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:8px\">" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Tipo de empaque</label>" +
    "<input id=\"gEmpTipo\" type=\"text\" value=\"Caja\" placeholder=\"Caja, Costal, Pieza...\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit\"></div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Contenido por empaque</label>" +
    "<input id=\"gEmpCont\" type=\"number\" min=\"1\" value=\"1\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit\"></div>" +
    "</div>" +
    "<div style=\"margin-top:8px\">" +
    "<label style=\"font-size:11px;color:var(--muted)\">Cantidad total a despachar</label>" +
    "<input id=\"gEmpCant\" type=\"number\" min=\"1\" value=\"\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:16px;font-weight:700;color:var(--primary)\">" +
    "</div>" +
    "</div>" +
    "</div>" +
    "<div class=\"modal-foot\" style=\"gap:8px;flex-wrap:wrap\">" +
    "<button class=\"btn\" onclick=\"this.closest('.modal').remove()\">Cancelar</button>" +
    "<button class=\"btn\" style=\"color:var(--primary);border-color:var(--primary)\" " +
    "onclick=\"_guiasAGranel('" + cat + "','" + desc.replace(/'/g,'&#39;') + "','" + um + "')\">" +
    "A granel</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarEmpaque('" + cat + "','" +
    desc.replace(/'/g,'&#39;') + "','" + um + "')\">Agregar</button>" +
    "</div></div>";

  document.body.appendChild(modal);
  // Foco en cantidad
  setTimeout(function(){ var el=document.getElementById("gEmpCant"); if(el) el.focus(); }, 100);
}

function _guiasSeleccionarEmpaque(cat, desc, um, tipo, cont){
  // Pedir solo la cantidad, empaque ya está seleccionado
  var cant = prompt("Cantidad total de " + cat + " (" + desc + "):", "");
  if(!cant || isNaN(cant) || parseInt(cant) < 1) return;
  cant = parseInt(cant);
  _guiasAgregarLinea(cat, desc, um, tipo, parseInt(cont), cant);
  _guiasBDActualizarEmpaque(cat, tipo, parseInt(cont), um);
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
  if(!cant || cant < 1){ alert("Ingresa la cantidad a despachar a granel."); return; }
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
  _guiasAgregarLinea(cat, desc, um, tipo, cont, cant);
  _guiasBDActualizarEmpaque(cat, tipo, cont, um);
}

function _guiasAgregarLinea(cat, desc, um, tipoEmp, contEmp, cant, esGranel){
  var bultos = contEmp > 1 ? Math.ceil(cant / contEmp) : cant;
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: tipoEmp, contEmp: contEmp,
    bultos: bultos, patio: false, granel: esGranel || false
  });
  // Limpiar input y cerrar modal
  var inp = document.getElementById("gCatInput");
  if(inp){ inp.value = ""; document.getElementById("gCatInfo").textContent = ""; }
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
    "<div>" +
    "<div class=\"empresa\">TELÉFONOS DE MÉXICO, S.A.B. DE C.V.</div>" +
    "<div class=\"sub\">ALMACÉN DISTRIBUIDOR PUEBLA</div>" +
    "<div class=\"sub\">Esquina Tepeyac, Calz. Ignacio Zaragoza, Los Pinos &nbsp; Tel. 2-57-73-90</div>" +
    "<div class=\"sub\">C.P. 72103 Heroica Puebla de Zaragoza, Pue.</div>" +
    "</div>" +
    "<div class=\"folio-box\">" +
    "<div class=\"label\">Guía de embarque</div>" +
    "<div class=\"label\">No. Guía</div>" +
    "<div class=\"num\">" + _guiaActual.folio + "</div>" +
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
    "<button onclick=\"window.print()\" style=\"margin-top:12px;padding:10px 24px;background:#001E6E;" +
    "color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer\">" +
    "&#128424; Imprimir</button>" +
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
