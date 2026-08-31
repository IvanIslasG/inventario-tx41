/* ============================================================
   guias.js  -  Módulo Guías de Embarque · Sistema TX41 Puebla
   Depende de: DB, mat(), almName(), $(), XLSX (SheetJS)
   ============================================================ */

// ── BD de empaques (seed inicial — crece orgánicamente en localStorage) ──────
const _GUIAS_BD_SEED = [{"cat": "1000320", "um": "PZ", "tipo": "Caja", "cont": 192, "freq": 3, "patio": false}, {"cat": "1000372", "um": "PZ", "tipo": "Caja", "cont": 15, "freq": 1, "patio": false}, {"cat": "1000416", "um": "PZ", "tipo": "Caja", "cont": 150, "freq": 1, "patio": false}, {"cat": "1000430", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1000430", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1000381", "um": "PZ", "tipo": "Costal", "cont": 50, "freq": 1, "patio": false}, {"cat": "1000418", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1000600", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1000713", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1000715", "um": "PZ", "tipo": "Caja", "cont": 3000, "freq": 1, "patio": false}, {"cat": "1000902", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1000903", "um": "PZ", "tipo": "Caja", "cont": 5000, "freq": 1, "patio": false}, {"cat": "1001224", "um": "M", "tipo": "Caja", "cont": 250, "freq": 1, "patio": false}, {"cat": "1026180", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": true}, {"cat": "1028665", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1028954", "um": "PZ", "tipo": "Caja", "cont": 20000, "freq": 1, "patio": false}, {"cat": "1034256", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1034325", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1035461", "um": "PZ", "tipo": "Caja", "cont": 12, "freq": 1, "patio": false}, {"cat": "1036183", "um": "PZ", "tipo": "Caja", "cont": 140, "freq": 1, "patio": false}, {"cat": "1036183", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1036713", "um": "PZ", "tipo": "Caja", "cont": 6000, "freq": 1, "patio": false}, {"cat": "1038272", "um": "PZ", "tipo": "Caja", "cont": 150, "freq": 1, "patio": false}, {"cat": "1038272", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1045853", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1046965", "um": "PZ", "tipo": "Caja", "cont": 30, "freq": 1, "patio": false}, {"cat": "1052980", "um": "PZ", "tipo": "Caja", "cont": 500, "freq": 1, "patio": false}, {"cat": "1054083", "um": "PZ", "tipo": "Caja", "cont": 1, "freq": 1, "patio": false}, {"cat": "1054594", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 2, "patio": false}, {"cat": "1054594", "um": "PZ", "tipo": "Caja", "cont": 1000, "freq": 1, "patio": false}, {"cat": "1000291", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1000295", "um": "PZ", "tipo": "Caja", "cont": 24, "freq": 1, "patio": false}, {"cat": "1000296", "um": "PZ", "tipo": "Caja", "cont": 24, "freq": 1, "patio": false}, {"cat": "1000337", "um": "PZ", "tipo": "Caja", "cont": 10000, "freq": 1, "patio": false}, {"cat": "1000371", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1000674", "um": "PZ", "tipo": "Caja", "cont": 100, "freq": 1, "patio": false}, {"cat": "1002447", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": true}, {"cat": "1002449", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002470", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002502", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002447", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1002535", "um": "PZ", "tipo": "Pieza", "cont": 1, "freq": 1, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Pieza", "cont": 120, "freq": 3, "patio": false}, {"cat": "1002433", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002521", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002522", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002554", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002556", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1034253", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1034254", "um": "PZ", "tipo": "Caja", "cont": 200, "freq": 1, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Caja", "cont": 12, "freq": 2, "patio": false}, {"cat": "1034445", "um": "PZ", "tipo": "Caja", "cont": 50, "freq": 1, "patio": false}, {"cat": "1000508", "um": "M", "tipo": "Tubo", "cont": 6, "freq": 1, "patio": true}, {"cat": "1002353", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002413", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1002551", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1002734", "um": "PZ", "tipo": "Caja", "cont": 10, "freq": 1, "patio": false}, {"cat": "1010457", "um": "PZ", "tipo": "Caja", "cont": 20, "freq": 1, "patio": false}, {"cat": "1025372", "um": "PZ", "tipo": "Caja", "cont": 40, "freq": 1, "patio": false}, {"cat": "1026092", "um": "PZ", "tipo": "Caja", "cont": 6, "freq": 1, "patio": false}, {"cat": "1038152", "um": "PZ", "tipo": "Caja", "cont": 10, "freq": 1, "patio": false}];
const _GUIAS_ALM_SEED = {"D041": {"nombre": "ALMACEN DISTRIBUIDOR PUEBLA", "atiende": "GERMAN PEREZ PORRAS", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "A08A": {"nombre": "ALMACEN AUXILIAR SAN PEDRO", "atiende": "ASSENETH ALAVARADO HERRERA", "domicilio": "25 NORTE No. 3617  COL NUEVA AURORA", "ciudad": "PUEBLA, PUE.", "tel": "01 222 2463193 , 01 222 2424783 FAX", "cp": "72070"}, "A08B": {"nombre": "ALMACEN AUXILIAR ZARAGOZA", "atiende": "ROBERTO COYOTZI MONTES", "domicilio": "CALZ IGNACIO ZARAGOZA No 247 COL TEPEYAC", "ciudad": "PUEBLA, PUE.", "tel": "12222352031", "cp": ""}, "A08C": {"nombre": "ALMACEN AUXILIAR ORIZABA", "atiende": "VANIA MARTINEZ JUAREZ", "domicilio": "31 ORIENTE No 308 CIRCUNVALACION", "ciudad": "ORIZABA, VERACRUZ", "tel": "", "cp": ""}, "A08D": {"nombre": "ALMACEN AUXILIAR TLAXCALA", "atiende": "CARLOS VERGARA", "domicilio": "AV INSTITUTO POLITECNICO NACIONAL S/N", "ciudad": "TLAXCALA, TLAXCALA", "tel": "", "cp": ""}, "A08E": {"nombre": "ALMACEN AUXILIAR TEHUACAN", "atiende": "OMAR EURESTI MARTINEZ", "domicilio": "16 NORTE No 413 COL SERDAN", "ciudad": "TEHUACAN, PUEBLA", "tel": "12383821393", "cp": "75750"}, "A08F": {"nombre": "ALMACEN AUXILIAR SAN BRUNO", "atiende": "GONZALO SILCA POZOS", "domicilio": "MARTIREZ 28 DE AGOSTO", "ciudad": "XALAPA, VERACRUZ", "tel": "12288153631", "cp": "91020"}, "A08G": {"nombre": "ALMACEN AUXILIAR CRISTAL", "atiende": "LAURA MANZO SILVESTRE, OMAR BAIZABAL", "domicilio": "ANTONIO CHEDRAHUI CARAM No 250", "ciudad": "XALAPA, VERACRUZ", "tel": "12288150444", "cp": "91180"}, "A08H": {"nombre": "ALMACEN AUXILIAR MOCAMBO", "atiende": "HIRAM SALAMANCA", "domicilio": "AV PALMERAS No 305 COL JARDINES DE VIRGINIA", "ciudad": "BOCA DEL RIO, VERACRUZ", "tel": "12299214122", "cp": ""}, "A08I": {"nombre": "ALMACEN AUXILIAR DOS CAMINOS", "atiende": "DIEGO GARCÍA", "domicilio": "AV 11 No 2627", "ciudad": "CORDOBA, VERACRUZ", "tel": "12717165433", "cp": ""}, "A08J": {"nombre": "ALMACEN AUXILIAR AQUILES SERDAN", "atiende": "JOSÉ ANTONIO LÓPEZ C", "domicilio": "35 PONIENTE No 723", "ciudad": "PUEBLA, PUE.", "tel": "12222433928", "cp": ""}, "A08K": {"nombre": "ALMACEN AUXILIAR ALTAMIRANO", "atiende": "LILIA CONTRERAS", "domicilio": "ALTAMIRANO No 1226", "ciudad": "VERACRUZ, VERACRUZ", "tel": "12299204842", "cp": "91700"}, "A08L": {"nombre": "ALMACEN AUXILIAR LERDO", "atiende": "JULIO CÉSAR GONZÁLEZ", "domicilio": "MARIANO ARISTA No 4424", "ciudad": "VERACRUZ, VERACRUZ", "tel": "12299204842", "cp": "91726"}, "A08M": {"nombre": "ALMACEN AUXILIAR PEÑUELA", "atiende": "JUAN LUNA", "domicilio": "KM 343 BOULEVARD S/N", "ciudad": "CORDOBA, VERACRUZ", "tel": "12717166707", "cp": "94501"}, "A08T": {"nombre": "ALMACEN AUXILIAR TULANCINGO", "atiende": "PEDRO MORALES LIRA", "domicilio": "RIVA PALACIOS 203 COL LOS ALAMOS", "ciudad": "TULANCINGO, HIDALGO", "tel": "17757536701", "cp": "43640"}, "A08V": {"nombre": "ALMACEN AUXILIAR PACHUCA", "atiende": "HIRAM ROSALES", "domicilio": "SAN MARTIN DE PORRES No 407 CFE", "ciudad": "PACHUCA DE SOTO, HIDALGO", "tel": "17717143677", "cp": "42090"}, "A08X": {"nombre": "ALMACEN AUXILIAR ATLIXCO", "atiende": "SERGIO NOE MILIAN PINTLE", "domicilio": "CALZ. OAXACA No 2710", "ciudad": "ATLIXCO, PUEBLA", "tel": "12444451868", "cp": "74294"}, "A08Y": {"nombre": "ALMACEN AUXILIAR MAYORAZGO", "atiende": "GLORIA MARCELA SILVA T", "domicilio": "CALLE BENITO JUAREZ No 12311", "ciudad": "PUEBLA, PUE.", "tel": "12222192800", "cp": ""}, "A08Z": {"nombre": "ALMACEN AUXILIAR POZA RICA", "atiende": "ENRIQUE LEANDRO BLANCO", "domicilio": "POZO 2 No 48 COL DIVISION NORTE", "ciudad": "POZA RICA, VERACRUZ", "tel": "17828233004", "cp": "93350"}, "A09C": {"nombre": "ALMACEN AUXILIAR COATZACOALCOS", "atiende": "JUANA MARIA RAMOS", "domicilio": "AV. JACINTO LEMARROY S/N FRACC. RANCHO ALEGRE 2", "ciudad": "COATZACOALCOS, VERACRUZ", "tel": "19212186550", "cp": "96558"}, "A09S": {"nombre": "ALMACEN AUXILIAR MINATITLAN", "atiende": "SERGIO VAZQUEZ", "domicilio": "EMILIANO ZAPATA No. 25 COL. INSURGENTES SUR", "ciudad": "MINATITLAN, VERACRUZ", "tel": "19222211979", "cp": "96710"}, "A09T": {"nombre": "ALMACEN AUXILIAR CHOLULA", "atiende": "ILIANA DE ITA, CARLOS GUTIERREZ", "domicilio": "22 ORIENTE No 606", "ciudad": "CHOLULA, PUEBLA", "tel": "12222474746", "cp": ""}, "A09U": {"nombre": "ALMACEN AUXILIAR AMALUCAN", "atiende": "JOSEFINA ONOFRE M", "domicilio": "BLVRD ATEOPAN ESQ 17A", "ciudad": "PUEBLA, PUE.", "tel": "12222868207", "cp": ""}, "A09V": {"nombre": "ALMACEN AUXILIAR TULA", "atiende": "PILAR HERNANDEZ, VERONICA MEZA", "domicilio": "AV SUR ESQUINA OTE No 25 COL CENTRO", "ciudad": "TULA DE ALLENDE, HIDALGO", "tel": "17737322888", "cp": "42800"}, "A0X4": {"nombre": "ALMACEN TELECO", "atiende": "", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "A0Y4": {"nombre": "ALMACEN CYCSA PUEBLA", "atiende": "ERNESTO CENTENO", "domicilio": "CALLE C No. 7 PARQUE INDUSTRIAL 2000", "ciudad": "PUEBLA, PUE.", "tel": "", "cp": ""}, "A082": {"nombre": "ALMACEN AUXILIAR TUXPAN", "atiende": "JAVIER CRUZ YAÑEZ", "domicilio": "DEMETRIO RUIZ MALERVA S/N COL CENTRO", "ciudad": "TUXPAN VERACRUZ", "tel": "17838340234", "cp": ""}, "A081": {"nombre": "ALMACEN AUXILIAR PACHOACAN", "atiende": "FELIPE OLANDES MORA", "domicilio": "AV FRNACISCO SARABIA No 100 COL CARLOS ROVIROSA", "ciudad": "PACHUCA DE SOTO, HIDALGO", "tel": "1771148651", "cp": "42082"}, "A083": {"nombre": "ALMACEN AUXILIAR SAN MARTIN", "atiende": "AGUSTIN BALBUENA SCHIAFINI", "domicilio": "AV. JUVENTUD S/N COL. LOS DICIOS", "ciudad": "SAN MARTIN TEXMELUCAN, PUEBLA", "tel": "", "cp": ""}, "D032": {"nombre": "ALMACEN DISTRIBUIDOR CELAYA", "atiende": "JOSE TORADO", "domicilio": "AV LAS FUENTES 10904 COL LAS FUENTES", "ciudad": "CELAYA, GTO", "tel": "", "cp": ""}, "D022": {"nombre": "ALMACEN DISTRIBUIDOR HERMOSILLO", "atiende": "BRUNO GENDA FERNANDEZ", "domicilio": "BLVRD JESUS GARCIA MORALES KM 5 No 145 COL EL LLANO", "ciudad": "HERMOSILLO, SONORA", "tel": "16622188701", "cp": "83210"}, "A09P": {"nombre": "ALMACEN AUXILIAR PINOTEPA", "atiende": "JUAN PEDRO GONZALEZ MARTINEZ", "domicilio": "AV AGUIRRE PALANCARES Y BAÑOS AGUIRRE", "ciudad": "PINOTEPA NACIONAL", "tel": "19545433906", "cp": "71600"}, "A09O": {"nombre": "ALMACEN AUXILIAR OAXACA 1", "atiende": "CARLOS TEJEDA", "domicilio": "HEROICO COLEGIO MILITAR No 1013 COL REFORMA", "ciudad": "OAXACA, OAX", "tel": "19515127100", "cp": "68050"}, "A091": {"nombre": "ALMACEN AUXILIAR APIZACO", "atiende": "LOURDES ESPINOSA", "domicilio": "VENUSTIANO CARRANZA ESQ JOSE A", "ciudad": "APIZACO, TLAXCALA", "tel": "12414174555", "cp": "90350"}, "A084": {"nombre": "ALMACEN AUXILIAR TEZIUTLAN", "atiende": "RAMON GALINDO BECERRA", "domicilio": "AV ENCINO Y AVELLANO SN", "ciudad": "TEZIUTLAN, PUEBLA", "tel": "12313130715", "cp": "73890"}, "A09Q": {"nombre": "ALMACEN AUXILIAR OAXACA II", "atiende": "LEON RUIS MATADAMAS", "domicilio": "ESMERALDA No 201 COL BUGAMBILIAS", "ciudad": "OAXACA, OAX", "tel": "19515127100", "cp": "68010"}, "A09R": {"nombre": "ALMACEN AUXILIAR HUAJUAPAN", "atiende": "J CARLOS GARCIA, VEDA CAROLINA OSORIO", "domicilio": "PROLONGACION DE MINA No 120", "ciudad": "H. CIUDAD DE HUAJUAPAN DE LEON, OAX", "tel": "19535324388", "cp": "69007"}, "D011": {"nombre": "ALMACEN DISTRIBUIDOR MONTERREY", "atiende": "ANTONIO AGUILAR", "domicilio": "CORDILLERA DE LOS ANDES No 701 JARDIN DE LAS PTES S/N0", "ciudad": "SN NICOLAS DE LOS GARZA, NUEVO LEON", "tel": "18183505572", "cp": "66460"}, "A0LE": {"nombre": "ALMACEN AUXILIAR LEGARIA", "atiende": "HERNESTO HERNANDEZ", "domicilio": "FELIPE CARRILLO PUERTO No. 750 TORRE BLANCA  MIGUEL HIDALGO", "ciudad": "MEXICO, DF", "tel": "", "cp": ""}, "D043": {"nombre": "ALMACEN DISTRIBUIDOR VILLAHERMOSA", "atiende": "EDUARDO TABOADA", "domicilio": "AV. ACERO S/N ESQUINA COBRE CD.INDUSTRIAL", "ciudad": "VILLAHERMOSA. TABASCO", "tel": "", "cp": ""}, "A09X": {"nombre": "ALMACEN AUXILIAR CHETUMAL", "atiende": "ARTURO ALONSO RAMIREZ", "domicilio": "AVENIDA 4 DE MARZO NO.30 COL.FIDEL VELAZQUEZ", "ciudad": "CHETUMAL QUINTANA ROO", "tel": "19838372365", "cp": "77080"}, "D021": {"nombre": "ALMACEN DISTRIBUIDOR GUADALAJARA", "atiende": "ENRIQUE MEDINA LOPEZ", "domicilio": "TRATADO DE TLALTELOCO No. 4114 COL. PARQUE AUDITORIO", "ciudad": "ZAPOPAN, JALISCO", "tel": "3336601554", "cp": ""}, "D013": {"nombre": "ALMACEN DISTRIBUIDOR CHIHUAHUA", "atiende": "CARLOS DURAN", "domicilio": "MIGUEL BARRAGAN No. 6903 COL. EL PARRAL", "ciudad": "CHIHUAHUA, CHIHUAHUA", "tel": "", "cp": ""}, "D007": {"nombre": "ALMACEN GENERAL LA PERLA", "atiende": "ROGER CANO", "domicilio": "CALLE NUEVA ESQUINA NEGRA MODELO COL. INDUSTRIAL", "ciudad": "NAUCALPAN EDO DE MEXICO", "tel": "", "cp": ""}, "MTZ": {"nombre": "ALMACEN AUXILIAR MARTINEZ DE LA TORRE", "atiende": "ROLANDO RAZGADO DE JESUS", "domicilio": "BOULEVARD ALFINO FLORES S/N COL. ADOLFO RUIZ CORTINA", "ciudad": "MARTINEZ DE LA TORRE, VER.", "tel": "", "cp": ""}, "C0TA": {"nombre": "CARSO TLAXCALA", "atiende": "", "domicilio": "CALLE AV.OCOTLAN S/N CONTRA ESQ. CALLE CONSTRUCTORES", "ciudad": "SANTA ANA CHIAUTEMPAN", "tel": ",012464620646", "cp": ""}, "ACTOPAN": {"nombre": "NUEVO ACTOPAN", "atiende": "PAREDES TENORIO FRANCISCA", "domicilio": "PEDRO MORENO No 49", "ciudad": "NUEVO ACTOPAN, HGO", "tel": "17727273187", "cp": "42500"}, "A09I": {"nombre": "MERIDA BUENAVISTA", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "MERIDA", "tel": "-", "cp": "-"}, "SAHAGUN": {"nombre": "CT SAHAGUN", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "CD SAHAGÚN", "tel": "-", "cp": "-"}, "TECA": {"nombre": "CT TECAMACHALCO", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "TECAMACHALCO", "tel": "-", "cp": "-"}, "A0AS": {"nombre": "ALMACEN ABASTOS", "atiende": "CONOCIDO", "domicilio": "CONOCIDO", "ciudad": "&", "tel": "&", "cp": "&"}, "A0CD": {"nombre": "ALMACEN AUXILIAR  DIANA", "atiende": "LUIS FELIPE ESTROP", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}, "SLPI": {"nombre": "ALMACEN AUXILIAR PINO", "atiende": "", "domicilio": "", "ciudad": "", "tel": "", "cp": ""}};
const _GUIAS_LS_BD    = "guias_bd_empaques_v1";
const _GUIAS_LS_HIST  = "guias_historial_v1";

// Semilla de tarimas conocidas: cat + tipo de contenedor + cuántos contenedores por tarima + contenido c/u
const _GUIAS_TARIMA_SEED = [
  {cat:"1054521", um:"M",   subTipo:"Bobina", subCant:60,  subCont:500, freq:1},
  {cat:"1054521", um:"M",   subTipo:"Bobina", subCant:48,  subCont:500, freq:1},
  {cat:"1001221", um:"M",   subTipo:"Rollo",  subCant:72,  subCont:300, freq:1},
  {cat:"1036187", um:"PZA", subTipo:"Caja",   subCant:117, subCont:1,   freq:1},
  {cat:"1036187", um:"PZA", subTipo:"Caja",   subCant:153, subCont:1,   freq:1},
  {cat:"1036188", um:"PZA", subTipo:"Caja",   subCant:117, subCont:1,   freq:1},
  {cat:"1036188", um:"PZA", subTipo:"Caja",   subCant:153, subCont:1,   freq:1},
  {cat:"1039312", um:"PZA", subTipo:"Caja",   subCant:117, subCont:1,   freq:1},
  {cat:"1039312", um:"PZA", subTipo:"Caja",   subCant:153, subCont:1,   freq:1}
];
const _GUIAS_LS_TARIMA = "guias_bd_tarimas_v1";
const _GUIAS_LS_DRAFT  = "guias_borrador_v1";
const _GUIAS_LS_ULTBK  = "guias_ultimo_backup_v1";
const _GUIAS_BACKUP_INTERVALO_DIAS = 7; // cada cuántos días se dispara el respaldo automático

// ── Cargar/guardar BD de empaques ────────────────────────────────────────────
var _guiasBDSeedMerged = false;
function _guiasBDCargar(){
  var bd = null;
  try{
    const raw = localStorage.getItem(_GUIAS_LS_BD);
    if(raw) bd = JSON.parse(raw);
  }catch(e){}

  if(bd === null){
    // Primera vez en este navegador — usar la semilla completa
    bd = JSON.parse(JSON.stringify(_GUIAS_BD_SEED));
    _guiasBDGuardar(bd);
    _guiasBDSeedMerged = true;
    return bd;
  }

  // Fusionar la semilla del código (nuevos empaques homologados) sin pisar lo que ya
  // tiene este navegador. Solo se hace una vez por sesión, no en cada llamada.
  if(!_guiasBDSeedMerged){
    var cambio = false;
    _GUIAS_BD_SEED.forEach(function(se){
      var yaExiste = bd.some(function(e){ return e.cat === se.cat && e.cont === se.cont; });
      if(!yaExiste){
        bd.push(JSON.parse(JSON.stringify(se)));
        cambio = true;
      } else if(se.preferido){
        // Si la semilla ya trae un preferido definido y este navegador aún no tiene uno para ese catálogo, adoptarlo
        var tienePreferido = bd.some(function(e){ return e.cat === se.cat && e.preferido; });
        if(!tienePreferido){
          var idx = bd.findIndex(function(e){ return e.cat === se.cat && e.cont === se.cont; });
          if(idx >= 0){ bd[idx].preferido = true; cambio = true; }
        }
      }
    });
    if(cambio) _guiasBDGuardar(bd);
    _guiasBDSeedMerged = true;
  }
  return bd;
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
           .sort(function(a,b){
             var pa = a.preferido ? 1 : 0, pb = b.preferido ? 1 : 0;
             if(pa !== pb) return pb - pa; // preferido siempre primero
             return (b.freq||0)-(a.freq||0);
           });
}

// Fija un empaque como preferido para un catálogo: siempre se propondrá primero,
// sin importar la frecuencia de uso de otras opciones.
function _guiasBDMarcarPreferido(cat, cont){
  var bd = _guiasBDCargar();
  bd.forEach(function(e){
    if(e.cat === cat) e.preferido = (e.cont === cont);
  });
  _guiasBDGuardar(bd);
}

function _guiasTogglePreferido(cat, desc, um, cont){
  _guiasBDMarcarPreferido(cat, cont);
  document.querySelector(".modal")?.remove();
  _guiasPedirEmpaque(cat, desc, um, _guiasBDOpciones(cat));
}

// Borra un empaque conocido (cat+contenido) de la memoria (localStorage)
function _guiasBDBorrarOpcion(cat, desc, um, cont){
  if(!confirm("¿Borrar este empaque de la memoria?\n\n" + cat + " — empaque de " + cont + " " + um + "\n\nYa no se propondrá para este catálogo.")) return;
  var bd = _guiasBDCargar();
  var idx = bd.findIndex(function(e){ return e.cat === cat && e.cont === cont; });
  if(idx >= 0){ bd.splice(idx, 1); _guiasBDGuardar(bd); }
  document.querySelector(".modal")?.remove();
  _guiasPedirEmpaque(cat, desc, um, _guiasBDOpciones(cat));
}

// ── Cargar/guardar BD de tarimas (mismo patrón que la BD de empaques) ─────────
var _guiasTarimaSeedMerged = false;
function _guiasTarimaBDCargar(){
  var bd = null;
  try{
    const raw = localStorage.getItem(_GUIAS_LS_TARIMA);
    if(raw) bd = JSON.parse(raw);
  }catch(e){}

  if(bd === null){
    bd = JSON.parse(JSON.stringify(_GUIAS_TARIMA_SEED));
    _guiasTarimaBDGuardar(bd);
    _guiasTarimaSeedMerged = true;
    return bd;
  }

  if(!_guiasTarimaSeedMerged){
    var cambio = false;
    _GUIAS_TARIMA_SEED.forEach(function(se){
      var yaExiste = bd.some(function(e){
        return e.cat === se.cat && e.subTipo === se.subTipo && e.subCant === se.subCant && e.subCont === se.subCont;
      });
      if(!yaExiste){ bd.push(JSON.parse(JSON.stringify(se))); cambio = true; }
    });
    if(cambio) _guiasTarimaBDGuardar(bd);
    _guiasTarimaSeedMerged = true;
  }
  return bd;
}

function _guiasTarimaBDGuardar(bd){
  try{ localStorage.setItem(_GUIAS_LS_TARIMA, JSON.stringify(bd)); }catch(e){}
}

function _guiasTarimaBDActualizar(cat, subTipo, subCant, subCont, um){
  var bd = _guiasTarimaBDCargar();
  var idx = bd.findIndex(function(e){
    return e.cat === cat && e.subTipo === subTipo && e.subCant === subCant && e.subCont === subCont;
  });
  if(idx >= 0){ bd[idx].freq = (bd[idx].freq || 1) + 1; }
  else{ bd.push({cat: cat, um: um || '', subTipo: subTipo, subCant: subCant, subCont: subCont, freq: 1}); }
  _guiasTarimaBDGuardar(bd);
}

function _guiasTarimaOpciones(cat){
  var bd = _guiasTarimaBDCargar();
  return bd.filter(function(e){ return e.cat === cat; })
           .sort(function(a,b){
             var pa = a.preferido ? 1 : 0, pb = b.preferido ? 1 : 0;
             if(pa !== pb) return pb - pa;
             return (b.freq||0)-(a.freq||0);
           });
}

function _guiasTarimaBDMarcarPreferido(cat, subTipo, subCant, subCont){
  var bd = _guiasTarimaBDCargar();
  bd.forEach(function(e){
    if(e.cat === cat) e.preferido = (e.subTipo === subTipo && e.subCant === subCant && e.subCont === subCont);
  });
  _guiasTarimaBDGuardar(bd);
}

function _guiasTarimaTogglePreferido(cat, desc, um, subTipo, subCant, subCont){
  _guiasTarimaBDMarcarPreferido(cat, subTipo, subCant, subCont);
  document.querySelector(".modal")?.remove();
  _guiasPedirEmpaque(cat, desc, um, _guiasBDOpciones(cat));
}

function _guiasTarimaBorrarOpcion(cat, desc, um, subTipo, subCant, subCont){
  if(!confirm("¿Borrar esta tarima de la memoria?\n\nTarima de " + subCant + " " + subTipo + " c/ " + subCont + " " + um + "\n\nYa no se propondrá para este catálogo.")) return;
  var bd = _guiasTarimaBDCargar();
  var idx = bd.findIndex(function(e){
    return e.cat === cat && e.subTipo === subTipo && e.subCant === subCant && e.subCont === subCont;
  });
  if(idx >= 0){ bd.splice(idx, 1); _guiasTarimaBDGuardar(bd); }
  document.querySelector(".modal")?.remove();
  _guiasPedirEmpaque(cat, desc, um, _guiasBDOpciones(cat));
}

// Agrega una línea de tarima (contenedor anidado: N tarimas de X contenedores de Y c/u)
function _guiasAgregarLineaTarima(cat, desc, um, subTipo, subCant, subCont, numTarimas){
  if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }
  var contEmp = subCant * subCont; // cantidad total que representa una tarima
  var cant = numTarimas * contEmp;
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: "Tarima", contEmp: contEmp,
    bultos: numTarimas, patio: false, granel: false,
    tarimaSubTipo: subTipo, tarimaSubCant: subCant, tarimaSubCont: subCont
  });
  _limpiarCatInput();
  document.querySelector(".modal")?.remove();
  _guiasRefrescarVista();
}

function _guiasSeleccionarTarima(cat, desc, um, subTipo, subCant, subCont){
  var num = prompt("¿Cuántas tarimas de " + subCant + " " + subTipo + " c/ " + subCont + " " + um + " van?", "1");
  if(!num || isNaN(num) || parseInt(num) < 1) return;
  _guiasTarimaBDActualizar(cat, subTipo, subCant, subCont, um);
  _guiasAgregarLineaTarima(cat, desc, um, subTipo, subCant, subCont, parseInt(num));
}

function _guiasConfirmarTarimaNueva(cat, desc, um){
  var subTipo = (document.getElementById("gTarimaTipo")?.value || "").trim();
  var subCant = parseInt(document.getElementById("gTarimaSubCant")?.value || "0");
  var subCont = parseInt(document.getElementById("gTarimaSubCont")?.value || "0");
  var numTarimas = parseInt(document.getElementById("gTarimaNum")?.value || "0");
  if(!subTipo){ alert("Indica el tipo de contenedor (Caja, Bobina, Rollo...)."); return; }
  if(!subCant || subCant < 1){ alert("Indica cuántos contenedores lleva cada tarima."); return; }
  if(!subCont || subCont < 1){ alert("Indica el contenido de cada contenedor."); return; }
  if(!numTarimas || numTarimas < 1){ alert("Indica cuántas tarimas van."); return; }
  _guiasTarimaBDActualizar(cat, subTipo, subCant, subCont, um);
  _guiasAgregarLineaTarima(cat, desc, um, subTipo, subCant, subCont, numTarimas);
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
  // Techo de seguridad muy alto (no un límite real de uso) — solo para evitar que localStorage
  // crezca sin control si algún día pasa algo raro. Con uso normal, nunca debería alcanzarse.
  if(hist.length > 5000) hist = hist.slice(0,5000);
  _guiasHistGuardar(hist);
}

// ── Info de almacén (BD propia + directorio DB) ───────────────────────────────
// ── Escapar texto para insertarlo seguro en atributos HTML ───────────────────
function _escAttr(str){
  return String(str==null?"":str).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

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
var _guiaEditandoOriginal = null; // { folio, area } de la guía que se está editando desde el historial, o null si es nueva

// ── Borrador automático de la guía en progreso ────────────────────────────────
// Guarda _guiaActual en localStorage cada pocos segundos MIENTRAS se está armando,
// independiente del guardado final al generar. Así, si se va la luz o se cierra el
// navegador a medias, al volver a entrar se puede recuperar lo que ya se había capturado.
var _guiasDraftUltimoJSON = null;
function _guiasDraftGuardar(){
  if(!_guiaActual) return;
  // Solo vale la pena respaldar si ya hay algo capturado (destino elegido o al menos una línea)
  if(!_guiaActual.destino && (!_guiaActual.lineas || !_guiaActual.lineas.length)) return;
  try{
    var snap = JSON.stringify({guardado: Date.now(), datos: _guiaActual, editando: _guiaEditandoOriginal});
    if(snap === _guiasDraftUltimoJSON) return; // nada cambió, no reescribir
    localStorage.setItem(_GUIAS_LS_DRAFT, snap);
    _guiasDraftUltimoJSON = snap;
  }catch(e){}
}
function _guiasDraftCargar(){
  try{
    var raw = localStorage.getItem(_GUIAS_LS_DRAFT);
    return raw ? JSON.parse(raw) : null;
  }catch(e){ return null; }
}
function _guiasDraftBorrar(){
  try{ localStorage.removeItem(_GUIAS_LS_DRAFT); }catch(e){}
  _guiasDraftUltimoJSON = null;
}
// Autoguardado periódico mientras haya una guía en progreso, más un intento extra
// al cambiar de pestaña/minimizar o cerrar — para no depender de un solo momento.
setInterval(_guiasDraftGuardar, 4000);
document.addEventListener("visibilitychange", function(){ if(document.hidden) _guiasDraftGuardar(); });
window.addEventListener("beforeunload", _guiasDraftGuardar);
function _guiasDraftRetomar(){
  var d = _guiasDraftCargar();
  if(!d) return;
  _guiaActual = d.datos;
  _guiaEditandoOriginal = d.editando || null;
  _guiasDraftBorrar();
  _guiasCapturaMateriales();
}
function _guiasDraftDescartar(){
  if(!confirm("¿Descartar el borrador sin terminar? Esta acción no se puede deshacer.")) return;
  _guiasDraftBorrar();
  modGuias();
}
var _guiasEditandoLineaIdx = null; // índice de línea que se está reemplazando (edición desde Revisión), o null

// Lotes reales (excluye el lote "NUEVO", que significa sin lote/no aplica)
function _guiasLotesReales(cat){
  try{
    if(typeof lotesDe !== "function") return [];
    return lotesDe(cat).filter(function(l){ return (l.lote||"").toUpperCase() !== "NUEVO"; });
  }catch(e){ return []; }
}

// ¿El material pertenece al área de Cables? (para forzar tipo "Bobina")
function _guiasEsCable(cat){
  try{
    var m = mat(cat);
    return (m.area || "").toLowerCase() === "cables" && (m.um || "").trim().toLowerCase() === "m";
  }catch(e){ return false; }
}

// Excepciones de tipo de empaque por catálogo: algunos cables no van en Bobina sino en otro contenedor.
var _GUIAS_TIPO_EXCEPCIONES = {
  "1001221": "Rollo"
};

// Tipo de empaque forzado para un catálogo (excepción específica, o "Bobina" si es cable en metros), o null si es libre
function _guiasTipoForzado(cat){
  if(_GUIAS_TIPO_EXCEPCIONES[cat]) return _GUIAS_TIPO_EXCEPCIONES[cat];
  if(_guiasEsCable(cat)) return "Bobina";
  return null;
}

var _guiasORPendiente = null; // { rows, hdrIdx, iCat, iXS } de un OR importado desde la pantalla de Datos, pendiente de aplicar a las líneas

// Si estamos en la pantalla de Revisión, guarda lo ya escrito en los campos de firma
// para no perderlo al refrescar la vista después de editar/borrar una línea.
function _guiasCapturarFirmasRevisionSiActiva(){
  if(!document.getElementById("gSurtio") || !_guiaActual) return;
  _guiaActual.pedido     = document.getElementById("gPedido")?.value.trim() || _guiaActual.pedido || "0";
  _guiaActual.siatel     = document.getElementById("gSiatel")?.value.trim() || _guiaActual.siatel || "0";
  _guiaActual.surtio     = document.getElementById("gSurtio")?.value || "";
  _guiaActual.transporte = document.getElementById("gTransporteRev")?.value || "";
  _guiaActual.operador   = document.getElementById("gOperador")?.value || "";
  _guiaActual.tipoVeh    = document.getElementById("gTipoVeh")?.value || "";
  _guiaActual.placas     = document.getElementById("gPlacas")?.value || "";
  _guiaActual.observaciones = document.getElementById("gObservaciones")?.value || "";
}

// Refresca la vista correcta según en qué pantalla estemos (Revisión o Captura de materiales)
function _guiasRefrescarVista(){
  if(document.getElementById("gSurtio")){
    if(_guiaActual.lineas.length === 0){
      _guiasCapturaMateriales(); // no quedan materiales: regresar a captura
    } else {
      _guiasRevision();
    }
  } else {
    _guiasRefrescarLineas();
  }
}


// ── Exportar / Importar guías y BD de empaques ───────────────────────────────
function _guiasExportarTodo(){
  _guiasExportarSeleccion(null); // null = todas
}

function _guiasExportarSeleccion(indices, nombreArchivo){
  var hist = _guiasHistCargar();
  var bd   = _guiasBDCargar();
  var datos = indices === null ? hist : indices.map(function(i){ return hist[i]; });

  if(!datos.length){ alert("No hay guías para exportar."); return; }

  var bundle = {
    version:   "tx41-guias-v1",
    exportado: new Date().toISOString(),
    guias:     datos,
    bd_empaques: bd
  };

  var json = JSON.stringify(bundle, null, 2);
  var blob = new Blob([json], {type: "application/json"});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement("a");
  a.href   = url;
  a.download = nombreArchivo || ("guias_tx41_" + new Date().toLocaleDateString("es-MX").replace(/\//g,"-") + ".json");
  a.click();
  URL.revokeObjectURL(url);
}

// ── Respaldo automático programado ────────────────────────────────────────────
// No es un cron real del sistema (esto es una página web, solo corre si está abierta):
// cada vez que se abre el menú de Guías, si ya pasó _GUIAS_BACKUP_INTERVALO_DIAS desde el
// último respaldo, se descarga solo el historial completo, sin pedir nada.
function _guiasBackupAutoCheck(){
  try{
    var hist = _guiasHistCargar();
    if(!hist.length) return; // nada que respaldar todavía
    var ultimo = +localStorage.getItem(_GUIAS_LS_ULTBK) || 0;
    var dias = (Date.now() - ultimo) / 86400000;
    if(dias < _GUIAS_BACKUP_INTERVALO_DIAS) return;
    var fecha = new Date().toLocaleDateString("es-MX").replace(/\//g,"-");
    _guiasExportarSeleccion(null, "Respaldo_automatico_guias_" + fecha + ".json");
    localStorage.setItem(_GUIAS_LS_ULTBK, String(Date.now()));
    _guiasToast("📦 Respaldo automático descargado — revisa tu carpeta de Descargas", "ok");
  }catch(e){}
}

function _guiasImportar(){
  var input = document.createElement("input");
  input.type   = "file";
  input.accept = ".json";
  input.onchange = function(){
    var file = input.files[0];
    if(!file) return;
    var reader = new FileReader();
    reader.onload = function(e){
      try{
        var bundle = JSON.parse(e.target.result);
        if(!bundle.version || !bundle.guias){
          alert("El archivo no es un respaldo válido de guías TX41."); return;
        }

        // Fusionar guías (por folio+area como clave, no duplicar)
        var hist = _guiasHistCargar();
        var importadas = 0;
        bundle.guias.forEach(function(g){
          var existe = hist.some(function(h){ return h.folio===g.folio && h.area===g.area; });
          if(!existe){ hist.unshift(g); importadas++; }
        });
        _guiasHistGuardar(hist);

        // Fusionar BD de empaques
        var bdLocal = _guiasBDCargar();
        var bdImport = bundle.bd_empaques || [];
        var empAgg = 0;
        bdImport.forEach(function(e){
          var idx = bdLocal.findIndex(function(l){ return l.cat===e.cat && l.cont===e.cont; });
          if(idx < 0){ bdLocal.push(e); empAgg++; }
          else if((e.freq||0) > (bdLocal[idx].freq||0)){ bdLocal[idx] = e; }
        });
        _guiasBDGuardar(bdLocal);

        alert("Importación completada:\n" +
          importadas + " guías nuevas\n" +
          empAgg + " empaques nuevos en BD");
        modGuias();
      } catch(err){
        alert("Error al leer el archivo: " + err.message);
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

// Genera el código de la semilla (_GUIAS_BD_SEED) a partir de la BD actual de este navegador,
// para pegarlo en guias.js y que todos los equipos lo reciban al cargar la página.
function _guiasBDExportarSemilla(){
  var bd = _guiasBDCargar();
  if(!bd.length){ alert("No hay empaques capturados todavía."); return; }
  var ordenado = bd.slice().sort(function(a,b){
    return a.cat.localeCompare(b.cat) || (a.cont - b.cont);
  });
  var jsonArray = JSON.stringify(ordenado);
  var codigo = "const _GUIAS_BD_SEED = " + jsonArray + ";";

  var modal = document.createElement("div");
  modal.className = "modal on";
  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:700px\">" +
    "<h3 style=\"margin:0 0 8px\">Semilla de empaques</h3>" +
    "<p style=\"font-size:12px;color:var(--muted);margin:0 0 14px\">" +
    ordenado.length + " empaques capturados en este navegador (incluye tus preferidos &#9733;).</p>" +

    "<div style=\"border:1.5px solid var(--primary);border-radius:10px;padding:14px;margin-bottom:14px;background:#f0f4ff\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--primary);margin-bottom:6px\">Opción rápida: actualizar tu guias.js automáticamente</div>" +
    "<div style=\"font-size:11px;color:var(--muted);margin-bottom:10px\">Sube tu archivo guias.js actual y te regreso el mismo archivo ya con la semilla sustituida, listo para subir a tu repo.</div>" +
    "<input type=\"file\" id=\"gSeedJSFile\" accept=\".js\" style=\"display:none\" onchange=\"_guiasBDActualizarArchivoJS(this, " +
    "'" + jsonArray.replace(/\\/g,"\\\\").replace(/'/g,"\\'").replace(/"/g,"&quot;") + "')\">" +
    "<button class=\"btn-prim\" onclick=\"document.getElementById('gSeedJSFile').click()\">&#8679; Subir guias.js y generar el actualizado</button>" +
    "</div>" +

    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);margin-bottom:6px\">Opción manual: copiar el código</div>" +
    "<textarea readonly onclick=\"this.select()\" style=\"width:100%;height:160px;font-family:monospace;" +
    "font-size:11px;padding:8px;border:1.5px solid var(--line);border-radius:8px;box-sizing:border-box\">" +
    codigo.replace(/</g,"&lt;") + "</textarea>" +
    "<div style=\"display:flex;justify-content:flex-end;gap:8px;margin-top:12px\">" +
    "<button class=\"btn\" onclick=\"this.closest('.modal').remove()\">Cerrar</button>" +
    "<button class=\"btn-prim\" onclick=\"navigator.clipboard.writeText(this.closest('.modal-box')." +
    "querySelector('textarea').value).then(function(){alert('Copiado al portapapeles.');})." +
    "catch(function(){alert('No se pudo copiar automático — selecciona el texto y usa Ctrl+C.');})\">" +
    "Copiar código</button>" +
    "</div></div>";
  document.body.appendChild(modal);
}

// Busca "const _GUIAS_BD_SEED = [ ... ];" en el texto del archivo y reemplaza el arreglo,
// contando corchetes para encontrar el cierre exacto sin importar el formato del archivo.
function _guiasReemplazarSeedEnTexto(texto, nuevoArrayJson){
  var marcador = "const _GUIAS_BD_SEED = ";
  var inicio = texto.indexOf(marcador);
  if(inicio < 0) return null;
  var posArray = inicio + marcador.length;
  if(texto[posArray] !== '[') return null;
  var depth = 0, i = posArray;
  for(; i < texto.length; i++){
    if(texto[i] === '[') depth++;
    else if(texto[i] === ']'){ depth--; if(depth === 0){ i++; break; } }
  }
  var finPunto = texto.indexOf(';', i);
  if(finPunto < 0) return null;
  return texto.slice(0, posArray) + nuevoArrayJson + texto.slice(finPunto);
}

function _guiasBDActualizarArchivoJS(input, jsonArray){
  var file = input.files[0];
  if(!file) return;
  input.value = "";
  var reader = new FileReader();
  reader.onload = function(e){
    var texto = e.target.result;
    var actualizado = _guiasReemplazarSeedEnTexto(texto, jsonArray);
    if(!actualizado){
      alert("No encontré la línea \"const _GUIAS_BD_SEED = [...]\" en ese archivo — asegúrate de subir tu guias.js.");
      return;
    }
    var blob = new Blob([actualizado], {type: "text/javascript"});
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement("a");
    a.href   = url;
    a.download = "guias.js";
    a.click();
    URL.revokeObjectURL(url);
    alert("Listo — se descargó guias.js con la semilla actualizada. Súbelo a tu repositorio para publicarlo.");
    document.querySelector(".modal")?.remove();
  };
  reader.readAsText(file);
}



function _guiasGetSeleccionados(){
  var checks = document.querySelectorAll(".guia-chk:checked");
  return Array.from(checks).map(function(c){ return parseInt(c.dataset.idx); });
}

function _guiasToggleExportBtn(){
  var n = document.querySelectorAll(".guia-chk:checked").length;
  var btnExp = document.getElementById("btnExportSel");
  if(btnExp){ btnExp.disabled = n === 0; btnExp.textContent = n > 0 ? "Exportar seleccionadas (" + n + ")" : "Exportar seleccionadas"; }
  var btnDel = document.getElementById("btnBorrarSel");
  if(btnDel){ btnDel.disabled = n === 0; btnDel.textContent = n > 0 ? "Borrar seleccionadas (" + n + ")" : "Borrar seleccionadas"; }
}

function _guiasBorrarSeleccionadas(){
  var indices = _guiasGetSeleccionados();
  if(indices.length === 0) return;
  if(!confirm("¿Borrar " + indices.length + " guía" + (indices.length>1?"s":"") + " seleccionada" + (indices.length>1?"s":"") + "?\n\nEsta acción no se puede deshacer.")) return;
  var hist = _guiasHistCargar();
  indices.sort(function(a,b){ return b-a; }).forEach(function(idx){ hist.splice(idx, 1); });
  _guiasHistGuardar(hist);
  modGuias();
}

// ── PANTALLA 1: Menú principal ────────────────────────────────────────────────
var _guiasAreaSel = null;   // área abierta en el listado (null = ninguna, tarjetas cerradas)

var _GUIAS_AREAS = ["Herramientas","Misceláneos","Papelería","Cables","Ropa y Calzado","General"];

function _guiasAgrupar(hist){
  var porArea = {};
  _GUIAS_AREAS.forEach(function(a){ porArea[a] = []; });
  hist.forEach(function(g,i){
    var a = g.area || "General";
    if(!porArea[a]) porArea[a] = [];
    porArea[a].push({g:g, i:i});
  });
  return porArea;
}

function modGuias(){
  _guiasBackupAutoCheck();
  var hist = _guiasHistCargar();
  var porArea = _guiasAgrupar(hist);

  // Si el área abierta se quedó sin guías, se cierra
  if(_guiasAreaSel && !(porArea[_guiasAreaSel] || []).length) _guiasAreaSel = null;

  var tarjetas = "";
  Object.keys(porArea).forEach(function(area){
    var n = (porArea[area] || []).length;
    var abierta = _guiasAreaSel === area;
    tarjetas +=
      "<div class=\"guia-area-card" + (abierta ? " abierta" : "") + (n === 0 ? " vacia" : "") + "\"" +
      " data-area=\"" + _escAttr(area) + "\">" +
      "<div class=\"ga-n\">" + n + "</div>" +
      "<div class=\"ga-t\">" + area + "</div>" +
      "<div class=\"ga-s\">" + (n === 0 ? "sin guías" : (n === 1 ? "1 guía" : n + " guías")) + "</div>" +
      "</div>";
  });

  // Si quedó un borrador de una guía sin terminar (ej. por un apagón o cierre inesperado),
  // se ofrece recuperarlo antes que nada — sin él, ese trabajo se hubiera perdido por completo.
  var draft = (!_guiaActual) ? _guiasDraftCargar() : null;
  var draftHtml = "";
  if(draft && draft.datos){
    var dd = draft.datos;
    var cuando = new Date(draft.guardado).toLocaleString('es-MX',{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'});
    draftHtml =
      "<div style=\"background:#fff8e1;border:1.5px solid #f0c93a;border-radius:12px;padding:14px 16px;margin-bottom:14px;" +
      "display:flex;align-items:center;gap:12px;flex-wrap:wrap\">" +
      "<div style=\"flex:1;min-width:220px\">" +
      "<div style=\"font-weight:700;font-size:13.5px\">📝 Hay una guía sin terminar</div>" +
      "<div style=\"font-size:12px;color:var(--muted);margin-top:2px\">" +
      (dd.area||"Sin área") + (dd.folio?(" · No. "+dd.folio):"") + (dd.destino?(" · "+dd.destino):"") +
      " · " + (dd.lineas?dd.lineas.length:0) + " línea(s) · respaldada " + cuando + "</div></div>" +
      "<button onclick=\"_guiasDraftRetomar()\" style=\"padding:8px 16px;background:var(--primary);color:#fff;" +
      "border:none;border-radius:8px;font-weight:700;font-size:12.5px;cursor:pointer;font-family:inherit\">Continuar</button>" +
      "<button onclick=\"_guiasDraftDescartar()\" style=\"padding:8px 14px;background:none;color:#a8390f;" +
      "border:1.5px solid #f0c93a;border-radius:8px;font-size:12.5px;cursor:pointer;font-family:inherit\">Descartar</button>" +
      "</div>";
  }

  var ultBk = +localStorage.getItem(_GUIAS_LS_ULTBK) || 0;
  var ultBkTxt = ultBk ? new Date(ultBk).toLocaleDateString('es-MX',{day:'2-digit',month:'short'}) : "nunca";

  $("#moduleView").innerHTML =
    "<div style=\"max-width:900px;margin:0 auto;padding:16px 0\">" +
    "<div class=\"menu-h\" style=\"margin-bottom:14px\">" +
    "<h1>Guías de Embarque</h1>" +
    "<p>" + DIST() + " &middot; " + hist.length + " guía" + (hist.length===1?"":"s") + " guardada" + (hist.length===1?"":"s") +
    " &middot; último respaldo automático: " + ultBkTxt + "</p>" +
    "</div>" +
    draftHtml +

    "<button onclick=\"_guiasDraftBorrar();_guiaActual=null;_guiaEditandoOriginal=null;_guiasNueva()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;" +
    "margin-bottom:10px\">+ Nueva guía de embarque</button>" +

    "<div style=\"display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap\">" +
    "<button id=\"btnExportSel\" disabled onclick=\"_guiasExportarSeleccion(_guiasGetSeleccionados())\"" +
    " style=\"flex:1;min-width:150px;padding:9px;background:white;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;color:var(--primary);\">" +
    "Exportar seleccionadas</button>" +
    "<button id=\"btnBorrarSel\" disabled onclick=\"_guiasBorrarSeleccionadas()\"" +
    " style=\"flex:1;min-width:150px;padding:9px;background:white;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;color:#dc2626;\">" +
    "Borrar seleccionadas</button>" +
    "</div>" +

    "<div style=\"display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap\">" +
    "<button onclick=\"_guiasExportarTodo()\"" +
    " style=\"flex:1;min-width:150px;padding:9px;background:white;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;color:var(--primary)\">" +
    "Exportar todo</button>" +
    "<button onclick=\"_guiasImportar()\"" +
    " style=\"flex:1;min-width:150px;padding:9px;background:white;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;color:var(--primary)\">" +
    "&#8679; Importar</button>" +
    "</div>" +

    "<details style=\"margin-bottom:18px\">" +
    "<summary style=\"font-size:11px;color:var(--muted);cursor:pointer;list-style:none;" +
    "display:inline-block\">&#9881; Opciones avanzadas</summary>" +
    "<button onclick=\"_guiasBDExportarSemilla()\"" +
    " style=\"margin-top:8px;width:100%;padding:7px;background:none;border:1px dashed var(--line);" +
    "border-radius:8px;font-size:11px;cursor:pointer;font-family:inherit;color:var(--muted)\">" +
    "Generar código de semilla de empaques</button>" +
    "</details>" +

    "<input type=\"search\" id=\"guiaSearch\" placeholder=\"Buscar por folio, destino o área...\"" +
    " style=\"width:100%;padding:9px 12px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-family:inherit;font-size:13px;margin-bottom:14px;box-sizing:border-box\">" +

    "<div class=\"guia-areas-grid\" id=\"guiaAreasGrid\">" + tarjetas + "</div>" +
    "<div id=\"guiaPanel\"></div>" +
    "</div>";

  $("#guiaAreasGrid").querySelectorAll(".guia-area-card").forEach(function(card){
    card.onclick = function(){
      var area = card.dataset.area;
      _guiasAreaSel = (_guiasAreaSel === area) ? null : area;
      modGuias();
    };
  });

  _guiasRenderPanel();

  var buscador = document.getElementById("guiaSearch");
  if(buscador){
    buscador.oninput = function(){ _guiasBuscar(this.value); };
  }
}

// ── Panel inferior: guías del área abierta ───────────────────────────────────
function _guiasRenderPanel(){
  var panel = document.getElementById("guiaPanel");
  if(!panel) return;
  if(!_guiasAreaSel){ panel.innerHTML = ""; return; }

  var items = _guiasAgrupar(_guiasHistCargar())[_guiasAreaSel] || [];
  panel.innerHTML =
    "<div class=\"guia-panel-head\">" +
    "<span>" + _guiasAreaSel + "</span>" +
    "<span class=\"pill\">" + items.length + "</span>" +
    "<button class=\"guia-cerrar\" onclick=\"_guiasAreaSel=null;modGuias()\">Cerrar &times;</button>" +
    "</div>" +
    "<div class=\"guia-lista\">" + items.map(_guiasFilaHtml).join("") + "</div>";
  _guiasToggleExportBtn();
}

function _guiasFilaHtml(item){
  var g = item.g, i = item.i;
  var area = g.area || "General";
  var buscable = (g.folio + " " + (g.destino||"") + " " + area + " " + (g.fecha||"")).toLowerCase();
  return "<div class=\"guia-item\" data-buscar=\"" + _escAttr(buscable) + "\">" +
    "<input type=\"checkbox\" class=\"guia-chk\" data-idx=\"" + i + "\"" +
    " onchange=\"_guiasToggleExportBtn()\" style=\"width:15px;height:15px;cursor:pointer;flex-shrink:0\">" +
    "<div style=\"flex:1;min-width:0;cursor:pointer\" onclick=\"_guiasAbrirHistorial(" + i + ")\" title=\"Ver / editar\">" +
    "<div class=\"gi-t\">No. " + g.folio + " &mdash; " + (g.destino||"") + "</div>" +
    "<div class=\"gi-s\">" + (g.fecha||"") + " &middot; " + (g.lineas||0) + " materiales</div>" +
    "</div>" +
    "<button onclick=\"_guiasReimprimirHistorial(" + i + ",event)\" title=\"Reimprimir\"" +
    " style=\"background:none;border:none;color:var(--primary);cursor:pointer;font-size:15px;padding:0\">&#128424;</button>" +
    "<button onclick=\"_guiasBorrarHistorial(" + i + ")\" title=\"Borrar\"" +
    " style=\"background:none;border:none;color:#dc2626;cursor:pointer;font-size:16px;padding:0\">&times;</button>" +
    "</div>";
}

// ── Búsqueda: resultados planos de todas las áreas ───────────────────────────
function _guiasBuscar(val){
  var q = (val||"").trim().toLowerCase();
  var grid = document.getElementById("guiaAreasGrid");
  var panel = document.getElementById("guiaPanel");
  if(!grid || !panel) return;

  if(!q){                       // sin búsqueda: se regresa a las tarjetas
    grid.style.display = "";
    _guiasRenderPanel();
    return;
  }

  grid.style.display = "none";  // durante la búsqueda las tarjetas estorban
  var hist = _guiasHistCargar();
  var res = hist.map(function(g,i){ return {g:g, i:i}; }).filter(function(item){
    var g = item.g;
    var s = (g.folio + " " + (g.destino||"") + " " + (g.area||"General") + " " + (g.fecha||"")).toLowerCase();
    return s.indexOf(q) !== -1;
  });

  if(!res.length){
    panel.innerHTML = "<div class=\"empty\">Sin resultados para esa búsqueda.</div>";
    return;
  }
  panel.innerHTML =
    "<div class=\"guia-panel-head\"><span>Resultados</span>" +
    "<span class=\"pill\">" + res.length + "</span></div>" +
    "<div class=\"guia-lista\">" + res.map(function(item){
      return _guiasFilaHtml(item).replace("<div class=\"gi-s\">",
        "<div class=\"gi-s\"><b style=\"color:var(--primary)\">" + (item.g.area||"General") + "</b> &middot; ");
    }).join("") + "</div>";
  _guiasToggleExportBtn();
}

// ── PANTALLA 2: Nueva guía — datos de cabecera ────────────────────────────────
var _guiasDestinoLista = [];      // [{sigla, nombre}] — se llena al construir la pantalla
var _guiasDestinoFiltrados = [];
var _guiasDestinoHighlight = -1;

function _guiasNueva(){
  var editando = !!_guiaActual; // si ya hay una guía en curso, estamos editando/regresando a Datos

  // Calcular próximo folio por área (el usuario puede cambiarlo)
  var hist = _guiasHistCargar();

  // Obtener áreas disponibles
  var areas = ["Herramientas","Misceláneos","Papelería","Cables","Ropa y Calzado","General"];

  var areasHtml = areas.map(function(a){
    var sel = editando && _guiaActual.area === a ? " selected" : "";
    return "<option value=\"" + a + "\"" + sel + ">" + a + "</option>";
  }).join("");

  // Almacenes para el autocompletado de destino (del directorio DB)
  var almsDb = Object.keys(DB.directorio?.almacenes||{}).sort();
  _guiasDestinoLista = almsDb.map(function(k){
    var info = _guiasAlmInfo(k);
    return { sigla: k, nombre: info.nombre || "" };
  });

  $("#moduleView").innerHTML =
    "<div style=\"max-width:560px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:24px\">" +
    "<button type=\"button\" onclick=\"modGuias()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Cancelar</button>" +
    "<h2 style=\"margin:0;font-size:18px\">" + (editando ? "Editar guía" : "Nueva guía") + "</h2>" +
    "</div>" +

    "<form id=\"gNuevaForm\" onsubmit=\"event.preventDefault();_guiasContinuarMateriales();return false;\">" +

    (editando ? "" :
    "<div style=\"margin-bottom:20px;display:flex;flex-direction:column;gap:6px\">" +
    "<button type=\"button\" onclick=\"_guiasCargarORDesdeNueva()\"" +
    " style=\"width:100%;padding:10px;background:white;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;" +
    "color:var(--primary)\">&#8679; Importar desde OR (Excel) &mdash; autocompleta Área, Destino y Fecha</button>" +
    "<input type=\"file\" id=\"gORFileNueva\" accept=\".xlsx\" style=\"display:none\"" +
    " onchange=\"_guiasProcesarORDesdeNueva(this)\">" +
    "<button type=\"button\" onclick=\"_guiasAbrirPegarTabla(true)\"" +
    " style=\"width:100%;padding:10px;background:white;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;" +
    "color:var(--primary)\">&#128203; Pegar tabla de un correo</button>" +
    "</div>") +

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
    "letter-spacing:.4px;display:block;margin-bottom:6px\">No. de Guía</label>" +
    "<input id=\"gFolio\" type=\"number\" min=\"1\"" +
    (editando ? " value=\"" + _escAttr(_guiaActual.folio) + "\"" : "") +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:16px;font-weight:700;font-family:inherit;color:var(--primary);box-sizing:border-box\">" +
    "</div>" +

    // Destino (autocompletado propio, con navegación por teclado)
    "<div style=\"margin-bottom:16px;position:relative\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Destino</label>" +
    "<input id=\"gDestino\" autocomplete=\"off\" placeholder=\"Escribe o selecciona almacén...\"" +
    (editando ? " value=\"" + _escAttr(_guiaActual.destino) + "\"" : "") +
    " oninput=\"_guiasDestinoInput()\" onfocus=\"_guiasDestinoInput()\"" +
    " onkeydown=\"_guiasDestinoKeydown(event)\"" +
    " onblur=\"setTimeout(function(){ var c=document.getElementById('gDestinoSugerencias'); if(c) c.style.display='none'; },150)\"" +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:14px;font-family:inherit;box-sizing:border-box\">" +
    "<div id=\"gDestinoSugerencias\" style=\"position:absolute;left:0;right:0;top:100%;margin-top:2px;" +
    "background:white;border:1.5px solid var(--line);border-radius:10px;max-height:220px;overflow-y:auto;" +
    "z-index:50;display:none;box-shadow:0 8px 20px rgba(0,0,0,.12)\"></div>" +
    "</div>" +

    // Info destinatario (se llena automático)
    "<div id=\"gDestinatarioInfo\" style=\"background:var(--lite,#f0f4ff);border-radius:10px;" +
    "padding:12px 16px;margin-bottom:16px;font-size:12px;color:var(--muted);display:none\">" +
    "</div>" +

    // Fecha
    "<div style=\"margin-bottom:24px\">" +
    "<label style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;display:block;margin-bottom:6px\">Fecha</label>" +
    "<input id=\"gFecha\" type=\"date\"" +
    (editando && _guiaActual.fecha ? " value=\"" + _escAttr(_guiaActual.fecha) + "\"" : "") +
    " style=\"width:100%;padding:10px 14px;border:1.5px solid var(--line);border-radius:10px;" +
    "font-size:14px;font-family:inherit;box-sizing:border-box\">" +
    "</div>" +

    // Botón continuar
    "<button type=\"submit\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit\">" +
    (editando ? "Continuar &rarr; Materiales" : "Continuar &rarr; Capturar materiales") + "</button>" +
    "</form>" +
    "</div>";

  if(editando){
    _guiasActualizarDestinatario();
  } else {
    // Poner fecha de hoy por default
    var hoy = new Date();
    var yyyy = hoy.getFullYear();
    var mm = String(hoy.getMonth()+1).padStart(2,'0');
    var dd = String(hoy.getDate()).padStart(2,'0');
    document.getElementById("gFecha").value = yyyy+"-"+mm+"-"+dd;
  }
}

function _guiasDestinoInput(){
  var inp = document.getElementById("gDestino");
  if(!inp) return;
  var q = inp.value.trim().toLowerCase();
  _guiasDestinoFiltrados = _guiasDestinoLista.filter(function(a){
    return !q || a.sigla.toLowerCase().includes(q) || a.nombre.toLowerCase().includes(q);
  }).slice(0, 30);
  _guiasDestinoHighlight = -1;
  _guiasDestinoRenderSugerencias();
  _guiasActualizarDestinatario();
}

function _guiasDestinoRenderSugerencias(){
  var cont = document.getElementById("gDestinoSugerencias");
  if(!cont) return;
  if(_guiasDestinoFiltrados.length === 0){ cont.style.display = "none"; cont.innerHTML = ""; return; }
  cont.innerHTML = _guiasDestinoFiltrados.map(function(a, idx){
    var resaltado = idx === _guiasDestinoHighlight;
    return "<div data-idx=\"" + idx + "\" onmousedown=\"event.preventDefault();_guiasDestinoElegir(" + idx + ")\"" +
      " style=\"padding:9px 14px;cursor:pointer;font-size:13px;" +
      (resaltado ? "background:var(--primary);color:white" : "background:white;color:inherit") + "\">" +
      "<b>" + a.sigla + "</b> &mdash; " + a.nombre + "</div>";
  }).join("");
  cont.style.display = "block";
  if(_guiasDestinoHighlight >= 0){
    var activo = cont.querySelector("[data-idx=\"" + _guiasDestinoHighlight + "\"]");
    if(activo) activo.scrollIntoView({block:"nearest"});
  }
}

function _guiasDestinoElegir(idx){
  var a = _guiasDestinoFiltrados[idx];
  if(!a) return;
  var inp = document.getElementById("gDestino");
  inp.value = a.sigla;
  var cont = document.getElementById("gDestinoSugerencias");
  if(cont) cont.style.display = "none";
  _guiasActualizarDestinatario();
}

function _guiasDestinoKeydown(ev){
  var cont = document.getElementById("gDestinoSugerencias");
  var visible = !!(cont && cont.style.display !== "none" && _guiasDestinoFiltrados.length > 0);

  if(ev.key === "ArrowDown"){
    ev.preventDefault();
    if(!visible){ _guiasDestinoInput(); return; }
    _guiasDestinoHighlight = Math.min(_guiasDestinoHighlight + 1, _guiasDestinoFiltrados.length - 1);
    _guiasDestinoRenderSugerencias();
    return;
  }
  if(ev.key === "ArrowUp"){
    ev.preventDefault();
    if(!visible) return;
    _guiasDestinoHighlight = Math.max(_guiasDestinoHighlight - 1, 0);
    _guiasDestinoRenderSugerencias();
    return;
  }
  if(ev.key === "Enter"){
    if(visible){
      ev.preventDefault();
      _guiasDestinoElegir(_guiasDestinoHighlight >= 0 ? _guiasDestinoHighlight : 0);
    }
    // Si no hay sugerencias visibles, se deja pasar el Enter para que el formulario continúe.
    return;
  }
  if(ev.key === "Tab"){
    if(visible){
      _guiasDestinoElegir(_guiasDestinoHighlight >= 0 ? _guiasDestinoHighlight : 0);
    }
    return; // no se bloquea: el Tab sigue avanzando al siguiente campo
  }
  if(ev.key === "Escape"){
    if(cont) cont.style.display = "none";
  }
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

  if(!folio){ alert("Ingresa el número de guía."); return; }
  if(!destino){ alert("Selecciona el almacén destino."); return; }

  var almInfo = _guiasAlmInfo(destino);
  var previa = _guiaActual; // si venimos de editar, aquí ya hay materiales y otros datos capturados
  _guiaActual = {
    area: area, folio: parseInt(folio), destino: destino,
    almInfo: almInfo, fecha: fecha, transporte: previa ? previa.transporte : '', // se captura en la pantalla de firma
    lineas: previa ? previa.lineas : [],
    surtio:    previa ? previa.surtio    : '',
    operador:  previa ? previa.operador  : '',
    placas:    previa ? previa.placas    : '',
    tipoVeh:   previa ? previa.tipoVeh   : '',
    pedido:    previa ? previa.pedido    : '',
    siatel:    previa ? previa.siatel    : '',
    cajasGranel: previa && previa.cajasGranel != null ? previa.cajasGranel : 0
  };

  // Si se importó un OR desde esta pantalla, aplicar ahora los materiales
  if(_guiasORPendiente){
    var res = _guiasImportarFilasOR(_guiasORPendiente.rows, _guiasORPendiente.hdrIdx, _guiasORPendiente.iCat, _guiasORPendiente.iXS);
    _guiasORPendiente = null;
    _guiasCapturaMateriales();
    if(res.paraRevisar > 0){
      _guiasToast(res.importados + " materiales guardados &mdash; " + res.paraRevisar + " por revisar", "warn");
      alert(res.importados + " materiales importados.\n\n" + res.paraRevisar +
        " se marcaron con \"Revisar\" (empaque ambiguo o desconocido). El sistema propuso lo más probable, pero conviene confirmarlo con el lápiz antes de generar la guía.");
    } else {
      _guiasToast(res.importados + " materiales procesados y guardados correctamente", "ok");
    }
    return;
  }

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
    "<button onclick=\"_guiasAbrirPegarTabla(false)\"" +
    " style=\"flex:1;padding:10px;background:white;border:1.5px solid var(--line);" +
    "border-radius:10px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;" +
    "color:var(--primary)\">&#128203; Pegar tabla de un correo</button>" +
    "</div>" +

    // Lista de líneas capturadas
    "<div id=\"gLineasLista\" style=\"display:flex;flex-direction:column;gap:8px\">" +
    lineasHtml +
    "</div>" +
    footerBtns +
    "</div>";
}

function _tplLineaGuia(l, idx){
  var esGranel = (!l.patio) && !!l.granel;
  var esPatio  = l.patio;
  var bgColor  = esPatio ? "#fff8e1" : l.revisar ? "#fff7ed" : esGranel ? "#f0fdf4" : "white";
  var badgeHtml =
    (esPatio ?
      "<span style=\"background:#f59e0b;color:white;font-size:9px;font-weight:700;" +
      "padding:1px 6px;border-radius:8px;margin-left:6px\">PATIO</span>" :
     esGranel ?
      "<span style=\"background:#16a34a;color:white;font-size:9px;font-weight:700;" +
      "padding:1px 6px;border-radius:8px;margin-left:6px\">GRANEL</span>" : "") +
    (l.revisar ?
      "<span style=\"background:#ea580c;color:white;font-size:9px;font-weight:700;" +
      "padding:1px 6px;border-radius:8px;margin-left:6px\">REVISAR</span>" : "");
  return (
    "<div style=\"background:" + bgColor + ";border:1px solid var(--line);border-radius:10px;" +
    "padding:12px 14px;display:flex;align-items:flex-start;gap:10px\">" +
    "<div style=\"flex:1\">" +
    "<div style=\"font-size:12px;font-weight:700;font-family:monospace;color:var(--primary)\">" + l.cat + badgeHtml + "</div>" +
    "<div style=\"font-size:12px;color:var(--text);margin:2px 0\">" + l.desc + "</div>" +
    "<div style=\"font-size:11px;color:var(--muted)\">" +
    l.cant + " " + l.um + " &mdash; " + l.bultos + " " + l.tipoEmp +
    (l.tipoEmp === "Tarima" ? " (" + l.tarimaSubCant + " " + l.tarimaSubTipo + " c/ " + l.tarimaSubCont + " " + l.um + ")"
      : l.contEmp > 1 ? " de " + l.contEmp + " " + l.um : "") +
    (l.patio ? " &mdash; <b>PATIO</b>" : "") +
    (l.lote ? " &mdash; <b>L - " + l.lote + "</b>" : "") +
    "</div>" +
    "</div>" +
    "<button onclick=\"_guiasEditarLineaCompleta(" + idx + ")\" title=\"Editar\"" +
    " style=\"background:none;border:none;color:var(--primary);cursor:pointer;font-size:16px;" +
    "padding:0;line-height:1\">&#9998;</button>" +
    "<button onclick=\"_guiasEliminarLinea(" + idx + ")\" title=\"Borrar\"" +
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

var _guiasLotesSeleccionados = []; // [{lote, libre}] mientras se arma la selección de bobinas

function _guiasPedirLote(cat, desc, um, lotes){
  _guiasLotesSeleccionados = [];
  _guiasRenderModalLotes(cat, desc, um, lotes);
}

function _guiasRenderModalLotes(cat, desc, um, lotes){
  var catEsc  = cat.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var descEsc = desc.replace(/'/g,"&#39;").replace(/"/g,"&quot;");

  var lotesConExistencia = lotes.filter(function(l){ return (l.lib||0) > 0; });

  var filas = lotesConExistencia.map(function(l){
    var loteEsc = String(l.lote).replace(/'/g,"&#39;").replace(/"/g,"&quot;");
    var sel = _guiasLotesSeleccionados.some(function(s){ return s.lote === l.lote; });
    return "<button onclick=\"_guiasToggleLoteSel('" + catEsc + "','" + descEsc + "','" + um + "','" + loteEsc + "'," + (l.lib||0) + ")\"" +
      " style=\"text-align:left;padding:14px 16px;background:" + (sel ? "#d1fae5" : "#f0f4ff") + ";" +
      "border:1.5px solid " + (sel ? "#16a34a" : "var(--primary)") + ";border-radius:10px;cursor:pointer;" +
      "font-family:inherit;font-size:13px;line-height:1.5;position:relative\">" +
      (sel ? "<span style=\"position:absolute;top:10px;right:12px;color:#16a34a;font-size:16px;font-weight:700\">&#10003;</span>" : "") +
      "<b>Lote " + l.lote + "</b><br>" +
      "<span style=\"font-size:11.5px;color:var(--muted)\">" + nfmt(l.lib||0) + " " + um + " libres" +
      ((l.tras||0) > 0 ? " &middot; " + nfmt(l.tras) + " en traslado" : "") + "</span>" +
      "</button>";
  }).join("");

  var totalSel = _guiasLotesSeleccionados.reduce(function(a,s){ return a + s.libre; }, 0);
  var n = _guiasLotesSeleccionados.length;

  document.querySelector(".modal")?.remove();
  var modal = document.createElement("div");
  modal.className = "modal on";
  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:640px;padding:24px\">" +
    "<h3 style=\"margin:0 0 6px\">" + cat + "</h3>" +
    "<div style=\"font-size:13px;color:var(--muted);margin-bottom:18px\">" + desc + "</div>" +
    "<div style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:12px\">Este material lleva lote &mdash; elige de cuáles surtir (puedes elegir varias)</div>" +
    (lotesConExistencia.length > 0 ?
      "<div style=\"display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:12px;margin-bottom:18px\">" +
      filas + "</div>"
      : "<div style=\"font-size:12px;color:var(--muted);margin-bottom:18px\">No hay lotes con existencia disponible &mdash; usa el campo manual.</div>") +
    "<div style=\"border-top:1px dashed var(--line);padding-top:16px;margin-top:8px\">" +
    "<label style=\"font-size:11px;color:var(--muted)\">Lote manual (si no aparece en la lista)</label>" +
    "<div style=\"display:flex;gap:8px;margin-top:6px\">" +
    "<input id=\"gLoteManual\" type=\"text\" placeholder=\"Ej. L-4521\"" +
    " style=\"flex:1;padding:10px;border:1.5px solid var(--line);border-radius:8px;font-family:inherit\">" +
    "<input id=\"gLoteManualCant\" type=\"number\" min=\"1\" placeholder=\"Cantidad\"" +
    " style=\"width:110px;padding:10px;border:1.5px solid var(--line);border-radius:8px;font-family:inherit\">" +
    "<button class=\"btn-prim\" onclick=\"_guiasAgregarLoteManualASeleccion('" + catEsc + "','" + descEsc + "','" + um + "')\">+ Agregar</button>" +
    "</div></div>" +
    (n > 0 ?
      "<div style=\"background:#f0fdf4;border:1.5px solid #16a34a;border-radius:10px;padding:12px 16px;margin-top:16px\">" +
      "<div style=\"font-size:12px;font-weight:700;color:#16a34a\">" + n + " bobina" + (n>1?"s":"") +
      " seleccionada" + (n>1?"s":"") + " &mdash; " + nfmt(totalSel) + " " + um + " en total</div>" +
      "</div>" : "") +
    "<div style=\"display:flex;gap:8px;margin-top:18px\">" +
    "<button class=\"btn\" onclick=\"_guiasLotesSeleccionados=[];_guiasEditandoLineaIdx=null;this.closest('.modal').remove()\" style=\"flex:1;padding:10px\">Cancelar</button>" +
    (n > 0 ?
      "<button class=\"btn-prim\" onclick=\"_guiasConfirmarLotesSeleccionados('" + catEsc + "','" + descEsc + "','" + um + "')\" style=\"flex:1;padding:10px\">" +
      "Agregar " + n + " bobina" + (n>1?"s":"") + "</button>" : "") +
    "</div>" +
    "</div>";
  document.body.appendChild(modal);
}

function _guiasToggleLoteSel(cat, desc, um, lote, libre){
  var idx = _guiasLotesSeleccionados.findIndex(function(s){ return s.lote === lote; });
  if(idx >= 0) _guiasLotesSeleccionados.splice(idx, 1);
  else _guiasLotesSeleccionados.push({ lote: lote, libre: libre });
  _guiasRenderModalLotes(cat, desc, um, _guiasLotesReales(cat));
}

function _guiasAgregarLoteManualASeleccion(cat, desc, um){
  var lote = document.getElementById("gLoteManual").value.trim();
  var cant = parseInt(document.getElementById("gLoteManualCant").value || "0");
  if(!lote){ alert("Escribe el número de lote."); return; }
  if(!cant || cant < 1){ alert("Escribe la cantidad de ese lote."); return; }
  _guiasLotesSeleccionados.push({ lote: lote, libre: cant });
  _guiasRenderModalLotes(cat, desc, um, _guiasLotesReales(cat));
}

// Agrega una línea por cada bobina seleccionada, completa (sin recortar), de un solo golpe
function _guiasConfirmarLotesSeleccionados(cat, desc, um){
  var tipoLote = _guiasTipoForzado(cat) || "Caja";
  var seleccion = _guiasLotesSeleccionados.slice();
  _guiasLotesSeleccionados = [];
  document.querySelector(".modal")?.remove();

  if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }

  seleccion.forEach(function(s){
    _guiaActual.lineas.push({
      cat: cat, desc: desc, um: um,
      cant: s.libre, tipoEmp: tipoLote, contEmp: s.libre,
      bultos: 1, patio: false, granel: false, lote: s.lote
    });
  });

  _limpiarCatInput();
  _guiasRefrescarVista();
}

function _guiasPedirEmpaque(cat, desc, um, opciones){
  var modal = document.createElement("div");
  modal.className = "modal on";

  var catEsc  = cat.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var descEsc = desc.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var esPatio = (mat(cat).ubic || "").toLowerCase() === "patio";

  // Si es patio — preguntar solo cantidad y agregar directo
  if(esPatio){
    var cant = prompt("Cantidad de " + cat + " (va a PATIO):", "");
    if(!cant || isNaN(cant) || parseInt(cant) < 1){ _guiasEditandoLineaIdx = null; return; }
    cant = parseInt(cant);
    if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }
    _guiaActual.lineas.push({
      cat: cat, desc: desc, um: um,
      cant: cant, tipoEmp: "Patio", contEmp: 1,
      bultos: cant, patio: true, granel: false
    });
    var inp = document.getElementById("gCatInput");
    if(inp) _limpiarCatInput();
    _guiasRefrescarVista();
    return;
  }

  // Si el material lleva lote real — se maneja aparte en el selector múltiple de bobinas
  var lotesReales = _guiasLotesReales(cat);
  if(lotesReales.length > 0){
    _guiasPedirLote(cat, desc, um, lotesReales);
    return;
  }
  var tipoForzado = _guiasTipoForzado(cat);

  // Botones de empaques conocidos
  var btnsConocidos = "";
  for(var i=0; i<opciones.length; i++){
    var op = opciones[i];
    var estrella = op.preferido ? "&#9733;" : "&#9734;"; // ★ / ☆
    btnsConocidos +=
      "<div style=\"display:flex;align-items:center;gap:4px;margin-bottom:6px\">" +
      "<button class=\"gEmpBtn\" onclick=\"_guiasSeleccionarEmpaque('" + catEsc + "','" + descEsc + "','" + um + "','" +
      op.tipo + "'," + op.cont + ")\"" +
      " style=\"flex:1;text-align:left;padding:10px 14px;background:#f0f4ff;" +
      "border:1.5px solid var(--primary);border-radius:8px;cursor:pointer;" +
      "font-family:inherit;font-size:13px\">" +
      "<b>" + op.tipo + "</b> de " + op.cont + " " + um +
      " <span style=\"color:var(--muted);font-size:11px\">(x" + (op.freq||1) + ")</span>" +
      (op.preferido ? " <span style=\"color:#f59e0b;font-size:10px;font-weight:700\">PREFERIDO</span>" : "") +
      "</button>" +
      "<button onclick=\"_guiasTogglePreferido('" + catEsc + "','" + descEsc + "','" + um + "'," + op.cont + ")\"" +
      " title=\"" + (op.preferido ? "Preferido — siempre se propondrá primero" : "Marcar como preferido") + "\"" +
      " style=\"background:none;border:none;cursor:pointer;font-size:20px;line-height:1;" +
      "color:" + (op.preferido ? "#f59e0b" : "#ccc") + ";flex-shrink:0\">" + estrella + "</button>" +
      "<button onclick=\"_guiasBDBorrarOpcion('" + catEsc + "','" + descEsc + "','" + um + "'," + op.cont + ")\"" +
      " title=\"Borrar este empaque de la memoria\"" +
      " style=\"background:none;border:none;cursor:pointer;font-size:18px;line-height:1;" +
      "color:#dc2626;flex-shrink:0\">&times;</button>" +
      "</div>";
  }

  // Botones de tarimas conocidas para este catálogo
  var tarimasCat = _guiasTarimaOpciones(cat);
  var btnsTarimas = "";
  for(var ti=0; ti<tarimasCat.length; ti++){
    var tp = tarimasCat[ti];
    var estrellaT = tp.preferido ? "&#9733;" : "&#9734;";
    btnsTarimas +=
      "<div style=\"display:flex;align-items:center;gap:4px;margin-bottom:6px\">" +
      "<button class=\"gEmpBtn\" onclick=\"_guiasSeleccionarTarima('" + catEsc + "','" + descEsc + "','" + um + "','" +
      tp.subTipo + "'," + tp.subCant + "," + tp.subCont + ")\"" +
      " style=\"flex:1;text-align:left;padding:10px 14px;background:#fdf4e7;" +
      "border:1.5px solid #b8722a;border-radius:8px;cursor:pointer;" +
      "font-family:inherit;font-size:13px\">" +
      "<b>Tarima</b> de " + tp.subCant + " " + tp.subTipo + " c/ " + tp.subCont + " " + um +
      (tp.preferido ? " <span style=\"color:#f59e0b;font-size:10px;font-weight:700\">PREFERIDA</span>" : "") +
      "</button>" +
      "<button onclick=\"_guiasTarimaTogglePreferido('" + catEsc + "','" + descEsc + "','" + um + "','" + tp.subTipo + "'," + tp.subCant + "," + tp.subCont + ")\"" +
      " title=\"" + (tp.preferido ? "Preferida — siempre se propondrá primero" : "Marcar como preferida") + "\"" +
      " style=\"background:none;border:none;cursor:pointer;font-size:20px;line-height:1;" +
      "color:" + (tp.preferido ? "#f59e0b" : "#ccc") + ";flex-shrink:0\">" + estrellaT + "</button>" +
      "<button onclick=\"_guiasTarimaBorrarOpcion('" + catEsc + "','" + descEsc + "','" + um + "','" + tp.subTipo + "'," + tp.subCant + "," + tp.subCont + ")\"" +
      " title=\"Borrar esta tarima de la memoria\"" +
      " style=\"background:none;border:none;cursor:pointer;font-size:18px;line-height:1;" +
      "color:#dc2626;flex-shrink:0\">&times;</button>" +
      "</div>";
  }

  // Columna izquierda — conocidos + tarimas + granel
  var izqHtml =
    "<div style=\"flex:1;min-width:0;padding:16px\">" +
    (opciones.length > 0 ?
      "<div style=\"font-size:11px;font-weight:700;color:var(--primary);text-transform:uppercase;" +
      "letter-spacing:.4px;margin-bottom:10px\">&#10003; Conocidos</div>" +
      btnsConocidos +
      "<div style=\"border-top:1px dashed var(--line);margin-top:12px;padding-top:12px\">" : "") +
    (tarimasCat.length > 0 ?
      "<div style=\"font-size:11px;font-weight:700;color:#b8722a;text-transform:uppercase;" +
      "letter-spacing:.4px;margin-bottom:10px\">&#128717; Tarimas conocidas</div>" +
      btnsTarimas +
      "<div style=\"border-top:1px dashed var(--line);margin-top:12px;padding-top:12px\">" : "") +
    "<div style=\"font-size:11px;font-weight:700;color:#16a34a;text-transform:uppercase;" +
    "letter-spacing:.4px;margin-bottom:6px\">&#9744; A granel</div>" +
    "<p style=\"font-size:11px;color:var(--muted);margin:0 0 8px\">Va a la caja colectiva del área</p>" +
    "<div style=\"display:flex;gap:6px\">" +
    "<input id=\"gGranelCant\" type=\"number\" min=\"1\" placeholder=\"Cantidad\"" +
    " onkeydown=\"if(event.key==='Enter'){event.preventDefault();_guiasAGranelModal('" + catEsc + "','" + descEsc + "','" + um + "');}\"" +
    " style=\"flex:1;padding:8px;border:1.5px solid #16a34a;border-radius:8px;font-family:inherit\"/>" +
    "<button onclick=\"_guiasAGranelModal('" + catEsc + "','" + descEsc + "','" + um + "')\"" +
    " style=\"padding:8px 14px;background:#16a34a;color:white;border:none;border-radius:8px;" +
    "font-weight:700;cursor:pointer;font-family:inherit\">Granel</button>" +
    "</div>" +
    (opciones.length > 0 ? "</div>" : "") +
    (tarimasCat.length > 0 ? "</div>" : "") +
    "</div>";

  // Columna derecha — nuevo empaque + nueva tarima (desplegables)
  var tipoInputHtml = tipoForzado
    ? "<input id=\"gEmpTipo\" type=\"text\" value=\"" + tipoForzado + "\" readonly" +
      " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
      "font-family:inherit;margin-top:3px;background:#f3f4f6;color:var(--muted)\">"
    : "<input id=\"gEmpTipo\" type=\"text\" value=\"Caja\" placeholder=\"Caja, Costal...\"" +
      " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
      "font-family:inherit;margin-top:3px\">";
  var tipoTarimaInputHtml = tipoForzado
    ? "<input id=\"gTarimaTipo\" type=\"text\" value=\"" + tipoForzado + "\" readonly" +
      " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
      "font-family:inherit;margin-top:3px;background:#f3f4f6;color:var(--muted)\">"
    : "<input id=\"gTarimaTipo\" type=\"text\" value=\"Caja\" placeholder=\"Caja, Bobina, Rollo...\"" +
      " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
      "font-family:inherit;margin-top:3px\">";
  var derechaHtml =
    "<div style=\"flex:1;min-width:0;padding:16px;border-left:1px solid var(--line)\">" +
    "<details" + (opciones.length === 0 ? " open" : "") + ">" +
    "<summary style=\"font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;cursor:pointer;padding:4px 0;list-style:none\">" +
    (opciones.length > 0 ? "+ Nuevo empaque diferente" : "Captura el empaque") +
    "</summary>" +
    "<div style=\"margin-top:10px;display:flex;flex-direction:column;gap:8px\">" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Tipo de empaque</label>" +
    tipoInputHtml + "</div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Contenido por empaque</label>" +
    "<input id=\"gEmpCont\" type=\"number\" min=\"1\" value=\"1\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;margin-top:3px\"></div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Cantidad total</label>" +
    "<input id=\"gEmpCant\" type=\"number\" min=\"1\" value=\"\"" +
    " onkeydown=\"if(event.key==='Enter'){event.preventDefault();_guiasConfirmarEmpaque('" + catEsc + "','" + descEsc + "','" + um + "');}\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:15px;font-weight:700;color:var(--primary);margin-top:3px\"></div>" +
    "</div></details>" +
    "<details style=\"margin-top:14px;border-top:1px dashed var(--line);padding-top:12px\">" +
    "<summary style=\"font-size:11px;font-weight:700;color:#b8722a;text-transform:uppercase;" +
    "letter-spacing:.4px;cursor:pointer;padding:4px 0;list-style:none\">+ Nueva tarima</summary>" +
    "<div style=\"margin-top:10px;display:flex;flex-direction:column;gap:8px\">" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Tipo de contenedor por tarima</label>" +
    tipoTarimaInputHtml + "</div>" +
    "<div style=\"display:flex;gap:8px\">" +
    "<div style=\"flex:1\"><label style=\"font-size:11px;color:var(--muted)\">Contenedores por tarima</label>" +
    "<input id=\"gTarimaSubCant\" type=\"number\" min=\"1\" placeholder=\"Ej. 60\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;margin-top:3px\"></div>" +
    "<div style=\"flex:1\"><label style=\"font-size:11px;color:var(--muted)\">Contenido c/u (" + um + ")</label>" +
    "<input id=\"gTarimaSubCont\" type=\"number\" min=\"1\" placeholder=\"Ej. 500\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;margin-top:3px\"></div>" +
    "</div>" +
    "<div><label style=\"font-size:11px;color:var(--muted)\">Cuántas tarimas van</label>" +
    "<input id=\"gTarimaNum\" type=\"number\" min=\"1\" value=\"1\"" +
    " onkeydown=\"if(event.key==='Enter'){event.preventDefault();_guiasConfirmarTarimaNueva('" + catEsc + "','" + descEsc + "','" + um + "');}\"" +
    " style=\"width:100%;padding:8px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:15px;font-weight:700;color:#b8722a;margin-top:3px\"></div>" +
    "<button onclick=\"_guiasConfirmarTarimaNueva('" + catEsc + "','" + descEsc + "','" + um + "')\"" +
    " style=\"width:100%;padding:9px;background:#b8722a;color:white;border:none;border-radius:8px;" +
    "font-weight:700;cursor:pointer;font-family:inherit;font-size:12.5px\">Agregar tarima</button>" +
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
    "<button class=\"btn\" onclick=\"_guiasEditandoLineaIdx=null;this.closest('.modal').remove()\">Cancelar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarEmpaque('" + catEsc + "','" + descEsc + "','" + um + "')\">Agregar empaque</button>" +
    "</div></div>";

  document.body.appendChild(modal);
  // Navegación con flechas entre los empaques/tarimas conocidas (sin tocar el mouse):
  // ↓/↑ mueve el foco solo entre los botones de selección (no las estrellas ni la ×).
  modal.addEventListener("keydown", function(e){
    if(e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    var btns = [].slice.call(modal.querySelectorAll(".gEmpBtn"));
    if(!btns.length) return;
    e.preventDefault();
    var idx = btns.indexOf(document.activeElement);
    if(idx === -1){ btns[0].focus(); return; }
    var next = e.key === "ArrowDown" ? (idx + 1) % btns.length : (idx - 1 + btns.length) % btns.length;
    btns[next].focus();
  });
  setTimeout(function(){
    // Si hay algo conocido, el foco arranca ahí (el primero es siempre el preferido/más usado)
    // para poder confirmar con solo Enter; si no, cae al campo de captura manual.
    var primero = modal.querySelector(".gEmpBtn");
    var el = primero || document.getElementById(opciones.length === 0 ? "gEmpCant" : "gGranelCant");
    if(el) el.focus();
  }, 100);
}

function _guiasSeleccionarEmpaque(cat, desc, um, tipo, cont){
  cont = parseInt(cont);
  var catEsc  = cat.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var descEsc = desc.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var tipoEsc = String(tipo).replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var box = document.querySelector(".modal .modal-box");
  if(!box) return;
  // Paso 2 dentro del mismo modal (sin prompt()): solo pide la cantidad total, con Enter para confirmar.
  box.innerHTML =
    "<div style=\"padding:24px\">" +
    "<h3 style=\"margin:0 0 4px\">" + cat + "</h3>" +
    "<div style=\"font-size:12px;color:var(--muted);margin-bottom:16px\">" + desc + "</div>" +
    "<div style=\"background:#f0f4ff;border:1.5px solid var(--primary);border-radius:10px;padding:12px 14px;margin-bottom:16px;font-size:13px\">" +
    "<b>" + tipo + "</b> de " + cont + " " + um + "</div>" +
    "<label style=\"font-size:11px;color:var(--muted)\">Cantidad total</label>" +
    "<input id=\"gEmpCantInline\" type=\"number\" min=\"1\" autocomplete=\"off\"" +
    " onkeydown=\"if(event.key==='Enter'){event.preventDefault();_guiasConfirmarCantidadInline('" + catEsc + "','" + descEsc + "','" + um + "','" + tipoEsc + "'," + cont + ");}\"" +
    " style=\"width:100%;padding:10px;border:1.5px solid var(--primary);border-radius:8px;" +
    "font-family:inherit;font-size:17px;font-weight:700;color:var(--primary);margin-top:4px\">" +
    "<div style=\"display:flex;gap:8px;margin-top:18px\">" +
    "<button class=\"btn\" onclick=\"_guiasPedirEmpaque('" + catEsc + "','" + descEsc + "','" + um + "',_guiasBDOpciones('" + catEsc + "'))\" style=\"flex:1;padding:10px\">&larr; Regresar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarCantidadInline('" + catEsc + "','" + descEsc + "','" + um + "','" + tipoEsc + "'," + cont + ")\" style=\"flex:1;padding:10px\">Agregar</button>" +
    "</div></div>";
  setTimeout(function(){ document.getElementById("gEmpCantInline")?.focus(); }, 50);
}

function _guiasSeleccionarTarima(cat, desc, um, subTipo, subCant, subCont){
  var catEsc  = cat.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var descEsc = desc.replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var subTipoEsc = String(subTipo).replace(/'/g,"&#39;").replace(/"/g,"&quot;");
  var box = document.querySelector(".modal .modal-box");
  if(!box) return;
  box.innerHTML =
    "<div style=\"padding:24px\">" +
    "<h3 style=\"margin:0 0 4px\">" + cat + "</h3>" +
    "<div style=\"font-size:12px;color:var(--muted);margin-bottom:16px\">" + desc + "</div>" +
    "<div style=\"background:#fdf4e7;border:1.5px solid #b8722a;border-radius:10px;padding:12px 14px;margin-bottom:16px;font-size:13px\">" +
    "<b>Tarima</b> de " + subCant + " " + subTipo + " c/ " + subCont + " " + um + "</div>" +
    "<label style=\"font-size:11px;color:var(--muted)\">¿Cuántas tarimas van?</label>" +
    "<input id=\"gTarimaNumInline\" type=\"number\" min=\"1\" value=\"1\" autocomplete=\"off\"" +
    " onkeydown=\"if(event.key==='Enter'){event.preventDefault();_guiasConfirmarTarimasInline('" + catEsc + "','" + descEsc + "','" + um + "','" + subTipoEsc + "'," + subCant + "," + subCont + ");}\"" +
    " style=\"width:100%;padding:10px;border:1.5px solid #b8722a;border-radius:8px;" +
    "font-family:inherit;font-size:17px;font-weight:700;color:#b8722a;margin-top:4px\">" +
    "<div style=\"display:flex;gap:8px;margin-top:18px\">" +
    "<button class=\"btn\" onclick=\"_guiasPedirEmpaque('" + catEsc + "','" + descEsc + "','" + um + "',_guiasBDOpciones('" + catEsc + "'))\" style=\"flex:1;padding:10px\">&larr; Regresar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarTarimasInline('" + catEsc + "','" + descEsc + "','" + um + "','" + subTipoEsc + "'," + subCant + "," + subCont + ")\" style=\"flex:1;padding:10px\">Agregar</button>" +
    "</div></div>";
  setTimeout(function(){ document.getElementById("gTarimaNumInline")?.focus(); }, 50);
}

// Confirma el paso 2 (cantidad) de un empaque conocido — sin prompt(), dentro del mismo modal, con Enter.
function _guiasConfirmarCantidadInline(cat, desc, um, tipo, cont){
  var cant = parseInt(document.getElementById("gEmpCantInline")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa una cantidad válida."); return; }
  document.querySelector(".modal")?.remove();
  var nCajas = Math.floor(cant / cont), residuo = cant % cont;
  if(nCajas > 0 && residuo > 0){
    _guiasAgregarLineaSilente(cat, desc, um, tipo, cont, nCajas * cont, false);
    _guiasAgregarLineaSilente(cat, desc, um, "Granel", 0, residuo, true);
    _limpiarCatInput();
    _guiasRefrescarVista();
  } else {
    _guiasAgregarLinea(cat, desc, um, tipo, cont, cant, nCajas === 0);
  }
  _guiasBDActualizarEmpaque(cat, tipo, cont, um);
}
// Confirma el paso 2 (número de tarimas) de una tarima conocida — mismo criterio, sin prompt().
function _guiasConfirmarTarimasInline(cat, desc, um, subTipo, subCant, subCont){
  var num = parseInt(document.getElementById("gTarimaNumInline")?.value || "0");
  if(!num || num < 1){ alert("Ingresa cuántas tarimas van."); return; }
  document.querySelector(".modal")?.remove();
  _guiasTarimaBDActualizar(cat, subTipo, subCant, subCont, um);
  _guiasAgregarLineaTarima(cat, desc, um, subTipo, subCant, subCont, num);
}

function _guiasAPatioModal(cat, desc, um){
  var cant = parseInt(document.getElementById("gPatioCant")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa la cantidad de patio."); return; }
  document.querySelector(".modal")?.remove();
  var bultos = cant;
  if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }
  _guiaActual.lineas.push({
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: "Patio", contEmp: 1,
    bultos: bultos, patio: true, granel: false
  });
  _limpiarCatInput();
  _guiasRefrescarVista();
}

function _guiasAGranelModal(cat, desc, um){
  var cant = parseInt(document.getElementById("gGranelCant")?.value || "0");
  if(!cant || cant < 1){ alert("Ingresa la cantidad."); return; }
  document.querySelector(".modal")?.remove();
  _guiasAgregarLinea(cat, desc, um, "Granel", 0, cant, true);
}

function _guiasAGranel(cat, desc, um){
  var cant = prompt("Cantidad de " + cat + " que va a granel:", "");
  if(!cant || isNaN(cant) || parseInt(cant) < 1){ _guiasEditandoLineaIdx = null; return; }
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
    // La línea vieja, si aplica, se limpia dentro de _guiasAgregarLineaSilente
    _guiasAgregarLineaSilente(cat, desc, um, tipo, cont, nCajas * cont, false);
    _guiasAgregarLineaSilente(cat, desc, um, "Granel", 0, residuo, true);
      _limpiarCatInput();
    _guiasRefrescarVista();
  } else {
    var esGranel = (nCajas === 0 && cont > 1);
    _guiasAgregarLinea(cat, desc, um, tipo, cont, cant, esGranel);
  }
  _guiasBDActualizarEmpaque(cat, tipo, cont, um);
}

function _limpiarCatInput(){
  var inp = document.getElementById("gCatInput");
  if(inp){
    inp.value = "";
    document.getElementById("gCatInfo").textContent = "";
    // Regresa el cursor al catálogo automáticamente — así se puede seguir capturando
    // solo con teclado, sin usar el mouse para volver a dar clic ahí cada vez.
    setTimeout(function(){ inp.focus(); }, 30);
  }
}

function _guiasAgregarLineaSilente(cat, desc, um, tipoEmp, contEmp, cant, esGranel){
  // Agrega sin limpiar input ni refrescar (para llamadas múltiples)
  if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }
  var bultos = contEmp > 1 ? Math.floor(cant / contEmp) : cant;
  var linea = {
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: tipoEmp, contEmp: contEmp,
    bultos: bultos, patio: false, granel: esGranel || false
  };
  _guiaActual.lineas.push(linea);
}

function _guiasAgregarLinea(cat, desc, um, tipoEmp, contEmp, cant, esGranel){
  if(_guiasEditandoLineaIdx !== null){ _guiaActual.lineas.splice(_guiasEditandoLineaIdx, 1); _guiasEditandoLineaIdx = null; }
  var bultos = contEmp > 1 ? Math.ceil(cant / contEmp) : cant;
  var linea = {
    cat: cat, desc: desc, um: um,
    cant: cant, tipoEmp: tipoEmp, contEmp: contEmp,
    bultos: bultos, patio: false, granel: esGranel || false
  };
  _guiaActual.lineas.push(linea);
  // Limpiar input y cerrar modal
  _limpiarCatInput();
  document.querySelector(".modal")?.remove();
  _guiasRefrescarVista();
}

function _guiasEliminarLinea(idx){
  _guiasCapturarFirmasRevisionSiActiva();
  _guiaActual.lineas.splice(idx, 1);
  _guiasRefrescarVista();
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

function _guiasCargarORDesdeNueva(){
  document.getElementById("gORFileNueva")?.click();
}

function _guiasProcesarORDesdeNueva(input){
  var file = input.files[0];
  if(!file) return;
  input.value = "";
  var reader = new FileReader();
  reader.onload = function(e){
    _guiasLeerFilasOR(e.target.result, function(rows, hdrIdx, iCat, iXS){
      var hdr = _guiasParsearHeaderOR(rows, hdrIdx);
      var detectado = [];

      if(hdr.area){
        var selArea = document.getElementById("gArea");
        if(selArea){ selArea.value = hdr.area; detectado.push("Área: " + hdr.area); }
      }
      if(hdr.destino){
        var inpDestino = document.getElementById("gDestino");
        if(inpDestino){ inpDestino.value = hdr.destino; _guiasActualizarDestinatario(); detectado.push("Destino: " + hdr.destino); }
      }
      if(hdr.fecha){
        var inpFecha = document.getElementById("gFecha");
        if(inpFecha){ inpFecha.value = hdr.fecha; detectado.push("Fecha: " + hdr.fecha); }
      }

      // Guardar filas de materiales para aplicarlas en cuanto se cree la guía (falta el folio)
      _guiasORPendiente = { rows: rows, hdrIdx: hdrIdx, iCat: iCat, iXS: iXS };
      _guiasToast("Archivo procesado y guardado &mdash; se importará al continuar", "ok");

      var msg = detectado.length ? detectado.join("\n") : "No se detectaron Área/Destino/Fecha en el archivo — verifica esos campos manualmente.";
      alert("Datos detectados del OR:\n\n" + msg + "\n\nLos materiales se importarán al continuar. Completa el número de guía (folio) y continúa.");
    });
  };
  reader.readAsArrayBuffer(file);
}

// Lee el archivo y regresa {rows, hdrIdx, iCat, iXS} o null (con alert) si no se pudo interpretar
// Detecta la fila de encabezados y las columnas de Catálogo/Cantidad en un arreglo de filas (rows).
// Se usa tanto para el Excel de OR como para tablas pegadas desde correo.
function _guiasDetectarHeaderYColumnas(rows){
  var hdrIdx = -1;
  for(var i=0; i<rows.length; i++){
    if(rows[i].some(function(c){
      var u = String(c||"").toUpperCase();
      return u.includes("CATALOGO") || u.includes("CATÁLOGO") || u.trim() === "MATERIAL";
    })){
      hdrIdx = i; break;
    }
  }
  if(hdrIdx < 0) return _guiasDetectarSinHeader(rows);
  var hdrs = rows[hdrIdx].map(function(h){ return String(h||"").toLowerCase().trim(); });
  var iCat = hdrs.findIndex(function(h){ return h.includes("cat") || h === "material"; });
  if(iCat < 0) return _guiasDetectarSinHeader(rows);
  var iXS  = hdrs.findIndex(function(h){
    return h.includes("surtir") || h.includes("cantidad") || h.includes("piezas") ||
           h.includes("pzas") || h.includes("pza") || /^cant\.?$/.test(h) || h.includes("qty");
  });
  if(iXS < 0){
    // Sin encabezado reconocible de cantidad (p.ej. el jefe pone el código del almacén destino,
    // como "TX8I") — por convención esa columna suele ser la última de la tabla.
    var iDesc = hdrs.findIndex(function(h){ return h.includes("desc"); });
    var candidatos = [];
    for(var c=0; c<hdrs.length; c++){ if(c!==iCat && c!==iDesc && hdrs[c]!=="") candidatos.push(c); }
    if(candidatos.length > 0) iXS = candidatos[candidatos.length-1];
  }
  if(iXS < 0) return null;
  return { hdrIdx: hdrIdx, iCat: iCat, iXS: iXS };
}

// Cuando no hay ninguna fila de encabezado reconocible (el jefe no puso títulos a las columnas):
// se asume que la primera columna es el catálogo, y se busca cuál otra columna es "casi toda numérica"
// para tratarla como la cantidad.
function _guiasDetectarSinHeader(rows){
  if(!rows || rows.length === 0) return null;
  var ncols = Math.max.apply(null, rows.map(function(r){ return r.length; }));
  if(ncols < 2) return null;
  var muestra = rows.slice(0, Math.min(rows.length, 20));

  // La columna 0 debe parecer catálogo: casi todas las filas con algo numérico (típico de claves SAP)
  var col0Num = 0, col0Total = 0;
  muestra.forEach(function(r){
    var v = String(r[0]||"").trim();
    if(!v) return;
    col0Total++;
    if(/^\d+$/.test(v.replace(/[,\s]/g,""))) col0Num++;
  });
  if(col0Total === 0 || (col0Num/col0Total) < 0.7) return null; // la col. 0 no parece catálogo — mejor no adivinar

  var mejorCol = -1, mejorScore = -1;
  for(var c=1; c<ncols; c++){
    var numericos = 0, total = 0;
    muestra.forEach(function(r){
      var v = String(r[c]||"").trim();
      if(!v) return;
      total++;
      // Acepta "4000", "4,000" o "4,000 PZAS"/"100 M" (número al inicio, seguido opcionalmente de texto/unidad)
      if(/^\d[\d,.\s]*(\s+[a-záéíóúñ.]+)?$/i.test(v)) numericos++;
    });
    var score = total > 0 ? numericos/total : 0;
    if(score > mejorScore){ mejorScore = score; mejorCol = c; }
  }
  if(mejorCol < 0 || mejorScore < 0.6) return null; // ninguna columna se ve suficientemente numérica

  return { hdrIdx: -1, iCat: 0, iXS: mejorCol };
}

// ── Pegar tabla copiada de un correo ──────────────────────────────────────────
function _guiasAbrirPegarTabla(desdeNueva){
  var modal = document.createElement("div");
  modal.className = "modal on";
  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:640px\">" +
    "<h3 style=\"margin:0 0 6px\">Pegar tabla del correo</h3>" +
    "<p style=\"font-size:12px;color:var(--muted);margin:0 0 12px\">" +
    "Copia la tabla completa desde el correo (incluyendo el encabezado con Catálogo y Cantidad), " +
    "y pégala aquí abajo con Ctrl+V.</p>" +
    "<div id=\"gPasteArea\" contenteditable=\"true\"" +
    " style=\"min-height:160px;max-height:320px;overflow:auto;border:1.5px dashed var(--line);" +
    "border-radius:8px;padding:12px;font-size:12px\"></div>" +
    "<div style=\"display:flex;justify-content:flex-end;gap:8px;margin-top:14px\">" +
    "<button class=\"btn\" onclick=\"this.closest('.modal').remove()\">Cancelar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasProcesarTablaPegada(" + (desdeNueva?"true":"false") + ")\">Procesar tabla</button>" +
    "</div></div>";
  document.body.appendChild(modal);
  setTimeout(function(){ document.getElementById("gPasteArea")?.focus(); }, 50);
}

// Aviso visual (toast) que confirma claramente que algo se procesó/guardó — no bloquea como alert()
function _guiasToast(mensaje, tipo){
  document.querySelectorAll(".guias-toast").forEach(function(t){ t.remove(); }); // solo uno a la vez
  var colores = {
    ok:  { bg: "#16a34a", icono: "&#10003;" },
    warn:{ bg: "#ea580c", icono: "&#9888;"  }
  };
  var c = colores[tipo] || colores.ok;
  var toast = document.createElement("div");
  toast.className = "guias-toast";
  toast.style.cssText =
    "position:fixed;top:16px;left:50%;transform:translateX(-50%);z-index:9999;" +
    "background:" + c.bg + ";color:white;padding:14px 22px;border-radius:12px;" +
    "font-family:inherit;font-size:14px;font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,.25);" +
    "display:flex;align-items:center;gap:10px;max-width:90vw;text-align:left;" +
    "animation:guiasToastIn .25s ease-out";
  toast.innerHTML = "<span style=\"font-size:18px\">" + c.icono + "</span><span>" + mensaje + "</span>";
  if(!document.getElementById("guiasToastStyle")){
    var style = document.createElement("style");
    style.id = "guiasToastStyle";
    style.textContent = "@keyframes guiasToastIn{from{opacity:0;transform:translate(-50%,-12px)}to{opacity:1;transform:translate(-50%,0)}}";
    document.head.appendChild(style);
  }
  document.body.appendChild(toast);
  setTimeout(function(){
    toast.style.transition = "opacity .3s";
    toast.style.opacity = "0";
    setTimeout(function(){ toast.remove(); }, 300);
  }, 3800);
}

function _guiasProcesarTablaPegada(desdeNueva){
  var area = document.getElementById("gPasteArea");
  if(!area) return;
  var rows = [];
  var tabla = area.querySelector("table");
  if(tabla){
    tabla.querySelectorAll("tr").forEach(function(tr){
      var celdas = Array.prototype.map.call(tr.querySelectorAll("td,th"), function(td){ return td.innerText.trim(); });
      if(celdas.length) rows.push(celdas);
    });
  }
  if(rows.length < 2){
    // Sin tabla HTML (o vacía) — intentar como texto plano tabulado / con columnas alineadas por espacios
    rows = [];
    var texto = area.innerText || "";
    texto.split(/\n/).forEach(function(linea){
      linea = linea.replace(/\s+$/,"");
      if(!linea.trim()) return;
      var celdas = linea.indexOf("\t") >= 0 ? linea.split("\t") : linea.split(/\s{2,}/);
      rows.push(celdas.map(function(c){ return c.trim(); }));
    });
  }
  if(rows.length < 2){
    alert("No se detectó una tabla válida. Copia la tabla completa del correo, incluyendo los encabezados.");
    return;
  }

  var det = _guiasDetectarHeaderYColumnas(rows);
  if(!det){
    alert("No pude identificar las columnas de Catálogo y Cantidad en lo que pegaste. " +
      "Verifica que la tabla tenga encabezados como \"Catálogo\" y \"Cantidad\" (o \"X Surtir\", \"Piezas\", etc.).");
    return;
  }

  // En vez de importar directo, se normaliza a pares [catálogo, cantidad] y se muestra
  // una tabla editable — así se pueden quitar líneas que no surten, corregir cantidades,
  // o agregar líneas sueltas antes de confirmar la importación.
  var pares = [];
  for(var r=det.hdrIdx+1; r<rows.length; r++){
    var row = rows[r];
    var cat = String(row[det.iCat]||"").trim();
    var xs  = String(row[det.iXS]||"").replace(/[,\s]/g,"").trim();
    if(!cat && !xs) continue;
    pares.push([cat, xs]);
  }

  document.querySelector(".modal")?.remove();
  _guiasAbrirRevisionTabla(pares, desdeNueva);
}

// ── Fila editable dentro de la tabla de revisión ─────────────────────────────
function _guiasTplFilaPegRevision(cat, xs){
  return "<div class=\"gPegRow\" style=\"display:flex;gap:8px;align-items:center;margin-bottom:6px\">" +
    "<input type=\"text\" class=\"gPegCat\" value=\"" + _escAttr(cat) + "\" placeholder=\"Catálogo\"" +
    " style=\"flex:1.4;padding:7px 9px;border:1px solid var(--line);border-radius:6px;font-family:inherit;font-size:13px\">" +
    "<input type=\"text\" inputmode=\"numeric\" class=\"gPegXS\" value=\"" + _escAttr(xs) + "\" placeholder=\"Cantidad\"" +
    " style=\"width:90px;padding:7px 9px;border:1px solid var(--line);border-radius:6px;font-family:inherit;font-size:13px;text-align:right\">" +
    "<button type=\"button\" onclick=\"this.closest('.gPegRow').remove()\" title=\"Quitar línea\"" +
    " style=\"border:none;background:none;color:var(--rojo,#dc2626);font-size:18px;line-height:1;cursor:pointer;padding:4px 8px;font-weight:700\">&times;</button>" +
    "</div>";
}

// ── Modal de revisión: tabla editable antes de importar ──────────────────────
function _guiasAbrirRevisionTabla(pares, desdeNueva){
  var filasHtml = pares.map(function(p){ return _guiasTplFilaPegRevision(p[0], p[1]); }).join("");
  var modal = document.createElement("div");
  modal.className = "modal on";
  modal.id = "gModalRevisionTabla";
  modal.innerHTML =
    "<div class=\"modal-box\" style=\"max-width:520px\">" +
    "<h3 style=\"margin:0 0 6px\">Revisar antes de importar</h3>" +
    "<p style=\"font-size:12px;color:var(--muted);margin:0 0 12px\">" +
    "Quita las líneas que no surten (&times;), corrige catálogo o cantidad, o agrega líneas sueltas. " +
    "Al confirmar se procesan igual que siempre (empaque, patio, lote, etc.).</p>" +
    "<div id=\"gPegRowsWrap\" style=\"max-height:340px;overflow-y:auto;padding-right:2px\">" +
    filasHtml +
    "</div>" +
    "<button type=\"button\" onclick=\"_guiasAgregarFilaPegRevision()\"" +
    " style=\"margin-top:8px;border:1.5px dashed var(--line);background:none;border-radius:6px;" +
    "padding:7px 12px;font-family:inherit;font-size:12px;color:var(--primary);cursor:pointer;width:100%\">" +
    "+ Agregar línea</button>" +
    "<div style=\"display:flex;justify-content:flex-end;gap:8px;margin-top:14px\">" +
    "<button class=\"btn\" onclick=\"this.closest('.modal').remove()\">Cancelar</button>" +
    "<button class=\"btn-prim\" onclick=\"_guiasConfirmarRevisionTabla(" + (desdeNueva?"true":"false") + ")\">Confirmar e importar</button>" +
    "</div></div>";
  document.body.appendChild(modal);
}

function _guiasAgregarFilaPegRevision(){
  var wrap = document.getElementById("gPegRowsWrap");
  if(!wrap) return;
  wrap.insertAdjacentHTML("beforeend", _guiasTplFilaPegRevision("", ""));
  var inputs = wrap.querySelectorAll(".gPegRow .gPegCat");
  var ultimo = inputs[inputs.length-1];
  if(ultimo) ultimo.focus();
}

function _guiasConfirmarRevisionTabla(desdeNueva){
  var wrap = document.getElementById("gPegRowsWrap");
  if(!wrap) return;
  var rowsEditadas = [];
  wrap.querySelectorAll(".gPegRow").forEach(function(fila){
    var cat = (fila.querySelector(".gPegCat")?.value || "").trim();
    var xs  = (fila.querySelector(".gPegXS")?.value || "").trim();
    if(!cat) return; // línea vacía o dejada en blanco tras editar — se ignora
    rowsEditadas.push([cat, xs]);
  });
  if(rowsEditadas.length === 0){
    alert("No quedó ninguna línea con catálogo. Agrega al menos una o cancela.");
    return;
  }

  document.getElementById("gModalRevisionTabla")?.remove();

  // hdrIdx:-1 porque rowsEditadas ya no trae encabezado — arranca en la fila 0
  if(desdeNueva){
    _guiasORPendiente = { rows: rowsEditadas, hdrIdx: -1, iCat: 0, iXS: 1 };
    _guiasToast("Tabla revisada y guardada &mdash; se importará al continuar", "ok");
  } else {
    var res = _guiasImportarFilasOR(rowsEditadas, -1, 0, 1);
    _guiasRefrescarLineas();
    if(res.paraRevisar > 0){
      _guiasToast(res.importados + " materiales guardados &mdash; " + res.paraRevisar + " por revisar", "warn");
      alert(res.importados + " materiales importados.\n\n" + res.paraRevisar +
        " se marcaron con \"Revisar\" (empaque ambiguo, desconocido, o requiere lote). Confírmalos con el lápiz antes de generar la guía.");
    } else {
      _guiasToast(res.importados + " materiales procesados y guardados correctamente", "ok");
    }
  }
}

function _guiasLeerFilasOR(data, cb){
  try{
    var wb = XLSX.read(data, {type:"array"});
    // Buscar primera hoja que no sea SAP
    var sheetName = wb.SheetNames.find(function(s){ return !s.startsWith("SAP"); });
    if(!sheetName){ alert("No se encontró hoja de datos en el archivo."); return; }
    var ws = wb.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
    var det = _guiasDetectarHeaderYColumnas(rows);
    if(!det){ alert("No se encontraron las columnas de Catálogo y Cantidad/X Surtir."); return; }
    cb(rows, det.hdrIdx, det.iCat, det.iXS);
  }catch(err){
    alert("Error al leer el archivo: " + err.message);
  }
}

// Extrae Área / Destino / Fecha de las filas de encabezado del OR (antes de la tabla de materiales)
var _GUIAS_AREAS = ["Herramientas","Misceláneos","Papelería","Cables","Ropa y Calzado","General"];
function _guiasParsearHeaderOR(rows, hdrIdx){
  var area = null, destino = null, fecha = null;
  var limite = Math.min(hdrIdx, 10);
  for(var i=0; i<limite; i++){
    var row = rows[i] || [];
    for(var c=0; c<row.length; c++){
      var val = String(row[c]||"").trim();
      if(!val) continue;

      if(!destino){
        var mAlm = val.match(/^Almac[eé]n:\s*(.+)$/i);
        if(mAlm){
          var mCod = mAlm[1].trim().match(/-\s*([A-Za-z0-9]{3,6})\s*$/);
          if(mCod) destino = mCod[1].toUpperCase();
        }
      }
      if(!area){
        var mOR = val.match(/^OR\s+(.+)$/i);
        if(mOR){
          var raw = mOR[1].trim();
          var norm = _GUIAS_AREAS.find(function(a){ return a.toLowerCase() === raw.toLowerCase(); });
          area = norm || raw;
        }
      }
      if(!fecha){
        var mFecha = val.match(/^Fecha:\s*(\d{1,2})\/(\d{1,2})\/(\d{4})$/i);
        if(mFecha){
          fecha = mFecha[3] + "-" + mFecha[2].padStart(2,'0') + "-" + mFecha[1].padStart(2,'0');
        }
      }
    }
  }
  return { area: area, destino: destino, fecha: fecha };
}

// Importa las filas de materiales del OR hacia _guiaActual.lineas (requiere _guiaActual ya creada)
function _guiasImportarFilasOR(rows, hdrIdx, iCat, iXS){
  var importados = 0;
  var paraRevisar = 0;
  for(var r=hdrIdx+1; r<rows.length; r++){
    var row = rows[r];
    var cat = String(row[iCat]||"").trim();
    var xs  = parseInt(String(row[iXS]||0).replace(/[,\s]/g,""));
    if(!cat || !xs || xs <= 0) continue;
    var m = mat(cat);
    var esPatioMat = (m.ubic || "").toLowerCase() === "patio";

    if(esPatioMat){
      // Material de patio — nunca lleva empaque, se marca directo como PATIO
      _guiaActual.lineas.push({
        cat: cat, desc: m.desc||cat, um: m.um||"PZ",
        cant: xs, tipoEmp: "Patio", contEmp: 1,
        bultos: xs, patio: true, granel: false
      });
      importados++;
      continue;
    }

    var lotesRealesMat = _guiasLotesReales(cat);
    if(lotesRealesMat.length > 0){
      // Lleva lote real — no se puede adivinar de cuál lote surtir, se marca para asignarlo con el lápiz
      _guiaActual.lineas.push({
        cat: cat, desc: m.desc||cat, um: m.um||"PZ", cant: xs,
        tipoEmp: (_guiasEsCable(cat) ? "Bobina" : "?"), contEmp: 0,
        bultos: xs, patio: false, granel: false, revisar: true, loteRequerido: true
      });
      importados++;
      continue;
    }

    var opciones = _guiasBDOpciones(cat);
    var antesLen = _guiaActual.lineas.length;

    if(opciones.length > 0){
      // Proponer el empaque más usado (ya viene ordenado por frecuencia)
      var op = opciones[0];
      var cont = op.cont > 0 ? op.cont : 1;
      var nCajas  = cont > 1 ? Math.floor(xs / cont) : xs;
      var residuo = cont > 1 ? xs % cont : 0;

      if(cont <= 1){
        // Empaque individual (1 pza por caja) — todo en cajas de 1
        _guiasAgregarLineaSilente(cat, m.desc||cat, m.um||"PZ", op.tipo, 1, xs, false);
      } else if(nCajas > 0 && residuo > 0){
        // División automática: cajas cerradas + residuo a granel
        _guiasAgregarLineaSilente(cat, m.desc||cat, m.um||"PZ", op.tipo, cont, nCajas*cont, false);
        _guiasAgregarLineaSilente(cat, m.desc||cat, m.um||"PZ", "Granel", 0, residuo, true);
      } else if(nCajas > 0){
        // Cabe exacto en cajas completas
        _guiasAgregarLineaSilente(cat, m.desc||cat, m.um||"PZ", op.tipo, cont, xs, false);
      } else {
        // No alcanza ni para una caja completa — todo a granel
        _guiasAgregarLineaSilente(cat, m.desc||cat, m.um||"PZ", "Granel", 0, xs, true);
      }

      // Si había más de un empaque conocido Y ninguno está marcado como preferido, pedir revisión
      if(opciones.length > 1 && !op.preferido){
        for(var k=antesLen; k<_guiaActual.lineas.length; k++) _guiaActual.lineas[k].revisar = true;
        paraRevisar++;
      }
    } else {
      // Sin empaque conocido — proponer granel por default, marcado para revisar
      _guiaActual.lineas.push({cat:cat, desc:m.desc||cat, um:m.um||"PZ", cant:xs,
        tipoEmp:"Granel", contEmp:0, bultos:xs, patio:false, granel:true, revisar:true});
      paraRevisar++;
    }
    importados++;
  }
  return { importados: importados, paraRevisar: paraRevisar };
}

function _guiasProcesarOR(input){
  var file = input.files[0];
  if(!file) return;
  input.value = "";
  var reader = new FileReader();
  reader.onload = function(e){
    _guiasLeerFilasOR(e.target.result, function(rows, hdrIdx, iCat, iXS){
      var res = _guiasImportarFilasOR(rows, hdrIdx, iCat, iXS);
      _guiasRefrescarLineas();
      if(res.paraRevisar > 0){
        _guiasToast(res.importados + " materiales guardados &mdash; " + res.paraRevisar + " por revisar", "warn");
        alert(res.importados + " materiales importados.\n\n" + res.paraRevisar +
          " se marcaron con \"Revisar\" (empaque ambiguo o desconocido). El sistema propuso lo más probable, pero conviene confirmarlo con el lápiz antes de generar la guía.");
      } else {
        _guiasToast(res.importados + " materiales procesados y guardados correctamente", "ok");
      }
    });
  };
  reader.readAsArrayBuffer(file);
}

// ── PANTALLA 4: Revisión y datos de firma ─────────────────────────────────────
// ── Transportes base conocidos (tarjetas de captura rápida) ───────────────────
var _GUIAS_TRANSPORTES_BASE = [
  { linea: "Mora", vehiculos: [
    { tipo: "Torton", placas: "76-AP-4X", operadores: ["Osvaldo Lara Juárez"] }
  ]},
  { linea: "Santana", vehiculos: [
    { tipo: "Torton", placas: "004-DX-7", operadores: ["Ángel Ramírez Alpízar", "Gustavo Ramírez", "Alejandro Galán"] },
    { tipo: "Plana",  placas: "660-AR-2", operadores: ["Ángel Ramírez Alpízar", "Gustavo Ramírez", "Alejandro Galán"] }
  ]}
];

// Personal que habitualmente surte las guías — tarjetas de llenado rápido
var _GUIAS_SURTIDORES = [
  "Misrrain Maldonado Villegas",
  "Iván Islas Galván",
  "Elisa Molina Cortés",
  "Carolina García Villegas",
  "Roberto Triano Vega",
  "Ana Reyes Ramírez"
];

function _guiasElegirSurtio(nombre){
  var el = document.getElementById("gSurtio");
  if(!el) return;
  el.value = nombre;
  if(_guiaActual) _guiaActual.surtio = nombre;
  _guiasMarcarSurtidor();
}

// Resalta la tarjeta que coincide con lo que hay escrito en el campo
function _guiasMarcarSurtidor(){
  var el = document.getElementById("gSurtio");
  if(!el) return;
  var val = (el.value || "").trim().toLowerCase();
  var cont = document.getElementById("gSurtidores");
  if(!cont) return;
  cont.querySelectorAll("button[data-surtidor]").forEach(function(b){
    var on = b.dataset.surtidor.toLowerCase() === val;
    b.style.background = on ? "var(--primary)" : "white";
    b.style.color      = on ? "white" : "var(--primary)";
  });
}

function _guiasTarjetasSurtidoresHtml(){
  return "<div style=\"font-size:10.5px;color:var(--muted);margin-bottom:6px\">" +
    "Qui&eacute;n surti&oacute; &mdash; toca un nombre</div>" +
    "<div id=\"gSurtidores\" style=\"display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px\">" +
    _GUIAS_SURTIDORES.map(function(n){
      return "<button type=\"button\" data-surtidor=\"" + _escAttr(n) + "\"" +
        " onclick=\"_guiasElegirSurtio('" + _escAttr(n).replace(/'/g,"\\'") + "')\"" +
        " style=\"padding:6px 11px;background:white;border:1.5px solid var(--primary);" +
        "color:var(--primary);border-radius:8px;font-size:11px;font-weight:600;" +
        "cursor:pointer;font-family:inherit\">" + n + "</button>";
    }).join("") +
    "</div>";
}

function _guiasElegirTransporteBase(linea, tipo, placas, operador){
  var f = function(id, val){ var el = document.getElementById(id); if(el) el.value = val; };
  f("gTransporteRev", linea);
  f("gTipoVeh", tipo);
  f("gPlacas", placas);
  f("gOperador", operador);
  // Marcar visualmente cuál quedó elegido, sin redibujar toda la pantalla
  var keySel = linea + "|" + tipo + "|" + placas + "|" + operador;
  document.querySelectorAll(".gTransBtn").forEach(function(btn){
    var esEsta = btn.getAttribute("data-key") === keySel;
    btn.style.cssText = (btn.style.width === "100%" ? "width:100%;" : "") + _guiasEstiloBtnTransporte(esEsta);
  });
}

function _guiasTarjetasTransporteHtml(){
  var html = "";
  _GUIAS_TRANSPORTES_BASE.forEach(function(t){
    t.vehiculos.forEach(function(v){
      var unOperador = v.operadores.length === 1;
      html +=
        "<div style=\"border:1.5px solid var(--line);border-radius:10px;padding:10px 12px;" +
        "background:#fafbff;min-width:170px;flex:1\">" +
        "<div style=\"font-size:12px;font-weight:800;color:var(--primary)\">" + t.linea + " &mdash; " + v.tipo + "</div>" +
        "<div style=\"font-size:11px;color:var(--muted);font-family:monospace;margin-bottom:6px\">" + v.placas + "</div>" +
        (unOperador ?
          _guiasTplBtnTransporte(t.linea, v.tipo, v.placas, v.operadores[0], true)
          :
          "<div style=\"font-size:10px;color:var(--muted);margin-bottom:3px\">Operador:</div>" +
          "<div style=\"display:flex;flex-wrap:wrap;gap:4px\">" +
          v.operadores.map(function(op){
            return _guiasTplBtnTransporte(t.linea, v.tipo, v.placas, op, false);
          }).join("") +
          "</div>") +
        "</div>";
    });
  });
  return html;
}

// Un botón de operador — marcado como seleccionado (fondo sólido) solo si coincide
// exactamente con lo que ya está elegido para esta guía en este momento.
function _guiasTplBtnTransporte(linea, tipo, placas, operador, ancho100){
  var key = linea + "|" + tipo + "|" + placas + "|" + operador;
  var sel = _guiaActual && _guiaActual.transporte === linea && _guiaActual.tipoVeh === tipo &&
            _guiaActual.placas === placas && _guiaActual.operador === operador;
  return "<button type=\"button\" class=\"gTransBtn\" data-key=\"" + _escAttr(key) + "\"" +
    " onclick=\"_guiasElegirTransporteBase('" + _escAttr(linea) + "','" + _escAttr(tipo) + "','" + _escAttr(placas) + "','" + _escAttr(operador) + "')\"" +
    " style=\"" + (ancho100 ? "width:100%;" : "") + _guiasEstiloBtnTransporte(sel) + "\">" + operador + "</button>";
}
function _guiasEstiloBtnTransporte(sel){
  return sel
    ? "padding:5px 9px;background:var(--primary);border:1.5px solid var(--primary);color:#fff;" +
      "border-radius:7px;font-size:10.5px;font-weight:700;cursor:pointer;font-family:inherit"
    : "padding:5px 9px;background:white;border:1.5px solid var(--primary);color:var(--primary);" +
      "border-radius:7px;font-size:10.5px;font-weight:600;cursor:pointer;font-family:inherit";
}

function _guiasRevision(){
  if(_guiaActual.lineas.length === 0){ alert("No hay materiales capturados."); return; }
  if(_guiaActual.lineas.some(function(l){ return l.pendiente; })){
    alert("Hay materiales con empaque pendiente (?). Defínelos antes de continuar.");
    return;
  }
  if(_guiaActual.lineas.some(function(l){ return l.loteRequerido; })){
    alert("Hay materiales que llevan lote y aún no se ha elegido de cuál surtir. Corrígelos con el lápiz antes de continuar.");
    return;
  }

  if(_guiaActual.cajasGranel == null) _guiaActual.cajasGranel = 0;
  var totalBultos = _guiasTotalBultos();

  var filasHtml = "";
  for(var i=0; i<_guiaActual.lineas.length; i++){
    var l = _guiaActual.lineas[i];
    var descEmp = l.bultos + " " + l.tipoEmp +
      (l.tipoEmp === "Tarima" ? " (" + l.tarimaSubCant + " " + l.tarimaSubTipo + " c/ " + l.tarimaSubCont + " " + l.um + ")"
        : l.contEmp>1?" de "+l.contEmp+" "+l.um:"") + (l.lote?" — L - "+l.lote:"");
    filasHtml +=
      "<tr style=\"border-bottom:1px solid var(--lite,#f4f6fb)\">" +
      "<td style=\"padding:6px 8px;font-size:12px;font-weight:700;font-family:monospace;" +
      "color:var(--primary)\">" + l.cat + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px\">" + l.desc + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px;text-align:center\">" + l.cant + " " + l.um + "</td>" +
      "<td style=\"padding:6px 8px;font-size:12px;text-align:center\">" + descEmp + "</td>" +
      "<td style=\"padding:6px 8px;text-align:center;white-space:nowrap\">" +
      "<button onclick=\"_guiasEditarLineaCompleta(" + i + ")\" title=\"Editar\"" +
      " style=\"background:none;border:none;cursor:pointer;color:var(--primary);font-size:14px\">&#9998;</button>" +
      "<button onclick=\"_guiasEliminarLinea(" + i + ")\" title=\"Borrar\"" +
      " style=\"background:none;border:none;cursor:pointer;color:#dc2626;font-size:15px;margin-left:6px\">&times;</button>" +
      "</td>" +
      "</tr>";
  }

  $("#moduleView").innerHTML =
    "<div style=\"max-width:720px;margin:0 auto;padding:24px 16px\">" +
    "<div style=\"display:flex;align-items:center;gap:12px;margin-bottom:20px\">" +
    "<button onclick=\"_guiasCapturarFirmasRevisionSiActiva();_guiasCapturaMateriales()\" style=\"background:none;border:1.5px solid var(--line);" +
    "border-radius:8px;padding:6px 14px;cursor:pointer;font-size:13px;font-family:inherit;" +
    "color:var(--muted)\">&lsaquo; Materiales</button>" +
    "<div><div style=\"font-size:18px;font-weight:800;color:var(--primary)\">" +
    "Guía " + _guiaActual.area + " No. " + _guiaActual.folio + " &mdash; Revisión</div>" +
    "<div style=\"font-size:12px;color:var(--muted)\">" + _guiaActual.destino + " &middot; " +
    _guiaActual.almInfo.nombre + " &middot; <span id=\"gTotalBultos\">" + totalBultos + " bultos total</span></div></div>" +
    "</div>" +

    // Panel de reparto de granel (vacío si no hay material a granel)
    "<div id=\"granelPanel\"></div>" +

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
    "<div style=\"font-size:10.5px;color:var(--muted);margin-bottom:6px\">Transporte conocido &mdash; toca para llenar automático</div>" +
    "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px\">" + _guiasTarjetasTransporteHtml() + "</div>" +
    _guiasTarjetasSurtidoresHtml() +
    "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:12px\">" +
    _tplCampoFirma("gPedido", "No. Pedido", _guiaActual.pedido || "0", true) +
    _tplCampoFirma("gSiatel", "Siatel", _guiaActual.siatel || "0", true) +
    _tplCampoFirma("gSurtio", "Surtió (quien despacha)", _guiaActual.surtio) +
    _tplCampoFirma("gTransporteRev", "Línea de transporte", _guiaActual.transporte) +
    _tplCampoFirma("gOperador", "Operador", _guiaActual.operador) +
    _tplCampoFirma("gTipoVeh", "Tipo de vehículo", _guiaActual.tipoVeh) +
    _tplCampoFirma("gPlacas", "Placas", _guiaActual.placas) +
    "<div style=\"grid-column:1/-1\">" +
    "<label style=\"font-size:11px;color:var(--muted);display:block;margin-bottom:4px\">" +
    "Observaciones (opcional &mdash; detalles del embarque o destino)</label>" +
    "<textarea id=\"gObservaciones\" rows=\"2\"" +
    " style=\"width:100%;padding:8px 12px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:13px;resize:vertical;box-sizing:border-box\">" +
    _escAttr(_guiaActual.observaciones||"").replace(/</g,"&lt;") + "</textarea>" +
    "</div>" +
    "</div></div>" +

    "<button onclick=\"_guiasGenerar()\"" +
    " style=\"width:100%;padding:14px;background:var(--primary);color:white;border:none;" +
    "border-radius:12px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit\">" +
    "&#128203; Generar e imprimir guía</button>" +
    "</div>";

  _grSel = null;
  _guiasRenderGranelPanel();
  _guiasMarcarSurtidor();
  var _sf = document.getElementById("gSurtio");
  if(_sf) _sf.addEventListener("input", _guiasMarcarSurtidor);
}

function _tplCampoFirma(id, label, valor, seleccionarTodo){
  return "<div><label style=\"font-size:11px;color:var(--muted);display:block;margin-bottom:4px\">" +
    label + "</label>" +
    "<input id=\"" + id + "\" type=\"text\" value=\"" + (valor||"") + "\"" +
    (seleccionarTodo ? " onfocus=\"this.select()\"" : "") +
    " style=\"width:100%;padding:8px 12px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:13px\"></div>";
}

function _guiasEditarLineaCompleta(idx){
  _guiasCapturarFirmasRevisionSiActiva();
  _guiasEditandoLineaIdx = idx;
  var l = _guiaActual.lineas[idx];
  // Abrir modal de empaque con los datos actuales precargados
  var opciones = _guiasBDOpciones(l.cat);
  _guiasPedirEmpaque(l.cat, l.desc, l.um, opciones);
  // La línea vieja se elimina solo si se confirma un reemplazo (ver funciones de agregar línea)
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
    } else if(l.granel || l.contEmp === 0){
      // granel explícito (elegido por el usuario), o sin empaque definido
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

/* ═══════════════════════════════════════════════════════════════════════════
   REPARTO DE GRANEL EN VARIAS CAJAS
   El usuario dice cuántas cajas colectivas va a usar y arrastra cada material
   a la que le toca. 0 cajas = comportamiento de siempre (un solo encabezado).
   ═══════════════════════════════════════════════════════════════════════════ */

// Índices (dentro de _guiaActual.lineas) de todo lo que va a granel
function _guiasGranelIdx(){
  var out = [];
  var ls = (_guiaActual && _guiaActual.lineas) || [];
  for(var i=0; i<ls.length; i++){
    var l = ls[i];
    if(!l.patio && (l.granel || l.contEmp === 0)) out.push(i);
  }
  return out;
}

function _guiasNumCajasGranel(){
  var n = parseInt((_guiaActual && _guiaActual.cajasGranel) || 0, 10);
  return (isNaN(n) || n < 0) ? 0 : n;
}

// Cuántas cajas colectivas cuentan como bulto (incluye la de "sin asignar")
function _guiasBultosGranel(){
  var idx = _guiasGranelIdx();
  if(!idx.length) return 0;
  var n = _guiasNumCajasGranel();
  if(n === 0) return 1;                       // una sola caja colectiva, como antes
  var sinAsignar = idx.some(function(i){
    var c = _guiaActual.lineas[i].cajaGranel;
    return !c || c < 1 || c > n;
  });
  return n + (sinAsignar ? 1 : 0);
}

function _guiasCambiarNumCajas(valor){
  var n = parseInt(valor, 10);
  if(isNaN(n) || n < 0) n = 0;
  if(n > 30) n = 30;
  _guiaActual.cajasGranel = n;
  // Si se reduce el número de cajas, lo que quedó fuera de rango vuelve a "sin asignar"
  _guiasGranelIdx().forEach(function(i){
    var l = _guiaActual.lineas[i];
    if(l.cajaGranel && (l.cajaGranel > n || n === 0)) delete l.cajaGranel;
  });
  _guiasRenderGranelPanel();
}

function _guiasAsignarGranel(idxLinea, caja){
  var l = _guiaActual.lineas[idxLinea];
  if(!l) return;
  if(caja == null) delete l.cajaGranel;
  else l.cajaGranel = caja;
  _guiasRenderGranelPanel();
}

// Reparte lo que falta, en orden, entre las cajas disponibles
function _guiasGranelAuto(){
  var n = _guiasNumCajasGranel();
  if(n < 1) return;
  var pend = _guiasGranelIdx().filter(function(i){
    var c = _guiaActual.lineas[i].cajaGranel;
    return !c || c < 1 || c > n;
  });
  if(!pend.length){ return; }
  // se reparte por cantidad de materiales, buscando siempre la caja menos llena
  var carga = {};
  for(var k=1;k<=n;k++) carga[k]=0;
  _guiasGranelIdx().forEach(function(i){
    var c=_guiaActual.lineas[i].cajaGranel;
    if(c>=1 && c<=n) carga[c]++;
  });
  pend.forEach(function(i){
    var mejor=1;
    for(var k=2;k<=n;k++) if(carga[k]<carga[mejor]) mejor=k;
    _guiaActual.lineas[i].cajaGranel = mejor;
    carga[mejor]++;
  });
  _guiasRenderGranelPanel();
}

function _guiasGranelLimpiar(){
  _guiasGranelIdx().forEach(function(i){ delete _guiaActual.lineas[i].cajaGranel; });
  _guiasRenderGranelPanel();
}

// ── Tarjeta de material a granel ─────────────────────────────────────────────
function _tplGranelCard(i){
  var l = _guiaActual.lineas[i];
  return "<div class=\"gr-card\" data-idx=\"" + i + "\">" +
    "<div class=\"gr-cat\">" + l.cat + "</div>" +
    "<div class=\"gr-desc\">" + l.desc + "</div>" +
    "<div class=\"gr-cant\">" + l.cant + " " + l.um + (l.lote ? " · L-" + l.lote : "") + "</div>" +
    "</div>";
}

function _tplGranelPanel(){
  var idx = _guiasGranelIdx();
  if(!idx.length) return "";   // sin granel no se pregunta nada

  var colectivo = _GUIAS_COLECTIVO[_guiaActual.area] || "Material en General";
  var n = _guiasNumCajasGranel();

  var html =
    "<div style=\"background:white;border:1px solid var(--line);border-radius:12px;" +
    "padding:16px;margin-bottom:20px\">" +
    "<div style=\"display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:6px\">" +
    "<div style=\"font-size:12px;font-weight:700;color:var(--muted);text-transform:uppercase;" +
    "letter-spacing:.4px;flex:1\">Cajas de granel</div>" +
    "<label style=\"font-size:12px;color:var(--muted)\">&iquest;Cu&aacute;ntas cajas?</label>" +
    "<input id=\"gNumCajasGranel\" type=\"number\" min=\"0\" max=\"30\" value=\"" + n + "\"" +
    " onchange=\"_guiasCambiarNumCajas(this.value)\"" +
    " style=\"width:70px;padding:7px 9px;border:1.5px solid var(--line);border-radius:8px;" +
    "font-family:inherit;font-size:14px;font-weight:700;text-align:center\">" +
    "</div>" +
    "<div style=\"font-size:11.5px;color:var(--muted);margin-bottom:12px\">" +
    idx.length + " material(es) a granel. " +
    (n === 0
      ? "Con 0 cajas todo sale junto bajo &laquo;" + colectivo + "&raquo;."
      : "Arrastra cada material a su caja, o t&oacute;calo y luego toca la caja.") +
    "</div>";

  if(n === 0){
    html += "<div class=\"gr-zona gr-zona-off\">" +
      idx.map(_tplGranelCard).join("") + "</div>";
  } else {
    var sinAsignar = idx.filter(function(i){
      var c = _guiaActual.lineas[i].cajaGranel;
      return !c || c < 1 || c > n;
    });

    html +=
      "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px\">" +
      "<button type=\"button\" onclick=\"_guiasGranelAuto()\" style=\"padding:6px 12px;background:white;" +
      "border:1.5px solid var(--line);border-radius:8px;font-size:11.5px;font-weight:600;" +
      "cursor:pointer;font-family:inherit;color:var(--primary)\">Repartir autom&aacute;tico</button>" +
      "<button type=\"button\" onclick=\"_guiasGranelLimpiar()\" style=\"padding:6px 12px;background:white;" +
      "border:1.5px solid var(--line);border-radius:8px;font-size:11.5px;font-weight:600;" +
      "cursor:pointer;font-family:inherit;color:var(--muted)\">Vaciar cajas</button>" +
      "</div>";

    // Bandeja de pendientes
    html +=
      "<div class=\"gr-pend" + (sinAsignar.length ? " alerta" : "") + "\">" +
      "<div class=\"gr-zona-tit\">" +
      (sinAsignar.length ? "Sin asignar &middot; " + sinAsignar.length : "Todo asignado &#10003;") +
      "</div>" +
      "<div class=\"gr-zona\" data-caja=\"0\">" +
      (sinAsignar.length ? sinAsignar.map(_tplGranelCard).join("")
                         : "<div class=\"gr-vacia\">Arrastra aqu&iacute; para sacar un material de su caja</div>") +
      "</div></div>";

    // Una zona por caja
    html += "<div class=\"gr-cajas\">";
    for(var c=1; c<=n; c++){
      var enCaja = idx.filter(function(i){ return _guiaActual.lineas[i].cajaGranel === c; });
      html +=
        "<div class=\"gr-caja\">" +
        "<div class=\"gr-zona-tit\">Caja " + c + " de " + colectivo +
        "<span class=\"gr-n\">" + enCaja.length + "</span></div>" +
        "<div class=\"gr-zona\" data-caja=\"" + c + "\">" +
        (enCaja.length ? enCaja.map(_tplGranelCard).join("")
                       : "<div class=\"gr-vacia\">Vac&iacute;a</div>") +
        "</div></div>";
    }
    html += "</div>";
  }

  return html + "</div>";
}

// Repinta solo el panel, sin rehacer toda la pantalla de revisión
function _guiasRenderGranelPanel(){
  var cont = document.getElementById("granelPanel");
  if(!cont) return;
  cont.innerHTML = _tplGranelPanel();
  _guiasGranelDnD();
  // el total de bultos del encabezado depende de las cajas de granel
  var bs = document.getElementById("gTotalBultos");
  if(bs) bs.textContent = _guiasTotalBultos() + " bultos total";
}

/* ── Arrastrar y soltar ──────────────────────────────────────────────────────
   Se usa Pointer Events en vez de HTML5 drag&drop porque este último no
   funciona con el dedo. Un toque corto selecciona; arrastrar mueve. */
var _grSel = null;   // índice seleccionado por toque (modo "tocar y colocar")

function _guiasGranelDnD(){
  var panel = document.getElementById("granelPanel");
  if(!panel) return;
  if(_guiasNumCajasGranel() < 1) return;

  panel.querySelectorAll(".gr-card").forEach(function(card){
    card.style.touchAction = "none";
    card.onpointerdown = function(ev){
      if(ev.button === 1 || ev.button === 2) return;
      var idx = +card.dataset.idx;
      var x0 = ev.clientX, y0 = ev.clientY;
      var arrastrando = false, fantasma = null, zonaPrev = null;

      var mover = function(e){
        if(!arrastrando){
          if(Math.abs(e.clientX-x0) < 6 && Math.abs(e.clientY-y0) < 6) return;
          arrastrando = true;
          card.classList.add("gr-origen");
          fantasma = card.cloneNode(true);
          fantasma.className = "gr-card gr-fantasma";
          fantasma.style.width = card.offsetWidth + "px";
          document.body.appendChild(fantasma);
        }
        fantasma.style.left = (e.clientX - fantasma.offsetWidth/2) + "px";
        fantasma.style.top  = (e.clientY - 22) + "px";

        fantasma.style.display = "none";
        var bajo = document.elementFromPoint(e.clientX, e.clientY);
        fantasma.style.display = "";
        var zona = bajo && bajo.closest ? bajo.closest(".gr-zona") : null;
        if(zona !== zonaPrev){
          if(zonaPrev) zonaPrev.classList.remove("gr-sobre");
          if(zona) zona.classList.add("gr-sobre");
          zonaPrev = zona;
        }
      };

      var soltar = function(e){
        card.releasePointerCapture && card.releasePointerCapture(ev.pointerId);
        window.removeEventListener("pointermove", mover);
        window.removeEventListener("pointerup", soltar);
        window.removeEventListener("pointercancel", soltar);
        if(fantasma) fantasma.remove();
        if(zonaPrev) zonaPrev.classList.remove("gr-sobre");
        card.classList.remove("gr-origen");

        if(arrastrando){
          if(zonaPrev){
            var c = +zonaPrev.dataset.caja;
            _guiasAsignarGranel(idx, c > 0 ? c : null);
          }
        } else {
          // Toque simple: seleccionar / deseleccionar
          _grSel = (_grSel === idx) ? null : idx;
          panel.querySelectorAll(".gr-card").forEach(function(c2){
            c2.classList.toggle("gr-sel", +c2.dataset.idx === _grSel);
          });
          panel.querySelectorAll(".gr-zona").forEach(function(z){
            z.classList.toggle("gr-destino", _grSel !== null);
          });
        }
      };

      try{ card.setPointerCapture(ev.pointerId); }catch(err){}
      window.addEventListener("pointermove", mover);
      window.addEventListener("pointerup", soltar);
      window.addEventListener("pointercancel", soltar);
    };
  });

  // Tocar una zona coloca ahí el material seleccionado
  panel.querySelectorAll(".gr-zona").forEach(function(zona){
    zona.onclick = function(){
      if(_grSel === null) return;
      var c = +zona.dataset.caja;
      var sel = _grSel;
      _grSel = null;
      _guiasAsignarGranel(sel, c > 0 ? c : null);
    };
  });
}

function _guiasTotalBultos(){
  var g = _guiasAgruparLineas(_guiaActual.lineas, _guiaActual.area);
  var t = 0;
  g.cajas.forEach(function(l){ t += Math.floor(l.cant/l.contEmp); });
  g.patio.forEach(function(l){ t += l.bultos || 1; });
  return t + _guiasBultosGranel();
}


// Abreviatura del tipo de empaque para el impreso (C/C para Caja, como ya se usaba; nuevo para los demás)
function _guiasAbrevTipo(tipo){
  var map = { "Caja": "C/C", "Bobina": "Bob c/" };
  return map[tipo] || (tipo + " c/");
}

function _guiasBloquesImpresion(lineas, area){
  var grupos = _guiasAgruparLineas(lineas, area);
  var colectivo = _GUIAS_COLECTIVO[area] || "Material en General";
  var bloques = [];

  // BLOQUE EMPAQUE: cajas cerradas + patio van siempre juntos, nunca se separan entre hojas
  var bloqueEmpaque = [];
  grupos.cajas.forEach(function(l){
    var nCajas = Math.floor(l.cant / l.contEmp);
    var descEmp;
    if(l.tipoEmp === "Tarima"){
      descEmp = "Tarima de " + l.tarimaSubCant + " " + l.tarimaSubTipo + " c/ " + l.tarimaSubCont + " " + l.um;
    } else {
      descEmp = _guiasAbrevTipo(l.tipoEmp) + " " + l.contEmp + " " + l.um;
    }
    if(l.lote) descEmp += " L - " + l.lote;
    bloqueEmpaque.push({
      cant:    nCajas,
      um:      descEmp,
      desc:    l.desc,
      cat:     l.cat,
      total:   l.cant + " " + l.um,
      tipo:    "caja",
      patio:   false
    });
  });
  grupos.patio.forEach(function(l){
    bloqueEmpaque.push({
      cant:    l.bultos || l.cant,
      um:      l.um,
      desc:    l.desc,
      cat:     l.cat,
      total:   l.cant + " " + l.um,
      tipo:    "patio",
      patio:   true
    });
  });
  if(bloqueEmpaque.length > 0) bloques.push(bloqueEmpaque);

  // BLOQUE(S) GRANEL: cada caja colectiva es un bloque que nunca se parte entre hojas
  if(grupos.granel.length > 0){
    var grupos_granel = grupos.granel;
    var nCajasG = parseInt((_guiaActual && _guiaActual.cajasGranel) || 0, 10);
    if(isNaN(nCajasG) || nCajasG < 0) nCajasG = 0;

    var filaGranel = function(l){
      return {
        cant:    l.cant,
        um:      l.um,
        desc:    l.desc + (l.lote ? " — L - " + l.lote : ""),
        cat:     l.cat,
        total:   l.cant + " " + l.um,
        tipo:    "granel",
        patio:   false
      };
    };
    var armarBloque = function(titulo, lineas){
      var b = [{ tipo:"separador" }, { tipo:"colectivo", desc:titulo }];
      lineas.forEach(function(l){ b.push(filaGranel(l)); });
      bloques.push(b);
    };

    // Se arman primero los grupos para saber cuántas cajas colectivas salen de verdad
    var grupos = [];   // {caja: n|null, lineas: []}
    if(nCajasG < 1){
      grupos.push({ caja:null, lineas: grupos_granel });
    } else {
      for(var c=1; c<=nCajasG; c++){
        var enCaja = grupos_granel.filter(function(l){ return l.cajaGranel === c; });
        if(enCaja.length) grupos.push({ caja:c, lineas:enCaja });
      }
      var sueltos = grupos_granel.filter(function(l){
        return !l.cajaGranel || l.cajaGranel < 1 || l.cajaGranel > nCajasG;
      });
      if(sueltos.length) grupos.push({ caja:null, lineas:sueltos });
    }

    // Una sola caja colectiva en toda la guía: se nombra "1 caja de …"
    var unaSola = grupos.length === 1;
    grupos.forEach(function(g){
      var titulo;
      if(unaSola)            titulo = "1 caja de " + colectivo + ":";
      else if(g.caja)        titulo = "Caja " + g.caja + " de " + colectivo + ":";
      else                   titulo = colectivo;
      armarBloque(titulo, g.lineas);
    });
  }

  return bloques;
}

// ── PANTALLA 5: Generar guía imprimible ───────────────────────────────────────
function _guiasGenerar(){
  var nCajasG = _guiasNumCajasGranel();
  if(nCajasG >= 1){
    var sueltos = _guiasGranelIdx().filter(function(i){
      var c = _guiaActual.lineas[i].cajaGranel;
      return !c || c < 1 || c > nCajasG;
    });
    if(sueltos.length){
      var nombres = sueltos.slice(0,6).map(function(i){ return "  · " + _guiaActual.lineas[i].cat; }).join("\n");
      if(!confirm("Hay " + sueltos.length + " material(es) a granel sin caja asignada:\n\n" + nombres +
        (sueltos.length>6 ? "\n  · …" : "") +
        "\n\nVan a salir al final, bajo el encabezado general y contando como una caja aparte.\n\n¿Generar así?")) return;
    }
  }

  var faltanLote = _guiaActual.lineas.filter(function(l){ return l.loteRequerido; });
  if(faltanLote.length > 0){
    alert("Hay " + faltanLote.length + " material(es) que llevan lote y aún no se ha elegido de cuál surtir " +
      "(marcados \"REVISAR\" con ? en la lista). Corrígelos con el lápiz antes de generar la guía.\n\n" +
      "Catálogos pendientes: " + faltanLote.map(function(l){ return l.cat; }).join(", "));
    return;
  }

  // Capturar datos de firma — SOLO si el formulario de firma está en pantalla (flujo normal
  // de captura). Al reimprimir desde el historial se llega aquí sin pasar por esa pantalla,
  // así que esos campos no existen en el DOM; si se leyeran de todos modos, se borraría con
  // "" lo que ya se había cargado correctamente desde el historial.
  if(document.getElementById("gSurtio")){
    _guiaActual.pedido     = document.getElementById("gPedido")?.value.trim() || "0";
    _guiaActual.siatel     = document.getElementById("gSiatel")?.value.trim() || "0";
    _guiaActual.surtio     = document.getElementById("gSurtio")?.value || "";
    _guiaActual.transporte = document.getElementById("gTransporteRev")?.value || "";
    _guiaActual.operador   = document.getElementById("gOperador")?.value || "";
    _guiaActual.tipoVeh    = document.getElementById("gTipoVeh")?.value || "";
    _guiaActual.placas     = document.getElementById("gPlacas")?.value || "";
    _guiaActual.observaciones = document.getElementById("gObservaciones")?.value || "";
  }

  var alm      = _guiaActual.almInfo;
  var hoy      = _guiaActual.fecha ? new Date(_guiaActual.fecha+"T12:00:00") : new Date();
  var fechaStr = hoy.toLocaleDateString("es-MX",{day:"2-digit",month:"2-digit",year:"numeric"});
  var totalMateriales = 0;
  _guiaActual.lineas.forEach(function(l){ totalMateriales += Number(l.cant) || 0; });

  // Calcular los bloques de impresión (empaque = cajas+patio juntos; granel aparte)
  var _bloques = _guiasBloquesImpresion(_guiaActual.lineas, _guiaActual.area);

  // Paginar por BLOQUES: un bloque nunca se separa entre hojas si cabe completo en una página
  var FILAS_POR_PAGINA = 21;
  var paginas = [];
  var paginaActual = [];
  _bloques.forEach(function(bloque){
    if(bloque.length > FILAS_POR_PAGINA){
      // Bloque más grande que una hoja completa: excepción, se reparte forzosamente
      var idx = 0;
      while(idx < bloque.length){
        var espacio = FILAS_POR_PAGINA - paginaActual.length;
        if(espacio <= 0){
          paginas.push(paginaActual);
          paginaActual = [];
          espacio = FILAS_POR_PAGINA;
        }
        var trozo = bloque.slice(idx, idx + espacio);
        paginaActual = paginaActual.concat(trozo);
        idx += trozo.length;
      }
    } else {
      if(paginaActual.length > 0 && paginaActual.length + bloque.length > FILAS_POR_PAGINA){
        // No cabe completo en lo que queda de la hoja: pasa entero a la siguiente
        paginas.push(paginaActual);
        paginaActual = [];
      }
      paginaActual = paginaActual.concat(bloque);
    }
  });
  if(paginaActual.length > 0 || paginas.length === 0) paginas.push(paginaActual);

  // Función que genera el HTML de una página completa
  function _htmlPagina(filasPag, esFinalPag, esPrimeraPag){
    var filasHtml = "";
    for(var fi=0; fi<filasPag.length; fi++){
      var f = filasPag[fi];
      if(f.tipo === "separador"){
        filasHtml += "<tr><td colspan=\"5\" style=\"padding:6px 0\">&nbsp;</td></tr>";
      } else if(f.tipo === "colectivo"){
        filasHtml += "<tr><td colspan=\"5\" style=\"padding:5px 8px;font-weight:700;font-size:12px;border-top:1px solid #ccc\">" + f.desc + "</td></tr>";
      } else if(f.tipo === "patio"){
        filasHtml += "<tr style=\"background:#fff8e1\"><td class=\"col-cant\">" + f.cant + " - Patio</td>" +
          "<td class=\"col-emp\">" + f.um + "</td><td class=\"col-desc\">" + f.desc + "</td>" +
          "<td class=\"col-cat\">" + f.cat + "</td><td class=\"col-tot\">" + f.total + "</td></tr>";
      } else {
        filasHtml += "<tr><td class=\"col-cant\">" + f.cant + "</td>" +
          "<td class=\"col-emp\">" + f.um + "</td><td class=\"col-desc\">" + f.desc + "</td>" +
          "<td class=\"col-cat\">" + f.cat + "</td><td class=\"col-tot\">" + f.total + "</td></tr>";
      }
    }
    // Total de materiales (HLog) solo en última página
    if(esFinalPag){
      filasHtml += "<tr><td colspan=\"5\" style=\"text-align:right;padding:5px 6px;font-weight:800;font-size:12px;border-top:2px solid #001E6E\">HLog: " + totalMateriales + "</td></tr>";
    }

    return "<div class=\"pagina\" style=\"page-break-after:always\">" +
      // ENCABEZADO
      "<div class=\"hdr\">" +
      "<div style=\"display:flex;align-items:center;gap:32px\">" +
      "<img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAL4AAABiCAYAAADuiGnJAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFxEAABcRAcom8z8AABwTSURBVHhe7Z0LmBxVmfdfLqKrkJgZJtNVdS5VPQFcvLAsKiir7Kr7Kayiq+YDRT8E9wNdlYuAK4qOggRYiJDMdNWpnklCAEEChMQQvCAi3rgl5DJddap6MrmRcFkuBghJIJfe5z013VNd3ZN0ksly2fN7nvfpmapTp05V/+vUe855z2kAjUaj0Wg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaDQazeucYplAb8Bft+ZJG3qkCd2lg7KXptGMjlv6PXjyL69bE3IReNEtUBj42+ylaTSj44avwPVrKq9bu+nJCnjhavBK789emkYzOp58Fopx5XVrM1dVwAskiOXvzV6aRjM6Wvia/5W0JPxyBfpaMEzXcOxo1uT4ZrarPFsT/gHvMM12QohFKTX3xPBYznkOAPbHDNl4NsGyLJJN14phXvnO/EQA2G+4fPtPyuU6Jg3nZ1mTSEdHx8H1l9AahJC/YYcyo9m1YnkZYxOyx6TJ5XIdza4L88N9eC+Hk76pg/Ncs/O0YnjccFnUPTAM49Aus4tW96uyjt95WRHe0VErAyGHWaZp0tbu3a6Ej+KbMVSBGSt3bjPxc2hYrE3yqbMW81S2izxbEH7+bZ0THUJ9h/LFDmUP7JExvsi2rF+OGzeuDb90MnHiFIfQJQ3pWjHGF1GT3l79gg455JB2m9KZecaWONR+APMlnZ1nZK+jFbhlnWgT9sc84w83nJeyJbZlefhwZI9DnEMPPdw2zblN7xPhi5lFru/s7LQxrd3efoRNyIJRztOC8cWMkJ90dna+DfMjBvmibZElNuMPqf14vwm9m3R2jtp2o5Z1nkPZg8qYjXkutSmba+dyPJu2ES/8a0MtW7X+QRTXy+DJ+6EQzofe4K6m5ir7Bfjyt1CUG6BPHdfccJ8IXwI3+DP0Br9oyCtteE5PPgi+3F57ALI2a3UFRBDtTPhWWxvBm3N4vqsyyck3mu3U2WH4mUlzWL6rkidsgzNxYifAMW/ihjVfpUtbNt9RDPOihK6dMGHCeCwf1tA2ocsOw30qH6fCLetH2etoBZvS0/OcVzCvhuty8hWbsPWjiYlb1nldbLRjHTy2xHLsSExLc7n35gn9a0PaJtfbzPAe2IT9grHxqla32izCDOuWLm7XfQ+csvt5jjf02FHT/IpD2fPV8+MnJ/RJ2zRPOeaYY96UTd+IK7fB7McqDXbDugrc9HglEf0gHU5dfTVnSbb3hIeDkEtULZwVfNVmYQ0tV8P00kdSxzbLN9nmDzEQ0QNwy9NJmbLlvPm/UPhroXfpsdkMqqDwOWUPKVHhjU0ZftFdzK7gDU9bNp36EizyPAr/hBNOOJBb5FZ1fNVY4zGjWfKF0scmTMgr4fMOnnMIfVB94ZiXbeP+72evoxU4IV90GNuCeaVFVP07z+zN1DS/nT1uouN02oT8CctWuwd1f9so1EXc4O/A9Mww/j5P6OPpNM3u22iG57EJXUCIeoMqurq6qE3IQnUPhs1hvIIPBLpC1XTUNE92CF01iY9cl03YE9Skp1TT7BpP/gg8eWWD+fGV4IZXQW/wUegJ28GTZ0NxxUXgywvrrH/wQvDK56rBJHfFRBDRIiXurOBrwscaOlwFXnwUTC1NAhF9u2m+uK2n9A2YtoxAMf6QKk9RXtVQTk9OBVdeAF5sZS+tyhHt7Ycwi53FDGMqN80rq2abuSncMG+0CdtQu9GUbWWELLRNc4ptmldU0zqGcTU1ze+jezJp0qQ31wmf2RVukD/bOfPydP6jGeblmOQCzvlbsHz7Uvh5xjc6lG1Ji9mm9JemabanjyOEnOEw9krquFccSl/sYnxbq8LnhPwB71n6vo1mtmFMdQxyWtbtMgzjGE5odJidH3mgGK9w05o+qa1tnDXROs6mPKg9zBwfZr6RW9b5o1Sge4GIroab1ldUrfuzp+rtlmcqMHN1BXrladBbyoEfLdp5jV8VfvlIcOOzYPa6Cvz82Sb5Pl2BGx+vQCH4abY4Y4lt20c5zF5VrRnzjG/KE/KFbLo0nZ3veVtV+FWxOuaeCRXZV8JPakIacIvMyzN7I/6PNW2e8g3EIB+vHnMkwEHMMNwuzkeuh5DF3DRvz9PkYWhF+F2MXVhXmD3ENDs/6VC6Ov2wOoRucUw2mxP2SO2cuJ3xjcwklxgGvDWbz95RkB8EL3pciTsrYrT+FRUQ8UvQKz+/W8IvDr4TvPj/QTF+Rfnv2XRo2GD25OMgwr/PFmussAn5sMP42hHhsy22Zf3/bLo0zYTPTeuqI4888qCdGb4p8DNbM+1L4TuURSRHPsEpv0f514m7U7Fy5lUjPSrt77Ata3n1/GjMJIJa9Md5bu9oVfic0kuz15w1vAet+ODMsM6yGX+x7sFKWeLm8B3cov85aRK8OXv83tF934FQiKYlAh+lwbovhY82ey2mvwn8IeUPjzV7K3z8ElBIqmFF2aM248tGM4fbASf0Hm6aR6fz28fCX0UmkndTi34XG72pGnRVZ2fnO/EYZhjn5BnfUd2Xp+yvtNP8lE3YJdVtrQjfpvQJh7Kl2euuM8JjZtEZ2XJnQTeQWdbU/LCrlTV032xCZhkHH1zz/ccON/oI+IOP7dRn39fCx31euBGmL/9MtnhjwVgIvyoOdCOw90N9DueXGNa0edXz4DD+DCHkhHR++1L4ecpWc5MfbZrmBx3Gnq4JFf1mg56HwrEt64/V7fjJCb0NAA50KL9sd4Sv7gFef9N7kNwH1bNkkb9ky90MPn782zmhfhe3t6dFn8c3kkVuJeNIrWE8dvjr3woiukX1pOysH31fCx/7/PtXVqAg5+6LWn8shI9fhMP4Nodig5BvdRjf6nC+Nc/5NmWMbUfD17NN+VrDMI5P57evhU86yfvRxbApvTMtVG6SecwwvtpF6ObquR3KXqCG8WnMyybs8t0Rfp7x7er6s/eA2dtwX3IP7IpjmPdmy90MAvA3nLArUfjZB8w26e/a2iySPWbvcYMvgx+/qPzsBjGmLCt8ES1W4s72t1cNA8uEXN268Gu1/iYoRF/KFnNv2Vvhqy/EdiqM0l9xk5zJDHYOY+wcm7Fz8pSem6f0PMeyzkezCbvItthZbW1tdT1Q+1r43OLH4T5qml9zaNKAVcIn9ElOSAnLXxWUQ+mD5gRTdV/vrvAdQu9iBvkWM9i38PqH78F5ecrPcyx6Pva82Bb9D8eyJmfLncUwjLdSw/hxniYPZbrGr5aJWfRO1tnpZI/dc65b3gkivAdmrWkUYdbSwhehAb5cpnpqGkZgh+2G9Vjjr4P+Fe9qWfho+BYpyPth1pK3Z4u7N4yF8PHTtshPsulaZZ8L3+QfwH1me/sRjkXWV4WUZ1xZqsauOJZ1WbXRaxM2ZXeEnyfsh9ny7Ck0Z37dpmxj2p1sZpyQ2eM5b0UTKtRk53il00AELzQdgVUjuitGYmiqwnejycoVEbIH/PCXIIIFo9ivQch+9ZB48isNwsf8mrlW/UPJPhx3qFTGrL92TISPxxE21sL/XjZdK+xM+ChomjO/nxZ71TAtJ2y9cahR60HbbeEze0yEb+Vyn7MJfbKad1I2upqa5o02pRtq25M+/K2c0suqYyJZcLyCm+bfccs6juTI+3CUPJsmwVtigVtanLg4WQGWK+DLHeBHIzE0ifA3ghf9X3X8tLvfDP3yEOgtHdzUcJ+/KOlzdYMzQUTDwh92g4rRSABb3QMXJ/u9YAl4Kw7LlHqPGSvhOxa9NJuuVZoJnxHynWy6VtiF8IF2dr7Ttpgat2gQvkX7CIwMKO2u8B3K9+hhTWNPNE/mhD6Wztem7ClmWZMxxMMmZCq6ZulRW4fxl6hhfDcVQKcghLTxzs6vsFzuS8wgXyUmOZWY5My2trYmbQMhv5c0KJvU9terrsUHwStdDEIuU4JF4bvyFSiU5oEXXg6F6DIohJcln80svAw8/MQR4/BuENG2WoCbH5egUPo2uNFvkvNnxb8Cz/eKKmN3pYVX164ZC+FjD4ltkT/ZhPzQJuSKXZlD6dUo7GYjt/ilYrcjM8nv0O2wCZmSPT5tnJArbUqvsXPsyyovQr6wM+G3tbWNYyb9WVb0ecqfYRPNj6avc3eFbxN6H45et1Rmy7qGGMYX0iO3GP9jW6SULptN6IvEML5VTaNqcEJvV+Wqlj9J9zQ+9NV0Kj/D+Gdqmp+iOfN07J5llvVlK2edbUw06tIBTFt6NLhR3LRXJhH5FnAH/i1JWz4aRHQzCLm15r9joxYNB7t2ZirdmiRPdS65HUR4J/SEyRfkDnwS3PDZpr5/0jheCddF76kv/J4xLHxVwwwL/2Xbss7KpkszLPw5NeGnBNSKYZceIxirkwSpofBtiz6UzivJr7U8VRchZQtVXoSc6jD2cnWfQ9matPBVGtM8M91oxE+b8puy4QM2YVfUC58uzgj/iYYy88byNTPs8nQImY9dlpifZXUea9ORe5CUnW6ihvWjbLnMQ83DOWG1Llg07Dp1KFvNTPbJ4WT7Wznrc1aHdZRtWT/hJjmDETLVsnIncYtenM4PoDe8SokrW9uj24EBYV4wE6aVR0bJrpKHgBv+APz4BSVm5f4MDvfg4GczG3ZjMC2ODwj5PIj4x3BjeVwt3ysXYVthXuL6ZGv9cvLweLK3rix7iE3ICQ5lT1SjBvOcb7ct6+xsujQYSstMcocK6Bq27Be7M0OhUkJXZqIzF+P59zQ/zvkCzAtrvTznO6rX4zC+nhmsrvsUY9ZtSh9WkY1Jmq3Uog3X7Fj0KlUWlU7F+CyrRkoywzjGoezpvSmzY7G5GPuUy+U4t6zfV/cN9zLtoJROw27YbLkQZrKPOowlDfXUcZzQlViZodtj5azPWpZ1lE3IhXgOahjnWh25zzPTvGQkJzc4GkQcKQFna1l0RXz5X1BY/i/pk9dw5cngBX9Wbo9q+GaOz1o1HXZ9isFT1AhxFj/4GIjoOdWozR6Pbxcveg6mBydmD9td0OelhBY55Qs44fM4Y7fahvF/sunScOBvYaZ5kU3Z3YyyO3bXOOPzbYt42GWn8hvP384tehWjbGE2bQs2F8tOraQxbOdyH+aMzcFzJNdEi9wwVC2dxrbo2Tbld3E0TJPLNYT+Ojn2JU7YfJuyOzmlC9A9wQkiuI92dHTZlM3kjM1vUqZd2VxO+AKeMy9A4duW9WVOGN77eck+Np8RMmWiCgEfHRS2TeltjLI7MV9VTkLvIqZ5YTvAIZZlfcQwjM9QSk/G9JTS9+IoNe00v5LkgINVbjBb1bANjcrBRPgiugmmpWrlLF7pSBDhdVCMdyQ1dRPBo+E+FHNBXg/uindls6nhL3oT9AZTYcaqpAx1+eCbAx8ceQN0Vxofmt3jQPR71Sjh+PFvz0+YMB6DtrKJsqBo2fjxE6rH7a61H3HEIamYnf2O7Og4eG/yq07owEkyeA3V7Xht2UbfMAfhfjxn6tg6sKGbPkd7ezuWuZrXAfjGypajVcPzDj/4+2MjVDVEU/tHq+mzpL87NDJuXJsKYZ48+QCcMcZy1mRqWRdjw9ayrHMZId/C2j85uiA/BW7wfNPBKvTB3ehZKIR1w+zNqewHxfgi6CuvVG+OtMuEf6vQh8E14IbdcG0L/fHu8mPAC4ZGbXMI+RyIsPlbSKMZrqCIaZ5BcuYlzLDOGZmdhbW4iOYmo60pH7zaVYnbeuW1u9WL0jPwjyDkr2sPDj5QWMvjpJbrgpOyyUdlTuUAEPJSVQY8Pl023Jb0Mt0D10T7IFBJ8wYjM/bjyc9DMdqioiCVSFOGwirIASiu3v1hYXfpRNUA9aPNUBzaDAXpqpCG3aUveg8IOahi87Plwwa3eqCi85u2EzSaUXHD74IrfwlCLgRP3lUzVy5UI7Bu8HWYPKeZj7hreu87GIT8GnjxN6H7vhZmvo+CG30dPIkjvo1l9ORvwQ8vhylh3YwijWbnzHi0Q00t7Fvb1mCzw/ax6DLca+6770CYWkrKky1jz7p26Cu1wVmLdjm54VVlWrkDppfeP5bhFhrNqwc20t3wOHCX57O7FBjCgbPYRPRrEPHd4K9/7bRHkqDCD6YWE0jARQPcoABCXgT9T2NPjkaToRCcAr4sgwi+lt2l8AaOBy94BLw4AlE+A25d23RNm1cFEf4ARBRBX/zZuu04fxrbVCIsgyh31e3TaBS4+gT2MGE7phk/fcSAYvRVKAzU16qvBYrSBT/aCv1R/ew27OkrxpeAKJ0Bc/Ty629csLepLz4J3PBUEPKf1IBePfuBu+w4FXrthV9QS6P0PnUw9AbHghdOAz/aBiLqAS86HqYFR6seJmwbCflu6A1OAj/6DLjhJ5UL0Qp9Q53glf8ZegY+B350Yt1D073oraq72A0mgx98pi5S9YYVE8Ef/CD48cegXyYLbOE5vWWfhwKmHV6CZWpsgRt/CEQ0H4rxJvDlj8ANPgTTl78DpkeOipnqWfKPUIj/VnUppylGTlKm8LNqxN6NPwbikfpQ3+5VbwERf1hF7HrBp8GVR9Tt17wGEPK9KhrUjTZCT4ArvP0V/NCDnoeSXqKpa9tU1GmxvA68YDP0lyvglTxwVZBeDH55KxSjHVCMNiUiiheoyTIiOAVEuBLcYAt4wRbw5RaYuWop9JZOyxahjh6JYyD3ghdtBKGC914AIZPoRBSdJ68HP3pOBQd68hUQ4aNKgIgb/Rt40TNQjDaAJ38O7uDHVRStL7eBH20Ct3QHTHu0A7zS2SoPrO0xzFyEWPYt4EU+CFkEET8DxRXPgghnwaxVwwON3furWXnFaDF4cjO4EUbJboK+8tPgB2fWyl94mIIb9oOIngEht4EnXwY/XgZekISta14DYE3qBb8HITeCG38Prh34OLjhjeDHKO6L1aCdF35Tzfn14t+CW/oieANfBTc4EaYusUDEF0CxfDf40XYQ0W+gb8UPQcSnwy9wvnLwNfCCu8Arn698ZhEKNdrshoMwPWqInVFgI9kLhxLBRHOgWO5WE3tc+XE1K84L74Q+DN8uL4Ti4PfUD2Lg3AgRrlFLsKA/LuQlIOQ6cOWTIKIAivHtIMr/AX45mQtdCK6AYvx3IMJuNT9axK+AV74D+ocuBRF/FvzwBPDCa1RaIe8Cf1EyzxldOhHhQzgEXnwp+PHZ4Mq5MAuDF0NczAkrCUt1MeMAowh/BX58MXjhz0DgwyUfVw1pzatNZT9V82H0pzdwqZokg/TgG0AOgogeV69oT96iguLc8BvZHBQoAHR1cBBthP1Ub081TwTHNLzSAvDkJiXkLNjN6cU3KCG7pe6au4UxS/dVDgQRX6SiUt3wVrh6Ka60XF0QYJoaYS+E09Q2nOjjBn+EWWuTtNX4KhxPUTFX8g+1cwpZgGL8EvjxqbVtSBGXggzwYZ4PNzzxNvVzS14Ug4jWKResCroy/fiAxMkbCWfUoeg9eQf0LzbVNmzMYwzXrc9hvNd16no0ryLu8glqDgA2TN3ofvDlPPCjhSDkPSCi56EYbVe/suIF31eBdX5UBjc4B3qWvbsufKMv+joUYxRJw3qU6uG6LnyX8vHxIcOfLvKjF0GUPpFNmfyeVxyDFz0Gheh9dftQ4H44R4WAuLJ+EoU79CFwww1QiBJBu2smqIV2RfQS9AZq3RyFiE5Rs9i8cET4XoyN280gMnliwKFb2qHuCT68OMlf1eJygQpFr+IHk4ffhknDXgQFtcCvF9ZP5MFI2xkrt4If3a/GjDSvItevtcCTi5Ro/fA2EKqR6oHAng7Zq5ZNxLAL/GE5N7wOvFCCULPA1oIrL6oJoBh9Q+XhxfXL51336GFqMS43KqlFb73wXrUqnB9ugOJgY43vD70PRLgePBmrRnGaKevaoSjvAhFuVm5WGmxroFvjRg+q/zFuyQsfBiHXwpzKSI9MUZ6WCF/+vrbNi71E+JnVK9yl71LCF+E8NU/Ci76jJg5hhG76R/bcgVNVjY/uILYB/Pg28KMtqgMgjVh5LHjyRfDjR9T6qppXkb4NbSDkb6A4uKMWeYruRm1kNTXCOmfOAcrtmY6u0QpsWG4CIT+l9qHw0cf3U66Ocg/CmeDKF4dDQo5Q/ncSavESiLAx3h97irxwLQi5BsSK+iXPuzGQUM6FvvJ2cAfVFMMaXnASeNhGkXPV/zhA5oYPg1daDVNTYwbNhO+WhRJ+Vqg14ctE+P7Ad+B6nPxT/nndSL4If6JW1EjcQIzMnZE0lmX9mEZv8GmVTkS3wdVPNA2B1vxPge4Krsqswq5lvwp9qO2770C4W33BqQcBAE7oPlD5yDhHwI1PV9vQx0cf249H1rP3VtlqKfRC8BBMezDxsYvRR0DIp8ALX4D+cqPw0Rf35R/UfAY36oYbliYCQaGpyftyqlqaBWvVq4drzZ41pvLDRbAD3DApD4ZGKOGHuxa+iAuqfeJjjY3X3Z24cCM1/nzl6rhLsVsV2zGL4No4mbCCA3MiHEraHVHS/vFw9Wx1b+bVumD715jK58d0BXmGDtl4LYB94F70AMzE1RswpEBOASGvAC+4HUTpAuiecxD4gydDcdAFEV8KIrgZvBD77P8AbpyEKLjRicr3Rt8cazxRPi/JFyflBy+DCGdAIbgWRPgn8MI14EUovgWqYYhvkjR98gy13IqnxgUWgIhmKHcD/XpsdBfLMllcK7oHvAjdlL+Aj43J6PpaIxbHAArhcvDCp+qEL8LTYbaar/xAbRtOG8U5En44CF48s+abo4+PUa8i+hXMeepg1Zj2gj8mM9/kveDHvSrEXIRLwJPbwZXnqONwzrYnw6RiiPBnZF1VUeCkoUJwM/iLXjvhGv/rUXE04ULwo9Xgx49BMVqj/Hm39AOYM+cg5crMGFoPfvQYiGgVuKU7wRs4qna86maUt4AXrYHi4HoQ0SzVcHaXnwkiDMCP14MblmB62A294TfBkyVwSxK84JqGUdHJlQOgIC8EN1wGnlwDxRgn7ZSgVybzYfG3Crzod+pcWFYl2KAPvLUjK7Nh4J4b3azGJu7eMeKWFOJ/ASEHlAtWBbtP/fBedW39Q+vAiy5X26+OHCjIEnjR9FpbRuC5S/i7wmvAj1ZBIegdXh4G31Dn1fIUgx8AL7xfrZaHvUCuXAFCzoZpAzr04TUHLoSlRmNL/wAiPlb549WGIdagKBDch4JHtyOLHx2qBsJ6ln9AjXwiGMqtukMHjoeZOLqKjb/oUPVD1F5wmIooHQ10EwoD71O9O72Dk2BWZWSxJBxYc+UxKl8cD8jOQ8C3CP6KDL6R0m4FBpthebIBafg/jkCL4FjVkEewAVscOlyN9KZD0tEdxHJ5K49S3ZLe8n9NenWCc1M5JvdTlTE6Xo3+7v30UI3mNYQXXKsG1LDm12jecGB8EPr/XvjvUIz/Vc3a86L/BIEhEzJWbwGN5g1HN845kL8FF3t2ws1Jl2z8Aoj4YfUQ6J4azRsSbEeI6BPgln8AbvBT8LDnS56mBvg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go1Go9FoNBqNRqPRaDQajUaj0Wg0Go1G8z/EfwPKygPvSBoo5AAAAABJRU5ErkJggg==\" style=\"height:96px;object-fit:contain\">" +
      "<div>" +
      "<div class=\"empresa\">TELÉFONOS DE MÉXICO, S.A.B. DE C.V.</div>" +
      "<div class=\"sub\">ALMACÉN DISTRIBUIDOR PUEBLA</div>" +
      "<div class=\"sub\">Esquina Tepeyac, Calz. Ignacio Zaragoza, Los Pinos &nbsp; Tel. 2-57-73-90</div>" +
      "<div class=\"sub\">C.P. 72103 Heroica Puebla de Zaragoza, Pue.</div>" +
      "</div></div>" +
      "<div class=\"folio-box\">" +
      "<div class=\"label\">Guía de embarque</div>" +
      "<div style=\"font-size:15px;font-weight:800;margin:3px 0\">No. Guía &nbsp; " + _guiaActual.folio + "</div>" +
      "<div style=\"font-size:13px;font-weight:700;margin:2px 0\">Pedido: " + (_guiaActual.pedido||"0") + "</div>" +
      "<div style=\"font-size:13px;font-weight:700;margin:2px 0\">Siatel: " + (_guiaActual.siatel||"0") + "</div>" +
      "<div class=\"label\">Área: " + _guiaActual.area + "</div>" +
      "</div></div>" +
      // DESTINATARIO
      "<div class=\"dest\"><table>" +
      "<tr><td><b>Destinatario:</b></td><td>" + alm.nombre + "</td><td><b>Tel.:</b></td><td>" + (alm.tel||"") + "</td></tr>" +
      "<tr><td><b>Atiende:</b></td><td>" + (alm.atiende||"") + "</td><td><b>C.P.:</b></td><td>" + (alm.cp||"") + "</td></tr>" +
      "<tr><td><b>Domicilio:</b></td><td colspan=\"3\">" + (alm.domicilio||"") + " &nbsp; " + (alm.ciudad||"") + "</td></tr>" +
      "</table></div>" +
      // OBSERVACIONES (solo en la primera página, solo si hay texto)
      (esPrimeraPag && _guiaActual.observaciones ?
        "<div class=\"observaciones-box\">" +
        "<div class=\"obs-label\">&#9888; Observaciones</div>" +
        "<div class=\"obs-texto\">" + _escAttr(_guiaActual.observaciones).replace(/\n/g,"<br>") + "</div>" +
        "</div>" : "") +
      // TABLA DE MATERIALES
      "<table class=\"items\">" +
      "<thead><tr>" +
      "<th class=\"col-cant\">Cantidad</th>" +
      "<th class=\"col-emp\">Unidad de embarque</th>" +
      "<th class=\"col-desc\">Descripción</th>" +
      "<th class=\"col-cat\">Catálogo</th>" +
      "<th class=\"col-tot\">Total</th>" +
      "</tr></thead><tbody>" + filasHtml + "</tbody></table>" +
      // TRANSPORTE Y FIRMAS (transporte | Surtió con espacio de firma | Sello)
      "<div class=\"firmas-wrap\">" +
      "<div class=\"transp\"><table>" +
      "<tr><td><b>Línea de transporte:</b></td><td>" + (_guiaActual.transporte||"") + "</td></tr>" +
      "<tr><td><b>Operador:</b></td><td>" + (_guiaActual.operador||"") + "</td></tr>" +
      "<tr><td><b>Tipo de vehículo:</b></td><td>" + (_guiaActual.tipoVeh||"") + "</td></tr>" +
      "<tr><td><b>Placas:</b></td><td>" + (_guiaActual.placas||"") + "</td></tr>" +
      "</table></div>" +
      "<div class=\"surtio-col\">" +
      "<div class=\"surtio-firma-espacio\"></div>" +
      "<div class=\"surtio-label\">Nombre y Firma quien surtió: <b>" + (_guiaActual.surtio||"") + "</b></div>" +
      "</div>" +
      "<div class=\"sello-box\">Sello</div>" +
      "</div>" +
      "<table class=\"firma-linea\">" +
      "<tr><td style=\"padding:32px 6px 5px;font-size:11px;font-weight:700;border-top:2px solid #000\">Fecha y firma &mdash; Transportista</td></tr>" +
      "</table>" +
      "</div>";
  }

  // Generar HTML completo con todas las páginas
  var paginasHtml = "";
  for(var p=0; p<paginas.length; p++){
    var esUltima = (p === paginas.length-1);
    var esPrimera = (p === 0);
    paginasHtml += _htmlPagina(paginas[p], esUltima, esPrimera);

  }

  var guiaHtml =
    "<!DOCTYPE html><html lang=\"es\"><head><meta charset=\"UTF-8\">" +
    "<title>Guía " + _guiaActual.area + " No." + _guiaActual.folio + " &mdash; D041</title>" +
    "<style>" +
    "body{font-family:Arial,sans-serif;font-size:11px;margin:0;padding:0;color:#000;font-weight:500}" +
    ".pagina{padding:12px;}" +
    "h1{font-size:13px;margin:0}" +
    ".hdr{display:grid;grid-template-columns:1fr auto;gap:12px;margin-bottom:8px}" +
    ".empresa{font-size:13px;font-weight:900}.sub{font-size:11px;font-weight:600}" +
    ".folio-box{border:3px solid #000;padding:6px 12px;text-align:right;min-width:160px;font-weight:700}" +
    ".label{font-size:10px;color:#555}" +
    ".dest{border:2px solid #555;padding:6px 10px;margin-bottom:6px;font-size:11px;font-weight:600;page-break-inside:avoid}" +
    ".dest table{width:100%;border-collapse:collapse}" +
    ".dest td{padding:2px 4px;vertical-align:top}" +
    "table.items{width:100%;border-collapse:collapse;margin:6px 0}" +
    "table.items th{background:#001E6E;color:#fff;padding:5px 6px;font-size:11px;font-weight:800;text-align:center;border:2px solid #001E6E}" +
    "table.items td{border:1.5px solid #999;padding:4px 6px;vertical-align:middle;font-weight:600}" +
    ".col-cant{text-align:center;width:60px}.col-emp{width:160px}" +
    ".col-desc{}.col-cat{text-align:center;width:90px;font-weight:800;font-family:monospace}" +
    ".col-tot{text-align:center;width:80px}" +
    ".observaciones-box{border:2px solid #c0392b;border-radius:4px;padding:8px 12px;margin:6px 0 8px;" +
    "background:#fff5f0;page-break-inside:avoid}" +
    ".obs-label{font-size:11px;font-weight:800;color:#c0392b;text-transform:uppercase;letter-spacing:.4px;margin-bottom:3px}" +
    ".obs-texto{font-size:13px;font-weight:700;line-height:1.4;color:#000}" +
    ".firmas-wrap{display:flex;align-items:stretch;margin-top:6px;border:2px solid #555;page-break-inside:avoid}" +
    ".firma-linea{width:100%;border-collapse:collapse;page-break-inside:avoid}" +
    ".firma-linea td{border-left:2px solid #555;border-right:2px solid #555;border-bottom:2px solid #555}" +
    ".transp{flex:1;font-size:11px;font-weight:600}" +
    ".transp table{width:100%;height:100%;border-collapse:collapse}" +
    ".transp td{padding:4px 6px;border-bottom:1px solid #ccc;font-weight:600}" +
    
    ".surtio-col{flex:1;border-left:1px solid #999;padding:6px 10px;display:flex;flex-direction:column;font-size:11px}" +
    ".surtio-firma-espacio{flex:1}" +
    ".surtio-label{border-top:1.5px solid #000;padding-top:3px;font-weight:600}" +
    
    ".sello-box{border-left:2px dashed #444;display:flex;align-items:center;" +
    "justify-content:center;color:#e2e2e2;font-size:9px;width:5cm;height:3.5cm;flex-shrink:0}" +
    
    "@page{size:letter;margin:6mm 9mm}" +
    "@media print{" +
    "html,body{padding:0;margin:0;width:100%}" +
    ".pagina{padding:8px;box-sizing:border-box;border:3px solid #000;margin:0;width:100%;page-break-inside:avoid}" +
    
    "button,.no-print{display:none!important}" +
    "table.items thead{display:table-header-group;-webkit-print-color-adjust:exact;print-color-adjust:exact}" +
    
    "}" +
    "</style></head><body>" +
    paginasHtml +
    "<div class=\"no-print\" style=\"padding:12px\">" +
    "<button onclick=\"window.print()\" style=\"padding:10px 24px;background:#001E6E;color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;margin-right:10px\">&#128424; Imprimir</button>" +
    "<button onclick=\"_guiasDescargarPDF()\" style=\"padding:10px 24px;background:#16a34a;color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer\">&#11015; Descargar PDF</button>" +
    "</div>" +
    "<script>function _guiasDescargarPDF(){var opt={margin:0,filename:'Guia_" + _guiaActual.area + "_No" + _guiaActual.folio + ".pdf',image:{type:'jpeg',quality:0.98},html2canvas:{scale:2,useCORS:true},jsPDF:{unit:'mm',format:'letter',orientation:'portrait'}};var el=document.body.cloneNode(true);el.querySelectorAll('.no-print').forEach(function(b){b.remove();});if(typeof html2pdf!=='undefined'){html2pdf().set(opt).from(el).save();}else{alert('Usa Imprimir → Guardar como PDF');window.print();}}<\/script>" +
    "<script src=\"https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js\"><\/script>" +
    "</body></html>";

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
  _guiasDraftBorrar(); // ya quedó a salvo en el historial, el borrador transitorio ya no hace falta

  // Si veníamos editando una guía existente y cambió folio/área, limpiar el registro anterior
  if(_guiaEditandoOriginal &&
     (_guiaEditandoOriginal.folio !== _guiaActual.folio || _guiaEditandoOriginal.area !== _guiaActual.area)){
    var histLimpio = _guiasHistCargar().filter(function(g){
      return !(g.folio === _guiaEditandoOriginal.folio && g.area === _guiaEditandoOriginal.area);
    });
    _guiasHistGuardar(histLimpio);
  }
  _guiaEditandoOriginal = null;
}

function _guiasBorrarHistorial(idx){
  var hist = _guiasHistCargar();
  var g = hist[idx];
  if(!g) return;
  if(!confirm("¿Borrar la guía No. " + g.folio + " — " + (g.destino||'') + "?\n\nEsta acción no se puede deshacer.")) return;
  hist.splice(idx, 1);
  _guiasHistGuardar(hist);
  modGuias();
}

function _guiasAbrirHistorial(idx){
  var hist = _guiasHistCargar();
  var g = hist[idx];
  if(!g) return;
  if(g.datos){
    // Cargar la guía completa en modo edición (materiales, cabecera, firmas)
    _guiaActual = JSON.parse(JSON.stringify(g.datos));
    _guiaEditandoOriginal = { folio: g.folio, area: g.area };
    _guiasCapturaMateriales();
  } else {
    alert("Guía " + g.area + " No. " + g.folio +
      "\nDestino: " + g.destino +
      "\nFecha: " + g.fecha +
      "\n\nEsta guía se generó con una versión anterior y no tiene datos para reimprimir o editar.");
  }
}

function _guiasReimprimirHistorial(idx, ev){
  if(ev) ev.stopPropagation();
  var hist = _guiasHistCargar();
  var g = hist[idx];
  if(!g) return;
  if(g.datos){
    _guiaActual = JSON.parse(JSON.stringify(g.datos));
    _guiaEditandoOriginal = { folio: g.folio, area: g.area };
    _guiasGenerar();
  } else {
    alert("Guía " + g.area + " No. " + g.folio +
      "\nDestino: " + g.destino +
      "\nFecha: " + g.fecha +
      "\n\nEsta guía se generó con una versión anterior y no tiene datos para reimprimir.");
  }
}

