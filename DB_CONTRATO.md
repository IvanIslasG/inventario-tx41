# Contrato de datos `DB` — Sistema TX41 Puebla

Este documento describe la estructura exacta del objeto `DB` que vive en memoria
una vez descifrado el bundle `datos.enc`. Todos los módulos (`index.html`, `or.js`)
leen de esta misma estructura. Si vas a tocar `publicar.py`, `bundleFinal()` en
`index.html`, o cualquier módulo que lea `DB`, este es el documento de referencia.

**Generado:** Julio 2026
**Mantenido por:** Iván Islas (Telmex Almacen Pue D041)
**Archivos que dependen de este contrato:** `index.html`, `or.js`, `publicar.py`

---

## Visión general

```
DB = {
  meta:              { ... },
  directorio:        { almacenes: { ... } },
  materiales:        { [catalogo]: { ... } },
  existencias:       { [catalogo]: { [almacen]: numero } },
  traslados:         { [catalogo]: { [almacen]: numero } },
  lotes:             { [catalogo]: [ { ... } ] },
  valoracion:        [ catalogo, catalogo, ... ],
  consumos:          { [almacen]: { [catalogo]: numero } },
  nombresGenericos:  { mapa: { [catalogo]: string }, lista: [ string ] },
  criticos:          [ catalogo, catalogo, ... ],
}
```

`DB` se reconstruye completo cada vez que se publica desde Configuración o desde
`publicar.py`. **No es incremental** — cada publicación reemplaza el bundle entero.

---

## `DB.meta`

Metadatos generales del bundle, generados automáticamente en cada publicación.

```js
DB.meta = {
  generado: "2026-06-30T10:48:15",      // ISO timestamp de cuándo se generó
  almacen_distribuidor: "D041",          // sigla del almacén distribuidor
  n_materiales: 5578,                    // conteo informativo, no autoritativo
  n_almacenes: 312,                      // conteo informativo
  n_con_existencia: 1672,                // conteo informativo
  meses_cpm: 6,                          // default para cálculo de CPM en OR
  meses_stock: 1.5,                      // default de stock objetivo en OR
}
```

**Quién lo usa:**
- `DIST()` en `index.html` lee `DB.meta.almacen_distribuidor` — es la fuente de
  verdad de cuál es el almacén distribuidor (actualmente siempre `"D041"`)
- `_modORPanel()` en `or.js` lee `meses_cpm` y `meses_stock` como valores
  default de los inputs CPM/Stock obj.

**Importante:** `n_materiales`, `n_almacenes`, `n_con_existencia` son solo para
mostrar en pantalla ("5,578 materiales · 312 almacenes" en el header). Nunca se
usan para lógica de negocio — siempre se itera sobre las claves reales del
objeto correspondiente.

---

## `DB.directorio.almacenes`

Catálogo de todos los almacenes/centros conocidos por el sistema.

```js
DB.directorio.almacenes = {
  "A08D": {
    desc: "TLAXCALA",
    centro: "RN59",
    recurrente: true,
  },
  "D041": {
    desc: "ALM. DISTRIBUIDOR PUEBLA",
    centro: "RN58",
    recurrente: true,
  },
}
```

| Campo | Tipo | Descripción |
|---|---|---|
| desc | string | Nombre legible del almacén |
| centro | string | Clave del centro SAP (RN51, RN55, RN56, RN57, RN58, RN59) — define la rama de exportación en OR |
| recurrente | boolean | Si es un almacén de uso frecuente — controla el filtro "Almacenes recurrentes" |

**Quién lo usa:**
- `almList()`, `almName(k)` en `index.html` — helpers globales
- Resumen OR en `or.js` lee `centro` para decidir la ruta SAP

**Regla de negocio:** el centro `RN58` corresponde a Puebla — si el almacén
destino de una OR pertenece a RN58, la exportación usa `MIGO 313` (traspaso
interno mismo centro). Cualquier otro centro usa `ME21N + MIGO 351`.

---

## `DB.materiales`

El maestro completo de materiales — fuente de verdad única para qué catálogos
existen en el sistema, sin importar si tienen existencia o consumo.

```js
DB.materiales = {
  "1000291": {
    desc: "LIMPIADOR/DESPLAZADOR HUMEDAD CRC 2.26",
    um: "BTE",
    area: "Misceláneos",
    ubic: "Anaquel 1",
  },
}
```

| Campo | Tipo | Descripción |
|---|---|---|
| desc | string | Descripción del material |
| um | string | Unidad de medida (PZA, BTE, KG, CAR, JGO, etc.) |
| area | string | Herramientas / Misceláneos / Papelería / Cables / Ropa y Calzado / vacío (Sin clasificar) |
| ubic | string | Ubicación física en el almacén. Puede estar vacío |

**Acceso seguro:** usar siempre `mat(cat)` en lugar de `DB.materiales[cat]`:

```js
const mat = c => DB.materiales[c] || {desc:"", um:"", area:"", ubic:""};
```

**Regla de negocio crítica (julio 2026):** tanto Inventario como OR iteran
sobre todos los catálogos de `DB.materiales`, sin filtrar por existencia o
consumo previo. En Inventario permite detectar rezagos físicos no registrados
en SAP. En OR permite monitorear materiales críticos aunque no tengan
consumo histórico.

---

## `DB.existencias`

Existencias por catálogo y por almacén, extraídas de `TOTAL.xlsx` (columna
"Libre utilización").

```js
DB.existencias = {
  "1000291": { "D041": 30, "A08D": 5 },
}
```

Un catálogo solo aparece aquí si tiene existencia > 0 en al menos un almacén.

**Acceso seguro:**
```js
const existencia = (DB.existencias?.[cat]?.[almacen]) || 0;
```

---

## `DB.traslados`

Mismo formato que existencias, columna "Tránsito y traslado" de `TOTAL.xlsx`.
Se suma a la existencia libre para obtener "Total SAP" en conteo físico.

---

## `DB.lotes`

Lotes con número real (no "NUEVO") presentes únicamente en el almacén
distribuidor D041. Se usa para exportación MIGO 351 con clase de valoración.

```js
DB.lotes = {
  "1000337": [
    { lote: "LT2024001", lib: 15, tras: 0 },
  ],
}
```

**Importante:** solo incluye lotes de D041, no de otros almacenes.

---

## `DB.valoracion`

Lista plana de catálogos que tienen algún lote registrado en cualquier
almacén. Bandera rápida para saber si un material requiere clase de
valoración al hacer MIGO.

---

## `DB.consumos`

Historial de consumo mensual por almacén, cargado desde los archivos
TX8A-TX9U en Configuración. Es el único campo que se fusiona (no reemplaza)
entre publicaciones — ver `bundleFinal()`.

```js
DB.consumos = {
  "A08D": { "1000291": 5, "1000295": 44 },
}
```

**Regla de negocio:** el consumo nunca decide si un material aparece en
Inventario u OR — solo se usa para calcular CPM (`consumo / meses_cpm`). Antes
de julio 2026 la presencia de consumo filtraba la tabla, ocultando materiales
críticos sin historial; esto se corrigió.

**Acceso seguro:**
```js
const cons = DB.consumos?.[almacen] || {};
const consumoDelCat = cons[cat] || 0;
```

---

## `DB.nombresGenericos`

Agrupación de catálogos por "nombre genérico" — el criterio para identificar
sustitutos entre sí.

```js
DB.nombresGenericos = {
  mapa: { "1000291": "LIMPIADOR DE CONTACTOS" },
  lista: ["LIMPIADOR DE CONTACTOS", "PILA AA"],
}
```

**Acceso seguro:**
```js
function ngDe(cat){ return (DB.nombresGenericos?.mapa||{})[cat] || ""; }
```

Un catálogo sin nombre genérico se trata como `"SIN SUSTITUTO"`. Solo los
materiales SIN SUSTITUTO reciben sugerencia automática de X Surtir; los que
tienen sustituto requieren decisión manual (para no duplicar pedido entre
variantes del mismo grupo).

---

## `DB.criticos`

Lista plana de catálogos marcados como críticos para el negocio (504
actualmente).

**Importante (revisión julio 2026):** este campo ya no determina qué aparece
en la tabla de OR — desde el cambio a "todos los materiales del maestro",
`DB.criticos` quedó sin uso activo en `calcularOR()`. Se conserva en el
bundle por si se reintroduce un filtro "Solo críticos" o para reportes de
monitoreo. Antes de eliminarlo del pipeline, confirmar que no hay un módulo
dependiendo de él.

---

## Cómo se construye DB (pipeline de publicación)

1. `publicar.py` (offline) lee TOTAL.xlsx, Almacenes_Hana.xlsx,
   Base_Maestra_General_41.xlsx, los TSV de consumo, NombresGenericos.xlsx y
   criticos.xlsx, arma el DB completo, lo cifra con AES-GCM/PBKDF2 (clave
   `TX41-Puebla-2026`) y genera `datos.enc`.

2. El módulo Configuración (en el navegador) permite actualizar solo
   existencias/traslados/lotes subiendo un nuevo TOTAL.xlsx, vía
   `parseTotalJS()` + `construirBundleJS()`. Esta ruta NO relee el maestro de
   materiales — toma `DB.materiales` tal como está en memoria y solo completa
   desc/um si faltan. Para reflejar cambios en el maestro (área, ubicación,
   nuevos catálogos) siempre hay que pasar por `publicar.py`.

3. El resultado se cifra en el navegador (Web Crypto, misma clave) y se
   puede publicar directo a GitHub Pages con un token fine-grained, o
   descargar el datos.enc para subirlo manualmente.

**Regla de oro:** si modificas el maestro de materiales, el cambio solo se
refleja corriendo `publicar.py` desde el equipo que tiene el Excel. La app
web nunca lee el maestro desde disco — solo confía en lo que ya viene en el
DB actual en memoria.

---

## Checklist al modificar este contrato

Si vas a agregar un campo nuevo a DB:

- [ ] Actualizar `publicar.py` para que lo incluya al generar el bundle
- [ ] Actualizar `bundleFinal()` / `construirBundleJS()` en index.html para
      que la ruta de publicación desde Configuración también lo preserve
- [ ] Documentar el campo en este archivo
- [ ] Usar siempre acceso seguro (`?.` y `|| default`) — nunca asumir que el
      campo existe, porque bundles antiguos en caché de usuarios pueden no
      tenerlo todavía
