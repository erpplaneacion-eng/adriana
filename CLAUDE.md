# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué hace este proyecto

Sistema para llenar automáticamente el **INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx** (multi-hoja) con datos contables mensuales extraídos de archivos XLS fuente. El mapeo fuente→destino vive en `chvs.db` (SQLite) y se edita visualmente con el Config Builder (Flask). `config_gastos.json` y `config_5105.json` solo se usan como fuente de migración inicial; la DB es la fuente de verdad en producción.

## Comandos principales

```bash
# Iniciar Config Builder (UI web para configurar mapeos)
py app.py
# Abre http://localhost:5000

# Procesar un mes (escribe valores en el INFORME REAL)
py procesar_todo.py MARZO
py procesar_todo.py FEBRERO-DL
```

## Arquitectura

### Flujo general
```
Archivos XLS mensuales (FEBRERO-DL/, MARZO/, etc.)
        ↓  leídos por
procesar_todo.py  ←  config_gastos.json  (qué código de qué archivo va a qué fila)
        ↓  escribe en
INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx  (archivo destino, se sobreescribe)
```

### Archivos clave

| Archivo | Rol |
|---|---|
| `app.py` | Flask backend del Config Builder. Sirve la UI y expone los endpoints API. |
| `db.py` | Capa de acceso a `chvs.db` (SQLite). `save_config/load_config`, `save_config_5105/load_config_5105`, `log_ejecucion`, `get_historial`, `diff_ejecuciones`. **`save_config()` NO borra `fuentes_valor`/`valores_escritos`** (son tablas de auditoría, no de config). |
| `db_postgres.py` | Capa de acceso a Postgres (Railway). Dos tablas: `chips_xls` (caché de datos XLS por mes) e `informe_resultados` (valores escritos al INFORME por mes). Conexión vía `DATABASE_URL`. |
| `procesar_todo.py` | Script principal. 4 pasos en `procesar_mes()`: (1) gastos ER + aux (todas las hojas), (2) salarios 5105 — corre después para no ser sobreescrito, (3) fórmulas cross-sheet en GASTOS ADMIN filas 89 y 114, (4) fórmulas cross-sheet en hojas 3-col/mes filas 148-151. También expone `calcular_preview()` y `_construir_archivos_cache()`. Al terminar una ejecución exitosa guarda los resultados en Postgres. **Nota**: el docstring de nivel de módulo del archivo lista el orden al revés (5105 primero) — ignorarlo, `procesar_mes()` es la fuente de verdad. |
| `procesar_gastos.py` | Script standalone de gastos (**legado**). BASE hardcodeada, apunta a `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx`. |
| `procesar_5105.py` | Script standalone de nómina (**legado**). BASE hardcodeada. |
| `config_gastos.json` | Solo para migración inicial a `chvs.db`. No es la fuente de verdad en producción. |
| `config_5105.json` | Solo para migración inicial a `chvs.db`. |
| `static/app.js` | Toda la lógica frontend: drag & drop, `mappings`, `buildConfig`, `renderChips`, modal de preview, panel de historial. |
| `INFORME REAL_2026...xlsx` | Plantilla/destino. **Nunca abrir mientras corre el script.** |

### Endpoints API (app.py)

| Endpoint | Método | Descripción |
|---|---|---|
| `/api/sheets` | GET | Hojas del INFORME REAL con filas (col A + B), filtrando fórmulas internas |
| `/api/folders` | GET | Carpetas mensuales disponibles en BASE con conteo de XLS |
| `/api/files/<folder>` | GET | Chips de todos los XLS de la carpeta. Usa caché Postgres; si no existe, lee XLS y guarda. |
| `/api/config` | GET | Lee config desde `chvs.db` |
| `/api/config/save` | POST | Guarda config en `chvs.db` (reemplaza tablas de config, preserva historial) |
| `/api/run/<folder>` | POST | Ejecuta procesar_todo.py para la carpeta indicada |
| `/api/preview/<folder>` | GET | Calcula todos los valores sin escribir en el Excel; retorna `{ok, items}` |
| `/api/historial` | GET | Últimas 50 ejecuciones desde `chvs.db` |
| `/api/historial/<mes>` | GET | Ejecuciones filtradas por mes |
| `/api/historial/diff` | GET | Compara dos ejecuciones: `?a=<id>&b=<id>` → lista de deltas por (hoja, fila) |
| `/api/upload/<folder>` | POST | Sube archivos XLS a una carpeta mensual en el servidor |
| `/api/upload/informe` | POST | Reemplaza el INFORME REAL base con el archivo subido |
| `/api/download/informe` | GET | Descarga el INFORME REAL actualizado |
| `/api/resultados/<mes>` | GET | Valores escritos al INFORME para ese mes desde Postgres. Si no hay datos, calcula preview y los guarda. |
| `/api/chips/reset/<mes>` | POST | Borra el caché de chips XLS de ese mes en Postgres. Al recargar `/api/files/<mes>` se regenera. |

### Persistencia — SQLite (`db.py`) + Postgres (`db_postgres.py`)

**SQLite `chvs.db`** — config y auditoría. La config vive aquí con 8 tablas:

| Tabla | Contenido |
|---|---|
| `file_sources` | Mapa `key → {prefijo, entidad}` |
| `dest_rows` | Cada ítem destino: `hoja, codigo_a, fila_dest, buscar_por` |
| `sources` | Sources por dest_row: `file_key, codes, op, sin_filtro, seccion, orden` |
| `valor_fijo` | Operaciones hardcodeadas por dest_row |
| `config_5105` | Reglas por entidad 5105: `only_code, subtract_codes` |
| `ejecuciones` | Log de cada corrida: `mes, fecha, exito, notas` |
| `valores_escritos` | Valores escritos por ejecución: `hoja, fila, codigo_a, valor_total` |
| `fuentes_valor` | Desglose de fuentes por valor escrito |

`init_db()` crea el schema. `migrate_from_json()` migra los JSON legacy la primera vez. `save_config()` solo borra `file_sources`, `dest_rows`, `sources`, `valor_fijo` — **nunca** `ejecuciones`, `valores_escritos` ni `fuentes_valor`.

**Postgres (Railway)** — caché de lectura y resultados. Conexión vía env var `DATABASE_URL`. Creado por `init_pg()` al arrancar `app.py`. Dos tablas:

| Tabla | Contenido |
|---|---|
| `chips_xls` | Caché de los items leídos de cada XLS: `(mes, archivo, idx)` PK. Campos: `file_key, seccion, codigo, descripcion, valor, debito, credito, op_jk, negrita, indent, separador`. Se puebla la primera vez que se carga la carpeta; se invalida con `POST /api/chips/reset/<mes>`. |
| `informe_resultados` | Valores escritos al INFORME por ejecución: `mes, hoja, fila, codigo_a, valor_total, fuentes (JSONB), ejecutado_en`. Se reemplaza completo en cada ejecución exitosa del mes. |

`file_key` en `chips_xls` usa el mismo criterio que `fileKey()` en app.js (ver sección abajo) — `ER_CHVS`, `AUX_CHVS`, `7205_UT_VSOL_BUGA2025`, `51055101_DOT_CHVS`, etc. **No** usar el filename raw como clave.

### Resolución de archivos — `_construir_archivos_cache()` (procesar_todo.py)

Helper que construye `{key: ruta_absoluta}` en 3 etapas, en orden:

1. **file_sources lookup**: para cada key en config, llama `encontrar_archivo(carpeta, prefijo, entidad, mes, anio)`.
2. **Auto-discovery**: escanea todos los `.xls` de la carpeta → genera key a partir del nombre sin `_MES_YYYY` normalizando espacios dobles (`re.sub(r'\s+', ' ', base).replace(' ', '_')`). Solo registra si el key aún no tiene ruta.
3. **Alias sin año**: para keys que siguen sin ruta, busca keys existentes que empiecen con ese prefijo y asigna la ruta del que queda último en orden alfabético (i.e., el año más reciente).

`encontrar_archivo()` intenta tres patrones en orden:
1. **Exacto**: `^{prefijo}_{entidad}_{mes}_{anio}.*\.xls`
2. **Flexible** (cualquier año): `^{prefijo}_{entidad}_\d{4}.*\.xls`
3. **Prefijo parcial** (`\d*` tras la entidad): `^{prefijo}_{entidad}\d*_{mes}_\d{4}.*\.xls` — cubre entidades con año pegado al nombre (ej. `CONS ALIM CALI2026`)

Esto cubre archivos de contratos anteriores (e.g., `7205_UT BUGA2025_MAR_2025.xls`) presentes en carpetas de 2026.

Tanto `procesar_gastos()` como `calcular_preview()` llaman a `_construir_archivos_cache()` — garantiza que ambas rutas producen resultados idénticos para los mismos archivos.

### Preview sin escritura — `calcular_preview()` (procesar_todo.py)

```python
items = calcular_preview(carpeta_mes)
# items: [{hoja, fila, codigo_a, valor_total, advertencias, fuentes}]
```

Reutiliza `_construir_archivos_cache()`, `leer_er()` y `leer_aux()` sin abrir ni guardar el Excel destino. Útil para validar valores antes de ejecutar. Llamado por `/api/preview/<folder>`.

### `inferEntidad()` en app.js

Extrae la entidad de un nombre de archivo o chip key. Tiene dos ramas:

1. **Normal** (`_MES_YYYY` presente): `^[^_]+_(.+?)_[A-Z]{3}_\d{4}` → captura la entidad del nombre de archivo.
2. **Fallback** (chip key sin mes, p.ej. `7205_CONS_ALIM_CALI`): `^(?:ER|AUX|5105|7205|7105|51355001)_(.+)$` → convierte `_` a espacios.

Sin este fallback, `file_sources` almacenaba `entidad=''` → `encontrar_archivo` construía patrón con doble guion bajo → archivo nunca encontrado.

### Tipos de archivos XLS fuente (en carpetas mensuales)

- `ESTADO DE RESULTADOS_<entidad>_<mes>_<año>.xls` — estado de resultados por unidad. Se lee con `leer_er()`: busca códigos en sección "99 GENERAL".
- `51355001_<entidad>_<mes>_<año>.xls` — auxiliar de proveedores. Se lee con `leer_aux()`. Sin secciones (NITs planos).
- `5105_<entidad>_<mes>_<año>.xls` — nómina. Se lee con `leer_5105()`. Tiene secciones por contrato.
- `7205_<entidad>_<mes>_<año>.xls` — costos de personal auxiliar, se lee con `leer_aux()`. Tiene secciones por contrato.
- `7105_<entidad>_<mes>_<año>.xls` — costos de raciones/alimentos, se lee con `leer_aux()`. Tiene secciones por contrato (misma estructura que 7205).
- `51055101 DOT_<entidad>_<mes>_<año>.xls` — dotación (costo 51055101). Se lee con `leer_aux()`. Prefijo con espacio: `"51055101 DOT"`.
- `72055101 DOT_<entidad>_<mes>_<año>.xls` — dotación operativa (costo 72055101). Se lee con `leer_aux()`. Prefijo con espacio: `"72055101 DOT"`.

**Importante**: `51055101 DOT` empieza con `5105` y `72055101 DOT` empieza con `7205`. Tanto `fileKey()` como `inferPrefijo()` y `_nombre_a_key()` los detectan **antes** del check genérico de `5105`/`7205` para evitar colisión de prefijo.

### Estructura de columnas en archivos XLS fuente

`detectar_col_valor()` busca en las primeras 25 filas el encabezado "Debitos" (modo `debito`) o "Netos" (modo `netos`, fallback). `detectar_col_creditos()` busca "Creditos" en las mismas filas.

`leer_aux()` y `leer_er()` **aplican resta J−K** (`Débitos - Créditos`) cuando encuentran ambas columnas en modo `debito`. Si no hay columna de Créditos, usan solo Débitos. El campo `op_jk: true` en los items del panel derecho indica que esa fila tuvo resta aplicada.

Los chips en la UI muestran el valor neto (J−K o solo J) según lo detectado en el archivo.

### Estructura de columnas en el INFORME REAL

- **Hojas estándar** (GASTOS ADMINISTRATIVOS, GASTOS OPERATIVOS, GASTOS VEHICULOS, etc.): columna = `MES_COLUMNA[mes]` → ENE=3, FEB=4, MAR=5 ... DIC=14
- **Hojas con 3 cols/mes**: columna = `mes_num * 3` → ENE=3, FEB=6, MAR=9...
  - Lista completa: CASINO, CALI, YUMBO, BUGA, COMEDORES CALI, COMEDORES PALMIRA, COMEDORES VALLE, CTAS EN PPACION, PYG TOTAL, RECREARTE
  - **`COMEDORES VALLE`** está en `COLUMNA_HOJA` (usa patrón 3-cols) pero **no** tiene entrada en `FORMULAS_CROSS_SHEET` — sus filas 148-151 no se escriben automáticamente.
  - En estas hojas, las columnas intermedias (4,5,7,8...) son de ratio/presupuesto, NO de valor mensual.
- Añadir una hoja al patrón 3-cols: agregar en `COLUMNA_HOJA` en `procesar_todo.py` y en `HOJAS_3COL` en `app.py`.

### config_gastos.json — campos por item

```json
{
  "codigo_a": "51055101",       // clave del destino (col A del Excel o "B:descripcion")
  "fila_dest": 8,               // fila Excel exacta (evita colisiones de código duplicado)
  "buscar_por": "B",            // "A" (default) o "B" (buscar por col B cuando no hay código)
  "hoja": "GASTOS ADMINISTRATIVOS", // hoja destino
  "sources": [
    {
      "key": "ER_CHVS",
      "codes": ["51055101"],
      "op": "-",          // opcional: "+" (default), "-", "*", "/" sobre el acumulado
      "sin_filtro": true, // opcional: suma todos los valores del archivo (ignora codes)
      "seccion": "PAECAL20 CONTR. CONSORCIO ALIMENTANDO  CALI 2026"  // opcional: sección del archivo
    }
  ],
  "valor_fijo": [{ "op": "+", "valor": 500000 }]  // operaciones hardcodeadas opcionales
}
```

`file_sources` mapea cada `key` a `{ "prefijo": "ESTADO DE RESULTADOS", "entidad": "CHVS" }` para que `encontrar_archivo()` localice el XLS en la carpeta del mes.

### config_5105.json — estructura por entidad

```json
{
  "CHVS":           { "subtract_codes": ["010103", "010303", "010501"] },
  "UT ALIM BUGA2026": { "subtract_codes": [] },
  "UT BUGA2025":    { "only_code": "010403" }
}
```

- `subtract_codes`: toma el total del código `5105` en la sección "99 GENERAL" y resta los subcódigos indicados.
- `only_code`: usa solo el valor de ese código (ignora el total 5105).
- La clave de cada entidad es el nombre extraído del nombre de archivo: `5105_<entidad>_<mes>_<año>.xls`.
- Si la entidad no aparece en el config, el archivo se omite (`[OMITIDO]` en el log).

**Claves `key` generadas por `fileKey()` en app.js y `_nombre_a_key()` en app.py (idénticas):**
- `ER_<ENTIDAD>` → archivos ESTADO DE RESULTADOS
- `AUX_<ENTIDAD>` → archivos 51355001
- `51055101_DOT_<ENTIDAD>` → archivos 51055101 DOT (detectado antes que 5105)
- `72055101_DOT_<ENTIDAD>` → archivos 72055101 DOT (detectado antes que 7205)
- `5105_<ENTIDAD>` → archivos 5105
- `7205_<ENTIDAD>` → archivos 7205
- `7105_<ENTIDAD>` → archivos 7105

Espacios en la entidad se convierten a `_`. Ambas funciones deben producir keys idénticos para el mismo archivo — si difieren, los chips del panel izquierdo no mostrarán valores al cargar la carpeta.

### rowKey en el frontend (app.js)

El key único de cada fila en `mappings` es `"${sheetName}::${codeKey}|${filaNum}"`. Al guardar, el nombre de hoja y `|filaNum` se extraen para guardar como `hoja` y `fila_dest` respectivamente. La función `parseRowKey()` hace el parsing inverso.

### Panel derecho — tabla de filas fuente (app.js)

El acordeón del panel derecho muestra los archivos XLS como una tabla con columnas **Código | Descripción | Valor**, respetando el orden y la indentación del Excel original.

- `renderDraggableChips(container, nombreArchivo, info)` — construye la tabla de filas `.source-row`.
- Filas en **negrita** (totales) → clase `.chip-total` (fondo ámbar).
- Filas normales (detalle) → fondo blanco.
- Barra de filtros: **Todos / Detalle / Totales** + buscador de texto.
- Selección múltiple: clic en filas las agrega a `selectedChips (Map)`. Arrastrar cualquier fila seleccionada envía el array completo como payload.
- `chipId = \`${fk}::${item.seccion || ''}::${item.codigo}\`` — clave única que incluye la sección para permitir que dos filas con el mismo código pero distinta unidad de negocio sean independientes.

### Secciones por contrato en archivos fuente (app.py)

`leer_todos_codigos()` detecta secciones en tres tipos de archivos:

**Archivos 5105:** busca filas sin código numérico dentro del bloque `5105` → separadores de sección.

**Archivos 7205 y 7105 (`es_seccionado`):** detecta cabeceras de sección como filas que:
- No tienen espacio inicial en col A
- No son un código puro de dígitos (`^\d{3,8}$`)
- Ejemplos: `99 GENERAL`, `PAECAL20 CONTR. CONSORCIO...`, `COMCAL2025-VI...`

En todos los casos:
- Cada ítem lleva el campo `seccion` con el nombre del contrato al que pertenece.
- Los separadores se renderizan como barras azules (`.source-row-sep`) en el panel derecho.
- Permite que el mismo código (ej. `029903 CASOS ESPECIALES`) aparezca en dos secciones distintas y sea arrastrable de forma independiente.

`leer_aux()` en `procesar_todo.py` respeta el campo `seccion`: cuando está presente y el archivo es 7205 o 7105, restringe la búsqueda solo a las filas de esa sección (entre su cabecera y la siguiente).

**Regla de matching en `leer_aux()`:**
- Códigos **con descripción** (contienen espacios): solo exact match → evita capturar filas de detalle por NIT.
- Códigos **puramente numéricos** (sin espacios): también acepta `startswith` → captura filas `"029903 NIT12345"` cuando el código buscado es `"029903"`.

### Operaciones entre chips (campo `op` en sources)

Cada chip en el panel izquierdo tiene un botón `+` (verde) o `−` (rojo). Al hacer clic alterna entre sumar y restar ese chip al total de la fila.

- El campo `op` se guarda en `config_gastos.json` dentro de cada source (solo se escribe si no es `"+"`; el `"+"` es el default para no inflar el JSON).
- `procesar_todo.py` aplica el `op` de cada source al calcular `valor_total`. Soporta `+`, `-`, `*`, `/`.
- El `Σ` del total en la UI refleja las restas en tiempo real.
- Retrocompatibilidad: sources sin campo `op` se tratan como `"+"`.

### Persistencia de valores y sección en mappings (app.js)

`actualizarValoresEnMappings(data)` — se llama al cargar una carpeta. Cruza los datos XLS cargados con los `mappings` existentes para poblar `src.valor` sin necesidad de reconfigurar:
- Retrocompatibilidad: si el código guardado en config tiene formato antiguo combinado (`"010102           GERENCIA"`), extrae la parte numérica con regex antes de buscar en el lookup.
- Para `sin_filtro: true` suma todos los valores del archivo.

`buildConfig()` — al guardar, persiste el campo `seccion` de cada source entry en el JSON.

`initMappingsFromConfig()` — al cargar, restaura el campo `seccion` desde el JSON, lo que permite que el badge de unidad de negocio (`.chip-sec`, cursiva morada) aparezca en los chips del panel izquierdo tras reabrir la app.

### config_gastos.json — campo seccion en sources

```json
"sources": [
  { "key": "5105_CHVS",  "codes": ["010304"], "seccion": "PAEBUGA08 CONTRATO UT BUGA 2026" },
  { "key": "7205_CHVS",  "codes": ["029903           CASOS ESPECIALES"], "seccion": "99 GENERAL" },
  { "key": "7205_CHVS",  "codes": ["029903           CASOS ESPECIALES"], "seccion": "PAECAL20 CONTR. CONSORCIO ALIMENTANDO  CALI 2026" }
]
```

El campo `seccion` es opcional. Se graba cuando el chip proviene de una sección específica de archivos 5105, 7205 o 7105. Sin él, `leer_aux` busca en todo el archivo.

### Exclusiones del Config Builder (app.py)

`FILAS_EXCLUIR` define por hoja qué filas NO se muestran en la UI:
- Fórmulas internas `=SUM(...)` (totales/subtotales que el Excel calcula solo)
- Porcentajes auto-calculados sobre ingresos
- Filas cross-sheet (referencian otras hojas del mismo libro vía `'Hoja'!celda`)
- Títulos sin valor

La detección automática de fórmulas (`es_formula`) en `app.py` excluye filas cuya columna de valor contiene una fórmula que:
- NO tiene referencias a libros externos (`[n]`)
- NO tiene referencias a otras hojas (`'NombreHoja'!`)
- SÍ referencia otras filas de la misma hoja

Para hojas de 3 cols/mes (ver lista completa en `HOJAS_3COL`) solo se revisan las columnas de valor (3,6,9,12...), no las intermedias de ratio/presupuesto.

Las filas cross-sheet que deben excluirse pero no son detectadas automáticamente se listan explícitamente en `FILAS_EXCLUIR` por hoja.

`HOJAS_EXCLUIR` = hojas que no se muestran en la UI (`%DIST C`, `%DIST C `).
`FILAS_EXCLUIR_GLOBAL` = filas excluidas en todas las hojas (encabezados corporativos: CORPORACIÓN HACIA UN VALLE SOLIDARIO, DIAS DE ATENCIÓN).
`TITULOS_SECCION` = filas de fórmula que se renderizan como **separador visual de sección** (barra azul con el nombre en mayúsculas, sin chips ni operaciones). Las filas de totales (TOTAL..., UTILIDAD..., COSTO NETO, etc.) de esas mismas hojas siguen ocultas. Actualmente cubre: GASTOS ADMINISTRATIVOS, GASTOS VEHICULOS, CASINO, CALI, YUMBO, COMEDORES CALI, COMEDORES PALMIRA, CTAS EN PPACION, PYG TOTAL. Para agregar una hoja: añadir entrada en `TITULOS_SECCION` en `app.py`.

### Limpieza de columna antes de escribir (procesar_todo.py)

Antes de escribir valores, `procesar_gastos()` limpia la columna del mes en **todas las hojas del libro** (no solo las que tienen items en el config). Esto evita residuos cuando una hoja pierde todos sus chips entre ejecuciones.

Solo se borran:
- Valores numéricos (escritos por el script en ejecuciones anteriores)
- Fórmulas cross-sheet que contienen `!` (escritas por el script)

Se preservan:
- Fórmulas internas de totales (`=SUM(...)`, `=E5+E13+...`) — no contienen `!`

### Fórmulas cross-sheet escritas por procesar_todo.py

Además de los valores numéricos, `procesar_mes()` escribe fórmulas directamente:

- **GASTOS ADMINISTRATIVOS fila 89**: `=CASINO!${col_3}$144` (ref a Casino)
- **GASTOS ADMINISTRATIVOS fila 114**: `=RECREARTE!${col_3}$141` (ref a Recrearte)
- **Hojas 3-col/mes filas 148-151**: fórmulas cross-sheet hacia `%DIST` y `GASTOS ADMINISTRATIVOS`, definidas en `FORMULAS_CROSS_SHEET`. Ver fila 149 de CALI como caso especial: contiene constantes hardcodeadas distintas por mes que no pueden restaurarse automáticamente.

Las columnas en estas fórmulas usan tres patrones distintos:
- `%DIST` col par: `chr(64 + 2*mes_num)` → ENE=B, FEB=D, MAR=F...
- `%DIST` col impar: `chr(65 + 2*mes_num)` → ENE=C, FEB=E, MAR=G...
- GASTOS ADMIN col estándar: `chr(66 + mes_num)` → ENE=C, FEB=D, MAR=E...

### Celdas fusionadas (merged cells)

En `procesar_gastos` y `procesar_5105`, antes de escribir se detecta si la celda pertenece a un rango fusionado y se redirige a `min_row/min_col` del rango.

### Preview modal (app.js / index.html)

Botón **"🔍 Vista previa"** en el footer llama `GET /api/preview/<folder>`. Muestra un modal con una tabla agrupada por hoja: columnas **Fila | Código | Valor**. Filas con advertencias (archivo no encontrado, error de lectura) se resaltan en rojo. El modal incluye botón "▶ Ejecutar mes" que dispara el flujo normal sin cerrar y reabrir. `confirmarEjecutar()` cierra el modal y llama `runMes()`.

### Panel historial (app.js / index.html)

Sección colapsable debajo del panel de log (`.hist-panel`). Al expandir carga `GET /api/historial` y muestra las últimas ejecuciones en tabla: **Fecha | Mes | Estado | Items**. Click en una fila expande el detalle de valores escritos. Botón "Comparar con anterior" llama `GET /api/historial/diff?a=N&b=N-1` y muestra tabla de diferencias con indicadores ▲/▼ coloreados (`delta-pos`/`delta-neg`).

## Rutas — desarrollo vs. Railway

`BASE` ya **no** está hardcodeada. Ambos scripts la resuelven dinámicamente:

| Contexto | Valor de BASE |
|---|---|
| Desarrollo local | directorio del script (`os.path.dirname(__file__)`) |
| Railway (producción) | variable de entorno `DATA_DIR` (volumen persistente) |

En `app.py`: `BASE = os.environ.get('DATA_DIR', _APP_DIR)`
En `procesar_todo.py`: `BASE = os.environ.get('CHVS_BASE', os.path.dirname(__file__))`

`app.py` pasa `CHVS_BASE=BASE` al subprocess de `procesar_todo.py` para que ambos trabajen sobre el mismo directorio.

**Nota**: `procesar_gastos.py` y `procesar_5105.py` (scripts legado) aún tienen `BASE` hardcodeada y apuntan a un archivo diferente. No usar en Railway.

## Deploy en Railway

### Archivos de deployment

| Archivo | Propósito |
|---|---|
| `requirements.txt` | Dependencias Python (Flask, openpyxl, xlrd, gunicorn, psycopg2-binary) |
| `Procfile` | Comando de arranque: `gunicorn app:app --bind 0.0.0.0:$PORT` |
| `.gitignore` | Excluye `__pycache__`, `.env`, backups de config |

### Pasos para hacer deploy

1. En Railway, crear nuevo proyecto desde el repositorio GitHub
2. Configurar un **volumen persistente** montado en `/data`
3. Agregar la variable de entorno `DATA_DIR=/data`
4. Railway detecta automáticamente el `Procfile` y despliega

### Cómo funciona en producción

**En el primer arranque**, `app.py` copia automáticamente al volumen `/data`:
- `INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx` (plantilla)
- `config_gastos.json` y `config_5105.json`

**Flujo mensual (ejemplo: ABRIL):**

1. Abrir la app en el navegador (`https://tu-app.railway.app`)
2. En el panel derecho → **"Subir archivos al servidor"** → escribir `ABRIL` como nombre de carpeta
3. Seleccionar todos los archivos XLS de abril → clic **Subir archivos**
   - Los archivos se guardan en `/data/ABRIL/` en el volumen de Railway
4. Seleccionar `ABRIL` en el selector de carpeta para ver los códigos disponibles
5. Configurar/verificar los mapeos con drag & drop
6. Clic **Guardar config**
7. En el footer → seleccionar `ABRIL` → clic **🔍 Vista previa** para verificar los valores antes de escribir, luego **▶ Ejecutar mes** (o ejecutar directamente desde el modal de preview)
   - El servidor corre `procesar_todo.py ABRIL` y escribe en `/data/INFORME REAL_2026...xlsx`
8. Clic **⬇ Descargar INFORME REAL** para bajar el Excel actualizado

**El archivo de salida** es siempre `/data/INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx` (el mismo que en desarrollo). Al ejecutar un mes se sobreescribe con los valores de ese mes. Para comparar meses distintos, descarga el archivo antes de procesar el siguiente.

### Variables de entorno en Railway

| Variable | Valor | Obligatoria |
|---|---|---|
| `DATA_DIR` | `/data` | Sí (para persistencia del volumen) |
| `DATABASE_URL` | `postgresql://...` | Sí (inyectada automáticamente por el servicio Postgres de Railway) |
| `PORT` | asignado por Railway | Auto |
| `RAILWAY_ENVIRONMENT` | asignado por Railway | Auto (desactiva debug mode) |

### Caché Postgres — operaciones de mantenimiento

```bash
# Invalidar caché de chips de un mes (p.ej. si se subieron nuevos XLS)
POST /api/chips/reset/MAYO
# → luego GET /api/files/MAYO regenera automáticamente

# Ver qué valores quedaron escritos al INFORME para un mes
GET /api/resultados/MAYO
# → si no hay datos guardados, calcula preview y los persiste
```

Si los chips del panel derecho muestran valores incorrectos al cambiar de mes, probablemente el caché tiene datos viejos. Reset + recarga resuelve.
