# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué hace este proyecto

Sistema para llenar automáticamente el **INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx** (multi-hoja) con datos contables mensuales extraídos de archivos XLS fuente. El mapeo fuente→destino vive en `config_gastos.json` y se edita visualmente con el Config Builder (Flask).

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
| `procesar_todo.py` | Script principal. 4 pasos: salarios (5105), gastos (ER + aux), fórmulas cross-sheet en GASTOS ADMIN, fórmulas cross-sheet en hojas 3-col/mes. |
| `procesar_gastos.py` | Script standalone de gastos (**legado**). BASE hardcodeada, apunta a `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx`. |
| `procesar_5105.py` | Script standalone de nómina (**legado**). BASE hardcodeada. |
| `config_gastos.json` | Mapeo principal. Cada item: `codigo_a`, `sources`, `hoja`, `fila_dest`, `buscar_por`, `valor_fijo`. Se respalda automáticamente al guardar. |
| `config_5105.json` | Reglas para archivos de nómina 5105 por entidad |
| `static/app.js` | Toda la lógica frontend: drag & drop, `mappings`, `buildConfig`, `renderChips` |
| `INFORME REAL_2026...xlsx` | Plantilla/destino. **Nunca abrir mientras corre el script.** |

### Endpoints API (app.py)

| Endpoint | Método | Descripción |
|---|---|---|
| `/api/sheets` | GET | Hojas del INFORME REAL con filas (col A + B), filtrando fórmulas internas |
| `/api/folders` | GET | Carpetas mensuales disponibles en BASE con conteo de XLS |
| `/api/files/<folder>` | GET | Códigos extraídos de todos los XLS en la carpeta |
| `/api/config` | GET | Lee config_gastos.json |
| `/api/config/save` | POST | Guarda config (hace backup automático por fecha) |
| `/api/run/<folder>` | POST | Ejecuta procesar_todo.py para la carpeta indicada |
| `/api/upload/<folder>` | POST | Sube archivos XLS a una carpeta mensual en el servidor |
| `/api/upload/informe` | POST | Reemplaza el INFORME REAL base con el archivo subido |
| `/api/download/informe` | GET | Descarga el INFORME REAL actualizado |

### Tipos de archivos XLS fuente (en carpetas mensuales)

- `ESTADO DE RESULTADOS_<entidad>_<mes>_<año>.xls` — estado de resultados por unidad. Se lee con `leer_er()`: busca códigos en sección "99 GENERAL".
- `51355001_<entidad>_<mes>_<año>.xls` — auxiliar de proveedores. Se lee con `leer_aux()`. Sin secciones (NITs planos).
- `5105_<entidad>_<mes>_<año>.xls` — nómina. Se lee con `leer_5105()`. Tiene secciones por contrato.
- `7205_<entidad>_<mes>_<año>.xls` — costos de personal auxiliar, se lee con `leer_aux()`. Tiene secciones por contrato.
- `7105_<entidad>_<mes>_<año>.xls` — costos de raciones/alimentos, se lee con `leer_aux()`. Tiene secciones por contrato (misma estructura que 7205).

### Estructura de columnas en archivos XLS fuente

Todos los archivos auxiliares (7205, 7105, 51355001) tienen 4 columnas de valor:
- **Col I** (idx 8): Saldo inicial
- **Col J** (idx 9): Débitos ← columna principal (`col_val`)
- **Col K** (idx 10): Créditos (no se resta; se usa solo col J)
- **Col L** (idx 11): Saldo final

`leer_aux()` y `leer_er()` leen únicamente la columna `col_val` (col J, Débitos). No se aplica resta J−K.

Los chips en la UI muestran el valor de col J tal como está en el archivo.

### Estructura de columnas en el INFORME REAL

- **Hojas estándar** (GASTOS ADMINISTRATIVOS, GASTOS OPERATIVOS, GASTOS VEHICULOS, etc.): columna = `MES_COLUMNA[mes]` → ENE=3, FEB=4, MAR=5 ... DIC=14
- **Hojas con 3 cols/mes**: columna = `mes_num * 3` → ENE=3, FEB=6, MAR=9...
  - Lista completa: CASINO, CALI, YUMBO, BUGA, COMEDORES CALI, COMEDORES PALMIRA, COMEDORES VALLE, CTAS EN PPACION, PYG TOTAL, RECREARTE
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
      "op": "-",          // opcional: "+" (default) o "-" para restar este chip al total
      "seccion": "PAECAL20 CONTR. CONSORCIO ALIMENTANDO  CALI 2026"  // opcional: sección del archivo
    }
  ],
  "valor_fijo": [{ "op": "+", "valor": 500000 }]  // operaciones hardcodeadas opcionales
}
```

`file_sources` mapea cada `key` a `{ "prefijo": "ESTADO DE RESULTADOS", "entidad": "CHVS" }` para que `encontrar_archivo()` localice el XLS en la carpeta del mes.

**Claves `key` generadas por `fileKey()` en app.js:**
- `ER_<ENTIDAD>` → archivos ESTADO DE RESULTADOS
- `AUX_<ENTIDAD>` → archivos 51355001
- `5105_<ENTIDAD>` → archivos 5105
- `7205_<ENTIDAD>` → archivos 7205
- `7105_<ENTIDAD>` → archivos 7105

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

- El campo `op` se guarda en `config_gastos.json` dentro de cada source (solo se escribe si es `"-"`; el `"+"` es el default para no inflar el JSON).
- `procesar_todo.py` aplica el `op` de cada source al calcular `valor_total`.
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
`TITULOS_SECCION` = filas de fórmula que se renderizan como **separador visual de sección** (barra azul con el nombre en mayúsculas, sin chips ni operaciones). Las filas de totales (TOTAL..., UTILIDAD..., COSTO NETO, etc.) de esas mismas hojas siguen ocultas. Para agregar una hoja: añadir entrada en `TITULOS_SECCION` en `app.py`.

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
| `requirements.txt` | Dependencias Python (Flask, openpyxl, xlrd, gunicorn) |
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
7. En el footer → seleccionar `ABRIL` → clic **▶ Ejecutar mes**
   - El servidor corre `procesar_todo.py ABRIL` y escribe en `/data/INFORME REAL_2026...xlsx`
8. Clic **⬇ Descargar INFORME REAL** para bajar el Excel actualizado

**El archivo de salida** es siempre `/data/INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx` (el mismo que en desarrollo). Al ejecutar un mes se sobreescribe con los valores de ese mes. Para comparar meses distintos, descarga el archivo antes de procesar el siguiente.

### Variables de entorno en Railway

| Variable | Valor | Obligatoria |
|---|---|---|
| `DATA_DIR` | `/data` | Sí (para persistencia) |
| `PORT` | asignado por Railway | Auto |
| `RAILWAY_ENVIRONMENT` | asignado por Railway | Auto (desactiva debug mode) |
