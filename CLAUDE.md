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
- `51355001_<entidad>_<mes>_<año>.xls` — auxiliar de proveedores. Se lee con `leer_aux()`.
- `5105_<entidad>_<mes>_<año>.xls` — nómina. Se lee con `leer_5105()`.
- `7205_<entidad>_<mes>_<año>.xls` — también auxiliar, se lee con `leer_aux()`.

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
    { "key": "ER_CHVS", "codes": ["51055101"] }
  ],
  "valor_fijo": [{ "op": "+", "valor": 500000 }]  // operaciones hardcodeadas opcionales
}
```

`file_sources` mapea cada `key` a `{ "prefijo": "ESTADO DE RESULTADOS", "entidad": "CHVS" }` para que `encontrar_archivo()` localice el XLS en la carpeta del mes.

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

### Archivos 5105 — secciones por unidad de negocio (app.py)

`leer_todos_codigos()` en `app.py` detecta secciones (unidades de negocio) dentro del archivo 5105:
- Busca filas sin código numérico dentro del bloque `5105` → se tratan como separadores de sección (`separador: true`).
- Cada ítem resultante lleva el campo `seccion` con el nombre de la unidad a la que pertenece (ej. `"99 GENERAL"`, `"PAEBUGA08 CONTRATO UT BUGA 2026"`).
- En el panel derecho, los separadores se renderizan como barras azules (`.source-row-sep`).
- Esto permite que dos filas con el mismo código (ej. `010304 APRENDICES TALENTO`) de unidades distintas sean ambas arrastrables e independientes.

### Persistencia de valores y sección en mappings (app.js)

`actualizarValoresEnMappings(data)` — se llama al cargar una carpeta. Cruza los datos XLS cargados con los `mappings` existentes para poblar `src.valor` sin necesidad de reconfigurar:
- Retrocompatibilidad: si el código guardado en config tiene formato antiguo combinado (`"010102           GERENCIA"`), extrae la parte numérica con regex antes de buscar en el lookup.
- Para `sin_filtro: true` suma todos los valores del archivo.

`buildConfig()` — al guardar, persiste el campo `seccion` de cada source entry en el JSON.

`initMappingsFromConfig()` — al cargar, restaura el campo `seccion` desde el JSON, lo que permite que el badge de unidad de negocio (`.chip-sec`, cursiva morada) aparezca en los chips del panel izquierdo tras reabrir la app.

### config_gastos.json — campo seccion en sources

```json
"sources": [
  { "key": "5105_CHVS", "codes": ["010304"], "seccion": "PAEBUGA08 CONTRATO UT BUGA 2026" }
]
```

El campo `seccion` es opcional. Solo se graba cuando el chip proviene de una sección secundaria del archivo 5105. Sin él, el chip se asume de la sección principal.

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
