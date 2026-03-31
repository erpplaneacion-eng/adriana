# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué hace este proyecto

Sistema para llenar automáticamente el **INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx** (multi-hoja) con datos contables mensuales extraídos de archivos XLS fuente. El mapeo fuente→destino vive en `config_gastos.json` y se edita visualmente con el Config Builder (Flask).

## Comandos principales

```bash
# Iniciar Config Builder (UI web para configurar mapeos)
cd C:\Users\User\OneDrive\Desktop\CHVS\adriana
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
| `app.py` | Flask backend del Config Builder. Sirve la UI y expone `/api/sheets`, `/api/files`, `/api/config/save`, `/api/run` |
| `procesar_todo.py` | Script principal. Paso 1: salarios (5105). Paso 2: gastos varios (ER + aux). Escribe en el INFORME REAL. |
| `procesar_gastos.py` | Script standalone de gastos (legado/independiente). Apunta a `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx`. |
| `procesar_5105.py` | Script standalone de nómina (legado/independiente). Apunta a `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx`. |
| `config_gastos.json` | Mapeo principal. Cada item: `codigo_a`, `sources`, `hoja`, `fila_dest`, `buscar_por`, `valor_fijo`. Se respalda automáticamente al guardar. |
| `config_5105.json` | Reglas para archivos de nómina 5105 por entidad |
| `config_gastos_backup_YYYY-MM-DD.json` | Backups automáticos del config al guardar desde la UI. |
| `static/app.js` | Toda la lógica frontend: drag & drop, `mappings`, `buildConfig`, `renderChips` |
| `INFORME REAL_2026...xlsx` | Plantilla/destino. **Nunca abrir mientras corre el script.** |

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

El key único de cada fila en `mappings` es `"${codeKey}|${fila.fila}"` (incluye número de fila Excel para evitar colisiones cuando el mismo código aparece en dos filas). Al guardar, el `|filaNum` se extrae y se guarda como `fila_dest`.

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
`TITULOS_SECCION` = filas de fórmula que se renderizan como **separador visual de sección** (barra azul con el nombre en mayúsculas, sin chips ni operaciones). Definido para: GASTOS ADMINISTRATIVOS, GASTOS VEHICULOS, CASINO, CALI, YUMBO, COMEDORES CALI, COMEDORES PALMIRA, CTAS EN PPACION, PYG TOTAL. Las filas de totales (TOTAL..., UTILIDAD..., COSTO NETO, etc.) de esas mismas hojas siguen ocultas. Para agregar una hoja: añadir entrada en `TITULOS_SECCION` en `app.py`.

### Fila especial: Gastos financieros (CALI fila 149)

Esta fila tiene fórmula cross-sheet con **constantes hardcodeadas distintas por mes** (`='GASTOS ADMINISTRATIVOS'!C109*'%DIST'!C52+520117.4`). Las fórmulas de MAR-DIC no existen en la plantilla y NO pueden restaurarse automáticamente porque el ajuste mensual varía. Requiere entrada manual en Excel para cada mes. Aparece visible en el Config Builder para que el usuario pueda asignarle un valor fijo o fuente.

### Celdas fusionadas (merged cells)

En `procesar_gastos`, antes de escribir se detecta si la celda pertenece a un rango fusionado y se redirige a `min_row/min_col` del rango. Esto es necesario en varias hojas del INFORME REAL.

## Estado del INFORME REAL (correcciones aplicadas 2026-03-30)

### Fórmulas restauradas por hoja

Durante la revisión se detectó que el script había sobreescrito fórmulas internas de Excel en varias hojas al procesar MARZO y ABRIL. Se restauraron:

| Hoja | Filas afectadas | Meses restaurados |
|---|---|---|
| GASTOS VEHICULOS | 9,11,15,19,22,25 (subtotales) | MAR, ABR |
| CASINO | 5,12,13,14,15,23,24,28,29,42,44,60,61,78,83,84,88,91,95,98,108,113,119,122,126,140,144 | MAR, ABR |
| CALI | 85,148,151 (cross-sheet %DIST) | MAR→DIC |
| YUMBO | 23,85,123,148,151 | OCT→DIC y MAR→DIC según fila |
| BUGA | 23,73,148,151 | DIC y MAR→DIC según fila |

### Fórmulas cross-sheet en hojas 3-col/mes

Las filas que referencian `'%DIST'` usan un patrón de 2 cols/mes:
- `%DIST` col par (B,D,F,H...): `chr(64 + 2*mes_num)` → ENE=B, FEB=D, MAR=F...
- `%DIST` col impar (C,E,G,I...): `chr(65 + 2*mes_num)` → ENE=C, FEB=E, MAR=G...
- Hoja referenciada (GASTOS ADMIN/OPERATIVOS) usa col estándar: `chr(66 + mes_num)` → ENE=C, FEB=D, MAR=E...

### Filas de título limpiadas

- **GASTOS OPERATIVOS fila 11** (CUOTA DE APOYO Y SOSTENIMIENTO): tenía valores hardcodeados en columnas MAR y ABR. Se limpiaron.

## Rutas — desarrollo vs. Railway

`BASE` ya **no** está hardcodeada. Ambos scripts la resuelven dinámicamente:

| Contexto | Valor de BASE |
|---|---|
| Desarrollo local | directorio del script (`os.path.dirname(__file__)`) |
| Railway (producción) | variable de entorno `DATA_DIR` (volumen persistente) |

En `app.py`: `BASE = os.environ.get('DATA_DIR', _APP_DIR)`
En `procesar_todo.py`: `BASE = os.environ.get('CHVS_BASE', os.path.dirname(__file__))`

`app.py` pasa `CHVS_BASE=BASE` al subprocess de `procesar_todo.py` para que ambos trabajen sobre el mismo directorio.

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
