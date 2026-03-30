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
| `config_gastos.json` | Mapeo. Cada item: `codigo_a` (fila destino), `sources` (archivos y códigos fuente), `hoja`, `fila_dest`, `buscar_por`, `valor_fijo` |
| `config_5105.json` | Reglas para archivos de nómina 5105 por entidad |
| `static/app.js` | Toda la lógica frontend: drag & drop, `mappings`, `buildConfig`, `renderChips` |
| `INFORME REAL_2026...xlsx` | Plantilla/destino. **Nunca abrir mientras corre el script.** |

### Tipos de archivos XLS fuente (en carpetas mensuales)

- `ESTADO DE RESULTADOS_<entidad>_<mes>_<año>.xls` — estado de resultados por unidad. Se lee con `leer_er()`: busca códigos en sección "99 GENERAL".
- `51355001_<entidad>_<mes>_<año>.xls` — auxiliar de proveedores. Se lee con `leer_aux()`.
- `5105_<entidad>_<mes>_<año>.xls` — nómina. Se lee con `leer_5105()`.
- `7205_<entidad>_<mes>_<año>.xls` — también auxiliar, se lee con `leer_aux()`.

### Estructura de columnas en el INFORME REAL

- **Hojas estándar** (GASTOS ADMINISTRATIVOS, GASTOS OPERATIVOS, GASTOS VEHICULOS, etc.): columna = `MES_COLUMNA[mes]` → ENE=3, FEB=4, MAR=5 ... DIC=14
- **Hojas con 3 cols/mes** (CASINO, CALI, YUMBO, BUGA): columna = `mes_num * 3` → ENE=3, FEB=6, MAR=9...
- Añadir una hoja al patrón 3-cols: agregar en `COLUMNA_HOJA` en `procesar_todo.py` y en `FILAS_EXCLUIR` en `app.py`.

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
- Fórmulas internas (totales/subtotales que el Excel calcula solo)
- Porcentajes auto-calculados sobre ingresos
- Filas cross-sheet (referencian otras hojas del mismo libro)
- Títulos sin valor

`HOJAS_EXCLUIR` = hojas que no se muestran en la UI (`%DIST C`).
`FILAS_EXCLUIR_GLOBAL` = filas excluidas en todas las hojas (encabezados corporativos).

### Celdas fusionadas (merged cells)

En `procesar_gastos`, antes de escribir se detecta si la celda pertenece a un rango fusionado y se redirige a `min_row/min_col` del rango. Esto es necesario en varias hojas del INFORME REAL.

## Rutas hardcodeadas

`BASE = r'C:\Users\User\OneDrive\Desktop\CHVS\adriana'` está en `app.py` y `procesar_todo.py`. Si se mueve el proyecto, actualizar en ambos archivos.
