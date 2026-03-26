# Pendientes — Continuación trabajo automatización Gastos Administrativos
**Fecha:** 2026-03-25

---

## Estado actual

El script `procesar_todo.py` procesa en un solo paso:
- **PASO 1 — Salarios `5105`** (config: `config_5105.json`)
- **PASO 2 — Gastos varios** (config: `config_gastos.json`)

De los 47 ítems de gastos, **46 se actualizan correctamente**.
Quedan **5 filas del archivo destino sin actualizar**.

---

## Ítems pendientes de resolver

### 1. `51201001   - 51355001` — Construcciones y Edificaciones
**Problema A — Fila no encontrada:**
El código en la columna A del archivo destino tiene espacios internos que no coinciden
con el string del config. Verificar el valor exacto con `repr()`.

**Problema B — Fuente incompleta:**
El config actual usa el total de la cuenta `51355001` del archivo auxiliar,
pero solo una parte corresponde a Construcciones. Falta identificar qué
proveedores/NITs del archivo `51355001_CHVS_{MES}.xls` corresponden a este concepto.

Referencia FEB destino: `='[8]Hoja 1'!$C$84+'[9]Hoja 1'!$G$16+'[9]Hoja 1'!$G$18`
- ER_CHVS código `51201001` → $37,439,181
- Diferencia desde auxiliar → $50,150,000 (proveedores a identificar)

---

### 2. `51451001      -   51451501` — (descripción a confirmar)
**Problema:** Fila con espacios internos en col A — no está en `config_gastos.json`.
Valor FEB destino: **$318,000**

**Pendiente:** Identificar de qué archivo y código proviene ese valor
y agregarlo al config con el código exacto (con espacios) que aparece en col A del destino.

---

### 3. `51953501` — (descripción a confirmar)
**Problema:** Código **no configurado** en `config_gastos.json`.
Valor FEB destino: **$1,186,883**

**Pendiente:** Identificar de qué archivo fuente y bajo qué código
viene este valor y agregar el ítem al config.

---

### 4. `51451501` — Plan SST
**Problema:** En el destino hay una fila con fórmula Excel antigua para este código.
El programa sí escribe el valor calculado en fila 58 (código `51451501` simple),
pero existe también la fila `51451001      -   51451501` con $318,000.

**Pendiente:** Confirmar si `51451501` (fila 58, $2,062,464) y
`51451001 - 51451501` (fila con $318,000) son el mismo concepto o dos separados.

---

### 5. `51956601` — (descripción a confirmar)
**Problema:** Código **no configurado** en `config_gastos.json`.
Valor FEB destino: fórmula Excel antigua (valor real desconocido).

**Pendiente:** Abrir el archivo destino, ver el valor real de FEB para este código
e identificar de qué archivo fuente proviene.

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `procesar_todo.py` | Script principal unificado (5105 + gastos) |
| `config_gastos.json` | Mapeo de ítems a archivos/códigos fuente |
| `config_5105.json` | Reglas por entidad para salarios |
| `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx` | Archivo destino |
| `RELACION_ITEMS_GASTOS.md` | Documentación completa de todos los ítems |
| `FEBRERO-DL\` | Carpeta con archivos fuente de febrero (con sufijo -DL) |
| `MARZO\`, `ABRIL\` | Carpetas de prueba (copia de febrero renombrada) |

## Uso del script

```
py procesar_todo.py FEBRERO-DL     # procesa febrero
py procesar_todo.py MARZO          # procesa marzo
py procesar_todo.py MARZO-DL       # también funciona con sufijo -DL
py procesar_todo.py C:\ruta\ABRIL  # ruta absoluta
```

> **Importante:** cerrar el archivo destino en Excel antes de ejecutar.


claude --resume 242343ae-77af-41fa-a9b7-385ad0c18f05