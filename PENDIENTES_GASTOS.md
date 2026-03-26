# Pendientes a Resolver — Automatización Gastos Administrativos

## Contexto
Programa `procesar_gastos.py` actualiza el archivo destino
`ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx` a partir de los archivos
de la carpeta mensual (`FEBRERO-DL`, `MARZO-DL`, etc.).
De 47 ítems configurados, **46 se procesan correctamente**.
Los siguientes requieren clarificación:

---

## 1. `51102001` — Asesoría Jurídica
**Valor en destino FEB:** $6,485,768
**Valor calculado actualmente:** $0

**Problema:** El código `51102001` no existe en el archivo
`ESTADO DE RESULTADOS_CHVS_FEB_2026-DL.xls`. El archivo
`match_columnaAycolumnaB.xlsx` indica que parte del valor viene de:
- `ESTADO DE RESULTADOS_CHVS` celda `$C$75`
- `51355001_CHVS` celda `$G$14`

Pero en los archivos DL esas referencias no corresponden.

**Pregunta:** ¿De dónde proviene el valor $6,485,768?
¿Está en algún otro archivo o bajo otro código?

---

## 2. `51355001` — Transporte, fletes y acarreos
**Valor en destino FEB:** $404,040
**Valor calculado actualmente:** $76,192,016 (total de la cuenta)

**Problema:** El archivo auxiliar `51355001_CHVS_FEB_2026-DL.xls`
contiene múltiples proveedores y el programa toma el total de la cuenta.
El valor correcto ($404,040) corresponde solo al proveedor **CASTRO**
(NIT 1069174807, descripción `CM1-2479 TRANS`).

**Pregunta:** ¿Se debe filtrar por proveedor específico, por descripción
(`CM1` o `TRANS`), o hay algún criterio claro para separar
Transporte de los demás conceptos del mismo archivo?

---

## 3. `51201001 - 51355001` — Construcciones y Edificaciones
**Valor en destino FEB:** $87,589,181
**Valor calculado actualmente:** No se actualizó (fila no encontrada)

Hay dos sub-problemas:

**3a. Fila no encontrada en el destino:**
El código en la columna A del archivo destino tiene espacios internos
que no coincidieron exactamente. Verificar cómo aparece escrito
(puede ser `51201001   - 51355001` con múltiples espacios).

**3b. Fuentes del valor:**
Según el `match_columnaAycolumnaB.xlsx`, este ítem suma tres fuentes:
- `ESTADO DE RESULTADOS_CHVS` → código `51201001` → $37,439,181
- `51355001_CHVS` celda `$G$16` → valor desconocido en archivo DL
- `51355001_CHVS` celda `$G$18` → valor desconocido en archivo DL

La diferencia entre el total ($87,589,181) y la parte del ER ($37,439,181)
es **$50,150,000**, que debería venir del archivo auxiliar `51355001`.
¿Cómo se identifica en ese archivo cuáles entradas corresponden
a Construcciones y Edificaciones?

---

## 4. `51356501` — Servicio Web / Hosting / Licencias
**Valor en destino FEB:** $8,718,886
**Valor calculado actualmente:** $5,741,346

**Problema:** El `match_columnaAycolumnaB.xlsx` indica que este ítem
suma valores de dos secciones del `ESTADO DE RESULTADOS_CHVS`:
- Sección U.N. 99 GENERAL (`$C$117`) → capturado correctamente
- Sección **CASINO** (`$C$312`) → NO capturado (el programa excluye
  la sección CASINO por defecto para evitar doble conteo)

Además debe sumar las otras entidades (YUMBO, BUGA, CALI) pero
el código `51356501` no se encontró en esos archivos.

**Pregunta:** ¿Se debe incluir la parte del CASINO en este ítem?
¿Y en cuáles de los otros archivos de entidades aparece `51356501`?

---

## 5. Códigos no encontrados en ESTADO DE RESULTADOS CHVS
Los siguientes códigos del `match_columnaAycolumnaB.xlsx` no aparecen
en columna A del archivo `ESTADO DE RESULTADOS_CHVS_FEB_2026-DL.xls`:

| Código | Ítem destino |
|--------|-------------|
| `51102001` | Asesoría Jurídica |
| `73201501` | Parte de Maquinaria y Equipo (se usa `51201501` como sustituto) |

¿Bajo qué código o archivo se encuentra cada uno?

---

## Archivos involucrados
- Destino: `ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx`
- Config ítems: `config_gastos.json`
- Config salarios: `config_5105.json`
- Script gastos: `procesar_gastos.py`
- Script salarios: `procesar_5105.py`
- Mapeo fuentes: `match_columnaAycolumnaB.xlsx`
