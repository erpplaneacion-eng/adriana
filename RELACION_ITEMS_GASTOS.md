# Relación de Ítems — Gastos Administrativos

> **Nota sobre nombres de archivo:** Los nombres varían según el mes y si tienen el sufijo `-DL`.
> Ejemplo para febrero: `ESTADO DE RESULTADOS_CHVS_FEB_2026-DL.xls`
> Ejemplo para marzo (sin -DL): `ESTADO DE RESULTADOS_CHVS_MAR_2026.xls`
> El programa detecta el mes y el año automáticamente desde el nombre de la carpeta.

---

## Archivos fuente utilizados

| Archivo (patrón) | Qué contiene |
|------------------|--------------|
| `ESTADO DE RESULTADOS_CHVS_{MES}_{AÑO}[-DL].xls` | Estado de resultados CHVS — solo se lee la sección **U.N. 99 GENERAL** |
| `ESTADO DE RESULTADOS_UT ALIM YUMBO2025_{MES}_{AÑO}[-DL].xls` | Estado de resultados UT ALIM YUMBO 2025 |
| `ESTADO DE RESULTADOS_UT ALIM YUMBO2026_{MES}_{AÑO}[-DL].xls` | Estado de resultados UT ALIM YUMBO 2026 |
| `ESTADO DE RESULTADOS_UT ALIM BUGA2026_{MES}_{AÑO}[-DL].xls` | Estado de resultados UT ALIM BUGA 2026 |
| `ESTADO DE RESULTADOS_CONS ALIM CALI2026_{MES}_{AÑO}[-DL].xls` | Estado de resultados CONS ALIM CALI 2026 |
| `51355001_CHVS_{MES}_{AÑO}[-DL].xls` | Auxiliar de la cuenta 51355001 — detalle por proveedor |

---

## Operaciones por ítem

### Ítems simples — 1 archivo, 1 código

Todos leen del archivo `ESTADO DE RESULTADOS_CHVS`, sección U.N. 99 GENERAL,
y toman el valor de la columna **Débitos** donde la columna A sea igual al código indicado.

| Código destino | Descripción | Código a buscar en col A |
|----------------|-------------|--------------------------|
| `51055101` | Dotación y Suministro | `51055101` |
| `51056301` | Capacitación | `51056301` |
| `51058401` | Gastos Médicos | `51058401` |
| `51101001` | Revisoría Fiscal | `51101001` |
| `51202001` | Equipo de Oficina | `51202001` |
| `51351502` | Laboratorio | `51351502` |
| `51351503` | Fumigación | `51351503` |
| `51352501` | Acueducto y Alcantarillado | `51352501` |
| `51353001` | Energía Eléctrica | `51353001` |
| `51353501` | Teléfono | `51353501` |
| `51353502` | Internet | `51353502` |
| `51354001` | Correo, Portes, Telegramas | `51354001` |
| `514095` | Certificados Cámara Comercio | `514095` |
| `51451001` | Construcciones y Edificaciones | `51451001` |
| `51451501` | Plan SST | `51451501` |
| `51951001` | Libros, Suscripciones | `51951001` |
| `51952501` | Elementos de Aseo y Cafetería | `51952501` |
| `51953001` | Papelería | `51953001` |
| `51954501` | Taxis y Buses | `51954501` |
| `51956001` | Casino y Restaurante (Refrigerios) | `51956001` |
| `51956501` | Parqueaderos | `51956501` |
| `51959501` | Otros | `51959501` |
| `51959502` | Compra Equipos de Oficina | `51959502` |
| `51959503` | Compra Suministros | `51959503` |
| `51959510` | Herramientas y Artículos Ferretería | `51959510` |
| `51959512` | Reuniones, Fiestas y Detalles | `51959512` |
| `51959514` | Equipos de Bodega | `51959514` |
| `51959515` | Compra Equipos de Cómputo | `51959515` |
| `51959504` | Cuota de Apoyo y Sostenimiento (SENA) | `51959504` |
| `51959518` | Compra Mercado IVH | `51959518` |
| `51959517` | Convenios Médicos - EMI | `51959517` |
| `51956002` | Donaciones | `51956002` |
| `53954001` | Ajuste al Peso | `53954001` |
| `530505` | Gastos Bancarios | `530505` |
| `530515` | Comisiones | `530515` |
| `53052001` | Intereses Corrientes Obligaciones | `53052001` |
| `53052004` | Intereses Corrientes Leasing | `53052004` |

---

### Ítems especiales — lógica diferente

---

#### `53954002` — Incapacidades
**Archivo:** `ESTADO DE RESULTADOS_UT ALIM YUMBO2025` (no CHVS)
**Operación:** buscar código `53954002` en col A → tomar valor col Débitos

---

#### `51103508` — Asesoría Técnica
**Archivo:** `ESTADO DE RESULTADOS_CHVS`
**Operación:** buscar código `511035` (código grupo padre) en col A → tomar valor col Débitos
> El código `51103508` no existe directamente; se usa el código agrupador `511035`

---

#### `51102001` — Asesoría Jurídica
**Operación:** suma de dos fuentes:
1. `ESTADO DE RESULTADOS_CHVS` → código `511025` → valor col Débitos
2. `51355001_CHVS` → NIT `16673822` (ROSALES MONTEMIRANDA ALVARO JOSE) → valor col Débitos

---

#### `51355001` — Transporte, Fletes y Acarreos
**Archivo:** `51355001_CHVS` (auxiliar de proveedores)
**Operación:** buscar NIT `1069174807` (CASTRO AVILA JENNY MARCELA) → tomar valor col Débitos
> Solo se toma el valor del proveedor CASTRO, no el total de la cuenta

---

#### `51201501 - 73201501` — Maquinaria y Equipo
**Archivo:** `ESTADO DE RESULTADOS_CHVS`
**Operación:** buscar código `51201501` + código `73201501` → sumar ambos valores

---

#### `51351001 - 51351004` — Temporales
**Archivo:** `ESTADO DE RESULTADOS_CHVS`
**Operación:** buscar código `51351001` + código `51351004` → sumar ambos valores

---

#### `51954001 - 71100501` — Envases y Empaques
**Archivo:** `ESTADO DE RESULTADOS_CHVS`
**Operación:** buscar código `51954001` + código `71100501` → sumar ambos valores

---

#### `51159502` — Gravamen a Movimientos Financieros
**Archivo:** `ESTADO DE RESULTADOS_CHVS`
**Operación:** buscar código `51159502` + código `53055001` → sumar ambos valores

---

#### `51356501` — Servicio Web / Hosting / Licencias Siesa y Planeación
**Operación:** suma de 5 fuentes:
1. `ESTADO DE RESULTADOS_CHVS` → código `51356501` → sección U.N.99
2. `ESTADO DE RESULTADOS_CHVS` → código `733565` → **todo el archivo** (incluye sección CASINO)
3. `ESTADO DE RESULTADOS_UT ALIM YUMBO2026` → código `733565` → sección U.N.99
4. `ESTADO DE RESULTADOS_UT ALIM BUGA2026` → código `733565` → sección U.N.99
5. `ESTADO DE RESULTADOS_CONS ALIM CALI2026` → código `733565` → sección U.N.99

---

#### `51201001 - 51355001` — *(PENDIENTE DE RESOLVER)*
**Operación prevista:** suma de dos fuentes:
1. `ESTADO DE RESULTADOS_CHVS` → código `51201001`
2. `51355001_CHVS` → entradas correspondientes a Construcciones y Edificaciones

> **Pendiente:** se debe identificar qué proveedores del archivo auxiliar `51355001`
> corresponden a este concepto, y resolver el espaciado exacto del código en el archivo destino.

---

---

#### `5105` — Salarios
**Script:** `procesar_5105.py` (config: `config_5105.json`)
**Archivos:** todos los `5105_{ENTIDAD}_{MES}_{AÑO}[-DL].xls` de la carpeta mensual
**Operación:** para cada entidad, busca la sección **U.N. 99 GENERAL**, toma el total de la cuenta `5105` y le resta los códigos indicados. Luego suma todas las entidades.

| Entidad | Archivo | Operación |
|---------|---------|-----------|
| CHVS | `5105_CHVS_{MES}.xls` | Total `5105` − códigos `010103` + `010303` + `010501` |
| UT ALIM YUMBO2026 | `5105_UT ALIM YUMBO2026_{MES}.xls` | Total `5105` − código `010303` |
| UT ALIM YUMBO2025 | `5105_UT ALIM YUMBO2025_{MES}.xls` | Total `5105` − código `010303` |
| UT ALIM BUGA2026 | `5105_UT ALIM BUGA2026_{MES}.xls` | Total `5105` sin restar nada |
| UT BUGA2025 | `5105_UT BUGA2025_{MES}.xls` | Solo el valor del código `010403` |
| CONS ALIM CALI2026 | `5105_CONS ALIM CALI2026_{MES}.xls` | Total `5105` − código `010303` |

> El resultado total (suma de todas las entidades) se escribe en la fila donde col A = `5105` del archivo destino.

---

## Resumen

| Tipo | Cantidad |
|------|----------|
| Ítems simples (1 archivo, 1 código) | 37 |
| Ítems con suma de códigos en mismo archivo | 4 |
| Ítems con múltiples archivos | 5 |
| Ítems con lógica especial (NIT, sección CASINO) | 2 |
| **Pendientes** | **1** |
| **Total** | **47** + salarios (5105) |
