"""
procesar_gastos.py
------------------
Lee los archivos ESTADO DE RESULTADOS y 51355001 de una carpeta mensual,
aplica el mapeo definido en config_gastos.json y escribe cada valor
en la columna correspondiente del archivo principal de Gastos Administrativos.

Uso:
    py procesar_gastos.py                   # usa FEBRERO-DL por defecto
    py procesar_gastos.py MARZO-DL          # nombre de carpeta relativo
    py procesar_gastos.py C:\\ruta\\MARZO-DL # ruta absoluta
"""

import os, re, sys, json, warnings
import xlrd, openpyxl

warnings.filterwarnings('ignore')

# ---------------------------------------------------------------------------
BASE             = r'C:\Users\User\OneDrive\Desktop\CHVS\adriana'
ARCHIVO_PRINCIPAL = os.path.join(BASE, 'ANALIIS PESTAÑA GASTOS ADMINISTRATIVOS.xlsx')
CONFIG_PATH       = os.path.join(BASE, 'config_gastos.json')

MES_COLUMNA = {
    "ENERO":3,"FEBRERO":4,"MARZO":5,"ABRIL":6,"MAYO":7,"JUNIO":8,
    "JULIO":9,"AGOSTO":10,"SEPTIEMBRE":11,"OCTUBRE":12,"NOVIEMBRE":13,"DICIEMBRE":14
}
MES_ABREV = {
    "ENE":"ENERO","FEB":"FEBRERO","MAR":"MARZO","ABR":"ABRIL",
    "MAY":"MAYO","JUN":"JUNIO","JUL":"JULIO","AGO":"AGOSTO",
    "SEP":"SEPTIEMBRE","OCT":"OCTUBRE","NOV":"NOVIEMBRE","DIC":"DICIEMBRE"
}

# ---------------------------------------------------------------------------
# Localizar archivo en carpeta
# ---------------------------------------------------------------------------
def encontrar_archivo(carpeta, prefijo, entidad, mes_abrev, anio):
    """
    Busca un archivo con patron: {prefijo}_{entidad}_{mes}_{anio}*.xls
    Ej: 'ESTADO DE RESULTADOS_CHVS_FEB_2026-DL.xls'
    """
    patron = re.compile(
        rf'^{re.escape(prefijo)}_{re.escape(entidad)}_{mes_abrev}_{anio}.*\.xls$',
        re.IGNORECASE
    )
    for f in os.listdir(carpeta):
        if patron.match(f) and os.path.isfile(os.path.join(carpeta, f)):
            return os.path.join(carpeta, f)
    return None


# ---------------------------------------------------------------------------
# Lectura de ESTADO DE RESULTADOS: busqueda dinamica por codigo en col A
# Solo lee dentro de la seccion U.N. 99 GENERAL
# ---------------------------------------------------------------------------
def leer_er(filepath, codigos_buscar, sin_filtro=False):
    """
    Abre un archivo ESTADO DE RESULTADOS y suma los valores (col J) de
    todos los codigos en codigos_buscar, buscando SOLO dentro de la
    seccion U.N. 99 GENERAL (excluye CASINO u otras secciones).
    Retorna el total.
    """
    wb  = xlrd.open_workbook(filepath)
    ws  = wb.sheet_by_name('Hoja 1')

    # --- Detectar columna de periodo (buscar "Netos al" o "Debito") ---------
    col_val = None
    for r in range(min(25, ws.nrows)):
        for c in range(ws.ncols):
            txt = str(ws.cell_value(r, c)).lower()
            if 'netos' in txt or 'debito' in txt or 'debito' in txt.replace('é','e'):
                col_val = c
                break
        if col_val is not None:
            break
    if col_val is None:
        col_val = 9  # fallback: columna J (indice 9)

    # --- Delimitar seccion U.N. 99 GENERAL ----------------------------------
    inicio_99 = None
    fin_99    = ws.nrows

    for r in range(ws.nrows):
        a = str(ws.cell_value(r, 0)).strip()
        if inicio_99 is None:
            if '99' in a and 'GENERAL' in a.upper():
                inicio_99 = r + 1
        else:
            # Detener al encontrar una nueva seccion (CASINO, JUNTA, etc.)
            if a and not re.match(r'^[\d\s]', a) and a not in ('Cuentas',):
                fin_99 = r
                break

    if sin_filtro:
        inicio_99 = 0
        fin_99    = ws.nrows
    elif inicio_99 is None:
        inicio_99 = 0   # si no encontro la seccion, busca en todo el archivo

    # --- Buscar cada codigo y sumar valores ---------------------------------
    total = 0.0
    for r in range(inicio_99, fin_99):
        raw = ws.cell_value(r, 0)
        a_val = str(int(raw)) if isinstance(raw, float) and raw == int(raw) else str(raw).strip()
        if a_val in codigos_buscar:
            v = ws.cell_value(r, col_val)
            total += float(v) if v else 0.0

    return total


# ---------------------------------------------------------------------------
# Lectura de archivos auxiliares (51355001): mismo mecanismo que ER
# pero sin filtro de seccion (archivo mas pequeno, sin secciones multiples)
# ---------------------------------------------------------------------------
def leer_aux(filepath, codigos_buscar):
    """
    Abre un archivo auxiliar (ej: 51355001) y toma el valor de col J
    para las filas donde col A coincida con alguno de los codigos.
    """
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_name('Hoja 1')

    # Detectar columna debitos
    col_val = None
    for r in range(min(20, ws.nrows)):
        for c in range(ws.ncols):
            txt = str(ws.cell_value(r, c)).lower()
            if 'debito' in txt or 'debito' in txt.replace('é','e'):
                col_val = c
                break
        if col_val is not None:
            break
    if col_val is None:
        col_val = 9

    total = 0.0
    for r in range(ws.nrows):
        raw = ws.cell_value(r, 0)
        a_val = str(int(raw)) if isinstance(raw, float) and raw == int(raw) else str(raw).strip()
        # Coincidencia exacta O empieza con el codigo (para filas NIT + nombre)
        matched = a_val in codigos_buscar or any(
            a_val.startswith(code + ' ') or a_val.startswith(code + '\t')
            for code in codigos_buscar
        )
        if matched:
            v = ws.cell_value(r, col_val)
            total += float(v) if v else 0.0

    return total


# ---------------------------------------------------------------------------
# Funcion principal
# ---------------------------------------------------------------------------
def procesar_mes(carpeta_mes):
    nombre_carpeta = os.path.basename(carpeta_mes.rstrip('/\\'))
    mes_nombre     = nombre_carpeta.replace('-DL','').upper()

    if mes_nombre not in MES_COLUMNA:
        print(f"ERROR: Mes '{mes_nombre}' no reconocido.")
        return

    # Buscar la abreviatura del mes para el nombre de archivo
    mes_abrev = next((k for k,v in MES_ABREV.items() if v == mes_nombre), None)
    if not mes_abrev:
        print(f"ERROR: No se pudo determinar abreviatura para '{mes_nombre}'")
        return

    # Detectar el anio desde los archivos en la carpeta
    anio = None
    for f in os.listdir(carpeta_mes):
        m = re.search(r'_(\d{4})[-\.]', f)
        if m:
            anio = m.group(1)
            break
    if not anio:
        print("ERROR: No se pudo detectar el anio desde los nombres de archivo.")
        return

    columna_xlsx = MES_COLUMNA[mes_nombre]
    col_letra    = chr(64 + columna_xlsx)

    print(f"{'='*65}")
    print(f"Mes: {mes_nombre} | Anio: {anio} | Columna destino: {col_letra}")
    print(f"Carpeta: {carpeta_mes}")
    print(f"{'='*65}\n")

    # --- Cargar configuracion -----------------------------------------------
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        config = json.load(f)

    file_sources = config['file_sources']
    items        = config['items']

    # --- Resolver rutas de archivos fuente (cache) --------------------------
    archivos_cache = {}
    for key, info in file_sources.items():
        ruta = encontrar_archivo(
            carpeta_mes, info['prefijo'], info['entidad'], mes_abrev, anio
        )
        archivos_cache[key] = ruta
        estado = "OK" if ruta else "NO ENCONTRADO"
        print(f"  [{estado}] {key:<20} -> {os.path.basename(ruta) if ruta else '-'}")

    print()

    # --- Procesar cada item -------------------------------------------------
    if not os.path.exists(ARCHIVO_PRINCIPAL):
        print(f"ERROR: Archivo principal no encontrado:\n  {ARCHIVO_PRINCIPAL}")
        return

    wb_dest = openpyxl.load_workbook(ARCHIVO_PRINCIPAL)
    hoja_dest = next(
        (wb_dest[sh] for sh in wb_dest.sheetnames
         if 'GASTO' in sh.upper() and 'ADMIN' in sh.upper()),
        None
    )
    if not hoja_dest:
        print("ERROR: No se encontro la hoja de Gastos Administrativos.")
        return

    # Indice de filas del destino: {codigo_normalizado: fila}
    # Se normalizan espacios internos para que "51201001   -   51355001"
    # coincida con "51201001 - 51355001" sin importar cuantos espacios tenga el Excel.
    def normalizar(texto):
        return re.sub(r'\s+', ' ', str(texto).strip())

    indice_destino = {}
    for row in hoja_dest.iter_rows():
        cod = normalizar(row[0].value) if row[0].value else ''
        if cod:
            indice_destino[cod] = row[0].row

    actualizados  = 0
    no_encontrados = []

    print(f"  {'Codigo destino':<30} {'Valor calculado':>18}   {'Fila destino'}")
    print(f"  {'-'*70}")

    for item in items:
        codigo_a = item['codigo_a']
        sources  = item['sources']

        # Calcular valor total sumando todas las fuentes
        valor_total = 0.0
        advertencias = []

        for src in sources:
            key    = src['key']
            codes  = src['codes']
            ruta   = archivos_cache.get(key)

            if not ruta:
                advertencias.append(f"[ARCHIVO NO ENCONTRADO: {key}]")
                continue

            try:
                prefijo = file_sources[key]['prefijo']
                if prefijo == 'ESTADO DE RESULTADOS':
                    v = leer_er(ruta, set(codes), sin_filtro=src.get('sin_filtro', False))
                else:
                    v = leer_aux(ruta, set(codes))
                valor_total += v
            except Exception as e:
                advertencias.append(f"[ERROR en {key}: {e}]")

        # Buscar fila en destino (normalizar codigo_a tambien)
        fila = indice_destino.get(normalizar(codigo_a))
        if fila is None:
            no_encontrados.append(codigo_a)
            print(f"  {codigo_a:<30} {'[SIN FILA DESTINO]':>18}")
            continue

        # Escribir en destino
        hoja_dest.cell(row=fila, column=columna_xlsx).value = valor_total
        actualizados += 1

        adv_str = '  ' + ' '.join(advertencias) if advertencias else ''
        print(f"  {codigo_a:<30} {valor_total:>18,.2f}   fila {fila}{adv_str}")

    print(f"\n  Items actualizados : {actualizados}")
    if no_encontrados:
        print(f"  Sin fila destino   : {no_encontrados}")

    # --- Guardar ------------------------------------------------------------
    try:
        wb_dest.save(ARCHIVO_PRINCIPAL)
        print(f"\n[OK] Archivo guardado: {ARCHIVO_PRINCIPAL}")
    except PermissionError:
        print(f"\n[ERROR] Cierre el archivo en Excel y vuelva a ejecutar.")


# ---------------------------------------------------------------------------
if __name__ == '__main__':
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        carpeta = arg if os.path.isabs(arg) else os.path.join(BASE, arg)
    else:
        carpeta = os.path.join(BASE, 'FEBRERO-DL')

    procesar_mes(carpeta)
