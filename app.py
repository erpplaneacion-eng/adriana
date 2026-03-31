"""
Config Builder - Flask backend
Interfaz visual drag & drop para configurar config_gastos.json
"""
import os, re, json, warnings, subprocess, shutil
from datetime import date
from flask import Flask, render_template, jsonify, request

warnings.filterwarnings('ignore')

BASE         = r'C:\Users\User\OneDrive\Desktop\CHVS\adriana'
CONFIG_PATH  = os.path.join(BASE, 'config_gastos.json')
INFORME_PATH = os.path.join(BASE, 'INFORME REAL_2026 - FORMATO VERSION ORIGINAL.xlsx')
SCRIPT_PATH  = os.path.join(BASE, 'procesar_todo.py')

MES_NOMBRES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
               'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

app = Flask(__name__)


# ---------------------------------------------------------------------------
# Utilidades XLS
# ---------------------------------------------------------------------------
def celda_str(raw):
    if isinstance(raw, float) and raw == int(raw):
        return str(int(raw))
    return str(raw).strip()


def leer_todos_codigos(filepath):
    """
    Extrae TODAS las filas con valor en col A y col J de cualquier archivo XLS.
    Col A = código, col B = descripción, col J (índice 9) = valor.
    """
    try:
        import xlrd
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_name('Hoja 1')

        # Detectar columna J (valor) - forzar indice 9 si no se detecta encabezado
        col_val = 9
        for r in range(min(25, ws.nrows)):
            for c in range(ws.ncols):
                txt = str(ws.cell_value(r, c)).lower().replace('é','e').replace('ó','o')
                if 'debito' in txt or 'neto' in txt:
                    col_val = c
                    break

        col_letra = chr(65 + col_val)
        items = []

        for r in range(ws.nrows):
            raw_a = ws.cell_value(r, 0)
            a_str = str(raw_a).strip()
            if not a_str or a_str in ('None', '0.0', '0'):
                continue

            # Convertir float sin decimales (ej: 51055101.0 -> "51055101")
            if isinstance(raw_a, float) and raw_a == int(raw_a):
                a_str = str(int(raw_a))

            # Descripción: col B (índice 1)
            desc = ''
            if ws.ncols > 1:
                raw_b = ws.cell_value(r, 1)
                desc  = str(raw_b).strip() if raw_b else ''

            # Valor col J
            val = 0.0
            if ws.ncols > col_val:
                raw_j = ws.cell_value(r, col_val)
                try:
                    val = float(raw_j) if raw_j else 0.0
                except (ValueError, TypeError):
                    val = 0.0

            # Solo incluir si hay código no vacío y (descripción o valor)
            if a_str and (desc or val != 0.0):
                items.append({
                    'codigo'     : a_str,
                    'descripcion': desc[:70],
                    'valor'      : val,
                    'celda'      : f'{col_letra}{r + 1}'
                })

        return items
    except Exception as e:
        return [{'error': str(e)}]


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/sheets')
def api_sheets():
    """Retorna las hojas del INFORME REAL con sus filas (col A + col B)."""
    try:
        import openpyxl
        wb    = openpyxl.load_workbook(INFORME_PATH, data_only=True)
        config = json.load(open(CONFIG_PATH, encoding='utf-8'))
        config_map = {item['codigo_a']: item['sources'] for item in config['items']}

        # Cargar también el Excel con fórmulas para detectar SUMs internos
        wb_f = openpyxl.load_workbook(INFORME_PATH, data_only=False)

        sheets = []
        HOJAS_EXCLUIR  = {'%DIST C', '%DIST C '}
        # Hojas con 3 cols/mes: solo cols 3,6,9,12... son columnas de valor
        # Las intermedias (4,5,7,8...) son ratios/presupuesto y no deben detectarse como fórmula
        HOJAS_3COL = {
            'CASINO', 'CALI', 'YUMBO', 'BUGA',
            'COMEDORES CALI', 'COMEDORES PALMIRA', 'COMEDORES VALLE',
            'CTAS EN PPACION'
        }

        # Filas a excluir en TODAS las hojas
        FILAS_EXCLUIR_GLOBAL = {
            'CORPORACI\u00d3N HACIA UN VALLE SOLIDARIO',
            'CORPORACI\u00d3N HACIA UN VALLE SOLIDARIO ',
            'DIAS DE ATENCI\u00d3N',
            'DIAS DE ATENCI\u00d3N ',
        }

        # Filas a excluir por hoja (col B) - titulos sin valor
        FILAS_EXCLUIR = {
            'GASTOS OPERATIVOS': {
                'CUOTA DE APOYO Y SOSTENIMIENTO',  # título sin valor configurable
                'GASTOS DE PERSONAL',               # fórmula interna: suma filas 6-10
                'DIVERSOS',                         # fórmula interna: suma filas 13-16
                'TOTAL GASTOS',                     # fórmula interna: suma totales
            },
            'CASINO': {
                'Intereses', 'Descuentos Comerciales', 'otros', 'Flete',
                'Industrializados PROPIOS', 'Flete Industrializados TERCEROS',
                'Transporte Materia Prima - Cavasa', 'Flete Preparados PROPIOS',
                'Flete Preparados TERCEROS', 'Material de empaque'
            },
            'YUMBO': {
                # Totales y fórmulas internas
                'TOTAL RACIONES', 'INGRESOS BRUTOS', 'OPERACIONALES',
                'INDUSTRIAS MANUFACTURERAS', 'NO OPERACIONALES', 'FINANCIEROS',
                'DESCUENTOS - GASTO OPERACIONAL', 'OTROS', 'TOTAL INGRESOS NETOS',
                'COSTOS DE PRODUCCION', 'INSUMOS CI', 'COSTOS PREPARADOS',
                'INSUMOS CP', 'FLETES CP', 'OTROS COSTOS Y GASTOS INDIRECTOS',
                'COSTOS DE PERSONAL', 'HONORARIOS', 'SERVICIOS', 'GASTOS LEGALES',
                'ADECUACIONES E INSTALACIONES', 'DIVERSOS', 'UTILIDAD BRUTA',
                'GASTOS ADMINISTRATIVOS Y NO OPERACI', 'TOTAL AJUSTES', 'UTILIDAD NETA',
                'MANO DE OBRA DIRECTA',
                # Descuentos automáticos por porcentaje
                'Procultura 0,5%', 'Prohospitales 1%', 'Rete Ica 0,6%',
                'Adulto Mayor 2%', 'Prodeporte 2,5%', '2% Juzgado de Familia',
                '0,5% Propacifico',
                # Cross-sheet: calculadas desde otras hojas del Excel
                'PREPARADOS PROPIOS', 'Costos Indirectos de Personal',
                'Gastos Fijos Administracion', 'Gastos financieros',
                'Gastos No Operacionales',
                # Sin valor / no configurables
                'Retefuente', 'COSTOS INDUSTRIALIZADOS',
                'RACION PARA PREPARAR EN CASA (Ins)', 'REFRIGERIOS (Ins)',
                'FLETES CI', 'INDUSTRIALIZADOS PROPIOS', 'INDUSTRIALIZADOS TERCEROS',
            },
            'CALI': {
                'TOTAL RACIONES', 'INGRESOS BRUTOS', 'OPERACIONALES',
                'INDUSTRIAS MANUFACTURERAS', 'NO OPERACIONALES', 'FINANCIEROS',
                'DESCUENTOS - GASTO OPERACIONAL', 'OTROS', 'TOTAL INGRESOS NETOS',
                'COSTOS DE PRODUCCION', 'COSTOS INDUSTRIALIZADOS', 'INSUMOS CI',
                'FLETES CI', 'COSTOS PREPARADOS', 'INSUMOS CP', 'FLETES CP',
                'OTROS COSTOS Y GASTOS INDIRECTOS', 'COSTOS DE PERSONAL',
                'HONORARIOS', 'SERVICIOS', 'GASTOS LEGALES',
                'ADECUACIONES E INSTALACIONES', 'DIVERSOS',
                'UTILIDAD BRUTA', 'GASTOS ADMINISTRATIVOS Y NO OPERACI',
                'TOTAL AJUSTES', 'UTILIDAD NETA',
                'MANO DE OBRA DIRECTA',
                'PREPARADOS PROPIOS',               # calculado: GASTOS VEHICULOS * %DIST
                # Cross-sheet: se calculan automáticamente desde otras hojas del libro
                'Costos Indirectos de Personal',
                'Gastos Fijos Administracion',
                'Gastos No Operacionales',
                # Nota: 'Gastos financieros' NO se excluye  necesita configuración manual
                '2% Estampillas Prounivalle',
                '0,88% Rte Ica',
                '1% Estampilla Prohospital',
                '0,5% Propacifico',
                'Retefuente',
                '2% Juzgado de Familia',
            },
            'BUGA': {
                # Totales y fórmulas internas
                'TOTAL RACIONES', 'INGRESOS BRUTOS', 'OPERACIONALES',
                'INDUSTRIAS MANUFACTURERAS', 'NO OPERACIONALES', 'FINANCIEROS',
                'DESCUENTOS - GASTO OPERACIONAL', 'OTROS', 'TOTAL INGRESOS NETOS',
                'COSTOS DE PRODUCCION', 'COSTOS INDUSTRIALIZADOS', 'INSUMOS CI',
                'FLETES CI', 'COSTOS PREPARADOS', 'INSUMOS CP', 'FLETES CP',
                'OTROS COSTOS Y GASTOS INDIRECTOS', 'COSTOS DE PERSONAL',
                'HONORARIOS', 'ARRENDAMIENTOS', 'SEGUROS', 'SERVICIOS',
                'GASTOS LEGALES', 'MANTENIMIENTO Y REPARACIONES',
                'ADECUACIONES E INSTALACIONES', 'GASTOS DE VIAJE', 'DIVERSOS',
                'UTILIDAD BRUTA', 'GASTOS ADMINISTRATIVOS Y NO OPERACIONALES',
                'TOTAL AJUSTES', 'UTILIDAD NETA',
                'MANO DE OBRA DIRECTA',
                # Nota: filas (Fac) visibles para configuración manual en UI
                # Porcentajes auto-calculados sobre ingresos
                '1% Estampilla Prohospital', '2,5% Estampilla Pro Deporte',
                '1% Estampilla Pro Univalle', '3% Estampilla Adulto Mayor',
                '0,5% Estampilla Universidad del Pacifico',
                '0,66% Rete Ica', 'Retefuente', '2% Juzgado de Familia',
                # Cross-sheet: calculadas desde otras hojas del mismo libro
                'PREPARADOS PROPIOS',
                'Gastos Fijos Administracion', 'Gastos financieros',
                'Gastos No Operacionales',
                # Sin valor / no configurables
                'Intereses', 'Descuentos Comerciales', 'otros Fin',
                'INDUSTRIALIZADOS TERCEROS',
            }
        }
        for sh_name in wb.sheetnames:
            ws      = wb[sh_name]
            ws_f    = wb_f[sh_name]
            if ws.sheet_state == 'hidden':
                continue
            if sh_name.strip() in HOJAS_EXCLUIR:
                continue
            rows = []
            for row in ws.iter_rows():
                a = str(row[0].value or '').strip()
                b = str(row[1].value or '').strip() if len(row) > 1 else ''

                # Incluir si col A tiene código numérico O si col B tiene texto descriptivo
                tiene_codigo = bool(a and re.match(r'^\d', a))
                if not tiene_codigo and not b:
                    continue  # fila vacía

                # Excluir filas globales y específicas por hoja
                if b in FILAS_EXCLUIR_GLOBAL:
                    continue
                if b in FILAS_EXCLUIR.get(sh_name.strip(), set()):
                    continue

                # Detectar fórmula interna: referencia otras filas de la MISMA hoja
                # Solo se revisan las columnas de valor (no las de ratio/presupuesto)
                es_formula = False
                current_row = row[0].row
                if sh_name.strip() in HOJAS_3COL:
                    cols_valor = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36]
                else:
                    cols_valor = list(range(3, 15))  # cols C a N (ENE-DIC estándar)
                for col_num in cols_valor:
                    v = ws_f.cell(row=current_row, column=col_num).value
                    if not v or not isinstance(v, str) or not v.startswith('='):
                        continue
                    # Si tiene referencia a libro externo o a otra hoja  NO es fórmula interna
                    if '[' in v or "'" in v:
                        continue
                    # Si referencia otras filas de la misma hoja  SÍ es fórmula interna
                    nums = re.findall(r'\d+', v)
                    otras_filas = [int(n) for n in nums if int(n) != current_row and 1 < int(n) < 500]
                    if otras_filas:
                        es_formula = True
                        break

                # Filas con fórmulas internas o cross-sheet: Excel las calcula solo,
                # no deben aparecer en la UI para evitar que el usuario las configure.
                if es_formula:
                    continue

                # Clave de búsqueda en config: código si existe, si no "B:<descripcion>"
                clave = re.sub(r'\s+', ' ', a) if tiene_codigo else f'B:{b}'
                fuentes = config_map.get(clave, config_map.get(a, []))

                rows.append({
                    'fila'        : row[0].row,
                    'codigo'      : a,
                    'descripcion' : b,
                    'fuentes'     : fuentes,
                    'tiene_codigo': tiene_codigo,
                })

            if rows:
                sheets.append({'nombre': sh_name.strip(), 'filas': rows})
        return jsonify(sheets)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/folders')
def api_folders():
    """Lista carpetas mensuales disponibles en BASE."""
    carpetas = []
    for nombre in sorted(os.listdir(BASE)):
        ruta = os.path.join(BASE, nombre)
        if not os.path.isdir(ruta):
            continue
        nombre_limpio = nombre.replace('-DL', '').upper()
        if nombre_limpio in MES_NOMBRES:
            # Contar archivos XLS
            n_xls = len([f for f in os.listdir(ruta) if f.lower().endswith('.xls')])
            if n_xls > 0:
                carpetas.append({'nombre': nombre, 'n_archivos': n_xls})
    return jsonify(carpetas)


@app.route('/api/files/<path:folder>')
def api_files(folder):
    """Lee todos los XLS de la carpeta y extrae códigos por archivo."""
    carpeta = os.path.join(BASE, folder) if not os.path.isabs(folder) else folder
    if not os.path.isdir(carpeta):
        return jsonify({'error': f'Carpeta no encontrada: {carpeta}'}), 404

    resultado = {}
    for nombre in sorted(os.listdir(carpeta)):
        if not nombre.lower().endswith('.xls'):
            continue
        ruta  = os.path.join(carpeta, nombre)
        items = leer_todos_codigos(ruta)
        resultado[nombre] = {'items': items}

    return jsonify(resultado)


@app.route('/api/config')
def api_config():
    with open(CONFIG_PATH, encoding='utf-8') as f:
        return jsonify(json.load(f))


@app.route('/api/config/save', methods=['POST'])
def api_config_save():
    """Recibe el nuevo config y lo guarda."""
    try:
        nuevo_config = request.get_json()
        # Backup del config anterior
        backup = CONFIG_PATH.replace('.json', f'_backup_{date.today()}.json')
        shutil.copy2(CONFIG_PATH, backup)
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(nuevo_config, f, ensure_ascii=False, indent=2)
        return jsonify({'ok': True, 'backup': backup})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/run/<path:folder>', methods=['POST'])
def api_run(folder):
    """Ejecuta procesar_todo.py para la carpeta indicada."""
    try:
        result = subprocess.run(
            ['py', SCRIPT_PATH, folder],
            capture_output=True, text=True, encoding='utf-8', errors='replace',
            timeout=300
        )
        return jsonify({
            'ok'    : result.returncode == 0,
            'stdout': result.stdout,
            'stderr': result.stderr
        })
    except subprocess.TimeoutExpired:
        return jsonify({'error': 'Timeout - el script tardo mas de 5 minutos'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print(f'\n  Config Builder iniciado')
    print(f'  Abre http://localhost:5000 en tu navegador\n')
    app.run(debug=True, port=5000)

