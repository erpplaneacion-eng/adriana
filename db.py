"""
db.py — Base de datos SQLite para configuración y auditoría de ejecuciones.
Reemplaza config_gastos.json y config_5105.json.
"""
import os
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime

_APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _get_base():
    return (os.environ.get('DATA_DIR') or
            os.environ.get('CHVS_BASE') or
            _APP_DIR)


def get_db_path():
    return os.path.join(_get_base(), 'chvs.db')


@contextmanager
def get_conn():
    conn = sqlite3.connect(get_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


_SCHEMA = """
CREATE TABLE IF NOT EXISTS file_sources (
    key     TEXT PRIMARY KEY,
    prefijo TEXT NOT NULL,
    entidad TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS dest_rows (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    hoja       TEXT NOT NULL,
    codigo_a   TEXT NOT NULL,
    fila_dest  INTEGER,
    buscar_por TEXT NOT NULL DEFAULT 'A',
    mes        TEXT
);

CREATE TABLE IF NOT EXISTS sources (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    dest_row_id INTEGER NOT NULL REFERENCES dest_rows(id) ON DELETE CASCADE,
    file_key    TEXT NOT NULL,
    codes       TEXT NOT NULL,
    op          TEXT NOT NULL DEFAULT '+',
    sin_filtro  INTEGER NOT NULL DEFAULT 0,
    seccion     TEXT,
    orden       INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS valor_fijo (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    dest_row_id INTEGER NOT NULL REFERENCES dest_rows(id) ON DELETE CASCADE,
    op          TEXT NOT NULL DEFAULT '+',
    valor       REAL NOT NULL,
    orden       INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS config_5105 (
    entidad        TEXT PRIMARY KEY,
    only_code      TEXT,
    subtract_codes TEXT
);

CREATE TABLE IF NOT EXISTS ejecuciones (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    mes   TEXT NOT NULL,
    fecha TEXT NOT NULL,
    exito INTEGER NOT NULL DEFAULT 1,
    notas TEXT
);

CREATE TABLE IF NOT EXISTS valores_escritos (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    ejecucion_id INTEGER NOT NULL REFERENCES ejecuciones(id),
    hoja         TEXT NOT NULL,
    fila         INTEGER NOT NULL,
    codigo_a     TEXT NOT NULL,
    valor_total  REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS fuentes_valor (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    valor_id INTEGER NOT NULL REFERENCES valores_escritos(id),
    file_key TEXT NOT NULL,
    codigo   TEXT NOT NULL,
    seccion  TEXT,
    op       TEXT NOT NULL DEFAULT '+',
    valor    REAL NOT NULL
);
"""


def init_db():
    """Crea el schema si no existe."""
    with get_conn() as conn:
        conn.executescript(_SCHEMA)
        try:
            conn.execute("ALTER TABLE dest_rows ADD COLUMN mes TEXT")
        except Exception:
            pass  # columna ya existe


def _ya_migrado():
    with get_conn() as conn:
        return conn.execute("SELECT COUNT(*) FROM dest_rows").fetchone()[0] > 0


def migrate_from_json():
    """
    Migra config_gastos.json y config_5105.json a la DB.
    No hace nada si la DB ya tiene datos.
    """
    if _ya_migrado():
        return

    base = _get_base()

    cfg_path = os.path.join(base, 'config_gastos.json')
    if os.path.exists(cfg_path):
        with open(cfg_path, encoding='utf-8') as f:
            cfg = json.load(f)
        save_config(cfg)
        print(f"[DB] Migrado config_gastos.json → {get_db_path()}")

    cfg5_path = os.path.join(base, 'config_5105.json')
    if os.path.exists(cfg5_path):
        with open(cfg5_path, encoding='utf-8') as f:
            cfg5 = json.load(f)
        save_config_5105(cfg5)
        print(f"[DB] Migrado config_5105.json → {get_db_path()}")


# ---------------------------------------------------------------------------
# Guardar / cargar config_gastos
# ---------------------------------------------------------------------------

def save_config(cfg, mes=None):
    """Guarda config en la DB.
    mes=None  → reemplaza todo (legado).
    mes='MAYO' → reemplaza solo los items de ese mes; file_sources siempre global.
    """
    mes_key = mes.upper() if mes else None
    with get_conn() as conn:
        # file_sources siempre globales
        conn.execute("DELETE FROM file_sources")
        for key, info in cfg.get('file_sources', {}).items():
            conn.execute(
                "INSERT INTO file_sources (key,prefijo,entidad) VALUES (?,?,?)",
                (key, info['prefijo'], info['entidad'])
            )

        # dest_rows: solo borrar el mes indicado (CASCADE elimina sources/valor_fijo)
        if mes_key:
            conn.execute("DELETE FROM dest_rows WHERE mes=?", (mes_key,))
        else:
            conn.execute("DELETE FROM valor_fijo")
            conn.execute("DELETE FROM sources")
            conn.execute("DELETE FROM dest_rows")

        for item in cfg.get('items', []):
            hoja       = item.get('hoja', 'GASTOS ADMINISTRATIVOS')
            codigo_a   = item.get('codigo_a', '')
            fila_dest  = item.get('fila_dest')
            buscar_por = item.get('buscar_por', 'A')

            cur = conn.execute(
                "INSERT INTO dest_rows (hoja,codigo_a,fila_dest,buscar_por,mes) VALUES (?,?,?,?,?)",
                (hoja, codigo_a, fila_dest, buscar_por, mes_key)
            )
            row_id = cur.lastrowid

            for i, src in enumerate(item.get('sources', [])):
                conn.execute(
                    "INSERT INTO sources "
                    "(dest_row_id,file_key,codes,op,sin_filtro,seccion,orden) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (
                        row_id,
                        src['key'],
                        json.dumps(src.get('codes', []), ensure_ascii=False),
                        src.get('op', '+'),
                        1 if src.get('sin_filtro') else 0,
                        src.get('seccion'),
                        i,
                    )
                )

            vfs = item.get('valor_fijo') or []
            if isinstance(vfs, dict):
                vfs = [vfs]
            for i, vf in enumerate(vfs):
                conn.execute(
                    "INSERT INTO valor_fijo (dest_row_id,op,valor,orden) VALUES (?,?,?,?)",
                    (row_id, vf.get('op', '+'), float(vf.get('valor', 0)), i)
                )


def load_config(mes=None):
    """
    Lee la config de la DB.
    mes='MAYO' → solo items de ese mes (panel vacío si nunca se configuró).
    mes=None   → todos los items (legado / procesar_todo sin mes especificado).
    """
    mes_key = mes.upper() if mes else None
    with get_conn() as conn:
        file_sources = {
            r['key']: {'prefijo': r['prefijo'], 'entidad': r['entidad']}
            for r in conn.execute("SELECT * FROM file_sources").fetchall()
        }

        items = []
        query = ("SELECT * FROM dest_rows WHERE mes=? ORDER BY id" if mes_key
                 else "SELECT * FROM dest_rows ORDER BY id")
        params = (mes_key,) if mes_key else ()
        for dr in conn.execute(query, params).fetchall():
            item = {
                'codigo_a'  : dr['codigo_a'],
                'hoja'      : dr['hoja'],
                'buscar_por': dr['buscar_por'],
            }
            if dr['fila_dest'] is not None:
                item['fila_dest'] = dr['fila_dest']

            srcs = conn.execute(
                "SELECT * FROM sources WHERE dest_row_id=? ORDER BY orden",
                (dr['id'],)
            ).fetchall()

            sources_list = []
            for s in srcs:
                sd = {'key': s['file_key'], 'codes': json.loads(s['codes'])}
                if s['op'] != '+':
                    sd['op'] = s['op']
                if s['sin_filtro']:
                    sd['sin_filtro'] = True
                if s['seccion']:
                    sd['seccion'] = s['seccion']
                sources_list.append(sd)
            item['sources'] = sources_list

            vfs = conn.execute(
                "SELECT * FROM valor_fijo WHERE dest_row_id=? ORDER BY orden",
                (dr['id'],)
            ).fetchall()
            if vfs:
                item['valor_fijo'] = [{'op': v['op'], 'valor': v['valor']} for v in vfs]

            items.append(item)

    return {'file_sources': file_sources, 'items': items}


# ---------------------------------------------------------------------------
# Guardar / cargar config_5105
# ---------------------------------------------------------------------------

def save_config_5105(cfg5):
    """Guarda el dict config_5105 (formato config_5105.json) en la DB."""
    with get_conn() as conn:
        conn.execute("DELETE FROM config_5105")
        for entidad, reglas in cfg5.items():
            conn.execute(
                "INSERT INTO config_5105 (entidad,only_code,subtract_codes) VALUES (?,?,?)",
                (
                    entidad,
                    reglas.get('only_code'),
                    json.dumps(reglas.get('subtract_codes', []), ensure_ascii=False),
                )
            )


def load_config_5105():
    """Retorna config_5105 en el mismo formato que config_5105.json."""
    with get_conn() as conn:
        result = {}
        for r in conn.execute("SELECT * FROM config_5105").fetchall():
            entry = {}
            if r['only_code']:
                entry['only_code'] = r['only_code']
            if r['subtract_codes']:
                entry['subtract_codes'] = json.loads(r['subtract_codes'])
            result[r['entidad']] = entry
        return result


# ---------------------------------------------------------------------------
# Log de ejecuciones
# ---------------------------------------------------------------------------

def log_ejecucion(mes, exito, valores, notas=None):
    """
    Registra una ejecución en la DB.

    valores: lista de dicts con estructura:
      {
        'hoja'       : str,
        'fila'       : int,
        'codigo_a'   : str,
        'valor_total': float,
        'fuentes'    : [{'file_key', 'codigo', 'seccion', 'op', 'valor'}, ...]
      }
    """
    fecha = datetime.now().isoformat(sep=' ', timespec='seconds')
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO ejecuciones (mes,fecha,exito,notas) VALUES (?,?,?,?)",
            (mes, fecha, 1 if exito else 0, notas)
        )
        ejec_id = cur.lastrowid

        for v in valores:
            cur2 = conn.execute(
                "INSERT INTO valores_escritos "
                "(ejecucion_id,hoja,fila,codigo_a,valor_total) VALUES (?,?,?,?,?)",
                (ejec_id, v['hoja'], v['fila'], v['codigo_a'], v['valor_total'])
            )
            val_id = cur2.lastrowid

            for f in v.get('fuentes', []):
                conn.execute(
                    "INSERT INTO fuentes_valor "
                    "(valor_id,file_key,codigo,seccion,op,valor) VALUES (?,?,?,?,?,?)",
                    (
                        val_id,
                        f['file_key'],
                        f['codigo'],
                        f.get('seccion'),
                        f.get('op', '+'),
                        f['valor'],
                    )
                )


def get_historial(mes=None):
    """
    Retorna el historial de ejecuciones, opcionalmente filtrado por mes.
    Cada ejecución incluye sus valores escritos y las fuentes por valor.
    """
    with get_conn() as conn:
        if mes:
            ejecuciones = conn.execute(
                "SELECT * FROM ejecuciones WHERE mes=? ORDER BY fecha DESC LIMIT 50",
                (mes,)
            ).fetchall()
        else:
            ejecuciones = conn.execute(
                "SELECT * FROM ejecuciones ORDER BY fecha DESC LIMIT 50"
            ).fetchall()

        result = []
        for ej in ejecuciones:
            valores = conn.execute(
                "SELECT * FROM valores_escritos WHERE ejecucion_id=? ORDER BY hoja,fila",
                (ej['id'],)
            ).fetchall()

            vals_list = []
            for v in valores:
                fuentes = conn.execute(
                    "SELECT * FROM fuentes_valor WHERE valor_id=?",
                    (v['id'],)
                ).fetchall()
                vals_list.append({
                    'hoja'       : v['hoja'],
                    'fila'       : v['fila'],
                    'codigo_a'   : v['codigo_a'],
                    'valor_total': v['valor_total'],
                    'fuentes'    : [dict(f) for f in fuentes],
                })

            result.append({
                'id'     : ej['id'],
                'mes'    : ej['mes'],
                'fecha'  : ej['fecha'],
                'exito'  : bool(ej['exito']),
                'notas'  : ej['notas'],
                'valores': vals_list,
            })

        return result


def diff_ejecuciones(id_a, id_b):
    """
    Compara dos ejecuciones por (hoja, fila).
    Retorna lista de {hoja, fila, codigo_a, valor_a, valor_b, delta}.
    """
    with get_conn() as conn:
        def _get(ejec_id):
            rows = conn.execute(
                "SELECT hoja, fila, codigo_a, valor_total FROM valores_escritos "
                "WHERE ejecucion_id=?", (ejec_id,)
            ).fetchall()
            return {(r['hoja'], r['fila']): dict(r) for r in rows}

        mapa_a = _get(id_a)
        mapa_b = _get(id_b)
        todas  = sorted(set(mapa_a) | set(mapa_b))

        result = []
        for key in todas:
            ra = mapa_a.get(key)
            rb = mapa_b.get(key)
            va = ra['valor_total'] if ra else None
            vb = rb['valor_total'] if rb else None
            result.append({
                'hoja'    : key[0],
                'fila'    : key[1],
                'codigo_a': (ra or rb)['codigo_a'],
                'valor_a' : va,
                'valor_b' : vb,
                'delta'   : (vb - va) if (va is not None and vb is not None) else None,
            })
        return result

