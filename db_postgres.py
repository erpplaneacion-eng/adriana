"""
Capa de acceso a Postgres.
- chips_xls        : caché de los valores leídos de cada XLS por mes
- informe_resultados: valores escritos al INFORME REAL por mes (qué quedó en cada celda)
"""
import os, json
import psycopg2
import psycopg2.extras


def _conn():
    return psycopg2.connect(os.environ['DATABASE_URL'])


def init_pg():
    """Crea las tablas si no existen."""
    with _conn() as conn, conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS chips_xls (
                mes         TEXT    NOT NULL,
                archivo     TEXT    NOT NULL,
                file_key    TEXT,
                idx         INTEGER NOT NULL,
                seccion     TEXT,
                codigo      TEXT,
                descripcion TEXT,
                valor       NUMERIC(18,2) DEFAULT 0,
                debito      NUMERIC(18,2) DEFAULT 0,
                credito     NUMERIC(18,2) DEFAULT 0,
                op_jk       BOOLEAN DEFAULT FALSE,
                negrita     BOOLEAN DEFAULT FALSE,
                indent      INTEGER DEFAULT 0,
                separador   BOOLEAN DEFAULT FALSE,
                PRIMARY KEY (mes, archivo, idx)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS informe_resultados (
                id           BIGSERIAL PRIMARY KEY,
                mes          TEXT NOT NULL,
                ejecutado_en TIMESTAMPTZ DEFAULT NOW(),
                hoja         TEXT NOT NULL,
                fila         INTEGER NOT NULL,
                codigo_a     TEXT,
                valor_total  NUMERIC(18,2),
                fuentes      JSONB
            )
        """)
        conn.commit()


def tiene_chips(mes, archivo):
    with _conn() as conn, conn.cursor() as cur:
        cur.execute(
            "SELECT 1 FROM chips_xls WHERE mes=%s AND archivo=%s LIMIT 1",
            (mes, archivo)
        )
        return cur.fetchone() is not None


def guardar_chips(mes, archivo, file_key, items):
    """Reemplaza los chips de (mes, archivo) con la lista nueva."""
    with _conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM chips_xls WHERE mes=%s AND archivo=%s", (mes, archivo))
        psycopg2.extras.execute_batch(cur, """
            INSERT INTO chips_xls
                (mes, archivo, file_key, idx, seccion, codigo, descripcion,
                 valor, debito, credito, op_jk, negrita, indent, separador)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, [
            (
                mes, archivo, file_key, i,
                item.get('seccion'),
                item.get('codigo', ''),
                item.get('descripcion', ''),
                float(item.get('valor',  0) or 0),
                float(item.get('debito', 0) or 0),
                float(item.get('credito',0) or 0),
                bool(item.get('op_jk',    False)),
                bool(item.get('negrita',  False)),
                int( item.get('indent',   0) or 0),
                bool(item.get('separador',False)),
            )
            for i, item in enumerate(items)
        ])
        conn.commit()


def cargar_chips(mes, archivo):
    """Devuelve los items en el mismo formato que leer_todos_codigos."""
    with _conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute("""
            SELECT seccion, codigo, descripcion, valor, debito, credito,
                   op_jk, negrita, indent, separador
            FROM chips_xls
            WHERE mes=%s AND archivo=%s
            ORDER BY idx
        """, (mes, archivo))
        return [dict(r) for r in cur.fetchall()]


def archivos_en_cache(mes):
    """Nombres de archivo ya cacheados para este mes."""
    with _conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT DISTINCT archivo FROM chips_xls WHERE mes=%s", (mes,))
        return {r[0] for r in cur.fetchall()}


# ---------------------------------------------------------------------------
# informe_resultados — qué quedó escrito en el INFORME REAL por mes
# ---------------------------------------------------------------------------

def guardar_resultados_mes(mes, items):
    """Reemplaza los resultados del mes con los valores recién escritos al INFORME."""
    with _conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM informe_resultados WHERE mes=%s", (mes,))
        psycopg2.extras.execute_batch(cur, """
            INSERT INTO informe_resultados
                (mes, hoja, fila, codigo_a, valor_total, fuentes)
            VALUES (%s,%s,%s,%s,%s,%s)
        """, [
            (
                mes,
                item['hoja'],
                int(item['fila']),
                item.get('codigo_a'),
                float(item.get('valor_total', 0) or 0),
                json.dumps(item.get('fuentes', []), ensure_ascii=False),
            )
            for item in items
        ])
        conn.commit()


def cargar_resultados_mes(mes):
    """Devuelve los valores escritos al INFORME para ese mes."""
    with _conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute("""
            SELECT hoja, fila, codigo_a, valor_total, fuentes, ejecutado_en
            FROM informe_resultados
            WHERE mes=%s
            ORDER BY hoja, fila
        """, (mes,))
        rows = []
        for r in cur.fetchall():
            d = dict(r)
            d['fuentes'] = d['fuentes'] if isinstance(d['fuentes'], list) else json.loads(d['fuentes'] or '[]')
            rows.append(d)
        return rows
