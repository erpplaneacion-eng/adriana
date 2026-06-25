"""
Capa de acceso a Postgres para caché de chips XLS.
Tabla: chips_xls — una fila por item de cada archivo XLS, indexada por (mes, archivo, idx).
"""
import os
import psycopg2
import psycopg2.extras


def _conn():
    return psycopg2.connect(os.environ['DATABASE_URL'])


def init_pg():
    """Crea la tabla si no existe."""
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
