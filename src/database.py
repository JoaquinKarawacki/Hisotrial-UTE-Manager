"""
database.py — Gestión de clientes y historial de descargas (SQLite).

Tablas:
    clientes     — datos de login UTE por cliente
    descargas    — registro de todas las descargas realizadas

El archivo .db se guarda en el mismo directorio que este script.
"""

import os
import sqlite3
import base64
from datetime import datetime
from typing import Optional, List

# Ruta del archivo de base de datos
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "seg_historial.db")


# ── Ofuscación simple de contraseñas ──────────────────────────────────────────

def _encode_pass(raw: str) -> str:
    return base64.b64encode(raw.encode()).decode()

def _decode_pass(encoded: str) -> str:
    try:
        return base64.b64decode(encoded.encode()).decode()
    except Exception:
        return encoded


# ── Conexión ─────────────────────────────────────────────────────────────────

def get_conn() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


# ── Inicialización ───────────────────────────────────────────────────────────

def init_db():
    """Crea las tablas si no existen."""
    with get_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS clientes (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre      TEXT    NOT NULL,
                usuario     TEXT    NOT NULL,
                password    TEXT    NOT NULL,
                carpeta     TEXT    DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS subcuentas (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id  INTEGER NOT NULL REFERENCES clientes(id) ON DELETE CASCADE,
                numero      TEXT    NOT NULL,
                creada_en   TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(cliente_id, numero)
            );

            CREATE TABLE IF NOT EXISTS descargas (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id      INTEGER REFERENCES clientes(id) ON DELETE CASCADE,
                usuario         TEXT NOT NULL,
                directorio      TEXT,
                fecha_inicio    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                fecha_fin       TIMESTAMP,
                estado          TEXT DEFAULT 'en_progreso',
                notas           TEXT DEFAULT ''
            );
        """)
        conn.commit()


# ── CLIENTES ──────────────────────────────────────────────────────────────────

def get_clientes() -> list[dict]:
    """Obtiene todos los clientes."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, nombre, usuario, carpeta FROM clientes ORDER BY nombre"
        ).fetchall()
    return [dict(r) for r in rows]


def get_cliente(cliente_id: int) -> Optional[dict]:
    """Obtiene un cliente específico con contraseña decodificada."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id, nombre, usuario, password, carpeta FROM clientes WHERE id=?",
            (cliente_id,)
        ).fetchone()
    if not row:
        return None
    d = dict(row)
    d["password_decoded"] = _decode_pass(d["password"])
    return d


def crear_cliente(nombre: str, usuario: str, password: str, carpeta: str = "") -> int:
    """Crea un nuevo cliente."""
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO clientes (nombre, usuario, password, carpeta) VALUES (?,?,?,?)",
            (nombre, usuario, _encode_pass(password), carpeta)
        )
        conn.commit()
        return cur.lastrowid


def actualizar_cliente(cliente_id: int, nombre: str, usuario: str, password: str = "", carpeta: str = ""):
    """Actualiza datos de un cliente. Si password viene vacío, conserva la contraseña actual."""
    with get_conn() as conn:
        if password:
            conn.execute(
                "UPDATE clientes SET nombre=?, usuario=?, password=?, carpeta=? WHERE id=?",
                (nombre, usuario, _encode_pass(password), carpeta, cliente_id)
            )
        else:
            conn.execute(
                "UPDATE clientes SET nombre=?, usuario=?, carpeta=? WHERE id=?",
                (nombre, usuario, carpeta, cliente_id)
            )
        conn.commit()


def eliminar_cliente(cliente_id: int):
    """Elimina un cliente (y sus descargas asociadas por FK CASCADE)."""
    with get_conn() as conn:
        conn.execute("DELETE FROM clientes WHERE id=?", (cliente_id,))
        conn.commit()


# ── SUBCUENTAS ───────────────────────────────────────────────────────────────

def get_subcuentas(cliente_id: int) -> list[dict]:
    """Obtiene las subcuentas configuradas de un cliente."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, cliente_id, numero FROM subcuentas WHERE cliente_id=? ORDER BY numero",
            (cliente_id,)
        ).fetchall()
    return [dict(r) for r in rows]


def crear_subcuenta(cliente_id: int, numero: str) -> int:
    """Crea una subcuenta para un cliente. Retorna el ID."""
    numero = (numero or "").strip()
    if not numero:
        raise ValueError("El numero de subcuenta es requerido")

    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO subcuentas (cliente_id, numero) VALUES (?, ?)",
            (cliente_id, numero)
        )
        conn.commit()
        return cur.lastrowid


def eliminar_subcuenta(subcuenta_id: int):
    """Elimina una subcuenta por ID."""
    with get_conn() as conn:
        conn.execute("DELETE FROM subcuentas WHERE id=?", (subcuenta_id,))
        conn.commit()


# ── DESCARGAS ─────────────────────────────────────────────────────────────────

def registrar_descarga(cliente_id: int, usuario: str, directorio: str = None) -> int:
    """Inicia un registro de descarga. Retorna el ID."""
    with get_conn() as conn:
        cur = conn.execute(
            """INSERT INTO descargas (cliente_id, usuario, directorio, estado)
               VALUES (?, ?, ?, 'en_progreso')""",
            (cliente_id, usuario, directorio or "")
        )
        conn.commit()
        return cur.lastrowid


def finalizar_descarga(descarga_id: int, estado: str = "exitosa", notas: str = None):
    """Marca una descarga como finalizada."""
    with get_conn() as conn:
        conn.execute(
            """UPDATE descargas
               SET estado = ?, fecha_fin = CURRENT_TIMESTAMP, notas = ?
               WHERE id = ?""",
            (estado, notas or "", descarga_id)
        )
        conn.commit()


def obtener_historial(cliente_id: int = None, limite: int = 50) -> List[dict]:
    """Obtiene los últimos N registros de descargas (opcionalmente filtrados por cliente)."""
    with get_conn() as conn:
        if cliente_id:
            rows = conn.execute(
                """SELECT id, cliente_id, usuario, directorio, fecha_inicio, fecha_fin, estado, notas
                   FROM descargas
                   WHERE cliente_id = ?
                   ORDER BY fecha_inicio DESC
                   LIMIT ?""",
                (cliente_id, limite)
            ).fetchall()
        else:
            rows = conn.execute(
                """SELECT id, cliente_id, usuario, directorio, fecha_inicio, fecha_fin, estado, notas
                   FROM descargas
                   ORDER BY fecha_inicio DESC
                   LIMIT ?""",
                (limite,)
            ).fetchall()
    return [dict(row) for row in rows]


def obtener_descarga(descarga_id: int) -> Optional[dict]:
    """Obtiene un registro específico de descarga."""
    with get_conn() as conn:
        row = conn.execute(
            """SELECT id, cliente_id, usuario, directorio, fecha_inicio, fecha_fin, estado, notas
               FROM descargas
               WHERE id = ?""",
            (descarga_id,)
        ).fetchone()
    return dict(row) if row else None


def limpiar_historial_antiguo(dias: int = 30):
    """Borra descargas más antiguas que N días."""
    with get_conn() as conn:
        conn.execute(
            """DELETE FROM descargas
               WHERE fecha_inicio < datetime('now', '-' || ? || ' days')""",
            (dias,)
        )
        conn.commit()


def get_estado_completo() -> list[dict]:
    """Devuelve todos los clientes con su información y últimas descargas."""
    clientes = get_clientes()
    for c in clientes:
        c["descargas"] = obtener_historial(c["id"], limite=5)
        c["subcuentas"] = get_subcuentas(c["id"])
        c["password_decoded"] = ""  # No incluir passwords en esta respuesta
    return clientes


# ── Inicializar al importar ──────────────────────────────────────────────────

init_db()
