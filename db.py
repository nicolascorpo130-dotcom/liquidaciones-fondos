"""Base de datos SQLite de Liquidaciones y Fondos."""
import os
import sqlite3

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, "datos.db")
COMPROBANTES_DIR = os.path.join(BASE, "comprobantes")

# Conceptos que se le pueden deber a cada persona (orden = orden de columnas en Pagos)
TIPOS = [
    ("LIQUIDACION", "Liquidación final"),
    ("FCL", "Fondo de cese laboral (FCL)"),
    ("AGUINALDO", "Aguinaldo"),
    ("BONO", "Bonos ADUA"),
]
TIPOS_NOMBRE = dict(TIPOS)
TIPOS_CORTO = {"LIQUIDACION": "Liquidación", "FCL": "FCL", "AGUINALDO": "Aguinaldo", "BONO": "Bonos"}

ESQUEMA = """
CREATE TABLE IF NOT EXISTS personal (
    id INTEGER PRIMARY KEY,
    cuil TEXT UNIQUE,
    nombre TEXT NOT NULL,
    fecha_baja TEXT,
    obra TEXT,
    cuenta_sueldo TEXT,
    cuenta_fondo TEXT,
    telegrama TEXT,
    notas TEXT,
    creado TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY,
    personal_id INTEGER NOT NULL REFERENCES personal(id) ON DELETE CASCADE,
    tipo TEXT NOT NULL,
    periodo TEXT,
    monto REAL,
    estado TEXT NOT NULL DEFAULT 'PENDIENTE',
    fecha_pago TEXT,
    comprobante TEXT,
    comprobante_nombre TEXT,
    notas TEXT,
    creado TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_items_personal ON items(personal_id);
"""


def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    return c


def init():
    os.makedirs(COMPROBANTES_DIR, exist_ok=True)
    with conn() as c:
        c.executescript(ESQUEMA)


def vacia():
    with conn() as c:
        return c.execute("SELECT COUNT(*) FROM personal").fetchone()[0] == 0
