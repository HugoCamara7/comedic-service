"""
history.py
----------
Historial local (SQLite) de generaciones de carnets.

Guarda metadatos de cada generacion (no datos personales sensibles fuera de
lo necesario para trazabilidad: DNI y numero de carnet, que son necesarios
para detectar "DNI previamente procesado" y permitir reimpresion).
"""
from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "history.db"
DATA_DIR.mkdir(parents=True, exist_ok=True)

SCHEMA = """
CREATE TABLE IF NOT EXISTS generations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    source_filename TEXT NOT NULL,
    total_records INTEGER NOT NULL,
    correct_records INTEGER NOT NULL,
    observed_records INTEGER NOT NULL,
    docx_path TEXT,
    pdf_path TEXT,
    params_json TEXT
);

CREATE TABLE IF NOT EXISTS processed_people (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    generation_id INTEGER NOT NULL,
    dni TEXT NOT NULL,
    numero_carnet TEXT,
    nombre_completo TEXT,
    fecha_emision TEXT,
    fecha_vencimiento TEXT,
    FOREIGN KEY (generation_id) REFERENCES generations(id)
);

CREATE INDEX IF NOT EXISTS idx_processed_dni ON processed_people(dni);
"""


@contextmanager
def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_connection() as conn:
        conn.executescript(SCHEMA)


@dataclass
class GenerationRecord:
    source_filename: str
    total_records: int
    correct_records: int
    observed_records: int
    docx_path: Optional[str] = None
    pdf_path: Optional[str] = None
    params: Optional[dict] = None


def save_generation(record: GenerationRecord, personas: List[dict]) -> int:
    """Guarda una generacion y las personas procesadas. Devuelve el id."""
    init_db()
    with get_connection() as conn:
        cur = conn.execute(
            """INSERT INTO generations
               (timestamp, source_filename, total_records, correct_records,
                observed_records, docx_path, pdf_path, params_json)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                datetime.now().isoformat(timespec="seconds"),
                record.source_filename,
                record.total_records,
                record.correct_records,
                record.observed_records,
                record.docx_path,
                record.pdf_path,
                json.dumps(record.params or {}, ensure_ascii=False),
            ),
        )
        generation_id = cur.lastrowid
        for p in personas:
            conn.execute(
                """INSERT INTO processed_people
                   (generation_id, dni, numero_carnet, nombre_completo,
                    fecha_emision, fecha_vencimiento)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (
                    generation_id,
                    p.get("dni", ""),
                    p.get("numero_carnet", ""),
                    p.get("nombre_completo", ""),
                    p.get("fecha_emision", ""),
                    p.get("fecha_vencimiento", ""),
                ),
            )
        return generation_id


def update_generation_paths(generation_id: int, docx_path: str, pdf_path: str) -> None:
    init_db()
    with get_connection() as conn:
        conn.execute(
            "UPDATE generations SET docx_path = ?, pdf_path = ? WHERE id = ?",
            (docx_path, pdf_path, generation_id),
        )


def list_generations(limit: int = 100) -> List[sqlite3.Row]:
    init_db()
    with get_connection() as conn:
        cur = conn.execute(
            "SELECT * FROM generations ORDER BY id DESC LIMIT ?", (limit,)
        )
        return cur.fetchall()


def get_generation(generation_id: int) -> Optional[sqlite3.Row]:
    init_db()
    with get_connection() as conn:
        cur = conn.execute("SELECT * FROM generations WHERE id = ?", (generation_id,))
        return cur.fetchone()


def get_generation_people(generation_id: int) -> List[sqlite3.Row]:
    init_db()
    with get_connection() as conn:
        cur = conn.execute(
            "SELECT * FROM processed_people WHERE generation_id = ?", (generation_id,)
        )
        return cur.fetchall()


def get_all_processed_dnis() -> set:
    init_db()
    with get_connection() as conn:
        cur = conn.execute("SELECT DISTINCT dni FROM processed_people")
        return {row["dni"] for row in cur.fetchall()}


def get_all_used_carnet_numbers() -> set:
    init_db()
    with get_connection() as conn:
        cur = conn.execute(
            "SELECT DISTINCT numero_carnet FROM processed_people WHERE numero_carnet != ''"
        )
        return {row["numero_carnet"] for row in cur.fetchall()}


def summary_counts() -> dict:
    init_db()
    with get_connection() as conn:
        total = conn.execute("SELECT COUNT(*) c FROM processed_people").fetchone()["c"]
        generations = conn.execute("SELECT COUNT(*) c FROM generations").fetchone()["c"]
        return {"total_personas": total, "total_generaciones": generations}
