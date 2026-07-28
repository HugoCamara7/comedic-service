from pathlib import Path

import pytest

from modules import history


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    """Aisla cada test con su propia base SQLite temporal."""
    db_path = tmp_path / "history.db"
    monkeypatch.setattr(history, "DB_PATH", db_path)
    history.init_db()
    yield


def test_save_and_list_generation():
    record = history.GenerationRecord(
        source_filename="test.xlsx",
        total_records=3,
        correct_records=2,
        observed_records=1,
        params={"columnas": 2, "filas": 2},
    )
    people = [
        {"dni": "11111111", "numero_carnet": "C-1", "nombre_completo": "A", "fecha_emision": "01/01/2026", "fecha_vencimiento": "01/01/2027"},
        {"dni": "22222222", "numero_carnet": "C-2", "nombre_completo": "B", "fecha_emision": "01/01/2026", "fecha_vencimiento": "01/01/2027"},
    ]
    gen_id = history.save_generation(record, people)
    assert gen_id > 0

    rows = history.list_generations()
    assert len(rows) == 1
    assert rows[0]["source_filename"] == "test.xlsx"

    stored_people = history.get_generation_people(gen_id)
    assert len(stored_people) == 2


def test_get_all_processed_dnis():
    record = history.GenerationRecord("a.xlsx", 1, 1, 0)
    history.save_generation(record, [{"dni": "12345678", "numero_carnet": "C-1", "nombre_completo": "X"}])
    dnis = history.get_all_processed_dnis()
    assert "12345678" in dnis


def test_get_all_used_carnet_numbers():
    record = history.GenerationRecord("a.xlsx", 1, 1, 0)
    history.save_generation(record, [{"dni": "12345678", "numero_carnet": "C-99", "nombre_completo": "X"}])
    numbers = history.get_all_used_carnet_numbers()
    assert "C-99" in numbers


def test_update_generation_paths():
    record = history.GenerationRecord("a.xlsx", 1, 1, 0)
    gen_id = history.save_generation(record, [])
    history.update_generation_paths(gen_id, "/tmp/x.docx", "/tmp/x.pdf")
    row = history.get_generation(gen_id)
    assert row["docx_path"] == "/tmp/x.docx"
    assert row["pdf_path"] == "/tmp/x.pdf"


def test_summary_counts():
    record = history.GenerationRecord("a.xlsx", 2, 2, 0)
    history.save_generation(
        record,
        [
            {"dni": "1", "numero_carnet": "C1", "nombre_completo": "A"},
            {"dni": "2", "numero_carnet": "C2", "nombre_completo": "B"},
        ],
    )
    summary = history.summary_counts()
    assert summary["total_personas"] == 2
    assert summary["total_generaciones"] == 1
