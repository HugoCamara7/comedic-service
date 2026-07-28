import io

import pandas as pd
import pytest

from modules import excel_reader


def _make_excel_bytes(rows, columns=None):
    columns = columns or excel_reader.COLUMNS
    df = pd.DataFrame(rows, columns=columns)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Personas")
    buf.seek(0)
    return buf


def test_template_bytes_has_expected_headers():
    data = excel_reader.template_bytes()
    df = pd.read_excel(io.BytesIO(data), sheet_name="Personas")
    assert list(df.columns) == excel_reader.COLUMNS


def test_load_excel_correcto(tmp_path):
    row = {c: "" for c in excel_reader.COLUMNS}
    row.update(
        {
            "DNI": "12345678",
            "Nombres": "Juan",
            "Apellido paterno": "Perez",
            "Apellido materno": "Lopez",
            "Fecha de nacimiento": "1990-01-01",
            "Fecha de emision": "2026-01-01",
        }
    )
    buf = _make_excel_bytes([row])
    result = excel_reader.load_excel(buf, "test.xlsx")
    assert result.missing_columns == []
    assert len(result.dataframe) == 1
    assert result.dataframe.iloc[0]["DNI"] == "12345678"


def test_load_excel_columnas_faltantes():
    columns = [c for c in excel_reader.COLUMNS if c != "Fecha de nacimiento"]
    row = {c: "x" for c in columns}
    buf = _make_excel_bytes([row], columns=columns)
    result = excel_reader.load_excel(buf, "test.xlsx")
    assert "Fecha de nacimiento" in result.missing_columns


def test_load_excel_extension_invalida():
    buf = io.BytesIO(b"not an excel file")
    with pytest.raises(excel_reader.ExcelFormatError):
        excel_reader.load_excel(buf, "test.csv")


def test_load_excel_tamano_excedido():
    row = {c: "x" for c in excel_reader.COLUMNS}
    buf = _make_excel_bytes([row])
    with pytest.raises(excel_reader.ExcelFormatError):
        excel_reader.load_excel(buf, "test.xlsx", max_size_mb=0)
