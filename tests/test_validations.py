import pandas as pd
import pytest

from modules.config import AppConfig
from modules import validations


def _row(**overrides):
    base = {
        "_fila_excel": 2,
        "DNI": "12345678",
        "Nombres": "Juan",
        "Apellido paterno": "Perez",
        "Apellido materno": "Lopez",
        "Fecha de nacimiento": "1990-01-01",
        "Numero de carnet": "C-00001",
        "Fecha de emision": "2026-01-01",
        "Fecha de vencimiento": "2027-01-01",
        "Empresa": "Empresa X",
        "Puesto": "Cocinero",
        "Fotografia": "",
    }
    base.update(overrides)
    return base


def _df(rows):
    return pd.DataFrame(rows)


@pytest.fixture
def config():
    return AppConfig()


def test_registro_correcto(config):
    df = _df([_row()])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido


def test_dni_vacio(config):
    df = _df([_row(DNI="")])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido
    assert any("DNI vacio" in e for e in records[0].errores)


def test_dni_duplicado_en_archivo(config):
    df = _df([_row(_fila_excel=2), _row(_fila_excel=3)])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido  # primera aparicion queda correcta
    assert not records[1].es_valido
    assert any("duplicado" in e for e in records[1].errores)


def test_fecha_invalida(config):
    df = _df([_row(**{"Fecha de nacimiento": "no-es-fecha"})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido
    assert any("invalida" in e for e in records[0].errores)


def test_vencimiento_anterior_a_emision(config):
    df = _df([_row(**{"Fecha de emision": "2027-01-01", "Fecha de vencimiento": "2026-01-01"})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido
    assert any("anterior" in e for e in records[0].errores)


def test_numero_carnet_repetido(config):
    df = _df(
        [
            _row(DNI="11111111", _fila_excel=2, **{"Numero de carnet": "C-1"}),
            _row(DNI="22222222", _fila_excel=3, **{"Numero de carnet": "C-1"}),
        ]
    )
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido
    assert not records[1].es_valido
    assert any("duplicado" in e for e in records[1].errores)


def test_fotografia_faltante_cuando_es_requerida(config):
    df = _df([_row(Fotografia="")])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=True)
    assert not records[0].es_valido
    assert any("Fotografia" in e for e in records[0].errores)


def test_fotografia_no_requerida_no_bloquea(config):
    df = _df([_row(Fotografia="")])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido


def test_nombres_muy_largos(config):
    df = _df(
        [
            _row(
                Nombres="Nombre" * 10,
                **{"Apellido paterno": "Paterno" * 5, "Apellido materno": "Materno" * 5},
            )
        ]
    )
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido
    assert any("largo" in e for e in records[0].errores)


def test_dni_previamente_procesado(config):
    df = _df([_row(DNI="99999999")])
    records = validations.validate_dataframe(
        df, config, plantilla_requiere_foto=False, dnis_previamente_procesados={"99999999"}
    )
    assert not records[0].es_valido
    assert any("previamente procesado" in e for e in records[0].errores)


def test_split_records(config):
    df = _df([_row(_fila_excel=2), _row(_fila_excel=3, DNI="")])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    correctos, observados = validations.split_records(records)
    assert len(correctos) == 1
    assert len(observados) == 1


def test_numeracion_vacia_sin_autogenerar_es_error(config):
    config.numbering.autogenerar = False
    df = _df([_row(**{"Numero de carnet": ""})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido


def test_numeracion_vacia_con_autogenerar_no_es_error(config):
    config.numbering.autogenerar = True
    df = _df([_row(**{"Numero de carnet": ""})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido


def test_vencimiento_vacio_sin_autocalcular_es_error(config):
    config.vigencia.autocalcular = False
    df = _df([_row(**{"Fecha de vencimiento": ""})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert not records[0].es_valido


def test_vencimiento_vacio_con_autocalcular_se_calcula(config):
    config.vigencia.autocalcular = True
    config.vigencia.meses_vigencia = 12
    df = _df([_row(**{"Fecha de vencimiento": ""})])
    records = validations.validate_dataframe(df, config, plantilla_requiere_foto=False)
    assert records[0].es_valido
    assert records[0].fecha_vencimiento is not None
    assert records[0].fecha_vencimiento.year == records[0].fecha_emision.year + 1
