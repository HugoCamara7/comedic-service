"""
config.py
---------
Configuracion persistente de la aplicacion (impresion, numeracion, vigencia).

Se guarda en data/config.json. No trae valores de negocio "adivinados": los
campos de numeracion y vigencia parten en None/deshabilitados hasta que el
usuario los define explicitamente en la pantalla de Configuracion.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
CONFIG_PATH = DATA_DIR / "config.json"

DATA_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class PrintConfig:
    # Papel
    paper_size: str = "A4"          # A4 | Carta | A3
    orientation: str = "vertical"   # vertical | horizontal

    # Grilla de carnets por hoja
    columnas: int = 2
    filas: int = 5

    # Tamano de cada carnet (mm) - referencial, lo define la plantilla Word,
    # pero se usa para calcular la grilla de la hoja maestra.
    carnet_ancho_mm: float = 85.6
    carnet_alto_mm: float = 54.0

    # Margenes de hoja (mm)
    margen_superior_mm: float = 10.0
    margen_inferior_mm: float = 10.0
    margen_izquierdo_mm: float = 10.0
    margen_derecho_mm: float = 10.0

    # Separacion entre carnets (mm)
    separacion_horizontal_mm: float = 4.0
    separacion_vertical_mm: float = 4.0

    # Calibracion / ajuste fino de impresora (mm) - desplaza toda la grilla
    desplazamiento_horizontal_mm: float = 0.0
    desplazamiento_vertical_mm: float = 0.0

    lineas_de_corte: bool = True
    escala_impresion: int = 100  # % - debe imprimirse siempre al 100%


@dataclass
class NumberingConfig:
    # Si esta deshabilitado, el usuario DEBE llenar "Numero de carnet" en el Excel.
    autogenerar: bool = False
    prefijo: str = ""
    # Cantidad de digitos del correlativo (con ceros a la izquierda)
    digitos: int = 5
    siguiente_correlativo: int = 1
    # Patron ilustrativo: {prefijo}-{correlativo}. No se asume ningun otro
    # formato salvo que el usuario lo configure aqui.


@dataclass
class VigenciaConfig:
    # Si esta deshabilitado, el usuario DEBE llenar "Fecha de vencimiento".
    autocalcular: bool = False
    meses_vigencia: int = 12


@dataclass
class AppConfig:
    print: PrintConfig = field(default_factory=PrintConfig)
    numbering: NumberingConfig = field(default_factory=NumberingConfig)
    vigencia: VigenciaConfig = field(default_factory=VigenciaConfig)
    dni_regex: str = r"^\d{8}$"  # DNI peruano: 8 digitos. Editable en Configuracion.
    max_excel_size_mb: int = 10
    max_word_template_size_mb: int = 15
    max_photo_size_mb: int = 5
    max_zip_size_mb: int = 50


def load_config() -> AppConfig:
    if not CONFIG_PATH.exists():
        cfg = AppConfig()
        save_config(cfg)
        return cfg
    raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    return AppConfig(
        print=PrintConfig(**raw.get("print", {})),
        numbering=NumberingConfig(**raw.get("numbering", {})),
        vigencia=VigenciaConfig(**raw.get("vigencia", {})),
        dni_regex=raw.get("dni_regex", AppConfig().dni_regex),
        max_excel_size_mb=raw.get("max_excel_size_mb", 10),
        max_word_template_size_mb=raw.get("max_word_template_size_mb", 15),
        max_photo_size_mb=raw.get("max_photo_size_mb", 5),
        max_zip_size_mb=raw.get("max_zip_size_mb", 50),
    )


def save_config(cfg: AppConfig) -> None:
    payload = {
        "print": asdict(cfg.print),
        "numbering": asdict(cfg.numbering),
        "vigencia": asdict(cfg.vigencia),
        "dni_regex": cfg.dni_regex,
        "max_excel_size_mb": cfg.max_excel_size_mb,
        "max_word_template_size_mb": cfg.max_word_template_size_mb,
        "max_photo_size_mb": cfg.max_photo_size_mb,
        "max_zip_size_mb": cfg.max_zip_size_mb,
    }
    CONFIG_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
