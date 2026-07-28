"""
validations.py
---------------
Reglas de validacion de registros de personas antes de generar carnets.

Un registro "observado" (con errores) NO detiene el procesamiento de los
demas. Cada observacion queda asociada a la fila del Excel y a un motivo
legible para el usuario.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd

from modules.config import AppConfig
from modules.excel_reader import REQUIRED_COLUMNS

DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]


def parse_date(value: str) -> Optional[datetime]:
    value = (value or "").strip()
    if not value:
        return None
    # pandas a veces entrega fechas ya como Timestamp-string tipo "2026-01-15 00:00:00"
    for fmt in DATE_FORMATS + ["%Y-%m-%d %H:%M:%S"]:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    # Ultimo intento: dejar que pandas interprete formatos comunes.
    try:
        parsed = pd.to_datetime(value, dayfirst=False, errors="raise")
        return parsed.to_pydatetime()
    except Exception:  # noqa: BLE001
        return None


@dataclass
class ValidatedRecord:
    fila_excel: int
    dni: str
    nombres: str
    apellido_paterno: str
    apellido_materno: str
    fecha_nacimiento: Optional[datetime]
    numero_carnet: str
    fecha_emision: Optional[datetime]
    fecha_vencimiento: Optional[datetime]
    empresa: str
    puesto: str
    fotografia: str
    nombre_completo: str = ""
    errores: List[str] = field(default_factory=list)

    @property
    def es_valido(self) -> bool:
        return len(self.errores) == 0


def _clean(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def validate_dataframe(
    df: pd.DataFrame,
    config: AppConfig,
    plantilla_requiere_foto: bool,
    fotos_disponibles: Optional[set] = None,
    dnis_previamente_procesados: Optional[set] = None,
) -> List[ValidatedRecord]:
    """Valida cada fila del DataFrame y devuelve una lista de ValidatedRecord.

    No lanza excepcion por errores de contenido: cada fila conserva sus
    propios `errores`. Solo se asume que el DataFrame ya paso por
    `excel_reader.load_excel` (columnas completas, aunque puedan faltar
    obligatorias que se validan aqui tambien fila por fila).
    """
    fotos_disponibles = fotos_disponibles or set()
    dnis_previamente_procesados = dnis_previamente_procesados or set()
    dni_regex = re.compile(config.dni_regex)

    records: List[ValidatedRecord] = []
    dnis_vistos: Dict[str, int] = {}
    carnets_vistos: Dict[str, int] = {}

    for _, row in df.iterrows():
        fila = int(row["_fila_excel"])
        dni = _clean(row.get("DNI"))
        nombres = _clean(row.get("Nombres"))
        ap_paterno = _clean(row.get("Apellido paterno"))
        ap_materno = _clean(row.get("Apellido materno"))
        fecha_nac_raw = _clean(row.get("Fecha de nacimiento"))
        numero_carnet = _clean(row.get("Numero de carnet"))
        fecha_emi_raw = _clean(row.get("Fecha de emision"))
        fecha_ven_raw = _clean(row.get("Fecha de vencimiento"))
        empresa = _clean(row.get("Empresa"))
        puesto = _clean(row.get("Puesto"))
        foto = _clean(row.get("Fotografia"))

        errores: List[str] = []

        # --- Obligatorios ---
        if not dni:
            errores.append("DNI vacio.")
        elif not dni_regex.match(dni):
            errores.append(f"DNI con formato incorrecto (se esperaba: {config.dni_regex}).")

        if not nombres:
            errores.append("Nombres vacio.")
        if not ap_paterno:
            errores.append("Apellido paterno vacio.")
        if not ap_materno:
            errores.append("Apellido materno vacio.")

        fecha_nacimiento = parse_date(fecha_nac_raw)
        if not fecha_nac_raw:
            errores.append("Fecha de nacimiento vacia.")
        elif fecha_nacimiento is None:
            errores.append(f"Fecha de nacimiento invalida: '{fecha_nac_raw}'.")

        fecha_emision = parse_date(fecha_emi_raw)
        if not fecha_emi_raw:
            errores.append("Fecha de emision vacia.")
        elif fecha_emision is None:
            errores.append(f"Fecha de emision invalida: '{fecha_emi_raw}'.")

        # --- Fecha de vencimiento: obligatoria salvo que se autocalcule ---
        fecha_vencimiento = parse_date(fecha_ven_raw)
        if not fecha_ven_raw:
            if config.vigencia.autocalcular and fecha_emision is not None:
                meses = config.vigencia.meses_vigencia
                mes = fecha_emision.month - 1 + meses
                anio = fecha_emision.year + mes // 12
                mes = mes % 12 + 1
                dia = min(fecha_emision.day, 28)
                fecha_vencimiento = fecha_emision.replace(year=anio, month=mes, day=dia)
            elif not config.vigencia.autocalcular:
                errores.append(
                    "Fecha de vencimiento vacia (la autogeneracion esta desactivada en Configuracion)."
                )
        elif fecha_vencimiento is None:
            errores.append(f"Fecha de vencimiento invalida: '{fecha_ven_raw}'.")

        if fecha_emision and fecha_vencimiento and fecha_vencimiento < fecha_emision:
            errores.append("Fecha de vencimiento anterior a la fecha de emision.")

        # --- Numero de carnet: obligatorio salvo autogeneracion ---
        if not numero_carnet:
            if not config.numbering.autogenerar:
                errores.append(
                    "Numero de carnet vacio (la autogeneracion esta desactivada en Configuracion)."
                )
            # si autogenerar=True, se asigna mas adelante (word_generator/app)
        else:
            if numero_carnet in carnets_vistos:
                errores.append(
                    f"Numero de carnet duplicado dentro del archivo (fila {carnets_vistos[numero_carnet]})."
                )
            else:
                carnets_vistos[numero_carnet] = fila

        # --- DNI duplicado dentro del archivo ---
        if dni:
            if dni in dnis_vistos:
                errores.append(f"DNI duplicado dentro del archivo (fila {dnis_vistos[dni]}).")
            else:
                dnis_vistos[dni] = fila

            if dni in dnis_previamente_procesados:
                errores.append("DNI previamente procesado (ya existe en el historial).")

        # --- Fotografia ---
        if plantilla_requiere_foto:
            if not foto:
                errores.append("Fotografia faltante (la plantilla seleccionada la requiere).")
            elif fotos_disponibles and foto not in fotos_disponibles:
                errores.append(f"Fotografia '{foto}' no encontrada entre los archivos cargados.")

        nombre_completo = " ".join(p for p in [nombres, ap_paterno, ap_materno] if p).strip()
        if len(nombre_completo) > 60:
            errores.append("Nombre completo muy largo (mas de 60 caracteres); puede no ajustarse a la plantilla.")

        records.append(
            ValidatedRecord(
                fila_excel=fila,
                dni=dni,
                nombres=nombres,
                apellido_paterno=ap_paterno,
                apellido_materno=ap_materno,
                fecha_nacimiento=fecha_nacimiento,
                numero_carnet=numero_carnet,
                fecha_emision=fecha_emision,
                fecha_vencimiento=fecha_vencimiento,
                empresa=empresa,
                puesto=puesto,
                fotografia=foto,
                nombre_completo=nombre_completo,
                errores=errores,
            )
        )

    return records


def split_records(records: List[ValidatedRecord]):
    correctos = [r for r in records if r.es_valido]
    observados = [r for r in records if not r.es_valido]
    return correctos, observados


def observados_dataframe(observados: List[ValidatedRecord]) -> pd.DataFrame:
    rows = []
    for r in observados:
        rows.append(
            {
                "Fila Excel": r.fila_excel,
                "DNI": r.dni,
                "Nombres": r.nombres,
                "Apellido paterno": r.apellido_paterno,
                "Apellido materno": r.apellido_materno,
                "Motivo(s)": " | ".join(r.errores),
            }
        )
    return pd.DataFrame(rows)
