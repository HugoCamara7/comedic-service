"""
assignment.py
-------------
Resuelve numero_carnet autogenerado (si esta habilitado en Configuracion) y
convierte ValidatedRecord -> dict plano listo para word_generator / history.
"""
from __future__ import annotations

from typing import Dict, List, Set

from modules.config import AppConfig
from modules.validations import ValidatedRecord


def _format_date(dt) -> str:
    if dt is None:
        return ""
    return dt.strftime("%d/%m/%Y")


def build_person_dicts(
    records: List[ValidatedRecord],
    config: AppConfig,
    used_carnet_numbers: Set[str],
) -> List[Dict]:
    """Devuelve una lista de dicts (uno por persona) con numero_carnet resuelto.

    No muta los ValidatedRecord originales.
    """
    people = []
    next_correlativo = config.numbering.siguiente_correlativo
    used = set(used_carnet_numbers)

    for r in records:
        numero_carnet = r.numero_carnet
        if not numero_carnet and config.numbering.autogenerar:
            while True:
                candidate = f"{config.numbering.prefijo}{next_correlativo:0{config.numbering.digitos}d}"
                next_correlativo += 1
                if candidate not in used:
                    numero_carnet = candidate
                    used.add(candidate)
                    break

        people.append(
            {
                "dni": r.dni,
                "nombre_completo": r.nombre_completo,
                "numero_carnet": numero_carnet,
                "fecha_nacimiento": _format_date(r.fecha_nacimiento),
                "fecha_emision": _format_date(r.fecha_emision),
                "fecha_vencimiento": _format_date(r.fecha_vencimiento),
                "empresa": r.empresa,
                "puesto": r.puesto,
                "fotografia": r.fotografia,
            }
        )
    return people
