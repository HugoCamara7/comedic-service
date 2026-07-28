"""
print_layout.py
----------------
Construye la "hoja maestra" (grilla) donde se insertan los carnets
individuales mediante subdocumentos de docxtpl.

La hoja maestra es un .docx generado dinamicamente segun la configuracion de
impresion (tamano de papel, orientacion, filas x columnas, margenes,
separacion y desplazamiento de calibracion). Cada celda de la tabla contiene
un placeholder de texto "{{ carnet_0 }}", "{{ carnet_1 }}", ... que luego
docxtpl reemplaza por el subdocumento renderizado de cada persona.

Usar PDF (obtenido por conversion de este mismo .docx con LibreOffice) como
documento principal de impresion garantiza que las posiciones del PDF
coincidan exactamente con las del Word editable, porque ambos provienen del
mismo archivo fuente.
"""
from __future__ import annotations

from pathlib import Path
from typing import Tuple

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Mm

from modules.config import PrintConfig

PAPER_SIZES_MM = {
    "A4": (210.0, 297.0),
    "A3": (297.0, 420.0),
    "Carta": (215.9, 279.4),
    "Oficio": (215.9, 355.6),
}


def _no_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.makeelement(qn("w:tblBorders"), {})
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.makeelement(qn(f"w:{edge}"), {qn("w:val"): "none"})
        borders.append(el)
    tbl_pr.append(borders)


def _dashed_borders(table) -> None:
    """Bordes punteados en cada celda, utiles como lineas de corte."""
    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            borders = tc_pr.makeelement(qn("w:tcBorders"), {})
            for edge in ("top", "left", "bottom", "right"):
                el = borders.makeelement(
                    qn(f"w:{edge}"),
                    {qn("w:val"): "dashed", qn("w:sz"): "4", qn("w:color"): "999999"},
                )
                borders.append(el)
            tc_pr.append(borders)


def page_size_mm(config: PrintConfig) -> Tuple[float, float]:
    width, height = PAPER_SIZES_MM.get(config.paper_size, PAPER_SIZES_MM["A4"])
    if config.orientation == "horizontal":
        width, height = height, width
    return width, height


def build_master_sheet(config: PrintConfig, output_path: Path, n_carnets: int) -> Path:
    """Genera un .docx de una sola hoja con `n_carnets` celdas dispuestas en
    la grilla filas x columnas definida en config, cada una con el
    placeholder {{ carnet_i }}."""
    doc = Document()
    section = doc.sections[0]

    page_w, page_h = page_size_mm(config)
    section.page_width = Mm(page_w)
    section.page_height = Mm(page_h)
    section.orientation = (
        WD_ORIENT.LANDSCAPE if config.orientation == "horizontal" else WD_ORIENT.PORTRAIT
    )

    # El desplazamiento de calibracion se aplica sumandolo al margen superior
    # e izquierdo (puede ser negativo).
    section.top_margin = Mm(max(0.0, config.margen_superior_mm + config.desplazamiento_vertical_mm))
    section.bottom_margin = Mm(config.margen_inferior_mm)
    section.left_margin = Mm(max(0.0, config.margen_izquierdo_mm + config.desplazamiento_horizontal_mm))
    section.right_margin = Mm(config.margen_derecho_mm)

    cols = config.columnas
    rows = config.filas

    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if config.lineas_de_corte:
        _dashed_borders(table)
    else:
        _no_borders(table)

    cell_w = config.carnet_ancho_mm
    cell_h = config.carnet_alto_mm

    # Ancho fijo de columnas (python-docx requiere fijarlo en cada celda).
    for r in range(rows):
        row_obj = table.rows[r]
        row_obj.height = Mm(cell_h)
        for c in range(cols):
            cell = table.cell(r, c)
            cell.width = Mm(cell_w)
    for c in range(cols):
        table.columns[c].width = Mm(cell_w)

    # Las celdas quedan vacias: modules/word_generator.py inserta el
    # contenido de cada carnet ya renderizado (copia de XML), por lo que no
    # se usan placeholders jinja a este nivel.

    # Espaciado entre celdas: Word no soporta cellSpacing facilmente via
    # python-docx de forma nativa entre columnas; se aproxima insertando un
    # ancho de "separador" como columnas vacias no es practico con celdas de
    # tamano fijo, por lo que se usa el atributo tblCellSpacing.
    tbl_pr = table._tbl.tblPr
    spacing_val = int(Mm(min(config.separacion_horizontal_mm, config.separacion_vertical_mm)).twips / 20)
    cell_spacing = tbl_pr.makeelement(
        qn("w:tblCellSpacing"), {qn("w:w"): str(max(0, spacing_val)), qn("w:type"): "dxa"}
    )
    tbl_pr.append(cell_spacing)

    doc.save(output_path)
    return output_path


def carnets_per_page(config: PrintConfig) -> int:
    return max(1, config.columnas * config.filas)
