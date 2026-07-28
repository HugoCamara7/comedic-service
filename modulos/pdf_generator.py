"""
pdf_generator.py
-----------------
Conversion del .docx final a PDF, usando LibreOffice en modo headless.

Decision de diseno: se evaluaron dos alternativas para producir el PDF de
impresion:

1. ReportLab: requeriria reimplementar todo el layout (tabla de carnets,
   fuentes, imagenes) en un motor separado del que genera el Word, con alto
   riesgo de que las posiciones del PDF no coincidan exactamente con las del
   Word editable.
2. Conversion del propio .docx a PDF con LibreOffice ("soffice --headless
   --convert-to pdf"): garantiza que el PDF sea un reflejo fiel del mismo
   documento que el usuario puede editar en Word, porque se genera a partir
   del mismo archivo fuente.

Se opto por la alternativa 2. Requiere tener instalado LibreOffice en el
servidor (paquete `libreoffice` en Debian/Ubuntu, o `libreoffice-writer` como
opcion minima). No se asume que Microsoft Word este instalado.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


class PdfConversionError(Exception):
    pass


def check_libreoffice_available() -> bool:
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def convert_docx_to_pdf(docx_path: Path, output_dir: Path, timeout: int = 120) -> Path:
    """Convierte un .docx a PDF usando LibreOffice headless.

    Devuelve la ruta del PDF generado. Lanza PdfConversionError si
    LibreOffice no esta disponible o si la conversion falla.
    """
    if not check_libreoffice_available():
        raise PdfConversionError(
            "LibreOffice no esta instalado en el servidor. Instale con: "
            "'sudo apt-get install libreoffice' (Debian/Ubuntu) para habilitar "
            "la generacion de PDF."
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")

    cmd = [
        soffice_bin,
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(docx_path),
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout, check=False
        )
    except subprocess.TimeoutExpired as exc:
        raise PdfConversionError("La conversion a PDF excedio el tiempo limite.") from exc

    expected_pdf = output_dir / (docx_path.stem + ".pdf")
    if result.returncode != 0 or not expected_pdf.exists():
        raise PdfConversionError(
            f"Fallo la conversion a PDF (codigo {result.returncode}). "
            f"Detalle: {result.stderr.strip()[:500]}"
        )
    return expected_pdf
