"""
session.py
----------
Utilidades de estado de sesion para la app Streamlit.

Cada sesion de navegador obtiene una carpeta de trabajo propia bajo
generated/sessions/<uuid>, usada para fotografias, plantillas subidas y
documentos generados temporalmente. No se comparte entre usuarios.
"""
from __future__ import annotations

import uuid
import zipfile
from pathlib import Path
from typing import List

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent.parent
SESSIONS_DIR = BASE_DIR / "generated" / "sessions"
SESSIONS_DIR.mkdir(parents=True, exist_ok=True)

MAX_PHOTO_ZIP_FILES = 2000


def get_session_dir() -> Path:
    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex[:12]
    session_dir = SESSIONS_DIR / st.session_state.session_id
    (session_dir / "fotos").mkdir(parents=True, exist_ok=True)
    (session_dir / "salida").mkdir(parents=True, exist_ok=True)
    return session_dir


def photos_dir() -> Path:
    d = get_session_dir() / "fotos"
    d.mkdir(parents=True, exist_ok=True)
    return d


def output_dir() -> Path:
    d = get_session_dir() / "salida"
    d.mkdir(parents=True, exist_ok=True)
    return d


def save_uploaded_photo(uploaded_file, max_size_mb: int = 5) -> Path:
    data = uploaded_file.getvalue()
    size_mb = len(data) / (1024 * 1024)
    if size_mb > max_size_mb:
        raise ValueError(
            f"La foto '{uploaded_file.name}' pesa {size_mb:.1f} MB y excede el limite de {max_size_mb} MB."
        )
    safe_name = Path(uploaded_file.name).name
    dest = photos_dir() / safe_name
    dest.write_bytes(data)
    return dest


def extract_photos_zip(uploaded_zip, max_size_mb: int = 50) -> List[str]:
    """Extrae fotos de un ZIP en la carpeta de fotos de la sesion.

    Medidas de seguridad basicas: bloquea rutas absolutas o con '..'
    (zip-slip), limita cantidad de archivos y tamano total.
    """
    data = uploaded_zip.getvalue()
    size_mb = len(data) / (1024 * 1024)
    if size_mb > max_size_mb:
        raise ValueError(f"El ZIP pesa {size_mb:.1f} MB y excede el limite de {max_size_mb} MB.")

    import io

    dest_dir = photos_dir()
    extracted = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = zf.namelist()
        if len(names) > MAX_PHOTO_ZIP_FILES:
            raise ValueError(f"El ZIP contiene demasiados archivos ({len(names)}).")
        for name in names:
            norm = Path(name)
            if norm.is_absolute() or ".." in norm.parts:
                continue  # zip-slip guard
            if name.endswith("/"):
                continue
            ext = norm.suffix.lower()
            if ext not in {".jpg", ".jpeg", ".png"}:
                continue
            safe_name = norm.name
            target = dest_dir / safe_name
            with zf.open(name) as src, open(target, "wb") as out:
                out.write(src.read())
            extracted.append(safe_name)
    return extracted


def list_available_photos() -> List[str]:
    return sorted(p.name for p in photos_dir().glob("*") if p.is_file())


def reset_session_files() -> None:
    """Limpia fotos y salidas de la sesion actual (no borra el historial)."""
    import shutil

    session_dir = get_session_dir()
    for sub in ("fotos", "salida"):
        d = session_dir / sub
        if d.exists():
            shutil.rmtree(d)
        d.mkdir(parents=True, exist_ok=True)
