from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import excel_reader, history, session, template_manager, validations
from modules.config import load_config

st.set_page_config(page_title="Nueva generacion", page_icon="📄", layout="wide")
st.title("📄 Nueva generacion")

config = load_config()

# ---------------------------------------------------------------------
# 1. Plantilla Excel estandar
# ---------------------------------------------------------------------
st.subheader("1. Excel estandar")
c1, c2 = st.columns([1, 2])
with c1:
    st.download_button(
        "⬇️ Descargar Excel estandar",
        data=excel_reader.template_bytes(),
        file_name="plantilla_carnets_sanidad.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    st.caption(
        "Contiene los encabezados exactos requeridos y una hoja de instrucciones. "
        "Completa tus datos sin modificar los nombres de columnas."
    )

st.divider()

# ---------------------------------------------------------------------
# 2. Plantilla Word del carnet
# ---------------------------------------------------------------------
st.subheader("2. Plantilla Word del carnet")

existing_templates = template_manager.list_templates()
default_tpl = template_manager.ensure_default_template()
options = ["Usar plantilla por defecto (incluida)"] + [t.name for t in existing_templates if t.name != default_tpl.name]

tpl_choice = st.selectbox("Selecciona una plantilla existente", options)
uploaded_tpl = st.file_uploader("O sube una nueva plantilla Word (.docx)", type=["docx"])

template_path = None
if uploaded_tpl is not None:
    try:
        template_path = template_manager.save_uploaded_template(uploaded_tpl, uploaded_tpl.name)
        st.success(f"Plantilla '{uploaded_tpl.name}' guardada correctamente.")
    except template_manager.TemplateError as exc:
        st.error(str(exc))
elif tpl_choice == options[0]:
    template_path = default_tpl
else:
    template_path = template_manager.TEMPLATES_DIR / tpl_choice

if template_path is not None:
    fields = template_manager.detect_fields(template_path)
    foto_requerida = "foto" in fields
    st.caption(
        f"Campos detectados en la plantilla: {', '.join(sorted(fields)) if fields else '(ninguno)'}"
    )
    if foto_requerida:
        st.info("Esta plantilla incluye el campo {{ foto }}: la fotografia sera obligatoria para cada persona.")
    st.session_state.template_path = str(template_path)
    st.session_state.foto_requerida = foto_requerida

st.divider()

# ---------------------------------------------------------------------
# 3. Fotografias (opcional, salvo que la plantilla las requiera)
# ---------------------------------------------------------------------
st.subheader("3. Fotografias (opcional)")
st.caption(
    "El nombre del archivo debe coincidir exactamente con el valor de la columna "
    "'Fotografia' en el Excel. Puedes subirlas una por una o en un ZIP."
)

col_a, col_b = st.columns(2)
with col_a:
    photos = st.file_uploader(
        "Subir fotografias individuales", type=["jpg", "jpeg", "png"], accept_multiple_files=True
    )
    if photos:
        for photo in photos:
            try:
                session.save_uploaded_photo(photo, max_size_mb=config.max_photo_size_mb)
            except ValueError as exc:
                st.error(str(exc))
        st.success(f"{len(photos)} fotografia(s) cargada(s).")

with col_b:
    photos_zip = st.file_uploader("Subir ZIP con fotografias", type=["zip"])
    if photos_zip is not None:
        try:
            extracted = session.extract_photos_zip(photos_zip, max_size_mb=config.max_zip_size_mb)
            st.success(f"{len(extracted)} fotografia(s) extraida(s) del ZIP.")
        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo procesar el ZIP: {exc}")

available_photos = session.list_available_photos()
if available_photos:
    with st.expander(f"Fotografias disponibles en esta sesion ({len(available_photos)})"):
        st.write(", ".join(available_photos))

st.divider()

# ---------------------------------------------------------------------
# 4. Cargar el Excel completado
# ---------------------------------------------------------------------
st.subheader("4. Cargar el Excel completado")
excel_file = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"], key="excel_uploader")

analizar = st.button("🔍 Analizar y validar registros", type="primary", disabled=excel_file is None)

if analizar and excel_file is not None:
    with st.spinner("Leyendo y validando el Excel..."):
        try:
            result = excel_reader.load_excel(excel_file, excel_file.name, max_size_mb=config.max_excel_size_mb)
        except excel_reader.ExcelFormatError as exc:
            st.error(str(exc))
            st.stop()

        if result.missing_columns:
            st.error(
                "Faltan columnas obligatorias en el Excel: " + ", ".join(result.missing_columns)
            )
            st.stop()

        foto_requerida = st.session_state.get("foto_requerida", False)
        fotos_disponibles = set(session.list_available_photos())
        dnis_previos = history.get_all_processed_dnis()

        records = validations.validate_dataframe(
            result.dataframe,
            config,
            plantilla_requiere_foto=foto_requerida,
            fotos_disponibles=fotos_disponibles,
            dnis_previamente_procesados=dnis_previos,
        )

        st.session_state.validated_records = records
        st.session_state.excel_filename = excel_file.name
        st.session_state.selected_dnis = {r.dni for r in records if r.es_valido}

        correctos, observados = validations.split_records(records)
        st.success(
            f"Analisis completo: {len(records)} registros ({len(correctos)} correctos, "
            f"{len(observados)} observados)."
        )
        if result.extra_columns:
            st.warning(f"Columnas no reconocidas (se ignoraron): {', '.join(result.extra_columns)}")

        st.switch_page("pages/2_Validacion.py")
