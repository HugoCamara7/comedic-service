from __future__ import annotations

import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import assignment, history, pdf_generator, session, template_manager, validations, word_generator
from modules.config import load_config
from modules.print_layout import carnets_per_page

st.set_page_config(page_title="Vista previa", page_icon="🖼️", layout="wide")
st.title("🖼️ Vista previa")

records = st.session_state.get("validated_records")
selected_dnis = st.session_state.get("selected_dnis", set())
if not records or not selected_dnis:
    st.warning("No hay personas seleccionadas. Ve a 'Seleccion' primero.")
    st.stop()

template_path_str = st.session_state.get("template_path")
if not template_path_str:
    st.warning("No hay plantilla Word seleccionada. Ve a 'Nueva generacion' primero.")
    st.stop()
template_path = Path(template_path_str)
foto_requerida = st.session_state.get("foto_requerida", False)

config = load_config()
correctos, _ = validations.split_records(records)
seleccionados = [r for r in correctos if r.dni in selected_dnis]
seleccionados.sort(key=lambda r: r.fila_excel)

used_numbers = history.get_all_used_carnet_numbers()
people = assignment.build_person_dicts(seleccionados, config, used_numbers)

n_per_page = carnets_per_page(config.print)
total_pages = max(1, (len(people) + n_per_page - 1) // n_per_page)

st.caption(
    f"{len(people)} personas seleccionadas · {config.print.columnas}x{config.print.filas} "
    f"carnets por hoja · {total_pages} hoja(s) en total."
)

if "preview_page_idx" not in st.session_state:
    st.session_state.preview_page_idx = 0

nav1, nav2, nav3 = st.columns([1, 2, 1])
with nav1:
    if st.button("⬅️ Anterior", disabled=st.session_state.preview_page_idx <= 0):
        st.session_state.preview_page_idx -= 1
with nav3:
    if st.button("Siguiente ➡️", disabled=st.session_state.preview_page_idx >= total_pages - 1):
        st.session_state.preview_page_idx += 1
with nav2:
    st.markdown(
        f"<div style='text-align:center'>Hoja {st.session_state.preview_page_idx + 1} de {total_pages}</div>",
        unsafe_allow_html=True,
    )

page_idx = st.session_state.preview_page_idx
batch = people[page_idx * n_per_page : (page_idx + 1) * n_per_page]

cache_key = f"preview_img_{page_idx}_{len(people)}_{config.print.columnas}x{config.print.filas}"

if cache_key not in st.session_state:
    with st.spinner("Generando vista previa de la hoja..."):
        try:
            work_dir = session.output_dir() / "preview"
            work_dir.mkdir(parents=True, exist_ok=True)
            page_docx = word_generator.compose_page(
                people_batch=batch,
                template_path=template_path,
                photos_dir=session.photos_dir(),
                foto_field_present=foto_requerida,
                config=config.print,
                work_dir=work_dir,
            )
            pdf_path = pdf_generator.convert_docx_to_pdf(page_docx, work_dir)
            img_prefix = work_dir / "preview"
            subprocess.run(
                ["pdftoppm", "-jpeg", "-r", "110", str(pdf_path), str(img_prefix)],
                check=True,
                capture_output=True,
                timeout=60,
            )
            img_candidates = sorted(work_dir.glob("preview*.jpg"))
            if img_candidates:
                st.session_state[cache_key] = img_candidates[0].read_bytes()
            else:
                st.session_state[cache_key] = None
        except pdf_generator.PdfConversionError as exc:
            st.error(str(exc))
            st.session_state[cache_key] = None
        except Exception as exc:  # noqa: BLE001
            st.error(f"No se pudo generar la vista previa: {exc}")
            st.session_state[cache_key] = None

img_bytes = st.session_state.get(cache_key)
if img_bytes:
    st.image(img_bytes, caption=f"Hoja {page_idx + 1} de {total_pages}", use_container_width=True)
else:
    st.warning("No se pudo renderizar la vista previa de esta hoja.")

with st.expander("Personas en esta hoja"):
    st.dataframe(
        [
            {"DNI": p["dni"], "Nombre": p["nombre_completo"], "N Carnet": p["numero_carnet"]}
            for p in batch
        ],
        use_container_width=True,
        hide_index=True,
    )

st.divider()
col_prev, col_next = st.columns(2)
with col_prev:
    if st.button("⬅️ Volver a Seleccion"):
        st.switch_page("pages/3_Seleccion.py")
with col_next:
    if st.button("Continuar a Generacion ➡️", type="primary"):
        st.switch_page("pages/5_Generacion.py")
