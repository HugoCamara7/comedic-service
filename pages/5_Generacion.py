from __future__ import annotations

import io
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import assignment, history, pdf_generator, session, validations, word_generator
from modules.config import load_config

st.set_page_config(page_title="Generacion", page_icon="🖨️", layout="wide")
st.title("🖨️ Generacion de carnets")

BASE_DIR = Path(__file__).resolve().parent.parent
HISTORY_FILES_DIR = BASE_DIR / "generated" / "history"

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
correctos, observados = validations.split_records(records)
seleccionados = [r for r in correctos if r.dni in selected_dnis]
seleccionados.sort(key=lambda r: r.fila_excel)

st.write(f"Se generaran carnets para **{len(seleccionados)}** persona(s).")
with st.expander("Configuracion de impresion actual"):
    p = config.print
    st.write(
        f"Papel: {p.paper_size} ({p.orientation}) · Grilla: {p.columnas}x{p.filas} por hoja · "
        f"Margenes: {p.margen_superior_mm}/{p.margen_inferior_mm}/{p.margen_izquierdo_mm}/{p.margen_derecho_mm} mm · "
        f"Separacion: {p.separacion_horizontal_mm}x{p.separacion_vertical_mm} mm"
    )
    st.caption("Puedes ajustar estos valores en la seccion Configuracion.")

confirmar = st.checkbox("Confirmo que revise la vista previa y deseo generar los documentos.")
generar = st.button("🚀 Generar Word y PDF", type="primary", disabled=not confirmar)

if generar:
    with st.spinner("Generando documentos... esto puede tardar segun la cantidad de registros."):
        try:
            used_numbers = history.get_all_used_carnet_numbers()
            people = assignment.build_person_dicts(seleccionados, config, used_numbers)

            out_dir = session.output_dir() / "final"
            out_dir.mkdir(parents=True, exist_ok=True)
            docx_path = out_dir / "carnets_sanidad.docx"

            word_generator.generate_word_document(
                people=people,
                template_path=template_path,
                photos_dir=session.photos_dir(),
                foto_field_present=foto_requerida,
                config=config.print,
                output_path=docx_path,
            )
            pdf_path = pdf_generator.convert_docx_to_pdf(docx_path, out_dir)

            excel_filename = st.session_state.get("excel_filename", "archivo.xlsx")
            gen_record = history.GenerationRecord(
                source_filename=excel_filename,
                total_records=len(records),
                correct_records=len(correctos),
                observed_records=len(observados),
                params={
                    "columnas": config.print.columnas,
                    "filas": config.print.filas,
                    "paper_size": config.print.paper_size,
                },
            )
            generation_id = history.save_generation(gen_record, people)

            # Copia permanente para el historial / reimpresion.
            hist_dir = HISTORY_FILES_DIR / str(generation_id)
            hist_dir.mkdir(parents=True, exist_ok=True)
            final_docx = hist_dir / "carnets_sanidad.docx"
            final_pdf = hist_dir / "carnets_sanidad.pdf"
            shutil.copy(docx_path, final_docx)
            shutil.copy(pdf_path, final_pdf)
            history.update_generation_paths(generation_id, str(final_docx), str(final_pdf))

            st.session_state.last_generation = {
                "id": generation_id,
                "docx": str(final_docx),
                "pdf": str(final_pdf),
            }
            st.success(f"¡Generacion completada! (ID de historial: {generation_id})")
        except pdf_generator.PdfConversionError as exc:
            st.error(str(exc))
        except Exception as exc:  # noqa: BLE001
            st.error(f"Ocurrio un error durante la generacion: {exc}")

last_gen = st.session_state.get("last_generation")
if last_gen:
    st.divider()
    st.subheader("Descargas")
    docx_bytes = Path(last_gen["docx"]).read_bytes()
    pdf_bytes = Path(last_gen["pdf"]).read_bytes()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "⬇️ Descargar Word (editable)",
            data=docx_bytes,
            file_name="carnets_sanidad.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "⬇️ Descargar PDF (para imprimir)",
            data=pdf_bytes,
            file_name="carnets_sanidad.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    with c3:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("carnets_sanidad.docx", docx_bytes)
            zf.writestr("carnets_sanidad.pdf", pdf_bytes)
        st.download_button(
            "⬇️ Descargar todo (ZIP)",
            data=zip_buf.getvalue(),
            file_name="carnets_sanidad.zip",
            mime="application/zip",
            use_container_width=True,
        )

st.divider()
if st.button("⬅️ Volver a Vista previa"):
    st.switch_page("pages/4_Vista_previa.py")
