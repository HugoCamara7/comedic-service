from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import history

st.set_page_config(page_title="Historial", page_icon="🗄️", layout="wide")
st.title("🗄️ Historial de generaciones")

rows = history.list_generations(limit=200)
if not rows:
    st.info("Todavia no hay generaciones registradas.")
    st.stop()

st.dataframe(
    [
        {
            "ID": r["id"],
            "Fecha": r["timestamp"],
            "Archivo origen": r["source_filename"],
            "Total": r["total_records"],
            "Correctos": r["correct_records"],
            "Observados": r["observed_records"],
        }
        for r in rows
    ],
    use_container_width=True,
    hide_index=True,
)

st.divider()
st.subheader("Reimprimir una generacion anterior")

ids = [r["id"] for r in rows]
selected_id = st.selectbox("Selecciona el ID de la generacion", ids)

selected_row = next((r for r in rows if r["id"] == selected_id), None)
if selected_row:
    st.write(f"**Archivo origen:** {selected_row['source_filename']}")
    st.write(f"**Fecha:** {selected_row['timestamp']}")
    st.write(
        f"**Registros:** {selected_row['total_records']} total, "
        f"{selected_row['correct_records']} correctos, {selected_row['observed_records']} observados"
    )

    people = history.get_generation_people(selected_id)
    with st.expander(f"Personas incluidas ({len(people)})"):
        st.dataframe(
            [
                {
                    "DNI": p["dni"],
                    "Nombre": p["nombre_completo"],
                    "N Carnet": p["numero_carnet"],
                    "Emision": p["fecha_emision"],
                    "Vencimiento": p["fecha_vencimiento"],
                }
                for p in people
            ],
            use_container_width=True,
            hide_index=True,
        )

    docx_path = Path(selected_row["docx_path"]) if selected_row["docx_path"] else None
    pdf_path = Path(selected_row["pdf_path"]) if selected_row["pdf_path"] else None

    c1, c2 = st.columns(2)
    with c1:
        if docx_path and docx_path.exists():
            st.download_button(
                "⬇️ Descargar Word",
                data=docx_path.read_bytes(),
                file_name=docx_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.caption("Archivo Word no disponible (pudo haberse limpiado del servidor).")
    with c2:
        if pdf_path and pdf_path.exists():
            st.download_button(
                "⬇️ Descargar PDF",
                data=pdf_path.read_bytes(),
                file_name=pdf_path.name,
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.caption("Archivo PDF no disponible (pudo haberse limpiado del servidor).")
