from __future__ import annotations

import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import validations

st.set_page_config(page_title="Validacion", page_icon="✅", layout="wide")
st.title("✅ Validacion de registros")

records = st.session_state.get("validated_records")
if not records:
    st.warning("Todavia no has cargado ningun Excel. Ve a 'Nueva generacion' primero.")
    st.stop()

correctos, observados = validations.split_records(records)

c1, c2, c3 = st.columns(3)
c1.metric("Total de registros", len(records))
c2.metric("Correctos", len(correctos))
c3.metric("Observados", len(observados))

st.divider()

st.subheader(f"✅ Registros correctos ({len(correctos)})")
if correctos:
    st.dataframe(
        [
            {
                "Fila": r.fila_excel,
                "DNI": r.dni,
                "Nombre completo": r.nombre_completo,
                "N Carnet": r.numero_carnet or "(se autogenerara)",
                "Empresa": r.empresa,
                "Puesto": r.puesto,
            }
            for r in correctos
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("No hay registros correctos todavia.")

st.divider()

st.subheader(f"⚠️ Registros observados ({len(observados)})")
if observados:
    st.dataframe(
        [
            {
                "Fila": r.fila_excel,
                "DNI": r.dni,
                "Nombre completo": r.nombre_completo,
                "Motivo(s)": " | ".join(r.errores),
            }
            for r in observados
        ],
        use_container_width=True,
        hide_index=True,
    )

    df_obs = validations.observados_dataframe(observados)
    buf = io.BytesIO()
    df_obs.to_excel(buf, index=False, sheet_name="Observados")
    st.download_button(
        "⬇️ Descargar reporte de registros observados (Excel)",
        data=buf.getvalue(),
        file_name="registros_observados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.success("No se detectaron registros observados. 🎉")

st.divider()

col_prev, col_next = st.columns(2)
with col_prev:
    if st.button("⬅️ Volver a Nueva generacion"):
        st.switch_page("pages/1_Nueva_generacion.py")
with col_next:
    if st.button("Continuar a Seleccion ➡️", type="primary", disabled=len(correctos) == 0):
        st.switch_page("pages/3_Seleccion.py")
