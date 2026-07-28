"""
app.py
------
Punto de entrada de la aplicacion Streamlit "Carnets de Sanidad".
Pantalla: Inicio.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import streamlit as st

from modules import history

st.set_page_config(
    page_title="Carnets de Sanidad",
    page_icon="🪪",
    layout="wide",
)

history.init_db()

st.title("🪪 Generador de Carnets de Sanidad")
st.caption(
    "Automatiza la carga de datos desde Excel, la validacion de registros y la "
    "generacion masiva de carnets en Word y PDF listos para imprimir."
)

st.divider()

col1, col2, col3 = st.columns(3)
summary = history.summary_counts()
col1.metric("Generaciones registradas", summary["total_generaciones"])
col2.metric("Personas procesadas (historico)", summary["total_personas"])
col3.metric("Estado del sistema", "Listo ✅")

st.divider()

st.subheader("Como funciona")
st.markdown(
    """
1. **Nueva generacion** — descarga el Excel estandar, cargalo con tus datos, sube la plantilla Word del carnet y (opcional) las fotografias.
2. **Validacion** — revisa los registros correctos y observados, con el motivo exacto de cada error.
3. **Seleccion** — elige que personas procesar.
4. **Vista previa** — revisa como quedaran los carnets antes de descargar.
5. **Generacion** — obtiene el Word editable y el PDF listo para imprimir.
6. **Historial** — consulta generaciones anteriores y reimprime si lo necesitas.
7. **Configuracion** — ajusta margenes, distribucion por hoja, numeracion y vigencia.

Usa el menu de la izquierda para navegar entre las secciones.
"""
)

st.divider()
st.subheader("Ultimas generaciones")
rows = history.list_generations(limit=5)
if not rows:
    st.info("Todavia no se ha generado ningun carnet. Ve a 'Nueva generacion' para empezar.")
else:
    table_rows = [
        {
            "Fecha": r["timestamp"],
            "Archivo": r["source_filename"],
            "Total": r["total_records"],
            "Correctos": r["correct_records"],
            "Observados": r["observed_records"],
        }
        for r in rows
    ]
    st.dataframe(table_rows, use_container_width=True, hide_index=True)
