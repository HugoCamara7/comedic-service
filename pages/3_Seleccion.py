from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules import validations

st.set_page_config(page_title="Seleccion", page_icon="🗂️", layout="wide")
st.title("🗂️ Seleccion de personas")

records = st.session_state.get("validated_records")
if not records:
    st.warning("Todavia no has cargado ningun Excel. Ve a 'Nueva generacion' primero.")
    st.stop()

correctos, _observados = validations.split_records(records)
if not correctos:
    st.warning("No hay registros correctos para seleccionar.")
    st.stop()

if "selected_dnis" not in st.session_state:
    st.session_state.selected_dnis = {r.dni for r in correctos}

st.caption(f"{len(correctos)} registros correctos disponibles para generar carnets.")

col1, col2, col3 = st.columns(3)
with col1:
    busqueda = st.text_input("🔎 Buscar por DNI o nombre")
with col2:
    empresas = sorted({r.empresa for r in correctos if r.empresa})
    filtro_empresa = st.selectbox("Filtrar por empresa", ["(Todas)"] + empresas)
with col3:
    puestos = sorted({r.puesto for r in correctos if r.puesto})
    filtro_puesto = st.selectbox("Filtrar por puesto", ["(Todos)"] + puestos)

filtered = correctos
if busqueda:
    q = busqueda.strip().lower()
    filtered = [r for r in filtered if q in r.dni.lower() or q in r.nombre_completo.lower()]
if filtro_empresa != "(Todas)":
    filtered = [r for r in filtered if r.empresa == filtro_empresa]
if filtro_puesto != "(Todos)":
    filtered = [r for r in filtered if r.puesto == filtro_puesto]

btn_col1, btn_col2, _ = st.columns([1, 1, 3])
with btn_col1:
    if st.button("Seleccionar todos (filtrados)"):
        st.session_state.selected_dnis |= {r.dni for r in filtered}
with btn_col2:
    if st.button("Quitar todos (filtrados)"):
        st.session_state.selected_dnis -= {r.dni for r in filtered}

st.divider()

for r in filtered:
    checked = r.dni in st.session_state.selected_dnis
    new_val = st.checkbox(
        f"**{r.nombre_completo}** — DNI {r.dni} — {r.empresa or 'Sin empresa'} / {r.puesto or 'Sin puesto'}",
        value=checked,
        key=f"chk_{r.dni}",
    )
    if new_val:
        st.session_state.selected_dnis.add(r.dni)
    else:
        st.session_state.selected_dnis.discard(r.dni)

st.divider()
n_seleccionados = len(st.session_state.selected_dnis)
st.metric("Personas seleccionadas", n_seleccionados)

col_prev, col_next = st.columns(2)
with col_prev:
    if st.button("⬅️ Volver a Validacion"):
        st.switch_page("pages/2_Validacion.py")
with col_next:
    if st.button("Continuar a Vista previa ➡️", type="primary", disabled=n_seleccionados == 0):
        st.switch_page("pages/4_Vista_previa.py")
