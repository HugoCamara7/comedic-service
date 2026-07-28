from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st

from modules.config import load_config, save_config
from modules.print_layout import PAPER_SIZES_MM

st.set_page_config(page_title="Configuracion", page_icon="⚙️", layout="wide")
st.title("⚙️ Configuracion")

config = load_config()

st.subheader("Distribucion e impresion")
with st.form("form_impresion"):
    c1, c2 = st.columns(2)
    with c1:
        paper_size = st.selectbox(
            "Tamano de papel", list(PAPER_SIZES_MM.keys()),
            index=list(PAPER_SIZES_MM.keys()).index(config.print.paper_size),
        )
        orientation = st.selectbox(
            "Orientacion", ["vertical", "horizontal"],
            index=["vertical", "horizontal"].index(config.print.orientation),
        )
        columnas = st.number_input("Carnets por fila (columnas)", min_value=1, max_value=10, value=config.print.columnas)
        filas = st.number_input("Filas de carnets por hoja", min_value=1, max_value=10, value=config.print.filas)
        lineas_de_corte = st.checkbox("Mostrar lineas de corte", value=config.print.lineas_de_corte)
    with c2:
        carnet_ancho = st.number_input("Ancho de carnet (mm)", min_value=10.0, value=float(config.print.carnet_ancho_mm))
        carnet_alto = st.number_input("Alto de carnet (mm)", min_value=10.0, value=float(config.print.carnet_alto_mm))
        escala = st.number_input("Escala de impresion (%)", min_value=50, max_value=100, value=config.print.escala_impresion)
        st.caption("La escala de impresion debe mantenerse en 100% al imprimir desde el visor de PDF.")

    st.markdown("**Margenes de hoja (mm)**")
    m1, m2, m3, m4 = st.columns(4)
    margen_sup = m1.number_input("Superior", min_value=0.0, value=float(config.print.margen_superior_mm))
    margen_inf = m2.number_input("Inferior", min_value=0.0, value=float(config.print.margen_inferior_mm))
    margen_izq = m3.number_input("Izquierdo", min_value=0.0, value=float(config.print.margen_izquierdo_mm))
    margen_der = m4.number_input("Derecho", min_value=0.0, value=float(config.print.margen_derecho_mm))

    st.markdown("**Separacion entre carnets (mm)**")
    s1, s2 = st.columns(2)
    sep_h = s1.number_input("Horizontal", min_value=0.0, value=float(config.print.separacion_horizontal_mm))
    sep_v = s2.number_input("Vertical", min_value=0.0, value=float(config.print.separacion_vertical_mm))

    st.markdown("**Calibracion de impresora (desplazamiento en mm)**")
    d1, d2 = st.columns(2)
    despl_h = d1.number_input("Desplazamiento horizontal", value=float(config.print.desplazamiento_horizontal_mm))
    despl_v = d2.number_input("Desplazamiento vertical", value=float(config.print.desplazamiento_vertical_mm))

    submitted_print = st.form_submit_button("💾 Guardar configuracion de impresion")

if submitted_print:
    config.print.paper_size = paper_size
    config.print.orientation = orientation
    config.print.columnas = int(columnas)
    config.print.filas = int(filas)
    config.print.lineas_de_corte = lineas_de_corte
    config.print.carnet_ancho_mm = float(carnet_ancho)
    config.print.carnet_alto_mm = float(carnet_alto)
    config.print.escala_impresion = int(escala)
    config.print.margen_superior_mm = float(margen_sup)
    config.print.margen_inferior_mm = float(margen_inf)
    config.print.margen_izquierdo_mm = float(margen_izq)
    config.print.margen_derecho_mm = float(margen_der)
    config.print.separacion_horizontal_mm = float(sep_h)
    config.print.separacion_vertical_mm = float(sep_v)
    config.print.desplazamiento_horizontal_mm = float(despl_h)
    config.print.desplazamiento_vertical_mm = float(despl_v)
    save_config(config)
    st.success("Configuracion de impresion guardada.")

st.divider()
st.subheader("Numeracion de carnets")
with st.form("form_numeracion"):
    autogenerar = st.checkbox("Autogenerar numero de carnet cuando venga vacio en el Excel", value=config.numbering.autogenerar)
    n1, n2, n3 = st.columns(3)
    prefijo = n1.text_input("Prefijo", value=config.numbering.prefijo)
    digitos = n2.number_input("Digitos del correlativo", min_value=1, max_value=10, value=config.numbering.digitos)
    siguiente = n3.number_input("Proximo correlativo", min_value=1, value=config.numbering.siguiente_correlativo)
    st.caption(f"Ejemplo de numero generado: {prefijo}{siguiente:0{digitos}d}")
    submitted_num = st.form_submit_button("💾 Guardar configuracion de numeracion")

if submitted_num:
    config.numbering.autogenerar = autogenerar
    config.numbering.prefijo = prefijo
    config.numbering.digitos = int(digitos)
    config.numbering.siguiente_correlativo = int(siguiente)
    save_config(config)
    st.success("Configuracion de numeracion guardada.")

st.divider()
st.subheader("Vigencia del carnet")
with st.form("form_vigencia"):
    autocalcular = st.checkbox(
        "Autocalcular fecha de vencimiento cuando venga vacia en el Excel", value=config.vigencia.autocalcular
    )
    meses = st.number_input("Meses de vigencia", min_value=1, max_value=120, value=config.vigencia.meses_vigencia)
    submitted_vig = st.form_submit_button("💾 Guardar configuracion de vigencia")

if submitted_vig:
    config.vigencia.autocalcular = autocalcular
    config.vigencia.meses_vigencia = int(meses)
    save_config(config)
    st.success("Configuracion de vigencia guardada.")

st.divider()
st.subheader("Validacion")
with st.form("form_validacion"):
    dni_regex = st.text_input("Expresion regular para validar el DNI", value=config.dni_regex)
    st.caption("Por defecto: 8 digitos numericos (DNI peruano). Ajustala si tu pais usa otro formato.")
    submitted_val = st.form_submit_button("💾 Guardar configuracion de validacion")

if submitted_val:
    config.dni_regex = dni_regex
    save_config(config)
    st.success("Configuracion de validacion guardada.")
