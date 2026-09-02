import pandas as pd
import streamlit as st

st.set_page_config(page_title="Consola Única de Compras - Enaex", layout="wide")

if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

st.title("🛒 Consola Única de Compras — Enaex")

# BARRA LATERAL: DATOS ERP
with st.sidebar:
    st.header("📌 Datos Solped")
    solped = st.text_input("N° Solped", value="10045982")
    material = st.text_input("Código Material", value="3001892")
    sociedad = st.selectbox("Sociedad", ["EC01", "EC06"])

# INGRESO MANUAL DE COTIZACIONES
st.subheader("➕ Carga Manual de Oferta")
with st.form("form_cotizacion", clear_on_submit=True):
    col1, col2, col3, col4 = st.columns(4)
    proveedor = col1.text_input("Proveedor")
    monto = col2.number_input("Monto Total", min_value=0.0, step=1000.0)
    moneda = col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"])
    plazo = col4.number_input("Plazo Entrega (Días)", min_value=1, value=5)
    obs = st.text_area("Observaciones Técnicas")

    if st.form_submit_button("Guardar en Cuadro Comparativo"):
        if proveedor and monto > 0:
            st.session_state["cotizaciones"].append({
                "Proveedor": proveedor,
                "Monto": monto,
                "Moneda": moneda,
                "Plazo (Días)": plazo,
                "Observaciones": obs,
            })
            st.success(f"Oferta de {proveedor} guardada.")
            st.rerun()

# CUADRO COMPARATIVO
st.divider()
st.subheader("📊 Cuadro Comparativo")

if st.session_state["cotizaciones"]:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    st.dataframe(df, use_container_width=True)

    if st.button("🚀 Emitir a SAP ME21N"):
        st.balloons()
        st.success(
            f"Orden emitida para Solped {solped} bajo Sociedad {sociedad}."
        )
else:
    st.info("Ingresa la primera cotización arriba para visualizar la tabla.")
