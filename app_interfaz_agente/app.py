import datetime
import io
import pandas as pd
import requests
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Consola de Compras - Enaex", layout="wide")

# 1. BASE HISTÓRICA COMPLETA (INCLUYE 2017, 2018 Y ANTERIORES)
HISTORICO_COMPRAS = pd.DataFrame([
    {"Material_ID": "EXPL-01", "Material": "EXPLOSIVOS HIGH POWER", "Fecha_OC": "2017-05-14", "Precio_CLP": 1250, "Proveedor": "EXPLOSIVOS CHILE"},
    {"Material_ID": "EXPL-01", "Material": "EXPLOSIVOS HIGH POWER", "Fecha_OC": "2018-11-20", "Precio_CLP": 1380, "Proveedor": "EXPLOSIVOS CHILE"},
    {"Material_ID": "QUIM-70", "Material": "MATERIAS PRIMAS QUIMICAS", "Fecha_OC": "2018-03-10", "Precio_CLP": 11000, "Proveedor": "CHEMICAL CORP"},
    {"Material_ID": "TOR-200", "Material": "TORNILLOS 200 KILOS", "Fecha_OC": "2017-09-01", "Precio_CLP": 42000, "Proveedor": "PERNOS S.A."}
])

st.title("🛒 Consola de Compras - Flujo Completo Felipe")

# MÓDULO HISTÓRICO SIN FILTRO DE AÑOS
st.subheader("📜 1. Consulta de Precios Históricos (Sin límite 2019)")
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    mat_buscar = st.text_input("Buscar precio histórico de material (Ej: EXPLOSIVOS, 2017, 2018):", value="EXPLOSIVOS")

df_hist_filtrado = HISTORICO_COMPRAS[HISTORICO_COMPRAS["Material"].str.contains(mat_buscar, case=False, na=False)]
st.dataframe(df_hist_filtrado, use_container_width=True)

st.divider()

# MÓDULO COMPARATIVO PARA 20+ MATERIALES
st.subheader("✏️ 2. Comparativo Dinámico por SOLPED (Soporta 20+ ítems)")

# Ejemplo con SOLPED 287 (3 materiales de prueba)
if "tabla_solped" not in st.session_state:
    st.session_state["tabla_solped"] = pd.DataFrame([
        {"Pos": 1, "Material": "EXPLOSIVOS HIGH POWER", "Cantidad": 20, "Precio Unitario": 1400, "Moneda": "CLP", "Proveedor": "EXPLOSIVOS CHILE", "Días Entrega": 5},
        {"Pos": 2, "Material": "MATERIAS PRIMAS QUIMICAS", "Cantidad": 70, "Precio Unitario": 12.5, "Moneda": "USD", "Proveedor": "CHEMICAL CORP", "Días Entrega": 14},
        {"Pos": 3, "Material": "TORNILLOS 200 KILOS", "Cantidad": 2, "Precio Unitario": 45000, "Moneda": "CLP", "Proveedor": "PERNOS S.A.", "Días Entrega": 3}
    ])

df_editado = st.data_editor(st.session_state["tabla_solped"], num_rows="dynamic", use_container_width=True)

# GRÁFICOS Y EVALUACIÓN DE MEJOR OFERTA
st.subheader("📈 Análisis y Elección de Mejor Oferta")
fig = px.bar(df_editado, x="Material", y="Precio Unitario", color="Proveedor", title="Comparativa de Precios por Ítem")
st.plotly_chart(fig, use_container_width=True)
