"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Correr local:  streamlit run src/app.py
"""
import streamlit as st
import pandas as pd

import config, loaders, transform

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

st.title("Dx Compradores — Nivel de Servicio")

# ---- Fuente de datos: local (desarrollo) o subida manual (Streamlit Cloud) ----
# En Streamlit Cloud no existe la ruta de red de la empresa, así que mientras
# no está la conexión a SharePoint/Azure, se sube el Excel a mano acá.
with st.sidebar:
    st.header("Datos de entrada")
    modo = st.radio("Origen de datos", ["Archivos locales (data/)", "Subir archivos"], index=1)

    archivo_data = archivo_resp_grupo = archivo_centro = archivo_mrp = None
    if modo == "Subir archivos":
        archivo_data = st.file_uploader("ME5A_con_Ariba.xlsx", type="xlsx")
        archivo_resp_grupo = st.file_uploader("Responsable_Grupo_Compras.xlsx", type="xlsx")
        archivo_centro = st.file_uploader("Centro_Sociedad_MRO.xlsx", type="xlsx")
        archivo_mrp = st.file_uploader("Responsable_MRP.xlsx", type="xlsx")
        if not all([archivo_data, archivo_resp_grupo, archivo_centro, archivo_mrp]):
            st.info("Sube los 4 archivos para generar el reporte.")
            st.stop()

# ---- Sidebar: parámetro FechaCorteReporte (hoy se ingresaba a mano en Power Query) ----
with st.sidebar:
    st.header("Parámetros")
    fecha_corte = st.date_input("Fecha de corte del reporte (FechaCorteReporte)", value=pd.Timestamp.today())
    st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")

# ---- Carga de datos (equivalente a refrescar las queries del pbix) ----
df_data = loaders.cargar_data_pr(archivo_data)
df_resp_grupo = loaders.cargar_responsable_grupo_compras(archivo_resp_grupo)
df_centro_sociedad = loaders.cargar_centro_sociedad_mro(archivo_centro)
df_resp_mrp = loaders.cargar_responsable_mrp(archivo_mrp)

df = transform.pipeline_completo(
    df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=pd.Timestamp(fecha_corte)
)

# ---- Filtros (slicers del pbix) ----
df["Año"] = df["Fecha de pedido"].dt.year
df["Mes"] = df["Fecha de pedido"].dt.month
df["Día"] = df["Fecha de pedido"].dt.day

st.subheader("Filtros")
c1, c2 = st.columns(2)
with c1:
    centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
with c2:
    aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))

df_f = df.copy()
if centros:
    df_f = df_f[df_f["Centro"].isin(centros)]
if aplica:
    df_f = df_f[df_f["Aplica?"].isin(aplica)]

# Jerarquía: Estado Solped -> Año -> Mes -> Día (de Fecha de pedido).
# OJO: el filtro de fecha solo restringe las filas "Pedido completo".
# "Sin pedido" y "Pedido incompleto", si están seleccionados, se muestran
# completos sin que la fecha los afecte (igual que en el pbix).
st.caption("Estado Solped (el filtro de fecha solo aplica dentro de 'Pedido completo')")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

df_f = df_f[df_f["Estado Solped"].isin(estados)] if estados else df_f

# Opciones de Año/Mes/Día calculadas SOLO sobre el subconjunto "Pedido completo"
df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]

with h2:
    años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
with h3:
    _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
    meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
with h4:
    _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
    dias = st.multiselect("Día", sorted(_base_dia["Día"].dropna().unique().astype(int)))

if años or meses or dias:
    es_pedido_completo = df_f["Estado Solped"] == "Pedido completo"
    cond_fecha = pd.Series(True, index=df_f.index)
    if años:
        cond_fecha &= df_f["Año"].isin(años)
    if meses:
        cond_fecha &= df_f["Mes"].isin(meses)
    if dias:
        cond_fecha &= df_f["Día"].isin(dias)
    # Se mantienen: filas que NO son "Pedido completo" (intactas) + las que
    # SÍ son "Pedido completo" y además cumplen la fecha seleccionada.
    df_f = df_f[~es_pedido_completo | (es_pedido_completo & cond_fecha)]

c3, c4 = st.columns(2)
with c3:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c4:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

for col, valores in [("Solped MRP", solped_mrp), ("Cumple", cumple)]:
    if valores:
        df_f = df_f[df_f[col].isin(valores)]

# ---- Tarjetas (cardVisual del pbix) ----
c1, c2, c3 = st.columns(3)
pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
promedio_dias = df_f["Nivel de Servicio"].mean()
total_pedidos = df_f["Pedido"].nunique()

c1.metric("% Cumplimiento", f"{pct_cumplimiento:.1f}%")
c2.metric("Promedio días de gestión", f"{promedio_dias:.1f}" if pd.notna(promedio_dias) else "-")
c3.metric("Pedidos (OC)", f"{total_pedidos:,}")

st.divider()

# ---- Tabla por Comprador (tableEx #1 del pbix) ----
st.subheader("Por comprador")
tabla_comprador = transform.calcular_metricas_por_grupo(df_f, ["Comprador por Grupo Compras"])
st.dataframe(tabla_comprador, use_container_width=True)

# ---- Tabla por Centro (tableEx #2 del pbix) ----
st.subheader("Por centro")
group_cols_centro = [c for c in ["Centro", "Nombre Centro"] if c in df_f.columns]
tabla_centro = transform.calcular_metricas_por_grupo(df_f, group_cols_centro)
st.dataframe(tabla_centro, use_container_width=True)

with st.expander("Ver detalle de solicitudes"):
    st.dataframe(df_f, use_container_width=True)

