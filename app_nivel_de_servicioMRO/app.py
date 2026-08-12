"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Correr local:  streamlit run app.py
"""
import streamlit as st
import pandas as pd

import config, loaders, transform

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

st.title("Dx Compradores — Nivel de Servicio")

# ---- Fuente de datos: local (desarrollo) o subida manual (Streamlit Cloud) ----
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

with st.sidebar:
    st.header("Parámetros")
    fecha_corte = st.date_input("Fecha de corte del reporte (FechaCorteReporte)", value=pd.Timestamp.today())
    st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")

# ---- Carga de datos ----
df_data = loaders.cargar_data_pr(archivo_data)
df_resp_grupo = loaders.cargar_responsable_grupo_compras(archivo_resp_grupo)
df_centro_sociedad = loaders.cargar_centro_sociedad_mro(archivo_centro)
df_resp_mrp = loaders.cargar_responsable_mrp(archivo_mrp)

df = transform.pipeline_completo(
    df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=pd.Timestamp(fecha_corte)
)
df["Año"] = df["Fecha de pedido"].dt.year
df["Mes"] = df["Fecha de pedido"].dt.month
df["Día"] = df["Fecha de pedido"].dt.day

# Registro de conteos por etapa - para el panel de diagnóstico al final
checkpoints = [("0. Total tras el pipeline (sin filtros)", len(df))]

# ---- Filtros: Centro y Aplica? ----
st.subheader("Filtros")
c1, c2 = st.columns(2)
with c1:
    centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
with c2:
    aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))

df_f = df.copy()
if centros:
    df_f = df_f[df_f["Centro"].isin(centros)]
checkpoints.append(("1. Tras filtro Centro", len(df_f)))

if aplica:
    df_f = df_f[df_f["Aplica?"].isin(aplica)]
checkpoints.append(("2. Tras filtro Aplica?", len(df_f)))

# ---- Estado Solped ----
st.caption("Estado Solped (el filtro de fecha de abajo solo aplica dentro de 'Pedido completo')")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]
checkpoints.append(("3. Tras filtro Estado Solped", len(df_f)))

conteo_sin_pedido_antes = (df_f["Estado Solped"] == "Sin pedido").sum()
conteo_incompleto_antes = (df_f["Estado Solped"] == "Pedido incompleto").sum()
conteo_completo_antes = (df_f["Estado Solped"] == "Pedido completo").sum()

# ---- Jerarquía Año / Mes / Día (opciones calculadas solo sobre "Pedido completo") ----
df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]

with h2:
    años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
with h3:
    _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
    meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
with h4:
    _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
    # Fechas completas (no solo el número de día) para no mezclar, por ejemplo,
    # el 30 de julio con el 30 de agosto si se seleccionan ambos meses.
    fechas_disponibles = sorted(_base_dia["Fecha de pedido"].dt.date.dropna().unique())
    fechas = st.multiselect(
        "Día", fechas_disponibles, format_func=lambda f: f.strftime("%d-%m-%Y")
    )

if años or meses or fechas:
    es_pedido_completo = df_f["Estado Solped"] == "Pedido completo"
    cond_fecha = pd.Series(True, index=df_f.index)
    if años:
        cond_fecha &= df_f["Año"].isin(años)
    if meses:
        cond_fecha &= df_f["Mes"].isin(meses)
    if fechas:
        cond_fecha &= df_f["Fecha de pedido"].dt.date.isin(fechas)
    df_f = df_f[~es_pedido_completo | (es_pedido_completo & cond_fecha)]

checkpoints.append(("4a. Sin pedido tras jerarquía (debe ser IGUAL al paso 3)", (df_f["Estado Solped"] == "Sin pedido").sum()))
checkpoints.append(("4b. Pedido incompleto tras jerarquía (debe ser IGUAL al paso 3)", (df_f["Estado Solped"] == "Pedido incompleto").sum()))
checkpoints.append(("4c. Pedido completo tras jerarquía (este SÍ puede bajar)", (df_f["Estado Solped"] == "Pedido completo").sum()))
checkpoints.append(("4. Total tras jerarquía de fecha", len(df_f)))

# ---- Solped MRP y Cumple ----
c3, c4 = st.columns(2)
with c3:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c4:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
checkpoints.append(("5. Tras filtro Solped MRP", len(df_f)))

if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]
checkpoints.append(("6. Tras filtro Cumple (RESULTADO FINAL)", len(df_f)))

# ---- Panel de diagnóstico ----
with st.expander("🔍 Diagnóstico de filtrado (para comparar contra el pbix)"):
    st.write(
        "Filas en cada etapa. Los pasos 4a y 4b deben quedar IGUAL al paso 3 "
        "(la jerarquía de fecha no debe tocar 'Sin pedido' ni 'Pedido incompleto'). "
        "Si en el pbix ves un número distinto, compara etapa por etapa hasta encontrar "
        "en cuál empiezan a diferir."
    )
    st.table(pd.DataFrame(checkpoints, columns=["Etapa", "Filas"]))

    st.write("Detalle de las solicitudes en el resultado final, para cruzar 1 a 1 contra el export del pbix:")
    st.dataframe(
        df_f[["Solicitud de pedido", "Centro", "Estado Solped", "Fecha de pedido", "Nivel de Servicio", "Cumple", "Solped MRP"]]
        .sort_values("Solicitud de pedido"),
        use_container_width=True,
    )
    csv = df_f.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar detalle filtrado (CSV)", csv, "detalle_filtrado.csv", "text/csv")

# ---- Tarjetas ----
c1, c2, c3 = st.columns(3)
pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
promedio_dias = df_f["Nivel de Servicio"].mean()
total_pedidos = df_f["Pedido"].nunique()

c1.metric("% Cumplimiento", f"{pct_cumplimiento:.1f}%")
c2.metric("Promedio días de gestión", f"{promedio_dias:.1f}" if pd.notna(promedio_dias) else "-")
c3.metric("Pedidos (OC)", f"{total_pedidos:,}")

st.divider()

# ---- Tabla por Comprador ----
st.subheader("Por comprador")
tabla_comprador = transform.calcular_metricas_por_grupo(df_f, ["Comprador por Grupo Compras"])
st.dataframe(tabla_comprador, use_container_width=True)

# ---- Tabla por Centro ----
st.subheader("Por centro")
group_cols_centro = [c for c in ["Centro", "Nombre Centro"] if c in df_f.columns]
tabla_centro = transform.calcular_metricas_por_grupo(df_f, group_cols_centro)
st.dataframe(tabla_centro, use_container_width=True)

with st.expander("Ver detalle de solicitudes"):
    st.dataframe(df_f, use_container_width=True)
