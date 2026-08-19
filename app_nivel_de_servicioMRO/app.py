"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Correr local: streamlit run app.py
"""

from io import BytesIO
import config
import loaders
import numpy as np
import pandas as pd
import streamlit as st
import transform

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")


# ---- 0. Función de Clasificación en 5 Categorías ----
def determinar_tipo_ariba(row):
    """
    Clasifica las solicitudes en 5 categorías independientes:
    - ⚙️ SAP ERP: Serie 1 (100...), Serie 19, CL...
    - ⚪ SAP MRP: Serie 5 (500...) o marca de Solped MRP.
    - 🟡 ARIBA DIRECTA: Flujo directo / automatizado.
    - 🟢 ARIBA CATALOGADA: Serie 6 con código de material/catálogo.
    - 🔵 ARIBA NO CATALOGADA: Serie 6 sin código de material o en Trazabilidad.
    """
    # Respetar clasificación explícita previa si ya viene en el DataFrame
    for col in ["Tipo Ariba", "Tipo_Ariba", "Origen Ariba", "Origen", "Tipo Flujo"]:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            val = str(row[col]).upper()
            if "DIRECTA" in val:
                return "🟡 ARIBA DIRECTA"
            elif "NO CATALOGAD" in val or "NOCATALOGAD" in val:
                return "🔵 ARIBA NO CATALOGADA"
            elif "CATALOGAD" in val:
                return "🟢 ARIBA CATALOGADA"
            elif "MRP" in val:
                return "⚪ SAP MRP"
            elif "ERP" in val:
                return "⚙️ SAP ERP"

    sol = str(row.get("Solicitud de pedido", "")).strip()
    material = str(row.get("Material", "")).strip()
    tiene_material = bool(material and material.lower() not in ["nan", "none", "n/a", "-", "0"])
    es_mrp_flag = str(row.get("Solped MRP", "")).strip().lower() in ["sí", "si", "true", "mrp", "1"]
    en_trazabilidad = bool(row.get("En_Trazabilidad", False) or row.get("En Trazabilidad", False))
    tipo_pedido = str(row.get("Tipo Pedido", "")).strip().upper()

    # 1. SAP MRP
    if sol.startswith("5") or es_mrp_flag:
        return "⚪ SAP MRP"

    # 2. SAP ERP
    if sol.startswith("1") or sol.startswith("19") or sol.upper().startswith("CL"):
        return "⚙️ SAP ERP"

    # 3. Flujos ARIBA (Serie 6)
    if sol.startswith("6"):
        if "DIRECTA" in tipo_pedido or "DIRECT" in tipo_pedido:
            return "🟡 ARIBA DIRECTA"
        elif en_trazabilidad or not tiene_material:
            return "🔵 ARIBA NO CATALOGADA"
        else:
            return "🟢 ARIBA CATALOGADA"

    return "⚪ OTROS"


# ---- 1. Autenticación y Reset de Caché ----
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("Dx Compradores — Nivel de Servicio")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password")
        if st.button("Ingresar"):
            if clave_ingresada == st.secrets["app_password"]:
                st.session_state["_autenticado"] = True
                loaders._descargar_onedrive.clear()
                st.session_state.pop("_clave_pipeline", None)
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

st.title("Dx Compradores — Nivel de Servicio")

# ---- 2. Configuración de Origen de Datos ----
with st.sidebar:
    st.header("Datos de entrada")
    modo = st.radio(
        "Origen de datos",
        ["OneDrive (automático)", "Subir archivos", "Archivos locales (data/)"],
        index=0,
    )

    archivo_data = archivo_resp_grupo = archivo_centro = archivo_mrp = None

    if modo == "OneDrive (automático)":
        archivo_data = "onedrive:me5a_parquet"
        archivo_resp_grupo = "onedrive:responsable_grupo_compras"
        archivo_centro = "onedrive:centro_sociedad_mro"
        archivo_mrp = "onedrive:responsable_mrp"

        if st.button("🔄 Forzar recarga desde OneDrive ahora"):
            loaders._descargar_onedrive.clear()
            st.session_state.pop("_clave_pipeline", None)
            st.rerun()

        if "onedrive" not in st.secrets:
            st.error("Falta configurar los Secrets de OneDrive en Streamlit.")
            st.stop()

    elif modo == "Subir archivos":
        archivo_data = st.file_uploader("ME5A_con_Ariba (.xlsx o .parquet)", type=["xlsx", "parquet"])
        archivo_resp_grupo = st.file_uploader("Responsable_Grupo_Compras.xlsx", type=["xlsx"])
        archivo_centro = st.file_uploader("Centro_Sociedad_MRO.xlsx", type=["xlsx"])
        archivo_mrp = st.file_uploader("Responsable_MRP.xlsx", type=["xlsx"])
        if not all([archivo_data, archivo_resp_grupo, archivo_centro, archivo_mrp]):
            st.info("Sube los 4 archivos para generar el reporte.")
            st.stop()
    else:
        archivo_data = "data/ME5A_con_Ariba.xlsx"
        archivo_resp_grupo = "data/Responsable_Grupo_Compras.xlsx"
        archivo_centro = "data/Centro_Sociedad_MRO.xlsx"
        archivo_mrp = "data/Responsable_MRP.xlsx"

    st.header("Parámetros")
    fecha_corte = st.date_input("Fecha de corte del reporte", value=pd.Timestamp.today())
    st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")


# ---- 3. Carga y Pipeline ----
def _clave_archivo(archivo):
    if hasattr(archivo, "name") and hasattr(archivo, "size"):
        return (archivo.name, archivo.size)
    return archivo


clave_actual = (
    _clave_archivo(archivo_data),
    _clave_archivo(archivo_resp_grupo),
    _clave_archivo(archivo_centro),
    _clave_archivo(archivo_mrp),
    pd.Timestamp(fecha_corte),
)

if st.session_state.get("_clave_pipeline") != clave_actual:
    df_data = loaders.cargar_data_pr(archivo_data)
    df_resp_grupo = loaders.cargar_responsable_grupo_compras(archivo_resp_grupo)
    df_centro_sociedad = loaders.cargar_centro_sociedad_mro(archivo_centro)
    df_resp_mrp = loaders.cargar_responsable_mrp(archivo_mrp)

    df_calculado = transform.pipeline_completo(
        df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=pd.Timestamp(fecha_corte)
    )
    df_calculado["Año"] = df_calculado["Fecha de pedido"].dt.year
    df_calculado["Mes"] = df_calculado["Fecha de pedido"].dt.month
    df_calculado["Día"] = df_calculado["Fecha de pedido"].dt.day

    # Asignación de la clasificación de 5 categorías
    df_calculado["Tipo Ariba"] = df_calculado.apply(determinar_tipo_ariba, axis=1)

    st.session_state["_df_pipeline"] = df_calculado
    st.session_state["_clave_pipeline"] = clave_actual

df = st.session_state["_df_pipeline"]

NS_MIN_GLOBAL = int(df["Nivel de Servicio"].min()) if len(df) and pd.notna(df["Nivel de Servicio"].min()) else 0
NS_MAX_GLOBAL = int(df["Nivel de Servicio"].max()) if len(df) and pd.notna(df["Nivel de Servicio"].max()) else 100

# ---- 4. Filtros de Usuario ----
st.subheader("Filtros")
c1, c2, c3 = st.columns(3)
with c1:
    centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
with c2:
    aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))
with c3:
    tipos_ariba = st.multiselect("Origen / Tipo Solicitud", sorted(df["Tipo Ariba"].dropna().unique()))

df_f = df.copy()
if centros:
    df_f = df_f[df_f["Centro"].isin(centros)]
if aplica:
    df_f = df_f[df_f["Aplica?"].isin(aplica)]
if tipos_ariba:
    df_f = df_f[df_f["Tipo Ariba"].isin(tipos_ariba)]

st.caption("Estado Solped")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]

# Jerarquía Temporal
df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]

with h2:
    años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
with h3:
    _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
    meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
with h4:
    _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
    fechas_disponibles = sorted(_base_dia["Fecha de pedido"].dt.date.dropna().unique())
    fechas = st.multiselect("Día", fechas_disponibles, format_func=lambda f: f.strftime("%d-%m-%Y"))

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

c4, c5 = st.columns(2)
with c4:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c5:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]

st.divider()

# ---- 5. Métricas y KPIs ----
pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
promedio_dias = df_f["Nivel de Servicio"].mean()
pedidos_distintos = df_f["Pedido"].nunique() + (1 if df_f["Pedido"].isna().any() else 0)


def tarjeta(titulo: str, valor: str, fondo: str = "rgba(64,75,85,0.07)", borde: str = "rgba(64,75,85,0.35)") -> str:
    return (
        f'<div style="background:{fondo};border:1.5px solid {borde};border-radius:8px;'
        f'padding:14px 18px;text-align:center;">'
        f'<div style="font-size:0.78rem;color:#404B55;font-weight:600;letter-spacing:.03em;'
        f'text-transform:uppercase;opacity:.85;margin-bottom:4px;">{titulo}</div>'
        f'<div style="font-size:2rem;font-weight:700;color:#404B55;line-height:1.1;">{valor}</div>'
        f'</div>'
    )


VERDE, VERDE_BORDE = "rgba(35, 145, 75, 0.16)", "rgba(35, 145, 75, 0.55)"
ROJO, ROJO_BORDE = "rgba(204, 0, 0, 0.14)", "rgba(204, 0, 0, 0.55)"

txt_dias = f"{promedio_dias:.0f}" if pd.notna(promedio_dias) else "-"
f_dias, b_dias = (ROJO, ROJO_BORDE) if pd.notna(promedio_dias) and promedio_dias > 10 else (VERDE, VERDE_BORDE)
f_pct, b_pct = (VERDE, VERDE_BORDE) if pct_cumplimiento >= 85 else (ROJO, ROJO_BORDE)

t1, t2, t3 = st.columns(3)
with t1:
    st.markdown(tarjeta("Promedio días de gestión", txt_dias, f_dias, b_dias), unsafe_allow_html=True)
with t2:
    st.markdown(tarjeta("% Cumplimiento", f"{pct_cumplimiento:.0f}%", f_pct, b_pct), unsafe_allow_html=True)
with t3:
    st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

st.divider()


# ---- 6. Tablas de Resultados ----
def tabla_enaex(tabla: pd.DataFrame) -> str:
    cols = list(tabla.columns)

    def _fmt(col, val):
        if pd.isna(val):
            return "-"
        if col == "Promedio días de gestión":
            return f"{val:,.0f}"
        if col == "% Cumplimiento":
            return f"{val:,.0f}%"
        if col == "Pos. OC generadas":
            return f"{val:,.0f}"
        return str(val)

    filas = []
    for i, r in enumerate(tabla.itertuples(index=False)):
        es_total = str(r[0]) == "TOTAL"
        fondo = "#404B55" if es_total else ("#ffffff" if i % 2 == 0 else "#f4f5f7")
        color = "#ffffff" if es_total else "#404B55"
        estilo_fila = f"background:{fondo};color:{color};font-weight:{'700' if es_total else '400'};"
        celdas = [
            f'<td style="padding:7px 12px;text-align:{"left" if j==0 else "right"};border-bottom:1px solid #e3e5e8;white-space:nowrap;">{_fmt(c, r[j])}</td>'
            for j, c in enumerate(cols)
        ]
        filas.append(f'<tr style="{estilo_fila}">{"".join(celdas)}</tr>')

    encabezados = "".join(
        f'<th style="padding:9px 12px;text-align:{"left" if j==0 else "right"};background:#404B55;color:#fff;font-weight:600;font-size:0.82rem;white-space:nowrap;">{c}</th>'
        for j, c in enumerate(cols)
    )
    tabla_html = f'<table style="width:100%;border-collapse:collapse;font-size:0.86rem;border:1px solid #d8dbdf;"><thead><tr>{encabezados}</tr></thead><tbody>{"".join(filas)}</tbody></table>'
    return f'<div style="overflow-x:auto;border:1px solid #d8dbdf;border-radius:4px;">{tabla_html}</div>'


st.subheader("Por comprador")
col_comprador = "Comprador (Grupo de compras)" if "Comprador (Grupo de compras)" in df_f.columns else "Comprador por Grupo Compras"
tabla_comprador = transform.calcular_metricas_por_grupo(df_f, [col_comprador])
tabla_comprador = transform.agregar_fila_total(tabla_comprador, df_f, [col_comprador])
st.markdown(tabla_enaex(tabla_comprador), unsafe_allow_html=True)

# ---- 7. Detalle de Solicitudes con Etiquetas ----
COLUMNAS_DETALLE = [
    "Centro", "Material", "Texto breve", "Solicitud de pedido", "Tipo Ariba",
    "Fecha de solicitud", "Fecha modificación", "Grupo de compras", "Cantidad pedida",
    "Pedido", "Fecha de pedido", "Posición de pedido", "Comprador (Grupo de compras)",
    "Solped MRP", "Nombre Centro 2", "Estado Solped", "Nivel de Servicio", "Comentario", "Cumple"
]


def preparar_detalle(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    for col_fecha in ["Fecha de solicitud", "Fecha modificación", "Fecha de pedido"]:
        if col_fecha in d.columns:
            d[col_fecha] = pd.to_datetime(d[col_fecha], errors="coerce").dt.date
    if "Comentario" not in d.columns:
        d["Comentario"] = ""
    return d[[c for c in COLUMNAS_DETALLE if c in d.columns]]


detalle = preparar_detalle(df_f)

with st.expander("Ver detalle de solicitudes", expanded=False):
    # Leyenda visual de las 5 categorías
    st.markdown(
        """
        <div style="display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap;">
            <div style="background:#d4edda; color:#155724; padding:5px 10px; border-radius:5px; font-weight:bold; font-size:0.78rem;">🟢 ARIBA CATALOGADA</div>
            <div style="background:#d1ecf1; color:#0c5460; padding:5px 10px; border-radius:5px; font-weight:bold; font-size:0.78rem;">🔵 ARIBA NO CATALOGADA</div>
            <div style="background:#fff3cd; color:#856404; padding:5px 10px; border-radius:5px; font-weight:bold; font-size:0.78rem;">🟡 ARIBA DIRECTA</div>
            <div style="background:#e2e3e5; color:#383d41; padding:5px 10px; border-radius:5px; font-weight:bold; font-size:0.78rem;">⚪ SAP MRP</div>
            <div style="background:#f8d7da; color:#721c24; padding:5px 10px; border-radius:5px; font-weight:bold; font-size:0.78rem;">⚙️ SAP ERP</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.data_editor(
        detalle,
        use_container_width=True,
        num_rows="fixed",
        key="editor_detalle",
        column_config={
            "Tipo Ariba": st.column_config.TextColumn("Origen / Tipo Solicitud", width="medium"),
            "Comentario": st.column_config.TextColumn("Comentario", width="medium"),
        },
        disabled=[c for c in detalle.columns if c != "Comentario"],
    )
