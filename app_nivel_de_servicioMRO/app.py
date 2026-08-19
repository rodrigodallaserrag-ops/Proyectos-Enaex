"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Correr local:  streamlit run app.py
"""
import streamlit as st
import pandas as pd
import numpy as np

import config, loaders, transform

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

# ---- Función de Clasificación Flujo SAP / Ariba ----
def determinar_tipo_ariba(row):
    """
    Clasifica la solicitud integrando los flujos SAP y Ariba según reglas unificadas:
    - Serie 1 (100) y Serie 5 (500):
        * Encargado Cesar -> ⚪ SAP MRP
        * Otro encargado -> ⚪ SAP ERP
    - Serie 6 (600):
        * Sin material o registrada en Trazabilidad -> 🔵 ARIBA NO CATALOGADA
        * Con código de material / catálogo -> 🟢 ARIBA DIRECTA / CATALOGADA
    """
    # 1. Respetar clasificación previa explícita si existe
    for col in ["Tipo Ariba", "Tipo_Ariba", "Origen Ariba", "Origen", "Tipo Flujo"]:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            val = str(row[col]).upper()
            if "NO CATALOGAD" in val or "NOCATALOGAD" in val:
                return "🔵 ARIBA NO CATALOGADA"
            elif "CATALOGAD" in val or "DIRECTA" in val:
                return "🟢 ARIBA DIRECTA / CATALOGADA"
            elif "MRP" in val:
                return "⚪ SAP MRP"
            elif "ERP" in val:
                return "⚪ SAP ERP"

    sol = str(row.get("Solicitud de pedido", "")).strip()
    material = str(row.get("Material", "")).strip()
    tiene_material = bool(material and material.lower() not in ["nan", "none", "n/a", "-", "0"])

    # Obtener nombre del encargado/responsable desde la fila
    encargado = str(
        row.get("Responsable MRP", row.get("Encargado", row.get("Responsable", row.get("Comprador (Grupo de compras)", ""))))
    ).upper()
    es_cesar = "CESAR" in encargado or "CÉSAR" in encargado

    # Serie 1 (100) y Serie 5 (500)
    if sol.startswith("1") or sol.startswith("5"):
        return "⚪ SAP MRP" if es_cesar else "⚪ SAP ERP"

    # Serie 6 (600)
    if sol.startswith("6"):
        if not tiene_material:
            return "🔵 ARIBA NO CATALOGADA"
        else:
            return "🟢 ARIBA DIRECTA / CATALOGADA"

    return "⚪ OTROS"


# ---- Acceso con contraseña ----
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

# ---- Fuente de datos: local, subida manual, u OneDrive automático ----
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
            st.error("Falta configurar los Secrets de OneDrive. Mientras tanto, usa 'Subir archivos'.")
            st.stop()

    elif modo == "Subir archivos":
        archivo_data = st.file_uploader(
            "ME5A_con_Ariba (.xlsx o .parquet)", type=["xlsx", "parquet"],
            help="Sube el archivo ME5A para procesar.",
        )
        archivo_resp_grupo = st.file_uploader("Responsable_Grupo_Compras.xlsx", type="xlsx")
        archivo_centro = st.file_uploader("Centro_Sociedad_MRO.xlsx", type="xlsx")
        archivo_mrp = st.file_uploader("Responsable_MRP.xlsx", type="xlsx")
        if not all([archivo_data, archivo_resp_grupo, archivo_centro, archivo_mrp]):
            st.info("Sube los 4 archivos para generar el reporte.")
            st.stop()
    else:
        archivo_data = "data/ME5A_con_Ariba.xlsx"
        archivo_resp_grupo = "data/Responsable_Grupo_Compras.xlsx"
        archivo_centro = "data/Centro_Sociedad_MRO.xlsx"
        archivo_mrp = "data/Responsable_MRP.xlsx"

with st.sidebar:
    st.header("Parámetros")
    fecha_corte = st.date_input("Fecha de corte del reporte (FechaCorteReporte)", value=pd.Timestamp.today())
    st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")

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
    
    # 🔹 Cálculo inmediato de la nueva clasificación al cargar
    df_calculado["Tipo Ariba"] = df_calculado.apply(determinar_tipo_ariba, axis=1)

    st.session_state["_df_pipeline"] = df_calculado
    st.session_state["_clave_pipeline"] = clave_actual

df = st.session_state["_df_pipeline"]

NS_MIN_GLOBAL = int(df["Nivel de Servicio"].min()) if len(df) and pd.notna(df["Nivel de Servicio"].min()) else 0
NS_MAX_GLOBAL = int(df["Nivel de Servicio"].max()) if len(df) and pd.notna(df["Nivel de Servicio"].max()) else 100

checkpoints = [("0. Total tras el pipeline (sin filtros)", len(df))]
metricas_por_etapa = []

def _snapshot(nombre, d):
    n = len(d)
    pct = (d["Cumple"] == "Cumple").sum() / n * 100 if n else 0
    prom = d["Nivel de Servicio"].mean() if n else float("nan")
    pos_oc = d["Pedido"].nunique() + (1 if d["Pedido"].isna().any() else 0)
    metricas_por_etapa.append(
        (nombre, f"{n:,}", f"{pct:.0f}%", f"{prom:.0f}" if pd.notna(prom) else "-", f"{pos_oc:,}")
    )

_snapshot("0. Sin filtros", df)

# ---- Sección de Filtros de Usuario ----
st.subheader("Filtros")
c1, c2, c3 = st.columns(3)
with c1:
    centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
with c2:
    aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))
with c3:
    tipos_ariba = st.multiselect("Tipo / Origen", sorted(df["Tipo Ariba"].dropna().unique()))

df_f = df.copy()
if centros:
    df_f = df_f[df_f["Centro"].isin(centros)]
if aplica:
    df_f = df_f[df_f["Aplica?"].isin(aplica)]
if tipos_ariba:
    df_f = df_f[df_f["Tipo Ariba"].isin(tipos_ariba)]

checkpoints.append(("1. Tras filtros principales", len(df_f)))
_snapshot("1. Tras filtros principales", df_f)

# ---- Estado Solped ----
st.caption("Estado Solped (el filtro de fecha de abajo solo aplica dentro de 'Pedido completo')")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]

# ---- Jerarquía Año / Mes / Día ----
df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]

with h2:
    años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
with h3:
    _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
    meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
with h4:
    _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
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

# ---- Solped MRP y Cumple ----
c4, c5 = st.columns(2)
with c4:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c5:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]

# ---- Filtro por Rango Nivel de Servicio ----
st.divider()
st.subheader("Filtro por días de gestión (Nivel de Servicio)")

if len(df_f):
    n_negativos = (df_f["Nivel de Servicio"] < 0).sum()

    f1, f2 = st.columns([1, 2])
    with f1:
        excluir_negativos = st.checkbox(
            "Excluir negativos (solo desde 0)",
            value=False,
            help="Excluye filas con días negativos para no distorsionar promedios."
        )
    with f2:
        if excluir_negativos:
            rango_ns = (0, NS_MAX_GLOBAL)
            st.caption(f"Rango aplicado: 0 a {NS_MAX_GLOBAL:,} días")
        elif NS_MIN_GLOBAL < NS_MAX_GLOBAL:
            rango_ns = st.slider(
                "Rango de días de gestión",
                min_value=NS_MIN_GLOBAL,
                max_value=NS_MAX_GLOBAL,
                value=(NS_MIN_GLOBAL, NS_MAX_GLOBAL),
                key="slider_rango_dias",
            )
        else:
            rango_ns = (NS_MIN_GLOBAL, NS_MAX_GLOBAL)

    df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

# ---- Métricas y Tablas Principales ----
pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
promedio_dias = df_f["Nivel de Servicio"].mean()
pedidos_distintos = df_f["Pedido"].nunique() + (1 if df_f["Pedido"].isna().any() else 0)

VERDE, VERDE_BORDE = "rgba(35, 145, 75, 0.16)", "rgba(35, 145, 75, 0.55)"
ROJO, ROJO_BORDE = "rgba(204, 0, 0, 0.14)", "rgba(204, 0, 0, 0.55)"

def tarjeta(titulo: str, valor: str, fondo: str = "rgba(64,75,85,0.07)", borde: str = "rgba(64,75,85,0.35)") -> str:
    return (
        f'<div style="background:{fondo};border:1.5px solid {borde};border-radius:8px;'
        f'padding:14px 18px;text-align:center;">'
        f'<div style="font-size:0.78rem;color:#404B55;font-weight:600;letter-spacing:.03em;'
        f'text-transform:uppercase;opacity:.85;margin-bottom:4px;">{titulo}</div>'
        f'<div style="font-size:2rem;font-weight:700;color:#404B55;line-height:1.1;">{valor}</div>'
        f"</div>"
    )

f_dias, b_dias, txt_dias = (ROJO, ROJO_BORDE, f"{promedio_dias:.0f}") if pd.notna(promedio_dias) and promedio_dias > 10 else (VERDE, VERDE_BORDE, f"{promedio_dias:.0f}" if pd.notna(promedio_dias) else "-")
f_pct, b_pct = (VERDE, VERDE_BORDE) if pct_cumplimiento >= 85 else (ROJO, ROJO_BORDE)

t1, t2, t3 = st.columns(3)
with t1:
    st.markdown(tarjeta("Promedio días de gestión", txt_dias, f_dias, b_dias), unsafe_allow_html=True)
with t2:
    st.markdown(tarjeta("% Cumplimiento", f"{pct_cumplimiento:.0f}%", f_pct, b_pct), unsafe_allow_html=True)
with t3:
    st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

st.divider()

# ---- Tablas de Resultados ----
ENAEX_GRIS, ENAEX_ROJO = "#404B55", "#CC0000"

def tabla_enaex(tabla: pd.DataFrame, max_height: int | None = None, compacta: bool = False) -> str:
    cols = list(tabla.columns)
    def _fmt(col, val):
        if pd.isna(val): return "-"
        if col == "Promedio días de gestión": return f"{val:,.0f}"
        if col == "% Cumplimiento": return f"{val:,.0f}%"
        if col == "Pos. OC generadas": return f"{val:,.0f}"
        return str(val)

    pad = "4px 6px" if compacta else "7px 12px"
    pad_th = "5px 6px" if compacta else "9px 12px"
    fuente = "0.72rem" if compacta else "0.86rem"
    fuente_th = "0.66rem" if compacta else "0.82rem"

    filas = []
    for i, r in enumerate(tabla.itertuples(index=False)):
        es_total = str(r[0]) == "TOTAL"
        fondo = ENAEX_GRIS if es_total else ("#ffffff" if i % 2 == 0 else "#f4f5f7")
        color_texto = "#fff" if es_total else ENAEX_GRIS
        estilo_fila = f"background:{fondo};color:{color_texto};font-weight:{'700' if es_total else '400'};"
        celdas = [f'<td style="padding:{pad};text-align:{"left" if j==0 else "right"};border-bottom:1px solid #e3e5e8;white-space:nowrap;">{_fmt(c, r[j])}</td>' for j, c in enumerate(cols)]
        filas.append(f'<tr style="{estilo_fila}">{"".join(celdas)}</tr>')

    encabezados = "".join(f'<th style="padding:{pad_th};text-align:{"left" if j==0 else "right"};background:{ENAEX_GRIS};color:#fff;font-weight:600;font-size:{fuente_th};sticky;top:0;z-index:2;white-space:nowrap;">{c}</th>' for j, c in enumerate(cols))
    tabla_html = f'<table style="width:100%;border-collapse:collapse;font-size:{fuente};font-family:inherit;border:1px solid #d8dbdf;"><thead><tr>{encabezados}</tr></thead><tbody>{"".join(filas)}</tbody></table>'
    alto = f"max-height:{max_height}px;overflow-y:auto;" if max_height else ""
    return f'<div style="{alto}overflow-x:auto;border:1px solid #d8dbdf;border-radius:4px;">{tabla_html}</div>'

st.subheader("Por comprador")
col_comprador = "Comprador (Grupo de compras)" if "Comprador (Grupo de compras)" in df_f.columns else "Comprador por Grupo Compras"
tabla_comprador = transform.calcular_metricas_por_grupo(df_f, [col_comprador])
tabla_comprador = transform.agregar_fila_total(tabla_comprador, df_f, [col_comprador])
st.markdown(tabla_enaex(tabla_comprador), unsafe_allow_html=True)

# ---- Detalle de Solicitudes ----
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
    st.caption("La columna **Comentario** es editable.")
    st.markdown(
        """
        <div style="display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap;">
            <div style="background:#d4edda; color:#155724; padding:6px 12px; border-radius:6px; font-weight:bold; font-size:0.8rem;">🟢 ARIBA DIRECTA / CATALOGADA</div>
            <div style="background:#d1ecf1; color:#0c5460; padding:6px 12px; border-radius:6px; font-weight:bold; font-size:0.8rem;">🔵 ARIBA NO CATALOGADA</div>
            <div style="background:#e2e3e5; color:#383d41; padding:6px 12px; border-radius:6px; font-weight:bold; font-size:0.8rem;">⚪ SAP ERP</div>
            <div style="background:#e2e3e5; color:#383d41; padding:6px 12px; border-radius:6px; font-weight:bold; font-size:0.8rem;">⚪ SAP MRP</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    detalle_editado = st.data_editor(
        detalle,
        use_container_width=True,
        num_rows="fixed",
        key="editor_detalle",
        column_config={
            "Tipo Ariba": st.column_config.TextColumn("Tipo / Origen", width="medium"),
            "Comentario": st.column_config.TextColumn("Comentario", width="medium"),
        },
        disabled=[c for c in detalle.columns if c != "Comentario"],
    )
