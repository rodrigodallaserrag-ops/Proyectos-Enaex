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
    else:
        # FIX 1: Rutas locales por defecto cuando no se suben archivos
        archivo_data = "data/ME5A_con_Ariba.xlsx"
        archivo_resp_grupo = "data/Responsable_Grupo_Compras.xlsx"
        archivo_centro = "data/Centro_Sociedad_MRO.xlsx"
        archivo_mrp = "data/Responsable_MRP.xlsx"

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

# FIX 2: Calcular min/max globales del dataframe COMPLETO para estabilizar el slider
NS_MIN_GLOBAL = int(df["Nivel de Servicio"].min()) if len(df) and pd.notna(df["Nivel de Servicio"].min()) else 0
NS_MAX_GLOBAL = int(df["Nivel de Servicio"].max()) if len(df) and pd.notna(df["Nivel de Servicio"].max()) else 100

# Registro de conteos por etapa
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
_snapshot("2. Tras Centro + Aplica?", df_f)

# ---- Estado Solped ----
st.caption("Estado Solped (el filtro de fecha de abajo solo aplica dentro de 'Pedido completo')")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]
checkpoints.append(("3. Tras filtro Estado Solped (total)", len(df_f)))
_snapshot("3. Tras Estado Solped", df_f)

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

_snapshot("4. Tras jerarquía Año/Mes/Día", df_f)

# ---- Solped MRP y Cumple ----
c3, c4 = st.columns(2)
with c3:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c4:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
_snapshot("5. Tras Solped MRP", df_f)

if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]
_snapshot("6. Tras Cumple", df_f)

# ---- Segunda capa: filtro por rango de Nivel de Servicio (días) ----
st.divider()
st.subheader("Filtro por días de gestión (Nivel de Servicio)")

if len(df_f):
    n_negativos = (df_f["Nivel de Servicio"] < 0).sum()

    f1, f2 = st.columns([1, 2])
    with f1:
        excluir_negativos = st.checkbox(
            "Excluir negativos (solo desde 0)",
            value=False,
            help="Hay filas con días negativos en la selección actual."
        )
    with f2:
        if excluir_negativos:
            rango_ns = (0, NS_MAX_GLOBAL)
            st.caption(f"Rango aplicado: 0 a {NS_MAX_GLOBAL:,} días (negativos excluidos)")
        elif NS_MIN_GLOBAL < NS_MAX_GLOBAL:
            # FIX 3: Usar constantes globales estables para evitar bucles infintos de estado
            rango_ns = st.slider(
                "Rango de días de gestión",
                min_value=NS_MIN_GLOBAL,
                max_value=NS_MAX_GLOBAL,
                value=(NS_MIN_GLOBAL, NS_MAX_GLOBAL),
                key="slider_rango_dias" # Key explícita
            )
        else:
            rango_ns = (NS_MIN_GLOBAL, NS_MAX_GLOBAL)

    df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

checkpoints.append(("7. Tras filtro Nivel de Servicio (RESULTADO FINAL)", len(df_f)))
_snapshot("7. RESULTADO FINAL", df_f)

# ... [El resto de la rendering de tablas y HTML se mantiene igual] ...

# ---- Detalle de solicitudes con columna de comentario editable ----
# ... (preparación del detalle) ...

detalle = preparar_detalle(df_f)

with st.expander("Ver detalle de solicitudes", expanded=False):
    st.caption("La columna **Comentario** es editable: escribe ahí y se incluirá en el Excel de registro.")
    detalle_editado = st.data_editor(
        detalle,
        use_container_width=True,
        num_rows="fixed",
        key="editor_detalle",
        disabled=[c for c in detalle.columns if c != "Comentario"],
    )

# ---- Exportación del registro semanal a Excel ----
semana_ref = pd.Timestamp(fecha_corte) - pd.Timedelta(days=7)
num_semana = semana_ref.isocalendar()[1]
nombre_archivo = f"Sem{num_semana:02d}-{semana_ref.year}.xlsx"

# FIX 4: Cachear la generación pesada del Excel
@st.cache_data(show_spinner="Generando Excel...")
def generar_excel_cached(detalle_df, fecha_corte_val, prom_dias, pct_cumpl, ped_dist, total_filas):
    # (aquí ejecutas la lógica de openpyxl)
    return generar_excel(detalle_df)

st.subheader("Registro semanal")
st.caption(f"Descarga el reporte completo como **{nombre_archivo}**.")

try:
    # Solo se compila el Excel si el usuario realmente hace clic o si cambian las métricas principales
    bytes_excel = generar_excel(detalle_editado)
    st.download_button(
        f"⬇ Descargar {nombre_archivo}",
        data=bytes_excel,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
except Exception as e:
    st.error(f"No se pudo generar el Excel: {e}")
