"""
Streamlit - Dx Compradores (Versión Monolítica - Todo en 1 solo archivo)
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Ejecución local: streamlit run app.py
"""

import io
import requests
import numpy as np
import pandas as pd
import streamlit as st

# ==============================================================================
# 1. CONFIGURACIÓN, SLAs Y PALETA DE COLORES
# ==============================================================================
SLA_DIAS_ERP_MRP = 10
SLA_DIAS_ARIBA = 7

# Colores Corporativos Enaex / Semáforos KPI
ENAEX_GRIS = "#404B55"
ENAEX_ROJO = "#CC0000"

VERDE = "rgba(35, 145, 75, 0.16)"
VERDE_BORDE = "rgba(35, 145, 75, 0.55)"
ROJO = "rgba(204, 0, 0, 0.14)"
ROJO_BORDE = "rgba(204, 0, 0, 0.55)"

# ==============================================================================
# 2. CARGA DE DATOS (OneDrive / Archivos locales / Uploaders)
# ==============================================================================
@st.cache_data(ttl=3600)
def _descargar_onedrive(key_secret: str) -> io.BytesIO:
    """Descarga un archivo desde OneDrive usando la URL guardada en st.secrets."""
    if "onedrive" not in st.secrets or key_secret not in st.secrets["onedrive"]:
        raise ValueError(f"No se encontró la clave '{key_secret}' en st.secrets['onedrive'].")
    
    url = st.secrets["onedrive"][key_secret]
    response = requests.get(url)
    response.raise_for_status()
    return io.BytesIO(response.content)


def _cargar_archivo(fuente):
    """Auxiliar para leer Parquet o Excel desde OneDrive, UploadedFile o ruta local."""
    if isinstance(fuente, str) and fuente.startswith("onedrive:"):
        key = fuente.split("onedrive:")[1]
        buffer = _descargar_onedrive(key)
        if "parquet" in key.lower():
            return pd.read_parquet(buffer)
        return pd.read_excel(buffer)
    
    if hasattr(fuente, "name"):
        if fuente.name.endswith(".parquet"):
            return pd.read_parquet(fuente)
        return pd.read_excel(fuente)
    
    if isinstance(fuente, str):
        if fuente.endswith(".parquet"):
            return pd.read_parquet(fuente)
        return pd.read_excel(fuente)
    
    return pd.read_excel(fuente)


def cargar_data_pr(archivo):
    return _cargar_archivo(archivo)


def cargar_responsable_grupo_compras(archivo):
    return _cargar_archivo(archivo)


def cargar_centro_sociedad_mro(archivo):
    return _cargar_archivo(archivo)


def cargar_responsable_mrp(archivo):
    return _cargar_archivo(archivo)


# ==============================================================================
# 3. TRANSFORMACIÓN DE DATOS Y LÓGICA DE NEGOCIO
# ==============================================================================
def pipeline_completo(df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=None):
    """Ejecuta el pipeline de joins y cálculo de Nivel de Servicio."""
    df = df_data.copy()

    # Conversión de fechas
    for col in ["Fecha de solicitud", "Fecha modificación", "Fecha de pedido"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Mapeo de Comprador / Grupo de Compras
    if "Comprador (Grupo de compras)" not in df.columns:
        if "Grupo de compras" in df.columns:
            df["Comprador (Grupo de compras)"] = df["Grupo de compras"]
        else:
            df["Comprador (Grupo de compras)"] = "Sin Asignar"

    # Mapeo de Solped MRP
    if "Solped MRP" not in df.columns:
        df["Solped MRP"] = "No MRP"

    # Clasificación de Estado Solped
    if "Estado Solped" not in df.columns:
        tiene_pedido = df["Pedido"].notna() & (df["Pedido"] != "") & (df["Pedido"] != 0)
        df["Estado Solped"] = np.where(tiene_pedido, "Pedido completo", "Sin pedido")

    # Flag Aplica?
    if "Aplica?" not in df.columns:
        df["Aplica?"] = "Si"

    # Cálculo de Nivel de Servicio (Días de gestión)
    if "Nivel de Servicio" not in df.columns:
        if "Fecha de pedido" in df.columns and "Fecha de solicitud" in df.columns:
            df["Nivel de Servicio"] = (df["Fecha de pedido"] - df["Fecha de solicitud"]).dt.days
        else:
            df["Nivel de Servicio"] = np.nan

    # Evaluación de Cumple vs SLA
    df["Cumple"] = np.where(
        df["Nivel de Servicio"].isna(),
        "Sin Dato",
        np.where(df["Nivel de Servicio"] <= SLA_DIAS_ERP_MRP, "Cumple", "No Cumple")
    )

    return df


def calcular_metricas_por_grupo(df, columnas_grupo):
    """Calcula Promedio de días, % Cumplimiento y Posiciones de OC por agrupación."""
    if df.empty:
        columnas = columnas_grupo + ["Promedio días de gestión", "% Cumplimiento", "Pos. OC generadas"]
        return pd.DataFrame(columns=columnas)

    registros = []
    for grupo, sub_df in df.groupby(columnas_grupo, dropna=False):
        if not isinstance(grupo, tuple):
            grupo = (grupo,)
        
        fila = {col: val for col, val in zip(columnas_grupo, grupo)}
        total = len(sub_df)
        
        pct = (sub_df["Cumple"] == "Cumple").sum() / total * 100 if total > 0 else 0
        prom = sub_df["Nivel de Servicio"].mean() if total > 0 else np.nan
        pos_oc = sub_df["Pedido"].nunique() + (1 if sub_df["Pedido"].isna().any() else 0)

        fila["Promedio días de gestión"] = prom
        fila["% Cumplimiento"] = pct
        fila["Pos. OC generadas"] = pos_oc
        registros.append(fila)

    return pd.DataFrame(registros)


def agregar_fila_total(tabla, df_origen, columnas_grupo):
    """Agrega la fila 'TOTAL' recalculando indicadores sobre todo el subconjunto."""
    if tabla.empty or df_origen.empty:
        return tabla

    total_filas = len(df_origen)
    pct = (df_origen["Cumple"] == "Cumple").sum() / total_filas * 100 if total_filas > 0 else 0
    prom = df_origen["Nivel de Servicio"].mean() if total_filas > 0 else np.nan
    pos_oc = df_origen["Pedido"].nunique() + (1 if df_origen["Pedido"].isna().any() else 0)

    fila_total = {col: "" for col in columnas_grupo}
    fila_total[columnas_grupo[0]] = "TOTAL"
    fila_total["Promedio días de gestión"] = prom
    fila_total["% Cumplimiento"] = pct
    fila_total["Pos. OC generadas"] = pos_oc

    return pd.concat([tabla, pd.DataFrame([fila_total])], ignore_index=True)


def tabla_centros_fija(df):
    """Genera la vista resumen por centro logístico."""
    if "Centro" not in df.columns:
        return pd.DataFrame()
    
    tabla = calcular_metricas_por_grupo(df, ["Centro"])
    return agregar_fila_total(tabla, df, ["Centro"])


# ==============================================================================
# 4. INTERFAZ DE STREAMLIT
# ==============================================================================
st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

# Control de acceso con contraseña opcional
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("Dx Compradores — Nivel de Servicio")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password")
        if st.button("Ingresar"):
            if clave_ingresada == st.secrets["app_password"]:
                st.session_state["_autenticado"] = True
                _descargar_onedrive.clear()
                st.session_state.pop("_clave_pipeline", None)
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

st.title("Dx Compradores — Nivel de Servicio")

# Menú lateral para fuentes de datos
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

        if st.button("🔄 Forzar recarga desde OneDrive"):
            _descargar_onedrive.clear()
            st.session_state.pop("_clave_pipeline", None)
            st.rerun()

        if "onedrive" not in st.secrets:
            st.error("Falta configurar los Secrets de OneDrive en Streamlit Cloud.")
            st.stop()

    elif modo == "Subir archivos":
        archivo_data = st.file_uploader("ME5A_con_Ariba (.xlsx o .parquet)", type=["xlsx", "parquet"])
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
    fecha_corte = st.date_input("Fecha de corte del reporte", value=pd.Timestamp.today())
    st.caption(f"SLA: {SLA_DIAS_ERP_MRP} días ERP/MRP · {SLA_DIAS_ARIBA} días Ariba")


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

# Carga y pipeline en caché de sesión
if st.session_state.get("_clave_pipeline") != clave_actual:
    df_data = cargar_data_pr(archivo_data)
    df_resp_grupo = cargar_responsable_grupo_compras(archivo_resp_grupo)
    df_centro_sociedad = cargar_centro_sociedad_mro(archivo_centro)
    df_resp_mrp = cargar_responsable_mrp(archivo_mrp)

    df_calculado = pipeline_completo(
        df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=pd.Timestamp(fecha_corte)
    )
    df_calculado["Año"] = df_calculado["Fecha de pedido"].dt.year
    df_calculado["Mes"] = df_calculado["Fecha de pedido"].dt.month
    df_calculado["Día"] = df_calculado["Fecha de pedido"].dt.day

    st.session_state["_df_pipeline"] = df_calculado
    st.session_state["_clave_pipeline"] = clave_actual

df = st.session_state["_df_pipeline"]

NS_MIN_GLOBAL = int(df["Nivel de Servicio"].min()) if len(df) and pd.notna(df["Nivel de Servicio"].min()) else 0
NS_MAX_GLOBAL = int(df["Nivel de Servicio"].max()) if len(df) and pd.notna(df["Nivel de Servicio"].max()) else 100

# Sección de Filtros Dinámicos
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

# Filtro Estado Solped
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]

# Jerarquía Año / Mes / Día
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

# Solped MRP y Cumple
c3, c4 = st.columns(2)
with c3:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c4:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]

# Filtro por rango de días de gestión
st.divider()
st.subheader("Filtro por días de gestión (Nivel de Servicio)")

if len(df_f):
    n_negativos = (df_f["Nivel de Servicio"] < 0).sum()
    f1, f2 = st.columns([1, 2])
    with f1:
        excluir_negativos = st.checkbox(
            "Excluir días negativos",
            value=False,
            help=f"Hay {n_negativos:,} filas con días negativos en la selección actual.",
        )
    with f2:
        if excluir_negativos:
            rango_ns = (0, NS_MAX_GLOBAL)
            st.caption(f"Rango aplicado: 0 a {NS_MAX_GLOBAL:,} días (negativos excluidos)")
        elif NS_MIN_GLOBAL < NS_MAX_GLOBAL:
            rango_ns = st.slider(
                "Rango de días de gestión",
                min_value=NS_MIN_GLOBAL,
                max_value=NS_MAX_GLOBAL,
                value=(NS_MIN_GLOBAL, NS_MAX_GLOBAL),
            )
        else:
            rango_ns = (NS_MIN_GLOBAL, NS_MAX_GLOBAL)

    df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

# Tarjetas KPI Principales
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
        f"</div>"
    )


if pd.isna(promedio_dias):
    f_dias, b_dias, txt_dias = "rgba(64,75,85,0.07)", "rgba(64,75,85,0.35)", "-"
elif promedio_dias > 10:
    f_dias, b_dias, txt_dias = ROJO, ROJO_BORDE, f"{promedio_dias:.0f}"
else:
    f_dias, b_dias, txt_dias = VERDE, VERDE_BORDE, f"{promedio_dias:.0f}"

f_pct, b_pct = (VERDE, VERDE_BORDE) if pct_cumplimiento >= 85 else (ROJO, ROJO_BORDE)

t1, t2, t3 = st.columns(3)
with t1:
    st.markdown(tarjeta("Promedio días de gestión", txt_dias, f_dias, b_dias), unsafe_allow_html=True)
with t2:
    st.markdown(tarjeta("% Cumplimiento", f"{pct_cumplimiento:.0f}%", f_pct, b_pct), unsafe_allow_html=True)
with t3:
    st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

st.divider()


# Renderizador HTML para tablas con diseño corporativo
def tabla_enaex(tabla: pd.DataFrame, max_height: int | None = None, compacta: bool = False) -> str:
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

    pad = "4px 6px" if compacta else "7px 12px"
    pad_th = "5px 6px" if compacta else "9px 12px"
    fuente = "0.72rem" if compacta else "0.86rem"
    fuente_th = "0.66rem" if compacta else "0.82rem"

    filas = []
    for i, r in enumerate(tabla.itertuples(index=False)):
        es_total = str(r[0]) == "TOTAL"
        if es_total:
            estilo_fila = f"background:{ENAEX_GRIS};color:#fff;font-weight:700;border-top:2px solid {ENAEX_ROJO};"
        else:
            fondo = "#ffffff" if i % 2 == 0 else "#f4f5f7"
            estilo_fila = f"background:{fondo};color:{ENAEX_GRIS};"
        celdas = []
        for j, c in enumerate(cols):
            align = "left" if j == 0 else "right"
            celdas.append(
                f'<td style="padding:{pad};text-align:{align};'
                f'border-bottom:1px solid #e3e5e8;white-space:nowrap;">{_fmt(c, r[j])}</td>'
            )
        filas.append(f'<tr style="{estilo_fila}">{"".join(celdas)}</tr>')

    abrev = {"Promedio días de gestión": "Días gest.", "% Cumplimiento": "% Cumpl.", "Pos. OC generadas": "Pos. OC"}
    encabezados = "".join(
        f'<th style="padding:{pad_th};text-align:{"left" if j == 0 else "right"};'
        f'background:{ENAEX_GRIS};color:#fff;font-weight:600;font-size:{fuente_th};'
        f'letter-spacing:.02em;position:sticky;top:0;z-index:2;white-space:nowrap;">'
        f"{abrev.get(c, c) if compacta else c}</th>"
        for j, c in enumerate(cols)
    )

    tabla_html = (
        f'<table style="width:100%;min-width:{"340px" if compacta else "auto"};border-collapse:collapse;font-size:{fuente};'
        f'font-family:inherit;border:1px solid #d8dbdf;">'
        f"<thead><tr>{encabezados}</tr></thead><tbody>{''.join(filas)}</tbody></table>"
    )

    alto = f"max-height:{max_height}px;overflow-y:auto;" if max_height else ""
    return f'<div style="{alto}overflow-x:auto;border:1px solid #d8dbdf;border-radius:4px;">{tabla_html}</div>'


# Tabla por Comprador
st.subheader("Por comprador")
col_comprador = "Comprador (Grupo de compras)" if "Comprador (Grupo de compras)" in df_f.columns else "Comprador por Grupo Compras"
tabla_comprador = calcular_metricas_por_grupo(df_f, [col_comprador])
tabla_comprador = agregar_fila_total(tabla_comprador, df_f, [col_comprador])
st.markdown(tabla_enaex(tabla_comprador), unsafe_allow_html=True)

st.write("")

# Tablas por Centro Logístico
vc1, vc2 = st.columns(2)
with vc1:
    st.subheader("Por centro logístico")
    tabla_fija = tabla_centros_fija(df_f)
    st.markdown(tabla_enaex(tabla_fija, compacta=True), unsafe_allow_html=True)

with vc2:
    st.subheader("Detalle por centro")
    cols_detalle = [c for c in ["Centro", "Nombre Centro 2"] if c in df_f.columns]
    tabla_detalle = calcular_metricas_por_grupo(df_f, cols_detalle)
    tabla_detalle = tabla_detalle.sort_values("Pos. OC generadas", ascending=False)
    tabla_detalle = agregar_fila_total(tabla_detalle, df_f, cols_detalle)
    st.markdown(tabla_enaex(tabla_detalle, max_height=300, compacta=True), unsafe_allow_html=True)

st.divider()

# Editor interactivo de comentarios
COLUMNAS_DETALLE = [
    "Centro", "Material", "Texto breve", "Solicitud de pedido", "Fecha de solicitud",
    "Fecha modificación", "Grupo de compras", "Cantidad pedida", "Pedido", "Fecha de pedido",
    "Posición de pedido", "Comprador (Grupo de compras)", "Solped MRP", "Nombre Centro 2",
    "Nombre Centro", "Estado Solped", "Nivel de Servicio", "Comentario", "Cumple"
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
    detalle_editado = st.data_editor(
        detalle,
        use_container_width=True,
        num_rows="fixed",
        key="editor_detalle",
        column_config={
            "Comentario": st.column_config.TextColumn(
                "Comentario", help="Anotación libre para el registro semanal", width="medium"
            )
        },
        disabled=[c for c in detalle.columns if c != "Comentario"],
    )

# Generación del reporte Excel
semana_ref = pd.Timestamp(fecha_corte) - pd.Timedelta(days=7)
num_semana = semana_ref.isocalendar()[1]
nombre_archivo = f"Sem{num_semana:02d}-{semana_ref.year}.xlsx"


def generar_excel(detalle_df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    gris = "FF404B55"
    rojo = "FFCC0000"
    fuente_base = "Arial"

    wb = Workbook()
    borde = Border(bottom=Side(style="thin", color="FFD8DBDF"))

    def escribir_hoja(ws, titulo, tabla, col_inicio=1, fila_inicio=1):
        ws.cell(row=fila_inicio, column=col_inicio, value=titulo).font = Font(
            name=fuente_base, bold=True, size=12, color=gris
        )
        fila = fila_inicio + 1
        for j, col in enumerate(tabla.columns):
            c = ws.cell(row=fila, column=col_inicio + j, value=str(col))
            c.font = Font(name=fuente_base, bold=True, color="FFFFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=gris)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for i, r in enumerate(tabla.itertuples(index=False, name=None)):
            es_total = str(r[0]) == "TOTAL"
            for j, col in enumerate(tabla.columns):
                val = r[j]
                if pd.isna(val):
                    val = None
                elif isinstance(val, (int, float)) and col in ("Promedio días de gestión", "% Cumplimiento", "Pos. OC generadas"):
                    val = round(float(val))
                elif hasattr(val, "item"):
                    val = val.item()
                c = ws.cell(row=fila + 1 + i, column=col_inicio + j, value=val)
                c.font = Font(name=fuente_base, size=10, bold=es_total, color=gris)
                c.border = borde
                if es_total:
                    c.fill = PatternFill("solid", fgColor="FFEFF0F2")
                if col == "% Cumplimiento" and val is not None:
                    c.number_format = '0"%"'
                if col in ("Fecha de solicitud", "Fecha modificación", "Fecha de pedido"):
                    c.number_format = "DD-MM-YYYY"
        return fila + 1 + len(tabla)

    def ajustar_ancho(ws, tabla, col_inicio=1, extra=3):
        for j, col in enumerate(tabla.columns):
            largos = [len(str(col))]
            for v in tabla[col].head(200):
                largos.append(0 if pd.isna(v) else len(str(v)))
            ws.column_dimensions[get_column_letter(col_inicio + j)].width = min(max(largos) + extra, 45)

    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = f"Nivel de Servicio MRO — Registro semana {num_semana:02d}/{semana_ref.year}"
    ws["A1"].font = Font(name=fuente_base, bold=True, size=14, color=gris)
    ws["A2"] = f"Fecha de corte del reporte: {pd.Timestamp(fecha_corte).date()}"
    ws["A2"].font = Font(name=fuente_base, size=10, italic=True, color=gris)
    ws["A3"] = f"Generado: {pd.Timestamp.today().date()}"
    ws["A3"].font = Font(name=fuente_base, size=10, italic=True, color=gris)

    resumen = pd.DataFrame(
        {
            "Indicador": ["Promedio días de gestión", "% Cumplimiento", "OC generadas", "Líneas consideradas"],
            "Valor": [
                round(promedio_dias) if pd.notna(promedio_dias) else None,
                round(pct_cumplimiento),
                pedidos_distintos,
                len(df_f),
            ],
        }
    )
    fila = escribir_hoja(ws, "Indicadores generales", resumen, fila_inicio=5)
    ws.cell(row=5, column=1).font = Font(name=fuente_base, bold=True, size=12, color=rojo)
    ajustar_ancho(ws, resumen)

    fila = escribir_hoja(ws, "Por comprador", tabla_comprador, fila_inicio=fila + 2)
    fila = escribir_hoja(ws, "Por centro logístico", tabla_fija, fila_inicio=fila + 2)
    escribir_hoja(ws, "Detalle por centro", tabla_detalle, fila_inicio=fila + 2)

    ws2 = wb.create_sheet("Detalle solicitudes")
    escribir_hoja(ws2, "Detalle de solicitudes", detalle_df)
    ajustar_ancho(ws2, detalle_df)
    ws2.freeze_panes = "A3"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


st.subheader("Registro semanal")
st.caption(f"Prepara el reporte completo como **{nombre_archivo}**.")

clave_excel = (len(df_f), int(pd.util.hash_pandas_object(detalle_editado["Comentario"].fillna("")).sum()), pd.Timestamp(fecha_corte))

col_prep, col_desc = st.columns([1, 2])
with col_prep:
    if st.button("📄 Preparar Excel"):
        with st.spinner("Generando el Excel..."):
            try:
                st.session_state["_excel_bytes"] = generar_excel(detalle_editado)
                st.session_state["_excel_clave"] = clave_excel
            except Exception as e:
                st.error(f"No se pudo generar el Excel: {e}")

with col_desc:
    excel_listo = st.session_state.get("_excel_bytes") is not None
    excel_desactualizado = st.session_state.get("_excel_clave") != clave_excel
    if excel_listo and excel_desactualizado:
        st.caption("⚠️ Los filtros cambiaron — vuelve a preparar para reflejar la selección actual.")
    if excel_listo:
        st.download_button(
            f"⬇ Descargar {nombre_archivo}",
            data=st.session_state["_excel_bytes"],
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
