"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores" y Trazabilidad.

Correr local: streamlit run app.py
"""
import pandas as pd
import streamlit as st

import config
import loaders
import transform

# Importación del archivo ariba_trazabilidad.py con alias para mantener compatibilidad
try:
    import ariba_trazabilidad as trazabilidad
    HAS_TRAZABILIDAD = True
except ImportError:
    HAS_TRAZABILIDAD = False

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

# ==============================================================================
# CONFIGURACIÓN TEMA (CLARO PREDETERMINADO / OSCURO)
# ==============================================================================
if "tema" not in st.session_state:
    st.session_state["tema"] = "claro"

# Botón flotante para cambiar tema
icono_tema = "🌙" if st.session_state["tema"] == "claro" else "☀️"
if st.button(icono_tema, key="theme_toggle", help="Alternar Modo Claro/Oscuro"):
    st.session_state["tema"] = "oscuro" if st.session_state["tema"] == "claro" else "claro"
    st.rerun()

if st.session_state["tema"] == "claro":
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important;
            right: 15px !important;
            z-index: 999999 !important;
            width: 45px !important;
            height: 45px !important;
            min-width: 0 !important; 
        }
        .st-key-theme_toggle button {
            background: #FFFFFF !important;
            border: 1px solid #E0E0E0 !important;
            border-radius: 50% !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1) !important;
            font-size: 1.4rem !important;
            padding: 0 !important;
            margin: 0 !important;
            color: #111111 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            line-height: 1 !important;
            width: 100% !important;
            height: 100% !important;
            min-height: unset !important;
        }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #F0F0F0 !important; }
        </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important;
            right: 15px !important;
            z-index: 999999 !important;
            width: 45px !important;
            height: 45px !important;
            min-width: 0 !important;
        }
        .st-key-theme_toggle button {
            background: #1E2329 !important;
            border: 1px solid #444444 !important;
            border-radius: 50% !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important;
            font-size: 1.4rem !important;
            padding: 0 !important;
            margin: 0 !important;
            color: #FF3333 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            line-height: 1 !important;
            width: 100% !important;
            height: 100% !important;
            min-height: unset !important;
        }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #2C323A !important; border-color: #FF3333 !important; }
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], html, body, [data-testid="stHeader"] { background-color: #0E1117 !important; color: #FF3333 !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, div, td, th, caption, .stMarkdown { color: #FF3333 !important; }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button) { background-color: #CC0000 !important; color: #FFFFFF !important; border: 1px solid #FF4D4D !important; font-weight: bold !important; }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button):hover { background-color: #FF0000 !important; color: #FFFFFF !important; border-color: #FF6666 !important; }
        input, select, textarea, div[data-baseweb="select"] { background-color: #1E2329 !important; color: #FF3333 !important; border-color: #CC0000 !important; }
        [data-testid="stForm"], div[data-testid="stVerticalBlock"] > div:has(input[type="password"]) { background-color: #1E2329 !important; padding: 2rem !important; border-radius: 12px !important; border: 1px solid #CC0000 !important; }
        </style>
    """, unsafe_allow_html=True)


# ---- 0. Función de Clasificación Corregida ----
def determinar_tipo_ariba(row):
    sol = str(row.get("Solicitud de pedido", "")).strip()
    material = str(row.get("Material", "")).strip()
    tiene_material = bool(material and material.lower() not in ["nan", "none", "n/a", "-", "0", "null"])
    es_mrp_flag = str(row.get("Solped MRP", "")).strip().lower() in ["sí", "si", "true", "mrp", "1"]
    en_trazabilidad = bool(row.get("En_Trazabilidad", False) or row.get("En Trazabilidad", False))

    texto_origen = ""
    for col in ["Tipo_Ariba", "Origen Ariba", "Origen", "Tipo Flujo", "Tipo Pedido"]:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != "":
            texto_origen += " " + str(row[col]).upper()

    if sol.startswith("5") or es_mrp_flag or "MRP" in texto_origen:
        return "⚪ SAP MRP"

    if sol.startswith("1") or sol.startswith("19") or sol.upper().startswith("CL") or "ERP" in texto_origen:
        return "⚙️ SAP ERP"

    if sol.startswith("6"):
        if "NO CATALOGAD" in texto_origen or "NOCATALOGAD" in texto_origen or "SIN CODIGO" in texto_origen:
            return "🔵 ARIBA NO CATALOGADA"
        elif en_trazabilidad or not tiene_material:
            return "🔵 ARIBA NO CATALOGADA"
        else:
            return "🟢 ARIBA CATALOGADA / DIRECTA"

    if "NO CATALOGAD" in texto_origen or "NOCATALOGAD" in texto_origen:
        return "🔵 ARIBA NO CATALOGADA"
    elif "DIRECTA" in texto_origen or "CATALOGAD" in texto_origen:
        return "🟢 ARIBA CATALOGADA / DIRECTA"

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

# ---- Pestañas Principales (Trazabilidad reemplazada visualmente por Preparar datos) ----
tab_dx, tab_preparar = st.tabs(["📊 Dx Compradores", "⚙️ Preparar datos"])

# ==============================================================================
# PESTAÑA 1: DX COMPRADORES
# ==============================================================================
with tab_dx:
    st.title("Dx Compradores — Nivel de Servicio")

    # ---- Fuente de datos ----
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
                st.error("Falta configurar los Secrets de OneDrive.")
                st.stop()

        elif modo == "Subir archivos":
            archivo_data = st.file_uploader(
                "ME5A_con_Ariba (.xlsx o .parquet)",
                type=["xlsx", "parquet"],
                help="Sube aquí tu archivo convertido a .parquet.",
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
        "df_trazabilidad_limpio" in st.session_state,
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

        if "df_trazabilidad_limpio" in st.session_state:
            df_traz = st.session_state["df_trazabilidad_limpio"]
            col_traz = "Solicitud de pedido" if "Solicitud de pedido" in df_traz.columns else "Solped SAP (600)"
            solpeds_no_cat = set(df_traz[col_traz].dropna().astype(str).str.strip().str.replace(r"\.0$", "", regex=True))
            solpeds_me5a = df_calculado["Solicitud de pedido"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            df_calculado["En_Trazabilidad"] = solpeds_me5a.isin(solpeds_no_cat)
        else:
            df_calculado["En_Trazabilidad"] = False

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
        metricas_por_etapa.append((nombre, f"{n:,}", f"{pct:.0f}%", f"{prom:.0f}" if pd.notna(prom) else "-", f"{pos_oc:,}"))

    _snapshot("0. Sin filtros", df)

    # ---- Filtros ----
    st.subheader("Filtros")
    c1, c2, c3 = st.columns(3)
    with c1: centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
    with c2: aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))
    with c3: tipos_ariba = st.multiselect("Origen / Tipo Solicitud", sorted(df["Tipo Ariba"].dropna().unique()))

    df_f = df.copy()
    if centros: df_f = df_f[df_f["Centro"].isin(centros)]
    checkpoints.append(("1. Tras filtro Centro", len(df_f)))

    if aplica: df_f = df_f[df_f["Aplica?"].isin(aplica)]
    checkpoints.append(("2. Tras filtro Aplica?", len(df_f)))

    if tipos_ariba: df_f = df_f[df_f["Tipo Ariba"].isin(tipos_ariba)]
    checkpoints.append(("2b. Tras filtro Origen / Tipo Solicitud", len(df_f)))
    _snapshot("2. Tras Centro + Aplica? + Origen", df_f)

    # ---- Estado Solped ----
    st.caption("Estado Solped")
    h1, h2, h3, h4 = st.columns(4)
    with h1: estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

    if estados: df_f = df_f[df_f["Estado Solped"].isin(estados)]
    _snapshot("3. Tras Estado Solped", df_f)

    df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]
    with h2: años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
    with h3:
        _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
        meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
    with h4:
        _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
        fechas = st.multiselect("Día", sorted(_base_dia["Fecha de pedido"].dt.date.dropna().unique()), format_func=lambda f: f.strftime("%d-%m-%Y"))

    if años or meses or fechas:
        es_pedido_completo = df_f["Estado Solped"] == "Pedido completo"
        cond_fecha = pd.Series(True, index=df_f.index)
        if años: cond_fecha &= df_f["Año"].isin(años)
        if meses: cond_fecha &= df_f["Mes"].isin(meses)
        if fechas: cond_fecha &= df_f["Fecha de pedido"].dt.date.isin(fechas)
        df_f = df_f[~es_pedido_completo | (es_pedido_completo & cond_fecha)]

    _snapshot("4. Tras jerarquía Año/Mes/Día", df_f)

    c4, c5 = st.columns(2)
    with c4: solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
    with c5: cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

    if solped_mrp: df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
    if cumple: df_f = df_f[df_f["Cumple"].isin(cumple)]

    st.divider()
    st.subheader("Filtro por días de gestión (Nivel de Servicio)")

    if len(df_f):
        f1, f2 = st.columns([1, 2])
        with f1: excluir_negativos = st.checkbox("Excluir negativos (solo desde 0)", value=False)
        with f2:
            if excluir_negativos:
                rango_ns = (0, NS_MAX_GLOBAL)
            elif NS_MIN_GLOBAL < NS_MAX_GLOBAL:
                rango_ns = st.slider("Rango de días de gestión", min_value=NS_MIN_GLOBAL, max_value=NS_MAX_GLOBAL, value=(NS_MIN_GLOBAL, NS_MAX_GLOBAL), key="slider_rango_dias")
            else:
                rango_ns = (NS_MIN_GLOBAL, NS_MAX_GLOBAL)
        df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

    # ---- Tarjetas KPI ----
    VERDE, VERDE_BORDE = "rgba(35, 145, 75, 0.16)", "rgba(35, 145, 75, 0.55)"
    ROJO, ROJO_BORDE = "rgba(204, 0, 0, 0.14)", "rgba(204, 0, 0, 0.55)"

    pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
    promedio_dias = df_f["Nivel de Servicio"].mean()
    promedio_lead_time = df_f["Lead Time Total"].mean() if "Lead Time Total" in df_f.columns else float("nan")
    pedidos_distintos = df_f["Pedido"].nunique() + (1 if df_f["Pedido"].isna().any() else 0)

    color_texto_card = "#FF3333" if st.session_state["tema"] == "oscuro" else "#404B55"
    color_sub_card = "#FF3333" if st.session_state["tema"] == "oscuro" else "#555"

    def tarjeta(titulo, valor, subtitulo="", fondo="rgba(64,75,85,0.07)", borde="rgba(64,75,85,0.35)"):
        html_sub = f'<div style="font-size:0.78rem;color:{color_sub_card};margin-top:4px;font-weight:500;">{subtitulo}</div>' if subtitulo else ""
        return (f'<div style="background:{fondo};border:1.5px solid {borde};border-radius:8px;padding:12px 18px;text-align:center;">'
                f'<div style="font-size:0.78rem;color:{color_texto_card};font-weight:600;letter-spacing:.03em;text-transform:uppercase;opacity:.85;margin-bottom:4px;">{titulo}</div>'
                f'<div style="font-size:2rem;font-weight:700;color:{color_texto_card};line-height:1.1;">{valor}</div>{html_sub}</div>')

    f_dias, b_dias, txt_dias = ("rgba(64,75,85,0.07)", "rgba(64,75,85,0.35)", "-") if pd.isna(promedio_dias) else (ROJO, ROJO_BORDE, f"{promedio_dias:.0f}") if promedio_dias > 10 else (VERDE, VERDE_BORDE, f"{promedio_dias:.0f}")
    f_pct, b_pct = (VERDE, VERDE_BORDE) if pct_cumplimiento >= 85 else (ROJO, ROJO_BORDE)
    txt_lt = f"{promedio_lead_time:.0f}" if pd.notna(promedio_lead_time) else "-"

    t1, t2, t3 = st.columns(3)
    with t1: st.markdown(tarjeta("Nivel de Servicio", f"{txt_dias} días", subtitulo=f"Lead Time Total: <b>{txt_lt}</b> días", fondo=f_dias, borde=b_dias), unsafe_allow_html=True)
    with t2: st.markdown(tarjeta("% Cumplimiento SLA", f"{pct_cumplimiento:.0f}%", fondo=f_pct, borde=b_pct), unsafe_allow_html=True)
    with t3: st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

    st.divider()

    # ---- Tablas Enaex ----
    ENAEX_GRIS = "#1E2329" if st.session_state["tema"] == "oscuro" else "#404B55"
    ENAEX_ROJO = "#CC0000"

    def tabla_enaex(tabla, max_height=None, compacta=False):
        cols = list(tabla.columns)
        def _fmt(col, val):
            if pd.isna(val): return "-"
            if col in ["Promedio días de gestión", "Promedio Lead Time Total"]: return f"{val:,.0f}"
            if col == "% Cumplimiento": return f"{val:,.0f}%"
            if col == "Pos. OC generadas": return f"{val:,.0f}"
            return str(val)

        pad = "4px 6px" if compacta else "7px 12px"
        pad_th = "5px 6px" if compacta else "9px 12px"
        fuente = "0.72rem" if compacta else "0.86rem"
        fuente_th = "0.66rem" if compacta else "0.82rem"
        color_texto_tabla = "#FF3333" if st.session_state["tema"] == "oscuro" else "#404B55"

        filas = []
        for i, r in enumerate(tabla.itertuples(index=False)):
            es_total = str(r[0]) == "TOTAL"
            estilo_fila = f"background:{ENAEX_GRIS};color:#fff;font-weight:700;border-top:2px solid {ENAEX_ROJO};" if es_total else f"background:{'#ffffff' if (i % 2 == 0 and st.session_state['tema'] == 'claro') else ('#f4f5f7' if st.session_state['tema'] == 'claro' else '#14181d')};color:{color_texto_tabla};"
            celdas = "".join(f'<td style="padding:{pad};text-align:{"left" if j==0 else "right"};border-bottom:1px solid #d8dbdf;white-space:nowrap;">{_fmt(c, r[j])}</td>' for j, c in enumerate(cols))
            filas.append(f'<tr style="{estilo_fila}">{celdas}</tr>')

        abrev = {"Promedio días de gestión": "Nivel Serv.", "Promedio Lead Time Total": "LT Total", "% Cumplimiento": "% Cumpl.", "Pos. OC generadas": "Pos. OC"}
        encabezados = "".join(f'<th style="padding:{pad_th};text-align:{"left" if j == 0 else "right"};background:{ENAEX_GRIS};color:#fff;font-weight:600;font-size:{fuente_th};letter-spacing:.02em;position:sticky;top:0;z-index:2;white-space:nowrap;">{abrev.get(c, c) if compacta else c}</th>' for j, c in enumerate(cols))
        
        tabla_html = f'<table style="width:100%;min-width:{"340px" if compacta else "auto"};border-collapse:collapse;font-size:{fuente};font-family:inherit;border:1px solid #d8dbdf;"><thead><tr>{encabezados}</tr></thead><tbody>{"".join(filas)}</tbody></table>'
        return f'<div style="{f"max-height:{max_height}px;overflow-y:auto;" if max_height else ""}overflow-x:auto;border:1px solid #d8dbdf;border-radius:4px;">{tabla_html}</div>'

    st.subheader("Por comprador")
    col_comprador = "Comprador (Grupo de compras)" if "Comprador (Grupo de compras)" in df_f.columns else "Comprador por Grupo Compras"
    tabla_comprador = transform.calcular_metricas_por_grupo(df_f, [col_comprador])
    tabla_comprador = transform.agregar_fila_total(tabla_comprador, df_f, [col_comprador])
    st.markdown(tabla_enaex(tabla_comprador), unsafe_allow_html=True)

    vc1, vc2 = st.columns(2)
    with vc1:
        st.subheader("Por centro logístico")
        tabla_fija = transform.tabla_centros_fija(df_f)
        st.markdown(tabla_enaex(tabla_fija, compacta=True), unsafe_allow_html=True)
    with vc2:
        st.subheader("Detalle por centro")
        cols_detalle = [c for c in ["Centro", "Nombre Centro 2"] if c in df_f.columns]
        tabla_detalle = transform.calcular_metricas_por_grupo(df_f, cols_detalle).sort_values("Pos. OC generadas", ascending=False)
        tabla_detalle = transform.agregar_fila_total(tabla_detalle, df_f, cols_detalle)
        st.markdown(tabla_enaex(tabla_detalle, max_height=300, compacta=True), unsafe_allow_html=True)


# ==============================================================================
# PESTAÑA 2: PREPARAR DATOS (Restaurada)
# ==============================================================================
with tab_preparar:
    st.title("⚙️ Preparar datos (Excel a Parquet)")
    st.markdown("Sube tu archivo Excel pesado para convertirlo a formato **.parquet**. Esto reduce drásticamente el peso del archivo y acelera la carga de la aplicación.")
    
    archivo_a_convertir = st.file_uploader("Cargar archivo Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    
    if archivo_a_convertir:
        if st.button("🚀 Convertir archivo"):
            with st.spinner("Leyendo el archivo Excel (esto puede tomar un momento)..."):
                try:
                    df_temp = pd.read_excel(archivo_a_convertir)
                    parquet_buffer = df_temp.to_parquet(index=False)
                    st.success("¡Archivo convertido con éxito!")
                    st.download_button(
                        label="⬇️ Descargar archivo .parquet",
                        data=parquet_buffer,
                        file_name=archivo_a_convertir.name.replace(".xlsx", ".parquet").replace(".xls", ".parquet"),
                        mime="application/octet-stream"
                    )
                except Exception as e:
                    st.error(f"Ocurrió un error durante la conversión: {e}")


# ==============================================================================
# SECCIÓN OCULTA: TRAZABILIDAD NO CATALOGADAS
# (Oculta visualmente mediante un bloque "if False", pero el código se mantiene)
# ==============================================================================
if False: 
    st.title("🔍 Trazabilidad PR No Catalogadas — Ariba")

    if not HAS_TRAZABILIDAD:
        st.warning("⚠️ **Módulo 'ariba_trazabilidad.py' no disponible.**")
    else:
        st.markdown("Procesa el reporte de **Trazabilidad Ariba** sin consolidar...")

        traz_col1, traz_col2 = st.columns([2, 1])
        with traz_col1:
            archivo_trazabilidad = st.file_uploader("Cargar Reporte PR No Catalogadas - Trazabilidad (.csv)", type=["csv"], key="uploader_trazabilidad")
        with traz_col2:
            empresa_id = st.text_input("Empresa compradora (ID)", value=getattr(trazabilidad, "EMPRESA_POR_DEFECTO", "1000"))

        if archivo_trazabilidad:
            if st.button("🚀 Procesar Trazabilidad", key="btn_procesar_traz"):
                with st.spinner("Procesando trazabilidad y reconstruyendo cadena de eventos..."):
                    try:
                        df_cadena, df_resumen = trazabilidad.procesar_trazabilidad_completa(archivo_trazabilidad, empresa=empresa_id)
                        st.session_state["df_trazabilidad_cadena"] = df_cadena
                        st.session_state["df_trazabilidad_limpio"] = df_resumen
                        st.session_state.pop("_clave_pipeline", None)
                        st.success("¡Trazabilidad procesada con éxito! La pestaña Dx Compradores ha sido actualizada.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al procesar el archivo: {e}")
