import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")
st.title("🎯 Dashboard OTIF (On Time, In Full)")
st.markdown("Medición del nivel de servicio de proveedores cruzando archivos de SAP.")

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
    # MODO CLARO: CSS ajustado para colocar el icono debajo de los 3 puntos y evitar que sea gigante
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important; /* Posicionado justo debajo de los 3 puntos */
            right: 15px !important; /* Esquina superior derecha */
            z-index: 999999 !important;
            width: 45px !important; /* Ancho fijo para evitar que sea gigante */
            height: 45px !important; /* Alto fijo */
            min-width: 0 !important; 
        }
        .st-key-theme_toggle button {
            background: #FFFFFF !important;
            border: 1px solid #E0E0E0 !important;
            border-radius: 50% !important; /* Botón circular */
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
        .st-key-theme_toggle button p {
            margin: 0 !important;
            padding: 0 !important;
            line-height: 1 !important;
            font-size: 1.4rem !important;
        }
        .st-key-theme_toggle button:hover {
            transform: scale(1.1) !important;
            background: #F0F0F0 !important;
        }
        </style>
    """, unsafe_allow_html=True)
else:
    # MODO OSCURO: CSS ajustado para colocar el icono debajo de los 3 puntos y evitar que sea gigante
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important; /* Posicionado justo debajo de los 3 puntos */
            right: 15px !important; /* Esquina superior derecha */
            z-index: 999999 !important;
            width: 45px !important; /* Ancho fijo para evitar que sea gigante */
            height: 45px !important; /* Alto fijo */
            min-width: 0 !important;
        }
        .st-key-theme_toggle button {
            background: #1E2329 !important;
            border: 1px solid #444444 !important;
            border-radius: 50% !important; /* Botón circular */
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
        .st-key-theme_toggle button p {
            margin: 0 !important;
            padding: 0 !important;
            line-height: 1 !important;
            font-size: 1.4rem !important;
        }
        .st-key-theme_toggle button:hover {
            transform: scale(1.1) !important;
            background: #2C323A !important;
            border-color: #FF3333 !important;
        }

        .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], html, body, [data-testid="stHeader"] {
            background-color: #0E1117 !important;
            color: #FF3333 !important;
        }

        p, span, label, h1, h2, h3, h4, h5, h6, div, td, th, caption, .stMarkdown {
            color: #FF3333 !important;
        }

        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button) {
            background-color: #CC0000 !important;
            color: #FFFFFF !important;
            border: 1px solid #FF4D4D !important;
            font-weight: bold !important;
        }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button):hover {
            background-color: #FF0000 !important;
            color: #FFFFFF !important;
            border-color: #FF6666 !important;
        }

        input, select, textarea, div[data-baseweb="select"] {
            background-color: #1E2329 !important;
            color: #FF3333 !important;
            border-color: #CC0000 !important;
        }

        [data-testid="stForm"], div[data-testid="stVerticalBlock"] > div:has(input[type="password"]) {
            background-color: #1E2329 !important;
            padding: 2rem !important;
            border-radius: 12px !important;
            border: 1px solid #CC0000 !important;
        }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# FUNCIÓN EN CACHÉ (Optimizada)
# ==========================================
@st.cache_data(show_spinner="Procesando cruce de datos SAP...")
def procesar_otif(file_me2m, file_me80fn):
    df_me2m = pd.read_excel(file_me2m, engine="openpyxl")
    cols_me80fn = ['Documento compras', 'Posición', 'Fe.contabilización']
    df_me80fn = pd.read_excel(file_me80fn, engine="openpyxl", usecols=lambda c: c in cols_me80fn)

    # 1. Descartar posiciones anuladas en SAP
    if 'Indicador de borrado' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Indicador de borrado'].isna()].copy()

    # 2. Última fecha de recepción real en ME80FN
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Fecha_Ingreso_SAP=('Fe.contabilización', 'max')
    ).reset_index()

    # 3. Cruzar ME2M con ME80FN
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

    # 4. Formato de fechas
    df_otif['Fecha_Estadistica'] = pd.to_datetime(df_otif['Fecha entrega estad.'], errors='coerce').dt.date
    df_otif['Fecha_Ingreso_SAP'] = pd.to_datetime(df_otif['Fecha_Ingreso_SAP'], errors='coerce').dt.date

    # 5. Cálculo de Reglas OTIF
    df_otif['In_Full'] = df_otif['Por entregar (cantidad)'] == 0
    df_otif['On_Time'] = (
        df_otif['Fecha_Ingreso_SAP'].notna() & 
        df_otif['Fecha_Estadistica'].notna() & 
        (df_otif['Fecha_Ingreso_SAP'] <= df_otif['Fecha_Estadistica'])
    )
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    # 6. Mapeo de estados visuales
    df_otif['Estado On Time'] = df_otif.apply(
        lambda r: '🔵 A Tiempo' if r['On_Time'] 
        else ('⏳ Pendiente' if pd.isna(r['Fecha_Ingreso_SAP']) else '🔴 Atrasado'), 
        axis=1
    )
    df_otif['Estado In Full'] = df_otif['In_Full'].map({True: '🔵 Completo', False: '🔴 Incompleto'})
    df_otif['Estado OTIF'] = df_otif['OTIF'].map({True: '🔵 Cumple OTIF', False: '🔴 No Cumple'})

    return df_otif

# ==========================================
# BARRA LATERAL (Carga)
# ==========================================
with st.sidebar:
    st.header("📂 Carga de Datos")
    archivo_me2m = st.file_uploader("1. Sube ME2M (.xlsx)", type=["xlsx"])
    archivo_me80fn = st.file_uploader("2. Sube ME80FN (.xlsx)", type=["xlsx"])

# ==========================================
# CUERPO PRINCIPAL
# ==========================================
if archivo_me2m and archivo_me80fn:
    df_base = procesar_otif(archivo_me2m, archivo_me80fn)

    # Filtros
    st.markdown("**🔍 Filtros de Análisis**")
    
    # Fila 1: Filtros Generales
    f_col1, f_col2, f_col3 = st.columns(3)
    with f_col1:
        centros = sorted(df_base['Centro'].dropna().unique())
        centro_sel = st.multiselect("Centro Logístico", centros)
    with f_col2:
        grupos = sorted(df_base['Grupo de compras'].dropna().unique())
        grupo_sel = st.multiselect("Grupo de Compras", grupos)
    with f_col3:
        proveedores = sorted(df_base['Proveedor/Centro suministrador'].dropna().astype(str).unique())
        prov_sel = st.multiselect("Proveedor", proveedores)

    # Fila 2: Filtros de Estados OTIF
    f_col4, f_col5, f_col6 = st.columns(3)
    with f_col4:
        estados_on_time = sorted(df_base['Estado On Time'].dropna().unique())
        on_time_sel = st.multiselect("Estado On Time", estados_on_time)
    with f_col5:
        estados_in_full = sorted(df_base['Estado In Full'].dropna().unique())
        in_full_sel = st.multiselect("Estado In Full", estados_in_full)
    with f_col6:
        estados_otif = sorted(df_base['Estado OTIF'].dropna().unique())
        otif_sel = st.multiselect("Estado OTIF", estados_otif)

    # Aplicar Filtros
    df = df_base.copy()
    if centro_sel:
        df = df[df['Centro'].isin(centro_sel)]
    if grupo_sel:
        df = df[df['Grupo de compras'].isin(grupo_sel)]
    if prov_sel:
        df = df[df['Proveedor/Centro suministrador'].astype(str).isin(prov_sel)]
    if on_time_sel:
        df = df[df['Estado On Time'].isin(on_time_sel)]
    if in_full_sel:
        df = df[df['Estado In Full'].isin(in_full_sel)]
    if otif_sel:
        df = df[df['Estado OTIF'].isin(otif_sel)]

    # KPIs
    total_lineas = len(df)
    pct_on_time = (df['On_Time'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_in_full = (df['In_Full'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_otif = (df['OTIF'].sum() / total_lineas * 100) if total_lineas > 0 else 0

    st.divider()

    st.markdown("**📊 Indicadores de Cumplimiento**")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Líneas Activas", f"{total_lineas:,}")
    k2.metric("⏱️ On Time", f"{pct_on_time:.1f}%")
    k3.metric("📦 In Full", f"{pct_in_full:.1f}%")
    k4.metric("⭐ OTIF Global", f"{pct_otif:.1f}%")

    st.divider()

    # Tabla de Detalle
    st.markdown("**📋 Detalle de Posiciones**")
    cols_mostrar = [
        'Documento compras', 'Posición', 'Centro', 'Proveedor/Centro suministrador',
        'Texto breve', 'Cantidad de pedido', 'Por entregar (cantidad)', 
        'Fecha_Estadistica', 'Fecha_Ingreso_SAP', 'Estado On Time', 'Estado In Full', 'Estado OTIF'
    ]
    df_mostrar = df[[c for c in cols_mostrar if c in df.columns]]

    # Configuración de columnas para ampliar ancho en la vista web
    config_columnas = {
        "Proveedor/Centro suministrador": st.column_config.TextColumn("Proveedor/Centro suministrador", width="large"),
        "Texto breve": st.column_config.TextColumn("Texto breve", width="large"),
        "Estado On Time": st.column_config.TextColumn("On Time", width="medium"),
        "Estado In Full": st.column_config.TextColumn("In Full", width="medium"),
        "Estado OTIF": st.column_config.TextColumn("OTIF", width="medium"),
    }

    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, column_config=config_columnas)

    # Descarga ligera y legible para Excel
    csv_bytes = df_mostrar.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

    st.download_button(
        label="📥 Descargar Reporte Completo (Formato Excel/CSV)",
        data=csv_bytes,
        file_name="Reporte_OTIF_SAP.csv",
        mime="text/csv"
    )
else:
    st.info("👈 Sube los archivos ME2M y ME80FN para desplegar las métricas.")
