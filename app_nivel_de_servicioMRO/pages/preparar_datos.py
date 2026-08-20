"""
Preparar datos — Conversión ME5A_con_Ariba.xlsx -> .parquet

Por qué esta página existe: el ME5A_con_Ariba.xlsx es, con diferencia, el
archivo más pesado de los 4 que carga el reporte. Leer un .xlsx con pandas
(vía openpyxl) tiene un pico de memoria varias veces mayor que el tamaño
final del DataFrame — el motor arma un árbol de objetos XML completo antes
de convertir a tabla. Parquet es binario y columnar: se lee casi directo a
memoria, con una fracción del pico de RAM y en menos tiempo.

Uso: sube acá tu ME5A_con_Ariba.xlsx una vez, descarga el .parquet resultante,
y desde ahora súbelo a la app principal en vez del .xlsx — el pipeline y los
filtros no cambian en nada, solo la fuente de carga.
"""
import io

import pandas as pd
import streamlit as st

import loaders

import streamlit as st

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

# ==============================================================================
# OCULTAR NAVEGACIÓN GLOBAL (TRAZABILIDAD ARIBA)
# ==============================================================================
st.markdown("""
    <style>
    /* Ocultar enlaces del menú lateral que contengan 'trazabilidad' o 'ariba' */
    [data-testid="stSidebarNav"] a[href*="trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="Trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="ariba"],
    [data-testid="stSidebarNav"] a[href*="Ariba"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

st.set_page_config(page_title="Preparar datos - Nivel de Servicio", layout="wide")

st.title("Preparar datos — Conversión a Parquet")
st.caption(
    "Convierte tu ME5A_con_Ariba.xlsx a .parquet: mismo contenido, mismo tipado, "
    "pero con una carga mucho más rápida y liviana en memoria dentro de la app principal."
)

archivo = st.file_uploader("ME5A_con_Ariba.xlsx", type="xlsx")

if archivo is None:
    st.info("Sube el archivo ME5A_con_Ariba.xlsx para convertirlo.")
    st.stop()

with st.spinner("Leyendo y tipando el Excel..."):
    df_crudo = pd.read_excel(archivo, sheet_name="Data")
    df_tipado = loaders._tipar_data_pr(df_crudo)

st.success(f"Leído correctamente: {len(df_tipado):,} filas, {len(df_tipado.columns)} columnas.")

with st.expander("Ver una muestra de los datos tipados"):
    st.dataframe(df_tipado.head(50), use_container_width=True)

# ---- Comparación de tamaño/memoria: xlsx original vs parquet resultante ----
buffer_parquet = io.BytesIO()
df_tipado.to_parquet(buffer_parquet, index=False)
bytes_parquet = buffer_parquet.getvalue()

tam_xlsx_mb = archivo.size / 1024 / 1024
tam_parquet_mb = len(bytes_parquet) / 1024 / 1024

c1, c2 = st.columns(2)
with c1:
    st.metric("Tamaño .xlsx original", f"{tam_xlsx_mb:.1f} MB")
with c2:
    st.metric(
        "Tamaño .parquet resultante",
        f"{tam_parquet_mb:.1f} MB",
        delta=f"{tam_parquet_mb - tam_xlsx_mb:.1f} MB",
        delta_color="inverse",
    )

st.download_button(
    "⬇ Descargar ME5A_con_Ariba.parquet",
    data=bytes_parquet,
    file_name="ME5A_con_Ariba.parquet",
    mime="application/octet-stream",
)

st.caption(
    "Una vez descargado, ve a la pestaña del reporte principal y súbelo ahí en vez "
    "del .xlsx — el resto del flujo (filtros, tablas, exportación) funciona exactamente igual."
)
