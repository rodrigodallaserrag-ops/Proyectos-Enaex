"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores" y Trazabilidad.

Correr local: streamlit run app.py
"""
import glob
import os
import re
import pandas as pd
import numpy as np
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
# CONFIGURACIÓN DE ENLACES SHAREPOINT / ONEDRIVE (URLs de descarga directa)
# ==============================================================================
URLS_SHAREPOINT = {
    "me5a": "https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQDXH7sdl9P7SJsmurEVREx8AdaUM4nE7AlJilbaTUTcaBQ?download=1",
    "responsable_grupo": "https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQD2_Sy3u0zVQafFMc9QybdlAamhW7o9erDNUwXCOOIa7v0?download=1",
    "responsable_mrp": "https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQAH7p88424LRpAkG9z6okZkASi1JfubHzWkgDpIIsTGqHg?download=1",
    "centro_sociedad": "https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQCDkE0EuXfKQZcOLPJ6o5mfAf1vRDJ4MIq2GnVqyuFwaGI?download=1",
}

# ==============================================================================
# FUNCIONES AUXILIARES PARA MANEJO DE ARCHIVOS DUPLICADOS
# ==============================================================================
def buscar_archivo_mas_reciente(patron_o_ruta: str) -> str:
    """
    Busca archivos que coincidan con un patrón (ej: 'data/ME5A_con_Ariba*.parquet'
    o 'data/ME5A_con_Ariba (1).xlsx') y retorna la ruta del más recientemente modificado.
    """
    if not isinstance(patron_o_ruta, str) or patron_o_ruta.startswith("onedrive:") or patron_o_ruta.startswith("http"):
        return patron_o_ruta

    nombre_base, ext = os.path.splitext(patron_o_ruta)
    patron_busqueda = f"{nombre_base}*{ext}"
    
    coincidencias = glob.glob(patron_busqueda)
    if coincidencias:
        return max(coincidencias, key=os.path.getmtime)
    
    return patron_o_ruta

def obtener_ultimo_subido(archivos):
    if isinstance(archivos, list):
        return archivos[-1] if len(archivos) > 0 else None
    return archivos

# ==============================================================================
# OCULTAR NAVEGACIÓN GLOBAL (TRAZABILIDAD ARIBA)
# ==============================================================================
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] a[href*="trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="Trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="ariba"],
    [data-testid="stSidebarNav"] a[href*="Ariba"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONFIGURACIÓN TEMA (CLARO PREDETERMINADO / OSCURO)
# ==============================================================================
if "tema" not in st.session_state:
    st.session_state["tema"] = "claro"

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


# ---- 0. Función de Clasificación Corregida ----
def determinar_tipo_ariba(row):
    sol = str(row.get("Solicitud de pedido", "")).strip()
    material = str(row.get("Material", "")).strip()
    tiene_material = bool(material and material.lower() not in ["nan", "none", "n/a", "-", "0", "null"])
    es_mrp_flag = str(row.get("Solped MRP", "")).strip().lower() in ["sí", "si", "true", "mrpEl ajuste reemplaza el límite inferior del filtro de días de gestión. En lugar de iniciar desde 0 como estaba configurado originalmente[cite: 1], el checkbox ahora establece el corte mínimo estrictamente desde el día 1, descartando valores en 0 o negativos para alinearse con la nueva regla de medición.

```python
"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores" y Trazabilidad.

Correr local: streamlit run app.py
"""
import glob
import os
import re
import pandas as pd
import numpy as np
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
# CONFIGURACIÓN DE ENLACES SHAREPOINT / ONEDRIVE (URLs de descarga directa)
# ==============================================================================
URLS_SHAREPOINT = {
    "me5a": "[https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQDXH7sdl9P7SJsmurEVREx8AdaUM4nE7AlJilbaTUTcaBQ?download=1](https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQDXH7sdl9P7SJsmurEVREx8AdaUM4nE7AlJilbaTUTcaBQ?download=1)",
    "responsable_grupo": "[https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQD2_Sy3u0zVQafFMc9QybdlAamhW7o9erDNUwXCOOIa7v0?download=1](https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQD2_Sy3u0zVQafFMc9QybdlAamhW7o9erDNUwXCOOIa7v0?download=1)",
    "responsable_mrp": "[https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQAH7p88424LRpAkG9z6okZkASi1JfubHzWkgDpIIsTGqHg?download=1](https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQAH7p88424LRpAkG9z6okZkASi1JfubHzWkgDpIIsTGqHg?download=1)",
    "centro_sociedad": "[https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQCDkE0EuXfKQZcOLPJ6o5mfAf1vRDJ4MIq2GnVqyuFwaGI?download=1](https://empresassk-my.sharepoint.com/:x:/g/personal/cristian_vasquez_enaex_com/IQCDkE0EuXfKQZcOLPJ6o5mfAf1vRDJ4MIq2GnVqyuFwaGI?download=1)",
}

# ==============================================================================
# FUNCIONES AUXILIARES PARA MANEJO DE ARCHIVOS DUPLICADOS
# ==============================================================================
def buscar_archivo_mas_reciente(patron_o_ruta: str) -> str:
    """
    Busca archivos que coincidan con un patrón (ej: 'data/ME5A_con_Ariba*.parquet'
    o 'data/ME5A_con_Ariba (1).xlsx') y retorna la ruta del más recientemente modificado.
    """
    if not isinstance(patron_o_ruta, str) or patron_o_ruta.startswith("onedrive:") or patron_o_ruta.startswith("http"):
        return patron_o_ruta

    nombre_base, ext = os.path.splitext(patron_o_ruta)
    patron_busqueda = f"{nombre_base}*{ext}"
    
    coincidencias = glob.glob(patron_busqueda)
    if coincidencias:
        return max(coincidencias, key=os.path.getmtime)
    
    return patron_o_ruta

def obtener_ultimo_subido(archivos):
    if isinstance(archivos, list):
        return archivos[-1] if len(archivos) > 0 else None
    return archivos

# ==============================================================================
# OCULTAR NAVEGACIÓN GLOBAL (TRAZABILIDAD ARIBA)
# ==============================================================================
st.markdown("""
    <style>
    [data-testid="stSidebarNav"] a[href*="trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="Trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="ariba"],
    [data-testid="stSidebarNav"] a[href*="Ariba"] {
        display: none !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONFIGURACIÓN TEMA (CLARO PREDETERMINADO / OSCURO)
# ==============================================================================
if "tema" not in st.session_state:
    st.session_state["tema"] = "claro"

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

# ---- Pestañas Principales ----
tab_dx, tab_trazabilidad = st.tabs(["📊 Dx Compradores", "🔗 Trazabilidad No Catalogadas"])

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
            archivo_data = URLS_SHAREPOINT["me5a"]
            archivo_resp_grupo = URLS_SHAREPOINT["responsable_grupo"]
            archivo_centro = URLS_SHAREPOINT["centro_sociedad"]
            archivo_mrp = URLS_SHAREPOINT["responsable_mrp"]

            if st.button("🔄 Forzar recarga desde OneDrive ahora"):
                loaders._descargar_onedrive.clear()
                st.session_state.pop("_clave_pipeline", None)
                st.rerun()

        elif modo == "Subir archivos":
            files_data = st.file_uploader(
                "ME5A_con_Ariba (.xlsx o .parquet)",
                type=["xlsx", "parquet"],
                accept_multiple_files=True,
                help="Puedes subir uno o varios archivos (incluso duplicados con '(1)'). Se tomará el último.",
            )
            files_resp = st.file_uploader("Responsable_Grupo_Compras.xlsx", type="xlsx", accept_multiple_files=True)
            files_centro = st.file_uploader("Centro_Sociedad_MRO.xlsx", type="xlsx", accept_multiple_files=True)
            files_mrp = st.file_uploader("Responsable_MRP.xlsx", type="xlsx", accept_multiple_files=True)

            archivo_data = obtener_ultimo_subido(files_data)
            archivo_resp_grupo = obtener_ultimo_subido(files_resp)
            archivo_centro = obtener_ultimo_subido(files_centro)
            archivo_mrp = obtener_ultimo_subido(files_mrp)

            if not all([archivo_data, archivo_resp_grupo, archivo_centro, archivo_mrp]):
                st.info("Sube los 4 archivos para generar el reporte.")
                st.stop()
        else:
            archivo_parquet_local = buscar_archivo_mas_reciente("data/ME5A_con_Ariba.parquet")
            archivo_excel_local = buscar_archivo_mas_reciente("data/ME5A_con_Ariba.xlsx")

            if os.path.exists(archivo_parquet_local):
                archivo_data = archivo_parquet_local
            else:
                archivo_data = archivo_excel_local

            archivo_resp_grupo = buscar_archivo_mas_reciente("data/Responsable_Grupo_Compras.xlsx")
            archivo_centro = buscar_archivo_mas_reciente("data/Centro_Sociedad_MRO.xlsx")
            archivo_mrp = buscar_archivo_mas_reciente("data/Responsable_MRP.xlsx")

        st.header("Parámetros")
        fecha_corte = st.date_input("Fecha de corte del reporte (FechaCorteReporte)", value=pd.Timestamp.today())
        st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")

    def _clave_archivo(archivo):
        if isinstance(archivo, list):
            return tuple(_clave_archivo(a) for a in archivo)
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
        try:
            df_data = loaders.cargar_data_pr(archivo_data)
            df_resp_grupo = loaders.cargar_responsable_grupo_compras(archivo_resp_grupo)
            df_centro_sociedad = loaders.cargar_centro_sociedad_mro(archivo_centro)
            df_resp_mrp = loaders.cargar_responsable_mrp(archivo_mrp)
            
            df_calculado = transform.pipeline_completo(
                df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp,
                fecha_corte=pd.Timestamp(fecha_corte),
                df_trazabilidad=st.session_state.get("df_trazabilidad_limpio"),
            )
            df_calculado["Año"] = df_calculado["Fecha de pedido"].dt.year
            df_calculado["Mes"] = df_calculado["Fecha de pedido"].dt.month
            df_calculado["Día"] = df_calculado["Fecha de pedido"].dt.day

            if "df_trazabilidad_limpio" in st.session_state:
                df_traz = st.session_state["df_trazabilidad_limpio"]
                col_traz = "Solicitud de pedido" if "Solicitud de pedido" in df_traz.columns else "Solped SAP (600)"
                
                solpeds_no_cat = set(
                    df_traz[col_traz]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                )

                solpeds_me5a = (
                    df_calculado["Solicitud de pedido"]
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                )

                df_calculado["En_Trazabilidad"] = solpeds_me5a.isin(solpeds_no_cat)
            else:
                df_calculado["En_Trazabilidad"] = False

            df_calculado["Tipo Ariba"] = df_calculado.apply(determinar_tipo_ariba, axis=1)

            df_calculado.loc[df_calculado["En_Trazabilidad"], "Tipo Ariba"] = "🔵 ARIBA NO CATALOGADA"

            # --- INICIO NUEVA LÓGICA DE NEGOCIO ---
            es_no_catalogada = df_calculado["Tipo Ariba"] == "🔵 ARIBA NO CATALOGADA"
            sin_pr_agregada = df_calculado["Fecha de pedido"].isna() | df_calculado["Pedido"].isna()
            tiene_pr_inicial = df_calculado["Fecha de liberación"].notna()
            
            casos_pendientes = es_no_catalogada & sin_pr_agregada & tiene_pr_inicial

            fecha_hoy = pd.Timestamp.today().normalize()
            fecha_liberacion = pd.to_datetime(df_calculado["Fecha de liberación"], errors="coerce").dt.normalize()
            dias_acumulados = (fecha_hoy - fecha_liberacion).dt.days
            
            df_calculado["Nivel de Servicio"] = np.where(
                casos_pendientes,
                dias_acumulados,
                df_calculado["Nivel de Servicio"]
            )
            # --- FIN NUEVA LÓGICA DE NEGOCIO ---

            st.session_state["_df_pipeline"] = df_calculado
            st.session_state["_clave_pipeline"] = clave_actual

        except Exception as e:
            st.error(f"🚨 **No se pudieron cargar los datos.**\n\nDetalle técnico: `{e}`\n\n**Solución recomendada:** Si la política de la empresa bloquea el acceso público directo de la API a SharePoint, cambia la opción en el menú izquierdo a **'Subir archivos'** o coloca las bases en la carpeta **'data/'**.")
            st.stop()

    df = st.session_state["_df_pipeline"].copy()

    # --- INICIO PERSISTENCIA DE COMENTARIOS ---
    if "comentarios_guardados" not in st.session_state:
        st.session_state["comentarios_guardados"] = {}

    # Mapear los comentarios guardados a la base de datos principal cruzando por el número de Solped
    df["Comentario"] = df["Solicitud de pedido"].astype(str).map(st.session_state["comentarios_guardados"]).fillna("")
    # --- FIN PERSISTENCIA DE COMENTARIOS ---

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

    # ---- Filtros ----
    st.subheader("Filtros")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
    with c2:
        aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))
    with c3:
        tipos_ariba = st.multiselect("Origen / Tipo Solicitud", sorted(df["Tipo Ariba"].dropna().unique()))
    with c4:
        filtro_comentario = st.selectbox(
            "Filtrado por Comentarios", 
            ["Todos", "Con comentario", "Sin comentario"],
            help="Permite buscar las filas que ya tengan observaciones cargadas."
        )

    # ---- Filtro para Excluir IDs de Solped ----
    solpeds_excluir_raw = st.text_area(
        "🚫 Excluir Solicitudes de Pedido (IDs)", 
        placeholder="Pega las IDs a excluir separadas por coma, espacio o línea. Ej: 10045982, 3001892",
        help="Ingresa las IDs de Solped que deseas ocultar y excluir del reporte.",
        height=80
    )

    df_f = df.copy()

    if solpeds_excluir_raw.strip():
        ids_excluir = set(re.split(r'[,\s\n]+', solpeds_excluir_raw.strip()))
        ids_excluir = {i for i in ids_excluir if i}
        
        if ids_excluir:
            solpeds_str = df_f["Solicitud de pedido"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            df_f = df_f[~solpeds_str.isin(ids_excluir)]
            checkpoints.append(("0b. Tras exclusión por ID Solped", len(df_f)))

    if centros:
        df_f = df_f[df_f["Centro"].isin(centros)]
    checkpoints.append(("1. Tras filtro Centro", len(df_f)))

    if aplica:
        df_f = df_f[df_f["Aplica?"].isin(aplica)]
    checkpoints.append(("2. Tras filtro Aplica?", len(df_f)))

    if tipos_ariba:
        df_f = df_f[df_f["Tipo Ariba"].isin(tipos_ariba)]
    checkpoints.append(("2b. Tras filtro Origen / Tipo Solicitud", len(df_f)))

    if filtro_comentario == "Con comentario":
        df_f = df_f[df_f["Comentario"].astype(str).str.strip() != ""]
    elif filtro_comentario == "Sin comentario":
        df_f = df_f[df_f["Comentario"].astype(str).str.strip() == ""]
    checkpoints.append(("2c. Tras filtro Comentarios", len(df_f)))

    _snapshot("2. Tras Centro + Aplica? + Origen + Comentarios + Exclusiones", df_f)

    # ---- Estado Solped ----
    st.caption("Estado Solped (el filtro de fecha de abajo solo aplica dentro de 'Pedido completo')")
    h1, h2, h3, h4 = st.columns(4)
    with h1:
        estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

    if estados:
        df_f = df_f[df_f["Estado Solped"].isin(estados)]
    checkpoints.append(("3. Tras filtro Estado Solped (total)", len(df_f)))
    checkpoints.append(("3a. Sin pedido (antes de jerarquía)", (df_f["Estado Solped"] == "Sin pedido").sum()))
    checkpoints.append(("3b. Pedido incompleto (antes de jerarquía)", (df_f["Estado Solped"] == "Pedido incompleto").sum()))
    checkpoints.append(("3c. Pedido completo (antes de jerarquía)", (df_f["Estado Solped"] == "Pedido completo").sum()))
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

    checkpoints.append(("4a. Sin pedido tras jerarquía", (df_f["Estado Solped"] == "Sin pedido").sum()))
    checkpoints.append(("4b. Pedido incompleto tras jerarquía", (df_f["Estado Solped"] == "Pedido incompleto").sum()))
    checkpoints.append(("4c. Pedido completo tras jerarquía", (df_f["Estado Solped"] == "Pedido completo").sum()))
    checkpoints.append(("4. Total tras jerarquía de fecha", len(df_f)))
    _snapshot("4. Tras jerarquía Año/Mes/Día", df_f)

    # ---- Solped MRP y Cumple ----
    c4_b, c5 = st.columns(2)
    with c4_b:
        solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
    with c5:
        cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

    if solped_mrp:
        df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
    checkpoints.append(("5. Tras filtro Solped MRP", len(df_f)))
    _snapshot("5. Tras Solped MRP", df_f)

    if cumple:
        df_f = df_f[df_f["Cumple"].isin(cumple)]
    checkpoints.append(("6. Tras filtro Cumple", len(df_f)))
    _snapshot("6. Tras Cumple", df_f)

    st.divider()
    st.subheader("Filtro por días de gestión (Nivel de Servicio)")

    if len(df_f):
        n_negativos = (df_f["Nivel de Servicio"] < 0).sum()

        f1, f2 = st.columns([1, 2])
        with f1:
            aplicar_corte_dia_1 = st.checkbox(
                "Corte desde el día 1 (Excluir <= 0)",
                value=False,
                help="Aplica el corte de medida para que empiece desde el día 1 en realidad, excluyendo días 0 y negativos.",
            )
        with f2:
            if aplicar_corte_dia_1:
                rango_ns = (1, NS_MAX_GLOBAL)
                st.caption(f"Rango aplicado: 1 a {NS_MAX_GLOBAL:,} días (corte desde el día 1)")
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
                st.caption(f"Todas las filas tienen {NS_MIN_GLOBAL} días")

        df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

    checkpoints.append(("7. Tras filtro Nivel de Servicio (RESULTADO FINAL)", len(df_f)))
    _snapshot("7. RESULTADO FINAL", df_f)

    # ---- Diagnóstico de filtrado ----
    with st.expander("🔍 Diagnóstico de filtrado (para comparar contra el pbix)"):
        st.write("Filas en cada etapa:")
        st.table(pd.DataFrame(checkpoints, columns=["Etapa", "Filas"]))
        st.write("**Las 3 métricas en cada etapa del filtro**:")
        st.table(pd.DataFrame(metricas_por_etapa, columns=["Etapa", "Filas", "% Cumplimiento", "Prom. días", "Pos. OC"]))
        st.dataframe(
            df_f[["Solicitud de pedido", "Centro", "Estado Solped", "Fecha de pedido", "Nivel de Servicio", "Cumple", "Solped MRP", "Tipo Ariba"]]
            .sort_values("Solicitud de pedido"),
            use_container_width=True,
        )
        csv = df_f.to_csv(index=False).encode("utf-8")
        st.download_button("Descargar detalle filtrado (CSV)", csv, "detalle_filtrado.csv", "text/csv")

    # ---- Tarjetas KPI ----
    VERDE, VERDE_BORDE = "rgba(35, 145, 75, 0.16)", "rgba(35, 145, 75, 0.55)"
    ROJO, ROJO_BORDE = "rgba(204, 0, 0, 0.14)", "rgba(204, 0, 0, 0.55)"

    pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
    promedio_dias = df_f["Nivel de Servicio"].mean()
    promedio_lead_time = df_f["Lead Time Total"].mean() if "Lead Time Total" in df_f.columns else float("nan")
    pedidos_distintos = df_f["Pedido"].nunique() + (1 if df_f["Pedido"].isna().any() else 0)

    color_texto_card = "#FF3333" if st.session_state["tema"] == "oscuro" else "#404B55"
    color_sub_card = "#FF3333" if st.session_state["tema"] == "oscuro" else "#555"

    def tarjeta(titulo: str, valor: str, subtitulo: str = "", fondo: str = "rgba(64,75,85,0.07)", borde: str = "rgba(64,75,85,0.35)") -> str:
        html_sub = f'<div style="font-size:0.78rem;color:{color_sub_card};margin-top:4px;font-weight:500;">{subtitulo}</div>' if subtitulo else ""
        return (
            f'<div style="background:{fondo};border:1.5px solid {borde};border-radius:8px;'
            f'padding:12px 18px;text-align:center;">'
            f'<div style="font-size:0.78rem;color:{color_texto_card};font-weight:600;letter-spacing:.03em;'
            f'text-transform:uppercase;opacity:.85;margin-bottom:4px;">{titulo}</div>'
            f'<div style="font-size:2rem;font-weight:700;color:{color_texto_card};line-height:1.1;">{valor}</div>'
            f'{html_sub}'
            f'</div>'
        )

    if pd.isna(promedio_dias):
        f_dias, b_dias, txt_dias = "rgba(64,75,85,0.07)", "rgba(64,75,85,0.35)", "-"
    elif promedio_dias > 10:
        f_dias, b_dias, txt_dias = ROJO, ROJO_BORDE, f"{promedio_dias:.0f}"
    else:
        f_dias, b_dias, txt_dias = VERDE, VERDE_BORDE, f"{promedio_dias:.0f}"

    f_pct, b_pct = (VERDE, VERDE_BORDE) if pct_cumplimiento >= 85 else (ROJO, ROJO_BORDE)
    txt_lt = f"{promedio_lead_time:.0f}" if pd.notna(promedio_lead_time) else "-"

    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown(
            tarjeta(
                "Nivel de Servicio",
                f"{txt_dias} días",
                fondo=f_dias,
                borde=b_dias,
            ),
            unsafe_allow_html=True,
        )
    with t2:
        st.markdown(tarjeta("% Cumplimiento SLA", f"{pct_cumplimiento:.0f}%", fondo=f_pct, borde=b_pct), unsafe_allow_html=True)
    with t3:
        st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

    st.divider()

    # ---- Estilo corporativo Enaex ----
    ENAEX_GRIS = "#1E2329" if st.session_state["tema"] == "oscuro" else "#404B55"
    ENAEX_ROJO = "#CC0000"

    def tabla_enaex(tabla: pd.DataFrame, max_height: int | None = None, compacta: bool = False) -> str:
        cols = list(tabla.columns)

        def _fmt(col, val):
            if pd.isna(val):
                return "-"
            if col in ["Promedio días de gestión", "Promedio Lead Time Total"]:
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

        color_texto_tabla = "#FF3333" if st.session_state["tema"] == "oscuro" else "#404B55"

        filas = []
        for i, r in enumerate(tabla.itertuples(index=False)):
            es_total = str(r[0]) == "TOTAL"
            if es_total:
                estilo_fila = f"background:{ENAEX_GRIS};color:#fff;font-weight:700;border-top:2px solid {ENAEX_ROJO};"
            else:
                fondo = "#ffffff" if (i % 2 == 0 and st.session_state["tema"] == "claro") else ("
