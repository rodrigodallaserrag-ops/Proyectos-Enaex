"""
Carga de datos - equivalente en Python a las queries de Power Query (M) del pbix:

    M (Power BI)                          ->  Python (acá)
    ------------------------------------------------------------------
    Data (2)                              -> cargar_data_pr()
    Responsable por Grupo de Compras      -> cargar_responsable_grupo_compras()
    CENTRO_SOCIEDAD Compra MRO            -> cargar_centro_sociedad_mro()
    Responsable de MRP                    -> cargar_responsable_mrp()

Fase actual: todo se lee de archivos locales (Excel).
Fase Azure: cada función cambia SOLO por dentro (SharePoint -> Graph API,
Data -> Blob Storage) - las funciones que las consumen (transform.py, app.py)
no se tocan.
"""
import io

import pandas as pd
import requests
import streamlit as st

import config

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

# ---- Carga automática desde OneDrive for Business (opcional) ----
ONEDRIVE_SENTINEL_PREFIX = "onedrive:"

@st.cache_data(show_spinner="Descargando desde OneDrive...", ttl=3600, max_entries=4)
def _descargar_onedrive(nombre_secreto: str) -> bytes:
    """
    Descarga el contenido crudo de un archivo compartido en OneDrive/SharePoint.
    """
    if "onedrive" not in st.secrets or nombre_secreto not in st.secrets["onedrive"]:
        raise KeyError(
            f"Falta el secreto 'onedrive.{nombre_secreto}'. Configúralo en "
            "Streamlit Cloud -> Manage app -> Settings -> Secrets con el link "
            "de descarga directa de OneDrive."
        )
    url = st.secrets["onedrive"][nombre_secreto]
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    return resp.content


def _resolver_fuente(archivo):
    """
    Si `archivo` es el sentinel 'onedrive:<nombre_secreto>', descarga desde
    OneDrive y devuelve un BytesIO listo para pandas. Si no, devuelve
    `archivo` tal cual.
    """
    if isinstance(archivo, str) and archivo.startswith(ONEDRIVE_SENTINEL_PREFIX):
        nombre_secreto = archivo[len(ONEDRIVE_SENTINEL_PREFIX):]
        contenido = _descargar_onedrive(nombre_secreto)
        return io.BytesIO(contenido)
    return archivo


def _columnas_normalizadas(df: pd.DataFrame) -> pd.DataFrame:
    """Quita espacios en blanco (incluyendo NBSP) al inicio/fin de cada nombre
    de columna."""
    df = df.copy()
    df.columns = df.columns.str.strip().str.replace("\xa0", " ", regex=False).str.strip()
    return df


def _requerir_columna(df: pd.DataFrame, nombre: str, archivo: str) -> None:
    """Falla con un mensaje claro en vez de un KeyError genérico."""
    if nombre not in df.columns:
        raise KeyError(
            f"No encontré la columna '{nombre}' en {archivo}. "
            f"Columnas disponibles: {list(df.columns)}"
        )


def _tipar_data_pr(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tipado equivalente al "Tipo cambiado" del M.
    """
    df = _columnas_normalizadas(df)
    df["Centro"] = df["Centro"].astype(str).str.strip()
    df["Material"] = pd.to_numeric(df["Material"], errors="coerce").astype("Int64")
    df["Pos.solicitud pedido"] = pd.to_numeric(df["Pos.solicitud pedido"], errors="coerce").astype("Int64")
    df["Solicitud de pedido"] = pd.to_numeric(df["Solicitud de pedido"], errors="coerce").astype("Int64")
    df["Fecha de solicitud"] = pd.to_datetime(df["Fecha de solicitud"], errors="coerce")
    df["Fecha modificación"] = pd.to_datetime(df["Fecha modificación"], errors="coerce")
    df["Fecha de liberación"] = pd.to_datetime(df["Fecha de liberación"], errors="coerce")
    df["Cantidad solicitada"] = pd.to_numeric(df["Cantidad solicitada"], errors="coerce")
    df["Valor total"] = pd.to_numeric(df["Valor total"], errors="coerce")
    df["Grupo de compras"] = df["Grupo de compras"].astype(str).str.strip()
    df["Cantidad pedida"] = pd.to_numeric(df["Cantidad pedida"], errors="coerce")
    df["Autor"] = df["Autor"].astype(str).str.strip()
    df["Pedido"] = pd.to_numeric(df["Pedido"].replace(r"^\s*$", pd.NA, regex=True), errors="coerce").astype("Int64")
    df["Fecha de pedido"] = pd.to_datetime(df["Fecha de pedido"], errors="coerce")
    df["Posición de pedido"] = pd.to_numeric(df["Posición de pedido"], errors="coerce").astype("Int64")
    df["Indicador liberación"] = df["Indicador liberación"].replace(r"^\s*$", pd.NA, regex=True)
    return df


def _es_parquet(archivo) -> bool:
    """Detecta si el archivo/ruta es .parquet."""
    nombre = getattr(archivo, "name", None) or str(archivo)
    return nombre.lower().endswith(".parquet")


@st.cache_data(show_spinner="Cargando datos de solicitudes de pedido (SAP/Ariba)...", max_entries=3, ttl=3600)
def cargar_data_pr(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a la query 'Data (2)' del pbix.
    """
ruta = ruta or config.RUTA_DATA_ME5A
    # es_onedrive_parquet = isinstance(ruta, str) and ruta == "onedrive:me5a_parquet" # Puedes borrar esta línea
    fuente = _resolver_fuente(ruta)

    # Solo usamos parquet si el archivo realmente tiene extensión .parquet
    if _es_parquet(fuente):
        return pd.read_parquet(fuente)

    # Si no es parquet (como el Excel de OneDrive), sigue de largo y lo lee como Excel:
    df = pd.read_excel(fuente, sheet_name="Data")
    return _tipar_data_pr(df)


@st.cache_data(show_spinner="Cargando responsables por grupo de compras...", max_entries=3, ttl=3600)
def cargar_responsable_grupo_compras(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a 'Responsable por Grupo de Compras' (lista de SharePoint).
    """
    ruta = ruta or config.RUTA_RESP_GRUPO_COMPRAS
    df = pd.read_excel(_resolver_fuente(ruta))
    df = _columnas_normalizadas(df)
    df = df.rename(columns={"Responsable.title": "Comprador por Grupo Compras"})
    _requerir_columna(df, "Grupo de Compras", "Responsable_Grupo_Compras.xlsx")
    _requerir_columna(df, "Comprador por Grupo Compras", "Responsable_Grupo_Compras.xlsx (columna 'Responsable.title')")
    df["Grupo de Compras"] = df["Grupo de Compras"].astype(str).str.strip()
    return df[["Grupo de Compras", "Comprador por Grupo Compras"]]


@st.cache_data(show_spinner="Cargando centros y sociedades MRO...", max_entries=3, ttl=3600)
def cargar_centro_sociedad_mro(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a 'CENTRO_SOCIEDAD Compras MRO'.
    """
    ruta = ruta or config.RUTA_CENTRO_SOCIEDAD
    df = pd.read_excel(_resolver_fuente(ruta))
    df = _columnas_normalizadas(df)
    _requerir_columna(df, "Título", "Centro_Sociedad_MRO.xlsx")
    _requerir_columna(df, "Nombre Centro", "Centro_Sociedad_MRO.xlsx")
    _requerir_columna(df, "Nombre Centro 2", "Centro_Sociedad_MRO.xlsx")
    df["Título"] = df["Título"].astype(str).str.strip()
    return df[["Título", "Nombre Centro", "Nombre Centro 2"]]


@st.cache_data(show_spinner="Cargando responsables MRP...", max_entries=3, ttl=3600)
def cargar_responsable_mrp(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a 'Responsable de MRP'.
    """
    ruta = ruta or config.RUTA_RESP_MRP
    df = pd.read_excel(_resolver_fuente(ruta))
    df = _columnas_normalizadas(df)
    df = df.rename(columns={"Responsable Compra.title": "Responsable de MRP.Responsable Compra.title"})
    _requerir_columna(df, "Title", "Responsable_MRP.xlsx")
    _requerir_columna(df, "Responsable de MRP.Responsable Compra.title", "Responsable_MRP.xlsx (columna 'Responsable Compra.title')")
    df["Title"] = df["Title"].astype(str).str.strip()
    return df[["Title", "Responsable de MRP.Responsable Compra.title"]]
