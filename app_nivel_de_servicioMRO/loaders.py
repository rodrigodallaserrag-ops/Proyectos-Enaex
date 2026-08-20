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

import streamlit as st

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
# Uso: en app.py, pasar el string "onedrive:<nombre_secreto>" como `ruta` a
# cualquiera de los cargar_* de más abajo, en vez de una ruta local o un
# UploadedFile. El nombre_secreto debe existir en st.secrets["onedrive"].
#
# Formato esperado en Secrets (Streamlit Cloud -> Manage app -> Settings -> Secrets):
#
#   [onedrive]
#   me5a_parquet = "https://enaex-my.sharepoint.com/:x:/g/personal/.../XXXX?download=1"
#   responsable_grupo_compras = "https://enaex-my.sharepoint.com/:x:/g/personal/.../XXXX?download=1"
#   centro_sociedad_mro = "https://enaex-my.sharepoint.com/:x:/g/personal/.../XXXX?download=1"
#   responsable_mrp = "https://enaex-my.sharepoint.com/:x:/g/personal/.../XXXX?download=1"
#
# Cómo conseguir cada URL: en OneDrive, clic derecho en el archivo -> "Compartir"
# -> "Copiar vínculo" (con permiso "Personas de Enaex con el vínculo") -> pegar
# esa URL y agregarle "&download=1" al final (o "?download=1" si no tiene "?").
# Eso hace que la URL entregue el archivo directo en vez de abrir el visor web.
#
# IMPORTANTE para que el link nunca se rompa: al actualizar el archivo,
# REEMPLAZA el contenido en el mismo lugar (arrastrarlo encima y elegir
# "Reemplazar"), no lo borres y subas uno nuevo — borrar+resubir cambia el ID
# interno del archivo y invalida el link, obligando a actualizar el secret.
ONEDRIVE_SENTINEL_PREFIX = "onedrive:"


@st.cache_data(show_spinner="Descargando desde OneDrive...", ttl=3600, max_entries=4)
def _descargar_onedrive(nombre_secreto: str) -> bytes:
    """
    Descarga el contenido crudo de un archivo compartido en OneDrive/SharePoint.

    Cacheado con ttl=3600 (1 hora): durante ese tiempo sirve la copia ya
    descargada sin volver a golpear la red en cada filtro/rerun; pasada la
    hora, se refresca sola en el próximo acceso — así el ME5A semanal se
    actualiza automáticamente sin que nadie tenga que tocar nada, y los 3
    archivos estáticos (que casi no cambian) igual se refrescan solos de vez
    en cuando por si acaso, sin costo real ya que casi siempre serán idénticos.
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
    `archivo` tal cual (ruta local o UploadedFile de Streamlit) — así el resto
    del código de cada cargar_* no necesita saber de dónde vino el archivo.
    """
    if isinstance(archivo, str) and archivo.startswith(ONEDRIVE_SENTINEL_PREFIX):
        nombre_secreto = archivo[len(ONEDRIVE_SENTINEL_PREFIX):]
        contenido = _descargar_onedrive(nombre_secreto)
        return io.BytesIO(contenido)
    return archivo


def _columnas_normalizadas(df: pd.DataFrame) -> pd.DataFrame:
    """Quita espacios en blanco (incluyendo NBSP) al inicio/fin de cada nombre
    de columna. Encabezados con espacios invisibles son la causa más común de
    un KeyError "misterioso" al leer Excel exportado desde SharePoint/SAP."""
    df = df.copy()
    df.columns = df.columns.str.strip().str.replace("\xa0", " ", regex=False).str.strip()
    return df


def _requerir_columna(df: pd.DataFrame, nombre: str, archivo: str) -> None:
    """Falla con un mensaje claro (columnas disponibles incluidas) en vez de
    un KeyError genérico, para diagnosticar de un vistazo un archivo con
    encabezados distintos a los esperados."""
    if nombre not in df.columns:
        raise KeyError(
            f"No encontré la columna '{nombre}' en {archivo}. "
            f"Columnas disponibles: {list(df.columns)}"
        )


def _tipar_data_pr(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tipado equivalente al "Tipo cambiado" del M. Se separó de la lectura para
    poder reutilizarla tanto al leer el .xlsx original como en la página de
    conversión a Parquet (preparar_datos.py) — el .parquet se genera A PARTIR
    de este mismo tipado, así que al leerlo de vuelta ya viene tipado y este
    paso no hay que repetirlo (ver cargar_data_pr más abajo).
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
    # Empty string / whitespace -> NaN, igual que Power Query al tipar a Int64.Type
    df["Pedido"] = pd.to_numeric(df["Pedido"].replace(r"^\s*$", pd.NA, regex=True), errors="coerce").astype("Int64")
    df["Fecha de pedido"] = pd.to_datetime(df["Fecha de pedido"], errors="coerce")
    df["Posición de pedido"] = pd.to_numeric(df["Posición de pedido"], errors="coerce").astype("Int64")
    # Indicador liberación: vacío y espacio en blanco se tratan como "sin indicador"
    df["Indicador liberación"] = df["Indicador liberación"].replace(r"^\s*$", pd.NA, regex=True)
    return df


def _es_parquet(archivo) -> bool:
    """Detecta si el archivo/ruta es .parquet, por nombre (UploadedFile) o
    extensión (ruta string)."""
    nombre = getattr(archivo, "name", None) or str(archivo)
    return nombre.lower().endswith(".parquet")


@st.cache_data(show_spinner="Cargando datos de solicitudes de pedido (SAP/Ariba)...", max_entries=3, ttl=3600)
def cargar_data_pr(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a la query 'Data (2)' del pbix (solo la carga + tipado,
    el resto de la lógica vive en transform.pipeline_completo).

    Acepta tanto .xlsx (se tipa al vuelo, más lento y con más overhead de
    memoria por el parseo de openpyxl) como .parquet (ya viene tipado desde
    preparar_datos.py, lectura columnar casi instantánea y con una fracción
    del pico de memoria de un .xlsx equivalente).
    """
    ruta = ruta or config.RUTA_DATA_ME5A
    es_onedrive_parquet = isinstance(ruta, str) and ruta == "onedrive:me5a_parquet"
    fuente = _resolver_fuente(ruta)

    if es_onedrive_parquet or _es_parquet(fuente):
        # El parquet se generó a partir de _tipar_data_pr, así que ya viene
        # con los dtypes correctos: no hace falta volver a tipar ni normalizar
        # columnas, eso es justamente lo que hace rápida esta rama.
        return pd.read_parquet(fuente)

    df = pd.read_excel(fuente, sheet_name="Data")
    return _tipar_data_pr(df)


@st.cache_data(show_spinner="Cargando responsables por grupo de compras...", max_entries=3, ttl=3600)
def cargar_responsable_grupo_compras(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a 'Responsable por Grupo de Compras' (lista de SharePoint).
    Columnas esperadas (según el M): "Grupo de Compras", "Responsable.title"
    (el ".title" es el nombre a mostrar de una columna Persona/Grupo de SharePoint).
    Si tu export trae la columna Persona con otro nombre (ej. "Responsable"),
    ajusta el rename de más abajo.
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
    Columnas esperadas (según el M): "Título" (clave = Centro SAP),
    "Nombre Centro", "Nombre Centro 2".
    OJO: el M original intercambia los nombres al expandir (posible
    inconsistencia de origen) - se replica tal cual para que el resultado
    calce con el pbix.
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
    Columnas esperadas (según el M): "Title" (clave = usuario SAP = campo
    "Autor" en Data), "Responsable Compra.title".
    """
    ruta = ruta or config.RUTA_RESP_MRP
    df = pd.read_excel(_resolver_fuente(ruta))
    df = _columnas_normalizadas(df)
    df = df.rename(columns={"Responsable Compra.title": "Responsable de MRP.Responsable Compra.title"})
    _requerir_columna(df, "Title", "Responsable_MRP.xlsx")
    _requerir_columna(df, "Responsable de MRP.Responsable Compra.title", "Responsable_MRP.xlsx (columna 'Responsable Compra.title')")
    df["Title"] = df["Title"].astype(str).str.strip()
    return df[["Title", "Responsable de MRP.Responsable Compra.title"]]
