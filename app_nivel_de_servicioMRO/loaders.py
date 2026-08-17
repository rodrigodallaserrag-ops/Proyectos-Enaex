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
import pandas as pd
import streamlit as st

import config


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


@st.cache_data(show_spinner="Cargando datos de solicitudes de pedido (SAP/Ariba)...", max_entries=3, ttl=3600)
def cargar_data_pr(ruta: str = None) -> pd.DataFrame:
    """
    Equivalente a la query 'Data (2)' del pbix (solo la carga + tipado,
    el resto de la lógica vive en transform.pipeline_completo).
    """
    ruta = ruta or config.RUTA_DATA_ME5A
    df = pd.read_excel(ruta, sheet_name="Data")
    df = _columnas_normalizadas(df)

    # Tipado equivalente al "Tipo cambiado" del M
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
    df = pd.read_excel(ruta)
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
    df = pd.read_excel(ruta)
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
    df = pd.read_excel(ruta)
    df = _columnas_normalizadas(df)
    df = df.rename(columns={"Responsable Compra.title": "Responsable de MRP.Responsable Compra.title"})
    _requerir_columna(df, "Title", "Responsable_MRP.xlsx")
    _requerir_columna(df, "Responsable de MRP.Responsable Compra.title", "Responsable_MRP.xlsx (columna 'Responsable Compra.title')")
    df["Title"] = df["Title"].astype(str).str.strip()
    return df[["Title", "Responsable de MRP.Responsable Compra.title"]]
