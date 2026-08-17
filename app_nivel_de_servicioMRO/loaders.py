"""
Carga de datos - equivalente en Python a las queries de Power Query (M) del pbix:

    M (Power BI)                          ->  Python (acá)
    ------------------------------------------------------------------
    Data (2)                              -> cargar_data_pr()
    Responsable por Grupo de Compras      -> cargar_responsable_grupo_compras()
    CENTRO_SOCIEDAD Compra MRO            -> cargar_centro_sociedad_mro()
    Responsable de MRP                    -> cargar_responsable_mrp()

Soporta lectura híbrida (.parquet y .xlsx) tanto desde ruta local como desde Streamlit UploadedFile.
"""
from typing import Any
import pandas as pd
import streamlit as st

import config


def _leer_archivo(origen: Any, sheet_name: str = None) -> pd.DataFrame:
    """
    Función auxiliar para cargar .parquet o .xlsx de manera transparente,
    ya sea desde una ruta (str/Path) o desde un UploadedFile de Streamlit.
    """
    nombre = getattr(origen, "name", str(origen)).lower()
    if nombre.endswith(".parquet"):
        return pd.read_parquet(origen)
    
    # Si es Excel
    if sheet_name:
        try:
            return pd.read_excel(origen, sheet_name=sheet_name)
        except Exception:
            # Si el sheet_name no existe o fue exportado directo, cae a la primera hoja
            return pd.read_excel(origen)
    return pd.read_excel(origen)


@st.cache_data(show_spinner="Cargando datos de solicitudes de pedido (SAP/Ariba)...")
def cargar_data_pr(ruta_o_archivo: Any = None) -> pd.DataFrame:
    """
    Equivalente a la query 'Data (2)' del pbix (solo la carga + tipado,
    el resto de la lógica vive en transform.pipeline_completo).
    """
    origen = ruta_o_archivo if ruta_o_archivo is not None else config.RUTA_DATA_ME5A
    df = _leer_archivo(origen, sheet_name="Data")

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
    df["Pedido"] = pd.to_numeric(df["Pedido"].astype(str).replace(r"^\s*$", pd.NA, regex=True), errors="coerce").astype("Int64")
    df["Fecha de pedido"] = pd.to_datetime(df["Fecha de pedido"], errors="coerce")
    df["Posición de pedido"] = pd.to_numeric(df["Posición de pedido"], errors="coerce").astype("Int64")
    
    # Indicador liberación: vacío y espacio en blanco se tratan como "sin indicador"
    if "Indicador liberación" in df.columns:
        df["Indicador liberación"] = df["Indicador liberación"].astype(str).replace(r"^\s*$", pd.NA, regex=True)

    return df


@st.cache_data(show_spinner="Cargando responsables por grupo de compras...")
def cargar_responsable_grupo_compras(ruta_o_archivo: Any = None) -> pd.DataFrame:
    """
    Equivalente a 'Responsable por Grupo de Compras' (lista de SharePoint).
    Columnas esperadas (según el M): "Grupo de Compras", "Responsable.title"
    """
    origen = ruta_o_archivo if ruta_o_archivo is not None else config.RUTA_RESP_GRUPO_COMPRAS
    df = _leer_archivo(origen)
    
    if "Responsable.title" in df.columns:
        df = df.rename(columns={"Responsable.title": "Comprador por Grupo Compras"})
    elif "Responsable" in df.columns and "Comprador por Grupo Compras" not in df.columns:
        df = df.rename(columns={"Responsable": "Comprador por Grupo Compras"})
        
    df["Grupo de Compras"] = df["Grupo de Compras"].astype(str).str.strip()
    return df[["Grupo de Compras", "Comprador por Grupo Compras"]]


@st.cache_data(show_spinner="Cargando centros y sociedades MRO...")
def cargar_centro_sociedad_mro(ruta_o_archivo: Any = None) -> pd.DataFrame:
    """
    Equivalente a 'CENTRO_SOCIEDAD Compras MRO'.
    Columnas esperadas (según el M): "Título" (clave = Centro SAP),
    "Nombre Centro", "Nombre Centro 2".
    """
    origen = ruta_o_archivo if ruta_o_archivo is not None else config.RUTA_CENTRO_SOCIEDAD
    df = _leer_archivo(origen)
    df["Título"] = df["Título"].astype(str).str.strip()
    return df[["Título", "Nombre Centro", "Nombre Centro 2"]]


@st.cache_data(show_spinner="Cargando responsables MRP...")
def cargar_responsable_mrp(ruta_o_archivo: Any = None) -> pd.DataFrame:
    """
    Equivalente a 'Responsable de MRP'.
    Columnas esperadas (según el M): "Title" (clave = usuario SAP = campo
    "Autor" en Data), "Responsable Compra.title".
    """
    origen = ruta_o_archivo if ruta_o_archivo is not None else config.RUTA_RESP_MRP
    df = _leer_archivo(origen)
    
    if "Responsable Compra.title" in df.columns:
        df = df.rename(columns={"Responsable Compra.title": "Responsable de MRP.Responsable Compra.title"})
    elif "Responsable Compra" in df.columns and "Responsable de MRP.Responsable Compra.title" not in df.columns:
        df = df.rename(columns={"Responsable Compra": "Responsable de MRP.Responsable Compra.title"})
        
    df["Title"] = df["Title"].astype(str).str.strip()
    return df[["Title", "Responsable de MRP.Responsable Compra.title"]]
