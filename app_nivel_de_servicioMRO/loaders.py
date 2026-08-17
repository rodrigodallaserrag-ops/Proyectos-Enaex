"""
Módulo de carga y optimización de datos para la aplicación de Nivel de Servicio.
Equivale a las capas de extracción (M) de Power Query.
"""
import pandas as pd
import streamlit as st


@st.cache_data(ttl="1h")
def cargar_data_pr(origen) -> pd.DataFrame:
    """
    Carga el reporte base (ME5A_con_Ariba.xlsx) ya sea desde ruta local o UploadedFile.
    Aplica tipado estricto para optimizar memoria y búsquedas.
    """
    df = pd.read_excel(origen)

    # Conversión a tipos enteros optimizados
    cols_entero = ["Solicitud de pedido", "Posición de pedido", "Material", "Centro"]
    for col in cols_entero:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    # Conversión a fechas datetime
    cols_fecha = ["Fecha de solicitud", "Fecha modificación", "Fecha de pedido"]
    for col in cols_fecha:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Limpieza de columna Pedido
    if "Pedido" in df.columns:
        df["Pedido"] = df["Pedido"].astype(str).str.strip().replace(["nan", "None", "<NA>", ""], None)

    return df


@st.cache_data(ttl="1h")
def cargar_responsable_grupo_compras(origen) -> pd.DataFrame:
    """
    Carga la tabla de mapeo Responsable_Grupo_Compras.xlsx.
    """
    df = pd.read_excel(origen)
    if "Grupo de compras" in df.columns:
        df["Grupo de compras"] = df["Grupo de compras"].astype(str).str.strip()
    return df


@st.cache_data(ttl="1h")
def cargar_centro_sociedad_mro(origen) -> pd.DataFrame:
    """
    Carga la tabla maestra Centro_Sociedad_MRO.xlsx.
    """
    df = pd.read_excel(origen)
    if "Centro" in df.columns:
        df["Centro"] = pd.to_numeric(df["Centro"], errors="coerce").astype("Int64")
    return df


@st.cache_data(ttl="1h")
def cargar_responsable_mrp(origen) -> pd.DataFrame:
    """
    Carga la tabla maestra Responsable_MRP.xlsx.
    """
    df = pd.read_excel(origen)
    if "Solped MRP" in df.columns:
        df["Solped MRP"] = df["Solped MRP"].astype(str).str.strip()
    return df
