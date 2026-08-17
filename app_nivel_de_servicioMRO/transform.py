"""
Motor de transformación y reglas de negocio.
Traduce la lógica de Power Query y DAX del modelo de Power BI original.
"""
import numpy as np
import pandas as pd

import config


def pipeline_completo(df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=None) -> pd.DataFrame:
    """
    Ejecuta el flujo completo de limpieza, cruzado de tablas compuestas y
    cálculo de indicadores SLA/Nivel de Servicio.
    """
    if fecha_corte is None:
        fecha_corte = pd.Timestamp.today()
    else:
        fecha_corte = pd.Timestamp(fecha_corte)

    df = df_data.copy()

    # 1. Cruzado con Responsable de Grupo de Compras
    if not df_resp_grupo.empty and "Grupo de compras" in df.columns:
        df = df.merge(df_resp_grupo, on="Grupo de compras", how="left")
        if "Comprador" in df.columns:
            df.rename(columns={"Comprador": "Comprador (Grupo de compras)"}, inplace=True)

    if "Comprador (Grupo de compras)" not in df.columns:
        df["Comprador (Grupo de compras)"] = "Sin Asignar"

    # 2. Cruzado con Maestros de Centro y Sociedad
    if not df_centro_sociedad.empty and "Centro" in df.columns:
        df = df.merge(df_centro_sociedad, on="Centro", how="left")

    for col_centro in ["Nombre Centro", "Nombre Centro 2"]:
        if col_centro not in df.columns:
            df[col_centro] = df["Centro"].astype(str)

    # 3. Identificación de Solped MRP y Origen de Solicitud
    if "Solped MRP" not in df.columns:
        df["Solped MRP"] = "ERP/Manual"

    es_ariba = df["Solicitud de pedido"] >= config.UMBRAL_SOLPED_ARIBA
    df["Origen"] = np.where(es_ariba, "Ariba", "ERP/MRP")

    # 4. Clasificación del Estado de la Solped
    tiene_pedido = df["Pedido"].notna() & (df["Pedido"] != "")
    df["Estado Solped"] = np.where(tiene_pedido, "Pedido completo", "Sin pedido")

    # 5. Evaluación del Indicador "Aplica?"
    dias_desde_solicitud = (fecha_corte - df["Fecha de solicitud"]).dt.days
    df["Aplica?"] = np.where(
        tiene_pedido | (dias_desde_solicitud >= config.DIAS_GRACIA_SIN_PEDIDO),
        "SI",
        "NO"
    )

    # 6. Cálculo del Nivel de Servicio (Días de Gestión)
    dias_con_pedido = (df["Fecha de pedido"] - df["Fecha de solicitud"]).dt.days
    df["Nivel de Servicio"] = np.where(tiene_pedido, dias_con_pedido, dias_desde_solicitud)

    # 7. Evaluación de SLA (Cumple / No cumple)
    sla_limite = np.where(df["Origen"] == "Ariba", config.SLA_DIAS_ARIBA, config.SLA_DIAS_ERP_MRP)
    df["Cumple"] = np.where(df["Nivel de Servicio"] <= sla_limite, "Cumple", "No cumple")

    return df


def calcular_metricas_por_grupo(df: pd.DataFrame, groupby_cols: list) -> pd.DataFrame:
    """
    Agrupa por las columnas indicadas y calcula las 3 métricas clave del dashboard.
    """
    if df.empty or not groupby_cols:
        columnas_salida = list(groupby_cols) + ["Promedio días de gestión", "% Cumplimiento", "Pos. OC generadas"]
        return pd.DataFrame(columns=columnas_salida)

    res = (
        df.groupby(groupby_cols, dropna=False)
        .apply(
            lambda g: pd.Series(
                {
                    "Promedio días de gestión": g["Nivel de Servicio"].mean(),
                    "% Cumplimiento": (g["Cumple"] == "Cumple").sum() / len(g) * 100 if len(g) > 0 else 0,
                    "Pos. OC generadas": g["Pedido"].nunique() + (1 if g["Pedido"].isna().any() else 0),
                }
            )
        )
        .reset_index()
    )

    return res


def tabla_centros_fija(df: pd.DataFrame) -> pd.DataFrame:
    """
    Genera el desglose resumido de métricas agrupando por la jerarquía fija de Centros.
    """
    col_agrupacion = ["Nombre Centro 2"] if "Nombre Centro 2" in df.columns else ["Centro"]
    return calcular_metricas_por_grupo(df, col_agrupacion)


def agregar_fila_total(tabla: pd.DataFrame, df_origen: pd.DataFrame, groupby_cols: list) -> pd.DataFrame:
    """
    Consolida y anexa la fila final de TOTAL a la tabla formateada.
    """
    if df_origen.empty:
        return tabla

    prom_dias = df_origen["Nivel de Servicio"].mean()
    pct_cumple = (df_origen["Cumple"] == "Cumple").sum() / len(df_origen) * 100 if len(df_origen) > 0 else 0
    pos_oc = df_origen["Pedido"].nunique() + (1 if df_origen["Pedido"].isna().any() else 0)

    total_row = {col: "TOTAL" if idx == 0 else "" for idx, col in enumerate(groupby_cols)}
    total_row["Promedio días de gestión"] = prom_dias
    total_row["% Cumplimiento"] = pct_cumple
    total_row["Pos. OC generadas"] = pos_oc

    fila_total_df = pd.DataFrame([total_row])
    return pd.concat([tabla, fila_total_df], ignore_index=True)
