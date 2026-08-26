"""
Transformaciones - traducción 1:1 del código M (Power Query) de la query
"Data (2)" del pbix Nivel_de_servicio_BI.pbix.

Incluye la separación de Fase 0 (SLA Comprador desde Fecha de liberación vs Lead Time Total).
"""
import numpy as np
import pandas as pd

import config


def _dias_entre(fecha_fin: pd.Series, fecha_inicio: pd.Series) -> pd.Series:
    """fecha_fin - fecha_inicio, en días (equivalente a la resta de fechas en M)."""
    return (fecha_fin - fecha_inicio).dt.days


def unir_responsables(
    df: pd.DataFrame,
    df_resp_grupo: pd.DataFrame,
    df_resp_mrp: pd.DataFrame,
) -> pd.DataFrame:
    """
    Pasos M: "Consultas combinadas", "Se expandió Responsable por Grupo de
    Compras", "Consultas combinadas1", "Se expandió Responsable de MRP",
    "Personalizada agregada" (Comprador por Grupo Compras2).
    """
    df = df.merge(
        df_resp_grupo, left_on="Grupo de compras", right_on="Grupo de Compras", how="left"
    ).drop(columns=["Grupo de Compras"])

    df = df.merge(df_resp_mrp, left_on="Autor", right_on="Title", how="left").drop(columns=["Title"])

    resp_mrp_col = "Responsable de MRP.Responsable Compra.title"
    df["Comprador por Grupo Compras2"] = np.where(
        df[resp_mrp_col].isna(), df["Comprador por Grupo Compras"], df[resp_mrp_col]
    )
    return df


def calcular_solped_mrp(df: pd.DataFrame) -> pd.DataFrame:
    """Paso M: 'Personalizada agregada1' -> columna 'Solped MRP'."""
    df = df.copy()
    resp_mrp_col = "Responsable de MRP.Responsable Compra.title"
    sin_resp_mrp = df[resp_mrp_col].isna()
    es_ariba = sin_resp_mrp & (df["Solicitud de pedido"] > config.UMBRAL_SOLICITUD_ARIBA)

    df["Solped MRP"] = np.select(
        [es_ariba, sin_resp_mrp],
        ["Ariba", "ERP"],
        default="MRP",
    )
    return df


def filtrar_solicitudes_vigentes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pasos M: 'Personalizado' (flag 2/0), 'Personalizado.3', 'Filas filtradas'.
    """
    df = df.copy()
    ind = df["Indicador liberación"]
    sin_indicador = ind.isna()

    personalizado = np.where(
        df["Solped MRP"].isin(["MRP", "Ariba"]) | ((df["Solped MRP"] == "ERP") & sin_indicador),
        2,
        0,
    )

    ind_num = pd.to_numeric(ind, errors="coerce")
    personalizado_3 = np.select(
        [sin_indicador, ind.isin(["X", "B"])],
        [personalizado, 0],
        default=ind_num,
    )

    return df[personalizado_3 == 2].copy()


def calcular_nivel_servicio_dias(df: pd.DataFrame, fecha_corte: pd.Timestamp) -> pd.DataFrame:
    """
    Calcula:
    1. SLA Comprador: Días transcurridos desde Fecha modificación.
       (Alineado con la query M: el inicio del reloj SIEMPRE es Fecha
       modificación, sin cadena de respaldo hacia Fecha de liberación ni
       Fecha de solicitud.)
    2. Lead Time Total: Días transcurridos desde Fecha de solicitud inicial (usuario).
    """
    df = df.copy()

    f_corte = pd.Timestamp(fecha_corte)
    f_ped = pd.to_datetime(df["Fecha de pedido"], errors="coerce")
    f_sol = pd.to_datetime(df["Fecha de solicitud"], errors="coerce")
    f_mod = pd.to_datetime(df["Fecha modificación"], errors="coerce")

    # Inicio SLA Comprador: SIEMPRE Fecha modificación (igual que la query M).
    f_inicio_sla = f_mod

    # Días para SLA Comprador
    df["_fecha_repor_inicio_sla"] = (f_corte - f_inicio_sla).dt.days
    df["_fecha_pedi_inicio_sla"] = (f_ped - f_inicio_sla).dt.days

    # Días para Lead Time Total
    df["_fecha_repor_solicitud"] = (f_corte - f_sol).dt.days
    df["_fecha_pedi_solicitud"] = (f_ped - f_sol).dt.days

    sin_pedido = df["Pedido"].isna()

    # Nivel de Servicio v1 (SLA Comprador)
    df["_nivel_servicio_v1"] = np.where(
        sin_pedido, df["_fecha_repor_inicio_sla"], df["_fecha_pedi_inicio_sla"]
    )

    # Lead Time Total (Usuario)
    df["Lead Time Total"] = np.where(
        sin_pedido, df["_fecha_repor_solicitud"], df["_fecha_pedi_solicitud"]
    )

    return df


def unir_centro_sociedad(df: pd.DataFrame, df_centro_sociedad: pd.DataFrame) -> pd.DataFrame:
    """
    Paso M: 'Consultas combinadas2' + expandir + rename.
    """
    df = df.merge(df_centro_sociedad, left_on="Centro", right_on="Título", how="left").drop(
        columns=["Título"]
    )
    df = df.rename(columns={"Nombre Centro": "Nombre Centro 2", "Nombre Centro 2": "Nombre Centro"})
    return df


def calcular_estado_solped_y_nivel_servicio(df: pd.DataFrame) -> pd.DataFrame:
    """
    Asigna el Estado Solped y consolida Nivel de Servicio (SLA Comprador final).
    """
    df = df.copy()
    sin_pedido = df["Pedido"].isna()
    pedido_completo = df["Cantidad pedida"] == df["Cantidad solicitada"]

    df["Estado Solped"] = np.select(
        [sin_pedido, pedido_completo],
        ["Sin pedido", "Pedido completo"],
        default="Pedido incompleto",
    )

    df["Nivel de Servicio"] = np.where(
        df["Estado Solped"] == "Pedido incompleto",
        df["_fecha_repor_inicio_sla"],
        df["_nivel_servicio_v1"],
    )
    return df


def calcular_aplica(df: pd.DataFrame) -> pd.DataFrame:
    """Determina si la solicitud aplica según días de gracia."""
    df = df.copy()
    check = np.where(df["Estado Solped"] == "Sin pedido", df["Nivel de Servicio"], np.nan)
    df["Aplica?"] = np.where(
        pd.isna(check) | (check >= config.DIAS_GRACIA_SIN_PEDIDO), "Aplica", "No aplica"
    )
    return df


def calcular_cumple(df: pd.DataFrame) -> pd.DataFrame:
    """SLA: 10 días para ERP/MRP, 7 días para Ariba."""
    df = df.copy()
    cumple_erp_mrp = (df["Nivel de Servicio"] <= config.SLA_DIAS_ERP_MRP) & df["Solped MRP"].isin(
        ["ERP", "MRP"]
    )
    cumple_ariba = (df["Nivel de Servicio"] <= config.SLA_DIAS_ARIBA) & (df["Solped MRP"] == "Ariba")
    df["Cumple"] = np.where(cumple_erp_mrp | cumple_ariba, "Cumple", "No cumple")
    return df


def pipeline_completo(
    df_data: pd.DataFrame,
    df_resp_grupo: pd.DataFrame,
    df_centro_sociedad: pd.DataFrame,
    df_resp_mrp: pd.DataFrame,
    fecha_corte: pd.Timestamp,
) -> pd.DataFrame:
    """Ejecuta el pipeline de transformación completo."""
    df = unir_responsables(df_data, df_resp_grupo, df_resp_mrp)
    df = calcular_solped_mrp(df)
    df = filtrar_solicitudes_vigentes(df)
    df = calcular_nivel_servicio_dias(df, fecha_corte)
    df = unir_centro_sociedad(df, df_centro_sociedad)
    df = calcular_estado_solped_y_nivel_servicio(df)
    df = calcular_aplica(df)
    df = calcular_cumple(df)

    df = df.sort_values("Fecha de solicitud", ascending=False)

    columnas_auxiliares = [
        "_fecha_repor_inicio_sla",
        "_fecha_pedi_inicio_sla",
        "_fecha_repor_solicitud",
        "_fecha_pedi_solicitud",
        "_nivel_servicio_v1",
        "Cantidad solicitada",
        "Unidad de medida",
        "Valor total",
        "Moneda",
        "Indicador liberación",
        "Grupo de artículos",
        "Autor",
        "Concluida",
        "Tipo de posición",
        "Tipo de imputación",
        "Consumo",
        "Pos.solicitud pedido",
        "Responsable de MRP.Responsable Compra.title",
    ]
    df = df.drop(columns=[c for c in columnas_auxiliares if c in df.columns])
    df = df.rename(
        columns={
            "Comprador por Grupo Compras": "Comprador (Grupo de compras)",
            "Comprador por Grupo Compras2": "Comprador por Grupo Compras",
        }
    )

    return df.reset_index(drop=True)


def calcular_metricas_por_grupo(df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """Calcula métricas agrupadas por comprador o centro logístico."""
    d = df.assign(
        _cumple_num=(df["Cumple"] == "Cumple").astype(int),
        _pedido_notna=df["Pedido"].notna().astype(int),
    )
    agg_dict = {
        "Promedio días de gestión": ("Nivel de Servicio", "mean"),
        "% Cumplimiento": ("_cumple_num", "mean"),
        "Pos. OC generadas": ("_pedido_notna", "sum"),
    }
    if "Lead Time Total" in df.columns:
        agg_dict["Promedio Lead Time Total"] = ("Lead Time Total", "mean")

    tabla = d.groupby(group_cols).agg(**agg_dict).reset_index()
    tabla["% Cumplimiento"] = tabla["% Cumplimiento"] * 100
    return tabla


def agregar_fila_total(tabla: pd.DataFrame, df_completo: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """Agrega fila de TOTAL general."""
    n = len(df_completo)
    total = {c: "" for c in group_cols}
    total[group_cols[0]] = "TOTAL"
    total["Promedio días de gestión"] = df_completo["Nivel de Servicio"].mean() if n else float("nan")
    if "Lead Time Total" in df_completo.columns:
        total["Promedio Lead Time Total"] = df_completo["Lead Time Total"].mean() if n else float("nan")
    total["% Cumplimiento"] = (df_completo["Cumple"] == "Cumple").sum() / n * 100 if n else 0
    total["Pos. OC generadas"] = tabla["Pos. OC generadas"].sum()

    return pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)


CENTROS_FIJOS = ["Planta Prillex", "Planta Rio Loa", "Planta Punta Teatinos", "Plantas de Servicio"]


def agregar_grupo_centro(df: pd.DataFrame) -> pd.DataFrame:
    """Crea la columna 'Grupo Centro' con las 4 categorías fijas."""
    df = df.copy()
    nombre = df["Nombre Centro"] if "Nombre Centro" in df.columns else pd.Series(index=df.index, dtype=object)
    df["Grupo Centro"] = nombre.where(nombre.isin(CENTROS_FIJOS[:3]), "Plantas de Servicio")
    return df


def tabla_centros_fija(df: pd.DataFrame) -> pd.DataFrame:
    """Vista fija por centro logístico."""
    d = agregar_grupo_centro(df)
    tabla = calcular_metricas_por_grupo(d, ["Grupo Centro"])
    tabla = (
        tabla.set_index("Grupo Centro")
        .reindex(CENTROS_FIJOS)
        .rename_axis("Centro")
        .reset_index()
    )
    tabla["Pos. OC generadas"] = tabla["Pos. OC generadas"].fillna(0)
    return agregar_fila_total(tabla, d, ["Centro"])
