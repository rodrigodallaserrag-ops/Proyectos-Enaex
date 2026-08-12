"""
Transformaciones - traducción 1:1 del código M (Power Query) de la query
"Data (2)" del pbix Nivel_de_servicio_BI.pbix.

Cada función de acá corresponde a uno o más pasos del "let...in" original.
Se mantiene el orden exacto porque varios pasos dependen de columnas
intermedias que después se descartan (igual que en el M).
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

    Traducción literal del M:
        Personalizado   = 2 si (Solped MRP in [MRP, Ariba]) o (ERP y sin indicador), si no 0
        Personalizado.3 = si indicador vacío/nulo -> Personalizado
                          si indicador "X" o "B"  -> 0
                          en cualquier otro caso  -> el indicador convertido a número
        Filtro final: conservar solo las filas donde Personalizado.3 = 2

    OJO: el último "else" es clave. Un indicador con valor "2" hace que la fila
    se conserve aunque sea ERP, porque su valor numérico es exactamente 2.
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
    Pasos M: 'Fecha Repor-Fecha mod', 'Fecha Pedi-Fech Modi',
    'Personalizada agregada5' (Nivel de Servicio v1).
    """
    df = df.copy()
    df["_fecha_repor_fecha_mod"] = _dias_entre(pd.Timestamp(fecha_corte), df["Fecha modificación"])
    df["_fecha_pedi_fecha_modi"] = _dias_entre(df["Fecha de pedido"], df["Fecha modificación"])

    sin_pedido = df["Pedido"].isna()
    df["_nivel_servicio_v1"] = np.where(
        sin_pedido, df["_fecha_repor_fecha_mod"], df["_fecha_pedi_fecha_modi"]
    )
    return df


def unir_centro_sociedad(df: pd.DataFrame, df_centro_sociedad: pd.DataFrame) -> pd.DataFrame:
    """
    Paso M: 'Consultas combinadas2' + expandir + rename (OJO: el M
    intercambia 'Nombre Centro' <-> 'Nombre Centro 2' al renombrar,
    se replica igual).
    """
    df = df.merge(df_centro_sociedad, left_on="Centro", right_on="Título", how="left").drop(
        columns=["Título"]
    )
    df = df.rename(columns={"Nombre Centro": "Nombre Centro 2", "Nombre Centro 2": "Nombre Centro"})
    return df


def calcular_estado_solped_y_nivel_servicio(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pasos M: 'Personalizada agregada6' (Estado Solped), 'Personalizada
    agregada7'/'Columnas con nombre cambiado5' (Nivel de Servicio final).
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
        df["_fecha_repor_fecha_mod"],
        df["_nivel_servicio_v1"],
    )
    return df


def calcular_aplica(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pasos M: 'Personalizada agregada8' (chequeo solo para 'Sin pedido'),
    'Personalizada agregada9' -> columna 'Aplica?'.
    """
    df = df.copy()
    check = np.where(df["Estado Solped"] == "Sin pedido", df["Nivel de Servicio"], np.nan)
    df["Aplica?"] = np.where(
        pd.isna(check) | (check >= config.DIAS_GRACIA_SIN_PEDIDO), "Aplica", "No aplica"
    )
    return df


def calcular_cumple(df: pd.DataFrame) -> pd.DataFrame:
    """Paso M: 'Personalizada agregada10' -> columna 'Cumple'. SLA: 10 días
    para ERP/MRP, 7 días para Ariba."""
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
    """Corre el pipeline completo, en el mismo orden que el 'let...in' del M."""
    df = unir_responsables(df_data, df_resp_grupo, df_resp_mrp)
    df = calcular_solped_mrp(df)
    df = filtrar_solicitudes_vigentes(df)
    df = calcular_nivel_servicio_dias(df, fecha_corte)
    df = unir_centro_sociedad(df, df_centro_sociedad)
    df = calcular_estado_solped_y_nivel_servicio(df)
    df = calcular_aplica(df)
    df = calcular_cumple(df)

    df = df.sort_values("Fecha de solicitud", ascending=False)

    # Limpieza final (equivalente a 'Columnas quitadas2' / 'Columnas quitadas3')
    # OJO: se conservan LAS DOS columnas de comprador:
    #   - "Comprador por Grupo Compras"  -> la del pbix, con override de MRP
    #   - "Comprador (Grupo de compras)" -> la pura, según el grupo de compras,
    #     que reparte las líneas MRP entre los compradores reales en vez de
    #     concentrarlas en el responsable de MRP.
    columnas_auxiliares = [
        "_fecha_repor_fecha_mod",
        "_fecha_pedi_fecha_modi",
        "_nivel_servicio_v1",
        "Fecha de liberación",
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
    """
    Reemplaza las medidas DAX del pbix, agregadas a nivel de comprador/centro/etc.
    Fórmulas confirmadas y validadas al 100% contra el pbix (ver conversación):

        % Cumplimiento          = DIVIDE(COUNTROWS(Cumple="Cumple"), COUNTROWS(todo))
        Promedio días de gestión = AVERAGE('Data (2)'[Nivel de servicio])
        Pos. OC generadas        = POSICIONES con pedido, SIN quitar duplicados,
                                   para que la columna sume verticalmente.

    OJO: la tarjeta general usa recuento DISTINTIVO (quita duplicados) porque
    ahí interesa cuántas OC distintas se generaron. Acá interesa cuántas
    posiciones gestionó cada comprador, así que se cuentan todas las líneas.
    """
    def _agg(g):
        return pd.Series(
            {
                "Promedio días de gestión": g["Nivel de Servicio"].mean(),
                "% Cumplimiento": (g["Cumple"] == "Cumple").sum() / max(len(g), 1) * 100,
                "Pos. OC generadas": g["Pedido"].notna().sum(),
            }
        )

    return df.groupby(group_cols).apply(_agg, include_groups=False).reset_index()


def agregar_fila_total(tabla: pd.DataFrame, df_completo: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """
    Agrega una fila TOTAL al final de la tabla, calculada sobre TODAS las líneas
    del conjunto filtrado (no promediando los promedios de cada fila).

    Esto importa porque:
      - % Cumplimiento total = líneas que cumplen / total de líneas.
        NO es el promedio de los porcentajes de cada comprador (eso daría el
        mismo peso a compradores con volúmenes muy distintos).
      - Promedio días total = promedio sobre todas las líneas, por el mismo motivo.
      - Pos. OC total = SUMA de la columna (criterio definido por el usuario).
        Ojo: puede quedar por encima del recuento distintivo global si un mismo
        pedido aparece en más de una fila de la tabla.
    """
    n = len(df_completo)
    total = {c: "" for c in group_cols}
    total[group_cols[0]] = "TOTAL"
    total["Promedio días de gestión"] = df_completo["Nivel de Servicio"].mean() if n else float("nan")
    total["% Cumplimiento"] = (df_completo["Cumple"] == "Cumple").sum() / n * 100 if n else 0
    total["Pos. OC generadas"] = tabla["Pos. OC generadas"].sum()

    return pd.concat([tabla, pd.DataFrame([total])], ignore_index=True)
