"""
Procesamiento del reporte 'PR No Catalogadas - Trazabilidad' de Ariba.

El reporte viene partido en dos bloques que NO comparten ninguna fila:

  Bloque A — PR iniciales
      Traen las fechas de trazabilidad y apuntan a su PR agregada,
      pero NO traen el código SAP (600) ni la orden de compra.

  Bloque B — PR agregadas
      Traen el código SAP (600), la PO y sus fechas,
      pero NO traen fechas de trazabilidad.

El puente es un auto-cruce dentro del mismo archivo: la columna
'ID de solicitud de compra agregada' del bloque A coincide con
'ID de solicitud de compra' del bloque B.

Este módulo NO cruza con la ME5A: solo deja el reporte en formato de cadena,
leyéndose de izquierda a derecha en orden cronológico.
"""
import numpy as np
import pandas as pd

# ---- Nombres de columna del reporte original ----
COL_EMPRESA = "[SOCC]Empresa compradora (ID de organización compradora)"
COL_ID = "[SOCC] ID de solicitud de compra"
COL_AGREGADA = "[SOCC] ID de solicitud de compra agregada"
COL_ORIGINAL = "[SOCC] ID de solicitud de compra original"
COL_ERP = "[SOCC] ID de solicitud de compra del ERP"
COL_PO = "[PC] ID de pedido"
COL_F_SOLICITUD = "[SOCC]Fecha de la solicitud de compra (Fecha)"
COL_F_APROB = "[SOCC]Fecha de aprobación (Fecha)"
COL_F_PEDIDO = "[PC]Fecha de pedido (Fecha)"
COL_F_ENVIO_ORIG = "[SOCC] Fecha de envío de la solicitud de compra original"
COL_F_ENVIO_AGR = "[SOCC] Fecha de envío de la solicitud de compra agregada"
COL_F_ASIGNADA = "[SOCC] Fecha asignada"
COL_ESTADO_AGR = "[SOCC] Estado de agregación"
COL_SOLICITANTE = "[SOCC]Solicitante (Usuario)"
COL_PROVEEDOR = "[SOCC]Proveedor (Nombre del proveedor (L1))"
COL_CECO = "[SOCC]Centro de costes (Centro de costes)"
COL_GRUPO_COMPRA = "[SOCC]Grupo de compra (ID de organización compradora)"
COL_LINEA = "[SOCC] Número de línea de la solicitud de compra"
COL_LINEA_ORIG = "[SOCC] Número de línea de la solicitud de compra original"
COL_LINEA_AGR = "[SOCC] Número de línea de la solicitud de compra agregada"

# En SAP la línea 1 corresponde a la posición 10, la 2 a la 20, etc.
FACTOR_POSICION = 10

EMPRESA_POR_DEFECTO = "CEN1"

# Orden de la cadena, de izquierda a derecha
COLUMNAS_CADENA = [
    # --- PR inicial ---
    "PR inicial",
    "Posición PR inicial",
    "Fecha creación PR inicial",
    "Fecha liberación PR inicial",
    "Fecha envío PR inicial",
    "Estado de agregación",
    # --- Toma del comprador ---
    "Fecha asignada al comprador",
    # --- PR agregada ---
    "PR agregada",
    "Posición PR agregada",
    "Fecha creación PR agregada",
    "Fecha envío PR agregada",
    "Fecha liberación PR agregada",
    # --- Salida a SAP ---
    "Solped SAP (600)",
    "PO",
    "Fecha PO",
    # --- Tiempos ---
    "Días espera en panel",
    "Días agregación",
    "Días liberación agregada",
    "Días gestión total",
    # --- Contexto ---
    "Cadena completa",
    "Solicitante",
    "Grupo de compra",
    "Centro de costes",
    "Proveedor",
]


def _limpiar(s: pd.Series) -> pd.Series:
    """'Sin clasificar' y cadenas vacías se tratan como nulo."""
    return s.replace(["Sin clasificar", "", " "], np.nan)


def _fecha(s: pd.Series, dayfirst: bool = True) -> pd.Series:
    return pd.to_datetime(_limpiar(s), dayfirst=dayfirst, errors="coerce")


def cargar_reporte(archivo, empresa: str = EMPRESA_POR_DEFECTO) -> pd.DataFrame:
    """Lee el CSV y deja solo la empresa compradora indicada."""
    df = pd.read_csv(archivo, low_memory=False)
    if COL_EMPRESA not in df.columns:
        raise ValueError(
            f"El archivo no tiene la columna '{COL_EMPRESA}'. "
            "¿Es el reporte 'PR No Catalogadas - Trazabilidad' de Ariba?"
        )
    return df[df[COL_EMPRESA].astype(str).str.strip() == empresa].copy()


def empresas_disponibles(archivo) -> list:
    """Lista de empresas compradoras presentes en el archivo."""
    df = pd.read_csv(archivo, low_memory=False, usecols=[COL_EMPRESA])
    return sorted(df[COL_EMPRESA].dropna().astype(str).str.strip().unique())


def _posicion(serie: pd.Series) -> pd.Series:
    """Convierte el número de línea de Ariba a la posición como se ve en SAP
    (línea 1 -> posición 10, línea 2 -> 20, ...)."""
    return (pd.to_numeric(serie, errors="coerce") * FACTOR_POSICION).astype("Int64")


def construir_cadena(d: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve UNA FILA POR LÍNEA de PR inicial, con su línea correspondiente en
    la PR agregada y la salida a SAP, dispuestas hacia la derecha en orden
    cronológico.
    """
    d = d.copy()
    d["_agregada"] = _limpiar(d[COL_AGREGADA])
    d["_erp"] = _limpiar(d[COL_ERP])

    # ---------- Bloque A: líneas de PR inicial ----------
    a = d[d[COL_ORIGINAL].notna()].copy()

    inicial = pd.DataFrame(
        {
            "PR inicial": a[COL_ORIGINAL],
            "Posición PR inicial": _posicion(a[COL_LINEA_ORIG]),
            "Fecha creación PR inicial": _fecha(a[COL_F_SOLICITUD]),
            "Fecha liberación PR inicial": _fecha(a[COL_F_APROB]),
            "Fecha envío PR inicial": _fecha(a[COL_F_ENVIO_ORIG], dayfirst=False),
            "Estado de agregación": a[COL_ESTADO_AGR],
            "Fecha asignada al comprador": _fecha(a[COL_F_ASIGNADA], dayfirst=False),
            "PR agregada": a["_agregada"],
            "Posición PR agregada": _posicion(a[COL_LINEA_AGR]),
            "Fecha envío PR agregada": _fecha(a[COL_F_ENVIO_AGR], dayfirst=False),
            "Solicitante": a[COL_SOLICITANTE],
            "Grupo de compra": a[COL_GRUPO_COMPRA],
            "Centro de costes": a[COL_CECO],
            "Proveedor": a[COL_PROVEEDOR],
        }
    )

    # ---------- Bloque B: líneas de PR agregada (traen el 600 y la PO) ----------
    b = d[d["_erp"].notna()].copy()
    b_lineas = pd.DataFrame(
        {
            "PR agregada": b[COL_ID],
            "Posición PR agregada": _posicion(b[COL_LINEA]),
            "Fecha creación PR agregada": _fecha(b[COL_F_SOLICITUD]),
            "Fecha liberación PR agregada": _fecha(b[COL_F_APROB]),
            "Solped SAP (600)": b["_erp"],
            "PO": b[COL_PO],
            "Fecha PO": _fecha(b[COL_F_PEDIDO]),
        }
    ).drop_duplicates(subset=["PR agregada", "Posición PR agregada"])

    # Cabecera de la PR agregada: respaldo cuando la línea no calza
    b_cabecera = (
        b_lineas.groupby("PR agregada")
        .agg(
            **{
                "Fecha creación PR agregada_h": ("Fecha creación PR agregada", "min"),
                "Fecha liberación PR agregada_h": ("Fecha liberación PR agregada", "min"),
                "Solped SAP (600)_h": ("Solped SAP (600)", "first"),
                "PO_h": ("PO", "first"),
                "Fecha PO_h": ("Fecha PO", "min"),
            }
        )
        .reset_index()
    )

    # ---------- Auto-cruce: primero por línea, con respaldo por cabecera ----------
    out = inicial.merge(b_lineas, on=["PR agregada", "Posición PR agregada"], how="left")
    out = out.merge(b_cabecera, on="PR agregada", how="left")
    for col in [
        "Fecha creación PR agregada",
        "Fecha liberación PR agregada",
        "Solped SAP (600)",
        "PO",
        "Fecha PO",
    ]:
        out[col] = out[col].fillna(out[f"{col}_h"])
        out = out.drop(columns=[f"{col}_h"])

    # ---------- Tiempos de cada tramo ----------
    def _dias(fin, ini):
        return ((out[fin] - out[ini]).dt.total_seconds() / 86400).round(1)

    out["Días espera en panel"] = _dias("Fecha asignada al comprador", "Fecha liberación PR inicial")
    out["Días agregación"] = _dias("Fecha envío PR agregada", "Fecha asignada al comprador")
    out["Días liberación agregada"] = _dias("Fecha liberación PR agregada", "Fecha envío PR agregada")
    out["Días gestión total"] = _dias("Fecha PO", "Fecha liberación PR inicial")

    out["Cadena completa"] = np.where(out["Solped SAP (600)"].notna(), "Sí", "No")

    out = out[[c for c in COLUMNAS_CADENA if c in out.columns]]
    return out.sort_values(
        ["Fecha creación PR inicial", "PR inicial", "Posición PR inicial"], ascending=[False, True, True]
    ).reset_index(drop=True)


def resumen_por_solped(cadena: pd.DataFrame) -> pd.DataFrame:
    """
    Una fila por Solped SAP (600). Es la tabla que más adelante se cruzará
    con la ME5A del dashboard principal.

    Cuando varias PR iniciales se consolidan en una misma agregada, se toma
    la liberación MÁS ANTIGUA: es cuando arrancó realmente el reloj.
    """
    c = cadena[cadena["Solped SAP (600)"].notna()].copy()
    r = (
        c.groupby("Solped SAP (600)")
        .agg(
            **{
                "PR agregada": ("PR agregada", "first"),
                "Fecha liberación PR inicial": ("Fecha liberación PR inicial", "min"),
                "Fecha asignada al comprador": ("Fecha asignada al comprador", "min"),
                "Fecha liberación PR agregada": ("Fecha liberación PR agregada", "min"),
                "PO": ("PO", "first"),
                "Fecha PO": ("Fecha PO", "min"),
                "PR iniciales consolidadas": ("PR inicial", "nunique"),
            }
        )
        .reset_index()
    )
    r["Días gestión real"] = (
        (r["Fecha PO"] - r["Fecha liberación PR inicial"]).dt.total_seconds() / 86400
    ).round(1)
    # Clave lista para el futuro cruce con ME5A
    r["Solicitud de pedido"] = pd.to_numeric(r["Solped SAP (600)"], errors="coerce").astype("Int64")
    return r.sort_values("Solicitud de pedido").reset_index(drop=True)


# ==============================================================================
# FUNCIÓN AÑADIDA PARA INTEGRACIÓN DIRECTA CON STREAMLIT / APP.PY
# ==============================================================================
def procesar_trazabilidad_completa(archivo, empresa: str = EMPRESA_POR_DEFECTO):
    """
    Procesa el archivo sucio ingresado en la UI de Streamlit.
    Devuelve:
      - df_cadena: Trazabilidad completa paso a paso.
      - df_resumen: Solpeds SAP 600 identificadas como NO CATALOGADAS.
    """
    raw = cargar_reporte(archivo, empresa=empresa)
    df_cadena = construir_cadena(raw)
    df_resumen = resumen_por_solped(df_cadena)
    return df_cadena, df_resumen


if __name__ == "__main__":
    import sys

    ruta = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "/mnt/user-data/uploads/Reporte_PR_No_Catalogadas_-_Trazabilidad__3_.csv"
    )
    d = cargar_reporte(ruta)
    cadena = construir_cadena(d)
    res = resumen_por_solped(cadena)
    print(f"PR iniciales:        {len(cadena):,}")
    print(f"Con cadena completa: {(cadena['Cadena completa']=='Sí').sum():,}")
    print(f"Solpeds 600 únicas:  {len(res):,}")
