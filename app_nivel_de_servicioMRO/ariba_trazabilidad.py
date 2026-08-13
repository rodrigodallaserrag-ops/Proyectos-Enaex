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

EMPRESA_POR_DEFECTO = "CEN1"

# Orden de la cadena, de izquierda a derecha
COLUMNAS_CADENA = [
    # --- PR inicial ---
    "PR inicial",
    "Fecha creación PR inicial",
    "Fecha liberación PR inicial",
    "Fecha envío PR inicial",
    "Estado de agregación",
    # --- Toma del comprador ---
    "Fecha asignada al comprador",
    # --- PR agregada ---
    "PR agregada",
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
    "Líneas PR inicial",
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


def construir_cadena(d: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve una fila por PR inicial, con su PR agregada y la salida a SAP
    dispuestas hacia la derecha en orden cronológico.
    """
    d = d.copy()
    d["_agregada"] = _limpiar(d[COL_AGREGADA])
    d["_erp"] = _limpiar(d[COL_ERP])

    # ---------- Bloque A: PR iniciales (se agrupan sus líneas) ----------
    a = d[d[COL_ORIGINAL].notna()].copy()
    a["_f_creacion"] = _fecha(a[COL_F_SOLICITUD])
    a["_f_aprob"] = _fecha(a[COL_F_APROB])
    a["_f_envio"] = _fecha(a[COL_F_ENVIO_ORIG], dayfirst=False)
    a["_f_envio_agr"] = _fecha(a[COL_F_ENVIO_AGR], dayfirst=False)
    a["_f_asignada"] = _fecha(a[COL_F_ASIGNADA], dayfirst=False)

    def _primero(s):
        s = s.dropna()
        return s.iloc[0] if len(s) else np.nan

    iniciales = (
        a.groupby(COL_ORIGINAL)
        .agg(
            **{
                "Fecha creación PR inicial": ("_f_creacion", "min"),
                "Fecha liberación PR inicial": ("_f_aprob", "min"),
                "Fecha envío PR inicial": ("_f_envio", "min"),
                "Fecha asignada al comprador": ("_f_asignada", "min"),
                "Fecha envío PR agregada": ("_f_envio_agr", "min"),
                "PR agregada": ("_agregada", _primero),
                "Estado de agregación": (COL_ESTADO_AGR, _primero),
                "Líneas PR inicial": (COL_ORIGINAL, "size"),
                "Solicitante": (COL_SOLICITANTE, _primero),
                "Grupo de compra": (COL_GRUPO_COMPRA, _primero),
                "Centro de costes": (COL_CECO, _primero),
                "Proveedor": (COL_PROVEEDOR, _primero),
            }
        )
        .reset_index()
        .rename(columns={COL_ORIGINAL: "PR inicial"})
    )

    # ---------- Bloque B: PR agregadas (una fila por PR) ----------
    b = d[d["_erp"].notna()].copy()
    b["_f_creacion"] = _fecha(b[COL_F_SOLICITUD])
    b["_f_aprob"] = _fecha(b[COL_F_APROB])
    b["_f_po"] = _fecha(b[COL_F_PEDIDO])

    agregadas = (
        b.groupby(COL_ID)
        .agg(
            **{
                "Fecha creación PR agregada": ("_f_creacion", "min"),
                "Fecha liberación PR agregada": ("_f_aprob", "min"),
                "Solped SAP (600)": ("_erp", "first"),
                "PO": (COL_PO, "first"),
                "Fecha PO": ("_f_po", "min"),
            }
        )
        .reset_index()
        .rename(columns={COL_ID: "PR agregada"})
    )

    # ---------- Auto-cruce ----------
    out = iniciales.merge(agregadas, on="PR agregada", how="left")

    # ---------- Tiempos de cada tramo ----------
    def _dias(fin, ini):
        return ((out[fin] - out[ini]).dt.total_seconds() / 86400).round(1)

    out["Días espera en panel"] = _dias("Fecha asignada al comprador", "Fecha liberación PR inicial")
    out["Días agregación"] = _dias("Fecha envío PR agregada", "Fecha asignada al comprador")
    out["Días liberación agregada"] = _dias("Fecha liberación PR agregada", "Fecha envío PR agregada")
    out["Días gestión total"] = _dias("Fecha PO", "Fecha liberación PR inicial")

    out["Cadena completa"] = np.where(out["Solped SAP (600)"].notna(), "Sí", "No")

    out = out[[c for c in COLUMNAS_CADENA if c in out.columns]]
    return out.sort_values("Fecha creación PR inicial", ascending=False).reset_index(drop=True)


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
