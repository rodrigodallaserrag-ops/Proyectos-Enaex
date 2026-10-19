# =============================================================================
# CUADRO COMPARATIVO DE OFERTAS - ABASTECIMIENTO ENAEX  (v3)
# -----------------------------------------------------------------------------
# Flujo:
#   1. Bandeja: al cargar la planilla se ven todas las SOLPED con su material.
#      Vista por material -> "todas las SOLPED con este material" -> agrupar.
#   2. Solicitud de cotización (Excel) para enviar al proveedor, sin precios
#      históricos internos. La respuesta del proveedor se importa de vuelta.
#   3. Comparación (total puesto en planta, CLP) -> sugerencia por ítem.
#   4. Adjudicación con aprobación del comprador -> Excel / PDF / borrador.
#
# requirements.txt:
#   streamlit>=1.35  pandas>=2.1  numpy  openpyxl  requests  fpdf2
#   python-calamine (opcional, lee Excel ~10x más rápido)
#   holidays        (opcional, feriados de Chile en fechas de entrega)
# =============================================================================
import io
import json
import re
import unicodedata
from datetime import date, datetime, timedelta

import numpy as np
import openpyxl
import pandas as pd
import requests
import streamlit as st
from fpdf import FPDF
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import python_calamine  # noqa: F401
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

try:
    import holidays as _holidays
    FERIADOS = np.array(
        sorted(_holidays.Chile(years=range(date.today().year - 1, date.today().year + 3)).keys()),
        dtype="datetime64[D]",
    )
except Exception:
    FERIADOS = np.array([], dtype="datetime64[D]")

# =============================================================================
# CONSTANTES DE NEGOCIO (de las planillas Cuadro MRO Rev3 y Autogestión)
# =============================================================================
MONEDAS = ["CLP", "USD", "EUR", "UF"]
TRANSPORTES = ["T. Gil", "T. Bello", "Pullman", "Retiramos", "EXW", "FCA", "FOB",
               "CFR", "CIF", "CPT", "CIP", "DAP", "DDP"]
TIPOS_ADJUDICACION = ["Mejor Oferta Económica", "Proveedor único",
                      "Aprobado Técnicamente", "Asignación Directa"]
TIPOS_PLAZO = ["Días hábiles", "Días corridos", "Semanas"]
TIPOS_COMPRA = ["Nacional", "Importación"]
LUGARES_ENTREGA = {                       # hoja "Fecha de entrega" de Autogestión
    "EDP - Entrega directa en planta": 0,
    "EBP - Entrega en bodega de proveedor": 0,
    "EBC - Entrega bodega consolidación": 7,
    "EXW A - Origen aéreo": 15,
    "EXW M - Origen marítimo": 45,
}
DIAS_EMISION_OC = 2
DIAS_INTERNACION = 10
DIAS_BODEGA_TRANSITO = 7
UMBRAL_TC_CLP = 1_000_000                 # >$1.000.000 CLP exige TC del día
UMBRAL_ALZA = 0.15                        # alerta si sube >15% vs última compra
SIN_ADJ = "— Sin adjudicar —"

# Solicitud de cotización (RFQ): campos que completa el proveedor
RFQ_HOJA = "Solicitud de cotización"
RFQ_CAMPOS = ["Razón social", "RUT", "Contacto (nombre y correo)", "Moneda de la oferta",
              "Costo total de flete", "Incoterm / forma de despacho",
              "Plazo de entrega general (días)", "Validez de la oferta (días)", "Condición de pago"]
RFQ_COLS = ["Ítem", "Código SAP", "Descripción", "Cantidad", "UM", "Lugar de entrega",
            "Precio unitario (sin IVA)", "Plazo entrega (días)", "Marca / modelo ofrecido",
            "Cumple especificación (Sí/No)", "Observaciones"]
RFQ_EDITABLES = RFQ_COLS[6:]
RFQ_FILA_CAMPOS = 9
RFQ_FILA_TABLA = RFQ_FILA_CAMPOS + len(RFQ_CAMPOS) + 2

# Sinónimos de columnas (comparación exacta tras normalizar)
SIN_REQ = {
    "solped": ["sp", "solped", "n solped", "solicitud de pedido", "sol pedido",
               "id solicitud de compra original", "solicitud pedido", "pr"],
    "pos": ["pos", "posicion", "pos solicitud pedido", "pos sol ped", "item"],
    "material": ["material", "codigo sap", "codigo material", "cod material", "codigo de articulo"],
    "texto": ["texto breve", "descripcion breve", "texto breve de material", "descripcion",
              "denominacion", "texto breve material"],
    "cantidad": ["cantidad", "ctd", "cantidad solicitada", "cant"],
    "um": ["um", "unidad", "unidad de medida", "unidad medida pedido", "umb"],
    "centro": ["centro", "planta"],
    "precio": ["precio", "precio unitario", "valor unitario", "precio de valoracion"],
    "moneda": ["moneda", "mon"],
    "fecha_sol": ["f solicitud", "fecha de solicitud", "fecha solicitud"],
    "dias": ["dias de tratamiento", "ns", "dias", "antiguedad"],
    "urgencia": ["urgencia", "prioridad"],
    "grupo": ["g compras", "grupo compras", "grupo de compras"],
    "solicitante": ["solicitante", "creado por"],
    "asignado": ["asignado a", "comprador"],
    "contrato": ["contrato marco", "contrato"],
}
SIN_HIST = {
    "material": ["material", "codigo sap", "codigo material"],
    "texto": ["texto breve", "descripcion", "texto breve de material"],
    "fecha": ["f documento", "fecha documento", "fecha de pedido", "fecha doc", "fecha"],
    "proveedor": ["proveedor", "proveedor centro suministrador", "nombre proveedor", "acreedor"],
    "precio": ["valor unitario", "precio unitario", "precio neto", "precio"],
    "valor_neto": ["valor neto de pedido", "valor neto", "importe"],
    "cantidad": ["cantidad de pedido", "cantidad", "ctd pedido"],
    "moneda": ["moneda", "mon"],
    "documento": ["documento", "documento compras", "pedido", "oc"],
}
REQ_OBLIGATORIAS = ["solped", "texto", "cantidad"]
HIST_OBLIGATORIAS = ["material", "fecha", "precio"]
COLS_HIST = ["Últ. compra", "Últ. proveedor", "Últ. precio", "Últ. moneda", "Últ. OC", "N° compras"]

# =============================================================================
# PÁGINA
# =============================================================================
st.set_page_config(page_title="Cuadro Comparativo de Ofertas - Enaex", layout="wide")
st.markdown("""
<style>
  [data-testid="stMainBlockContainer"] { padding-bottom: 6rem; max-width: 100%; }
  .titulo { font-size: 1.7rem; font-weight: 700; color: #1E3A8A; margin-bottom: .2rem; }
  .subtitulo { color: #4B5563; margin-bottom: 1rem; }
  div.stButton > button, div.stDownloadButton > button { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# UTILIDADES
# =============================================================================
def norm(txt) -> str:
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", " ", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def clave_material(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s).lstrip("0")
    return s if len(s) >= 4 else ""


def clave_item(material, texto) -> str:
    """Agrupa por código SAP; si no hay código (Ariba), por descripción normalizada."""
    return material if material else "TXT:" + norm(texto)


def parse_monto(v):
    """'$49.100CLP', '1.234,56', '389,00 USD', 59500 -> (valor, moneda|None)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0, None
    if isinstance(v, (int, float, np.number)):
        return float(v), None
    s = str(v).strip().upper()
    moneda = next((m for m in MONEDAS if m in s), None)
    s = re.sub(r"[^\d.,-]", "", s)
    if not s:
        return 0.0, moneda
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) != 3 else s.replace(",", "")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[-1]) == 3):
        s = s.replace(".", "")
    try:
        return float(s), moneda
    except ValueError:
        return 0.0, moneda


def a_num(v, defecto=0.0) -> float:
    val, _ = parse_monto(v)
    return val if val else defecto


def unir_unicos(serie) -> str:
    return ", ".join(dict.fromkeys(str(x) for x in serie if str(x).strip() not in ("", "nan", "None")))


def mapear_columnas(columnas, sinonimos) -> dict:
    normalizadas = {c: norm(c) for c in columnas}
    usadas, mapa = set(), {}
    for campo, sins in sinonimos.items():
        encontrada = None
        for sin in sins:
            encontrada = next((c for c, n in normalizadas.items() if n == sin and c not in usadas), None)
            if encontrada:
                break
        if not encontrada:
            for sin in sins:
                if len(sin) >= 5:
                    encontrada = next((c for c, n in normalizadas.items()
                                       if n.startswith(sin) and c not in usadas), None)
                    if encontrada:
                        break
        if encontrada:
            mapa[campo] = encontrada
            usadas.add(encontrada)
    return mapa


def deduplicar(columnas):
    vistos, salida = {}, []
    for c in columnas:
        c = str(c).strip() if c is not None and str(c).strip() not in ("", "nan", "None") else "Col"
        if c in vistos:
            vistos[c] += 1
            salida.append(f"{c}_{vistos[c]}")
        else:
            vistos[c] = 0
            salida.append(c)
    return salida

# =============================================================================
# LECTURA DE ARCHIVOS
# =============================================================================
@st.cache_data(show_spinner=False)
def listar_hojas(file_bytes: bytes, nombre: str):
    if nombre.lower().endswith(".csv"):
        return ["CSV"]
    return pd.ExcelFile(io.BytesIO(file_bytes), engine=EXCEL_ENGINE).sheet_names


def puntaje_hoja(nombre_hoja: str, preferidas) -> int:
    n = norm(nombre_hoja)
    for i, p in enumerate(preferidas):
        if n == norm(p):
            return 100 - i
    return 0


@st.cache_data(show_spinner="Leyendo planilla…")
def leer_hoja(file_bytes: bytes, nombre: str, hoja: str, sinonimos_json: str) -> pd.DataFrame:
    """Lee la hoja y ubica la fila de encabezados en las primeras 40 filas."""
    sinonimos = json.loads(sinonimos_json)
    if nombre.lower().endswith(".csv"):
        raw = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=object, sep=None, engine="python")
    else:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=hoja, header=None,
                            dtype=object, engine=EXCEL_ENGINE)
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if raw.empty:
        return raw
    todos = {s for sins in sinonimos.values() for s in sins}
    mejor_fila, mejor_puntaje = 0, -1
    for i in range(min(40, len(raw))):
        celdas = {norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)}
        puntaje = len(celdas & todos)
        if puntaje > mejor_puntaje:
            mejor_fila, mejor_puntaje = i, puntaje
    df = raw.iloc[mejor_fila + 1:].copy()
    df.columns = deduplicar(raw.iloc[mejor_fila].tolist())
    df = df.loc[:, [c for c in df.columns if not c.startswith("Col")]]
    return df.dropna(how="all").reset_index(drop=True)


def normalizar_requerimientos(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """Lleva la carga (ME5A o Ariba) a una fila por posición de SOLPED."""
    g = lambda c: df[mapa[c]] if c in mapa else pd.Series([None] * len(df), index=df.index)
    out = pd.DataFrame({
        "SOLPED": g("solped").astype(str).str.strip().str.replace(r"\.0$", "", regex=True),
        "Pos": g("pos"),
        "Material": g("material").apply(clave_material),
        "Texto": g("texto").astype(str).str.strip().str.strip('"').replace({"None": "", "nan": ""}),
        "Cantidad": g("cantidad").apply(lambda v: a_num(v, 1.0)),
        "UM": g("um").fillna("C/U").astype(str).str.strip(),
        "Centro": g("centro").fillna("").astype(str),
        "Precio ref.": g("precio").apply(lambda v: parse_monto(v)[0]),
        "F. solicitud": pd.to_datetime(g("fecha_sol"), errors="coerce", dayfirst=True),
        "Días": pd.to_numeric(g("dias"), errors="coerce"),
        "Urgencia": g("urgencia").fillna("").astype(str),
        "G. compras": g("grupo").fillna("").astype(str),
        "Asignado a": g("asignado").fillna("").astype(str),
        "Contrato marco": g("contrato").fillna("").astype(str),
    })
    out = out[~out["SOLPED"].isin(["", "None", "nan"]) & (out["Texto"] != "")]
    sin_pos = out["Pos"].isna() | out["Pos"].astype(str).str.strip().isin(["", "None", "nan"])
    out.loc[sin_pos, "Pos"] = out[sin_pos].groupby("SOLPED").cumcount().add(1).mul(10)
    out["Pos"] = out["Pos"].apply(lambda v: int(a_num(v, 0)))
    if out["Días"].isna().all() and out["F. solicitud"].notna().any():
        out["Días"] = (pd.Timestamp(date.today()) - out["F. solicitud"]).dt.days
    out["Clave"] = [clave_item(m, t) for m, t in zip(out["Material"], out["Texto"])]
    return out.drop_duplicates(subset=["SOLPED", "Pos"]).reset_index(drop=True)


@st.cache_data(show_spinner="Indexando histórico de compras…")
def construir_historico(df: pd.DataFrame, mapa: dict):
    """Devuelve (última compra por código SAP, última compra por descripción)."""
    g = lambda c: df[mapa[c]] if c in mapa else pd.Series([None] * len(df), index=df.index)
    h = pd.DataFrame({
        "mat_key": g("material").apply(clave_material),
        "txt_key": g("texto").apply(norm),
        "Fecha": pd.to_datetime(g("fecha"), errors="coerce", dayfirst=True),
        "Proveedor": g("proveedor").fillna("").astype(str).str.replace(r"^\d+\s+", "", regex=True).str.strip(),
        "Precio": g("precio").apply(lambda v: parse_monto(v)[0]),
        "Moneda": g("moneda").fillna("CLP").astype(str).str.upper().str.strip(),
        "OC": g("documento").fillna("").astype(str),
    })
    if "precio" not in mapa and {"valor_neto", "cantidad"} <= set(mapa):
        cant = g("cantidad").apply(lambda v: a_num(v, 0))
        h["Precio"] = np.where(cant > 0, g("valor_neto").apply(lambda v: a_num(v, 0)) / cant.replace(0, np.nan), 0)
    h = h[(h["Precio"] > 0) & h["Fecha"].notna()].sort_values("Fecha", ascending=False)

    def ultima(clave):
        sub = h[h[clave] != ""]
        n = sub.groupby(clave).size().rename("N° compras")
        return sub.drop_duplicates(clave).set_index(clave).join(n)
    return ultima("mat_key"), ultima("txt_key")


def adjuntar_historico(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas de última compra a un DataFrame con 'Material' y 'Texto'."""
    df = df.copy()
    if ss.get("hist") is None:
        for c in COLS_HIST:
            df[c] = np.nan if c in ("Últ. precio", "N° compras") else ""
        df["Últ. compra"] = None
        df["Con historial"], df["Historial por"] = False, ""
        return df
    ult_mat, ult_txt = ss.hist
    txt = df["Texto"].apply(norm)
    por_codigo = df["Material"].isin(ult_mat.index) & (df["Material"] != "")
    por_texto = ~por_codigo & txt.isin(ult_txt.index)
    for destino, origen in [("Últ. compra", "Fecha"), ("Últ. proveedor", "Proveedor"), ("Últ. precio", "Precio"),
                            ("Últ. moneda", "Moneda"), ("Últ. OC", "OC"), ("N° compras", "N° compras")]:
        v_cod = df["Material"].map(ult_mat[origen])
        v_txt = txt.map(ult_txt[origen])
        df[destino] = v_cod.where(por_codigo, v_txt.where(por_texto))
    df["Últ. compra"] = pd.to_datetime(df["Últ. compra"], errors="coerce").dt.date
    df["N° compras"] = df["N° compras"].fillna(0).astype(int)
    df["Con historial"] = por_codigo | por_texto
    df["Historial por"] = np.where(por_codigo, "Código", np.where(por_texto, "Descripción", ""))
    return df

# =============================================================================
# CONSOLIDACIÓN DE LÍNEAS (varias SOLPED -> ítems a cotizar)
# =============================================================================
def consolidar_lineas(lineas: pd.DataFrame, agrupar: bool = True):
    """Asigna 'Ítem' a cada línea. Si agrupar=True, mismo material + misma UM = un ítem."""
    l = lineas.copy().reset_index(drop=True)
    if "Clave" not in l:
        l["Clave"] = [clave_item(m, t) for m, t in zip(l["Material"], l["Texto"])]
    l["UM"] = l["UM"].fillna("C/U").astype(str).str.strip()
    if agrupar:
        l["Ítem"] = l.groupby(["Clave", "UM"], sort=False).ngroup() + 1
    else:
        l["Ítem"] = np.arange(1, len(l) + 1)
    l["Ref"] = l["SOLPED"].astype(str) + "/" + l["Pos"].astype(str)
    items = (l.groupby("Ítem", sort=True)
             .agg(Material=("Material", "first"),
                  Texto=("Texto", lambda s: max(s.astype(str), key=len)),
                  Cantidad=("Cantidad", "sum"),
                  UM=("UM", "first"),
                  Centro=("Centro", unir_unicos),
                  SOLPEDs=("Ref", lambda s: ", ".join(s)),
                  **{"N° SOLPED": ("SOLPED", "nunique")})
             .reset_index())
    return l, items

# =============================================================================
# TIPOS DE CAMBIO Y FECHAS
# =============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def obtener_indicadores():
    try:
        r = requests.get("https://mindicador.cl/api", timeout=6)
        if r.ok:
            d = r.json()
            return {"USD": float(d["dolar"]["valor"]), "EUR": float(d["euro"]["valor"]),
                    "UF": float(d["uf"]["valor"]), "fecha": str(d["dolar"]["fecha"])[:10],
                    "fuente": "mindicador.cl (Banco Central)", "ok": True}
    except Exception:
        pass
    return {"USD": 0.0, "EUR": 0.0, "UF": 0.0, "fecha": "", "fuente": "Sin conexión", "ok": False}


def a_clp(monto, moneda, tc) -> float:
    moneda = str(moneda or "CLP").upper().strip()
    factor = {"CLP": 1.0, "USD": tc["USD"], "EUR": tc["EUR"], "UF": tc["UF"]}.get(moneda, 1.0)
    return float(monto or 0) * factor


def sumar_habiles(f, n: int):
    return np.busday_offset(f, int(n), roll="forward", holidays=FERIADOS)


def calcular_fecha_planta(plazo, tipo_plazo, tipo_compra, lugar, bodega_transito, hoy=None):
    """Réplica de la hoja 'Fecha de entrega' de Autogestión."""
    hoy = np.datetime64(hoy or date.today(), "D")
    plazo = int(a_num(plazo, 0))
    emision = sumar_habiles(hoy, DIAS_EMISION_OC)
    if tipo_plazo == "Días hábiles":
        f_prov = sumar_habiles(emision, plazo)
    elif tipo_plazo == "Semanas":
        f_prov = emision + np.timedelta64(7 * plazo, "D")
    else:
        f_prov = emision + np.timedelta64(plazo, "D")
    extra = (DIAS_INTERNACION if tipo_compra == "Importación" else 0) \
        + LUGARES_ENTREGA.get(lugar, 0) + (DIAS_BODEGA_TRANSITO if bodega_transito else 0)
    f_planta = sumar_habiles(sumar_habiles(f_prov, 2) + np.timedelta64(extra, "D"), 1)
    return pd.Timestamp(f_prov).date(), pd.Timestamp(f_planta).date(), int((f_planta - hoy).astype(int))

# =============================================================================
# MOTOR DE COMPARACIÓN (por ítem)
# =============================================================================
def calcular_comparativo(items, prov, ofertas, tc):
    prov = prov.set_index("Proveedor")
    det = ofertas.drop(columns=["Texto", "Cantidad", "UM"], errors="ignore").merge(
        items[["Ítem", "Cantidad", "Últ. precio CLP"]], on="Ítem", how="left")
    det["Precio mejorado"] = det["Precio mejorado"].where(det["Precio mejorado"].fillna(0) > 0, det["Precio inicial"])
    det = det[det["Proveedor"].isin(prov.index)].copy()
    det["Moneda"] = det["Proveedor"].map(prov["Moneda"])
    det["Cotiza"] = det["Precio mejorado"].fillna(0) > 0
    det["Subtotal inicial"] = det["Cantidad"] * det["Precio inicial"].fillna(0)
    det["Subtotal mejorado"] = det["Cantidad"] * det["Precio mejorado"].fillna(0)
    det["Subtotal CLP"] = [a_clp(m, mo, tc) for m, mo in zip(det["Subtotal mejorado"], det["Moneda"])]
    det["Ahorro spot CLP"] = [a_clp(i - m, mo, tc) if i > 0 else 0.0
                              for i, m, mo in zip(det["Subtotal inicial"], det["Subtotal mejorado"], det["Moneda"])]
    det["Flete CLP"] = 0.0
    for p, fila in prov.iterrows():
        flete_clp = a_clp(fila.get("Flete total", 0), fila.get("Moneda flete", "CLP"), tc)
        m = (det["Proveedor"] == p) & det["Cotiza"]
        base = det.loc[m, "Subtotal CLP"].sum()
        if flete_clp > 0 and base > 0:
            det.loc[m, "Flete CLP"] = flete_clp * det.loc[m, "Subtotal CLP"] / base
    det["Total CLP"] = det["Subtotal CLP"] + det["Flete CLP"]
    det["Precio unit. CLP"] = det["Subtotal CLP"] / det["Cantidad"].replace(0, np.nan)
    det["Unit. puesto CLP"] = det["Total CLP"] / det["Cantidad"].replace(0, np.nan)
    ult = det["Últ. precio CLP"].where(det["Últ. precio CLP"] > 0)
    det["Var. vs últ. compra"] = (det["Precio unit. CLP"] - ult) / ult
    fechas = []
    for _, r in det.iterrows():
        cfg = prov.loc[r["Proveedor"]]
        plazo = r["Plazo"] if pd.notna(r.get("Plazo")) and a_num(r.get("Plazo")) > 0 else cfg.get("Plazo", 0)
        _, f_planta, dias = calcular_fecha_planta(plazo, cfg.get("Tipo plazo"), cfg.get("Tipo compra"),
                                                  cfg.get("Lugar entrega"), bool(cfg.get("Bodega tránsito")))
        fechas.append((f_planta, dias))
    det["Entrega planta"] = [f for f, _ in fechas]
    det["Días a planta"] = [d for _, d in fechas]
    det["Válida"] = det["Cotiza"] & det["Cumple técnico"].fillna(False).astype(bool)

    filas = []
    for _, it in items.iterrows():
        d = det[(det["Ítem"] == it["Ítem"]) & det["Válida"]].sort_values(["Total CLP", "Días a planta"])
        n_validas = len(d)
        mejor = d.iloc[0] if n_validas else None
        alertas = []
        if n_validas == 0:
            alertas.append("Sin ofertas válidas")
        if not it["Con historial"]:
            alertas.append("Sin historial: requiere validación técnica")
        if mejor is not None and pd.notna(mejor["Var. vs últ. compra"]) and mejor["Var. vs últ. compra"] > UMBRAL_ALZA:
            alertas.append(f"Alza {mejor['Var. vs últ. compra']:.0%} vs última compra")
        filas.append({
            "Ítem": it["Ítem"],
            "Sugerido": mejor["Proveedor"] if mejor is not None else SIN_ADJ,
            "Tipo sugerido": ("Proveedor único" if n_validas == 1 else "Mejor Oferta Económica") if n_validas else "",
            "Mejor total CLP": mejor["Total CLP"] if mejor is not None else np.nan,
            "Ofertas válidas": n_validas,
            "Alertas": "; ".join(alertas),
        })
    return det, pd.DataFrame(filas)

# =============================================================================
# SOLICITUD DE COTIZACIÓN PARA EL PROVEEDOR (y lectura de su respuesta)
# =============================================================================
AZUL, VERDE, AMARILLO = "1E3A8A", "D1FAE5", "FEF9C3"
BORDE = Border(*(Side(style="thin", color="D1D5DB"),) * 4)


def generar_rfq(ref, items, lineas, cfg, incluir_detalle=True) -> bytes:
    """Excel para el proveedor: ítems a cotizar + celdas amarillas a completar.
    No incluye precios históricos ni información interna de costos."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = RFQ_HOJA
    ws["A1"] = "ENAEX - SOLICITUD DE COTIZACIÓN"
    ws["A1"].font = Font(size=15, bold=True, color=AZUL)
    ws["A2"] = f"Referencia: {ref}"
    ws["A3"] = f"Fecha de emisión: {date.today():%d/%m/%Y}"
    ws["A4"] = f"Responder antes de: {cfg['fecha_limite']:%d/%m/%Y}"
    ws["A5"] = f"Comprador: {cfg['comprador']}" + (f"  ·  {cfg['correo']}" if cfg.get("correo") else "")
    ws["A6"] = f"Entrega: {cfg['entrega']}" if cfg.get("entrega") else ""
    ws["A7"] = ("Instrucciones: complete solo las celdas amarillas, no cambie encabezados ni el orden de las filas "
                "y devuelva este mismo archivo en Excel. Precios unitarios netos, sin IVA.")
    ws["A7"].font = Font(italic=True, color="6B7280")
    for i, campo in enumerate(RFQ_CAMPOS):
        r = RFQ_FILA_CAMPOS + i
        ws.cell(r, 1, campo).font = Font(bold=True)
        celda = ws.cell(r, 2)
        celda.fill = PatternFill("solid", start_color=AMARILLO)
        celda.border = BORDE
    dv_mon = DataValidation(type="list", formula1='"' + ",".join(MONEDAS) + '"', allow_blank=True)
    dv_inc = DataValidation(type="list", formula1='"' + ",".join(TRANSPORTES) + '"', allow_blank=True)
    dv_si = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
    for dv in (dv_mon, dv_inc, dv_si):
        ws.add_data_validation(dv)
    dv_mon.add(ws.cell(RFQ_FILA_CAMPOS + RFQ_CAMPOS.index("Moneda de la oferta"), 2))
    dv_inc.add(ws.cell(RFQ_FILA_CAMPOS + RFQ_CAMPOS.index("Incoterm / forma de despacho"), 2))

    h = RFQ_FILA_TABLA
    for j, col in enumerate(RFQ_COLS, start=1):
        c = ws.cell(h, j, col)
        c.fill = PatternFill("solid", start_color=AZUL if col not in RFQ_EDITABLES else "B45309")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    for i, (_, it) in enumerate(items.iterrows(), start=1):
        r = h + i
        valores = [int(it["Ítem"]), it["Material"] or "", it["Texto"], float(it["Cantidad"]), it["UM"], it["Centro"]]
        for j, v in enumerate(valores, start=1):
            ws.cell(r, j, v).border = BORDE
        for j in range(len(valores) + 1, len(RFQ_COLS) + 1):
            c = ws.cell(r, j)
            c.fill = PatternFill("solid", start_color=AMARILLO)
            c.border = BORDE
        ws.cell(r, 7).number_format = "#,##0.00"
        dv_si.add(ws.cell(r, RFQ_COLS.index("Cumple especificación (Sí/No)") + 1))
    anchos = [7, 13, 50, 10, 9, 24, 18, 14, 26, 16, 34]
    for j, w in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.column_dimensions["A"].width = 30
    ws.row_dimensions[h].height = 32
    ws.freeze_panes = ws.cell(h + 1, 1)

    if incluir_detalle:
        wd = wb.create_sheet("Detalle por SOLPED")
        det = lineas[["SOLPED", "Pos", "Ítem", "Material", "Texto", "Cantidad", "UM", "Centro"]]
        wd.append(["SOLPED", "Pos", "Ítem", "Código SAP", "Descripción", "Cantidad", "UM", "Centro"])
        for fila in det.itertuples(index=False):
            wd.append(list(fila))
        for c in wd[1]:
            c.fill = PatternFill("solid", start_color=AZUL)
            c.font = Font(bold=True, color="FFFFFF")
        for j, w in enumerate([14, 7, 7, 13, 50, 10, 9, 24], start=1):
            wd.column_dimensions[get_column_letter(j)].width = w

    meta = wb.create_sheet("_meta")
    meta["A1"], meta["A2"] = "ENAEX_RFQ_V1", ref
    meta.sheet_state = "hidden"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def leer_respuesta_rfq(file_bytes: bytes):
    """Lee la solicitud devuelta por el proveedor. Retorna (referencia, datos proveedor, ítems)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb[RFQ_HOJA] if RFQ_HOJA in wb.sheetnames else wb.worksheets[0]
    ref = wb["_meta"]["A2"].value if "_meta" in wb.sheetnames else None
    datos = {}
    for r in range(1, RFQ_FILA_TABLA + 5):
        etiqueta = ws.cell(r, 1).value
        if etiqueta in RFQ_CAMPOS:
            datos[etiqueta] = ws.cell(r, 2).value
    fila_h = next((r for r in range(1, 60) if ws.cell(r, 1).value == "Ítem"), None)
    if fila_h is None:
        raise ValueError("No se encontró la tabla de ítems (encabezado 'Ítem').")
    encabezados = [ws.cell(fila_h, j).value for j in range(1, ws.max_column + 1)]
    filas = []
    for r in range(fila_h + 1, ws.max_row + 1):
        if ws.cell(r, 1).value in (None, ""):
            break
        filas.append({encabezados[j - 1]: ws.cell(r, j).value for j in range(1, len(encabezados) + 1)
                      if encabezados[j - 1]})
    return ref, datos, pd.DataFrame(filas)

# =============================================================================
# EXPORTACIÓN DEL CUADRO (Excel y PDF)
# =============================================================================
def estilo_hoja(ws, fila_header, cols_monto=(), cols_pct=(), cols_fecha=()):
    for c in ws[fila_header]:
        c.fill = PatternFill("solid", start_color=AZUL)
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    headers = {c.column: str(c.value) for c in ws[fila_header]}
    for fila in ws.iter_rows(min_row=fila_header + 1, max_row=ws.max_row):
        for c in fila:
            c.border = BORDE
            h = headers.get(c.column, "")
            if h in cols_monto:
                c.number_format = '#,##0.00'
            elif h in cols_pct:
                c.number_format = '0.0%'
            elif h in cols_fecha:
                c.number_format = 'DD/MM/YYYY'
    for col in ws.columns:
        largo = max((len(str(c.value)) for c in col if c.value is not None and c.row >= fila_header), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(largo + 2, 10), 48)
    ws.freeze_panes = ws.cell(row=fila_header + 1, column=1)


def detalle_por_solped(lineas, adj):
    """Reparte la adjudicación de cada ítem en sus líneas SOLPED/Pos (para emitir las OC)."""
    d = lineas[["SOLPED", "Pos", "Ítem", "Material", "Texto", "Cantidad", "UM", "Centro"]].merge(
        adj[["Ítem", "Proveedor adjudicado", "Moneda", "Precio mejorado", "Unit. puesto CLP"]], on="Ítem", how="left")
    d["Monto línea CLP"] = d["Unit. puesto CLP"] * d["Cantidad"]
    return d.rename(columns={"Precio mejorado": "Precio unit. (moneda oferta)"})


def generar_excel(ref, items, lineas, prov, det, adj, tc, aprobacion) -> bytes:
    buf = io.BytesIO()
    base = items[["Ítem", "SOLPEDs", "Material", "Texto", "Cantidad", "UM", "Últ. compra", "Últ. proveedor",
                  "Últ. precio", "Últ. moneda"]].copy()
    for p in prov["Proveedor"]:
        d = det[det["Proveedor"] == p].set_index("Ítem")
        base[f"{p} | P. inicial"] = base["Ítem"].map(d["Precio inicial"])
        base[f"{p} | P. mejorado"] = base["Ítem"].map(d["Precio mejorado"])
        base[f"{p} | Total CLP"] = base["Ítem"].map(d["Total CLP"]).where(base["Ítem"].map(d["Válida"]).fillna(False))
        base[f"{p} | Entrega"] = base["Ítem"].map(d["Entrega planta"])
    base = base.merge(adj[["Ítem", "Sugerido", "Proveedor adjudicado", "Tipo adjudicación", "Monto adjudicado CLP",
                           "Ahorro spot CLP", "Ahorro vs histórico CLP", "Var. vs últ. compra", "Comentario",
                           "Alertas"]], on="Ítem", how="left")
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        base.to_excel(xw, sheet_name="Cuadro comparativo", index=False, startrow=5)
        detalle_por_solped(lineas, adj).to_excel(xw, sheet_name="Detalle por SOLPED", index=False)
        det[["Ítem", "Proveedor", "Moneda", "Precio inicial", "Precio mejorado", "Cantidad", "Subtotal CLP",
             "Flete CLP", "Total CLP", "Ahorro spot CLP", "Var. vs últ. compra", "Plazo", "Entrega planta",
             "Días a planta", "Cumple técnico", "Observación"]].to_excel(xw, sheet_name="Detalle ofertas", index=False)
        prov.to_excel(xw, sheet_name="Proveedores", index=False)
        pd.DataFrame({
            "Parámetro": ["Referencia", "SOLPED incluidas", "Fecha", "USD/CLP", "EUR/CLP", "UF/CLP", "Fecha TC",
                          "Fuente TC", "Aprobado por", "Fecha aprobación", "Estado"],
            "Valor": [ref, unir_unicos(lineas["SOLPED"]), date.today().strftime("%d/%m/%Y"), tc["USD"], tc["EUR"],
                      tc["UF"], tc["fecha"], tc["fuente"], aprobacion.get("nombre", ""), aprobacion.get("fecha", ""),
                      "APROBADO" if aprobacion.get("ok") else "BORRADOR - pendiente de aprobación"],
        }).to_excel(xw, sheet_name="Parámetros", index=False)

        ws = xw.sheets["Cuadro comparativo"]
        ws["A1"] = f"ENAEX - CUADRO COMPARATIVO DE OFERTAS | {ref}"
        ws["A1"].font = Font(size=15, bold=True, color=AZUL)
        ws["A2"] = (f"Fecha: {date.today():%d/%m/%Y} | SOLPED: {unir_unicos(lineas['SOLPED'])} | USD {tc['USD']:,.2f} | "
                    f"EUR {tc['EUR']:,.2f} | UF {tc['UF']:,.2f} (TC al {tc['fecha'] or 'ingreso manual'})")
        ws["A3"] = (f"APROBADO por {aprobacion.get('nombre', '')} el {aprobacion.get('fecha', '')}"
                    if aprobacion.get("ok") else "BORRADOR - pendiente de aprobación del comprador")
        ws["A3"].font = Font(bold=True, color="065F46" if aprobacion.get("ok") else "B45309")
        ws["A4"] = (f"Monto adjudicado: $ {adj['Monto adjudicado CLP'].sum():,.0f} CLP | "
                    f"Ahorro spot: $ {adj['Ahorro spot CLP'].sum():,.0f} CLP")
        montos = [c for c in base.columns if any(k in c for k in ["precio", "P. inicial", "P. mejorado",
                                                                    "Total CLP", "Monto", "Ahorro"])]
        estilo_hoja(ws, 6, cols_monto=montos, cols_pct=["Var. vs últ. compra"],
                    cols_fecha=["Últ. compra"] + [c for c in base.columns if c.endswith("| Entrega")])
        col_idx = {c.value: c.column for c in ws[6]}
        for r, (_, fila) in enumerate(base.iterrows(), start=7):
            p = fila.get("Proveedor adjudicado")
            if p and f"{p} | Total CLP" in col_idx:
                ws.cell(row=r, column=col_idx[f"{p} | Total CLP"]).fill = PatternFill("solid", start_color=VERDE)
        estilo_hoja(xw.sheets["Detalle por SOLPED"], 1,
                    cols_monto=["Precio unit. (moneda oferta)", "Unit. puesto CLP", "Monto línea CLP"])
        estilo_hoja(xw.sheets["Detalle ofertas"], 1,
                    cols_monto=["Precio inicial", "Precio mejorado", "Subtotal CLP", "Flete CLP", "Total CLP",
                                "Ahorro spot CLP"], cols_pct=["Var. vs últ. compra"], cols_fecha=["Entrega planta"])
        estilo_hoja(xw.sheets["Proveedores"], 1, cols_monto=["Flete total"])
        estilo_hoja(xw.sheets["Parámetros"], 1)
    return buf.getvalue()


def txt_pdf(s) -> str:
    return unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("latin-1", "ignore").decode("latin-1")


def generar_pdf(ref, adj, lineas, tc, aprobacion) -> bytes:
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 8, txt_pdf(f"ENAEX - Cuadro comparativo {ref}"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(80, 80, 80)
    pdf.multi_cell(0, 5, txt_pdf(f"Fecha {date.today():%d/%m/%Y} | SOLPED: {unir_unicos(lineas['SOLPED'])} | "
                                 f"USD {tc['USD']:,.2f} | EUR {tc['EUR']:,.2f} | UF {tc['UF']:,.2f} | "
                                 f"TC al {tc['fecha'] or 'ingreso manual'}"), new_x="LMARGIN", new_y="NEXT")
    estado = (f"APROBADO por {aprobacion.get('nombre', '')} el {aprobacion.get('fecha', '')}"
              if aprobacion.get("ok") else "BORRADOR - pendiente de aprobacion del comprador")
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, txt_pdf(estado), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    cols = [("Item", 10), ("SOLPED/Pos", 36), ("Descripcion", 58), ("Cant.", 13), ("Adjudicado", 42),
            ("Tipo", 32), ("Monto CLP", 26), ("Ahorro spot", 22), ("Var.", 14), ("Entrega", 20)]
    pdf.set_fill_color(30, 58, 138)
    pdf.set_text_color(255, 255, 255)
    for n, w in cols:
        pdf.cell(w, 7, n, border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_font("Helvetica", "", 7.5)
    pdf.set_text_color(0, 0, 0)
    for i, (_, r) in enumerate(adj.iterrows()):
        pdf.set_fill_color(*((249, 250, 251) if i % 2 else (255, 255, 255)))
        var = f"{r['Var. vs últ. compra']:.0%}" if pd.notna(r["Var. vs últ. compra"]) else "-"
        ent = r["Entrega planta"].strftime("%d/%m/%Y") if hasattr(r.get("Entrega planta"), "strftime") else "-"
        vals = [str(r["Ítem"]), str(r["SOLPEDs"])[:24], str(r["Texto"])[:36], f"{r['Cantidad']:g}",
                str(r["Proveedor adjudicado"])[:26], str(r["Tipo adjudicación"])[:20],
                f"{r['Monto adjudicado CLP']:,.0f}", f"{r['Ahorro spot CLP']:,.0f}", var, ent]
        for (n, w), v in zip(cols, vals):
            pdf.cell(w, 6, txt_pdf(v), border=1, fill=True,
                     align="R" if n in ("Monto CLP", "Ahorro spot", "Cant.") else "L")
        pdf.ln()
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, txt_pdf(f"Total adjudicado: $ {adj['Monto adjudicado CLP'].sum():,.0f} CLP   |   "
                           f"Ahorro spot: $ {adj['Ahorro spot CLP'].sum():,.0f} CLP"), new_x="LMARGIN", new_y="NEXT")
    com = adj[adj["Comentario"].astype(str).str.strip() != ""]
    if not com.empty:
        pdf.set_font("Helvetica", "", 8)
        for _, r in com.iterrows():
            pdf.multi_cell(0, 5, txt_pdf(f"Item {r['Ítem']}: {r['Comentario']}"), new_x="LMARGIN", new_y="NEXT")
    return bytes(pdf.output())

# =============================================================================
# ESTADO DE SESIÓN Y PROYECTOS (un "cuadro" = una o varias SOLPED)
# =============================================================================
ss = st.session_state
ss.setdefault("proyectos", {})
ss.setdefault("req", None)
ss.setdefault("hist", None)

PROV_DEFECTO = {"Moneda": "CLP", "Transporte": "T. Gil", "Flete total": 0.0, "Moneda flete": "CLP", "Plazo": 10,
                "Tipo plazo": "Días hábiles", "Tipo compra": "Nacional",
                "Lugar entrega": list(LUGARES_ENTREGA)[0], "Bodega tránsito": False}


def nuevo_proyecto(ref: str, lineas: pd.DataFrame, agrupar: bool = True):
    lin, items = consolidar_lineas(lineas, agrupar)
    items = adjuntar_historico(items)
    sugeridos = [p for p in items["Últ. proveedor"].dropna().unique() if str(p).strip()][:3] or ["Proveedor 1"]
    prov = pd.DataFrame([{"Proveedor": p, **PROV_DEFECTO} for p in sugeridos])
    ss.proyectos[ref] = {"lineas": lin, "items": items, "prov_src": prov, "prov_v": 0, "prov_names": None,
                         "of_src": None, "of_v": 0, "of_last": None, "adj_override": {},
                         "aprobacion": {"ok": False, "nombre": "", "fecha": ""}}
    ss.sp_activa = ref


def armar_ofertas(items, nombres, previas):
    base = pd.DataFrame([(i, t, c, u, n) for i, t, c, u in zip(items["Ítem"], items["Texto"], items["Cantidad"],
                                                                 items["UM"]) for n in nombres],
                        columns=["Ítem", "Texto", "Cantidad", "UM", "Proveedor"])
    for c, v in [("Precio inicial", np.nan), ("Precio mejorado", np.nan), ("Plazo", np.nan),
                 ("Cumple técnico", True), ("Observación", "")]:
        base[c] = v
    if previas is not None and not previas.empty:
        prev = previas.drop(columns=["Texto", "Cantidad", "UM"], errors="ignore").drop_duplicates(["Ítem", "Proveedor"])
        base = base[["Ítem", "Texto", "Cantidad", "UM", "Proveedor"]].merge(prev, on=["Ítem", "Proveedor"], how="left")
        base["Cumple técnico"] = base["Cumple técnico"].astype("boolean").fillna(True).astype(bool)
        base["Observación"] = base["Observación"].fillna("").astype(str)
    return base


def incorporar_respuesta(P, datos, resp, nombre_archivo):
    """Agrega o actualiza el proveedor y carga sus precios. Una segunda respuesta
    con precio distinto se registra como 'Precio mejorado' (ronda de negociación)."""
    nombre = str(datos.get("Razón social") or "").strip() or nombre_archivo.rsplit(".", 1)[0]
    prov = (P.get("prov_actual") if P.get("prov_actual") is not None else P["prov_src"]).copy()
    moneda = str(datos.get("Moneda de la oferta") or "").upper().strip()
    transporte = str(datos.get("Incoterm / forma de despacho") or "").strip()
    nuevos = {
        "Moneda": moneda if moneda in MONEDAS else None,
        "Moneda flete": moneda if moneda in MONEDAS else None,
        "Transporte": transporte if transporte in TRANSPORTES else None,
        "Flete total": a_num(datos.get("Costo total de flete"), 0.0),
        "Plazo": int(a_num(datos.get("Plazo de entrega general (días)"), 0)) or None,
    }
    if nombre in set(prov["Proveedor"]):
        for k, v in nuevos.items():
            if v is not None:
                prov.loc[prov["Proveedor"] == nombre, k] = v
    else:
        fila = {"Proveedor": nombre, **PROV_DEFECTO, **{k: v for k, v in nuevos.items() if v is not None}}
        prov = pd.concat([prov, pd.DataFrame([fila])], ignore_index=True)
    of = armar_ofertas(P["items"], tuple(prov["Proveedor"]), P["of_last"])
    n = 0
    for _, r in resp.iterrows():
        item = int(a_num(r.get("Ítem"), 0))
        precio = parse_monto(r.get("Precio unitario (sin IVA)"))[0]
        m = (of["Ítem"] == item) & (of["Proveedor"] == nombre)
        if not m.any() or precio <= 0:
            continue
        previo = of.loc[m, "Precio inicial"].iloc[0]
        if pd.notna(previo) and previo > 0 and abs(previo - precio) > 1e-9:
            of.loc[m, "Precio mejorado"] = precio
        else:
            of.loc[m, "Precio inicial"] = precio
        plazo = a_num(r.get("Plazo entrega (días)"), 0)
        if plazo:
            of.loc[m, "Plazo"] = plazo
        cumple = str(r.get("Cumple especificación (Sí/No)") or "").strip().lower()
        of.loc[m, "Cumple técnico"] = not cumple.startswith("n")
        obs = " · ".join(str(x) for x in [r.get("Marca / modelo ofrecido"), r.get("Observaciones")]
                         if x not in (None, "") and str(x) != "nan")
        if obs:
            of.loc[m, "Observación"] = obs
        n += 1
    P["prov_src"], P["prov_v"] = prov, P["prov_v"] + 1
    P["of_last"], P["of_src"], P["prov_names"] = of, None, None
    return nombre, n


def proyecto_a_json(ref, P) -> str:
    dj = lambda df: None if df is None else df.to_json(orient="records", date_format="iso", force_ascii=False)
    return json.dumps({"ref": ref, "lineas": dj(P["lineas"]), "items": dj(P["items"]), "prov": dj(P["prov_src"]),
                       "ofertas": dj(P["of_last"]), "adj_override": P["adj_override"],
                       "aprobacion": P["aprobacion"]}, ensure_ascii=False)


def json_a_proyecto(texto: str):
    d = json.loads(texto)
    leer = lambda s: pd.read_json(io.StringIO(s), orient="records", dtype=False) if s else None
    items = leer(d["items"])
    items["Últ. compra"] = pd.to_datetime(items["Últ. compra"], errors="coerce").dt.date
    items["Material"] = items["Material"].fillna("").astype(str)
    ss.proyectos[d["ref"]] = {
        "lineas": leer(d["lineas"]), "items": items, "prov_src": leer(d["prov"]), "prov_v": 1, "prov_names": None,
        "of_src": None, "of_v": 1, "of_last": leer(d["ofertas"]),
        "adj_override": {int(k): v for k, v in d.get("adj_override", {}).items()},
        "aprobacion": d.get("aprobacion", {"ok": False, "nombre": "", "fecha": ""}),
    }
    ss.sp_activa = d["ref"]
    return d["ref"]

# =============================================================================
# BARRA LATERAL
# =============================================================================
st.markdown("<div class='titulo'>Cuadro Comparativo de Ofertas</div>"
            "<div class='subtitulo'>SOLPED y materiales → solicitud de cotización → ofertas → adjudicación "
            "con aprobación del comprador</div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("Tipos de cambio")
    ind = obtener_indicadores()
    if ind["ok"]:
        st.success(f"TC al {ind['fecha']} · {ind['fuente']}")
    else:
        st.error("Sin conexión a mindicador.cl: ingresa el TC del día manualmente.")
    if st.button("Actualizar tasas", use_container_width=True):
        obtener_indicadores.clear()
        st.rerun()
    tc = {"USD": st.number_input("USD / CLP", value=ind["USD"] or 0.0, step=1.0, format="%.2f"),
          "EUR": st.number_input("EUR / CLP", value=ind["EUR"] or 0.0, step=1.0, format="%.2f"),
          "UF": st.number_input("UF / CLP", value=ind["UF"] or 0.0, step=10.0, format="%.2f"),
          "fecha": ind["fecha"], "fuente": ind["fuente"]}
    if (tc["USD"], tc["EUR"], tc["UF"]) != (ind["USD"], ind["EUR"], ind["UF"]):
        tc["fecha"], tc["fuente"] = date.today().isoformat(), "Ingreso manual"
    tc_hoy = tc["fecha"] == date.today().isoformat()
    tc_ok = min(tc["USD"], tc["EUR"], tc["UF"]) > 0
    if not tc_ok:
        st.warning("Completa los tres tipos de cambio para convertir monedas.")

    st.divider()
    st.header("Archivos")
    f_req = st.file_uploader("1. Carga de SOLPED (Autogestión, ME5A o Ariba)", type=["xlsx", "xlsm", "xls", "csv"])
    if f_req:
        hojas = listar_hojas(f_req.getvalue(), f_req.name)
        hoja_req = st.selectbox("Hoja de requerimientos", hojas, index=int(np.argmax(
            [puntaje_hoja(h, ["Carga ARIBA", "Carga", "ME5A", "Planilla anterior"]) for h in hojas])))
    f_hist = st.file_uploader("2. Histórico de compras (Cuadro MRO, ME2M o CSV)", type=["xlsx", "xlsm", "xls", "csv"])
    if f_hist:
        hojas_h = listar_hojas(f_hist.getvalue(), f_hist.name)
        hoja_hist = st.selectbox("Hoja de histórico", hojas_h, index=int(np.argmax(
            [puntaje_hoja(h, ["Ultima COMPRA", "BASE OC", "ME2M", "Data"]) for h in hojas_h])))
    if st.button("Procesar archivos", type="primary", use_container_width=True):
        if f_req:
            df = leer_hoja(f_req.getvalue(), f_req.name, hoja_req, json.dumps(SIN_REQ))
            ss.req_raw, ss.req_map = df, mapear_columnas(df.columns, SIN_REQ)
        if f_hist:
            dfh = leer_hoja(f_hist.getvalue(), f_hist.name, hoja_hist, json.dumps(SIN_HIST))
            ss.hist_raw, ss.hist_map = dfh, mapear_columnas(dfh.columns, SIN_HIST)

    for etiqueta, raw_k, map_k, sins, oblig in [("requerimientos", "req_raw", "req_map", SIN_REQ, REQ_OBLIGATORIAS),
                                                ("histórico", "hist_raw", "hist_map", SIN_HIST, HIST_OBLIGATORIAS)]:
        if ss.get(raw_k) is not None:
            faltan = [c for c in oblig if c not in ss[map_k]]
            with st.expander(f"Columnas de {etiqueta}" + (" ⚠" if faltan else " ✓"), expanded=bool(faltan)):
                opciones = ["(no usar)"] + list(ss[raw_k].columns)
                for campo in sins:
                    actual = ss[map_k].get(campo)
                    elegido = st.selectbox(campo, opciones, key=f"map_{map_k}_{campo}",
                                           index=opciones.index(actual) if actual in opciones else 0)
                    if elegido == "(no usar)":
                        ss[map_k].pop(campo, None)
                    else:
                        ss[map_k][campo] = elegido

    if ss.get("hist_raw") is not None and all(c in ss.hist_map for c in HIST_OBLIGATORIAS):
        ss.hist = construir_historico(ss.hist_raw, ss.hist_map)
    if ss.get("req_raw") is not None and all(c in ss.req_map for c in REQ_OBLIGATORIAS):
        ss.req = normalizar_requerimientos(ss.req_raw, ss.req_map)
    if ss.req is not None:
        st.caption(f"{ss.req['SOLPED'].nunique()} SOLPED · {len(ss.req)} posiciones · "
                   f"{ss.req['Clave'].nunique()} materiales distintos")
    if ss.hist is not None:
        st.caption(f"Histórico: {len(ss.hist[0]):,} materiales con última compra")

    st.divider()
    st.header("Borradores")
    f_json = st.file_uploader("Retomar borrador (.json)", type=["json"])
    if f_json and st.button("Cargar borrador", use_container_width=True):
        st.success(f"Borrador {json_a_proyecto(f_json.getvalue().decode('utf-8'))} cargado.")

# =============================================================================
# PANEL DE ACCIONES SOBRE LÍNEAS SELECCIONADAS (agrupar, cotizar, crear cuadro)
# =============================================================================
def panel_lineas(lineas_sel: pd.DataFrame, ref_defecto: str, clave: str):
    firma = str(abs(hash(tuple(lineas_sel["SOLPED"] + "/" + lineas_sel["Pos"].astype(str)))))
    st.markdown(f"##### {len(lineas_sel)} línea(s) de {lineas_sel['SOLPED'].nunique()} SOLPED")
    editable = lineas_sel[["SOLPED", "Pos", "Material", "Texto", "Cantidad", "UM", "Centro", "Días",
                           "Asignado a"]].copy()
    editable.insert(0, "Incluir", True)
    incl = st.data_editor(editable, key=f"incl_{clave}_{firma}", hide_index=True, use_container_width=True,
                          disabled=[c for c in editable.columns if c != "Incluir"],
                          height=min(70 + 35 * len(editable), 420),
                          column_config={"Días": st.column_config.NumberColumn(format="%d")})
    elegidas = lineas_sel[incl["Incluir"].to_numpy()]
    if elegidas.empty:
        st.info("Marca al menos una línea.")
        return
    c1, c2 = st.columns([1, 2])
    agrupar = c1.checkbox("Consolidar mismo material en un ítem", value=True, key=f"agr_{clave}",
                          help="Suma las cantidades de todas las SOLPED: el proveedor cotiza una vez por material.")
    ref = c2.text_input("Nombre del cuadro", value=ref_defecto, key=f"ref_{clave}_{firma}")
    lin, items = consolidar_lineas(elegidas, agrupar)
    st.caption(f"Se cotizarán {len(items)} ítem(s).")
    st.dataframe(items[["Ítem", "Material", "Texto", "Cantidad", "UM", "Centro", "SOLPEDs"]],
                 hide_index=True, use_container_width=True)

    with st.expander("Datos de la solicitud de cotización", expanded=False):
        d1, d2, d3 = st.columns(3)
        comprador = d1.text_input("Comprador", value=unir_unicos(elegidas["Asignado a"]) or "",
                                  key=f"comp_{clave}")
        correo = d2.text_input("Correo de respuesta", key=f"mail_{clave}")
        limite = d3.date_input("Responder antes de", value=date.today() + timedelta(days=5), key=f"lim_{clave}")
        entrega = st.text_input("Lugar / condición de entrega", value=unir_unicos(items["Centro"]),
                                key=f"ent_{clave}")
        detalle = st.checkbox("Incluir hoja con detalle por SOLPED", value=True, key=f"det_{clave}")
    cfg = {"comprador": comprador, "correo": correo, "fecha_limite": limite, "entrega": entrega}

    b1, b2 = st.columns(2)
    b1.download_button("Descargar solicitud de cotización (Excel para proveedor)",
                       data=generar_rfq(ref, items, lin, cfg, detalle),
                       file_name=f"Solicitud_cotizacion_{re.sub(r'[^A-Za-z0-9_-]', '_', ref)}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True, key=f"rfq_{clave}")
    if b2.button("Crear cuadro comparativo con estas líneas", type="primary", use_container_width=True,
                 key=f"crear_{clave}"):
        if ref in ss.proyectos:
            st.warning(f"Ya existe un cuadro llamado {ref}; cambia el nombre.")
        else:
            nuevo_proyecto(ref, elegidas, agrupar)
            st.success(f"Cuadro {ref} creado. Continúa en la pestaña 2.")

# =============================================================================
# PESTAÑAS
# =============================================================================
tab_bandeja, tab_ofertas, tab_adj = st.tabs(["1. Bandeja: SOLPED y materiales", "2. Proveedores y ofertas",
                                              "3. Adjudicación y reportes"])

# ---------------------------------------------------------------- TAB 1
with tab_bandeja:
    if ss.req is not None and not ss.req.empty:
        req = ss.req
        f1, f2, f3 = st.columns([2, 3, 2])
        compradores = sorted(c for c in req["Asignado a"].unique() if c)
        filtro_comp = f1.multiselect("Comprador", compradores) if compradores else []
        busca = f2.text_input("Buscar SOLPED, código o descripción")
        vista_por = f3.radio("Ver por", ["Material", "SOLPED", "Todas las líneas"], horizontal=True)
        base = req if not filtro_comp else req[req["Asignado a"].isin(filtro_comp)]
        if busca:
            t = norm(busca)
            base = base[(base["SOLPED"] + " " + base["Material"] + " " + base["Texto"]).apply(norm).str.contains(t, regex=False)]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("SOLPED", base["SOLPED"].nunique())
        k2.metric("Posiciones", len(base))
        k3.metric("Materiales distintos", base["Clave"].nunique())
        repetidos = base.groupby(["Clave", "UM"])["SOLPED"].nunique()
        k4.metric("Materiales en 2+ SOLPED", int((repetidos > 1).sum()))

        if vista_por == "Material":
            solo_rep = st.checkbox("Mostrar solo materiales pedidos en más de una SOLPED", value=False)
            mats = (base.groupby(["Clave", "UM"], sort=False)
                    .agg(Material=("Material", "first"), Texto=("Texto", lambda s: max(s, key=len)),
                         **{"N° SOLPED": ("SOLPED", "nunique"), "Posiciones": ("Pos", "size"),
                            "Cantidad total": ("Cantidad", "sum"), "SOLPEDs": ("SOLPED", unir_unicos),
                            "Centros": ("Centro", unir_unicos), "Días máx.": ("Días", "max")})
                    .reset_index())
            if solo_rep:
                mats = mats[mats["N° SOLPED"] > 1]
            mats = adjuntar_historico(mats).sort_values(["N° SOLPED", "Días máx."], ascending=False).reset_index(drop=True)
            st.caption("Selecciona uno o más materiales para ver todas las SOLPED que los piden.")
            sel = st.dataframe(
                mats[["Material", "Texto", "UM", "N° SOLPED", "Cantidad total", "SOLPEDs", "Centros", "Días máx.",
                      "Últ. compra", "Últ. proveedor", "Últ. precio", "Últ. moneda"]],
                hide_index=True, use_container_width=True, height=360, on_select="rerun",
                selection_mode="multi-row", key="sel_mat",
                column_config={"Últ. precio": st.column_config.NumberColumn(format="%.2f"),
                               "Días máx.": st.column_config.NumberColumn(format="%d"),
                               "Últ. compra": st.column_config.DateColumn(format="DD/MM/YYYY")})
            if sel.selection.rows:
                elegidos = mats.iloc[sel.selection.rows][["Clave", "UM"]]
                lineas_sel = base.merge(elegidos, on=["Clave", "UM"]).sort_values(["Clave", "SOLPED", "Pos"])
                ref_def = (f"MAT-{elegidos.iloc[0]['Clave'][:20]}" if len(elegidos) == 1 else "AGRUPADO") \
                    + f"-{date.today():%Y%m%d}"
                st.divider()
                panel_lineas(lineas_sel.reset_index(drop=True), ref_def.replace("TXT:", ""), "mat")

        elif vista_por == "SOLPED":
            bandeja = (base.groupby("SOLPED")
                       .agg(Posiciones=("Pos", "size"), Días=("Días", "max"), Centro=("Centro", "first"),
                            Comprador=("Asignado a", "first"), Urgencia=("Urgencia", "first"),
                            Materiales=("Texto", lambda s: " | ".join(list(s)[:3]) + (" …" if len(s) > 3 else "")))
                       .reset_index().sort_values("Días", ascending=False).reset_index(drop=True))
            bandeja["Tramo"] = pd.cut(bandeja["Días"], [-1, 5, 15, 30, 10_000],
                                      labels=["0-5 días", "6-15 días", "16-30 días", "30+ días"]).astype(str)
            st.caption("Ordenadas por antigüedad. Puedes seleccionar varias SOLPED para cotizarlas juntas.")
            sel = st.dataframe(bandeja, hide_index=True, use_container_width=True, height=360,
                               on_select="rerun", selection_mode="multi-row", key="sel_sp",
                               column_config={"Días": st.column_config.NumberColumn(format="%d"),
                                              "Materiales": st.column_config.TextColumn(width="large")})
            if sel.selection.rows:
                sps = bandeja.iloc[sel.selection.rows]["SOLPED"].tolist()
                lineas_sel = base[base["SOLPED"].isin(sps)].sort_values(["SOLPED", "Pos"]).reset_index(drop=True)
                st.divider()
                panel_lineas(lineas_sel, sps[0] if len(sps) == 1 else f"AGRUPADO-{date.today():%Y%m%d}", "sp")

        else:
            st.caption("Todas las posiciones cargadas. Selecciona las que quieras cotizar.")
            vista = base.sort_values(["Texto", "SOLPED"]).reset_index(drop=True)
            sel = st.dataframe(vista[["SOLPED", "Pos", "Material", "Texto", "Cantidad", "UM", "Centro", "Días",
                                      "Urgencia", "Asignado a"]],
                               hide_index=True, use_container_width=True, height=420, on_select="rerun",
                               selection_mode="multi-row", key="sel_lin",
                               column_config={"Días": st.column_config.NumberColumn(format="%d")})
            if sel.selection.rows:
                st.divider()
                panel_lineas(vista.iloc[sel.selection.rows].reset_index(drop=True),
                             f"SEL-{date.today():%Y%m%d}", "lin")
    else:
        st.info("Carga la planilla de SOLPED en la barra lateral y presiona 'Procesar archivos'. "
                "También puedes crear un cuadro manual aquí abajo.")

    with st.expander("Crear cuadro manual (sin planilla)", expanded=ss.req is None):
        sp_man = st.text_input("N° SOLPED", key="sp_manual")
        plantilla = pd.DataFrame({"Pos": [10], "Material": [""], "Texto": [""], "Cantidad": [1.0],
                                  "UM": ["C/U"], "Centro": [""]})
        pos_man = st.data_editor(plantilla, num_rows="dynamic", use_container_width=True, key="pos_manual")
        if st.button("Crear cuadro manual") and sp_man.strip():
            pm = pos_man.dropna(subset=["Texto"])
            pm = pm[pm["Texto"].astype(str).str.strip() != ""].copy()
            pm["Material"] = pm["Material"].apply(clave_material)
            pm["SOLPED"] = sp_man.strip()
            pm["Asignado a"], pm["Días"] = "", np.nan
            nuevo_proyecto(sp_man.strip(), pm, agrupar=False)
            st.success("Cuadro creado. Continúa en la pestaña 2.")

    if ss.proyectos:
        st.caption("Cuadros abiertos en esta sesión: " + ", ".join(ss.proyectos))

# ---------------------------------------------------------------- CUADRO ACTIVO
ref = ss.get("sp_activa")
P = ss.proyectos.get(ref) if ref else None

# ---------------------------------------------------------------- TAB 2
with tab_ofertas:
    if P is None:
        st.info("Crea un cuadro desde la pestaña 1.")
    else:
        if len(ss.proyectos) > 1:
            ref = st.selectbox("Cuadro activo", list(ss.proyectos), index=list(ss.proyectos).index(ref))
            ss.sp_activa, P = ref, ss.proyectos[ref]
        items, lineas = P["items"], P["lineas"]
        items["Últ. precio CLP"] = [a_clp(p, m, tc) if pd.notna(p) else np.nan
                                    for p, m in zip(items["Últ. precio"], items["Últ. moneda"])]
        items["Últ. precio CLP"] = items["Últ. precio CLP"].where(items["Últ. precio CLP"] > 0)

        st.subheader(f"{ref} · {len(items)} ítem(s) de {lineas['SOLPED'].nunique()} SOLPED")
        st.dataframe(items[["Ítem", "Material", "Texto", "Cantidad", "UM", "SOLPEDs", "Últ. compra",
                            "Últ. proveedor", "Últ. precio", "Últ. moneda", "N° compras", "Historial por"]],
                     hide_index=True, use_container_width=True,
                     column_config={"Últ. precio": st.column_config.NumberColumn(format="%.2f"),
                                    "Últ. compra": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                    "SOLPEDs": st.column_config.TextColumn(width="medium")})
        sin_hist = int((~items["Con historial"].astype(bool)).sum())
        if sin_hist:
            st.warning(f"{sin_hist} ítem(s) sin historial de compra: requieren validación técnica del especialista.")

        with st.expander("Solicitud de cotización e importación de respuestas", expanded=False):
            r1, r2 = st.columns(2)
            with r1:
                st.markdown("**Enviar a proveedores**")
                cfg = {"comprador": st.text_input("Comprador", key=f"rc_{ref}",
                                                  value=unir_unicos(lineas.get("Asignado a", pd.Series(dtype=str)))),
                       "correo": st.text_input("Correo de respuesta", key=f"rm_{ref}"),
                       "fecha_limite": st.date_input("Responder antes de", key=f"rl_{ref}",
                                                     value=date.today() + timedelta(days=5)),
                       "entrega": st.text_input("Lugar / condición de entrega", key=f"re_{ref}",
                                                value=unir_unicos(items["Centro"]))}
                st.download_button("Descargar solicitud de cotización", data=generar_rfq(ref, items, lineas, cfg),
                                   file_name=f"Solicitud_cotizacion_{re.sub(r'[^A-Za-z0-9_-]', '_', ref)}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            with r2:
                st.markdown("**Importar respuestas de proveedores**")
                st.caption("Sube la misma planilla devuelta por cada proveedor. Si ya tenía precio, "
                           "el nuevo se registra como precio mejorado.")
                respuestas = st.file_uploader("Planillas respondidas (.xlsx)", type=["xlsx"],
                                              accept_multiple_files=True, key=f"resp_{ref}")
                if respuestas and st.button("Incorporar respuestas", type="primary", use_container_width=True):
                    for f in respuestas:
                        try:
                            ref_resp, datos, resp = leer_respuesta_rfq(f.getvalue())
                            if ref_resp and ref_resp != ref:
                                st.error(f"{f.name}: corresponde al cuadro '{ref_resp}', no a '{ref}'. Se omitió.")
                                continue
                            nombre, n = incorporar_respuesta(P, datos, resp, f.name)
                            st.success(f"{nombre}: {n} precio(s) cargados desde {f.name}.")
                        except Exception as e:
                            st.error(f"{f.name}: no se pudo leer ({e}).")
                    st.rerun()

        st.markdown("#### Proveedores")
        st.caption("Una fila por proveedor que cotizó. El flete se prorratea entre sus ítems según valor.")
        prov_edit = st.data_editor(
            P["prov_src"], key=f"prov_{ref}_{P['prov_v']}", num_rows="dynamic", use_container_width=True,
            hide_index=True,
            column_config={
                "Proveedor": st.column_config.TextColumn(required=True),
                "Moneda": st.column_config.SelectboxColumn(options=MONEDAS, required=True, default="CLP"),
                "Transporte": st.column_config.SelectboxColumn(options=TRANSPORTES, default="T. Gil"),
                "Flete total": st.column_config.NumberColumn(min_value=0.0, format="%.2f", default=0.0),
                "Moneda flete": st.column_config.SelectboxColumn(options=MONEDAS, default="CLP"),
                "Plazo": st.column_config.NumberColumn("Plazo por defecto", min_value=0, step=1, default=10),
                "Tipo plazo": st.column_config.SelectboxColumn(options=TIPOS_PLAZO, default="Días hábiles"),
                "Tipo compra": st.column_config.SelectboxColumn(options=TIPOS_COMPRA, default="Nacional"),
                "Lugar entrega": st.column_config.SelectboxColumn(options=list(LUGARES_ENTREGA),
                                                                  default=list(LUGARES_ENTREGA)[0]),
                "Bodega tránsito": st.column_config.CheckboxColumn(default=False),
            })
        prov = prov_edit.dropna(subset=["Proveedor"]).copy()
        prov["Proveedor"] = prov["Proveedor"].astype(str).str.strip()
        prov = prov[prov["Proveedor"] != ""].drop_duplicates("Proveedor")
        for c, v in PROV_DEFECTO.items():
            prov[c] = prov[c].fillna(v) if c in prov else v
        P["prov_actual"] = prov
        nombres = tuple(prov["Proveedor"])
        if nombres != P["prov_names"] or P["of_src"] is None:
            P["of_src"] = armar_ofertas(items, nombres, P["of_last"])
            P["of_v"] += 1
            P["prov_names"] = nombres

        st.markdown("#### Ofertas por ítem")
        st.caption("Precio mejorado vacío = se usa el inicial. Desmarca 'Cumple técnico' si no cumple la ficha. "
                   "Plazo vacío = plazo por defecto del proveedor.")
        of_edit = st.data_editor(
            P["of_src"], key=f"of_{ref}_{P['of_v']}", use_container_width=True, hide_index=True,
            height=min(80 + 35 * len(P["of_src"]), 600), disabled=["Ítem", "Texto", "Cantidad", "UM", "Proveedor"],
            column_config={"Texto": st.column_config.TextColumn(width="large"),
                           "Precio inicial": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                           "Precio mejorado": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                           "Plazo": st.column_config.NumberColumn(min_value=0, step=1),
                           "Cumple técnico": st.column_config.CheckboxColumn()})
        P["of_last"] = of_edit

        if not tc_ok:
            st.error("Faltan tipos de cambio: no se puede comparar entre monedas.")
            P["res"] = None
        elif of_edit["Precio inicial"].fillna(0).gt(0).any():
            det, res = calcular_comparativo(items, prov, of_edit, tc)
            P["det"], P["res"] = det, res
            st.markdown("#### Comparación (total puesto en planta, CLP)")
            matriz = det.pivot_table(index="Ítem", columns="Proveedor", values="Total CLP", aggfunc="first")
            validas = det.pivot_table(index="Ítem", columns="Proveedor", values="Válida", aggfunc="first")
            matriz = matriz.where(validas.fillna(False).astype(bool))
            matriz.insert(0, "Descripción", matriz.index.map(items.set_index("Ítem")["Texto"]))

            def verde_min(fila):
                valores = pd.to_numeric(fila.drop("Descripción"), errors="coerce")
                if valores.notna().sum() == 0:
                    return [""] * len(fila)
                return [""] + ["background-color:#D1FAE5;font-weight:600" if v == valores.min() else ""
                               for v in valores]

            num_cols = [c for c in matriz.columns if c != "Descripción"]
            st.dataframe(matriz.style.apply(verde_min, axis=1).format("$ {:,.0f}", subset=num_cols, na_rep="—"),
                         use_container_width=True)
            st.caption("Verde = menor total válido (incluye flete prorrateado). '—' = no cotiza o no cumple técnico.")
        else:
            P["res"] = None
            st.info("Ingresa precios (o importa las respuestas de proveedores) para ver la comparación.")

# ---------------------------------------------------------------- TAB 3
with tab_adj:
    if P is None or P.get("res") is None:
        st.info("Completa proveedores y ofertas en la pestaña 2.")
    else:
        items, lineas, det, res, prov = P["items"], P["lineas"], P["det"], P["res"], P["prov_actual"]
        st.subheader(f"Adjudicación {ref}")
        st.caption("El sistema propone el menor total válido. El comprador decide; si se aparta de la sugerencia, "
                   "el comentario es obligatorio.")
        src = res.merge(items[["Ítem", "Texto", "Cantidad", "UM", "SOLPEDs"]], on="Ítem")
        src["Proveedor adjudicado"] = src["Sugerido"]
        src["Tipo adjudicación"] = src["Tipo sugerido"].replace("", TIPOS_ADJUDICACION[0])
        src["Comentario"] = ""
        for it_n, ov in P["adj_override"].items():
            m = src["Ítem"] == it_n
            if m.any():
                for k, v in ov.items():
                    src.loc[m, k] = v
        src = src[["Ítem", "SOLPEDs", "Texto", "Cantidad", "UM", "Sugerido", "Mejor total CLP", "Ofertas válidas",
                   "Alertas", "Proveedor adjudicado", "Tipo adjudicación", "Comentario"]]
        firma = str(pd.util.hash_pandas_object(src.astype(str), index=False).sum())
        adj_edit = st.data_editor(
            src, key=f"adj_{ref}_{firma}", hide_index=True, use_container_width=True,
            disabled=["Ítem", "SOLPEDs", "Texto", "Cantidad", "UM", "Sugerido", "Mejor total CLP", "Ofertas válidas",
                      "Alertas"],
            column_config={"Texto": st.column_config.TextColumn(width="medium"),
                           "Mejor total CLP": st.column_config.NumberColumn(format="$ %.0f"),
                           "Proveedor adjudicado": st.column_config.SelectboxColumn(
                               options=[SIN_ADJ] + list(prov["Proveedor"])),
                           "Tipo adjudicación": st.column_config.SelectboxColumn(options=TIPOS_ADJUDICACION),
                           "Comentario": st.column_config.TextColumn(width="large")})
        P["adj_override"] = {int(r["Ítem"]): {"Proveedor adjudicado": r["Proveedor adjudicado"],
                                              "Tipo adjudicación": r["Tipo adjudicación"],
                                              "Comentario": r["Comentario"] or ""} for _, r in adj_edit.iterrows()}

        adj = adj_edit.merge(
            det.rename(columns={"Proveedor": "Proveedor adjudicado"})[
                ["Ítem", "Proveedor adjudicado", "Moneda", "Precio mejorado", "Total CLP", "Ahorro spot CLP",
                 "Var. vs últ. compra", "Entrega planta", "Válida", "Precio unit. CLP", "Unit. puesto CLP"]],
            on=["Ítem", "Proveedor adjudicado"], how="left")
        adj = adj.merge(items[["Ítem", "Últ. precio CLP"]], on="Ítem", how="left")
        adj["Monto adjudicado CLP"] = adj["Total CLP"].fillna(0)
        adj["Ahorro spot CLP"] = adj["Ahorro spot CLP"].fillna(0)
        adj["Ahorro vs histórico CLP"] = ((adj["Últ. precio CLP"] - adj["Precio unit. CLP"]) * adj["Cantidad"]).fillna(0)
        adj["Válida"] = adj["Válida"].astype("boolean").fillna(False).astype(bool)

        problemas = []
        for _, r in adj.iterrows():
            if r["Proveedor adjudicado"] == SIN_ADJ:
                problemas.append(f"Ítem {r['Ítem']}: sin proveedor adjudicado.")
            elif not r["Válida"]:
                problemas.append(f"Ítem {r['Ítem']}: {r['Proveedor adjudicado']} no tiene oferta válida (precio o técnico).")
            elif r["Proveedor adjudicado"] != r["Sugerido"] and not str(r["Comentario"]).strip():
                problemas.append(f"Ítem {r['Ítem']}: se adjudica distinto a la sugerencia; falta comentario.")

        total = adj["Monto adjudicado CLP"].sum()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Monto adjudicado", f"$ {total:,.0f}")
        k2.metric("Ahorro spot (negociación)", f"$ {adj['Ahorro spot CLP'].sum():,.0f}")
        k3.metric("Ahorro vs última compra", f"$ {adj['Ahorro vs histórico CLP'].sum():,.0f}")
        k4.metric("Ítems adjudicados", f"{int((adj['Proveedor adjudicado'] != SIN_ADJ).sum())} / {len(adj)}")
        if total > UMBRAL_TC_CLP and not tc_hoy:
            st.error(f"Monto sobre $ {UMBRAL_TC_CLP:,.0f} CLP: actualiza el tipo de cambio del día "
                     f"(el actual es del {tc['fecha'] or 'sin fecha'}).")
        for p in problemas:
            st.warning(p)

        st.markdown("#### Reparto por SOLPED (para emitir las OC)")
        reparto = detalle_por_solped(lineas, adj)
        st.dataframe(reparto, hide_index=True, use_container_width=True,
                     column_config={"Precio unit. (moneda oferta)": st.column_config.NumberColumn(format="%.2f"),
                                    "Unit. puesto CLP": st.column_config.NumberColumn(format="$ %.0f"),
                                    "Monto línea CLP": st.column_config.NumberColumn(format="$ %.0f")})

        st.divider()
        st.markdown("#### Aprobación del comprador")
        ap = P["aprobacion"]
        c1, c2 = st.columns([2, 1])
        nombre = c1.text_input("Nombre de quien aprueba", value=ap.get("nombre", ""))
        puede = not problemas and (total <= UMBRAL_TC_CLP or tc_hoy)
        aprobado = c2.checkbox("Revisé y apruebo esta adjudicación", value=ap.get("ok", False) and puede,
                               disabled=not puede or not nombre.strip())
        if not puede:
            st.caption("La aprobación se habilita cuando no quedan observaciones pendientes.")
        P["aprobacion"] = {"ok": bool(aprobado), "nombre": nombre.strip(),
                           "fecha": datetime.now().strftime("%d/%m/%Y %H:%M") if aprobado else ""}

        st.markdown("#### Reportes")
        sufijo = "" if P["aprobacion"]["ok"] else "_BORRADOR"
        nombre_arch = re.sub(r"[^A-Za-z0-9_-]", "_", ref)
        d1, d2, d3 = st.columns(3)
        d1.download_button("Excel cuadro comparativo",
                           data=generar_excel(ref, items, lineas, prov, det, adj, tc, P["aprobacion"]),
                           file_name=f"Cuadro_comparativo_{nombre_arch}_{date.today():%Y%m%d}{sufijo}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)
        d2.download_button("PDF resumen", data=generar_pdf(ref, adj, lineas, tc, P["aprobacion"]),
                           file_name=f"Cuadro_comparativo_{nombre_arch}_{date.today():%Y%m%d}{sufijo}.pdf",
                           mime="application/pdf", use_container_width=True)
        d3.download_button("Guardar borrador (.json)", data=proyecto_a_json(ref, P),
                           file_name=f"Borrador_{nombre_arch}.json", mime="application/json",
                           use_container_width=True)
