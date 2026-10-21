# =============================================================================
# CUADRO COMPARATIVO DE OFERTAS - ABASTECIMIENTO ENAEX  (v4 · asistente guiado)
# -----------------------------------------------------------------------------
# Un asistente de 5 pasos, un paso a la vez:
#   1. Sube tus archivos        (planilla de SOLPED + histórico opcional)
#   2. Elige qué cotizar        (por material o por SOLPED; se agrupan solas)
#   3. Pide la cotización       (descargas la planilla, subes las respuestas)
#   4. Compara y elige          (tarjeta por ítem con el más conveniente)
#   5. Aprueba y descarga       (Excel, PDF y guardar para seguir después)
#
# requirements.txt:
#   streamlit>=1.37  pandas>=2.1  numpy  openpyxl  requests  fpdf2
#   python-calamine (opcional, lee Excel ~10x más rápido)
#   holidays        (opcional, feriados de Chile en fechas de entrega)
# =============================================================================
import hashlib
import html
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
LUGARES_ENTREGA = {
    "EDP - Entrega directa en planta": 0,
    "EBP - Entrega en bodega de proveedor": 0,
    "EBC - Entrega bodega consolidación": 7,
    "EXW A - Origen aéreo": 15,
    "EXW M - Origen marítimo": 45,
}
DIAS_EMISION_OC = 2
DIAS_INTERNACION = 10
DIAS_BODEGA_TRANSITO = 7
UMBRAL_TC_CLP = 1_000_000
UMBRAL_ALZA = 0.15
SIN_ADJ = "— Sin adjudicar —"

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

PASOS = [("📂", "Sube tus archivos"), ("🔎", "Elige qué cotizar"), ("✉️", "Pide la cotización"),
         ("⚖️", "Compara y elige"), ("✅", "Aprueba y descarga")]

# =============================================================================
# PÁGINA Y ESTILO VISUAL
# =============================================================================
st.set_page_config(page_title="Cuadro Comparativo - Enaex", page_icon="⚖️", layout="wide",
                   initial_sidebar_state="collapsed")
st.markdown("""
<style>
  [data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none; }
  [data-testid="stMainBlockContainer"] { max-width: 1180px; padding-top: 1.5rem; padding-bottom: 5rem; }
  h1, h2, h3 { color: #1E3A8A; }
  .marca { font-size: .85rem; color: #6B7280; letter-spacing: .04em; text-transform: uppercase; }
  .pasos { display: flex; gap: .5rem; margin: .4rem 0 1.4rem 0; }
  .paso { flex: 1; text-align: center; padding: .65rem .3rem; border-radius: 12px; font-size: .9rem;
          background: #F3F4F6; color: #6B7280; border: 2px solid transparent; }
  .paso b { display: block; font-size: 1.25rem; margin-bottom: .15rem; }
  .paso.hecho { background: #ECFDF5; color: #065F46; }
  .paso.actual { background: #EEF2FF; color: #1E3A8A; border-color: #1E3A8A; font-weight: 600; }
  .ayuda { background: #F8FAFC; border-left: 4px solid #1E3A8A; padding: .8rem 1rem; border-radius: 8px;
           color: #374151; margin-bottom: 1rem; }
  .ok { background: #ECFDF5; border-left: 4px solid #059669; padding: .7rem 1rem; border-radius: 8px;
        color: #065F46; margin: .4rem 0; }
  .ojo { background: #FFFBEB; border-left: 4px solid #D97706; padding: .7rem 1rem; border-radius: 8px;
         color: #92400E; margin: .4rem 0; }
  .mini { display: flex; gap: .8rem; margin: .5rem 0 1rem 0; }
  .mini div { flex: 1; background: #F8FAFC; border-radius: 12px; padding: .9rem; text-align: center; color: #374151; }
  .mini div b { display: block; font-size: 1.6rem; color: #1E3A8A; }
  .fila-oferta { display: grid; grid-template-columns: 210px 1fr 150px 120px; gap: .6rem; align-items: center;
                 padding: .35rem 0; font-size: .92rem; }
  .barra { height: 14px; border-radius: 7px; background: #CBD5E1; }
  .barra.gana { background: #059669; }
  .etiqueta { font-size: .75rem; padding: .1rem .5rem; border-radius: 999px; background: #D1FAE5; color: #065F46;
              margin-left: .4rem; white-space: nowrap; }
  .etiqueta.gris { background: #F3F4F6; color: #6B7280; }
  .etiqueta.roja { background: #FEE2E2; color: #991B1B; }
  .sub { color: #6B7280; font-size: .88rem; }
  div.stButton > button, div.stDownloadButton > button { border-radius: 10px; padding: .55rem 1.1rem; }
  button[kind="primary"], [data-testid="stBaseButton-primary"] {
      background: #1E3A8A !important; border-color: #1E3A8A !important; color: #FFFFFF !important; font-weight: 600; }
  button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover { background: #2548A8 !important; }
  button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled {
      background: #E5E7EB !important; border-color: #E5E7EB !important; color: #9CA3AF !important; }
  [data-testid="stBaseButton-segmented_controlActive"], [data-testid="stBaseButton-pillsActive"] {
      background: #EEF2FF !important; border-color: #1E3A8A !important; color: #1E3A8A !important; }
  [data-testid="stCheckbox"] input:checked + div, [role="radio"][aria-checked="true"] { accent-color: #1E3A8A; }
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
# ESTADO DE LA SESIÓN
# =============================================================================
ss = st.session_state
ss.setdefault("paso", 1)
ss.setdefault("req", None)          # posiciones de SOLPED normalizadas
ss.setdefault("hist", None)         # (última compra por código, por descripción)
ss.setdefault("cuadro", None)       # el cuadro comparativo en curso
ss.setdefault("mat_sel", set())     # materiales elegidos  (clave|UM)
ss.setdefault("sp_sel", set())      # SOLPED elegidas
ss.setdefault("sel_sig", None)
ss.setdefault("sel_ver", 0)
ss.setdefault("archivos", {})       # firma de archivos ya leídos

PROV_DEFECTO = {"Moneda": "CLP", "Flete total": 0.0, "Plazo": 10, "Tipo compra": "Nacional",
                "Lugar entrega": list(LUGARES_ENTREGA)[0], "Transporte": "T. Gil", "Tipo plazo": "Días hábiles",
                "Bodega tránsito": False, "Moneda flete": "CLP"}
COLS_PROV_VISIBLES = ["Proveedor", "Moneda", "Flete total", "Plazo", "Tipo compra", "Lugar entrega"]

if "tc" not in ss:
    ind = obtener_indicadores()
    ss.tc = {"USD": ind["USD"], "EUR": ind["EUR"], "UF": ind["UF"], "fecha": ind["fecha"], "fuente": ind["fuente"]}
tc = ss.tc


def firma(*partes) -> str:
    return hashlib.md5("|".join(map(str, partes)).encode()).hexdigest()[:10]


def dinero(v, moneda="CLP") -> str:
    if v is None or pd.isna(v):
        return "—"
    if moneda == "CLP":
        return "$ " + f"{v:,.0f}".replace(",", ".")
    return f"{moneda} " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nombre_archivo(txt) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", str(txt))


def caja(clase, texto):
    st.markdown(f"<div class='{clase}'>{texto}</div>", unsafe_allow_html=True)


def ir_a(paso: int):
    """Cambia de paso y 'congela' las grillas editables para no perder lo escrito."""
    C = ss.cuadro
    if C is not None:
        if C.get("of_last") is not None:
            C["of_src"], C["of_v"] = C["of_last"].copy(), C["of_v"] + 1
        if C.get("prov_actual") is not None:
            C["prov_src"], C["prov_v"] = C["prov_actual"].copy(), C["prov_v"] + 1
    ss.sel_sig = None
    ss.paso = paso


def lineas_elegidas() -> pd.DataFrame:
    req = ss.req
    if req is None or (not ss.mat_sel and not ss.sp_sel):
        return pd.DataFrame()
    clave = req["Clave"] + "|" + req["UM"]
    return req[clave.isin(ss.mat_sel) | req["SOLPED"].isin(ss.sp_sel)].sort_values(["Clave", "SOLPED", "Pos"])


def crear_cuadro(lineas: pd.DataFrame, agrupar: bool = True):
    lin, items = consolidar_lineas(lineas, agrupar)
    items = adjuntar_historico(items)
    sugeridos = [p for p in items["Últ. proveedor"].dropna().unique() if str(p).strip()][:3]
    prov = pd.DataFrame([{"Proveedor": p, **PROV_DEFECTO} for p in sugeridos],
                        columns=["Proveedor", *PROV_DEFECTO])
    ref = (f"MAT-{items.iloc[0]['Material'] or 'SINCOD'}" if len(items) == 1 else
           (lin["SOLPED"].iloc[0] if lin["SOLPED"].nunique() == 1 else "AGRUPADO")) + f"-{date.today():%Y%m%d}"
    ss.cuadro = {"ref": ref, "firma": firma(sorted(lin["SOLPED"] + "/" + lin["Pos"].astype(str)), agrupar),
                 "lineas": lin, "items": items, "prov_src": prov, "prov_v": 0, "prov_actual": prov,
                 "of_src": None, "of_v": 0, "of_last": None, "decision": {}, "importados": {},
                 "aprobacion": {"ok": False, "nombre": "", "fecha": "", "firma": ""}}


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


def limpiar_proveedores(df: pd.DataFrame) -> pd.DataFrame:
    prov = df.dropna(subset=["Proveedor"]).copy()
    prov["Proveedor"] = prov["Proveedor"].astype(str).str.strip()
    prov = prov[prov["Proveedor"] != ""].drop_duplicates("Proveedor")
    for c, v in PROV_DEFECTO.items():
        prov[c] = prov[c].fillna(v) if c in prov else v
    prov["Moneda flete"] = prov["Moneda"]          # el flete se informa en la moneda de la oferta
    return prov.reset_index(drop=True)


def incorporar_respuesta(C, datos, resp, nombre_arch):
    """Agrega o actualiza el proveedor y carga sus precios. Una segunda respuesta
    con precio distinto queda como 'Precio mejorado' (ronda de negociación)."""
    nombre = str(datos.get("Razón social") or "").strip() or nombre_arch.rsplit(".", 1)[0]
    prov = C["prov_actual"].copy()
    moneda = str(datos.get("Moneda de la oferta") or "").upper().strip()
    transporte = str(datos.get("Incoterm / forma de despacho") or "").strip()
    nuevos = {"Moneda": moneda if moneda in MONEDAS else None,
              "Transporte": transporte if transporte in TRANSPORTES else None,
              "Flete total": a_num(datos.get("Costo total de flete"), 0.0),
              "Plazo": int(a_num(datos.get("Plazo de entrega general (días)"), 0)) or None}
    if nombre in set(prov["Proveedor"]):
        for k, v in nuevos.items():
            if v is not None:
                prov.loc[prov["Proveedor"] == nombre, k] = v
    else:
        prov = pd.concat([prov, pd.DataFrame([{"Proveedor": nombre, **PROV_DEFECTO,
                                               **{k: v for k, v in nuevos.items() if v is not None}}])],
                         ignore_index=True)
    prov = limpiar_proveedores(prov)
    of = armar_ofertas(C["items"], tuple(prov["Proveedor"]), C["of_last"])
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
    C["prov_src"], C["prov_actual"], C["prov_v"] = prov, prov, C["prov_v"] + 1
    C["of_src"], C["of_last"], C["of_v"] = of, of, C["of_v"] + 1
    C["prov_names"] = tuple(prov["Proveedor"])
    return nombre, n


def cuadro_a_json(C) -> str:
    dj = lambda df: None if df is None else df.to_json(orient="records", date_format="iso", force_ascii=False)
    return json.dumps({"ref": C["ref"], "firma": C["firma"], "lineas": dj(C["lineas"]), "items": dj(C["items"]),
                       "prov": dj(C["prov_actual"]), "ofertas": dj(C["of_last"]), "decision": C["decision"],
                       "aprobacion": C["aprobacion"]}, ensure_ascii=False)


def json_a_cuadro(texto: str):
    d = json.loads(texto)
    leer = lambda s: pd.read_json(io.StringIO(s), orient="records", dtype=False) if s else None
    items = leer(d["items"])
    items["Últ. compra"] = pd.to_datetime(items["Últ. compra"], errors="coerce").dt.date
    items["Material"] = items["Material"].fillna("").astype(str)
    prov = limpiar_proveedores(leer(d["prov"]))
    of = leer(d["ofertas"])
    ss.cuadro = {"ref": d["ref"], "firma": d.get("firma", ""), "lineas": leer(d["lineas"]), "items": items,
                 "prov_src": prov, "prov_actual": prov, "prov_v": 1, "of_src": of, "of_last": of, "of_v": 1,
                 "prov_names": tuple(prov["Proveedor"]), "importados": {},
                 "decision": {int(k): v for k, v in d.get("decision", {}).items()},
                 "aprobacion": d.get("aprobacion", {"ok": False, "nombre": "", "fecha": "", "firma": ""})}


def calcular(C):
    """Comparación del cuadro actual; None si aún no hay precios o falta el tipo de cambio."""
    items = C["items"]
    items["Últ. precio CLP"] = [a_clp(p, m, tc) if pd.notna(p) else np.nan
                                for p, m in zip(items["Últ. precio"], items["Últ. moneda"])]
    items["Últ. precio CLP"] = items["Últ. precio CLP"].where(items["Últ. precio CLP"] > 0)
    of = C["of_last"]
    if of is None or not of["Precio inicial"].fillna(0).gt(0).any() or min(tc["USD"], tc["EUR"], tc["UF"]) <= 0:
        return None, None
    return calcular_comparativo(items, C["prov_actual"], of, tc)


def tabla_adjudicacion(C, det, res) -> pd.DataFrame:
    """Une la decisión del comprador con los valores calculados (insumo de reportes)."""
    src = res.merge(C["items"][["Ítem", "Texto", "Cantidad", "UM", "SOLPEDs", "Últ. precio CLP"]], on="Ítem")
    elegido, tipo, comentario = [], [], []
    for _, r in src.iterrows():
        d = C["decision"].get(int(r["Ítem"]), {})
        p = d.get("prov", r["Sugerido"])
        elegido.append(p)
        if p == r["Sugerido"]:
            tipo.append(r["Tipo sugerido"] or TIPOS_ADJUDICACION[0])
        else:
            tipo.append(d.get("tipo", "Aprobado Técnicamente"))
        comentario.append(d.get("comentario", ""))
    src["Proveedor adjudicado"], src["Tipo adjudicación"], src["Comentario"] = elegido, tipo, comentario
    adj = src.merge(det.rename(columns={"Proveedor": "Proveedor adjudicado"})[
        ["Ítem", "Proveedor adjudicado", "Moneda", "Precio mejorado", "Total CLP", "Ahorro spot CLP",
         "Var. vs últ. compra", "Entrega planta", "Válida", "Precio unit. CLP", "Unit. puesto CLP"]],
        on=["Ítem", "Proveedor adjudicado"], how="left")
    adj["Monto adjudicado CLP"] = adj["Total CLP"].fillna(0)
    adj["Ahorro spot CLP"] = adj["Ahorro spot CLP"].fillna(0)
    adj["Ahorro vs histórico CLP"] = ((adj["Últ. precio CLP"] - adj["Precio unit. CLP"]) * adj["Cantidad"]).fillna(0)
    adj["Válida"] = adj["Válida"].astype("boolean").fillna(False).astype(bool)
    return adj


def pendientes(adj) -> list:
    msgs = []
    for _, r in adj.iterrows():
        if r["Proveedor adjudicado"] == SIN_ADJ:
            msgs.append(f"Ítem {r['Ítem']} ({r['Texto'][:40]}): aún no eliges proveedor.")
        elif not r["Válida"]:
            msgs.append(f"Ítem {r['Ítem']}: {r['Proveedor adjudicado']} no tiene un precio válido.")
        elif r["Proveedor adjudicado"] != r["Sugerido"] and not str(r["Comentario"]).strip():
            msgs.append(f"Ítem {r['Ítem']}: elegiste un proveedor distinto al más conveniente; falta explicar por qué.")
    if adj["Monto adjudicado CLP"].sum() > UMBRAL_TC_CLP and tc["fecha"] != date.today().isoformat():
        msgs.append("La compra supera $1.000.000: confirma el tipo de cambio de hoy (paso 4, arriba).")
    return msgs

# =============================================================================
# ENCABEZADO Y BARRA DE PASOS
# =============================================================================
st.markdown("<div class='marca'>Enaex · Abastecimiento</div>", unsafe_allow_html=True)
st.title("Cuadro comparativo de ofertas")
barra = "".join(
    f"<div class='paso {'actual' if i == ss.paso else 'hecho' if i < ss.paso else ''}'>"
    f"<b>{'✓' if i < ss.paso else ic}</b>{i}. {txt}</div>"
    for i, (ic, txt) in enumerate(PASOS, start=1))
st.markdown(f"<div class='pasos'>{barra}</div>", unsafe_allow_html=True)


def navegacion(puede_seguir=True, motivo="", texto_siguiente=None):
    st.write("")
    st.divider()
    c1, c2, c3 = st.columns([1, 2, 1.4])
    if ss.paso > 1:
        c1.button("← Volver", on_click=ir_a, args=(ss.paso - 1,), use_container_width=True)
    if ss.paso < len(PASOS):
        etiqueta = texto_siguiente or f"Siguiente: {PASOS[ss.paso][1]} →"
        c3.button(etiqueta, type="primary", disabled=not puede_seguir, on_click=ir_a, args=(ss.paso + 1,),
                  use_container_width=True)
        if not puede_seguir and motivo:
            c2.markdown(f"<div class='sub' style='text-align:right;padding-top:.6rem'>{motivo}</div>",
                        unsafe_allow_html=True)

# =============================================================================
# PASO 1 · SUBE TUS ARCHIVOS
# =============================================================================
if ss.paso == 1:
    caja("ayuda", "Sube la planilla donde están tus SOLPED (Autogestión, exportación de ME5A o de Ariba). "
                  "Si además subes el histórico de compras, verás cuánto se pagó la última vez por cada material.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 1️⃣ Planilla de SOLPED")
        st.caption("Obligatoria · Excel o CSV")
        f_req = st.file_uploader("Planilla de SOLPED", type=["xlsx", "xlsm", "xls", "csv"],
                                 label_visibility="collapsed", key="up_req")
    with c2:
        st.markdown("### 2️⃣ Histórico de compras")
        st.caption("Opcional · Cuadro MRO (hoja Ultima COMPRA), ME2M o CSV")
        f_hist = st.file_uploader("Histórico de compras", type=["xlsx", "xlsm", "xls", "csv"],
                                  label_visibility="collapsed", key="up_hist")

    # Lectura automática: sin botón "procesar"
    for f, tipo, sins, pref in [(f_req, "req", SIN_REQ, ["Carga ARIBA", "Carga", "ME5A", "Planilla anterior"]),
                                (f_hist, "hist", SIN_HIST, ["Ultima COMPRA", "BASE OC", "ME2M", "Data"])]:
        if f is None:
            continue
        hojas = listar_hojas(f.getvalue(), f.name)
        hoja_auto = hojas[int(np.argmax([puntaje_hoja(h, pref) for h in hojas]))]
        hoja = ss.get(f"hoja_{tipo}") if ss.get(f"hoja_{tipo}") in hojas else hoja_auto
        clave_arch = firma(f.name, f.size, hoja)
        if ss.archivos.get(tipo) != clave_arch:
            df = leer_hoja(f.getvalue(), f.name, hoja, json.dumps(sins))
            ss[f"{tipo}_raw"], ss[f"{tipo}_map"] = df, mapear_columnas(df.columns, sins)
            ss[f"{tipo}_hojas"], ss[f"{tipo}_hoja_usada"] = hojas, hoja
            ss.archivos[tipo] = clave_arch

    if ss.get("hist_raw") is not None and all(c in ss.hist_map for c in HIST_OBLIGATORIAS):
        ss.hist = construir_historico(ss.hist_raw, ss.hist_map)
    if ss.get("req_raw") is not None and all(c in ss.req_map for c in REQ_OBLIGATORIAS):
        ss.req = normalizar_requerimientos(ss.req_raw, ss.req_map)

    if ss.req is not None:
        caja("ok", f"✔ Leímos <b>{ss.req['SOLPED'].nunique()} SOLPED</b> con <b>{len(ss.req)} posiciones</b> "
                   f"y <b>{ss.req['Clave'].nunique()} materiales</b> distintos.")
    elif ss.get("req_raw") is not None:
        caja("ojo", "No pudimos reconocer las columnas de la planilla. Abre <b>¿No se ve bien?</b> aquí abajo "
                    "e indica cuál es la columna de SOLPED, descripción y cantidad.")
    if ss.hist is not None:
        caja("ok", f"✔ Histórico listo: última compra de <b>{len(ss.hist[0]):,}</b> materiales.".replace(",", "."))

    if ss.get("req_raw") is not None or ss.get("hist_raw") is not None:
        with st.expander("¿No se ve bien? Ajustar cómo leímos el archivo"):
            for tipo, etiqueta, sins in [("req", "Planilla de SOLPED", SIN_REQ), ("hist", "Histórico", SIN_HIST)]:
                if ss.get(f"{tipo}_raw") is None:
                    continue
                st.markdown(f"**{etiqueta}**")
                hojas = ss[f"{tipo}_hojas"]
                st.selectbox("Hoja del Excel", hojas, index=hojas.index(ss[f"{tipo}_hoja_usada"]),
                             key=f"hoja_{tipo}")
                opciones = ["(no usar)"] + list(ss[f"{tipo}_raw"].columns)
                cols = st.columns(4)
                for i, campo in enumerate(sins):
                    actual = ss[f"{tipo}_map"].get(campo)
                    elegido = cols[i % 4].selectbox(campo, opciones, key=f"map_{tipo}_{campo}",
                                                    index=opciones.index(actual) if actual in opciones else 0)
                    nuevo_valor = None if elegido == "(no usar)" else elegido
                    if nuevo_valor != actual:
                        if nuevo_valor is None:
                            ss[f"{tipo}_map"].pop(campo, None)
                        else:
                            ss[f"{tipo}_map"][campo] = nuevo_valor
                        if tipo == "req":
                            ss.req = None
                        st.rerun()

    st.write("")
    o1, o2 = st.columns(2)
    with o1.expander("✍️ No tengo planilla: escribir los materiales a mano"):
        sp_man = st.text_input("N° de SOLPED", key="sp_manual")
        plantilla = pd.DataFrame({"Material": [""], "Texto": [""], "Cantidad": [1.0], "UM": ["C/U"], "Centro": [""]})
        pos_man = st.data_editor(plantilla, num_rows="dynamic", use_container_width=True, key="pos_manual",
                                 column_config={"Material": "Código SAP (opcional)", "Texto": "Descripción"})
        if st.button("Usar estos materiales", disabled=not sp_man.strip()):
            pm = pos_man.dropna(subset=["Texto"])
            pm = pm[pm["Texto"].astype(str).str.strip() != ""].copy()
            pm["SOLPED"], pm["Pos"] = sp_man.strip(), np.arange(1, len(pm) + 1) * 10
            pm["Material"] = pm["Material"].apply(clave_material)
            for c, v in [("Precio ref.", 0.0), ("F. solicitud", pd.NaT), ("Días", np.nan), ("Urgencia", ""),
                         ("G. compras", ""), ("Asignado a", ""), ("Contrato marco", "")]:
                pm[c] = v
            pm["Clave"] = [clave_item(m, t) for m, t in zip(pm["Material"], pm["Texto"])]
            ss.req, ss.sp_sel = pm, {sp_man.strip()}
            ir_a(2)
            st.rerun()
    with o2.expander("💾 Retomar un trabajo guardado"):
        f_json = st.file_uploader("Archivo .json guardado en el paso 5", type=["json"], key="up_json")
        if f_json is not None:
            if ss.archivos.get("json") != firma(f_json.name, f_json.size):
                json_a_cuadro(f_json.getvalue().decode("utf-8"))
                ss.archivos["json"] = firma(f_json.name, f_json.size)
            st.success(f"Trabajo «{ss.cuadro['ref']}» recuperado.")
            st.button("Continuar donde quedé →", type="primary", on_click=ir_a, args=(4,))

    navegacion(ss.req is not None, "Primero sube la planilla de SOLPED.")

# =============================================================================
# PASO 2 · ELIGE QUÉ COTIZAR
# =============================================================================
elif ss.paso == 2:
    req = ss.req
    caja("ayuda", "Marca los materiales que quieres cotizar. Si un material lo piden varias SOLPED, "
                  "<b>se juntan solas en una sola línea</b> con la cantidad total, así el proveedor cotiza una sola vez.")
    c1, c2, c3 = st.columns([3, 1.6, 1.4])
    busca = c1.text_input("🔍 Buscar", placeholder="Escribe un material, código o N° de SOLPED…",
                          label_visibility="collapsed")
    vista = c2.segmented_control("Ver", ["Por material", "Por SOLPED"], default="Por material",
                                 label_visibility="collapsed") or "Por material"
    solo_rep = c3.toggle("Solo repetidos", help="Materiales que piden 2 o más SOLPED")
    compradores = sorted(c for c in req["Asignado a"].unique() if c)
    filtro_comp = st.pills("Comprador", compradores, selection_mode="multi") if len(compradores) > 1 else []

    base = req if not filtro_comp else req[req["Asignado a"].isin(filtro_comp)]
    if busca:
        t = norm(busca)
        base = base[(base["SOLPED"] + " " + base["Material"] + " " + base["Texto"]).apply(norm).str.contains(t, regex=False)]

    sig = firma(vista, busca, solo_rep, sorted(filtro_comp or []))
    if ss.sel_sig != sig:
        ss.sel_sig, ss.sel_ver = sig, ss.sel_ver + 1
        if vista == "Por material":
            tabla = (base.groupby(["Clave", "UM"], sort=False)
                     .agg(Material=("Material", "first"), Texto=("Texto", lambda s: max(s, key=len)),
                          **{"SOLPED": ("SOLPED", "nunique"), "Cantidad": ("Cantidad", "sum"),
                             "Días": ("Días", "max"), "Detalle SOLPED": ("SOLPED", unir_unicos)})
                     .reset_index())
            if solo_rep:
                tabla = tabla[tabla["SOLPED"] > 1]
            tabla = adjuntar_historico(tabla).sort_values(["SOLPED", "Días"], ascending=False).reset_index(drop=True)
            tabla["id"] = tabla["Clave"] + "|" + tabla["UM"]
            tabla.insert(0, "Cotizar", tabla["id"].isin(ss.mat_sel))
        else:
            tabla = (base.groupby("SOLPED")
                     .agg(Materiales=("Pos", "size"), Días=("Días", "max"), Centro=("Centro", "first"),
                          Contenido=("Texto", lambda s: " · ".join(list(s)[:3]) + (" …" if len(s) > 3 else "")))
                     .reset_index().sort_values("Días", ascending=False).reset_index(drop=True))
            if solo_rep:
                rep = base.groupby(["Clave", "UM"])["SOLPED"].transform("nunique") > 1
                tabla = tabla[tabla["SOLPED"].isin(base[rep]["SOLPED"])].reset_index(drop=True)
            tabla["id"] = tabla["SOLPED"]
            tabla.insert(0, "Cotizar", tabla["id"].isin(ss.sp_sel))
        ss.sel_tabla = tabla

    tabla = ss.sel_tabla
    if vista == "Por material":
        cfg = {"Cotizar": st.column_config.CheckboxColumn("✔", width="small"),
               "Texto": st.column_config.TextColumn("Material", width="large"),
               "Material": st.column_config.TextColumn("Código"),
               "SOLPED": st.column_config.NumberColumn("Lo piden", format="%d SOLPED"),
               "Cantidad": st.column_config.NumberColumn("Cantidad total", format="%g"),
               "Días": st.column_config.NumberColumn("Antigüedad", format="%d días"),
               "Últ. proveedor": st.column_config.TextColumn("Último proveedor"),
               "Últ. precio": st.column_config.NumberColumn("Último precio", format="localized")}
        orden = ["Cotizar", "Texto", "Material", "UM", "SOLPED", "Cantidad", "Días", "Últ. proveedor", "Últ. precio",
                 "Últ. moneda"]
    else:
        cfg = {"Cotizar": st.column_config.CheckboxColumn("✔", width="small"),
               "Materiales": st.column_config.NumberColumn("Posiciones", format="%d"),
               "Días": st.column_config.NumberColumn("Antigüedad", format="%d días"),
               "Contenido": st.column_config.TextColumn("Qué pide", width="large")}
        orden = ["Cotizar", "SOLPED", "Contenido", "Materiales", "Días", "Centro"]
    editado = st.data_editor(tabla, key=f"sel_{ss.sel_ver}", hide_index=True, use_container_width=True,
                             height=420, column_order=orden, column_config=cfg,
                             disabled=[c for c in tabla.columns if c != "Cotizar"])
    destino = ss.mat_sel if vista == "Por material" else ss.sp_sel
    for i, marcado in zip(editado["id"], editado["Cotizar"]):
        (destino.add if marcado else destino.discard)(i)

    lin = lineas_elegidas()
    if lin.empty:
        caja("sub", "Marca la casilla ✔ de lo que quieras cotizar.")
    else:
        _, items_prev = consolidar_lineas(lin, True)
        st.markdown(f"<div class='mini'><div><b>{len(items_prev)}</b>ítem(s) a cotizar</div>"
                    f"<div><b>{lin['SOLPED'].nunique()}</b>SOLPED incluidas</div>"
                    f"<div><b>{len(lin)}</b>posiciones</div></div>", unsafe_allow_html=True)
        with st.expander("Ver detalle de lo elegido"):
            st.dataframe(lin[["SOLPED", "Pos", "Material", "Texto", "Cantidad", "UM", "Centro"]],
                         hide_index=True, use_container_width=True)
            if st.button("Quitar todo y empezar de nuevo"):
                ss.mat_sel, ss.sp_sel, ss.sel_sig = set(), set(), None
                st.rerun()

    def preparar_y_seguir():
        lineas = lineas_elegidas()
        f = firma(sorted(lineas["SOLPED"] + "/" + lineas["Pos"].astype(str)), True)
        if ss.cuadro is None or ss.cuadro["firma"] != f:
            crear_cuadro(lineas, True)
        ir_a(3)

    st.write("")
    st.divider()
    b1, b2, b3 = st.columns([1, 2, 1.4])
    b1.button("← Volver", on_click=ir_a, args=(1,), use_container_width=True)
    if ss.cuadro is not None and not lin.empty and ss.cuadro["firma"] != firma(
            sorted(lin["SOLPED"] + "/" + lin["Pos"].astype(str)), True):
        b2.markdown("<div class='sub' style='text-align:right;padding-top:.6rem'>Cambiaste la selección: "
                    "se empezará un cuadro nuevo.</div>", unsafe_allow_html=True)
    b3.button("Siguiente: Pide la cotización →", type="primary", disabled=lin.empty,
              on_click=preparar_y_seguir, use_container_width=True)

# =============================================================================
# PASO 3 · PIDE LA COTIZACIÓN
# =============================================================================
elif ss.paso == 3:
    C = ss.cuadro
    items, lineas = C["items"], C["lineas"]
    st.markdown(f"<div class='mini'><div><b>{len(items)}</b>ítem(s)</div><div><b>{lineas['SOLPED'].nunique()}</b>"
                f"SOLPED</div><div><b>{len(C['prov_actual'])}</b>proveedor(es)</div></div>", unsafe_allow_html=True)

    st.markdown("### ① Descarga la planilla y envíala a tus proveedores")
    c1, c2, c3 = st.columns(3)
    comprador = c1.text_input("Tu nombre", value=ss.get("mi_nombre", unir_unicos(lineas.get("Asignado a", []))))
    correo = c2.text_input("Tu correo (para que te respondan)", value=ss.get("mi_correo", ""))
    limite = c3.date_input("Que respondan antes del", value=date.today() + timedelta(days=5), format="DD/MM/YYYY")
    ss.mi_nombre, ss.mi_correo = comprador, correo
    st.download_button("⬇️  Descargar planilla para proveedores", type="primary", use_container_width=True,
                       data=generar_rfq(C["ref"], items, lineas,
                                        {"comprador": comprador, "correo": correo, "fecha_limite": limite,
                                         "entrega": unir_unicos(items["Centro"])}),
                       file_name=f"Solicitud_cotizacion_{nombre_archivo(C['ref'])}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("La planilla no muestra precios históricos ni información interna. El proveedor completa solo las "
               "celdas amarillas y te la devuelve.")

    st.markdown("### ② Cuando te respondan, sube aquí sus planillas")
    respuestas = st.file_uploader("Planillas respondidas", type=["xlsx"], accept_multiple_files=True,
                                  label_visibility="collapsed", key=f"resp_{C['ref']}")
    for f in respuestas or []:
        clave_f = firma(f.name, f.size)
        if clave_f in C["importados"]:
            continue
        try:
            ref_resp, datos, resp = leer_respuesta_rfq(f.getvalue())
            if ref_resp and ref_resp != C["ref"]:
                C["importados"][clave_f] = ("error", f"{f.name}: es de otra cotización ({ref_resp}).")
            else:
                nombre, n = incorporar_respuesta(C, datos, resp, f.name)
                C["importados"][clave_f] = ("ok", f"{nombre}: {n} precio(s) cargados.")
        except Exception as e:
            C["importados"][clave_f] = ("error", f"{f.name}: no se pudo leer ({e}).")
    for estado, msg in C["importados"].values():
        caja("ok" if estado == "ok" else "ojo", ("✔ " if estado == "ok" else "⚠ ") + html.escape(msg))

    st.markdown("### ③ Revisa los precios recibidos")
    st.caption("Se llenan solos al subir las planillas. Si te respondieron por correo o PDF, escríbelos aquí. "
               "Si un proveedor mejora su oferta, anota el nuevo valor en «Precio mejorado».")

    st.markdown("**Proveedores**")
    prov_edit = st.data_editor(
        C["prov_src"], key=f"prov_{C['ref']}_{C['prov_v']}", num_rows="dynamic", hide_index=True,
        use_container_width=True, column_order=COLS_PROV_VISIBLES,
        column_config={
            "Proveedor": st.column_config.TextColumn("Proveedor", required=True, width="large"),
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=MONEDAS, required=True, default="CLP"),
            "Flete total": st.column_config.NumberColumn("Flete (total)", min_value=0.0, default=0.0,
                                                         help="Costo de envío total, en la moneda de la oferta"),
            "Plazo": st.column_config.NumberColumn("Plazo (días hábiles)", min_value=0, step=1, default=10),
            "Tipo compra": st.column_config.SelectboxColumn("Origen", options=TIPOS_COMPRA, default="Nacional",
                                                            help="Importación suma 10 días de internación"),
            "Lugar entrega": st.column_config.SelectboxColumn("Dónde entrega", options=list(LUGARES_ENTREGA),
                                                              default=list(LUGARES_ENTREGA)[0]),
        })
    prov = limpiar_proveedores(prov_edit)
    C["prov_actual"] = prov
    nombres = tuple(prov["Proveedor"])
    if C.get("prov_names") != nombres or C["of_src"] is None:
        C["of_src"] = armar_ofertas(items, nombres, C["of_last"])
        C["of_v"] += 1
        C["prov_names"] = nombres

    if not nombres:
        caja("sub", "Agrega al menos un proveedor (o sube una planilla respondida).")
    else:
        st.markdown("**Precios por ítem** (precio unitario, en la moneda del proveedor)")
        of_edit = st.data_editor(
            C["of_src"], key=f"of_{C['ref']}_{C['of_v']}", hide_index=True, use_container_width=True,
            height=min(80 + 35 * len(C["of_src"]), 520), disabled=["Ítem", "Texto", "Cantidad", "UM", "Proveedor"],
            column_order=["Proveedor", "Texto", "Cantidad", "UM", "Precio inicial", "Precio mejorado", "Plazo",
                          "Cumple técnico", "Observación"],
            column_config={"Texto": st.column_config.TextColumn("Material", width="large"),
                           "Cantidad": st.column_config.NumberColumn(format="%g"),
                           "Precio inicial": st.column_config.NumberColumn("Precio", min_value=0.0, format="%.2f"),
                           "Precio mejorado": st.column_config.NumberColumn("Precio mejorado", min_value=0.0,
                                                                            format="%.2f",
                                                                            help="Opcional: si negociaste un mejor precio"),
                           "Plazo": st.column_config.NumberColumn("Plazo (días)", min_value=0, step=1,
                                                                  help="Vacío = plazo general del proveedor"),
                           "Cumple técnico": st.column_config.CheckboxColumn("¿Cumple?",
                                                                             help="Desmarca si no cumple la ficha técnica"),
                           "Observación": st.column_config.TextColumn("Notas")})
        C["of_last"] = of_edit

    hay_precios = C.get("of_last") is not None and C["of_last"]["Precio inicial"].fillna(0).gt(0).any()
    navegacion(hay_precios, "Necesitas al menos un precio para comparar.")

# =============================================================================
# PASO 4 · COMPARA Y ELIGE
# =============================================================================
elif ss.paso == 4:
    C = ss.cuadro

    def campos_tc():
        nombres_tc = {"USD": "Dólar (USD)", "EUR": "Euro (EUR)", "UF": "UF"}
        cols_tc = st.columns(3)
        nuevo = {k: cols_tc[i].number_input(nombres_tc[k], value=float(tc[k]), min_value=0.0, step=1.0,
                                            format="%.2f", key=f"tc_in_{k}")
                 for i, k in enumerate(["USD", "EUR", "UF"])}
        if any(abs(nuevo[k] - tc[k]) > 1e-9 for k in nuevo):
            ss.tc = {**nuevo, "fecha": date.today().isoformat(), "fuente": "Ingreso manual"}
            st.rerun()

    if min(tc["USD"], tc["EUR"], tc["UF"]) <= 0:
        caja("ojo", "No pudimos traer el tipo de cambio desde internet. Escríbelo aquí (valor de hoy en pesos):")
        campos_tc()
    else:
        t1, t2 = st.columns([4, 1])
        hoy_txt = "de hoy" if tc["fecha"] == date.today().isoformat() else f"del {tc['fecha'] or '—'}"
        t1.markdown(f"<div class='sub'>Tipo de cambio {hoy_txt}: Dólar <b>{dinero(tc['USD'])}</b> · "
                    f"Euro <b>{dinero(tc['EUR'])}</b> · UF <b>{dinero(tc['UF'])}</b></div>", unsafe_allow_html=True)
        with t2.popover("Cambiar", use_container_width=True):
            campos_tc()
            if st.button("Traer de internet"):
                obtener_indicadores.clear()
                ind = obtener_indicadores()
                if ind["ok"]:
                    ss.tc = {"USD": ind["USD"], "EUR": ind["EUR"], "UF": ind["UF"], "fecha": ind["fecha"],
                             "fuente": ind["fuente"]}
                    for k in ["USD", "EUR", "UF"]:
                        ss.pop(f"tc_in_{k}", None)
                    st.rerun()
                st.warning("Sin conexión: escríbelo a mano.")

    det, res = calcular(C)
    if det is None:
        if min(tc["USD"], tc["EUR"], tc["UF"]) > 0:
            caja("ojo", "Aún no hay precios cargados. Vuelve al paso 3 y sube las respuestas de los proveedores.")
        navegacion(False, "Completa el tipo de cambio para ver la comparación.")
        st.stop()
    C["det"], C["res"] = det, res
    items = C["items"]
    caja("ayuda", "Para cada ítem te mostramos el <b>total puesto en planta</b> de cada proveedor (incluye el flete). "
                  "El verde es el más conveniente y ya viene elegido; puedes cambiarlo si tienes un motivo.")

    for _, it in items.iterrows():
        n_item = int(it["Ítem"])
        r = res[res["Ítem"] == n_item].iloc[0]
        ofertas = det[det["Ítem"] == n_item].sort_values(["Válida", "Total CLP"], ascending=[False, True])
        validas = ofertas[ofertas["Válida"]]
        maximo = validas["Total CLP"].max() if not validas.empty else 1
        with st.container(border=True):
            st.markdown(f"#### {n_item}. {html.escape(str(it['Texto']))}")
            ult = (f"Última compra: <b>{dinero(it['Últ. precio'], it['Últ. moneda'] or 'CLP')}</b> c/u a "
                   f"{html.escape(str(it['Últ. proveedor']))} ({it['Últ. compra']:%d/%m/%Y})"
                   if it["Con historial"] and pd.notna(it["Últ. precio"]) and pd.notna(it["Últ. compra"])
                   else "<span class='etiqueta roja'>Sin compras anteriores · requiere visto bueno técnico</span>")
            st.markdown(f"<div class='sub'>{it['Cantidad']:g} {html.escape(str(it['UM']))} · "
                        f"{int(it['N° SOLPED'])} SOLPED · {ult}</div>", unsafe_allow_html=True)
            filas = []
            for _, o in ofertas.iterrows():
                nom = html.escape(str(o["Proveedor"]))
                if not o["Cotiza"]:
                    filas.append(f"<div class='fila-oferta'><div>{nom}</div><div class='sub'>No cotizó</div>"
                                 f"<div></div><div></div></div>")
                    continue
                gana = o["Proveedor"] == r["Sugerido"]
                ancho = max(6, o["Total CLP"] / maximo * 100) if o["Válida"] else 0
                badge = ("<span class='etiqueta'>Más conveniente</span>" if gana else
                         "" if o["Válida"] else "<span class='etiqueta roja'>No cumple</span>")
                var = o["Var. vs últ. compra"]
                var_txt = ("" if pd.isna(var) else
                           f"<span class='etiqueta {'roja' if var > UMBRAL_ALZA else 'gris' if var > 0 else ''}'>"
                           f"{abs(var):.0%} {'más caro' if var > 0 else 'más barato'} que la última vez</span>")
                filas.append(
                    f"<div class='fila-oferta'><div><b>{nom}</b>{badge}</div>"
                    f"<div><div class='barra {'gana' if gana else ''}' style='width:{ancho:.0f}%'></div>{var_txt}</div>"
                    f"<div style='text-align:right'><b>{dinero(o['Total CLP'])}</b><br>"
                    f"<span class='sub'>{dinero(o['Unit. puesto CLP'])} c/u</span></div>"
                    f"<div class='sub'>Llega {o['Entrega planta']:%d/%m/%Y}</div></div>")
            st.markdown("".join(filas), unsafe_allow_html=True)

            opciones = list(validas["Proveedor"]) + [SIN_ADJ]
            previo = C["decision"].get(n_item, {}).get("prov", r["Sugerido"])
            elegido = st.radio("¿A quién le compras?", opciones, horizontal=True, key=f"dec_{C['ref']}_{n_item}",
                               index=opciones.index(previo) if previo in opciones else len(opciones) - 1)
            dec = {"prov": elegido}
            if elegido not in (r["Sugerido"], SIN_ADJ):
                e1, e2 = st.columns([1, 2])
                dec["tipo"] = e1.selectbox("Motivo", ["Aprobado Técnicamente", "Asignación Directa"],
                                           key=f"tipo_{C['ref']}_{n_item}",
                                           index=["Aprobado Técnicamente", "Asignación Directa"].index(
                                               C["decision"].get(n_item, {}).get("tipo", "Aprobado Técnicamente")))
                dec["comentario"] = e2.text_input("¿Por qué eliges este proveedor? (obligatorio)",
                                                  value=C["decision"].get(n_item, {}).get("comentario", ""),
                                                  key=f"com_{C['ref']}_{n_item}")
            else:
                dec["comentario"] = ""
            C["decision"][n_item] = dec

    adj = tabla_adjudicacion(C, det, res)
    C["adj"] = adj
    st.markdown(f"<div class='mini'><div><b>{dinero(adj['Monto adjudicado CLP'].sum())}</b>total a comprar</div>"
                f"<div><b>{dinero(adj['Ahorro spot CLP'].sum())}</b>ahorro por negociación</div>"
                f"<div><b>{dinero(adj['Ahorro vs histórico CLP'].sum())}</b>ahorro vs última compra</div></div>",
                unsafe_allow_html=True)
    navegacion(True)

# =============================================================================
# PASO 5 · APRUEBA Y DESCARGA
# =============================================================================
elif ss.paso == 5:
    C = ss.cuadro
    det, res = calcular(C)
    if det is None:
        caja("ojo", "Faltan precios o el tipo de cambio. Vuelve al paso anterior.")
        navegacion(False)
        st.stop()
    adj = tabla_adjudicacion(C, det, res)
    items, lineas, prov = C["items"], C["lineas"], C["prov_actual"]
    total = adj["Monto adjudicado CLP"].sum()
    st.markdown(f"<div class='mini'><div><b>{dinero(total)}</b>total a comprar</div>"
                f"<div><b>{int((adj['Proveedor adjudicado'] != SIN_ADJ).sum())} de {len(adj)}</b>ítems con proveedor</div>"
                f"<div><b>{dinero(adj['Ahorro spot CLP'].sum())}</b>ahorro por negociación</div>"
                f"<div><b>{dinero(adj['Ahorro vs histórico CLP'].sum())}</b>ahorro vs última compra</div></div>",
                unsafe_allow_html=True)

    st.markdown("### Resumen de la compra")
    resumen = adj[["Ítem", "Texto", "Cantidad", "UM", "Proveedor adjudicado", "Monto adjudicado CLP",
                   "Entrega planta", "Tipo adjudicación"]]
    st.dataframe(resumen, hide_index=True, use_container_width=True,
                 column_config={"Texto": st.column_config.TextColumn("Material", width="large"),
                                "Proveedor adjudicado": "Proveedor",
                                "Monto adjudicado CLP": st.column_config.NumberColumn("Total ($)", format="localized"),
                                "Entrega planta": st.column_config.DateColumn("Llega", format="DD/MM/YYYY")})
    with st.expander("Ver cómo se reparte por SOLPED (para emitir las OC)"):
        st.dataframe(detalle_por_solped(lineas, adj), hide_index=True, use_container_width=True,
                     column_config={"Precio unit. (moneda oferta)": st.column_config.NumberColumn(format="localized"),
                                    "Unit. puesto CLP": st.column_config.NumberColumn("Unitario puesto ($)", format="localized"),
                                    "Monto línea CLP": st.column_config.NumberColumn("Monto línea ($)", format="localized")})

    st.markdown("### Aprobación")
    faltan = pendientes(adj)
    huella = firma(adj[["Ítem", "Proveedor adjudicado", "Monto adjudicado CLP", "Comentario"]].to_json(),
                   tc["USD"], tc["EUR"], tc["UF"])
    ap = C["aprobacion"]
    aprobado = ap.get("ok") and ap.get("firma") == huella
    if faltan:
        caja("ojo", "<b>Antes de aprobar, revisa:</b><br>" + "<br>".join("• " + html.escape(m) for m in faltan))
    elif aprobado:
        caja("ok", f"✔ Aprobado por <b>{html.escape(ap['nombre'])}</b> el {ap['fecha']}.")
    else:
        a1, a2 = st.columns([2, 1])
        quien = a1.text_input("Tu nombre", value=ss.get("mi_nombre", ""), placeholder="Quién aprueba")
        if a2.button("✅ Aprobar compra", type="primary", disabled=not quien.strip(), use_container_width=True):
            C["aprobacion"] = {"ok": True, "nombre": quien.strip(),
                               "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"), "firma": huella}
            st.rerun()
        st.caption("Al aprobar, los reportes dejan de decir BORRADOR. Si cambias algo después, habrá que aprobar de nuevo.")

    st.markdown("### Descargas")
    ap_rep = C["aprobacion"] if aprobado else {"ok": False, "nombre": "", "fecha": ""}
    sufijo = "" if aprobado else "_BORRADOR"
    base_nombre = f"{nombre_archivo(C['ref'])}_{date.today():%Y%m%d}{sufijo}"
    d1, d2, d3 = st.columns(3)
    d1.download_button("📊  Cuadro comparativo (Excel)", type="primary", use_container_width=True,
                       data=generar_excel(C["ref"], items, lineas, prov, det, adj, tc, ap_rep),
                       file_name=f"Cuadro_comparativo_{base_nombre}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("📄  Resumen (PDF)", use_container_width=True,
                       data=generar_pdf(C["ref"], adj, lineas, tc, ap_rep),
                       file_name=f"Cuadro_comparativo_{base_nombre}.pdf", mime="application/pdf")
    d3.download_button("💾  Guardar para seguir después", use_container_width=True,
                       data=cuadro_a_json(C), file_name=f"Trabajo_{nombre_archivo(C['ref'])}.json",
                       mime="application/json",
                       help="Para retomarlo: paso 1 → «Retomar un trabajo guardado»")

    st.write("")
    st.divider()
    n1, _, n3 = st.columns([1, 2, 1.4])
    n1.button("← Volver", on_click=ir_a, args=(4,), use_container_width=True)

    def nueva_cotizacion():
        ss.cuadro, ss.mat_sel, ss.sp_sel = None, set(), set()
        ir_a(2)

    n3.button("Empezar otra cotización", on_click=nueva_cotizacion, use_container_width=True)
