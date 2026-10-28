# =============================================================================
# CUADRO COMPARATIVO DE OFERTAS - ABASTECIMIENTO ENAEX  (v5 · asistente guiado)
# -----------------------------------------------------------------------------
# Un asistente de 5 pasos, un paso a la vez:
#   1. Sube tus archivos        (planilla de SOLPED + histórico opcional)
#   2. Elige qué cotizar        (por material o SOLPED; tú decides qué posiciones entran)
#   3. Ingresa las ofertas      (precio, moneda, transporte y plazo por proveedor)
#   4. Compara y elige          (tarjeta por ítem con el más conveniente)
#   5. Aprueba y descarga       (Excel, PDF y guardar para seguir después)
#
# requirements.txt:
#   streamlit>=1.46  pandas>=2.1  numpy  openpyxl  requests  fpdf2
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
from pathlib import Path

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

# Íconos dibujados en SVG (trazo, heredan el color del texto con currentColor)
_SVG = {
    "carpeta": '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4'
               'a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
    "buscar": '<circle cx="11" cy="11" r="7.5"/><path d="m21 21-4.3-4.3"/>',
    "editar": '<path d="M12.5 22H18a2 2 0 0 0 2-2V7l-5-5H6a2 2 0 0 0-2 2v9.5"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/>'
              '<path d="M13.4 15.6a1 1 0 1 0-3-3l-5 5a2 2 0 0 0-.5.85l-.84 2.87a.5.5 0 0 0 .62.62l2.87-.84'
              'a2 2 0 0 0 .85-.5z"/>',
    "balanza": '<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1'
               's-2.13-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>',
    "aprobado": '<circle cx="12" cy="12" r="10"/><path d="m8.5 12 2.5 2.5 4.5-5"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "alerta": '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/>'
              '<path d="M12 9v4"/><path d="M12 17h.01"/>',
    "empresa": '<rect width="16" height="20" x="4" y="2" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01M12 6h.01'
               'M16 6h.01M8 10h.01M12 10h.01M16 10h.01M8 14h.01M12 14h.01M16 14h.01"/>',
    "mas": '<path d="M5 12h14"/><path d="M12 5v14"/>',
}


def ico(nombre: str, tam: float = 1.1, clase: str = "") -> str:
    """Devuelve un ícono SVG en línea; tam en em para que escale con el texto que lo rodea."""
    return (f"<svg class='ico {clase}' width='{tam}em' height='{tam}em' viewBox='0 0 24 24' fill='none' "
            f"stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round' "
            f"aria-hidden='true'>{_SVG[nombre]}</svg>")


def num(n) -> str:
    """Número dentro de una insignia redonda, para pasos y secciones numeradas."""
    return f"<span class='num'>{n}</span>"


PASOS = [("carpeta", "Sube tus archivos"), ("buscar", "Elige qué cotizar"), ("editar", "Ingresa las ofertas"),
         ("balanza", "Compara y elige"), ("aprobado", "Aprueba y descarga")]
AGREGAR = ":material/add: Agregar proveedor"

# =============================================================================
# PÁGINA Y ESTILO VISUAL
# =============================================================================
st.set_page_config(page_title="Cuadro Comparativo - Enaex", page_icon=":material/balance:", layout="wide",
                   initial_sidebar_state="collapsed")
st.markdown("""
<style>
  :root { --rojo: #9E2B2F; --rojo-osc: #7F1D1D; }
  [data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none; }
  [data-testid="stMainBlockContainer"] { max-width: 1200px; padding-top: 2.6rem; padding-bottom: 5rem; }
  .banner { background: linear-gradient(135deg, var(--rojo) 0%, var(--rojo-osc) 100%); border-radius: 16px;
            padding: 1.6rem 2rem; color: #FFFFFF; margin: .4rem 0 1.1rem 0;
            box-shadow: 0 6px 18px rgba(127,29,29,.18);
            display: grid; grid-template-columns: 1fr auto; align-items: center; column-gap: 2rem; row-gap: 1rem; }
  .banner .izq { display: flex; flex-direction: column; justify-content: center; gap: .35rem; }
  .banner .marca { font-size: .78rem; letter-spacing: .12em; text-transform: uppercase; opacity: .85; }
  .banner .titulo { font-size: 1.75rem; font-weight: 700; line-height: 1.15; }
  .banner .fuente { font-size: .74rem; opacity: .72; margin-top: .2rem; }
  .tasas { display: grid; grid-template-columns: repeat(3, 128px); gap: .6rem; }
  .tasa { background: rgba(0,0,0,.2); border: 1px solid rgba(255,255,255,.08); border-radius: 12px;
          padding: .65rem .5rem; text-align: center; }
  .tasa span { display: block; font-size: .7rem; opacity: .85; text-transform: uppercase; letter-spacing: .08em;
               margin-bottom: .15rem; }
  .tasa b { font-size: 1.2rem; font-variant-numeric: tabular-nums; }
  .tasa.usd b { color: #7DD3FC; } .tasa.eur b { color: #FDE047; } .tasa.uf b { color: #FDBA74; }
  @media (max-width: 860px) {
    .banner { grid-template-columns: 1fr; padding: 1.3rem 1.3rem; }
    .tasas { grid-template-columns: repeat(3, 1fr); }
  }
  .pasos { display: flex; gap: .5rem; margin: 1.3rem 0 1.2rem 0; }
  .paso { flex: 1; text-align: center; padding: .6rem .3rem; border-radius: 12px; font-size: .88rem;
          background: rgba(128,128,128,.12); opacity: .75; border: 2px solid transparent; }
  .paso .ico-paso { display: flex; justify-content: center; margin-bottom: .3rem; }
  .paso.hecho .ico-paso { color: #059669; }
  .paso.actual .ico-paso { color: var(--rojo); }
  .ico { display: inline-block; vertical-align: -0.18em; flex-shrink: 0; }
  .num { display: inline-flex; align-items: center; justify-content: center; min-width: 1.55em; height: 1.55em;
         padding: 0 .3em; border-radius: 7px; background: var(--rojo); color: #FFFFFF; font-size: .8em;
         font-weight: 700; line-height: 1; margin-right: .45rem; vertical-align: .12em; }
  .ok .ico { color: #059669; margin-right: .35rem; }
  .ojo .ico, .alerta-tc .ico { color: #D97706; margin-right: .35rem; }
  .titulo-ico { color: var(--rojo); margin-right: .4rem; }
  .paso.hecho { background: rgba(5,150,105,.16); opacity: 1; }
  .paso.actual { background: rgba(158,43,47,.14); border-color: var(--rojo); opacity: 1; font-weight: 700; }
  .ayuda { background: rgba(30,58,138,.08); border-left: 4px solid var(--rojo); padding: .8rem 1rem;
           border-radius: 8px; margin-bottom: 1rem; }
  .ok { background: rgba(5,150,105,.13); border-left: 4px solid #059669; padding: .7rem 1rem; border-radius: 8px; margin: .4rem 0; }
  .ojo { background: rgba(217,119,6,.15); border-left: 4px solid #D97706; padding: .7rem 1rem; border-radius: 8px; margin: .4rem 0; }
  .mini { display: flex; gap: .7rem; margin: .5rem 0 1rem 0; flex-wrap: wrap; }
  .mini div { flex: 1; min-width: 150px; background: var(--rojo); border-radius: 12px; padding: .85rem .6rem;
              text-align: center; color: rgba(255,255,255,.92); font-size: .9rem;
              box-shadow: inset 0 -3px 0 rgba(0,0,0,.18); }
  .mini div b { display: block; font-size: 1.55rem; line-height: 1.3; }
  .mini .c1 { color: #7DD3FC; } .mini .c2 { color: #FDE047; } .mini .c3 { color: #FDBA74; }
  .mini .c4 { color: #A5F3FC; } .mini .c5 { color: #86EFAC; } .mini .c6 { color: #F9A8D4; }
  .fila-oferta { display: grid; grid-template-columns: 220px 1fr 150px 120px; gap: .6rem; align-items: center;
                 padding: .35rem 0; font-size: .92rem; }
  .barra { height: 14px; border-radius: 7px; background: #94A3B8; }
  .barra.gana { background: #059669; }
  .etiqueta { font-size: .75rem; padding: .1rem .5rem; border-radius: 999px; background: #D1FAE5; color: #065F46;
              margin-left: .4rem; white-space: nowrap; }
  .etiqueta.gris { background: #E5E7EB; color: #374151; }
  .etiqueta.roja { background: #FEE2E2; color: #991B1B; }
  .sub { opacity: .72; font-size: .88rem; }
  div.stButton > button, div.stDownloadButton > button { border-radius: 10px; padding: .55rem 1.1rem; }
  button[kind="primary"], [data-testid="stBaseButton-primary"] {
      background: var(--rojo) !important; border-color: var(--rojo) !important; color: #FFFFFF !important; font-weight: 600; }
  button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover { background: var(--rojo-osc) !important; }
  button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled {
      background: rgba(128,128,128,.25) !important; border-color: transparent !important; color: rgba(128,128,128,.9) !important; }
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
        "Moneda ref.": [parse_monto(v)[1] or (str(mo).upper().strip() if str(mo).upper().strip() in MONEDAS else "CLP")
                        for v, mo in zip(g("precio"), g("moneda"))],
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
    for c, v in [("Precio ref.", 0.0), ("Moneda ref.", "CLP")]:
        if c not in l:
            l[c] = v
    l["Precio ref."] = pd.to_numeric(l["Precio ref."], errors="coerce").fillna(0.0)
    l["Moneda ref."] = l["Moneda ref."].fillna("CLP").astype(str)
    items = (l.groupby("Ítem", sort=True)
             .agg(Material=("Material", "first"),
                  Texto=("Texto", lambda s: max(s.astype(str), key=len)),
                  Cantidad=("Cantidad", "sum"),
                  UM=("UM", "first"),
                  Centro=("Centro", unir_unicos),
                  SOLPEDs=("Ref", lambda s: ", ".join(s)),
                  **{"N° SOLPED": ("SOLPED", "nunique"),
                     "Precio SOLPED": ("Precio ref.", lambda s: next((float(x) for x in s if x and x > 0), 0.0)),
                     "Moneda SOLPED": ("Moneda ref.", "first")})
             .reset_index())
    return l, items

# =============================================================================
# TIPOS DE CAMBIO Y FECHAS
# =============================================================================
CACHE_TC = Path(__file__).with_name("tipo_cambio_guardado.json")
NOMBRE_TC = {"USD": "Dólar", "EUR": "Euro", "UF": "UF"}
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre",
         "octubre", "noviembre", "diciembre"]
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}


def leer_cache_tc() -> dict:
    try:
        return json.loads(CACHE_TC.read_text(encoding="utf-8"))
    except Exception:
        return {}


def guardar_cache_tc(valores: dict):
    """Guarda el último valor conocido de cada moneda (sirve si un día falla internet)."""
    if not valores:
        return
    try:
        datos = leer_cache_tc()
        for k, (v, f) in valores.items():
            if v and v > 0:
                datos[k] = {"valor": float(v), "fecha": str(f)}
        CACHE_TC.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass


def valor_sii(texto: str, dia: date):
    """Lee el valor de un día en las tablas anuales del SII (UF o dólar observado).
    Soporta el formato por mes (div id='mes_enero' con pares día/valor) y la tabla anual
    (columnas Ene…Dic, una fila por día). Si el día no tiene valor (fin de semana), toma el anterior."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(texto, "html.parser")

    def buscar(d: date):
        cont = soup.find(id=f"mes_{MESES[d.month - 1]}")
        if cont is not None:
            celdas = cont.find_all(["th", "td"])
            for a, b in zip(celdas, celdas[1:]):
                if a.get_text(strip=True) == str(d.day) and b.name == "td":
                    v = parse_monto(b.get_text(strip=True))[0]
                    if v > 0:
                        return v
        for tabla in soup.find_all("table"):
            filas = tabla.find_all("tr")
            if not filas:
                continue
            encabezado = [norm(c.get_text()) for c in filas[0].find_all(["th", "td"])]
            col = next((i for i, h in enumerate(encabezado) if h[:3] == MESES[d.month - 1][:3]), None)
            if col is None or not any(h.startswith("ene") for h in encabezado):
                continue
            for fila in filas[1:]:
                celdas = fila.find_all(["th", "td"])
                if celdas and celdas[0].get_text(strip=True) == str(d.day) and len(celdas) > col:
                    v = parse_monto(celdas[col].get_text(strip=True))[0]
                    if v > 0:
                        return v
        return None

    for atras in range(0, 8):
        d = dia - timedelta(days=atras)
        if d.year != dia.year:
            break
        v = buscar(d)
        if v:
            return v, d
    return None, None


@st.cache_data(ttl=900, show_spinner=False)
def obtener_indicadores():
    """Dólar, euro y UF del día, moneda por moneda, probando varias fuentes:
       1) mindicador.cl (Banco Central)   2) SII (UF y dólar observado)
       3) open.er-api.com (dólar y euro)  4) último valor guardado en este computador."""
    hoy = date.today()
    hora = datetime.now().strftime("%H:%M")
    val = {"USD": 0.0, "EUR": 0.0, "UF": 0.0}
    fuente, fecha = {}, {}

    def poner(k, v, f, fe):
        if v and v > 0 and not val[k]:
            val[k], fuente[k], fecha[k] = float(v), f, str(fe)

    try:
        d = requests.get("https://mindicador.cl/api", timeout=10, headers=UA).json()
        for k, clave in [("USD", "dolar"), ("EUR", "euro"), ("UF", "uf")]:
            poner(k, d[clave]["valor"], "Banco Central", str(d[clave]["fecha"])[:10])
    except Exception:
        pass
    if not val["UF"]:
        try:
            d = requests.get(f"https://mindicador.cl/api/uf/{hoy:%d-%m-%Y}", timeout=10, headers=UA).json()
            poner("UF", d["serie"][0]["valor"], "Banco Central", str(d["serie"][0]["fecha"])[:10])
        except Exception:
            pass
    for k, pagina in [("UF", "uf/uf"), ("USD", "dolar/dolar")]:
        if not val[k]:
            try:
                r = requests.get(f"https://www.sii.cl/valores_y_fechas/{pagina}{hoy.year}.htm", timeout=12, headers=UA)
                r.encoding = r.apparent_encoding or "latin-1"
                v, dia = valor_sii(r.text, hoy)
                poner(k, v, "SII", dia)
            except Exception:
                pass
    if not val["USD"] or not val["EUR"]:
        try:
            tasas = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10, headers=UA).json()["rates"]
            poner("USD", tasas["CLP"], "open.er-api.com", hoy)
            poner("EUR", float(tasas["CLP"]) / float(tasas["EUR"]), "open.er-api.com", hoy)
        except Exception:
            pass
    en_linea = {k: (val[k], fecha[k]) for k in val if val[k]}
    guardar_cache_tc(en_linea)
    cache = leer_cache_tc()
    for k in val:
        if not val[k] and cache.get(k):
            poner(k, cache[k]["valor"], f"último valor guardado ({cache[k]['fecha']})", cache[k]["fecha"])
    return {**val, "fuente_por_moneda": fuente, "fecha_por_moneda": fecha,
            "fecha": fecha.get("USD", hoy.isoformat()), "hora": hora, "ok": all(val.values())}


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
ss.setdefault("mat_sel", set())     # materiales marcados  (clave|UM)
ss.setdefault("sp_sel", set())      # SOLPED marcadas
ss.setdefault("excluidas", set())   # posiciones que el comprador dejó fuera (SOLPED/Pos)
ss.setdefault("sel_sig", None)
ss.setdefault("sel_ver", 0)
ss.setdefault("lin_ver", 0)
ss.setdefault("archivos", {})

PROV_DEFECTO = {"Moneda": "CLP", "Transporte": "T. Gil", "Flete total": 0.0, "Moneda flete": "CLP", "Plazo": 10,
                "Tipo plazo": "Días hábiles", "Tipo compra": "Nacional",
                "Lugar entrega": list(LUGARES_ENTREGA)[0], "Bodega tránsito": False}

# ---------------------------------------------------------------- tipo de cambio en línea
ss.setdefault("tc_override", {})     # valores escritos a mano, por moneda
ind = obtener_indicadores()
ss.tc = {k: float(ss.tc_override.get(k, ind[k]) or 0.0) for k in ["USD", "EUR", "UF"]}
ss.tc.update({"fecha": ind["fecha"], "hora": ind["hora"],
              "fuente_por_moneda": {k: ("ingresado a mano" if k in ss.tc_override else ind["fuente_por_moneda"].get(k, ""))
                                    for k in ["USD", "EUR", "UF"]}})
fuentes_usadas = list(dict.fromkeys(v for v in ss.tc["fuente_por_moneda"].values() if v))
ss.tc["fuente"] = " / ".join(fuentes_usadas) or "Sin conexión"
tc = ss.tc
TC_OK = min(tc["USD"], tc["EUR"], tc["UF"]) > 0


def faltan_tc(monedas) -> list:
    """Monedas que se necesitan para convertir y que no tienen valor."""
    return [NOMBRE_TC[m] for m in ["USD", "EUR", "UF"] if m in monedas and tc[m] <= 0]


def monedas_en_uso(C) -> set:
    if C is None:
        return set()
    usadas = set()
    for p in C["prov"].values():
        usadas |= {p.get("Moneda", "CLP"), p.get("Moneda flete", "CLP")}
    return usadas - {"CLP"}


def firma(*partes) -> str:
    return hashlib.md5("|".join(map(str, partes)).encode()).hexdigest()[:10]


def dinero(v, moneda="CLP") -> str:
    if v is None or pd.isna(v):
        return "—"
    if moneda == "CLP":
        return "$ " + f"{v:,.0f}".replace(",", ".")
    simbolo = {"USD": "US$", "EUR": "€", "UF": "UF"}.get(moneda, moneda)
    return f"{simbolo} " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def nombre_archivo(txt) -> str:
    return re.sub(r"[^A-Za-z0-9_-]", "_", str(txt))


def caja(clase, texto):
    st.markdown(f"<div class='{clase}'>{texto}</div>", unsafe_allow_html=True)


def tarjetas(datos):
    """Tarjetas de estadísticas: fondo rojo mate, cada número con su color."""
    html_t = "".join(f"<div><b class='c{(i % 6) + 1}'>{v}</b>{html.escape(str(e))}</div>"
                     for i, (v, e) in enumerate(datos))
    st.markdown(f"<div class='mini'>{html_t}</div>", unsafe_allow_html=True)


def ir_a(paso: int):
    """Cambia de paso y 'congela' las grillas para no perder lo escrito."""
    C = ss.cuadro
    if C is not None and C.get("of_last") is not None:
        C["of_src"], C["of_v"] = C["of_last"].copy(), C["of_v"] + 1
    ss.sel_sig = None
    ss.lin_ver += 1
    ss.paso = paso


def candidatas() -> pd.DataFrame:
    """Posiciones de los materiales y SOLPED marcados (antes de quitar las excluidas)."""
    req = ss.req
    if req is None or (not ss.mat_sel and not ss.sp_sel):
        return pd.DataFrame()
    clave = req["Clave"] + "|" + req["UM"]
    out = req[clave.isin(ss.mat_sel) | req["SOLPED"].isin(ss.sp_sel)].copy()
    out["id"] = out["SOLPED"].astype(str) + "/" + out["Pos"].astype(str)
    return out.sort_values(["SOLPED", "Pos"]).reset_index(drop=True)


def lineas_elegidas() -> pd.DataFrame:
    c = candidatas()
    return c if c.empty else c[~c["id"].isin(ss.excluidas)].reset_index(drop=True)


def valor_referencial(lineas: pd.DataFrame) -> float:
    if lineas.empty or "Precio ref." not in lineas:
        return 0.0
    monedas = lineas["Moneda ref."] if "Moneda ref." in lineas else ["CLP"] * len(lineas)
    return float(sum(a_clp(p * q, m, tc) for p, q, m in zip(lineas["Precio ref."].fillna(0), lineas["Cantidad"], monedas)))


def crear_cuadro(lineas: pd.DataFrame, agrupar: bool = True):
    lin, items = consolidar_lineas(lineas, agrupar)
    items = adjuntar_historico(items)
    vistos, sugeridos = set(), []
    for p in items["Últ. proveedor"].dropna():
        if str(p).strip() and norm(p).replace(" s a", "") not in vistos:
            vistos.add(norm(p).replace(" s a", ""))
            sugeridos.append(str(p).strip())
    sugeridos = sugeridos[:6]
    ref = (f"MAT-{items.iloc[0]['Material'] or 'SINCOD'}" if len(items) == 1 else
           (lin["SOLPED"].iloc[0] if lin["SOLPED"].nunique() == 1 else "AGRUPADO")) + f"-{date.today():%Y%m%d}"
    ss.cuadro = {"ref": ref, "firma": firma(sorted(lin["SOLPED"] + "/" + lin["Pos"].astype(str))),
                 "lineas": lin, "items": items, "prov": {}, "sugeridos": sugeridos,
                 "of_src": None, "of_v": 0, "of_last": None, "prov_names": None,
                 "decision": {}, "importados": {},
                 "aprobacion": {"ok": False, "nombre": "", "fecha": "", "firma": ""}}


def prov_df(C) -> pd.DataFrame:
    return pd.DataFrame([{"Proveedor": n, **{**PROV_DEFECTO, **v}} for n, v in C["prov"].items()],
                        columns=["Proveedor", *PROV_DEFECTO])


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


def sincronizar_ofertas(C):
    """Reconstruye la grilla de ofertas cuando cambia la lista de proveedores."""
    nombres = tuple(C["prov"])
    if C.get("prov_names") != nombres or C["of_src"] is None:
        of = armar_ofertas(C["items"], nombres, C["of_last"])
        C["of_src"], C["of_last"], C["of_v"], C["prov_names"] = of, of.copy(), C["of_v"] + 1, nombres


def incorporar_respuesta(C, datos, resp, nombre_arch):
    """Carga la planilla devuelta por un proveedor. Un segundo precio distinto queda como 'Precio mejorado'."""
    nombre = str(datos.get("Razón social") or "").strip() or nombre_arch.rsplit(".", 1)[0]
    d = {**PROV_DEFECTO, **C["prov"].get(nombre, {})}
    moneda = str(datos.get("Moneda de la oferta") or "").upper().strip()
    transporte = str(datos.get("Incoterm / forma de despacho") or "").strip()
    if moneda in MONEDAS:
        d["Moneda"] = d["Moneda flete"] = moneda
    if transporte in TRANSPORTES:
        d["Transporte"] = transporte
    if a_num(datos.get("Costo total de flete"), 0.0):
        d["Flete total"] = a_num(datos.get("Costo total de flete"), 0.0)
    if a_num(datos.get("Plazo de entrega general (días)"), 0):
        d["Plazo"] = int(a_num(datos.get("Plazo de entrega general (días)"), 0))
    C["prov"][nombre] = d
    of = armar_ofertas(C["items"], tuple(C["prov"]), C["of_last"])
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
        if a_num(r.get("Plazo entrega (días)"), 0):
            of.loc[m, "Plazo"] = a_num(r.get("Plazo entrega (días)"), 0)
        cumple = str(r.get("Cumple especificación (Sí/No)") or "").strip().lower()
        of.loc[m, "Cumple técnico"] = not cumple.startswith("n")
        obs = " · ".join(str(x) for x in [r.get("Marca / modelo ofrecido"), r.get("Observaciones")]
                         if x not in (None, "") and str(x) != "nan")
        if obs:
            of.loc[m, "Observación"] = obs
        n += 1
    C["of_src"], C["of_last"], C["of_v"], C["prov_names"] = of, of.copy(), C["of_v"] + 1, tuple(C["prov"])
    return nombre, n


def cuadro_a_json(C) -> str:
    dj = lambda df: None if df is None else df.to_json(orient="records", date_format="iso", force_ascii=False)
    return json.dumps({"ref": C["ref"], "firma": C["firma"], "lineas": dj(C["lineas"]), "items": dj(C["items"]),
                       "prov": C["prov"], "sugeridos": C.get("sugeridos", []), "ofertas": dj(C["of_last"]),
                       "decision": C["decision"], "aprobacion": C["aprobacion"]}, ensure_ascii=False, default=str)


def json_a_cuadro(texto: str):
    d = json.loads(texto)
    leer = lambda s: pd.read_json(io.StringIO(s), orient="records", dtype=False) if s else None
    items = leer(d["items"])
    items["Últ. compra"] = pd.to_datetime(items["Últ. compra"], errors="coerce").dt.date
    items["Material"] = items["Material"].fillna("").astype(str)
    prov = d.get("prov") or {}
    if isinstance(prov, list):                       # borradores de la versión anterior
        prov = {p.pop("Proveedor"): p for p in prov}
    of = leer(d["ofertas"])
    ss.cuadro = {"ref": d["ref"], "firma": d.get("firma", ""), "lineas": leer(d["lineas"]), "items": items,
                 "prov": prov, "sugeridos": d.get("sugeridos", []), "of_src": of, "of_last": of,
                 "of_v": 1, "prov_names": tuple(prov), "importados": {},
                 "decision": {int(k): v for k, v in d.get("decision", {}).items()},
                 "aprobacion": d.get("aprobacion", {"ok": False, "nombre": "", "fecha": "", "firma": ""})}


def preparar_items(C):
    """Columnas en pesos chilenos para comparar (última compra y precio SOLPED)."""
    items = C["items"]
    for c, v in [("Precio SOLPED", 0.0), ("Moneda SOLPED", "CLP")]:
        if c not in items:
            items[c] = v
    items["Últ. precio CLP"] = [a_clp(p, m, tc) if pd.notna(p) else np.nan
                                for p, m in zip(items["Últ. precio"], items["Últ. moneda"])]
    items["Últ. precio CLP"] = items["Últ. precio CLP"].where(items["Últ. precio CLP"] > 0)
    items["Precio SOLPED CLP"] = [a_clp(p, m, tc) if p else np.nan
                                  for p, m in zip(items["Precio SOLPED"].fillna(0), items["Moneda SOLPED"])]
    return items


def calcular(C):
    items = preparar_items(C)
    of = C["of_last"]
    if not C["prov"] or of is None or not of["Precio inicial"].fillna(0).gt(0).any() or faltan_tc(monedas_en_uso(C)):
        return None, None
    return calcular_comparativo(items, prov_df(C), of, tc)


def tabla_adjudicacion(C, det, res) -> pd.DataFrame:
    src = res.merge(C["items"][["Ítem", "Texto", "Cantidad", "UM", "SOLPEDs", "Últ. precio CLP"]], on="Ítem")
    elegido, tipo, comentario = [], [], []
    for _, r in src.iterrows():
        d = C["decision"].get(int(r["Ítem"]), {})
        p = d.get("prov", r["Sugerido"])
        elegido.append(p)
        tipo.append((r["Tipo sugerido"] or TIPOS_ADJUDICACION[0]) if p == r["Sugerido"]
                    else d.get("tipo", "Aprobado Técnicamente"))
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
    return msgs

# =============================================================================
# ENCABEZADO CORPORATIVO CON TIPO DE CAMBIO EN LÍNEA
# =============================================================================
def tasa_html(clase, nombre, valor):
    return f"<div class='tasa {clase}'><span>{nombre}</span><b>{dinero(valor) if valor else '—'}</b></div>"


detalle_fuente = ", ".join(f"{NOMBRE_TC[k]}: {v}" for k, v in tc["fuente_por_moneda"].items() if v) \
    if len(fuentes_usadas) > 1 else (fuentes_usadas[0] if fuentes_usadas else "sin conexión")
origen_tc = f"{detalle_fuente} · consultado a las {tc['hora']}"
st.markdown(
    "<div class='banner'>"
    "<div class='izq'><div class='marca'>Enaex · Gerencia de Abastecimiento</div>"
    "<div class='titulo'>Cuadro comparativo de ofertas</div>"
    f"<div class='fuente'>{html.escape(origen_tc)}</div></div>"
    f"<div class='tasas'>{tasa_html('usd', 'Dólar', tc['USD'])}{tasa_html('eur', 'Euro', tc['EUR'])}"
    f"{tasa_html('uf', 'UF', tc['UF'])}</div></div>",
    unsafe_allow_html=True)


def refrescar_tc():
    obtener_indicadores.clear()
    ss.tc_override = {}


def fijar_tc(moneda, valor):
    if valor and valor > 0:
        ss.tc_override[moneda] = float(valor)
        guardar_cache_tc({moneda: (valor, date.today().isoformat())})


_, a1, a2 = st.columns([5.6, 1.2, 1.2], gap="small")
a2.button("Actualizar", icon=":material/refresh:", on_click=refrescar_tc, use_container_width=True,
          help="Vuelve a consultar los valores del día y descarta lo escrito a mano")
with a1.popover("Ajustar", icon=":material/edit:", use_container_width=True):
    st.caption("Escribe el valor si necesitas otro (por ejemplo, el de la fecha de la oferta).")
    for k in ["USD", "EUR", "UF"]:
        v = st.number_input(NOMBRE_TC[k], value=float(tc[k]), min_value=0.0, step=1.0, format="%.2f",
                            key=f"tc_{k}_{firma(tc[k])}")
        if abs(v - tc[k]) > 1e-9:
            fijar_tc(k, v)
            st.rerun()

sin_valor = [k for k in ["USD", "EUR", "UF"] if tc[k] <= 0]
if sin_valor:
    nombres_sv = [NOMBRE_TC[k] for k in sin_valor]
    nombres_sv = nombres_sv[0] if len(nombres_sv) == 1 else ", ".join(nombres_sv[:-1]) + " y " + nombres_sv[-1]
    with st.container(border=True):
        st.markdown(f"<span class='alerta-tc'>{ico('alerta')}</span>No pudimos traer **{nombres_sv}** desde internet "
                    "(tu red puede estar bloqueando las páginas del Banco Central y del SII). Escribe el valor de hoy "
                    "y quedará guardado para la próxima vez. "
                    "Solo es obligatorio si alguna oferta o transporte está en esa moneda.", unsafe_allow_html=True)
        cols_sv = st.columns(len(sin_valor) + 1)
        for col, k in zip(cols_sv, sin_valor):
            v = col.number_input(f"{NOMBRE_TC[k]} de hoy ($)", min_value=0.0, step=1.0, format="%.2f",
                                 value=None, placeholder="Ej: 39.612,40" if k == "UF" else "Ej: 945,30",
                                 key=f"tc_falta_{k}")
            if v:
                fijar_tc(k, v)
                st.rerun()
        cols_sv[-1].markdown("<div class='sub' style='padding-top:1.9rem'>Valor oficial en sii.cl → "
                             "Valores y fechas</div>", unsafe_allow_html=True)

barra = "".join(
    f"<div class='paso {'actual' if i == ss.paso else 'hecho' if i < ss.paso else ''}'>"
    f"<span class='ico-paso'>{ico('check' if i < ss.paso else ic, 1.35)}</span>{i}. {txt}</div>"
    for i, (ic, txt) in enumerate(PASOS, start=1))
st.markdown(f"<div class='pasos'>{barra}</div>", unsafe_allow_html=True)


def navegacion(puede_seguir=True, motivo="", al_seguir=None):
    st.write("")
    st.divider()
    c1, c2, c3 = st.columns([1, 2, 1.4])
    if ss.paso > 1:
        c1.button("← Volver", on_click=ir_a, args=(ss.paso - 1,), use_container_width=True)
    if ss.paso < len(PASOS):
        c3.button(f"Siguiente: {PASOS[ss.paso][1]} →", type="primary", disabled=not puede_seguir,
                  on_click=al_seguir or ir_a, args=() if al_seguir else (ss.paso + 1,), use_container_width=True)
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
        st.markdown(f"#### {num(1)}Planilla de SOLPED", unsafe_allow_html=True)
        st.caption("Obligatoria · Excel o CSV")
        f_req = st.file_uploader("Planilla de SOLPED", type=["xlsx", "xlsm", "xls", "csv"],
                                 label_visibility="collapsed", key="up_req")
    with c2:
        st.markdown(f"#### {num(2)}Histórico de compras", unsafe_allow_html=True)
        st.caption("Opcional · Cuadro MRO (hoja Ultima COMPRA), ME2M o CSV")
        f_hist = st.file_uploader("Histórico de compras", type=["xlsx", "xlsm", "xls", "csv"],
                                  label_visibility="collapsed", key="up_hist")

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
        tarjetas([(ss.req["SOLPED"].nunique(), "SOLPED"), (len(ss.req), "posiciones"),
                  (ss.req["Clave"].nunique(), "materiales distintos"),
                  (f"{len(ss.hist[0]):,}".replace(",", ".") if ss.hist is not None else "—",
                   "materiales con historial")])
    elif ss.get("req_raw") is not None:
        caja("ojo", "No pudimos reconocer las columnas de la planilla. Abre <b>¿No se ve bien?</b> aquí abajo "
                    "e indica cuál es la columna de SOLPED, descripción y cantidad.")

    if ss.get("req_raw") is not None or ss.get("hist_raw") is not None:
        with st.expander("¿No se ve bien? Ajustar cómo leímos el archivo"):
            for tipo, etiqueta, sins in [("req", "Planilla de SOLPED", SIN_REQ), ("hist", "Histórico", SIN_HIST)]:
                if ss.get(f"{tipo}_raw") is None:
                    continue
                st.markdown(f"**{etiqueta}**")
                hojas = ss[f"{tipo}_hojas"]
                st.selectbox("Hoja del Excel", hojas, index=hojas.index(ss[f"{tipo}_hoja_usada"]), key=f"hoja_{tipo}")
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
    with o1.expander("No tengo planilla: escribir los materiales a mano", icon=":material/edit_note:"):
        sp_man = st.text_input("N° de SOLPED", key="sp_manual")
        plantilla = pd.DataFrame({"Material": [""], "Texto": [""], "Cantidad": [1.0], "UM": ["C/U"],
                                  "Precio ref.": [0.0], "Moneda ref.": ["CLP"], "Centro": [""]})
        pos_man = st.data_editor(plantilla, num_rows="dynamic", use_container_width=True, key="pos_manual",
                                 column_config={"Material": "Código SAP (opcional)", "Texto": "Descripción",
                                                "Precio ref.": st.column_config.NumberColumn("Precio SOLPED"),
                                                "Moneda ref.": st.column_config.SelectboxColumn("Moneda",
                                                                                                options=MONEDAS)})
        if st.button("Usar estos materiales", disabled=not sp_man.strip()):
            pm = pos_man.dropna(subset=["Texto"])
            pm = pm[pm["Texto"].astype(str).str.strip() != ""].copy()
            pm["SOLPED"], pm["Pos"] = sp_man.strip(), np.arange(1, len(pm) + 1) * 10
            pm["Material"] = pm["Material"].apply(clave_material)
            for c, v in [("F. solicitud", pd.NaT), ("Días", np.nan), ("Urgencia", ""), ("G. compras", ""),
                         ("Asignado a", ""), ("Contrato marco", "")]:
                pm[c] = v
            pm["Clave"] = [clave_item(m, t) for m, t in zip(pm["Material"], pm["Texto"])]
            ss.req, ss.sp_sel, ss.excluidas = pm, {sp_man.strip()}, set()
            ir_a(2)
            st.rerun()
    with o2.expander("Retomar un trabajo guardado", icon=":material/history:"):
        f_json = st.file_uploader("Archivo .json guardado en el paso 5", type=["json"], key="up_json")
        if f_json is not None:
            if ss.archivos.get("json") != firma(f_json.name, f_json.size):
                json_a_cuadro(f_json.getvalue().decode("utf-8"))
                ss.archivos["json"] = firma(f_json.name, f_json.size)
            st.success(f"Trabajo «{ss.cuadro['ref']}» recuperado.")
            st.button("Continuar donde quedé →", type="primary", on_click=ir_a, args=(3,))

    navegacion(ss.req is not None, "Primero sube la planilla de SOLPED.")

# =============================================================================
# PASO 2 · ELIGE QUÉ COTIZAR
# =============================================================================
elif ss.paso == 2:
    req = ss.req
    caja("ayuda", f"{num(1)}Marca {ico('check', 1)} los materiales o SOLPED que quieres cotizar. "
                  f"<br>{num(2)}Abajo aparecen todas sus posiciones: desmarca las que <b>no</b> quieres incluir "
                  "(se quedan pendientes para otra cotización). Lo que piden varias SOLPED se junta en una sola línea.")
    c1, c2, c3 = st.columns([3, 1.6, 1.4])
    busca = c1.text_input("Buscar", placeholder="Escribe un material, código o N° de SOLPED…",
                          label_visibility="collapsed")
    vista = c2.segmented_control("Ver", ["Por SOLPED", "Por material"], default="Por SOLPED",
                                 label_visibility="collapsed") or "Por SOLPED"
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
                             "Días": ("Días", "max")})
                     .reset_index())
            if solo_rep:
                tabla = tabla[tabla["SOLPED"] > 1]
            tabla = adjuntar_historico(tabla).sort_values(["SOLPED", "Días"], ascending=False).reset_index(drop=True)
            tabla["id"] = tabla["Clave"] + "|" + tabla["UM"]
            tabla.insert(0, "Cotizar", tabla["id"].isin(ss.mat_sel))
        else:
            tabla = (base.groupby("SOLPED")
                     .agg(Posiciones=("Pos", "size"), Días=("Días", "max"), Centro=("Centro", "first"),
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
        cfg = {"Cotizar": st.column_config.CheckboxColumn("Cotizar", width="small"),
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
        cfg = {"Cotizar": st.column_config.CheckboxColumn("Cotizar", width="small"),
               "Posiciones": st.column_config.NumberColumn("Posiciones", format="%d"),
               "Días": st.column_config.NumberColumn("Antigüedad", format="%d días"),
               "Contenido": st.column_config.TextColumn("Qué pide", width="large")}
        orden = ["Cotizar", "SOLPED", "Contenido", "Posiciones", "Días", "Centro"]
    st.markdown(f"##### {num(1)}¿Qué quieres cotizar?", unsafe_allow_html=True)
    editado = st.data_editor(tabla, key=f"sel_{ss.sel_ver}", hide_index=True, use_container_width=True,
                             height=300, column_order=orden, column_config=cfg,
                             disabled=[c for c in tabla.columns if c != "Cotizar"])
    destino = ss.mat_sel if vista == "Por material" else ss.sp_sel
    for i, marcado in zip(editado["id"], editado["Cotizar"]):
        (destino.add if marcado else destino.discard)(i)

    cand = candidatas()
    if cand.empty:
        caja("sub", f"Marca la casilla {ico('check', 1)} de una SOLPED o de un material para ver sus posiciones.")
    else:
        st.markdown(f"##### {num(2)}Elige qué posiciones entran al cuadro", unsafe_allow_html=True)
        b1, b2, _ = st.columns([1, 1, 3])

        def marcar_todas(valor: bool):
            ids = set(candidatas()["id"])
            ss.excluidas = (ss.excluidas - ids) if valor else (ss.excluidas | ids)
            ss.lin_ver += 1

        b1.button("Marcar todas", icon=":material/check_box:", on_click=marcar_todas, args=(True,), use_container_width=True)
        b2.button("Desmarcar todas", icon=":material/check_box_outline_blank:", on_click=marcar_todas, args=(False,), use_container_width=True)
        vista_lin = cand[["id", "SOLPED", "Pos", "Material", "Texto", "Cantidad", "UM", "Precio ref.",
                          "Moneda ref.", "Centro"]].copy()
        vista_lin.insert(0, "Incluir", ~vista_lin["id"].isin(ss.excluidas))
        vista_lin["Ref. en pesos"] = [a_clp(p * q, m, tc) if p else np.nan for p, q, m in
                                      zip(vista_lin["Precio ref."], vista_lin["Cantidad"], vista_lin["Moneda ref."])]
        lin_sig = firma(sorted(cand["id"]), ss.lin_ver)
        incl = st.data_editor(
            vista_lin, key=f"lin_{lin_sig}", hide_index=True, use_container_width=True,
            height=min(80 + 35 * len(vista_lin), 420),
            column_order=["Incluir", "SOLPED", "Pos", "Texto", "Material", "Cantidad", "UM", "Precio ref.",
                          "Moneda ref.", "Ref. en pesos", "Centro"],
            disabled=[c for c in vista_lin.columns if c != "Incluir"],
            column_config={"Incluir": st.column_config.CheckboxColumn("Entra", width="small"),
                           "Texto": st.column_config.TextColumn("Material", width="large"),
                           "Material": "Código",
                           "Cantidad": st.column_config.NumberColumn(format="%g"),
                           "Precio ref.": st.column_config.NumberColumn("Precio SOLPED", format="localized"),
                           "Moneda ref.": "Moneda",
                           "Ref. en pesos": st.column_config.NumberColumn("Total en $ CLP", format="localized")})
        for i, entra in zip(incl["id"], incl["Incluir"]):
            (ss.excluidas.discard if entra else ss.excluidas.add)(i)

    lin = lineas_elegidas()
    if not lin.empty:
        _, items_prev = consolidar_lineas(lin, True)
        ref_clp = valor_referencial(lin)
        tarjetas([(len(lin), "posiciones elegidas"), (lin["SOLPED"].nunique(), "SOLPED"),
                  (len(items_prev), "ítems a cotizar"), (dinero(ref_clp) if ref_clp else "—", "valor referencial SOLPED")])
        if ref_clp > UMBRAL_TC_CLP:
            caja("ok", f"Supera $1.000.000: todo se convertirá con el tipo de cambio del día "
                       f"(dólar {dinero(tc['USD'])}, euro {dinero(tc['EUR'])}, UF {dinero(tc['UF'])}).")

    def preparar_y_seguir():
        lineas = lineas_elegidas()
        f = firma(sorted(lineas["SOLPED"] + "/" + lineas["Pos"].astype(str)))
        if ss.cuadro is None or ss.cuadro["firma"] != f:
            crear_cuadro(lineas, True)
        ir_a(3)

    aviso = ""
    if ss.cuadro is not None and not lin.empty and ss.cuadro["firma"] != firma(sorted(lin["SOLPED"] + "/" + lin["Pos"].astype(str))):
        aviso = "Cambiaste la selección: se armará un cuadro nuevo."
    navegacion(not lin.empty, aviso or "Marca al menos una posición.", al_seguir=preparar_y_seguir)
    if aviso and not lin.empty:
        st.caption(aviso)

# =============================================================================
# PASO 3 · INGRESA LAS OFERTAS
# =============================================================================
elif ss.paso == 3:
    C = ss.cuadro
    items, lineas = preparar_items(C), C["lineas"]
    sincronizar_ofertas(C)
    caja("ayuda", "Registra lo que ofreció cada proveedor: <b>precio, moneda, transporte y plazo</b>. "
                  "Todo se convierte a pesos chilenos con el tipo de cambio de hoy, así puedes comparar euros, "
                  "dólares y pesos en la misma escala.")
    ref_clp = float((items["Precio SOLPED CLP"].fillna(0) * items["Cantidad"]).sum())
    tarjetas([(len(items), "ítems"), (lineas["SOLPED"].nunique(), "SOLPED"), (len(C["prov"]), "proveedores"),
              (dinero(ref_clp) if ref_clp else "—", "valor referencial SOLPED")])

    clave_sel = f"psel_{C['ref']}"
    opciones = list(C["prov"]) + [AGREGAR]
    if ss.get(clave_sel) not in opciones:
        ss[clave_sel] = opciones[0]

    def congelar():
        if C.get("of_last") is not None:
            C["of_src"], C["of_v"] = C["of_last"].copy(), C["of_v"] + 1

    def agregar_proveedor(nombre=None):
        nombre = (nombre or ss.get(f"nuevo_{C['ref']}", "")).strip()
        if nombre and nombre not in C["prov"]:
            C["prov"][nombre] = dict(PROV_DEFECTO)
            congelar()
            ss[clave_sel] = nombre
            ss[f"nuevo_{C['ref']}"] = ""

    def quitar_proveedor(nombre):
        C["prov"].pop(nombre, None)
        congelar()
        ss[clave_sel] = (list(C["prov"]) or [AGREGAR])[0]

    st.segmented_control("Proveedor", opciones, key=clave_sel, on_change=congelar, label_visibility="collapsed")
    actual = ss.get(clave_sel) or opciones[0]

    if actual == AGREGAR:
        with st.container(border=True):
            st.markdown(f"#### <span class='titulo-ico'>{ico('mas')}</span>Nuevo proveedor", unsafe_allow_html=True)
            n1, n2 = st.columns([3, 1])
            n1.text_input("Nombre del proveedor", key=f"nuevo_{C['ref']}", placeholder="Ej: Sonepar Chile S.A.")
            n2.write("")
            n2.button("Agregar", type="primary", on_click=agregar_proveedor, use_container_width=True)
            sugeridos = [s for s in C.get("sugeridos", []) if s not in C["prov"]]
            if sugeridos:
                st.caption("Proveedores que ya vendieron estos materiales (un clic para agregar):")
                cols_s = st.columns(min(3, len(sugeridos)))
                for i, s in enumerate(sugeridos):
                    cols_s[i % len(cols_s)].button(f"＋ {s}", key=f"sug_{C['ref']}_{i}", on_click=agregar_proveedor,
                                                   args=(s,), use_container_width=True)
    else:
        nombre = actual
        P = {**PROV_DEFECTO, **C["prov"][nombre]}
        k = lambda campo: f"pv_{C['ref']}_{firma(nombre)}_{campo}"
        with st.container(border=True):
            h1, h2 = st.columns([4, 1])
            h1.markdown(f"#### <span class='titulo-ico'>{ico('empresa')}</span>{html.escape(nombre)}",
                        unsafe_allow_html=True)
            h2.button("Quitar", icon=":material/delete:", key=k("quitar"), on_click=quitar_proveedor, args=(nombre,), use_container_width=True)

            st.markdown("**Oferta y transporte**")
            f1, f2, f3, f4 = st.columns(4)
            P["Moneda"] = f1.selectbox("Moneda de los precios", MONEDAS, index=MONEDAS.index(P["Moneda"]),
                                       key=k("mon"))
            P["Transporte"] = f2.selectbox("Transporte", TRANSPORTES, index=TRANSPORTES.index(P["Transporte"])
                                           if P["Transporte"] in TRANSPORTES else 0, key=k("trans"))
            P["Flete total"] = f3.number_input("Costo del transporte", value=float(P["Flete total"]), min_value=0.0,
                                               step=1000.0 if P["Moneda flete"] == "CLP" else 10.0, key=k("flete"))
            P["Moneda flete"] = f4.selectbox("Moneda del transporte", MONEDAS, index=MONEDAS.index(P["Moneda flete"]),
                                             key=k("monfl"))
            st.markdown("**Envío y entrega**")
            g1, g2, g3, g4 = st.columns(4)
            P["Plazo"] = g1.number_input("Plazo de despacho", value=int(a_num(P["Plazo"], 0)), min_value=0, step=1,
                                         key=k("plazo"))
            P["Tipo plazo"] = g2.selectbox("Unidad", TIPOS_PLAZO, index=TIPOS_PLAZO.index(P["Tipo plazo"]),
                                           key=k("tplazo"))
            P["Tipo compra"] = g3.selectbox("Origen", TIPOS_COMPRA, index=TIPOS_COMPRA.index(P["Tipo compra"]),
                                            key=k("origen"), help="Importación suma 10 días de internación")
            P["Lugar entrega"] = g4.selectbox("Dónde entrega", list(LUGARES_ENTREGA),
                                              index=list(LUGARES_ENTREGA).index(P["Lugar entrega"]), key=k("lugar"))
            P["Bodega tránsito"] = st.checkbox("Pasa por bodega de tránsito (Isaver / Bello, +7 días)",
                                               value=bool(P["Bodega tránsito"]), key=k("bodega"))
            C["prov"][nombre] = P
            f_desp, f_planta, _ = calcular_fecha_planta(P["Plazo"], P["Tipo plazo"], P["Tipo compra"],
                                                        P["Lugar entrega"], P["Bodega tránsito"])

            st.markdown(f"**Precios unitarios** (en {P['Moneda']}, sin IVA)")
            src = C["of_src"][C["of_src"]["Proveedor"] == nombre].merge(
                items[["Ítem", "Material", "Precio SOLPED", "Moneda SOLPED", "Precio SOLPED CLP"]], on="Ítem", how="left")
            src["Ref. SOLPED"] = [f"{dinero(p, m)}" + (f"  (≈ {dinero(pc)})" if m != "CLP" and pd.notna(pc) else "")
                                  if p else "—" for p, m, pc in zip(src["Precio SOLPED"], src["Moneda SOLPED"],
                                                                     src["Precio SOLPED CLP"])]
            ed = st.data_editor(
                src.reset_index(drop=True), key=f"of_{C['ref']}_{firma(nombre)}_{C['of_v']}", hide_index=True,
                use_container_width=True, height=min(80 + 35 * len(src), 420),
                column_order=["Ítem", "Texto", "Material", "Cantidad", "UM", "Ref. SOLPED", "Precio inicial",
                              "Precio mejorado", "Plazo", "Cumple técnico", "Observación"],
                disabled=["Ítem", "Texto", "Material", "Cantidad", "UM", "Ref. SOLPED"],
                column_config={"Texto": st.column_config.TextColumn("Material", width="large"),
                               "Material": "Código",
                               "Cantidad": st.column_config.NumberColumn(format="%g"),
                               "Ref. SOLPED": "Precio SOLPED",
                               "Precio inicial": st.column_config.NumberColumn(f"Precio ({P['Moneda']})", min_value=0.0,
                                                                               format="localized"),
                               "Precio mejorado": st.column_config.NumberColumn(
                                   "Precio negociado", min_value=0.0, format="localized",
                                   help="Opcional: si el proveedor mejoró su precio"),
                               "Plazo": st.column_config.NumberColumn("Plazo distinto", min_value=0, step=1,
                                                                      help="Solo si este ítem tiene otro plazo"),
                               "Cumple técnico": st.column_config.CheckboxColumn("¿Cumple?"),
                               "Observación": st.column_config.TextColumn("Notas")})
            mascara = C["of_last"]["Proveedor"] == nombre
            editados = ed.set_index("Ítem")
            for col in ["Precio inicial", "Precio mejorado", "Plazo", "Cumple técnico", "Observación"]:
                C["of_last"].loc[mascara, col] = C["of_last"].loc[mascara, "Ítem"].map(editados[col]).values

            precio = ed["Precio mejorado"].where(ed["Precio mejorado"].fillna(0) > 0, ed["Precio inicial"]).fillna(0)
            mat = float((precio * ed["Cantidad"]).sum())
            mat_clp, fl_clp = a_clp(mat, P["Moneda"], tc), a_clp(P["Flete total"], P["Moneda flete"], tc)
            tarjetas([(dinero(mat, P["Moneda"]) if P["Moneda"] != "CLP" else dinero(mat), "materiales"),
                      (dinero(mat_clp), "materiales en pesos"), (dinero(fl_clp), "transporte en pesos"),
                      (dinero(mat_clp + fl_clp), "total puesto en planta"),
                      (f"{f_desp:%d/%m/%Y}", "despacho del proveedor"),
                      (f"{f_planta:%d/%m/%Y}", "llega a planta")])

    with st.expander("Opcional: enviar planilla de cotización por correo e importar respuestas", icon=":material/mail:"):
        caja("sub", "Si prefieres pedir la oferta por escrito, descarga esta planilla (no muestra precios internos). "
                    "Cuando el proveedor la devuelva, súbela y sus precios se cargan solos.")
        r1, r2 = st.columns(2)
        with r1:
            comprador = st.text_input("Tu nombre", value=ss.get("mi_nombre", unir_unicos(lineas.get("Asignado a", []))))
            correo = st.text_input("Tu correo", value=ss.get("mi_correo", ""))
            limite = st.date_input("Responder antes del", value=date.today() + timedelta(days=3), format="DD/MM/YYYY")
            ss.mi_nombre, ss.mi_correo = comprador, correo
            st.download_button("Descargar planilla para proveedor", icon=":material/download:", use_container_width=True,
                               data=generar_rfq(C["ref"], items, lineas,
                                                {"comprador": comprador, "correo": correo, "fecha_limite": limite,
                                                 "entrega": unir_unicos(items["Centro"])}),
                               file_name=f"Solicitud_cotizacion_{nombre_archivo(C['ref'])}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with r2:
            respuestas = st.file_uploader("Planillas respondidas (.xlsx)", type=["xlsx"], accept_multiple_files=True,
                                          key=f"resp_{C['ref']}")
            for f in respuestas or []:
                clave_f = firma(f.name, f.size)
                if clave_f in C["importados"]:
                    continue
                try:
                    ref_resp, datos, resp = leer_respuesta_rfq(f.getvalue())
                    if ref_resp and ref_resp != C["ref"]:
                        C["importados"][clave_f] = ("error", f"{f.name}: es de otra cotización ({ref_resp}).")
                    else:
                        n_prov, n = incorporar_respuesta(C, datos, resp, f.name)
                        C["importados"][clave_f] = ("ok", f"{n_prov}: {n} precio(s) cargados.")
                        st.rerun()
                except Exception as e:
                    C["importados"][clave_f] = ("error", f"{f.name}: no se pudo leer ({e}).")
            for estado, msg in C["importados"].values():
                caja("ok" if estado == "ok" else "ojo", ico("check" if estado == "ok" else "alerta") + html.escape(msg))

    hay_precios = bool(C["prov"]) and C.get("of_last") is not None and C["of_last"]["Precio inicial"].fillna(0).gt(0).any()
    faltan = faltan_tc(monedas_en_uso(C))
    navegacion(hay_precios and not faltan,
               f"Falta el valor de {' y '.join(faltan)} (cuadro de arriba): hay ofertas en esa moneda." if faltan
               else "Ingresa al menos un precio.")

# =============================================================================
# PASO 4 · COMPARA Y ELIGE
# =============================================================================
elif ss.paso == 4:
    C = ss.cuadro
    det, res = calcular(C)
    if det is None:
        faltan = faltan_tc(monedas_en_uso(C))
        caja("ojo", f"Falta el valor de {' y '.join(faltan)} (arriba)." if faltan else "Faltan precios (paso 3).")
        navegacion(False)
        st.stop()
    items = C["items"]
    adj_prev = tabla_adjudicacion(C, det, res)
    tarjetas([(dinero(res["Mejor total CLP"].sum()), "mejor combinación posible"),
              (len(C["prov"]), "proveedores comparados"),
              (int((res["Ofertas válidas"] > 1).sum()), "ítems con 2+ ofertas"),
              (dinero(adj_prev["Ahorro vs histórico CLP"].sum()), "ahorro vs última compra")])

    st.markdown("##### Vista rápida: total puesto en planta por proveedor (en pesos)")
    matriz = det.pivot_table(index="Ítem", columns="Proveedor", values="Total CLP", aggfunc="first")
    validas = det.pivot_table(index="Ítem", columns="Proveedor", values="Válida", aggfunc="first")
    matriz = matriz.where(validas.fillna(False).astype(bool))
    totales = matriz.sum(min_count=1)
    matriz.index = [f"{i}. {str(t)[:45]}" for i, t in zip(matriz.index, matriz.index.map(items.set_index("Ítem")["Texto"]))]
    matriz.loc["TOTAL si le compro todo"] = totales

    def resaltar(fila):
        v = pd.to_numeric(fila, errors="coerce")
        if v.notna().sum() == 0:
            return [""] * len(fila)
        return ["background-color:#D1FAE5;color:#065F46;font-weight:700" if x == v.min() else "" for x in v]

    st.dataframe(matriz.style.apply(resaltar, axis=1).format(lambda x: dinero(x) if pd.notna(x) else "—"),
                 use_container_width=True)
    st.caption("Verde = más conveniente en esa fila. «—» = no cotizó o no cumple. Incluye el transporte repartido.")

    st.markdown("##### Decide por ítem")
    for _, it in items.iterrows():
        n_item = int(it["Ítem"])
        r = res[res["Ítem"] == n_item].iloc[0]
        ofertas = det[det["Ítem"] == n_item].sort_values(["Válida", "Total CLP"], ascending=[False, True])
        validas_it = ofertas[ofertas["Válida"]]
        maximo = validas_it["Total CLP"].max() if not validas_it.empty else 1
        with st.container(border=True):
            st.markdown(f"**{n_item}. {html.escape(str(it['Texto']))}**")
            ult = (f"Última compra: <b>{dinero(it['Últ. precio'], it['Últ. moneda'] or 'CLP')}</b> c/u a "
                   f"{html.escape(str(it['Últ. proveedor']))} ({it['Últ. compra']:%d/%m/%Y})"
                   if it["Con historial"] and pd.notna(it["Últ. precio"]) and pd.notna(it["Últ. compra"])
                   else "<span class='etiqueta roja'>Sin compras anteriores · requiere visto bueno técnico</span>")
            ref_sp = (f" · Precio SOLPED: <b>{dinero(it['Precio SOLPED'], it['Moneda SOLPED'])}</b> c/u"
                      + (f" (≈ {dinero(it['Precio SOLPED CLP'])})" if it["Moneda SOLPED"] != "CLP" else "")
                      if it.get("Precio SOLPED") else "")
            st.markdown(f"<div class='sub'>{it['Cantidad']:g} {html.escape(str(it['UM']))} · "
                        f"{int(it['N° SOLPED'])} SOLPED · {ult}{ref_sp}</div>", unsafe_allow_html=True)
            filas = []
            for _, o in ofertas.iterrows():
                nom = html.escape(str(o["Proveedor"]))
                if not o["Cotiza"]:
                    filas.append(f"<div class='fila-oferta'><div>{nom}</div><div class='sub'>No cotizó</div><div></div><div></div></div>")
                    continue
                gana = o["Proveedor"] == r["Sugerido"]
                ancho = max(6, o["Total CLP"] / maximo * 100) if o["Válida"] else 0
                badge = ("<span class='etiqueta'>Más conveniente</span>" if gana else
                         "" if o["Válida"] else "<span class='etiqueta roja'>No cumple</span>")
                var = o["Var. vs últ. compra"]
                var_txt = ("" if pd.isna(var) else
                           f"<span class='etiqueta {'roja' if var > UMBRAL_ALZA else 'gris' if var > 0 else ''}'>"
                           f"{abs(var):.0%} {'más caro' if var > 0 else 'más barato'} que la última vez</span>")
                orig = (f"<br><span class='sub'>{dinero(o['Precio mejorado'], o['Moneda'])} c/u</span>"
                        if o["Moneda"] != "CLP" else "")
                filas.append(
                    f"<div class='fila-oferta'><div><b>{nom}</b>{badge}{orig}</div>"
                    f"<div><div class='barra {'gana' if gana else ''}' style='width:{ancho:.0f}%'></div>{var_txt}</div>"
                    f"<div style='text-align:right'><b>{dinero(o['Total CLP'])}</b><br>"
                    f"<span class='sub'>{dinero(o['Unit. puesto CLP'])} c/u</span></div>"
                    f"<div class='sub'>Llega {o['Entrega planta']:%d/%m/%Y}</div></div>")
            st.markdown("".join(filas), unsafe_allow_html=True)

            opciones = list(validas_it["Proveedor"]) + [SIN_ADJ]
            previo = C["decision"].get(n_item, {}).get("prov", r["Sugerido"])
            elegido = st.radio("¿A quién le compras?", opciones, horizontal=True, key=f"dec_{C['ref']}_{n_item}",
                               index=opciones.index(previo) if previo in opciones else len(opciones) - 1)
            dec = {"prov": elegido, "comentario": ""}
            if elegido not in (r["Sugerido"], SIN_ADJ):
                e1, e2 = st.columns([1, 2])
                motivos = ["Aprobado Técnicamente", "Asignación Directa"]
                dec["tipo"] = e1.selectbox("Motivo", motivos, key=f"tipo_{C['ref']}_{n_item}",
                                           index=motivos.index(C["decision"].get(n_item, {}).get("tipo", motivos[0])))
                dec["comentario"] = e2.text_input("¿Por qué eliges este proveedor? (obligatorio)",
                                                  value=C["decision"].get(n_item, {}).get("comentario", ""),
                                                  key=f"com_{C['ref']}_{n_item}")
            C["decision"][n_item] = dec

    adj = tabla_adjudicacion(C, det, res)
    tarjetas([(dinero(adj["Monto adjudicado CLP"].sum()), "total de tu elección"),
              (dinero(adj["Ahorro spot CLP"].sum()), "ahorro por negociación"),
              (dinero(adj["Ahorro vs histórico CLP"].sum()), "ahorro vs última compra")])
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
    items, lineas, prov = C["items"], C["lineas"], prov_df(C)
    total = adj["Monto adjudicado CLP"].sum()
    tarjetas([(dinero(total), "total a comprar"),
              (f"{int((adj['Proveedor adjudicado'] != SIN_ADJ).sum())} de {len(adj)}", "ítems con proveedor"),
              (dinero(adj["Ahorro spot CLP"].sum()), "ahorro por negociación"),
              (dinero(adj["Ahorro vs histórico CLP"].sum()), "ahorro vs última compra")])
    if total > UMBRAL_TC_CLP:
        usados = ", ".join(f"{NOMBRE_TC[k].lower()} {dinero(tc[k])}" for k in ["USD", "EUR", "UF"] if tc[k] > 0)
        caja("ok", f"Compra sobre $1.000.000: convertida con el tipo de cambio del día ({usados}; "
                   f"fuente: {html.escape(tc['fuente'])}).")

    st.markdown("##### Resumen de la compra")
    st.dataframe(adj[["Ítem", "Texto", "Cantidad", "UM", "Proveedor adjudicado", "Monto adjudicado CLP",
                      "Entrega planta", "Tipo adjudicación"]], hide_index=True, use_container_width=True,
                 column_config={"Texto": st.column_config.TextColumn("Material", width="large"),
                                "Proveedor adjudicado": "Proveedor",
                                "Monto adjudicado CLP": st.column_config.NumberColumn("Total ($)", format="localized"),
                                "Entrega planta": st.column_config.DateColumn("Llega", format="DD/MM/YYYY")})
    with st.expander("Ver cómo se reparte por SOLPED (para emitir las OC)"):
        st.dataframe(detalle_por_solped(lineas, adj), hide_index=True, use_container_width=True,
                     column_config={"Precio unit. (moneda oferta)": st.column_config.NumberColumn(format="localized"),
                                    "Unit. puesto CLP": st.column_config.NumberColumn("Unitario puesto ($)", format="localized"),
                                    "Monto línea CLP": st.column_config.NumberColumn("Monto línea ($)", format="localized")})

    st.markdown("##### Aprobación")
    faltan = pendientes(adj)
    huella = firma(adj[["Ítem", "Proveedor adjudicado", "Monto adjudicado CLP", "Comentario"]].to_json(),
                   tc["USD"], tc["EUR"], tc["UF"])
    ap = C["aprobacion"]
    aprobado = ap.get("ok") and ap.get("firma") == huella
    if faltan:
        caja("ojo", ico("alerta") + "<b>Antes de aprobar, revisa:</b><br>" + "<br>".join("• " + html.escape(m) for m in faltan))
    elif aprobado:
        caja("ok", f"{ico('aprobado')}Aprobado por <b>{html.escape(ap['nombre'])}</b> el {ap['fecha']}.")
    else:
        a1, a2 = st.columns([2, 1])
        quien = a1.text_input("Tu nombre", value=ss.get("mi_nombre", ""), placeholder="Quién aprueba")
        if a2.button("Aprobar compra", icon=":material/task_alt:", type="primary", disabled=not quien.strip(), use_container_width=True):
            C["aprobacion"] = {"ok": True, "nombre": quien.strip(),
                               "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"), "firma": huella}
            st.rerun()
        st.caption("Al aprobar, los reportes dejan de decir BORRADOR. Si cambias algo después, habrá que aprobar de nuevo.")

    st.markdown("##### Descargas")
    ap_rep = C["aprobacion"] if aprobado else {"ok": False, "nombre": "", "fecha": ""}
    base_nombre = f"{nombre_archivo(C['ref'])}_{date.today():%Y%m%d}{'' if aprobado else '_BORRADOR'}"
    d1, d2, d3 = st.columns(3)
    d1.download_button("Cuadro comparativo (Excel)", icon=":material/table_view:", type="primary", use_container_width=True,
                       data=generar_excel(C["ref"], items, lineas, prov, det, adj, tc, ap_rep),
                       file_name=f"Cuadro_comparativo_{base_nombre}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("Resumen (PDF)", icon=":material/picture_as_pdf:", use_container_width=True,
                       data=generar_pdf(C["ref"], adj, lineas, tc, ap_rep),
                       file_name=f"Cuadro_comparativo_{base_nombre}.pdf", mime="application/pdf")
    d3.download_button("Guardar para seguir después", icon=":material/save:", use_container_width=True,
                       data=cuadro_a_json(C), file_name=f"Trabajo_{nombre_archivo(C['ref'])}.json",
                       mime="application/json", help="Para retomarlo: paso 1 → «Retomar un trabajo guardado»")

    st.write("")
    st.divider()
    n1, _, n3 = st.columns([1, 2, 1.4])
    n1.button("← Volver", on_click=ir_a, args=(4,), use_container_width=True)

    def nueva_cotizacion():
        ss.cuadro, ss.mat_sel, ss.sp_sel, ss.excluidas = None, set(), set(), set()
        ir_a(2)

    n3.button("Empezar otra cotización", on_click=nueva_cotizacion, use_container_width=True)
