# =============================================================================
# CUADRO COMPARATIVO DE OFERTAS - ABASTECIMIENTO ENAEX  (v2)
# -----------------------------------------------------------------------------
# Flujo: Requerimiento (SOLPED) -> Histórico (última compra ME2M)
#        -> Proveedores y ofertas -> Motor de comparación
#        -> Adjudicación con aprobación humana -> Excel / PDF / Borrador JSON
#
# Requisitos (requirements.txt):
#   streamlit>=1.35  pandas  numpy  openpyxl  requests  fpdf2
#   python-calamine  (opcional, lee Excel ~10x más rápido)
#   holidays         (opcional, feriados de Chile en fechas de entrega)
# =============================================================================
import io
import json
import re
import unicodedata
from datetime import date, datetime

import numpy as np
import pandas as pd
import requests
import streamlit as st
from fpdf import FPDF
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
# CONSTANTES DE NEGOCIO (tomadas de las planillas Cuadro MRO Rev3 y Autogestión)
# =============================================================================
MONEDAS = ["CLP", "USD", "EUR", "UF"]
TRANSPORTES = ["T. Gil", "T. Bello", "Pullman", "Retiramos", "EXW", "FCA", "FOB",
               "CFR", "CIF", "CPT", "CIP", "DAP", "DDP"]
TIPOS_ADJUDICACION = ["Mejor Oferta Económica", "Proveedor único",
                      "Aprobado Técnicamente", "Asignación Directa"]
TIPOS_PLAZO = ["Días hábiles", "Días corridos", "Semanas"]
TIPOS_COMPRA = ["Nacional", "Importación"]
# Hoja "Fecha de entrega" de Autogestión: días adicionales según lugar de entrega
LUGARES_ENTREGA = {
    "EDP - Entrega directa en planta": 0,
    "EBP - Entrega en bodega de proveedor": 0,
    "EBC - Entrega bodega consolidación": 7,
    "EXW A - Origen aéreo": 15,
    "EXW M - Origen marítimo": 45,
}
DIAS_EMISION_OC = 2          # WORKDAY(hoy, 2) en la planilla
DIAS_INTERNACION = 10        # IF(Tipo="Internación", 10, 0)
DIAS_BODEGA_TRANSITO = 7     # Bodega tránsito Isaver / Bello
UMBRAL_TC_CLP = 1_000_000    # Regla: >$1.000.000 CLP exige TC del día
UMBRAL_ALZA = 0.15           # Alerta si el precio sube más de 15% vs última compra
SIN_ADJ = "— Sin adjudicar —"

# Sinónimos de columnas (comparación exacta tras normalizar; nunca subcadenas cortas)
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

# =============================================================================
# PÁGINA Y ESTILOS
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
# UTILIDADES GENERALES
# =============================================================================
def norm(txt) -> str:
    """Minúsculas, sin tildes, sin signos; para comparar encabezados y textos."""
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    s = re.sub(r"[^a-z0-9]+", " ", s.lower())
    return re.sub(r"\s+", " ", s).strip()


def clave_material(v) -> str:
    """'30006049', 30006049.0, ' 0030006049 ' -> '30006049'. Vacío si no es código."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s).lstrip("0")
    return s if len(s) >= 4 else ""


def parse_monto(v):
    """Convierte '$49.100CLP', '1.234,56', '389,00 USD', 59500 en (valor, moneda|None)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return 0.0, None
    if isinstance(v, (int, float, np.number)):
        return float(v), None
    s = str(v).strip().upper()
    moneda = next((m for m in ["CLP", "USD", "EUR", "UF"] if m in s), None)
    s = re.sub(r"[^\d.,-]", "", s)
    if not s:
        return 0.0, moneda
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) != 3 else s.replace(",", "")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[-1]) == 3):
        s = s.replace(".", "")  # punto como separador de miles (formato chileno)
    try:
        return float(s), moneda
    except ValueError:
        return 0.0, moneda


def a_num(v, defecto=0.0) -> float:
    val, _ = parse_monto(v)
    return val if val else defecto


def mapear_columnas(columnas, sinonimos) -> dict:
    """Asigna cada campo canónico a una columna real: primero igualdad exacta, luego prefijo."""
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
# LECTURA DE ARCHIVOS (con caché y detección de encabezado)
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
    """Lee la hoja completa y ubica la fila de encabezados en las primeras 40 filas."""
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
    """Lleva la planilla de carga (ME5A o Ariba) a columnas estándar."""
    g = lambda campo: df[mapa[campo]] if campo in mapa else pd.Series([None] * len(df), index=df.index)
    out = pd.DataFrame({
        "SOLPED": g("solped").astype(str).str.strip().str.replace(r"\.0$", "", regex=True),
        "Pos": g("pos"),
        "Material": g("material").apply(clave_material),
        "Texto": g("texto").astype(str).str.strip().str.strip('"').replace({"None": "", "nan": ""}),
        "Cantidad": g("cantidad").apply(lambda v: a_num(v, 1.0)),
        "UM": g("um").fillna("C/U").astype(str),
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
    # Ariba no trae posición: se numera 10, 20, 30… dentro de cada SOLPED (como SAP)
    sin_pos = out["Pos"].isna() | (out["Pos"].astype(str).str.strip().isin(["", "None", "nan"]))
    out.loc[sin_pos, "Pos"] = out[sin_pos].groupby("SOLPED").cumcount().add(1).mul(10)
    out["Pos"] = out["Pos"].apply(lambda v: int(a_num(v, 0)))
    if out["Días"].isna().all() and out["F. solicitud"].notna().any():
        out["Días"] = (pd.Timestamp(date.today()) - out["F. solicitud"]).dt.days
    return out.drop_duplicates(subset=["SOLPED", "Pos"]).reset_index(drop=True)


@st.cache_data(show_spinner="Indexando histórico de compras…")
def construir_historico(df: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """Una fila por material con la última compra y estadísticas simples."""
    g = lambda campo: df[mapa[campo]] if campo in mapa else pd.Series([None] * len(df), index=df.index)
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
    stats = h.groupby("mat_key").agg(**{"N° compras": ("Precio", "size")})
    ult = h.drop_duplicates("mat_key").set_index("mat_key")
    ult = ult.join(stats)
    return ult.reset_index()


def buscar_historico(hist: pd.DataFrame, material: str, texto: str):
    """Busca por código SAP; si no hay código (Ariba), intenta por descripción exacta."""
    if hist is None or hist.empty:
        return None, ""
    if material:
        fila = hist[hist["mat_key"] == material]
        if not fila.empty:
            return fila.iloc[0], "Código"
    t = norm(texto)
    if t:
        fila = hist[hist["txt_key"] == t]
        if not fila.empty:
            return fila.iloc[0], "Descripción"
    return None, ""

# =============================================================================
# TIPOS DE CAMBIO
# =============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def obtener_indicadores():
    try:
        r = requests.get("https://mindicador.cl/api", timeout=6)
        if r.ok:
            d = r.json()
            return {
                "USD": float(d["dolar"]["valor"]), "EUR": float(d["euro"]["valor"]),
                "UF": float(d["uf"]["valor"]), "fecha": str(d["dolar"]["fecha"])[:10],
                "fuente": "mindicador.cl (Banco Central)", "ok": True,
            }
    except Exception:
        pass
    return {"USD": 0.0, "EUR": 0.0, "UF": 0.0, "fecha": "", "fuente": "Sin conexión", "ok": False}


def a_clp(monto, moneda, tc) -> float:
    moneda = str(moneda or "CLP").upper().strip()
    factor = {"CLP": 1.0, "USD": tc["USD"], "EUR": tc["EUR"], "UF": tc["UF"]}.get(moneda, 1.0)
    return float(monto or 0) * factor

# =============================================================================
# FECHAS DE ENTREGA (réplica de la hoja "Fecha de entrega" de Autogestión)
# =============================================================================
def sumar_habiles(f: np.datetime64, n: int) -> np.datetime64:
    return np.busday_offset(f, int(n), roll="forward", holidays=FERIADOS)


def calcular_fecha_planta(plazo, tipo_plazo, tipo_compra, lugar, bodega_transito, hoy=None):
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
# MOTOR DE COMPARACIÓN
# =============================================================================
def calcular_comparativo(pos: pd.DataFrame, prov: pd.DataFrame, ofertas: pd.DataFrame, tc: dict):
    """Devuelve (detalle por oferta, resumen por posición con sugerencia)."""
    prov = prov.set_index("Proveedor")
    det = ofertas.merge(pos[["Pos", "Cantidad", "Últ. precio CLP"]], on="Pos", how="left")
    det["Precio mejorado"] = det["Precio mejorado"].where(det["Precio mejorado"].fillna(0) > 0,
                                                          det["Precio inicial"])
    det = det[det["Proveedor"].isin(prov.index)].copy()
    det["Moneda"] = det["Proveedor"].map(prov["Moneda"])
    det["Cotiza"] = det["Precio mejorado"].fillna(0) > 0
    det["Subtotal inicial"] = det["Cantidad"] * det["Precio inicial"].fillna(0)
    det["Subtotal mejorado"] = det["Cantidad"] * det["Precio mejorado"].fillna(0)
    det["Subtotal CLP"] = [a_clp(m, mo, tc) for m, mo in zip(det["Subtotal mejorado"], det["Moneda"])]
    det["Ahorro spot CLP"] = [a_clp(i - m, mo, tc) if i > 0 else 0.0
                              for i, m, mo in zip(det["Subtotal inicial"], det["Subtotal mejorado"], det["Moneda"])]
    # Flete total del proveedor prorrateado según el valor de cada posición cotizada
    det["Flete CLP"] = 0.0
    for p, fila in prov.iterrows():
        flete_clp = a_clp(fila.get("Flete total", 0), fila.get("Moneda flete", "CLP"), tc)
        m = (det["Proveedor"] == p) & det["Cotiza"]
        base = det.loc[m, "Subtotal CLP"].sum()
        if flete_clp > 0 and base > 0:
            det.loc[m, "Flete CLP"] = flete_clp * det.loc[m, "Subtotal CLP"] / base
    det["Total CLP"] = det["Subtotal CLP"] + det["Flete CLP"]
    det["Precio unit. CLP"] = det["Subtotal CLP"] / det["Cantidad"].replace(0, np.nan)
    ult = det["Últ. precio CLP"].where(det["Últ. precio CLP"] > 0)
    det["Var. vs últ. compra"] = (det["Precio unit. CLP"] - ult) / ult
    # Fecha de entrega en planta por oferta
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
    for _, p in pos.iterrows():
        d = det[(det["Pos"] == p["Pos"]) & det["Válida"]].sort_values(["Total CLP", "Días a planta"])
        n_validas = len(d)
        mejor = d.iloc[0] if n_validas else None
        alertas = []
        if n_validas == 0:
            alertas.append("Sin ofertas válidas")
        if not p["Con historial"]:
            alertas.append("Sin historial: requiere validación técnica")
        if mejor is not None and pd.notna(mejor["Var. vs últ. compra"]) and mejor["Var. vs últ. compra"] > UMBRAL_ALZA:
            alertas.append(f"Alza {mejor['Var. vs últ. compra']:.0%} vs última compra")
        filas.append({
            "Pos": p["Pos"],
            "Sugerido": mejor["Proveedor"] if mejor is not None else SIN_ADJ,
            "Tipo sugerido": ("Proveedor único" if n_validas == 1 else "Mejor Oferta Económica") if n_validas else "",
            "Mejor total CLP": mejor["Total CLP"] if mejor is not None else np.nan,
            "Ofertas válidas": n_validas,
            "Alertas": "; ".join(alertas),
        })
    return det, pd.DataFrame(filas)

# =============================================================================
# EXPORTACIÓN
# =============================================================================
AZUL, VERDE, GRIS = "1E3A8A", "D1FAE5", "F9FAFB"
BORDE = Border(*(Side(style="thin", color="E5E7EB"),) * 4)


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


def generar_excel(sp, pos, prov, det, adj, tc, aprobacion) -> bytes:
    buf = io.BytesIO()
    proveedores = list(prov["Proveedor"])
    # Hoja 1: matriz como el Cuadro MRO (posiciones en filas, proveedores en columnas)
    base = pos[["Pos", "Material", "Texto", "Cantidad", "UM", "Últ. compra", "Últ. proveedor",
                "Últ. precio", "Últ. moneda"]].copy()
    for p in proveedores:
        d = det[det["Proveedor"] == p].set_index("Pos")
        base[f"{p} | P. inicial"] = base["Pos"].map(d["Precio inicial"])
        base[f"{p} | P. mejorado"] = base["Pos"].map(d["Precio mejorado"])
        base[f"{p} | Total CLP"] = base["Pos"].map(d["Total CLP"]).where(base["Pos"].map(d["Válida"]).fillna(False))
        base[f"{p} | Entrega"] = base["Pos"].map(d["Entrega planta"])
    base = base.merge(adj[["Pos", "Sugerido", "Proveedor adjudicado", "Tipo adjudicación",
                           "Monto adjudicado CLP", "Ahorro spot CLP", "Ahorro vs histórico CLP",
                           "Var. vs últ. compra", "Comentario", "Alertas"]], on="Pos", how="left")
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        base.to_excel(xw, sheet_name="Cuadro comparativo", index=False, startrow=5)
        det_out = det[["Pos", "Proveedor", "Moneda", "Precio inicial", "Precio mejorado", "Cantidad",
                       "Subtotal CLP", "Flete CLP", "Total CLP", "Ahorro spot CLP", "Var. vs últ. compra",
                       "Plazo", "Entrega planta", "Días a planta", "Cumple técnico", "Observación"]]
        det_out.to_excel(xw, sheet_name="Detalle ofertas", index=False)
        prov.to_excel(xw, sheet_name="Proveedores", index=False)
        pd.DataFrame({
            "Parámetro": ["SOLPED", "Fecha", "USD/CLP", "EUR/CLP", "UF/CLP", "Fecha TC", "Fuente TC",
                          "Aprobado por", "Fecha aprobación", "Estado"],
            "Valor": [sp, date.today().strftime("%d/%m/%Y"), tc["USD"], tc["EUR"], tc["UF"], tc["fecha"],
                      tc["fuente"], aprobacion.get("nombre", ""), aprobacion.get("fecha", ""),
                      "APROBADO" if aprobacion.get("ok") else "BORRADOR - pendiente de aprobación"],
        }).to_excel(xw, sheet_name="Parámetros", index=False)

        ws = xw.sheets["Cuadro comparativo"]
        ws["A1"] = f"ENAEX - CUADRO COMPARATIVO DE OFERTAS | SOLPED {sp}"
        ws["A1"].font = Font(size=15, bold=True, color=AZUL)
        ws["A2"] = (f"Fecha: {date.today():%d/%m/%Y} | USD {tc['USD']:,.2f} | EUR {tc['EUR']:,.2f} | "
                    f"UF {tc['UF']:,.2f} (TC al {tc['fecha'] or 'ingreso manual'})")
        ws["A3"] = ("APROBADO por " + aprobacion.get("nombre", "") + " el " + aprobacion.get("fecha", "")
                    if aprobacion.get("ok") else "BORRADOR - pendiente de aprobación del comprador")
        ws["A3"].font = Font(bold=True, color="065F46" if aprobacion.get("ok") else "B45309")
        ws["A4"] = (f"Monto adjudicado: $ {adj['Monto adjudicado CLP'].sum():,.0f} CLP | "
                    f"Ahorro spot: $ {adj['Ahorro spot CLP'].sum():,.0f} CLP")
        montos = [c for c in base.columns if any(k in c for k in ["precio", "P. inicial", "P. mejorado",
                                                                    "Total CLP", "Monto", "Ahorro"])]
        estilo_hoja(ws, 6, cols_monto=montos, cols_pct=["Var. vs últ. compra"],
                    cols_fecha=["Últ. compra"] + [c for c in base.columns if c.endswith("| Entrega")])
        # Resalta en verde el total adjudicado de cada fila
        col_idx = {c.value: c.column for c in ws[6]}
        for r, (_, fila) in enumerate(base.iterrows(), start=7):
            p = fila.get("Proveedor adjudicado")
            if p and f"{p} | Total CLP" in col_idx:
                ws.cell(row=r, column=col_idx[f"{p} | Total CLP"]).fill = PatternFill("solid", start_color=VERDE)
        estilo_hoja(xw.sheets["Detalle ofertas"], 1,
                    cols_monto=["Precio inicial", "Precio mejorado", "Subtotal CLP", "Flete CLP",
                                "Total CLP", "Ahorro spot CLP"],
                    cols_pct=["Var. vs últ. compra"], cols_fecha=["Entrega planta"])
        estilo_hoja(xw.sheets["Proveedores"], 1, cols_monto=["Flete total"])
        estilo_hoja(xw.sheets["Parámetros"], 1)
    return buf.getvalue()


def txt_pdf(s) -> str:
    return unicodedata.normalize("NFKD", str(s if s is not None else "")).encode("latin-1", "ignore").decode("latin-1")


def generar_pdf(sp, adj, tc, aprobacion) -> bytes:
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(30, 58, 138)
    pdf.cell(0, 8, txt_pdf(f"ENAEX - Cuadro comparativo SOLPED {sp}"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 5, txt_pdf(f"Fecha {date.today():%d/%m/%Y} | USD {tc['USD']:,.2f} | EUR {tc['EUR']:,.2f} | "
                           f"UF {tc['UF']:,.2f} | TC al {tc['fecha'] or 'ingreso manual'}"),
             new_x="LMARGIN", new_y="NEXT")
    estado = (f"APROBADO por {aprobacion.get('nombre', '')} el {aprobacion.get('fecha', '')}"
              if aprobacion.get("ok") else "BORRADOR - pendiente de aprobacion del comprador")
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, txt_pdf(estado), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    cols = [("Pos", 12), ("Descripcion", 70), ("Cant.", 14), ("Adjudicado", 48), ("Tipo", 38),
            ("Monto CLP", 30), ("Ahorro spot", 26), ("Var. hist.", 18), ("Entrega", 20)]
    pdf.set_fill_color(30, 58, 138)
    pdf.set_text_color(255, 255, 255)
    for n, w in cols:
        pdf.cell(w, 7, n, border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    for i, (_, r) in enumerate(adj.iterrows()):
        pdf.set_fill_color(*((249, 250, 251) if i % 2 else (255, 255, 255)))
        var = f"{r['Var. vs últ. compra']:.0%}" if pd.notna(r["Var. vs últ. compra"]) else "-"
        ent = r["Entrega planta"].strftime("%d/%m/%Y") if pd.notna(r.get("Entrega planta")) else "-"
        vals = [str(r["Pos"]), str(r["Texto"])[:45], f"{r['Cantidad']:g}", str(r["Proveedor adjudicado"])[:28],
                str(r["Tipo adjudicación"])[:24], f"{r['Monto adjudicado CLP']:,.0f}",
                f"{r['Ahorro spot CLP']:,.0f}", var, ent]
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
            pdf.multi_cell(0, 5, txt_pdf(f"Pos {r['Pos']}: {r['Comentario']}"), new_x="LMARGIN", new_y="NEXT")
    return bytes(pdf.output())

# =============================================================================
# ESTADO DE SESIÓN
# =============================================================================
ss = st.session_state
ss.setdefault("proyectos", {})       # SOLPED -> datos del cuadro
ss.setdefault("req", None)           # requerimientos normalizados
ss.setdefault("hist", None)          # histórico indexado por material


def nuevo_proyecto(sp: str, posiciones: pd.DataFrame):
    pos = posiciones.copy().reset_index(drop=True)
    hist = ss.hist
    filas = []
    for _, r in pos.iterrows():
        h, via = buscar_historico(hist, r.get("Material", ""), r.get("Texto", ""))
        filas.append({
            "Con historial": h is not None,
            "Historial por": via,
            "Últ. compra": h["Fecha"].date() if h is not None else None,
            "Últ. proveedor": h["Proveedor"] if h is not None else "",
            "Últ. precio": h["Precio"] if h is not None else np.nan,
            "Últ. moneda": h["Moneda"] if h is not None else "",
            "Últ. OC": h["OC"] if h is not None else "",
            "N° compras": int(h["N° compras"]) if h is not None else 0,
        })
    pos = pd.concat([pos, pd.DataFrame(filas)], axis=1)
    # Proveedores sugeridos: los que ya vendieron estos materiales (scouting por historial)
    sugeridos = [p for p in pos["Últ. proveedor"].dropna().unique() if p][:3] or ["Proveedor 1"]
    prov = pd.DataFrame({
        "Proveedor": sugeridos, "Moneda": "CLP", "Transporte": "T. Gil", "Flete total": 0.0,
        "Moneda flete": "CLP", "Plazo": 10, "Tipo plazo": "Días hábiles", "Tipo compra": "Nacional",
        "Lugar entrega": list(LUGARES_ENTREGA)[0], "Bodega tránsito": False,
    })
    ss.proyectos[sp] = {"pos": pos, "prov_src": prov, "prov_v": 0, "prov_names": None,
                        "of_src": None, "of_v": 0, "of_last": None, "adj_override": {},
                        "aprobacion": {"ok": False, "nombre": "", "fecha": ""}}


def armar_ofertas(pos, nombres, previas):
    """Grilla Pos x Proveedor, conservando lo que ya se había digitado."""
    base = pd.DataFrame([(p, t, n) for p, t in zip(pos["Pos"], pos["Texto"]) for n in nombres],
                        columns=["Pos", "Texto", "Proveedor"])
    for c, v in [("Precio inicial", np.nan), ("Precio mejorado", np.nan), ("Plazo", np.nan),
                 ("Cumple técnico", True), ("Observación", "")]:
        base[c] = v
    if previas is not None and not previas.empty:
        prev = previas.drop(columns=["Texto"], errors="ignore").drop_duplicates(["Pos", "Proveedor"])
        base = base[["Pos", "Texto", "Proveedor"]].merge(prev, on=["Pos", "Proveedor"], how="left")
        base["Cumple técnico"] = base["Cumple técnico"].fillna(True).astype(bool)
        base["Observación"] = base["Observación"].fillna("")
    return base


def proyecto_a_json(sp, P) -> str:
    def df_json(df):
        return None if df is None else df.to_json(orient="records", date_format="iso", force_ascii=False)
    return json.dumps({"solped": sp, "pos": df_json(P["pos"]), "prov": df_json(P["prov_src"]),
                       "ofertas": df_json(P["of_last"]), "adj_override": P["adj_override"],
                       "aprobacion": P["aprobacion"]}, ensure_ascii=False)


def json_a_proyecto(texto: str):
    d = json.loads(texto)
    leer = lambda s: pd.read_json(io.StringIO(s), orient="records") if s else None
    pos = leer(d["pos"])
    if "Últ. compra" in pos:
        pos["Últ. compra"] = pd.to_datetime(pos["Últ. compra"], errors="coerce").dt.date
    ss.proyectos[d["solped"]] = {
        "pos": pos, "prov_src": leer(d["prov"]), "prov_v": 1, "prov_names": None,
        "of_src": None, "of_v": 1, "of_last": leer(d["ofertas"]),
        "adj_override": {int(k): v for k, v in d.get("adj_override", {}).items()},
        "aprobacion": d.get("aprobacion", {"ok": False, "nombre": "", "fecha": ""}),
    }
    return d["solped"]

# =============================================================================
# BARRA LATERAL: TIPOS DE CAMBIO Y ARCHIVOS
# =============================================================================
st.markdown("<div class='titulo'>Cuadro Comparativo de Ofertas</div>"
            "<div class='subtitulo'>SOLPED → histórico de compras → ofertas → adjudicación con aprobación del comprador</div>",
            unsafe_allow_html=True)

with st.sidebar:
    st.header("Tipos de cambio")
    ind = obtener_indicadores()
    if ind["ok"]:
        st.success(f"TC al {ind['fecha']} · {ind['fuente']}")
    else:
        st.error("Sin conexión a mindicador.cl: ingresa el TC del día manualmente.")
    if st.button("Actualizar tasas", use_container_width=True):
        obtener_indicadores.clear()   # limpia solo esta caché, no la de archivos
        st.rerun()
    tc = {
        "USD": st.number_input("USD / CLP", value=ind["USD"] or 0.0, step=1.0, format="%.2f"),
        "EUR": st.number_input("EUR / CLP", value=ind["EUR"] or 0.0, step=1.0, format="%.2f"),
        "UF": st.number_input("UF / CLP", value=ind["UF"] or 0.0, step=10.0, format="%.2f"),
        "fecha": ind["fecha"], "fuente": ind["fuente"],
    }
    manual = (tc["USD"], tc["EUR"], tc["UF"]) != (ind["USD"], ind["EUR"], ind["UF"])
    if manual:
        tc["fecha"], tc["fuente"] = date.today().isoformat(), "Ingreso manual"
    tc_hoy = tc["fecha"] == date.today().isoformat()
    if min(tc["USD"], tc["EUR"], tc["UF"]) <= 0:
        st.warning("Completa los tres tipos de cambio para convertir monedas.")

    st.divider()
    st.header("Archivos")
    f_req = st.file_uploader("1. Carga de SOLPED (Autogestión, ME5A o Ariba)", type=["xlsx", "xlsm", "xls", "csv"])
    if f_req:
        hojas = listar_hojas(f_req.getvalue(), f_req.name)
        pref = ["Carga ARIBA", "Carga", "ME5A", "Planilla anterior"]
        hoja_req = st.selectbox("Hoja de requerimientos", hojas,
                                index=int(np.argmax([puntaje_hoja(h, pref) for h in hojas])))
    f_hist = st.file_uploader("2. Histórico de compras (Cuadro MRO, ME2M o CSV)", type=["xlsx", "xlsm", "xls", "csv"])
    if f_hist:
        hojas_h = listar_hojas(f_hist.getvalue(), f_hist.name)
        pref_h = ["Ultima COMPRA", "BASE OC", "ME2M", "Data"]
        hoja_hist = st.selectbox("Hoja de histórico", hojas_h,
                                 index=int(np.argmax([puntaje_hoja(h, pref_h) for h in hojas_h])))
        st.caption("Tip: exportar el histórico a CSV acelera mucho la carga.")
    if st.button("Procesar archivos", type="primary", use_container_width=True):
        if f_req:
            df = leer_hoja(f_req.getvalue(), f_req.name, hoja_req, json.dumps(SIN_REQ))
            ss.req_raw, ss.req_map = df, mapear_columnas(df.columns, SIN_REQ)
        if f_hist:
            dfh = leer_hoja(f_hist.getvalue(), f_hist.name, hoja_hist, json.dumps(SIN_HIST))
            ss.hist_raw, ss.hist_map = dfh, mapear_columnas(dfh.columns, SIN_HIST)
        ss.req = ss.hist = None   # se reconstruyen abajo con el mapeo confirmado

    # Mapeo de columnas revisable (evita que una columna mal detectada contamine el cuadro)
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

    if ss.get("req_raw") is not None and all(c in ss.req_map for c in REQ_OBLIGATORIAS):
        ss.req = normalizar_requerimientos(ss.req_raw, ss.req_map)
    if ss.get("hist_raw") is not None and all(c in ss.hist_map for c in HIST_OBLIGATORIAS):
        ss.hist = construir_historico(ss.hist_raw, ss.hist_map)
    if ss.req is not None:
        st.caption(f"{ss.req['SOLPED'].nunique()} SOLPED · {len(ss.req)} posiciones")
    if ss.hist is not None:
        st.caption(f"Histórico: {len(ss.hist):,} materiales con última compra")

    st.divider()
    st.header("Borradores")
    f_json = st.file_uploader("Retomar borrador (.json)", type=["json"])
    if f_json and st.button("Cargar borrador", use_container_width=True):
        sp_cargada = json_a_proyecto(f_json.getvalue().decode("utf-8"))
        ss.sp_activa = sp_cargada
        st.success(f"Borrador de SOLPED {sp_cargada} cargado.")

# =============================================================================
# PESTAÑAS
# =============================================================================
tab_bandeja, tab_ofertas, tab_adj = st.tabs(["1. Bandeja de SOLPED", "2. Proveedores y ofertas",
                                              "3. Adjudicación y reportes"])

# ---------------------------------------------------------------- TAB 1
with tab_bandeja:
    if ss.req is not None and not ss.req.empty:
        req = ss.req
        compradores = sorted([c for c in req["Asignado a"].unique() if c])
        c1, c2 = st.columns([2, 3])
        filtro_comp = c1.multiselect("Comprador", compradores) if compradores else []
        texto_busca = c2.text_input("Buscar por SOLPED, material o descripción")
        vista = req if not filtro_comp else req[req["Asignado a"].isin(filtro_comp)]
        if texto_busca:
            t = norm(texto_busca)
            vista = vista[vista.apply(lambda r: t in norm(f"{r['SOLPED']} {r['Material']} {r['Texto']}"), axis=1)]
        bandeja = (vista.groupby("SOLPED")
                   .agg(Posiciones=("Pos", "size"), Días=("Días", "max"), Centro=("Centro", "first"),
                        Comprador=("Asignado a", "first"), Urgencia=("Urgencia", "first"),
                        Ejemplo=("Texto", "first"))
                   .reset_index().sort_values("Días", ascending=False))
        bandeja["Tramo"] = pd.cut(bandeja["Días"], [-1, 5, 15, 30, 10_000],
                                  labels=["0-5 días", "6-15 días", "16-30 días", "30+ días"]).astype(str)
        k1, k2, k3 = st.columns(3)
        k1.metric("SOLPED pendientes", len(bandeja))
        k2.metric("Sobre 15 días", int((bandeja["Días"] > 15).sum()))
        k3.metric("Sobre 30 días", int((bandeja["Días"] > 30).sum()))
        st.caption("Ordenadas por antigüedad. Selecciona una fila para abrir su cuadro comparativo.")
        sel = st.dataframe(bandeja, use_container_width=True, hide_index=True, height=320,
                           on_select="rerun", selection_mode="single-row",
                           column_config={"Días": st.column_config.NumberColumn(format="%d")})
        if sel.selection.rows:
            sp = str(bandeja.iloc[sel.selection.rows[0]]["SOLPED"])
            if st.button(f"Trabajar SOLPED {sp}", type="primary"):
                if sp not in ss.proyectos:
                    nuevo_proyecto(sp, req[req["SOLPED"] == sp])
                ss.sp_activa = sp
                st.success(f"SOLPED {sp} lista. Continúa en la pestaña 2.")
    else:
        st.info("Carga la planilla de SOLPED en la barra lateral, o crea un cuadro manual aquí abajo.")

    with st.expander("Crear cuadro manual (sin planilla)", expanded=ss.req is None):
        sp_man = st.text_input("N° SOLPED", key="sp_manual")
        plantilla = pd.DataFrame({"Pos": [10], "Material": [""], "Texto": [""], "Cantidad": [1.0],
                                  "UM": ["C/U"], "Centro": [""]})
        pos_man = st.data_editor(plantilla, num_rows="dynamic", use_container_width=True, key="pos_manual")
        if st.button("Crear cuadro manual") and sp_man.strip():
            pm = pos_man.dropna(subset=["Texto"])
            pm = pm[pm["Texto"].astype(str).str.strip() != ""].copy()
            pm["Material"] = pm["Material"].apply(clave_material)
            nuevo_proyecto(sp_man.strip(), pm)
            ss.sp_activa = sp_man.strip()
            st.success("Cuadro creado. Continúa en la pestaña 2.")

    if ss.proyectos:
        st.caption("Cuadros abiertos en esta sesión: " + ", ".join(ss.proyectos))

# ---------------------------------------------------------------- SOLPED ACTIVA
sp = ss.get("sp_activa")
P = ss.proyectos.get(sp) if sp else None

# ---------------------------------------------------------------- TAB 2
with tab_ofertas:
    if P is None:
        st.info("Elige una SOLPED en la pestaña 1.")
    else:
        if len(ss.proyectos) > 1:
            sp = st.selectbox("SOLPED activa", list(ss.proyectos), index=list(ss.proyectos).index(sp))
            ss.sp_activa, P = sp, ss.proyectos[sp]
        pos = P["pos"]
        tc_ok = min(tc["USD"], tc["EUR"], tc["UF"]) > 0
        pos["Últ. precio CLP"] = [a_clp(p, m, tc) if pd.notna(p) else np.nan
                                  for p, m in zip(pos["Últ. precio"], pos["Últ. moneda"])]
        pos["Últ. precio CLP"] = pos["Últ. precio CLP"].where(pos["Últ. precio CLP"] > 0)

        st.subheader(f"SOLPED {sp} · {len(pos)} posiciones")
        st.dataframe(
            pos[["Pos", "Material", "Texto", "Cantidad", "UM", "Centro", "Últ. compra", "Últ. proveedor",
                 "Últ. precio", "Últ. moneda", "N° compras", "Historial por"]],
            hide_index=True, use_container_width=True,
            column_config={"Últ. precio": st.column_config.NumberColumn(format="%.2f"),
                           "Últ. compra": st.column_config.DateColumn(format="DD/MM/YYYY")})
        sin_hist = int((~pos["Con historial"]).sum())
        if sin_hist:
            st.warning(f"{sin_hist} posición(es) sin historial de compra: requieren validación técnica del especialista.")
        if ss.hist is None:
            st.caption("Sin histórico cargado: no se calcula variación vs última compra ni bypass técnico.")

        st.markdown("#### Proveedores")
        st.caption("Agrega una fila por proveedor que cotizó. El flete se prorratea entre sus posiciones según valor.")
        prov_edit = st.data_editor(
            P["prov_src"], key=f"prov_{sp}_{P['prov_v']}", num_rows="dynamic", use_container_width=True,
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
        for c, v in [("Moneda", "CLP"), ("Moneda flete", "CLP"), ("Flete total", 0.0), ("Plazo", 0),
                     ("Tipo plazo", "Días hábiles"), ("Tipo compra", "Nacional"),
                     ("Lugar entrega", list(LUGARES_ENTREGA)[0]), ("Bodega tránsito", False)]:
            prov[c] = prov[c].fillna(v) if c in prov else v
        nombres = tuple(prov["Proveedor"])

        if nombres != P["prov_names"] or P["of_src"] is None:
            P["of_src"] = armar_ofertas(pos, nombres, P["of_last"])
            P["of_v"] += 1
            P["prov_names"] = nombres

        st.markdown("#### Ofertas por posición")
        st.caption("Precio mejorado vacío = se usa el inicial. Desmarca 'Cumple técnico' si la oferta no cumple la ficha. "
                   "Plazo vacío = plazo por defecto del proveedor.")
        of_edit = st.data_editor(
            P["of_src"], key=f"of_{sp}_{P['of_v']}", use_container_width=True, hide_index=True,
            height=min(80 + 35 * len(P["of_src"]), 600),
            disabled=["Pos", "Texto", "Proveedor"],
            column_config={
                "Texto": st.column_config.TextColumn(width="large"),
                "Precio inicial": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                "Precio mejorado": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                "Plazo": st.column_config.NumberColumn(min_value=0, step=1),
                "Cumple técnico": st.column_config.CheckboxColumn(),
            })
        P["of_last"] = of_edit
        P["prov_actual"] = prov

        if not tc_ok:
            st.error("Faltan tipos de cambio: no se puede comparar entre monedas.")
        elif of_edit["Precio inicial"].fillna(0).gt(0).any():
            det, res = calcular_comparativo(pos, prov, of_edit, tc)
            P["det"], P["res"] = det, res
            st.markdown("#### Comparación (total puesto en planta, CLP)")
            matriz = det.pivot_table(index="Pos", columns="Proveedor", values="Total CLP", aggfunc="first")
            validas = det.pivot_table(index="Pos", columns="Proveedor", values="Válida", aggfunc="first")
            matriz = matriz.where(validas.fillna(False).astype(bool))

            def verde_min(fila):
                if fila.notna().sum() == 0:
                    return [""] * len(fila)
                return ["background-color:#D1FAE5;font-weight:600" if v == fila.min() else "" for v in fila]

            st.dataframe(matriz.style.apply(verde_min, axis=1).format("$ {:,.0f}", na_rep="—"),
                         use_container_width=True)
            st.caption("Verde = menor total válido (incluye flete prorrateado). '—' = no cotiza o no cumple técnico.")
        else:
            st.info("Ingresa precios en la grilla de ofertas para ver la comparación.")

# ---------------------------------------------------------------- TAB 3
with tab_adj:
    if P is None or P.get("res") is None:
        st.info("Completa proveedores y ofertas en la pestaña 2.")
    else:
        pos, det, res, prov = P["pos"], P["det"], P["res"], P["prov_actual"]
        st.subheader(f"Adjudicación SOLPED {sp}")
        st.caption("El sistema propone el menor total válido. El comprador decide; si se aparta de la sugerencia, "
                   "el comentario es obligatorio.")

        src = res.merge(pos[["Pos", "Texto", "Cantidad"]], on="Pos")
        src["Proveedor adjudicado"] = src["Sugerido"]
        src["Tipo adjudicación"] = src["Tipo sugerido"].replace("", TIPOS_ADJUDICACION[0])
        src["Comentario"] = ""
        for pos_n, ov in P["adj_override"].items():
            m = src["Pos"] == pos_n
            if m.any():
                for k, v in ov.items():
                    src.loc[m, k] = v
        src = src[["Pos", "Texto", "Cantidad", "Sugerido", "Mejor total CLP", "Ofertas válidas", "Alertas",
                   "Proveedor adjudicado", "Tipo adjudicación", "Comentario"]]
        firma = str(pd.util.hash_pandas_object(src.astype(str), index=False).sum())
        adj_edit = st.data_editor(
            src, key=f"adj_{sp}_{firma}", hide_index=True, use_container_width=True,
            disabled=["Pos", "Texto", "Cantidad", "Sugerido", "Mejor total CLP", "Ofertas válidas", "Alertas"],
            column_config={
                "Texto": st.column_config.TextColumn(width="medium"),
                "Mejor total CLP": st.column_config.NumberColumn(format="$ %.0f"),
                "Proveedor adjudicado": st.column_config.SelectboxColumn(options=[SIN_ADJ] + list(prov["Proveedor"])),
                "Tipo adjudicación": st.column_config.SelectboxColumn(options=TIPOS_ADJUDICACION),
                "Comentario": st.column_config.TextColumn(width="large"),
            })
        P["adj_override"] = {int(r["Pos"]): {"Proveedor adjudicado": r["Proveedor adjudicado"],
                                             "Tipo adjudicación": r["Tipo adjudicación"],
                                             "Comentario": r["Comentario"] or ""}
                             for _, r in adj_edit.iterrows()}

        # Valores finales según lo que decidió el comprador
        adj = adj_edit.merge(
            det.rename(columns={"Proveedor": "Proveedor adjudicado"})[
                ["Pos", "Proveedor adjudicado", "Total CLP", "Ahorro spot CLP", "Var. vs últ. compra",
                 "Entrega planta", "Válida", "Precio unit. CLP"]],
            on=["Pos", "Proveedor adjudicado"], how="left")
        adj = adj.merge(pos[["Pos", "Últ. precio CLP"]], on="Pos", how="left")
        adj["Monto adjudicado CLP"] = adj["Total CLP"].fillna(0)
        adj["Ahorro spot CLP"] = adj["Ahorro spot CLP"].fillna(0)
        adj["Ahorro vs histórico CLP"] = ((adj["Últ. precio CLP"] - adj["Precio unit. CLP"]) * adj["Cantidad"]).fillna(0)
        adj["Válida"] = adj["Válida"].astype("boolean").fillna(False).astype(bool)

        problemas = []
        for _, r in adj.iterrows():
            if r["Proveedor adjudicado"] == SIN_ADJ:
                problemas.append(f"Pos {r['Pos']}: sin proveedor adjudicado.")
            elif not r["Válida"]:
                problemas.append(f"Pos {r['Pos']}: {r['Proveedor adjudicado']} no tiene oferta válida (precio o técnico).")
            elif r["Proveedor adjudicado"] != r["Sugerido"] and not str(r["Comentario"]).strip():
                problemas.append(f"Pos {r['Pos']}: se adjudica distinto a la sugerencia; falta comentario.")

        total = adj["Monto adjudicado CLP"].sum()
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Monto adjudicado", f"$ {total:,.0f}")
        k2.metric("Ahorro spot (negociación)", f"$ {adj['Ahorro spot CLP'].sum():,.0f}")
        k3.metric("Ahorro vs última compra", f"$ {adj['Ahorro vs histórico CLP'].sum():,.0f}")
        k4.metric("Posiciones adjudicadas", f"{int((adj['Proveedor adjudicado'] != SIN_ADJ).sum())} / {len(adj)}")

        if total > UMBRAL_TC_CLP and not tc_hoy:
            st.error(f"Monto sobre $ {UMBRAL_TC_CLP:,.0f} CLP: actualiza el tipo de cambio del día "
                     f"(el actual es del {tc['fecha'] or 'sin fecha'}).")
        for p in problemas:
            st.warning(p)

        por_prov = adj[adj["Proveedor adjudicado"] != SIN_ADJ].groupby("Proveedor adjudicado")["Monto adjudicado CLP"].sum()
        if len(por_prov) > 1:
            st.bar_chart(por_prov, horizontal=True, height=60 + 40 * len(por_prov))

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
        adj_rep = adj.merge(pos[["Pos"]], on="Pos")
        sufijo = "" if P["aprobacion"]["ok"] else "_BORRADOR"
        d1, d2, d3 = st.columns(3)
        d1.download_button("Excel cuadro comparativo",
                           data=generar_excel(sp, pos, prov, det, adj_rep, tc, P["aprobacion"]),
                           file_name=f"Cuadro_comparativo_{sp}_{date.today():%Y%m%d}{sufijo}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)
        d2.download_button("PDF resumen", data=generar_pdf(sp, adj_rep, tc, P["aprobacion"]),
                           file_name=f"Cuadro_comparativo_{sp}_{date.today():%Y%m%d}{sufijo}.pdf",
                           mime="application/pdf", use_container_width=True)
        d3.download_button("Guardar borrador (.json)", data=proyecto_a_json(sp, P),
                           file_name=f"Borrador_{sp}.json", mime="application/json", use_container_width=True)
