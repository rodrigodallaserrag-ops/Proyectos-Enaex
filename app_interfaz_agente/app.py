import datetime
import io
import re
import pandas as pd
import requests
import streamlit as st
import urllib3
import plotly.express as px

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. CONEXIÓN A SHAREPOINT / ONEDRIVE
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def cargar_maestro_solpeds(url_sharepoint=""):
    if url_sharepoint:
        try:
            direct_url = url_sharepoint.replace("?e=", "&download=1").replace("p=1", "download=1")
            df = pd.read_excel(direct_url, dtype={"SOLPED": str, "CODIGO_SAP": str, "POS": int})
            return df, "SharePoint Sincronizado 🟢"
        except Exception:
            pass
            
    data_demo = {
        "SOLPED": ["10045982", "10045982", "10045983", "10045984"],
        "POS": [10, 20, 10, 10],
        "SOCIEDAD": ["EC01", "EC01", "EC06", "EC01"],
        "CODIGO_SAP": ["3001892", "3001893", "4001002", "3002550"],
        "DESCRIPCION": ["VÁLVULA DE BOLA 2 INCH ANSI 300", "EMPADRÓN BASTIDOR SOPORTE", "MANGUERA ALTA PRESIÓN 1/2", "KITS DE EMPADRADO COMPLETO"],
        "UM": ["C/U", "C/U", "MTR", "SET"],
        "ULTIMA_COMPRA_MONTO": [180000, 450000, 25000, 1200000],
        "ULTIMA_COMPRA_MONEDA": ["CLP", "CLP", "CLP", "CLP"],
        "PROVEEDOR_HISTORICO": ["MCM CHILE", "INDURA", "PARKER", "SKF CHILE"]
    }
    df_demo = pd.DataFrame(data_demo)
    df_demo["SOLPED"] = df_demo["SOLPED"].astype(str)
    return df_demo, "Modo Demostración Local 🟡"

# -----------------------------------------------------------------------------
# 2. MOTOR FINANCIERO DE MONEDAS (MULTI-FUENTE ACTIVO)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1800)
def obtener_indicadores_financieros():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    # Intento 1: Mindicador.cl
    try:
        r = requests.get("https://mindicador.cl/api", headers=headers, verify=False, timeout=5)
        if r.status_code == 200:
            data = r.json()
            if "dolar" in data and "euro" in data and "uf" in data:
                return {
                    "dolar": float(data["dolar"]["valor"]),
                    "euro": float(data["euro"]["valor"]),
                    "uf": float(data["uf"]["valor"]),
                    "fecha": data["dolar"]["fecha"][:10],
                    "estado": "Online (Mindicador) 🟢",
                }
    except Exception:
        pass

    # Intento 2: DolarAPI Chile
    try:
        r_usd = requests.get("https://cl.dolarapi.com/v1/cotizaciones/usd", headers=headers, verify=False, timeout=5)
        r_eur = requests.get("https://cl.dolarapi.com/v1/cotizaciones/eur", headers=headers, verify=False, timeout=5)
        r_uf = requests.get("https://cl.dolarapi.com/v1/cotizaciones/uf", headers=headers, verify=False, timeout=5)
        if r_usd.status_code == 200:
            return {
                "dolar": float(r_usd.json().get("promedio", 938.0)),
                "euro": float(r_eur.json().get("promedio", 1020.0)) if r_eur.status_code == 200 else 1020.0,
                "uf": float(r_uf.json().get("promedio", 40875.0)) if r_uf.status_code == 200 else 40875.0,
                "fecha": datetime.date.today().strftime("%Y-%m-%d"),
                "estado": "Online (DolarAPI) 🟢",
            }
    except Exception:
        pass

    # Respaldo Offline
    return {
        "dolar": 938.0,
        "euro": 1020.0,
        "uf": 40875.0,
        "fecha": datetime.date.today().strftime("%d-%m-%Y"),
        "estado": "Offline (Red Estricta) 🛡️",
    }

# -----------------------------------------------------------------------------
# 3. MOTOR DE VINCULACIÓN AUTOMÁTICA DE EXCEL DE PROVEEDORES
# -----------------------------------------------------------------------------
def vincular_cotizaciones_excel(df_proveedores, solped_id, solped_info, indicadores):
    if df_proveedores is None or df_proveedores.empty:
        return []

    df_p = df_proveedores.copy()
    df_p.columns = [str(c).strip().upper() for c in df_p.columns]

    # Identificar columna SOLPED/ID
    col_solped = next((c for c in df_p.columns if any(k in c for k in ["SOLPED", "ID", "REQ", "NUMERO"])), None)
    if not col_solped:
        return []

    df_filtered = df_p[df_p[col_solped].astype(str).str.strip() == str(solped_id).strip()]
    if df_filtered.empty:
        return []

    dolar_actual = indicadores["dolar"]
    euro_actual = indicadores["euro"]
    uf_actual = indicadores["uf"]
    ult_compra = solped_info.get("ULTIMA_COMPRA_MONTO", 0)

    # Identificación flexible de columnas
    col_prov = next((c for c in df_p.columns if any(k in c for k in ["PROV", "NOMBRE", "VENDOR", "EMPRESA"])), "PROVEEDOR")
    col_monto = next((c for c in df_p.columns if any(k in c for k in ["MONTO", "PRECIO", "VALOR", "OFERTA"])), "MONTO")
    col_moneda = next((c for c in df_p.columns if any(k in c for k in ["MONEDA", "CURRENCY"])), "MONEDA")
    col_fecha = next((c for c in df_p.columns if any(k in c for k in ["FECHA", "PLAZO", "ENTREGA"])), "FECHA_ENTREGA")
    col_obs = next((c for c in df_p.columns if any(k in c for k in ["OBS", "NOTA", "COMENTARIO"])), "OBSERVACIONES")

    cotizaciones_mapeadas = []
    for _, row in df_filtered.iterrows():
        prov = str(row.get(col_prov, "PROVEEDOR EXCEL")).strip()
        moneda = str(row.get(col_moneda, "CLP")).strip().upper()
        if moneda not in ["CLP", "USD", "EUR", "UF"]:
            moneda = "CLP"

        try:
            monto = float(row.get(col_monto, 0))
        except (ValueError, TypeError):
            monto = 0.0

        if monto <= 0 or not prov:
            continue

        monto_clp = monto
        if moneda == "USD": monto_clp = monto * dolar_actual
        elif moneda == "EUR": monto_clp = monto * euro_actual
        elif moneda == "UF": monto_clp = monto * uf_actual

        monto_usd = monto_clp / dolar_actual if dolar_actual > 0 else 0
        var_pct = ((monto_clp - ult_compra) / ult_compra * 100) if ult_compra > 0 else 0.0

        cotizaciones_mapeadas.append({
            "POS": solped_info.get("POS", 10),
            "Proveedor": prov,
            "Monto Original": monto,
            "Moneda": moneda,
            "Equiv. CLP ($)": round(monto_clp, 2),
            "Equiv. USD ($)": round(monto_usd, 2),
            "Var % vs Hist": round(var_pct, 1),
            "Fecha de Entrega": str(row.get(col_fecha, "A convenir")),
            "Observaciones": str(row.get(col_obs, "")) if pd.notna(row.get(col_obs)) else "",
        })

    return cotizaciones_mapeadas

# Formatters y Generadores de PDF/Excel
indicadores_cache = obtener_indicadores_financieros()
indicadores = dict(indicadores_cache)

def formato_clp(valor):
    return f"${int(valor):,}".replace(",", ".")

def aplicar_formato_regional(monto, moneda):
    if moneda == "CLP": return f"$ {int(monto):,}".replace(",", ".")
    elif moneda == "USD": return f"$ {monto:,.2f}"
    elif moneda in ["EUR", "UF"]: return f"{'€' if moneda=='EUR' else 'UF'} {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(monto)

def generar_pdf_ejecutivo(solped_info, cotizaciones, datos_indicadores):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, leading=18, textColor=colors.HexColor('#0F172A'))
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=9, leading=11, textColor=colors.HexColor('#475569'), spaceAfter=10)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontSize=8, leading=10)
    header_table_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=7, leading=9, textColor=colors.whitesmoke, fontName='Helvetica-Bold')
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7, leading=9)

    story.append(Paragraph("<b>ENAEX — Consola de Compras</b>", title_style))
    story.append(Paragraph(f"Reporte Ejecutivo de Adjudicación — Solped N° <b>{solped_info['SOLPED']}</b>", subtitle_style))

    meta_data = [
        [Paragraph(f"<b>N° Solped:</b> {solped_info['SOLPED']}", normal_style), Paragraph(f"<b>POS:</b> {solped_info['POS']}", normal_style), Paragraph(f"<b>Sociedad:</b> {solped_info['SOCIEDAD']}", normal_style)],
        [Paragraph(f"<b>Código SAP:</b> {solped_info['CODIGO_SAP']}", normal_style), Paragraph(f"<b>Descripción:</b> {solped_info['DESCRIPCION']}", normal_style), Paragraph(f"<b>UM:</b> {solped_info['UM']}", normal_style)],
        [Paragraph(f"<b>Última Compra:</b> {formato_clp(solped_info['ULTIMA_COMPRA_MONTO'])}", normal_style), Paragraph(f"<b>Prov. Histórico:</b> {solped_info['PROVEEDOR_HISTORICO']}", normal_style), Paragraph(f"<b>Fecha Emisión:</b> {datetime.date.today().strftime('%d-%m-%Y')}", normal_style)]
    ]
    t_meta = Table(meta_data, colWidths=[180, 180, 180])
    t_meta.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')), ('PADDING', (0,0), (-1,-1), 5), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1'))]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    table_data = [[Paragraph("<b>Proveedor</b>", header_table_style), Paragraph("<b>Monto Orig.</b>", header_table_style), Paragraph("<b>Mon.</b>", header_table_style), Paragraph("<b>Equiv. CLP ($)</b>", header_table_style), Paragraph("<b>Var. % Hist.</b>", header_table_style), Paragraph("<b>Plazo Entrega</b>", header_table_style), Paragraph("<b>Observaciones</b>", header_table_style)]]
    min_clp = min([c["Equiv. CLP ($)"] for c in cotizaciones]) if cotizaciones else 0

    for c in cotizaciones:
        prov_text = f"<b>{c['Proveedor']}</b>"
        if len(cotizaciones) > 1 and c["Equiv. CLP ($)"] == min_clp: prov_text += "<br/><font color='#16A34A'><b>★ Mejor Oferta</b></font>"
        var_pct = c.get("Var % vs Hist", 0)
        table_data.append([
            Paragraph(prov_text, cell_style), Paragraph(aplicar_formato_regional(c["Monto Original"], c["Moneda"]), cell_style),
            Paragraph(c["Moneda"], cell_style), Paragraph(f"$ {int(c['Equiv. CLP ($)']):,}".replace(",", "."), cell_style),
            Paragraph(f"<font color='{'#16A34A' if var_pct <= 0 else '#DC2626'}'><b>{var_pct:+.1f}%</b></font>", cell_style),
            Paragraph(str(c["Fecha de Entrega"]), cell_style), Paragraph(c.get("Observaciones", "-") or "-", cell_style)
        ])

    t_quotes = Table(table_data, colWidths=[100, 65, 30, 75, 60, 90, 120])
    t_quotes.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')), ('PADDING', (0,0), (-1,-1), 5)]))
    story.append(t_quotes)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def generar_excel_estilizado(df, solped_info, datos_indicadores):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cuadro Comparativo"
    navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    thin_border = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    ws.cell(row=1, column=1, value="ENAEX — Consola de Compras").font = Font(name="Calibri", size=15, bold=True, color="0F172A")
    headers = ["POS", "Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Var. % Hist.", "Fecha Entrega", "Observaciones"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = header_font

    min_clp = df["Equiv. CLP ($)"].min() if not df.empty else 0
    for r_offset, (_, row) in enumerate(df.iterrows()):
        r = 5 + r_offset
        is_best = (len(df) > 1 and row["Equiv. CLP ($)"] == min_clp)
        vals = [solped_info["POS"], row["Proveedor"], row["Monto Original"], row["Moneda"], row["Equiv. CLP ($)"], row["Equiv. USD ($)"], f"{row.get('Var % vs Hist', 0):+.1f}%", row["Fecha de Entrega"], row.get("Observaciones", "")]
        for c_i, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c_i, value=v)
            cell.font = bold_font if is_best else regular_font
            if is_best: cell.fill = green_fill
            cell.border = thin_border

    wb.save(output)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. INICIALIZACIÓN DE ESTADO Y SIDEBAR INTEGRA CON SHAREPOINT + EXCEL PROVEEDORES
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state: st.session_state["cotizaciones"] = []
if "monto_input" not in st.session_state: st.session_state["monto_input"] = ""
if "moneda_input" not in st.session_state: st.session_state["moneda_input"] = "CLP"

with st.sidebar:
    st.header("🔗 Conexión OneDrive / SharePoint")
    url_sharepoint = st.text_input("Enlace Excel SharePoint (Opcional)", key="url_sharepoint")
    if st.button("🔄 Recargar SharePoint"):
        st.cache_data.clear()
        st.toast("Base de SharePoint actualizada", icon="🔄")

    df_maestro, estado_sp = cargar_maestro_solpeds(url_sharepoint)
    st.caption(f"Estado SharePoint: **{estado_sp}**")

    st.divider()
    # NUEVO MÓDULO: CARGAR EXCEL DE PROVEEDORES
    st.header("📂 Cargar Excel de Proveedores")
    uploaded_prov = st.file_uploader("Adjuntar Excel/CSV Cotizaciones", type=["xlsx", "xls", "csv"], key="excel_prov")
    df_prov_excel = None
    if uploaded_prov is not None:
        try:
            df_prov_excel = pd.read_csv(uploaded_prov) if uploaded_prov.name.endswith(".csv") else pd.read_excel(uploaded_prov)
            st.success(f" Archivo activo ({len(df_prov_excel)} filas)")
        except Exception as e:
            st.error("Error al leer Excel de proveedores")

    st.divider()
    st.header("📌 Búsqueda de Solped")
    solped_ingresada = st.text_input("Ingresar N° Solped", value="10045982")
    df_solped_match = df_maestro[df_maestro["SOLPED"] == solped_ingresada.strip()]

    if not df_solped_match.empty:
        pos_seleccionada = st.selectbox("Posición (POS)", df_solped_match["POS"].tolist())
        row_solped = df_solped_match[df_solped_match["POS"] == pos_seleccionada].iloc[0]
        solped_info = {
            "SOLPED": str(row_solped["SOLPED"]), "POS": int(row_solped["POS"]),
            "SOCIEDAD": str(row_solped["SOCIEDAD"]), "CODIGO_SAP": str(row_solped["CODIGO_SAP"]),
            "DESCRIPCION": str(row_solped["DESCRIPCION"]), "UM": str(row_solped["UM"]),
            "ULTIMA_COMPRA_MONTO": float(row_solped["ULTIMA_COMPRA_MONTO"]),
            "ULTIMA_COMPRA_MONEDA": str(row_solped["ULTIMA_COMPRA_MONEDA"]),
            "PROVEEDOR_HISTORICO": str(row_solped["PROVEEDOR_HISTORICO"])
        }
        st.success("✅ Solped auto-cargada con éxito")
    else:
        st.warning("⚠️ Solped no encontrada en Maestro. Ingrese datos:")
        solped_info = {
            "SOLPED": solped_ingresada, "POS": st.number_input("POS", value=10),
            "SOCIEDAD": st.selectbox("Sociedad", ["EC01", "EC06"]), "CODIGO_SAP": st.text_input("Código SAP", value="3001892"),
            "DESCRIPCION": st.text_input("Descripción", value="MATERIAL GENÉRICO"), "UM": st.text_input("UM", value="C/U"),
            "ULTIMA_COMPRA_MONTO": st.number_input("Última Compra ($)", value=100000.0),
            "ULTIMA_COMPRA_MONEDA": "CLP", "PROVEEDOR_HISTORICO": st.text_input("Prov. Histórico", value="PROVEEDOR BASE")
        }

    # BINDING AUTOMÁTICO DESDE EXCEL DE PROVEEDORES
    if df_prov_excel is not None:
        ofertas_auto = vincular_cotizaciones_excel(df_prov_excel, solped_ingresada, solped_info, indicadores)
        if ofertas_auto:
            st.info(f"💡 Se detectaron {len(ofertas_auto)} oferta(s) para ID {solped_ingresada}")
            if st.button("🔗 Vincular Ofertas de Excel", type="primary"):
                st.session_state["cotizaciones"] = ofertas_auto
                st.toast(f"✅ {len(ofertas_auto)} cotizaciones cargadas automáticamente", icon="🎉")
                st.rerun()

    st.divider()
    if st.button("🌐 Sincronizar Divisas API"):
        st.cache_data.clear()
        indicadores = dict(obtener_indicadores_financieros())
        st.rerun()

def formatear_caja_monto():
    raw = str(st.session_state["monto_input"]).strip()
    moneda = st.session_state["moneda_input"]
    if not raw: return
    solo_numeros = re.sub(r'[^0-9.,]', '', raw)
    if not solo_numeros: return
    try:
        limpio = solo_numeros.replace(",", "") if moneda == "USD" else solo_numeros.replace(".", "").replace(",", ".")
        num = float(limpio)
        if moneda == "CLP": st.session_state["monto_input"] = f"{int(num):,}".replace(",", ".")
        elif moneda == "USD": st.session_state["monto_input"] = f"{num:,.2f}"
        elif moneda in ["EUR", "UF"]: st.session_state["monto_input"] = f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError: pass

# -----------------------------------------------------------------------------
# 5. PANEL PRINCIPAL
# -----------------------------------------------------------------------------
st.title("🛒 Consola de Compras — Enaex")
st.caption(f"🗓️ Valores del día ({indicadores['fecha']}) - API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1.5, 1.5, 1.5, 1])
col_uf.metric("UF", formato_clp(indicadores['uf']))
col_usd.metric("Dólar", formato_clp(indicadores['dolar']))
col_eur.metric("Euro", formato_clp(indicadores['euro']))

st.divider()
st.markdown(f"""
<div style="background-color: #F8FAFC; padding: 12px; border-radius: 8px; border: 1px solid #E2E8F0; margin-bottom: 15px;">
    <h4 style="margin: 0 0 8px 0; color: #0F172A;">📋 Ficha TÉCNICA REQUERIMIENTO (SOLPED: {solped_info['SOLPED']})</h4>
    <p style="margin: 0; font-size: 14px; color: #334155;">
        <b>POS:</b> {solped_info['POS']} | <b>Código SAP:</b> {solped_info['CODIGO_SAP']} | 
        <b>Descripción:</b> {solped_info['DESCRIPCION']} | <b>UM:</b> {solped_info['UM']} | <b>Sociedad:</b> {solped_info['SOCIEDAD']}<br>
        <b>Última Compra Benchmark:</b> {formato_clp(solped_info['ULTIMA_COMPRA_MONTO'])} CLP ({solped_info['PROVEEDOR_HISTORICO']})
    </p>
</div>
""", unsafe_allow_html=True)

st.subheader("➕ Carga Manual / Directa de Oferta")
def procesar_guardado():
    raw = str(st.session_state.get("monto_input", "")).strip()
    moneda = st.session_state.get("moneda_input", "CLP")
    proveedor = st.session_state.get("proveedor_input", "")
    fecha_entrega = st.session_state.get("fecha_entrega_input", datetime.date.today())
    obs = st.session_state.get("obs_input", "")

    try: monto = float(raw.replace(",", "")) if moneda == "USD" else float(raw.replace(".", "").replace(",", "."))
    except ValueError: monto = 0.0

    if not proveedor or monto <= 0:
        st.toast("⚠️ Ingrese Proveedor y Monto válido.", icon="🚨")
    else:
        monto_clp = monto * (indicadores["dolar"] if moneda == "USD" else indicadores["euro"] if moneda == "EUR" else indicadores["uf"] if moneda == "UF" else 1)
        monto_usd = monto_clp / indicadores["dolar"] if indicadores["dolar"] > 0 else 0
        ult_compra = solped_info["ULTIMA_COMPRA_MONTO"]
        var_pct = ((monto_clp - ult_compra) / ult_compra * 100) if ult_compra > 0 else 0.0

        st.session_state["cotizaciones"].append({
            "POS": solped_info["POS"], "Proveedor": proveedor, "Monto Original": monto, "Moneda": moneda,
            "Equiv. CLP ($)": round(monto_clp, 2), "Equiv. USD ($)": round(monto_usd, 2),
            "Var % vs Hist": round(var_pct, 1), "Fecha de Entrega": fecha_entrega.strftime('%d-%m-%Y'),
            "Observaciones": obs,
        })
        st.session_state["monto_input"] = ""
        st.session_state["proveedor_input"] = ""
        st.session_state["obs_input"] = ""
        st.toast(f"✅ Oferta de {proveedor} guardada.", icon="✅")

col1, col2, col3, col4 = st.columns(4)
col1.text_input("Proveedor*", key="proveedor_input")
col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"], key="moneda_input", on_change=formatear_caja_monto)
col2.text_input("Monto Original*", key="monto_input", on_change=formatear_caja_monto)
col4.date_input("Fecha de Entrega", min_value=datetime.date.today(), key="fecha_entrega_input")
st.text_area("Observaciones Técnicas", key="obs_input")
st.button("Guardar en Cuadro Comparativo", on_click=procesar_guardado)

st.divider()
st.subheader("📊 Cuadro Comparativo Homogeneizado")

if not st.session_state["cotizaciones"]:
    st.info("👆 Registra ofertas manualmente o sube un Excel en el panel lateral para vincularlas por Solped.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    df_visual = df.copy()
    df_visual["Monto Original"] = df_visual.apply(lambda r: aplicar_formato_regional(r["Monto Original"], r["Moneda"]), axis=1)
    df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(x):,}".replace(",", "."))
    df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")
    df_visual["Var % vs Hist"] = df_visual["Var % vs Hist"].apply(lambda x: f"{x:+.1f}%")
    st.dataframe(df_visual, use_container_width=True)

    fig_precio = px.bar(df, x="Proveedor", y="Equiv. CLP ($)", text_auto=',.0f', title="Comparativa de Ofertas vs Última Compra (Línea Roja)", color="Proveedor")
    fig_precio.add_hline(y=solped_info["ULTIMA_COMPRA_MONTO"], line_dash="dash", line_color="red", annotation_text="Última Compra")
    st.plotly_chart(fig_precio, use_container_width=True)

    col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
    with col_btn1:
        st.download_button("📥 Descargar Excel Corporativo", data=generar_excel_estilizado(df, solped_info, indicadores), file_name=f"cuadro_solped_{solped_info['SOLPED']}.xlsx")
    with col_btn2:
        st.download_button("📄 Descargar Reporte PDF", data=generar_pdf_ejecutivo(solped_info, st.session_state["cotizaciones"], indicadores), file_name=f"reporte_solped_{solped_info['SOLPED']}.pdf")
    with col_btn3:
        if st.button("Limpiar Cuadro", type="secondary"):
            st.session_state["cotizaciones"] = []
            st.rerun()
