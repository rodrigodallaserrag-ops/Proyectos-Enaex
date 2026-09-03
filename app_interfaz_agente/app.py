import datetime
import io
import json
import re
import ssl
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
# 1. CONEXIÓN A SHAREPOINT / ONEDRIVE + CACHE AUTOMÁTICO
# -----------------------------------------------------------------------------
@st.cache_data(ttl=300)
def cargar_maestro_solpeds(url_sharepoint=""):
    """
    Carga el Maestro de Solpeds desde OneDrive/SharePoint o genera una base de datos local de respaldo.
    """
    if url_sharepoint:
        try:
            # Transformación de enlace compartido a enlace directo de descarga
            direct_url = url_sharepoint.replace("?e=", "&download=1").replace("p=1", "download=1")
            df = pd.read_excel(direct_url, dtype={"SOLPED": str, "CODIGO_SAP": str, "POS": int})
            return df, "SharePoint Sincronizado 🟢"
        except Exception:
            pass
            
    # Base de Datos de prueba (Fallback si no hay URL activa)
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
# 2. MOTOR FINANCIERO DE MONEDAS MULTIFUENTE (FAILOVER)
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1800)  
def obtener_indicadores_financieros():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }

    # Fuente 1: mindicador.cl
    try:
        response = requests.get("https://mindicador.cl/api", headers=headers, verify=False, timeout=5)
        if response.status_code == 200:
            data = response.json()
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

    # Fuente 2: DolarApi Chile
    try:
        res_usd = requests.get("https://cl.dolarapi.com/v1/cotizaciones/usd", headers=headers, verify=False, timeout=5)
        res_eur = requests.get("https://cl.dolarapi.com/v1/cotizaciones/eur", headers=headers, verify=False, timeout=5)
        res_uf = requests.get("https://cl.dolarapi.com/v1/cotizaciones/uf", headers=headers, verify=False, timeout=5)
        
        if res_usd.status_code == 200 and res_eur.status_code == 200 and res_uf.status_code == 200:
            d_usd = res_usd.json()
            d_eur = res_eur.json()
            d_uf = res_uf.json()
            
            val_usd = float(d_usd.get("valor") or d_usd.get("compra") or d_usd.get("venta"))
            val_eur = float(d_eur.get("valor") or d_eur.get("compra") or d_eur.get("venta"))
            val_uf = float(d_uf.get("valor") or d_uf.get("compra") or d_uf.get("venta"))

            return {
                "dolar": val_usd,
                "euro": val_eur,
                "uf": val_uf,
                "fecha": datetime.date.today().strftime("%Y-%m-%d"),
                "estado": "Online (DolarApi) 🟢",
            }
    except Exception:
        pass

    # Fuente 3: Gael API
    try:
        res = requests.get("https://api.gael.cl/general/public/monedas", headers=headers, verify=False, timeout=5)
        if res.status_code == 200:
            items = res.json()
            m_dict = {}
            for item in items:
                if 'Codigo' in item and 'Valor' in item:
                    val_clean = str(item['Valor']).replace('.', '').replace(',', '.')
                    m_dict[item['Codigo']] = float(val_clean)
            
            if 'UF' in m_dict and 'USD' in m_dict and 'EUR' in m_dict:
                return {
                    "dolar": m_dict['USD'],
                    "euro": m_dict['EUR'],
                    "uf": m_dict['UF'],
                    "fecha": datetime.date.today().strftime("%Y-%m-%d"),
                    "estado": "Online (Gael API) 🟢",
                }
    except Exception:
        pass

    # Respaldo si todo lo demás falla
    return {
        "dolar": 938.0,
        "euro": 1020.0,
        "uf": 40875.0,
        "fecha": "Valores Estimados",
        "estado": "Offline (Red Estricta) 🛡️",
    }

indicadores_cache = obtener_indicadores_financieros()
indicadores = dict(indicadores_cache)

def formato_clp(valor):
    return f"${int(valor):,}".replace(",", ".")

def aplicar_formato_regional(monto, moneda):
    if moneda == "CLP":
        return f"$ {int(monto):,}".replace(",", ".")
    elif moneda == "USD":
        return f"$ {monto:,.2f}"
    elif moneda == "EUR":
        return f"€ {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    elif moneda == "UF":
        return f"UF {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(monto)

# -----------------------------------------------------------------------------
# 3. GENERADOR DE REPORTES PDF EJECUTIVOS
# -----------------------------------------------------------------------------
def generar_pdf_ejecutivo(solped_info, cotizaciones, datos_indicadores):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
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
        [
            Paragraph(f"<b>N° Solped:</b> {solped_info['SOLPED']}", normal_style),
            Paragraph(f"<b>POS:</b> {solped_info['POS']}", normal_style),
            Paragraph(f"<b>Sociedad:</b> {solped_info['SOCIEDAD']}", normal_style)
        ],
        [
            Paragraph(f"<b>Código SAP:</b> {solped_info['CODIGO_SAP']}", normal_style),
            Paragraph(f"<b>Descripción:</b> {solped_info['DESCRIPCION']}", normal_style),
            Paragraph(f"<b>UM:</b> {solped_info['UM']}", normal_style)
        ],
        [
            Paragraph(f"<b>Última Compra:</b> {formato_clp(solped_info['ULTIMA_COMPRA_MONTO'])}", normal_style),
            Paragraph(f"<b>Prov. Histórico:</b> {solped_info['PROVEEDOR_HISTORICO']}", normal_style),
            Paragraph(f"<b>Fecha Emisión:</b> {datetime.date.today().strftime('%d-%m-%Y')}", normal_style)
        ]
    ]
    t_meta = Table(meta_data, colWidths=[180, 180, 180])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('PADDING', (0,0), (-1,-1), 5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1'))
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    table_data = [[
        Paragraph("<b>Proveedor</b>", header_table_style),
        Paragraph("<b>Monto Orig.</b>", header_table_style),
        Paragraph("<b>Mon.</b>", header_table_style),
        Paragraph("<b>Equiv. CLP ($)</b>", header_table_style),
        Paragraph("<b>Var. % Hist.</b>", header_table_style),
        Paragraph("<b>Plazo Entrega</b>", header_table_style),
        Paragraph("<b>Observaciones</b>", header_table_style)
    ]]

    min_clp = min([c["Equiv. CLP ($)"] for c in cotizaciones]) if cotizaciones else 0

    for c in cotizaciones:
        monto_orig_fmt = aplicar_formato_regional(c["Monto Original"], c["Moneda"])
        clp_fmt = f"$ {int(c['Equiv. CLP ($)']):,}".replace(",", ".")
        
        prov_text = f"<b>{c['Proveedor']}</b>"
        if len(cotizaciones) > 1 and c["Equiv. CLP ($)"] == min_clp:
            prov_text += "<br/><font color='#16A34A'><b>★ Mejor Oferta</b></font>"

        var_pct = c.get("Var % vs Hist", 0)
        color_var = "#16A34A" if var_pct <= 0 else "#DC2626"
        var_text = f"<font color='{color_var}'><b>{var_pct:+.1f}%</b></font>"

        table_data.append([
            Paragraph(prov_text, cell_style),
            Paragraph(monto_orig_fmt, cell_style),
            Paragraph(c["Moneda"], cell_style),
            Paragraph(clp_fmt, cell_style),
            Paragraph(var_text, cell_style),
            Paragraph(c["Fecha de Entrega"], cell_style),
            Paragraph(c.get("Observaciones", "-") or "-", cell_style)
        ])

    t_quotes = Table(table_data, colWidths=[100, 65, 30, 75, 60, 90, 120])
    t_quotes.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_quotes)
    story.append(Spacer(1, 10))

    max_monto = max([c["Equiv. CLP ($)"] for c in cotizaciones]) if cotizaciones else 0
    if max_monto > 1000000:
        warn_p = Paragraph(
            f"<b>⚠️ Nota de Control Financiero:</b> El requerimiento supera $1.000.000 CLP "
            f"(Máximo detectado: {formato_clp(max_monto)} CLP). Requiere validación según matriz de firmas.",
            normal_style
        )
        story.append(warn_p)

    story.append(Spacer(1, 20))
    sig_data = [[
        Paragraph("___________________________________<br/><b>Elaborado por:</b> Analista de Compras", cell_style),
        Paragraph("___________________________________<br/><b>Aprobado por:</b> Jefatura de Abastecimiento", cell_style)
    ]]
    t_sig = Table(sig_data, colWidths=[270, 270])
    t_sig.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(t_sig)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# 4. GENERADOR DE EXCEL CORPORATIVO CON HISTÓRICO
# -----------------------------------------------------------------------------
def generar_excel_estilizado(df, solped_info, datos_indicadores):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cuadro Comparativo"
    ws.views.sheetView[0].showGridLines = True

    navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=15, bold=True, color="0F172A")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    meta_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws.cell(row=1, column=1, value="ENAEX — Consola de Compras").font = title_font
    ws.cell(row=2, column=1, value=f"Cuadro Comparativo - Solped N° {solped_info['SOLPED']} (POS: {solped_info['POS']})").font = bold_font

    meta_data = [
        [f"Código SAP: {solped_info['CODIGO_SAP']}", f"Descripción: {solped_info['DESCRIPCION']}", f"UM: {solped_info['UM']}"],
        [f"Sociedad: {solped_info['SOCIEDAD']}", f"Última Compra: ${solped_info['ULTIMA_COMPRA_MONTO']:,} CLP", f"Prov. Histórico: {solped_info['PROVEEDOR_HISTORICO']}"]
    ]
    for r_idx, row_data in enumerate(meta_data, start=4):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.fill = meta_fill
            cell.border = thin_border

    headers = ["POS", "Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Var. % Hist.", "Fecha Entrega", "Observaciones"]
    start_row = 7
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    min_clp = df["Equiv. CLP ($)"].min() if not df.empty else 0

    for r_offset, (_, row) in enumerate(df.iterrows()):
        current_row = start_row + 1 + r_offset
        is_best = (len(df) > 1 and row["Equiv. CLP ($)"] == min_clp)
        
        prov_text = str(row["Proveedor"]) + (" ★ (Mejor Oferta)" if is_best else "")
        
        c0 = ws.cell(row=current_row, column=1, value=solped_info["POS"])
        c1 = ws.cell(row=current_row, column=2, value=prov_text)
        c2 = ws.cell(row=current_row, column=3, value=row["Monto Original"])
        c3 = ws.cell(row=current_row, column=4, value=row["Moneda"])
        c4 = ws.cell(row=current_row, column=5, value=row["Equiv. CLP ($)"])
        c5 = ws.cell(row=current_row, column=6, value=row["Equiv. USD ($)"])
        c6 = ws.cell(row=current_row, column=7, value=f"{row.get('Var % vs Hist', 0):+.1f}%")
        c7 = ws.cell(row=current_row, column=8, value=str(row["Fecha de Entrega"]))
        c8 = ws.cell(row=current_row, column=9, value=str(row.get("Observaciones", "")))

        c2.number_format = '$ #,##0.00' if row["Moneda"] in ["USD", "EUR", "UF"] else '$ #,##0'
        c4.number_format = '$ #,##0'
        c5.number_format = '$ #,##0.00'

        for c in [c0, c1, c2, c3, c4, c5, c6, c7, c8]:
            c.font = bold_font if is_best else regular_font
            if is_best:
                c.fill = green_fill
            c.border = thin_border

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 5. INICIALIZACIÓN DE ESTADO Y SIDEBAR INTEGRA CON SHAREPOINT
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

if "monto_input" not in st.session_state:
    st.session_state["monto_input"] = ""

if "moneda_input" not in st.session_state:
    st.session_state["moneda_input"] = "CLP"

with st.sidebar:
    st.header("🔗 Conexión OneDrive / SharePoint")
    url_sharepoint = st.text_input("Enlace Excel SharePoint (Opcional)", key="url_sharepoint")
    
    col_sp1, col_sp2 = st.columns([1, 1])
    if col_sp1.button("🔄 Recargar Datos"):
        st.cache_data.clear()
        st.toast("Base de datos y caché de monedas actualizados", icon="🔄")

    df_maestro, estado_sp = cargar_maestro_solpeds(url_sharepoint)
    st.caption(f"Estado repositorio: **{estado_sp}**")

    st.divider()
    st.header("📌 Búsqueda de Solped")
    
    solped_ingresada = st.text_input("Ingresar N° Solped", value="10045982")
    
    # Filtrar datos de la Solped ingresada
    df_solped_match = df_maestro[df_maestro["SOLPED"] == solped_ingresada.strip()]

    if not df_solped_match.empty:
        posiciones_disponibles = df_solped_match["POS"].tolist()
        pos_seleccionada = st.selectbox("Posición (POS)", opciones := posiciones_disponibles)
        
        row_solped = df_solped_match[df_solped_match["POS"] == pos_seleccionada].iloc[0]
        
        # Datos extraídos automáticamente
        solped_info = {
            "SOLPED": str(row_solped["SOLPED"]),
            "POS": int(row_solped["POS"]),
            "SOCIEDAD": str(row_solped["SOCIEDAD"]),
            "CODIGO_SAP": str(row_solped["CODIGO_SAP"]),
            "DESCRIPCION": str(row_solped["DESCRIPCION"]),
            "UM": str(row_solped["UM"]),
            "ULTIMA_COMPRA_MONTO": float(row_solped["ULTIMA_COMPRA_MONTO"]),
            "ULTIMA_COMPRA_MONEDA": str(row_solped["ULTIMA_COMPRA_MONEDA"]),
            "PROVEEDOR_HISTORICO": str(row_solped["PROVEEDOR_HISTORICO"])
        }
        st.success("✅ Solped auto-cargada con éxito")
    else:
        st.warning("⚠️ Solped no encontrada en SharePoint. Ingrese valores manuales:")
        solped_info = {
            "SOLPED": solped_ingresada,
            "POS": st.number_input("POS", value=10),
            "SOCIEDAD": st.selectbox("Sociedad", ["EC01", "EC06"]),
            "CODIGO_SAP": st.text_input("Código SAP", value="3001892"),
            "DESCRIPCION": st.text_input("Descripción", value="MATERIAL GENÉRICO"),
            "UM": st.text_input("UM", value="C/U"),
            "ULTIMA_COMPRA_MONTO": st.number_input("Última Compra ($)", value=100000.0),
            "ULTIMA_COMPRA_MONEDA": "CLP",
            "PROVEEDOR_HISTORICO": st.text_input("Prov. Histórico", value="PROVEEDOR BASE")
        }

    # Edición manual de respaldo
    if "Offline" in indicadores["estado"]:
        st.divider()
        st.warning("⚠️ Ajuste manual de monedas (API Offline):")
        
        val_uf = st.text_input("Valor UF", value=str(indicadores["uf"]))
        val_usd = st.text_input("Valor Dólar", value=str(indicadores["dolar"]))
        val_eur = st.text_input("Valor Euro", value=str(indicadores["euro"]))
        
        try:
            indicadores["uf"] = float(val_uf.replace(",", "."))
        except ValueError:
            pass
            
        try:
            indicadores["dolar"] = float(val_usd.replace(",", "."))
        except ValueError:
            pass
            
        try:
            indicadores["euro"] = float(val_eur.replace(",", "."))
        except ValueError:
            pass

def formatear_caja_monto():
    raw = str(st.session_state["monto_input"]).strip()
    moneda = st.session_state["moneda_input"]
    if not raw: return
    solo_numeros = re.sub(r'[^0-9.,]', '', raw)
    if not solo_numeros:
        st.session_state["monto_input"] = ""
        return
    try:
        limpio = solo_numeros.replace(",", "") if moneda == "USD" else solo_numeros.replace(".", "").replace(",", ".")
        num = float(limpio)
        if moneda == "CLP":
            st.session_state["monto_input"] = f"{int(num):,}".replace(",", ".")
        elif moneda == "USD":
            st.session_state["monto_input"] = f"{num:,.2f}"
        elif moneda in ["EUR", "UF"]:
            st.session_state["monto_input"] = f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        st.session_state["monto_input"] = ""

st.title("🛒 Consola de Compras — Enaex")

# -----------------------------------------------------------------------------
# 6. PANEL DE CONTROL E INFORMACIÓN MAESTRA
# -----------------------------------------------------------------------------
st.caption(f"🗓️ Valores del día ({indicadores['fecha']}) - API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1.5, 1.5, 1.5, 1])
col_uf.metric("UF", formato_clp(indicadores['uf']))
col_usd.metric("Dólar", formato_clp(indicadores['dolar']))
col_eur.metric("Euro", formato_clp(indicadores['euro']))

st.divider()

# Tarjeta Ficha Técnica Solped
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

# -----------------------------------------------------------------------------
# 7. CARGA MANUAL DE OFERTAS
# -----------------------------------------------------------------------------
st.subheader("➕ Carga Manual de Oferta")

def procesar_guardado():
    raw = str(st.session_state.get("monto_input", "")).strip()
    moneda = st.session_state.get("moneda_input", "CLP")
    proveedor = st.session_state.get("proveedor_input", "")
    fecha_entrega = st.session_state.get("fecha_entrega_input", datetime.date.today())
    obs = st.session_state.get("obs_input", "")
    
    dolar_actual = indicadores["dolar"]
    euro_actual = indicadores["euro"]
    uf_actual = indicadores["uf"]

    try:
        monto = float(raw.replace(",", "")) if moneda == "USD" else float(raw.replace(".", "").replace(",", "."))
    except ValueError:
        monto = 0.0

    if not proveedor or monto <= 0:
        st.toast("⚠️ Ingrese un Proveedor y Monto válido.", icon="🚨")
    else:
        monto_clp = monto
        if moneda == "USD": monto_clp = monto * dolar_actual
        elif moneda == "EUR": monto_clp = monto * euro_actual
        elif moneda == "UF": monto_clp = monto * uf_actual

        monto_usd = monto_clp / dolar_actual if dolar_actual > 0 else 0
        
        # Cálculo Variación % vs Última Compra
        ult_compra = solped_info["ULTIMA_COMPRA_MONTO"]
        var_pct = ((monto_clp - ult_compra) / ult_compra * 100) if ult_compra > 0 else 0.0

        dias_diferencia = (fecha_entrega - datetime.date.today()).days
        plazo_final = f"{fecha_entrega.strftime('%d-%m-%Y')} ({dias_diferencia} días)"

        st.session_state["cotizaciones"].append({
            "POS": solped_info["POS"],
            "Proveedor": proveedor,
            "Monto Original": monto, 
            "Moneda": moneda,
            "Equiv. CLP ($)": round(monto_clp, 2),
            "Equiv. USD ($)": round(monto_usd, 2),
            "Var % vs Hist": round(var_pct, 1),
            "Fecha de Entrega": plazo_final,
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

# -----------------------------------------------------------------------------
# 8. CUADRO COMPARATIVO CON MÉTRICAS HISTÓRICAS
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo Homogeneizado")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=["POS", "Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Var % vs Hist", "Fecha de Entrega", "Observaciones"]
    )
    st.dataframe(df_empty, use_container_width=True)
    st.info("👆 Registra cotizaciones para construir el cuadro comparativo.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    
    df_visual = df.copy()
    df_visual["Monto Original"] = df_visual.apply(lambda r: aplicar_formato_regional(r["Monto Original"], r["Moneda"]), axis=1)
    df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(x):,}".replace(",", "."))
    df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")
    df_visual["Var % vs Hist"] = df_visual["Var % vs Hist"].apply(lambda x: f"{x:+.1f}%")

    st.dataframe(df_visual, use_container_width=True)

    # Métricas de Variación vs Última Compra
    st.markdown("### 📈 Métricas de Adjudicación & Benchmark")
    
    monto_min = df["Equiv. CLP ($)"].min()
    monto_max = df["Equiv. CLP ($)"].max()
    prov_min = df.loc[df["Equiv. CLP ($)"] == monto_min, "Proveedor"].values[0]
    
    ult_compra_monto = solped_info["ULTIMA_COMPRA_MONTO"]
    ahorro_vs_hist = ult_compra_monto - monto_min
    pct_vs_hist = ((monto_min - ult_compra_monto) / ult_compra_monto * 100) if ult_compra_monto > 0 else 0

    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    
    with kpi_col1:
        st.metric(
            label="🏆 Oferta Adjudicable",
            value=prov_min,
            delta=formato_clp(monto_min),
            delta_color="normal"
        )
    with kpi_col2:
        st.metric(
            label="📉 Variación vs. Última Compra",
            value=f"{pct_vs_hist:+.1f}%",
            delta=f"{formato_clp(abs(ahorro_vs_hist))} {'Ahorro' if ahorro_vs_hist >= 0 else 'Sobrecosto'}",
            delta_color="inverse" if pct_vs_hist > 0 else "normal"
        )
    with kpi_col3:
        st.metric(
            label="📌 Benchmark Histórico",
            value=formato_clp(ult_compra_monto),
            delta=f"Prov: {solped_info['PROVEEDOR_HISTORICO']}",
            delta_color="off"
        )

    # Gráfico de barras
    df_chart = df.copy()
    fig_precio = px.bar(
        df_chart,
        x="Proveedor",
        y="Equiv. CLP ($)",
        text_auto=',.0f',
        title="Comparativa de Ofertas vs Última Compra (Línea Roja)",
        color="Proveedor"
    )
    fig_precio.add_hline(y=ult_compra_monto, line_dash="dash", line_color="red", annotation_text="Última Compra")
    st.plotly_chart(fig_precio, use_container_width=True)

    # Gestión y Descargas
    st.markdown("#### 🛠️ Gestión de Filas")
    for i, c in enumerate(st.session_state["cotizaciones"]):
        c_col1, c_col2 = st.columns([5, 1])
        c_col1.write(f"• **{c['Proveedor']}**: {formato_clp(c['Equiv. CLP ($)'])} CLP (Var: {c['Var % vs Hist']:+.1f}%)")
        if c_col2.button("Eliminar", key=f"del_{i}"):
            st.session_state["cotizaciones"].pop(i)
            st.rerun()

    st.divider()

    col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
    with col_btn1:
        excel_bytes = generar_excel_estilizado(df, solped_info, indicadores)
        st.download_button(
            label="📥 Descargar Excel Corporativo",
            data=excel_bytes,
            file_name=f"cuadro_comparativo_solped_{solped_info['SOLPED']}_pos{solped_info['POS']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_btn2:
        pdf_bytes = generar_pdf_ejecutivo(solped_info, st.session_state["cotizaciones"], indicadores)
        st.download_button(
            label="📄 Descargar Reporte PDF",
            data=pdf_bytes,
            file_name=f"reporte_ejecutivo_solped_{solped_info['SOLPED']}.pdf",
            mime="application/pdf"
        )
    with col_btn3:
        if st.button("Limpiar Cuadro", type="secondary"):
            st.session_state["cotizaciones"] = []
            st.rerun()
