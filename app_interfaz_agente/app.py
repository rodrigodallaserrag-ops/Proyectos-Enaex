import datetime
import io
import json
import re
import ssl
import urllib.request
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
# 1. MOTOR FINANCIERO: EVASIÓN SSL CORPORATIVA + SISTEMA CASCADA
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)  
def obtener_indicadores_financieros():
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        
        req = urllib.request.Request(
            "https://mindicador.cl/api", 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req, timeout=8, context=ctx) as response:
            data = json.loads(response.read().decode('utf-8'))
            if "dolar" in data:
                return {
                    "dolar": float(data["dolar"]["valor"]),
                    "euro": float(data["euro"]["valor"]),
                    "uf": float(data["uf"]["valor"]),
                    "fecha": data["dolar"]["fecha"][:10],
                    "estado": "Online (Proxy Local) 🟢",
                }
    except Exception:
        pass
        
    urls_intento = [
        "https://mindicador.cl/api",
        "http://mindicador.cl/api",
        "https://api.allorigins.win/raw?url=https://mindicador.cl/api"
    ]
    
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    for url in urls_intento:
        try:
            with requests.Session() as s:
                s.trust_env = True
                response = s.get(url, timeout=8, headers=headers, verify=False)
                response.raise_for_status()
                data = response.json()
                
                if "dolar" in data and "valor" in data["dolar"]:
                    return {
                        "dolar": float(data["dolar"]["valor"]),
                        "euro": float(data["euro"]["valor"]),
                        "uf": float(data["uf"]["valor"]),
                        "fecha": data["dolar"]["fecha"][:10],
                        "estado": "Online (Ruta Alterna) 🟢",
                    }
        except Exception:
            continue 

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
# 2. GENERADOR DE REPORTES PDF EJECUTIVOS
# -----------------------------------------------------------------------------
def generar_pdf_ejecutivo(solped, material, sociedad, cotizaciones, datos_indicadores):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20,
        textColor=colors.HexColor('#0F172A'), spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'SubTitleStyle', parent=styles['Normal'], fontSize=10, leading=12,
        textColor=colors.HexColor('#475569'), spaceAfter=14
    )
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontSize=9, leading=12)
    header_table_style = ParagraphStyle(
        'HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10,
        textColor=colors.whitesmoke, fontName='Helvetica-Bold'
    )
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10)

    story.append(Paragraph("<b>ENAEX — Consola de Compras</b>", title_style))
    story.append(Paragraph(f"Reporte Ejecutivo de Adjudicación — Solped N° <b>{solped}</b>", subtitle_style))

    meta_data = [
        [
            Paragraph(f"<b>N° Solped:</b> {solped}", normal_style),
            Paragraph(f"<b>Código Material:</b> {material}", normal_style),
            Paragraph(f"<b>Sociedad:</b> {sociedad}", normal_style)
        ],
        [
            Paragraph(f"<b>Fecha Emisión:</b> {datetime.date.today().strftime('%d-%m-%Y')}", normal_style),
            Paragraph(f"<b>Dólar Ref.:</b> ${datos_indicadores['dolar']:,.2f}", normal_style),
            Paragraph(f"<b>UF Ref.:</b> ${datos_indicadores['uf']:,.2f}", normal_style)
        ]
    ]
    t_meta = Table(meta_data, colWidths=[180, 180, 180])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('PADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LINEBELOW', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E1'))
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 14))

    table_data = [[
        Paragraph("<b>Proveedor</b>", header_table_style),
        Paragraph("<b>Monto Orig.</b>", header_table_style),
        Paragraph("<b>Mon.</b>", header_table_style),
        Paragraph("<b>Equiv. CLP ($)</b>", header_table_style),
        Paragraph("<b>Equiv. USD ($)</b>", header_table_style),
        Paragraph("<b>Plazo Entrega</b>", header_table_style),
        Paragraph("<b>Observaciones</b>", header_table_style)
    ]]

    min_clp = min([c["Equiv. CLP ($)"] for c in cotizaciones]) if cotizaciones else 0

    for c in cotizaciones:
        monto_orig_fmt = aplicar_formato_regional(c["Monto Original"], c["Moneda"])
        clp_fmt = f"$ {int(c['Equiv. CLP ($)']):,}".replace(",", ".")
        usd_fmt = f"$ {c['Equiv. USD ($)']:,.2f}"
        
        prov_text = f"<b>{c['Proveedor']}</b>"
        if len(cotizaciones) > 1 and c["Equiv. CLP ($)"] == min_clp:
            prov_text += "<br/><font color='#16A34A'><b>★ Mejor Oferta</b></font>"

        table_data.append([
            Paragraph(prov_text, cell_style),
            Paragraph(monto_orig_fmt, cell_style),
            Paragraph(c["Moneda"], cell_style),
            Paragraph(clp_fmt, cell_style),
            Paragraph(usd_fmt, cell_style),
            Paragraph(c["Fecha de Entrega"], cell_style),
            Paragraph(c.get("Observaciones", "-") or "-", cell_style)
        ])

    t_quotes = Table(table_data, colWidths=[105, 65, 30, 75, 70, 85, 110])
    t_quotes.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_quotes)
    story.append(Spacer(1, 15))

    max_monto = max([c["Equiv. CLP ($)"] for c in cotizaciones]) if cotizaciones else 0
    if max_monto > 1000000:
        warn_p = Paragraph(
            f"<b>⚠️ Nota de Control Financiero:</b> El requerimiento supera $1.000.000 CLP "
            f"(Máximo detectado: {formato_clp(max_monto)} CLP). Requiere validación de acuerdo a matriz de firmas vigente.",
            normal_style
        )
        story.append(warn_p)
        story.append(Spacer(1, 15))

    story.append(Spacer(1, 25))
    sig_data = [
        [
            Paragraph("___________________________________<br/><b>Elaborado por:</b> Analista de Compras", cell_style),
            Paragraph("___________________________________<br/><b>Aprobado por:</b> Jefatura de Abastecimiento", cell_style)
        ]
    ]
    t_sig = Table(sig_data, colWidths=[270, 270])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_sig)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# -----------------------------------------------------------------------------
# 3. GENERADOR DE EXCEL CORPORATIVO ESTILIZADO
# -----------------------------------------------------------------------------
def generar_excel_estilizado(df, solped, material, sociedad, datos_indicadores):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cuadro Comparativo"
    ws.views.sheetView[0].showGridLines = True

    # Paleta de colores y fuentes Enaex
    navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="0F172A")
    sub_font = Font(name="Calibri", size=10, italic=True, color="475569")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    meta_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. Título principal
    ws.cell(row=1, column=1, value="ENAEX — Consola de Compras").font = title_font
    ws.cell(row=2, column=1, value=f"Cuadro Comparativo de Cotizaciones — Solped N° {solped}").font = sub_font

    # 2. Bloque de Metadatos
    meta_data = [
        [f"N° Solped: {solped}", f"Código Material: {material}", f"Sociedad: {sociedad}"],
        [f"Fecha Emisión: {datetime.date.today().strftime('%d-%m-%Y')}", f"Dólar Ref.: ${datos_indicadores['dolar']:,.2f}", f"UF Ref.: ${datos_indicadores['uf']:,.2f}"]
    ]
    
    for r_idx, row_data in enumerate(meta_data, start=4):
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = bold_font
            cell.fill = meta_fill
            cell.border = thin_border

    # 3. Encabezados de Tabla
    headers = ["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Fecha de Entrega", "Observaciones"]
    start_row = 7
    
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center" if c_idx in [2,3,4,5,6] else "left", vertical="center")

    # 4. Filas de Datos y Resaltado
    min_clp = df["Equiv. CLP ($)"].min() if not df.empty else 0

    for r_offset, (_, row) in enumerate(df.iterrows()):
        current_row = start_row + 1 + r_offset
        is_best = (len(df) > 1 and row["Equiv. CLP ($)"] == min_clp)
        
        prov_text = str(row["Proveedor"]) + (" ★ (Mejor Oferta)" if is_best else "")
        
        c1 = ws.cell(row=current_row, column=1, value=prov_text)
        c2 = ws.cell(row=current_row, column=2, value=row["Monto Original"])
        c3 = ws.cell(row=current_row, column=3, value=row["Moneda"])
        c4 = ws.cell(row=current_row, column=4, value=row["Equiv. CLP ($)"])
        c5 = ws.cell(row=current_row, column=5, value=row["Equiv. USD ($)"])
        c6 = ws.cell(row=current_row, column=6, value=str(row["Fecha de Entrega"]))
        c7 = ws.cell(row=current_row, column=7, value=str(row.get("Observaciones", "")))

        # Formatos numéricos contables
        c2.number_format = '$ #,##0.00' if row["Moneda"] in ["USD", "EUR", "UF"] else '$ #,##0'
        c3.alignment = Alignment(horizontal="center")
        c4.number_format = '$ #,##0'
        c5.number_format = '$ #,##0.00'
        c6.alignment = Alignment(horizontal="center")

        # Aplicar colores y bordes
        for c in [c1, c2, c3, c4, c5, c6, c7]:
            c.font = bold_font if is_best else regular_font
            if is_best:
                c.fill = green_fill
            c.border = thin_border

    # 5. Ajuste automático de ancho de columnas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format and '$' in cell.number_format and isinstance(cell.value, (int, float)):
                val = f"$ {cell.value:,.2f}"
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. INICIALIZAR ESTADO Y BARRA LATERAL
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

if "monto_input" not in st.session_state:
    st.session_state["monto_input"] = ""

if "moneda_input" not in st.session_state:
    st.session_state["moneda_input"] = "CLP"

with st.sidebar:
    st.header("📌 Datos Solped")
    solped = st.text_input("N° Solped", value="10045982")
    material = st.text_input("Código Material", value="3001892")
    sociedad = st.selectbox("Sociedad", ["EC01", "EC06"])
    
    if "Offline" in indicadores["estado"]:
        st.divider()
        st.warning("⚠️ **Red Corporativa Bloqueada**\nAjusta los valores manualmente para continuar:")
        indicadores["uf"] = st.number_input("Valor UF de hoy", value=float(indicadores["uf"]), key="uf_manual")
        indicadores["dolar"] = st.number_input("Valor Dólar de hoy", value=float(indicadores["dolar"]), key="dolar_manual")
        indicadores["euro"] = st.number_input("Valor Euro de hoy", value=float(indicadores["euro"]), key="euro_manual")
        indicadores["fecha"] = datetime.date.today().strftime("%d-%m-%Y (Manual)")

def formatear_caja_monto():
    raw = str(st.session_state["monto_input"]).strip()
    moneda = st.session_state["moneda_input"]
    if not raw: return
    solo_numeros = re.sub(r'[^0-9.,]', '', raw)
    if not solo_numeros:
        st.session_state["monto_input"] = ""
        return
    try:
        if moneda == "USD":
            limpio = solo_numeros.replace(",", "")
        else:
            limpio = solo_numeros.replace(".", "").replace(",", ".")
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
# 5. PANEL CENTRAL DE MONEDAS
# -----------------------------------------------------------------------------
st.caption(f"🗓️ Valores del día ({indicadores['fecha']}) - Estado API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1.5, 1.5, 1.5, 1])
col_uf.metric("UF", formato_clp(indicadores['uf']))
col_usd.metric("Dólar", formato_clp(indicadores['dolar']))
col_eur.metric("Euro", formato_clp(indicadores['euro']))

st.divider()

# -----------------------------------------------------------------------------
# 6. INGRESO DE COTIZACIONES
# -----------------------------------------------------------------------------
st.subheader("➕ Carga Manual de Oferta")

def procesar_guardado():
    raw = str(st.session_state.get("monto_input", "")).strip()
    moneda = st.session_state.get("moneda_input", "CLP")
    proveedor = st.session_state.get("proveedor_input", "")
    fecha_entrega = st.session_state.get("fecha_entrega_input", datetime.date.today())
    obs = st.session_state.get("obs_input", "")
    
    uf_actual = st.session_state.get("uf_manual", indicadores["uf"])
    dolar_actual = st.session_state.get("dolar_manual", indicadores["dolar"])
    euro_actual = st.session_state.get("euro_manual", indicadores["euro"])
    
    try:
        if moneda == "USD":
            monto = float(raw.replace(",", ""))
        else:
            monto = float(raw.replace(".", "").replace(",", "."))
    except ValueError:
        monto = 0.0

    if not proveedor or monto <= 0:
        st.toast("⚠️ Debes ingresar un Proveedor y un Monto numérico válido.", icon="🚨")
    else:
        monto_clp = monto
        if moneda == "USD":
            monto_clp = monto * dolar_actual
        elif moneda == "EUR":
            monto_clp = monto * euro_actual
        elif moneda == "UF":
            monto_clp = monto * uf_actual

        monto_usd = monto_clp / dolar_actual if dolar_actual > 0 else 0
        
        hoy = datetime.date.today()
        dias_diferencia = (fecha_entrega - hoy).days
        fecha_formateada = fecha_entrega.strftime("%d-%m-%Y")
        
        texto_dias = "día" if dias_diferencia == 1 else "días"
        plazo_final = f"{fecha_formateada} ({dias_diferencia} {texto_dias})"

        st.session_state["cotizaciones"].append({
            "Proveedor": proveedor,
            "Monto Original": monto, 
            "Moneda": moneda,
            "Equiv. CLP ($)": round(monto_clp, 2),
            "Equiv. USD ($)": round(monto_usd, 2),
            "Fecha de Entrega": plazo_final,
            "Observaciones": obs,
        })
        
        st.session_state["monto_input"] = ""
        st.session_state["proveedor_input"] = ""
        st.session_state["obs_input"] = ""
        st.toast(f"✅ Oferta de {proveedor} ingresada correctamente.", icon="✅")

col1, col2, col3, col4 = st.columns(4)

col1.text_input("Proveedor*", key="proveedor_input")
col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"], key="moneda_input", on_change=formatear_caja_monto)
col2.text_input("Monto Original*", placeholder="Solo números (Ej: 190000)", key="monto_input", on_change=formatear_caja_monto)
col4.date_input("Fecha de Entrega", min_value=datetime.date.today(), key="fecha_entrega_input")

st.text_area("Observaciones Técnicas", key="obs_input")
st.button("Guardar en Cuadro Comparativo", on_click=procesar_guardado)

# -----------------------------------------------------------------------------
# 7. CUADRO COMPARATIVO Y EXPORTACIÓN
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo (Homogeneizado)")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Fecha de Entrega", "Observaciones"]
    )
    st.dataframe(df_empty, use_container_width=True)
    st.info("👆 Agrega ofertas arriba para visualizar el cuadro comparativo.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    
    df_visual = df.copy()
    df_visual["Monto Original"] = df_visual.apply(lambda fila: aplicar_formato_regional(fila["Monto Original"], fila["Moneda"]), axis=1)
    df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(x):,}".replace(",", "."))
    df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")

    st.dataframe(df_visual, use_container_width=True)

    # -------------------------------------------------------------------------
    # ANÁLISIS VISUAL Y MÉTRICAS DE AHORRO
    # -------------------------------------------------------------------------
    st.markdown("### 📈 Análisis Visual y Métricas de Ahorro")
    
    if len(st.session_state["cotizaciones"]) < 2:
        st.info("💡 Ingresa al menos **2 cotizaciones** para habilitar el gráfico comparativo y el análisis de ahorro.")
    else:
        monto_min = df["Equiv. CLP ($)"].min()
        monto_max = df["Equiv. CLP ($)"].max()
        monto_prom = df["Equiv. CLP ($)"].mean()
        
        prov_min = df.loc[df["Equiv. CLP ($)"] == monto_min, "Proveedor"].values[0]
        prov_max = df.loc[df["Equiv. CLP ($)"] == monto_max, "Proveedor"].values[0]
        
        ahorro_vs_max = monto_max - monto_min
        pct_vs_max = (ahorro_vs_max / monto_max * 100) if monto_max > 0 else 0
        
        ahorro_vs_prom = monto_prom - monto_min
        pct_vs_prom = (ahorro_vs_prom / monto_prom * 100) if monto_prom > 0 else 0

        kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
        
        with kpi_col1:
            st.metric(
                label="🏆 Oferta Recomendada",
                value=prov_min,
                delta=formato_clp(monto_min),
                delta_color="normal"
            )
        with kpi_col2:
            st.metric(
                label="💰 Ahorro Máximo (vs. Menos Económica)",
                value=formato_clp(ahorro_vs_max),
                delta=f"-{pct_vs_max:.1f}% vs {prov_max}",
                delta_color="normal"
            )
        with kpi_col3:
            st.metric(
                label="📊 Ahorro vs. Promedio del Mercado",
                value=formato_clp(ahorro_vs_prom),
                delta=f"-{pct_vs_prom:.1f}% vs Promedio",
                delta_color="normal"
            )

        df_chart = df.copy()
        
        def extraer_dias(cadena):
            match = re.search(r'\((\d+)\s+días?\)', str(cadena))
            return int(match.group(1)) if match else 0
            
        df_chart["Días de Entrega"] = df_chart["Fecha de Entrega"].apply(extraer_dias)
        
        def asignar_categoria_color(val):
            if val == monto_min:
                return "Mejor Opción (Mínimo)"
            elif val == monto_max:
                return "Mayor Precio"
            return "Opción Intermedia"
            
        df_chart["Evaluación"] = df_chart["Equiv. CLP ($)"].apply(asignar_categoria_color)

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            fig_precio = px.bar(
                df_chart,
                x="Proveedor",
                y="Equiv. CLP ($)",
                color="Evaluación",
                text_auto=',.0f',
                title="Comparativo de Precios Homogeneizados (CLP)",
                color_discrete_map={
                    "Mejor Opción (Mínimo)": "#16A34A",
                    "Mayor Precio": "#DC2626",
                    "Opción Intermedia": "#0284C7"
                }
            )
            fig_precio.update_layout(yaxis_title="Monto CLP ($)", xaxis_title="Proveedor", showlegend=True)
            st.plotly_chart(fig_precio, use_container_width=True)

        with chart_col2:
            fig_plazo = px.bar(
                df_chart,
                x="Proveedor",
                y="Días de Entrega",
                color="Evaluación",
                text_auto=True,
                title="Plazos de Entrega Estimados (Días Corridos)",
                color_discrete_map={
                    "Mejor Opción (Mínimo)": "#16A34A",
                    "Mayor Precio": "#DC2626",
                    "Opción Intermedia": "#0284C7"
                }
            )
            fig_plazo.update_layout(yaxis_title="Días de Entrega", xaxis_title="Proveedor", showlegend=False)
            st.plotly_chart(fig_plazo, use_container_width=True)

    # -------------------------------------------------------------------------
    # GESTIÓN Y EXPORTACIÓN (EXCEL / PDF)
    # -------------------------------------------------------------------------
    st.markdown("#### 🛠️ Gestionar Ofertas Ingresadas")
    
    for i, cotizacion in enumerate(st.session_state["cotizaciones"]):
        col_info, col_btn = st.columns([5, 1])
        monto_clp_formateado = f"$ {int(cotizacion['Equiv. CLP ($)']):,}".replace(",", ".")
        col_info.markdown(f"Oferta de **{cotizacion['Proveedor']}** por **{monto_clp_formateado} CLP**")
        
        if col_btn.button("Eliminar", type="primary", key=f"eliminar_fila_{i}"):
            st.session_state["cotizaciones"].pop(i)
            st.rerun()
            
    st.write("---")

    max_monto_clp = df["Equiv. CLP ($)"].max()
    if max_monto_clp > 1000000:
        st.warning(
            f"**Control Financiero (> 1M CLP):** Requerimiento alcanza {formato_clp(max_monto_clp)}. "
            f"Se aplicaron los tipos de cambio mostrados en la parte superior.",
            icon="⚠️"
        )

    col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
    with col_btn1:
        excel_bytes = generar_excel_estilizado(df, solped, material, sociedad, indicadores)
        st.download_button(
            label="📥 Descargar Excel Corporativo",
            data=excel_bytes,
            file_name=f"cuadro_comparativo_solped_{solped}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_btn2:
        pdf_bytes = generar_pdf_ejecutivo(solped, material, sociedad, st.session_state["cotizaciones"], indicadores)
        st.download_button(
            label="📄 Descargar Reporte PDF",
            data=pdf_bytes,
            file_name=f"reporte_ejecutivo_solped_{solped}.pdf",
            mime="application/pdf"
        )
    with col_btn3:
        if st.button("Limpiar Cuadro", type="secondary"):
            st.session_state["cotizaciones"] = []
            st.rerun()
