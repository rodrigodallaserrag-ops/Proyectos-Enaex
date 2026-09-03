import datetime
import io
import json
import re
import ssl
import urllib.request
import urllib3
import pandas as pd
import requests
import streamlit as st
import plotly.express as px

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola de Compras — Enaex", layout="wide")

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
    try:
        val = float(monto)
    except (ValueError, TypeError):
        return str(monto)
        
    if moneda == "CLP":
        return f"$ {int(val):,}".replace(",", ".")
    elif moneda == "USD":
        return f"$ {val:,.2f}"
    elif moneda == "EUR":
        return f"€ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    elif moneda == "UF":
        return f"UF {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(monto)

# -----------------------------------------------------------------------------
# 2. CARGADOR INTELIGENTE DE PLANILLAS (DETECCION DE ENCABEZADOS REALES)
# -----------------------------------------------------------------------------
def cargar_planilla_inteligente(file):
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave = ['sp', 'solped', 'pr', 'código', 'descripción', 'cantidad', 'precio', 'proveedor', 'monto', 'material', 'texto breve', 'centro']
            mejor_sheet, mejor_fila_idx, max_coincidencias = None, None, 0

            for sheet_name in excel_file.sheet_names:
                try:
                    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                except Exception:
                    continue

                if df_raw.empty or len(df_raw) < 2:
                    continue

                for idx in range(min(50, len(df_raw))):
                    row = df_raw.iloc[idx]
                    celdas = [str(val).strip().lower() for val in row.values if pd.notna(val)]
                    coincidencias = sum(1 for kw in palabras_clave if any(kw in c for c in celdas))
                    if coincidencias > max_coincidencias:
                        max_coincidencias = coincidencias
                        mejor_sheet = sheet_name
                        mejor_fila_idx = idx

            if mejor_sheet is not None and max_coincidencias >= 2:
                df_raw = pd.read_excel(excel_file, sheet_name=mejor_sheet, header=None)
                df_clean = df_raw.iloc[mejor_fila_idx:].copy()
                df_clean.columns = [str(c).strip() if pd.notna(c) and str(c).strip() != 'None' else f"Col_{i}" for i, c in enumerate(df_clean.iloc[0])]
                df_clean = df_clean.iloc[1:].reset_index(drop=True).dropna(how='all')
                cols_utiles = [c for c in df_clean.columns if not str(c).startswith("Col_")]
                return df_clean[cols_utiles] if cols_utiles else df_clean

            return pd.read_excel(excel_file, sheet_name=0).dropna(how='all')

        elif nombre.endswith('.csv'):
            try:
                df = pd.read_csv(file_bytes, sep=None, engine='python', encoding='utf-8')
            except Exception:
                file_bytes.seek(0)
                df = pd.read_csv(file_bytes, sep=None, engine='python', encoding='latin-1')
            return df.dropna(how='all')

    except Exception as e:
        st.error(f"Error al procesar el archivo ({file.name}): {e}")
        return None

def extraer_materiales_de_masivo(df, id_solped):
    col_sp = next((c for c in df.columns if any(kw in str(c).lower() for kw in ['sp', 'solped', 'solicitud', 'pr'])), None)
    if not col_sp:
        return []
    
    df_filtrado = df[df[col_sp].astype(str).str.strip().str.upper() == str(id_solped).strip().upper()]
    if df_filtrado.empty:
        return []

    posiciones = []
    for idx, row in enumerate(df_filtrado.to_dict('records')):
        def get_val(keys, default):
            return next((row[c] for c in row.keys() if any(k in str(c).lower() for k in keys) and pd.notna(row[c])), default)

        posiciones.append({
            "Pos": idx + 1,
            "Material": get_val(['texto', 'desc', 'material', 'item'], "(Material)"),
            "Centro": get_val(['centro', 'plant', 'almacen'], "(Centro)"),
            "Cantidad": float(get_val(['cant', 'cantidad'], 1.0)),
            "UM": str(get_val(['um', 'unidad'], "C/U")).upper(),
            "Precio Unitario": float(get_val(['precio', 'monto', 'val', 'costo'], 0.0)),
            "Moneda": str(get_val(['moneda', 'curr'], "CLP")).upper(),
            "Proveedor": str(get_val(['proveedor', 'vendor', 'prov'], "(Proveedor)")),
            "Calendario de entrega": str(get_val(['calendario', 'fecha', 'entrega', 'plazo'], "(Calendario de entrega)")),
            "Días Entrega": int(float(get_val(['dias', 'lead', 'tratamiento'], 0))),
            "Observaciones": "(Observaciones)"
        })
    return posiciones

# -----------------------------------------------------------------------------
# 3. GENERADOR DE REPORTES PDF Y EXCEL ESTILIZADOS
# -----------------------------------------------------------------------------
def generar_pdf_ejecutivo(solped, material, sociedad, cotizaciones, datos_indicadores):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor('#0F172A'), spaceAfter=4)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=10, leading=12, textColor=colors.HexColor('#475569'), spaceAfter=14)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontSize=9, leading=12)
    header_table_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.whitesmoke, fontName='Helvetica-Bold')
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

    min_clp = min([c.get("Equiv. CLP ($)", 0) for c in cotizaciones]) if cotizaciones else 0

    for c in cotizaciones:
        monto_orig_fmt = aplicar_formato_regional(c.get("Monto Original", c.get("Precio Unitario", 0)), c.get("Moneda", "CLP"))
        clp_fmt = f"$ {int(c.get('Equiv. CLP ($)', 0)):,}".replace(",", ".")
        usd_fmt = f"$ {c.get('Equiv. USD ($)', 0):,.2f}"
        
        prov_text = f"<b>{c.get('Proveedor', '(Proveedor)')}</b>"
        if len(cotizaciones) > 1 and c.get("Equiv. CLP ($)", 0) == min_clp and min_clp > 0:
            prov_text += "<br/><font color='#16A34A'><b>★ Mejor Oferta</b></font>"

        table_data.append([
            Paragraph(prov_text, cell_style),
            Paragraph(monto_orig_fmt, cell_style),
            Paragraph(str(c.get("Moneda", "CLP")), cell_style),
            Paragraph(clp_fmt, cell_style),
            Paragraph(usd_fmt, cell_style),
            Paragraph(str(c.get("Fecha de Entrega", c.get("Calendario de entrega", "(Calendario de entrega)"))), cell_style),
            Paragraph(str(c.get("Observaciones", "(Observaciones)")), cell_style)
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

    max_monto = max([c.get("Equiv. CLP ($)", 0) for c in cotizaciones]) if cotizaciones else 0
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

def generar_excel_estilizado(df, solped, material, sociedad, datos_indicadores):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cuadro Comparativo"
    ws.views.sheetView[0].showGridLines = True

    navy_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=16, bold=True, color="0F172A")
    sub_font = Font(name="Calibri", size=10, italic=True, color="475569")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    
    green_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    meta_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws.cell(row=1, column=1, value="ENAEX — Consola de Compras").font = title_font
    ws.cell(row=2, column=1, value=f"Cuadro Comparativo de Cotizaciones — Solped N° {solped}").font = sub_font

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

    headers = list(df.columns)
    start_row = 7
    
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_clp_name = next((c for c in df.columns if "CLP" in str(c)), None)
    min_clp = df[col_clp_name].min() if col_clp_name and not df.empty else 0

    for r_offset, (_, row) in enumerate(df.iterrows()):
        current_row = start_row + 1 + r_offset
        is_best = (len(df) > 1 and col_clp_name and row[col_clp_name] == min_clp and min_clp > 0)
        
        for c_idx, col_name in enumerate(headers, start=1):
            cell_val = row[col_name]
            cell = ws.cell(row=current_row, column=c_idx, value=cell_val)
            cell.font = bold_font if is_best else regular_font
            if is_best:
                cell.fill = green_fill
            cell.border = thin_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. INICIALIZACION Y BARRA LATERAL
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

if "monto_input" not in st.session_state:
    st.session_state["monto_input"] = ""

if "moneda_input" not in st.session_state:
    st.session_state["moneda_input"] = "CLP"

with st.sidebar:
    st.header("📌 Datos Solped")
    solped = st.text_input("N° Solped", value="", placeholder="(ID SOLPED)")
    material = st.text_input("Código Material", value="", placeholder="(Código Material)")
    sociedad = st.selectbox("Sociedad", ["EC01", "EC06"])
    
    if "Offline" in indicadores["estado"]:
        st.divider()
        st.warning("⚠️ **Red Corporativa Bloqueada**\nAjusta los valores manualmente:")
        indicadores["uf"] = st.number_input("Valor UF", value=float(indicadores["uf"]), key="uf_manual")
        indicadores["dolar"] = st.number_input("Valor Dólar", value=float(indicadores["dolar"]), key="dolar_manual")
        indicadores["euro"] = st.number_input("Valor Euro", value=float(indicadores["euro"]), key="euro_manual")
        indicadores["fecha"] = datetime.date.today().strftime("%d-%m-%Y (Manual)")

st.title("🛒 Consola de Compras — Enaex")
st.caption(f"🗓️ Valores del día ({indicadores['fecha']}) - Estado API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1.5, 1.5, 1.5, 1])
col_uf.metric("UF", formato_clp(indicadores['uf']))
col_usd.metric("Dólar", formato_clp(indicadores['dolar']))
col_eur.metric("Euro", formato_clp(indicadores['euro']))

st.divider()

# -----------------------------------------------------------------------------
# 5. PESTAÑAS PRINCIPALES: MODO AUTOMÁTICO Y MODO MANUAL
# -----------------------------------------------------------------------------
tab_auto, tab_manual, tab_historico = st.tabs([
    "🤖 1. Autogestión y Carga Masiva", 
    "✍️ 2. Carga Manual de Ofertas", 
    "📜 3. Histórico General de OC (Previo a 2019)"
])

# =============================================================================
# TAB 1: PROCESAMIENTO AUTOMÁTICO (PLANILLA AUTOGESTIÓN + +20 MATERIALES)
# =============================================================================
with tab_auto:
    st.subheader("📋 Planilla Diaria de Autogestión")
    uploaded_file = st.file_uploader("Suba la Planilla de Autogestión Diaria (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="masivo_diario")

    if uploaded_file is not None:
        df_masivo = cargar_planilla_inteligente(uploaded_file)
        if df_masivo is not None and not df_masivo.empty:
            st.session_state["df_masivo"] = df_masivo
            st.success(f"✅ Archivo de requerimientos cargado ({len(df_masivo)} filas detectadas).")
            with st.expander("Ver Vista Previa de la Tabla Detectada"):
                st.dataframe(df_masivo.head(10), use_container_width=True)

    st.subheader("✏️ Evaluación por SOLPED (Soporta 20+ Materiales sin límite de filas)")
    col_search, col_btn = st.columns([3, 1])
    with col_search:
        solped_auto = st.text_input("Buscar ID SOLPED en la planilla:", value="", placeholder="(ID SOLPED)", key="auto_sp_input")
    with col_btn:
        st.write(" ")
        st.write(" ")
        cargar_btn = st.button("📥 Extraer Materiales", use_container_width=True, type="primary")

    if "datos_solped_actual" not in st.session_state or cargar_btn:
        sp_id = solped_auto.strip() if solped_auto.strip() else "(ID SOLPED)"
        materiales_cargados = []
        
        if "df_masivo" in st.session_state and solped_auto.strip():
            materiales_cargados = extraer_materiales_de_masivo(st.session_state["df_masivo"], sp_id)
        
        if not materiales_cargados:
            materiales_cargados = [
                {
                    "Pos": 1, 
                    "Material": "(Material)", 
                    "Centro": "(Centro)", 
                    "Cantidad": 1.0, 
                    "UM": "C/U", 
                    "Precio Unitario": 0.0, 
                    "Moneda": "CLP", 
                    "Proveedor": "(Proveedor)", 
                    "Calendario de entrega": "(Calendario de entrega)", 
                    "Días Entrega": 0, 
                    "Observaciones": "(Observaciones)"
                }
            ]
                
        st.session_state["datos_solped_actual"] = pd.DataFrame(materiales_cargados)
        st.session_state["solped_id_cargada"] = sp_id

    df_trabajo = st.session_state["datos_solped_actual"]

    df_auto_editado = st.data_editor(
        df_trabajo,
        num_rows="dynamic",
        use_container_width=True,
        height=300,
        column_config={
            "Pos": st.column_config.NumberColumn("Pos"),
            "Material": st.column_config.TextColumn("Material"),
            "Centro": st.column_config.TextColumn("Centro"),
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=0.01, format="%.2f"),
            "UM": st.column_config.SelectboxColumn("UM", options=["C/U", "KG", "LTS", "SET", "MTR", "TON"], default="C/U"),
            "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "EUR", "UF"], default="CLP"),
            "Proveedor": st.column_config.TextColumn("Proveedor"),
            "Calendario de entrega": st.column_config.TextColumn("Calendario de entrega"),
            "Días Entrega": st.column_config.NumberColumn("Días Entrega", min_value=0),
            "Observaciones": st.column_config.TextColumn("Observaciones")
        },
        key="editor_multi_posicion"
    )

    def calcular_monto_clp_row(row):
        try:
            precio = float(row.get("Precio Unitario", 0))
            cant = float(row.get("Cantidad", 1))
            moneda = str(row.get("Moneda", "CLP")).upper()
            
            uf_actual = st.session_state.get("uf_manual", indicadores["uf"])
            dolar_actual = st.session_state.get("dolar_manual", indicadores["dolar"])
            euro_actual = st.session_state.get("euro_manual", indicadores["euro"])
            
            factor = 1.0
            if moneda == "USD": factor = dolar_actual
            elif moneda == "EUR": factor = euro_actual
            elif moneda == "UF": factor = uf_actual
            
            return int(precio * cant * factor)
        except Exception:
            return 0

    df_auto_editado["Equiv. CLP ($)"] = df_auto_editado.apply(calcular_monto_clp_row, axis=1)
    st.session_state["datos_solped_actual"] = df_auto_editado

    if not df_auto_editado.empty and df_auto_editado["Equiv. CLP ($)"].sum() > 0:
        st.markdown("#### 📈 Gráfico de Evaluación Multimaterial")
        fig_auto = px.bar(
            df_auto_editado, x="Material", y="Equiv. CLP ($)", color="Proveedor",
            title="Comparativo de Montos por Material [CLP]", text_auto=',.0f'
        )
        st.plotly_chart(fig_auto, use_container_width=True)

# =============================================================================
# TAB 2: CARGA MANUAL DE OFERTAS (CODIGO BASE MEJORADO)
# =============================================================================
with tab_manual:
    st.subheader("➕ Carga Manual de Oferta Paso a Paso")

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
    col1.text_input("Proveedor*", key="proveedor_input", placeholder="(Proveedor)")
    col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"], key="moneda_input", on_change=formatear_caja_monto)
    col2.text_input("Monto Original*", placeholder="(Monto)", key="monto_input", on_change=formatear_caja_monto)
    col4.date_input("Calendario de entrega", min_value=datetime.date.today(), key="fecha_entrega_input")

    st.text_area("Observaciones Técnicas", key="obs_input", placeholder="(Observaciones)")
    st.button("Guardar en Cuadro Comparativo", on_click=procesar_guardado)

    st.divider()
    st.subheader("📊 Cuadro Comparativo Homogeneizado")

    if not st.session_state["cotizaciones"]:
        df_empty = pd.DataFrame(
            columns=["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Fecha de Entrega", "Observaciones"]
        )
        st.dataframe(df_empty, use_container_width=True)
    else:
        df_man = pd.DataFrame(st.session_state["cotizaciones"])
        
        df_visual = df_man.copy()
        df_visual["Monto Original"] = df_visual.apply(lambda fila: aplicar_formato_regional(fila["Monto Original"], fila["Moneda"]), axis=1)
        df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(x):,}".replace(",", "."))
        df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")

        st.dataframe(df_visual, use_container_width=True)

        if len(st.session_state["cotizaciones"]) >= 2:
            monto_min = df_man["Equiv. CLP ($)"].min()
            monto_max = df_man["Equiv. CLP ($)"].max()
            monto_prom = df_man["Equiv. CLP ($)"].mean()
            
            prov_min = df_man.loc[df_man["Equiv. CLP ($)"] == monto_min, "Proveedor"].values[0]
            prov_max = df_man.loc[df_man["Equiv. CLP ($)"] == monto_max, "Proveedor"].values[0]
            
            ahorro_vs_max = monto_max - monto_min
            pct_vs_max = (ahorro_vs_max / monto_max * 100) if monto_max > 0 else 0
            
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            kpi_col1.metric("🏆 Oferta Recomendada", prov_min, delta=formato_clp(monto_min))
            kpi_col2.metric("💰 Ahorro Máximo", formato_clp(ahorro_vs_max), delta=f"-{pct_vs_max:.1f}% vs {prov_max}")
            kpi_col3.metric("📊 Ahorro vs Promedio", formato_clp(monto_prom - monto_min))

# =============================================================================
# TAB 3: HISTÓRICO COMPLETO DE OC (SIN FILTRO DE AÑO 2019)
# =============================================================================
with tab_historico:
    st.subheader("📜 Consulta de Precios Históricos Completa")
    st.markdown("Suba el maestro de órdenes de compra antiguas para resolver las búsquedas de materiales de **2018, 2017 y anteriores**.")

    uploaded_hist = st.file_uploader("Suba la Base Histórica Completa (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="historico_oc")

    if uploaded_hist is not None:
        df_historico = cargar_planilla_inteligente(uploaded_hist)
        if df_historico is not None and not df_historico.empty:
            st.session_state["df_historico"] = df_historico
            st.success(f"✅ Histórico cargado con éxito ({len(df_historico)} filas disponibles sin límite de años).")

    if "df_historico" in st.session_state and st.session_state["df_historico"] is not None:
        busqueda_hist = st.text_input("Buscar por Material, Código, Proveedor o Año:", value="", placeholder="(Material) / (Proveedor) / (Año)")
        
        if busqueda_hist.strip():
            df_h = st.session_state["df_historico"]
            mask = df_h.astype(str).apply(lambda col: col.str.contains(busqueda_hist.strip(), case=False, na=False)).any(axis=1)
            res_hist = df_h[mask]
            st.markdown(f"**Registros Históricos Encontrados ({len(res_hist)}):**")
            st.dataframe(res_hist, use_container_width=True)
    else:
        st.info("💡 Suba un archivo histórico para consultar precios anteriores a 2019.")

# -----------------------------------------------------------------------------
# 6. GESTIÓN Y EXPORTACIÓN FINAL (EXCEL / PDF)
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📥 Exportar Informe Final")

df_final_export = pd.DataFrame()
if st.session_state["cotizaciones"]:
    df_final_export = pd.DataFrame(st.session_state["cotizaciones"])
elif "datos_solped_actual" in st.session_state:
    df_final_export = st.session_state["datos_solped_actual"]

if not df_final_export.empty:
    col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1.5, 3])
    with col_btn1:
        excel_bytes = generar_excel_estilizado(df_final_export, solped or "(ID SOLPED)", material or "(Material)", sociedad, indicadores)
        st.download_button(
            label="📥 Descargar Excel Corporativo",
            data=excel_bytes,
            file_name=f"cuadro_comparativo_solped_{solped or 'export'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_btn2:
        records_pdf = df_final_export.to_dict('records')
        pdf_bytes = generar_pdf_ejecutivo(solped or "(ID SOLPED)", material or "(Material)", sociedad, records_pdf, indicadores)
        st.download_button(
            label="📄 Descargar Reporte PDF",
            data=pdf_bytes,
            file_name=f"reporte_ejecutivo_solped_{solped or 'export'}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with col_btn3:
        if st.button("Limpiar Datos", type="secondary"):
            st.session_state["cotizaciones"] = []
            st.session_state.pop("datos_solped_actual", None)
            st.rerun()
else:
    st.info("💡 Cargue datos en el Modo Automático o Manual para habilitar las descargas en PDF y Excel corporativo.")
