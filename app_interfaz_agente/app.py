import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from datetime import datetime, date
import io

# Librerías para diseño de Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Librería para generación de PDF
from fpdf import FPDF

# =============================================================================
# CONFIGURACIÓN DE PÁGINA
# =============================================================================
st.set_page_config(
    page_title="Sistema Integrado de Evaluación de Ofertas - Enaex",
    page_icon="⚡",
    layout="wide"
)

st.markdown("""
<style>
    .main-header { font-size: 1.8rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1rem; color: #4B5563; margin-bottom: 1.5rem; }
    .stTable { font-size: 0.85rem; }
    .metric-card { background-color: #F3F4F6; padding: 1rem; border-radius: 0.5rem; border-left: 4px solid #1E3A8A; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# OBTENCIÓN DE INDICADORES FINANCIEROS EN TIEMPO REAL 
# =============================================================================
@st.cache_data(ttl=3600)
def obtener_indicadores_tiempo_real():
    valores_defecto = {"USD": 950.0, "EUR": 1020.0, "UF": 38000.0, "estado": False}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    try:
        response = requests.get("https://mindicador.cl/api", headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            usd = float(data.get("dolar", {}).get("valor", 950.0))
            eur = float(data.get("euro", {}).get("valor", 1020.0))
            uf = float(data.get("uf", {}).get("valor", 38000.0))
            if usd > 0 and eur > 0 and uf > 0:
                return {"USD": usd, "EUR": eur, "UF": uf, "estado": True}
    except Exception:
        pass

    try:
        response_alt = requests.get("https://open.er-api.com/v6/latest/USD", headers=headers, timeout=5)
        if response_alt.status_code == 200:
            data_alt = response_alt.json()
            rates = data_alt.get("rates", {})
            if "CLP" in rates:
                usd_clp = float(rates["CLP"])
                eur_rate = float(rates.get("EUR", 0.92))
                eur_clp = usd_clp / eur_rate if eur_rate > 0 else 1020.0
                
                return {
                    "USD": round(usd_clp, 2),
                    "EUR": round(eur_clp, 2),
                    "UF": 38300.0,
                    "estado": True
                }
    except Exception:
        pass

    return valores_defecto

# =============================================================================
# FUNCIONES DE EXPORTACIÓN Y FORMATO (EXCEL Y PDF)
# =============================================================================
def generar_excel_estilizado(df, moneda_vista, transporte_reporte="No Especificado"):
    buffer_excel = io.BytesIO()
    cols_export = [
        'SOLPED', 'Pos', 'Material', 'Centro', 'Cantidad', 'UM', 
        'Precio Unitario', 'Moneda', 'Proveedor Visual', 'Transporte',
        'Calendario de entrega', 'Días para Entrega', 'Monto Total Visualizado'
    ]
    df_export = df[[c for c in cols_export if c in df.columns]].copy()
    df_export.rename(columns={
        'Proveedor Visual': 'Proveedor', 
        'Monto Total Visualizado': f'Total ({moneda_vista})'
    }, inplace=True)
    
    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Cuadro Comparativo', startrow=3)
        workbook = writer.book
        worksheet = writer.sheets['Cuadro Comparativo']
        
        HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
        TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="1E3A8A")
        SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="6B7280")
        HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        DATA_FONT = Font(name="Calibri", size=10)
        
        THIN_BORDER = Border(
            left=Side(style='thin', color='E5E7EB'), right=Side(style='thin', color='E5E7EB'),
            top=Side(style='thin', color='E5E7EB'), bottom=Side(style='thin', color='E5E7EB')
        )
        
        worksheet['A1'] = "ENAEX - CUADRO COMPARATIVO DE OFERTAS"
        worksheet['A1'].font = TITLE_FONT
        
        subtitulo = f"Fecha de informe: {date.today().strftime('%d/%m/%Y')} | Moneda base: {moneda_vista}"
        if transporte_reporte and transporte_reporte != "No Especificado":
            subtitulo += f" | Transporte General: {transporte_reporte}"
            
        worksheet['A2'] = subtitulo
        worksheet['A2'].font = SUBTITLE_FONT
        
        for col_num in range(1, len(df_export.columns) + 1):
            cell = worksheet.cell(row=4, column=col_num)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = THIN_BORDER
        
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=5, max_row=4 + len(df_export), min_col=1, max_col=len(df_export.columns)), start=5):
            use_zebra = (row_idx % 2 == 0)
            for cell in row:
                cell.font = DATA_FONT
                cell.border = THIN_BORDER
                if use_zebra:
                    cell.fill = ZEBRA_FILL
                
                col_header = worksheet.cell(row=4, column=cell.column).value
                if col_header in ['Precio Unitario', f'Total ({moneda_vista})']:
                    cell.number_format = '$#,##0.00'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                elif col_header in ['Cantidad', 'Pos', 'Días para Entrega']:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col_header in ['SOLPED', 'Moneda', 'UM', 'Centro', 'Transporte', 'Calendario de entrega']:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row < 4: continue
                if cell.value: max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    return buffer_excel.getvalue()

def clean_str_pdf(txt):
    s = str(txt or '')
    reemplazos = {'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','á':'a','é':'e','í':'i','ó':'o','ú':'u','Ñ':'N','ñ':'n','°':''}
    for k, v in reemplazos.items(): s = s.replace(k, v)
    return s.encode('latin-1', 'ignore').decode('latin-1')

class PDFReport(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(30, 58, 138)
        self.cell(0, 8, 'ENAEX - EVALUACION COMPARATIVA DE OFERTAS', ln=True, align='C')
        self.set_font('Helvetica', 'I', 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, f'Fecha de Emision: {date.today().strftime("%d/%m/%Y")}', ln=True, align='C')
        self.ln(4)
    def footer(self):
        self.set_y(-12)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Pagina {self.page_no()}/{{nb}} - Documento Generado Automaticamente', align='C')

def generar_pdf(df, moneda_vista, transporte_reporte="No Especificado"):
    try:
        pdf = PDFReport(orientation='L', unit='mm', format='A4')
        pdf.alias_nb_pages()
        pdf.add_page()
        
        monto_total = df["Monto Total Visualizado"].sum() if "Monto Total Visualizado" in df.columns else 0.0
        pdf.set_fill_color(243, 244, 246)
        pdf.rect(10, pdf.get_y(), 277, 10, style='F')
        pdf.set_font('Helvetica', 'B', 9)
        pdf.set_text_color(30, 58, 138)
        
        resumen_txt = f'  RESUMEN GENERAL: Total Ofertas Evaluadas: {len(df)}    |    Monto Acumulado ({moneda_vista}): ${monto_total:,.2f}'
        if transporte_reporte and transporte_reporte != "No Especificado":
            resumen_txt += f'    |    Transporte General: {transporte_reporte}'
            
        pdf.cell(0, 8, resumen_txt, ln=True)
        pdf.ln(4)

        cols = [("SOLPED", 22), ("Material", 70), ("Proveedor", 42), ("Transporte", 20),
                ("Cant.", 15), ("Mon", 15), (f"Total ({moneda_vista})", 35), ("Entrega", 30)]

        pdf.set_font('Helvetica', 'B', 8)
        pdf.set_fill_color(30, 58, 138)
        pdf.set_text_color(255, 255, 255)

        for name, width in cols: pdf.cell(width, 7, name, border=1, align='C', fill=True)
        pdf.ln()

        pdf.set_font('Helvetica', '', 8)
        pdf.set_text_color(0, 0, 0)
        
        fill = False
        for _, row in df.iterrows():
            if pdf.get_y() > 180:
                pdf.add_page()
                pdf.set_font('Helvetica', 'B', 8)
                pdf.set_fill_color(30, 58, 138)
                pdf.set_text_color(255, 255, 255)
                for name, width in cols: pdf.cell(width, 7, name, border=1, align='C', fill=True)
                pdf.ln()
                pdf.set_font('Helvetica', '', 8)
                pdf.set_text_color(0, 0, 0)

            if fill: pdf.set_fill_color(249, 250, 251)
            else: pdf.set_fill_color(255, 255, 255)

            solped = clean_str_pdf(row.get('SOLPED', ''))[:15]
            material = clean_str_pdf(row.get('Material', ''))[:45]
            proveedor = clean_str_pdf(row.get('Proveedor Visual', ''))[:28]
            transporte = clean_str_pdf(row.get('Transporte', ''))[:12]
            cant = f"{row.get('Cantidad', 0):,.0f}"
            mon = clean_str_pdf(row.get('Moneda', 'CLP'))
            monto = f"${row.get('Monto Total Visualizado', 0):,.2f}"
            
            if pd.notna(row.get('Calendario de entrega')):
                try:
                    fecha_str = pd.to_datetime(row['Calendario de entrega']).strftime('%d/%m/%Y')
                except:
                    dias_val = row.get('Días para Entrega', 0)
                    fecha_str = f"{0 if pd.isna(dias_val) else int(dias_val)} dias"
            else:
                dias_val = row.get('Días para Entrega', 0)
                fecha_str = f"{0 if pd.isna(dias_val) else int(dias_val)} dias"

            pdf.cell(cols[0][1], 6, solped, border=1, align='C', fill=True)
            pdf.cell(cols[1][1], 6, material, border=1, align='L', fill=True)
            pdf.cell(cols[2][1], 6, proveedor, border=1, align='L', fill=True)
            pdf.cell(cols[3][1], 6, transporte, border=1, align='C', fill=True)
            pdf.cell(cols[4][1], 6, cant, border=1, align='C', fill=True)
            pdf.cell(cols[5][1], 6, mon, border=1, align='C', fill=True)
            pdf.cell(cols[6][1], 6, monto, border=1, align='R', fill=True)
            pdf.cell(cols[7][1], 6, fecha_str, border=1, align='C', fill=True)
            pdf.ln()
            fill = not fill

        out = pdf.output(dest='S') if hasattr(pdf, 'output') else b""
        if isinstance(out, str): return out.encode('latin-1')
        return bytes(out)
    except Exception: return b""

# =============================================================================
# FUNCIONES AUXILIARES Y LECTURA DE PLANILLAS (OPTIMIZADAS)
# =============================================================================
def deduplicar_columnas(columnas):
    vistos = {}
    columnas_unicas = []
    for c in columnas:
        c_str = str(c).strip()
        if c_str in vistos:
            vistos[c_str] += 1
            columnas_unicas.append(f"{c_str}_{vistos[c_str]}")
        else:
            vistos[c_str] = 0
            columnas_unicas.append(c_str)
    return columnas_unicas

def procesar_y_reparar_planilla(df):
    if df is None or df.empty: return df

    df.columns = deduplicar_columnas(df.columns)

    palabras_clave = ['solped', 'material', 'pos', 'texto', 'centro', 'cant', 'cantidad', 'proveedor', 'acreedor', 'vendor', 'documento', 'precio', 'moneda', 'solicitud', 'requerimiento', 'denominacion', 'descripcion']
    
    current_cols_lower = " ".join([str(c).lower() for c in df.columns])
    current_matches = sum(1 for kw in palabras_clave if kw in current_cols_lower)
    
    header_idx = -1
    best_matches = current_matches
    
    for idx in range(min(25, len(df))):
        row_str = " ".join([str(val).lower() for val in df.iloc[idx]])
        matches = sum(1 for kw in palabras_clave if kw in row_str)
        if matches > best_matches:
            header_idx = idx
            best_matches = matches

    if header_idx != -1:
        nuevas_columnas = []
        for i, val in enumerate(df.iloc[header_idx]):
            val_str = str(val).strip()
            if val_str.lower() in ['nan', 'none', '']:
                col_orig = str(df.columns[i])
                nuevas_columnas.append(col_orig if not col_orig.startswith("Unnamed") else f"Col_Vacia_{i}")
            else: nuevas_columnas.append(val_str)
        
        df.columns = deduplicar_columnas(nuevas_columnas)
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        if not df.empty:
            col_0_series = df.iloc[:, 0].astype(str).str.strip()
            col_0_nombre = str(df.columns[0]).strip()
            df = df[col_0_series != col_0_nombre].reset_index(drop=True)

    nuevos_nombres = {}
    for idx, col in enumerate(df.columns):
        col_str = str(col).strip()
        if col_str.startswith("Col_Vacia_") or col_str.startswith("Unnamed:"):
            series_clean = df.iloc[:, idx].dropna().astype(str).str.strip()
            series_clean = series_clean[~series_clean.str.lower().isin(['nan', 'none', ''])]
            if not series_clean.empty:
                nuevos_nombres[col] = series_clean.iloc[0]

    if nuevos_nombres: 
        df = df.rename(columns=nuevos_nombres)

    df.columns = deduplicar_columnas(df.columns)

    cols = list(df.columns)
    id_col = None
    for c in cols:
        c_low = str(c).lower().strip()
        if c_low in ['sp', 'solped', 'solicitud', 'n° solped', 'num solped', 'solped/pos', 'sol.pedido']:
            id_col = c
            break
            
    if not id_col:
        for c in cols:
            c_low = str(c).lower().strip()
            if any(kw in c_low for kw in ['solped', 'solicitud', 'requerimiento', 'pr', 'pedido']):
                id_col = c
                break
                
    if id_col and id_col in cols:
        df = df.rename(columns={id_col: 'SOLPED'})
        cols = list(df.columns)
        cols.remove('SOLPED')
        cols.insert(0, 'SOLPED')
        df = df[cols]

    return df.dropna(how='all')

def extraer_materiales_de_masivo(df, id_solped):
    if df is None or df.empty: return []
        
    raw_search = str(id_solped).strip()
    if not raw_search or raw_search.lower() in ["(id solped)", "none", "nan", "", "n/a"]: return []
        
    digits_search = re.sub(r'\D', '', raw_search)
    clean_search = re.sub(r'[^a-zA-Z0-9]', '', raw_search).lower()

    sp_cols = [c for c in df.columns if any(kw in str(c).lower() for kw in ['solped', 'solicitud', 'sp', 'requerimiento', 'pedido', 'pr'])]
    if not sp_cols:
        sp_cols = [df.columns[0]]

    df_filtrado = pd.DataFrame()
    for col in sp_cols:
        col_series = df[col].astype(str).str.strip()
        col_series_clean = col_series.apply(lambda x: re.sub(r'[^a-zA-Z0-9]', '', str(x)).lower())
        
        mask = col_series_clean == clean_search
        
        if not mask.any() and len(digits_search) >= 3:
            col_digits = col_series.str.replace(r'\D', '', regex=True)
            mask = col_digits == digits_search
            
        if not mask.any() and len(clean_search) >= 4:
            mask = col_series_clean.str.contains(clean_search, regex=False)
            
        if mask.any():
            df_filtrado = df[mask]
            break

    if df_filtrado.empty: return []

    posiciones = []
    for idx, row in enumerate(df_filtrado.to_dict('records')):
        def get_val(keys, default):
            candidates = []
            for k in keys:
                for col, val in row.items():
                    if k in str(col).lower() and pd.notna(val):
                        val_str = str(val).strip()
                        if val_str != "" and val_str.lower() not in ["nan", "none", "null"]:
                            candidates.append(val_str)
            if not candidates: return default
            sin_truncar = [c for c in candidates if '...' not in c and '…' not in c]
            if sin_truncar: return max(sin_truncar, key=len)
            return max(candidates, key=len)

        def clean_num(val, default=0.0):
            try:
                if isinstance(val, (int, float)): return float(val)
                s = re.sub(r'[^\d.,-]', '', str(val).strip())
                if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
                elif '.' in s and len(s.split('.')[-1]) == 3: s = s.replace('.', '')
                elif ',' in s: s = s.replace(',', '.')
                return float(s)
            except: return default

        solped_real = get_val(['solped', 'solicitud', 'sp', 'requerimiento', 'pedido'], raw_search)
        mat_desc = get_val(['texto breve de material', 'texto breve', 'denominación del material', 'denominacion del material', 'descripción del material', 'descripcion del material', 'descripción', 'descripcion', 'denominacion', 'denominación', 'material', 'texto', 'item', 'artículo', 'articulo', 'breve'], f"Material {idx+1}")
        centro_desc = get_val(['nombre centro', 'nombre del centro', 'denominación centro', 'denominacion centro', 'descripción centro', 'descripcion centro', 'texto centro', 'centro', 'plant', 'almacen', 'alm'], "E001")
        proveedor_sugerido = str(get_val(['proveedor', 'vendor', 'prov', 'nam', 'razon social', 'acreedor', 'nombre proveedor', 'nombre_proveedor', 'nom_prov', 'lifnr', 'nombre del proveedor', 'supplier', 'nombre', 'distribuidor'], ""))
        
        pos_val = get_val(['pos', 'posicion', 'posición', 'item', 'linea'], idx + 1)
        try: pos_num = int(clean_num(pos_val, idx + 1))
        except: pos_num = idx + 1

        cant_raw = clean_num(get_val(['cant', 'cantidad', 'ctd'], 1.0), 1.0)
        cant_clean = int(cant_raw) if float(cant_raw).is_integer() else cant_raw

        fecha_hist = get_val(['fecha', 'date', 'entrega', 'creacion', 'f.pedido'], None)
        fecha_parsed = date.today()
        if fecha_hist and fecha_hist != "":
            try: fecha_parsed = pd.to_datetime(fecha_hist).date()
            except: fecha_parsed = date.today()

        posiciones.append({
            "SOLPED": str(solped_real),
            "Pos": pos_num,
            "Material": str(mat_desc),
            "Centro": str(centro_desc),
            "Cantidad": cant_clean,
            "UM": str(get_val(['um', 'unidad', 'unid', 'medida'], "C/U")).upper(),
            "Precio Unitario": clean_num(get_val(['precio', 'monto', 'val', 'costo', 'p.u', 'neto', 'p.unitario'], 0.0), 0.0),
            "Moneda": str(get_val(['moneda', 'curr', 'mon'], "CLP")).upper(),
            "Proveedor": proveedor_sugerido,
            "Transporte": str(get_val(['transporte', 'incoterm', 'despacho'], "EXW")),
            "Calendario de entrega": fecha_parsed,
            "Observaciones": str(get_val(['obs', 'observacion', 'comentario', 'notas'], ""))
        })
    return posiciones

def convertir_moneda(monto, moneda_origen, tc_usd, tc_uf, tc_eur):
    monto, moneda_origen = float(monto or 0.0), str(moneda_origen).upper()
    if moneda_origen == "CLP": clp = monto
    elif moneda_origen == "USD": clp = monto * tc_usd
    elif moneda_origen == "UF": clp = monto * tc_uf
    elif moneda_origen == "EUR": clp = monto * tc_eur
    else: clp = monto
    usd = clp / tc_usd if tc_usd > 0 else 0.0
    eur = clp / tc_eur if tc_eur > 0 else 0.0
    return clp, usd, eur

@st.cache_data(show_spinner="Procesando archivo inteligentemente...")
def leer_archivo_cached(file_bytes, file_name):
    try:
        if file_name.endswith(".csv"): 
            df_raw = pd.read_csv(io.BytesIO(file_bytes))
        else:
            xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
            
            best_sheet = None
            best_score = -1
            keywords_scoring = ['solped', 'material', 'pos', 'texto', 'centro', 'cant', 'cantidad', 'proveedor', 'acreedor', 'vendor', 'precio', 'monto', 'pr', 'requerimiento']
            
            for sheet in xls.sheet_names:
                try:
                    df_tmp = pd.read_excel(xls, sheet_name=sheet, nrows=30)
                    if df_tmp.empty: continue
                    
                    sheet_text = " ".join(df_tmp.columns.astype(str).str.lower()) + " " + " ".join(df_tmp.fillna('').astype(str).values.flatten()[:200]).lower()
                    score = sum(1 for kw in keywords_scoring if kw in sheet_text)
                    
                    if any(k in sheet.lower() for k in ['base', 'dato', 'solped', 'detalle', 'cuadro', 'mat', 'reporte', 'pr']):
                        score += 3
                        
                    if score > best_score:
                        best_score = score
                        best_sheet = sheet
                except Exception:
                    continue
            
            df_raw = pd.read_excel(xls, sheet_name=best_sheet if best_sheet else xls.sheet_names[0])
            
        df_clean = df_raw.dropna(axis=1, how='all').dropna(axis=0, how='all')
        df_clean.columns = deduplicar_columnas(df_clean.columns)
        df_procesado = procesar_y_reparar_planilla(df_clean)
        
        if df_procesado is not None and not df_procesado.empty:
            df_procesado.columns = deduplicar_columnas(df_procesado.columns)
            df_procesado['cantidad_nulos'] = df_procesado.isnull().sum(axis=1)
            df_procesado = df_procesado.sort_values(by='cantidad_nulos').drop(columns=['cantidad_nulos']).reset_index(drop=True)
            
        return df_procesado
    except Exception as e:
        st.error(f"Error al leer el archivo {file_name}: {e}")
        return None

def leer_archivo(file_uploader):
    if file_uploader is None: return None
    return leer_archivo_cached(file_uploader.getvalue(), file_uploader.name)

# =============================================================================
# INICIALIZACIÓN DE ESTADO
# =============================================================================
if "df_masivo" not in st.session_state: st.session_state.df_masivo = None
if "df_historico" not in st.session_state: st.session_state.df_historico = None
if "ofertas_manuales" not in st.session_state: st.session_state.ofertas_manuales = []

OPCIONES_TRANSPORTE = ["T. Gil", "T. Bello", "Pullman", "Retiramos", "EXW", "FCA", "FOB", "CFR", "CIF", "CPT", "CIP", "DAT", "DDP"]

# =============================================================================
# ENCABEZADO Y PARÁMETROS GLOBALES
# =============================================================================
st.markdown("<div class='main-header'>⚡ Sistema Integrado de Evaluación de Ofertas</div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Parámetros de Cambio")
    indicadores = obtener_indicadores_tiempo_real()
    
    if indicadores["estado"]: st.success("🟢 Indicadores actualizados en vivo")
    else: st.warning("⚠️ Usando valores por defecto (Sin conexión).")
        
    if st.button("🔄 Actualizar Tasas API", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    tc_usd = st.number_input("Tipo de Cambio USD / CLP", value=indicadores["USD"], step=1.0, format="%.2f")
    tc_uf = st.number_input("Tipo de Cambio UF / CLP", value=indicadores["UF"], step=100.0, format="%.2f")
    tc_eur = st.number_input("Tipo de Cambio EUR / CLP", value=indicadores["EUR"], step=1.0, format="%.2f")
    st.divider()
    
    st.header("📂 Carga de Archivos Base")
    file_autogestion = st.file_uploader("1. Plantilla Autogestión (Estructura Base)", type=["xlsx", "xls", "csv", "xlsm"])
    file_cuadro = st.file_uploader("2. Cuadro Comparativo (Datos Brutos / Histórico)", type=["xlsx", "xls", "csv", "xlsm"])
    
    if st.button("Procesar y Unir Archivos", type="primary", use_container_width=True):
        if file_autogestion:
            df_base = leer_archivo(file_autogestion)
            st.session_state.df_masivo = df_base
            if df_base is not None:
                st.success(f"Plantilla Base cargada ({len(df_base)} filas).")
            
        if file_cuadro:
            df_raw_hist = leer_archivo(file_cuadro)
            st.session_state.df_historico = df_raw_hist
            if df_raw_hist is not None:
                st.success(f"Cuadro Bruto / Histórico cargado ({len(df_raw_hist)} filas).")

        if st.session_state.df_masivo is not None and st.session_state.df_historico is not None:
            col_mat_base = next((c for c in st.session_state.df_masivo.columns if 'material' in str(c).lower()), None)
            col_mat_hist = next((c for c in st.session_state.df_historico.columns if 'material' in str(c).lower()), None)
            
            if col_mat_base and col_mat_hist:
                st.session_state.df_masivo[col_mat_base] = st.session_state.df_masivo[col_mat_base].astype(str).str.strip()
                st.session_state.df_historico[col_mat_hist] = st.session_state.df_historico[col_mat_hist].astype(str).str.strip()

                col_fecha_hist = next((c for c in st.session_state.df_historico.columns if any(k in str(c).lower() for k in ['fecha', 'date', 'ano', 'año', 'creacion'])), None)
                if col_fecha_hist:
                    st.session_state.df_historico['_fecha_tmp'] = pd.to_datetime(st.session_state.df_historico[col_fecha_hist], errors='coerce')
                    st.session_state.df_historico = st.session_state.df_historico.sort_values(by='_fecha_tmp', ascending=False)

                df_hist_unique = st.session_state.df_historico.drop_duplicates(subset=[col_mat_hist], keep='first')
                if '_fecha_tmp' in df_hist_unique.columns:
                    df_hist_unique = df_hist_unique.drop(columns=['_fecha_tmp'])

                df_merged = pd.merge(st.session_state.df_masivo, df_hist_unique, left_on=col_mat_base, right_on=col_mat_hist, how='left', suffixes=('', '_Bruto'))
                st.session_state.df_masivo = df_merged
                st.success("✅ Datos brutos e histórico integrados.")
            else:
                st.warning("No se encontró la columna 'Material' en ambas planillas para realizar el cruce.")

if st.session_state.df_masivo is not None:
    with st.expander("👀 Vista Previa de la Base Integrada", expanded=False):
        st.dataframe(st.session_state.df_masivo, use_container_width=True)

tabs = st.tabs(["✏️ Evaluación por SOLPED", "➕ Carga Manual / Directa", "📊 Cuadro Comparativo Integrado"])

# =============================================================================
# TAB 1: EVALUACIÓN POR SOLPED (FLUIDA - SIN CONTENEDOR RÍGIDO DE ALTURA)
# =============================================================================
with tabs[0]:
    st.subheader("✏️ Evaluación por SOLPED")
    
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        solped_id = st.text_input("Buscar ID SOLPED en la plantilla:", placeholder="Ej: PR151762")
    with col_btn:
        st.write("")
        st.write("")
        btn_extraer = st.button("📤 Extraer Materiales", type="primary", use_container_width=True)

    if btn_extraer and solped_id.strip():
        if st.session_state.df_masivo is not None:
            mats = extraer_materiales_de_masivo(st.session_state.df_masivo, solped_id)
            if mats:
                st.session_state[f"lista_solped_{solped_id}"] = mats
                st.success(f"Se encontraron {len(mats)} posiciones para la SOLPED **{solped_id}**")
            else:
                st.warning(f"No se encontraron registros exactos para la SOLPED '{solped_id}'.")
        else:
            st.info("Carga las planillas en el menú lateral y haz clic en Procesar.")

    key_lista = f"lista_solped_{solped_id}" if (solped_id and f"lista_solped_{solped_id}" in st.session_state) else "lista_solped_default"

    if key_lista not in st.session_state:
        st.session_state[key_lista] = [{
            "SOLPED": "PR151762", "Pos": 1, "Material": "PROYECTOR LED 50W DS1", "Centro": "E024", "Cantidad": 10, 
            "UM": "CADA UNO", "Precio Unitario": 8190.0, "Moneda": "CLP", 
            "Proveedor": "", "Transporte": "EXW", "Calendario de entrega": date.today(), "Observaciones": ""
        }]

    lista_materiales = st.session_state[key_lista]

    if lista_materiales:
        col_m_head, col_m_save = st.columns([2, 1])
        with col_m_head:
            st.write(f"**Materiales extraídos ({len(lista_materiales)} ítems):**")
        with col_m_save:
            if st.button("💾 Guardar Oferta en Cuadro Comparativo", type="primary", key="btn_save_top_t1", use_container_width=True):
                for r in st.session_state[key_lista]:
                    clp, usd, eur = convertir_moneda(float(r["Precio Unitario"]) * float(r["Cantidad"]), r["Moneda"], tc_usd, tc_uf, tc_eur)
                    item_guardar = r.copy()
                    if not item_guardar.get("SOLPED") or item_guardar.get("SOLPED") in ["", "N/A", "None", "nan"]:
                        item_guardar["SOLPED"] = solped_id if solped_id else "N/A"
                    item_guardar["Total CLP"] = clp
                    item_guardar["Total USD"] = usd
                    item_guardar["Total EUR"] = eur
                    st.session_state.ofertas_manuales.append(item_guardar)
                st.success("¡Oferta guardada exitosamente en el Cuadro Comparativo!")

        idx_a_eliminar = None
        
        # Flujo natural sin límite fijo de altura
        for idx, item in enumerate(lista_materiales):
            with st.container(border=True):
                col_title, col_del = st.columns([8, 2])
                with col_title:
                    st.markdown(f"### SOLPED: {item.get('SOLPED', solped_id)} | Pos {item['Pos']}: {item['Material']}")
                    st.markdown(
                        f"<div style='color: #8C8C8C; font-size: 0.9em; padding-bottom: 10px;'>"
                        f"<b>Centro:</b> {item['Centro']} | <b>UM Original:</b> {item['UM']}</div>", 
                        unsafe_allow_html=True
                    )
                with col_del:
                    if st.button("🗑️ Eliminar", key=f"btn_del_t1_{idx}", type="primary", use_container_width=True):
                        idx_a_eliminar = idx

                c1, c2, c3, c4, c5, c6 = st.columns([1, 1.5, 1, 1.5, 1.5, 1.5])
                
                cant_val = float(item.get("Cantidad", 1.0))
                cant_val_clean = int(cant_val) if cant_val.is_integer() else cant_val
                item["Cantidad"] = c1.number_input("Cantidad", value=cant_val_clean, format="%g", key=f"cant_{key_lista}_{idx}")
                
                item["Precio Unitario"] = c2.number_input("Precio Unitario", value=float(item.get("Precio Unitario", 0.0)), format="%.2f", key=f"pu_{key_lista}_{idx}")
                
                moneda_opts = ["CLP", "USD", "EUR"]
                m_idx = moneda_opts.index(item.get("Moneda", "CLP")) if item.get("Moneda") in moneda_opts else 0
                item["Moneda"] = c3.selectbox("Moneda", moneda_opts, index=m_idx, key=f"mon_{key_lista}_{idx}")
                
                item["Proveedor"] = c4.text_input("Proveedor", value=str(item.get("Proveedor", "")), key=f"prov_{key_lista}_{idx}")
                
                trans_opts = OPCIONES_TRANSPORTE
                t_idx = trans_opts.index(item.get("Transporte", "EXW")) if item.get("Transporte") in trans_opts else 4
                item["Transporte"] = c5.selectbox("Transporte", trans_opts, index=t_idx, key=f"trans_{key_lista}_{idx}")

                valor_fecha = item.get("Calendario de entrega", date.today())
                if isinstance(valor_fecha, str):
                    try: valor_fecha = datetime.strptime(valor_fecha, "%Y-%m-%d").date()
                    except: valor_fecha = date.today()
                
                item["Calendario de entrega"] = c6.date_input("Fecha Entrega", value=valor_fecha, key=f"date_{key_lista}_{idx}")

        if idx_a_eliminar is not None:
            st.session_state[key_lista].pop(idx_a_eliminar)
            st.rerun()

        st.divider()
        if st.button("💾 Guardar Oferta en Cuadro Comparativo", type="primary", key="btn_save_bot_t1"):
            for r in st.session_state[key_lista]:
                clp, usd, eur = convertir_moneda(float(r["Precio Unitario"]) * float(r["Cantidad"]), r["Moneda"], tc_usd, tc_uf, tc_eur)
                item_guardar = r.copy()
                if not item_guardar.get("SOLPED") or item_guardar.get("SOLPED") in ["", "N/A", "None", "nan"]:
                    item_guardar["SOLPED"] = solped_id if solped_id else "N/A"
                item_guardar["Total CLP"] = clp
                item_guardar["Total USD"] = usd
                item_guardar["Total EUR"] = eur
                st.session_state.ofertas_manuales.append(item_guardar)
            st.success("¡Oferta guardada exitosamente en el Cuadro Comparativo!")

# =============================================================================
# TAB 2: CARGA MANUAL DIRECTA (FLUIDA - SIN CONTENEDOR RÍGIDO DE ALTURA)
# =============================================================================
with tabs[1]:
    st.subheader("➕ Carga Manual de Oferta Paso a Paso")
    
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1: manual_solped = st.text_input("Ingresar N° SOLPED para Autocompletar:", placeholder="Ej: PR175798")
    with col_s2:
        st.write("")
        st.write("")
        btn_cargar_manual = st.button("📥 Cargar Requerimiento", use_container_width=True)

    if btn_cargar_manual and manual_solped:
        if st.session_state.df_masivo is not None:
            mats = extraer_materiales_de_masivo(st.session_state.df_masivo, manual_solped)
            if mats:
                st.session_state["manual_items_list"] = mats
                st.success(f"Materiales cargados automáticamente desde la SOLPED {manual_solped}")
            else: st.warning(f"No se encontró la SOLPED {manual_solped}.")
        else: st.info("Sube las planillas en la barra lateral.")

    if "manual_items_list" not in st.session_state:
        st.session_state["manual_items_list"] = [{
            "SOLPED": "MANUAL", "Pos": 1, "Material": "Ítem Manual 1", "Centro": "E001", "Cantidad": 1, "UM": "C/U",
            "Precio Unitario": 0.0, "Moneda": "CLP", "Proveedor": "", "Transporte": "EXW",
            "Calendario de entrega": date.today(), "Observaciones": ""
        }]

    lista_manual = st.session_state["manual_items_list"]
    
    col_man_btn1, col_man_btn2 = st.columns([1, 1])
    with col_man_btn1:
        if st.button("➕ Agregar Nuevo Material a la Lista", use_container_width=True):
            lista_manual.append({
                "SOLPED": manual_solped if manual_solped else "MANUAL",
                "Pos": len(lista_manual) + 1, "Material": "Nuevo Material", "Centro": "E001", "Cantidad": 1, "UM": "C/U",
                "Precio Unitario": 0.0, "Moneda": "CLP", "Proveedor": "", "Transporte": "EXW",
                "Calendario de entrega": date.today(), "Observaciones": ""
            })
            st.rerun()
    with col_man_btn2:
        if st.button("💾 Guardar Cotización Manual Completa", type="primary", key="btn_save_top_t2", use_container_width=True):
            for item in st.session_state["manual_items_list"]:
                clp, usd, eur = convertir_moneda(float(item["Precio Unitario"]) * float(item["Cantidad"]), item["Moneda"], tc_usd, tc_uf, tc_eur)
                item_guardar = item.copy()
                if not item_guardar.get("SOLPED") or item_guardar.get("SOLPED") in ["", "N/A", "None", "nan"]:
                    item_guardar["SOLPED"] = manual_solped if manual_solped else "MANUAL"
                item_guardar["Total CLP"] = clp
                item_guardar["Total USD"] = usd
                item_guardar["Total EUR"] = eur
                st.session_state.ofertas_manuales.append(item_guardar)
            st.success("¡Cotización agregada al Cuadro Comparativo!")

    idx_del_manual = None
    # Flujo natural sin límite fijo de altura
    for idx, item in enumerate(lista_manual):
        with st.container(border=True):
            col_m_title, col_m_del = st.columns([8, 2])
            with col_m_title:
                item["Material"] = st.text_input("Descripción / Material", value=item.get("Material", ""), key=f"man_mat_{idx}")
            with col_m_del:
                if st.button("🗑️ Eliminar", key=f"btn_del_t2_{idx}", type="primary", use_container_width=True):
                    idx_del_manual = idx

            c1, c2, c3, c4, c5, c6 = st.columns([1, 1.5, 1, 1.5, 1.5, 1.5])
            
            man_cant_val = float(item.get("Cantidad", 1.0))
            man_cant_clean = int(man_cant_val) if man_cant_val.is_integer() else man_cant_val
            item["Cantidad"] = c1.number_input("Cantidad", value=man_cant_clean, format="%g", key=f"man_cant_{idx}")
            
            item["Precio Unitario"] = c2.number_input("Precio Unitario", value=float(item.get("Precio Unitario", 0.0)), format="%.2f", key=f"man_pu_{idx}")
            
            moneda_opts = ["CLP", "USD", "EUR"]
            m_idx = moneda_opts.index(item.get("Moneda", "CLP")) if item.get("Moneda") in moneda_opts else 0
            item["Moneda"] = c3.selectbox("Moneda", moneda_opts, index=m_idx, key=f"man_mon_{idx}")
            
            item["Proveedor"] = c4.text_input("Proveedor", value=str(item.get("Proveedor", "")), key=f"man_prov_{idx}")
            
            trans_opts = OPCIONES_TRANSPORTE
            t_idx = trans_opts.index(item.get("Transporte", "EXW")) if item.get("Transporte") in trans_opts else 4
            item["Transporte"] = c5.selectbox("Transporte", trans_opts, index=t_idx, key=f"man_trans_{idx}")

            valor_fecha_man = item.get("Calendario de entrega", date.today())
            if isinstance(valor_fecha_man, str):
                try: valor_fecha_man = datetime.strptime(valor_fecha_man, "%Y-%m-%d").date()
                except: valor_fecha_man = date.today()
            
            item["Calendario de entrega"] = c6.date_input("Fecha Entrega", value=valor_fecha_man, key=f"man_date_{idx}")

    if idx_del_manual is not None:
        st.session_state["manual_items_list"].pop(idx_del_manual)
        st.rerun()

    if st.button("💾 Guardar Cotización Manual Completa", type="primary", key="btn_save_bot_t2"):
        for item in st.session_state["manual_items_list"]:
            clp, usd, eur = convertir_moneda(float(item["Precio Unitario"]) * float(item["Cantidad"]), item["Moneda"], tc_usd, tc_uf, tc_eur)
            item_guardar = item.copy()
            if not item_guardar.get("SOLPED") or item_guardar.get("SOLPED") in ["", "N/A", "None", "nan"]:
                item_guardar["SOLPED"] = manual_solped if manual_solped else "MANUAL"
            item_guardar["Total CLP"] = clp
            item_guardar["Total USD"] = usd
            item_guardar["Total EUR"] = eur
            st.session_state.ofertas_manuales.append(item_guardar)
        st.success("¡Cotización agregada al Cuadro Comparativo!")

# =============================================================================
# TAB 3: CUADRO COMPARATIVO INTEGRADO & DESCARGAS
# =============================================================================
with tabs[2]:
    st.subheader("📊 Cuadro Comparativo Integrado")
    
    if st.session_state.ofertas_manuales:
        df_comp = pd.DataFrame(st.session_state.ofertas_manuales)
        
        if 'SOLPED' not in df_comp.columns:
            df_comp['SOLPED'] = 'N/A'
        df_comp['SOLPED'] = df_comp['SOLPED'].fillna('N/A').astype(str).replace({'': 'N/A', 'none': 'N/A', 'None': 'N/A', 'nan': 'N/A'})
        
        df_comp['Proveedor'] = df_comp['Proveedor'].fillna('Sin Especificar').astype(str).str.strip()
        df_comp['Proveedor Visual'] = df_comp['Proveedor'].apply(
            lambda x: 'Sin Especificar' if x in ['', 'none', 'None', 'nan', 'NULL', 'null'] else x
        )

        cols_orden = [
            'SOLPED', 'Pos', 'Material', 'Centro', 'Cantidad', 'UM', 
            'Precio Unitario', 'Moneda', 'Proveedor Visual', 'Transporte', 
            'Calendario de entrega', 'Total CLP', 'Total USD', 'Total EUR', 'Observaciones'
        ]
        
        existing_cols = [c for c in cols_orden if c in df_comp.columns]
        other_cols = [c for c in df_comp.columns if c not in cols_orden and c not in ['Total CLP', 'Total USD', 'Total EUR']]
        df_comp = df_comp[existing_cols + other_cols]

        col_opt1, col_opt2 = st.columns([1, 1])
        with col_opt1:
            moneda_vista = st.radio("💱 Seleccionar Moneda de Visualización:", options=["CLP", "USD", "EUR"], horizontal=True)
        with col_opt2:
            transporte_reporte = st.selectbox("🚚 Transporte General (Cabecera reporte):", options=["No Especificado"] + OPCIONES_TRANSPORTE, index=0)

        if moneda_vista == "CLP": df_comp["Monto Total Visualizado"] = df_comp["Total CLP"]
        elif moneda_vista == "USD": df_comp["Monto Total Visualizado"] = df_comp["Total USD"]
        elif moneda_vista == "EUR": df_comp["Monto Total Visualizado"] = df_comp["Total EUR"]

        df_comp['Calendario de entrega'] = pd.to_datetime(df_comp['Calendario de entrega'], errors='coerce')
        hoy = pd.Timestamp(date.today())
        
        df_comp['Días para Entrega'] = (df_comp['Calendario de entrega'] - hoy).dt.days
        df_comp['Días para Entrega'] = df_comp['Días para Entrega'].apply(lambda x: int(x) if pd.notna(x) and x > 0 else 0)
        df_comp['Calendario de entrega'] = df_comp['Calendario de entrega'].dt.strftime('%d/%m/%Y').fillna('N/A')

        # BARRA SUPERIOR DE ACCIONES DIRECTAS
        bytes_excel = generar_excel_estilizado(df_comp, moneda_vista, transporte_reporte)
        bytes_pdf = generar_pdf(df_comp, moneda_vista, transporte_reporte)
        
        col_top_d1, col_top_d2, col_top_d3 = st.columns([1, 1, 1])
        with col_top_d1:
            if bytes_excel:
                st.download_button(label="📊 Reporte Excel", data=bytes_excel, file_name=f"Reporte_Comparativo_{date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="dl_excel_top")
        with col_top_d2:
            if bytes_pdf:
                st.download_button(label="📄 Reporte PDF", data=bytes_pdf, file_name=f"Reporte_Comparativo_{date.today()}.pdf", mime="application/pdf", use_container_width=True, key="dl_pdf_top")
        with col_top_d3:
            if st.button("🗑️ Limpiar Todo", use_container_width=True, key="btn_clear_top"):
                st.session_state.ofertas_manuales = []
                st.rerun()

        st.divider()

        # CONTENIDO DIRECTO EN LA PESTAÑA
        st.markdown("### 🏆 Motor de Recomendación")
        
        def highlight_best(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            group_cols = [c for c in ['SOLPED', 'Material'] if c in df.columns]
            if group_cols:
                for name, group in df.groupby(group_cols):
                    if len(group) > 1:
                        if 'Monto Total Visualizado' in df.columns:
                            styles.loc[group['Monto Total Visualizado'].idxmin(), 'Monto Total Visualizado'] = 'background-color: #D1FAE5; color: #065F46; font-weight: bold;'
                        if 'Días para Entrega' in df.columns:
                            styles.loc[group['Días para Entrega'].idxmin(), 'Días para Entrega'] = 'background-color: #DBEAFE; color: #1E3A8A; font-weight: bold;'
            return styles

        if not df_comp.empty:
            styled_df_comp = df_comp.style.apply(highlight_best, axis=None).format({
                "Cantidad": lambda x: f"{int(x)}" if pd.notna(x) and float(x).is_integer() else (f"{x:,.2f}" if pd.notna(x) else ""),
                "Monto Total Visualizado": "$ {:,.2f}", 
                "Precio Unitario": "$ {:,.2f}",
                "Total CLP": "$ {:,.2f}", 
                "Total USD": "$ {:,.2f}", 
                "Total EUR": "$ {:,.2f}"
            })

            st.dataframe(styled_df_comp, height=350, use_container_width=True)

            col_c1, col_c2 = st.columns(2)
            with col_c1: st.metric("Total Ofertas Registradas", len(df_comp))
            with col_c2: st.metric(f"Monto Total Acumulado ({moneda_vista})", f"$ {df_comp['Monto Total Visualizado'].sum():,.2f}")
                
            st.divider()
            st.subheader("📈 Gráficos Comparativos por SOLPED")
            col_graf1, col_graf2 = st.columns(2)
            
            with col_graf1:
                st.markdown(f"**💰 Comparativa de Monto Total por SOLPED ({moneda_vista})**")
                st.bar_chart(df_comp.groupby("SOLPED")["Monto Total Visualizado"].sum().reset_index(), x="SOLPED", y="Monto Total Visualizado", height=250)
                
            with col_graf2:
                st.markdown("**⏳ Promedio Días de Entrega por SOLPED**")
                st.bar_chart(df_comp.groupby("SOLPED")["Días para Entrega"].mean().reset_index(), x="SOLPED", y="Días para Entrega", height=250)

            st.divider()
            st.subheader("📥 Exportar Reportes")
            
            col_down1, col_down2, _ = st.columns([1, 1, 2])
            with col_down1:
                if bytes_excel:
                    st.download_button(label="📊 Reporte Excel", data=bytes_excel, file_name=f"Reporte_Comparativo_{date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="dl_excel_bot")
            with col_down2:
                if bytes_pdf:
                    st.download_button(label="📄 Descargar Reporte PDF", data=bytes_pdf, file_name=f"Reporte_Comparativo_{date.today()}.pdf", mime="application/pdf", use_container_width=True, key="dl_pdf_bot")

        st.write("")
        if st.button("🗑️ Limpiar TODO el Cuadro Comparativo", key="btn_clear_bot"):
            st.session_state.ofertas_manuales = []
            st.rerun()
    else:
        st.info("Aún no hay ofertas registradas en el Cuadro Comparativo.")
