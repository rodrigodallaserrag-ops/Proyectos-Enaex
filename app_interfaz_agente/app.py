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

# Intentar importar FPDF para generación de PDF
try:
    from fpdf import FPDF
    PDF_HABILITADO = True
except ImportError:
    PDF_HABILITADO = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. MOTOR FINANCIERO DE MONEDAS
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1800)
def obtener_indicadores_financieros():
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        r = requests.get("https://mindicador.cl/api", headers=headers, verify=False, timeout=5)
        if r.status_code == 200:
            data = r.json()
            return {
                "dolar": float(data["dolar"]["valor"]),
                "euro": float(data["euro"]["valor"]),
                "uf": float(data["uf"]["valor"]),
                "fecha": data["dolar"]["fecha"][:10],
                "estado": "Online 🟢",
            }
    except Exception:
        pass
    return {"dolar": 890.33, "euro": 1044.25, "uf": 39894.61, "fecha": datetime.date.today().strftime("%d-%m-%Y"), "estado": "Offline 🛡️"}

indicadores = obtener_indicadores_financieros()

# -----------------------------------------------------------------------------
# 2. CARGA INTELIGENTE DE PLANILLA DE AUTOGESTIÓN (.XLSM)
# -----------------------------------------------------------------------------
def cargar_planilla_autogestion(file):
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave_fuertes = [
                'solped', 'código', 'codigo', 'descripción', 'descripcion', 
                'cantidad', 'um', 'precio', 'proveedor', 'oferta', 
                'incoterm', 'adjudicado', 'monto', 'lead time', 'material', 'posición'
            ]
            
            mejor_sheet, mejor_fila_idx, max_coincidencias = None, None, 0

            for sheet_name in excel_file.sheet_names:
                try:
                    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                except Exception:
                    continue

                if df_raw.empty or len(df_raw) < 2:
                    continue

                for idx in range(min(40, len(df_raw))):
                    row = df_raw.iloc[idx]
                    celdas_texto = [str(val).strip().lower() for val in row.values if pd.notna(val) and str(val).strip() != '']
                    if len(celdas_texto) < 3:
                        continue
                    coincidencias = sum(1 for kw in palabras_clave_fuertes if any(kw in celda for celda in celdas_texto))
                    if coincidencias > max_coincidencias:
                        max_coincidencias = coincidencias
                        mejor_sheet = sheet_name
                        mejor_fila_idx = idx

            if mejor_sheet is not None and max_coincidencias >= 2:
                df_raw = pd.read_excel(excel_file, sheet_name=mejor_sheet, header=None)
                df_clean = df_raw.iloc[mejor_fila_idx:].copy()
                
                encabezados = []
                for i, col_val in enumerate(df_clean.iloc[0]):
                    val_str = str(col_val).strip() if pd.notna(col_val) else ""
                    if val_str != "" and val_str.lower() not in ["none", "nan", "unnamed"]:
                        encabezados.append(val_str)
                    else:
                        encabezados.append(f"Columna_{i+1}")
                
                df_clean.columns = encabezados
                df_clean = df_clean.iloc[1:].reset_index(drop=True).dropna(how='all')
                cols_utiles = [col for col in df_clean.columns if not str(col).startswith("Columna_")]
                if len(cols_utiles) >= 2:
                    df_clean = df_clean[cols_utiles]
                return df_clean.dropna(thresh=2)

            return pd.read_excel(excel_file, sheet_name=0).dropna(how='all').dropna(how='all', axis=1)

        elif nombre.endswith('.csv'):
            try:
                df = pd.read_csv(file_bytes, sep=None, engine='python', encoding='utf-8')
            except Exception:
                file_bytes.seek(0)
                df = pd.read_csv(file_bytes, sep=None, engine='python', encoding='latin-1')
            return df.dropna(how='all').dropna(how='all', axis=1)

    except Exception as e:
        st.error(f"Error al procesar la planilla ({file.name}): {e}")
        return None

def generar_matriz_ejemplo():
    data = []
    materiales_demo = [
        ("LIMPIA CONTACTOS 279 CHESTERTON.", "C/U", 72, 54500, "CLP", "PRINTEC S A"),
        ("ACEITE HIDRAULICO ISO 68", "LTS", 500, 3200, "CLP", "LUBRICANTES CHILE"),
        ("VALVULA DE BOLA 2 INCH ANSI 300", "C/U", 15, 180000, "CLP", "MCM CHILE"),
        ("MANGUERA ALTA PRESION 1/2", "MTR", 120, 25000, "CLP", "PARKER"),
        ("EMPADRON BASTIDOR SOPORTE", "C/U", 4, 450000, "CLP", "INDURA"),
    ]
    for i in range(1, 16):
        idx = (i - 1) % len(materiales_demo)
        desc, um, cant, precio, mon, prov = materiales_demo[idx]
        solped_id = "1002610200" if i <= 8 else "1002610300"
        data.append({
            "Pos": i * 10,
            "Solped": solped_id,
            "Código SAP": f"2000{6120 + i}",
            "Descripción breve": f"{desc} #{i}",
            "Cantidad": cant,
            "UM": um,
            "Última compra": "19-06-2025",
            "Proveedor Histórico": prov,
            "Último Precio [Unit]": precio,
            "Moneda Hist.": mon,
            "Oferta Inicial Moneda": precio * 1.05,
            "Oferta Mejorada Moneda": precio * 0.95,
            "Lead Time": "5 DIAS",
            "Validación Técnica": True,
            "Proveedor Sugerido": prov,
            "Monto Adjudicado": int(cant * precio * 0.95),
        })
    return pd.DataFrame(data)

# GENERADOR DE REPORTES PDF
def exportar_reporte_pdf(id_solicitud, comprador, comentarios, df_data, total_monto):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"REPORTE DE ADJUDICACIÓN - SOLICITUD #{id_solicitud}", ln=True, align='C')
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, f"Fecha: {datetime.date.today().strftime('%d/%m/%Y')} | Comprador: {comprador}", ln=True, align='C')
    pdf.ln(5)

    # Resumen Ejecutivo
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "1. Resumen de Adjudicación", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Total de Posiciones: {len(df_data)}", ln=True)
    pdf.cell(0, 6, f"Monto Total Adjudicado: CLP ${total_monto:,.0f}".replace(",", "."), ln=True)
    pdf.ln(4)

    # Comentarios / Justificación
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "2. Observaciones y Justificación Técnica/Económica", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.multi_cell(0, 6, comentarios)
    pdf.ln(6)

    # Detalle de Materiales (Tabla Resumida)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "3. Detalle de Ítems Comparados", ln=True)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 7, "Pos", 1)
    pdf.cell(75, 7, "Descripción", 1)
    pdf.cell(20, 7, "Cant.", 1)
    pdf.cell(40, 7, "Proveedor", 1)
    pdf.cell(35, 7, "Monto", 1, ln=True)

    pdf.set_font("Arial", '', 8)
    for _, row in df_data.iterrows():
        pos = str(row.get("Pos", ""))
        desc = str(row.get("Descripción breve", row.get("Descripción", "")))[:35]
        cant = str(row.get("Cantidad", ""))
        prov = str(row.get("Proveedor Sugerido", row.get("Proveedor Histórico", "")))[:20]
        monto = f"${float(row.get('Monto Adjudicado', 0)):,.0f}".replace(",", ".")
        
        pdf.cell(20, 6, pos, 1)
        pdf.cell(75, 6, desc, 1)
        pdf.cell(20, 6, cant, 1)
        pdf.cell(40, 6, prov, 1)
        pdf.cell(35, 6, monto, 1, ln=True)

    return pdf.output(dest='S').encode('latin1')

# -----------------------------------------------------------------------------
# 3. INTERFAZ LATERAL
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("📋 Planilla de Autogestión")
    uploaded_auto = st.file_uploader(
        "Subir Planilla de Autogestión (Excel/XLSM/CSV)", 
        type=["xlsx", "xlsm", "xls", "csv"]
    )
    st.divider()
    st.header("💱 Indicadores del Día")
    st.write(f"**USD:** ${indicadores['dolar']:,.2f}")
    st.write(f"**EUR:** ${indicadores['euro']:,.2f}")
    st.caption(f"Estado API: {indicadores['estado']}")

# -----------------------------------------------------------------------------
# 4. PANEL DE CONTROL DE REPORTE E ID DE SOLICITUD
# -----------------------------------------------------------------------------
st.title("🛒 Cuadro Comparativo Multimaterial - Autogestión")

if uploaded_auto is not None:
    df_cargado = cargar_planilla_autogestion(uploaded_auto)
    if df_cargado is not None and not df_cargado.empty:
        st.session_state["df_matriz"] = df_cargado

if "df_matriz" not in st.session_state:
    st.session_state["df_matriz"] = generar_matriz_ejemplo()

df_matriz = st.session_state["df_matriz"]

# Búsqueda/Detección de ID Solicitud
col_solped = [c for c in df_matriz.columns if 'solped' in str(c).lower() or 'id' in str(c).lower()]
lista_ids = ["Todas las solicitudes"]
if col_solped:
    unicos = df_matriz[col_solped[0]].dropna().astype(str).unique().tolist()
    lista_ids.extend(unicos)

st.subheader("🆔 Control de ID de Solicitud y Filtros")
col_f1, col_f2, col_f3 = st.columns([2, 2, 2])

with col_f1:
    id_seleccionado = st.selectbox("Seleccionar ID Solicitud (SOLPED)", lista_ids)
with col_f2:
    id_reporte = st.text_input("ID Personalizada para Reporte", value=id_seleccionado if id_seleccionado != "Todas las solicitudes" else "SOL-200")
with col_f3:
    comprador = st.text_input("Comprador / Evaluador", value="Felipe Martínez")

# Filtrado dinámico
if id_seleccionado != "Todas las solicitudes" and col_solped:
    df_filtrado = df_matriz[df_matriz[col_solped[0]].astype(str) == id_seleccionado]
else:
    df_filtrado = df_matriz

# -----------------------------------------------------------------------------
# 5. MATRIZ DE COTIZACIONES EDITABLE
# -----------------------------------------------------------------------------
st.subheader(f"📊 Matriz de Cotizaciones ({len(df_filtrado)} Posiciones)")

df_edited = st.data_editor(
    df_filtrado,
    num_rows="dynamic",
    use_container_width=True,
    height=400
)

# -----------------------------------------------------------------------------
# 6. COMENTARIOS EJECUTIVOS Y GENERACIÓN DE INFORMES
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📝 Dictamen y Comentarios del Reporte")

comentarios_reporte = st.text_area(
    "Ingrese las observaciones de adjudicación para el informe final:",
    value="Se realiza adjudicación considerando el menor costo total, cumplimiento de plazo de entrega (Lead Time) y la validación técnica favorable por parte del área usuaria."
)

col_monto = [col for col in df_edited.columns if 'monto' in str(col).lower() or 'adjudicado' in str(col).lower()]
total_adjudicado = float(pd.to_numeric(df_edited[col_monto[0]], errors='coerce').fillna(0).sum()) if col_monto else 0.0

st.metric("Total Adjudicado para esta Solicitud", f"$ {total_adjudicado:,.0f}".replace(",", "."))

col_d1, col_d2 = st.columns(2)

with col_d1:
    # Descarga Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_edited.to_excel(writer, sheet_name=f'Solicitud_{id_reporte}', index=False)
    
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=output_excel.getvalue(),
        file_name=f"Reporte_Solicitud_{id_reporte}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

with col_d2:
    # Descarga PDF
    if PDF_HABILITADO:
        try:
            pdf_bytes = exportar_reporte_pdf(id_reporte, comprador, comentarios_reporte, df_edited, total_adjudicado)
            st.download_button(
                label="📄 Descargar Reporte en PDF",
                data=pdf_bytes,
                file_name=f"Informe_Adjudicacion_{id_reporte}.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )
        except Exception as e:
            st.warning(f"No se pudo generar PDF: {e}")
    else:
        st.info("💡 Instala `fpdf` (`pip install fpdf`) para activar la descarga directa en PDF.")
