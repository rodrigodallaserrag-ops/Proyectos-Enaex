import datetime
import io
import pandas as pd
import requests
import streamlit as st
import urllib3
import plotly.express as px

try:
    from fpdf import FPDF
    PDF_HABILITADO = True
except ImportError:
    PDF_HABILITADO = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. TASAS DE CAMBIO
# -----------------------------------------------------------------------------
@st.cache_data(ttl=1800)
def obtener_indicadores():
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        r = requests.get("https://mindicador.cl/api", headers=headers, verify=False, timeout=5)
        if r.status_code == 200:
            data = r.json()
            return {"USD": float(data["dolar"]["valor"]), "EUR": float(data["euro"]["valor"]), "CLP": 1.0}
    except Exception:
        pass
    return {"USD": 890.33, "EUR": 1044.25, "CLP": 1.0}

tasas = obtener_indicadores()

# -----------------------------------------------------------------------------
# 2. CARGADOR INTELIGENTE DE EXCEL (DETECTA LA FILA REAL DE ENCABEZADOS)
# -----------------------------------------------------------------------------
def cargar_planilla_inteligente(file):
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave = ['sp', 'solped', 'pr', 'código', 'descripción', 'cantidad', 'precio', 'proveedor', 'monto', 'material', 'texto breve']
            mejor_sheet, mejor_fila_idx, max_coincidencias = None, None, 0

            # Buscar la fila que tenga más coincidencias con columnas de compras
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

            # Extraer tabla omitiendo las filas de títulos superiores
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
            "Material": get_val(['texto', 'desc', 'material', 'item'], f"Material ID {id_solped}"),
            "Cantidad": float(get_val(['cant', 'cantidad'], 1.0)),
            "UM": str(get_val(['um', 'unidad'], "C/U")).upper(),
            "Precio Unitario": float(get_val(['precio', 'monto', 'val', 'costo'], 0.0)),
            "Moneda": str(get_val(['moneda', 'curr'], "CLP")).upper(),
            "Proveedor": str(get_val(['proveedor', 'vendor', 'prov'], "POR DEFINIR")),
            "Días Entrega": int(float(get_val(['dias', 'plazo', 'lead', 'tratamiento'], 10))),
            "Comentario": "Autocompletado desde planilla masiva"
        })
    return posiciones

SOLPEDS_BASE = {
    "287": [
        {"Pos": 1, "Material": "EXPLOSIVOS HIGH POWER", "Cantidad": 20, "UM": "C/U", "Precio Unitario": 8.0, "Moneda": "CLP", "Proveedor": "EXPLOSIVOS CHILE", "Días Entrega": 5, "Comentario": "⚠️ Precio de 8 pesos (corregir valor)"},
        {"Pos": 2, "Material": "MATERIAS PRIMAS QUIMICAS", "Cantidad": 70, "UM": "KG", "Precio Unitario": 12.5, "Moneda": "USD", "Proveedor": "CHEMICAL CORP", "Días Entrega": 14, "Comentario": "Importación directa"},
        {"Pos": 3, "Material": "TORNILLOS 200 KILOS", "Cantidad": 2, "UM": "C/U", "Precio Unitario": 45000.0, "Moneda": "CLP", "Proveedor": "PERNOS S.A.", "Días Entrega": 3, "Comentario": "Stock local disponible"},
    ]
}

def generar_pdf(id_solped, comprador, df_data, total_monto):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"REPORTE DE ADJUDICACION - SOLPED #{id_solped}", ln=True, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, f"Fecha: {datetime.date.today().strftime('%d/%m/%Y')} | Comprador: {comprador}", ln=True, align='C')
    pdf.ln(5)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "Resumen de Evaluación", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Total Items evaluados: {len(df_data)}", ln=True)
    pdf.cell(0, 6, f"Monto Total Evaluado (CLP): ${total_monto:,.0f}".replace(",", "."), ln=True)
    pdf.ln(4)

    pdf.set_font("Arial", 'B', 9)
    pdf.cell(10, 7, "Pos", 1)
    pdf.cell(70, 7, "Material / Descripcion", 1)
    pdf.cell(15, 7, "Cant", 1)
    pdf.cell(45, 7, "Proveedor", 1)
    pdf.cell(15, 7, "Dias", 1)
    pdf.cell(35, 7, "Total CLP", 1, ln=True)

    pdf.set_font("Arial", '', 8)
    for _, row in df_data.iterrows():
        pdf.cell(10, 6, str(row.get("Pos", "")), 1)
        pdf.cell(70, 6, str(row.get("Material", ""))[:32], 1)
        pdf.cell(15, 6, str(row.get("Cantidad", "")), 1)
        pdf.cell(45, 6, str(row.get("Proveedor", ""))[:22], 1)
        pdf.cell(15, 6, str(row.get("Días Entrega", "")), 1)
        monto_clp = float(row.get("Monto Total CLP", 0))
        pdf.cell(35, 6, f"${monto_clp:,.0f}".replace(",", "."), 1, ln=True)

    return pdf.output(dest='S').encode('latin1')

# -----------------------------------------------------------------------------
# INTERFAZ Streamlit
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("💱 Tipo de Cambio")
    st.write(f"**USD:** ${tasas['USD']:,.2f}")
    st.write(f"**EUR:** ${tasas['EUR']:,.2f}")

st.title("🛒 Consola de Compras - Autogestión Integral")

# MÓDULO 1: CARGA MASIVA AUTOMÁTICA
st.header("📋 Módulo 1: Base de Datos Maestra (Carga Masiva)")
uploaded_file = st.file_uploader("Suba su Planilla de Autogestión (Excel/CSV) para nutrir el sistema", type=["xlsx", "xlsm", "xls", "csv"])

if uploaded_file is not None:
    df_masivo = cargar_planilla_inteligente(uploaded_file)
    if df_masivo is not None and not df_masivo.empty:
        st.session_state["df_masivo"] = df_masivo
        st.success(f"✅ Archivo cargado e interpretado correctamente ({len(df_masivo)} filas detectadas).")
        with st.expander("Ver Vista Previa de la Tabla Detectada"):
            st.dataframe(df_masivo.head(10), use_container_width=True)

st.divider()

# MÓDULO 2: BÚSQUEDA Y EDICIÓN MULTI-POSICIÓN
st.header("✏️ Módulo 2: Edición y Evaluación por SOLPED")
st.markdown("Busque una ID. El sistema extraerá **todos sus materiales** desde la base maestra para corregir precios (ej. montos en 8 pesos), ajustar monedas o cambiar proveedores.")

col_search, col_btn = st.columns([3, 1])
with col_search:
    solped_input = st.text_input("Ingrese ID SOLPED (Ejemplo: 287):", value="287")
with col_btn:
    st.write(" ")
    st.write(" ")
    cargar_btn = st.button("📥 Cargar Materiales", use_container_width=True, type="primary")

if "datos_solped_actual" not in st.session_state or cargar_btn:
    sp_id = solped_input.strip()
    materiales_cargados = []
    
    if "df_masivo" in st.session_state:
        materiales_cargados = extraer_materiales_de_masivo(st.session_state["df_masivo"], sp_id)
    
    if not materiales_cargados:
        if sp_id in SOLPEDS_BASE:
            materiales_cargados = SOLPEDS_BASE[sp_id]
        else:
            materiales_cargados = [
                {"Pos": 1, "Material": f"NUEVO MATERIAL - SOLPED {sp_id}", "Cantidad": 1, "UM": "C/U", "Precio Unitario": 0.0, "Moneda": "CLP", "Proveedor": "", "Días Entrega": 1, "Comentario": "Agregue datos reales"}
            ]
            
    st.session_state["datos_solped_actual"] = pd.DataFrame(materiales_cargados)
    st.session_state["solped_id_cargada"] = sp_id

df_trabajo = st.session_state["datos_solped_actual"]

st.subheader(f"🛠️ Ajuste de Materiales - SOLPED #{st.session_state.get('solped_id_cargada', '')}")
df_editado = st.data_editor(
    df_trabajo,
    num_rows="dynamic",
    use_container_width=True,
    height=240,
    column_config={
        "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "EUR"], default="CLP"),
        "UM": st.column_config.SelectboxColumn("UM", options=["C/U", "KG", "LTS", "SET", "MTR"], default="C/U"),
        "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
        "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=0.1),
        "Días Entrega": st.column_config.NumberColumn("Días Entrega", min_value=0)
    },
    key="editor_multi_posicion"
)

def calcular_monto_clp(row):
    try:
        precio = float(row.get("Precio Unitario", 0))
        cant = float(row.get("Cantidad", 1))
        moneda = str(row.get("Moneda", "CLP")).upper()
        factor = tasas.get(moneda, 1.0)
        return int(precio * cant * factor)
    except Exception:
        return 0

df_editado["Monto Total CLP"] = df_editado.apply(calcular_monto_clp, axis=1)

# MÓDULO 3: GRÁFICOS Y EXPORTACIÓN
st.divider()
st.subheader("📈 Análisis de Ofertas y Lead Time")

col_g1, col_g2 = st.columns(2)
with col_g1:
    fig_costo = px.bar(
        df_editado, x="Material", y="Monto Total CLP", color="Proveedor",
        text_auto=',.0f', title="💰 Monto Total Normalizado [CLP]"
    )
    fig_costo.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_costo, use_container_width=True)

with col_g2:
    fig_dias = px.bar(
        df_editado, x="Material", y="Días Entrega", color="Proveedor",
        text_auto=True, title="⏱️ Días Prometidos de Entrega (Lead Time)"
    )
    fig_dias.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_dias, use_container_width=True)

st.subheader("📥 Exportar Informe Final")
c_rep1, c_rep2 = st.columns(2)
with c_rep1:
    comprador_nombre = st.text_input("Nombre del Comprador:", value="Felipe Martínez")
with c_rep2:
    id_sp_reporte = st.session_state.get('solped_id_cargada', '287')
    total_evaluado = df_editado["Monto Total CLP"].sum()
    st.metric("Monto Total Acumulado", f"$ {total_evaluado:,.0f}".replace(",", "."))

col_exp1, col_exp2 = st.columns(2)
with col_exp1:
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df_editado.to_excel(writer, sheet_name=f'SOLPED_{id_sp_reporte}'[:31], index=False)
    st.download_button(
        label="📥 Descargar Reporte Excel",
        data=out_excel.getvalue(),
        file_name=f"Evaluacion_SOLPED_{id_sp_reporte}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary"
    )

with col_exp2:
    if PDF_HABILITADO:
        try:
            bytes_pdf = generar_pdf(id_sp_reporte, comprador_nombre, df_editado, total_evaluado)
            st.download_button(
                label="📄 Descargar Reporte PDF",
                data=bytes_pdf,
                file_name=f"Informe_SOLPED_{id_sp_reporte}.pdf",
                mime="application/pdf",
                use_container_width=True, type="secondary"
            )
        except Exception as err:
            st.warning(f"Error al generar PDF: {err}")
    else:
        st.info("Instale `fpdf` para habilitar la descarga en PDF.")
