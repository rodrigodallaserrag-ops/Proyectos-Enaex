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
# 2. CARGADOR INTELIGENTE DE EXCEL
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
            "Comentario": "(Comentario)"
        })
    return posiciones

# Plantilla base por defecto con casillas tipo plantilla
SOLPEDS_BASE = {
    "(ID SOLPED)": [
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
            "Comentario": "(Comentario)"
        }
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

    pdf.set_font("Arial", 'B', 8)
    pdf.cell(8, 7, "Pos", 1)
    pdf.cell(50, 7, "Material", 1)
    pdf.cell(20, 7, "Centro", 1)
    pdf.cell(12, 7, "Cant", 1)
    pdf.cell(35, 7, "Proveedor", 1)
    pdf.cell(35, 7, "Cal. Entrega", 1)
    pdf.cell(30, 7, "Total CLP", 1, ln=True)

    pdf.set_font("Arial", '', 7)
    for _, row in df_data.iterrows():
        pdf.cell(8, 6, str(row.get("Pos", "")), 1)
        pdf.cell(50, 6, str(row.get("Material", ""))[:25], 1)
        pdf.cell(20, 6, str(row.get("Centro", ""))[:10], 1)
        pdf.cell(12, 6, str(row.get("Cantidad", "")), 1)
        pdf.cell(35, 6, str(row.get("Proveedor", ""))[:18], 1)
        pdf.cell(35, 6, str(row.get("Calendario de entrega", ""))[:18], 1)
        monto_clp = float(row.get("Monto Total CLP", 0))
        pdf.cell(30, 6, f"${monto_clp:,.0f}".replace(",", "."), 1, ln=True)

    return pdf.output(dest='S').encode('latin1')

# -----------------------------------------------------------------------------
# INTERFAZ STREAMLIT
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("💱 Tipo de Cambio")
    st.write(f"**USD:** ${tasas['USD']:,.2f}")
    st.write(f"**EUR:** ${tasas['EUR']:,.2f}")

st.title("🛒 Consola de Compras - Autogestión Integral")

# =============================================================================
# MÓDULO 1: CARGA MASIVA DE REQUERIMIENTOS DIARIOS
# =============================================================================
st.header("📋 Módulo 1: Base de Datos Maestra (Requerimientos Diarios)")
uploaded_file = st.file_uploader("Suba la Planilla de Autogestión Diaria (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="masivo_diario")

if uploaded_file is not None:
    df_masivo = cargar_planilla_inteligente(uploaded_file)
    if df_masivo is not None and not df_masivo.empty:
        st.session_state["df_masivo"] = df_masivo
        st.success(f"✅ Archivo de requerimientos cargado ({len(df_masivo)} filas detectadas).")
        with st.expander("Ver Vista Previa de la Tabla Detectada"):
            st.dataframe(df_masivo.head(10), use_container_width=True)

st.divider()

# =============================================================================
# MÓDULO 2: BÚSQUEDA HISTÓRICA COMPLETA (SIN LÍMITE DE FECHA - PREVIO A 2019)
# =============================================================================
st.header("📜 Módulo 2: Historico de OC (Consulta Completa sin Filtro de Fecha)")
st.markdown("Cargue el archivo histórico maestro de Órdenes de Compra para consultar antecedentes previos a 2018/2019.")

uploaded_hist = st.file_uploader("Suba la Base Histórica Completa de Órdenes de Compra (Excel/CSV)", type=["xlsx", "xlsm", "xls", "csv"], key="historico_oc")

if uploaded_hist is not None:
    df_historico = cargar_planilla_inteligente(uploaded_hist)
    if df_historico is not None and not df_historico.empty:
        st.session_state["df_historico"] = df_historico
        st.success(f"✅ Histórico cargado con éxito ({len(df_historico)} registros cargados sin restricción de años).")

if "df_historico" in st.session_state and st.session_state["df_historico"] is not None:
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        busqueda_hist = st.text_input("Buscar en histórico (Ej: Código de Material, Texto, Proveedor, Año):", value="", placeholder="(Material) / (Proveedor) / (Año)")
    
    if busqueda_hist.strip():
        df_h = st.session_state["df_historico"]
        mask = df_h.astype(str).apply(lambda col: col.str.contains(busqueda_hist.strip(), case=False, na=False)).any(axis=1)
        res_hist = df_h[mask]
        st.markdown(f"**Resultados encontrados ({len(res_hist)}):**")
        st.dataframe(res_hist, use_container_width=True)
else:
    st.info("💡 Suba una base histórica para habilitar la consulta de precios anteriores a 2019.")

st.divider()

# =============================================================================
# MÓDULO 3: COMPARATIVO Y EDICIÓN MULTI-POSICIÓN (SOPORTA +20 MATERIALES)
# =============================================================================
st.header("✏️ Módulo 3: Comparativo Dinámico por SOLPED")
st.markdown("Extraiga **todos los materiales** de una SOLPED para evaluar ofertas sin límites de filas (Soporta 20+ posiciones).")

col_search, col_btn = st.columns([3, 1])
with col_search:
    solped_input = st.text_input("Ingrese ID SOLPED:", value="", placeholder="(ID SOLPED)")
with col_btn:
    st.write(" ")
    st.write(" ")
    cargar_btn = st.button("📥 Cargar Materiales", use_container_width=True, type="primary")

if "datos_solped_actual" not in st.session_state or cargar_btn:
    sp_id = solped_input.strip() if solped_input.strip() else "(ID SOLPED)"
    materiales_cargados = []
    
    if "df_masivo" in st.session_state and solped_input.strip():
        materiales_cargados = extraer_materiales_de_masivo(st.session_state["df_masivo"], sp_id)
    
    if not materiales_cargados:
        if sp_id in SOLPEDS_BASE:
            materiales_cargados = SOLPEDS_BASE[sp_id]
        else:
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
                    "Comentario": "(Comentario)"
                }
            ]
            
    st.session_state["datos_solped_actual"] = pd.DataFrame(materiales_cargados)
    st.session_state["solped_id_cargada"] = sp_id

df_trabajo = st.session_state["datos_solped_actual"]

st.subheader(f"🛠️ Evaluación e Inserción de Datos - SOLPED #{st.session_state.get('solped_id_cargada', '')}")
df_editado = st.data_editor(
    df_trabajo,
    num_rows="dynamic",
    use_container_width=True,
    height=300,
    column_config={
        "Pos": st.column_config.NumberColumn("Pos", disabled=False),
        "Material": st.column_config.TextColumn("Material"),
        "Centro": st.column_config.TextColumn("Centro"),
        "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=0.01, format="%.2f"),
        "UM": st.column_config.SelectboxColumn("UM", options=["C/U", "KG", "LTS", "SET", "MTR", "TON"], default="C/U"),
        "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
        "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "EUR"], default="CLP"),
        "Proveedor": st.column_config.TextColumn("Proveedor"),
        "Calendario de entrega": st.column_config.TextColumn("Calendario de entrega"),
        "Días Entrega": st.column_config.NumberColumn("Días Entrega", min_value=0),
        "Comentario": st.column_config.TextColumn("Comentario")
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

# =============================================================================
# MÓDULO 4: GRÁFICOS Y EXPORTACIÓN DEL REPORTE
# =============================================================================
st.divider()
st.subheader("📈 Análisis de Ofertas y Tiempos de Entrega")

col_g1, col_g2 = st.columns(2)
with col_g1:
    fig_costo = px.bar(
        df_editado, x="Material", y="Monto Total CLP", color="Proveedor",
        text_auto=',.0f', title="💰 Monto Total Evaluado [CLP]"
    )
    fig_costo.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_costo, use_container_width=True)

with col_g2:
    fig_dias = px.bar(
        df_editado, x="Material", y="Días Entrega", color="Proveedor",
        text_auto=True, title="⏱️ Días Prometidos de Entrega"
    )
    fig_dias.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_dias, use_container_width=True)

st.subheader("📥 Exportar Reporte Comparativo")
c_rep1, c_rep2 = st.columns(2)
with c_rep1:
    comprador_nombre = st.text_input("Nombre del Comprador:", value="", placeholder="(Nombre del Comprador)")
with c_rep2:
    id_sp_reporte = st.session_state.get('solped_id_cargada', '(ID SOLPED)')
    total_evaluado = df_editado["Monto Total CLP"].sum()
    st.metric("Monto Total Acumulado Evaluado", f"$ {total_evaluado:,.0f}".replace(",", "."))

col_exp1, col_exp2 = st.columns(2)
with col_exp1:
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df_editado.to_excel(writer, sheet_name='Comparativo', index=False)
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
            bytes_pdf = generar_pdf(id_sp_reporte, comprador_nombre if comprador_nombre else "(Comprador)", df_editado, total_evaluado)
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
        st.info("Instale la librería `fpdf` para activar la exportación a PDF.")
