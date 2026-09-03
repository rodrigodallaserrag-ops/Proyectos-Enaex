import datetime
import io
import pandas as pd
import requests
import streamlit as st
import urllib3
import plotly.express as px

# Intentar importar FPDF para generación de PDF
try:
    from fpdf import FPDF
    PDF_HABILITADO = True
except ImportError:
    PDF_HABILITADO = False

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. TASAS DE CAMBIO EN TIEMPO REAL
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
                "estado": "Online 🟢",
            }
    except Exception:
        pass
    return {"dolar": 890.33, "euro": 1044.25, "uf": 39894.61, "estado": "Offline 🛡️"}

indicadores = obtener_indicadores_financieros()

# -----------------------------------------------------------------------------
# 2. FUNCIONES DE BASE DE DATOS Y PARSER DE PLANILLAS
# -----------------------------------------------------------------------------
SOLPEDS_DEMO = {
    "289283": {"desc": "DISCO RUPTURA GRAFITO 2`", "cant": 10, "um": "C/U", "precio": 58.00, "moneda": "USD", "prov": "MCM CHILE", "dias": 7, "comentario": "Proveedor con mejor lead time."},
    "1609": {"desc": "VALVULA DE BOLA 2 INCH ANSI 300", "cant": 5, "um": "C/U", "precio": 180000.0, "moneda": "CLP", "prov": "INDURA", "dias": 15, "comentario": "Precio estándar nacional."},
    "83723": {"desc": "ACEITE HIDRAULICO ISO 68", "cant": 200, "um": "LTS", "precio": 3.50, "moneda": "EUR", "prov": "TOTAL ENERGIES", "dias": 10, "comentario": "Importación directa."},
}

def generar_matriz_ejemplo():
    return pd.DataFrame([
        {"SP": "PR176577", "F. solicitud": "2026-07-27 16:07:00", "Dias de tratamiento": 3, "Material": "30066886", "Texto breve": "DISCO RUPTURA GRAFITO 2`", "Cantidad": 10, "UM": "C/U", "Proveedor Sugerido": "PRINTEC S A", "Monto Adjudicado": 545000},
        {"SP": "PR172030", "F. solicitud": "2026-07-02 11:43:00", "Dias de tratamiento": 28, "Material": "30051314", "Texto breve": "CONTADOR DIGITAL H", "Cantidad": 5, "UM": "C/U", "Proveedor Sugerido": "MCM CHILE", "Monto Adjudicado": 900000},
        {"SP": "PR172041", "F. solicitud": "2026-07-02 10:42:00", "Dias de tratamiento": 28, "Material": "20028618", "Texto breve": "JUEGO PERILLEROS", "Cantidad": 12, "UM": "SET", "Proveedor Sugerido": "PARKER", "Monto Adjudicado": 300000},
    ])

def cargar_planilla_autogestion(file):
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave = ['sp', 'solped', 'pr', 'código', 'descripción', 'cantidad', 'precio', 'proveedor', 'monto', 'material']
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

def buscar_solped_en_dataframe(id_sp, df):
    if df is None or df.empty:
        return None
    
    id_clean = str(id_sp).strip().upper()
    col_sp = None
    for col in df.columns:
        c_norm = str(col).strip().lower()
        if c_norm in ['sp', 'solped', 'solicitud', 'pr'] or 'solped' in c_norm or 'sp' in c_norm:
            col_sp = col
            break
            
    if not col_sp:
        return None

    coincidencias = df[df[col_sp].astype(str).str.strip().str.upper() == id_clean]
    if coincidencias.empty:
        return None
    
    row = coincidencias.iloc[0]

    def get_val(keys, default=""):
        for k in keys:
            for col in df.columns:
                if k in str(col).strip().lower():
                    val = row[col]
                    if pd.notna(val):
                        return val
        return default

    desc = get_val(['texto breve', 'descrip', 'material', 'item'], f"MATERIAL SOLPED #{id_clean}")
    cant = get_val(['cantidad', 'cant'], 1)
    um = str(get_val(['um', 'unidad'], "C/U")).upper()
    precio = get_val(['monto', 'precio', 'val', 'costo'], 100000.0)
    moneda = str(get_val(['moneda', 'curr'], "CLP")).upper()
    prov = get_val(['proveedor', 'vendor', 'prov'], "PROVEEDOR POR DEFINIR")
    dias = get_val(['dias de tratamiento', 'dias', 'plazo', 'lead'], 10)

    try: cant_num = int(float(cant))
    except Exception: cant_num = 1

    try: precio_num = float(precio)
    except Exception: precio_num = 100000.0

    try: dias_num = int(float(dias))
    except Exception: dias_num = 10

    moneda_val = moneda if moneda in ["CLP", "USD", "EUR"] else "CLP"
    um_val = um if um in ["C/U", "LTS", "MTR", "SET", "KG"] else "C/U"

    return {
        "desc": str(desc),
        "cant": max(1, cant_num),
        "um": um_val,
        "precio": max(0.0, precio_num),
        "moneda": moneda_val,
        "prov": str(prov),
        "dias": max(1, dias_num),
        "comentario": f"Autocompletado desde Planilla de Autogestión (SOLPED #{id_clean})"
    }

def exportar_reporte_pdf(id_solicitud, comprador, comentarios, df_data, total_monto):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"REPORTE DE ADJUDICACION - SOLPED #{id_solicitud}", ln=True, align='C')
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 8, f"Fecha: {datetime.date.today().strftime('%d/%m/%Y')} | Comprador: {comprador}", ln=True, align='C')
    pdf.ln(5)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "1. Resumen de Adjudicacion", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Total de Posiciones: {len(df_data)}", ln=True)
    pdf.cell(0, 6, f"Monto Total Adjudicado: CLP ${total_monto:,.0f}".replace(",", "."), ln=True)
    pdf.ln(4)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "2. Observaciones y Justificacion Tecnica", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.multi_cell(0, 6, comentarios)
    pdf.ln(6)

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "3. Detalle de Items Comparados", ln=True)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(30, 7, "SP / SOLPED", 1)
    pdf.cell(70, 7, "Descripcion", 1)
    pdf.cell(20, 7, "Cant.", 1)
    pdf.cell(40, 7, "Proveedor", 1)
    pdf.cell(30, 7, "Monto", 1, ln=True)

    pdf.set_font("Arial", '', 8)
    for _, row in df_data.iterrows():
        sp = str(row.get("SP", row.get("SP / SOLPED", "")))
        desc = str(row.get("Texto breve", row.get("Descripción", "")))[:30]
        cant = str(row.get("Cantidad", ""))
        prov = str(row.get("Proveedor Sugerido", row.get("Proveedor", "")))[:20]
        monto_val = row.get('Monto Adjudicado', row.get('Monto Total CLP', 0))
        monto = f"${float(monto_val):,.0f}".replace(",", ".")
        
        pdf.cell(30, 6, sp, 1)
        pdf.cell(70, 6, desc, 1)
        pdf.cell(20, 6, cant, 1)
        pdf.cell(40, 6, prov, 1)
        pdf.cell(30, 6, monto, 1, ln=True)

    return pdf.output(dest='S').encode('latin1')

# -----------------------------------------------------------------------------
# 3. ESTADO INICIAL DE SESIÓN Y CALLBACK
# -----------------------------------------------------------------------------
if "df_matriz_archivo" not in st.session_state:
    st.session_state["df_matriz_archivo"] = generar_matriz_ejemplo()

if "matriz_acumulada" not in st.session_state:
    st.session_state["matriz_acumulada"] = pd.DataFrame(columns=[
        "SP / SOLPED", "Descripción", "Cantidad", "UM", "Precio Unitario",
        "Moneda", "Precio Unit. CLP Norm.", "Proveedor",
        "Fecha Entrega", "Días Entrega", "Monto Total CLP", "Comentarios"
    ])

def cargar_solped_callback():
    sp_id = st.session_state.get("input_solped_id", "").strip()
    if sp_id:
        df_auto = st.session_state.get("df_matriz_archivo", None)
        datos = buscar_solped_en_dataframe(sp_id, df_auto)
        
        if not datos:
            id_clean = str(sp_id).strip().upper()
            if id_clean in SOLPEDS_DEMO:
                datos = SOLPEDS_DEMO[id_clean]
            else:
                datos = {
                    "desc": f"MATERIAL ASOCIADO A SOLPED #{id_clean}",
                    "cant": 1,
                    "um": "C/U",
                    "precio": 100000.0,
                    "moneda": "CLP",
                    "prov": "PROVEEDOR POR DEFINIR",
                    "dias": 10,
                    "comentario": "Cargado automáticamente. Modifique los campos necesarios."
                }

        st.session_state["f_desc"] = datos["desc"]
        st.session_state["f_cant"] = int(datos["cant"])
        st.session_state["f_um"] = datos["um"]
        st.session_state["f_precio"] = float(datos["precio"])
        st.session_state["f_moneda"] = datos["moneda"]
        st.session_state["f_prov"] = datos["prov"]
        st.session_state["f_dias"] = int(datos["dias"])
        st.session_state["f_comentario"] = datos["comentario"]

if "f_desc" not in st.session_state:
    st.session_state["input_solped_id"] = "PR176577"
    cargar_solped_callback()

# -----------------------------------------------------------------------------
# 4. LATERAL - INDICADORES
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("💱 Indicadores del Día")
    st.write(f"**USD:** ${indicadores['dolar']:,.2f}")
    st.write(f"**EUR:** ${indicadores['euro']:,.2f}")
    st.write(f"**UF:** ${indicadores['uf']:,.2f}")
    st.caption(f"Estado API: {indicadores['estado']}")

st.title("🛒 Consola de Compras y Autogestión - Enaex")

# =============================================================================
# MÓDULO 1: INGRESO MANUAL Y BÚSQUEDA INSTANTÁNEA POR SOLPED
# =============================================================================
st.header("✍️ Módulo 1: Carga Rápida y Comparador Manual por SOLPED")

st.markdown("**🔍 Búsqueda Rápida de SOLPED (Escriba cualquier ID de la planilla o externa)**")

col_b1, col_b2 = st.columns([3, 1])
with col_b1:
    st.text_input(
        "Ingrese ID SOLPED (ej: PR176577, PR172030, 289283):",
        key="input_solped_id",
        on_change=cargar_solped_callback
    )
with col_b2:
    st.write(" ")
    st.write(" ")
    if st.button("🔄 Cargar Información SOLPED", use_container_width=True):
        cargar_solped_callback()

# Formulario de revisión y edición manual
with st.container(border=True):
    st.subheader(f"⚙️ Detalle Cargado para SOLPED #{st.session_state.get('input_solped_id', '')}")
    
    c1, c2, c3 = st.columns([4, 1.5, 1.5])
    with c1:
        v_desc = st.text_input("Descripción Breve del Material", key="f_desc")
    with c2:
        v_cant = st.number_input("Cantidad", min_value=1, key="f_cant")
    with c3:
        v_um = st.selectbox("Unidad Medida", ["C/U", "LTS", "MTR", "SET", "KG"], key="f_um")

    c4, c5, c6, c7 = st.columns([2, 1.5, 2.5, 2])
    with c4:
        v_precio = st.number_input("Precio Unitario", min_value=0.0, step=100.0, key="f_precio")
    with c5:
        v_moneda = st.selectbox("Moneda", ["CLP", "USD", "EUR"], key="f_moneda")
    with c6:
        v_prov = st.text_input("Proveedor Oferente", key="f_prov")
    with c7:
        v_fecha = st.date_input(
            "Fecha Entrega (Calendario)", 
            value=datetime.date.today() + datetime.timedelta(days=st.session_state.get("f_dias", 10))
        )

    v_comentarios = st.text_input("Comentarios / Observaciones", key="f_comentario")

    if st.button("➕ Agregar esta SOLPED a la Matriz Comparativa Manual", type="primary", use_container_width=True):
        factor = 1.0
        if v_moneda == "USD":
            factor = indicadores["dolar"]
        elif v_moneda == "EUR":
            factor = indicadores["euro"]

        precio_clp_norm = int(v_precio * factor)
        monto_total_clp = int(v_cant * precio_clp_norm)
        dias_calculados = max(0, (v_fecha - datetime.date.today()).days)

        nueva_posicion = {
            "SP / SOLPED": str(st.session_state["input_solped_id"]),
            "Descripción": v_desc,
            "Cantidad": v_cant,
            "UM": v_um,
            "Precio Unitario": v_precio,
            "Moneda": v_moneda,
            "Precio Unit. CLP Norm.": precio_clp_norm,
            "Proveedor": v_prov,
            "Fecha Entrega": v_fecha,
            "Días Entrega": dias_calculados,
            "Monto Total CLP": monto_total_clp,
            "Comentarios": v_comentarios
        }

        df_act = st.session_state["matriz_acumulada"]
        st.session_state["matriz_acumulada"] = pd.concat([df_act, pd.DataFrame([nueva_posicion])], ignore_index=True)
        st.success(f"✅ SOLPED #{st.session_state['input_solped_id']} agregada a la matriz manual.")

# Tabla manual y gráficos
df_matriz_manual = st.session_state["matriz_acumulada"]

st.subheader(f"📊 Matriz Comparativa Multi-SOLPED Manual ({len(df_matriz_manual)} Registros)")

if df_matriz_manual.empty:
    st.info("💡 La matriz está vacía. Ingrese una SOLPED arriba y presione **'Agregar esta SOLPED a la Matriz Comparativa Manual'**.")
else:
    df_edited_manual = st.data_editor(
        df_matriz_manual,
        num_rows="dynamic",
        use_container_width=True,
        height=240,
        key="editor_manual"
    )
    st.session_state["matriz_acumulada"] = df_edited_manual

    st.subheader("📈 Análisis Comparativo Unitario (Carga Manual)")
    g1, g2 = st.columns(2)

    with g1:
        fig_precio = px.bar(
            df_edited_manual,
            x="SP / SOLPED",
            y="Precio Unit. CLP Norm.",
            color="Proveedor",
            text_auto=',.0f',
            title="💰 Precio Unitario Normalizado [CLP]",
            labels={"Precio Unit. CLP Norm.": "Precio Unit. (CLP)"}
        )
        fig_precio.update_layout(height=280)
        st.plotly_chart(fig_precio, use_container_width=True)

    with g2:
        fig_dias = px.bar(
            df_edited_manual,
            x="SP / SOLPED",
            y="Días Entrega",
            color="Proveedor",
            text_auto=True,
            title="⏱️ Días Prometidos de Entrega (Lead Time)",
            labels={"Días Entrega": "Días"}
        )
        fig_dias.update_layout(height=280)
        st.plotly_chart(fig_dias, use_container_width=True)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        out_excel_m = io.BytesIO()
        with pd.ExcelWriter(out_excel_m, engine='openpyxl') as writer:
            df_edited_manual.to_excel(writer, sheet_name='Comparativo_Manual', index=False)

        st.download_button(
            label="📥 Descargar Comparativo Manual en Excel",
            data=out_excel_m.getvalue(),
            file_name=f"Reporte_Manual_SOLPEDs_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    with col_m2:
        if st.button("🧹 Vaciar Matriz Manual", use_container_width=True):
            st.session_state["matriz_acumulada"] = pd.DataFrame(columns=df_matriz_manual.columns)
            st.rerun()

st.divider()

# =============================================================================
# MÓDULO 2: CARGA MASIVA Y PROCESAMIENTO AUTOMÁTICO DE PLANILLA (.XLSM)
# =============================================================================
st.header("📋 Módulo 2: Carga Masiva y Gráficos Automáticos (.XLSM / .XLSX / .CSV)")

uploaded_auto = st.file_uploader(
    "Subir Planilla de Autogestión Oficial", 
    type=["xlsx", "xlsm", "xls", "csv"],
    help="Cargue su archivo para visualizar la matriz completa, sus gráficos automáticos e informes."
)

if uploaded_auto is not None:
    df_cargado = cargar_planilla_autogestion(uploaded_auto)
    if df_cargado is not None and not df_cargado.empty:
        st.session_state["df_matriz_archivo"] = df_cargado
        st.success(f"✅ Archivo '{uploaded_auto.name}' cargado exitosamente.")

df_archivo = st.session_state["df_matriz_archivo"]

# Detección de columna de SOLPED en el archivo
col_solped = None
for c in df_archivo.columns:
    c_clean = str(c).strip().lower()
    if c_clean in ['sp', 'solped', 'solicitud', 'pr'] or 'solped' in c_clean or 'sp' in c_clean:
        col_solped = c
        break

lista_ids_archivo = ["Todas las solicitudes (SOLPED)"]
if col_solped:
    unicos = [str(x) for x in df_archivo[col_solped].dropna().unique() if str(x).strip() != ""]
    lista_ids_archivo.extend(unicos)

st.subheader("🆔 Filtros de Planilla y Control de Reporte")
col_f1, col_f2, col_f3 = st.columns([2, 2, 2])

with col_f1:
    id_sel_archivo = st.selectbox("Seleccionar SP / SOLPED del Archivo", lista_ids_archivo)
with col_f2:
    val_defecto = id_sel_archivo if id_sel_archivo != "Todas las solicitudes (SOLPED)" else "SOLPED-CONSOLIDADA"
    id_reporte_archivo = st.text_input("ID Personalizada para Reporte Masivo", value=val_defecto)
with col_f3:
    comprador_archivo = st.text_input("Comprador / Evaluador", value="Felipe Martínez", key="comp_archivo")

# Filtrado de la planilla subida
if id_sel_archivo != "Todas las solicitudes (SOLPED)" and col_solped:
    df_filtrado_archivo = df_archivo[df_archivo[col_solped].astype(str) == id_sel_archivo]
else:
    df_filtrado_archivo = df_archivo

st.subheader(f"📊 Matriz de Cotizaciones desde Planilla ({len(df_filtrado_archivo)} Posiciones)")

df_edited_archivo = st.data_editor(
    df_filtrado_archivo,
    num_rows="dynamic",
    use_container_width=True,
    height=280,
    key="editor_archivo"
)

# -----------------------------------------------------------------------------
# GRÁFICOS AUTOMÁTICOS DEL MÓDULO 2 (PLANILLA DE AUTOGESTIÓN)
# -----------------------------------------------------------------------------
st.subheader("📈 Gráficos Automáticos del Archivo Cargado")

cols_a = df_edited_archivo.columns
col_monto_auto = next((c for c in cols_a if any(kw in str(c).lower() for kw in ['monto', 'adjudicado', 'total', 'precio', 'val'])), None)
col_prov_auto = next((c for c in cols_a if any(kw in str(c).lower() for kw in ['proveedor', 'vendor', 'prov'])), None)
col_label_auto = col_solped if col_solped else next((c for c in cols_a if any(kw in str(c).lower() for kw in ['texto', 'desc', 'material', 'item'])), df_edited_archivo.columns[0])
col_dias_auto = next((c for c in cols_a if any(kw in str(c).lower() for kw in ['dias', 'plazo', 'lead', 'tratamiento'])), None)

if col_monto_auto:
    g_auto1, g_auto2 = st.columns(2)
    with g_auto1:
        fig_monto_auto = px.bar(
            df_edited_archivo,
            x=col_label_auto,
            y=col_monto_auto,
            color=col_prov_auto if col_prov_auto else None,
            text_auto=',.0f',
            title=f"💰 Monto por Posición ({col_monto_auto})",
            labels={col_monto_auto: "Monto", col_label_auto: "Item / SP"}
        )
        fig_monto_auto.update_layout(height=300)
        st.plotly_chart(fig_monto_auto, use_container_width=True)

    with g_auto2:
        if col_dias_auto:
            fig_dias_auto = px.bar(
                df_edited_archivo,
                x=col_label_auto,
                y=col_dias_auto,
                color=col_prov_auto if col_prov_auto else None,
                text_auto=True,
                title=f"⏱️ Lead Time / Días de Tratamiento ({col_dias_auto})",
                labels={col_dias_auto: "Días", col_label_auto: "Item / SP"}
            )
            fig_dias_auto.update_layout(height=300)
            st.plotly_chart(fig_dias_auto, use_container_width=True)
        elif col_prov_auto:
            df_prov_sum = df_edited_archivo.groupby(col_prov_auto)[col_monto_auto].sum().reset_index()
            fig_pie_auto = px.pie(
                df_prov_sum,
                names=col_prov_auto,
                values=col_monto_auto,
                title="📊 Distribución de Monto por Proveedor"
            )
            fig_pie_auto.update_layout(height=300)
            st.plotly_chart(fig_pie_auto, use_container_width=True)

# Dictamen y descarga
st.subheader("📝 Dictamen e Informe del Archivo Cargado")

comentarios_reporte_archivo = st.text_area(
    "Observaciones de adjudicación para el informe final:",
    value="Se realiza adjudicación considerando el menor costo total, cumplimiento de plazo de entrega (Lead Time) y la validación técnica favorable por parte del área usuaria."
)

if col_monto_auto:
    total_adjudicado = float(pd.to_numeric(df_edited_archivo[col_monto_auto], errors='coerce').fillna(0).sum())
else:
    total_adjudicado = 0.0

st.metric(f"Total Adjudicado SOLPED ({id_reporte_archivo})", f"$ {total_adjudicado:,.0f}".replace(",", "."))

col_a1, col_a2 = st.columns(2)

with col_a1:
    output_excel_a = io.BytesIO()
    with pd.ExcelWriter(output_excel_a, engine='openpyxl') as writer:
        df_edited_archivo.to_excel(writer, sheet_name=f'SOLPED_{id_reporte_archivo}'[:31], index=False)
    
    st.download_button(
        label="📥 Descargar Reporte Completo en Excel",
        data=output_excel_a.getvalue(),
        file_name=f"Reporte_Planilla_SOLPED_{id_reporte_archivo}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

with col_a2:
    if PDF_HABILITADO:
        try:
            pdf_bytes = exportar_reporte_pdf(id_reporte_archivo, comprador_archivo, comentarios_reporte_archivo, df_edited_archivo, total_adjudicado)
            st.download_button(
                label="📄 Descargar Reporte en PDF",
                data=pdf_bytes,
                file_name=f"Informe_Adjudicacion_SOLPED_{id_reporte_archivo}.pdf",
                mime="application/pdf",
                type="secondary",
                use_container_width=True
            )
        except Exception as e:
            st.warning(f"No se pudo generar el PDF: {e}")
    else:
        st.info("💡 Instale `fpdf` (`pip install fpdf`) para habilitar la exportación directa a PDF.")
