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
# 2. CARGA DE PLANILLA Y DATOS DE EJEMPLO
# -----------------------------------------------------------------------------
def cargar_planilla_autogestion(file):
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave = ['sp', 'solped', 'pr', 'código', 'descripción', 'cantidad', 'precio', 'proveedor', 'oferta', 'monto']
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

    except Exception as e:
        st.error(f"Error al procesar la planilla: {e}")
        return None

def generar_matriz_ejemplo():
    hoy = datetime.date.today()
    return pd.DataFrame([
        {
            "SP": "PR176577",
            "Texto breve": "DISCO RUPTURA GRAFITO 2`",
            "Cantidad": 10,
            "UM": "C/U",
            "Precio Unitario": 54500,
            "Moneda": "CLP",
            "Precio Normalizado [CLP]": 54500,
            "Proveedor": "PRINTEC S A",
            "Fecha Entrega": hoy + datetime.timedelta(days=15),
            "Días Entrega": 15,
            "Monto Total [CLP]": 545000,
            "Comentarios": "Proveedor histórico con pronta entrega."
        },
        {
            "SP": "PR176577",
            "Texto breve": "DISCO RUPTURA GRAFITO 2`",
            "Cantidad": 10,
            "UM": "C/U",
            "Precio Unitario": 58.00,
            "Moneda": "USD",
            "Precio Normalizado [CLP]": int(58.00 * indicadores["dolar"]),
            "Proveedor": "MCM CHILE",
            "Fecha Entrega": hoy + datetime.timedelta(days=7),
            "Días Entrega": 7,
            "Monto Total [CLP]": int(10 * 58.00 * indicadores["dolar"]),
            "Comentarios": "Oferta en USD. Entrega más rápida."
        },
        {
            "SP": "PR172030",
            "Texto breve": "CONTADOR DIGITAL H",
            "Cantidad": 5,
            "UM": "C/U",
            "Precio Unitario": 180000,
            "Moneda": "CLP",
            "Precio Normalizado [CLP]": 180000,
            "Proveedor": "INDURA",
            "Fecha Entrega": hoy + datetime.timedelta(days=20),
            "Días Entrega": 20,
            "Monto Total [CLP]": 900000,
            "Comentarios": "Opción estándar en stock."
        }
    ])

# -----------------------------------------------------------------------------
# 3. BARRA LATERAL
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("📋 Carga Masiva (Excel/XLSM)")
    uploaded_auto = st.file_uploader("Subir Planilla Autogestión", type=["xlsx", "xlsm", "xls", "csv"])
    st.divider()
    st.header("💱 Tasas de Cambio")
    st.write(f"**USD:** ${indicadores['dolar']:,.2f}")
    st.write(f"**EUR:** ${indicadores['euro']:,.2f}")
    st.caption(f"Estado API: {indicadores['estado']}")

# Initialize state
if "df_matriz" not in st.session_state:
    st.session_state["df_matriz"] = generar_matriz_ejemplo()

if uploaded_auto is not None:
    df_c = cargar_planilla_autogestion(uploaded_auto)
    if df_c is not None and not df_c.empty:
        st.session_state["df_matriz"] = df_c

df_matriz = st.session_state["df_matriz"]

# -----------------------------------------------------------------------------
# 4. FORMULARIO DE INGRESO MANUAL DE OFERTAS
# -----------------------------------------------------------------------------
st.title("🛒 Cuadro Comparativo Multimaterial - Autogestión")

with st.expander("➕ **Ingreso Manual de Ofertas / Posiciones**", expanded=True):
    with st.form("form_agregar_oferta", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2, 3, 1.5, 1.5])
        with c1:
            m_sp = st.text_input("ID SOLPED (SP)", value="PR176577", help="Ejemplo: PR176577")
        with c2:
            m_desc = st.text_input("Descripción breve del material", value="VÁLVULA BOLA 2 INCH ANSI 300")
        with c3:
            m_cant = st.number_input("Cantidad", min_value=1, value=10)
        with c4:
            m_um = st.selectbox("UM", ["C/U", "LTS", "MTR", "SET", "KG"])

        c5, c6, c7, c8 = st.columns([2, 1.5, 2, 2.5])
        with c5:
            m_precio = st.number_input("Precio Unitario Oferta", min_value=0.0, value=150000.0, step=1000.0)
        with c6:
            m_moneda = st.selectbox("Moneda", ["CLP", "USD", "EUR"])
        with c7:
            m_fecha_entrega = st.date_input("Fecha Prometida de Entrega", value=datetime.date.today() + datetime.timedelta(days=10))
        with c8:
            m_prov = st.text_input("Proveedor Oferente", value="PARKER CHILE")

        m_comentarios = st.text_input("Comentarios / Observaciones", value="Garantía de 12 meses. Incluye flete a planta.")

        btn_agregar = st.form_submit_button("💾 Agregar Oferta a la Matriz", type="primary", use_container_width=True)

    if btn_agregar:
        # Tasa de conversión
        factor_moneda = 1.0
        if m_moneda == "USD":
            factor_moneda = indicadores["dolar"]
        elif m_moneda == "EUR":
            factor_moneda = indicadores["euro"]

        precio_clp = int(m_precio * factor_moneda)
        monto_total_clp = int(m_cant * precio_clp)
        dias_entrega = max(0, (m_fecha_entrega - datetime.date.today()).days)

        nueva_fila = {
            "SP": m_sp,
            "Texto breve": m_desc,
            "Cantidad": m_cant,
            "UM": m_um,
            "Precio Unitario": m_precio,
            "Moneda": m_moneda,
            "Precio Normalizado [CLP]": precio_clp,
            "Proveedor": m_prov,
            "Fecha Entrega": m_fecha_entrega,
            "Días Entrega": dias_entrega,
            "Monto Total [CLP]": monto_total_clp,
            "Comentarios": m_comentarios
        }

        st.session_state["df_matriz"] = pd.concat([st.session_state["df_matriz"], pd.DataFrame([nueva_fila])], ignore_index=True)
        st.success(f"✅ Oferta de **{m_prov}** agregada exitosamente a la SOLPED **{m_sp}**.")
        st.rerun()

# -----------------------------------------------------------------------------
# 5. FILTROS Y CONTROL POR SOLPED
# -----------------------------------------------------------------------------
st.divider()

col_solped = [c for c in df_matriz.columns if any(k in str(c).lower() for k in ['sp', 'solped', 'pr'])][0] if df_matriz.columns.any() else None

lista_ids = ["Todas las solicitudes"]
if col_solped and col_solped in df_matriz.columns:
    lista_ids.extend([str(x) for x in df_matriz[col_solped].dropna().unique() if str(x).strip() != ""])

f1, f2, f3 = st.columns([2.5, 2.5, 2])
with f1:
    id_sel = st.selectbox("🔍 Filtrar por SP / SOLPED", lista_ids)
with f2:
    val_def = id_sel if id_sel != "Todas las solicitudes" else "REPORTE-CONSOLIDADO"
    id_reporte = st.text_input("ID para Informe Final", value=val_def)
with f3:
    comprador = st.text_input("Evaluador / Comprador", value="Felipe Martínez")

df_filtrado = df_matriz[df_matriz[col_solped].astype(str) == id_sel] if id_sel != "Todas las solicitudes" and col_solped else df_matriz

# -----------------------------------------------------------------------------
# 6. ANÁLISIS GRÁFICO (COMPARATIVA DE OFERTAS Y TIEMPOS DE ENTREGA)
# -----------------------------------------------------------------------------
if not df_filtrado.empty and "Proveedor" in df_filtrado.columns:
    st.subheader("📈 Análisis Comparativo Económico y de Plazos")
    g1, g2 = st.columns(2)

    col_precio_norm = "Precio Normalizado [CLP]" if "Precio Normalizado [CLP]" in df_filtrado.columns else "Monto Total [CLP]"
    col_dias = "Días Entrega" if "Días Entrega" in df_filtrado.columns else "Dias de tratamiento"

    with g1:
        if col_precio_norm in df_filtrado.columns:
            fig_precio = px.bar(
                df_filtrado,
                x="Proveedor",
                y=col_precio_norm,
                color="Proveedor",
                text_auto=',.0f',
                title="💰 Comparativa de Precios (Normalizados CLP)",
                labels={col_precio_norm: "Monto CLP", "Proveedor": "Proveedor"}
            )
            fig_precio.update_layout(showlegend=False, height=320)
            st.plotly_chart(fig_precio, use_container_width=True)

    with g2:
        if col_dias in df_filtrado.columns:
            fig_dias = px.bar(
                df_filtrado,
                x="Proveedor",
                y=col_dias,
                color="Proveedor",
                text_auto=True,
                title="⏱️ Días Prometidos de Entrega (Lead Time)",
                labels={col_dias: "Días", "Proveedor": "Proveedor"}
            )
            fig_dias.update_layout(showlegend=False, height=320)
            st.plotly_chart(fig_dias, use_container_width=True)

# -----------------------------------------------------------------------------
# 7. MATRIZ EDITABLE Y EXPORTACIÓN
# -----------------------------------------------------------------------------
st.subheader(f"📊 Matriz de Cotizaciones ({len(df_filtrado)} Registros)")

df_edited = st.data_editor(
    df_filtrado,
    num_rows="dynamic",
    use_container_width=True,
    height=380
)

st.divider()
st.subheader("📝 Dictamen del Reporte")
comentarios_reporte = st.text_area(
    "Justificación de Selección / Adjudicación:",
    value="Se selecciona la propuesta óptima considerando equilibrio entre mejor precio unitario en CLP normalizado y el tiempo de respuesta requerido por operación."
)

col_d1, col_d2 = st.columns(2)
with col_d1:
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df_edited.to_excel(writer, sheet_name=f'SOLPED_{id_reporte}'[:31], index=False)

    st.download_button(
        label="📥 Descargar Comparativo en Excel",
        data=out_excel.getvalue(),
        file_name=f"Cuadro_Comparativo_{id_reporte}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

with col_d2:
    if st.button("🧹 Limpiar y Reiniciar Matriz", use_container_width=True):
        st.session_state["df_matriz"] = generar_matriz_ejemplo()
        st.rerun()
