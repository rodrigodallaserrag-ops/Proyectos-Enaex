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
# 2. BASE DE DATOS Y CARGA INSTANTÁNEA POR SOLPED
# -----------------------------------------------------------------------------
SOLPEDS_DEMO = {
    "289283": {"desc": "DISCO RUPTURA GRAFITO 2`", "cant": 10, "um": "C/U", "precio": 58.00, "moneda": "USD", "prov": "MCM CHILE", "dias": 7, "comentario": "Proveedor con mejor lead time."},
    "1609": {"desc": "VALVULA DE BOLA 2 INCH ANSI 300", "cant": 5, "um": "C/U", "precio": 180000.0, "moneda": "CLP", "prov": "INDURA", "dias": 15, "comentario": "Precio estándar nacional."},
    "83723": {"desc": "ACEITE HIDRAULICO ISO 68", "cant": 200, "um": "LTS", "precio": 3.50, "moneda": "EUR", "prov": "TOTAL ENERGIES", "dias": 10, "comentario": "Importación directa."},
}

def obtener_datos_solped(id_sp):
    id_clean = str(id_sp).strip().upper()
    if id_clean in SOLPEDS_DEMO:
        return SOLPEDS_DEMO[id_clean]
    return {
        "desc": f"MATERIAL ASOCIADO A SOLPED #{id_clean}",
        "cant": 1,
        "um": "C/U",
        "precio": 100000.0,
        "moneda": "CLP",
        "prov": "PROVEEDOR POR DEFINIR",
        "dias": 10,
        "comentario": "Cargado automáticamente. Modifique los campos necesarios."
    }

# Inicializar sesión acumulativa
if "matriz_acumulada" not in st.session_state:
    st.session_state["matriz_acumulada"] = pd.DataFrame(columns=[
        "SP / SOLPED", "Descripción", "Cantidad", "UM", "Precio Unitario",
        "Moneda", "Precio Unit. CLP Norm.", "Proveedor",
        "Fecha Entrega", "Días Entrega", "Monto Total CLP", "Comentarios"
    ])

# Callback para cargar datos al presionar Enter o cambiar ID
def cargar_solped_callback():
    sp_id = st.session_state.get("input_solped_id", "").strip()
    if sp_id:
        datos = obtener_datos_solped(sp_id)
        st.session_state["f_desc"] = datos["desc"]
        st.session_state["f_cant"] = int(datos["cant"])
        st.session_state["f_um"] = datos["um"]
        st.session_state["f_precio"] = float(datos["precio"])
        st.session_state["f_moneda"] = datos["moneda"]
        st.session_state["f_prov"] = datos["prov"]
        st.session_state["f_dias"] = int(datos["dias"])
        st.session_state["f_comentario"] = datos["comentario"]

# Configuración inicial de campos
if "f_desc" not in st.session_state:
    st.session_state["input_solped_id"] = "289283"
    cargar_solped_callback()

# -----------------------------------------------------------------------------
# 3. INTERFAZ Y PANEL DE BÚSQUEDA RAPIDA
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("💱 Indicadores del Día")
    st.write(f"**USD:** ${indicadores['dolar']:,.2f}")
    st.write(f"**EUR:** ${indicadores['euro']:,.2f}")
    st.caption(f"Estado API: {indicadores['estado']}")

st.title("🛒 Consola de Autogestión y Cuadro Comparativo Multi-SOLPED")

**🔍 Carga Rápida de SOLPED (Escriba la ID y presione Ctrl + Enter)**

col_b1, col_b2 = st.columns([3, 1])
with col_b1:
    st.text_input(
        "Ingrese ID SOLPED (Ejemplos sugeridos: 289283, 1609, 83723 o cualquiera nueva):",
        key="input_solped_id",
        on_change=cargar_solped_callback
    )
with col_b2:
    st.write(" ")
    st.write(" ")
    if st.button("🔄 Cargar Informacion", use_container_width=True):
        cargar_solped_callback()

# -----------------------------------------------------------------------------
# 4. FORMULARIO DE REVISIÓN Y EDICIÓN
# -----------------------------------------------------------------------------
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

    if st.button("➕ Agregar esta SOLPED a la Matriz Comparativa", type="primary", use_container_width=True):
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
        st.success(f"✅ SOLPED #{st.session_state['input_solped_id']} agregada a la matriz.")

# -----------------------------------------------------------------------------
# 5. MATRIZ ACUMULADA Y ANÁLISIS COMPARATIVO
# -----------------------------------------------------------------------------
df_matriz = st.session_state["matriz_acumulada"]

st.divider()
st.subheader(f"📊 Matriz Comparativa Multi-SOLPED ({len(df_matriz)} Ítems Agregados)")

if df_matriz.empty:
    st.info("💡 La matriz está vacía. Ingrese una SOLPED arriba y presione **'Agregar esta SOLPED a la Matriz Comparativa'**.")
else:
    # Editor interactivo
    df_edited = st.data_editor(
        df_matriz,
        num_rows="dynamic",
        use_container_width=True,
        height=280
    )
    st.session_state["matriz_acumulada"] = df_edited

    # Gráficos comparativos
    st.subheader("📈 Análisis Comparativo Unitario")
    g1, g2 = st.columns(2)

    with g1:
        fig_precio = px.bar(
            df_edited,
            x="SP / SOLPED",
            y="Precio Unit. CLP Norm.",
            color="Proveedor",
            text_auto=',.0f',
            title="💰 Comparativa de Precio Unitario Normalizado [CLP]",
            labels={"Precio Unit. CLP Norm.": "Precio Unit. (CLP)"}
        )
        fig_precio.update_layout(height=300)
        st.plotly_chart(fig_precio, use_container_width=True)

    with g2:
        fig_dias = px.bar(
            df_edited,
            x="SP / SOLPED",
            y="Días Entrega",
            color="Proveedor",
            text_auto=True,
            title="⏱️ Días Prometidos de Entrega (Lead Time)",
            labels={"Días Entrega": "Días"}
        )
        fig_dias.update_layout(height=300)
        st.plotly_chart(fig_dias, use_container_width=True)

    # Exportación
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
            df_edited.to_excel(writer, sheet_name='Comparativo_SOLPEDs', index=False)

        st.download_button(
            label="📥 Descargar Reporte en Excel",
            data=out_excel.getvalue(),
            file_name=f"Reporte_Comparativo_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    with col_d2:
        if st.button("🧹 Vaciar Matriz Comparativa", use_container_width=True):
            st.session_state["matriz_acumulada"] = pd.DataFrame(columns=df_matriz.columns)
            st.rerun()
