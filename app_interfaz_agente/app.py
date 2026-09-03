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
# 1. TASAS DE CAMBIO EN TIEMPO REAL
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
# 2. BASE DE DATOS Y MULTI-POSICIÓN (EJEMPLO SOLPED 287)
# -----------------------------------------------------------------------------
SOLPEDS_BASE = {
    "287": [
        {"Pos": 1, "Material": "EXPLOSIVOS HIGH POWER", "Cantidad": 20, "UM": "C/U", "Precio Unitario": 8.0, "Moneda": "CLP", "Proveedor": "EXPLOSIVOS CHILE", "Días Entrega": 5, "Comentario": "⚠️ Precio de 8 pesos (corregir valor)"},
        {"Pos": 2, "Material": "MATERIAS PRIMAS QUIMICAS", "Cantidad": 70, "UM": "KG", "Precio Unitario": 12.5, "Moneda": "USD", "Proveedor": "CHEMICAL CORP", "Días Entrega": 14, "Comentario": "Importación directa"},
        {"Pos": 3, "Material": "TORNILLOS 200 KILOS", "Cantidad": 2, "UM": "C/U", "Precio Unitario": 45000.0, "Moneda": "CLP", "Proveedor": "PERNOS S.A.", "Días Entrega": 3, "Comentario": "Stock local disponible"},
    ],
    "PR176577": [
        {"Pos": 1, "Material": "DISCO RUPTURA GRAFITO 2`", "Cantidad": 10, "UM": "C/U", "Precio Unitario": 54500.0, "Moneda": "CLP", "Proveedor": "PRINTEC S A", "Días Entrega": 3, "Comentario": "Entrega inmediata"},
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
    pdf.cell(0, 6, f"Total Ítems evaluados: {len(df_data)}", ln=True)
    pdf.cell(0, 6, f"Monto Total Evaluado (CLP): ${total_monto:,.0f}".replace(",", "."), ln=True)
    pdf.ln(4)

    pdf.set_font("Arial", 'B', 9)
    pdf.cell(15, 7, "Pos", 1)
    pdf.cell(65, 7, "Material / Descripcion", 1)
    pdf.cell(15, 7, "Cant", 1)
    pdf.cell(45, 7, "Proveedor", 1)
    pdf.cell(20, 7, "Dias", 1)
    pdf.cell(30, 7, "Total CLP", 1, ln=True)

    pdf.set_font("Arial", '', 8)
    for _, row in df_data.iterrows():
        pdf.cell(15, 6, str(row.get("Pos", "")), 1)
        pdf.cell(65, 6, str(row.get("Material", ""))[:28], 1)
        pdf.cell(15, 6, str(row.get("Cantidad", "")), 1)
        pdf.cell(45, 6, str(row.get("Proveedor", ""))[:22], 1)
        pdf.cell(20, 6, str(row.get("Días Entrega", "")), 1)
        monto_clp = float(row.get("Monto Total CLP", 0))
        pdf.cell(30, 6, f"${monto_clp:,.0f}".replace(",", "."), 1, ln=True)

    return pdf.output(dest='S').encode('latin1')

# -----------------------------------------------------------------------------
# 3. INTERFAZ PRINCIPAL Y CARGA DE SOLPED MULTI-POSICIÓN
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("💱 Tipo de Cambio")
    st.write(f"**USD:** ${tasas['USD']:,.2f}")
    st.write(f"**EUR:** ${tasas['EUR']:,.2f}")

st.title("🛒 Consola de Compras - Gestión por SOLPED")

st.subheader("🔍 Cargar SOLPED con Múltiples Materiales")
col_search, col_btn = st.columns([3, 1])

with col_search:
    solped_input = st.text_input("Ingrese ID SOLPED (Ejemplo: 287 o PR176577):", value="287")

with col_btn:
    st.write(" ")
    st.write(" ")
    cargar_btn = st.button("📥 Cargar Materiales", use_container_width=True, type="primary")

# Cargar datos en la sesión al presionar el botón o inicio
if "datos_solped_actual" not in st.session_state or cargar_btn:
    sp_id = solped_input.strip()
    if sp_id in SOLPEDS_BASE:
        st.session_state["datos_solped_actual"] = pd.DataFrame(SOLPEDS_BASE[sp_id])
    else:
        st.session_state["datos_solped_actual"] = pd.DataFrame([
            {"Pos": 1, "Material": f"MATERIAL 1 DE SOLPED #{sp_id}", "Cantidad": 10, "UM": "C/U", "Precio Unitario": 1000.0, "Moneda": "CLP", "Proveedor": "PROVEEDOR A", "Días Entrega": 5, "Comentario": "Nuevo registro"},
            {"Pos": 2, "Material": f"MATERIAL 2 DE SOLPED #{sp_id}", "Cantidad": 5, "UM": "KG", "Precio Unitario": 50.0, "Moneda": "USD", "Proveedor": "PROVEEDOR B", "Días Entrega": 10, "Comentario": "Nuevo registro"}
        ])
    st.session_state["solped_id_cargada"] = sp_id

df_trabajo = st.session_state["datos_solped_actual"]

# -----------------------------------------------------------------------------
# 4. TABLA INTERACTIVA (EDICIÓN EN VIVO Y BORRADO DE POSICIONES)
# -----------------------------------------------------------------------------
st.subheader(f"✏️ Posiciones Cargadas para SOLPED #{st.session_state.get('solped_id_cargada', '287')}")
st.caption("💡 Modifique directamente los precios (ej. de 8 a valor real), cambie monedas, ajuste días de entrega o elimine/agregue filas.")

df_editado = st.data_editor(
    df_trabajo,
    num_rows="dynamic",
    use_container_width=True,
    height=240,
    column_config={
        "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "EUR"], default="CLP"),
        "UM": st.column_config.SelectboxColumn("UM", options=["C/U", "KG", "LTS", "SET", "MTR"], default="C/U"),
        "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
        "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=1),
        "Días Entrega": st.column_config.NumberColumn("Días Entrega", min_value=1)
    },
    key="editor_multi_posicion"
)

# Recalcular valores normalizados en CLP en tiempo real
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

# -----------------------------------------------------------------------------
# 5. GRÁFICOS AUTOMÁTICOS Y ANÁLISIS DE MEJOR OFERTA
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📈 Análisis de Materiales, Costos y Lead Time")

col_g1, col_g2 = st.columns(2)

with col_g1:
    fig_costo = px.bar(
        df_editado,
        x="Material",
        y="Monto Total CLP",
        color="Proveedor",
        text_auto=',.0f',
        title="💰 Monto Total Normalizado por Material [CLP]",
        labels={"Monto Total CLP": "Monto Total (CLP)"}
    )
    fig_costo.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_costo, use_container_width=True)

with col_g2:
    fig_dias = px.bar(
        df_editado,
        x="Material",
        y="Días Entrega",
        color="Proveedor",
        text_auto=True,
        title="⏱️ Días Prometidos de Entrega (Lead Time)",
        labels={"Días Entrega": "Días"}
    )
    fig_dias.update_layout(height=320, xaxis_tickangle=-15)
    st.plotly_chart(fig_dias, use_container_width=True)

# Resumen técnico
if not df_editado.empty:
    mejor_tiempo = df_editado.loc[df_editado['Días Entrega'].idxmin()]
    total_evaluado = df_editado["Monto Total CLP"].sum()
    
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Monto Total Acumulado SOLPED", f"$ {total_evaluado:,.0f}".replace(",", "."))
    col_m2.metric("Entrega Más Rápida", f"{mejor_tiempo['Material']} ({mejor_tiempo['Días Entrega']} días)")

# -----------------------------------------------------------------------------
# 6. EXPORTACIÓN DE REPORTES (EXCEL / PDF)
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📥 Exportar Informe de Evaluación")

c_rep1, c_rep2 = st.columns(2)
with c_rep1:
    comprador_nombre = st.text_input("Nombre del Comprador / Evaluador:", value="Felipe Martínez")
with c_rep2:
    id_sp_reporte = st.session_state.get('solped_id_cargada', '287')
    st.text_input("ID SOLPED para Documento:", value=f"SOLPED-{id_sp_reporte}", disabled=True)

col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    out_excel = io.BytesIO()
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        df_editado.to_excel(writer, sheet_name=f'SOLPED_{id_sp_reporte}', index=False)

    st.download_button(
        label="📥 Descargar Reporte Excel",
        data=out_excel.getvalue(),
        file_name=f"Evaluacion_SOLPED_{id_sp_reporte}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
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
                type="secondary",
                use_container_width=True
            )
        except Exception as err:
            st.warning(f"Error al generar PDF: {err}")
    else:
        st.info("Instale `fpdf` para habilitar la descarga en PDF.")
