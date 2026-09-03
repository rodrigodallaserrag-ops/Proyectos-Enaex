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
# 2. CARGA Y PROCESAMIENTO ROBUSTO DE PLANILLA DE AUTOGESTIÓN (.XLSM)
# -----------------------------------------------------------------------------
def cargar_planilla_autogestion(file):
    """
    Escanea las pestañas del archivo Excel/XLSM, detecta automáticamente la fila 
    donde comienza la tabla de datos omitiendo títulos superiores y convirtiendo
    cada celda a string para prevenir errores de tipo float.
    """
    try:
        nombre = file.name.lower()
        file_bytes = io.BytesIO(file.getvalue())

        if nombre.endswith(('.xlsx', '.xlsm', '.xls')):
            engine = 'openpyxl' if nombre.endswith(('.xlsx', '.xlsm')) else None
            excel_file = pd.ExcelFile(file_bytes, engine=engine)
            
            palabras_clave = [
                'pos', 'solped', 'código', 'codigo', 'descripción', 'descripcion', 
                'cantidad', 'um', 'último', 'ultimo', 'oferta', 'proveedor', 'monto'
            ]
            
            df_final = None

            for sheet_name in excel_file.sheet_names:
                try:
                    df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
                except Exception:
                    continue

                if df_raw.empty:
                    continue

                fila_encabezado_idx = None
                
                # Buscar fila de encabezados convirtiendo cada celda de forma ultra segura a texto
                for idx, row in df_raw.iterrows():
                    celdas_texto = [str(val).lower() for val in row.values if pd.notna(val)]
                    
                    coincidencias = 0
                    for kw in palabras_clave:
                        if any(kw in celda for celda in celdas_texto):
                            coincidencias += 1
                            
                    if coincidencias >= 2:
                        fila_encabezado_idx = idx
                        break

                if fila_encabezado_idx is not None:
                    df_clean = df_raw.iloc[fila_encabezado_idx:].copy()
                    
                    # Generar encabezados limpios
                    nuevos_encabezados = []
                    for i, col_val in enumerate(df_clean.iloc[0]):
                        val_str = str(col_val).strip() if pd.notna(col_val) else ""
                        if val_str != "" and val_str.lower() != "none" and val_str.lower() != "nan":
                            nuevos_encabezados.append(val_str)
                        else:
                            nuevos_encabezados.append(f"Columna_{i+1}")
                    
                    df_clean.columns = nuevos_encabezados
                    df_clean = df_clean.iloc[1:].reset_index(drop=True)
                    df_clean = df_clean.dropna(how='all')
                    
                    # Filtrar columnas genéricas residuales
                    cols_validas = [c for c in df_clean.columns if not str(c).startswith("Columna_")]
                    if cols_validas:
                        df_clean = df_clean[cols_validas]
                    
                    df_final = df_clean
                    break

            if df_final is None or df_final.empty:
                df_final = pd.read_excel(excel_file, sheet_name=0)
                df_final = df_final.dropna(how='all').dropna(how='all', axis=1)

            return df_final

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
    """Genera matriz inicial editable con +20 materiales idéntica a la plantilla Excel"""
    data = []
    materiales_demo = [
        ("LIMPIA CONTACTOS 279 CHESTERTON.", "C/U", 72, 54500, "CLP", "PRINTEC S A"),
        ("ACEITE HIDRAULICO ISO 68", "LTS", 500, 3200, "CLP", "LUBRICANTES CHILE"),
        ("VALVULA DE BOLA 2 INCH ANSI 300", "C/U", 15, 180000, "CLP", "MCM CHILE"),
        ("MANGUERA ALTA PRESION 1/2", "MTR", 120, 25000, "CLP", "PARKER"),
        ("EMPADRON BASTIDOR SOPORTE", "C/U", 4, 450000, "CLP", "INDURA"),
        ("RODAMIENTO BOLA 6204-2RSH", "C/U", 50, 18500, "CLP", "SKF CHILE"),
        ("KITS DE EMPADRADO COMPLETO", "SET", 2, 1200000, "CLP", "SKF CHILE"),
        ("ELECTRODO AWS E7018 1/8", "KG", 200, 4800, "CLP", "INDURA"),
        ("FILTRO DE AIRE PRIMARIO", "C/U", 24, 65000, "CLP", "CATERPILLAR"),
        ("CORREA TRANSMISION V B52", "C/U", 30, 12500, "CLP", "GATES CHILE"),
    ]
    
    for i in range(1, 21):
        idx = (i - 1) % len(materiales_demo)
        desc, um, cant, precio, mon, prov = materiales_demo[idx]
        precio_usd = round(precio / indicadores["dolar"], 2) if mon == "CLP" else precio
        
        data.append({
            "Pos": i * 10,
            "Solped": "1002610299",
            "Código SAP": f"2000{6120 + i}",
            "Descripción breve": f"{desc} #{i}",
            "Cantidad": cant,
            "UM": um,
            "Última compra": "19-06-2025",
            "Proveedor Histórico": prov,
            "Último Precio [Unit]": precio,
            "Moneda Hist.": mon,
            "Último precio [USD]": precio_usd,
            "Incoterm": "T. Gil",
            "Oferta Inicial Moneda": precio * 1.05,
            "Oferta Mejorada Moneda": precio * 0.95,
            "Lead Time": "5 DIAS",
            "Validación Técnica": True,
            "Proveedor Sugerido": prov,
            "Tipo Adjudicación": "Proveedor único",
            "Monto Adjudicado": int(cant * precio * 0.95),
            "Comentario": "Se adjudica por historial de compra y validación técnica favorable."
        })
    return pd.DataFrame(data)

# -----------------------------------------------------------------------------
# 3. INTERFAZ LATERAL (CONFIGURACIÓN Y PLANILLA DE AUTOGESTIÓN)
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
    st.write(f"**UF:** ${indicadores['uf']:,.2f}")
    st.caption(f"Estado API: {indicadores['estado']}")

# -----------------------------------------------------------------------------
# 4. CUADRO COMPARATIVO MULTIMATERIAL CENTRAL (+20 MATERIALES)
# -----------------------------------------------------------------------------
st.title("🛒 Cuadro Comparativo Multimaterial - Autogestión")
st.caption("Consolidado de cotizaciones y adjudicación automatizada en una sola pantalla.")

if uploaded_auto is not None:
    df_cargado = cargar_planilla_autogestion(uploaded_auto)
    if df_cargado is not None and not df_cargado.empty:
        st.session_state["df_matriz"] = df_cargado
        st.success(f"✅ Planilla '{uploaded_auto.name}' procesada y cargada correctamente.")

if "df_matriz" not in st.session_state:
    st.session_state["df_matriz"] = generar_matriz_ejemplo()

df_matriz = st.session_state["df_matriz"]

# Controles de acción directa
col_b1, col_b2, col_b3 = st.columns([2, 2, 4])
with col_b1:
    if st.button("🧹 Limpiar Cuadro Comparativo", type="secondary"):
        st.session_state["df_matriz"] = pd.DataFrame(columns=df_matriz.columns)
        st.rerun()
with col_b2:
    if st.button("🔄 Recargar Datos Demostración"):
        st.session_state["df_matriz"] = generar_matriz_ejemplo()
        st.rerun()

st.subheader(f"📊 Matriz de Cotizaciones ({len(df_matriz)} Materiales / Posiciones)")

# Editor de datos interactivo
df_edited = st.data_editor(
    df_matriz,
    num_rows="dynamic",
    use_container_width=True,
    height=550
)

st.session_state["df_matriz"] = df_edited

# -----------------------------------------------------------------------------
# 5. RESUMEN DE ADJUDICACIÓN Y EXPORTACIÓN MASIVA
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📈 Resumen Estadístico de Adjudicación")

if not df_edited.empty:
    col_monto = [col for col in df_edited.columns if 'monto' in str(col).lower() or 'adjudicado' in str(col).lower()]
    if col_monto:
        monto_series = pd.to_numeric(df_edited[col_monto[0]], errors='coerce').fillna(0)
        total_adjudicado = float(monto_series.sum())
    else:
        total_adjudicado = 0.0
        
    total_items = len(df_edited)
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Materiales Comparados", f"{total_items} Pos")
    col_m2.metric("Monto Total Adjudicado (CLP)", f"$ {total_adjudicado:,.0f}".replace(",", "."))
    col_m3.metric("Promedio por Ítem", f"$ {total_adjudicado/max(1, total_items):,.0f}".replace(",", "."))

    # Exportación a Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_edited.to_excel(writer, sheet_name='Cuadro Comparativo', index=False)
    
    st.download_button(
        label="📥 Descargar Cuadro Comparativo Automatizado (Excel)",
        data=output_excel.getvalue(),
        file_name=f"Cuadro_Comparativo_Autogestion_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
