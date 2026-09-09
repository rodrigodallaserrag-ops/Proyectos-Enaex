"""
Streamlit - Cuadro Comparativo y Planilla de Gestión
Gestión multimoneda y análisis comparativo de ofertas y proveedores.

Correr local: streamlit run app.py
"""
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="Cuadro Comparativo y Planilla de Gestión", layout="wide")

# ==============================================================================
# FUNCIONES AUXILIARES DE CONVERSIÓN DE MONEDA
# ==============================================================================
def convertir_a_moneda_base(monto, moneda_origen, moneda_destino="USD", usd_clp=950.0, eur_clp=1020.0):
    """
    Convierte un monto desde su moneda de origen a una moneda de destino unificada[cite: 9].
    """
    if pd.isna(monto) or monto == 0:
        return 0.0
    
    mon_orig = str(moneda_origen).upper().strip() if pd.notna(moneda_origen) else "CLP"
    mon_dest = str(moneda_destino).upper().strip()
    
    # 1. Convertir primero a CLP como moneda intermedia de referencia[cite: 9]
    if mon_orig in ["CLP", "CLF", "PESO", "PESOS"]:
        monto_clp = float(monto)
    elif mon_orig in ["USD", "US$", "DOLARES", "DOLAR"]:
        monto_clp = float(monto) * usd_clp
    elif mon_orig in ["EUR", "EUR$", "EUROS"]:
        monto_clp = float(monto) * eur_clp
    else:
        monto_clp = float(monto)  # Por defecto si no se reconoce[cite: 9]
        
    # 2. Convertir de CLP a la Moneda Destino elegida[cite: 9]
    if mon_dest == "CLP":
        return monto_clp
    elif mon_dest == "USD":
        return monto_clp / usd_clp if usd_clp > 0 else monto_clp
    elif mon_dest == "EUR":
        return monto_clp / eur_clp if eur_clp > 0 else monto_clp
    
    return monto_clp

def aplicar_conversion_multimoneda(df, moneda_base, tasa_usd_clp, tasa_eur_clp):
    """
    Calcula los valores monetarios normalizados[cite: 9].
    (Ajusta los nombres de las columnas según tu nuevo archivo de proveedores).
    """
    df = df.copy()
    
    # Identificación/fallback de columnas de precio y moneda[cite: 9]
    col_precio_mat = "Precio Unitario" if "Precio Unitario" in df.columns else None
    col_moneda_mat = "Moneda" if "Moneda" in df.columns else None
    
    # Conversión
    if col_precio_mat and col_precio_mat in df.columns:
        monedas_mat = df[col_moneda_mat] if (col_moneda_mat and col_moneda_mat in df.columns) else "USD"
        df["Precio Normalizado (Base)"] = [
            convertir_a_moneda_base(p, m, moneda_base, tasa_usd_clp, tasa_eur_clp)
            for p, m in zip(df[col_precio_mat].fillna(0), monedas_mat)
        ]
    else:
        df["Precio Normalizado (Base)"] = 0.0

    df["Moneda Base"] = moneda_base
    return df

# ==============================================================================
# CONFIGURACIÓN TEMA (CLARO PREDETERMINADO / OSCURO)
# ==============================================================================
if "tema" not in st.session_state:
    st.session_state["tema"] = "claro"

icono_tema = "🌙" if st.session_state["tema"] == "claro" else "☀️"
if st.button(icono_tema, key="theme_toggle", help="Alternar Modo Claro/Oscuro"):
    st.session_state["tema"] = "oscuro" if st.session_state["tema"] == "claro" else "claro"
    st.rerun()

if st.session_state["tema"] == "claro":
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important;
            right: 15px !important;
            z-index: 999999 !important;
            width: 45px !important;
            height: 45px !important;
            min-width: 0 !important; 
        }
        .st-key-theme_toggle button {
            background: #FFFFFF !important;
            border: 1px solid #E0E0E0 !important;
            border-radius: 50% !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1) !important;
            font-size: 1.4rem !important;
            padding: 0 !important;
            margin: 0 !important;
            color: #111111 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            line-height: 1 !important;
            width: 100% !important;
            height: 100% !important;
            min-height: unset !important;
        }
        .st-key-theme_toggle button p {
            margin: 0 !important;
            padding: 0 !important;
            line-height: 1 !important;
            font-size: 1.4rem !important;
        }
        .st-key-theme_toggle button:hover {
            transform: scale(1.1) !important;
            background: #F0F0F0 !important;
        }
        </style>
    """, unsafe_allow_html=True) #[cite: 9]
else:
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important;
            top: 65px !important;
            right: 15px !important;
            z-index: 999999 !important;
            width: 45px !important;
            height: 45px !important;
            min-width: 0 !important;
        }
        .st-key-theme_toggle button {
            background: #1E2329 !important;
            border: 1px solid #444444 !important;
            border-radius: 50% !important;
            box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important;
            font-size: 1.4rem !important;
            padding: 0 !important;
            margin: 0 !important;
            color: #FF3333 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            line-height: 1 !important;
            width: 100% !important;
            height: 100% !important;
            min-height: unset !important;
        }
        .st-key-theme_toggle button p {
            margin: 0 !important;
            padding: 0 !important;
            line-height: 1 !important;
            font-size: 1.4rem !important;
        }
        .st-key-theme_toggle button:hover {
            transform: scale(1.1) !important;
            background: #2C323A !important;
            border-color: #FF3333 !important;
        }

        .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], html, body, [data-testid="stHeader"] {
            background-color: #0E1117 !important;
            color: #FF3333 !important;
        }

        p, span, label, h1, h2, h3, h4, h5, h6, div, td, th, caption, .stMarkdown {
            color: #FF3333 !important;
        }

        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button) {
            background-color: #CC0000 !important;
            color: #FFFFFF !important;
            border: 1px solid #FF4D4D !important;
            font-weight: bold !important;
        }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button):hover {
            background-color: #FF0000 !important;
            color: #FFFFFF !important;
            border-color: #FF6666 !important;
        }

        input, select, textarea, div[data-baseweb="select"] {
            background-color: #1E2329 !important;
            color: #FF3333 !important;
            border-color: #CC0000 !important;
        }

        [data-testid="stForm"], div[data-testid="stVerticalBlock"] > div:has(input[type="password"]) {
            background-color: #1E2329 !important;
            padding: 2rem !important;
            border-radius: 12px !important;
            border: 1px solid #CC0000 !important;
        }
        </style>
    """, unsafe_allow_html=True) #[cite: 9]


# ==============================================================================
# ACCESO CON CONTRASEÑA
# ==============================================================================
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("📊 Cuadro Comparativo y Planilla de Gestión")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password") #[cite: 9]
        if st.button("Ingresar"): #[cite: 9]
            if clave_ingresada == st.secrets["app_password"]: #[cite: 9]
                st.session_state["_autenticado"] = True #[cite: 9]
                st.rerun() #[cite: 9]
            else:
                st.error("Contraseña incorrecta.") #[cite: 9]
        st.stop() #[cite: 9]


# ==============================================================================
# INTERFAZ PRINCIPAL
# ==============================================================================
st.title("📊 Cuadro Comparativo y Planilla de Gestión")

with st.sidebar:
    st.header("Datos de entrada")
    archivo_ofertas = st.file_uploader("Subir matriz de ofertas (.xlsx)", type=["xlsx"])
    
    st.divider() #[cite: 9]
    st.header("💱 Conversión Multimoneda") #[cite: 9]
    moneda_base = st.selectbox("Moneda Consolidada Base", ["USD", "CLP", "EUR"], index=0) #[cite: 9]
    tasa_usd_clp = st.number_input("Tasa USD / CLP", value=950.0, step=1.0) #[cite: 9]
    tasa_eur_clp = st.number_input("Tasa EUR / CLP", value=1020.0, step=1.0) #[cite: 9]

# ---- Pestañas Principales ----
tab_comparativo, tab_planilla = st.tabs(["⚖️ Cuadro Comparativo", "📝 Planilla de Gestión"])

# ==============================================================================
# PESTAÑA 1: CUADRO COMPARATIVO
# ==============================================================================
with tab_comparativo:
    st.subheader("Análisis de Ofertas y Proveedores")
    
    if archivo_ofertas:
        try:
            df_ofertas = pd.read_excel(archivo_ofertas)
            # Aplicar conversión multimoneda (asegúrate de que las columnas coincidan con tu excel)
            df_ofertas_convertido = aplicar_conversion_multimoneda(df_ofertas, moneda_base, tasa_usd_clp, tasa_eur_clp)
            st.dataframe(df_ofertas_convertido, use_container_width=True)
            
            # Aquí puedes añadir tus tarjetas KPI o gráficos de comparación
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
    else:
        st.info("Sube un archivo en la barra lateral para generar el cuadro comparativo.")

# ==============================================================================
# PESTAÑA 2: PLANILLA DE GESTIÓN
# ==============================================================================
with tab_planilla:
    st.subheader("Seguimiento y Control de Adjudicaciones")
    st.write("Gestiona el estado de las evaluaciones y añade comentarios.")
    
    if archivo_ofertas and 'df_ofertas_convertido' in locals():
        # Crear columnas de gestión si no existen
        df_gestion = df_ofertas_convertido.copy()
        if "Estado" not in df_gestion.columns:
            df_gestion["Estado"] = "En Evaluación"
        if "Comentario" not in df_gestion.columns:
            df_gestion["Comentario"] = ""
            
        detalle_editado = st.data_editor(
            df_gestion,
            use_container_width=True,
            num_rows="dynamic",
            key="editor_gestion",
            column_config={
                "Estado": st.column_config.SelectboxColumn(
                    "Estado de Adjudicación",
                    options=["En Evaluación", "Adjudicado", "Rechazado", "Stand-by"],
                    required=True
                ),
                "Comentario": st.column_config.TextColumn("Comentarios del Comprador")
            }
        )
    else:
        st.info("Sube los datos iniciales en la barra lateral para habilitar la planilla de gestión editable.")
