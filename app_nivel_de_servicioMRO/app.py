"""
Streamlit - Cuadro Comparativo y Planilla de Gestión
Gestión multimoneda (Material + Transporte) y análisis comparativo con tarjetas interactivas SOLPED.

Correr local: streamlit run app.py
"""
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="Cuadro Comparativo y Planilla de Gestión", layout="wide")

# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================
def convertir_a_moneda_base(monto, moneda_origen, moneda_destino="USD", usd_clp=950.0, eur_clp=1020.0):
    """Convierte un monto desde su moneda de origen a una moneda de destino unificada."""
    if pd.isna(monto) or monto == 0:
        return 0.0
    
    mon_orig = str(moneda_origen).upper().strip() if pd.notna(moneda_origen) else "CLP"
    mon_dest = str(moneda_destino).upper().strip()
    
    if mon_orig in ["CLP", "CLF", "PESO", "PESOS"]:
        monto_clp = float(monto)
    elif mon_orig in ["USD", "US$", "DOLARES", "DOLAR"]:
        monto_clp = float(monto) * usd_clp
    elif mon_orig in ["EUR", "EUR$", "EUROS"]:
        monto_clp = float(monto) * eur_clp
    else:
        monto_clp = float(monto)
        
    if mon_dest == "CLP":
        return monto_clp
    elif mon_dest == "USD":
        return monto_clp / usd_clp if usd_clp > 0 else monto_clp
    elif mon_dest == "EUR":
        return monto_clp / eur_clp if eur_clp > 0 else monto_clp
    
    return monto_clp

def encontrar_columna(df, posibles_nombres):
    """Busca coincidencias de nombres de columnas ignorando mayúsculas/minúsculas."""
    for col in df.columns:
        if str(col).strip().lower() in [n.lower() for n in posibles_nombres]:
            return col
    return None

def aplicar_conversion_multimoneda(df, moneda_base, tasa_usd_clp, tasa_eur_clp):
    """Detecta montos y monedas de Material y Transporte, calcula equivalentes en Moneda Base."""
    df = df.copy()
    
    nombres_precio_mat = ["Valor neto de pedido", "Valor unitario", "Precio Material", "Precio Unitario", "Precio Neto", "Valor Material"]
    col_precio_mat = encontrar_columna(df, nombres_precio_mat)
    
    nombres_moneda_mat = ["Moneda", "Moneda Material", "Moneda Mat", "CM"]
    col_moneda_mat = encontrar_columna(df, nombres_moneda_mat)
    
    nombres_precio_trans = [
        "Precio Transporte", "Precio de Transporte", "Valor Flete", "Flete", 
        "Transporte", "Precio Flete", "Envio", "Envío", "Costo Envio", 
        "Coste de Envio", "Transporte Costo", "Costo Transporte"
    ]
    col_precio_trans = encontrar_columna(df, nombres_precio_trans)
    
    nombres_moneda_trans = ["Moneda Transporte", "Moneda Flete", "Moneda Trans", "Moneda Envio"]
    col_moneda_trans = encontrar_columna(df, nombres_moneda_trans)
    
    if col_precio_mat:
        monedas_mat = df[col_moneda_mat] if col_moneda_mat else "CLP"
        df["Monto Material (Base)"] = [
            convertir_a_moneda_base(p, m, moneda_base, tasa_usd_clp, tasa_eur_clp)
            for p, m in zip(df[col_precio_mat].fillna(0), monedas_mat)
        ]
    else:
        df["Monto Material (Base)"] = 0.0

    if col_precio_trans:
        monedas_trans = df[col_moneda_trans] if col_moneda_trans else (df[col_moneda_mat] if col_moneda_mat else "CLP")
        df["Monto Transporte (Base)"] = [
            convertir_a_moneda_base(p, m, moneda_base, tasa_usd_clp, tasa_eur_clp)
            for p, m in zip(df[col_precio_trans].fillna(0), monedas_trans)
        ]
    else:
        df["Monto Transporte (Base)"] = 0.0

    df["Monto Total (Base)"] = df["Monto Material (Base)"] + df["Monto Transporte (Base)"]
    df["Moneda Base"] = moneda_base
    return df

# ==============================================================================
# CONFIGURACIÓN TEMA
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
        .st-key-theme_toggle { position: fixed !important; top: 65px !important; right: 15px !important; z-index: 999999 !important; width: 45px !important; height: 45px !important; min-width: 0 !important; }
        .st-key-theme_toggle button { background: #FFFFFF !important; border: 1px solid #E0E0E0 !important; border-radius: 50% !important; box-shadow: 0 2px 5px rgba(0,0,0,0.1) !important; font-size: 1.4rem !important; padding: 0 !important; margin: 0 !important; color: #111111 !important; display: flex !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; min-height: unset !important; }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #F0F0F0 !important; }
        </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
        <style>
        .st-key-theme_toggle { position: fixed !important; top: 65px !important; right: 15px !important; z-index: 999999 !important; width: 45px !important; height: 45px !important; min-width: 0 !important; }
        .st-key-theme_toggle button { background: #1E2329 !important; border: 1px solid #444444 !important; border-radius: 50% !important; box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important; font-size: 1.4rem !important; padding: 0 !important; margin: 0 !important; color: #FF3333 !important; display: flex !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; min-height: unset !important; }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #2C323A !important; border-color: #FF3333 !important; }
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], html, body, [data-testid="stHeader"] { background-color: #0E1117 !important; color: #FF3333 !important; }
        p, span, label, h1, h2, h3, h4, h5, h6, div, td, th, caption, .stMarkdown { color: #FF3333 !important; }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button) { background-color: #CC0000 !important; color: #FFFFFF !important; border: 1px solid #FF4D4D !important; font-weight: bold !important; }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button):hover { background-color: #FF0000 !important; color: #FFFFFF !important; border-color: #FF6666 !important; }
        input, select, textarea, div[data-baseweb="select"] { background-color: #1E2329 !important; color: #FF3333 !important; border-color: #CC0000 !important; }
        [data-testid="stForm"], div[data-testid="stVerticalBlock"] > div:has(input[type="password"]) { background-color: #1E2329 !important; padding: 2rem !important; border-radius: 12px !important; border: 1px solid #CC0000 !important; }
        </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# ACCESO CON CONTRASEÑA
# ==============================================================================
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("📊 Cuadro Comparativo y Planilla de Gestión")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password")
        if st.button("Ingresar"):
            if clave_ingresada == st.secrets["app_password"]:
                st.session_state["_autenticado"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

# ==============================================================================
# INTERFAZ PRINCIPAL
# ==============================================================================
st.title("📊 Cuadro Comparativo y Planilla de Gestión")

with st.sidebar:
    st.header("⚙️ Configuración de Datos")
    archivo_ofertas = st.file_uploader("Subir matriz de ofertas (.xlsx)", type=["xlsx"])
    
    st.divider()
    st.header("💱 Conversión Multimoneda")
    moneda_base = st.selectbox("Moneda Consolidada Base", ["USD", "CLP", "EUR"], index=0)
    tasa_usd_clp = st.number_input("Tasa USD / CLP", value=950.0, step=1.0)
    tasa_eur_clp = st.number_input("Tasa EUR / CLP", value=1020.0, step=1.0)

tab_comparativo, tab_planilla = st.tabs(["⚖️ Cuadro Comparativo", "📝 Evaluación SOLPED"])

# ==============================================================================
# PESTAÑA 1: CUADRO COMPARATIVO
# ==============================================================================
with tab_comparativo:
    st.subheader("Análisis de Ofertas y Proveedores")
    
    if archivo_ofertas:
        try:
            df_raw = pd.read_excel(archivo_ofertas)
            df_procesado = aplicar_conversion_multimoneda(df_raw, moneda_base, tasa_usd_clp, tasa_eur_clp)
            
            total_mat = df_procesado["Monto Material (Base)"].sum()
            total_trans = df_procesado["Monto Transporte (Base)"].sum()
            total_gen = df_procesado["Monto Total (Base)"].sum()
            
            m1, m2, m3 = st.columns(3)
            m1.metric(f"Total Materiales ({moneda_base})", f"{total_mat:,.2f}")
            m2.metric(f"Total Transporte/Envío ({moneda_base})", f"{total_trans:,.2f}")
            m3.metric(f"Monto Total Consolidado ({moneda_base})", f"{total_gen:,.2f}")
            
            st.divider()
            
            st.dataframe(
                df_procesado,
                use_container_width=True,
                height=600, 
                column_config={
                    "Monto Material (Base)": st.column_config.NumberColumn(f"Costo Material ({moneda_base})", format="%.2f"),
                    "Monto Transporte (Base)": st.column_config.NumberColumn(f"Precio Transporte ({moneda_base})", format="%.2f"),
                    "Monto Total (Base)": st.column_config.NumberColumn(f"Total Final ({moneda_base})", format="%.2f"),
                }
            )
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
    else:
        st.info("Sube una planilla Excel en la barra lateral para procesar los costos de materiales y transporte.")

# ==============================================================================
# PESTAÑA 2: EVALUACIÓN SOLPED (TARJETAS INTERACTIVAS)
# ==============================================================================
with tab_planilla:
    st.subheader("Seguimiento y Control de Adjudicaciones (SOLPED)")
    
    if archivo_ofertas and 'df_procesado' in locals():
        df_gestion = df_procesado.copy()
        
        # Búsqueda dinámica de nombres de columnas originales para inyectar en la UI
        col_desc = encontrar_columna(df_gestion, ["Descripción", "Material", "Texto breve", "Texto de material", "Detalle"])
        col_cant = encontrar_columna(df_gestion, ["Cantidad", "Cant", "CANTIDAD", "Cant."])
        col_prov = encontrar_columna(df_gestion, ["Proveedor", "Licitante", "Nombre Proveedor"])
        col_precio_mat = encontrar_columna(df_gestion, ["Valor neto de pedido", "Valor unitario", "Precio Material", "Precio Unitario", "Precio Neto"])
        col_moneda_mat = encontrar_columna(df_gestion, ["Moneda", "Moneda Material", "CM"])
        col_precio_trans = encontrar_columna(df_gestion, ["Precio Transporte", "Valor Flete", "Flete", "Transporte", "Envio", "Costo Envio"])

        st.markdown("Revisa, ajusta cantidades o edita los **Costos de Transporte** inyectados. El Total se actualizará en tiempo real.")

        for idx, row in df_gestion.iterrows():
            # Inyección de datos detectados (o valores por defecto)
            desc_inyectada = row[col_desc] if col_desc and pd.notna(row[col_desc]) else f"Posición {idx+1}"
            cant_inyectada = float(row[col_cant]) if col_cant and pd.notna(row[col_cant]) else 1.0
            pu_inyectado = float(row[col_precio_mat]) if col_precio_mat and pd.notna(row[col_precio_mat]) else 0.0
            moneda_inyectada = str(row[col_moneda_mat]) if col_moneda_mat and pd.notna(row[col_moneda_mat]) else "CLP"
            prov_inyectado = str(row[col_prov]) if col_prov and pd.notna(row[col_prov]) else "Proveedor Sin Definir"
            transporte_inyectado = float(row[col_precio_trans]) if col_precio_trans and pd.notna(row[col_precio_trans]) else 0.0
            
            with st.container():
                st.markdown(f"### Pos {(idx+1)*10}: {desc_inyectada}")
                st.caption(f"Centro: General | UM Original: UN")
                
                # Fila 1: Datos Base
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    cantidad = st.number_input("Cantidad", value=cant_inyectada, min_value=0.0, step=1.0, key=f"cant_{idx}")
                with col2:
                    precio_unitario = st.number_input("Precio Unitario", value=pu_inyectado, step=100.0, key=f"pu_{idx}")
                with col3:
                    moneda = st.text_input("Moneda", value=moneda_inyectada, key=f"moneda_{idx}", disabled=True)
                with col4:
                    proveedor = st.text_input("Proveedor", value=prov_inyectado, key=f"prov_{idx}")

                # Fila 2: Transporte y Total
                col5, col6, col7, col8 = st.columns(4)
                with col5:
                    incoterm = st.selectbox("Condición Transp.", ["EXW", "DDP", "FOB", "CIF", "Otro"], key=f"inco_{idx}")
                with col6:
                    costo_transporte = st.number_input(
                        "Costo de Transporte (Inyectado)", 
                        value=transporte_inyectado, 
                        step=1000.0, 
                        help="Valor detectado. Edita si el precio de flete/envío es incorrecto.",
                        key=f"costo_trans_{idx}"
                    )
                with col7:
                    fecha_entrega = st.date_input("Fecha Entrega Aprox.", key=f"fecha_{idx}")
                
                with col8:
                    # Cálculo en Tiempo Real
                    costo_material_total = cantidad * precio_unitario
                    total_final = costo_material_total + costo_transporte
                    st.info(f"**Monto Total ({moneda})**\n### {total_final:,.2f}")
                    
                st.divider()
    else:
        st.info("Carga la planilla de ofertas en la barra lateral para visualizar las posiciones SOLPED a evaluar.")
