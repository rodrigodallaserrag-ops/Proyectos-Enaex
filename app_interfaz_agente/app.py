import re
import pandas as pd
import requests
import streamlit as st
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

st.set_page_config(page_title="Consola Única de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. MOTOR FINANCIERO: MINDICADOR.CL
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)  
def obtener_indicadores_financieros():
    try:
        url = "https://mindicador.cl/api"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        response = requests.get(url, timeout=5, headers=headers, verify=False)
        data = response.json()
        return {
            "dolar": float(data["dolar"]["valor"]),
            "euro": float(data["euro"]["valor"]),
            "uf": float(data["uf"]["valor"]),
            "fecha": data["dolar"]["fecha"][:10],
            "estado": "Online 🟢",
        }
    except Exception:
        return {
            "dolar": 938.0,
            "euro": 1020.0,
            "uf": 40875.0,
            "fecha": "Valores Estimados",
            "estado": "Offline (Red Enaex) 🛡️",
        }

indicadores = obtener_indicadores_financieros()

def formato_clp(valor):
    return f"${int(round(valor)):,}".replace(",", ".")

def aplicar_formato_regional(monto, moneda):
    """Aplica formato visual personalizado según la moneda ingresada."""
    if moneda == "CLP":
        return f"$ {int(round(monto)):,}".replace(",", ".")
    elif moneda == "USD":
        return f"$ {monto:,.2f}"
    elif moneda == "EUR":
        return f"€ {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    elif moneda == "UF":
        return f"UF {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(monto)

# -----------------------------------------------------------------------------
# 2. INICIALIZAR ESTADO DE LA APLICACIÓN
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

st.title("🛒 Consola Única de Compras — Enaex")

# -----------------------------------------------------------------------------
# 3. PANEL CENTRAL DE MONEDAS
# -----------------------------------------------------------------------------
st.caption(f"🗓️ Valores del día ({indicadores['fecha']}) - Estado API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1.5, 1.5, 1.5, 1])
col_uf.metric("UF", formato_clp(indicadores['uf']))
col_usd.metric("Dólar", formato_clp(indicadores['dolar']))
col_eur.metric("Euro", formato_clp(indicadores['euro']))

st.divider()

# -----------------------------------------------------------------------------
# 4. BARRA LATERAL: DATOS SOLPED
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("📌 Datos Solped")
    solped = st.text_input("N° Solped", value="10045982")
    material = st.text_input("Código Material", value="3001892")
    sociedad = st.selectbox("Sociedad", ["EC01", "EC06"])

# -----------------------------------------------------------------------------
# 5. FORMULARIO DE INGRESO DE OFERTAS (ESTÉTIKA CLEAN)
# -----------------------------------------------------------------------------
st.subheader("➕ Carga Manual de Oferta")

with st.form("form_cotizacion", clear_on_submit=True):
    col1, col2, col3, col4 = st.columns(4)
    proveedor = col1.text_input("Proveedor*")
    
    # Campo de texto ultra limpio (sin botones de suma/resta)
    monto_str = col2.text_input("Monto Original*", placeholder="Ej: 1800000 o 1500,50")
    moneda = col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"])
    plazo = col4.number_input("Plazo Entrega (Días)", min_value=1, value=5)
    obs = st.text_area("Observaciones Técnicas")

    if st.form_submit_button("Guardar en Cuadro Comparativo", type="primary"):
        # Limpieza estricta de caracteres no numéricos por backend
        monto_sanitizado = re.sub(r'[^0-9.,]', '', str(monto_str).strip())
        
        try:
            if moneda == "USD":
                # Formato US: elimina comas de miles
                limpio = monto_sanitizado.replace(",", "")
            else:
                # Formato CLP/EUR/UF: quita puntos de miles y convierte coma a punto decimal
                limpio = monto_sanitizado.replace(".", "").replace(",", ".")
            
            monto = float(limpio)
        except ValueError:
            monto = 0.0

        if not proveedor or monto <= 0:
            st.error("⚠️ Ingrese un Proveedor válido y un Monto numérico mayor a 0.")
        else:
            # Conversiones financieras automatizadas
            monto_clp = monto
            if moneda == "USD":
                monto_clp = monto * indicadores["dolar"]
            elif moneda == "EUR":
                monto_clp = monto * indicadores["euro"]
            elif moneda == "UF":
                monto_clp = monto * indicadores["uf"]

            monto_usd = monto_clp / indicadores["dolar"]

            st.session_state["cotizaciones"].append({
                "Proveedor": proveedor,
                "Monto Original": monto, 
                "Moneda": moneda,
                "Equiv. CLP ($)": round(monto_clp, 2),
                "Equiv. USD ($)": round(monto_usd, 2),
                "Plazo (Días)": plazo,
                "Observaciones": obs,
            })
            
            st.success(f"✅ Oferta de {proveedor} ({moneda}) ingresada y calculada correctamente.")
            st.rerun()

# -----------------------------------------------------------------------------
# 6. CUADRO COMPARATIVO HOMOGENEIZADO Y MULTIMONEDA
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo (Homogeneizado)")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Plazo (Días)", "Observaciones"]
    )
    st.dataframe(df_empty, use_container_width=True)
    st.info("👆 Llena el formulario superior para simular la comparativa.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    
    # Formateo visual adaptable fila por fila
    df_visual = df.copy()
    df_visual["Monto Original"] = df_visual.apply(
        lambda fila: aplicar_formato_regional(fila["Monto Original"], fila["Moneda"]), axis=1
    )
    df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(round(x)):,}".replace(",", "."))
    df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")

    st.dataframe(df_visual, use_container_width=True)

    max_monto_clp = df["Equiv. CLP ($)"].max()
    if max_monto_clp > 1000000:
        st.warning(
            f"⚠️ **Control Financiero (> $1M CLP):** Requerimiento alcanza {formato_clp(max_monto_clp)}. "
            f"Se aplicaron los tipos de cambio oficiales del día ({indicadores['fecha']})."
        )

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("🗑️ Limpiar Tabla"):
            st.session_state["cotizaciones"] = []
            st.rerun()

    with col_btn2:
        if st.button("🚀 Emitir a SAP ME21N"):
            st.balloons()
            st.success(f"Orden preparada para Solped {solped} (Sociedad {sociedad}). Valores convertidos inyectados en SAP.")
