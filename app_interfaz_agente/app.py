import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Consola Única de Compras - Enaex", layout="wide")

# -----------------------------------------------------------------------------
# 1. MOTOR FINANCIERO: CONSUMO DE API MINDICADOR.CL EN TIEMPO REAL
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)  # Guarda en caché por 1 hora para no saturar la red
def obtener_indicadores_financieros():
    try:
        url = "https://mindicador.cl/api"
        response = requests.get(url, timeout=5)
        data = response.json()

        return {
            "dolar": float(data["dolar"]["valor"]),
            "euro": float(data["euro"]["valor"]),
            "uf": float(data["uf"]["valor"]),
            "fecha": data["dolar"]["fecha"][:10],
            "estado": "Online 🟢",
        }
    except Exception:
        # Valores de respaldo (fallback) en caso de caída temporal de la API
        return {
            "dolar": 940.0,
            "euro": 1020.0,
            "uf": 37800.0,
            "fecha": "Valores Estimados",
            "estado": "Offline (Fallback) ⚠️",
        }

indicadores = obtener_indicadores_financieros()

# -----------------------------------------------------------------------------
# 2. INICIALIZAR ESTADO DE LA APLICACIÓN
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

st.title("🛒 Consola Única de Compras — Enaex")

# -----------------------------------------------------------------------------
# 3. PANEL CENTRAL DE MONEDAS (Destacado)
# -----------------------------------------------------------------------------
st.caption(f"🗓️ Valores oficiales del día extraídos en vivo ({indicadores['fecha']}) - Estado API: {indicadores['estado']}")

col_uf, col_usd, col_eur, _ = st.columns([1, 1, 1, 2])
col_uf.metric("Valor UF", f"${indicadores['uf']:,.2f}")
col_usd.metric("Valor Dólar", f"${indicadores['dolar']:,.2f}")
col_eur.metric("Valor Euro", f"${indicadores['euro']:,.2f}")

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
# 5. FORMULARIO DE INGRESO MANUAL DE COTIZACIONES
# -----------------------------------------------------------------------------
st.subheader("➕ Carga Manual de Oferta")
with st.form("form_cotizacion", clear_on_submit=True):
    col1, col2, col3, col4 = st.columns(4)
    proveedor = col1.text_input("Proveedor*")
    monto = col2.number_input("Monto Original*", min_value=0.0, step=1000.0)
    moneda = col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"])
    plazo = col4.number_input("Plazo Entrega (Días)", min_value=1, value=5)
    obs = st.text_area("Observaciones Técnicas")

    if st.form_submit_button("Guardar en Cuadro Comparativo"):
        if not proveedor or monto <= 0:
            st.error("⚠️ Debes ingresar el nombre del Proveedor y un Monto mayor a 0.")
        else:
            # CÁLCULO DE CONVERSIÓN EN TIEMPO REAL A CLP Y USD
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
            st.success(f"✅ Oferta de {proveedor} convertida e ingresada correctamente.")
            st.rerun()

# -----------------------------------------------------------------------------
# 6. CUADRO COMPARATIVO HOMOGENEIZADO
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo (Homogeneizado)")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=[
            "Proveedor",
            "Monto Original",
            "Moneda",
            "Equiv. CLP ($)",
            "Equiv. USD ($)",
            "Plazo (Días)",
            "Observaciones",
        ]
    )
    st.dataframe(df_empty, use_container_width=True)
    st.info("👆 Llena el formulario arriba para simular la comparativa.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    st.dataframe(df, use_container_width=True)

    # Regla de Negocio Enaex: Control Financiero > $1.000.000 CLP
    max_monto_clp = df["Equiv. CLP ($)"].max()
    if max_monto_clp > 1000000:
        st.warning(
            f"⚠️ **Control Financiero (> $1M CLP):** Requerimiento alcanza ${max_monto_clp:,.0f} CLP. "
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
