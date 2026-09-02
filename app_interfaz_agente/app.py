import io
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
    return f"${int(valor):,}".replace(",", ".")

def aplicar_formato_regional(monto, moneda):
    if moneda == "CLP":
        return f"$ {int(monto):,}".replace(",", ".")
    elif moneda == "USD":
        return f"$ {monto:,.2f}"
    elif moneda == "EUR":
        return f"€ {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    elif moneda == "UF":
        return f"UF {monto:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(monto)

# -----------------------------------------------------------------------------
# 2. INICIALIZAR ESTADO Y CALLBACK DE SANITIZACIÓN NUMÉRICA
# -----------------------------------------------------------------------------
if "cotizaciones" not in st.session_state:
    st.session_state["cotizaciones"] = []

if "monto_input" not in st.session_state:
    st.session_state["monto_input"] = ""

if "moneda_input" not in st.session_state:
    st.session_state["moneda_input"] = "CLP"

def formatear_caja_monto():
    """Filtra y elimina cualquier caracter no numérico (letras, símbolos)"""
    raw = str(st.session_state["monto_input"]).strip()
    moneda = st.session_state["moneda_input"]
    
    if not raw:
        return

    # 1. Filtro estricto: Elimina cualquier letra o símbolo que NO sea un número, punto o coma
    solo_numeros = re.sub(r'[^0-9.,]', '', raw)
    
    # Si al quitar letras no queda nada (ej: escribieron "AWDADW"), borra el contenido de la caja
    if not solo_numeros:
        st.session_state["monto_input"] = ""
        return

    try:
        # 2. Conversión según la norma de la moneda seleccionada
        if moneda == "USD":
            limpio = solo_numeros.replace(",", "")
        else:
            limpio = solo_numeros.replace(".", "").replace(",", ".")
            
        num = float(limpio)
        
        # 3. Reescribe la caja con el formato correcto aplicado
        if moneda == "CLP":
            st.session_state["monto_input"] = f"{int(num):,}".replace(",", ".")
        elif moneda == "USD":
            st.session_state["monto_input"] = f"{num:,.2f}"
        elif moneda in ["EUR", "UF"]:
            st.session_state["monto_input"] = f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
    except ValueError:
        st.session_state["monto_input"] = ""

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
# 5. INGRESO DE COTIZACIONES
# -----------------------------------------------------------------------------
st.subheader("➕ Carga Manual de Oferta")

# Función que procesa los datos y limpia las cajas ANTES de redibujar la pantalla
def procesar_guardado():
    raw = str(st.session_state.get("monto_input", "")).strip()
    moneda = st.session_state.get("moneda_input", "CLP")
    proveedor = st.session_state.get("proveedor_input", "")
    plazo = st.session_state.get("plazo_input", 5)
    obs = st.session_state.get("obs_input", "")
    
    try:
        if moneda == "USD":
            monto = float(raw.replace(",", ""))
        else:
            monto = float(raw.replace(".", "").replace(",", "."))
    except ValueError:
        monto = 0.0

    if not proveedor or monto <= 0:
        st.toast("⚠️ Debes ingresar un Proveedor y un Monto numérico válido.", icon="🚨")
    else:
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
        
        # Limpiamos las cajas de forma segura sin generar el error de instanciación
        st.session_state["monto_input"] = ""
        st.session_state["proveedor_input"] = ""
        st.session_state["obs_input"] = ""
        st.toast(f"✅ Oferta de {proveedor} ingresada correctamente.", icon="✅")

col1, col2, col3, col4 = st.columns(4)

col1.text_input("Proveedor*", key="proveedor_input")
col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"], key="moneda_input", on_change=formatear_caja_monto)
col2.text_input("Monto Original*", placeholder="Solo números (Ej: 190000)", key="monto_input", on_change=formatear_caja_monto)
col4.number_input("Plazo Entrega (Días)", min_value=1, value=5, key="plazo_input")
st.text_area("Observaciones Técnicas", key="obs_input")

st.button("Guardar en Cuadro Comparativo", type="primary", on_click=procesar_guardado)

# -----------------------------------------------------------------------------
# 6. CUADRO COMPARATIVO HOMOGENEIZADO Y DESCARGA A EXCEL
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo (Homogeneizado)")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Plazo (Días)", "Observaciones"]
    )
    st.dataframe(df_empty, use_container_width=True)
    st.info("👆 Agrega ofertas arriba para visualizar el cuadro comparativo.")
else:
    df = pd.DataFrame(st.session_state["cotizaciones"])
    
    df_visual = df.copy()
    df_visual["Monto Original"] = df_visual.apply(lambda fila: aplicar_formato_regional(fila["Monto Original"], fila["Moneda"]), axis=1)
    df_visual["Equiv. CLP ($)"] = df_visual["Equiv. CLP ($)"].apply(lambda x: f"$ {int(x):,}".replace(",", "."))
    df_visual["Equiv. USD ($)"] = df_visual["Equiv. USD ($)"].apply(lambda x: f"$ {x:,.2f}")

    st.dataframe(df_visual, use_container_width=True)

    max_monto_clp = df["Equiv. CLP ($)"].max()
    if max_monto_clp > 1000000:
        st.warning(
            f"⚠️ **Control Financiero (> $1M CLP):** Requerimiento alcanza {formato_clp(max_monto_clp)}. "
            f"Se aplicaron los tipos de cambio oficiales del día ({indicadores['fecha']})."
        )

    # Función para convertir el DataFrame a Excel en memoria
    def convertir_excel(dataframe):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name='Comparativo')
        return output.getvalue()

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("🗑️ Limpiar Tabla"):
            st.session_state["cotizaciones"] = []
            st.rerun()

    with col_btn2:
        # Preparamos el archivo Excel usando el DataFrame con datos numéricos puros
        excel_data = convertir_excel(df)
        
        st.download_button(
            label="🚀 Emitir a SAP ME21N (Descargar Excel)",
            data=excel_data,
            file_name=f"Comparativo_Solped_{solped}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
