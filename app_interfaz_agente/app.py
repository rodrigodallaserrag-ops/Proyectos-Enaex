import datetime
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

    solo_numeros = re.sub(r'[^0-9.,]', '', raw)
    
    if not solo_numeros:
        st.session_state["monto_input"] = ""
        return

    try:
        if moneda == "USD":
            limpio = solo_numeros.replace(",", "")
        else:
            limpio = solo_numeros.replace(".", "").replace(",", ".")
            
        num = float(limpio)
        
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

def procesar_guardado():
    raw = str(st.session_state.get("monto_input", "")).strip()
    moneda = st.session_state.get("moneda_input", "CLP")
    proveedor = st.session_state.get("proveedor_input", "")
    fecha_entrega = st.session_state.get("fecha_entrega_input", datetime.date.today())
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
        
        hoy = datetime.date.today()
        dias_diferencia = (fecha_entrega - hoy).days
        fecha_formateada = fecha_entrega.strftime("%d-%m-%Y")
        
        texto_dias = "día" if dias_diferencia == 1 else "días"
        plazo_final = f"{fecha_formateada} ({dias_diferencia} {texto_dias})"

        st.session_state["cotizaciones"].append({
            "Proveedor": proveedor,
            "Monto Original": monto, 
            "Moneda": moneda,
            "Equiv. CLP ($)": round(monto_clp, 2),
            "Equiv. USD ($)": round(monto_usd, 2),
            "Fecha de Entrega": plazo_final,
            "Observaciones": obs,
        })
        
        st.session_state["monto_input"] = ""
        st.session_state["proveedor_input"] = ""
        st.session_state["obs_input"] = ""
        st.toast(f"✅ Oferta de {proveedor} ingresada correctamente.", icon="✅")

col1, col2, col3, col4 = st.columns(4)

col1.text_input("Proveedor*", key="proveedor_input")
col3.selectbox("Moneda", ["CLP", "USD", "EUR", "UF"], key="moneda_input", on_change=formatear_caja_monto)
col2.text_input("Monto Original*", placeholder="Solo números (Ej: 190000)", key="monto_input", on_change=formatear_caja_monto)
col4.date_input("Fecha de Entrega", min_value=datetime.date.today(), key="fecha_entrega_input")

st.text_area("Observaciones Técnicas", key="obs_input")
st.button("Guardar en Cuadro Comparativo", on_click=procesar_guardado)

# -----------------------------------------------------------------------------
# 6. CUADRO COMPARATIVO HOMOGENEIZADO Y DESCARGA A EXCEL
# -----------------------------------------------------------------------------
st.divider()
st.subheader("📊 Cuadro Comparativo (Homogeneizado)")

if not st.session_state["cotizaciones"]:
    df_empty = pd.DataFrame(
        columns=["Proveedor", "Monto Original", "Moneda", "Equiv. CLP ($)", "Equiv. USD ($)", "Fecha de Entrega", "Observaciones"]
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
    
    # --- SISTEMA DE ELIMINACIÓN INDIVIDUAL ---
    st.markdown("#### 🛠️ Gestionar Ofertas Ingresadas")
    
    for i, cotizacion in enumerate(st.session_state["cotizaciones"]):
        col_info, col_btn = st.columns([5, 1])
        
        monto_clp_formateado = f"$ {int(cotizacion['Equiv. CLP ($)']):,}".replace(",", ".")
        col_info.markdown(f"Oferta de **{cotizacion['Proveedor']}** por **{monto_clp_formateado} CLP**")
        
        # Botón rojo ("primary") y con texto exacto "Eliminar"
        if col_btn.button("Eliminar", type="primary", key=f"eliminar_fila_{i}"):
            st.session_state["cotizaciones"].pop(i)
            st.rerun()
            
    st.write("---")

    max_monto_clp = df["Equiv. CLP ($)"].max()
    if max_monto_clp > 1000000:
        st.warning(
            f"⚠️ **Control Financiero (> $1M CLP):** Requerimiento alcanza {formato_clp(max_monto_clp)}. "
            f"Se aplicaron los tipos de cambio oficiales del día ({indicadores['fecha']})."
        )

    def convertir_excel(dataframe):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            dataframe.to_excel(writer, index=False, sheet_name='Comparativo')
        return output.getvalue()

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("Aquí tienes un par de formas de hacerlo, dependiendo de la tecnología que estés utilizando:

**HTML y CSS (Estilos en línea)**
La forma más rápida si solo necesitas un botón simple:
```html
<button style="background-color: red; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer;">
  Eliminar
</button>
