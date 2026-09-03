import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from datetime import datetime, date

# =============================================================================
# CONFIGURACIÓN DE PÁGINA
# =============================================================================
st.set_page_config(
    page_title="Sistema Integrado de Evaluación de Ofertas - Enaex",
    page_icon="⚡",
    layout="wide"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main-header { font-size: 1.8rem; font-weight: 700; color: #1E3A8A; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1rem; color: #4B5563; margin-bottom: 1.5rem; }
    .stTable { font-size: 0.85rem; }
    .metric-card { background-color: #F3F4F6; padding: 1rem; border-radius: 0.5rem; border-left: 4px solid #1E3A8A; }
    .badge-best { background-color: #D1FAE5; color: #065F46; padding: 0.2rem 0.5rem; border-radius: 0.25rem; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# OBTENCIÓN DE INDICADORES FINANCIEROS EN TIEMPO REAL (API)
# =============================================================================
@st.cache_data(ttl=3600)
def obtener_indicadores_tiempo_real():
    """Consulta la API de mindicador.cl para obtener USD, EUR y UF actualizados"""
    valores_defecto = {"USD": 950.0, "EUR": 1020.0, "UF": 38000.0, "estado": False}
    try:
        response = requests.get("https://mindicador.cl/api", timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                "USD": float(data.get("dolar", {}).get("valor", 950.0)),
                "EUR": float(data.get("euro", {}).get("valor", 1020.0)),
                "UF": float(data.get("uf", {}).get("valor", 38000.0)),
                "estado": True
            }
    except Exception:
        pass
    return valores_defecto

# =============================================================================
# FUNCIONES AUXILIARES Y BÚSQUEDA ROBUSTA
# =============================================================================
def procesar_y_reparar_planilla(df):
    if df is None or df.empty:
        return df

    palabras_clave = ['sp', 'solped', 'material', 'pos', 'texto breve', 'centro', 'cantidad']
    header_idx = -1
    
    for idx in range(min(20, len(df))):
        row_values = [str(val).lower() for val in df.iloc[idx]]
        matches = sum(1 for val in row_values for kw in palabras_clave if kw in val)
        if matches >= 2:
            header_idx = idx
            break

    if header_idx != -1:
        nuevas_columnas = []
        for i, val in enumerate(df.iloc[header_idx]):
            val_str = str(val).strip()
            if val_str.lower() in ['nan', 'none', '']:
                col_orig = str(df.columns[i])
                if not col_orig.startswith("Unnamed"):
                    nuevas_columnas.append(col_orig)
                else:
                    nuevas_columnas.append(f"Col_Vacia_{i}")
            else:
                nuevas_columnas.append(val_str)
        
        df.columns = nuevas_columnas
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
        
        if not df.empty:
            primer_col = df.columns[0]
            df = df[df[primer_col].astype(str).str.strip() != str(primer_col).strip()].reset_index(drop=True)

    mapeo_columnas = {
        'Unnamed: 6': 'UM', 'Unnamed: 7': 'Solicitante', 'Unnamed: 8': 'Centro',
        'Unnamed: 9': 'Tipo de posición', 'Unnamed: 10': 'G. compras', 'Unnamed: 11': 'Mod. el',
        'Unnamed: 12': 'Urgencia', 'Unnamed: 13': 'NS', 'Unnamed: 14': 'Contrato marco',
        'Unnamed: 15': 'Observación', 'Unnamed: 16': 'Responsable', 'Unnamed: 41': 'Total general'
    }
    df = df.rename(columns={k: v for k, v in mapeo_columnas.items() if k in df.columns})

    vistos = {}
    columnas_deduplicadas = []
    for c in df.columns:
        c_str = str(c).strip()
        if c_str in vistos:
            vistos[c_str] += 1
            columnas_deduplicadas.append(f"{c_str}_{vistos[c_str]}")
        else:
            vistos[c_str] = 0
            columnas_deduplicadas.append(c_str)
    df.columns = columnas_deduplicadas

    cols = list(df.columns)
    id_col = None
    
    for c in cols:
        if str(c).lower() in ['sp', 'solped', 'solicitud']:
            id_col = c
            break
            
    if not id_col:
        for c in cols:
            if any(kw in str(c).lower() for kw in ['sp', 'solped', 'solicitud', 'pr', 'requerimiento', 'pedido']):
                id_col = c
                break
                
    if id_col and id_col in cols:
        cols.remove(id_col)
        cols.insert(0, id_col)
        df = df[cols]

    df = df.dropna(how='all')
    return df

def extraer_materiales_de_masivo(df, id_solped):
    if df is None or df.empty:
        return []
        
    raw_search = str(id_solped).strip()
    if not raw_search or raw_search.lower() in ["(id solped)", "none", "nan"]:
        return []
        
    digits_search = re.sub(r'\D', '', raw_search)
    
    sp_cols = [c for c in df.columns if any(kw in str(c).lower() for kw in ['sp', 'solped', 'solicitud', 'pr', 'requerimiento', 'doc', 'pedido', 'compra'])]
    if not sp_cols:
        sp_cols = list(df.columns)

    df_filtrado = pd.DataFrame()
    
    for col in sp_cols:
        col_str = df[col].astype(str).str.strip()
        mask = col_str.str.lower() == raw_search.lower()
        
        if not mask.any() and digits_search:
            col_digits = col_str.apply(lambda x: re.sub(r'\D', '', str(x)))
            mask = col_digits == digits_search
            
        if not mask.any():
            mask = col_str.str.lower().str.contains(raw_search.lower(), regex=False)

        if mask.any():
            df_filtrado = df[mask]
            break

    if df_filtrado.empty:
        return []

    posiciones = []
    for idx, row in enumerate(df_filtrado.to_dict('records')):
        def get_val(keys, default):
            for k in keys:
                for col in row.keys():
                    if k in str(col).lower() and pd.notna(row[col]) and str(row[col]).strip() != "":
                        return row[col]
            return default

        def clean_num(val, default=0.0):
            try:
                if isinstance(val, (int, float)): return float(val)
                s = re.sub(r'[^0-9.,-]', '', str(val)).replace(',', '.')
                return float(s)
            except Exception:
                return default

        posiciones.append({
            "Pos": int(idx + 1),
            "Material": str(get_val(['texto', 'desc', 'material', 'denominacion', 'item', 'artículo', 'articulo', 'breve'], f"Material {idx+1}")),
            "Centro": str(get_val(['centro', 'plant', 'almacen', 'alm'], "E001")),
            "Cantidad": clean_num(get_val(['cant', 'cantidad', 'ctd'], 1.0), 1.0),
            "UM": str(get_val(['um', 'unidad', 'unid', 'medida'], "C/U")).upper(),
            "Precio Unitario": clean_num(get_val(['precio', 'monto', 'val', 'costo', 'p.u', 'neto'], 0.0), 0.0),
            "Moneda": str(get_val(['moneda', 'curr', 'mon'], "CLP")).upper(),
            "Proveedor": str(get_val(['proveedor', 'vendor', 'prov', 'nam'], "")),
            "Calendario de entrega": date.today(),
            "Observaciones": str(get_val(['obs', 'observacion', 'comentario'], ""))
        })
        
    return posiciones

def convertir_moneda(monto, moneda_origen, tc_usd, tc_uf, tc_eur):
    """Calcula la equivalencia en las tres monedas principales"""
    monto = float(monto or 0.0)
    moneda_origen = str(moneda_origen).upper()
    
    if moneda_origen == "CLP":
        clp = monto
    elif moneda_origen == "USD":
        clp = monto * tc_usd
    elif moneda_origen == "UF":
        clp = monto * tc_uf
    elif moneda_origen == "EUR":
        clp = monto * tc_eur
    else:
        clp = monto
        
    usd = clp / tc_usd if tc_usd > 0 else 0.0
    eur = clp / tc_eur if tc_eur > 0 else 0.0
    return clp, usd, eur

# =============================================================================
# INICIALIZACIÓN DE ESTADO
# =============================================================================
if "df_masivo" not in st.session_state:
    st.session_state.df_masivo = None
if "ofertas_manuales" not in st.session_state:
    st.session_state.ofertas_manuales = []

# =============================================================================
# ENCABEZADO Y PARÁMETROS GLOBALES
# =============================================================================
st.markdown("<div class='main-header'>⚡ Sistema Integrado de Evaluación de Ofertas</div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Parámetros de Cambio")
    
    indicadores = obtener_indicadores_tiempo_real()
    
    if indicadores["estado"]:
        st.success("🟢 Indicadores actualizados en tiempo real")
    else:
        st.warning("⚠️ Sin conexión a API. Usando valores por defecto.")
        
    if st.button("🔄 Actualizar Tasas API", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    tc_usd = st.number_input("Tipo de Cambio USD / CLP", value=indicadores["USD"], step=1.0, format="%.2f")
    tc_uf = st.number_input("Tipo de Cambio UF / CLP", value=indicadores["UF"], step=100.0, format="%.2f")
    tc_eur = st.number_input("Tipo de Cambio EUR / CLP", value=indicadores["EUR"], step=1.0, format="%.2f")
    st.divider()
    
    st.header("📂 Carga de Archivo Base")
    file_masivo = st.file_uploader("Cargar Planilla Maestro/SOLPEDs (Excel/CSV)", type=["xlsx", "xls", "csv", "xlsm"])
    
    if file_masivo:
        try:
            if file_masivo.name.endswith(".csv"):
                df_raw = pd.read_csv(file_masivo)
            else:
                dict_dfs = pd.read_excel(file_masivo, sheet_name=None, engine='openpyxl')
                df_raw = pd.concat(dict_dfs.values(), ignore_index=True)
                
            df_clean = df_raw.dropna(axis=1, how='all').dropna(axis=0, how='all')
            df_procesado = procesar_y_reparar_planilla(df_clean)
            
            df_procesado['cantidad_nulos'] = df_procesado.isnull().sum(axis=1)
            df_procesado = df_procesado.sort_values(by='cantidad_nulos').drop(columns=['cantidad_nulos']).reset_index(drop=True)
            
            st.session_state.df_masivo = df_procesado
            st.success(f"Planilla cargada y limpiada correctamente ({len(st.session_state.df_masivo)} filas)")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

# =============================================================================
# VISTA PREVIA DE DATOS CARGADOS
# =============================================================================
if st.session_state.df_masivo is not None:
    with st.expander("👀 Vista Previa de la Planilla Base Cargada", expanded=False):
        st.write(f"Mostrando los datos procesados. La columna 'SP' ha sido priorizada en la primera posición para fácil lectura.")
        st.dataframe(st.session_state.df_masivo, use_container_width=True)

tabs = st.tabs(["✏️ Evaluación por SOLPED", "➕ Carga Manual / Directa", "📊 Cuadro Comparativo Integrado"])

# =============================================================================
# TAB 1: EVALUACIÓN POR SOLPED (AUTOGESTIÓN)
# =============================================================================
with tabs[0]:
    st.subheader("✏️ Evaluación por SOLPED")
    
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        solped_id = st.text_input("Buscar ID SOLPED en la planilla:", placeholder="Ej: PR175798 o 175798")
    with col_btn:
        st.write("")
        st.write("")
        btn_extraer = st.button("📤 Extraer Materiales", type="primary", use_container_width=True)

    if (btn_extraer or solped_id) and solped_id.strip():
        if st.session_state.df_masivo is not None:
            materiales = extraer_materiales_de_masivo(st.session_state.df_masivoPara implementar un selector que cambie dinámicamente la moneda en la que se visualizan los montos totales, la mejor opción es modificar la pestaña del **Cuadro Comparativo Integrado** (Tab 3). 

Dado que el archivo app (17).py ya calcula y almacena el `Total CLP` y cuenta con las variables de conversión en tiempo real (`tc_usd`, `tc_eur`, `tc_uf`)[cite: 1], puedes usar un `st.radio` para aplicar el tipo de cambio al vuelo sobre la tabla y las métricas.

Reemplaza todo el bloque de código de tu **TAB 3** con la siguiente versión:

```python
# =============================================================================
# TAB 3: CUADRO COMPARATIVO INTEGRADO
# =============================================================================
with tabs[2]:
    st.subheader("📊 Cuadro Comparativo Integrado")
    
    if st.session_state.ofertas_manuales:
        df_comp = pd.DataFrame(st.session_state.ofertas_manuales)
        
        # 1. Selector de moneda para la visualización
        moneda_vista = st.radio(
            "💱 Mostrar valores convertidos en:", 
            options=["CLP", "USD", "EUR", "UF"], 
            horizontal=True
        )
        
        # 2. Copia del dataframe para no alterar los datos base
        df_vista = df_comp.copy()
        
        # 3. Conversión dinámica basada en el Total CLP y los parámetros de la barra lateral
        if moneda_vista == "CLP":
            df_vista["Total Visualizado"] = df_vista["Total CLP"]
        elif moneda_vista == "USD":
            df_vista["Total Visualizado"] = df_vista["Total CLP"] / tc_usd
        elif moneda_vista == "EUR":
            df_vista["Total Visualizado"] = df_vista["Total CLP"] / tc_eur
        elif moneda_vista == "UF":
            df_vista["Total Visualizado"] = df_vista["Total CLP"] / tc_uf
            
        # Opcional: Dar formato visual a la nueva columna
        df_vista["Total Visualizado"] = df_vista["Total Visualizado"].apply(lambda x: f"{x:,.2f} {moneda_vista}")
        
        # Reordenar columnas para que el Total Visualizado destaque (al final)
        cols = list(df_vista.columns)
        cols.append(cols.pop(cols.index("Total Visualizado")))
        df_vista = df_vista[cols]
        
        st.dataframe(df_vista, use_container_width=True)
        
        # 4. Actualización dinámica de las métricas inferiores
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.metric("Total Ofertas Registradas", len(df_vista))
        with col_c2:
            total_base_clp = df_comp['Total CLP'].sum()
            
            if moneda_vista == "CLP":
                monto_final = total_base_clp
            elif moneda_vista == "USD":
                monto_final = total_base_clp / tc_usd
            elif moneda_vista == "EUR":
                monto_final = total_base_clp / tc_eur
            elif moneda_vista == "UF":
                monto_final = total_base_clp / tc_uf
                
            st.metric(f"Monto Total Acumulado ({moneda_vista})", f"{monto_final:,.2f}")
            
        if st.button("🗑️ Limpiar Cuadro Comparativo"):
            st.session_state.ofertas_manuales = []
            st.rerun()
    else:
        st.info("Aún no hay ofertas registradas en el Cuadro Comparativo.")
