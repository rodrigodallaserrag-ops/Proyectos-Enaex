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
    """
    Renombra las columnas 'Unnamed' explícitamente, promueve encabezados si es necesario, 
    deduplica columnas idénticas y mueve la columna de ID al principio del DataFrame.
    """
    if df is None or df.empty:
        return df

    # 1. Mapeo explícito basado en las imágenes de referencia
    mapeo_columnas = {
        'Unnamed: 6': 'UM',
        'Unnamed: 7': 'Solicitante',
        'Unnamed: 8': 'Centro',
        'Unnamed: 9': 'Tipo de posición',
        'Unnamed: 10': 'G. compras',
        'Unnamed: 11': 'Mod. el',
        'Unnamed: 12': 'Urgencia',
        'Unnamed: 13': 'NS',
        'Unnamed: 14': 'Contrato marco',
        'Unnamed: 15': 'Observación',
        'Unnamed: 16': 'Responsable',
        'Unnamed: 41': 'Total general'
    }
    df = df.rename(columns=mapeo_columnas)

    # 2. Intentar promover la fila de encabezados real si aún quedan muchas "Unnamed"
    unnamed_cols = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if len(unnamed_cols) > 0:
        for idx in range(min(10, len(df))):
            row = df.iloc[idx]
            celdas_validas = [str(v).strip() for v in row if pd.notna(v) and str(v).strip().lower() not in ['nan', 'none', '']]
            
            if len(celdas_validas) >= (len(df.columns) * 0.3):
                nuevas_columnas = []
                for i, v in enumerate(row):
                    v_str = str(v).strip()
                    if pd.notna(v) and v_str != '' and not v_str.startswith("Unnamed:"):
                        nuevas_columnas.append(v_str)
                    else:
                        nuevas_columnas.append(df.columns[i]) # Mantener el nombre ya mapeado
                
                df.columns = nuevas_columnas
                df = df.iloc[idx + 1:].reset_index(drop=True)
                break

    # 3. Eliminar filas duplicadas (ej. la fila que se usó como encabezado)
    if not df.empty:
        primer_col = df.columns[0]
        mask_repetida = df[primer_col].astype(str).str.lower().str.strip() == str(primer_col).lower().strip()
        df = df[~mask_repetida].reset_index(drop=True)

    df.columns = [str(c).strip() for c in df.columns]

    # --- FIX: Deduplicar nombres de columnas para evitar el error de PyArrow ---
    vistos = {}
    columnas_deduplicadas = []
    for c in df.columns:
        if c in vistos:
            vistos[c] += 1
            columnas_deduplicadas.append(f"{c}_{vistos[c]}")
        else:
            vistos[c] = 0
            columnas_deduplicadas.append(c)
    df.columns = columnas_deduplicadas
    # ---------------------------------------------------------------------------

    # 4. Encontrar la columna de ID (SOLPED) y moverla al principio
    cols = list(df.columns)
    id_col = None
    for c in cols:
        if any(kw in str(c).lower() for kw in ['sp', 'solped', 'solicitud', 'pr', 'requerimiento', 'pedido']):
            id_col = c
            break
            
    if id_col and id_col in cols:
        cols.remove(id_col)
        cols.insert(0, id_col)
        df = df[cols]

    return df

def extraer_materiales_de_masivo(df, id_solped):
    """
    Busca de manera flexible el ID SOLPED en la planilla masiva.
    """
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
    """Convierte importes a CLP y USD"""
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
    return clp, usd

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
                
            # Limpieza inicial
            df_clean = df_raw.dropna(axis=1, how='all').dropna(axis=0, how='all')
            
            # Aplicar la nueva función de mapeo y ordenamiento
            df_procesado = procesar_y_reparar_planilla(df_clean)
            
            # Ordenar por completitud (opcional, para visualización limpia)
            df_procesado['cantidad_nulos'] = df_procesado.isnull().sum(axis=1)
            df_procesado = df_procesado.sort_values(by='cantidad_nulos').drop(columns=['cantidad_nulos']).reset_index(drop=True)
            
            st.session_state.df_masivo = df_procesado
            st.success(f"Planilla cargada y mapeada correctamente ({len(st.session_state.df_masivo)} filas)")
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

# =============================================================================
# VISTA PREVIA DE DATOS CARGADOS
# =============================================================================
if st.session_state.df_masivo is not None:
    with st.expander("👀 Vista Previa de la Planilla Base Cargada", expanded=False):
        st.write(f"Mostrando los datos procesados con columnas corregidas e ID al inicio ({len(st.session_state.df_masivo)} registros en total):")
        st.dataframe(st.session_state.df_masivo, use_container_width=True)

tabs = st.tabs(["✏️ Evaluación por SOLPED", "➕ Carga Manual / Directa", "📊 Cuadro Comparativo Integrado"])

# =============================================================================
# TAB 1: EVALUACIÓN POR SOLPED (AUTOGESTIÓN)
# =============================================================================
with tabs[0]:
    st.subheader("✏️ Evaluación por SOLPED (Soporta 20+ Materiales sin límite de filas)")
    
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        solped_id = st.text_input("Buscar ID SOLPED en la planilla:", placeholder="Ej: PR175798 o 175798")
    with col_btn:
        st.write("")
        st.write("")
        btn_extraer = st.button("📤 Extraer Materiales", type="primary", use_container_width=True)

    if (btn_extraer or solped_id) and solped_id.strip():
        if st.session_state.df_masivo is not None:
            materiales = extraer_materiales_de_masivo(st.session_state.df_masivo, solped_id)
            if materiales:
                st.session_state[f"editor_{solped_id}"] = pd.DataFrame(materiales)
                st.success(f"Se encontraron {len(materiales)} posiciones para la SOLPED **{solped_id}**")
            else:
                st.warning(f"No se encontraron registros para la SOLPED '{solped_id}'. Verifica si fue cargada en el panel lateral.")
        else:
            st.info("Carga una planilla maestra en el menú lateral para realizar la búsqueda automática por SOLPED.")

    key_editor = f"editor_{solped_id}" if (solped_id and f"editor_{solped_id}" in st.session_state) else "editor_default"
    
    df_inicial = st.session_state.get(key_editor, pd.DataFrame([{
        "Pos": 1, "Material": "(Material)", "Centro": "(Centro)", "Cantidad": 1.0, 
        "UM": "C/U", "Precio Unitario": 0.0, "Moneda": "CLP", 
        "Proveedor": "", "Calendario de entrega": date.today(), "Observaciones": ""
    }]))

    if not df_inicial.empty:
        df_inicial["Precio Unitario"] = pd.to_numeric(df_inicial["Precio Unitario"], errors='coerce').fillna(0.0)
        df_inicial["Cantidad"] = pd.to_numeric(df_inicial["Cantidad"], errors='coerce').fillna(1.0)
        df_inicial["Calendario de entrega"] = pd.to_datetime(df_inicial["Calendario de entrega"]).dt.date

    edited_df = st.data_editor(
        df_inicial,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Pos": st.column_config.NumberColumn("Pos", disabled=True),
            "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "UF", "EUR"]),
            "Calendario de entrega": st.column_config.DateColumn("Fecha Entrega")
        }
    )

    if st.button("💾 Guardar Oferta de SOLPED en Comparativo", type="primary"):
        registros = edited_df.to_dict('records')
        for r in registros:
            clp, usd = convertir_moneda(r["Precio Unitario"] * r["Cantidad"], r["Moneda"], tc_usd, tc_uf, tc_eur)
            r["SOLPED"] = solped_id if solped_id else "N/A"
            r["Total CLP"] = clp
            r["Total USD"] = usd
            st.session_state.ofertas_manuales.append(r)
        st.success("¡Oferta guardada exitosamente en el Cuadro Comparativo!")

# =============================================================================
# TAB 2: CARGA MANUAL INTEGRA / EDICIÓN DIRECTA POR SOLPED
# =============================================================================
with tabs[1]:
    st.subheader("➕ Carga Manual de Oferta Paso a Paso")
    
    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        manual_solped = st.text_input("Ingresar N° SOLPED para Autocompletar:", placeholder="Ej: PR175798")
    with col_s2:
        st.write("")
        st.write("")
        btn_cargar_manual = st.button("📥 Cargar Requerimiento", use_container_width=True)

    if btn_cargar_manual and manual_solped:
        if st.session_state.df_masivo is not None:
            mats = extraer_materiales_de_masivo(st.session_state.df_masivo, manual_solped)
            if mats:
                st.session_state["manual_grid_df"] = pd.DataFrame(mats)
                st.success(f"Materiales cargados automáticamente desde la SOLPED {manual_solped}")
            else:
                st.warning(f"No se encontró la SOLPED {manual_solped} en el archivo base.")
        else:
            st.info("Sube una planilla en la barra lateral para autocompletar posiciones por SOLPED.")

    if "manual_grid_df" not in st.session_state:
        st.session_state["manual_grid_df"] = pd.DataFrame([{
            "Pos": 1, "Material": "Ítem Manual", "Cantidad": 1.0, "UM": "C/U",
            "Precio Unitario": 0.0, "Moneda": "CLP", "Proveedor": "", 
            "Calendario de entrega": date.today(), "Observaciones": ""
        }])

    df_manual = st.session_state["manual_grid_df"]
    if not df_manual.empty:
        df_manual["Precio Unitario"] = pd.to_numeric(df_manual["Precio Unitario"], errors='coerce').fillna(0.0)
        df_manual["Cantidad"] = pd.to_numeric(df_manual["Cantidad"], errors='coerce').fillna(1.0)
        df_manual["Calendario de entrega"] = pd.to_datetime(df_manual["Calendario de entrega"]).dt.date

    st.write("### Tabla de Cotización de Proveedor")
    
    cotizacion_df = st.data_editor(
        df_manual,
        num_rows="dynamic",
        use_container_width=True,
        key="cotizacion_manual_editor",
        column_config={
            "Precio Unitario": st.column_config.NumberColumn("Precio Unitario", format="$ %.2f"),
            "Moneda": st.column_config.SelectboxColumn("Moneda", options=["CLP", "USD", "UF", "EUR"]),
            "Calendario de entrega": st.column_config.DateColumn("Calendario de entrega")
        }
    )

    if st.button("💾 Guardar Cotización Manual Completa", type="primary"):
        items = cotizacion_df.to_dict('records')
        for item in items:
            clp, usd = convertir_moneda(item["Precio Unitario"] * item["Cantidad"], item["Moneda"], tc_usd, tc_uf, tc_eur)
            item["SOLPED"] = manual_solped if manual_solped else "MANUAL"
            item["Total CLP"] = clp
            item["Total USD"] = usd
            st.session_state.ofertas_manuales.append(item)
        st.success("¡Cotización agregada al Cuadro Comparativo!")

# =============================================================================
# TAB 3: CUADRO COMPARATIVO INTEGRADO
# =============================================================================
with tabs[2]:
    st.subheader("📊 Cuadro Comparativo Integrado")
    
    if st.session_state.ofertas_manuales:
        df_comp = pd.DataFrame(st.session_state.ofertas_manuales)
        st.dataframe(df_comp, use_container_width=True)
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.metric("Total Ofertas Registradas", len(df_comp))
        with col_c2:
            st.metric("Monto Total Acumulado (CLP)", f"$ {df_comp['Total CLP'].sum():,.2f}")
            
        if st.button("🗑️ Limpiar Cuadro Comparativo"):
            st.session_state.ofertas_manuales = []
            st.rerun()
    else:
        st.info("Aún no hay ofertas registradas en el Cuadro Comparativo.")
