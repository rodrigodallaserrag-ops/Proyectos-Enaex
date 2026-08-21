import pandas as pd
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")
st.title("🎯 Dashboard OTIF (On Time, In Full)")
st.markdown("Analiza el cumplimiento de entregas de los proveedores.")

# ==========================================
# 2. FUNCIÓN DE PROCESAMIENTO (Caché para que sea rápido)
# ==========================================
@st.cache_data(show_spinner="Procesando cruce de datos...")
def procesar_otif(file_me2m, file_me80fn):
    # Leer los archivos
    df_me2m = pd.read_excel(file_me2m)
    df_me80fn = pd.read_excel(file_me80fn)

    # Agrupar recepciones (ME80FN) por OC y Posición
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Cantidad_Recibida=('Cantidad', 'sum'),
        Fecha_Real_Entrega=('Fe.contabilización', 'max')
    ).reset_index()

    # Cruzar ME2M con las recepciones (Left join)
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

    # Rellenar con 0 las cantidades de lo que aún no ha llegado nada
    df_otif['Cantidad_Recibida'] = df_otif['Cantidad_Recibida'].fillna(0)

    # Asegurar formato de fechas
    df_otif['Fecha de entrega'] = pd.to_datetime(df_otif['Fecha de entrega']).dt.date
    df_otif['Fecha_Real_Entrega'] = pd.to_datetime(df_otif['Fecha_Real_Entrega']).dt.date

    # Lógica de los indicadores
    # In Full: Recibió todo lo pedido
    df_otif['In_Full'] = df_otif['Cantidad_Recibida'] >= df_otif['Cantidad de pedido']
    
    # On Time: Tiene fecha de entrega real y llegó antes o igual a la promesa
    df_otif['On_Time'] = df_otif.apply(
        lambda x: True if pd.notna(x['Fecha_Real_Entrega']) and (x['Fecha_Real_Entrega'] <= x['Fecha de entrega']) else False, 
        axis=1
    )
    
    # OTIF: Cumple In Full y On Time simultáneamente
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    return df_otif

# ==========================================
# 3. BARRA LATERAL (Carga de archivos)
# ==========================================
with st.sidebar:
    st.header("📂 Carga de Datos")
    st.info("Sube los dos reportes descargados de SAP para calcular el OTIF.")
    
    archivo_me2m = st.file_uploader("1. Sube el reporte ME2M (Promesas)", type=["xlsx"])
    archivo_me80fn = st.file_uploader("2. Sube el reporte ME80FN (Entregas Reales)", type=["xlsx"])

# ==========================================
# 4. CUERPO PRINCIPAL (Filtros y KPIs)
# ==========================================
if archivo_me2m and archivo_me80fn:
    # Procesar los datos
    df_resultado = procesar_otif(archivo_me2m, archivo_me80fn)

    st.divider()
    
    # --- FILTROS ---
    st.subheader("🔍 Filtros de Análisis")
    col1, col2 = st.columns(2)
    
    with col1:
        centros = sorted(df_resultado['Centro'].dropna().unique())
        centro_sel = st.multiselect("Filtrar por Centro Logístico", centros, help="Deja en blanco para ver todos.")
        
    with col2:
        grupos = sorted(df_resultado['Grupo de compras'].dropna().unique())
        grupo_sel = st.multiselect("Filtrar por Grupo de Compras", grupos, help="Deja en blanco para ver todos.")

    # Aplicar filtros
    df_filtrado = df_resultado.copy()
    if centro_sel:
        df_filtrado = df_filtrado[df_filtrado['Centro'].isin(centro_sel)]
    if grupo_sel:
        df_filtrado = df_filtrado[df_filtrado['Grupo de compras'].isin(grupo_sel)]

    # --- MÉTRICAS ---
    total_lineas = len(df_filtrado)
    if total_lineas > 0:
        pct_on_time = (df_filtrado['On_Time'].sum() / total_lineas) * 100
        pct_in_full = (df_filtrado['In_Full'].sum() / total_lineas) * 100
        pct_otif = (df_filtrado['OTIF'].sum() / total_lineas) * 100
    else:
        pct_on_time = pct_in_full = pct_otif = 0

    st.divider()
    st.subheader("📊 Resumen General")
    
    # Usamos las tarjetas nativas de Streamlit (metrics) que son muy limpias
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Líneas de Pedido", f"{total_lineas:,}")
    kpi2.metric("⏱️ On Time (A Tiempo)", f"{pct_on_time:.1f}%")
    kpi3.metric("📦 In Full (Completos)", f"{pct_in_full:.1f}%")
    kpi4.metric("⭐ Nivel OTIF", f"{pct_otif:.1f}%")

    # --- TABLA DE DETALLES ---
    st.divider()
    st.subheader("📋 Detalle de Órdenes de Compra")
    
    # Preparamos una tabla más limpia para mostrar al usuario
    columnas_mostrar = [
        'Documento compras', 'Posición', 'Centro', 'Proveedor/Centro suministrador', 
        'Texto breve', 'Cantidad de pedido', 'Cantidad_Recibida', 
        'Fecha de entrega', 'Fecha_Real_Entrega', 'On_Time', 'In_Full', 'OTIF'
    ]
    
    # Dejamos solo las columnas que importan y renombramos algunas para que se vean mejor
    df_mostrar = df_filtrado[[c for c in columnas_mostrar if c in df_filtrado.columns]].copy()
    
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

else:
    # Mensaje de bienvenida si no hay archivos cargados
    st.warning("👈 Por favor, sube los archivos **ME2M** y **ME80FN** en el menú de la izquierda para comenzar.")
