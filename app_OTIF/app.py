import pandas as pd
import streamlit as st

st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")
st.title("🎯 Dashboard OTIF (On Time, In Full)")
st.markdown("Medición del nivel de servicio de proveedores cruzando archivos de SAP.")

# ==========================================
# FUNCIÓN EN CACHÉ (Optimizada para memoria)
# ==========================================
@st.cache_data(show_spinner="Procesando cruce de datos...")
def procesar_otif(file_me2m, file_me80fn):
    # Leemos los Excel
    df_me2m = pd.read_excel(file_me2m, engine="openpyxl")
    # Para ME80FN, solo nos interesan ciertas columnas para ahorrar memoria
    cols_me80fn = ['Documento compras', 'Posición', 'Fe.contabilización']
    df_me80fn = pd.read_excel(file_me80fn, engine="openpyxl", usecols=lambda c: c in cols_me80fn)

    # 1. Descartar posiciones anuladas en SAP
    if 'Indicador de borrado' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Indicador de borrado'].isna()].copy()

    # 2. Obtener fecha de la última recepción real en ME80FN
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Fecha_Ingreso_SAP=('Fe.contabilización', 'max')
    ).reset_index()

    # 3. Cruzar ME2M con ME80FN
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

    # 4. Formato de fechas
    df_otif['Fecha_Estadistica'] = pd.to_datetime(df_otif['Fecha entrega estad.'], errors='coerce').dt.date
    df_otif['Fecha_Ingreso_SAP'] = pd.to_datetime(df_otif['Fecha_Ingreso_SAP'], errors='coerce').dt.date

    # 5. Evaluación de Reglas
    df_otif['In_Full'] = df_otif['Por entregar (cantidad)'] == 0
    df_otif['On_Time'] = (
        df_otif['Fecha_Ingreso_SAP'].notna() & 
        df_otif['Fecha_Estadistica'].notna() & 
        (df_otif['Fecha_Ingreso_SAP'] <= df_otif['Fecha_Estadistica'])
    )
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    return df_otif

# ==========================================
# BARRA LATERAL (Carga)
# ==========================================
with st.sidebar:
    st.header("📂 Carga de Datos")
    archivo_me2m = st.file_uploader("1. Sube ME2M (.xlsx)", type=["xlsx"])
    archivo_me80fn = st.file_uploader("2. Sube ME80FN (.xlsx)", type=["xlsx"])

# ==========================================
# CUERPO PRINCIPAL
# ==========================================
if archivo_me2m and archivo_me80fn:
    df_base = procesar_otif(archivo_me2m, archivo_me80fn)

    # Filtros
    st.markdown("**🔍 Filtros de Análisis**")
    f_col1, f_col2, f_col3 = st.columns(3)

    with f_col1:
        centros = sorted(df_base['Centro'].dropna().unique())
        centro_sel = st.multiselect("Centro Logístico", centros)

    with f_col2:
        grupos = sorted(df_base['Grupo de compras'].dropna().unique())
        grupo_sel = st.multiselect("Grupo de Compras", grupos)

    with f_col3:
        proveedores = sorted(df_base['Proveedor/Centro suministrador'].dropna().astype(str).unique())
        prov_sel = st.multiselect("Proveedor", proveedores)

    # Filtrar datos
    df = df_base.copy()
    if centro_sel:
        df = df[df['Centro'].isin(centro_sel)]
    if grupo_sel:
        df = df[df['Grupo de compras'].isin(grupo_sel)]
    if prov_sel:
        df = df[df['Proveedor/Centro suministrador'].astype(str).isin(prov_sel)]

    # KPIs
    total_lineas = len(df)
    pct_on_time = (df['On_Time'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_in_full = (df['In_Full'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_otif = (df['OTIF'].sum() / total_lineas * 100) if total_lineas > 0 else 0

    st.divider()

    st.markdown("**📊 Indicadores de Cumplimiento**")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Líneas Activas", f"{total_lineas:,}")
    k2.metric("⏱️ On Time", f"{pct_on_time:.1f}%")
    k3.metric("📦 In Full", f"{pct_in_full:.1f}%")
    k4.metric("⭐ OTIF Global", f"{pct_otif:.1f}%")

    st.divider()

    # Tabla de Detalle
    st.markdown("**📋 Detalle de Posiciones**")
    cols_mostrar = [
        'Documento compras', 'Posición', 'Centro', 'Proveedor/Centro suministrador',
        'Texto breve', 'Cantidad de pedido', 'Por entregar (cantidad)', 
        'Fecha_Estadistica', 'Fecha_Ingreso_SAP', 'On_Time', 'In_Full', 'OTIF'
    ]
    df_mostrar = df[[c for c in cols_mostrar if c in df.columns]]
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

    # ==========================================
    # DESCARGA LIGERA Y FORMATEADA (BOM UTF-8 y Semicolon)
    # ==========================================
    csv_bytes = df_mostrar.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

    st.download_button(
        label="📥 Descargar Reporte Completo (Formato Excel/CSV)",
        data=csv_bytes,
        file_name="Reporte_OTIF_SAP.csv",
        mime="text/csv"
    )
else:
    st.info("👈 Sube los archivos ME2M y ME80FN para desplegar las métricas.")
