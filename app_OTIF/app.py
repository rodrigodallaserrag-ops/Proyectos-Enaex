import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")
st.title("🎯 Dashboard OTIF - Nivel de Servicio")

# Función de procesamiento con limpieza de registros borrados
@st.cache_data(show_spinner="Procesando datos de SAP...")
def procesar_otif(file_me2m, file_me80fn):
    df_me2m = pd.read_excel(file_me2m, engine="openpyxl")
    df_me80fn = pd.read_excel(file_me80fn, engine="openpyxl")

    # 1. Filtrar posiciones borradas en SAP (L o S)
    if 'Indicador de borrado' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Indicador de borrado'].isna()].copy()

    # 2. Agrupar recepciones reales en ME80FN
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Cantidad_Recibida=('Cantidad', 'sum'),
        Fecha_Real_Entrega=('Fe.contabilización', 'max')
    ).reset_index()

    # 3. Cruzar ME2M con ME80FN
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

    # 4. Normalizar datos
    df_otif['Cantidad_Recibida'] = df_otif['Cantidad_Recibida'].fillna(0)
    df_otif['Fecha de entrega'] = pd.to_datetime(df_otif['Fecha de entrega']).dt.date
    df_otif['Fecha_Real_Entrega'] = pd.to_datetime(df_otif['Fecha_Real_Entrega']).dt.date

    # 5. Lógica de Indicadores
    df_otif['In_Full'] = df_otif['Cantidad_Recibida'] >= df_otif['Cantidad de pedido']
    df_otif['On_Time'] = df_otif.apply(
        lambda x: True if pd.notna(x['Fecha_Real_Entrega']) and (x['Fecha_Real_Entrega'] <= x['Fecha de entrega']) else False, 
        axis=1
    )
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    # Etiqueta de estado simplificada
    def definir_estado(row):
        if row['OTIF']:
            return "Cumple OTIF"
        elif not row['In_Full'] and row['On_Time']:
            return "Solo A Tiempo (Incompleto)"
        elif row['In_Full'] and not row['On_Time']:
            return "Solo Completo (Atrasado)"
        else:
            return "Incumple Ambos"

    df_otif['Estado_OTIF'] = df_otif.apply(definir_estado, axis=1)

    return df_otif

# Barra Lateral
with st.sidebar:
    st.header("📂 Carga de Archivos")
    archivo_me2m = st.file_uploader("1. Reporte ME2M (.xlsx)", type=["xlsx"])
    archivo_me80fn = st.file_uploader("2. Reporte ME80FN (.xlsx)", type=["xlsx"])

if archivo_me2m and archivo_me80fn:
    df_base = procesar_otif(archivo_me2m, archivo_me80fn)

    # Filtros Superiores
    st.markdown("### 🔍 Filtros de Análisis")
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

    # Filtrado dinámico
    df = df_base.copy()
    if centro_sel:
        df = df[df['Centro'].isin(centro_sel)]
    if grupo_sel:
        df = df[df['Grupo de compras'].isin(grupo_sel)]
    if prov_sel:
        df = df[df['Proveedor/Centro suministrador'].astype(str).isin(prov_sel)]

    # Cálculo de Métricas
    total_lineas = len(df)
    pct_on_time = (df['On_Time'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_in_full = (df['In_Full'].sum() / total_lineas * 100) if total_lineas > 0 else 0
    pct_otif = (df['OTIF'].sum() / total_lineas * 100) if total_lineas > 0 else 0

    st.divider()

    # Tarjetas Métricas
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Líneas Activas", f"{total_lineas:,}")
    kpi2.metric("⏱️ % On Time", f"{pct_on_time:.1f}%")
    kpi3.metric("📦 % In Full", f"{pct_in_full:.1f}%")
    kpi4.metric("⭐ Nivel OTIF", f"{pct_otif:.1f}%")

    st.divider()

    # Visualizaciones Gráficas
    g_col1, g_col2 = st.columns([1, 2])

    with g_col1:
        st.markdown("#### Distribución de Cumplimiento")
        conteo_estados = df['Estado_OTIF'].value_counts()
        st.bar_chart(conteo_estados, color="#0068c9")

    with g_col2:
        st.markdown("#### Top 10 Proveedores con Menor OTIF (%)")
        if total_lineas > 0:
            prov_otif = df.groupby('Proveedor/Centro suministrador').agg(
                Total=('OTIF', 'count'),
                OTIF_OK=('OTIF', 'sum')
            )
            prov_otif = prov_otif[prov_otif['Total'] >= 5]  # Mínimo 5 pedidos
            prov_otif['% OTIF'] = (prov_otif['OTIF_OK'] / prov_otif['Total']) * 100
            top_peores = prov_otif.sort_values(by='% OTIF').head(10)['% OTIF']
            st.bar_chart(top_peores, color="#ff2b2b")

    st.divider()

    # Tabla de Detalle y Exportación
    st.markdown("#### 📋 Detalle de Posiciones")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Botón Descargar Resultados
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Reporte OTIF en CSV",
        data=csv_data,
        file_name="Reporte_OTIF_Filtrado.csv",
        mime="text/csv"
    )
else:
    st.info("💡 Por favor, sube los archivos en el menú lateral para actualizar el panel.")
