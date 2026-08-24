import pandas as pd
import streamlit as st
import io
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")

# ==========================================
# INYECCIÓN DE CSS (Adaptable a Claro/Oscuro)
# ==========================================
st.markdown("""
<style>
    /* Estilos específicos para la tabla resumen que se adaptan al tema activo */
    table.custom-summary-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        margin-bottom: 20px;
        font-size: 14px;
    }
    table.custom-summary-table thead th {
        background-color: #3b4852 !important; /* Gris oscuro para el header (se ve bien en ambos modos) */
        color: #ffffff !important;
        text-align: right;
        padding: 10px 12px;
        border: none;
        font-weight: 600;
    }
    table.custom-summary-table thead th:first-child {
        text-align: left; /* La primera columna alineada a la izquierda */
    }
    table.custom-summary-table tbody td {
        padding: 10px 12px;
        text-align: right;
        border-bottom: 1px solid rgba(128, 128, 128, 0.2); /* Borde sutil semi-transparente */
        /* Al no definir background-color ni color, hereda los del tema (claro u oscuro) */
    }
    table.custom-summary-table tbody td:first-child {
        text-align: left;
    }
    /* Estilo para la última fila (TOTAL) */
    table.custom-summary-table tbody tr:last-child td {
        background-color: #3b4852 !important;
        color: #ffffff !important;
        font-weight: bold;
        border-top: 3px solid #d93836 !important; /* Línea roja */
        border-bottom: none;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# SISTEMA DE LOGIN (Contraseña)
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔒 Acceso Restringido")
    st.markdown("Por favor, ingresa la contraseña para acceder al Dashboard.")
    
    clave = st.text_input("Contraseña de acceso", type="password")
    
    if st.button("Ingresar"):
        if clave == "ENAEX2026":
            st.session_state["autenticado"] = True
            st.rerun()
        elif clave != "":
            st.error("❌ Contraseña incorrecta. Intenta nuevamente.")
            
    st.stop()


# ==========================================
# SI YA INICIÓ SESIÓN, CONTINÚA EL DASHBOARD
# ==========================================
st.title("🎯 Dashboard OTIF (On Time, In Full)")
st.markdown("Medición del nivel de servicio de proveedores cruzando archivos de SAP y mapeos locales.")

# ==========================================
# FUNCIONES EN CACHÉ (Optimizadas)
# ==========================================
@st.cache_data(show_spinner="Procesando cruce de datos SAP y Mapeos...")
def procesar_otif(file_me2m, file_me80fn, file_centro, file_grupo):
    # 1. Carga de los archivos SAP
    df_me2m = pd.read_excel(file_me2m, engine="openpyxl")
    cols_me80fn = ['Documento compras', 'Posición', 'Fe.contabilización']
    df_me80fn = pd.read_excel(file_me80fn, engine="openpyxl", usecols=lambda c: c in cols_me80fn)

    # Descartar posiciones anuladas en SAP
    if 'Indicador de borrado' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Indicador de borrado'].isna()].copy()

    # 2. Carga de Archivos de Mapeo Locales
    try:
        df_centro = pd.read_excel(file_centro, engine="openpyxl")
        df_centro['Título'] = df_centro['Título'].astype(str)
        df_me2m['Centro'] = df_me2m['Centro'].astype(str)
        df_me2m = pd.merge(df_me2m, df_centro[['Título', 'Nombre Centro', 'Nombre Centro 2']], 
                           left_on='Centro', right_on='Título', how='left')
    except Exception:
        df_me2m['Nombre Centro'] = df_me2m['Centro']
        df_me2m['Nombre Centro 2'] = df_me2m['Centro']
        
    try:
        df_grupo = pd.read_excel(file_grupo, engine="openpyxl")
        df_grupo['Grupo de Compras'] = df_grupo['Grupo de Compras'].astype(str)
        df_me2m['Grupo de compras'] = df_me2m['Grupo de compras'].astype(str)
        df_me2m = pd.merge(df_me2m, df_grupo[['Grupo de Compras', 'Responsable.title']], 
                           left_on='Grupo de compras', right_on='Grupo de Compras', how='left')
        df_me2m.rename(columns={'Responsable.title': 'Comprador'}, inplace=True)
    except Exception:
        df_me2m['Comprador'] = df_me2m['Grupo de compras']

    # Llenar vacíos 
    df_me2m['Nombre Centro'] = df_me2m['Nombre Centro'].fillna(df_me2m['Centro'])
    df_me2m['Nombre Centro 2'] = df_me2m['Nombre Centro 2'].fillna(df_me2m['Centro'])
    df_me2m['Comprador'] = df_me2m['Comprador'].fillna(df_me2m['Grupo de compras'])

    # ==========================================
    # NUEVA REGLA: AGRUPACIÓN VISTA DE CENTRO
    # ==========================================
    def agrupar_centro_logistico(nombre):
        n_upper = str(nombre).upper()
        if 'PRILLEX' in n_upper:
            return 'Prillex'
        elif 'RIO LOA' in n_upper or 'RÍO LOA' in n_upper:
            return 'Rio Loa'
        elif 'TEATINOS' in n_upper:
            return 'Teatinos'
        else:
            return 'Plantas de servicio'
            
    df_me2m['Nombre Centro 2'] = df_me2m['Nombre Centro 2'].apply(agrupar_centro_logistico)

    # 3. Última fecha de recepción real en ME80FN
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Fecha_Ingreso_SAP=('Fe.contabilización', 'max')
    ).reset_index()

    # 4. Cruzar ME2M con ME80FN
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

    # 5. Formato de fechas y cálculo de Semana (Formato visual dd/mm/aa)
    df_otif['Fecha_Estadistica'] = pd.to_datetime(df_otif['Fecha entrega estad.'], errors='coerce').dt.date
    df_otif['Fecha_Ingreso_SAP'] = pd.to_datetime(df_otif['Fecha_Ingreso_SAP'], errors='coerce').dt.date

    def formatear_semana(fecha_val):
        if pd.isna(fecha_val):
            return 'Sin Fecha'
        d = pd.to_datetime(fecha_val)
        lunes = d - pd.Timedelta(days=d.weekday())
        sem = d.isocalendar()[1]
        return f"Sem {sem:02d} - {lunes.strftime('%d/%m/%y')}"
        
    df_otif['Semana/Año'] = df_otif['Fecha_Estadistica'].apply(formatear_semana)

    # 6. Cálculo de Reglas OTIF
    df_otif['In_Full'] = df_otif['Por entregar (cantidad)'] == 0
    df_otif['On_Time'] = (
        df_otif['Fecha_Ingreso_SAP'].notna() & 
        df_otif['Fecha_Estadistica'].notna() & 
        (df_otif['Fecha_Ingreso_SAP'] <= df_otif['Fecha_Estadistica'])
    )
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    # 7. Mapeo de estados visuales
    df_otif['Estado On Time'] = df_otif.apply(
        lambda r: '🔵 A Tiempo' if r['On_Time'] 
        else ('⏳ Pendiente' if pd.isna(r['Fecha_Ingreso_SAP']) else '🔴 Atrasado'), 
        axis=1
    )
    df_otif['Estado In Full'] = df_otif['In_Full'].map({True: '🔵 Completo', False: '🔴 Incompleto'})
    df_otif['Estado OTIF'] = df_otif['OTIF'].map({True: '🔵 Cumple OTIF', False: '🔴 No Cumple'})

    return df_otif

# Función genérica para crear las tablas resumen con totales y formato numérico
def generar_tabla_resumen(df_filtrado, col_agrupacion, nombre_columna):
    if df_filtrado.empty:
        return pd.DataFrame(columns=[nombre_columna, 'Pos. OC', 'OTIF'])
        
    res = df_filtrado.groupby(col_agrupacion).agg(
        Pos_OC=('OTIF', 'count'),
        OTIF_pct=('OTIF', 'mean')
    ).reset_index()
    
    res.rename(columns={col_agrupacion: nombre_columna, 'Pos_OC': 'Pos. OC'}, inplace=True)
    
    # Calcular Totales
    total_pos = res['Pos. OC'].sum()
    total_pct = df_filtrado['OTIF'].mean()
    
    total_row = pd.DataFrame({
        nombre_columna: ['TOTAL'],
        'Pos. OC': [total_pos],
        'OTIF_pct': [total_pct]
    })
    
    res = pd.concat([res, total_row], ignore_index=True)
    
    # Darle formato de porcentaje a la columna OTIF (ej: "87%")
    res['OTIF'] = (res['OTIF_pct'] * 100).fillna(0).round(0).astype(int).astype(str) + "%"
    res = res.drop(columns=['OTIF_pct'])
    
    # Ordenar excluyendo la fila "TOTAL"
    res_body = res.iloc[:-1].sort_values(by='Pos. OC', ascending=False)
    res_final = pd.concat([res_body, res.iloc[[-1]]])
    
    # Aplicar formato de miles a la columna de Posiciones (ej: 10,999)
    res_final['Pos. OC'] = res_final['Pos. OC'].apply(lambda x: f"{x:,}")
    
    return res_final

@st.cache_data(show_spinner="Preparando Excel Resumen Multitabla...")
def generar_excel_resumen(df_detalle, df_t1, df_t2, df_t3):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_detalle.to_excel(writer, index=False, sheet_name='Detalle Posiciones')
        ws0 = writer.sheets['Detalle Posiciones']
        ws0.auto_filter.ref = ws0.dimensions
        for i, col_name in enumerate(df_detalle.columns, 1):
            col_letter = get_column_letter(i)
            if col_name in ['Proveedor/Centro suministrador', 'Texto breve', 'Comprador']:
                ws0.column_dimensions[col_letter].width = 40
            else:
                ws0.column_dimensions[col_letter].width = 18

        def format_summary_sheet(sheet_name, df_res):
            df_res.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 15

        format_summary_sheet('Resumen Comprador', df_t1)
        format_summary_sheet('Resumen Planta Macro', df_t2)
        format_summary_sheet('Resumen Detalle Centro', df_t3)
                
    return output.getvalue()

# Función para renderizar el DataFrame como la tabla HTML estilizada
def renderizar_tabla_html(df):
    html = df.to_html(index=False, classes="custom-summary-table")
    st.markdown(html, unsafe_allow_html=True)


# ==========================================
# BARRA LATERAL (Carga)
# ==========================================
with st.sidebar:
    st.header("📂 Carga de Datos")
    archivo_me2m = st.file_uploader("1. Sube ME2M (.xlsx)", type=["xlsx"])
    archivo_me80fn = st.file_uploader("2. Sube ME80FN (.xlsx)", type=["xlsx"])
    archivo_grupo = st.file_uploader("3. Sube Responsable Grupo (.xlsx)", type=["xlsx"])
    archivo_centro = st.file_uploader("4. Sube Centro Sociedad (.xlsx)", type=["xlsx"])
    
    # Botón de cierre
    st.divider()
    if st.button("Cerrar Sesión"):
        st.session_state["autenticado"] = False
        st.rerun()

# ==========================================
# CUERPO PRINCIPAL
# ==========================================
if archivo_me2m and archivo_me80fn and archivo_grupo and archivo_centro:
    df_base = procesar_otif(archivo_me2m, archivo_me80fn, archivo_centro, archivo_grupo)

    # Filtro de semana en el sidebar 
    with st.sidebar:
        st.divider()
        st.markdown("**📅 Filtro de Tiempo**")
        semanas_disp = sorted([s for s in df_base['Semana/Año'].unique() if s != 'Sin Fecha'])
        if 'Sin Fecha' in df_base['Semana/Año'].unique():
            semanas_disp.append('Sin Fecha')
        semana_sel = st.multiselect("Semana / Año", semanas_disp)

    st.markdown("**🔍 Filtros de Análisis**")
    
    # FILA 1: Entidades (Centro, Grupo, Comprador, Proveedor)
    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1:
        centros = sorted(df_base['Nombre Centro 2'].dropna().unique())
        centro_sel = st.multiselect("Centro Logístico", centros)
    with f_col2:
        grupos = sorted(df_base['Grupo de compras'].dropna().unique())
        grupo_sel = st.multiselect("Grupo de Compras", grupos)
    with f_col3:
        nombres_permitidos = ['Consuelo', 'Sofia', 'Sofía', 'Felipe', 'Constanza']
        compradores = sorted([
            c for c in df_base['Comprador'].dropna().astype(str).unique() 
            if any(n in c for n in nombres_permitidos)
        ])
        comprador_sel = st.multiselect("Comprador", compradores)
    with f_col4:
        proveedores = sorted(df_base['Proveedor/Centro suministrador'].dropna().astype(str).unique())
        prov_sel = st.multiselect("Proveedor", proveedores)

    # FILA 2: Estados OTIF
    f_col5, f_col6, f_col7 = st.columns(3)
    with f_col5:
        estados_on_time = sorted(df_base['Estado On Time'].dropna().unique())
        on_time_sel = st.multiselect("Estado On Time", estados_on_time)
    with f_col6:
        estados_in_full = sorted(df_base['Estado In Full'].dropna().unique())
        in_full_sel = st.multiselect("Estado In Full", estados_in_full)
    with f_col7:
        estados_otif = sorted(df_base['Estado OTIF'].dropna().unique())
        otif_sel = st.multiselect("Estado OTIF", estados_otif)

    # APLICAR TODOS LOS FILTROS
    df = df_base.copy()
    if centro_sel:
        df = df[df['Nombre Centro 2'].isin(centro_sel)]
    if grupo_sel:
        df = df[df['Grupo de compras'].isin(grupo_sel)]
    if comprador_sel:
        df = df[df['Comprador'].isin(comprador_sel)]
    if prov_sel:
        df = df[df['Proveedor/Centro suministrador'].astype(str).isin(prov_sel)]
    if on_time_sel:
        df = df[df['Estado On Time'].isin(on_time_sel)]
    if in_full_sel:
        df = df[df['Estado In Full'].isin(in_full_sel)]
    if otif_sel:
        df = df[df['Estado OTIF'].isin(otif_sel)]
    if semana_sel:
        df = df[df['Semana/Año'].isin(semana_sel)]

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

    # ==========================================
    # TABLAS RESUMEN (Nuevo Layout)
    # ==========================================
    
    # 1. Preparar datos de Comprador
    compradores_obj = [
        'Consuelo Valenzuela Fuenzalida', 
        'Sofia Oporto Oporto', 
        'Felipe Martínez Ulloa', 
        'Constanza Caruz Ruiz'
    ]
    df_comp = df[df['Comprador'].isin(compradores_obj)]
    
    df_t1 = generar_tabla_resumen(df_comp, 'Comprador', 'Comprador (Grupo de compras)')

    # 2. Preparar datos de Plantas
    df_t2 = generar_tabla_resumen(df, 'Nombre Centro 2', 'Centro')
    df_t3 = generar_tabla_resumen(df, 'Nombre Centro', 'Centro (Macro)')

    # 3. Dibujar Layout renderizando a HTML
    st.markdown("### Por comprador")
    renderizar_tabla_html(df_t1)

    st.write("") # Espacio en blanco

    col_izq, col_der = st.columns(2)
    with col_izq:
        st.markdown("### Por centro logístico")
        st.caption("Vista fija — el total calza con la vista por comprador.")
        renderizar_tabla_html(df_t2)
        
    with col_der:
        st.markdown("### Detalle por centro")
        st.caption("Centros activos según los filtros aplicados.")
        renderizar_tabla_html(df_t3)

    st.divider()

    # ==========================================
    # TABLA DE DETALLE
    # ==========================================
    st.markdown("**📋 Detalle de Posiciones y Trazabilidad**")
    cols_mostrar = [
        'Documento compras', 'Posición', 'Semana/Año', 'Centro', 'Nombre Centro', 
        'Comprador', 'Proveedor/Centro suministrador', 'Texto breve', 'Cantidad de pedido', 
        'Por entregar (cantidad)', 'Fecha_Estadistica', 'Fecha_Ingreso_SAP', 
        'Estado On Time', 'Estado In Full', 'Estado OTIF'
    ]
    df_mostrar = df[[c for c in cols_mostrar if c in df.columns]]

    config_columnas = {
        "Proveedor/Centro suministrador": st.column_config.TextColumn("Proveedor", width="large"),
        "Texto breve": st.column_config.TextColumn("Texto breve", width="large"),
        "Comprador": st.column_config.TextColumn("Comprador", width="medium"),
    }
    # La tabla general la mantenemos con st.dataframe para no perder el scroll y filtrado nativo
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, column_config=config_columnas)

    # ==========================================
    # DESCARGA DEL SÚPER EXCEL
    # ==========================================
    st.divider()
    st.markdown("### 📥 Descarga de Reportes")
    
    excel_bytes = generar_excel_resumen(df_mostrar, df_t1, df_t2, df_t3)

    st.download_button(
        label="📊 Descargar Reporte de Trazabilidad y Resumen (.xlsx)",
        data=excel_bytes,
        file_name="Trazabilidad_OTIF_Semanal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.info("👈 Sube los 4 archivos en la barra lateral para desplegar las métricas correspondientes.")
