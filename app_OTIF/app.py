import pandas as pd
import streamlit as st
import io
import requests
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard OTIF", page_icon="🎯", layout="wide")

# ==========================================
# CONFIGURACIÓN DE RUTAS Y ENLACES ONEDRIVE
# ==========================================
ONEDRIVE_URLS = {
    "me2m": "https://empresassk-my.sharepoint.com/:x:/g/personal/rodrigo_dallaserra_enaex_com/IQCsqrxqNHT5QbsHdXTEX6j3ATbS-oKte1km1xtAm4xtMrY?e=1OLXgF&download=1",
    "me80fn": "https://empresassk-my.sharepoint.com/:x:/g/personal/rodrigo_dallaserra_enaex_com/IQD0KZXh49t6QJiNk-8Db8GLAQEodvYZizKumm57X2B8PBo?e=pxeBjd&download=1",
    "grupo": "https://empresassk-my.sharepoint.com/:x:/g/personal/rodrigo_dallaserra_enaex_com/IQCWw06SLb9DT7t5WVOemA19AT663LxrU6u9e7ZXygjKGDE?e=6uKY02&download=1",
    "centro": "https://empresassk-my.sharepoint.com/:x:/g/personal/rodrigo_dallaserra_enaex_com/IQB6Ny2KoGQURp6nWehTspgGAfwNUsxEcW4xzfszzAfUtAM?e=Y0aQsB&download=1"
}

LOCAL_PATHS = {
    "me2m": "data/ME2M.xlsx",
    "me80fn": "data/ME80FN.xlsx",
    "grupo": "data/Grupo.xlsx",
    "centro": "data/Centro.xlsx"
}

# ==========================================
# INYECCIÓN DE CSS (Adaptable a Claro/Oscuro con Scroll)
# ==========================================
st.markdown("""
<style>
    .table-container {
        max-height: 400px;
        overflow-y: auto;
        margin-bottom: 20px;
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 5px;
    }
    
    table.custom-summary-table {
        width: 100%;
        border-collapse: collapse;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 14px;
        margin-bottom: 0px; 
    }
    table.custom-summary-table thead th {
        background-color: #3b4852 !important;
        color: #ffffff !important;
        text-align: right;
        padding: 10px 12px;
        border: none;
        font-weight: 600;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    table.custom-summary-table thead th:first-child {
        text-align: left;
    }
    table.custom-summary-table tbody td {
        padding: 10px 12px;
        text-align: right;
        border-bottom: 1px solid rgba(128, 128, 128, 0.2);
    }
    table.custom-summary-table tbody td:first-child {
        text-align: left;
    }
    table.custom-summary-table tbody tr:last-child td {
        background-color: #3b4852 !important;
        color: #ffffff !important;
        font-weight: bold;
        border-top: 3px solid #d93836 !important;
        border-bottom: none;
    }
    
    .table-container::-webkit-scrollbar {
        width: 8px;
    }
    .table-container::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.1); 
    }
    .table-container::-webkit-scrollbar-thumb {
        background: rgba(128, 128, 128, 0.5); 
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# SISTEMA DE LOGIN
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
# DASHBOARD PRINCIPAL
# ==========================================
st.title("🎯 Dashboard OTIF (On Time, In Full)")
st.markdown("Medición del nivel de servicio de proveedores cruzando archivos de SAP y mapeos locales.")

# ==========================================
# FUNCIONES EN CACHÉ Y DESCARGA
# ==========================================
def obtener_buffer_archivo(origen):
    if isinstance(origen, str) and origen.startswith("http"):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
        }
        response = requests.get(origen, headers=headers)
        response.raise_for_status()
        return io.BytesIO(response.content)
    return origen

@st.cache_data(show_spinner="Cargando y procesando datos...")
def procesar_otif(file_me2m, file_me80fn, file_centro, file_grupo):
    buffer_me2m = obtener_buffer_archivo(file_me2m)
    buffer_me80fn = obtener_buffer_archivo(file_me80fn)
    buffer_centro = obtener_buffer_archivo(file_centro)
    buffer_grupo = obtener_buffer_archivo(file_grupo)

    # 1. LECTURA Y FILTRO DE ME80FN (Agregada Clase de movimiento 101)
    df_me2m = pd.read_excel(buffer_me2m, engine="openpyxl")
    cols_me80fn = ['Documento compras', 'Posición', 'Fe.contabilización', 'Clase de movimiento']
    df_me80fn = pd.read_excel(buffer_me80fn, engine="openpyxl", usecols=lambda c: c in cols_me80fn)

    if 'Clase de movimiento' in df_me80fn.columns:
        df_me80fn = df_me80fn[df_me80fn['Clase de movimiento'].astype(str) == '101']

    # 2. FILTROS EN ME2M (Reglas de negocio)
    if 'Indicador de borrado' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Indicador de borrado'].isna()].copy()
        
    if 'Ind.liberación' in df_me2m.columns:
        df_me2m = df_me2m[df_me2m['Ind.liberación'] == 'B'].copy()

    # (Opcional) Filtrar años 2018-2021. Descomentar si usas la 'Fecha del documento'
    # if 'Fecha del documento' in df_me2m.columns:
    #     df_me2m['Fecha del documento'] = pd.to_datetime(df_me2m['Fecha del documento'], errors='coerce')
    #     df_me2m = df_me2m[df_me2m['Fecha del documento'].dt.year > 2021].copy()

    # (Opcional) Excluir OC Masiva. Descomentar y cambiar 'Tipo Documento' por tu columna real
    # if 'Clase de documento' in df_me2m.columns:
    #     df_me2m = df_me2m[df_me2m['Clase de documento'] != 'ZMAS'].copy()

    # 3. MERGE DE CENTROS Y GRUPOS
    try:
        df_centro = pd.read_excel(buffer_centro, engine="openpyxl")
        df_centro['Título'] = df_centro['Título'].astype(str)
        df_me2m['Centro'] = df_me2m['Centro'].astype(str)
        df_me2m = pd.merge(df_me2m, df_centro[['Título', 'Nombre Centro', 'Nombre Centro 2']], 
                           left_on='Centro', right_on='Título', how='left')
    except Exception:
        df_me2m['Nombre Centro'] = df_me2m['Centro']
        df_me2m['Nombre Centro 2'] = df_me2m['Centro']
        
    try:
        df_grupo = pd.read_excel(buffer_grupo, engine="openpyxl")
        df_grupo['Grupo de Compras'] = df_grupo['Grupo de Compras'].astype(str)
        df_me2m['Grupo de compras'] = df_me2m['Grupo de compras'].astype(str)
        df_me2m = pd.merge(df_me2m, df_grupo[['Grupo de Compras', 'Responsable.title']], 
                           left_on='Grupo de compras', right_on='Grupo de Compras', how='left')
        df_me2m.rename(columns={'Responsable.title': 'Comprador'}, inplace=True)
    except Exception:
        df_me2m['Comprador'] = df_me2m['Grupo de compras']

    df_me2m['Nombre Centro'] = df_me2m['Nombre Centro'].fillna(df_me2m['Centro'])
    df_me2m['Nombre Centro 2'] = df_me2m['Nombre Centro 2'].fillna(df_me2m['Centro'])
    df_me2m['Comprador'] = df_me2m['Comprador'].fillna(df_me2m['Grupo de compras'])

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

    # 4. AGRUPACIÓN DE RECEPCIONES (Deduplicación)
    df_recepciones = df_me80fn.groupby(['Documento compras', 'Posición']).agg(
        Fecha_Ingreso_SAP=('Fe.contabilización', 'max')
    ).reset_index()

    # 5. MERGE FINAL Y CÁLCULOS
    df_otif = pd.merge(df_me2m, df_recepciones, on=['Documento compras', 'Posición'], how='left')

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

    df_otif['In_Full'] = df_otif['Por entregar (cantidad)'] == 0
    df_otif['On_Time'] = (
        df_otif['Fecha_Ingreso_SAP'].notna() & 
        df_otif['Fecha_Estadistica'].notna() & 
        (df_otif['Fecha_Ingreso_SAP'] <= df_otif['Fecha_Estadistica'])
    )
    df_otif['OTIF'] = df_otif['In_Full'] & df_otif['On_Time']

    df_otif['Estado On Time'] = df_otif.apply(
        lambda r: '🔵 A Tiempo' if r['On_Time'] 
        else ('⏳ Pendiente' if pd.isna(r['Fecha_Ingreso_SAP']) else '🔴 Atrasado'), 
        axis=1
    )
    df_otif['Estado In Full'] = df_otif['In_Full'].map({True: '🔵 Completo', False: '🔴 Incompleto'})
    df_otif['Estado OTIF'] = df_otif['OTIF'].map({True: '🔵 Cumple OTIF', False: '🔴 No Cumple'})

    return df_otif

def generar_tabla_resumen(df_filtrado, col_agrupacion, nombre_columna):
    if df_filtrado.empty:
        return pd.DataFrame(columns=[nombre_columna, 'Pos. OC', 'OTIF'])
        
    res = df_filtrado.groupby(col_agrupacion).agg(
        Pos_OC=('OTIF', 'count'),
        OTIF_pct=('OTIF', 'mean')
    ).reset_index()
    
    res.rename(columns={col_agrupacion: nombre_columna, 'Pos_OC': 'Pos. OC'}, inplace=True)
    
    total_pos = res['Pos. OC'].sum()
    total_pct = df_filtrado['OTIF'].mean()
    
    total_row = pd.DataFrame({
        nombre_columna: ['TOTAL'],
        'Pos. OC': [total_pos],
        'OTIF_pct': [total_pct]
    })
    
    res = pd.concat([res, total_row], ignore_index=True)
    res['OTIF'] = (res['OTIF_pct'] * 100).fillna(0).round(0).astype(int).astype(str) + "%"
    res = res.drop(columns=['OTIF_pct'])
    
    res_body = res.iloc[:-1].sort_values(by='Pos. OC', ascending=False)
    res_final = pd.concat([res_body, res.iloc[[-1]]])
    res_final['Pos. OC'] = res_final['Pos. OC'].apply(lambda x: f"{x:,}")
    
    return res_final

@st.cache_data(show_spinner="Preparando Excel Resumen...")
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

def renderizar_tabla_html(df):
    html_table = df.to_html(index=False, classes="custom-summary-table")
    html_final = f'<div class="table-container">{html_table}</div>'
    st.markdown(html_final, unsafe_allow_html=True)


# ==========================================
# BARRA LATERAL: ORIGEN DE DATOS Y FILTRO TIEMPO
# ==========================================
with st.sidebar:
    st.header("Origen de datos")
    origen_datos = st.radio(
        "Selecciona el origen:",
        ["OneDrive (automático)", "Subir archivos", "Archivos locales (data/)"],
        label_visibility="collapsed"
    )

    file_me2m, file_me80fn, file_grupo, file_centro = None, None, None, None

    if origen_datos == "OneDrive (automático)":
        if st.button("🔄 Forzar recarga desde OneDrive ahora", use_container_width=True):
            st.cache_data.clear()
            st.toast("Caché borrada. Volviendo a descargar desde OneDrive...")
            st.rerun()
            
        file_me2m = ONEDRIVE_URLS["me2m"]
        file_me80fn = ONEDRIVE_URLS["me80fn"]
        file_grupo = ONEDRIVE_URLS["grupo"]
        file_centro = ONEDRIVE_URLS["centro"]

    elif origen_datos == "Subir archivos":
        st.markdown("**📂 Carga manual de archivos**")
        file_me2m = st.file_uploader("1. Sube ME2M (.xlsx)", type=["xlsx"])
        file_me80fn = st.file_uploader("2. Sube ME80FN (.xlsx)", type=["xlsx"])
        file_grupo = st.file_uploader("3. Sube Responsable Grupo (.xlsx)", type=["xlsx"])
        file_centro = st.file_uploader("4. Sube Centro Sociedad (.xlsx)", type=["xlsx"])

    elif origen_datos == "Archivos locales (data/)":
        file_me2m = LOCAL_PATHS["me2m"]
        file_me80fn = LOCAL_PATHS["me80fn"]
        file_grupo = LOCAL_PATHS["grupo"]
        file_centro = LOCAL_PATHS["centro"]


# ==========================================
# PROCESAMIENTO Y RENDERIZADO
# ==========================================
listos_para_procesar = False

if origen_datos == "Subir archivos":
    if file_me2m and file_me80fn and file_grupo and file_centro:
        listos_para_procesar = True
    else:
        st.info("👈 Sube los 4 archivos en la barra lateral para desplegar las métricas.")
else:
    listos_para_procesar = True

if listos_para_procesar:
    try:
        df_base = procesar_otif(file_me2m, file_me80fn, file_centro, file_grupo)
    except Exception as e:
        st.error(f"❌ Error al cargar los datos: {e}")
        st.stop()

    # Devolvemos el filtro de tiempo a la barra lateral
    with st.sidebar:
        st.divider()
        st.markdown("**📅 Filtro de Tiempo**")
        semanas_disp = sorted([s for s in df_base['Semana/Año'].unique() if s != 'Sin Fecha'])
        if 'Sin Fecha' in df_base['Semana/Año'].unique():
            semanas_disp.append('Sin Fecha')
        semana_sel = st.multiselect("Semana / Año", semanas_disp, placeholder="Vacío = Todas las semanas")
        
        st.divider()
        if st.button("Cerrar Sesión", use_container_width=True):
            st.session_state["autenticado"] = False
            st.rerun()

    # Pre-calcular listas para los filtros
    centros = sorted(df_base['Nombre Centro 2'].dropna().unique())
    grupos = sorted(df_base['Grupo de compras'].dropna().unique())
    nombres_permitidos = ['Consuelo', 'Sofia', 'Sofía', 'Felipe', 'Constanza']
    compradores = sorted([
        c for c in df_base['Comprador'].dropna().astype(str).unique() 
        if any(n in c for n in nombres_permitidos)
    ])
    proveedores = sorted(df_base['Proveedor/Centro suministrador'].dropna().astype(str).unique())
    estados_on_time = sorted(df_base['Estado On Time'].dropna().unique())
    estados_in_full = sorted(df_base['Estado In Full'].dropna().unique())
    estados_otif = sorted(df_base['Estado OTIF'].dropna().unique())

    st.info("💡 **Solución al 'congelamiento':** Los filtros complejos ahora están dentro de un **Panel**. Puedes agregar o quitar las etiquetas que necesites con total fluidez. Solo debes dar clic en el botón **Aplicar Filtros** para ver los resultados.")

    # ==========================================
    # NUEVO PANEL DE FILTROS (ST.FORM)
    # ==========================================
    with st.form("panel_filtros"):
        st.markdown("**🔍 Panel de Filtros**")
        
        # Fila 1 de filtros
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)
        with f_col1:
            centro_sel = st.multiselect("Centro Logístico", centros, placeholder="Vacío = Todos")
        with f_col2:
            grupo_sel = st.multiselect("Grupo de Compras", grupos, placeholder="Vacío = Todos")
        with f_col3:
            comprador_sel = st.multiselect("Comprador", compradores, placeholder="Vacío = Todos")
        with f_col4:
            prov_sel = st.multiselect("Proveedor", proveedores, placeholder="Vacío = Todos")

        # Fila 2 de filtros
        f_col5, f_col6, f_col7 = st.columns(3)
        with f_col5:
            on_time_sel = st.multiselect("Estado On Time", estados_on_time, placeholder="Vacío = Todos")
        with f_col6:
            in_full_sel = st.multiselect("Estado In Full", estados_in_full, placeholder="Vacío = Todos")
        with f_col7:
            otif_sel = st.multiselect("Estado OTIF", estados_otif, placeholder="Vacío = Todos")

        btn_aplicar = st.form_submit_button("🔄 Aplicar Filtros", type="primary", use_container_width=True)

    df = df_base.copy()
    
    # Aplicar lógica de filtrado
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

    compradores_obj = [
        'Consuelo Valenzuela Fuenzalida', 
        'Sofia Oporto Oporto', 
        'Felipe Martínez Ulloa', 
        'Constanza Caruz Ruiz'
    ]
    df_comp = df[df['Comprador'].isin(compradores_obj)]
    
    df_t1 = generar_tabla_resumen(df_comp, 'Comprador', 'Comprador (Grupo de compras)')
    df_t2 = generar_tabla_resumen(df, 'Nombre Centro 2', 'Centro')
    df_t3 = generar_tabla_resumen(df, 'Nombre Centro', 'Centro (Macro)')

    st.markdown("### Por comprador")
    renderizar_tabla_html(df_t1)

    st.write("")

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
    st.dataframe(df_mostrar, use_container_width=True, hide_index=True, column_config=config_columnas)

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
