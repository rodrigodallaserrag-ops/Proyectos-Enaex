import html
import io

import pandas as pd
import requests
import streamlit as st
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw


# ==========================================
# ÍCONO DE LA PESTAÑA (dibujado, sin emojis)
# ==========================================
def crear_icono_pagina():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((4, 4, 60, 60), outline=(217, 56, 54, 255), width=6)
    d.ellipse((18, 18, 46, 46), outline=(59, 72, 82, 255), width=6)
    d.ellipse((27, 27, 37, 37), fill=(217, 56, 54, 255))
    return img


st.set_page_config(page_title="Dashboard OTIF", page_icon=crear_icono_pagina(), layout="wide")

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
# PALETA Y ESTADOS
# ==========================================
COLOR_ROJO = "#d93836"
COLOR_PIZARRA = "#3b4852"
COLOR_AZUL = "#2f7fc1"
COLOR_AMBAR = "#c98a12"
COLOR_VERDE = "#2e8b57"

COLOR_ESTADO = {
    "A Tiempo": COLOR_AZUL,
    "Completo": COLOR_AZUL,
    "Cumple OTIF": COLOR_AZUL,
    "Pendiente": COLOR_AMBAR,
    "Atrasado": COLOR_ROJO,
    "Incompleto": COLOR_ROJO,
    "No Cumple": COLOR_ROJO,
}

# ==========================================
# ÍCONOS SVG (trazos simples, heredan el color del texto)
# ==========================================
ICONOS = {
    "objetivo": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.5" fill="currentColor"/>',
    "candado": '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/><path d="M12 15v2"/>',
    "reloj": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "caja": '<path d="M21 8l-9-5-9 5v8l9 5 9-5z"/><path d="M3 8l9 5 9-5"/><path d="M12 13v8"/>',
    "check_circulo": '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.8 2.8L16 9.5"/>',
    "lista": '<path d="M9 6h11M9 12h11M9 18h11"/><circle cx="4.5" cy="6" r="1" fill="currentColor"/><circle cx="4.5" cy="12" r="1" fill="currentColor"/><circle cx="4.5" cy="18" r="1" fill="currentColor"/>',
    "filtro": '<path d="M3 4h18l-7 8.5V19l-4 2v-8.5z"/>',
    "calendario": '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/>',
    "grafico": '<path d="M4 4v16h16"/><path d="M8 16v-4M12 16V8M16 16v-6"/>',
    "persona": '<circle cx="12" cy="8" r="4"/><path d="M4 21v-1a6 6 0 0 1 6-6h4a6 6 0 0 1 6 6v1"/>',
    "planta": '<path d="M3 21V9l9-5 9 5v12"/><path d="M7 21v-8h10v8"/><path d="M7 17h10"/>',
    "ubicacion": '<path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>',
    "tabla": '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 10v10"/>',
    "descarga": '<path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M5 21h14"/>',
    "base_datos": '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
    "archivo": '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/><path d="M8 14h8M8 18h5"/>',
    "subir": '<path d="M12 21V9"/><path d="M7 14l5-5 5 5"/><path d="M5 3h14"/>',
    "info": '<circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><circle cx="12" cy="8" r="0.6" fill="currentColor"/>',
    "alerta": '<circle cx="12" cy="12" r="9"/><path d="M12 7v6"/><circle cx="12" cy="16.5" r="0.6" fill="currentColor"/>',
    "check": '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    "pendiente": '<circle cx="12" cy="12" r="8"/>',
}


def icono(nombre, size=20, stroke=2):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" '
        f'fill="none" stroke="currentColor" stroke-width="{stroke}" stroke-linecap="round" '
        f'stroke-linejoin="round" aria-hidden="true">{ICONOS[nombre]}</svg>'
    )


# ==========================================
# INYECCIÓN DE CSS (Adaptable a Claro/Oscuro con Scroll)
# ==========================================
st.markdown(f"""
<style>
    /* ---------- Encabezado principal ---------- */
    .app-header {{
        display: flex;
        align-items: center;
        gap: 16px;
        padding-bottom: 18px;
        margin-bottom: 12px;
        border-bottom: 1px solid rgba(128, 128, 128, 0.25);
    }}
    .app-logo {{
        width: 52px;
        height: 52px;
        flex-shrink: 0;
        border-radius: 12px;
        background: {COLOR_ROJO};
        color: #ffffff;
        display: flex;
        align-items: center;
        justify-content: center;
    }}
    .app-titulo {{
        font-size: 1.75rem;
        font-weight: 700;
        line-height: 1.2;
    }}
    .app-sub {{
        font-size: 0.95rem;
        opacity: 0.75;
        margin-top: 4px;
    }}

    /* ---------- Encabezados de sección ---------- */
    .sec-head {{
        display: flex;
        align-items: flex-start;
        gap: 10px;
        margin: 6px 0 12px 0;
    }}
    .sec-icon {{
        color: {COLOR_ROJO};
        display: flex;
        padding-top: 2px;
    }}
    .sec-titulo {{
        font-size: 1.12rem;
        font-weight: 650;
        line-height: 1.3;
    }}
    .sec-desc {{
        font-size: 0.85rem;
        opacity: 0.7;
        margin-top: 2px;
    }}

    /* ---------- Avisos ---------- */
    .aviso {{
        display: flex;
        align-items: flex-start;
        gap: 10px;
        padding: 12px 14px;
        margin: 8px 0 14px 0;
        border-radius: 8px;
        border: 1px solid rgba(128, 128, 128, 0.25);
        border-left: 4px solid var(--c);
        background: rgba(128, 128, 128, 0.05);
        font-size: 0.92rem;
    }}
    .aviso-icon {{
        color: var(--c);
        display: flex;
        padding-top: 1px;
    }}

    /* ---------- Login ---------- */
    .login-head {{
        text-align: center;
        margin: 8vh 0 20px 0;
    }}
    .login-icon {{
        width: 64px;
        height: 64px;
        border-radius: 50%;
        background: {COLOR_PIZARRA};
        color: #ffffff;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 14px;
    }}
    .login-titulo {{
        font-size: 1.5rem;
        font-weight: 700;
    }}
    .login-sub {{
        opacity: 0.75;
        margin-top: 6px;
        font-size: 0.95rem;
    }}

    /* ---------- Lista de archivos por cargar ---------- */
    .carga {{
        max-width: 560px;
        border: 1px solid rgba(128, 128, 128, 0.25);
        border-radius: 10px;
        padding: 18px 20px;
        background: rgba(128, 128, 128, 0.04);
    }}
    .checklist {{
        list-style: none;
        padding: 0;
        margin: 10px 0 0 0;
    }}
    .checklist li {{
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 9px 0;
        margin: 0;
        border-bottom: 1px solid rgba(128, 128, 128, 0.15);
    }}
    .checklist li:last-child {{
        border-bottom: none;
    }}
    .checklist .estado-ok {{
        color: {COLOR_VERDE};
        display: flex;
    }}
    .checklist .estado-pend {{
        opacity: 0.45;
        display: flex;
    }}
    .checklist .nombre {{
        font-weight: 600;
    }}
    .checklist .detalle {{
        font-size: 0.85rem;
        opacity: 0.7;
    }}

    /* ---------- Tarjetas de indicadores ---------- */
    .kpi-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 16px;
        margin-bottom: 8px;
    }}
    @media (max-width: 1000px) {{
        .kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    }}
    @media (max-width: 560px) {{
        .kpi-grid {{ grid-template-columns: 1fr; }}
    }}
    .kpi {{
        border: 1px solid rgba(128, 128, 128, 0.25);
        border-radius: 10px;
        padding: 16px 18px;
        background: rgba(128, 128, 128, 0.04);
    }}
    .kpi.kpi-principal {{
        border-top: 3px solid {COLOR_ROJO};
    }}
    .kpi-head {{
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 0.92rem;
        font-weight: 600;
    }}
    .kpi-icon {{
        width: 34px;
        height: 34px;
        flex-shrink: 0;
        border-radius: 8px;
        background: rgba(128, 128, 128, 0.15);
        display: flex;
        align-items: center;
        justify-content: center;
    }}
    .kpi.kpi-principal .kpi-icon {{
        background: {COLOR_ROJO};
        color: #ffffff;
    }}
    .kpi-valor {{
        font-size: 2rem;
        font-weight: 700;
        line-height: 1.1;
        margin: 12px 0 10px 0;
        font-variant-numeric: tabular-nums;
    }}
    .kpi-barra {{
        height: 6px;
        border-radius: 3px;
        background: rgba(128, 128, 128, 0.2);
        overflow: hidden;
    }}
    .kpi-barra span {{
        display: block;
        height: 100%;
        background: {COLOR_AZUL};
    }}
    .kpi.kpi-principal .kpi-barra span {{
        background: {COLOR_ROJO};
    }}
    .kpi-ayuda {{
        font-size: 0.8rem;
        opacity: 0.7;
        margin-top: 8px;
    }}

    /* ---------- Tablas resumen ---------- */
    .table-container {{
        max-height: 400px;
        overflow-y: auto;
        margin-bottom: 20px;
        border: 1px solid rgba(128, 128, 128, 0.2);
        border-radius: 8px;
    }}
    table.custom-summary-table {{
        width: 100%;
        border-collapse: collapse;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 14px;
        margin-bottom: 0px;
    }}
    table.custom-summary-table thead th {{
        background-color: {COLOR_PIZARRA} !important;
        color: #ffffff !important;
        text-align: right;
        padding: 10px 12px;
        border: none;
        font-weight: 600;
        position: sticky;
        top: 0;
        z-index: 10;
    }}
    table.custom-summary-table thead th:first-child {{
        text-align: left;
    }}
    table.custom-summary-table tbody td {{
        padding: 9px 12px;
        text-align: right;
        border: none;
        border-bottom: 1px solid rgba(128, 128, 128, 0.2);
        font-variant-numeric: tabular-nums;
    }}
    table.custom-summary-table tbody td:first-child {{
        text-align: left;
    }}
    table.custom-summary-table tbody tr:hover td {{
        background-color: rgba(128, 128, 128, 0.08);
    }}
    table.custom-summary-table tbody tr.fila-total td {{
        background-color: {COLOR_PIZARRA} !important;
        color: #ffffff !important;
        font-weight: bold;
        border-top: 3px solid {COLOR_ROJO} !important;
        border-bottom: none;
    }}
    table.custom-summary-table tbody tr.fila-vacia td {{
        text-align: center;
        padding: 22px 12px;
        opacity: 0.7;
    }}
    .otif-cell {{
        display: flex;
        align-items: center;
        justify-content: flex-end;
        gap: 10px;
    }}
    .otif-bar {{
        width: 90px;
        height: 6px;
        border-radius: 3px;
        background: rgba(128, 128, 128, 0.22);
        overflow: hidden;
    }}
    .otif-bar span {{
        display: block;
        height: 100%;
        background: {COLOR_AZUL};
    }}
    tr.fila-total .otif-bar {{
        background: rgba(255, 255, 255, 0.25);
    }}
    tr.fila-total .otif-bar span {{
        background: #ffffff;
    }}
    .otif-val {{
        min-width: 40px;
    }}

    .table-container::-webkit-scrollbar {{
        width: 8px;
    }}
    .table-container::-webkit-scrollbar-track {{
        background: rgba(0, 0, 0, 0.1);
    }}
    .table-container::-webkit-scrollbar-thumb {{
        background: rgba(128, 128, 128, 0.5);
        border-radius: 4px;
    }}
</style>
""", unsafe_allow_html=True)


# ==========================================
# COMPONENTES VISUALES (HTML)
# ==========================================
def encabezado_seccion(nombre_icono, titulo, descripcion=None):
    desc = f'<div class="sec-desc">{descripcion}</div>' if descripcion else ""
    st.markdown(
        f'<div class="sec-head"><span class="sec-icon">{icono(nombre_icono, 22)}</span>'
        f'<div><div class="sec-titulo">{titulo}</div>{desc}</div></div>',
        unsafe_allow_html=True,
    )


def aviso(texto, tipo="info"):
    config = {
        "info": ("info", COLOR_AZUL),
        "error": ("alerta", COLOR_ROJO),
        "ok": ("check_circulo", COLOR_VERDE),
    }
    nombre_icono, color = config.get(tipo, config["info"])
    st.markdown(
        f'<div class="aviso" style="--c:{color}"><span class="aviso-icon">{icono(nombre_icono, 20)}</span>'
        f'<div>{texto}</div></div>',
        unsafe_allow_html=True,
    )


def tarjetas_kpi(tarjetas):
    bloques = []
    for t in tarjetas:
        clase = "kpi kpi-principal" if t.get("principal") else "kpi"
        barra = ""
        if t.get("pct") is not None:
            ancho = max(0.0, min(100.0, t["pct"]))
            barra = f'<div class="kpi-barra"><span style="width:{ancho:.1f}%"></span></div>'
        bloques.append(
            f'<div class="{clase}">'
            f'<div class="kpi-head"><span class="kpi-icon">{icono(t["icono"], 18)}</span>{t["titulo"]}</div>'
            f'<div class="kpi-valor">{t["valor"]}</div>{barra}'
            f'<div class="kpi-ayuda">{t["ayuda"]}</div></div>'
        )
    st.markdown(f'<div class="kpi-grid">{"".join(bloques)}</div>', unsafe_allow_html=True)


# ==========================================
# SISTEMA DE LOGIN
# ==========================================
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    _, col_login, _ = st.columns([1, 1.2, 1])
    with col_login:
        st.markdown(
            f'<div class="login-head"><div class="login-icon">{icono("candado", 30)}</div>'
            f'<div class="login-titulo">Acceso restringido</div>'
            f'<div class="login-sub">Ingresa la contraseña para ver el Dashboard OTIF.</div></div>',
            unsafe_allow_html=True,
        )
        with st.form("form_login"):
            clave = st.text_input("Contraseña de acceso", type="password")
            ingresar = st.form_submit_button("Ingresar", type="primary", use_container_width=True)

        if ingresar:
            if clave == "ENAEX2026":
                st.session_state["autenticado"] = True
                st.rerun()
            elif clave != "":
                aviso("La contraseña no es correcta. Revisa mayúsculas y vuelve a intentarlo.", "error")

    st.stop()


# ==========================================
# DASHBOARD PRINCIPAL
# ==========================================
st.markdown(
    f'<div class="app-header"><div class="app-logo">{icono("objetivo", 30)}</div>'
    f'<div><div class="app-titulo">Dashboard OTIF (On Time, In Full)</div>'
    f'<div class="app-sub">Nivel de servicio de proveedores: entregas a tiempo y completas, '
    f'cruzando archivos de SAP con los mapeos de centros y compradores.</div></div></div>',
    unsafe_allow_html=True,
)

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
        lambda r: 'A Tiempo' if r['On_Time']
        else ('Pendiente' if pd.isna(r['Fecha_Ingreso_SAP']) else 'Atrasado'),
        axis=1
    )
    df_otif['Estado In Full'] = df_otif['In_Full'].map({True: 'Completo', False: 'Incompleto'})
    df_otif['Estado OTIF'] = df_otif['OTIF'].map({True: 'Cumple OTIF', False: 'No Cumple'})

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
    columnas = list(df.columns)
    cabecera = "".join(f"<th>{html.escape(str(c))}</th>" for c in columnas)

    filas = []
    if df.empty:
        filas.append(
            f'<tr class="fila-vacia"><td colspan="{len(columnas)}">'
            f'No hay posiciones para los filtros seleccionados.</td></tr>'
        )
    else:
        ultima = len(df) - 1
        for i, (_, fila) in enumerate(df.iterrows()):
            celdas = []
            for c in columnas:
                valor = str(fila[c])
                if c == "OTIF":
                    try:
                        pct = max(0, min(100, int(valor.rstrip("%"))))
                    except ValueError:
                        pct = 0
                    celdas.append(
                        f'<td><div class="otif-cell"><div class="otif-bar"><span style="width:{pct}%"></span></div>'
                        f'<span class="otif-val">{html.escape(valor)}</span></div></td>'
                    )
                else:
                    celdas.append(f"<td>{html.escape(valor)}</td>")
            clase = ' class="fila-total"' if i == ultima else ""
            filas.append(f"<tr{clase}>{''.join(celdas)}</tr>")

    tabla = (
        f'<div class="table-container"><table class="custom-summary-table">'
        f'<thead><tr>{cabecera}</tr></thead><tbody>{"".join(filas)}</tbody></table></div>'
    )
    st.markdown(tabla, unsafe_allow_html=True)

def estilo_estado(valor):
    color = COLOR_ESTADO.get(valor)
    return f"color: {color}; font-weight: 600;" if color else ""

def preparar_tabla_detalle(df_mostrar):
    """Colorea los estados en la tabla de detalle. Si no es posible, devuelve la tabla sin estilo."""
    cols_estado = [c for c in ['Estado On Time', 'Estado In Full', 'Estado OTIF'] if c in df_mostrar.columns]
    if not cols_estado or df_mostrar.size > 250_000:
        return df_mostrar
    try:
        def fmt_num(v):
            if pd.isna(v):
                return ""
            return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}"

        styler = df_mostrar.style.format(na_rep="")
        cols_float = df_mostrar.select_dtypes(include="float").columns.tolist()
        if cols_float:
            styler = styler.format(fmt_num, subset=cols_float)
        aplicar = getattr(styler, "map", None) or styler.applymap
        return aplicar(estilo_estado, subset=cols_estado)
    except Exception:
        return df_mostrar


# ==========================================
# BARRA LATERAL: ORIGEN DE DATOS Y FILTRO TIEMPO
# ==========================================
with st.sidebar:
    encabezado_seccion("base_datos", "Origen de datos", "Desde dónde se leen los archivos de SAP.")
    origen_datos = st.radio(
        "Selecciona el origen:",
        ["OneDrive (automático)", "Subir archivos", "Archivos locales (data/)"],
        label_visibility="collapsed"
    )

    file_me2m, file_me80fn, file_grupo, file_centro = None, None, None, None

    if origen_datos == "OneDrive (automático)":
        if st.button("Recargar datos desde OneDrive", use_container_width=True):
            st.cache_data.clear()
            st.toast("Caché borrada. Descargando de nuevo desde OneDrive...")
            st.rerun()

        file_me2m = ONEDRIVE_URLS["me2m"]
        file_me80fn = ONEDRIVE_URLS["me80fn"]
        file_grupo = ONEDRIVE_URLS["grupo"]
        file_centro = ONEDRIVE_URLS["centro"]

    elif origen_datos == "Subir archivos":
        st.divider()
        encabezado_seccion("subir", "Carga manual", "Sube los 4 archivos en formato .xlsx.")
        file_me2m = st.file_uploader("ME2M: pedidos de compra", type=["xlsx"])
        file_me80fn = st.file_uploader("ME80FN: recepciones", type=["xlsx"])
        file_grupo = st.file_uploader("Responsable por grupo de compras", type=["xlsx"])
        file_centro = st.file_uploader("Centro y sociedad", type=["xlsx"])

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
        archivos = [
            ("ME2M", "Pedidos de compra", file_me2m),
            ("ME80FN", "Recepciones de mercancía", file_me80fn),
            ("Grupo", "Responsable por grupo de compras", file_grupo),
            ("Centro", "Centro y sociedad", file_centro),
        ]
        cargados = sum(1 for *_, f in archivos if f)
        items = []
        for nombre, detalle, f in archivos:
            if f:
                marca = f'<span class="estado-ok">{icono("check_circulo", 20)}</span>'
            else:
                marca = f'<span class="estado-pend">{icono("pendiente", 20)}</span>'
            items.append(
                f'<li>{marca}<div><div class="nombre">{nombre}</div>'
                f'<div class="detalle">{detalle}</div></div></li>'
            )
        st.markdown(
            f'<div class="carga"><div class="sec-head" style="margin:0">'
            f'<span class="sec-icon">{icono("archivo", 22)}</span>'
            f'<div><div class="sec-titulo">Carga los archivos para comenzar</div>'
            f'<div class="sec-desc">{cargados} de 4 cargados. Usa la barra lateral para subir los que faltan.</div>'
            f'</div></div><ul class="checklist">{"".join(items)}</ul></div>',
            unsafe_allow_html=True,
        )
else:
    listos_para_procesar = True

if listos_para_procesar:
    try:
        df_base = procesar_otif(file_me2m, file_me80fn, file_centro, file_grupo)
    except Exception as e:
        aviso(f"<b>No se pudieron cargar los datos.</b><br>{html.escape(str(e))}", "error")
        st.stop()

    # Filtro de tiempo en la barra lateral
    with st.sidebar:
        st.divider()
        encabezado_seccion("calendario", "Periodo", "Vacío incluye todas las semanas.")
        semanas_disp = sorted([s for s in df_base['Semana/Año'].unique() if s != 'Sin Fecha'])
        if 'Sin Fecha' in df_base['Semana/Año'].unique():
            semanas_disp.append('Sin Fecha')
        semana_sel = st.multiselect("Semana / Año", semanas_disp, placeholder="Todas las semanas")

        st.divider()
        if st.button("Cerrar sesión", use_container_width=True):
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

    # ==========================================
    # PANEL DE FILTROS (ST.FORM)
    # ==========================================
    with st.form("panel_filtros"):
        encabezado_seccion(
            "filtro",
            "Filtros",
            "Elige los valores y presiona <b>Aplicar filtros</b> para actualizar. Un filtro vacío incluye todo.",
        )

        # Fila 1 de filtros
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)
        with f_col1:
            centro_sel = st.multiselect("Centro logístico", centros, placeholder="Todos")
        with f_col2:
            grupo_sel = st.multiselect("Grupo de compras", grupos, placeholder="Todos")
        with f_col3:
            comprador_sel = st.multiselect("Comprador", compradores, placeholder="Todos")
        with f_col4:
            prov_sel = st.multiselect("Proveedor", proveedores, placeholder="Todos")

        # Fila 2 de filtros
        f_col5, f_col6, f_col7 = st.columns(3)
        with f_col5:
            on_time_sel = st.multiselect("Estado On Time", estados_on_time, placeholder="Todos")
        with f_col6:
            in_full_sel = st.multiselect("Estado In Full", estados_in_full, placeholder="Todos")
        with f_col7:
            otif_sel = st.multiselect("Estado OTIF", estados_otif, placeholder="Todos")

        btn_aplicar = st.form_submit_button("Aplicar filtros", type="primary", use_container_width=True)

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

    encabezado_seccion("grafico", "Indicadores de cumplimiento", "Calculados sobre las posiciones que cumplen los filtros.")
    tarjetas_kpi([
        {"icono": "lista", "titulo": "Líneas activas", "valor": f"{total_lineas:,}", "pct": None,
         "ayuda": "Posiciones de OC incluidas en el cálculo."},
        {"icono": "reloj", "titulo": "On Time", "valor": f"{pct_on_time:.1f}%", "pct": pct_on_time,
         "ayuda": "Recepción en SAP en o antes de la fecha estadística."},
        {"icono": "caja", "titulo": "In Full", "valor": f"{pct_in_full:.1f}%", "pct": pct_in_full,
         "ayuda": "Sin cantidad pendiente por entregar."},
        {"icono": "check_circulo", "titulo": "OTIF global", "valor": f"{pct_otif:.1f}%", "pct": pct_otif,
         "ayuda": "Cumple a tiempo y completo a la vez.", "principal": True},
    ])

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

    encabezado_seccion("persona", "Por comprador", "Posiciones y porcentaje OTIF de cada comprador.")
    renderizar_tabla_html(df_t1)

    st.write("")

    col_izq, col_der = st.columns(2)
    with col_izq:
        encabezado_seccion("planta", "Por centro logístico", "Vista fija: el total calza con la vista por comprador.")
        renderizar_tabla_html(df_t2)

    with col_der:
        encabezado_seccion("ubicacion", "Detalle por centro", "Centros activos según los filtros aplicados.")
        renderizar_tabla_html(df_t3)

    st.divider()

    encabezado_seccion(
        "tabla",
        "Detalle de posiciones y trazabilidad",
        "Estados en azul cumplen, en rojo no cumplen y en ámbar siguen pendientes de recepción.",
    )
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
        "Fecha_Estadistica": st.column_config.DateColumn("Fecha estadística", format="DD/MM/YYYY"),
        "Fecha_Ingreso_SAP": st.column_config.DateColumn("Fecha ingreso SAP", format="DD/MM/YYYY"),
    }
    st.dataframe(
        preparar_tabla_detalle(df_mostrar),
        use_container_width=True,
        hide_index=True,
        column_config=config_columnas,
    )

    st.divider()
    encabezado_seccion(
        "descarga",
        "Descarga de reportes",
        "Excel con el detalle filtrado y las tres tablas resumen en hojas separadas.",
    )

    excel_bytes = generar_excel_resumen(df_mostrar, df_t1, df_t2, df_t3)

    st.download_button(
        label="Descargar reporte de trazabilidad y resumen (.xlsx)",
        data=excel_bytes,
        file_name="Trazabilidad_OTIF_Semanal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
