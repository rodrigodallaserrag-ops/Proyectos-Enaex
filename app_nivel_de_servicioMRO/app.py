"""
Streamlit - Cuadro Comparativo Integrado y Planilla de Gestión SOLPED
Gestión multimoneda reactiva con persistencia de estado para Costo Transporte y Carga Manual.



st.set_page_config(page_title="Cuadro Comparativo y Planilla de Gestión", layout="wide")

# ==============================================================================
# ESTADO GLOBAL DE SESIÓN (SESSION STATE)
# ==============================================================================
if "items_data" not in st.session_state:
    st.session_state["items_data"] = []

if "uploaded_file_name" not in st.session_state:
    st.session_state["uploaded_file_name"] = None

if "costo_transporte_global" not in st.session_state:
    st.session_state["costo_transporte_global"] = 0.0

# ==============================================================================
# FUNCIONES AUXILIARES DE CONVERSIÓN DE MONEDA
# ==============================================================================
def convertir_a_moneda_base(monto, moneda_origen, moneda_destino="USD", usd_clp=950.0, eur_clp=1020.0):
    if pd.isna(monto) or monto == 0:
        return 0.0
    
    mon_orig = str(moneda_origen).upper().strip() if pd.notna(moneda_origen) else "CLP"
    mon_dest = str(moneda_destino).upper().strip()
    
    if mon_orig in ["CLP", "CLF", "PESO", "PESOS"]:
        monto_clp = float(monto)
    elif mon_orig in ["USD", "US$", "DOLARES", "DOLAR"]:
        monto_clp = float(monto) * usd_clp
    elif mon_orig in ["EUR", "EUR$", "EUROS"]:
        monto_clp = float(monto) * eur_clp
    else:
        monto_clp = float(monto)
        
    if mon_dest == "CLP":
        return monto_clp
    elif mon_dest == "USD":
        return monto_clp / usd_clp if usd_clp > 0 else monto_clp
    elif mon_dest == "EUR":
        return monto_clp / eur_clp if eur_clp > 0 else monto_clp
    
    return monto_clp

def encontrar_columna(df, posibles_nombres):
    """Busca coincidencias de nombres de columnas ignorando mayúsculas/minúsculas."""
    for col in df.columns:
        if str(col).strip().lower() in [n.lower() for n in posibles_nombres]:
            return col
    return None

def cargar_excel_a_state(archivo):
    """Parsea el archivo Excel cargado y lo inyecta en el estado global."""
    df_raw = pd.read_excel(archivo)
    col_desc = encontrar_columna(df_raw, ["Descripción", "Material", "Texto breve", "Texto de material", "Detalle"])
    col_cant = encontrar_columna(df_raw, ["Cantidad", "Cant", "CANTIDAD", "Cant."])
    col_prov = encontrar_columna(df_raw, ["Proveedor", "Licitante", "Nombre Proveedor"])
    col_precio_mat = encontrar_columna(df_raw, ["Valor neto de pedido", "Valor unitario", "Precio Material", "Precio Unitario", "Precio Neto"])
    col_moneda_mat = encontrar_columna(df_raw, ["Moneda", "Moneda Material", "CM"])
    col_precio_trans = encontrar_columna(df_raw, ["Precio Transporte", "Valor Flete", "Flete", "Transporte", "Envio", "Costo Envio", "Costo Transporte"])
    col_centro = encontrar_columna(df_raw, ["Centro", "Planta", "Ubicación"])

    items = []
    for idx, row in df_raw.iterrows():
        pos_num = (idx + 1) * 10
        items.append({
            "posicion": pos_num,
            "descripcion": str(row[col_desc]) if col_desc and pd.notna(row[col_desc]) else f"POSICIÓN {pos_num}",
            "centro": str(row[col_centro]) if col_centro and pd.notna(row[col_centro]) else "E024 (Planta Rinconada)",
            "cantidad": float(row[col_cant]) if col_cant and pd.notna(row[col_cant]) else 1.0,
            "precio_unitario": float(row[col_precio_mat]) if col_precio_mat and pd.notna(row[col_precio_mat]) else 0.0,
            "moneda": str(row[col_moneda_mat]).upper().strip() if col_moneda_mat and pd.notna(row[col_moneda_mat]) else "CLP",
            "proveedor": str(row[col_prov]) if col_prov and pd.notna(row[col_prov]) else "Proveedor Desconocido",
            "incoterm": "EXW",
            "costo_transporte": float(row[col_precio_trans]) if col_precio_trans and pd.notna(row[col_precio_trans]) else 0.0,
            "fecha_entrega": date.today()
        })
    st.session_state["items_data"] = items

# ==============================================================================
# CONFIGURACIÓN TEMA (CLARO PREDETERMINADO / OSCURO)
# ==============================================================================
if "tema" not in st.session_state:
    st.session_state["tema"] = "claro"

icono_tema = "🌙" if st.session_state["tema"] == "claro" else "☀️"
if st.button(icono_tema, key="theme_toggle", help="Alternar Modo Claro/Oscuro"):
    st.session_state["tema"] = "oscuro" if st.session_state["tema"] == "claro" else "claro"
    st.rerun()

if st.session_state["tema"] == "claro":
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important; top: 65px !important; right: 15px !important;
            z-index: 999999 !important; width: 45px !important; height: 45px !important; min-width: 0 !important; 
        }
        .st-key-theme_toggle button {
            background: #FFFFFF !important; border: 1px solid #E0E0E0 !important;
            border-radius: 50% !important; box-shadow: 0 2px 5px rgba(0,0,0,0.1) !important;
            font-size: 1.4rem !important; padding: 0 !important; margin: 0 !important;
            color: #111111 !important; display: flex !important; align-items: center !important;
            justify-content: center !important; width: 100% !important; height: 100% !important; min-height: unset !important;
        }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #F0F0F0 !important; }
        </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
        <style>
        .st-key-theme_toggle {
            position: fixed !important; top: 65px !important; right: 15px !important;
            z-index: 999999 !important; width: 45px !important; height: 45px !important; min-width: 0 !important;
        }
        .st-key-theme_toggle button {
            background: #1E2329 !important; border: 1px solid #444444 !important;
            border-radius: 50% !important; box-shadow: 0 2px 5px rgba(0,0,0,0.3) !important;
            font-size: 1.4rem !important; padding: 0 !important; margin: 0 !important;
            color: #FF3333 !important; display: flex !important; align-items: center !important;
            justify-content: center !important; width: 100% !important; height: 100% !important; min-height: unset !important;
        }
        .st-key-theme_toggle button p { margin: 0 !important; padding: 0 !important; line-height: 1 !important; font-size: 1.4rem !important; }
        .st-key-theme_toggle button:hover { transform: scale(1.1) !important; background: #2C323A !important; border-color: #FF3333 !important; }
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], html, body, [data-testid="stHeader"] {
            background-color: #0E1117 !important; color: #FF3333 !important;
        }
        p, span, label, h1, h2, h3, h4, h5, h6, div, td, th, caption, .stMarkdown { color: #FF3333 !important; }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button) {
            background-color: #CC0000 !important; color: #FFFFFF !important; border: 1px solid #FF4D4D !important; font-weight: bold !important;
        }
        div[data-testid="stButton"] > button:not(.st-key-theme_toggle button):hover { background-color: #FF0000 !important; color: #FFFFFF !important; border-color: #FF6666 !important; }
        input, select, textarea, div[data-baseweb="select"] { background-color: #1E2329 !important; color: #FF3333 !important; border-color: #CC0000 !important; }
        [data-testid="stForm"], div[data-testid="stVerticalBlock"] > div:has(input[type="password"]) {
            background-color: #1E2329 !important; padding: 2rem !important; border-radius: 12px !important; border: 1px solid #CC0000 !important;
        }
        </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# ACCESO CON CONTRASEÑA
# ==============================================================================
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("📊 Cuadro Comparativo y Planilla de Gestión")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password")
        if st.button("Ingresar"):
            if clave_ingresada == st.secrets["app_password"]:
                st.session_state["_autenticado"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

# ==============================================================================
# INTERFAZ PRINCIPAL Y BARRA LATERAL
# ==============================================================================
st.title("📊 Cuadro Comparativo y Planilla de Gestión")

with st.sidebar:
    st.header("⚙️ Configuración de Datos")
    archivo_ofertas = st.file_uploader("Subir matriz de ofertas (.xlsx)", type=["xlsx"])
    
    if archivo_ofertas is not None:
        if st.session_state["uploaded_file_name"] != archivo_ofertas.name:
            st.session_state["uploaded_file_name"] = archivo_ofertas.name
            cargar_excel_a_state(archivo_ofertas)
            st.success("Planilla Excel cargada correctamente.")

    st.divider()
    st.header("💱 Conversión Multimoneda")
    moneda_base = st.selectbox("Moneda Consolidada Base", ["USD", "CLP", "EUR"], index=0)
    tasa_usd_clp = st.number_input("Tasa USD / CLP", value=950.0, step=1.0)
    tasa_eur_clp = st.number_input("Tasa EUR / CLP", value=1020.0, step=1.0)

# ==============================================================================
# SECCIÓN: CARGA MANUAL / DIRECTA
# ==============================================================================
with st.expander("➕ Carga Manual / Directa de Posición SOLPED", expanded=False):
    with st.form("form_carga_manual", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            man_pos = st.number_input("Posición", value=10, step=10)
            man_desc = st.text_input("Descripción Material", value="PARTIDOR SUAVE ABB 16A")
            man_centro = st.text_input("Centro / Planta", value="E024 (Planta Rinconada)")
        with col2:
            man_cant = st.number_input("Cantidad", value=1.0, step=1.0)
            man_pu = st.number_input("Precio Unitario", value=283000.0, step=1000.0)
            man_moneda = st.selectbox("Moneda", ["CLP", "USD", "EUR"])
        with col3:
            man_prov = st.text_input("Proveedor", value="AUTOMATION BUSS")
            man_inco = st.selectbox("Incoterm / Transporte", ["EXW", "DDP", "FOB", "CIF", "CPT"])
            man_costo_trans = st.number_input("Costo Transporte ($)", value=109000.0, step=1000.0, help="Ejemplo: 109.000 CLP")
            man_fecha = st.date_input("Fecha Entrega", value=date.today())
        
        btn_add = st.form_submit_button("➕ Agregar Posición a Evaluación")
        if btn_add:
            st.session_state["items_data"].append({
                "posicion": man_pos,
                "descripcion": man_desc,
                "centro": man_centro,
                "cantidad": man_cant,
                "precio_unitario": man_pu,
                "moneda": man_moneda,
                "proveedor": man_prov,
                "incoterm": man_inco,
                "costo_transporte": man_costo_trans,
                "fecha_entrega": man_fecha
            })
            st.success(f"Posición {man_pos} agregada correctamente con Costo Transporte de ${man_costo_trans:,.0f}.")
            st.rerun()

# ==============================================================================
# PESTAÑAS PRINCIPALES
# ==============================================================================
tab_planilla, tab_comparativo = st.tabs(["📝 Evaluación SOLPED", "📊 Cuadro Comparativo Integrado"])

# ------------------------------------------------------------------------------
# PESTAÑA 1: EVALUACIÓN SOLPED (EDICIÓN DIRECTA EN TARJETAS)
# ------------------------------------------------------------------------------
with tab_planilla:
    st.subheader("Seguimiento y Control de Adjudicaciones (SOLPED)")
    
    if not st.session_state["items_data"]:
        st.info("No hay posiciones cargadas. Sube un archivo Excel en la barra lateral o usa '➕ Carga Manual / Directa' arriba.")
    else:
        st.caption("Modifica el **Costo Transporte** o cualquier valor en las tarjetas. Los datos se actualizarán automáticamente en el Cuadro Comparativo Integrado.")
        
        indices_a_eliminar = []
        
        for idx, item in enumerate(st.session_state["items_data"]):
            with st.container(border=True):
                col_header, col_del = st.columns([6, 1])
                with col_header:
                    st.markdown(f"### SOLPED: 2026-02-20 10:08:00 | Pos {item['posicion']}: {item['descripcion']}")
                    st.caption(f"Centro: {item['centro']} | UM Original: 2023-02-13 00:00:00")
                with col_del:
                    if st.button("🗑️ Eliminar", key=f"del_{idx}"):
                        indices_a_eliminar.append(idx)

                c1, c2, c3, c4, c5, c6, c7 = st.columns([1.0, 1.8, 1.2, 2.2, 1.2, 1.6, 1.5])
                
                with c1:
                    st.session_state["items_data"][idx]["cantidad"] = st.number_input(
                        "Cantidad", value=float(item["cantidad"]), min_value=0.0, step=1.0, key=f"cant_{idx}"
                    )
                with c2:
                    st.session_state["items_data"][idx]["precio_unitario"] = st.number_input(
                        "Precio Unitario", value=float(item["precio_unitario"]), step=100.0, key=f"pu_{idx}"
                    )
                with c3:
                    st.session_state["items_data"][idx]["moneda"] = st.selectbox(
                        "Moneda", ["CLP", "USD", "EUR"], 
                        index=["CLP", "USD", "EUR"].index(item["moneda"]) if item["moneda"] in ["CLP", "USD", "EUR"] else 0, 
                        key=f"mon_{idx}"
                    )
                with c4:
                    st.session_state["items_data"][idx]["proveedor"] = st.text_input(
                        "Proveedor", value=item["proveedor"], key=f"prov_{idx}"
                    )
                with c5:
                    st.session_state["items_data"][idx]["incoterm"] = st.selectbox(
                        "Transporte", ["EXW", "DDP", "FOB", "CIF", "CPT"], 
                        index=["EXW", "DDP", "FOB", "CIF", "CPT"].index(item["incoterm"]) if item["incoterm"] in ["EXW", "DDP", "FOB", "CIF", "CPT"] else 0,
                        key=f"inco_{idx}"
                    )
                with c6:
                    st.session_state["items_data"][idx]["costo_transporte"] = st.number_input(
                        "Costo Transp.", value=float(item["costo_transporte"]), step=1000.0, key=f"ct_{idx}"
                    )
                with c7:
                    st.session_state["items_data"][idx]["fecha_entrega"] = st.date_input(
                        "Fecha Entrega", value=item["fecha_entrega"], key=f"fecha_{idx}"
                    )

                subtotal = st.session_state["items_data"][idx]["cantidad"] * st.session_state["items_data"][idx]["precio_unitario"]
                ct = st.session_state["items_data"][idx]["costo_transporte"]
                mon = st.session_state["items_data"][idx]["moneda"]
                st.info(f"**Total Posición ({mon}):** {(subtotal + ct):,.2f} *(Subtotal Material: {subtotal:,.2f} + Costo Transporte: {ct:,.2f})*")

        if indices_a_eliminar:
            for idx in sorted(indices_a_eliminar, reverse=True):
                st.session_state["items_data"].pop(idx)
            st.rerun()

# ------------------------------------------------------------------------------
# PESTAÑA 2: CUADRO COMPARATIVO INTEGRADO (CÁLCULO DINÁMICO EN TIEMPO REAL)
# ------------------------------------------------------------------------------
with tab_comparativo:
    st.subheader("Análisis de Ofertas y Proveedores (Integrado)")
    
    if not st.session_state["items_data"]:
        st.info("No hay datos para comparar. Carga una planilla Excel o agrega ítems manualmente.")
    else:
        df_comp = pd.DataFrame(st.session_state["items_data"])
        
        monto_mat_base = []
        monto_trans_base = []
        
        for _, row in df_comp.iterrows():
            subtotal_mat = row["cantidad"] * row["precio_unitario"]
            m_mat = convertir_a_moneda_base(subtotal_mat, row["moneda"], moneda_base, tasa_usd_clp, tasa_eur_clp)
            m_trans = convertir_a_moneda_base(row["costo_transporte"], row["moneda"], moneda_base, tasa_usd_clp, tasa_eur_clp)
            
            monto_mat_base.append(m_mat)
            monto_trans_base.append(m_trans)
            
        df_comp["Subtotal Material (Origen)"] = df_comp["cantidad"] * df_comp["precio_unitario"]
        df_comp[f"Monto Material ({moneda_base})"] = monto_mat_base
        df_comp[f"Costo Transporte ({moneda_base})"] = monto_trans_base
        df_comp[f"Monto Total ({moneda_base})"] = df_comp[f"Monto Material ({moneda_base})"] + df_comp[f"Costo Transporte ({moneda_base})"]
        
        total_mat_gen = sum(monto_mat_base)
        total_trans_items = sum(monto_trans_base)

        # ----------------------------------------------------------------------
        # CAMPO INTERACTIVO: COSTO TRANSPORTE GLOBAL AL LADO DE LOS TOTALES
        # ----------------------------------------------------------------------
        col_input_trans, col_m1, col_m2, col_m3 = st.columns([1.3, 1, 1, 1])
        
        with col_input_trans:
            costo_trans_global = st.number_input(
                f"🚚 Costo Transporte Global ({moneda_base})",
                value=float(st.session_state.get("costo_transporte_global", 0.0)),
                step=1000.0,
                key="input_costo_transporte_global",
                help="Ingresa un costo de transporte adicional para calcularlo dinámicamente con el total."
            )
            st.session_state["costo_transporte_global"] = costo_trans_global

        total_trans_gen = total_trans_items + costo_trans_global
        total_total_gen = total_mat_gen + total_trans_gen

        with col_m1:
            st.metric(f"Total Materiales ({moneda_base})", f"{total_mat_gen:,.2f}")
        with col_m2:
            st.metric(f"Total Transporte ({moneda_base})", f"{total_trans_gen:,.2f}")
        with col_m3:
            st.metric(f"Monto Total Consolidado ({moneda_base})", f"{total_total_gen:,.2f}")
        
        st.divider()
        
        cols_mostrar = [
            "posicion", "descripcion", "centro", "cantidad", "precio_unitario", 
            "moneda", "Subtotal Material (Origen)", "costo_transporte", "proveedor", "incoterm",
            f"Monto Material ({moneda_base})", f"Costo Transporte ({moneda_base})", f"Monto Total ({moneda_base})"
        ]
        
        st.dataframe(
            df_comp[cols_mostrar],
            use_container_width=True,
            height=500,
            column_config={
                "posicion": "Pos",
                "descripcion": "Descripción Material",
                "centro": "Centro",
                "cantidad": "Cant",
                "precio_unitario": st.column_config.NumberColumn("P. Unitario", format="%.2f"),
                "moneda": "Moneda",
                "Subtotal Material (Origen)": st.column_config.NumberColumn("Subtotal Material", format="%.2f"),
                "costo_transporte": st.column_config.NumberColumn("Costo Transporte (Origen)", format="%.2f"),
                "proveedor": "Proveedor",
                "incoterm": "Incoterm",
                f"Monto Material ({moneda_base})": st.column_config.NumberColumn(f"Mat. ({moneda_base})", format="%.2f"),
                f"Costo Transporte ({moneda_base})": st.column_config.NumberColumn(f"Transp. ({moneda_base})", format="%.2f"),
                f"Monto Total ({moneda_base})": st.column_config.NumberColumn(f"Total ({moneda_base})", format="%.2f"),
            }
        )
