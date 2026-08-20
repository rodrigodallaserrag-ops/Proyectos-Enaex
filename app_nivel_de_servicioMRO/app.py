st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

# ---- Ocultar ítem de navegación del sidebar en todo momento ----
st.markdown("""
    <style>
    /* Ocultar enlaces del menú lateral que contengan 'trazabilidad' o 'ariba' */
    [data-testid="stSidebarNav"] a[href*="trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="Trazabilidad"],
    [data-testid="stSidebarNav"] a[href*="ariba"] {
        display: none !important;
    }
    
    /* Si deseas ocultar por completo todo el bloque de navegación automático del sidebar: */
    /* [data-testid="stSidebarNav"] { display: none !important; } */
    </style>
""", unsafe_allow_html=True)
