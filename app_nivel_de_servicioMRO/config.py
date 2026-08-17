"""
Configuración centralizada y reglas de negocio para la aplicación
de Nivel de Servicio MRO.
"""

# ---- Rutas de archivos por defecto (Entorno Local) ----
PATH_ME5A = "data/ME5A_con_Ariba.xlsx"
PATH_RESPONSABLE_GRUPO = "data/Responsable_Grupo_Compras.xlsx"
PATH_CENTRO_SOCIEDAD = "data/Centro_Sociedad_MRO.xlsx"
PATH_RESPONSABLE_MRP = "data/Responsable_MRP.xlsx"

# ---- Reglas de SLA (Acuerdos de Nivel de Servicio) ----
SLA_DIAS_ERP_MRP = 10  # Días límite para considerar "Cumple" en solicitudes ERP y MRP
SLA_DIAS_ARIBA = 7     # Días límite para considerar "Cumple" en solicitudes Ariba

# ---- Umbrales y Parámetros del Negocio ----
UMBRAL_SOLPED_ARIBA = 6_000_000_000  # Solicitudes >= a este número provienen de Ariba
DIAS_GRACIA_SIN_PEDIDO = 7           # Días de margen antes de clasificar como "Aplica" sin pedido
