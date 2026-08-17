"""
Configuración central del reporte Dx Compradores.

Hoy lee archivos locales. Cuando se migre a Azure, solo cambian
las funciones de carga en loaders.py (SharePoint -> Graph API,
Data -> Blob Storage), este archivo no debería necesitar tocarse
salvo para leer las variables desde Key Vault / variables de entorno.
"""
from pathlib import Path
import os
from datetime import date

# ---- Rutas locales (fase actual: trabajo local) ----
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

RUTA_DATA_ME5A = os.getenv("RUTA_DATA_ME5A", str(DATA_DIR / "ME5A_con_Ariba.xlsx"))
RUTA_RESP_GRUPO_COMPRAS = os.getenv("RUTA_RESP_GRUPO_COMPRAS", str(DATA_DIR / "Responsable_Grupo_Compras.xlsx"))
RUTA_CENTRO_SOCIEDAD = os.getenv("RUTA_CENTRO_SOCIEDAD", str(DATA_DIR / "Centro_Sociedad_MRO.xlsx"))
RUTA_RESP_MRP = os.getenv("RUTA_RESP_MRP", str(DATA_DIR / "Responsable_MRP.xlsx"))

# ---- Parámetro que hoy se ingresa manualmente en Power Query ----
# En el pbix: "FechaCorteReporte" se escribe a mano antes de actualizar.
# Acá se calcula solo (hoy), pero se puede sobreescribir por variable de entorno
# o por un selector en el sidebar de Streamlit.
FECHA_CORTE_REPORTE_DEFAULT = date.today()

# ---- Reglas de negocio (extraídas del código M real de la query "Data (2)") ----
# SLA de días de gestión: distinto según el origen de la solicitud.
SLA_DIAS_ERP_MRP = 10   # Solped MRP == "ERP" o "MRP"
SLA_DIAS_ARIBA = 7      # Solped MRP == "Ariba"

# Umbral de "Solicitud de pedido" que distingue Ariba (números grandes) de ERP manual.
UMBRAL_SOLICITUD_ARIBA = 6_000_000_000

# Días de gracia para "Aplica?": una solicitud "Sin pedido" con menos de estos días
# de antigüedad no se considera aún en el análisis de cumplimiento.
DIAS_GRACIA_SIN_PEDIDO = 7
