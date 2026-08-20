import streamlit as st
import pandas as pd

# Configuración básica de la página
st.set_page_config(page_title="Dashboard OTIF - KPI 2", layout="wide")

st.title("📦 Dashboard OTIF — Nivel de Servicio Proveedores")
st.write("Bienvenido a la nueva aplicación de OTIF. Por favor, carga los archivos necesarios en el panel lateral.")

# Panel lateral para subir las 5 tablas que tenías en Power Query
with st.sidebar:
    st.header("Carga de Datos")
    
    archivo_base_oc = st.file_uploader("1. Cargar Base OC", type=["xlsx", "csv"])
    archivo_recepcion = st.file_uploader("2. Cargar Recepción1", type=["xlsx", "csv"])
    archivo_resp = st.file_uploader("3. Cargar Responsable por Grupo", type=["xlsx", "csv"])
    archivo_listado = st.file_uploader("4. Cargar Listado OC Masiva", type=["xlsx", "csv"])
    archivo_centro = st.file_uploader("5. Cargar CENTRO_SOCIEDAD", type=["xlsx", "csv"])

# Lógica básica para mostrar que la app funciona cuando suban la Base OC
if archivo_base_oc:
    st.success("¡Base OC cargada correctamente!")
    
    # Leemos el archivo (asumiendo que es excel, puedes cambiar a csv si es necesario)
    df_base_oc = pd.read_excel(archivo_base_oc)
    
    st.subheader("Vista previa de Base OC")
    st.dataframe(df_base_oc.head())
    
    st.info("Próximos pasos: Aquí programaremos el cruce con la tabla Recepción1 para calcular el On-Time y el In-Full.")
else:
    st.info("👈 Sube los archivos en el menú de la izquierda para comenzar.")
