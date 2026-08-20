"""
Preparar datos — Conversión ME5A_con_Ariba.xlsx -> .parquet

Por qué esta página existe: el ME5A_con_Ariba.xlsx es, con diferencia, el
archivo más pesado de los 4 que carga el reporte. Leer un .xlsx con pandas
(vía openpyxl) tiene un pico de memoria varias veces mayor que el tamaño
final del DataFrame — el motor arma un árbol de objetos XML completo antes
de convertir a tabla. Parquet es binario y columnar: se lee casi directo a
memoria, con una fracción del pico de RAM y en menos tiempo.

Uso: sube acá tu ME5A_con_Ariba.xlsx una vez, descarga el .parquet resultante,
y desde ahora súbelo a la app principal en vez del .xlsx — el pipeline y los
filtros no cambian en nada, solo la fuente de carga.
"""
import io

import pandas as pd
import streamlit as st

import loaders

st.set_page_config(page_title="Preparar datos - Nivel de Servicio", layout="wide")

st.title("Preparar datos — Conversión a Parquet")
st.caption(
    "Convierte tu ME5A_con_Ariba.xlsx a .parquet: mismo contenido, mismo tipado, "
    "pero con una carga mucho más rápida y liviana en memoria dentro de la app principal."
)

archivo = st.file_uploader("ME5A_con_Ariba.xlsx", type="xlsx")

if archivo is None:
    st.info("Sube el archivo ME5A_con_Ariba.xlsx para convertirlo.")
    st.stop()

with st.spinner("Leyendo y tipando el Excel..."):
    df_crudo = pd.read_excel(archivo, sheet_name="Data")
    df_tipado = loaders._tipar_data_pr(df_crudo)

st.success(f"Leído correctamente: {len(df_tipado):,} filas, {len(df_tipado.columns)} columnas.")

with st.expander("Ver una muestra de los datos tipados"):
    st.dataframe(df_tipado.head(50), use_container_width=True)

# ---- Comparación de tamaño/memoria: xlsx original vs parquet resultante ----
buffer_parquet = io.BytesIO()
df_tipado.to_parquet(buffer_parquet, index=False)
bytes_parquet = buffer_parquet.getvalue()

tam_xlsx_mb = archivo.size / 1024 / 1024
tam_parquet_mb = len(bytes_parquet) / 1024 / 1024

c1, c2 = st.columns(2)
with c1:
    st.metric("Tamaño .xlsx original", f"{tam_xlsx_mb:.1f} MB")
with c2:
    st.metric(
        "Tamaño .parquet resultante",
        f"{tam_parquet_mb:.1f} MB",
        delta=f"{tam_parquet_mb - tam_xlsx_mb:.1f} MB",
        delta_color="inverse",
    )

st.download_button(
    "⬇ Descargar ME5A_con_Ariba.parquet",
    data=bytes_parquet,
    file_name="ME5A_con_Ariba.parquet",
    mime="application/octet-stream",
)

st.caption(
    "Una vez descargado, ve a la pestaña del reporte principal y súbelo ahí en vez "
    "del .xlsx — el resto del flujo (filtros, tablas, exportación) funciona exactamente igual."
)
