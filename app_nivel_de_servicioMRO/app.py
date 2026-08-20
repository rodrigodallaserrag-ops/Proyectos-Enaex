import pandas as pd
import streamlit as st

# Importación del archivo ariba_trazabilidad.py con alias para mantener compatibilidad
try:
    import ariba_trazabilidad as trazabilidad
    HAS_TRAZABILIDAD = True
except ImportError:
    HAS_TRAZABILIDAD = False

st.set_page_config(page_title="Trazabilidad Ariba", layout="wide")

# (Opcional) Si quieres mantener el botón de modo oscuro en esta página, 
# puedes copiar aquí el bloque de "CONFIGURACIÓN TEMA" de tu app.py original.

st.title("🔍 Trazabilidad PR No Catalogadas — Ariba")

if not HAS_TRAZABILIDAD:
    st.warning(
        "⚠️ **Módulo 'ariba_trazabilidad.py' no disponible.**\n\n"
        "Asegúrate de que el archivo esté subido en GitHub en la misma carpeta raíz."
    )
else:
    st.markdown(
        "Procesa el reporte de **Trazabilidad Ariba** sin consolidar para vincular "
        "las solicitudes de compra iniciales, sus agregadas y su salida a SAP ERP (Solped 600)."
    )

    traz_col1, traz_col2 = st.columns([2, 1])
    with traz_col1:
        archivo_trazabilidad = st.file_uploader(
            "Cargar Reporte PR No Catalogadas - Trazabilidad (.csv)",
            type=["csv"],
            key="uploader_trazabilidad",
        )
    with traz_col2:
        empresa_id = st.text_input(
            "Empresa compradora (ID)",
            value=getattr(trazabilidad, "EMPRESA_POR_DEFECTO", "1000"),
        )

    if archivo_trazabilidad:
        if st.button("🚀 Procesar Trazabilidad", key="btn_procesar_traz"):
            with st.spinner("Procesando trazabilidad y reconstruyendo cadena de eventos..."):
                try:
                    df_cadena, df_resumen = trazabilidad.procesar_trazabilidad_completa(
                        archivo_trazabilidad, empresa=empresa_id
                    )

                    # Guardamos en session_state para que 'app.py' pueda leerlo
                    st.session_state["df_trazabilidad_cadena"] = df_cadena
                    st.session_state["df_trazabilidad_limpio"] = df_resumen

                    # Borrar la caché del pipeline para obligar a recalcular la Pestaña Dx Compradores
                    st.session_state.pop("_clave_pipeline", None)
                    st.success("¡Trazabilidad procesada con éxito! La pestaña Dx Compradores ha sido actualizada.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al procesar el archivo: {e}")

    if "df_trazabilidad_limpio" in st.session_state and "df_trazabilidad_cadena" in st.session_state:
        df_cadena = st.session_state["df_trazabilidad_cadena"]
        df_resumen = st.session_state["df_trazabilidad_limpio"]

        kpi_t1, kpi_t2, kpi_t3 = st.columns(3)
        with kpi_t1:
            st.metric("Total PR Iniciales", f"{len(df_cadena):,}")
        with kpi_t2:
            completas = (df_cadena["Cadena completa"] == "Sí").sum() if "Cadena completa" in df_cadena.columns else 0
            st.metric("Cadenas Completas (con SAP)", f"{completas:,}")
        with kpi_t3:
            st.metric("Solpeds 600 Identificadas", f"{len(df_resumen):,}")

        st.divider()

        tab_cad, tab_res = st.tabs(["📋 Cadena Detallada", "📊 Resumen por Solped (SAP 600)"])

        with tab_cad:
            st.subheader("Vista Cadena de Eventos")
            st.dataframe(df_cadena, use_container_width=True)
            csv_cadena = df_cadena.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="⬇ Descargar Cadena Completa (CSV)",
                data=csv_cadena,
                file_name="Trazabilidad_Cadena_Completa.csv",
                mime="text/csv",
            )

        with tab_res:
            st.subheader("Vista Consolidada por Solped SAP 600")
            st.caption("Esta información servirá de cruce directo con la base ME5A.")
            st.dataframe(df_resumen, use_container_width=True)
            csv_resumen = df_resumen.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="⬇ Descargar Resumen Solpeds 600 (CSV)",
                data=csv_resumen,
                file_name="Resumen_Solped_600_No_Catalogadas.csv",
                mime="text/csv",
            )
