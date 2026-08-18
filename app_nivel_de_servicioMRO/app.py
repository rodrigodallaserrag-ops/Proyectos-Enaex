"""
Streamlit - Dx Compradores
Réplica del pbix "Nivel_de_servicio_BI.pbix", página "Dx Compradores".

Correr local:  streamlit run app.py
"""
import streamlit as st
import pandas as pd

import config, loaders, transform

st.set_page_config(page_title="Dx Compradores - Nivel de Servicio", layout="wide")

# ---- Acceso con contraseña: cada ingreso fuerza datos frescos desde OneDrive ----
# st.cache_data se comparte entre TODAS las sesiones/usuarios de esta misma
# instancia corriendo (no es por sesión individual) — así que aunque el ttl
# de _descargar_onedrive sea de 1 hora, otro usuario pudo haber disparado esa
# descarga hace 40 minutos y tú heredarías esos bytes. Para garantizar "datos
# frescos cada vez que ingreso la contraseña" sin depender del ttl, al
# autenticarse correctamente se limpia el caché explícitamente y se obliga
# una descarga nueva.
if "app_password" in st.secrets:
    if not st.session_state.get("_autenticado"):
        st.title("Dx Compradores — Nivel de Servicio")
        clave_ingresada = st.text_input("Contraseña de acceso", type="password")
        if st.button("Ingresar"):
            if clave_ingresada == st.secrets["app_password"]:
                st.session_state["_autenticado"] = True
                loaders._descargar_onedrive.clear()
                st.session_state.pop("_clave_pipeline", None)
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        st.stop()

st.title("Dx Compradores — Nivel de Servicio")

# ---- Fuente de datos: local (desarrollo), subida manual, u OneDrive automático ----
with st.sidebar:
    st.header("Datos de entrada")
    modo = st.radio(
        "Origen de datos",
        ["OneDrive (automático)", "Subir archivos", "Archivos locales (data/)"],
        index=0,
    )

    archivo_data = archivo_resp_grupo = archivo_centro = archivo_mrp = None
    if modo == "OneDrive (automático)":
        # Los 4 archivos se descargan solos desde los links guardados en
        # Secrets (ver loaders._descargar_onedrive para el formato esperado).
        # El caché con ttl=1h evita re-descargar en cada filtro, y se refresca
        # solo — no hace falta tocar nada cuando actualizas el ME5A semanal.
        archivo_data = "onedrive:me5a_parquet"
        archivo_resp_grupo = "onedrive:responsable_grupo_compras"
        archivo_centro = "onedrive:centro_sociedad_mro"
        archivo_mrp = "onedrive:responsable_mrp"

        if st.button("🔄 Forzar recarga desde OneDrive ahora"):
            loaders._descargar_onedrive.clear()
            st.session_state.pop("_clave_pipeline", None)
            st.rerun()

        if "onedrive" not in st.secrets:
            st.error(
                "Falta configurar los Secrets de OneDrive (Settings → Secrets). "
                "Mientras tanto, usa 'Subir archivos'."
            )
            st.stop()

    elif modo == "Subir archivos":
        archivo_data = st.file_uploader(
            "ME5A_con_Ariba (.xlsx o .parquet)", type=["xlsx", "parquet"],
            help="Si ya lo convertiste en la pestaña 'Preparar datos', sube el .parquet — carga mucho más rápido y liviano que el .xlsx.",
        )
        archivo_resp_grupo = st.file_uploader("Responsable_Grupo_Compras.xlsx", type="xlsx")
        archivo_centro = st.file_uploader("Centro_Sociedad_MRO.xlsx", type="xlsx")
        archivo_mrp = st.file_uploader("Responsable_MRP.xlsx", type="xlsx")
        if not all([archivo_data, archivo_resp_grupo, archivo_centro, archivo_mrp]):
            st.info("Sube los 4 archivos para generar el reporte.")
            st.stop()
    else:
        # Rutas locales por defecto cuando no se suben archivos manualmente.
        archivo_data = "data/ME5A_con_Ariba.xlsx"
        archivo_resp_grupo = "data/Responsable_Grupo_Compras.xlsx"
        archivo_centro = "data/Centro_Sociedad_MRO.xlsx"
        archivo_mrp = "data/Responsable_MRP.xlsx"

with st.sidebar:
    st.header("Parámetros")
    fecha_corte = st.date_input("Fecha de corte del reporte (FechaCorteReporte)", value=pd.Timestamp.today())
    st.caption(f"SLA: {config.SLA_DIAS_ERP_MRP} días ERP/MRP · {config.SLA_DIAS_ARIBA} días Ariba")

# ---- Carga de datos + pipeline: computar UNA sola vez por sesión ----
# Streamlit re-ejecuta este script completo en cada interacción (cada filtro
# que tocas). Sin este bloque, eso significa releer los 4 archivos y correr
# TODO el pipeline (merges + cálculo de Nivel de Servicio) en cada filtro,
# aunque filtrar en sí sea barato (isin()/between() sobre datos en memoria).
#
# @st.cache_data no es la herramienta correcta acá: para decidir si hay
# cache hit, tiene que hashear el CONTENIDO COMPLETO de los DataFrames en
# cada rerun — con un dataset grande, ese hashing es más caro que el propio
# cálculo, y además retiene copias completas en RAM (max_entries).
#
# La alternativa correcta para "calcular una vez, filtrar muchas veces"
# es st.session_state con una CLAVE DE INVALIDACIÓN barata (nombre + tamaño
# del archivo, no su contenido) — comparar esa clave es O(1), no O(n).
def _clave_archivo(archivo):
    """Identifica un archivo sin leer su contenido: (nombre, tamaño) para
    un UploadedFile, o la ruta tal cual para un path local."""
    if hasattr(archivo, "name") and hasattr(archivo, "size"):
        return (archivo.name, archivo.size)
    return archivo


clave_actual = (
    _clave_archivo(archivo_data),
    _clave_archivo(archivo_resp_grupo),
    _clave_archivo(archivo_centro),
    _clave_archivo(archivo_mrp),
    pd.Timestamp(fecha_corte),
)

if st.session_state.get("_clave_pipeline") != clave_actual:
    df_data = loaders.cargar_data_pr(archivo_data)
    df_resp_grupo = loaders.cargar_responsable_grupo_compras(archivo_resp_grupo)
    df_centro_sociedad = loaders.cargar_centro_sociedad_mro(archivo_centro)
    df_resp_mrp = loaders.cargar_responsable_mrp(archivo_mrp)

    df_calculado = transform.pipeline_completo(
        df_data, df_resp_grupo, df_centro_sociedad, df_resp_mrp, fecha_corte=pd.Timestamp(fecha_corte)
    )
    df_calculado["Año"] = df_calculado["Fecha de pedido"].dt.year
    df_calculado["Mes"] = df_calculado["Fecha de pedido"].dt.month
    df_calculado["Día"] = df_calculado["Fecha de pedido"].dt.day

    st.session_state["_df_pipeline"] = df_calculado
    st.session_state["_clave_pipeline"] = clave_actual

# A partir de acá, cada rerun por filtro SOLO lee de session_state — no
# vuelve a leer archivos ni a correr el pipeline.
df = st.session_state["_df_pipeline"]

# Min/max globales sobre el dataframe COMPLETO (sin filtrar), para que el
# slider de más abajo no cambie de rango cada vez que filtras — si el rango
# se recalcula sobre df_f (que cambia con cada filtro), el slider "salta" y
# puede generar reruns en cadena.
NS_MIN_GLOBAL = int(df["Nivel de Servicio"].min()) if len(df) and pd.notna(df["Nivel de Servicio"].min()) else 0
NS_MAX_GLOBAL = int(df["Nivel de Servicio"].max()) if len(df) and pd.notna(df["Nivel de Servicio"].max()) else 100

# Registro de conteos por etapa - para el panel de diagnóstico al final
checkpoints = [("0. Total tras el pipeline (sin filtros)", len(df))]

# Registro de las 3 métricas en cada etapa, para localizar dónde diverge del pbix
metricas_por_etapa = []


def _snapshot(nombre, d):
    """Guarda las 3 métricas del pbix para el subconjunto d en esta etapa."""
    n = len(d)
    pct = (d["Cumple"] == "Cumple").sum() / n * 100 if n else 0
    prom = d["Nivel de Servicio"].mean() if n else float("nan")
    pos_oc = d["Pedido"].nunique() + (1 if d["Pedido"].isna().any() else 0)
    metricas_por_etapa.append(
        (nombre, f"{n:,}", f"{pct:.0f}%", f"{prom:.0f}" if pd.notna(prom) else "-", f"{pos_oc:,}")
    )


_snapshot("0. Sin filtros", df)

# ---- Filtros: Centro y Aplica? ----
st.subheader("Filtros")
c1, c2 = st.columns(2)
with c1:
    centros = st.multiselect("Centro", sorted(df["Centro"].dropna().unique()))
with c2:
    aplica = st.multiselect("Aplica?", sorted(df["Aplica?"].dropna().unique()))

df_f = df.copy()
if centros:
    df_f = df_f[df_f["Centro"].isin(centros)]
checkpoints.append(("1. Tras filtro Centro", len(df_f)))

if aplica:
    df_f = df_f[df_f["Aplica?"].isin(aplica)]
checkpoints.append(("2. Tras filtro Aplica?", len(df_f)))
_snapshot("2. Tras Centro + Aplica?", df_f)

# ---- Estado Solped ----
st.caption("Estado Solped (el filtro de fecha de abajo solo aplica dentro de 'Pedido completo')")
h1, h2, h3, h4 = st.columns(4)
with h1:
    estados = st.multiselect("Estado Solped", sorted(df_f["Estado Solped"].dropna().unique()))

if estados:
    df_f = df_f[df_f["Estado Solped"].isin(estados)]
checkpoints.append(("3. Tras filtro Estado Solped (total)", len(df_f)))
checkpoints.append(("3a. Sin pedido (antes de jerarquía)", (df_f["Estado Solped"] == "Sin pedido").sum()))
checkpoints.append(("3b. Pedido incompleto (antes de jerarquía)", (df_f["Estado Solped"] == "Pedido incompleto").sum()))
checkpoints.append(("3c. Pedido completo (antes de jerarquía)", (df_f["Estado Solped"] == "Pedido completo").sum()))
_snapshot("3. Tras Estado Solped", df_f)

conteo_sin_pedido_antes = (df_f["Estado Solped"] == "Sin pedido").sum()
conteo_incompleto_antes = (df_f["Estado Solped"] == "Pedido incompleto").sum()
conteo_completo_antes = (df_f["Estado Solped"] == "Pedido completo").sum()

# ---- Jerarquía Año / Mes / Día (opciones calculadas solo sobre "Pedido completo") ----
df_pedido_completo = df_f[df_f["Estado Solped"] == "Pedido completo"]

with h2:
    años = st.multiselect("Año", sorted(df_pedido_completo["Año"].dropna().unique().astype(int)))
with h3:
    _base_mes = df_pedido_completo[df_pedido_completo["Año"].isin(años)] if años else df_pedido_completo
    meses = st.multiselect("Mes", sorted(_base_mes["Mes"].dropna().unique().astype(int)))
with h4:
    _base_dia = _base_mes[_base_mes["Mes"].isin(meses)] if meses else _base_mes
    # Fechas completas (no solo el número de día) para no mezclar, por ejemplo,
    # el 30 de julio con el 30 de agosto si se seleccionan ambos meses.
    fechas_disponibles = sorted(_base_dia["Fecha de pedido"].dt.date.dropna().unique())
    fechas = st.multiselect(
        "Día", fechas_disponibles, format_func=lambda f: f.strftime("%d-%m-%Y")
    )

if años or meses or fechas:
    es_pedido_completo = df_f["Estado Solped"] == "Pedido completo"
    cond_fecha = pd.Series(True, index=df_f.index)
    if años:
        cond_fecha &= df_f["Año"].isin(años)
    if meses:
        cond_fecha &= df_f["Mes"].isin(meses)
    if fechas:
        cond_fecha &= df_f["Fecha de pedido"].dt.date.isin(fechas)
    df_f = df_f[~es_pedido_completo | (es_pedido_completo & cond_fecha)]

checkpoints.append(("4a. Sin pedido tras jerarquía (debe ser IGUAL a 3a)", (df_f["Estado Solped"] == "Sin pedido").sum()))
checkpoints.append(("4b. Pedido incompleto tras jerarquía (debe ser IGUAL a 3b)", (df_f["Estado Solped"] == "Pedido incompleto").sum()))
checkpoints.append(("4c. Pedido completo tras jerarquía (debe ser MENOR O IGUAL a 3c)", (df_f["Estado Solped"] == "Pedido completo").sum()))
checkpoints.append(("4. Total tras jerarquía de fecha", len(df_f)))
_snapshot("4. Tras jerarquía Año/Mes/Día", df_f)

# ---- Solped MRP y Cumple ----
c3, c4 = st.columns(2)
with c3:
    solped_mrp = st.multiselect("Solped MRP", sorted(df_f["Solped MRP"].dropna().unique()))
with c4:
    cumple = st.multiselect("Nivel de Servicio (Cumple)", sorted(df_f["Cumple"].dropna().unique()))

if solped_mrp:
    df_f = df_f[df_f["Solped MRP"].isin(solped_mrp)]
checkpoints.append(("5. Tras filtro Solped MRP", len(df_f)))
_snapshot("5. Tras Solped MRP", df_f)

if cumple:
    df_f = df_f[df_f["Cumple"].isin(cumple)]
checkpoints.append(("6. Tras filtro Cumple", len(df_f)))
_snapshot("6. Tras Cumple", df_f)

# ---- Segunda capa: filtro por rango de Nivel de Servicio (días) ----
# Los valores negativos aparecen cuando la solped se modificó DESPUÉS de que
# ya existía la OC (Fecha modificación > Fecha de pedido). No representan
# gestión real y, por la regla de Cumple (<=10 o <=7), siempre puntúan como
# "Cumple", lo que infla el indicador.
st.divider()
st.subheader("Filtro por días de gestión (Nivel de Servicio)")

if len(df_f):
    n_negativos = (df_f["Nivel de Servicio"] < 0).sum()

    f1, f2 = st.columns([1, 2])
    with f1:
        excluir_negativos = st.checkbox(
            "Excluir negativos (solo desde 0)",
            value=False,
            help=(
                f"Hay {n_negativos:,} filas con días negativos en la selección actual. "
                "Ocurren cuando la solped se modificó después de generada la OC. "
                "Al excluirlas el promedio deja de estar distorsionado."
            ),
        )
    with f2:
        if excluir_negativos:
            rango_ns = (0, NS_MAX_GLOBAL)
            st.caption(f"Rango aplicado: 0 a {NS_MAX_GLOBAL:,} días (negativos excluidos)")
        elif NS_MIN_GLOBAL < NS_MAX_GLOBAL:
            # Rango y key fijos (basados en NS_MIN/MAX_GLOBAL, no en df_f) para
            # que el slider no se redefina en cada rerun con límites distintos.
            rango_ns = st.slider(
                "Rango de días de gestión",
                min_value=NS_MIN_GLOBAL,
                max_value=NS_MAX_GLOBAL,
                value=(NS_MIN_GLOBAL, NS_MAX_GLOBAL),
                key="slider_rango_dias",
            )
        else:
            rango_ns = (NS_MIN_GLOBAL, NS_MAX_GLOBAL)
            st.caption(f"Todas las filas tienen {NS_MIN_GLOBAL} días")

    df_f = df_f[df_f["Nivel de Servicio"].between(rango_ns[0], rango_ns[1])]

    if n_negativos and not excluir_negativos:
        st.caption(
            f"⚠️ La selección incluye {n_negativos:,} filas con días negativos, "
            "que siempre cuentan como 'Cumple' y bajan el promedio."
        )

checkpoints.append(("7. Tras filtro Nivel de Servicio (RESULTADO FINAL)", len(df_f)))
_snapshot("7. RESULTADO FINAL", df_f)

# ---- Panel de diagnóstico ----
with st.expander("🔍 Diagnóstico de filtrado (para comparar contra el pbix)"):
    st.write(
        "Filas en cada etapa. Los pasos 4a y 4b deben quedar IGUAL al paso 3 "
        "(la jerarquía de fecha no debe tocar 'Sin pedido' ni 'Pedido incompleto'). "
        "Si en el pbix ves un número distinto, compara etapa por etapa hasta encontrar "
        "en cuál empiezan a diferir."
    )
    st.table(pd.DataFrame(checkpoints, columns=["Etapa", "Filas"]))

    st.write("**Las 3 métricas en cada etapa del filtro** — compara contra el pbix aplicando los mismos slicers, uno a uno, para ver en cuál se separan:")
    st.table(pd.DataFrame(metricas_por_etapa, columns=["Etapa", "Filas", "% Cumplimiento", "Prom. días", "Pos. OC"]))

    st.write("Detalle de las solicitudes en el resultado final, para cruzar 1 a 1 contra el export del pbix:")
    st.dataframe(
        df_f[["Solicitud de pedido", "Centro", "Estado Solped", "Fecha de pedido", "Nivel de Servicio", "Cumple", "Solped MRP"]]
        .sort_values("Solicitud de pedido"),
        use_container_width=True,
    )
    csv = df_f.to_csv(index=False).encode("utf-8")
    st.download_button("Descargar detalle filtrado (CSV)", csv, "detalle_filtrado.csv", "text/csv")

# ---- Tarjetas con semáforo ----
VERDE = "rgba(35, 145, 75, 0.16)"
VERDE_BORDE = "rgba(35, 145, 75, 0.55)"
ROJO = "rgba(204, 0, 0, 0.14)"
ROJO_BORDE = "rgba(204, 0, 0, 0.55)"

pct_cumplimiento = (df_f["Cumple"] == "Cumple").sum() / max(len(df_f), 1) * 100
promedio_dias = df_f["Nivel de Servicio"].mean()
pedidos_distintos = df_f["Pedido"].nunique() + (1 if df_f["Pedido"].isna().any() else 0)


def tarjeta(titulo: str, valor: str, fondo: str = "rgba(64,75,85,0.07)", borde: str = "rgba(64,75,85,0.35)") -> str:
    return (
        f'<div style="background:{fondo};border:1.5px solid {borde};border-radius:8px;'
        f'padding:14px 18px;text-align:center;">'
        f'<div style="font-size:0.78rem;color:#404B55;font-weight:600;letter-spacing:.03em;'
        f'text-transform:uppercase;opacity:.85;margin-bottom:4px;">{titulo}</div>'
        f'<div style="font-size:2rem;font-weight:700;color:#404B55;line-height:1.1;">{valor}</div>'
        f"</div>"
    )


# Semáforo: días de gestión <= 10 verde, > 10 rojo
if pd.isna(promedio_dias):
    f_dias, b_dias, txt_dias = "rgba(64,75,85,0.07)", "rgba(64,75,85,0.35)", "-"
elif promedio_dias > 10:
    f_dias, b_dias, txt_dias = ROJO, ROJO_BORDE, f"{promedio_dias:.0f}"
else:
    f_dias, b_dias, txt_dias = VERDE, VERDE_BORDE, f"{promedio_dias:.0f}"

# Semáforo: % cumplimiento >= 85 verde, < 85 rojo
if pct_cumplimiento >= 85:
    f_pct, b_pct = VERDE, VERDE_BORDE
else:
    f_pct, b_pct = ROJO, ROJO_BORDE

t1, t2, t3 = st.columns(3)
with t1:
    st.markdown(tarjeta("Promedio días de gestión", txt_dias, f_dias, b_dias), unsafe_allow_html=True)
with t2:
    st.markdown(tarjeta("% Cumplimiento", f"{pct_cumplimiento:.0f}%", f_pct, b_pct), unsafe_allow_html=True)
with t3:
    st.markdown(tarjeta("OC generadas", f"{pedidos_distintos:,}"), unsafe_allow_html=True)

st.divider()

# ---- Estilo corporativo Enaex para las tablas ----
ENAEX_GRIS = "#404B55"
ENAEX_ROJO = "#CC0000"


def tabla_enaex(tabla: pd.DataFrame, max_height: int | None = None, compacta: bool = False) -> str:
    """
    Renderiza la tabla como HTML con la paleta corporativa: encabezado gris
    oscuro, filas alternadas, y la fila TOTAL destacada en rojo.
    max_height (px) activa el scroll vertical con encabezado fijo.
    """
    cols = list(tabla.columns)

    def _fmt(col, val):
        if pd.isna(val):
            return "-"
        if col == "Promedio días de gestión":
            return f"{val:,.0f}"
        if col == "% Cumplimiento":
            return f"{val:,.0f}%"
        if col == "Pos. OC generadas":
            return f"{val:,.0f}"
        return str(val)

    pad = "4px 6px" if compacta else "7px 12px"
    pad_th = "5px 6px" if compacta else "9px 12px"
    fuente = "0.72rem" if compacta else "0.86rem"
    fuente_th = "0.66rem" if compacta else "0.82rem"

    # OPTIMIZACIÓN: itertuples() en vez de iterrows(). Con listas de miles de
    # filas, iterrows() reconstruye una Series por fila (~30x más lento acá
    # en benchmark local); itertuples() entrega tuplas planas, mucho más rápido
    # para solo leer valores como hacemos en este loop.
    filas = []
    for i, r in enumerate(tabla.itertuples(index=False)):
        es_total = str(r[0]) == "TOTAL"
        if es_total:
            estilo_fila = f"background:{ENAEX_GRIS};color:#fff;font-weight:700;border-top:2px solid {ENAEX_ROJO};"
        else:
            fondo = "#ffffff" if i % 2 == 0 else "#f4f5f7"
            estilo_fila = f"background:{fondo};color:{ENAEX_GRIS};"
        celdas = []
        for j, c in enumerate(cols):
            align = "left" if j == 0 else "right"
            celdas.append(
                f'<td style="padding:{pad};text-align:{align};'
                f'border-bottom:1px solid #e3e5e8;white-space:nowrap;">{_fmt(c, r[j])}</td>'
            )
        filas.append(f'<tr style="{estilo_fila}">{"".join(celdas)}</tr>')

    # En modo compacto los títulos largos se abrevian: son los que hacen que la
    # tabla se salga del recuadro cuando va en media pantalla.
    abrev = {
        "Promedio días de gestión": "Días gest.",
        "% Cumplimiento": "% Cumpl.",
        "Pos. OC generadas": "Pos. OC",
    }
    encabezados = "".join(
        f'<th style="padding:{pad_th};text-align:{"left" if j == 0 else "right"};'
        f'background:{ENAEX_GRIS};color:#fff;font-weight:600;font-size:{fuente_th};'
        f'letter-spacing:.02em;position:sticky;top:0;z-index:2;white-space:nowrap;">'
        f"{abrev.get(c, c) if compacta else c}</th>"
        for j, c in enumerate(cols)
    )

    tabla_html = (
        f'<table style="width:100%;min-width:{"340px" if compacta else "auto"};border-collapse:collapse;font-size:{fuente};'
        f'font-family:inherit;border:1px solid #d8dbdf;">'
        f"<thead><tr>{encabezados}</tr></thead><tbody>{''.join(filas)}</tbody></table>"
    )

    # Scroll horizontal siempre disponible (para tablas angostas en dos columnas),
    # y vertical solo si se pidió max_height.
    alto = f"max-height:{max_height}px;overflow-y:auto;" if max_height else ""
    return (
        f'<div style="{alto}overflow-x:auto;border:1px solid #d8dbdf;'
        f'border-radius:4px;">{tabla_html}</div>'
    )


# ---- Tabla por Comprador ----
st.subheader("Por comprador")
st.caption(
    "Asignación por **grupo de compras**: las líneas MRP se reparten entre los "
    "compradores responsables de cada grupo, en vez de concentrarse en el responsable de MRP."
)
col_comprador = (
    "Comprador (Grupo de compras)"
    if "Comprador (Grupo de compras)" in df_f.columns
    else "Comprador por Grupo Compras"
)
tabla_comprador = transform.calcular_metricas_por_grupo(df_f, [col_comprador])
tabla_comprador = transform.agregar_fila_total(tabla_comprador, df_f, [col_comprador])
st.markdown(tabla_enaex(tabla_comprador), unsafe_allow_html=True)

st.write("")

# ---- Dos vistas por centro, en paralelo ----
vc1, vc2 = st.columns(2)

with vc1:
    st.subheader("Por centro logístico")
    st.caption("Vista fija — el total calza con la vista por comprador.")
    tabla_fija = transform.tabla_centros_fija(df_f)
    st.markdown(tabla_enaex(tabla_fija, compacta=True), unsafe_allow_html=True)

with vc2:
    st.subheader("Detalle por centro")
    st.caption("Centros activos según los filtros aplicados.")
    cols_detalle = [c for c in ["Centro", "Nombre Centro 2"] if c in df_f.columns]
    tabla_detalle = transform.calcular_metricas_por_grupo(df_f, cols_detalle)
    tabla_detalle = tabla_detalle.sort_values("Pos. OC generadas", ascending=False)
    tabla_detalle = transform.agregar_fila_total(tabla_detalle, df_f, cols_detalle)
    st.markdown(tabla_enaex(tabla_detalle, max_height=300, compacta=True), unsafe_allow_html=True)

st.divider()

# ---- Detalle de solicitudes con columna de comentario editable ----
COLUMNAS_DETALLE = [
    "Centro",
    "Material",
    "Texto breve",
    "Solicitud de pedido",
    "Fecha de solicitud",
    "Fecha modificación",
    "Grupo de compras",
    "Cantidad pedida",
    "Pedido",
    "Fecha de pedido",
    "Posición de pedido",
    "Comprador (Grupo de compras)",
    "Solped MRP",
    "Nombre Centro 2",
    "Nombre Centro",
    "Estado Solped",
    "Nivel de Servicio",
    "Comentario",
    "Cumple",
]


def preparar_detalle(d: pd.DataFrame) -> pd.DataFrame:
    """Ordena columnas, quita las auxiliares y deja las fechas sin hora."""
    d = d.copy()
    for col_fecha in ["Fecha de solicitud", "Fecha modificación", "Fecha de pedido"]:
        if col_fecha in d.columns:
            d[col_fecha] = pd.to_datetime(d[col_fecha], errors="coerce").dt.date
    if "Comentario" not in d.columns:
        d["Comentario"] = ""
    return d[[c for c in COLUMNAS_DETALLE if c in d.columns]]


detalle = preparar_detalle(df_f)

with st.expander("Ver detalle de solicitudes", expanded=False):
    st.caption("La columna **Comentario** es editable: escribe ahí y se incluirá en el Excel de registro.")
    detalle_editado = st.data_editor(
        detalle,
        use_container_width=True,
        num_rows="fixed",
        key="editor_detalle",
        column_config={
            "Comentario": st.column_config.TextColumn(
                "Comentario", help="Anotación libre para el registro semanal", width="medium"
            )
        },
        disabled=[c for c in detalle.columns if c != "Comentario"],
    )

# ---- Exportación del registro semanal a Excel ----
semana_ref = pd.Timestamp(fecha_corte) - pd.Timedelta(days=7)
num_semana = semana_ref.isocalendar()[1]
nombre_archivo = f"Sem{num_semana:02d}-{semana_ref.year}.xlsx"


def generar_excel(detalle_df: pd.DataFrame) -> bytes:
    """Arma el registro semanal: resumen, tablas por comprador/centro y detalle."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    gris = "FF404B55"
    rojo = "FFCC0000"
    fuente_base = "Arial"

    wb = Workbook()
    borde = Border(bottom=Side(style="thin", color="FFD8DBDF"))

    def escribir_hoja(ws, titulo, tabla, col_inicio=1, fila_inicio=1):
        ws.cell(row=fila_inicio, column=col_inicio, value=titulo).font = Font(
            name=fuente_base, bold=True, size=12, color=gris
        )
        fila = fila_inicio + 1
        for j, col in enumerate(tabla.columns):
            c = ws.cell(row=fila, column=col_inicio + j, value=str(col))
            c.font = Font(name=fuente_base, bold=True, color="FFFFFFFF", size=10)
            c.fill = PatternFill("solid", fgColor=gris)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # OPTIMIZACIÓN: itertuples() en vez de iterrows() — mismo motivo que en
        # tabla_enaex. Acá además necesitamos los nombres de columna, así que
        # usamos itertuples(name=None) para tuplas planas y las emparejamos con
        # tabla.columns por índice.
        for i, r in enumerate(tabla.itertuples(index=False, name=None)):
            es_total = str(r[0]) == "TOTAL"
            for j, col in enumerate(tabla.columns):
                val = r[j]
                if pd.isna(val):
                    val = None
                elif isinstance(val, (int, float)) and col in (
                    "Promedio días de gestión",
                    "% Cumplimiento",
                    "Pos. OC generadas",
                ):
                    val = round(float(val))
                elif hasattr(val, "item"):
                    # numpy.int64 / numpy.float64 -> tipo nativo de Python
                    val = val.item()
                c = ws.cell(row=fila + 1 + i, column=col_inicio + j, value=val)
                c.font = Font(name=fuente_base, size=10, bold=es_total, color=gris)
                c.border = borde
                if es_total:
                    c.fill = PatternFill("solid", fgColor="FFEFF0F2")
                if col == "% Cumplimiento" and val is not None:
                    c.number_format = '0"%"'
                if col in ("Fecha de solicitud", "Fecha modificación", "Fecha de pedido"):
                    c.number_format = "DD-MM-YYYY"
        return fila + 1 + len(tabla)

    def ajustar_ancho(ws, tabla, col_inicio=1, extra=3):
        for j, col in enumerate(tabla.columns):
            # Se recorre en Python puro: los métodos de pandas (fillna/replace)
            # fallan o cambian dtypes en columnas Int64 anulables con vacíos.
            largos = [len(str(col))]
            for v in tabla[col].head(200):
                largos.append(0 if pd.isna(v) else len(str(v)))
            ws.column_dimensions[get_column_letter(col_inicio + j)].width = min(max(largos) + extra, 45)

    # --- Hoja 1: Resumen ---
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = f"Nivel de Servicio MRO — Registro semana {num_semana:02d}/{semana_ref.year}"
    ws["A1"].font = Font(name=fuente_base, bold=True, size=14, color=gris)
    ws["A2"] = f"Fecha de corte del reporte: {pd.Timestamp(fecha_corte).date()}"
    ws["A2"].font = Font(name=fuente_base, size=10, italic=True, color=gris)
    ws["A3"] = f"Generado: {pd.Timestamp.today().date()}"
    ws["A3"].font = Font(name=fuente_base, size=10, italic=True, color=gris)

    resumen = pd.DataFrame(
        {
            "Indicador": ["Promedio días de gestión", "% Cumplimiento", "OC generadas", "Líneas consideradas"],
            "Valor": [
                round(promedio_dias) if pd.notna(promedio_dias) else None,
                round(pct_cumplimiento),
                pedidos_distintos,
                len(df_f),
            ],
        }
    )
    fila = escribir_hoja(ws, "Indicadores generales", resumen, fila_inicio=5)
    ws.cell(row=5, column=1).font = Font(name=fuente_base, bold=True, size=12, color=rojo)
    ajustar_ancho(ws, resumen)

    fila = escribir_hoja(ws, "Por comprador", tabla_comprador, fila_inicio=fila + 2)
    fila = escribir_hoja(ws, "Por centro logístico", tabla_fija, fila_inicio=fila + 2)
    escribir_hoja(ws, "Detalle por centro", tabla_detalle, fila_inicio=fila + 2)

    # --- Hoja 2: Detalle de solicitudes (con comentarios) ---
    ws2 = wb.create_sheet("Detalle solicitudes")
    escribir_hoja(ws2, "Detalle de solicitudes", detalle_df)
    ajustar_ancho(ws2, detalle_df)
    ws2.freeze_panes = "A3"

    from io import BytesIO

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


st.subheader("Registro semanal")
st.caption(f"Prepara el reporte completo (indicadores, tablas y detalle con comentarios) como **{nombre_archivo}**.")

# Patrón de dos pasos: generar_excel() SOLO corre cuando el usuario hace clic
# en "Preparar", no en cada rerun. Escribir un Excel celda por celda con
# openpyxl es lento a escala (varios segundos con miles de filas de detalle)
# — si esto corriera en cada rerun (como antes), cada filtro que tocas
# dispararía esa reconstrucción completa aunque no fueras a descargar nada.
# Guardamos los bytes ya generados en session_state, listos para servir.
clave_excel = (len(df_f), int(pd.util.hash_pandas_object(detalle_editado["Comentario"].fillna("")).sum()), pd.Timestamp(fecha_corte))

col_prep, col_desc = st.columns([1, 2])
with col_prep:
    if st.button("📄 Preparar Excel"):
        with st.spinner("Generando el Excel..."):
            try:
                st.session_state["_excel_bytes"] = generar_excel(detalle_editado)
                st.session_state["_excel_clave"] = clave_excel
            except Exception as e:
                st.error(f"No se pudo generar el Excel: {e}")

with col_desc:
    excel_listo = st.session_state.get("_excel_bytes") is not None
    excel_desactualizado = st.session_state.get("_excel_clave") != clave_excel
    if excel_listo and excel_desactualizado:
        st.caption("⚠️ Los filtros cambiaron desde que preparaste este Excel — vuelve a preparar para reflejar la selección actual.")
    if excel_listo:
        st.download_button(
            f"⬇ Descargar {nombre_archivo}",
            data=st.session_state["_excel_bytes"],
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )
    else:
        st.caption("Haz clic en \"Preparar Excel\" para generar el archivo de descarga.")
