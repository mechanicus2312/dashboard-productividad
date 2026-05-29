import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="Dashboard Productividad",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stMetric"] {
    background: white;
    border-radius: 10px;
    padding: 1rem 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
[data-testid="stMetricValue"] { color: #003366; font-weight: 700; }
.header-box {
    background: linear-gradient(90deg, #003366, #0077CC);
    color: white;
    padding: 1.2rem 2rem;
    border-radius: 10px;
    margin-bottom: 1.2rem;
}
</style>
""", unsafe_allow_html=True)

SEMAFORO_COLOR = {
    "Verde":    "#00B050",
    "Amarillo": "#FFBF00",
    "Alerta":   "#E74C3C",
    "Rojo":     "#C0392B",
}
PALETA = ["#003366", "#0077CC", "#66B2FF", "#FF6B35", "#FFD166"]

META_CIRCULAR_DEF = 0.86
META_FTE_DEF      = 115.83
DIAS_HAB_DEF      = 104



# ── Helpers ───────────────────────────────────────────────────────────────────
def _calcular_semaforo(pct):
    if pct >= 0.95:
        return "Verde"
    elif pct >= META_CIRCULAR_DEF:
        return "Amarillo"
    return "Alerta"


def _construir_resumenes(det):
    # Resumen por banca y línea
    banca = (
        det.groupby(["Banca_Transformada", "Linea_Proceso"])
        .agg(
            Total_Ops=("Cantidad_Solicitudes", "sum"),
            Cumple=("Cumplimiento", lambda x: (x == "CUMPLE").sum()),
            Min_Ejecucion=("Min_Ejecucion", "sum"),
        )
        .reset_index()
    )
    banca["Pct_Cumplimiento"] = banca["Cumple"] / banca["Total_Ops"]

    # Resumen por funcionario
    func = (
        det.groupby(["Funcionario_Ejecutor", "Banca_Transformada"])
        .agg(
            Operaciones_Asignadas=("Cantidad_Solicitudes", "sum"),
            Operaciones_Ejecutadas=("Cantidad_Solicitudes", "sum"),
            Min_Ejecucion_Total=("Min_Ejecucion", "sum"),
            Min_Asignacion_Total=("Min_Asignacion", "sum"),
            Cumple=("Cumplimiento", lambda x: (x == "CUMPLE").sum()),
            No_Cumple=("Cumplimiento", lambda x: (x != "CUMPLE").sum()),
            Ultima_Actividad=("Fecha_Fin", "max"),
        )
        .reset_index()
    )
    func["Pct_Cumplimiento"] = func["Cumple"] / func["Operaciones_Ejecutadas"]
    func["Eficiencia"] = func["Min_Ejecucion_Total"] / func["Min_Asignacion_Total"]
    func["Semaforo"] = func["Pct_Cumplimiento"].apply(_calcular_semaforo)
    func["Productividad_Estandarizada"] = (
        func["Operaciones_Ejecutadas"] / META_FTE_DEF
    ).round(4)
    func["Ranking"] = func["Operaciones_Ejecutadas"].rank(ascending=False).astype(int)
    func["Eficacia"] = 1.0
    func["Efectividad"] = func["Operaciones_Ejecutadas"] / (func["Operaciones_Ejecutadas"].mean())

    params = {
        "Circular Reglamentaria":        META_CIRCULAR_DEF,
        "Meta Teorica x FTE (16 dias)":  META_FTE_DEF,
        "Días Hábiles Periodo":          DIAS_HAB_DEF,
    }
    return func, banca, params


# ── Carga desde Excel (respaldo) ──────────────────────────────────────────────
def _parsear_excel(buf):
    det = pd.read_excel(buf, sheet_name="Detalle",
                        parse_dates=["Fecha_Llegada", "Fecha_Asignacion", "Fecha_Fin"])
    buf.seek(0)
    func = pd.read_excel(buf, sheet_name="Resumen_Funcionario",
                         parse_dates=["Ultima_Actividad"])
    buf.seek(0)
    banca = pd.read_excel(buf, sheet_name="Resumen_Banca")
    buf.seek(0)
    raw_params = pd.read_excel(buf, sheet_name="Parametros")

    num_func = ["Eficacia", "Efectividad", "Eficiencia",
                "Productividad", "Productividad_Estandarizada", "Pct_Cumplimiento"]
    for c in num_func:
        if func[c].dtype == object:
            func[c] = func[c].str.replace(",", ".", regex=False).astype(float)
    if banca["Pct_Cumplimiento"].dtype == object:
        banca["Pct_Cumplimiento"] = (
            banca["Pct_Cumplimiento"].str.replace(",", ".", regex=False).astype(float)
        )
    p = dict(zip(raw_params["Parametro"], raw_params["Valor"]))
    return det, func, banca, p


def _tiempo_a_minutos(serie):
    """Convierte 'HH:MM' a minutos desde medianoche."""
    def _conv(val):
        try:
            partes = str(val).strip().split(":")
            if len(partes) == 2:
                return int(partes[0]) * 60 + int(partes[1])
        except Exception:
            pass
        return None
    return serie.apply(_conv)


@st.cache_data
def cargar_desde_csv(archivo_bytes):
    for enc in ["utf-8-sig", "latin-1", "cp1252"]:
        try:
            raw = pd.read_csv(BytesIO(archivo_bytes), encoding=enc, low_memory=False)
            break
        except UnicodeDecodeError:
            continue

    det = pd.DataFrame()

    # ── Identificadores y categorías ─────────────────────────────────────────
    det["ID"]                 = raw.get("Consecutivo", raw.get("ID"))
    def _normalizar_banca(val):
        v = str(val).upper().strip()
        if any(k in v for k in ["PERSONA", "BP", "BANCA P"]):
            return "Personas"
        if any(k in v for k in ["EMPRESA", "BE", "BANCA E"]):
            return "Empresas"
        return val  # deja el valor original si no coincide

    col_banca = "Banca_Transformada" if "Banca_Transformada" in raw.columns else "Banca"
    det["Banca_Transformada"] = raw.get(col_banca, pd.Series(dtype=str)).apply(_normalizar_banca)
    det["Linea_Proceso"]      = raw.get("Linea de Proceso", raw.get("Linea_Proceso"))
    det["Funcionario_Asignado"] = raw.get("Funcionario asignado", raw.get("Funcionario_Asignado"))
    det["Funcionario_Ejecutor"] = raw.get("Funcionario Ejecutor", raw.get("Funcionario_Ejecutor"))
    det["Estado_Solicitud"]   = raw.get("Estado de la solicitud", raw.get("Estado_Solicitud"))
    det["Cargo_Comercial"]    = raw.get("Cargo Comercial", raw.get("Cargo_Comercial"))
    det["Segmento"]           = raw.get("Segmento", "")
    det["Actividad"]          = raw.get("Actividad realizada", raw.get("Actividad", ""))
    det["Canal"]              = raw.get("Servicio/Canal", raw.get("Canal", ""))
    det["Cantidad_Solicitudes"] = pd.to_numeric(
        raw.get("Cantidad Solicitudes", raw.get("Cantidad_Solicitudes", 1)), errors="coerce"
    ).fillna(1).astype(int)

    # ── Fechas ────────────────────────────────────────────────────────────────
    for col_dest, col_src in [
        ("Fecha_Llegada",    "Fecha llegada correo"),
        ("Fecha_Asignacion", "Fecha asignacion"),
        ("Fecha_Fin",        "Fecha fin"),
    ]:
        det[col_dest] = pd.to_datetime(
            raw.get(col_src, raw.get(col_dest)), dayfirst=True, errors="coerce"
        )

    # ── Tiempos en minutos ────────────────────────────────────────────────────
    min_llegada   = _tiempo_a_minutos(raw.get("Hora llegada correo",   pd.Series()))
    min_asignado  = _tiempo_a_minutos(raw.get("Hora asignacion",       pd.Series()))
    min_ini_ejec  = _tiempo_a_minutos(raw.get("Hora inicio ejecucion", pd.Series()))
    min_fin_ejec  = _tiempo_a_minutos(raw.get("Hora fin ejecucion",    pd.Series()))

    det["Min_Asignacion"] = (min_asignado - min_llegada).clip(lower=0)

    raw_ejec = pd.to_numeric(raw.get("Min_Fin_Ejecucion"), errors="coerce") - \
               pd.to_numeric(raw.get("Min_Inicio_Ejecucion"), errors="coerce")
    # Calcular desde horas si están disponibles; si no, usar columnas Min_*
    ejec_horas = (min_fin_ejec - min_ini_ejec)
    det["Min_Ejecucion"] = ejec_horas.where(ejec_horas.notna(), raw_ejec).clip(lower=0)

    det["Min_Espera_Asignacion"] = (det["Min_Asignacion"] - det["Min_Ejecucion"]).clip(lower=0)
    det["Min_Ciclo"]             = det["Min_Asignacion"] + det["Min_Ejecucion"]
    det["Min_Ciclo_Ajustado"]    = det["Min_Ciclo"]

    # ── Cumplimiento (umbral configurable, default 480 min = 1 día hábil) ────
    umbral = st.session_state.get("umbral_ciclo", 480)
    det["Cumplimiento"]          = det["Min_Ciclo"].apply(
        lambda x: "CUMPLE" if pd.notna(x) and x <= umbral else "NO CUMPLE"
    )
    det["Cumplimiento_Ajustado"] = det["Cumplimiento"]

    # ── Dimensiones temporales ────────────────────────────────────────────────
    det["Año"]       = det["Fecha_Llegada"].dt.year
    det["Mes"]       = det["Fecha_Llegada"].dt.month
    det["Semana"]    = det["Fecha_Llegada"].dt.isocalendar().week.astype("Int64")
    det["Mes_Nombre"] = det["Fecha_Llegada"].dt.strftime("%B %Y")

    func, banca, params = _construir_resumenes(det)
    return det, func, banca, params


@st.cache_data
def cargar_desde_xlsx(archivo_bytes):
    return _parsear_excel(BytesIO(archivo_bytes))


# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 📂 Datos")

archivo = st.sidebar.file_uploader(
    "Sube el archivo de datos",
    type=["csv", "xlsx"],
    help="CSV exportado de SharePoint o Excel con las 4 hojas procesadas",
)

if archivo is None:
    st.markdown(
        '<div class="header-box"><h2 style="margin:0">📊 Dashboard de Productividad</h2></div>',
        unsafe_allow_html=True,
    )
    st.markdown("### Para comenzar, sube el archivo desde el panel izquierdo.")
    st.info(
        "**Opción 1 — CSV de SharePoint** *(recomendado)*\n"
        "1. Entra a la lista en SharePoint\n"
        "2. Exporta a Excel → descarga el `.csv`\n"
        "3. Súbelo aquí\n\n"
        "**Opción 2 — Excel procesado** `.xlsx` con hojas: "
        "`Detalle`, `Resumen_Funcionario`, `Resumen_Banca`, `Parametros`"
    )
    st.stop()

if archivo.name.endswith(".csv"):
    st.session_state["umbral_ciclo"] = st.sidebar.number_input(
        "Umbral cumplimiento (min)", min_value=30, max_value=2880,
        value=st.session_state.get("umbral_ciclo", 480),
        step=30,
        help="Minutos máximos de ciclo para considerar CUMPLE. Default: 480 min (1 día hábil)",
    )

archivo_bytes = archivo.read()
if archivo.name.endswith(".csv"):
    det, func, banca, params = cargar_desde_csv(archivo_bytes)
else:
    det, func, banca, params = cargar_desde_xlsx(archivo_bytes)

st.sidebar.success(f"✅ {archivo.name}")

META_CIRCULAR = float(params.get("Circular Reglamentaria", 0.86))
META_FTE      = float(params.get("Meta Teorica x FTE (16 dias)", 115.83))
DIAS_HABILES  = int(params.get("Días Hábiles Periodo", 104))

st.sidebar.divider()

# ── Filtros ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("## ⚙️ Filtros")
bancas_opts = ["Todas"] + sorted(det["Banca_Transformada"].dropna().unique())
banca_sel = st.sidebar.selectbox("Banca", bancas_opts)

lineas_opts = ["Todas"] + sorted(det["Linea_Proceso"].dropna().unique())
linea_sel = st.sidebar.selectbox("Línea de proceso", lineas_opts)

st.sidebar.divider()
pagina = st.sidebar.radio(
    "Página",
    ["📊 Resumen Ejecutivo", "🏆 Ranking Funcionarios", "🔍 Detalle por Línea"],
)
st.sidebar.divider()
st.sidebar.caption(
    f"Meta Circular: **{META_CIRCULAR:.0%}** | Meta FTE: **{META_FTE:.1f}** | "
    f"Días hábiles: **{DIAS_HABILES}**"
)


# ── Filtrado ──────────────────────────────────────────────────────────────────
def filtrar_det(df):
    d = df.copy()
    if banca_sel != "Todas":
        d = d[d["Banca_Transformada"] == banca_sel]
    if linea_sel != "Todas":
        d = d[d["Linea_Proceso"] == linea_sel]
    return d

def filtrar_func(df):
    d = df.copy()
    if banca_sel != "Todas":
        d = d[d["Banca_Transformada"] == banca_sel]
    return d

def filtrar_banca(df):
    d = df.copy()
    if banca_sel != "Todas":
        d = d[d["Banca_Transformada"] == banca_sel]
    if linea_sel != "Todas":
        d = d[d["Linea_Proceso"] == linea_sel]
    return d


det_f   = filtrar_det(det)
func_f  = filtrar_func(func)
banca_f = filtrar_banca(banca)

# ── KPIs globales ─────────────────────────────────────────────────────────────
total_ops     = int(det_f["Cantidad_Solicitudes"].sum())
ops_cumple    = int(det_f.loc[det_f["Cumplimiento"] == "CUMPLE", "Cantidad_Solicitudes"].sum())
pct_cumple    = ops_cumple / total_ops if total_ops else 0
n_func        = det_f["Funcionario_Ejecutor"].nunique()
avg_ejecucion = det_f["Min_Ejecucion"].mean()
avg_ciclo     = det_f["Min_Ciclo"].mean()


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — RESUMEN EJECUTIVO
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "📊 Resumen Ejecutivo":
    st.markdown(
        '<div class="header-box"><h2 style="margin:0">📊 Resumen Ejecutivo de Productividad</h2></div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Operaciones", f"{total_ops:,}")
    c2.metric(
        "% Cumplimiento",
        f"{pct_cumple:.1%}",
        f"{pct_cumple - META_CIRCULAR:+.1%} vs meta",
    )
    c3.metric("Funcionarios Activos", n_func)
    c4.metric("Avg. Min Ejecución", f"{avg_ejecucion:.1f} min")
    c5.metric("Avg. Min Ciclo", f"{avg_ciclo:.1f} min")

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Cumplimiento por Banca y Línea")
        fig = px.bar(
            banca_f.sort_values("Pct_Cumplimiento"),
            x="Pct_Cumplimiento", y="Linea_Proceso",
            color="Banca_Transformada", barmode="group",
            orientation="h", text_auto=".1%",
            labels={"Pct_Cumplimiento": "% Cumplimiento", "Linea_Proceso": ""},
            color_discrete_sequence=PALETA,
        )
        fig.add_vline(
            x=META_CIRCULAR, line_dash="dash", line_color="#E74C3C",
            annotation_text=f"Meta {META_CIRCULAR:.0%}",
            annotation_position="top right",
        )
        fig.update_layout(height=400, plot_bgcolor="white", xaxis_tickformat=".0%",
                          legend_title="Banca")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Distribución por Semáforo")
        sem = func_f["Semaforo"].value_counts().reset_index()
        sem.columns = ["Semaforo", "Cantidad"]
        fig2 = px.pie(
            sem, values="Cantidad", names="Semaforo", hole=0.5,
            color="Semaforo", color_discrete_map=SEMAFORO_COLOR,
        )
        fig2.update_traces(textinfo="percent+label")
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Tendencia Semanal de Cumplimiento")
        semanal = (
            det_f.groupby(["Semana", "Banca_Transformada"])
            .agg(
                Ops=("Cantidad_Solicitudes", "sum"),
                Cumple=("Cumplimiento", lambda x: (x == "CUMPLE").sum()),
            )
            .reset_index()
        )
        semanal["Pct"] = semanal["Cumple"] / semanal["Ops"]
        fig3 = px.line(
            semanal, x="Semana", y="Pct",
            color="Banca_Transformada", markers=True,
            labels={"Pct": "% Cumplimiento", "Semana": "Semana"},
            color_discrete_sequence=PALETA,
        )
        fig3.add_hline(
            y=META_CIRCULAR, line_dash="dash", line_color="#E74C3C",
            annotation_text=f"Meta {META_CIRCULAR:.0%}",
        )
        fig3.update_layout(height=330, plot_bgcolor="white", yaxis_tickformat=".0%")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("Operaciones por Línea de Proceso")
        ops_linea = (
            det_f.groupby("Linea_Proceso")["Cantidad_Solicitudes"]
            .sum().reset_index()
            .sort_values("Cantidad_Solicitudes")
        )
        fig4 = px.bar(
            ops_linea, x="Cantidad_Solicitudes", y="Linea_Proceso",
            orientation="h", text_auto=True,
            labels={"Cantidad_Solicitudes": "Operaciones", "Linea_Proceso": ""},
            color_discrete_sequence=["#003366"],
        )
        fig4.update_layout(height=330, plot_bgcolor="white")
        st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — RANKING FUNCIONARIOS
# ══════════════════════════════════════════════════════════════════════════════
elif pagina == "🏆 Ranking Funcionarios":
    st.markdown(
        '<div class="header-box"><h2 style="margin:0">🏆 Ranking de Funcionarios</h2></div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Top 15 por Productividad Estandarizada")
        top15 = func_f.nlargest(15, "Productividad_Estandarizada").sort_values(
            "Productividad_Estandarizada"
        )
        colores_bar = [SEMAFORO_COLOR.get(s, "#999") for s in top15["Semaforo"]]
        fig = go.Figure(
            go.Bar(
                x=top15["Productividad_Estandarizada"],
                y=top15["Funcionario_Ejecutor"],
                orientation="h",
                marker_color=colores_bar,
                text=top15["Productividad_Estandarizada"].round(2),
                textposition="outside",
            )
        )
        fig.add_vline(x=1.0, line_dash="dot", line_color="#003366",
                      annotation_text="Meta = 1.0")
        fig.update_layout(
            height=500, plot_bgcolor="white",
            xaxis_title="Productividad Estandarizada",
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Eficiencia vs % Cumplimiento")
        fig2 = px.scatter(
            func_f,
            x="Eficiencia", y="Pct_Cumplimiento",
            color="Semaforo", size="Operaciones_Ejecutadas",
            hover_name="Funcionario_Ejecutor",
            hover_data={"Ranking": True, "Productividad_Estandarizada": ":.2f"},
            color_discrete_map=SEMAFORO_COLOR,
            labels={"Pct_Cumplimiento": "% Cumplimiento"},
        )
        fig2.add_hline(
            y=META_CIRCULAR, line_dash="dash", line_color="#E74C3C",
            annotation_text=f"Meta {META_CIRCULAR:.0%}",
        )
        fig2.update_layout(height=500, plot_bgcolor="white", yaxis_tickformat=".0%")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Tabla Completa")
    cols_show = [
        "Ranking", "Funcionario_Ejecutor", "Banca_Transformada",
        "Operaciones_Ejecutadas", "Pct_Cumplimiento",
        "Productividad_Estandarizada", "Eficacia", "Efectividad", "Eficiencia",
        "Semaforo",
    ]
    tabla = func_f[cols_show].copy().sort_values("Ranking")
    tabla["Pct_Cumplimiento"] = tabla["Pct_Cumplimiento"].map("{:.1%}".format)
    for c in ["Productividad_Estandarizada", "Eficacia", "Efectividad", "Eficiencia"]:
        tabla[c] = tabla[c].round(2)

    def _color_sem(val):
        return {
            "Verde":    "background-color:#C6EFCE;color:#276221",
            "Amarillo": "background-color:#FFEB9C;color:#9C6500",
            "Alerta":   "background-color:#FFC7CE;color:#9C0006",
            "Rojo":     "background-color:#FFC7CE;color:#9C0006",
        }.get(val, "")

    styled = tabla.style.map(_color_sem, subset=["Semaforo"])
    st.dataframe(styled, use_container_width=True, height=420)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — DETALLE POR LÍNEA
# ══════════════════════════════════════════════════════════════════════════════
else:
    st.markdown(
        '<div class="header-box"><h2 style="margin:0">🔍 Detalle por Línea de Proceso</h2></div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Tiempos promedio por Línea")
        tiempos = (
            det_f.groupby("Linea_Proceso")
            .agg(
                Ejecucion=("Min_Ejecucion", "mean"),
                Espera=("Min_Espera_Asignacion", "mean"),
            )
            .reset_index()
            .sort_values("Ejecucion", ascending=False)
        )
        fig = px.bar(
            tiempos, x="Linea_Proceso", y=["Ejecucion", "Espera"],
            barmode="stack",
            labels={"value": "Minutos promedio", "variable": "Componente", "Linea_Proceso": ""},
            color_discrete_map={"Ejecucion": "#003366", "Espera": "#66B2FF"},
        )
        fig.update_layout(height=360, plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Cumplimiento mensual")
        mensual = (
            det_f.groupby(["Mes", "Mes_Nombre", "Banca_Transformada"])
            .agg(
                Ops=("Cantidad_Solicitudes", "sum"),
                Cumple=("Cumplimiento", lambda x: (x == "CUMPLE").sum()),
            )
            .reset_index()
            .sort_values("Mes")
        )
        mensual["Pct"] = mensual["Cumple"] / mensual["Ops"]
        fig2 = px.line(
            mensual, x="Mes_Nombre", y="Pct",
            color="Banca_Transformada", markers=True,
            labels={"Pct": "% Cumplimiento", "Mes_Nombre": "Mes"},
            color_discrete_sequence=PALETA,
        )
        fig2.add_hline(
            y=META_CIRCULAR, line_dash="dash", line_color="#E74C3C",
            annotation_text=f"Meta {META_CIRCULAR:.0%}",
        )
        fig2.update_layout(height=360, plot_bgcolor="white", yaxis_tickformat=".0%")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Heatmap: % Cumplimiento por Banca × Línea")
    pivot = banca.pivot_table(
        values="Pct_Cumplimiento",
        index="Linea_Proceso",
        columns="Banca_Transformada",
        aggfunc="mean",
    )
    fig3 = px.imshow(
        pivot, text_auto=".1%", aspect="auto",
        color_continuous_scale=["#E74C3C", "#FFBF00", "#00B050"],
        zmin=0.80, zmax=1.0,
        labels={"color": "% Cumplimiento"},
    )
    fig3.update_layout(height=380)
    st.plotly_chart(fig3, use_container_width=True)

    st.subheader("Detalle de operaciones")
    cols_det = [
        "Fecha_Llegada", "Banca_Transformada", "Linea_Proceso",
        "Funcionario_Ejecutor", "Estado_Solicitud",
        "Min_Ejecucion", "Min_Espera_Asignacion", "Min_Ciclo", "Cumplimiento",
    ]
    muestra = det_f[cols_det].head(500)
    st.dataframe(muestra, use_container_width=True, height=300)
    st.caption(f"Mostrando primeras 500 filas de {len(det_f):,} totales filtradas.")
