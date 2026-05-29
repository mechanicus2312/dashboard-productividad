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

SP_CONFIGURADO = (
    "SHAREPOINT_URL" in st.secrets
    and "SHAREPOINT_USER" in st.secrets
    and "SHAREPOINT_PASS" in st.secrets
    and "SHAREPOINT_FILE" in st.secrets
)


# ── Carga de datos ────────────────────────────────────────────────────────────
def _parsear(buf):
    det = pd.read_excel(
        buf, sheet_name="Detalle",
        parse_dates=["Fecha_Llegada", "Fecha_Asignacion", "Fecha_Fin"],
    )
    buf.seek(0)
    func = pd.read_excel(buf, sheet_name="Resumen_Funcionario",
                         parse_dates=["Ultima_Actividad"])
    buf.seek(0)
    banca = pd.read_excel(buf, sheet_name="Resumen_Banca")
    buf.seek(0)
    params = pd.read_excel(buf, sheet_name="Parametros")

    num_func = [
        "Eficacia", "Efectividad", "Eficiencia",
        "Productividad", "Productividad_Estandarizada", "Pct_Cumplimiento",
    ]
    for c in num_func:
        if func[c].dtype == object:
            func[c] = func[c].str.replace(",", ".", regex=False).astype(float)
    if banca["Pct_Cumplimiento"].dtype == object:
        banca["Pct_Cumplimiento"] = (
            banca["Pct_Cumplimiento"].str.replace(",", ".", regex=False).astype(float)
        )
    p = dict(zip(params["Parametro"], params["Valor"]))
    return det, func, banca, p


@st.cache_data(ttl=3600)
def cargar_desde_sharepoint():
    from office365.runtime.auth.user_credential import UserCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File

    ctx = ClientContext(st.secrets["SHAREPOINT_URL"]).with_credentials(
        UserCredential(st.secrets["SHAREPOINT_USER"], st.secrets["SHAREPOINT_PASS"])
    )
    response = File.open_binary(ctx, st.secrets["SHAREPOINT_FILE"])
    return _parsear(BytesIO(response.content))


@st.cache_data
def cargar_desde_bytes(archivo_bytes):
    return _parsear(BytesIO(archivo_bytes))


# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 📂 Datos")

if SP_CONFIGURADO:
    st.sidebar.info("Conectado a SharePoint")
    if st.sidebar.button("🔄 Actualizar datos"):
        cargar_desde_sharepoint.clear()
    try:
        det, func, banca, params = cargar_desde_sharepoint()
        st.sidebar.success("✅ Datos cargados desde SharePoint")
    except Exception as e:
        st.sidebar.error(f"Error SharePoint: {e}")
        st.stop()
else:
    archivo = st.sidebar.file_uploader(
        "Sube el archivo Excel",
        type=["xlsx"],
        help="Hojas requeridas: Detalle, Resumen_Funcionario, Resumen_Banca, Parametros",
    )
    if archivo is None:
        st.markdown(
            '<div class="header-box"><h2 style="margin:0">📊 Dashboard de Productividad</h2></div>',
            unsafe_allow_html=True,
        )
        st.markdown("### Para comenzar, sube el archivo de datos desde el panel izquierdo.")
        st.info(
            "**Formato esperado:** archivo `.xlsx` con las hojas:\n"
            "- `Detalle` — operaciones individuales\n"
            "- `Resumen_Funcionario` — métricas por funcionario\n"
            "- `Resumen_Banca` — resumen por banca y línea\n"
            "- `Parametros` — metas y parámetros del período"
        )
        st.stop()
    det, func, banca, params = cargar_desde_bytes(archivo.read())
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
