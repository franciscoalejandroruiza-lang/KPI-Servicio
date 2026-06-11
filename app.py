"""
KPI Servicio Técnico - SenIntegral
Cómo correr:  streamlit run app.py
"""

import streamlit as st           # Framework principal de la app web
import pandas as pd               # Manipulación de DataFrames
from datetime import datetime     # Para obtener mes/año actuales como defaults
import io                         # Para envolver bytes del archivo en buffer legible por pandas

# ── CONFIGURACIÓN DE PÁGINA ──────────────────────────────────────────────────
st.set_page_config(page_title="KPI Servicio Técnico", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title { font-size:2rem; font-weight:700; color:#1f2937; margin-bottom:0.2rem; }
    .subtitle   { color:#6b7280; margin-bottom:1.5rem; }

    .step-card {
        background: #f8fafc;
        border-left: 4px solid #3b82f6;
        border-radius: 0 8px 8px 0;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    .step-number {
        font-size: 1.6rem;
        font-weight: 800;
        color: #3b82f6;
        line-height: 1;
        margin-bottom: 0.3rem;
    }
    .step-title {
        font-size: 1rem;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 0.5rem;
    }
    .step-body {
        font-size: 0.9rem;
        color: #374151;
        line-height: 1.6;
    }
    .alert-xlsx {
        background: #fff7ed;
        border: 1px solid #fb923c;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        margin: 0.8rem 0;
        font-size: 0.88rem;
        color: #9a3412;
    }
    .tip-box {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        font-size: 0.85rem;
        color: #166534;
        margin-top: 0.5rem;
    }
    .tab-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.88rem;
        margin-top: 0.5rem;
    }
    .tab-table th {
        background: #eff6ff;
        color: #1e40af;
        padding: 6px 10px;
        text-align: left;
        border-bottom: 2px solid #bfdbfe;
    }
    .tab-table td {
        padding: 6px 10px;
        border-bottom: 1px solid #e5e7eb;
        color: #374151;
    }
    .tab-table tr:last-child td { border-bottom: none; }
</style>
""", unsafe_allow_html=True)


# ── CONSTANTES ───────────────────────────────────────────────────────────────
MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}
EXCLUIDOS_REINCIDENCIA = {"CHI SISTEMAS"}

COL_MAP = {
    "Folio":            "folio",
    "N.° de contrato":  "contrato",
    "Tipo de contrato": "tipo_contrato",
    "N.° de equipo":    "num_equipo",
    "N.° de serie":     "num_serie",
    "Modelo":           "modelo",
    "Nombre legal":     "nombre_legal",
    "Nombre comercial": "nombre_comercial",
    "Grupo Empresarial":"grupo",
    "Técnico":          "tecnico",
    "Referencia":       "referencia",
    "Fecha recepción":  "fecha_recepcion",
    "Última visita":    "ultima_visita",
    "Categoría":        "categoria",
    "Estatus":          "estatus",
    "Prioridad":        "prioridad",
    "Origen":           "origen",
    "Ruta Servicio":    "ruta",
    "Problema reportado":                   "problema",
    "Nombre Legal Cliente de Servicio":     "cliente_legal",
    "Nombre Comercial Cliente de Servicio": "cliente_comercial",
}


# ── FUNCIONES ─────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def cargar_excel(archivo_bytes: bytes) -> pd.DataFrame:
    """Lee el Excel desde bytes, renombra columnas y normaliza tipos."""
    if len(archivo_bytes) == 0:
        raise ValueError("El archivo llegó vacío. Vuelve a cargarlo.")
    try:
        df = pd.read_excel(io.BytesIO(archivo_bytes), engine="openpyxl")
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo: {e}")

    df = df.rename(columns=COL_MAP)

    df["fecha_recepcion"] = pd.to_datetime(df["fecha_recepcion"], errors="coerce")
    df["ultima_visita"]   = pd.to_datetime(df["ultima_visita"],   errors="coerce")

    for col in ["tecnico", "categoria", "estatus", "num_serie"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    return df


def filtrar_mes(df: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    """Filtra el DataFrame al año/mes de fecha_recepcion indicados."""
    mask = (
        (df["fecha_recepcion"].dt.year  == anio) &
        (df["fecha_recepcion"].dt.month == mes)
    )
    return df[mask].copy()


def calcular_reportes_resueltos(
    df_mes: pd.DataFrame,
    pesos: dict
) -> tuple:
    """
    Retorna:
      - pivot_cats : tabla técnico x categoría con conteos
      - pts_base   : Series con puntos ponderados por técnico
      - total_res  : Series con conteo total resueltos por técnico
      - resueltos  : DataFrame filtrado solo a RESUELTA
    """
    resueltos = df_mes[df_mes["estatus"] == "RESUELTA"].copy()
    if resueltos.empty:
        return pd.DataFrame(), pd.Series(dtype=float), pd.Series(dtype=int), resueltos

    resueltos["puntos"] = resueltos["categoria"].map(
        lambda c: pesos.get(c, 1)
    )
    pivot = resueltos.pivot_table(
        index="tecnico", columns="categoria",
        values="folio", aggfunc="count", fill_value=0
    )
    pts_base  = resueltos.groupby("tecnico")["puntos"].sum().rename("pts_base")
    total_res = resueltos.groupby("tecnico")["folio"].count().rename("total_resueltos")
    return pivot, pts_base, total_res, resueltos


def calcular_reincidencias(
    df_completo: pd.DataFrame,
    df_mes: pd.DataFrame,
    meses_historial: int,
    excluidos: set
) -> pd.DataFrame:
    """
    Para cada CORRECTIVO resuelto en df_mes, busca visitas previas al mismo
    num_serie en los meses_historial meses anteriores y genera penalizaciones.
    """
    correctivos_mes = df_mes[
        (df_mes["estatus"] == "RESUELTA") &
        (df_mes["categoria"] == "CORRECTIVO")
    ].copy()

    if correctivos_mes.empty:
        return pd.DataFrame()

    fecha_min_mes = df_mes["fecha_recepcion"].min()
    if pd.isna(fecha_min_mes):
        return pd.DataFrame()

    mes_inicio = fecha_min_mes - pd.DateOffset(months=meses_historial)

    historico = df_completo[
        (df_completo["fecha_recepcion"] >= mes_inicio) &
        (df_completo["fecha_recepcion"] <  fecha_min_mes) &
        (df_completo["categoria"] == "CORRECTIVO")
    ].copy()

    penalizaciones = []
    for _, folio_actual in correctivos_mes.iterrows():
        serie = folio_actual["num_serie"]
        if serie in ("NAN", "", "NONE"):
            continue
        for _, visita in historico[historico["num_serie"] == serie].iterrows():
            tec_prev = visita["tecnico"]
            if tec_prev in excluidos:
                continue
            penalizaciones.append({
                "folio_actual":   folio_actual["folio"],
                "num_serie":      serie,
                "fecha_actual":   folio_actual["fecha_recepcion"],
                "tecnico_actual": folio_actual["tecnico"],
                "folio_previo":   visita["folio"],
                "fecha_previa":   visita["fecha_recepcion"],
                "tecnico_previo": tec_prev,
                "penalizacion":   -1,
            })

    return pd.DataFrame(penalizaciones) if penalizaciones else pd.DataFrame()


def construir_resumen(
    tecnicos: list,
    pts_base: pd.Series,
    total_res: pd.Series,
    df_pen: pd.DataFrame
) -> pd.DataFrame:
    """Consolida pts_base, penalizaciones y total neto en un DataFrame de resumen."""
    resumen = pd.DataFrame(index=sorted(tecnicos))
    resumen.index.name = "Técnico"
    resumen["Pts_Base"]           = pts_base.reindex(resumen.index, fill_value=0)
    resumen["Puntos Extra"]       = 0
    resumen["Penalización"]       = 0
    resumen["Reportes Resueltos"] = total_res.reindex(resumen.index, fill_value=0)

    if not df_pen.empty:
        pen = df_pen.groupby("tecnico_previo")["penalizacion"].sum()
        resumen["Penalización"] = pen.reindex(resumen.index, fill_value=0)

    resumen["TOTAL NETO"] = (
        resumen["Pts_Base"] + resumen["Puntos Extra"] + resumen["Penalización"]
    )
    return resumen.sort_values("TOTAL NETO", ascending=False)


# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    st.divider()

    st.markdown("**📂 Cargar Reporte Excel**")
    archivo = st.file_uploader(
        "Archivo", type=["xlsx", "xls"],
        label_visibility="collapsed"
    )
    st.divider()

    st.markdown("**📅 Período a Evaluar**")
    mes_sel  = st.selectbox(
        "Mes", list(MESES_ES.keys()),
        format_func=lambda m: MESES_ES[m],
        index=datetime.now().month - 1
    )
    anio_sel = st.number_input(
        "Año", min_value=2020, max_value=2030,
        value=datetime.now().year, step=1
    )
    st.divider()

    st.markdown("**🔁 Meses de historial (reincidencias)**")
    meses_historial = st.slider(
        "Meses retroactivos", 1, 12, 3, label_visibility="collapsed"
    )
    st.caption(f"Se revisan los {meses_historial} mes(es) anteriores al período evaluado.")
    st.divider()
    st.caption("KPI Servicio Técnico v1.0 · SenIntegral")


# ── PANTALLA DE BIENVENIDA / INSTRUCTIVO ──────────────────────────────────────
if archivo is None:
    st.markdown('<p class="main-title">📊 KPI Servicio Técnico</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="subtitle">SenIntegral · Sistema de evaluación de productividad técnica</p>',
        unsafe_allow_html=True,
    )

    st.info("👈 Carga tu archivo Excel en el panel izquierdo para comenzar.", icon="📂")

    st.divider()
    st.markdown("## 📖 Instructivo de Uso")
    st.markdown("Sigue estos pasos para generar el reporte de KPI correctamente.")
    st.markdown("")

    # ── FILA 1: Pasos 1 y 2 ──────────────────────────────────────────────────
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown("""
<div class="step-card">
    <div class="step-number">01</div>
    <div class="step-title">Prepara el archivo Excel</div>
    <div class="step-body">
        El sistema lee el reporte exportado desde tu plataforma de servicio técnico.
        Antes de cargarlo, verifica que el archivo esté en el formato correcto.
    </div>
</div>
""", unsafe_allow_html=True)

        st.markdown("""
<div class="alert-xlsx">
    ⚠️ <strong>El archivo DEBE estar en formato .xlsx</strong><br>
    Para convertirlo: abre el archivo en Excel → <em>Archivo → Guardar como</em>
    → en "Tipo" selecciona <strong>Libro de Excel (.xlsx)</strong> → Guardar.
    <br><br>
    ❌ No uses <code>.xls</code>, <code>.csv</code> ni <code>.ods</code> — estos formatos
    pueden causar errores de lectura.
</div>
""", unsafe_allow_html=True)

        st.markdown("""
<div class="tip-box">
    💡 <strong>Tip:</strong> La hoja de datos debe ser la <strong>primera hoja</strong>
    del archivo. Si tienes varias hojas, mueve la de datos al primer lugar.
</div>
""", unsafe_allow_html=True)

    with col2:
        st.markdown("""
<div class="step-card">
    <div class="step-number">02</div>
    <div class="step-title">Carga el archivo y configura el período</div>
    <div class="step-body">
        <ol style="margin: 0.5rem 0 0 1rem; padding: 0;">
            <li style="margin-bottom:6px">En el <strong>panel izquierdo</strong>, haz clic en <strong>Browse files</strong> bajo "Cargar Reporte Excel"</li>
            <li style="margin-bottom:6px">Selecciona tu archivo <code>.xlsx</code> — al aparecer el nombre del archivo, la carga fue exitosa</li>
            <li style="margin-bottom:6px">Elige el <strong>Mes</strong> y <strong>Año</strong> que deseas evaluar</li>
            <li style="margin-bottom:6px">Ajusta los <strong>Meses de historial</strong> para controlar qué tan atrás se buscan reincidencias (default: 3 meses)</li>
        </ol>
    </div>
</div>
""", unsafe_allow_html=True)

        st.markdown("""
<div class="tip-box">
    💡 <strong>Tip:</strong> Si cambias el mes o el año, la app recalcula automáticamente
    todos los indicadores sin necesidad de recargar el archivo.
</div>
""", unsafe_allow_html=True)

    st.markdown("")

    # ── FILA 2: Pasos 3 y 4 ──────────────────────────────────────────────────
    col3, col4 = st.columns(2, gap="large")

    with col3:
        st.markdown("""
<div class="step-card">
    <div class="step-number">03</div>
    <div class="step-title">Navega por las pestañas de resultados</div>
    <div class="step-body">
        Una vez cargado el archivo, verás cuatro pestañas con distintos niveles de detalle:
    </div>
</div>
""", unsafe_allow_html=True)

        st.markdown("""
<table class="tab-table">
    <tr>
        <th>Pestaña</th>
        <th>Contenido</th>
    </tr>
    <tr>
        <td>📋 <strong>Puntaje Final</strong></td>
        <td>Ranking general con puntos base, penalizaciones y total neto por técnico</td>
    </tr>
    <tr>
        <td>📊 <strong>Matriz por Categoría</strong></td>
        <td>Tabla cruzada: cuántos reportes resolvió cada técnico por tipo de servicio</td>
    </tr>
    <tr>
        <td>⚠️ <strong>Penalizaciones</strong></td>
        <td>Detalle de reincidencias detectadas por número de serie de equipo</td>
    </tr>
    <tr>
        <td>⚙️ <strong>Pesos</strong></td>
        <td>Configura cuántos puntos vale cada categoría de reporte</td>
    </tr>
</table>
""", unsafe_allow_html=True)

    with col4:
        st.markdown("""
<div class="step-card">
    <div class="step-number">04</div>
    <div class="step-title">Entiende el sistema de puntuación</div>
    <div class="step-body">
        <ul style="margin: 0.5rem 0 0 1rem; padding: 0;">
            <li style="margin-bottom:6px">Cada reporte con estatus <strong>RESUELTA</strong> suma puntos según su categoría</li>
            <li style="margin-bottom:6px"><strong>PREVENTIVO</strong> vale <strong>2 puntos</strong> por defecto; el resto vale 1 (configurable en ⚙️ Pesos)</li>
            <li style="margin-bottom:6px"><strong>Reincidencia:</strong> si un equipo (mismo N° de serie) tuvo una visita correctiva previa en el historial, el técnico que la atendió recibe <strong>−1 punto</strong></li>
            <li style="margin-bottom:6px"><strong>CHI Sistemas</strong> está excluido de penalizaciones</li>
        </ul>
    </div>
</div>
""", unsafe_allow_html=True)

        st.markdown("""
<div class="tip-box">
    💡 <strong>Tip:</strong> Ajusta los pesos en la pestaña ⚙️ <strong>Pesos</strong>
    antes de revisar el puntaje final — los cambios se aplican en tiempo real a todo el ranking.
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.caption("KPI Servicio Técnico v1.0 · SenIntegral · Ante cualquier duda contacta al administrador del sistema.")
    st.stop()


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
nombre_archivo = archivo.name
if (
    "archivo_bytes"  not in st.session_state or
    st.session_state.get("archivo_nombre") != nombre_archivo
):
    st.session_state["archivo_bytes"]  = archivo.read()
    st.session_state["archivo_nombre"] = nombre_archivo

archivo_bytes = st.session_state["archivo_bytes"]

if len(archivo_bytes) == 0:
    st.error("El archivo llegó vacío. Intenta cargarlo de nuevo.")
    st.stop()

# ── ENCABEZADO ────────────────────────────────────────────────────────────────
st.markdown('<p class="main-title">📊 KPI Servicio Técnico</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="subtitle">Período: <b>{MESES_ES[mes_sel]} {anio_sel}</b> · '
    f'Historial: {meses_historial} mes(es) retroactivo(s)</p>',
    unsafe_allow_html=True,
)

with st.spinner("Procesando archivo..."):
    df_raw = cargar_excel(archivo_bytes)

df_mes = filtrar_mes(df_raw, anio_sel, mes_sel)
if df_mes.empty:
    st.warning(f"No hay registros para {MESES_ES[mes_sel]} {anio_sel}.")
    st.stop()

tecnicos_en_mes = sorted(df_mes["tecnico"].dropna().unique().tolist())


# ── PESTAÑAS ──────────────────────────────────────────────────────────────────
tab_resumen, tab_matriz, tab_pen, tab_pesos = st.tabs([
    "📋 Puntaje Final",
    "📊 Matriz por Categoría",
    "⚠️ Penalizaciones por Técnico",
    "⚙️ Pesos de Categorías",
])


# ════════════════════════════════════════════════════════════════════════════
#  TAB 4 — PESOS (se procesa primero porque los demás lo necesitan)
# ════════════════════════════════════════════════════════════════════════════
with tab_pesos:
    st.subheader("⚙️ Puntaje asignado a cada Categoría")
    st.markdown(
        "Ajusta cuántos puntos vale cada tipo de reporte resuelto. "
        "Por defecto **todas valen 1**, excepto **PREVENTIVO que vale 2**."
    )
    todas_cats = sorted(df_raw["categoria"].dropna().unique().tolist())
    pesos = {}
    cols_p = st.columns(3)
    for i, cat in enumerate(todas_cats):
        with cols_p[i % 3]:
            pesos[cat] = st.number_input(
                label=cat.title(),
                min_value=0, max_value=10,
                value=2 if cat == "PREVENTIVO" else 1,
                step=1, key=f"p_{cat}"
            )


# ── CÁLCULOS PRINCIPALES ──────────────────────────────────────────────────────
pivot_cats, pts_base, total_res, df_resueltos = calcular_reportes_resueltos(df_mes, pesos)
df_pen_df = calcular_reincidencias(df_raw, df_mes, meses_historial, EXCLUIDOS_REINCIDENCIA)
resumen   = construir_resumen(tecnicos_en_mes, pts_base, total_res, df_pen_df)


# ════════════════════════════════════════════════════════════════════════════
#  TAB 1 — PUNTAJE FINAL
# ════════════════════════════════════════════════════════════════════════════
with tab_resumen:
    st.subheader(f"Resumen de Puntuación: {MESES_ES[mes_sel]} {anio_sel}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📥 Reportes en el mes",    len(df_mes))
    c2.metric("✅ Resueltos",             len(df_resueltos))
    c3.metric("⚠️ Penalizaciones tot.",   int(resumen["Penalización"].sum()))
    mejor = resumen["TOTAL NETO"].idxmax() if not resumen.empty else "—"
    c4.metric("🏆 Mejor técnico", mejor.title() if mejor != "—" else "—")

    st.divider()

    disp = resumen[["Pts_Base", "Puntos Extra", "Penalización", "TOTAL NETO", "Reportes Resueltos"]].copy()

    def color_pen(val):
        if isinstance(val, (int, float)) and val < 0:
            return "color:#dc2626; font-weight:bold"
        return ""

    st.dataframe(
        disp.style.map(color_pen, subset=["Penalización"]).format("{:.0f}"),
        use_container_width=True,
        height=300,
    )

    st.divider()
    st.markdown("#### 🏅 Ranking de Productividad")
    max_pts = max(float(resumen["Pts_Base"].max()), 1.0)

    for i, (tec, row) in enumerate(resumen.iterrows(), 1):
        medalla = {1: "🥇", 2: "🥈", 3: "🥉"}.get(i, f"#{i}")
        color   = "#22c55e" if row["TOTAL NETO"] >= 0 else "#ef4444"
        ancho   = max(int((float(row["Pts_Base"]) / max_pts) * 100), 2)
        pen_txt = f"· {int(row['Penalización'])} pen." if row["Penalización"] != 0 else ""
        st.markdown(
            f"<div style='margin-bottom:10px;'>"
            f"<div style='display:flex;justify-content:space-between;margin-bottom:3px;'>"
            f"<span><b>{medalla} {tec.title()}</b></span>"
            f"<span style='color:#6b7280;'>{int(row['Pts_Base'])} pts base {pen_txt}"
            f" → <b>{int(row['TOTAL NETO'])} total</b></span></div>"
            f"<div style='background:#e5e7eb;border-radius:4px;height:12px;'>"
            f"<div style='background:{color};width:{ancho}%;height:12px;border-radius:4px;'>"
            f"</div></div></div>",
            unsafe_allow_html=True,
        )


# ════════════════════════════════════════════════════════════════════════════
#  TAB 2 — MATRIZ POR CATEGORÍA
# ════════════════════════════════════════════════════════════════════════════
with tab_matriz:
    st.subheader("Matriz Operativa — Reportes Resueltos por Técnico y Categoría")
    st.caption("Solo cuenta Estatus = RESUELTA en el período seleccionado.")

    if pivot_cats.empty:
        st.info("No hay reportes resueltos en el período.")
    else:
        piv = pivot_cats.copy()
        piv["TOTAL"] = piv.sum(axis=1)
        piv = piv.sort_values("TOTAL", ascending=False)

        cols_color = [c for c in piv.columns if c != "TOTAL"]
        max_val = float(piv[cols_color].max().max()) if cols_color else 1.0

        def color_verde(val):
            if not isinstance(val, (int, float)) or val == 0 or max_val == 0:
                return "background-color:#ffffff"
            ratio = val / max_val
            r = int(220 - ratio * 120)
            g = int(240 - ratio * 80)
            b = int(220 - ratio * 150)
            return f"background-color:rgb({r},{g},{b}); color:#1a1a1a"

        st.dataframe(
            piv.style.map(color_verde, subset=cols_color).format("{:.0f}"),
            use_container_width=True,
            height=350,
        )

        st.divider()
        st.markdown("#### 📄 Detalle de Reportes Resueltos")
        tec_f = st.selectbox("Filtrar por técnico:", ["Todos"] + list(resumen.index))
        det = df_resueltos.copy()
        if tec_f != "Todos":
            det = det[det["tecnico"] == tec_f]

        cols_det = [c for c in
                    ["folio", "tecnico", "num_serie", "categoria",
                     "fecha_recepcion", "ultima_visita", "nombre_comercial", "problema"]
                    if c in det.columns]
        st.dataframe(
            det[cols_det].rename(columns={
                "folio":           "Folio",
                "tecnico":         "Técnico",
                "num_serie":       "N° Serie",
                "categoria":       "Categoría",
                "fecha_recepcion": "Fecha Recepción",
                "ultima_visita":   "Última Visita",
                "nombre_comercial":"Cliente",
                "problema":        "Problema",
            }).sort_values("Fecha Recepción", ascending=False),
            use_container_width=True,
            height=400,
        )


# ════════════════════════════════════════════════════════════════════════════
#  TAB 3 — PENALIZACIONES
# ════════════════════════════════════════════════════════════════════════════
with tab_pen:
    st.subheader("⚠️ Penalizaciones por Reincidencia (solo CORRECTIVO)")
    st.markdown(
        f"Para cada CORRECTIVO resuelto en **{MESES_ES[mes_sel]} {anio_sel}**, "
        f"se busca si el mismo número de serie tuvo visitas previas en los "
        f"**{meses_historial} mes(es) anteriores**. "
        f"Cada visita previa = **-1 punto** al técnico que la atendió. "
        f"_(CHI Sistemas excluido de penalizaciones)_"
    )

    if df_pen_df.empty:
        st.success("✅ No se encontraron reincidencias en el período analizado.")
    else:
        pen_res = df_pen_df.groupby("tecnico_previo").agg(
            Reincidencias=("penalizacion", "count"),
            Puntos_Restados=("penalizacion", "sum"),
        ).sort_values("Puntos_Restados")

        col_a, col_b = st.columns([1, 2])

        with col_a:
            st.markdown("**Resumen por técnico**")
            st.dataframe(pen_res, use_container_width=True)

        with col_b:
            st.markdown("**Detalle completo**")
            st.dataframe(
                df_pen_df[[
                    "folio_actual", "num_serie", "fecha_actual", "tecnico_actual",
                    "folio_previo", "fecha_previa", "tecnico_previo", "penalizacion",
                ]].rename(columns={
                    "folio_actual":   "Folio Actual",
                    "num_serie":      "N° Serie",
                    "fecha_actual":   "Fecha Actual",
                    "tecnico_actual": "Técnico Actual",
                    "folio_previo":   "Folio Previo",
                    "fecha_previa":   "Fecha Previa",
                    "tecnico_previo": "Técnico Penalizado",
                    "penalizacion":   "Puntos",
                }).sort_values("Fecha Actual", ascending=False),
                use_container_width=True,
                height=400,
            )

        st.divider()
        st.markdown("#### 🔍 Ver reincidencias de un técnico específico")
        tec_p = st.selectbox(
            "Técnico:",
            sorted(df_pen_df["tecnico_previo"].unique()),
            key="sp"
        )
        sub = df_pen_df[df_pen_df["tecnico_previo"] == tec_p]
        st.markdown(
            f"**{tec_p.title()}** → "
            f"**{len(sub)} reincidencia(s)** = "
            f"**{int(sub['penalizacion'].sum())} puntos**"
        )
        for _, r in sub.sort_values("fecha_actual", ascending=False).iterrows():
            fa = r["fecha_actual"].strftime("%d/%m/%Y") if pd.notna(r["fecha_actual"]) else "—"
            fp = r["fecha_previa"].strftime("%d/%m/%Y") if pd.notna(r["fecha_previa"]) else "—"
            st.markdown(
                f"- Folio **{int(r['folio_actual'])}** · "
                f"Serie `{r['num_serie']}` · "
                f"Recibido: `{fa}` · "
                f"Visita previa: `{fp}` "
                f"(folio {int(r['folio_previo'])})"
            )