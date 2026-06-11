"""
KPI Servicio Técnico - SenIntegral
Cómo correr:  streamlit run app.py
"""
 
import streamlit as st
import pandas as pd
from datetime import datetime
import io
 
# ── CONFIGURACIÓN DE PÁGINA ──────────────────────────────────────────────────
st.set_page_config(page_title="KPI Servicio Técnico", page_icon="📊", layout="wide")
 
st.markdown("""
<style>
    .main-title { font-size:2rem; font-weight:700; color:#1f2937; margin-bottom:0.2rem; }
    .subtitle   { color:#6b7280; margin-bottom:1.5rem; }
</style>
""", unsafe_allow_html=True)
 
 
# ── CONSTANTES ───────────────────────────────────────────────────────────────
MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}
EXCLUIDOS_REINCIDENCIA = {"CHI SISTEMAS"}
 
# Mapa de renombrado de columnas del Excel
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
    "Problema reportado":               "problema",
    "Nombre Legal Cliente de Servicio":     "cliente_legal",
    "Nombre Comercial Cliente de Servicio": "cliente_comercial",
}
 
 
# ── FUNCIONES ─────────────────────────────────────────────────────────────────
 
@st.cache_data
def cargar_excel(archivo_bytes):
    """Lee el Excel, renombra columnas y convierte fechas."""
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
 
 
def filtrar_mes(df, anio, mes):
    mask = (
        (df["fecha_recepcion"].dt.year  == anio) &
        (df["fecha_recepcion"].dt.month == mes)
    )
    return df[mask].copy()
 
 
def calcular_reportes_resueltos(df_mes, pesos):
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
 
 
def calcular_reincidencias(df_completo, df_mes, meses_historial, excluidos):
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
 
 
def construir_resumen(tecnicos, pts_base, total_res, df_pen):
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
    archivo = st.file_uploader("Archivo", type=["xlsx", "xls"], label_visibility="collapsed")
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
 
 
# ── ENCABEZADO ────────────────────────────────────────────────────────────────
st.markdown('<p class="main-title">📊 KPI Servicio Técnico</p>', unsafe_allow_html=True)
st.markdown(
    f'<p class="subtitle">Período: <b>{MESES_ES[mes_sel]} {anio_sel}</b> · '
    f'Historial: {meses_historial} mes(es) retroactivo(s)</p>',
    unsafe_allow_html=True,
)
 
if archivo is None:
    st.info("👈 Carga tu archivo Excel en el panel izquierdo para comenzar.", icon="📂")
    st.stop()
 
 
# ── CARGA DE DATOS (session_state para evitar buffer vacío) ───────────────────
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
 
    # Tabla de resumen con color en penalizaciones
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
 
        # Color verde manual (sin matplotlib)
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