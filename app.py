# ==========================
# TABLERO POSVENTA — V2.1 PRO (Tabs + Incluir/Excluir + Labels)
# Macro → Micro (Semanal + Acumulado)
# ==========================

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import gdown

st.set_page_config(page_title="Tablero Posventa", layout="wide")

# ==========================
# DRIVE CONFIG
# ==========================
DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# ==========================
# CSS (chips verde suave + estética)
# ==========================
st.markdown(
    """
<style>
/* tags multiselect (chips) */
div[data-baseweb="tag"]{
  background-color:#dff5e6 !important;
  color:#14532d !important;
  border:1px solid #b7e4c7 !important;
}
/* borde del tag */
div[data-baseweb="tag"] span{
  color:#14532d !important;
}
/* títulos */
h1, h2, h3 { letter-spacing: -0.2px; }
/* reduce padding superior */
.block-container { padding-top: 1.2rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ==========================
# HELPERS
# ==========================

def safe_ratio(n, d):
    try:
        if d is None or pd.isna(d) or float(d) == 0:
            return np.nan
        return float(n) / float(d)
    except Exception:
        return np.nan

def money0(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"${float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def num0(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def pct_str(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "—"

def estado_from_cumpl(c):
    if c is None or pd.isna(c):
        return "—"
    if c >= 1:
        return "Verde"
    if c >= 0.9:
        return "Amarillo"
    return "Rojo"

def parse_semana_num(series: pd.Series) -> pd.Series:
    # Acepta: "Semana 1", "1", "1.0", "Semana 1.0", etc.
    s = series.astype(str).str.strip()
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]
    num = num.str.replace(",", ".", regex=False)
    numf = pd.to_numeric(num, errors="coerce")
    return np.floor(numf).astype("Int64")

def ensure_cols(df0: pd.DataFrame):
    # Limpia columnas "Unnamed"
    df0 = df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")]
    return df0

def bar_with_labels(df_plot, x, y, title="", is_percent=False, height=420, xaxis_title=None):
    fig = px.bar(df_plot, x=x, y=y, orientation="h", title=title)
    if is_percent:
        fig.update_traces(
            texttemplate="%{x:.1%}",
            textposition="inside",
            insidetextanchor="end",
        )
        fig.update_xaxes(tickformat=".0%")
    else:
        fig.update_traces(
            texttemplate="%{x:,.0f}",
            textposition="inside",
            insidetextanchor="end",
        )
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=40 if title else 10, b=10),
        uniformtext_minsize=10,
        uniformtext_mode="hide",
        yaxis_title=None,
        xaxis_title=xaxis_title if xaxis_title else x,
        showlegend=False,
    )
    return fig

# ==========================
# CARGA
# ==========================

@st.cache_data(ttl=300)
def load():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df0 = pd.read_excel(EXCEL_LOCAL)
    df0 = ensure_cols(df0)
    return df0

df = load()

# ==========================
# VALIDACIÓN MÍNIMA
# ==========================
required = ["Semana", "Sucursal", "KPI", "Categoria_KPI", "Tipo_KPI", "Real_$", "Real_Q", "Objetivo_$", "Objetivo_Q"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error("Faltan columnas requeridas en el Excel:")
    st.write(missing)
    st.stop()

# ==========================
# NORMALIZACIÓN
# ==========================
df["Semana_Num"] = parse_semana_num(df["Semana"])
df = df[~df["Semana_Num"].isna()].copy()

# Números robustos
for col in ["Real_$", "Costo_$", "Margen_$", "Margen_%", "Real_Q", "Objetivo_$", "Objetivo_Q", "Cumplimiento_%"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()

# ==========================
# SIDEBAR (Filtros)
# ==========================
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
default_sem = 1 if 1 in semanas else semanas[0]
default_idx = semanas.index(default_sem)

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal_sel = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

# Corte acumulado
df_cut = df[df["Semana_Num"] <= semana_corte].copy()

# Filtro sucursal operativo global
if sucursal_sel != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal_sel].copy()

# ==========================
# LISTAS BASE
# ==========================
kpi_macro = ["Repuestos", "Servicios"]

# Aperturas disponibles en P&L por segmento (Tipo $)
rep_aperturas_all = sorted(
    df[df["KPI"].astype(str).str.upper() == "REPUESTOS"]
    .loc[df["Tipo_KPI"].astype(str).str.strip() == "$", "Categoria_KPI"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

srv_aperturas_all = sorted(
    df[df["KPI"].astype(str).str.upper() == "SERVICIOS"]
    .loc[df["Tipo_KPI"].astype(str).str.strip() == "$", "Categoria_KPI"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

# Sidebar: incluir aperturas (P&L)
st.sidebar.markdown("---")
st.sidebar.subheader("Incluir variables (P&L)")

rep_incl = st.sidebar.multiselect(
    "Repuestos: aperturas incluidas",
    rep_aperturas_all,
    default=rep_aperturas_all,
)

srv_incl = st.sidebar.multiselect(
    "Servicios: aperturas incluidas",
    srv_aperturas_all,
    default=srv_aperturas_all,
)

# Sidebar ranking selector (para tab 1 + tab 2)
st.sidebar.markdown("---")
rank_metric = st.sidebar.selectbox(
    "Ranking por sucursal (macro)",
    ["Cumplimiento %", "Gap ($)"],
    index=0,
)

# ==========================
# FUNCIONES DE CÁLCULO
# ==========================

def build_pl_df(df_base: pd.DataFrame, kpi_name: str, aperturas_incl: list[str]):
    """
    Devuelve acumulado P&L (Tipo $) por:
    - total segmento
    - por apertura
    - por sucursal (macro)
    - por sucursal + apertura (micro pro)
    """
    seg = df_base[
        (df_base["KPI"].astype(str).str.upper() == kpi_name.upper())
        & (df_base["Tipo_KPI"].astype(str).str.strip() == "$")
    ].copy()

    if aperturas_incl:
        seg = seg[seg["Categoria_KPI"].astype(str).isin(aperturas_incl)].copy()

    # Total
    total_real = seg["Real_$"].sum(skipna=True)
    total_obj = seg["Objetivo_$"].sum(skipna=True)
    total_c = safe_ratio(total_real, total_obj)

    # Por apertura
    by_ap = seg.groupby(["Categoria_KPI"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_ap["Cumpl"] = by_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_ap["Gap"] = by_ap["Obj"] - by_ap["Real"]

    # Por sucursal (macro)
    by_suc = seg.groupby(["Sucursal"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

    # Por sucursal + apertura (micro pro)
    by_suc_ap = seg.groupby(["Sucursal", "Categoria_KPI"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_suc_ap["Cumpl"] = by_suc_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc_ap["Gap"] = by_suc_ap["Obj"] - by_suc_ap["Real"]

    return dict(
        seg=seg,
        total_real=total_real,
        total_obj=total_obj,
        total_c=total_c,
        by_ap=by_ap,
        by_suc=by_suc,
        by_suc_ap=by_suc_ap,
    )

def build_kpi_resto(df_base: pd.DataFrame):
    """KPIs que NO son Repuestos/Servicios."""
    resto = df_base[~df_base["KPI"].astype(str).str.upper().isin(["REPUESTOS", "SERVICIOS"])].copy()
    # lista KPI
    kpis = sorted(resto["KPI"].dropna().astype(str).unique().tolist())
    return resto, kpis

def filter_gestion(df_base: pd.DataFrame, suc_gest: str):
    # Gestión: mostramos desvíos relevantes
    d = df_base.copy()
    if suc_gest != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_gest].copy()

    # Calculamos cumplimiento robusto dependiendo tipo
    def real_val(r):
        return r["Real_$"] if str(r["Tipo_KPI"]).strip() == "$" else r["Real_Q"]

    def obj_val(r):
        return r["Objetivo_$"] if str(r["Tipo_KPI"]).strip() == "$" else r["Objetivo_Q"]

    d["Real_val"] = d.apply(real_val, axis=1)
    d["Obj_val"] = d.apply(obj_val, axis=1)
    d["Cumpl_calc"] = d.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)

    # "Desvío" como Gap (Obj-Real) para $; para Q también
    d["Gap"] = d["Obj_val"] - d["Real_val"]

    # Consideramos desvío si:
    # - Obj > 0 y Cumpl < 1
    # - o Gap positivo grande
    d_rel = d[(d["Obj_val"].fillna(0) > 0) & (d["Cumpl_calc"].fillna(1) < 1)].copy()
    d_rel = d_rel.sort_values(["Tipo_KPI", "Gap"], ascending=[True, False])

    return d_rel

# ==========================
# HEADER
# ==========================
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: {sucursal_sel} | Corte semana {float(semana_corte):.1f}")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ==========================
# TAB 1 — P&L
# ==========================
with tab1:
    # Construcción P&L por segmento
    rep = build_pl_df(df_cut, "Repuestos", rep_incl)
    srv = build_pl_df(df_cut, "Servicios", srv_incl)

    # Macro cards
    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.subheader("🧩 REPUESTOS (P&L)")
        st.metric("Cumplimiento (Acum.)", pct_str(rep["total_c"]), f"Real {money0(rep['total_real'])} / Obj {money0(rep['total_obj'])}")
        st.write(f"Estado: **{estado_from_cumpl(rep['total_c'])}**")

    with c2:
        st.subheader("🧩 SERVICIOS (P&L)")
        st.metric("Cumplimiento (Acum.)", pct_str(srv["total_c"]), f"Real {money0(srv['total_real'])} / Obj {money0(srv['total_obj'])}")
        st.write(f"Estado: **{estado_from_cumpl(srv['total_c'])}**")

    st.divider()

    # Micro por apertura (cumplimiento)
    st.subheader("Aperturas — micro (cumplimiento acumulado)")

    m1, m2 = st.columns(2, gap="large")

    with m1:
        df_rep_ap = rep["by_ap"].copy()
        df_rep_ap = df_rep_ap.sort_values("Cumpl", ascending=False)
        fig_rep_ap = bar_with_labels(
            df_rep_ap,
            x="Cumpl",
            y="Categoria_KPI",
            title="Repuestos — por apertura",
            is_percent=True,
            height=420,
            xaxis_title="Cumplimiento (Acum.)",
        )
        st.plotly_chart(fig_rep_ap, use_container_width=True)

    with m2:
        df_srv_ap = srv["by_ap"].copy()
        df_srv_ap = df_srv_ap.sort_values("Cumpl", ascending=False)
        fig_srv_ap = bar_with_labels(
            df_srv_ap,
            x="Cumpl",
            y="Categoria_KPI",
            title="Servicios — por apertura",
            is_percent=True,
            height=420,
            xaxis_title="Cumplimiento (Acum.)",
        )
        st.plotly_chart(fig_srv_ap, use_container_width=True)

    st.divider()

    # Ranking macro por sucursal
    st.subheader("🏁 Ranking por sucursal (Macro)")

    r1, r2 = st.columns(2, gap="large")

    def rank_plot(df_suc, title_prefix):
        d = df_suc.copy()
        if rank_metric == "Cumplimiento %":
            d = d.sort_values("Cumpl", ascending=False)
            fig = bar_with_labels(d, x="Cumpl", y="Sucursal", title=title_prefix, is_percent=True, height=420, xaxis_title="Cumplimiento (Acum.)")
        else:
            d = d.sort_values("Gap", ascending=False)
            fig = bar_with_labels(d, x="Gap", y="Sucursal", title=title_prefix, is_percent=False, height=420, xaxis_title="Gap (Obj - Real)")
        return fig

    with r1:
        st.plotly_chart(rank_plot(rep["by_suc"], "Repuestos — por sucursal"), use_container_width=True)

    with r2:
        st.plotly_chart(rank_plot(srv["by_suc"], "Servicios — por sucursal"), use_container_width=True)

    st.divider()

    # Micro PRO: sucursal + apertura
    st.subheader("🎯 Micro — ranking sucursal + apertura")

    cA, cB, cC, cD = st.columns([1.2, 1.2, 0.7, 0.9], gap="large")
    with cA:
        rep_ap_pick = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + sorted(rep["by_suc_ap"]["Categoria_KPI"].astype(str).unique().tolist()))
    with cB:
        srv_ap_pick = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + sorted(srv["by_suc_ap"]["Categoria_KPI"].astype(str).unique().tolist()))
    with cC:
        top_n = st.selectbox("Top N", [5, 10, 15, 20], index=1)
    with cD:
        show_zero = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=True)

    def micro_df(d0, ap_pick):
        d = d0.copy()
        if ap_pick != "Todas las aperturas":
            d = d[d["Categoria_KPI"].astype(str) == ap_pick].copy()

        # si no quiere ceros: filtra obj>0 y real==0
        if not show_zero:
            d = d[~((d["Obj"].fillna(0) > 0) & (d["Real"].fillna(0) == 0))].copy()

        d["Label"] = d["Sucursal"].astype(str) + " — " + d["Categoria_KPI"].astype(str)
        d = d.sort_values("Cumpl", ascending=False).head(top_n)
        return d

    mm1, mm2 = st.columns(2, gap="large")

    with mm1:
        dmr = micro_df(rep["by_suc_ap"].rename(columns={"Categoria_KPI": "Categoria_KPI", "Real": "Real", "Obj": "Obj"}), rep_ap_pick)
        if dmr.empty:
            st.info("Sin datos para este filtro (Repuestos).")
        else:
            fig = bar_with_labels(
                dmr,
                x="Cumpl",
                y="Label",
                title="Repuestos — sucursal + apertura (micro)",
                is_percent=True,
                height=520,
                xaxis_title="Cumplimiento (Acum.)",
            )
            st.plotly_chart(fig, use_container_width=True)

    with mm2:
        dms = micro_df(srv["by_suc_ap"].rename(columns={"Categoria_KPI": "Categoria_KPI", "Real": "Real", "Obj": "Obj"}), srv_ap_pick)
        if dms.empty:
            st.info("Sin datos para este filtro (Servicios).")
        else:
            fig = bar_with_labels(
                dms,
                x="Cumpl",
                y="Label",
                title="Servicios — sucursal + apertura (micro)",
                is_percent=True,
                height=520,
                xaxis_title="Cumplimiento (Acum.)",
            )
            st.plotly_chart(fig, use_container_width=True)

# ==========================
# TAB 2 — KPIs RESTO
# ==========================
with tab2:
    resto, kpis_resto = build_kpi_resto(df_cut)

    st.subheader("📌 KPIs (resto) — Macro → Micro")

    if not kpis_resto:
        st.info("No hay KPIs adicionales cargados (fuera de Repuestos/Servicios).")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto, index=0)

        d = resto[resto["KPI"].astype(str) == str(kpi_sel)].copy()

        # Determinar tipo predominante (si mezcla, priorizamos $ si hay)
        tipos = d["Tipo_KPI"].astype(str).str.strip().unique().tolist()
        tipo_pref = "$" if "$" in tipos else (tipos[0] if tipos else "$")

        # Real/Obj según tipo
        if tipo_pref == "$":
            d["Real_val"] = d["Real_$"].fillna(0.0)
            d["Obj_val"] = d["Objetivo_$"].fillna(0.0)
            unit_title = "$"
        else:
            d["Real_val"] = d["Real_Q"].fillna(0.0)
            d["Obj_val"] = d["Objetivo_Q"].fillna(0.0)
            unit_title = "Q"

        # Acumulado por KPI (consolidado)
        real_tot = d["Real_val"].sum()
        obj_tot = d["Obj_val"].sum()
        cumpl_tot = safe_ratio(real_tot, obj_tot)

        a1, a2, a3 = st.columns([1.1, 1, 1], gap="large")
        with a1:
            st.markdown(f"### {kpi_sel} ({unit_title})")
            st.write(f"Estado: **{estado_from_cumpl(cumpl_tot)}**")
        with a2:
            st.metric("Cumplimiento (Acum.)", pct_str(cumpl_tot))
        with a3:
            st.metric("Real / Objetivo (Acum.)", f"{money0(real_tot) if unit_title=='$' else num0(real_tot)} / {money0(obj_tot) if unit_title=='$' else num0(obj_tot)}")

        st.divider()

        # Ranking por sucursal (este KPI)
        by_suc = d.groupby(["Sucursal"], as_index=False).agg(
            Real=("Real_val", "sum"),
            Obj=("Obj_val", "sum"),
        )
        by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

        # Orden “ranking”
        if rank_metric == "Cumplimiento %":
            by_suc = by_suc.sort_values("Cumpl", ascending=False)
            fig = bar_with_labels(by_suc, x="Cumpl", y="Sucursal", title="Ranking por sucursal — este KPI", is_percent=True, height=520, xaxis_title="Cumplimiento (Acum.)")
        else:
            by_suc = by_suc.sort_values("Gap", ascending=False)
            fig = bar_with_labels(by_suc, x="Gap", y="Sucursal", title="Ranking por sucursal — este KPI", is_percent=False, height=520, xaxis_title="Gap (Obj - Real)")

        st.plotly_chart(fig, use_container_width=True)

# ==========================
# TAB 3 — GESTIÓN (DESVÍOS)
# ==========================
with tab3:
    st.subheader("🧪 Gestión (desvíos)")

    # Filtro de sucursal específico para gestión (como pediste)
    suc_gest = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)

    d_rel = filter_gestion(df_cut if sucursal_sel == "TODAS (Consolidado)" else df[df["Semana_Num"] <= semana_corte], suc_gest)

    if d_rel.empty:
        st.success("No hay desvíos relevantes con objetivo válido (en este corte).")
    else:
        # Filtros extra de gestión
        g1, g2, g3 = st.columns([1.1, 1.1, 1], gap="large")
        with g1:
            tipo_g = st.selectbox("Tipo", ["Todos", "$", "Q"], index=0)
        with g2:
            kpi_g = st.selectbox("KPI", ["Todos"] + sorted(d_rel["KPI"].dropna().astype(str).unique().tolist()), index=0)
        with g3:
            topg = st.selectbox("Top desvíos", [10, 20, 30, 50], index=1)

        dg = d_rel.copy()
        if tipo_g != "Todos":
            dg = dg[dg["Tipo_KPI"].astype(str).str.strip() == tipo_g].copy()
        if kpi_g != "Todos":
            dg = dg[dg["KPI"].astype(str) == kpi_g].copy()

        dg["Cumpl_str"] = dg["Cumpl_calc"].apply(lambda x: pct_str(x))
        dg["Obj_str"] = dg.apply(lambda r: money0(r["Obj_val"]) if str(r["Tipo_KPI"]).strip() == "$" else num0(r["Obj_val"]), axis=1)
        dg["Real_str"] = dg.apply(lambda r: money0(r["Real_val"]) if str(r["Tipo_KPI"]).strip() == "$" else num0(r["Real_val"]), axis=1)
        dg["Gap_str"] = dg.apply(lambda r: money0(r["Gap"]) if str(r["Tipo_KPI"]).strip() == "$" else num0(r["Gap"]), axis=1)

        # Tabla de gestión (top por Gap)
        dg = dg.sort_values("Gap", ascending=False).head(topg)

        st.caption("Listado de desvíos (Obj>0 y Cumpl<100%).")
        show_cols = [
            "Semana",
            "Sucursal",
            "KPI",
            "Categoria_KPI",
            "Tipo_KPI",
            "Real_str",
            "Obj_str",
            "Cumpl_str",
            "Gap_str",
        ]
        # Comentario si existe
        if "Comentario / Acción" in dg.columns:
            show_cols.append("Comentario / Acción")

        st.dataframe(dg[show_cols], use_container_width=True, height=480)

        st.divider()

        # Gráfico: Top gaps (labels incluidos)
        dg_plot = dg.copy()
        dg_plot["Item"] = dg_plot["Sucursal"].astype(str) + " — " + dg_plot["KPI"].astype(str) + " — " + dg_plot["Categoria_KPI"].astype(str)

        fig = bar_with_labels(
            dg_plot.sort_values("Gap", ascending=False),
            x="Gap",
            y="Item",
            title="Top desvíos (Gap = Obj - Real)",
            is_percent=False,
            height=600,
            xaxis_title="Gap",
        )
        st.plotly_chart(fig, use_container_width=True)
