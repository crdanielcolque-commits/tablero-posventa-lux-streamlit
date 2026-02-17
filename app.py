# ==========================
# TABLERO POSVENTA — V2.1 PRO (FIX DEFINITIVO)
# - Parse numérico AR robusto (coma decimal, % y $)
# - Tabs + filtros incluir/excluir
# - Macro → Micro + ranking + micro PRO
# - Gestión con filtro sucursal
# - Labels dentro de barras
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
/* multiselect tags (chips) - selector más específico */
div.stMultiSelect div[data-baseweb="tag"]{
  background-color:#d1fae5 !important;
  color:#065f46 !important;
  border:1px solid #a7f3d0 !important;
}
div.stMultiSelect div[data-baseweb="tag"] span{
  color:#065f46 !important;
}

/* títulos */
h1, h2, h3 { letter-spacing: -0.2px; }
.block-container { padding-top: 1.1rem; }
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

def clean_unnamed(df0: pd.DataFrame):
    return df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")].copy()

def parse_number_ar(series: pd.Series) -> pd.Series:
    """
    Parser robusto para números en formato AR:
    - "9.808.131,76" -> 9808131.76
    - "9808131,76"   -> 9808131.76
    - "$ 1.234.567"  -> 1234567
    - "57,33%" o "57.33%" -> 57.33
    - vacío -> NaN
    """
    s = series.astype(str).str.strip()

    # Normalizar vacíos
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan})

    # Quitar símbolos comunes
    s = s.str.replace(r"[\$\s]", "", regex=True)
    s = s.str.replace("%", "", regex=False)

    # Dejar solo dígitos, separadores y signo
    s = s.str.replace(r"[^0-9\-\.,]", "", regex=True)

    def _one(x):
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return np.nan
        x = str(x)
        if x == "" or x.lower() == "nan":
            return np.nan

        has_dot = "." in x
        has_com = "," in x

        # Caso "1.234.567,89" => miles con punto, decimal con coma
        if has_dot and has_com:
            x = x.replace(".", "")
            x = x.replace(",", ".")
        # Caso "1234567,89" => decimal coma
        elif has_com and not has_dot:
            x = x.replace(",", ".")
        # Caso "1234567.89" => decimal punto (ok)
        # Caso "1.234.567" => probablemente miles con punto => sacamos puntos
        elif has_dot and not has_com:
            # si hay más de 1 punto: miles
            if x.count(".") > 1:
                x = x.replace(".", "")
            else:
                # si tiene 1 punto pero parece miles (ej "1.234" sin decimales)
                # asumimos miles si hay 3 dígitos después del punto y total corto
                parts = x.split(".")
                if len(parts) == 2 and len(parts[1]) == 3 and len(parts[0]) <= 3:
                    x = x.replace(".", "")
        try:
            return float(x)
        except Exception:
            return np.nan

    return s.map(_one)

def bar_with_labels(df_plot, x, y, title="", is_percent=False, height=420, xaxis_title=None, clip_max=None):
    d = df_plot.copy()

    # opcional: limitar extremos para que no arruinen la escala
    if clip_max is not None and x in d.columns:
        d[x] = d[x].clip(upper=clip_max)

    fig = px.bar(d, x=x, y=y, orientation="h", title=title)

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
    df0 = clean_unnamed(df0)
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
# NORMALIZACIÓN + PARSE AR
# ==========================
df["Semana_Num"] = parse_semana_num(df["Semana"])
df = df[~df["Semana_Num"].isna()].copy()

df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()

# Parse robusto numérico (AR) para columnas críticas
for col in ["Real_$", "Costo_$", "Margen_$", "Margen_%", "Real_Q", "Objetivo_$", "Objetivo_Q", "Cumplimiento_%"]:
    if col in df.columns:
        df[col] = parse_number_ar(df[col])

# ==========================
# SIDEBAR (Filtros)
# ==========================
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
default_sem = 1 if 1 in semanas else semanas[0]
default_idx = semanas.index(default_sem)

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().astype(str).unique().tolist())
sucursal_sel = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

# Corte acumulado
df_cut = df[df["Semana_Num"] <= semana_corte].copy()

# Filtro sucursal global
if sucursal_sel != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"].astype(str) == sucursal_sel].copy()

# ==========================
# LISTAS BASE PARA P&L
# (P&L = Repuestos/Servicios en $)
# ==========================
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
    seg = df_base[
        (df_base["KPI"].astype(str).str.upper() == kpi_name.upper())
        & (df_base["Tipo_KPI"].astype(str).str.strip() == "$")
    ].copy()

    if aperturas_incl:
        seg = seg[seg["Categoria_KPI"].astype(str).isin(aperturas_incl)].copy()

    # ALCANCE PARA CUMPLIMIENTO: solo filas con Obj > 0
    seg_scope = seg[seg["Objetivo_$"].fillna(0) > 0].copy()

    total_real = seg_scope["Real_$"].sum(skipna=True)
    total_obj = seg_scope["Objetivo_$"].sum(skipna=True)
    total_c = safe_ratio(total_real, total_obj)

    # Por apertura
    by_ap = seg_scope.groupby(["Categoria_KPI"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_ap["Cumpl"] = by_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_ap["Gap"] = by_ap["Obj"] - by_ap["Real"]

    # Por sucursal (macro)
    by_suc = seg_scope.groupby(["Sucursal"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

    # Por sucursal + apertura (micro)
    by_suc_ap = seg_scope.groupby(["Sucursal", "Categoria_KPI"], as_index=False).agg(
        Real=("Real_$", "sum"),
        Obj=("Objetivo_$", "sum"),
    )
    by_suc_ap["Cumpl"] = by_suc_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc_ap["Gap"] = by_suc_ap["Obj"] - by_suc_ap["Real"]

    return dict(
        seg=seg,
        seg_scope=seg_scope,
        total_real=total_real,
        total_obj=total_obj,
        total_c=total_c,
        by_ap=by_ap,
        by_suc=by_suc,
        by_suc_ap=by_suc_ap,
    )

def build_kpi_resto(df_base: pd.DataFrame):
    resto = df_base[~df_base["KPI"].astype(str).str.upper().isin(["REPUESTOS", "SERVICIOS"])].copy()
    kpis = sorted(resto["KPI"].dropna().astype(str).unique().tolist())
    return resto, kpis

def filter_gestion(df_base: pd.DataFrame, suc_gest: str):
    d = df_base.copy()
    if suc_gest != "TODAS (Consolidado)":
        d = d[d["Sucursal"].astype(str) == suc_gest].copy()

    def real_val(r):
        return r["Real_$"] if str(r["Tipo_KPI"]).strip() == "$" else r["Real_Q"]

    def obj_val(r):
        return r["Objetivo_$"] if str(r["Tipo_KPI"]).strip() == "$" else r["Objetivo_Q"]

    d["Real_val"] = d.apply(real_val, axis=1)
    d["Obj_val"] = d.apply(obj_val, axis=1)
    d["Cumpl_calc"] = d.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)
    d["Gap"] = d["Obj_val"] - d["Real_val"]

    d_rel = d[(d["Obj_val"].fillna(0) > 0) & (d["Cumpl_calc"].fillna(1) < 1)].copy()
    d_rel = d_rel.sort_values(["Tipo_KPI", "Gap"], ascending=[True, False])
    return d_rel

# ==========================
# HEADER
# ==========================
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: {sucursal_sel} | Corte semana {int(semana_corte)}")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ==========================
# TAB 1 — P&L
# ==========================
with tab1:
    rep = build_pl_df(df_cut, "Repuestos", rep_incl)
    srv = build_pl_df(df_cut, "Servicios", srv_incl)

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.subheader("🧩 REPUESTOS (P&L)")
        st.metric(
            "Cumplimiento (Acum.)",
            pct_str(rep["total_c"]),
            f"Real {money0(rep['total_real'])} / Obj {money0(rep['total_obj'])}",
        )
        st.write(f"Estado: **{estado_from_cumpl(rep['total_c'])}**")

        if rep["total_obj"] == 0:
            st.warning("Repuestos: el Objetivo_$ está en 0 (o no se pudo parsear). Revisá formato del Excel.")

    with c2:
        st.subheader("🧩 SERVICIOS (P&L)")
        st.metric(
            "Cumplimiento (Acum.)",
            pct_str(srv["total_c"]),
            f"Real {money0(srv['total_real'])} / Obj {money0(srv['total_obj'])}",
        )
        st.write(f"Estado: **{estado_from_cumpl(srv['total_c'])}**")

        if srv["total_obj"] == 0:
            st.warning("Servicios: el Objetivo_$ está en 0 (o no se pudo parsear). Revisá formato del Excel.")

    st.divider()

    st.subheader("Aperturas — micro (cumplimiento acumulado)")
    m1, m2 = st.columns(2, gap="large")

    with m1:
        df_rep_ap = rep["by_ap"].copy().sort_values("Cumpl", ascending=False)
        if df_rep_ap.empty:
            st.info("Repuestos: sin aperturas con objetivo válido (>0) en este corte.")
        else:
            st.plotly_chart(
                bar_with_labels(
                    df_rep_ap,
                    x="Cumpl",
                    y="Categoria_KPI",
                    title="Repuestos — por apertura",
                    is_percent=True,
                    height=420,
                    xaxis_title="Cumplimiento (Acum.)",
                ),
                use_container_width=True,
            )

    with m2:
        df_srv_ap = srv["by_ap"].copy().sort_values("Cumpl", ascending=False)
        if df_srv_ap.empty:
            st.info("Servicios: sin aperturas con objetivo válido (>0) en este corte.")
        else:
            st.plotly_chart(
                bar_with_labels(
                    df_srv_ap,
                    x="Cumpl",
                    y="Categoria_KPI",
                    title="Servicios — por apertura",
                    is_percent=True,
                    height=420,
                    xaxis_title="Cumplimiento (Acum.)",
                ),
                use_container_width=True,
            )

    st.divider()

    st.subheader("🏁 Ranking por sucursal (Macro)")
    r1, r2 = st.columns(2, gap="large")

    def rank_plot(df_suc, title_prefix):
        d = df_suc.copy()
        if d.empty:
            return None
        if rank_metric == "Cumplimiento %":
            d = d.sort_values("Cumpl", ascending=False)
            return bar_with_labels(d, x="Cumpl", y="Sucursal", title=title_prefix, is_percent=True, height=420, xaxis_title="Cumplimiento (Acum.)")
        else:
            d = d.sort_values("Gap", ascending=False)
            return bar_with_labels(d, x="Gap", y="Sucursal", title=title_prefix, is_percent=False, height=420, xaxis_title="Gap (Obj - Real)")

    with r1:
        fig = rank_plot(rep["by_suc"], "Repuestos — por sucursal")
        if fig is None:
            st.info("Repuestos: sin ranking (no hay obj>0).")
        else:
            st.plotly_chart(fig, use_container_width=True)

    with r2:
        fig = rank_plot(srv["by_suc"], "Servicios — por sucursal")
        if fig is None:
            st.info("Servicios: sin ranking (no hay obj>0).")
        else:
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    st.subheader("🎯 Micro — ranking sucursal + apertura")

    cA, cB, cC, cD = st.columns([1.2, 1.2, 0.7, 0.9], gap="large")
    with cA:
        rep_ap_pick = st.selectbox(
            "Repuestos (micro)",
            ["Todas las aperturas"] + sorted(rep["by_suc_ap"]["Categoria_KPI"].astype(str).unique().tolist()),
        )
    with cB:
        srv_ap_pick = st.selectbox(
            "Servicios (micro)",
            ["Todas las aperturas"] + sorted(srv["by_suc_ap"]["Categoria_KPI"].astype(str).unique().tolist()),
        )
    with cC:
        top_n = st.selectbox("Top N", [5, 10, 15, 20], index=1)
    with cD:
        show_zero = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=True)

    def micro_df(d0, ap_pick):
        d = d0.copy()
        if d.empty:
            return d
        if ap_pick != "Todas las aperturas":
            d = d[d["Categoria_KPI"].astype(str) == ap_pick].copy()

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
            st.plotly_chart(
                bar_with_labels(dmr, x="Cumpl", y="Label", title="Repuestos — sucursal + apertura (micro)", is_percent=True, height=520, xaxis_title="Cumplimiento (Acum.)"),
                use_container_width=True,
            )

    with mm2:
        dms = micro_df(srv["by_suc_ap"].rename(columns={"Categoria_KPI": "Categoria_KPI", "Real": "Real", "Obj": "Obj"}), srv_ap_pick)
        if dms.empty:
            st.info("Sin datos para este filtro (Servicios).")
        else:
            st.plotly_chart(
                bar_with_labels(dms, x="Cumpl", y="Label", title="Servicios — sucursal + apertura (micro)", is_percent=True, height=520, xaxis_title="Cumplimiento (Acum.)"),
                use_container_width=True,
            )

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
        if d.empty:
            st.info("Sin datos para este KPI.")
        else:
            tipos = d["Tipo_KPI"].astype(str).str.strip().unique().tolist()
            tipo_pref = "$" if "$" in tipos else (tipos[0] if tipos else "$")

            if tipo_pref == "$":
                d["Real_val"] = d["Real_$"].fillna(0.0)
                d["Obj_val"] = d["Objetivo_$"].fillna(0.0)
                unit_title = "$"
            else:
                d["Real_val"] = d["Real_Q"].fillna(0.0)
                d["Obj_val"] = d["Objetivo_Q"].fillna(0.0)
                unit_title = "Q"

            # Scope para cumplimiento: Obj>0
            scope = d[d["Obj_val"].fillna(0) > 0].copy()

            real_tot = scope["Real_val"].sum()
            obj_tot = scope["Obj_val"].sum()
            cumpl_tot = safe_ratio(real_tot, obj_tot)

            a1, a2, a3 = st.columns([1.1, 1, 1], gap="large")
            with a1:
                st.markdown(f"### {kpi_sel} ({unit_title})")
                st.write(f"Estado: **{estado_from_cumpl(cumpl_tot)}**")
            with a2:
                st.metric("Cumplimiento (Acum.)", pct_str(cumpl_tot))
            with a3:
                st.metric(
                    "Real / Objetivo (Acum.)",
                    f"{money0(real_tot) if unit_title=='$' else num0(real_tot)} / {money0(obj_tot) if unit_title=='$' else num0(obj_tot)}",
                )

            st.divider()

            by_suc = scope.groupby(["Sucursal"], as_index=False).agg(
                Real=("Real_val", "sum"),
                Obj=("Obj_val", "sum"),
            )
            by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
            by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

            if by_suc.empty:
                st.info("Sin ranking por sucursal (no hay objetivos > 0 en este corte).")
            else:
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

    suc_gest = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)
    d_rel = filter_gestion(df_cut, suc_gest)

    if d_rel.empty:
        st.success("No hay desvíos relevantes con objetivo válido (en este corte).")
    else:
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

        dg = dg.sort_values("Gap", ascending=False).head(topg)

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
        if "Comentario / Acción" in dg.columns:
            show_cols.append("Comentario / Acción")

        st.dataframe(dg[show_cols], use_container_width=True, height=480)

        st.divider()

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
