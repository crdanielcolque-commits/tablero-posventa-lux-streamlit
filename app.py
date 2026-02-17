# ==========================
# TABLERO POSVENTA — V2.2 "LEVEL DIOS" (Macro → Micro)
# - Visual premium (tarjetas HTML, badges, spacing)
# - Chips multiselect VERDE SUAVE (CSS reforzado)
# - P&L Repuestos vs Servicios (macro) + Aperturas (micro)
# - Ranking sucursal (macro) + Micro PRO sucursal+apertura
# - KPIs resto + Gestión con filtro sucursal
# - Labels dentro de barras (valores y %)
# - Parse numérico AR robusto
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
# CSS PREMIUM + CHIPS VERDE SUAVE
# ==========================
st.markdown(
    r"""
<style>
/* Layout */
.block-container { padding-top: 1.1rem; max-width: 1400px; }
[data-testid="stSidebar"] { width: 320px; }
hr { margin: 1.2rem 0; }

/* Títulos */
h1 { letter-spacing: -0.6px; }
h2 { letter-spacing: -0.3px; }
h3 { letter-spacing: -0.2px; }

/* Cards */
.lux-card{
  background: white;
  border: 1px solid rgba(15, 23, 42, 0.08);
  border-radius: 16px;
  padding: 16px 16px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
.lux-card h4{
  margin: 0 0 8px 0;
  font-size: 14px;
  color: rgba(15, 23, 42, 0.75);
  font-weight: 600;
}
.lux-big{
  font-size: 34px;
  font-weight: 800;
  letter-spacing: -0.8px;
  color: #0f172a;
  line-height: 1.1;
}
.lux-sub{
  margin-top: 6px;
  font-size: 13px;
  color: rgba(15, 23, 42, 0.65);
}
.lux-row{
  display:flex;
  gap:12px;
  flex-wrap:wrap;
  margin-top: 10px;
}
.lux-pill{
  display:inline-flex;
  align-items:center;
  gap:8px;
  padding: 6px 10px;
  border-radius: 999px;
  border: 1px solid rgba(15, 23, 42, 0.10);
  background: rgba(15, 23, 42, 0.03);
  font-size: 12px;
  color: rgba(15, 23, 42, 0.75);
}
.lux-dot{
  width:10px; height:10px; border-radius:999px;
  background: rgba(148,163,184,1);
}
.badge{
  display:inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 700;
  border: 1px solid rgba(15, 23, 42, 0.10);
}
.badge-green{ background: rgba(16,185,129,0.12); color:#065f46; border-color: rgba(16,185,129,0.25); }
.badge-amber{ background: rgba(245,158,11,0.12); color:#92400e; border-color: rgba(245,158,11,0.25); }
.badge-red{ background: rgba(239,68,68,0.12); color:#7f1d1d; border-color: rgba(239,68,68,0.25); }
.badge-gray{ background: rgba(148,163,184,0.15); color:#334155; border-color: rgba(148,163,184,0.25); }

/* Hero header */
.hero{
  padding: 18px 18px;
  border-radius: 18px;
  border: 1px solid rgba(15, 23, 42, 0.08);
  background: linear-gradient(135deg, rgba(15, 23, 42, 0.96), rgba(30, 64, 175, 0.88));
  color: white;
  box-shadow: 0 10px 28px rgba(15, 23, 42, 0.18);
}
.hero h1{
  margin: 0;
  font-size: 34px;
  font-weight: 900;
  letter-spacing: -1px;
}
.hero .meta{
  margin-top: 6px;
  font-size: 13px;
  color: rgba(255,255,255,0.82);
}

/* ========= CHIPS MULTISELECT (forzado) =========
   Streamlit cambia internamente, así que vamos a:
   - agarrar tags por data-baseweb="tag"
   - forzar background/border/text
   - también el "x" del chip
*/
div[data-testid="stMultiSelect"] div[data-baseweb="tag"]{
  background-color:#d1fae5 !important;
  border: 1px solid #a7f3d0 !important;
}
div[data-testid="stMultiSelect"] div[data-baseweb="tag"] *{
  color:#065f46 !important;
}
div[data-testid="stMultiSelect"] div[data-baseweb="tag"] svg{
  fill:#065f46 !important;
}

/* Quitar sensación "alerta" general */
[data-testid="stSidebar"] .stMarkdown { color: rgba(15,23,42,0.85); }

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
        return ("—", "badge-gray")
    if c >= 1:
        return ("VERDE", "badge-green")
    if c >= 0.9:
        return ("AMARILLO", "badge-amber")
    return ("ROJO", "badge-red")

def parse_semana_num(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]
    num = num.str.replace(",", ".", regex=False)
    numf = pd.to_numeric(num, errors="coerce")
    return np.floor(numf).astype("Int64")

def clean_unnamed(df0: pd.DataFrame):
    return df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")].copy()

def parse_number_ar(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    s = s.str.replace(r"[\$\s]", "", regex=True)
    s = s.str.replace("%", "", regex=False)
    s = s.str.replace(r"[^0-9\-\.,]", "", regex=True)

    def _one(x):
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return np.nan
        x = str(x)
        if x == "" or x.lower() == "nan":
            return np.nan

        has_dot = "." in x
        has_com = "," in x
        if has_dot and has_com:
            x = x.replace(".", "")
            x = x.replace(",", ".")
        elif has_com and not has_dot:
            x = x.replace(",", ".")
        elif has_dot and not has_com:
            if x.count(".") > 1:
                x = x.replace(".", "")
            else:
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
    if clip_max is not None and x in d.columns:
        d[x] = d[x].clip(upper=clip_max)

    fig = px.bar(d, x=x, y=y, orientation="h", title=title)

    if is_percent:
        fig.update_traces(texttemplate="%{x:.1%}", textposition="inside", insidetextanchor="end")
        fig.update_xaxes(tickformat=".0%")
    else:
        fig.update_traces(texttemplate="%{x:,.0f}", textposition="inside", insidetextanchor="end")

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

def hero(title, meta):
    st.markdown(
        f"""
<div class="hero">
  <h1>{title}</h1>
  <div class="meta">{meta}</div>
</div>
""",
        unsafe_allow_html=True,
    )

def card(title, big, sub, badge_text=None, badge_class="badge-gray"):
    badge_html = f'<span class="badge {badge_class}">{badge_text}</span>' if badge_text else ""
    st.markdown(
        f"""
<div class="lux-card">
  <h4>{title} {badge_html}</h4>
  <div class="lux-big">{big}</div>
  <div class="lux-sub">{sub}</div>
</div>
""",
        unsafe_allow_html=True,
    )

# ==========================
# LOAD
# ==========================
@st.cache_data(ttl=300)
def load():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df0 = pd.read_excel(EXCEL_LOCAL)
    return clean_unnamed(df0)

df = load()

# ==========================
# VALIDACIÓN
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

df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()
df["KPI"] = df["KPI"].astype(str).str.strip()
df["Categoria_KPI"] = df["Categoria_KPI"].astype(str).str.strip()
df["Sucursal"] = df["Sucursal"].astype(str).str.strip()

for col in ["Real_$", "Costo_$", "Margen_$", "Margen_%", "Real_Q", "Objetivo_$", "Objetivo_Q", "Cumplimiento_%"]:
    if col in df.columns:
        df[col] = parse_number_ar(df[col])

# ==========================
# SIDEBAR
# ==========================
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
default_sem = 1 if 1 in semanas else semanas[0]
default_idx = semanas.index(default_sem)

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal_sel = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

# corte acumulado
df_cut = df[df["Semana_Num"] <= semana_corte].copy()
if sucursal_sel != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal_sel].copy()

# aperturas disponibles P&L (solo $)
rep_ap_all = sorted(df[df["KPI"].str.upper() == "REPUESTOS"].loc[df["Tipo_KPI"] == "$", "Categoria_KPI"].dropna().unique().tolist())
srv_ap_all = sorted(df[df["KPI"].str.upper() == "SERVICIOS"].loc[df["Tipo_KPI"] == "$", "Categoria_KPI"].dropna().unique().tolist())

st.sidebar.markdown("---")
st.sidebar.subheader("Incluir variables (P&L)")

rep_incl = st.sidebar.multiselect("Repuestos: aperturas incluidas", rep_ap_all, default=rep_ap_all)
srv_incl = st.sidebar.multiselect("Servicios: aperturas incluidas", srv_ap_all, default=srv_ap_all)

st.sidebar.markdown("---")
rank_metric = st.sidebar.selectbox("Ranking por sucursal (macro)", ["Cumplimiento %", "Gap ($)"], index=0)

# ==========================
# CÁLCULOS P&L
# ==========================
def pl_segment(df_base: pd.DataFrame, kpi_name: str, aperturas_incl: list[str]):
    seg = df_base[(df_base["KPI"].str.upper() == kpi_name.upper()) & (df_base["Tipo_KPI"] == "$")].copy()
    if aperturas_incl:
        seg = seg[seg["Categoria_KPI"].isin(aperturas_incl)].copy()

    scope = seg[seg["Objetivo_$"].fillna(0) > 0].copy()

    total_real = scope["Real_$"].sum()
    total_obj = scope["Objetivo_$"].sum()
    total_c = safe_ratio(total_real, total_obj)
    total_gap = total_obj - total_real

    by_ap = scope.groupby("Categoria_KPI", as_index=False).agg(Real=("Real_$", "sum"), Obj=("Objetivo_$", "sum"))
    by_ap["Cumpl"] = by_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_ap["Gap"] = by_ap["Obj"] - by_ap["Real"]

    by_suc = scope.groupby("Sucursal", as_index=False).agg(Real=("Real_$", "sum"), Obj=("Objetivo_$", "sum"))
    by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

    by_suc_ap = scope.groupby(["Sucursal", "Categoria_KPI"], as_index=False).agg(Real=("Real_$", "sum"), Obj=("Objetivo_$", "sum"))
    by_suc_ap["Cumpl"] = by_suc_ap.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    by_suc_ap["Gap"] = by_suc_ap["Obj"] - by_suc_ap["Real"]

    return dict(total_real=total_real, total_obj=total_obj, total_c=total_c, total_gap=total_gap,
                by_ap=by_ap, by_suc=by_suc, by_suc_ap=by_suc_ap)

rep = pl_segment(df_cut, "Repuestos", rep_incl)
srv = pl_segment(df_cut, "Servicios", srv_incl)

# ==========================
# HERO
# ==========================
hero(
    "Tablero Posventa — Macro → Micro",
    f"Sucursal: <b>{sucursal_sel}</b> &nbsp;|&nbsp; Corte: <b>Semana {int(semana_corte)}</b> &nbsp;|&nbsp; Vista: <b>Semanal + Acumulado</b>",
)

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ==========================
# TAB 1
# ==========================
with tab1:
    st.markdown("### P&L — visión ejecutiva (macro)")
    c1, c2 = st.columns(2, gap="large")

    btxt, bcls = estado_from_cumpl(rep["total_c"])
    with c1:
        card(
            "REPUESTOS (Acumulado)",
            pct_str(rep["total_c"]),
            f"Real {money0(rep['total_real'])} · Obj {money0(rep['total_obj'])} · Gap {money0(rep['total_gap'])}",
            badge_text=btxt,
            badge_class=bcls,
        )

    btxt, bcls = estado_from_cumpl(srv["total_c"])
    with c2:
        card(
            "SERVICIOS (Acumulado)",
            pct_str(srv["total_c"]),
            f"Real {money0(srv['total_real'])} · Obj {money0(srv['total_obj'])} · Gap {money0(srv['total_gap'])}",
            badge_text=btxt,
            badge_class=bcls,
        )

    st.divider()

    st.markdown("### Aperturas — micro (cumplimiento acumulado)")
    m1, m2 = st.columns(2, gap="large")

    with m1:
        d = rep["by_ap"].copy().sort_values("Cumpl", ascending=False)
        if d.empty:
            st.info("Repuestos: sin aperturas con objetivo válido (>0) en este corte.")
        else:
            st.plotly_chart(
                bar_with_labels(d, x="Cumpl", y="Categoria_KPI", title="Repuestos — por apertura", is_percent=True, height=460, xaxis_title="Cumplimiento (Acum.)"),
                use_container_width=True,
            )

    with m2:
        d = srv["by_ap"].copy().sort_values("Cumpl", ascending=False)
        if d.empty:
            st.info("Servicios: sin aperturas con objetivo válido (>0) en este corte.")
        else:
            st.plotly_chart(
                bar_with_labels(d, x="Cumpl", y="Categoria_KPI", title="Servicios — por apertura", is_percent=True, height=460, xaxis_title="Cumplimiento (Acum.)"),
                use_container_width=True,
            )

    st.divider()

    st.markdown("### Ranking por sucursal (macro) + micro por apertura")
    r1, r2 = st.columns(2, gap="large")

    def rank_plot(df_suc, title_prefix):
        if df_suc.empty:
            return None
        if rank_metric == "Cumplimiento %":
            d = df_suc.sort_values("Cumpl", ascending=False)
            return bar_with_labels(d, x="Cumpl", y="Sucursal", title=title_prefix, is_percent=True, height=480, xaxis_title="Cumplimiento (Acum.)")
        d = df_suc.sort_values("Gap", ascending=False)
        return bar_with_labels(d, x="Gap", y="Sucursal", title=title_prefix, is_percent=False, height=480, xaxis_title="Gap (Obj - Real)")

    with r1:
        fig = rank_plot(rep["by_suc"], "Repuestos — por sucursal")
        st.plotly_chart(fig, use_container_width=True) if fig else st.info("Repuestos: sin ranking (obj>0).")

    with r2:
        fig = rank_plot(srv["by_suc"], "Servicios — por sucursal")
        st.plotly_chart(fig, use_container_width=True) if fig else st.info("Servicios: sin ranking (obj>0).")

    st.divider()

    st.markdown("### 🎯 Micro — ranking sucursal + apertura")
    a, b, c, d = st.columns([1.2, 1.2, 0.7, 0.9], gap="large")
    with a:
        rep_ap_pick = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + sorted(rep["by_suc_ap"]["Categoria_KPI"].unique().tolist()))
    with b:
        srv_ap_pick = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + sorted(srv["by_suc_ap"]["Categoria_KPI"].unique().tolist()))
    with c:
        top_n = st.selectbox("Top N", [5, 10, 15, 20], index=1)
    with d:
        show_zero = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=True)

    def micro_df(by_suc_ap, ap_pick):
        dd = by_suc_ap.copy()
        if ap_pick != "Todas las aperturas":
            dd = dd[dd["Categoria_KPI"] == ap_pick].copy()
        if not show_zero:
            dd = dd[~((dd["Obj"].fillna(0) > 0) & (dd["Real"].fillna(0) == 0))].copy()
        dd["Label"] = dd["Sucursal"] + " — " + dd["Categoria_KPI"]
        dd = dd.sort_values("Cumpl", ascending=False).head(top_n)
        return dd

    mm1, mm2 = st.columns(2, gap="large")
    with mm1:
        dmr = micro_df(rep["by_suc_ap"], rep_ap_pick)
        if dmr.empty:
            st.info("Sin datos (Repuestos).")
        else:
            st.plotly_chart(bar_with_labels(dmr, x="Cumpl", y="Label", title="Repuestos — sucursal + apertura (micro)", is_percent=True, height=560, xaxis_title="Cumplimiento (Acum.)"), use_container_width=True)

    with mm2:
        dms = micro_df(srv["by_suc_ap"], srv_ap_pick)
        if dms.empty:
            st.info("Sin datos (Servicios).")
        else:
            st.plotly_chart(bar_with_labels(dms, x="Cumpl", y="Label", title="Servicios — sucursal + apertura (micro)", is_percent=True, height=560, xaxis_title="Cumplimiento (Acum.)"), use_container_width=True)

# ==========================
# TAB 2 — KPIs RESTO
# ==========================
with tab2:
    st.markdown("### KPIs (resto) — Macro → Micro")
    resto = df_cut[~df_cut["KPI"].str.upper().isin(["REPUESTOS", "SERVICIOS"])].copy()
    kpis_resto = sorted(resto["KPI"].dropna().unique().tolist())

    if not kpis_resto:
        st.info("No hay KPIs adicionales cargados (fuera de Repuestos/Servicios).")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto, index=0)
        d0 = resto[resto["KPI"] == kpi_sel].copy()

        tipos = sorted(d0["Tipo_KPI"].dropna().unique().tolist())
        tipo_pref = "$" if "$" in tipos else (tipos[0] if tipos else "$")
        tipo_pref = st.radio("Tipo", tipos if tipos else ["$"], horizontal=True, index=(tipos.index(tipo_pref) if tipo_pref in tipos else 0))

        if tipo_pref == "$":
            d0["Real_val"] = d0["Real_$"].fillna(0.0)
            d0["Obj_val"] = d0["Objetivo_$"].fillna(0.0)
            fmt_real = lambda x: money0(x)
        else:
            d0["Real_val"] = d0["Real_Q"].fillna(0.0)
            d0["Obj_val"] = d0["Objetivo_Q"].fillna(0.0)
            fmt_real = lambda x: num0(x)

        scope = d0[d0["Obj_val"].fillna(0) > 0].copy()
        real_tot = scope["Real_val"].sum()
        obj_tot = scope["Obj_val"].sum()
        ctot = safe_ratio(real_tot, obj_tot)
        btxt, bcls = estado_from_cumpl(ctot)

        c1, c2, c3 = st.columns([1.2, 1, 1], gap="large")
        with c1:
            card(f"{kpi_sel} ({tipo_pref})", pct_str(ctot), f"Real {fmt_real(real_tot)} · Obj {fmt_real(obj_tot)}", badge_text=btxt, badge_class=bcls)
        with c2:
            card("Real (Acum.)", fmt_real(real_tot), "Suma acumulada al corte", None, "badge-gray")
        with c3:
            card("Objetivo (Acum.)", fmt_real(obj_tot), "Suma de objetivos al corte", None, "badge-gray")

        st.divider()

        by_suc = scope.groupby("Sucursal", as_index=False).agg(Real=("Real_val", "sum"), Obj=("Obj_val", "sum"))
        by_suc["Cumpl"] = by_suc.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        by_suc["Gap"] = by_suc["Obj"] - by_suc["Real"]

        if by_suc.empty:
            st.info("Sin ranking por sucursal (no hay objetivos > 0).")
        else:
            if rank_metric == "Cumplimiento %":
                by_suc = by_suc.sort_values("Cumpl", ascending=False)
                fig = bar_with_labels(by_suc, x="Cumpl", y="Sucursal", title="Ranking por sucursal — este KPI", is_percent=True, height=560, xaxis_title="Cumplimiento (Acum.)")
            else:
                by_suc = by_suc.sort_values("Gap", ascending=False)
                fig = bar_with_labels(by_suc, x="Gap", y="Sucursal", title="Ranking por sucursal — este KPI", is_percent=False, height=560, xaxis_title="Gap (Obj - Real)")
            st.plotly_chart(fig, use_container_width=True)

# ==========================
# TAB 3 — GESTIÓN
# ==========================
with tab3:
    st.markdown("### Gestión (desvíos) — Torre de control")
    suc_gest = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)

    d = df_cut.copy()
    if suc_gest != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_gest].copy()

    def _real(r):
        return r["Real_$"] if r["Tipo_KPI"] == "$" else r["Real_Q"]
    def _obj(r):
        return r["Objetivo_$"] if r["Tipo_KPI"] == "$" else r["Objetivo_Q"]

    d["Real_val"] = d.apply(_real, axis=1)
    d["Obj_val"] = d.apply(_obj, axis=1)
    d["Cumpl_calc"] = d.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)
    d["Gap"] = d["Obj_val"] - d["Real_val"]

    d_rel = d[(d["Obj_val"].fillna(0) > 0) & (d["Cumpl_calc"].fillna(1) < 1)].copy()

    g1, g2, g3 = st.columns([1.1, 1.1, 1], gap="large")
    with g1:
        tipo_g = st.selectbox("Tipo", ["Todos", "$", "Q"], index=0)
    with g2:
        kpi_g = st.selectbox("KPI", ["Todos"] + sorted(d_rel["KPI"].dropna().unique().tolist()), index=0)
    with g3:
        topg = st.selectbox("Top desvíos", [10, 20, 30, 50], index=1)

    if d_rel.empty:
        st.success("No hay desvíos relevantes con objetivo válido (en este corte).")
    else:
        dg = d_rel.copy()
        if tipo_g != "Todos":
            dg = dg[dg["Tipo_KPI"] == tipo_g].copy()
        if kpi_g != "Todos":
            dg = dg[dg["KPI"] == kpi_g].copy()

        dg = dg.sort_values("Gap", ascending=False).head(topg)

        dg["Cumpl_str"] = dg["Cumpl_calc"].apply(pct_str)
        dg["Obj_str"] = dg.apply(lambda r: money0(r["Obj_val"]) if r["Tipo_KPI"] == "$" else num0(r["Obj_val"]), axis=1)
        dg["Real_str"] = dg.apply(lambda r: money0(r["Real_val"]) if r["Tipo_KPI"] == "$" else num0(r["Real_val"]), axis=1)
        dg["Gap_str"] = dg.apply(lambda r: money0(r["Gap"]) if r["Tipo_KPI"] == "$" else num0(r["Gap"]), axis=1)

        cols = ["Semana", "Sucursal", "KPI", "Categoria_KPI", "Tipo_KPI", "Real_str", "Obj_str", "Cumpl_str", "Gap_str"]
        if "Comentario / Acción" in dg.columns:
            cols.append("Comentario / Acción")

        st.dataframe(dg[cols], use_container_width=True, height=520)

        st.divider()

        dg_plot = dg.copy()
        dg_plot["Item"] = dg_plot["Sucursal"] + " — " + dg_plot["KPI"] + " — " + dg_plot["Categoria_KPI"]
        fig = bar_with_labels(dg_plot, x="Gap", y="Item", title="Top desvíos (Gap = Obj - Real)", is_percent=False, height=640, xaxis_title="Gap")
        st.plotly_chart(fig, use_container_width=True)
