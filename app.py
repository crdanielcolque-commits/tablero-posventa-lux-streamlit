# ==========================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado)
# V2.1 PRO (un solo archivo, listo para copiar/pegar)
# ==========================================================

import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import gdown

# --------------------------
# CONFIG
# --------------------------
st.set_page_config(page_title="Tablero Posventa", layout="wide")

# 👉 Pegá tu ID de Drive acá (el de tu Google Sheet)
DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# --------------------------
# ESTILO (chips verde suave + layout limpio)
# --------------------------
st.markdown(
    """
<style>
/* Sidebar un poco más agradable */
section[data-testid="stSidebar"] {
  background: #f7f9fb;
}

/* Chips (tags) de multiselect -> verde suave */
[data-baseweb="tag"] {
  background-color: #d8f5e3 !important;
  color: #0f5132 !important;
  border: 1px solid #a9e7c2 !important;
}
[data-baseweb="tag"] svg {
  color: #0f5132 !important;
}

/* Titulares */
h1, h2, h3 {
  letter-spacing: -0.3px;
}

/* Cards suaves */
.kpi-card {
  background: white;
  border: 1px solid #eef2f6;
  border-radius: 16px;
  padding: 16px 16px 12px 16px;
  box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
.kpi-title {
  font-size: 12px;
  color: #667085;
  margin-bottom: 6px;
}
.kpi-value {
  font-size: 28px;
  font-weight: 750;
  color: #101828;
  line-height: 1.05;
}
.kpi-sub {
  font-size: 12px;
  color: #667085;
  margin-top: 8px;
}
.badge {
  display:inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: 700;
  border: 1px solid transparent;
}
.badge-green { background:#ecfdf3; color:#027a48; border-color:#abefc6; }
.badge-yellow{ background:#fffaeb; color:#b54708; border-color:#fedf89; }
.badge-red   { background:#fef3f2; color:#b42318; border-color:#fecdca; }
.badge-gray  { background:#f2f4f7; color:#344054; border-color:#eaecf0; }

hr { border:none; border-top: 1px solid #eef2f6; margin: 16px 0; }

.small-note { font-size:12px; color:#667085; }
</style>
""",
    unsafe_allow_html=True
)

# --------------------------
# HELPERS
# --------------------------
def safe_ratio(n, d):
    try:
        if d is None or pd.isna(d) or float(d) == 0:
            return np.nan
        return float(n) / float(d)
    except Exception:
        return np.nan

def money(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"${float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def num(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def pct(x):
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

def badge_html(est):
    if est == "Verde":
        return '<span class="badge badge-green">VERDE</span>'
    if est == "Amarillo":
        return '<span class="badge badge-yellow">AMARILLO</span>'
    if est == "Rojo":
        return '<span class="badge badge-red">ROJO</span>'
    return '<span class="badge badge-gray">—</span>'

def parse_semana_num(series: pd.Series) -> pd.Series:
    """
    Acepta: 'Semana 1', '1', '1.0', 'Semana 1.0', etc.
    Devuelve Int64 (nullable)
    """
    s = series.astype(str).str.strip()
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]
    num = num.str.replace(",", ".", regex=False)
    numf = pd.to_numeric(num, errors="coerce")
    numi = np.floor(numf).astype("Int64")
    return numi

def clean_unnamed_cols(df):
    return df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")].copy()

def to_float_series(s):
    # admite números con coma, % en string, etc.
    if s is None:
        return pd.Series(dtype=float)
    x = s.astype(str).str.replace(".", "", regex=False)  # miles
    x = x.str.replace(",", ".", regex=False)
    x = x.str.replace("%", "", regex=False)
    return pd.to_numeric(x, errors="coerce")

def ensure_col(df, name):
    if name not in df.columns:
        df[name] = np.nan
    return df

def kpi_card(title, value, sub=None, badge=None):
    b = badge_html(badge) if badge else ""
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
<div class="kpi-card">
  <div class="kpi-title">{title}</div>
  <div style="display:flex; align-items:center; gap:10px;">
    <div class="kpi-value">{value}</div>
    <div>{b}</div>
  </div>
  {sub_html}
</div>
""",
        unsafe_allow_html=True
    )

def plot_bar_h(df, x, y, title, x_suffix="", height=380, text=None, sort_desc=True, x_range=None):
    if df.empty:
        st.info("Sin datos para mostrar.")
        return

    d = df.copy()
    if sort_desc:
        d = d.sort_values(x, ascending=True)  # horizontal bar: ascending-> menor arriba, mayor abajo (visual)
    fig = px.bar(d, x=x, y=y, orientation="h", text=text if text else x)
    fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
    fig.update_layout(
        title=title,
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        xaxis_title="",
        yaxis_title="",
    )
    if x_range is not None:
        fig.update_xaxes(range=x_range)
    if x_suffix:
        fig.update_traces(texttemplate="%{text} " + x_suffix)
    st.plotly_chart(fig, use_container_width=True)

# --------------------------
# LOAD DATA (Drive -> xlsx)
# --------------------------
@st.cache_data(ttl=300)
def load_from_drive(file_id: str) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df0 = pd.read_excel(EXCEL_LOCAL)
    df0 = clean_unnamed_cols(df0)
    df0.columns = [c.strip() if isinstance(c, str) else c for c in df0.columns]
    return df0

df_raw = load_from_drive(DRIVE_FILE_ID)

# --------------------------
# VALIDACIÓN / NORMALIZACIÓN
# --------------------------
# Columnas esperadas (pero lo hacemos tolerante)
df = df_raw.copy()
df = ensure_col(df, "Fecha")
df = ensure_col(df, "Semana")
df = ensure_col(df, "Sucursal")
df = ensure_col(df, "KPI")
df = ensure_col(df, "Categoria_KPI")
df = ensure_col(df, "Tipo_KPI")
df = ensure_col(df, "Real_$")
df = ensure_col(df, "Costo_$")
df = ensure_col(df, "Margen_$")
df = ensure_col(df, "Margen_%")
df = ensure_col(df, "Real_Q")
df = ensure_col(df, "Objetivo_$")
df = ensure_col(df, "Objetivo_Q")
df = ensure_col(df, "Cumplimiento_%")
df = ensure_col(df, "Estado")
df = ensure_col(df, "Comentario / Acción")

# Semana num robusta
df["Semana_Num"] = parse_semana_num(df["Semana"])
df = df[~df["Semana_Num"].isna()].copy()

# Normalizar texto
for col in ["Sucursal", "KPI", "Categoria_KPI", "Tipo_KPI", "Estado"]:
    df[col] = df[col].astype(str).str.strip()

# Normalizar numéricos
df["Real_$"] = to_float_series(df["Real_$"]).fillna(0.0)
df["Costo_$"] = to_float_series(df["Costo_$"]).fillna(0.0)
df["Margen_$"] = to_float_series(df["Margen_$"])
df["Margen_%"] = to_float_series(df["Margen_%"]) / 100.0  # si venía 30.64% -> 0.3064
df["Real_Q"] = to_float_series(df["Real_Q"]).fillna(0.0)
df["Objetivo_$"] = to_float_series(df["Objetivo_$"]).fillna(0.0)
df["Objetivo_Q"] = to_float_series(df["Objetivo_Q"]).fillna(0.0)

# --------------------------
# SIDEBAR — filtros obligatorios
# --------------------------
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
if not semanas:
    st.error("No se encontraron semanas válidas en la columna 'Semana'.")
    st.stop()

# Default: semana 1 si existe, sino la menor
default_sem = 1 if 1 in semanas else semanas[0]
default_idx = semanas.index(default_sem)

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal_sel = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

st.sidebar.markdown("<hr/>", unsafe_allow_html=True)

# --------------------------
# RECORTE acumulado + sucursal
# --------------------------
df_cut = df[df["Semana_Num"] <= semana_corte].copy()

if sucursal_sel != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal_sel].copy()

# --------------------------
# Segmentos macro
# --------------------------
# KPI "macro" esperado: REPUESTOS / SERVICIOS
df_cut["KPI_UP"] = df_cut["KPI"].astype(str).str.upper()

# --------------------------
# SIDEBAR — incluir/excluir aperturas (P&L)
# --------------------------
st.sidebar.markdown("### Incluir variables (P&L)")

rep_aperturas_all = sorted(df_cut[df_cut["KPI_UP"] == "REPUESTOS"]["Categoria_KPI"].dropna().unique().tolist())
srv_aperturas_all = sorted(df_cut[df_cut["KPI_UP"] == "SERVICIOS"]["Categoria_KPI"].dropna().unique().tolist())

rep_aperturas_sel = st.sidebar.multiselect(
    "Repuestos: aperturas incluidas",
    rep_aperturas_all,
    default=rep_aperturas_all
)

srv_aperturas_sel = st.sidebar.multiselect(
    "Servicios: aperturas incluidas",
    srv_aperturas_all,
    default=srv_aperturas_all
)

st.sidebar.markdown("<hr/>", unsafe_allow_html=True)

# Ranking micro: top N + show zeros
st.sidebar.markdown("### Ranking (micro)")
top_n = st.sidebar.selectbox("Top N", [5, 10, 15, 20], index=1)
show_zeros = st.sidebar.checkbox("Mostrar 0% (obj>0 y real=0)", value=False)

# --------------------------
# TABs
# --------------------------
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal_sel}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ==========================================================
# UTIL: construir acumulados por tipo ($ o Q)
# ==========================================================
def build_acum(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve df acumulado por Sucursal/KPI/Categoria_KPI/Tipo_KPI con:
    Real_Acum, Obj_Acum, Cumpl_Acum, Gap_Acum, Margen_Acum (si existe)
    """
    d = df_in.copy()

    # Definir Real/Obj según Tipo_KPI
    # Tipo puede venir '$' o 'Q'
    d["Tipo_KPI"] = d["Tipo_KPI"].astype(str).str.strip()
    d["Tipo_UP"] = d["Tipo_KPI"].str.upper()

    # Real
    d["Real_val"] = np.where(d["Tipo_UP"].isin(["$", "S", "USD", "ARS"]), d["Real_$"], d["Real_Q"])
    # Obj
    d["Obj_val"] = np.where(d["Tipo_UP"].isin(["$", "S", "USD", "ARS"]), d["Objetivo_$"], d["Objetivo_Q"])

    # Gap
    d["Gap_val"] = d["Obj_val"] - d["Real_val"]

    # margen (solo para $)
    d["Margen_val"] = np.where(d["Tipo_UP"].isin(["$", "S", "USD", "ARS"]), d["Margen_$"].fillna(0.0), np.nan)

    out = d.groupby(["Sucursal", "KPI", "Categoria_KPI", "Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_val", "sum"),
        Obj_Acum=("Obj_val", "sum"),
        Gap_Acum=("Gap_val", "sum"),
        Margen_Acum=("Margen_val", "sum"),
    )

    out["Cumpl_Acum"] = out.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)

    # Etiqueta para gráfico %
    out["Cumpl_Acum_plot"] = out["Cumpl_Acum"].copy()
    # Si Obj=0 => NaN para no romper gráficos de cumplimiento
    out.loc[(out["Obj_Acum"].isna()) | (out["Obj_Acum"] == 0), "Cumpl_Acum_plot"] = np.nan

    # Estado
    out["Estado_Acum"] = out["Cumpl_Acum"].apply(estado_from_cumpl)

    return out

df_acum = build_acum(df_cut)

# ==========================================================
# TAB 1 — P&L (macro → micro)
# ==========================================================
with tab1:
    st.markdown("### 🧩 P&L — Macro → Micro")

    # Aplicar filtros de aperturas incluidas
    d_rep = df_acum[df_acum["KPI"].astype(str).str.upper() == "REPUESTOS"].copy()
    d_srv = df_acum[df_acum["KPI"].astype(str).str.upper() == "SERVICIOS"].copy()

    if rep_aperturas_sel:
        d_rep = d_rep[d_rep["Categoria_KPI"].isin(rep_aperturas_sel)]
    if srv_aperturas_sel:
        d_srv = d_srv[d_srv["Categoria_KPI"].isin(srv_aperturas_sel)]

    # Macro cards: preferimos Tipo_KPI "$" para P&L (si existiera mezcla, filtramos por $)
    d_rep_money = d_rep[d_rep["Tipo_KPI"].astype(str).str.strip() == "$"].copy()
    d_srv_money = d_srv[d_srv["Tipo_KPI"].astype(str).str.strip() == "$"].copy()

    rep_real = d_rep_money["Real_Acum"].sum()
    rep_obj = d_rep_money["Obj_Acum"].sum()
    rep_c = safe_ratio(rep_real, rep_obj)
    rep_est = estado_from_cumpl(rep_c)

    srv_real = d_srv_money["Real_Acum"].sum()
    srv_obj = d_srv_money["Obj_Acum"].sum()
    srv_c = safe_ratio(srv_real, srv_obj)
    srv_est = estado_from_cumpl(srv_c)

    colA, colB = st.columns(2, gap="large")
    with colA:
        kpi_card(
            "REPUESTOS (P&L) — Cumplimiento (Acum.)",
            pct(rep_c),
            sub=f"Real {money(rep_real)}  |  Obj {money(rep_obj)}",
            badge=rep_est
        )
    with colB:
        kpi_card(
            "SERVICIOS (P&L) — Cumplimiento (Acum.)",
            pct(srv_c),
            sub=f"Real {money(srv_real)}  |  Obj {money(srv_obj)}",
            badge=srv_est
        )

    st.markdown("<hr/>", unsafe_allow_html=True)

    # --------------------------
    # Micro: aperturas por segmento (cumplimiento %)
    # --------------------------
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    c1, c2 = st.columns(2, gap="large")

    # Repuestos por apertura
    with c1:
        rep_micro = d_rep_money.groupby(["Categoria_KPI"], as_index=False).agg(
            Real=("Real_Acum", "sum"),
            Obj=("Obj_Acum", "sum")
        )
        rep_micro["Cumpl"] = rep_micro.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

        # filtro show zeros
        if not show_zeros:
            rep_micro = rep_micro[~((rep_micro["Obj"] > 0) & (rep_micro["Real"] == 0))].copy()

        rep_micro["Texto"] = rep_micro["Cumpl"].apply(pct)
        rep_micro["Cumpl_plot"] = rep_micro["Cumpl"]

        # Orden correcto: ranking por cumplimiento
        rep_micro = rep_micro.sort_values("Cumpl_plot", ascending=True)

        fig_rep = px.bar(
            rep_micro,
            x="Cumpl_plot",
            y="Categoria_KPI",
            orientation="h",
            text="Texto",
            title="Repuestos — por apertura"
        )
        fig_rep.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig_rep.update_layout(height=360, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig_rep.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig_rep, use_container_width=True)

    # Servicios por apertura
    with c2:
        srv_micro = d_srv_money.groupby(["Categoria_KPI"], as_index=False).agg(
            Real=("Real_Acum", "sum"),
            Obj=("Obj_Acum", "sum")
        )
        srv_micro["Cumpl"] = srv_micro.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

        if not show_zeros:
            srv_micro = srv_micro[~((srv_micro["Obj"] > 0) & (srv_micro["Real"] == 0))].copy()

        srv_micro["Texto"] = srv_micro["Cumpl"].apply(pct)
        srv_micro["Cumpl_plot"] = srv_micro["Cumpl"]
        srv_micro = srv_micro.sort_values("Cumpl_plot", ascending=True)

        fig_srv = px.bar(
            srv_micro,
            x="Cumpl_plot",
            y="Categoria_KPI",
            orientation="h",
            text="Texto",
            title="Servicios — por apertura"
        )
        fig_srv.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig_srv.update_layout(height=360, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig_srv.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig_srv, use_container_width=True)

    st.markdown("<hr/>", unsafe_allow_html=True)

    # --------------------------
    # Ranking macro por sucursal (repuestos y servicios)
    # --------------------------
    st.markdown("### 🏁 Ranking por sucursal (Macro)")
    st.markdown('<div class="small-note">Ordenado por Cumplimiento acumulado. Incluye Real/Obj dentro de cada barra.</div>', unsafe_allow_html=True)

    # construir ranking macro por sucursal (solo $)
    d_money = df_acum[df_acum["Tipo_KPI"].astype(str).str.strip() == "$"].copy()

    # aplicar filtros P&L a nivel aperturas incluidas
    d_rep_rank = d_money[d_money["KPI"].astype(str).str.upper() == "REPUESTOS"].copy()
    d_srv_rank = d_money[d_money["KPI"].astype(str).str.upper() == "SERVICIOS"].copy()
    if rep_aperturas_sel:
        d_rep_rank = d_rep_rank[d_rep_rank["Categoria_KPI"].isin(rep_aperturas_sel)]
    if srv_aperturas_sel:
        d_srv_rank = d_srv_rank[d_srv_rank["Categoria_KPI"].isin(srv_aperturas_sel)]

    rep_rank = d_rep_rank.groupby("Sucursal", as_index=False).agg(Real=("Real_Acum", "sum"), Obj=("Obj_Acum", "sum"))
    rep_rank["Cumpl"] = rep_rank.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    rep_rank["Texto"] = rep_rank.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)

    srv_rank = d_srv_rank.groupby("Sucursal", as_index=False).agg(Real=("Real_Acum", "sum"), Obj=("Obj_Acum", "sum"))
    srv_rank["Cumpl"] = srv_rank.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    srv_rank["Texto"] = srv_rank.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)

    colR, colS = st.columns(2, gap="large")

    with colR:
        rep_rank_plot = rep_rank.dropna(subset=["Cumpl"]).copy()
        rep_rank_plot = rep_rank_plot.sort_values("Cumpl", ascending=True)
        fig = px.bar(rep_rank_plot, x="Cumpl", y="Sucursal", orientation="h", text="Texto", title="Repuestos — por sucursal")
        fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig.update_xaxes(tickformat=".0%")
        if not rep_rank_plot.empty:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Repuestos: sin ranking (obj>0).")

    with colS:
        srv_rank_plot = srv_rank.dropna(subset=["Cumpl"]).copy()
        srv_rank_plot = srv_rank_plot.sort_values("Cumpl", ascending=True)
        fig = px.bar(srv_rank_plot, x="Cumpl", y="Sucursal", orientation="h", text="Texto", title="Servicios — por sucursal")
        fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig.update_xaxes(tickformat=".0%")
        if not srv_rank_plot.empty:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Servicios: sin ranking (obj>0).")

    st.markdown("<hr/>", unsafe_allow_html=True)

    # --------------------------
    # ✅ Micro PRO: ranking sucursal + apertura (lo pediste)
    # --------------------------
    st.markdown("### 🎯 Micro — ranking sucursal + apertura")

    micro_col1, micro_col2, micro_col3, micro_col4 = st.columns([2, 2, 1, 2], gap="large")
    with micro_col1:
        rep_apertura_focus = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + rep_aperturas_all, index=0)
    with micro_col2:
        srv_apertura_focus = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + srv_aperturas_all, index=0)
    with micro_col3:
        micro_topn = st.selectbox("Top N", [5, 10, 15, 20], index=1)
    with micro_col4:
        micro_show0 = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=show_zeros)

    # dataset micro por sucursal + apertura (solo $)
    micro_money = df_acum[df_acum["Tipo_KPI"].astype(str).str.strip() == "$"].copy()

    micro_rep = micro_money[micro_money["KPI"].astype(str).str.upper() == "REPUESTOS"].copy()
    micro_srv = micro_money[micro_money["KPI"].astype(str).str.upper() == "SERVICIOS"].copy()

    # aplicar aperturas incluidas sidebar (baseline)
    if rep_aperturas_sel:
        micro_rep = micro_rep[micro_rep["Categoria_KPI"].isin(rep_aperturas_sel)]
    if srv_aperturas_sel:
        micro_srv = micro_srv[micro_srv["Categoria_KPI"].isin(srv_aperturas_sel)]

    # foco opcional
    if rep_apertura_focus != "Todas las aperturas":
        micro_rep = micro_rep[micro_rep["Categoria_KPI"] == rep_apertura_focus]
    if srv_apertura_focus != "Todas las aperturas":
        micro_srv = micro_srv[micro_srv["Categoria_KPI"] == srv_apertura_focus]

    # agrupar por sucursal + apertura
    micro_rep_g = micro_rep.groupby(["Sucursal", "Categoria_KPI"], as_index=False).agg(
        Real=("Real_Acum", "sum"),
        Obj=("Obj_Acum", "sum")
    )
    micro_rep_g["Cumpl"] = micro_rep_g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

    micro_srv_g = micro_srv.groupby(["Sucursal", "Categoria_KPI"], as_index=False).agg(
        Real=("Real_Acum", "sum"),
        Obj=("Obj_Acum", "sum")
    )
    micro_srv_g["Cumpl"] = micro_srv_g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

    # remover 0% si no quiere
    if not micro_show0:
        micro_rep_g = micro_rep_g[~((micro_rep_g["Obj"] > 0) & (micro_rep_g["Real"] == 0))].copy()
        micro_srv_g = micro_srv_g[~((micro_srv_g["Obj"] > 0) & (micro_srv_g["Real"] == 0))].copy()

    # ranking top N: por menor cumplimiento (peores) o por gap? acá por cumplimiento (peores)
    micro_rep_g["Label"] = micro_rep_g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    micro_srv_g["Label"] = micro_srv_g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)

    micro_rep_g = micro_rep_g.sort_values("Cumpl", ascending=True).head(micro_topn)
    micro_srv_g = micro_srv_g.sort_values("Cumpl", ascending=True).head(micro_topn)

    # construir etiqueta y-axis combinada
    micro_rep_g["Item"] = micro_rep_g["Sucursal"] + " — " + micro_rep_g["Categoria_KPI"]
    micro_srv_g["Item"] = micro_srv_g["Sucursal"] + " — " + micro_srv_g["Categoria_KPI"]

    cc1, cc2 = st.columns(2, gap="large")
    with cc1:
        fig = px.bar(micro_rep_g, x="Cumpl", y="Item", orientation="h", text="Label", title="Repuestos — sucursal + apertura (micro)")
        fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig.update_layout(height=460, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    with cc2:
        fig = px.bar(micro_srv_g, x="Cumpl", y="Item", orientation="h", text="Label", title="Servicios — sucursal + apertura (micro)")
        fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig.update_layout(height=460, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

# ==========================================================
# TAB 2 — KPIs (resto)
# ==========================================================
with tab2:
    st.markdown("### 📌 KPIs (resto) — Macro → Micro")
    st.markdown('<div class="small-note">Acá van los KPIs que no son Repuestos/Servicios (ej: Accesorios, Neumáticos, Campañas, CPUS, etc.).</div>', unsafe_allow_html=True)

    # Determinar lista de KPIs resto
    kpi_resto = sorted(df_cut[~df_cut["KPI_UP"].isin(["REPUESTOS", "SERVICIOS"])]["KPI"].dropna().unique().tolist())
    if not kpi_resto:
        st.info("No se detectaron KPIs 'resto' en la base (además de Repuestos/Servicios).")
    else:
        kpi_pick = st.selectbox("Elegí un KPI (resto)", kpi_resto, index=0)

        d_k = df_acum[df_acum["KPI"] == kpi_pick].copy()

        # Determinar si es $ o Q (si hay mezcla, prioriza $ si existe)
        tipos = d_k["Tipo_KPI"].dropna().unique().tolist()
        tipo_pick = "$" if "$" in tipos else (tipos[0] if tipos else "$")
        d_k = d_k[d_k["Tipo_KPI"] == tipo_pick].copy()

        # Macro: consolidado (acumulado)
        k_real = d_k["Real_Acum"].sum()
        k_obj = d_k["Obj_Acum"].sum()
        k_c = safe_ratio(k_real, k_obj)
        k_est = estado_from_cumpl(k_c)

        if tipo_pick == "$":
            sub = f"Real {money(k_real)}  |  Obj {money(k_obj)}"
            val = pct(k_c)
        else:
            sub = f"Real {num(k_real)}  |  Obj {num(k_obj)}"
            val = pct(k_c)

        kpi_card(f"{kpi_pick} ({tipo_pick}) — Cumplimiento (Acum.)", val, sub=sub, badge=k_est)

        st.markdown("<hr/>", unsafe_allow_html=True)

        # Ranking por sucursal — este KPI
        rank = d_k.groupby("Sucursal", as_index=False).agg(Real=("Real_Acum", "sum"), Obj=("Obj_Acum", "sum"))
        rank["Cumpl"] = rank.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

        # Q o $ labels
        if tipo_pick == "$":
            rank["Label"] = rank.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
        else:
            rank["Label"] = rank.apply(lambda r: f"{pct(r['Cumpl'])} | {num(r['Real'])}/{num(r['Obj'])}", axis=1)

        rank_plot = rank.dropna(subset=["Cumpl"]).copy()
        rank_plot = rank_plot.sort_values("Cumpl", ascending=True)

        fig = px.bar(rank_plot, x="Cumpl", y="Sucursal", orientation="h", text="Label", title="Ranking por sucursal — este KPI")
        fig.update_traces(textposition="inside", insidetextanchor="end", cliponaxis=False)
        fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10), xaxis_title="", yaxis_title="")
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

# ==========================================================
# TAB 3 — Gestión (desvíos)
# ==========================================================
with tab3:
    st.markdown("### 🧪 Gestión (desvíos) — Accionable")
    st.markdown('<div class="small-note">Objetivo: detectar qué explica el gap (Obj - Real) y priorizar acciones.</div>', unsafe_allow_html=True)

    # Filtro independiente de sucursal en Gestión (lo pediste)
    st.markdown("#### Filtros de Gestión")
    g_col1, g_col2, g_col3 = st.columns([2, 2, 2], gap="large")
    with g_col1:
        g_suc = st.selectbox("Sucursal (Gestión)", ["TODAS"] + sorted(df["Sucursal"].dropna().unique().tolist()), index=0)
    with g_col2:
        g_seg = st.selectbox("Segmento", ["Total", "Repuestos", "Servicios", "KPIs (resto)"], index=0)
    with g_col3:
        g_tipo = st.selectbox("Tipo", ["$", "Q"], index=0)

    # dataset base para gestión (acumulado en corte)
    dg = build_acum(df[df["Semana_Num"] <= semana_corte].copy())  # gestión siempre corte, no depende del filtro principal
    dg["KPI_UP"] = dg["KPI"].astype(str).str.upper()

    if g_suc != "TODAS":
        dg = dg[dg["Sucursal"] == g_suc].copy()

    # filtrar tipo
    dg = dg[dg["Tipo_KPI"].astype(str).str.strip() == g_tipo].copy()

    # filtrar segmento
    if g_seg == "Repuestos":
        dg = dg[dg["KPI_UP"] == "REPUESTOS"].copy()
    elif g_seg == "Servicios":
        dg = dg[dg["KPI_UP"] == "SERVICIOS"].copy()
    elif g_seg == "KPIs (resto)":
        dg = dg[~dg["KPI_UP"].isin(["REPUESTOS", "SERVICIOS"])].copy()
    else:
        dg = dg.copy()

    # Tabla de desvíos principales (por gap)
    st.markdown("<hr/>", unsafe_allow_html=True)
    st.markdown("#### Top desvíos por impacto (Gap = Obj - Real)")

    # gap positivo = falta (peor), gap negativo = superávit
    dtop = dg.groupby(["KPI", "Categoria_KPI"], as_index=False).agg(
        Real=("Real_Acum", "sum"),
        Obj=("Obj_Acum", "sum"),
    )
    dtop["Gap"] = dtop["Obj"] - dtop["Real"]
    dtop["Cumpl"] = dtop.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    dtop["Estado"] = dtop["Cumpl"].apply(estado_from_cumpl)

    dtop = dtop.sort_values("Gap", ascending=False)

    show_n = st.slider("Cantidad de ítems", min_value=5, max_value=30, value=12, step=1)

    show_df = dtop.head(show_n).copy()
    if g_tipo == "$":
        show_df["Real"] = show_df["Real"].apply(money)
        show_df["Obj"] = show_df["Obj"].apply(money)
        show_df["Gap"] = show_df["Gap"].apply(money)
    else:
        show_df["Real"] = show_df["Real"].apply(num)
        show_df["Obj"] = show_df["Obj"].apply(num)
        show_df["Gap"] = show_df["Gap"].apply(num)

    show_df["Cumpl"] = show_df["Cumpl"].apply(pct)

    st.dataframe(show_df[["KPI", "Categoria_KPI", "Real", "Obj", "Gap", "Cumpl", "Estado"]], use_container_width=True, hide_index=True)

    # Waterfall drivers (por categoría)
    st.markdown("<hr/>", unsafe_allow_html=True)
    st.markdown("#### Drivers del desvío — Waterfall")

    wf = dtop.copy()
    wf = wf.sort_values("Gap", ascending=False).head(15)  # top drivers

    if wf.empty:
        st.info("Sin drivers para este filtro.")
    else:
        fig = go.Figure(go.Waterfall(
            x=wf["KPI"].astype(str) + " | " + wf["Categoria_KPI"].astype(str),
            y=wf["Obj"].astype(float) - wf["Real"].astype(float),
            measure=["relative"] * len(wf)
        ))
        fig.update_layout(height=420, margin=dict(l=20, r=20, t=10, b=20))
        st.plotly_chart(fig, use_container_width=True)

    # Narrativa automática
    st.markdown("<hr/>", unsafe_allow_html=True)
    st.markdown("#### Lectura automática (para Dirección)")

    if dtop.empty:
        st.info("No hay datos para narrar.")
    else:
        p = dtop.iloc[0]
        if g_tipo == "$":
            st.info(
                f"Principal desvío ({g_seg}, {g_tipo}) → **{p['KPI']} / {p['Categoria_KPI']}** "
                f"con impacto **{money(p['Obj']-p['Real'])}** "
                f"(Obj {money(p['Obj'])} vs Real {money(p['Real'])}) — Cumpl {pct(p['Cumpl'])}."
            )
        else:
            st.info(
                f"Principal desvío ({g_seg}, {g_tipo}) → **{p['KPI']} / {p['Categoria_KPI']}** "
                f"con impacto **{num(p['Obj']-p['Real'])}** "
                f"(Obj {num(p['Obj'])} vs Real {num(p['Real'])}) — Cumpl {pct(p['Cumpl'])}."
            )
