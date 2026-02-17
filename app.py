# ============================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado) v2.2.1
# FIX CRÍTICO: parse de números AR (coma decimal / $ / miles)
# ============================================================

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

# ---------------------------
# CONFIG
# ---------------------------
st.set_page_config(page_title="Tablero Posventa", layout="wide")

DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# ---------------------------
# HELPERS
# ---------------------------
def parse_semana_num(series: pd.Series) -> pd.Series:
    """Convierte 'Semana 1', '1', '1.0', 'Semana 1.0' a Int64 robusto."""
    s = series.astype(str).str.strip()
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]
    num = num.str.replace(",", ".", regex=False)
    numf = pd.to_numeric(num, errors="coerce")
    return np.floor(numf).astype("Int64")

def to_num_ar(x):
    """
    Convierte números en formato AR a float:
    - $ 1.234.567,89  -> 1234567.89
    - 9808131,76      -> 9808131.76
    - 0.00 / $0.00    -> 0.0
    - 30,64% / 30.64% -> 0.3064 (si viene con %)
    """
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return np.nan
    s = str(x).strip()
    if s == "" or s.lower() in {"nan", "none"}:
        return np.nan

    is_pct = "%" in s
    s = s.replace("%", "")

    # quitar símbolos y espacios comunes
    s = (
        s.replace("$", "")
         .replace("AR$", "")
         .replace(" ", "")
         .replace("\u00A0", "")
    )

    # normalizar separadores:
    # si hay coma, asumimos coma decimal y punto miles
    # 1.234.567,89 -> 1234567.89
    if "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    # si no hay coma, puede venir ya en formato punto decimal (ok)

    try:
        v = float(s)
        if is_pct:
            v = v / 100.0
        return v
    except Exception:
        return np.nan

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

def qty(x):
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

def estado(c):
    if c is None or pd.isna(c):
        return "—"
    if c >= 1:
        return "Verde"
    if c >= 0.9:
        return "Amarillo"
    return "Rojo"

def badge_estado_html(est):
    color = {"Verde":"#198754", "Amarillo":"#d39e00", "Rojo":"#dc3545", "—":"#6c757d"}.get(est, "#6c757d")
    bg    = {"Verde":"#d1e7dd", "Amarillo":"#fff3cd", "Rojo":"#f8d7da", "—":"#e9ecef"}.get(est, "#e9ecef")
    return f"""
    <span style="display:inline-block;padding:4px 10px;border-radius:999px;
                 background:{bg};color:{color};font-weight:700;font-size:12px;border:1px solid {color}33;">
        {est.upper()}
    </span>
    """

def card_html(title, value, sub, estado_txt=None):
    estado_block = ""
    if estado_txt is not None:
        estado_block = f"<div style='margin-top:8px;'>{badge_estado_html(estado_txt)}</div>"
    return f"""
    <div style="border:1px solid #eee;border-radius:14px;padding:16px;background:#fff;
                box-shadow:0 2px 10px rgba(0,0,0,0.04);">
        <div style="font-size:12px;color:#6c757d;font-weight:700;">{title}</div>
        <div style="font-size:28px;font-weight:800;margin-top:6px;">{value}</div>
        <div style="font-size:12px;color:#6c757d;margin-top:6px;">{sub}</div>
        {estado_block}
    </div>
    """

# ---------------------------
# LOAD
# ---------------------------
@st.cache_data(ttl=300)
def load_from_drive():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df0 = pd.read_excel(EXCEL_LOCAL)
    df0 = df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")]
    return df0

df = load_from_drive()

# ---------------------------
# VALIDACIÓN
# ---------------------------
required = [
    "Fecha","Semana","Sucursal","KPI","Categoria_KPI","Tipo_KPI",
    "Real_$","Real_Q","Objetivo_$","Objetivo_Q","Cumplimiento_%","Estado"
]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error("❌ Faltan columnas requeridas en el Excel:")
    st.write(missing)
    st.stop()

# ---------------------------
# NORMALIZACIÓN
# ---------------------------
df["Semana_Num"] = parse_semana_num(df["Semana"])
df = df[~df["Semana_Num"].isna()].copy()

# Parse AR numérico (CRÍTICO)
for c in ["Real_$","Costo_$","Margen_$","Margen_%","Real_Q","Objetivo_$","Objetivo_Q","Cumplimiento_%"]:
    if c in df.columns:
        df[c] = df[c].apply(to_num_ar)

# Limpieza strings
df["KPI"] = df["KPI"].astype(str).str.strip()
df["Categoria_KPI"] = df["Categoria_KPI"].astype(str).str.strip()
df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()

# Calcular Real/Obj según Tipo
def build_real_obj(row):
    if row["Tipo_KPI"] == "$":
        return row["Real_$"], row["Objetivo_$"]
    else:
        return row["Real_Q"], row["Objetivo_Q"]

tmp = df.apply(build_real_obj, axis=1, result_type="expand")
df["Real_val"] = pd.to_numeric(tmp[0], errors="coerce").fillna(0.0)
df["Obj_val"]  = pd.to_numeric(tmp[1], errors="coerce").fillna(0.0)
df["Cumpl_calc"] = df.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)

# ---------------------------
# SIDEBAR
# ---------------------------
st.sidebar.markdown("## Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
default_sem = 1 if 1 in semanas else (min(semanas) if semanas else 1)
default_idx = semanas.index(default_sem) if default_sem in semanas else 0

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

st.sidebar.markdown("---")
st.sidebar.markdown("### Cálculo")
show_obj0 = st.sidebar.checkbox("Incluir filas con Obj=0 (puede distorsionar %)", value=False)

# Corte
df_cut = df[df["Semana_Num"] <= semana_corte].copy()
if sucursal != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal].copy()

def apply_obj0_filter(d):
    if show_obj0:
        return d.copy()
    return d[d["Obj_val"] > 0].copy()

# ---------------------------
# Filtros P&L aperturas
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("## Incluir variables (P&L)")

rep_open = sorted(df_cut[(df_cut["KPI"].str.upper()=="REPUESTOS") & (df_cut["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
srv_open = sorted(df_cut[(df_cut["KPI"].str.upper()=="SERVICIOS") & (df_cut["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())

rep_sel = st.sidebar.multiselect("Repuestos: aperturas incluidas", rep_open, default=rep_open)
srv_sel = st.sidebar.multiselect("Servicios: aperturas incluidas", srv_open, default=srv_open)

# ---------------------------
# UI
# ---------------------------
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ============================================================
# TAB 1 — P&L
# ============================================================
with tab1:
    st.markdown("## 🧩 P&L — Macro → Micro")
    st.markdown("---")

    d_pl = df_cut[df_cut["Tipo_KPI"]=="$"].copy()

    d_rep = d_pl[d_pl["KPI"].str.upper()=="REPUESTOS"].copy()
    d_rep = d_rep[d_rep["Categoria_KPI"].isin(rep_sel)].copy()
    d_rep = apply_obj0_filter(d_rep)

    d_srv = d_pl[d_pl["KPI"].str.upper()=="SERVICIOS"].copy()
    d_srv = d_srv[d_srv["Categoria_KPI"].isin(srv_sel)].copy()
    d_srv = apply_obj0_filter(d_srv)

    def macro_cards(d):
        real = d["Real_val"].sum()
        obj  = d["Obj_val"].sum()
        c    = safe_ratio(real, obj)
        st.markdown(
            card_html("Cumplimiento (Acum.)", pct(c), f"Real {money(real)} | Obj {money(obj)}", estado(c)),
            unsafe_allow_html=True
        )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        macro_cards(d_rep)
    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        macro_cards(d_srv)

    st.markdown("---")
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    def micro_aperturas(d):
        g = d.groupby("Categoria_KPI", as_index=False).agg(
            Real=("Real_val","sum"),
            Obj=("Obj_val","sum")
        )
        g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        g = g[~g["Cumpl"].isna()].sort_values("Cumpl", ascending=False)
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
        return g

    l, r = st.columns(2)
    with l:
        st.markdown("**Repuestos — por apertura**")
        g = micro_aperturas(d_rep)
        if g.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with r:
        st.markdown("**Servicios — por apertura**")
        g = micro_aperturas(d_srv)
        if g.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 2 — KPIs resto
# ============================================================
with tab2:
    st.markdown("## 📌 KPIs (resto) — Macro → Micro")
    st.markdown("---")

    resto = df_cut[~df_cut["KPI"].str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()
    resto = apply_obj0_filter(resto)

    kpis_resto = sorted(resto["KPI"].unique().tolist())
    if not kpis_resto:
        st.info("No hay KPIs (resto) con Obj>0 en este corte.")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto)

        x = resto[resto["KPI"]==kpi_sel].copy()
        tipos = sorted(x["Tipo_KPI"].unique().tolist())

        for t in tipos:
            xt = x[x["Tipo_KPI"]==t].copy()
            real = xt["Real_val"].sum()
            obj  = xt["Obj_val"].sum()
            c    = safe_ratio(real, obj)

            st.markdown(
                card_html(
                    f"{kpi_sel} ({t}) — Cumplimiento (Acum.)",
                    pct(c),
                    f"Real {(money(real) if t=='$' else qty(real))} | Obj {(money(obj) if t=='$' else qty(obj))}",
                    estado(c)
                ),
                unsafe_allow_html=True
            )

            g = xt.groupby("Sucursal", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
            g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
            g = g[~g["Cumpl"].isna()].sort_values("Cumpl", ascending=False)
            g["label"] = g.apply(
                lambda r: f"{pct(r['Cumpl'])} | {(money(r['Real']) if t=='$' else qty(r['Real']))}/{(money(r['Obj']) if t=='$' else qty(r['Obj']))}",
                axis=1
            )

            st.markdown("### Ranking por sucursal — este KPI")
            if g.empty:
                st.info("Sin ranking (Obj=0 o sin datos).")
            else:
                fig = px.bar(g, x="Cumpl", y="Sucursal", orientation="h", text="label")
                fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
                fig.update_traces(textposition="inside")
                st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 3 — Gestión
# ============================================================
with tab3:
    st.markdown("## 🧪 Gestión (desvíos)")
    st.markdown("---")

    suc_g = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)

    d = df[df["Semana_Num"] <= semana_corte].copy()
    if suc_g != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_g].copy()

    d = apply_obj0_filter(d)

    g = d.groupby(["KPI","Categoria_KPI","Tipo_KPI"], as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
    g["Gap"] = g["Obj"] - g["Real"]
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g.sort_values("Gap", ascending=False)

    st.markdown("### Top desvíos (Gap) — Obj - Real")
    if g.empty:
        st.info("Sin desvíos (o todo Obj=0).")
    else:
        show_n = st.selectbox("Top N desvíos", [10,20,30,50], index=1)
        gg = g.head(show_n).copy()
        gg["key"] = gg["KPI"].astype(str) + " — " + gg["Categoria_KPI"].astype(str) + " (" + gg["Tipo_KPI"].astype(str) + ")"
        gg["label"] = gg.apply(lambda r: f"Gap {money(r['Gap']) if r['Tipo_KPI']=='$' else qty(r['Gap'])} | {pct(r['Cumpl'])}", axis=1)

        fig = px.bar(gg, x="Gap", y="key", orientation="h", text="label")
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10))
        fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)
