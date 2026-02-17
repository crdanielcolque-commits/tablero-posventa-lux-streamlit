# ============================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado) v2.2
# Robusto: Obj=0, filtros P&L por aperturas, ranking + labels,
# 3 tabs (P&L / KPIs resto / Gestión)
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

DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"  # <- tu Google Sheet ID
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
    numi = np.floor(numf).astype("Int64")
    return numi

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

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

def chip_html(text):
    # Verde suave (no alerta)
    return f"""
    <span style="
        display:inline-block;
        padding:6px 10px;
        border-radius:12px;
        background:#DFF3E6;
        color:#0F5132;
        font-weight:600;
        font-size:12px;
        border:1px solid #BFE6CF;">
        {text}
    </span>
    """

def badge_estado_html(est):
    color = {"Verde":"#198754", "Amarillo":"#d39e00", "Rojo":"#dc3545", "—":"#6c757d"}.get(est, "#6c757d")
    bg    = {"Verde":"#d1e7dd", "Amarillo":"#fff3cd", "Rojo":"#f8d7da", "—":"#e9ecef"}.get(est, "#e9ecef")
    return f"""
    <span style="
        display:inline-block;
        padding:4px 10px;
        border-radius:999px;
        background:{bg};
        color:{color};
        font-weight:700;
        font-size:12px;
        border:1px solid {color}33;">
        {est.upper()}
    </span>
    """

def card_html(title, value, sub, estado_txt=None):
    estado_block = ""
    if estado_txt is not None:
        estado_block = f"<div style='margin-top:8px;'>{badge_estado_html(estado_txt)}</div>"
    return f"""
    <div style="
        border:1px solid #eee;
        border-radius:14px;
        padding:16px 16px;
        background:#fff;
        box-shadow:0 2px 10px rgba(0,0,0,0.04);">
        <div style="font-size:12px;color:#6c757d;font-weight:700;letter-spacing:.2px;">{title}</div>
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
# VALIDACIÓN MÍNIMA
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

# Numerización
for c in ["Real_$","Costo_$","Margen_$","Margen_%","Real_Q","Objetivo_$","Objetivo_Q","Cumplimiento_%"]:
    if c in df.columns:
        df[c] = to_num(df[c])

df["KPI"] = df["KPI"].astype(str).str.strip()
df["Categoria_KPI"] = df["Categoria_KPI"].astype(str).str.strip()
df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()

# Para cálculos: definimos columnas "Real_val" y "Obj_val" según Tipo_KPI
def build_real_obj(row):
    if row["Tipo_KPI"] == "$":
        return row["Real_$"], row["Objetivo_$"]
    else:
        return row["Real_Q"], row["Objetivo_Q"]

tmp = df.apply(build_real_obj, axis=1, result_type="expand")
df["Real_val"] = to_num(tmp[0]).fillna(0.0)
df["Obj_val"]  = to_num(tmp[1]).fillna(0.0)

# Cumplimiento calculado robusto
df["Cumpl_calc"] = df.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)

# ---------------------------
# SIDEBAR (obligatorio)
# ---------------------------
st.sidebar.markdown("## Filtros obligatorios")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
default_sem = 1 if 1 in semanas else (min(semanas) if semanas else 1)
default_idx = semanas.index(default_sem) if default_sem in semanas else 0

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)

sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

st.sidebar.markdown("---")

# Corte semanal acumulado
df_cut = df[df["Semana_Num"] <= semana_corte].copy()

# Filtro sucursal
if sucursal != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal].copy()

# Filtro: incluir/excluir filas Obj=0
st.sidebar.markdown("### Cálculo")
show_obj0 = st.sidebar.checkbox("Incluir filas con Obj=0 (puede distorsionar %)", value=False)

# ---------------------------
# Filtros P&L (aperturas incluidas)
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("## Incluir variables (P&L)")

# Aperturas Repuestos y Servicios (Tipo $)
rep_open = sorted(df_cut[(df_cut["KPI"].str.upper()=="REPUESTOS") & (df_cut["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
srv_open = sorted(df_cut[(df_cut["KPI"].str.upper()=="SERVICIOS") & (df_cut["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())

rep_sel = st.sidebar.multiselect("Repuestos: aperturas incluidas", rep_open, default=rep_open)
srv_sel = st.sidebar.multiselect("Servicios: aperturas incluidas", srv_open, default=srv_open)

# ---------------------------
# TAB LAYOUT
# ---------------------------
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ============================================================
# UTIL: agregaciones
# ============================================================
def apply_obj0_filter(d):
    if show_obj0:
        return d.copy()
    return d[d["Obj_val"] > 0].copy()

def agg_segment(d, kpi_name, tipo):
    x = d[(d["KPI"].str.upper()==kpi_name.upper()) & (d["Tipo_KPI"]==tipo)].copy()
    x = apply_obj0_filter(x)
    real = x["Real_val"].sum()
    obj  = x["Obj_val"].sum()
    c    = safe_ratio(real, obj)
    return real, obj, c, x

def barh_rank(df_rank, x, y, title, label_mode="pct"):
    if df_rank.empty:
        st.info("Sin datos para mostrar.")
        return
    fig = px.bar(df_rank, x=x, y=y, orientation="h", text="label")
    fig.update_layout(height=420, margin=dict(l=10, r=10, t=40, b=10), title=title)
    fig.update_traces(textposition="inside")
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 1 — P&L Macro → Micro
# ============================================================
with tab1:
    st.markdown("## 🧩 P&L — Macro → Micro")
    st.markdown(
        "<div style='color:#6c757d'>Primero lo macro (Repuestos / Servicios). Después el micro (aperturas y ranking por sucursal).</div>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    # Datos P&L filtrados por aperturas seleccionadas
    d_pl = df_cut[(df_cut["Tipo_KPI"]=="$")].copy()

    d_rep = d_pl[d_pl["KPI"].str.upper()=="REPUESTOS"].copy()
    d_rep = d_rep[d_rep["Categoria_KPI"].isin(rep_sel)].copy()

    d_srv = d_pl[d_pl["KPI"].str.upper()=="SERVICIOS"].copy()
    d_srv = d_srv[d_srv["Categoria_KPI"].isin(srv_sel)].copy()

    # Macro cards
    rep_real, rep_obj, rep_c, rep_df = agg_segment(d_rep, "REPUESTOS", "$")
    srv_real, srv_obj, srv_c, srv_df = agg_segment(d_srv, "SERVICIOS", "$")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        st.markdown(
            card_html(
                "Cumplimiento (Acum.)",
                pct(rep_c),
                f"Real {money(rep_real)} | Obj {money(rep_obj)}",
                estado(rep_c)
            ),
            unsafe_allow_html=True
        )
    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        st.markdown(
            card_html(
                "Cumplimiento (Acum.)",
                pct(srv_c),
                f"Real {money(srv_real)} | Obj {money(srv_obj)}",
                estado(srv_c)
            ),
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    # Micro aperturas (Repuestos / Servicios)
    def micro_by_open(d, kpi_upper):
        m = d[d["KPI"].str.upper()==kpi_upper].copy()
        m = apply_obj0_filter(m)
        g = m.groupby(["Categoria_KPI"], as_index=False).agg(
            Real=("Real_val","sum"),
            Obj=("Obj_val","sum")
        )
        g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
        g = g.sort_values("Cumpl", ascending=False)
        return g

    left, right = st.columns(2)
    with left:
        st.markdown("**Repuestos — por apertura**")
        g_rep = micro_by_open(d_rep, "REPUESTOS")
        if g_rep.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g_rep, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with right:
        st.markdown("**Servicios — por apertura**")
        g_srv = micro_by_open(d_srv, "SERVICIOS")
        if g_srv.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g_srv, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("## 🎯 Micro — ranking sucursal + apertura")

    topn = st.selectbox("Top N", [5,10,15,20], index=1)
    show_obj0_rank = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=True)

    def micro_sucursal_apertura(d, kpi_upper, aperturas_sel):
        x = d[(d["KPI"].str.upper()==kpi_upper) & (d["Tipo_KPI"]=="$")].copy()
        x = x[x["Categoria_KPI"].isin(aperturas_sel)].copy()
        x = apply_obj0_filter(x)

        g = x.groupby(["Sucursal","Categoria_KPI"], as_index=False).agg(
            Real=("Real_val","sum"),
            Obj=("Obj_val","sum")
        )
        g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)

        # limpiar NaN
        g = g[~g["Cumpl"].isna()].copy()

        # si no quiere ver 0% (cuando Obj>0 y Real=0)
        if not show_obj0_rank:
            g = g[~((g["Obj"]>0) & (g["Real"]==0))].copy()

        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
        g["key"] = g["Sucursal"].astype(str) + " — " + g["Categoria_KPI"].astype(str)
        g = g.sort_values("Cumpl", ascending=False).head(topn)
        return g

    l2, r2 = st.columns(2)
    with l2:
        rep_pick = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + rep_sel)
        rep_use = rep_sel if rep_pick == "Todas las aperturas" else [rep_pick]
        g = micro_sucursal_apertura(df_cut, "REPUESTOS", rep_use)
        st.markdown("**Repuestos — sucursal + apertura (micro)**")
        if g.empty:
            st.info("Sin datos para este ranking.")
        else:
            fig = px.bar(g, x="Cumpl", y="key", orientation="h", text="label")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with r2:
        srv_pick = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + srv_sel)
        srv_use = srv_sel if srv_pick == "Todas las aperturas" else [srv_pick]
        g = micro_sucursal_apertura(df_cut, "SERVICIOS", srv_use)
        st.markdown("**Servicios — sucursal + apertura (micro)**")
        if g.empty:
            st.info("Sin datos para este ranking.")
        else:
            fig = px.bar(g, x="Cumpl", y="key", orientation="h", text="label")
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("## 🏁 Ranking por sucursal (macro) + micro por apertura")

    def rank_sucursal_macro(d, kpi_upper):
        x = d[(d["KPI"].str.upper()==kpi_upper) & (d["Tipo_KPI"]=="$")].copy()
        x = apply_obj0_filter(x)
        g = x.groupby(["Sucursal"], as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
        g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        g = g[~g["Cumpl"].isna()].copy()
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
        g = g.sort_values("Cumpl", ascending=False)
        return g

    l3, r3 = st.columns(2)
    with l3:
        st.markdown("**Repuestos — por sucursal (macro)**")
        rk = rank_sucursal_macro(d_rep, "REPUESTOS")
        if rk.empty:
            st.info("Sin datos para ranking.")
        else:
            fig = px.bar(rk, x="Cumpl", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

            # micro: apertura dentro de sucursal (repuestos)
            suc_pick = st.selectbox("Ver micro de Repuestos para sucursal:", ["(elegir)"] + rk["Sucursal"].tolist())
            if suc_pick != "(elegir)":
                micro = d_rep[(d_rep["Sucursal"]==suc_pick) & (d_rep["KPI"].str.upper()=="REPUESTOS")].copy()
                micro = apply_obj0_filter(micro)
                g = micro.groupby("Categoria_KPI", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
                g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
                g = g[~g["Cumpl"].isna()].sort_values("Cumpl", ascending=False)
                g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
                st.markdown("**Micro (aperturas) — Repuestos**")
                if g.empty:
                    st.info("Sin micro (Obj=0 o sin datos).")
                else:
                    fig = px.bar(g, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
                    fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
                    fig.update_traces(textposition="inside")
                    st.plotly_chart(fig, use_container_width=True)

    with r3:
        st.markdown("**Servicios — por sucursal (macro)**")
        rk = rank_sucursal_macro(d_srv, "SERVICIOS")
        if rk.empty:
            st.info("Sin datos para ranking.")
        else:
            fig = px.bar(rk, x="Cumpl", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

            # micro: apertura dentro de sucursal (servicios)
            suc_pick = st.selectbox("Ver micro de Servicios para sucursal:", ["(elegir)"] + rk["Sucursal"].tolist(), key="srv_micro_pick")
            if suc_pick != "(elegir)":
                micro = d_srv[(d_srv["Sucursal"]==suc_pick) & (d_srv["KPI"].str.upper()=="SERVICIOS")].copy()
                micro = apply_obj0_filter(micro)
                g = micro.groupby("Categoria_KPI", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
                g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
                g = g[~g["Cumpl"].isna()].sort_values("Cumpl", ascending=False)
                g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
                st.markdown("**Micro (aperturas) — Servicios**")
                if g.empty:
                    st.info("Sin micro (Obj=0 o sin datos).")
                else:
                    fig = px.bar(g, x="Cumpl", y="Categoria_KPI", orientation="h", text="label")
                    fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
                    fig.update_traces(textposition="inside")
                    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 2 — KPIs resto (no Repuestos/Servicios)
# ============================================================
with tab2:
    st.markdown("## 📌 KPIs (resto) — Macro → Micro")
    st.markdown("<div style='color:#6c757d'>KPIs que no son Repuestos/Servicios (ej: Accesorios, Neumáticos, Campañas, CPUS, etc.).</div>", unsafe_allow_html=True)
    st.markdown("---")

    resto = df_cut[~df_cut["KPI"].str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()
    resto = apply_obj0_filter(resto)

    kpis_resto = sorted(resto["KPI"].unique().tolist())
    if not kpis_resto:
        st.info("No hay KPIs (resto) con Obj>0 en este corte.")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto)

        x = resto[resto["KPI"]==kpi_sel].copy()

        # Macro KPI seleccionado (según Tipo: si hay mezclas, mostramos separado)
        tipos = sorted(x["Tipo_KPI"].unique().tolist())
        for t in tipos:
            xt = x[x["Tipo_KPI"]==t].copy()
            real = xt["Real_val"].sum()
            obj  = xt["Obj_val"].sum()
            c    = safe_ratio(real, obj)
            est  = estado(c)
            st.markdown(
                card_html(
                    f"{kpi_sel} ({t}) — Cumplimiento (Acum.)",
                    pct(c),
                    f"Real {money(real) if t=='$' else qty(real)} | Obj {money(obj) if t=='$' else qty(obj)}",
                    est
                ),
                unsafe_allow_html=True
            )

            # Ranking por sucursal para ese KPI
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
# TAB 3 — Gestión (desvíos) + filtro sucursal interno
# ============================================================
with tab3:
    st.markdown("## 🧪 Gestión (desvíos)")
    st.markdown("<div style='color:#6c757d'>Vista de control: dónde está el gap (Obj-Real) y qué lo explica.</div>", unsafe_allow_html=True)
    st.markdown("---")

    # Filtro sucursal dentro de Gestión (pedido tuyo)
    suc_g = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)

    d = df[df["Semana_Num"] <= semana_corte].copy()
    if suc_g != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_g].copy()

    d = apply_obj0_filter(d)

    # Gap por KPI + Categoria
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
        gg["label"] = gg.apply(
            lambda r: f"Gap {money(r['Gap']) if r['Tipo_KPI']=='$' else qty(r['Gap'])} | {pct(r['Cumpl'])}",
            axis=1
        )
        gg["key"] = gg["KPI"].astype(str) + " — " + gg["Categoria_KPI"].astype(str) + f" ({gg['Tipo_KPI']})"

        fig = px.bar(gg, x="Gap", y="key", orientation="h", text="label")
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10))
        fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### Nota (para no volver a frustrarse)")
    st.info("Si ves % absurdos (ej: 3000%+), casi seguro hay filas con Objetivo=0 o muy chico. Por defecto este tablero las excluye del % para no distorsionar.")
