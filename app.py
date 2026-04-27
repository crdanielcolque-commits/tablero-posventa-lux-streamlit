import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(layout="wide")

# ============================================================
# HELPERS
# ============================================================
def safe_ratio(a, b):
    return a / b if b != 0 else 0

def money(x):
    try:
        return f"$ {int(x):,}".replace(",", ".")
    except:
        return "$ 0"

def pct(x):
    try:
        return f"{x:.1%}"
    except:
        return "-"

# ============================================================
# DATA
# ============================================================
@st.cache_data
def load_data():
    return pd.read_excel("base_posventa.xlsx")

df = load_data()

# ============================================================
# NORMALIZACIÓN (CLAVE)
# ============================================================
df["Real_val"] = np.where(df["Objetivo $"] != 0, df["Importe FC"], df["Q"])
df["Obj_val"] = np.where(df["Objetivo $"] != 0, df["Objetivo $"], df["Objetivo Q"])

# ============================================================
# FILTROS
# ============================================================
st.sidebar.title("Filtros")

semanas = sorted(df["Semana"].unique())
semana_sel = st.sidebar.selectbox("Semana corte", semanas)

df_cut = df[df["Semana"] <= semana_sel]

sucursales = df_cut["Sucursal"].unique()
suc_sel = st.sidebar.multiselect("Sucursal", sucursales, default=sucursales)

df_cut = df_cut[df_cut["Sucursal"].isin(suc_sel)]

# ============================================================
# TAB 1 - RESUMEN
# ============================================================
def tab_resumen(df):

    st.title("📊 Resumen Ejecutivo")

    total_real = df["Real_val"].sum()
    total_obj = df["Obj_val"].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Real", money(total_real))
    c2.metric("Objetivo", money(total_obj))
    c3.metric("Cumplimiento", pct(safe_ratio(total_real, total_obj)))

# ============================================================
# TAB NUEVO - GAP
# ============================================================
def tab_gap(df):

    st.title("🎯 Cierre de Mes – GAP a Objetivo")

    df_work = df.copy()

    # FILTROS INTELIGENTES
    rep_mask = df_work["Categoria KPI"].str.contains("Repuesto", case=False, na=False)
    srv_mask = df_work["Categoria KPI"].str.contains("Servicio", case=False, na=False)

    rep_opts = sorted(df_work.loc[rep_mask, "Categoria KPI"].unique())
    srv_opts = sorted(df_work.loc[srv_mask, "Categoria KPI"].unique())

    c1, c2 = st.columns(2)

    with c1:
        rep_sel = st.multiselect("Repuestos", rep_opts, default=rep_opts)

    with c2:
        srv_sel = st.multiselect("Servicios", srv_opts, default=srv_opts)

    incluir_cpus = st.checkbox("Incluir CPUs", True)
    incluir_neum = st.checkbox("Incluir Neumáticos", True)

    df_work = df_work[df_work["Categoria KPI"].isin(rep_sel + srv_sel)]

    if incluir_cpus:
        df_work = pd.concat([df_work, df[df["KPI"].str.contains("CPU", case=False, na=False)]])

    if incluir_neum:
        df_work = pd.concat([df_work, df[df["KPI"].str.contains("NEUM", case=False, na=False)]])

    # AGRUPACIÓN
    gap = df_work.groupby(["Sucursal", "Categoria KPI"], as_index=False).agg({
        "Real_val": "sum",
        "Obj_val": "sum"
    })

    gap["GAP"] = gap["Real_val"] - gap["Obj_val"]
    gap["Cumpl"] = gap.apply(lambda r: safe_ratio(r["Real_val"], r["Obj_val"]), axis=1)
    gap["Falta"] = np.where(gap["GAP"] < 0, abs(gap["GAP"]), 0)

    # KPIs
    total_real = gap["Real_val"].sum()
    total_obj = gap["Obj_val"].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Cumplimiento", pct(safe_ratio(total_real, total_obj)))
    c2.metric("GAP", money(total_real - total_obj))
    c3.metric("Falta", money(abs(total_real - total_obj)))

    # GAP POR SUCURSAL
    st.subheader("📍 GAP por sucursal")

    suc = gap.groupby("Sucursal", as_index=False)["GAP"].sum().sort_values("GAP")

    fig = px.bar(suc, x="GAP", y="Sucursal", orientation="h")
    fig.add_vline(x=0)
    st.plotly_chart(fig, use_container_width=True)

    # TABLA
    st.subheader("📊 Detalle")

    gap_show = gap.copy()
    gap_show["Real"] = gap_show["Real_val"].apply(money)
    gap_show["Objetivo"] = gap_show["Obj_val"].apply(money)
    gap_show["GAP"] = gap_show["GAP"].apply(money)
    gap_show["Falta"] = gap_show["Falta"].apply(money)
    gap_show["Cumpl"] = gap_show["Cumpl"].apply(pct)

    st.dataframe(gap_show, use_container_width=True)

    # PRIORIDADES
    st.subheader("⚔️ Prioridades de acción")

    top = gap.sort_values("Falta", ascending=False).head(10)

    fig2 = px.bar(top, x="Falta", y="Categoria KPI", color="Sucursal", orientation="h")
    st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# TABS
# ============================================================
tab1, tab2 = st.tabs([
    "📊 Resumen",
    "🎯 Cierre GAP"
])

with tab1:
    tab_resumen(df_cut)

with tab2:
    tab_gap(df_cut)
