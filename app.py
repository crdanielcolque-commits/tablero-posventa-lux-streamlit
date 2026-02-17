# ==========================
# TABLERO POSVENTA V2.1
# Dirección + Diagnóstico Inteligente
# ==========================

import re
import unicodedata
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
# HELPERS
# ==========================

def safe_ratio(n, d):
    if d is None or pd.isna(d) or d == 0:
        return None
    return n / d

def money(x):
    if pd.isna(x): return "—"
    return f"${x:,.0f}".replace(",", ".")

def pct(x):
    if pd.isna(x): return "—"
    return f"{x*100:.1f}%"

def estado(c):
    if pd.isna(c): return "—"
    if c >= 1: return "Verde"
    if c >= 0.9: return "Amarillo"
    return "Rojo"

# ==========================
# CARGA
# ==========================

@st.cache_data(ttl=300)
def load():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df = pd.read_excel(EXCEL_LOCAL)
    return df

df = load()

# ==========================
# NORMALIZACIÓN BÁSICA
# ==========================

df["Semana_Num"] = df["Semana"].str.extract("(\d+)").astype(int)
df["Real"] = df["Real_$"]
df["Obj"] = df["Objetivo_$"]

# ==========================
# SIDEBAR
# ==========================

st.sidebar.title("Filtros")

semana = st.sidebar.selectbox("Semana corte", sorted(df["Semana_Num"].unique()))
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS"] + sorted(df["Sucursal"].unique()))

df = df[df["Semana_Num"] <= semana]

if sucursal != "TODAS":
    df = df[df["Sucursal"] == sucursal]

# ==========================
# ACUMULADO
# ==========================

df_acum = df.groupby(["Sucursal","KPI","Categoria_KPI"], as_index=False).agg(
    Real=("Real","sum"),
    Obj=("Obj","sum")
)

df_acum["Cumpl"] = df_acum.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
df_acum["Gap"] = df_acum["Obj"] - df_acum["Real"]

# ==========================
# RESUMEN EJECUTIVO
# ==========================

st.title("Tablero Posventa — V2.1 Diagnóstico Inteligente")

rep = df_acum[df_acum["KPI"]=="Repuestos"]
srv = df_acum[df_acum["KPI"]=="Servicios"]

rep_real = rep["Real"].sum()
rep_obj = rep["Obj"].sum()
rep_c = safe_ratio(rep_real, rep_obj)
rep_gap = rep_obj - rep_real

srv_real = srv["Real"].sum()
srv_obj = srv["Obj"].sum()
srv_c = safe_ratio(srv_real, srv_obj)
srv_gap = srv_obj - srv_real

tot_real = rep_real + srv_real
tot_obj = rep_obj + srv_obj
tot_c = safe_ratio(tot_real, tot_obj)
tot_gap = tot_obj - tot_real

c1,c2,c3 = st.columns(3)

with c1:
    st.metric("Repuestos", pct(rep_c), money(rep_gap))

with c2:
    st.metric("Servicios", pct(srv_c), money(srv_gap))

with c3:
    st.metric("Total Posventa", pct(tot_c), money(tot_gap))

st.divider()

# ==========================
# WATERFALL DRIVERS
# ==========================

st.subheader("Drivers del Desvío — Waterfall")

seg = st.selectbox("Segmento", ["Repuestos","Servicios","Total"])

if seg == "Total":
    wf = df_acum[df_acum["KPI"].isin(["Repuestos","Servicios"])]
else:
    wf = df_acum[df_acum["KPI"]==seg]

wf = wf.sort_values("Gap", ascending=False)

fig = go.Figure(go.Waterfall(
    x=wf["Categoria_KPI"],
    y=wf["Gap"],
    measure=["relative"]*len(wf)
))

fig.update_layout(height=400)
st.plotly_chart(fig, use_container_width=True)

# ==========================
# TOP SUCURSALES
# ==========================

st.subheader("Sucursales que explican el desvío")

top_suc = df.groupby(["Sucursal"], as_index=False).agg(
    Real=("Real","sum"),
    Obj=("Obj","sum")
)

top_suc["Gap"] = top_suc["Obj"] - top_suc["Real"]
top_suc = top_suc.sort_values("Gap", ascending=False)

fig2 = px.bar(top_suc, x="Gap", y="Sucursal", orientation="h")
st.plotly_chart(fig2, use_container_width=True)

# ==========================
# NARRATIVA AUTOMÁTICA
# ==========================

if len(wf) > 0:
    principal = wf.iloc[0]
    st.info(
        f"El principal driver del desvío es {principal['Categoria_KPI']} "
        f"con un impacto de {money(principal['Gap'])}."
    )
