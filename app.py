# ==========================
# TABLERO POSVENTA V2.1 (FIX)
# Diagnóstico + Waterfall
# ==========================

import re
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
# HELPERS
# ==========================

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
    if c >= 1: return "Verde"
    if c >= 0.9: return "Amarillo"
    return "Rojo"

def parse_semana_num(series: pd.Series) -> pd.Series:
    """
    Convierte una columna Semana a número robustamente.
    Acepta: "Semana 1", "1", "1.0", "Semana 1.0", etc.
    Devuelve Int64 (nullable).
    """
    s = series.astype(str).str.strip()

    # Buscar primer número (entero o decimal)
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]

    # Normalizar coma -> punto y convertir a float
    num = num.str.replace(",", ".", regex=False)

    # to_numeric robusto
    numf = pd.to_numeric(num, errors="coerce")

    # Si vino como 1.0 lo llevamos a 1
    numi = np.floor(numf).astype("Int64")

    return numi

# ==========================
# CARGA
# ==========================

@st.cache_data(ttl=300)
def load():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)
    df0 = pd.read_excel(EXCEL_LOCAL)

    # Limpieza de columnas por si aparece "Unnamed"
    df0 = df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")]

    return df0

df = load()

# ==========================
# VALIDACIÓN MINIMA
# ==========================
required = ["Semana","Sucursal","KPI","Categoria_KPI","Real_$","Objetivo_$"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error("Faltan columnas requeridas en el Excel:")
    st.write(missing)
    st.stop()

# ==========================
# NORMALIZACIÓN
# ==========================

df["Semana_Num"] = parse_semana_num(df["Semana"])

# Si hay filas sin semana, las descartamos (o podrías asignar 0)
df = df[~df["Semana_Num"].isna()].copy()

df["Real"] = pd.to_numeric(df["Real_$"], errors="coerce").fillna(0.0)
df["Obj"]  = pd.to_numeric(df["Objetivo_$"], errors="coerce").fillna(0.0)

# ==========================
# SIDEBAR
# ==========================

st.sidebar.title("Filtros")

semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
if not semanas:
    st.error("No se encontraron semanas válidas en la columna 'Semana'.")
    st.stop()

# Default: Semana 1 si existe, sino la menor
default_sem = 1 if 1 in semanas else semanas[0]
default_idx = semanas.index(default_sem)

semana = st.sidebar.selectbox("Semana corte", semanas, index=default_idx)
sucursales = sorted(df["Sucursal"].dropna().unique().tolist())
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS"] + sucursales)

# Corte acumulado
df = df[df["Semana_Num"] <= semana].copy()

# Filtro sucursal (si aplica)
if sucursal != "TODAS":
    df = df[df["Sucursal"] == sucursal].copy()

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
# HEADER
# ==========================

st.title("Tablero Posventa — V2.1 (Diagnóstico Inteligente)")

# ==========================
# KPIs MACRO
# ==========================

rep = df_acum[df_acum["KPI"].astype(str).str.upper()=="REPUESTOS"]
srv = df_acum[df_acum["KPI"].astype(str).str.upper()=="SERVICIOS"]

rep_real = rep["Real"].sum()
rep_obj  = rep["Obj"].sum()
rep_c    = safe_ratio(rep_real, rep_obj)
rep_gap  = rep_obj - rep_real

srv_real = srv["Real"].sum()
srv_obj  = srv["Obj"].sum()
srv_c    = safe_ratio(srv_real, srv_obj)
srv_gap  = srv_obj - srv_real

tot_real = rep_real + srv_real
tot_obj  = rep_obj + srv_obj
tot_c    = safe_ratio(tot_real, tot_obj)
tot_gap  = tot_obj - tot_real

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
    wf = df_acum[df_acum["KPI"].astype(str).str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()
else:
    wf = df_acum[df_acum["KPI"].astype(str).str.upper()==seg.upper()].copy()

# Orden por impacto (Gap)
wf = wf.sort_values("Gap", ascending=False)

if wf.empty:
    st.warning("No hay datos para el segmento seleccionado.")
else:
    fig = go.Figure(go.Waterfall(
        x=wf["Categoria_KPI"].astype(str),
        y=wf["Gap"].astype(float),
        measure=["relative"]*len(wf)
    ))
    fig.update_layout(height=420, margin=dict(l=20,r=20,t=10,b=20))
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# ==========================
# TOP SUCURSALES (GAP)
# ==========================

st.subheader("Sucursales que explican el desvío (Gap)")

top_suc = df.groupby(["Sucursal"], as_index=False).agg(
    Real=("Real","sum"),
    Obj=("Obj","sum")
)
top_suc["Gap"] = top_suc["Obj"] - top_suc["Real"]
top_suc = top_suc.sort_values("Gap", ascending=False)

if top_suc.empty:
    st.warning("Sin datos para ranking de sucursales.")
else:
    fig2 = px.bar(top_suc, x="Gap", y="Sucursal", orientation="h")
    fig2.update_layout(height=420, margin=dict(l=20,r=20,t=10,b=20))
    st.plotly_chart(fig2, use_container_width=True)

# ==========================
# NARRATIVA AUTOMÁTICA
# ==========================

st.subheader("Lectura automática (para Dirección)")

if wf.empty:
    st.info("No hay drivers para narrar en este corte.")
else:
    principal = wf.iloc[0]
    st.info(
        f"El principal driver del desvío en **{seg}** es **{principal['Categoria_KPI']}** "
        f"con impacto de **{money(principal['Gap'])}** (Obj {money(principal['Obj'])} vs Real {money(principal['Real'])})."
    )
