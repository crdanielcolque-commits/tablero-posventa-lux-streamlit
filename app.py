import re
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

st.set_page_config(page_title="Tablero Posventa", layout="wide")

# =========================================================
# CONFIGURACIÓN GOOGLE DRIVE (YA COMPLETA)
# =========================================================
DRIVE_FILE_ID = "12J0gKlKfRvztWnInHg9XvT8vRq5oLlfQ"
EXCEL_LOCAL = "base_posventa.xlsx"

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def parse_semana_num(x):
    if pd.isna(x):
        return None
    m = re.search(r"(\d+)", str(x))
    return int(m.group(1)) if m else None

def money_fmt(x):
    try:
        return f"${x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def pct_fmt(x):
    if x is None or pd.isna(x):
        return "—"
    return f"{x*100:.1f}%"

def estado_por_umbral(cumpl, umbral_amar, umbral_verde):
    if cumpl is None or pd.isna(cumpl):
        return "—"
    if cumpl >= umbral_verde:
        return "Verde"
    if cumpl >= umbral_amar:
        return "Amarillo"
    return "Rojo"

def chip_estado(estado):
    if estado == "Verde":
        return "🟩 Verde"
    if estado == "Amarillo":
        return "🟨 Amarillo"
    if estado == "Rojo":
        return "🟥 Rojo"
    return "—"

# =========================================================
# CARGA DEL EXCEL DESDE GOOGLE DRIVE
# =========================================================
@st.cache_data(show_spinner=True, ttl=300)
def load_from_drive():
    url = f"https://drive.google.com/uc?id={DRIVE_FILE_ID}"
    gdown.download(url, EXCEL_LOCAL, quiet=True, fuzzy=True)

    df = pd.read_excel(EXCEL_LOCAL, sheet_name="BASE_INPUT")
    dim_kpi = pd.read_excel(EXCEL_LOCAL, sheet_name="DIM_KPI")

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Semana_Num"] = df["Semana"].apply(parse_semana_num)

    if "Umbral_Amarillo" not in dim_kpi.columns:
        dim_kpi["Umbral_Amarillo"] = 0.90
    if "Umbral_Verde" not in dim_kpi.columns:
        dim_kpi["Umbral_Verde"] = 1.00

    return df, dim_kpi

# =========================================================
# TRANSFORMACIONES
# =========================================================
def build_kpi_week(df):
    df = df.copy()

    df["Real_val"] = df.apply(lambda r: r["Real_$"] if r["Tipo_KPI"] == "$" else r["Real_Q"], axis=1)
    df["Obj_val"]  = df.apply(lambda r: r["Objetivo_$"] if r["Tipo_KPI"] == "$" else r["Objetivo_Q"], axis=1)

    agg = df.groupby(["Semana_Num", "Sucursal", "KPI", "Tipo_KPI"], as_index=False).agg(
        Real_Sem=("Real_val", "sum"),
        Obj_Sem=("Obj_val", "max"),
        Costo_Sem=("Costo_$", "sum"),
        Margen_Sem=("Margen_$", "sum")
    )

    agg["Cumpl_Sem"] = agg["Real_Sem"] / agg["Obj_Sem"]
    return agg

def add_acumulado(df):
    df = df.sort_values(["Sucursal", "KPI", "Semana_Num"]).copy()

    df["Real_Acum"] = df.groupby(["Sucursal", "KPI"])["Real_Sem"].cumsum()
    df["Obj_Acum"]  = df.groupby(["Sucursal", "KPI"])["Obj_Sem"].cumsum()
    df["Cumpl_Acum"] = df["Real_Acum"] / df["Obj_Acum"]

    df["Margen_Acum"] = df.groupby(["Sucursal", "KPI"])["Margen_Sem"].cumsum()
    df["MargenPct_Acum"] = (df["Margen_Acum"] / df["Real_Acum"]).where(df["Real_Acum"] != 0)

    return df

def aplicar_reglas(df, dim_kpi):
    df = df.merge(dim_kpi[["KPI", "Umbral_Amarillo", "Umbral_Verde"]], on="KPI", how="left")
    df["Estado_Acum"] = df.apply(
        lambda r: estado_por_umbral(r["Cumpl_Acum"], r["Umbral_Amarillo"], r["Umbral_Verde"]),
        axis=1
    )
    return df

# =========================================================
# APP
# =========================================================
df_base, dim_kpi = load_from_drive()
df_week = add_acumulado(build_kpi_week(df_base))

# =========================================================
# FILTROS
# =========================================================
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df_week["Semana_Num"].dropna().unique())
sucursales = sorted(df_week["Sucursal"].dropna().unique())

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=len(semanas)-1)
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

df_cut = df_week[df_week["Semana_Num"] <= semana_corte]
df_last = df_cut[df_cut["Semana_Num"] == semana_corte]

if sucursal != "TODAS (Consolidado)":
    df_last = df_last[df_last["Sucursal"] == sucursal]
else:
    df_last = df_last.groupby(["KPI", "Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum", "sum"),
        Obj_Acum=("Obj_Acum", "sum"),
        Margen_Acum=("Margen_Acum", "sum")
    )
    df_last["Cumpl_Acum"] = df_last["Real_Acum"] / df_last["Obj_Acum"]
    df_last["MargenPct_Acum"] = df_last["Margen_Acum"] / df_last["Real_Acum"]

df_last = aplicar_reglas(df_last, dim_kpi)

# =========================================================
# VISUAL
# =========================================================
st.title("Tablero Posventa — Semanal + Acumulado")
