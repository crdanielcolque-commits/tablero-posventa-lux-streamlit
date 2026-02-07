import re
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

st.set_page_config(page_title="Tablero Posventa", layout="wide")

# ==========================
# CONFIG DRIVE
# ==========================
DRIVE_FILE_ID = "12J0gKlKfRvztWnInHg9XvT8vRq5oLlfQ"
EXCEL_LOCAL = "base_posventa.xlsx"

# ==========================
# Helpers
# ==========================
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
    if estado == "Verde": return "🟩 Verde"
    if estado == "Amarillo": return "🟨 Amarillo"
    if estado == "Rojo": return "🟥 Rojo"
    return "—"

# ==========================
# Carga desde Google Sheets (export a XLSX)
# ==========================
@st.cache_data(show_spinner=True, ttl=300)
def load_from_drive():
    # Fuerza descarga como Excel desde Google Sheets
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
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

# ==========================
# Transformaciones
# ==========================
def build_kpi_week(df):
    df = df.copy()

    df["Real_val"] = df.apply(lambda r: r["Real_$"] if r["Tipo_KPI"] == "$" else r["Real_Q"], axis=1)
    df["Obj_val"]  = df.apply(lambda r: r["Objetivo_$"] if r["Tipo_KPI"] == "$" else r["Objetivo_Q"], axis=1)

    agg = df.groupby(["Semana_Num", "Sucursal", "KPI", "Tipo_KPI"], as_index=False).agg(
        Real_Sem=("Real_val", "sum"),
        Obj_Sem=("Obj_val", "max"),       # evita duplicar objetivo
        Costo_Sem=("Costo_$", "sum"),
        Margen_Sem=("Margen_$", "sum"),
    )

    agg["Cumpl_Sem"] = agg["Real_Sem"] / agg["Obj_Sem"]
    agg["Real_Acum"] = agg.groupby(["Sucursal","KPI"])["Real_Sem"].cumsum()
    agg["Obj_Acum"]  = agg.groupby(["Sucursal","KPI"])["Obj_Sem"].cumsum()
    agg["Cumpl_Acum"] = agg["Real_Acum"] / agg["Obj_Acum"]

    agg["Margen_Acum"] = agg.groupby(["Sucursal","KPI"])["Margen_Sem"].cumsum()
    agg["MargenPct_Acum"] = (agg["Margen_Acum"] / agg["Real_Acum"]).where(agg["Real_Acum"] != 0)

    return agg

def aplicar_reglas(df_last, dim_kpi):
    out = df_last.merge(dim_kpi[["KPI","Umbral_Amarillo","Umbral_Verde"]], on="KPI", how="left")
    out["Umbral_Amarillo"] = out["Umbral_Amarillo"].fillna(0.90)
    out["Umbral_Verde"] = out["Umbral_Verde"].fillna(1.00)
    out["Estado_Acum"] = out.apply(lambda r: estado_por_umbral(r["Cumpl_Acum"], r["Umbral_Amarillo"], r["Umbral_Verde"]), axis=1)
    return out

def consolidar_todas(df_last_suc):
    cons = df_last_suc.groupby(["KPI","Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
    )
    cons["Cumpl_Acum"] = cons["Real_Acum"] / cons["Obj_Acum"]
    cons["MargenPct_Acum"] = (cons["Margen_Acum"] / cons["Real_Acum"]).where(cons["Real_Acum"] != 0)
    return cons

# ==========================
# App
# ==========================
df_base, dim_kpi = load_from_drive()
df_week = build_kpi_week(df_base)

# Sidebar filtros
st.sidebar.title("Filtros obligatorios")
semanas = sorted(df_week["Semana_Num"].dropna().unique())
sucursales = sorted(df_week["Sucursal"].dropna().unique())

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=len(semanas)-1)
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

df_last = df_week[df_week["Semana_Num"] == semana_corte].copy()

if sucursal != "TODAS (Consolidado)":
    df_last = df_last[df_last["Sucursal"] == sucursal].copy()
else:
    df_last = consolidar_todas(df_last)

df_last = aplicar_reglas(df_last, dim_kpi)

# ==========================
# VISUAL
# ==========================
st.title("Tablero Posventa — Semanal + Acumulado")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🏠 Resumen Ejecutivo", "📈 Seguimiento", "🧩 Gestión"])

with tab1:
    econ = df_last[df_last["Tipo_KPI"] == "$"].copy()
    oper = df_last[df_last["Tipo_KPI"] == "Q"].copy()

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.subheader("🔵 Económico ($)")
        st.metric("Cumplimiento Acumulado", pct_fmt(econ["Cumpl_Acum"].sum() / len(econ) if len(econ) else None))
        st.metric("Facturación Acum", money_fmt(econ["Real_Acum"].sum() if len(econ) else 0))
        st.metric("Margen % Acum", pct_fmt(econ["MargenPct_Acum"].mean() if "MargenPct_Acum" in econ.columns and len(econ) else None))

        if len(econ):
            fig = px.bar(
                econ.sort_values("Cumpl_Acum"),
                x="Cumpl_Acum", y="KPI", orientation="h",
                text=econ["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%")
            )
            fig.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay KPIs económicos ($) en este corte.")

    with col2:
        st.subheader("🟢 Operativo (Q)")
        st.metric("Cumplimiento Acumulado", pct_fmt(oper["Cumpl_Acum"].mean() if len(oper) else None))

        if len(oper):
            fig2 = px.bar(
                oper.sort_values("Cumpl_Acum"),
                x="Cumpl_Acum", y="KPI", orientation="h",
                text=oper["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%")
            )
            fig2.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No hay KPIs operativos (Q) en este corte.")

with tab2:
    st.subheader("Seguimiento por KPI (semanal vs acumulado)")

    if sucursal == "TODAS (Consolidado)":
        st.info("Para seguimiento semanal por KPI, elegí una sucursal (en consolidado mezclaría bases).")

    else:
        kpis = sorted(df_week[df_week["Sucursal"] == sucursal]["KPI"].dropna().unique())
        kpi_sel = st.selectbox("KPI", kpis)

        serie = df_week[(df_week["Sucursal"] == sucursal) & (df_week["KPI"] == kpi_sel)].sort_values("Semana_Num")

        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.line(serie, x="Semana_Num", y="Cumpl_Sem", markers=True, title="Cumplimiento semanal")
            fig1.update_layout(yaxis_tickformat=".0%")
            st.plotly_chart(fig1, use_container_width=True)

        with c2:
            fig2 = px.line(serie, x="Semana_Num", y="Cumpl_Acum", markers=True, title="Cumplimiento acumulado")
            fig2.update_layout(yaxis_tickformat=".0%")
            st.plotly_chart(fig2, use_container_width=True)

with tab3:
    st.subheader("Gestión por desvíos (acumulado)")

    g = df_last.copy()
    g["Gap"] = g["Obj_Acum"] - g["Real_Acum"]
    g = g.sort_values(["Estado_Acum","Gap"], ascending=[True, False])

    g["Estado"] = g["Estado_Acum"].apply(chip_estado)
    g["Cumpl_Acum"] = g["Cumpl_Acum"].apply(pct_fmt)

    # Formatos
    def fmt_val(tipo, v):
        if pd.isna(v): return "—"
        return money_fmt(v) if tipo == "$" else f"{v:,.0f}".replace(",", ".")

    g["Real_Acum_fmt"] = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Real_Acum"]), axis=1)
    g["Obj_Acum_fmt"]  = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Obj_Acum"]), axis=1)
    g["Gap_fmt"]       = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Gap"]), axis=1)

    st.dataframe(
        g[["KPI","Tipo_KPI","Estado","Cumpl_Acum","Real_Acum_fmt","Obj_Acum_fmt","Gap_fmt"]],
        use_container_width=True,
        hide_index=True
    )
