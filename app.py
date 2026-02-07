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

def estado_global(serie_estados):
    s = set([x for x in serie_estados.dropna().tolist()])
    if "Rojo" in s: return "Rojo"
    if "Amarillo" in s: return "Amarillo"
    if "Verde" in s: return "Verde"
    return "—"

# ==========================
# Carga desde Google Sheets (export a XLSX)
# ==========================
@st.cache_data(show_spinner=True, ttl=300)
def load_from_drive():
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

    # Semanal
    agg["Cumpl_Sem"] = agg["Real_Sem"] / agg["Obj_Sem"]

    # Acumulado por sucursal + KPI
    agg = agg.sort_values(["Sucursal", "KPI", "Semana_Num"]).copy()
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
    out["Estado_Acum"] = out.apply(
        lambda r: estado_por_umbral(r["Cumpl_Acum"], r["Umbral_Amarillo"], r["Umbral_Verde"]),
        axis=1
    )
    return out

def consolidar_todas(df_last_suc):
    cons = df_last_suc.groupby(["KPI","Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
        Real_Sem=("Real_Sem","sum"),
        Obj_Sem=("Obj_Sem","sum"),
        Margen_Sem=("Margen_Sem","sum"),
    )
    cons["Cumpl_Acum"] = cons["Real_Acum"] / cons["Obj_Acum"]
    cons["Cumpl_Sem"]  = cons["Real_Sem"] / cons["Obj_Sem"]
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

# Snapshots semana corte
df_last_suc = df_week[df_week["Semana_Num"] == semana_corte].copy()
df_last_suc = aplicar_reglas(df_last_suc, dim_kpi)

# Vista principal
if sucursal != "TODAS (Consolidado)":
    df_last = df_last_suc[df_last_suc["Sucursal"] == sucursal].copy()
else:
    df_last = consolidar_todas(df_last_suc)
    df_last = aplicar_reglas(df_last, dim_kpi)

# ==========================
# VISUAL
# ==========================
st.title("Tablero Posventa — Semanal + Acumulado")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🏠 Resumen Ejecutivo", "📈 Seguimiento", "🧩 Gestión"])

# ==========================
# TAB 1: Resumen Ejecutivo
# ==========================
with tab1:
    econ = df_last[df_last["Tipo_KPI"] == "$"].copy()
    oper = df_last[df_last["Tipo_KPI"] == "Q"].copy()

    # (1) Consolidados correctos SUM/SUM
    econ_real = float(econ["Real_Acum"].sum()) if len(econ) else 0.0
    econ_obj  = float(econ["Obj_Acum"].sum())  if len(econ) else 0.0
    econ_cump = (econ_real / econ_obj) if econ_obj else None

    econ_margen = float(econ["Margen_Acum"].sum()) if ("Margen_Acum" in econ.columns and len(econ)) else 0.0
    econ_margen_pct = (econ_margen / econ_real) if econ_real else None

    oper_real = float(oper["Real_Acum"].sum()) if len(oper) else 0.0
    oper_obj  = float(oper["Obj_Acum"].sum())  if len(oper) else 0.0
    oper_cump = (oper_real / oper_obj) if oper_obj else None

    # (2) Semáforo global + conteos
    econ_rojo = int((econ["Estado_Acum"] == "Rojo").sum()) if len(econ) else 0
    econ_amar = int((econ["Estado_Acum"] == "Amarillo").sum()) if len(econ) else 0
    econ_ver  = int((econ["Estado_Acum"] == "Verde").sum()) if len(econ) else 0

    oper_rojo = int((oper["Estado_Acum"] == "Rojo").sum()) if len(oper) else 0
    oper_amar = int((oper["Estado_Acum"] == "Amarillo").sum()) if len(oper) else 0
    oper_ver  = int((oper["Estado_Acum"] == "Verde").sum()) if len(oper) else 0

    glob_econ = estado_global(econ["Estado_Acum"]) if len(econ) else "—"
    glob_oper = estado_global(oper["Estado_Acum"]) if len(oper) else "—"

    a,b,c,d = st.columns([1,1,1,1])
    a.metric("Estado Global ($)", chip_estado(glob_econ))
    b.metric("KPIs ($) por color", f"🟥{econ_rojo}  🟨{econ_amar}  🟩{econ_ver}")
    c.metric("Estado Global (Q)", chip_estado(glob_oper))
    d.metric("KPIs (Q) por color", f"🟥{oper_rojo}  🟨{oper_amar}  🟩{oper_ver}")
    st.divider()

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.subheader("🔵 Económico ($)")
        st.metric("Cumplimiento Acumulado ($)", pct_fmt(econ_cump))
        st.metric("Facturación Acumulada", money_fmt(econ_real))
        st.metric("Margen % Acumulado", pct_fmt(econ_margen_pct))

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
        st.metric("Cumplimiento Acumulado (Q)", pct_fmt(oper_cump))
        st.metric("Real Acumulado (Q)", f"{oper_real:,.0f}".replace(",", "."))
        st.metric("Objetivo Acumulado (Q)", f"{oper_obj:,.0f}".replace(",", "."))

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

    # (3) Heatmap KPI × Sucursal (solo TODAS)
    if sucursal == "TODAS (Consolidado)":
        st.divider()
        st.subheader("🔥 Heatmap KPI × Sucursal (Cumplimiento Acumulado)")

        heat = df_last_suc.pivot_table(
            index="KPI",
            columns="Sucursal",
            values="Cumpl_Acum",
            aggfunc="mean"
        ).reset_index().melt(id_vars="KPI", var_name="Sucursal", value_name="Cumpl_Acum")

        fig_h = px.density_heatmap(
            heat,
            x="Sucursal",
            y="KPI",
            z="Cumpl_Acum",
            histfunc="avg",
            title="Cumplimiento Acumulado por KPI y Sucursal",
            labels={"Cumpl_Acum": "Cumpl. Acum"}
        )
        fig_h.update_layout(coloraxis_colorbar=dict(tickformat=".0%"))
        st.plotly_chart(fig_h, use_container_width=True)

# ==========================
# TAB 2: Seguimiento
# ==========================
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

# ==========================
# TAB 3: Gestión
# ==========================
with tab3:
    st.subheader("Gestión por desvíos (acumulado)")

    g = df_last.copy()
    g["Gap"] = g["Obj_Acum"] - g["Real_Acum"]

    # Orden: rojos primero, luego amarillos, luego verdes, y por gap descendente
    order_map = {"Rojo": 0, "Amarillo": 1, "Verde": 2, "—": 9}
    g["OrdenEstado"] = g["Estado_Acum"].map(order_map).fillna(9)
    g = g.sort_values(["OrdenEstado", "Gap"], ascending=[True, False])

    g["Estado"] = g["Estado_Acum"].apply(chip_estado)
    g["Cumpl_Acum"] = g["Cumpl_Acum"].apply(pct_fmt)

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
