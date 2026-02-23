# ============================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado) v2.3
# + Modo Presentación (sin sidebar)
# + Resumen ejecutivo 1 línea
# + Estados con íconos premium
# + Cap visual de % (ranking/charts) sin tocar cálculos
# + Tabla auditoría (expander)
# + Export Excel (detalle + acumulados)
#
# v2.3.2 (MOD) — KPI Cards Semanal + Acumulado + Proyección fin de mes
#               usando hoja "Dias habiles" (Mes, Semana, Dias habiles)
# + Mini evolución semanal (Acum vs Proy)
# ============================================================

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown
from io import BytesIO

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

    s = (
        s.replace("$", "")
         .replace("AR$", "")
         .replace(" ", "")
         .replace("\u00A0", "")
    )

    # Si hay coma => coma decimal y punto miles
    if "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
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

def estado_icon(est_txt: str) -> str:
    return {"Verde": "✅", "Amarillo": "⚠️", "Rojo": "🔴", "—": "—"}.get(est_txt, "—")

def badge_estado_html(est):
    icon = estado_icon(est)
    color = {"Verde":"#198754", "Amarillo":"#d39e00", "Rojo":"#dc3545", "—":"#6c757d"}.get(est, "#6c757d")
    bg    = {"Verde":"#d1e7dd", "Amarillo":"#fff3cd", "Rojo":"#f8d7da", "—":"#e9ecef"}.get(est, "#e9ecef")
    return f"""
    <span style="display:inline-block;padding:4px 10px;border-radius:999px;
                 background:{bg};color:{color};font-weight:800;font-size:12px;border:1px solid {color}33;">
        {icon} {est.upper()}
    </span>
    """

def card_html(title, value, sub, estado_txt=None):
    estado_block = ""
    if estado_txt is not None:
        estado_block = f"<div style='margin-top:8px;'>{badge_estado_html(estado_txt)}</div>"
    return f"""
    <div style="border:1px solid #eee;border-radius:14px;padding:16px;background:#fff;
                box-shadow:0 2px 10px rgba(0,0,0,0.04);">
        <div style="font-size:12px;color:#6c757d;font-weight:800;letter-spacing:0.2px;">{title}</div>
        <div style="font-size:28px;font-weight:900;margin-top:6px;">{value}</div>
        <div style="font-size:12px;color:#6c757d;margin-top:6px;">{sub}</div>
        {estado_block}
    </div>
    """

def chips_css_soft_green():
    # Chips (multiselect) en verde suave, no “alerta”
    return """
    <style>
    /* Streamlit multiselect selected chips */
    div[data-baseweb="tag"]{
        background-color: #d1e7dd !important;
        border: 1px solid rgba(25,135,84,0.25) !important;
    }
    div[data-baseweb="tag"] span{
        color: #0f5132 !important;
        font-weight: 700 !important;
    }
    </style>
    """

def hide_sidebar_css():
    return """
    <style>
      section[data-testid="stSidebar"] {display: none !important;}
      div[data-testid="stSidebarNav"] {display: none !important;}
      /* Ajuste padding main */
      .block-container {padding-left: 2.2rem; padding-right: 2.2rem;}
    </style>
    """

def month_name_es(month_num: int) -> str:
    m = {
        1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
        7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
    }
    return m.get(int(month_num), "")

def norm_text(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower()
    # normalización simple de tildes comunes
    s = (s.replace("á","a").replace("é","e").replace("í","i")
           .replace("ó","o").replace("ú","u").replace("ü","u")
           .replace("ñ","n"))
    return s

# ---------------------------
# LOAD (hoja principal + hoja "Dias habiles")
# ---------------------------
@st.cache_data(ttl=300)
def load_from_drive():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)

    xls = pd.ExcelFile(EXCEL_LOCAL)
    df0 = pd.read_excel(xls, sheet_name=0)
    df0 = df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")]

    # Hoja Dias habiles (si no existe, devolvemos vacía)
    try:
        df_dias = pd.read_excel(xls, sheet_name="Dias habiles")
    except Exception:
        df_dias = pd.DataFrame(columns=["Mes","Semana","Dias habiles"])

    return df0, df_dias

df, df_dias_habiles = load_from_drive()

# ---------------------------
# VALIDACIÓN base
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
# NORMALIZACIÓN base
# ---------------------------
df["Semana_Num"] = parse_semana_num(df["Semana"])
df = df[~df["Semana_Num"].isna()].copy()

# Fecha a datetime
df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
df = df[~df["Fecha"].isna()].copy()

# Mes técnico (YYYY-MM) + Mes nombre (Enero/Febrero/...)
df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)
df["Mes_Nombre"] = df["Fecha"].dt.month.apply(month_name_es)
df["Mes_Nombre_norm"] = df["Mes_Nombre"].apply(norm_text)

# Parse AR numérico
for c in ["Real_$","Costo_$","Margen_$","Margen_%","Real_Q","Objetivo_$","Objetivo_Q","Cumplimiento_%"]:
    if c in df.columns:
        df[c] = df[c].apply(to_num_ar)

# Limpieza strings
df["KPI"] = df["KPI"].astype(str).str.strip()
df["Categoria_KPI"] = df["Categoria_KPI"].astype(str).str.strip()
df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()
df["Sucursal"] = df["Sucursal"].astype(str).str.strip()

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
# NORMALIZACIÓN hoja Dias habiles
# ---------------------------
if not df_dias_habiles.empty:
    # Forzar columnas esperadas
    for col in ["Mes","Semana","Dias habiles"]:
        if col not in df_dias_habiles.columns:
            df_dias_habiles[col] = np.nan

    df_dias_habiles["Mes_norm"] = df_dias_habiles["Mes"].apply(norm_text)
    df_dias_habiles["Semana"] = pd.to_numeric(df_dias_habiles["Semana"], errors="coerce").fillna(0).astype(int)
    df_dias_habiles["Dias habiles"] = pd.to_numeric(df_dias_habiles["Dias habiles"], errors="coerce").fillna(0).astype(float)

# CSS chips soft green (siempre)
st.markdown(chips_css_soft_green(), unsafe_allow_html=True)

# ---------------------------
# CONTROLES (estado)
# ---------------------------
if "modo_presentacion" not in st.session_state:
    st.session_state["modo_presentacion"] = False

if "cap_visual" not in st.session_state:
    st.session_state["cap_visual"] = True

if "cap_val" not in st.session_state:
    st.session_state["cap_val"] = 2.0  # 200%

# ---------------------------
# UTIL: filtro Obj=0 y cap visual
# ---------------------------
def apply_obj0_filter(d, show_obj0: bool):
    if show_obj0:
        return d.copy()
    return d[d["Obj_val"] > 0].copy()

def apply_cap_visual(d, cap_on: bool, cap_value: float):
    # no toca Cumpl real, crea Cumpl_plot
    out = d.copy()
    if "Cumpl" not in out.columns:
        return out
    if cap_on:
        out["Cumpl_plot"] = out["Cumpl"].clip(upper=cap_value)
    else:
        out["Cumpl_plot"] = out["Cumpl"]
    return out

# ---------------------------
# FILTROS DISPONIBLES
# ---------------------------
semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
if not semanas:
    st.error("No se encontraron semanas válidas en la columna 'Semana'.")
    st.stop()

default_sem = 1 if 1 in semanas else min(semanas)
sucursales = sorted(df["Sucursal"].dropna().unique().tolist())

# ---------------------------
# TOP BAR: Modo Presentación + Cap
# ---------------------------
topc1, topc2, topc3, topc4 = st.columns([1.1, 1.2, 1.2, 1.5])
with topc1:
    st.session_state["modo_presentacion"] = st.toggle("Modo presentación", value=st.session_state["modo_presentacion"])
with topc2:
    st.session_state["cap_visual"] = st.toggle("Cap visual %", value=st.session_state["cap_visual"])
with topc3:
    cap_options = {"150%": 1.5, "200%": 2.0, "300%": 3.0, "Sin cap": 999.0}
    cap_label = st.selectbox("Máx visual", list(cap_options.keys()), index=1)
    st.session_state["cap_val"] = cap_options[cap_label]
with topc4:
    st.caption("Tip: el cap es **solo visual** (ranking/gráficos). No altera el cálculo base.")

# Si presentación: ocultar sidebar
if st.session_state["modo_presentacion"]:
    st.markdown(hide_sidebar_css(), unsafe_allow_html=True)

# ---------------------------
# INPUTS: sidebar normal o barra superior (presentación)
# ---------------------------
def render_filters(area="sidebar"):
    # defaults robustos
    if "semana_corte" not in st.session_state:
        st.session_state["semana_corte"] = default_sem
    if "sucursal" not in st.session_state:
        st.session_state["sucursal"] = "TODAS (Consolidado)"
    if "show_obj0" not in st.session_state:
        st.session_state["show_obj0"] = False

    container = st.sidebar if area == "sidebar" else st.container()

    with container:
        if area != "sidebar":
            st.markdown("### Filtros")
        else:
            st.sidebar.markdown("## Filtros obligatorios")

        st.session_state["semana_corte"] = st.selectbox(
            "Semana corte", semanas,
            index=semanas.index(st.session_state["semana_corte"]) if st.session_state["semana_corte"] in semanas else 0,
            key=f"semana_{area}"
        )

        st.session_state["sucursal"] = st.selectbox(
            "Sucursal", ["TODAS (Consolidado)"] + sucursales,
            index=(["TODAS (Consolidado)"] + sucursales).index(st.session_state["sucursal"])
            if st.session_state["sucursal"] in (["TODAS (Consolidado)"] + sucursales) else 0,
            key=f"sucursal_{area}"
        )

        st.markdown("---")
        st.markdown("### Cálculo")
        st.session_state["show_obj0"] = st.checkbox(
            "Incluir filas con Obj=0 (puede distorsionar %)",
            value=st.session_state["show_obj0"],
            key=f"obj0_{area}"
        )

# Render filtros
if st.session_state["modo_presentacion"]:
    with st.expander("Abrir filtros (presentación)", expanded=False):
        render_filters(area="top")
else:
    render_filters(area="sidebar")

semana_corte = int(st.session_state["semana_corte"])
sucursal = st.session_state["sucursal"]
show_obj0 = bool(st.session_state["show_obj0"])
cap_on = bool(st.session_state["cap_visual"])
cap_val = float(st.session_state["cap_val"])

# ---------------------------
# CORTE
# ---------------------------
df_cut = df[df["Semana_Num"] <= semana_corte].copy()
if sucursal != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal].copy()

# ---------------------------
# Mes de referencia para proyección
# Tomamos el mes del último dato disponible al corte.
# ---------------------------
if not df_cut.empty:
    mes_ref = df_cut["Mes"].max()
    mes_nombre_ref_norm = df_cut.loc[df_cut["Mes"] == mes_ref, "Mes_Nombre_norm"].iloc[0]
else:
    mes_ref = df["Mes"].max()
    mes_nombre_ref_norm = df.loc[df["Mes"] == mes_ref, "Mes_Nombre_norm"].iloc[0]

df_month = df[df["Mes"] == mes_ref].copy()
if sucursal != "TODAS (Consolidado)":
    df_month = df_month[df_month["Sucursal"] == sucursal].copy()

# ---------------------------
# Filtros P&L aperturas
# ---------------------------
def compute_openings_pl(dfc):
    rep_open = sorted(dfc[(dfc["KPI"].str.upper()=="REPUESTOS") & (dfc["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
    srv_open = sorted(dfc[(dfc["KPI"].str.upper()=="SERVICIOS") & (dfc["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
    return rep_open, srv_open

rep_open, srv_open = compute_openings_pl(df_cut)

# Estado session para selecciones
if "rep_sel" not in st.session_state:
    st.session_state["rep_sel"] = rep_open
if "srv_sel" not in st.session_state:
    st.session_state["srv_sel"] = srv_open

# Ajuste por cambios de base: si aparece algo nuevo, sumarlo por defecto
for x in rep_open:
    if x not in st.session_state["rep_sel"]:
        st.session_state["rep_sel"].append(x)
for x in srv_open:
    if x not in st.session_state["srv_sel"]:
        st.session_state["srv_sel"].append(x)

# Eliminar seleccionados que ya no existen
st.session_state["rep_sel"] = [x for x in st.session_state["rep_sel"] if x in rep_open]
st.session_state["srv_sel"] = [x for x in st.session_state["srv_sel"] if x in srv_open]

# Sidebar chips (si presentación, van dentro del expander de filtros)
def render_pl_multiselect(area="sidebar"):
    container = st.sidebar if area == "sidebar" else st.container()
    with container:
        st.markdown("## Incluir variables (P&L)")
        st.session_state["rep_sel"] = st.multiselect(
            "Repuestos: aperturas incluidas",
            rep_open, default=st.session_state["rep_sel"],
            key=f"rep_sel_{area}"
        )
        st.session_state["srv_sel"] = st.multiselect(
            "Servicios: aperturas incluidas",
            srv_open, default=st.session_state["srv_sel"],
            key=f"srv_sel_{area}"
        )

if st.session_state["modo_presentacion"]:
    with st.expander("Abrir variables P&L (presentación)", expanded=False):
        render_pl_multiselect(area="top")
else:
    st.sidebar.markdown("---")
    render_pl_multiselect(area="sidebar")

rep_sel = st.session_state["rep_sel"]
srv_sel = st.session_state["srv_sel"]

# ---------------------------
# NUEVO: Serie semanal + proyección usando hoja Dias habiles
# ---------------------------
def kpi_weekly_series_for_scope(
    df_scope_month: pd.DataFrame,
    *,
    cap_on: bool,
    cap_val: float
) -> pd.DataFrame:
    """
    Serie por Semana_Num:
      - Real_sem, Obj_sem
      - Real_acum, Obj_acum
      - Cumpl_sem, Cumpl_acum
      - Obj_mes (suma Obj mes completo)
      - Proy_real_eom (run-rate por días hábiles acumulados hasta la semana)
      - Cumpl_proy_eom
    """
    if df_scope_month.empty:
        return pd.DataFrame()

    g = df_scope_month.groupby("Semana_Num", as_index=False).agg(
        Real_sem=("Real_val", "sum"),
        Obj_sem=("Obj_val", "sum")
    ).sort_values("Semana_Num")

    g["Real_acum"] = g["Real_sem"].cumsum()
    g["Obj_acum"]  = g["Obj_sem"].cumsum()

    g["Cumpl_sem"]  = g.apply(lambda r: safe_ratio(r["Real_sem"], r["Obj_sem"]), axis=1)
    g["Cumpl_acum"] = g.apply(lambda r: safe_ratio(r["Real_acum"], r["Obj_acum"]), axis=1)

    obj_mes = float(df_scope_month["Obj_val"].sum())
    g["Obj_mes"] = obj_mes

    # Buscar días hábiles del mes (por nombre: Febrero, etc.)
    mes_norm = df_scope_month["Mes_Nombre_norm"].iloc[0]
    dias_df = df_dias_habiles[df_dias_habiles.get("Mes_norm", "") == mes_norm].copy() if not df_dias_habiles.empty else pd.DataFrame()

    if dias_df.empty:
        # Sin hoja o sin mes coincidente → proyección nula
        g["Dias_sem"] = np.nan
        g["Dias_acum"] = np.nan
        g["Proy_real_eom"] = np.nan
        g["Cumpl_proy_eom"] = np.nan
    else:
        total_dias_mes = float(dias_df["Dias habiles"].sum())

        # Merge días por semana
        g = g.merge(
            dias_df[["Semana", "Dias habiles"]],
            left_on="Semana_Num",
            right_on="Semana",
            how="left"
        )
        g["Dias_sem"] = pd.to_numeric(g["Dias habiles"], errors="coerce").fillna(0.0)
        g["Dias_acum"] = g["Dias_sem"].cumsum()

        # Run rate y proyección
        g["Run_rate"] = g["Real_acum"] / g["Dias_acum"].replace(0, np.nan)
        g["Proy_real_eom"] = g["Run_rate"] * total_dias_mes
        g["Cumpl_proy_eom"] = g.apply(lambda r: safe_ratio(r["Proy_real_eom"], obj_mes), axis=1)

    # Cap visual SOLO para el gráfico
    def cap_series(s):
        if cap_on:
            return s.clip(upper=cap_val)
        return s

    g["Cumpl_acum_plot"] = cap_series(g["Cumpl_acum"])
    g["Cumpl_proy_plot"] = cap_series(g["Cumpl_proy_eom"])

    return g

def kpi_card_weekly_projection(
    title: str,
    *,
    df_scope_month: pd.DataFrame,   # mes completo (para proyección + series)
    df_scope_cut: pd.DataFrame,     # corte (para valor grande 100% consistente con v2.3)
    semana_corte: int,
    tipo: str,                      # "$" o "Q" (solo para formateo)
    cap_on: bool,
    cap_val: float
):
    series = kpi_weekly_series_for_scope(df_scope_month, cap_on=cap_on, cap_val=cap_val)
    if series.empty:
        st.markdown(card_html(title, "—", "Sin datos", "—"), unsafe_allow_html=True)
        return

    s_cut = series[series["Semana_Num"] <= semana_corte].copy()
    if s_cut.empty:
        st.markdown(card_html(title, "—", f"Sin datos hasta Semana {semana_corte}", "—"), unsafe_allow_html=True)
        return

    last = s_cut.iloc[-1]

    # Valor grande (corte real)
    df_cut_local = df_scope_cut.copy()
    df_cut_local = df_cut_local[df_cut_local["Semana_Num"] <= semana_corte].copy()

    real_acum = float(df_cut_local["Real_val"].sum()) if not df_cut_local.empty else 0.0
    acum_obj  = float(df_cut_local["Obj_val"].sum())  if not df_cut_local.empty else 0.0
    acum_c    = safe_ratio(real_acum, acum_obj)

    # Semana puntual (desde corte)
    df_sem = df_cut_local[df_cut_local["Semana_Num"] == semana_corte].copy()
    if not df_sem.empty:
        sem_real = float(df_sem["Real_val"].sum())
        sem_obj  = float(df_sem["Obj_val"].sum())
        sem_c    = safe_ratio(sem_real, sem_obj)
    else:
        sem_real, sem_obj, sem_c = np.nan, np.nan, np.nan

    # Proyección (desde mes completo + hoja días hábiles)
    obj_mes   = float(last["Obj_mes"]) if not pd.isna(last["Obj_mes"]) else 0.0
    proy_real = float(last["Proy_real_eom"]) if not pd.isna(last["Proy_real_eom"]) else np.nan
    proy_c    = float(last["Cumpl_proy_eom"]) if not pd.isna(last["Cumpl_proy_eom"]) else np.nan

    est = estado(proy_c if not pd.isna(proy_c) else acum_c)

    if tipo == "$":
        main_value = money(real_acum)
        sub = (
            f"Semana {semana_corte}: {money(sem_real)} / {money(sem_obj)} ({pct(sem_c)}) | "
            f"Acum: {money(real_acum)} / {money(acum_obj)} ({pct(acum_c)}) | "
            f"Proy EOM: {money(proy_real)} / {money(obj_mes)} ({pct(proy_c)})"
        )
    else:
        main_value = qty(real_acum)
        sub = (
            f"Semana {semana_corte}: {qty(sem_real)} / {qty(sem_obj)} ({pct(sem_c)}) | "
            f"Acum: {qty(real_acum)} / {qty(acum_obj)} ({pct(acum_c)}) | "
            f"Proy EOM: {qty(proy_real)} / {qty(obj_mes)} ({pct(proy_c)})"
        )

    st.markdown(card_html(title, main_value, sub, est), unsafe_allow_html=True)

    # Sparkline (mini evolución semanal)
    plot_df = s_cut[["Semana_Num", "Cumpl_acum_plot", "Cumpl_proy_plot"]].copy()
    plot_df = plot_df.rename(columns={"Semana_Num": "Semana"})

    m1 = plot_df[["Semana", "Cumpl_acum_plot"]].rename(columns={"Cumpl_acum_plot": "Cumpl"})
    m1["Serie"] = "Acumulado"
    m2 = plot_df[["Semana", "Cumpl_proy_plot"]].rename(columns={"Cumpl_proy_plot": "Cumpl"})
    m2["Serie"] = "Proyectado"
    spark = pd.concat([m1, m2], ignore_index=True)

    fig = px.line(spark, x="Semana", y="Cumpl", color="Serie", markers=False)
    fig.update_layout(
        height=130,
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis_title="",
        yaxis_title="",
    )
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

# ---------------------------
# HEADER
# ---------------------------
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}** | Mes ref: **{mes_ref}** ({month_name_es(int(mes_ref.split('-')[1]))})")

tab1, tab2, tab3 = st.tabs(["🧩 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧪 Gestión (desvíos)"])

# ============================================================
# FUNCIONES DE RESUMEN / MICRO
# ============================================================
def summarize_segment(dseg: pd.DataFrame, tipo: str):
    dseg = apply_obj0_filter(dseg, show_obj0)
    real = dseg["Real_val"].sum()
    obj = dseg["Obj_val"].sum()
    c = safe_ratio(real, obj)
    est = estado(c)
    sub = f"Real {(money(real) if tipo=='$' else qty(real))} | Obj {(money(obj) if tipo=='$' else qty(obj))}"
    return real, obj, c, est, sub

def micro_aperturas(d: pd.DataFrame, tipo: str):
    g = d.groupby("Categoria_KPI", as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy()
    g = g.sort_values("Cumpl", ascending=False)

    g = apply_cap_visual(g, cap_on, cap_val)

    if tipo == "$":
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    else:
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {qty(r['Real'])}/{qty(r['Obj'])}", axis=1)
    return g

def ranking_sucursal_apertura_micro(d: pd.DataFrame, tipo: str, top_n: int, show_zero: bool):
    # Ranking micro por combinación Sucursal-Apertura (Categoria)
    x = d.copy()
    if not show_zero:
        x = x[x["Obj_val"] > 0].copy()

    g = x.groupby(["Sucursal","Categoria_KPI"], as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy()
    g = g.sort_values("Cumpl", ascending=False).head(top_n).copy()

    g = apply_cap_visual(g, cap_on, cap_val)

    if tipo == "$":
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    else:
        g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {qty(r['Real'])}/{qty(r['Obj'])}", axis=1)

    g["key"] = g["Sucursal"].astype(str) + " — " + g["Categoria_KPI"].astype(str)
    return g

def principal_driver_gap(d_pl: pd.DataFrame):
    # Driver principal del desvío (Obj-Real), usando selección P&L actual
    x = d_pl.copy()
    x = apply_obj0_filter(x, show_obj0)
    if x.empty:
        return None

    g = x.groupby(["KPI","Categoria_KPI"], as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
    g["Gap"] = g["Obj"] - g["Real"]
    g = g.sort_values("Gap", ascending=False)
    row = g.iloc[0]
    return {
        "KPI": str(row["KPI"]),
        "Cat": str(row["Categoria_KPI"]),
        "Gap": float(row["Gap"]),
        "Real": float(row["Real"]),
        "Obj": float(row["Obj"])
    }

def export_excel_bytes(detail_df: pd.DataFrame, acum_df: pd.DataFrame, name_detail="Detalle", name_acum="Acumulado"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, index=False, sheet_name=name_detail)
        acum_df.to_excel(writer, index=False, sheet_name=name_acum)
    output.seek(0)
    return output

# ============================================================
# TAB 1 — P&L
# ============================================================
with tab1:
    st.markdown("## 🧩 P&L — Macro → Micro")
    st.markdown("---")

    # Mes completo (para proyección) + corte (para resto del tablero)
    d_pl_month = df_month[df_month["Tipo_KPI"]=="$"].copy()
    d_pl_cut   = df_cut[df_cut["Tipo_KPI"]=="$"].copy()

    d_rep_month = d_pl_month[d_pl_month["KPI"].str.upper()=="REPUESTOS"].copy()
    d_rep_month = d_rep_month[d_rep_month["Categoria_KPI"].isin(rep_sel)].copy()

    d_srv_month = d_pl_month[d_pl_month["KPI"].str.upper()=="SERVICIOS"].copy()
    d_srv_month = d_srv_month[d_srv_month["Categoria_KPI"].isin(srv_sel)].copy()

    d_rep_cut = d_pl_cut[d_pl_cut["KPI"].str.upper()=="REPUESTOS"].copy()
    d_rep_cut = d_rep_cut[d_rep_cut["Categoria_KPI"].isin(rep_sel)].copy()

    d_srv_cut = d_pl_cut[d_pl_cut["KPI"].str.upper()=="SERVICIOS"].copy()
    d_srv_cut = d_srv_cut[d_srv_cut["Categoria_KPI"].isin(srv_sel)].copy()

    # Resumen ejecutivo 1 línea (acumulado al corte)
    rep_real, rep_obj, rep_c, rep_est, _ = summarize_segment(d_rep_cut, "$")
    srv_real, srv_obj, srv_c, srv_est, _ = summarize_segment(d_srv_cut, "$")

    driver = principal_driver_gap(pd.concat([d_rep_cut, d_srv_cut], ignore_index=True))
    if driver:
        driver_txt = f"Principal desvío: **{driver['KPI']} / {driver['Cat']}** (Gap {money(driver['Gap'])})"
    else:
        driver_txt = "Principal desvío: —"

    st.info(
        f"**Resumen Ejecutivo:** "
        f"Repuestos {pct(rep_c)} {estado_icon(rep_est)} | "
        f"Servicios {pct(srv_c)} {estado_icon(srv_est)} | "
        f"{driver_txt}"
    )

    # Macro cards (Semana + Acum + Proyección + sparkline)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        kpi_card_weekly_projection(
            "Repuestos — Real (Acum.)",
            df_scope_month=apply_obj0_filter(d_rep_month, show_obj0),
            df_scope_cut=apply_obj0_filter(d_rep_cut, show_obj0),
            semana_corte=semana_corte,
            tipo="$",
            cap_on=cap_on,
            cap_val=cap_val
        )
    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        kpi_card_weekly_projection(
            "Servicios — Real (Acum.)",
            df_scope_month=apply_obj0_filter(d_srv_month, show_obj0),
            df_scope_cut=apply_obj0_filter(d_srv_cut, show_obj0),
            semana_corte=semana_corte,
            tipo="$",
            cap_on=cap_on,
            cap_val=cap_val
        )

    st.markdown("---")
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    l, r = st.columns(2)
    with l:
        st.markdown("**Repuestos — por apertura**")
        g = micro_aperturas(apply_obj0_filter(d_rep_cut, show_obj0), "$")
        if g.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with r:
        st.markdown("**Servicios — por apertura**")
        g = micro_aperturas(apply_obj0_filter(d_srv_cut, show_obj0), "$")
        if g.empty:
            st.info("Sin datos (revisar aperturas seleccionadas / Obj=0).")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("## 🎯 Micro — ranking sucursal + apertura")

    cA, cB, cC, cD = st.columns([1.1, 1.2, 1.2, 1.5])
    with cA:
        top_n = st.selectbox("Top N", [5,10,15,20,30], index=1)
    with cB:
        rep_micro_choice = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + rep_open, index=0)
    with cC:
        srv_micro_choice = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + srv_open, index=0)
    with cD:
        show_zero_rank = st.checkbox("Mostrar 0% (Obj=0 y real=0)", value=False)

    # Rep micro ranking
    rep_rank_base = d_rep_cut.copy()
    if rep_micro_choice != "Todas las aperturas":
        rep_rank_base = rep_rank_base[rep_rank_base["Categoria_KPI"] == rep_micro_choice].copy()

    srv_rank_base = d_srv_cut.copy()
    if srv_micro_choice != "Todas las aperturas":
        srv_rank_base = srv_rank_base[srv_rank_base["Categoria_KPI"] == srv_micro_choice].copy()

    rr, ss = st.columns(2)
    with rr:
        st.markdown("### Repuestos — sucursal + apertura (micro)")
        g = ranking_sucursal_apertura_micro(rep_rank_base, "$", top_n=top_n, show_zero=show_zero_rank)
        if g.empty:
            st.info("Sin ranking (Obj=0 o sin datos).")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="key", orientation="h", text="label")
            fig.update_layout(height=460, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with ss:
        st.markdown("### Servicios — sucursal + apertura (micro)")
        g = ranking_sucursal_apertura_micro(srv_rank_base, "$", top_n=top_n, show_zero=show_zero_rank)
        if g.empty:
            st.info("Sin ranking (Obj=0 o sin datos).")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="key", orientation="h", text="label")
            fig.update_layout(height=460, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    # Auditoría + Export
    st.markdown("---")
    with st.expander("🔎 Auditoría y export (P&L)", expanded=False):
        # Detalle filtrado P&L (al corte)
        detail = pd.concat([d_rep_cut, d_srv_cut], ignore_index=True).copy()
        detail = detail.sort_values(["Semana_Num","Sucursal","KPI","Categoria_KPI"], ascending=[True, True, True, True])

        # Acumulado P&L por KPI/Categoria/Sucursal (al corte)
        acum = detail.groupby(["KPI","Categoria_KPI","Sucursal"], as_index=False).agg(
            Real=("Real_val","sum"),
            Obj=("Obj_val","sum"),
        )
        acum["Cumpl"] = acum.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
        acum["Gap"] = acum["Obj"] - acum["Real"]
        acum = acum.sort_values(["KPI","Gap"], ascending=[True, False])

        st.markdown("**Vista rápida (acumulado):**")
        st.dataframe(acum, use_container_width=True, hide_index=True)

        excel_bytes = export_excel_bytes(detail, acum, name_detail="Detalle_P&L", name_acum="Acumulado_P&L")
        st.download_button(
            "⬇️ Descargar Excel (Detalle + Acumulado P&L)",
            data=excel_bytes,
            file_name=f"Tablero_P&L_Sem{semana_corte}_{sucursal.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ============================================================
# TAB 2 — KPIs resto
# ============================================================
with tab2:
    st.markdown("## 📌 KPIs (resto) — Macro → Micro")
    st.markdown("---")

    resto_month = df_month[~df_month["KPI"].str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()
    resto_cut   = df_cut[~df_cut["KPI"].str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()

    resto_month = apply_obj0_filter(resto_month, show_obj0)
    resto_cut   = apply_obj0_filter(resto_cut, show_obj0)

    # Tarjetas por KPI
    st.markdown("### 🧩 Tarjetas KPI — Semana / Acum / Proyección (fin de mes)")
    kpi_pairs = (
        resto_month.groupby(["KPI","Tipo_KPI"], as_index=False)
        .size()[["KPI","Tipo_KPI"]]
        .sort_values(["Tipo_KPI","KPI"])
    )

    if kpi_pairs.empty:
        st.info("No hay KPIs (resto) con Obj>0 en este mes.")
    else:
        cols = st.columns(3)
        for i, row in enumerate(kpi_pairs.itertuples(index=False)):
            kpi_name = row.KPI
            tipo_kpi = row.Tipo_KPI

            scope_month = resto_month[(resto_month["KPI"] == kpi_name) & (resto_month["Tipo_KPI"] == tipo_kpi)].copy()
            scope_cut   = resto_cut[(resto_cut["KPI"] == kpi_name) & (resto_cut["Tipo_KPI"] == tipo_kpi)].copy()

            with cols[i % 3]:
                kpi_card_weekly_projection(
                    f"{kpi_name} ({tipo_kpi})",
                    df_scope_month=scope_month,
                    df_scope_cut=scope_cut,
                    semana_corte=semana_corte,
                    tipo=("$" if tipo_kpi == "$" else "Q"),
                    cap_on=cap_on,
                    cap_val=cap_val
                )

    st.markdown("---")

    # Deep dive (tu lógica original)
    kpis_resto = sorted(resto_cut["KPI"].unique().tolist())
    if not kpis_resto:
        st.info("No hay KPIs (resto) con Obj>0 en este corte.")
    else:
        st.markdown("### 🔍 Detalle (ranking por sucursal) — Elegí un KPI")
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto)

        x = resto_cut[resto_cut["KPI"] == kpi_sel].copy()
        tipos = sorted(x["Tipo_KPI"].unique().tolist())

        for t in tipos:
            xt = x[x["Tipo_KPI"] == t].copy()
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

            g = xt.groupby("Sucursal", as_index=False).agg(
                Real=("Real_val","sum"),
                Obj=("Obj_val","sum")
            )
            g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
            g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
            g = apply_cap_visual(g, cap_on, cap_val)

            g["label"] = g.apply(
                lambda r: f"{pct(r['Cumpl'])} | {(money(r['Real']) if t=='$' else qty(r['Real']))}/{(money(r['Obj']) if t=='$' else qty(r['Obj']))}",
                axis=1
            )

            st.markdown("### Ranking por sucursal — este KPI")
            if g.empty:
                st.info("Sin ranking (Obj=0 o sin datos).")
            else:
                fig = px.bar(g, x="Cumpl_plot", y="Sucursal", orientation="h", text="label")
                fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
                fig.update_traces(textposition="inside")
                st.plotly_chart(fig, use_container_width=True)

        # Auditoría + Export
        st.markdown("---")
        with st.expander("🔎 Auditoría y export (KPIs resto)", expanded=False):
            detail = x.copy().sort_values(["Semana_Num","Sucursal","KPI","Categoria_KPI"])
            acum = detail.groupby(["KPI","Tipo_KPI","Sucursal"], as_index=False).agg(
                Real=("Real_val","sum"),
                Obj=("Obj_val","sum"),
            )
            acum["Cumpl"] = acum.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
            acum["Gap"] = acum["Obj"] - acum["Real"]
            acum = acum.sort_values(["Tipo_KPI","Gap"], ascending=[True, False])

            st.dataframe(acum, use_container_width=True, hide_index=True)

            excel_bytes = export_excel_bytes(detail, acum, name_detail="Detalle_KPIsResto", name_acum="Acumulado_KPIsResto")
            st.download_button(
                "⬇️ Descargar Excel (Detalle + Acumulado KPIs resto)",
                data=excel_bytes,
                file_name=f"Tablero_KPIsResto_{kpi_sel}_Sem{semana_corte}_{sucursal.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ============================================================
# TAB 3 — Gestión
# ============================================================
with tab3:
    st.markdown("## 🧪 Gestión (desvíos)")
    st.markdown("---")

    # Filtro sucursal dentro del tab (pedido)
    suc_g = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)

    d = df[df["Semana_Num"] <= semana_corte].copy()
    if suc_g != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_g].copy()

    d = apply_obj0_filter(d, show_obj0)

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

        gg_plot = gg.copy()
        gg_plot["label"] = gg_plot.apply(
            lambda r: f"Gap {(money(r['Gap']) if r['Tipo_KPI']=='$' else qty(r['Gap']))} | {pct(r['Cumpl'])}",
            axis=1
        )

        fig = px.bar(gg_plot, x="Gap", y="key", orientation="h", text="label")
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Gap (Obj - Real)")
        fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("🔎 Auditoría y export (Gestión)", expanded=False):
            st.dataframe(g, use_container_width=True, hide_index=True)
            detail = d.copy().sort_values(["Semana_Num","Sucursal","KPI","Categoria_KPI"])
            excel_bytes = export_excel_bytes(detail, g, name_detail="Detalle_Gestion", name_acum="Desvios_Gestion")
            st.download_button(
                "⬇️ Descargar Excel (Detalle + Desvíos Gestión)",
                data=excel_bytes,
                file_name=f"Tablero_Gestion_Sem{semana_corte}_{suc_g.replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
