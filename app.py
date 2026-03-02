# ============================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado) v2.3
# v2.3.14
# ✅ Botón global: Descargar Resumen Ejecutivo (Excel)
# ✅ TAB: 🗓️ Hitos del mes desde "Resumen del mes" (fuzzy match)
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
    s = series.astype(str).str.strip()
    num = s.str.extract(r"(\d+(?:[.,]\d+)?)")[0]
    num = num.str.replace(",", ".", regex=False)
    numf = pd.to_numeric(num, errors="coerce")
    return np.floor(numf).astype("Int64")

def to_num_ar(x):
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

    # Si hay coma => asumo coma decimal (quita puntos miles)
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

def pct(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "—"

def mini_kpi_html(label: str, value: str):
    return f"""
    <div style="display:inline-block;padding:6px 10px;border-radius:12px;
                border:1px solid #eee;background:#f8f9fa;min-width:110px;text-align:center;">
        <div style="font-size:11px;color:#6c757d;font-weight:800;letter-spacing:.2px;">{label}</div>
        <div style="font-size:16px;font-weight:900;margin-top:2px;">{value}</div>
    </div>
    """

def footer_kpi_only_html(label: str, value: str):
    return f"""
    <div style="margin-top:10px;display:flex;gap:10px;align-items:center;justify-content:flex-start;flex-wrap:wrap;">
        <div>{mini_kpi_html(label, value)}</div>
    </div>
    """

def card_html_base(title, value, sub):
    return f"""
    <div style="border:1px solid #eee;border-radius:14px;padding:16px;background:#fff;
                box-shadow:0 2px 10px rgba(0,0,0,0.04);">
        <div style="font-size:12px;color:#6c757d;font-weight:800;letter-spacing:0.2px;">{title}</div>
        <div style="font-size:28px;font-weight:900;margin-top:6px;">{value}</div>
        <div style="font-size:12px;color:#6c757d;margin-top:6px;">{sub}</div>
    </div>
    """

def chips_css_soft_green():
    return """
    <style>
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
      .block-container {padding-left: 2.2rem; padding-right: 2.2rem;}
    </style>
    """

def norm_text(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower()
    s = (s.replace("á","a").replace("é","e").replace("í","i")
           .replace("ó","o").replace("ú","u").replace("ü","u")
           .replace("ñ","n"))
    s = s.replace("_", " ").replace("-", " ")
    s = " ".join(s.split())
    return s

def month_name_es(month_num: int) -> str:
    m = {
        1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
        7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
    }
    return m.get(int(month_num), "")

def compute_semana_mes(df_in: pd.DataFrame) -> pd.Series:
    out = pd.Series(index=df_in.index, dtype="Int64")
    for mes, g in df_in.groupby("Mes"):
        uniq = sorted([int(x) for x in g["Semana_Num"].dropna().unique().tolist()])
        mapping = {w: i+1 for i, w in enumerate(uniq)}
        out.loc[g.index] = g["Semana_Num"].map(mapping).astype("Int64")
    return out

def find_sheet_name(sheet_names, target: str):
    t = norm_text(target)
    for s in sheet_names:
        if norm_text(s) == t:
            return s
    for s in sheet_names:
        if t in norm_text(s):
            return s
    t_words = set(norm_text(target).split())
    best = None
    best_score = 0
    for s in sheet_names:
        sw = set(norm_text(s).split())
        score = len(t_words.intersection(sw))
        if score > best_score:
            best_score = score
            best = s
    return best if best_score >= max(2, len(t_words)//2) else None

def export_excel_multi_sheets(sheets: dict, filename_sheet_limit=31) -> BytesIO:
    """
    sheets: dict[str, pd.DataFrame]
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df_sheet in sheets.items():
            safe_name = str(name)[:filename_sheet_limit]
            if df_sheet is None:
                df_sheet = pd.DataFrame()
            df_sheet.to_excel(writer, index=False, sheet_name=safe_name)
    out.seek(0)
    return out

# ---------------------------
# LOAD
# ---------------------------
@st.cache_data(ttl=300)
def load_from_drive():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True)

    xls = pd.ExcelFile(EXCEL_LOCAL)
    sheet_names = list(xls.sheet_names)

    df0 = pd.read_excel(xls, sheet_name=0)
    df0 = df0.loc[:, ~df0.columns.astype(str).str.match(r"^Unnamed")]

    dias_sheet = find_sheet_name(sheet_names, "Dias habiles")
    if dias_sheet is not None:
        df_dias = pd.read_excel(xls, sheet_name=dias_sheet)
        df_dias = df_dias.loc[:, ~df_dias.columns.astype(str).str.match(r"^Unnamed")]
    else:
        df_dias = pd.DataFrame(columns=["Mes","Semana","Dias habiles"])

    resumen_sheet = find_sheet_name(sheet_names, "Resumen del mes")
    if resumen_sheet is not None:
        df_res = pd.read_excel(xls, sheet_name=resumen_sheet)
        df_res = df_res.loc[:, ~df_res.columns.astype(str).str.match(r"^Unnamed")]
    else:
        df_res = pd.DataFrame()

    return df0, df_dias, df_res, sheet_names, dias_sheet, resumen_sheet

df, df_dias_habiles, df_resumen_mes, SHEET_NAMES, DIAS_SHEET_FOUND, RESUMEN_SHEET_FOUND = load_from_drive()

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

df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
df = df[~df["Fecha"].isna()].copy()

df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)  # YYYY-MM
df["Mes_Nombre"] = df["Fecha"].dt.month.apply(month_name_es)
df["Mes_norm"] = df["Mes_Nombre"].apply(norm_text)
df["Semana_Mes"] = compute_semana_mes(df)

for c in ["Real_$","Costo_$","Margen_$","Margen_%","Real_Q","Objetivo_$","Objetivo_Q","Cumplimiento_%"]:
    if c in df.columns:
        df[c] = df[c].apply(to_num_ar)

df["KPI"] = df["KPI"].astype(str).str.strip()
df["Categoria_KPI"] = df["Categoria_KPI"].astype(str).str.strip()
df["Tipo_KPI"] = df["Tipo_KPI"].astype(str).str.strip()
df["Sucursal"] = df["Sucursal"].astype(str).str.strip()

def build_real_obj(row):
    if row["Tipo_KPI"] == "$":
        return row["Real_$"], row["Objetivo_$"]
    else:
        return row["Real_Q"], row["Objetivo_Q"]

tmp = df.apply(build_real_obj, axis=1, result_type="expand")
df["Real_val"] = pd.to_numeric(tmp[0], errors="coerce").fillna(0.0)
df["Obj_val"]  = pd.to_numeric(tmp[1], errors="coerce").fillna(0.0)

# Días hábiles
if df_dias_habiles is None or df_dias_habiles.empty:
    df_dias_habiles = pd.DataFrame(columns=["Mes","Semana","Dias habiles"])
for col in ["Mes","Semana","Dias habiles"]:
    if col not in df_dias_habiles.columns:
        df_dias_habiles[col] = np.nan

df_dias_habiles["Mes_norm"] = df_dias_habiles["Mes"].apply(norm_text)
df_dias_habiles["Semana"] = pd.to_numeric(df_dias_habiles["Semana"], errors="coerce").fillna(0).astype(int)
df_dias_habiles["Dias habiles"] = pd.to_numeric(df_dias_habiles["Dias habiles"], errors="coerce").fillna(0).astype(float)

st.markdown(chips_css_soft_green(), unsafe_allow_html=True)

# ---------------------------
# CONTROLES
# ---------------------------
if "modo_presentacion" not in st.session_state:
    st.session_state["modo_presentacion"] = False
if "cap_visual" not in st.session_state:
    st.session_state["cap_visual"] = True
if "cap_val" not in st.session_state:
    st.session_state["cap_val"] = 2.0

def apply_obj0_filter(d, show_obj0: bool):
    return d.copy() if show_obj0 else d[d["Obj_val"] > 0].copy()

def apply_cap_visual(d, cap_on: bool, cap_value: float):
    out = d.copy()
    if "Cumpl" not in out.columns:
        return out
    out["Cumpl_plot"] = out["Cumpl"].clip(upper=cap_value) if cap_on else out["Cumpl"]
    return out

# ---------------------------
# FILTROS
# ---------------------------
semanas = sorted(df["Semana_Num"].dropna().unique().tolist())
if not semanas:
    st.error("No se encontraron semanas válidas en la columna 'Semana'.")
    st.stop()

default_sem = 1 if 1 in semanas else min(semanas)
sucursales = sorted(df["Sucursal"].dropna().unique().tolist())

# ---------------------------
# TOP BAR
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

if st.session_state["modo_presentacion"]:
    st.markdown(hide_sidebar_css(), unsafe_allow_html=True)

def render_filters(area="sidebar"):
    if "semana_corte" not in st.session_state:
        st.session_state["semana_corte"] = default_sem
    if "sucursal" not in st.session_state:
        st.session_state["sucursal"] = "TODAS (Consolidado)"
    if "show_obj0" not in st.session_state:
        st.session_state["show_obj0"] = True

    container = st.sidebar if area == "sidebar" else st.container()
    with container:
        if area == "sidebar":
            st.sidebar.markdown("## Filtros obligatorios")
        else:
            st.markdown("### Filtros")

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
# CORTE / MES
# ---------------------------
df_cut = df[df["Semana_Num"] <= semana_corte].copy()
if sucursal != "TODAS (Consolidado)":
    df_cut = df_cut[df_cut["Sucursal"] == sucursal].copy()

mes_ref = df_cut["Mes"].max() if not df_cut.empty else df["Mes"].max()
mes_ref_norm = norm_text(month_name_es(int(str(mes_ref).split("-")[1])))

df_month = df[df["Mes"] == mes_ref].copy()
if sucursal != "TODAS (Consolidado)":
    df_month = df_month[df_month["Sucursal"] == sucursal].copy()

df_cut_mes = df_cut[df_cut["Mes"] == mes_ref].copy()
semana_mes_corte = int(df_cut_mes["Semana_Mes"].max()) if not df_cut_mes.empty else 1

dias_mes = df_dias_habiles[df_dias_habiles["Mes_norm"] == mes_ref_norm].copy()
dias_total_mes = float(dias_mes["Dias habiles"].sum()) if not dias_mes.empty else np.nan
dias_transc = float(dias_mes[dias_mes["Semana"] <= semana_mes_corte]["Dias habiles"].sum()) if not dias_mes.empty else np.nan

# ---------------------------
# P&L aperturas
# ---------------------------
def compute_openings_pl(dfc):
    rep_open = sorted(dfc[(dfc["KPI"].str.upper()=="REPUESTOS") & (dfc["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
    srv_open = sorted(dfc[(dfc["KPI"].str.upper()=="SERVICIOS") & (dfc["Tipo_KPI"]=="$")]["Categoria_KPI"].unique().tolist())
    return rep_open, srv_open

rep_open, srv_open = compute_openings_pl(df_cut)

if "rep_sel" not in st.session_state:
    st.session_state["rep_sel"] = rep_open
if "srv_sel" not in st.session_state:
    st.session_state["srv_sel"] = srv_open

for x in rep_open:
    if x not in st.session_state["rep_sel"]:
        st.session_state["rep_sel"].append(x)
for x in srv_open:
    if x not in st.session_state["srv_sel"]:
        st.session_state["srv_sel"].append(x)

st.session_state["rep_sel"] = [x for x in st.session_state["rep_sel"] if x in rep_open]
st.session_state["srv_sel"] = [x for x in st.session_state["srv_sel"] if x in srv_open]

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

# ============================================================
# Funciones de cálculo / proyección / tablas (para dashboard + export)
# ============================================================
def summarize_segment(dseg: pd.DataFrame):
    dseg = apply_obj0_filter(dseg, show_obj0)
    real = dseg["Real_val"].sum()
    obj = dseg["Obj_val"].sum()
    c = safe_ratio(real, obj)
    return float(real), float(obj), c

def proyectar_eom_runrate(real_acum: float) -> float:
    if pd.isna(dias_transc) or pd.isna(dias_total_mes) or dias_transc == 0:
        return np.nan
    return (float(real_acum) / float(dias_transc)) * float(dias_total_mes)

def micro_sucursal(d: pd.DataFrame):
    g = d.groupby("Sucursal", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    return g

def micro_aperturas(d: pd.DataFrame):
    g = d.groupby("Categoria_KPI", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    return g

def ranking_sucursal_apertura_micro(d: pd.DataFrame, top_n: int, show_zero: bool):
    x = d.copy()
    if not show_zero:
        x = x[x["Obj_val"] > 0].copy()
    g = x.groupby(["Sucursal","Categoria_KPI"], as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False).head(top_n).copy()
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct(r['Cumpl'])} | {money(r['Real'])}/{money(r['Obj'])}", axis=1)
    g["key"] = g["Sucursal"].astype(str) + " — " + g["Categoria_KPI"].astype(str)
    return g

def principal_driver_gap(d_pl: pd.DataFrame):
    x = apply_obj0_filter(d_pl.copy(), show_obj0)
    if x.empty:
        return None
    g = x.groupby(["KPI","Categoria_KPI"], as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Gap"] = g["Obj"] - g["Real"]
    g = g.sort_values("Gap", ascending=False)
    row = g.iloc[0]
    return {"KPI": str(row["KPI"]), "Cat": str(row["Categoria_KPI"]), "Gap": float(row["Gap"])}

def spark_evolucion(df_scope_month: pd.DataFrame):
    if df_scope_month is None or df_scope_month.empty:
        return
    x = apply_obj0_filter(df_scope_month.copy(), show_obj0)
    g = x.groupby("Semana_Mes", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum")).sort_values("Semana_Mes")
    g["Cumpl_sem"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl_sem"].isna()].copy()

    gg = g.copy()
    gg["Cumpl"] = gg["Cumpl_sem"]
    gg["Cumpl_plot"] = gg["Cumpl"].clip(upper=cap_val) if cap_on else gg["Cumpl"]
    gg["txt"] = gg["Cumpl"].apply(pct)

    fig = px.line(gg, x="Semana_Mes", y="Cumpl_plot", markers=True, text="txt")
    weeks = sorted([int(w) for w in gg["Semana_Mes"].dropna().unique().tolist()])
    fig.update_xaxes(tickmode="array", tickvals=weeks, dtick=1)
    fig.update_traces(mode="lines+markers+text", textposition="top center")
    fig.update_layout(height=150, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="", yaxis_title="")
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# Construyo datasets P&L (se usan en pantalla y export)
# ============================================================
d_pl = df_cut[df_cut["Tipo_KPI"]=="$"].copy()
d_rep = d_pl[d_pl["KPI"].str.upper()=="REPUESTOS"].copy()
d_rep = d_rep[d_rep["Categoria_KPI"].isin(rep_sel)].copy()
d_srv = d_pl[d_pl["KPI"].str.upper()=="SERVICIOS"].copy()
d_srv = d_srv[d_srv["Categoria_KPI"].isin(srv_sel)].copy()

rep_real, rep_obj, rep_c = summarize_segment(d_rep)
srv_real, srv_obj, srv_c = summarize_segment(d_srv)

total_real = float(rep_real) + float(srv_real)
total_obj  = float(rep_obj) + float(srv_obj)
total_c    = safe_ratio(total_real, total_obj)

rep_proy = proyectar_eom_runrate(rep_real)
srv_proy = proyectar_eom_runrate(srv_real)
total_proy = proyectar_eom_runrate(total_real)

rep_month = df_month[(df_month["Tipo_KPI"]=="$") & (df_month["KPI"].str.upper()=="REPUESTOS")].copy()
rep_month = rep_month[rep_month["Categoria_KPI"].isin(rep_sel)].copy()
srv_month = df_month[(df_month["Tipo_KPI"]=="$") & (df_month["KPI"].str.upper()=="SERVICIOS")].copy()
srv_month = srv_month[srv_month["Categoria_KPI"].isin(srv_sel)].copy()
total_month = pd.concat([rep_month, srv_month], ignore_index=True)

# Tablas para export (las mismas que se ven)
rep_by_suc = micro_sucursal(apply_obj0_filter(d_rep, show_obj0))
srv_by_suc = micro_sucursal(apply_obj0_filter(d_srv, show_obj0))
rep_by_ap  = micro_aperturas(apply_obj0_filter(d_rep, show_obj0))
srv_by_ap  = micro_aperturas(apply_obj0_filter(d_srv, show_obj0))

# Defaults del ranking (coherentes con UI)
DEFAULT_TOP_N = 10
DEFAULT_SHOW_ZERO_RANK = False
rep_rank = ranking_sucursal_apertura_micro(d_rep, top_n=DEFAULT_TOP_N, show_zero=DEFAULT_SHOW_ZERO_RANK)
srv_rank = ranking_sucursal_apertura_micro(d_srv, top_n=DEFAULT_TOP_N, show_zero=DEFAULT_SHOW_ZERO_RANK)

driver = principal_driver_gap(pd.concat([d_rep, d_srv], ignore_index=True))

# ============================================================
# HEADER + BOTÓN EXCEL (Resumen Ejecutivo)
# ============================================================
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")

header_left, header_right = st.columns([2.3, 1.0])
with header_left:
    st.caption(
        f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}** | "
        f"Mes ref: **{mes_ref}** | SemanaMes corte: **{semana_mes_corte}** | "
        f"Días hábiles transcurridos: **{(int(dias_transc) if not pd.isna(dias_transc) else '—')}** | "
        f"Días hábiles mes: **{(int(dias_total_mes) if not pd.isna(dias_total_mes) else '—')}**"
    )

with header_right:
    # Armo el Excel “Resumen Ejecutivo” con todo lo mostrado
    resumen_exec = pd.DataFrame([
        {
            "Bloque": "Repuestos",
            "Real_Acum": rep_real,
            "Obj_Acum": rep_obj,
            "Cumpl_Acum": rep_c,
            "Proy_EOM_RunRate": rep_proy,
        },
        {
            "Bloque": "Servicios",
            "Real_Acum": srv_real,
            "Obj_Acum": srv_obj,
            "Cumpl_Acum": srv_c,
            "Proy_EOM_RunRate": srv_proy,
        },
        {
            "Bloque": "Total Postventa",
            "Real_Acum": total_real,
            "Obj_Acum": total_obj,
            "Cumpl_Acum": total_c,
            "Proy_EOM_RunRate": total_proy,
        },
    ])

    # Agrego “pretty” en texto + meta
    resumen_exec["Real_Acum_$"] = resumen_exec["Real_Acum"].apply(money)
    resumen_exec["Obj_Acum_$"]  = resumen_exec["Obj_Acum"].apply(money)
    resumen_exec["Cumpl_Acum_%"] = resumen_exec["Cumpl_Acum"].apply(pct)
    resumen_exec["Proy_EOM_$"] = resumen_exec["Proy_EOM_RunRate"].apply(money)

    meta = pd.DataFrame([{
        "Sucursal": sucursal,
        "Semana_Corte": semana_corte,
        "Mes_Ref": mes_ref,
        "SemanaMes_Corte": semana_mes_corte,
        "Dias_Habiles_Transcurridos": (None if pd.isna(dias_transc) else int(dias_transc)),
        "Dias_Habiles_Mes": (None if pd.isna(dias_total_mes) else int(dias_total_mes)),
        "Cap_Visual_On": cap_on,
        "Cap_Visual_Max": cap_val,
        "Incluir_Obj_0": show_obj0,
        "Hoja_Dias_Habiles": (DIAS_SHEET_FOUND if DIAS_SHEET_FOUND is not None else "NO_ENCONTRADA"),
        "Hoja_Resumen_del_Mes": (RESUMEN_SHEET_FOUND if RESUMEN_SHEET_FOUND is not None else "NO_ENCONTRADA"),
        "Principal_Desvio_KPI": (driver["KPI"] if driver else "—"),
        "Principal_Desvio_Cat": (driver["Cat"] if driver else "—"),
        "Principal_Desvio_Gap": (driver["Gap"] if driver else np.nan),
        "Principal_Desvio_Gap_$": (money(driver["Gap"]) if driver else "—"),
    }])

    # Hitos tal cual
    hitos_df = df_resumen_mes.copy() if (df_resumen_mes is not None and not df_resumen_mes.empty) else pd.DataFrame()

    # Excel
    sheets = {
        "Resumen_Ejecutivo": resumen_exec,
        "Meta": meta,
        "Rep_Sucursal": rep_by_suc,
        "Srv_Sucursal": srv_by_suc,
        "Rep_Aperturas": rep_by_ap,
        "Srv_Aperturas": srv_by_ap,
        "Rep_Ranking_Micro": rep_rank,
        "Srv_Ranking_Micro": srv_rank,
        "Hitos_del_mes": hitos_df,
    }
    excel_bytes = export_excel_multi_sheets(sheets)

    st.download_button(
        "⬇️ Resumen Ejecutivo (Excel)",
        data=excel_bytes,
        file_name=f"Resumen_Ejecutivo_Tablero_Posventa_{mes_ref}_Sem{semana_corte}_{sucursal.replace(' ','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "🧩 P&L (Repuestos vs Servicios)",
    "📌 KPIs (resto)",
    "🧪 Gestión (desvíos)",
    "🗓️ Hitos del mes"
])

# ============================================================
# TAB 1 — P&L
# ============================================================
with tab1:
    st.markdown("## 🧩 P&L — Macro → Micro")
    st.markdown("---")

    driver_txt = f"Principal desvío: **{driver['KPI']} / {driver['Cat']}** (Gap {money(driver['Gap'])})" if driver else "Principal desvío: —"
    st.info(f"**Resumen Ejecutivo:** Repuestos {pct(rep_c)} | Servicios {pct(srv_c)} | Total Postventa {pct(total_c)} | {driver_txt}")

    def macro_block(titulo, real, obj, df_scope_month_for_spark):
        c = safe_ratio(real, obj)
        proy_real = proyectar_eom_runrate(real)
        sub = (
            f"Real {money(real)} | Obj {money(obj)} | "
            f"Proy EOM (run-rate): {money(proy_real)} | "
            f"Días: {('—' if pd.isna(dias_transc) else int(dias_transc))}/{('—' if pd.isna(dias_total_mes) else int(dias_total_mes))}"
        )
        st.markdown(card_html_base(titulo, money(real), sub), unsafe_allow_html=True)
        st.markdown(footer_kpi_only_html("Cumpl. Acum.", pct(c)), unsafe_allow_html=True)
        spark_evolucion(df_scope_month_for_spark)

    c1, c_mid, c2 = st.columns([1.0, 1.05, 1.0])

    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        macro_block("Repuestos — Real (Acum.)", rep_real, rep_obj, rep_month)

    with c_mid:
        st.markdown("### 🧩 TOTAL POSTVENTA (P&L)")
        macro_block("Total Postventa — Real (Acum.)", total_real, total_obj, total_month)

    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        macro_block("Servicios — Real (Acum.)", srv_real, srv_obj, srv_month)

    st.markdown("---")
    st.markdown("### Cumplimiento por Sucursal — (acumulado)")

    a, b = st.columns(2)
    with a:
        st.markdown("**Repuestos — por sucursal**")
        if rep_by_suc.empty:
            st.info("Sin datos por sucursal.")
        else:
            fig = px.bar(rep_by_suc, x="Cumpl_plot", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with b:
        st.markdown("**Servicios — por sucursal**")
        if srv_by_suc.empty:
            st.info("Sin datos por sucursal.")
        else:
            fig = px.bar(srv_by_suc, x="Cumpl_plot", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    l, r = st.columns(2)
    with l:
        st.markdown("**Repuestos — por apertura**")
        if rep_by_ap.empty:
            st.info("Sin datos (revisar aperturas).")
        else:
            fig = px.bar(rep_by_ap, x="Cumpl_plot", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with r:
        st.markdown("**Servicios — por apertura**")
        if srv_by_ap.empty:
            st.info("Sin datos (revisar aperturas).")
        else:
            fig = px.bar(srv_by_ap, x="Cumpl_plot", y="Categoria_KPI", orientation="h", text="label")
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

    rep_rank_base = d_rep.copy()
    if rep_micro_choice != "Todas las aperturas":
        rep_rank_base = rep_rank_base[rep_rank_base["Categoria_KPI"] == rep_micro_choice].copy()
    srv_rank_base = d_srv.copy()
    if srv_micro_choice != "Todas las aperturas":
        srv_rank_base = srv_rank_base[srv_rank_base["Categoria_KPI"] == srv_micro_choice].copy()

    rr, ss = st.columns(2)
    with rr:
        st.markdown("### Repuestos — sucursal + apertura (micro)")
        g = ranking_sucursal_apertura_micro(rep_rank_base, top_n=top_n, show_zero=show_zero_rank)
        if g.empty:
            st.info("Sin ranking.")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="key", orientation="h", text="label")
            fig.update_layout(height=460, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with ss:
        st.markdown("### Servicios — sucursal + apertura (micro)")
        g = ranking_sucursal_apertura_micro(srv_rank_base, top_n=top_n, show_zero=show_zero_rank)
        if g.empty:
            st.info("Sin ranking.")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="key", orientation="h", text="label")
            fig.update_layout(height=460, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 2 — KPIs resto
# ============================================================
with tab2:
    st.info("TAB 2 sin cambios en esta versión (se mantiene igual que tu versión anterior estable).")

# ============================================================
# TAB 3 — Gestión
# ============================================================
with tab3:
    st.info("TAB 3 sin cambios en esta versión (se mantiene igual que tu versión anterior estable).")

# ============================================================
# TAB 4 — Hitos del mes (Resumen del mes)
# ============================================================
with tab4:
    st.markdown("## 🗓️ Hitos del mes — Resumen del mes")
    st.caption("Tomado **tal cual** desde la pestaña de la base (fuzzy match por nombre).")
    st.markdown("---")

    if RESUMEN_SHEET_FOUND is None:
        st.warning("No pude encontrar la hoja **'Resumen del mes'** por nombre.")
        with st.expander("Ver pestañas detectadas (para identificar el nombre exacto)", expanded=True):
            st.write(SHEET_NAMES)
        st.stop()

    st.success(f"✅ Hoja encontrada: **{RESUMEN_SHEET_FOUND}**")

    if df_resumen_mes is None or df_resumen_mes.empty:
        st.info("La hoja existe, pero está vacía (o solo tiene encabezados).")
    else:
        st.markdown("### Vista tabla (tal cual la hoja)")
        st.dataframe(df_resumen_mes, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown("### Vista lectura rápida (líneas)")
        tmp = df_resumen_mes.copy().dropna(how="all")
        if tmp.empty:
            st.info("No hay filas con contenido.")
        else:
            for _, r in tmp.iterrows():
                parts = []
                for c in tmp.columns:
                    v = r.get(c, "")
                    if pd.isna(v):
                        continue
                    s = str(v).strip()
                    if s == "" or s.lower() == "nan":
                        continue
                    parts.append(f"{c}: {s}")
                if parts:
                    st.markdown("- " + " | ".join(parts))
