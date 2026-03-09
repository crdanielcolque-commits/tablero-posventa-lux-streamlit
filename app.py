# ============================================================
# TABLERO POSVENTA — MACRO → MICRO (Semanal + Acumulado) v2.3
# v2.3.16
# ✅ Filtro de MES
# ✅ Semana corte limitada al mes seleccionado
# ✅ Proyección / días hábiles / export Excel según mes elegido
# ✅ TAB2 y TAB3 completos
# ✅ Excel profesional resumido
# ============================================================

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

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

    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_dot > last_comma:
            s = s.replace(",", "")
        else:
            s = s.replace(".", "")
            s = s.replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        pass

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

def money_str(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"${float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def qty_str(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "—"

def pct_str(x):
    if x is None or pd.isna(x):
        return "—"
    try:
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "—"

def card_html_base(title, value, sub):
    return f"""
    <div style="border:1px solid #eee;border-radius:14px;padding:16px;background:#fff;
                box-shadow:0 2px 10px rgba(0,0,0,0.04);">
        <div style="font-size:12px;color:#6c757d;font-weight:800;letter-spacing:0.2px;">{title}</div>
        <div style="font-size:28px;font-weight:900;margin-top:6px;">{value}</div>
        <div style="font-size:12px;color:#6c757d;margin-top:6px;">{sub}</div>
    </div>
    """

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

def build_month_label(mes_key: str) -> str:
    try:
        p = pd.Period(mes_key, freq="M")
        return f"{month_name_es(p.month)} {p.year}"
    except Exception:
        return str(mes_key)

# ---------------------------
# EXCEL EXPORT PROFESIONAL
# ---------------------------
def _autosize_ws(ws, min_w=10, max_w=55):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(val))
            except Exception:
                pass
        width = max(min_w, min(max_w, max_len + 2))
        ws.column_dimensions[col_letter].width = width

def _style_table(ws, freeze_row=1):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    ws.freeze_panes = ws[f"A{freeze_row+1}"]
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = left

def _format_columns(ws, col_formats: dict):
    headers = [c.value for c in ws[1]]
    name_to_idx = {str(h): i+1 for i, h in enumerate(headers) if h is not None}
    for col_name, fmt in col_formats.items():
        if col_name in name_to_idx:
            idx = name_to_idx[col_name]
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=idx).number_format = fmt

def build_exec_excel_professional(
    meta_df: pd.DataFrame,
    resumen_df: pd.DataFrame,
    pl_sucursal_rep: pd.DataFrame,
    pl_sucursal_srv: pd.DataFrame,
    pl_ap_rep: pd.DataFrame,
    pl_ap_srv: pd.DataFrame,
    ranking_rep: pd.DataFrame,
    ranking_srv: pd.DataFrame,
    hitos_df: pd.DataFrame,
) -> BytesIO:
    def clean(df_in: pd.DataFrame, keep_cols=None, drop_cols=None):
        if df_in is None:
            return pd.DataFrame()
        dfc = df_in.copy()
        if drop_cols:
            for c in drop_cols:
                if c in dfc.columns:
                    dfc = dfc.drop(columns=[c])
        if keep_cols:
            cols = [c for c in keep_cols if c in dfc.columns]
            dfc = dfc[cols].copy()
        return dfc

    resumen_clean = clean(
        resumen_df,
        keep_cols=["Bloque","Real_Acum","Obj_Acum","Cumpl_Acum","Proy_EOM_RunRate","Dias_Transc","Dias_Mes"]
    )

    rep_s = clean(pl_sucursal_rep, drop_cols=["label","Cumpl_plot"])
    srv_s = clean(pl_sucursal_srv, drop_cols=["label","Cumpl_plot"])
    if not rep_s.empty:
        rep_s.insert(0, "Bloque", "Repuestos")
    if not srv_s.empty:
        srv_s.insert(0, "Bloque", "Servicios")
    pl_suc = pd.concat([rep_s, srv_s], ignore_index=True) if (not rep_s.empty or not srv_s.empty) else pd.DataFrame()

    rep_a = clean(pl_ap_rep, drop_cols=["label","Cumpl_plot"])
    srv_a = clean(pl_ap_srv, drop_cols=["label","Cumpl_plot"])
    if not rep_a.empty:
        rep_a.insert(0, "Bloque", "Repuestos")
    if not srv_a.empty:
        srv_a.insert(0, "Bloque", "Servicios")
    pl_ap = pd.concat([rep_a, srv_a], ignore_index=True) if (not rep_a.empty or not srv_a.empty) else pd.DataFrame()

    rep_r = clean(ranking_rep, drop_cols=["label","Cumpl_plot","key"])
    srv_r = clean(ranking_srv, drop_cols=["label","Cumpl_plot","key"])
    if not rep_r.empty:
        rep_r.insert(0, "Bloque", "Repuestos")
    if not srv_r.empty:
        srv_r.insert(0, "Bloque", "Servicios")
    rank = pd.concat([rep_r, srv_r], ignore_index=True) if (not rep_r.empty or not srv_r.empty) else pd.DataFrame()

    hitos = pd.DataFrame() if hitos_df is None else hitos_df.copy()

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        meta_df.to_excel(writer, index=False, sheet_name="00_Meta")
        resumen_clean.to_excel(writer, index=False, sheet_name="01_Resumen")
        pl_suc.to_excel(writer, index=False, sheet_name="02_P&L_Sucursal")
        pl_ap.to_excel(writer, index=False, sheet_name="03_P&L_Aperturas")
        rank.to_excel(writer, index=False, sheet_name="04_Ranking_Micro")
        hitos.to_excel(writer, index=False, sheet_name="05_Hitos_mes")

    out.seek(0)
    wb = load_workbook(out)

    ws = wb["00_Meta"]
    _style_table(ws)
    _autosize_ws(ws)

    ws = wb["01_Resumen"]
    _style_table(ws)
    _format_columns(ws, {
        "Real_Acum": '"$"#,##0',
        "Obj_Acum": '"$"#,##0',
        "Cumpl_Acum": '0.0%',
        "Proy_EOM_RunRate": '"$"#,##0',
        "Dias_Transc": '0',
        "Dias_Mes": '0',
    })
    _autosize_ws(ws, min_w=12, max_w=40)

    ws = wb["02_P&L_Sucursal"]
    _style_table(ws)
    _format_columns(ws, {
        "Real": '"$"#,##0',
        "Obj": '"$"#,##0',
        "Cumpl": '0.0%',
    })
    _autosize_ws(ws)

    ws = wb["03_P&L_Aperturas"]
    _style_table(ws)
    _format_columns(ws, {
        "Real": '"$"#,##0',
        "Obj": '"$"#,##0',
        "Cumpl": '0.0%',
    })
    _autosize_ws(ws)

    ws = wb["04_Ranking_Micro"]
    _style_table(ws)
    _format_columns(ws, {
        "Real": '"$"#,##0',
        "Obj": '"$"#,##0',
        "Cumpl": '0.0%',
    })
    _autosize_ws(ws)

    ws = wb["05_Hitos_mes"]
    _style_table(ws)
    _autosize_ws(ws, min_w=12, max_w=70)

    out2 = BytesIO()
    wb.save(out2)
    out2.seek(0)
    return out2

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

df["Mes"] = df["Fecha"].dt.to_period("M").astype(str)
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

if df_dias_habiles is None or df_dias_habiles.empty:
    df_dias_habiles = pd.DataFrame(columns=["Mes","Semana","Dias habiles"])
for col in ["Mes","Semana","Dias habiles"]:
    if col not in df_dias_habiles.columns:
        df_dias_habiles[col] = np.nan

df_dias_habiles["Mes_norm"] = df_dias_habiles["Mes"].apply(norm_text)
df_dias_habiles["Semana"] = pd.to_numeric(df_dias_habiles["Semana"], errors="coerce").fillna(0).astype(int)
df_dias_habiles["Dias habiles"] = pd.to_numeric(df_dias_habiles["Dias habiles"], errors="coerce").fillna(0).astype(float)

# ---------------------------
# CONTROLES VISUALES
# ---------------------------
if "modo_presentacion" not in st.session_state:
    st.session_state["modo_presentacion"] = False
if "cap_visual" not in st.session_state:
    st.session_state["cap_visual"] = True
if "cap_val" not in st.session_state:
    st.session_state["cap_val"] = 2.0

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
    st.caption("Tip: el cap es solo visual (ranking/gráficos). No altera el cálculo base.")

if st.session_state["modo_presentacion"]:
    st.markdown(hide_sidebar_css(), unsafe_allow_html=True)

# ---------------------------
# FILTROS BASE
# ---------------------------
def apply_obj0_filter(d, show_obj0: bool):
    return d.copy() if show_obj0 else d[d["Obj_val"] > 0].copy()

def apply_cap_visual(d, cap_on: bool, cap_value: float):
    out = d.copy()
    if "Cumpl" not in out.columns:
        return out
    out["Cumpl_plot"] = out["Cumpl"].clip(upper=cap_value) if cap_on else out["Cumpl"]
    return out

meses_disponibles = sorted(df["Mes"].dropna().unique().tolist())
if not meses_disponibles:
    st.error("No se encontraron meses válidos en la base.")
    st.stop()

month_labels = {m: build_month_label(m) for m in meses_disponibles}
meses_labels_ordenados = [month_labels[m] for m in meses_disponibles]
label_to_mes = {v: k for k, v in month_labels.items()}

default_mes = meses_disponibles[-1]

def render_filters(area="sidebar"):
    if "mes_sel" not in st.session_state:
        st.session_state["mes_sel"] = default_mes
    if "semana_corte" not in st.session_state:
        st.session_state["semana_corte"] = None
    if "sucursal" not in st.session_state:
        st.session_state["sucursal"] = "TODAS (Consolidado)"
    if "show_obj0" not in st.session_state:
        st.session_state["show_obj0"] = True

    container = st.sidebar if area == "sidebar" else st.container()

    # Semanas del mes elegido
    semanas_mes = sorted(
        df.loc[df["Mes"] == st.session_state["mes_sel"], "Semana_Num"].dropna().unique().tolist()
    )
    if not semanas_mes:
        semanas_mes = []

    if st.session_state["semana_corte"] not in semanas_mes:
        st.session_state["semana_corte"] = semanas_mes[-1] if semanas_mes else None

    sucursales = sorted(df["Sucursal"].dropna().unique().tolist())

    with container:
        if area == "sidebar":
            st.sidebar.markdown("## Filtros obligatorios")
        else:
            st.markdown("### Filtros")

        mes_label_actual = month_labels.get(st.session_state["mes_sel"], st.session_state["mes_sel"])
        mes_sel_label = st.selectbox(
            "Mes",
            meses_labels_ordenados,
            index=meses_labels_ordenados.index(mes_label_actual) if mes_label_actual in meses_labels_ordenados else len(meses_labels_ordenados)-1,
            key=f"mes_{area}"
        )
        st.session_state["mes_sel"] = label_to_mes[mes_sel_label]

        # recalcular semanas del mes luego del cambio
        semanas_mes = sorted(
            df.loc[df["Mes"] == st.session_state["mes_sel"], "Semana_Num"].dropna().unique().tolist()
        )
        if st.session_state["semana_corte"] not in semanas_mes:
            st.session_state["semana_corte"] = semanas_mes[-1] if semanas_mes else None

        st.session_state["semana_corte"] = st.selectbox(
            "Semana corte",
            semanas_mes,
            index=semanas_mes.index(st.session_state["semana_corte"]) if st.session_state["semana_corte"] in semanas_mes else 0,
            key=f"semana_{area}"
        )

        st.session_state["sucursal"] = st.selectbox(
            "Sucursal",
            ["TODAS (Consolidado)"] + sucursales,
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

mes_sel = st.session_state["mes_sel"]
semana_corte = int(st.session_state["semana_corte"])
sucursal = st.session_state["sucursal"]
show_obj0 = bool(st.session_state["show_obj0"])
cap_on = bool(st.session_state["cap_visual"])
cap_val = float(st.session_state["cap_val"])

# ---------------------------
# DATOS DEL MES SELECCIONADO
# ---------------------------
df_mes = df[df["Mes"] == mes_sel].copy()
if sucursal != "TODAS (Consolidado)":
    df_mes = df_mes[df_mes["Sucursal"] == sucursal].copy()

df_cut = df_mes[df_mes["Semana_Num"] <= semana_corte].copy()

mes_ref = mes_sel
mes_ref_norm = norm_text(month_name_es(int(str(mes_ref).split("-")[1])))

df_month = df[df["Mes"] == mes_ref].copy()
if sucursal != "TODAS (Consolidado)":
    df_month = df_month[df_month["Sucursal"] == sucursal].copy()

df_cut_mes = df_cut.copy()
semana_mes_corte = int(df_cut_mes["Semana_Mes"].max()) if not df_cut_mes.empty else 1

dias_mes = df_dias_habiles[df_dias_habiles["Mes_norm"] == mes_ref_norm].copy()
dias_total_mes = float(dias_mes["Dias habiles"].sum()) if not dias_mes.empty else np.nan
dias_transc = float(dias_mes[dias_mes["Semana"] <= semana_mes_corte]["Dias habiles"].sum()) if not dias_mes.empty else np.nan

# ---------------------------
# P&L aperturas del mes elegido
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

# ---------------------------
# FUNCIONES DE NEGOCIO
# ---------------------------
def summarize_segment(dseg: pd.DataFrame):
    dseg = apply_obj0_filter(dseg, show_obj0)
    real = dseg["Real_val"].sum()
    obj = dseg["Obj_val"].sum()
    c = safe_ratio(real, obj)
    return float(real), float(obj), c

def proyectar_eom_runrate(real_acum: float, dias_trans: float, dias_mes_total: float) -> float:
    if pd.isna(dias_trans) or pd.isna(dias_mes_total) or dias_trans == 0:
        return np.nan
    return (float(real_acum) / float(dias_trans)) * float(dias_mes_total)

def micro_sucursal(d: pd.DataFrame):
    g = d.groupby("Sucursal", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct_str(r['Cumpl'])} | {money_str(r['Real'])}/{money_str(r['Obj'])}", axis=1)
    return g

def micro_aperturas(d: pd.DataFrame):
    g = d.groupby("Categoria_KPI", as_index=False).agg(Real=("Real_val","sum"), Obj=("Obj_val","sum"))
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct_str(r['Cumpl'])} | {money_str(r['Real'])}/{money_str(r['Obj'])}", axis=1)
    return g

def ranking_sucursal_apertura_micro(d: pd.DataFrame, top_n: int, show_zero: bool):
    x = d.copy()
    if not show_zero:
        x = x[x["Obj_val"] > 0].copy()

    g = x.groupby(["Sucursal","Categoria_KPI"], as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
    g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False).head(top_n).copy()
    g = apply_cap_visual(g, cap_on, cap_val)
    g["label"] = g.apply(lambda r: f"{pct_str(r['Cumpl'])} | {money_str(r['Real'])}/{money_str(r['Obj'])}", axis=1)
    g["key"] = g["Sucursal"].astype(str) + " — " + g["Categoria_KPI"].astype(str)
    return g

def principal_driver_gap(d_pl: pd.DataFrame):
    x = apply_obj0_filter(d_pl.copy(), show_obj0)
    if x.empty:
        return None
    g = x.groupby(["KPI","Categoria_KPI"], as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )
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
    gg["txt"] = gg["Cumpl"].apply(pct_str)

    fig = px.line(gg, x="Semana_Mes", y="Cumpl_plot", markers=True, text="txt")
    weeks = sorted([int(w) for w in gg["Semana_Mes"].dropna().unique().tolist()])
    fig.update_xaxes(tickmode="array", tickvals=weeks, dtick=1)
    fig.update_traces(mode="lines+markers+text", textposition="top center")
    fig.update_layout(height=150, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="", yaxis_title="")
    fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

def filter_hitos_by_month(df_hitos: pd.DataFrame, mes_key: str) -> pd.DataFrame:
    """
    Si existe una columna tipo Mes/Periodo/Fecha, intenta filtrar por mes.
    Si no puede, devuelve toda la hoja.
    """
    if df_hitos is None or df_hitos.empty:
        return pd.DataFrame()

    tmp = df_hitos.copy()
    month_cols = []
    for c in tmp.columns:
        nc = norm_text(c)
        if nc in {"mes", "periodo", "periodo mes", "fecha", "month"} or "mes" in nc or "periodo" in nc:
            month_cols.append(c)

    if not month_cols:
        return tmp

    mes_label = build_month_label(mes_key)
    mes_norm = norm_text(mes_label)
    mes_key_norm = norm_text(mes_key)
    nombre_mes = norm_text(month_name_es(int(str(mes_key).split("-")[1])))

    mask = pd.Series(False, index=tmp.index)
    for c in month_cols:
        s = tmp[c].astype(str).apply(norm_text)
        mask = mask | s.str.contains(mes_key_norm, na=False) | s.str.contains(nombre_mes, na=False) | s.str.contains(mes_norm, na=False)

    filtrado = tmp[mask].copy()
    return filtrado if not filtrado.empty else tmp

# ---------------------------
# HEADER + EXCEL
# ---------------------------
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")

hl, hr = st.columns([2.2, 1.0])
with hl:
    st.caption(
        f"Sucursal: **{sucursal}** | Mes: **{build_month_label(mes_sel)}** | "
        f"Corte semana **{semana_corte}** | SemanaMes corte: **{semana_mes_corte}** | "
        f"Días hábiles transcurridos: **{(int(dias_transc) if not pd.isna(dias_transc) else '—')}** | "
        f"Días hábiles mes: **{(int(dias_total_mes) if not pd.isna(dias_total_mes) else '—')}**"
    )

with hr:
    semanas_mes_export = sorted(df.loc[df["Mes"] == mes_sel, "Semana_Num"].dropna().unique().tolist())
    sem_export = st.selectbox(
        "Semana para Excel",
        semanas_mes_export,
        index=semanas_mes_export.index(semana_corte) if semana_corte in semanas_mes_export else len(semanas_mes_export)-1,
        help="El Excel se calcula hasta esta semana del mes seleccionado."
    )

# dataset export según mes_sel + sem_export
df_cut_xls = df[df["Mes"] == mes_sel].copy()
if sucursal != "TODAS (Consolidado)":
    df_cut_xls = df_cut_xls[df_cut_xls["Sucursal"] == sucursal].copy()
df_cut_xls = df_cut_xls[df_cut_xls["Semana_Num"] <= int(sem_export)].copy()

mes_ref_xls = mes_sel
mes_ref_norm_xls = norm_text(month_name_es(int(str(mes_ref_xls).split("-")[1])))

df_month_xls = df[df["Mes"] == mes_ref_xls].copy()
if sucursal != "TODAS (Consolidado)":
    df_month_xls = df_month_xls[df_month_xls["Sucursal"] == sucursal].copy()

df_cut_mes_xls = df_cut_xls.copy()
semana_mes_corte_xls = int(df_cut_mes_xls["Semana_Mes"].max()) if not df_cut_mes_xls.empty else 1

dias_mes_xls = df_dias_habiles[df_dias_habiles["Mes_norm"] == mes_ref_norm_xls].copy()
dias_total_mes_xls = float(dias_mes_xls["Dias habiles"].sum()) if not dias_mes_xls.empty else np.nan
dias_transc_xls = float(dias_mes_xls[dias_mes_xls["Semana"] <= semana_mes_corte_xls]["Dias habiles"].sum()) if not dias_mes_xls.empty else np.nan

# P&L export
d_pl_xls = df_cut_xls[df_cut_xls["Tipo_KPI"]=="$"].copy()
d_rep_xls = d_pl_xls[d_pl_xls["KPI"].str.upper()=="REPUESTOS"].copy()
d_rep_xls = d_rep_xls[d_rep_xls["Categoria_KPI"].isin(rep_sel)].copy()
d_srv_xls = d_pl_xls[d_pl_xls["KPI"].str.upper()=="SERVICIOS"].copy()
d_srv_xls = d_srv_xls[d_srv_xls["Categoria_KPI"].isin(srv_sel)].copy()

rep_real_xls, rep_obj_xls, rep_c_xls = summarize_segment(d_rep_xls)
srv_real_xls, srv_obj_xls, srv_c_xls = summarize_segment(d_srv_xls)

total_real_xls = rep_real_xls + srv_real_xls
total_obj_xls  = rep_obj_xls + srv_obj_xls
total_c_xls    = safe_ratio(total_real_xls, total_obj_xls)

rep_proy_xls = proyectar_eom_runrate(rep_real_xls, dias_transc_xls, dias_total_mes_xls)
srv_proy_xls = proyectar_eom_runrate(srv_real_xls, dias_transc_xls, dias_total_mes_xls)
total_proy_xls = proyectar_eom_runrate(total_real_xls, dias_transc_xls, dias_total_mes_xls)

rep_by_suc_xls = micro_sucursal(apply_obj0_filter(d_rep_xls, show_obj0))
srv_by_suc_xls = micro_sucursal(apply_obj0_filter(d_srv_xls, show_obj0))
rep_by_ap_xls  = micro_aperturas(apply_obj0_filter(d_rep_xls, show_obj0))
srv_by_ap_xls  = micro_aperturas(apply_obj0_filter(d_srv_xls, show_obj0))

top_n_xls = int(st.session_state.get("top_n_rank", 10))
show_zero_rank_xls = bool(st.session_state.get("show_zero_rank", False))
rep_rank_xls = ranking_sucursal_apertura_micro(d_rep_xls, top_n=top_n_xls, show_zero=show_zero_rank_xls)
srv_rank_xls = ranking_sucursal_apertura_micro(d_srv_xls, top_n=top_n_xls, show_zero=show_zero_rank_xls)

driver_xls = principal_driver_gap(pd.concat([d_rep_xls, d_srv_xls], ignore_index=True))

resumen_xls = pd.DataFrame([
    {"Bloque":"Repuestos", "Real_Acum":rep_real_xls, "Obj_Acum":rep_obj_xls, "Cumpl_Acum":rep_c_xls, "Proy_EOM_RunRate":rep_proy_xls,
     "Dias_Transc": (np.nan if pd.isna(dias_transc_xls) else dias_transc_xls), "Dias_Mes": (np.nan if pd.isna(dias_total_mes_xls) else dias_total_mes_xls)},
    {"Bloque":"Servicios", "Real_Acum":srv_real_xls, "Obj_Acum":srv_obj_xls, "Cumpl_Acum":srv_c_xls, "Proy_EOM_RunRate":srv_proy_xls,
     "Dias_Transc": (np.nan if pd.isna(dias_transc_xls) else dias_transc_xls), "Dias_Mes": (np.nan if pd.isna(dias_total_mes_xls) else dias_total_mes_xls)},
    {"Bloque":"Total Postventa", "Real_Acum":total_real_xls, "Obj_Acum":total_obj_xls, "Cumpl_Acum":total_c_xls, "Proy_EOM_RunRate":total_proy_xls,
     "Dias_Transc": (np.nan if pd.isna(dias_transc_xls) else dias_transc_xls), "Dias_Mes": (np.nan if pd.isna(dias_total_mes_xls) else dias_total_mes_xls)},
])

meta_xls = pd.DataFrame([{
    "Sucursal": sucursal,
    "Mes_Seleccionado": build_month_label(mes_sel),
    "Mes_Key": mes_sel,
    "Semana_Corte_Dashboard": semana_corte,
    "Semana_Export_Excel": int(sem_export),
    "SemanaMes_Corte_Excel": semana_mes_corte_xls,
    "Dias_Habiles_Transcurridos_Excel": (None if pd.isna(dias_transc_xls) else int(dias_transc_xls)),
    "Dias_Habiles_Mes_Excel": (None if pd.isna(dias_total_mes_xls) else int(dias_total_mes_xls)),
    "Cap_Visual_On": cap_on,
    "Cap_Visual_Max": cap_val,
    "Incluir_Obj_0": show_obj0,
    "Hoja_Dias_Habiles": (DIAS_SHEET_FOUND if DIAS_SHEET_FOUND is not None else "NO_ENCONTRADA"),
    "Hoja_Resumen_del_Mes": (RESUMEN_SHEET_FOUND if RESUMEN_SHEET_FOUND is not None else "NO_ENCONTRADA"),
    "Ranking_TopN": top_n_xls,
    "Ranking_ShowZero": show_zero_rank_xls,
    "Principal_Desvio_KPI": (driver_xls["KPI"] if driver_xls else "—"),
    "Principal_Desvio_Cat": (driver_xls["Cat"] if driver_xls else "—"),
    "Principal_Desvio_Gap": (driver_xls["Gap"] if driver_xls else np.nan),
}])

hitos_export = filter_hitos_by_month(df_resumen_mes, mes_sel)

excel_bytes_prof = build_exec_excel_professional(
    meta_df=meta_xls,
    resumen_df=resumen_xls,
    pl_sucursal_rep=rep_by_suc_xls,
    pl_sucursal_srv=srv_by_suc_xls,
    pl_ap_rep=rep_by_ap_xls,
    pl_ap_srv=srv_by_ap_xls,
    ranking_rep=rep_rank_xls,
    ranking_srv=srv_rank_xls,
    hitos_df=hitos_export,
)

with hr:
    st.download_button(
        "⬇️ Resumen Ejecutivo (Excel)",
        data=excel_bytes_prof,
        file_name=f"Resumen_Ejecutivo_Tablero_Posventa_{mes_sel}_Sem{int(sem_export)}_{sucursal.replace(' ','_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ---------------------------
# TABS
# ---------------------------
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

    d_pl = df_cut[df_cut["Tipo_KPI"]=="$"].copy()

    d_rep = d_pl[d_pl["KPI"].str.upper()=="REPUESTOS"].copy()
    d_rep = d_rep[d_rep["Categoria_KPI"].isin(rep_sel)].copy()

    d_srv = d_pl[d_pl["KPI"].str.upper()=="SERVICIOS"].copy()
    d_srv = d_srv[d_srv["Categoria_KPI"].isin(srv_sel)].copy()

    rep_real, rep_obj, rep_c = summarize_segment(d_rep)
    srv_real, srv_obj, srv_c = summarize_segment(d_srv)

    total_real = rep_real + srv_real
    total_obj  = rep_obj + srv_obj
    total_c    = safe_ratio(total_real, total_obj)

    rep_proy = proyectar_eom_runrate(rep_real, dias_transc, dias_total_mes)
    srv_proy = proyectar_eom_runrate(srv_real, dias_transc, dias_total_mes)
    total_proy = proyectar_eom_runrate(total_real, dias_transc, dias_total_mes)

    driver = principal_driver_gap(pd.concat([d_rep, d_srv], ignore_index=True))
    driver_txt = f"Principal desvío: **{driver['KPI']} / {driver['Cat']}** (Gap {money_str(driver['Gap'])})" if driver else "Principal desvío: —"

    st.info(
        f"**Resumen Ejecutivo:** "
        f"Repuestos {pct_str(rep_c)} | "
        f"Servicios {pct_str(srv_c)} | "
        f"Total Postventa {pct_str(total_c)} | "
        f"{driver_txt}"
    )

    rep_month = df_month[(df_month["Tipo_KPI"]=="$") & (df_month["KPI"].str.upper()=="REPUESTOS")].copy()
    rep_month = rep_month[rep_month["Categoria_KPI"].isin(rep_sel)].copy()
    srv_month = df_month[(df_month["Tipo_KPI"]=="$") & (df_month["KPI"].str.upper()=="SERVICIOS")].copy()
    srv_month = srv_month[srv_month["Categoria_KPI"].isin(srv_sel)].copy()
    total_month = pd.concat([rep_month, srv_month], ignore_index=True)

    def macro_block(titulo, real, obj, proy, df_scope_month_for_spark):
        c = safe_ratio(real, obj)
        sub = (
            f"Real {money_str(real)} | Obj {money_str(obj)} | "
            f"Proy EOM (run-rate): {money_str(proy)} | "
            f"Días: {('—' if pd.isna(dias_transc) else int(dias_transc))}/{('—' if pd.isna(dias_total_mes) else int(dias_total_mes))}"
        )
        st.markdown(card_html_base(titulo, money_str(real), sub), unsafe_allow_html=True)
        st.markdown(footer_kpi_only_html("Cumpl. Acum.", pct_str(c)), unsafe_allow_html=True)
        spark_evolucion(df_scope_month_for_spark)

    c1, c_mid, c2 = st.columns([1.0, 1.05, 1.0])

    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        macro_block("Repuestos — Real (Acum.)", rep_real, rep_obj, rep_proy, rep_month)

    with c_mid:
        st.markdown("### 🧩 TOTAL POSTVENTA (P&L)")
        macro_block("Total Postventa — Real (Acum.)", total_real, total_obj, total_proy, total_month)

    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        macro_block("Servicios — Real (Acum.)", srv_real, srv_obj, srv_proy, srv_month)

    st.markdown("---")
    st.markdown("### Cumplimiento por Sucursal — (acumulado)")

    a, b = st.columns(2)
    with a:
        st.markdown("**Repuestos — por sucursal**")
        g = micro_sucursal(apply_obj0_filter(d_rep, show_obj0))
        if g.empty:
            st.info("Sin datos por sucursal.")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with b:
        st.markdown("**Servicios — por sucursal**")
        g = micro_sucursal(apply_obj0_filter(d_srv, show_obj0))
        if g.empty:
            st.info("Sin datos por sucursal.")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="Sucursal", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### Aperturas — micro (cumplimiento acumulado)")

    l, r = st.columns(2)
    with l:
        st.markdown("**Repuestos — por apertura**")
        g = micro_aperturas(apply_obj0_filter(d_rep, show_obj0))
        if g.empty:
            st.info("Sin datos (revisar aperturas).")
        else:
            fig = px.bar(g, x="Cumpl_plot", y="Categoria_KPI", orientation="h", text="label")
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Cumplimiento (visual)")
            fig.update_traces(textposition="inside")
            st.plotly_chart(fig, use_container_width=True)

    with r:
        st.markdown("**Servicios — por apertura**")
        g = micro_aperturas(apply_obj0_filter(d_srv, show_obj0))
        if g.empty:
            st.info("Sin datos (revisar aperturas).")
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
        st.session_state["top_n_rank"] = int(top_n)
    with cB:
        rep_micro_choice = st.selectbox("Repuestos (micro)", ["Todas las aperturas"] + rep_open, index=0)
    with cC:
        srv_micro_choice = st.selectbox("Servicios (micro)", ["Todas las aperturas"] + srv_open, index=0)
    with cD:
        show_zero_rank = st.checkbox("Mostrar 0% (Obj=0 y real=0)", value=False)
        st.session_state["show_zero_rank"] = bool(show_zero_rank)

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
    st.markdown("## 📌 KPIs (resto) — Macro → Micro")
    st.markdown("---")

    resto = df_cut[~df_cut["KPI"].str.upper().isin(["REPUESTOS","SERVICIOS"])].copy()
    resto = apply_obj0_filter(resto, show_obj0)

    kpis_resto = sorted(resto["KPI"].unique().tolist())
    if not kpis_resto:
        st.info("No hay KPIs (resto) con Obj>0 en este corte.")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto)

        x = resto[resto["KPI"] == kpi_sel].copy()
        tipos = sorted(x["Tipo_KPI"].unique().tolist())

        for t in tipos:
            xt = x[x["Tipo_KPI"] == t].copy()
            real = xt["Real_val"].sum()
            obj  = xt["Obj_val"].sum()
            c    = safe_ratio(real, obj)

            st.markdown(
                card_html_base(
                    f"{kpi_sel} ({t}) — Cumplimiento (Acum.)",
                    (money_str(real) if t=="$" else qty_str(real)),
                    f"Real {(money_str(real) if t=='$' else qty_str(real))} | Obj {(money_str(obj) if t=='$' else qty_str(obj))}"
                ),
                unsafe_allow_html=True
            )
            st.markdown(footer_kpi_only_html("Cumpl. Acum.", pct_str(c)), unsafe_allow_html=True)

            g = xt.groupby("Sucursal", as_index=False).agg(
                Real=("Real_val","sum"),
                Obj=("Obj_val","sum")
            )
            g["Cumpl"] = g.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
            g = g[~g["Cumpl"].isna()].copy().sort_values("Cumpl", ascending=False)
            g = apply_cap_visual(g, cap_on, cap_val)

            g["label"] = g.apply(
                lambda r: f"{pct_str(r['Cumpl'])} | {(money_str(r['Real']) if t=='$' else qty_str(r['Real']))}/{(money_str(r['Obj']) if t=='$' else qty_str(r['Obj']))}",
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

        st.markdown("---")
        with st.expander("🔎 Auditoría (KPIs resto)", expanded=False):
            detail = x.copy().sort_values(["Semana_Num","Sucursal","KPI","Categoria_KPI"])
            st.dataframe(detail, use_container_width=True, hide_index=True)

# ============================================================
# TAB 3 — Gestión
# ============================================================
with tab3:
    st.markdown("## 🧪 Gestión (desvíos)")
    st.markdown("---")

    suc_g = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sorted(df["Sucursal"].dropna().unique().tolist()), index=0)

    d = df[df["Mes"] == mes_sel].copy()
    if suc_g != "TODAS (Consolidado)":
        d = d[d["Sucursal"] == suc_g].copy()

    d = d[d["Semana_Num"] <= semana_corte].copy()
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
        gg["label"] = gg.apply(
            lambda r: f"Gap {(money_str(r['Gap']) if r['Tipo_KPI']=='$' else qty_str(r['Gap']))} | {pct_str(r['Cumpl'])}",
            axis=1
        )

        fig = px.bar(gg, x="Gap", y="key", orientation="h", text="label")
        fig.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="Gap (Obj - Real)")
        fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("🔎 Auditoría (Gestión)", expanded=False):
            st.dataframe(g, use_container_width=True, hide_index=True)

# ============================================================
# TAB 4 — Hitos del mes
# ============================================================
with tab4:
    st.markdown("## 🗓️ Hitos del mes — Resumen del mes")
    st.caption(f"Mes seleccionado: **{build_month_label(mes_sel)}**")
    st.markdown("---")

    if RESUMEN_SHEET_FOUND is None:
        st.warning("No pude encontrar la hoja 'Resumen del mes' por nombre.")
        with st.expander("Ver pestañas detectadas", expanded=True):
            st.write(SHEET_NAMES)
        st.stop()

    st.success(f"✅ Hoja encontrada: **{RESUMEN_SHEET_FOUND}**")

    hitos_mes = filter_hitos_by_month(df_resumen_mes, mes_sel)

    if hitos_mes is None or hitos_mes.empty:
        st.info("La hoja existe, pero no tiene contenido para este mes.")
    else:
        st.dataframe(hitos_mes, use_container_width=True, hide_index=True)
