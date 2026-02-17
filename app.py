import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

st.set_page_config(page_title="Tablero Posventa", layout="wide")

# ==========================
# CONFIG DRIVE
# ==========================
DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# ==========================
# ESTILO (mejorado)
# ==========================
st.markdown("""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
section[data-testid="stSidebar"] .block-container {padding-top: 1.1rem;}

h1 {margin-bottom: 0.2rem;}
.small-muted {opacity: 0.7; font-size: 0.9rem;}

.card {
  background:#fff; border:1px solid rgba(0,0,0,0.08);
  border-radius:16px; padding:16px 18px;
  box-shadow: 0 10px 28px rgba(0,0,0,0.04);
}
.card-title {font-size:0.85rem; opacity:0.75; margin-bottom:6px;}
.card-value {font-size:1.8rem; font-weight:800; line-height:1.05;}
.card-sub {font-size:0.95rem; opacity:0.78; margin-top:8px;}
.hr {height:1px; background:rgba(0,0,0,0.08); margin: 14px 0;}

.badge {
  display:inline-block; padding: 6px 10px; border-radius: 999px;
  font-size: 0.85rem; font-weight: 700; color: white;
}
.badge-red {background:#d64545;}
.badge-yellow {background:#d1a100;}
.badge-green {background:#2c9f6b;}
.badge-gray {background:#6c757d;}

.kpi-grid {display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:12px;}
@media (max-width: 1200px){ .kpi-grid {grid-template-columns: repeat(2, minmax(0, 1fr));} }
@media (max-width: 800px){ .kpi-grid {grid-template-columns: repeat(1, minmax(0, 1fr));} }

</style>
""", unsafe_allow_html=True)

# ==========================
# Helpers
# ==========================
def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = s.replace(" ", "").replace("-", "").replace(".", "").replace("/", "")
    s = s.replace("__", "_")
    return s

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidates:
        key = _norm(cand)
        if key in norm_map:
            return norm_map[key]
    for cand in candidates:
        key = _norm(cand)
        for k, real in norm_map.items():
            if key in k:
                return real
    return None

def parse_semana_num(x):
    if pd.isna(x):
        return None
    m = re.search(r"(\d+)", str(x))
    return int(m.group(1)) if m else None

def money_fmt(x):
    if x is None or pd.isna(x): return "—"
    try:
        return f"${x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def num_fmt(x):
    if x is None or pd.isna(x): return "—"
    try:
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return str(x)

def pct_fmt_ratio(x):
    if x is None or pd.isna(x): return "—"
    return f"{x*100:.1f}%"

def badge_html(estado):
    if estado == "Rojo":
        return '<span class="badge badge-red">ROJO</span>'
    if estado == "Amarillo":
        return '<span class="badge badge-yellow">AMARILLO</span>'
    if estado == "Verde":
        return '<span class="badge badge-green">VERDE</span>'
    return '<span class="badge badge-gray">—</span>'

def estado_por_umbral(cumpl, umbral_amar, umbral_verde):
    if cumpl is None or pd.isna(cumpl):
        return "—"
    if cumpl >= umbral_verde:
        return "Verde"
    if cumpl >= umbral_amar:
        return "Amarillo"
    return "Rojo"

def estado_global(serie_estados):
    s = set([x for x in serie_estados.dropna().tolist()])
    if "Rojo" in s: return "Rojo"
    if "Amarillo" in s: return "Amarillo"
    if "Verde" in s: return "Verde"
    return "—"

# ✅ PARSER NUMÉRICO ARGENTINA
def to_number_ar(s):
    if pd.isna(s):
        return None
    if isinstance(s, (int, float)):
        return float(s)

    txt = str(s).strip()
    if txt == "":
        return None

    txt = re.sub(r"[^0-9,\.\-]", "", txt)

    if txt.count(",") >= 1 and txt.count(".") >= 1:
        if re.search(r",\d{1,2}$", txt):
            txt = txt.replace(".", "")
            txt = txt.replace(",", ".")
        else:
            txt = txt.replace(",", "")
    elif txt.count(",") >= 1 and txt.count(".") == 0:
        txt = txt.replace(",", ".")

    try:
        return float(txt)
    except Exception:
        return None

def coerce_numeric_ar(series: pd.Series) -> pd.Series:
    return series.apply(to_number_ar).astype("float64")

def norm_tipo_kpi(x):
    if pd.isna(x):
        return None
    t = str(x).strip().upper()
    if "$" in t:
        return "$"
    if "Q" in t:
        return "Q"
    return t

def safe_ratio(n, d):
    if d is None or pd.isna(d) or d == 0:
        return None
    if n is None or pd.isna(n):
        return None
    return n / d

# ==========================
# Carga desde Google Sheets
# ==========================
@st.cache_data(show_spinner=True, ttl=300)
def load_from_drive():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True, fuzzy=True)
    df = pd.read_excel(EXCEL_LOCAL, sheet_name=0)

    try:
        dim_kpi = pd.read_excel(EXCEL_LOCAL, sheet_name="DIM_KPI")
    except Exception:
        dim_kpi = pd.DataFrame(columns=["KPI", "Umbral_Amarillo", "Umbral_Verde"])

    df.columns = [str(c).strip() for c in df.columns]
    dim_kpi.columns = [str(c).strip() for c in dim_kpi.columns]
    return df, dim_kpi

def resolve_schema(df: pd.DataFrame) -> dict:
    col = {}
    col["fecha"] = find_col(df, ["Fecha"])
    col["semana"] = find_col(df, ["Semana"])
    col["sucursal"] = find_col(df, ["Sucursal"])
    col["kpi"] = find_col(df, ["KPI"])
    col["categoria_kpi"] = find_col(df, ["Categoria_KPI", "Categoría_KPI", "Categoria KPI", "Categoría KPI"])
    col["tipo_kpi"] = find_col(df, ["Tipo_KPI"])

    col["real_$"] = find_col(df, ["Real_$", "Real $", "Real$"])
    col["obj_$"]  = find_col(df, ["Objetivo_$", "Objetivo $", "Objetivo$"])
    col["real_q"] = find_col(df, ["Real_Q", "Real Q"])
    col["obj_q"]  = find_col(df, ["Objetivo_Q", "Objetivo Q"])

    col["costo_$"]  = find_col(df, ["Costo_$", "Costo $", "Costo$"])
    col["margen_$"] = find_col(df, ["Margen_$", "Margen $", "Margen$"])

    # categoria_kpi es opcional, todo lo demás es clave
    required = ["semana","sucursal","kpi","tipo_kpi","real_$","obj_$","real_q","obj_q","costo_$","margen_$"]
    missing = [k for k in required if col.get(k) is None]

    if missing:
        return {"ok": False, "missing": missing, "found_cols": list(df.columns), "col": col}
    return {"ok": True, "col": col}

def normalize_dim_kpi(dim_kpi: pd.DataFrame) -> pd.DataFrame:
    if dim_kpi is None or len(dim_kpi) == 0:
        return pd.DataFrame(columns=["KPI", "Umbral_Amarillo", "Umbral_Verde"])

    kpi_col = find_col(dim_kpi, ["KPI"])
    ua_col  = find_col(dim_kpi, ["Umbral_Amarillo", "Amarillo"])
    uv_col  = find_col(dim_kpi, ["Umbral_Verde", "Verde"])

    out = pd.DataFrame()
    out["KPI"] = dim_kpi[kpi_col] if kpi_col else None
    out["Umbral_Amarillo"] = pd.to_numeric(dim_kpi[ua_col], errors="coerce") if ua_col else 0.90
    out["Umbral_Verde"] = pd.to_numeric(dim_kpi[uv_col], errors="coerce") if uv_col else 1.00
    out["Umbral_Amarillo"] = out["Umbral_Amarillo"].fillna(0.90)
    out["Umbral_Verde"] = out["Umbral_Verde"].fillna(1.00)
    out = out.dropna(subset=["KPI"]).copy()
    return out

# ==========================
# Build semanal/acumulado con "Apertura"
# ==========================
def build_kpi_week(df_raw: pd.DataFrame, schema: dict) -> pd.DataFrame:
    c = schema["col"]
    df = df_raw.copy()

    if c["fecha"]:
        df[c["fecha"]] = pd.to_datetime(df[c["fecha"]], errors="coerce")

    df["Semana_Num"] = df[c["semana"]].apply(parse_semana_num)
    df[c["tipo_kpi"]] = df[c["tipo_kpi"]].apply(norm_tipo_kpi)

    for k in ["real_$","obj_$","real_q","obj_q","costo_$","margen_$"]:
        df[c[k]] = coerce_numeric_ar(df[c[k]])

    df["Real_val"] = df.apply(lambda r: r[c["real_$"]] if r[c["tipo_kpi"]] == "$" else r[c["real_q"]], axis=1)
    df["Obj_val"]  = df.apply(lambda r: r[c["obj_$"]]  if r[c["tipo_kpi"]] == "$" else r[c["obj_q"]], axis=1)

    # Apertura solo para Repuestos/Servicios (macro→micro)
    if c["categoria_kpi"]:
        df["Apertura"] = df.apply(
            lambda r: r[c["categoria_kpi"]] if str(r[c["kpi"]]).strip().lower() in ["repuestos","servicios"] else "TOTAL",
            axis=1
        )
    else:
        df["Apertura"] = "TOTAL"

    agg = df.groupby(
        ["Semana_Num", c["sucursal"], c["kpi"], "Apertura", c["tipo_kpi"]],
        as_index=False
    ).agg(
        Real_Sem=("Real_val", "sum"),
        Obj_Sem=("Obj_val", "sum"),
        Costo_Sem=(c["costo_$"], "sum"),
        Margen_Sem=(c["margen_$"], "sum"),
    ).rename(columns={
        c["sucursal"]: "Sucursal",
        c["kpi"]: "KPI",
        c["tipo_kpi"]: "Tipo_KPI",
    })

    agg["Cumpl_Sem"] = agg.apply(lambda r: safe_ratio(r["Real_Sem"], r["Obj_Sem"]), axis=1)
    agg = agg.sort_values(["Sucursal","KPI","Apertura","Semana_Num"]).copy()

    agg["Real_Acum"] = agg.groupby(["Sucursal","KPI","Apertura"])["Real_Sem"].cumsum()
    agg["Obj_Acum"]  = agg.groupby(["Sucursal","KPI","Apertura"])["Obj_Sem"].cumsum()
    agg["Cumpl_Acum"] = agg.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)

    agg["Margen_Acum"] = agg.groupby(["Sucursal","KPI","Apertura"])["Margen_Sem"].cumsum()
    agg["MargenPct_Acum"] = agg.apply(lambda r: safe_ratio(r["Margen_Acum"], r["Real_Acum"]), axis=1)

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
    # consolidado por KPI + Apertura + Tipo (respetando filtros posteriores)
    df_valid = df_last_suc[
        ((pd.notna(df_last_suc["Obj_Sem"])) & (df_last_suc["Obj_Sem"] > 0)) |
        ((pd.notna(df_last_suc["Obj_Acum"])) & (df_last_suc["Obj_Acum"] > 0))
    ].copy()

    cons = df_valid.groupby(["KPI","Apertura","Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
        Real_Sem=("Real_Sem","sum"),
        Obj_Sem=("Obj_Sem","sum"),
        Margen_Sem=("Margen_Sem","sum"),
    )
    cons["Cumpl_Acum"] = cons.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)
    cons["Cumpl_Sem"]  = cons.apply(lambda r: safe_ratio(r["Real_Sem"], r["Obj_Sem"]), axis=1)
    cons["MargenPct_Acum"] = cons.apply(lambda r: safe_ratio(r["Margen_Acum"], r["Real_Acum"]), axis=1)
    return cons

# ==========================
# UI helpers: tarjetas macro
# ==========================
def card_html(title, value, sub=None):
    sub_html = f'<div class="card-sub">{sub}</div>' if sub else ""
    return f"""
    <div class="card">
      <div class="card-title">{title}</div>
      <div class="card-value">{value}</div>
      {sub_html}
    </div>
    """

def segment_metrics(df_seg):
    # df_seg ya filtrado por aperturas elegidas y con Tipo_KPI "$"
    real = float(df_seg["Real_Acum"].sum()) if len(df_seg) else 0.0
    obj  = float(df_seg["Obj_Acum"].sum()) if len(df_seg) else 0.0
    cump = safe_ratio(real, obj)
    margen = float(df_seg["Margen_Acum"].sum()) if ("Margen_Acum" in df_seg.columns and len(df_seg)) else 0.0
    margen_pct = safe_ratio(margen, real)
    return real, obj, cump, margen, margen_pct

# ==========================
# CARGA
# ==========================
df_raw, dim_kpi_raw = load_from_drive()
schema = resolve_schema(df_raw)
if not schema["ok"]:
    st.error("⚠️ Faltan columnas necesarias (cambió el archivo).")
    st.write("Faltan:", schema["missing"])
    st.code("\n".join(schema["found_cols"]))
    st.stop()

dim_kpi = normalize_dim_kpi(dim_kpi_raw)
df_week = build_kpi_week(df_raw, schema)

# ==========================
# SIDEBAR
# ==========================
st.sidebar.title("Filtros obligatorios")

semanas = sorted(df_week["Semana_Num"].dropna().unique())
sucursales = sorted(df_week["Sucursal"].dropna().unique())

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=0 if len(semanas) else 0)
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

df_last_suc = df_week[df_week["Semana_Num"] == semana_corte].copy()
df_last_suc = aplicar_reglas(df_last_suc, dim_kpi)

if sucursal != "TODAS (Consolidado)":
    df_last = df_last_suc[df_last_suc["Sucursal"] == sucursal].copy()
else:
    df_last = consolidar_todas(df_last_suc)
    df_last = aplicar_reglas(df_last, dim_kpi)

# Filtros por aperturas (solo para P&L)
rep_aperturas = sorted(
    df_last[(df_last["KPI"] == "Repuestos") & (df_last["Apertura"] != "TOTAL")]["Apertura"].dropna().unique()
)
srv_aperturas = sorted(
    df_last[(df_last["KPI"] == "Servicios") & (df_last["Apertura"] != "TOTAL")]["Apertura"].dropna().unique()
)

st.sidebar.markdown("---")
st.sidebar.subheader("Incluir variables (P&L)")

rep_sel = st.sidebar.multiselect(
    "Repuestos: aperturas incluidas",
    options=rep_aperturas,
    default=rep_aperturas
)
srv_sel = st.sidebar.multiselect(
    "Servicios: aperturas incluidas",
    options=srv_aperturas,
    default=srv_aperturas
)

st.sidebar.markdown("---")
rank_metric = st.sidebar.selectbox(
    "Ranking por sucursal (macro)",
    ["Cumplimiento %", "Real (monto)"],
    index=0
)

# ==========================
# HEADER
# ==========================
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🏆 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧩 Gestión (desvíos)"])

# ==========================
# TAB 1 — P&L
# ==========================
with tab1:
    # Macro: Repuestos y Servicios (solo $)
    rep = df_last[(df_last["KPI"] == "Repuestos") & (df_last["Tipo_KPI"] == "$")].copy()
    srv = df_last[(df_last["KPI"] == "Servicios") & (df_last["Tipo_KPI"] == "$")].copy()

    if len(rep_aperturas):
        rep = rep[(rep["Apertura"].isin(rep_sel)) | (rep["Apertura"] == "TOTAL")]
    if len(srv_aperturas):
        srv = srv[(srv["Apertura"].isin(srv_sel)) | (srv["Apertura"] == "TOTAL")]

    # Para macro, usamos TOTAL del segmento: sumamos aperturas seleccionadas
    # (si tu archivo trae TOTAL ya armado, igual lo ignoramos y construimos el total real)
    rep_total = rep[rep["Apertura"].isin(rep_sel)].copy() if rep_sel else rep[rep["Apertura"] != "TOTAL"].copy()
    srv_total = srv[srv["Apertura"].isin(srv_sel)].copy() if srv_sel else srv[srv["Apertura"] != "TOTAL"].copy()

    rep_real, rep_obj, rep_cump, rep_margen, rep_margen_pct = segment_metrics(rep_total)
    srv_real, srv_obj, srv_cump, srv_margen, srv_margen_pct = segment_metrics(srv_total)

    # Estado macro por umbral (si existe DIM_KPI para Repuestos/Servicios)
    rep_estado = estado_por_umbral(rep_cump, 0.90, 1.00)
    srv_estado = estado_por_umbral(srv_cump, 0.90, 1.00)

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        st.markdown(
            "<div class='kpi-grid'>"
            + card_html("Cumplimiento (Acum.)", pct_fmt_ratio(rep_cump), f"Estado: {badge_html(rep_estado)}")
            + card_html("Real (Acum.)", money_fmt(rep_real), f"Objetivo: {money_fmt(rep_obj)}")
            + card_html("Margen % (Acum.)", pct_fmt_ratio(rep_margen_pct), f"Margen: {money_fmt(rep_margen)}")
            + "</div>",
            unsafe_allow_html=True
        )

        st.markdown("<div class='hr'></div>", unsafe_allow_html=True)
        st.markdown("#### Aperturas (Repuestos) — micro")
        rep_micro = rep_total.copy()
        if len(rep_micro):
            rep_micro["Cumpl_Acum_plot"] = rep_micro["Cumpl_Acum"]
            rep_micro = rep_micro.dropna(subset=["Cumpl_Acum_plot"]).copy()
            if len(rep_micro):
                fig = px.bar(
                    rep_micro.sort_values("Cumpl_Acum_plot"),
                    x="Cumpl_Acum_plot", y="Apertura", orientation="h",
                    text=rep_micro["Cumpl_Acum_plot"].apply(lambda x: f"{x*100:.1f}%")
                )
                fig.update_layout(xaxis_tickformat=".0%")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Repuestos: sin objetivos válidos para aperturas seleccionadas.")
        else:
            st.info("Repuestos: no hay datos con aperturas seleccionadas.")

    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        st.markdown(
            "<div class='kpi-grid'>"
            + card_html("Cumplimiento (Acum.)", pct_fmt_ratio(srv_cump), f"Estado: {badge_html(srv_estado)}")
            + card_html("Real (Acum.)", money_fmt(srv_real), f"Objetivo: {money_fmt(srv_obj)}")
            + card_html("Margen % (Acum.)", pct_fmt_ratio(srv_margen_pct), f"Margen: {money_fmt(srv_margen)}")
            + "</div>",
            unsafe_allow_html=True
        )

        st.markdown("<div class='hr'></div>", unsafe_allow_html=True)
        st.markdown("#### Aperturas (Servicios) — micro")
        srv_micro = srv_total.copy()
        if len(srv_micro):
            srv_micro["Cumpl_Acum_plot"] = srv_micro["Cumpl_Acum"]
            srv_micro = srv_micro.dropna(subset=["Cumpl_Acum_plot"]).copy()
            if len(srv_micro):
                fig2 = px.bar(
                    srv_micro.sort_values("Cumpl_Acum_plot"),
                    x="Cumpl_Acum_plot", y="Apertura", orientation="h",
                    text=srv_micro["Cumpl_Acum_plot"].apply(lambda x: f"{x*100:.1f}%")
                )
                fig2.update_layout(xaxis_tickformat=".0%")
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Servicios: sin objetivos válidos para aperturas seleccionadas.")
        else:
            st.info("Servicios: no hay datos con aperturas seleccionadas.")

    st.markdown("<div class='hr'></div>", unsafe_allow_html=True)
    st.markdown("### 🏁 Ranking por sucursal (Macro)")

    # Ranking por sucursal siempre se calcula con df_last_suc (no consolidado)
    df_rank_base = df_last_suc.copy()
    # Repuestos
    rep_rank = df_rank_base[(df_rank_base["KPI"] == "Repuestos") & (df_rank_base["Tipo_KPI"] == "$") & (df_rank_base["Apertura"].isin(rep_sel))].copy()
    srv_rank = df_rank_base[(df_rank_base["KPI"] == "Servicios") & (df_rank_base["Tipo_KPI"] == "$") & (df_rank_base["Apertura"].isin(srv_sel))].copy()

    rep_rank_suc = rep_rank.groupby("Sucursal", as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
    )
    rep_rank_suc["Cumpl_Acum"] = rep_rank_suc.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)

    srv_rank_suc = srv_rank.groupby("Sucursal", as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
    )
    srv_rank_suc["Cumpl_Acum"] = srv_rank_suc.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)

    rc1, rc2 = st.columns(2, gap="large")

    with rc1:
        st.markdown("#### Repuestos — por sucursal")
        if len(rep_rank_suc):
            y = "Cumpl_Acum" if rank_metric == "Cumplimiento %" else "Real_Acum"
            fig3 = px.bar(
                rep_rank_suc.sort_values(y, ascending=True),
                x=y, y="Sucursal", orientation="h",
                text=rep_rank_suc[y].apply(lambda v: f"{v*100:.1f}%" if y=="Cumpl_Acum" and pd.notna(v) else money_fmt(v))
            )
            if y == "Cumpl_Acum":
                fig3.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No hay datos de ranking de Repuestos con las aperturas seleccionadas.")

    with rc2:
        st.markdown("#### Servicios — por sucursal")
        if len(srv_rank_suc):
            y = "Cumpl_Acum" if rank_metric == "Cumplimiento %" else "Real_Acum"
            fig4 = px.bar(
                srv_rank_suc.sort_values(y, ascending=True),
                x=y, y="Sucursal", orientation="h",
                text=srv_rank_suc[y].apply(lambda v: f"{v*100:.1f}%" if y=="Cumpl_Acum" and pd.notna(v) else money_fmt(v))
            )
            if y == "Cumpl_Acum":
                fig4.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("No hay datos de ranking de Servicios con las aperturas seleccionadas.")

# ==========================
# TAB 2 — resto de KPIs
# ==========================
with tab2:
    st.markdown("### 📌 KPIs (resto) — Macro → Micro")
    st.markdown("<span class='small-muted'>Acá van los KPIs que no son Repuestos/Servicios.</span>", unsafe_allow_html=True)

    resto = df_last[~df_last["KPI"].isin(["Repuestos","Servicios"])].copy()

    if sucursal == "TODAS (Consolidado)":
        base_kpis = sorted(resto["KPI"].dropna().unique())
    else:
        base_kpis = sorted(resto["KPI"].dropna().unique())

    if not base_kpis:
        st.info("No hay KPIs adicionales en este corte.")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", base_kpis)
        df_k = resto[resto["KPI"] == kpi_sel].copy()

        # Tarjeta macro KPI
        tipo = df_k["Tipo_KPI"].iloc[0] if len(df_k) else "—"
        real = float(df_k["Real_Acum"].sum()) if len(df_k) else 0.0
        obj  = float(df_k["Obj_Acum"].sum()) if len(df_k) else 0.0
        cump = safe_ratio(real, obj)
        estado = estado_por_umbral(cump, 0.90, 1.00)

        st.markdown(
            "<div class='kpi-grid'>"
            + card_html("KPI", f"{kpi_sel} ({tipo})", f"Estado: {badge_html(estado)}")
            + card_html("Cumplimiento (Acum.)", pct_fmt_ratio(cump), f"Real: {money_fmt(real) if tipo=='$' else num_fmt(real)}")
            + card_html("Objetivo (Acum.)", money_fmt(obj) if tipo=="$" else num_fmt(obj), None)
            + "</div>",
            unsafe_allow_html=True
        )

        st.markdown("<div class='hr'></div>", unsafe_allow_html=True)
        st.markdown("#### Ranking por sucursal — este KPI")

        df_rank_kpi = df_last_suc[df_last_suc["KPI"] == kpi_sel].copy()
        if len(df_rank_kpi):
            rk = df_rank_kpi.groupby(["Sucursal","Tipo_KPI"], as_index=False).agg(
                Real_Acum=("Real_Acum","sum"),
                Obj_Acum=("Obj_Acum","sum"),
            )
            rk["Cumpl_Acum"] = rk.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)
            y = "Cumpl_Acum" if rank_metric == "Cumplimiento %" else "Real_Acum"
            fig = px.bar(
                rk.sort_values(y, ascending=True),
                x=y, y="Sucursal", orientation="h",
                text=rk[y].apply(lambda v: f"{v*100:.1f}%" if y=="Cumpl_Acum" and pd.notna(v) else (money_fmt(v) if tipo=="$" else num_fmt(v)))
            )
            if y == "Cumpl_Acum":
                fig.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay datos por sucursal para este KPI en la semana seleccionada.")

# ==========================
# TAB 3 — Gestión
# ==========================
with tab3:
    st.markdown("### 🧩 Gestión (desvíos acumulados)")
    g = df_last.copy()
    g["Gap"] = g["Obj_Acum"] - g["Real_Acum"]
    order_map = {"Rojo": 0, "Amarillo": 1, "Verde": 2, "—": 9}
    g["OrdenEstado"] = g["Estado_Acum"].map(order_map).fillna(9)
    g = g.sort_values(["OrdenEstado","Gap"], ascending=[True, False])

    def fmt_val(tipo, v):
        if pd.isna(v): return "—"
        return money_fmt(v) if tipo == "$" else num_fmt(v)

    g["Cumpl_Acum_fmt"] = g["Cumpl_Acum"].apply(pct_fmt_ratio)
    g["Real_Acum_fmt"] = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Real_Acum"]), axis=1)
    g["Obj_Acum_fmt"]  = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Obj_Acum"]), axis=1)
    g["Gap_fmt"]       = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Gap"]), axis=1)

    st.dataframe(
        g[["KPI","Apertura","Tipo_KPI","Estado_Acum","Cumpl_Acum_fmt","Real_Acum_fmt","Obj_Acum_fmt","Gap_fmt"]],
        use_container_width=True,
        hide_index=True
    )
