import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

# ==========================
# CONFIG PÁGINA
# ==========================
st.set_page_config(page_title="Tablero Posventa", layout="wide")

# ==========================
# CONFIG DRIVE (TU ARCHIVO)
# ==========================
DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# ==========================
# ESTILO
# ==========================
st.markdown(r"""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
section[data-testid="stSidebar"] .block-container {padding-top: 1.1rem;}
.small-muted {opacity: 0.7; font-size: 0.9rem;}
.hr {height:1px; background:rgba(0,0,0,0.10); margin: 16px 0;}

.badge {
  display:inline-block; padding: 6px 10px; border-radius: 999px;
  font-size: 0.85rem; font-weight: 800; color: white;
}
.badge-red {background:#d64545;}
.badge-yellow {background:#d1a100;}
.badge-green {background:#2c9f6b;}
.badge-gray {background:#6c757d;}

.metric-wrap {
  border: 1px solid rgba(0,0,0,0.08);
  border-radius: 16px;
  padding: 14px 16px;
  box-shadow: 0 10px 28px rgba(0,0,0,0.04);
  background: white;
}
.metric-sub {opacity: 0.72; font-size: 0.9rem; margin-top: 6px;}

.exec-card{
  border: 1px solid rgba(0,0,0,0.08);
  border-radius: 18px;
  padding: 16px 16px;
  box-shadow: 0 12px 30px rgba(0,0,0,0.05);
  background: linear-gradient(180deg, rgba(248,250,255,1) 0%, rgba(255,255,255,1) 100%);
}
.exec-title{font-size:0.95rem; opacity:0.75; font-weight:800; letter-spacing:0.3px;}
.exec-main{font-size:1.65rem; font-weight:900; margin-top:6px;}
.exec-sub{opacity:0.78; font-size:0.92rem; margin-top:6px; line-height:1.25;}

/* =========================================================
   ✅ Multiselect "chips" (tags) -> VERDE SUAVE
   ========================================================= */
section[data-testid="stSidebar"] div[data-baseweb="tag"],
section[data-testid="stSidebar"] span[data-baseweb="tag"],
section[data-testid="stSidebar"] div[data-baseweb="tag"] > span,
section[data-testid="stSidebar"] div[data-baseweb="tag"] > div {
  background-color: #E8F5E9 !important;
  border: 1px solid #B7E1C1 !important;
  box-shadow: none !important;
}
section[data-testid="stSidebar"] div[data-baseweb="tag"] * ,
section[data-testid="stSidebar"] span[data-baseweb="tag"] * {
  color: #1B5E20 !important;
  font-weight: 700 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="tag"] svg,
section[data-testid="stSidebar"] span[data-baseweb="tag"] svg {
  color: #1B5E20 !important;
  fill: #1B5E20 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="tag"]:hover,
section[data-testid="stSidebar"] span[data-baseweb="tag"]:hover {
  background-color: #DFF2E3 !important;
  border-color: #9ED3AD !important;
}
section[data-testid="stSidebar"] div[role="listbox"] + div div[role="button"]{
  background-color: #E8F5E9 !important;
  border: 1px solid #B7E1C1 !important;
  color: #1B5E20 !important;
}
</style>
""", unsafe_allow_html=True)

# ==========================
# HELPERS
# ==========================
def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = s.replace(" ", "").replace("-", "").replace(".", "").replace("/", "")
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

def money_fmt_signed(x):
    if x is None or pd.isna(x): return "—"
    sign = "-" if x < 0 else ""
    return sign + money_fmt(abs(x))

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

def estado_por_umbral(cumpl, umbral_amar=0.90, umbral_verde=1.00):
    if cumpl is None or pd.isna(cumpl):
        return "—"
    if cumpl >= umbral_verde:
        return "Verde"
    if cumpl >= umbral_amar:
        return "Amarillo"
    return "Rojo"

# ✅ PARSER NUMÉRICO (ARG)
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

def metric_box(title, value, sub_html=""):
    st.markdown(f"""
    <div class="metric-wrap">
      <div class="small-muted">{title}</div>
      <div style="font-size: 1.7rem; font-weight: 900; margin-top: 2px;">{value}</div>
      <div class="metric-sub">{sub_html}</div>
    </div>
    """, unsafe_allow_html=True)

def exec_card(title, cumpl, real, obj, gap):
    estado = estado_por_umbral(cumpl)
    st.markdown(f"""
    <div class="exec-card">
      <div class="exec-title">{title}</div>
      <div class="exec-main">{pct_fmt_ratio(cumpl)} {badge_html(estado)}</div>
      <div class="exec-sub">
        <b>Real:</b> {money_fmt(real)} &nbsp;·&nbsp;
        <b>Obj:</b> {money_fmt(obj)}<br/>
        <b>Gap (Obj-Real):</b> {money_fmt_signed(gap)}
      </div>
    </div>
    """, unsafe_allow_html=True)

# ==========================
# CARGA DESDE GOOGLE SHEETS (XLSX)
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
    df_valid = df_last_suc[
        ((pd.notna(df_last_suc["Obj_Acum"])) & (df_last_suc["Obj_Acum"] > 0))
    ].copy()

    cons = df_valid.groupby(["KPI","Apertura","Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
    )
    cons["Cumpl_Acum"] = cons.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)
    cons["MargenPct_Acum"] = cons.apply(lambda r: safe_ratio(r["Margen_Acum"], r["Real_Acum"]), axis=1)
    return cons

def segment_metrics(df_seg):
    real = float(df_seg["Real_Acum"].sum()) if len(df_seg) else 0.0
    obj  = float(df_seg["Obj_Acum"].sum()) if len(df_seg) else 0.0
    cump = safe_ratio(real, obj)
    margen = float(df_seg["Margen_Acum"].sum()) if ("Margen_Acum" in df_seg.columns and len(df_seg)) else 0.0
    margen_pct = safe_ratio(margen, real)
    gap = obj - real  # positivo = faltante
    return real, obj, cump, margen, margen_pct, gap

# ==========================
# CARGA + NORMALIZA
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

default_semana = 0
if 1 in semanas:
    default_semana = semanas.index(1)

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=default_semana)
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

df_last_suc = df_week[df_week["Semana_Num"] == semana_corte].copy()
df_last_suc = aplicar_reglas(df_last_suc, dim_kpi)

if sucursal != "TODAS (Consolidado)":
    df_last = df_last_suc[df_last_suc["Sucursal"] == sucursal].copy()
else:
    df_last = consolidar_todas(df_last_suc)
    df_last = aplicar_reglas(df_last, dim_kpi)

rep_aperturas = sorted(df_last[(df_last["KPI"]=="Repuestos") & (df_last["Apertura"]!="TOTAL")]["Apertura"].dropna().unique())
srv_aperturas = sorted(df_last[(df_last["KPI"]=="Servicios") & (df_last["Apertura"]!="TOTAL")]["Apertura"].dropna().unique())

st.sidebar.markdown("---")
st.sidebar.subheader("Incluir variables (P&L)")
rep_sel = st.sidebar.multiselect("Repuestos: aperturas incluidas", options=rep_aperturas, default=rep_aperturas)
srv_sel = st.sidebar.multiselect("Servicios: aperturas incluidas", options=srv_aperturas, default=srv_aperturas)

st.sidebar.markdown("---")
rank_metric = st.sidebar.selectbox("Ranking (micro)", ["Cumplimiento %", "Real (monto)"], index=0)

st.sidebar.markdown("---")
rank_gap_scope = st.sidebar.selectbox("Ranking GAP $ (macro)", ["Repuestos", "Servicios", "Total Posventa"], index=2)

# ==========================
# HEADER
# ==========================
st.title("Tablero Posventa — Macro → Micro (Semanal + Acumulado)")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🏆 P&L (Repuestos vs Servicios)", "📌 KPIs (resto)", "🧩 Gestión (desvíos)"])

# ==========================
# TAB 1 — P&L + V2.0 (Strip + Ranking Gap + Tendencia)
# ==========================
with tab1:
    # --- Segmentos base (en $)
    rep = df_last[(df_last["KPI"]=="Repuestos") & (df_last["Tipo_KPI"]=="$")].copy()
    srv = df_last[(df_last["KPI"]=="Servicios") & (df_last["Tipo_KPI"]=="$")].copy()

    rep_total = rep[rep["Apertura"].isin(rep_sel)].copy() if rep_sel else rep[rep["Apertura"]!="TOTAL"].copy()
    srv_total = srv[srv["Apertura"].isin(srv_sel)].copy() if srv_sel else srv[srv["Apertura"]!="TOTAL"].copy()

    rep_real, rep_obj, rep_cump, rep_margen, rep_margen_pct, rep_gap = segment_metrics(rep_total)
    srv_real, srv_obj, srv_cump, srv_margen, srv_margen_pct, srv_gap = segment_metrics(srv_total)

    tot_real = rep_real + srv_real
    tot_obj  = rep_obj + srv_obj
    tot_cump = safe_ratio(tot_real, tot_obj)
    tot_gap  = tot_obj - tot_real

    # ==========================
    # V2.0 — STRIP EJECUTIVO
    # ==========================
    st.markdown("## 🏛️ Resumen Ejecutivo (Dirección)")
    ec1, ec2, ec3 = st.columns(3, gap="large")
    with ec1:
        exec_card("REPUESTOS", rep_cump, rep_real, rep_obj, rep_gap)
    with ec2:
        exec_card("SERVICIOS", srv_cump, srv_real, srv_obj, srv_gap)
    with ec3:
        exec_card("TOTAL POSVENTA", tot_cump, tot_real, tot_obj, tot_gap)

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # ==========================
    # V2.0 — RANKING POR GAP $ (macro por sucursal)
    # ==========================
    st.markdown("## 💰 Ranking por GAP $ (Obj - Real) — por sucursal")

    base_rank = df_last_suc.copy()
    base_rank = base_rank[base_rank["Tipo_KPI"] == "$"].copy()
    base_rank = base_rank[(pd.notna(base_rank["Obj_Acum"])) & (base_rank["Obj_Acum"] > 0)].copy()

    # Filtrar aperturas incluidas
    rep_rank = base_rank[(base_rank["KPI"]=="Repuestos") & (base_rank["Apertura"].isin(rep_sel))].copy()
    srv_rank = base_rank[(base_rank["KPI"]=="Servicios") & (base_rank["Apertura"].isin(srv_sel))].copy()

    rep_suc = rep_rank.groupby("Sucursal", as_index=False).agg(Real=("Real_Acum","sum"), Obj=("Obj_Acum","sum"))
    rep_suc["Gap"] = rep_suc["Obj"] - rep_suc["Real"]
    rep_suc["Segmento"] = "Repuestos"

    srv_suc = srv_rank.groupby("Sucursal", as_index=False).agg(Real=("Real_Acum","sum"), Obj=("Obj_Acum","sum"))
    srv_suc["Gap"] = srv_suc["Obj"] - srv_suc["Real"]
    srv_suc["Segmento"] = "Servicios"

    tot_suc = pd.merge(rep_suc[["Sucursal","Real","Obj","Gap"]], srv_suc[["Sucursal","Real","Obj","Gap"]],
                       on="Sucursal", how="outer", suffixes=("_Rep","_Srv"))
    tot_suc = tot_suc.fillna(0)
    tot_suc["Real"] = tot_suc["Real_Rep"] + tot_suc["Real_Srv"]
    tot_suc["Obj"]  = tot_suc["Obj_Rep"] + tot_suc["Obj_Srv"]
    tot_suc["Gap"]  = tot_suc["Obj"] - tot_suc["Real"]
    tot_suc["Segmento"] = "Total Posventa"
    tot_suc = tot_suc[["Sucursal","Real","Obj","Gap","Segmento"]].copy()

    if rank_gap_scope == "Repuestos":
        rk = rep_suc.copy()
    elif rank_gap_scope == "Servicios":
        rk = srv_suc.copy()
    else:
        rk = tot_suc.copy()

    rk = rk.sort_values("Gap", ascending=False).copy()
    rk["Gap_txt"] = rk["Gap"].apply(money_fmt_signed)
    order = rk["Sucursal"].tolist()

    fig_gap = px.bar(
        rk,
        x="Gap",
        y="Sucursal",
        orientation="h",
        text="Gap_txt",
        category_orders={"Sucursal": order}
    )
    fig_gap.update_layout(
        height=420,
        xaxis_title="GAP $ (Obj - Real)  |  (+) faltante  ·  (-) sobrecumplimiento",
        yaxis_title=""
    )
    st.plotly_chart(fig_gap, use_container_width=True)

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # ==========================
    # V2.0 — TENDENCIA SEMANAL (acumulado)
    # ==========================
    st.markdown("## 📈 Tendencia semanal acumulada (Cumplimiento %)")

    trend = df_week[df_week["Tipo_KPI"] == "$"].copy()
    trend = trend[(trend["Semana_Num"].notna()) & (trend["Semana_Num"] <= semana_corte)].copy()

    # aplicar sucursal filtro global para tendencia
    if sucursal != "TODAS (Consolidado)":
        trend = trend[trend["Sucursal"] == sucursal].copy()

    # Repuestos/Servicios con aperturas incluidas
    tr_rep = trend[(trend["KPI"]=="Repuestos") & (trend["Apertura"].isin(rep_sel))].copy()
    tr_srv = trend[(trend["KPI"]=="Servicios") & (trend["Apertura"].isin(srv_sel))].copy()

    tr_rep = tr_rep.groupby("Semana_Num", as_index=False).agg(Real=("Real_Acum","sum"), Obj=("Obj_Acum","sum"))
    tr_rep["Cumpl"] = tr_rep.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    tr_rep["Segmento"] = "Repuestos"

    tr_srv = tr_srv.groupby("Semana_Num", as_index=False).agg(Real=("Real_Acum","sum"), Obj=("Obj_Acum","sum"))
    tr_srv["Cumpl"] = tr_srv.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    tr_srv["Segmento"] = "Servicios"

    tr_tot = pd.merge(tr_rep[["Semana_Num","Real","Obj"]], tr_srv[["Semana_Num","Real","Obj"]],
                      on="Semana_Num", how="outer", suffixes=("_Rep","_Srv")).fillna(0)
    tr_tot["Real"] = tr_tot["Real_Rep"] + tr_tot["Real_Srv"]
    tr_tot["Obj"]  = tr_tot["Obj_Rep"] + tr_tot["Obj_Srv"]
    tr_tot["Cumpl"] = tr_tot.apply(lambda r: safe_ratio(r["Real"], r["Obj"]), axis=1)
    tr_tot["Segmento"] = "Total Posventa"
    tr_tot = tr_tot[["Semana_Num","Cumpl","Segmento"]].copy()

    tr_rep = tr_rep[["Semana_Num","Cumpl","Segmento"]].copy()
    tr_srv = tr_srv[["Semana_Num","Cumpl","Segmento"]].copy()

    trend_all = pd.concat([tr_rep, tr_srv, tr_tot], ignore_index=True)
    trend_all = trend_all.dropna(subset=["Cumpl"]).copy()
    trend_all = trend_all.sort_values(["Segmento","Semana_Num"]).copy()

    if len(trend_all) == 0:
        st.info("No hay datos suficientes para tendencia con los filtros actuales.")
    else:
        fig_tr = px.line(
            trend_all,
            x="Semana_Num",
            y="Cumpl",
            color="Segmento",
            markers=True
        )
        fig_tr.update_yaxes(tickformat=".0%")
        fig_tr.update_layout(height=360, xaxis_title="Semana", yaxis_title="Cumplimiento acumulado")
        st.plotly_chart(fig_tr, use_container_width=True)

    # ==========================
    # Macro → Micro (lo existente)
    # ==========================
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    st.markdown("## 🔎 Macro → Micro (detalle por aperturas)")

    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown("### 🧩 REPUESTOS (P&L)")
        a,b,c = st.columns(3)
        with a: metric_box("Cumplimiento (Acum.)", pct_fmt_ratio(rep_cump), f"Estado: {badge_html(estado_por_umbral(rep_cump))}")
        with b: metric_box("Real (Acum.)", money_fmt(rep_real), f"Objetivo: {money_fmt(rep_obj)}")
        with c: metric_box("Margen % (Acum.)", pct_fmt_ratio(rep_margen_pct), f"Margen: {money_fmt(rep_margen)}")

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("#### Aperturas (Repuestos) — micro")
        rep_micro = rep_total.dropna(subset=["Cumpl_Acum"]).copy()
        rep_micro = rep_micro[(pd.notna(rep_micro["Obj_Acum"])) & (rep_micro["Obj_Acum"] > 0)].copy()
        if len(rep_micro):
            rep_micro = rep_micro.sort_values("Cumpl_Acum", ascending=False).copy()
            order = rep_micro["Apertura"].tolist()
            fig = px.bar(
                rep_micro, x="Cumpl_Acum", y="Apertura", orientation="h",
                text=rep_micro["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%"),
                category_orders={"Apertura": order}
            )
            fig.update_layout(xaxis_tickformat=".0%", yaxis_title="Apertura", height=360)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Repuestos: sin objetivos válidos para aperturas seleccionadas.")

    with c2:
        st.markdown("### 🧩 SERVICIOS (P&L)")
        a,b,c = st.columns(3)
        with a: metric_box("Cumplimiento (Acum.)", pct_fmt_ratio(srv_cump), f"Estado: {badge_html(estado_por_umbral(srv_cump))}")
        with b: metric_box("Real (Acum.)", money_fmt(srv_real), f"Objetivo: {money_fmt(srv_obj)}")
        with c: metric_box("Margen % (Acum.)", pct_fmt_ratio(srv_margen_pct), f"Margen: {money_fmt(srv_margen)}")

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("#### Aperturas (Servicios) — micro")
        srv_micro = srv_total.dropna(subset=["Cumpl_Acum"]).copy()
        srv_micro = srv_micro[(pd.notna(srv_micro["Obj_Acum"])) & (srv_micro["Obj_Acum"] > 0)].copy()
        if len(srv_micro):
            srv_micro = srv_micro.sort_values("Cumpl_Acum", ascending=False).copy()
            order = srv_micro["Apertura"].tolist()
            fig2 = px.bar(
                srv_micro, x="Cumpl_Acum", y="Apertura", orientation="h",
                text=srv_micro["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%"),
                category_orders={"Apertura": order}
            )
            fig2.update_layout(xaxis_tickformat=".0%", yaxis_title="Apertura", height=360)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Servicios: sin objetivos válidos para aperturas seleccionadas.")

    # ========= MICRO PRO (renombrado a Micro —)
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    st.markdown("## 🎯 Micro — ranking sucursal + apertura")

    cA, cB, cC, cD = st.columns([1.2, 1.2, 1, 1])
    with cA:
        rep_micro_mode = st.selectbox("Repuestos (micro)", ["Todas las aperturas", "Elegir 1 apertura"], index=0)
    with cB:
        srv_micro_mode = st.selectbox("Servicios (micro)", ["Todas las aperturas", "Elegir 1 apertura"], index=0)
    with cC:
        top_n_micro = st.selectbox("Top N", [5, 10, 15, 20], index=1)
    with cD:
        ver_ceros = st.checkbox("Mostrar 0% (obj>0 y real=0)", value=True)

    rep_ap_sel = None
    srv_ap_sel = None
    if rep_micro_mode == "Elegir 1 apertura":
        rep_ap_sel = st.selectbox("Apertura Repuestos", rep_sel if rep_sel else rep_aperturas, index=0)
    if srv_micro_mode == "Elegir 1 apertura":
        srv_ap_sel = st.selectbox("Apertura Servicios", srv_sel if srv_sel else srv_aperturas, index=0)

    def micro_rank_chart(df_rank_base, kpi_name, aperturas_incluidas, modo, apertura_unica, top_n, tipo="$"):
        dfm = df_rank_base[
            (df_rank_base["KPI"] == kpi_name) &
            (df_rank_base["Tipo_KPI"] == tipo) &
            (df_rank_base["Apertura"].isin(aperturas_incluidas))
        ].copy()

        if modo == "Elegir 1 apertura" and apertura_unica:
            dfm = dfm[dfm["Apertura"] == apertura_unica].copy()

        dfm = dfm[(pd.notna(dfm["Obj_Acum"])) & (dfm["Obj_Acum"] > 0)].copy()
        if not ver_ceros:
            dfm = dfm[dfm["Real_Acum"] > 0].copy()

        xcol = "Cumpl_Acum" if rank_metric == "Cumplimiento %" else "Real_Acum"
        dfm = dfm.dropna(subset=[xcol]).copy()

        if not len(dfm):
            return None

        dfm["Etiqueta"] = dfm["Sucursal"].astype(str) + " — " + dfm["Apertura"].astype(str)
        dfm = dfm.sort_values(xcol, ascending=False).head(top_n).copy()
        order = dfm["Etiqueta"].tolist()

        txt = dfm[xcol].apply(lambda v: f"{v*100:.1f}%" if xcol=="Cumpl_Acum" else money_fmt(v))

        fig = px.bar(
            dfm, x=xcol, y="Etiqueta", orientation="h",
            text=txt, category_orders={"Etiqueta": order}
        )
        if xcol == "Cumpl_Acum":
            fig.update_layout(xaxis_tickformat=".0%")
        fig.update_layout(yaxis_title="", height=420)
        return fig

    mc1, mc2 = st.columns(2, gap="large")
    with mc1:
        st.markdown("#### Repuestos — sucursal + apertura (micro)")
        figm1 = micro_rank_chart(
            df_rank_base=df_last_suc,
            kpi_name="Repuestos",
            aperturas_incluidas=rep_sel if rep_sel else rep_aperturas,
            modo=rep_micro_mode,
            apertura_unica=rep_ap_sel,
            top_n=top_n_micro
        )
        if figm1:
            st.plotly_chart(figm1, use_container_width=True)
        else:
            st.info("No hay datos micro de Repuestos con los filtros actuales.")

    with mc2:
        st.markdown("#### Servicios — sucursal + apertura (micro)")
        figm2 = micro_rank_chart(
            df_rank_base=df_last_suc,
            kpi_name="Servicios",
            aperturas_incluidas=srv_sel if srv_sel else srv_aperturas,
            modo=srv_micro_mode,
            apertura_unica=srv_ap_sel,
            top_n=top_n_micro
        )
        if figm2:
            st.plotly_chart(figm2, use_container_width=True)
        else:
            st.info("No hay datos micro de Servicios con los filtros actuales.")

# ==========================
# TAB 2 — KPIs resto
# ==========================
with tab2:
    st.markdown("### 📌 KPIs (resto) — Macro → Micro")
    st.markdown('<span class="small-muted">KPIs que NO son Repuestos/Servicios (incluye $ y Q).</span>', unsafe_allow_html=True)

    base = df_week[df_week["Semana_Num"] == semana_corte].copy()
    base = base[~base["KPI"].isin(["Repuestos","Servicios"])].copy()

    kpis_resto = sorted(base["KPI"].dropna().unique().tolist())
    if not kpis_resto:
        st.info("No hay KPIs adicionales en este corte.")
    else:
        kpi_sel = st.selectbox("Elegí un KPI (resto)", kpis_resto, index=0)

        if sucursal != "TODAS (Consolidado)":
            df_k = base[(base["Sucursal"] == sucursal) & (base["KPI"] == kpi_sel)].copy()
            df_k = df_k.groupby(["KPI","Tipo_KPI"], as_index=False).agg(Real_Acum=("Real_Acum","sum"), Obj_Acum=("Obj_Acum","sum"))
        else:
            df_k = base[base["KPI"] == kpi_sel].copy()
            df_k = df_k.groupby(["KPI","Tipo_KPI"], as_index=False).agg(Real_Acum=("Real_Acum","sum"), Obj_Acum=("Obj_Acum","sum"))

        tipo = df_k["Tipo_KPI"].iloc[0] if len(df_k) else "—"
        real = float(df_k["Real_Acum"].sum()) if len(df_k) else 0.0
        obj  = float(df_k["Obj_Acum"].sum()) if len(df_k) else 0.0
        cump = safe_ratio(real, obj)
        estado = estado_por_umbral(cump)

        a,b,c = st.columns(3)
        with a: metric_box("KPI", f"{kpi_sel} ({tipo})", f"Estado: {badge_html(estado)}")
        with b:
            val = money_fmt(real) if tipo=="$" else num_fmt(real)
            metric_box("Real (Acum.)", val, f"Cumpl.: {pct_fmt_ratio(cump)}")
        with c:
            val = money_fmt(obj) if tipo=="$" else num_fmt(obj)
            metric_box("Objetivo (Acum.)", val, "&nbsp;")

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.markdown("#### Ranking por sucursal — este KPI")

        rk = base[base["KPI"] == kpi_sel].copy()
        rk = rk.groupby(["Sucursal","Tipo_KPI"], as_index=False).agg(Real_Acum=("Real_Acum","sum"), Obj_Acum=("Obj_Acum","sum"))
        rk["Cumpl_Acum"] = rk.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)

        xcol = "Cumpl_Acum" if rank_metric=="Cumplimiento %" else "Real_Acum"
        rk = rk.sort_values(xcol, ascending=False).copy()
        order = rk["Sucursal"].tolist()

        fig = px.bar(
            rk, x=xcol, y="Sucursal", orientation="h",
            text=rk[xcol].apply(lambda v: f"{v*100:.1f}%" if xcol=="Cumpl_Acum" and pd.notna(v) else (money_fmt(v) if tipo=="$" else num_fmt(v))),
            category_orders={"Sucursal": order}
        )
        if xcol=="Cumpl_Acum": fig.update_layout(xaxis_tickformat=".0%")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

# ==========================
# TAB 3 — Gestión (con filtro sucursal + Drivers)
# ==========================
with tab3:
    st.markdown("### 🧩 Gestión (desvíos acumulados)")
    st.markdown('<span class="small-muted">Filtro de gestión independiente del filtro global.</span>', unsafe_allow_html=True)

    colf1, colf2 = st.columns([1.2, 3])
    with colf1:
        suc_gestion = st.selectbox("Sucursal (Gestión)", ["TODAS (Consolidado)"] + sucursales, index=0)
    with colf2:
        st.write("")

    if suc_gestion != "TODAS (Consolidado)":
        g = df_last_suc[df_last_suc["Sucursal"] == suc_gestion].copy()
        g = g.groupby(["KPI","Apertura","Tipo_KPI"], as_index=False).agg(
            Real_Acum=("Real_Acum","sum"),
            Obj_Acum=("Obj_Acum","sum"),
            Margen_Acum=("Margen_Acum","sum"),
        )
        g["Cumpl_Acum"] = g.apply(lambda r: safe_ratio(r["Real_Acum"], r["Obj_Acum"]), axis=1)
        g = aplicar_reglas(g, dim_kpi)
    else:
        g = consolidar_todas(df_last_suc)
        g = aplicar_reglas(g, dim_kpi)

    g["Gap"] = g["Obj_Acum"] - g["Real_Acum"]

    order_map = {"Rojo": 0, "Amarillo": 1, "Verde": 2, "—": 9}
    g["OrdenEstado"] = g["Estado_Acum"].map(order_map).fillna(9)
    g = g.sort_values(["OrdenEstado","Gap"], ascending=[True, False]).copy()

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

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
    st.markdown("### 🎯 Drivers del desvío (qué explica el gap)")

    c1, c2, c3, c4 = st.columns([1.2, 1.8, 1, 1])
    with c1:
        tipo_drv = st.selectbox("Tipo", ["$", "Q"], index=0, key="drv_tipo")
    with c2:
        kpi_drv = st.selectbox("KPI", sorted(g["KPI"].dropna().unique().tolist()), index=0, key="drv_kpi")
    with c3:
        top_n = st.selectbox("Top", [5, 10, 15, 20], index=1, key="drv_top")
    with c4:
        solo_gap_pos = st.checkbox("Solo gap positivo", value=True, key="drv_pos")

    gp = g[(g["Tipo_KPI"] == tipo_drv) & (g["KPI"] == kpi_drv)].copy()
    gp = gp[(pd.notna(gp["Obj_Acum"])) & (gp["Obj_Acum"] > 0)].copy()

    gp["Gap"] = gp["Obj_Acum"] - gp["Real_Acum"]
    if solo_gap_pos:
        gp = gp[gp["Gap"] > 0].copy()

    if len(gp) == 0:
        st.info("No hay desvíos para mostrar con los filtros actuales.")
    else:
        gap_total = float(gp["Gap"].sum())
        gap_total_safe = gap_total if gap_total != 0 else 1.0

        gp["Aporte_%"] = gp["Gap"] / gap_total_safe
        gp = gp.sort_values("Gap", ascending=False).head(top_n).copy()

        v_gap = money_fmt(gap_total) if tipo_drv == "$" else num_fmt(gap_total)
        metric_box("Gap total (Obj - Real)", v_gap, f"KPI: <b>{kpi_drv}</b> ({tipo_drv})")

        gp_plot = gp.copy()
        gp_plot["Gap_txt"] = gp_plot["Gap"].apply(lambda v: money_fmt(v) if tipo_drv=="$" else num_fmt(v))
        gp_plot["Aporte_txt"] = gp_plot["Aporte_%"].apply(lambda v: f"{v*100:.1f}%")
        gp_plot = gp_plot.sort_values("Gap", ascending=True)

        figd = px.bar(
            gp_plot,
            x="Gap",
            y="Apertura",
            orientation="h",
            text=gp_plot["Gap_txt"] + "  (" + gp_plot["Aporte_txt"] + ")"
        )
        figd.update_layout(
            height=420,
            xaxis_title="Gap (Obj - Real)  |  (Aporte al gap total)",
            yaxis_title="Apertura"
        )
        st.plotly_chart(figd, use_container_width=True)

        tbl = gp[["Apertura","Estado_Acum","Cumpl_Acum","Real_Acum","Obj_Acum","Gap","Aporte_%"]].copy()
        tbl["Cumpl_Acum"] = tbl["Cumpl_Acum"].apply(pct_fmt_ratio)
        tbl["Real_Acum"] = tbl["Real_Acum"].apply(lambda v: money_fmt(v) if tipo_drv=="$" else num_fmt(v))
        tbl["Obj_Acum"]  = tbl["Obj_Acum"].apply(lambda v: money_fmt(v) if tipo_drv=="$" else num_fmt(v))
        tbl["Gap"]       = tbl["Gap"].apply(lambda v: money_fmt(v) if tipo_drv=="$" else num_fmt(v))
        tbl["Aporte_%"]  = tbl["Aporte_%"].apply(lambda v: f"{v*100:.1f}%")
        st.dataframe(tbl, use_container_width=True, hide_index=True)
